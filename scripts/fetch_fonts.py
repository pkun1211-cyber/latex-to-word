# -*- coding: utf-8 -*-
r"""从 CTAN 官方镜像获取 Latin Roman OTF，转成 Word 可用的 TTF。

为什么仓库不直接放字体
----------------------
1. 二进制字体 2.8MB，占仓库体积的 90%
2. 避免字体再分发的许可细节 —— 直接从官方源取，用户拿到的是原版

为什么必须转 TTF
----------------
Word 对 CFF 轮廓的 OTF 支持不完整：把 w:rFonts 指向 LM Roman（OTF/CFF）后，
导出 PDF 里西文字符**整体消失**（Word 渲染阶段直接丢弃字形）。
必须转成 TrueType 二次曲线轮廓。

为什么不能用 Times New Roman 代替
---------------------------------
Latin Modern 是**光学尺寸族**，5/6/7/8/9/10/12/17pt 字形宽度各不相同。
替代字体会让西文行宽系统性偏差，逐行对齐永远收不敛。

用法
----
    python fetch_fonts.py                 # 下载 + 转换（不安装）
    python fetch_fonts.py --install       # 下载 + 转换 + 用户级安装
    python fetch_fonts.py --from <目录>   # 用本地已有 OTF（如 MiKTeX 自带）
    python fetch_fonts.py --mirror <url>  # 指定镜像

本地已有 MiKTeX 时，OTF 通常在：
    C:\\Users\\<user>\\AppData\\Local\\Programs\\MiKTeX\\fonts\\opentype\\public\\lm
"""
import os, sys, io, zipfile, tempfile, shutil, urllib.request

if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import l2wconf as C

# CTAN 镜像，按可用性排序。中科大在国内最快且稳定。
MIRRORS = [
    'https://mirrors.ustc.edu.cn/CTAN/fonts/lm.zip',
    'https://mirrors.tuna.tsinghua.edu.cn/CTAN/fonts/lm.zip',
    'https://mirrors.ctan.org/fonts/lm.zip',
]

NEED = [
    'lmroman5-regular', 'lmroman6-regular', 'lmroman7-regular',
    'lmroman8-regular', 'lmroman9-regular', 'lmroman9-bold',
    'lmroman9-italic', 'lmroman10-regular', 'lmroman10-bold',
    'lmroman10-italic', 'lmroman12-regular', 'lmroman17-regular',
    'lmromandemi10-regular',
]

# MiKTeX / TeX Live 常见位置
LOCAL_HINTS = [
    os.path.expandvars(r'%LOCALAPPDATA%\Programs\MiKTeX\fonts\opentype\public\lm'),
    r'C:\texlive\2024\texmf-dist\fonts\opentype\public\lm',
    r'C:\texlive\2023\texmf-dist\fonts\opentype\public\lm',
    '/usr/share/texlive/texmf-dist/fonts/opentype/public/lm',
    '/usr/local/texlive/2024/texmf-dist/fonts/opentype/public/lm',
]


def find_local():
    """找本机已有的 OTF 目录（省去下载）。"""
    for d in LOCAL_HINTS:
        if os.path.isdir(d):
            hit = sum(1 for n in NEED
                      if os.path.exists(os.path.join(d, n + '.otf')))
            if hit >= len(NEED) - 2:      # 容忍个别缺失
                return d
    return None


def download(mirror=None):
    """下载 lm.zip 并解出 OTF 目录。返回临时目录路径。"""
    urls = [mirror] if mirror else MIRRORS
    tmpd = tempfile.mkdtemp(prefix='lm_otf_')
    zp = os.path.join(tmpd, 'lm.zip')
    for u in urls:
        try:
            print('  下载 %s' % u)
            req = urllib.request.Request(u, headers={'User-Agent': 'latex-to-word'})
            with urllib.request.urlopen(req, timeout=120) as r, open(zp, 'wb') as f:
                shutil.copyfileobj(r, f)
            sz = os.path.getsize(zp)
            if sz < 100000:
                raise ValueError('文件过小(%d字节)，可能是错误页' % sz)
            print('  已下载 %.2f MB' % (sz / 1048576))
            break
        except Exception as e:
            print('  失败: %s: %s' % (type(e).__name__, e))
    else:
        raise SystemExit('所有镜像均下载失败。请检查网络，或用 --from 指定本地 OTF 目录。')

    with zipfile.ZipFile(zp) as z:
        names = [n for n in z.namelist() if n.lower().endswith('.otf')]
        if not names:
            raise SystemExit('压缩包内未找到 OTF')
        for n in names:
            base = os.path.basename(n)
            with z.open(n) as src, open(os.path.join(tmpd, base), 'wb') as dst:
                shutil.copyfileobj(src, dst)
    os.remove(zp)
    print('  解出 %d 个 OTF' % len(names))
    return tmpd


def main():
    args = sys.argv[1:]
    src = None
    mirror = None
    do_install = '--install' in args

    if '--from' in args:
        src = args[args.index('--from') + 1]
        if not os.path.isdir(src):
            raise SystemExit('目录不存在: %s' % src)
        print('使用指定目录: %s' % src)
    if '--mirror' in args:
        mirror = args[args.index('--mirror') + 1]

    cleanup = None
    if not src:
        src = find_local()
        if src:
            print('发现本机已有 OTF: %s' % src)
        else:
            print('本机未找到 OTF，从 CTAN 下载...')
            src = cleanup = download(mirror)

    outdir = C.TTFDIR
    os.makedirs(outdir, exist_ok=True)

    # 复用 otf2ttf 的转换核心
    import otf2ttf
    otf2ttf.LM_DIR = src

    print('\n转换 OTF -> TTF（CFF 三次曲线 -> TrueType 二次曲线）...')
    from fontTools.ttLib import TTFont
    ok = fail = 0
    for n in NEED:
        s = os.path.join(src, n + '.otf')
        d = os.path.join(outdir, n + '.ttf')
        if not os.path.exists(s):
            print('  跳过 %s（源缺失）' % n)
            fail += 1
            continue
        try:
            otf2ttf.convert(s, d)
            g = TTFont(d, lazy=True)
            assert 'glyf' in g and 'CFF ' not in g
            upm, ng = g['head'].unitsPerEm, g['maxp'].numGlyphs
            g.close()
            print('  OK   %-26s upm=%d glyphs=%d' % (n, upm, ng))
            ok += 1
        except Exception as e:
            print('  FAIL %s: %s: %s' % (n, type(e).__name__, e))
            fail += 1

    if cleanup:
        shutil.rmtree(cleanup, ignore_errors=True)

    print('\n转换完成 OK=%d FAIL=%d' % (ok, fail))
    print('输出目录: %s' % outdir)

    if ok == 0:
        raise SystemExit('无一成功，请检查 fontTools 是否已安装：pip install fontTools')

    # 修 name 表，让 Word 能正确匹配粗体/斜体
    print('\n修正 name 表（Word 字体匹配）...')
    import subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    subprocess.run([sys.executable, os.path.join(here, 'fix_ttf_names.py')],
                   check=False)

    if do_install:
        print('\n用户级安装...')
        subprocess.run([sys.executable, os.path.join(here, 'install_ttf.py')],
                       check=False)
        print('\n★ 安装完成，必须重启 Word 才能生效。')
        print('  校验：python audit_lmfont.py')
    else:
        print('\n下一步：python install_ttf.py  （装完重启 Word）')


if __name__ == '__main__':
    main()
