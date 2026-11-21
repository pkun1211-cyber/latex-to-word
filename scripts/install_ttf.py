# -*- coding: utf-8 -*-
r"""用户级安装 TrueType 版 LM Roman（免管理员，可完整卸载）。

为什么必须装字体
----------------
LaTeX 的 Latin Modern 是**光学尺寸族**：同一字体在 5/6/7/8/9/10/12/17pt
下字形宽度不同。Word 若用 Times New Roman 代替，西文行宽会系统性偏差，
逐行硬对齐永远收不敛。装 LM Roman TTF 后 Word 与 TeX 用同一套字形度量。

为什么用 TTF 而非 OTF
---------------------
Word 对 CFF/OTF 的字距（w:spacing）处理与 TrueType 不同，
且部分 OTF 在 Word 里 kerning 行为异常。先用 otf2ttf.py 转换。

安装位置与可逆性
----------------
文件 → %LOCALAPPDATA%\Microsoft\Windows\Fonts（用户级，不动系统目录）
注册 → HKCU\Software\Microsoft\Windows NT\CurrentVersion\Fonts
清单 → C.tmp('install_ttf_map.json')，供 uninstall_ttf.py 精确回滚
已存在的同名文件不覆盖（existed=True），卸载时也不会删除。

用法: python install_ttf.py
"""
import os, sys, io, json, shutil, hashlib, winreg
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from fontTools.ttLib import TTFont
import l2wconf as C

SRC = C.TTFDIR
DEST = os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Windows\Fonts')
REG_KEY = r'Software\Microsoft\Windows NT\CurrentVersion\Fonts'
MAP = C.tmp('install_ttf_map.json')


def main():
    if not os.path.isdir(SRC):
        raise SystemExit('字体目录不存在: %s\n先运行 otf2ttf.py 生成 TTF。' % SRC)
    os.makedirs(DEST, exist_ok=True)
    plan = []
    for fn in sorted(os.listdir(SRC)):
        if not fn.endswith('.ttf'):
            continue
        s = os.path.join(SRC, fn)
        f = TTFont(s, lazy=True)
        nm = {r.nameID: r.toUnicode() for r in f['name'].names if r.platformID == 3}
        full = nm.get(4)
        f.close()
        if not full:
            print('  SKIP %s（无 nameID 4，先跑 fix_ttf_names.py）' % fn)
            continue
        rn = full + ' (TrueType)'
        dst = os.path.join(DEST, fn)
        sha = hashlib.sha256(open(s, 'rb').read()).hexdigest()
        existed = os.path.exists(dst)
        if not existed:
            shutil.copy2(s, dst)
        assert hashlib.sha256(open(dst, 'rb').read()).hexdigest() == sha
        plan.append({'file': fn, 'src': s, 'dst': dst, 'reg': rn,
                     'sha256': sha, 'existed': existed})
        print(f'  {"SKIP" if existed else "COPY"} {fn:<28s} -> {rn}')

    k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE)
    for it in plan:
        winreg.SetValueEx(k, it['reg'], 0, winreg.REG_SZ, it['dst'])
    winreg.CloseKey(k)
    json.dump(plan, open(MAP, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'\n已安装 {len(plan)} 个 TTF，注册表已写。')
    print(f'卸载清单: {MAP}')
    print('卸载命令: python uninstall_ttf.py')
    print('注意：Word 需重启才能识别新字体。')


if __name__ == '__main__':
    main()
