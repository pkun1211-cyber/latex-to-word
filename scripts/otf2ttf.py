# -*- coding: utf-8 -*-
r"""把 LM Roman 的 OTF(CFF 三次曲线) 转为 TTF(TrueType 二次曲线)。

为什么必须转
------------
Word 对 CFF 轮廓的 OTF 支持不完整：实测把 w:rFonts 指向 LM Roman（OTF/CFF）后，
导出的 PDF 里西文字符**整体消失**（字体表无 LM Roman，字符数骤减），
说明 Word 渲染阶段直接丢弃了这些字形。转成 TrueType 轮廓即可正常渲染。

转换要点
--------
1. CFF charstring → cu2qu 二次贝塞尔 → glyf/loca 表
2. upm 必须为 2 的幂（1000 → 2048），否则 TrueType 渲染精度不足
3. 重建 head/hhea/maxp/post/OS/2，删除 CFF 相关表
4. name 表 family 保持 "LM Roman 9"，subfamily 保持 Regular/Bold/Italic
   并补 fsSelection / macStyle 位，使 Word 能正确匹配 b/i
5. 保留 hmtx 度量（字宽必须与 LaTeX 完全一致，这是对齐的前提）

用法: python otf2ttf.py [输出目录]
"""
import os, sys, io, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from fontTools.ttLib import TTFont, newTable
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.ttLib.scaleUpem import scale_upem
import l2wconf as C

LM_DIR = r'C:\Users\20546\AppData\Local\Programs\MiKTeX\fonts\opentype\public\lm'
OUT = C.TTFDIR

NEED = [
    'lmroman5-regular.otf', 'lmroman6-regular.otf', 'lmroman7-regular.otf',
    'lmroman8-regular.otf', 'lmroman9-regular.otf', 'lmroman9-bold.otf',
    'lmroman9-italic.otf', 'lmroman10-regular.otf', 'lmroman10-bold.otf',
    'lmroman10-italic.otf', 'lmroman12-regular.otf', 'lmroman17-regular.otf',
    'lmromandemi10-regular.otf',
]


def convert(src, dst, max_err=0.6):
    f = TTFont(src, lazy=False)
    assert 'CFF ' in f, 'not a CFF font'
    glyphOrder = f.getGlyphOrder()
    glyphSet = f.getGlyphSet()

    # 1) 先把 upm 提到 2048（在 CFF 阶段做，度量一并缩放）
    if f['head'].unitsPerEm != 2048:
        scale_upem(f, 2048)

    # 2) CFF -> glyf
    glyf = newTable('glyf')
    glyf.glyphOrder = glyphOrder
    glyf.glyphs = {}
    for name in glyphOrder:
        pen = TTGlyphPen(glyphSet)
        cu2qu = Cu2QuPen(pen, max_err, reverse_direction=True)
        try:
            glyphSet[name].draw(cu2qu)
            glyf[name] = pen.glyph()
        except Exception:
            # 个别字形转换失败则置空，避免整体失败
            glyf[name] = pen.glyph()
    f['glyf'] = glyf

    # 3) loca（由 glyf 编译时自动生成，占位即可）
    f['loca'] = newTable('loca')

    # 4) head：indexToLocFormat 由编译决定；改 sfntVersion
    f.sfntVersion = '\x00\x01\x00\x00'
    f['head'].indexToLocFormat = 1

    # 5) maxp：TrueType 版本 1.0
    maxp = newTable('maxp')
    maxp.tableVersion = 0x00010000
    maxp.numGlyphs = len(glyphOrder)
    maxp.maxPoints = 0
    maxp.maxContours = 0
    maxp.maxCompositePoints = 0
    maxp.maxCompositeContours = 0
    maxp.maxZones = 2
    maxp.maxTwilightPoints = 0
    maxp.maxStorage = 0
    maxp.maxFunctionDefs = 0
    maxp.maxInstructionDefs = 0
    maxp.maxStackElements = 0
    maxp.maxSizeOfInstructions = 0
    maxp.maxComponentElements = 0
    maxp.maxComponentDepth = 0
    f['maxp'] = maxp

    # 6) post：格式 2.0 带字形名
    post = newTable('post')
    post.formatType = 2.0
    post.isFixedPitch = f['post'].isFixedPitch if 'post' in f else 0
    post.italicAngle = f['post'].italicAngle if 'post' in f else 0
    post.underlinePosition = f['post'].underlinePosition if 'post' in f else -100
    post.underlineThickness = f['post'].underlineThickness if 'post' in f else 50
    post.minMemType42 = 0
    post.maxMemType42 = 0
    post.minMemType1 = 0
    post.maxMemType1 = 0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = glyphOrder
    f['post'] = post

    # 7) 删除 CFF 专有表
    for t in ('CFF ', 'CFF2', 'VORG'):
        if t in f:
            del f[t]

    # 8) 编译期重算 maxp 的极值（recalc）
    f['glyf'].compile(f)  # 触发一次以填充极值
    f.save(dst)
    f.close()
    return dst


if __name__ == '__main__':
    args = [a for a in sys.argv[1:]]
    only = None
    if '--only' in args:
        i = args.index('--only')
        only = args[i + 1]
        del args[i:i + 2]
    outdir = args[0] if args else OUT
    os.makedirs(outdir, exist_ok=True)
    files = [only] if only else NEED
    ok = fail = 0
    for fn in files:
        src = os.path.join(LM_DIR, fn)
        dst = os.path.join(outdir, fn.replace('.otf', '.ttf'))
        try:
            convert(src, dst)
            # 校验：能被重新解析且含 glyf
            g = TTFont(dst, lazy=True)
            assert 'glyf' in g and 'CFF ' not in g
            upm = g['head'].unitsPerEm
            ng = g['maxp'].numGlyphs
            # 字宽校验：与源字体在 9pt 下逐字符比对
            s = TTFont(src, lazy=True)
            sc, dc = s.getBestCmap(), g.getBestCmap()
            worst = 0.0
            for ch in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ,.;:%-':
                if ord(ch) not in sc or ord(ch) not in dc:
                    continue
                a = s['hmtx'][sc[ord(ch)]][0] / s['head'].unitsPerEm * 9
                b = g['hmtx'][dc[ord(ch)]][0] / upm * 9
                worst = max(worst, abs(a - b))
            s.close(); g.close()
            print(f'  OK  {fn:<28s} -> {os.path.basename(dst):<28s} '
                  f'upm={upm} glyphs={ng} 字宽最大偏差={worst:.4f}pt')
            ok += 1
        except Exception as e:
            print(f'  FAIL {fn}: {type(e).__name__}: {e}')
            fail += 1
    print(f'\n转换完成 OK={ok} FAIL={fail}  输出目录={outdir}')
