# -*- coding: utf-8 -*-
"""修正转换后 TTF 的 name / OS/2 / head 样式位，使 Word 能按 b/i 正确匹配。

上游 OTF 的 name 表把光学尺寸写进 subfamily（如 "9 Regular"），
Word 只认 Regular/Bold/Italic/Bold Italic，必须规范化：
  name1 (family)    = LM Roman 9
  name2 (subfamily) = Regular / Bold / Italic
  name4 (full)      = LM Roman 9 Bold
  name6 (postscript)= LMRoman9-Bold
  name16/17 (typo)  = LM Roman 9 / Regular
同时同步 OS/2.fsSelection 与 head.macStyle。
"""
import os, sys, io, re
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from fontTools.ttLib import TTFont
import l2wconf as C

DIR = C.TTFDIR

# fsSelection 位
FS_ITALIC = 1 << 0
FS_BOLD = 1 << 5
FS_REGULAR = 1 << 6
FS_USE_TYPO = 1 << 7
FS_OBLIQUE = 1 << 9


def style_of(fn):
    if 'bolditalic' in fn:
        return 'Bold Italic'
    if 'bold' in fn:
        return 'Bold'
    if 'italic' in fn or 'oblique' in fn:
        return 'Italic'
    return 'Regular'


def set_name(font, nid, val):
    """同时写 Windows(3,1,0x409) 与 Mac(1,0,0) 两套，确保各平台可读。"""
    nt = font['name']
    for plat, enc, lang in ((3, 1, 0x409), (1, 0, 0)):
        nt.setName(val, nid, plat, enc, lang)


def main():
    for fn in sorted(os.listdir(DIR)):
        if not fn.endswith('.ttf'):
            continue
        p = os.path.join(DIR, fn)
        f = TTFont(p, lazy=False)
        nm = {r.nameID: r.toUnicode() for r in f['name'].names if r.platformID == 3}
        fam = nm.get(1) or 'LM Roman 9'
        # 去掉 subfamily 里的光学尺寸前缀
        sub = style_of(fn)
        full = fam if sub == 'Regular' else f'{fam} {sub}'
        ps = re.sub(r'[^A-Za-z0-9]', '', fam) + ('-' + sub.replace(' ', '') if sub != 'Regular' else '')

        set_name(f, 1, fam)
        set_name(f, 2, sub)
        set_name(f, 4, full)
        set_name(f, 6, ps)
        set_name(f, 16, fam)
        set_name(f, 17, sub)

        bold = sub in ('Bold', 'Bold Italic')
        ital = sub in ('Italic', 'Bold Italic')
        fs = FS_USE_TYPO
        fs |= FS_BOLD if bold else 0
        fs |= FS_ITALIC if ital else 0
        fs |= FS_REGULAR if sub == 'Regular' else 0
        f['OS/2'].fsSelection = fs
        f['OS/2'].usWeightClass = 700 if bold else 400
        f['head'].macStyle = (1 if bold else 0) | (2 if ital else 0)
        # post.italicAngle 保持原值；OS/2 版本需 >=4 才支持 USE_TYPO_METRICS
        if f['OS/2'].version < 4:
            f['OS/2'].version = 4
        f.save(p)
        f.close()
        print(f'  {fn:<28s} fam={fam:<20s} sub={sub:<12s} fsSel=0x{fs:03X} weight={700 if bold else 400}')
    print('\nname/样式位修正完成')


if __name__ == '__main__':
    main()
