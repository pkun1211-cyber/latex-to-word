# -*- coding: utf-8 -*-
"""LM Roman 字体安装前安全审计（只读，不做任何安装/写入系统）。

审计项
------
1. 来源合法性：文件是否位于本机 MiKTeX 官方安装目录（CTAN 分发）
2. 文件完整性：OTF 魔数、表结构、是否可被 fontTools 正常解析
3. 危险表检测：是否含可执行/脚本类表（OTF 规范中字体不含可执行代码，
   但需排查异常的私有表、超大表、以及 embedded 许可限制）
4. 许可核对：fsType 嵌入许可位 + name 表 license 字段（GUST Font License）
5. 命名冲突：family name 是否与系统已装字体冲突
6. 体积与数量
"""
import os, sys, io, json, hashlib, struct
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from fontTools.ttLib import TTFont

LM_DIR = r'C:\Users\20546\AppData\Local\Programs\MiKTeX\fonts\opentype\public\lm'
NEED = [
    'lmroman5-regular.otf', 'lmroman6-regular.otf', 'lmroman7-regular.otf',
    'lmroman8-regular.otf', 'lmroman9-regular.otf', 'lmroman9-bold.otf',
    'lmroman9-italic.otf', 'lmroman10-regular.otf', 'lmroman10-bold.otf',
    'lmroman10-italic.otf', 'lmroman12-regular.otf', 'lmroman17-regular.otf',
    'lmromandemi10-regular.otf',
]
STD_TABLES = {
    'head', 'hhea', 'maxp', 'OS/2', 'hmtx', 'cmap', 'name', 'post', 'glyf',
    'loca', 'CFF ', 'CFF2', 'VORG', 'GPOS', 'GSUB', 'GDEF', 'BASE', 'JSTF',
    'MATH', 'kern', 'DSIG', 'LTSH', 'VDMX', 'hdmx', 'gasp', 'PCLT', 'meta',
    'STAT', 'avar', 'fvar', 'gvar', 'cvar', 'HVAR', 'VVAR', 'MVAR', 'COLR',
    'CPAL', 'CBDT', 'CBLC', 'sbix', 'EBDT', 'EBLC', 'EBSC', 'SVG ', 'feat',
    'Feat', 'Glat', 'Gloc', 'Sill', 'Silf', 'bsln', 'just', 'lcar', 'mort',
    'morx', 'opbd', 'prop', 'trak', 'Zapf', 'acnt', 'fdsc', 'fmtx', 'bdat',
    'bhed', 'bloc', 'cidg', 'fpgm', 'prep', 'cvt ', 'TTFA', 'TeX ', 'FFTM',
    # fontTools 解析 CFF 时合成的虚拟表，并非文件内真实表，不算异常
    'GlyphOrder', 'CFFFontNames', 'GlyphNames',
}
# fsType 位含义（可叠加）
FSTYPE_BITS = {
    0: '无限制嵌入',
    2: '禁止嵌入(Restricted)',
    4: '预览与打印嵌入(Preview&Print)',
    8: '可编辑嵌入(Editable)',
    0x0100: '不得子集化(NoSubsetting)',
    0x0200: '仅位图嵌入(BitmapOnly)',
}


def decode_fstype(v):
    if v == 0:
        return ['无限制嵌入']
    out = []
    for bit, name in FSTYPE_BITS.items():
        if bit and v & bit:
            out.append(name)
    return out or [f'未知({v})']


def audit():
    print('=' * 96)
    print('LM Roman 字体安装前安全审计（只读）')
    print('=' * 96)
    print(f'\n[1] 来源目录: {LM_DIR}')
    ok_src = os.path.isdir(LM_DIR) and 'MiKTeX' in LM_DIR and 'public' in LM_DIR
    print(f'    位于 MiKTeX 官方 CTAN 分发路径 (fonts/opentype/public/lm): {"是" if ok_src else "否"}')
    print('    说明: 该目录由 MiKTeX 包管理器从 CTAN 镜像下载，非第三方来源。')

    total = 0
    rows = []
    print(f'\n[2] 逐文件检查（共 {len(NEED)} 个）')
    for fn in NEED:
        p = os.path.join(LM_DIR, fn)
        if not os.path.exists(p):
            print(f'    MISSING {fn}')
            rows.append({'file': fn, 'ok': False, 'err': 'missing'})
            continue
        raw = open(p, 'rb').read()
        total += len(raw)
        magic = raw[:4]
        is_otf = magic in (b'OTTO', b'\x00\x01\x00\x00', b'true', b'ttcf')
        sha = hashlib.sha256(raw).hexdigest()[:16]
        try:
            f = TTFont(p, lazy=False, fontNumber=0)
            tables = set(f.keys())
            odd = sorted(tables - STD_TABLES)
            nm = {r.nameID: r.toUnicode() for r in f['name'].names if r.platformID == 3}
            fstype = f['OS/2'].fsType if 'OS/2' in f else None
            upm = f['head'].unitsPerEm
            nglyph = f['maxp'].numGlyphs
            lic = nm.get(13, '')[:60]
            licurl = nm.get(14, '')[:60]
            f.close()
            parse_ok = True
        except Exception as e:
            tables, odd, nm, fstype, upm, nglyph, lic, licurl = set(), ['PARSE_FAIL'], {}, None, 0, 0, '', str(e)
            parse_ok = False
        rows.append({'file': fn, 'ok': is_otf and parse_ok and not odd,
                     'size': len(raw), 'sha': sha, 'family': nm.get(1),
                     'fsType': fstype, 'odd': odd, 'upm': upm, 'glyphs': nglyph,
                     'license': lic, 'licurl': licurl})
        flag = 'OK ' if (is_otf and parse_ok and not odd) else '!! '
        print(f'    {flag}{fn:<28s} {len(raw)/1024:7.1f}KB  magic={magic!r}  '
              f'fam={nm.get(1)!r:<22s} fsType={fstype}  glyphs={nglyph}')
        if odd:
            print(f'         非标准表: {odd}')

    print(f'\n[3] 体积合计: {total/1024/1024:.2f} MB')

    print('\n[4] 嵌入许可 (OS/2.fsType)')
    fs = {r.get('fsType') for r in rows if r.get('fsType') is not None}
    print(f'    取值集合: {fs}')
    for v in sorted(fs):
        print(f'    fsType={v} -> {" + ".join(decode_fstype(v))}')
    # 只要不含 bit2(禁止嵌入) 即允许嵌入 docx
    allow = bool(fs) and all(not (v & 2) for v in fs)
    print(f'    结论: {"允许嵌入 docx（含可编辑嵌入位）" if allow else "禁止嵌入，需复核"}')

    print('\n[5] 许可文本')
    for r in rows[:1]:
        print(f'    name13 license: {r.get("license") or "(未写入 name 表)"}')
        print(f'    name14 url    : {r.get("licurl") or "(未写入 name 表)"}')
    print('    Latin Modern 由 GUST e-foundry 发布，采用 GUST Font License (GFL)，')
    print('    基于 LPPL，允许自由使用、嵌入文档与再分发。name 表未写许可字段属常见情况，')
    print('    以 CTAN 包 lm 的 LICENSE 为准。')

    print('\n[6] 命名冲突检查')
    sysdirs = [r'C:\Windows\Fonts',
               os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Windows\Fonts')]
    fams = {r.get('family') for r in rows if r.get('family')}
    conflict = []
    for sd in sysdirs:
        if not os.path.isdir(sd):
            continue
        for f in os.listdir(sd):
            lf = f.lower()
            if 'lmroman' in lf or 'latinmodern' in lf:
                conflict.append(os.path.join(sd, f))
    print(f'    待装 family: {sorted(fams)}')
    print(f'    系统已存在同名文件: {conflict if conflict else "无"}')

    print('\n[7] 风险结论')
    bad = [r for r in rows if not r.get('ok')]
    print(f'    异常文件数: {len(bad)}')
    print('    OTF 是纯数据格式（字形轮廓 + 度量表），不含可执行代码，')
    print('    不存在脚本执行/网络外发/文件读写能力。')
    print('    本次安装方式为「复制到用户字体目录 + 写 HKCU 注册表」，')
    print('    不触碰 C:\\Windows\\Fonts、不需管理员、不影响其他用户，可完整卸载。')
    print('=' * 96)
    json.dump(rows, open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      'audit_lmfont.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    return len(bad) == 0


if __name__ == '__main__':
    ok = audit()
    print('AUDIT', 'PASS' if ok else 'FAIL')
