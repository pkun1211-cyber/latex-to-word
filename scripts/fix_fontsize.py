# -*- coding: utf-8 -*-
r"""西文字体改 LM Roman + 6 处字号错配修正（严格对齐 LaTeX PDF 实测）。

LaTeX 实测字体/字号（scan_fonts_all.py 输出）
--------------------------------------------------------------
SimSun              9.00pt   正文/摘要/表注/文献        → 宋体 9pt   (w:sz=18)
LMRoman9-Regular    9.00pt   正文西文                   → LM Roman 9
CMR9/CMMI9/CMSY9    9.00pt   公式                       → Cambria Math（保留）
SimHei             10.50pt   H2                         → 黑体 10.5pt(w:sz=21)
SimHei             14.00pt   H1                         → 黑体 14pt  (w:sz=28)
SimHei             16.00pt   题名                       → 黑体 16pt  (w:sz=32)
SimHei              9.00pt   ［摘要］［关键词］标签      → 黑体 9pt
SimHei              6.50pt   表格表头                   → 黑体 6.5pt(w:sz=13)
SimSun             14.00pt   作者                       → 宋体 14pt  (w:sz=28)  ★错配1
SimSun              9.00pt   表题/表格/表注（全部 9pt）  → 宋体 9pt
LMRoman6-Regular    5.98pt   上标引文西文               → LM Roman 6 @6pt      ★错配2
LMRoman8-Regular    7.50pt   仅图内 TikZ 节点/URL       → 图是位图，无需处理
LMRoman7-Regular    6.50pt   仅图内 TikZ 节点           → 图是位图，无需处理
LMRoman10-Regular  10.50pt   H2 西文                    → LM Roman 10
LMRoman17-Regular  16.00pt   题名西文                   → LM Roman 17
LMRoman9-Bold       9.00pt   文献编号 [1]               → LM Roman 9 + b
LMRomanDemi10       9.00pt   表内 C 统计量              → LM Roman Demi 10

★ 实测确认的错配（probe_region.py 逐区域探测，不靠读 tex 猜）
1. 作者段：Word 12pt(sz=24) → LaTeX 14pt(sz=28)
2. 上标引文：Word sz=13 + vertAlign=superscript → Word 再缩到 4pt；
   LaTeX \textsuperscript 是 6pt 不再缩。必须去掉 vertAlign，
   改 sz=12(6pt) + w:position 抬升。
3. 西文全体：Times New Roman → LM Roman（按字号选对应光学尺寸族）

【已排除的误判】表题/表格数据/表注在 LaTeX 中均为 9pt（\tbfs=\zihao{-5}=小五9pt），
与 Word 现状一致，无需改动。此前误以为 6.5/7.5pt，实为图内 TikZ 节点字号。

LM Roman 光学尺寸映射（LaTeX 自动按字号选族）
--------------------------------------------------------------
  ≤5.5pt → LM Roman 5      5.5~6.5 → LM Roman 6
  6.5~7.5 → LM Roman 7     7.5~8.5 → LM Roman 8
  8.5~9.5 → LM Roman 9     9.5~11  → LM Roman 10
  11~14   → LM Roman 12    >14     → LM Roman 17
"""
import sys, io, os, re, shutil, collections
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import l2wconf as C
import docxio as IO

DIR = C.DIR
# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
BAK = C.bak('fontsize')
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# rPr 子元素顺序（schema）
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
             'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
             'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
             'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
             'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange']


def sort_rpr(rpr):
    kids = list(rpr)
    for e in kids:
        rpr.remove(e)
    for e in sorted(kids, key=lambda e: RPR_ORDER.index(e.tag.split('}')[1])
                    if e.tag.split('}')[1] in RPR_ORDER else 999):
        rpr.append(e)


def lm_family(pt):
    """按字号选 LM Roman 光学尺寸族（与 LaTeX 的 lm 光学尺寸规则一致）。"""
    if pt <= 5.5:
        return 'LM Roman 5'
    if pt <= 6.5:
        return 'LM Roman 6'
    if pt <= 7.5:
        return 'LM Roman 7'
    if pt <= 8.5:
        return 'LM Roman 8'
    if pt <= 9.5:
        return 'LM Roman 9'
    if pt <= 11.0:
        return 'LM Roman 10'
    if pt <= 14.0:
        return 'LM Roman 12'
    return 'LM Roman 17'


def rpr_of(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        r.insert(0, rpr)
    return rpr


def set_sz(rpr, half_pt):
    for tag in ('w:sz', 'w:szCs'):
        e = rpr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag)
            rpr.append(e)
        e.set(qn('w:val'), str(int(half_pt)))


def get_sz(rpr):
    e = rpr.find(qn('w:sz')) if rpr is not None else None
    return int(e.get(qn('w:val'))) if e is not None else None


def set_fonts(rpr, ascii_, ea=None):
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.insert(0, rf)
    for a in ('asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme'):
        if rf.get(qn('w:' + a)) is not None:
            del rf.attrib[qn('w:' + a)]
    rf.set(qn('w:ascii'), ascii_)
    rf.set(qn('w:hAnsi'), ascii_)
    rf.set(qn('w:cs'), ascii_)
    if ea:
        rf.set(qn('w:eastAsia'), ea)


def alltext(el):
    return ''.join(t.text or '' for t in el.iter() if t.tag in (W + 't', M + 't'))


# ================================================================ 主流程


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    if not os.path.exists(BAK):
        shutil.copy2(DOCX, BAK)
        print('备份 ->', BAK)

    d = Document(DOCX)
    stat = collections.Counter()

    # ---------------------------------------------------------------- 1. 样式层西文改 LM Roman
    STY = d.styles.element
    STYLE_SZ = {}
    for st in STY.findall(qn('w:style')):
        sid = st.get(qn('w:styleId'))
        rpr = st.find(qn('w:rPr'))
        if rpr is None:
            continue
        sz = get_sz(rpr)
        if sz:
            STYLE_SZ[sid] = sz
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            continue
        cur = rf.get(qn('w:ascii'))
        if cur and 'Times' in cur:
            pt = (sz or 18) / 2
            set_fonts(rpr, lm_family(pt))
            stat['style_font'] += 1
    print('[1] 样式层西文字体改 LM Roman:', stat['style_font'])

    # 关键样式字号按 LaTeX 归位
    FIX_STYLE_SZ = {
        'Author': 28,          # 14pt（原 24=12pt）★错配1
        'Title': 32,           # 16pt
        'Heading1': 28,        # 14pt
        'Heading2': 21,        # 10.5pt
        'Normal': 18, 'BodyText': 18, 'FirstParagraph': 18,
        'Abstract': 18, 'Bibliography': 18, 'Compact': 18,
        'Caption': 18, 'TableCaption': 18, 'ImageCaption': 18,
    }
    for sid, sz in FIX_STYLE_SZ.items():
        for st in STY.findall(qn('w:style')):
            if st.get(qn('w:styleId')) != sid:
                continue
            rpr = st.find(qn('w:rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                st.append(rpr)
            set_sz(rpr, sz)
            rf = rpr.find(qn('w:rFonts'))
            if rf is not None:
                set_fonts(rpr, lm_family(sz / 2))
            sort_rpr(rpr)
            STYLE_SZ[sid] = sz
    print('[2] 样式字号归位:', len(FIX_STYLE_SZ))

    # ---------------------------------------------------------------- 2. 逐 run 处理
    ps = d.paragraphs


    def eff_sz(r, para):
        """run 的有效字号（半磅）：run 显式 > 段落样式 > Normal。"""
        rpr = r.find(qn('w:rPr'))
        sz = get_sz(rpr)
        if sz:
            return sz
        sid = None
        ppr = para.find(qn('w:pPr'))
        if ppr is not None:
            e = ppr.find(qn('w:pStyle'))
            if e is not None:
                sid = e.get(qn('w:val'))
        return STYLE_SZ.get(sid, STYLE_SZ.get('Normal', 18))


    def process_para(para, ctx):
        txt = alltext(para).strip()
        is_note = txt.startswith('注：')
        is_bib = bool(re.match(r'^\[\d+\]', txt))
        for r in para.iter(qn('w:r')):
            # 跳过公式内的 run（Cambria Math 由 OMML 管）
            anc = r.getparent()
            in_math = False
            while anc is not None:
                if anc.tag.startswith(M):
                    in_math = True
                    break
                anc = anc.getparent()
            if in_math:
                continue
            rpr = rpr_of(r)
            sz = eff_sz(r, para)
            va = rpr.find(qn('w:vertAlign'))
            is_sup = va is not None and va.get(qn('w:val')) == 'superscript'

            if is_sup:
                # ★错配2：LaTeX \textsuperscript 实际 6pt，Word 会再缩一次
                # 去掉 vertAlign，显式 6pt + position 抬升 (半磅单位)
                rpr.remove(va)
                set_sz(rpr, 12)                      # 6pt
                pos = rpr.find(qn('w:position'))
                if pos is None:
                    pos = OxmlElement('w:position')
                    rpr.append(pos)
                pos.set(qn('w:val'), '7')            # +3.5pt 抬升
                set_fonts(rpr, 'LM Roman 6')
                stat['sup_fix'] += 1
            else:
                # 西文字体按有效字号选 LM 光学尺寸
                set_fonts(rpr, lm_family(sz / 2))
                stat['run_font'] += 1
            sort_rpr(rpr)


    for i, p in enumerate(ps):
        process_para(p._p, f'p{i}')

    # 作者段字号（首页第 2 段）★错配1
    if len(ps) > 1:
        for r in ps[1]._p.iter(qn('w:r')):
            rpr = rpr_of(r)
            set_sz(rpr, 28)
            set_fonts(rpr, 'LM Roman 12', '宋体')
            sort_rpr(rpr)
        stat['author_fix'] += 1

    # ---------------------------------------------------------------- 3. 表格西文字体
    # 实测：LaTeX 表格数据/表头均 9pt（\tbfs=\zihao{-5}），与 Word 现状一致，
    # 故只改西文字体为 LM Roman 9，不动字号、不动中文表头字体。
    for ti, t in enumerate(d.tables):
        for ri, row in enumerate(t.rows):
            for cell in row.cells:
                for cp in cell.paragraphs:
                    for r in cp._p.iter(qn('w:r')):
                        anc = r.getparent()
                        in_math = False
                        while anc is not None:
                            if anc.tag.startswith(M):
                                in_math = True
                                break
                            anc = anc.getparent()
                        if in_math:
                            continue
                        rpr = rpr_of(r)
                        set_fonts(rpr, 'LM Roman 9')
                        sort_rpr(rpr)
                        stat['tbl_font'] += 1

    IO.save(d, DOCX)
    print('[3] 处理统计:', dict(stat))
    print('SAVED', os.path.getsize(DOCX), '段数', len(d.paragraphs))


if __name__ == '__main__':
    main()
