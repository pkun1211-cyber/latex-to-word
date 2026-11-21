# -*- coding: utf-8 -*-
r"""LaTeX→Word 版面精修（统一版，取代 patch_metrics/patch_front/patch_fig）。

全部参数由 LaTeX PDF（pypdfium2 逐行 baseline）实测反推，单位 twips = 1/20 pt。

实测对照（LaTeX → Word 目标）
--------------------------------------------------------------
正文基线距        13.5pt  → line=270 exact
表格数据行距      10.8pt  → line=216 exact
表题→表头         24.5pt  → caption after=220 (11pt)
表末行→表注       17.3pt  → note before=70  (3.5pt)
表注行距          10.8pt  → line=216 exact
表注末→正文/H2    19.6pt  → 下一段 before=120 (6pt)
H1 前            33.5pt  → before=390 (19.5pt) + 正文行盒余量
H1→正文          18.9pt  → after=40   (2pt)
H2 前            24.0pt  → before=200 (10pt)
H2→正文          17.1pt  → after=40   (2pt)
单行公式主行距    34.5pt  → before=200 after=200 (10pt+10pt)
双行公式内部行距  21.0pt  → eqArr 段 line=420 exact
公式→正文        20.3pt
题名→作者        21.7pt  → title line=400 exact after=100
作者→单位        18.8pt  → author line=340 exact after=130
单位→摘要        27.4pt  → affil line=240 exact after=310
摘要→关键词      15.5pt  → after=40
关键词→Abstract  15.5pt  → after=40
Abstract→Keyw    15.5pt  → after=40
Keywords→基金    17.5pt  → after=80
末作者→H1        39.3pt  → after=110 + H1 before
文献标题→首条    45.1pt  → H1(参考文献) after=520 (26pt)
文献条目基线距    17.4pt  → line=348 exact
图上文→图顶       10.5pt  → drawing 段 before=140 auto
图底→图题         10.1pt  → caption before=100 auto
图题→H2          34.5pt  → caption after=120 + H2 before

关键坑
--------------------------------------------------------------
1. docDefaults 的 pPrDefault spacing after=200 会污染所有未显式设 after 的段落
   （H1/H2/表题/首页段全中招）→ 必须清零。
2. 含 w:drawing 的段落严禁 lineRule=exact，否则 inline 图被压到行高、溢出覆盖文字。
3. 段中硬拆：COM InsertBreak 会在段内插入分页符并把段拆成两段，中间留空段。
   本脚本自动检测「前段无句末标点 + 空段 + 后段小写/中文续接」并合并。
4. 标题字体走 majorEastAsia 主题（等线 Light），必须显式改黑体。
5. Caption 样式带 w:i（斜体），LaTeX caption 是正体，须清除。
"""
import sys, io, os, re, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
import l2wconf as C
import docxio as IO

DIR = C.DIR
# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
BAK = C.bak('refine')
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    d = Document(DOCX)
    body = d.element.body


    # ---------------------------------------------------------------- 工具
    def ppr_of(el):
        ppr = el.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr'); el.insert(0, ppr)
        return ppr


    def sp_of(el):
        ppr = ppr_of(el)
        sp = ppr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); ppr.append(sp)
        return sp


    def set_sp(el, line=None, rule=None, before=None, after=None):
        sp = sp_of(el)
        if line is not None:
            sp.set(qn('w:line'), str(line))
        if rule is not None:
            sp.set(qn('w:lineRule'), rule)
        if before is not None:
            sp.set(qn('w:before'), str(before)); sp.set(qn('w:beforeAutospacing'), '0')
        if after is not None:
            sp.set(qn('w:after'), str(after)); sp.set(qn('w:afterAutospacing'), '0')


    def set_ind(el, chars=2, pt=9):
        ppr = ppr_of(el)
        ind = ppr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind'); ppr.append(ind)
        ind.set(qn('w:firstLineChars'), str(chars * 100))
        ind.set(qn('w:firstLine'), str(int(chars * pt * 20)))
        for a in ('w:left', 'w:leftChars', 'w:right', 'w:rightChars', 'w:hanging', 'w:hangingChars'):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]


    def clear_ind(el):
        ppr = ppr_of(el)
        ind = ppr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind'); ppr.append(ind)
        for a in ('w:firstLine', 'w:firstLineChars', 'w:left', 'w:leftChars',
                  'w:right', 'w:rightChars', 'w:hanging', 'w:hangingChars'):
            if ind.get(qn(a)) is not None:
                del ind.attrib[qn(a)]
        ind.set(qn('w:firstLine'), '0'); ind.set(qn('w:firstLineChars'), '0')


    def alltext(el):
        W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        return ''.join(t.text or '' for t in el.iter() if t.tag in (W + 't', M + 't'))


    def has_math(el):
        return el.find(f'.//{M}oMath') is not None


    def has_disp(el):
        return el.find(f'{M}oMathPara') is not None


    def has_img(el):
        return bool(el.findall('.//' + qn('w:drawing')))


    # ---------------------------------------------------------------- 0. docDefaults 清零
    dd = d.styles.element.find(qn('w:docDefaults'))
    if dd is not None:
        pd = dd.find(qn('w:pPrDefault'))
        if pd is not None:
            pp = pd.find(qn('w:pPr'))
            if pp is not None:
                sp = pp.find(qn('w:spacing'))
                if sp is not None:
                    sp.set(qn('w:after'), '0'); sp.set(qn('w:before'), '0')
                    sp.set(qn('w:line'), '270'); sp.set(qn('w:lineRule'), 'exact')
                    print('[0] docDefaults spacing 归零')

    # ---------------------------------------------------------------- 1. 样式层
    STY = d.styles.element


    def style_by_id(sid):
        for st in STY.findall(qn('w:style')):
            if st.get(qn('w:styleId')) == sid:
                return st
        return None


    def style_rpr(st):
        rpr = st.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr'); st.append(rpr)
        return rpr


    def set_style_font(st, ea, ascii_, sz):
        rpr = style_rpr(st)
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
        for a in ('asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme'):
            if rf.get(qn('w:' + a)) is not None:
                del rf.attrib[qn('w:' + a)]
        rf.set(qn('w:ascii'), ascii_); rf.set(qn('w:hAnsi'), ascii_)
        rf.set(qn('w:eastAsia'), ea); rf.set(qn('w:cs'), ascii_)
        for tag in ('w:sz', 'w:szCs'):
            e = rpr.find(qn(tag))
            if e is None:
                e = OxmlElement(tag); rpr.append(e)
            e.set(qn('w:val'), str(sz))


    def set_style_sp(sid, line=None, rule=None, before=None, after=None):
        st = style_by_id(sid)
        if st is None:
            return
        ppr = st.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr'); st.append(ppr)
        sp = ppr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); ppr.append(sp)
        if line is not None:
            sp.set(qn('w:line'), str(line))
        if rule is not None:
            sp.set(qn('w:lineRule'), rule)
        if before is not None:
            sp.set(qn('w:before'), str(before)); sp.set(qn('w:beforeAutospacing'), '0')
        if after is not None:
            sp.set(qn('w:after'), str(after)); sp.set(qn('w:afterAutospacing'), '0')


    # Normal：宋体 9pt / 13.5pt 固定行距 / 两端对齐
    # 【关键】LaTeX 正文是 justify，每行撑满版心。Word 默认左对齐会让英文行提前断行、
    # 中文行右边界参差，进而整页少装/多装行。必须在 Normal 样式上显式写 jc=both。
    set_style_sp('Normal', line=270, rule='exact', before=0, after=0)
    _st = style_by_id('Normal')
    _ppr = _st.find(qn('w:pPr'))
    _jc = _ppr.find(qn('w:jc'))
    if _jc is None:
        _jc = OxmlElement('w:jc'); _ppr.append(_jc)
    _jc.set(qn('w:val'), 'both')
    # pPr 子元素顺序：jc 必须在 spacing/ind 之后
    _ord = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'numPr',
            'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'jc', 'outlineLvl', 'rPr']
    _kids = list(_ppr)
    for _e in _kids:
        _ppr.remove(_e)
    for _e in sorted(_kids, key=lambda e: _ord.index(e.tag.split('}')[1])
                     if e.tag.split('}')[1] in _ord else 99):
        _ppr.append(_e)
    # H1：黑体四号(14pt=28)，前 19.5pt 后 2pt
    st = style_by_id('Heading1')
    set_style_font(st, '黑体', 'Times New Roman', 28)
    set_style_sp('Heading1', line=340, rule='exact', before=390, after=40)
    # H2：黑体五号(10.5pt=21)，前 10pt 后 2pt
    st = style_by_id('Heading2')
    set_style_font(st, '黑体', 'Times New Roman', 21)
    set_style_sp('Heading2', line=260, rule='exact', before=200, after=40)
    # Caption 系：宋体小五(9pt)，去斜体
    for sid in ('Caption', 'TableCaption', 'ImageCaption'):
        st = style_by_id(sid)
        if st is None:
            continue
        rpr = style_rpr(st)
        for it in rpr.findall(qn('w:i')) + rpr.findall(qn('w:iCs')):
            rpr.remove(it)
        set_style_font(st, '宋体', 'Times New Roman', 18)
    set_style_sp('Caption', line=240, rule='exact', before=0, after=0)
    set_style_sp('TableCaption', line=240, rule='exact', before=100, after=220)
    set_style_sp('ImageCaption', line=240, rule='auto', before=100, after=120)
    print('[1] 样式层：Normal/H1/H2/Caption 字体与间距已对齐')

    # ---------------------------------------------------------------- 2. 清除全部分页符
    nclr = 0
    for p in d.paragraphs:
        ppr = p._p.find(qn('w:pPr'))
        if ppr is not None:
            for pb in ppr.findall(qn('w:pageBreakBefore')):
                ppr.remove(pb); nclr += 1
        for r in p._p.iter(qn('w:r')):
            for br in r.findall(qn('w:br')):
                if br.get(qn('w:type')) == 'page':
                    r.remove(br); nclr += 1
    print('[2] 清除分页符:', nclr)

    # ---------------------------------------------------------------- 3. 合并被硬拆的段落
    END_PUNCT = '。？！；：”’）)】》.!?'


    def merge_split_paras():
        merged = 0
        while True:
            ps = d.paragraphs
            hit = None
            for i in range(len(ps) - 2):
                a, mid, b = ps[i], ps[i + 1], ps[i + 2]
                if mid.text.strip():
                    continue
                if mid._p.find(f'.//{M}oMath') is not None or has_img(mid._p):
                    continue
                if a.style.name not in ('Body Text', 'First Paragraph'):
                    continue
                if b.style.name not in ('Body Text', 'First Paragraph'):
                    continue
                ta = alltext(a._p).rstrip()
                tb = alltext(b._p).lstrip()
                if not ta or not tb:
                    continue
                if ta[-1] in END_PUNCT:
                    continue
                # 后段以小写/中文/数字续接，且不是新句起头标志
                if re.match(r'^[（(【\[]', tb) or re.match(r'^\d+[.、]', tb):
                    continue
                hit = (i, a, mid, b)
                break
            if hit is None:
                break
            i, a, mid, b = hit
            # 把 b 的内容（除 pPr）搬到 a 末尾
            for ch in list(b._p):
                if ch.tag == qn('w:pPr'):
                    continue
                a._p.append(ch)
            body.remove(mid._p)
            body.remove(b._p)
            merged += 1
            print('    合并段 %d+%d+%d → %s…' % (i, i + 1, i + 2, alltext(a._p)[:24]))
        return merged


    nmerge = merge_split_paras()
    print('[3] 合并硬拆段落:', nmerge)

    # ---------------------------------------------------------------- 4. 首页度量
    ps = d.paragraphs


    def front(i, line, rule, before, after, indent=False):
        el = ps[i]._p
        set_sp(el, line=line, rule=rule, before=before, after=after)
        if indent:
            set_ind(el, 2, 9)
        else:
            clear_ind(el)


    front(0, 400, 'exact', 0, 100)     # 题名 16pt 黑体
    front(1, 340, 'exact', 0, 130)     # 作者 12pt
    front(2, 240, 'exact', 0, 310)     # 单位 9pt
    front(3, 270, 'exact', 0, 40)      # 摘要
    front(4, 270, 'exact', 0, 40)      # 关键词
    front(5, 270, 'exact', 0, 40)      # Abstract
    front(6, 270, 'exact', 0, 80)      # Keywords
    for i in (7, 8, 9, 10, 11):
        front(i, 270, 'exact', 0, 0)   # 基金/中图/作者简介
    front(12, 270, 'exact', 0, 110)    # 末位作者
    print('[4] 首页 13 段度量已设定')

    # ---------------------------------------------------------------- 5. 逐块度量
    ps = d.paragraphs
    n_body = n_note = n_eq1 = n_eq2 = n_bib = n_img = 0
    for i, p in enumerate(ps):
        el = p._p
        style = p.style.name
        txt = alltext(el).strip()

        if has_img(el):
            # inline 图：必须 auto，否则被行距压扁
            set_sp(el, line=240, rule='auto', before=140, after=0)
            clear_ind(el)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            n_img += 1
            continue
        if style == 'Image Caption':
            set_sp(el, line=240, rule='auto', before=100, after=120)
            clear_ind(el)
            continue
        if style == 'Table Caption':
            set_sp(el, line=240, rule='exact', before=100, after=220)
            clear_ind(el)
            continue
        if style in ('Heading 1', 'Heading 2'):
            # 参考文献标题：LaTeX 标题→首条 45.1pt，比常规 H1 大
            if style == 'Heading 1' and '参考文献' in txt:
                set_sp(el, line=340, rule='exact', before=390, after=520)
            else:
                set_sp(el, line=340 if style == 'Heading 1' else 260, rule='exact',
                       before=390 if style == 'Heading 1' else 200, after=40)
            clear_ind(el)
            continue
        if i <= 12:
            continue
        if has_disp(el):
            omp = el.find(f'{M}oMathPara')
            narr = len(omp.findall(f'.//{M}eqArr'))
            if narr > 0:
                set_sp(el, line=420, rule='exact', before=140, after=140)
                n_eq2 += 1
            else:
                set_sp(el, line=270, rule='exact', before=200, after=200)
                n_eq1 += 1
            clear_ind(el)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            continue
        if txt.startswith('注：'):
            set_sp(el, line=216, rule='exact', before=70, after=0)
            clear_ind(el)
            n_note += 1
            continue
        if re.match(r'^\[\d+\]', txt):
            set_sp(el, line=348, rule='exact', before=0, after=0)
            ppr = ppr_of(el)
            ind = ppr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind'); ppr.append(ind)
            ind.set(qn('w:left'), '400'); ind.set(qn('w:hanging'), '400')
            for a in ('w:firstLine', 'w:firstLineChars'):
                if ind.get(qn(a)) is not None:
                    del ind.attrib[qn(a)]
            n_bib += 1
            continue
        if not txt:
            set_sp(el, line=60, rule='exact', before=0, after=0)
            continue
        # 正文
        set_sp(el, line=270, rule='exact', before=0, after=0)
        set_ind(el, 2, 9)
        n_body += 1

    print('[5] 正文:%d 表注:%d 单行式:%d 双行式:%d 文献:%d 图:%d' %
          (n_body, n_note, n_eq1, n_eq2, n_bib, n_img))

    # ---------------------------------------------------------------- 6. 表注后续段 before 6pt
    ps = d.paragraphs
    n_after = 0
    for i, p in enumerate(ps):
        if not alltext(p._p).strip().startswith('注：'):
            continue
        for j in range(i + 1, min(i + 3, len(ps))):
            q = ps[j]
            if not alltext(q._p).strip():
                continue
            if q.style.name in ('Heading 1', 'Heading 2', 'Table Caption'):
                break
            sp = sp_of(q._p)
            cur = int(sp.get(qn('w:before')) or 0)
            if cur < 120:
                set_sp(q._p, before=120)
                n_after += 1
            break
    print('[6] 表注后正文段 before=6pt:', n_after)

    # ---------------------------------------------------------------- 7. 表格单元格
    ncell = 0
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                for cp in cell.paragraphs:
                    set_sp(cp._p, line=216, rule='exact', before=0, after=0)
                    clear_ind(cp._p)
                    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    ncell += 1
    print('[7] 表格单元格段落:', ncell)

    # ---------------------------------------------------------------- 8. 表后空段压缩
    children = list(body)
    for i, el in enumerate(children):
        if el.tag != qn('w:tbl'):
            continue
        if i + 1 < len(children):
            nxt = children[i + 1]
            if nxt.tag == qn('w:p') and not alltext(nxt).strip() and not has_math(nxt):
                set_sp(nxt, line=60, rule='exact', before=0, after=0)

    IO.save(d, DOCX)
    print('SAVED', os.path.getsize(DOCX), '段数', len(d.paragraphs))


if __name__ == '__main__':
    main()
