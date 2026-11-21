# -*- coding: utf-8 -*-
"""11.20 版 Pandoc 产物 OOXML 后处理。

1. 文献段拆 44 条 + 书签 _RefN
2. ⟦CITE:n,n⟧ -> 上标平级 hyperlink（严禁嵌套）
3. ⟦REF:表N⟧ -> 普通文本
4. ⟦TABNOTE⟧ -> 「注：」小五段落，紧贴表格之后
5. ⟦FOOTNOTE:n⟧ -> 页脚脚注（真实 w:footnoteReference）
6. 11 张三线表：顶/底 12、栏目线 6，单元格居中
7. 标题中文编号 一、/（一），章内重置
8. 表题「表N　」前缀、图题「图1　」
9. 正文首行缩进 2em、1.5 倍行距、页边距
"""
import os, re, json, sys
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
import l2wconf as C
import docxio as IO

BASE = C.BASE
DIR = C.DIR
SRC = C.tmp("pandoc_raw.docx")   # Pandoc 直出，未后处理


def make_run(text, size=9, bold=False, superscript=False, ea='宋体'):
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Times New Roman'); rf.set(qn('w:hAnsi'), 'Times New Roman')
    rf.set(qn('w:eastAsia'), ea); rf.set(qn('w:cs'), 'Times New Roman')
    rpr.append(rf)
    if bold:
        rpr.append(OxmlElement('w:b'))
    for tag in ('w:sz', 'w:szCs'):
        e = OxmlElement(tag); e.set(qn('w:val'), str(int(size * 2))); rpr.append(e)
    if superscript:
        va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rpr.append(va)
    r.append(rpr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t)
    return r


def make_hyperlink(text, anchor, size=6.5):
    """平级 hyperlink，绝不嵌套。"""
    hl = OxmlElement('w:hyperlink'); hl.set(qn('w:anchor'), anchor)
    hl.append(make_run(text, size=size, superscript=True))
    return hl


def set_indent(para, em=2, size=9):
    ppr = para._p.get_or_add_pPr()
    ind = ppr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind'); ppr.append(ind)
    ind.set(qn('w:firstLineChars'), str(em * 100))
    ind.set(qn('w:firstLine'), str(int(em * size * 20)))


def set_spacing(para, multiple=1.5):
    """兼容旧调用；实际按 LaTeX 实测固定行距 13.5pt = 270 twips exact。"""
    ppr = para._p.get_or_add_pPr()
    sp = ppr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing'); ppr.append(sp)
    sp.set(qn('w:line'), '270'); sp.set(qn('w:lineRule'), 'exact')
    sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')


def main():
    OUT = C.require_docx()
    d = Document(SRC)
    body = d.element.body

    # ============ 1. 文献拆分 ============
    bib_para = next((p for p in d.paragraphs if '⟦BIB:' in p.text), None)
    if bib_para is None:
        raise RuntimeError('未找到文献段落')
    full_text = ''.join(r.text for r in bib_para.runs)
    parts = re.split(r'⟦BIB:(\d+)⟧\s*', full_text)
    entries = [(int(parts[i]), parts[i + 1].strip()) for i in range(1, len(parts), 2)]
    print('文献条目:', len(entries))

    idx = list(body).index(bib_para._p)
    body.remove(bib_para._p)
    bid = 1000
    for num, content in entries:
        np = OxmlElement('w:p')
        ppr = OxmlElement('w:pPr')
        sp = OxmlElement('w:spacing'); sp.set(qn('w:line'), '348'); sp.set(qn('w:lineRule'), 'exact')
        sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0'); ppr.append(sp)
        ind = OxmlElement('w:ind'); ind.set(qn('w:left'), '400'); ind.set(qn('w:hanging'), '400')
        ppr.append(ind)
        np.append(ppr)
        bs = OxmlElement('w:bookmarkStart'); bs.set(qn('w:id'), str(bid)); bs.set(qn('w:name'), f'_Ref{num}')
        np.append(bs)
        np.append(make_run(f'[{num}] ', size=9))
        if content:
            np.append(make_run(content, size=9))
        be = OxmlElement('w:bookmarkEnd'); be.set(qn('w:id'), str(bid)); np.append(be)
        body.insert(idx, np); idx += 1; bid += 1

    # ============ 2~3. CITE / REF ============
    PAT = r'⟦(CITE|REF):([^⟧]+)⟧'
    for para in list(d.paragraphs):
        if '⟦CITE:' not in para.text and '⟦REF:' not in para.text:
            continue
        for r in list(para._p.findall(qn('w:r'))):
            tn = r.find(qn('w:t'))
            if tn is None or not tn.text or '⟦' not in tn.text:
                continue
            txt = tn.text
            if '⟦CITE:' not in txt and '⟦REF:' not in txt:
                continue
            parent = r.getparent(); ridx = list(parent).index(r); parent.remove(r)
            nodes, pos = [], 0
            for m in re.finditer(PAT, txt):
                if m.start() > pos:
                    nodes.append(make_run(txt[pos:m.start()]))
                kind, val = m.group(1), m.group(2)
                if kind == 'CITE':
                    nums = [int(x) for x in val.split(',')]
                    nodes.append(make_run('[', size=6.5, superscript=True))
                    for j, n in enumerate(nums):
                        nodes.append(make_hyperlink(str(n), f'_Ref{n}'))
                        if j < len(nums) - 1:
                            nodes.append(make_run(',', size=6.5, superscript=True))
                    nodes.append(make_run(']', size=6.5, superscript=True))
                else:
                    # val 形如 "表|3"；若前文已出现「表/式/图」字，只补编号，避免「表表3」
                    typ, _, num = val.partition('|')
                    preceding = ''.join(
                        (nd.find(qn('w:t')).text or '') if nd.find(qn('w:t')) is not None else ''
                        for nd in nodes)
                    tail = preceding.rstrip()
                    if typ and (tail.endswith(typ) or (typ == '式' and tail.endswith('（'))):
                        nodes.append(make_run(num))
                    else:
                        nodes.append(make_run(typ + num))
                pos = m.end()
            if pos < len(txt):
                nodes.append(make_run(txt[pos:]))
            for j, nd in enumerate(nodes):
                parent.insert(ridx + j, nd)

    # ============ 4. 表注归位 ============
    note_paras = [p for p in d.paragraphs if p.text.strip().startswith('⟦TABNOTE⟧')]
    print('表注段:', len(note_paras))
    for p in note_paras:
        # 【关键】只把首个含占位符的 w:t 就地改写为「注：」，绝不整段重建。
        # 整段重建会丢掉 m:oMath 的相对位置（p.text 不含 m:t），
        # 导致段内全部行内公式被挤到段首、正文缺失公式占位。
        done = False
        for r in p._p.findall(qn('w:r')):
            tn = r.find(qn('w:t'))
            if tn is None or not tn.text or '⟦TABNOTE⟧' not in tn.text:
                continue
            tn.text = tn.text.replace('⟦TABNOTE⟧', '注：', 1)
            tn.set(qn('xml:space'), 'preserve')
            done = True
            break
        if not done:
            # 占位符被 Pandoc 拆到多个 run：逐 run 清除占位符字符后在最前插「注：」
            for r in p._p.findall(qn('w:r')):
                tn = r.find(qn('w:t'))
                if tn is not None and tn.text:
                    tn.text = tn.text.replace('⟦TABNOTE⟧', '')
            first = p._p.find(qn('w:r'))
            node = make_run('注：', size=9)
            if first is not None:
                first.addprevious(node)
            else:
                p._p.append(node)
        # 统一表注字号为小五
        for r in p._p.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr'); r.insert(0, rpr)
            for tag in ('w:sz', 'w:szCs'):
                e = rpr.find(qn(tag))
                if e is None:
                    e = OxmlElement(tag); rpr.append(e)
                e.set(qn('w:val'), '18')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        ppr = p._p.get_or_add_pPr()
        ind = ppr.find(qn('w:ind'))
        if ind is not None:
            ppr.remove(ind)
        ind = OxmlElement('w:ind'); ind.set(qn('w:firstLine'), '0'); ind.set(qn('w:firstLineChars'), '0')
        ppr.append(ind)
        set_spacing(p, 1.15)

    # ============ 5. 脚注（真实 footnote） ============
    foots = json.load(open(C.tmp('footnotes.json'), encoding='utf-8'))
    from lxml import etree
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    fpart = next(rel.target_part for rel in d.part.rels.values()
                 if rel.reltype.endswith('/footnotes'))   # 复用已有 part，避免 zip 重名
    froot = etree.fromstring(fpart.blob)
    # 清掉模板里可能存在的示例脚注，仅保留 separator / continuationSeparator
    for fn in list(froot):
        if fn.get(f'{{{W}}}type') is None:
            froot.remove(fn)

    def save_froot():
        fpart._blob = etree.tostring(froot, xml_declaration=True, encoding='UTF-8', standalone=True)

    fid = 1
    for para in list(d.paragraphs):
        for r in list(para._p.findall(qn('w:r'))):
            tn = r.find(qn('w:t'))
            if tn is None or not tn.text or '⟦FOOTNOTE:' not in tn.text:
                continue
            txt = tn.text
            parent = r.getparent(); ridx = list(parent).index(r); parent.remove(r)
            nodes, pos = [], 0
            for m in re.finditer(r'⟦FOOTNOTE:(\d+)⟧', txt):
                if m.start() > pos:
                    nodes.append(make_run(txt[pos:m.start()]))
                n = int(m.group(1))
                content = foots[n - 1]
                # 建 footnote 节点
                fn = etree.SubElement(froot, f'{{{W}}}footnote')
                fn.set(f'{{{W}}}id', str(fid))
                fp = etree.SubElement(fn, f'{{{W}}}p')
                fppr = etree.SubElement(fp, f'{{{W}}}pPr')
                fst = etree.SubElement(fppr, f'{{{W}}}pStyle'); fst.set(f'{{{W}}}val', 'FootnoteText')
                fr = etree.SubElement(fp, f'{{{W}}}r')
                frpr = etree.SubElement(fr, f'{{{W}}}rPr')
                va = etree.SubElement(frpr, f'{{{W}}}vertAlign'); va.set(f'{{{W}}}val', 'superscript')
                etree.SubElement(fr, f'{{{W}}}footnoteRef')
                fr2 = etree.SubElement(fp, f'{{{W}}}r')
                frpr2 = etree.SubElement(fr2, f'{{{W}}}rPr')
                rf = etree.SubElement(frpr2, f'{{{W}}}rFonts')
                rf.set(f'{{{W}}}ascii', 'Times New Roman'); rf.set(f'{{{W}}}hAnsi', 'Times New Roman')
                rf.set(f'{{{W}}}eastAsia', '宋体')
                sz = etree.SubElement(frpr2, f'{{{W}}}sz'); sz.set(f'{{{W}}}val', '16')
                ft = etree.SubElement(fr2, f'{{{W}}}t')
                ft.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                ft.text = ' ' + content
                # 正文引用
                rr = OxmlElement('w:r')
                rrpr = OxmlElement('w:rPr')
                v = OxmlElement('w:vertAlign'); v.set(qn('w:val'), 'superscript'); rrpr.append(v)
                rr.append(rrpr)
                ref = OxmlElement('w:footnoteReference'); ref.set(qn('w:id'), str(fid))
                rr.append(ref)
                nodes.append(rr)
                fid += 1
                pos = m.end()
            if pos < len(txt):
                nodes.append(make_run(txt[pos:]))
            for j, nd in enumerate(nodes):
                parent.insert(ridx + j, nd)
    save_froot()
    print('脚注写入:', fid - 1)

    # ============ 6. 三线表 ============
    # 按 tex 原始列规格分配列宽（Pandoc 默认等分，会让 p{} 长文本列竖排）
    TEXTW = 9072  # 版心宽 twips：(21 - 2.5*2) cm
    COLSPEC = [
        [0.60, 0.20, 0.20],                       # 表1 p{.60}cc
        [0.10, 0.26, 0.10, 0.11, 0.11, 0.11, 0.11, 0.10],  # 表2 llcccccc
        [0.40, 0.30, 0.30],                       # 表3 lcc
        [0.34, 0.20, 0.28, 0.18],                 # 表4 lccc
        [0.09, 0.25, 0.24, 0.24, 0.18],           # 表5 llccc
        [0.07, 0.20, 0.11, 0.21, 0.21, 0.20],     # 表6 lllccc
        [0.32, 0.17, 0.19, 0.14, 0.18],           # 表7 lcccc
        [0.32, 0.17, 0.17, 0.17, 0.17],           # 表8 lcccc
        [0.40, 0.20, 0.20, 0.20],                 # 表9 lccc
        [0.13, 0.30, 0.24, 0.33],                 # 表10 p{.13}p{.28}p{.22}p{.28}
        [0.16, 0.25, 0.18, 0.41],                 # 表11 p{.16}p{.24}p{.17}p{.34}
    ]

    for ti, t in enumerate(d.tables):
        tblPr = t._element.tblPr
        # 固定布局 + 满版心宽
        for tag in ('w:tblLayout', 'w:tblW'):
            old = tblPr.find(qn(tag))
            if old is not None:
                tblPr.remove(old)
        lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
        tw = OxmlElement('w:tblW'); tw.set(qn('w:w'), str(TEXTW)); tw.set(qn('w:type'), 'dxa')
        tblPr.append(tw)
        # 列宽
        spec = COLSPEC[ti] if ti < len(COLSPEC) else None
        grid = t._element.find(qn('w:tblGrid'))
        ncol = len(grid.findall(qn('w:gridCol'))) if grid is not None else 0
        if spec and len(spec) == ncol:
            widths = [int(TEXTW * f) for f in spec]
            widths[-1] += TEXTW - sum(widths)
            for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
                gc.set(qn('w:w'), str(w))
            for row in t.rows:
                cells = row._tr.findall(qn('w:tc'))
                for cell, w in zip(cells, widths):
                    tcPr = cell.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = OxmlElement('w:tcPr'); cell.insert(0, tcPr)
                    cw = tcPr.find(qn('w:tcW'))
                    if cw is None:
                        cw = OxmlElement('w:tcW'); tcPr.append(cw)
                    cw.set(qn('w:w'), str(w)); cw.set(qn('w:type'), 'dxa')
        # 行不跨页断行；表头重复
        for ri, row in enumerate(t.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.append(OxmlElement('w:cantSplit'))
            if ri == 0 and trPr.find(qn('w:tblHeader')) is None:
                trPr.append(OxmlElement('w:tblHeader'))
        old = tblPr.find(qn('w:tblBorders'))
        if old is not None:
            tblPr.remove(old)
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            e = OxmlElement(f'w:{edge}')
            if edge in ('top', 'bottom'):
                e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), '12'); e.set(qn('w:color'), '000000')
            else:
                e.set(qn('w:val'), 'nil')
            borders.append(e)
        tblPr.append(borders)
        jc = tblPr.find(qn('w:jc'))
        if jc is None:
            jc = OxmlElement('w:jc'); tblPr.append(jc)
        jc.set(qn('w:val'), 'center')
        if t.rows:
            for cell in t.rows[0].cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcb = tcPr.find(qn('w:tcBorders'))
                if tcb is None:
                    tcb = OxmlElement('w:tcBorders'); tcPr.append(tcb)
                bot = OxmlElement('w:bottom')
                bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6'); bot.set(qn('w:color'), '000000')
                tcb.append(bot)
        for row in t.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                va = tcPr.find(qn('w:vAlign'))
                if va is None:
                    va = OxmlElement('w:vAlign'); tcPr.append(va)
                va.set(qn('w:val'), 'center')
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    ppr = para._p.get_or_add_pPr()
                    sp = ppr.find(qn('w:spacing'))
                    if sp is None:
                        sp = OxmlElement('w:spacing'); ppr.append(sp)
                    sp.set(qn('w:line'), '216'); sp.set(qn('w:lineRule'), 'exact')
                    sp.set(qn('w:before'), '0'); sp.set(qn('w:after'), '0')
                    ind = ppr.find(qn('w:ind'))
                    if ind is not None:
                        ppr.remove(ind)
                    for r in para.runs:
                        rpr = r._element.get_or_add_rPr()
                        for tag in ('w:sz', 'w:szCs'):
                            e = rpr.find(qn(tag))
                            if e is None:
                                e = OxmlElement(tag); rpr.append(e)
                            e.set(qn('w:val'), '18')
    print('三线表:', len(d.tables))

    # ============ 7. 标题中文编号 ============
    cn = '一二三四五六七八九十'
    h1 = h2 = 0
    for para in d.paragraphs:
        st = para.style.name
        if st == 'Heading 1':
            h1 += 1; h2 = 0
            if not re.match(r'^[一二三四五六七八九十]+、', para.text):
                pre = (cn[h1 - 1] if h1 <= 10 else str(h1)) + '、'
                if para.runs:
                    para.runs[0].text = pre + para.runs[0].text
        elif st == 'Heading 2':
            h2 += 1
            if not re.match(r'^（[一二三四五六七八九十]+）', para.text):
                pre = '（' + (cn[h2 - 1] if h2 <= 10 else str(h2)) + '）'
                if para.runs:
                    para.runs[0].text = pre + para.runs[0].text
    print('H1:', h1, 'H2总数:', sum(1 for p in d.paragraphs if p.style.name == 'Heading 2'))

    # ============ 8. 表题/图题 ============
    ti = 0
    fi = 0
    for para in d.paragraphs:
        if para.style.name == 'Table Caption':
            ti += 1
            if not re.match(r'^表\s*\d', para.text):
                if para.runs:
                    para.runs[0].text = f'表{ti}　' + para.runs[0].text
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif para.style.name == 'Image Caption':
            fi += 1
            if not re.match(r'^图\s*\d', para.text):
                if para.runs:
                    para.runs[0].text = f'图{fi}　' + para.runs[0].text
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    print('表题:', ti)

    # ============ 9. 正文缩进/行距（LaTeX 实测基线距 13.5pt = 270 twips exact） ============
    BODY_LINE = 1.125  # 保留参数名兼容；set_spacing 内部已改为 exact 270
    skip = {'Heading 1', 'Heading 2', 'Heading 3', 'Table Caption', 'Image Caption',
            'Bibliography', 'Title', 'Subtitle', 'Author', 'Date', 'Captioned Figure'}
    for para in d.paragraphs:
        if para.style.name in skip:
            continue
        txt = para.text.strip()
        if not txt:
            continue
        if txt.startswith('注：'):
            continue
        # 首页元信息段不缩进：摘要/关键词/基金/中图分类/作者简介，
        # 以及「姓名（单位…）」形式的作者行。用通用判据而非硬编码人名。
        if txt.startswith(('［摘', '［关键词', 'Abstract', 'Keywords', '［基金',
                           '［中图', '［作者')) or re.match(r'^[\u4e00-\u9fff]{2,4}（', txt):
            set_spacing(para, BODY_LINE)
            continue
        if re.match(r'^\[\d+\]', txt):
            continue
        set_indent(para, 2, 9)
        set_spacing(para, BODY_LINE)
    # 杂散宽度参数段（thebibliography 残留）清除
    for para in list(d.paragraphs):
        if para.text.strip() in ('99', '44') and para.style.name in ('First Paragraph', 'Body Text'):
            body.remove(para._p)

    # ============ 10. 页边距 ============
    sec = d.sections[0]
    _m = C._C.get('margins_cm', {})
    sec.top_margin = Cm(_m.get('top', 2.4)); sec.bottom_margin = Cm(_m.get('bottom', 2.4))
    sec.left_margin = Cm(_m.get('left', 2.5)); sec.right_margin = Cm(_m.get('right', 2.5))

    IO.save(d, OUT)
    print('SAVED:', OUT, os.path.getsize(OUT))

    # ============ 11. 版面精修链 ============
    # 本段只跑**通用**精修脚本。项目专属的首页拆行补丁（如把题名拆成三行）
    # 应由调用方在跑本脚本之前完成，因为它会改变段落索引。
    import subprocess
    PY = sys.executable
    SCRIPTS = os.path.dirname(os.path.abspath(__file__))
    # 顺序不可调换：
    #   refine_layout 按索引设首页度量 + 全文间距（依赖上游已定稿的段落索引）
    #   fix_settings  pPr 子元素排序 + snapToGrid=0 + characterSpacingControl
    #   fix_justify   段落级 jc=both + autoSpaceDE/DN=0
    #   fix_fontsize  西文改 LM Roman（按字号选光学尺寸族）+ 上标字号
    #   fix_lig2      连字序列替换为 U+FB00..FB04（TeX 默认启用连字）
    #   fix_ns        补 w14/w15/mc 命名空间声明（必须在 fix_lig2 之后）
    CHAIN = ('refine_layout.py', 'fix_settings.py', 'fix_justify.py',
             'fix_fontsize.py', 'fix_lig2.py', 'fix_ns.py')
    for name in CHAIN:
        path = os.path.join(SCRIPTS, name)
        if not os.path.exists(path):
            print('缺少', name); continue
        r = subprocess.run([PY, path], capture_output=True, text=True, encoding='utf-8')
        print('--- %s ---' % name)
        print(r.stdout or '')
        if r.returncode != 0:
            print('%s 失败:' % name, r.stderr)

    print('\n后处理完成。接下来按需运行硬对齐链（见 SKILL.md 第 4 步）。')


if __name__ == '__main__':
    main()
