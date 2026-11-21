# -*- coding: utf-8 -*-
"""基于 Pandoc 官方模板，生成钉死 PDF 版式的 reference.docx。

PDF 基准（小论文 LaTeX 版）：
- A4；上下 2.4cm、左右 2.5cm
- 正文：中文宋体 / 西文 Times New Roman，小五(9pt)，1.5 倍行距
- 一级标题：黑体四号(14pt)；二级标题：黑体五号(10.5pt)
- 题名：黑体三号(16pt)；作者：宋体小四(12pt)
- 表题/图题：宋体小五(9pt) 居中；文献：宋体小五
"""
import os
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import l2wconf as C
import docxio as IO

BASE = C.BASE
SRC = os.path.join(BASE, ".workbuddy", "tmp", "reference_base.docx")
OUT = os.path.join(BASE, ".workbuddy", "tmp", "reference_custom.docx")


def set_font(style, ascii_font, ea_font, size=None, bold=None):
    style.font.name = ascii_font
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.insert(0, rf)
    rf.set(qn('w:ascii'), ascii_font)
    rf.set(qn('w:hAnsi'), ascii_font)
    rf.set(qn('w:eastAsia'), ea_font)
    rf.set(qn('w:cs'), ascii_font)
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold

def set_para(style, line=1.5, before=0, after=0, align=None, keep_next=False):
    pf = style.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if align is not None:
        style.paragraph_format.alignment = align
    if keep_next:
        ppr = style.element.get_or_add_pPr()
        kn = OxmlElement('w:keepNext'); ppr.append(kn)

def sty(doc, sid):
    """按 style_id 取样式。doc 必须显式传入——模块级不得持有 Document 实例，
    否则 import 本模块就会打开文件（早期版本因此损坏过用户成品）。"""
    for s in doc.styles:
        if s.style_id == sid:
            return s
    return None


def main():
    d = Document(SRC)
    # ---------- 正文 ----------
    normal = sty(d, 'Normal')
    set_font(normal, 'Times New Roman', '宋体', 9)
    set_para(normal, line=1.5, before=0, after=0)

    # ---------- 题名 / 作者 / 日期 ----------
    for sid, sz, ea in [('Title', 16, '黑体'), ('Subtitle', 12, '宋体'),
                        ('Author', 12, '宋体'), ('Date', 9, '宋体')]:
        s = sty(d, sid)
        if s:
            set_font(s, 'Times New Roman', ea, sz)
            set_para(s, line=1.5, before=0, after=3)

    # ---------- 摘要 ----------
    for sid in ['AbstractTitle', 'Abstract']:
        s = sty(d, sid)
        if s:
            set_font(s, 'Times New Roman', '宋体', 9, bold=(sid == 'AbstractTitle'))
            set_para(s, line=1.5, before=0, after=0)

    # ---------- 标题层级 ----------
    set_font(sty(d, 'Heading1'), 'Times New Roman', '黑体', 14, bold=False)
    set_para(sty(d, 'Heading1'), line=1.5, before=12, after=6, keep_next=True)
    set_font(sty(d, 'Heading2'), 'Times New Roman', '黑体', 10.5, bold=False)
    set_para(sty(d, 'Heading2'), line=1.5, before=8, after=4, keep_next=True)
    for i in range(3, 10):
        s = sty(d, f'Heading{i}')
        if s:
            set_font(s, 'Times New Roman', '黑体', 10.5)
            set_para(s, line=1.5, before=6, after=3, keep_next=True)

    # ---------- 图表题 ----------
    for sid in ['Caption', 'TableCaption', 'ImageCaption']:
        s = sty(d, sid)
        if s:
            set_font(s, 'Times New Roman', '宋体', 9)
            set_para(s, line=1.5, before=3, after=3)

    # ---------- 正文块 / 参考文献 ----------
    for sid in ['BodyText', 'FirstParagraph', 'Compact', 'BlockText', 'Bibliography']:
        s = sty(d, sid)
        if s:
            set_font(s, 'Times New Roman', '宋体', 9)
            set_para(s, line=1.5, before=0, after=0)

    # ---------- 页面设置 ----------
    sec = d.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.4)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    IO.save(d, OUT)
    print('saved:', OUT, os.path.getsize(OUT))

    # 校验
    d2 = Document(OUT)
    n = d2.styles['Normal']
    print('Normal font:', n.font.name, n.font.size.pt, 'eastAsia=',
          n.element.rPr.rFonts.get(qn('w:eastAsia')))
    s2 = d2.sections[0]
    print('page margins cm:', round(s2.top_margin.cm,2), round(s2.left_margin.cm,2))


if __name__ == '__main__':
    main()
