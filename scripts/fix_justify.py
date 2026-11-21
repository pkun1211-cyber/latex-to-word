# -*- coding: utf-8 -*-
r"""段落级显式两端对齐。

样式层 jc=both 不足以生效：docDefaults 与 compat 会让 Word 沿用左对齐。
必须在每个正文段的 pPr 上直接写 <w:jc w:val="both"/>。
LaTeX 正文是 justify（每行撑满 453.6pt 版心），左对齐会导致：
  - 中文行右边界参差（515—528pt 波动 vs LaTeX 稳定 524pt）
  - 英文行提前断行，整段多占 1—2 行
"""
import sys, io, os
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import l2wconf as C
import docxio as IO

# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
ORD = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
       'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
       'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
       'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
       'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
       'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
       'sectPr', 'pPrChange']
RANK = {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}' + n: i
        for i, n in enumerate(ORD)}


def sort_ppr(ppr):
    kids = list(ppr)
    want = sorted(kids, key=lambda e: RANK.get(e.tag, 999))
    if [e.tag for e in kids] != [e.tag for e in want]:
        for e in kids:
            ppr.remove(e)
        for e in want:
            ppr.append(e)


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    d = Document(DOCX)
    CENTER = {'Table Caption', 'Image Caption', 'Captioned Figure', 'Title', 'Author'}
    n_both = n_center = 0
    for i, p in enumerate(d.paragraphs):
        el = p._p
        ppr = el.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr'); el.insert(0, ppr)
        jc = ppr.find(qn('w:jc'))
        if jc is None:
            jc = OxmlElement('w:jc'); ppr.append(jc)
        st = p.style.name
        txt = p.text.strip()
        if st in CENTER or i <= 2 or el.find(f'{M}oMathPara') is not None:
            jc.set(qn('w:val'), 'center'); n_center += 1
        elif st in ('Heading 1', 'Heading 2', 'Heading 3'):
            jc.set(qn('w:val'), 'left')
        elif txt.startswith('注：'):
            jc.set(qn('w:val'), 'both'); n_both += 1
        else:
            jc.set(qn('w:val'), 'both'); n_both += 1
        # 关闭中英文/中数字之间的自动加空格。
        # LaTeX(ctex) 不插这类空隙，Word 默认 autoSpaceDE/DN=1 会在每个西文词两侧
        # 加 1/8 em，累积使中文行右边界在 517—528pt 间抖动、断行位置偏移。
        for tag in ('w:autoSpaceDE', 'w:autoSpaceDN'):
            e = ppr.find(qn(tag))
            if e is None:
                e = OxmlElement(tag); ppr.append(e)
            e.set(qn('w:val'), '0')
        sort_ppr(ppr)
    print('两端对齐段:', n_both, '居中/左对齐段:', n_center + 0)
    IO.save(d, DOCX)
    print('SAVED', os.path.getsize(DOCX))


if __name__ == '__main__':
    main()
