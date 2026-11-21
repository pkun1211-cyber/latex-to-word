# -*- coding: utf-8 -*-
r"""复刻 LaTeX 摘要/关键词段的悬挂缩进。

实测 LaTeX：
  ［摘要］首行 x0=66.05，续行 x0=70.87  → 首行左移 4.82pt
  ［关键词］首行 x0=66.05
  正文段首行 x0=88.87（缩进 2em=18pt），续行 70.87

Word 现状：摘要段 firstLine=0，全部行 x0=70.8 → 首行少了 4.82pt 可用宽度，
导致 br 前的末字被挤到下一行。

做法：给摘要/关键词段设 w:ind left=0 firstLine=-96（−4.8pt = −96 twips）
      并把 left 设为 96（4.8pt），使续行仍在 70.87。
      即 left=96, firstLine=-96 → 首行 0、续行 4.8pt？方向相反。

正确换算：Word 的 ind left 是整段左缩进，firstLine 是首行相对 left 的增量。
  目标：首行 = 版心左 − 4.82pt，续行 = 版心左
  → left = 0, firstLine = -96（负值首行外凸），续行 = left = 0
  Word 允许 firstLine 负值（悬挂），但需 hanging 属性表达：
  w:ind w:left="0" w:hanging="96" → 首行左移 4.8pt，续行 0。
"""
import sys, io, os, re, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO

TW = 20  # 1pt = 20 twips


def fix(docx=None, apply=False):
    docx = docx or HA.DOCX
    d = Document(docx)
    hits = []
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        if not (t.startswith('［摘要］') or t.startswith('［关键词］')
                or t.startswith('［基金项目］') or t.startswith('［作者简介］')):
            continue
        ppr = p._p.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr')
            p._p.insert(0, ppr)
        ind = ppr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            ppr.append(ind)
        old = dict(ind.attrib)
        # 悬挂 4.8pt：首行外凸，续行回到版心左界
        for k in ('firstLine', 'firstLineChars', 'hangingChars'):
            ind.attrib.pop(qn('w:' + k), None)
        ind.set(qn('w:left'), '0')
        ind.set(qn('w:leftChars'), '0')
        ind.set(qn('w:hanging'), '96')
        HA.sort_el(ppr, HA.PPR_ORDER)
        hits.append((i, t[:14], old))
    print(f'命中 {len(hits)} 段')
    for i, t, old in hits:
        print(f'  P{i} {t}')
    if not apply:
        print('（未修改，加 --apply）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_hang.docx')
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
    IO.save(d, docx)
    print('已写入悬挂缩进')


if __name__ == '__main__':
    fix(apply='--apply' in sys.argv)
