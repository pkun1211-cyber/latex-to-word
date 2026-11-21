# -*- coding: utf-8 -*-
r"""复刻 LaTeX 参考文献区的行距结构。

实测（LaTeX PDF）：
  文献条目**内部**行距 = 13.5pt   （\zihao{-5} 9pt × 1.5 = 13.5）
  文献条目**之间**行距 = 17.49pt  （13.5 + \itemsep ≈ 4pt）
Word 现状：全部 17.4pt（exact 348）→ 每条多行文献都被拉高，
累积后文献区整体后移一页（p12 多 28 行、p13 溢出）。

做法：
  文献段 w:spacing line=270 (13.5pt) lineRule=exact
  段后 after=80 (4pt) → 条目间距 13.5+4 = 17.5pt，与 LaTeX 一致
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


def fix(docx=None, apply=False, line=270, after=80):
    docx = docx or HA.DOCX
    d = Document(docx)
    hits = []
    in_ref = False
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        if '参考文献' in t and len(t) <= 10:
            in_ref = True
            continue
        if not in_ref:
            continue
        if not re.match(r'^\[\d+\]', t):
            continue
        ppr = p._p.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr')
            p._p.insert(0, ppr)
        sp = ppr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing')
            ppr.append(sp)
        old = (sp.get(qn('w:line')), sp.get(qn('w:after')))
        sp.set(qn('w:line'), str(line))
        sp.set(qn('w:lineRule'), 'exact')
        sp.set(qn('w:after'), str(after))
        sp.set(qn('w:before'), '0')
        sp.set(qn('w:beforeAutospacing'), '0')
        sp.set(qn('w:afterAutospacing'), '0')
        HA.sort_el(ppr, HA.PPR_ORDER)
        hits.append((i, t[:20], old))
    print(f'文献段={len(hits)}  行距 {line/20}pt  段后 {after/20}pt')
    for i, t, old in hits[:5]:
        print(f'  P{i} 原={old} | {t}')
    if not apply:
        print('（未修改，加 --apply）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_refspace.docx')
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
    IO.save(d, docx)
    print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    ln = int(a[a.index('--line') + 1]) if '--line' in a else 270
    af = int(a[a.index('--after') + 1]) if '--after' in a else 80
    fix(apply='--apply' in a, line=ln, after=af)
