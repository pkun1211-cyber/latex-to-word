# -*- coding: utf-8 -*-
r"""统一公式段间距，对齐 LaTeX amsmath equation 的垂直节奏。

LaTeX 实测（PDF 基线差）：
  式(1) y=680.5 → 式(2) y=714.9 → 式(3) y=749.4   间隔 34.4pt
  式(4) y=323.0 → 式(5) y=362.2                   间隔 39.2pt（含第二行）
  正文末行 → 首个公式：约 20pt
  末个公式 → 后续正文：约 20pt

Word 现状不一致：
  P39/40/41  line=270(13.5pt) before=200 after=200(10pt)  → 33.5pt
  P46/47     line=420(21pt)   before=140 after=140(7pt)   → 35pt
但 OMML 的行盒实际高度大于 line 设定值（分式/上下标撑高），
exact 行距会裁切，atLeast 会膨胀 → 统一改 auto 倍数行距不可控。

策略：保留 exact，但把 before/after 统一为 100(5pt)，
line 统一 270(13.5pt)，让公式段总节奏 = 13.5+10 = 23.5pt 起，
再由 OMML 实际高度决定。通过导出实测反馈调整。
"""
import sys, io, os, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO


def fix(docx=None, apply=False, line=270, before=100, after=100):
    docx = docx or HA.DOCX
    d = Document(docx)
    hits = []
    for i, p in enumerate(d.paragraphs):
        el = p._p
        if el.find('.//' + HA.M + 'oMath') is None:
            continue
        ppr = el.find(qn('w:pPr'))
        if ppr is None:
            continue
        jc = ppr.find(qn('w:jc'))
        # 只处理独立公式段（居中），不动行内公式所在正文段
        if jc is None or jc.get(qn('w:val')) != 'center':
            continue
        # 表题/图题也可能居中且含行内公式（如「治理×碳管理」），必须排除
        if p.style.name in ('Table Caption', 'Image Caption', 'Caption'):
            continue
        txt = HA.norm(''.join(x.text or '' for x in el.iter(HA.W + 't')))
        if txt.startswith(('表', '图', '注：')):
            continue
        sp = ppr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing')
            ppr.append(sp)
        old = (sp.get(qn('w:line')), sp.get(qn('w:before')), sp.get(qn('w:after')))
        sp.set(qn('w:line'), str(line))
        sp.set(qn('w:lineRule'), 'exact')
        sp.set(qn('w:before'), str(before))
        sp.set(qn('w:after'), str(after))
        HA.sort_el(ppr, HA.PPR_ORDER)
        hits.append((i, old))
    print(f'独立公式段={len(hits)}  line={line/20}pt before={before/20}pt after={after/20}pt')
    for i, old in hits:
        print(f'  P{i} 原={old}')
    if not apply:
        print('（未修改，加 --apply）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_eqspace.docx')
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
    IO.save(d, docx)
    print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    g = lambda k, dv: int(a[a.index(k) + 1]) if k in a else dv
    fix(apply='--apply' in a, line=g('--line', 270),
        before=g('--before', 100), after=g('--after', 100))
