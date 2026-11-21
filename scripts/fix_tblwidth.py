# -*- coding: utf-8 -*-
r"""把表格宽度对齐 LaTeX 版心宽，消除单元格内多余换行。

实测问题
--------
表6 列宽合计 7920 twips = 396pt，而版心宽 453.5pt = 9070 twips。
Word 因此把「RKS 治理评分」折成「RKS 治理评」+「分」两行，
表头折成 3 行，整表比 LaTeX 高约 36pt，累积后把后续内容挤到下一页。

LaTeX 侧表格用 \resizebox{\textwidth}{!}{...} 或 tabular* 撑满版心，
故 Word 也应按版心宽等比放大列宽。

用法:
  python fix_tblwidth.py --list
  python fix_tblwidth.py --apply            # 全部表格按版心宽等比放大
  python fix_tblwidth.py --apply --only 5   # 只处理第 5 个表
"""
import sys, io, os, re, shutil, json
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO

# 版心：A4 宽 595.3pt − 左右各 2.5cm(70.87pt) = 453.5pt
TEXT_W_PT = 453.53
TEXT_W = int(round(TEXT_W_PT * 20))   # 9071 twips


def table_width(t):
    grid = t._tbl.find(qn('w:tblGrid'))
    if grid is None:
        return None, None
    cols = [int(c.get(qn('w:w'))) for c in grid]
    return sum(cols), cols


def set_width(t, target):
    grid = t._tbl.find(qn('w:tblGrid'))
    cols = [int(c.get(qn('w:w'))) for c in grid]
    tot = sum(cols)
    if tot <= 0:
        return False
    k = target / tot
    new = [max(1, int(round(c * k))) for c in cols]
    # 修正累计舍入误差
    new[-1] += target - sum(new)
    for c, v in zip(grid, new):
        c.set(qn('w:w'), str(v))
    # tblW
    pr = t._tbl.find(qn('w:tblPr'))
    if pr is not None:
        w = pr.find(qn('w:tblW'))
        if w is not None:
            w.set(qn('w:w'), str(target))
            w.set(qn('w:type'), 'dxa')
    # 每个单元格的 tcW 也要同步，否则 Word 仍按旧宽排版
    for row in t._tbl.findall(qn('w:tr')):
        tcs = row.findall(qn('w:tc'))
        ci = 0
        for tc in tcs:
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                ci += 1
                continue
            gs = tcpr.find(qn('w:gridSpan'))
            span = int(gs.get(qn('w:val'))) if gs is not None else 1
            w = tcpr.find(qn('w:tcW'))
            if w is not None:
                w.set(qn('w:w'), str(sum(new[ci:ci + span])))
                w.set(qn('w:type'), 'dxa')
            ci += span
    return True


def main():
    a = sys.argv[1:]
    d = Document(HA.DOCX)
    only = int(a[a.index('--only') + 1]) if '--only' in a else None
    rows = []
    for ti, t in enumerate(d.tables):
        tot, cols = table_width(t)
        if tot is None:
            continue
        head = HA.norm(''.join(x.text or '' for x in t._tbl.iter(HA.W + 't')))[:30]
        rows.append((ti, tot, tot / 20, head))
    if '--list' in a or '--apply' not in a:
        print(f'版心宽 = {TEXT_W} twips = {TEXT_W_PT}pt')
        for ti, tot, pt, head in rows:
            flag = '  <<偏窄' if tot < TEXT_W - 200 else ''
            print(f'  表[{ti}] 宽={tot}({pt:.1f}pt){flag} | {head}')
        if '--apply' not in a:
            return
    bak = os.path.join(HA.TMP, 'bak_before_tblwidth.docx')
    if not os.path.exists(bak):
        shutil.copy2(HA.DOCX, bak)
        print('备份 ->', bak)
    n = 0
    for ti, t in enumerate(d.tables):
        if only is not None and ti != only:
            continue
        tot, cols = table_width(t)
        if tot is None or abs(tot - TEXT_W) < 60:
            continue
        if set_width(t, TEXT_W):
            n += 1
            print(f'  表[{ti}] {tot} → {TEXT_W}')
    IO.save(d, HA.DOCX)
    print(f'已调整 {n} 个表')


if __name__ == '__main__':
    main()
