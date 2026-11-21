# -*- coding: utf-8 -*-
r"""增量硬对齐：只对「尚未插 br 但已可配对」的段落补插 br + 字距。

用途
----
主流程 hard_apply 跑完后，若改进了配对判据（如文献首行前缀分母修正），
新增可配对段落需要补插 br，但不能重跑全量（会重复插入）。
本脚本按「段内是否已有 w:br」判断，只处理缺 br 的段。
"""
import sys, io, os, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_apply as AP
import hard_spacing as HS
import docxio as IO


def run(apply=False):
    d = Document(HA.DOCX)
    pairs, TP, WP = HA.plan(d)
    todo, skip = [], []
    for w, tp, sc, j in pairs:
        el = w['el']
        if el.find('.//' + qn('w:br')) is not None:
            continue          # 已硬对齐
        if len(tp) < 2:
            continue
        wn, tn = HA.mnorm(w['text']), HA.mnorm(''.join(L['t'] for L in tp))
        lenr = min(len(wn), len(tn)) / max(len(wn), len(tn))
        if lenr < 0.9:
            skip.append((w['i'], f'长度不符 {len(wn)}/{len(tn)}'))
            continue
        breaks, logs = AP.find_breaks(el, tp)
        if breaks is None:
            skip.append((w['i'], '; '.join(logs)))
            continue
        todo.append((w, tp, breaks))
    print(f'待补插段={len(todo)}  跳过={len(skip)}')
    for w, tp, b in todo:
        print(f"  P{w['i']:<3d} T行={len(tp)} br={len(b)} | {HA.mnorm(w['text'])[:36]}")
    for i, why in skip:
        print(f'  跳过 P{i}: {why}')
    if not apply:
        print('（未修改，加 --apply）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_incr.docx')
    if not os.path.exists(bak):
        shutil.copy2(HA.DOCX, bak)
    n = 0
    for w, tp, b in todo:
        n += AP.insert_breaks(w['el'], b)
        AP.set_left(w['el'])
        # 同步写首版字距补偿（末行除外）
        segs = HS.line_segments(w['el'])
        if len(segs) == len(tp):
            for li, (seg, L) in enumerate(zip(segs, tp)):
                if li == len(tp) - 1 or not seg:
                    continue
                nat, cnt = HS.seg_width(seg)
                if cnt == 0:
                    continue
                tgt = L['x1'] - L['x0']
                delta = (tgt - nat) / cnt
                if abs(delta) <= 0.6:
                    HS.apply_seg(seg, delta)
    IO.save(d, HA.DOCX)
    print(f'已补插 {n} 个 br，{len(todo)} 段')


if __name__ == '__main__':
    run(apply='--apply' in sys.argv)
