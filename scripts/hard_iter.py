# -*- coding: utf-8 -*-
r"""字距补偿的闭环迭代校正。

为什么必须迭代
--------------
静态度量（fontTools hmtx + 中文全角假设）与 Word 实际渲染存在系统偏差：
  · Word 对中文标点做挤压（kinsoku / 标点压缩），每个标点省 2~4pt
  · w:spacing 作用于 run 的每个字符间隙，含 run 末尾
  · 中西文混排时 Word 可能插入额外间距
实测：某行静态度量 441.00pt，Word 实际渲染 427.06pt，差 13.94pt。
故改用「导出 PDF → 实测行宽 → 反馈修正」的闭环，2~3 轮即可收敛。

算法
----
对每个已硬对齐段落的每一行（末行除外）：
  当前补偿 c（twips/字符）
  实测宽 w_act（从 Word 导出 PDF 提取）
  目标宽 w_tgt（LaTeX PDF）
  误差 e = w_tgt − w_act
  新补偿 c' = c + e / n × 20      （n = 该行非引文字符数）
阻尼系数 0.8 防振荡；单轮限幅 ±40 twips。

行匹配
------
Word 导出 PDF 的行序列与段落 br 切分一一对应的前提是「没有二次断行」。
若出现孤字行（br 前溢出），行数会多于 br 段数 → 该段本轮跳过并报告，
下一轮补偿增大后孤字行消失，自然恢复对齐。
"""
import sys, io, os, re, json, copy, shutil, subprocess
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_spacing as HS
import docxio as IO

W = HA.W
PY = r"C:\Users\20546\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
TMP = HA.TMP


def get_spacing(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        return 0
    e = rpr.find(qn('w:spacing'))
    if e is None:
        return 0
    return int(e.get(qn('w:val')))


def word_lines_by_page(tag):
    arr = json.load(open(os.path.join(TMP, tag + '.json'), encoding='utf-8'))
    out = []
    for P in arr:
        for L in P['lines']:
            out.append({'p': P['page'], 'y': L['y'], 'x0': L['x0'],
                        'x1': L['x1'], 'size': L['size'], 't': L['t']})
    return out


def clean_wlines(lines):
    """剔除页码、上标引文碎片、公式上下标碎片，口径与 tex 侧一致。"""
    CITE = re.compile(r'(?:\[\d+(?:[,，\-–]\s*\d+)*\]\s*)+$')
    out = []
    for L in lines:
        n = HA.norm(L['t'])
        if not n:
            continue
        if re.fullmatch(r'\d{1,2}', n) and L['y'] > 780:
            continue
        if L['size'] < 8.0:
            if CITE.fullmatch(n):
                continue
            if re.fullmatch(r"[′'∗*·°\d\s\-+=()]+", n):
                continue
            if len(n) <= 6:
                continue
        out.append(L)
    return out


def locate(wlines, tp):
    """在 Word 行序列中定位该段的连续行块（按首行前缀 + 行数）。

    返回 (start, ok)；ok=False 表示行数不符（存在二次断行）。
    """
    first = HA.mnorm(tp[0]['t'])
    if len(first) < 12:
        return None, False
    key = first[:12]
    for i, L in enumerate(wlines):
        if HA.mnorm(L['t']).startswith(key):
            return i, True
    return None, False


def iterate(tag, apply=False, damp=0.8, cap=40, max_delta=1.2):
    d = Document(HA.DOCX)
    pairs, TP, WP = HA.plan(d)
    wl = clean_wlines(word_lines_by_page(tag))
    stats = {'seg': 0, 'upd': 0, 'mismatch': 0, 'notfound': 0, 'last': 0}
    report = []
    for w, tp, sc, j in pairs:
        el = w['el']
        if el.find('.//' + qn('w:br')) is None:
            continue
        segs = HS.line_segments(el)
        if len(segs) != len(tp):
            stats['mismatch'] += 1
            report.append((w['i'], f'br段数={len(segs)} T行数={len(tp)}'))
            continue
        st, ok = locate(wl, tp)
        if not ok:
            stats['notfound'] += 1
            report.append((w['i'], 'Word行序列未找到首行'))
            continue
        # Word 侧行数必须足够
        if st + len(tp) > len(wl):
            stats['notfound'] += 1
            continue
        # 按「文本消耗」把 Word 行归并到 br 段：
        # 一个 br 段可能因二次断行对应多个 Word 行（孤字行），
        # 把它们合并后再量宽度，才能算出正确的补偿。
        wblk = []
        pos = st
        fail = False
        for li, T in enumerate(tp):
            need = HA.mnorm(T['t'])
            got = ''
            used = []
            while pos < len(wl) and len(got) < len(need):
                L = wl[pos]
                s = HA.mnorm(L['t'])
                if not s:
                    pos += 1
                    continue
                # 该 Word 行必须与剩余目标前缀一致
                rest = need[len(got):]
                k = min(len(s), len(rest), 10)
                if k and sum(1 for x, y in zip(s[:k], rest[:k]) if x == y) / k < 0.7:
                    fail = True
                    break
                got += s
                used.append(L)
                pos += 1
                # 单段最多吃 4 个 Word 行（超过说明定位错）
                if len(used) > 4:
                    fail = True
                    break
            if fail or not used:
                fail = True
                break
            # 合并宽度：各 Word 行宽度之和 = 该 br 段内容的真实渲染总宽
            # （二次断行时首行装满版心、溢出部分在后续行）
            tot = sum(u['x1'] - u['x0'] for u in used)
            wblk.append({'x0': used[0]['x0'], 'x1': used[0]['x0'] + tot,
                         'nline': len(used), 't': got})
        if fail:
            stats['mismatch'] += 1
            report.append((w['i'], f'文本消耗失败（行内容不符）'))
            continue
        for li, (seg, T, L) in enumerate(zip(segs, tp, wblk)):
            stats['seg'] += 1
            if li == len(tp) - 1:
                stats['last'] += 1
                continue
            if not seg:
                continue
            nat, n = HS.seg_width(seg)
            if n == 0:
                continue
            tgt = T['x1'] - T['x0']
            act = L['x1'] - L['x0']   # 已按多行累加
            err = tgt - act
            overflow = L['nline'] > 1
            # 保守策略：
            #   · 溢出行（br 前内容装不下，产生孤字行）→ 必须修，按误差全额补偿
            #   · 未溢出但宽度偏差 > 1.2pt → 修（右界明显不齐）
            #   · 未溢出且偏差小 → 不动（避免累积过度导致页数膨胀）
            # 只修溢出行。未溢出的行即使右界略不齐也不动 ——
            # 实测改这些行会引发后续断行连锁变化，反馈环路震荡，
            # 两轮后行数偏差反而从 44 涨到 69。
            if not overflow:
                continue
            cur = get_spacing(seg[0][0])
            add = err / n * 20 * (1.0 if overflow else damp)
            add = max(-cap, min(cap, add))
            # 溢出行哪怕只需 0.2pt 也必须至少减 1 twip，否则末字仍被推走
            if overflow and abs(add) < 1.0:
                add = 1.0 if add > 0 else -1.0
            new = int(round(cur + add))
            new = max(-160, min(160, new))
            if new == cur:
                continue
            if abs(new) / 20.0 > max_delta:
                report.append((w['i'], f'行{li} 补偿超限 {new/20:+.2f}pt err={err:+.2f}'))
                continue
            if apply:
                HS.apply_seg(seg, new / 20.0)
            stats['upd'] += 1
    print(f"行片段={stats['seg']} 更新={stats['upd']} 末行={stats['last']} "
          f"行数不符={stats['mismatch']} 未定位={stats['notfound']}")
    if report:
        print('\n异常:')
        for i, why in report[:40]:
            print(f'  P{i}  {why}')
    if apply:
        IO.save(d, HA.DOCX)
        print('已写入')
    return stats


if __name__ == '__main__':
    a = sys.argv[1:]
    tag = a[0] if a and not a[0].startswith('--') else 'ha1'
    iterate(tag, apply='--apply' in a)
