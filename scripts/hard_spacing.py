# -*- coding: utf-8 -*-
r"""硬对齐第二步：按行写字符间距补偿，使每行宽度精确等于 LaTeX 实测宽度。

为什么需要
----------
插 w:br 只锁定了「断点」，但 Word 仍会在 br 之前自动断行：
若该行自然宽度 > 可用宽度，末词被挤到下一行（实测 Abstract 出现
"ﬁrms." / "increase" / "full" / "scores" 单独成行）。
反之中文行自然宽度普遍比 LaTeX 窄 3~25pt（TeX 的 \CJKglue 在汉字间
插了可伸缩胶水），右界参差。

做法
----
对每个已硬对齐的段落，按 br 切出「行片段」，逐行：
  自然宽 = Σ 字符 advance（中文全角 = 字号，西文查 LM Roman hmtx）
  目标宽 = LaTeX 该行 x1 − x0
  每字符补偿 = (目标 − 自然) / 字符数
写 <w:spacing w:val="N"/>，N = round(补偿 × 20)（单位 1/20 pt）。

末行不补偿（LaTeX 末行是自然宽，不需要撑满）。
补偿超过 ±0.6pt/字符 的行跳过（说明配对或度量有误，硬拉会变形）。

实现细节
--------
w:spacing 是 run 级属性，故需把行片段内的 run 按补偿值再拆一次。
为减少 run 爆炸，同一行内相同 rPr 的连续 run 合并处理。
"""
import sys, io, os, re, json, copy, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_apply as AP
import docxio as IO

W = HA.W
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def lm_family(pt):
    if pt <= 5.5:
        return 'LM Roman 5'
    if pt <= 6.5:
        return 'LM Roman 6'
    if pt <= 7.5:
        return 'LM Roman 7'
    if pt <= 8.5:
        return 'LM Roman 8'
    if pt <= 9.5:
        return 'LM Roman 9'
    if pt <= 11.0:
        return 'LM Roman 10'
    if pt <= 14.0:
        return 'LM Roman 12'
    return 'LM Roman 17'


def run_size(r, default=9.0):
    rpr = r.find(qn('w:rPr'))
    if rpr is not None:
        e = rpr.find(qn('w:sz'))
        if e is not None:
            return int(e.get(qn('w:val'))) / 2.0
    return default


def walk_runs(p_el):
    """按文档顺序产出 (run, [w:t...])，遇到 w:br 产出哨兵 ('BR', None)。"""
    for r in p_el.iter():
        if r.tag != W + 'r':
            continue
        # run 内可能 br 与 t 混排，需按子节点顺序
        for kid in r:
            if kid.tag == W + 'br':
                yield ('BR', None)
            elif kid.tag == W + 't':
                yield ('T', (r, kid))


def line_segments(p_el):
    """按 w:br 切分，返回 [[(run, tnode), ...], ...]。"""
    segs = [[]]
    for kind, payload in walk_runs(p_el):
        if kind == 'BR':
            segs.append([])
        else:
            segs[-1].append(payload)
    return segs


def is_cite_run(r):
    """引文 run：字号 <= 6pt 且内容是 [数字,] 片段或纯数字（含 hyperlink 内）。

    LaTeX 行提取时上标引文被剔除，故其宽度不计入目标行宽；
    若度量时把引文算进自然宽，会算出错误的负补偿（实测 −2 twips）。
    """
    sz = run_size(r, default=99)
    if sz > 6.5:
        return False
    txt = ''.join(t.text or '' for t in r.findall(qn('w:t')))
    if re.fullmatch(r'[\[\],\d\s]+', txt or ' '):
        return True
    par = r.getparent()
    if par is not None and par.tag == W + 'hyperlink' and re.fullmatch(r'\d+', txt or ''):
        return True
    return False


def seg_width(seg, exclude_cite=True):
    """返回 (自然宽, 可补偿字符数)。

    引文字符仍占实际排版宽度，必须计入自然宽；
    但补偿只能施加到非引文 run 上（引文是上标，拉宽会破坏视觉），
    故字符数只统计非引文部分，补偿值 = 差值 / 非引文字符数。
    """
    w = 0.0
    n = 0
    for r, t in seg:
        sz = run_size(r)
        fam = lm_family(sz)
        cite = is_cite_run(r) if exclude_cite else False
        for ch in (t.text or ''):
            w += HA.MET.adv(ch, fam, sz)
            if not cite:
                n += 1
    return w, n


def set_spacing(r, twips):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        rpr = OxmlElement('w:rPr')
        r.insert(0, rpr)
    e = rpr.find(qn('w:spacing'))
    if e is None:
        e = OxmlElement('w:spacing')
        rpr.append(e)
    e.set(qn('w:val'), str(int(twips)))
    HA.sort_el(rpr, HA.RPR_ORDER)


def apply_seg(seg, delta_per_char):
    """把补偿写到该行每个 run 上。

    注意 w:spacing 会在 run 的**每个字符后**加间距（含最后一个），
    故 n 个字符实际增加 n × delta。补偿公式已按字符数均摊，一致。
    """
    tw = int(round(delta_per_char * 20))
    # 注意：tw == 0 时**不能提前 return**。
    # 若该 run 已有非零 w:spacing，必须显式写 0 覆盖，否则旧值残留，
    # 表现为「脚本报告已改、导出后行宽却没变」（实测 [18][35] 因此修不掉）。
    touched = set()
    for r, t in seg:
        if id(r) in touched:
            continue
        if is_cite_run(r):
            continue
        touched.add(id(r))
        set_spacing(r, tw)
    return len(touched)


def run(docx=None, apply=False, max_delta=0.6, report=None):
    docx = docx or HA.DOCX
    d = Document(docx)
    pairs, TP, WP = HA.plan(d)
    stats = {'seg': 0, 'ok': 0, 'skip_big': 0, 'skip_last': 0, 'runs': 0}
    rows = []
    for w, tp, sc, j in pairs:
        el = w['el']
        # 只处理已插 br 的段落
        if el.find('.//' + qn('w:br')) is None:
            continue
        segs = line_segments(el)
        if len(segs) != len(tp):
            rows.append((w['i'], f'段行数不符 W={len(segs)} T={len(tp)}'))
            continue
        for li, (seg, L) in enumerate(zip(segs, tp)):
            stats['seg'] += 1
            if not seg:
                continue
            nat, n = seg_width(seg)
            tgt = L['x1'] - L['x0']
            if li == len(tp) - 1:
                stats['skip_last'] += 1
                continue
            delta = (tgt - nat) / max(1, n)
            if abs(delta) > max_delta:
                stats['skip_big'] += 1
                rows.append((w['i'], f'行{li} 补偿过大 {delta:+.3f}pt/字符 '
                                     f'自然={nat:.1f} 目标={tgt:.1f}'))
                continue
            if apply:
                stats['runs'] += apply_seg(seg, delta)
            stats['ok'] += 1
    print(f"行片段={stats['seg']}  可补偿={stats['ok']}  "
          f"末行跳过={stats['skip_last']}  补偿过大跳过={stats['skip_big']}")
    if rows:
        print('\n异常明细:')
        for i, why in rows[:60]:
            print(f'  P{i}  {why}')
    if not apply:
        print('\n（未修改文件，加 --apply 执行）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_spacing.docx')
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
        print(f'已备份 → {bak}')
    IO.save(d, docx)
    print(f"已写入字距补偿：{stats['ok']} 行，{stats['runs']} 个 run")
    if report:
        json.dump(stats, open(report, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    a = sys.argv[1:]
    md = 0.6
    if '--max-delta' in a:
        md = float(a[a.index('--max-delta') + 1])
    run(apply='--apply' in a, max_delta=md)
