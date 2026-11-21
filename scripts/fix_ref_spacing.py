# -*- coding: utf-8 -*-
r"""参考文献区专项字距校正。

文献段特点
----------
· 悬挂缩进：首行左界 = 版心左（70.87），续行左界 = 70.87 + 缩进
· 条目编号 [n] 占首行宽度
· LaTeX 行提取时把 [n] 保留在行首，与正文段（上标引文被剔除）不同

因此通用 hard_iter 的「文本消耗」定位在文献区容易错位，
这里改为直接按段落顺序 + 条目编号锚定，逐条校正。

用法:
  python fix_ref_spacing.py <word_tag> [--apply]
"""
import sys, io, os, re, json, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_spacing as HS
import hard_iter as HI
import l2wconf as C
import docxio as IO


def ref_lines(tag):
    """从 Word PDF 提取文献区行，按条目编号分组。"""
    arr = json.load(open(os.path.join(HA.TMP, tag + '.json'), encoding='utf-8'))
    groups = {}
    cur = None
    started = False
    for P in arr:
        for L in P['lines']:
            n = HA.norm(L['t'])
            if '参考文献' in n and len(n) <= 10:
                started = True
                continue
            if not started:
                continue
            if L['y'] > 780:
                continue
            m = re.match(r'^\[(\d+)\]', n)
            if m and L['x0'] < 80:
                cur = int(m.group(1))
                groups[cur] = []
            if cur is not None:
                groups[cur].append(L)
    return groups


def tex_ref_lines():
    arr = json.load(open(C.tmp(C.TEX_TAG + '.json'), encoding='utf-8'))
    groups = {}
    cur = None
    started = False
    for P in arr:
        for L in P['lines']:
            n = HA.norm(L['t'])
            if n.startswith('参考文献'):
                started = True
                continue
            if not started:
                continue
            if L['y'] > 780:
                continue
            m = re.match(r'^\[(\d+)\]', n)
            if m and L['x0'] < 80:
                cur = int(m.group(1))
                groups[cur] = []
            if cur is not None:
                groups[cur].append(L)
    return groups


def run(tag, apply=False, damp=1.0):
    d = Document(HA.DOCX)
    TG = tex_ref_lines()
    WG = ref_lines(tag)
    # Word 段落按编号索引
    WP = {}
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        m = re.match(r'^\[(\d+)\]', t)
        if m:
            WP[int(m.group(1))] = (i, p)
    print(f'LaTeX条目={len(TG)} Word PDF条目={len(WG)} Word段落={len(WP)}')
    n_upd = 0
    for no in sorted(TG):
        if no not in WG or no not in WP:
            continue
        tl, wl = TG[no], WG[no]
        i, p = WP[no]
        segs = HS.line_segments(p._p)
        if len(segs) != len(tl):
            print(f'  [{no}] P{i} br段={len(segs)} T行={len(tl)}  <<段数不符')
            continue
        if len(wl) != len(tl):
            print(f'  [{no}] P{i} W行={len(wl)} T行={len(tl)}  <<溢出')
        for li, (seg, T) in enumerate(zip(segs, tl)):
            if li == len(tl) - 1 or not seg:
                continue
            nat, n = HS.seg_width(seg)
            if n == 0:
                continue
            tgt = T['x1'] - T['x0']
            # 该行的实际渲染宽：优先用 PDF 对应行；行数不符时用静态度量
            if li < len(wl):
                act = wl[li]['x1'] - wl[li]['x0']
                if len(wl) > len(tl):
                    act = None   # 溢出，改用静态度量
            else:
                act = None
            cur = HI.get_spacing(seg[0][0])
            if act is None:
                # 静态度量 + 当前 spacing 的实际效果
                act = nat + cur / 20.0 * n
            err = tgt - act
            if abs(err) < 0.15:
                continue
            add = err / n * 20 * damp
            add = max(-30, min(30, add))
            if abs(add) < 1.0:
                add = 1.0 if add > 0 else -1.0
            new = int(round(cur + add))
            new = max(-100, min(100, new))
            if new == cur:
                continue
            print(f'  [{no}] P{i} 行{li} 目标={tgt:.1f} 实际={act:.1f} '
                  f'err={err:+.2f} spacing {cur}→{new}')
            if apply:
                HS.apply_seg(seg, new / 20.0)
            n_upd += 1
    print(f'更新行数={n_upd}')
    if apply:
        bak = os.path.join(HA.TMP, 'bak_before_refsp.docx')
        if not os.path.exists(bak):
            shutil.copy2(HA.DOCX, bak)
        IO.save(d, HA.DOCX)
        print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    tag = a[0] if a and not a[0].startswith('--') else 'it4'
    run(tag, apply='--apply' in a)
