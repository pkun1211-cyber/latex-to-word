# -*- coding: utf-8 -*-
r"""参考文献区字距精校（第二版，基于实测宽度模型）。

第一版 fix_ref_spacing.py 失败的原因
------------------------------------
它用「静态度量 nat + cur/20*n」预测行宽，但 Word 实际渲染还叠加了
两项静态度量抓不到的量：
  1. 常数偏置 ≈ +0.90pt（行首 ink 偏移 / 字距舍入累积）
  2. **西文空格 → CJK 边界的额外间距 ≈ +1.75pt/处**
     （Word 的中西文混排间距，autoSpaceDE=0 也压不掉）
中文文献条目只有 2~3 个空格但全是「空格→汉字」，故被低估 4~7pt，
导致第一版把本该压缩的行判成需要拉伸，越改越糟（42→49）。

第二版做法
----------
直接从 Word 导出 PDF 的 **字符级 bbox** 取真实自然宽（Σ 字符 advance），
再实测该行「空格→CJK」边界的额外间隙，两者相加得到可信基准：
    act = nat_pdf + Σ(边界间隙) + cur/20 × n
反解目标 spacing：
    cur' = (tgt − nat_pdf − Σ边界间隙) / n × 20
只改「行数溢出」或「宽度偏差 > 0.6pt」的行，其余不动。

用法:
  python fix_ref2.py <word_tag> [--apply] [--only-overflow] [--thresh 0.6]
"""
import sys, io, os, re, json, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
import pymupdf

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_spacing as HS
import hard_iter as HI
import fix_ref_spacing as FR
import docxio as IO


def pdf_row_chars(pdf_path, key, pages=None):
    """在 Word 导出 PDF 中找以 key 开头的行，返回字符级 bbox 列表。"""
    doc = pymupdf.open(pdf_path)
    rng = range(len(doc)) if pages is None else pages
    for pno in rng:
        dd = doc[pno].get_text('rawdict')
        for blk in dd['blocks']:
            if blk['type'] != 0:
                continue
            for ln in blk['lines']:
                cs = [c for s in ln['spans'] for c in s['chars']]
                if ''.join(c['c'] for c in cs).startswith(key):
                    doc.close()
                    return cs
    doc.close()
    return None


def is_cjk(ch):
    return ord(ch) > 0x2000


def boundary_extra(cs):
    """实测「西文空格 → CJK」边界的额外间隙之和。

    只统计 gap > 0.8pt 的边界（真混排间距）；
    个别条目该间距被 Word 压掉（如 [43]），实测为负则不计。
    """
    tot = 0.0
    cnt = 0
    for k in range(len(cs) - 1):
        if cs[k]['c'] == ' ' and is_cjk(cs[k + 1]['c']):
            g = cs[k + 1]['bbox'][0] - cs[k]['bbox'][2]
            if g > 0.8:
                tot += g
                cnt += 1
    return tot, cnt


def run(tag, apply=False, only_overflow=False, thresh=0.6):
    docx = HA.DOCX
    pdf = os.path.join(HA.TMP, tag + '.pdf')
    d = Document(docx)
    TG = FR.tex_ref_lines()
    WG = FR.ref_lines(tag)
    WP = {}
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        m = re.match(r'^\[(\d+)\]', t)
        if m:
            WP[int(m.group(1))] = (i, p)
    print(f'LaTeX条目={len(TG)} Word条目={len(WG)} Word段落={len(WP)}')
    n_upd = 0
    rows = []
    for no in sorted(TG):
        if no not in WG or no not in WP:
            continue
        tl, wl = TG[no], WG[no]
        i, p = WP[no]
        segs = HS.line_segments(p._p)
        if len(segs) != len(tl):
            rows.append((no, i, f'br段={len(segs)} T行={len(tl)} 段数不符'))
            continue
        over = len(wl) > len(tl)
        for li, (seg, T) in enumerate(zip(segs, tl)):
            if li == len(tl) - 1 or not seg:
                continue
            txt = ''.join(t.text or '' for r, t in seg)
            cs = pdf_row_chars(pdf, txt[:8])
            if cs is None:
                rows.append((no, i, f'行{li} PDF未定位'))
                continue
            # 自然宽必须用**静态度量**（覆盖整段字符）。
            # PDF 字符宽只含已渲染部分：溢出行末词被推到下一行，
            # 用 PDF 宽会把「少了 3 个词」误算成「需要大幅拉伸」。
            nat, n = HS.seg_width(seg)
            if n == 0:
                continue
            # PDF 侧只用来实测「空格→CJK」边界的额外间隙
            pt = ''.join(c['c'] for c in cs).rstrip()
            if pt == txt.rstrip():
                cs = cs[:len(txt)]
            elif txt.rstrip().startswith(pt):
                cs = cs[:len(pt)]
            else:
                cs = []
            extra, cnt = boundary_extra(cs)
            if not cs:
                # 文本含连字/变音符导致 PDF 字符流对不上时，
                # 按「西文空格后接汉字」的出现次数估算（实测 ≈1.75pt/处）
                cnt = sum(1 for k in range(len(txt) - 1)
                          if txt[k] == ' ' and is_cjk(txt[k + 1]))
                extra = 1.75 * cnt
            cur = HI.get_spacing(seg[0][0])
            tgt = T['x1'] - T['x0']
            need = (tgt - nat - extra) / max(1, n) * 20
            pred = nat + extra + cur / 20.0 * n
            err = tgt - pred
            if only_overflow and not over:
                continue
            if not over and abs(err) < thresh:
                continue
            new = int(round(need))
            new = max(-160, min(160, new))
            if new == cur:
                continue
            print(f'  [{no:2d}] P{i} 行{li} n={n} 边界={cnt} nat={nat:.2f} '
                  f'tgt={tgt:.2f} pred={pred:.2f} err={err:+.2f} '
                  f'spacing {cur}→{new}{"  溢出" if over else ""}')
            if apply:
                HS.apply_seg(seg, new / 20.0)
            n_upd += 1
    if rows:
        print('\n异常:')
        for no, i, why in rows[:30]:
            print(f'  [{no}] P{i} {why}')
    print(f'更新行数={n_upd}')
    if apply:
        bak = os.path.join(HA.TMP, 'bak_before_ref2.docx')
        if not os.path.exists(bak):
            shutil.copy2(docx, bak)
        IO.save(d, docx)
        print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    tag = a[0] if a and not a[0].startswith('--') else 'fn2'
    th = 0.6
    if '--thresh' in a:
        th = float(a[a.index('--thresh') + 1])
    run(tag, apply='--apply' in a, only_overflow='--only-overflow' in a, thresh=th)