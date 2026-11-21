# -*- coding: utf-8 -*-
r"""顽固文献行兜底：用 PDF 实测宽度反解，逐步加大压缩直到行数收敛。

适用场景
--------
fix_ref3 的静态度量模型对含**连字（ﬁ/ﬂ）或变音符（Ü/Ñ/Ó）**的行会低估自然宽，
导致 spacing 已压到 0 仍溢出。此时改用 Word 导出 PDF 的实测数据：

    真实宽(当前sp) = ink_已渲染 + gaps_已渲染 + 剩余字符宽度 + 剩余字符数×sp/20

其中「剩余字符宽度」用静态度量近似（只有末词 1~2 个词，误差可控）。
反解目标 sp 后再额外留 SAFETY 余量，确保一次到位。

用法: python fix_ref4.py <word_tag> [--apply] [--safety 0.6] [--nos 18,33,35] [--extra 2]
"""
import sys, io, os, re, math, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import hard_spacing as HS
import hard_iter as HI
import fix_ref_spacing as FR
import fix_ref2 as F2
import docxio as IO

EXTRA = 2


def run(tag, apply=False, safety=0.6, nos=None):
    docx = HA.DOCX
    pdf = os.path.join(HA.TMP, tag + '.pdf')
    d = Document(docx)
    s = d.sections[0]
    UW = s.page_width.pt - s.left_margin.pt - s.right_margin.pt
    TG = FR.tex_ref_lines()
    WG = FR.ref_lines(tag)
    WP = {}
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        m = re.match(r'^\[(\d+)\]', t)
        if m:
            WP[int(m.group(1))] = (i, p)
    n_upd = 0
    for no in sorted(TG):
        if no not in WP or no not in WG:
            continue
        if nos and no not in nos:
            continue
        tl, wl = TG[no], WG[no]
        if len(wl) <= len(tl):
            continue
        i, p = WP[no]
        segs = HS.line_segments(p._p)
        print(f'=== [{no}] P{i} T={len(tl)} W={len(wl)} br段={len(segs)}')
        for li, seg in enumerate(segs):
            if not seg or (li == len(tl) - 1 and len(tl) > 1):
                continue
            txt = ''.join(t.text or '' for r, t in seg)
            nat, n = HS.seg_width(seg)
            cur = HI.get_spacing(seg[0][0])
            cs = F2.pdf_row_chars(pdf, txt[:10])
            if cs is None:
                print(f'   行{li} PDF未定位')
                continue
            pt = ''.join(c['c'] for c in cs).rstrip()
            # PDF 把 Ü/Ñ/Ó 等变音符渲染成独立 span（'MÜ LLER'），
            # 字符流与 docx 文本对不上。此时**不能用实测 ink**：
            # 前缀只能匹配到变音符之前，剩余字符占比过大会把 spacing 算爆。
            # 退回静态度量 + 强制额外压缩 EXTRA twip。
            if not txt.startswith(pt) or len(pt) < len(txt) * 0.7:
                nb = sum(1 for k in range(len(txt) - 1)
                         if txt[k] == ' ' and F2.is_cjk(txt[k + 1]))
                act = nat + 1.75 * nb + cur / 20.0 * n + 0.10
                need = (UW - safety - nat - 1.75 * nb - 0.10) / max(1, n) * 20.0
                new = max(-160, min(160, math.floor(need) - EXTRA))
                print(f'   行{li} sp={cur} n={n} nat={nat:.2f} 字符流不符'
                      f'(PDF {len(pt)}/{len(txt)}) act={act:.2f} '
                      f'超={act-UW:+.2f} → sp={new} (含额外-{EXTRA})')
                if new != cur:
                    if apply:
                        HS.apply_seg(seg, new / 20.0)
                    n_upd += 1
                continue
            ink = cs[-1]['bbox'][2] - cs[0]['bbox'][0]
            nrender = len(pt)
            rest = txt[nrender:]
            # 比例法：按字符数比例切分静态度量（同字号同字体，误差 < 1pt）
            nat_rest = nat * len(rest) / max(1, len(txt))
            act = ink + nat_rest + cur / 20.0 * n
            # 该行本身已不溢出（末词没被推走）→ 不要动它，
            # 放松 spacing 反而会把下一行的内容吸上来、破坏已对齐的断点。
            if act <= UW and not rest.strip():
                print(f'   行{li} sp={cur} act={act:.2f} 未溢出，跳过')
                continue
            need = (UW - safety - ink - nat_rest) / max(1, n) * 20.0
            new = max(-160, min(160, math.floor(need)))
            print(f'   行{li} sp={cur} n={n} nat={nat:.2f} ink={ink:.2f} '
                  f'渲染{nrender}/{len(txt)} rest={rest!r} nat_rest={nat_rest:.2f} '
                  f'act={act:.2f} 超={act-UW:+.2f} → sp={new}')
            if new == cur:
                continue
            if apply:
                HS.apply_seg(seg, new / 20.0)
            n_upd += 1
    print(f'更新行数={n_upd}')
    if apply and n_upd:
        bak = os.path.join(HA.TMP, 'bak_before_ref4.docx')
        if not os.path.exists(bak):
            shutil.copy2(docx, bak)
        IO.save(d, docx)
        print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    tag = a[0] if a and not a[0].startswith('--') else 'rf4'
    sf = 0.6
    if '--safety' in a:
        sf = float(a[a.index('--safety') + 1])
    ns = None
    if '--nos' in a:
        ns = set(int(x) for x in a[a.index('--nos') + 1].split(','))
    run(tag, apply='--apply' in a, safety=sf, nos=ns)
