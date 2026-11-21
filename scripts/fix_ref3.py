# -*- coding: utf-8 -*-
r"""参考文献区字距精校（第三版：实测校准模型 + 逐行 floor 求解）。

模型（本论文实测标定）
----------------------
    真实行宽 = nat + Σ(空格→CJK 边界额外间距) + cur × n / 20 + BIAS

  nat  = HS.seg_width(seg) 静态度量（LM Roman 字形 advance 之和，含引文上标）
  n    = 可施加字距的字符数（不含引文上标）
  cur  = 当前 w:spacing（twip）
  BIAS = +0.10pt 常数（行首 ink 偏移），由非溢出行反解标定
  边界间距 ≈ 1.75pt/处，Word 中西文混排间距，autoSpaceDE=0 压不掉

关键校正（相对第一、二版）
--------------------------
* spacing 生效系数 = **1.0**，不是 0.968。
  证据：字符间隙中位数恰等于 cur/20（cur=5→0.252，cur=1→0.045）。
* BIAS = 0.10pt，不是 0.90pt。第二版把边界间距重复计入才拟合出 0.90。
* 自然宽必须用静态度量：溢出行的 PDF 只渲染了前半部分（末词被推走），
  用 PDF advance 会把「少了 3 个词」误算成「需要大幅拉伸」。
* 目标宽用 **Word 实际可用宽**（版心宽 − 首行左缩进），不用 TeX 行宽。
  TeX 行宽含行末空白回退误差，会低估约 1pt。
* 求解用 **floor** 而非 round：溢出行哪怕只超 0.1pt 也会断行，
  round 会留下残余溢出（实测 7 行全部未修好）。

用法:
  python fix_ref3.py [--apply] [--margin 0.15] [--all]
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
import docxio as IO

BIAS = 0.10
BOUND = 1.75          # 空格→CJK 边界间距，pt/处
RIGHT_EPS = 0.10      # 行首 ink 偏移已在 BIAS 中体现


def is_cjk(ch):
    return ord(ch) > 0x2000


def cjk_bounds(txt):
    """统计「西文空格 → CJK」边界数。"""
    return sum(1 for k in range(len(txt) - 1)
               if txt[k] == ' ' and is_cjk(txt[k + 1]))


def usable_width(doc):
    s = doc.sections[0]
    return s.page_width.pt - s.left_margin.pt - s.right_margin.pt


def run(apply=False, margin=0.15, do_all=False):
    docx = HA.DOCX
    d = Document(docx)
    UW = usable_width(d)
    TG = FR.tex_ref_lines()
    WP = {}
    for i, p in enumerate(d.paragraphs):
        t = HA.norm(''.join(x.text or '' for x in p._p.iter(HA.W + 't')))
        m = re.match(r'^\[(\d+)\]', t)
        if m:
            WP[int(m.group(1))] = (i, p)
    WG = FR.ref_lines(os.environ.get('REFTAG','rf3'))
    print(f'Word 可用宽={UW:.2f}pt  条目 T={len(TG)} W={len(WP)}')
    n_upd = 0
    for no in sorted(TG):
        if no not in WP:
            continue
        tl = TG[no]
        i, p = WP[no]
        segs = HS.line_segments(p._p)
        if len(segs) != len(tl):
            print(f'  [{no}] P{i} br段={len(segs)} T行={len(tl)} 段数不符，跳过')
            continue
        over = len(WG.get(no, [])) > len(tl)
        if not over and not do_all:
            continue
        for li, seg in enumerate(segs):
            if not seg:
                continue
            txt = ''.join(t.text or '' for r, t in seg)
            nat, n = HS.seg_width(seg)
            if n == 0:
                continue
            cur = HI.get_spacing(seg[0][0])
            nb = cjk_bounds(txt)
            extra = BOUND * nb
            pred = nat + extra + cur / 20.0 * n + BIAS
            last = (li == len(tl) - 1)
            # 末行是自然宽（TeX 不拉伸），绝不能拉满；
            # 唯一例外是单行段（li==0 且 len(tl)==1），它既是首行也是末行，
            # 溢出时必须压缩（如 [12]）。
            if last and len(tl) > 1:
                continue
            avail = UW - margin
            need = (avail - nat - extra - BIAS) / n * 20.0
            new = max(-160, min(160, math.floor(need)))
            # 注意：不能用「pred 与可用宽差 < 0.8pt 就跳过」做过滤。
            # 溢出行往往只超 0.1~0.5pt，正是这类行必须压掉 1 twip。
            if new == cur:
                continue
            print(f'  [{no:2d}] P{i} 行{li} n={n} 边界={nb} nat={nat:.2f} '
                  f'pred={pred:.2f} 超={pred-UW:+.2f} spacing {cur}→{new}'
                  f'{"  溢出" if over else ""}{"  末行" if last else ""}')
            if apply:
                HS.apply_seg(seg, new / 20.0)
            n_upd += 1
    print(f'更新行数={n_upd}')
    if apply and n_upd:
        bak = os.path.join(HA.TMP, 'bak_before_ref3.docx')
        if not os.path.exists(bak):
            shutil.copy2(docx, bak)
        IO.save(d, docx)
        print('已写入')


if __name__ == '__main__':
    a = sys.argv[1:]
    mg = 0.15
    if '--margin' in a:
        mg = float(a[a.index('--margin') + 1])
    run(apply='--apply' in a, margin=mg, do_all='--all' in a)
