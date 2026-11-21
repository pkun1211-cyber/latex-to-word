# -*- coding: utf-8 -*-
r"""逐行硬对齐：按 LaTeX PDF 的实际断行结果锁定 Word 的每一行。

原理
----
TeX 的 \spaceskip 允许词间空格双向伸缩（实测 −0.74 ~ +1.81pt/空格），
Word 两端对齐只拉伸不压缩，故同一段落断行必然不同。
唯一彻底的办法：把 LaTeX 的断行结果直接写进 Word。

做法
----
1. 从 tex PDF 提取每段的行序列（已按视觉行聚合）
2. 与 Word 段落按文本内容配对（归一化后比对，容忍公式/连字差异）
3. 对每个 Word 段落：
   a. 在 LaTeX 的断点处插入 <w:br/>（软回车，段内换行，不拆段）
   b. 关闭该段两端对齐改 left（避免 Word 再次拉伸）
   c. 按行计算空格字距补偿：
      目标行宽 = LaTeX 实测行宽
      自然行宽 = 字符 advance 之和
      每空格补偿 = (目标 − 自然) / 空格数
      用 <w:spacing w:val="N"/>（单位 1/20 pt）写到该行的空格 run 上
4. 保留公式节点原位（绝不整段重建，避免 m:oMath 错位）

关键约束
--------
- 只在**纯文本节点**内插 br，公式节点前后不动
- 含公式的段落：跳过硬对齐（公式位置无法用 br 精确控制），保留原样
- 表格单元格、图段、公式段一律跳过
- 首页题名/作者/单位/摘要标签段特殊处理
- 全程备份 + 逐段日志，可完整回滚

用法:
  python hard_align.py --plan          # 只出计划，不改文件
  python hard_align.py --apply         # 执行
  python hard_align.py --undo          # 回滚
  python hard_align.py --apply --only-body   # 只处理正文段（保守）
"""
import sys, io, os, re, json, shutil, collections
# 仅作为主程序运行时才包装 stdout，避免被 import 时产生副作用
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from fontTools.ttLib import TTFont
import l2wconf as C

DIR = C.DIR
# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
TMP = C.TMP
BAK = os.path.join(TMP, 'bak_before_hard.docx')
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
             'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
             'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
             'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap',
             'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps',
             'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint',
             'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
             'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang',
             'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange']


def sort_el(el, order):
    kids = list(el)
    for e in kids:
        el.remove(e)
    for e in sorted(kids, key=lambda e: order.index(e.tag.split('}')[1])
                    if e.tag.split('}')[1] in order else 900):
        el.append(e)


def tx(el):
    return ''.join(t.text or '' for t in el.iter() if t.tag in (W + 't', M + 't'))


def norm(s):
    """归一化：去空白、连字还原、数学斜体转普通、去零宽字符。"""
    s = ''.join(s.split())
    for a, b in (('\ufb00', 'ff'), ('\ufb01', 'fi'), ('\ufb02', 'fl'),
                 ('\ufb03', 'ffi'), ('\ufb04', 'ffl')):
        s = s.replace(a, b)
    s = s.replace('\u200b', '').replace('\u2060', '')
    # 数学字母数字符号 → ASCII
    out = []
    for ch in s:
        o = ord(ch)
        if 0x1d400 <= o <= 0x1d7ff:
            base = (o - 0x1d400) % 52
            out.append(chr(ord('A') + base) if base < 26 else chr(ord('a') + base - 26))
        elif 0x1d7ce <= o <= 0x1d7ff:
            out.append(chr(ord('0') + (o - 0x1d7ce) % 10))
        else:
            out.append(ch)
    return ''.join(out)


# ---------------------------------------------------------------- 字体度量
class Metrics:
    def __init__(self):
        self.cache = {}
        self.ttfdir = C.TTFDIR
        self.song = r'C:\Windows\Fonts\simsun.ttc'

    def load(self, fam):
        if fam in self.cache:
            return self.cache[fam]
        fn = {'LM Roman 5': 'lmroman5-regular.ttf', 'LM Roman 6': 'lmroman6-regular.ttf',
              'LM Roman 7': 'lmroman7-regular.ttf', 'LM Roman 8': 'lmroman8-regular.ttf',
              'LM Roman 9': 'lmroman9-regular.ttf', 'LM Roman 10': 'lmroman10-regular.ttf',
              'LM Roman 12': 'lmroman12-regular.ttf', 'LM Roman 17': 'lmroman17-regular.ttf',
              }.get(fam)
        f = None
        if fn:
            p = os.path.join(self.ttfdir, fn)
            if os.path.exists(p):
                f = TTFont(p, lazy=True)
        if f is None:
            f = TTFont(os.path.join(self.ttfdir, 'lmroman9-regular.ttf'), lazy=True)
        self.cache[fam] = (f, f.getBestCmap(), f['hmtx'], f['head'].unitsPerEm)
        return self.cache[fam]

    def adv(self, ch, fam='LM Roman 9', size=9.0):
        o = ord(ch)
        # CJK 全角
        if ('\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f'
                or '\uff00' <= ch <= '\uffef'):
            return size
        f, cmap, hmtx, upm = self.load(fam)
        g = cmap.get(o)
        if g is None:
            return size * 0.5
        return hmtx[g][0] / upm * size


MET = Metrics()


# ---------------------------------------------------------------- LaTeX 行序列
def tex_lines():
    arr = json.load(open(C.tmp(C.TEX_TAG + '.json'), encoding='utf-8'))
    out = []
    for P in arr:
        for L in P['lines']:
            out.append({'p': P['page'], 'y': L['y'], 'x0': L['x0'], 'x1': L['x1'],
                        'size': L['size'], 't': L['t']})
    return out


def group_tex_paragraphs(lines):
    """把 LaTeX 行按段落聚合。

    两个必须先剔除的干扰：
    1. 上标引文 [1] 自成一行（基线比正文高约 9.7pt，字号 5.98），
       若不剔除会把一个正文段截成多段。
    2. 页码（y>780 且纯数字）。
    段落起点判据：正文首行缩进 2em → x0≈88.9（版心左界 70.9）。
    """
    # 先过滤上标碎片与页码。
    # 注意：上标引文常常是多个并排，如 "[17][18, 19]"、"[33] [34] [35, 36]"，
    # 用 fullmatch(r'\[[\d,\s]+\]') 会漏掉 → 必须用「整行只由引文括号构成」的判据。
    CITE = re.compile(r'(?:\[\d+(?:[,，\-–]\s*\d+)*\]\s*)+$')
    clean = []
    for L in lines:
        t = L['t'].strip()
        if not t:
            continue
        n = norm(t)
        if re.fullmatch(r'\d{1,2}', n) and L['y'] > 780:
            continue
        if L['size'] < 8.0:
            # 上标引文（可多个并排）
            if CITE.fullmatch(n):
                continue
            # 公式上下标 / 转置符 / 显著性星号 / 单个数字
            if re.fullmatch(r"[′'∗*·°\d\s\-+=()]+", n):
                continue
            if len(n) <= 6:
                continue
        clean.append(L)

    paras = []
    cur = []
    RIGHT = 524.4      # 版心右界
    LEFT = 70.87       # 版心左界
    in_ref = False     # 是否已进入参考文献区
    ref_no = 0         # 已见的文献序号（必须递增才算新条目）

    def head_end(s):
        return (bool(re.match(r'^[一二三四五六七八九十]、', s))
                or bool(re.match(r'^（[一二三四五六]）', s))
                or s.startswith('注：'))

    for L in clean:
        n = norm(L['t'])
        if n.startswith('参考文献'):
            in_ref = True
            if cur:
                paras.append(cur)
                cur = []
            cur.append(L)
            paras.append(cur)
            cur = []
            continue

        if in_ref:
            # 参考文献：新条目 = 行首在版心左界（悬挂缩进）+ [n] 且 n 递增。
            # 续行常以 [J]. / [M]. / [2023] 开头，只看 ^\[\d+\] 会全部误判。
            m = re.match(r'^\[(\d+)\]', n)
            at_left = L['x0'] < LEFT + 6
            newref = bool(m) and at_left and int(m.group(1)) == ref_no + 1
            if newref:
                if cur:
                    paras.append(cur)
                    cur = []
                ref_no = int(m.group(1))
            cur.append(L)
            continue

        # 新段起点判据（必须同时满足）：
        #   a. 首行缩进 2em → x0≈88.9
        #   b. 前一行是段末（右界明显缩进，未撑满版心）或前一行是标题
        # 仅靠 x0 会把「表格/公式附近恰好缩进的续行」误判为新段。
        indent = 87.5 < L['x0'] < 90.5
        if indent and cur:
            prev = cur[-1]
            pn = norm(prev['t'])
            prev_end = prev['x1'] < RIGHT - 12
            prev_head = head_end(pn)
            # 跨页且本页首行：只有前段已收尾才认新段（LaTeX 段落常跨页）
            cross_page = L['p'] != prev['p'] and prev_end
            if not (prev_end or prev_head or cross_page):
                indent = False
        is_head = (bool(re.match(r'^[一二三四五六七八九十]、', n))
                   or bool(re.match(r'^（[一二三四五六]）', n)))
        is_special = (n.startswith('注：')
                      or (re.match(r'^表\d+', n) and len(n) <= 40
                          and not re.search(r'[。，；：、]', n))
                      or re.match(r'^图\d+', n)
                      or n.startswith('［') or n.startswith('Abstract')
                      or n.startswith('Keywords'))
        if cur and (indent or is_head or is_special):
            paras.append(cur)
            cur = []
        cur.append(L)
        # 标题自成一段：flush 后 cur 已空，再 append 会重复存储
        if is_head:
            paras.append(cur)
            cur = []
            continue
    if cur:
        paras.append(cur)
    return paras


# ---------------------------------------------------------------- Word 段落
def word_paras(d):
    out = []
    for i, p in enumerate(d.paragraphs):
        el = p._p
        has_math = el.find(f'.//{M}oMath') is not None
        has_img = bool(el.findall('.//' + qn('w:drawing')))
        out.append({'i': i, 'p': p, 'el': el, 'style': p.style.name,
                    'text': tx(el), 'math': has_math, 'img': has_img})
    return out


CITE_INLINE = re.compile(r'\[\d+(?:[,，\-–]\d+)*\]')


def mnorm(s):
    """匹配用归一化：在 norm 基础上再去掉引文标记。

    TeX 行提取时上标引文被当作独立碎片剔除，而 Word 段内引文仍在文本里，
    不去掉会导致前缀匹配在第一个引文处中断。
    """
    s = norm(s)
    s = CITE_INLINE.sub('', s)
    return s


def plan(d, only_body=False):
    """单调序列对齐：Word 段序与 LaTeX 段序同向，只在上一命中之后向后找。

    贪心「全局最佳 + used 集合」会让靠前的 Word 段抢走靠后的 LaTeX 段，
    造成后续指针错位（实测 P22/P24 只配到 1 行）。
    """
    TL = tex_lines()
    TP = group_tex_paragraphs(TL)
    WP = word_paras(d)
    tn_list = [mnorm(''.join(L['t'] for L in tp)) for tp in TP]
    pairs = []
    j0 = 0
    for w in WP:
        if w['img'] or w['math']:
            continue
        if w['style'] in ('Table Caption', 'Image Caption', 'Heading 1', 'Heading 2'):
            continue
        if only_body and w['style'] not in ('Body Text', 'First Paragraph'):
            continue
        wn = mnorm(w['text'])
        if len(wn) < 20:
            continue
        best, bestscore = None, 0.0
        # 向后窗口：允许跳过表格/公式/图对应的 LaTeX 段
        for j in range(j0, min(j0 + 40, len(TP))):
            tn = tn_list[j]
            if len(tn) < 20:
                continue
            k = min(len(wn), len(tn))
            same = 0
            for a, b in zip(wn[:k], tn[:k]):
                if a != b:
                    break
                same += 1
            # 前缀一致性：以「LaTeX 首行长度」为分母，而非整段长度。
            # 文献条目常首行完全一致、次行才分歧，用整段长度做分母会把
            # pre 压到 0.53~0.86，导致 [9][17][19][24] 等条目漏配对。
            first_len = len(mnorm(TP[j][0]['t'])) or k
            denom = min(k, max(20, first_len))
            pre = same / denom
            lenr = min(len(wn), len(tn)) / max(len(wn), len(tn))
            score = pre * 0.75 + lenr * 0.25
            if pre > 0.9 and score > bestscore:
                bestscore, best = score, j
        if best is not None:
            pairs.append((w, TP[best], bestscore, best))
            j0 = best + 1
    return pairs, TP, WP


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    args = sys.argv[1:]
    if '--undo' in args:
        shutil.copy2(BAK, DOCX)
        print('已回滚')
        return
    d = Document(DOCX)
    pairs, TP, WP = plan(d, only_body='--only-body' in args)
    print(f'LaTeX 段落={len(TP)}  Word 段落={len(WP)}  可配对={len(pairs)}')
    tot_lines = 0
    bad = []
    for w, tp, sc, j in pairs:
        wn, tn = mnorm(w['text']), mnorm(''.join(L['t'] for L in tp))
        lenr = min(len(wn), len(tn)) / max(len(wn), len(tn))
        flag = '' if lenr > 0.9 else '  <<长度不符'
        if flag:
            bad.append((w['i'], j, len(wn), len(tn)))
        if len(pairs) <= 400:
            print(f"  P{w['i']:<3d}→T{j:<3d} [{w['style'][:14]:<14s}] 匹配={sc:.2f} "
                  f"T行={len(tp):2d} W字={len(wn):4d} T字={len(tn):4d}{flag} "
                  f"| {wn[:26]}")
        tot_lines += len(tp)
    print(f'\n配对段落覆盖 LaTeX 行数={tot_lines}  长度不符段={len(bad)}')
    for i, j, a, b in bad:
        print(f'   P{i}→T{j}  W={a} T={b}')
    if '--apply' not in args:
        print('\n（未修改文件，加 --apply 执行）')
        return
    print('\n--apply 尚未实现写入逻辑，先确认配对质量')


if __name__ == '__main__':
    main()
