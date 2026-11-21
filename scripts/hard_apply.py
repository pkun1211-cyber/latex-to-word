# -*- coding: utf-8 -*-
r"""逐行硬对齐的写入实现。

核心思路
--------
1. 建立 Word 段落的「字符流」：遍历所有 w:t（含 hyperlink 内），
   记录每个字符对应的 (节点, 偏移)。引文 [n] 的字符标记为 cite=True。
2. LaTeX 行文本已在提取时剔除上标引文，故用「去引文的归一化字符流」
   与 LaTeX 行序列做逐行消耗匹配，得到每个断点在字符流中的位置。
3. 在断点位置把 w:t 拆成两个 run，中间插 <w:br/>。
   拆分时深拷贝原 rPr，保证字体/字号/position 不丢。
4. 段落 jc 改 left（不再让 Word 拉伸），并按行写空格字距补偿。

空格补偿
--------
LaTeX 实测行宽 x1-x0，Word 自然行宽 = 各字符 advance 之和。
差值均摊到该行的西文空格上，写 <w:spacing w:val="N"/>（1/20 pt）。
只在该行有空格时补偿；纯中文行无空格可调，靠 br 已能锁定断点。

安全边界
--------
- 含 oMath / drawing 的段落跳过
- 表格单元格内段落跳过（由调用方过滤）
- 长度不符（Word 与 LaTeX 字数差 >10%）的段落跳过
- 断点匹配失败（累计误差 >3 字符）的段落跳过并记日志
"""
import sys, io, os, re, json, copy, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO

W = HA.W
M = HA.M


# ------------------------------------------------------------ 字符流
def char_stream(p_el):
    """返回 [(node, idx_in_text, ch, is_cite)]，按文档顺序。

    is_cite 判定：该 w:t 所在 run 的字号 <= 12（6pt）且处于引文括号内。
    简化实现：先扫一遍标记所有「sz<=12 且文本是 [数字,] 片段」的 run。
    """
    out = []
    for r in p_el.iter():
        if r.tag != W + 'r':
            continue
        # 该 run 是否为引文碎片
        rpr = r.find(qn('w:rPr'))
        sz = None
        if rpr is not None:
            e = rpr.find(qn('w:sz'))
            if e is not None:
                sz = int(e.get(qn('w:val')))
        for t in r.findall(qn('w:t')):
            txt = t.text or ''
            cite = (sz is not None and sz <= 12
                    and re.fullmatch(r'[\[\],\d\s]+', txt) is not None)
            # hyperlink 内的数字也算引文
            if not cite and sz is not None and sz <= 12 and re.fullmatch(r'\d+', txt):
                par = r.getparent()
                if par is not None and par.tag == W + 'hyperlink':
                    cite = True
            for k, ch in enumerate(txt):
                out.append((t, k, ch, cite))
    return out


def build_index(stream, drop_lead_cite=False):
    """去引文、去空白后的字符 → 原始流下标。

    drop_lead_cite: 参考文献段首的 [n] 是正常字号（不属于上标引文），
    但 LaTeX 提取时把它当作条目编号保留在行首；而 HA.mnorm 会删掉所有
    [数字]。为与 mnorm 口径一致，这里也删掉段首的 [n]。
    """
    idx = []
    chars = []
    for i, (t, k, ch, cite) in enumerate(stream):
        if cite:
            continue
        n = HA.norm(ch)
        if not n:
            continue
        for c in n:
            idx.append(i)
            chars.append(c)
    s = ''.join(chars)
    if drop_lead_cite:
        m = re.match(r'\[\d+\]', s)
        if m:
            s = s[m.end():]
            idx = idx[m.end():]
    return s, idx


# ------------------------------------------------------------ 断点定位
def find_breaks(p_el, tlines):
    """返回断点在字符流中的下标列表（在该字符之后换行）。

    逐行消耗：把 LaTeX 每行归一化后与 Word 归一化字符流对齐。
    LaTeX 行内可能残留公式符号（′ ∗ 数字上下标），用「跳过不匹配字符」容忍。
    """
    stream = char_stream(p_el)
    lead = re.match(r'^\[\d+\]', HA.norm(''.join(c for _, _, c, _ in stream)))
    wstr, idx = build_index(stream, drop_lead_cite=bool(lead))
    breaks = []
    pos = 0
    logs = []
    for li, L in enumerate(tlines[:-1]):
        ln = HA.mnorm(L['t'])
        if not ln:
            continue
        # 在 wstr[pos:] 中找 ln 的结束位置：允许少量字符不匹配
        end = match_end(wstr, pos, ln)
        if end is None:
            logs.append(f'行{li} 匹配失败 pos={pos} len={len(ln)}')
            return None, logs
        if end >= len(idx):
            logs.append(f'行{li} 越界')
            return None, logs
        breaks.append(idx[end - 1])
        pos = end
    # 末行必须能覆盖剩余（容忍尾部差异）
    last = HA.mnorm(tlines[-1]['t'])
    rest = wstr[pos:]
    if last and rest:
        r = min(len(last), len(rest))
        same = sum(1 for a, b in zip(last[:r], rest[:r]) if a == b)
        if same / r < 0.85:
            logs.append(f'末行相似度低 {same/r:.2f}')
            return None, logs
    return breaks, logs


def match_end(s, start, pat, max_skip=6):
    """在 s[start:] 中匹配 pat，返回消耗到的结束下标（exclusive）。

    容忍两类差异：
    1. s 中多出 pat 没有的字符（公式碎片、残留符号），最多 max_skip 个
    2. **TeX 自动断词连字符**：LaTeX 行末可能是 "The Ac-"，
       Word 对应位置是 "The Accounting"（无连字符）。
       pat 里的 '-' / '‐' / '­' 若在此处不匹配，直接跳过该 pat 字符。
    """
    HYPH = '-\u2010\u00ad\u2011'
    i, j = start, 0
    skip = 0
    while j < len(pat) and i < len(s):
        if s[i] == pat[j]:
            i += 1
            j += 1
            continue
        # TeX 断词连字符：pat 有 '-' 而 s 没有 → 跳过 pat 的这个字符
        if pat[j] in HYPH:
            j += 1
            continue
        # s 多出一个字符
        if skip < max_skip and i + 1 < len(s) and s[i + 1] == pat[j]:
            i += 1
            skip += 1
            continue
        return None
    # pat 末尾若剩连字符，同样跳过
    while j < len(pat) and pat[j] in HYPH:
        j += 1
    if j < len(pat):
        return None
    return i


# ------------------------------------------------------------ 插入 br
def insert_breaks(p_el, breaks):
    """在字符流下标处插 w:br。从后往前处理，避免下标失效。"""
    stream = char_stream(p_el)
    n = 0
    for bi in sorted(set(breaks), reverse=True):
        if bi >= len(stream):
            continue
        t, k, ch, cite = stream[bi]
        r = t.getparent()
        if r is None or r.tag != W + 'r':
            continue
        txt = t.text or ''
        # 断点在该字符之后 → 拆分点 k+1
        cut = k + 1
        if cut >= len(txt):
            # 正好在 run 末尾：直接在该 run 后插一个只含 br 的 run
            br_r = OxmlElement('w:r')
            rpr = r.find(qn('w:rPr'))
            if rpr is not None:
                br_r.append(copy.deepcopy(rpr))
            br_r.append(OxmlElement('w:br'))
            r.addnext(br_r)
            n += 1
            continue
        # run 内拆分
        left, right = txt[:cut], txt[cut:]
        t.text = left
        if left and (left[0] == ' ' or left[-1] == ' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_r = copy.deepcopy(r)
        # new_r 里只保留一个 w:t
        for extra in new_r.findall(qn('w:t'))[1:]:
            new_r.remove(extra)
        nt = new_r.find(qn('w:t'))
        nt.text = right
        if right and (right[0] == ' ' or right[-1] == ' '):
            nt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        br_r = OxmlElement('w:r')
        rpr = r.find(qn('w:rPr'))
        if rpr is not None:
            br_r.append(copy.deepcopy(rpr))
        br_r.append(OxmlElement('w:br'))
        r.addnext(br_r)
        br_r.addnext(new_r)
        n += 1
    return n


def set_left(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = OxmlElement('w:pPr')
        p_el.insert(0, ppr)
    jc = ppr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        ppr.append(jc)
    jc.set(qn('w:val'), 'left')
    HA.sort_el(ppr, HA.PPR_ORDER)


# ------------------------------------------------------------ 主流程
def run(docx=None, apply=False, only_body=False, limit=None):
    docx = docx or HA.DOCX
    d = Document(docx)
    pairs, TP, WP = HA.plan(d, only_body=only_body)
    ok, skip = [], []
    for w, tp, sc, j in pairs:
        wn = HA.mnorm(w['text'])
        tn = HA.mnorm(''.join(L['t'] for L in tp))
        lenr = min(len(wn), len(tn)) / max(len(wn), len(tn))
        if lenr < 0.9:
            skip.append((w['i'], j, f'长度不符 {len(wn)}/{len(tn)}'))
            continue
        if len(tp) < 2:
            skip.append((w['i'], j, '单行段无需 br'))
            continue
        breaks, logs = find_breaks(w['el'], tp)
        if breaks is None:
            skip.append((w['i'], j, '; '.join(logs)))
            continue
        ok.append((w, tp, breaks))
    print(f'可硬对齐段={len(ok)}  跳过={len(skip)}')
    tot_br = sum(len(b) for _, _, b in ok)
    print(f'将插入 br 总数={tot_br}')
    for w, tp, b in ok[:200]:
        print(f"  P{w['i']:<3d} T行={len(tp):2d} br={len(b):2d} | {HA.mnorm(w['text'])[:30]}")
    if skip:
        print('\n跳过明细:')
        for i, j, why in skip:
            print(f'  P{i}→T{j}  {why}')
    if not apply:
        print('\n（未修改文件，加 --apply 执行）')
        return
    bak = HA.BAK
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
        print(f'已备份 → {bak}')
    # limit 按「段落数」截断，且 jc 与 br 必须成对修改，
    # 否则未插 br 的段落被改成 left 会失去两端对齐、右界参差不齐。
    todo = ok[:limit] if limit else ok
    nbr = 0
    for w, tp, b in todo:
        nbr += insert_breaks(w['el'], b)
        set_left(w['el'])
    IO.save(d, docx)
    print(f'\n已写入：{len(todo)} 段，{nbr} 个 br → {os.path.basename(docx)}')


if __name__ == '__main__':
    a = sys.argv[1:]
    lim = None
    if '--limit' in a:
        k = a.index('--limit')
        lim = int(a[k + 1])
    run(apply='--apply' in a, only_body='--only-body' in a, limit=lim)
