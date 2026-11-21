# -*- coding: utf-8 -*-
r"""模拟 LaTeX 浮动体：把「表题+表格+表注」整块移到正文之后。

为什么不能用 pageBreakBefore
----------------------------
LaTeX 的 \begin{table}[htbp] 是**浮动体**：它被移到次页顶部，
而原本位置的正文会**继续向前填充本页**。
Word 的表格是流内元素，给表题加 pageBreakBefore 只会把
「表题及其后所有内容」一起推到次页，本页留下大片空白
（实测 p3 在 y=554 就停了，留白 219pt，页数还从 12 涨到 13）。

正确做法
--------
把表块整体**移动到后续正文之后**，让正文先填满本页，
表格自然落到下一页顶部 —— 这才是浮动体的等价实现。

块边界识别
----------
表块 = [表题段, (空段), w:tbl, (空段), 表注段]
从表题段起，到「注：」段止（含中间的空段与表格）。

用法:
  python float_table.py 表2 --after 43     # 移到 P43 之后
  python float_table.py 表2 --auto         # 自动选落点（下一个 H1/H2 之前）
  python float_table.py --list             # 列出所有表块及其 body 索引
  python float_table.py --undo             # 从备份恢复
"""
import sys, io, os, re, json, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import l2wconf as C
import docxio as IO

DIR = C.DIR
# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
TMP = C.TMP
BAK = os.path.join(TMP, 'bak_before_float.docx')
MAP = os.path.join(TMP, 'float_map.json')
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def tx(el):
    return ''.join(t.text or '' for t in el.iter() if t.tag in (W + 't', M + 't'))


def is_tabcap(n):
    return bool(re.match(r'^表\d+', n)) and 3 < len(n) <= 40 and not re.search(r'[。，；：、]', n)


def scan(d):
    """返回 [{num, start, end, cap_idx, body_els, para_range}]"""
    body = d.element.body
    kids = list(body)
    # body 索引 -> 段落序号
    pidx = {}
    pi = 0
    for i, el in enumerate(kids):
        if el.tag == qn('w:p'):
            pidx[i] = pi
            pi += 1
    blocks = []
    for i, el in enumerate(kids):
        if el.tag != qn('w:p'):
            continue
        n = ''.join(tx(el).split())
        if not is_tabcap(n):
            continue
        num = re.match(r'^表(\d+)', n).group(1)
        # 向后找到 w:tbl，再找到「注：」段
        j = i + 1
        tbl_at = None
        while j < len(kids) and j <= i + 4:
            if kids[j].tag == qn('w:tbl'):
                tbl_at = j
                break
            j += 1
        if tbl_at is None:
            continue
        k = tbl_at + 1
        note_at = None
        while k < len(kids) and k <= tbl_at + 4:
            if kids[k].tag == qn('w:p'):
                s = ''.join(tx(kids[k]).split())
                if s.startswith('注：'):
                    note_at = k
                    break
                if s:
                    break
            k += 1
        end = note_at if note_at is not None else tbl_at
        blocks.append({'num': num, 'start': i, 'end': end,
                       'cap_p': pidx.get(i), 'note_p': pidx.get(end),
                       'has_note': note_at is not None,
                       'title': n[:36]})
    return blocks, kids, pidx


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    args = sys.argv[1:]
    d = Document(DOCX)
    blocks, kids, pidx = scan(d)
    if '--list' in args:
        print(f'表块数={len(blocks)}')
        for b in blocks:
            print(f"  表{b['num']:<3s} body[{b['start']}:{b['end']}] "
                  f"P{b['cap_p']}→P{b['note_p']} note={b['has_note']}  {b['title']}")
        return
    if '--undo' in args:
        shutil.copy2(BAK, DOCX)
        print('已从备份恢复')
        return
    tgt = next((a for a in args if a.startswith('表')), None)
    if tgt is None:
        print('需指定表号，如 表2')
        return
    num = tgt[1:]
    b = next((x for x in blocks if x['num'] == num), None)
    if b is None:
        print(f'未找到 表{num}')
        return
    if not os.path.exists(BAK):
        shutil.copy2(DOCX, BAK)
        print('备份 ->', BAK)

    # 确定落点
    after_p = None
    if '--after' in args:
        after_p = int(args[args.index('--after') + 1])
    else:
        # 自动：表块之后的第一个 H1/H2 之前的最后一段
        for i in range(b['end'] + 1, len(kids)):
            el = kids[i]
            if el.tag != qn('w:p'):
                continue
            st = el.find(qn('w:pPr'))
            sid = None
            if st is not None:
                e = st.find(qn('w:pStyle'))
                if e is not None:
                    sid = e.get(qn('w:val'))
            if sid in ('Heading1', 'Heading2'):
                after_p = pidx[i] - 1
                break
        if after_p is None:
            print('自动落点失败，请用 --after 指定')
            return
    # 落点 body 索引
    dst_i = next((i for i, p in pidx.items() if p == after_p), None)
    if dst_i is None:
        print(f'落点 P{after_p} 不存在')
        return
    if dst_i <= b['end']:
        print(f'落点 P{after_p}(body{dst_i}) 必须在表块之后(body{b["end"]})')
        return

    els = kids[b['start']:b['end'] + 1]
    anchor = kids[dst_i]
    print(f"移动 表{num} 块 body[{b['start']}:{b['end']}] ({len(els)}个元素) "
          f"→ P{after_p}(body{dst_i}) 之后")
    print(f"  落点段落: {tx(anchor).strip()[:50]}")
    for el in els:
        anchor.addnext(el)
        anchor = el
    # 记录映射
    log = []
    if os.path.exists(MAP):
        log = json.load(open(MAP, encoding='utf-8'))
    log.append({'table': num, 'from_body': [b['start'], b['end']],
                'after_para': after_p, 'n_els': len(els)})
    json.dump(log, open(MAP, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    IO.save(d, DOCX)
    print('SAVED', os.path.getsize(DOCX))


if __name__ == '__main__':
    main()
