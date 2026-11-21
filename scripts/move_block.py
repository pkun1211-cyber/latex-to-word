# -*- coding: utf-8 -*-
r"""通用块移动：把「表题+表格+表注」或「图+图题」整块前移/后移。

相对 float_table.py 的改进
--------------------------
1. is_tabcap 不再排除含「、」的表题
   （表10「协同机制的子机制、功能与问题对应」因此被漏掉）
2. 支持 --before（向前移动），float_table 只能 --after
3. 支持图块（w:drawing 段 + 图题段）
4. 落点可用文本锚点指定，避免固定索引随改动失效

用法:
  python move_block.py --list
  python move_block.py 表10 --before-text "六、协同机制构建"
  python move_block.py 图1  --after-text  "机制的形成路径见图1。"
  python move_block.py --undo
"""
import sys, io, os, re, json, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO

DOCX = HA.DOCX
TMP = HA.TMP
BAK = os.path.join(TMP, 'bak_before_move.docx')
MAP = os.path.join(TMP, 'move_map.json')
W = HA.W


def tx(el):
    return ''.join(t.text or '' for t in el.iter() if t.tag in (W + 't', HA.M + 't'))


def is_tabcap(n):
    # 表题：以「表N」开头、长度适中、无句末标点（允许顿号）
    return (bool(re.match(r'^表\d+', n)) and 3 < len(n) <= 46
            and not re.search(r'[。；：]', n))


def is_figcap(n):
    return bool(re.match(r'^图\d+', n)) and 3 < len(n) <= 46 and not re.search(r'[。；：]', n)


def scan(d):
    body = d.element.body
    kids = list(body)
    blocks = []
    for i, el in enumerate(kids):
        if el.tag != qn('w:p'):
            continue
        n = HA.norm(tx(el))
        if is_tabcap(n):
            num = re.match(r'^表(\d+)', n).group(1)
            j, tbl_at = i + 1, None
            while j < len(kids) and j <= i + 4:
                if kids[j].tag == qn('w:tbl'):
                    tbl_at = j
                    break
                j += 1
            if tbl_at is None:
                continue
            k, note_at = tbl_at + 1, None
            while k < len(kids) and k <= tbl_at + 4:
                if kids[k].tag == qn('w:p'):
                    s = HA.norm(tx(kids[k]))
                    if s.startswith('注：'):
                        note_at = k
                        break
                    if s:
                        break
                k += 1
            end = note_at if note_at is not None else tbl_at
            blocks.append({'kind': '表', 'num': num, 'start': i, 'end': end,
                           'title': n[:40]})
        elif is_figcap(n):
            num = re.match(r'^图(\d+)', n).group(1)
            # 图题可能在图之前或之后：向前后各找 2 个含 drawing 的段
            s, e = i, i
            for j in range(max(0, i - 2), i):
                if kids[j].find('.//' + qn('w:drawing')) is not None:
                    s = j
                    break
            for j in range(i + 1, min(len(kids), i + 3)):
                if kids[j].find('.//' + qn('w:drawing')) is not None:
                    e = j
                    break
            blocks.append({'kind': '图', 'num': num, 'start': s, 'end': e,
                           'title': n[:40]})
    return blocks, kids


def find_anchor(kids, text):
    key = HA.norm(text)
    for i, el in enumerate(kids):
        if el.tag != qn('w:p'):
            continue
        if key in HA.norm(tx(el)):
            return i
    return None


def main():
    a = sys.argv[1:]
    d = Document(DOCX)
    blocks, kids = scan(d)
    if '--list' in a:
        print(f'块数={len(blocks)}')
        for b in blocks:
            print(f"  {b['kind']}{b['num']:<3s} body[{b['start']}:{b['end']}]  {b['title']}")
        return
    if '--undo' in a:
        shutil.copy2(BAK, DOCX)
        print('已回滚')
        return
    tgt = next((x for x in a if re.match(r'^[表图]\d+$', x)), None)
    if not tgt:
        print('需指定块号，如 表10 / 图1')
        return
    b = next((x for x in blocks if x['kind'] == tgt[0] and x['num'] == tgt[1:]), None)
    if b is None:
        print(f'未找到 {tgt}')
        return
    mode = None
    if '--before-text' in a:
        mode, key = 'before_text', a[a.index('--before-text') + 1]
    elif '--after-text' in a:
        mode, key = 'after_text', a[a.index('--after-text') + 1]
    elif '--before' in a:
        mode, key = 'before_idx', a[a.index('--before') + 1]
    elif '--after' in a:
        mode, key = 'after_idx', a[a.index('--after') + 1]
    else:
        print('需指定落点：--before-text / --after-text / --before / --after')
        return
    if mode.endswith('text'):
        ai = find_anchor(kids, key)
        if ai is None:
            print(f'锚点未找到: {key}')
            return
    else:
        ai = int(key)
    els = kids[b['start']:b['end'] + 1]
    anchor = kids[ai]
    print(f"移动 {tgt} body[{b['start']}:{b['end']}] ({len(els)}元素) "
          f"→ body{ai} {'之前' if mode.startswith('before') else '之后'}")
    print(f"  锚点: {HA.norm(tx(anchor))[:50]}")
    if not os.path.exists(BAK):
        shutil.copy2(DOCX, BAK)
        print('备份 ->', BAK)
    if mode.startswith('before'):
        for el in els:
            anchor.addprevious(el)
    else:
        cur = anchor
        for el in els:
            cur.addnext(el)
            cur = el
    log = json.load(open(MAP, encoding='utf-8')) if os.path.exists(MAP) else []
    log.append({'block': tgt, 'from': [b['start'], b['end']],
                'mode': mode, 'anchor': ai, 'n': len(els)})
    json.dump(log, open(MAP, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    IO.save(d, DOCX)
    print('SAVED', os.path.getsize(DOCX))


if __name__ == '__main__':
    main()
