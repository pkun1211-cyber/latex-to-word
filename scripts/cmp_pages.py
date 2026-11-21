# -*- coding: utf-8 -*-
"""逐页对齐诊断：页数 / 行数 / 首末行 / 累积偏移 / 字号分布。

用法: python cmp_pages.py [tex.json] [word.json]
"""
import os, sys, json, collections, difflib
import l2wconf as C

TMP = C.TMP


def load(tag):
    return json.load(open(os.path.join(TMP, tag + '.json'), encoding='utf-8'))


def norm(s):
    """归一化：去空白、去数学斜体差异，便于跨引擎比对。"""
    s = ''.join(s.split())
    return s


def main(tex_tag='tex2', word_tag='word2'):
    T, W = load(tex_tag), load(word_tag)
    print(f'页数  LaTeX={len(T)}  Word={len(W)}  {"OK" if len(T)==len(W) else "MISMATCH"}')
    print()
    print(f"{'页':>3} {'T行数':>6} {'W行数':>6} {'差':>5}   {'首行匹配':<8} {'末行匹配':<8}")
    print('-' * 74)
    for i in range(max(len(T), len(W))):
        t = T[i]['lines'] if i < len(T) else []
        w = W[i]['lines'] if i < len(W) else []
        # 首末行取「非空且长度>4」的行，避开公式碎片
        tf = next((L['t'] for L in t if len(norm(L['t'])) > 4), '')
        wf = next((L['t'] for L in w if len(norm(L['t'])) > 4), '')
        tl = next((L['t'] for L in reversed(t) if len(norm(L['t'])) > 4), '')
        wl = next((L['t'] for L in reversed(w) if len(norm(L['t'])) > 4), '')
        smf = difflib.SequenceMatcher(None, norm(tf)[:40], norm(wf)[:40]).ratio()
        sml = difflib.SequenceMatcher(None, norm(tl)[-40:], norm(wl)[-40:]).ratio()
        print(f'{i+1:>3} {len(t):>6} {len(w):>6} {len(w)-len(t):>+5}   '
              f'{smf:>7.2f}  {sml:>7.2f}')
    print()
    # 全文行序列对齐：找出第一个分歧点
    tl_all = [norm(L['t']) for P in T for L in P['lines'] if len(norm(L['t'])) > 3]
    wl_all = [norm(L['t']) for P in W for L in P['lines'] if len(norm(L['t'])) > 3]
    sm = difflib.SequenceMatcher(None, tl_all, wl_all, autojunk=False)
    print(f'全文有效行  LaTeX={len(tl_all)}  Word={len(wl_all)}  相似度={sm.ratio():.4f}')
    print('\n前 8 个分歧块:')
    n = 0
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            continue
        n += 1
        if n > 8:
            break
        print(f'  [{op}] T[{i1}:{i2}] W[{j1}:{j2}]')
        for s in tl_all[i1:min(i2, i1 + 2)]:
            print(f'      T: {s[:70]}')
        for s in wl_all[j1:min(j2, j1 + 2)]:
            print(f'      W: {s[:70]}')
    # 字号分布对比
    print('\n字号分布（按行数）:')
    for tag, arr in (('LaTeX', T), ('Word ', W)):
        c = collections.Counter()
        for P in arr:
            for L in P['lines']:
                c[L['size']] += 1
        print(f'  {tag}: ' + '  '.join(f'{k}pt×{v}' for k, v in sorted(c.items())))


if __name__ == '__main__':
    main(*(sys.argv[1:3] if len(sys.argv) > 2 else []))
