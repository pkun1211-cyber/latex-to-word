# -*- coding: utf-8 -*-
"""可靠的 PDF 行提取（PyMuPDF span 版），替代 pypdfium2 charbox 方案。

为什么换实现
------------
pypdfium2 的 get_charbox() 对部分 CFF/Type1 字形返回异常 box（高度可达 400pt），
按 baseline 聚类时会把同一视觉行拆成两条互补碎片（如 "overnance" / "eectsestates"），
导致行数统计与断行比对全部失真。

PyMuPDF 的 page.get_text("dict") 直接给出 span 级 bbox / size / font / origin，
用 span 的 origin.y（基线）聚类即可稳定还原视觉行，且天然保持阅读顺序。

输出每行：
  y      基线（PDF 坐标，原点左上，向下为正）
  size   该行主字号（按字符数加权众数）
  x0/x1  行左右界（已剔除行尾空白字符的 advance）
  fonts  该行用到的字体集合
  t      行文本
"""
import os, sys, json, collections
import pymupdf
import l2wconf as C

BASE = C.BASE
TMP = C.TMP
TEX = C.TEXPDF          # LaTeX 编译产物，作为对齐基准


def page_lines(page, tol=1.6, merge_scripts=True):
    """按 span 基线聚类还原视觉行。

    merge_scripts: Word 的 OMML 公式把上下标渲染成独立 span，
    基线比主行低/高 1~2pt，会被聚类成额外的"行"，虚增行数
    （实测同一公式 LaTeX 记 1 行、Word 记 2 行）。
    这里把「字号明显更小且基线差 < 3pt」的碎片并回主行。
    """
    d = page.get_text('dict')
    spans = []
    for blk in d['blocks']:
        if blk['type'] != 0:
            continue
        for ln in blk['lines']:
            for sp in ln['spans']:
                txt = sp['text']
                if not txt.strip():
                    # 纯空白 span 仍参与，但不影响 ink 边界
                    pass
                spans.append({
                    'y': sp['origin'][1],
                    'x0': sp['bbox'][0], 'x1': sp['bbox'][2],
                    'size': sp['size'], 'font': sp['font'],
                    'text': txt, 'asc': sp.get('ascender', 0), 'desc': sp.get('descender', 0),
                })
    if not spans:
        return [], []
    spans.sort(key=lambda s: (s['y'], s['x0']))
    groups = []
    for sp in spans:
        if groups and abs(groups[-1][0]['y'] - sp['y']) <= tol:
            groups[-1].append(sp)
        else:
            groups.append([sp])
    # 合并公式上下标碎片：小字号 + 基线接近的相邻组并入主行
    if merge_scripts and len(groups) > 1:
        merged = [groups[0]]
        for g in groups[1:]:
            prev = merged[-1]
            py = prev[0]['y']
            gy = g[0]['y']
            psz = max(s['size'] for s in prev)
            gsz = max(s['size'] for s in g)
            # 碎片特征：基线差 < 3pt 且字号比主行小 15% 以上
            if abs(gy - py) < 3.0 and gsz < psz * 0.85:
                prev.extend(g)
            else:
                merged.append(g)
        groups = merged

    out = []
    for g in groups:
        g.sort(key=lambda s: s['x0'])
        txt = ''.join(s['text'] for s in g)
        # 剔除行尾空白后的右界：逐字符估算不便，改为剔除末尾纯空白 span
        gg = list(g)
        while gg and not gg[-1]['text'].strip():
            gg.pop()
        x1 = max(s['x1'] for s in gg) if gg else max(s['x1'] for s in g)
        # 若末 span 以空格结尾，按其空格数 × 空格宽近似回退
        if gg:
            last = gg[-1]
            tail = len(last['text']) - len(last['text'].rstrip())
            if tail:
                # 用该 span 平均字宽的一半近似空格宽
                n = max(1, len(last['text'].strip()))
                sw = (last['x1'] - last['x0']) / max(1, len(last['text'])) * 0.5
                x1 -= tail * sw
        szc = collections.Counter()
        for s in g:
            szc[round(s['size'], 2)] += len(s['text'].strip())
        size = szc.most_common(1)[0][0] if szc else 0
        out.append({
            'y': round(g[0]['y'], 2),
            'size': size,
            'x0': round(min(s['x0'] for s in g), 2),
            'x1': round(x1, 2),
            'fonts': sorted({s['font'].split('+')[-1] for s in g}),
            't': txt,
        })
    imgs = []
    for blk in d['blocks']:
        if blk['type'] == 1:
            imgs.append([round(v, 2) for v in blk['bbox']])
    return out, imgs


def extract(path, out_json=None, out_txt=None):
    doc = pymupdf.open(path)
    pages = []
    for pno in range(len(doc)):
        lines, imgs = page_lines(doc[pno])
        pages.append({'page': pno + 1, 'lines': lines, 'images': imgs})
    doc.close()
    if out_json:
        json.dump(pages, open(out_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    if out_txt:
        with open(out_txt, 'w', encoding='utf-8') as f:
            for P in pages:
                f.write(f"\n{'='*100}\nPAGE {P['page']}  行数={len(P['lines'])}  图={P['images']}\n{'='*100}\n")
                for L in P['lines']:
                    f.write(f"  y={L['y']:7.2f} sz={L['size']:5.2f} x0={L['x0']:6.1f} x1={L['x1']:6.1f} "
                            f"[{','.join(L['fonts'])[:34]:<34s}] {L['t']}\n")
    return pages


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else TEX
    tag = sys.argv[2] if len(sys.argv) > 2 else 'tex2'
    pages = extract(src, os.path.join(TMP, tag + '.json'), os.path.join(TMP, tag + '.txt'))
    print(f'{os.path.basename(src)}  页数={len(pages)}')
    for P in pages:
        ls = P['lines']
        print(f"  p{P['page']:2d} 行数={len(ls):3d}  首={ls[0]['t'][:38] if ls else ''!r}")
