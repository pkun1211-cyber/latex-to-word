# -*- coding: utf-8 -*-
r"""直接把西文连字序列替换为 Unicode 连字字符（最可靠，不依赖 Word 的 OTF 特性开关）。

背景
----
w14:ligatures 实测无效（补齐命名空间后行数/断点仍完全不变），
Word 对该特性的支持依赖版本与渲染路径，不可靠。

方案
----
LM Roman 的 cmap 已直接映射 U+FB00..FB04（ﬀ ﬁ ﬂ ﬃ ﬄ），
把文本里的 ff/fi/fl/ffi/ffl 序列换成对应单字符即可，
字形与宽度与 TeX 完全一致（连字比分量之和窄 0.25~0.52pt @9pt）。

安全边界
--------
1. 只处理**纯西文 run**（run 文本全为 ASCII/拉丁扩展），避免动中文与公式
2. 跳过 OMML 内的一切内容
3. 跳过 URL / DOI / 文件路径（连字会破坏可点击性与可复制性）——
   实测 LaTeX 的 \url 也不连字，行为一致
4. 长序列优先：ffi/ffl 先于 ff/fi/fl，避免拆错
5. 保留 w:t 的 xml:space="preserve"
6. 记录替换清单便于回滚核对
"""
import sys, io, os, re, json, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
import l2wconf as C
import docxio as IO

DIR = C.DIR
# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
BAK = C.bak('lig2')
TMPD = C.TMP
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

# 长序列优先
LIG = [('ffi', '\ufb03'), ('ffl', '\ufb04'), ('ff', '\ufb00'),
       ('fi', '\ufb01'), ('fl', '\ufb02')]

URL_RE = re.compile(r'(https?://|www\.|doi[:/]|\.com|\.cn|\.org|\.pdf|\.docx)', re.I)


def is_latin(s):
    """判断是否纯西文（允许拉丁扩展、常用标点、空白）。"""
    for ch in s:
        o = ord(ch)
        if o < 0x250:            # ASCII + Latin-1 + Latin Extended-A/B
            continue
        if ch in '\u2013\u2014\u2018\u2019\u201c\u201d\u2026\u00a0':
            continue
        if 0xfb00 <= o <= 0xfb04:  # 已是连字
            continue
        return False
    return True


def convert(s):
    n = 0
    for a, b in LIG:
        c = s.count(a)
        if c:
            s = s.replace(a, b)
            n += c
    return s, n


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    if not os.path.exists(BAK):
        shutil.copy2(DOCX, BAK)
        print('备份 ->', BAK)
    d = Document(DOCX)
    log = []
    total = 0
    skipped_url = 0
    for r in d.element.iter(qn('w:r')):
        anc = r.getparent()
        in_math = False
        while anc is not None:
            if anc.tag.startswith(M):
                in_math = True
                break
            anc = anc.getparent()
        if in_math:
            continue
        for tn in r.findall(qn('w:t')):
            s = tn.text or ''
            if not s or not is_latin(s):
                continue
            if URL_RE.search(s):
                skipped_url += 1
                continue
            new, n = convert(s)
            if n:
                tn.text = new
                tn.set(qn('xml:space'), 'preserve')
                total += n
                log.append({'n': n, 'old': s[:70], 'new': new[:70]})
    IO.save(d, DOCX)
    json.dump(log, open(os.path.join(TMPD, 'lig2_log.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    print(f'连字替换 {total} 处，涉及 {len(log)} 个 w:t；跳过 URL 段 {skipped_url}')
    for it in log[:8]:
        print(f"  ×{it['n']}  {it['old'][:56]}")
    print('SAVED', os.path.getsize(DOCX))


if __name__ == '__main__':
    main()
