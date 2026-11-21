# -*- coding: utf-8 -*-
r"""settings.xml 与 pPr 合法性修正。

1. characterSpacingControl: doNotCompress → compressPunctuation
   LaTeX/ctex 默认挤压中文标点（句末标点占半宽），Word 默认不压缩，
   导致每行少装 1—2 字、断行位置与 LaTeX 不一致。
2. 关闭 docGrid 对行的吸附（snapToGrid），避免固定行距被网格改写。
3. w:pPr 子元素必须按 OOXML schema 顺序排列，否则 Word 静默丢弃部分属性。
   正确顺序：pStyle → keepNext → keepLines → pageBreakBefore → framePr →
   widowControl → numPr → suppressLineNumbers → pBdr → shd → tabs →
   suppressAutoHyphens → kinsoku → wordWrap → overflowPunct →
   topLinePunct → autoSpaceDE → autoSpaceDN → bidi → adjustRightInd →
   snapToGrid → spacing → ind → contextualSpacing → mirrorIndents →
   suppressOverlap → jc → textDirection → textAlignment →
   textboxTightWrap → outlineLvl → divId → cnfStyle → rPr → sectPr → pPrChange
"""
import sys, io, os, re, zipfile, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import l2wconf as C
import docxio as IO

# 容错取值：import 阶段不校验配置，否则脱离项目目录时无法加载本模块。
# 真正的存在性校验在 main() 入口做（C.require_docx()）。
DOCX = C.DOCX
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
             'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
             'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
             'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
             'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
             'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr',
             'pPrChange']
RANK = {W + ':' + n: i for i, n in enumerate(PPR_ORDER)}
RANK = {'{' + W + '}' + n: i for i, n in enumerate(PPR_ORDER)}


def sort_ppr(ppr):
    kids = list(ppr)
    want = sorted(kids, key=lambda e: RANK.get(e.tag, 999))
    if [e.tag for e in kids] == [e.tag for e in want]:
        return False
    for e in kids:
        ppr.remove(e)
    for e in want:
        ppr.append(e)
    return True


def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验，import 不受影响
    d = Document(DOCX)
    n = 0
    for ppr in d.element.body.iter(qn('w:pPr')):
        if sort_ppr(ppr):
            n += 1
    # 样式里的 pPr 也要排序
    for ppr in d.styles.element.iter(qn('w:pPr')):
        if sort_ppr(ppr):
            n += 1
    print('pPr 重排:', n)

    # 关闭 snapToGrid（段落级），保证 exact 行距不被网格改写
    nsnap = 0
    for p in d.paragraphs:
        ppr = p._p.find(qn('w:pPr'))
        if ppr is None:
            continue
        sg = ppr.find(qn('w:snapToGrid'))
        if sg is None:
            sg = etree.Element(qn('w:snapToGrid'))
            sp = ppr.find(qn('w:spacing'))
            if sp is not None:
                sp.addprevious(sg)
            else:
                ppr.append(sg)
            sort_ppr(ppr)
        sg.set(qn('w:val'), '0')
        nsnap += 1
    print('snapToGrid=0 段落:', nsnap)

    IO.save(d, DOCX)

    # ---- settings.xml：zip 内原地替换 ----
    tmp = DOCX + '.tmp'
    zin = zipfile.ZipFile(DOCX, 'r')
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    changed = []
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/settings.xml':
            s = data.decode('utf-8')
            before = s
            # compressPunctuation 会连西文字距一起压，英文 Abstract 会少 2 行；
            # 用 compressPunctuationAndJapaneseKana 仍影响西文，故保持 doNotCompress，
            # 中文断行差异改由段落级 w:w（字符缩放）微调解决。
            s = re.sub(r'<w:characterSpacingControl w:val="[^"]*"/>',
                       '<w:characterSpacingControl w:val="doNotCompress"/>', s)
            if 'characterSpacingControl' not in s:
                s = s.replace('</w:settings>',
                              '<w:characterSpacingControl w:val="doNotCompress"/></w:settings>')
            if s != before:
                changed.append('characterSpacingControl→doNotCompress')
            data = s.encode('utf-8')
        zout.writestr(item, data)
    zin.close(); zout.close()
    IO.atomic_replace(tmp, DOCX)
    print('settings 修改:', changed or '无')
    print('SAVED', os.path.getsize(DOCX))


if __name__ == '__main__':
    main()
