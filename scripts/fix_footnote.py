# -*- coding: utf-8 -*-
r"""复刻 LaTeX 脚注排版。

LaTeX 实测（\footnotesize 在 zihao=5 下约 7.5pt）：
  字号 7.5pt，行距 9pt（y=745.4 → 754.4 → 763.4）
  脚注区起点 y≈742（正文末行 727.3 + 约 15pt）
Word 现状：
  FootnoteText 样式，run 字号 16(8pt)，行距继承样式（约 13.6pt）
  → 脚注比 LaTeX 高约 14pt，把正文末 2 行挤到下一页
  （实测 p7 少装 2 行、p8 多出 2 行）

做法：
  1. 改 FootnoteText 样式：sz=15(7.5pt)、spacing line=180(9pt) exact
  2. 改脚注 run：sz=15、西文 LM Roman 7（按 7.5pt 光学尺寸）
  3. 脚注编号上标保持 vertAlign（LaTeX 也是上标）
"""
import sys, io, os, re, zipfile, shutil
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import hard_align as HA
import docxio as IO


def lm_family(pt):
    if pt <= 5.5:
        return 'LM Roman 5'
    if pt <= 6.5:
        return 'LM Roman 6'
    if pt <= 7.5:
        return 'LM Roman 7'
    if pt <= 8.5:
        return 'LM Roman 8'
    if pt <= 9.5:
        return 'LM Roman 9'
    return 'LM Roman 10'


def fix_style(d, sz=15, line=180):
    n = 0
    for s in d.styles.element.findall(qn('w:style')):
        sid = s.get(qn('w:styleId'))
        if sid not in ('FootnoteText', 'FootnoteReference'):
            continue
        rpr = s.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr')
            s.append(rpr)
        for tag, val in (('sz', str(sz)), ('szCs', str(sz))):
            e = rpr.find(qn('w:' + tag))
            if e is None:
                e = OxmlElement('w:' + tag)
                rpr.append(e)
            e.set(qn('w:val'), val)
        f = rpr.find(qn('w:rFonts'))
        if f is None:
            f = OxmlElement('w:rFonts')
            rpr.insert(0, f)
        fam = lm_family(sz / 2.0)
        f.set(qn('w:ascii'), fam)
        f.set(qn('w:hAnsi'), fam)
        HA.sort_el(rpr, HA.RPR_ORDER)
        if sid == 'FootnoteText':
            ppr = s.find(qn('w:pPr'))
            if ppr is None:
                ppr = OxmlElement('w:pPr')
                s.insert(1, ppr)
            sp = ppr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                ppr.append(sp)
            sp.set(qn('w:line'), str(line))
            sp.set(qn('w:lineRule'), 'exact')
            sp.set(qn('w:before'), '0')
            sp.set(qn('w:after'), '0')
            HA.sort_el(ppr, HA.PPR_ORDER)
        n += 1
    return n


def fix_footnotes_xml(docx, sz=15, line=180):
    """直接改 word/footnotes.xml：run 字号 + 段落 spacing。"""
    tmp = docx + '.tmp'
    zin = zipfile.ZipFile(docx)
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    fam = lm_family(sz / 2.0)
    changed = False
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/footnotes.xml':
            x = data.decode('utf-8')
            # run 字号
            x = x.replace('<w:sz w:val="16"/>', f'<w:sz w:val="{sz}"/>')
            x = x.replace('<w:sz w:val="18"/>', f'<w:sz w:val="{sz}"/>')
            # 西文字体
            x = x.replace('w:ascii="Times New Roman" w:hAnsi="Times New Roman"',
                          f'w:ascii="{fam}" w:hAnsi="{fam}"')
            # 给每个脚注段加 spacing（在 pStyle 之后）
            x = re.sub(r'(<w:pStyle w:val="FootnoteText"/>)(?!<w:spacing)',
                       r'\1' + f'<w:spacing w:line="{line}" w:lineRule="exact" '
                               f'w:before="0" w:after="0"/>', x)
            data = x.encode('utf-8')
            changed = True
        zout.writestr(item, data)
    zin.close()
    zout.close()
    IO.atomic_replace(tmp, docx)
    return changed


def run(apply=False, sz=15, line=180):
    d = Document(HA.DOCX)
    n = fix_style(d, sz, line)
    print(f'样式修改={n}  字号={sz/2}pt 行距={line/20}pt 西文={lm_family(sz/2.0)}')
    if not apply:
        print('（未修改，加 --apply）')
        return
    bak = os.path.join(HA.TMP, 'bak_before_footnote.docx')
    if not os.path.exists(bak):
        shutil.copy2(HA.DOCX, bak)
    IO.save(d, HA.DOCX)
    ok = fix_footnotes_xml(HA.DOCX, sz, line)
    print('footnotes.xml 已改' if ok else 'footnotes.xml 未找到')


if __name__ == '__main__':
    a = sys.argv[1:]
    g = lambda k, dv: int(a[a.index(k) + 1]) if k in a else dv
    run(apply='--apply' in a, sz=g('--sz', 15), line=g('--line', 180))
