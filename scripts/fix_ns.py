# -*- coding: utf-8 -*-
"""在 document.xml / styles.xml 根元素补 w14 等命名空间声明。

python-docx 的默认 nsmap 不含 w14，导致 <w14:ligatures> 写出去后
Word 因命名空间未声明而整体忽略（实测启用连字后行数完全无变化）。
本脚本在 zip 层原地补声明，不改动其他内容。
"""
import sys, io, os, re, zipfile, shutil
import l2wconf as C
import docxio as IO
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DOCX = C.DOCX
TMPD = C.TMP

NS = {
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


def patch(xml, part):
    s = xml.decode('utf-8')
    # 找根元素起始标签
    m = re.search(r'<w:(document|styles)\b[^>]*>', s)
    if not m:
        print(f'  {part}: 未找到根元素')
        return xml, 0
    tag = m.group(0)
    add = []
    for pfx, uri in NS.items():
        if f'xmlns:{pfx}=' not in tag:
            add.append(f'xmlns:{pfx}="{uri}"')
    if not add:
        print(f'  {part}: 命名空间已齐全')
        return xml, 0
    newtag = tag[:-1] + ' ' + ' '.join(add) + '>'
    # mc:Ignorable 需含 w14 w15，否则 Word 可能报修复
    if 'mc:Ignorable' in newtag:
        newtag = re.sub(r'mc:Ignorable="([^"]*)"',
                        lambda mm: 'mc:Ignorable="' + ' '.join(
                            sorted(set(mm.group(1).split()) | {'w14', 'w15'})) + '"',
                        newtag)
    else:
        newtag = newtag[:-1] + ' mc:Ignorable="w14 w15">'
    s = s[:m.start()] + newtag + s[m.end():]
    print(f'  {part}: 补 {len(add)} 个命名空间声明')
    return s.encode('utf-8'), len(add)


def main():
    tmp = os.path.join(TMPD, '_ns_fixed.docx')
    zin = zipfile.ZipFile(DOCX, 'r')
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    total = 0
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename in ('word/document.xml', 'word/styles.xml'):
            data, n = patch(data, it.filename)
            total += n
        zout.writestr(it, data)
    zin.close()
    zout.close()
    for _ in range(6):
        try:
            IO.atomic_replace(tmp, DOCX)
            break
        except PermissionError:
            import time
            time.sleep(1.5)
    print(f'完成，共补 {total} 处；文件 {os.path.getsize(DOCX)} 字节')


if __name__ == '__main__':
    main()
