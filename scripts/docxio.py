# -*- coding: utf-8 -*-
r"""docx 安全读写：关闭占用进程 + 原子替换 + 自动备份。

为什么需要
----------
Word 打开着目标 docx 时，`os.replace(tmp, docx)` 会抛
`PermissionError: [WinError 5] 拒绝访问`。
python-docx 的 d.save() 同样会失败。所有写盘操作都应走本模块。

用法
----
    import docxio as IO
    IO.save(doc, C.DOCX, tag='settings')      # 自动备份 + 原子写
    IO.replace_zip(C.DOCX, transform_fn)      # zip 内原地改写
"""
import os, sys, time, shutil, zipfile


def close_word(verbose=True):
    """关闭所有 Word 文档与实例，释放文件占用。"""
    try:
        import win32com.client as w
    except ImportError:
        return False
    try:
        app = w.GetActiveObject('Word.Application')
    except Exception:
        return False
    try:
        for i in range(app.Documents.Count, 0, -1):
            app.Documents.Item(i).Close(0)
        app.Quit()
        if verbose:
            print('  已关闭占用中的 Word 实例')
        time.sleep(1.0)
        return True
    except Exception:
        return False


def atomic_replace(tmp, dst, retries=4):
    """原子替换；遇占用先关 Word 再重试。"""
    for k in range(retries):
        try:
            os.replace(tmp, dst)
            return
        except PermissionError:
            if k == retries - 1:
                raise
            close_word()
            time.sleep(1.0)


def backup(docx, tag, tmpdir=None):
    """首次修改前备份（同名备份已存在则不覆盖，保留最早原件）。"""
    if not docx or not os.path.exists(docx):
        return None
    tmpdir = tmpdir or os.path.dirname(docx)
    # 目录在此按需创建：l2wconf.tmp() 只拼路径不建目录（避免 import 副作用）
    os.makedirs(tmpdir, exist_ok=True)
    bak = os.path.join(tmpdir, 'bak_before_%s.docx' % tag)
    if not os.path.exists(bak):
        shutil.copy2(docx, bak)
    return bak


def save(doc, docx, tag=None, tmpdir=None, close=True):
    """保存 python-docx Document，自动备份并处理文件占用。"""
    if tag:
        backup(docx, tag, tmpdir)
    if close:
        close_word(verbose=False)
    for k in range(4):
        try:
            doc.save(docx)
            return
        except PermissionError:
            if k == 3:
                raise
            close_word()
            time.sleep(1.0)


def replace_zip(docx, transform, tag=None, tmpdir=None):
    """zip 内原地改写。transform(filename, bytes) -> bytes。"""
    if tag:
        backup(docx, tag, tmpdir)
    close_word(verbose=False)
    tmp = docx + '.tmp'
    changed = []
    zin = zipfile.ZipFile(docx, 'r')
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        new = transform(item.filename, data)
        if new is not None and new != data:
            changed.append(item.filename)
            data = new
        zout.writestr(item, data)
    zin.close()
    zout.close()
    atomic_replace(tmp, docx)
    return changed
