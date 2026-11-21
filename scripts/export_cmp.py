# -*- coding: utf-8 -*-
"""通用：docx -> Word COM 导出 PDF -> PyMuPDF 行提取 -> 与 LaTeX 逐页对比。

用法:
  python export_cmp.py <docx路径> <标签>
例:
  python export_cmp.py exp_punct_on.docx pON
输出:
  <标签>.pdf / <标签>.json / <标签>.txt  + 控制台对比摘要
"""
import os, sys, io, time, json, collections, difflib
if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import pythoncom
import win32com.client as w

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdflines2 import extract
import l2wconf as C

TMP = C.TMP


def kill_word():
    try:
        app = w.GetActiveObject('Word.Application')
        for i in range(app.Documents.Count, 0, -1):
            app.Documents.Item(i).Close(0)
        app.Quit()
    except Exception:
        pass


def get_app():
    for _ in range(5):
        try:
            app = w.Dispatch('Word.Application')
            app.Visible = False
            app.DisplayAlerts = 0
            return app
        except pythoncom.com_error:
            kill_word()
            time.sleep(2)
    raise RuntimeError('无法启动 Word')


def export(docx, pdf):
    kill_word()
    app = get_app()
    try:
        doc = app.Documents.Open(os.path.abspath(docx))
        doc.Repaginate()
        n = doc.ComputeStatistics(2)
        if os.path.exists(pdf):
            for _ in range(5):
                try:
                    os.remove(pdf)
                    break
                except PermissionError:
                    time.sleep(1)
        doc.ExportAsFixedFormat(pdf, 17)
        doc.Close(0)
        return n
    finally:
        try:
            app.Quit()
        except Exception:
            pass


def norm(s):
    return ''.join(s.split())


def compare(word_tag, tex_tag='tex2', brief=True):
    T = json.load(open(os.path.join(TMP, tex_tag + '.json'), encoding='utf-8'))
    W = json.load(open(os.path.join(TMP, word_tag + '.json'), encoding='utf-8'))
    print(f'\n页数 LaTeX={len(T)} Word={len(W)} {"OK" if len(T)==len(W) else "MISMATCH"}')
    print(f"{'页':>3} {'T行':>4} {'W行':>4} {'差':>4}  {'首匹配':>6} {'末匹配':>6}")
    tot = 0
    for i in range(max(len(T), len(W))):
        t = T[i]['lines'] if i < len(T) else []
        ww = W[i]['lines'] if i < len(W) else []
        tf = next((L['t'] for L in t if len(norm(L['t'])) > 4), '')
        wf = next((L['t'] for L in ww if len(norm(L['t'])) > 4), '')
        tl = next((L['t'] for L in reversed(t) if len(norm(L['t'])) > 4), '')
        wl = next((L['t'] for L in reversed(ww) if len(norm(L['t'])) > 4), '')
        a = difflib.SequenceMatcher(None, norm(tf)[:40], norm(wf)[:40]).ratio()
        b = difflib.SequenceMatcher(None, norm(tl)[-40:], norm(wl)[-40:]).ratio()
        tot += abs(len(ww) - len(t))
        print(f'{i+1:>3} {len(t):>4} {len(ww):>4} {len(ww)-len(t):>+4}  {a:>6.2f} {b:>6.2f}')
    print(f'行数绝对偏差合计 = {tot}')
    ta = [norm(L['t']) for P in T for L in P['lines'] if len(norm(L['t'])) > 3]
    wa = [norm(L['t']) for P in W for L in P['lines'] if len(norm(L['t'])) > 3]
    sm = difflib.SequenceMatcher(None, ta, wa, autojunk=False)
    print(f'全文行相似度 = {sm.ratio():.4f}  (T={len(ta)} W={len(wa)})')
    # 中文满行右界档位
    for tag, arr in (('LaTeX', T), ('Word ', W)):
        c = collections.Counter()
        for P in arr:
            for L in P['lines']:
                if L['size'] == 9.0 and sum(1 for ch in L['t'] if '\u4e00' <= ch <= '\u9fff') > 20:
                    c[round(L['x1'])] += 1
        top = sorted(c.items(), key=lambda z: -z[1])[:6]
        print(f'  {tag} 中文满行右界档位: {top}')
    return tot


if __name__ == '__main__':
    docx = sys.argv[1]
    tag = sys.argv[2]
    if not os.path.isabs(docx):
        docx = os.path.join(TMP, docx)
    pdf = os.path.join(TMP, tag + '.pdf')
    n = export(docx, pdf)
    print(f'COM 页数 = {n}')
    extract(pdf, os.path.join(TMP, tag + '.json'), os.path.join(TMP, tag + '.txt'))
    compare(tag)
