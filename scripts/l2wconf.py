# -*- coding: utf-8 -*-
r"""latex-to-word 统一路径配置。

所有脚本通过本模块取路径，不再硬编码项目目录。

配置来源（优先级从高到低）
--------------------------
1. 环境变量 L2W_CONF 指向的 JSON 文件
2. 当前工作目录下的 l2w.json
3. 环境变量 L2W_BASE / L2W_DOCX / L2W_TEX / L2W_TMP
4. 默认值（BASE=cwd，TMP=BASE/.workbuddy/tmp）

l2w.json 字段
-------------
{
  "base":  "D:/proj",                     # 项目根
  "docx":  "D:/proj/out/paper.docx",      # 目标 Word（会被就地修改）
  "tex":   "D:/proj/paper.tex",           # LaTeX 源
  "texpdf":"D:/proj/paper.pdf",           # LaTeX 编译产物（对齐基准）
  "tmp":   "D:/proj/.workbuddy/tmp",      # 中间产物目录
  "ttfdir":".../skills/latex-to-word/assets/ttf_lm",  # LM Roman TTF
  "tex_tag": "tex",                       # LaTeX 行提取 json 的标签
  "outname": "paper_word.docx"            # 首次转换的输出名
}

用法
----
    import l2wconf as C
    d = Document(C.DOCX)
    tmp = C.TMP
"""
import os, json

_HERE = os.path.dirname(os.path.abspath(__file__))
_SKILL = os.path.dirname(_HERE)


def _load():
    cfg = {}
    p = os.environ.get('L2W_CONF')
    if p and os.path.exists(p):
        cfg = json.load(open(p, encoding='utf-8'))
    elif os.path.exists('l2w.json'):
        cfg = json.load(open('l2w.json', encoding='utf-8'))
    for k, env in (('base', 'L2W_BASE'), ('docx', 'L2W_DOCX'),
                   ('tex', 'L2W_TEX'), ('texpdf', 'L2W_TEXPDF'),
                   ('tmp', 'L2W_TMP'), ('ttfdir', 'L2W_TTFDIR')):
        v = os.environ.get(env)
        if v:
            cfg[k] = v
    return cfg


_C = _load()

BASE = os.path.abspath(_C.get('base') or os.getcwd())
TMP = os.path.abspath(_C.get('tmp') or os.path.join(BASE, '.workbuddy', 'tmp'))
DOCX = os.path.abspath(_C['docx']) if _C.get('docx') else ''
DIR = os.path.dirname(DOCX) if DOCX else BASE
TEX = os.path.abspath(_C['tex']) if _C.get('tex') else ''
TEXPDF = os.path.abspath(_C['texpdf']) if _C.get('texpdf') else ''
TTFDIR = os.path.abspath(_C.get('ttfdir') or os.path.join(_SKILL, 'assets', 'ttf_lm'))
TEX_TAG = _C.get('tex_tag', 'tex')
OUTNAME = _C.get('outname', 'paper_word.docx')
SKILL = _SKILL
ASSETS = os.path.join(_SKILL, 'assets')

def ensure_tmp():
    """按需创建中间目录。

    注意：**不要在模块顶层调用**。早期版本顶层执行 os.makedirs(TMP)，
    导致 import 本模块就在当前工作目录留下 .workbuddy/tmp（副作用），
    在技能目录里跑自检时会污染技能包。
    """
    os.makedirs(TMP, exist_ok=True)
    return TMP


def tmp(*parts):
    """中间产物路径。**纯路径拼接，不建目录**——顶层调用不得有副作用。

    目录由写盘环节创建（docxio.save / ensure_tmp）。
    """
    return os.path.join(TMP, *parts)


def bak(name):
    """备份路径（统一放 TMP，命名 bak_before_<name>.docx）。"""
    return tmp('bak_before_%s.docx' % name)


def require_docx():
    if not DOCX or not os.path.exists(DOCX):
        raise SystemExit(
            '未配置目标 docx。请设置环境变量 L2W_DOCX，或在当前目录放 l2w.json。\n'
            '示例: {"base":"D:/proj","docx":"D:/proj/out/paper.docx",'
            '"tex":"D:/proj/paper.tex","texpdf":"D:/proj/paper.pdf"}')
    return DOCX


def describe():
    return ('BASE=%s\nTMP=%s\nDOCX=%s\nTEX=%s\nTEXPDF=%s\nTTFDIR=%s'
            % (BASE, TMP, DOCX, TEX, TEXPDF, TTFDIR))


if __name__ == '__main__':
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    print(describe())
