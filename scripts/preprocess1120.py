# -*- coding: utf-8 -*-
"""把当前 ctex 中文 LaTeX 预处理为 Pandoc 可高保真解析的形式（11.20 版）。

占位标记（Pandoc 原样保留，供后处理重建）：
- 引用：⟦CITE:1,2⟧
- 交叉引用：⟦REF:表3⟧ / ⟦REF:式1⟧ / ⟦REF:图1⟧
- 文献条目：⟦BIB:n⟧
- 表注：⟦TABNOTE⟧ 前缀
- 脚注：⟦FOOTNOTE:...⟧
"""
import os, re
import l2wconf as C

BASE = C.BASE
OUT = C.tmp("pre.tex")

def convert_bib_block(block, bib_key2num):
    """把 \\bibitem 转成 ⟦BIB:n⟧ 占位，并回填 key→编号映射。

    映射表必须显式传入：模块级不得持有可变状态，
    否则 import 本模块就会带上一次运行的残留。
    """
    out, n = [], 0
    for ln in block.split("\n"):
        m = re.match(r'(\s*)\\bibitem(?:\[(\d+)\])?\{([^}]+)\}\s*(.*)', ln)
        if m:
            lead, opt, key, content = m.groups()
            n = int(opt) if opt else n + 1
            bib_key2num[key] = n
            out.append(f"{lead}⟦BIB:{n}⟧ {content}")
        else:
            out.append(ln)
    return "\n".join(out)


def main():
    # tex 源在运行时才解析，避免 import 阶段就强制要求配置
    SRC = C.TEX
    if not SRC or not os.path.exists(SRC):
        raise SystemExit('未配置 tex 源文件（l2w.json 的 "tex" 字段或环境变量 L2W_TEX）。')
    raw = open(SRC, encoding="utf-8").read()
    t = raw

    # ---------- 0. 剥离 makeatletter 块与导言区自定义 ----------
    t = re.sub(r'\\makeatletter.*?\\makeatother', '', t, flags=re.S)

    # ---------- 1. 文献 key -> 编号 ----------
    bib_key2num = {}

    def repl_bib_env(m):
        return ("\\begin{thebibliography}{99}\n"
                + convert_bib_block(m.group(1), bib_key2num)
                + "\\end{thebibliography}")

    t = re.sub(r'\\begin\{thebibliography\}\{\d+\}(.*?)\\end\{thebibliography\}',
               repl_bib_env, t, flags=re.S)
    print("文献条数:", len(bib_key2num))

    # ---------- 2. label -> (类型, 编号) ----------
    label2ref = {}
    tab_n = eq_n = fig_n = 0
    for m in re.finditer(r'\\begin\{(table|figure|equation)\}(.*?)\\label\{([^}]+)\}', t, flags=re.S):
        env, label = m.group(1), m.group(3)
        if env == 'table':
            tab_n += 1; label2ref[label] = ('表', tab_n)
        elif env == 'equation':
            eq_n += 1; label2ref[label] = ('式', eq_n)
        else:
            fig_n += 1; label2ref[label] = ('图', fig_n)
    print("label映射:", label2ref)

    # ---------- 3. \cite / \ref ----------
    def repl_cite(m):
        nums = [str(bib_key2num[k.strip()]) for k in m.group(1).split(',') if k.strip() in bib_key2num]
        return '⟦CITE:' + ','.join(nums) + '⟧' if nums else ''
    t = re.sub(r'\\cite\{([^}]+)\}', repl_cite, t)

    def repl_ref(m):
        lb = m.group(1)
        if lb in label2ref:
            typ, num = label2ref[lb]
            return f'⟦REF:{typ}|{num}⟧'      # 类型与编号分开，后处理据前文决定是否补「表/式/图」
        return f'⟦REF:|{lb}⟧'
    t = re.sub(r'\\ref\{([^}]+)\}', repl_ref, t)

    # ---------- 4. \tabnote 展开为普通段落 ----------
    def repl_tabnote(m):
        return '\n\n⟦TABNOTE⟧' + m.group(1) + '\n'
    t = re.sub(r'\\tabnote\{((?:[^{}]|\{[^{}]*\})*)\}', repl_tabnote, t)

    # ---------- 5. \footnote 转占位（Pandoc 脚注在表格附近易错位，改文末注） ----------
    foot_notes = []
    def repl_foot(m):
        inner = m.group(1)
        inner = re.sub(r'\\url\{([^}]*)\}', r'\1', inner)
        foot_notes.append(inner)
        return f'⟦FOOTNOTE:{len(foot_notes)}⟧'
    t = re.sub(r'\\footnote\{((?:[^{}]|\{[^{}]*\})*)\}', repl_foot, t)
    print("脚注数:", len(foot_notes))

    # ---------- 6. ctex 命令清理 ----------
    t = re.sub(r'\\documentclass\[[^\]]*\]\{ctexart\}', r'\\documentclass[a4paper]{article}', t)
    t = re.sub(r'\\ctexset\s*\{(?:[^{}]|\{[^{}]*\})*\}', '', t, flags=re.S)
    t = re.sub(r'\\usepackage\[[^\]]*\]\{[^}]*\}', '', t)
    t = re.sub(r'\\usepackage\{[^}]*\}', '', t)
    t = re.sub(r'\\usetikzlibrary\{[^}]*\}', '', t)
    t = re.sub(r'\\urlstyle\{[^}]*\}', '', t)
    t = re.sub(r'\\zihao\{-5\}', r'\\small ', t)
    t = re.sub(r'\\zihao\{3\}', r'\\LARGE ', t)
    t = re.sub(r'\\zihao\{4\}', r'\\Large ', t)
    t = re.sub(r'\\zihao\{5\}', r'\\large ', t)
    t = re.sub(r'\\zihao\{[^{}]*\}', '', t)
    t = re.sub(r'\\heiti|\\kaishu|\\songti|\\bfseries', '', t)
    t = re.sub(r'\\onehalfspacing|\\singlespacing|\\raggedright', '', t)
    t = re.sub(r'\\setlength\{[^}]*\}\{[^}]*\}', '', t)
    t = re.sub(r'\\captionsetup\{[^}]*\}', '', t)
    t = re.sub(r'\\setcounter\{[^}]*\}\{[^}]*\}', '', t)
    t = re.sub(r'\\pagestyle\{[^}]*\}', '', t)
    t = re.sub(r'\\renewcommand\{\\refname\}\{\}', '', t)
    t = re.sub(r'\\renewcommand\{[^}]*\}\{[^}]*\}', '', t)
    # thebibliography 宽度参数 {44} 会被 Pandoc 当正文输出（"99"/"44"杂散段），置空
    t = re.sub(r'\\begin\{thebibliography\}\{\d+\}', r'\\begin{thebibliography}{}', t)
    t = re.sub(r'\\newcommand\{\\[a-zA-Z]+\}\{[^\n]*\}', '', t)
    t = re.sub(r'\\hrule[^\n]*', '', t)
    t = re.sub(r'\\vspace\*?\{[^}]*\}', '', t)
    t = re.sub(r'\\hspace\*?\{[^}]*\}', ' ', t)
    t = re.sub(r'\\smallskip|\\medskip|\\bigskip|\\par\b', '', t)
    t = re.sub(r'\\noindent', '', t)
    t = re.sub(r'\\quad', ' ', t)
    t = re.sub(r'\\qquad', '  ', t)
    t = re.sub(r'\\newpage|\\clearpage', '', t)
    t = re.sub(r'\\tbfs', '', t)
    t = re.sub(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}',
               r'\\includegraphics[width=0.9\\textwidth]{fig1.png}', t, flags=re.S)
    # 中文千分位细空格 \, -> 空
    t = t.replace('\\,', '')
    # 标题里的 \\[6pt] 换行保留为 Pandoc 硬换行
    t = re.sub(r'\\\\\[\d+pt\]', r'\\\\', t)

    open(OUT, "w", encoding="utf-8").write(t)
    print("pre.tex:", len(t), "chars")
    print("CITE:", t.count('⟦CITE:'), "REF:", t.count('⟦REF:'),
          "BIB:", t.count('⟦BIB:'), "TABNOTE:", t.count('⟦TABNOTE⟧'))
    keys = set()
    for m in re.finditer(r'\\cite\{([^}]+)\}', raw):
        keys.update(k.strip() for k in m.group(1).split(','))
    print("未映射cite key:", keys - set(bib_key2num))
    # 脚注文本落盘供后处理
    import json
    json.dump(foot_notes, open(C.tmp('footnotes.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
