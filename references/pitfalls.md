# LaTeX → Word 实测坑清单（33 条）

全部来自一篇 12 页中文 ctexart 论文的真实转换过程，含量化数据。
按「转换期 → 版面期 → 硬对齐期 → 度量期」分组。

---

## 一、Pandoc 转换期（1~12）

### 1. Pandoc 无 bib 文件时直接丢弃 `\cite`
`\cite{porter1985}` 在输出里**完全消失**，不报错。
→ 预处理阶段把 cite 转成占位 `⟦CITE:1,2⟧`，后处理再重建为上标 run。
编号映射从 `\bibitem` 顺序建立，不要依赖 bibtex。

### 2. `hyperlink` 不能嵌套
上标引文若做成 `<w:hyperlink><w:hyperlink>` Word 会修复文档并丢内容。
→ 引文上标只用普通 run + `w:vertAlign=superscript`，内部锚点单层。

### 3. 标题变蓝：`themeColor` 未清
Pandoc 官方 reference.docx 的 Heading 样式带
`<w:color w:val="345A8A" w:themeColor="accent1"/>`。
只改 `w:val` 不够，**必须删掉 `w:themeColor`/`w:themeShade`/`w:themeTint`**，
否则 Word 优先用主题色。

### 4. `thebibliography` 的宽度参数变成正文
`\begin{thebibliography}{99}` 的 `{99}` 被 Pandoc 当普通文本输出，
文档里出现孤立的「99」段。
→ 预处理置空为 `\begin{thebibliography}{}`；后处理再兜底删除
内容为 `99`/`44` 且样式为 FirstParagraph/BodyText 的段。

### 5. `w:pPr` 子元素顺序非法 → Word 静默丢属性
OOXML schema 对 pPr 子元素有**严格顺序**。顺序错了 Word 不报错，
直接忽略部分属性（表现为行距/缩进设了没生效）。
正确顺序见 `hard_align.PPR_ORDER`，rPr 见 `RPR_ORDER`。
→ 每次改完 pPr/rPr 都调 `sort_el()`。`fix_settings.py` 做全文兜底排序。

### 6. `\zihao` / `\ctexset` 等 ctex 命令 Pandoc 不认
→ 预处理全部剥离，字号信息改由 reference.docx 的样式承载。
`\zihao{-5}`→`\small`、`{3}`→`\LARGE`、`{4}`→`\Large`、`{5}`→`\large`。

### 7. tikzpicture 整块丢失
→ 预处理替换为 `\includegraphics{fig1.png}`，图另行导出 PNG。
**注意**：图内 TikZ 文字会混进 LaTeX PDF 行提取，导致该段永久无法配对（见坑 25）。

### 8. Pandoc 不保留表格列宽比例
所有列被等分。实测表 6 列宽合计 7920 twips = 396pt，
而版心宽 453.5pt = 9070 twips → 「RKS 治理评分」被折成两行，
表头折成 3 行，整表比 LaTeX 高 36pt，把后续内容挤到下一页。
→ `fix_tblwidth.py` 按版心宽等比放大所有列。

### 9. 三线表必须手工重建
LaTeX 的 booktabs `\toprule/\midrule/\bottomrule` 转不过来。
→ 后处理清空所有单元格边框，只给首行上/下、末行下加 `w:tcBorders`。
线宽按 LaTeX 实测：top/bottom 0.75pt，mid 0.5pt。

### 10. `\ref` 的中文前缀要看上下文
LaTeX 里写 `表\ref{tab:x}`，转出来是「表⟦REF:表|3⟧」；
但也有 `\ref{tab:x}` 单独用（前文已有「表」字）。
→ 占位保留**类型与编号分离**（`⟦REF:表|3⟧`），后处理检查前一个字符
是否已是「表/式/图」，避免「表表3」「式式1」。

### 11. 自定义命令 `\tabnote` 被吞
→ 转成 `⟦TABNOTE⟧` 前缀的独立段，后处理按标记把它移到对应表格之后，
并设小字号 + 无缩进。

### 12. Pandoc 脚注在表格附近错位 / footnotes part 复用
Pandoc 生成的脚注可能落到文末，或表格内的脚注引用错乱。
→ 预处理把 `\footnote` 转占位 + 文本落盘 JSON，
后处理**自己写 `word/footnotes.xml`** 并插 `w:footnoteReference`。
注意 docx 里 footnotes part 可能被多处复用，改一处会影响全部。

---

## 二、版面期（13~19）

### 13. `docGrid` / `snapToGrid` 改写 exact 行距
中文 Word 模板默认有文档网格，会把 `lineRule="exact"` 的行距吸附到网格。
→ 段落级 `<w:snapToGrid w:val="0"/>`，`fix_settings.py` 全文设置。

### 14. `characterSpacingControl` 两难
`compressPunctuation` 能复刻 ctex 的中文标点挤压，
但**会连西文字距一起压**，英文 Abstract 会少 2 行。
`compressPunctuationAndJapaneseKana` 同样影响西文。
→ 保持 `doNotCompress`，中文断行差异改由逐行 `w:spacing` 补偿解决。

### 15. `autoSpaceDE` / `autoSpaceDN` 必须关
Word 默认在西文与中文/数字之间自动加间距。
→ 段落级 `<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>`。
**但注意：这压不掉「西文空格→汉字」边界的间距**（见坑 28）。

### 16. 行距要用 exact twips，不能用倍数
LaTeX 实测正文基线距 13.5pt → `<w:spacing w:line="270" w:lineRule="exact"/>`。
用 `line_spacing=1.5` 之类的倍数会随字号变化，无法精确对齐。
**代价**：twips 是整数，±0.1pt 抖动累积（见坑 26）。

### 17. 摘要/关键词是悬挂缩进，不是首行缩进
LaTeX 实测：［摘要］首行 x0=66.05，续行 x0=70.87 → 首行**左移** 4.82pt。
正文段则是首行 x0=88.87（缩进 2em=18pt），续行 70.87。
Word 若给摘要段 firstLine=0，首行少 4.82pt 可用宽 → br 前末字被挤到下一行。
→ `fix_hang.py` 设 `w:ind left=0 firstLine=-96`（−4.8pt）。

### 18. 脚注字号行距差异会挤走正文
LaTeX `\footnotesize` 在 zihao=5 下约 7.5pt，行距 9pt。
Word 默认 FootnoteText 是 8pt / 13.6pt → 脚注区高 14pt，
把正文末 2 行挤到下一页（实测 p7 少 2 行、p8 多 2 行）。
→ `fix_footnote.py --sz 15 --line 180`（7.5pt / 9pt）。

### 19. 公式段垂直节奏不一致
LaTeX amsmath equation 间隔实测 34.4pt，正文→首公式约 20pt。
Word 里不同公式段的 before/after 五花八门（33.5pt / 35pt 混杂）。
→ `fix_eqspace.py` 统一 line/before/after。
**但公式区总高度仍可能差 9 行**——OMML 与 TeX 的公式排版高度本质不同。

---

## 三、硬对齐期（20~27）

### 20. 字距闭环迭代**只能修溢出行**，全量更新必然发散
判据：`overflow = Word该行实际行数 > 1`。
实测全量更新：行偏差 44→69、42→49，越改越糟。
→ `hard_iter.py` 里 `if not overflow: continue`。

### 21. `w:br` 只锁断点，不防二次断行
插了 br 之后，若该行**自然宽 > 可用宽**，Word 仍会在 br 之前断行，
末词被挤到下一行，产生「孤字行」。
实测文献区表现：`and` / `Organizational` / `than` / `123.` / `Pro` /
`Economics` / `Infor` / `of` / `Inter` / `Supply` / `the` 各自单独成行。
→ 必须配合字距压缩，把自然宽压到可用宽以内。

### 22. 求解 spacing 用 `floor` 不能用 `round`
溢出行哪怕只超 0.1pt 也会断行。`round` 会留下残余溢出。
实测：7 行用 round 修正后**全部没修好**。
→ `new = math.floor(need)`。

### 23. 配对必须单调序列对齐，不能贪心全局最佳
「全局最佳 + used 集合」会让靠前的 Word 段抢走靠后的 LaTeX 段，
后续指针错位（实测 P22/P24 只配到 1 行）。
→ 只在上一命中之后向后找（窗口 40 段）。

### 24. 前缀一致性的分母用「LaTeX 首行长度」，不是整段长度
文献条目常首行完全一致、次行才分歧。用整段长度做分母会把
相似度压到 0.53~0.86，导致 [9][17][19][24] 漏配对。
→ `denom = min(k, max(20, len(mnorm(TP[j][0]['t']))))`。
改这一处后配对数 92→98。

### 25. 四类段落永久无法配对，必须跳过
含 `m:oMath` 的段、含 `w:drawing` 的段、表格单元格、图表题。
另有少数段因**图内 TikZ 文字 / 表注混入 LaTeX 行提取**而文本对不上
（本论文 P9/P75/P93/P100）。
→ 不要试图强配，会污染后续指针。

### 26. `exact` 行距 twips 取整抖动无法消除
±0.1pt × 500 行 ≈ 4 行偏差。这是格式本身的量化误差，不是 bug。
→ 如实告知用户，不要承诺 100% 一致。

### 27. ★ `apply_seg` 遇 `tw == 0` 提前 return → 旧值残留
```python
tw = int(round(delta * 20))
if tw == 0: return 0      # ← BUG
```
若该 run 已有非零 `w:spacing`，目标为 0 时**必须显式写 0 覆盖**。
症状：脚本报告「已改 spacing 1→0」，导出 PDF 后行宽纹丝不动。
本论文 [18][35] 两条文献因此连续三轮修不掉。
→ 删掉这个提前 return。

---

## 四、度量期（28~31）★ 最关键

### 28. ★ Word 在「西文空格 → CJK 汉字」边界额外插 ≈1.75pt
`autoSpaceDE=0` / `autoSpaceDN=0` **也压不掉**。
实测 gap：[1] 1.747/1.713/1.750；[15] 1.710/1.913/1.917；
[16] 1.734/1.713/1.765；[42] 1.957/1.713/1.935。

中文文献条目只有 2~3 个空格，但**全是「空格→汉字」**，
所以静态度量低估 4~7pt。第一版脚本因此把本该压缩的行判成需要拉伸，
越改越糟（42→49）。

**例外**：[43] 两处边界 gap 为**负**（−0.054/−0.097），
说明 Word 在特定条件下会压掉该间距。
→ **实测优先、估算兜底**，只计 gap > 0.8pt 的边界，不要一律加常数。

### 29. ★ spacing 生效系数 = 1.0，常数偏置 = +0.10pt
**曾经的错误结论**：拟合出「系数 0.968 + 常数 0.90pt」。
那是把 CJK 边界间距**重复计入**造成的假象。

**正确验证方法**：看字符间隙的**中位数**。
cur=5 → gap 中位 0.252；cur=1 → gap 中位 0.045。
恰好等于 `cur/20`，故系数就是 1.0。

最终模型：
```
真实行宽 = nat + Σ(空格→CJK边界间距) + cur × n / 20 + 0.10
nat = 静态度量（LM Roman hmtx advance 之和，含引文上标）
n   = 可施加字距的字符数（不含引文上标）
```
反解：`cur' = (可用宽 − nat − Σ边界 − 0.10) / n × 20`，取 floor。

### 30. ★ 溢出行的自然宽**必须用静态度量**，不能用 PDF 字符宽和
溢出行的 PDF **只渲染了前半部分**（末词被推到下一行）。
用 PDF advance 求和会把「少了 3 个词」误算成「需要大幅拉伸」。
实测：[6] PDF 81 字符 vs 段 95 字符（少了 `on Organizational`），
用 PDF 宽算出 spacing 5→17，完全错误。

→ 自然宽一律用 `HS.seg_width(seg)`（覆盖整段字符）。
PDF 只用来**实测边界间隙**，或在兜底脚本里用
`ink_已渲染 + nat×剩余字符占比` 反解。

### 31. PDF 把变音符/连字拆成独立 span，字符流对不上
`MUÑOZ` → `MUÑ OZ`，`KÖLBEL` → `KÖ LBEL`，`RIGOBÓN` → `RIGOBÓ N`。
`Definition` 里的 `fi` 是连字 `ﬁ`（U+FB01）。

若按前缀匹配，[35] 只能匹配到 18/99 字符，
剩余占比过大 → spacing 算爆（−69）。
→ 判据：`if not txt.startswith(pt) or len(pt) < len(txt)*0.7:` 
则**放弃实测**，退回静态度量 + 强制额外压缩 2 twip。

另：`pymupdf` 的 `rawdict` 里 span **没有 `'text'` 键**，
要用 `''.join(c['c'] for c in s['chars'])`。

---

## 五、工程期（32~33）★ 会损坏用户文件

### 32. ★★ 脚本主逻辑写在模块顶层 → `import` 即写盘，静默损坏成品
`refine_layout` / `fix_fontsize` / `fix_justify` / `fix_settings` /
`build_reference` / `preprocess` / `postprocess` 原本把主流程写在顶层。
任何 `import`（包括自检、被别的脚本引用）都会执行并写盘。

**实测事故**：一次 `_smoke.py` 逐个 import，把 326489 字节的成品改成
318668 字节，行偏差从 35 退到 56，且没有任何报错提示。

→ 全部包进 `def main()` + `if __name__ == '__main__': main()`。
**切点判据**：跳过 import / 函数定义 / 类定义 / 文档字符串 /
不含实例化的纯赋值 / 已有的 `__main__` 守卫块；
遇到第一条真正执行语句（`Document(...)`、`open(...)`、for/with/try 执行体）才切。

**连带问题**：顶层函数若引用了模块级实例（如 `def sty(sid): for s in d.styles`），
移走 `d` 后会 `NameError`。必须改成显式传参 `def sty(doc, sid)`。
同理 `bib_key2num` 这类可变全局状态也要改成参数传入。

**配置读取也要延迟**：`SRC = C.TEX or C.require_docx()` 写在顶层会让
import 阶段就强制要求配置。改为在 main 内取值。

### 33. ★ 自检必须校验文件哈希，不能只看「导入成功」
「33/33 导入成功」不等于安全——坑 32 的事故里所有 import 都"成功"了。
→ `_smoke.py` 在前后各算一次目标 docx 的 SHA256，
不一致就报 `[!!] 目标 docx 被修改`，退出码非 0。

同时：**每轮改动前存一个明确命名的基准备份**（如 `GOOD_35_07236.docx`）。
自动备份名（`bak_before_*`）会被后续脚本覆盖，
`move_block.py --undo` 恢复的是它自己首次运行时的存档，
不一定是你想要的上一状态。

### 34. ★ 顶层 `require_docx()` 让 skill 脱离项目目录就不可用
封装完成时自检 33/33 通过，是**在项目目录下**跑的（那里有 `l2w.json`）。
换到技能目录再跑，只剩 **12/33**：21 个脚本在 import 阶段就 `SystemExit`。
根因是 7 个脚本顶层写了 `DOCX = C.require_docx()`，
其余 14 个是 import 它们的连带失败。

后果：新用户照 SKILL.md 第一步想看用法，直接被"未配置"挡住，
连 `--help` 都跑不到。**这是可用性缺陷，不是环境问题。**

→ 顶层改容错取值，校验推迟到执行入口：
```python
DOCX = C.DOCX          # 顶层：可能为空串，不校验

def main():
    global DOCX
    DOCX = C.require_docx()   # 执行时才校验
```
验收必须**两种环境各跑一次**：
- 技能目录内（无配置）→ 应 33/33 加载成功
- 项目目录 / `L2W_CONF` 指向配置 → 33/33 且 SHA256 不变
- 无配置下直接执行脚本 → 应打印可操作提示，而非 traceback

### 35. 配置模块顶层 `os.makedirs` 会污染技能包
`l2wconf.py` 顶层的 `os.makedirs(TMP, exist_ok=True)`，
使得任何 import 都在**当前工作目录**建出 `.workbuddy/tmp`。
在技能目录里跑自检，就把技能包写脏了。

→ 拆成两个函数，职责分离：
```python
def ensure_tmp():        # 显式创建，不在顶层调用
    os.makedirs(TMP, exist_ok=True); return TMP

def tmp(*parts):         # 纯路径拼接，零副作用
    return os.path.join(TMP, *parts)
```
目录创建下移到写盘环节（`docxio.backup()` 里 `os.makedirs(tmpdir)`）。

**通用判据**：配置模块的顶层只允许「读环境变量 / 算路径 / 定义函数」。
任何 `makedirs`、`open(...,'w')`、`Document(...)`、网络请求都必须延迟。
自检除了比 SHA256，还要 `ls -a` 看技能目录有没有多出文件。

---

## 附：LaTeX 行提取的两个判据

**上标引文行是多个并排**：`[17][18, 19]`、`[33] [34] [35, 36]`。
用 `fullmatch(r'\[[\d,\s]+\]')` 会漏掉。
→ `(?:\[\d+(?:[,，\-–]\s*\d+)*\]\s*)+$`

**参考文献新条目**：行首在版心左界（x0 < LEFT+6，悬挂缩进）
+ `[n]` 且 **n 递增**。
续行常以 `[J].` / `[M].` / `[2023]` 开头，只看 `^\[\d+\]` 会全部误判。

**TeX 断词连字符**：行末可能是 `TheAc-` / `GasPro-` / `Quan-` / `Infor-`，
Word 侧没有这个 `-`。匹配行尾时要容忍（`match_end` 的 HYPH 跳过逻辑）。
