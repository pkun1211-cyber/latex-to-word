---
name: latex-to-word
description: 把中文 LaTeX 论文（ctex/ctexart）高保真转换为 Word docx，并逐行对齐 LaTeX 编译 PDF 的版面。当用户要求「latex转word」「tex转docx」「Word版要和PDF一模一样」「保留LaTeX排版的Word」「论文转Word投稿」「按PDF逐行对齐Word」时使用。覆盖 Pandoc 转换链、三线表/上标引文/中文编号重建、LM Roman 字体安装、逐行硬对齐与字距闭环校正、逐页行数验收。仅适用于 Windows + 本机 Word（需 COM 导出 PDF 做闭环）。
---

# LaTeX → Word 高保真转换

## 这个 skill 解决什么

Pandoc 直出的 docx 与 LaTeX PDF 差异巨大：字体不同、行距不同、表格不是三线表、
引文变成裸文本、中文标题没编号、断行位置全不一样。本 skill 提供一条**七步链**，
把差异压到「页数相同、逐页行数接近、断行位置逐行锁定」。

**核心认知：Word 与 TeX 的断行永远不可能自动一致。**
TeX 的 `\spaceskip` 允许词间空格**双向伸缩**（实测 −0.74 ~ +1.81pt/空格），
Word 两端对齐**只拉伸不压缩**。所以必须把 LaTeX 的断行结果**直接写进 Word**
（插 `w:br` + 逐行字距补偿），而不是指望调参数让它们自然收敛。

## 何时用 / 何时不用

**用**：中文学术论文（ctexart/article）、有 LaTeX 编译好的 PDF 作为基准、
目标是投稿或归档的 Word 版、要求版面高度一致。

**不用**：
- 只要内容不要版面 → 直接 `pandoc x.tex -o x.docx` 就够
- 需要 Word 里继续自由编辑 → 硬对齐后段落是硬排版，改文字必须重新生成（见下方权衡）
- macOS / Linux → 闭环校正依赖 Word COM 导出 PDF，本 skill 的验收环节不可用
  （转换链本身可跑，但无法逐行校准）

## 保真强度：先和用户确认

| 档位 | 做法 | 代价 |
|---|---|---|
| A 样式对齐 | 只对齐字体/字号/行距/页边距/三线表 | 断行与分页仍不同 |
| B 分页对齐 | A + 浮动体位置 + 段间距 + 表格宽度 | 页数一致，行内容仍有差异 |
| **C 逐行硬对齐** | B + 按 LaTeX 断点插 `w:br` + 逐行字距 | **段落变硬排版，改文字需重新生成** |

默认走 C（用户说「一模一样」时）。**必须先告知 C 的代价并取得确认**。

## 环境前提

```
Python: C:\Users\<user>\.workbuddy\binaries\python\envs\default\Scripts\python.exe
        （PATH 上的 python 是基础解释器，没装 python-docx/pymupdf，会静默降级）
依赖:   python-docx  lxml  pymupdf  fontTools  pywin32
外部:   pandoc（转换）、本机 Microsoft Word（COM 导出 PDF）
        pandoc 不必在 PATH：run_chain.py 会自动回退到 pypandoc 内置版本
        （只装了 pypandoc-binary 的机器也能跑，见第 2 步）
字体:   本机 assets/ttf_lm/ 已含 13 个 LM Roman TTF，可直接 install_ttf.py；
        新克隆的仓库不含字体二进制，用 fetch_fonts.py 自动获取
        （本机 MiKTeX 优先，否则从 CTAN 镜像下载）→ 转 TTF → 修 name 表 → 安装
```

## 配置：所有脚本统一读 l2wconf

在项目 tmp 目录放 `l2w.json`，用 `L2W_CONF` 指向它：

```json
{
  "base":   "D:/proj",
  "docx":   "D:/proj/out/paper.docx",
  "tex":    "D:/proj/paper.tex",
  "texpdf": "D:/proj/paper.pdf",
  "tmp":    "D:/proj/.workbuddy/tmp",
  "tex_tag": "tex",
  "margins_cm": {"top": 2.4, "bottom": 2.4, "left": 2.5, "right": 2.5}
}
```

```bash
export L2W_CONF="D:/proj/.workbuddy/tmp/l2w.json"
```

`tex_tag` 是 LaTeX 行提取 JSON 的标签，后续 `hard_align` 读 `<tex_tag>.json`。
**所有写盘都走 `docxio.py`**（自动关 Word 占用 + 原子替换 + 首次备份）。

---

## 快速通道：统一入口

```bash
export L2W_CONF="D:/proj/.workbuddy/tmp/l2w.json"
$PY $S/run_chain.py all --level C --dry-run   # 先看会跑什么
$PY $S/run_chain.py base pre post             # 分阶段跑
$PY $S/run_chain.py hard                      # 逐行硬对齐
$PY $S/run_chain.py iter --rounds 3           # 闭环校正 3 轮
$PY $S/run_chain.py check                     # 只验收
```

阶段：`font` / `base` / `pre` / `post` / `hard` / `iter` / `ref` / `check`。
`all --level B` 只到分页对齐，`--level C` 含逐行硬对齐。
配置样例见 `assets/l2w.example.json`。

下面是各阶段的手工命令与原理，排错时按此逐步定位。

## 七步链

设 `S=<skill>/scripts`，`PY=<托管python>`，`T=<tmp目录>`。

### 第 0 步：装字体（一次性）

**情况 A：assets/ttf_lm/ 已有 TTF**（本机技能目录就是这种）

```bash
$PY $S/install_ttf.py          # 用户级安装，写 HKCU，不动系统目录
$PY $S/audit_lmfont.py         # 校验 Word 能否识别
```

**情况 B：没有字体二进制**（从 GitHub 新克隆——仓库不含 TTF，只有许可与下载脚本）

```bash
$PY $S/fetch_fonts.py --install   # 获取 + 转 TTF + 修 name + 用户级安装（写 HKCU）
$PY $S/audit_lmfont.py            # 校验 Word 能否识别
```

`fetch_fonts.py` 优先用本机 MiKTeX/TeX Live 的 OTF，没有则从 CTAN 镜像下载
（中科大 → 清华 → 官方，自动降级）。也可 `--from <目录>` / `--mirror <url>`。
分步执行等价于：`otf2ttf.py` → `fix_ttf_names.py` → `install_ttf.py`。

**OTF 必须转 TTF**：Word 对 CFF 轮廓支持不完整，直接用 OTF 会让导出 PDF 里
西文字符整体消失。转换须把 upm 提到 2048 且保留 hmtx 度量。

**装完必须重启 Word。** 卸载：`uninstall_ttf.py`（按安装清单精确回滚）。

为什么非装不可：Latin Modern 是**光学尺寸族**，5/6/7/8/9/10/12/17pt 字形宽度各不同。
用 Times New Roman 代替会让西文行宽系统性偏差，逐行对齐永远收不敛。

### 第 1 步：提取 LaTeX 版面基准

```bash
$PY $S/pdflines2.py "<paper.pdf>" tex
```

产出 `T/tex.json` + `T/tex.txt`：每行的基线 y、字号、左右界 x0/x1、字体、文本。
**这是后续一切对齐的唯一基准**，`l2w.json` 的 `tex_tag` 要与此处标签一致。

同时人工记录：版心宽、正文基线距（→ `w:spacing line` twips）、
各级标题字号、首行缩进量、摘要悬挂缩进量。

### 第 2 步：预处理 tex

```bash
$PY $S/preprocess1120.py       # 读 C.TEX，写 T/pre.tex + T/footnotes.json
```

把 Pandoc 会弄丢的东西转成**占位标记**（Pandoc 原样保留，后处理再重建）：

| 原 LaTeX | 占位 | 为什么 |
|---|---|---|
| `\cite{k1,k2}` | `⟦CITE:1,2⟧` | Pandoc 无 bib 时直接丢弃 cite |
| `\ref{lb}` | `⟦REF:表\|3⟧` | 类型与编号分开，后处理据前文决定是否补「表/式/图」 |
| `\bibitem{k}` | `⟦BIB:n⟧` | 保留编号，后处理拆成独立段 |
| `\tabnote{}` | `⟦TABNOTE⟧` | 自定义命令 Pandoc 不认 |
| `\footnote{}` | `⟦FOOTNOTE:n⟧` | Pandoc 脚注在表格附近会错位 |

同时剥离 ctex 专属命令（`\ctexset`/`\zihao`/`\heiti` 等）、
把 tikzpicture 换成 `\includegraphics`、清空 `thebibliography` 的宽度参数
（否则 `{99}` 会变成正文里的杂散段）。

### 第 3 步：建模板 + Pandoc 转换

```bash
pandoc --print-default-data-file reference.docx > $T/reference_base.docx
$PY $S/build_reference.py      # 写 T/reference_custom.docx
pandoc $T/pre.tex -o $T/pandoc_raw.docx \
       --reference-doc=$T/reference_custom.docx --from=latex --to=docx
```

`build_reference.py` 按第 1 步记录的参数钉死 Normal/Heading/Title/Caption 等样式。
**必须基于 Pandoc 官方 reference.docx 改**，不能自己新建空 docx——
Pandoc 依赖里面一整套样式 ID（`Compact`/`FirstParagraph`/`TableCaption`…）。

**pandoc 找不到怎么办**：`run_chain.py pre` 内置探测顺序为
PATH → pypandoc 内置 → 常见安装目录，所以只装 `pypandoc-binary` 也能跑：

```bash
$PY -m pip install pypandoc-binary   # 自带 pandoc，不必改 PATH
```

手工敲上面两条 pandoc 命令时若报「不是内部或外部命令」，
把 `pandoc` 换成 `$PY -c "import pypandoc;print(pypandoc.get_pandoc_path())"` 输出的路径。

### 第 4 步：OOXML 后处理

```bash
$PY $S/postprocess1120.py      # 读 T/pandoc_raw.docx，写 C.DOCX
```

十件事：文献拆段 → CITE 转上标 → REF 转中文交叉引用 → 表注归位 →
真实脚注 → **三线表** → 标题中文编号（一、/（一））→ 图表题编号居中 →
正文缩进 2em + exact 行距 → 页边距。

尾部自动串跑精修链（顺序不可换）：
`refine_layout` → `fix_settings` → `fix_justify` → `fix_fontsize` → `fix_lig2` → `fix_ns`

**项目专属的首页拆行补丁要在这一步之前跑**，因为它会改变段落索引。

### 第 5 步：逐行硬对齐（档位 C 才做）

```bash
$PY $S/hard_align.py                    # 只看配对计划，不改文件
$PY $S/hard_apply.py                    # 看写入计划
$PY $S/hard_apply.py --apply            # 插 w:br + jc=left
$PY $S/hard_spacing.py                  # 看字距补偿计划
$PY $S/hard_spacing.py --apply          # 写 w:spacing
$PY $S/hard_incr.py --apply             # 补配对新命中的段（改过判据后用）
```

配对逻辑在 `hard_align.plan()`：**单调序列对齐**（Word 段序与 LaTeX 段序同向，
只在上一命中之后向后找），前缀一致性以 **LaTeX 首行长度**为分母。
含公式/图片的段、表格单元格、图表题一律跳过。

### 第 6 步：闭环校正（关键，不可省）

静态度量与 Word 实际渲染有系统偏差，必须「导出 PDF → 实测 → 反馈」：

```bash
$PY $S/export_cmp.py "<C.DOCX>" it1     # 导出 T/it1.pdf + 逐页对比
$PY $S/hard_iter.py it1                 # 看修正计划
$PY $S/hard_iter.py it1 --apply         # 只修溢出行
# 重复 2~3 轮，每轮换标签 it2 / it3
```

**`hard_iter` 只修「行数溢出」的行，绝不做全量更新**——全量必然发散（实测 44→69）。

文献区单独精校（模型更准，见 references 第 28~31 条）：

```bash
$PY $S/fix_ref3.py                      # 干跑
$PY $S/fix_ref3.py --apply              # 静态度量模型 + floor 求解
$PY $S/fix_ref4.py <tag> --nos 18,33    # 顽固行：PDF 实测反解兜底
```

### 第 7 步：局部版面修补（按需）

```bash
$PY $S/fix_tblwidth.py --list           # 表格宽度对齐版心（消除单元格内折行）
$PY $S/fix_tblwidth.py --apply
$PY $S/fix_hang.py --apply              # 摘要/关键词悬挂缩进
$PY $S/fix_refspace.py --apply --line 270 --after 80    # 文献区行距段距
$PY $S/fix_eqspace.py --apply --line 270 --before 100 --after 100  # 公式节奏
$PY $S/fix_footnote.py --apply --sz 15 --line 180       # 脚注字号行距
$PY $S/move_block.py --list             # 浮动体整块移动
$PY $S/move_block.py 表10 --before-text "六、协同机制构建"
```

每改一项都要重跑 `export_cmp.py` 看指标，**变差立刻回滚**（`T/bak_before_*.docx`）。

---

## 验收清单

```bash
$PY $S/export_cmp.py "<C.DOCX>" final
$PY $S/cmp_pages.py tex final           # 分歧块定位 + 字号分布
```

必须报告给用户的指标：

1. **页数** LaTeX vs Word —— 必须相等
2. **逐页行数差** 与绝对偏差合计
3. **全文行相似度**（difflib 行序列）
4. **中文满行右界档位分布** —— 反映断行是否稳定
5. 每页首/末行匹配度 —— 定位内容错位

参考量级（一篇 12 页中文论文，走完 C 档）：
页数 12=12，行偏差 35，行相似度 0.727。
**不要承诺 100% 一致**——见下方硬边界。

## 已知硬边界（必须如实告知）

- **`exact` 行距 twips 取整抖动**：±0.1pt × 500 行 ≈ 4 行，无法消除
- **含公式的段落无法逐行锁定**：OMML 的 `m:oMath` 位置不能用 `w:br` 精确控制
- **图内文字/表注混入 LaTeX 行提取**：少数段落永久无法配对
- **公式区与浮动体高度**差异是独立问题，字距补偿解决不了
- **硬对齐后段落是硬排版**：改文字必须从 tex 重新走全链

## 排错

先查 `references/pitfalls.md`（36 条实测坑）。高频：

| 症状 | 原因 | 查 |
|---|---|---|
| 改了 spacing 但导出没变 | `apply_seg` 遇 tw==0 提前 return | 坑 27 |
| 字距越改越糟 | 全量更新 / 漏算 CJK 边界间距 | 坑 20、28 |
| 中文行右界超出版心十几 pt | 中文标点悬挂 | 坑 28 |
| Word 打不开 docx | pPr 子元素顺序非法 | 坑 5 |
| 标题变蓝 | themeColor 未清 | 坑 3 |
| 表格列宽等分 | Pandoc 不保留列宽比例 | 坑 8 |
| 脚注跑到文末 | footnotes part 复用 | 坑 12 |
| 导出报 WinError 5 | Word 占用文件 | 用 docxio |
| **文件莫名变小、指标退化** | **脚本顶层执行，import 即写盘** | **坑 32** |
| 加守卫后 NameError | 顶层函数引用了模块级实例 | 坑 32 |
| 换目录后大量脚本"未配置"就退出 | 顶层 `require_docx()` | 坑 34 |
| 技能目录里多出 `.workbuddy/tmp` | 配置模块顶层 `makedirs` | 坑 35 |

## 目录

```
latex-to-word/
├── SKILL.md
├── scripts/
│   ├── run_chain.py       统一入口（阶段编排）
│   ├── l2wconf.py         路径配置（所有脚本都读它）
│   ├── docxio.py          安全写盘（关 Word 占用 + 原子替换 + 备份）
│   ├── _smoke.py          自检：import 全脚本 + SHA256 校验文件未被改动
│   └── ...                34 个功能脚本
├── references/
│   └── pitfalls.md        36 条实测坑与量化数据
└── assets/
    ├── ttf_lm/            13 个 LM Roman TTF（2.8MB）
    └── l2w.example.json   配置样例
```

## 自检（改动脚本后必跑）

`python _smoke.py` 要在**两种环境各跑一次**，缺一不可：

```bash
# ① 技能目录内，无任何配置 —— 验证可加载性
cd <skill>/scripts && python _smoke.py        # 期望 34/34

# ② 指向真实项目 —— 验证功能且零副作用
L2W_CONF=/path/to/l2w.json python _smoke.py   # 期望 34/34 + SHA256 一致
```

只跑 ② 会漏掉「脱离项目目录就不可用」的缺陷（坑 34）；
只看"导入成功"会漏掉静默写盘（坑 32）。跑完还要 `ls -a` 确认
技能目录没多出 `.workbuddy/`（坑 35）。

**新增脚本的硬规范**：模块级只允许 import、常量、纯函数定义。
不得实例化 `Document`、不得持有可变全局状态、不得建目录、
不得在顶层调用 `C.require_docx()`（放 main 内，见坑 34）。
