# latex-to-word

> 把中文 LaTeX 论文高保真转换为 Word，并**逐行对齐 LaTeX 编译出的 PDF 版面**。

一条经过真实论文打磨的转换链：**38 个脚本 + 35 条实测坑清单 + 字体自动获取**。
不是「跑个 Pandoc 就完事」，而是把 Pandoc 直出稿与 LaTeX PDF 之间的每一处版面差异
逐项量化、逐项修正，直到页数相同、断行位置逐行锁定。

```
LaTeX PDF  ─┐
            ├─→ 逐行比对 ─→ 字距闭环校正 ─→ Word docx
Word docx  ─┘        （导出 PDF 实测，反馈修正）
```

---

## 目录

- [为什么需要它](#为什么需要它)
- [核心认知：断行永远不可能自动一致](#核心认知断行永远不可能自动一致)
- [效果](#效果)
- [三档保真强度](#三档保真强度)
- [环境要求](#环境要求)
- [快速开始](#快速开始)
- [七步链详解](#七步链详解)
- [关键技术原理](#关键技术原理)
- [脚本总览](#脚本总览)
- [已知硬边界](#已知硬边界)
- [工程纪律](#工程纪律)
- [作为 WorkBuddy Skill 使用](#作为-workbuddy-skill-使用)
- [许可与致谢](#许可与致谢)

---

## 为什么需要它

学术论文常遇到这种要求：**「用 LaTeX 写，但投稿/归档要交 Word，而且格式不能变。」**

直接 `pandoc paper.tex -o paper.docx` 的结果是：

| 项目 | LaTeX PDF | Pandoc 直出 Word |
|:---|:---|:---|
| 西文字体 | Latin Modern（光学尺寸族） | Times New Roman / Calibri |
| 表格 | 三线表 | 全网格普通表 |
| 引文 | 上标 `[1,2]` | 裸文本，甚至**直接丢失** |
| 中文标题 | 一、／（一）／1. | 无编号 |
| 交叉引用 | 表 3／式(2) | `??` 或裸数字 |
| 行距 | 固定基线距 | 单倍/多倍行距 |
| 断行位置 | TeX 断行算法 | Word 两端对齐 |
| 页数 | 12 页 | 常常 14~16 页 |

本项目把这些差异**逐项拆解并修正**，而不是指望换个模板就好。

---

## 核心认知：断行永远不可能自动一致

这是整个项目最重要的一条结论，也是所有设计的出发点：

> **TeX 的词间空格可以双向伸缩，Word 两端对齐只拉伸不压缩。**

实测 TeX 在一行内对空格施加的伸缩量是 **−0.74 ~ +1.81 pt/空格**。
也就是说 TeX 会**把空格压窄**来塞进一个词，而 Word 永远不会这么做——
它只会把词推到下一行。

**推论：无论怎么调字体、字号、行距、页边距，Word 都不可能自动复现 TeX 的断行。**

所以本项目采取的做法是：

1. 从 LaTeX PDF 提取**每一行的实际文本与坐标**（`pdflines2.py`）
2. 把 LaTeX 的断行结果**直接写进 Word**——在每个断点插入 `w:br`
3. 对每一行**单独计算字距补偿**（`w:spacing`），让它刚好填满版心宽度
4. 导出 Word 为 PDF，**实测行宽再反馈修正**（闭环，2~3 轮）

这就是「逐行硬对齐」。代价见[三档保真强度](#三档保真强度)。

---

## 效果

一篇 12 页中文社科论文（含 8 个三线表、多个公式、44 条参考文献、TikZ 插图）：

| 指标 | 结果 |
|:---|:---|
| 页数 | **12 = 12**（完全一致） |
| 逐页行数绝对偏差合计 | **35 行** |
| 全文行相似度（difflib） | **0.724** |
| 参考文献区断行 | 基本逐行一致 |

优化过程中的真实收敛轨迹：

```
初始 Pandoc 直出   →  页数不等，断行全不同
样式对齐（A 档）   →  字体/行距/三线表就位
分页对齐（B 档）   →  页数 12=12
逐行硬对齐（C 档） →  行偏差 77 → 44 → 41 → 35
```

> **不承诺 100% 一致。** 剩余偏差有明确的物理原因，见[已知硬边界](#已知硬边界)。

---

## 三档保真强度

**使用前必须先确定档位**，因为 C 档有明确代价。

| 档位 | 做法 | 代价 |
|:---|:---|:---|
| **A 样式对齐** | 字体 / 字号 / 行距 / 页边距 / 三线表 / 上标引文 / 中文编号 | 断行与分页仍不同 |
| **B 分页对齐** | A + 浮动体位置 + 段间距 + 表格宽度 | 页数一致，行内容仍有差异 |
| **C 逐行硬对齐** | B + 按 LaTeX 断点插 `w:br` + 逐行字距补偿 | **段落变硬排版，改文字必须重新生成** |

**什么时候不该用 C 档**：需要在 Word 里继续自由编辑文字。
硬对齐后每个断点都是硬回车，改一个字就可能让整段错位——
必须回到 `.tex` 改，然后重跑全链。

---

## 环境要求

| 项 | 要求 | 说明 |
|:---|:---|:---|
| 操作系统 | **Windows** | 闭环校正依赖 Word COM 导出 PDF |
| Microsoft Word | 本机安装 | 用于 `ExportAsFixedFormat` 导出 PDF 做实测 |
| Python | 3.10+ | 需能 `pip install` |
| Pandoc | 2.x+ | 承担 tex → docx 的实际转换 |
| LaTeX | 已编译出 PDF | **PDF 是唯一的对齐基准** |
| 字体 | Latin Modern Roman | 用 `fetch_fonts.py` 自动获取，见下 |

Python 依赖：

```bash
pip install python-docx lxml pymupdf fontTools pywin32
```

> **macOS / Linux**：转换链本身（A/B 档）可跑，但**闭环校正不可用**——
> 没有 Word COM 就无法导出 PDF 实测，C 档的逐行校准做不了。

---

## 快速开始

### 1. 获取并安装字体（一次性）

```bash
python scripts/fetch_fonts.py --install    # 获取 + 转换 + 用户级安装
python scripts/audit_lmfont.py             # 校验 Word 能否识别
```

`fetch_fonts.py` 会自动选择来源：

1. **本机已有 MiKTeX / TeX Live** → 直接用其自带的 OTF（最快，无需下载）
2. **没有** → 从 CTAN 镜像下载 `lm.zip`（中科大 / 清华 / 官方，自动降级重试）

拿到 OTF 后转成 TrueType 轮廓（13 个光学尺寸），修正 `name` 表与样式位，
再用户级安装（写 `HKCU`，**不动系统目录**）。

其他用法：

```bash
python scripts/fetch_fonts.py                    # 只获取转换，不安装
python scripts/fetch_fonts.py --from "C:\...\lm"  # 指定本地 OTF 目录
python scripts/fetch_fonts.py --mirror <url>     # 指定镜像
```

**装完必须重启 Word。** 卸载用 `uninstall_ttf.py`（按安装清单精确回滚）。

#### 为什么字体必须专门处理

**① 为什么不能用 Times New Roman 代替**
Latin Modern 是**光学尺寸族**，5/6/7/8/9/10/12/17pt 的字形宽度各不相同。
替代字体会让西文行宽产生系统性偏差，逐行对齐永远收不敛。

**② 为什么必须把 OTF 转成 TTF**
Word 对 CFF 轮廓的 OTF 支持不完整。实测把 `w:rFonts` 指向 LM Roman（OTF/CFF）后，
导出 PDF 里**西文字符整体消失**——Word 渲染阶段直接丢弃了这些字形。
转成 TrueType 二次曲线后正常。转换时 `upm` 必须提到 2048（2 的幂），
且**保留 `hmtx` 度量**（字宽必须与 LaTeX 完全一致，这是对齐的前提）。

**③ 为什么仓库不直接放 TTF**
二进制字体 2.8MB 占仓库体积 90%；改为从 CTAN 官方源获取，
用户拿到的是原版，也规避了再分发的许可细节。

### 2. 写配置

在项目目录放 `l2w.json`（样例见 `assets/l2w.example.json`）：

```json
{
  "base":   "D:/proj",
  "docx":   "D:/proj/out/paper.docx",
  "tex":    "D:/proj/paper.tex",
  "texpdf": "D:/proj/paper.pdf",
  "tmp":    "D:/proj/tmp",
  "tex_tag": "tex",
  "margins_cm": {"top": 2.4, "bottom": 2.4, "left": 2.5, "right": 2.5}
}
```

```bash
export L2W_CONF="D:/proj/l2w.json"
```

配置优先级：`L2W_CONF` 指向的 JSON > 当前目录 `l2w.json` > 环境变量 > 默认值。

### 3. 跑统一入口

```bash
python scripts/run_chain.py all --level C --dry-run   # 先看会跑什么
python scripts/run_chain.py all --level C             # 实际执行
python scripts/run_chain.py iter --rounds 3           # 闭环校正 3 轮
python scripts/run_chain.py check                     # 只验收
```

阶段：`font` / `base` / `pre` / `post` / `hard` / `iter` / `ref` / `check`。
`--level B` 只到分页对齐，`--level C` 含逐行硬对齐。
全链开始前会自动存一份 `bak_chain_start.docx`。

### 4. 验收

```bash
python scripts/export_cmp.py "<你的docx>" final
python scripts/cmp_pages.py tex final
```

---

## 七步链详解

统一入口内部就是这七步。排错时按此逐步定位。

### 第 1 步：提取 LaTeX 版面基准

```bash
python scripts/pdflines2.py paper.pdf tex
```

产出 `tex.json` + `tex.txt`：每行的**基线 y、字号、左右界 x0/x1、字体、文本**。
这是后续一切对齐的唯一基准。

同时需人工记录：版心宽、正文基线距（→ `w:spacing line` twips）、
各级标题字号、首行缩进量、摘要悬挂缩进量。

### 第 2 步：预处理 tex —— 占位标记法

Pandoc 会弄丢很多东西。做法是**先转成占位标记**（Pandoc 原样保留），后处理再重建：

| 原 LaTeX | 占位标记 | 为什么 |
|:---|:---|:---|
| `\cite{k1,k2}` | `⟦CITE:1,2⟧` | Pandoc 无 bib 时**直接丢弃** cite |
| `\ref{lb}` | `⟦REF:表\|3⟧` | 类型与编号分开，后处理据前文决定是否补「表/式/图」 |
| `\bibitem{k}` | `⟦BIB:n⟧` | 保留编号，后处理拆成独立段 |
| `\tabnote{}` | `⟦TABNOTE⟧` | 自定义命令 Pandoc 不认 |
| `\footnote{}` | `⟦FOOTNOTE:n⟧` | Pandoc 脚注在表格附近会错位 |

同时剥离 ctex 专属命令（`\ctexset`/`\zihao`/`\heiti`）、
把 `tikzpicture` 换成 `\includegraphics`、
清空 `thebibliography` 的宽度参数（否则 `{99}` 会变成正文里的杂散段）。

### 第 3 步：建模板 + Pandoc 转换

```bash
pandoc --print-default-data-file reference.docx > tmp/reference_base.docx
python scripts/build_reference.py
pandoc tmp/pre.tex -o tmp/pandoc_raw.docx \
       --reference-doc=tmp/reference_custom.docx --from=latex --to=docx
```

**必须基于 Pandoc 官方 `reference.docx` 改**，不能自己新建空 docx——
Pandoc 依赖里面一整套样式 ID（`Compact` / `FirstParagraph` / `TableCaption`…），
少了就会静默丢格式。

### 第 4 步：OOXML 后处理

```bash
python scripts/postprocess1120.py
```

十件事：文献拆段 → CITE 转上标 → REF 转中文交叉引用 → 表注归位 →
真实脚注 → **三线表** → 标题中文编号 → 图表题编号居中 →
正文缩进 2em + exact 行距 → 页边距。

尾部自动串跑精修链（**顺序不可换**）：
`refine_layout` → `fix_settings` → `fix_justify` → `fix_fontsize` → `fix_lig2` → `fix_ns`

### 第 5 步：逐行硬对齐

```bash
python scripts/hard_align.py             # 只看配对计划，不改文件
python scripts/hard_apply.py --apply     # 插 w:br + jc=left
python scripts/hard_spacing.py --apply   # 写 w:spacing 字距补偿
python scripts/hard_incr.py --apply      # 补配对新命中的段
```

配对逻辑是**单调序列对齐**：Word 段序与 LaTeX 段序同向，
只在上一命中之后向后找；前缀一致性以 **LaTeX 首行长度**为分母。
含公式/图片的段、表格单元格、图表题一律跳过。

### 第 6 步：闭环校正（关键，不可省）

静态度量与 Word 实际渲染有系统偏差，必须「导出 PDF → 实测 → 反馈」：

```bash
python scripts/export_cmp.py "<docx>" it1   # 导出 it1.pdf + 逐页对比
python scripts/hard_iter.py it1 --apply     # 只修溢出行
# 重复 2~3 轮，每轮换标签 it2 / it3
```

> **`hard_iter` 只修「行数溢出」的行，绝不做全量更新。**
> 全量更新必然发散——实测行偏差从 44 恶化到 69。

参考文献区单独精校（模型更准）：

```bash
python scripts/fix_ref3.py --apply           # 静态度量模型 + floor 求解
python scripts/fix_ref4.py it3 --nos 18,33   # 顽固行：PDF 实测反解兜底
```

### 第 7 步：局部版面修补（按需）

```bash
python scripts/fix_tblwidth.py --apply    # 表格宽度对齐版心
python scripts/fix_hang.py --apply        # 摘要/关键词悬挂缩进
python scripts/fix_refspace.py --apply --line 270 --after 80
python scripts/fix_eqspace.py --apply --line 270 --before 100 --after 100
python scripts/fix_footnote.py --apply --sz 15 --line 180
python scripts/move_block.py --list       # 浮动体整块移动
```

**每改一项都要重跑 `export_cmp.py` 看指标，变差立刻回滚。**

---

## 关键技术原理

这一节是项目里最费时间标定出来的量化结论，直接复用可省大量试错。

### 1. 字距（`w:spacing`）的真实生效模型

```
真实行宽 = nat + Σ(空格→CJK边界间距) + cur × n / 20 + 0.10
```

- `nat`：静态度量的自然宽度（含引文上标）
- `cur`：当前 `w:spacing` 值（单位 twip/字符）
- `n`：可施加字距的字符数（**不含**引文上标）
- **生效系数 = 1.0**（不是 0.968）
- **常数偏置 = +0.10 pt**（不是 0.90 pt）

> 早期把 CJK 边界间距重复计入，拟合出了错误的 0.968 / 0.90。
> 正确判据：字符间隙的**中位数**恰好等于 `cur / 20`。

反解所需字距：

```python
cur_new = floor((可用宽 - nat - Σ边界 - 0.10) / n * 20)
```

**必须用 `floor` 而不是 `round`**：溢出行哪怕只超 0.1pt 也会断行，
`round` 会留下残余（实测 7 行全部没修好）。

### 2. 空格→CJK 边界间距 ≈ 1.75 pt/处

中英文混排时，每个「空格紧跟中文」的边界会额外产生约 1.75pt 间距，
且 `autoSpaceDE=0` **压不掉**。算宽度时必须逐处累加。

### 3. 溢出行的自然宽必须用静态度量

PDF 里溢出行只渲染了前半部分。若用 PDF 实测 advance 当自然宽，
会把「少了 3 个词」误算成「需要大幅拉伸」，方向完全反了。

### 4. `w:br` 只锁断点，不防二次断行

插了 `w:br` 不等于万事大吉：**若该行自然宽 > 可用宽，末词仍会被推走**，
产生孤字行（`and` / `of` / `the` / `管`）。所以必须配合字距压缩。

### 5. 中文标点悬挂

中文行的右界 x1 可以超出版心右界约 12pt（标点悬挂），
这是**正常现象**，不是错误。判断溢出时要区分中西文。

### 6. 变音符与连字会拆 span

PDF 里 `MUÑOZ` 可能渲染成 `MUÑ` + `OZ` 两个 span，
`Definition` 含连字 `ﬁ`(U+FB01)。
前缀匹配失败时须放弃实测、退回静态度量 + 额外压缩 2 twip。

---

## 脚本总览

### 基础设施（3）

| 脚本 | 作用 |
|:---|:---|
| `l2wconf.py` | 统一路径配置，所有脚本都读它 |
| `docxio.py` | 安全写盘：关 Word 占用 + 原子替换 + 首次备份 |
| `run_chain.py` | 统一入口，阶段编排 |

### 字体（6）

`fetch_fonts.py`（**入口**：自动获取 + 转换 + 安装）/ `install_ttf.py` /
`uninstall_ttf.py` / `audit_lmfont.py` / `otf2ttf.py` / `fix_ttf_names.py`

### 转换链（4）

`pdflines2.py`（提取基准）/ `preprocess1120.py`（占位标记）/
`build_reference.py`（模板）/ `postprocess1120.py`（OOXML 后处理）

### 硬对齐（5）

`hard_align.py`（配对）/ `hard_apply.py`（插 br）/ `hard_spacing.py`（字距）/
`hard_incr.py`（增量补配）/ `hard_iter.py`（闭环修正）

### 文献区精校（5）

`fix_ref2.py` / `fix_ref3.py`（主力：静态模型 + floor）/
`fix_ref4.py`（顽固行兜底）/ `fix_refspace.py` / `fix_ref_spacing.py`

### 局部修补（9）

`refine_layout.py` / `fix_settings.py` / `fix_justify.py` / `fix_fontsize.py` /
`fix_lig2.py` / `fix_ns.py` / `fix_tblwidth.py` / `fix_hang.py` /
`fix_eqspace.py` / `fix_footnote.py`

### 浮动体与验收（4）

`float_table.py` / `move_block.py` / `export_cmp.py` / `cmp_pages.py`

### 自检（1）

`_smoke.py` —— import 全部脚本 + **SHA256 校验目标文件未被改动**

---

## 已知硬边界

这些是物理限制，**不是 bug，也无法通过调参消除**：

1. **`exact` 行距 twips 取整抖动**：±0.1pt × 500 行 ≈ 4 行偏差
2. **含公式的段落无法逐行锁定**：OMML 的 `m:oMath` 位置不能用 `w:br` 精确控制
3. **图内文字/表注混入行提取**：少数段落永久无法配对
4. **浮动体高度差异**是独立问题，字距补偿解决不了
5. **硬对齐后是硬排版**：改文字必须从 tex 重走全链

---

## 工程纪律

项目踩过的坑已整理成 **`references/pitfalls.md`（35 条）**，每条都带实测数据。
其中三条是血泪教训，值得单独强调：

### 纪律一：脚本模块级不得有副作用

早期 7 个脚本的主逻辑写在模块顶层，结果**自检时逐个 import 就把成品写坏了**：
文件从 326489 字节变成 318668，行偏差从 35 恶化到 56，**全程没有任何报错**。

> 所有脚本必须有 `main()` 守卫。模块级只允许 import、常量、纯函数定义。
> 不得实例化 `Document`、不得持有可变全局状态、不得建目录。

### 纪律二：自检必须校验文件哈希

「33/33 导入成功」**不等于安全**——上面那次事故里所有 import 都「成功」了。

```bash
# 必须在两种环境各跑一次
cd scripts && python _smoke.py                    # ① 无配置：验证可加载性
L2W_CONF=/path/l2w.json python _smoke.py          # ② 有配置：验证功能 + SHA256 一致
```

只跑 ② 会漏掉「脱离项目目录就不可用」的缺陷；只看「导入成功」会漏掉静默写盘。

### 纪律三：每轮存明确命名的基准备份

自动备份名（`bak_before_*`）会被后续脚本覆盖。
`move_block.py --undo` 恢复的是**它自己首次运行时的存档**，不一定是你想要的上一状态。
所以每轮改动前手动存一份如 `GOOD_35_0724.docx` 的明确基准。

---

## 作为 WorkBuddy Skill 使用

本仓库同时是一个可直接使用的 WorkBuddy Skill（根目录有 `SKILL.md`）。

```bash
# Windows
git clone <repo> "%USERPROFILE%\.workbuddy\skills\latex-to-word"
```

WorkBuddy 靠扫描 `~/.workbuddy/skills/*/SKILL.md` 自动发现技能，
**无需注册或打包**。安装后需**重启应用或新开会话**刷新技能缓存。

之后直接说「把这个 tex 转成 Word，要和 PDF 一模一样」即可触发。

---

## 许可与致谢

代码部分见 [LICENSE](LICENSE)（MIT）。

本仓库**不再分发字体二进制**。`fetch_fonts.py` 从 [CTAN](https://ctan.org/pkg/lm)
官方镜像获取 Latin Modern Roman，其版权归 GUST e-foundry 所有，
遵循 **GUST Font License (GFL)**，条款摘要见
[assets/ttf_lm/LICENSE-GFL.txt](assets/ttf_lm/LICENSE-GFL.txt)。
用户自行获取的字体受该许可约束，与本仓库代码的 MIT 许可无关。

本项目的方法论建立在一篇真实中文社科论文的转换实践上，
所有量化结论（字距系数、边界间距、收敛轨迹）均为实测所得，非理论推导。
