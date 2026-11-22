# latex-to-word

> Convert Chinese LaTeX papers to Word with **line-by-line layout fidelity** to the LaTeX-compiled PDF.

**[中文文档](README.md)** · English summary below.

---

## The Problem

Academic workflows often demand: *"Write in LaTeX, but submit a Word file — and the layout must not change."*

Running `pandoc paper.tex -o paper.docx` produces a file that differs from the PDF in almost every respect: fonts, tables (grid instead of three-line), citations (dropped entirely), Chinese heading numbering (missing), line spacing, line breaks, and page count.

This project fixes each difference **quantitatively**, not by swapping templates.

## The Core Insight

> **TeX allows inter-word spaces to shrink AND stretch. Word's justified alignment only stretches.**

Measured TeX space adjustment: **−0.74 to +1.81 pt per space**. TeX will *compress* a space to fit one more word on a line; Word never does — it pushes the word to the next line.

**Consequence: no amount of font/size/spacing/margin tuning will make Word reproduce TeX's line breaks automatically.**

So instead of tuning, we:

1. Extract **every line's text and coordinates** from the LaTeX PDF (`pdflines2.py`)
2. Write TeX's line breaks **directly into Word** as `w:br` elements
3. Compute **per-line character-spacing compensation** (`w:spacing`) so each line exactly fills the text width
4. Export Word back to PDF, **measure actual line widths, and feed corrections back** (closed loop, 2–3 rounds)

## Results

A 12-page Chinese social-science paper (8 three-line tables, equations, 44 references, TikZ figures):

| Metric | Result |
|:---|:---|
| Page count | **12 = 12** (identical) |
| Total per-page line deviation | **35 lines** |
| Full-text line similarity (difflib) | **0.724** |

Convergence trajectory: `77 → 44 → 41 → 35` line deviation.

**We do not claim 100% fidelity.** Remaining deviations have concrete physical causes — see [Hard Limits](#hard-limits).

## Three Fidelity Levels

| Level | What it does | Trade-off |
|:---|:---|:---|
| **A — Style** | Fonts, sizes, spacing, margins, three-line tables, superscript citations, Chinese numbering | Line breaks and pagination still differ |
| **B — Pagination** | A + float placement, paragraph spacing, table widths | Page count matches, line content still differs |
| **C — Line-locked** | B + `w:br` at TeX break points + per-line character spacing | **Paragraphs become hard-formatted; editing text requires regeneration** |

Do **not** use level C if you need to keep editing text freely in Word.

## Requirements

- **Windows** (closed-loop calibration requires Word COM to export PDF)
- Microsoft Word installed locally
- Python 3.10+ with `python-docx lxml pymupdf fontTools pywin32`
- Pandoc 2.x+ — **need not be on PATH**; `pip install pypandoc-binary` is enough
  (the bundled pandoc is auto-detected)
- A LaTeX-compiled PDF as the alignment reference

On macOS/Linux the A/B conversion chain runs, but **level C calibration is unavailable** (no Word COM).

## Quick Start

```bash
# 1. Fetch + convert + install Latin Modern fonts (one-time, user-level, HKCU only)
python scripts/fetch_fonts.py --install
python scripts/audit_lmfont.py
# Restart Word afterwards.

# 2. Point to your config
export L2W_CONF="/path/to/l2w.json"

# 3. Run the chain
python scripts/run_chain.py all --level C --dry-run   # preview
python scripts/run_chain.py all --level C             # execute
python scripts/run_chain.py iter --rounds 3           # closed-loop calibration

# 4. Verify
python scripts/export_cmp.py "<your.docx>" final
python scripts/cmp_pages.py tex final
```

Config example: `assets/l2w.example.json`. Full pipeline documentation: **[README.md](README.md)** (Chinese).

## Why Latin Modern Needs Special Handling

**Optical-size family.** The 5/6/7/8/9/10/12/17pt variants have *different glyph widths*. Substituting Times New Roman introduces systematic Western-text width errors, and line-level alignment will never converge.

**OTF must be converted to TTF.** Word's support for CFF-outline OTF is incomplete: pointing `w:rFonts` at LM Roman (OTF/CFF) makes **Western characters vanish entirely** from the exported PDF — Word's renderer silently drops those glyphs. Converting to TrueType quadratic outlines fixes it. During conversion `upm` must be raised to 2048 (a power of two) while **preserving `hmtx` metrics** — glyph widths must match LaTeX exactly, which is the precondition for alignment.

**Fonts are fetched, not bundled.** `fetch_fonts.py` uses local MiKTeX/TeX Live OTFs if present, otherwise downloads `lm.zip` from CTAN mirrors (USTC / Tsinghua / official, with automatic fallback). This keeps the repo at ~380 KB instead of 3 MB and avoids font-redistribution licensing concerns. You receive the original upstream fonts under the GUST Font License.

Verified: 13/13 fonts convert successfully (upm=2048, 821 glyphs each); the download path yields all 13 required OTFs from a 17.8 MB archive.

## Key Quantitative Findings

Hard-won calibration constants — reuse them and skip the trial-and-error:

```
actual_line_width = nat + Σ(space→CJK boundary gaps) + cur × n / 20 + 0.10
```

- Character-spacing **effective coefficient = 1.0** (not 0.968)
- **Constant bias = +0.10 pt** (not 0.90 pt)
- Each space→CJK boundary adds **≈1.75 pt**, and `autoSpaceDE=0` cannot suppress it
- Solve with **`floor`, never `round`** — a line overflowing by even 0.1 pt still breaks
- For overflowing lines, natural width **must** come from static metrics, not PDF advance (the PDF only rendered the first half)

See `references/pitfalls.md` for all **36 documented pitfalls** with measured data.

## Hard Limits

Not bugs — physical constraints that no parameter tuning can remove:

1. `exact` line-spacing twips rounding jitter: ±0.1 pt × 500 lines ≈ 4 lines
2. Paragraphs containing equations cannot be line-locked (OMML `m:oMath` positioning)
3. In-figure text and table notes leak into PDF line extraction; some paragraphs never pair
4. Float height differences are an independent problem
5. After line-locking, paragraphs are hard-formatted — text edits require full regeneration

## Engineering Discipline

Four rules learned the hard way:

1. **No module-level side effects.** Seven scripts once had main logic at module top level; merely importing them during a smoke test silently corrupted the deliverable (326489 → 318668 bytes, deviation 35 → 56) **with no error raised**. Every script must have a `main()` guard.
2. **Smoke tests must verify file hashes.** "33/33 imports succeeded" does *not* mean safe — every import "succeeded" during that corruption incident. Run `_smoke.py` in **two environments**: bare (loadability) and configured (function + SHA256 unchanged).
3. **Keep explicitly named baseline backups each round.** Auto-backup names get overwritten; `--undo` restores the tool's *own first-run* snapshot, not your previous state.
4. **Never assume an external binary is on PATH.** `run_chain.py` resolves pandoc via PATH → pypandoc's bundled copy → common install dirs. Machines with only `pypandoc-binary` installed work out of the box; a bare `['pandoc', ...]` call would abort the whole chain with `FileNotFoundError`.

## Repository Layout

```
SKILL.md              WorkBuddy skill definition (auto-discovered)
README.md             Full documentation (Chinese)
scripts/              38 Python scripts
references/
  pitfalls.md         36 measured pitfalls
assets/
  ttf_lm/             Font license notice (TTFs generated by fetch_fonts.py)
  l2w.example.json    Config template
```

## Use as a WorkBuddy Skill

```bash
git clone <repo> "$HOME/.workbuddy/skills/latex-to-word"
```

WorkBuddy auto-discovers skills by scanning `~/.workbuddy/skills/*/SKILL.md` — no registration or packaging needed. Restart the app to refresh the skill cache.

## License

Code: **MIT** (see [LICENSE](LICENSE)).

Fonts are **not redistributed** by this repository. `fetch_fonts.py` retrieves Latin Modern Roman from [CTAN](https://ctan.org/pkg/lm); those fonts are copyright GUST e-foundry under the **GUST Font License** (summary in [assets/ttf_lm/LICENSE-GFL.txt](assets/ttf_lm/LICENSE-GFL.txt)). Fonts you fetch are governed by that license, independent of this repo's MIT code license.
