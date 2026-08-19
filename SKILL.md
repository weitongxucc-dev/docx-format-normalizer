---
name: "docx-format-normalizer"
description: "Auto-normalize .docx formatting to industry standards (government docs, bidding dark-bids, CMA/CNAS testing reports, court documents, academic papers, construction plans). Invoke when user uploads a .docx and asks to fix/normalize/standardize document formatting."
---

# DOCX Multi-Industry Document Format Normalizer

## Overview

This Skill normalizes the formatting of uploaded .docx files to match industry-specific standards. It does NOT modify business text content — only layout properties (margins, fonts, font sizes, line spacing, indentation, alignment, headers/footers, page numbers, table styles).

### Critical Principle: Audit Before Formatting

The Skill does NOT blindly accept and format whatever the user uploads. It first **audits** the document for non-compliant content, then **auto-cleans** these issues before applying industry-standard formatting. Issues found and actions taken are reported in the modification report.

**Auto-detected and cleaned issues include:**
- Colored text (non-black fonts) → set to black
- Colored cell backgrounds → cleared to white
- Paragraph decorative borders (lines below headings) → removed
- Non-standard font references (Japanese fonts, Noto Sans, Courier) → replaced with standard Chinese fonts
- Non-standard font sizes → normalized to template specifications
- Theme accent colors → set to black
- Table border colors → set to black
- Style-level shading → cleared
- Cover page navigation labels (dot-separated taglines) → removed
- Cover page data tables → removed
- Cover page text boxes/shapes → removed
- Cover page layout non-compliance → restructured (title at top, date at bottom, proper spacing)

**Output filename rule**: Output files MUST use Chinese names matching the industry, e.g. `01_党政机关公文.docx`, NOT English names like `government_document_output.docx`.

### Supported Industries

| ID | Industry | Standard/Reference |
|----|----------|-------------------|
| 1 | Government Official Document | GB/T 9704-2012 |
| 2 | Bidding Dark-Bid (Jiangsu) | Province-specific |
| 3 | Bidding Dark-Bid (Guizhou) | Province-specific |
| 4 | Bidding Dark-Bid (Ji'an) | City-specific |
| 5 | CMA/CNAS Testing Report | DB61/T 1327.5-2020 |
| 6 | Court Litigation Document | Supreme Court Standard |
| 7 | Academic Paper | GB/T 7714 |
| 8 | Construction Plan | Industry Convention |

### Constraints

- **Input**: Only `.docx` files (not `.doc`, `.pdf`)
- **Scope**: Formatting/layout correction only; never alter body text content
- **Dark-bid mode**: Force-clears bold, italic, underline, colored fonts, headers, footers, page numbers
- **Government docs**: Precise margins (top 37mm, left 28mm), 22 lines/page, 28 chars/line
- **Testing reports**: Per-page-zone formatting (cover/home/data/attachment pages differ)

---

## Workflow

### Step 0: Environment Setup (MANDATORY)

Before processing any document, set up the Python environment:

```bash
# Run the setup script (creates a virtual environment with correct dependencies)
bash scripts/setup.sh
```

This script will:
1. Find a suitable Python (3.11+)
2. Create an isolated virtual environment in `.venv/`
3. Install `python-docx` and `lxml<6` (pinned to avoid compatibility issues)
4. Output the Python path to use for processing

**The script outputs the Python path** — use this path (not `python3`) for all subsequent commands. Example output: `SUCCESS: Python path = /path/to/.venv/bin/python`

If setup fails, see Exception Handling below.

### Step 1: Receive Input

When the user uploads a .docx file and asks to normalize/fix/standardize its formatting:

1. Confirm the file is `.docx` format. If not, inform the user: "仅支持 .docx 格式文件，请转换后重新上传。"
2. Ask the user to select an industry type from the 8 supported options above.
3. If the user selects a bidding dark-bid type, confirm the province/city variant.

### Step 2: Load Template Configuration

Based on the user's selected industry, identify the corresponding JSON template file:

| Industry | Template File |
|----------|--------------|
| Government Document | `templates/government_document.json` |
| Dark-Bid (Jiangsu) | `templates/bidding_dark_jiangsu.json` |
| Dark-Bid (Guizhou) | `templates/bidding_dark_guizhou.json` |
| Dark-Bid (Ji'an) | `templates/bidding_dark_jian.json` |
| Testing Report | `templates/testing_report.json` |
| Court Document | `templates/court_document.json` |
| Academic Paper | `templates/academic_paper.json` |
| Construction Plan | `templates/construction_plan.json` |

### Step 3: Execute Processing Engine

Run the script using the Python path from Step 0:

```bash
<PYTHON_PATH> scripts/docx_formatter.py \
  --input <input.docx> \
  --template templates/<template_name>.json \
  --output <output.docx> \
  --report <modification_report.json>
```

Example:
```bash
/path/to/.venv/bin/python scripts/docx_formatter.py \
  --input /path/to/document.docx \
  --template templates/government_document.json \
  --output /path/to/output_government.docx \
  --report /path/to/modification_report.json
```

The script will:
1. Analyze document structure (identify cover page, TOC, body sections)
2. **Audit & auto-cleanup**: Detect colored text, colored cell backgrounds, non-standard font sizes; auto-fix by setting text to black, clearing cell backgrounds, and flagging for normalization
3. Modify style definitions (styles.xml) for persistent formatting
4. Apply page setup (margins, paper size, grid)
5. Apply title heading formatting (if `title_heading` in template)
6. Apply body text formatting (font, size, line spacing, indent)
7. Apply heading formatting (unified style + text pattern detection)
8. Apply hierarchy font differentiation (黑体/楷体/仿宋 etc.)
9. Apply header/footer rules (clear or update)
10. Apply table formatting (borders, cell fonts)
11. Execute force-clear rules (if dark-bid mode)
12. Apply per-page-zone formatting (if testing report)
13. Generate a modification report JSON (includes audit findings + formatting changes)

### Step 4: Verify Output (RECOMMENDED)

Run the validation script to check key formatting properties:

```bash
<PYTHON_PATH> scripts/validate.py \
  --input <output.docx> \
  --template templates/<template_name>.json
```

This checks: page margins, body font, heading fonts, table formatting, **content compliance** (no colored text, no colored cell backgrounds), and reports PASS/FAIL for each.

### Step 5: Visual Check (RECOMMENDED)

Run the visual check script to verify the output **looks** correct by rendering to images and analyzing for colored content:

```bash
<PYTHON_PATH> scripts/visual_check.py \
  --input <output.docx> \
  --template templates/<template_name>.json \
  --save-images <output_dir>
```

This performs:
1. Converts docx to PDF via LibreOffice
2. Converts PDF to page images (30 DPI)
3. Analyzes each page image for colored (non-grayscale) pixels
4. Reports PASS/FAIL for color detection
5. Optionally saves page images for AI visual inspection

The visual check catches issues that XML-level validation cannot, such as:
- Table border colors inherited from table styles
- Theme accent colors referenced by style definitions
- Any residual colored visual elements

### Step 6: Output Results

Present the results to the user:

1. **Confirmation**: Display the selected industry and template parameters
2. **Processing Complete**: Brief summary of modifications made
3. **Modification Report**: Table showing `序号 | 修改位置 | 修改前 | 修改后`
4. **Output File**: Provide the normalized .docx file path
5. **Verification Report**: Show PASS/FAIL results from validate.py
6. **Manual Review Suggestions**: List items the user should visually verify

---

## Quick Start (One-Liner)

```bash
# Setup (run once)
bash scripts/setup.sh

# Process a document (replace paths)
$(cat .venv/.python_path 2>/dev/null || echo ".venv/bin/python") scripts/docx_formatter.py \
  --input "your_document.docx" \
  --template templates/government_document.json \
  --output "output.docx" \
  --report "report.json"
```

---

## Industry Format Rules (Authoritative Reference)

### 1. Government Official Document (GB/T 9704-2012)

| Parameter | Value |
|-----------|-------|
| Paper | A4 (210mm × 297mm) |
| Top margin (天头) | 37mm ± 1mm |
| Left margin (订口) | 28mm ± 1mm |
| Content area (版心) | 156mm × 225mm |
| Lines per page | 22 |
| Chars per line | 28 |
| Body text font | 仿宋_GB2312 |
| Body text size | 3号 (16pt) |
| Title font | 小标宋体 |
| Title size | 2号 (22pt) |
| Text color | Black |
| Page number font | 宋体, 4号 (14pt) |
| Printing | Double-sided, left binding |

**Hierarchy font differentiation**:
- Level 1 (一、): 黑体 (SimHei)
- Level 2 (（一）): 楷体 (KaiTi)
- Level 3 (1.): 仿宋 (FangSong)
- Level 4 (（1）): 仿宋 (FangSong)

### 2. Bidding Dark-Bid

**Force-clear rules (ALL variants)**:
- Clear ALL bold, italic, underline, strikethrough
- Clear ALL colored text (set to black)
- Remove ALL headers, footers, page numbers
- No table of contents
- Clear hyperlink coloring and custom char spacing/scale

**Jiangsu variant**: 宋体 小四号 12pt, single spacing, 2.5cm margins all sides
**Guizhou variant**: 宋体 四号 14pt, chart text 五号 10.5pt, Arabic hierarchy
**Ji'an variant**: 宋体 四号 14pt, 1.5x spacing, margins top/bottom 2.5cm left 3.0cm right 2.5cm

### 3. CMA/CNAS Testing Report

**Per-page-zone formatting (CRITICAL)**:
- Cover: title 黑体 60pt bold, report number 仿宋 三号, lab name 黑体 一号 bold
- Notice page: title 黑体 小二号 bold, content 黑体 小三号
- Home/Data/Attachment pages: header 黑体 小二号 bold, content 仿宋 小四号

### 4. Court Litigation Document

| Parameter | Value |
|-----------|-------|
| Court name font | 宋体 bold, 2号 (22pt), char spacing +2pt |
| Document name font | 大标宋体 bold, 1号 (26pt) |
| Case number & body | 仿宋, 3号 (16pt) |
| Top/Bottom margins | 30mm / 27mm |
| Page number | 4号 white Arabic numerals |

### 5. Academic Paper (GB/T 7714)

| Parameter | Value |
|-----------|-------|
| Chinese body font | 宋体, 小四号 (12pt) |
| English/digit font | Times New Roman, 小四号 (12pt) |
| Line spacing | 1.5x |
| Title font | 黑体, 小四号 (12pt) |
| Abstract font | 宋体, 五号 (10.5pt) |
| References | 宋体/TNR, 五号 (10.5pt) |
| CJK/Latin font separation | REQUIRED |

### 6. Construction Plan

| Parameter | Value |
|-----------|-------|
| Margins | Left 3.17cm, others 2.5cm |
| Body font | 宋体, 四号 (14pt) |
| Title font | 黑体, 三号 (16pt) |
| Line spacing | 1.5x |
| Hierarchy | Arabic decimal (1, 1.1, 1.1.1) |
| Western font | Times New Roman |
| Table text | One size smaller than body |

---

## Exception Handling

1. **Not a .docx file**: "仅支持 .docx 格式文件，请将文件转换为 .docx 后重新上传。"
2. **File corrupted**: "文档无法解析，文件可能已损坏，请检查后重新上传。"
3. **No industry selected**: "请选择目标行业类型：1.党政机关公文 2.招投标暗标 3.CMA/CNAS检测报告 4.法院诉讼文书 5.学术论文 6.工程施工方案。（暗标需补充省份参数）"
4. **python-docx not installed / lxml crash**:
   - Run `bash scripts/setup.sh` to create a fresh virtual environment
   - If lxml still crashes, try: `pip install 'lxml>=4.9,<6' --force-reinstall`
   - If using system Python, create an isolated venv: `python3 -m venv .venv && .venv/bin/pip install 'python-docx>=0.8.11' 'lxml>=4.9,<6'`
5. **Document has no Word heading styles**: The script will still process body text and apply hierarchy font rules via text pattern matching.
6. **Document has nested tables**: The script traverses all table cells including nested tables.
7. **Paragraphs with no runs (images/objects)**: These are skipped during run-level formatting.

---

## Key Technical Constraints

1. **Never modify body text** — only change formatting/layout properties
2. **Cover page protection** — cover page paragraphs keep their original alignment, only font/size is updated
3. **Style-level + run-level modification** — both styles.xml definitions AND direct run formatting are updated
4. **Testing reports must use per-page-zone formatting** — not uniform whole-document styling
5. **Government documents need precise page geometry** — 37mm/28mm margins, 22 lines, 28 chars
6. **Dark-bids need deep traversal** — walk every run, paragraph, table cell, header, footer to clear forbidden formatting
7. **Fixed line spacing requires zero paragraph spacing** — when line_spacing_rule is "fixed", space_before and space_after must be 0
8. **CJK/Latin font separation** — Chinese uses font_cn (eastAsia), English/digits use font_en (ascii/hAnsi)
