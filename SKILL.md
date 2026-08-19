---
name: "docx-format-normalizer"
description: "Auto-normalize .docx formatting to industry standards (government docs, bidding dark-bids, CMA/CNAS testing reports, court documents, academic papers, construction plans). Invoke when user uploads a .docx and asks to fix/normalize/standardize document formatting."
---

# DOCX Multi-Industry Document Format Normalizer

## Overview

This Skill normalizes the formatting of uploaded .docx files to match industry-specific standards. It does NOT modify business text content — only layout properties (margins, fonts, font sizes, line spacing, indentation, alignment, headers/footers, page numbers, table styles).

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

### Step 1: Receive Input

When the user uploads a .docx file and asks to normalize/fix/standardize its formatting:

1. Confirm the file is `.docx` format. If not, inform the user: "仅支持 .docx 格式文件，请转换后重新上传。"
2. Ask the user to select an industry type from the 8 supported options above.
3. If the user selects a bidding dark-bid type, confirm the province/city variant.

### Step 2: Load Template Configuration

Based on the user's selected industry, load the corresponding JSON configuration from `templates/`:

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

Read the JSON file to obtain the full parameter set: page size, margins, fonts, font sizes, line spacing, indentation, alignment, header/footer rules, table rules, force-clear rules.

### Step 3: Generate Processing Configuration

Output a JSON configuration object that the Python processing engine (`scripts/docx_formatter.py`) will consume. The configuration structure:

```json
{
  "industry": "government_document",
  "input_file": "/path/to/input.docx",
  "output_file": "/path/to/output.docx",
  "report_file": "/path/to/modification_report.json",
  "page_setup": {
    "paper_size": "A4",
    "margins": { "top": 37, "bottom": 35, "left": 28, "right": 26, "unit": "mm" },
    "orientation": "portrait"
  },
  "body_text": {
    "font_cn": "仿宋",
    "font_en": "Times New Roman",
    "font_size": "16pt",
    "line_spacing": "fixed",
    "lines_per_page": 22,
    "chars_per_line": 28,
    "first_line_indent": "2chars",
    "alignment": "justify"
  },
  "headings": [ ... ],
  "hierarchy_fonts": [ ... ],
  "header_footer": { ... },
  "tables": { ... },
  "force_clear": [ ... ]
}
```

### Step 4: Execute Python Processing Engine

Run the processing script:

```bash
python3 scripts/docx_formatter.py --config <config_json_path>
```

The script will:
1. Read the original .docx using `python-docx`
2. Extract current formatting information
3. Apply the template configuration
4. Execute force-clear rules (if dark-bid mode)
5. Output the normalized .docx file
6. Generate a modification report JSON

### Step 5: Output Results

Present the results to the user in this format:

1. **Confirmation**: Display the selected industry and sub-parameters
2. **Processing Complete**: Brief summary of what was done
3. **Modification Report**: Table showing: `序号 | 文档位置 | 修改前 | 修改后`
4. **Output File**: Provide the normalized .docx file path
5. **Verification Suggestions**: List items the user should manually review

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
| Issue org logo color | Red |
| Page number font | 半角宋体 Arabic numerals |
| Page number size | 4号 (14pt) |
| Printing | Double-sided |
| Binding | Left side |

**Hierarchy font differentiation (CRITICAL)**:
- Level 1 (一、): 黑体 (SimHei)
- Level 2 (（一）): 楷体 (KaiTi)
- Level 3 (1.): 仿宋 (FangSong)
- Level 4 (（1）): 仿宋 (FangSong)

### 2. Bidding Dark-Bid

**Force-clear rules (ALL variants)**:
- Clear ALL bold
- Clear ALL italic
- Clear ALL underline
- Clear ALL colored text (set to black)
- Remove ALL headers
- Remove ALL footers
- Remove ALL page numbers
- No table of contents

**Jiangsu variant**:
- Body font: 宋体 (SimSun), 小四号 (12pt)
- Line spacing: single (单倍)
- Margins: all sides 2.5cm
- First line indent: 2 chars
- Table border: 0.5pt solid
- Image caption: below image, centered

**Guizhou variant**:
- Body font: 宋体, 四号 (14pt)
- Chart text: 宋体, 五号 (10.5pt)
- No bold, color, italic, underline
- Hierarchy: Arabic numerals (1, 1.1, 1.1.1, ...)

**Ji'an variant**:
- Paper: A4, portrait, white bg black text
- Body font: 宋体, 四号 (14pt)
- Line spacing: 1.5x
- Margins: top 2.5cm, bottom 2.5cm, left 3.0cm, right 2.5cm
- Char spacing: standard, scale 100%
- First line indent: 2 chars
- Left aligned

### 3. CMA/CNAS Testing Report

**Per-page-zone formatting (CRITICAL — do NOT apply uniform formatting)**:

| Page zone | Position | Font | Size |
|-----------|----------|------|------|
| Cover - title | Line 2 | 黑体 bold | 60pt |
| Cover - report number | Line 3 | 仿宋 | 3号 (16pt) |
| Cover - info fields | Lower section | 仿宋 | 小三号 (15pt) |
| Cover - lab name | Last line | 黑体 bold | 1号 (26pt) |
| Notice page - title | Line 1 | 黑体 bold | 小二号 (18pt) |
| Notice page - content | Line 2+ | 黑体 | 小三号 (15pt) |
| Home/Data/Attachment - header | Lines 1-2 | 黑体 bold | 小二号 (18pt) |
| Home/Data/Attachment - number/page | Line 3 | 仿宋 | 小四号 (12pt) |
| Home/Data/Attachment - content | Middle | 仿宋 | 小四号 (12pt) |
| Attachment - content | Middle | 仿宋 bold | 小四号 (12pt) |

- Paper: A4, left binding
- Continuous page numbering
- Cover and TOC on separate pages

### 4. Court Litigation Document

| Parameter | Value |
|-----------|-------|
| Paper | A4 (297mm) |
| Court name font | 宋体 bold, 2号 (22pt), char spacing +2pt |
| Document name font | 大标宋体 bold, 1号 (26pt), char spacing +2pt |
| Case number & body | 仿宋, 3号 (16pt) |
| Lines per page | 22 |
| Chars per line | 28 |
| Top margin (天头) | 30mm |
| Bottom margin (地脚) | 27mm |
| Left margin > Right margin | Yes |
| Page number | 4号 half-width, white font Arabic numerals |
| Page number position | Single page right, double page left, 1 char from edge |
| Printing | Double-sided |
| Multi-page binding | Paste method (no staples) |

### 5. Academic Paper (GB/T 7714)

| Parameter | Value |
|-----------|-------|
| Paper | A4 |
| Chinese body font | 宋体 (SimSun), 小四号 (12pt) |
| English/digit font | Times New Roman, 小四号 (12pt) |
| Line spacing | 1.5x |
| Title font | 黑体 (SimHei), 小四号 (12pt) |
| Author name | 宋体 bold, 小四号 |
| Abstract font | 宋体, 五号 (10.5pt) |
| References title | 黑体/宋体 bold, 三号/四号 |
| References entries | 宋体/Times New Roman, 五号 (10.5pt) |
| Alignment | Left-aligned (all elements) |
| Note title | 黑体, 五号, flush left |
| Note content | 宋体, 小五号 |
| Chinese/English font separation | REQUIRED |

### 6. Construction Plan

| Parameter | Value |
|-----------|-------|
| Paper | A4 |
| Left margin | 3.17cm |
| Top/Right/Bottom margins | 2.5cm |
| Line spacing | 1.5x |
| Body font | 宋体 (SimSun), 四号 (14pt) |
| Title font | 黑体 (SimHei), 三号 (16pt) |
| Heading spacing | 1.5x, 0.5 line before and after |
| Body spacing | 0 before and after |
| Numbering | 1, 1.1, 1.1.1 (Arabic decimal) |
| Western font | Times New Roman |
| Table text | One size smaller than body |
| Foreign symbols | Italic for scalars, upright for units/functions |

---

## Exception Handling

1. **Not a .docx file**: "仅支持 .docx 格式文件，请将文件转换为 .docx 后重新上传。"
2. **File corrupted**: "文档无法解析，文件可能已损坏，请检查后重新上传。"
3. **No industry selected**: "请选择目标行业类型：1.党政机关公文 2.招投标暗标 3.CMA/CNAS检测报告 4.法院诉讼文书 5.学术论文 6.工程施工方案。（暗标需补充省份参数）"
4. **Missing python-docx**: "处理引擎缺少依赖，请运行 pip install python-docx 安装。"

---

## Key Technical Constraints

1. **Never modify body text** — only change formatting/layout properties
2. **Testing reports must use per-page-zone formatting** — not uniform whole-document styling
3. **Government documents need precise page geometry** — 37mm/28mm margins, 22 lines, 28 chars, 156mm×225mm content area
4. **Dark-bids need deep traversal** — walk every run, paragraph, table cell, header, footer to clear forbidden formatting
5. **Academic papers need CJK/Latin font separation** — Chinese uses 宋体, English/digits use Times New Roman
6. **Modification report must be clear** — tell the user exactly what was changed, before and after values
