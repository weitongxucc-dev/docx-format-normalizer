# DOCX Multi-Industry Document Format Normalizer

Automatically normalize `.docx` formatting to match industry-specific standards. Supports 8 industry formats with audit-first workflow: review → clean → format → validate.

## Quick Start

### Step 1: Setup Environment (MANDATORY)

```bash
bash scripts/setup.sh
```

This creates an isolated virtual environment with `python-docx` and `lxml<6` (pinned to avoid macOS compatibility issues). The script outputs the Python path to use for all subsequent commands.

### Step 2: Format a Document

```bash
# Use the Python path from setup.sh output
PYTHON_PATH=".venv/bin/python"

$PYTHON_PATH scripts/docx_formatter.py \
  --input your_document.docx \
  --template templates/government_document.json \
  --output output.docx
```

### Step 3: Validate Output

```bash
$PYTHON_PATH scripts/validate.py \
  --input output.docx \
  --template templates/government_document.json
```

### Step 4: Visual Check (Optional)

```bash
$PYTHON_PATH scripts/visual_check.py \
  --input output.docx \
  --template templates/government_document.json
```

Visual check requires LibreOffice installed (`brew install --cask libreoffice` on macOS).

## Supported Industries

| ID | Template File | Industry | Standard |
|----|--------------|----------|----------|
| 01 | government_document.json | 党政机关公文 | GB/T 9704-2012 |
| 02 | bidding_dark_jiangsu.json | 江苏暗标 | Province-specific |
| 03 | bidding_dark_guizhou.json | 贵州暗标 | Province-specific |
| 04 | bidding_dark_jian.json | 吉安暗标 | City-specific |
| 05 | testing_report.json | CMA/CNAS检测报告 | DB61/T 1327.5-2020 |
| 06 | court_document.json | 法院诉讼文书 | Supreme Court Standard |
| 07 | academic_paper.json | 学术论文 | GB/T 7714 |
| 08 | construction_plan.json | 工程施工方案 | Industry Convention |

## Key Features

### Audit-First Workflow
- Scans document for non-compliant content before formatting
- Auto-cleans: colored text, colored cells, paragraph borders, navigation labels, cover page tables
- Reports all issues found and actions taken

### Cover Page Layout Compliance
- Validates cover page element positions (title at top, date at bottom)
- Removes non-standard elements (navigation labels, data tables, text boxes)
- Restructures layout with proper spacing between elements
- Each template defines industry-specific layout rules

### Font Fallback Chain
- Sets correct GB/T standard font names in XML (e.g., 小标宋体, 仿宋_GB2312)
- Adds macOS-compatible fallback fonts via `w:cs` attribute
- Document renders correctly on both Windows (with standard fonts) and macOS

### Style-Level Modification
- Modifies `styles.xml` definitions, not just individual runs
- Ensures consistent formatting throughout the document

## Scope

This tool **only modifies formatting/layout** — it does NOT:
- Alter body text content
- Generate structural elements (发文字号, 印章, 红色双线, etc.)
- Convert file formats (.doc, .pdf → .docx)
- Process password-protected documents

## Output

- **Formatted `.docx` file** with industry-standard formatting
- **`modification_report.json`** detailing all changes made
- **Validation results** (PASS/FAIL per check item)

## macOS Font Note

GB/T 9704 standard fonts (小标宋体, 仿宋_GB2312) are not pre-installed on macOS. The tool adds fallback fonts (STSong, STFangsong) via the `w:cs` XML attribute. For best visual results, install the standard font pack or open the document in Windows Word.

## Project Structure

```
docx-format-normalizer/
├── SKILL.md              # Agent-facing skill documentation
├── README.md             # This file
├── scripts/
│   ├── docx_formatter.py # Main formatting engine
│   ├── validate.py       # Output validator
│   ├── visual_check.py   # Visual color detection (LibreOffice + image analysis)
│   ├── setup.sh          # Environment setup script
│   └── requirements.txt  # Python dependencies
├── templates/
│   ├── government_document.json
│   ├── bidding_dark_jiangsu.json
│   ├── bidding_dark_guizhou.json
│   ├── bidding_dark_jian.json
│   ├── testing_report.json
│   ├── court_document.json
│   ├── academic_paper.json
│   └── construction_plan.json
└── tests/
    ├── test_pipeline.py  # Automated test pipeline
    └── test_cases.md     # Manual test cases
```

## License

MIT
