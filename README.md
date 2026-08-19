# DOCX Format Normalizer Skill

Auto-normalize `.docx` formatting to industry-specific standards. Supports 6 industries with 8 template configurations.

## Supported Industries

| # | Industry | Template | Standard |
|---|----------|----------|----------|
| 1 | Government Official Document | `government_document.json` | GB/T 9704-2012 |
| 2 | Bidding Dark-Bid (Jiangsu) | `bidding_dark_jiangsu.json` | Province standard |
| 3 | Bidding Dark-Bid (Guizhou) | `bidding_dark_guizhou.json` | 黔发改法规〔2026〕196号 |
| 4 | Bidding Dark-Bid (Ji'an) | `bidding_dark_jian.json` | City standard |
| 5 | CMA/CNAS Testing Report | `testing_report.json` | DB61/T 1327.5-2020 |
| 6 | Court Litigation Document | `court_document.json` | Supreme Court standard |
| 7 | Academic Paper | `academic_paper.json` | GB/T 7714 |
| 8 | Construction Plan | `construction_plan.json` | Industry convention |

## Quick Start

### Install Dependencies

```bash
pip install python-docx
```

### Process a Document

```bash
# Method 1: Use a config file
python3 scripts/docx_formatter.py --config config.json

# Method 2: Direct arguments
python3 scripts/docx_formatter.py \
  --input your_document.docx \
  --template templates/government_document.json \
  --output normalized.docx \
  --report report.json
```

### Config File Example

```json
{
  "input_file": "input.docx",
  "output_file": "output.docx",
  "report_file": "report.json",
  "template_path": "templates/government_document.json"
}
```

## Architecture

```
docx-format-normalizer/
├── SKILL.md                    # Skill definition (system prompt + rules)
├── README.md                   # This file
├── templates/                  # Industry format templates (JSON)
│   ├── government_document.json
│   ├── bidding_dark_jiangsu.json
│   ├── bidding_dark_guizhou.json
│   ├── bidding_dark_jian.json
│   ├── testing_report.json
│   ├── court_document.json
│   ├── academic_paper.json
│   └── construction_plan.json
├── scripts/
│   ├── docx_formatter.py       # Python processing engine
│   └── requirements.txt        # Dependencies
├── tests/
│   └── test_cases.md            # Test cases & verification checklist
└── docs/
    ├── docx-format-research.html  # Industry research report
    └── skill-tutorial.html        # Full tutorial
```

## Key Features

1. **Only formats, never modifies text content** — layout properties only
2. **Force-clear mode for dark-bids** — removes bold, italic, underline, colors, headers, footers, page numbers
3. **Per-page-zone formatting** — testing reports use different fonts per page section
4. **Hierarchy font differentiation** — government docs use 4 different fonts for heading levels
5. **CJK/Latin font separation** — academic papers set Chinese and English fonts separately
6. **Province-specific variants** — dark-bid templates differ by province
7. **Modification report** — outputs a JSON report of all changes made

## Industry Format Rules Summary

### Government Document (GB/T 9704-2012)
- Margins: top 37mm, left 28mm | Body: 仿宋 3号 (16pt) | 22 lines × 28 chars per page
- Title: 小标宋 2号 | Hierarchy: 黑体→楷体→仿宋→仿宋

### Bidding Dark-Bid
- **Force-clear**: bold, italic, underline, colors, headers, footers, page numbers
- Jiangsu: 宋体 小四号, single line spacing, margins all 2.5cm
- Guizhou: 宋体 四号, 1.5x spacing
- Ji'an: 宋体 四号, 1.5x, margins left 3.0cm

### Testing Report
- Per-page-zone: cover (黑体 60pt) / notice (黑体 小二号) / data (仿宋 小四号) / attachment (仿宋 小四号 bold)

### Court Document
- Court name: 宋体 bold 2号 (+2pt char spacing) | Body: 仿宋 3号 | 22 lines × 28 chars

### Academic Paper (GB/T 7714)
- Chinese: 宋体 小四号 | English: Times New Roman 小四号 | 1.5x line spacing | Left aligned

### Construction Plan
- Body: 宋体 四号 | Title: 黑体 三号 | 1.5x spacing | Left margin 3.17cm

## Testing

See `tests/test_cases.md` for 5 test cases covering all major scenarios:
1. Government document format correction
2. Jiangsu dark-bid force-clear
3. Academic paper CJK/Latin font separation
4. Testing report per-page-zone formatting
5. Court document formatting

## License

MIT
