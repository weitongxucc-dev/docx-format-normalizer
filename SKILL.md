---
name: "docx-format-normalizer"
description: "多行业文档格式规范工具：自动将 .docx 文档格式规范化为行业标准（党政机关公文、招投标暗标、CMA/CNAS检测报告、法院诉讼文书、学术论文、工程施工方案）。当用户上传 .docx 并要求修正/规范化/标准化文档格式时调用。"
---

# 多行业文档格式规范 Skill

将用户上传的 .docx 文档格式规范化为行业标准格式。格式规则的唯一权威来源是
`templates/*.json`；本文件只描述触发方式、工作流与处置原则。

## 核心原则：精准审查、分级处置、全程可追溯

- **默认保留内容**：任何内容修改都必须由可解释、可验证的规则驱动，禁止仅凭
  粗略位置判断直接删除。
- **分级处置**：
  - 高置信度（如分隔符占比≥40% 的装饰导航行）→ 自动修复/删除；
  - 中置信度（如首个标题前的封面表格/图形，可能是合法内容）→ 保留并提示人工确认；
  - 低置信度 → 不处理。
- **授权删除**：仅当用户显式确认时，才用 `--sanitize` 删除中置信度内容。
- **全程可追溯**：所有处置记录规则 ID、位置、证据链与置信度（`--report` 输出，
  仅内部自检用，不交付用户）。
- **格式 vs 内容**：排版属性（页边距、字体、字号、行距、缩进、对齐、页眉页脚、
  页码、表格样式）自动规范；正文表述不改动，但文档自带的明显格式错误
  （彩色字、非标字体、装饰元素）会被规则化修复。

## 支持的行业格式

| 行业 | 模板文件 | 标准依据 |
|------|----------|----------|
| 党政机关公文 | `government_document.json` | GB/T 9704-2012《党政机关公文格式》 |
| 招投标暗标（江苏） | `bidding_dark_jiangsu.json` | 江苏省公共资源交易暗标编制规范 |
| 招投标暗标（贵州） | `bidding_dark_guizhou.json` | 黔发改法规〔2026〕196号 |
| 招投标暗标（吉安） | `bidding_dark_jian.json` | 吉安市发改委暗标盲评通知 |
| CMA/CNAS检测报告 | `testing_report.json` | RB/T 214-2017 + CNAS-CL01 |
| 法院诉讼文书 | `court_document.json` | 法〔2016〕221号文书样式 |
| 学术论文 | `academic_paper.json` | GB/T 7714—2025（过渡期可沿用 2015 版） |
| 工程施工方案 | `construction_plan.json` | GB/T 50502-2009 |

约束：仅支持 `.docx`（不支持 `.doc`/`.pdf`）；输入前校验扩展名与 ZIP 完整性。

## 工作流程

### 第0步：环境配置（必选）

```bash
bash scripts/setup.sh                # macOS / Linux
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1   # Windows
```

创建 `.venv` 并安装 `python-docx`、`lxml<6`（锁定版本避免兼容性问题），
输出 `SUCCESS: Python path = ...`，后续命令使用该路径。

### 第1步：接收输入

1. 确认文件为 `.docx`，否则提示转换后重新上传。
2. 向用户列出行业清单，**等用户选定目标行业后再生成**；暗标需确认省份变体。
3. 默认只生成用户选定的行业，不要一次性生成全部 8 个。

### 第2步：执行格式化

```bash
<PYTHON_PATH> scripts/docx_formatter.py \
  --input <input.docx> \
  --template templates/<template>.json \
  --output <输出中文名.docx>
```

- 不传 `--output` 时自动命名：`<原文件名>_<行业名称>_<YYYYMMDD>.docx`。
  输出必须使用中文名（如 `01_党政机关公文.docx`）。
- `--sanitize`：仅在用户已确认删除封面可疑表格/图形等中置信度内容时追加。
- `--report`：可选，仅内部自检；默认不加，不向用户目录输出报告。

### 第3步：验证（推荐）

```bash
<PYTHON_PATH> scripts/validate.py --input <output.docx> --template templates/<template>.json
```

失败时退出码为 1；可加 `--json <path>` 输出机器可读结果供流水线判断。
封面表格/导航标签存在时为 WARN（保留待人工确认），布局问题仍为 FAIL。

### 第4步：视觉检查（推荐）

```bash
<PYTHON_PATH> scripts/visual_check.py --input <output.docx> --template templates/<template>.json
```

LibreOffice 转 PDF → 页面图片 → 彩色像素检测；失败退出码为 1。
默认不落盘报告与图片；需要人工看图时加 `--save-images <dir>`。

### 第5步：交付

只交付：规范化后的 .docx（中文文件名）+ 一句说明（行业 + 主要修改项数）+
1-3 条建议人工确认项（如封面布局、表格）。

**不要交付**：`modification_report.json`、`visual_check_report.json`、
validate 逐项明细、渲染页面图片等中间产物（除非用户明确索要）。
若存在"保留并提示"项，必须在说明中告知用户待确认内容。

## 异常处理

1. 非 .docx / 文件损坏：提示转换或检查后重新上传。
2. 未选择行业：列出 6 类行业让用户选择（暗标需补充省份）。
3. 依赖问题：重跑 `setup.sh`；lxml 崩溃时 `pip install 'lxml>=4.9,<6' --force-reinstall`。
4. 无 Word 标题样式的文档：通过文本模式匹配应用层次字体规则。
5. 检测报告无分区关键词命中：分区检测退化为按首页处理，需人工复核输出。

## 关键技术约束

1. 封面清理为证据驱动分级处置，禁止弱特征盲删。
2. 检测报告必须按页面分区格式化（封面/说明/首页/数据/附件格式不同），
   分区按逻辑页切段、区域内相对序号定位。
3. 党政公文需精确几何参数：上37mm/左28mm 边距，22行×28字。
4. 暗标深度遍历清除加粗/斜体/下划线/彩色/页眉页脚/页码。
5. 固定行距（line_spacing_rule="fixed"）时段前段后距必须为 0。
6. 中英文分离：中文用 font_cn（eastAsia），西文用 font_en（ascii/hAnsi）。
7. 新增行业只需添加 JSON 模板，不改 Python 代码。
