---
name: "docx-format-normalizer"
description: "Auto-normalize .docx formatting to industry standards (government docs, bidding dark-bids, CMA/CNAS testing reports, court documents, academic papers, construction plans). Invoke when user uploads a .docx and asks to fix/normalize/standardize document formatting."
---

# 多行业文档格式规范 Skill

## 概述

本 Skill 将用户上传的 .docx 文档格式规范化为行业标准格式。**不修改正文文字内容** —— 仅修改排版属性（页边距、字体、字号、行距、缩进、对齐、页眉页脚、页码、表格样式）。

### 核心原则：先审查后格式化

本 Skill **不会**盲目接受和格式化用户上传的任何内容。它首先**审查**文档中的不合规内容，然后**自动清理**这些问题，最后再应用行业标准格式。所有发现的问题和操作均记录在修改报告中。

**自动检测和清理的问题包括：**
- 彩色文字（非黑色字体）→ 设为黑色
- 彩色单元格背景 → 清除为白色
- 段落装饰边框（标题下横线）→ 删除
- 非标准字体引用（日文字体、Noto Sans、Courier）→ 替换为标准中文字体
- 非标准字号 → 规范化为模板规格
- 主题强调色 → 设为黑色
- 表格边框色 → 设为黑色
- 样式级底纹 → 清除
- 封面导航标签（点号分隔的标签行）→ 删除
- 封面数据表格 → 删除
- 封面文本框/图形 → 删除
- 封面布局不合规 → 重构（标题在上、日期在下、适当留白）

**输出文件名规则**：输出文件必须使用与行业匹配的中文名，如 `01_党政机关公文.docx`，不能用英文名如 `government_document_output.docx`。

### 支持的行业格式

| 编号 | 行业 | 标准依据 |
|------|------|----------|
| 1 | 党政机关公文 | GB/T 9704-2012 |
| 2 | 招投标暗标（江苏） | 省级规范 |
| 3 | 招投标暗标（贵州） | 省级规范 |
| 4 | 招投标暗标（吉安） | 市级规范 |
| 5 | CMA/CNAS检测报告 | DB61/T 1327.5-2020 |
| 6 | 法院诉讼文书 | 最高法院标准 |
| 7 | 学术论文 | GB/T 7714 |
| 8 | 工程施工方案 | 行业惯例 |

### 约束条件

- **输入**：仅支持 `.docx` 格式文件（不支持 `.doc`、`.pdf`）
- **范围**：仅修改格式/排版，不修改正文文字内容
- **暗标模式**：强制清除加粗、斜体、下划线、彩色字体、页眉、页脚、页码
- **党政公文**：精确页边距（上37mm、左28mm），每页22行×28字
- **检测报告**：按页面分区格式化（封面/首页/数据页/附件页格式不同）

---

## 工作流程

### 第0步：环境配置（必选）

处理任何文档前，先配置 Python 环境：

```bash
# 运行配置脚本（创建带正确依赖的虚拟环境）
bash scripts/setup.sh
```

此脚本会：
1. 查找合适的 Python（3.11+）
2. 在 `.venv/` 中创建独立虚拟环境
3. 安装 `python-docx` 和 `lxml<6`（锁定版本避免兼容性问题）
4. 输出用于处理的 Python 路径

**脚本会输出 Python 路径** —— 后续所有命令使用此路径（而非 `python3`）。示例输出：`SUCCESS: Python path = /path/to/.venv/bin/python`

如配置失败，见下方"异常处理"。

### 第1步：接收输入

当用户上传 .docx 文件并要求规范化/修复/标准化格式时：

1. 确认文件为 `.docx` 格式。如不是，告知用户："仅支持 .docx 格式文件，请转换后重新上传。"
2. **先向用户列出可生成的行业**（党政机关公文 / 招投标暗标 / CMA检测报告 /
   法院诉讼文书 / 学术论文 / 工程施工方案），**等用户选定目标行业后再生成**。
3. 如用户选择暗标类型，确认省份/城市变体。

> **重要**：默认只生成用户选定的那一个行业，**不要一次性生成全部8个**。
> 仅当用户明确要求"对比测试/全部生成"时才批量处理。

### 第2步：加载模板配置

根据用户选择的行业，识别对应的 JSON 模板文件：

| 行业 | 模板文件 |
|------|----------|
| 党政机关公文 | `templates/government_document.json` |
| 江苏暗标 | `templates/bidding_dark_jiangsu.json` |
| 贵州暗标 | `templates/bidding_dark_guizhou.json` |
| 吉安暗标 | `templates/bidding_dark_jian.json` |
| CMA/CNAS检测报告 | `templates/testing_report.json` |
| 法院诉讼文书 | `templates/court_document.json` |
| 学术论文 | `templates/academic_paper.json` |
| 工程施工方案 | `templates/construction_plan.json` |

### 第3步：执行处理引擎

使用第0步输出的 Python 路径运行脚本：

```bash
<PYTHON_PATH> scripts/docx_formatter.py \
  --input <input.docx> \
  --template templates/<template_name>.json \
  --output <output.docx>
```

> `--report <modification_report.json>` 为**可选**参数，仅内部自检时使用，
> 默认不要加，避免向用户目录输出报告文件。

示例：
```bash
/path/to/.venv/bin/python scripts/docx_formatter.py \
  --input /path/to/document.docx \
  --template templates/government_document.json \
  --output /path/to/01_党政机关公文.docx
```

脚本执行流程：
1. 分析文档结构（识别封面页、目录、正文区域）
2. **审查与自动清理**：检测彩色文字、彩色单元格、非标准字号；自动修复——文字设为黑色、清除单元格背景、标记为规范化
3. 修改样式定义（styles.xml）确保持久格式化
4. 应用页面设置（边距、纸张大小、网格）
5. 应用标题格式（如模板中有 `title_heading`）
6. 应用正文格式（字体、字号、行距、缩进）
7. 应用标题格式（统一样式 + 文本模式识别）
8. 应用层次字体区分（黑体/楷体/仿宋等）
9. 应用页眉页脚规则（清除或更新）
10. 应用表格格式（边框、单元格字体）
11. 执行强制清除规则（暗标模式）
12. 按页面分区格式化（检测报告）
13. 生成修改报告 JSON（含审查结果 + 格式变更）

### 第4步：验证输出（推荐）

运行验证脚本检查关键格式属性：

```bash
<PYTHON_PATH> scripts/validate.py \
  --input <output.docx> \
  --template templates/<template_name>.json
```

检查项：页边距、正文字体、标题字体、表格格式、**内容合规性**（无彩色文字、无彩色单元格），并报告每项 PASS/FAIL。

### 第5步：视觉检查（推荐）

运行视觉检查脚本，通过渲染为图片并分析彩色内容来验证输出**看起来**正确：

```bash
<PYTHON_PATH> scripts/visual_check.py \
  --input <output.docx> \
  --template templates/<template_name>.json \
  --save-images <output_dir>
```

执行内容：
1. 通过 LibreOffice 将 docx 转为 PDF
2. 将 PDF 转为页面图片（30 DPI）
3. 分析每页图片是否有彩色（非灰度）像素
4. 报告颜色检测 PASS/FAIL
5. 可选保存页面图片供 AI 视觉检查

视觉检查能发现 XML 级验证无法发现的问题：
- 表格样式继承的边框色
- 样式定义引用的主题强调色
- 任何残留的彩色视觉元素

### 第6步：输出结果

**交付给用户的内容必须精简**——用户只需要最终的 .docx 和一句简要说明，
不需要看一堆中间报告文件。

1. **输出文件**：只提供规范化后的 .docx 文件路径（中文文件名）
2. **简要说明**：一句话说明行业类型 + 主要修改项数
3. **人工检查建议**：列 1-3 条用户应视觉确认的项目（如封面、表格）

**不要向用户呈现的中间产物**（仅供内部验证，默认不生成到用户目录）：
- `modification_report.json`、`visual_check_report.json` 等报告文件
- validate.py 的逐项 PASS/FAIL 明细
- 视觉检查渲染的页面图片

> 报告与验证只在内部用于自检。除非用户明确索要，不要把 report.json
> 等文件放进交付目录或逐条展示。运行 formatter 时不加 `--report` 参数
> 即不会生成报告文件。

---

## 快速开始（一行命令）

```bash
# 环境配置（运行一次）
bash scripts/setup.sh

# 处理文档（替换路径）
$(cat .venv/.python_path 2>/dev/null || echo ".venv/bin/python") scripts/docx_formatter.py \
  --input "your_document.docx" \
  --template templates/government_document.json \
  --output "output.docx" \
  --report "report.json"
```

---

## 行业格式规则（权威参考）

### 1. 党政机关公文（GB/T 9704-2012）

| 参数 | 值 |
|------|------|
| 纸张 | A4（210mm × 297mm） |
| 上边距（天头） | 37mm ± 1mm |
| 左边距（订口） | 28mm ± 1mm |
| 版心 | 156mm × 225mm |
| 每页行数 | 22 |
| 每行字数 | 28 |
| 正文字体 | 仿宋_GB2312 |
| 正文字号 | 三号（16pt） |
| 标题字体 | 小标宋体 |
| 标题字号 | 二号（22pt） |
| 文字颜色 | 黑色 |
| 页码字体 | 宋体，四号（14pt） |
| 印刷 | 双面，左侧装订 |

**层次字体区分**：
- 一级（一、）：黑体
- 二级（（一））：楷体
- 三级（1.）：仿宋
- 四级（（1））：仿宋

### 2. 招投标暗标

**强制清除规则（所有变体）**：
- 清除所有加粗、斜体、下划线、删除线
- 清除所有彩色文字（设为黑色）
- 删除所有页眉、页脚、页码
- 不设目录
- 清除超链接颜色和自定义字符间距/缩放

**江苏变体**：宋体 小四号 12pt，单倍行距，四周2.5cm边距
**贵州变体**：宋体 四号 14pt，图表文字五号 10.5pt，阿拉伯数字层次
**吉安变体**：宋体 四号 14pt，1.5倍行距，上下2.5cm 左3.0cm 右2.5cm边距

### 3. CMA/CNAS检测报告

**按页面分区格式化（关键）**：
- 封面：标题 黑体 60pt 加粗，报告编号 仿宋 三号，实验室名称 黑体 一号 加粗
- 说明页：标题 黑体 小二号 加粗，内容 黑体 小三号
- 首页/数据页/附件页：表头 黑体 小二号 加粗，内容 仿宋 小四号

### 4. 法院诉讼文书

| 参数 | 值 |
|------|------|
| 法院名称字体 | 宋体 加粗，二号（22pt），字符间距+2pt |
| 文书名称字体 | 大标宋体 加粗，一号（26pt） |
| 案号及正文 | 仿宋，三号（16pt） |
| 上下边距 | 30mm / 27mm |
| 页码 | 四号 白体阿拉伯数字 |

### 5. 学术论文（GB/T 7714）

| 参数 | 值 |
|------|------|
| 中文正文字体 | 宋体，小四号（12pt） |
| 英文/数字字体 | Times New Roman，小四号（12pt） |
| 行距 | 1.5倍 |
| 标题字体 | 黑体，小四号（12pt） |
| 摘要字体 | 宋体，五号（10.5pt） |
| 参考文献 | 宋体/TNR，五号（10.5pt） |
| 中英文分离 | 必须 |

### 6. 工程施工方案

| 参数 | 值 |
|------|------|
| 边距 | 左3.17cm，其余2.5cm |
| 正文字体 | 宋体，四号（14pt） |
| 标题字体 | 黑体，三号（16pt） |
| 行距 | 1.5倍 |
| 层次编号 | 阿拉伯小数（1, 1.1, 1.1.1） |
| 西文字体 | Times New Roman |
| 表格文字 | 比正文小一号 |

---

## 异常处理

1. **非 .docx 文件**："仅支持 .docx 格式文件，请将文件转换为 .docx 后重新上传。"
2. **文件损坏**："文档无法解析，文件可能已损坏，请检查后重新上传。"
3. **未选择行业**："请选择目标行业类型：1.党政机关公文 2.招投标暗标 3.CMA/CNAS检测报告 4.法院诉讼文书 5.学术论文 6.工程施工方案。（暗标需补充省份参数）"
4. **python-docx 未安装 / lxml 崩溃**：
   - 运行 `bash scripts/setup.sh` 创建全新虚拟环境
   - 如 lxml 仍崩溃，尝试：`pip install 'lxml>=4.9,<6' --force-reinstall`
   - 如使用系统 Python，创建独立 venv：`python3 -m venv .venv && .venv/bin/pip install 'python-docx>=0.8.11' 'lxml>=4.9,<6'`
5. **文档无 Word 标题样式**：脚本仍会处理正文，并通过文本模式匹配应用层次字体规则。
6. **文档有嵌套表格**：脚本会遍历所有表格单元格，包括嵌套表格。
7. **段落无 run（图片/对象）**：在 run 级格式化时跳过。

---

## 关键技术约束

1. **不修改正文** —— 仅修改格式/排版属性
2. **封面保护** —— 封面段落保留原始对齐方式，仅更新字体/字号
3. **样式级 + run 级修改** —— 同时更新 styles.xml 定义和直接 run 格式
4. **检测报告必须按页面分区格式化** —— 不能统一全文样式
5. **党政公文需精确页面几何参数** —— 37mm/28mm 边距，22行，28字
6. **暗标需深度遍历** —— 遍历每个 run、段落、表格单元格、页眉、页脚以清除禁止格式
7. **固定行距需零段间距** —— 当 line_spacing_rule 为 "fixed" 时，space_before 和 space_after 必须为 0
8. **中英文分离** —— 中文使用 font_cn（eastAsia），英文/数字使用 font_en（ascii/hAnsi）
