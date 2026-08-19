# DOCX 多行业文档格式规范化工具

自动将 `.docx` 文档格式规范化为行业标准格式。支持 8 种行业格式，采用"审查→清理→格式化→验证"的工作流。

## 快速开始

### 第一步：环境配置（必选）

```bash
bash scripts/setup.sh
```

此脚本会创建独立的 Python 虚拟环境，安装 `python-docx` 和 `lxml<6`（锁定版本以避免 macOS 兼容性问题）。脚本会输出 Python 路径，后续命令需使用该路径。

### 第二步：格式化文档

```bash
# 使用 setup.sh 输出的 Python 路径
PYTHON_PATH=".venv/bin/python"

$PYTHON_PATH scripts/docx_formatter.py \
  --input your_document.docx \
  --template templates/government_document.json \
  --output output.docx
```

### 第三步：验证输出

```bash
$PYTHON_PATH scripts/validate.py \
  --input output.docx \
  --template templates/government_document.json
```

### 第四步：视觉检查（可选）

```bash
$PYTHON_PATH scripts/visual_check.py \
  --input output.docx \
  --template templates/government_document.json
```

视觉检查需要安装 LibreOffice（macOS 可通过 `brew install --cask libreoffice` 安装）。

## 支持的行业格式

| 编号 | 模板文件 | 行业 | 标准依据 |
|------|----------|------|----------|
| 01 | government_document.json | 党政机关公文 | GB/T 9704-2012 |
| 02 | bidding_dark_jiangsu.json | 江苏暗标 | 省级规范 |
| 03 | bidding_dark_guizhou.json | 贵州暗标 | 省级规范 |
| 04 | bidding_dark_jian.json | 吉安暗标 | 市级规范 |
| 05 | testing_report.json | CMA/CNAS检测报告 | DB61/T 1327.5-2020 |
| 06 | court_document.json | 法院诉讼文书 | 最高法院标准 |
| 07 | academic_paper.json | 学术论文 | GB/T 7714 |
| 08 | construction_plan.json | 工程施工方案 | 行业惯例 |

## 核心特性

### 审查优先工作流
- 格式化前先扫描文档中的不合规内容
- 自动清理：彩色文字、彩色单元格、段落装饰线、封面导航标签、封面数据表格
- 所有发现的问题和操作均记录在修改报告中

### 封面布局合规
- 校验封面元素位置（标题在上、日期在下）
- 删除非标准元素（导航标签、数据表格、文本框）
- 重构布局，确保元素间留白符合标准
- 每个模板定义行业专属的封面布局规则

### 字体兜底链
- 在 XML 中设置正确的 GB/T 标准字体名（如小标宋体、仿宋_GB2312）
- 通过 `w:cs` 属性添加 macOS 兼容的兜底字体
- 文档在 Windows（有标准字体）和 macOS 上均可正确渲染

### 样式级修改
- 直接修改 `styles.xml` 中的样式定义，而非仅修改个别 run
- 确保全文格式一致性

## 能力边界

本工具**仅修改格式/排版** —— 不会：
- 修改正文文字内容
- 生成结构性要素（发文字号、印章、红色双线等）
- 转换文件格式（.doc、.pdf → .docx）
- 处理加密文档

## 输出物

- **格式化后的 `.docx` 文件**（中文文件名，如 `01_党政机关公文.docx`）
- **`modification_report.json`** 修改报告，记录所有变更
- **验证结果**（每项检查的 PASS/FAIL）

## macOS 字体说明

GB/T 9704 标准字体（小标宋体、仿宋_GB2312）在 macOS 上默认未安装。本工具通过 `w:cs` XML 属性添加兜底字体（STSong、STFangsong 等）。如需最佳视觉效果，请安装标准字体包或在 Windows Word 中打开文档。

## 项目结构

```
docx-format-normalizer/
├── SKILL.md              # Skill 说明文档（Agent 调用用）
├── README.md             # 本文件
├── scripts/
│   ├── docx_formatter.py # 核心格式化引擎
│   ├── validate.py       # 输出验证器
│   ├── visual_check.py   # 视觉检查工具（LibreOffice + 图像分析）
│   ├── setup.sh          # 环境配置脚本
│   └── requirements.txt  # Python 依赖
├── templates/
│   ├── government_document.json    # 党政机关公文
│   ├── bidding_dark_jiangsu.json   # 江苏暗标
│   ├── bidding_dark_guizhou.json   # 贵州暗标
│   ├── bidding_dark_jian.json      # 吉安暗标
│   ├── testing_report.json        # CMA/CNAS检测报告
│   ├── court_document.json        # 法院诉讼文书
│   ├── academic_paper.json        # 学术论文
│   └── construction_plan.json     # 工程施工方案
└── tests/
    ├── test_pipeline.py  # 自动化测试流水线
    └── test_cases.md     # 手工测试用例
```

## 许可证

MIT
