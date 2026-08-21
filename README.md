# 多行业文档格式规范 Skill

自动将 `.docx` 文档格式规范化为行业标准格式。支持 8 种行业格式，采用"审查→清理→格式化→验证"的工作流。

## 快速开始

### 第一步：环境配置（必选）

按操作系统二选一：

```bash
# macOS / Linux
bash scripts/setup.sh
```

```powershell
# Windows (PowerShell)
powershell -ExecutionPolicy Bypass -File scripts\setup.ps1
```

脚本会创建独立的 Python 虚拟环境，安装 `python-docx` 和 `lxml<6`（锁定版本以避免兼容性问题），并输出 Python 路径，后续命令需使用该路径。

### 第二步：格式化文档

```bash
# 使用 setup 输出的 Python 路径
# macOS/Linux: .venv/bin/python   Windows: .venv\Scripts\python.exe
PYTHON_PATH=".venv/bin/python"

$PYTHON_PATH scripts/docx_formatter.py \
  --input your_document.docx \
  --template templates/government_document.json
```

不传 `--output` 时自动生成中文文件名：`<原文件名>_<行业>_<日期>.docx`，
如 `AI心灵陪伴机器人_报告_党政机关公文_20260821.docx`。也可用 `--output` 指定。

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
| 01 | government_document.json | 党政机关公文 | GB/T 9704-2012《党政机关公文格式》 |
| 02 | bidding_dark_jiangsu.json | 江苏暗标 | 江苏省公共资源交易暗标编制规范 |
| 03 | bidding_dark_guizhou.json | 贵州暗标 | 黔发改法规〔2026〕196号 |
| 04 | bidding_dark_jian.json | 吉安暗标 | 吉安市发改委暗标盲评通知 |
| 05 | testing_report.json | CMA/CNAS检测报告 | RB/T 214-2017 + CNAS-CL01 |
| 06 | court_document.json | 法院诉讼文书 | 法〔2016〕221号文书样式 |
| 07 | academic_paper.json | 学术论文 | GB/T 7714—2025（过渡期可沿用 2015 版） |
| 08 | construction_plan.json | 工程施工方案 | GB/T 50502-2009 |

## 核心特性

### 精准审查、分级处置、全程可追溯
- 格式化前先扫描文档中的不合规内容，任何内容修改都由可解释的规则驱动
- 高置信度问题自动修复：彩色文字/单元格、段落装饰线、高分隔符占比的导航行
- 中置信度内容（如首个标题前的封面表格/图形）默认保留并提示，
  经用户确认后可用 `--sanitize` 授权删除
- 所有处置记录规则 ID、位置、证据链与置信度（`--report`，仅内部自检用）

### 封面布局合规
- 校验封面元素位置（标题在上、日期在下）
- 清理装饰性元素（导航标签、装饰线）；表格/图形默认保留待人工确认
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
- 修改报告、验证明细、视觉图片等中间产物默认不生成、不交付；
  仅内部自检时通过 `--report` / `--json` / `--save-images` 显式开启

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
