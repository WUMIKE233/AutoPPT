# AutoPPT

**中文** | [English](#english)

AutoPPT 是一个基于 Python、PyQt5、DeepSeek API 与 python-pptx 的桌面工具，用于生成 PPT、数据图表，并提供中英互译辅助能力。

它面向课堂课件、政务汇报、学术答辩和简约演示等场景，支持从主题或 Markdown 大纲生成结构化幻灯片，也支持从 Excel 数据生成常见图表。

## 功能亮点

- **AI PPT 生成**：根据主题或大纲生成章节内容、备注与演示文稿。
- **模板选择**：内置教学、政务、学术、简约等 PPT 模板。
- **图表生成**：从 Excel 数据生成饼图、柱状图、折线图、散点图、热力图、直方图和气泡图。
- **中英互译**：提供桌面端文本翻译入口，适合演示文稿内容整理。
- **桌面界面**：使用 PyQt5 构建，支持文件选择、输出目录配置、进度提示与取消任务。

## 环境要求

- Python 3.10+
- PyQt5
- python-pptx
- pandas
- matplotlib
- seaborn
- requests
- Pillow

## 快速开始

```bash
pip install PyQt5 python-pptx pandas matplotlib seaborn requests Pillow
```

配置 DeepSeek API Key：

```bash
# Windows PowerShell
$env:DEEPSEEK_API_KEY="your-key-here"

# macOS / Linux
export DEEPSEEK_API_KEY="your-key-here"
```

启动整合版应用：

```bash
python main.py
```

也可以运行旧版 PPT 生成入口：

```bash
python PPT.py
```

## 使用流程

1. 选择“生成 PPT”或“生成图表”。
2. 输入主题、粘贴 Markdown 大纲，或上传本地文件。
3. 选择模板与输出目录。
4. 先生成大纲并确认，再生成最终 PPT。
5. 图表模式下选择 Excel 文件、图表类型和输出格式。

## 安全说明

API Key 不应写入源码或提交到仓库。当前版本从 `DEEPSEEK_API_KEY` 环境变量读取密钥。若历史提交中曾出现真实密钥，请立即到服务商控制台吊销并重新生成。

## 仓库结构

```text
main.py                 整合版桌面应用
PPT.py                  旧版 PPT 生成入口
template_*.pptx         演示文稿模板
data_init/              UI 图片与初始化资源
output/                 示例或生成结果目录
old_version/            历史版本代码
```

## English

AutoPPT is a desktop application built with Python, PyQt5, the DeepSeek API, and python-pptx. It helps generate PowerPoint decks, data charts, and Chinese-English translation drafts.

It is designed for teaching materials, government reports, academic defenses, and clean business presentations. Users can generate structured slides from a topic or Markdown outline, and create charts from Excel data.

## Features

- **AI-assisted PPT generation**: creates section content, speaker notes, and presentation files from a topic or outline.
- **Template selection**: includes teaching, government, academic, and simple presentation templates.
- **Chart generation**: creates pie, bar, line, scatter, heatmap, histogram, and bubble charts from Excel files.
- **Chinese-English translation**: provides a desktop entry point for preparing bilingual presentation content.
- **Desktop UI**: built with PyQt5, including file pickers, output folder settings, progress states, and task cancellation.

## Requirements

- Python 3.10+
- PyQt5
- python-pptx
- pandas
- matplotlib
- seaborn
- requests
- Pillow

## Quick Start

```bash
pip install PyQt5 python-pptx pandas matplotlib seaborn requests Pillow
```

Set the DeepSeek API key:

```bash
# Windows PowerShell
$env:DEEPSEEK_API_KEY="your-key-here"

# macOS / Linux
export DEEPSEEK_API_KEY="your-key-here"
```

Run the integrated app:

```bash
python main.py
```

Or run the legacy PPT generator:

```bash
python PPT.py
```

## Workflow

1. Choose PPT generation or chart generation.
2. Enter a topic, paste a Markdown outline, or upload a local file.
3. Select a template and output folder.
4. Generate and review the outline, then create the final PPT.
5. In chart mode, select an Excel file, chart type, and output format.

## Security Notes

API keys should not be hardcoded or committed to the repository. This version reads the key from the `DEEPSEEK_API_KEY` environment variable. If a real key appeared in repository history, revoke it in the provider console and create a new one.

## Repository Layout

```text
main.py                 Integrated desktop application
PPT.py                  Legacy PPT generation entry point
template_*.pptx         Presentation templates
data_init/              UI images and startup assets
output/                 Example or generated outputs
old_version/            Historical versions
```

## License

No license file is currently included. Add one before distributing or accepting external contributions.