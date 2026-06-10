<div align="center">

# 📄 MinerU Batch Processor

**MinerU 文档批量处理工具 — 图形化界面**

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python)
![MinerU SDK](https://img.shields.io/badge/MinerU_SDK-0.2.5+-green)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-orange)
![License](https://img.shields.io/badge/License-Apache_2.0-blue)

<img src="screenshot.png" alt="MinerU Batch Processor" width="700"/>

一个基于 **MinerU API** 的桌面端文档批量处理工具，提供直观的图形界面，支持批量将 PDF、Office 文档、图片等转换为 Markdown、DOCX、HTML、LaTeX 等结构化格式。

[功能特性](#-功能特性) •
[界面预览](#-界面预览) •
[快速开始](#-快速开始) •
[使用指南](#-使用指南) •
[常见问题](#-常见问题)

</div>

---

## ✨ 功能特性

### 📥 文件管理
| 功能 | 说明 |
|------|------|
| **多文件选择** | 支持 Ctrl/Shift 多选，一次添加多个文件 |
| **文件夹导入** | 递归扫描文件夹内所有支持格式的文档 |
| **文件列表管理** | 查看文件大小、状态，支持移除/清空操作 |
| **广泛格式支持** | PDF、Word、Excel、PPT、图片、文本、电子书、压缩包等 |

### 🔧 处理选项
- **OCR 文字识别** — 对扫描件和图片文档自动进行文字识别
- **公式识别** — 自动提取文档中的数学公式（LaTeX 格式）
- **表格识别** — 识别并结构化文档中的表格内容
- **多语言支持** — 中文、英文、日文、韩文、法文、德文等 12 种语言
- **模型选择** — 默认模型 / 精确模式（p2p）/ 快速模式（fast）

### 💾 输出格式
- **Markdown** — 结构化文本，代码友好
- **DOCX** — Word 文档格式
- **HTML** — 网页格式
- **LaTeX** — 学术排版格式
- **保存全部** — 一键导出所有格式 + 提取的图片

### 🎨 用户界面
- 标签式布局，功能分区清晰
- 实时进度条 + 每文件状态更新
- 深色主题日志控制台，彩色高亮
- 内置帮助文档

---

## 🖥 界面预览

```
┌──────────────────────────────────────────────────────────┐
│  📄 MinerU 文档批量处理工具                                │
├──────────────────────────────────────────────────────────┤
│  ┌─ API 配置 ──────────────────────────────────────────┐ │
│  │ Token / API Key: [●●●●●●●●●●●●●●●●●●●●●●●●] [👁]  │ │
│  │ Base URL:       [https://mineru.net/api/v4     ]    │ │
│  └──────────────────────────────────────────────────────┘ │
│  ┌─ 文件选择 ──────────────────────────────────────────┐ │
│  │ ○ 选择单个/多个文件    ● 选择整个文件夹   已选择 5 个  │ │
│  │ [📂 添加文件] [📁 添加文件夹] [🗑 清空] [✂ 移除]    │ │
│  │ ┌────────────────────────────────────────────────┐  │ │
│  │ │ # │ 文件名           │ 大小    │ 状态          │  │ │
│  │ │ 1 │ report.pdf       │ 2.3 MB │ ✓ 完成        │  │ │
│  │ │ 2 │ paper.docx       │ 1.1 MB │ ✓ 完成        │  │ │
│  │ │ 3 │ slides.pptx      │ 5.7 MB │ 处理中...     │  │ │
│  │ └────────────────────────────────────────────────┘  │ │
│  └──────────────────────────────────────────────────────┘ │
│  ┌─ 输出格式 ────┐  ┌─ 处理选项 ────────────────────┐   │
│  │ ☑ Markdown    │  │ ☑ 启用 OCR                     │   │
│  │ ☐ DOCX        │  │ ☑ 启用公式识别                 │   │
│  │ ☐ HTML        │  │ ☑ 启用表格识别                 │   │
│  │ ☐ LaTeX       │  │ 文档语言: [中文     ▼]         │   │
│  │ ☐ 保存全部    │  │ 处理模型: [默认模型 ▼]         │   │
│  └───────────────┘  └────────────────────────────────┘   │
│  [▶ 开始处理] [■ 取消] ████████████░░░░ 3/5 就绪         │
│  状态: 处理完成 - 成功 5, 失败 0, 用时 12.3s             │
└──────────────────────────────────────────────────────────┘
```

---

## 🚀 快速开始

### 前置要求
- **Python 3.8+**
- **MinerU API Token** — 前往 [mineru.net](https://mineru.net) 注册获取

### 安装

```bash
# 1. 安装 MinerU SDK
pip install mineru-open-sdk

# 2. 克隆仓库
git clone https://github.com/B0Y-Studio/MinerU-Batch.git
cd MinerU-Batch

# 3. 运行程序
python MinerU_Batch_Processor.py
```

Windows 用户也可以直接双击 **`启动工具.bat`** 运行。

### 环境变量（可选）
```bash
# 设置 Token 环境变量，程序启动时自动读取
export MINERU_TOKEN=your_token_here
```

---

## 📖 使用指南

### 基本流程

1. **填写 Token** — 在 API 配置区输入你的 MinerU Token
2. **选择文件** — 点击「添加文件」或「添加文件夹」导入文档
3. **配置输出** — 勾选需要的格式和处理选项
4. **开始处理** — 点击「开始处理」，等待处理完成
5. **查看结果** — 每个文件在处理完成后自动保存到独立的子目录

### 输出目录结构
```
输出目录/
├── 文档1名称/
│   ├── 文档1名称.md
│   ├── 文档1名称.docx
│   ├── 文档1名称.html
│   └── images/
│       ├── img_001.png
│       └── img_002.png
├── 文档2名称/
│   ├── 文档2名称.md
│   └── images/
│       └── ...
└── ...
```

### 支持的文件格式

| 类别 | 格式 |
|------|------|
| 文档 | PDF, DOC, DOCX, PPT, PPTX, XLS, XLSX |
| 图片 | JPG, JPEG, PNG, BMP, TIFF |
| 文本 | TXT, MD, HTML, HTM |
| 电子书 | EPUB, MOBI |
| 数据 | CSV, JSON, XML |
| 压缩包 | ZIP |

---

## ⚙ 高级用法

### 自定义 Base URL
如果需要使用私有部署的 MinerU 服务，可以在界面中修改 Base URL。

### 模型选择
| 模型 | 适用场景 |
|------|----------|
| **默认模型** | 通用文档处理 |
| **精确模式 (p2p)** | 含复杂排版的 PDF，追求更高准确率 |
| **快速模式 (fast)** | 对速度有要求的批量处理 |

---

## ❓ 常见问题

**Q: 如何处理 Token？**
A: Token 可在界面上直接输入（支持密码遮挡显示），也可以设置在 `MINERU_TOKEN` 环境变量中。

**Q: 文件处理失败怎么办？**
A: 查看「运行日志」Tab 获取详细错误信息。常见原因：文件过大、页数超限、Token 配额不足。

**Q: 支持并发处理吗？**
A: 目前使用 MinerU SDK 的 `extract()` 方法逐文件处理，每个文件等待上一个完成。如需更高的并发量，可以在代码中改用 `extract_batch()`。

**Q: 输出文件在哪里？**
A: 默认输出到桌面 `MinerU_Output` 目录，可在界面中自定义。

---

## 🔧 技术栈

- **Python 3.8+** — 编程语言
- **Tkinter / ttk** — 图形界面框架
- **MinerU Open SDK** — 文档智能处理 API
- **Threading** — 后台异步处理

---

## 📄 License

本项目基于 [Apache License 2.0](LICENSE) 开源。

---

<div align="center">
  <sub>Built with ❤️ using <a href="https://github.com/opendatalab/MinerU">MinerU</a> | 
  <a href="https://github.com/B0Y-Studio/MinerU-Batch/issues">Report Issue</a></sub>
</div>
