---
name: word-latex-mathtype
description: "Convert LaTeX text in Word .docx files into MathType or native Word equation objects on Windows. Use when working with docx formula conversion, turning standalone or inline LaTeX into equations, cleaning stray display-$ characters, or safely handling Chinese \\text{} content. 在 Windows 上将 Word .docx 中的 LaTeX 文本转换为 MathType 或 Word 公式对象，用于 docx 公式转换、显示公式残留 $ 清理以及中文 \\text{} 安全处理。"
---

# Word LaTeX MathType

## 中文说明

### 概述

这个 skill 用于把 Word `.docx` 文档中的 LaTeX 公式文本转换成真正的公式对象。默认优先使用 MathType；如果 MathType 不可用或单条转换失败，则回退到 Word 原生公式。

它内置了两类专项修复：显示公式转换后残留字面量 `$` 的清理，以及 `\text{中文}` 场景下把中文拆成普通文本、数学部分保留为公式对象，避免 MathType 乱码。

### 何时使用

- 用户明确提到 Word、`.docx`、LaTeX、公式对象、MathType、Word 公式
- 需要把独立公式段落或高置信行内公式从纯文本变成可编辑公式
- 需要保留已有 OMath、OLE 或现成公式对象，不要破坏现有对象
- 需要批量处理文档，并生成转换日志方便回查

### 前置条件

- Windows 环境
- 已安装 Microsoft Word
- 建议安装 MathType；未安装时可回退为 Word 原生公式
- 执行覆盖写入前，尽量关闭目标文档，避免锁文件

### 工作流

1. 运行环境检查脚本，补齐 Python 和 `PyYAML`，并验证 Word/MathType。
2. 运行公式转换脚本，输入源 `.docx`，输出新文档和 `.json` 日志。
3. 抽样检查显示公式、分式、分段函数和带中文 `\text{}` 的公式。

### 常用命令

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_docx_latex_to_equations.ps1 `
  -SourcePath "C:\path\to\input.docx" `
  -TargetPath "C:\path\to\input_formula.docx" `
  -LogPath "C:\path\to\input_formula_log.json"
```

### 脚本

- `scripts/setup_environment.ps1`
  用于检查或安装 Python 3.12、`PyYAML`，并验证 Word 与 MathType。
- `scripts/convert_docx_latex_to_equations.ps1`
  用于复制源文档、转换公式对象、输出日志。

### 限制

- 当前版本仅支持 Windows
- 依赖 Word COM 自动化；没有 Word 就不能转换 `.docx` 公式对象
- MathType 是推荐项，不是强制项
- `\text{中文}` 的处理优先保证中文可读性，不强求保持为单一 MathType 对象

## English Guide

### Overview

This skill converts LaTeX-like formula text inside Word `.docx` files into real equation objects. It uses MathType first when available, then falls back to native Word equations when MathType is unavailable or a single formula fails.

It also removes stray `$` characters left by display-style conversion and splits `\text{CJK}` content into plain text plus equation objects when needed for readability.

### When To Use

- The task involves Word, `.docx`, LaTeX, equation objects, MathType, or Word equations
- Standalone or inline LaTeX should become editable equations
- Existing OMath, OLE, or inline objects should be preserved
- The workflow needs a JSON conversion log

### Prerequisites

- Windows
- Microsoft Word installed
- MathType recommended; native Word equations remain available as fallback
- Close the target output document before overwrite operations

### Workflow

1. Run the environment setup script to install or verify Python and `PyYAML`, then verify Word and MathType.
2. Run the conversion script with a source `.docx` file.
3. Review the output document and JSON log.

### Common Commands

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_docx_latex_to_equations.ps1 `
  -SourcePath "C:\path\to\input.docx" `
  -TargetPath "C:\path\to\input_formula.docx" `
  -LogPath "C:\path\to\input_formula_log.json"
```

### Scripts

- `scripts/setup_environment.ps1`
  Installs or verifies Python 3.12 and `PyYAML`, then checks for Word and MathType.
- `scripts/convert_docx_latex_to_equations.ps1`
  Copies the source document, converts formulas into equation objects, and writes the JSON log.

### Limitations

- Windows-only
- Word COM automation is required for `.docx` equation conversion
- MathType is optional but recommended
- For `\text{Chinese}` cases, readability is prioritized over keeping the entire expression as one MathType object
