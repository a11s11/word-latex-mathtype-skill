---
name: word-latex-mathtype-skill
description: "Convert LaTeX in Word .docx files into MathType or Word equations on Windows, including display-$ cleanup and CJK-safe handling. 在 Windows 上将 Word .docx 中的 LaTeX 转换为 MathType 或 Word 公式对象，并处理显示公式残留 $ 与中文 \\text 乱码。"
---

# Word Latex Mathtype Skill

## 中文说明

### 概述

这个 skill 用于把 Word `.docx` 文档中的 LaTeX 公式文本转换成真正的公式对象。默认优先使用 MathType；如果 MathType 不可用或单条转换失败，则回退到 Word 原生公式。

它内置了两类专项修复：显示公式转换后残留字面量 `$` 的清理，以及 `\text{中文}` 场景下把中文拆成普通文本、数学部分保留为公式对象，避免 MathType 乱码。

### 何时使用

- 用户明确提到 Word、`.docx`、LaTeX、公式对象、MathType、Word 公式。
- 需要把独立公式段落或高置信行内公式从纯文本变成可编辑公式。
- 需要保留已有 OMath、OLE 或现成公式对象，不要破坏现有对象。
- 需要批量处理文档，并生成转换日志方便回查。

### 前置条件

- Windows 环境。
- 已安装 Microsoft Word。
- 建议安装 MathType；未安装时可回退为 Word 原生公式。
- 执行覆盖写入前，尽量关闭源文档和目标文档，避免锁文件。
- 发布到 GitHub 时只使用浏览器登录，不使用密码登录。

### 工作流

1. 先运行环境检查与安装脚本，补齐 Python、Git、GitHub CLI，并验证 Word/MathType。
2. 运行公式转换脚本，输入源 `.docx`，输出新文档和 `.json` 日志。
3. 抽样检查显示公式、分式、分段函数、带中文 `\text{}` 的公式。
4. 如果要分享 skill，再运行 GitHub 发布脚本进行浏览器登录、建仓库、提交和推送。

### 常用命令

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_docx_latex_to_equations.ps1 `
  -SourcePath "I:\path\input.docx" `
  -TargetPath "I:\path\input_formula.docx" `
  -LogPath "I:\path\input_formula_log.json"
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/publish_to_github.ps1 `
  -RepoRoot "I:\path\word-latex-mathtype-skill" `
  -RepoName "word-latex-mathtype-skill"
```

### 转换行为

- 独立公式：优先走 MathType 显示公式路径，自动清理新建对象两侧残留的字面量 `$`。
- 行内公式：只处理高置信命中的片段，避免误伤函数名、标识符和代码样式文本。
- 含中文公式：检测 `\text{...中文...}` 后拆成“公式对象 + 普通文本”混排。
- 已存在对象：跳过已有 OMath、InlineShape、OLE 等对象。
- 失败回退：MathType 失败时尝试 Word OMath；两者都失败则保留原 LaTeX 文本并写入日志。

### 输出

- 新的 `.docx` 输出文件。
- 与输出文件配套的 `.json` 日志。
- 日志状态包含 `display-mathtype`、`display-mathtype-cjk-split`、`inline-mathtype`、`word-fallback`、`failed` 等。

### 脚本

- `scripts/setup_environment.ps1`
  用于检查或安装 Python 3.12、GitHub CLI、Git，并验证 Word 与 MathType。
- `scripts/convert_docx_latex_to_equations.ps1`
  用于复制源文档、转换公式对象、输出日志。
- `scripts/publish_to_github.ps1`
  用于浏览器登录 GitHub、初始化仓库、创建远端并推送。

### 限制

- 当前版本仅支持 Windows。
- 依赖 Word COM 自动化；没有 Word 就不能转换 `.docx` 公式对象。
- MathType 是“推荐但非强制”；没有 MathType 时只能生成 Word 原生公式。
- `\text{中文}` 的处理优先保证中文可读性，不强求保持为单一 MathType 对象。

## English Guide

### Overview

This skill converts LaTeX-like formula text inside Word `.docx` files into real equation objects. It uses MathType first when available, then falls back to native Word equations when MathType is unavailable or a single formula fails.

It also includes two focused fixes: removing stray literal `$` characters left around newly created display equations, and splitting `\text{CJK}` content into plain text plus equation objects so Chinese text stays readable.

### When To Use

- The user mentions Word, `.docx`, LaTeX, equation objects, MathType, or Word equations.
- The task is to turn standalone formula paragraphs or high-confidence inline formulas into editable equations.
- Existing equation objects, OLE objects, or Word OMath objects should be preserved.
- The workflow needs a JSON conversion log for review or troubleshooting.

### Prerequisites

- Windows.
- Microsoft Word installed.
- MathType recommended; native Word equations remain available as fallback.
- Source and target documents should be closed before overwrite operations.
- GitHub publishing uses browser login only, never password auth.

### Workflow

1. Run the environment setup script to install or verify Python, Git, and GitHub CLI, then verify Word and MathType.
2. Run the conversion script with a source `.docx` file and let it produce a copied output document plus a `.json` log.
3. Spot-check display formulas, fractions, cases, and formulas that contain `\text{}` with Chinese.
4. If the skill should be shared, run the GitHub publish script to authenticate in the browser, create the repo, commit, and push.

### Common Commands

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup_environment.ps1
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/convert_docx_latex_to_equations.ps1 `
  -SourcePath "I:\path\input.docx" `
  -TargetPath "I:\path\input_formula.docx" `
  -LogPath "I:\path\input_formula_log.json"
```

```powershell
powershell -ExecutionPolicy Bypass -File scripts/publish_to_github.ps1 `
  -RepoRoot "I:\path\word-latex-mathtype-skill" `
  -RepoName "word-latex-mathtype-skill"
```

### Conversion Behavior

- Display formulas: MathType is preferred for display-style conversion, and stray literal `$` delimiters are removed around newly created objects.
- Inline formulas: only high-confidence matches are converted to reduce false positives.
- Formulas containing Chinese in `\text{...}` are split into mixed equation/text content for readability.
- Existing OMath, InlineShape, and OLE objects are skipped.
- If MathType fails, the script tries native Word equations; if both fail, the original LaTeX text is preserved and logged.

### Outputs

- A new `.docx` output file.
- A companion `.json` log file.
- Log statuses such as `display-mathtype`, `display-mathtype-cjk-split`, `inline-mathtype`, `word-fallback`, and `failed`.

### Scripts

- `scripts/setup_environment.ps1`
  Installs or verifies Python 3.12, GitHub CLI, and Git, then checks for Word and MathType.
- `scripts/convert_docx_latex_to_equations.ps1`
  Copies the source document, converts formulas into equation objects, and writes the JSON log.
- `scripts/publish_to_github.ps1`
  Opens browser-based GitHub auth, initializes the repo, creates the remote repository, and pushes the skill.

### Limitations

- Version 1 is Windows-only.
- Word COM automation is required for `.docx` equation conversion.
- MathType is optional but recommended; without it, only native Word equations can be produced.
- For `\text{Chinese}` cases, readability is prioritized over keeping the entire expression as one MathType object.
