# word-latex-mathtype

## 这个 skill 做什么

`word-latex-mathtype` 用于把 Word `.docx` 文档中的 LaTeX 公式文本转换成真正的公式对象。

它的默认策略是：

- 优先转换为 MathType 对象
- MathType 不可用或单条公式失败时，回退为 Word 原生公式
- 自动清理显示公式转换后残留的 `$`
- 对包含中文 `\text{}` 的公式采用“数学对象 + 普通文本”的混排方式，优先保证中文可读性

这个 skill 适合处理已有 Word 报告、技术说明、实验文档中的 LaTeX 公式清洗和批量替换。

## 环境要求

- Windows
- Microsoft Word
- 建议安装 MathType
- PowerShell
- Python 3.12
- Python 包 `PyYAML`
- 如果希望脚本自动安装 Python，建议系统可用 `winget`

`setup_environment.ps1` 会自动检查或安装 Python 3.12，并补齐 `PyYAML`。它还会验证 Word 和 MathType 是否可用。

## 打包内容

```text
word-latex-mathtype/
├── SKILL.md
├── README.md
├── .gitignore
├── agents/
│   └── openai.yaml
└── scripts/
    ├── setup_environment.ps1
    └── convert_docx_latex_to_equations.ps1
```

## Codex 安装方式

推荐使用项目级安装。

1. 在你的项目根目录创建 `.agents/skills/word-latex-mathtype/`
2. 将这个 skill 整个目录复制进去
3. 最终结构应类似于：

```text
.agents/
└── skills/
    └── word-latex-mathtype/
        ├── SKILL.md
        ├── agents/
        └── scripts/
```

在 Codex 中有两种使用方式：

- 直接在任务里提到 Word、`.docx`、LaTeX、MathType、公式转换等需求，让 Codex 自动触发这个 skill
- 显式写出 `$word-latex-mathtype`，要求 Codex 使用该 skill

## Claude Code 安装方式

推荐使用项目级安装。

1. 在你的项目根目录创建 `.claude/skills/word-latex-mathtype/`
2. 将这个 skill 整个目录复制进去
3. 最终结构应类似于：

```text
.claude/
└── skills/
    └── word-latex-mathtype/
        ├── SKILL.md
        ├── agents/
        └── scripts/
```

在 Claude Code 中推荐直接用 slash 命令调用：

```text
/word-latex-mathtype
```

也可以直接描述任务，让 Claude 根据 `SKILL.md` 的描述自动匹配。

## 准备环境

在 skill 根目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup_environment.ps1
```

这个脚本会：

- 检测或安装 Python 3.12
- 检测并安装 `PyYAML`
- 验证 Word 是否存在
- 检查 MathType 是否存在

如果没有安装 MathType，脚本会给出警告，但转换流程仍可继续，后续会更多依赖 Word 原生公式回退。

## 使用说明

### 最简用法

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\convert_docx_latex_to_equations.ps1 `
  -SourcePath "C:\path\to\input.docx"
```

当只提供 `-SourcePath` 时，脚本会自动生成：

- `input_formula.docx`
- `input_formula_log.json`

### 指定输出文档和日志

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\convert_docx_latex_to_equations.ps1 `
  -SourcePath "C:\path\to\input.docx" `
  -TargetPath "C:\path\to\input_formula.docx" `
  -LogPath "C:\path\to\input_formula_log.json"
```

### 使用前建议

- 尽量关闭目标输出文档，避免覆盖时被 Word 锁定
- 如果源文档正在被打开，脚本通常仍可读取，但建议关闭不必要的占用窗口
- 转换后抽样检查显示公式、行内公式和中文混排公式

## 输出结果与日志说明

脚本会输出一个新的 `.docx` 文件和一个 `.json` 日志文件。

常见日志状态包括：

- `display-mathtype`
  独立公式成功转换为 MathType 对象
- `display-mathtype-cjk-split`
  包含中文的独立公式被拆成“公式对象 + 普通文本”混排
- `inline-mathtype`
  行内公式成功转换为 MathType 对象
- `word-fallback`
  MathType 失败后回退为 Word 原生公式成功
- `skipped-existing-object`
  该段已有现成对象，因此跳过
- `failed`
  MathType 和 Word 公式都失败，保留原文本

## 常见问题

### 1. 运行时提示目标文档被锁定

说明目标输出文件正在被 Word 或其他程序占用。关闭对应文档后重新运行即可。

### 2. 没有安装 MathType 可以用吗

可以。脚本会继续尝试转换，并在需要时回退到 Word 原生公式，只是最终结果中 MathType 对象会减少。

### 3. 为什么不是所有行内公式都被转换

这是有意为之。正文里经常混有函数名、英文标识符和代码风格文本，如果匹配太激进，误伤会明显增加。当前策略只转换高置信片段。

### 4. 为什么中文公式不是整个都变成一个 MathType 对象

因为整段送入 MathType 时，中文 `\text{}` 更容易出现乱码或显示不稳定。当前实现优先保证中文可读性，因此会拆成“数学对象 + 普通中文文本”。

## 已知限制

- 当前版本只支持 Windows
- 依赖 Word COM 自动化
- 对特别复杂或嵌套较深的 LaTeX 结构，仍可能需要人工复查
- 对行内公式采用保守识别，不追求“全部命中”
- 中文混排场景优先可读性，不保证始终保留为单一 MathType 对象
