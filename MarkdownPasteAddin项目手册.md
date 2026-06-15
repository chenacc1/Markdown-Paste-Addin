# MarkdownPasteAddin v3.2 项目手册

## 执行操作指南 · 架构说明 · 完整参考

---

# 第一部分：项目概述

## 1.1 这是什么？

**MarkdownPasteAddin** 是一个将 Markdown/HTML 富文本内容转换为 Microsoft Word 原生格式的工具集。主要解决以下痛点：

> 从 DeepSeek、ChatGPT、网页等地方复制了包含表格、流程图、数学公式、代码块的内容，粘贴到 Word 时格式丢失，需要手动重建。

本工具一键完成：
- Markdown 表格 → Word 原生表格（蓝底表头 + 自动对齐）
- Mermaid 流程图代码 → 渲染为 PNG 嵌入
- LaTeX 数学公式 → Word 原生 OMML 公式对象
- 代码块 → Pygments 语法高亮着色
- 图片 URL → 自动下载嵌入
- 任务列表、引用块、分割线、嵌套列表 → 对应 Word 样式
- 图表自动编号（图1、图2、表1、表2）
- 目录自动生成（TOC 域代码）

## 1.2 核心能力矩阵

```
输入源              解析引擎          输出
─────────────────────────────────────────────
剪贴板 HTML     →  parser_html    ┐
剪贴板 Text     →                 │
Markdown 文件   →  parser_md      ├→ builder_docx → .docx
DeepSeek API    →  (同 md 解析)   │
手动输入 (GUI)  →                 ┘
                                   │
                                   └→ format_docx → 格式统一
```

## 1.3 版本历史

| 版本 | 主要能力 |
|------|----------|
| v1.0 | 表格、Mermaid 流程图、图片、基础文本 |
| v2.0 | 数学公式、代码高亮、任务列表、引用块、分割线、嵌套列表、TOC、图表编号、批量转换、文件监控、GUI、DeepSeek API |
| v3.0 | Chrome/Edge 扩展、桥接服务、报告格式预设、封面页、页眉页脚、HTML解析器、中文格式 |
| v3.2 | 离线 JS docx 引擎 (docx-builder.js)、LaTeX 引擎重写、扩展离线优先架构、GUI 界面重构 |

---

# 第二部分：环境安装

## 2.1 一键安装

```bash
pip install -r requirements.txt
```

## 2.2 依赖清单

| 包名 | 最低版本 | 用途 | 必须 |
|------|---------|------|------|
| python-docx | 0.8.11 | Word 文档读写 | 是 |
| requests | 2.28.0 | HTTP 请求（在线渲染/图片下载）| 是 |
| pywin32 | 305 | Windows 剪贴板读取 | 是 |
| beautifulsoup4 | 4.12.0 | HTML 内容解析 | 是 |
| html5lib | 1.1 | HTML5 解析器 | 是 |
| lxml | 4.9.0 | XML 处理（OMML 公式）| 是 |
| Pygments | 2.15.0 | 代码语法高亮 | 推荐 |
| tqdm | 4.65.0 | 进度条显示 | 推荐 |
| openai | 1.0.0 | DeepSeek API 调用 | 可选 |
| Node.js + mermaid-cli | — | Mermaid 本地渲染 | 可选 |

## 2.3 Mermaid 本地渲染（可选）

```bash
winget install OpenJS.NodeJS          # 安装 Node.js
npm install -g @mermaid-js/mermaid-cli # 安装 mermaid-cli
```

> 不安装时自动回退到 mermaid.ink 在线 API（需联网）。

## 2.4 验证安装

```bash
python -c "from md2docx_lib import parse_markdown; print('OK')"
```

---

# 第三部分：快速上手（5 分钟）

## 场景 1：从 DeepSeek 复制回答转 Word

```
第一步：在 DeepSeek 中 Ctrl+C 复制内容
第二步：双击 「粘贴转Word.bat」
第三步：自动打开 paste_result.docx，含目录和格式化
```

## 场景 2：从 Markdown 文件转 Word

```bash
python md2docx.py 报告.md 报告.docx --toc --format --title "季度报告"
```

## 场景 3：直接向 DeepSeek 提问并生成 Word

```bash
# 先设置 API Key
set DEEPSEEK_API_KEY=sk-xxxx

# 提问
python deepseek_api.py "解释量子计算的基本原理" output.docx --toc --format
```

## 场景 4：批量转换整个文件夹

```bash
python batch_convert.py ./docs/ ./output/
```

## 场景 5：GUI 图形界面

```bash
python gui_app.py
```

---

# 第四部分：工具详解

## 4.1 md2docx.py — 核心转换工具

### 命令格式

```
python md2docx.py [输入] [输出] [选项]
```

### 参数说明

| 参数 | 说明 | 示例 |
|------|------|------|
| 位置参数1 | 输出路径（剪贴板模式）或 输入 .md 文件 | `output.docx` / `report.md` |
| 位置参数2 | 输出路径（文件模式）| `output.docx` |
| `-c` / `--clipboard` | 强制从剪贴板读取 | `-c output.docx` |
| `--toc` | 在文档开头添加目录 | `--toc` |
| `--title "标题"` | 设置文档标题 | `--title "年度报告"` |
| `--template T.docx` | 使用 Word 模板 | `--template template.docx` |
| `--format` | 转换后自动格式化 | `--format` |
| `--no-captions` | 禁用图表自动编号 | `--no-captions` |
| `--image-width 6.0` | 图片最大宽度（英寸）| `--image-width 6.0` |

### 使用示例

```bash
# 剪贴板 → Word（最常用）
python md2docx.py output.docx

# 文件 → Word
python md2docx.py input.md output.docx

# 文件 → Word（含目录+标题+格式化）
python md2docx.py input.md output.docx --toc --title "我的文档" --format

# 基于模板生成
python md2docx.py input.md output.docx --template company_template.docx
```

### 输入格式优先级

1. 剪贴板 HTML（最丰富，保留表格/图片/格式）
2. 剪贴板纯文本
3. .md 文件（UTF-8 编码）

### 解析流程

```
原始内容
  │
  ├─ 判断类型 ──→ HTML 格式 ──→ parser_html.parse_html()
  │                              ├─ BeautifulSoup 解析
  │                              ├─ 提取 <table> → 表格 chunk
  │                              ├─ 提取 <img>   → 图片 chunk
  │                              ├─ 提取 <pre>   → 代码 chunk
  │                              ├─ 提取 <h1-h6> → 标题 chunk
  │                              └─ 提取 <ul>/<ol> → 列表
  │
  └─ 判断类型 ──→ 文本格式 ──→ parser_markdown.parse_markdown()
                                 ├─ 正则匹配表格 (|...|)
                                 ├─ 正则匹配代码块 (```)
                                 ├─ 正则匹配 Mermaid (```mermaid)
                                 ├─ 正则匹配标题 (#)
                                 ├─ 正则匹配任务列表 (- [ ])
                                 ├─ 正则匹配引用块 (> )
                                 ├─ 正则匹配分割线 (---)
                                 ├─ 正则匹配图片 (![]())
                                 └─ 剩余 → 文本块
        │
        ▼
  builder_docx.DocumentBuilder.build()
        │
        ├─ 表格 → Word Table (蓝底表头)
        ├─ Mermaid → mmdc 本地渲染 / mermaid.ink 在线 → PNG
        ├─ 图片 → HTTP 下载 → 嵌入
        ├─ 代码 → Pygments 着色 → 有色 Consolas 字体
        ├─ 数学 → LaTeX 递归解析 → OMML XML
        ├─ 任务列表 → ☐/☒ 符号
        ├─ 引用块 → 左缩进 + 灰色背景
        ├─ 分割线 → 段落底部边框
        └─ 文本 → 标题/列表/正文对应 Word 样式
        │
        ▼
  .docx 文件
```

---

## 4.2 format_docx.py — 格式统一工具

### 命令格式

```
python format_docx.py [文档.docx] [输出.docx] [选项]
```

### 参数说明

| 参数 | 说明 |
|------|------|
| 位置参数1 | 输入 .docx 文件 |
| 位置参数2 | 输出文件（默认覆盖原文件）|
| `--active` / `-a` | 格式化当前 Word 中打开的文档 |
| `--show` / `-s` | 显示当前格式预设 |
| `--lang cn` | 中文格式（宋体/黑体，默认）|
| `--lang en` | 英文格式（Times New Roman/Arial）|
| `--list-presets` | 列出所有预设名称 |

### 识别规则

```
段落文本 ──→ 已有 Heading 1 样式？  ──→ 一级标题
         │
         ├─→ 已有 Heading 2 样式？  ──→ 二级标题
         │
         ├─→ 已有 Heading 3 样式？  ──→ 三级标题
         │
         ├─→ 以 图1/Figure 1 开头？  ──→ 图题
         │
         ├─→ 以 表1/Table 1 开头？  ──→ 表题
         │
         ├─→ 第X章/一、/Chapter ？  ──→ 一级标题
         │
         ├─→ （一）/(1)/数字编号？  ──→ 二级标题
         │
         ├─→ 短文本+加粗+大字号？  ──→ 按字号判定级别
         │
         └─→ 其他                    ──→ 正文
```

### 默认格式预设

| 类型 | 中文字体 | 英文字体 | 字号 | 加粗 | 对齐 | 行距 | 段前 | 段后 | 首行缩进 | 大纲级别 |
|------|---------|---------|------|------|------|------|------|------|----------|----------|
| 一级标题 | 黑体 | Arial | 16pt | 是 | 左 | 1.5 | 18pt | 12pt | — | 1 |
| 二级标题 | 黑体 | Arial | 14pt | 是 | 左 | 1.5 | 14pt | 8pt | — | 2 |
| 三级标题 | 黑体 | Arial | 12pt | 是 | 左 | 1.5 | 10pt | 6pt | — | 3 |
| 图题 | 宋体 | Times NR | 10pt | 否 | 居中 | 1.5 | 6pt | 12pt | — | — |
| 表题 | 宋体 | Times NR | 10pt | 是 | 居中 | 1.5 | 12pt | 6pt | — | — |
| 正文 | 宋体 | Times NR | 12pt | 否 | 左 | 1.5 | 0 | 0 | 2字符 | — |

### 自定义格式

编辑 `md2docx_lib/formatter.py` 中的 `FORMAT` 字典：

```python
FORMAT = {
    "一级标题": {
        "font_name": "黑体",           # 修改字体
        "font_size": Pt(16),           # 修改字号
        "bold": True,                  # 修改粗体
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,  # 修改对齐
        "space_before": Pt(18),        # 修改段前
        "space_after": Pt(12),         # 修改段后
        "line_spacing": 1.5,           # 修改行距
        "outline_level": 1,            # 修改大纲级别
    },
    # ... 以此类推
}
```

---

## 4.3 deepseek_api.py — DeepSeek API 集成

### 前提条件

1. 安装 openai 包：`pip install openai`
2. 获取 DeepSeek API Key
3. 设置环境变量：`set DEEPSEEK_API_KEY=sk-xxxx`

### 命令格式

```
python deepseek_api.py "你的问题" [输出.docx] [选项]
```

### 参数说明

| 参数 | 说明 |
|------|------|
| 位置参数1 | 向 DeepSeek 提出的问题 |
| 位置参数2 | 输出 .docx 路径 |
| `--interactive` / `-i` | 交互对话模式 |
| `--model` / `-m` | 模型名称（默认 `deepseek-chat`）|
| `--api-key` / `-k` | API Key（不设环境变量时使用）|
| `--conversation` / `-c` | 加载历史对话 JSON |
| `--save-conversation` / `-s` | 保存对话到 JSON |

### 交互模式命令

```
You: 你的问题               ← 向 DeepSeek 提问
You: /save                  ← 导出对话到 Word
You: /exit                  ← 退出并自动保存
```

### DeepSeek 系统提示

工具自动要求 DeepSeek 输出结构化的 Markdown 回答，包括：
- 使用 `#` 标题层级
- 使用表格展示数据
- 使用 mermaid 代码块绘制流程图
- 使用 LaTeX 公式
- 使用任务列表

---

## 4.4 gui_app.py — GUI 图形界面

### 功能布局

```
┌──────────────────────────────────────────────┐
│  [粘贴] [打开文件] [清空] │ [转换] [预览]      │  工具栏
├──────────────────────────────────────────────┤
│                                              │
│          Markdown 编辑区                      │  主编辑
│          (等宽字体，滚动文本)                   │
│                                              │
├──────────────────────────────────────────────┤
│          解析结果预览                          │  结果预览
│  文本: 5  表格: 2  流程图: 1  代码: 1        │
│  任务列表: 1  引用: 1  共 11 个块            │
├──────────────────────────────────────────────┤
│  输出: [output.docx] [浏览] ☑编号 ☐目录 ☑格式化  图片宽:[5.5] │  设置
│  状态: 就绪                                    │  状态栏
└──────────────────────────────────────────────┘
```

### 使用流程

1. 点击「从剪贴板粘贴」或「打开 Markdown 文件」
2. 在编辑区查看/修改内容
3. 点击「预览」查看解析结果
4. 设置输出路径和选项
5. 点击「转换为 Word」

---

## 4.5 batch_convert.py — 批量转换

```bash
# 转换目录下所有 .md 文件
python batch_convert.py input_dir/ output_dir/

# 转换匹配的文件
python batch_convert.py "reports/*.md" output_dir/

# 监控目录，文件变化自动转换
python batch_convert.py --watch input_dir/ output_dir/

# 自定义轮询间隔
python batch_convert.py --watch input_dir/ output_dir/ --interval 5
```

## 4.6 watch_convert.py — 单文件监控

```bash
# 监控 report.md，变化时自动重新生成
python watch_convert.py report.md report.docx

# 自定义检查间隔
python watch_convert.py report.md report.docx --interval 3
```

---

# 第五部分：架构详解

## 5.1 核心库模块图

```
                  md2docx_lib (核心库)
                       │
        ┌──────────────┼──────────────┐
        │              │              │
   ┌────▼────┐   ┌────▼────┐   ┌────▼────┐
   │ Parser  │   │Renderer │   │ Builder │
   │  Layer  │   │  Layer  │   │  Layer  │
   └────┬────┘   └────┬────┘   └────┬────┘
        │              │              │
        │              │              │
  ┌─────┴──────┐  ┌────┴─────┐  ┌────┴──────┐
  │parser_md   │  │renderer  │  │builder_   │
  │parser_html │  │_mermaid  │  │docx       │
  │parser_math │  │_code     │  │builder_   │
  │            │  │_image    │  │toc        │
  │clipboard   │  │          │  │_numbering │
  └────────────┘  └──────────┘  │formatter  │
                                │template   │
                                └───────────┘
```

## 5.2 数据流

```
外部输入                  解析                  中间表示            构建             输出
───────────────────────────────────────────────────────────────────────────────────────
Clipboard ──→ clipboard.py ──→ list[dict] ──→ builder_docx.py ──→ .docx
DeepSeek  ──→ deepseek_api ──→  chunks     DocumentBuilder
.md file  ──→ parse_md()   ──→           build(chunks, path)
HTML      ──→ parse_html() ──→
```

### Chunk 数据结构（中间表示）

```python
chunks = [
    # 标题
    {"type": "heading",    "level": 1, "text": "Title"},

    # 文本
    {"type": "text",       "text": "Paragraph text\nwith newlines"},

    # 表格
    {"type": "table",      "headers": ["A","B"], "alignments": ["left","center"],
                           "rows": [["1","2"], ["3","4"]]},

    # 代码块
    {"type": "code",       "language": "python", "code": "def foo():\n    pass"},

    # 流程图
    {"type": "mermaid",    "code": "graph LR\n  A-->B"},

    # 数学公式
    {"type": "math",       "latex": "x^2 + y^2 = z^2", "display": True},

    # 图片
    {"type": "image",      "src": "https://...", "alt": "Description"},

    # 任务列表
    {"type": "task_list",  "items": [{"checked": False, "text": "Do X"}, ...]},

    # 引用块
    {"type": "blockquote", "text": "Quoted text"},

    # 分割线
    {"type": "hr"},
]
```

## 5.3 数学公式引擎 (parser_math.py)

### 支持的 LaTeX 语法

| 类别 | 支持的命令 |
|------|-----------|
| 分式 | `\frac{num}{den}` |
| 根号 | `\sqrt[n]{x}` |
| 上下标 | `x^{sup}`, `x_{sub}`, `x_{sub}^{sup}` |
| 希腊字母 | `\alpha` ~ `\omega`, `\Gamma` ~ `\Omega` |
| 运算符 | `\times \div \pm \cdot \leq \geq \neq \approx` |
| 箭头 | `\rightarrow \leftarrow \Rightarrow \Leftarrow` |
| 大型算符 | `\sum \prod \int \oint` (含上下限) |
| 重音 | `\bar \vec \hat \tilde \dot \ddot` |
| 分隔符 | `\left( \right)` `\left[ \right]` |
| 矩阵 | `\begin{matrix} ... \end{matrix}` (及 p/b/v 变体) |
| 特殊符号 | `\infty \partial \nabla \forall \exists \cdots \ldots` |
| 文本 | `\text{...}` `\mbox{...}` |

### 渲染流程

```
LaTeX 字符串
    │
    ▼
_LaTeXParser (递归下降)
    │  tokenize + parse
    ▼
AST (dict tree)
    │  _omml_from_tree()
    ▼
OMML XML 字符串
    │  插入 python-docx run._r
    ▼
Word 原生公式对象
```

---

# 第六部分：常见场景操作手册

## 6.1 日常使用：DeepSeek → Word

```bash
# 方式一：双击 bat（最简单）
1. DeepSeek 中 Ctrl+C 复制
2. 双击 「粘贴转Word.bat」
3. Word 打开，文档已含目录和格式化

# 方式二：命令行（可定制）
python md2docx.py result.docx --toc --format --title "DeepSeek 回答"

# 方式三：GUI（可视化）
python gui_app.py
```

## 6.2 处理已有 Markdown 文件

```bash
# 单个文件
python md2docx.py report.md report.docx --toc --format --title "报告标题"

# 整个文件夹
python batch_convert.py ./markdown_files/ ./word_output/

# 文件变化自动转（适合边写边看）
python watch_convert.py draft.md draft.docx
```

## 6.3 格式化现有 Word 文档

```bash
# 命令行格式化
python format_docx.py document.docx

# 在 Word 中直接格式化
python format_docx.py --active

# 查看当前格式预设
python format_docx.py --show
```

## 6.4 基于公司模板生成文档

```bash
# 1. 准备模板（在要插入内容的位置写 {{CONTENT}}）
# 2. 转换并指定模板
python md2docx.py content.md output.docx --template company_template.docx --format
```

## 6.5 批量处理 + 格式化管线

```bash
# 一步完成：转换所有 .md → .docx → 格式统一
for f in *.md; do
    python md2docx.py "$f" "${f%.md}.docx" --format
done

# 或使用批处理工具
python batch_convert.py ./input/ ./output/
for f in ./output/*.docx; do
    python format_docx.py "$f"
done
```

## 6.6 DeepSeek 交互写作

```bash
python deepseek_api.py --interactive

# 对话中：
# You: 请写一份关于 AI 发展趋势的报告大纲
# You: 展开第二点，加入具体数据和表格
# You: 把这些内容整理成一个流程图
# You: /save        ← 导出为 Word
# You: /exit        ← 退出
```

---

# 第七部分：故障排除

## 7.1 双击 bat 闪退

**现象**：双击 `.bat` 文件后窗口一闪而过

**解决**：
1. 检查 Python 安装路径：`where python`
2. 编辑 bat 文件，将 `D:\anaconda3\python.exe` 改为你的 Python 路径
3. 或在 bat 文件末尾加 `pause` 查看错误信息

## 7.2 剪贴板读取失败

**现象**：提示 "clipboard is empty"

**解决**：
1. 确认已复制内容（先 Ctrl+C）
2. 安装 pywin32：`pip install pywin32`
3. 某些应用（如 WPS）使用的剪贴板格式可能不兼容

## 7.3 Mermaid 流程图不显示

**现象**：显示为代码而不是图

**解决**：
1. 检查 Markdown 中是否使用了 ` ```mermaid ` 标记
2. 安装 Node.js：`winget install OpenJS.NodeJS`
3. 安装 mermaid-cli：`npm install -g @mermaid-js/mermaid-cli`
4. 确保网络畅通（自动回退在线渲染）

## 7.4 数学公式显示为原始 LaTeX

**现象**：看到 `$x^2$` 而不是公式

**解决**：
1. 安装 lxml：`pip install lxml`
2. 使用块级公式 `$$x^2$$`（单独一行）
3. 需要 Word 2019 或更新版本（支持 OMML）

## 7.5 代码没有语法高亮

**现象**：代码是单色等宽字体，没有颜色

**解决**：
1. 安装 Pygments：`pip install Pygments`
2. 确保代码块有语言标识：` ```python` 而不是 ` ``` `

## 7.6 目录显示为 `[ 请在 Word 中右键此处... ]`

**现象**：目录位置显示灰色提示文字

**解决**：
1. 在 Word 中右键该区域 → 选择「更新域」
2. 这是 Word 域代码的正常行为，不是错误

## 7.7 DeepSeek API 报错

**现象**：`Error: DEEPSEEK_API_KEY not set`

**解决**：
```bash
# 设置环境变量（当前会话）
set DEEPSEEK_API_KEY=sk-xxxx

# 或写在 bat 文件中
# 或用命令行参数
python deepseek_api.py --api-key sk-xxxx "问题" out.docx
```

---

# 第八部分：C# VSTO 插件版本

## 8.1 编译要求

| 组件 | 版本 |
|------|------|
| Visual Studio | 2022+ |
| Office/SharePoint 工作负载 | 已安装 |
| .NET Framework | 4.7.2 |
| Microsoft Office | 2019+ |

## 8.2 构建步骤

```
1. 用 Visual Studio 打开 MarkdownPasteAddin.sln
2. NuGet 还原：右键解决方案 → 还原 NuGet 包
3. 按 F5 调试运行（自动启动 Word 并加载插件）
4. 发布：生成 → 发布 → 选择发布位置
```

## 8.3 与 Python 版本功能对比

| 功能 | Python | C# VSTO |
|------|--------|---------|
| Markdown 表格 | 是 | 是 |
| HTML 表格 | 是 | 否（仅 text 剪贴板）|
| Mermaid 流程图 | 是（本地+在线）| 是（仅本地）|
| 图片下载嵌入 | 是 | 是 |
| 数学公式 | 是（OMML）| 否 |
| 代码高亮 | 是（Pygments）| 否（等宽字体）|
| 任务列表 | 是 | 是 |
| 引用块 | 是 | 是 |
| 分割线 | 是 | 是 |
| 图表编号 | 是 | 是 |
| TOC 目录 | 是 | 否 |
| 格式统一 | 独立工具 | 否 |
| 批量转换 | 是 | 否 |
| GUI 界面 | tkinter | Word 内嵌按钮 |
| DeepSeek API | 是 | 否 |

> **推荐**：日常使用 Python 版本（功能最全），C# 版本适合需要 Word 内嵌体验的场景。

---

# 第九部分：开发指南

## 9.1 添加新 Chunk 类型

1. 在 `parser_markdown.py` 中添加检测逻辑
2. 在 `builder_docx.py` 的 `_process_chunk()` 中添加处理分支
3. 在 C# 版本的 Models 中添加对应类

## 9.2 添加新的格式预设

编辑 `md2docx_lib/formatter.py` 的 `FORMAT` 字典：

```python
FORMAT["代码注释"] = {
    "font_name": "Consolas",
    "font_size": Pt(9),
    "bold": False,
    "color": RGBColor(0x99, 0x99, 0x99),
    "alignment": WD_ALIGN_PARAGRAPH.LEFT,
    "space_before": Pt(0),
    "space_after": Pt(0),
    "line_spacing": 1.0,
}
```

然后添加对应的检测规则到 `detect_paragraph_type()`。

## 9.3 扩展数学公式支持

在 `parser_math.py` 的 `_COMMANDS` 字典中添加新符号，或在 `_omml_command()` 方法中添加新的命令处理分支。

## 9.4 运行测试

```bash
# Python 功能测试
python -c "
from md2docx_lib import parse_markdown
md = open('tests/MarkdownPasteAddin.Tests/TestData/sample-mixed.md').read()
chunks = parse_markdown(md)
print(f'Parsed {len(chunks)} chunks')
"

# C# 测试（在 Visual Studio 中）
# 测试 → 运行所有测试
```

---

# 第十部分：快速参考卡片

```
┌─────────────────────────────────────────────────────────┐
│              MarkdownPasteAddin v3.2                    │
│                   快速参考                               │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  粘贴转 Word:    python md2docx.py output.docx           │
│  文件转 Word:    python md2docx.py in.md out.docx        │
│  格式化 Word:    python format_docx.py doc.docx          │
│  批量转换:       python batch_convert.py docs/ out/      │
│  监控转换:       python watch_convert.py in.md out.docx  │
│  DeepSeek 提问:  python deepseek_api.py "问题" out.docx  │
│  GUI 界面:       python gui_app.py                       │
│                                                         │
│  常用选项:                                               │
│    --toc           添加目录                              │
│    --format        转换后格式化                          │
│    --title "T"     设置文档标题                          │
│    --template T    使用模板                              │
│    --lang en       英文格式                              │
│                                                         │
│  快捷方式（双击即用）:                                    │
│    粘贴转Word.bat       剪贴板 → Word                    │
│    格式化Word文档.bat   拖入 .docx 格式化                │
│    GUI启动.bat          打开图形界面                     │
│    批量转换.bat          批量 .md → .docx                │
│    DeepSeek提问转Word.bat  API 直连提问                  │
│                                                         │
│  核心库: from md2docx_lib import *                      │
│  依赖安装: pip install -r requirements.txt              │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

---

*文档版本: 3.2 | 最后更新: 2026-06-15 | 项目许可: MIT*
