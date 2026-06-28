# MarkdownPasteAddin v3.2 — Markdown 智能粘贴 Word 工具

将从 DeepSeek、网页、Markdown 文件复制的含**表格、流程图、数学公式、代码、图片**的内容，自动转换为 Word 原生格式；并对 Word 文档进行**标题、图题、表题、正文**格式统一。

## 新功能 (v3.2)

| 功能 | 说明 |
|------|------|
| **Chrome/Edge 扩展** | 一键从网页导出为 Word，支持离线转换 |
| **离线 JS 引擎** | `docx-builder.js` 浏览器内直接生成 .docx |
| **桥接服务** | 扩展可选连接本地 Python 服务获得完整功能 |
| **数学公式** | `$...$` / `$$...$$` → Word 原生 OMML 公式 |
| **代码高亮** | ```python 代码块 → 语法高亮着色 (Pygments) |
| **任务列表** | `- [ ]` / `- [x]` → 复选框列表 |
| **引用块** | `> 引用文本` → 带左边框的引用样式 |
| **嵌套列表** | 多级 `-` / `1.` 缩进 → 多级列表样式 |
| **分割线** | `---` / `***` → Word 水平分割线 |
| **目录生成** | 自动插入 TOC 域代码 (Word 中右键更新) |
| **图表自动编号** | 图1、图2、表1、表2 自动递增 |
| **报告格式预设** | 政府公文 / 学术论文 / 商务报告 |
| **页眉页脚** | 自动页码 / 文档标题 |
| **Word 模板** | `--template T.docx` 基于模板生成 |
| **批量转换** | `batch_convert.py *.md output/` |
| **文件监控** | `watch_convert.py` 文件变化自动转换 |
| **错误恢复** | 单个 chunk 失败不影响后续 (红色标记) |
| **GUI 桌面版** | `gui_app.py` tkinter 图形界面 |
| **DeepSeek API** | `deepseek_api.py` 直接提问/对话转 Word |

---

## 项目结构

```
MarkdownPasteAddin/
├── chrome-extension/             # Chrome/Edge 扩展
│   ├── manifest.json             # Manifest V3
│   ├── background.js             # Service Worker
│   ├── lib/docx-builder.js       # 离线 docx 生成引擎
│   ├── content/content_script.js # 页面注入 + 内容捕获
│   ├── popup/                    # 工具栏弹窗 + 侧边栏
│   └── options/                  # 设置页面
├── bridge_server.py              # 扩展桥接服务 (localhost:9876)
├── md2docx.py                    # 核心工具：剪贴板/文件 → Word (.docx)
├── format_docx.py                # 格式工具：统一 Word 文档的段落格式
├── batch_convert.py              # 批量转换工具
├── watch_convert.py              # 文件监控自动转换
├── deepseek_api.py               # DeepSeek API 集成
├── gui_app.py                    # GUI 桌面版
├── md2docx_lib/                  # 核心库
│   ├── __init__.py
│   ├── clipboard.py              # 剪贴板读取
│   ├── parser_markdown.py        # Markdown 解析器
│   ├── parser_html.py            # HTML 解析器
│   ├── parser_math.py            # LaTeX → OMML 转换器
│   ├── inline_processor.py       # 内联格式处理
│   ├── renderer_mermaid.py       # Mermaid 渲染
│   ├── renderer_code.py          # 代码语法高亮
│   ├── renderer_image.py         # 图片下载
│   ├── builder_docx.py           # Word 文档构建器
│   ├── builder_toc.py            # TOC 目录生成
│   ├── builder_numbering.py      # 图表自动编号
│   ├── report_standards.py       # 报告格式标准
│   ├── formatter.py              # 格式统一
│   └── template.py               # 模板支持
├── 粘贴转Word.bat                 # 一键快捷方式
├── 格式化Word文档.bat             # 格式统一快捷方式
├── GUI启动.bat                   # 图形界面快捷方式
├── 批量转换.bat                   # 批量转换快捷方式
├── DeepSeek提问转Word.bat         # API 提问快捷方式
├── 启动桥接服务.bat               # 桥接服务启动
├── requirements.txt              # Python 依赖
├── src/                          # C# VSTO 插件版
│   └── MarkdownPasteAddin/
└── tests/                        # 测试
```

---

## 环境依赖

| 依赖 | 用途 | 安装命令 |
|------|------|----------|
| Python 3.9+ | 脚本运行环境 | `winget install python` |
| python-docx | Word 文档读写 | `pip install python-docx` |
| requests | HTTP 请求 | `pip install requests` |
| pywin32 | Windows 剪贴板读取 | `pip install pywin32` |
| beautifulsoup4 | HTML 内容解析 | `pip install beautifulsoup4` |
| html5lib | HTML 解析器 | `pip install html5lib` |
| lxml | XML 处理 (OMML) | `pip install lxml` |
| Pygments | 代码语法高亮 | `pip install Pygments` |
| tqdm | 进度条 | `pip install tqdm` |
| openai | DeepSeek API 调用 | `pip install openai` |
| Node.js | Mermaid 本地渲染 | `winget install OpenJS.NodeJS` |
| mermaid-cli | 命令行渲染 Mermaid | `npm install -g @mermaid-js/mermaid-cli` |

一键安装：
```bash
pip install -r requirements.txt
```

---

## 工具一：Chrome/Edge 扩展

### 功能

一键从任意网页导出内容为 Word，无需安装 Python。

- **离线模式** — 安装即用，浏览器内直接生成 .docx
- **桥接模式** — 连接本地 Python 服务获得流程图、公式、报告排版等完整功能

### 安装

```
Chrome:  chrome://extensions/ → 开发者模式 → 加载已解压的扩展 → 选择 chrome-extension/
Edge:    edge://extensions/ → 开发人员模式 → 加载解压缩的扩展 → 选择 chrome-extension/
```

### 桥接服务（可选）

```bash
pip install -r requirements.txt
启动桥接服务.bat
# 或: python bridge_server.py
```

---

## 工具二：md2docx — 粘贴转 Word

### 功能

- **表格** — 蓝底表头 + 居中 + 对齐
- **Mermaid 流程图** — 本地/在线渲染为 PNG
- **数学公式** — `$x^2$` / `$$...$$` → Word 原生公式对象
- **代码高亮** — Pygments 语法着色 → Word 彩色字符
- **图片** — URL / data:URI → 下载嵌入
- **任务列表** — `- [ ]` → ☐ 复选框
- **引用块** — `>` → 左侧灰色缩进引用
- **分割线** — `---` → Word 水平线
- **标题/列表** — 对应 Word 样式

### 用法

```bash
# 从剪贴板转换
python md2docx.py output.docx

# 从 Markdown 文件转换
python md2docx.py input.md output.docx

# 带目录
python md2docx.py input.md output.docx --toc

# 带标题
python md2docx.py input.md output.docx --title "我的报告"

# 带模板
python md2docx.py input.md output.docx --template template.docx

# 转换后自动格式化
python md2docx.py input.md output.docx --format

# 设置图片最大宽度
python md2docx.py input.md output.docx --image-width 6.0
```

---

## 工具三：format_docx — Word 格式统一

### 格式预设

| 类型 | 字体 | 字号 | 加粗 | 对齐 | 其他 |
|------|------|------|------|------|------|
| 一级标题 | 黑体 | 16pt | 是 | 左 | 大纲级别 1 |
| 二级标题 | 黑体 | 14pt | 是 | 左 | 大纲级别 2 |
| 三级标题 | 黑体 | 12pt | 是 | 左 | 大纲级别 3 |
| 图题 | 宋体 | 10pt | 否 | 居中 | — |
| 表题 | 宋体 | 10pt | 是 | 居中 | — |
| 正文 | 宋体 | 12pt | 否 | 左 | 首行缩进 1.5倍行距 |

### 用法

```bash
python format_docx.py 文档.docx                   # 原地格式化
python format_docx.py 原.docx 新.docx             # 输出到新文件
python format_docx.py --active                    # 格式化当前 Word 文档
python format_docx.py --show                      # 显示格式预设
python format_docx.py --lang en 文档.docx         # 英文格式 (Times New Roman)
```

---

## 工具四：deepseek_api — DeepSeek 直连

### 用法

```bash
# 设置 API key
set DEEPSEEK_API_KEY=sk-xxxx

# 单次提问
python deepseek_api.py "解释量子纠缠" output.docx

# 交互对话模式
python deepseek_api.py --interactive

# 加载对话历史
python deepseek_api.py --conversation chat.json "继续解释" out.docx

# 保存对话记录
python deepseek_api.py "问题" out.docx --save-conversation chat.json
```

---

## 工具五：GUI 桌面版

```bash
python gui_app.py
```

提供图形界面：粘贴/编辑 Markdown → 预览解析结果 → 一键生成 Word。

---

## 工具六：批量转换 & 文件监控

```bash
# 批量转换目录
python batch_convert.py input_dir/ output_dir/

# 批量转换文件通配符
python batch_convert.py "docs/*.md" output_dir/

# 监控目录自动转换
python batch_convert.py --watch input_dir/ output_dir/

# 监控单文件
python watch_convert.py input.md output.docx
```

---

## 典型工作流

```
1. DeepSeek 提问 → 复制回答 (Ctrl+C)
2. 双击 粘贴转Word.bat → paste_result.docx (自动含目录+格式化)
3. 打开文件直接使用
```

或一键直达：

```
1. 双击 DeepSeek提问转Word.bat → 输入问题 → 自动生成 Word
```

---

## C# VSTO 插件版本

`src/` 目录下提供 Word VSTO Add-in 版本，安装后会在 Word 功能区增加"Markdown Paste"选项卡。

功能与 Python 版本基本对齐：表格、Mermaid、图片下载、代码块、任务列表、引用块、分割线、错误恢复。

---

## 自定义格式

编辑 `md2docx_lib/formatter.py` 中的 `FORMAT` 字典即可修改格式预设。

---

## 常见问题

**Q: 双击 bat 闪退？**
A: 检查 Python 路径，编辑 bat 中的 python.exe 路径。

**Q: 数学公式显示为 $...$？**
A: 需安装 `lxml` 库：`pip install lxml`。OMML 公式需要 Word 2019+ 支持。

**Q: 代码没有语法高亮？**
A: 需安装 Pygments：`pip install Pygments`。未安装时自动降级为无高亮等宽字体。

**Q: DeepSeek API 报错？**
A: 设置环境变量 `DEEPSEEK_API_KEY=sk-xxxx` 或使用 `--api-key` 参数。

**Q: 目录显示为域代码？**
A: 在 Word 中右键 TOC → 更新域 即可生成。Word 打开时自动提示。

**Q: Mermaid 流程图不显示？**
A: 安装 Node.js + `npm install -g @mermaid-js/mermaid-cli`，或确保联网（自动回退在线 API）。

---

## 支持项目 / Sponsorship

如果这个项目帮你省了时间，欢迎请作者喝杯咖啡。

| 方式 | 链接 / 二维码 |
|------|-------------|
| 爱发电 | [afdian.com/a/markdownpaste](https://afdian.com/a/markdownpaste)（待替换为你的链接） |
| GitHub Sponsors | [github.com/sponsors/chenacc1](https://github.com/sponsors/chenacc1) |
| 微信赞赏 | 将赞赏码图片保存为 `sponsor/wechat-qr.png` |

所有赞助者 ID 会记录在 `SPONSORS.md` 中（可选匿名）。

---

MIT License
