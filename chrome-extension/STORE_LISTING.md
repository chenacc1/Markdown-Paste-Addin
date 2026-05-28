# Microsoft Edge Add-ons Store 上架指南

## 一、准备工作

### 1.1 开发者注册

1. 访问 [Microsoft Partner Center](https://partner.microsoft.com/)
2. 使用 Microsoft 账户登录（个人账户即可，不需要公司）
3. 进入「Windows & Office」→「Edge Add-ons」
4. 填写开发者信息，同意协议
5. 注册免费

### 1.2 准备物料

| 物料 | 尺寸 | 格式 | 说明 |
|------|------|------|------|
| 扩展图标 | 128x128 | PNG | manifest.json 中的 icon128 |
| 商店宣传图(小) | 300x300 | PNG | 搜索列表中的缩略图 |
| 商店宣传图(中) | 440x280 | PNG | 扩展详情页顶部 |
| 商店宣传图(大) | 1400x560 | PNG | 精选推荐使用 |
| 截图 | 1280x800 | PNG/JPG | 至少 1 张，最多 10 张 |
| 隐私政策 | URL | HTML | 如果有外部通信需要 |

### 1.3 打包扩展

```bash
# 在 chrome-extension/ 目录下打包
cd chrome-extension
zip -r ../store-submission/edge-addons.zip . -x "*.md" "generate_icons.py" "generate_store_assets.py" ".git/*"
```

或使用 Python 打包脚本（推荐）：

```bash
cd D:\claude-project\p8\MarkdownPasteAddin
python _repack.py
```

## 二、商店信息填写模板

### 基本信息

```
名称（中文）: MarkdownPasteAddin - 智能粘贴Word
名称（英文）: MarkdownPasteAddin - Smart Paste to Word
简短描述（45字）: 将网页Markdown/表格/流程图/公式一键导出为Word
详细描述: 见下方「商店描述」文本
开发者: [你的名字/组织名称]
网站: https://github.com/user/MarkdownPasteAddin
隐私政策URL: [如果收集数据才需要，本扩展不收集任何数据]
支持邮箱: [你的邮箱]
```

### 分类

```
主分类: 生产力
次分类: 开发者工具
```

### 商店描述（中文）

```
# MarkdownPasteAddin v3.0

一键将网页中的 Markdown 内容、表格、流程图、数学公式、代码块转换为 Microsoft Word 原生格式 (.docx)。

## 主要功能

📝 **智能内容捕获**
- 自动识别网页中的 Markdown / 代码块 / 表格 / Mermaid 流程图
- 支持手动框选内容 → 右键「导出选中内容为Word」
- 支持 DeepSeek、ChatGPT、Claude、豆包等所有 AI 平台
- 右下角浮动"W"按钮，一键导出整页内容

📊 **完整格式转换**
- Markdown / HTML 表格 → Word 原生表格（蓝底表头 + 交替行色）
- Mermaid 流程图/时序图/甘特图 → 自动渲染为高清 PNG 嵌入
- LaTeX 数学公式 → Word 原生 OMML 公式对象（可编辑）
- 代码块 → Pygments 语法高亮着色（50+ 语言）
- 任务列表 → ☐ / ☒ 复选框符号
- 引用块 → 左侧缩进 + 灰色背景 + 斜体样式
- 嵌套列表 → 多级缩进 List Bullet / List Number
- 图片 → 自动下载 URL / data:URI 并嵌入
- 分割线 → Word 水平线

📑 **文档增强**
- 自动生成目录 (TOC 域代码，在 Word 中右键更新)
- 图表自动编号 (图1、图2、表1、表2 递增标注)
- 三种报告格式预设：商务报告 / 学术论文 / 政府公文
- 可选封面页（标题 / 作者 / 日期）
- 页眉页脚自动设置（页码域 + 分割线）
- 全文格式统一（字体 / 字号 / 行距 / 页边距）

⚙️ **灵活配置**
- 导出选项：目录 / 编号 / 格式化 / 封面
- 报告预设一键切换
- 图片最大宽度可调
- 桥接服务地址可自定义

## 如何使用

1. 安装此扩展
2. 下载并运行本地桥接服务（见下方说明）
3. 浏览任何包含结构化内容的网页（如 DeepSeek 对话）
4. 点击右下角蓝色"W"浮动按钮，或选中内容后右键导出
5. .docx 文件自动下载，直接在 Microsoft Word 中打开编辑

## 本地桥接服务

本扩展通过本地 HTTP 服务 (127.0.0.1:9876) 实现完整的 Word 文档生成。所有数据在本地处理，不会上传到任何远程服务器。

桥接服务下载和安装说明: https://github.com/user/MarkdownPasteAddin

一键启动: 双击 `启动桥接服务.bat`（Windows）或运行 `python bridge_server.py`

## 兼容性

- Microsoft Edge 102+
- Google Chrome 114+
- 其他基于 Chromium 的浏览器（360极速、QQ浏览器等）

## 隐私声明

- 本扩展不会收集、存储、传输任何用户数据
- 所有内容处理仅限于用户本地计算机
- 与桥接服务的通信仅限于 127.0.0.1 (本机回环地址)
- 不使用任何分析或追踪工具
- 无需登录、无需注册
```

### 商店描述（English）

```
# MarkdownPasteAddin v3.0

One-click export of Markdown content, tables, flowcharts, math formulas, and code blocks to native Microsoft Word format (.docx).

## Key Features

📝 **Smart Content Capture**
- Auto-detect Markdown / code blocks / tables / Mermaid diagrams in any web page
- Right-click selection → "Export selected as Word"
- Works with DeepSeek, ChatGPT, Claude, Doubao, and all AI platforms
- Floating "W" button at bottom-right for one-click full-page export

📊 **Complete Format Conversion**
- Markdown / HTML tables → Word native tables (blue header + alternating rows)
- Mermaid diagrams (flowchart/sequence/gantt) → auto-rendered high-res PNG
- LaTeX math formulas → Word native OMML equation objects (editable)
- Code blocks → Pygments syntax highlighting (50+ languages)
- Task lists → ☐ / ☒ checkbox symbols
- Blockquotes → indented grey-background italic style
- Nested lists → multi-level List Bullet / List Number
- Images → auto-download URL / data:URI and embed
- Horizontal rules → Word horizontal lines

📑 **Document Enhancement**
- Auto Table of Contents (TOC field, right-click to update in Word)
- Auto figure/table numbering (Fig 1, Fig 2, Table 1, Table 2)
- Three Report Format Presets: Business / Academic / Government
- Optional cover page (title / author / date)
- Auto headers & footers (page numbers + separator lines)
- Full document formatting (font / size / line spacing / margins)

⚙️ **Flexible Configuration**
- Export options: TOC / Numbering / Formatting / Cover Page
- One-click report preset switching
- Adjustable image max width
- Customizable bridge server address

## How to Use

1. Install this extension
2. Download and run the local bridge server (see below)
3. Browse any page with structured content (e.g., DeepSeek conversations)
4. Click the blue "W" floating button at bottom-right, or right-click selection
5. .docx file auto-downloads, open directly in Microsoft Word

## Local Bridge Server

This extension communicates with a local HTTP bridge service (127.0.0.1:9876) for full-featured Word document generation. All data is processed locally on your computer — nothing is uploaded to any remote server.

Bridge server download: https://github.com/user/MarkdownPasteAddin

One-click launch: Run `启动桥接服务.bat` (Windows) or `python bridge_server.py`

## Compatibility

- Microsoft Edge 102+
- Google Chrome 114+
- Other Chromium-based browsers (360 Browser, QQ Browser, etc.)

## Privacy

- No data collection, storage, or transmission
- All content processed locally on user's computer
- Communication limited to 127.0.0.1 (localhost loopback)
- No analytics, no tracking
- No login, no registration required
```

## 三、审核要点

### Edge 扩展审核 vs Chrome

| 项目 | Chrome | Edge |
|------|--------|------|
| 审核时间 | 1-3 天 | 1-3 天 |
| 严格度 | 较高 | 中等 |
| 权限要求 | 最小权限原则 | 最小权限原则 |
| 外部通信 | 需声明 | 需声明 |
| 自动批准 | 不常见 | 偶尔 |
| 费用 | $5 一次性 | 免费 |

### 常见审核问题及预案

| 问题 | 解决方案 |
|------|----------|
| "host_permissions 过于宽泛" | 说明仅用于 localhost 桥接服务，无外部通信 |
| "缺少隐私政策" | 在扩展描述中明确 "No data collection" |
| "外部代码" | 说明 sidePanel 为浏览器原生 API，无远程代码 |
| "功能描述不匹配" | 确保截图与功能描述一致 |
| "clipboardRead 权限" | 说明仅用于读取用户主动复制的内容 |

## 四、提交清单

- [ ] 扩展文件打包为 .zip (31 files)
- [ ] 128x128 图标 PNG
- [ ] 至少 1 张 1280x800 截图
- [ ] 商店描述（中英文）
- [ ] 简短描述（45 字）
- [ ] 支持邮箱
- [ ] 隐私说明
- [ ] 测试通过：Edge 中正常加载和运行

## 五、Edge 侧载测试

提交前必须在 Edge 中测试：

```
1. 打开 Edge → edge://extensions/
2. 打开左下角「开发人员模式」
3. 点击「加载解压缩的扩展」
4. 选择 chrome-extension/ 目录
5. 验证：右下角浮动按钮出现
6. 验证：工具栏图标可点击弹出弹窗
7. 验证：侧边栏可打开 (edge://extensions 中配置)
8. 验证：选中内容右键 → 导出为Word
9. 验证：点击浮动按钮 → 整页导出
10. 验证：Paste & Export 按钮读取剪贴板
```

## 六、Edge 特有功能

| 功能 | Edge 支持 | Chrome 支持 |
|------|-----------|------------|
| Side Panel (侧边栏) | 原生支持 v102+ | Chrome 114+ |
| Bing Chat 集成 | 可通过 sidePanel | 不适用 |
| Collections 集成 | 可扩展 | 不适用 |
| Vertical Tabs 适配 | 自动适配 | 不适用 |

## 七、应用商店链接

- **Edge Add-ons 提交**: https://partner.microsoft.com/dashboard/windows/EdgeAddOns
- **Edge Add-ons 列表页**: https://microsoftedge.microsoft.com/addons/
- **Edge 扩展文档**: https://learn.microsoft.com/en-us/microsoft-edge/extensions-chromium/
- **Chrome Web Store 开发者控制台**: https://chrome.google.com/webstore/devconsole
