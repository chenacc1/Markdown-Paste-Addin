# MarkdownPasteAddin 浏览器扩展 v3.2

**支持浏览器**: Google Chrome 88+ | Microsoft Edge 102+ | 所有 Chromium 浏览器

一键将网页中的 Markdown/表格/代码块导出为 Word 文档。**安装即用，离线优先**，无需安装 Python 或启动额外服务。桥接服务为可选增强。

---

## 快速开始

### 1. 加载扩展

**Chrome:**
```
chrome://extensions/ → 开发者模式 → 加载已解压的扩展程序
→ 选择 chrome-extension/ 目录
```

**Microsoft Edge:**
```
edge://extensions/ → 开发人员模式 → 加载解压缩的扩展
→ 选择 chrome-extension/ 目录
```

### 2. 开始使用

- 浏览 AI 对话页面 → 右下角蓝色 **W** 按钮 → 点击导出
- 或：选中内容 → 右键 → "导出选中内容为Word"
- 或：点击工具栏图标 → 配置选项 → 导出

---

## 功能清单

| 功能 | 离线模式 | 桥接增强 |
|------|---------|---------|
| 标题 (H1-H6) | 支持 | 支持 |
| 正文 (粗体/斜体/行内代码) | 支持 | 支持 |
| 表格 (蓝底表头) | 支持 | 支持 |
| 代码块 (等宽字体) | 支持 | 语法高亮 |
| 任务列表 (☐/☒) | 支持 | 支持 |
| 引用块 | 支持 | 支持 |
| 分割线 | 支持 | 支持 |
| 流程图 (Mermaid) | — | 渲染为 PNG |
| 数学公式 (LaTeX) | — | Word OMML |
| 目录 (TOC) | — | TOC 域代码 |
| 格式预设 (商务/学术/公文) | — | 全文格式化 |

---

## 离线模式 vs 桥接服务

**离线模式**（默认）：扩展内建 JS docx 引擎，安装即用，覆盖日常所需。

**桥接服务**（可选增强）：启动本地 Python 服务可获得完整功能。

```bash
pip install -r requirements.txt
启动桥接服务.bat
```

---

## 故障排除

| 问题 | 解决 |
|------|------|
| 右下角按钮不出现 | 刷新页面，检查扩展是否已启用 |
| 导出的 .docx 格式不理想 | 启动桥接服务获得完整格式化 |
| Edge 中无法加载 | 在 edge://extensions 打开开发人员模式 |

---

## 文件结构

```
chrome-extension/
├── manifest.json              # Manifest V3 (Chrome + Edge)
├── background.js              # Service Worker (bridge + offline)
├── lib/
│   ├── browser-polyfill.js    # 跨浏览器 API 标准化
│   └── docx-builder.js        # 离线 docx 生成引擎
├── content/
│   ├── content_script.js      # 页面注入 + 内容捕获
│   └── content_style.css      # 浮动按钮样式
├── popup/
│   ├── popup.html/css/js      # 工具栏弹窗
│   ├── sidepanel.html/js      # 侧边栏 (Edge 原生支持)
├── options/
│   └── options.html/js        # 设置页面
├── _locales/zh_CN/            # 中文本地化
└── icons/                     # 扩展图标 (16/48/128)
```
