# MarkdownPasteAddin 浏览器扩展 v2.0

**支持浏览器**: Google Chrome 88+ | Microsoft Edge 102+ | 所有 Chromium 浏览器

一键将网页中的 Markdown/表格/流程图/数学公式导出为 Word 文档。通过本地桥接服务实现完整的 Python 转换能力。

---

## 快速开始

### 1. 启动桥接服务

```bash
# 确保已安装 Python 依赖
pip install -r ../requirements.txt

# 启动服务（保持窗口打开）
..\启动桥接服务.bat

# 或后台静默运行
..\启动桥接服务_后台.vbs
```

### 2. 加载扩展

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

### 3. 开始使用

- 浏览 AI 对话页面 → 右下角蓝色 **W** 按钮 → 点击导出
- 或：选中内容 → 右键 → "导出选中内容为Word"
- 或：点击工具栏图标 → 配置选项 → 导出

---

## 功能清单

| 功能 | 说明 |
|------|------|
| 表格 | Markdown/HTML → Word 原生表格 |
| 流程图 | Mermaid → 渲染为 PNG |
| 数学公式 | LaTeX → Word OMML |
| 代码高亮 | Pygments 语法着色 |
| 任务列表 | `- [ ]` → 复选框 |
| 引用块 | `>` → 引用样式 |
| 目录 | 自动插入 TOC 域 |
| 图表编号 | 图1/表1 自动递增 |
| 格式化 | 导出后统一字体/字号/行距 |

---

## 跨浏览器差异

| 特性 | Chrome | Edge |
|------|--------|------|
| 扩展 API | chrome.* | chrome.* (完全兼容) |
| 侧边栏 | v114+ (实验性) | v102+ (原生) |
| 右键菜单 | 支持 | 支持 |
| 下载API | downloads | downloads |
| 扩展商店 | Chrome Web Store ($5) | Edge Add-ons (免费) |
| 侧载 | 开发者模式 | 开发者模式 |

---

## 上架商店

- **Chrome Web Store**: 见 [STORE_LISTING.md](./STORE_LISTING.md)
- **Edge Add-ons**: `STORE_LISTING.md` 中有完整的提交指南

两个商店可以使用同一份扩展代码包。

---

## 故障排除

| 问题 | 解决 |
|------|------|
| 右下角按钮不出现 | 刷新页面，检查扩展是否已启用 |
| 状态灯红色 | 启动桥接服务 `启动桥接服务.bat` |
| 导出无响应 | 访问 http://localhost:9876/api/health 检查服务 |
| Edge 中无法加载 | 在 edge://extensions 打开开发人员模式 |

---

## 文件结构

```
chrome-extension/
├── manifest.json              # Manifest V3 (Chrome + Edge)
├── background.js              # Service Worker
├── lib/
│   └── browser-polyfill.js    # 跨浏览器 API 标准化
├── content/
│   ├── content_script.js      # 页面注入 + 内容捕获
│   └── content_style.css      # 浮动按钮样式
├── popup/
│   ├── popup.html/css/js      # 工具栏弹窗
│   ├── sidepanel.html         # 侧边栏 (Edge 原生支持)
│   └── sidepanel.js           # 侧边栏逻辑
├── options/
│   └── options.html/js        # 设置页面
├── icons/                     # 扩展图标
├── _locales/zh_CN/            # 中文本地化
└── README.md
```
