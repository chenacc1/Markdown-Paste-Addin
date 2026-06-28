# Freemium 免费增值 — 待办清单

## 整体架构

```
扩展（免费 JS）→ 基础功能（表格/文本/标题/列表/任务/引用/分割线）
桥接服务（Pro）→ 高级功能（数学公式 OMML / Mermaid 渲染 / 代码高亮 / 报告预设）
                    ↑
             校验 License Key
```

---

## Phase 1：License 系统（核心基础设施）

### 1.1 generate_license.py — Key 生成器

- [ ] 命令行工具，支持参数：
  - `--email` 用户邮箱
  - `--tier` 档位（pro / premium）
  - `--months` 有效期月数
- [ ] Key 格式：`MP-{tier}-{email_hash}-{expiry}-{hmac}`
- [ ] Key 内嵌：邮箱 SHA256 前 8 位、档位、到期时间戳、HMAC-SHA256 签名防伪造
- [ ] 输出为 `license_key.txt` 或直接打印

### 1.2 license_manager.py — md2docx_lib 新模块

- [ ] `validate_key(key: str) -> dict` 校验 key 是否有效，返回到期时间、档位
- [ ] `is_expired(expiry: str) -> bool` 判断是否过期
- [ ] `get_tier_features(tier: str) -> list` 返回该档位解锁的功能列表
- [ ] 无 key 时默认返回基础功能集

### 1.3 bridge_server.py — 加 License 校验层

- [ ] `POST /api/convert` 新增 `license_key` 字段
- [ ] 请求到达时先校验 key，根据 tier 决定开启哪些 Pro 功能
- [ ] 无 key / key 失效时返回仅基础功能的 .docx
- [ ] 响应中增加 `warning` 字段提示升级："数学公式需要 Pro 版，前往 xxx 升级"

### 1.4 扩展 options.html/options.js — 加 License 输入

- [ ] 新增 License Key 输入框 + 保存按钮
- [ ] 保存到 `chrome.storage.local`
- [ ] `background.js` 调用桥接服务时自动附带 key

### 1.5 扩展 popup.js — Pro 状态提示

- [ ] 显示当前 License 状态（Pro / 免费 / 已过期）
- [ ] 免费版 Pro 功能灰显 + 链接导向付费页

---

## Phase 2：付费与分发

### 2.1 爱发电 Webhook → 自动发 Key

- [ ] 搭一个简单的 Flask 端点接收爱发电付费通知
- [ ] 验证 Webhook 签名
- [ ] 自动调用 `generate_license.py` → 发邮件给用户
- [ ] 备选：初期手动发（用户量 < 50 时完全可行）

### 2.2 付费页面

- [ ] 单页 HTML（放在 GitHub Pages）：套餐说明 + 购买按钮
- [ ] 三档：免费 / Pro（9.9元/月）/ 企业（49元/年）
- [ ] 购买按钮跳转爱发电对应方案

---

## Phase 3：用户引导与体验

### 3.1 首次安装引导

- [ ] 安装后弹出介绍页：功能对比表 + "免费试用 Pro 7 天"
- [ ] 试用期自动生成 7 天临时 key（本地生成，不上报）

### 3.2 功能对比表

| 功能 | 免费版（扩展直接） | Pro 版（桥接服务） |
|------|-------------------|-------------------|
| 表格 | ✓ | ✓ |
| 标题/文本/列表 | ✓ | ✓ |
| 任务列表 | ✓ | ✓ |
| 引用块/分割线 | ✓ | ✓ |
| 数学公式 OMML | ✗ | ✓ |
| Mermaid 流程图 | ✗ | ✓ |
| 代码语法高亮 | ✗ | ✓ |
| 报告格式预设 | ✗ | ✓ |
| 封面页/页眉页脚 | ✗ | ✓ |
| 图表自动编号 | ✗ | ✓ |
| TOC 目录 | ✗ | ✓ |
| 批量转换 | ✗ | ✓ |

---

## 技术要点

- **本地校验**：License 校验完全在用户本地 Python 服务完成，不调用外部 API，隐私安全
- **Key 不可伪造**：HMAC-SHA256 签名的密钥确保只有你能生成有效 Key
- **离线友好**：基础功能不依赖任何服务，Pro 功能只需本地桥接服务
- **平滑降级**：无 key 时 Pro 功能自动跳过，不报错

---

*状态：规划阶段 | 最后更新：2026-06-29*
