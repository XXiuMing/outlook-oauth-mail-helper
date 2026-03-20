# outlook-oauth-mail-helper

[English README](./README.md)

一个面向 **Outlook / Microsoft Graph** 的实用型 Python 命令行工具，支持：

- 读取邮件
- 发送邮件
- 创建草稿
- 回复邮件
- 搜索邮件
- 导出正文
- 下载附件
- 管理文件夹与邮件状态
- 基于 **OAuth refresh token** 的自动续期工作流

如果你已经有可用的 `client_id + refresh_token`，又不想上很重的 Web 面板，这个工具就是给这种场景准备的。

---

## 项目定位

这个项目想解决的问题很明确：

很多 Outlook 自动化方案要么太重、要么太散、要么只够演示，真正落到“日常收发邮件 + 附件 + 草稿 + 导出”时并不好用。

这个项目的目标是：

- **足够轻**：一个 Python CLI 就能跑
- **足够实用**：覆盖真实邮箱工作流
- **足够直接**：拿 refresh token 就能干活
- **足够脚本化**：适合服务器、VPS、本地终端、自动化流程

---

## 首屏介绍文案（可用于仓库首页理解）

**outlook-oauth-mail-helper** 是一个小而实用的 Outlook 邮件 CLI。它把 Outlook OAuth 凭据直接变成可用的命令行邮箱工作流。

它擅长的不是“做一个很大的平台”，而是把这些高频动作做好：

- 快速看收件箱
- 读单封邮件
- 搜索邮件
- 创建草稿
- 发邮件
- 回复邮件
- 上传大附件
- 导出正文和附件
- 移动 / 删除 / 标记邮件

---

## 支持的功能

### 邮件读取
- 列出收件箱邮件
- 读取单封邮件全文
- 在指定文件夹内搜索邮件
- 列出草稿
- 列出邮箱文件夹

### 邮件发送
- 发送新邮件
- 创建草稿
- 发送草稿
- 回复邮件
- 回复全部

### 附件处理
- 给草稿加附件
- **小附件直接上传**
- **大附件走 Graph upload session 分片上传**
- 列出邮件附件
- 下载单个附件
- 下载全部文件附件

### 邮件管理
- 标记已读 / 未读
- 移动邮件
- 删除邮件
- 保存邮件正文为 `.txt` 或 `.html`

### OAuth / Token 流程
- 支持已有 `access_token` 的配置
- 支持 `refresh_token` 优先配置
- 自动刷新 access token
- 自动把新 token 写回配置文件

---

## 功能总览

| 功能 | 是否支持 |
|---|---|
| 收件箱读取 | 是 |
| 单封邮件阅读 | 是 |
| 搜索邮件 | 是 |
| 发送邮件 | 是 |
| 草稿 | 是 |
| 回复 / 回复全部 | 是 |
| 小附件上传 | 是 |
| 大附件上传 | 是 |
| 列附件 | 是 |
| 下载单个附件 | 是 |
| 下载全部附件 | 是 |
| 列文件夹 | 是 |
| 移动 / 删除邮件 | 是 |
| 标记已读未读 | 是 |
| 保存正文到文件 | 是 |

---

## 安装

### 1）克隆仓库

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
```

### 2）安装依赖

```bash
pip install -r requirements.txt
```

### 3）给脚本执行权限

```bash
chmod +x outlook_oauth_mail.py
```

### 4）可选：加一个 alias

```bash
alias outlook-mail='python3 /绝对路径/outlook_oauth_mail.py'
```

以后就可以直接：

```bash
outlook-mail inbox -n 10
```

---

## 配置

默认配置路径：

```bash
~/.outlook-oauth/config.json
```

也可以手动指定：

```bash
python3 outlook_oauth_mail.py --config /path/to/config.json inbox
```

### 方案 A：refresh token 优先

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

### 方案 B：access token + refresh token

```json
{
  "access_token": "<ACCESS_TOKEN>",
  "refresh_token": "<REFRESH_TOKEN>",
  "client_id": "<CLIENT_ID>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default",
  "expires_at": 1760000000
}
```

### 字段说明

- `client_id`：微软应用 Client ID
- `refresh_token`：OAuth 刷新令牌
- `access_token`：当前访问令牌（可选）
- `tenant`：通常是 `common` / `consumers` / 租户 ID
- `scope`：刷新时使用的 scope，默认是 `https://graph.microsoft.com/.default`
- `expires_at`：access token 过期时间戳（可选）

---

## Token 刷新逻辑

以下情况会自动刷新：

- 没有 `access_token`，但有有效 `refresh_token`
- token 快过期
- Graph 返回 `401`

刷新成功后，会自动把新的 token 写回配置文件。

---

## 终端示例

### 看收件箱

```bash
outlook-mail inbox -n 5
```

示例输出：

```text
{
  "value": [
    {
      "id": "AQMk...",
      "receivedDateTime": "2026-03-19T23:44:14Z",
      "subject": "Body save test",
      "isRead": true
    }
  ]
}
```

### 发邮件

```bash
outlook-mail send \
  --to someone@example.com \
  --subject "测试" \
  --body "Hello from outlook-oauth-mail-helper"
```

示例输出：

```text
Mail sent.
```

### 建草稿 + 加附件 + 发出

```bash
outlook-mail draft --to someone@example.com --subject "报告" --body "见附件"
outlook-mail attach <draft_id> ./report.pdf
outlook-mail send-draft <draft_id>
```

---

## 常用命令

### 收件箱

```bash
python3 outlook_oauth_mail.py inbox -n 10
python3 outlook_oauth_mail.py inbox -u
```

### 读单封邮件

```bash
python3 outlook_oauth_mail.py read <message_id>
```

### 搜索邮件

```bash
python3 outlook_oauth_mail.py search "关键词"
python3 outlook_oauth_mail.py search "invoice" --folder inbox -n 20
```

### 发邮件

```bash
python3 outlook_oauth_mail.py send \
  --to someone@example.com \
  --subject "Test" \
  --body "Hello"
```

### 创建草稿

```bash
python3 outlook_oauth_mail.py draft \
  --to someone@example.com \
  --subject "Draft title" \
  --body "Draft content"
```

### 回复邮件

```bash
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks"
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks" --all --send-now
```

### 附件

```bash
python3 outlook_oauth_mail.py attach <draft_id> /path/to/file.pdf
python3 outlook_oauth_mail.py attachments <message_id>
python3 outlook_oauth_mail.py download-attachment <message_id> <attachment_id> --outdir ./downloads
python3 outlook_oauth_mail.py download-all-attachments <message_id> --outdir ./downloads
```

### 邮件管理

```bash
python3 outlook_oauth_mail.py folders
python3 outlook_oauth_mail.py mark <message_id> --read
python3 outlook_oauth_mail.py move <message_id> archive
python3 outlook_oauth_mail.py delete <message_id>
```

### 导出正文

```bash
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./saved-mails
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./saved-mails --html
```

---

## 典型工作流

### 看邮件后回复

```bash
python3 outlook_oauth_mail.py inbox -n 5
python3 outlook_oauth_mail.py read <message_id>
python3 outlook_oauth_mail.py reply <message_id> --body "收到，谢谢。"
```

### 建草稿并发送附件

```bash
python3 outlook_oauth_mail.py draft --to someone@example.com --subject "报告" --body "见附件"
python3 outlook_oauth_mail.py attach <draft_id> ./report.pdf
python3 outlook_oauth_mail.py send-draft <draft_id>
```

### 导出正文和附件

```bash
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./export
python3 outlook_oauth_mail.py download-all-attachments <message_id> --outdir ./export
```

---

## 安全说明

- **不要把真实 token 提交进 git**
- 真实配置建议放在 `~/.outlook-oauth/config.json`
- 仓库里只保留 `config.example.json`
- 使用前请确认 refresh token 来源与权限范围

---

## 当前限制

- 大附件虽然支持 upload session，但还没做断点续传
- 搜索目前按单文件夹进行，不是全邮箱搜索
- HTML 转纯文本目前只是简单去标签，不是完整渲染

---

## 适合谁

如果你需要这些东西，这个项目就很适合：

- 一个轻量级 Outlook CLI
- 基于 Microsoft Graph 的邮件自动化
- 基于 refresh token 的邮箱工作流
- 在 Linux / VPS / 本地终端里跑邮箱脚本

---

## 发布状态

当前推荐起点：**v0.1.0**

这个版本已经能覆盖真实邮箱场景：
- 收发信
- 草稿
- 回复
- 大附件
- 下载附件
- 文件夹操作
- 正文导出

---

## License

MIT
