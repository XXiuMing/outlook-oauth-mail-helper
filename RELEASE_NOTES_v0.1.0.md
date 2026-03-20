# v0.1.0

## English

First public release of `outlook-oauth-mail-helper`.

### Included in this release

- Outlook mail access through Microsoft Graph
- Refresh-token-first workflow support
- Automatic access-token refresh
- Inbox reading and full message reading
- Search within folders
- New mail sending
- Draft creation and draft sending
- Reply / reply-all support
- Small attachment upload
- Large attachment upload with Graph upload sessions
- Attachment listing and download
- Bulk attachment download
- Folder listing
- Mark read / unread
- Move and delete messages
- Save message body as text or HTML

### Installation

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
pip install -r requirements.txt
chmod +x outlook_oauth_mail.py
```

Optional editable install:

```bash
pip install -e .
```

### Minimal configuration

Default config path:

```bash
~/.outlook-oauth/config.json
```

Example:

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

### Example commands

```bash
outlook-mail inbox -n 10
outlook-mail send --to someone@example.com --subject "Test" --body "Hello"
outlook-mail draft --to someone@example.com --subject "Draft" --body "Body"
outlook-mail attachments <message_id>
outlook-mail download-all-attachments <message_id> --outdir ./downloads
```

### Notes

This release is intended as a practical first version for everyday mailbox work.

### Verification note

A local end-to-end workflow check was completed after release, covering inbox read, message read, search, folders, mark read/unread, draft creation, small and large attachment upload, attachment download, body export, draft listing, move, and delete.

### Security reminder

Do not commit real OAuth credentials. Use `config.example.json` as a template only.

---

## 中文

`outlook-oauth-mail-helper` 的第一个公开版本。

### 本次版本包含

- 基于 Microsoft Graph 的 Outlook 邮件访问能力
- refresh token 优先工作流
- 自动刷新 access token
- 收件箱读取与单封邮件阅读
- 文件夹内搜索邮件
- 新邮件发送
- 草稿创建与发送
- 回复 / 回复全部
- 小附件上传
- 基于 upload session 的大附件上传
- 附件列表与下载
- 批量下载附件
- 文件夹列表
- 标记已读 / 未读
- 移动和删除邮件
- 保存邮件正文为文本或 HTML

### 安装

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
pip install -r requirements.txt
chmod +x outlook_oauth_mail.py
```

可选的本地安装方式：

```bash
pip install -e .
```

### 最小配置

默认配置路径：

```bash
~/.outlook-oauth/config.json
```

示例：

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

### 常用命令示例

```bash
outlook-mail inbox -n 10
outlook-mail send --to someone@example.com --subject "测试" --body "你好"
outlook-mail draft --to someone@example.com --subject "草稿" --body "内容"
outlook-mail attachments <message_id>
outlook-mail download-all-attachments <message_id> --outdir ./downloads
```

### 说明

这个版本的目标是提供一套适合日常邮箱工作的可用 CLI。

### 安全提醒

不要提交真实 OAuth 凭据。仓库中的 `config.example.json` 仅用于示例。

### 排障提示

如果刷新 token 时出现 `invalid_scope`，先确认配置里使用的是：

```json
"scope": "https://graph.microsoft.com/.default"
```

不要把 `.default` 和资源级 scope 混在同一次刷新请求里。
示例。
