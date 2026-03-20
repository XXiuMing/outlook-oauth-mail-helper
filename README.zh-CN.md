# outlook-oauth-mail-helper

[English README](./README.md)

[![Release](https://img.shields.io/github/v/release/XXiuMing/outlook-oauth-mail-helper)](https://github.com/XXiuMing/outlook-oauth-mail-helper/releases)
[![Python Check](https://img.shields.io/github/actions/workflow/status/XXiuMing/outlook-oauth-mail-helper/python-check.yml?branch=main&label=python%20check)](https://github.com/XXiuMing/outlook-oauth-mail-helper/actions/workflows/python-check.yml)
[![License](https://img.shields.io/github/license/XXiuMing/outlook-oauth-mail-helper)](./LICENSE)

`outlook-oauth-mail-helper` 是一个基于 Microsoft Graph 的 Outlook 邮件命令行工具。

它面向的是比较直接的邮箱工作：读邮件、发邮件、建草稿、回复、处理附件，以及导出正文和附件内容。项目同时支持基于 refresh token 的配置方式，所以如果你已经有可用的 `client_id + refresh_token`，可以把它整理成一套适合在终端、服务器、VPS 或自动化脚本中使用的邮件工作流。

## 主要功能

- 查看收件箱和读取单封邮件
- 发送新邮件
- 创建草稿并发送草稿
- 回复和回复全部
- 上传小附件和大附件
- 列出并下载附件
- 在指定文件夹中搜索邮件
- 标记已读 / 未读
- 移动和删除邮件
- 保存邮件正文为文本或 HTML
- 需要时自动刷新 access token

## 安装

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
pip install -r requirements.txt
chmod +x outlook_oauth_mail.py
```

如果你想直接使用更短的命令，可以本地安装：

```bash
pip install -e .
```

或者自己加一个 alias：

```bash
alias outlook-mail='python3 /绝对路径/outlook_oauth_mail.py'
```

## 配置

默认配置文件路径：

```bash
~/.outlook-oauth/config.json
```

最常见的配置方式是 refresh token 优先：

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

如果你已经有当前 access token，也可以写成：

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

字段含义：

- `client_id`：微软应用 Client ID
- `refresh_token`：OAuth refresh token
- `access_token`：当前 access token（可选）
- `tenant`：通常是 `common`、`consumers` 或具体租户 ID
- `scope`：刷新 token 时使用的 scope，默认是 `https://graph.microsoft.com/.default`
- `expires_at`：可选，Unix 时间戳，表示 access token 的过期时间

## Token 刷新逻辑

以下情况会自动刷新 token：

- 当前没有 `access_token`，但有可用的 `refresh_token`
- 当前 token 即将过期
- Graph API 返回 `401`

刷新成功后，新的 token 会自动写回配置文件。

## 常用命令

### 收件箱

```bash
outlook-mail inbox -n 10
outlook-mail inbox -u
```

### 读取单封邮件

```bash
outlook-mail read <message_id>
```

### 搜索邮件

```bash
outlook-mail search "invoice"
outlook-mail search "ChatGPT" --folder inbox -n 20
```

### 发邮件

```bash
outlook-mail send \
  --to someone@example.com \
  --subject "测试" \
  --body "你好"
```

### 草稿

```bash
outlook-mail draft --to someone@example.com --subject "草稿标题" --body "草稿内容"
outlook-mail drafts -n 20
outlook-mail send-draft <draft_id>
```

### 回复

```bash
outlook-mail reply <message_id> --body "收到，谢谢。"
outlook-mail reply <message_id> --body "收到，谢谢。" --all --send-now
```

### 附件

```bash
outlook-mail attach <draft_id> ./report.pdf
outlook-mail attachments <message_id>
outlook-mail download-attachment <message_id> <attachment_id> --outdir ./downloads
outlook-mail download-all-attachments <message_id> --outdir ./downloads
```

### 邮件管理

```bash
outlook-mail folders
outlook-mail mark <message_id> --read
outlook-mail move <message_id> archive
outlook-mail delete <message_id>
```

### 导出正文

```bash
outlook-mail save-body <message_id> --outdir ./saved-mails
outlook-mail save-body <message_id> --outdir ./saved-mails --html
```

## 使用示例

```bash
outlook-mail inbox -n 5
outlook-mail read <message_id>
outlook-mail reply <message_id> --body "收到，谢谢。"
```

如果是带附件的草稿流程：

```bash
outlook-mail draft --to someone@example.com --subject "报告" --body "见附件"
outlook-mail attach <draft_id> ./report.pdf
outlook-mail send-draft <draft_id>
```

## 适用场景

这个项目比较适合下面这些情况：

- 想在 Linux 终端里直接处理 Outlook 邮件
- 想在 VPS 或远程机器上跑邮件脚本
- 想基于 Microsoft Graph 做轻量级邮件自动化
- 已经有 refresh token，不想再依赖浏览器界面做日常操作

## FAQ

### 需要先自己做完整的微软应用集成吗？

通常还是需要一个有效的 Microsoft OAuth 配置，不过这个工具的重点是：当你已经有可用的 `client_id + refresh_token` 时，可以直接把它变成一套可用的 CLI 工作流。

### 支持大附件吗？

支持。3 MB 以内直接上传，更大的文件会走 Graph upload session。

### 会自动保存新的 token 吗？

会。刷新成功后，新的 token 会自动写回配置文件。

### 它是完整邮件客户端吗？

不是。它更像一个面向实际任务的命令行工具，而不是桌面邮件客户端的替代品。

## Roadmap

后续如果继续打磨，比较值得做的方向包括：

- 更细的附件下载过滤
- 更好的 HTML 正文转文本效果
- 比 smoke test 更完整的测试覆盖
- 更完整的打包和发布流程

## 仓库中的工程化文件

当前仓库还包含：

- `CHANGELOG.md`
- `CONTRIBUTING.md`
- `SECURITY.md`
- GitHub Actions 基础检查工作流
- issue 模板

## 安全说明

- 不要提交真实 OAuth 凭据
- 建议把真实配置保存在 `~/.outlook-oauth/config.json` 或其他被忽略的路径中
- 仓库中只保留 `config.example.json`

## 当前版本说明

`v0.1.0` 是第一个公开版本，已经覆盖常见邮箱工作，包括发信、草稿、回复、附件处理、文件夹操作，以及正文导出。

## License

MIT
