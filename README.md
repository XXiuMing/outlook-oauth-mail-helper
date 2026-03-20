# outlook-oauth-mail-helper

[中文说明](./README.zh-CN.md)

[![Release](https://img.shields.io/github/v/release/XXiuMing/outlook-oauth-mail-helper)](https://github.com/XXiuMing/outlook-oauth-mail-helper/releases)
[![Python Check](https://img.shields.io/github/actions/workflow/status/XXiuMing/outlook-oauth-mail-helper/python-check.yml?branch=main&label=python%20check)](https://github.com/XXiuMing/outlook-oauth-mail-helper/actions/workflows/python-check.yml)
[![License](https://img.shields.io/github/license/XXiuMing/outlook-oauth-mail-helper)](./LICENSE)

`outlook-oauth-mail-helper` is a command-line tool for working with Outlook mail through Microsoft Graph.

It covers the common tasks you would expect in a mailbox workflow: reading messages, sending mail, creating drafts, replying, handling attachments, exporting message content, and performing basic folder operations. It also supports refresh-token-based setups, which makes it practical for server-side scripts, remote machines, and other non-interactive environments.

## Main features

- Read inbox messages and full message content
- Send new mail
- Create drafts and send drafts
- Reply and reply-all
- Upload small and large attachments
- List and download attachments
- Search within folders
- Mark read / unread
- Move and delete messages
- Save message bodies as text or HTML
- Refresh access tokens automatically when needed

## Installation

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
pip install -r requirements.txt
chmod +x outlook_oauth_mail.py
```

If you want a shorter command, you can install it locally:

```bash
pip install -e .
```

or define an alias:

```bash
alias outlook-mail='python3 /absolute/path/to/outlook_oauth_mail.py'
```

## Configuration

By default the tool reads:

```bash
~/.outlook-oauth/config.json
```

A minimal refresh-token-based config looks like this:

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

You can also include a current access token and expiry metadata:

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

Important fields:

- `client_id`: Microsoft application client ID
- `refresh_token`: OAuth refresh token
- `access_token`: optional current access token
- `tenant`: usually `common`, `consumers`, or a tenant ID
- `scope`: refresh scope, defaulting to `https://graph.microsoft.com/.default`
- `expires_at`: optional Unix timestamp for token expiry

## Token handling

The tool refreshes access tokens automatically when:

- there is no current `access_token`, but a valid `refresh_token` exists
- the current token is near expiry
- Microsoft Graph returns `401`

If refresh succeeds, updated token values are written back to the config file.

## Common commands

### Inbox

```bash
outlook-mail inbox -n 10
outlook-mail inbox -u
```

### Read one message

```bash
outlook-mail read <message_id>
```

### Search

```bash
outlook-mail search "invoice"
outlook-mail search "ChatGPT" --folder inbox -n 20
```

### Send mail

```bash
outlook-mail send \
  --to someone@example.com \
  --subject "Test" \
  --body "Hello"
```

### Drafts

```bash
outlook-mail draft --to someone@example.com --subject "Draft title" --body "Draft content"
outlook-mail drafts -n 20
outlook-mail send-draft <draft_id>
```

### Replies

```bash
outlook-mail reply <message_id> --body "Thanks"
outlook-mail reply <message_id> --body "Thanks" --all --send-now
```

### Attachments

```bash
outlook-mail attach <draft_id> ./report.pdf
outlook-mail attachments <message_id>
outlook-mail download-attachment <message_id> <attachment_id> --outdir ./downloads
outlook-mail download-all-attachments <message_id> --outdir ./downloads
```

### Mail management

```bash
outlook-mail folders
outlook-mail mark <message_id> --read
outlook-mail move <message_id> archive
outlook-mail delete <message_id>
```

### Export message body

```bash
outlook-mail save-body <message_id> --outdir ./saved-mails
outlook-mail save-body <message_id> --outdir ./saved-mails --html
```

## Example session

```bash
outlook-mail inbox -n 5
outlook-mail read <message_id>
outlook-mail reply <message_id> --body "Got it, thanks."
```

For a draft with an attachment:

```bash
outlook-mail draft --to someone@example.com --subject "Report" --body "See attached."
outlook-mail attach <draft_id> ./report.pdf
outlook-mail send-draft <draft_id>
```

## Use cases

This project is useful if you want to:

- work with Outlook mail from a Linux shell
- automate mailbox tasks on a VPS or remote machine
- build simple scripts around Microsoft Graph mail operations
- manage a mailbox with refresh-token-based authentication rather than a browser UI

## FAQ

### Do I need a Microsoft OAuth setup first?

Yes. You still need a valid Microsoft OAuth configuration. The tool is mainly intended for the stage after that, when you already have usable credentials and want a practical CLI workflow.

### Can it handle large attachments?

Yes. Files up to 3 MB are uploaded directly. Larger files use a Graph upload session.

### Does it store refreshed tokens automatically?

Yes. When token refresh succeeds, updated token values are written back to the config file.

### Is this a full mail client?

No. It is a command-line helper for common mailbox operations, not a replacement for a desktop mail application.

## Roadmap

Near-term improvements that would make sense for future versions include:

- better filtering for attachment downloads
- cleaner plain-text export for HTML-heavy messages
- broader test coverage beyond smoke checks
- incremental packaging and release improvements

## Project files

This repository includes:

- `CHANGELOG.md`
- `CONTRIBUTING.md`
- `SECURITY.md`
- a GitHub Actions workflow for basic Python checks
- issue templates for bug reports and feature requests

## Security notes

- Never commit real OAuth credentials
- Keep your real config in `~/.outlook-oauth/config.json` or another ignored path
- This repository only includes `config.example.json`

## Current scope

`v0.1.0` is the first public release. It already covers the core day-to-day tasks: sending, drafting, replying, attachment handling, folder operations, and body export.

## License

MIT
