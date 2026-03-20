# v0.1.0

First public release of `outlook-oauth-mail-helper`.

## What this release includes

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

## Installation

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

## Minimal configuration

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

## Example commands

```bash
outlook-mail inbox -n 10
outlook-mail send --to someone@example.com --subject "Test" --body "Hello"
outlook-mail draft --to someone@example.com --subject "Draft" --body "Body"
outlook-mail attachments <message_id>
outlook-mail download-all-attachments <message_id> --outdir ./downloads
```

## Notes

This release is meant to be a practical first version for real mailbox workflows, not just a demo.

## Security reminder

Do not commit real OAuth credentials. Use `config.example.json` as a template only.
