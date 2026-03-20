# outlook-oauth-mail-helper

A lightweight Python CLI for reading and sending Outlook mail through Microsoft Graph using OAuth refresh tokens.

It supports the refresh-token format commonly exported by the `assast/outlookEmail` project and can also work with a normal access-token-based config.

## Features

- Read inbox messages
- Read a single message
- Search messages in a folder
- Send mail
- Create drafts and send drafts
- Reply / reply-all
- Upload small and large attachments
- List mail folders
- Mark read / unread
- Move and delete messages
- List attachments
- Download one or all attachments
- Save message body to a file

## Installation

```bash
pip install -r requirements.txt
chmod +x outlook_oauth_mail.py
```

Optional shell alias:

```bash
alias outlook-mail='python3 /path/to/outlook_oauth_mail.py'
```

## Configuration

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

## Common commands

### Inbox

```bash
python3 outlook_oauth_mail.py inbox -n 10
python3 outlook_oauth_mail.py inbox -u
```

### Read one message

```bash
python3 outlook_oauth_mail.py read <message_id>
```

### Search

```bash
python3 outlook_oauth_mail.py search "keyword"
python3 outlook_oauth_mail.py search "invoice" --folder inbox -n 20
```

### Send mail

```bash
python3 outlook_oauth_mail.py send \
  --to someone@example.com \
  --subject "Test" \
  --body "Hello"
```

### Drafts and attachments

```bash
python3 outlook_oauth_mail.py draft --to someone@example.com --subject "Draft" --body "Body"
python3 outlook_oauth_mail.py attach <draft_id> /path/to/file.pdf
python3 outlook_oauth_mail.py send-draft <draft_id>
```

### Replies

```bash
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks"
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks" --all --send-now
```

### Attachments

```bash
python3 outlook_oauth_mail.py attachments <message_id>
python3 outlook_oauth_mail.py download-attachment <message_id> <attachment_id> --outdir ./downloads
python3 outlook_oauth_mail.py download-all-attachments <message_id> --outdir ./downloads
```

### Mail management

```bash
python3 outlook_oauth_mail.py folders
python3 outlook_oauth_mail.py mark <message_id> --read
python3 outlook_oauth_mail.py move <message_id> archive
python3 outlook_oauth_mail.py delete <message_id>
```

### Save message body

```bash
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./saved-mails
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./saved-mails --html
```

## Security notes

- Never commit real OAuth tokens.
- Keep your real config in `~/.outlook-oauth/config.json` or another ignored path.
- This repo includes only `config.example.json`.

## License

MIT
