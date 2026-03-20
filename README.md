# outlook-oauth-mail-helper

[中文说明 / Chinese README](./README.zh-CN.md)

A practical Python CLI for **reading, sending, drafting, replying, searching, exporting, and managing Outlook email** through **Microsoft Graph**, with first-class support for **OAuth refresh-token workflows**.

It is designed for people who already have a usable Outlook / Microsoft OAuth refresh token and want a lightweight command-line tool instead of a full web panel.

---

## Hero / Project intro

**Outlook OAuth Mail Helper** is a small but capable command-line tool that turns Outlook OAuth credentials into a usable mailbox workflow.

It focuses on the things people actually need:

- check inbox quickly
- read one message in full
- search mail
- create drafts
- send mail
- reply to mail
- upload large attachments
- export message body and attachments
- move / delete / mark messages

If you already have a `client_id + refresh_token` pair, this project gives you a direct and scriptable way to work with Outlook mail on a server, VPS, local machine, or automation box.

---

## Why this project exists

Most Outlook automation examples fall into one of these buckets:

- too heavyweight
- too tied to a web dashboard
- too incomplete for real mailbox work
- or assume you want to build an entire Microsoft app integration from scratch

This project aims to be the opposite:

- **small enough to understand quickly**
- **practical enough for daily use**
- **good at real mailbox tasks**
- **friendly to refresh-token-based setups**

It works especially well if you already have a `client_id + refresh_token` pair from an Outlook OAuth flow and want to turn that into a usable mail CLI.

---

## Terminal examples

### Read inbox

```bash
outlook-mail inbox -n 5
```

Example output:

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

### Send a message

```bash
outlook-mail send \
  --to someone@example.com \
  --subject "Test" \
  --body "Hello from outlook-oauth-mail-helper"
```

Example output:

```text
Mail sent.
```

### Create a draft, attach a file, then send

```bash
outlook-mail draft --to someone@example.com --subject "Report" --body "See attached."
outlook-mail attach <draft_id> ./report.pdf
outlook-mail send-draft <draft_id>
```

### Export an email

```bash
outlook-mail save-body <message_id> --outdir ./export
outlook-mail download-all-attachments <message_id> --outdir ./export
```

---

## What it can do

### Mail reading
- List inbox messages
- Read a full message
- Search messages inside a folder
- List draft messages
- List mail folders

### Mail sending
- Send a new email
- Create drafts
- Send existing drafts
- Reply to a message
- Reply-all to a message

### Attachments
- Attach files to drafts
- Upload **small attachments directly**
- Upload **large attachments with Graph upload sessions**
- List attachments on a message
- Download one attachment
- Download all file attachments

### Message management
- Mark read / unread
- Move messages to another folder
- Delete messages
- Save message body as `.txt` or `.html`

### OAuth / token flow
- Use a normal `access_token` config
- Or use **refresh-token-first config**
- Automatically refresh access tokens when needed
- Persist updated tokens back to config

---

## Feature summary

| Area | Supported |
|---|---|
| Inbox listing | Yes |
| Read full message | Yes |
| Search | Yes |
| Send mail | Yes |
| Drafts | Yes |
| Reply / reply-all | Yes |
| Small attachment upload | Yes |
| Large attachment upload | Yes |
| Attachment listing | Yes |
| Single attachment download | Yes |
| Bulk attachment download | Yes |
| Folder listing | Yes |
| Move / delete | Yes |
| Mark read / unread | Yes |
| Save body to file | Yes |

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
```

### 2. Install dependency

```bash
pip install -r requirements.txt
```

### 3. Make the script executable

```bash
chmod +x outlook_oauth_mail.py
```

### 4. Optional: add a shell alias

```bash
alias outlook-mail='python3 /absolute/path/to/outlook_oauth_mail.py'
```

After that, you can use:

```bash
outlook-mail inbox -n 10
```

---

## Configuration

Default config path:

```bash
~/.outlook-oauth/config.json
```

You can also pass a custom config path with:

```bash
python3 outlook_oauth_mail.py --config /path/to/config.json inbox
```

### Option A: refresh-token-first config

```json
{
  "client_id": "<CLIENT_ID>",
  "refresh_token": "<REFRESH_TOKEN>",
  "tenant": "common",
  "scope": "https://graph.microsoft.com/.default"
}
```

### Option B: access token + refresh token config

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

### Field meanings

- `client_id`: Microsoft app client ID
- `refresh_token`: OAuth refresh token
- `access_token`: optional current access token
- `tenant`: usually `common`, `consumers`, or a tenant ID
- `scope`: token refresh scope; by default this project uses `https://graph.microsoft.com/.default`
- `expires_at`: optional Unix timestamp for access token expiry

---

## Token refresh behavior

The CLI automatically refreshes tokens when:

- there is **no** `access_token` but a valid `refresh_token` exists
- the current token is near expiry
- Microsoft Graph returns `401`

If refresh succeeds, the tool writes the new token values back into your config file.

---

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

### Search in a folder

```bash
python3 outlook_oauth_mail.py search "keyword"
python3 outlook_oauth_mail.py search "invoice" --folder inbox -n 20
```

### Send a new message

```bash
python3 outlook_oauth_mail.py send \
  --to someone@example.com \
  --subject "Test" \
  --body "Hello"
```

### Create a draft

```bash
python3 outlook_oauth_mail.py draft \
  --to someone@example.com \
  --subject "Draft title" \
  --body "Draft content"
```

### Reply

```bash
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks"
python3 outlook_oauth_mail.py reply <message_id> --body "Thanks" --all --send-now
```

### Attachments

```bash
python3 outlook_oauth_mail.py attach <draft_id> /path/to/file.pdf
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

---

## Practical workflows

### Read then reply

```bash
python3 outlook_oauth_mail.py inbox -n 5
python3 outlook_oauth_mail.py read <message_id>
python3 outlook_oauth_mail.py reply <message_id> --body "Got it, thanks."
```

### Create draft with attachment and send

```bash
python3 outlook_oauth_mail.py draft --to someone@example.com --subject "Report" --body "See attached."
python3 outlook_oauth_mail.py attach <draft_id> ./report.pdf
python3 outlook_oauth_mail.py send-draft <draft_id>
```

### Export an email and its attachments

```bash
python3 outlook_oauth_mail.py save-body <message_id> --outdir ./export
python3 outlook_oauth_mail.py download-all-attachments <message_id> --outdir ./export
```

---

## Security notes

- **Never commit real OAuth tokens**
- Keep your real config in `~/.outlook-oauth/config.json` or another ignored path
- This repository only includes `config.example.json`
- Review your refresh token source and permissions before use

---

## Current limitations

- no resumable recovery for interrupted large-file uploads yet
- search currently works folder-by-folder instead of whole-mailbox search
- HTML-to-text body export uses simple tag stripping, not full HTML rendering

---

## Who this is for

This project is useful if you want:

- a small Outlook CLI
- Microsoft Graph mail automation without a heavy app framework
- refresh-token-based mailbox operations
- scripting-friendly mail workflows on Linux servers or remote boxes

---

## Release status

Current recommended starting point: **v0.1.0**

This version is already capable of real-world mailbox tasks, including reading, sending, drafting, replying, uploading large attachments, downloading attachments, folder operations, and exporting message bodies.

---

## License

MIT
