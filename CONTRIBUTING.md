# Contributing

Thanks for your interest in contributing.

## How to contribute

### 1. Open an issue first when appropriate

Please open an issue for:
- bug reports
- feature requests
- design discussions
- behavior changes that may affect existing users

For small typo or documentation fixes, a direct pull request is usually fine.

## Development setup

```bash
git clone https://github.com/XXiuMing/outlook-oauth-mail-helper.git
cd outlook-oauth-mail-helper
pip install -r requirements.txt
python -m py_compile outlook_oauth_mail.py
```

## Before opening a PR

Please make sure:
- the script still runs
- syntax check passes
- no real tokens or private credentials are committed
- README / docs are updated if behavior changes

## Coding style

This project currently prefers:
- simple Python
- minimal dependencies
- explicit behavior over framework abstraction
- practical CLI-first design

## Security-sensitive rule

Do not submit:
- real OAuth tokens
- refresh tokens
- personal mailbox dumps
- private configs

Use placeholders only.

## Pull request notes

A good PR should explain:
- what changed
- why it changed
- how it was tested
- whether it affects backward compatibility
