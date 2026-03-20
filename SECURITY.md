# Security Policy

## Supported versions

At this stage, the latest version on `main` should be considered the supported version.

## Reporting a vulnerability

Please do **not** open a public issue for sensitive security problems.

If you discover a security issue, report it privately to the repository owner first.

Suggested report content:
- summary of the issue
- impact
- reproduction steps
- proof of concept if safe to share
- suggested mitigation if known

## Sensitive data warning

This project deals with OAuth credentials and mailbox operations.

Never include any of the following in issues or pull requests:
- access tokens
- refresh tokens
- private mailbox content
- downloaded attachments containing personal data
- local config files with real secrets

Use placeholders and redacted samples only.

## Scope notes

Common risk areas include:
- token handling
- config file safety
- accidental credential logging
- unsafe attachment writes
- unsafe shell usage around exported content
