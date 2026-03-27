# Claude Development Instructions — M365Email

## Master Coding Standards

**All code in this repo MUST follow the master standards at:**
https://github.com/theteam123/virtual_dev_server/blob/main/CLAUDE.md

Read it completely before writing any code. The rules below are
repo-specific additions — they do not override the master.

---

## CRITICAL: This Is NOT ERPNext

This platform runs on **Frappe Framework** — NOT ERPNext.

- **Do NOT** assume any ERPNext doctypes exist
- **Do NOT** reference ERPNext modules, controllers, or utilities
- **Do NOT** import from `erpnext.*` — it is not installed
- All doctypes in this system are **custom-built** by Team Group

---

## App Identity

- **App Name:** m365email
- **Title:** M365Email
- **Publisher:** Team Group Pty Ltd
- **Description:** Microsoft 365 Email Integration

## Key Features

- Microsoft 365 email send and receive
- OAuth2 authentication with M365
- See `README_SETUP.md` and `SENDING_SETUP.md` for configuration

## Staging Note

M365 Email integration is auto-disabled after production database refresh.

## Build

```bash
bench build --app m365email
```
