# AM_CRM Outlook VBA Helper

The browser app can open a draft with `mailto:` as a fallback, but that does not prove an email was sent.

Use this VBA helper path for future Outlook Desktop integration:

1. Open Outlook draft with prefilled To/Subject/Body.
2. Export recent email history for client email addresses.
3. Export actually sent client emails so AM_CRM can count sent emails as client contact.
4. Convert client emails into JSON task candidates for LLM analysis/import.

The `.bas` file is intentionally a starter. Codex should complete it after the browser data engine is stable.
