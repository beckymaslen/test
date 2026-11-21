# Becky's Husband Screening Process™

- `index.html` – the questionnaire people fill in.
- `inbox.html` – Becky's private inbox (needs the inbox key).
- `apps-script/Code.gs` – the script inside the Google Sheet that saves applications.

To connect saving, put the Google Apps Script web app URL (ends in `/exec`) into `const ENDPOINT = "";` in both `index.html` and `inbox.html`.
