# AM_CRM Codex Starter Project

This folder is a VS Code-ready starter scaffold for **AM_CRM**, a local-only client management app for a wealth manager/financial advisor.

It was generated from:

- `source_workbooks/schema.xlsx`
- `source_workbooks/clients.xlsx`

## What is already included

- `index.html` — opens the starter app.
- `styles/styles.css` — VS Code-inspired dark theme, compact table, 0px default radius.
- `src/app.js` — working starter app with:
  - Home grid using the uploaded client seed list.
  - Status icons.
  - Fixed top-center search.
  - Home/Settings navigation.
  - Popup-card framework.
  - Notes and tasks stored in browser localStorage.
  - JSON backup/export scaffold.
  - LLM export/import scaffold.
  - Data Management settings scaffold.
- `src/generated/schemaCatalog.js` — generated JavaScript global from the schema workbook.
- `src/generated/seedClients.js` — generated JavaScript global from the client list.
- `data/*.json` — implementation-ready JSON schema, metadata, relationship mappings, seed clients, app state template, and target allocation template.
- `docs/CODEX_INITIAL_PROMPT.md` — paste this into Codex inside VS Code.
- `docs/DATA_SCHEMA_SUMMARY.md` — readable schema/relationship summary.
- `vba/AM_CRM_Outlook_Helper.bas` — starter Outlook VBA helper.

## How to open

Open this folder in VS Code and open `index.html` in a browser.

The starter app does not require a server for the seeded-client view. For full Excel folder import, Codex should implement a browser file/folder picker and local SheetJS integration.

## Important local-only rule

Do not use CDN links, external APIs, cloud storage, or third-party hosted services. Any Excel parser library must be stored locally, for example:

```text
lib/xlsx.full.min.js
```

## Current generated counts

- Seed clients: `102`
- Schema tables: `8`
- Relationship mappings: `7`

## Known schema gaps Codex must handle gracefully

The attached schema workbook does not currently include schemas for:

- `data_GPS.xlsx`
- `data_CLIENTCONTACT.xlsx`
- `data_POSITIONS_EXTERNAL.xlsx`
- `ref_ALLOCATION_TARGET.xlsx`

The app should still support those files through Settings > Data Management once their headers are available. Missing files should trigger warnings, not crashes.

Also, the current `ref_ACCOUNT` schema does not show explicit last-login / username fields for Online Activity. The Online Activity column should remain `N/A` until those fields are mapped.
