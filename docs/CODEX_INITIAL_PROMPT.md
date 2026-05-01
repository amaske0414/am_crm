You are working inside a VS Code folder named `AM_CRM_Codex_Starter`. Build the AM_CRM app in this repo.

Your job is to turn the starter scaffold into a complete, local-only, browser-based wealth management client command center. Do not ask clarifying questions. Make sensible implementation decisions and code the app directly.

Start by reading these files in order:

1. `README.md`
2. `docs/APP_SPEC.md`
3. `docs/DATA_SCHEMA_SUMMARY.md`
4. `docs/MISSING_SCHEMA_ITEMS.md`
5. `data/schema_catalog.json`
6. `data/relationship_mappings.json`
7. `data/seed_clients.json`
8. `src/app.js`
9. `styles/styles.css`

Hard constraints:

- Use only HTML, CSS, vanilla JavaScript, JSON, and optional VBA.
- No React, Vue, Angular, TypeScript build step, backend server, cloud service, database server, CDN, or external hosted API.
- The app must be usable by opening `index.html`.
- If you need SheetJS for `.xlsx` parsing, assume the local file will be placed at `lib/xlsx.full.min.js`. Add graceful detection and instructions if it is missing. Do not load it from a CDN.
- Keep all client data local. Do not send data to any external endpoint.
- Keep the visual design close to VS Code default dark theme.
- Keep the UI compact/dense.
- Every box, button, table cell, popup, card, modal, input, and rectangle must default to `0px` corner radius, controlled by a Settings value.
- The app has only two pages/views: Home and Settings.
- Home contains one table grid with exactly these columns:
  1. Status
  2. Client Name
  3. Email
  4. Next Contact
  5. T12 Revenue
  6. NNA
  7. Available Cash
  8. Total Assets
  9. Advisory Assets
  10. Asset Allocation
  11. Financial Plan
  12. Online Activity
  13. Tasks
  14. Notes

Preserve the starter’s current structure unless there is a strong reason to improve it. You may split `src/app.js` into additional vanilla JS files, but keep loading simple from `index.html`, with no bundler required.

Core implementation tasks:

## 1. Make the current starter fully stable

- Verify `index.html` opens and renders the seeded client list.
- Keep the fixed top-center search.
- Search should start only after 3+ characters and filter in real time.
- Status and Client Name must stay pinned.
- Default sort should be T12 Revenue descending. Since revenue may be missing before import, sort missing values after real values.
- All main-grid values must be imported/calculated/read-only. No inline editing.
- Clicking any cell opens the correct popup.
- Only one popup can be open at a time.
- Clicking outside a popup closes it.
- Popup mode must be configurable in Settings: floating card, right drawer, centered modal.
- Missing values display as `N/A` in dimmed text.

## 2. Implement local Excel import

Build an Excel import engine that can read a user-selected local folder or multiple selected `.xlsx` files.

Target files:

- `ref_ALG.xlsx`
- `ref_ACCOUNT.xlsx`
- `data_POSITIONS.xlsx`
- `data_ACTIVITY.xlsx`
- `data_REVENUE.xlsx`
- `ref_PRODUCT.xlsx`
- `ref_SECURITY.xlsx`
- `ref_WEALTHDESK.xlsx`
- `data_GPS.xlsx` optional/missing schema
- `data_CLIENTCONTACT.xlsx` optional/missing schema
- `data_POSITIONS_EXTERNAL.xlsx` optional/missing schema
- `ref_ALLOCATION_TARGET.xlsx` optional/missing schema

Rules:

- Source Excel files are read-only.
- Preserve leading zeroes.
- Treat join keys as text.
- Trim source values before comparisons.
- Read the first worksheet unless a mapping says otherwise.
- Validate expected column names using `data/schema_catalog.json`.
- If a file is missing, stale, or has changed columns, show a warning in Settings > Data Management. Do not crash.
- If optional files are missing, dependent columns should show `N/A`.
- Store imported raw rows in memory only, not in backups.
- After a refresh, store calculated main-grid and popup summaries as snapshots so a JSON backup can restore the visible state without immediately re-importing Excel.

## 3. Implement relationship mapping exactly from schema

Use `data/relationship_mappings.json` as the approved relationship source.

Implement helper functions:

- `normalizeTextKey(value)` = upper-case trimmed string.
- `leftAccountKey(value)` = first 10 characters of a trimmed account number string.
- `buildAccountKey(branchOrOffice, accountNumber)` = zero-padded 3-digit branch/office + `-` + zero-padded 6-digit account number.
- `toNumber(value)` safely parses numbers/currency.
- `toDate(value)` safely parses Excel dates, serial dates, and strings.

Implement these joins:

- R001: `data_REVENUE.[Cde Prd]` -> `ref_PRODUCT.[Cde Prd]`
- R002: revenue account key from `Num Alt Ofc` + `Num Acc` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R003: positions account key from `Branch Number` + `Account Number` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R004: positions account key -> first 10 chars of `ref_WEALTHDESK.[Account Number]`
- R005: first 10 chars of `data_ACTIVITY.[Account Number]` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R006: `data_POSITIONS.[Cde Msdw Sec]` -> `ref_SECURITY.[ID_MSDW_SECURITY]`
- R007: `UPPER(TRIM(ref_ACCOUNT.[Sort Name]))` -> `UPPER(TRIM(ref_ALG.[ID_NAME_ALG]))`

## 4. Implement Settings > Data Management visual relationship editor

Build a local visual mapping editor:

- Show a card for every imported/generated table.
- Each card lists fields from the schema or loaded workbook headers.
- Preload the approved relationships from `data/relationship_mappings.json`.
- Show the canonical predicate for each approved mapping.
- Allow manual add/delete/edit of relationship mappings.
- Use simple HTML/SVG for connector lines. No external library.
- Persist user mapping edits in app state and backup.
- Include validation checks:
  - Duplicate keys in lookup tables.
  - Blank keys in detail tables.
  - Detail rows with no matching lookup row.
  - Changed/missing columns.

## 5. Implement one-row-per-client-group aggregation

The main grid row key is client group ID / household.

Use this row source priority:

1. Imported `ref_ALG.[ID_NAME_ALG]`
2. Seed clients in `data/seed_clients.json`
3. Manually added clients stored in app state

Client name shown in the grid is the client group ID/name.

For each client group ID, connect:

- Accounts from `ref_ACCOUNT` by `Sort Name` -> `ID_NAME_ALG`.
- Positions from `data_POSITIONS` through account key joins.
- Activity from `data_ACTIVITY` through account key joins.
- Revenue from `data_REVENUE` through account key joins.
- Product category from `ref_PRODUCT`.
- Security classification from `ref_SECURITY`.
- Strategy data from `ref_WEALTHDESK`.

## 6. Implement all main-grid column calculations

### Status

- Display as icon only.
- Display status source priority:
  1. `ref_ALG.STATUS`
  2. seed client `STATUS`
  3. best/highest tier among linked `ref_ACCOUNT.Status`
  4. `N/A`
- Valid values: A, B, C, D.
- A = platinum circle, B = gold circle, C = silver circle, D = bronze circle.
- Settings > Icons must allow local `.png` overrides.
- Popup compares imported/display status with calculated status.
- Calculated status:
  - A if T12 Revenue >= 30000 or Total Assets >= 5000000
  - B if T12 Revenue >= 15000 or Total Assets >= 2500000
  - C if T12 Revenue >= 5000 or Total Assets >= 1000000
  - D otherwise
- Support Parent Client Group ID override:
  - Child calculated tier is max of standard calculated tier and one tier below parent.

### Client Name / Profile

Popup fields:

- Client Group ID, read-only.
- Parent Client Group ID, editable.
- Risk Profile dropdown:
  - 1 - Wealth Conservation
  - 2 - Income Oriented
  - 3 - Balanced Growth
  - 4 - Market Growth
  - 5 - Opportunistic Growth
- Include Alts? Yes/No.
- Include Munis? Yes/No.
- Total Net Worth from `ref_ACCOUNT.[Total Net Worth]`, read-only.
- Annual Income from `ref_ACCOUNT.[Total Annual Income]`, read-only.
- Minimum Available Cash, numeric, default 0.
- Maximum Available Cash, numeric, default 100000.
- Number of Reviews: 0, 1, 2.
- Review Month: Jan-Dec.
- Contact Frequency: 0, 1, 4, 6, 12.
- Static relationship notes.
- Family/relationship manager.
- Client activity history.

Default profile values based on Status:
- A: Contact Frequency 12 and Number of Reviews 2.
- B: Contact Frequency 6 and Number of Reviews 2.
- C: Contact Frequency 1 and Number of Reviews 0.
- D: Contact Frequency 0 and Number of Reviews 0.

Auto-save editable profile fields to app state.

### Email

- Show all client emails.
- Clicking Email opens Outlook draft.
- Use `mailto:` as fallback.
- Keep a VBA helper pathway for Outlook Desktop.
- Select template based on client data:
  - overdue contact
  - excess cash
  - upcoming review
  - default
- Settings > Email must support add/edit/delete templates and template variables.
- Do not log an email because a draft opened.
- Only sent email logging should happen via VBA helper/imported sent-email log.

### Next Contact

- Display days until next contact:
  - future: `+12d`
  - today: `Due`
  - overdue: `-8d`
  - request-only/missing: `N/A`
- Contact frequency:
  - 0 = request only
  - 1 = annual / 365 days
  - 4 = quarterly / 91 days
  - 6 = bi-monthly / 61 days
  - 12 = monthly / 30 days
- Last contact comes from:
  1. app logged contact interactions
  2. `data_CLIENTCONTACT.xlsx`, when available
  3. seed/client last review date
- Clicking opens interaction logger.
- Contact types:
  - Sent email
  - Sent text message
  - Zoom meeting
  - In-person meeting
  - Logged call
  - Voicemail
- Saving adds a note to Next Contact, logs activity, and recalculates Next Contact.
- Overdue contact should turn the grid text red and add/update a note in the Next Contact popup.

### T12 Revenue

- Source: `data_REVENUE.xlsx`.
- Period: last 365 days from today.
- Date field: `Dt2 Trd Entr`.
- Default amount field: `Amt Cmp Grs`.
- Add setting to switch amount field to `Amt Pyout`.
- Join revenue to accounts via R002 and products via R001.
- Aggregate to client group ID.
- Popup: pivot-like table with selectable row fields, column fields, value field, and date range.
- Suggested default:
  - Rows: product `GROUP` or `PRODUCT_L1`
  - Columns: month
  - Values: sum selected revenue amount

### NNA

- Source: `data_ACTIVITY.xlsx`.
- Period: last 365 days.
- Amount field: `Amount($)`.
- Include activity that changes balance.
- Exclude dividends defensively if Activity Type or Activity Description contains `dividend`.
- Aggregate by client group ID.
- Popup default:
  - Rows: Activity Type
  - Columns: Month
  - Values: Sum Amount($)
- Large NNA activity notes are visible only inside the NNA popup.

### Available Cash

- Source: `data_POSITIONS.xlsx`; optional external positions when available and enabled.
- Calculate at client group ID.
- Include taxable, non-IRA, non-managed positions only.
- Exclude positions where account is advisory/managed.
- Include:
  - pure cash
  - money market
  - Savings Deposit products
- Use ref_SECURITY flags when available:
  - isCASH
  - isBDPS
- Also use data_POSITIONS fields when needed:
  - Nme Iss Typ
  - Nme Iss Sub Typ
  - Nme Asset Typ
  - Txt Lvl fields
- Main-grid format: `$0.0K / 0.0%`.
- Display dollar value should be excess cash = total available cash - minimum available cash.
- Percent = excess cash / Total Assets.
- Red if raw total available cash is below min or above max.
- Popup shows:
  - raw total cash
  - min threshold
  - max threshold
  - excess cash
  - breakdown by account type IRA/Taxable
  - columns for cash, money market, savings deposits
  - securities maturing in next 30 days
  - notes
- Securities maturing in next 10 days should automatically add/update a note.
- An intentional-cash note disables cash threshold alerts until archived/deleted.

### Total Assets

- Source: `data_POSITIONS.xlsx` plus optional `data_POSITIONS_EXTERNAL.xlsx`.
- Internal market value field: `Amt Mkt Val`.
- Aggregate by client group ID.
- Settings should allow product category toggles.
- Popup:
  - Morgan Stanley total
  - external total
  - custodian/platform breakdown
  - product category breakdown
  - notes
- If Total Assets changes plus/minus 10% from prior refresh, add/update note inside Total Assets popup.

### Advisory Assets

- Source: `data_POSITIONS.xlsx`.
- Include positions where `Managed = Y`.
- Market value field: `Amt Mkt Val`.
- Percent = Advisory Assets / Total Assets.
- Main format: `$0.0MM / 0.0%`.
- Popup shows advisory assets by strategy from `ref_WEALTHDESK.[Investment Type]`.
- Include notes section.
- Include placeholder for conversion opportunity logic.

### Asset Allocation

- Use positions by market value.
- Classify using `ref_SECURITY`; missing classification = `UNKNOWN`.
- Support manual asset class mapping in Settings.
- Settings must support upload/mapping of target allocation files:
  - one with alternatives
  - one without alternatives
- Target allocation depends on profile:
  - Risk Profile
  - Include Alts?
  - Include Munis?
- Settings must support drift threshold by asset class.
- Main grid shows:
  - `OK` in green
  - `DRIFT` in red
- Popup table rows are asset classes with:
  - Current Allocation
  - Target Allocation
  - Drift
  - Threshold
  - Status
- Any single-stock position greater than 10% of total assets should add/update a note inside Asset Allocation popup.

### Financial Plan

- `data_GPS.xlsx` schema is not attached, so implement configurable mapping.
- Main grid shows Likelihood of Success.
- >= 85 green.
- 0-84 red.
- Missing `N/A` gray.
- Popup shows GPS status, likelihood, last plan update date, next review date, notes.
- No plan/outdated plan flags appear only in Financial Plan popup notes.

### Online Activity

- Current `ref_ACCOUNT` schema does not include last-login/username fields, so implement configurable mapping.
- When mapped:
  - last login <31 days: green circle
  - >=31 and <180 days: yellow circle
  - >=180 days: red circle
  - missing/never registered: `N/A` gray
- Popup shows last login date, username, notes.

### Tasks

- Tasks are managed entirely inside the app.
- Fields:
  - id
  - clientGroupId
  - title
  - dueDate
  - priority
  - status
  - notes
  - createdDate
  - completedDate
  - recurrence
  - archivedDate
  - source
- Status values:
  - Open
  - In Progress
  - Waiting
  - Completed
  - Deferred
  - Cancelled
- Priority:
  - Low
  - Normal
  - High
  - Urgent
- Main grid shows open task count.
- Popup supports add/edit/complete/archive.
- Completing task adds task note and activity log event.
- Support recurring task placeholder.

### Notes

- Every note must be tied to a client group ID and column key.
- Notes appear as mini cards.
- Support star, edit, archive/delete.
- Main grid shows notepad icon.
- Global search includes notes.
- Activity log records note actions.

## 7. Implement backup, restore, export

Backup:

- One-click JSON backup.
- Filename: `AM_CRM_backup_YYYYMMDD_HHMMSS.json`.
- Includes:
  - settings
  - table layout
  - data management mappings
  - main grid calculated snapshot
  - popup summaries
  - notes
  - tasks
  - activity log
  - email templates
  - icon overrides
  - manual clients
  - manual profile overrides
  - prior refresh snapshot
- Excludes raw imported Excel rows.
- Restore should load the saved snapshot without refreshing Excel until Refresh Data is clicked.
- Settings toggle for automatic backups.

Export:

- Export current grid to Excel when local SheetJS is available.
- Fallback to CSV if SheetJS is missing.
- Export LLM JSON to `AM_CRM_LLM_EXPORT_YYYYMMDD_HHMMSS.json`.

## 8. Implement LLM export/import

LLM export should include one object per client group ID:

- client profile summary
- imported and calculated status
- T12 revenue
- NNA
- available cash and thresholds
- total assets
- advisory assets
- asset allocation/drift
- financial plan status
- online activity
- last contact
- next contact
- open tasks
- important notes
- recent activity log
- column-specific alerts
- suggested rule-based next best actions

LLM task import should accept:

```json
{
  "generatedAt": "2026-05-01T00:00:00",
  "source": "Internal ChatGPT",
  "tasks": [
    {
      "clientGroupId": "TOM & KATHY BROWN",
      "title": "Schedule review meeting",
      "priority": "High",
      "dueDate": "2026-05-03",
      "notes": "Reasoning from model",
      "sourceColumn": "Next Contact"
    }
  ]
}
```

Create imported tasks and log activity.

## 9. Implement settings panels

Complete the Settings submenus:

- Main Table: popup mode, radius, tooltips, field source, reset layout.
- Data Management: file status, table cards, field cards, visual joins, validation.
- Status: icon config, tier rules, calculated status preview.
- Client Profile: dropdown defaults and status-driven defaults.
- Email: templates add/edit/delete, trigger rules, variables.
- Next Contact: cadence settings and contact type list.
- T12 Revenue: amount field selection and pivot defaults.
- NNA: excluded activity keyword rules.
- Available Cash: held-away toggle, cash type mappings, threshold defaults.
- Total Assets: product category toggles.
- Advisory Assets: strategy mapping and conversion-opportunity placeholder.
- Asset Allocation: target file mapping, class mapping, drift thresholds.
- Financial Plan: data_GPS field mapping.
- Online Activity: login date/username field mapping.
- Tasks: statuses/priorities and recurrence placeholder.
- Notes: note display and archive behavior.
- Icons: upload `.png` overrides for app icons/status indicators.
- Backups: backup, restore, auto-backup toggle.
- Activity Log: view/edit/delete activity entries.
- LLM Export / Import: export structured JSON, import task JSON.

## 10. Improve the VBA helper

Complete `vba/AM_CRM_Outlook_Helper.bas` enough to support:

- Opening Outlook Desktop draft with To/Subject/Body.
- Exporting recent emails for known client email addresses to a JSON file.
- Limiting to 20 most recent emails per client email address.
- Exporting sent-email log so the app can log actually sent emails.
- Never send emails automatically without explicit user action.

## 11. Manual acceptance tests

After coding, verify these manually:

1. `index.html` opens without a server and shows the seeded client list.
2. Search filters after 3 characters.
3. Clicking every column type opens a popup.
4. Notes can be added, starred, archived, searched, and backed up.
5. Tasks can be added, completed, archived, exported, and imported from LLM JSON.
6. Backup downloads JSON and restore reloads app state.
7. Data Management shows schema tables and mappings from the generated JSON.
8. Missing optional files show warnings but do not crash the app.
9. Excel import preserves leading zeroes and uses mappings R001-R007.
10. T12 Revenue, NNA, Available Cash, Total Assets, Advisory Assets, and Asset Allocation recalculate after refresh.
11. App remains local-only and contains no CDN or external API calls.
12. UI remains compact, dark, and 0px radius by default.

Deliver code, not pseudo-code. Keep the app practical, fast, and readable. Prioritize the data engine and LLM export/import because those are the most important workflows.
