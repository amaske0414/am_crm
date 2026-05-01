# AM_CRM Implementation Spec

## Product Goal

AM_CRM is a local-only, browser-based client management app for a wealth manager/financial advisor with roughly 100 client group IDs and 600-800 accounts. It should consolidate Excel-based client, account, holding, activity, revenue, planning, and strategy data into one dense Home grid that can be exported to JSON for LLM analysis and later re-import LLM-generated tasks.

## Non-Negotiable Constraints

- Local-only.
- Open from `index.html`.
- Use HTML, CSS, vanilla JavaScript, and optional VBA only.
- No React, Angular, Vue, backend server, cloud sync, external APIs, or CDN dependencies.
- Excel source files are read-only imports.
- App-created state is stored as JSON/local browser state and included in one-click JSON backup.
- Default visual style should generally match VS Code's default dark theme.
- Default corner radius for every rectangle/box/button/card/control is `0px`.
- Only two app pages/views:
  - Home
  - Settings

## Home Grid Columns

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

## Main Grid Behavior

- One row per client group ID / household.
- Default sort: T12 Revenue descending.
- Fixed top-center search box.
- Search starts after at least 3 characters and filters in real time.
- Status and Client Name columns are pinned.
- No inline editing in the grid.
- Clicking any cell opens a popup card for that column/client.
- Missing data displays as `N/A` with dimmed font.
- Numeric columns show totals where meaningful.
- Last data refresh appears bottom-right in small gray text.
- Settings can reset column widths/layout.

## Settings Page

Settings has a left submenu:

- Main Table
- Data Management
- Status
- Client Profile
- Email
- Next Contact
- T12 Revenue
- NNA
- Available Cash
- Total Assets
- Advisory Assets
- Asset Allocation
- Financial Plan
- Online Activity
- Tasks
- Notes
- Icons
- Backups
- Activity Log
- LLM Export / Import

## Data Management

Settings > Data Management must become a visual data relationship editor:

- Card for each imported table.
- Each card lists column names.
- User can create/edit links between columns.
- Persist mappings in app state.
- Preload mappings from `data/relationship_mappings.json`.
- Show missing, stale, and changed-column warnings.
- Preserve leading zeroes in account numbers and keys.
- Treat join keys as text.
- Trim values before comparison.
- Use left joins when enriching fact/detail tables.

## Relationship Mapping Source of Truth

Use the generated file:

```text
data/relationship_mappings.json
```

The canonical account key is `###-######`.

Important mappings:

- R001: `data_REVENUE.[Cde Prd]` -> `ref_PRODUCT.[Cde Prd]`
- R002: revenue account key from `Num Alt Ofc` + `Num Acc` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R003: positions account key from `Branch Number` + `Account Number` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R004: positions account key -> first 10 chars of `ref_WEALTHDESK.[Account Number]`
- R005: first 10 chars of `data_ACTIVITY.[Account Number]` -> first 10 chars of `ref_ACCOUNT.[Account Number]`
- R006: `data_POSITIONS.[Cde Msdw Sec]` -> `ref_SECURITY.[ID_MSDW_SECURITY]`
- R007: `UPPER(TRIM(ref_ACCOUNT.[Sort Name]))` -> `UPPER(TRIM(ref_ALG.[ID_NAME_ALG]))`

## Column Logic Summary

### Status

- Display as icon only.
- Display status should come from `ref_ALG.STATUS` if available; otherwise use the uploaded client seed `STATUS`; otherwise derive from the highest/best account-level `ref_ACCOUNT.Status` tied to the client group.
- Valid values: A, B, C, D.
- A = platinum circle, B = gold circle, C = silver circle, D = bronze circle.
- Popup compares imported/display status with calculated status.
- Calculated status:
  - A if Revenue >= 30000 or Total Assets >= 5000000
  - B if Revenue >= 15000 or Total Assets >= 2500000
  - C if Revenue >= 5000 or Total Assets >= 1000000
  - D otherwise
- Client profile can assign Parent Client Group ID. Child tier is max of standard calculated tier and one tier below parent.

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
- Total Net Worth, imported/read-only.
- Annual Income, imported/read-only.
- Minimum Available Cash, editable numeric, default 0.
- Maximum Available Cash, editable numeric, default 100000.
- Number of Reviews: 0, 1, 2.
- Review Month: Jan-Dec.
- Contact Frequency: 0, 1, 4, 6, 12.
- Static relationship notes.
- Family/relationship tracking.
- Activity history for client group ID.

Status-driven defaults:

- A: Monthly contact, semi-annual reviews.
- B: Bi-monthly contact, semi-annual reviews.
- C: Annual contact, no reviews.
- D: Request-only contact, no reviews.

### Email

- Show multiple email addresses.
- Clicking opens Outlook Desktop draft.
- Starter can use `mailto`; complete version should use VBA helper where possible.
- Select email template based on client data.
- Do not log email just because draft opened.
- Log only actually sent emails, through the VBA helper.

### Next Contact

- Display days until next contact.
- Calculation: last contact date + cadence days.
- Contact frequency:
  - 0 = request only
  - 1 = annual
  - 4 = quarterly
  - 6 = bi-monthly
  - 12 = monthly
- Counts as contact:
  - Sent email
  - Sent text message
  - Zoom meeting
  - In-person meeting
  - Logged call
  - Voicemail
- Clicking opens interaction logger.
- Saving an interaction adds a note to the Next Contact popup and an activity log entry.

### T12 Revenue

- Calculate from `data_REVENUE.xlsx` over last 365 days.
- Default revenue amount field should be `Amt Cmp Grs`, with a setting to switch to `Amt Pyout`.
- Join to `ref_ACCOUNT` and `ref_PRODUCT`.
- Popup should behave like a lightweight pivot table with selectable rows/columns/value/date filters.

### NNA

- Calculate from `data_ACTIVITY.xlsx` over last 365 days.
- Use `Amount($)`.
- Include balance-changing activity.
- Exclude dividends defensively using Activity Type/Description matching.
- Popup defaults to Activity Type by month.

### Available Cash

- Calculate at client group ID level.
- Use taxable, non-IRA, non-managed/non-advisory positions.
- Include pure cash, money market funds, and Savings Deposit products.
- Include/exclude external held-away cash through Settings.
- Main display format: `$0.0K / 0.0%`.
- Use min/max thresholds from profile.
- Red if total available cash is below min or above max.
- Popup shows breakdown by IRA/Taxable and cash type.
- Popup lists securities maturing in next 30 days.
- Notes alert for maturities in next 10 days.
- Intentional cash note suppresses threshold alerts.

### Total Assets

- Calculate from `data_POSITIONS.xlsx` + optional `data_POSITIONS_EXTERNAL.xlsx`.
- Include all product categories unless disabled in Settings.
- Popup shows Morgan Stanley vs external and custodian/platform breakdown.
- Add note if assets change +/- 10% from prior refresh.

### Advisory Assets

- Calculate all positions where `data_POSITIONS.Managed = Y`.
- Display format: `$0.0MM / 0.0%`.
- Percent = Advisory Assets / Total Assets.
- Popup shows advisory assets by strategy using `ref_WEALTHDESK.Investment Type`.
- Include conversion-opportunity placeholder and notes.

### Asset Allocation

- Main grid shows `OK` or `DRIFT`.
- Classify positions using `ref_SECURITY.xlsx`; missing classification = `UNKNOWN`.
- Compare to uploaded/mapped target allocation files.
- Manual drift thresholds by asset class.
- Popup shows Asset Class, Current Allocation, Target Allocation, Drift, Threshold, Status.
- Flag single-stock position > 10% of total assets as popup note.

### Financial Plan

- Source file `data_GPS.xlsx` is not in current schema; build configurable mapping.
- Main grid shows Likelihood of Success.
- >= 85 green, 0-84 red, missing N/A gray.
- Popup shows GPS status, likelihood, last update, next review, notes.
- Add note for no plan/outdated plan.

### Online Activity

- Current schema lacks explicit last-login/username fields.
- Build configurable mapping.
- Main grid displays status circle when mapped:
  - <31 days green
  - >=31 and <180 days yellow
  - >=180 days red
  - missing/never registered N/A gray
- Popup shows last login date, username, and notes.

### Tasks

- Managed inside AM_CRM only.
- Fields: title, due date, priority, status, notes, created date, completed date, recurrence.
- Statuses: Open, In Progress, Waiting, Completed, Deferred, Cancelled.
- Priorities: Low, Normal, High, Urgent.
- Main grid shows open task count.
- Popup supports add/edit/complete/archive.
- Completing a task adds task note and activity log entry.
- LLM import can create tasks.

### Notes

- Notes are always tied to a client group ID and column key.
- Notes appear as mini cards in popups.
- Support star/archive/edit/delete.
- Main grid shows notepad icon.
- Global search includes notes.
- Add/edit/archive actions are logged.

## Backup / Restore

- One-click JSON backup.
- Backup includes:
  - Settings
  - Table layout
  - Main grid snapshot
  - Popup summaries
  - Notes
  - Tasks
  - Activity log
  - Email templates
  - Icon overrides
  - Manual clients
  - Manual profile overrides
  - Prior refresh snapshot
- Backup excludes raw imported Excel rows.
- Restoring backup should show saved snapshot and not refresh Excel data until Refresh Data is clicked.

## LLM Export / Import

Export JSON should include each client group ID with:

- Client profile summary
- Imported and calculated status
- T12 revenue
- NNA
- Available cash and thresholds
- Total assets
- Advisory assets
- Allocation/drift
- Plan status
- Online activity
- Last/next contact
- Open tasks
- Notes
- Alerts
- Recommended next action prompt

Import JSON should create tasks from a structure like:

```json
{
  "tasks": [
    {
      "clientGroupId": "CLIENT GROUP",
      "title": "Review excess cash",
      "priority": "High",
      "dueDate": "2026-05-03",
      "notes": "Reasoning from LLM",
      "sourceColumn": "Available Cash"
    }
  ]
}
```

## Outlook VBA Helper

The JS app can open drafts with mailto initially. The full implementation should include a VBA helper for:

- Opening Outlook desktop draft windows.
- Exporting last 20 recent client emails by email address.
- Logging actually sent emails.
- Supporting email-to-task workflows.
