# Manual Acceptance Test Plan

Run these checks after Codex implements the app.

## Startup

- Open `index.html`.
- App loads with no console errors.
- Home grid appears.
- Client list has seeded client group IDs.
- Settings page is accessible from left sidebar.

## Home Grid

- Columns appear in this exact order:
  Status, Client Name, Email, Next Contact, T12 Revenue, NNA, Available Cash, Total Assets, Advisory Assets, Asset Allocation, Financial Plan, Online Activity, Tasks, Notes.
- Status and Client Name remain pinned when horizontally scrolling.
- Search input is fixed top-center.
- Search does nothing until 3 characters are entered.
- Search filters in real time after 3+ characters.
- Default sort is T12 Revenue descending when values exist.
- Missing values show `N/A` in dimmed text.

## Popups

- Clicking every column opens a popup.
- Only one popup remains open at a time.
- Clicking outside popup closes it.
- Popup mode can be changed in Settings.

## Notes

- Add note from each popup type.
- Note is linked to client group ID and column key.
- Star note.
- Archive note.
- Search finds note text.
- Activity log records note actions.

## Tasks

- Add task from Tasks popup.
- Complete task.
- Archive task.
- Open task count updates in grid.
- LLM task import creates tasks.

## Data Import

- Select/import local `.xlsx` files.
- Missing optional files do not crash the app.
- Changed columns show warnings.
- Leading zeroes are preserved.
- Relationship mappings R001-R007 are applied.
- Refresh recalculates main grid.

## Calculations

- T12 Revenue aggregates by client group over last 365 days.
- NNA aggregates by client group over last 365 days and excludes dividends.
- Available Cash respects taxable/non-managed/cash-equivalent rules.
- Total Assets aggregates position values.
- Advisory Assets uses `Managed = Y`.
- Asset Allocation classifies through `ref_SECURITY` and flags drift.
- Financial Plan shows `N/A` until `data_GPS` mapping is configured.
- Online Activity shows `N/A` until login field mapping is configured.

## Backup / Restore

- JSON backup downloads with timestamp filename.
- Backup excludes raw Excel rows.
- Restore reloads notes/tasks/settings/grid snapshot.
- Restore does not refresh Excel data until Refresh Data is clicked.

## Local-Only Verification

- No CDN URLs in HTML/CSS/JS.
- No cloud API calls.
- No external data transmission.
- App can run from local files.
