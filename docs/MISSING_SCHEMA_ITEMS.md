# Missing or Future Schema Items

The uploaded schema workbook covers:

- ref_ALG
- ref_ACCOUNT
- data_POSITIONS
- data_ACTIVITY
- data_REVENUE
- ref_PRODUCT
- ref_SECURITY
- ref_WEALTHDESK
- METADATA
- MAPPING

The broader AM_CRM spec also references these future/optional files:

- data_GPS.xlsx
- data_CLIENTCONTACT.xlsx
- data_POSITIONS_EXTERNAL.xlsx
- ref_ALLOCATION_TARGET.xlsx

Implementation rule:

1. The app must never fail if an expected optional file is missing.
2. Settings > Data Management must show missing/stale/changed-file warnings.
3. Once a future file is present, the app should read headers and allow mapping.
4. Calculations depending on missing files should display `N/A` with dimmed text.
5. Popups should include implementation notes and notes/tasks functionality even while source data is missing.

Online Activity note:

The user spec says `ref_ACCOUNT.xlsx` should provide last login date and username, but the uploaded `ref_ACCOUNT` schema does not currently include those columns. Build Online Activity with configurable field mapping rather than hardcoding nonexistent column names.
