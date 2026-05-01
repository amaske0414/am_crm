# Local JavaScript Library Folder

Place local third-party JavaScript files here. Do not use CDN URLs.

Recommended for Excel parsing:

```text
lib/xlsx.full.min.js
```

The app should detect whether `window.XLSX` exists. If it does not, Excel import/export-to-xlsx should display a clear local setup message and fall back to seed data/CSV export where possible.
