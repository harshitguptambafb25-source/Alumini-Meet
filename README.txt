# GLA Alumni Get-Together 2026 — Dashboard deployment

## Included
- `Code.gs`: Google Apps Script backend.
- `Index.html`: Premium responsive web dashboard.
- `appsscript.json`: Apps Script manifest.
- Uses your provided spreadsheet IDs and tab IDs (gid values).
- Creates a `Check-in Log` tab in the master spreadsheet when first check-in is recorded.

## Before deploying
1. Open the master spreadsheet and confirm the tab in the URL has gid `1277655877`.
2. Open the Form responses spreadsheet and confirm the tab in the URL has gid `904347459`.
3. In the Google account that owns/has edit access to the master sheet, visit `script.google.com` and create a new project.
4. Add `Code.gs` and `Index.html` from this bundle.
5. In `Code.gs`, replace `REPLACE_WITH_YOUR_EMAIL` in `ADMIN_EMAILS` with your exact Google account email.
6. Save. Run `getDashboardData` once from the Apps Script editor and approve the requested Google Sheets access. Inspect any error.
7. If key fields are not detected, update `findKeyIndex_()` so the correct unique identifier column is selected. The script tries to identify common labels like Roll Number, Alumni ID, Email, or Phone; header names differ between forms, so this must be checked against your actual sheets.
8. Deploy > New deployment > Web app.
   - Execute as: **Me**
   - Who has access: **Only myself** (recommended for initial testing), or a restricted organization audience if your Workspace administrator permits it.
9. Open the deployment URL while signed into the authorized account.

## Important security notes
- Do NOT deploy as "Anyone" while personal alumni details are returned by the backend.
- Do not paste Google passwords or OAuth tokens into code.
- Apps Script authorization is generally a one-time approval per user/script/scopes, but Google may ask again after permission changes, scope changes, account changes, or security events.
- The code checks `Session.getActiveUser().getEmail()` against `ADMIN_EMAILS`. For consumer accounts, the email may be blank depending on deployment/access mode; use a restricted deployment and test. If it is blank, don't remove access controls and publish publicly; adjust to an organization-restricted deployment or add a secure authentication approach.
- Google Forms file-upload answers often appear as Drive URLs. The dashboard opens the URL if the viewer has permission; it cannot bypass Drive sharing restrictions. Keep uploaded photos restricted to the authorized team.
- The app refreshes every 60 seconds. Apps Script quotas, execution time, and concurrent use may affect this; this is near-live polling, not true push streaming.
- `Check-in Log` records identifier, name, timestamp, active user email when available, and notes. Confirm identity before clicking check-in.
- Data is read from the exact configured tabs by sheet ID (gid), not necessarily the first tab.

## Known starter limitations to validate against the real headers
- Column headers and identifiers have not been inspected from the linked spreadsheets in this environment. You must verify the auto-detected key fields before using this operationally.
- Current mismatch comparison compares fields only where normalized header names are identical. If headers differ (e.g., "Name" vs "Full Name"), map them explicitly before relying on mismatch counts.
- Timeline uses the timestamp column if one is present.
- Excel export currently provides a CSV export that opens in Excel; print/save-to-PDF is available in the browser. Native XLSX/PDF generation can be added after schema validation.
- Pass generation is intentionally deferred, as requested.
