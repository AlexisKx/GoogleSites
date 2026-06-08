# NAP Data Converter — Google Apps Script + Google Sites port

Port of the Streamlit app `AlexisKx/NAP_CONVERTER_APP` (`app_to_xlsx1.py`) to run entirely inside Google Workspace, so it can be embedded in a Google Site as required by company restrictions.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server — parsing, enrichment, snapshot storage, admin auth, XLSX export. |
| `Lookups.gs` | All hardcoded lookup tables (PLA IDs, cabinet → tech, prefix → area/province/territory). Ported verbatim from the Python. |
| `Index.html` | Single-page shell with sidebar navigation. |
| `Styles.html` | Shared CSS, included from `Index.html`. |
| `Scripts.html` | Front-end JS — page routing + all five views. |
| `appsscript.json` | Manifest. Web app deployed as `USER_ACCESSING` with `DOMAIN` access. |

## Pages (mapped from the Streamlit app)

| Streamlit | Apps Script | Notes |
|---|---|---|
| 📡 Converter | ✅ Converter | CSV → parse → merge dupes → enrich → save snapshot → 3-tab output. |
| 🕓 Data History | ✅ Data History | Lists every saved snapshot. Click *Excel* to rebuild and download. |
| 📁 GEO Reference | ✅ GEO Reference (admin) | Add/Edit (with confirmed delete), Bulk Upload, View All. |
| 👥 User Management | ✅ Admin Management (admin) | Workspace SSO handles identity. This page only manages the admin allowlist. |
| 🔒 Change Password | ❌ Removed | Not needed with SSO. |

## Storage

One Google Sheet holds every tab — no Supabase, no external DB.

| Tab | What's in it |
|---|---|
| `ADMINS` | email allowlist for admin-only pages |
| `GEO_REFERENCE` | NAP ID → CITY_NAME / BRGY_NAME / LOCATION TAGGING overrides |
| `NAP_DATA` | every saved snapshot, one row per NAP per date |
| `SNAPSHOT_INDEX` | one row per saved snapshot, for the History page |

`setupSheets()` creates them on first run.

> The old PLA / Tech / Area / Province / Territory lookups that used to require their own tabs are now hardcoded in `Lookups.gs`, matching the Python. There's nothing for you to fill in by hand for those.

## Deployment

1. Create a Google Sheet for the data tabs. Copy its ID (the string between `/d/` and `/edit` in the URL).
2. https://script.google.com → **New project**.
3. Paste `Code.gs` into the default code file. Add another file (➕ → Script) named `Lookups`, paste in `Lookups.gs`.
4. Project Settings → tick "Show appsscript.json" → paste `appsscript.json`.
5. Add three HTML files: `Index`, `Styles`, `Scripts` — paste each file's contents.
6. In `Code.gs`, set `CONFIG.LOOKUP_SPREADSHEET_ID` to the ID from step 1. Optionally set `CONFIG.SEED_ADMIN_EMAIL` to your email.
7. Services (left sidebar) → add **Drive API** (`Drive`, v2) — needed for bulk GEO upload.
8. Run `setupSheets()` once. Approve the OAuth scopes.
9. Deploy → New deployment → "Web app". **Execute as: User accessing**. **Access: Anyone within [your domain]** (or "Anyone" for personal Gmail). Note the `/exec` URL.
10. In Google Sites: Insert → Embed → By URL → paste the `/exec` URL.

## Logic notes (matches `app_to_xlsx1.py` exactly)

- **Parsing.** `split(';')`, last `TRAILING_COLS = 12` fields are fixed. **Cabinet is `tail[4]`** (not `fields[0]`). Junk header rows (e.g. "NAP facility summary report") are skipped.
- **Territory filter.** Records whose NAP prefix is not in `PREFIX_TERRITORY` are dropped (counted as "skipped"). The strict prefix match requires the next char after the prefix to be `_`, `-`, a digit, or `L`.
- **Merge key.** `stripSuffix_()` — strips one trailing letter that follows a digit (e.g. `DVO123A` → `DVO123`).
- **Port totals special case.** If both the existing and the incoming `Ports Total` are exactly 16, the merged value stays at 16 instead of summing.
- **Utilization.** `round(assigned / total, 4)`. If total is 0, value is 0.
- **PLA ID.** Derived from cabinet: split by `_`, try the first two parts joined, then just the first part, against `PLA_ID_LOOKUP`.
- **Tech.** Derived from cabinet: `LSA` in cabinet → GPON; `-M` in cabinet → `CABINET_TECH_LOOKUP` (default `ADSL/VDSL`); else GPON.
- **Sales Area / Province / Territory.** Derived from the leading prefix of the NAP ID via `NAP_AREA_LOOKUP` / `NAP_PROVINCE_LOOKUP` / `PREFIX_TERRITORY`.
- **City / Barangay / Location Tagging.** Read from the `GEO_REFERENCE` tab keyed by full NAP ID. Anything missing surfaces in the Missing Location tab.

## Known limits

- Apps Script execution cap: **6 minutes**. Very large CSVs (hundreds of thousands of rows) will time out — keep the Streamlit tool for those.
- Google Sheets cap: **10M cells per spreadsheet**. At ~25k rows × 19 cols per snapshot, `NAP_DATA` fills up in ~20 daily snapshots. Rotate to a fresh spreadsheet monthly or implement a retention cleanup.
- Auth is the Google account viewing the page. No bcrypt logins. Make sure the deployment access setting matches who should see it.
- Re-uploading on the same snapshot date **overwrites** that date's rows, matching the Streamlit behavior.

## Updating the lookup data

When a new cabinet or NAP prefix appears, open `Lookups.gs` in the script editor and add an entry to the relevant dict. No data migration needed — these are plain JS object literals.
