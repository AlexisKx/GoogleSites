# NAP Data Converter — Google Apps Script + Google Sites port

Port of the Streamlit app `AlexisKx/NAP_CONVERTER_APP` to run entirely inside Google Workspace, so it can be embedded in a Google Site as required by company restrictions.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Server side. Parsing, enrichment, snapshot storage, admin auth, XLSX export. |
| `Index.html` | Single-page shell with sidebar navigation. |
| `Styles.html` | Shared CSS, included from `Index.html`. |
| `Scripts.html` | Front-end JS — page routing + all five views. |
| `appsscript.json` | Manifest. Web app deployed as `USER_ACCESSING` (so `Session.getActiveUser()` returns the real viewer) with `DOMAIN` access. |

## Pages (mapped from the Streamlit app)

| Streamlit | Apps Script | Notes |
|---|---|---|
| 📡 Converter | ✅ Converter | CSV → parse → merge dupes → enrich → save snapshot → 3-tab output (Summary / Preview / Missing Location). |
| 🕓 Data History | ✅ Data History | Lists every saved snapshot. Click *Excel* to rebuild and download. |
| 📁 GEO Reference | ✅ GEO Reference (admin) | Add/Edit (with confirmed delete), Bulk Upload, View All — same bulk format as Streamlit. |
| 👥 User Management | ✅ Admin Management (admin) | Workspace SSO handles identity. This page only manages who has admin role. |
| 🔒 Change Password | ❌ Removed | Not needed with SSO. |

## Storage

One Google Sheet holds every tab — no Supabase, no separate DB.

| Tab | Columns |
|---|---|
| `ADMINS` | email, added_at, added_by |
| `MASTER` | NAP ID, PLA ID, Tech, Territory, Area, Province, City, BRGY, Location (optional override) |
| `PLA_BY_CABINET` | Cabinet, PLA ID, Tech |
| `AREA_BY_PREFIX` | Prefix, Sales Area, Province, Territory |
| `GEO_REFERENCE` | NAP ID, CITY_NAME, BRGY_NAME, LOCATION TAGGING, updated_at, updated_by |
| `NAP_DATA` | snapshot_date, uploaded_by, + the 17 output columns |
| `SNAPSHOT_INDEX` | snapshot_date, uploaded_by, uploaded_at, row_count |

`setupSheets()` creates them all on first run.

## Deployment

1. Create a Google Sheet to hold the lookup tabs. Copy its ID (the string between `/d/` and `/edit` in the URL).
2. Go to https://script.google.com → New project.
3. Paste `Code.gs` into the code file. Paste `appsscript.json` into the manifest (Project settings → "Show appsscript.json").
4. Add three HTML files: `Index`, `Styles`, `Scripts` — paste each file's contents.
5. In `Code.gs`, set `CONFIG.LOOKUP_SPREADSHEET_ID` to the ID from step 1. Optionally set `CONFIG.SEED_ADMIN_EMAIL` to your email so you get auto-promoted.
6. From the script editor, run `setupSheets()` once. Approve the OAuth scopes.
7. Deploy → New deployment → "Web app". Execute as: **User accessing the web app**. Access: **Anyone within [your domain]**. Note the `/exec` URL.
8. In Google Sites: Insert → Embed → By URL → paste the `/exec` URL. Save and publish.

## Logic notes (carried over from the existing Code.gs in this folder)

Parsing matches the Python:
- `split(';')`, last `TRAILING_COLS = 12` fields are fixed
- coordinates kept only if they match `^-?\d{1,3}\.\d{4,}$`
- utilization is `round(assigned/total, 4)`

Enrichment uses **lookup tables** (MASTER → PLA_BY_CABINET / AREA_BY_PREFIX / GEO_REFERENCE), priority order: MASTER first, then fallback. The Streamlit handover describes PLA / Tech / Sales Area / Territory / Province as "derived"; if the Python code computes these from strings rather than looking them up, you'll want to replace the lookup with that derivation. The three open items from the previous handover (merge key, NAP prefix rule, PLA/Tech derivation) still apply — reconcile against `app.py` when you have it.

`baseNapId_()` strips a trailing `-NNN` / `_NNN` suffix before using it as the merge key. **Verify this matches `merge_duplicates()` in the Python.**

## Known limits

- Apps Script execution cap: **6 minutes**. Very large CSVs (hundreds of thousands of rows) will time out — keep the Streamlit tool for those.
- Google Sheets cap: **10M cells per spreadsheet**. At ~25k rows × 19 cols per snapshot, `NAP_DATA` fills up in ~20 daily snapshots. Rotate to a fresh spreadsheet monthly or implement a retention cleanup.
- No password flow — auth is the Google account viewing the page. Make sure the deployment's access setting matches who should see it.
- Re-uploading on the same snapshot date **overwrites** that date's rows (upsert), matching the Streamlit behavior.
