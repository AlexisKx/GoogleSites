# NAP Data Converter — Google Apps Script + Google Sites port

Port of the Streamlit app `AlexisKx/NAP_CONVERTER_APP` (`app_to_xlsx1.py`), restructured so the **converter runs entirely in the browser**. The Apps Script side is now a thin auth + GEO_REFERENCE backend. This lets the app handle the user's real-world file sizes (100–700 MB) which Apps Script's own 6-min / 50 MB limits cannot.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server — auth, GEO_REFERENCE CRUD, admin allowlist. ~250 lines. |
| `Index.html` | Page shell. Loads SheetJS from CDN, includes Styles + Scripts + LookupsClient. |
| `Styles.html` | CSS. |
| `Scripts.html` | All client-side logic — streaming CSV parse, merge, enrich, XLSX build, download. |
| `LookupsClient.html` | Lookup dictionaries (PLA, tech, prefix maps) ported verbatim from the Python. Loaded into the browser as global consts. |
| `appsscript.json` | Manifest. Web app deployed as `USER_ACCESSING` with `DOMAIN` access. |

## Architecture

```
Browser                                         Apps Script
─────────                                       ───────────
File picker (HTML <input type=file>)
       │
       ▼ file.stream().pipeThrough(TextDecoderStream)
Streaming line reader (memory-safe)
       │
       ▼
Parse + Territory filter + Merge dupes
       │
       ▼  geoData (one-time fetch)        ◀──── getGeoData()  (cached 5 min)
Enrich (PLA, Tech, Area, Territory,
        Province from JS dicts;
        City/BRGY/Loc from geoData)
       │
       ▼ SheetJS XLSX.utils.aoa_to_sheet
Build .xlsx in memory
       │
       ▼ Blob → anchor download
User's Downloads folder
```

The Sheet only stores:
- `ADMINS` — email allowlist for admin pages
- `GEO_REFERENCE` — City / Barangay / Location Tagging by NAP ID

Conversions are **not saved anywhere** — the user re-runs if they need the file again.

## Pages

| Page | Notes |
|---|---|
| 📡 Converter | All-in-browser. Progress bar + live row count. 3-tab output (Summary / Preview / Missing Location). |
| 📁 GEO Reference (admin) | Add/Edit (confirmed delete), Bulk Upload (parses Excel in browser, sends rows in chunks), View All. |
| 👥 Admin Management (admin) | Workspace SSO handles identity. This page only controls the admin allowlist. |
| ℹ️ About | How it works + caveats. |

## Deployment

1. Create a Google Sheet to hold the data tabs. Copy its ID (between `/d/` and `/edit` in the URL).
2. https://script.google.com → **New project**.
3. Paste `Code.gs` into the default code file. Set `CONFIG.LOOKUP_SPREADSHEET_ID` to the ID from step 1. Set `CONFIG.SEED_ADMIN_EMAIL` to your email if you want auto-admin.
4. Add four HTML files: `Index`, `Styles`, `Scripts`, `LookupsClient`. Paste each file's contents.
5. Project Settings → tick "Show appsscript.json" → paste `appsscript.json`.
6. Run `setupSheets()` once. Approve the OAuth scopes (Sheets + your email).
7. Deploy → New deployment → "Web app". **Execute as: User accessing**. **Access: Anyone** (or "Anyone within domain" for Workspace). Note the `/exec` URL.
8. In Google Sites: Insert → Embed → By URL → paste the `/exec` URL.

> If you used the older version of this app, **delete** any old files in your script project: `Lookups.gs`, any references to `NAP_DATA` / `SNAPSHOT_INDEX` / `MASTER` / `PLA_BY_CABINET` / `AREA_BY_PREFIX` tabs in your Sheet. They're unused now. The `Drive` Advanced Service can also be removed.

## Logic notes (matches `app_to_xlsx1.py` exactly)

- **Parsing.** `split(';')`, last `TRAILING_COLS = 12` fields are fixed. **Cabinet is `tail[4]`** (not `fields[0]`). Junk header rows (e.g. "NAP facility summary report") are skipped.
- **Territory filter.** Records whose NAP prefix is not in `PREFIX_TERRITORY` are dropped. Strict prefix match requires the next char after the prefix to be `_`, `-`, a digit, or `L`.
- **Merge key.** `stripSuffix()` — strips one trailing letter that follows a digit (e.g. `DVO123A` → `DVO123`).
- **Port totals special case.** If both the existing and the incoming `Ports Total` are exactly 16, the merged value stays at 16 instead of summing.
- **Utilization.** `round(assigned / total, 4)`. If total is 0, value is 0.
- **PLA ID.** Derived from cabinet: split by `_`, try the first two parts joined, then just the first part, against `PLA_ID_LOOKUP`.
- **Tech.** Derived from cabinet: `LSA` in cabinet → GPON; `-M` in cabinet → `CABINET_TECH_LOOKUP` (default `ADSL/VDSL`); else GPON.
- **Sales Area / Province / Territory.** Derived from the leading prefix of the NAP ID.
- **City / Barangay / Location Tagging.** Read from `GEO_REFERENCE` keyed by full NAP ID. Anything missing surfaces in the Missing Location tab.

## Practical limits

| | Limit | Why |
|---|---|---|
| CSV input size | ~700 MB (desktop) | Browser RAM. Streaming the read keeps the file off the JS heap, but the parsed/merged result needs to fit. Mobile devices will choke much earlier. |
| Output XLSX size | ~200 MB | SheetJS holds the workbook in memory before writing. |
| Conversion time | unbounded | No 6-min cap because nothing runs on Apps Script. A 500 MB file takes 1–5 minutes on a modern desktop. |
| GEO_REFERENCE size | ~100K entries | Server returns the full dict to the browser on first conversion; 100K × ~150 bytes ≈ 15 MB JSON, fine. |

If the user has a 2 GB CSV someday: split it, or fall back to the original Streamlit tool — the right tool for true mega-dumps.

## Updating the lookup data

When a new cabinet or NAP prefix appears, open `LookupsClient.html` in the script editor and add an entry to the relevant dict. Save → redeploy as a new version → the browser picks it up on next load.
