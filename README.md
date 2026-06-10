# NAP Data Converter — Google Apps Script + Google Sites port

Port of the Streamlit app `AlexisKx/NAP_CONVERTER_APP` (`app_to_xlsx1.py`). Browser-first architecture for the converter (handles 100–700 MB CSVs), with classic username/password authentication so end users never see Google's "Unverified app" warning.

## Files

| File | What it is |
|---|---|
| `Code.gs` | Server — auth (login/logout/sessions), user management, GEO_REFERENCE CRUD. |
| `Index.html` | Page shell + page templates (login, converter, geo, users, password, about). |
| `Styles.html` | CSS. |
| `Scripts.html` | All client-side logic — login flow, streaming CSV parse, merge, enrich, XLSX build, download. |
| `LookupsClient.html` | Lookup dictionaries (PLA, tech, prefix maps) bundled into the page. |
| `appsscript.json` | Manifest. Web app deployed as `USER_DEPLOYING` (you) + `ANYONE_ANONYMOUS` access. |

## Authentication

Username + password. No Google sign-in required from end users.

- Web app deployed as **Execute as: Me (you, the deployer)** → users do NOT see the OAuth warning.
- Access set to **Anyone** → anyone with the `/exec` URL can open the login page.
- The login page calls `login(username, password)`, server returns a session token (256-bit random hex) with a 24h TTL.
- Token is stored in the browser's `localStorage`. Subsequent calls pass the token; the server validates against the `USER_SESSIONS` tab.
- Passwords are stored hashed (iterated SHA-256 + per-user salt; 1000 iterations).
- Changing your password invalidates all your other sessions.
- Admins can reset other users' passwords (which signs them out everywhere) and revoke admin / delete accounts.

## Storage tabs (auto-created by `setupSheets()`)

| Tab | Columns |
|---|---|
| `USERS` | username, password_hash, salt, is_admin, created_at, created_by, last_login |
| `USER_SESSIONS` | token, username, created_at, expires_at |
| `GEO_REFERENCE` | NAP ID, CITY_NAME, BRGY_NAME, LOCATION TAGGING, updated_at, updated_by |

Expired sessions are cleaned up on each login (cheap, runs once per signin).

## Pages

| Page | Who sees it |
|---|---|
| 🔐 Sign in | Anyone (before authenticating) |
| 📡 Converter | Any signed-in user |
| 📁 GEO Reference | Admins only |
| 👥 Users | Admins only — add users, reset passwords, toggle admin, delete |
| 🔒 Change Password | Any signed-in user |
| ℹ️ About | Any signed-in user |

## Deployment

1. Create a Google Sheet for the data tabs. Copy its ID (between `/d/` and `/edit`).
2. https://script.google.com → **New project**.
3. Paste `Code.gs`. Set `CONFIG.LOOKUP_SPREADSHEET_ID` to the Sheet ID. Optionally change `CONFIG.SEED_ADMIN.username` / `password` (default `admin` / `changeme123`).
4. Add four HTML files: `Index`, `Styles`, `Scripts`, `LookupsClient`. Paste each file's contents.
5. Project Settings → tick "Show appsscript.json" → paste `appsscript.json`.
6. Run `setupSheets()` once. Approve the OAuth scopes (Sheets only — no user identity needed). On the first run, it creates the seed admin and prints the credentials.
7. **Deploy → New deployment → Web app**.
   - **Execute as: Me** (the deployer).
   - **Who has access: Anyone**.
   - Note the `/exec` URL.
8. Open `/exec` → log in with the seed admin credentials → **immediately go to Change Password** and pick a real password.
9. Go to **Users** → create accounts for your supervisor / teammates. Give them admin if needed.
10. In Google Sites: Insert → Embed → By URL → paste the `/exec` URL.

## Bootstrapping a forgotten password

If you (the only admin) forget your password:

1. Open the lookup Sheet → `USERS` tab.
2. Delete your row (or all rows).
3. Re-run `setupSheets()` from the script editor. The seed admin is recreated with the credentials from `CONFIG.SEED_ADMIN`.

## Logic notes (matches `app_to_xlsx1.py` exactly)

- **Parsing.** `split(';')`, last `TRAILING_COLS = 12` fields are fixed. Cabinet is `tail[4]`. Junk header rows skipped.
- **Territory filter.** Records whose NAP prefix isn't in `PREFIX_TERRITORY` are dropped. Prefix match requires the next char to be `_`, `-`, a digit, or `L`.
- **Merge key.** `stripSuffix()` — strips one trailing letter after a digit (`DVO123A` → `DVO123`).
- **Port totals special case.** If both existing and incoming `Ports Total` are exactly 16, the merged value stays at 16.
- **PLA ID / Tech / Sales Area / Territory / Province** all derived from cabinet / NAP ID via the lookup dicts.
- **City / Barangay / Location Tagging** from `GEO_REFERENCE` keyed by full NAP ID.

## Practical limits

| | Limit | Why |
|---|---|---|
| CSV input size | ~700 MB (desktop) | Browser RAM. Streaming the read keeps the file off the JS heap, but the parsed/merged result needs to fit. |
| Output XLSX size | ~200 MB | ExcelJS holds the workbook in memory before writing. |
| Conversion time | unbounded | No 6-min cap because nothing runs on Apps Script. |
| GEO_REFERENCE size | ~100K entries | Server returns the full dict to the browser; cached server-side for 5 minutes. |
| Session length | 24 hours | TTL on the token row in `USER_SESSIONS`. |

## Security notes

- **You own password security.** Apps Script doesn't provide bcrypt; we use iterated SHA-256 with per-user salt (1000 iterations). Strong enough for an internal tool with a controlled user base; not a public-internet auth system.
- Anyone with edit access to the lookup Sheet can read password hashes. Limit Sheet sharing to script administrators only.
- HTTPS is automatic (Apps Script web apps are always HTTPS).
- No password complexity rules besides "≥ 8 characters". Add stronger validation in `login` / `adminCreateUser` / `changeMyPassword` if your org requires it.
