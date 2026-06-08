/**
 * NAP Data Converter — Google Apps Script port of app_to_xlsx1.py.
 *
 * Parsing, merge, and enrichment match the Python exactly:
 *   - parseRow reads cabinet from tail[4], not fields[0]
 *   - PLA ID + Tech are derived from the cabinet string (not looked up in a tab)
 *   - Sales Area / Province / Territory are derived from NAP-ID prefix
 *   - merge key = strip_suffix(nap_id) — one trailing letter after a digit
 *   - junk header rows are skipped
 *   - records outside Territory-7 prefixes are filtered out
 *   - ports_total special case: two 16s stay 16, otherwise summed
 *
 * Big lookup dictionaries (CABINET_TECH_LOOKUP, PLA_ID_LOOKUP,
 * NAP_AREA_LOOKUP, NAP_PROVINCE_LOOKUP, PREFIX_TERRITORY) live in
 * Lookups.gs.
 *
 * Storage: one Google Sheet ("Lookup Spreadsheet") with these tabs:
 *   ADMINS, GEO_REFERENCE, NAP_DATA, SNAPSHOT_INDEX.
 *
 * Auth: Workspace SSO via Session.getActiveUser(). The ADMINS tab is the
 * allowlist for admin-only pages.
 *
 * Run setupSheets() once from the script editor to create the tabs.
 */

// ============================ CONFIG ============================
const CONFIG = {
  LOOKUP_SPREADSHEET_ID: 'PASTE_YOUR_LOOKUP_SPREADSHEET_ID_HERE',

  TABS: {
    ADMINS:         'ADMINS',
    GEO_REFERENCE:  'GEO_REFERENCE',
    NAP_DATA:       'NAP_DATA',
    SNAPSHOT_INDEX: 'SNAPSHOT_INDEX'
  },

  TRAILING_COLS:       12,
  OUTPUT_FOLDER_NAME:  'NAP Converter Output',
  GEO_CACHE_SECONDS:   300,
  SEED_ADMIN_EMAIL:    ''   // optional: email auto-promoted to admin on first setupSheets()
};

const OUTPUT_COLS = ['Cabinet', 'NAP ID', 'Discovered When', 'PLA ID', 'Tech',
  'Ports Assigned', 'Ports Reserved', 'Ports Total', 'UTILIZATION',
  'Latitude', 'Longitude', 'SALES_AREA', 'TERRITORY', 'BRGY_NAME',
  'CITY_NAME', 'PROVINCE_NAME', 'LOCATION TAGGING'];

const SNAPSHOT_COLS = ['snapshot_date', 'uploaded_by'].concat(OUTPUT_COLS);
const SNAPSHOT_INDEX_COLS = ['snapshot_date', 'uploaded_by', 'uploaded_at', 'row_count'];
const GEO_HEADERS = ['NAP ID', 'CITY_NAME', 'BRGY_NAME', 'LOCATION TAGGING', 'updated_at', 'updated_by'];
const ADMIN_HEADERS = ['email', 'added_at', 'added_by'];

// Junk header lines (output from the source system) to skip.
const JUNK_PATTERNS = [
  /^\s*nap facility summary report/i,
  /^\s*object\s*:/i,
  /^\s*specified report/i,
  /^\s*nap name pattern/i,
  /^\s*report results/i,
  /^\s*\d+\s+rows?\s+are\s+displayed/i,
  /^\s*location\s*$/i
];

// ============================ WEB APP ENTRY ============================
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('NAP Data Converter')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ============================ AUTH ============================
function getSession() {
  const email = Session.getActiveUser().getEmail() || '';
  const admin = email && isAdmin_(email);
  return { email: email, isAdmin: admin };
}

function isAdmin_(email) {
  if (!email) return false;
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.ADMINS);
  if (!sh) return false;
  const last = sh.getLastRow();
  if (last < 2) return false;
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === email.toLowerCase()) return true;
  }
  return false;
}

function requireAdmin_() {
  const s = getSession();
  if (!s.isAdmin) throw new Error('Admin access required.');
  return s;
}

// ============================ HELPERS ============================
function s_(x) { return String(x == null ? '' : x).trim(); }

function toInt_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? '' : n;
}

function isInt_(v) { return typeof v === 'number' && Math.floor(v) === v; }

function openLookupSheet_() {
  if (!CONFIG.LOOKUP_SPREADSHEET_ID || CONFIG.LOOKUP_SPREADSHEET_ID.indexOf('PASTE_') === 0) {
    throw new Error('CONFIG.LOOKUP_SPREADSHEET_ID is not set. Edit Code.gs.');
  }
  return SpreadsheetApp.openById(CONFIG.LOOKUP_SPREADSHEET_ID);
}

// ============================ PARSING (matches Python) ============================
function isJunkRow_(raw) {
  const firstField = raw.split(';')[0].trim();
  for (let i = 0; i < JUNK_PATTERNS.length; i++) {
    if (JUNK_PATTERNS[i].test(firstField)) return true;
  }
  return false;
}

function parseRaw_(raw) {
  const fields = raw.split(';');
  const n = fields.length;
  if (n < CONFIG.TRAILING_COLS + 6) return null;
  const tail = fields.slice(n - CONFIG.TRAILING_COLS);
  return {
    cabinet:        (tail[4]    || '').trim(),
    napId:          (fields[1]  || '').trim(),
    status:         (fields[2]  || '').trim(),
    lat:            (tail[0]    || '').trim(),
    lon:            (tail[1]    || '').trim(),
    discovered:     (tail[2]    || '').trim(),
    portsTotal:     (tail[6]    || '').trim(),
    portsAssigned:  (tail[7]    || '').trim(),
    portsReserved:  (tail[8]    || '').trim()
  };
}

// Coordinates: Python trims and returns as a string; no regex validation.
function toCoord_(v) { return s_(v); }

// strip_suffix: remove one trailing letter that follows a digit (e.g. DVO123A → DVO123).
function stripSuffix_(napId) {
  return String(napId || '').replace(/(\d)[A-Za-z]$/, '$1');
}

// ============================ ENRICHMENT (matches Python) ============================
let _sortedPrefixes = null;
function sortedPrefixes_() {
  if (_sortedPrefixes) return _sortedPrefixes;
  _sortedPrefixes = Object.keys(PREFIX_TERRITORY).slice()
    .sort(function (a, b) { return b.length - a.length; });
  return _sortedPrefixes;
}

function getNapPrefix_(napId) {
  if (!napId) return '';
  const u = String(napId).trim().toUpperCase();
  const ps = sortedPrefixes_();
  for (let i = 0; i < ps.length; i++) {
    if (u.indexOf(ps[i]) === 0) return ps[i];
  }
  return '';
}

function getSalesArea_(napId) { return NAP_AREA_LOOKUP[getNapPrefix_(napId)] || ''; }
function getProvince_(napId)  { return NAP_PROVINCE_LOOKUP[getNapPrefix_(napId)] || ''; }

// Stricter prefix match: requires next char to be _, -, digit, or L.
function getTerritory_(napId) {
  if (!napId) return '';
  const u = String(napId).trim().toUpperCase();
  const ps = sortedPrefixes_();
  for (let i = 0; i < ps.length; i++) {
    const p = ps[i].toUpperCase();
    if (u === p) return PREFIX_TERRITORY[ps[i]];
    if (u.length > p.length) {
      const next = u.charAt(p.length);
      const isDelim = (next === '_' || next === '-' || next === 'L' || (next >= '0' && next <= '9'));
      if (isDelim && u.indexOf(p) === 0) return PREFIX_TERRITORY[ps[i]];
    }
  }
  return '';
}

function getTech_(cabinet) {
  if (!cabinet) return '';
  const u = String(cabinet).toUpperCase();
  if (u.indexOf('LSA') !== -1) return 'GPON';
  if (cabinet.indexOf('-M') !== -1) return CABINET_TECH_LOOKUP[cabinet] || 'ADSL/VDSL';
  return 'GPON';
}

function getPlaId_(cabinet) {
  if (!cabinet) return '';
  const parts = String(cabinet).trim().split('_');
  if (parts.length >= 2) {
    const k2 = parts[0] + '_' + parts[1];
    if (PLA_ID_LOOKUP[k2] != null) return PLA_ID_LOOKUP[k2];
  }
  if (PLA_ID_LOOKUP[parts[0]] != null) return PLA_ID_LOOKUP[parts[0]];
  return '';
}

function calcUtilization_(pa, pt) {
  const a = (typeof pa === 'number') ? pa : parseInt(pa, 10);
  const t = (typeof pt === 'number') ? pt : parseInt(pt, 10);
  if (isNaN(t)) return '';
  if (t === 0) return 0;
  if (isNaN(a)) return '';
  return Math.round((a / t) * 10000) / 10000;
}

// ============================ GEO LOOKUP ============================
function loadGeoLookup_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('geo_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const out = {};
  if (sh) {
    const last = sh.getLastRow();
    if (last >= 2) {
      const v = sh.getRange(2, 1, last - 1, 4).getValues();
      for (let i = 0; i < v.length; i++) {
        const nap = s_(v[i][0]);
        if (!nap) continue;
        out[nap] = { city: s_(v[i][1]), brgy: s_(v[i][2]), loc: s_(v[i][3]) };
      }
    }
  }
  try { cache.put('geo_v1', JSON.stringify(out), CONFIG.GEO_CACHE_SECONDS); } catch (e) {}
  return out;
}

function invalidateGeoCache_() {
  try { CacheService.getScriptCache().remove('geo_v1'); } catch (e) {}
}

// ============================ CONVERTER ============================
function processNapCsv(csvText, snapshotDateStr) {
  const session = getSession();
  if (!session.email) throw new Error('Sign in with your Google account first.');

  const geo = loadGeoLookup_();
  const lines = csvText.split(/\r\n|\n|\r/);

  // Pass 1: parse + filter
  const recs = [];
  let readCount = 0, skipped = 0;
  for (let i = 0; i < lines.length; i++) {
    if (i === 0) continue;                              // header row
    const raw = (lines[i] || '');
    if (!raw.replace(/;/g, '').trim()) continue;        // blank
    if (isJunkRow_(raw)) continue;                      // header garbage
    readCount++;

    const rec = parseRaw_(raw);
    if (rec === null) { skipped++; continue; }
    if (!getTerritory_(rec.napId)) { skipped++; continue; }
    recs.push(rec);
  }

  // Pass 2: merge by base NAP ID (strip_suffix)
  const merged = {};
  const order = [];
  for (let i = 0; i < recs.length; i++) {
    const rec = recs[i];
    const base = stripSuffix_(rec.napId);
    if (!merged[base]) {
      order.push(base);
      merged[base] = {
        cabinet:   rec.cabinet,
        napId:     base,
        discovered: rec.discovered,
        lat:       rec.lat,
        lon:       rec.lon,
        pa:        toInt_(rec.portsAssigned),
        pr:        toInt_(rec.portsReserved),
        pt:        toInt_(rec.portsTotal),
        territory: getTerritory_(rec.napId),
        firstPt:   toInt_(rec.portsTotal)
      };
    } else {
      const e = merged[base];
      const newPa = toInt_(rec.portsAssigned), newPr = toInt_(rec.portsReserved), newPt = toInt_(rec.portsTotal);
      const pa = isInt_(e.pa) ? e.pa : 0;
      const pr = isInt_(e.pr) ? e.pr : 0;
      const pt = isInt_(e.pt) ? e.pt : 0;
      const firstPt = isInt_(e.firstPt) ? e.firstPt : 0;
      e.pa = pa + (isInt_(newPa) ? newPa : 0);
      e.pr = pr + (isInt_(newPr) ? newPr : 0);
      if (isInt_(newPt) && newPt === 16 && firstPt === 16) e.pt = 16;
      else e.pt = pt + (isInt_(newPt) ? newPt : 0);
    }
  }

  // Pass 3: build output rows
  const rows = [];
  const missing = [];
  for (let i = 0; i < order.length; i++) {
    const m = merged[order[i]];
    const pa = m.pa, pt = m.pt;
    const util = (isInt_(pa) && isInt_(pt)) ? calcUtilization_(pa, pt) : (pt === 0 ? 0 : '');
    const g = geo[m.napId] || { city: '', brgy: '', loc: '' };

    rows.push([
      m.cabinet,
      m.napId,
      m.discovered,
      getPlaId_(m.cabinet),
      getTech_(m.cabinet),
      pa, m.pr, pt, util,
      toCoord_(m.lat), toCoord_(m.lon),
      getSalesArea_(m.napId),
      m.territory,
      g.brgy, g.city,
      getProvince_(m.napId),
      g.loc
    ]);

    if (!g.city || !g.brgy || !g.loc) {
      missing.push({ napId: m.napId, city: g.city, brgy: g.brgy, location: g.loc });
    }
  }

  const snapshotDate = normalizeDate_(snapshotDateStr);
  saveSnapshot_(snapshotDate, session.email, rows);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const downloadUrl = buildXlsx_('NAP_cleaned_' + snapshotDate + '_' + stamp, rows);

  return {
    downloadUrl: downloadUrl,
    written: rows.length,
    read: readCount,
    skipped: skipped,
    snapshotDate: snapshotDate,
    preview: rows.slice(0, 50).map(function (r) { return rowToObject_(r); }),
    missing: missing
  };
}

function rowToObject_(r) {
  const o = {};
  for (let i = 0; i < OUTPUT_COLS.length; i++) o[OUTPUT_COLS[i]] = r[i];
  return o;
}

function normalizeDate_(d) {
  if (!d) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const t = new Date(d);
  if (isNaN(t.getTime())) return String(d).slice(0, 10);
  return Utilities.formatDate(t, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ============================ SNAPSHOT STORAGE ============================
function saveSnapshot_(snapshotDate, email, rows) {
  const ss = openLookupSheet_();
  const sh = ss.getSheetByName(CONFIG.TABS.NAP_DATA);
  if (!sh) throw new Error('NAP_DATA tab missing — run setupSheets().');

  const last = sh.getLastRow();
  if (last > 1) {
    const dateCol = sh.getRange(2, 1, last - 1, 1).getValues();
    const keep = [];
    for (let i = 0; i < dateCol.length; i++) {
      if (String(dateCol[i][0]) !== snapshotDate) keep.push(i);
    }
    if (keep.length !== dateCol.length) {
      const all = sh.getRange(2, 1, last - 1, SNAPSHOT_COLS.length).getValues();
      const kept = keep.map(function (i) { return all[i]; });
      sh.getRange(2, 1, last - 1, SNAPSHOT_COLS.length).clearContent();
      if (kept.length) sh.getRange(2, 1, kept.length, SNAPSHOT_COLS.length).setValues(kept);
    }
  }

  if (rows.length) {
    const data = rows.map(function (r) { return [snapshotDate, email].concat(r); });
    const writeRow = sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, data.length, SNAPSHOT_COLS.length).setValues(data);
  }

  upsertSnapshotIndex_(snapshotDate, email, rows.length);
}

function upsertSnapshotIndex_(snapshotDate, email, count) {
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.SNAPSHOT_INDEX);
  if (!sh) return;
  const last = sh.getLastRow();
  const now = new Date();
  if (last > 1) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      if (String(dates[i][0]) === snapshotDate) {
        sh.getRange(i + 2, 1, 1, SNAPSHOT_INDEX_COLS.length)
          .setValues([[snapshotDate, email, now, count]]);
        return;
      }
    }
  }
  sh.appendRow([snapshotDate, email, now, count]);
}

function listSnapshots() {
  getSession();
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.SNAPSHOT_INDEX);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, SNAPSHOT_INDEX_COLS.length).getValues();
  return v.map(function (r) {
    return {
      snapshotDate: normalizeDate_(r[0]),
      uploadedBy:   s_(r[1]),
      uploadedAt:   r[2] ? new Date(r[2]).toISOString() : '',
      rowCount:     Number(r[3]) || 0
    };
  }).sort(function (a, b) { return a.snapshotDate < b.snapshotDate ? 1 : -1; });
}

function exportSnapshot(snapshotDate) {
  const session = getSession();
  if (!session.email) throw new Error('Sign in first.');
  snapshotDate = normalizeDate_(snapshotDate);
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.NAP_DATA);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('No data for ' + snapshotDate);
  const all = sh.getRange(2, 1, last - 1, SNAPSHOT_COLS.length).getValues();
  const rows = all.filter(function (r) { return String(r[0]) === snapshotDate; })
                  .map(function (r) { return r.slice(2); });
  if (!rows.length) throw new Error('No data for ' + snapshotDate);
  return { downloadUrl: buildXlsx_('NAP_cleaned_' + snapshotDate, rows), rowCount: rows.length };
}

// ============================ XLSX BUILD ============================
function buildXlsx_(baseName, dataRows) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const tmp = SpreadsheetApp.create(baseName + '_' + stamp);
  const sh = tmp.getSheets()[0];
  sh.getRange(1, 1, 1, OUTPUT_COLS.length).setValues([OUTPUT_COLS])
    .setFontWeight('bold').setBackground('#e8eefc');
  if (dataRows.length) {
    sh.getRange(2, 1, dataRows.length, OUTPUT_COLS.length).setValues(dataRows);
  }
  sh.setFrozenRows(1);
  SpreadsheetApp.flush();

  const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName(baseName + '.xlsx');
  const file = getOutputFolder_().createFile(blob);
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return file.getUrl();
}

function buildMissingXlsx(missing) {
  getSession();
  const headers = ['NAP ID', 'CITY_NAME', 'BRGY_NAME', 'LOCATION TAGGING'];
  const rows = (missing || []).map(function (m) {
    return [m.napId, m.city || '', m.brgy || '', m.location || ''];
  });
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const tmp = SpreadsheetApp.create('NAP_missing_' + stamp);
  const sh = tmp.getSheets()[0];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#fff2cc');
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName('NAP_missing_' + stamp + '.xlsx');
  const file = getOutputFolder_().createFile(blob);
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return { downloadUrl: file.getUrl(), rowCount: rows.length };
}

function getOutputFolder_() {
  const it = DriveApp.getFoldersByName(CONFIG.OUTPUT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.OUTPUT_FOLDER_NAME);
}

// ============================ GEO REFERENCE (admin) ============================
function listGeo(filter) {
  requireAdmin_();
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, GEO_HEADERS.length).getValues();
  const q = String(filter || '').trim().toLowerCase();
  const out = [];
  for (let i = 0; i < v.length; i++) {
    const r = v[i];
    if (!s_(r[0])) continue;
    const row = { napId: s_(r[0]), city: s_(r[1]), brgy: s_(r[2]), location: s_(r[3]),
                  updatedAt: r[4] ? new Date(r[4]).toISOString() : '', updatedBy: s_(r[5]) };
    if (!q || [row.napId, row.city, row.brgy, row.location].join(' ').toLowerCase().indexOf(q) !== -1) {
      out.push(row);
    }
  }
  return out;
}

function getGeo(napId) {
  requireAdmin_();
  napId = s_(napId);
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const v = sh.getRange(2, 1, last - 1, GEO_HEADERS.length).getValues();
  for (let i = 0; i < v.length; i++) {
    if (s_(v[i][0]) === napId) {
      return { napId: napId, city: s_(v[i][1]), brgy: s_(v[i][2]), location: s_(v[i][3]) };
    }
  }
  return null;
}

function upsertGeo(entry) {
  const session = requireAdmin_();
  const napId = s_(entry && entry.napId);
  if (!napId) throw new Error('NAP ID is required.');
  const city = s_(entry.city), brgy = s_(entry.brgy), location = s_(entry.location);

  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const last = sh.getLastRow();
  const now = new Date();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < v.length; i++) {
      if (s_(v[i][0]) === napId) {
        sh.getRange(i + 2, 1, 1, GEO_HEADERS.length)
          .setValues([[napId, city, brgy, location, now, session.email]]);
        invalidateGeoCache_();
        return { napId: napId, action: 'updated' };
      }
    }
  }
  sh.appendRow([napId, city, brgy, location, now, session.email]);
  invalidateGeoCache_();
  return { napId: napId, action: 'added' };
}

function deleteGeo(napId) {
  requireAdmin_();
  napId = s_(napId);
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const last = sh.getLastRow();
  if (last < 2) return { deleted: false };
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (s_(v[i][0]) === napId) {
      sh.deleteRow(i + 2);
      invalidateGeoCache_();
      return { deleted: true };
    }
  }
  return { deleted: false };
}

function bulkUpsertGeo(xlsxBase64, mimeType) {
  const session = requireAdmin_();
  if (!xlsxBase64) throw new Error('No file uploaded.');
  const blob = Utilities.newBlob(Utilities.base64Decode(xlsxBase64),
                                 mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 'bulk_geo.xlsx');
  const tempFile = Drive.Files.insert({ title: 'NAP_bulk_geo_tmp', mimeType: MimeType.GOOGLE_SHEETS }, blob);
  const ss = SpreadsheetApp.openById(tempFile.id);
  const sh = ss.getSheets()[0];
  const last = sh.getLastRow();
  if (last < 2) { DriveApp.getFileById(tempFile.id).setTrashed(true); return { processed: 0, added: 0, updated: 0 }; }
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return s_(h).toUpperCase(); });
  const idx = {
    nap:  header.indexOf('NAP ID'),
    city: header.indexOf('CITY_NAME'),
    brgy: header.indexOf('BRGY_NAME'),
    loc:  header.indexOf('LOCATION TAGGING')
  };
  for (const k in idx) if (idx[k] < 0) {
    DriveApp.getFileById(tempFile.id).setTrashed(true);
    throw new Error('Missing column: ' + k.toUpperCase());
  }
  const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  DriveApp.getFileById(tempFile.id).setTrashed(true);

  const target = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const targetLast = target.getLastRow();
  const existing = {};
  if (targetLast > 1) {
    const cur = target.getRange(2, 1, targetLast - 1, GEO_HEADERS.length).getValues();
    for (let i = 0; i < cur.length; i++) existing[s_(cur[i][0])] = i + 2;
  }
  const now = new Date();
  let added = 0, updated = 0;
  const appendBuf = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const nap = s_(r[idx.nap]);
    if (!nap) continue;
    const row = [nap, s_(r[idx.city]), s_(r[idx.brgy]), s_(r[idx.loc]), now, session.email];
    if (existing[nap]) {
      target.getRange(existing[nap], 1, 1, GEO_HEADERS.length).setValues([row]);
      updated++;
    } else {
      appendBuf.push(row);
      added++;
    }
  }
  if (appendBuf.length) {
    target.getRange(target.getLastRow() + 1, 1, appendBuf.length, GEO_HEADERS.length).setValues(appendBuf);
  }
  invalidateGeoCache_();
  return { processed: added + updated, added: added, updated: updated };
}

// ============================ ADMIN MANAGEMENT ============================
function listAdmins() {
  requireAdmin_();
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.ADMINS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, ADMIN_HEADERS.length).getValues();
  return v.filter(function (r) { return s_(r[0]); }).map(function (r) {
    return { email: s_(r[0]), addedAt: r[1] ? new Date(r[1]).toISOString() : '', addedBy: s_(r[2]) };
  });
}

function addAdmin(email) {
  const session = requireAdmin_();
  email = s_(email).toLowerCase();
  if (!email || email.indexOf('@') < 0) throw new Error('Valid email required.');
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.ADMINS);
  const last = sh.getLastRow();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < v.length; i++) {
      if (s_(v[i][0]).toLowerCase() === email) return { email: email, action: 'exists' };
    }
  }
  sh.appendRow([email, new Date(), session.email]);
  return { email: email, action: 'added' };
}

function removeAdmin(email) {
  const session = requireAdmin_();
  email = s_(email).toLowerCase();
  if (email === session.email.toLowerCase()) throw new Error('Refusing to remove yourself.');
  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.ADMINS);
  const last = sh.getLastRow();
  if (last < 2) return { removed: false };
  const v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) {
    if (s_(v[i][0]).toLowerCase() === email) {
      sh.deleteRow(i + 2);
      return { removed: true };
    }
  }
  return { removed: false };
}

// ============================ SETUP ============================
function setupSheets() {
  const ss = openLookupSheet_();
  const need = [
    [CONFIG.TABS.ADMINS,         ADMIN_HEADERS],
    [CONFIG.TABS.GEO_REFERENCE,  GEO_HEADERS],
    [CONFIG.TABS.NAP_DATA,       SNAPSHOT_COLS],
    [CONFIG.TABS.SNAPSHOT_INDEX, SNAPSHOT_INDEX_COLS]
  ];
  need.forEach(function (pair) {
    const name = pair[0], headers = pair[1];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#e8eefc');
      sh.setFrozenRows(1);
    }
  });

  if (CONFIG.SEED_ADMIN_EMAIL) {
    const adminSh = ss.getSheetByName(CONFIG.TABS.ADMINS);
    if (adminSh.getLastRow() < 2) {
      adminSh.appendRow([CONFIG.SEED_ADMIN_EMAIL.toLowerCase(), new Date(), 'setup']);
    }
  }
  return 'Setup complete. Tabs: ' + need.map(function (p) { return p[0]; }).join(', ');
}
