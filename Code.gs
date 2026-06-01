/**
 * NAP Data Converter — Google Apps Script port of the Streamlit app
 * (AlexisKx/NAP_CONVERTER_APP → Google Sites + Apps Script).
 *
 * All five Streamlit pages are reproduced as a single web app:
 *   - 📡 Converter
 *   - 🕓 Data History
 *   - 📁 GEO Reference   (admin)
 *   - 👥 Admin Management (admin)   ← replaces "User Management"; SSO handles identity
 *   - ℹ️  About / Handover
 *
 * Auth: Workspace SSO via Session.getActiveUser(). An ADMINS tab in the lookup
 * sheet controls who sees admin-only pages.
 *
 * Storage: one Google Sheet ("Lookup Spreadsheet") holds every tab.
 *   - ADMINS, MASTER, PLA_BY_CABINET, AREA_BY_PREFIX, GEO_REFERENCE,
 *     NAP_DATA (snapshots), SNAPSHOT_INDEX (one row per saved snapshot).
 *
 * Run setupSheets() once from the script editor to create the tabs.
 */

// ============================ CONFIG ============================
const CONFIG = {
  LOOKUP_SPREADSHEET_ID: 'PASTE_YOUR_LOOKUP_SPREADSHEET_ID_HERE',

  TABS: {
    ADMINS:          'ADMINS',
    MASTER:          'MASTER',
    PLA_BY_CABINET:  'PLA_BY_CABINET',
    AREA_BY_PREFIX:  'AREA_BY_PREFIX',
    GEO_REFERENCE:   'GEO_REFERENCE',
    NAP_DATA:        'NAP_DATA',
    SNAPSHOT_INDEX:  'SNAPSHOT_INDEX'
  },

  TRAILING_COLS: 12,
  OUTPUT_FOLDER_NAME: 'NAP Converter Output',
  GEO_CACHE_SECONDS: 300,
  SEED_ADMIN_EMAIL: ''   // optional: email auto-promoted to admin on first setupSheets()
};

const OUTPUT_COLS = ['Cabinet', 'NAP ID', 'Discovered When', 'PLA ID', 'Tech',
  'Ports Assigned', 'Ports Reserved', 'Ports Total', 'UTILIZATION',
  'Latitude', 'Longitude', 'SALES_AREA', 'TERRITORY', 'BRGY_NAME',
  'CITY_NAME', 'PROVINCE_NAME', 'LOCATION TAGGING'];

const SNAPSHOT_COLS = ['snapshot_date', 'uploaded_by'].concat(OUTPUT_COLS);
const SNAPSHOT_INDEX_COLS = ['snapshot_date', 'uploaded_by', 'uploaded_at', 'row_count'];
const GEO_HEADERS = ['NAP ID', 'CITY_NAME', 'BRGY_NAME', 'LOCATION TAGGING', 'updated_at', 'updated_by'];
const ADMIN_HEADERS = ['email', 'added_at', 'added_by'];

const COORD_RE = /^-?\d{1,3}\.\d{4,}$/;

// ============================ WEB APP ENTRY ============================
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('NAP Data Converter')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);  // needed for Google Sites embed
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
  const ss = openLookupSheet_();
  const sh = ss.getSheetByName(CONFIG.TABS.ADMINS);
  if (!sh) return false;
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
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

// ============================ PARSING (matches Python) ============================
function parseRow_(raw) {
  const fields = raw.split(';');
  const n = fields.length;
  if (n < CONFIG.TRAILING_COLS + 6) return null;
  const tail = fields.slice(n - CONFIG.TRAILING_COLS);
  return {
    cabinet:       (fields[0] || '').trim(),
    napId:         (fields[1] || '').trim(),
    discovered:    (tail[2]  || '').trim(),
    portsAssigned: (tail[7]  || '').trim(),
    portsReserved: (tail[8]  || '').trim(),
    portsTotal:    (tail[6]  || '').trim(),
    lat:           (tail[0]  || '').trim(),
    lon:           (tail[1]  || '').trim()
  };
}

function toCoord_(v) {
  v = (v || '').trim();
  return COORD_RE.test(v) ? Number(v) : '';
}

function calcUtilization_(pa, pt) {
  const a = parseInt(pa, 10), t = parseInt(pt, 10);
  if (!t || isNaN(t) || isNaN(a)) return '';
  return Math.round((a / t) * 10000) / 10000;
}

// Leading alphabetic portion — used for AREA_BY_PREFIX lookup.
function napPrefix_(napId) {
  const m = (napId || '').match(/^[A-Za-z]+/);
  return m ? m[0].toUpperCase() : '';
}

// "Base NAP ID" used as the merge key. The Streamlit version says duplicates
// are merged by base NAP ID with a suffix stripped; the exact suffix rule
// isn't in the handover, so we strip a trailing "-NNN" / "_NNN" / final
// "-X" if present. >>> RECONCILE against app.py.merge_duplicates() <<<
function baseNapId_(napId) {
  if (!napId) return '';
  return String(napId).replace(/[-_][A-Za-z0-9]+$/, '').trim() || String(napId).trim();
}

// ============================ LOOKUPS ============================
function openLookupSheet_() {
  if (!CONFIG.LOOKUP_SPREADSHEET_ID || CONFIG.LOOKUP_SPREADSHEET_ID.indexOf('PASTE_') === 0) {
    throw new Error('CONFIG.LOOKUP_SPREADSHEET_ID is not set. Edit Code.gs.');
  }
  return SpreadsheetApp.openById(CONFIG.LOOKUP_SPREADSHEET_ID);
}

function readTab_(name) {
  const sh = openLookupSheet_().getSheetByName(name);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  return v.length > 1 ? v.slice(1) : [];
}

function s_(x) { return String(x == null ? '' : x).trim(); }

function loadLookups_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('lookups_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }

  const master = {};
  readTab_(CONFIG.TABS.MASTER).forEach(function (r) {
    const nap = s_(r[0]);
    if (!nap) return;
    master[nap] = {
      pla: s_(r[1]), tech: s_(r[2]), territory: s_(r[3]), area: s_(r[4]),
      province: s_(r[5]), city: s_(r[6]), brgy: s_(r[7]), location: s_(r[8])
    };
  });

  const plaByCabinet = {};
  readTab_(CONFIG.TABS.PLA_BY_CABINET).forEach(function (r) {
    const cab = s_(r[0]);
    if (!cab) return;
    plaByCabinet[cab] = { pla: s_(r[1]), tech: s_(r[2]) };
  });

  const areaByPrefix = {};
  readTab_(CONFIG.TABS.AREA_BY_PREFIX).forEach(function (r) {
    const pfx = s_(r[0]).toUpperCase();
    if (!pfx) return;
    areaByPrefix[pfx] = { area: s_(r[1]), province: s_(r[2]), territory: s_(r[3]) };
  });

  const geo = {};
  readTab_(CONFIG.TABS.GEO_REFERENCE).forEach(function (r) {
    const nap = s_(r[0]);
    if (!nap) return;
    geo[nap] = { city: s_(r[1]), brgy: s_(r[2]), location: s_(r[3]) };
  });

  const out = { master: master, plaByCabinet: plaByCabinet, areaByPrefix: areaByPrefix, geo: geo };
  try { cache.put('lookups_v1', JSON.stringify(out), CONFIG.GEO_CACHE_SECONDS); } catch (e) {}
  return out;
}

function invalidateLookupCache_() {
  try { CacheService.getScriptCache().remove('lookups_v1'); } catch (e) {}
}

// ============================ CONVERTER ============================
function processNapCsv(csvText, snapshotDateStr) {
  const session = getSession();
  if (!session.email) throw new Error('Sign in with your Google account first.');

  const lookups = loadLookups_();
  const lines = csvText.split(/\r\n|\n|\r/);

  const merged = {};
  let readCount = 0, skipped = 0;

  for (let i = 0; i < lines.length; i++) {
    if (i === 0) continue;
    const raw = lines[i];
    if (!raw || !raw.replace(/;/g, '').trim()) continue;
    const rec = parseRow_(raw);
    if (!rec) { skipped++; continue; }
    readCount++;

    const key = baseNapId_(rec.napId) || (rec.cabinet + '|' + i);
    if (!merged[key]) merged[key] = {
      cabinet: rec.cabinet, napId: baseNapId_(rec.napId) || rec.napId, discovered: rec.discovered,
      lat: rec.lat, lon: rec.lon, _pa: 0, _pr: 0, _pt: 0
    };
    const m = merged[key];
    m._pa += parseInt(rec.portsAssigned, 10) || 0;
    m._pr += parseInt(rec.portsReserved, 10) || 0;
    m._pt += parseInt(rec.portsTotal, 10)    || 0;
  }

  const rows = [];
  const missing = [];
  Object.keys(merged).forEach(function (key) {
    const m   = merged[key];
    const nap = m.napId;
    const ref = lookups.master[nap]            || {};
    const cab = lookups.plaByCabinet[m.cabinet] || {};
    const pfx = lookups.areaByPrefix[napPrefix_(nap)] || {};
    const g   = lookups.geo[nap]               || {};

    const pla       = ref.pla       || cab.pla       || '';
    const tech      = ref.tech      || cab.tech      || '';
    const territory = ref.territory || pfx.territory || '';
    const area      = ref.area      || pfx.area      || '';
    const province  = ref.province  || pfx.province  || '';
    const city      = ref.city      || g.city        || '';
    const brgy      = ref.brgy      || g.brgy        || '';
    const location  = ref.location  || g.location    || '';

    rows.push([
      m.cabinet, nap, m.discovered, pla, tech,
      m._pa, m._pr, m._pt, calcUtilization_(m._pa, m._pt),
      toCoord_(m.lat), toCoord_(m.lon),
      area, territory, brgy, city, province, location
    ]);

    if (!city || !brgy || !location) {
      missing.push({
        napId: nap, cabinet: m.cabinet,
        city: city, brgy: brgy, location: location
      });
    }
  });

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

  // Upsert: remove existing rows for this date, then append new ones.
  const last = sh.getLastRow();
  if (last > 1) {
    const dateCol = sh.getRange(2, 1, last - 1, 1).getValues();
    const keepRows = [];
    for (let i = 0; i < dateCol.length; i++) {
      if (String(dateCol[i][0]) !== snapshotDate) keepRows.push(i);
    }
    if (keepRows.length !== dateCol.length) {
      const all = sh.getRange(2, 1, last - 1, SNAPSHOT_COLS.length).getValues();
      const kept = keepRows.map(function (i) { return all[i]; });
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
      uploadedBy: s_(r[1]),
      uploadedAt: r[2] ? new Date(r[2]).toISOString() : '',
      rowCount: Number(r[3]) || 0
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
        invalidateLookupCache_();
        return { napId: napId, action: 'updated' };
      }
    }
  }
  sh.appendRow([napId, city, brgy, location, now, session.email]);
  invalidateLookupCache_();
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
      invalidateLookupCache_();
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
  invalidateLookupCache_();
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
    [CONFIG.TABS.MASTER,         ['NAP ID', 'PLA ID', 'Tech', 'Territory', 'Area', 'Province', 'City', 'BRGY', 'Location']],
    [CONFIG.TABS.PLA_BY_CABINET, ['Cabinet', 'PLA ID', 'Tech']],
    [CONFIG.TABS.AREA_BY_PREFIX, ['Prefix', 'Sales Area', 'Province', 'Territory']],
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
  return 'Setup complete.';
}
