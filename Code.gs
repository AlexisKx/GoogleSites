/**
 * NAP Data Converter — slim server.
 *
 * In this version the converter runs ENTIRELY in the browser:
 *   - CSV is streamed and parsed client-side
 *   - merge + enrichment + XLSX build are all client-side
 *   - file is downloaded directly from the browser, never touches the server
 *
 * The server only:
 *   - serves the static HTML/JS shell
 *   - authenticates the viewer (Workspace SSO via Session.getActiveUser)
 *   - reads / writes the GEO_REFERENCE tab (the only thing we persist)
 *   - manages the ADMINS allowlist
 *   - returns the full GEO dict to the browser once per session
 *
 * Storage: one Google Sheet with two tabs — ADMINS and GEO_REFERENCE.
 *
 * Run setupSheets() once from the script editor after pasting the
 * Sheet ID into CONFIG.LOOKUP_SPREADSHEET_ID.
 */

// ============================ CONFIG ============================
const CONFIG = {
  LOOKUP_SPREADSHEET_ID: 'PASTE_YOUR_LOOKUP_SPREADSHEET_ID_HERE',

  TABS: {
    ADMINS:        'ADMINS',
    GEO_REFERENCE: 'GEO_REFERENCE'
  },

  GEO_CACHE_SECONDS: 300,
  SEED_ADMIN_EMAIL:  ''   // optional: email auto-promoted to admin on first setupSheets()
};

const GEO_HEADERS = ['NAP ID', 'CITY_NAME', 'BRGY_NAME', 'LOCATION TAGGING', 'updated_at', 'updated_by'];
const ADMIN_HEADERS = ['email', 'added_at', 'added_by'];

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

function openLookupSheet_() {
  if (!CONFIG.LOOKUP_SPREADSHEET_ID || CONFIG.LOOKUP_SPREADSHEET_ID.indexOf('PASTE_') === 0) {
    throw new Error('CONFIG.LOOKUP_SPREADSHEET_ID is not set. Edit Code.gs.');
  }
  return SpreadsheetApp.openById(CONFIG.LOOKUP_SPREADSHEET_ID);
}

// ============================ GEO — BULK LOAD FOR CLIENT ============================
// Returns the entire GEO_REFERENCE as { napId: [city, brgy, loc] }.
// Cached server-side for 5 minutes. The browser caches it for the session.
function getGeoData() {
  getSession();
  const cache = CacheService.getScriptCache();
  const cached = cache.get('geo_v2');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  const sh = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const out = {};
  if (sh) {
    const last = sh.getLastRow();
    if (last >= 2) {
      const v = sh.getRange(2, 1, last - 1, 4).getValues();
      for (let i = 0; i < v.length; i++) {
        const nap = s_(v[i][0]);
        if (!nap) continue;
        out[nap] = [s_(v[i][1]), s_(v[i][2]), s_(v[i][3])];
      }
    }
  }
  try { cache.put('geo_v2', JSON.stringify(out), CONFIG.GEO_CACHE_SECONDS); } catch (e) {}
  return out;
}

function invalidateGeoCache_() {
  try { CacheService.getScriptCache().remove('geo_v2'); } catch (e) {}
}

// ============================ GEO REFERENCE (admin UI) ============================
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

// Bulk upsert from the browser. The browser parses the Excel (with SheetJS),
// extracts {napId, city, brgy, location} rows, and sends them as JSON —
// no need for the Drive Advanced Service anymore.
function bulkUpsertGeoRows(rows) {
  const session = requireAdmin_();
  if (!Array.isArray(rows)) throw new Error('Bad payload.');

  const target = openLookupSheet_().getSheetByName(CONFIG.TABS.GEO_REFERENCE);
  const targetLast = target.getLastRow();
  const existing = {};
  if (targetLast > 1) {
    const cur = target.getRange(2, 1, targetLast - 1, 1).getValues();
    for (let i = 0; i < cur.length; i++) existing[s_(cur[i][0])] = i + 2;
  }
  const now = new Date();
  let added = 0, updated = 0;
  const appendBuf = [];

  // Batch updates: collect by row then write in chunks.
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const nap = s_(r.napId);
    if (!nap) continue;
    const out = [nap, s_(r.city), s_(r.brgy), s_(r.location), now, session.email];
    if (existing[nap]) {
      target.getRange(existing[nap], 1, 1, GEO_HEADERS.length).setValues([out]);
      updated++;
    } else {
      appendBuf.push(out);
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
    [CONFIG.TABS.ADMINS,        ADMIN_HEADERS],
    [CONFIG.TABS.GEO_REFERENCE, GEO_HEADERS]
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
