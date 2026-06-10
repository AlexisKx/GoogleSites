/**
 * NAP Data Converter — slim server with username/password auth.
 *
 * Auth model: classic username + password, sessions tracked in a Sheet.
 *   - No Google sign-in required from end users.
 *   - Web app deployed as "Execute as: ME" (you, the deployer), so users
 *     don't see Google's "Unverified" warning.
 *   - Passwords are stored hashed (iterated SHA-256 + per-user salt).
 *   - Sessions are random 256-bit tokens stored in USER_SESSIONS,
 *     with a TTL (default 24h).
 *
 * Storage tabs (auto-created by setupSheets):
 *   - USERS:         username, password_hash, salt, is_admin, created_at,
 *                    created_by, last_login
 *   - USER_SESSIONS: token, username, created_at, expires_at
 *   - GEO_REFERENCE: NAP ID, CITY_NAME, BRGY_NAME, LOCATION TAGGING, updated_at, updated_by
 *
 * Bootstrap: if USERS is empty, setupSheets() seeds an admin account
 * from CONFIG.SEED_ADMIN. Log in once with those credentials, then
 * IMMEDIATELY change the password from the Change Password page.
 */

// ============================ CONFIG ============================
const CONFIG = {
  LOOKUP_SPREADSHEET_ID: 'PASTE_YOUR_LOOKUP_SPREADSHEET_ID_HERE',

  TABS: {
    USERS:         'USERS',
    USER_SESSIONS: 'USER_SESSIONS',
    GEO_REFERENCE: 'GEO_REFERENCE'
  },

  SESSION_TTL_HOURS: 24,
  HASH_ITERATIONS:   1000,
  GEO_CACHE_SECONDS: 300,

  // Seeded once if USERS is empty. CHANGE THE PASSWORD on first login.
  SEED_ADMIN: {
    username: 'admin',
    password: 'changeme123'
  }
};

const USERS_HEADERS    = ['username', 'password_hash', 'salt', 'is_admin', 'created_at', 'created_by', 'last_login'];
const SESSIONS_HEADERS = ['token', 'username', 'created_at', 'expires_at'];
const GEO_HEADERS      = ['NAP ID', 'CITY_NAME', 'BRGY_NAME', 'LOCATION TAGGING', 'updated_at', 'updated_by'];

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

// ============================ HELPERS ============================
function s_(x)        { return String(x == null ? '' : x).trim(); }
function lc_(x)       { return s_(x).toLowerCase(); }

function openLookupSheet_() {
  if (!CONFIG.LOOKUP_SPREADSHEET_ID || CONFIG.LOOKUP_SPREADSHEET_ID.indexOf('PASTE_') === 0) {
    throw new Error('CONFIG.LOOKUP_SPREADSHEET_ID is not set. Edit Code.gs.');
  }
  return SpreadsheetApp.openById(CONFIG.LOOKUP_SPREADSHEET_ID);
}

// ============================ CRYPTO ============================
function bytesToHex_(bytes) {
  let h = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i] & 0xff;
    h += (b < 16 ? '0' : '') + b.toString(16);
  }
  return h;
}

function generateSalt_() {
  return Utilities.getUuid().replace(/-/g, '');
}

function generateToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function hashPassword_(password, salt) {
  let cur = String(password || '') + String(salt || '');
  for (let i = 0; i < CONFIG.HASH_ITERATIONS; i++) {
    cur = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cur));
  }
  return cur;
}

function constantTimeEqual_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

// ============================ USER STORAGE ============================
function usersSheet_()    { return openLookupSheet_().getSheetByName(CONFIG.TABS.USERS); }
function sessionsSheet_() { return openLookupSheet_().getSheetByName(CONFIG.TABS.USER_SESSIONS); }

function findUserByUsername_(username) {
  const u = lc_(username);
  if (!u) return null;
  const sh = usersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, USERS_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (lc_(data[i][0]) === u) {
      return {
        row: i + 2,
        record: {
          username:      s_(data[i][0]),
          password_hash: s_(data[i][1]),
          salt:          s_(data[i][2]),
          is_admin:      data[i][3] === true || lc_(data[i][3]) === 'true',
          created_at:    data[i][4],
          created_by:    s_(data[i][5]),
          last_login:    data[i][6]
        }
      };
    }
  }
  return null;
}

function updateUserFields_(username, fields) {
  const found = findUserByUsername_(username);
  if (!found) throw new Error('User not found.');
  const sh = usersSheet_();
  for (const key in fields) {
    const idx = USERS_HEADERS.indexOf(key);
    if (idx < 0) continue;
    sh.getRange(found.row, idx + 1).setValue(fields[key]);
  }
}

// ============================ SESSIONS ============================
function createSession_(username) {
  const sh = sessionsSheet_();
  const token = generateToken_();
  const now = new Date();
  const exp = new Date(now.getTime() + CONFIG.SESSION_TTL_HOURS * 3600 * 1000);
  sh.appendRow([token, username, now, exp]);
  return { token: token, expiresAt: exp };
}

function getSessionUser_(token) {
  token = s_(token);
  if (!token) return null;
  const sh = sessionsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, SESSIONS_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (s_(data[i][0]) === token) {
      const exp = new Date(data[i][3]);
      if (isNaN(exp.getTime()) || exp.getTime() < Date.now()) {
        sh.deleteRow(i + 2);
        return null;
      }
      const u = findUserByUsername_(data[i][1]);
      if (!u) { sh.deleteRow(i + 2); return null; }
      return { username: u.record.username, isAdmin: u.record.is_admin };
    }
  }
  return null;
}

function destroySession_(token) {
  token = s_(token);
  if (!token) return;
  const sh = sessionsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (s_(data[i][0]) === token) { sh.deleteRow(i + 2); return; }
  }
}

function destroyAllSessionsForUser_(username, exceptToken) {
  const u = lc_(username);
  const sh = sessionsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, SESSIONS_HEADERS.length).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (lc_(data[i][1]) === u && s_(data[i][0]) !== s_(exceptToken)) {
      sh.deleteRow(i + 2);
    }
  }
}

function cleanExpiredSessions_() {
  const sh = sessionsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, SESSIONS_HEADERS.length).getValues();
  const now = Date.now();
  for (let i = data.length - 1; i >= 0; i--) {
    const exp = new Date(data[i][3]);
    if (isNaN(exp.getTime()) || exp.getTime() < now) sh.deleteRow(i + 2);
  }
}

function requireSession_(token) {
  const u = getSessionUser_(token);
  if (!u) throw new Error('Your session has expired. Please sign in again.');
  return u;
}
function requireAdmin_(token) {
  const u = requireSession_(token);
  if (!u.isAdmin) throw new Error('Admin access required.');
  return u;
}

// ============================ AUTH API ============================
function login(username, password) {
  username = s_(username);
  if (!username || !password) throw new Error('Username and password are required.');

  // Opportunistic cleanup of expired sessions (cheap, runs once per login).
  try { cleanExpiredSessions_(); } catch (e) {}

  const found = findUserByUsername_(username);
  // Always do a dummy hash if the user doesn't exist, to avoid timing leaks.
  const salt = found ? found.record.salt : 'no_such_user_salt';
  const givenHash = hashPassword_(password, salt);
  if (!found || !constantTimeEqual_(givenHash, found.record.password_hash)) {
    throw new Error('Incorrect username or password.');
  }

  const sess = createSession_(found.record.username);
  updateUserFields_(found.record.username, { last_login: new Date() });

  return {
    token:     sess.token,
    username:  found.record.username,
    isAdmin:   found.record.is_admin,
    expiresAt: sess.expiresAt.toISOString()
  };
}

function logout(token) {
  destroySession_(token);
  return { ok: true };
}

// Used on page load to validate an existing token from localStorage.
function ping(token) {
  const u = getSessionUser_(token);
  if (!u) return { ok: false };
  return { ok: true, username: u.username, isAdmin: u.isAdmin };
}

function changeMyPassword(token, oldPassword, newPassword) {
  const u = requireSession_(token);
  newPassword = String(newPassword || '');
  if (newPassword.length < 8) throw new Error('New password must be at least 8 characters.');
  const found = findUserByUsername_(u.username);
  if (!found) throw new Error('User not found.');
  const oldHash = hashPassword_(oldPassword, found.record.salt);
  if (!constantTimeEqual_(oldHash, found.record.password_hash)) {
    throw new Error('Current password is incorrect.');
  }
  const newSalt = generateSalt_();
  const newHash = hashPassword_(newPassword, newSalt);
  updateUserFields_(u.username, { password_hash: newHash, salt: newSalt });
  destroyAllSessionsForUser_(u.username, token);
  return { ok: true };
}

// ============================ USER MGMT (admin) ============================
function adminListUsers(token) {
  requireAdmin_(token);
  const sh = usersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, USERS_HEADERS.length).getValues();
  return data.filter(function (r) { return s_(r[0]); }).map(function (r) {
    return {
      username:  s_(r[0]),
      isAdmin:   r[3] === true || lc_(r[3]) === 'true',
      createdAt: r[4] ? new Date(r[4]).toISOString() : '',
      createdBy: s_(r[5]),
      lastLogin: r[6] ? new Date(r[6]).toISOString() : ''
    };
  });
}

function adminCreateUser(token, username, password, isAdmin) {
  const me = requireAdmin_(token);
  username = s_(username);
  password = String(password || '');
  if (!username) throw new Error('Username is required.');
  if (password.length < 8) throw new Error('Password must be at least 8 characters.');
  if (findUserByUsername_(username)) throw new Error('That username already exists.');
  const salt = generateSalt_();
  const hash = hashPassword_(password, salt);
  usersSheet_().appendRow([username, hash, salt, !!isAdmin, new Date(), me.username, '']);
  return { username: username };
}

function adminResetPassword(token, username, newPassword) {
  requireAdmin_(token);
  username = s_(username);
  newPassword = String(newPassword || '');
  if (newPassword.length < 8) throw new Error('New password must be at least 8 characters.');
  if (!findUserByUsername_(username)) throw new Error('User not found.');
  const salt = generateSalt_();
  const hash = hashPassword_(newPassword, salt);
  updateUserFields_(username, { password_hash: hash, salt: salt });
  destroyAllSessionsForUser_(username, '');
  return { ok: true };
}

function adminToggleAdmin(token, username, isAdmin) {
  const me = requireAdmin_(token);
  username = s_(username);
  if (lc_(username) === lc_(me.username) && !isAdmin) {
    throw new Error('You cannot remove your own admin role.');
  }
  if (!findUserByUsername_(username)) throw new Error('User not found.');
  updateUserFields_(username, { is_admin: !!isAdmin });
  return { ok: true };
}

function adminDeleteUser(token, username) {
  const me = requireAdmin_(token);
  username = s_(username);
  if (lc_(username) === lc_(me.username)) throw new Error('You cannot delete yourself.');
  const found = findUserByUsername_(username);
  if (!found) throw new Error('User not found.');
  usersSheet_().deleteRow(found.row);
  destroyAllSessionsForUser_(username, '');
  return { ok: true };
}

// ============================ GEO (token-gated) ============================
function getGeoData(token) {
  requireSession_(token);
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

function listGeo(token, filter) {
  requireAdmin_(token);
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

function getGeo(token, napId) {
  requireAdmin_(token);
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

function upsertGeo(token, entry) {
  const me = requireAdmin_(token);
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
          .setValues([[napId, city, brgy, location, now, me.username]]);
        invalidateGeoCache_();
        return { napId: napId, action: 'updated' };
      }
    }
  }
  sh.appendRow([napId, city, brgy, location, now, me.username]);
  invalidateGeoCache_();
  return { napId: napId, action: 'added' };
}

function deleteGeo(token, napId) {
  requireAdmin_(token);
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

function bulkUpsertGeoRows(token, rows) {
  const me = requireAdmin_(token);
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
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const nap = s_(r.napId);
    if (!nap) continue;
    const out = [nap, s_(r.city), s_(r.brgy), s_(r.location), now, me.username];
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

// ============================ SETUP ============================
function setupSheets() {
  const ss = openLookupSheet_();
  const need = [
    [CONFIG.TABS.USERS,         USERS_HEADERS],
    [CONFIG.TABS.USER_SESSIONS, SESSIONS_HEADERS],
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

  // Seed the first admin if USERS is empty.
  const usersSh = ss.getSheetByName(CONFIG.TABS.USERS);
  if (usersSh.getLastRow() < 2 && CONFIG.SEED_ADMIN && CONFIG.SEED_ADMIN.username) {
    const salt = generateSalt_();
    const hash = hashPassword_(CONFIG.SEED_ADMIN.password, salt);
    usersSh.appendRow([CONFIG.SEED_ADMIN.username, hash, salt, true, new Date(), 'setup', '']);
    return 'Setup complete. Seed admin created: ' + CONFIG.SEED_ADMIN.username +
           ' / ' + CONFIG.SEED_ADMIN.password + ' — CHANGE THIS PASSWORD on first login.';
  }
  return 'Setup complete. Tabs: ' + need.map(function (p) { return p[0]; }).join(', ');
}
