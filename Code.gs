const VERSION = '1.4.0'; // LINE 1 = the only version line in this file. Bump on every change.

// EcoOils Project - Site Walkthrough Close-Out backend (Google Apps Script web app).
// Setup (once, on iggifernandes@gmail.com): Extensions > Apps Script > paste this file > run setup() >
//   run installDigestTrigger() > Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// PINs: 6 digits, set only through the email "Forgot PIN" flow. Never store a PIN or password in this file or in a sheet.

const APP_URL = 'https://iggifernandes.github.io/EcoOils-Project-Site-Walkthrough-Close-Out/';
const F_NAME = 'Findings';
const F_COLS = ['id', 'createdAt', 'ic', 'location', 'category', 'risk', 'description', 'dueDate', 'status',
  'beforePhotos', 'afterPhotos', 'closeNote', 'closedBy', 'closedAt', 'verifiedBy', 'verifiedAt', 'hseNote',
  'reopenCount', 'loggedBy', 'updatedAt'];
const C_NAME = 'Contacts';
const C_COLS = ['id', 'ic', 'name', 'email', 'role', 'updatedAt'];
const U_NAME = 'Users';
const U_COLS = ['id', 'name', 'perms', 'active', 'updatedAt']; // id = the person's email (lowercase)
// What a ticked box allows. Owners (ADMIN_EMAILS) always have everything.
// view=read, pdf=print/PDF report, csv=CSV export, log=log findings, verify=verify/send back/close/edit, resend=resend email,
// contactsView=see contacts, contacts=edit contacts + run digest, delete=delete findings, users=manage users.
const PERMS = ['view', 'pdf', 'csv', 'log', 'verify', 'resend', 'contactsView', 'contacts', 'delete', 'users'];
const DATE_COLS = ['dueDate'];
const RISKS = ['High', 'Medium', 'Low'];
const OPEN_STATUS = ['Open', 'Reopened'];
const DEVICE_DAYS = 60; // "keep me signed in on this device" lasts this many days

/* ---------- entry points ---------- */
function doGet(e) {
  const f = e && e.parameter && e.parameter.f;
  if (f === 'ping') return json_({ ok: true, version: VERSION }); // lets the Hub show a live version per app
  return json_({ ok: true, version: VERSION });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  try {
    switch (req.action) {
      case 'ping': return json_({ ok: true, version: VERSION });
      case 'requestReset': return json_(requestReset_(req));
      case 'setPin': return json_(setPin_(req));
      case 'login': return json_(login_(req));
      case 'resume': return json_(resume_(req));
      case 'logout': CacheService.getScriptCache().remove('sess_' + req.token); if (req.device) revokeDevices_('', s_(req.device, 100)); return json_({ ok: true });
      case 'icView': return json_(icView_(req));
      case 'icClose': return json_(icClose_(req));
      case 'me': return json_(me_(req));
      case 'list': requirePerm_(req, 'view'); return json_(list_());
      case 'photos': requirePerm_(req, 'view'); return json_(photos_(req));
      case 'create': return json_(create_(req, requirePerm_(req, 'log')));
      case 'update': requirePerm_(req, 'verify'); return json_(update_(req));
      case 'verify': return json_(setStatus_(req, requirePerm_(req, 'verify')));
      case 'adminClose': return json_(adminClose_(req, requirePerm_(req, 'verify')));
      case 'delete': requirePerm_(req, 'delete'); return json_(delete_(req));
      case 'resend': requirePerm_(req, 'resend'); return json_(resend_(req));
      case 'contactsList': requirePerm_(req, ['contactsView', 'contacts']); return json_(contactsList_());
      case 'ctSave': requirePerm_(req, 'contacts'); return json_(ctSave_(req));
      case 'ctDelete': requirePerm_(req, 'contacts'); return json_(ctDelete_(req));
      case 'digestNow': requirePerm_(req, 'contacts'); return json_(digest_({ force: true }));
      case 'usersList': requirePerm_(req, 'users'); return json_(usersList_());
      case 'userSave': return json_(userSave_(req, requirePerm_(req, 'users')));
      case 'userDelete': requirePerm_(req, 'users'); return json_(userDelete_(req));
      default: return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- template helpers (EcoOils Common Rules) ---------- */
// Every write goes through this, so two people saving at once cannot overwrite each other.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Every read of a whole sheet goes through this. After any write, call CacheService.getScriptCache().remove(key).
function cached_(key, seconds, loader) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const value = loader();
  const text = JSON.stringify(value);
  if (text.length < 90000) cache.put(key, text, seconds); // cache values are limited to 100 KB
  return value;
}

// Write many rows in one call (faster than appendRow or setValue in a loop).
function appendRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- sheet helpers ---------- */
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function s_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 2000); }
function clearCache_() { CacheService.getScriptCache().removeAll(['findings', 'contacts', 'users']); }

function tab_(name, cols) {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() < cols.length) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    cols.forEach(function (k, i) { if (DATE_COLS.indexOf(k) >= 0) sh.getRange(2, i + 1, 2000, 1).setNumberFormat('@'); });
  }
  return sh;
}

function readTab_(name, cols) {
  const sh = tab_(name, cols), n = sh.getLastRow();
  if (n < 2) return [];
  const tz = SpreadsheetApp.openById(prop_('SHEET_ID')).getSpreadsheetTimeZone();
  return sh.getRange(2, 1, n - 1, cols.length).getValues().map(function (v) {
    const o = {};
    cols.forEach(function (k, i) {
      let x = v[i];
      if (x instanceof Date) x = DATE_COLS.indexOf(k) >= 0 ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : x.toISOString();
      o[k] = x;
    });
    return o;
  }).filter(function (o) { return o.id; });
}

function findRow_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function findings_() { return cached_('findings', 60, function () { return readTab_(F_NAME, F_COLS); }); }
function contacts_() { return cached_('contacts', 300, function () { return readTab_(C_NAME, C_COLS); }); }

function writeRow_(sh, row, o, cols) { sh.getRange(row, 1, 1, cols.length).setValues([cols.map(function (c) { return o[c] == null ? '' : o[c]; })]); }

function nextId_() {
  const p = PropertiesService.getScriptProperties();
  const n = Number(p.getProperty('SEQ') || 0) + 1;
  p.setProperty('SEQ', String(n));
  return 'WT-' + ('0000' + n).slice(-4);
}

function isoDate_(v) {
  const t = s_(v, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) throw new Error('Bad date');
  return t;
}
function todayIso_() { return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd'); }
function fmtD_(iso) { return iso ? Utilities.formatDate(new Date(String(iso).slice(0, 10) + 'T12:00:00+07:00'), 'Asia/Jakarta', 'dd MMM yyyy') : ''; }
function daysBetween_(a, b) { return Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 86400000); }

/* ---------- photos (Drive) ---------- */
function savePhotos_(arr, prefix) {
  const folder = DriveApp.getFolderById(prop_('FOLDER_ID')), ids = [];
  (arr || []).slice(0, 3).forEach(function (uri, i) {
    const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+\/=]+)$/.exec(String(uri || ''));
    if (!m) throw new Error('Invalid photo');
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > 3 * 1024 * 1024) throw new Error('Photo too large');
    const blob = Utilities.newBlob(bytes, m[1], prefix + '-' + (i + 1) + '.' + m[1].split('/')[1]);
    ids.push(folder.createFile(blob).getId());
  });
  return ids;
}

function photoUri_(id) {
  const f = DriveApp.getFileById(id), it = f.getParents(), fid = prop_('FOLDER_ID');
  let ok = false;
  while (it.hasNext()) if (it.next().getId() === fid) ok = true;
  if (!ok) throw new Error('Not allowed');
  const b = f.getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

function idsOf_(csv) { return String(csv || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }

function photos_(r) {
  const out = {};
  (r.ids || []).slice(0, 6).forEach(function (id) { try { out[id] = photoUri_(s_(id, 80)); } catch (e) { out[id] = ''; } });
  return { ok: true, photos: out };
}

function trash_(ids) { ids.forEach(function (f) { try { DriveApp.getFileById(f).setTrashed(true); } catch (e) {} }); }

/* ---------- admin auth: email-code PIN reset, no stored passwords ---------- */
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, prop_('PIN_SALT') + '|' + s)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function adminEmails_() { return (prop_('ADMIN_EMAILS') || '').toLowerCase().split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function isAdminEmail_(e) { return adminEmails_().indexOf(String(e).toLowerCase().trim()) >= 0; }
function users_() { return cached_('users', 300, function () { return readTab_(U_NAME, U_COLS); }); }
// Who may sign in, and what they may do. null = no access. Looked up on every call, so a change takes effect at once.
function userInfo_(email) {
  email = String(email || '').toLowerCase().trim();
  if (!email) return null;
  if (isAdminEmail_(email)) return { email: email, admin: true, perms: PERMS.slice(), name: '' };
  const u = users_().filter(function (x) { return String(x.id).toLowerCase() === email && String(x.active) !== 'false' && String(x.active) !== 'FALSE'; })[0];
  if (!u) return null;
  const p = String(u.perms || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return PERMS.indexOf(x) >= 0; });
  return { email: email, admin: false, perms: p, name: u.name || '' };
}
function me_(r) { const i = userInfo_(requirePerm_(r, '')); if (!i) throw new Error('SESSION'); return { ok: true, email: i.email, admin: i.admin, perms: i.perms, name: i.name }; }

function requestReset_(r) {
  const email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  const rk = 'resetrate_' + email, n = Number(cache.get(rk) || 0);
  if (n >= 5) return { ok: true };
  cache.put(rk, String(n + 1), 3600);
  if (userInfo_(email)) {
    const code = '' + Math.floor(100000 + Math.random() * 900000);
    cache.put('reset_' + email, hash_('code|' + email + '|' + code), 900);
    mail_({ to: email, subject: 'EcoOils Walkthrough - PIN reset code', body: 'Your one-time code is ' + code + '. It expires in 15 minutes.\nIf you did not ask for it, ignore this email.', name: 'EcoOils Project HSSE' });
  }
  return { ok: true }; // same answer either way
}

function setPin_(r) {
  const email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  if (!/^\d{6}$/.test(String(r.pin || ''))) throw new Error('PIN must be exactly 6 digits');
  const fk = 'fail_' + email, fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Wait 15 minutes.');
  const want = cache.get('reset_' + email);
  if (!want || want !== hash_('code|' + email + '|' + s_(r.code, 10))) {
    cache.put(fk, String(fails + 1), 900);
    throw new Error('Wrong or expired code');
  }
  cache.remove('reset_' + email);
  PropertiesService.getScriptProperties().setProperty('PIN_' + email, hash_('pin|' + email + '|' + r.pin));
  revokeDevices_(email, ''); // a PIN reset signs the email out of every remembered device
  return { ok: true };
}

function login_(r) {
  const email = s_(r.email, 120).toLowerCase(), cache = CacheService.getScriptCache();
  const fk = 'fail_' + email, fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Too many attempts. Wait 15 minutes.');
  const stored = prop_('PIN_' + email);
  if (!userInfo_(email) || !stored) { cache.put(fk, String(fails + 1), 900); throw new Error('No PIN set yet or wrong details. Use "Forgot PIN / set PIN".'); }
  if (stored !== hash_('pin|' + email + '|' + r.pin)) { cache.put(fk, String(fails + 1), 900); throw new Error('Wrong PIN'); }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('sess_' + token, email, 21600);
  const inf = userInfo_(email);
  return { ok: true, token: token, email: email, perms: inf.perms, admin: inf.admin, name: inf.name, device: r.remember ? newDevice_(email) : '' };
}

/* Remembered devices: a random token is kept (hashed) in Script Properties for DEVICE_DAYS, so the browser can ask for a new session without the PIN. */
function newDevice_(email) {
  const p = PropertiesService.getScriptProperties(), all = p.getProperties(), now = Date.now();
  Object.keys(all).forEach(function (k) { if (k.indexOf('DEV_') === 0 && Number(String(all[k]).split('|')[1]) < now) p.deleteProperty(k); });
  const dev = Utilities.getUuid() + Utilities.getUuid();
  p.setProperty('DEV_' + hash_('dev|' + dev), email + '|' + (now + DEVICE_DAYS * 86400000));
  return dev;
}

// With dev: removes that one device. Without dev: removes every device of the email.
function revokeDevices_(email, dev) {
  const p = PropertiesService.getScriptProperties();
  if (dev) { p.deleteProperty('DEV_' + hash_('dev|' + dev)); return; }
  const all = p.getProperties();
  Object.keys(all).forEach(function (k) { if (k.indexOf('DEV_') === 0 && String(all[k]).split('|')[0] === email) p.deleteProperty(k); });
}

function resume_(r) {
  const key = 'DEV_' + hash_('dev|' + s_(r.device, 100)), v = prop_(key);
  if (!v) throw new Error('SESSION');
  const parts = v.split('|'), email = parts[0];
  if (Number(parts[1]) < Date.now() || !userInfo_(email)) { PropertiesService.getScriptProperties().deleteProperty(key); throw new Error('SESSION'); }
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, email, 21600);
  const inf = userInfo_(email);
  return { ok: true, token: token, email: email, perms: inf.perms, admin: inf.admin, name: inf.name };
}

// Returns the signed-in email. With a perm name, also checks that the person has that box ticked ('' = just signed in).
function requirePerm_(r, perm) {
  const e = r.token && CacheService.getScriptCache().get('sess_' + r.token);
  if (!e) throw new Error('SESSION');
  const i = userInfo_(e);
  if (!i) throw new Error('SESSION'); // access was removed
  const need = perm ? [].concat(perm) : [];
  if (need.length && !need.some(function (p) { return i.perms.indexOf(p) >= 0; })) throw new Error('You do not have access to this. Ask the HSSE Manager.');
  return e;
}

/* ---------- users and what they may do (Users tab) ---------- */
function usersList_() {
  return { ok: true, perms: PERMS, owners: adminEmails_(), users: users_().map(function (u) {
    return { email: String(u.id), name: u.name, perms: String(u.perms || '').split(',').filter(Boolean), active: String(u.active) !== 'false' && String(u.active) !== 'FALSE' };
  }) };
}

function userSave_(r, caller) {
  const u = r.user || {}, email = s_(u.email, 120).toLowerCase(), name = s_(u.name, 120);
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Name and a valid email are required');
  if (isAdminEmail_(email)) throw new Error('This email is an owner (ADMIN_EMAILS) and already has full access.');
  const me = userInfo_(caller);
  let perms = (u.perms || []).filter(function (x) { return PERMS.indexOf(x) >= 0; });
  if (perms.indexOf('users') >= 0 && !me.admin) throw new Error('Only an owner can give the Users access.');
  const active = u.active === false ? 'false' : 'true';
  return withLock_(function () {
    const sh = tab_(U_NAME, U_COLS), row = findRow_(sh, email), isNew = row < 0;
    const old = isNew ? null : sh.getRange(row, 1, 1, U_COLS.length).getValues()[0];
    if (!isNew && !me.admin && String(old[2]).split(',').indexOf('users') >= 0) perms = perms.concat(['users']).filter(function (x, i, a) { return a.indexOf(x) === i; }); // a non-owner cannot strip the Users access
    const o = { id: email, name: name, perms: perms.join(','), active: active, updatedAt: new Date().toISOString() };
    if (isNew) appendRows_(sh, [U_COLS.map(function (k) { return o[k]; })]); else writeRow_(sh, row, o, U_COLS);
    clearCache_();
    if (active === 'false') revokeDevices_(email, '');
    let mailed = false;
    if (isNew && active === 'true') {
      mailed = mail_({ to: email, subject: 'Site Walkthrough Close-Out - you now have access / Anda kini memiliki akses',
        body: 'Hello ' + name + ',\n\nYou have been given access to the Site Walkthrough Close-Out app.\n1. Open ' + APP_URL + '\n2. Tap "Forgot PIN / set PIN", enter this email address, and set your own 6-digit PIN with the code we email you.\n3. Sign in.' +
          '\n\nHalo ' + name + ',\n\nAnda telah diberi akses ke aplikasi Site Walkthrough Close-Out. Buka ' + APP_URL + ', ketuk "Forgot PIN / set PIN", masukkan email ini, atur PIN 6 digit Anda dengan kode yang kami kirim, lalu masuk.' + SIG, name: 'EcoOils Project HSSE', replyTo: prop_('REPLY_TO') || adminEmails_()[0] });
    }
    return { ok: true, mailed: mailed };
  });
}

function userDelete_(r) {
  const email = s_(r.email, 120).toLowerCase();
  return withLock_(function () {
    const sh = tab_(U_NAME, U_COLS), row = findRow_(sh, email);
    if (row > 0) sh.deleteRow(row);
    clearCache_();
    PropertiesService.getScriptProperties().deleteProperty('PIN_' + email);
    revokeDevices_(email, '');
    return { ok: true };
  });
}

/* ---------- close-out link token (derived from the finding id and the secret salt) ---------- */
function ctok_(id) { return hash_('close|' + id).slice(0, 16); }
function closeLink_(id) { return APP_URL + '?c=' + id + '.' + ctok_(id); }
function parseLink_(v) {
  const p = String(v || '').split('.'), id = p[0], tok = p[1] || '';
  if (!id || tok.length !== 16 || tok !== ctok_(id)) throw new Error('This link is not valid.');
  return id;
}

/* ---------- admin: findings ---------- */
function list_() {
  const rows = findings_().map(function (o) { const c = {}; Object.keys(o).forEach(function (k) { c[k] = o[k]; }); c.link = closeLink_(o.id); return c; });
  return { ok: true, rows: rows, today: todayIso_() };
}

function create_(r, email) {
  const f = r.fields || {};
  const need = ['ic', 'location', 'category', 'risk', 'description', 'dueDate'];
  for (let i = 0; i < need.length; i++) if (!s_(f[need[i]])) throw new Error('Missing: ' + need[i]);
  if (RISKS.indexOf(s_(f.risk)) < 0) throw new Error('Bad risk');
  if (s_(f.description).length < 10) throw new Error('Description too short');
  if (!r.before || !r.before.length) throw new Error('A before photo is required');
  const due = isoDate_(f.dueDate);
  const row = withLock_(function () {
    const id = nextId_(), now = new Date().toISOString();
    const ids = savePhotos_(r.before, id + '-before');
    const o = {
      id: id, createdAt: now, ic: s_(f.ic, 60).toUpperCase(), location: s_(f.location, 120), category: s_(f.category, 60), risk: s_(f.risk),
      description: s_(f.description, 1500), dueDate: due, status: 'Open', beforePhotos: ids.join(','), afterPhotos: '', closeNote: '', closedBy: '',
      closedAt: '', verifiedBy: '', verifiedAt: '', hseNote: '', reopenCount: 0, loggedBy: email, updatedAt: now
    };
    appendRows_(tab_(F_NAME, F_COLS), [F_COLS.map(function (c) { return o[c] == null ? '' : o[c]; })]);
    clearCache_();
    return o;
  });
  const mail = notifyNew_(row);
  return { ok: true, id: row.id, link: closeLink_(row.id), mail: mail };
}

const EDITABLE = ['ic', 'location', 'category', 'risk', 'description', 'dueDate', 'hseNote'];
function update_(r) {
  const f = r.fields || {};
  return withLock_(function () {
    const sh = tab_(F_NAME, F_COLS), row = findRow_(sh, r.id);
    if (row < 0) throw new Error('Finding not found');
    const o = readRow_(sh, row);
    EDITABLE.forEach(function (k) {
      if (f[k] === undefined) return;
      if (k === 'dueDate') o[k] = isoDate_(f[k]);
      else if (k === 'risk') { if (RISKS.indexOf(s_(f[k])) < 0) throw new Error('Bad risk'); o[k] = s_(f[k]); }
      else if (k === 'ic') o[k] = s_(f[k], 60).toUpperCase();
      else o[k] = s_(f[k], k === 'description' || k === 'hseNote' ? 1500 : 120);
    });
    o.updatedAt = new Date().toISOString();
    writeRow_(sh, row, o, F_COLS);
    clearCache_();
    return { ok: true };
  });
}

function readRow_(sh, row) {
  const v = sh.getRange(row, 1, 1, F_COLS.length).getValues()[0], o = {};
  F_COLS.forEach(function (k, i) { let x = v[i]; if (x instanceof Date) x = k === 'dueDate' ? Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd') : x.toISOString(); o[k] = x; });
  return o;
}

// HSSE verifies the IC's close-out, or sends it back.
function setStatus_(r, email) {
  const st = s_(r.status, 20), note = s_(r.note, 1000);
  if (['Verified', 'Reopened'].indexOf(st) < 0) throw new Error('Bad status');
  if (st === 'Reopened' && note.length < 5) throw new Error('Write what is still missing');
  let reopened = null;
  withLock_(function () {
    const sh = tab_(F_NAME, F_COLS), row = findRow_(sh, r.id);
    if (row < 0) throw new Error('Finding not found');
    const o = readRow_(sh, row), now = new Date().toISOString();
    o.status = st; o.updatedAt = now;
    if (st === 'Verified') { o.verifiedBy = email; o.verifiedAt = now; if (note) o.hseNote = note; }
    else { o.reopenCount = Number(o.reopenCount || 0) + 1; o.hseNote = note; o.verifiedBy = ''; o.verifiedAt = ''; reopened = o; }
    writeRow_(sh, row, o, F_COLS);
    clearCache_();
  });
  if (reopened) notifyReopen_(reopened);
  return { ok: true };
}

// HSSE closes a finding on behalf of the IC (for example the IC sent the after photo by WhatsApp). It is verified straight away.
function adminClose_(r, email) {
  const name = s_(r.closedBy, 120), note = s_(r.note, 1000);
  if (name.length < 2) throw new Error('Write who closed it');
  if (!r.after || !r.after.length) throw new Error('An after photo is required');
  return withLock_(function () {
    const sh = tab_(F_NAME, F_COLS), row = findRow_(sh, r.id);
    if (row < 0) throw new Error('Finding not found');
    const o = readRow_(sh, row), now = new Date().toISOString();
    const ids = savePhotos_(r.after, o.id + '-after-' + (idsOf_(o.afterPhotos).length + 1));
    o.afterPhotos = idsOf_(o.afterPhotos).concat(ids).join(',');
    o.closeNote = note; o.closedBy = name + ' (via HSSE)'; o.closedAt = now;
    o.status = 'Verified'; o.verifiedBy = email; o.verifiedAt = now; o.updatedAt = now;
    writeRow_(sh, row, o, F_COLS);
    clearCache_();
    return { ok: true };
  });
}

function delete_(r) {
  const ids = (r.ids || []).map(String);
  return withLock_(function () {
    const sh = tab_(F_NAME, F_COLS);
    let n = 0;
    ids.forEach(function (id) {
      const row = findRow_(sh, id);
      if (row < 0) return;
      const o = readRow_(sh, row);
      trash_(idsOf_(o.beforePhotos).concat(idsOf_(o.afterPhotos)));
      sh.deleteRow(row); n++;
    });
    clearCache_();
    return { ok: true, deleted: n };
  });
}

/* ---------- IC close-out (no login: the emailed link is the key; HSSE verifies afterwards) ---------- */
function icView_(r) {
  const id = parseLink_(r.c);
  const o = findings_().filter(function (x) { return x.id === id; })[0];
  if (!o) throw new Error('This finding was removed.');
  const before = idsOf_(o.beforePhotos).slice(0, 3).map(photoUri_);
  const after = idsOf_(o.afterPhotos).slice(-3).map(photoUri_);
  return {
    ok: true, today: todayIso_(), before: before, after: after,
    f: { id: o.id, ic: o.ic, location: o.location, category: o.category, risk: o.risk, description: o.description, dueDate: o.dueDate,
      status: o.status, closeNote: o.closeNote, closedBy: o.closedBy, hseNote: o.hseNote, reopenCount: o.reopenCount, createdAt: o.createdAt }
  };
}

function icClose_(r) {
  const id = parseLink_(r.c), name = s_(r.closedBy, 120), note = s_(r.note, 1000);
  if (name.length < 2) throw new Error('Write your name');
  if (!r.after || !r.after.length) throw new Error('An after photo is required');
  let o;
  withLock_(function () {
    const sh = tab_(F_NAME, F_COLS), row = findRow_(sh, id);
    if (row < 0) throw new Error('This finding was removed.');
    o = readRow_(sh, row);
    if (OPEN_STATUS.indexOf(o.status) < 0) throw new Error('This finding is already closed.');
    const ids = savePhotos_(r.after, id + '-after-' + (idsOf_(o.afterPhotos).length + 1));
    const now = new Date().toISOString();
    o.afterPhotos = idsOf_(o.afterPhotos).concat(ids).join(',');
    o.closeNote = note; o.closedBy = name; o.closedAt = now; o.status = 'Closed by IC'; o.updatedAt = now;
    writeRow_(sh, row, o, F_COLS);
    clearCache_();
  });
  notifyClosed_(o);
  return { ok: true };
}

/* ---------- contacts: who receives findings for each company ---------- */
function contactsList_() {
  const trig = (function () { try { return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'autoDigest'; }); } catch (e) { return null; } })();
  return { ok: true, contacts: contacts_(), digestOn: trig, digestDays: prop_('DIGEST_DAYS') || 'Mon,Tue,Wed,Thu,Fri,Sat', digestLast: prop_('DIGEST_LAST') || '' };
}

function ctSave_(r) {
  const c = r.contact || {}, email = s_(c.email, 120).toLowerCase(), ic = s_(c.ic, 60).toUpperCase(), name = s_(c.name, 120);
  if (!ic || !name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Company, name and a valid email are required');
  const role = c.role === 'PM' ? 'PM' : 'HSSE';
  return withLock_(function () {
    const sh = tab_(C_NAME, C_COLS), row = c.id ? findRow_(sh, c.id) : -1;
    const o = { id: row > 0 ? c.id : 'CT-' + Utilities.getUuid().slice(0, 6).toUpperCase(), ic: ic, name: name, email: email, role: role, updatedAt: new Date().toISOString() };
    if (row > 0) writeRow_(sh, row, o, C_COLS); else appendRows_(sh, [C_COLS.map(function (k) { return o[k]; })]);
    clearCache_();
    return { ok: true, id: o.id };
  });
}

function ctDelete_(r) {
  return withLock_(function () {
    const sh = tab_(C_NAME, C_COLS);
    (r.ids || []).forEach(function (id) { const row = findRow_(sh, id); if (row > 0) sh.deleteRow(row); });
    clearCache_();
    return { ok: true };
  });
}

function recipients_(ic) {
  const cs = contacts_().filter(function (c) { return String(c.ic).toUpperCase() === String(ic).toUpperCase() && c.email; });
  const hsse = cs.filter(function (c) { return c.role !== 'PM'; }).map(function (c) { return c.email; });
  const pm = cs.filter(function (c) { return c.role === 'PM'; }).map(function (c) { return c.email; });
  return { to: hsse.length ? hsse : pm, pm: hsse.length ? pm : [], names: cs.map(function (c) { return c.name; }) };
}

/* ---------- email ---------- */
const SIG = '\n\nThank you / Terima kasih,\nHSSE - EcoOils Project, Sei Mangkei';
function send_(opt) {
  const m = { to: opt.to, subject: opt.subject, body: opt.body, name: 'EcoOils Project HSSE' };
  if (opt.cc) m.cc = opt.cc;
  if (opt.attachments && opt.attachments.length) m.attachments = opt.attachments;
  const reply = prop_('REPLY_TO') || adminEmails_()[0];
  if (reply) m.replyTo = reply;
  return mail_(m);
}

/* One door for every email. If Script Properties RELAY_URL + RELAY_SECRET are set, the mail is
   handed to the Mail Relay script (owned by ptecooilsmentari@gmail.com) so it is sent from that
   account and uses ITS daily quota. If they are not set, it sends from this account as before.
   If the relay fails, nothing is sent from this account unless RELAY_FALLBACK = 1. */
function mail_(m) {
  const url = prop_('RELAY_URL'), secret = prop_('RELAY_SECRET');
  if (url && secret) {
    try {
      const payload = { secret: secret, to: m.to, cc: m.cc || '', subject: m.subject, body: m.body, replyTo: m.replyTo || '', name: m.name || '', attachments: [] };
      (m.attachments || []).forEach(function (b) {
        payload.attachments.push({ name: b.getName(), type: b.getContentType(), b64: Utilities.base64Encode(b.getBytes()) });
      });
      const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: true });
      const j = JSON.parse(res.getContentText() || '{}');
      if (j && j.ok) return true;
      console.error('Relay refused: ' + (j && j.error));
    } catch (e) { console.error('Relay error: ' + e); }
    if (prop_('RELAY_FALLBACK') !== '1') return false;
  }
  if (MailApp.getRemainingDailyQuota() < 3) return false;
  const l = { to: m.to, subject: m.subject, body: m.body, name: m.name || 'EcoOils Project HSSE' };
  if (m.cc) l.cc = m.cc;
  if (m.replyTo) l.replyTo = m.replyTo;
  if (m.attachments && m.attachments.length) l.attachments = m.attachments;
  MailApp.sendEmail(l);
  return true;
}

/* Run from the editor to prove the relay works: sends a test email to the first admin. */
function testRelay() {
  const to = adminEmails_()[0];
  Logger.log(mail_({ to: to, subject: 'Walkthrough relay test', body: 'If this arrived from ptecooilsmentari@gmail.com, the relay works.', name: 'EcoOils Project HSSE' }) ? 'SENT to ' + to : 'NOT SENT - check RELAY_URL / RELAY_SECRET');
}

function notifyNew_(o) {
  const rc = recipients_(o.ic);
  if (!rc.to.length) return { sent: 0, noContacts: true };
  const link = closeLink_(o.id);
  const subject = '[Walkthrough ' + o.id + '] ' + o.risk.toUpperCase() + ' - ' + o.location + ' - close by / tutup sebelum ' + fmtD_(o.dueDate);
  const en = 'Dear ' + rc.names.join(', ') + ',\n\nA site walkthrough finding has been raised for ' + o.ic + '.\n\nRef: ' + o.id + '\nRisk: ' + o.risk + '\nLocation: ' + o.location + '\nCategory: ' + o.category + '\nFinding: ' + o.description + '\nClose by: ' + fmtD_(o.dueDate) + '\n\nThe BEFORE photo is attached. Please fix it, then open the link below, add the AFTER photo and your name, and submit:\n' + link;
  const id = 'Yth. ' + rc.names.join(', ') + ',\n\nTemuan walkthrough lokasi kerja telah dicatat untuk ' + o.ic + '.\n\nNo. Ref: ' + o.id + '\nRisiko: ' + o.risk + '\nLokasi: ' + o.location + '\nKategori: ' + o.category + '\nTemuan: ' + o.description + '\nTutup sebelum: ' + fmtD_(o.dueDate) + '\n\nFoto SEBELUM terlampir. Mohon diperbaiki, lalu buka tautan di bawah, tambahkan foto SESUDAH dan nama Anda, kemudian kirim:\n' + link;
  const att = idsOf_(o.beforePhotos).map(function (fid) { try { return DriveApp.getFileById(fid).getBlob(); } catch (e) { return null; } }).filter(Boolean);
  const cc = (o.risk === 'High' ? rc.pm : []).concat(adminEmails_()).join(',');
  const ok = send_({ to: rc.to.join(','), cc: cc, subject: subject, body: en + '\n\n----------\n\n' + id + SIG, attachments: att });
  return { sent: ok ? rc.to.length : 0, noContacts: false, quota: !ok };
}

function notifyReopen_(o) {
  const rc = recipients_(o.ic);
  if (!rc.to.length) return;
  const link = closeLink_(o.id);
  const subject = '[Walkthrough ' + o.id + '] Not accepted, please redo / Belum diterima, mohon ulangi - ' + o.location;
  const en = 'HSSE has not accepted the close-out of ' + o.id + ' (' + o.location + ').\nReason: ' + o.hseNote + '\nNew close-by date: ' + fmtD_(o.dueDate) + '\nPlease fix it and submit a new AFTER photo here:\n' + link;
  const id = 'HSSE belum menerima penutupan ' + o.id + ' (' + o.location + ').\nAlasan: ' + o.hseNote + '\nBatas penutupan: ' + fmtD_(o.dueDate) + '\nMohon perbaiki dan kirim foto SESUDAH yang baru di sini:\n' + link;
  send_({ to: rc.to.join(','), cc: rc.pm.join(','), subject: subject, body: en + '\n\n----------\n\n' + id + SIG });
}

function notifyClosed_(o) {
  const to = adminEmails_().join(',');
  if (!to) return;
  send_({ to: to, subject: '[Walkthrough ' + o.id + '] Closed by ' + o.ic + ', please verify - ' + o.location,
    body: o.id + ' (' + o.location + ') was closed by ' + o.closedBy + ' (' + o.ic + ').\nNote: ' + (o.closeNote || '-') + '\n\nOpen the app > Findings to check the before and after photos, then Verify or send it back:\n' + APP_URL });
}

function resend_(r) {
  const o = findings_().filter(function (x) { return x.id === r.id; })[0];
  if (!o) throw new Error('Finding not found');
  if (OPEN_STATUS.indexOf(o.status) < 0) throw new Error('This finding is not open');
  return { ok: true, mail: notifyNew_(o) };
}

/* ---------- daily digest: one email per company listing its open and overdue findings ---------- */
function digest_(opt) {
  opt = opt || {};
  const today = todayIso_();
  if (!opt.force && prop_('DIGEST_LAST') === today) return { ok: true, results: [], skipped: 'already sent today' };
  const open = findings_().filter(function (o) { return OPEN_STATUS.indexOf(o.status) >= 0; });
  const by = {};
  open.forEach(function (o) { (by[o.ic] = by[o.ic] || []).push(o); });
  const results = [];
  Object.keys(by).forEach(function (ic) {
    const rc = recipients_(ic);
    if (!rc.to.length) { results.push({ ic: ic, status: 'noemail', open: by[ic].length }); return; }
    const items = by[ic].sort(function (a, b) { return String(a.dueDate).localeCompare(String(b.dueDate)); });
    const late = items.filter(function (o) { return o.dueDate < today; });
    const lines = items.map(function (o) {
      const d = daysBetween_(o.dueDate, today);
      const tag = d > 0 ? 'OVERDUE ' + d + ' d / LEWAT ' + d + ' h' : (d === 0 ? 'DUE TODAY / JATUH TEMPO HARI INI' : 'due ' + fmtD_(o.dueDate));
      return '- ' + o.id + ' | ' + o.risk + ' | ' + o.location + ' | ' + tag + '\n  ' + closeLink_(o.id);
    }).join('\n');
    const subject = '[Walkthrough] Open findings / Temuan terbuka - ' + ic + ': ' + items.length + ' open, ' + late.length + ' overdue (' + fmtD_(today) + ')';
    const body = 'Dear ' + rc.names.join(', ') + ',\n\nOpen walkthrough findings for ' + ic + ': ' + items.length + ' (overdue: ' + late.length + ').\nOpen each link, add the AFTER photo and submit to close it.\n\n' + lines +
      '\n\n----------\n\nTemuan walkthrough yang masih terbuka untuk ' + ic + ': ' + items.length + ' (lewat tenggat: ' + late.length + ').\nBuka setiap tautan, tambahkan foto SESUDAH lalu kirim untuk menutup.' + SIG;
    const cc = (late.length ? rc.pm : []).concat(adminEmails_()).join(',');
    const ok = send_({ to: rc.to.join(','), cc: cc, subject: subject, body: body });
    results.push({ ic: ic, status: ok ? 'sent' : 'quota', open: items.length, overdue: late.length });
  });
  if (!opt.force) PropertiesService.getScriptProperties().setProperty('DIGEST_LAST', today);
  return { ok: true, results: results };
}

/* A daily trigger calls this; it only sends on the days in script property DIGEST_DAYS (default Mon to Sat). */
function autoDigest() {
  const days = (prop_('DIGEST_DAYS') || 'Mon,Tue,Wed,Thu,Fri,Sat').split(',').map(function (x) { return x.trim(); });
  if (days.indexOf(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'EEE')) < 0) return;
  Logger.log(JSON.stringify(digest_({})));
}
/* Run once from the editor (select installDigestTrigger > Run, then Allow). Runs autoDigest every day about 07:00. */
function installDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoDigest').timeBased().everyDays(1).atHour(7).create();
  Logger.log('Daily digest trigger installed (about 07:00 Jakarta time).');
}
function removeDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoDigest') ScriptApp.deleteTrigger(t); });
}

/* ---------- one-time setup ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  const ss = p.getProperty('SHEET_ID') ? SpreadsheetApp.openById(p.getProperty('SHEET_ID')) : SpreadsheetApp.create('EcoOils Project - Site Walkthrough');
  p.setProperty('SHEET_ID', ss.getId());
  tab_(F_NAME, F_COLS); tab_(C_NAME, C_COLS); tab_(U_NAME, U_COLS);
  if (!p.getProperty('FOLDER_ID')) p.setProperty('FOLDER_ID', DriveApp.createFolder('EcoOils Project - Walkthrough photos').getId());
  if (!p.getProperty('ADMIN_EMAILS')) p.setProperty('ADMIN_EMAILS', 'Ignatius.Ambrosio@ecooils.com.my');
  if (!p.getProperty('PIN_SALT')) p.setProperty('PIN_SALT', Utilities.getUuid());
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('Next: run installDigestTrigger, then Deploy > New deployment > Web app (Execute as: Me, Access: Anyone).');
}
