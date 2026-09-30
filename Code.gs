/************ CONFIG ************/
const CONFIG = {
  SHEET_NAME: 'Sheet1',              // tab name
  FOLDER_ID: 'PASTE_DRIVE_FOLDER_ID', // folder where all photos are stored
  PIN: '',                           // optional access code, e.g. '4821'. Leave '' to disable
  CACHE_SECONDS: 300
};
// Column order: A Last Updated | B Employee Details | C Designation | D Zone | E Ward | F Photo
/*******************************/

function doGet(e) {
  try {
    if (CONFIG.PIN && (e.parameter.pin || '') !== CONFIG.PIN) return out({ error: 'PIN' });
    return out({ rows: getRows() });
  } catch (err) { return out({ error: String(err) }); }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (CONFIG.PIN && (d.pin || '') !== CONFIG.PIN) return out({ error: 'PIN' });
    if (d.action !== 'save') return out({ error: 'Bad request' });
    return out(saveRecord(d));
  } catch (err) { return out({ error: String(err) }); }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- read (cached in chunks; CacheService limit is 100KB per key) ---------- */
function getRows() {
  const cache = CacheService.getScriptCache();
  const n = cache.get('rows_n');
  if (n) {
    const keys = []; for (let i = 0; i < +n; i++) keys.push('rows_' + i);
    const parts = cache.getAll(keys);
    if (keys.every(k => parts[k] !== undefined)) return JSON.parse(keys.map(k => parts[k]).join(''));
  }
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const rows = sh.getRange(2, 1, last - 1, 6).getValues().map(r => [
    String(r[1]), String(r[2]), String(r[3]), String(r[4]), String(r[5] || ''),
    r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd-MMM-yyyy HH:mm:ss') : String(r[0] || '')
  ]).filter(r => r[0]);
  const json = JSON.stringify(rows), size = 90000, obj = {};
  let c = 0; for (let i = 0; i < json.length; i += size) obj['rows_' + (c++)] = json.substr(i, size);
  obj.rows_n = String(c);
  try { cache.putAll(obj, CONFIG.CACHE_SECONDS); } catch (e) {}
  return rows;
}
function clearRowsCache() {
  const cache = CacheService.getScriptCache(), n = +(cache.get('rows_n') || 0), keys = ['rows_n'];
  for (let i = 0; i < n; i++) keys.push('rows_' + i);
  cache.removeAll(keys);
}

/* ---------- save ---------- */
function saveRecord(d) {
  const details = String(d.details || '').trim();
  if (!details) return { error: 'Employee missing' };

  // 1) Upload photo first (outside the lock: each employee has its own file, so no clash)
  let photoUrl = '';
  if (d.image) {
    const b64 = String(d.image).replace(/^data:image\/\w+;base64,/, '');
    const name = details.replace(/\s*\|\s*/g, '_').replace(/[\\/:*?"<>]/g, '_').trim() + '.jpg';
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);       // keep exactly one photo per employee
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    photoUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view';
  }

  // 2) Write to sheet under a lock (only this short step is serialized)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { error: 'BUSY' };
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    const cell = sh.getRange('B:B').createTextFinder(details).matchEntireCell(true).findNext();
    if (!cell || cell.getRow() < 2) return { error: 'Employee not found in sheet' };
    const row = cell.getRow();
    const now = new Date();
    const rng = sh.getRange(row, 3, 1, 4);                   // C:F
    const cur = rng.getValues()[0];
    rng.setValues([[String(d.designation || ''), String(d.zone || ''), String(d.ward || ''), photoUrl || cur[3]]]);
    sh.getRange(row, 1).setValue(now);                        // A: Last Updated Timestamp
    SpreadsheetApp.flush();
    clearRowsCache();
    return { ok: true, photoUrl: photoUrl, timestamp: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm:ss') };
  } finally { lock.releaseLock(); }
}
