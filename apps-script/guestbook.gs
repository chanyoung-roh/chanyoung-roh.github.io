/**
 * Bear guestbook backend — paste into Google Sheets → Extensions → Apps Script, then deploy as a Web app.
 * Moderation: delete a row, or type anything in column G ("hidden") to hide that note.
 */
const SHEET_NAME = 'notes';
const ANCHORS = ['hero', 'statement', 'about', 'work', 'gallery', 'contact'];
const MAX_NAME = 10, MAX_MEMO = 8;
const MAX_PER_MIN = 12;     // all visitors together, per minute (cheap spam brake)
const MAX_NOTES = 500;      // newest N notes are served
const TTL_HOURS = 48;       // notes older than this are no longer served (rows stay in the sheet)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['time', 'name', 'memo', 'anchor', 'x', 'y', 'hidden']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// strip control / invisible / bidi characters, collapse spaces, enforce length by code point (한글 1자 = 1)
function clean_(s, max) {
  s = String(s == null ? '' : s)
    .replace(/[\u0000-\u001F\u007F\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g, '')
    .replace(/\s+/g, ' ').trim();
  const n = Array.from(s).length;
  return n >= 1 && n <= max ? s : null;
}
// never let a visitor's text become a spreadsheet formula
function cell_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function uncell_(s) { s = String(s); return s.charAt(0) === "'" ? s.slice(1) : s; }

function doGet() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return out_({ ok: true, notes: [] });
  const start = Math.max(2, last - MAX_NOTES + 1);
  const rows = sh.getRange(start, 1, last - start + 1, 7).getValues();
  const notes = rows
    .filter(r => r[1] !== '' && r[2] !== '' && String(r[6]).trim() === '' && ANCHORS.indexOf(r[3]) >= 0)
    .filter(r => Date.now() - new Date(r[0]).getTime() < TTL_HOURS * 3600 * 1000)
    .map(r => ({ t: new Date(r[0]).getTime(), name: uncell_(r[1]), memo: uncell_(r[2]), a: r[3], x: Number(r[4]), y: Number(r[5]) }));
  return out_({ ok: true, notes: notes });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  if (d.site) return out_({ ok: true });                       // honeypot filled → bot; pretend success
  const name = clean_(d.name, MAX_NAME), memo = clean_(d.memo, MAX_MEMO);
  if (!name || !memo) return out_({ ok: false, error: 'length' });
  if (ANCHORS.indexOf(d.a) < 0) return out_({ ok: false, error: 'anchor' });
  const x = Math.round(Number(d.x) * 1e4) / 1e4, y = Math.round(Number(d.y) * 1e4) / 1e4;
  if (!(x >= 0 && x <= 1 && y >= 0 && y <= 1)) return out_({ ok: false, error: 'position' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return out_({ ok: false, error: 'busy' });
  try {
    const cache = CacheService.getScriptCache(), key = 'n' + Math.floor(Date.now() / 60000);
    const n = Number(cache.get(key) || 0);
    if (n >= MAX_PER_MIN) return out_({ ok: false, error: 'busy' });
    cache.put(key, String(n + 1), 120);
    const now = new Date();
    sheet_().appendRow([now, cell_(name), cell_(memo), d.a, x, y, '']);
    return out_({ ok: true, note: { t: now.getTime(), name: name, memo: memo, a: d.a, x: x, y: y } });
  } finally {
    lock.releaseLock();
  }
}
