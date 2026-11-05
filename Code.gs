/**
 * Code.gs — the spreadsheet backend for Slate.
 *
 * DEPLOY: Apps Script > paste Code.local.gs (made by `node make-local.js`,
 * it carries the real token and sheet id) > Deploy > New deployment > Web app
 * > execute as ME > access ANYONE. Paste the /exec url and the token into the
 * app under Settings > Spreadsheet, press Test.
 *
 * It runs as YOU, so it opens the spreadsheet with your own permission and
 * the app never sees a Google credential. The token is the only thing
 * standing between the url and your sheet — make it long and random.
 *
 * SEVERAL PHONES, ONE SHEET. Every tab the app keeps is a list of records,
 * one per row, each with an Id and an Updated column. A phone sends only the
 * records it changed, each with the row as the phone last saw it (its base);
 * this merges them in, record by record, under a lock:
 *   - a row whose base is still what the sheet holds is taken;
 *   - a row the sheet has changed since (another phone, or a hand edit) is
 *     not: its id goes back as a conflict, and the phone merges the two and
 *     sends the result;
 *   - a row with no base (a phone from before v11) is taken when its Updated
 *     is the same or later.
 * No row is ever removed. A deleted record stays as a row with Deleted
 * filled, so a phone that had it cannot bring it back.
 *
 * SAFE WRITES. A tab is written over in place, never cleared first, so a run
 * stopped half way leaves the old rows, not an empty tab. Once a day, before
 * its first write, each tab is copied to "<tab> backup" - unless the tab has
 * shrunk to under half the backup, which then stays as it was.
 *
 * STALE PHONES. A phone says which version of the app it is. The newest one
 * seen is kept in the Config tab (row "app"); a phone older than that is
 * refused, and told to reload, so an out-of-date copy never writes.
 *
 * DELIBERATELY GENERIC: the app sends tab names, headers and rows; a column
 * this does not know is written all the same, and a column the sheet has but
 * the app did not send is kept. Adding a field never requires a redeploy.
 */

var BACKEND_VERSION = 2;

/* ── configure ─────────────────────────────────────────────────────────── */
var SHEET_ID = 'PUT_YOUR_SPREADSHEET_ID_HERE';
var TOKEN    = 'PUT_A_LONG_RANDOM_STRING_HERE';
/* ──────────────────────────────────────────────────────────────────────── */

/* RUN THIS ONCE from the editor (pick "authorize" in the dropdown, press Run)
   so the web app has its permissions. */
function authorize() {
  SpreadsheetApp.openById(SHEET_ID).getName();
  return 'ok';
}

function doGet(e)  { return handle(e, null); }
function doPost(e) {
  var body = null;
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = null; }
  return handle(e, body);
}

function handle(e, body) {
  var p = (e && e.parameter) || {};
  var token = (body && body.token) || p.token;
  if (token !== TOKEN) return json({ ok: false, error: 'bad token' });

  var action = (body && body.action) || p.action || 'ping';
  try {
    if (action === 'ping') return json({ ok: true, backend: BACKEND_VERSION });
    if (action === 'pull') return json({ ok: true, backend: BACKEND_VERSION, tabs: pull() });
    if (action === 'sync') return json(Object.assign({ ok: true, backend: BACKEND_VERSION }, sync(body)));
    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function book() { return SpreadsheetApp.openById(SHEET_ID); }

/* A cell as text: a date someone typed into the sheet by hand comes back as
   a Date, and the app wants YYYY-MM-DD. */
function cell(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || 'UTC', 'yyyy-MM-dd');
  return v;
}

/* A sheet's values, or nothing for a tab that was never written. */
function values(sh) {
  var v = sh.getDataRange().getValues();
  if (!v.length || !v[0].join('')) return [];
  return v;
}

/* Every tab, as { name, headers, rows } - the backups stay in the sheet. */
function pull() {
  var ss = book(), tz = ss.getSpreadsheetTimeZone();
  return ss.getSheets().filter(function (sh) { return !/ backup$/.test(sh.getName()); }).map(function (sh) {
    var v = values(sh).map(function (r) { return r.map(function (x) { return cell(x, tz); }); });
    return { name: sh.getName(), headers: v.length ? v[0] : [], rows: v.slice(1) };
  });
}

/* Merge what a phone sent, then hand every tab back, all under one lock so
   two phones syncing at once take turns. */
function sync(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ss = book(), wrote = {}, conflicts = {};
    /* an out-of-date phone writes nothing */
    var app = Number(body && body.app) || 0, newest = Number(config(ss, 'app')) || 0;
    if (app < newest) throw new Error('Slate v' + newest + ' is out - reload the app (About, Hard reload) to sync.');
    if (app > newest) setConfig(ss, 'app', String(app));
    ((body && body.tabs) || []).forEach(function (t) {
      backup(ss, t.name);
      var r = mergeTab(ss, t);
      wrote[t.name] = r.took;
      if (r.conflicts.length) conflicts[t.name] = r.conflicts;
    });
    SpreadsheetApp.flush();
    return { wrote: wrote, conflicts: conflicts, tabs: pull() };
  } finally {
    lock.releaseLock();
  }
}

/* ── Config: one value a row (Id, Value, Updated) ─────────────────────── */
function configRows(ss) {
  var sh = ss.getSheetByName('Config');
  return sh ? values(sh) : [];
}
function config(ss, id) {
  var v = configRows(ss);
  if (!v.length) return '';
  var h = v[0].map(String), I = h.indexOf('Id'), V = h.indexOf('Value');
  for (var i = 1; i < v.length; i++) if (String(v[i][I]) === id) return v[i][V];
  return '';
}
function setConfig(ss, id, value) {
  mergeTab(ss, { name: 'Config', headers: ['Id', 'Value', 'Updated'], rows: [[id, value, new Date().toISOString()]] });
}

/* ── the daily backup ─────────────────────────────────────────────────── */
function today(ss) { return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'UTC', 'yyyy-MM-dd'); }
function backup(ss, name) {
  var day = today(ss), key = 'backup ' + name;
  if (String(config(ss, key)) === day) return;
  var sh = ss.getSheetByName(name);
  var v = sh ? values(sh) : [];
  if (v.length < 2) return;                    /* nothing to copy yet: try again at the next write */
  var bname = name + ' backup', bk = ss.getSheetByName(bname);
  var had = bk ? values(bk).length : 0;
  /* a tab emptied (or mostly) does not overwrite the copy taken before */
  if (v.length * 2 >= had) {
    bk = bk || ss.insertSheet(bname);
    bk.clearContents();
    var range = bk.getRange(1, 1, v.length, v[0].length);
    range.setNumberFormat('@');
    range.setValues(v);
  }
  setConfig(ss, key, day);
}

/* t: { name, headers, rows, bases }. Rows are keyed by the Id column. An
   incoming row with a base (bases[i]: the row as that phone last saw it,
   header to value) is taken when the sheet's row still matches the base on
   every column the phone sent, else its id is a conflict. With no base, it is
   taken when its Updated is the same or later. Rows keep their place; new
   ones go at the bottom. Returns { took, conflicts }. */
function mergeTab(ss, t) {
  var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
  var tz = ss.getSpreadsheetTimeZone();
  var v = values(sh);
  var have = v.length ? v[0].map(String) : [];
  var heads = (t.headers || []).map(String);
  have.forEach(function (h) { if (h && heads.indexOf(h) < 0) heads.push(h); });
  var ID = heads.indexOf('Id'), UP = heads.indexOf('Updated');
  if (ID < 0) throw new Error(t.name + ': no Id column');

  /* the sheet's rows, laid out on the merged headers */
  /* a row typed by hand with no Id gets one, so it is never written away */
  var rows = [], at = {}, stamp = new Date().getTime().toString(36);
  v.slice(1).forEach(function (r, n) {
    var o = heads.map(function (h) { var i = have.indexOf(h); return i < 0 ? '' : cell(r[i], tz); });
    if (!o.join('')) return;
    if (!String(o[ID] || '')) o[ID] = 'h' + stamp + n;
    var id = String(o[ID]);
    at[id] = rows.length; rows.push(o);
  });

  var took = 0, conflicts = [], sent = (t.headers || []).map(String);
  var str = function (x) { return x == null ? '' : String(x); };
  (t.rows || []).forEach(function (r, k) {
    var o = heads.map(function (h) { var i = (t.headers || []).indexOf(h); return i < 0 ? null : r[i]; });
    var id = String(o[ID] || '');
    if (!id) return;
    if (at[id] == null) {
      at[id] = rows.length; rows.push(o.map(function (x) { return x == null ? '' : x; })); took++;
      return;
    }
    var cur = rows[at[id]], base = t.bases && t.bases[k];
    if (base) {
      for (var j = 0; j < sent.length; j++) {
        var hi = heads.indexOf(sent[j]);
        if (str(cur[hi]) !== str(base[sent[j]])) { conflicts.push(id); return; }
      }
    } else if (UP >= 0 && String(o[UP] || '') < String(cur[UP] || '')) return;
    rows[at[id]] = o.map(function (x, i) { return x == null ? cur[i] : x; }); took++;
  });

  var out = [heads].concat(rows).map(function (r) {
    return r.map(function (x) { var s = x == null ? '' : x; return typeof s === 'string' && /^=/.test(s) ? "'" + s : s; });
  });
  /* written over in place: a stop half way leaves rows, never an empty tab;
     what lay beyond the new rows (rows removed by hand) is cleared after */
  var range = sh.getRange(1, 1, out.length, heads.length);
  range.setNumberFormat('@');
  range.setValues(out);
  if (v.length > out.length) sh.getRange(out.length + 1, 1, v.length - out.length, Math.max(heads.length, have.length)).clearContent();
  sh.setFrozenRows(1);
  return { took: took, conflicts: conflicts };
}
