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
 * records it changed; this merges them in, record by record, under a lock:
 * the row with the later Updated wins, every other row is left as it is. So
 * two people editing two different tasks at the same moment both keep their
 * edit. A deleted record stays as a row with Deleted filled, so a phone that
 * had it cannot bring it back.
 *
 * DELIBERATELY GENERIC: the app sends tab names, headers and rows; a column
 * this does not know is written all the same, and a column the sheet has but
 * the app did not send is kept. Adding a field never requires a redeploy.
 */

var BACKEND_VERSION = 1;

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

/* Every tab, as { name, headers, rows }. */
function pull() {
  var ss = book(), tz = ss.getSpreadsheetTimeZone();
  return ss.getSheets().map(function (sh) {
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
    var ss = book(), wrote = {};
    ((body && body.tabs) || []).forEach(function (t) { wrote[t.name] = mergeTab(ss, t); });
    SpreadsheetApp.flush();
    return { wrote: wrote, tabs: pull() };
  } finally {
    lock.releaseLock();
  }
}

/* t: { name, headers, rows }. Rows are keyed by the Id column; an incoming
   row replaces the sheet's row only when its Updated is the same or later.
   Rows keep their place; new ones go at the bottom. Returns how many rows
   were taken. */
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

  var took = 0;
  (t.rows || []).forEach(function (r) {
    var o = heads.map(function (h) { var i = (t.headers || []).indexOf(h); return i < 0 ? null : r[i]; });
    var id = String(o[ID] || '');
    if (!id) return;
    if (at[id] == null) {
      at[id] = rows.length; rows.push(o.map(function (x) { return x == null ? '' : x; })); took++;
      return;
    }
    var cur = rows[at[id]];
    if (UP >= 0 && String(o[UP] || '') < String(cur[UP] || '')) return;
    rows[at[id]] = o.map(function (x, i) { return x == null ? cur[i] : x; }); took++;
  });

  var out = [heads].concat(rows).map(function (r) {
    return r.map(function (x) { var s = x == null ? '' : x; return typeof s === 'string' && /^=/.test(s) ? "'" + s : s; });
  });
  sh.clearContents();
  var range = sh.getRange(1, 1, out.length, heads.length);
  range.setNumberFormat('@');
  range.setValues(out);
  sh.setFrozenRows(1);
  return took;
}
