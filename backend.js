/**
 * backend.js — the only file in Slate that touches the network.
 *
 * It talks to ONE url: the Apps Script /exec you paste in under Settings >
 * Data > Spreadsheet. No analytics, no CDN, nothing else. The script runs as
 * YOU on Google's side, so it opens your spreadsheet with your own permission
 * and this code never sees a Google credential. The token is the only thing
 * between the url and your sheet.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SlateBackend = factory();
}(typeof self !== 'undefined' ? self : globalThis, function () {

  /* Apps Script answers a cross-origin POST only for simple content types, so
     the body goes as text/plain and the script parses it. A JSON content type
     triggers a preflight that /exec does not answer. */
  async function call(cfg, action, extra) {
    if (!cfg || !cfg.url) throw new Error('No script url.');
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(cfg.url)) {
      throw new Error('That is not an Apps Script /exec url.');
    }
    const body = Object.assign({ token: cfg.token || '', action }, extra || {});
    let res;
    try {
      res = await fetch(cfg.url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body),
        redirect: 'follow',
      });
    } catch (e) {
      /* The browser's "Failed to fetch": no response at all. With Apps
         Script that is nearly always the deployment answering with a login
         redirect, which carries no CORS headers - access is "only myself" or
         "anyone with a Google account" instead of "anyone". */
      throw new Error('Could not reach the script. Deployment access must be "Anyone", the url the /exec one - or you are offline.');
    }
    if (!res.ok) throw new Error('Backend said ' + res.status + '.');
    const text = await res.text();
    let j;
    try { j = JSON.parse(text); }
    catch (e) {
      /* Almost always the deployment's access setting: Google serves a login
         page instead of the script, and a login page is not JSON. */
      throw new Error('Not JSON — is the deployment set to "anyone with the link"?');
    }
    if (!j.ok) throw new Error(j.error || 'Backend refused.');
    return j;
  }

  const ping = (cfg) => call(cfg, 'ping');
  const pull = (cfg) => call(cfg, 'pull');
  /* tabs: [{ name, headers, rows, bases }] — only the records this phone changed,
     each with the row as this phone last saw it.
     The script merges them into the sheet record by record (the later
     Updated wins) and answers with every tab, merged. */
  /* app: this copy's version - one older than the newest seen is refused */
  const sync = (cfg, tabs, app) => call(cfg, 'sync', { tabs, app });

  return { call, ping, pull, sync };
}));
