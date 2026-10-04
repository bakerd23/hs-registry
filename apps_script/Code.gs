/**
 * High School Registry: shared decision store.
 *
 * A Google Apps Script web app backed by the Google Sheet it's attached to.
 * The matching page (GitHub Pages) reads and writes team decisions here.
 *
 * Sheet "decisions": one row per school name (alias_id), latest decision wins.
 * Sheet "log": every change, append-only, so nothing is ever lost.
 *
 * Setup: see ../README.md. The team passcode lives in Script Properties
 * (key PASSCODE) and never in the public repo.
 */

const HEADERS = ['alias_id', 'status', 'ceeb_code', 'decided_by', 'decided_at', 'server_ms'];
const LOG_HEADERS = ['server_time', 'alias_id', 'status', 'ceeb_code', 'decided_by', 'decided_at'];
const STATUSES = ['match', 'research', 'nohs', 'open'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Keep codes like "010370" as text so Sheets doesn't drop the leading zero.
    sh.getRange('A:C').setNumberFormat('@');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function authorized_(key) {
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  return pass && key === pass;
}

/** GET ?key=PASSCODE&since=<server_ms> → decisions changed after `since` (all of them when since=0). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!authorized_(p.key)) return json_({ ok: false, error: 'bad_passcode' });
  const since = Number(p.since || 0);
  const sh = sheet_('decisions', HEADERS);
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
  const out = [];
  for (const r of rows) {
    if (Number(r[5]) > since) out.push({ id: String(r[0]), s: r[1], c: String(r[2] || ''), u: r[3], t: r[4] });
  }
  return json_({ ok: true, now: Date.now(), rows: out });
}

/**
 * POST body (sent as text/plain to avoid a CORS preflight):
 * {"key": "...", "name": "Alex", "changes": [{"id": "<alias_id>", "s": "match", "c": "110505"}]}
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (!authorized_(body.key)) return json_({ ok: false, error: 'bad_passcode' });
  const name = String(body.name || '').slice(0, 60);
  const changes = (body.changes || []).filter(c =>
    /^[0-9a-f]{12}$/.test(c.id) && STATUSES.indexOf(c.s) >= 0 && (c.s !== 'match' || /^\d{6}$/.test(c.c)));
  if (!changes.length) return json_({ ok: false, error: 'no_valid_changes' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_('decisions', HEADERS);
    const log = sheet_('log', LOG_HEADERS);
    const n = sh.getLastRow() - 1;
    const ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
    const index = {};
    ids.forEach((id, i) => { index[id] = i + 2; });
    const now = Date.now();
    const iso = new Date(now).toISOString();
    const logRows = [];
    const append = [];
    for (const c of changes) {
      const row = [c.id, c.s, c.s === 'match' ? c.c : '', name, iso, now];
      if (index[c.id]) sh.getRange(index[c.id], 1, 1, HEADERS.length).setValues([row]);
      else { append.push(row); index[c.id] = -1; }
      logRows.push([iso, c.id, c.s, row[2], name, iso]);
    }
    if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, HEADERS.length).setValues(append);
    log.getRange(log.getLastRow() + 1, 1, logRows.length, LOG_HEADERS.length).setValues(logRows);
    return json_({ ok: true, now: now, saved: changes.length });
  } finally {
    lock.releaseLock();
  }
}
