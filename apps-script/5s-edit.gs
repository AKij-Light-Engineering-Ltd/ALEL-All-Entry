/**
 * ============================================================================
 *  ALEL 5S Audit — EDIT / DELETE backend   (Google Apps Script Web App)
 *  File: apps-script/5s-edit.gs
 * ----------------------------------------------------------------------------
 *  Ki kore:
 *    • update  — ekta 5S audit row er value change kore (Total/Max/Score% nijei
 *                abar hiseb kore), ar "5S Edit Log" sheet e audit-trail rakhe
 *    • delete  — ekta 5S audit row delete kore (log e rekhe dey)
 *    • verify  — JSONP GET; dashboard save ta sotti hoyeche kina check kore
 *
 *  SETUP (ekbar, ~3 minute):
 *    1) script.google.com → New project → ei code ta paste korun (Code.gs)
 *    2) Save (💾) → Run ekbar `pingTest` → Google permission "Allow"
 *    3) Deploy → New deployment → type: Web app
 *         Execute as        : Me (your email)
 *         Who has access    : Anyone with a Google account     <-- ei ta din
 *       → Deploy → copy the /exec URL
 *       ("Anyone with a Google account" dile script caller er ASOL Google account
 *        check kore, tai EDITORS list er baire keu edit korte parbe na ✓)
 *    4) Oi /exec URL ta dashboard e paste korun:
 *         https://akij-light-engineering-ltd.github.io/ALEL-All-Entry/5s-dashboard.html
 *       → ✏️ Edit → setup box e URL paste → "Save URL"
 *    5) (Strongest) niche CONFIG e EDIT_KEY set korun — jemon 'ALEL-5S-2026'.
 *       Dashboard er setup box e oi code ekbar dile ei browser theke edit hobe;
 *       onno karo browser e Edit/Delete button-i asbe na ✓
 *
 *  NOTE: ei script ta tomar 5S FORM er script theke ALADA rakha bhalo —
 *        tahole form er /exec URL kono rokome change hobe na.
 * ============================================================================
 */

// ------------------------- CONFIG -------------------------------------------
var SHEET_ID   = '1iVlAdl9bcmZE1Uk_24z7NUl4FtivFTa2h7wOI1T8T0w';  // "ALEL All Entry Central Sheet"
var TAB_NAME   = '5S Form';        // je tab e 5S audit data ache
var LOG_TAB    = '5S Edit Log';    // audit trail (script nijei banabe)
var KEY_COLS   = ['Email', 'Audit Date', 'Floor / Area'];  // row identify korar key
var SCORE_RE   = /^\d\.\d\s/;      // "1.1 ..." = score column
var REMARK_RE  = /Remarks/i;
var MAX_PER_Q  = 3;                // protiti question er max score

// ------------------------- EDIT PERMISSION ----------------------------------
// Ei list er email chara keu edit/delete korte parbe na (server-side check).
// Khali array = sob allowed (recommended na).
var EDITORS = [
  'md.marufhossain@akijlightengineering.com'
];

// Extra lock (optional but recommended): ei code chara kono browser theke
// edit/delete kora jabe na. Code ta dashboard er "setup" box e ekbar dite hoy
// (browser e save thake). Eta khali rakhle sudhu Google account check cholbe.
var EDIT_KEY = '';   // e.g. 'ALEL-5S-2026' — set korle eta mandatory hoye jay

// Deployment access "Anyone with a Google account" dile script caller er asol
// Google account o check kore (Session.getActiveUser) — eta sob theke strong.
// "Anyone" dile caller account paowa jay na, tai payload er email check hoy.

// ------------------------- WEB ENTRY ----------------------------------------
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || 'ping');
  var res;
  try {
    if (action === 'verify')      res = doVerify_(p);
    else if (action === 'whoami') res = doWhoami_(p);
    else if (action === 'ping')   res = doPing_();
    else                          res = { ok: false, error: 'Unknown GET action: ' + action };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return respond_(res, e);
}

function doPost(e) {
  var res;
  try {
    var body = (e && e.postData && e.postData.contents) || '{}';
    var p = JSON.parse(body);
    var action = String(p.action || '');
    if (action === 'update')      res = doUpdate_(p);
    else if (action === 'delete') res = doDelete_(p);
    else                          res = { ok: false, error: 'Unknown POST action: ' + action };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return respond_(res, e);
}

/** JSONP support (dashboard theke response PORA jay) */
function respond_(obj, e) {
  var cb = e && e.parameter && e.parameter.callback;
  var text = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------- HELPERS ------------------------------------------
function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB_NAME);
  if (!sh) throw new Error('Tab pelam na: ' + TAB_NAME);
  return sh;
}
function logSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.appendRow(['When', 'Action', 'Row', 'By (name)', 'By (email)', 'Verified via', 'Key (email | date | floor)', 'Changes']);
    sh.getRange('A1:G1').setFontWeight('bold');
  }
  return sh;
}
function norm_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function day_(v) {
  if (v instanceof Date) {
    var t = new Date(v.getTime() - v.getTimezoneOffset() * 60000);
    return t.toISOString().slice(0, 10);
  }
  var m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? (m[1] + '-' + m[2] + '-' + m[3]) : String(v).trim();
}
/** Caller er asol Google account (deployment "Anyone with a Google account" hole). */
function callerEmail_() {
  try {
    var e = Session.getActiveUser().getEmail();
    return e ? String(e).trim().toLowerCase() : '';
  } catch (err) { return ''; }
}

/**
 * Edit permission gate.
 *  1) Google session er email ache  -> oi ta-i bichar hobe (sob theke strong)
 *  2) session nei (deployment "Anyone") -> payload er email check
 * Return: { ok, via, email, error }
 */
function editorOk_(p) {
  var list = (EDITORS || []).map(norm_);
  var code = String((p && p.code) || '');
  var sess = callerEmail_();

  // 1) Google session ache -> oi ta-i final (fake kora jay na)
  if (sess) {
    if (!list.length || list.indexOf(sess) >= 0) return { ok: true, via: 'google-session', email: sess };
    return {
      ok: false, via: 'google-session', email: sess,
      error: 'Ei Google account (' + sess + ') er edit permission nei. Only ' + EDITORS.join(', ') + ' — oi account diye login kore abar chesta korun.'
    };
  }

  // 2) EDIT_KEY set kora ache -> code mile kina dekho
  if (EDIT_KEY) {
    if (code && code === EDIT_KEY) return { ok: true, via: 'access-key', email: '(access code)' };
    return {
      ok: false, via: 'access-key', email: '(no session / no code)',
      error: 'Edit permission nei — tomar Google account chine na, ar access code o nei.'
    };
  }

  // 3) EDIT_KEY khali -> payload email check (weak fallback)
  var claim = norm_((p && p.by && p.by.email) || (p && p.email) || '');
  if (!list.length) return { ok: true, via: 'open', email: '(open)' };
  if (claim && list.indexOf(claim) >= 0) return { ok: true, via: 'email', email: claim };
  return {
    ok: false, via: 'email', email: claim || '(khali)',
    error: 'Edit permission nei — only ' + EDITORS.join(', ') + ' ei dashboard edit korte parbe.'
  };
}

/** Row khoja: row number age (key mile kina check kore), na mille key diye scan. */
function findRow_(values, head, key, rowNo) {
  var ci = {}; for (var i = 0; i < head.length; i++) ci[head[i]] = i;
  function match(r) {
    for (var k = 0; k < KEY_COLS.length; k++) {
      var col = KEY_COLS[k]; if (ci[col] === undefined) continue;
      var want = (col === 'Audit Date') ? day_(key && key.date) : norm_((key || {})[col === 'Email' ? 'email' : 'floor']);
      var have = (col === 'Audit Date') ? day_(r[ci[col]]) : norm_(r[ci[col]]);
      if (want !== '' && want !== have) return false;
    }
    return true;
  }
  var n = Number(rowNo || 0);
  if (n >= 2 && n <= values.length) {
    if (match(values[n - 1])) return n - 1;            // 0-based index
  }
  if (key) {
    for (var r = 1; r < values.length; r++) if (match(values[r])) return r;
  }
  return -1;
}

function rowToObj_(head, row) {
  var o = {};
  for (var i = 0; i < head.length; i++) o[head[i]] = (row[i] === undefined || row[i] === null) ? '' : String(row[i]);
  return o;
}

// ------------------------- ACTIONS ------------------------------------------
function doPing_() {
  var sh = sheet_();
  return { ok: true, info: 'ALEL 5S edit backend', tab: TAB_NAME, rows: sh.getLastRow(), editors: EDITORS };
}

/** Caller er permission jiggasha — dashboard modal kholar somoy ei action call kore. */
function doWhoami_(p) {
  var g = editorOk_(p);
  return {
    ok: true,
    allowed: g.ok,
    caller: g.email,
    via: g.via,
    editors: EDITORS,
    error: g.ok ? '' : g.error
  };
}

function doVerify_(p) {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var head = values[0].map(function (h) { return String(h).trim(); });
  var idx = findRow_(values, head, p.key || null, p.row);
  if (idx < 1) return { ok: false, error: 'Row pelam na' };
  return { ok: true, row: idx + 1, values: rowToObj_(head, values[idx]) };
}

function doUpdate_(p) {
  var gate = editorOk_(p);
  if (!gate.ok) return { ok: false, code: 'FORBIDDEN', error: gate.error, caller: gate.email };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var head = values[0].map(function (h) { return String(h).trim(); });
    var idx = findRow_(values, head, p.key || null, p.row);
    if (idx < 1) return { ok: false, error: 'Row khuje pelam na — Refresh kore abar chesta koro' };

    var row = values[idx].slice();
    var v = p.values || {};
    var changed = [];
    Object.keys(v).forEach(function (k) {
      var key = String(k).trim();
      if (key === 'sheet') return;
      var c = head.indexOf(key);
      if (c < 0) return;                                  // ei column nei → skip
      var oldV = row[c], newV = v[k];
      if (String(oldV) !== String(newV)) {
        changed.push(key + ': "' + oldV + '" -> "' + newV + '"');
        row[c] = newV;
      }
    });

    // Total / Max / Score % abar hiseb
    var total = 0, cnt = 0;
    for (var c2 = 0; c2 < head.length; c2++) {
      if (SCORE_RE.test(head[c2]) && !REMARK_RE.test(head[c2])) {
        cnt++;
        total += Number(row[c2]) || 0;
      }
    }
    var max = cnt * MAX_PER_Q;
    var tC = head.indexOf('Total Score'), mC = head.indexOf('Max Score'), pC = head.indexOf('Score %');
    if (tC >= 0) row[tC] = total;
    if (mC >= 0) row[mC] = max;
    if (pC >= 0) row[pC] = max ? Math.round(total / max * 100) : 0;

    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);

    logSheet_().appendRow([
      new Date(), 'UPDATE', idx + 1,
      (p.by && p.by.name) || '', gate.email,
      gate.via,
      [day_(row[head.indexOf('Email')]), day_(row[head.indexOf('Audit Date')]), row[head.indexOf('Floor / Area')]].join(' | '),
      changed.join(' ; ') || '(no value change)'
    ]);

    return { ok: true, row: idx + 1, total: total, max: max, pct: max ? Math.round(total / max * 100) : 0, changed: changed.length };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doDelete_(p) {
  var gate = editorOk_(p);
  if (!gate.ok) return { ok: false, code: 'FORBIDDEN', error: gate.error, caller: gate.email };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var head = values[0].map(function (h) { return String(h).trim(); });
    var idx = findRow_(values, head, p.key || null, p.row);
    if (idx < 1) return { ok: false, error: 'Row khuje pelam na' };

    var snap = rowToObj_(head, values[idx]);
    logSheet_().appendRow([
      new Date(), 'DELETE', idx + 1,
      (p.by && p.by.name) || '', gate.email,
      gate.via,
      [snap['Email'], day_(snap['Audit Date']), snap['Floor / Area']].join(' | '),
      'Total=' + snap['Total Score'] + ' Score%=' + snap['Score %'] + ' Auditor=' + snap['Auditor Name']
    ]);
    sh.deleteRow(idx + 1);
    return { ok: true, deletedRow: idx + 1 };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ------------------------- TEST ---------------------------------------------
/** Editor e ekbar Run korun — permission allow hoye jabe. */
function pingTest() {
  Logger.log(JSON.stringify(doPing_()));
}
