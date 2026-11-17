/**
 * ALEL Operation Bulletin - INSTANT SYNC
 * -------------------------------------
 * Problem it solves: the GitHub Action "Auto-Sync Bulletin from Google Sheets" is
 * scheduled every 5 minutes, but GitHub heavily throttles scheduled runs on this
 * repo, so in practice the page only refreshed every 4-6 hours.
 *
 * This file makes a sheet edit reach the live page in ~40 seconds.
 *
 * SETUP (one time, ~60 seconds)
 * 1. Open the Google Sheet  ->  Extensions  ->  Apps Script.
 *    (Use the SAME script project that already powers the bulletin backend.)
 * 2. Add this file (File > New > Script, name it PgListInstantSync) and paste all of this.
 * 3. Create a GitHub fine-grained token:
 *       GitHub > Settings > Developer settings > Personal access tokens > Fine-grained tokens
 *       - Repository access: only  AKij-Light-Engineering-Ltd/ALEL-All-Entry
 *       - Permissions: Actions = Read and write     (nothing else needed)
 *       - Expiration: 365 days or less  (the AKij org rejects tokens with >366 days)
 * 4. Apps Script > Project Settings > Script properties > add property
 *       GH_PAT = <paste the token>          (never paste it in chat or in the sheet)
 * 5. Run setupInstantSync() once and approve the permission prompt.
 *
 * After that: edit PG LIST (or any product tab) -> the page updates itself.
 * Note: onEdit fires for MANUAL edits. Bulk paste/imports from another script do
 * not fire it - use testDispatchNow() or the "Run workflow" button in that case.
 */

var SYNC_REPO = 'AKij-Light-Engineering-Ltd/ALEL-All-Entry';
var SYNC_WORKFLOW = 'sync-bulletin.yml';
var SYNC_REF = 'main';
var SYNC_SHEET_TAB = 'PG LIST';       // '' = any tab of this spreadsheet
var SYNC_MIN_GAP_MS = 60000;          // max one dispatch per minute

/** Run this once to install the trigger. */
function setupInstantSync() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('GH_PAT')) {
    throw new Error('Add script property GH_PAT first (Project Settings > Script properties).');
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onSheetEditSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSheetEditSync')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  return 'Trigger installed. Edit a cell in ' + SYNC_SHEET_TAB + ' to test.';
}

/** Installable onEdit handler. */
function onSheetEditSync(e) {
  try {
    if (SYNC_SHEET_TAB && e && e.range) {
      var sh = e.range.getSheet();
      if (sh && sh.getName() !== SYNC_SHEET_TAB) return;
    }
    dispatchBulletinSync_();
  } catch (err) {
    console.error('onSheetEditSync: ' + err);
  }
}

/** Manual test - run from the editor any time. */
function testDispatchNow() {
  return dispatchBulletinSync_();
}

function dispatchBulletinSync_() {
  var props = PropertiesService.getScriptProperties();
  var pat = props.getProperty('GH_PAT');
  if (!pat) return 'GH_PAT script property is missing.';

  var now = Date.now();
  var last = Number(props.getProperty('LAST_DISPATCH') || 0);
  if (now - last < SYNC_MIN_GAP_MS) return 'Skipped (already dispatched < 1 min ago).';
  props.setProperty('LAST_DISPATCH', String(now));

  var url = 'https://api.github.com/repos/' + SYNC_REPO +
            '/actions/workflows/' + SYNC_WORKFLOW + '/dispatches';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + pat,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify({ ref: SYNC_REF }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code === 204 || code === 200) return 'Sync dispatched. Page updates in ~40s.';
  return 'Dispatch failed (' + code + '): ' + res.getContentText().slice(0, 300);
}
