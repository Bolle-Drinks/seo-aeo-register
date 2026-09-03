/**
 * Backend for the BOLLE SEO & AEO worklist.
 *
 * Reads and writes the "BOLLE SEO & AEO worklist" spreadsheet so the
 * register page at https://bolle-drinks.github.io/seo-aeo-register/
 * can save ticks and names for everyone, not just per browser.
 *
 * Deploy: Extensions > Apps Script from the sheet, paste this in,
 * then Deploy > New deployment > Web app,
 *   Execute as:  Me
 *   Who has access:  Anyone
 * Copy the /exec URL it gives you.
 *
 * Column map (row 1 is the header):
 *   A Done | B Task ID | C Group | D Who can do it | E Task
 *   F Done by | G Date done | H Register ref | I Notes
 */

var TOKEN = '8jDH_3Er5ETO4RE0';   // must match the token in the page
var COL = { done: 1, id: 2, who: 6, date: 7 };

// The worklist tab is sheet index 0 and every read and write assumes it.
// Adding a second tab makes that an ordering dependency rather than a fact,
// so this checks the sheet is the worklist before returning it, and throws
// rather than returning the wrong one.
var WORKLIST_HEADER = ['Done','Task ID','Group','Who can do it','Task',
                       'Done by','Date done','Register ref','Notes'];

function sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (!sh) throw new Error('worklist_tab_missing');
  var head = sh.getRange(1, 1, 1, WORKLIST_HEADER.length).getValues()[0];
  for (var i = 0; i < WORKLIST_HEADER.length; i++) {
    if (String(head[i]).trim() !== WORKLIST_HEADER[i]) {
      throw new Error('worklist_tab_mismatch: sheet 0 is "' + sh.getName() +
                      '", column ' + (i + 1) + ' header is "' + head[i] +
                      '", expected "' + WORKLIST_HEADER[i] + '"');
    }
  }
  return sh;
}

function readState_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return {};
  var rows = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][COL.id - 1] || '').trim();
    if (!id) continue;
    var doneCell = rows[i][COL.done - 1];
    var done = doneCell === true ||
               String(doneCell).toLowerCase() === 'true' ||
               String(doneCell).toLowerCase() === 'y' ||
               String(doneCell).toLowerCase() === 'yes' ||
               String(doneCell) === '✓';
    var d = rows[i][COL.date - 1];
    out[id.toLowerCase()] = {
      done: done,
      who: String(rows[i][COL.who - 1] || '').trim(),
      date: d instanceof Date ? Utilities.formatDate(d, 'UTC', 'dd MMM') : String(d || '')
    };
  }
  return out;
}

function writeOne_(id, done, who) {
  var sh = sheet_();
  var last = sh.getLastRow();
  var ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim().toLowerCase() === String(id).trim().toLowerCase()) {
      var row = i + 2;
      if (done !== null && done !== undefined) {
        sh.getRange(row, COL.done).setValue(done ? true : false);
        sh.getRange(row, COL.date).setValue(done ? new Date() : '');
      }
      if (who !== null && who !== undefined) {
        sh.getRange(row, COL.who).setValue(String(who).slice(0, 40));
      }
      return true;
    }
  }
  return false;
}

/* ---------- SUGGESTIONS. Additive. Nothing below touches the worklist tab. ---------- */

var SUGGEST_TAB = 'Suggestions';
var SUGGEST_MIN_INTERVAL_MS = 20000;              // 20 seconds, global
var SUGGEST_ALLOWED_KEYS = { token: 1, action: 1, name: 1, text: 1 };
var SUGGEST_NAME_RE = /^[A-Za-z][A-Za-z '\-]{0,39}$/;
var SUGGEST_TEXT_MAX = 500;

// Resolves the Suggestions tab BY NAME and returns null if it is absent.
// There is deliberately no fallback to sheet_() or getSheets()[0]: the
// absence of a fallback is what makes a write to the worklist tab
// unreachable from this path, rather than merely unintended.
function suggestSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUGGEST_TAB);
}

function readSuggestions_() {
  var sh = suggestSheet_();
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var when = rows[i][0];
    var name = String(rows[i][1] || '').trim();
    var text = String(rows[i][2] || '').trim();
    if (!text) continue;
    out.push({
      when: when instanceof Date ? Utilities.formatDate(when, 'UTC', 'dd MMM') : String(when || ''),
      name: name,
      text: text
    });
  }
  return out;
}

function addSuggestion_(name, text) {
  var sh = suggestSheet_();
  if (!sh) return 'no_tab';
  var last = sh.getLastRow();
  if (last >= 2) {
    var prev = sh.getRange(last, 1).getValue();
    if (prev instanceof Date && (new Date().getTime() - prev.getTime()) < SUGGEST_MIN_INTERVAL_MS) {
      return 'rate_limited';
    }
  }
  sh.appendRow([new Date(), name, text]);      // fixed width, always three cells
  return 'ok';
}

function handleSuggest_(body) {
  // Reject UNEXPECTED KEYS rather than ignoring them, so a field added later
  // cannot arrive unnoticed. This also rejects a payload carrying `updates`.
  for (var k in body) {
    if (!SUGGEST_ALLOWED_KEYS[k]) return json_({ ok: false, error: 'unexpected_key' });
  }
  if (typeof body.name !== 'string' || typeof body.text !== 'string') {
    return json_({ ok: false, error: 'bad_type' });
  }
  var name = body.name.trim();
  var text = body.text.trim();
  if (!SUGGEST_NAME_RE.test(name)) return json_({ ok: false, error: 'bad_name' });
  if (text.length < 1 || text.length > SUGGEST_TEXT_MAX) return json_({ ok: false, error: 'bad_text' });

  var lock = LockService.getScriptLock();        // the SAME script lock the tick path uses
  try { lock.waitLock(10000); }
  catch (err) { return json_({ ok: false, error: 'busy' }); }
  try {
    var r = addSuggestion_(name, text);
    if (r === 'no_tab')       return json_({ ok: false, error: 'no_suggestions_tab' });
    if (r === 'rate_limited') return json_({ ok: false, error: 'rate_limited' });
    return json_({ ok: true, suggestions: readSuggestions_() });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.what === 'suggestions') {
    return json_({ ok: true, suggestions: readSuggestions_() });
  }
  return json_({ ok: true, tasks: readState_() });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  if (body.token !== TOKEN) return json_({ ok: false, error: 'bad_token' });

  if (body.action === 'suggest') return handleSuggest_(body);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var updates = body.updates || [];
    for (var i = 0; i < updates.length; i++) {
      writeOne_(updates[i].id, updates[i].done, updates[i].who);
    }
    return json_({ ok: true, tasks: readState_() });
  } finally {
    lock.releaseLock();
  }
}
