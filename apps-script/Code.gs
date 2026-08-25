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

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
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
