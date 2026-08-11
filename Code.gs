/**
 * Roundtable live board — backend (Google Apps Script, bound to a Google Sheet)
 *
 * ARCHITECTURE NOTE:
 * Pages are served via HtmlService — that's the only way Apps Script can
 * return a page browsers actually render (ContentService can only serve raw
 * data: JSON, plain text, CSV — not renderable HTML). The client talks to the
 * backend via ordinary fetch() calls to the JSON endpoints below, rather than
 * google.script.run, since that RPC bridge is the more likely source of the
 * Safari-specific loading failure — fetch() is just a normal network request.
 *
 * Data lives in a sheet tab called "Entries", created automatically on first run.
 * Columns: ID, Timestamp, Category, Organisation, Text, Votes
 */

var SHEET_NAME = 'Entries';
var CATEGORIES = ['challenge', 'experience', 'message'];
var HEADER = ['ID', 'Timestamp', 'Category', 'Organisation', 'Text', 'Votes'];

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'list')   return jsonOut_(serverListEntries());
  if (action === 'add')    { serverAddEntry(e.parameter.category, e.parameter.org, e.parameter.text); return jsonOut_({ ok: true }); }
  if (action === 'edit')   { serverEditEntry(e.parameter.id, e.parameter.category, e.parameter.org, e.parameter.text); return jsonOut_({ ok: true }); }
  if (action === 'delete') { serverDeleteEntry(e.parameter.id); return jsonOut_({ ok: true }); }
  if (action === 'vote')   { serverVoteEntry(e.parameter.id); return jsonOut_({ ok: true }); }

  // No action -> serve an actual rendered page via HtmlService.
  var page = (e && e.parameter && e.parameter.page) || 'board';
  var fileName = page === 'input' ? 'Input' : 'Board';
  return HtmlService.createHtmlOutputFromFile(fileName)
    .setTitle(page === 'input' ? 'Roundtable — Add Answer' : 'Roundtable — Live Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER);
    return sheet;
  }
  // Self-heal older sheets created before the Organisation column existed.
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (firstRow.indexOf('Organisation') === -1) {
    if (sheet.getLastRow() < 1) {
      sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    } else {
      sheet.insertColumnBefore(4); // shifts old Text/Votes right, opens column D
      sheet.getRange(1, 4).setValue('Organisation');
    }
  }
  return sheet;
}

function serverListEntries() {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  values.shift(); // drop header row
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    out.push({
      id: row[0],
      category: row[2],
      org: row[3] || '',
      text: row[4],
      votes: Number(row[5]) || 0
    });
  }
  return out;
}

function serverAddEntry(category, org, text) {
  text = (text || '').trim();
  if (!text) return;
  category = (category || '').toLowerCase();
  if (CATEGORIES.indexOf(category) === -1) category = 'challenge';
  org = (org || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var sheet = getSheet_();
    var id = Utilities.getUuid();
    sheet.appendRow([id, new Date(), category, org, text, 0]);
  } finally {
    lock.releaseLock();
  }
}

function serverEditEntry(id, category, org, text) {
  text = (text || '').trim();
  if (!text) return;
  category = (category || '').toLowerCase();
  if (CATEGORIES.indexOf(category) === -1) category = 'challenge';
  org = (org || '').trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 3, 1, 3).setValues([[category, org, text]]); // C:category D:org E:text
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function serverDeleteEntry(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function serverVoteEntry(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var sheet = getSheet_();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        sheet.getRange(i + 1, 6).setValue((Number(values[i][5]) || 0) + 1); // column F = Votes
        break;
      }
    }
  } finally {
    lock.releaseLock();
  }
}
