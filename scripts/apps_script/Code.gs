/**
 * Paste this into Extensions > Apps Script for the target Google Sheet,
 * then Deploy > New deployment > type "Web app":
 *   - Execute as: Me
 *   - Who has access: Anyone
 * Copy the resulting Web App URL into the APPS_SCRIPT_URL GitHub secret.
 *
 * Set TOKEN below to a random string of your choosing, and put the same
 * value in the APPS_SCRIPT_TOKEN GitHub secret — this is the only thing
 * stopping a stranger who finds the URL from writing junk into your sheet.
 */

var TOKEN = 'REPLACE_WITH_A_RANDOM_SECRET';
var SHEET_TAB = 'raw_credit_benchmark';
var HEADERS = ['reference_month', 'product', 'balance_brl_bn', 'nominal_rate_pa', 'npl90', 'updated_at'];

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  if (body.token !== TOKEN) {
    return jsonResponse({ ok: false, error: 'invalid token' });
  }

  var rows = body.rows || [];
  var sheet = getOrCreateSheet();
  var now = new Date().toISOString();

  var data = sheet.getDataRange().getValues();
  var keyToRow = {}; // "reference_month|product" -> 1-based row number
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0] + '|' + data[i][1];
    keyToRow[key] = i + 1;
  }

  var updated = 0;
  var appended = 0;
  var toAppend = [];

  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    var key = r.reference_month + '|' + r.product;
    var values = [r.reference_month, r.product, r.balance_brl_bn, r.nominal_rate_pa, r.npl90, now];

    if (keyToRow[key]) {
      sheet.getRange(keyToRow[key], 1, 1, HEADERS.length).setValues([values]);
      updated++;
    } else {
      toAppend.push(values);
      appended++;
    }
  }

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, HEADERS.length).setValues(toAppend);
  }

  return jsonResponse({ ok: true, updated: updated, appended: appended });
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TAB);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
