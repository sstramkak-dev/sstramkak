/**
 * Smart 5G Dashboard — Google Apps Script Web App
 *
 * Handles POST requests from the web application to sync data with Google Sheets.
 *
 * Supported actions (sent as JSON in the POST body):
 *   { sheet: 'SheetName', action: 'sync',   data: [ {...}, ... ] }  — full sheet upsert
 *   { sheet: 'SheetName', action: 'read'                          }  — return all rows as JSON
 *   { sheet: 'SheetName', action: 'delete', data: { id: '...' }   }  — delete row by id
 *
 * Sheets managed:
 *   Sales, Customers, TopUp, Terminations, OutCoverage,
 *   Promotions, Deposits, KPI, Items, Coverage, Staff
 *
 * TopUp sheet columns (auto-created if missing):
 *   id, customerId, name, phone, amount, agent, branch, date,
 *   endDate,   ← expire date entered by user in the TopUp form
 *   tariff, remark, tuStatus, lat, lng
 *
 * Deployment:
 *   1. Open the Google Apps Script project linked to your spreadsheet.
 *   2. Replace (or paste) this code into Code.gs.
 *   3. Click "Deploy" → "Manage deployments" → "New deployment"
 *      (or update an existing one) as a Web App:
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   4. Copy the new Web App URL and update GS_URL in app.js if it changed.
 *   5. Save — no other configuration is required; columns are created on first sync.
 */

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var sheetName = payload.sheet;
    var action    = payload.action;
    var data      = payload.data;

    if (!sheetName) return jsonResponse({ error: 'Missing sheet name' });

    if (action === 'sync')   return handleSync(sheetName, data);
    if (action === 'read')   return handleRead(sheetName);
    if (action === 'delete') return handleDelete(sheetName, data);

    return jsonResponse({ error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Allow GET for quick health-check / CORS preflight
function doGet() {
  return ContentService.createTextOutput('OK');
}

// ---------------------------------------------------------------------------
// Action handlers
// ---------------------------------------------------------------------------

/**
 * Full upsert: replace all rows in the sheet with the supplied data array.
 * - Merges any new field keys into the header row (adds columns as needed).
 * - Existing rows without a new column get an empty string in that column.
 * - Backward-compatible: sheets that already have all columns are unaffected.
 */
function handleSync(sheetName, data) {
  if (!Array.isArray(data) || data.length === 0) {
    // Nothing to write — leave the sheet as-is but confirm success.
    return jsonResponse({ status: 'ok', rows: 0 });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, sheetName);

  // Derive the ordered list of columns from the existing header + any new keys
  // found in the incoming data.
  var headers = ensureHeaders(sheet, data);

  // Clear data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }

  // Write each record as a row
  var rows = data.map(function(record) {
    return headers.map(function(col) {
      var val = record[col];
      return (val === undefined || val === null) ? '' : val;
    });
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return jsonResponse({ status: 'ok', rows: rows.length });
}

/**
 * Read all rows and return them as an array of objects keyed by header name.
 */
function handleRead(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ data: [] });

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return jsonResponse({ data: [] });

  var values  = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(String);
  var result  = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    // Skip entirely empty rows
    if (row.every(function(c) { return c === '' || c === null || c === undefined; })) continue;
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i] !== undefined ? row[i] : '';
    });
    result.push(obj);
  }

  return jsonResponse({ data: result });
}

/**
 * Delete the row whose 'id' column matches data.id.
 */
function handleDelete(sheetName, data) {
  if (!data || !data.id) return jsonResponse({ error: 'Missing id' });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ status: 'ok', deleted: 0 });

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return jsonResponse({ status: 'ok', deleted: 0 });

  var values  = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(String);
  var idCol   = headers.indexOf('id');
  if (idCol < 0) return jsonResponse({ status: 'ok', deleted: 0 });

  var deleted = 0;
  // Iterate from the bottom so row indices stay valid as rows are removed
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][idCol]) === String(data.id)) {
      sheet.deleteRow(r + 1); // +1 because sheet rows are 1-indexed
      deleted++;
    }
  }

  return jsonResponse({ status: 'ok', deleted: deleted });
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Return (or create) a sheet with the given name in the active spreadsheet.
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Ensure the sheet has a header row that contains every key present in the
 * data array.  New keys are appended to the right of existing headers.
 *
 * Returns the final, ordered list of header strings.
 */
function ensureHeaders(sheet, data) {
  // Collect all unique keys from the incoming data, preserving first-seen order
  var keySeen = {};
  var allKeys = [];
  data.forEach(function(record) {
    Object.keys(record).forEach(function(k) {
      if (!keySeen[k]) {
        keySeen[k] = true;
        allKeys.push(k);
      }
    });
  });

  var lastCol = sheet.getLastColumn();
  var existingHeaders = [];

  if (lastCol > 0 && sheet.getLastRow() > 0) {
    existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  }

  // Find keys that are not yet in the header row
  var newKeys = allKeys.filter(function(k) {
    return existingHeaders.indexOf(k) < 0;
  });

  // Append missing header cells
  if (newKeys.length > 0) {
    var startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, newKeys.length).setValues([newKeys]);
    existingHeaders = existingHeaders.concat(newKeys);
  }

  return existingHeaders;
}

/**
 * Wrap a plain JS object as a JSON ContentService response.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
