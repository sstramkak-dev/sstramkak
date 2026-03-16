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
 *   Expire Date,   ← expire date entered by user in the TopUp form (mapped from endDate)
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
// TopUp sheet configuration
// ---------------------------------------------------------------------------

/**
 * Canonical column order for the TopUp sheet.
 * These headers are written on first use and any missing ones are appended.
 */
var TOPUP_HEADERS = [
  'id', 'customerId', 'name', 'phone', 'amount',
  'agent', 'branch', 'date', 'Expire Date',
  'tariff', 'remark', 'tuStatus', 'lat', 'lng'
];

/**
 * Maps incoming record field names to their canonical column header names for
 * the TopUp sheet.  Fields not listed here are written under their own key.
 */
var TOPUP_FIELD_MAP = {
  endDate:    'Expire Date',
  expireDate: 'Expire Date'
};

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
 * - For the TopUp sheet, fields are remapped via TOPUP_FIELD_MAP and date
 *   values are normalised to YYYY-MM-DD before writing.
 */
function handleSync(sheetName, data) {
  if (!Array.isArray(data) || data.length === 0) {
    // Nothing to write — leave the sheet as-is but confirm success.
    return jsonResponse({ status: 'ok', rows: 0 });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, sheetName);

  // For the TopUp sheet, remap field keys and normalise date values.
  var normalised = (sheetName === 'TopUp')
    ? data.map(normaliseTopUpRecord)
    : data;

  // Seed canonical headers for TopUp before deriving from the data keys so
  // that the column order is stable and "Expire Date" is always present.
  if (sheetName === 'TopUp') {
    ensureTopUpHeaders(sheet);
  }

  // Derive the ordered list of columns from the existing header + any new keys
  // found in the incoming data.
  var headers = ensureHeaders(sheet, normalised);

  // Clear data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }

  // Write each record as a row
  var rows = normalised.map(function(record) {
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

// ---------------------------------------------------------------------------
// TopUp-specific helpers
// ---------------------------------------------------------------------------

/**
 * Write the canonical TopUp headers to the sheet if no header row exists yet,
 * or append any missing canonical headers to the right of existing ones.
 */
function ensureTopUpHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  var existing = [];

  if (lastCol > 0 && sheet.getLastRow() > 0) {
    existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  }

  var hasHeaders = existing.some(function(h) { return h.trim() !== ''; });

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, TOPUP_HEADERS.length).setValues([TOPUP_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  var existingSet = {};
  existing.forEach(function(h) { if (h) existingSet[h] = true; });

  var toAdd = TOPUP_HEADERS.filter(function(h) { return !existingSet[h]; });
  if (toAdd.length > 0) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  }
}

/**
 * Remap a raw TopUp record so that:
 *  - `endDate` / `expireDate` → `Expire Date` (the canonical column name)
 *  - Date values stored under `Expire Date` and `date` are normalised to YYYY-MM-DD.
 *  - The original source field is removed to avoid creating a duplicate column.
 */
function normaliseTopUpRecord(record) {
  var out = {};

  Object.keys(record).forEach(function(key) {
    var canonical = TOPUP_FIELD_MAP[key] || key;
    var value     = record[key];

    // Normalise both the expire-date column and the regular date column to YYYY-MM-DD.
    if (canonical === 'Expire Date' || canonical === 'date') {
      value = toYmd(value);
    }

    // If two source keys map to the same canonical column, prefer a non-empty
    // value: keep the existing value if it is already set and non-empty.
    if (canonical in out && (out[canonical] === '' || out[canonical] === null || out[canonical] === undefined)) {
      out[canonical] = value;
    } else if (!(canonical in out)) {
      out[canonical] = value;
    }
  });

  return out;
}

/**
 * Normalise a value to a YYYY-MM-DD string.
 * Returns an empty string if the value is absent or not parseable.
 */
function toYmd(v) {
  if (v === null || v === undefined || v === '') return '';

  // Google Apps Script Date object
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var s = String(v).trim();
  if (!s) return '';

  // Already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // Try generic Date parsing as a fallback
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return '';
}

/**
 * Wrap a plain JS object as a JSON ContentService response.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
