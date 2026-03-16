/**
 * Smart 5G Dashboard — Google Apps Script Web App
 *
 * Handles POST requests from the web application to sync data with Google Sheets.
 *
 * Supported actions (sent as JSON in the POST body):
 *   { sheet: 'SheetName', action: 'sync',   data: [ {...}, ... ] }  — full sheet upsert
 *   { sheet: 'SheetName', action: 'read'                          }  — return all rows as JSON
 *   { sheet: 'SheetName', action: 'delete', data: { id: '...' }   }  — delete row by id
 *   { sheet: 'DailySale', action: 'append', data: { ... }         }  — append one Daily Sale row
 *
 * Sheets managed:
 *   Sales, Customers, TopUp, Terminations, OutCoverage,
 *   Promotions, Deposits, KPI, Items, Coverage, Staff, DailySale
 *
 * DailySale sheet columns (auto-created / guaranteed, including Expire Date):
 *   Date, Expire Date, Agent, Branch,
 *   Buy Number, ChangeSIM, Recharge, SC Dealer, Device + Accessories, Total Revenue,
 *   Gross Ads, Smart@Home, Smart Fiber+, SmartNas, Monthly Upsell,
 *   Note, Submitted At, ID
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
// Configuration
// ---------------------------------------------------------------------------

/** Name of the Daily Sale sheet (change here if your sheet tab has a different name). */
var DAILY_SALE_SHEET_NAME = 'DailySale';

/**
 * Fixed column order for the DailySale sheet.
 * 'Expire Date' is always guaranteed — it will be blank when not supplied.
 * Any extra headers already present in the sheet are preserved to the right.
 */
var DAILY_SALE_HEADERS = [
  'Date', 'Expire Date', 'Agent', 'Branch',
  'Buy Number', 'ChangeSIM', 'Recharge', 'SC Dealer', 'Device + Accessories', 'Total Revenue',
  'Gross Ads', 'Smart@Home', 'Smart Fiber+', 'SmartNas', 'Monthly Upsell',
  'Note', 'Submitted At', 'ID'
];

/** Maps payload dollarItems object keys → DailySale column names. */
var DOLLAR_ITEM_MAP = {
  'i9':  'Buy Number',
  'i6':  'ChangeSIM',
  'i7':  'Recharge',
  'i10': 'SC Dealer',
  'i11': 'Device + Accessories',
  'i8':  'Total Revenue'
};

/** Maps payload items (unit group) object keys → DailySale column names. */
var UNIT_ITEM_MAP = {
  'i1': 'Gross Ads',
  'i2': 'Smart@Home',
  'i3': 'Smart Fiber+',
  'i4': 'SmartNas',
  'i5': 'Monthly Upsell'
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
    if (action === 'append') return handleAppend(sheetName, data);

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

/**
 * Append a single row to the specified sheet.
 *
 * For the DailySale sheet this function:
 *   - Guarantees the fixed DAILY_SALE_HEADERS exist (including 'Expire Date').
 *   - Flattens nested items / dollarItems objects into individual columns.
 *   - Normalises Date and Expire Date to YYYY-MM-DD strings.
 *   - Leaves 'Expire Date' blank when not provided in the payload.
 *
 * For any other sheet the function ensures headers from the data keys and
 * appends the row as-is.
 *
 * Expected payload:
 *   {
 *     sheet:       'DailySale',
 *     action:      'append',
 *     data: {
 *       id:          '<record id>',
 *       agent:       '<name>',
 *       branch:      '<branch>',
 *       date:        'YYYY-MM-DD',
 *       expireDate:  'YYYY-MM-DD',   // optional — blank if omitted
 *       note:        '<remark>',
 *       submittedAt: '<ISO timestamp>',
 *       items:       { i1: 3, i2: 1, ... },      // unit group (object or JSON string)
 *       dollarItems: { i9: 50, i7: 100, ... }     // dollar group (object or JSON string)
 *     }
 *   }
 */
function handleAppend(sheetName, data) {
  if (!data || typeof data !== 'object') {
    return jsonResponse({ error: 'Missing or invalid data' });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, sheetName);

  var isDaily = (sheetName === DAILY_SALE_SHEET_NAME);
  var headers = isDaily ? ensureDailySaleHeaders(sheet) : ensureHeaders(sheet, [data]);
  var flat    = isDaily ? flattenDailySaleRecord(data)  : data;

  var row = headers.map(function(col) {
    var val = flat[col];
    return (val === undefined || val === null) ? '' : val;
  });

  sheet.appendRow(row);

  return jsonResponse({ status: 'ok', columns: row.length });
}

// ---------------------------------------------------------------------------
// DailySale-specific helpers
// ---------------------------------------------------------------------------

/**
 * Ensure the DailySale sheet has every column listed in DAILY_SALE_HEADERS.
 * Any columns already in the sheet are preserved; missing ones are appended.
 * Returns the final ordered list of header strings.
 */
function ensureDailySaleHeaders(sheet) {
  var lastCol  = sheet.getLastColumn();
  var existing = [];

  if (lastCol > 0 && sheet.getLastRow() > 0) {
    existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  }

  var missing = DAILY_SALE_HEADERS.filter(function(h) {
    return existing.indexOf(h) < 0;
  });

  if (missing.length > 0) {
    var startCol = existing.length + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    existing = existing.concat(missing);
  }

  return existing;
}

/**
 * Flatten a Daily Sale record (with nested items / dollarItems as objects or
 * JSON strings) into a plain object keyed by DAILY_SALE_HEADERS column names.
 *
 * Rules:
 *   - 'Date'        → formatDateStr(data.date)
 *   - 'Expire Date' → formatDateStr(data.expireDate); blank string if absent
 *   - Dollar/unit item columns → numeric value (0 when not present in payload)
 */
function flattenDailySaleRecord(data) {
  var flat = {};

  flat['Date']         = formatDateStr(data.date        || '');
  flat['Expire Date']  = formatDateStr(data.expireDate  || '');  // blank when omitted
  flat['Agent']        = data.agent       || '';
  flat['Branch']       = data.branch      || '';
  flat['Note']         = data.note        || data.remark || '';
  flat['Submitted At'] = data.submittedAt || '';
  flat['ID']           = data.id          || '';

  var dollarItems = parseJsonOrObj(data.dollarItems);
  Object.keys(DOLLAR_ITEM_MAP).forEach(function(key) {
    var colName = DOLLAR_ITEM_MAP[key];
    var raw = dollarItems[key];
    flat[colName] = (raw !== undefined && raw !== null) ? (parseFloat(raw) || 0) : 0;
  });

  var unitItems = parseJsonOrObj(data.items);
  Object.keys(UNIT_ITEM_MAP).forEach(function(key) {
    var colName = UNIT_ITEM_MAP[key];
    var raw = unitItems[key];
    flat[colName] = (raw !== undefined && raw !== null) ? (parseInt(raw, 10) || 0) : 0;
  });

  return flat;
}

/**
 * Normalise a date value to a YYYY-MM-DD string.
 * Returns an empty string when the value is absent or cannot be parsed.
 */
function formatDateStr(val) {
  if (!val) return '';
  if (typeof val === 'string') {
    // Already YYYY-MM-DD — return as-is.
    if (/^\d{4}-\d{2}-\d{2}$/.test(val)) return val;
    // ISO timestamp — take only the date part.
    if (/^\d{4}-\d{2}-\d{2}T/.test(val)) return val.split('T')[0];
  }
  try {
    var d = new Date(val);
    if (isNaN(d.getTime())) return '';
    var y   = d.getFullYear();
    var m   = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
  } catch (e) {
    return '';
  }
}

/**
 * Parse a value that is either already an object or a JSON-encoded string.
 * Returns an empty object on failure or when the value is absent.
 */
function parseJsonOrObj(val) {
  if (!val) return {};
  if (typeof val === 'object') return val;
  try { return JSON.parse(val); } catch (e) { return {}; }
}

// ---------------------------------------------------------------------------
// Generic sheet helpers
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
