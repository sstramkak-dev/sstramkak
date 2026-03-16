// ============================================================
// Smart 5G Dashboard — app.js
// ============================================================

// ------------------------------------------------------------
// State & Constants
// ------------------------------------------------------------
let currentRole = 'admin'; // 'admin', 'supervisor', 'user'
let currentUser = null;
let currentPage = 'dashboard';
let currentSaleTab = 'report';
let currentCustomerTab = 'new-customer';
let currentSettingsTab = 'permission';
let currentCoverageTab = 'smart-home';
let currentPromoView = 'new'; // 'new' or 'expired'
let currentReportView = 'table'; // 'table' or 'summary'
let filteredSales = [];
let itemGroupSelected = 'unit'; // 'unit' or 'dollar'
let kpiValueMode = 'unit'; // 'unit' or 'currency'
let kpiTypeSelected = 'Sales';
let kpiForSelected = 'shop'; // 'shop' or 'agent'
let kpiSelectedMonth = ''; // '' means no filter (show all)

// Chart instances
let _cTrend = null, _cMix = null, _cAgent = null, _cGrowth = null;
let _cSaleMix = null, _cSaleAgent = null;
let _cDepositPerf = null;
let _cKpiGauges = [];

// Coverage map instances
var _covMapSmartHome = null, _covMapSmartHome5G = null, _covMapSmartFiber = null;
var _covPickerMap = null, _covPickerMarker = null, _covPickerHighlight = null;
var _tuPickerMap = null, _tuPickerMarker = null;

// Commune autocomplete state
var _covCommuneDebounce = null;

// Approval form state
var _approvalFormData = null;
var _sigCanvas = null, _sigCtx = null, _sigDrawing = false;

// Constants
const TAB_PERM = { admin: ['permission'], cluster: ['permission'], supervisor: [], agent: [], user: [] };
const TAB_LBL = { permission: 'Permission' };
const AV_COLORS = ['#E53935','#8E24AA','#1565C0','#00838F','#2E7D32','#F57F17','#4E342E','#37474F'];
const CHART_PAL = ['#1B7D3D','#2196F3','#FF9800','#9C27B0','#F44336','#00BCD4','#FFEB3B','#795548'];
const KNOWN_CURS = ['USD','KHR','THB','VND'];
const KNOWN_UNITS = ['Unit','SIM','GB','MB','Minutes','SMS','Voucher'];

// Item ID constants for key KPI calculations
const ITEM_ID_REVENUE = 'i8';
const ITEM_ID_RECHARGE = 'i7';
const ITEM_ID_GROSS_ADS = 'i1';
const ITEM_ID_SMART_HOME = 'i2';
const ITEM_ID_SMART_FIBER = 'i3';
const ITEM_ID_BUY_NUMBER = 'i9';

// Fixed sale form items (Dollar Group and Unit Group)
const DOLLAR_SALE_ITEMS = [
  { id: 'i9',  name: 'Buy Number' },
  { id: 'i6',  name: 'ChangeSIM' },
  { id: 'i7',  name: 'Recharge' },
  { id: 'i10', name: 'SC Dealer' },
  { id: 'i11', name: 'Device + Accessories' },
];
const UNIT_SALE_ITEMS = [
  { id: 'i1', name: 'Gross Ads' },
  { id: 'i2', name: 'Smart@Home' },
  { id: 'i3', name: 'Smart Fiber+' },
  { id: 'i4', name: 'SmartNas' },
  { id: 'i5', name: 'Monthly Upsell' },
];

const BRANCHES = ['Phnom Penh', 'Siem Reap', 'Battambang', 'Sihanoukville', 'Kampong Cham', 'Express_Tramkak'];

const SUPPORT_CONTACT = { email: 'support@smart5g.com', phone: '+855 23 123 456' };

// ── Google Sheets Sync ──────────────────────────────────────
const GS_URL = 'https://script.google.com/macros/s/AKfycbxokMtXAEvhyUrtfCSFCDmmdv6Cr6rOFVxkBxtH_eUbQc4okwCcVNVVvOv02nmanfPdTA/exec';

function _gsPost(payload, retries) {
  if (!GS_URL) return Promise.resolve();
  retries = retries === undefined ? 2 : retries;
  return fetch(GS_URL, {
    method: 'POST',
    mode: 'no-cors',
    body: JSON.stringify(payload)
  }).catch(function(err) {
    if (retries > 0) {
      return new Promise(function(resolve) { setTimeout(resolve, 1500); })
        .then(function() { return _gsPost(payload, retries - 1); });
    }
    console.warn('GS post failed after retries:', err);
    throw err;
  });
}

function normalizeStaffRecord(u) {
  var out = {};
  Object.keys(u).forEach(function(k) {
    var val = u[k];
    // Normalize key to lowercase so that column headers from Google Sheets
    // (e.g. 'Username', 'Password') are accessible via the expected lowercase keys.
    var key = k.toLowerCase();
    // Trim all string values and handle null/undefined
    if (typeof val === 'string') {
      out[key] = val.trim();
    } else {
      out[key] = val;
    }
  });
  // Lowercase status so login comparison works regardless of how the sheet stores it
  // (role VALUE is kept as-is because roleMap lookups depend on the original casing, e.g. 'Admin')
  if (out.status) out.status = out.status.toLowerCase();
  // Ensure username exists and is not empty
  if (!out.username || out.username === '') {
    console.warn('Staff record missing username:', out);
  }
  return out;
}

function isAdminUser(u) {
  return (u.username || '').toLowerCase() === 'admin' && (u.role || '').toLowerCase() === 'admin';
}

function fetchStaffFromSheet() {
  if (!GS_URL) {
    console.warn('[SYNC] No Google Sheets URL configured');
    return Promise.resolve();
  }

  console.log('[SYNC] Fetching staff from Google Sheets...');

  return readSheet('Staff')
    .then(function(data) {
      console.log('[SYNC] Raw staff data received:', data.length, 'records');

      if (!Array.isArray(data) || data.length === 0) {
        console.warn('[SYNC] No staff data returned from sheet');
        return;
      }

      console.log('[SYNC] Processing', data.length, 'staff records');

      // Normalize staff records from Google Sheets (trim whitespace, lowercase status)
      var normalized = data.map(normalizeStaffRecord);
      console.log('[SYNC] Normalized staff records:', normalized.map(function(u) { return { username: u.username, status: u.status, role: u.role, hasPassword: !!u.password }; }));

      // Keep local admin as fallback if sheet doesn't contain one
      if (!normalized.some(isAdminUser)) {
        var localAdmin = staffList.find(isAdminUser);
        if (localAdmin) {
          console.log('[SYNC] Adding local admin as fallback');
          normalized.unshift(localAdmin);
        }
      }
      // Keep default agent users as fallback if sheet doesn't contain them
      DEFAULT_AGENT_USERS.forEach(function(du) {
        if (!normalized.some(function(u) { return u.username === du.username; })) {
          console.log('[SYNC] Adding default user:', du.username);
          normalized.push(du);
        }
      });
      staffList = normalized;
      lsSave(LS_KEYS.staff, staffList);
      console.log('[SYNC] ✓ Staff list updated:', staffList.length, 'total users');
    })
    .catch(function(e) {
      console.error('[SYNC] ✗ Staff fetch error:', e);
      var errEl = g('login-error');
      if (errEl) {
        var hasCachedStaff = localStorage.getItem(LS_KEYS.staff) !== null;
        errEl.textContent = hasCachedStaff
          ? 'Could not reach the server. Signing in with cached data.'
          : 'Could not reach the server. Please check your internet connection and try again.';
        errEl.style.display = '';
        setTimeout(function() { if (errEl) errEl.style.display = 'none'; }, 5000);
      }
    });
}

// Serialize all object/array field values to JSON strings so rows can be stored
// as flat strings in Google Sheets. Used by both syncSheet and syncUpAll.
function normalizeRowForSheet(row) {
  var out = {};
  Object.keys(row).forEach(function(k) {
    var v = row[k];
    out[k] = (v !== null && v !== undefined && typeof v === 'object') ? JSON.stringify(v) : (v !== undefined ? v : '');
  });
  return out;
}

function normalizeArrayForSheet(dataArray) {
  return Array.isArray(dataArray) ? dataArray.map(normalizeRowForSheet) : [];
}

function syncSheet(sheetName, dataArray) {
  if (!GS_URL) return;
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Syncing\u2026';

  _gsPost({ sheet: sheetName, action: 'sync', data: normalizeArrayForSheet(dataArray) })
    .then(function() {
      if (ind) ind.className = '';
      if (lbl) lbl.textContent = 'Synced \u2713';
      setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    })
    .catch(function(err) {
      console.warn('GS sync error:', err);
      if (ind) ind.className = 'error';
      if (lbl) lbl.textContent = 'Sync failed';
    });
}

function deleteFromSheet(sheetName, id) {
  if (!GS_URL) return;
  _gsPost({ sheet: sheetName, action: 'delete', data: { id: id } })
    .catch(function(err) { console.warn('GS delete error:', err); });
}

// Read a single sheet from Google Apps Script (returns a Promise<array>).
// Sends { sheet, action:'read' } via POST and accepts both a plain array
// and a { status:'ok', data:[...] } envelope in the response.
function readSheet(sheetName) {
  if (!GS_URL) return Promise.resolve([]);

  console.log('[SYNC] Reading sheet:', sheetName);

  return fetch(GS_URL, {
    method: 'POST',
    body: JSON.stringify({ sheet: sheetName, action: 'read' })
  })
    .then(function(r) {
      if (!r.ok) {
        console.error('[SYNC] HTTP error:', r.status);
        throw new Error('HTTP ' + r.status);
      }
      return r.json();
    })
    .then(function(resp) {
      console.log('[SYNC] Sheet response for', sheetName + ':', resp);

      if (Array.isArray(resp)) {
        console.log('[SYNC] Received array with', resp.length, 'items');
        return resp;
      }
      if (resp && Array.isArray(resp.data)) {
        console.log('[SYNC] Received data envelope with', resp.data.length, 'items');
        return resp.data;
      }
      console.warn('[SYNC] Unexpected response format');
      return [];
    })
    .catch(function(e) {
      console.error('[SYNC] readSheet(' + sheetName + ') failed:', e);
      return [];
    });
}

// Push all local data to Google Sheets. Exposed globally so admin can call it
// from the browser console or a "Sync Up" button.
function syncUpAll() {
  if (!GS_URL) {
    showToast('No sync URL configured.', 'error');
    return;
  }
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  var upBtn = document.getElementById('sync-up-btn');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Uploading\u2026';
  if (upBtn) upBtn.disabled = true;

  var sheets = [
    { name: 'Sales',        data: function() { return saleRecords; } },
    { name: 'Customers',    data: function() { return newCustomers; } },
    { name: 'TopUp',        data: function() { return topUpList; } },
    { name: 'Terminations', data: function() { return terminationList; } },
    { name: 'OutCoverage',  data: function() { return outCoverageList; } },
    { name: 'Promotions',   data: function() { return promotionList; } },
    { name: 'Deposits',     data: function() { return depositList; } },
    { name: 'KPI',          data: function() { return kpiList; } },
    { name: 'Items',        data: function() { return itemCatalogue; } },
    { name: 'Coverage',     data: function() { return coverageLocations; } },
    { name: 'Staff',        data: function() { return staffList; } },
  ];

  var promises = sheets.map(function(s) {
    return _gsPost({ sheet: s.name, action: 'sync', data: normalizeArrayForSheet(s.data()) });
  });

  Promise.all(promises).then(function() {
    if (ind) ind.className = '';
    if (lbl) lbl.textContent = 'Uploaded \u2713';
    if (upBtn) upBtn.disabled = false;
    setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    showToast('All data uploaded to Google Sheets.', 'success');
  }).catch(function(err) {
    console.warn('Sync up error:', err);
    if (ind) ind.className = 'error';
    if (lbl) lbl.textContent = 'Upload failed';
    if (upBtn) upBtn.disabled = false;
    showToast('Upload failed. Please try again.', 'error');
  });
}

// Sync all sheets down from Google Sheets into local memory and localStorage,
// then re-render the current page. Exposed globally so admin can call it from
// the browser console or a "Sync Down" button.
function syncDownAll() {
  if (!GS_URL) {
    showToast('No sync URL configured.', 'error');
    return Promise.resolve();
  }
  var ind = document.getElementById('gs-sync-indicator');
  var lbl = document.getElementById('gs-sync-status');
  var btn = document.getElementById('sync-down-btn');
  if (ind) ind.className = 'syncing';
  if (lbl) lbl.textContent = 'Syncing\u2026';
  if (btn) btn.disabled = true;

  var sheets = [
    { name: 'Sales',        lsKey: LS_KEYS.sales,        assign: function(d) { saleRecords = d; } },
    { name: 'Customers',    lsKey: LS_KEYS.customers,    assign: function(d) { newCustomers = d; } },
    { name: 'TopUp',        lsKey: LS_KEYS.topup,        assign: function(d) { topUpList = d; } },
    { name: 'Terminations', lsKey: LS_KEYS.terminations, assign: function(d) { terminationList = d; } },
    { name: 'OutCoverage',  lsKey: LS_KEYS.outCoverage,  assign: function(d) { outCoverageList = d; } },
    { name: 'Promotions',   lsKey: LS_KEYS.promotions,   assign: function(d) { promotionList = d; } },
    { name: 'Deposits',     lsKey: LS_KEYS.deposits,     assign: function(d) { depositList = d; } },
    { name: 'KPI',          lsKey: LS_KEYS.kpis,         assign: function(d) { kpiList = d; } },
    { name: 'Items',        lsKey: LS_KEYS.items,        assign: function(d) { if (d.length) itemCatalogue = d; } },
    { name: 'Coverage',     lsKey: LS_KEYS.coverage,     assign: function(d) { coverageLocations = d; } },
    { name: 'Staff',        lsKey: LS_KEYS.staff,        assign: function(d) {
      var normalizedStaff = d.map(normalizeStaffRecord);
      if (!normalizedStaff.some(isAdminUser)) {
        var localAdmin = staffList.find(isAdminUser);
        if (localAdmin) normalizedStaff.unshift(localAdmin);
      }
      staffList = normalizedStaff;
    }}
  ];

  var promises = sheets.map(function(s) {
    return readSheet(s.name)
      .then(function(data) {
        if (!Array.isArray(data) || data.length === 0) return;
        // Parse JSON-serialized object fields back to objects (e.g. items, dollarItems, cashDetail)
        var parsed = data.map(function(row) {
          var out = {};
          Object.keys(row).forEach(function(k) {
            var v = row[k];
            if (typeof v === 'string' && v.length > 0 && (v[0] === '{' || v[0] === '[')) {
              try { out[k] = JSON.parse(v); } catch(parseErr) { out[k] = v; }
            } else {
              out[k] = v;
            }
          });
          return out;
        });
        s.assign(parsed);
        lsSave(s.lsKey, parsed);
      });
  });

  return Promise.all(promises).then(function() {
    // Re-sync currentUser from the refreshed staffList
    if (currentUser) {
      var freshUser = staffList.find(function(u) { return u.id === currentUser.id; });
      if (freshUser) {
        currentUser = freshUser;
        var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = freshUser.name;
        var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(freshUser.name);
      } else if (!isAdminUser(currentUser)) {
        currentUser = null;
        var as2 = g('app-shell'); if (as2) as2.style.display = 'none';
        var ls2 = g('login-screen'); if (ls2) ls2.style.display = 'flex';
        var lf2 = g('login-form'); if (lf2) lf2.reset();
        showToast('Your account was not found. You have been signed out.', 'error');
        return;
      }
    }
    if (ind) ind.className = '';
    if (lbl) lbl.textContent = 'Synced \u2713';
    if (btn) btn.disabled = false;
    setTimeout(function() { if (lbl) lbl.textContent = ''; }, 3000);
    // Re-render current page with fresh data
    if (currentPage === 'dashboard') renderDashboard();
    else if (currentPage === 'promotionPage') renderPromotionCards();
    else if (currentPage === 'kpi') renderKpiTable();
    else if (currentPage === 'sale') applyReportFilters();
    else if (currentPage === 'customer') { renderNewCustomerTable(); renderTopUpTable(); renderTerminationTable(); renderOutCoverageTable(); }
    else if (currentPage === 'deposit') { renderDepositTable(); updateDepositKpis(); }
    else if (currentPage === 'settings') { renderStaffTable(); renderAccessContent(currentSettingsTab); }
    showToast('Data synced from Google Sheets.', 'success');
  }).catch(function(err) {
    console.warn('Sync down error:', err);
    if (ind) ind.className = 'error';
    if (lbl) lbl.textContent = 'Sync failed';
    if (btn) btn.disabled = false;
    showToast('Sync failed. App will use cached data.', 'error');
  });
}

// Expose sync functions globally so they can be called from buttons or console
window.syncDownAll = syncDownAll;
window.syncUpAll = syncUpAll;

function refreshAllData() {
  return syncDownAll();
}

// ------------------------------------------------------------
// Sample Data
// ------------------------------------------------------------
let itemCatalogue = [
  { id: 'i1',  name: 'Gross Ads',            shortcut: 'GA', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Gross Ads' },
  { id: 'i2',  name: 'Smart@Home',            shortcut: 'SH', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Smart@Home package' },
  { id: 'i3',  name: 'Smart Fiber+',          shortcut: 'SF', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Smart Fiber+' },
  { id: 'i4',  name: 'SmartNas',              shortcut: 'SN', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'SmartNas' },
  { id: 'i5',  name: 'Monthly Upsell',        shortcut: 'MU', group: 'unit',   unit: 'Unit', category: 'Sales', status: 'active', desc: 'Monthly Upsell' },
  { id: 'i6',  name: 'ChangeSIM',             shortcut: 'CS', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Change SIM ($)' },
  { id: 'i7',  name: 'Recharge',              shortcut: 'RC', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Recharge ($)' },
  { id: 'i9',  name: 'Buy Number',            shortcut: 'BN', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Buy Number ($)' },
  { id: 'i10', name: 'SC Dealer',             shortcut: 'BN', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'SC Dealer ($)' },
  { id: 'i11', name: 'Device + Accessories',  shortcut: 'DA', group: 'dollar', currency: '$', price: 1, category: 'Sales', status: 'active', desc: 'Device + Accessories ($)' },
  { id: 'i8',  name: 'Total Revenue',         shortcut: 'RV', group: 'dollar', currency: '$', price: 1, noAutoSum: true, category: 'Sales', status: 'active', desc: 'Total Revenue ($)' },
];

let saleRecords = [];

let newCustomers = [];

let topUpList = [];

let terminationList = [];

let outCoverageList = [];

const DEFAULT_AGENT_USERS = [
  { id: 'u2', name: 'RIM SARAY', username: 'rim.saray', password: 'Tramkak@2026', role: 'Supervisor', branch: 'Express_Tramkak', status: 'active', email: 'rim.saray1@smart.com.kh' },
  { id: 'u3', name: 'KUN CHAMNAN', username: 'kun.chamnan', password: 'Tramkak@2026', role: 'Agent', branch: '', status: 'active', email: 'kun.chamnan@smart.com.kh' },
];

let staffList = [
  { id: 'u1', name: 'Admin', username: 'admin', password: 'Tramkak@2026', role: 'Admin', branch: '', status: 'active' },
].concat(DEFAULT_AGENT_USERS);

let kpiList = [];

let promotionList = [];

let depositList = [];

let coverageLocations = [];

let notifications = [];

// ------------------------------------------------------------
// localStorage Persistence
// ------------------------------------------------------------
const LS_KEYS = {
  items: 'smart5g_items',
  sales: 'smart5g_sales',
  customers: 'smart5g_customers',
  topup: 'smart5g_topup',
  terminations: 'smart5g_terminations',
  outCoverage: 'smart5g_out_coverage',
  staff: 'smart5g_staff',
  kpis: 'smart5g_kpis',
  promotions: 'smart5g_promotions',
  deposits: 'smart5g_deposits',
  coverage: 'smart5g_coverage',
  session: 'smart5g_session'
};

function lsSave(key, data) {
  try { localStorage.setItem(key, JSON.stringify(data)); } catch (e) { console.warn('localStorage save failed:', e); }
}

function lsLoad(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) { return fallback; }
}

function saveAllData() {
  lsSave(LS_KEYS.items, itemCatalogue);
  lsSave(LS_KEYS.sales, saleRecords);
  lsSave(LS_KEYS.customers, newCustomers);
  lsSave(LS_KEYS.topup, topUpList);
  lsSave(LS_KEYS.terminations, terminationList);
  lsSave(LS_KEYS.outCoverage, outCoverageList);
  lsSave(LS_KEYS.staff, staffList);
  lsSave(LS_KEYS.kpis, kpiList);
  lsSave(LS_KEYS.promotions, promotionList);
  lsSave(LS_KEYS.deposits, depositList);
  lsSave(LS_KEYS.coverage, coverageLocations);
}

function loadAllData() {
  itemCatalogue = lsLoad(LS_KEYS.items, itemCatalogue);
  saleRecords = lsLoad(LS_KEYS.sales, saleRecords);
  newCustomers = lsLoad(LS_KEYS.customers, newCustomers);
  topUpList = lsLoad(LS_KEYS.topup, topUpList);
  terminationList = lsLoad(LS_KEYS.terminations, terminationList);
  outCoverageList = lsLoad(LS_KEYS.outCoverage, outCoverageList);
  staffList = lsLoad(LS_KEYS.staff, staffList);
  kpiList = lsLoad(LS_KEYS.kpis, kpiList);
  promotionList = lsLoad(LS_KEYS.promotions, promotionList);
  depositList = lsLoad(LS_KEYS.deposits, depositList);
  coverageLocations = lsLoad(LS_KEYS.coverage, coverageLocations);
  // Ensure admin user always exists
  var hasAdmin = staffList.some(function(u) { return u.username === 'admin' && u.role === 'Admin'; });
  if (!hasAdmin) {
    staffList.unshift({ id: 'u1', name: 'Admin', username: 'admin', password: 'Tramkak@2026', role: 'Admin', branch: '', status: 'active' });
  }
  // Ensure default agent users always exist
  DEFAULT_AGENT_USERS.forEach(function(du) {
    if (!staffList.some(function(u) { return u.username === du.username; })) {
      staffList.push(du);
    }
  });
}

// ------------------------------------------------------------
// Helper Functions
// ------------------------------------------------------------
function g(id) { return document.getElementById(id); }
function rv(id) { const el = g(id); return el ? el.value.trim() : ''; }
function rt(id) { const el = g(id); return el ? el.value : ''; }
function $$(sel) { return document.querySelectorAll(sel); }
function uid() { return '_' + Math.random().toString(36).substr(2, 9); }
function ini(name) { return name.split(' ').map(w => w[0]).join('').toUpperCase().substring(0, 2); }
function fmtMoney(n, cur) { cur = cur !== undefined ? cur : '$'; return cur + Number(n).toFixed(2); }
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function showToast(message, type) {
  var toast = document.createElement('div');
  var bg = (type === 'success') ? '#1B7D3D' : (type === 'error') ? '#C62828' : (type === 'warning') ? '#E65100' : '#333';
  toast.style.cssText = 'position:fixed;bottom:24px;right:24px;background:' + bg + ';color:#fff;padding:12px 20px;border-radius:10px;font-size:.875rem;font-weight:500;box-shadow:0 4px 16px rgba(0,0,0,.2);z-index:10000;opacity:0;transition:opacity .25s;max-width:320px;';
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(function() { toast.style.opacity = '1'; }, 10);
  setTimeout(function() { toast.style.opacity = '0'; setTimeout(function() { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 300); }, 3000);
}

// Professional alert dialog (replaces native alert)
function showAlert(message, type, title) {
  type = type || 'error';
  var iconMap = { error: 'fa-circle-xmark', success: 'fa-circle-check', warning: 'fa-triangle-exclamation', info: 'fa-circle-info' };
  var titleMap = { error: 'Error', success: 'Success', warning: 'Warning', info: 'Information' };
  var classMap = { error: 'dialog-icon-error', success: 'dialog-icon-success', warning: 'dialog-icon-warning', info: 'dialog-icon-info' };
  var iconEl = g('alert-icon');
  var iconWrap = g('alert-icon-wrap');
  var titleEl = g('alert-title');
  var msgEl = g('alert-message');
  if (iconEl) { iconEl.className = 'fas ' + (iconMap[type] || iconMap.error); }
  if (iconWrap) { iconWrap.className = 'dialog-icon-wrap ' + (classMap[type] || classMap.error); }
  if (titleEl) { titleEl.textContent = title || titleMap[type] || 'Alert'; }
  if (msgEl) { msgEl.textContent = message; }
  openModal('modal-alert');
}

// Professional confirm dialog (replaces native confirm)
var _confirmCallback = null;
function showConfirm(message, onConfirm, title, confirmText, isDanger) {
  _confirmCallback = onConfirm || null;
  var titleEl = g('confirm-title');
  var msgEl = g('confirm-message');
  var btn = g('confirm-ok-btn');
  var iconEl = g('confirm-icon');
  var iconWrap = g('confirm-icon-wrap');
  if (titleEl) titleEl.textContent = title || 'Confirm Action';
  if (msgEl) msgEl.textContent = message;
  if (btn) {
    btn.textContent = confirmText || 'Confirm';
    btn.className = isDanger === false ? 'btn btn-primary' : 'btn btn-danger';
  }
  if (iconEl) { iconEl.className = isDanger === false ? 'fas fa-circle-check' : 'fas fa-trash-can'; }
  if (iconWrap) { iconWrap.className = isDanger === false ? 'dialog-icon-wrap dialog-icon-info' : 'dialog-icon-wrap dialog-icon-delete'; }
  openModal('modal-confirm');
}

function _onConfirmAction() {
  closeModal('modal-confirm');
  if (typeof _confirmCallback === 'function') {
    var cb = _confirmCallback;
    _confirmCallback = null;
    cb();
  }
}

// Populate branch dropdowns
function getBranches() {
  return [...new Set(staffList.map(function(u) { return u.branch; }).filter(Boolean))].sort();
}

function populateBranchSelects() {
  const branches = getBranches();
  const branchSelectIds = ['sale-branch', 'nc-branch', 'tu-branch', 'term-branch', 'dep-branch'];
  branchSelectIds.forEach(function(id) {
    const sel = g(id);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = '<option value="">Select branch</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (current === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  });
}

// ------------------------------------------------------------
// Date Helpers
// ------------------------------------------------------------
function ymOf(dateStr) { return dateStr ? dateStr.substring(0, 7) : ''; }
function ymNow() { return new Date().toISOString().substring(0, 7); }
function ymPrev() { const d = new Date(); d.setMonth(d.getMonth() - 1); return d.toISOString().substring(0, 7); }
function ymLabel(ym) {
  const parts = ym.split('-');
  return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+parts[1] - 1] + ' ' + parts[0];
}
function last7Months() {
  const r = [];
  const d = new Date();
  for (let i = 6; i >= 0; i--) {
    const t = new Date(d.getFullYear(), d.getMonth() - i, 1);
    r.push(t.toISOString().substring(0, 7));
  }
  return r;
}
function pctChange(curr, prev) {
  if (!prev) return null;
  return Math.round((curr - prev) / prev * 100);
}
function setTrend(elId, curr, prev) {
  const el = g(elId);
  if (!el) return;
  const pct = pctChange(curr, prev);
  if (pct === null) {
    el.innerHTML = '<i class="fas fa-minus"></i> N/A';
    el.className = 'trend-badge trend-flat';
  } else if (pct > 0) {
    el.innerHTML = '<i class="fas fa-arrow-up"></i> ' + pct + '%';
    el.className = 'trend-badge trend-up';
  } else if (pct < 0) {
    el.innerHTML = '<i class="fas fa-arrow-down"></i> ' + Math.abs(pct) + '%';
    el.className = 'trend-badge trend-down';
  } else {
    el.innerHTML = '<i class="fas fa-minus"></i> 0%';
    el.className = 'trend-badge trend-flat';
  }
}
function destroyChart(ref) {
  if (ref) { try { ref.destroy(); } catch (e) {} }
  return null;
}
function clearCanvas(id) {
  const c = g(id);
  if (c) { const ctx = c.getContext('2d'); ctx.clearRect(0, 0, c.width, c.height); }
}

// ------------------------------------------------------------
// Navigation
// ------------------------------------------------------------
function navigateTo(page, btn) {
  $$('.page').forEach(function(p) { p.classList.remove('active'); });
  const pg = g('page-' + page);
  if (pg) pg.classList.add('active');
  currentPage = page;

  $$('.nav-item').forEach(function(li) { li.classList.remove('active'); });
  if (btn) {
    const li = btn.closest ? btn.closest('.nav-item') : null;
    if (li) li.classList.add('active');
  }

  const titles = {
    dashboard: 'Dashboard',
    promotionPage: currentPromoView === 'expired' ? 'Expired Promotion' : 'New Promotion',
    deposit: 'Deposit',
    sale: 'Sale',
    kpi: 'KPI Setting',
    customer: 'Customer',
    settings: 'Settings',
    'inv-sale-stock': 'Inventory – Sale Stock',
    'inv-shop-stock': 'Inventory – Shop Stock',
    coverage: 'Coverage'
  };
  const titleEl = g('page-title');
  if (titleEl) titleEl.textContent = titles[page] || page;

  populateBranchSelects();

  if (page === 'dashboard') renderDashboard();
  if (page === 'promotionPage') renderPromotionCards();
  if (page === 'kpi') { initKpiMonthPicker(); renderKpiTable(); }
  if (page === 'deposit') { renderDepositTable(); updateDepositKpis(); }
  if (page === 'sale') { renderItemChips(); applyReportFilters(); }
  if (page === 'customer') {
    renderNewCustomerTable();
    renderTopUpTable();
    renderTerminationTable();
    renderOutCoverageTable();
  }
  if (page === 'settings') {
    renderStaffTable();
    renderAccessContent(currentSettingsTab);
  }
  if (page === 'inv-sale-stock') {
    // Reset to stock tab
    $$('.inv-tab-content').forEach(function(c) { c.classList.remove('active'); });
    var tc = g('invtab-content-stock'); if (tc) tc.classList.add('active');
    $$('.tab-btn[id^="invtab-"]').forEach(function(b) { b.classList.remove('active'); });
    var stockTabBtn = g('invtab-stock'); if (stockTabBtn) stockTabBtn.classList.add('active');
    renderInvSaleStock();
  }
  if (page === 'inv-shop-stock') {
    renderShopStockTable();
  }
  if (page === 'coverage') {
    initCoveragePage();
  }
}

function toggleSubmenu(id, elOrId) {
  const sub = g(id);
  if (!sub) return;
  const isOpen = sub.classList.contains('open');
  $$('.submenu').forEach(function(s) { s.classList.remove('open'); });
  $$('.has-submenu').forEach(function(li) { li.classList.remove('submenu-open'); });
  if (!isOpen) {
    sub.classList.add('open');
    const li = typeof elOrId === 'string' ? g(elOrId) : (elOrId && elOrId.closest ? elOrId.closest('.has-submenu') : null);
    if (li) li.classList.add('submenu-open');
  }
}

function openSaleTab(tab, btn) {
  switchSaleTab(tab);
  $$('.sale-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}

function switchSaleTab(tab) {
  currentSaleTab = tab;
  $$('.sale-tab-content').forEach(function(c) { c.classList.remove('active'); });
  const tc = g('sale-tab-' + tab);
  if (tc) tc.classList.add('active');
}

function openCustomerTab(tab, btn) {
  switchCustomerTab(tab);
  $$('.customer-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}

function switchCustomerTab(tab) {
  currentCustomerTab = tab;
  $$('.customer-tab-content').forEach(function(c) { c.classList.remove('active'); });
  $$('.tab-content').forEach(function(c) { c.classList.remove('active'); });
  // Update tab button states
  $$('.tab-btn').forEach(function(b) {
    if (b.getAttribute('data-tab') === tab) b.classList.add('active');
    else if (['new-customer','topup','termination','out-coverage'].includes(b.getAttribute('data-tab'))) b.classList.remove('active');
  });
  const tc = g('tab-content-' + tab);
  if (tc) tc.classList.add('active');
}

function openSettingsTab(tab, btn) {
  switchSettingsTab(tab);
  $$('.settings-tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  renderAccessContent(tab);
}

function switchSettingsTab(tab) {
  currentSettingsTab = tab;
  $$('.settings-tab-content').forEach(function(c) { c.classList.remove('active'); });
  // Update tab button states
  $$('.tab-btn').forEach(function(b) {
    if (b.getAttribute('data-tab') === tab) b.classList.add('active');
    else if (['permission'].includes(b.getAttribute('data-tab'))) b.classList.remove('active');
  });
  const tc = g('stab-content-' + tab);
  if (tc) tc.classList.add('active');
}

function renderAccessContent(tab) {
  const allowed = TAB_PERM[currentRole] || [];
  var banner = g('settings-contact-banner');
  if (banner) banner.style.display = (currentRole !== 'admin' && currentRole !== 'cluster') ? '' : 'none';
  if (!allowed.includes(tab)) {
    const tc = g('stab-content-' + tab);
    if (tc) {
      tc.innerHTML = '<div class="access-denied"><i class="fas fa-lock fa-3x" style="color:#BDBDBD;margin-bottom:12px;"></i><h3 style="color:#555;">Access Denied</h3><p style="color:#999;">You do not have permission to access this section.</p></div>';
    }
  } else {
    if (tab === 'permission') renderStaffTable();
  }
}

function setPromoView(view) {
  currentPromoView = view;
  var sectionAvailable = g('promo-section-available');
  var sectionExpired = g('promo-section-expired');
  var addBtn = g('promo-new-btn');
  var titleEl = g('promo-page-title');
  if (view === 'new') {
    if (sectionAvailable) sectionAvailable.style.display = '';
    if (sectionExpired) sectionExpired.style.display = 'none';
    if (addBtn) addBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';
    if (titleEl) titleEl.innerHTML = '<i class="fas fa-tags" style="color:#1B7D3D;margin-right:8px"></i>New Promotion';
    var pageTitleEl = g('page-title');
    if (pageTitleEl) pageTitleEl.textContent = 'New Promotion';
  } else {
    if (sectionAvailable) sectionAvailable.style.display = 'none';
    if (sectionExpired) sectionExpired.style.display = '';
    if (addBtn) addBtn.style.display = 'none';
    if (titleEl) titleEl.innerHTML = '<i class="fas fa-clock-rotate-left" style="color:#1B7D3D;margin-right:8px"></i>Expired Promotion';
    var pageTitleEl = g('page-title');
    if (pageTitleEl) pageTitleEl.textContent = 'Expired Promotion';
  }
}

function openPromoSubMenu(view, el) {
  navigateTo('promotionPage', null);
  setPromoView(view);
  setActiveSubItem(el);
}

function openSettingsMenu(el) {
  navigateTo('settings', null);
  switchSettingsTab('permission');
  setActiveSubItem(el);
}

function setActiveSubItem(el) {
  $$('.submenu-item').forEach(function(li) { li.classList.remove('active'); });
  if (el) el.classList.add('active');
}

// ------------------------------------------------------------
// Role Switcher
// ------------------------------------------------------------
function switchRole(role) {
  currentRole = role;
  const roleNames = { admin: 'Admin User', cluster: 'Cluster User', supervisor: 'Supervisor User', agent: 'Agent User', user: 'Agent User' };
  const roleBadges = { admin: 'Admin', cluster: 'Cluster', supervisor: 'Supervisor', agent: 'Agent', user: 'Agent' };
  const roleColors = { admin: '#1B7D3D', cluster: '#6A1B9A', supervisor: '#1565C0', agent: '#E65100', user: '#E65100' };

  // For demo role switching: pick an appropriate representative user from staffList
  if (role === 'supervisor') {
    var sup = staffList.find(function(u) { return u.role === 'Supervisor' && u.status === 'active'; });
    if (sup) currentUser = sup;
  } else if (role === 'agent') {
    var agent = staffList.find(function(u) { return u.role === 'Agent' && u.status === 'active'; });
    if (agent) currentUser = agent;
  } else if (role === 'admin' || role === 'cluster') {
    var adminUser = staffList.find(function(u) { return (u.role === 'Admin' || u.role === 'Cluster') && u.status === 'active'; });
    if (adminUser) currentUser = adminUser;
  }

  const nameEl = g('topbar-name');
  const roleEl = g('topbar-role');
  const avatarEl = g('topbar-avatar');

  if (nameEl) nameEl.textContent = currentUser ? currentUser.name : roleNames[role];
  if (roleEl) { roleEl.textContent = roleBadges[role]; roleEl.style.background = roleColors[role]; }
  if (avatarEl) { avatarEl.textContent = ini(currentUser ? currentUser.name : roleNames[role]); avatarEl.style.background = roleColors[role]; }

  const rb = g('role-widget-btn');
  if (rb) { const lbl = rb.querySelector('#role-widget-label'); if (lbl) lbl.textContent = roleBadges[role]; }

  const wd = g('role-widget-dropdown');
  if (wd) wd.style.display = 'none';

  var banner = g('settings-contact-banner');
  if (banner) banner.style.display = (currentRole !== 'admin' && currentRole !== 'cluster') ? '' : 'none';

  var newBtn = g('promo-new-btn');
  if (newBtn) newBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  var saleNewBtn = g('sale-new-btn');
  if (saleNewBtn) saleNewBtn.style.display = (currentRole === 'cluster') ? 'none' : '';

  var covAddBtn = g('cov-add-btn');
  if (covAddBtn) covAddBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  if (currentPage === 'dashboard') renderDashboard();
  if (currentPage === 'settings') renderAccessContent(currentSettingsTab);
  if (currentPage === 'kpi') renderKpiTable();
  if (currentPage === 'sale') applyReportFilters();
  if (currentPage === 'coverage') initCoveragePage();
  if (currentPage === 'customer') { renderNewCustomerTable(); renderTopUpTable(); renderTerminationTable(); renderOutCoverageTable(); }
  if (currentPage === 'deposit') { renderDepositTable(); updateDepositKpis(); }

  // Always refresh item chips so add-button and chip interactivity reflect new role
  renderItemChips();
}

function toggleRoleWidget() {
  const wd = g('role-widget-dropdown');
  if (!wd) return;
  wd.style.display = wd.style.display === 'block' ? 'none' : 'block';
}

// ------------------------------------------------------------
// Modal Helpers
// ------------------------------------------------------------
function openModal(id) {
  populateBranchSelects();
  const el = g(id);
  if (el) { el.style.display = 'flex'; setTimeout(function() { el.classList.add('active'); }, 10); }
}

function closeModal(id) {
  const el = g(id);
  if (el) { el.classList.remove('active'); setTimeout(function() { el.style.display = 'none'; }, 300); }
  if (id === 'modal-topUp') {
    if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
  }
}

function handleOverlay(e, id) {
  if (e.target === e.currentTarget) closeModal(id);
}

function openAddModal(type) {
  if (type === 'item') openItemModal();
  else if (type === 'sale') openNewSaleModal();
  else if (type === 'new-customer') openCustomerModal('new-customer');
  else if (type === 'topup') openCustomerModal('topup');
  else if (type === 'termination') openCustomerModal('termination');
  else if (type === 'kpi') openKpiModal();
  else if (type === 'user') openUserModal();
  else if (type === 'promotion') openPromotionModal();
  else if (type === 'deposit') openDepositModal();
}

function togglePwd(inputId, eyeId) {
  const inp = g(inputId);
  const eye = g(eyeId);
  if (!inp) return;
  if (inp.type === 'password') {
    inp.type = 'text';
    if (eye) eye.className = 'fas fa-eye-slash';
  } else {
    inp.type = 'password';
    if (eye) eye.className = 'fas fa-eye';
  }
}

// ------------------------------------------------------------
// Item Catalogue
// ------------------------------------------------------------
function openItemModal(item) {
  if (currentRole !== 'admin') { showAlert('Only admin can manage items.', 'error'); return; }
  const form = g('form-addItem');
  if (form) form.reset();

  selectItemGroup('unit');
  g('item-edit-id').value = '';

  const title = g('modal-addItem-title');
  const btn = g('item-submit-btn');

  if (item) {
    if (title) title.textContent = 'Edit Item';
    if (btn) btn.textContent = 'Update Item';
    g('item-edit-id').value = item.id;
    g('item-name').value = item.name || '';
    g('item-shortcut').value = item.shortcut || '';
    g('item-category').value = item.category || '';
    g('item-status').value = item.status || 'active';
    g('item-desc').value = item.desc || '';

    selectItemGroup(item.group || 'unit');

    if (item.group === 'unit') {
      const unitSel = g('item-unit');
      if (unitSel) {
        if (KNOWN_UNITS.includes(item.unit)) {
          unitSel.value = item.unit;
        } else {
          unitSel.value = 'custom';
          const cu = g('item-custom-unit');
          if (cu) { cu.value = item.unit || ''; cu.style.display = ''; }
        }
      }
    } else {
      const curSel = g('item-currency');
      if (curSel) {
        if (KNOWN_CURS.includes(item.currency)) {
          curSel.value = item.currency;
        } else {
          curSel.value = 'custom';
          const cc = g('item-custom-currency');
          if (cc) { cc.value = item.currency || ''; cc.style.display = ''; }
        }
      }
      const priceEl = g('item-price');
      if (priceEl) priceEl.value = item.price || '';
    }
  } else {
    if (title) title.textContent = 'Add Item';
    if (btn) btn.textContent = 'Add Item';
  }

  openModal('modal-addItem');
}

function selectItemGroup(grp) {
  itemGroupSelected = grp;
  const unitBtn = g('grp-btn-unit');
  const dollarBtn = g('grp-btn-dollar');
  const unitFields = g('item-unit-fields');
  const dollarFields = g('item-dollar-fields');

  if (grp === 'unit') {
    if (unitBtn) unitBtn.classList.add('active');
    if (dollarBtn) dollarBtn.classList.remove('active');
    if (unitFields) unitFields.style.display = '';
    if (dollarFields) dollarFields.style.display = 'none';
  } else {
    if (dollarBtn) dollarBtn.classList.add('active');
    if (unitBtn) unitBtn.classList.remove('active');
    if (unitFields) unitFields.style.display = 'none';
    if (dollarFields) dollarFields.style.display = '';
  }
}

function handleCurrencySelectChange() {
  const sel = g('item-currency');
  const inp = g('item-custom-currency');
  if (!inp) return;
  inp.style.display = sel && sel.value === 'custom' ? '' : 'none';
}

function handleUnitSelectChange() {
  const sel = g('item-unit');
  const inp = g('item-custom-unit');
  if (!inp) return;
  inp.style.display = sel && sel.value === 'custom' ? '' : 'none';
}

function submitItem(e) {
  e.preventDefault();
  const editId = rv('item-edit-id');
  const name = rv('item-name');
  const shortcut = rv('item-shortcut');
  const category = rv('item-category');
  const status = rv('item-status');
  const desc = rv('item-desc');
  const grp = itemGroupSelected;

  if (!name) { showAlert('Please enter item name'); return; }

  let unit = '', currency = '', price = 0;
  if (grp === 'unit') {
    const unitSel = g('item-unit');
    unit = unitSel && unitSel.value === 'custom' ? rv('item-custom-unit') : rv('item-unit');
  } else {
    const curSel = g('item-currency');
    currency = curSel && curSel.value === 'custom' ? rv('item-custom-currency') : rv('item-currency');
    price = parseFloat(rv('item-price')) || 0;
  }

  const obj = { id: editId || uid(), name: name, shortcut: shortcut, group: grp, unit: unit, currency: currency, price: price, category: category, status: status, desc: desc };

  if (editId) {
    const idx = itemCatalogue.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) itemCatalogue[idx] = obj;
  } else {
    itemCatalogue.push(obj);
  }

  closeModal('modal-addItem');
  renderItemChips();
  if (currentPage === 'dashboard') renderDashboard();
  saveAllData();
  showToast(editId ? 'Item updated.' : 'Item added successfully.', 'success');
}

function editItem(id) {
  if (currentRole !== 'admin') { showAlert('Only admin can edit items.', 'error'); return; }
  const item = itemCatalogue.find(function(x) { return x.id === id; });
  if (item) openItemModal(item);
}

function deleteItem(id) {
  if (currentRole !== 'admin') { showAlert('Only admin can delete items.', 'error'); return; }
  showConfirm('Are you sure you want to delete this item? This action cannot be undone.', function() {
    itemCatalogue = itemCatalogue.filter(function(x) { return x.id !== id; });
    renderItemChips();
    if (currentPage === 'dashboard') renderDashboard();
    saveAllData();
    showToast('Item deleted.', 'success');
  }, 'Delete Item', 'Delete');
}

function renderItemChips() {
  const strip = g('items-strip');
  if (!strip) return;
  const isAdmin = currentRole === 'admin';
  const addBtn = g('item-add-btn');
  if (addBtn) addBtn.style.display = isAdmin ? '' : 'none';
  const active = itemCatalogue.filter(function(x) { return x.status === 'active'; });
  if (!active.length) {
    strip.innerHTML = isAdmin
      ? '<span style="color:#999;font-size:0.85rem;">No items in catalogue. <a href="#" onclick="openAddModal(\'item\');return false;">Add Item</a></span>'
      : '<span style="color:#999;font-size:0.85rem;">No items in catalogue.</span>';
    return;
  }
  strip.innerHTML = active.map(function(item) {
    const chipClass = item.group === 'unit' ? 'item-chip-unit' : 'item-chip-dollar';
    if (isAdmin) {
      return '<span class="item-chip ' + chipClass + '" onclick="editItem(\'' + esc(item.id) + '\')" title="Click to edit: ' + esc(item.name) + '">' +
        esc(item.shortcut || item.name) + '</span>';
    }
    return '<span class="item-chip ' + chipClass + '" style="cursor:default;" title="' + esc(item.name) + '">' +
      esc(item.shortcut || item.name) + '</span>';
  }).join('');
}

// ------------------------------------------------------------
// New Sale Modal
// ------------------------------------------------------------
function openNewSaleModal(sale) {
  const form = g('form-newSale');
  if (form) form.reset();
  g('sale-edit-id').value = '';

  const title = g('modal-newSale-title');
  const btn = g('sale-submit-btn');

  const unitContainer = g('sale-unit-items');
  const dollarContainer = g('sale-dollar-items');

  if (unitContainer) {
    unitContainer.innerHTML = '<div class="sale-items-grid">' + UNIT_SALE_ITEMS.map(function(item) {
      return '<div class="sic-card sic-card-unit">' +
        '<div class="sic-label">' + esc(item.name) + '</div>' +
        '<input type="number" class="sic-input" id="sic-' + esc(item.id) + '" min="0" value="" placeholder="0">' +
        '</div>';
    }).join('') + '</div>';
  }

  if (dollarContainer) {
    dollarContainer.innerHTML = '<div class="sale-items-grid">' + DOLLAR_SALE_ITEMS.map(function(item) {
      return '<div class="sic-card sic-card-dollar">' +
        '<div class="sic-label">' + esc(item.name) + '</div>' +
        '<input type="number" class="sic-input sic-dollar-input" id="sic-' + esc(item.id) + '" min="0" step="0.01" value="" placeholder="0.00">' +
        '</div>';
    }).join('') + '</div>';
  }

  const revTotalEl = g('sale-revenue-total');
  if (revTotalEl) {
    revTotalEl.style.display = '';
    revTotalEl.innerHTML = '<div class="sale-revenue-total-bar"><i class="fas fa-calculator"></i> Total Revenue (Auto Sum): <span id="sale-revenue-total-value">$0.00</span></div>';
  }

  function updateSaleRevenueTotal() {
    var sum = 0;
    DOLLAR_SALE_ITEMS.forEach(function(item) {
      var inp = g('sic-' + item.id);
      if (inp) sum += parseFloat(inp.value) || 0;
    });
    var el = g('sale-revenue-total-value');
    if (el) el.textContent = '$' + sum.toFixed(2);
  }

  if (dollarContainer) {
    dollarContainer.querySelectorAll('.sic-dollar-input').forEach(function(inp) {
      inp.addEventListener('input', function() { updateSaleRevenueTotal(); });
    });
  }

  populateBranchSelects();

  if (sale) {
    if (title) title.textContent = 'Edit Sale';
    if (btn) btn.textContent = 'Update Sale';
    g('sale-edit-id').value = sale.id;
    g('sale-agent-name').value = sale.agent || '';
    const brSel = g('sale-branch');
    if (brSel) brSel.value = sale.branch || '';
    g('sale-date').value = sale.date || '';
    g('sale-remark').value = sale.remark || sale.note || '';
    const expireDateEl = g('sale-expire-date');
    if (expireDateEl) expireDateEl.value = sale.expireDate || '';

    if (sale.items) {
      Object.keys(sale.items).forEach(function(iid) {
        const inp = g('sic-' + iid);
        if (inp) inp.value = sale.items[iid];
      });
    }
    if (sale.dollarItems) {
      Object.keys(sale.dollarItems).forEach(function(iid) {
        const inp = g('sic-' + iid);
        if (inp) inp.value = sale.dollarItems[iid];
      });
    }
    updateSaleRevenueTotal();
  } else {
    if (title) title.textContent = 'New Sale';
    if (btn) btn.textContent = 'Save Sale';
    g('sale-date').value = new Date().toISOString().split('T')[0];
    if (currentUser) {
      const agentEl = g('sale-agent-name');
      if (agentEl) { agentEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agentEl.readOnly = true; }
      const brEl = g('sale-branch');
      if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
    }
  }

  openModal('modal-newSale');
}

function submitSale(e) {
  e.preventDefault();
  const editId = rv('sale-edit-id');
  const agent = rv('sale-agent-name');
  const branch = rv('sale-branch');
  const date = rv('sale-date');
  const expireDate = rv('sale-expire-date');
  const note = rv('sale-remark');

  if (!agent) { showAlert('Please enter agent name'); return; }
  if (!date) { showAlert('Please select date'); return; }

  const items = {}, dollarItems = {};
  let autoRevenue = 0;
  UNIT_SALE_ITEMS.forEach(function(item) {
    const inp = g('sic-' + item.id);
    if (!inp) return;
    const val = parseFloat(inp.value) || 0;
    if (val > 0) items[item.id] = val;
  });
  DOLLAR_SALE_ITEMS.forEach(function(item) {
    const inp = g('sic-' + item.id);
    if (!inp) return;
    const val = parseFloat(inp.value) || 0;
    if (val > 0) dollarItems[item.id] = val;
    autoRevenue += val;
  });
  if (autoRevenue > 0) dollarItems[ITEM_ID_REVENUE] = autoRevenue;

  const now = new Date().toISOString();
  const existingRecord = editId ? saleRecords.find(function(x) { return x.id === editId; }) : null;
  const obj = { id: editId || uid(), agent: agent, branch: branch, date: date, expireDate: expireDate || '', submittedAt: (existingRecord && existingRecord.submittedAt) || now, note: note, items: items, dollarItems: dollarItems };

  if (editId) {
    if (existingRecord && !canModifySaleRecord(existingRecord)) { showSalePermissionError('edit'); return; }
    const idx = saleRecords.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) saleRecords[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a sale record.');
  } else {
    saleRecords.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a new sale record.');
  }

  closeModal('modal-newSale');
  applyReportFilters();
  if (currentPage === 'dashboard') renderDashboard();
  syncSheet('Sales', saleRecords);
  if (!editId) appendDailySale(obj);
  saveAllData();
  showToast(editId ? 'Sale record updated.' : 'Sale record added successfully.', 'success');
}

function editSale(id) {
  const sale = saleRecords.find(function(x) { return x.id === id; });
  if (!sale) return;
  if (!canModifySaleRecord(sale)) { showSalePermissionError('edit'); return; }
  openNewSaleModal(sale);
}

function deleteSale(id) {
  const sale = saleRecords.find(function(x) { return x.id === id; });
  if (!sale) return;
  if (!canModifySaleRecord(sale)) { showSalePermissionError('delete'); return; }
  showConfirm('Are you sure you want to delete this sale record? This action cannot be undone.', function() {
    saleRecords = saleRecords.filter(function(x) { return x.id !== id; });
    applyReportFilters();
    if (currentPage === 'dashboard') renderDashboard();
    syncSheet('Sales', saleRecords);
    saveAllData();
    showToast('Sale record deleted.', 'success');
  }, 'Delete Sale Record', 'Delete');
}

/**
 * Append a single sale record to the DailySale sheet via the Apps Script
 * 'append' action.  The Apps Script handler flattens nested items and
 * dollarItems into individual columns and guarantees the 'Expire Date' column.
 * expireDate is included automatically — blank when not entered by the user.
 */
function appendDailySale(record) {
  if (!GS_URL || !record) return;
  var flat = normalizeRowForSheet(record);
  _gsPost({ sheet: 'DailySale', action: 'append', data: flat })
    .catch(function(err) { console.warn('DailySale append error:', err); });
}

// ------------------------------------------------------------
// Sale Filters & Table
// ------------------------------------------------------------

// Returns the base sale records filtered by the current user's role.
// Agents only see their own records; supervisors see all records in their branch/shop; admin/cluster see all.
function getSaleBaseRecords() {
  if (currentRole === 'agent' && currentUser) {
    return saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  }
  if (currentRole === 'supervisor' && currentUser) {
    return saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  }
  return saleRecords.slice();
}

// Returns true if the current user is permitted to edit or delete the given sale record.
function canModifySaleRecord(sale) {
  if (!sale) return false;
  if (currentRole === 'admin') return true;
  if (currentRole === 'cluster') return false;
  if (!currentUser) return false;
  if (currentRole === 'supervisor') return sale.branch === currentUser.branch;
  if (currentRole === 'agent') return sale.agent === currentUser.name;
  return false;
}

// Returns records filtered by the current user's role.
// Agents only see their own records; supervisors see all records in their branch/shop.
function getBaseRecordsForRole(list) {
  if (currentRole === 'agent' && currentUser) {
    return list.filter(function(r) { return r.agent === currentUser.name; });
  }
  if (currentRole === 'supervisor' && currentUser) {
    return list.filter(function(r) { return r.branch === currentUser.branch; });
  }
  return list.slice();
}

// Returns true if the current user can edit or delete the given customer/deposit record.
function canModifyRecord(record) {
  if (!record) return false;
  if (currentRole === 'admin') return true;
  if (currentRole === 'cluster') return false;
  if (!currentUser) return false;
  if (currentRole === 'supervisor') return record.branch === currentUser.branch;
  if (currentRole === 'agent') return record.agent === currentUser.name;
  return false;
}

// Returns true if the current user can create, edit, or delete KPIs.
// Admin and cluster have full access; supervisors can manage KPIs for their shop/agents; agents cannot.
function canManageKpis() {
  return currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor';
}

// Returns true if a supervisor can modify the given KPI (must be their shop KPI or an agent KPI in their branch).
function canSupervisorModifyKpi(kpi) {
  if (!kpi || !currentUser) return false;
  return (kpi.kpiFor === 'shop' && kpi.assigneeId === currentUser.id) ||
         (kpi.kpiFor === 'agent' && kpi.assigneeBranch === currentUser.branch);
}

// Returns the subset of kpiList visible to the current user.
// Admin/cluster see all KPIs; supervisor/agent see only their own branch's KPIs.
function getVisibleKpis() {
  if (currentRole === 'admin' || currentRole === 'cluster') return kpiList.slice();
  if (!currentUser || !currentUser.branch) return [];
  var userBranch = currentUser.branch;
  return kpiList.filter(function(k) {
    if (k.kpiFor === 'shop') {
      var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
      return sup && sup.branch === userBranch;
    } else if (k.kpiFor === 'agent') {
      return k.assigneeBranch === userBranch;
    }
    return false;
  });
}

// Shows a role-appropriate permission error for sale report modification attempts.
function showSalePermissionError(action) {
  if (currentRole === 'cluster') {
    showAlert('You do not have permission to ' + action + ' sale reports.', 'error');
  } else if (currentRole === 'agent') {
    showAlert('You can only ' + action + ' your own sale reports.', 'error');
  } else {
    showAlert('You can only ' + action + ' sale reports within your branch.', 'error');
  }
}

function applyReportFilters() {
  const baseRecords = getSaleBaseRecords();

  const dateFrom = rv('sale-date-from');
  const dateTo = rv('sale-date-to');
  const agent = rv('sale-filter-agent');
  const branch = rv('sale-filter-branch');

  filteredSales = baseRecords.filter(function(s) {
    if (dateFrom && s.date < dateFrom) return false;
    if (dateTo && s.date > dateTo) return false;
    if (agent && s.agent !== agent) return false;
    if (branch && s.branch !== branch) return false;
    return true;
  });

  renderSaleTable();
  updateSaleKpis();
  // If the summary/chart view is currently active, refresh it too so charts stay up to date
  if (currentReportView === 'summary') {
    var unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
    var dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active'; });
    renderSummaryView(filteredSales, unitItems, dollarItems);
    // Defer chart render by one tick so the DOM update from renderSummaryView completes first
    setTimeout(renderSaleCharts, 50);
  }
}

function clearReportFilters() {
  ['sale-date-from', 'sale-date-to', 'sale-date-search', 'sale-filter-agent', 'sale-filter-branch'].forEach(function(id) {
    const el = g(id); if (el) el.value = '';
  });
  filteredSales = getSaleBaseRecords();
  renderSaleTable();
  updateSaleKpis();
}

function setSaleDateSearch(dateVal) {
  // Set both From and To to the same date for a quick single-date search
  const fromEl = g('sale-date-from');
  const toEl = g('sale-date-to');
  if (fromEl) fromEl.value = dateVal;
  if (toEl) toEl.value = dateVal;
  applyReportFilters();
}

function clearSaleDateSearch() {
  const el = g('sale-date-search');
  if (el) el.value = '';
}

function setReportView(view) {
  currentReportView = view;
  $$('.view-toggle-btn').forEach(function(b) { b.classList.remove('active'); });
  const btn = g('btn-view-' + view);
  if (btn) btn.classList.add('active');

  const tableCard = g('sale-table-card');
  const summaryView = g('sale-summary-view');

  if (view === 'table') {
    if (tableCard) tableCard.style.display = '';
    if (summaryView) summaryView.style.display = 'none';
  } else {
    if (tableCard) tableCard.style.display = 'none';
    if (summaryView) summaryView.style.display = '';
    const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
    const dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active'; });
    renderSummaryView(filteredSales, unitItems, dollarItems);
    renderSaleCharts();
  }
}

function updateSaleKpis() {
  const data = filteredSales;
  let totalRevenue = 0, totalRecharge = 0, totalGrossAds = 0, totalHomeInternet = 0;

  data.forEach(function(s) {
    // Total Revenue: sum of Revenue (RV) dollar item
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) totalRevenue += s.dollarItems[ITEM_ID_REVENUE];
    // Total Recharge: sum of Recharge (RC) dollar item
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) totalRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    // Total Gross Ads: sum of Gross Ads (GA) unit item
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) totalGrossAds += s.items[ITEM_ID_GROSS_ADS];
    // Total Home Internet: sum of Smart@Home (SH) + Smart Fiber+ (SF)
    if (s.items && s.items[ITEM_ID_SMART_HOME]) totalHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) totalHomeInternet += s.items[ITEM_ID_SMART_FIBER];
  });

  const el1 = g('sale-kpi-revenue'); if (el1) el1.textContent = fmtMoney(totalRevenue);
  const el2 = g('sale-kpi-recharge'); if (el2) el2.textContent = fmtMoney(totalRecharge);
  const el3 = g('sale-kpi-gross-ads'); if (el3) el3.textContent = totalGrossAds;
  const el4 = g('sale-kpi-home-internet'); if (el4) el4.textContent = totalHomeInternet;
}

function renderSaleTable() {
  const table = g('sale-table');
  if (!table) return;

  // Role-based base records for filter dropdowns
  const baseRecords = getSaleBaseRecords();

  // Show/hide branch filter: hidden for agent and supervisor (supervisor is locked to their branch)
  const branchFilterWrap = g('sale-branch-filter-wrap');
  if (branchFilterWrap) {
    branchFilterWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor') ? 'none' : '';
  }

  // Show/hide New Sale button: cluster cannot create/edit reports
  const saleNewBtn = g('sale-new-btn');
  if (saleNewBtn) {
    saleNewBtn.style.display = (currentRole === 'cluster') ? 'none' : '';
  }

  // Populate filter dropdowns
  const agentFilter = g('sale-filter-agent');
  const branchFilter = g('sale-filter-branch');
  if (agentFilter) {
    const agents = [...new Set(baseRecords.map(function(s) { return s.agent; }))];
    const curAgent = agentFilter.value;
    agentFilter.innerHTML = '<option value="">All Agents</option>' +
      agents.map(function(a) { return '<option value="' + esc(a) + '"' + (curAgent === a ? ' selected' : '') + '>' + esc(a) + '</option>'; }).join('');
  }
  if (branchFilter) {
    const branches = [...new Set(baseRecords.map(function(s) { return s.branch; }))];
    const curBranch = branchFilter.value;
    branchFilter.innerHTML = '<option value="">All Branches</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (curBranch === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  const data = filteredSales;

  if (!data.length) {
    table.innerHTML = '<thead></thead><tbody><tr><td colspan="20" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-inbox" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No records found</td></tr></tbody>';
    updateTotalBar(0, 0);
    return;
  }

  // Always show all active item columns regardless of whether they have data
  const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  const dollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active' && x.id !== ITEM_ID_REVENUE; });

  let headerRow1 = '<tr><th rowspan="2">Agent</th><th rowspan="2">Branch</th><th rowspan="2">Submit Date</th>';
  if (unitItems.length) headerRow1 += '<th colspan="' + unitItems.length + '" class="th-group-unit">Unit Group</th>';
  if (dollarItems.length) headerRow1 += '<th colspan="' + dollarItems.length + '" class="th-group-dollar">Dollar Group</th>';
  headerRow1 += '<th rowspan="2" class="td-buy-number">Total Revenue</th><th rowspan="2">Remark</th><th rowspan="2">Actions</th></tr>';

  let headerRow2 = '<tr>';
  unitItems.forEach(function(item) { headerRow2 += '<th class="th-unit">' + esc(item.shortcut || item.name) + '</th>'; });
  dollarItems.forEach(function(item) { headerRow2 += '<th class="th-dollar">' + esc(item.shortcut || item.name) + '</th>'; });
  headerRow2 += '</tr>';

  let totalUnits = 0, totalDollar = 0;

  const bodyRows = data.map(function(s) {
    const unitCells = unitItems.map(function(item) {
      const qty = s.items && s.items[item.id] ? s.items[item.id] : 0;
      totalUnits += qty;
      return '<td class="td-unit">' + (qty || '') + '</td>';
    }).join('');
    const dollarCells = dollarItems.map(function(item) {
      const amt = s.dollarItems && s.dollarItems[item.id] ? s.dollarItems[item.id] : 0;
      totalDollar += amt;
      return '<td class="td-dollar">' + (amt > 0 ? fmtMoney(amt, esc(item.currency) + ' ') : '') + '</td>';
    }).join('');
    const saleRev = s.dollarItems && s.dollarItems[ITEM_ID_REVENUE] ? s.dollarItems[ITEM_ID_REVENUE] : 0;

    const submitDate = s.submittedAt ? s.submittedAt.split('T')[0] : s.date;

    // Determine if the current user can edit/delete this record
    const canEdit = canModifySaleRecord(s);

    const avIdx = Math.abs((s.agent.charCodeAt(0) || 0)) % 8;
    return '<tr>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(s.agent)) + '</span>' + esc(s.agent) + '</div></td>' +
      '<td>' + esc(s.branch) + '</td>' +
      '<td>' + esc(submitDate) + '</td>' +
      unitCells +
      dollarCells +
      '<td class="td-buy-number">' + fmtMoney(saleRev) + '</td>' +
      '<td style="color:#888;font-size:0.8rem;">' + esc(s.note || s.remark || '') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editSale(\'' + esc(s.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteSale(\'' + esc(s.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');

  const thead = table.querySelector('thead') || document.createElement('thead');
  const tbody = table.querySelector('tbody') || document.createElement('tbody');
  thead.innerHTML = headerRow1 + headerRow2;
  tbody.innerHTML = bodyRows;
  if (!table.contains(tbody)) table.appendChild(tbody);
  // Ensure thead is always the first child in the DOM for correct sticky-header and display ordering
  if (table.firstChild !== thead) table.insertBefore(thead, table.firstChild);

  updateTotalBar(totalUnits, totalDollar);
}

function updateTotalBar(units, dollar) {
  const bar = g('sale-total-bar');
  if (!bar) return;
  bar.innerHTML =
    '<span class="total-label"><strong>Total Units:</strong> ' + units + '</span>' +
    '<span class="total-label"><strong>Total $:</strong> ' + fmtMoney(dollar) + '</span>';
}

// ------------------------------------------------------------
// Sale CSV Download
// ------------------------------------------------------------
function openDownloadModal() {
  const today = new Date().toISOString().split('T')[0];
  const fromEl = g('dl-date-from');
  const toEl = g('dl-date-to');
  if (fromEl && !fromEl.value) fromEl.value = today.slice(0, 7) + '-01';
  if (toEl && !toEl.value) toEl.value = today;

  // Show/hide role-based filter sections
  const shopSection = g('dl-shop-section');
  const agentSection = g('dl-agent-section');
  if (shopSection) shopSection.style.display = (currentRole === 'cluster') ? '' : 'none';
  if (agentSection) agentSection.style.display = (currentRole === 'supervisor') ? '' : 'none';

  // Populate cluster shop (branch) filter
  if (currentRole === 'cluster') {
    const shopFilter = g('dl-shop-filter');
    if (shopFilter) {
      const branches = [...new Set(saleRecords.map(function(s) { return s.branch; }).filter(Boolean))].sort();
      shopFilter.innerHTML = '<option value="">All Shops</option>' +
        branches.map(function(b) { return '<option value="' + esc(b) + '">' + esc(b) + '</option>'; }).join('');
    }
  }

  // Populate supervisor agent filter (scoped to their branch)
  if (currentRole === 'supervisor' && currentUser) {
    const agentFilter = g('dl-agent-filter');
    if (agentFilter) {
      const branchRecords = saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
      const agents = [...new Set(branchRecords.map(function(s) { return s.agent; }).filter(Boolean))].sort();
      agentFilter.innerHTML = '<option value="">All Agents</option>' +
        agents.map(function(a) { return '<option value="' + esc(a) + '">' + esc(a) + '</option>'; }).join('');
    }
  }

  openModal('modal-saleDownload');
}

function downloadSaleCSV() {
  const dateFrom = rv('dl-date-from');
  const dateTo = rv('dl-date-to');
  const shopFilter = (currentRole === 'cluster') ? rv('dl-shop-filter') : '';
  const agentFilter = (currentRole === 'supervisor') ? rv('dl-agent-filter') : '';

  const base = getSaleBaseRecords();
  const rows = base.filter(function(s) {
    if (dateFrom && s.date < dateFrom) return false;
    if (dateTo && s.date > dateTo) return false;
    if (shopFilter && s.branch !== shopFilter) return false;
    if (agentFilter && s.agent !== agentFilter) return false;
    return true;
  });

  if (!rows.length) { showAlert('No records found for the selected filters.'); return; }

  // Only include item columns that have at least one non-zero value in the filtered rows
  const allUnitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  const allDollarItems = itemCatalogue.filter(function(x) { return x.group === 'dollar' && x.status === 'active' && x.id !== ITEM_ID_REVENUE; });
  const unitItems = allUnitItems.filter(function(item) { return rows.some(function(s) { return s.items && s.items[item.id] > 0; }); });
  const dollarItems = allDollarItems.filter(function(item) { return rows.some(function(s) { return s.dollarItems && s.dollarItems[item.id] > 0; }); });

  const escape = function(v) {
    const s = String(v === null || v === undefined ? '' : v);
    return '"' + s.replace(/"/g, '""') + '"';
  };

  const unitHeaders = unitItems.map(function(x) { return escape(x.shortcut || x.name); });
  const dollarHeaders = dollarItems.map(function(x) { return escape(x.shortcut || x.name); });
  const header = ['Agent', 'Branch', 'Submit Date'].concat(unitHeaders).concat(dollarHeaders).concat(['Total Revenue', 'Remark']).map(escape).join(',');

  const lines = rows.map(function(s) {
    const submitDate = s.submittedAt ? s.submittedAt.split('T')[0] : s.date;
    const unitCols = unitItems.map(function(item) { return escape(s.items && s.items[item.id] ? s.items[item.id] : ''); });
    const dollarCols = dollarItems.map(function(item) { return escape(s.dollarItems && s.dollarItems[item.id] ? s.dollarItems[item.id] : ''); });
    const rev = s.dollarItems && s.dollarItems[ITEM_ID_REVENUE] ? s.dollarItems[ITEM_ID_REVENUE] : '';
    return [escape(s.agent), escape(s.branch), escape(submitDate)]
      .concat(unitCols).concat(dollarCols)
      .concat([escape(rev), escape(s.note || s.remark || '')])
      .join(',');
  });

  const csv = header + '\n' + lines.join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const suffix = (dateFrom || 'earliest') + '_to_' + (dateTo || 'latest');
  a.href = url;
  a.download = 'daily_sale_' + suffix + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  closeModal('modal-saleDownload');
  showToast('Download complete.', 'success');
}

function renderSummaryView(data, unitItems, dollarItems) {
  const container = g('sale-summary-view');
  if (!container) return;

  if (!data.length) {
    container.innerHTML = '<div style="text-align:center;padding:40px;color:#999;"><i class="fas fa-inbox fa-3x" style="display:block;margin-bottom:12px;"></i>No records found</div>';
    return;
  }

  const agentMap = {};
  data.forEach(function(s) {
    if (!agentMap[s.agent]) agentMap[s.agent] = { units: {}, dollars: {}, totalUnits: 0, totalRev: 0 };
    const ag = agentMap[s.agent];
    Object.keys(s.items || {}).forEach(function(iid) {
      ag.units[iid] = (ag.units[iid] || 0) + s.items[iid];
      ag.totalUnits += s.items[iid];
    });
    Object.keys(s.dollarItems || {}).forEach(function(iid) {
      ag.dollars[iid] = (ag.dollars[iid] || 0) + s.dollarItems[iid];
      if (iid === ITEM_ID_REVENUE) ag.totalRev += s.dollarItems[iid];
    });
  });

  const cards = Object.keys(agentMap).map(function(agent, idx) {
    const ag = agentMap[agent];
    const unitRows = unitItems.map(function(item) {
      const qty = ag.units[item.id] || 0;
      return qty ? '<div style="display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:0.8125rem;"><span>' + esc(item.name) + '</span><span style="font-weight:600;color:#1B7D3D;">' + qty + '</span></div>' : '';
    }).join('');
    const dollarRows = dollarItems.map(function(item) {
      const amt = ag.dollars[item.id] || 0;
      return amt ? '<div style="display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #f0f0f0;font-size:0.8125rem;"><span>' + esc(item.name) + '</span><span style="font-weight:600;color:#E65100;">' + fmtMoney(amt, esc(item.currency) + ' ') + '</span></div>' : '';
    }).join('');
    return '<div class="summary-card">' +
      '<div class="summary-card-header">' +
        '<span class="sc-avatar av-' + (idx % 8) + '">' + esc(ini(agent)) + '</span>' +
        '<div><div class="sc-name">' + esc(agent) + '</div><div style="font-size:0.72rem;opacity:0.8;">Rev: ' + fmtMoney(ag.totalRev) + '</div></div>' +
      '</div>' +
      '<div class="summary-card-body">' + (unitRows + dollarRows || '<div style="color:#999;font-size:0.8rem;">No sales</div>') + '</div>' +
      '</div>';
  }).join('');

  const chartRow = '<div class="chart-row" style="margin-top:20px;">' +
    '<div class="chart-card"><div class="chart-card-header"><span class="chart-card-title">Sales by Item</span></div><div class="chart-card-body"><canvas id="cSaleMix"></canvas></div></div>' +
    '<div class="chart-card"><div class="chart-card-header"><span class="chart-card-title">Sales by Agent</span></div><div class="chart-card-body"><canvas id="cSaleAgent"></canvas></div></div>' +
    '</div>';

  container.innerHTML = '<div class="summary-grid">' + cards + '</div>' + chartRow;
  setTimeout(renderSaleCharts, 50);
}

function renderSaleCharts() {
  _cSaleMix = destroyChart(_cSaleMix);
  _cSaleAgent = destroyChart(_cSaleAgent);

  const data = filteredSales.length ? filteredSales : saleRecords;
  const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });

  const mixLabels = unitItems.map(function(x) { return x.name; });
  const mixData = unitItems.map(function(item) {
    let t = 0; data.forEach(function(s) { t += (s.items && s.items[item.id]) ? s.items[item.id] : 0; }); return t;
  });

  const mixCtx = g('cSaleMix');
  if (mixCtx && typeof Chart !== 'undefined' && mixData.some(function(v) { return v > 0; })) {
    _cSaleMix = new Chart(mixCtx, {
      type: 'doughnut',
      data: { labels: mixLabels, datasets: [{ data: mixData, backgroundColor: CHART_PAL, borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
          tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
        }
      }
    });
  }

  const agentMap = {};
  data.forEach(function(s) {
    if (!agentMap[s.agent]) agentMap[s.agent] = { rev: 0, branch: s.branch || '' };
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) {
      agentMap[s.agent].rev += s.dollarItems[ITEM_ID_REVENUE];
    }
  });

  // Group agents by branch for color coding
  const branchList = [];
  Object.keys(agentMap).forEach(function(a) {
    const br = agentMap[a].branch;
    if (br && branchList.indexOf(br) === -1) branchList.push(br);
  });

  const agentLabels = Object.keys(agentMap);
  const agentVals = agentLabels.map(function(a) { return agentMap[a].rev; });
  const agentColors = agentLabels.map(function(a) {
    const bIdx = branchList.indexOf(agentMap[a].branch);
    return CHART_PAL[bIdx >= 0 ? bIdx % CHART_PAL.length : 0];
  });
  const agentBranchLabels = agentLabels.map(function(a) {
    return agentMap[a].branch ? a + ' (' + agentMap[a].branch + ')' : a;
  });

  const agCtx = g('cSaleAgent');
  if (agCtx && typeof Chart !== 'undefined' && agentLabels.length) {
    _cSaleAgent = new Chart(agCtx, {
      type: 'bar',
      data: {
        labels: agentBranchLabels,
        datasets: [{
          label: 'Revenue ($)',
          data: agentVals,
          backgroundColor: agentColors,
          borderColor: agentColors,
          borderWidth: 1,
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
            callbacks: {
              label: function(ctx) { return 'Revenue: $' + Number(ctx.parsed.y).toFixed(2); }
            }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } } },
          y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: function(v) { return '$' + v; } } }
        }
      }
    });
  }
}

// ------------------------------------------------------------
// Dashboard
// ------------------------------------------------------------
function renderDashboard() {
  const ym = ymNow();
  const ymP = ymPrev();

  // Role-based data filtering
  let viewSales = saleRecords;
  if (currentRole === 'agent' && currentUser) {
    viewSales = saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  } else if (currentRole === 'supervisor' && currentUser) {
    viewSales = saleRecords.filter(function(s) { return s.branch === currentUser.branch; });
  }
  // Apply branch filter (available for admin/cluster)
  const branchFilterVal = g('dash-branch-filter') ? g('dash-branch-filter').value : '';
  if (branchFilterVal) {
    viewSales = viewSales.filter(function(s) { return s.branch === branchFilterVal; });
  }
  // Hide branch filter for agent and supervisor (they auto-display their own branch)
  var branchFilterWrap = g('dash-branch-filter-wrap');
  if (branchFilterWrap) branchFilterWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor') ? 'none' : '';

  // Show branch summary table for admin, cluster, and supervisor (supervisor sees only their own branch via filtered viewSales)
  var branchSection = g('dash-branch-section');
  if (branchSection) branchSection.style.display = (currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor') ? '' : 'none';

  const currSales = viewSales.filter(function(s) { return ymOf(s.date) === ym; });
  const prevSales = viewSales.filter(function(s) { return ymOf(s.date) === ymP; });

  let currRevenue = 0, prevRevenue = 0;
  let currRecharge = 0, prevRecharge = 0;
  let currGrossAds = 0, prevGrossAds = 0;
  let currHomeInternet = 0, prevHomeInternet = 0;

  currSales.forEach(function(s) {
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) currRevenue += s.dollarItems[ITEM_ID_REVENUE];
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) currRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) currGrossAds += s.items[ITEM_ID_GROSS_ADS];
    if (s.items && s.items[ITEM_ID_SMART_HOME]) currHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) currHomeInternet += s.items[ITEM_ID_SMART_FIBER];
  });

  prevSales.forEach(function(s) {
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) prevRevenue += s.dollarItems[ITEM_ID_REVENUE];
    if (s.dollarItems && s.dollarItems[ITEM_ID_RECHARGE]) prevRecharge += s.dollarItems[ITEM_ID_RECHARGE];
    if (s.items && s.items[ITEM_ID_GROSS_ADS]) prevGrossAds += s.items[ITEM_ID_GROSS_ADS];
    if (s.items && s.items[ITEM_ID_SMART_HOME]) prevHomeInternet += s.items[ITEM_ID_SMART_HOME];
    if (s.items && s.items[ITEM_ID_SMART_FIBER]) prevHomeInternet += s.items[ITEM_ID_SMART_FIBER];
  });

  const kr = g('kv-revenue'); if (kr) kr.textContent = fmtMoney(currRevenue);
  const krc = g('kv-recharge'); if (krc) krc.textContent = fmtMoney(currRecharge);
  const kg = g('kv-gross-ads'); if (kg) kg.textContent = currGrossAds;
  const kh = g('kv-home-internet'); if (kh) kh.textContent = currHomeInternet;

  setTrend('tr-revenue', currRevenue, prevRevenue);
  setTrend('tr-recharge', currRecharge, prevRecharge);
  setTrend('tr-gross-ads', currGrossAds, prevGrossAds);
  setTrend('tr-home-internet', currHomeInternet, prevHomeInternet);

  // Chart 1: Monthly Trend / Shop KPI Achievement Gauge
  _cTrend = destroyChart(_cTrend);
  clearCanvas('cTrend');
  var trendGaugeOverlay = g('trend-gauge-overlay');

  // For supervisor/agent: show KPI achievement gauge for their own shop
  var shopKpiForGauge = null;
  if ((currentRole === 'supervisor' || currentRole === 'agent') && currentUser) {
    var shopSup = (currentRole === 'supervisor') ? currentUser :
      staffList.find(function(u) { return u.role === 'Supervisor' && u.branch === currentUser.branch; });
    if (shopSup) {
      shopKpiForGauge = kpiList.find(function(k) { return k.kpiFor === 'shop' && k.assigneeId === shopSup.id; });
    }
  }

  const tCtx = g('cTrend');
  if (tCtx && shopKpiForGauge && typeof Chart !== 'undefined') {
    var kpiActualVal = calcKpiActual(shopKpiForGauge);
    var kpiPct = shopKpiForGauge.target > 0 ? Math.round(kpiActualVal / shopKpiForGauge.target * 100) : 0;
    var fillPct = Math.min(kpiPct, 100);
    var gColor = kpiPct >= 100 ? '#1B7D3D' : kpiPct >= 70 ? '#FF9800' : '#E53935';
    _cTrend = new Chart(tCtx, {
      type: 'doughnut',
      data: {
        datasets: [{
          data: [fillPct, 100 - fillPct],
          backgroundColor: [gColor, '#EEEEEE'],
          borderWidth: 0,
          circumference: 180,
          rotation: -90
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        animation: { animateRotate: true, duration: 700 }
      }
    });
    if (trendGaugeOverlay) {
      trendGaugeOverlay.style.display = '';
      var pctEl = g('trend-gauge-pct-text');
      var detailEl = g('trend-gauge-detail');
      if (pctEl) { pctEl.style.color = gColor; pctEl.textContent = kpiPct + '%'; }
      if (detailEl) {
        var tgValueDisplay = shopKpiForGauge.valueMode === 'currency'
          ? fmtMoney(kpiActualVal, shopKpiForGauge.currency + ' ') + ' / ' + fmtMoney(shopKpiForGauge.target, shopKpiForGauge.currency + ' ')
          : Math.round(kpiActualVal * 100) / 100 + ' / ' + shopKpiForGauge.target + (shopKpiForGauge.unit ? ' ' + esc(shopKpiForGauge.unit) : '');
        detailEl.textContent = esc(shopKpiForGauge.name) + ': ' + tgValueDisplay;
      }
    }
  } else {
    if (trendGaugeOverlay) trendGaugeOverlay.style.display = 'none';
    if (tCtx && typeof Chart !== 'undefined') {
      const months = last7Months();
      const monthLabels = months.map(ymLabel);
      const unitsPerMonth = months.map(function(m) {
        let u = 0;
        viewSales.filter(function(s) { return ymOf(s.date) === m; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { u += v; });
        });
        return u;
      });
      const revPerMonth = months.map(function(m) {
        let r = 0;
        viewSales.filter(function(s) { return ymOf(s.date) === m; }).forEach(function(s) {
          Object.keys(s.dollarItems || {}).forEach(function(iid) {
            const item = itemCatalogue.find(function(x) { return x.id === iid; });
            if (item && !item.noAutoSum && !item.noAutoRevenue) r += s.dollarItems[iid] * (item.price || 1);
          });
        });
        return r;
      });
      _cTrend = new Chart(tCtx, {
        type: 'line',
        data: {
          labels: monthLabels,
          datasets: [
            { label: 'Units', data: unitsPerMonth, borderColor: '#1B7D3D', backgroundColor: 'rgba(27,125,61,0.08)', yAxisID: 'y', tension: 0.4, fill: true, pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: '#1B7D3D', borderWidth: 2 },
            { label: 'Revenue ($)', data: revPerMonth, borderColor: '#FF9800', backgroundColor: 'rgba(255,152,0,0.08)', yAxisID: 'y1', tension: 0.4, fill: true, pointRadius: 4, pointHoverRadius: 6, pointBackgroundColor: '#FF9800', borderWidth: 2 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 16, font: { size: 11 } } },
            tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, titleFont: { size: 12 }, bodyFont: { size: 11 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: { position: 'left', title: { display: true, text: 'Units', font: { size: 11 } }, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } } },
            y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Revenue ($)', font: { size: 11 } }, ticks: { font: { size: 11 } } }
          }
        }
      });
    }
  }

  // Chart 2: Item Mix (doughnut)
  _cMix = destroyChart(_cMix);
  clearCanvas('cMix');
  const unitItemsDash = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  const mixData = unitItemsDash.map(function(item) {
    let total = 0;
    viewSales.forEach(function(s) { total += (s.items && s.items[item.id]) ? s.items[item.id] : 0; });
    return total;
  });
  const mCtx = g('cMix');
  if (mCtx && unitItemsDash.length && typeof Chart !== 'undefined') {
    _cMix = new Chart(mCtx, {
      type: 'doughnut',
      data: {
        labels: unitItemsDash.map(function(x) { return x.name; }),
        datasets: [{ data: mixData, backgroundColor: CHART_PAL, borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'circle', padding: 12, font: { size: 11 } } },
          tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
        }
      }
    });
  }

  // Chart 3: Agent Performance (Revenue)
  _cAgent = destroyChart(_cAgent);
  clearCanvas('cAgent');
  const agentRevenue = {};
  currSales.forEach(function(s) {
    if (!(s.agent in agentRevenue)) agentRevenue[s.agent] = 0;
    if (s.dollarItems && s.dollarItems[ITEM_ID_REVENUE]) agentRevenue[s.agent] += s.dollarItems[ITEM_ID_REVENUE];
  });
  const agentNames = Object.keys(agentRevenue);
  const agentVals = agentNames.map(function(a) { return agentRevenue[a]; });
  const aCtx = g('cAgent');
  if (aCtx && agentNames.length && typeof Chart !== 'undefined') {
    _cAgent = new Chart(aCtx, {
      type: 'bar',
      data: {
        labels: agentNames,
        datasets: [{ label: 'Revenue This Month ($)', data: agentVals, backgroundColor: CHART_PAL, borderRadius: 4, borderSkipped: false }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
            callbacks: { label: function(ctx) { return ' ' + fmtMoney(ctx.parsed.x || 0); } }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: function(v) { return fmtMoney(v); } } },
          y: { grid: { display: false }, ticks: { font: { size: 11 } } }
        }
      }
    });
  }

  // Chart 4: Growth vs Last Month
  _cGrowth = destroyChart(_cGrowth);
  clearCanvas('cGrowth');
  // Sync dropdown state
  var growthViewSel = g('growth-chart-view');
  if (growthViewSel && growthViewSel.value !== _growthChartView) growthViewSel.value = _growthChartView;
  const growthLabels = unitItemsDash.map(function(x) { return x.shortcut || x.name; });
  const currItemUnits = unitItemsDash.map(function(item) {
    let t = 0; currSales.forEach(function(s) { t += (s.items && s.items[item.id]) || 0; }); return t;
  });
  const prevItemUnits = unitItemsDash.map(function(item) {
    let t = 0; prevSales.forEach(function(s) { t += (s.items && s.items[item.id]) || 0; }); return t;
  });
  const gCtx = g('cGrowth');
  if (gCtx && unitItemsDash.length && typeof Chart !== 'undefined') {
    if (_growthChartView === 'pct') {
      // Show percentage growth per item
      const pctData = currItemUnits.map(function(curr, i) {
        var prev = prevItemUnits[i];
        if (!prev) return null;
        return Math.round((curr - prev) / prev * 100);
      });
      const barColors = pctData.map(function(v) {
        if (v === null) return '#BDBDBD';
        return v >= 0 ? '#1B7D3D' : '#E53935';
      });
      _cGrowth = new Chart(gCtx, {
        type: 'bar',
        data: {
          labels: growthLabels,
          datasets: [{
            label: '% Growth vs Last Month',
            data: pctData,
            backgroundColor: barColors,
            borderRadius: 5,
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
            tooltip: {
              backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
              callbacks: {
                label: function(ctx) {
                  var v = ctx.parsed.y;
                  if (v === null) return 'N/A';
                  return (v >= 0 ? '+' : '') + v + '%';
                }
              }
            }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: {
              grid: { color: 'rgba(0,0,0,0.05)' },
              ticks: {
                font: { size: 11 },
                callback: function(v) { return v + '%'; }
              }
            }
          }
        }
      });
    } else {
      _cGrowth = new Chart(gCtx, {
        type: 'line',
        data: {
          labels: growthLabels,
          datasets: [
            { label: 'This Month', data: currItemUnits, borderColor: '#1B7D3D', backgroundColor: 'rgba(27,125,61,0.08)', tension: 0.3, fill: false, pointRadius: 5, pointHoverRadius: 7, pointBackgroundColor: '#1B7D3D', pointBorderColor: '#fff', pointBorderWidth: 2, borderWidth: 2 },
            { label: 'Last Month', data: prevItemUnits, borderColor: '#A5D6A7', backgroundColor: 'rgba(165,214,167,0.08)', tension: 0.3, fill: false, pointRadius: 5, pointHoverRadius: 7, pointBackgroundColor: '#A5D6A7', pointBorderColor: '#fff', pointBorderWidth: 2, borderWidth: 2 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { position: 'top', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14, font: { size: 11 } } },
            tooltip: { backgroundColor: 'rgba(26,26,46,0.9)', padding: 10, cornerRadius: 8, bodyFont: { size: 11 } }
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 11 } } },
            y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 } }, beginAtZero: true }
          }
        }
      });
    }
  }

  // Branch summary table (visible for admin/cluster/supervisor)
  const branchTableBody = g('branch-table');
  if (branchTableBody) {
    const branches = [];
    viewSales.forEach(function(s) { if (branches.indexOf(s.branch) < 0) branches.push(s.branch); });
    if (!branches.length) {
      branchTableBody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px;color:#999;">No data</td></tr>';
    } else {
      branchTableBody.innerHTML = branches.map(function(branch) {
        let cU = 0, pU = 0;
        currSales.filter(function(s) { return s.branch === branch; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { cU += v; });
        });
        prevSales.filter(function(s) { return s.branch === branch; }).forEach(function(s) {
          Object.values(s.items || {}).forEach(function(v) { pU += v; });
        });
        const pct = pctChange(cU, pU);
        const trendHtml = pct === null ? '<span class="pill pill-gray">N/A</span>' :
          pct > 0 ? '<span class="pill pill-green">+' + pct + '%</span>' :
          pct < 0 ? '<span class="pill pill-red">' + pct + '%</span>' : '<span class="pill pill-gray">0%</span>';
        return '<tr><td>' + esc(branch) + '</td><td>' + cU + '</td><td>' + pU + '</td><td>' + trendHtml + '</td></tr>';
      }).join('');
    }
  }

  // Branch filter dropdown
  const branchFilter = g('dash-branch-filter');
  if (branchFilter) {
    const branches = getBranches();
    const cur = branchFilter.value;
    branchFilter.innerHTML = '<option value="">All Branches</option>' +
      branches.map(function(b) { return '<option value="' + esc(b) + '"' + (cur === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  // KPI Achievement section
  renderDashboardKpiSection();
}

// ------------------------------------------------------------
// KPI vs Actual Achievement Dashboard
// ------------------------------------------------------------
function getKpiUnitLabel(kpi) {
  if (kpi.itemId) {
    var item = itemCatalogue.find(function(x) { return x.id === kpi.itemId; });
    if (item) return item.name;
  }
  return kpi.unit || '';
}

function calcKpiActual(kpi, ym) {
  if (!ym) ym = ymNow();
  const currSales = saleRecords.filter(function(s) { return ymOf(s.date) === ym; });
  let filtered = currSales;
  if (kpi.kpiFor === 'agent') {
    const agent = staffList.find(function(u) { return u.id === kpi.assigneeId; });
    if (agent) filtered = currSales.filter(function(s) { return s.agent === agent.name; });
  } else if (kpi.kpiFor === 'shop') {
    const sup = staffList.find(function(u) { return u.id === kpi.assigneeId; });
    if (sup) filtered = currSales.filter(function(s) { return s.branch === sup.branch; });
  }
  let actual = 0;
  if (kpi.valueMode === 'unit') {
    if (kpi.itemId) {
      filtered.forEach(function(s) { actual += (s.items && s.items[kpi.itemId]) || 0; });
    } else {
      filtered.forEach(function(s) { Object.values(s.items || {}).forEach(function(v) { actual += v; }); });
    }
  } else {
    filtered.forEach(function(s) {
      Object.keys(s.dollarItems || {}).forEach(function(iid) {
        const item = itemCatalogue.find(function(x) { return x.id === iid; });
        if (item && !item.noAutoSum && !item.noAutoRevenue) actual += s.dollarItems[iid] * (item.price || 1);
      });
    });
  }
  return actual;
}

var _dashKpiView = 'all'; // 'all', 'shop', 'agent', 'branch'
var _growthChartView = 'compare'; // 'compare' or 'pct'

function setGrowthChartView(v) {
  _growthChartView = v;
  renderDashboard();
}

function setDashKpiView(view) {
  _dashKpiView = view;
  $$('.dash-kpi-view-btn').forEach(function(b) {
    b.classList.toggle('active', b.getAttribute('data-view') === view);
  });
  renderDashboardKpiSection();
}

function renderDashboardKpiSection() {
  var section = g('dash-kpi-section');
  if (!section) return;

  // Show/hide the "By Branch" button (only for cluster/admin)
  var branchBtn = g('dash-kpi-branch-btn');
  if (branchBtn) branchBtn.style.display = (currentRole === 'cluster' || currentRole === 'admin') ? '' : 'none';

  // Populate and show/hide branch dropdown for "By Branch" view
  var branchFilterWrap = g('dash-kpi-branch-filter-wrap');
  var branchFilterSel = g('dash-kpi-branch-filter');
  if (branchFilterWrap) branchFilterWrap.style.display = (_dashKpiView === 'branch' && (currentRole === 'cluster' || currentRole === 'admin')) ? '' : 'none';
  if (branchFilterSel && _dashKpiView === 'branch') {
    var allBranches = getBranches();
    var curBranch = branchFilterSel.value;
    branchFilterSel.innerHTML = '<option value="">All Branches</option>' +
      allBranches.map(function(b) { return '<option value="' + esc(b) + '"' + (curBranch === b ? ' selected' : '') + '>' + esc(b) + '</option>'; }).join('');
  }

  // Determine relevant KPIs based on role
  var relevantKpis = [];
  if (currentRole === 'agent' && currentUser) {
    // Agents can view their own agent KPIs AND the shop KPIs for their branch
    var branchSups = staffList.filter(function(u) {
      return u.role === 'Supervisor' && u.branch === currentUser.branch;
    });
    var supIds = branchSups.map(function(s) { return s.id; });
    relevantKpis = kpiList.filter(function(k) {
      return (k.kpiFor === 'agent' && k.assigneeId === currentUser.id) ||
             (k.kpiFor === 'shop' && supIds.indexOf(k.assigneeId) !== -1);
    });
  } else if (currentRole === 'supervisor' && currentUser) {
    relevantKpis = kpiList.filter(function(k) {
      return (k.kpiFor === 'shop' && k.assigneeId === currentUser.id) ||
             (k.kpiFor === 'agent' && k.assigneeBranch === currentUser.branch);
    });
  } else if (currentRole === 'cluster' || currentRole === 'admin') {
    relevantKpis = kpiList.slice();
  }

  // Apply view filter
  if (_dashKpiView === 'shop') {
    relevantKpis = relevantKpis.filter(function(k) { return k.kpiFor === 'shop'; });
  } else if (_dashKpiView === 'agent') {
    relevantKpis = relevantKpis.filter(function(k) { return k.kpiFor === 'agent'; });
  } else if (_dashKpiView === 'branch' && branchFilterSel && branchFilterSel.value) {
    // Filter by branch: for shop KPIs use assignee's branch, for agent KPIs use assigneeBranch
    var selectedBranch = branchFilterSel.value;
    relevantKpis = relevantKpis.filter(function(k) {
      if (k.kpiFor === 'agent') return k.assigneeBranch === selectedBranch;
      if (k.kpiFor === 'shop') {
        var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
        return sup ? sup.branch === selectedBranch : false;
      }
      return false;
    });
  }

  // Show/hide the view toggle buttons based on role
  var toggleWrap = g('dash-kpi-view-toggle');
  if (toggleWrap) {
    toggleWrap.style.display = (currentRole === 'agent' || currentRole === 'supervisor' || currentRole === 'cluster' || currentRole === 'admin') ? '' : 'none';
  }

  if (!relevantKpis.length) {
    section.style.display = 'none';
    return;
  }
  section.style.display = '';

  // Build KPI data with actuals and percentages
  var kpiData = relevantKpis.map(function(k) {
    var assignee = staffList.find(function(u) { return u.id === k.assigneeId; });
    var actual = calcKpiActual(k);
    var pct = k.target > 0 ? Math.round(actual / k.target * 100) : 0;
    return { k: k, assignee: assignee, actual: Math.round(actual * 100) / 100, pct: pct };
  });

  // Destroy existing gauge charts
  _cKpiGauges.forEach(function(c) { if (c) { try { c.destroy(); } catch (e) { console.warn('Chart destroy error:', e); } } });
  _cKpiGauges = [];

  // Color helper for gauge achievement level
  function gaugeColor(pct) { return pct >= 100 ? '#1B7D3D' : pct >= 70 ? '#FF9800' : '#E53935'; }
  function gaugePillClass(pct) { return pct >= 100 ? 'pill-green' : pct >= 70 ? 'pill-orange' : 'pill-red'; }

  // Render gauge cards
  var gaugeGrid = g('dash-kpi-gauge-grid');
  if (gaugeGrid) {
    gaugeGrid.innerHTML = kpiData.map(function(d, i) {
      var color = gaugeColor(d.pct);
      var pctClass = gaugePillClass(d.pct);
      var forPill = d.k.kpiFor === 'shop'
        ? '<span class="pill pill-blue" style="font-size:.7rem;padding:2px 7px;">Shop</span>'
        : '<span class="pill pill-orange" style="font-size:.7rem;padding:2px 7px;">Agent</span>';
      var assigneeName = d.assignee ? esc(d.assignee.name) : (d.k.assigneeBranch ? esc(d.k.assigneeBranch) : '—');
      var valueDisplay = d.k.valueMode === 'currency'
        ? fmtMoney(d.k.target, esc(d.k.currency) + ' ') + ' / ' + fmtMoney(d.actual, esc(d.k.currency) + ' ')
        : (function() { var ul = getKpiUnitLabel(d.k); return d.k.target + ' / ' + d.actual + (ul ? ' ' + esc(ul) : ''); })();
      return '<div class="kpi-gauge-card">' +
        '<div class="kpi-gauge-canvas-wrap">' +
        '<canvas id="kpiGauge_' + i + '" height="130"></canvas>' +
        '<div class="kpi-gauge-pct-text" style="color:' + color + '">' + d.pct + '%</div>' +
        '</div>' +
        '<div class="kpi-gauge-name">' + esc(d.k.name) + '</div>' +
        '<div class="kpi-gauge-assignee">' + forPill + ' <span>' + assigneeName + '</span></div>' +
        '<div class="kpi-gauge-value"><span class="pill ' + pctClass + '" style="font-size:.72rem;">' + valueDisplay + '</span></div>' +
        '</div>';
    }).join('');

    // Create Chart.js gauge (semicircle doughnut) for each KPI
    if (typeof Chart !== 'undefined') {
      kpiData.forEach(function(d, i) {
        var canvas = document.getElementById('kpiGauge_' + i);
        if (!canvas) return;
        var color = gaugeColor(d.pct);
        var fillPct = Math.min(d.pct, 100);
        var chart = new Chart(canvas, {
          type: 'doughnut',
          data: {
            datasets: [{
              data: [fillPct, 100 - fillPct],
              backgroundColor: [color, '#EEEEEE'],
              borderWidth: 0,
              circumference: 180,
              rotation: -90
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '72%',
            plugins: {
              legend: { display: false },
              tooltip: { enabled: false }
            },
            animation: { animateRotate: true, duration: 700 }
          }
        });
        _cKpiGauges.push(chart);
      });
    }
  }

  // Render summary table
  var tableWrap = g('dash-kpi-table-wrap');
  if (tableWrap) {
    var rows = kpiData.map(function(d) {
      var pct = d.pct;
      var pctClass = gaugePillClass(pct);
      var forLabel = d.k.kpiFor === 'shop' ? '<span class="pill pill-blue">Shop</span>' : '<span class="pill pill-orange">Agent</span>';
      var assigneeName = d.assignee ? esc(d.assignee.name) : (d.k.assigneeBranch ? esc(d.k.assigneeBranch) : '—');
      var valueDisplay = d.k.valueMode === 'currency'
        ? fmtMoney(d.k.target, esc(d.k.currency) + ' ') + ' / ' + fmtMoney(d.actual, esc(d.k.currency) + ' ')
        : (function() { var ul = getKpiUnitLabel(d.k); return d.k.target + ' / ' + d.actual + (ul ? ' ' + esc(ul) : ''); })();
      return '<tr>' +
        '<td>' + esc(d.k.name) + '</td>' +
        '<td>' + forLabel + ' <small style="color:#888;">' + assigneeName + '</small></td>' +
        '<td>' + valueDisplay + '</td>' +
        '<td><span class="pill ' + pctClass + '">' + pct + '%</span></td>' +
        '</tr>';
    }).join('');
    tableWrap.innerHTML = '<table class="data-table" style="margin-top:8px;"><thead><tr><th>KPI</th><th>Assignee</th><th>Target / Actual</th><th>Achievement</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }
}
// ------------------------------------------------------------
function openCustomerModal(type, item) {
  if (type === 'new-customer') {
    const form = g('form-newCustomer');
    if (form) form.reset();
    g('nc-edit-id').value = '';
    const title = g('modal-newCustomer-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit New Customer';
      g('nc-edit-id').value = item.id;
      g('nc-name').value = item.name || '';
      g('nc-phone').value = item.phone || '';
      g('nc-id').value = item.idNum || '';
      const tariffSel = g('nc-tariff'); if (tariffSel) tariffSel.value = item.tariff || item.pkg || '';
      g('nc-agent').value = item.agent || '';
      const bSel = g('nc-branch'); if (bSel) bSel.value = item.branch || '';
      g('nc-date').value = item.date || '';
      const statusSel = g('nc-status'); if (statusSel) statusSel.value = item.status || 'follow';
      if (g('nc-lat')) g('nc-lat').value = item.lat || '';
      if (g('nc-lng')) g('nc-lng').value = item.lng || '';
    } else {
      if (title) title.textContent = 'Add New Customer';
      g('nc-date').value = new Date().toISOString().split('T')[0];
      if (currentUser) {
        const agEl = g('nc-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('nc-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-newCustomer');
    // Initialize Leaflet map after modal animation completes (150ms matches CSS transition)
    setTimeout(function() {
      var latVal = parseFloat(g('nc-lat') ? g('nc-lat').value : '') || 0;
      var lngVal = parseFloat(g('nc-lng') ? g('nc-lng').value : '') || 0;
      // Default center: Phnom Penh, Cambodia
      var defaultCenter = (latVal && lngVal) ? [latVal, lngVal] : [11.5564, 104.9282];
      if (window._ncMap) { window._ncMap.remove(); window._ncMap = null; }
      var map = L.map('nc-map', {
        fullscreenControl: true,
        fullscreenControlOptions: { position: 'topleft' }
      }).setView(defaultCenter, latVal && lngVal ? 15 : 12);
      var streetLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 19,
        attribution: '© OpenStreetMap contributors'
      });
      var satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
        maxZoom: 19,
        attribution: 'Tiles &copy; Esri'
      });
      satelliteLayer.addTo(map);
      L.control.layers({ 'Satellite': satelliteLayer, 'Street Map': streetLayer }, null, { position: 'topright' }).addTo(map);
      var marker = null;
      if (latVal && lngVal) {
        marker = L.marker([latVal, lngVal]).addTo(map);
      }
      map.on('click', function(e) {
        var lat = e.latlng.lat.toFixed(6); // 6 decimal places ≈ 0.11 m precision
        var lng = e.latlng.lng.toFixed(6);
        if (g('nc-lat')) g('nc-lat').value = lat;
        if (g('nc-lng')) g('nc-lng').value = lng;
        if (marker) { map.removeLayer(marker); }
        marker = L.marker([e.latlng.lat, e.latlng.lng]).addTo(map);
      });
      window._ncMap = map;
    }, 150);

  } else if (type === 'topup') {
    const form = g('form-topUp');
    if (form) form.reset();
    g('tu-edit-id').value = '';
    const title = g('modal-topUp-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit Top Up';
      g('tu-edit-id').value = item.id;
      g('tu-name').value = item.name || '';
      g('tu-phone').value = item.phone || '';
      g('tu-amount').value = item.amount || '';
      g('tu-agent').value = item.agent || '';
      const bSel = g('tu-branch'); if (bSel) bSel.value = item.branch || '';
      g('tu-date').value = item.date || '';
      const endDateEl = g('tu-end-date'); if (endDateEl) endDateEl.value = item.endDate || '';
      const tuStatusSel = g('tu-status'); if (tuStatusSel) tuStatusSel.value = item.tuStatus || 'active';
      const tuTariffEl = g('tu-tariff'); if (tuTariffEl) tuTariffEl.value = item.tariff || '';
      const tuRemarkEl = g('tu-remark'); if (tuRemarkEl) tuRemarkEl.value = item.remark || '';
      const tuLatEl = g('tu-lat'); if (tuLatEl) tuLatEl.value = item.lat || '';
      const tuLngEl = g('tu-lng'); if (tuLngEl) tuLngEl.value = item.lng || '';
      toggleTuLatlongRow(item.tuStatus || 'active', item.lat, item.lng);
    } else {
      if (title) title.textContent = 'Add Top Up';
      g('tu-date').value = new Date().toISOString().split('T')[0];
      toggleTuLatlongRow('active');
      if (currentUser) {
        const agEl = g('tu-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('tu-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-topUp');

  } else if (type === 'termination') {
    const form = g('form-termination');
    if (form) form.reset();
    g('term-edit-id').value = '';
    const title = g('modal-termination-title');
    populateBranchSelects();
    if (item) {
      if (title) title.textContent = 'Edit Termination';
      g('term-edit-id').value = item.id;
      g('term-name').value = item.name || '';
      g('term-phone').value = item.phone || '';
      g('term-reason').value = item.reason || '';
      g('term-agent').value = item.agent || '';
      const bSel = g('term-branch'); if (bSel) bSel.value = item.branch || '';
      g('term-date').value = item.date || '';
      const termLatEl = g('term-lat'); if (termLatEl) termLatEl.value = item.lat || '';
      const termLngEl = g('term-lng'); if (termLngEl) termLngEl.value = item.lng || '';
    } else {
      if (title) title.textContent = 'Add Termination';
      g('term-date').value = new Date().toISOString().split('T')[0];
      if (currentUser) {
        const agEl = g('term-agent'); if (agEl) { agEl.value = currentUser.name || ''; if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true; }
        const brEl = g('term-branch'); if (brEl && currentUser.branch) { brEl.value = currentUser.branch; if (currentRole === 'agent' || currentRole === 'supervisor') brEl.disabled = true; }
      }
    }
    openModal('modal-termination');
  }
}

function submitNewCustomer(e) {
  e.preventDefault();
  const editId = rv('nc-edit-id');
  const tariffEl = g('nc-tariff');
  const statusEl = g('nc-status');
  const obj = {
    id: editId || uid(),
    name: rv('nc-name'), phone: rv('nc-phone'), idNum: rv('nc-id'),
    tariff: tariffEl ? tariffEl.value : '', pkg: tariffEl ? tariffEl.value : '',
    agent: rv('nc-agent'), branch: rv('nc-branch'), date: rv('nc-date'),
    status: statusEl ? statusEl.value : 'follow',
    lat: rv('nc-lat') || '', lng: rv('nc-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  const prevStatus = editId ? (newCustomers.find(function(x){return x.id===editId;})||{}).status : null;
  if (editId) {
    const idx = newCustomers.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) newCustomers[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a customer record.');
  } else {
    newCustomers.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a new customer.');
  }
  // Auto-add to TopUp when status is Close and remove from New Customer list
  if (obj.status === 'close' && prevStatus !== 'close') {
    var existingTopUpRecord = topUpList.find(function(t) { return t.customerId === obj.id; });
    if (!existingTopUpRecord) {
      topUpList.push({
        id: uid(), customerId: obj.id, name: obj.name, phone: obj.phone,
        tariff: obj.tariff, agent: obj.agent, branch: obj.branch, date: obj.date,
        tuStatus: 'active', amount: 0, note: 'Auto-added (status: Close)'
      });
      syncSheet('TopUp', topUpList);
    }
    // Remove from new customer list and navigate to Top Up tab
    newCustomers = newCustomers.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-newCustomer');
    renderNewCustomerTable();
    renderTopUpTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('topup');
    // Update tab button active state
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var topupBtn = g('tab-topup'); if (topupBtn) topupBtn.classList.add('active');
    showToast('Customer marked as Closed and moved to Top Up.', 'info');
    return;
  }
  // Move to Out Coverage List when status is Out Coverage
  if (obj.status === 'out-coverage' && prevStatus !== 'out-coverage') {
    var existingOutCovRecord = outCoverageList.find(function(t) { return t.customerId === obj.id; });
    if (!existingOutCovRecord) {
      outCoverageList.push({
        id: uid(), customerId: obj.id, name: obj.name, phone: obj.phone,
        idNum: obj.idNum, tariff: obj.tariff, agent: obj.agent, branch: obj.branch,
        date: obj.date, lat: obj.lat, lng: obj.lng, note: 'Out of coverage area'
      });
      syncSheet('OutCoverage', outCoverageList);
    }
    // Remove from new customer list and navigate to Out Coverage tab
    newCustomers = newCustomers.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-newCustomer');
    renderNewCustomerTable();
    renderOutCoverageTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('out-coverage');
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var ocBtn = g('tab-out-coverage'); if (ocBtn) ocBtn.classList.add('active');
    showToast('Customer marked as Out Coverage and moved to Out Coverage List.', 'info');
    return;
  }
  closeModal('modal-newCustomer');
  renderNewCustomerTable();
  renderTopUpTable();
  syncSheet('Customers', newCustomers);
  saveAllData();
  showToast(editId ? 'Customer record updated.' : 'Customer added successfully.', 'success');
}

function editNewCustomer(id) {
  const item = newCustomers.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('new-customer', item);
}

function deleteNewCustomer(id) {
  const item = newCustomers.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this customer record? This action cannot be undone.', function() {
    newCustomers = newCustomers.filter(function(x) { return x.id !== id; });
    renderNewCustomerTable();
    syncSheet('Customers', newCustomers);
    saveAllData();
    showToast('Customer record deleted.', 'success');
  }, 'Delete Customer', 'Delete');
}

function renderNewCustomerTable() {
  const tbody = g('new-customer-table');
  if (!tbody) return;
  const searchVal = (rv('nc-search') || '').toLowerCase().trim();
  const baseList = getBaseRecordsForRole(newCustomers);
  const list = searchVal
    ? baseList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(searchVal) ||
               (c.phone || '').toLowerCase().includes(searchVal);
      })
    : baseList;
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-users" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' + (searchVal ? 'No results found' : 'No customers yet') + '</td></tr>';
    return;
  }
  const statusPillMap = { follow: 'pill-gray', lead: 'pill-blue', 'hot-prospect': 'pill-orange', close: 'pill-green', 'out-coverage': 'pill-red' };
  const statusLabelMap = { follow: 'Follow', lead: 'Lead', 'hot-prospect': 'Hot Prospect', close: 'Close', 'out-coverage': 'Out Coverage' };
  tbody.innerHTML = list.map(function(c, i) {
    const avIdx = i % 8;
    const st = c.status || 'follow';
    const stPill = statusPillMap[st] || 'pill-gray';
    const stLabel = statusLabelMap[st] || esc(st);
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.idNum || '') + '</td>' +
      '<td>' + esc(c.tariff || c.pkg || '') + '</td>' +
      '<td><span class="pill ' + stPill + '">' + stLabel + '</span></td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td style="white-space:nowrap;">' + (c.lat && c.lng
        ? '<a href="https://www.openstreetmap.org/?mlat=' + esc(c.lat) + '&mlon=' + esc(c.lng) + '&zoom=15" target="_blank" title="' + esc(c.lat) + ', ' + esc(c.lng) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-map-marker-alt"></i> ' + esc(c.lat) + ', ' + esc(c.lng) + '</a>'
        : '<span style="color:#ccc;font-size:0.8rem;">—</span>') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editNewCustomer(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteNewCustomer(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function toggleTuLatlongRow(status, existingLat, existingLng) {
  const row = g('tu-latlong-row');
  if (!row) return;
  if (status === 'terminate') {
    row.style.display = '';
    setTimeout(function() {
      if (typeof L === 'undefined') return;
      if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
      var pickerEl = g('tu-picker-map');
      if (!pickerEl) return;
      var lat0 = parseFloat(existingLat) || null;
      var lng0 = parseFloat(existingLng) || null;
      var center = (lat0 && lng0) ? [lat0, lng0] : KH_CENTER;
      var zoom = (lat0 && lng0) ? 14 : 7;
      _tuPickerMap = L.map('tu-picker-map').setView(center, zoom);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        maxZoom: 18,
        attribution: '© OpenStreetMap contributors'
      }).addTo(_tuPickerMap);
      if (lat0 && lng0) {
        _tuPickerMarker = L.marker([lat0, lng0]).addTo(_tuPickerMap);
      }
      _tuPickerMap.on('click', function(e) {
        var lat = parseFloat(e.latlng.lat.toFixed(6));
        var lng = parseFloat(e.latlng.lng.toFixed(6));
        var latEl = g('tu-lat'); var lngEl = g('tu-lng');
        if (latEl) latEl.value = lat;
        if (lngEl) lngEl.value = lng;
        if (_tuPickerMarker) _tuPickerMap.removeLayer(_tuPickerMarker);
        _tuPickerMarker = L.marker([lat, lng]).addTo(_tuPickerMap);
      });
    }, 200);
  } else {
    row.style.display = 'none';
    if (_tuPickerMap) { _tuPickerMap.remove(); _tuPickerMap = null; _tuPickerMarker = null; }
    var latEl = g('tu-lat'); var lngEl = g('tu-lng');
    if (latEl) latEl.value = '';
    if (lngEl) lngEl.value = '';
  }
}

function submitTopUp(e) {
  e.preventDefault();
  const editId = rv('tu-edit-id');
  const tuStatusSel = g('tu-status');
  const tuStatus = tuStatusSel ? tuStatusSel.value : 'active';
  const existingRecord = editId ? topUpList.find(function(x) { return x.id === editId; }) : null;
  const obj = {
    id: editId || uid(),
    customerId: existingRecord ? (existingRecord.customerId || '') : '',
    name: rv('tu-name'), phone: rv('tu-phone'), amount: parseFloat(rv('tu-amount')) || 0,
    agent: rv('tu-agent'), branch: rv('tu-branch'), date: rv('tu-date'),
    endDate: rv('tu-end-date') || '',
    tariff: rv('tu-tariff') || '',
    remark: rv('tu-remark') || '',
    tuStatus: tuStatus,
    lat: rv('tu-lat') || '', lng: rv('tu-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  const prevStatus = existingRecord ? existingRecord.tuStatus : null;
  if (editId) {
    const idx = topUpList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) topUpList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a top-up record.');
  } else {
    topUpList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' submitted a top-up.');
  }
  // Auto-add to termination when status = Terminate and remove from Top Up list
  if (tuStatus === 'terminate' && prevStatus !== 'terminate') {
    const existingTerminationRecord = terminationList.find(function(t) { return (obj.customerId && t.customerId === obj.customerId) || (t.name === obj.name && t.phone === obj.phone); });
    if (!existingTerminationRecord) {
      terminationList.push({
        id: uid(), customerId: obj.customerId || '', name: obj.name, phone: obj.phone, reason: 'Service terminated',
        agent: obj.agent, branch: obj.branch, date: obj.date, lat: obj.lat || '', lng: obj.lng || ''
      });
      syncSheet('Terminations', terminationList);
      renderTerminationTable();
    }
    // Remove from Top Up list and navigate to Termination tab
    topUpList = topUpList.filter(function(x) { return x.id !== obj.id; });
    closeModal('modal-topUp');
    renderTopUpTable();
    syncSheet('TopUp', topUpList);
    saveAllData();
    navigateTo('customer', null);
    switchCustomerTab('termination');
    // Update tab button active state
    $$('.tab-btn').forEach(function(b) { b.classList.remove('active'); });
    var termBtn = g('tab-termination'); if (termBtn) termBtn.classList.add('active');
    showToast('Customer terminated and moved to Termination list.', 'info');
    return;
  }
  closeModal('modal-topUp');
  renderTopUpTable();
  syncSheet('TopUp', topUpList);
  saveAllData();
  showToast(editId ? 'Top-up record updated.' : 'Top-up submitted successfully.', 'success');
}

function onTuExistingCustomerChange() {
  const sel = g('tu-existing-customer');
  if (!sel || !sel.value) return;
  const customer = newCustomers.find(function(c) { return c.id === sel.value; });
  if (!customer) return;
  const nameEl = g('tu-name'); if (nameEl) nameEl.value = customer.name || '';
  const phoneEl = g('tu-phone'); if (phoneEl) phoneEl.value = customer.phone || '';
  const agEl = g('tu-agent'); if (agEl && !agEl.value) agEl.value = customer.agent || '';
  const brEl = g('tu-branch'); if (brEl && !brEl.value) brEl.value = customer.branch || '';
}

function editTopUp(id) {
  const item = topUpList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('topup', item);
}

function deleteTopUp(id) {
  const item = topUpList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this top-up record? This action cannot be undone.', function() {
    topUpList = topUpList.filter(function(x) { return x.id !== id; });
    renderTopUpTable();
    syncSheet('TopUp', topUpList);
    saveAllData();
    showToast('Top-up record deleted.', 'success');
  }, 'Delete Top-Up Record', 'Delete');
}

function renderTopUpTable() {
  const tbody = g('topup-table');
  if (!tbody) return;

  // Render nearly-expired notice
  const MS_PER_DAY = 86400000;
  const today = new Date(new Date().toISOString().split('T')[0]);
  const in7Days = new Date(today); in7Days.setDate(in7Days.getDate() + 7);
  const baseTopUpList = getBaseRecordsForRole(topUpList);
  const nearlyExpired = baseTopUpList.filter(function(c) {
    if (!c.endDate || c.tuStatus === 'terminate') return false;
    const exp = new Date(c.endDate);
    return exp >= today && exp <= in7Days;
  });
  const noticeEl = g('topup-expiry-notice');
  if (noticeEl) {
    if (nearlyExpired.length) {
      noticeEl.innerHTML = '<div class="expiry-notice-banner">' +
        '<div class="expiry-notice-icon"><i class="fas fa-bell"></i></div>' +
        '<div class="expiry-notice-content">' +
          '<div class="expiry-notice-title">Follow-up Required — ' + nearlyExpired.length + ' customer' + (nearlyExpired.length > 1 ? 's' : '') + ' expiring within 7 days</div>' +
          '<div class="expiry-notice-list">' +
            nearlyExpired.map(function(c) {
              const daysLeft = Math.round((new Date(c.endDate) - today) / MS_PER_DAY);
              return '<span class="expiry-notice-item">' + esc(c.name) + ' (' + esc(c.phone) + ') — <strong>' + (daysLeft === 0 ? 'Today' : daysLeft + ' day' + (daysLeft > 1 ? 's' : '')) + '</strong></span>';
            }).join('') +
          '</div>' +
        '</div>' +
      '</div>';
    } else {
      noticeEl.innerHTML = '';
    }
  }

  if (!baseTopUpList.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-coins" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No top up records yet</td></tr>';
    return;
  }
  const tuSearchVal = (rv('tu-search') || '').toLowerCase().trim();
  const tuList = tuSearchVal
    ? baseTopUpList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(tuSearchVal) ||
               (c.phone || '').toLowerCase().includes(tuSearchVal);
      })
    : baseTopUpList;
  if (!tuList.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-coins" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No results found</td></tr>';
    return;
  }
  tbody.innerHTML = tuList.map(function(c, i) {
    const avIdx = i % 8;
    const tuSt = c.tuStatus || 'active';
    const stPill = tuSt === 'active' ? 'pill-green' : 'pill-red';
    const stLabel = tuSt === 'active' ? 'Active' : 'Terminate';
    const canEdit = canModifyRecord(c);
    var expiryCell = '<td>—</td>';
    var rowClass = '';
    if (c.endDate) {
      const exp = new Date(c.endDate);
      const daysLeft = Math.round((exp - today) / MS_PER_DAY);
      if (tuSt !== 'terminate' && daysLeft >= 0 && daysLeft <= 7) {
        rowClass = ' class="tr-nearly-expired"';
        expiryCell = '<td><span class="expiry-badge expiry-badge-warn">' + esc(c.endDate) + ' <span class="expiry-days-left">(' + (daysLeft === 0 ? 'Today' : daysLeft + 'd') + ')</span></span></td>';
      } else if (tuSt !== 'terminate' && daysLeft < 0) {
        rowClass = ' class="tr-expired"';
        expiryCell = '<td><span class="expiry-badge expiry-badge-expired">' + esc(c.endDate) + ' <span class="expiry-days-left">(Expired)</span></span></td>';
      } else {
        expiryCell = '<td>' + esc(c.endDate) + '</td>';
      }
    }
    return '<tr' + rowClass + '>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + fmtMoney(c.amount) + '</td>' +
      '<td><span class="pill ' + stPill + '">' + stLabel + '</span></td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      expiryCell +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editTopUp(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteTopUp(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function submitTermination(e) {
  e.preventDefault();
  const editId = rv('term-edit-id');
  const obj = {
    id: editId || uid(),
    name: rv('term-name'), phone: rv('term-phone'), reason: rv('term-reason'),
    agent: rv('term-agent'), branch: rv('term-branch'), date: rv('term-date'),
    lat: rv('term-lat') || '', lng: rv('term-lng') || ''
  };
  if (!obj.name) { showAlert('Please enter customer name'); return; }
  if (!obj.phone) { showAlert('Please enter phone number'); return; }
  if (!/^\d{6,15}$/.test(obj.phone.replace(/[\s\-+()]/g, ''))) { showAlert('Please enter a valid phone number (6–15 digits, separators allowed)'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  if (editId) {
    const idx = terminationList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) terminationList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a termination record.');
  } else {
    terminationList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' submitted a termination.');
  }
  closeModal('modal-termination');
  renderTerminationTable();
  syncSheet('Terminations', terminationList);
  saveAllData();
  showToast(editId ? 'Termination record updated.' : 'Termination submitted successfully.', 'success');
}

function editTermination(id) {
  const item = terminationList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openCustomerModal('termination', item);
}

function deleteTermination(id) {
  const item = terminationList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this termination record? This action cannot be undone.', function() {
    terminationList = terminationList.filter(function(x) { return x.id !== id; });
    renderTerminationTable();
    syncSheet('Terminations', terminationList);
    saveAllData();
    showToast('Termination record deleted.', 'success');
  }, 'Delete Termination Record', 'Delete');
}

function renderTerminationTable() {
  const tbody = g('termination-table');
  if (!tbody) return;
  const baseTermList = getBaseRecordsForRole(terminationList);
  if (!baseTermList.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-times-circle" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No termination records yet</td></tr>';
    return;
  }
  const termSearchVal = (rv('term-search') || '').toLowerCase().trim();
  const termList = termSearchVal
    ? baseTermList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(termSearchVal) ||
               (c.phone || '').toLowerCase().includes(termSearchVal);
      })
    : baseTermList;
  if (!termList.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-times-circle" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No results found</td></tr>';
    return;
  }
  tbody.innerHTML = termList.map(function(c, i) {
    const avIdx = i % 8;
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.reason || '') + '</td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td>' + (c.lat || c.lng ? [c.lat, c.lng].filter(Boolean).map(esc).join(', ') : '—') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-edit" onclick="editTermination(\'' + esc(c.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteTermination(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Out Coverage Functions
// ------------------------------------------------------------
function renderOutCoverageTable() {
  const tbody = g('out-coverage-table');
  if (!tbody) return;
  const baseList = getBaseRecordsForRole(outCoverageList);
  const searchVal = (rv('oc-search') || '').toLowerCase().trim();
  const list = searchVal
    ? baseList.filter(function(c) {
        return (c.name || '').toLowerCase().includes(searchVal) ||
               (c.phone || '').toLowerCase().includes(searchVal);
      })
    : baseList;
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-signal" style="font-size:2rem;display:block;margin-bottom:8px;"></i>' + (searchVal ? 'No results found' : 'No out of coverage records yet') + '</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(function(c, i) {
    const avIdx = i % 8;
    const canEdit = canModifyRecord(c);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(c.name)) + '</span>' + esc(c.name) + '</div></td>' +
      '<td>' + esc(c.phone) + '</td>' +
      '<td>' + esc(c.idNum || '') + '</td>' +
      '<td>' + esc(c.tariff || '') + '</td>' +
      '<td>' + esc(c.agent || '') + '</td>' +
      '<td>' + esc(c.branch || '') + '</td>' +
      '<td>' + esc(c.date || '') + '</td>' +
      '<td style="white-space:nowrap;">' + (c.lat && c.lng
        ? '<a href="https://www.openstreetmap.org/?mlat=' + esc(c.lat) + '&mlon=' + esc(c.lng) + '&zoom=15" target="_blank" title="' + esc(c.lat) + ', ' + esc(c.lng) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-map-marker-alt"></i> ' + esc(c.lat) + ', ' + esc(c.lng) + '</a>'
        : '<span style="color:#ccc;font-size:0.8rem;">—</span>') + '</td>' +
      '<td style="white-space:nowrap;">' +
        (canEdit ? '<button class="btn-delete" onclick="deleteOutCoverage(\'' + esc(c.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function deleteOutCoverage(id) {
  const item = outCoverageList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Delete this out coverage record?', function() {
    outCoverageList = outCoverageList.filter(function(x) { return x.id !== id; });
    syncSheet('OutCoverage', outCoverageList);
    saveAllData();
    renderOutCoverageTable();
    showToast('Out coverage record deleted.', 'success');
  });
}

// ------------------------------------------------------------
// Promotion Functions
// ------------------------------------------------------------
function isPromoExpired(p) {
  if (!p.endDate) return false;
  const today = new Date(new Date().toISOString().split('T')[0]);
  return new Date(p.endDate) < today;
}

function openNewPromotionModal(item) {
  const form = g('form-newPromotion');
  if (form) form.reset();
  const editEl = g('np-edit-id');
  if (editEl) editEl.value = '';
  const title = g('modal-newPromotion-title');
  const btn = g('np-submit-btn');
  if (item) {
    if (title) title.textContent = 'Edit Promotion';
    if (btn) btn.textContent = 'Update Promotion';
    if (editEl) editEl.value = item.id;
    const c = g('np-campaign'); if (c) c.value = item.campaign || '';
    const ch = g('np-channel'); if (ch) ch.value = item.channel || '';
    const s = g('np-start'); if (s) s.value = item.startDate || '';
    const e = g('np-end'); if (e) e.value = item.endDate || '';
    const t = g('np-terms'); if (t) t.value = item.terms || '';
  } else {
    if (title) title.textContent = 'New Promotion';
    if (btn) btn.textContent = 'Add Promotion';
    const s = g('np-start'); if (s) s.value = new Date().toISOString().split('T')[0];
  }
  openModal('modal-newPromotion');
}

function submitNewPromotion(e) {
  e.preventDefault();
  const editId = rv('np-edit-id');
  const obj = {
    id: editId || uid(),
    campaign: rv('np-campaign'),
    channel: rv('np-channel'),
    startDate: rv('np-start'),
    endDate: rv('np-end'),
    terms: g('np-terms') ? g('np-terms').value : ''
  };
  if (!obj.campaign) { showAlert('Please enter campaign name'); return; }
  if (!obj.startDate || !obj.endDate) { showAlert('Please enter the promotion period'); return; }
  if (editId) {
    const idx = promotionList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) promotionList[idx] = obj;
  } else {
    promotionList.push(obj);
  }
  closeModal('modal-newPromotion');
  renderPromotionCards();
  renderPromoSettingTable();
  syncSheet('Promotions', promotionList);
  saveAllData();
  showToast(editId ? 'Promotion updated.' : 'Promotion created successfully.', 'success');
}

function editNewPromotion(id) {
  const item = promotionList.find(function(x) { return x.id === id; });
  if (item) openNewPromotionModal(item);
}

function deleteNewPromotion(id) {
  showConfirm('Are you sure you want to delete this promotion? This action cannot be undone.', function() {
    promotionList = promotionList.filter(function(x) { return x.id !== id; });
    renderPromotionCards();
    renderPromoSettingTable();
    syncSheet('Promotions', promotionList);
    saveAllData();
    showToast('Promotion deleted.', 'success');
  }, 'Delete Promotion', 'Delete');
}

function renderPromotionCards() {
  const available = promotionList.filter(function(p) { return !isPromoExpired(p); });
  const expired = promotionList.filter(function(p) { return isPromoExpired(p); });

  const avCount = g('promo-available-count');
  const exCount = g('promo-expired-count');
  if (avCount) avCount.textContent = available.length;
  if (exCount) exCount.textContent = expired.length;

  const avGrid = g('promo-available-grid');
  const avEmpty = g('promo-available-empty');
  if (avGrid) {
    if (!available.length) {
      avGrid.innerHTML = '';
      if (avEmpty) avEmpty.style.display = '';
    } else {
      if (avEmpty) avEmpty.style.display = 'none';
      avGrid.innerHTML = available.map(function(p) { return buildPromoCard(p, false); }).join('');
    }
  }

  const exGrid = g('promo-expired-grid');
  const exEmpty = g('promo-expired-empty');
  if (exGrid) {
    if (!expired.length) {
      exGrid.innerHTML = '';
      if (exEmpty) exEmpty.style.display = '';
    } else {
      if (exEmpty) exEmpty.style.display = 'none';
      exGrid.innerHTML = expired.map(function(p) { return buildPromoCard(p, true); }).join('');
    }
  }
  var newBtn = g('promo-new-btn');
  if (newBtn) newBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';
  // Apply current promo view to show/hide sections
  setPromoView(currentPromoView);
}

function buildPromoCard(p, isExpired) {
  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');
  var statusBadge = isExpired
    ? '<span class="promo-status-badge promo-status-expired">Expired</span>'
    : '<span class="promo-status-badge promo-status-active">Active</span>';
  var termsHtml = p.terms ? '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-file-lines"></i></span><span class="promo-info-value promo-terms-text">' + esc(p.terms) + '</span></div>' : '';
  var actionsHtml;
  if (isAdmin) {
    if (isExpired) {
      actionsHtml = '<button class="btn-restore" onclick="restorePromotion(\'' + esc(p.id) + '\')"><i class="fas fa-rotate-left"></i> Restore</button>' +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>';
    } else {
      actionsHtml = '<button class="btn-edit" onclick="editNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-edit"></i> Edit</button>' +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>';
    }
  } else {
    actionsHtml = '<button class="btn-view-promo" onclick="openPromoViewModal(\'' + esc(p.id) + '\')"><i class="fas fa-eye"></i> View</button>';
  }
  return '<div class="promo-card-v2' + (isExpired ? ' expired' : '') + '">' +
    '<div class="promo-card-v2-header">' +
      '<span class="promo-card-v2-title">' + esc(p.campaign) + '</span>' +
      statusBadge +
    '</div>' +
    '<div class="promo-card-v2-body">' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-broadcast-tower"></i></span><span><span class="promo-info-label">Channel: </span><span class="promo-info-value">' + esc(p.channel || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar-range"></i></span><span><span class="promo-info-label">Period: </span><span class="promo-info-value">' + esc(p.startDate || '') + ' \u2192 ' + esc(p.endDate || '') + '</span></span></div>' +
      termsHtml +
    '</div>' +
    '<div class="promo-card-v2-footer">' + actionsHtml + '</div>' +
  '</div>';
}

function restorePromotion(id) {
  var idx = promotionList.findIndex(function(x) { return x.id === id; });
  if (idx < 0) return;
  var today = new Date();
  var future = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 30);
  promotionList[idx].endDate = future.toISOString().split('T')[0];
  renderPromotionCards();
  showToast('Promotion restored for 30 days.', 'success');
  syncSheet('Promotions', promotionList);
  saveAllData();
}

function openPromoViewModal(id) {
  var p = promotionList.find(function(x) { return x.id === id; });
  if (!p) return;
  var body = g('modal-viewPromo-body');
  if (body) {
    body.innerHTML =
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-tag"></i></span><span><span class="promo-info-label">Campaign: </span><span class="promo-info-value">' + esc(p.campaign) + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-broadcast-tower"></i></span><span><span class="promo-info-label">Channel: </span><span class="promo-info-value">' + esc(p.channel || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar"></i></span><span><span class="promo-info-label">Start: </span><span class="promo-info-value">' + esc(p.startDate || '—') + '</span></span></div>' +
      '<div class="promo-info-row"><span class="promo-info-icon"><i class="fas fa-calendar-xmark"></i></span><span><span class="promo-info-label">End: </span><span class="promo-info-value">' + esc(p.endDate || '—') + '</span></span></div>' +
      (p.terms ? '<div class="promo-info-row" style="margin-top:8px;"><span class="promo-info-icon"><i class="fas fa-file-lines"></i></span><span><span class="promo-info-label">Terms: </span><span class="promo-info-value" style="font-style:italic;color:#777;">' + esc(p.terms) + '</span></span></div>' : '');
  }
  openModal('modal-viewPromo');
}

function renderPromotionTable() {
  renderPromotionCards();
}

function renderPromoSettingTable() {
  const tbody = g('promo-setting-table') ? g('promo-setting-table').querySelector('tbody') : null;
  if (!tbody) return;
  if (!promotionList.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:30px;color:#999;">No promotions defined</td></tr>';
    return;
  }
  tbody.innerHTML = promotionList.map(function(p, i) {
    const expired = isPromoExpired(p);
    const statusPill = expired ? 'pill-gray' : 'pill-green';
    const statusLabel = expired ? 'Expired' : 'Active';
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><strong>' + esc(p.campaign || p.name || '') + '</strong></td>' +
      '<td>' + esc(p.channel || '') + '</td>' +
      '<td>' + esc(p.startDate || '') + ' – ' + esc(p.endDate || '') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + statusLabel + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        (expired ? '' : '<button class="btn-edit" onclick="editNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-edit"></i></button> ') +
        '<button class="btn-delete" onclick="deleteNewPromotion(\'' + esc(p.id) + '\')"><i class="fas fa-trash"></i></button>' +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Deposit Functions
// ------------------------------------------------------------
var USD_DENOMS = [100, 50, 20, 10, 5, 1];
var KHR_DENOMS = [100000, 50000, 20000, 10000, 5000, 2000, 1000, 500, 100];

function formatKHR(n) {
  return (n || 0).toLocaleString('en-US') + ' ៛';
}

function toggleDenomSection() {
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (!section) return;
  var isOpen = section.style.display !== 'none';
  section.style.display = isOpen ? 'none' : 'block';
  if (btn) btn.classList.toggle('open', !isOpen);
}

function calcDenomTotals() {
  var usdTotal = 0;
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    var sub = qty * d;
    usdTotal += sub;
    var subEl = g('usd-sub-' + d);
    if (subEl) {
      subEl.textContent = qty > 0 ? '$' + sub.toFixed(2) : '—';
      subEl.classList.toggle('has-value', qty > 0);
    }
  });
  var usdHeaderEl = g('usd-denom-total');
  if (usdHeaderEl) usdHeaderEl.textContent = '$' + usdTotal.toFixed(2);
  var usdRowEl = g('usd-denom-total-row');
  if (usdRowEl) usdRowEl.textContent = '$' + usdTotal.toFixed(2);
  var cashEl = g('dep-cash'); if (cashEl) cashEl.value = usdTotal > 0 ? usdTotal.toFixed(2) : '';

  var khrTotal = 0;
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    var sub = qty * d;
    khrTotal += sub;
    var subEl = g('khr-sub-' + d);
    if (subEl) {
      subEl.textContent = qty > 0 ? formatKHR(sub) : '—';
      subEl.classList.toggle('has-value', qty > 0);
    }
  });
  var khrHeaderEl = g('khr-denom-total');
  if (khrHeaderEl) khrHeaderEl.textContent = formatKHR(khrTotal);
  var khrRowEl = g('khr-denom-total-row');
  if (khrRowEl) khrRowEl.textContent = formatKHR(khrTotal);
  var rielEl = g('dep-riel'); if (rielEl) rielEl.value = khrTotal > 0 ? khrTotal : '';
}

function _resetDenomSection() {
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (section) section.style.display = 'none';
  if (btn) btn.classList.remove('open');
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d); if (el) el.value = '';
    var subEl = g('usd-sub-' + d); if (subEl) { subEl.textContent = '—'; subEl.classList.remove('has-value'); }
  });
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d); if (el) el.value = '';
    var subEl = g('khr-sub-' + d); if (subEl) { subEl.textContent = '—'; subEl.classList.remove('has-value'); }
  });
  var usdH = g('usd-denom-total'); if (usdH) usdH.textContent = '$0.00';
  var usdR = g('usd-denom-total-row'); if (usdR) usdR.textContent = '$0.00';
  var khrH = g('khr-denom-total'); if (khrH) khrH.textContent = '0 ៛';
  var khrR = g('khr-denom-total-row'); if (khrR) khrR.textContent = '0 ៛';
}

function _loadDenomSection(cashDetail) {
  if (!cashDetail) return;
  var section = g('denom-section');
  var btn = g('denom-toggle-btn');
  if (section) section.style.display = 'block';
  if (btn) btn.classList.add('open');
  (cashDetail.usd || []).forEach(function(entry) {
    var el = g('usd-qty-' + entry.denom); if (el) el.value = entry.qty;
  });
  (cashDetail.khr || []).forEach(function(entry) {
    var el = g('khr-qty-' + entry.denom); if (el) el.value = entry.qty;
  });
  calcDenomTotals();
}

function openAddDeposit(el) {
  navigateTo('deposit', null);
  setActiveSubItem(el);
  openDepositModal(null);
}

function openDepositModal(item) {
  const form = g('form-addDeposit');
  if (form) form.reset();
  const editEl = g('dep-edit-id');
  if (editEl) editEl.value = '';
  _resetDenomSection();

  // Always reset field lock states so they don't persist between modal opens
  const agEl = g('dep-agent');
  const brEl = g('dep-branch');
  const brTextEl = g('dep-branch-text');
  if (agEl) agEl.readOnly = false;
  if (brEl) { brEl.disabled = false; brEl.style.display = ''; brEl.required = true; }
  if (brTextEl) brTextEl.style.display = 'none';

  const title = g('modal-addDeposit-title');
  const btn = g('dep-submit-btn');
  populateBranchSelects();

  if (item) {
    if (title) title.textContent = 'Edit Deposit';
    if (btn) btn.textContent = 'Update Deposit';
    if (editEl) editEl.value = item.id;
    if (agEl) agEl.value = item.agent || '';
    if (brEl) brEl.value = item.branch || '';
    const cashEl = g('dep-cash'); if (cashEl) cashEl.value = item.cash || '';
    const creditEl = g('dep-credit'); if (creditEl) creditEl.value = item.credit || '';
    const rielEl = g('dep-riel'); if (rielEl) rielEl.value = item.riel || '';
    const dtEl = g('dep-date'); if (dtEl) dtEl.value = item.date || '';
    const ntEl = g('dep-remark'); if (ntEl) ntEl.value = item.remark || item.note || '';
    if (item.cashDetail) _loadDenomSection(item.cashDetail);
    if (currentUser && (currentRole === 'agent' || currentRole === 'supervisor')) {
      if (agEl) agEl.readOnly = true;
      if (brEl && brTextEl) {
        brTextEl.textContent = item.branch || '';
        brEl.style.display = 'none';
        brEl.required = false;
        brTextEl.style.display = '';
      }
    }
  } else {
    if (title) title.textContent = 'Add Deposit';
    if (btn) btn.textContent = 'Add Deposit';
    const dtEl = g('dep-date'); if (dtEl) dtEl.value = new Date().toISOString().split('T')[0];
    if (currentUser) {
      if (agEl) {
        agEl.value = currentUser.name || '';
        if (currentRole === 'agent' || currentRole === 'supervisor') agEl.readOnly = true;
      }
      if (brEl) {
        if (currentUser.branch) brEl.value = currentUser.branch;
        if (currentRole === 'agent' || currentRole === 'supervisor') {
          if (brTextEl) {
            brTextEl.textContent = currentUser.branch || '';
            brTextEl.style.display = '';
          }
          brEl.style.display = 'none';
          brEl.required = false;
        }
      }
    }
  }
  openModal('modal-addDeposit');
}

function submitDeposit(e) {
  e.preventDefault();
  const editId = rv('dep-edit-id');
  const cash = parseFloat(rv('dep-cash')) || 0;
  const credit = parseFloat(rv('dep-credit')) || 0;
  const riel = parseFloat(rv('dep-riel')) || 0;
  const cashDetailUsd = [];
  const cashDetailKhr = [];
  USD_DENOMS.forEach(function(d) {
    var el = g('usd-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    if (qty > 0) cashDetailUsd.push({ denom: d, qty: qty });
  });
  KHR_DENOMS.forEach(function(d) {
    var el = g('khr-qty-' + d);
    var qty = el ? (parseInt(el.value) || 0) : 0;
    if (qty > 0) cashDetailKhr.push({ denom: d, qty: qty });
  });
  const obj = {
    id: editId || uid(),
    agent: rv('dep-agent'),
    branch: rv('dep-branch'),
    cash: cash,
    credit: credit,
    riel: riel,
    cashDetail: (cashDetailUsd.length || cashDetailKhr.length) ? { usd: cashDetailUsd, khr: cashDetailKhr } : null,
    amount: cash + credit,
    date: rv('dep-date'),
    remark: rv('dep-remark'),
    status: editId ? (depositList.find(function(x){return x.id===editId;})||{}).status || 'pending' : 'pending'
  };
  if (!obj.agent) { showAlert('Please enter agent name'); return; }
  if (obj.amount <= 0 && riel <= 0) { showAlert('Please enter a cash, credit, and/or KHR riel amount greater than zero'); return; }
  if (!obj.date) { showAlert('Please select a date'); return; }
  if (editId) {
    const idx = depositList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) depositList[idx] = obj;
    addNotification((currentUser ? currentUser.name : 'User') + ' updated a deposit record.');
  } else {
    depositList.push(obj);
    addNotification((currentUser ? currentUser.name : 'User') + ' added a deposit.');
  }
  closeModal('modal-addDeposit');
  renderDepositTable();
  updateDepositKpis();
  syncSheet('Deposits', depositList);
  saveAllData();
  showToast(editId ? 'Deposit record updated.' : 'Deposit submitted successfully.', 'success');
}

function editDeposit(id) {
  const item = depositList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to edit this record.', 'error'); return; }
  openDepositModal(item);
}

function deleteDeposit(id) {
  const item = depositList.find(function(x) { return x.id === id; });
  if (!item) return;
  if (!canModifyRecord(item)) { showAlert('You do not have permission to delete this record.', 'error'); return; }
  showConfirm('Are you sure you want to delete this deposit record? This action cannot be undone.', function() {
    depositList = depositList.filter(function(x) { return x.id !== id; });
    renderDepositTable();
    updateDepositKpis();
    syncSheet('Deposits', depositList);
    saveAllData();
    showToast('Deposit record deleted.', 'success');
  }, 'Delete Deposit Record', 'Delete');
}

function approveDeposit(id) {
  var canApprove = (currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster');
  if (!canApprove) { showAlert('Only supervisor, admin, or cluster can approve deposits.', 'warning'); return; }
  showConfirm('Approve this deposit record?', function() {
    var idx = depositList.findIndex(function(x) { return x.id === id; });
    if (idx >= 0) {
      depositList[idx].status = 'approved';
      depositList[idx].approvedBy = currentUser ? currentUser.name : 'Supervisor';
      depositList[idx].approvedAt = new Date().toISOString().split('T')[0];
      renderDepositTable();
      updateDepositKpis();
      syncSheet('Deposits', depositList);
      saveAllData();
      showToast('Deposit approved.', 'success');
      showApprovalFormModal('deposit', depositList[idx]);
    }
  }, 'Approve Deposit', 'Approve', false);
}

function updateDepositKpis() {
  const baseDeposits = getBaseRecordsForRole(depositList);
  let totalCash = 0, totalCredit = 0;
  const agents = new Set();
  baseDeposits.forEach(function(d) { totalCash += (d.cash || 0); totalCredit += (d.credit || 0); agents.add(d.agent); });
  const total = totalCash + totalCredit;
  const el1 = g('dep-kpi-total'); if (el1) el1.textContent = fmtMoney(total);
  const el2 = g('dep-kpi-count'); if (el2) el2.textContent = baseDeposits.length;
  const el3 = g('dep-kpi-agents'); if (el3) el3.textContent = agents.size;
  const el4 = g('dep-kpi-cash'); if (el4) el4.textContent = fmtMoney(totalCash);
  const el5 = g('dep-kpi-credit'); if (el5) el5.textContent = fmtMoney(totalCredit);
  renderDepositChart();
}

function renderDepositChart() {
  _cDepositPerf = destroyChart(_cDepositPerf);
  const canvas = g('cDepositPerf');
  if (!canvas || typeof Chart === 'undefined') return;
  const periodEl = g('dep-chart-period');
  const period = periodEl ? periodEl.value : 'monthly';
  const baseDepositList = getBaseRecordsForRole(depositList);

  const now = new Date();
  let labels = [], cashData = [], creditData = [];

  if (period === 'weekly') {
    for (let i = 6; i >= 0; i--) {
      const d = new Date(now); d.setDate(d.getDate() - i);
      const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
      labels.push(key.slice(5));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date === key) { c += (dep.cash||0); cr += (dep.credit||0); } });
      cashData.push(c); creditData.push(cr);
    }
  } else if (period === 'monthly') {
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      labels.push(ymLabel(key));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date && dep.date.startsWith(key)) { c += (dep.cash||0); cr += (dep.credit||0); } });
      cashData.push(c); creditData.push(cr);
    }
  } else {
    for (let i = 2; i >= 0; i--) {
      const yr = now.getFullYear() - i;
      labels.push(String(yr));
      let c = 0, cr = 0;
      baseDepositList.forEach(function(dep) { if (dep.date && dep.date.startsWith(String(yr))) { c += (dep.cash||0); cr += (dep.credit||0); } });
      cashData.push(c); creditData.push(cr);
    }
  }

  _cDepositPerf = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Cash ($)', data: cashData, backgroundColor: 'rgba(27,125,61,0.8)', borderColor: '#1B7D3D', borderWidth: 1 },
        { label: 'Credit ($)', data: creditData, backgroundColor: 'rgba(21,101,192,0.8)', borderColor: '#1565C0', borderWidth: 1 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } } },
      scales: {
        x: { stacked: true, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, stacked: true, ticks: { font: { size: 10 } } }
      }
    }
  });
}

function renderDepositTable() {
  const table = g('deposit-table');
  if (!table) return;
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if (!tbody) return;

  const baseDepositList = getBaseRecordsForRole(depositList);

  // Update header to include cash, credit, status columns
  if (thead) {
    thead.innerHTML = '<tr><th>#</th><th>Agent</th><th>Branch</th><th>Cash ($)</th><th>Credit ($)</th><th>Total</th><th>Date</th><th>Remark</th><th>Status</th><th>Actions</th></tr>';
  }

  if (!baseDepositList.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-piggy-bank" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No deposit records yet</td></tr>';
    return;
  }
  const canApprove = (currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster');
  tbody.innerHTML = baseDepositList.map(function(d, i) {
    const avIdx = i % 8;
    const status = d.status || 'pending';
    const statusPill = status === 'approved' ? 'pill-green' : 'pill-orange';
    const statusLabel = status === 'approved' ? 'Approved' : 'Pending';
    const approveBtn = (canApprove && status !== 'approved') ? '<button class="btn-edit" onclick="approveDeposit(\'' + esc(d.id) + '\')" title="Approve"><i class="fas fa-check-circle"></i></button> ' : '';
    const canEdit = canModifyRecord(d);
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(d.agent)) + '</span>' + esc(d.agent) + '</div></td>' +
      '<td>' + esc(d.branch || '') + '</td>' +
      '<td style="color:#1B7D3D;font-weight:600;">' + (d.cash ? '$' + Number(d.cash).toFixed(2) : '—') + '</td>' +
      '<td style="color:#1565C0;font-weight:600;">' + (d.credit ? '$' + Number(d.credit).toFixed(2) : '—') + '</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">$' + Number((d.cash || 0) + (d.credit || 0)).toFixed(2) + '</td>' +
      '<td>' + esc(d.date || '') + '</td>' +
      '<td style="color:#888;font-size:0.8rem;">' + esc(d.remark || d.note || '') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + statusLabel + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        approveBtn +
        (canEdit ? '<button class="btn-edit" onclick="editDeposit(\'' + esc(d.id) + '\')"><i class="fas fa-edit"></i></button> ' : '') +
        (canEdit ? '<button class="btn-delete" onclick="deleteDeposit(\'' + esc(d.id) + '\')"><i class="fas fa-trash"></i></button>' : '') +
      '</td>' +
      '</tr>';
  }).join('');
}

function setDepositView(view) {
  var depBtns = document.querySelectorAll('#page-deposit .view-toggle-btn');
  depBtns.forEach(function(b) { b.classList.remove('active'); });
  var btn = g(view === 'table' ? 'dep-view-btn-table' : 'dep-view-btn-summary');
  if (btn) btn.classList.add('active');

  var tableCard = g('deposit-table-card');
  var summaryView = g('deposit-summary-view');

  if (view === 'table') {
    if (tableCard) tableCard.style.display = '';
    if (summaryView) summaryView.style.display = 'none';
  } else {
    if (tableCard) tableCard.style.display = 'none';
    if (summaryView) summaryView.style.display = '';
    renderDepositSummaryView();
  }
}

function renderDepositSummaryView() {
  var container = g('deposit-summary-view');
  if (!container) return;

  const baseDepositList = getBaseRecordsForRole(depositList);
  if (!baseDepositList.length) {
    container.innerHTML = '<div style="text-align:center;padding:40px;color:#999;"><i class="fas fa-inbox fa-3x" style="display:block;margin-bottom:12px;"></i>No deposit records found</div>';
    return;
  }

  var agentMap = {};
  baseDepositList.forEach(function(d) {
    if (!agentMap[d.agent]) agentMap[d.agent] = { cash: 0, credit: 0, total: 0, count: 0, branch: d.branch || '' };
    var ag = agentMap[d.agent];
    ag.cash += (d.cash || 0);
    ag.credit += (d.credit || 0);
    ag.total += (d.cash || 0) + (d.credit || 0);
    ag.count++;
  });

  var cards = Object.keys(agentMap).map(function(agent, i) {
    var ag = agentMap[agent];
    var avIdx = i % 8;
    return '<div class="summary-card">' +
      '<div class="summary-card-header">' +
        '<span class="avatar-circle av-' + avIdx + '" style="width:36px;height:36px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.8rem;font-weight:700;color:#fff;">' + esc(ini(agent)) + '</span>' +
        '<div><div class="sc-name">' + esc(agent) + '</div><div style="font-size:0.72rem;opacity:0.8;">' + esc(ag.branch) + ' &middot; ' + ag.count + ' deposit' + (ag.count !== 1 ? 's' : '') + '</div></div>' +
      '</div>' +
      '<div class="summary-card-body">' +
        '<div class="summary-row"><span>Cash</span><span style="color:#1B7D3D;font-weight:600;">$' + Number(ag.cash).toFixed(2) + '</span></div>' +
        '<div class="summary-row"><span>Credit</span><span style="color:#1565C0;font-weight:600;">$' + Number(ag.credit).toFixed(2) + '</span></div>' +
        '<div class="summary-row"><span>Total</span><span style="font-weight:700;color:#1A1A2E;">$' + Number(ag.total).toFixed(2) + '</span></div>' +
      '</div>' +
    '</div>';
  }).join('');

  container.innerHTML = '<div class="summary-grid">' + cards + '</div>';
}

// ------------------------------------------------------------
// Staff Functions
// ------------------------------------------------------------
function openUserModal(user) {
  const form = g('form-addUser');
  if (form) form.reset();
  g('user-edit-id').value = '';
  const title = g('modal-addUser-title');
  const btn = g('user-submit-btn');
  populateBranchSelects();

  if (user) {
    if (title) title.textContent = 'Edit User';
    if (btn) btn.textContent = 'Update User';
    g('user-edit-id').value = user.id;
    g('user-name').value = user.name || '';
    g('user-username').value = user.username || '';
    g('user-password').value = user.password || '';
    g('user-role').value = user.role || 'Agent';
    const bInput = g('user-branch'); if (bInput) bInput.value = user.branch || '';
    g('user-status').value = user.status || 'active';
    const emailInput = g('user-email'); if (emailInput) emailInput.value = user.email || '';
  } else {
    if (title) title.textContent = 'Add User';
    if (btn) btn.textContent = 'Add User';
  }
  openModal('modal-addUser');
}

function submitUser(e) {
  e.preventDefault();
  const editId = rv('user-edit-id');
  const obj = {
    id: editId || uid(),
    name: rv('user-name'), username: rv('user-username'), password: rv('user-password'),
    role: rv('user-role'), branch: rv('user-branch'), status: rv('user-status'),
    email: rv('user-email') || ''
  };
  if (!obj.name) { showAlert('Please enter user name'); return; }
  if (!obj.username) { showAlert('Please enter username'); return; }
  if (!editId && !obj.password) { showAlert('Please enter a password for the new user'); return; }
  const dupUser = staffList.find(function(x) { return x.username.toLowerCase() === obj.username.toLowerCase() && x.id !== editId; });
  if (dupUser) { showAlert('Username already exists. Please choose a different username.'); return; }
  if (editId) {
    const idx = staffList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) {
      // If password field was left blank on edit, preserve the existing password
      if (!obj.password) { obj.password = staffList[idx].password || ''; }
      staffList[idx] = obj;
    }
  } else {
    staffList.push(obj);
  }
  closeModal('modal-addUser');
  renderStaffTable();
  syncSheet('Staff', staffList);
  saveAllData();
  showToast(editId ? 'User updated successfully.' : 'User added successfully.', 'success');
}

function editUser(id) {
  const user = staffList.find(function(x) { return x.id === id; });
  if (user) openUserModal(user);
}

function deleteUser(id) {
  showConfirm('Are you sure you want to delete this user? This action cannot be undone.', function() {
    staffList = staffList.filter(function(x) { return x.id !== id; });
    renderStaffTable();
    syncSheet('Staff', staffList);
    saveAllData();
    showToast('User deleted.', 'success');
  }, 'Delete User', 'Delete');
}

function renderStaffTable() {
  const tbody = g('staff-table');
  if (!tbody) return;
  if (!staffList.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-users-cog" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No users yet</td></tr>';
    return;
  }
  tbody.innerHTML = staffList.map(function(u, i) {
    const rolePill = u.role === 'Admin' ? 'pill-green' : u.role === 'Cluster' ? 'pill-purple' : u.role === 'Supervisor' ? 'pill-blue' : u.role === 'Agent' ? 'pill-orange' : 'pill-gray';
    const statusPill = u.status === 'active' ? 'pill-green' : 'pill-red';
    const avIdx = i % 8;
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td><div class="name-cell"><span class="avatar-circle av-' + avIdx + '" style="width:30px;height:30px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:0.7rem;font-weight:700;color:#fff;margin-right:8px;">' + esc(ini(u.name)) + '</span>' + esc(u.name) + '</div></td>' +
      '<td>' + esc(u.username) + '</td>' +
      '<td><span class="pill ' + rolePill + '">' + esc(u.role) + '</span></td>' +
      '<td>' + esc(u.branch || '') + '</td>' +
      '<td>' + (u.email
        ? '<a href="mailto:' + encodeURIComponent(u.email) + '" style="color:#1B7D3D;text-decoration:none;"><i class="fas fa-envelope" style="margin-right:4px;font-size:.8rem;"></i>' + esc(u.email) + '</a>'
        : '<span style="color:#ccc;">—</span>') + '</td>' +
      '<td><span class="pill ' + statusPill + '">' + esc(u.status) + '</span></td>' +
      '<td style="white-space:nowrap;">' +
        '<button class="btn-edit" onclick="editUser(\'' + esc(u.id) + '\')"><i class="fas fa-edit"></i></button> ' +
        '<button class="btn-delete" onclick="deleteUser(\'' + esc(u.id) + '\')"><i class="fas fa-trash"></i></button>' +
      '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// KPI Functions
// ------------------------------------------------------------
function onUserRoleChange() {
  // reserved for role-specific UI changes
}

function onKpiShopNameInput(val) {
  var hiddenEl = g('kpi-shop-assignee');
  if (!hiddenEl) return;
  var sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });
  var searchVal = (val || '').trim().toLowerCase();
  var found = sups.find(function(u) { return u.name.toLowerCase() === searchVal; });
  hiddenEl.value = found ? found.id : '';
}

function selectKpiFor(type) {
  kpiForSelected = type;
  const shopBtn = g('kpi-for-shop');
  const agentBtn = g('kpi-for-agent');
  if (shopBtn) { shopBtn.classList.toggle('active', type === 'shop'); }
  if (agentBtn) { agentBtn.classList.toggle('active', type === 'agent'); }

  const shopGroup = g('kpi-shop-assignee-group');
  const agentBranchGroup = g('kpi-agent-branch-group');
  const agentAssigneeGroup = g('kpi-agent-assignee-group');

  if (type === 'shop') {
    if (shopGroup) shopGroup.style.display = '';
    if (agentBranchGroup) agentBranchGroup.style.display = 'none';
    if (agentAssigneeGroup) agentAssigneeGroup.style.display = 'none';
    populateKpiShopAssignee();
  } else {
    if (shopGroup) shopGroup.style.display = 'none';
    if (agentBranchGroup) agentBranchGroup.style.display = '';
    if (agentAssigneeGroup) agentAssigneeGroup.style.display = '';
    populateKpiAgentBranch();
  }
}

function populateKpiShopAssignee(preselectedId) {
  const textEl = g('kpi-shop-assignee-name');
  const hiddenEl = g('kpi-shop-assignee');
  const datalist = g('kpi-sup-list');
  if (!textEl) return;
  const sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });

  if (currentRole === 'supervisor' && currentUser) {
    // Supervisor can only assign KPI to themselves — show readonly textbox
    textEl.value = currentUser.name;
    textEl.readOnly = true;
    textEl.style.background = '#f5f5f5';
    if (hiddenEl) hiddenEl.value = currentUser.id;
    if (datalist) datalist.innerHTML = '';
  } else {
    // Admin: populate datalist with all supervisors, allow typing to choose
    textEl.readOnly = false;
    textEl.style.background = '';
    if (datalist) {
      datalist.innerHTML = sups.map(function(u) {
        return '<option value="' + esc(u.name) + '"></option>';
      }).join('');
    }
    if (preselectedId) {
      var sup = sups.find(function(u) { return u.id === preselectedId; });
      textEl.value = sup ? sup.name : '';
      if (hiddenEl) hiddenEl.value = preselectedId;
    } else {
      textEl.value = '';
      if (hiddenEl) hiddenEl.value = '';
    }
  }
}

function populateKpiAgentBranch() {
  const branchSel = g('kpi-agent-branch');
  if (!branchSel) return;

  if (currentRole === 'supervisor' && currentUser) {
    // Lock branch dropdown to supervisor's own branch
    branchSel.innerHTML = '<option value="' + esc(currentUser.branch) + '">' + esc(currentUser.branch) + '</option>';
    branchSel.disabled = true;
    branchSel.value = currentUser.branch;
  } else {
    branchSel.disabled = false;
    branchSel.innerHTML = '<option value="">Select branch</option>' +
      getBranches().map(function(b) { return '<option value="' + esc(b) + '">' + esc(b) + '</option>'; }).join('');
  }

  const agentSel = g('kpi-agent-assignee');
  if (agentSel) agentSel.innerHTML = '<option value="">Select Agent</option>';

  // Auto-load agents when supervisor's branch is pre-set
  if (currentRole === 'supervisor' && currentUser) {
    onKpiBranchChange();
  }
}

function onKpiBranchChange() {
  const branch = (currentRole === 'supervisor' && currentUser) ? currentUser.branch : rv('kpi-agent-branch');
  const agentSel = g('kpi-agent-assignee');
  if (!agentSel) return;
  const agents = staffList.filter(function(u) { return u.role === 'Agent' && u.branch === branch; });
  agentSel.innerHTML = '<option value="">Select Agent</option>' +
    agents.map(function(u) { return '<option value="' + esc(u.id) + '">' + esc(u.name) + '</option>'; }).join('');
}

function openKpiModal(item) {
  const form = g('form-kpi');
  if (form) form.reset();
  g('kpi-edit-id').value = '';
  kpiTypeSelected = 'Sales';
  setValueMode('unit');
  kpiForSelected = (item && item.kpiFor) ? item.kpiFor : 'shop';
  selectKpiFor(kpiForSelected);

  const title = g('modal-kpi-title');
  const btn = g('kpi-submit-btn');

  $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
  const firstChip = g('kpi-chip-Sales');
  if (firstChip) firstChip.classList.add('active');

  if (item) {
    if (title) title.textContent = 'Edit KPI';
    if (btn) btn.textContent = 'Update KPI';
    g('kpi-edit-id').value = item.id;
    g('kpi-name').value = item.name || '';
    g('kpi-target').value = item.target || '';
    g('kpi-period').value = item.period || 'Monthly';

    kpiTypeSelected = item.type || 'Sales';
    $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
    const chip = g('kpi-chip-' + kpiTypeSelected);
    if (chip) chip.classList.add('active');

    setValueMode(item.valueMode || 'unit');
    if (item.valueMode === 'unit') {
      const itemSel = g('kpi-item-sel'); if (itemSel && item.itemId) itemSel.value = item.itemId;
    } else {
      const csEl = g('kpi-currency-sel'); if (csEl) csEl.value = item.currency || 'USD';
    }
    if (item.kpiFor === 'shop') {
      populateKpiShopAssignee(item.assigneeId);
    } else if (item.kpiFor === 'agent') {
      populateKpiAgentBranch();
      const branchSel = g('kpi-agent-branch');
      if (branchSel && item.assigneeBranch) {
        branchSel.value = item.assigneeBranch;
        onKpiBranchChange();
        const agentSel = g('kpi-agent-assignee');
        if (agentSel && item.assigneeId) agentSel.value = item.assigneeId;
      }
    }
  } else {
    if (title) title.textContent = 'Add KPI';
    if (btn) btn.textContent = 'Add KPI';
  }
  openModal('modal-kpi');
}

function selectKpiType(el) {
  const type = el ? el.getAttribute('data-value') : 'Sales';
  kpiTypeSelected = type;
  $$('.kpi-type-chip').forEach(function(c) { c.classList.remove('active'); });
  if (el) el.classList.add('active');
}

function setValueMode(mode) {
  kpiValueMode = mode;
  const unitField = g('kpi-unit-field');
  const curField = g('kpi-currency-field');
  const unitToggle = g('kpi-unit-toggle');
  const curToggle = g('kpi-currency-toggle');

  if (mode === 'unit') {
    if (unitField) unitField.style.display = '';
    if (curField) curField.style.display = 'none';
    if (unitToggle) unitToggle.classList.add('active');
    if (curToggle) curToggle.classList.remove('active');
    populateKpiItemSel();
  } else {
    if (unitField) unitField.style.display = 'none';
    if (curField) curField.style.display = '';
    if (unitToggle) unitToggle.classList.remove('active');
    if (curToggle) curToggle.classList.add('active');
  }
}

function populateKpiItemSel() {
  const sel = g('kpi-item-sel');
  if (!sel) return;
  const unitItems = itemCatalogue.filter(function(x) { return x.group === 'unit' && x.status === 'active'; });
  sel.innerHTML = '<option value="">All Unit Items</option>' +
    unitItems.map(function(item) {
      return '<option value="' + esc(item.id) + '">' + esc(item.name) + '</option>';
    }).join('');
}

function submitKpi(e) {
  e.preventDefault();
  const editId = rv('kpi-edit-id');

  // Permission check: only admin, cluster, and supervisor can create/edit KPIs
  if (!canManageKpis()) { showAlert('You do not have permission to manage KPIs.', 'error'); return; }

  // Resolve supervisor id from text input (for admin role)
  if (kpiForSelected === 'shop') {
    var shopName = rv('kpi-shop-assignee-name');
    var hiddenEl = g('kpi-shop-assignee');
    var isSupervisorWithAutoFill = currentRole === 'supervisor' && currentUser;
    if (!isSupervisorWithAutoFill && shopName && hiddenEl && !hiddenEl.value) {
      var sups = staffList.filter(function(u) { return u.role === 'Supervisor'; });
      var found = sups.find(function(u) { return u.name.toLowerCase() === shopName.trim().toLowerCase(); });
      if (found) hiddenEl.value = found.id;
    }
  }

  const obj = {
    id: editId || uid(),
    name: rv('kpi-name'),
    type: kpiTypeSelected,
    kpiFor: kpiForSelected,
    assigneeId: kpiForSelected === 'shop' ? rv('kpi-shop-assignee') : rv('kpi-agent-assignee'),
    assigneeBranch: kpiForSelected === 'agent' ? rv('kpi-agent-branch') : '',
    target: parseFloat(rv('kpi-target')) || 0,
    valueMode: kpiValueMode,
    itemId: kpiValueMode === 'unit' ? rv('kpi-item-sel') : '',
    unit: kpiValueMode === 'unit' ? rv('kpi-unit-val') : '',
    currency: kpiValueMode === 'currency' ? rv('kpi-currency-sel') : '',
    period: rv('kpi-period')
  };
  if (!obj.name) { showAlert('Please enter KPI name'); return; }
  if (kpiForSelected === 'shop' && !obj.assigneeId) { showAlert('Please select a supervisor'); return; }
  if (kpiForSelected === 'agent' && !obj.assigneeId) { showAlert('Please select an agent'); return; }
  if (editId) {
    const idx = kpiList.findIndex(function(x) { return x.id === editId; });
    if (idx >= 0) kpiList[idx] = obj;
  } else {
    kpiList.push(obj);
  }
  closeModal('modal-kpi');
  renderKpiTable();
  syncSheet('KPI', kpiList);
  saveAllData();
  showToast(editId ? 'KPI updated successfully.' : 'KPI added successfully.', 'success');
  // Refresh dashboard KPI section if on dashboard
  if (currentPage === 'dashboard') renderDashboardKpiSection();
}

function editKpi(id) {
  const item = kpiList.find(function(x) { return x.id === id; });
  if (!item) return;
  // Permission check: only admin, cluster, or supervisor (for their own KPIs) can edit
  if (!canManageKpis()) { showAlert('You do not have permission to edit KPIs.', 'error'); return; }
  if (currentRole === 'supervisor' && !canSupervisorModifyKpi(item)) {
    showAlert('You can only edit KPIs assigned to your shop or agents.', 'error'); return;
  }
  openKpiModal(item);
}

function deleteKpi(id) {
  const item = kpiList.find(function(x) { return x.id === id; });
  if (!item) return;
  // Permission check
  if (!canManageKpis()) { showAlert('You do not have permission to delete KPIs.', 'error'); return; }
  if (currentRole === 'supervisor' && !canSupervisorModifyKpi(item)) {
    showAlert('You can only delete KPIs assigned to your shop or agents.', 'error'); return;
  }
  showConfirm('Are you sure you want to delete this KPI? This action cannot be undone.', function() {
    kpiList = kpiList.filter(function(x) { return x.id !== id; });
    renderKpiTable();
    syncSheet('KPI', kpiList);
    saveAllData();
    showToast('KPI deleted.', 'success');
    if (currentPage === 'dashboard') renderDashboardKpiSection();
  }, 'Delete KPI', 'Delete');
}

function setKpiMonth(mode) {
  const picker = g('kpi-month-picker');
  if (mode === 'current') {
    kpiSelectedMonth = ymNow();
  } else if (mode === 'prev') {
    const d = new Date();
    d.setMonth(d.getMonth() - 1);
    kpiSelectedMonth = d.toISOString().substring(0, 7);
  } else {
    kpiSelectedMonth = '';
  }
  if (picker) picker.value = kpiSelectedMonth;
  renderKpiTable();
}

function onKpiMonthChange() {
  const picker = g('kpi-month-picker');
  kpiSelectedMonth = picker ? picker.value : '';
  renderKpiTable();
}

function initKpiMonthPicker() {
  if (!kpiSelectedMonth) kpiSelectedMonth = ymNow();
  const picker = g('kpi-month-picker');
  if (picker) picker.value = kpiSelectedMonth;
}

// Returns the branch associated with a KPI entry.
// For agent KPIs the branch is stored directly; for shop KPIs it comes from the assignee supervisor.
function getKpiBranch(k) {
  if (k.kpiFor === 'agent') return k.assigneeBranch || '';
  if (k.kpiFor === 'shop') {
    var sup = staffList.find(function(u) { return u.id === k.assigneeId; });
    return sup ? (sup.branch || '') : '';
  }
  return '';
}

function renderKpiTable() {
  const tbody = g('kpi-table');
  if (!tbody) return;

  // Show/hide KPI add button based on role
  var kpiAddBtn = g('kpi-add-btn');
  if (kpiAddBtn) {
    kpiAddBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster' || currentRole === 'supervisor') ? '' : 'none';
  }

  // Only show KPIs visible to the current user's branch
  var visibleKpis = getVisibleKpis();

  if (!visibleKpis.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-chart-line" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No KPIs defined yet</td></tr>';
    return;
  }

  // Group KPIs by branch
  var branchMap = {};
  var branchOrder = [];
  visibleKpis.forEach(function(k) {
    var branch = getKpiBranch(k) || 'Unassigned';
    if (!branchMap[branch]) {
      branchMap[branch] = [];
      branchOrder.push(branch);
    }
    branchMap[branch].push(k);
  });
  branchOrder.sort();

  var html = '';
  branchOrder.forEach(function(branch) {
    var kpis = branchMap[branch];
    var rowNum = 0;
    html += '<tr><td colspan="9" class="kpi-branch-header">' +
      '<i class="fas fa-building" style="margin-right:6px;"></i>' + esc(branch) +
      ' <span class="kpi-branch-count">(' + kpis.length + ' KPI' + (kpis.length !== 1 ? 's' : '') + ')</span>' +
      '</td></tr>';
    kpis.forEach(function(k) {
      rowNum++;
      const typePill = k.type === 'Sales' ? 'pill-green' : k.type === 'Revenue' ? 'pill-orange' : k.type === 'Units' ? 'pill-blue' : 'pill-purple';
      const unitLabel = getKpiUnitLabel(k);
      const valueDisplay = k.valueMode === 'currency'
        ? fmtMoney(k.target, esc(k.currency) + ' ')
        : k.target + (unitLabel ? ' ' + esc(unitLabel) : '');
      const assignee = staffList.find(function(u) { return u.id === k.assigneeId; });
      const forLabel = k.kpiFor === 'shop' ? '<span class="pill pill-blue"><i class="fas fa-store"></i> Shop</span>' : '<span class="pill pill-orange"><i class="fas fa-user"></i> Agent</span>';
      const assigneeName = assignee ? esc(assignee.name) : (k.assigneeBranch ? esc(k.assigneeBranch) : '—');
      // Compute actual & achievement for selected month
      const ym = kpiSelectedMonth || ymNow();
      const actual = Math.round(calcKpiActual(k, ym) * 100) / 100;
      const pct = k.target > 0 ? Math.round(actual / k.target * 100) : 0;
      const pctClass = pct >= 100 ? 'pill-green' : pct >= 70 ? 'pill-orange' : 'pill-red';
      const actualDisplay = k.valueMode === 'currency'
        ? fmtMoney(actual, esc(k.currency) + ' ')
        : actual + (unitLabel ? ' ' + esc(unitLabel) : '');
      const progressBar = '<div style="background:#eee;border-radius:4px;height:6px;width:80px;display:inline-block;vertical-align:middle;margin-right:4px;">' +
        '<div style="background:' + (pct >= 100 ? '#1B7D3D' : pct >= 70 ? '#FF9800' : '#E53935') + ';width:' + Math.min(pct, 100) + '%;height:100%;border-radius:4px;"></div></div>';
      // Determine if the current user can modify this KPI using centralised helpers
      var canModifyKpi = canManageKpis() &&
        (currentRole !== 'supervisor' || canSupervisorModifyKpi(k));
      var actionBtns = canModifyKpi
        ? '<button class="btn-edit" onclick="editKpi(\'' + esc(k.id) + '\')"><i class="fas fa-edit"></i></button> ' +
          '<button class="btn-delete" onclick="deleteKpi(\'' + esc(k.id) + '\')"><i class="fas fa-trash"></i></button>'
        : '<span style="color:#bbb;font-size:.75rem;">View only</span>';
      html += '<tr>' +
        '<td>' + rowNum + '</td>' +
        '<td>' + esc(k.name) + '</td>' +
        '<td><span class="pill ' + typePill + '">' + esc(k.type) + '</span></td>' +
        '<td>' + forLabel + '<br><small style="color:#888;">' + assigneeName + '</small></td>' +
        '<td>' + valueDisplay + '</td>' +
        '<td>' + actualDisplay + '</td>' +
        '<td>' + progressBar + '<span class="pill ' + pctClass + '" style="font-size:.72rem;">' + pct + '%</span></td>' +
        '<td>' + esc(k.period || '') + '</td>' +
        '<td style="white-space:nowrap;">' + actionBtns + '</td>' +
        '</tr>';
    });
  });
  tbody.innerHTML = html;
}

// ------------------------------------------------------------
// Login / Logout
// ------------------------------------------------------------
function handleLogin(e) {
  e.preventDefault();
  var username = rv('login-username');
  var password = rv('login-password');
  var errEl = g('login-error');
  var btn = g('login-submit-btn');

  console.log('[AUTH] Login attempt started');
  console.log('[AUTH] Username entered:', username);

  if (!username || !password) {
    console.warn('[AUTH] Missing credentials');
    if (errEl) { errEl.textContent = 'Please enter username and password.'; errEl.style.display = ''; }
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Signing in\u2026'; }
  function doAuth() {
    console.log('[AUTH] Starting authentication check');
    console.log('[AUTH] Total staff records loaded:', staffList.length);
    console.log('[AUTH] Staff usernames available:', staffList.map(function(u) { return u.username; }));

    var user = staffList.find(function(u) {
      var usernameMatch = (u.username || '').toLowerCase() === username.toLowerCase();
      var passwordMatch = u.password === password;
      var statusMatch = (u.status || '').toLowerCase() === 'active';

      if (usernameMatch) {
        console.log('[AUTH] Username match found:', u.username);
        console.log('[AUTH] Password match:', passwordMatch);
        console.log('[AUTH] Status:', u.status, '| Status active:', statusMatch);
      }

      return usernameMatch && passwordMatch && statusMatch;
    });
    if (user) {
      console.log('[AUTH] \u2713 Authentication successful for user:', user.username);
      var roleMap = { 'Admin': 'admin', 'Cluster': 'cluster', 'Supervisor': 'supervisor', 'Agent': 'agent' };
      currentUser = user;
      if (errEl) errEl.style.display = 'none';
      var ls = g('login-screen'); if (ls) ls.style.display = 'none';
      var as = g('app-shell'); if (as) { as.style.display = 'flex'; }
      switchRole(roleMap[user.role] || 'user');
      // switchRole overwrites currentUser with a representative user for demo switching;
      // restore it to the authenticated user so all data filtering uses the correct identity.
      currentUser = user;
      lsSave(LS_KEYS.session, user);
      var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = user.name;
      var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(user.name);
      navigateTo('dashboard', null);
      // Auto-refresh all data from Google Sheets on login so users always see the latest records
      refreshAllData();
    } else {
      console.error('[AUTH] \u2717 Authentication failed');
      console.error('[AUTH] No matching user found for username:', username);
      console.error('[AUTH] Available users:', staffList.map(function(u) {
        return { username: u.username, status: u.status, hasPassword: !!u.password };
      }));
      if (errEl) { errEl.textContent = 'Invalid username or password, or account is inactive.'; errEl.style.display = ''; }
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-right-to-bracket"></i> Sign In'; }
    }
  }
  console.log('[AUTH] Fetching staff data from Google Sheets...');
  fetchStaffFromSheet()
    .then(function() {
      console.log('[AUTH] Staff data fetch completed');
      doAuth();
    })
    .catch(function(err) {
      console.error('[AUTH] Staff data fetch error:', err);
      doAuth(); // Still attempt auth with cached data
    });
}

function toggleLoginPwd() {
  var inp = g('login-password');
  var eye = g('login-pwd-eye');
  if (!inp) return;
  if (inp.type === 'password') { inp.type = 'text'; if (eye) eye.className = 'fas fa-eye-slash'; }
  else { inp.type = 'password'; if (eye) eye.className = 'fas fa-eye'; }
}

function handleLogout() {
  showConfirm('Are you sure you want to sign out of Smart 5G Dashboard?', function() {
    currentUser = null;
    localStorage.removeItem(LS_KEYS.session);
    var as = g('app-shell'); if (as) as.style.display = 'none';
    var ls = g('login-screen'); if (ls) ls.style.display = 'flex';
    var lf = g('login-form'); if (lf) lf.reset();
    var errEl = g('login-error'); if (errEl) errEl.style.display = 'none';
    var btn = g('login-submit-btn'); if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-right-to-bracket"></i> Sign In'; }
  }, 'Sign Out', 'Sign Out', false);
}

function addNotification(msg) {
  notifications.push({ id: uid(), msg: msg, time: new Date().toLocaleString(), read: false });
  updateNotificationBadge();
}

function updateNotificationBadge() {
  var unread = notifications.filter(function(n) { return !n.read; }).length;
  var badge = g('notif-badge');
  if (badge) {
    badge.textContent = unread > 9 ? '9+' : (unread || '');
    badge.style.display = unread > 0 ? 'flex' : 'none';
  }
}

function toggleNotifications() {
  var panel = g('notif-panel');
  if (!panel) return;
  var isOpen = panel.style.display === 'block';
  panel.style.display = isOpen ? 'none' : 'block';
  if (!isOpen) {
    notifications.forEach(function(n) { n.read = true; });
    updateNotificationBadge();
    renderNotificationPanel();
  }
}

function renderNotificationPanel() {
  var list = g('notif-list');
  if (!list) return;
  if (!notifications.length) {
    list.innerHTML = '<div style="text-align:center;padding:24px;color:#999;font-size:.85rem;"><i class="fas fa-bell-slash" style="font-size:1.5rem;display:block;margin-bottom:8px;opacity:.3;"></i>No notifications yet</div>';
    return;
  }
  list.innerHTML = notifications.slice().reverse().map(function(n) {
    return '<div style="padding:10px 16px;border-bottom:1px solid #F5F5F5;font-size:.82rem;">' +
      '<div style="color:#1A1A2E;font-weight:500;">' + esc(n.msg) + '</div>' +
      '<div style="color:#999;font-size:.72rem;margin-top:2px;">' + esc(n.time) + '</div>' +
    '</div>';
  }).join('');
}

function loginContactSupport() {
  showAlert('Email: ' + SUPPORT_CONTACT.email + '\nPhone: ' + SUPPORT_CONTACT.phone, 'info', 'Contact Support');
}

// ------------------------------------------------------------
// Contact Support Modal
// ------------------------------------------------------------
function openContactSupportModal() {
  if (currentUser) {
    var nameEl = g('cs-name');
    var usernameEl = g('cs-username');
    var branchEl = g('cs-branch');
    if (nameEl) nameEl.value = currentUser.name || '';
    if (usernameEl) usernameEl.value = currentUser.username || '';
    if (branchEl) branchEl.value = currentUser.branch || '';
  }
  var msgEl = g('cs-message');
  if (msgEl) msgEl.value = '';
  openModal('modal-contactSupport');
}

function submitContactSupport() {
  var msg = g('cs-message') ? g('cs-message').value.trim() : '';
  if (!msg) { showToast('Please enter a message.', 'error'); return; }
  closeModal('modal-contactSupport');
  showToast('Support request sent successfully!', 'success');
}

// ------------------------------------------------------------
// Compatibility Aliases
// ------------------------------------------------------------
function applySaleFilters() { applyReportFilters(); }
function clearSaleFilters() { clearReportFilters(); }
function loadDashboard() { renderDashboard(); }
function selectKpiMode(mode) { setValueMode(mode); }
function submitKPI(e) { submitKpi(e); }
function switchSaleView(view) { setReportView(view); }
function togglePasswordVisibility(inputId, eyeId) { togglePwd(inputId, eyeId); }
function toggleSidebar() {
  const sidebar = g('sidebar');
  if (sidebar) sidebar.classList.toggle('sidebar-collapsed');
}

// ------------------------------------------------------------
// Social Login
// ------------------------------------------------------------
function loginWithSocial(provider) {
  if (provider === 'telegram') {
    window.open('https://t.me/saray2026123', '_blank', 'noopener,noreferrer');
    return;
  }
  if (provider === 'phone') {
    openModal('modal-phone-login');
    return;
  }
  var names = { facebook: 'Facebook', google: 'Google' };
  showAlert(
    'To register with ' + names[provider] + ', please contact the admin on Telegram:\n\n@saray2026123\n\nThey will help set up your account.',
    'info',
    'Register with ' + names[provider]
  );
}

function submitPhoneLogin() {
  var phone = g('phone-login-number') ? g('phone-login-number').value.trim() : '';
  var name = g('phone-login-name') ? g('phone-login-name').value.trim() : '';
  if (!phone || !name) { showToast('Please fill in all required fields.', 'error'); return; }
  closeModal('modal-phone-login');
  showToast('Request submitted! Please contact @saray2026123 on Telegram to complete setup.', 'success');
}

// ------------------------------------------------------------
// Inventory Management
// ------------------------------------------------------------
var LS_INV_STOCK = 'smart5g_inv_stock';
var LS_INV_REQUESTS = 'smart5g_inv_requests';
var LS_INV_HISTORY = 'smart5g_inv_history';

var invStock = [];       // { id, itemId, itemName, category, inStock, allocated, lastUpdated }
var invRequests = [];    // { id, date, requestedBy, itemId, itemName, qty, purpose, status, reviewedBy, reviewNote }
var invHistory = [];     // { id, date, itemId, itemName, type('in'/'out'), qty, before, after, by, note }
var _reviewAllocId = null;

function loadInvData() {
  try { invStock = JSON.parse(localStorage.getItem(LS_INV_STOCK)) || []; } catch(e) { invStock = []; }
  try { invRequests = JSON.parse(localStorage.getItem(LS_INV_REQUESTS)) || []; } catch(e) { invRequests = []; }
  try { invHistory = JSON.parse(localStorage.getItem(LS_INV_HISTORY)) || []; } catch(e) { invHistory = []; }
}

function saveInvData() {
  localStorage.setItem(LS_INV_STOCK, JSON.stringify(invStock));
  localStorage.setItem(LS_INV_REQUESTS, JSON.stringify(invRequests));
  localStorage.setItem(LS_INV_HISTORY, JSON.stringify(invHistory));
}

function getAvailableQty(itemId) {
  var entry = invStock.find(function(s) { return s.itemId === itemId; });
  if (!entry) return 0;
  return Math.max(0, (entry.inStock || 0) - (entry.allocated || 0));
}

function openInvTab(tab, btn) {
  $$('.inv-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('invtab-content-' + tab);
  if (tc) tc.classList.add('active');
  $$('.tab-btn[id^="invtab-"]').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
  if (tab === 'requests') renderInvRequestsTable();
  if (tab === 'history') renderInvHistoryTable();
}

function renderInvSaleStock() {
  loadInvData();
  // Update KPI
  var totalIn = invStock.reduce(function(s, x) { return s + (x.inStock || 0); }, 0);
  var totalAlloc = invStock.reduce(function(s, x) { return s + (x.allocated || 0); }, 0);
  var pending = invRequests.filter(function(r) { return r.status === 'pending'; }).length;
  var kvTotal = g('inv-kv-total'); if (kvTotal) kvTotal.textContent = totalIn;
  var kvAlloc = g('inv-kv-allocated'); if (kvAlloc) kvAlloc.textContent = totalAlloc;
  var kvPending = g('inv-kv-pending'); if (kvPending) kvPending.textContent = pending;

  // Role-based UI
  var isSup = currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster';
  var supActions = g('inv-sale-sup-actions'); if (supActions) supActions.style.display = isSup ? '' : 'none';
  var allocBtn = g('inv-allocate-btn'); if (allocBtn) allocBtn.style.display = isSup ? '' : 'none';
  var newReqBtn = g('inv-new-request-btn'); if (newReqBtn) newReqBtn.style.display = isSup ? '' : 'none';
  var actionsCol = g('inv-stock-actions-col'); if (actionsCol) actionsCol.style.display = isSup ? '' : 'none';

  // Render stock table
  var tbody = g('inv-stock-table');
  if (!tbody) return;
  if (!invStock.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No stock items yet.' + (isSup ? ' Click <strong>Add Stock</strong> to get started.' : '') + '</td></tr>';
    return;
  }
  tbody.innerHTML = invStock.map(function(s, idx) {
    var avail = Math.max(0, (s.inStock || 0) - (s.allocated || 0));
    var statusBadge = avail <= 5
      ? '<span class="inv-badge-status inv-badge-low"><i class="fas fa-triangle-exclamation"></i> Low</span>'
      : '<span class="inv-badge-status inv-badge-ok"><i class="fas fa-check"></i> OK</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td style="font-weight:600;">' + esc(s.itemName) + '</td>' +
      '<td>' + esc(s.category || '-') + '</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">' + (s.inStock || 0) + '</td>' +
      '<td style="color:#1565C0;">' + (s.allocated || 0) + '</td>' +
      '<td style="font-weight:700;">' + avail + ' ' + statusBadge + '</td>' +
      '<td style="color:#888;font-size:.8rem;">' + esc(s.lastUpdated || '-') + '</td>' +
      (isSup ? '<td><button class="btn-edit" onclick="openAddStockModal(\'' + esc(s.id) + '\')"><i class="fas fa-edit"></i></button> <button class="btn-delete" onclick="deleteStockItem(\'' + esc(s.id) + '\')"><i class="fas fa-trash"></i></button></td>' : '<td></td>') +
      '</tr>';
  }).join('');
}

function renderInvRequestsTable() {
  loadInvData();
  var filter = g('inv-req-status-filter') ? g('inv-req-status-filter').value : '';
  var isSup = currentRole === 'supervisor' || currentRole === 'admin' || currentRole === 'cluster';
  var tbody = g('inv-requests-table');
  if (!tbody) return;
  var data = invRequests.filter(function(r) { return !filter || r.status === filter; });
  if (!data.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No requests found.</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(function(r, idx) {
    var statusClass = { pending: 'inv-badge-pending', approved: 'inv-badge-approved', rejected: 'inv-badge-rejected' }[r.status] || 'inv-badge-pending';
    var statusLabel = r.status ? (r.status.charAt(0).toUpperCase() + r.status.slice(1)) : 'Pending';
    var actions = '';
    if (isSup && r.status === 'pending') {
      actions = '<button class="btn btn-sm btn-primary" onclick="openReviewAllocModal(\'' + esc(r.id) + '\')"><i class="fas fa-clipboard-check"></i> Review</button>';
    }
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(r.date || '-') + '</td>' +
      '<td>' + esc(r.requestedBy || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(r.itemName || '-') + '</td>' +
      '<td style="font-weight:700;">' + (r.qty || 0) + '</td>' +
      '<td style="color:#888;">' + esc(r.purpose || '-') + '</td>' +
      '<td><span class="inv-badge-status ' + statusClass + '">' + statusLabel + '</span></td>' +
      '<td>' + (actions || '<span style="color:#ccc;font-size:.8rem;">—</span>') + '</td>' +
      '</tr>';
  }).join('');
}

function renderInvHistoryTable() {
  loadInvData();
  var tbody = g('inv-history-table');
  if (!tbody) return;
  var data = invHistory.slice().reverse();
  if (!data.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-inbox fa-2x" style="display:block;margin-bottom:10px;"></i>No history yet.</td></tr>';
    return;
  }
  tbody.innerHTML = data.map(function(h, idx) {
    var typeLabel = h.type === 'in'
      ? '<span class="inv-type-in"><i class="fas fa-arrow-down"></i> IN</span>'
      : '<span class="inv-type-out"><i class="fas fa-arrow-up"></i> OUT</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(h.date || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(h.itemName || '-') + '</td>' +
      '<td>' + typeLabel + '</td>' +
      '<td style="font-weight:700;">' + (h.qty || 0) + '</td>' +
      '<td>' + (h.before || 0) + '</td>' +
      '<td>' + (h.after || 0) + '</td>' +
      '<td>' + esc(h.by || '-') + '</td>' +
      '<td style="color:#888;font-size:.8rem;">' + esc(h.note || '-') + '</td>' +
      '</tr>';
  }).join('');
}

function openAddStockModal(editId) {
  var form = g('form-addStock');
  if (form) form.reset();
  // Populate item select
  var sel = g('addstock-item');
  if (sel) {
    sel.innerHTML = '<option value="">Select item</option>' +
      itemCatalogue.filter(function(x) { return x.status === 'active'; }).map(function(it) {
        return '<option value="' + esc(it.id) + '">' + esc(it.name) + '</option>';
      }).join('');
  }
  var editIdEl = g('addstock-edit-id');
  if (editIdEl) editIdEl.value = '';
  if (editId) {
    loadInvData();
    var entry = invStock.find(function(s) { return s.id === editId; });
    if (entry) {
      if (editIdEl) editIdEl.value = editId;
      if (sel) sel.value = entry.itemId;
      var qtyEl = g('addstock-qty'); if (qtyEl) qtyEl.value = entry.inStock || '';
      var noteEl = g('addstock-note'); if (noteEl) noteEl.value = entry.note || '';
    }
  }
  openModal('modal-addStock');
}

function submitAddStock(e) {
  e.preventDefault();
  loadInvData();
  var itemId = rv('addstock-item');
  var qty = parseInt(rv('addstock-qty'), 10);
  var note = rv('addstock-note');
  if (!itemId) { showToast('Please select an item.', 'error'); return; }
  if (!qty || qty < 1) { showToast('Please enter a valid quantity.', 'error'); return; }
  var item = itemCatalogue.find(function(x) { return x.id === itemId; });
  if (!item) { showToast('Item not found.', 'error'); return; }
  var editId = rv('addstock-edit-id');
  var now = new Date().toISOString().slice(0, 10);
  var byUser = currentUser ? currentUser.name : currentRole;

  var existing = invStock.find(function(s) { return s.id === editId; }) ||
                 invStock.find(function(s) { return s.itemId === itemId; });

  if (existing && !editId) {
    // Add to existing stock
    var before = existing.inStock || 0;
    existing.inStock = before + qty;
    existing.lastUpdated = now;
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: before, after: existing.inStock, by: byUser, note: note || 'Stock added' });
  } else if (existing && editId) {
    var before2 = existing.inStock || 0;
    existing.inStock = qty;
    existing.lastUpdated = now;
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: before2, after: qty, by: byUser, note: note || 'Stock adjusted' });
  } else {
    invStock.push({ id: 'is' + Date.now(), itemId: itemId, itemName: item.name, category: item.category || '-', inStock: qty, allocated: 0, lastUpdated: now });
    invHistory.push({ id: 'ih' + Date.now(), date: now, itemId: itemId, itemName: item.name, type: 'in', qty: qty, before: 0, after: qty, by: byUser, note: note || 'Initial stock' });
  }
  saveInvData();
  closeModal('modal-addStock');
  showToast('Stock updated successfully!', 'success');
  renderInvSaleStock();
}

function deleteStockItem(id) {
  showConfirm('Are you sure you want to delete this stock entry?', function() {
    loadInvData();
    invStock = invStock.filter(function(s) { return s.id !== id; });
    saveInvData();
    showToast('Stock entry deleted.', 'success');
    renderInvSaleStock();
  }, 'Delete Stock Entry', 'Delete', true);
}

function openAllocateModal() {
  var form = g('form-allocate');
  if (form) form.reset();
  loadInvData();
  var sel = g('allocate-item');
  if (sel) {
    sel.innerHTML = '<option value="">Select item</option>' +
      invStock.map(function(s) {
        var avail = Math.max(0, (s.inStock || 0) - (s.allocated || 0));
        return '<option value="' + esc(s.itemId) + '">' + esc(s.itemName) + ' (Available: ' + avail + ')</option>';
      }).join('');
  }
  var infoEl = g('allocate-available-info');
  if (infoEl) infoEl.style.display = 'none';
  openModal('modal-allocate');
}

function updateAllocateAvailable() {
  var itemId = g('allocate-item') ? g('allocate-item').value : '';
  var infoEl = g('allocate-available-info');
  var qtyEl = g('allocate-avail-qty');
  if (!itemId) { if (infoEl) infoEl.style.display = 'none'; return; }
  var avail = getAvailableQty(itemId);
  if (infoEl) infoEl.style.display = '';
  if (qtyEl) qtyEl.textContent = avail;
}

function submitAllocationRequest(e) {
  e.preventDefault();
  loadInvData();
  var itemId = rv('allocate-item');
  var qty = parseInt(rv('allocate-qty'), 10);
  var purpose = rv('allocate-purpose');
  if (!itemId) { showToast('Please select an item.', 'error'); return; }
  if (!qty || qty < 1) { showToast('Please enter a valid quantity.', 'error'); return; }

  // Verify stock exists
  var stockEntry = invStock.find(function(s) { return s.itemId === itemId; });
  if (!stockEntry) { showToast('No stock found for this item.', 'error'); return; }

  // Verify available quantity
  var avail = getAvailableQty(itemId);
  if (qty > avail) {
    showToast('Requested quantity (' + qty + ') exceeds available stock (' + avail + ').', 'error');
    return;
  }

  var now = new Date().toISOString().slice(0, 10);
  var byUser = currentUser ? currentUser.name : currentRole;
  invRequests.push({
    id: 'ir' + Date.now(),
    date: now,
    requestedBy: byUser,
    itemId: itemId,
    itemName: stockEntry.itemName,
    qty: qty,
    purpose: purpose,
    status: 'pending'
  });
  saveInvData();
  closeModal('modal-allocate');
  showToast('Allocation request submitted! Awaiting supervisor approval.', 'success');
  renderInvRequestsTable();
  renderInvSaleStock();
}

function openReviewAllocModal(requestId) {
  loadInvData();
  var req = invRequests.find(function(r) { return r.id === requestId; });
  if (!req) return;
  _reviewAllocId = requestId;
  var avail = getAvailableQty(req.itemId);
  var body = g('review-alloc-body');
  if (body) {
    var match = req.itemName && req.qty;
    body.innerHTML =
      '<div class="form-group">' +
        '<div style="background:#F5F5F5;border-radius:10px;padding:16px;">' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Request ID</span>' +
            '<span style="font-size:.8125rem;font-weight:600;">' + esc(req.id) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Requested By</span>' +
            '<span style="font-size:.8125rem;font-weight:600;">' + esc(req.requestedBy) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Date</span>' +
            '<span style="font-size:.8125rem;">' + esc(req.date) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;align-items:center;">' +
            '<span style="font-size:.8125rem;color:#888;">Item Name</span>' +
            '<span style="font-size:.875rem;font-weight:700;color:#1A1A2E;">' + esc(req.itemName) + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;align-items:center;">' +
            '<span style="font-size:.8125rem;color:#888;">Qty Requested</span>' +
            '<span style="font-size:1rem;font-weight:800;color:#1565C0;">' + req.qty + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:10px;">' +
            '<span style="font-size:.8125rem;color:#888;">Available in Stock</span>' +
            '<span style="font-size:.875rem;font-weight:700;color:' + (avail >= req.qty ? '#2E7D32' : '#C62828') + ';">' + avail + '</span>' +
          '</div>' +
          '<div style="display:flex;justify-content:space-between;margin-bottom:4px;">' +
            '<span style="font-size:.8125rem;color:#888;">Purpose</span>' +
            '<span style="font-size:.8125rem;max-width:55%;text-align:right;">' + esc(req.purpose) + '</span>' +
          '</div>' +
        '</div>' +
      '</div>' +
      (avail < req.qty ? '<div style="background:#FFEBEE;border-radius:8px;padding:10px 14px;font-size:.8125rem;color:#C62828;margin-top:8px;"><i class="fas fa-triangle-exclamation"></i> <strong>Warning:</strong> Available stock (' + avail + ') is less than requested (' + req.qty + '). Approval will not be possible.</div>' : '') +
      '<div class="form-group" style="margin-top:12px;">' +
        '<label class="form-label">Review Note (optional)</label>' +
        '<input type="text" class="form-input" id="review-note-input" placeholder="e.g., Approved for weekly allocation" />' +
      '</div>';
  }
  // Disable approve if not enough stock
  var approveBtn = g('review-approve-btn');
  if (approveBtn) approveBtn.disabled = avail < req.qty;
  openModal('modal-reviewAlloc');
}

function processAllocation(action) {
  loadInvData();
  var req = invRequests.find(function(r) { return r.id === _reviewAllocId; });
  if (!req) return;
  var reviewNote = g('review-note-input') ? g('review-note-input').value.trim() : '';
  var now = new Date().toISOString().slice(0, 10);
  var byUser = currentUser ? currentUser.name : currentRole;

  if (action === 'approved') {
    // Verify item name and amount one more time
    var stockEntry = invStock.find(function(s) { return s.itemId === req.itemId; });
    if (!stockEntry) { showToast('Stock entry not found. Cannot approve.', 'error'); return; }
    var avail = getAvailableQty(req.itemId);
    if (req.qty > avail) {
      showToast('Insufficient stock (' + avail + ' available, ' + req.qty + ' requested). Cannot approve.', 'error');
      return;
    }
    // Auto-calculate: deduct from available (add to allocated)
    var before = stockEntry.inStock || 0;
    stockEntry.allocated = (stockEntry.allocated || 0) + req.qty;
    invHistory.push({
      id: 'ih' + Date.now(),
      date: now,
      itemId: req.itemId,
      itemName: req.itemName,
      type: 'out',
      qty: req.qty,
      before: before,
      after: before,  // inStock doesn't change, allocated increases
      by: byUser,
      note: 'Allocated: ' + (reviewNote || req.purpose)
    });
    req.status = 'approved';
    req.reviewedBy = byUser;
    req.reviewNote = reviewNote;
    req.reviewedAt = now;
    saveInvData();
    closeModal('modal-reviewAlloc');
    showToast('Allocation approved! ' + req.qty + ' units of ' + req.itemName + ' allocated.', 'success');
    showApprovalFormModal('allocation', req);
  } else {
    req.status = 'rejected';
    req.reviewedBy = byUser;
    req.reviewNote = reviewNote;
    saveInvData();
    closeModal('modal-reviewAlloc');
    showToast('Allocation request rejected.', 'info');
  }
  _reviewAllocId = null;
  renderInvRequestsTable();
  renderInvSaleStock();
}

function renderShopStockTable() {
  loadInvData();
  var shopFilterEl = g('shop-stock-branch-filter');
  // Auto-apply branch filter for supervisor and agent roles
  if (shopFilterEl && currentUser && (currentRole === 'supervisor' || currentRole === 'agent') && !shopFilterEl.value) {
    shopFilterEl.value = currentUser.branch || '';
  }
  var branchFilter = shopFilterEl ? shopFilterEl.value : '';
  var tbody = g('shop-stock-table');
  if (!tbody) return;

  // Build shop stock from approved allocations
  var approved = invRequests.filter(function(r) {
    return r.status === 'approved';
  });

  if (!approved.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#999;"><i class="fas fa-store fa-2x" style="display:block;margin-bottom:10px;"></i>No allocated stock for shops yet.</td></tr>';
    return;
  }

  tbody.innerHTML = approved.filter(function(r) {
    return !branchFilter || (r.purpose && r.purpose.toLowerCase().indexOf(branchFilter.toLowerCase()) !== -1);
  }).map(function(r, idx) {
    var stockEntry = invStock.find(function(s) { return s.itemId === r.itemId; });
    var remaining = r.qty;  // simplified: show full allocated qty
    var statusBadge = '<span class="inv-badge-status inv-badge-approved">Active</span>';
    return '<tr>' +
      '<td>' + (idx + 1) + '</td>' +
      '<td>' + esc(r.purpose || '-') + '</td>' +
      '<td style="font-weight:600;">' + esc(r.itemName) + '</td>' +
      '<td style="font-weight:700;color:#1565C0;">' + (r.qty || 0) + '</td>' +
      '<td>0</td>' +
      '<td style="font-weight:700;color:#1B7D3D;">' + remaining + '</td>' +
      '<td>' + esc(r.date || '-') + '</td>' +
      '<td>' + statusBadge + '</td>' +
      '</tr>';
  }).join('');
}

// ------------------------------------------------------------
// Approval Form — PDF & Email
// ------------------------------------------------------------

function initApprovalSignaturePad() {
  var oldCanvas = g('approval-sig-canvas');
  if (!oldCanvas || !oldCanvas.parentNode) return;
  // Replace canvas with a fresh clone to remove any stale event listeners
  var canvas = oldCanvas.cloneNode(false);
  oldCanvas.parentNode.replaceChild(canvas, oldCanvas);
  _sigCanvas = canvas;
  _sigCtx = canvas.getContext('2d');
  _sigCtx.strokeStyle = '#1A1A2E';
  _sigCtx.lineWidth = 2;
  _sigCtx.lineCap = 'round';
  _sigCtx.lineJoin = 'round';

  function getPos(e) {
    var rect = canvas.getBoundingClientRect();
    var scaleX = canvas.width / rect.width;
    var scaleY = canvas.height / rect.height;
    var src = e.touches ? e.touches[0] : e;
    return { x: (src.clientX - rect.left) * scaleX, y: (src.clientY - rect.top) * scaleY };
  }

  canvas.addEventListener('mousedown', function(e) { _sigDrawing = true; var p = getPos(e); _sigCtx.beginPath(); _sigCtx.moveTo(p.x, p.y); });
  canvas.addEventListener('mousemove', function(e) { if (!_sigDrawing) return; var p = getPos(e); _sigCtx.lineTo(p.x, p.y); _sigCtx.stroke(); });
  canvas.addEventListener('mouseup', function() { _sigDrawing = false; });
  canvas.addEventListener('mouseleave', function() { _sigDrawing = false; });
  canvas.addEventListener('touchstart', function(e) { e.preventDefault(); _sigDrawing = true; var p = getPos(e); _sigCtx.beginPath(); _sigCtx.moveTo(p.x, p.y); }, { passive: false });
  canvas.addEventListener('touchmove', function(e) { e.preventDefault(); if (!_sigDrawing) return; var p = getPos(e); _sigCtx.lineTo(p.x, p.y); _sigCtx.stroke(); }, { passive: false });
  canvas.addEventListener('touchend', function() { _sigDrawing = false; });
}

function clearApprovalSignature() {
  if (_sigCtx && _sigCanvas) {
    _sigCtx.clearRect(0, 0, _sigCanvas.width, _sigCanvas.height);
  }
}

function showApprovalFormModal(type, data) {
  _approvalFormData = { type: type, data: data };
  var titleEl = g('approval-form-title');
  var contentEl = g('approval-form-content');
  var submitterEl = g('approval-submitter-name');
  var approverEl = g('approval-approver-name');

  var v = _getApprovalFormVars(type, data);
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-file-contract" style="color:#1B7D3D;margin-right:8px;"></i>' + v.formTitle;

  var detailRows = '';
  if (v.isDeposit) {
    detailRows =
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Cash Amount</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1B7D3D;">$' + (data.cash ? Number(data.cash).toFixed(2) : '0.00') + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Credit Amount</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1565C0;">$' + (data.credit ? Number(data.credit).toFixed(2) : '0.00') + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Total Amount</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;color:#E65100;">$' + ((Number(data.cash)||0) + (Number(data.credit)||0)).toFixed(2) + '</td></tr>';
  } else {
    detailRows =
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Item Name</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;">' + esc(data.itemName || '') + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Quantity Allocated</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:700;color:#1565C0;">' + (data.qty || 0) + '</td></tr>' +
      '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Purpose</td>' +
        '<td style="padding:7px 12px;border-bottom:1px solid #eee;">' + esc(data.purpose || '') + '</td></tr>';
  }

  if (contentEl) {
    contentEl.innerHTML =
      '<div style="border:2px solid #1B7D3D;border-radius:8px;overflow:hidden;">' +
        '<div style="background:#1B7D3D;color:#fff;padding:12px 16px;text-align:center;">' +
          '<div style="font-size:1.1rem;font-weight:700;letter-spacing:.5px;">SMART 5G DASHBOARD</div>' +
          '<div style="font-size:.85rem;opacity:.9;margin-top:2px;">' + v.formTitle.toUpperCase() + '</div>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;">' +
          '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;width:45%;">Submitter Name</td>' +
            '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:600;">' + esc(v.submitter) + '</td></tr>' +
          (v.branch ? '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Branch</td><td style="padding:7px 12px;border-bottom:1px solid #eee;">' + esc(v.branch) + '</td></tr>' : '') +
          '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Date Submitted</td>' +
            '<td style="padding:7px 12px;border-bottom:1px solid #eee;">' + esc(v.dateSubmitted) + '</td></tr>' +
          detailRows +
          (v.remark ? '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Remark / Note</td><td style="padding:7px 12px;border-bottom:1px solid #eee;">' + esc(v.remark) + '</td></tr>' : '') +
          '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;border-bottom:1px solid #eee;">Approved By</td>' +
            '<td style="padding:7px 12px;border-bottom:1px solid #eee;font-weight:600;color:#1B7D3D;">' + esc(v.approver) + '</td></tr>' +
          '<tr><td style="padding:7px 12px;font-weight:600;color:#555;background:#f5f7fa;">Date Approved</td>' +
            '<td style="padding:7px 12px;font-weight:600;">' + esc(v.dateApproved) + '</td></tr>' +
        '</table>' +
      '</div>';
  }

  if (submitterEl) submitterEl.textContent = v.submitter;
  if (approverEl) approverEl.textContent = v.approver;

  // Clear previous signature
  clearApprovalSignature();
  openModal('modal-approvalForm');

  // Init signature pad (after modal is visible)
  setTimeout(initApprovalSignaturePad, 50);
}

function _escHtml(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function _getApprovalFormVars(type, data) {
  var isDeposit = (type === 'deposit');
  return {
    isDeposit: isDeposit,
    formTitle: isDeposit ? 'Daily Cash & Credit Deposit Approval' : 'Stock Allocation Approval',
    submitter: isDeposit ? (data.agent || '') : (data.requestedBy || ''),
    approver:  isDeposit ? (data.approvedBy || '') : (data.reviewedBy || ''),
    dateSubmitted: data.date || '',
    dateApproved:  isDeposit ? (data.approvedAt || '') : (data.reviewedAt || ''),
    branch: isDeposit ? (data.branch || '') : '',
    remark: isDeposit ? (data.remark || '') : (data.reviewNote || data.purpose || '')
  };
}

function printApprovalForm() {
  if (!_approvalFormData) return;
  var v = _getApprovalFormVars(_approvalFormData.type, _approvalFormData.data);
  var data = _approvalFormData.data;
  var sigDataUrl = (_sigCanvas && _sigCanvas.getContext) ? _sigCanvas.toDataURL('image/png') : '';

  var detailHtml = '';
  if (v.isDeposit) {
    detailHtml =
      '<tr><td class="lbl">Cash Amount</td><td class="val" style="color:#1B7D3D;">$' + (data.cash ? Number(data.cash).toFixed(2) : '0.00') + '</td></tr>' +
      '<tr><td class="lbl">Credit Amount</td><td class="val" style="color:#1565C0;">$' + (data.credit ? Number(data.credit).toFixed(2) : '0.00') + '</td></tr>' +
      '<tr><td class="lbl">Total Amount</td><td class="val" style="font-weight:700;color:#E65100;">$' + ((Number(data.cash)||0) + (Number(data.credit)||0)).toFixed(2) + '</td></tr>';
  } else {
    detailHtml =
      '<tr><td class="lbl">Item Name</td><td class="val">' + _escHtml(data.itemName) + '</td></tr>' +
      '<tr><td class="lbl">Quantity Allocated</td><td class="val" style="color:#1565C0;font-weight:700;">' + _escHtml(String(data.qty || 0)) + '</td></tr>' +
      '<tr><td class="lbl">Purpose</td><td class="val">' + _escHtml(data.purpose) + '</td></tr>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>' + _escHtml(v.formTitle) + '</title><style>' +
    'body{font-family:Arial,sans-serif;padding:32px;color:#1A1A2E;font-size:13px;max-width:700px;margin:0 auto;}' +
    '.header{background:#1B7D3D;color:#fff;padding:14px 20px;border-radius:8px 8px 0 0;text-align:center;}' +
    '.header h1{margin:0;font-size:1.1rem;letter-spacing:.5px;}' +
    '.header p{margin:3px 0 0;font-size:.8rem;opacity:.9;}' +
    'table{width:100%;border-collapse:collapse;border:2px solid #1B7D3D;border-top:none;border-radius:0 0 8px 8px;overflow:hidden;}' +
    'td{padding:8px 14px;border-bottom:1px solid #e8e8e8;}' +
    '.lbl{background:#f5f7fa;font-weight:600;color:#555;width:45%;}' +
    '.val{font-weight:500;}' +
    '.sig-section{display:flex;gap:32px;margin-top:28px;}' +
    '.sig-box{flex:1;}' +
    '.sig-box p{font-weight:600;font-size:.85rem;color:#555;margin:0 0 6px;}' +
    '.sig-line{border-bottom:2px solid #333;height:60px;margin-bottom:4px;}' +
    '.sig-name{font-size:.75rem;color:#333;font-style:italic;}' +
    '.sig-img{max-width:100%;height:60px;object-fit:contain;display:block;border:1px solid #ddd;border-radius:4px;background:#fafafa;}' +
    '.footer{margin-top:20px;text-align:center;font-size:.7rem;color:#999;border-top:1px solid #eee;padding-top:10px;}' +
    '@media print{body{padding:16px;} button{display:none!important;} .no-print{display:none!important;}}' +
    '</style></head><body>' +
    '<div class="no-print" style="margin-bottom:16px;text-align:right;">' +
      '<button onclick="window.print()" style="background:#1B7D3D;color:#fff;border:none;padding:9px 18px;border-radius:6px;cursor:pointer;font-size:.875rem;">\uD83D\uDDB6 Print / Save as PDF</button>' +
    '</div>' +
    '<div class="header"><h1>SMART 5G DASHBOARD</h1><p>' + _escHtml(v.formTitle).toUpperCase() + '</p></div>' +
    '<table>' +
      '<tr><td class="lbl">Submitter Name</td><td class="val">' + _escHtml(v.submitter) + '</td></tr>' +
      (v.branch ? '<tr><td class="lbl">Branch</td><td class="val">' + _escHtml(v.branch) + '</td></tr>' : '') +
      '<tr><td class="lbl">Date Submitted</td><td class="val">' + _escHtml(v.dateSubmitted) + '</td></tr>' +
      detailHtml +
      (v.remark ? '<tr><td class="lbl">Remark / Note</td><td class="val">' + _escHtml(v.remark) + '</td></tr>' : '') +
      '<tr><td class="lbl">Approved By</td><td class="val" style="color:#1B7D3D;font-weight:600;">' + _escHtml(v.approver) + '</td></tr>' +
      '<tr><td class="lbl">Date Approved</td><td class="val">' + _escHtml(v.dateApproved) + '</td></tr>' +
    '</table>' +
    '<div class="sig-section">' +
      '<div class="sig-box"><p>Submitter Signature:</p><div class="sig-line"></div><div class="sig-name">' + _escHtml(v.submitter) + '</div></div>' +
      '<div class="sig-box"><p>Supervisor / Approver Signature:</p>' +
        (sigDataUrl ? '<img src="' + _escHtml(sigDataUrl) + '" class="sig-img" alt="Signature" />' : '<div class="sig-line"></div>') +
        '<div class="sig-name">' + _escHtml(v.approver) + '</div></div>' +
    '</div>' +
    '<div class="footer">Generated by Smart 5G Dashboard &bull; ' + new Date().toLocaleString() + '</div>' +
    '</body></html>';

  var win = window.open('', '_blank', 'width=780,height=700,scrollbars=yes');
  if (win) {
    win.document.open();
    win.document.write(html);
    win.document.close();
  }
}

function emailApprovalForm() {
  if (!_approvalFormData) return;
  var v = _getApprovalFormVars(_approvalFormData.type, _approvalFormData.data);
  var data = _approvalFormData.data;

  // Look up submitter's email from staffList by name (case-insensitive, trimmed)
  var submitterNameNorm = (v.submitter || '').trim().toLowerCase();
  var submitterUser = staffList.find(function(u) { return (u.name || '').trim().toLowerCase() === submitterNameNorm; });
  var submitterEmail = (submitterUser && submitterUser.email) ? submitterUser.email : '';
  if (!submitterEmail) {
    showToast('No email on file for ' + (v.submitter || 'submitter') + '. Add an email in Settings \u2192 Permission to auto-fill the recipient.', 'warning');
  }

  var subject = 'Smart 5G \u2014 ' + v.formTitle + ' \u2014 ' + v.dateApproved;

  var bodyLines = [
    'SMART 5G DASHBOARD',
    v.formTitle.toUpperCase(),
    '========================================',
    '',
    'Submitter Name  : ' + v.submitter,
  ];
  if (v.branch) bodyLines.push('Branch          : ' + v.branch);
  bodyLines.push('Date Submitted  : ' + v.dateSubmitted);

  if (v.isDeposit) {
    bodyLines.push('Cash Amount     : $' + (data.cash ? Number(data.cash).toFixed(2) : '0.00'));
    bodyLines.push('Credit Amount   : $' + (data.credit ? Number(data.credit).toFixed(2) : '0.00'));
    bodyLines.push('Total Amount    : $' + ((Number(data.cash)||0) + (Number(data.credit)||0)).toFixed(2));
  } else {
    bodyLines.push('Item Name       : ' + (data.itemName || ''));
    bodyLines.push('Quantity        : ' + (data.qty || 0));
    bodyLines.push('Purpose         : ' + (data.purpose || ''));
  }

  if (v.remark) bodyLines.push('Remark / Note   : ' + v.remark);
  bodyLines.push('Approved By     : ' + v.approver);
  bodyLines.push('Date Approved   : ' + v.dateApproved);
  bodyLines.push('');
  bodyLines.push('---');
  bodyLines.push('Note: Please print the approval form (use "Print / Save PDF" button), sign it, and attach to this email for records.');
  bodyLines.push('');
  bodyLines.push('Generated by Smart 5G Dashboard \u2014 ' + new Date().toLocaleString());

  var mailtoLink = 'mailto:' + encodeURIComponent(submitterEmail) + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(bodyLines.join('\r\n'));
  window.open(mailtoLink);
}

// ------------------------------------------------------------
// Init
// ------------------------------------------------------------
document.addEventListener('DOMContentLoaded', function() {
  loadAllData();
  loadInvData();
  filteredSales = saleRecords.slice();
  populateBranchSelects();
  renderItemChips();
  renderNewCustomerTable();
  renderTopUpTable();
  renderTerminationTable();
  renderStaffTable();
  renderKpiTable();
  renderPromotionCards();
  renderDepositTable();
  updateDepositKpis();
  renderSaleTable();
  updateSaleKpis();

  // Restore login session if user was previously authenticated
  var savedSession = lsLoad(LS_KEYS.session, null);
  if (savedSession && savedSession.username) {
    var roleMap = { 'Admin': 'admin', 'Cluster': 'cluster', 'Supervisor': 'supervisor', 'Agent': 'agent' };
    currentUser = savedSession;
    var ls = g('login-screen'); if (ls) ls.style.display = 'none';
    var as = g('app-shell'); if (as) as.style.display = 'flex';
    switchRole(roleMap[savedSession.role] || 'user');
    currentUser = savedSession;
    var nameEl = g('topbar-name'); if (nameEl) nameEl.textContent = savedSession.name;
    var avatarEl = g('topbar-avatar'); if (avatarEl) avatarEl.textContent = ini(savedSession.name);
    navigateTo('dashboard', null);
  }

  // Populate shop stock branch filter
  var shopFilter = g('shop-stock-branch-filter');
  if (shopFilter) {
    BRANCHES.forEach(function(b) {
      var opt = document.createElement('option');
      opt.value = b; opt.textContent = b;
      shopFilter.appendChild(opt);
    });
  }
});

document.addEventListener('click', function(e) {
  var panel = g('notif-panel');
  var btn = g('notif-btn');
  if (panel && panel.style.display === 'block') {
    if (!panel.contains(e.target) && (!btn || !btn.contains(e.target))) {
      panel.style.display = 'none';
    }
  }
});

// ============================================================
// Coverage Feature
// ============================================================

var COV_TABS = [
  { id: 'smart-home',    field: 'smartHome',    color: '#1B7D3D', mapVar: '_covMapSmartHome' },
  { id: 'smart-home-5g', field: 'smartHome5G',  color: '#1565C0', mapVar: '_covMapSmartHome5G' },
  { id: 'smart-fiber',   field: 'smartFiber',   color: '#FB8C00', mapVar: '_covMapSmartFiber' }
];

var KH_CENTER = [12.5657, 104.9910]; // Cambodia center

function switchCoverageTab(tab) {
  currentCoverageTab = tab;
  $$('.cov-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('cov-content-' + tab);
  if (tc) tc.classList.add('active');

  // Invalidate the map so it renders correctly when becoming visible
  setTimeout(function() {
    if (typeof L !== 'undefined') {
      if (tab === 'smart-home' && _covMapSmartHome) { _covMapSmartHome.invalidateSize(); }
      if (tab === 'smart-home-5g' && _covMapSmartHome5G) { _covMapSmartHome5G.invalidateSize(); }
      if (tab === 'smart-fiber' && _covMapSmartFiber) { _covMapSmartFiber.invalidateSize(); }
    }
    renderCoverageMap(tab);
    renderCoverageTable(tab);
  }, 50);
}

// Called from sidebar nav — shows only the selected coverage type (tab header hidden)
function openCoverageFromNav(tab, navItem) {
  navigateTo('coverage', null);
  setActiveSubItem(navItem);
  var tabHeader = document.querySelector('#page-coverage .tab-header');
  if (tabHeader) tabHeader.style.display = 'none';
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  var activeBtn = g('cov-tab-' + tab);
  if (activeBtn) activeBtn.classList.add('active');
  switchCoverageTab(tab);
}

// Called from the tab buttons inside the page — shows the tab header
function openCoverageTab(tab, btn) {
  var tabHeader = document.querySelector('#page-coverage .tab-header');
  if (tabHeader) tabHeader.style.display = '';
  switchCoverageTab(tab);
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}

function initCoveragePage() {
  // Show Add button for admin/cluster only
  var addBtn = g('cov-add-btn');
  if (addBtn) addBtn.style.display = (currentRole === 'admin' || currentRole === 'cluster') ? '' : 'none';

  // Show/hide Actions column based on role
  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');
  $$('.cov-admin-col').forEach(function(el) { el.style.display = isAdmin ? '' : 'none'; });

  // Initialize maps if not already done
  initCoverageMap('smart-home');
  initCoverageMap('smart-home-5g');
  initCoverageMap('smart-fiber');

  // Render tables
  renderCoverageTable('smart-home');
  renderCoverageTable('smart-home-5g');
  renderCoverageTable('smart-fiber');

  // Activate current tab button
  $$('#page-coverage .tab-btn').forEach(function(b) { b.classList.remove('active'); });
  var activeBtn = g('cov-tab-' + currentCoverageTab);
  if (activeBtn) activeBtn.classList.add('active');

  // Show/hide correct tab content
  $$('.cov-tab-content').forEach(function(c) { c.classList.remove('active'); });
  var tc = g('cov-content-' + currentCoverageTab);
  if (tc) tc.classList.add('active');

  setTimeout(function() {
    if (typeof L !== 'undefined') {
      if (_covMapSmartHome) _covMapSmartHome.invalidateSize();
      if (_covMapSmartHome5G) _covMapSmartHome5G.invalidateSize();
      if (_covMapSmartFiber) _covMapSmartFiber.invalidateSize();
    }
    renderCoverageMap('smart-home');
    renderCoverageMap('smart-home-5g');
    renderCoverageMap('smart-fiber');
  }, 100);
}

function initCoverageMap(tab) {
  if (typeof L === 'undefined') return; // Leaflet not loaded
  var mapId = 'cov-map-' + tab;
  var el = g(mapId);
  if (!el) return;

  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  // Destroy existing map instance if any
  if (tab === 'smart-home' && _covMapSmartHome) { _covMapSmartHome.remove(); _covMapSmartHome = null; }
  if (tab === 'smart-home-5g' && _covMapSmartHome5G) { _covMapSmartHome5G.remove(); _covMapSmartHome5G = null; }
  if (tab === 'smart-fiber' && _covMapSmartFiber) { _covMapSmartFiber.remove(); _covMapSmartFiber = null; }

  var map = L.map(mapId).setView(KH_CENTER, 7);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 18,
    attribution: '© OpenStreetMap contributors'
  }).addTo(map);

  if (tab === 'smart-home') _covMapSmartHome = map;
  if (tab === 'smart-home-5g') _covMapSmartHome5G = map;
  if (tab === 'smart-fiber') _covMapSmartFiber = map;
}

function getMapForTab(tab) {
  if (tab === 'smart-home') return _covMapSmartHome;
  if (tab === 'smart-home-5g') return _covMapSmartHome5G;
  if (tab === 'smart-fiber') return _covMapSmartFiber;
  return null;
}

function renderCoverageMap(tab) {
  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  // Update count badge regardless of Leaflet availability
  var locs = coverageLocations.filter(function(loc) { return loc[tabInfo.field]; });
  var badge = g('cov-count-' + tab);
  if (badge) badge.textContent = locs.length + ' location' + (locs.length !== 1 ? 's' : '');

  if (typeof L === 'undefined') return; // Leaflet not loaded
  var map = getMapForTab(tab);
  if (!map) return;

  // Remove existing layers (except tile layer)
  map.eachLayer(function(layer) {
    if (layer instanceof L.CircleMarker || layer instanceof L.Marker) {
      map.removeLayer(layer);
    }
  });

  locs.forEach(function(loc) {
    if (!loc.lat || !loc.lng) return;
    var circle = L.circleMarker([loc.lat, loc.lng], {
      radius: 10,
      fillColor: tabInfo.color,
      color: '#fff',
      weight: 2,
      opacity: 1,
      fillOpacity: 0.85
    }).addTo(map);
    circle.bindPopup(
      '<strong>' + esc(loc.commune) + '</strong><br/>' +
      '<small>' + esc(loc.district) + ', ' + esc(loc.province) + '</small>'
    );
  });

}

function renderCoverageTable(tab) {
  var tbody = g('cov-table-' + tab);
  if (!tbody) return;

  var tabInfo = COV_TABS.find(function(t) { return t.id === tab; });
  if (!tabInfo) return;

  var searchEl = g('cov-search-' + tab);
  var q = searchEl ? searchEl.value.trim().toLowerCase() : '';

  var locs = coverageLocations.filter(function(loc) { return loc[tabInfo.field]; });
  if (q) {
    locs = locs.filter(function(loc) {
      return (loc.commune || '').toLowerCase().includes(q) ||
             (loc.district || '').toLowerCase().includes(q) ||
             (loc.province || '').toLowerCase().includes(q);
    });
  }

  var isAdmin = (currentRole === 'admin' || currentRole === 'cluster');

  if (!locs.length) {
    tbody.innerHTML = '<tr><td colspan="' + (isAdmin ? 6 : 5) + '" style="text-align:center;padding:40px;color:#999;"><i class="fas fa-map" style="font-size:2rem;display:block;margin-bottom:8px;"></i>No coverage locations found</td></tr>';
    return;
  }

  tbody.innerHTML = locs.map(function(loc, i) {
    return '<tr>' +
      '<td>' + (i + 1) + '</td>' +
      '<td>' + esc(loc.commune) + '</td>' +
      '<td>' + esc(loc.district) + '</td>' +
      '<td>' + esc(loc.province) + '</td>' +
      '<td>' + esc(loc.date || '') + '</td>' +
      (isAdmin ? '<td><button class="btn-edit" onclick="editCoverageLocation(\'' + loc.id + '\')" title="Edit"><i class="fas fa-pen"></i></button> <button class="btn-delete" onclick="deleteCoverageLocation(\'' + loc.id + '\')" title="Delete"><i class="fas fa-trash-can"></i></button></td>' : '<td style="display:none;"></td>') +
      '</tr>';
  }).join('');
}

function openAddCoverageModal() {
  var f = g('form-coverage');
  if (f) f.reset();
  var titleEl = g('modal-coverage-title');
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-map-location-dot" style="color:#1B7D3D;margin-right:8px;"></i>Add Coverage Location';
  if (f) f.removeAttribute('data-edit-id');

  openModal('modal-coverage');

  setTimeout(function() {
    if (typeof L === 'undefined') return; // Leaflet not loaded
    if (_covPickerMap) { _covPickerMap.remove(); _covPickerMap = null; _covPickerMarker = null; _covPickerHighlight = null; }
    var pickerEl = g('cov-picker-map');
    if (!pickerEl) return;
    _covPickerMap = L.map('cov-picker-map').setView(KH_CENTER, 7);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 18,
      attribution: '© OpenStreetMap contributors'
    }).addTo(_covPickerMap);
    _covPickerMap.on('click', function(e) {
      var lat = parseFloat(e.latlng.lat.toFixed(6));
      var lng = parseFloat(e.latlng.lng.toFixed(6));
      var latEl = g('cov-lat'); var lngEl = g('cov-lng');
      if (latEl) latEl.value = lat;
      if (lngEl) lngEl.value = lng;
      if (_covPickerMarker) _covPickerMap.removeLayer(_covPickerMarker);
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    });
  }, 200);
}

function editCoverageLocation(id) {
  var loc = coverageLocations.find(function(l) { return l.id === id; });
  if (!loc) return;

  var titleEl = g('modal-coverage-title');
  if (titleEl) titleEl.innerHTML = '<i class="fas fa-pen" style="color:#1B7D3D;margin-right:8px;"></i>Edit Coverage Location';

  var f = g('form-coverage');
  if (f) f.setAttribute('data-edit-id', id);

  var prov = g('cov-province'); if (prov) prov.value = loc.province || '';
  var dist = g('cov-district'); if (dist) dist.value = loc.district || '';
  var comm = g('cov-commune'); if (comm) comm.value = loc.commune || '';
  var latEl = g('cov-lat'); if (latEl) latEl.value = loc.lat || '';
  var lngEl = g('cov-lng'); if (lngEl) lngEl.value = loc.lng || '';
  var sh = g('cov-prod-smart-home'); if (sh) sh.checked = !!loc.smartHome;
  var sh5g = g('cov-prod-smart-home-5g'); if (sh5g) sh5g.checked = !!loc.smartHome5G;
  var sf = g('cov-prod-smart-fiber'); if (sf) sf.checked = !!loc.smartFiber;

  openModal('modal-coverage');

  setTimeout(function() {
    if (typeof L === 'undefined') return; // Leaflet not loaded
    if (_covPickerMap) { _covPickerMap.remove(); _covPickerMap = null; _covPickerMarker = null; _covPickerHighlight = null; }
    var pickerEl = g('cov-picker-map');
    if (!pickerEl) return;
    var center = (loc.lat && loc.lng) ? [loc.lat, loc.lng] : KH_CENTER;
    var zoom = (loc.lat && loc.lng) ? 14 : 7;
    _covPickerMap = L.map('cov-picker-map').setView(center, zoom);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 18,
      attribution: '© OpenStreetMap contributors'
    }).addTo(_covPickerMap);
    if (loc.lat && loc.lng) {
      _covPickerMarker = L.marker([loc.lat, loc.lng]).addTo(_covPickerMap);
    }
    _covPickerMap.on('click', function(e) {
      var lat = parseFloat(e.latlng.lat.toFixed(6));
      var lng = parseFloat(e.latlng.lng.toFixed(6));
      var latEl = g('cov-lat'); var lngEl = g('cov-lng');
      if (latEl) latEl.value = lat;
      if (lngEl) lngEl.value = lng;
      if (_covPickerMarker) _covPickerMap.removeLayer(_covPickerMarker);
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    });
  }, 200);
}

function submitCoverageLocation(e) {
  e.preventDefault();
  var province = (g('cov-province').value || '').trim();
  var district = (g('cov-district').value || '').trim();
  var commune  = (g('cov-commune').value || '').trim();
  var lat      = parseFloat(g('cov-lat').value) || null;
  var lng      = parseFloat(g('cov-lng').value) || null;
  var smartHome    = g('cov-prod-smart-home').checked;
  var smartHome5G  = g('cov-prod-smart-home-5g').checked;
  var smartFiber   = g('cov-prod-smart-fiber').checked;

  if (!province || !district || !commune) {
    showToast('Please fill in Province, District, and Commune.', 'error'); return;
  }
  if (!smartHome && !smartHome5G && !smartFiber) {
    showToast('Please select at least one product.', 'error'); return;
  }

  var today = new Date().toISOString().slice(0, 10);
  var f = g('form-coverage');
  var editId = f ? f.getAttribute('data-edit-id') : null;

  if (editId) {
    var idx = coverageLocations.findIndex(function(l) { return l.id === editId; });
    if (idx !== -1) {
      coverageLocations[idx] = Object.assign(coverageLocations[idx], {
        province: province, district: district, commune: commune,
        lat: lat, lng: lng,
        smartHome: smartHome, smartHome5G: smartHome5G, smartFiber: smartFiber,
        date: today
      });
    }
  } else {
    coverageLocations.push({
      id: 'cv_' + Date.now(),
      province: province, district: district, commune: commune,
      lat: lat, lng: lng,
      smartHome: smartHome, smartHome5G: smartHome5G, smartFiber: smartFiber,
      date: today
    });
  }

  lsSave(LS_KEYS.coverage, coverageLocations);
  closeModal('modal-coverage');
  showToast('Coverage location saved.', 'success');

  // Refresh all maps and tables
  renderCoverageMap('smart-home');
  renderCoverageMap('smart-home-5g');
  renderCoverageMap('smart-fiber');
  renderCoverageTable('smart-home');
  renderCoverageTable('smart-home-5g');
  renderCoverageTable('smart-fiber');
}

function deleteCoverageLocation(id) {
  showConfirm('Delete this coverage location?', function() {
    coverageLocations = coverageLocations.filter(function(l) { return l.id !== id; });
    lsSave(LS_KEYS.coverage, coverageLocations);
    showToast('Coverage location deleted.', 'success');
    renderCoverageMap('smart-home');
    renderCoverageMap('smart-home-5g');
    renderCoverageMap('smart-fiber');
    renderCoverageTable('smart-home');
    renderCoverageTable('smart-home-5g');
    renderCoverageTable('smart-fiber');
  });
}

// ------------------------------------------------------------
// Commune Autocomplete (Nominatim)
// ------------------------------------------------------------
function onCommuneInput(input) {
  var q = input.value.trim();
  var dropdown = g('cov-commune-suggestions');
  if (!dropdown) return;

  clearTimeout(_covCommuneDebounce);
  if (q.length < 2) { dropdown.style.display = 'none'; return; }

  _covCommuneDebounce = setTimeout(function() {
    var url = 'https://nominatim.openstreetmap.org/search?q=' +
      encodeURIComponent(q + ' Cambodia') +
      '&countrycodes=kh&format=json&addressdetails=1&limit=6';
    fetch(url)
      .then(function(r) { return r.json(); })
      .then(function(results) {
        if (!results || !results.length) { dropdown.style.display = 'none'; return; }
        dropdown.innerHTML = results.map(function(r, i) {
          var displayName = r.display_name || '';
          return '<div class="cov-suggestion-item" onmousedown="selectCommuneSuggestion(' + i + ')" data-idx="' + i + '">' +
            '<i class="fas fa-location-dot" style="color:#1B7D3D;margin-right:6px;flex-shrink:0;"></i>' +
            '<span>' + esc(displayName) + '</span>' +
            '</div>';
        }).join('');
        dropdown._results = results;
        dropdown.style.display = 'block';
      })
      .catch(function() { dropdown.style.display = 'none'; });
  }, 320);
}

function selectCommuneSuggestion(idx) {
  var dropdown = g('cov-commune-suggestions');
  if (!dropdown || !dropdown._results) return;
  var r = dropdown._results[idx];
  if (!r) return;
  dropdown.style.display = 'none';

  var addr = r.address || {};

  // Fill commune field
  var communeVal = addr.village || addr.commune || addr.suburb || addr.neighbourhood || addr.quarter || r.name || '';
  var communeEl = g('cov-commune');
  if (communeEl) communeEl.value = communeVal;

  // Fill district
  var districtVal = addr.county || addr.city_district || addr.district || addr.town || addr.city || '';
  var districtEl = g('cov-district');
  if (districtEl) districtEl.value = districtVal;

  // Fill province
  var provinceVal = addr.state || addr.province || '';
  var provinceEl = g('cov-province');
  if (provinceEl) provinceEl.value = provinceVal;

  // Fill lat/lng with centre
  var lat = parseFloat(r.lat);
  var lng = parseFloat(r.lon);
  if (!isNaN(lat) && !isNaN(lng)) {
    var latEl = g('cov-lat'); var lngEl = g('cov-lng');
    if (latEl) latEl.value = lat.toFixed(6);
    if (lngEl) lngEl.value = lng.toFixed(6);
  }

  // Highlight on picker map
  if (typeof L !== 'undefined' && _covPickerMap) {
    // Remove old highlight and marker
    if (_covPickerHighlight) { _covPickerMap.removeLayer(_covPickerHighlight); _covPickerHighlight = null; }
    if (_covPickerMarker)    { _covPickerMap.removeLayer(_covPickerMarker);    _covPickerMarker = null; }

    // Draw bounding box highlight
    if (r.boundingbox && r.boundingbox.length === 4) {
      var s = parseFloat(r.boundingbox[0]), n = parseFloat(r.boundingbox[1]);
      var w = parseFloat(r.boundingbox[2]), e = parseFloat(r.boundingbox[3]);
      if (!isNaN(s) && !isNaN(n) && !isNaN(w) && !isNaN(e)) {
        _covPickerHighlight = L.rectangle([[s, w], [n, e]], {
          color: '#1B7D3D', weight: 2, fillColor: '#1B7D3D', fillOpacity: 0.2
        }).addTo(_covPickerMap);
        _covPickerMap.fitBounds([[s, w], [n, e]], { padding: [30, 30] });
      } else if (!isNaN(lat) && !isNaN(lng)) {
        _covPickerMap.setView([lat, lng], 14);
      }
    } else if (!isNaN(lat) && !isNaN(lng)) {
      _covPickerMap.setView([lat, lng], 14);
    }

    // Add centre marker
    if (!isNaN(lat) && !isNaN(lng)) {
      _covPickerMarker = L.marker([lat, lng]).addTo(_covPickerMap);
    }
  }
}

function closeCommuneSuggestions() {
  var dropdown = g('cov-commune-suggestions');
  if (dropdown) dropdown.style.display = 'none';
}
