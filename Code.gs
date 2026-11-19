/*****************************************************************
 * STOCK MANAGEMENT SYSTEM — Code.gs
 * Database: Google Sheets | Backend: Apps Script | UI: Index/Style/JavaScript.html
 * FIRST RUN: run setupStockManagement() once.
 *****************************************************************/

// ======================== CONFIGURATION ========================
var CONFIG = {
  // CONFIGURATION: Enter your Spreadsheet ID here.
  // Leave '' if you opened Apps Script from the Sheet (Extensions → Apps Script).
  SPREADSHEET_ID: '',
  // CONFIGURATION: Timezone for all dates and "Submitted At" times.
  TIMEZONE: 'Asia/Dhaka',
  // CONFIGURATION: Products at or below this quantity are shown as Low Stock.
  LOW_STOCK_THRESHOLD: 10,
  APP_TITLE: 'Stock Manager'
};

var SHEET = {
  dashboard: 'Dashboard', products: 'Product List', IN: 'Stock In', OUT: 'Outbound Product',
  RET: 'Return Product', DMG: 'Damage Product', OTH: 'Others', ledger: 'Stock Ledger', summary: 'Stock Summary'
};
var TYPE = {
  IN:  { label: 'Stock In',  col: 'Stock In (+)', sign: 1 },
  OUT: { label: 'Outbound',  col: 'Outbound (-)', sign: -1 },
  RET: { label: 'Return',    col: 'Return (+)',   sign: 1 },
  DMG: { label: 'Damage',    col: 'Damage (-)',   sign: -1 },
  OTH: { label: 'Others',    col: 'Others (-)',   sign: -1 }
};
var TXN_HEAD = ['Transaction Date', 'Transaction ID', 'Product ID', 'Product Name', 'Quantity', 'Note', 'Submitted At', 'Updated At'];
var HEADERS = {
  products: ['Product ID', 'Product Name', 'Category', 'Opening Stock', 'Current Stock', 'Status', 'Created Date', 'Updated Date'],
  IN: TXN_HEAD, OUT: TXN_HEAD, RET: TXN_HEAD, DMG: TXN_HEAD,
  OTH: ['Transaction Date', 'Transaction ID', 'Product ID', 'Product Name', 'Quantity', 'Reason', 'Note', 'Submitted At', 'Updated At'],
  ledger: ['Transaction Date', 'Transaction ID', 'Product ID', 'Product Name', 'Transaction Type', 'Stock In (+)', 'Outbound (-)',
    'Return (+)', 'Damage (-)', 'Others (-)', 'Balance After Transaction', 'Note', 'Submitted At', 'Backdated', 'Updated At'],
  summary: ['Date', 'Product ID', 'Product Name', 'Previous Stock', 'Stock In', 'Outbound', 'Return', 'Damage', 'Others', 'Closing Stock']
};
var DATE_HEADS = ['Transaction Date', 'Submitted At', 'Updated At', 'Created Date', 'Updated Date', 'Date'];
var NUM_HEADS = ['Quantity', 'Opening Stock', 'Current Stock', 'Stock In (+)', 'Outbound (-)', 'Return (+)', 'Damage (-)', 'Others (-)',
  'Balance After Transaction', 'Previous Stock', 'Stock In', 'Outbound', 'Return', 'Damage', 'Others', 'Closing Stock'];

// ======================== WEB APP ENTRY ========================
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Stock System')
      .addItem('Run setup / re-apply formatting', 'setupStockManagement')
      .addItem('Reconcile stock', 'reconcileStockMenu').addToUi();
  } catch (e) {}
}

// ======================== BASIC HELPERS ========================
var _ss = null;
function ss_() { if (!_ss) _ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); return _ss; }
function sh_(key) {
  var s = ss_().getSheetByName(SHEET[key]);
  if (!s) fail_('Setup is not complete. Run setupStockManagement() first.');
  return s;
}
function fail_(m) { var e = new Error(m); e.name = 'UserError'; throw e; }
function now_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(v || '').trim().substring(0, 10); }
function dtstr_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss') : String(v || ''); }
function num_(v) { var n = Number(v); return isNaN(n) ? 0 : n; }
function addDays_(d, n) { var p = d.split('-'); return Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)), 'UTC', 'yyyy-MM-dd'); }
function vdate_(s) {
  s = String(s || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || addDays_(s, 0) !== s) fail_('Please choose a valid date.');
  return s;
}
function colLetter_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function headRow_(s) { var lc = s.getLastColumn(); return lc ? s.getRange(1, 1, 1, lc).getValues()[0].map(String) : []; }
function table_(s) {
  var head = headRow_(s), lr = s.getLastRow();
  return { head: head, values: (lr > 1 && head.length) ? s.getRange(2, 1, lr - 1, head.length).getValues() : [] };
}
function col_(head, name) { var i = head.indexOf(name); if (i < 0) fail_('Column "' + name + '" is missing. Run setupStockManagement() again.'); return i; }
function obj_(head, row) { var o = {}; head.forEach(function (h, i) { o[h] = row[i]; }); return o; }
function maxNum_(ids) { var m = 0; ids.forEach(function (i) { var x = /(\d+)$/.exec(String(i)); if (x) m = Math.max(m, Number(x[1])); }); return m; }
function nextId_(key, prefix, pad, existing) {
  var pr = PropertiesService.getScriptProperties();
  var n = Math.max(Number(pr.getProperty(key) || 0), existing) + 1;
  pr.setProperty(key, String(n));
  return prefix + ('000000000' + n).slice(-pad);
}
function locked_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_('The system is busy. Please try again in a moment.');
  try { return fn(); } finally { lock.releaseLock(); }
}
function api_(fn) {
  try { return { ok: true, data: fn() }; }
  catch (e) {
    if (e.name !== 'UserError') console.error(e.stack || e);
    return { ok: false, error: e.name === 'UserError' ? e.message : 'Unexpected error: ' + e.message };
  }
}

// Snapshot / rollback: every write operation snapshots the sheets it touches and restores them on failure.
function snap_(keys) {
  return keys.map(function (k) {
    var s = sh_(k), lr = s.getLastRow(), lc = s.getLastColumn();
    return { key: k, lc: lc, vals: lr > 1 ? s.getRange(2, 1, lr - 1, lc).getValues() : [] };
  });
}
function restore_(snaps) {
  snaps.forEach(function (sn) {
    var s = sh_(sn.key), lr = s.getLastRow();
    if (lr > 1) s.getRange(2, 1, lr - 1, sn.lc).clearContent();
    if (sn.vals.length) s.getRange(2, 1, sn.vals.length, sn.lc).setValues(sn.vals);
  });
}
function mutate_(keys, fn) {
  return locked_(function () {
    var sn = snap_(keys);
    try { return fn(); }
    catch (e) { try { restore_(sn); } catch (x) { console.error(x); } throw e; }
  });
}

// ======================== DATA LAYER ========================
function loadProducts_() {
  var t = table_(sh_('products')), out = [];
  t.values.forEach(function (r, i) {
    var o = obj_(t.head, r); if (!o['Product ID']) return;
    out.push({ row: i + 2, id: String(o['Product ID']), name: String(o['Product Name']), category: String(o['Category'] || ''),
      opening: num_(o['Opening Stock']), current: num_(o['Current Stock']), status: String(o['Status'] || 'Active'),
      created: dtstr_(o['Created Date']), updated: dtstr_(o['Updated Date']) });
  });
  return out;
}
function loadEvents_() {
  var t = table_(sh_('ledger')), res = [];
  t.values.forEach(function (r) {
    var o = obj_(t.head, r); if (!o['Transaction ID']) return;
    var type = null;
    for (var k in TYPE) if (String(o['Transaction Type']) === TYPE[k].label) type = k;
    if (!type) return;
    res.push({ date: dstr_(o['Transaction Date']), id: String(o['Transaction ID']), pid: String(o['Product ID']), type: type,
      qty: Math.abs(num_(o[TYPE[type].col])), note: String(o['Note'] || ''), submitted: dtstr_(o['Submitted At']), updated: dtstr_(o['Updated At']) });
  });
  return res;
}
// Chronological order by TRANSACTION DATE (never submission date). Inflows first inside a day, then submission time.
function sortEv_(ev) {
  ev.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    var sa = TYPE[a.type].sign, sb = TYPE[b.type].sign; if (sa !== sb) return sb - sa;
    if (a.submitted !== b.submitted) return a.submitted < b.submitted ? -1 : 1;
    return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0);
  });
}
// Replays the whole ledger from Opening Stock. watch = product IDs that must never go negative.
function replay_(events, products, watch) {
  sortEv_(events);
  var bal = {}, pm = {}, err = null;
  products.forEach(function (p) { bal[p.id] = p.opening; pm[p.id] = p; });
  events.forEach(function (e) {
    if (bal[e.pid] === undefined) bal[e.pid] = 0;
    bal[e.pid] += TYPE[e.type].sign * e.qty; e.balance = bal[e.pid];
    if (e.balance < 0 && !err && (!watch || watch.indexOf(e.pid) >= 0))
      err = 'Not enough stock: "' + (pm[e.pid] ? pm[e.pid].name : e.pid) + '" would drop to ' + e.balance + ' on ' + e.date +
        ' (transaction ' + e.id + '). Stock cannot be negative on any date.';
  });
  return { bal: bal, error: err };
}
function ctx_() {
  var products = loadProducts_(), events = loadEvents_(), r = replay_(events, products), pm = {}, by = {};
  products.forEach(function (p) { pm[p.id] = p; by[p.id] = []; });
  events.forEach(function (e) { (by[e.pid] = by[e.pid] || []).push(e); });
  return { products: products, events: events, bal: r.bal, error: r.error, pm: pm, by: by };
}
function sumTypes_(list) { var t = { IN: 0, OUT: 0, RET: 0, DMG: 0, OTH: 0 }; list.forEach(function (e) { t[e.type] += e.qty; }); return t; }

function replaceRows_(sheet, objs) {
  var head = headRow_(sheet), lr = sheet.getLastRow();
  if (lr > 1) sheet.getRange(2, 1, lr - 1, head.length).clearContent();
  if (objs.length) sheet.getRange(2, 1, objs.length, head.length).setValues(objs.map(function (o) { return head.map(function (h) { return o[h] === undefined ? '' : o[h]; }); }));
}
function writeProductRow_(p) {
  var s = sh_('products'), head = headRow_(s);
  var o = { 'Product ID': p.id, 'Product Name': p.name, 'Category': p.category, 'Opening Stock': p.opening, 'Current Stock': p.current,
    'Status': p.status, 'Created Date': p.created, 'Updated Date': p.updated };
  var cur = p.row <= s.getLastRow() ? s.getRange(p.row, 1, 1, head.length).getValues()[0] : head.map(function () { return ''; });
  s.getRange(p.row, 1, 1, head.length).setValues([head.map(function (h, i) { return o[h] === undefined ? cur[i] : o[h]; })]);
}
function summaryRows_(events, products) {
  var by = {}, rows = [];
  events.forEach(function (e) { (by[e.pid] = by[e.pid] || []).push(e); });
  products.forEach(function (p) {
    var prev = p.opening, days = {};
    (by[p.id] || []).forEach(function (e) { (days[e.date] = days[e.date] || { IN: 0, OUT: 0, RET: 0, DMG: 0, OTH: 0 })[e.type] += e.qty; });
    Object.keys(days).sort().forEach(function (d) {
      var x = days[d], close = prev + x.IN + x.RET - x.OUT - x.DMG - x.OTH;
      rows.push({ 'Date': d, 'Product ID': p.id, 'Product Name': p.name, 'Previous Stock': prev, 'Stock In': x.IN, 'Outbound': x.OUT,
        'Return': x.RET, 'Damage': x.DMG, 'Others': x.OTH, 'Closing Stock': close });
      prev = close;
    });
  });
  rows.sort(function (a, b) { return a.Date !== b.Date ? (a.Date < b.Date ? -1 : 1) : (a['Product Name'] < b['Product Name'] ? -1 : 1); });
  return rows;
}
// Writes ledger (with recalculated balances), cached current stock and the Stock Summary sheet.
function commit_(events, products, bal) {
  var pm = {};
  products.forEach(function (p) { pm[p.id] = p; if (bal[p.id] !== undefined) p.current = bal[p.id]; });
  var ps = sh_('products'), head = headRow_(ps), ci = col_(head, 'Current Stock') + 1, lr = ps.getLastRow();
  if (lr > 1) {
    var vals = ps.getRange(2, ci, lr - 1, 1).getValues();
    products.forEach(function (p) { vals[p.row - 2][0] = p.current; });
    ps.getRange(2, ci, lr - 1, 1).setValues(vals);
  }
  replaceRows_(sh_('ledger'), events.map(function (e) {
    var o = { 'Transaction Date': e.date, 'Transaction ID': e.id, 'Product ID': e.pid, 'Product Name': pm[e.pid] ? pm[e.pid].name : '',
      'Transaction Type': TYPE[e.type].label, 'Balance After Transaction': e.balance, 'Note': e.note, 'Submitted At': e.submitted,
      'Backdated': e.submitted.substring(0, 10) > e.date ? 'Yes' : 'No', 'Updated At': e.updated || '' };
    o[TYPE[e.type].col] = e.qty; return o;
  }));
  replaceRows_(sh_('summary'), summaryRows_(events, products));
}
function renameInSheets_(pid, name) {
  ['IN', 'OUT', 'RET', 'DMG', 'OTH'].forEach(function (k) {
    var s = sh_(k), t = table_(s); if (!t.values.length) return;
    var ci = col_(t.head, 'Product ID'), ni = col_(t.head, 'Product Name'), ch = false;
    t.values.forEach(function (r) { if (String(r[ci]) === pid) { r[ni] = name; ch = true; } });
    if (ch) s.getRange(2, 1, t.values.length, t.head.length).setValues(t.values);
  });
}

// ======================== VALIDATION ========================
function validTxn_(type, p, products, keepPid) {
  var date = vdate_(p.date);
  if (date > today_()) fail_('Future dates are not allowed. Choose today or an earlier date.');
  var prod = products.filter(function (x) { return x.id === p.pid; })[0];
  if (!prod) fail_('Please select a valid product.');
  if (prod.status !== 'Active' && prod.id !== keepPid) fail_('"' + prod.name + '" is inactive. Reactivate it to record new transactions.');
  var qty = Number(p.qty);
  if (!isFinite(qty) || qty <= 0 || Math.floor(qty) !== qty || qty > 1000000000) fail_('Quantity must be a positive whole number.');
  var note = String(p.note || '').trim().substring(0, 500), reason = String(p.reason || '').trim().substring(0, 200);
  if (type === 'OTH' && !reason) fail_('Reason is required for Others.');
  return { date: date, pid: prod.id, pname: prod.name, qty: qty, note: note, reason: reason };
}
function ledgerNote_(type, v) { return type === 'OTH' ? v.reason + (v.note ? ' — ' + v.note : '') : v.note; }
function pub_(p, b) {
  return { id: p.id, name: p.name, category: p.category, opening: p.opening, current: b === undefined ? p.current : b, status: p.status, created: p.created, updated: p.updated };
}
function evOut_(e, c) {
  return { id: e.id, date: e.date, pid: e.pid, pname: c.pm[e.pid] ? c.pm[e.pid].name : e.pid, type: e.type, typeLabel: TYPE[e.type].label,
    sign: TYPE[e.type].sign, qty: e.qty, balance: e.balance, note: e.note, submitted: e.submitted, updated: e.updated,
    backdated: e.submitted.substring(0, 10) > e.date };
}
function summaryFor_(c, date) {
  var rows = [], tot = { prev: 0, in: 0, out: 0, ret: 0, dmg: 0, oth: 0, close: 0 };
  c.products.forEach(function (p) {
    var prev = p.opening, d = { IN: 0, OUT: 0, RET: 0, DMG: 0, OTH: 0 }, act = false;
    (c.by[p.id] || []).forEach(function (e) {
      if (e.date < date) prev += TYPE[e.type].sign * e.qty;
      else if (e.date === date) { d[e.type] += e.qty; act = true; }
    });
    var close = prev + d.IN + d.RET - d.OUT - d.DMG - d.OTH;
    if (p.status !== 'Active' && !act && prev === 0) return;
    rows.push({ id: p.id, name: p.name, category: p.category, status: p.status, prev: prev, in: d.IN, out: d.OUT, ret: d.RET, dmg: d.DMG, oth: d.OTH, close: close });
    tot.prev += prev; tot.in += d.IN; tot.out += d.OUT; tot.ret += d.RET; tot.dmg += d.DMG; tot.oth += d.OTH; tot.close += close;
  });
  rows.sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
  return { rows: rows, totals: tot };
}

// ======================== READ API (called from the web app) ========================
function getInitData() {
  return api_(function () {
    var c = ctx_();
    return { today: today_(), lowStock: CONFIG.LOW_STOCK_THRESHOLD, products: c.products.map(function (p) { return pub_(p, c.bal[p.id]); }) };
  });
}
function getDashboard(date) {
  return api_(function () {
    var c = ctx_(); date = vdate_(date || today_());
    var sm = summaryFor_(c, date), active = c.products.filter(function (p) { return p.status === 'Active'; });
    var totalStock = active.reduce(function (a, p) { return a + (c.bal[p.id] || 0); }, 0);
    var recent = c.events.slice().sort(function (a, b) { return a.submitted < b.submitted ? 1 : (a.submitted > b.submitted ? -1 : 0); })
      .slice(0, 10).map(function (e) { return evOut_(e, c); });
    var low = active.map(function (p) { return { id: p.id, name: p.name, current: c.bal[p.id] || 0 }; })
      .filter(function (x) { return x.current <= CONFIG.LOW_STOCK_THRESHOLD; }).sort(function (a, b) { return a.current - b.current; }).slice(0, 10);
    var days = {};
    c.events.forEach(function (e) { var d = days[e.date] = days[e.date] || { in: 0, out: 0 }; if (TYPE[e.type].sign > 0) d.in += e.qty; else d.out += e.qty; });
    var mov = [];
    for (var i = 6; i >= 0; i--) { var d = addDays_(date, -i); mov.push({ date: d, in: (days[d] || {}).in || 0, out: (days[d] || {}).out || 0 }); }
    return { date: date, totalProducts: active.length, totalStock: totalStock, totals: sm.totals, recent: recent, low: low, movement: mov, lowLimit: CONFIG.LOW_STOCK_THRESHOLD };
  });
}
function getCurrentStock() {
  return api_(function () {
    var c = ctx_();
    return c.products.filter(function (p) { return p.status === 'Active'; }).map(function (p) {
      var t = sumTypes_(c.by[p.id] || []);
      return { id: p.id, name: p.name, category: p.category, opening: p.opening, in: t.IN, out: t.OUT, ret: t.RET, dmg: t.DMG, oth: t.OTH, current: c.bal[p.id] || 0 };
    });
  });
}
function getProductDetails(id) {
  return api_(function () {
    var c = ctx_(), p = c.pm[id]; if (!p) fail_('Product not found.');
    var list = c.by[id] || [], t = sumTypes_(list);
    return { product: pub_(p, c.bal[id]), totals: { in: t.IN, out: t.OUT, ret: t.RET, dmg: t.DMG, oth: t.OTH },
      history: list.slice().reverse().slice(0, 100).map(function (e) { return evOut_(e, c); }) };
  });
}
// Historical availability: stock at end of the chosen date, limited by every later balance (so a backdated outbound can never push a later day negative).
function getAvailableStock(pid, date, excludeId) {
  return api_(function () {
    var c = ctx_(), p = c.pm[pid]; if (!p) fail_('Product not found.');
    date = vdate_(date || today_());
    var bal = p.opening, endBal = p.opening, minAfter = Infinity;
    (c.by[pid] || []).forEach(function (e) {
      if (e.id === excludeId) return;
      bal += TYPE[e.type].sign * e.qty;
      if (e.date <= date) endBal = bal; else minAfter = Math.min(minAfter, bal);
    });
    return { current: bal, endOfDate: endBal, available: Math.max(0, Math.min(endBal, minAfter)) };
  });
}
function getStockSummary(date) {
  return api_(function () { var c = ctx_(); date = vdate_(date || today_()); var s = summaryFor_(c, date); return { date: date, rows: s.rows, totals: s.totals }; });
}
function getTransactions(f) {
  return api_(function () {
    f = f || {}; var c = ctx_(), q = String(f.q || '').trim().toLowerCase();
    var list = c.events.filter(function (e) {
      return (!f.from || e.date >= f.from) && (!f.to || e.date <= f.to) && (!f.pid || e.pid === f.pid) && (!f.type || e.type === f.type) && (!q || e.id.toLowerCase().indexOf(q) >= 0);
    });
    var dir = f.sort === 'asc' ? 1 : -1;
    list.sort(function (a, b) {
      if (a.date !== b.date) return a.date < b.date ? -dir : dir;
      if (a.submitted !== b.submitted) return a.submitted < b.submitted ? -dir : dir;
      return 0;
    });
    var size = Math.min(Math.max(Number(f.pageSize) || 20, 1), 100), pages = Math.max(1, Math.ceil(list.length / size));
    var page = Math.min(Math.max(Number(f.page) || 1, 1), pages);
    return { total: list.length, page: page, pages: pages, rows: list.slice((page - 1) * size, page * size).map(function (e) { return evOut_(e, c); }) };
  });
}
function getTransaction(id) {
  return api_(function () {
    var c = ctx_(), e = c.events.filter(function (x) { return x.id === id; })[0]; if (!e) fail_('Transaction not found.');
    var t = table_(sh_(e.type)), ii = col_(t.head, 'Transaction ID'), row = t.values.filter(function (r) { return String(r[ii]) === id; })[0];
    var o = row ? obj_(t.head, row) : {}, out = evOut_(e, c);
    out.note = String(o['Note'] || ''); out.reason = String(o['Reason'] || ''); return out;
  });
}

// ======================== WRITE API ========================
function saveProduct(p) {
  return api_(function () {
    p = p || {};
    return mutate_(['products', 'ledger', 'summary', 'IN', 'OUT', 'RET', 'DMG', 'OTH'], function () {
      var name = String(p.name || '').trim();
      if (!name) fail_('Product name is required.');
      if (name.length > 100) fail_('Product name is too long (max 100 characters).');
      var category = String(p.category || '').trim().substring(0, 60);
      var opening = (p.opening === '' || p.opening == null) ? 0 : Number(p.opening);
      if (!isFinite(opening) || opening < 0 || Math.floor(opening) !== opening) fail_('Opening stock must be a whole number, 0 or more.');
      var products = loadProducts_(), lname = name.toLowerCase(), ts = now_();
      if (products.some(function (x) { return x.id !== p.id && x.name.toLowerCase() === lname; }))
        fail_('A product named "' + name + '" already exists. Duplicate names are not allowed.');
      var events = loadEvents_(), pr, renamed = false;
      if (!p.id) {
        pr = { row: sh_('products').getLastRow() + 1, id: nextId_('PRD_SEQ', 'P-', 4, maxNum_(products.map(function (x) { return x.id; }))),
          name: name, category: category, opening: opening, current: opening, status: 'Active', created: ts, updated: ts };
        products.push(pr);
      } else {
        pr = products.filter(function (x) { return x.id === p.id; })[0]; if (!pr) fail_('Product not found.');
        renamed = pr.name !== name; pr.name = name; pr.category = category; pr.opening = opening; pr.updated = ts;
        if (p.status === 'Active' || p.status === 'Inactive') pr.status = p.status;
      }
      var r = replay_(events, products, [pr.id]);
      if (r.error) fail_('This opening stock would make history negative. ' + r.error);
      writeProductRow_(pr);
      if (renamed) renameInSheets_(pr.id, name);
      commit_(events, products, r.bal);
      return { id: pr.id };
    });
  });
}
function setProductStatus(id, status) {
  return api_(function () {
    if (status !== 'Active' && status !== 'Inactive') fail_('Invalid status.');
    return mutate_(['products'], function () {
      var pr = loadProducts_().filter(function (x) { return x.id === id; })[0]; if (!pr) fail_('Product not found.');
      pr.status = status; pr.updated = now_(); writeProductRow_(pr); return { id: id, status: status };
    });
  });
}
function addTransaction(p) {
  return api_(function () {
    if (!p || !TYPE[p.type]) fail_('Invalid transaction type.');
    var rid = String(p.requestId || '').substring(0, 100);
    if (!rid) fail_('Missing request ID. Reload the page and try again.');
    return mutate_(['products', 'ledger', 'summary', p.type], function () {
      var cache = CacheService.getScriptCache();
      if (cache.get('req_' + rid)) fail_('This entry was already submitted. Duplicate submission blocked.');
      var products = loadProducts_(), events = loadEvents_(), v = validTxn_(p.type, p, products, null);
      var id = nextId_('TXN_SEQ', p.type + '-', 6, maxNum_(events.map(function (e) { return e.id; })));
      var ev = { date: v.date, id: id, pid: v.pid, type: p.type, qty: v.qty, note: ledgerNote_(p.type, v), submitted: now_(), updated: '' };
      events.push(ev);
      var r = replay_(events, products, [v.pid]);
      if (r.error) fail_(r.error);
      var s = sh_(p.type), head = headRow_(s);
      var o = { 'Transaction Date': v.date, 'Transaction ID': id, 'Product ID': v.pid, 'Product Name': v.pname, 'Quantity': v.qty,
        'Note': v.note, 'Reason': v.reason, 'Submitted At': ev.submitted, 'Updated At': '' };
      s.getRange(s.getLastRow() + 1, 1, 1, head.length).setValues([head.map(function (h) { return o[h] === undefined ? '' : o[h]; })]);
      commit_(events, products, r.bal);
      cache.put('req_' + rid, id, 21600);
      return { id: id, date: v.date, backdated: ev.submitted.substring(0, 10) > v.date, balance: ev.balance };
    });
  });
}
function updateTransaction(p) {
  return api_(function () {
    p = p || {};
    var pre = loadEvents_().filter(function (e) { return e.id === p.id; })[0];
    if (!pre) fail_('Transaction not found.');
    return mutate_(['products', 'ledger', 'summary', pre.type], function () {
      var products = loadProducts_(), events = loadEvents_();
      var ev = events.filter(function (e) { return e.id === p.id; })[0]; if (!ev) fail_('Transaction not found.');
      var v = validTxn_(ev.type, p, products, ev.pid), oldPid = ev.pid;
      ev.date = v.date; ev.pid = v.pid; ev.qty = v.qty; ev.note = ledgerNote_(ev.type, v); ev.updated = now_();
      var r = replay_(events, products, [oldPid, v.pid]);
      if (r.error) fail_(r.error);
      var s = sh_(ev.type), t = table_(s), ii = col_(t.head, 'Transaction ID');
      var ri = -1; t.values.forEach(function (row, i) { if (String(row[ii]) === p.id) ri = i; });
      if (ri < 0) fail_('Transaction row not found in ' + SHEET[ev.type] + '.');
      var upd = { 'Transaction Date': v.date, 'Product ID': v.pid, 'Product Name': v.pname, 'Quantity': v.qty, 'Note': v.note, 'Reason': v.reason, 'Updated At': ev.updated };
      t.head.forEach(function (h, i) { if (upd[h] !== undefined) t.values[ri][i] = upd[h]; });   // Submitted At is preserved
      s.getRange(ri + 2, 1, 1, t.head.length).setValues([t.values[ri]]);
      commit_(events, products, r.bal);
      return { id: p.id };
    });
  });
}
function deleteTransaction(id) {
  return api_(function () {
    var pre = loadEvents_().filter(function (e) { return e.id === id; })[0];
    if (!pre) fail_('Transaction not found.');
    return mutate_(['products', 'ledger', 'summary', pre.type], function () {
      var products = loadProducts_(), events = loadEvents_();
      var ev = events.filter(function (e) { return e.id === id; })[0]; if (!ev) fail_('Transaction not found.');
      events = events.filter(function (e) { return e.id !== id; });
      var r = replay_(events, products, [ev.pid]);
      if (r.error) fail_('Cannot delete: later transactions depend on this stock. ' + r.error);
      var s = sh_(ev.type), t = table_(s), ii = col_(t.head, 'Transaction ID');
      for (var i = 0; i < t.values.length; i++) if (String(t.values[i][ii]) === id) { s.deleteRow(i + 2); break; }
      commit_(events, products, r.bal);
      return { id: id };
    });
  });
}

// ======================== RECONCILIATION ========================
function reconcile_() {
  return mutate_(['products', 'ledger', 'summary'], function () {
    var c = ctx_(), fixed = [];
    c.products.forEach(function (p) { var a = c.bal[p.id] || 0; if (p.current !== a) fixed.push({ id: p.id, name: p.name, stored: p.current, actual: a }); });
    commit_(c.events, c.products, c.bal);
    return { checked: c.products.length, fixed: fixed, warning: c.error };
  });
}
function reconcileStock() { return api_(reconcile_); }
function reconcileStockMenu() {
  var r = reconcile_(), msg = 'Checked ' + r.checked + ' products. Corrected: ' + r.fixed.length + '.' + (r.warning ? '\n\nWarning: ' + r.warning : '');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

// ======================== SETUP (safe to run many times) ========================
function setupStockManagement() {
  var ss = ss_(); ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  ['dashboard', 'products', 'IN', 'OUT', 'RET', 'DMG', 'OTH', 'ledger', 'summary'].forEach(function (k, i) {
    var s = ss.getSheetByName(SHEET[k]) || ss.insertSheet(SHEET[k], i);
    if (k !== 'dashboard') { ensureHeaders_(s, HEADERS[k]); formatSheet_(s); }
  });
  buildDashboard_(ss.getSheetByName(SHEET.dashboard));
  ['Sheet1', 'Sheet 1'].forEach(function (n) {
    var d = ss.getSheetByName(n);
    if (d && d.getLastRow() === 0 && d.getLastColumn() === 0 && ss.getSheets().length > 1) ss.deleteSheet(d);
  });
  var pr = PropertiesService.getScriptProperties();
  if (!pr.getProperty('TXN_SEQ')) pr.setProperty('TXN_SEQ', String(maxNum_(loadEvents_().map(function (e) { return e.id; }))));
  if (!pr.getProperty('PRD_SEQ')) pr.setProperty('PRD_SEQ', String(maxNum_(loadProducts_().map(function (p) { return p.id; }))));
  ss.setActiveSheet(ss.getSheetByName(SHEET.dashboard));
  var msg = 'Setup complete. Existing data was kept. You can now deploy the web app.';
  console.log(msg); try { ss.toast(msg, 'Stock System', 8); } catch (e) {}
  return msg;
}
function ensureHeaders_(s, headers) {
  if (s.getLastColumn() === 0 || s.getLastRow() === 0) { s.getRange(1, 1, 1, headers.length).setValues([headers]); return; }
  var cur = headRow_(s), missing = headers.filter(function (h) { return cur.indexOf(h) < 0; });
  if (missing.length) s.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]);
}
function formatSheet_(s) {
  var head = headRow_(s), lc = head.length;
  if (s.getMaxRows() < 3000) s.insertRowsAfter(s.getMaxRows(), 3000 - s.getMaxRows());
  if (s.getMaxColumns() > lc) { try { s.deleteColumns(lc + 1, s.getMaxColumns() - lc); } catch (e) {} }
  var maxR = s.getMaxRows(), all = s.getRange(1, 1, maxR, lc);
  s.setFrozenRows(1); s.setHiddenGridlines(true); s.setTabColor('#C5A15B');
  all.setFontFamily('Inter').setFontSize(10).setFontColor('#F4F4F4').setBackground('#17191D').setVerticalAlignment('middle').setHorizontalAlignment('left');
  all.setBorder(true, true, true, true, true, true, '#2C3038', SpreadsheetApp.BorderStyle.SOLID);
  s.getBandings().forEach(function (b) { b.remove(); });
  var b = s.getRange(2, 1, maxR - 1, lc).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  b.setFirstRowColor('#17191D').setSecondRowColor('#202329');
  s.getRange(1, 1, 1, lc).setBackground('#101114').setFontColor('#C5A15B').setFontWeight('bold')
    .setBorder(false, false, true, false, false, false, '#C5A15B', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  head.forEach(function (h, i) {
    if (DATE_HEADS.indexOf(h) >= 0) s.getRange(2, i + 1, maxR - 1, 1).setNumberFormat('@');
    if (NUM_HEADS.indexOf(h) >= 0) s.getRange(1, i + 1, maxR, 1).setHorizontalAlignment('right');
    if (NUM_HEADS.indexOf(h) >= 0) s.getRange(2, i + 1, maxR - 1, 1).setNumberFormat('#,##0');
  });
  if (s.getFilter()) s.getFilter().remove();
  s.getRange(1, 1, maxR, lc).createFilter();
  s.autoResizeColumns(1, lc);
  for (var i = 1; i <= lc; i++) { var w = s.getColumnWidth(i); s.setColumnWidth(i, Math.min(Math.max(w + 16, 110), 320)); }
}
function buildDashboard_(s) {
  var lh = headRow_(sh_('ledger')), ph = headRow_(sh_('products'));
  var L = function (h) { return colLetter_(lh.indexOf(h) + 1); }, P = function (h) { return colLetter_(ph.indexOf(h) + 1); };
  var q = "'" + SHEET.ledger + "'!", pq = "'" + SHEET.products + "'!";
  var rng = function (h) { return q + '$' + L(h) + '$2:$' + L(h); };
  var day = function (h) { return '=SUMPRODUCT((' + rng('Transaction Date') + '=TEXT($B$4,"yyyy-mm-dd"))*(' + rng(h) + '))'; };
  var keep = s.getRange('B4').getValue();
  s.clear(); s.setHiddenGridlines(true); s.setTabColor('#C5A15B');
  s.getRange('A1').setValue('STOCK MANAGEMENT DASHBOARD');
  s.getRange('A2').setValue('Live overview. Change the date in B4 to see that day\'s movement. Full web app: deploy as Web App.');
  s.getRange('A4').setValue('Select Date');
  if (keep) s.getRange('B4').setValue(keep); else s.getRange('B4').setFormula('=TODAY()');
  s.getRange('B4').setNumberFormat('yyyy-mm-dd').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  var rows = [
    ['Total Products (Active)', '=COUNTIF(' + pq + '$' + P('Status') + '$2:$' + P('Status') + ',"Active")'],
    ['Total Current Stock (Active)', '=SUMIF(' + pq + '$' + P('Status') + '$2:$' + P('Status') + ',"Active",' + pq + '$' + P('Current Stock') + '$2:$' + P('Current Stock') + ')'],
    ['Stock In (selected date)', day('Stock In (+)')], ['Outbound (selected date)', day('Outbound (-)')],
    ['Return (selected date)', day('Return (+)')], ['Damage (selected date)', day('Damage (-)')], ['Others (selected date)', day('Others (-)')]
  ];
  rows.forEach(function (r, i) { s.getRange(6 + i, 1).setValue(r[0]); s.getRange(6 + i, 2).setFormula(r[1]); });
  s.getRange('A1:D14').setBackground('#17191D').setFontFamily('Inter').setFontColor('#F4F4F4');
  s.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#C5A15B');
  s.getRange('A2').setFontColor('#9298A3');
  s.getRange('A4').setFontColor('#9298A3'); s.getRange('B4').setFontColor('#C5A15B').setFontWeight('bold').setHorizontalAlignment('left');
  s.getRange('A6:A12').setFontColor('#9298A3'); s.getRange('B6:B12').setFontWeight('bold').setNumberFormat('#,##0').setHorizontalAlignment('left');
  s.getRange('A6:B12').setBorder(null, null, true, null, null, true, '#2C3038', SpreadsheetApp.BorderStyle.SOLID);
  s.setColumnWidth(1, 260); s.setColumnWidth(2, 160);
}
