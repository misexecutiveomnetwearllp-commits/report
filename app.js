/* ============================================================
   StockLedger — ERP Sales / Purchase / Stock Analyzer
   Pure client-side. No data ever leaves the browser.
   ============================================================ */
(function () {
"use strict";

/* ---------------------------------------------------------------
   1. CONSTANTS & CANONICAL SCHEMA
   --------------------------------------------------------------- */
const CANONICAL_FIELDS = [
  'Date', 'Transaction Type', 'Item Code', 'Article No', 'Brand', 'Colour',
  'Style', 'Section', 'Sub Section', 'Supplier', 'Size', 'Item Type',
  'Quantity', 'Opening Qty', 'Closing Qty', 'Price', 'Amount', 'HSN Code',
  'City', 'Discount', 'Purchase Bill Date'
];

const FIELD_KIND = {
  'Date': 'date', 'Purchase Bill Date': 'date',
  'Quantity': 'number', 'Opening Qty': 'number', 'Closing Qty': 'number',
  'Price': 'number', 'Amount': 'number'
};

const SYNONYMS = {
  'transactiontype': 'Transaction Type', 'type': 'Transaction Type', 'txntype': 'Transaction Type',
  'transactiondate': 'Date', 'date': 'Date', 'invoicedate': 'Date', 'billdate': 'Date', 'txndate': 'Date',
  'articleno': 'Article No', 'article': 'Article No', 'articlenumber': 'Article No', 'articlecode': 'Article No',
  'itemcode': 'Item Code', 'sku': 'Item Code', 'code': 'Item Code', 'skucode': 'Item Code',
  'sizename': 'Size', 'size': 'Size',
  'brandname': 'Brand', 'brand': 'Brand',
  'colourname': 'Colour', 'colorname': 'Colour', 'colour': 'Colour', 'color': 'Colour',
  'stylename': 'Style', 'style': 'Style', 'styleno': 'Style', 'stylenumber': 'Style', 'stylecode': 'Style',
  'sectionname': 'Section', 'section': 'Section', 'category': 'Section', 'department': 'Section',
  'subsectionname': 'Sub Section', 'subsection': 'Sub Section', 'subcategory': 'Sub Section',
  'suppliername': 'Supplier', 'supplier': 'Supplier', 'vendor': 'Supplier', 'vendorname': 'Supplier',
  'party': 'Supplier', 'partyname': 'Supplier',
  'transactionquantity': 'Quantity', 'quantity': 'Quantity', 'qty': 'Quantity',
  'stockqty': 'Quantity', 'balanceqty': 'Quantity',
  // OBS = opening balance, CBS = closing balance. Dono alag rakhte hain taaki
  // ek hi report mein dono aa sakein aur movement nikala ja sake.
  'cbsqty': 'Closing Qty', 'closingqty': 'Closing Qty', 'closingstock': 'Closing Qty',
  'closingbalance': 'Closing Qty', 'closingbalanceqty': 'Closing Qty', 'cbs': 'Closing Qty',
  'obsqty': 'Opening Qty', 'openingqty': 'Opening Qty', 'openingstock': 'Opening Qty',
  'openingbalance': 'Opening Qty', 'openingbalanceqty': 'Opening Qty', 'obs': 'Opening Qty',
  'itemtype': 'Item Type', 'itype': 'Item Type',
  'discname': 'Discount', 'discount': 'Discount', 'discountname': 'Discount', 'disc': 'Discount',
  'purchasebilldate': 'Purchase Bill Date',
  'transactionhsncode': 'HSN Code', 'hsncode': 'HSN Code', 'hsn': 'HSN Code',
  'transactionlocationcity': 'City', 'city': 'City', 'location': 'City', 'locationcity': 'City',
  'price': 'Price', 'rate': 'Price', 'mrp': 'Price', 'sellingprice': 'Price', 'unitprice': 'Price', 'costprice': 'Price',
  'amount': 'Amount', 'totalamount': 'Amount', 'value': 'Amount', 'saleamount': 'Amount', 'netamount': 'Amount', 'totalvalue': 'Amount'
};

const BLANK_TOKENS = new Set(['', 'na', 'n/a', 'n.a', 'n.a.', 'null', 'nil', '-', 'undefined']);
const MONTHS = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
const MONTH_LABELS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const AGG_LABELS = { sum: 'Sum', count: 'Count', avg: 'Average', min: 'Min', max: 'Max', distinct: 'Distinct count' };
const CHART_COLORS = ['#A6402C', '#1F6F5C', '#B9862F', '#4A6FA5', '#7A4CA0', '#C6784B', '#3E8E7E', '#8C5B3F', '#5B7553', '#9C4F6B'];

/* ---------------------------------------------------------------
   1b. LIBRARY LOADER — app.js khud libraries load karta hai
   ---------------------------------------------------------------
   Isse ye faayda hai ki agar index.html purana ho (jisme galat CDN
   URL tha) tab bhi site chal jayegi. Ek CDN block ho (ad-blocker,
   office firewall) to agla try hota hai.
   --------------------------------------------------------------- */
const XLSX_URLS = [
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
];
const CHART_URLS = [
  'https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js',
  'https://unpkg.com/chart.js@4.4.4/dist/chart.umd.min.js'
];

const Libs = { xlsx: null, chart: null };

function loadScriptWithFallback(urls) {
  return new Promise(function (resolve, reject) {
    let i = 0;
    (function next() {
      if (i >= urls.length) { reject(new Error('sabhi CDN fail ho gaye')); return; }
      const s = document.createElement('script');
      s.src = urls[i++];
      s.async = false;
      s.onload = function () { resolve(); };
      s.onerror = function () { if (s.parentNode) s.parentNode.removeChild(s); next(); };
      document.head.appendChild(s);
    })();
  });
}

function ensureXLSX() {
  if (typeof XLSX !== 'undefined') return Promise.resolve(true);
  if (!Libs.xlsx) {
    Libs.xlsx = loadScriptWithFallback(XLSX_URLS)
      .then(function () { return typeof XLSX !== 'undefined'; })
      .catch(function () { return false; });
  }
  return Libs.xlsx;
}

function ensureChart() {
  if (typeof Chart !== 'undefined') return Promise.resolve(true);
  if (!Libs.chart) {
    Libs.chart = loadScriptWithFallback(CHART_URLS)
      .then(function () { return typeof Chart !== 'undefined'; })
      .catch(function () { return false; });
  }
  return Libs.chart;
}

function showLibError(msg) {
  let b = document.getElementById('lib-error');
  if (!b) {
    // Purane index.html mein ye element nahi hai — bana dete hain.
    b = document.createElement('div');
    b.id = 'lib-error';
    b.className = 'lib-error';
    b.style.cssText = 'background:#F7DED7;border:1px solid #A6402C;color:#7A2716;' +
      'padding:12px 16px;border-radius:8px;font-size:13px;margin-bottom:18px;line-height:1.5;';
    const main = document.querySelector('.main');
    if (main) main.insertBefore(b, main.firstChild);
    else document.body.insertBefore(b, document.body.firstChild);
  }
  b.style.display = '';
  b.textContent = msg;
}

/* ---------------------------------------------------------------
   2. UTILITIES
   --------------------------------------------------------------- */
function normKey(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

function suggestField(header) {
  const key = normKey(header);
  if (!key) return null;
  if (SYNONYMS[key]) return SYNONYMS[key];
  for (const k in SYNONYMS) { if (key.includes(k) || k.includes(key)) return SYNONYMS[k]; }
  return null;
}

function cleanValue(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  if (BLANK_TOKENS.has(s.toLowerCase())) return null;
  return s;
}

function parseDateLoose(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + v * 86400000);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})-([A-Za-z]{3,})-(\d{2,4})$/);
  if (m) {
    const mon = MONTHS[m[2].toLowerCase().slice(0, 3)];
    if (mon !== undefined) {
      let yr = parseInt(m[3], 10); if (yr < 100) yr += 2000;
      return new Date(Date.UTC(yr, mon, parseInt(m[1], 10)));
    }
  }
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let yr = parseInt(m[3], 10); if (yr < 100) yr += 2000;
    return new Date(Date.UTC(yr, parseInt(m[2], 10) - 1, parseInt(m[1], 10)));
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d) {
  if (!d) return '';
  if (!(d instanceof Date)) { d = parseDateLoose(d); if (!d) return ''; }
  if (isNaN(d.getTime())) return '';
  return String(d.getUTCDate()).padStart(2, '0') + '-' + MONTH_LABELS[d.getUTCMonth()] + '-' + d.getUTCFullYear();
}

function fmtNum(n, decimals) {
  if (n === null || n === undefined || isNaN(n)) return '';
  const dp = decimals === undefined ? (Number.isInteger(n) ? 0 : 2) : decimals;
  return Number(n).toLocaleString('en-IN', { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

/**
 * ERP exports apne aakhir mein "Total", uska duplicate, aur "Printed on ..."
 * jaisi rows jodte hain. Inhe data samajh liya jaye to har figure double ho
 * jata hai, isliye ingest ke waqt hata dete hain.
 */
function isJunkRow(row) {
  let textCount = 0, numCount = 0, hasPrinted = false, hasTotalLabel = false;
  for (const c of row) {
    if (c === null || c === undefined || c === '') continue;
    if (typeof c === 'number') { numCount++; continue; }
    const s = String(c).trim();
    if (!s) continue;
    textCount++;
    if (/^printed\s+on/i.test(s)) hasPrinted = true;
    if (/^(grand\s+|sub\s*)?total$/i.test(s)) hasTotalLabel = true;
  }
  if (hasPrinted) return true;                          // footer line
  if (hasTotalLabel && textCount <= 2) return true;     // "Total | 16144"
  if (textCount === 0 && numCount > 0 && numCount <= 2) return true; // bare totals line
  return false;
}

/** ERP header rows mein "Reporting Period From : 01-08-2026 to 25-08-2026"
 *  jaisi line hoti hai — usse report ka asli date range nikalte hain. */
function extractReportPeriod(rows, headerIdx) {
  const scanTo = Math.min(headerIdx >= 0 ? headerIdx : 12, rows.length);
  const re = /reporting\s+period\s*(?:from)?\s*:?\s*([0-9]{1,2}[-\/][0-9]{1,2}[-\/][0-9]{2,4}|[0-9]{1,2}-[A-Za-z]{3,}-[0-9]{2,4})\s*(?:to|-|through|till)\s*([0-9]{1,2}[-\/][0-9]{1,2}[-\/][0-9]{2,4}|[0-9]{1,2}-[A-Za-z]{3,}-[0-9]{2,4})/i;
  for (let i = 0; i < scanTo; i++) {
    const row = rows[i] || [];
    for (const cell of row) {
      if (cell === null || cell === undefined) continue;
      const m = String(cell).match(re);
      if (m) {
        const from = parseDateLoose(m[1]), to = parseDateLoose(m[2]);
        if (from && to) return { from, to, raw: String(cell).trim() };
      }
    }
  }
  return null;
}

function detectHeaderRow(rows, maxScan) {
  maxScan = Math.min(maxScan || 25, rows.length);
  let best = { idx: 0, score: -1 };
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i] || [];
    let score = 0, nonEmpty = 0;
    for (const cell of row) {
      if (cell === null || cell === undefined || cell === '') continue;
      nonEmpty++;
      if (typeof cell === 'string' && suggestField(cell)) score += 2;
      else if (typeof cell === 'string' && cell.length < 30) score += 0.3;
    }
    if (nonEmpty >= 3 && score > best.score) best = { idx: i, score };
  }
  return best.idx;
}

function guessDatasetType(filename, columns, sampleRows) {
  const fn = (filename || '').toLowerCase();
  if (/stock/.test(fn)) return 'stock';
  if (/purchase/.test(fn)) return 'purchase';
  if (/sale/.test(fn)) return 'sales';
  const headers = columns.map(c => normKey(c.header));
  if (headers.some(h => h.includes('cbsqty') || h.includes('obsqty') || h.includes('closingstock') ||
                       h.includes('closingbalance') || h.includes('balanceqty'))) return 'stock';
  const typeCol = columns.find(c => c.suggested === 'Transaction Type');
  if (typeCol && sampleRows.length) {
    const vals = sampleRows.map(r => String(r[typeCol.colIdx] || '').toLowerCase()).join(' ');
    if (vals.includes('purchase')) return 'purchase';
    if (vals.includes('sale')) return 'sales';
  }
  return 'other';
}

function uid() { return Math.random().toString(36).slice(2, 10); }

/* localStorage kabhi-kabhi in-app browsers / private mode mein block hota hai,
   isliye har call try-catch mein wrapped hai. Fail ho to app normal chalta rahe. */
const Store = {
  get(k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { window.localStorage.setItem(k, v); return true; } catch (e) { return false; } },
  remove(k) { try { window.localStorage.removeItem(k); } catch (e) {} }
};

function debounce(fn, ms) {
  let t;
  return function (...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
}

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime || 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

function toCSV(headers, rows) {
  const esc = v => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const lines = [headers.map(esc).join(',')];
  for (const r of rows) lines.push(r.map(esc).join(','));
  return lines.join('\r\n');
}

let toastTimer = null;
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 2600);
}

/* ---------------------------------------------------------------
   3. APP STATE
   --------------------------------------------------------------- */
const App = {
  datasets: [],     // { id, name, type, fields[], records[], rowCount }
  relationships: [],// { id, fromDsId, fromField, toDsId, toField, enabled }
  nextDsColor: 0
};

/* Lookup cache — relationships ke through dusre dataset se value uthane ke liye.
   Data ya relationship badalne par saaf ho jata hai. */
let LookupCache = new Map();
function clearLookups() { LookupCache = new Map(); }

function getLookup(dsId, keyField, valueField) {
  const ck = dsId + '|' + keyField + '|' + valueField;
  let lut = LookupCache.get(ck);
  if (lut) return lut;
  lut = new Map();
  const ds = App.datasets.find(d => d.id === dsId);
  if (ds) {
    for (const rec of ds.records) {
      const k = rec[keyField];
      if (k === null || k === undefined || k === '') continue;
      const v = rec[valueField];
      if (v === null || v === undefined || v === '') continue;
      const ks = String(k);
      if (!lut.has(ks)) lut.set(ks, v);
    }
  }
  LookupCache.set(ck, lut);
  return lut;
}

/**
 * Kisi record se field ki value nikalta hai. Agar us record ke dataset mein
 * wo field nahi hai, to active relationships ke through jude hue dataset se
 * uthane ki koshish karta hai — jaise Power BI mein related table se column
 * aa jata hai.
 */
function resolveField(rec, field) {
  const v = rec[field];
  if (v !== undefined && v !== null && v !== '') return v;
  const dsId = rec.__ds;
  if (!dsId || !App.relationships.length) return null;

  for (const rel of App.relationships) {
    if (!rel.enabled) continue;
    // Sirf asli key ke through value uthana safe hai (many-to-one lookup).
    // Warna ek Section ke hazaron items mein se koi bhi random value aa sakti hai.
    if (!rel.score || !rel.score.canEnrich) continue;
    let myField, otherId, otherField;
    if (rel.fromDsId === dsId) { myField = rel.fromField; otherId = rel.toDsId; otherField = rel.toField; }
    else if (rel.toDsId === dsId) { myField = rel.toField; otherId = rel.fromDsId; otherField = rel.fromField; }
    else continue;

    const other = App.datasets.find(d => d.id === otherId);
    if (!other || other.fields.indexOf(field) === -1) continue;

    const key = rec[myField];
    if (key === null || key === undefined || key === '') continue;
    const got = getLookup(otherId, otherField, field).get(String(key));
    if (got !== undefined) return got;
  }
  return null;
}

function fieldsOfDataset(ds) {
  return ds.fields.slice();
}

function allLoadedFields() {
  const set = new Set();
  App.datasets.forEach(ds => ds.fields.forEach(f => set.add(f)));
  return CANONICAL_FIELDS.filter(f => set.has(f)).concat([...set].filter(f => !CANONICAL_FIELDS.includes(f)));
}

// "combined" pseudo dataset: union of records from chosen datasets, tagged with Source Type / Source File
function getRecordsForSelection(selection) {
  // selection: dataset id, or '__all__', or '__type:sales__' etc.
  if (selection === '__all__') {
    return App.datasets.flatMap(ds => ds.records);
  }
  if (selection && selection.startsWith('__type:')) {
    const t = selection.slice(7, -2);
    return App.datasets.filter(d => d.type === t).flatMap(ds => ds.records);
  }
  const ds = App.datasets.find(d => d.id === selection);
  return ds ? ds.records : [];
}

function datasetsOfType(type) { return App.datasets.filter(d => d.type === type); }

/* ---------------------------------------------------------------
   3b. AUTO-PERSISTENCE (IndexedDB) — data survives closing the tab.
   Removed only when the user clicks Remove on a dataset.
   --------------------------------------------------------------- */
const IDB_NAME = 'StockLedgerDB', IDB_VERSION = 1, IDB_STORE = 'datasets';
let _idbDb = null;

function idbAvailable() { return typeof indexedDB !== 'undefined'; }

function idbOpen() {
  if (_idbDb) return Promise.resolve(_idbDb);
  return new Promise((resolve, reject) => {
    if (!idbAvailable()) { reject(new Error('IndexedDB not available')); return; }
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = () => { req.result.createObjectStore(IDB_STORE, { keyPath: 'id' }); };
    req.onsuccess = () => { _idbDb = req.result; resolve(_idbDb); };
    req.onerror = () => reject(req.error);
  });
}

function serializeDate(d) { return d instanceof Date ? { __date: d.toISOString() } : d; }
function deserializeDate(v) { return (v && typeof v === 'object' && v.__date) ? new Date(v.__date) : v; }

function serializeDatasetForIdb(ds) {
  return {
    id: ds.id, name: ds.name, type: ds.type, fields: ds.fields,
    origin: ds.origin || null, mapping: ds.mapping || null, headerIdx: ds.headerIdx || 0,
    reportPeriod: ds.reportPeriod ? { from: serializeDate(ds.reportPeriod.from), to: serializeDate(ds.reportPeriod.to), raw: ds.reportPeriod.raw } : null,
    records: ds.records.map(r => {
      const o = {};
      ds.fields.forEach(f => { o[f] = serializeDate(r[f]); });
      return o;
    })
  };
}

function hydrateDatasetFromIdb(raw) {
  const records = raw.records.map(r => {
    const o = { __ds: raw.id };
    raw.fields.forEach(f => { o[f] = deserializeDate(r[f]); });
    return o;
  });
  return {
    id: raw.id, name: raw.name, type: raw.type, fields: raw.fields, records,
    rowCount: records.length, colorIdx: App.nextDsColor++,
    origin: raw.origin || null, mapping: raw.mapping || null, headerIdx: raw.headerIdx || 0,
    reportPeriod: raw.reportPeriod ? { from: deserializeDate(raw.reportPeriod.from), to: deserializeDate(raw.reportPeriod.to), raw: raw.reportPeriod.raw } : null
  };
}

/** Fire-and-forget save — never blocks the UI, and a failure here (private
 *  browsing, storage disabled) never breaks the app; it just means this
 *  dataset won't survive a browser close. */
function idbSaveDataset(ds) {
  if (!idbAvailable()) return;
  idbOpen().then(db => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).put(serializeDatasetForIdb(ds));
  }).catch(err => console.error('StockLedger: could not auto-save dataset', err));
}

function idbDeleteDataset(id) {
  if (!idbAvailable()) return;
  idbOpen().then(db => {
    const tx = db.transaction(IDB_STORE, 'readwrite');
    tx.objectStore(IDB_STORE).delete(id);
  }).catch(err => console.error('StockLedger: could not remove auto-saved dataset', err));
}

function idbLoadAllDatasets() {
  if (!idbAvailable()) return Promise.resolve([]);
  return idbOpen().then(db => new Promise((resolve, reject) => {
    const tx = db.transaction(IDB_STORE, 'readonly');
    const req = tx.objectStore(IDB_STORE).getAll();
    req.onsuccess = () => resolve((req.result || []).map(hydrateDatasetFromIdb));
    req.onerror = () => reject(req.error);
  })).catch(err => { console.error('StockLedger: could not restore auto-saved data', err); return []; });
}

/** App start hote hi pehle se load ki hui files wapas la deta hai. */
function restorePersistedDatasets() {
  idbLoadAllDatasets().then(saved => {
    if (!saved.length) return;
    App.datasets = App.datasets.concat(saved);
    refreshAfterDataChange();
    toast(saved.length + ' previously loaded file(s) restored.');
  });
}

/* ---------------------------------------------------------------
   4. FILE IMPORT
   --------------------------------------------------------------- */
function initImport() {
  const dz = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  document.getElementById('browse-btn').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => handleFiles(e.target.files));

  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('dragover'); }));
  dz.addEventListener('drop', e => handleFiles(e.dataTransfer.files));
  dz.addEventListener('click', e => { if (e.target === dz || e.target.closest('.dropzone-inner')) fileInput.click(); });
}

function handleFiles(fileList) {
  Array.from(fileList).forEach(file => readWorkbook(file));
}

function readWorkbook(file) {
  if (file.size === 0) { toast(file.name + ': file is empty (0 bytes) — please download it again.'); return; }
  if (file.size > 80 * 1024 * 1024) { toast(file.name + ': this file is very large (' + (file.size / 1024 / 1024).toFixed(0) + ' MB) — this may be slow to load in the browser.'); }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    // Library abhi tak load na hui ho to pehle usko laate hain, phir parse.
    ensureXLSX().then(function (ok) {
      if (!ok) {
        showLibError('The Excel reader library could not be loaded. Please check your internet connection, ad-blocker ' +
          'or office firewall — file reading will not work without it.');
        return;
      }
      parseWorkbookBuffer(data, file);
    });
  };
  reader.onerror = () => toast(file.name + ': the browser could not read this file (' + (reader.error ? reader.error.name : 'unknown') + ').');
  reader.readAsArrayBuffer(file);
}

/**
 * XLSX.read kई formats support karta hai, lekin ek hi option-set har file ke
 * saath kaam nahi karta — kuch ERP ".xls" asal mein HTML table hote hain, kuch
 * mein encoding alag hoti hai. Isliye 3 tareeke try karte hain, pehla jo chale
 * wahi use hota hai. Sab fail ho to asli error dikhate hain, generic nahi.
 */
function parseWorkbookBuffer(data, file) {
  if (typeof XLSX === 'undefined') {
    showLibError('The Excel reader library could not be loaded. Please refresh the page; if it still fails, check your ad-blocker or firewall.');
    return;
  }

  const attempts = [
    () => XLSX.read(data, { type: 'array' }),
    () => XLSX.read(data, { type: 'array', codepage: 65001 }),
    () => XLSX.read(data, { type: 'array', raw: true, cellText: false }),
    () => {
      // kuch exports asal mein HTML table hote hain jinka naam .xls hota hai
      let str;
      try { str = new TextDecoder('utf-8').decode(data); }
      catch (e) { str = new TextDecoder('windows-1252').decode(data); }
      return XLSX.read(str, { type: 'string' });
    }
  ];

  let lastErr = null;
  for (const attempt of attempts) {
    try {
      const wb = attempt();
      if (wb && wb.SheetNames && wb.SheetNames.length) {
        if (wb.SheetNames.length === 1) loadSheet(file, wb, wb.SheetNames[0]);
        else renderMultiSheetChoice(file, wb);
        return;
      }
      lastErr = new Error('The file was read but contains no sheets.');
    } catch (err) {
      lastErr = err;
    }
  }

  console.error('StockLedger: could not parse', file.name, lastErr);
  toast(file.name + ': could not be read — ' + (lastErr ? lastErr.message : 'unknown error') +
    '. Try opening it in Excel and using "Save As -> .xlsx".');
}

function renderMultiSheetChoice(file, wb) {
  const queue = document.getElementById('import-queue');
  const card = document.createElement('div');
  card.className = 'import-card';
  card.innerHTML =
    '<div class="import-card-head"><span class="fname">' + escapeHtml(file.name) + '</span></div>' +
    '<div class="import-card-row">This file has multiple sheets — pick one to import: ' +
    '<select class="select sheet-pick">' + wb.SheetNames.map(n => '<option value="' + escapeHtml(n) + '">' + escapeHtml(n) + '</option>').join('') + '</select>' +
    '<button class="ghost-btn small sheet-pick-go">Use this sheet</button></div>';
  queue.appendChild(card);
  card.querySelector('.sheet-pick-go').addEventListener('click', () => {
    const sel = card.querySelector('.sheet-pick').value;
    card.remove();
    loadSheet(file, wb, sel);
  });
}

function loadSheet(file, wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  ingestRows(rows, file.name, sheetName, null);
}

/**
 * Shared entry point for BOTH sources (uploaded file and Google Sheet).
 * Takes a raw 2D array, finds the header row, and opens the mapping card.
 * `origin` is null for files, or {url, key, sheet} for Google Sheets.
 */
function ingestRows(rows, sourceName, sheetName, origin) {
  const headerIdx = detectHeaderRow(rows);
  const headerRow = rows[headerIdx] || [];
  const columns = [];
  headerRow.forEach((h, colIdx) => {
    if (h === null || h === undefined || String(h).trim() === '') return;
    columns.push({ colIdx, header: String(h).trim(), suggested: suggestField(h) });
  });
  const allDataRows = rows.slice(headerIdx + 1).filter(r => r && r.some(c => c !== null && c !== undefined && c !== ''));
  const dataRows = allDataRows.filter(r => !isJunkRow(r));
  const droppedCount = allDataRows.length - dataRows.length;

  if (!columns.length || !dataRows.length) {
    toast(sourceName + ': could not find a header row with data — check the source.');
    return;
  }

  const guessedType = guessDatasetType(sourceName + ' ' + (sheetName || ''), columns, dataRows.slice(0, 30));
  const reportPeriod = extractReportPeriod(rows, headerIdx);
  renderImportCard(sourceName, sheetName, columns, dataRows, guessedType, origin, headerIdx, droppedCount, reportPeriod);
}

function renderImportCard(filename, sheetName, columns, dataRows, guessedType, origin, headerIdx, droppedCount, reportPeriod) {
  const queue = document.getElementById('import-queue');
  const cardId = uid();
  const card = document.createElement('div');
  card.className = 'import-card';
  card.dataset.cardId = cardId;

  const typeOptions = ['sales', 'purchase', 'stock', 'other']
    .map(t => '<option value="' + t + '"' + (t === guessedType ? ' selected' : '') + '>' + t[0].toUpperCase() + t.slice(1) + '</option>').join('');

  const rowsHtml = columns.map((c, i) => {
    const sample = dataRows.slice(0, 3).map(r => r[c.colIdx]).filter(v => v !== null && v !== undefined && v !== '').slice(0, 1)[0];
    const options = ['<option value="">— ignore —</option>']
      .concat(CANONICAL_FIELDS.map(f => '<option value="' + f + '"' + (c.suggested === f ? ' selected' : '') + '>' + f + '</option>'))
      .concat(['<option value="__custom__' + i + '"' + (!c.suggested ? ' selected' : '') + '>Keep as "' + escapeHtml(c.header) + '"</option>']);
    return '<tr data-colidx="' + c.colIdx + '" data-header="' + escapeHtml(c.header) + '">' +
      '<td>' + escapeHtml(c.header) + '</td>' +
      '<td class="sample-cell">' + escapeHtml(sample === undefined ? '' : String(sample)) + '</td>' +
      '<td><select class="map-select">' + options.join('') + '</select></td>' +
      '</tr>';
  }).join('');

  card.innerHTML =
    '<div class="import-card-head">' +
      '<span class="fname">' + escapeHtml(filename) + (sheetName ? ' &middot; ' + escapeHtml(sheetName) : '') + '</span>' +
      '<span>' + (reportPeriod
          ? '<span class="rp-badge" title="' + escapeHtml(reportPeriod.raw) + '">\uD83D\uDCC5 ' +
            fmtDate(reportPeriod.from) + ' \u2192 ' + fmtDate(reportPeriod.to) + '</span> '
          : '') +
        dataRows.length.toLocaleString('en-IN') + ' rows detected' +
        (droppedCount ? ' <span class="drop-note" title="ERP total / footer rows hata diye gaye — warna har figure double ho jata">· ' + droppedCount + ' total/footer row' + (droppedCount > 1 ? 's' : '') + ' skipped</span>' : '') +
      '</span>' +
    '</div>' +
    '<div class="import-card-row">' +
      '<label class="toolbar-label">Name:</label><input type="text" class="text-input ds-name" value="' + escapeHtml(filename.replace(/\.[^.]+$/, '')) + '">' +
      '<label class="toolbar-label">Type:</label><select class="select ds-type">' + typeOptions + '</select>' +
      '<span class="spacer"></span>' +
      '<button class="ghost-btn primary confirm-import">Add to workspace</button>' +
      '<button class="ghost-btn small discard-import">Discard</button>' +
    '</div>' +
    '<details><summary style="cursor:pointer;font-size:12.5px;color:var(--text-soft);">Review column mapping (' + columns.length + ' columns)</summary>' +
    '<table class="map-table"><thead><tr><th>Column in file</th><th>Sample value</th><th>Maps to</th></tr></thead><tbody>' + rowsHtml + '</tbody></table>' +
    '</details>';

  queue.appendChild(card);

  card.querySelector('.discard-import').addEventListener('click', () => card.remove());
  card.querySelector('.confirm-import').addEventListener('click', () => {
    confirmImport(card, filename, columns, dataRows, origin, headerIdx, reportPeriod);
  });
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** Turns raw rows + a column mapping into normalized record objects. */
function buildRecords(dataRows, mapping, dsId) {
  const dateFields = new Set(mapping.filter(m => FIELD_KIND[m.field] === 'date').map(m => m.field));
  const numberFields = new Set(mapping.filter(m => FIELD_KIND[m.field] === 'number').map(m => m.field));
  return dataRows.map(r => {
    const rec = { __ds: dsId };
    mapping.forEach(m => {
      let v = cleanValue(r[m.colIdx]);
      if (v !== null && dateFields.has(m.field)) v = parseDateLoose(v);
      else if (v !== null && numberFields.has(m.field)) { const n = Number(v); v = isNaN(n) ? null : n; }
      // if same canonical field appears twice, keep first non-null
      if (rec[m.field] === undefined || rec[m.field] === null) rec[m.field] = v;
    });
    return rec;
  });
}

function confirmImport(card, filename, columns, dataRows, origin, headerIdx, reportPeriod) {
  const name = card.querySelector('.ds-name').value.trim() || filename;
  const type = card.querySelector('.ds-type').value;
  const mapRows = card.querySelectorAll('.map-table tbody tr');
  const mapping = []; // { colIdx, field }
  mapRows.forEach(tr => {
    const colIdx = parseInt(tr.dataset.colidx, 10);
    const header = tr.dataset.header;
    const sel = tr.querySelector('.map-select').value;
    let field = sel;
    if (sel.startsWith('__custom__')) field = header;
    if (field) mapping.push({ colIdx, field });
  });
  if (!mapping.length) { toast('Map at least one column before adding.'); return; }

  const dsId = uid();
  const records = buildRecords(dataRows, mapping, dsId);
  const fields = [...new Set(mapping.map(m => m.field))];
  const ds = {
    id: dsId, name, type, fields, records, rowCount: records.length,
    colorIdx: App.nextDsColor++,
    origin: origin || null,     // {url, key, sheet} when pulled from Google Sheets
    mapping: mapping,           // remembered so Refresh needs no re-mapping
    headerIdx: headerIdx || 0,
    reportPeriod: reportPeriod || null   // ERP header ki "Reporting Period" line
  };
  App.datasets.push(ds);
  idbSaveDataset(ds);
  card.remove();
  toast('Added "' + name + '" — ' + records.length.toLocaleString('en-IN') + ' rows.');
  refreshAfterDataChange();
}

function refreshAfterDataChange() {
  clearLookups();
  clearAnchorCache();
  // dusri file aate hi connections khud detect kar lete hain
  if (Prefs.autoDetectLinks !== false && App.datasets.length > 1 && !App.relationships.length) autoDetectRelationships(true);
  rescoreRelationships();
  renderSidebarDatasets();
  renderLoadedTable();
  populateDatasetSelects();
  renderExplore();
  renderPivotFieldList();
  computePivot();
  populateQuickSelects();
  renderQuickFieldList();
  renderQuickReport();
  renderDashboard();
  renderInsights();
  renderPerformance();
  renderCatalog();
  renderRelations();
  updateRangeNotes();
  if (Drill.open) renderDrill();
  maybeAutoShowSnapshot();
}

function removeDataset(id) {
  App.datasets = App.datasets.filter(d => d.id !== id);
  idbDeleteDataset(id);
  refreshAfterDataChange();
}

function typeTagClass(t) { return { sales: 'tag-sales', purchase: 'tag-purchase', stock: 'tag-stock' }[t] || 'tag-other'; }

function renderSidebarDatasets() {
  const wrap = document.getElementById('sidebar-dataset-list');
  if (!App.datasets.length) { wrap.innerHTML = '<div class="empty-hint">No files yet — start on Import Data.</div>'; return; }
  wrap.innerHTML = App.datasets.map(ds =>
    '<div class="sd-item">' +
      '<span class="sd-name">' + escapeHtml(ds.name) + '</span>' +
      '<span class="sd-meta">' + ds.rowCount.toLocaleString('en-IN') + ' rows</span>' +
      '<span class="sd-type-tag ' + typeTagClass(ds.type) + '">' + ds.type + '</span>' +
    '</div>'
  ).join('');
}

function dateRangeOf(ds) {
  // Stock / OBS-CBS reports mein Date column nahi hoti, unka period ERP ke
  // header ki "Reporting Period" line se aata hai.
  if (ds.reportPeriod) {
    return fmtDate(ds.reportPeriod.from) + ' \u2192 ' + fmtDate(ds.reportPeriod.to);
  }
  const dField = ds.fields.includes('Date') ? 'Date' : (ds.fields.includes('Purchase Bill Date') ? 'Purchase Bill Date' : null);
  if (!dField) return '—';
  let min = null, max = null;
  ds.records.forEach(r => {
    const d = r[dField];
    if (!d || !(d instanceof Date) || isNaN(d.getTime())) return;
    if (!min || d < min) min = d;
    if (!max || d > max) max = d;
  });
  if (!min) return '—';
  return fmtDate(min) + ' → ' + fmtDate(max);
}

function renderLoadedTable() {
  const panel = document.getElementById('loaded-datasets-panel');
  const tbody = document.querySelector('#loaded-datasets-table tbody');
  if (!App.datasets.length) { panel.style.display = 'none'; return; }
  panel.style.display = '';
  tbody.innerHTML = App.datasets.map(ds =>
    '<tr>' +
      '<td>' + escapeHtml(ds.name) + (ds.origin ? ' <span class="src-badge">Sheet</span>' : '') + '</td>' +
      '<td><span class="sd-type-tag ' + typeTagClass(ds.type) + '">' + ds.type + '</span></td>' +
      '<td>' + ds.rowCount.toLocaleString('en-IN') + '</td>' +
      '<td>' + dateRangeOf(ds) + '</td>' +
      '<td>' + ds.fields.join(', ') + '</td>' +
      '<td>' + (ds.origin ? '<button class="ghost-btn small refresh-ds" data-id="' + ds.id + '">↻ Refresh</button> ' : '') +
        '<button class="ghost-btn small remove-ds" data-id="' + ds.id + '">Remove</button></td>' +
    '</tr>'
  ).join('');
  tbody.querySelectorAll('.remove-ds').forEach(btn => btn.addEventListener('click', () => removeDataset(btn.dataset.id)));
  tbody.querySelectorAll('.refresh-ds').forEach(btn => btn.addEventListener('click', () => refreshDataset(btn.dataset.id)));
}

function populateDatasetSelects() {
  const opts = ['<option value="__all__">All files combined</option>']
    .concat(['sales', 'purchase', 'stock'].filter(t => datasetsOfType(t).length).map(t =>
      '<option value="__type:' + t + '__">All ' + t + ' files</option>'))
    .concat(App.datasets.map(ds => '<option value="' + ds.id + '">' + escapeHtml(ds.name) + '</option>'));
  ['explore-dataset-select', 'pivot-dataset-select', 'quick-dataset-select'].forEach(id => {
    const el = document.getElementById(id);
    const prev = el.value;
    el.innerHTML = opts.join('');
    if ([...el.options].some(o => o.value === prev)) el.value = prev;
  });
}

/* ---------------------------------------------------------------
   4b. GOOGLE SHEETS CONNECTOR (Apps Script web app)
   --------------------------------------------------------------- */
const GS = { url: '', key: '', meta: null };

function initSheets() {
  // source switcher
  document.querySelectorAll('.source-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.source-tab').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('source-file').style.display = btn.dataset.source === 'file' ? '' : 'none';
      document.getElementById('source-sheet').style.display = btn.dataset.source === 'sheet' ? '' : 'none';
    });
  });

  const urlEl = document.getElementById('gs-url');
  const keyEl = document.getElementById('gs-key');
  const savedUrl = Store.get('sl_gs_url'), savedKey = Store.get('sl_gs_key');
  if (savedUrl) urlEl.value = savedUrl;
  if (savedKey) keyEl.value = savedKey;

  document.getElementById('gs-connect').addEventListener('click', connectSheet);
  document.getElementById('gs-forget').addEventListener('click', () => {
    Store.remove('sl_gs_url'); Store.remove('sl_gs_key');
    urlEl.value = ''; keyEl.value = '';
    GS.url = ''; GS.key = ''; GS.meta = null;
    document.getElementById('gs-sheet-list').style.display = 'none';
    setGsStatus('Saved details cleared.', '');
    updateGsOnlyButtons();
  });

  document.getElementById('pivot-to-sheet').addEventListener('click', pivotToSheet);
  const insSheet = document.getElementById('insights-to-sheet');   // Reorder tab removed
  if (insSheet) insSheet.addEventListener('click', insightsToSheet);
}

function setGsStatus(msg, cls) {
  const el = document.getElementById('gs-status');
  el.textContent = msg;
  el.className = 'connect-status' + (cls ? ' ' + cls : '');
}

function updateGsOnlyButtons() {
  const on = !!(GS.url && GS.meta && GS.meta.canWrite);
  document.querySelectorAll('.gs-only').forEach(el => { el.style.display = on ? '' : 'none'; });
}

function gsUrlWith(params) {
  const u = GS.url + (GS.url.indexOf('?') === -1 ? '?' : '&');
  return u + new URLSearchParams(Object.assign({ key: GS.key }, params)).toString();
}

function gsGet(params) {
  return fetch(gsUrlWith(params), { method: 'GET', redirect: 'follow' })
    .then(r => r.text())
    .then(text => {
      let data;
      try { data = JSON.parse(text); }
      catch (e) {
        throw new Error('The script returned HTML instead of JSON. Is the deployment set to "Who has access: Anyone"? The URL must end in /exec.');
      }
      if (!data.ok) throw new Error(data.error || 'Unknown error from Apps Script.');
      return data;
    });
}

function gsPost(body) {
  // text/plain rakhna zaroori hai — warna browser CORS preflight bhejta hai
  // jise Apps Script handle nahi kar pata.
  return fetch(GS.url, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ key: GS.key }, body))
  })
    .then(r => r.text())
    .then(text => {
      let data;
      try { data = JSON.parse(text); } catch (e) { throw new Error('Sheet ne unexpected response diya.'); }
      if (!data.ok) throw new Error(data.error || 'Write failed.');
      return data;
    });
}

function connectSheet() {
  const url = document.getElementById('gs-url').value.trim();
  const key = document.getElementById('gs-key').value.trim();
  if (!url) { setGsStatus('Enter the web app URL.', 'err'); return; }
  if (!/^https:\/\/script\.google\.com\/.*\/exec/.test(url)) {
    setGsStatus('The URL should look like https://script.google.com/.../exec', 'err');
    return;
  }
  if (!key) { setGsStatus('Enter the API key.', 'err'); return; }

  GS.url = url; GS.key = key;
  setGsStatus('Connecting…', 'busy');

  gsGet({ action: 'meta' }).then(meta => {
    GS.meta = meta;
    if (document.getElementById('gs-remember').checked) {
      Store.set('sl_gs_url', url); Store.set('sl_gs_key', key);
    }
    setGsStatus('Connected: ' + meta.spreadsheetName + ' (' + meta.sheets.length + ' sheets)', 'ok');
    renderSheetList(meta);
    updateGsOnlyButtons();
    pullSnapshotConfigFromSheet();
  }).catch(err => {
    GS.meta = null;
    setGsStatus(err.message, 'err');
    document.getElementById('gs-sheet-list').style.display = 'none';
    updateGsOnlyButtons();
  });
}

function renderSheetList(meta) {
  const wrap = document.getElementById('gs-sheet-list');
  wrap.style.display = '';
  if (!meta.sheets.length) { wrap.innerHTML = '<div class="empty-hint">No visible sheets found in this spreadsheet.</div>'; return; }

  wrap.innerHTML = '<h3>' + escapeHtml(meta.spreadsheetName) + '</h3>' +
    meta.sheets.map(s =>
      '<div class="sheet-row" data-sheet="' + escapeHtml(s.name) + '">' +
        '<span class="sheet-name">' + escapeHtml(s.name) + '</span>' +
        '<span class="sheet-meta">' + s.rows.toLocaleString('en-IN') + ' rows × ' + s.cols + ' cols</span>' +
        '<select class="select sheet-type">' +
          ['sales', 'purchase', 'stock', 'other'].map(t =>
            '<option value="' + t + '"' + (t === s.guessedType ? ' selected' : '') + '>' + t[0].toUpperCase() + t.slice(1) + '</option>').join('') +
        '</select>' +
        '<button class="ghost-btn small pull-sheet">Pull data</button>' +
        '<div class="progress-track" style="display:none;width:100%;"><div class="progress-fill"></div></div>' +
      '</div>'
    ).join('');

  wrap.querySelectorAll('.pull-sheet').forEach(btn => {
    btn.addEventListener('click', () => {
      const row = btn.closest('.sheet-row');
      pullSheet(row.dataset.sheet, row.querySelector('.sheet-type').value, row);
    });
  });
}

/** Sheet ko chunks mein kheenchta hai taaki 50,000+ rows bhi aa jayen. */
function pullSheet(sheetName, type, rowEl) {
  const btn = rowEl.querySelector('.pull-sheet');
  const track = rowEl.querySelector('.progress-track');
  const fill = rowEl.querySelector('.progress-fill');
  btn.disabled = true; btn.textContent = 'Loading…';
  track.style.display = '';
  fill.style.width = '0%';

  const chunkSize = (GS.meta && GS.meta.maxRowsPerRequest) || 5000;
  const all = [];

  function next(offset) {
    return gsGet({ action: 'data', sheet: sheetName, offset: offset, limit: chunkSize }).then(res => {
      res.rows.forEach(r => all.push(r));
      const pct = res.totalRows ? Math.min(100, Math.round((all.length / res.totalRows) * 100)) : 100;
      fill.style.width = pct + '%';
      btn.textContent = 'Loading… ' + pct + '%';
      if (!res.done && res.rows.length) return next(offset + res.rows.length);
      return null;
    });
  }

  next(0).then(() => {
    btn.disabled = false; btn.textContent = 'Pull data';
    setTimeout(() => { track.style.display = 'none'; }, 600);
    if (!all.length) { toast(sheetName + ': no rows found.'); return; }

    // Type ko user ki choice se force karte hain, guess se nahi.
    ingestRows(all, sheetName, null, { url: GS.url, key: GS.key, sheet: sheetName });
    const card = document.querySelector('#import-queue .import-card:last-child');
    if (card) {
      const sel = card.querySelector('.ds-type');
      if (sel) sel.value = type;
      card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    // Import queue Upload tab ke neeche hai, isliye user ko wahan le jaate hain
    toast(sheetName + ' loaded — neeche mapping confirm karke "Add to workspace" dabao.');
  }).catch(err => {
    btn.disabled = false; btn.textContent = 'Pull data';
    track.style.display = 'none';
    toast('Pull failed: ' + err.message);
  });
}

/** Pehle se load ki hui sheet ko dobara kheench kar refresh karta hai. */
function refreshDataset(id) {
  const ds = App.datasets.find(d => d.id === id);
  if (!ds || !ds.origin) return;
  toast('Refreshing "' + ds.name + '"…');

  const savedUrl = GS.url, savedKey = GS.key;
  GS.url = ds.origin.url; GS.key = ds.origin.key;
  const chunkSize = (GS.meta && GS.meta.maxRowsPerRequest) || 5000;
  const all = [];

  function next(offset) {
    return gsGet({ action: 'data', sheet: ds.origin.sheet, offset: offset, limit: chunkSize }).then(res => {
      res.rows.forEach(r => all.push(r));
      if (!res.done && res.rows.length) return next(offset + res.rows.length);
      return null;
    });
  }

  next(0).then(() => {
    const dataRows = all.slice(ds.headerIdx + 1).filter(r => r && r.some(c => c !== null && c !== undefined && c !== ''));
    ds.records = buildRecords(dataRows, ds.mapping, ds.id);
    ds.rowCount = ds.records.length;
    idbSaveDataset(ds);
    GS.url = savedUrl || GS.url; GS.key = savedKey || GS.key;
    refreshAfterDataChange();
    toast('"' + ds.name + '" refreshed — ' + ds.rowCount.toLocaleString('en-IN') + ' rows.');
  }).catch(err => {
    GS.url = savedUrl; GS.key = savedKey;
    toast('Refresh failed: ' + err.message);
  });
}

function pivotToSheet() {
  const grid = pivotToGrid();
  if (!grid) { toast('Build a pivot first.'); return; }
  const sheetName = prompt('Which sheet should this be written to? (it will be created if missing)', 'StockLedger Pivot');
  if (!sheetName) return;
  toast('Writing to sheet...');
  gsPost({ action: 'write', sheet: sheetName, values: [grid.headers].concat(grid.rows), mode: 'replace' })
    .then(res => toast('Done - "' + res.sheet + '" now has ' + res.rowsWritten + ' rows.'))
    .catch(err => toast('Write failed: ' + err.message));
}

function insightsToSheet() {
  const grid = insightsToGrid();
  if (!grid) { toast('There is nothing to export yet.'); return; }
  const sheetName = prompt('Which sheet should this be written to? (it will be created if missing)', 'StockLedger Reorder');
  if (!sheetName) return;
  toast('Writing to sheet...');
  gsPost({ action: 'write', sheet: sheetName, values: [grid.headers].concat(grid.rows), mode: 'replace' })
    .then(res => toast('Done - "' + res.sheet + '" now has ' + res.rowsWritten + ' rows.'))
    .catch(err => toast('Write failed: ' + err.message));
}

/* ---------------------------------------------------------------
   5. TAB NAVIGATION
   --------------------------------------------------------------- */
function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });
}

/* ---------------------------------------------------------------
   6. EXPLORE TAB
   --------------------------------------------------------------- */
const ExploreState = { page: 1, pageSize: 100, sortField: null, sortDir: 1, search: '', filters: {} };

function initExplore() {
  document.getElementById('explore-dataset-select').addEventListener('change', () => {
    ExploreState.page = 1; ExploreState.filters = {}; ExploreState.sortField = null; renderExplore();
  });
  document.getElementById('explore-search').addEventListener('input', debounce(e => { ExploreState.search = e.target.value; ExploreState.page = 1; renderExplore(); }, 200));
  document.getElementById('explore-export').addEventListener('click', exportExploreCSV);
  document.getElementById('explore-add-filter').addEventListener('click', openExploreFilterPicker);
}

/** Column chunkar uski values mein se filter lagane ka popup. */
function openExploreFilterPicker() {
  const fields = exploreFields();
  if (!fields.length) { toast('Load a file first.'); return; }
  const recs = currentExploreRecords();

  const popup = document.createElement('div');
  popup.className = 'modal-backdrop';
  popup.innerHTML = '<div class="modal-box">' +
    '<h3>Filter lagao</h3>' +
    '<label class="toolbar-label">Column:</label> ' +
    '<select id="efp-field" class="select">' + fields.map(f => '<option value="' + escapeHtml(f) + '">' + escapeHtml(f) + '</option>').join('') + '</select>' +
    '<div id="efp-values" class="efp-values"></div>' +
    '<div class="modal-actions">' +
      '<button class="ghost-btn small" id="efp-all">Select all</button>' +
      '<button class="ghost-btn small" id="efp-none">Clear</button>' +
      '<span class="spacer"></span>' +
      '<button class="ghost-btn primary small" id="efp-apply">Apply</button>' +
      '<button class="ghost-btn small" id="efp-cancel">Cancel</button>' +
    '</div></div>';
  document.body.appendChild(popup);

  function loadValues() {
    const f = popup.querySelector('#efp-field').value;
    const counts = new Map();
    recs.forEach(r => {
      let v = r[f];
      if (v instanceof Date) v = fmtDate(v);
      v = (v === null || v === undefined || v === '') ? '(blank)' : String(v);
      counts.set(v, (counts.get(v) || 0) + 1);
    });
    const sel = ExploreState.filters[f];
    const vals = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 500);
    popup.querySelector('#efp-values').innerHTML = vals.map(([v, c]) =>
      '<label class="efp-row"><input type="checkbox" value="' + escapeHtml(v) + '"' +
      (!sel || sel.has(v) ? ' checked' : '') + '> <span>' + escapeHtml(v) + '</span>' +
      '<span class="efp-count">' + c.toLocaleString('en-IN') + '</span></label>').join('') +
      (counts.size > 500 ? '<div class="empty-hint">' + counts.size.toLocaleString('en-IN') + ' unique values — showing the top 500.</div>' : '');
  }
  loadValues();
  popup.querySelector('#efp-field').addEventListener('change', loadValues);
  popup.querySelector('#efp-all').onclick = () => popup.querySelectorAll('#efp-values input').forEach(c => c.checked = true);
  popup.querySelector('#efp-none').onclick = () => popup.querySelectorAll('#efp-values input').forEach(c => c.checked = false);
  popup.querySelector('#efp-cancel').onclick = () => popup.remove();
  popup.addEventListener('click', e => { if (e.target === popup) popup.remove(); });
  popup.querySelector('#efp-apply').onclick = () => {
    const f = popup.querySelector('#efp-field').value;
    const boxes = [...popup.querySelectorAll('#efp-values input')];
    const checked = boxes.filter(b => b.checked).map(b => b.value);
    if (checked.length === boxes.length) delete ExploreState.filters[f];
    else ExploreState.filters[f] = new Set(checked);
    popup.remove();
    ExploreState.page = 1;
    renderExplore();
  };
}

function renderExploreFilterChips() {
  const wrap = document.getElementById('explore-filters');
  const keys = Object.keys(ExploreState.filters);
  wrap.innerHTML = keys.map(f =>
    '<span class="filter-chip">' + escapeHtml(f) + ': ' + ExploreState.filters[f].size + ' selected' +
    '<button data-f="' + escapeHtml(f) + '">&times;</button></span>').join('');
  wrap.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    delete ExploreState.filters[b.dataset.f];
    renderExplore();
  }));
}

function applyExploreFilters(recs) {
  const keys = Object.keys(ExploreState.filters);
  if (!keys.length) return recs;
  return recs.filter(r => keys.every(f => {
    let v = r[f];
    if (v instanceof Date) v = fmtDate(v);
    v = (v === null || v === undefined || v === '') ? '(blank)' : String(v);
    return ExploreState.filters[f].has(v);
  }));
}

function currentExploreRecords() {
  const sel = document.getElementById('explore-dataset-select').value;
  let recs = getRecordsForSelection(sel);
  if (sel === '__all__' || sel.startsWith('__type:')) {
    // tag with source so it's visible when combined
  }
  return recs;
}

function exploreFields() {
  const sel = document.getElementById('explore-dataset-select').value;
  if (sel === '__all__') return allLoadedFields();
  if (sel.startsWith('__type:')) {
    const t = sel.slice(7, -2);
    const set = new Set(); datasetsOfType(t).forEach(d => d.fields.forEach(f => set.add(f)));
    return CANONICAL_FIELDS.filter(f => set.has(f));
  }
  const ds = App.datasets.find(d => d.id === sel);
  return ds ? ds.fields : [];
}

function renderExplore() {
  const table = document.getElementById('explore-table');
  if (!App.datasets.length) { table.innerHTML = '<tr><td class="empty-hint">Load a file first.</td></tr>'; document.getElementById('explore-pager').innerHTML = ''; return; }

  let recs = applyExploreFilters(currentExploreRecords());
  const fields = exploreFields();
  renderExploreFilterChips();

  if (ExploreState.search) {
    const q = ExploreState.search.toLowerCase();
    recs = recs.filter(r => fields.some(f => {
      const v = r[f];
      if (v === null || v === undefined) return false;
      const s = v instanceof Date ? fmtDate(v) : String(v);
      return s.toLowerCase().includes(q);
    }));
  }

  if (ExploreState.sortField) {
    const f = ExploreState.sortField, dir = ExploreState.sortDir;
    recs = recs.slice().sort((a, b) => {
      let va = a[f], vb = b[f];
      if (va === null || va === undefined) va = '';
      if (vb === null || vb === undefined) vb = '';
      if (va instanceof Date && vb instanceof Date) return (va - vb) * dir;
      if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
      return String(va).localeCompare(String(vb)) * dir;
    });
  }

  const total = recs.length;
  const totalPages = Math.max(1, Math.ceil(total / ExploreState.pageSize));
  ExploreState.page = Math.min(ExploreState.page, totalPages);
  const start = (ExploreState.page - 1) * ExploreState.pageSize;
  const pageRecs = recs.slice(start, start + ExploreState.pageSize);

  const thead = '<thead><tr>' + fields.map(f =>
    '<th data-field="' + f + '">' + f + (ExploreState.sortField === f ? '<span class="sort-arrow">' + (ExploreState.sortDir === 1 ? '▲' : '▼') + '</span>' : '') + '</th>'
  ).join('') + '</tr></thead>';

  const isNum = f => FIELD_KIND[f] === 'number';
  const tbody = '<tbody>' + pageRecs.map(r =>
    '<tr>' + fields.map(f => {
      let v = r[f];
      if (v instanceof Date) v = fmtDate(v);
      else if (typeof v === 'number') v = isNum(f) ? fmtNum(v) : String(v);
      return '<td class="' + (isNum(f) ? 'num' : '') + '">' + escapeHtml(v === null || v === undefined ? '' : v) + '</td>';
    }).join('') + '</tr>'
  ).join('') + '</tbody>';

  table.innerHTML = thead + tbody + exploreFootHtml(recs, fields, isNum);
  makeTableResizable(table);
  table.querySelectorAll('thead th').forEach(th => th.addEventListener('click', () => {
    const f = th.dataset.field;
    if (ExploreState.sortField === f) ExploreState.sortDir *= -1; else { ExploreState.sortField = f; ExploreState.sortDir = 1; }
    renderExplore();
  }));

  const pager = document.getElementById('explore-pager');
  pager.innerHTML =
    '<button id="ep-prev"' + (ExploreState.page <= 1 ? ' disabled' : '') + '>← Prev</button>' +
    '<span>Page ' + ExploreState.page + ' of ' + totalPages + ' &middot; ' + total.toLocaleString('en-IN') + ' rows</span>' +
    '<button id="ep-next"' + (ExploreState.page >= totalPages ? ' disabled' : '') + '>Next →</button>';
  const prevBtn = document.getElementById('ep-prev'), nextBtn = document.getElementById('ep-next');
  if (prevBtn) prevBtn.addEventListener('click', () => { ExploreState.page--; renderExplore(); });
  if (nextBtn) nextBtn.addEventListener('click', () => { ExploreState.page++; renderExplore(); });
}

/** Filtered rows par numeric columns ka total \u2014 paginated view mein bhi
 *  poore filtered set ka sum dikhata hai, na ki sirf current page ka. */
function exploreFootHtml(recs, fields, isNum) {
  if (!fields.some(isNum)) return '';
  let out = '<tfoot><tr>';
  fields.forEach((f, i) => {
    if (i === 0) { out += '<td>Total (' + recs.length.toLocaleString('en-IN') + ' rows)</td>'; return; }
    if (isNum(f)) {
      const sum = recs.reduce((s, r) => s + (typeof r[f] === 'number' ? r[f] : 0), 0);
      out += '<td class="num">' + fmtNum(sum) + '</td>';
    } else out += '<td></td>';
  });
  return out + '</tr></tfoot>';
}

function exportExploreCSV() {
  const fields = exploreFields();
  let recs = applyExploreFilters(currentExploreRecords());
  if (ExploreState.search) {
    const q = ExploreState.search.toLowerCase();
    recs = recs.filter(r => fields.some(f => {
      const v = r[f]; if (v === null || v === undefined) return false;
      const s = v instanceof Date ? fmtDate(v) : String(v);
      return s.toLowerCase().includes(q);
    }));
  }
  const rows = recs.map(r => fields.map(f => { const v = r[f]; return v instanceof Date ? fmtDate(v) : v; }));
  downloadBlob(toCSV(fields, rows), 'stockledger-export.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   7. PIVOT BUILDER
   --------------------------------------------------------------- */
const Pivot = { filters: [], rows: [], columns: [], values: [] };
// filters: {field, excluded:Set}
// rows/columns: {field, grain} grain used only for date fields: 'day'|'month'|'quarter'|'year'
// values: {field, agg}

let pivotChart = null;
let pivotLastResult = null;

function initPivot() {
  document.getElementById('pivot-dataset-select').addEventListener('change', () => { resetPivotZones(); renderPivotFieldList(); computePivot(); });
  document.getElementById('pivot-show-as').addEventListener('change', computePivot);
  document.getElementById('pivot-export').addEventListener('click', exportPivotCSV);
  document.getElementById('pivot-add-chart').addEventListener('click', chartCurrentPivot);
  document.getElementById('pivot-chart-close').addEventListener('click', () => {
    document.getElementById('pivot-chart-holder').style.display = 'none';
    if (pivotChart) { pivotChart.destroy(); pivotChart = null; }
  });

  ['filters', 'rows', 'columns', 'values'].forEach(zone => {
    const body = document.querySelector('.zone-body[data-drop="' + zone + '"]');
    body.addEventListener('dragover', e => { e.preventDefault(); body.classList.add('dragover'); });
    body.addEventListener('dragleave', () => body.classList.remove('dragover'));
    body.addEventListener('drop', e => {
      e.preventDefault(); body.classList.remove('dragover');
      const field = e.dataTransfer.getData('text/field');
      if (field) addFieldToZone(zone, field);
    });
  });
}

function resetPivotZones() { Pivot.filters = []; Pivot.rows = []; Pivot.columns = []; Pivot.values = []; renderPivotZones(); }

function currentPivotFields() {
  const sel = document.getElementById('pivot-dataset-select').value;
  if (!sel) return [];
  if (sel === '__all__') return allLoadedFields();
  if (sel.startsWith('__type:')) {
    const t = sel.slice(7, -2); const set = new Set();
    datasetsOfType(t).forEach(d => d.fields.forEach(f => set.add(f)));
    return CANONICAL_FIELDS.filter(f => set.has(f));
  }
  const ds = App.datasets.find(d => d.id === sel);
  return ds ? ds.fields : [];
}

function renderPivotFieldList() {
  const wrap = document.getElementById('pivot-field-list');
  const fields = currentPivotFields();
  if (!fields.length) { wrap.innerHTML = '<div class="empty-hint">Load data first.</div>'; return; }
  wrap.innerHTML = fields.map(f => {
    const kind = FIELD_KIND[f] || 'text';
    return '<div class="field-chip kind-' + kind + '" draggable="true" data-field="' + f + '"><span class="chip-dot"></span>' + f + '</div>';
  }).join('');
  wrap.querySelectorAll('.field-chip').forEach(chip => {
    chip.addEventListener('dragstart', e => {
      e.dataTransfer.setData('text/field', chip.dataset.field);
      chip.classList.add('dragging');
    });
    chip.addEventListener('dragend', () => chip.classList.remove('dragging'));
  });
}

function addFieldToZone(zone, field) {
  if (zone === 'values') {
    Pivot.values.push({ field, agg: FIELD_KIND[field] === 'number' ? 'sum' : 'count' });
  } else if (zone === 'filters') {
    if (Pivot.filters.some(f => f.field === field)) return;
    Pivot.filters.push({ field, excluded: new Set() });
  } else {
    const arr = Pivot[zone];
    if (arr.some(f => f.field === field)) return;
    arr.push({ field, grain: FIELD_KIND[field] === 'date' ? 'month' : undefined });
  }
  renderPivotZones();
  computePivot();
}

function removeFromZone(zone, idx) { Pivot[zone].splice(idx, 1); renderPivotZones(); computePivot(); }

function renderPivotZones() {
  renderZone('rows', item => zoneChipHtml(item, 'rows'));
  renderZone('columns', item => zoneChipHtml(item, 'columns'));
  renderZone('values', item => valueChipHtml(item));
  renderZone('filters', item => filterChipHtml(item));
  wireZoneEvents();
}

function renderZone(zone, htmlFn) {
  const body = document.querySelector('.zone-body[data-drop="' + zone + '"]');
  body.innerHTML = Pivot[zone].map(htmlFn).join('');
}

function zoneChipHtml(item, zone) {
  const idx = Pivot[zone].indexOf(item);
  let grainSel = '';
  if (FIELD_KIND[item.field] === 'date') {
    grainSel = '<select class="grain-select" data-zone="' + zone + '" data-idx="' + idx + '">' +
      ['day', 'month', 'quarter', 'year'].map(g => '<option value="' + g + '"' + (item.grain === g ? ' selected' : '') + '>' + g[0].toUpperCase() + g.slice(1) + '</option>').join('') +
      '</select>';
  }
  return '<div class="zone-chip"><span class="zc-name">' + item.field + '</span>' + grainSel +
    '<button class="chip-remove-btn" data-zone="' + zone + '" data-idx="' + idx + '">✕</button></div>';
}

function valueChipHtml(item) {
  const idx = Pivot.values.indexOf(item);
  const aggOpts = Object.keys(AGG_LABELS).filter(a => FIELD_KIND[item.field] === 'number' || a === 'count' || a === 'distinct');
  return '<div class="zone-chip"><span class="zc-name">' + item.field + '</span>' +
    '<select class="agg-select" data-idx="' + idx + '">' + aggOpts.map(a => '<option value="' + a + '"' + (item.agg === a ? ' selected' : '') + '>' + AGG_LABELS[a] + '</option>').join('') + '</select>' +
    '<button class="chip-remove-btn" data-zone="values" data-idx="' + idx + '">✕</button></div>';
}

function filterChipHtml(item) {
  const idx = Pivot.filters.indexOf(item);
  const activeCount = item.excluded.size;
  return '<div class="zone-chip"><span class="zc-name">' + item.field + (activeCount ? ' (' + activeCount + ' excluded)' : '') + '</span>' +
    '<button class="ghost-btn small filter-edit-btn" data-idx="' + idx + '" style="padding:2px 8px;">Edit</button>' +
    '<button class="chip-remove-btn" data-zone="filters" data-idx="' + idx + '">✕</button></div>';
}

function wireZoneEvents() {
  document.querySelectorAll('.chip-remove-btn[data-zone]').forEach(btn => {
    btn.onclick = () => removeFromZone(btn.dataset.zone, parseInt(btn.dataset.idx, 10));
  });
  document.querySelectorAll('.grain-select').forEach(sel => {
    sel.onchange = () => { Pivot[sel.dataset.zone][parseInt(sel.dataset.idx, 10)].grain = sel.value; computePivot(); };
  });
  document.querySelectorAll('.agg-select').forEach(sel => {
    sel.onchange = () => { Pivot.values[parseInt(sel.dataset.idx, 10)].agg = sel.value; computePivot(); };
  });
  document.querySelectorAll('.filter-edit-btn').forEach(btn => {
    btn.onclick = () => openFilterEditor(parseInt(btn.dataset.idx, 10));
  });
}

function openFilterEditor(idx) {
  const item = Pivot.filters[idx];
  const recs = getRecordsForSelection(document.getElementById('pivot-dataset-select').value);
  const values = [...new Set(recs.map(r => valueKey(r[item.field], item.field)))].sort();
  const listHtml = values.map(v =>
    '<label style="display:flex;gap:6px;align-items:center;padding:3px 0;font-size:12.5px;">' +
    '<input type="checkbox" value="' + escapeHtml(v) + '"' + (item.excluded.has(v) ? '' : ' checked') + '> ' + escapeHtml(v || '(blank)') + '</label>'
  ).join('');
  const popup = document.createElement('div');
  popup.style.cssText = 'position:fixed;inset:0;background:rgba(22,33,44,0.45);z-index:1000;display:flex;align-items:center;justify-content:center;';
  popup.innerHTML = '<div style="background:#FFFDF8;border-radius:10px;padding:18px 20px;max-width:360px;max-height:70vh;overflow:auto;box-shadow:0 10px 40px rgba(0,0,0,0.3);">' +
    '<h3 style="margin-bottom:10px;">Filter: ' + item.field + '</h3>' +
    '<div style="margin-bottom:8px;"><button class="ghost-btn small" id="filt-all">Select all</button> <button class="ghost-btn small" id="filt-none">Clear all</button></div>' +
    '<div id="filt-list">' + listHtml + '</div>' +
    '<div style="margin-top:14px;text-align:right;"><button class="ghost-btn primary small" id="filt-apply">Apply</button> <button class="ghost-btn small" id="filt-cancel">Cancel</button></div>' +
    '</div>';
  document.body.appendChild(popup);
  popup.querySelector('#filt-all').onclick = () => popup.querySelectorAll('#filt-list input').forEach(cb => cb.checked = true);
  popup.querySelector('#filt-none').onclick = () => popup.querySelectorAll('#filt-list input').forEach(cb => cb.checked = false);
  popup.querySelector('#filt-cancel').onclick = () => popup.remove();
  popup.querySelector('#filt-apply').onclick = () => {
    const excluded = new Set();
    popup.querySelectorAll('#filt-list input').forEach(cb => { if (!cb.checked) excluded.add(cb.value); });
    item.excluded = excluded;
    popup.remove();
    renderPivotZones();
    computePivot();
  };
}

function valueKey(v, field) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return dateKeyForGrain(v, 'day');
  return String(v);
}

function dateKeyForGrain(d, grain) {
  const y = d.getUTCFullYear(), m = d.getUTCMonth();
  if (grain === 'year') return String(y);
  if (grain === 'quarter') return y + ' Q' + (Math.floor(m / 3) + 1);
  if (grain === 'month') return MONTH_LABELS[m] + ' ' + y;
  return fmtDate(d);
}

function groupKeyFor(rec, fieldDef) {
  const v = rec[fieldDef.field];
  if (v === null || v === undefined) return '(blank)';
  if (v instanceof Date) return dateKeyForGrain(v, fieldDef.grain || 'month');
  return String(v);
}

function computePivot() {
  const sel = document.getElementById('pivot-dataset-select').value;
  const table = document.getElementById('pivot-table');
  if (!sel || !App.datasets.length) { table.innerHTML = ''; pivotLastResult = null; return; }
  let recs = getRecordsForSelection(sel);

  Pivot.filters.forEach(f => {
    if (!f.excluded.size) return;
    recs = recs.filter(r => !f.excluded.has(valueKey(r[f.field], f.field)));
  });

  if (!Pivot.rows.length && !Pivot.columns.length && !Pivot.values.length) {
    table.innerHTML = '<tr><td class="empty-hint">Drag fields into Rows / Columns / Values to build a pivot.</td></tr>';
    pivotLastResult = null;
    return;
  }
  if (!Pivot.values.length) {
    table.innerHTML = '<tr><td class="empty-hint">Add at least one field to Values.</td></tr>';
    pivotLastResult = null;
    return;
  }

  const rowKeyOf = rec => Pivot.rows.length ? Pivot.rows.map(f => groupKeyFor(rec, f)).join(' | ') : '__all__';
  const colKeyOf = rec => Pivot.columns.length ? Pivot.columns.map(f => groupKeyFor(rec, f)).join(' | ') : '__all__';

  const cellMap = new Map(); // rowKey -> colKey -> [{field, agg}] accumulator state
  const rowKeys = new Set(), colKeys = new Set();

  recs.forEach(rec => {
    const rk = rowKeyOf(rec), ck = colKeyOf(rec);
    rowKeys.add(rk); colKeys.add(ck);
    const cellId = rk + '\u0001' + ck;
    let acc = cellMap.get(cellId);
    if (!acc) { acc = Pivot.values.map(() => ({ sum: 0, count: 0, min: Infinity, max: -Infinity, distinctSet: new Set() })); cellMap.set(cellId, acc); }
    Pivot.values.forEach((vDef, vi) => {
      const raw = rec[vDef.field];
      const num = typeof raw === 'number' ? raw : (raw instanceof Date ? null : Number(raw));
      const a = acc[vi];
      if (raw !== null && raw !== undefined) {
        a.count++;
        a.distinctSet.add(raw instanceof Date ? fmtDate(raw) : String(raw));
        if (!isNaN(num) && num !== null) { a.sum += num; if (num < a.min) a.min = num; if (num > a.max) a.max = num; }
      }
    });
  });

  function resolveAgg(acc, vDef) {
    switch (vDef.agg) {
      case 'sum': return acc.sum;
      case 'count': return acc.count;
      case 'avg': return acc.count ? acc.sum / acc.count : 0;
      case 'min': return acc.min === Infinity ? 0 : acc.min;
      case 'max': return acc.max === -Infinity ? 0 : acc.max;
      case 'distinct': return acc.distinctSet.size;
      default: return acc.sum;
    }
  }

  const sortedRowKeys = [...rowKeys].sort();
  const sortedColKeys = [...colKeys].sort();
  const showAs = document.getElementById('pivot-show-as').value;

  // build numeric grid per value field
  const grid = {}; // vi -> rowKey -> colKey -> number
  Pivot.values.forEach((vDef, vi) => {
    grid[vi] = {};
    sortedRowKeys.forEach(rk => {
      grid[vi][rk] = {};
      sortedColKeys.forEach(ck => {
        const acc = cellMap.get(rk + '\u0001' + ck);
        grid[vi][rk][ck] = acc ? resolveAgg(acc[vi], vDef) : 0;
      });
    });
  });

  // totals for % calculations
  Pivot.values.forEach((vDef, vi) => {
    const rowTotals = {}, colTotals = {}; let grand = 0;
    sortedRowKeys.forEach(rk => { rowTotals[rk] = sortedColKeys.reduce((s, ck) => s + grid[vi][rk][ck], 0); });
    sortedColKeys.forEach(ck => { colTotals[ck] = sortedRowKeys.reduce((s, rk) => s + grid[vi][rk][ck], 0); });
    grand = sortedRowKeys.reduce((s, rk) => s + rowTotals[rk], 0);
    if (showAs !== 'value') {
      sortedRowKeys.forEach(rk => sortedColKeys.forEach(ck => {
        const v = grid[vi][rk][ck];
        let base = 1;
        if (showAs === 'rowpct') base = rowTotals[rk] || 1;
        else if (showAs === 'colpct') base = colTotals[ck] || 1;
        else if (showAs === 'totalpct') base = grand || 1;
        grid[vi][rk][ck] = base ? (v / base) * 100 : 0;
      }));
    }
  });

  pivotLastResult = { rowKeys: sortedRowKeys, colKeys: sortedColKeys, grid, values: Pivot.values, showAs, rowLabel: Pivot.rows.map(r => r.field).join(' / ') || '(all)' };
  renderPivotTable(pivotLastResult);
}

function renderPivotTable(res) {
  const table = document.getElementById('pivot-table');
  const isPct = res.showAs !== 'value';
  const singleValue = res.values.length === 1;

  let headRow1 = '<th rowspan="2">' + (res.rowLabel) + '</th>';
  res.colKeys.forEach(ck => {
    headRow1 += '<th colspan="' + res.values.length + '">' + (ck === '__all__' ? 'Total' : ck) + '</th>';
  });
  if (res.colKeys.length > 1 || !singleValue) headRow1 += '<th colspan="' + res.values.length + '">Grand Total</th>';

  let headRow2 = '';
  res.colKeys.forEach(() => { res.values.forEach(v => { headRow2 += '<th class="num">' + v.field + ' (' + AGG_LABELS[v.agg] + ')</th>'; }); });
  if (res.colKeys.length > 1 || !singleValue) res.values.forEach(v => { headRow2 += '<th class="num">' + v.field + ' (' + AGG_LABELS[v.agg] + ')</th>'; });

  let body = '';
  res.rowKeys.forEach(rk => {
    body += '<tr><td>' + (rk === '__all__' ? 'Total' : rk) + '</td>';
    let rowGrand = res.values.map(() => 0);
    res.colKeys.forEach(ck => {
      res.values.forEach((v, vi) => {
        const val = res.grid[vi][rk][ck] || 0;
        rowGrand[vi] += val;
        body += '<td class="num">' + (isPct ? fmtNum(val, 1) + '%' : fmtNum(val)) + '</td>';
      });
    });
    if (res.colKeys.length > 1 || !singleValue) {
      res.values.forEach((v, vi) => { body += '<td class="num">' + (isPct ? '' : fmtNum(rowGrand[vi])) + '</td>'; });
    }
    body += '</tr>';
  });

  // grand total row
  let footRow = '<td>Grand Total</td>';
  res.colKeys.forEach(ck => {
    res.values.forEach((v, vi) => {
      const colSum = res.rowKeys.reduce((s, rk) => s + (res.grid[vi][rk][ck] || 0), 0);
      footRow += '<td class="num">' + (isPct ? '' : fmtNum(colSum)) + '</td>';
    });
  });
  if (res.colKeys.length > 1 || !singleValue) {
    res.values.forEach((v, vi) => {
      const total = res.rowKeys.reduce((s, rk) => s + res.colKeys.reduce((s2, ck) => s2 + (res.grid[vi][rk][ck] || 0), 0), 0);
      footRow += '<td class="num">' + (isPct ? '' : fmtNum(total)) + '</td>';
    });
  }

  table.innerHTML = '<thead><tr>' + headRow1 + '</tr><tr>' + headRow2 + '</tr></thead><tbody>' + body + '</tbody><tfoot><tr>' + footRow + '</tr></tfoot>';
}

function pivotToGrid() {
  if (!pivotLastResult) return null;
  const res = pivotLastResult;
  const headers = [res.rowLabel];
  res.colKeys.forEach(ck => res.values.forEach(v => headers.push((ck === '__all__' ? 'Total' : ck) + ' — ' + v.field + ' (' + AGG_LABELS[v.agg] + ')')));
  const rows = res.rowKeys.map(rk => {
    const row = [rk === '__all__' ? 'Total' : rk];
    res.colKeys.forEach(ck => res.values.forEach((v, vi) => row.push(res.grid[vi][rk][ck] || 0)));
    return row;
  });
  return { headers, rows };
}

function exportPivotCSV() {
  const grid = pivotToGrid();
  if (!grid) { toast('Build a pivot first.'); return; }
  downloadBlob(toCSV(grid.headers, grid.rows), 'pivot-export.csv', 'text/csv');
}

function chartCurrentPivot() {
  if (!pivotLastResult) { toast('Build a pivot first.'); return; }
  const res = pivotLastResult;
  const holder = document.getElementById('pivot-chart-holder');
  holder.style.display = '';
  const ctx = document.getElementById('pivot-chart-canvas').getContext('2d');
  if (pivotChart) pivotChart.destroy();

  const labels = res.rowKeys.map(rk => rk === '__all__' ? 'Total' : rk);
  const datasets = [];
  if (res.colKeys.length <= 1) {
    res.values.forEach((v, vi) => {
      datasets.push({ label: v.field + ' (' + AGG_LABELS[v.agg] + ')', data: res.rowKeys.map(rk => res.grid[vi][rk][res.colKeys[0]] || 0), backgroundColor: CHART_COLORS[vi % CHART_COLORS.length] });
    });
  } else {
    res.colKeys.forEach((ck, ci) => {
      datasets.push({ label: ck, data: res.rowKeys.map(rk => res.grid[0][rk][ck] || 0), backgroundColor: CHART_COLORS[ci % CHART_COLORS.length] });
    });
  }
  const chartType = (Pivot.rows.length && FIELD_KIND[Pivot.rows[0].field] === 'date') ? 'line' : 'bar';
  pivotChart = makeChart(ctx, { type: labels.length > 25 ? 'line' : chartType, data: { labels, datasets }, options: chartOptions() });
}

/** Chart.js na mile to app crash na ho — background mein load karke dobara draw karte hain. */
let chartReloadQueued = false;
function makeChart(ctx, cfg) {
  if (typeof Chart === 'undefined') {
    if (!chartReloadQueued) {
      chartReloadQueued = true;
      ensureChart().then(function (ok) {
        chartReloadQueued = false;
        if (ok) { renderDashboard(); renderPerformance(); }
      });
    }
    return null;
  }
  try { return new Chart(ctx, cfg); }
  catch (e) { console.error('Chart render failed', e); return null; }
}

function chartOptions(extra) {
  return Object.assign({
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { labels: { font: { family: "'IBM Plex Sans', sans-serif" } } } },
    scales: { x: { ticks: { font: { size: 10.5 } } }, y: { beginAtZero: true } }
  }, extra || {});
}

/* ---------------------------------------------------------------
   7b. ANALYSIS ENGINE — period filter + per-item performance
   --------------------------------------------------------------- */

// Period ka "aaj" browser ki date nahi, data ki last sale date hoti hai.
// ERP export purana ho to bhi "Last 30 days" sahi window pakde.
App.period = { mode: 'all', from: null, to: null };

function salesRecords() { return datasetsOfType('sales').flatMap(d => d.records); }
function purchaseRecords() { return datasetsOfType('purchase').flatMap(d => d.records); }
function stockRecords() { return datasetsOfType('stock').flatMap(d => d.records); }

/** Math.max(...arr) bade arrays par stack overflow de sakta hai (57,000+ rows),
 *  isliye loop se min/max nikalte hain. */
function minMaxTime(dates) {
  let mn = null, mx = null;
  for (let i = 0; i < dates.length; i++) {
    const t = dates[i].getTime();
    if (mn === null || t < mn) mn = t;
    if (mx === null || t > mx) mx = t;
  }
  return { min: mn, max: mx };
}

// Anchor date poore data par nikalti hai, isliye cache karte hain — warna har
// filter call par saara data dobara scan hota hai (popup 48 sec le raha tha).
let _anchorCache = null;
function clearAnchorCache() { _anchorCache = null; }

function dataAnchorDate() {
  if (_anchorCache) return _anchorCache;
  let dates = salesRecords().concat(purchaseRecords()).map(r => r.Date).filter(Boolean);
  if (!dates.length) {
    // Agar file "sales"/"purchase" type mein import nahi hui to bhi anchor data
    // se hi banni chahiye, warna aaj ki date lag jayegi aur "This week" khaali dikhega.
    dates = App.datasets
      .filter(d => d.fields.includes('Date'))
      .flatMap(d => d.records).map(r => r.Date).filter(Boolean);
  }
  if (!dates.length) return new Date();
  _anchorCache = new Date(minMaxTime(dates).max);
  return _anchorCache;
}

/** Jo data abhi dikh raha hai, wo asal mein kab se kab tak ka hai — ye batata hai.
 *  Filter "Last 30 days" ho ya "All data", user ko exact tareekh dikhni chahiye. */
function effectiveRange(range) {
  const recs = salesRecords().concat(purchaseRecords()).filter(r => inPeriod(r, range));
  const dates = recs.map(r => r.Date).filter(Boolean);
  if (!dates.length) {
    return { from: null, to: null, days: 0, rows: recs.length, text: 'no dated rows in this window' };
  }
  const mm = minMaxTime(dates);
  const from = new Date(mm.min), to = new Date(mm.max);
  const days = Math.max(1, Math.round((to - from) / 86400000) + 1);
  return {
    from, to, days, rows: recs.length,
    text: fmtDate(from) + '  \u2192  ' + fmtDate(to)
  };
}

/** Har toolbar ke neeche exact range likh deta hai. */
function updateRangeNotes() {
  const notes = document.querySelectorAll('.period-range-note');
  if (!notes.length) return;
  if (!App.datasets.length) {
    notes.forEach(n => { n.textContent = ''; });
    return;
  }
  const er = effectiveRange(periodRange());
  const txt = er.from
    ? 'Data: ' + er.text + '   \u00b7   ' + er.days + ' days   \u00b7   ' + er.rows.toLocaleString('en-IN') + ' rows'
    : er.text;
  notes.forEach(n => { n.textContent = txt; });
}

function periodRange() {
  const p = App.period;
  const anchor = dataAnchorDate();
  const endOfAnchor = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), anchor.getUTCDate()));
  if (p.mode === 'all') return { from: null, to: null, label: 'All data' };
  if (p.mode === 'custom') {
    const from = p.from ? parseDateLoose(p.from) : null;
    const to = p.to ? parseDateLoose(p.to) : null;
    const label = (from && to) ? (fmtDate(from) + ' → ' + fmtDate(to)) : 'Custom range';
    return { from, to, label };
  }
  if (p.mode === 'thismonth') {
    return { from: new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), 1)), to: endOfAnchor, label: 'This month' };
  }
  if (p.mode === 'thisyear') {
    return { from: new Date(Date.UTC(anchor.getUTCFullYear(), 0, 1)), to: endOfAnchor, label: 'This year' };
  }
  const days = parseInt(p.mode, 10);
  if (!isNaN(days)) {
    return { from: new Date(endOfAnchor.getTime() - (days - 1) * 86400000), to: endOfAnchor, label: 'Last ' + days + ' days' };
  }
  return { from: null, to: null, label: 'All data' };
}

function inPeriod(rec, range) {
  if (!range.from && !range.to) return true;
  const d = rec.Date;
  if (!d) return false;
  if (range.from && d < range.from) return false;
  if (range.to && d > range.to) return false;
  return true;
}

function periodDayCount(range, recs) {
  if (range.from && range.to) return Math.max(1, Math.round((range.to - range.from) / 86400000) + 1);
  const dates = recs.map(r => r.Date).filter(Boolean);
  if (!dates.length) return 30;
  const mm = minMaxTime(dates);
  const min = new Date(mm.min), max = new Date(mm.max);
  return Math.max(1, Math.round((max - min) / 86400000) + 1);
}

/** Row ki quantity. OBS/CBS report mein "Quantity" nahi hoti — wahan
 *  Closing Qty (CBS) hi asli balance hai, isliye uspar fallback karte hain. */
function recQty(r) {
  if (typeof r.Quantity === 'number') return r.Quantity;
  if (typeof r['Closing Qty'] === 'number') return r['Closing Qty'];
  return 0;
}

function recOpeningQty(r) {
  if (typeof r['Opening Qty'] === 'number') return r['Opening Qty'];
  return null;
}

function dimKey(rec, dim) {
  let key = rec[dim];
  if (key === null || key === undefined || key === '') key = resolveField(rec, dim);
  if (key === null || key === undefined || key === '') return '(blank)';
  if (key instanceof Date) return fmtDate(key);
  return String(key);
}

/* ---------------------------------------------------------------
   7c. RELATIONSHIPS — auto-detect + quality scoring
   --------------------------------------------------------------- */

/** Do datasets ke beech ek field par kitna overlap hai, ye naapta hai.
 *  Saath hi batata hai ki target field ek asli "key" hai ya nahi — yani uski
 *  values unique hain ya nahi. Ye zaroori hai: Section jaisa field 100% match
 *  dikhata hai lekin usse doosri file se value uthana galat hoga, kyunki ek
 *  Section ke andar hazaron alag-alag rows hoti hain. */
function scoreRelationship(fromDs, fromField, toDs, toField, sampleLimit) {
  const limit = sampleLimit || 4000;
  const toSet = new Set();
  let toNonNull = 0;
  for (const r of toDs.records) {
    const v = r[toField];
    if (v === null || v === undefined || v === '') continue;
    toNonNull++;
    toSet.add(String(v));
  }
  if (!toSet.size) return { matched: 0, checked: 0, pct: 0, toDistinct: 0, uniqueness: 0, canEnrich: false };

  let checked = 0, matched = 0;
  const fromSet = new Set();
  const step = Math.max(1, Math.floor(fromDs.records.length / limit));
  for (let i = 0; i < fromDs.records.length; i += step) {
    const v = fromDs.records[i][fromField];
    if (v === null || v === undefined || v === '') continue;
    checked++;
    fromSet.add(String(v));
    if (toSet.has(String(v))) matched++;
  }

  const uniqueness = toNonNull ? toSet.size / toNonNull : 0;
  return {
    matched, checked,
    pct: checked ? (matched / checked) * 100 : 0,
    toDistinct: toSet.size,
    fromDistinct: fromSet.size,
    uniqueness,
    // Enrichment (dusri file se column uthana) sirf tab safe hai jab target
    // field lagbhag unique ho — warna kaunsi row uthaayen, ye tay hi nahi hota.
    canEnrich: uniqueness >= 0.9
  };
}

function relationshipId(a, b, c, d) { return [a, b, c, d].join('~'); }

/** Har dataset-pair ke liye sabse achha joining field dhoondta hai.
 *  Sirf match % dekhna kaafi nahi — kam distinct values wala field (jaise
 *  Section) 100% match dikhata hai par join ke liye bekaar hai. Isliye
 *  cardinality ko bhi weight dete hain. */
function suggestRelationships() {
  const out = [];
  const KEY_PREFERENCE = ['Item Code', 'Article No', 'Style', 'Brand', 'Supplier', 'Sub Section', 'Section'];
  for (let i = 0; i < App.datasets.length; i++) {
    for (let j = i + 1; j < App.datasets.length; j++) {
      const a = App.datasets[i], b = App.datasets[j];
      const shared = a.fields.filter(f => b.fields.indexOf(f) !== -1 && FIELD_KIND[f] !== 'number' && FIELD_KIND[f] !== 'date');
      if (!shared.length) continue;
      const candidates = [];
      shared.forEach(f => {
        const sc = scoreRelationship(a, f, b, f);
        if (sc.toDistinct < 5) return;                 // 4 values wala field key nahi hota
        const card = sc.toDistinct >= 500 ? 1 : (sc.toDistinct >= 100 ? 0.9 : (sc.toDistinct >= 20 ? 0.7 : 0.4));
        const pref = KEY_PREFERENCE.indexOf(f);
        const prefBonus = pref >= 0 ? (KEY_PREFERENCE.length - pref) * 2 : 0;
        candidates.push({ field: f, score: sc, rank: sc.pct * card + prefBonus });
      });
      if (!candidates.length) continue;

      // Jo connection sach mein kaam ki hai (jiske through column lookup ho
      // sakta hai) usko pehle chunte hain — chahe uska match % kam ho.
      const enrichable = candidates.filter(c => c.score.canEnrich && c.score.pct >= 10);
      const pool = enrichable.length ? enrichable : candidates;
      const best = pool.reduce((x, y) => (y.rank > x.rank ? y : x));

      if (best && best.score.pct >= 20) {
        out.push({
          id: relationshipId(a.id, best.field, b.id, best.field),
          fromDsId: a.id, fromField: best.field,
          toDsId: b.id, toField: best.field,
          enabled: true, score: best.score
        });
      }
    }
  }
  return out;
}

function autoDetectRelationships(silent) {
  const found = suggestRelationships();
  let added = 0;
  found.forEach(r => {
    if (!App.relationships.some(x => x.id === r.id)) { App.relationships.push(r); added++; }
  });
  clearLookups();
  if (!silent) toast(added ? added + ' connection(s) detected.' : 'No new connections were found.');
  return added;
}

function rescoreRelationships() {
  App.relationships.forEach(rel => {
    const a = App.datasets.find(d => d.id === rel.fromDsId);
    const b = App.datasets.find(d => d.id === rel.toDsId);
    rel.score = (a && b) ? scoreRelationship(a, rel.fromField, b, rel.toField) : { pct: 0, matched: 0, checked: 0 };
  });
}

/**
 * Poora per-item picture: kitna bika, kitna stock hai, aakhri bikri kab hui,
 * stock kitna purana hai, sell-through kya hai, ABC class kya hai, aur
 * status (Best seller / Slow / Dead stock / Out of stock).
 */
function buildAnalysis(dim, targetDays) {
  const range = periodRange();
  const sales = salesRecords().filter(r => inPeriod(r, range));
  const purchases = purchaseRecords().filter(r => inPeriod(r, range));
  const stock = stockRecords();
  const days = periodDayCount(range, sales.length ? sales : purchases);
  const anchor = dataAnchorDate();

  const map = new Map();
  function slot(key) {
    let s = map.get(key);
    if (!s) {
      s = { key, sold: 0, purchased: 0, stock: 0, opening: 0, hasOpening: false, saleLines: 0,
            firstSale: null, lastSale: null, oldestStock: null, newestStock: null,
            meta: {} };
      map.set(key, s);
    }
    return s;
  }

  sales.forEach(r => {
    const s = slot(dimKey(r, dim));
    const q = recQty(r);
    s.sold += q; s.saleLines++;
    if (r.Date) {
      if (!s.firstSale || r.Date < s.firstSale) s.firstSale = r.Date;
      if (!s.lastSale || r.Date > s.lastSale) s.lastSale = r.Date;
    }
    captureMeta(s, r);
  });

  purchases.forEach(r => {
    const s = slot(dimKey(r, dim));
    s.purchased += recQty(r);
    captureMeta(s, r);
  });

  stock.forEach(r => {
    const s = slot(dimKey(r, dim));
    const q = recQty(r);
    s.stock += q;                                  // CBS = closing balance
    const ob = recOpeningQty(r);
    if (ob !== null) { s.opening += ob; s.hasOpening = true; }   // OBS
    const bd = r['Purchase Bill Date'];
    if (bd) {
      if (!s.oldestStock || bd < s.oldestStock) s.oldestStock = bd;
      if (!s.newestStock || bd > s.newestStock) s.newestStock = bd;
      s.ageQtySum = (s.ageQtySum || 0) + q;
      s.ageWeighted = (s.ageWeighted || 0) + q * bd.getTime();
    }
    captureMeta(s, r);
  });

  let rows = [...map.values()].map(s => {
    const avgDaily = s.sold / days;
    const daysCover = avgDaily > 0 ? s.stock / avgDaily : (s.stock > 0 ? Infinity : 0);
    const suggested = Math.max(0, Math.round(avgDaily * targetDays - s.stock));
    const opening = s.sold + s.stock;
    const sellThrough = opening > 0 ? (s.sold / opening) * 100 : 0;
    const daysSinceLastSale = s.lastSale ? Math.round((anchor - s.lastSale) / 86400000) : null;
    // Stock age = quantity-weighted average lot age, not just the oldest piece.
    const stockAgeDays = (s.ageQtySum > 0)
      ? Math.round((anchor - (s.ageWeighted / s.ageQtySum)) / 86400000)
      : (s.oldestStock ? Math.round((anchor - s.oldestStock) / 86400000) : null);
    const oldestAgeDays = s.oldestStock ? Math.round((anchor - s.oldestStock) / 86400000) : null;
    // OBS - CBS = is period mein stock kitna ghata (ya badha)
    const movement = s.hasOpening ? (s.opening - s.stock) : null;
    return Object.assign({}, s, { avgDaily, daysCover, suggested, sellThrough, daysSinceLastSale, stockAgeDays, oldestAgeDays, movement });
  });

  // ABC classification — Pareto on quantity sold
  const totalSold = rows.reduce((a, r) => a + r.sold, 0);
  rows.sort((a, b) => b.sold - a.sold);
  let cum = 0;
  rows.forEach(r => {
    if (totalSold <= 0) { r.abc = '—'; r.cumPct = 0; return; }
    cum += r.sold;
    r.cumPct = (cum / totalSold) * 100;
    r.abc = r.sold === 0 ? '—' : (r.cumPct <= 80 ? 'A' : (r.cumPct <= 95 ? 'B' : 'C'));
  });

  // Status
  const DEAD_DAYS = 90;
  rows.forEach(r => {
    if (r.sold === 0 && r.stock > 0) r.status = 'Dead stock';
    else if (r.stock === 0 && r.sold > 0) r.status = 'Out of stock';
    else if (r.daysSinceLastSale !== null && r.daysSinceLastSale > DEAD_DAYS && r.stock > 0) r.status = 'Dead stock';
    else if (r.abc === 'A') r.status = 'Best seller';
    else if (r.abc === 'B') r.status = 'Steady';
    else if (r.sold > 0) r.status = 'Slow mover';
    else r.status = 'No activity';
    r.reorder = r.suggested >= 1;
    // Overstocked = itna stock ki target cover se kai guna zyada chal jaye.
    r.overstocked = r.stock > 0 && r.daysCover > targetDays * 3;
    r.excessQty = r.overstocked && r.avgDaily > 0
      ? Math.max(0, Math.round(r.stock - r.avgDaily * targetDays))
      : (r.sold === 0 ? r.stock : 0);
  });

  return { rows, days, range, anchor, totalSold };
}

// Item-level rows par brand/section/supplier bhi dikha saken isliye pehli
// non-empty value yaad rakh lete hain.
const META_FIELDS = ['Brand', 'Section', 'Sub Section', 'Supplier', 'Style', 'Colour', 'Size', 'Article No'];
function captureMeta(slotObj, rec) {
  META_FIELDS.forEach(f => {
    if (!slotObj.meta[f] && rec[f]) slotObj.meta[f] = rec[f] instanceof Date ? fmtDate(rec[f]) : String(rec[f]);
  });
}

function periodSelectHtml(id) {
  const p = App.period.mode;
  const opts = [['all', 'All data'], ['30', 'Last 30 days'], ['90', 'Last 90 days'],
                ['180', 'Last 180 days'], ['365', 'Last 365 days'],
                ['thismonth', 'Latest month'], ['thisyear', 'Latest year'],
                ['custom', 'Custom range…']];
  const isCustom = p === 'custom';
  return '<select id="' + id + '" class="select period-select">' +
    opts.map(([v, l]) => '<option value="' + v + '"' + (p === v ? ' selected' : '') + '>' + l + '</option>').join('') +
    '</select>' +
    '<span class="period-custom-range" style="' + (isCustom ? '' : 'display:none;') + '">' +
      '<input type="date" class="text-input period-from" value="' + (App.period.from || '') + '"> to ' +
      '<input type="date" class="text-input period-to" value="' + (App.period.to || '') + '">' +
    '</span>';
}

/** Purane markup mein "Custom range" option aur date boxes nahi hote — yahan
 *  khud jod dete hain, taaki sirf app.js update karne se bhi feature mile. */
function ensurePeriodCustomInputs() {
  document.querySelectorAll('.period-select').forEach(sel => {
    if (![...sel.options].some(o => o.value === 'custom')) {
      const opt = document.createElement('option');
      opt.value = 'custom';
      opt.textContent = 'Custom range\u2026';
      sel.appendChild(opt);
    }
    const nxt = sel.nextElementSibling;
    if (!nxt || !nxt.classList.contains('period-custom-range')) {
      const span = document.createElement('span');
      span.className = 'period-custom-range date-range-inline';
      span.style.display = sel.value === 'custom' ? '' : 'none';
      span.innerHTML = '<span class="dr-label">From</span><input type="date" class="text-input period-from">' +
                       '<span class="dr-label">To</span><input type="date" class="text-input period-to">';
      sel.parentNode.insertBefore(span, sel.nextSibling);
    }
    // exact date range note
    const toolbar = sel.closest('.toolbar');
    if (toolbar && !toolbar.querySelector('.period-range-note')) {
      const note = document.createElement('div');
      note.className = 'period-range-note';
      toolbar.appendChild(note);
    }
  });
  updateRangeNotes();
}

function wirePeriodSelects() {
  ensurePeriodCustomInputs();
  document.querySelectorAll('.period-select').forEach(sel => {
    sel.addEventListener('change', () => {
      App.period.mode = sel.value;
      document.querySelectorAll('.period-select').forEach(s => { s.value = sel.value; });
      document.querySelectorAll('.period-custom-range').forEach(el => {
        el.style.display = sel.value === 'custom' ? '' : 'none';
      });
      if (sel.value !== 'custom' || (App.period.from && App.period.to)) renderAllPeriodViews();
    });
  });
  document.querySelectorAll('.period-from').forEach(inp => {
    inp.addEventListener('change', function () {
      App.period.from = this.value || null;
      document.querySelectorAll('.period-from').forEach(el => { el.value = App.period.from || ''; });
      if (App.period.from && App.period.to) renderAllPeriodViews();
    });
  });
  document.querySelectorAll('.period-to').forEach(inp => {
    inp.addEventListener('change', function () {
      App.period.to = this.value || null;
      document.querySelectorAll('.period-to').forEach(el => { el.value = App.period.to || ''; });
      if (App.period.from && App.period.to) renderAllPeriodViews();
    });
  });
}

function renderAllPeriodViews() {
  updateRangeNotes();
  renderInsights();
  renderPerformance();
  renderCatalog();
  renderDashboard();
  if (typeof Drill !== 'undefined' && Drill.open) renderDrill();
}

/* ---------------------------------------------------------------
   7d. QUICK REPORT — checkbox se nested report (Google Sheet wale
       R-Data pivot jaisa: tick karo, click ka order hi grouping order)
   --------------------------------------------------------------- */
const QuickReport = {
  order: [],            // tick karne ka kram — yahi grouping ka kram hai
  desc: true,
  measure: 'Quantity',
  agg: 'sum',
  collapsed: {},
  lastRows: null,
  dateGrain: 'none',    // none | day | week | month | quarter | year
  period: { mode: 'all', from: null, to: null },
  filters: {},          // column -> Set(values)  (Excel jaisa column filter)
  search: '',
  sourcePicked: false   // user ne khud data source chuna hai ya nahi
};

// Date se nikle hue extra grouping options — ab ye tick-list mein nahi,
// alag "Date grouping" dropdown se aate hain.
const DERIVED_DIMS = ['Year', 'Quarter', 'Month', 'Week', 'Transaction Date'];
const GRAIN_TO_DIM = { day: 'Transaction Date', week: 'Week', month: 'Month', quarter: 'Quarter', year: 'Year' };

function quickReportFields() {
  const base = currentQuickFields();
  const out = [];
  base.forEach(f => {
    if (f === 'Date' || f === 'Purchase Bill Date') return;  // date ab alag dropdown se
    if (FIELD_KIND[f] === 'number') return;                  // measure hai, grouping nahi
    out.push(f);
  });
  return out;
}

function currentQuickFields() {
  const sel = document.getElementById('quick-dataset-select');
  const v = sel ? sel.value : '__all__';
  if (v === '__all__') return allLoadedFields();
  if (v.startsWith('__type:')) {
    const t = v.slice(7, -2); const set = new Set();
    datasetsOfType(t).forEach(d => d.fields.forEach(f => set.add(f)));
    return CANONICAL_FIELDS.filter(f => set.has(f));
  }
  const ds = App.datasets.find(d => d.id === v);
  return ds ? ds.fields : [];
}

function weekKeyOf(d) {
  const mon = mondayOfWeekUTC(new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())));
  const sun = new Date(mon.getTime() + 6 * 86400000);
  return fmtDate(mon) + ' \u2013 ' + fmtDate(sun);
}

function quickGroupKey(rec, dim) {
  if (dim === 'Year') return rec.Date ? String(rec.Date.getUTCFullYear()) : '(blank)';
  if (dim === 'Quarter') return rec.Date ? dateKeyForGrain(rec.Date, 'quarter') : '(blank)';
  if (dim === 'Month') return rec.Date ? dateKeyForGrain(rec.Date, 'month') : '(blank)';
  if (dim === 'Week') return rec.Date ? weekKeyOf(rec.Date) : '(blank)';
  if (dim === 'Transaction Date') return rec.Date ? fmtDate(rec.Date) : '(blank)';
  return dimKey(rec, dim);
}

function quickSortValue(dim, key) {
  if (dim === 'Month') return grainSort(key, 'month');
  if (dim === 'Year') return parseInt(key, 10) || 0;
  if (dim === 'Quarter') { const m = key.match(/^(\d{4}) Q(\d)$/); return m ? parseInt(m[1],10)*10 + parseInt(m[2],10) : 0; }
  if (dim === 'Week') { const d = parseDateLoose(key.split(' \u2013 ')[0]); return d ? d.getTime() : 0; }
  if (dim === 'Transaction Date') { const d = parseDateLoose(key); return d ? d.getTime() : 0; }
  return null;
}

function initQuickReport() {
  const modeBtns = document.querySelectorAll('#pivot-mode .seg-btn');
  modeBtns.forEach(btn => btn.addEventListener('click', () => {
    modeBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const quick = btn.dataset.mode === 'quick';
    document.getElementById('pivot-body-builder').style.display = quick ? 'none' : '';
    document.getElementById('pivot-body-quick').style.display = quick ? '' : 'none';
    if (quick) renderQuickReport();
  }));

  document.getElementById('quick-dataset-select').addEventListener('change', () => {
    QuickReport.sourcePicked = true;
    QuickReport.order = []; QuickReport.filters = {};
    // The measure list depends on the chosen file: a stock/OBS-CBS report has
    // Opening/Closing Qty but no "Quantity", so it must be rebuilt here.
    refreshQuickMeasures();
    renderQuickFieldList(); renderQuickReport();
  });
  document.getElementById('quick-measure').addEventListener('change', e => { QuickReport.measure = e.target.value; renderQuickReport(); });
  document.getElementById('quick-agg').addEventListener('change', e => { QuickReport.agg = e.target.value; renderQuickReport(); });
  document.getElementById('quick-sort').addEventListener('change', e => { QuickReport.desc = e.target.value === 'desc'; renderQuickReport(); });
  document.getElementById('quick-clear').addEventListener('click', () => {
    QuickReport.order = []; QuickReport.collapsed = {}; QuickReport.filters = {};
    QuickReport.search = '';
    const si = document.getElementById('quick-search'); if (si) si.value = '';
    renderQuickFieldList(); renderQuickReport();
  });
  const grainSel = document.getElementById('quick-date-grain');
  if (grainSel) grainSel.addEventListener('change', e => {
    QuickReport.dateGrain = e.target.value; QuickReport.collapsed = {}; renderQuickReport();
  });
  const perSel = document.getElementById('quick-period');
  if (perSel) perSel.addEventListener('change', e => {
    QuickReport.period.mode = e.target.value;
    const cr = document.getElementById('quick-custom-range');
    if (cr) cr.style.display = e.target.value === 'custom' ? '' : 'none';
    QuickReport.collapsed = {};
    if (e.target.value !== 'custom' || (QuickReport.period.from && QuickReport.period.to)) renderQuickReport();
  });
  ['quick-from', 'quick-to'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('change', () => {
      QuickReport.period.from = document.getElementById('quick-from').value || null;
      QuickReport.period.to = document.getElementById('quick-to').value || null;
      QuickReport.collapsed = {};
      if (QuickReport.period.from && QuickReport.period.to) renderQuickReport();
    });
  });
  const qs = document.getElementById('quick-search');
  if (qs) qs.addEventListener('input', debounce(e => {
    QuickReport.search = e.target.value; QuickReport.collapsed = {}; renderQuickReport();
  }, 220));

  const qea = document.getElementById('quick-expand-all');
  if (qea) qea.addEventListener('click', () => {
    (QuickReport.lastPaths || []).forEach(p => { QuickReport.collapsed[p] = false; });
    renderQuickReport();
  });
  const qca = document.getElementById('quick-collapse-all');
  if (qca) qca.addEventListener('click', () => { QuickReport.collapsed = {}; renderQuickReport(); });
  document.getElementById('quick-export').addEventListener('click', exportQuickCSV);
  const toSheet = document.getElementById('quick-to-sheet');
  if (toSheet) toSheet.addEventListener('click', quickToSheet);
}

function populateQuickSelects() {
  const el = document.getElementById('quick-dataset-select');
  if (!el) return;
  const opts = ['<option value="__all__">\u26A0 All files combined (mixed)</option>']
    .concat(['sales', 'purchase', 'stock'].filter(t => datasetsOfType(t).length).map(t =>
      '<option value="__type:' + t + '__">' + t.charAt(0).toUpperCase() + t.slice(1) + ' data</option>'))
    .concat(App.datasets.map(ds => '<option value="' + ds.id + '">' + escapeHtml(ds.name) + '</option>'));
  const prev = el.value;
  el.innerHTML = opts.join('');
  // User ne khud choose kiya ho to wahi rakho. Warna auto-pick — "All combined"
  // default NEVER, kyunki wo Sales+Purchase+Stock ki qty jod deta hai jiska
  // koi business matlab nahi banta.
  if (QuickReport.sourcePicked && prev && [...el.options].some(o => o.value === prev)) {
    el.value = prev;
  } else if (datasetsOfType('sales').length) el.value = '__type:sales__';
  else if (datasetsOfType('purchase').length) el.value = '__type:purchase__';
  else if (datasetsOfType('stock').length) el.value = '__type:stock__';
  else if (App.datasets.length) el.value = App.datasets[0].id;

  refreshQuickMeasures();
}

/** Measure dropdown ko current data source ke hisaab se banata hai. */
function refreshQuickMeasures() {
  const mEl = document.getElementById('quick-measure');
  if (!mEl) return;
  const nums = currentQuickFields().filter(f => FIELD_KIND[f] === 'number');
  const list = nums.length ? nums : ['Quantity'];
  const pm = mEl.value;
  mEl.innerHTML = list.map(f => '<option value="' + f + '">' + f + '</option>').join('');
  // Keep the previous measure if it still exists; otherwise prefer the
  // closing balance for stock reports, else the first available number.
  if (list.indexOf(pm) !== -1) mEl.value = pm;
  else if (list.indexOf('Closing Qty') !== -1) mEl.value = 'Closing Qty';
  else if (list.indexOf('Quantity') !== -1) mEl.value = 'Quantity';
  QuickReport.measure = mEl.value;
}

function renderQuickFieldList() {
  const wrap = document.getElementById('quick-field-list');
  if (!wrap) return;
  const fields = quickReportFields();
  if (!fields.length) { wrap.innerHTML = '<div class="empty-hint">Load a file first.</div>'; return; }

  wrap.innerHTML = fields.map(f => {
    const idx = QuickReport.order.indexOf(f);
    const on = idx !== -1;
    return '<label class="qr-check' + (on ? ' on' : '') + '">' +
      '<input type="checkbox" value="' + escapeHtml(f) + '"' + (on ? ' checked' : '') + '>' +
      '<span class="qr-name">' + escapeHtml(f) + '</span>' +
      (on ? '<span class="qr-order">' + (idx + 1) + '</span>' : '') +
    '</label>';
  }).join('');

  wrap.querySelectorAll('input').forEach(cb => cb.addEventListener('change', () => {
    const f = cb.value;
    if (cb.checked) { if (!QuickReport.order.includes(f)) QuickReport.order.push(f); }
    else QuickReport.order = QuickReport.order.filter(x => x !== f);
    QuickReport.collapsed = {};
    renderQuickFieldList();
    renderQuickReport();
  }));
}

/** Nested groups + har level par subtotal — bilkul sheet wale pivot jaisa. */
function buildQuickTree(recs, dims, depth, parentPath) {
  if (depth >= dims.length) return null;
  const dim = dims[depth];
  const groups = new Map();
  recs.forEach(r => {
    const k = quickGroupKey(r, dim);
    let g = groups.get(k);
    if (!g) { g = { key: k, recs: [], total: 0, count: 0 }; groups.set(k, g); }
    g.recs.push(r);
    const v = r[QuickReport.measure];
    if (typeof v === 'number') { g.total += v; g.count++; }
    else if (v !== null && v !== undefined) g.count++;
  });

  let list = [...groups.values()].map(g => {
    const path = parentPath ? parentPath + '|' + g.key : g.key;
    let value = g.total;
    if (QuickReport.agg === 'count') value = g.recs.length;
    else if (QuickReport.agg === 'avg') value = g.count ? g.total / g.count : 0;
    return {
      dim, key: g.key, value, rows: g.recs.length, path, depth,
      children: buildQuickTree(g.recs, dims, depth + 1, path)
    };
  });

  const natural = quickSortValue(dim, list.length ? list[0].key : '');
  if (natural !== null) {
    list.sort((a, b) => (quickSortValue(dim, a.key) - quickSortValue(dim, b.key)) * (QuickReport.desc ? -1 : 1));
  } else {
    list.sort((a, b) => (a.value - b.value) * (QuickReport.desc ? -1 : 1));
  }
  return list;
}

/** Quick Report ka apna date window — dashboard wale filter se alag rehta hai. */
function quickPeriodRange() {
  const p = QuickReport.period;
  const anchor = dataAnchorDate();
  const mid = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), anchor.getUTCDate()));
  const mk = (from, to, label) => ({ from, to, label });

  if (p.mode === 'all') return mk(null, null, 'All data');
  if (p.mode === 'custom') {
    const from = p.from ? parseDateLoose(p.from) : null;
    const to = p.to ? parseDateLoose(p.to) : null;
    return mk(from, to, (from && to) ? (fmtDate(from) + ' \u2192 ' + fmtDate(to)) : 'Custom range');
  }
  if (p.mode === 'thisweek') {
    const mon = mondayOfWeekUTC(mid), sun = new Date(mon.getTime() + 6 * 86400000);
    return mk(mon, sun, 'This week: ' + fmtDate(mon) + ' \u2192 ' + fmtDate(sun));
  }
  if (p.mode === 'lastweek') {
    const cur = mondayOfWeekUTC(mid);
    const mon = new Date(cur.getTime() - 7 * 86400000), sun = new Date(cur.getTime() - 86400000);
    return mk(mon, sun, 'Last week: ' + fmtDate(mon) + ' \u2192 ' + fmtDate(sun));
  }
  if (p.mode === 'thismonth') {
    const from = new Date(Date.UTC(mid.getUTCFullYear(), mid.getUTCMonth(), 1));
    return mk(from, mid, 'This month: ' + fmtDate(from) + ' \u2192 ' + fmtDate(mid));
  }
  if (p.mode === 'lastmonth') {
    const from = new Date(Date.UTC(mid.getUTCFullYear(), mid.getUTCMonth() - 1, 1));
    const to = new Date(Date.UTC(mid.getUTCFullYear(), mid.getUTCMonth(), 0));
    return mk(from, to, 'Last month: ' + fmtDate(from) + ' \u2192 ' + fmtDate(to));
  }
  if (p.mode === 'thisyear') {
    const from = new Date(Date.UTC(mid.getUTCFullYear(), 0, 1));
    return mk(from, mid, 'This year: ' + fmtDate(from) + ' \u2192 ' + fmtDate(mid));
  }
  const days = parseInt(p.mode, 10);
  if (!isNaN(days)) {
    const from = new Date(mid.getTime() - (days - 1) * 86400000);
    return mk(from, mid, 'Last ' + days + ' days: ' + fmtDate(from) + ' \u2192 ' + fmtDate(mid));
  }
  return mk(null, null, 'All data');
}

/** Source rows par date window + column filters + search lagata hai. */
function quickFilteredRecords() {
  const sel = document.getElementById('quick-dataset-select');
  let recs = getRecordsForSelection(sel ? sel.value : '__all__');

  const range = quickPeriodRange();
  if (range.from || range.to) recs = recs.filter(r => (r.Date ? inPeriod(r, range) : false));

  const cols = Object.keys(QuickReport.filters);
  if (cols.length) {
    recs = recs.filter(r => cols.every(c => QuickReport.filters[c].has(quickGroupKey(r, c))));
  }

  const q = (QuickReport.search || '').trim().toLowerCase();
  if (q) {
    const dims = QuickReport.order.slice();
    recs = recs.filter(r => dims.some(d => quickGroupKey(r, d).toLowerCase().includes(q)));
  }
  return recs;
}

/** Kisi column ke saare values (count ke saath) — filter popup ke liye. */
function quickColumnValues(dim) {
  const sel = document.getElementById('quick-dataset-select');
  let recs = getRecordsForSelection(sel ? sel.value : '__all__');
  const range = quickPeriodRange();
  if (range.from || range.to) recs = recs.filter(r => (r.Date ? inPeriod(r, range) : false));
  // baaki columns ke filters lagey rahen taaki list relevant rahe
  Object.keys(QuickReport.filters).forEach(c => {
    if (c === dim) return;
    recs = recs.filter(r => QuickReport.filters[c].has(quickGroupKey(r, c)));
  });
  const counts = new Map();
  recs.forEach(r => { const k = quickGroupKey(r, dim); counts.set(k, (counts.get(k) || 0) + 1); });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

/** Excel jaisa column filter popup — search box ke saath. */
function openQuickColumnFilter(dim, anchorEl) {
  const values = quickColumnValues(dim);
  const active = QuickReport.filters[dim];

  const pop = document.createElement('div');
  pop.className = 'modal-backdrop';
  pop.innerHTML = '<div class="modal-box qcf-box">' +
    '<h3>Filter: ' + escapeHtml(dim) + '</h3>' +
    '<input type="search" id="qcf-search" class="text-input" placeholder="Type to search values\u2026" style="width:100%;margin-bottom:8px;">' +
    '<div class="qcf-actions">' +
    '<button class="ghost-btn small primary" id="qcf-only" title="Keep only the values matching your search">Only these</button>' +
    '<button class="ghost-btn small" id="qcf-all">Select all</button>' +
    '<button class="ghost-btn small" id="qcf-none">Clear</button>' +
    '<span class="qcf-count" id="qcf-count"></span></div>' +
    '<div id="qcf-list" class="efp-values"></div>' +
    '<div class="modal-actions"><button class="ghost-btn primary small" id="qcf-apply">Apply</button>' +
    '<button class="ghost-btn small" id="qcf-clear">Remove filter</button>' +
    '<span class="spacer"></span><button class="ghost-btn small" id="qcf-cancel">Cancel</button></div></div>';
  document.body.appendChild(pop);

  const listEl = pop.querySelector('#qcf-list');
  const searchEl = pop.querySelector('#qcf-search');
  const checkedNow = new Set(active ? [...active] : values.map(v => v[0]));

  function paint() {
    const q = searchEl.value.trim().toLowerCase();
    const shown = q ? values.filter(([v]) => v.toLowerCase().includes(q)) : values;
    pop.querySelector('#qcf-count').textContent = shown.length.toLocaleString('en-IN') + ' of ' + values.length.toLocaleString('en-IN') + ' values';
    listEl.innerHTML = shown.slice(0, 800).map(([v, c]) =>
      '<label class="efp-row"><input type="checkbox" value="' + escapeHtml(v) + '"' + (checkedNow.has(v) ? ' checked' : '') + '>' +
      '<span>' + escapeHtml(v) + '</span><span class="efp-count">' + c.toLocaleString('en-IN') + '</span></label>').join('') +
      (shown.length > 800 ? '<div class="empty-hint">Showing the top 800 - use search to narrow further.</div>' : '');
    listEl.querySelectorAll('input').forEach(cb => cb.addEventListener('change', () => {
      if (cb.checked) checkedNow.add(cb.value); else checkedNow.delete(cb.value);
    }));
  }
  paint();
  searchEl.addEventListener('input', debounce(paint, 150));
  searchEl.focus();

  // "Only these" — bade data mein sabse kaam ka: search karo, ek click,
  // sirf wahi values rah jaati hain.
  pop.querySelector('#qcf-only').onclick = () => {
    const q = searchEl.value.trim().toLowerCase();
    const shown = q ? values.filter(([v]) => v.toLowerCase().includes(q)) : values;
    checkedNow.clear();
    shown.forEach(([v]) => checkedNow.add(v));
    paint();
  };
  pop.querySelector('#qcf-all').onclick = () => { values.forEach(([v]) => checkedNow.add(v)); paint(); };
  pop.querySelector('#qcf-none').onclick = () => { checkedNow.clear(); paint(); };
  pop.querySelector('#qcf-cancel').onclick = () => pop.remove();
  pop.addEventListener('click', e => { if (e.target === pop) pop.remove(); });
  pop.querySelector('#qcf-clear').onclick = () => {
    delete QuickReport.filters[dim];
    pop.remove(); QuickReport.collapsed = {}; renderQuickReport();
  };
  pop.querySelector('#qcf-apply').onclick = () => {
    if (checkedNow.size === 0) { toast('Select at least one value.'); return; }
    if (checkedNow.size === values.length) delete QuickReport.filters[dim];
    else QuickReport.filters[dim] = new Set(checkedNow);
    pop.remove(); QuickReport.collapsed = {}; renderQuickReport();
  };
}

/** Active column filters ko chips ki tarah dikhata hai. */
function renderQuickFilterChips() {
  const wrap = document.getElementById('quick-filter-chips');
  if (!wrap) return;
  const cols = Object.keys(QuickReport.filters);
  if (!cols.length && !QuickReport.search) { wrap.innerHTML = ''; return; }
  wrap.innerHTML =
    (QuickReport.search ? '<span class="filter-chip">Search: "' + escapeHtml(QuickReport.search) + '"<button data-clear="__search">&times;</button></span>' : '') +
    cols.map(c => '<span class="filter-chip">' + escapeHtml(c) + ': ' + QuickReport.filters[c].size +
      ' selected<button data-clear="' + escapeHtml(c) + '">&times;</button></span>').join('') +
    (cols.length > 1 ? '<button class="ghost-btn small" id="qr-clear-filters">Clear all filters</button>' : '');

  wrap.querySelectorAll('[data-clear]').forEach(b => b.addEventListener('click', () => {
    if (b.dataset.clear === '__search') {
      QuickReport.search = '';
      const si = document.getElementById('quick-search'); if (si) si.value = '';
    } else delete QuickReport.filters[b.dataset.clear];
    QuickReport.collapsed = {};
    renderQuickReport();
  }));
  const clearAll = wrap.querySelector('#qr-clear-filters');
  if (clearAll) clearAll.addEventListener('click', () => {
    QuickReport.filters = {}; QuickReport.collapsed = {}; renderQuickReport();
  });
}

/** Abhi kaunsa data dikh raha hai — Sales / Purchase / Stock — aur agar
 *  mixed hai to saaf warning, kyunki alag-alag type ki qty jodna galat hai. */
function quickSourceInfo() {
  const sel = document.getElementById('quick-dataset-select');
  const v = sel ? sel.value : '__all__';
  let types = [], label = '';

  if (v === '__all__') {
    types = [...new Set(App.datasets.map(d => d.type))];
    label = 'All files combined';
  } else if (v.startsWith('__type:')) {
    const t = v.slice(7, -2);
    types = [t];
    label = t.charAt(0).toUpperCase() + t.slice(1) + ' data';
  } else {
    const ds = App.datasets.find(d => d.id === v);
    if (ds) { types = [ds.type]; label = ds.name; }
  }
  return { types, label, mixed: types.length > 1 };
}

function renderQuickSourceNote(rowCount) {
  const el = document.getElementById('quick-source-note');
  if (!el) return;
  const info = quickSourceInfo();
  const cls = { sales: 'tag-sales', purchase: 'tag-purchase', stock: 'tag-stock' }[info.types[0]] || 'tag-other';

  el.innerHTML =
    '<span class="qsrc-label">Showing:</span> ' +
    '<span class="sd-type-tag ' + (info.mixed ? 'tag-other' : cls) + '">' + escapeHtml(info.label) + '</span> ' +
    '<span class="qsrc-count">' + rowCount.toLocaleString('en-IN') + ' rows</span>';

  const warn = document.getElementById('quick-mixed-warning');
  if (warn) {
    if (info.mixed) {
      warn.style.display = '';
      warn.textContent = 'Note: this is "All files combined" - Sales, Purchase and Stock quantities are being added together, ' +
        'which is not meaningful (sold + purchased + on-hand in a single number). ' +
        'Use the "Data" dropdown above to pick Sales, Purchase or Stock separately.';
    } else warn.style.display = 'none';
  }
}

function renderQuickReport() {
  const table = document.getElementById('quick-table');
  const head = document.getElementById('quick-heading');
  if (!table) return;

  if (!App.datasets.length) {
    table.innerHTML = '<tr><td class="empty-hint">Load data from the Import tab first.</td></tr>';
    head.textContent = ''; QuickReport.lastRows = null; return;
  }
  // Date grouping dropdown se aaya level hamesha sabse pehle lagta hai
  const dims = QuickReport.order.slice();
  const grainDim = GRAIN_TO_DIM[QuickReport.dateGrain];
  if (grainDim) dims.unshift(grainDim);

  const range = quickPeriodRange();
  const noteEl = document.getElementById('quick-range-note');
  if (noteEl) noteEl.textContent = 'Showing: ' + range.label;

  if (!dims.length) {
    table.innerHTML = '<tr><td class="empty-hint">Tick columns on the left, or pick a "Date grouping" above. Columns group in the order you tick them.</td></tr>';
    head.textContent = ''; QuickReport.lastRows = null;
    renderQuickSourceNote(quickFilteredRecords().length);
    renderQuickFilterChips();
    return;
  }

  head.textContent = dims.join('  \u2794  ');

  const recs = quickFilteredRecords();
  renderQuickSourceNote(recs.length);
  const tree = buildQuickTree(recs, dims, 0, '') || [];
  QuickReport.lastPaths = (function collect(ns, out) {
    (ns || []).forEach(n => { if (n.children && n.children.length) { out.push(n.path); collect(n.children, out); } });
    return out;
  })(tree, []);
  renderQuickFilterChips();

  const srcInfo = quickSourceInfo();
  const label = AGG_LABELS[QuickReport.agg] + ' of ' + QuickReport.measure +
                (srcInfo.mixed ? '' : ' (' + srcInfo.types[0] + ')');
  const grand = tree.reduce((s, n) => s + n.value, 0);

  let body = '';
  const flat = [];
  const MAX_RENDER = 4000;   // bade reports mein browser hang na ho
  let capped = false;
  (function walk(nodes) {
    nodes.forEach(n => {
      if (flat.length >= MAX_RENDER) { capped = true; return; }
      flat.push(n);
      const hasKids = n.children && n.children.length;
      // closed by default - opens only when the user clicks it
      const isCollapsed = QuickReport.collapsed[n.path] !== false;
      const isLeaf = !hasKids;
      body += '<tr class="qr-row depth-' + Math.min(n.depth, 4) + (isLeaf ? ' qr-leaf' : ' qr-group') + '" data-path="' + escapeHtml(n.path) + '">' +
        '<td class="qr-cell" style="padding-left:' + (10 + n.depth * 22) + 'px">' +
          (hasKids
            ? '<span class="qr-toggle" data-toggle="' + escapeHtml(n.path) + '">' + (isCollapsed ? '\u25B8' : '\u25BE') + '</span>'
            : '<span class="qr-bullet"></span>') +
          escapeHtml(n.key) +
          (hasKids ? ' <span class="qr-total-tag">Total</span>' : '') +
        '</td>' +
        '<td class="num qr-val">' + fmtNum(n.value) + '</td>' +
        '<td class="num qr-share">' + (grand ? fmtNum(n.value / grand * 100, 1) + '%' : '') + '</td>' +
        '<td class="num qr-rows">' + n.rows.toLocaleString('en-IN') + '</td>' +
      '</tr>';
      if (hasKids && !isCollapsed) walk(n.children);
    });
  })(tree);

  QuickReport.lastRows = { flat, dims, label, grand };

  const colHead = dims.map(d => {
    const on = !!QuickReport.filters[d];
    return '<span class="qr-colchip' + (on ? ' on' : '') + '" data-col="' + escapeHtml(d) + '" title="' + escapeHtml(d) + ' par filter lagao">' +
      escapeHtml(d) + '<span class="qr-funnel">\u25BE</span></span>';
  }).join('<span class="qr-arrow">\u2794</span>');

  table.innerHTML =
    '<thead><tr>' +
      '<th class="qr-th-main">' + colHead + '</th>' +
      '<th class="num">' + escapeHtml(label) + '</th>' +
      '<th class="num">Share</th>' +
      '<th class="num">Rows</th>' +
    '</tr></thead><tbody>' + body + '</tbody>' +
    '<tfoot><tr><td>Grand Total</td><td class="num">' + fmtNum(grand) + '</td><td class="num">100%</td>' +
    '<td class="num">' + recs.length.toLocaleString('en-IN') + '</td></tr></tfoot>';

  table.querySelectorAll('.qr-colchip').forEach(chip => chip.addEventListener('click', e => {
    e.stopPropagation();
    openQuickColumnFilter(chip.dataset.col, chip);
  }));

  table.querySelectorAll('.qr-toggle').forEach(t => t.addEventListener('click', e => {
    e.stopPropagation();
    const p = t.dataset.toggle;
    QuickReport.collapsed[p] = (QuickReport.collapsed[p] === false);
    renderQuickReport();
  }));
  makeTableResizable(table);
  table.querySelectorAll('.qr-row').forEach(tr => tr.addEventListener('click', () => {
    const node = flat.find(n => n.path === tr.dataset.path);
    if (!node) return;
    const parts = node.path.split('|');
    const filters = [];
    for (let i = 0; i < parts.length && i < dims.length; i++) {
      if (DERIVED_DIMS.includes(dims[i])) continue;   // drill sirf asli columns par
      filters.push({ field: dims[i], value: parts[i] });
    }
    if (filters.length) openDrillPath(filters);
  }));

  document.getElementById('quick-count').textContent =
    flat.length.toLocaleString('en-IN') + ' rows shown \u00b7 ' + recs.length.toLocaleString('en-IN') + ' source rows' +
    (capped ? ' \u00b7 pehli ' + MAX_RENDER.toLocaleString('en-IN') + ' rows shown (use Export CSV for the full list)' : '');
}

function quickToGrid() {
  if (!QuickReport.lastRows || !QuickReport.lastRows.flat.length) return null;
  const { flat, dims, label } = QuickReport.lastRows;
  const headers = dims.concat([label, 'Rows']);
  const rows = flat.map(n => {
    const parts = n.path.split('|');
    const cells = dims.map((d, i) => (i < parts.length ? parts[i] : ''));
    return cells.concat([n.value, n.rows]);
  });
  return { headers, rows };
}

function exportQuickCSV() {
  const g = quickToGrid();
  if (!g) { toast('Please tick at least one column first.'); return; }
  downloadBlob(toCSV(g.headers, g.rows), 'quick-report.csv', 'text/csv');
}

function quickToSheet() {
  const g = quickToGrid();
  if (!g) { toast('Please tick at least one column first.'); return; }
  const sheetName = prompt('Which sheet should this be written to?', 'StockLedger Report');
  if (!sheetName) return;
  toast('Writing to sheet...');
  gsPost({ action: 'write', sheet: sheetName, mode: 'replace', values: [g.headers].concat(g.rows) })
    .then(res => toast('Done - "' + res.sheet + '" now has ' + res.rowsWritten + ' rows.'))
    .catch(err => toast('Write failed: ' + err.message));
}

/* ---------------------------------------------------------------
   8. REORDER PLANNER
   --------------------------------------------------------------- */
function initInsights() {
  if (!document.getElementById('insights-groupby')) return;   // tab removed
  document.getElementById('insights-groupby').addEventListener('change', renderInsights);
  document.getElementById('insights-target-days').addEventListener('input', debounce(renderInsights, 250));
  document.getElementById('insights-only-flagged').addEventListener('change', renderInsights);
  document.getElementById('insights-export').addEventListener('click', exportInsightsCSV);
}

function aggregateByDimension(records, dim) {
  const map = new Map();
  records.forEach(r => {
    const key = dimKey(r, dim);
    map.set(key, (map.get(key) || 0) + (recQty(r)));
  });
  return map;
}

let lastInsightsRows = null;

function renderInsights() {
  const wrap = document.getElementById('insights-table');
  if (!wrap) return;   // Reorder Planner tab was removed
  const kpiWrap = document.getElementById('insights-kpis');
  const note = document.getElementById('insights-missing-note');
  const dim = document.getElementById('insights-groupby').value;
  const tdEl = document.getElementById('insights-target-days');
  const targetDays = Math.max(1, parseInt(tdEl ? tdEl.value : (Prefs.targetDays || 30), 10) || 30);
  const onlyFlagged = document.getElementById('insights-only-flagged').checked;

  if (!App.datasets.length) {
    wrap.innerHTML = '<tr><td class="empty-hint">Load a Sales, Purchase or Stock file to see reorder suggestions.</td></tr>';
    kpiWrap.innerHTML = ''; note.style.display = 'none'; lastInsightsRows = null;
    return;
  }

  const missing = [];
  if (!datasetsOfType('sales').length) missing.push('Sales');
  if (!datasetsOfType('purchase').length) missing.push('Purchase');
  if (!datasetsOfType('stock').length) missing.push('Stock');
  if (missing.length) {
    note.style.display = '';
    note.textContent = 'No ' + missing.join(' / ') + ' data loaded yet — those columns will show as 0 until you add it on Import.';
  } else note.style.display = 'none';

  const A = buildAnalysis(dim, targetDays);
  let rows = A.rows.slice().sort((a, b) => b.suggested - a.suggested || b.sold - a.sold);
  if (onlyFlagged) rows = rows.filter(r => r.reorder);
  lastInsightsRows = { rows, dim, periodDays: A.days, targetDays };

  const totalSold = A.rows.reduce((s, r) => s + r.sold, 0);
  const totalStock = A.rows.reduce((s, r) => s + r.stock, 0);
  const totalPurchased = A.rows.reduce((s, r) => s + r.purchased, 0);
  const reorderCount = A.rows.filter(r => r.reorder).length;
  const reorderQty = A.rows.reduce((s, r) => s + (r.reorder ? r.suggested : 0), 0);

  const erI = effectiveRange(A.range);
  if (kpiWrap) kpiWrap.innerHTML = [
    ['Analysis window', erI.from ? erI.text : A.range.label, A.range.label + ' · ' + A.days + ' days'],
    ['Sold qty', fmtNum(totalSold), 'in this window'],
    ['Purchased qty', fmtNum(totalPurchased), 'in this window'],
    ['Stock on hand', fmtNum(totalStock), 'current snapshot'],
    ['Reorder candidates', reorderCount.toLocaleString('en-IN'), fmtNum(reorderQty) + ' pcs suggested']
  ].map(([label, value, sub]) => '<div class="kpi-card"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>').join('');

  const head = '<thead><tr><th>' + dim + '</th><th class="num">Sold</th><th class="num">Purchased</th><th class="num">Stock</th>' +
    '<th class="num">Avg/day</th><th class="num">Days cover</th><th>Last sold</th><th class="num">Suggested reorder</th></tr></thead>';
  const body = '<tbody>' + rows.slice(0, 1000).map(r =>
    '<tr class="drillable" data-key="' + escapeHtml(r.key) + '">' +
    '<td>' + escapeHtml(r.key) + ' <span class="drill-hint">▸</span>' + (r.reorder ? ' <span class="row-flag">●</span>' : '') + '</td>' +
    '<td class="num">' + fmtNum(r.sold) + '</td>' +
    '<td class="num">' + fmtNum(r.purchased) + '</td>' +
    '<td class="num">' + fmtNum(r.stock) + '</td>' +
    '<td class="num">' + fmtNum(r.avgDaily, 2) + '</td>' +
    '<td class="num">' + (r.daysCover === Infinity ? '∞' : fmtNum(r.daysCover, 1)) + '</td>' +
    '<td>' + (r.lastSale ? fmtDate(r.lastSale) : '—') + '</td>' +
    '<td class="num"><strong>' + fmtNum(r.suggested) + '</strong></td>' +
    '</tr>'
  ).join('') + '</tbody>';

  const foot = '<tfoot><tr><td>Total (' + rows.length.toLocaleString('en-IN') + ' rows)</td>' +
    '<td class="num">' + fmtNum(rows.reduce((s, r) => s + r.sold, 0)) + '</td>' +
    '<td class="num">' + fmtNum(rows.reduce((s, r) => s + r.purchased, 0)) + '</td>' +
    '<td class="num">' + fmtNum(rows.reduce((s, r) => s + r.stock, 0)) + '</td>' +
    '<td></td><td></td><td></td>' +
    '<td class="num"><strong>' + fmtNum(rows.reduce((s, r) => s + (r.reorder ? r.suggested : 0), 0)) + '</strong></td>' +
    '</tr></tfoot>';

  wrap.innerHTML = head + body + foot;
  makeTableResizable(wrap);
  wrap.querySelectorAll('tbody tr').forEach(tr => tr.addEventListener('click', () => {
    if (tr.dataset.key) openDrill(dim, tr.dataset.key);
  }));
}

function insightsToGrid() {
  if (!lastInsightsRows || !lastInsightsRows.rows.length) return null;
  const { rows, dim } = lastInsightsRows;
  const headers = [dim, 'Sold Qty', 'Purchased Qty', 'Stock Qty', 'Avg Daily Sale', 'Days Cover', 'Last Sold', 'Suggested Reorder', 'Reorder Candidate'];
  const data = rows.map(r => [r.key, r.sold, r.purchased, r.stock, Number(r.avgDaily.toFixed(2)),
    r.daysCover === Infinity ? '' : Number(r.daysCover.toFixed(1)),
    r.lastSale ? fmtDate(r.lastSale) : '', r.suggested, r.reorder ? 'Yes' : 'No']);
  return { headers, rows: data };
}

function exportInsightsCSV() {
  const grid = insightsToGrid();
  if (!grid) { toast('Nothing to export yet.'); return; }
  downloadBlob(toCSV(grid.headers, grid.rows), 'reorder-plan.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   8b. PERFORMANCE — best sellers, dead stock, ABC, ageing
   --------------------------------------------------------------- */
const PerfState = { view: 'all', search: '', sortKey: 'sold', sortDir: -1, open: {} };
let perfCharts = {};
let lastPerfRows = null;

function initPerformance() {
  document.getElementById('perf-groupby').addEventListener('change', () => { PerfState.open = {}; renderPerformance(); });
  document.getElementById('perf-search').addEventListener('input', debounce(e => { PerfState.search = e.target.value; renderPerformance(); }, 200));
  document.getElementById('perf-export').addEventListener('click', exportPerfCSV);
  document.querySelectorAll('#perf-view .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#perf-view .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      PerfState.view = btn.dataset.view;
      renderPerformance();
    });
  });
}

function renderPerformance() {
  const dim = document.getElementById('perf-groupby').value;
  const tableEl = document.getElementById('perf-table');
  const kpiWrap = document.getElementById('perf-kpis');   // removed from the UI
  const chartsWrap = document.getElementById('perf-charts');

  if (!App.datasets.length) {
    tableEl.innerHTML = '<tr><td class="empty-hint">Load your data on the Import tab first.</td></tr>';
    if (kpiWrap) kpiWrap.innerHTML = ''; chartsWrap.style.display = 'none'; lastPerfRows = null;
    return;
  }
  chartsWrap.style.display = '';

  const tdEl = document.getElementById('insights-target-days');
  const targetDays = Math.max(1, parseInt(tdEl ? tdEl.value : (Prefs.targetDays || 30), 10) || 30);
  const A = buildAnalysis(dim, targetDays);

  // KPIs
  const dead = A.rows.filter(r => r.status === 'Dead stock');
  const best = A.rows.filter(r => r.status === 'Best seller');
  const oos = A.rows.filter(r => r.status === 'Out of stock');
  const over = A.rows.filter(r => r.overstocked);
  const deadQty = dead.reduce((s, r) => s + r.stock, 0);
  const excessQty = over.reduce((s, r) => s + r.excessQty, 0);
  const totalStock = A.rows.reduce((s, r) => s + r.stock, 0);
  const overallST = (A.totalSold + totalStock) > 0 ? (A.totalSold / (A.totalSold + totalStock)) * 100 : 0;

  const erP = effectiveRange(A.range);
  if (kpiWrap) kpiWrap.innerHTML = [
    ['Window', erP.from ? erP.text : A.range.label, A.range.label + ' · ' + A.days + ' days'],
    ['Best sellers', best.length.toLocaleString('en-IN'), 'A-class, 80% of sales'],
    ['Dead / non-moving', dead.length.toLocaleString('en-IN'), fmtNum(deadQty) + ' pcs stuck'],
    ['Overstocked', over.length.toLocaleString('en-IN'), fmtNum(excessQty) + ' pcs excess'],
    ['Out of stock', oos.length.toLocaleString('en-IN'), 'sold but zero balance'],
    ['Overall sell-through', fmtNum(overallST, 1) + '%', 'sold ÷ (sold + stock)']
  ].map(([label, value, sub]) => '<div class="kpi-card"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>').join('');

  renderBoard('perf');

  // filter for table
  let rows = A.rows.slice();
  if (PerfState.view === 'best') rows = rows.filter(r => r.status === 'Best seller');
  else if (PerfState.view === 'dead') rows = rows.filter(r => r.status === 'Dead stock');
  else if (PerfState.view === 'slow') rows = rows.filter(r => r.status === 'Slow mover');
  else if (PerfState.view === 'oos') rows = rows.filter(r => r.status === 'Out of stock');
  else if (PerfState.view === 'over') rows = rows.filter(r => r.overstocked);
  if (PerfState.search) {
    const q = PerfState.search.toLowerCase();
    rows = rows.filter(r => r.key.toLowerCase().includes(q) ||
      Object.values(r.meta).some(v => String(v).toLowerCase().includes(q)));
  }

  const sk = PerfState.sortKey, sd = PerfState.sortDir;
  rows.sort((a, b) => {
    let va = a[sk], vb = b[sk];
    if (va instanceof Date) va = va.getTime();
    if (vb instanceof Date) vb = vb.getTime();
    if (va === null || va === undefined) va = -Infinity;
    if (vb === null || vb === undefined) vb = -Infinity;
    if (typeof va === 'string' || typeof vb === 'string') return String(va).localeCompare(String(vb)) * sd;
    return (va - vb) * sd;
  });
  lastPerfRows = { rows, dim };

  // OBS/CBS report load hui ho to opening balance aur movement bhi dikhao,
  // aur "Stock" ko साफ likho ki wo closing balance (CBS) hai.
  const hasOBS = A.rows.some(r => r.hasOpening);
  const stockLabel = hasOBS ? 'Closing (CBS)' : 'Stock';
  const cols = [
    ['key', dim, false], ['sold', 'Sold', true], ['purchased', 'Purchased', true]
  ]
  .concat(hasOBS ? [['opening', 'Opening (OBS)', true]] : [])
  .concat([['stock', stockLabel, true]])
  .concat(hasOBS ? [['movement', 'Moved (OBS\u2212CBS)', true]] : [])
  .concat([
    ['sellThrough', 'Sell-through', true],
    ['lastSale', 'Last sold', false],
    ['abc', 'ABC', false], ['status', 'Status', false]
  ]);
  PerfState.cols = cols;
  const head = '<thead><tr>' + cols.map(([k, label, isNum]) =>
    '<th data-key="' + k + '" class="' + (isNum ? 'num' : '') + '">' + label +
    (sk === k ? '<span class="sort-arrow">' + (sd === 1 ? '▲' : '▼') + '</span>' : '') + '</th>').join('') + '</tr></thead>';

  const bodyRows = [];
  rows.slice(0, Prefs.perfLimit || 800).forEach(r => {
    bodyRows.push(perfRowHtml(r, dim, 0, [{ field: dim, value: r.key }]));
    // Row khuli ho to uske andar ki details usi table mein, neeche.
    if (PerfState.open[r.key]) {
      bodyRows.push(perfChildRowsHtml([{ field: dim, value: r.key }], 1, cols.length));
    }
  });
  const body = '<tbody>' + bodyRows.join('') + '</tbody>';

  const foot = perfFootHtml(rows, cols);
  tableEl.innerHTML = head + body + foot;
  tableEl.querySelectorAll('thead th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.key;
    if (PerfState.sortKey === k) PerfState.sortDir *= -1;
    else { PerfState.sortKey = k; PerfState.sortDir = (k === 'key' || k === 'status' || k === 'abc') ? 1 : -1; }
    renderPerformance();
  }));
  wirePerfRowEvents(tableEl, dim);
  makeTableResizable(tableEl);

  renderPerfLegend(hasOBS);
  document.getElementById('perf-count').textContent =
    rows.length.toLocaleString('en-IN') + ' rows' + (rows.length > 800 ? ' (showing first 800 — export for full list)' : '');
}

/* ---- Inline expand: row ke neeche hi uski details, bina naya window khole ---- */

const PERF_CHAIN_DEFAULT = ['Style', 'Colour', 'Size', 'Brand', 'Supplier', 'Sub Section', 'Item Code'];
const PERF_CHAIN_CHOICES = ['Style', 'Colour', 'Size', 'Article No', 'Brand', 'Supplier',
                            'Sub Section', 'Section', 'Item Code'];

/** The expand order, set in Settings -> Behaviour (like the mind map levels). */
function perfChain() {
  const c = (Behaviour.perfChain && Behaviour.perfChain.length) ? Behaviour.perfChain : PERF_CHAIN_DEFAULT;
  return c.filter(function (v, i) { return v && c.indexOf(v) === i; });
}

/** Kis level par kaunsa dimension khulega. */
function perfChildDim(usedFields) {
  const have = new Set();
  App.datasets.forEach(d => d.fields.forEach(f => have.add(f)));
  return perfChain().find(d => have.has(d) && usedFields.indexOf(d) === -1) || null;
}

/** Metric cells — PerfState.cols ke hisaab se, taaki OBS/CBS columns
 *  hone par header aur rows hamesha match karein. */
/** Har column aur status ka matlab — taaki table dekhte hi samajh aaye. */
function renderPerfLegend(hasOBS) {
  const el = document.getElementById('perf-legend');
  if (!el) return;
  const open = el.dataset.open === '1';
  const metricRows = []
    .concat(hasOBS ? [
      ['Opening (OBS)', 'Stock at the start of the period - from the OBS Qty column'],
      ['Closing (CBS)', 'Stock left at the end of the period - from the CBS Qty column'],
      ['Moved (OBS\u2212CBS)', '\u2193 means stock went down (issued), \u2191 means it went up (received)']
    ] : [['Stock', 'How much stock is on hand right now']])
    .concat([
      ['Sold', 'Quantity sold in the selected window'],
      ['Sell-through', 'sold / (sold + stock) - how much of the stock moved'],
      ['Days cover', 'stock / average daily sale - how many days of cover you hold'],
      ['Days since', 'Days since the last sale'],
      ['Stock age', 'Quantity-weighted average age from the purchase bill date'],
      ['Excess', 'Stock above the target cover level'],
      ['ABC', 'A = top items making up 80% of sales, B = next 15%, C = the rest']
    ]);

  const statusRows = [
    ['Best seller', 'A-class - your strongest sellers'],
    ['Steady', 'B-class - selling steadily'],
    ['Slow mover', 'C-class - selling, but very slowly'],
    ['Dead stock', 'Stock on hand but zero sales in this window, or no sale for 90+ days'],
    ['Out of stock', 'Sold, but the balance is now zero - may need reordering'],
    ['Overstocked', 'Days cover is more than 3x the target'],
    ['No activity', 'No sales and no stock']
  ];

  el.innerHTML =
    '<button class="legend-toggle" id="perf-legend-toggle">' + (open ? '\u25BE' : '\u25B8') +
      ' What the columns and statuses mean</button>' +
    (open ? '<div class="legend-body">' +
      '<div class="legend-col"><h4>Columns</h4>' +
        metricRows.map(([k, v]) => '<div class="legend-item"><strong>' + k + '</strong><span>' + v + '</span></div>').join('') +
      '</div>' +
      '<div class="legend-col"><h4>Status</h4>' +
        statusRows.map(([k, v]) => '<div class="legend-item">' +
          '<span class="status-tag st-' + k.replace(/\s+/g, '-').toLowerCase() + '">' + k + '</span>' +
          '<span>' + v + '</span></div>').join('') +
      '</div></div>' : '');

  const t = document.getElementById('perf-legend-toggle');
  if (t) t.addEventListener('click', () => {
    el.dataset.open = open ? '0' : '1';
    renderPerfLegend(hasOBS);
  });
}

function perfMetricCells(r) {
  const cols = PerfState.cols || [];
  let out = '';
  cols.forEach(([k]) => {
    if (k === 'key') return;
    if (k === 'sold') out += '<td class="num">' + fmtNum(r.sold) + '</td>';
    else if (k === 'purchased') out += '<td class="num cat-purch">' + fmtNum(r.purchased || 0) + '</td>';
    else if (k === 'opening') out += '<td class="num obs-col">' + (r.hasOpening ? fmtNum(r.opening) : '\u2014') + '</td>';
    else if (k === 'stock') out += '<td class="num cbs-col">' + fmtNum(r.stock) + '</td>';
    else if (k === 'movement') out += '<td class="num ' + (r.movement > 0 ? 'mv-down' : (r.movement < 0 ? 'mv-up' : '')) + '">' +
        (r.movement === null || r.movement === undefined ? '\u2014'
          : (r.movement > 0 ? '\u2193 ' : (r.movement < 0 ? '\u2191 ' : '')) + fmtNum(Math.abs(r.movement))) + '</td>';
    else if (k === 'sellThrough') out += '<td class="num">' + fmtNum(r.sellThrough, 1) + '%</td>';
    else if (k === 'daysCover') out += '<td class="num">' + (r.daysCover === Infinity ? '\u221E' : fmtNum(r.daysCover, 0)) + '</td>';
    else if (k === 'lastSale') out += '<td>' + (r.lastSale ? fmtDate(r.lastSale) : '\u2014') + '</td>';
    else if (k === 'daysSinceLastSale') out += '<td class="num">' + (r.daysSinceLastSale === null || r.daysSinceLastSale === undefined ? '\u2014' : r.daysSinceLastSale) + '</td>';
    else if (k === 'stockAgeDays') out += '<td class="num">' + (r.stockAgeDays === null || r.stockAgeDays === undefined ? '\u2014' : r.stockAgeDays) + '</td>';
    else if (k === 'excessQty') out += '<td class="num">' + (r.excessQty ? fmtNum(r.excessQty) : '\u2014') + '</td>';
    else if (k === 'abc') out += '<td><span class="abc-tag abc-' + (r.abc || '\u2014') + '">' + (r.abc || '\u2014') + '</span></td>';
    else if (k === 'status') out += '<td><span class="status-tag st-' + String(r.status || '').replace(/\s+/g, '-').toLowerCase() + '">' + escapeHtml(r.status || '') + '</span></td>';
    else out += '<td></td>';
  });
  return out;
}

/** Open state for a row. Nothing opens on its own: a row is open only if the
 *  user clicked it. The "open the top item too" behaviour is optional and off
 *  by default (Settings -> 02 Performance). */
function perfIsOpen(id, isFirstChild) {
  const v = PerfState.open[id];
  if (v !== undefined) return v;
  return Behaviour.autoOpenFirst ? !!isFirstChild : false;
}

function perfRowHtml(r, dim, depth, path, isFirstChild) {
  const key = path.map(p => p.field + '=' + p.value).join('|');
  const open = perfIsOpen(depth === 0 ? r.key : key, isFirstChild);
  const hasKids = !!perfChildDim(path.map(p => p.field));
  return '<tr class="perf-row depth-' + depth + (open ? ' is-open' : '') + '" data-key="' + escapeHtml(r.key) + '" data-path="' + escapeHtml(key) + '">' +
    '<td class="perf-name" style="--indent:' + (10 + depth * 18) + 'px">' +
      (hasKids
        ? '<button class="perf-caret' + (open ? ' open' : '') + '" title="Show details inside this row">\u25B8</button>'
        : '<span class="perf-caret-spacer"></span>') +
      '<span class="perf-key" title="' + escapeHtml(Object.values(r.meta || {}).join(' \u00b7 ')) + '">' + escapeHtml(r.key) + '</span>' +
      '<button class="perf-popout" title="Open full details in a separate window">\u29C9</button>' +
    '</td>' +
    perfMetricCells(r) +
  '</tr>';
}

/** Ek khuli row ke andar ki rows (aur unke andar ki, recursively). */
function perfChildRowsHtml(path, depth, colCount) {
  const usedFields = path.map(p => p.field);
  const childDim = perfChildDim(usedFields);
  if (!childDim) return '';

  const range = periodRange();
  const match = rec => path.every(p => dimKey(rec, p.field) === p.value);
  const sales = salesRecords().filter(r => inPeriod(r, range)).filter(match);
  const stock = stockRecords().filter(match);
  const days = periodDayCount(range, sales);
  const anchor = dataAnchorDate();

  const map = new Map();
  const slot = k => {
    let x = map.get(k);
    if (!x) { x = { key: k, sold: 0, purchased: 0, stock: 0, opening: 0, hasOpening: false, lastSale: null, meta: {} }; map.set(k, x); }
    return x;
  };
  purchaseRecords().filter(r => inPeriod(r, range)).filter(match).forEach(r => {
    slot(dimKey(r, childDim)).purchased += recQty(r);
  });
  sales.forEach(r => {
    const x = slot(dimKey(r, childDim));
    x.sold += recQty(r);
    if (r.Date && (!x.lastSale || r.Date > x.lastSale)) x.lastSale = r.Date;
  });
  stock.forEach(r => {
    const x = slot(dimKey(r, childDim));
    x.stock += recQty(r);
    const ob = recOpeningQty(r);
    if (ob !== null) { x.opening = (x.opening || 0) + ob; x.hasOpening = true; }
  });

  let kids = [...map.values()].map(x => {
    const opening = x.sold + x.stock;
    const avgDaily = x.sold / days;
    return Object.assign(x, {
      movement: x.hasOpening ? (x.opening - x.stock) : null,
      sellThrough: opening > 0 ? (x.sold / opening) * 100 : 0,
      daysCover: avgDaily > 0 ? x.stock / avgDaily : (x.stock > 0 ? Infinity : 0),
      daysSinceLastSale: x.lastSale ? Math.round((anchor - x.lastSale) / 86400000) : null,
      stockAgeDays: null, excessQty: 0, abc: '\u2014',
      status: x.sold === 0 && x.stock > 0 ? 'Dead stock' : (x.stock === 0 && x.sold > 0 ? 'Out of stock' : (x.sold > 0 ? 'Moving' : 'No activity'))
    });
  }).sort((a, b) => b.sold - a.sold);

  const LIMIT = 25;
  const shown = kids.slice(0, LIMIT);
  if (!shown.length) {
    return '<tr class="perf-child-note"><td colspan="' + colCount + '">No ' + escapeHtml(childDim) + ' data found here.</td></tr>';
  }

  // No section heading row — child rows follow the parent directly.
  let html = '';

  shown.forEach((k, i) => {
    const childPath = path.concat([{ field: childDim, value: k.key }]);
    const key = childPath.map(p => p.field + '=' + p.value).join('|');
    const isFirst = i === 0;
    html += perfRowHtml(k, childDim, depth, childPath, isFirst);
    if (perfIsOpen(key, isFirst)) html += perfChildRowsHtml(childPath, depth + 1, colCount);
  });
  return html;
}

/** Marks the first child of a path as open (one level), so expanding a row
 *  immediately reveals the leading item's detail instead of an empty list. */
function autoOpenFirstChild(pathStr) {
  const path = String(pathStr).split('|').filter(Boolean).map(function (seg) {
    const i = seg.indexOf('=');
    return { field: seg.slice(0, i), value: seg.slice(i + 1) };
  });
  if (!path.length) return;
  const childDim = perfChildDim(path.map(function (p) { return p.field; }));
  if (!childDim) return;

  const range = periodRange();
  const match = function (rec) { return path.every(function (p) { return dimKey(rec, p.field) === p.value; }); };
  const sales = salesRecords().filter(function (r) { return inPeriod(r, range); }).filter(match);
  const stock = stockRecords().filter(match);

  const totals = new Map();
  sales.forEach(function (r) { const k = dimKey(r, childDim); totals.set(k, (totals.get(k) || 0) + recQty(r)); });
  if (!totals.size) stock.forEach(function (r) { const k = dimKey(r, childDim); totals.set(k, (totals.get(k) || 0) + recQty(r)); });
  if (!totals.size) return;

  const first = [...totals.entries()].sort(function (a, b) { return b[1] - a[1]; })[0][0];
  PerfState.open[pathStr + '|' + childDim + '=' + first] = true;
}

function wirePerfRowEvents(tableEl, dim) {
  tableEl.querySelectorAll('.perf-caret').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const tr = btn.closest('tr');
      const depth = parseInt((tr.className.match(/depth-(\d+)/) || [0, 0])[1], 10);
      const id = depth === 0 ? tr.dataset.key : tr.dataset.path;
      const opening = !PerfState.open[id];
      PerfState.open[id] = opening;
      // When a row opens, show the top child's details straight away; the rest
      // stay closed until clicked.
      if (opening && Behaviour.autoOpenFirst) autoOpenFirstChild(tr.dataset.path || (dim + '=' + tr.dataset.key));
      renderPerformance();
    });
  });
  tableEl.querySelectorAll('.perf-popout').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const tr = btn.closest('tr');
      const path = (tr.dataset.path || '').split('|').filter(Boolean).map(seg => {
        const i = seg.indexOf('=');
        return { field: seg.slice(0, i), value: seg.slice(i + 1) };
      });
      if (path.length) openDrillPath(path);
      else openDrill(dim, tr.dataset.key);
    });
  });
  // poori row par click = expand/collapse (jahan possible ho)
  tableEl.querySelectorAll('tr.perf-row').forEach(tr => {
    tr.addEventListener('click', () => {
      if (!Behaviour.inlineExpand) { tr.querySelector('.perf-popout').click(); return; }
      const caret = tr.querySelector('.perf-caret');
      if (caret) caret.click();
    });
  });
}

/** Column-wise totals footer \u2014 sums over ALL filtered rows, not just the 800 shown. */
function perfFootHtml(rows, cols) {
  let out = '<tfoot><tr>';
  cols.forEach(([k], i) => {
    if (i === 0) { out += '<td>Total (' + rows.length.toLocaleString('en-IN') + ' rows)</td>'; return; }
    if (k === 'sold') out += '<td class="num">' + fmtNum(rows.reduce((s, r) => s + r.sold, 0)) + '</td>';
    else if (k === 'purchased') out += '<td class="num">' + fmtNum(rows.reduce((s, r) => s + (r.purchased || 0), 0)) + '</td>';
    else if (k === 'opening') out += '<td class="num">' + fmtNum(rows.reduce((s, r) => s + (r.hasOpening ? r.opening : 0), 0)) + '</td>';
    else if (k === 'stock') out += '<td class="num">' + fmtNum(rows.reduce((s, r) => s + r.stock, 0)) + '</td>';
    else if (k === 'movement') {
      const m = rows.reduce((s, r) => s + (r.movement || 0), 0);
      out += '<td class="num ' + (m > 0 ? 'mv-down' : (m < 0 ? 'mv-up' : '')) + '">' + (m ? fmtNum(Math.abs(m)) : '\u2014') + '</td>';
    }
    else if (k === 'sellThrough') {
      const sold = rows.reduce((s, r) => s + r.sold, 0), stock = rows.reduce((s, r) => s + r.stock, 0);
      out += '<td class="num">' + (sold + stock > 0 ? fmtNum(sold / (sold + stock) * 100, 1) + '%' : '\u2014') + '</td>';
    }
    else out += '<td></td>';
  });
  return out + '</tr></tfoot>';
}

function destroyPerfChart(id) { if (perfCharts[id]) { perfCharts[id].destroy(); delete perfCharts[id]; } }

function renderParetoChart(rows) {
  if (!document.getElementById('chart-pareto')) return;   // replaced by the chart board
  destroyPerfChart('chart-pareto');
  const top = rows.filter(r => r.sold > 0).slice(0, 15);
  if (!top.length) return;
  const ctx = document.getElementById('chart-pareto').getContext('2d');
  const paretoKeys = top.map(r => r.key);
  const paretoDim = document.getElementById('perf-groupby').value;
  perfCharts['chart-pareto'] = makeChart(ctx, {
    data: {
      labels: top.map(r => r.key.length > 22 ? r.key.slice(0, 22) + '…' : r.key),
      datasets: [
        { type: 'bar', label: 'Qty sold', data: top.map(r => r.sold), backgroundColor: CHART_COLORS[0], yAxisID: 'y' },
        { type: 'line', label: 'Cumulative %', data: top.map(r => r.cumPct), borderColor: CHART_COLORS[2], backgroundColor: 'transparent', yAxisID: 'y1', tension: .3, pointRadius: 2 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      onClick: (evt, els) => { if (els && els.length) openDrill(paretoDim, paretoKeys[els[0].index]); },
      plugins: { tooltip: { callbacks: { title: items => paretoKeys[items[0].dataIndex] + '  (click for details)' } } },
      scales: {
        x: { ticks: { font: { size: 9.5 }, maxRotation: 60, minRotation: 40 } },
        y: { beginAtZero: true, position: 'left', title: { display: true, text: 'Qty' } },
        y1: { beginAtZero: true, max: 100, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: '%' } }
      }
    }
  });
}

function renderStatusChart(rows) {
  if (!document.getElementById('chart-pareto')) return;   // replaced by the chart board
  destroyPerfChart('chart-status');
  const counts = {};
  rows.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
  const labels = Object.keys(counts);
  if (!labels.length) return;
  const ctx = document.getElementById('chart-status').getContext('2d');
  perfCharts['chart-status'] = makeChart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data: labels.map(l => counts[l]), backgroundColor: labels.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } } } }
  });
}

function renderAgeingChart(rows) {
  if (!document.getElementById('chart-pareto')) return;   // replaced by the chart board
  destroyPerfChart('chart-ageing');
  const buckets = [['0–30 d', 0, 30], ['31–90 d', 31, 90], ['91–180 d', 91, 180], ['181–365 d', 181, 365], ['1 yr+', 366, Infinity]];
  const data = buckets.map(([, lo, hi]) =>
    rows.reduce((s, r) => s + ((r.stockAgeDays !== null && r.stockAgeDays >= lo && r.stockAgeDays <= hi) ? r.stock : 0), 0));
  if (!data.some(d => d > 0)) return;
  const ctx = document.getElementById('chart-ageing').getContext('2d');
  perfCharts['chart-ageing'] = makeChart(ctx, {
    type: 'bar',
    data: { labels: buckets.map(b => b[0]), datasets: [{ label: 'Stock qty', data, backgroundColor: buckets.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) }] },
    options: Object.assign(chartOptions(), { plugins: { legend: { display: false } } })
  });
}

function renderBottomChart(rows) {
  if (!document.getElementById('chart-pareto')) return;   // replaced by the chart board
  destroyPerfChart('chart-bottom');
  const stuck = rows.filter(r => r.stock > 0 && r.sold === 0).sort((a, b) => b.stock - a.stock).slice(0, 10);
  if (!stuck.length) return;
  const ctx = document.getElementById('chart-bottom').getContext('2d');
  const stuckKeys = stuck.map(r => r.key);
  const stuckDim = document.getElementById('perf-groupby').value;
  perfCharts['chart-bottom'] = makeChart(ctx, {
    type: 'bar',
    data: {
      labels: stuck.map(r => r.key.length > 24 ? r.key.slice(0, 24) + '…' : r.key),
      datasets: [{ label: 'Stock lying unsold', data: stuck.map(r => r.stock), backgroundColor: CHART_COLORS[5] }]
    },
    options: Object.assign(chartOptions(), {
      indexAxis: 'y',
      plugins: { legend: { display: false },
        tooltip: { callbacks: { title: items => stuckKeys[items[0].dataIndex] + '  (click for details)' } } },
      onClick: (evt, els) => { if (els && els.length) openDrill(stuckDim, stuckKeys[els[0].index]); }
    })
  });
}

function perfToGrid() {
  if (!lastPerfRows || !lastPerfRows.rows.length) return null;
  const { rows, dim } = lastPerfRows;
  const headers = [dim, 'Brand', 'Section', 'Supplier', 'Sold Qty', 'Purchased Qty', 'Stock Qty',
    'Sell-through %', 'Days Cover', 'Last Sold', 'Days Since Last Sale', 'Stock Age (days)',
    'Excess Qty', 'ABC', 'Status'];
  const data = rows.map(r => [r.key, r.meta.Brand || '', r.meta.Section || '', r.meta.Supplier || '',
    r.sold, r.purchased, r.stock, Number(r.sellThrough.toFixed(1)),
    r.daysCover === Infinity ? '' : Math.round(r.daysCover),
    r.lastSale ? fmtDate(r.lastSale) : '', r.daysSinceLastSale === null ? '' : r.daysSinceLastSale,
    r.stockAgeDays === null ? '' : r.stockAgeDays, r.excessQty || 0, r.abc, r.status]);
  return { headers, rows: data };
}

function exportPerfCSV() {
  const grid = perfToGrid();
  if (!grid) { toast('Nothing to export yet.'); return; }
  downloadBlob(toCSV(grid.headers, grid.rows), 'product-performance.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   8c. DRILL-DOWN — kisi bhi cheez par click karke depth mein jaao
   --------------------------------------------------------------- */
const Drill = { filters: [], dim: 'Colour', open: false, showRaw: false, sortKey: 'sold', sortDir: -1 };
let drillChart = null;

const DRILL_DIMS = ['Article No', 'Style', 'Colour', 'Size', 'Brand', 'Sub Section',
                    'Section', 'Supplier', 'Item Code', 'Month'];

function availableDrillDims() {
  const have = new Set();
  App.datasets.forEach(d => d.fields.forEach(f => have.add(f)));
  // relationships se jude fields bhi available maante hain
  if (App.relationships.some(r => r.enabled)) App.datasets.forEach(d => d.fields.forEach(f => have.add(f)));
  return DRILL_DIMS.filter(d => d === 'Month' || have.has(d));
}

function drillMatches(rec) {
  for (const f of Drill.filters) {
    let v;
    if (f.field === 'Month') v = rec.Date ? dateKeyForGrain(rec.Date, 'month') : '(blank)';
    else v = dimKey(rec, f.field);
    if (v !== f.value) return false;
  }
  return true;
}

function drillRecords(type) {
  const range = periodRange();
  let recs;
  if (type === 'sales') recs = salesRecords().filter(r => inPeriod(r, range));
  else if (type === 'purchase') recs = purchaseRecords().filter(r => inPeriod(r, range));
  else recs = stockRecords();
  return recs.filter(drillMatches);
}

function openDrill(field, value) {
  Drill.filters = [{ field, value }];
  Drill.open = true;
  Drill.showRaw = false;
  // default breakdown: pehla available dim jo filter mein nahi hai
  const dims = availableDrillDims().filter(d => d !== field);
  Drill.dim = dims.includes('Colour') ? 'Colour' : (dims[0] || 'Style');
  document.getElementById('drill-overlay').style.display = 'flex';
  modalOpen('drill-overlay', closeDrill);
  renderDrill();
}

function toISODate(d) {
  if (!(d instanceof Date)) return null;
  return d.getUTCFullYear() + '-' +
         String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
         String(d.getUTCDate()).padStart(2, '0');
}

/** Snapshot ki window ko baaki app ke period filter par bhi laga deta hai,
 *  taaki "This week" ka node click karne par drill bhi wahi week dikhaye. */
function syncPeriodToRange(range) {
  if (!range || !range.from || !range.to) return;
  App.period = { mode: 'custom', from: toISODate(range.from), to: toISODate(range.to) };
  document.querySelectorAll('.period-select').forEach(s => { s.value = 'custom'; });
  document.querySelectorAll('.period-custom-range').forEach(el => { el.style.display = ''; });
  document.querySelectorAll('.period-from').forEach(el => { el.value = App.period.from; });
  document.querySelectorAll('.period-to').forEach(el => { el.value = App.period.to; });
  renderInsights(); renderPerformance(); renderDashboard();
}

/** Mind map se aaye poore raste (Sub Section > Style > Colour) ke saath drill kholta hai. */
function openDrillPath(filters, range) {
  if (!filters || !filters.length) return;
  if (range) syncPeriodToRange(range);
  Drill.filters = filters.slice();
  Drill.open = true;
  Drill.showRaw = false;
  const used = filters.map(f => f.field);
  const dims = availableDrillDims().filter(d => used.indexOf(d) === -1);
  Drill.dim = dims[0] || 'Style';
  document.getElementById('drill-overlay').style.display = 'flex';
  modalOpen('drill-overlay', closeDrill);
  renderDrill();
}

function pushDrill(field, value) {
  if (Drill.filters.some(f => f.field === field)) return;
  Drill.filters.push({ field, value });
  const dims = availableDrillDims().filter(d => !Drill.filters.some(f => f.field === d));
  Drill.dim = dims[0] || Drill.dim;
  renderDrill();
}

function popDrillTo(idx) {
  Drill.filters = Drill.filters.slice(0, idx + 1);
  renderDrill();
}

function closeDrill() {
  Drill.open = false;
  modalClose('drill-overlay');
  document.getElementById('drill-overlay').style.display = 'none';
  if (drillChart) { drillChart.destroy(); drillChart = null; }
}

function initDrill() {
  document.getElementById('drill-close').addEventListener('click', closeDrill);
  document.getElementById('drill-overlay').addEventListener('click', e => {
    if (e.target.id === 'drill-overlay' && modalIsTop('drill-overlay')) closeDrill();
  });
  document.getElementById('drill-dim').addEventListener('change', e => { Drill.dim = e.target.value; renderDrill(); });
  document.getElementById('drill-raw-toggle').addEventListener('click', () => {
    Drill.showRaw = !Drill.showRaw; renderDrill();
  });
  document.getElementById('drill-export').addEventListener('click', exportDrillCSV);
}

let lastDrillRows = null;

function renderDrill() {
  const sales = drillRecords('sales');
  const purch = drillRecords('purchase');
  const stock = drillRecords('stock');
  const range = periodRange();
  const days = periodDayCount(range, sales.length ? sales : purch);
  const anchor = dataAnchorDate();

  const sumQ = rs => rs.reduce((s, r) => s + (recQty(r)), 0);
  const sold = sumQ(sales), purchased = sumQ(purch), inStock = sumQ(stock);
  const sellThrough = (sold + inStock) > 0 ? (sold / (sold + inStock)) * 100 : 0;
  const avgDaily = sold / days;
  const daysCover = avgDaily > 0 ? inStock / avgDaily : (inStock > 0 ? Infinity : 0);
  const dates = sales.map(r => r.Date).filter(Boolean);
  const lastSale = dates.length ? new Date(minMaxTime(dates).max) : null;

  // title + breadcrumb
  const last = Drill.filters[Drill.filters.length - 1];
  document.getElementById('drill-title').textContent = last ? last.value : 'Details';
  const erDr = effectiveRange(range);
  document.getElementById('drill-subtitle').textContent =
    Drill.filters.map(f => f.field + ': ' + f.value).join('  ·  ') +
    (erDr.from ? '      |      Data: ' + erDr.text : '');

  document.getElementById('drill-breadcrumb').innerHTML =
    Drill.filters.map((f, i) =>
      '<button class="crumb" data-idx="' + i + '">' + escapeHtml(f.field) + ': <strong>' + escapeHtml(f.value) + '</strong>' +
      (i < Drill.filters.length - 1 ? '' : '') + '</button>'
    ).join('<span class="crumb-sep">›</span>');
  document.querySelectorAll('#drill-breadcrumb .crumb').forEach(b =>
    b.addEventListener('click', () => popDrillTo(parseInt(b.dataset.idx, 10))));

  // KPIs
  document.getElementById('drill-kpis').innerHTML = [
    ['Sold', fmtNum(sold), sales.length.toLocaleString('en-IN') + ' bill lines'],
    ['Purchased', fmtNum(purchased), purch.length.toLocaleString('en-IN') + ' lines'],
    ['Stock', fmtNum(inStock), stock.length.toLocaleString('en-IN') + ' stock rows'],
    ['Sell-through', fmtNum(sellThrough, 1) + '%', 'sold ÷ (sold + stock)'],
    ['Days cover', daysCover === Infinity ? '∞' : fmtNum(daysCover, 0), 'at ' + fmtNum(avgDaily, 2) + '/day'],
    ['Last sold', lastSale ? fmtDate(lastSale) : '—', lastSale ? Math.round((anchor - lastSale) / 86400000) + ' days ago' : 'no sale in window']
  ].map(([l, v, s]) => '<div class="kpi-card small"><div class="kpi-label">' + l + '</div><div class="kpi-value">' + v + '</div><div class="kpi-sub">' + s + '</div></div>').join('');

  // breakdown dim selector
  const dims = availableDrillDims().filter(d => !Drill.filters.some(f => f.field === d));
  const dimSel = document.getElementById('drill-dim');
  dimSel.innerHTML = dims.map(d => '<option value="' + d + '"' + (d === Drill.dim ? ' selected' : '') + '>' + d + '</option>').join('');
  if (!dims.includes(Drill.dim)) Drill.dim = dims[0];

  renderDrillTrend(sales);
  renderDrillBreakdown(sales, purch, stock, days);

  document.getElementById('drill-raw-toggle').textContent = Drill.showRaw ? '▴ Hide raw rows' : '▾ Show raw rows';
  const rawWrap = document.getElementById('drill-raw');
  if (Drill.showRaw) { rawWrap.style.display = ''; renderDrillRaw(sales); }
  else rawWrap.style.display = 'none';
}

function renderDrillTrend(sales) {
  if (drillChart) { drillChart.destroy(); drillChart = null; }
  const m = new Map();
  sales.forEach(r => {
    if (!r.Date) return;
    const k = dateKeyForGrain(r.Date, 'month');
    m.set(k, (m.get(k) || 0) + (recQty(r)));
  });
  const keys = [...m.keys()].sort((a, b) => grainSort(a, 'month') - grainSort(b, 'month'));
  const el = document.getElementById('drill-trend');
  if (!keys.length) { el.parentElement.style.display = 'none'; return; }
  el.parentElement.style.display = '';
  drillChart = makeChart(el.getContext('2d'), {
    type: 'bar',
    data: { labels: keys, datasets: [{ label: 'Qty sold', data: keys.map(k => m.get(k)), backgroundColor: CHART_COLORS[0] }] },
    options: Object.assign(chartOptions(), { plugins: { legend: { display: false } } })
  });
}

function renderDrillBreakdown(sales, purch, stock, days) {
  const dim = Drill.dim;
  const map = new Map();
  function slot(k) {
    let s = map.get(k);
    if (!s) { s = { key: k, sold: 0, purchased: 0, stock: 0, lastSale: null }; map.set(k, s); }
    return s;
  }
  const keyOf = r => dim === 'Month'
    ? (r.Date ? dateKeyForGrain(r.Date, 'month') : '(blank)')
    : dimKey(r, dim);

  sales.forEach(r => {
    const s = slot(keyOf(r));
    s.sold += recQty(r);
    if (r.Date && (!s.lastSale || r.Date > s.lastSale)) s.lastSale = r.Date;
  });
  purch.forEach(r => { slot(keyOf(r)).purchased += recQty(r); });
  stock.forEach(r => { slot(keyOf(r)).stock += recQty(r); });

  let rows = [...map.values()].map(s => {
    const opening = s.sold + s.stock;
    return Object.assign(s, {
      sellThrough: opening > 0 ? (s.sold / opening) * 100 : 0,
      daysCover: (s.sold / days) > 0 ? s.stock / (s.sold / days) : (s.stock > 0 ? Infinity : 0)
    });
  });

  const sk = Drill.sortKey, sd = Drill.sortDir;
  rows.sort((a, b) => {
    let va = a[sk], vb = b[sk];
    if (va instanceof Date) va = va.getTime();
    if (vb instanceof Date) vb = vb.getTime();
    if (va === null || va === undefined) va = -Infinity;
    if (vb === null || vb === undefined) vb = -Infinity;
    if (typeof va === 'string' || typeof vb === 'string') return String(va).localeCompare(String(vb)) * sd;
    return (va - vb) * sd;
  });
  lastDrillRows = { rows, dim };

  const totalSold = rows.reduce((s, r) => s + r.sold, 0);
  const cols = [['key', dim, false], ['sold', 'Sold', true], ['share', 'Share', true],
                ['stock', 'Stock', true], ['sellThrough', 'Sell-thru', true],
                ['daysCover', 'Cover', true], ['lastSale', 'Last sold', false]];

  const canDrillDeeper = availableDrillDims().filter(d => !Drill.filters.some(f => f.field === d)).length > 1;

  document.getElementById('drill-breakdown').innerHTML =
    '<thead><tr>' + cols.map(([k, l, n]) =>
      '<th data-key="' + k + '" class="' + (n ? 'num' : '') + '">' + l +
      (sk === k ? '<span class="sort-arrow">' + (sd === 1 ? '▲' : '▼') + '</span>' : '') + '</th>').join('') +
    '</tr></thead><tbody>' +
    rows.slice(0, 300).map(r =>
      '<tr class="' + (canDrillDeeper ? 'drillable' : '') + '" data-value="' + escapeHtml(r.key) + '">' +
        '<td>' + escapeHtml(r.key) + (canDrillDeeper ? ' <span class="drill-hint">▸</span>' : '') + '</td>' +
        '<td class="num">' + fmtNum(r.sold) + '</td>' +
        '<td class="num">' + (totalSold > 0 ? fmtNum(r.sold / totalSold * 100, 1) + '%' : '—') + '</td>' +
        '<td class="num">' + fmtNum(r.stock) + '</td>' +
        '<td class="num">' + fmtNum(r.sellThrough, 1) + '%</td>' +
        '<td class="num">' + (r.daysCover === Infinity ? '∞' : fmtNum(r.daysCover, 0)) + '</td>' +
        '<td>' + (r.lastSale ? fmtDate(r.lastSale) : '—') + '</td>' +
      '</tr>').join('') + '</tbody>' +
    // column-wise totals, matching the other tables
    (function () {
      const tSold = rows.reduce(function (a, x) { return a + x.sold; }, 0);
      const tStock = rows.reduce(function (a, x) { return a + x.stock; }, 0);
      const tST = (tSold + tStock) > 0 ? (tSold / (tSold + tStock)) * 100 : 0;
      return '<tfoot><tr>' +
        '<td>Total · ' + rows.length.toLocaleString('en-IN') + ' ' + escapeHtml(dim) + '</td>' +
        '<td class="num">' + fmtNum(tSold) + '</td>' +
        '<td class="num">100%</td>' +
        '<td class="num">' + fmtNum(tStock) + '</td>' +
        '<td class="num">' + fmtNum(tST, 1) + '%</td>' +
        '<td class="num"></td><td></td></tr></tfoot>';
    })();

  makeTableResizable(document.getElementById('drill-breakdown'));
  document.querySelectorAll('#drill-breakdown thead th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.key;
    if (Drill.sortKey === k) Drill.sortDir *= -1;
    else { Drill.sortKey = k; Drill.sortDir = (k === 'key') ? 1 : -1; }
    renderDrill();
  }));
  if (canDrillDeeper) {
    document.querySelectorAll('#drill-breakdown tbody tr.drillable').forEach(tr =>
      tr.addEventListener('click', () => pushDrill(Drill.dim, tr.dataset.value)));
  }
  document.getElementById('drill-breakdown-count').textContent =
    rows.length.toLocaleString('en-IN') + ' ' + dim + ' values' + (rows.length > 300 ? ' (top 300 shown)' : '');

  // Agar zyadatar sale rows mein ye column khaali hai to breakdown bharosemand
  // nahi hai — user ko saaf bata dete hain.
  const blankRow = rows.find(r => r.key === '(blank)');
  const noteEl = document.getElementById('drill-blank-note');
  if (blankRow && totalSold > 0 && (blankRow.sold / totalSold) >= 0.3) {
    noteEl.style.display = '';
    noteEl.textContent = 'Note: ' + fmtNum(blankRow.sold / totalSold * 100, 0) +
      '% of the sale rows in this selection have "' + dim + '" empty, so this breakdown is incomplete. ' +
      'Filling in ' + dim + ' in your ERP will make this analysis work.';
  } else {
    noteEl.style.display = 'none';
  }
}

function renderDrillRaw(sales) {
  const fields = ['Date', 'Item Code', 'Article No', 'Style', 'Colour', 'Size', 'Brand', 'Supplier', 'Quantity'];
  const have = fields.filter(f => App.datasets.some(d => d.fields.includes(f)));
  const rows = sales.slice(0, 200);
  document.getElementById('drill-raw-table').innerHTML =
    '<thead><tr>' + have.map(f => '<th class="' + (FIELD_KIND[f] === 'number' ? 'num' : '') + '">' + f + '</th>').join('') + '</tr></thead>' +
    '<tbody>' + rows.map(r => '<tr>' + have.map(f => {
      let v = r[f];
      if (v === null || v === undefined || v === '') v = resolveField(r, f);
      if (v instanceof Date) v = fmtDate(v);
      else if (typeof v === 'number' && FIELD_KIND[f] === 'number') v = fmtNum(v);
      return '<td class="' + (FIELD_KIND[f] === 'number' ? 'num' : '') + '">' + escapeHtml(v == null ? '' : v) + '</td>';
    }).join('') + '</tr>').join('') + '</tbody>';
  document.getElementById('drill-raw-count').textContent =
    sales.length.toLocaleString('en-IN') + ' sale rows' + (sales.length > 200 ? ' (first 200 shown)' : '');
}

function exportDrillCSV() {
  if (!lastDrillRows || !lastDrillRows.rows.length) { toast('There is nothing to export yet.'); return; }
  const { rows, dim } = lastDrillRows;
  const headers = [dim, 'Sold Qty', 'Purchased Qty', 'Stock Qty', 'Sell-through %', 'Days Cover', 'Last Sold'];
  const data = rows.map(r => [r.key, r.sold, r.purchased, r.stock, Number(r.sellThrough.toFixed(1)),
    r.daysCover === Infinity ? '' : Math.round(r.daysCover), r.lastSale ? fmtDate(r.lastSale) : '']);
  const ctx = Drill.filters.map(f => f.field + '-' + f.value).join('_').replace(/[^\w-]+/g, '');
  downloadBlob(toCSV(headers, data), 'drill-' + ctx.slice(0, 60) + '.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   8d. TOP ITEMS SNAPSHOT — mind map + top lists, configurable
   --------------------------------------------------------------- */
const SNAPSHOT_ALL_DIMS = ['Sub Section', 'Style', 'Section', 'Colour', 'Brand', 'Supplier', 'Size', 'Article No'];
const SNAPSHOT_DEFAULT_CONFIG = {
  levels: ['Sub Section', 'Style', 'Colour'],          // mind map ki hierarchy (upar se neeche)
  dims: ['Sub Section', 'Style', 'Colour', 'Supplier'], // purana "Top Lists" view
  mapStyle: 'tree-h',
  autoShow: true,
  topN: 5
};

const Snapshot = {
  config: Object.assign({}, SNAPSHOT_DEFAULT_CONFIG),
  periodMode: 'thisweek',   // thisweek | lastweek | thismonth | custom
  from: null, to: null,
  shownThisSession: false,
  view: 'map'               // 'map' | 'cards' | 'settings'
};

// Mind map ka apna zoom/pan/expand state
const SnapMap = { zoom: 1, panX: 40, panY: 0, expanded: {}, drag: null, nodes: [], w: 0, h: 0 };

function loadSnapshotConfig() {
  try {
    const raw = Store.get('sl_snapshot_config');
    if (raw) {
      const parsed = JSON.parse(raw);
      // Purani setting mein sirf "dims" tha — usse hierarchy bana lete hain.
      if (!parsed.levels && parsed.dims) parsed.levels = parsed.dims.slice(0, 3);
      Snapshot.config = Object.assign({}, SNAPSHOT_DEFAULT_CONFIG, parsed);
    }
  } catch (e) {}
}

function saveSnapshotConfigLocal() {
  Store.set('sl_snapshot_config', JSON.stringify(Snapshot.config));
}

/** Google Sheet se connect hote hi purani settings (agar save ki thi) khinch leta hai. */
function pullSnapshotConfigFromSheet() {
  if (!GS.url) return;
  gsGet({ action: 'data', sheet: 'StockLedger Settings', offset: 0, limit: 20 }).then(res => {
    const row = (res.rows || []).find(r => r[0] === 'snapshot_config');
    if (row && row[1]) {
      try {
        const parsed = JSON.parse(row[1]);
        if (!parsed.levels && parsed.dims) parsed.levels = parsed.dims.slice(0, 3);
        Snapshot.config = Object.assign({}, SNAPSHOT_DEFAULT_CONFIG, parsed);
        saveSnapshotConfigLocal();
        if (document.getElementById('snapshot-overlay').style.display !== 'none') renderSnapshot();
      } catch (e) {}
    }
  }).catch(() => { /* sheet abhi nahi bani — default settings chalengi */ });
}

/** Settings ko local aur (connected ho to) Google Sheet dono jagah save karta hai. */
function pushSnapshotConfig() {
  saveSnapshotConfigLocal();
  if (GS.url && GS.meta && GS.meta.canWrite) {
    gsPost({ action: 'write', sheet: 'StockLedger Settings', mode: 'replace',
      values: [['key', 'value'], ['snapshot_config', JSON.stringify(Snapshot.config)]] })
      .then(() => toast('Settings saved, including to your Google Sheet.'))
      .catch(() => toast('Settings saved on this device (could not save to the Sheet).'));
  } else {
    toast('Settings saved on this device.');
  }
}

function mondayOfWeekUTC(d) {
  const day = d.getUTCDay();
  const diff = (day === 0 ? -6 : 1 - day);
  return new Date(d.getTime() + diff * 86400000);
}

function snapshotRange() {
  const anchor = dataAnchorDate();
  const anchorMid = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), anchor.getUTCDate()));
  if (Snapshot.periodMode === 'lastweek') {
    const curMon = mondayOfWeekUTC(anchorMid);
    const lastMon = new Date(curMon.getTime() - 7 * 86400000);
    const lastSun = new Date(curMon.getTime() - 1 * 86400000);
    return { from: lastMon, to: lastSun, label: 'Last week (' + fmtDate(lastMon) + ' – ' + fmtDate(lastSun) + ')' };
  }
  if (Snapshot.periodMode === 'thismonth') {
    const from = new Date(Date.UTC(anchorMid.getUTCFullYear(), anchorMid.getUTCMonth(), 1));
    return { from, to: anchorMid, label: 'This month (' + fmtDate(from) + ' \u2013 ' + fmtDate(anchorMid) + ')' };
  }
  if (Snapshot.periodMode === 'custom') {
    const from = Snapshot.from ? parseDateLoose(Snapshot.from) : null;
    const to = Snapshot.to ? parseDateLoose(Snapshot.to) : null;
    return { from, to, label: (from && to) ? ('Custom (' + fmtDate(from) + ' \u2013 ' + fmtDate(to) + ')') : 'Custom range - pick both dates' };
  }
  const mon = mondayOfWeekUTC(anchorMid);
  const sun = new Date(mon.getTime() + 6 * 86400000);
  return { from: mon, to: sun, label: 'This week (' + fmtDate(mon) + ' – ' + fmtDate(sun) + ')' };
}

/** Agar index.html purana/cached ho to popup ka markup hum khud bana dete hain. */
function ensureSnapshotDom() {
  if (!document.getElementById('snapshot-overlay')) {
    const div = document.createElement('div');
    div.id = 'snapshot-overlay';
    div.className = 'drill-overlay snapshot-overlay';
    div.style.display = 'none';
    div.innerHTML = snapshotPanelHtml();
    document.body.appendChild(div);
  } else if (!document.getElementById('snapshot-body-map')) {
    // Markup purana hai (sirf cards wala) — poora andar ka hissa refresh kar dete hain.
    document.getElementById('snapshot-overlay').innerHTML = snapshotPanelHtml();
  }

  if (!document.getElementById('btn-open-snapshot')) {
    const btn = document.createElement('button');
    btn.id = 'btn-open-snapshot';
    btn.className = 'ghost-btn';
    btn.textContent = '\uD83D\uDCCC Top Items Snapshot';
    const footer = document.querySelector('.sidebar-footer');
    if (footer) footer.insertBefore(btn, footer.firstChild);
    else document.body.appendChild(btn);
  }
}

function snapshotPanelHtml() {
  return '<div class="drill-panel snapshot-panel">' +
    '<div class="drill-head"><div>' +
      '<h2>Top Items Snapshot</h2>' +
      '<div id="snapshot-subtitle" class="drill-subtitle"></div>' +
    '</div><button id="snapshot-close" class="drill-close" title="Close (Esc)">&times;</button></div>' +

    '<div class="snapshot-controls">' +
      '<div class="seg-control inline" id="snapshot-view-tabs">' +
        '<button class="seg-btn active" data-view="map">Mind Map</button>' +
        '<button class="seg-btn" data-view="cards">Top Lists</button>' +
      '</div>' +

      '<div class="seg-control inline" id="snapshot-period-btns">' +
        '<button class="seg-btn active" data-p="thisweek">This week</button>' +
        '<button class="seg-btn" data-p="lastweek">Last week</button>' +
        '<button class="seg-btn" data-p="thismonth">This month</button>' +
        '<button class="seg-btn" data-p="custom">Custom</button>' +
      '</div>' +
      '<span class="date-range-inline" id="snapshot-custom-range" style="display:none;">' +
        '<span class="dr-label">From</span><input type="date" id="snapshot-from" class="text-input">' +
        '<span class="dr-label">To</span><input type="date" id="snapshot-to" class="text-input">' +
      '</span>' +
    '</div>' +

    '<div id="snapshot-total" class="drill-subtitle snapshot-total"></div>' +

    '<div id="snapshot-body-map" class="snapshot-body">' +
      '<div class="map-toolbar">' +
        '<label class="toolbar-label">Design:</label>' +
        '<select id="snapmap-style" class="select">' +
          SNAP_STYLES.map(function (st) {
            return '<option value="' + st.id + '"' + ((Snapshot.config.mapStyle || 'tree-h') === st.id ? ' selected' : '') + '>' + st.name + '</option>';
          }).join('') +
        '</select>' +
        '<span class="map-sep"></span>' +
        '<button class="ghost-btn small" id="snapmap-zoom-out">&minus;</button>' +
        '<span id="snapmap-zoom-label" class="drill-count">100%</span>' +
        '<button class="ghost-btn small" id="snapmap-zoom-in">+</button>' +
        '<button class="ghost-btn small" id="snapmap-fit">Fit</button>' +
        '<button class="ghost-btn small" id="snapmap-expand-all">Expand all</button>' +
        '<button class="ghost-btn small" id="snapmap-collapse-all">Collapse</button>' +
        '<span class="drill-count map-hint">Click = details \u00b7 drag = move \u00b7 Ctrl+scroll = zoom</span>' +
      '</div>' +
      '<div class="snapmap-wrap" id="snapmap-wrap"><svg id="snapmap-svg"></svg></div>' +
    '</div>' +

    '<div id="snapshot-body-cards" class="snapshot-body" style="display:none;">' +
      '<div id="snapshot-grid" class="snap-grid"></div>' +
    '</div>' +

  '</div>';
}

function initSnapshot() {
  ensureSnapshotDom();
  loadSnapshotConfig();

  document.getElementById('snapshot-close').addEventListener('click', closeSnapshot);
  document.getElementById('snapshot-overlay').addEventListener('click', e => {
    if (e.target.id === 'snapshot-overlay' && modalIsTop('snapshot-overlay')) closeSnapshot();
  });

  document.querySelectorAll('#snapshot-period-btns .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#snapshot-period-btns .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      Snapshot.periodMode = btn.dataset.p;
      document.getElementById('snapshot-custom-range').style.display = btn.dataset.p === 'custom' ? '' : 'none';
      if (btn.dataset.p !== 'custom' || (Snapshot.from && Snapshot.to)) renderSnapshot();
    });
  });
  document.querySelectorAll('#snapshot-custom-range input').forEach(inp => {
    inp.addEventListener('change', () => {
      Snapshot.from = document.getElementById('snapshot-from').value || null;
      Snapshot.to = document.getElementById('snapshot-to').value || null;
      if (Snapshot.from && Snapshot.to) renderSnapshot();
    });
  });

  document.querySelectorAll('#snapshot-view-tabs .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#snapshot-view-tabs .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      Snapshot.view = btn.dataset.view;
      showSnapshotView();
      renderSnapshot();
    });
  });

  document.getElementById('btn-open-snapshot').addEventListener('click', () => openSnapshot());
  const gotoSet = document.getElementById('snapshot-goto-settings');
  if (gotoSet) gotoSet.addEventListener('click', function (e) {
    e.stopPropagation();
    settingsTab = 'snapshot';
    openSettings();
    document.querySelectorAll('#settings-tabs .seg-btn').forEach(function (b) {
      b.classList.toggle('active', b.dataset.stab === 'snapshot');
    });
    renderSettingsBody();
  });
  initSnapMapControls();
}

function showSnapshotView() {
  const v = Snapshot.view;
  document.getElementById('snapshot-body-map').style.display = v === 'map' ? '' : 'none';
  document.getElementById('snapshot-body-cards').style.display = v === 'cards' ? '' : 'none';
  document.getElementById('snapshot-total').style.display = '';
}

function openSnapshot() {
  document.getElementById('snapshot-overlay').style.display = 'flex';
  modalOpen('snapshot-overlay', closeSnapshot);
  showSnapshotView();
  renderSnapshot();
}
function closeSnapshot() {
  document.getElementById('snapshot-overlay').style.display = 'none';
  modalClose('snapshot-overlay');
}

function maybeAutoShowSnapshot() {
  if (Snapshot.shownThisSession) return;
  if (!Snapshot.config.autoShow) return;
  if (!salesRecords().length) return;
  Snapshot.shownThisSession = true;
  Snapshot.view = 'map';
  document.querySelectorAll('#snapshot-view-tabs .seg-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.view === 'map'));
  openSnapshot();
}

function snapshotRecords() {
  // range ko bahar nikalna zaroori hai — warna ye har record ke liye dobara
  // calculate hota hai aur poora data baar-baar scan hota hai (bahut slow).
  const range = snapshotRange();
  return salesRecords().filter(r => inPeriod(r, range));
}

function renderSnapshot() {
  const range = snapshotRange();
  document.getElementById('snapshot-subtitle').textContent = range.label;
  const recs = snapshotRecords();
  const totalSold = recs.reduce((s, r) => s + (recQty(r)), 0);
  document.getElementById('snapshot-total').textContent =
    recs.length ? (fmtNum(totalSold) + ' pcs sold across ' + recs.length.toLocaleString('en-IN') + ' bill lines')
                : 'No sales found in this period.';

  if (Snapshot.view === 'map') renderSnapMap(recs, totalSold);
  else if (Snapshot.view === 'cards') renderSnapshotCards(recs);
  else renderSnapshotSettings();
}

/* ---------- MIND MAP — 5 designs ---------- */

const SNAP_STYLES = [
  { id: 'tree-h',   name: 'Horizontal Tree' },
  { id: 'tree-v',   name: 'Vertical Org Chart' },
  { id: 'radial',   name: 'Radial Burst' },
  { id: 'sunburst', name: 'Sunburst Rings' },
  { id: 'treemap',  name: 'Treemap Blocks' }
];

const SNAP_PALETTE = ['#A6402C', '#1F6F5C', '#B9862F', '#4A6FA5', '#7A4CA0', '#C6784B', '#3E8E7E', '#8C5B3F'];
function snapColor(i) { return SNAP_PALETTE[i % SNAP_PALETTE.length]; }

/** Hierarchy banata hai: har level par top-N, aur unke andar agla level. */
function buildSnapTree(recs, levels, topN, depth, pathFilters) {
  if (depth >= levels.length) return [];
  const dim = levels[depth];
  if (!dim) return [];
  const map = new Map();
  recs.forEach(r => {
    const k = dimKey(r, dim);
    map.set(k, (map.get(k) || 0) + (recQty(r)));
  });
  const top = [...map.entries()].filter(([k]) => k !== '(blank)').sort((a, b) => b[1] - a[1]).slice(0, topN);

  return top.map(([key, qty]) => {
    const subset = recs.filter(r => dimKey(r, dim) === key);
    const filters = pathFilters.concat([{ field: dim, value: key }]);
    const path = filters.map(f => f.field + '=' + f.value).join('|');
    return {
      dim, key, qty, lines: subset.length, filters, path,
      children: buildSnapTree(subset, levels, topN, depth + 1, filters)
    };
  });
}

/** Collapse state lagata hai (sirf tree/radial designs mein kaam ka hai). */
function visibleChildren(node) {
  return SnapMap.expanded[node.path] ? node.children : [];
}

function truncateLabel(s, max) {
  s = String(s);
  return s.length > max ? s.slice(0, max - 1) + '\u2026' : s;
}

function snapTooltip(n) {
  return escapeHtml(n.dim + ': ' + n.key + ' \u2014 ' + fmtNum(n.qty) + ' pcs, ' + n.lines + ' bill lines');
}

/** Har design ek {body, w, h} lautata hai; sizing yahin ek jagah hoti hai. */
function renderSnapMap(recs, totalSold) {
  const svg = document.getElementById('snapmap-svg');
  const levels = (Snapshot.config.levels || []).filter(l => l);
  if (!levels.length) {
    svg.innerHTML = '<text x="20" y="30" class="snap-empty-text">Choose levels in the Settings tab.</text>';
    return;
  }
  if (!recs.length) {
    svg.innerHTML = '<text x="20" y="30" class="snap-empty-text">No sales found in this period.</text>';
    return;
  }

  const tree = buildSnapTree(recs, levels, Snapshot.config.topN || 5, 0, []);
  // Default collapsed — user manually expands whichever branch they want to see.

  const style = Snapshot.config.mapStyle || 'tree-h';
  let out;
  if (style === 'tree-v') out = layoutTreeV(tree, totalSold);
  else if (style === 'radial') out = layoutRadial(tree, totalSold);
  else if (style === 'sunburst') out = layoutSunburst(tree, totalSold);
  else if (style === 'treemap') out = layoutTreemap(tree, totalSold);
  else out = layoutTreeH(tree, totalSold);

  SnapMap.w = out.w; SnapMap.h = out.h;
  // viewBox ki jagah asli pixel size — warna expand karne par sab kuch
  // itna chhota ho jata tha ki dikhta hi nahi tha.
  svg.removeAttribute('viewBox');
  svg.innerHTML = '<g id="snapmap-viewport">' + out.body + '</g>';
  applySnapMapTransform();
  wireSnapMapNodes();
}

function applySnapMapTransform() {
  const svg = document.getElementById('snapmap-svg');
  const vp = document.getElementById('snapmap-viewport');
  const z = SnapMap.zoom;
  if (vp) vp.setAttribute('transform', 'scale(' + z + ')');
  if (svg) {
    svg.setAttribute('width', Math.max(10, Math.round(SnapMap.w * z)));
    svg.setAttribute('height', Math.max(10, Math.round(SnapMap.h * z)));
  }
  const lbl = document.getElementById('snapmap-zoom-label');
  if (lbl) lbl.textContent = Math.round(z * 100) + '%';
}

function wireSnapMapNodes() {
  const svg = document.getElementById('snapmap-svg');
  svg.querySelectorAll('[data-toggle]').forEach(g => {
    g.addEventListener('click', e => {
      e.stopPropagation();
      const p = g.dataset.toggle;
      SnapMap.expanded[p] = !SnapMap.expanded[p];
      renderSnapshot();
    });
  });
  svg.querySelectorAll('[data-path]').forEach(g => {
    g.addEventListener('click', e => {
      e.stopPropagation();
      const node = findSnapNode(g.dataset.path);
      if (!node) return;
      closeSnapshot();
      openDrillPath(node.filters, snapshotRange());
    });
  });
}

let _snapIndex = {};
function indexNodes(nodes) {
  nodes.forEach(n => { _snapIndex[n.path] = n; indexNodes(n.children); });
}
function findSnapNode(path) { return _snapIndex[path]; }

/* ===== 1. HORIZONTAL TREE ===== */
function layoutTreeH(tree, totalSold) {
  _snapIndex = {}; indexNodes(tree);
  const NODE_W = 210, NODE_H = 30, COL_W = 268, ROW_H = 40, ROOT_W = 150, PAD = 24;
  const state = { leaf: 0, maxDepth: 0, all: [] };

  (function walk(nodes, depth) {
    nodes.forEach(n => {
      n.depth = depth;
      n.x = PAD + ROOT_W + 60 + depth * COL_W;
      const kids = visibleChildren(n);
      if (kids.length) {
        walk(kids, depth + 1);
        n.y = (kids[0].y + kids[kids.length - 1].y) / 2;
      } else {
        n.y = PAD + state.leaf * ROW_H; state.leaf++;
      }
      state.maxDepth = Math.max(state.maxDepth, depth);
      state.all.push(n);
    });
  })(tree, 0);

  const rootY = tree.length ? (tree[0].y + tree[tree.length - 1].y) / 2 : PAD;
  let links = '', nodes = '';

  function link(x1, y1, x2, y2, color) {
    const mx = (x1 + x2) / 2;
    return '<path class="snap-link" style="stroke:' + color + '" d="M' + x1 + ',' + y1 +
           ' C' + mx + ',' + y1 + ' ' + mx + ',' + y2 + ' ' + x2 + ',' + y2 + '"></path>';
  }

  tree.forEach(n => { links += link(PAD + ROOT_W, rootY + NODE_H / 2, n.x, n.y + NODE_H / 2, snapColor(0)); });
  (function walkLinks(nodes2) {
    nodes2.forEach(p => {
      visibleChildren(p).forEach(c => {
        links += link(p.x + NODE_W, p.y + NODE_H / 2, c.x, c.y + NODE_H / 2, snapColor(c.depth));
      });
      walkLinks(visibleChildren(p));
    });
  })(tree);

  nodes += '<g class="snap-node snap-root" transform="translate(' + PAD + ',' + rootY + ')">' +
    '<rect width="' + ROOT_W + '" height="' + NODE_H + '" rx="7"></rect>' +
    '<text x="12" y="' + (NODE_H / 2 + 4) + '" class="snap-node-label">All sales</text>' +
    '<text x="' + (ROOT_W - 12) + '" y="' + (NODE_H / 2 + 4) + '" class="snap-node-qty" text-anchor="end">' + fmtNum(totalSold) + '</text></g>';

  const maxByDepth = {};
  state.all.forEach(n => { maxByDepth[n.depth] = Math.max(maxByDepth[n.depth] || 0, n.qty); });

  state.all.forEach(n => {
    const share = maxByDepth[n.depth] ? n.qty / maxByDepth[n.depth] : 0;
    const col = snapColor(n.depth);
    nodes += '<g class="snap-node" transform="translate(' + n.x + ',' + n.y + ')" data-path="' + escapeHtml(n.path) + '">' +
      '<rect class="snap-node-bar" style="fill:' + col + '" width="' + Math.max(4, NODE_W * share) + '" height="' + NODE_H + '" rx="6"></rect>' +
      '<rect class="snap-node-box" width="' + NODE_W + '" height="' + NODE_H + '" rx="6"></rect>' +
      '<text x="10" y="' + (NODE_H / 2 + 4) + '" class="snap-node-label">' + escapeHtml(truncateLabel(n.key, 20)) + '</text>' +
      '<text x="' + (NODE_W - 10) + '" y="' + (NODE_H / 2 + 4) + '" class="snap-node-qty" text-anchor="end">' + fmtNum(n.qty) + '</text>' +
      '<title>' + snapTooltip(n) + '</title></g>';
    if (n.children.length) {
      nodes += '<g class="snap-toggle" transform="translate(' + (n.x + NODE_W + 9) + ',' + (n.y + NODE_H / 2) + ')" data-toggle="' + escapeHtml(n.path) + '">' +
        '<circle r="9"></circle><text y="4" text-anchor="middle">' + (SnapMap.expanded[n.path] ? '\u2212' : '+') + '</text></g>';
    }
  });

  return {
    body: links + nodes,
    w: PAD * 2 + ROOT_W + 60 + (state.maxDepth + 1) * COL_W + 40,
    h: Math.max(PAD * 2 + state.leaf * ROW_H, 240)
  };
}

/* ===== 2. VERTICAL ORG CHART ===== */
function layoutTreeV(tree, totalSold) {
  _snapIndex = {}; indexNodes(tree);
  const NODE_W = 150, NODE_H = 44, COL_W = 168, ROW_H = 108, PAD = 30;
  const state = { leaf: 0, maxDepth: 0, all: [] };

  (function walk(nodes, depth) {
    nodes.forEach(n => {
      n.depth = depth;
      n.y = PAD + 70 + depth * ROW_H;
      const kids = visibleChildren(n);
      if (kids.length) {
        walk(kids, depth + 1);
        n.x = (kids[0].x + kids[kids.length - 1].x) / 2;
      } else {
        n.x = PAD + state.leaf * COL_W; state.leaf++;
      }
      state.maxDepth = Math.max(state.maxDepth, depth);
      state.all.push(n);
    });
  })(tree, 0);

  const rootX = tree.length ? (tree[0].x + tree[tree.length - 1].x) / 2 : PAD;
  const rootY = PAD;
  let links = '', nodes = '';

  function elbow(x1, y1, x2, y2, color) {
    const my = (y1 + y2) / 2;
    return '<path class="snap-link" style="stroke:' + color + '" d="M' + x1 + ',' + y1 + ' V' + my + ' H' + x2 + ' V' + y2 + '"></path>';
  }

  tree.forEach(n => { links += elbow(rootX + NODE_W / 2, rootY + NODE_H, n.x + NODE_W / 2, n.y, snapColor(0)); });
  (function walkLinks(ns) {
    ns.forEach(p => {
      visibleChildren(p).forEach(c => {
        links += elbow(p.x + NODE_W / 2, p.y + NODE_H, c.x + NODE_W / 2, c.y, snapColor(c.depth));
      });
      walkLinks(visibleChildren(p));
    });
  })(tree);

  nodes += '<g class="snap-node snap-root" transform="translate(' + rootX + ',' + rootY + ')">' +
    '<rect width="' + NODE_W + '" height="' + NODE_H + '" rx="8"></rect>' +
    '<text x="' + NODE_W / 2 + '" y="19" class="snap-node-label" text-anchor="middle">All sales</text>' +
    '<text x="' + NODE_W / 2 + '" y="35" class="snap-node-qty" text-anchor="middle">' + fmtNum(totalSold) + '</text></g>';

  state.all.forEach(n => {
    const col = snapColor(n.depth);
    nodes += '<g class="snap-node snap-card-node" transform="translate(' + n.x + ',' + n.y + ')" data-path="' + escapeHtml(n.path) + '">' +
      '<rect class="snap-node-box" width="' + NODE_W + '" height="' + NODE_H + '" rx="8"></rect>' +
      '<rect style="fill:' + col + '" width="5" height="' + NODE_H + '" rx="2.5"></rect>' +
      '<text x="' + (NODE_W / 2 + 3) + '" y="19" class="snap-node-label" text-anchor="middle">' + escapeHtml(truncateLabel(n.key, 16)) + '</text>' +
      '<text x="' + (NODE_W / 2 + 3) + '" y="35" class="snap-node-qty" text-anchor="middle">' + fmtNum(n.qty) + '</text>' +
      '<title>' + snapTooltip(n) + '</title></g>';
    if (n.children.length) {
      nodes += '<g class="snap-toggle" transform="translate(' + (n.x + NODE_W / 2) + ',' + (n.y + NODE_H + 11) + ')" data-toggle="' + escapeHtml(n.path) + '">' +
        '<circle r="9"></circle><text y="4" text-anchor="middle">' + (SnapMap.expanded[n.path] ? '\u2212' : '+') + '</text></g>';
    }
  });

  return {
    body: links + nodes,
    w: Math.max(PAD * 2 + state.leaf * COL_W, 400),
    h: PAD * 2 + 70 + (state.maxDepth + 1) * ROW_H
  };
}

/* ===== 3. RADIAL BURST ===== */
function layoutRadial(tree, totalSold) {
  _snapIndex = {}; indexNodes(tree);
  const RING = 165, PAD = 60;
  const state = { leaf: 0, maxDepth: 0, all: [] };

  (function count(nodes, depth) {
    nodes.forEach(n => {
      n.depth = depth;
      const kids = visibleChildren(n);
      if (kids.length) count(kids, depth + 1); else state.leaf++;
      state.maxDepth = Math.max(state.maxDepth, depth);
    });
  })(tree, 0);

  const totalLeaves = Math.max(1, state.leaf);
  let leafIdx = 0;
  (function place(nodes, depth) {
    nodes.forEach(n => {
      const kids = visibleChildren(n);
      if (kids.length) {
        place(kids, depth + 1);
        n.angle = (kids[0].angle + kids[kids.length - 1].angle) / 2;
      } else {
        n.angle = (leafIdx + 0.5) / totalLeaves * Math.PI * 2;
        leafIdx++;
      }
      n.r = (depth + 1) * RING;
      state.all.push(n);
    });
  })(tree, 0);

  const R = (state.maxDepth + 1) * RING + PAD + 90;
  const cx = R, cy = R;
  const px = n => cx + Math.cos(n.angle - Math.PI / 2) * n.r;
  const py = n => cy + Math.sin(n.angle - Math.PI / 2) * n.r;

  let links = '', nodes = '';
  function curve(x1, y1, x2, y2, color) {
    return '<path class="snap-link" style="stroke:' + color + '" d="M' + x1 + ',' + y1 +
           ' Q' + ((x1 + x2) / 2 + (cx - (x1 + x2) / 2) * 0.25) + ',' + ((y1 + y2) / 2 + (cy - (y1 + y2) / 2) * 0.25) +
           ' ' + x2 + ',' + y2 + '"></path>';
  }
  tree.forEach(n => { links += curve(cx, cy, px(n), py(n), snapColor(0)); });
  (function walkLinks(ns) {
    ns.forEach(p => {
      visibleChildren(p).forEach(c => { links += curve(px(p), py(p), px(c), py(c), snapColor(c.depth)); });
      walkLinks(visibleChildren(p));
    });
  })(tree);

  nodes += '<g class="snap-node snap-root" transform="translate(' + (cx - 62) + ',' + (cy - 20) + ')">' +
    '<rect width="124" height="40" rx="20"></rect>' +
    '<text x="62" y="18" class="snap-node-label" text-anchor="middle">All sales</text>' +
    '<text x="62" y="32" class="snap-node-qty" text-anchor="middle">' + fmtNum(totalSold) + '</text></g>';

  const maxByDepth = {};
  state.all.forEach(n => { maxByDepth[n.depth] = Math.max(maxByDepth[n.depth] || 0, n.qty); });

  state.all.forEach(n => {
    const x = px(n), y = py(n), col = snapColor(n.depth);
    const share = maxByDepth[n.depth] ? n.qty / maxByDepth[n.depth] : 0;
    const rad = 7 + share * 13;
    const left = Math.cos(n.angle - Math.PI / 2) < 0;
    nodes += '<g class="snap-node snap-radial-node" data-path="' + escapeHtml(n.path) + '">' +
      '<circle cx="' + x + '" cy="' + y + '" r="' + rad + '" style="fill:' + col + '"></circle>' +
      '<text x="' + (x + (left ? -(rad + 7) : rad + 7)) + '" y="' + (y - 1) + '" class="snap-node-label" text-anchor="' + (left ? 'end' : 'start') + '">' + escapeHtml(truncateLabel(n.key, 16)) + '</text>' +
      '<text x="' + (x + (left ? -(rad + 7) : rad + 7)) + '" y="' + (y + 12) + '" class="snap-node-qty" text-anchor="' + (left ? 'end' : 'start') + '">' + fmtNum(n.qty) + '</text>' +
      '<title>' + snapTooltip(n) + '</title></g>';
    if (n.children.length) {
      nodes += '<g class="snap-toggle" transform="translate(' + x + ',' + (y + rad + 13) + ')" data-toggle="' + escapeHtml(n.path) + '">' +
        '<circle r="8"></circle><text y="4" text-anchor="middle">' + (SnapMap.expanded[n.path] ? '\u2212' : '+') + '</text></g>';
    }
  });

  return { body: links + nodes, w: R * 2, h: R * 2 };
}

/* ===== 4. SUNBURST RINGS ===== */
function layoutSunburst(tree, totalSold) {
  _snapIndex = {}; indexNodes(tree);
  const RING = 92, PAD = 40;
  let maxDepth = 0;
  (function d(ns, depth) { ns.forEach(n => { n.depth = depth; maxDepth = Math.max(maxDepth, depth); d(n.children, depth + 1); }); })(tree, 0);

  const R = (maxDepth + 2) * RING + PAD;
  const cx = R, cy = R;
  let body = '';

  function arcPath(a0, a1, r0, r1) {
    const large = (a1 - a0) > Math.PI ? 1 : 0;
    const p = (a, r) => [cx + Math.cos(a - Math.PI / 2) * r, cy + Math.sin(a - Math.PI / 2) * r];
    const [x0, y0] = p(a0, r0), [x1, y1] = p(a1, r0), [x2, y2] = p(a1, r1), [x3, y3] = p(a0, r1);
    return 'M' + x0 + ',' + y0 + ' A' + r0 + ',' + r0 + ' 0 ' + large + ' 1 ' + x1 + ',' + y1 +
           ' L' + x2 + ',' + y2 + ' A' + r1 + ',' + r1 + ' 0 ' + large + ' 0 ' + x3 + ',' + y3 + ' Z';
  }

  (function draw(nodes, depth, startAngle, endAngle) {
    const sum = nodes.reduce((s, n) => s + n.qty, 0) || 1;
    let a = startAngle;
    nodes.forEach((n, i) => {
      const span = (endAngle - startAngle) * (n.qty / sum);
      const a0 = a, a1 = a + span;
      a = a1;
      const r0 = (depth + 1) * RING, r1 = (depth + 2) * RING - 6;
      const col = snapColor(depth === 0 ? i : n.depth + i);
      body += '<g class="snap-node snap-arc" data-path="' + escapeHtml(n.path) + '">' +
        '<path d="' + arcPath(a0, a1, r0, r1) + '" style="fill:' + col + '"></path>' +
        '<title>' + snapTooltip(n) + '</title></g>';
      // label agar segment kaafi bada ho
      if (span > 0.16) {
        const mid = (a0 + a1) / 2, rm = (r0 + r1) / 2;
        const lx = cx + Math.cos(mid - Math.PI / 2) * rm;
        const ly = cy + Math.sin(mid - Math.PI / 2) * rm;
        let deg = mid * 180 / Math.PI - 90;
        if (deg > 90 || deg < -90) deg += 180;
        body += '<text class="snap-arc-label" x="' + lx + '" y="' + ly + '" text-anchor="middle" transform="rotate(' + deg + ',' + lx + ',' + ly + ')">' +
          escapeHtml(truncateLabel(n.key, 14)) + '</text>' +
          '<text class="snap-arc-qty" x="' + lx + '" y="' + (ly + 13) + '" text-anchor="middle" transform="rotate(' + deg + ',' + lx + ',' + ly + ')">' + fmtNum(n.qty) + '</text>';
      }
      if (n.children.length) draw(n.children, depth + 1, a0, a1);
    });
  })(tree, 0, 0, Math.PI * 2);

  body += '<circle cx="' + cx + '" cy="' + cy + '" r="' + (RING - 6) + '" class="snap-sun-center"></circle>' +
    '<text x="' + cx + '" y="' + (cy - 4) + '" class="snap-node-label snap-sun-text" text-anchor="middle">All sales</text>' +
    '<text x="' + cx + '" y="' + (cy + 14) + '" class="snap-node-qty snap-sun-text" text-anchor="middle">' + fmtNum(totalSold) + '</text>';

  return { body, w: R * 2, h: R * 2 };
}

/* ===== 5. TREEMAP BLOCKS ===== */
function layoutTreemap(tree, totalSold) {
  _snapIndex = {}; indexNodes(tree);
  const W = 1180, H = 660, PAD = 16;
  let body = '';

  // slice-and-dice: har level par direction badalti hai
  function draw(nodes, x, y, w, h, depth, horizontal) {
    const sum = nodes.reduce((s, n) => s + n.qty, 0) || 1;
    let off = 0;
    nodes.forEach((n, i) => {
      const frac = n.qty / sum;
      const nw = horizontal ? w * frac : w;
      const nh = horizontal ? h : h * frac;
      const nx = horizontal ? x + off : x;
      const ny = horizontal ? y : y + off;
      off += horizontal ? nw : nh;

      const gap = depth === 0 ? 4 : 2;
      const bx = nx + gap, by = ny + gap, bw = Math.max(1, nw - gap * 2), bh = Math.max(1, nh - gap * 2);
      const col = snapColor(depth === 0 ? i : depth + i);
      const showLabel = bw > 62 && bh > 26;

      body += '<g class="snap-node snap-block" data-path="' + escapeHtml(n.path) + '">' +
        '<rect x="' + bx + '" y="' + by + '" width="' + bw + '" height="' + bh + '" rx="4" style="fill:' + col + ';fill-opacity:' + (depth === 0 ? 0.9 : 0.55) + '"></rect>' +
        (showLabel
          ? '<text x="' + (bx + 7) + '" y="' + (by + 16) + '" class="snap-block-label">' + escapeHtml(truncateLabel(n.key, Math.floor(bw / 8))) + '</text>' +
            '<text x="' + (bx + 7) + '" y="' + (by + 30) + '" class="snap-block-qty">' + fmtNum(n.qty) + '</text>'
          : '') +
        '<title>' + snapTooltip(n) + '</title></g>';

      // bachche andar, thoda header chhod kar
      if (n.children.length && bw > 90 && bh > 60) {
        const head = showLabel ? 34 : 4;
        draw(n.children, bx + 3, by + head, bw - 6, bh - head - 3, depth + 1, !horizontal);
      }
    });
  }
  draw(tree, PAD, PAD, W - PAD * 2, H - PAD * 2, 0, true);

  return { body, w: W, h: H };
}

function fitSnapMap() {
  const wrap = document.getElementById('snapmap-wrap');
  if (!wrap || !SnapMap.w) return;
  const availW = (wrap.clientWidth || 1000) - 24, availH = (wrap.clientHeight || 520) - 24;
  const z = Math.min(availW / SnapMap.w, availH / SnapMap.h);
  SnapMap.zoom = Math.max(0.15, Math.min(2, z));
  applySnapMapTransform();
  wrap.scrollLeft = 0; wrap.scrollTop = 0;
}

function setSnapZoom(z) {
  SnapMap.zoom = Math.max(0.15, Math.min(4, z));
  applySnapMapTransform();
}

function setAllSnapExpanded(val) {
  function walk(nodes) {
    nodes.forEach(n => { if (n.children.length) { SnapMap.expanded[n.path] = val; walk(n.children); } });
  }
  const levels = (Snapshot.config.levels || []).filter(l => l);
  walk(buildSnapTree(snapshotRecords(), levels, Snapshot.config.topN || 5, 0, []));
  renderSnapshot();
}

function initSnapMapControls() {
  const zi = document.getElementById('snapmap-zoom-in');
  const zo = document.getElementById('snapmap-zoom-out');
  const fit = document.getElementById('snapmap-fit');
  const ea = document.getElementById('snapmap-expand-all');
  const ca = document.getElementById('snapmap-collapse-all');
  const styleSel = document.getElementById('snapmap-style');
  const wrap = document.getElementById('snapmap-wrap');
  if (!wrap) return;

  if (zi) zi.addEventListener('click', () => setSnapZoom(SnapMap.zoom * 1.25));
  if (zo) zo.addEventListener('click', () => setSnapZoom(SnapMap.zoom / 1.25));
  if (fit) fit.addEventListener('click', fitSnapMap);
  if (ea) ea.addEventListener('click', () => setAllSnapExpanded(true));
  if (ca) ca.addEventListener('click', () => setAllSnapExpanded(false));
  if (styleSel) {
    styleSel.addEventListener('change', () => {
      Snapshot.config.mapStyle = styleSel.value;
      saveSnapshotConfigLocal();
      renderSnapshot();
      setTimeout(fitSnapMap, 0);
    });
  }

  wrap.addEventListener('wheel', e => {
    if (!e.ctrlKey && !e.metaKey) return;   // normal scroll chalta rahe
    e.preventDefault();
    setSnapZoom(SnapMap.zoom * (e.deltaY < 0 ? 1.12 : 1 / 1.12));
  }, { passive: false });

  // khaali jagah se drag karke scroll
  wrap.addEventListener('mousedown', e => {
    if (e.target.closest('[data-path]') || e.target.closest('[data-toggle]')) return;
    SnapMap.drag = { x: e.clientX, y: e.clientY, sl: wrap.scrollLeft, st: wrap.scrollTop };
    wrap.classList.add('grabbing');
    e.preventDefault();
  });
  window.addEventListener('mousemove', e => {
    if (!SnapMap.drag) return;
    wrap.scrollLeft = SnapMap.drag.sl - (e.clientX - SnapMap.drag.x);
    wrap.scrollTop = SnapMap.drag.st - (e.clientY - SnapMap.drag.y);
  });
  window.addEventListener('mouseup', () => {
    SnapMap.drag = null;
    wrap.classList.remove('grabbing');
  });
}

/* ---------- TOP LISTS (purana cards view) ---------- */

function renderSnapshotCards(recs) {
  const dims = (Snapshot.config.dims || []).filter(d => SNAPSHOT_ALL_DIMS.includes(d));
  const wrap = document.getElementById('snapshot-grid');
  if (!dims.length) {
    wrap.innerHTML = '<div class="empty-hint">No types selected - choose them in the Settings tab.</div>';
    return;
  }
  wrap.innerHTML = dims.map(dim => {
    const map = aggregateByDimension(recs, dim);
    const top = [...map.entries()].filter(([k]) => k !== '(blank)').sort((a, b) => b[1] - a[1]).slice(0, Snapshot.config.topN || 5);
    return '<div class="snap-card">' +
      '<h4>' + escapeHtml(dim) + '</h4>' +
      (top.length ? '<table class="snap-table">' + top.map(([k, v], i) =>
        '<tr class="snap-row" data-dim="' + escapeHtml(dim) + '" data-value="' + escapeHtml(k) + '">' +
          '<td class="snap-rank">' + (i + 1) + '</td>' +
          '<td class="snap-name">' + escapeHtml(k) + '</td>' +
          '<td class="num snap-qty">' + fmtNum(v) + '</td>' +
        '</tr>').join('') + '</table>'
        : '<div class="empty-hint">No data in this period.</div>') +
      '</div>';
  }).join('');

  wrap.querySelectorAll('.snap-row').forEach(tr => tr.addEventListener('click', () => {
    closeSnapshot();
    openDrillPath([{ field: tr.dataset.dim, value: tr.dataset.value }], snapshotRange());
  }));
}

/* ---------- SETTINGS ---------- */

function renderSnapshotSettings(hostId) {
  // Ye settings do jagah dikhti hain — snapshot popup mein aur Settings panel
  // mein. Isliye container ka id bahar se aata hai (duplicate id na bane).
  const wrap = document.getElementById(hostId || 'snapshot-settings-body');
  if (!wrap) return;
  const levels = Snapshot.config.levels || [];
  // Jitne columns available hain, utne hi level rakh sakte ho.
  const maxLevels = SNAPSHOT_ALL_DIMS.length;
  const levelSelect = (i) =>
    '<select class="select snap-level" data-i="' + i + '">' +
      (i > 0 ? '<option value="">\u2014 none \u2014</option>' : '') +
      SNAPSHOT_ALL_DIMS.map(d => '<option value="' + d + '"' + (levels[i] === d ? ' selected' : '') + '>' + d + '</option>').join('') +
    '</select>';

  let levelRows = '';
  for (let i = 0; i < maxLevels; i++) {
    levelRows += '<div><label class="toolbar-label">Level ' + (i + 1) + '</label>' + levelSelect(i) + '</div>';
  }

  wrap.innerHTML =
    '<h3 class="snap-set-title">Mind Map design</h3>' +
    '<div class="snap-style-grid">' +
      SNAP_STYLES.map(st =>
        '<label class="snap-style-opt' + ((Snapshot.config.mapStyle || 'tree-h') === st.id ? ' active' : '') + '">' +
          '<input type="radio" name="mapstyle" value="' + st.id + '"' + ((Snapshot.config.mapStyle || 'tree-h') === st.id ? ' checked' : '') + '>' +
          '<span class="snap-style-name">' + st.name + '</span>' +
        '</label>').join('') +
    '</div>' +

    '<h3 class="snap-set-title">Mind Map hierarchy</h3>' +
    '<p class="drill-subtitle">Choose which level opens inside which, from top to bottom. Use as many levels as you like - leave the rest as "none".</p>' +
    '<div class="snap-levels">' + levelRows + '</div>' +

    '<h3 class="snap-set-title">Which types to show in Top Lists</h3>' +
    '<div class="snap-checklist">' +
      SNAPSHOT_ALL_DIMS.map(d =>
        '<label class="snap-check"><input type="checkbox" value="' + d + '"' +
        ((Snapshot.config.dims || []).includes(d) ? ' checked' : '') + '> ' + d + '</label>').join('') +
    '</div>' +

    '<div class="connect-row" style="margin-top:16px;">' +
      '<label class="toolbar-label">Top N per level/type:</label>' +
      '<input type="number" id="snap-topn" class="text-input narrow" min="3" max="15" value="' + (Snapshot.config.topN || 5) + '">' +
    '</div>' +
    '<label class="toolbar-checkbox" style="margin-top:10px;">' +
      '<input type="checkbox" id="snap-autoshow"' + (Snapshot.config.autoShow ? ' checked' : '') + '> ' +
      'Open this popup automatically when data is loaded' +
    '</label>' +
    '';

  wrap.querySelectorAll('input[name="mapstyle"]').forEach(r => r.addEventListener('change', () => {
    wrap.querySelectorAll('.snap-style-opt').forEach(o => o.classList.toggle('active', o.querySelector('input').checked));
  }));

  // The snapshot page saves through the shared Save bar at the bottom, so this
  // block only runs if an older markup still has its own button.
  // Snapshot settings apply the moment you change them; the shared Save bar at
  // the bottom is what writes them to disk (and to the Google Sheet).
  function applySnapshotSettings(rerender) {
    const lv = [...wrap.querySelectorAll('.snap-level')].map(x => x.value).filter(v => v);
    Snapshot.config.levels = lv.filter((v, i) => lv.indexOf(v) === i);
    if (!Snapshot.config.levels.length) Snapshot.config.levels = ['Sub Section'];
    const styleRadio = wrap.querySelector('input[name="mapstyle"]:checked');
    if (styleRadio) Snapshot.config.mapStyle = styleRadio.value;
    Snapshot.config.dims = [...wrap.querySelectorAll('.snap-check input:checked')].map(c => c.value);
    const tn = wrap.querySelector('#snap-topn');
    if (tn) Snapshot.config.topN = Math.max(3, Math.min(15, parseInt(tn.value, 10) || 5));
    const as = wrap.querySelector('#snap-autoshow');
    if (as) Snapshot.config.autoShow = as.checked;
    SnapMap.expanded = {};
    saveSnapshotConfigLocal();
    if (rerender && document.getElementById('snapshot-overlay') &&
        document.getElementById('snapshot-overlay').style.display !== 'none') {
      renderSnapshot();
      setTimeout(fitSnapMap, 0);
    }
  }

  wrap.querySelectorAll('.snap-level, .snap-check input, #snap-topn, #snap-autoshow, input[name="mapstyle"]')
    .forEach(el => el.addEventListener('change', () => applySnapshotSettings(true)));
}


/* ---------------------------------------------------------------
   9. DASHBOARD
   --------------------------------------------------------------- */
let dashCharts = {};
const DashState = { grain: 'month' };

function initDashboard() {
  document.querySelectorAll('#dash-grain .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#dash-grain .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      DashState.grain = btn.dataset.grain;
      renderDashboard();
    });
  });
}

function renderDashboard() {
  renderBoard('dash');
  if (!document.getElementById('chart-trend')) return;   // old fixed charts were replaced by the builder

  const empty = document.getElementById('dashboard-empty');
  const body = document.getElementById('dashboard-body');
  if (!App.datasets.length) { empty.style.display = ''; body.style.display = 'none'; return; }
  empty.style.display = 'none'; body.style.display = '';

  const range = periodRange();
  const salesRecs = salesRecords().filter(r => inPeriod(r, range));
  const purchaseRecs = purchaseRecords().filter(r => inPeriod(r, range));
  const stockRecs = stockRecords();

  const sumQty = recs => recs.reduce((s, r) => s + (recQty(r)), 0);
  const distinctItems = recs => new Set(recs.map(r => r['Item Code']).filter(Boolean)).size;

  const soldQty = sumQty(salesRecs), stockQty = sumQty(stockRecs);
  const sellThrough = (soldQty + stockQty) > 0 ? (soldQty / (soldQty + stockQty)) * 100 : 0;

  const A = App.datasets.length ? buildAnalysis('Item Code', 30) : null;
  const deadCount = A ? A.rows.filter(r => r.status === 'Dead stock').length : 0;

  const erD = effectiveRange(range);
  // KPI boxes removed from the Dashboard on request.


  renderTrendChart(salesRecs, purchaseRecs);
  renderTopChart('chart-topbrands', salesRecs, 'Brand', 'Qty sold');
  renderTopChart('chart-topsections', salesRecs, 'Section', 'Qty sold');
  renderTopChart('chart-topsuppliers', salesRecs.length ? salesRecs : purchaseRecs, 'Supplier', 'Qty');
  renderTopChart('chart-topsizes', salesRecs, 'Size', 'Qty sold');
  renderStockSplitChart(stockRecs);
}

function destroyChart(id) { if (dashCharts[id]) { dashCharts[id].destroy(); delete dashCharts[id]; } }

function grainKey(d, grain) {
  if (grain === 'day') return fmtDate(d);
  if (grain === 'year') return String(d.getUTCFullYear());
  return MONTH_LABELS[d.getUTCMonth()] + ' ' + d.getUTCFullYear();
}

function grainSort(key, grain) {
  if (grain === 'year') return parseInt(key, 10) * 10000;
  if (grain === 'day') { const d = parseDateLoose(key); return d ? d.getTime() : 0; }
  const [mon, yr] = key.split(' ');
  return parseInt(yr, 10) * 12 + MONTH_LABELS.indexOf(mon);
}

/** Chart heading me batata hai ki grouping kya hai aur kaunsi date se kaunsi tak. */
function updateTrendHeading(labels) {
  const el = document.getElementById('trend-heading');
  if (!el) return;
  const g = DashState.grain;
  const word = g === 'day' ? 'day' : (g === 'year' ? 'year' : 'month');
  if (!labels || !labels.length) {
    el.innerHTML = 'Sales vs purchase trend <span class="h3-note">(no dated rows)</span>';
    return;
  }
  el.innerHTML = 'Sales vs purchase trend ' +
    '<span class="h3-note">by ' + word + ' \u00b7 ' +
    escapeHtml(labels[0]) + ' \u2192 ' + escapeHtml(labels[labels.length - 1]) +
    ' \u00b7 ' + labels.length + ' ' + word + (labels.length === 1 ? '' : 's') + '</span>';
}

function renderTrendChart(salesRecs, purchaseRecs) {
  destroyChart('chart-trend');
  const grain = DashState.grain;
  const byGrain = recs => {
    const m = new Map();
    recs.forEach(r => {
      if (!r.Date) return;
      const key = grainKey(r.Date, grain);
      m.set(key, (m.get(key) || 0) + (recQty(r)));
    });
    return m;
  };
  const salesM = byGrain(salesRecs), purchM = byGrain(purchaseRecs);
  const allKeys = [...new Set([...salesM.keys(), ...purchM.keys()])]
    .sort((a, b) => grainSort(a, grain) - grainSort(b, grain));
  updateTrendHeading(allKeys);

  const ctx = document.getElementById('chart-trend').getContext('2d');
  dashCharts['chart-trend'] = makeChart(ctx, {
    type: grain === 'year' ? 'bar' : 'line',
    data: {
      labels: allKeys,
      datasets: [
        { label: 'Sold qty', data: allKeys.map(k => salesM.get(k) || 0), borderColor: CHART_COLORS[0], backgroundColor: CHART_COLORS[0] + '33', tension: .25, fill: grain !== 'year', pointRadius: grain === 'day' ? 0 : 3 },
        { label: 'Purchased qty', data: allKeys.map(k => purchM.get(k) || 0), borderColor: CHART_COLORS[1], backgroundColor: CHART_COLORS[1] + '33', tension: .25, fill: grain !== 'year', pointRadius: grain === 'day' ? 0 : 3 }
      ]
    },
    options: Object.assign(chartOptions(), {
      scales: { x: { ticks: { font: { size: 10 }, maxTicksLimit: grain === 'day' ? 14 : 24 } }, y: { beginAtZero: true } }
    })
  });
}

function renderTopChart(canvasId, recs, field, label) {
  destroyChart(canvasId);
  const el = document.getElementById(canvasId);
  if (!el) return;
  const ctx = el.getContext('2d');
  if (!recs.length) { el.style.display = 'none'; return; }
  el.style.display = '';
  const map = aggregateByDimension(recs, field);
  const top = [...map.entries()].filter(([k]) => k !== '(blank)').sort((a, b) => b[1] - a[1]).slice(0, 8);
  if (!top.length) { el.style.display = 'none'; return; }
  const fullLabels = top.map(t => t[0]);
  dashCharts[canvasId] = makeChart(ctx, {
    type: 'bar',
    data: { labels: top.map(t => t[0].length > 22 ? t[0].slice(0, 22) + '…' : t[0]), datasets: [{ label, data: top.map(t => t[1]), backgroundColor: CHART_COLORS[3] }] },
    options: Object.assign(chartOptions(), {
      indexAxis: 'y',
      plugins: { legend: { display: false },
        tooltip: { callbacks: { title: items => fullLabels[items[0].dataIndex] + '  (click for details)' } } },
      onClick: (evt, els) => { if (els && els.length) openDrill(field, fullLabels[els[0].index]); }
    })
  });
}

function renderStockSplitChart(stockRecs) {
  destroyChart('chart-stocksplit');
  const el = document.getElementById('chart-stocksplit');
  if (!el) return;
  if (!stockRecs.length) { el.style.display = 'none'; return; }
  el.style.display = '';
  el.style.cursor = 'pointer';
  const map = aggregateByDimension(stockRecs, 'Section');
  const top = [...map.entries()].filter(([k]) => k !== '(blank)').sort((a, b) => b[1] - a[1]).slice(0, 8);
  const labels = top.map(t => t[0]);
  dashCharts['chart-stocksplit'] = makeChart(el.getContext('2d'), {
    type: 'doughnut',
    data: { labels, datasets: [{ data: top.map(t => t[1]), backgroundColor: top.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } },
        tooltip: { callbacks: { title: items => labels[items[0].dataIndex] + '  (click for stock details)' } }
      },
      onClick: (evt, els) => { if (els && els.length) openDrill('Section', labels[els[0].index]); }
    }
  });
}

/* ---------------------------------------------------------------
   9b. CONNECTIONS — Power BI jaisa wire-based data model
   --------------------------------------------------------------- */
const RelUI = { positions: {}, dragNode: null, dragPort: null, tempLine: null, selected: null };

function initRelations() {
  return;   // Connections tab removed
  if (!document.getElementById('rel-autodetect')) return;
  document.getElementById('rel-autodetect').addEventListener('click', () => {
    autoDetectRelationships();
    renderRelations();
    refreshAnalysisViews();
  });
  document.getElementById('rel-clear').addEventListener('click', () => {
    if (!App.relationships.length) { toast('There are no connections.'); return; }
    App.relationships = [];
    clearLookups();
    renderRelations();
    refreshAnalysisViews();
    toast('All connections removed.');
  });
  document.getElementById('rel-arrange').addEventListener('click', () => {
    RelUI.positions = {};
    renderRelations();
  });

  const canvas = document.getElementById('rel-canvas');
  canvas.addEventListener('mousemove', onRelMouseMove);
  canvas.addEventListener('mouseup', onRelMouseUp);
  canvas.addEventListener('mouseleave', onRelMouseUp);
}

function defaultPosition(idx) {
  const col = idx % 3, row = Math.floor(idx / 3);
  return { x: 30 + col * 300, y: 24 + row * 340 };
}

function renderRelations() {
  return;   // Connections tab removed
  const canvas = document.getElementById('rel-canvas');
  if (!canvas) return;   // Connections tab removed from UI — detection still runs in background
  const empty = document.getElementById('rel-empty');
  if (!App.datasets.length) {
    empty.style.display = '';
    canvas.style.display = 'none';
    document.getElementById('rel-list').innerHTML = '';
    return;
  }
  empty.style.display = 'none';
  canvas.style.display = '';

  App.datasets.forEach((ds, i) => {
    if (!RelUI.positions[ds.id]) RelUI.positions[ds.id] = defaultPosition(i);
  });

  const nodesHtml = App.datasets.map(ds => {
    const p = RelUI.positions[ds.id];
    return '<div class="rel-node" data-ds="' + ds.id + '" style="left:' + p.x + 'px;top:' + p.y + 'px;">' +
      '<div class="rel-node-head ' + typeTagClass(ds.type) + '">' +
        '<span class="rel-node-title">' + escapeHtml(ds.name) + '</span>' +
        '<span class="rel-node-type">' + ds.type + '</span>' +
      '</div>' +
      '<div class="rel-node-meta">' + ds.rowCount.toLocaleString('en-IN') + ' rows</div>' +
      '<div class="rel-fields">' +
        ds.fields.map(f =>
          '<div class="rel-field" data-ds="' + ds.id + '" data-field="' + escapeHtml(f) + '">' +
            '<span class="rel-port" data-ds="' + ds.id + '" data-field="' + escapeHtml(f) + '" title="Drag to connect"></span>' +
            '<span class="rel-field-name">' + escapeHtml(f) + '</span>' +
            (FIELD_KIND[f] ? '<span class="rel-field-kind">' + FIELD_KIND[f].slice(0, 3) + '</span>' : '') +
          '</div>').join('') +
      '</div></div>';
  }).join('');

  canvas.innerHTML = '<svg id="rel-wires"></svg>' + nodesHtml;
  wireRelNodeEvents();
  drawWires();
  renderRelList();
}

function wireRelNodeEvents() {
  document.querySelectorAll('.rel-node-head').forEach(head => {
    head.addEventListener('mousedown', e => {
      const node = head.closest('.rel-node');
      const canvasRect = document.getElementById('rel-canvas').getBoundingClientRect();
      RelUI.dragNode = {
        id: node.dataset.ds,
        offX: e.clientX - canvasRect.left - RelUI.positions[node.dataset.ds].x,
        offY: e.clientY - canvasRect.top - RelUI.positions[node.dataset.ds].y
      };
      e.preventDefault();
    });
  });

  document.querySelectorAll('.rel-port').forEach(port => {
    port.addEventListener('mousedown', e => {
      RelUI.dragPort = { dsId: port.dataset.ds, field: port.dataset.field };
      e.stopPropagation();
      e.preventDefault();
    });
    port.addEventListener('mouseup', e => {
      if (!RelUI.dragPort) return;
      const from = RelUI.dragPort;
      const to = { dsId: port.dataset.ds, field: port.dataset.field };
      RelUI.dragPort = null;
      removeTempLine();
      if (from.dsId === to.dsId) { toast('You cannot connect two columns from the same file.'); return; }
      addRelationship(from.dsId, from.field, to.dsId, to.field);
      e.stopPropagation();
    });
  });
}

function addRelationship(fromDsId, fromField, toDsId, toField) {
  const id = relationshipId(fromDsId, fromField, toDsId, toField);
  const rev = relationshipId(toDsId, toField, fromDsId, fromField);
  if (App.relationships.some(r => r.id === id || r.id === rev)) { toast('That connection already exists.'); return; }
  const a = App.datasets.find(d => d.id === fromDsId), b = App.datasets.find(d => d.id === toDsId);
  const score = scoreRelationship(a, fromField, b, toField);
  App.relationships.push({ id, fromDsId, fromField, toDsId, toField, enabled: true, score });
  clearLookups();
  renderRelations();
  refreshAnalysisViews();
  toast('Connection created: ' + fromField + ' ↔ ' + toField + ' (' + fmtNum(score.pct, 0) + '% match)');
}

function onRelMouseMove(e) {
  const canvas = document.getElementById('rel-canvas');
  const rect = canvas.getBoundingClientRect();

  if (RelUI.dragNode) {
    const p = RelUI.positions[RelUI.dragNode.id];
    p.x = Math.max(0, e.clientX - rect.left - RelUI.dragNode.offX);
    p.y = Math.max(0, e.clientY - rect.top - RelUI.dragNode.offY);
    const node = canvas.querySelector('.rel-node[data-ds="' + RelUI.dragNode.id + '"]');
    if (node) { node.style.left = p.x + 'px'; node.style.top = p.y + 'px'; }
    drawWires();
    return;
  }

  if (RelUI.dragPort) {
    const svg = document.getElementById('rel-wires');
    const start = portCenter(RelUI.dragPort.dsId, RelUI.dragPort.field);
    if (!start) return;
    let line = document.getElementById('rel-temp-line');
    if (!line) {
      line = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      line.id = 'rel-temp-line';
      line.setAttribute('class', 'rel-wire temp');
      svg.appendChild(line);
    }
    const ex = e.clientX - rect.left + canvas.scrollLeft;
    const ey = e.clientY - rect.top + canvas.scrollTop;
    line.setAttribute('d', bezier(start.x, start.y, ex, ey));
  }
}

function onRelMouseUp() {
  RelUI.dragNode = null;
  if (RelUI.dragPort) { RelUI.dragPort = null; removeTempLine(); }
}

function removeTempLine() {
  const l = document.getElementById('rel-temp-line');
  if (l) l.remove();
}

function portCenter(dsId, field) {
  const canvas = document.getElementById('rel-canvas');
  const port = canvas.querySelector('.rel-port[data-ds="' + dsId + '"][data-field="' + cssEscape(field) + '"]');
  if (!port) return null;
  const pr = port.getBoundingClientRect(), cr = canvas.getBoundingClientRect();
  return {
    x: pr.left - cr.left + canvas.scrollLeft + pr.width / 2,
    y: pr.top - cr.top + canvas.scrollTop + pr.height / 2
  };
}

function cssEscape(s) { return String(s).replace(/"/g, '\\"'); }

function bezier(x1, y1, x2, y2) {
  const dx = Math.max(40, Math.abs(x2 - x1) * 0.5);
  return 'M' + x1 + ',' + y1 + ' C' + (x1 - dx) + ',' + y1 + ' ' + (x2 + dx) + ',' + y2 + ' ' + x2 + ',' + y2;
}

function drawWires() {
  const svg = document.getElementById('rel-wires');
  if (!svg) return;
  const canvas = document.getElementById('rel-canvas');
  svg.setAttribute('width', canvas.scrollWidth);
  svg.setAttribute('height', canvas.scrollHeight);

  const parts = App.relationships.map(rel => {
    const a = portCenter(rel.fromDsId, rel.fromField);
    const b = portCenter(rel.toDsId, rel.toField);
    if (!a || !b) return '';
    const cls = 'rel-wire' + (rel.enabled ? '' : ' disabled') + (RelUI.selected === rel.id ? ' selected' : '');
    const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
    const pct = rel.score ? Math.round(rel.score.pct) : 0;
    return '<path class="' + cls + '" d="' + bezier(a.x, a.y, b.x, b.y) + '" data-rel="' + rel.id + '"></path>' +
      '<circle class="rel-wire-dot" cx="' + a.x + '" cy="' + a.y + '" r="4"></circle>' +
      '<circle class="rel-wire-dot" cx="' + b.x + '" cy="' + b.y + '" r="4"></circle>' +
      '<rect class="rel-wire-label-bg" x="' + (mx - 20) + '" y="' + (my - 9) + '" width="40" height="18" rx="9"></rect>' +
      '<text class="rel-wire-label" x="' + mx + '" y="' + (my + 4) + '" text-anchor="middle">' + pct + '%</text>';
  }).join('');

  svg.innerHTML = parts;
  svg.querySelectorAll('.rel-wire').forEach(p => {
    p.addEventListener('click', () => {
      RelUI.selected = RelUI.selected === p.dataset.rel ? null : p.dataset.rel;
      drawWires(); renderRelList();
    });
  });
}

function renderRelList() {
  const wrap = document.getElementById('rel-list');
  if (!App.relationships.length) {
    wrap.innerHTML = '<div class="empty-hint">No connections yet. Click "Auto-detect", or drag from a column dot onto a column in another file.</div>';
    return;
  }
  const nameOf = id => { const d = App.datasets.find(x => x.id === id); return d ? d.name : '?'; };

  wrap.innerHTML = '<table class="simple-table"><thead><tr>' +
    '<th>From</th><th>Column</th><th>To</th><th>Column</th><th>Match</th><th>Unique keys</th><th>Can look up</th><th>Active</th><th></th></tr></thead><tbody>' +
    App.relationships.map(rel => {
      const sc = rel.score || {};
      const pct = sc.pct || 0;
      const cls = pct >= 70 ? 'ok' : (pct >= 30 ? 'warn' : 'bad');
      return '<tr class="' + (RelUI.selected === rel.id ? 'row-selected' : '') + '">' +
        '<td>' + escapeHtml(nameOf(rel.fromDsId)) + '</td>' +
        '<td><code>' + escapeHtml(rel.fromField) + '</code></td>' +
        '<td>' + escapeHtml(nameOf(rel.toDsId)) + '</td>' +
        '<td><code>' + escapeHtml(rel.toField) + '</code></td>' +
        '<td><span class="match-pill ' + cls + '">' + fmtNum(pct, 0) + '%</span></td>' +
        '<td>' + (sc.toDistinct || 0).toLocaleString('en-IN') + '</td>' +
        '<td>' + (sc.canEnrich
            ? '<span class="match-pill ok" title="This is a true key, so columns from the other file can be used through it">Yes</span>'
            : '<span class="match-pill warn" title="Values are not unique, so this link is informational only - column lookup is disabled">No</span>') + '</td>' +
        '<td><input type="checkbox" class="rel-toggle" data-id="' + rel.id + '"' + (rel.enabled ? ' checked' : '') + '></td>' +
        '<td><button class="ghost-btn small rel-del" data-id="' + rel.id + '">Remove</button></td>' +
      '</tr>';
    }).join('') + '</tbody></table>' +
    '<p class="rel-help"><strong>Match %</strong> = how many rows in the first file found a matching value in the second. ' +
    '<strong>Unique keys</strong> = how many distinct values the target column has. ' +
    '<strong>Can look up</strong> = whether a column from the other file can be used in this file analysis. ' +
    'This is only "Yes" when the values are nearly unique (like Item Code / barcode). ' +
    'A column like Section shows a 100% match, but looking values up through it would be wrong - one Section contains thousands of different items.</p>';

  wrap.querySelectorAll('.rel-toggle').forEach(cb => cb.addEventListener('change', () => {
    const rel = App.relationships.find(r => r.id === cb.dataset.id);
    if (rel) rel.enabled = cb.checked;
    clearLookups(); drawWires(); refreshAnalysisViews();
  }));
  wrap.querySelectorAll('.rel-del').forEach(btn => btn.addEventListener('click', () => {
    App.relationships = App.relationships.filter(r => r.id !== btn.dataset.id);
    clearLookups(); renderRelations(); refreshAnalysisViews();
  }));
}

function refreshAnalysisViews() {
  renderDashboard();
  renderPerformance();
  renderInsights();
  if (Drill.open) renderDrill();
}

/* ---------------------------------------------------------------
   10. SESSION SAVE / LOAD
   --------------------------------------------------------------- */
function initSession() {
  document.getElementById('btn-save-session').addEventListener('click', saveSession);
  document.getElementById('session-load-input').addEventListener('change', e => { if (e.target.files[0]) loadSession(e.target.files[0]); });
}

function saveSession() {
  if (!App.datasets.length) { toast('Nothing loaded yet.'); return; }
  const out = {
    version: 2,
    savedAt: new Date().toISOString(),
    datasets: App.datasets.map(ds => ({
      id: ds.id, name: ds.name, type: ds.type, fields: ds.fields,
      origin: ds.origin || null, mapping: ds.mapping || null, headerIdx: ds.headerIdx || 0,
      records: ds.records.map(r => {
        const o = {};
        ds.fields.forEach(f => { const v = r[f]; o[f] = v instanceof Date ? { __date: v.toISOString() } : v; });
        return o;
      })
    })),
    relationships: App.relationships.map(r => ({
      id: r.id, fromDsId: r.fromDsId, fromField: r.fromField,
      toDsId: r.toDsId, toField: r.toField, enabled: r.enabled
    })),
    positions: RelUI.positions,
    period: App.period
  };
  downloadBlob(JSON.stringify(out), 'stockledger-session.json', 'application/json');
  toast('Session saved.');
}

function loadSession(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const parsed = JSON.parse(e.target.result);
      // v1 = plain array of datasets; v2 = object with relationships too
      const list = Array.isArray(parsed) ? parsed : parsed.datasets;
      if (!list || !list.length) throw new Error('no datasets in file');

      const idMap = {};
      list.forEach(ds => {
        const newId = uid();
        idMap[ds.id || newId] = newId;
        const records = ds.records.map(r => {
          const o = { __ds: newId };
          ds.fields.forEach(f => {
            let v = r[f];
            if (v && typeof v === 'object' && v.__date) v = new Date(v.__date);
            // Purani ya haath se badli hui session file mein date string bhi
            // ho sakti hai — usse wapas asli Date bana lete hain.
            else if (typeof v === 'string' && FIELD_KIND[f] === 'date') v = parseDateLoose(v);
            if (v instanceof Date && isNaN(v.getTime())) v = null;
            o[f] = v;
          });
          return o;
        });
        App.datasets.push({
          id: newId, name: ds.name, type: ds.type, fields: ds.fields, records,
          rowCount: records.length, colorIdx: App.nextDsColor++,
          origin: ds.origin || null, mapping: ds.mapping || null, headerIdx: ds.headerIdx || 0
        });
      });

      if (!Array.isArray(parsed)) {
        (parsed.relationships || []).forEach(rel => {
          const f = idMap[rel.fromDsId], t = idMap[rel.toDsId];
          if (!f || !t) return;
          App.relationships.push({
            id: relationshipId(f, rel.fromField, t, rel.toField),
            fromDsId: f, fromField: rel.fromField, toDsId: t, toField: rel.toField,
            enabled: rel.enabled !== false
          });
        });
        if (parsed.positions) {
          Object.keys(parsed.positions).forEach(oldId => {
            if (idMap[oldId]) RelUI.positions[idMap[oldId]] = parsed.positions[oldId];
          });
        }
        if (parsed.period) App.period = parsed.period;
      }

      clearLookups();
      rescoreRelationships();
      refreshAfterDataChange();
      toast('Session loaded — ' + list.length + ' file(s), ' + App.relationships.length + ' connection(s).');
    } catch (err) {
      console.error(err);
      toast('That file does not look like a StockLedger session export.');
    }
  };
  reader.readAsText(file);
}

/* ---------------------------------------------------------------
   12. MODAL STACK — jo baad mein khula, wahi pehle band ho
   ---------------------------------------------------------------
   Pehle har overlay apne aap Esc par band ho jata tha, isliye do
   window khuli hon to dono ek saath band ho jati thi. Ab stack
   chalta hai: Esc / backdrop click sirf sabse upar wali band karta
   hai, ulte kram mein (LIFO) — jaise browser mein hota hai.
   --------------------------------------------------------------- */
const ModalStack = { items: [] };

function modalOpen(id, closeFn) {
  // agar pehle se stack mein hai to usko upar le aao
  ModalStack.items = ModalStack.items.filter(m => m.id !== id);
  ModalStack.items.push({ id, close: closeFn });
  updateModalDepths();
}

function modalClose(id) {
  ModalStack.items = ModalStack.items.filter(m => m.id !== id);
  updateModalDepths();
}

/** Sabse upar wali window band karta hai. Kuch khula ho to true. */
function modalCloseTop() {
  const top = ModalStack.items[ModalStack.items.length - 1];
  if (!top) return false;
  try { top.close(); } catch (e) { console.error(e); }
  ModalStack.items = ModalStack.items.filter(m => m !== top);
  updateModalDepths();
  return true;
}

function modalIsTop(id) {
  const top = ModalStack.items[ModalStack.items.length - 1];
  return !!top && top.id === id;
}

/** Baad mein khuli window hamesha upar dikhe. */
function updateModalDepths() {
  ModalStack.items.forEach((m, i) => {
    const el = typeof m.id === 'string' ? document.getElementById(m.id) : null;
    if (el) el.style.zIndex = String(900 + i * 10);
  });
}

function initModalStack() {
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    if (modalCloseTop()) { e.preventDefault(); e.stopPropagation(); }
  }, true);   // capture phase — purane individual Esc handlers se pehle
}

/* ---------------------------------------------------------------
   13. SETTINGS — ek jagah sab control (sidebar ke bottom se)
   --------------------------------------------------------------- */
const THEME_DEFAULT = {
  accent: '#A6402C',
  font: 'slab',        // slab | sans | serif | mono
  density: 'normal',   // compact | normal | comfortable
  fontSize: 13,        // table font size px
  zebra: true,
  gridLines: true,
  caretSize: 20,       // expand/collapse triangle size in px
  rowPad: 4,           // table row padding in px (row height)
  childRowScale: 70,   // drill-down row height as a % of the parent row
  gridColor: '#111111',      // table grid line colour
  gridWidth: 1,              // table grid line thickness (px)
  hoverColor: '#FFE3BF',     // row colour under the mouse (normal rows)
  drillHoverColor: '#D7E8F7',// row colour under the mouse (drill-down rows)
  openFill: '#FDF3E7',       // fill of a row opened in the drill-down
  openBorderColor: '#A6402C',// border of a row opened in the drill-down
  openBorderWidth: 2         // that border's thickness (px)
};

const ACCENT_CHOICES = [
  ['#A6402C', 'Rust (default)'],
  ['#1F6F5C', 'Teal'],
  ['#4A6FA5', 'Indigo'],
  ['#7A4CA0', 'Purple'],
  ['#B9862F', 'Gold'],
  ['#2F5D3A', 'Forest'],
  ['#8C2E1B', 'Brick'],
  ['#26303A', 'Charcoal']
];

const FONT_CHOICES = [
  ['slab',  'Zilla Slab + IBM Plex', "'Zilla Slab', Georgia, serif", "'IBM Plex Sans', system-ui, sans-serif"],
  ['sans',  'All Sans (clean)',      "'IBM Plex Sans', system-ui, sans-serif", "'IBM Plex Sans', system-ui, sans-serif"],
  ['serif', 'Serif (classic)',       "Georgia, 'Times New Roman', serif", "Georgia, 'Times New Roman', serif"],
  ['mono',  'Mono (data-heavy)',     "'IBM Plex Mono', monospace", "'IBM Plex Mono', monospace"]
];

const Theme = Object.assign({}, THEME_DEFAULT);

function loadTheme() {
  try {
    const raw = Store.get('sl_theme');
    if (raw) Object.assign(Theme, THEME_DEFAULT, JSON.parse(raw));
  } catch (e) {}
  applyTheme();
}

function saveTheme() {
  Store.set('sl_theme', JSON.stringify(Theme));
  applyTheme();
}

function applyTheme() {
  const root = document.documentElement;
  root.style.setProperty('--rust', Theme.accent);
  root.style.setProperty('--rust-dark', shadeColor(Theme.accent, -18));

  const f = FONT_CHOICES.find(x => x[0] === Theme.font) || FONT_CHOICES[0];
  root.style.setProperty('--font-display', f[2]);
  root.style.setProperty('--font-body', f[3]);
  root.style.setProperty('--table-font-size', Theme.fontSize + 'px');
  root.style.setProperty('--caret-size', (Theme.caretSize || 20) + 'px');
  root.style.setProperty('--row-pad', (Theme.rowPad === undefined ? 4 : Theme.rowPad) + 'px');
  // Drill-down rows sit shorter than their parent row (default 70%).
  const crs = (Theme.childRowScale === undefined ? 70 : Theme.childRowScale) / 100;
  root.style.setProperty('--child-row-scale', String(crs));

  // Table grid, hover colour and the drill-down open-row highlight
  root.style.setProperty('--grid-color', Theme.gridColor || '#111111');
  root.style.setProperty('--drill-hover', Theme.drillHoverColor || '#D7E8F7');
  root.style.setProperty('--grid-width', (Theme.gridWidth === undefined ? 1 : Theme.gridWidth) + 'px');
  root.style.setProperty('--xl-hover', Theme.hoverColor || '#FFE3BF');
  root.style.setProperty('--lvl0-open', Theme.openFill || '#FDF3E7');
  root.style.setProperty('--open-bc', Theme.openBorderColor || '#A6402C');
  root.style.setProperty('--open-bw', (Theme.openBorderWidth === undefined ? 2 : Theme.openBorderWidth) + 'px');

  const b = document.body;
  ['density-compact', 'density-normal', 'density-comfortable'].forEach(c => b.classList.remove(c));
  b.classList.add('density-' + Theme.density);
  b.classList.toggle('no-zebra', !Theme.zebra);
  b.classList.toggle('no-gridlines', !Theme.gridLines);
}

/** Hex color ko halka/gehra karta hai. */
function shadeColor(hex, percent) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  if (!m) return hex;
  const adj = v => Math.max(0, Math.min(255, Math.round(parseInt(v, 16) * (100 + percent) / 100)));
  const to2 = v => v.toString(16).padStart(2, '0');
  return '#' + to2(adj(m[1])) + to2(adj(m[2])) + to2(adj(m[3]));
}

function ensureSettingsDom() {
  if (!document.getElementById('settings-overlay')) {
    const d = document.createElement('div');
    d.id = 'settings-overlay';
    d.className = 'drill-overlay settings-overlay';
    d.style.display = 'none';
    d.innerHTML =
      '<div class="drill-panel settings-panel">' +
        '<div class="drill-head"><div>' +
          '<h2>Settings</h2>' +
          '<div class="drill-subtitle">Appearance, snapshot and behaviour - all in one place</div>' +
        '</div><button id="settings-close" class="drill-close" title="Close (Esc)">&times;</button></div>' +
        '<div class="seg-control" id="settings-tabs" style="margin:14px 0 12px;">' +
          '<button class="seg-btn active" data-stab="look">Look &amp; Feel</button>' +
          '<button class="seg-btn" data-stab="dashboard">01 Dashboard</button>' +
          '<button class="seg-btn" data-stab="performance">02 Performance</button>' +
          '<button class="seg-btn" data-stab="catalog">03 Catalog</button>' +
          '<button class="seg-btn" data-stab="pivot">04 Pivot</button>' +
          '<button class="seg-btn" data-stab="explore">05 Explore</button>' +
          '<button class="seg-btn" data-stab="import">06 Import</button>' +
          '<button class="seg-btn" data-stab="snapshot">Snapshot</button>' +
        '</div>' +
        '<div id="settings-body" class="settings-body"></div>' +
      '</div>';
    document.body.appendChild(d);
  }
  if (!document.getElementById('btn-open-settings')) {
    const btn = document.createElement('button');
    btn.id = 'btn-open-settings';
    btn.className = 'ghost-btn';
    btn.innerHTML = '\u2699 Settings';
    const footer = document.querySelector('.sidebar-footer');
    if (footer) footer.appendChild(btn); else document.body.appendChild(btn);
  }
}

let settingsTab = 'look';

function initSettings() {
  ensureSettingsDom();
  loadTheme();

  document.getElementById('btn-open-settings').addEventListener('click', openSettings);
  document.getElementById('settings-close').addEventListener('click', closeSettings);
  document.getElementById('settings-overlay').addEventListener('click', e => {
    if (e.target.id === 'settings-overlay' && modalIsTop('settings-overlay')) closeSettings();
  });
  document.querySelectorAll('#settings-tabs .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#settings-tabs .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      settingsTab = btn.dataset.stab;
      renderSettingsBody();
    });
  });
}

function openSettings() {
  document.getElementById('settings-overlay').style.display = 'flex';
  modalOpen('settings-overlay', closeSettings);
  renderSettingsBody();
}
function closeSettings() {
  document.getElementById('settings-overlay').style.display = 'none';
  modalClose('settings-overlay');
}

/** Settings apply as you change them; this button is the visible confirmation
 *  and forces everything to disk in one go. */
function settingsSaveBarHtml() {
  return '<div class="settings-savebar">' +
    '<button class="ghost-btn small primary" id="set-save">Save settings</button>' +
    '<span class="drill-count" id="set-saved-note"></span>' +
    '</div>';
}

function wireSettingsSave(wrap) {
  const btn = wrap.querySelector('#set-save');
  if (!btn) return;
  btn.addEventListener('click', () => {
    saveTheme(); saveBehaviour(); savePrefs(); saveCatPrefs();
    saveSnapshotConfigLocal(); saveReplen(); saveReplenBulk(); saveColWidths();
    const n = wrap.querySelector('#set-saved-note');
    if (n) {
      n.textContent = 'Saved at ' + new Date().toLocaleTimeString();
      setTimeout(() => { if (n) n.textContent = ''; }, 4000);
    }
    toast('Settings saved.');
  });
}

function renderSettingsBody() {
  const wrap = document.getElementById('settings-body');
  const finish = () => { wrap.insertAdjacentHTML('beforeend', settingsSaveBarHtml()); wireSettingsSave(wrap); };
  if (settingsTab === 'snapshot') {
    wrap.innerHTML = '<div id="settings-snap-body"></div>';
    renderSnapshotSettings('settings-snap-body');
    finish();
    return;
  }
  // ---- 01 Dashboard ----
  if (settingsTab === 'dashboard') {
    wrap.innerHTML =
      '<h3 class="snap-set-title">Default trend grouping</h3>' +
      '<p class="drill-subtitle">Which grouping the Dashboard trend chart opens with.</p>' +
      '<div class="settings-row">' +
        '<select id="set-dash-grain" class="select">' +
          ['day', 'month', 'year'].map(function (g) {
            return '<option value="' + g + '"' + ((Prefs.dashGrain || 'month') === g ? ' selected' : '') + '>' +
              g.charAt(0).toUpperCase() + g.slice(1) + ' wise</option>';
          }).join('') +
        '</select>' +
      '</div>' +
      '<h3 class="snap-set-title">Chart layout</h3>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="set-boards-lock"' +
        (Boards.locked ? ' checked' : '') + '> Lock the chart layout (last saved arrangement stays put and cannot be edited)</label>' +
      '<p class="drill-subtitle" style="margin-top:6px;">Applies to both the Dashboard and the Product Performance charts. Unlock to move, resize, add or remove charts.</p>' +
      '<h3 class="snap-set-title">Default sales window</h3>' +
      '<div class="settings-row"><select id="set-dash-window" class="select">' +
        PERIOD_CHOICES.map(function (o) {
          return '<option value="' + o[0] + '"' + ((Prefs.defaultPeriod || 'all') === o[0] ? ' selected' : '') + '>' + o[1] + '</option>';
        }).join('') +
      '</select></div>' +
      '<p class="drill-subtitle" style="margin-top:10px;">The trend chart heading always shows the exact dates it covers, so Day / Month / Year is never ambiguous.</p>';

    const lk = wrap.querySelector('#set-boards-lock');
    if (lk) lk.addEventListener('change', function (e) { setBoardsLocked(e.target.checked); });
    renderBoardBackgroundSettings(wrap);
    wrap.querySelector('#set-dash-grain').addEventListener('change', function (e) {
      Prefs.dashGrain = e.target.value; savePrefs();
      DashState.grain = Prefs.dashGrain;
      document.querySelectorAll('#dash-grain .seg-btn').forEach(function (b) {
        b.classList.toggle('active', b.dataset.grain === Prefs.dashGrain);
      });
      renderDashboard();
    });
    wrap.querySelector('#set-dash-window').addEventListener('change', function (e) {
      Prefs.defaultPeriod = e.target.value; savePrefs();
    });
    finish();
    return;
  }

  // ---- 02 Product Performance ----
  if (settingsTab === 'performance') {
    wrap.innerHTML =
      '<h3 class="snap-set-title">How a row opens</h3>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="set-inline"' +
        (Behaviour.inlineExpand ? ' checked' : '') + '> Clicking a row opens its details inside the same table ' +
        '(turn this off to open a separate window instead)</label>' +
      '<label class="toolbar-checkbox" style="margin-top:8px;"><input type="checkbox" id="set-autofirst"' +
        (Behaviour.autoOpenFirst ? ' checked' : '') + '> Also open the top item automatically ' +
        '(off by default \u2014 rows open only when you click them)</label>' +

      '<h3 class="snap-set-title">Expand order</h3>' +
      '<p class="drill-subtitle">When you open a row, this is the order the levels open in \u2014 same idea as the mind map levels.</p>' +
      '<div class="snap-levels">' +
        PERF_CHAIN_CHOICES.map(function (_, i) {
          const cur = perfChain();
          return '<div><label class="toolbar-label">Level ' + (i + 1) + '</label>' +
            '<select class="select perf-chain-level" data-i="' + i + '">' +
              (i > 0 ? '<option value="">\u2014 none \u2014</option>' : '') +
              PERF_CHAIN_CHOICES.map(function (d) {
                return '<option value="' + d + '"' + (cur[i] === d ? ' selected' : '') + '>' + d + '</option>';
              }).join('') +
            '</select></div>';
        }).join('') +
      '</div>' +
      '<div style="margin-top:10px;"><button class="ghost-btn small primary" id="set-chain-save">Save expand order</button></div>' +

      '<h3 class="snap-set-title">Default grouping</h3>' +
      '<div class="settings-row"><select id="set-perf-group" class="select">' +
        PERF_CHAIN_CHOICES.map(function (d) {
          return '<option value="' + d + '"' + ((Prefs.perfGroupBy || 'Article No') === d ? ' selected' : '') + '>' + d + '</option>';
        }).join('') +
      '</select>' +
      '<label class="toolbar-label">Rows shown</label>' +
      '<input type="number" id="set-perf-limit" class="text-input narrow" min="100" max="5000" step="100" value="' + (Prefs.perfLimit || 800) + '"></div>';

    wrap.querySelector('#set-inline').addEventListener('change', function (e) {
      Behaviour.inlineExpand = e.target.checked; saveBehaviour(); renderPerformance();
    });
    wrap.querySelector('#set-autofirst').addEventListener('change', function (e) {
      Behaviour.autoOpenFirst = e.target.checked; saveBehaviour();
    });
    wrap.querySelector('#set-chain-save').addEventListener('click', function () {
      const picked = [...wrap.querySelectorAll('.perf-chain-level')].map(function (x) { return x.value; }).filter(Boolean);
      Behaviour.perfChain = picked.filter(function (v, i) { return picked.indexOf(v) === i; });
      if (!Behaviour.perfChain.length) Behaviour.perfChain = PERF_CHAIN_DEFAULT.slice();
      saveBehaviour(); PerfState.open = {}; renderPerformance();
      toast('Expand order saved.');
    });
    wrap.querySelector('#set-perf-group').addEventListener('change', function (e) {
      Prefs.perfGroupBy = e.target.value; savePrefs();
      const g = document.getElementById('perf-groupby');
      if (g) { g.value = Prefs.perfGroupBy; PerfState.open = {}; renderPerformance(); }
    });
    wrap.querySelector('#set-perf-limit').addEventListener('change', function (e) {
      Prefs.perfLimit = Math.max(100, Math.min(5000, parseInt(e.target.value, 10) || 800));
      savePrefs(); renderPerformance();
    });
    finish();
    return;
  }

  // ---- 03 Catalog ----
  if (settingsTab === 'catalog') { renderCatalogSettings(wrap); finish(); return; }

  // ---- 04 Pivot Builder ----
  if (settingsTab === 'pivot') {
    wrap.innerHTML =
      '<h3 class="snap-set-title">Defaults</h3>' +
      '<div class="settings-row">' +
        '<label class="toolbar-label">Opens in</label>' +
        '<select id="set-pivot-mode" class="select">' +
          '<option value="quick"' + ((Prefs.pivotMode || 'quick') === 'quick' ? ' selected' : '') + '>Quick Report (tick columns)</option>' +
          '<option value="drag"' + (Prefs.pivotMode === 'drag' ? ' selected' : '') + '>Drag &amp; Drop Builder</option>' +
        '</select>' +
        '<label class="toolbar-label">Date grouping</label>' +
        '<select id="set-pivot-grain" class="select">' +
          [['none', 'None'], ['day', 'Day'], ['week', 'Week'], ['month', 'Month'], ['quarter', 'Quarter'], ['year', 'Year']].map(function (o) {
            return '<option value="' + o[0] + '"' + ((Prefs.quickGrain || 'none') === o[0] ? ' selected' : '') + '>' + o[1] + '</option>';
          }).join('') +
        '</select>' +
      '</div>' +
      '<div class="settings-row">' +
        '<label class="toolbar-label">Sort</label>' +
        '<select id="set-pivot-sort" class="select">' +
          '<option value="desc"' + (Prefs.quickDesc !== false ? ' selected' : '') + '>Descending (largest first)</option>' +
          '<option value="asc"' + (Prefs.quickDesc === false ? ' selected' : '') + '>Ascending</option>' +
        '</select>' +
      '</div>';

    wrap.querySelector('#set-pivot-mode').addEventListener('change', function (e) {
      Prefs.pivotMode = e.target.value; savePrefs();
    });
    wrap.querySelector('#set-pivot-grain').addEventListener('change', function (e) {
      Prefs.quickGrain = e.target.value; savePrefs();
      QuickReport.dateGrain = Prefs.quickGrain;
      const g = document.getElementById('quick-date-grain');
      if (g) { g.value = Prefs.quickGrain; renderQuickReport(); }
    });
    wrap.querySelector('#set-pivot-sort').addEventListener('change', function (e) {
      Prefs.quickDesc = e.target.value === 'desc'; savePrefs();
      QuickReport.desc = Prefs.quickDesc;
      const g = document.getElementById('quick-sort');
      if (g) { g.value = e.target.value; renderQuickReport(); }
    });
    finish();
    return;
  }

  // ---- 05 Explore Rows ----
  if (settingsTab === 'explore') {
    wrap.innerHTML =
      '<h3 class="snap-set-title">Rows per page</h3>' +
      '<div class="settings-row">' +
        '<select id="set-explore-page" class="select">' +
          [50, 100, 200, 500, 1000].map(function (n) {
            return '<option value="' + n + '"' + ((Prefs.explorePageSize || 100) === n ? ' selected' : '') + '>' + n + ' rows</option>';
          }).join('') +
        '</select>' +
      '</div>' +
      '<p class="drill-subtitle" style="margin-top:10px;">Larger pages mean less clicking but a slower first draw on very big files.</p>';

    wrap.querySelector('#set-explore-page').addEventListener('change', function (e) {
      Prefs.explorePageSize = parseInt(e.target.value, 10) || 100;
      savePrefs();
      ExploreState.pageSize = Prefs.explorePageSize;
      ExploreState.page = 1;
      renderExplore();
    });
    finish();
    return;
  }

  // ---- 06 Import Data ----
  if (settingsTab === 'import') {
    wrap.innerHTML =
      '<h3 class="snap-set-title">After a file is loaded</h3>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="set-autosnap"' +
        (Snapshot.config.autoShow ? ' checked' : '') + '> Open the Top Items Snapshot automatically</label>' +
      '<label class="toolbar-checkbox" style="margin-top:8px;"><input type="checkbox" id="set-autolink"' +
        (Prefs.autoDetectLinks !== false ? ' checked' : '') + '> Detect links between files automatically</label>' +
      '<h3 class="snap-set-title">Saved data</h3>' +
      '<p class="drill-subtitle">Loaded files stay in this browser until you remove them on the Import tab.</p>' +
      '<button class="ghost-btn small" id="set-clear-data">Remove all loaded files</button>' +
      '<h3 class="snap-set-title">Reset</h3>' +
      '<button class="ghost-btn small" id="set-reset-theme">Reset Look &amp; Feel to defaults</button> ' +
      '<button class="ghost-btn small" id="set-reset-all">Clear all settings</button>' +
      '<p class="drill-subtitle" style="margin-top:10px;">Build ' + BUILD_VERSION + '</p>';

    wrap.querySelector('#set-autosnap').addEventListener('change', function (e) {
      Snapshot.config.autoShow = e.target.checked; saveSnapshotConfigLocal();
    });
    wrap.querySelector('#set-autolink').addEventListener('change', function (e) {
      Prefs.autoDetectLinks = e.target.checked; savePrefs();
    });
    wrap.querySelector('#set-clear-data').addEventListener('click', function () {
      App.datasets.slice().forEach(function (d) { removeDataset(d.id); });
      toast('All loaded files removed.');
    });
    wrap.querySelector('#set-reset-theme').addEventListener('click', function () {
      Object.assign(Theme, THEME_DEFAULT); saveTheme(); renderSettingsBody(); toast('Look & Feel has been reset.');
    });
    wrap.querySelector('#set-reset-all').addEventListener('click', function () {
      Store.remove('sl_theme'); Store.remove('sl_snapshot_config'); Store.remove('sl_behaviour');
      Store.remove('sl_prefs'); Store.remove('sl_colwidths');
      Object.assign(Theme, THEME_DEFAULT); saveTheme();
      toast('All settings cleared - please refresh the page.');
    });
    finish();
    return;
  }

  // --- Look & Feel ---
  wrap.innerHTML =
    '<h3 class="snap-set-title">Accent colour</h3>' +
    '<div class="swatch-grid">' +
      ACCENT_CHOICES.map(([hex, name]) =>
        '<button class="swatch' + (Theme.accent.toLowerCase() === hex.toLowerCase() ? ' active' : '') + '" ' +
        'data-accent="' + hex + '" title="' + name + '"><span style="background:' + hex + '"></span>' + name + '</button>').join('') +
    '</div>' +

    '<h3 class="snap-set-title">Font style</h3>' +
    '<div class="snap-style-grid">' +
      FONT_CHOICES.map(([id, name, disp]) =>
        '<label class="snap-style-opt' + (Theme.font === id ? ' active' : '') + '">' +
        '<input type="radio" name="themefont" value="' + id + '"' + (Theme.font === id ? ' checked' : '') + '>' +
        '<span class="snap-style-name" style="font-family:' + disp + '">' + name + '</span></label>').join('') +
    '</div>' +

    '<h3 class="snap-set-title">Table look</h3>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">Row height</label>' +
      '<input type="range" id="set-rowpad" min="1" max="14" step="1" value="' + (Theme.rowPad === undefined ? 4 : Theme.rowPad) + '">' +
      '<span class="drill-count" id="set-rowpad-val">' + (Theme.rowPad === undefined ? 4 : Theme.rowPad) + 'px</span>' +
      '<label class="toolbar-label">Font size</label>' +
      '<input type="range" id="set-fontsize" min="11" max="17" step="0.5" value="' + Theme.fontSize + '">' +
      '<span class="drill-count" id="set-fontsize-val">' + Theme.fontSize + 'px</span>' +
    '</div>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">Drill-down row height</label>' +
      '<input type="range" id="set-childrow" min="40" max="100" step="5" value="' +
        (Theme.childRowScale === undefined ? 70 : Theme.childRowScale) + '">' +
      '<span class="drill-count" id="set-childrow-val">' +
        (Theme.childRowScale === undefined ? 70 : Theme.childRowScale) + '% of the parent row</span>' +
    '</div>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">Expand triangle size</label>' +
      '<input type="range" id="set-caret" min="12" max="30" step="1" value="' + (Theme.caretSize || 20) + '">' +
      '<span class="drill-count" id="set-caret-val">' + (Theme.caretSize || 20) + 'px</span>' +
      '<span class="caret-demo"><button class="perf-caret">\u25B8</button><button class="perf-caret open">\u25B8</button></span>' +
    '</div>' +
    '<div class="settings-row">' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="set-zebra"' + (Theme.zebra ? ' checked' : '') + '> Alternate row shading</label>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="set-grid"' + (Theme.gridLines ? ' checked' : '') + '> Row separator lines</label>' +
    '</div>' +

    '<h3 class="snap-set-title">Table grid lines</h3>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Line colour</label>' +
      '<input type="color" id="set-gridcolor" value="' + (Theme.gridColor || '#111111') + '">' +
      '<span class="hexcode" id="set-gridcolor-val">' + (Theme.gridColor || '#111111') + '</span>' +
      '<button class="ghost-btn small" id="set-grid-black">Black</button>' +
      '<button class="ghost-btn small" id="set-grid-soft">Soft grey</button>' +
    '</div>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">Line thickness</label>' +
      '<input type="range" id="set-gridwidth" min="0" max="3" step="0.5" value="' + (Theme.gridWidth === undefined ? 1 : Theme.gridWidth) + '">' +
      '<span class="drill-count" id="set-gridwidth-val">' + (Theme.gridWidth === undefined ? 1 : Theme.gridWidth) + 'px</span>' +
    '</div>' +

    '<h3 class="snap-set-title">Row colours</h3>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Mouse-over row</label>' +
      '<input type="color" id="set-hovercolor" value="' + (Theme.hoverColor || '#FFE3BF') + '">' +
      '<span class="hexcode" id="set-hovercolor-val">' + (Theme.hoverColor || '#FFE3BF') + '</span>' +
    '</div>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Mouse-over drill row</label>' +
      '<input type="color" id="set-drillhover" value="' + (Theme.drillHoverColor || '#D7E8F7') + '">' +
      '<span class="hexcode" id="set-drillhover-val">' + (Theme.drillHoverColor || '#D7E8F7') + '</span>' +
      '<span class="drill-count">kept different so drill rows stand out</span>' +
    '</div>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Opened drill row fill</label>' +
      '<input type="color" id="set-openfill" value="' + (Theme.openFill || '#FDF3E7') + '">' +
      '<span class="hexcode" id="set-openfill-val">' + (Theme.openFill || '#FDF3E7') + '</span>' +
    '</div>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Opened drill row border</label>' +
      '<input type="color" id="set-openborder" value="' + (Theme.openBorderColor || '#A6402C') + '">' +
      '<span class="hexcode" id="set-openborder-val">' + (Theme.openBorderColor || '#A6402C') + '</span>' +
    '</div>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">Open row border thickness</label>' +
      '<input type="range" id="set-openbw" min="0" max="5" step="0.5" value="' + (Theme.openBorderWidth === undefined ? 2 : Theme.openBorderWidth) + '">' +
      '<span class="drill-count" id="set-openbw-val">' + (Theme.openBorderWidth === undefined ? 2 : Theme.openBorderWidth) + 'px</span>' +
    '</div>' +

    '<div class="theme-preview">' +
      '<table class="data-table"><thead><tr><th>Article</th><th class="num">Sold</th><th class="num">Stock</th></tr></thead>' +
      '<tbody><tr><td>T-SHIRT-M</td><td class="num">1,538</td><td class="num">2,828</td></tr>' +
      '<tr><td>SHIRT-M</td><td class="num">656</td><td class="num">1,204</td></tr>' +
      '<tr><td>TOP-L</td><td class="num">458</td><td class="num">903</td></tr></tbody></table>' +
    '</div>';

  wrap.querySelectorAll('.swatch').forEach(b => b.addEventListener('click', () => {
    Theme.accent = b.dataset.accent; saveTheme(); renderSettingsBody();
  }));
  wrap.querySelectorAll('input[name="themefont"]').forEach(r => r.addEventListener('change', () => {
    Theme.font = r.value; saveTheme(); renderSettingsBody();
  }));
  const rp = wrap.querySelector('#set-rowpad');
  if (rp) rp.addEventListener('input', e => {
    Theme.rowPad = parseInt(e.target.value, 10);
    wrap.querySelector('#set-rowpad-val').textContent = Theme.rowPad + 'px';
    saveTheme();
  });
  const fs = wrap.querySelector('#set-fontsize');
  fs.addEventListener('input', e => {
    Theme.fontSize = parseFloat(e.target.value);
    wrap.querySelector('#set-fontsize-val').textContent = Theme.fontSize + 'px';
    saveTheme();
  });
  function bindColor(id, key) {
    const el = wrap.querySelector('#' + id);
    if (!el) return;
    el.addEventListener('input', function (e) {
      Theme[key] = e.target.value;
      const lab = wrap.querySelector('#' + id + '-val');
      if (lab) lab.textContent = e.target.value;
      saveTheme();
    });
  }
  bindColor('set-gridcolor', 'gridColor');
  bindColor('set-hovercolor', 'hoverColor');
  bindColor('set-drillhover', 'drillHoverColor');
  bindColor('set-openfill', 'openFill');
  bindColor('set-openborder', 'openBorderColor');

  const gb = wrap.querySelector('#set-grid-black');
  if (gb) gb.addEventListener('click', function () { Theme.gridColor = '#111111'; saveTheme(); renderSettingsBody(); });
  const gs = wrap.querySelector('#set-grid-soft');
  if (gs) gs.addEventListener('click', function () { Theme.gridColor = '#D8D3C6'; saveTheme(); renderSettingsBody(); });

  const gw = wrap.querySelector('#set-gridwidth');
  if (gw) gw.addEventListener('input', function (e) {
    Theme.gridWidth = parseFloat(e.target.value);
    wrap.querySelector('#set-gridwidth-val').textContent = Theme.gridWidth + 'px';
    saveTheme();
  });
  const obw = wrap.querySelector('#set-openbw');
  if (obw) obw.addEventListener('input', function (e) {
    Theme.openBorderWidth = parseFloat(e.target.value);
    wrap.querySelector('#set-openbw-val').textContent = Theme.openBorderWidth + 'px';
    saveTheme();
  });

  const crEl = wrap.querySelector('#set-childrow');
  if (crEl) crEl.addEventListener('input', function (e) {
    Theme.childRowScale = parseInt(e.target.value, 10);
    wrap.querySelector('#set-childrow-val').textContent = Theme.childRowScale + '% of the parent row';
    saveTheme();
  });
  const cs = wrap.querySelector('#set-caret');
  if (cs) cs.addEventListener('input', e => {
    Theme.caretSize = parseInt(e.target.value, 10);
    wrap.querySelector('#set-caret-val').textContent = Theme.caretSize + 'px';
    saveTheme();
  });
  wrap.querySelector('#set-zebra').addEventListener('change', e => { Theme.zebra = e.target.checked; saveTheme(); });
  wrap.querySelector('#set-grid').addEventListener('change', e => { Theme.gridLines = e.target.checked; saveTheme(); });
  finish();
}

/* ---- Per-tab preferences (Settings panel, sections 01-06) ---- */
const PERIOD_CHOICES = [
  ['all', 'All data'], ['30', 'Last 30 days'], ['90', 'Last 90 days'],
  ['180', 'Last 180 days'], ['365', 'Last 365 days'],
  ['thismonth', 'Latest month'], ['thisyear', 'Latest year']
];

const PREFS_DEFAULT = {
  dashGrain: 'month',
  defaultPeriod: 'all',
  perfGroupBy: 'Article No',
  perfLimit: 800,
  reorderGroupBy: 'Article No',
  targetDays: 30,
  reorderOnlyFlagged: true,
  pivotMode: 'quick',
  quickGrain: 'none',
  quickDesc: true,
  explorePageSize: 100,
  autoDetectLinks: true
};

const Prefs = Object.assign({}, PREFS_DEFAULT);

function loadPrefs() {
  try {
    const raw = Store.get('sl_prefs');
    if (raw) Object.assign(Prefs, PREFS_DEFAULT, JSON.parse(raw));
  } catch (e) {}
}
function savePrefs() { Store.set('sl_prefs', JSON.stringify(Prefs)); }

/** Saved defaults ko controls par lagata hai (page khulte hi). */
function applyPrefsToControls() {
  DashState.grain = Prefs.dashGrain || 'month';
  document.querySelectorAll('#dash-grain .seg-btn').forEach(function (b) {
    b.classList.toggle('active', b.dataset.grain === DashState.grain);
  });

  if (Prefs.defaultPeriod && Prefs.defaultPeriod !== 'all') {
    App.period.mode = Prefs.defaultPeriod;
    document.querySelectorAll('.period-select').forEach(function (sl) { sl.value = Prefs.defaultPeriod; });
  }

  const pg = document.getElementById('perf-groupby');
  if (pg && Prefs.perfGroupBy) pg.value = Prefs.perfGroupBy;

  const rg = document.getElementById('insights-groupby');
  if (rg && Prefs.reorderGroupBy) rg.value = Prefs.reorderGroupBy;
  const td = document.getElementById('insights-target-days');
  if (td && Prefs.targetDays) td.value = Prefs.targetDays;
  const of = document.getElementById('insights-only-flagged');
  if (of) of.checked = Prefs.reorderOnlyFlagged !== false;

  QuickReport.dateGrain = Prefs.quickGrain || 'none';
  const qg = document.getElementById('quick-date-grain');
  if (qg) qg.value = QuickReport.dateGrain;
  QuickReport.desc = Prefs.quickDesc !== false;
  const qs = document.getElementById('quick-sort');
  if (qs) qs.value = QuickReport.desc ? 'desc' : 'asc';

  ExploreState.pageSize = Prefs.explorePageSize || 100;

  if (Prefs.pivotMode === 'drag') {
    const b = document.querySelector('#pivot-mode .seg-btn[data-mode="drag"]');
    if (b) b.click();
  }
}

/* Behaviour settings */
const Behaviour = { inlineExpand: true, perfChain: PERF_CHAIN_DEFAULT.slice(), autoOpenFirst: false };
function loadBehaviour() {
  try { const raw = Store.get('sl_behaviour'); if (raw) Object.assign(Behaviour, JSON.parse(raw)); } catch (e) {}
}
function saveBehaviour() { Store.set('sl_behaviour', JSON.stringify(Behaviour)); }

/* ---------------------------------------------------------------
   14. COLUMN RESIZING — drag a header edge to widen/narrow a column
   --------------------------------------------------------------- */
const ColWidths = { map: {}, drag: null };

function loadColWidths() {
  try {
    const raw = Store.get('sl_colwidths');
    if (raw) ColWidths.map = JSON.parse(raw) || {};
  } catch (e) { ColWidths.map = {}; }
}
function saveColWidths() { Store.set('sl_colwidths', JSON.stringify(ColWidths.map)); }

function colKey(tableId, index) { return tableId + '#' + index; }

/** Applies saved widths and adds a drag handle to every header cell. */
function makeTableResizable(table) {
  if (!table) return;
  const tableId = table.id || 'tbl';
  const ths = table.querySelectorAll('thead th');
  if (!ths.length) return;

  table.classList.add('resizable-table');

  ths.forEach((th, i) => {
    const saved = ColWidths.map[colKey(tableId, i)];
    if (saved) {
      th.style.width = saved + 'px';
      th.style.minWidth = saved + 'px';
      th.style.maxWidth = saved + 'px';
    }
    if (th.querySelector('.col-resizer')) return;
    const grip = document.createElement('span');
    grip.className = 'col-resizer';
    grip.title = 'Drag to resize this column';
    th.appendChild(grip);

    grip.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation();           // don't trigger the header's sort handler
      ColWidths.drag = {
        tableId: tableId, index: i, th: th,
        startX: e.clientX,
        startW: th.getBoundingClientRect().width
      };
      document.body.classList.add('col-resizing');
    });
    // double-click resets this column to automatic width
    grip.addEventListener('dblclick', e => {
      e.preventDefault(); e.stopPropagation();
      delete ColWidths.map[colKey(tableId, i)];
      th.style.width = th.style.minWidth = th.style.maxWidth = '';
      saveColWidths();
    });
  });
}

function initColResize() {
  loadColWidths();
  window.addEventListener('mousemove', e => {
    const d = ColWidths.drag;
    if (!d) return;
    const w = Math.max(48, Math.round(d.startW + (e.clientX - d.startX)));
    d.th.style.width = w + 'px';
    d.th.style.minWidth = w + 'px';
    d.th.style.maxWidth = w + 'px';
  });
  window.addEventListener('mouseup', () => {
    const d = ColWidths.drag;
    if (!d) return;
    const w = parseInt(d.th.style.width, 10);
    if (w) { ColWidths.map[colKey(d.tableId, d.index)] = w; saveColWidths(); }
    ColWidths.drag = null;
    document.body.classList.remove('col-resizing');
  });
}

/** Called after any table re-render so grips and widths stay in place. */
function refreshResizableTables() {
  ['perf-table', 'insights-table', 'quick-table', 'explore-table', 'drill-breakdown', 'drill-raw-table']
    .forEach(id => makeTableResizable(document.getElementById(id)));
}

/* ---------------------------------------------------------------
   15. CATALOG CONSOLE — alphabetical product browser
   ---------------------------------------------------------------
   Article No  = the design      (top level row)
   Colour Name = the variant     (first expand)
   Size Name   = the size        (second expand)
   Every level shows sold / opening / closing, cover days and a
   colour-coded status, plus an optional photo per design.
   --------------------------------------------------------------- */

const Catalog = {
  letter: 'all',
  section: 'all',
  sub: 'all',
  search: '',
  flags: { stockout: false, nosale: false, overstock: false, lowcover: false },
  sort: 'worstcover',
  sortCol: null,        // clicked column key
  sortDir: -1,
  expanded: {},
  images: {},          // Article No -> data URL
  lastRows: null
};

/* ---- colour swatches: map a colour name to something visual ---- */
const COLOUR_WORDS = {
  black: '#1B1B1B', white: '#F5F5F5', offwhite: '#F2EFE6', 'off white': '#F2EFE6',
  cream: '#F3E9D2', ivory: '#F6F1E0', beige: '#DCC9A6', khaki: '#B5A16B',
  grey: '#8B8B8B', gray: '#8B8B8B', charcoal: '#3A3F44', silver: '#C0C0C0',
  red: '#C0392B', maroon: '#7B241C', rust: '#A6402C', wine: '#722F37',
  pink: '#E88BA6', rose: '#D96A87', peach: '#F2B49A', coral: '#E9705B',
  orange: '#E08A2E', mustard: '#D3A625', yellow: '#E8C547', gold: '#C9A227',
  green: '#2E7D4F', olive: '#6B7A32', pista: '#9CCB8E', mint: '#A8D8C4',
  teal: '#1F6F5C', bottle: '#1F4B3F', 'b-green': '#2E7D4F',
  blue: '#2C5AA0', navy: '#1E2A4A', 'r-blue': '#2C5AA0', sky: '#7FB3E0',
  firozi: '#28B5B5', turquoise: '#28B5B5', 'air force': '#5A7FA6',
  purple: '#6B4C9A', violet: '#7A4CA0', lavender: '#B9A7D6', magenta: '#B5348C',
  brown: '#6B4A2F', coffee: '#4B3621', tan: '#B08D57', camel: '#C19A6B',
  multi: 'linear-gradient(135deg,#C0392B 0 25%,#E8C547 25% 50%,#2E7D4F 50% 75%,#2C5AA0 75%)',
  assorted: 'linear-gradient(135deg,#C0392B 0 25%,#E8C547 25% 50%,#2E7D4F 50% 75%,#2C5AA0 75%)',
  asstd: 'linear-gradient(135deg,#C0392B 0 25%,#E8C547 25% 50%,#2E7D4F 50% 75%,#2C5AA0 75%)'
};

/** Turns a colour name into a CSS background. Known names get their real
 *  colour; anything else gets a stable colour derived from the text, so the
 *  same name always looks the same. */
function colourSwatch(name) {
  const key = String(name || '').trim().toLowerCase();
  if (!key || key === '(blank)') return '#D8D3C6';
  if (COLOUR_WORDS[key]) return COLOUR_WORDS[key];
  for (const w in COLOUR_WORDS) {
    if (key.indexOf(w) !== -1) return COLOUR_WORDS[w];
  }
  let h = 0;
  for (let i = 0; i < key.length; i++) h = (h * 31 + key.charCodeAt(i)) % 360;
  return 'hsl(' + h + ', 42%, 58%)';
}

/* ---- build the catalog index from the loaded data ---- */
function buildCatalog() {
  const range = periodRange();
  const sales = salesRecords().filter(r => inPeriod(r, range));
  const purch = purchaseRecords().filter(r => inPeriod(r, range));
  const stock = stockRecords();
  const days = periodDayCount(range, sales.length ? sales : purch);
  const anchor = dataAnchorDate();

  const designs = new Map();
  function design(key) {
    let d = designs.get(key);
    if (!d) {
      d = { key, sold: 0, purchased: 0, obs: 0, cbs: 0, hasOBS: false,
            lastSale: null, colours: new Map(), meta: {} };
      designs.set(key, d);
    }
    return d;
  }
  function colour(d, key) {
    let c = d.colours.get(key);
    if (!c) { c = { key, sold: 0, purchased: 0, obs: 0, cbs: 0, hasOBS: false, lastSale: null, sizes: new Map() }; d.colours.set(key, c); }
    return c;
  }
  function size(c, key) {
    let z = c.sizes.get(key);
    if (!z) { z = { key, sold: 0, purchased: 0, obs: 0, cbs: 0, hasOBS: false }; c.sizes.set(key, z); }
    return z;
  }
  function meta(d, r) {
    ['Section', 'Sub Section', 'Brand', 'Supplier', 'Style'].forEach(f => {
      if (!d.meta[f] && r[f]) d.meta[f] = String(r[f]);
    });
  }

  sales.forEach(r => {
    const d = design(dimKey(r, 'Article No'));
    const c = colour(d, dimKey(r, 'Colour'));
    const z = size(c, dimKey(r, 'Size'));
    const q = recQty(r);
    d.sold += q; c.sold += q; z.sold += q;
    if (r.Date) {
      if (!d.lastSale || r.Date > d.lastSale) d.lastSale = r.Date;
      if (!c.lastSale || r.Date > c.lastSale) c.lastSale = r.Date;
    }
    meta(d, r);
  });

  purch.forEach(r => {
    const d = design(dimKey(r, 'Article No'));
    const c = colour(d, dimKey(r, 'Colour'));
    const z = size(c, dimKey(r, 'Size'));
    const q = recQty(r);
    d.purchased += q; c.purchased += q; z.purchased += q;
    meta(d, r);
  });

  stock.forEach(r => {
    const d = design(dimKey(r, 'Article No'));
    const c = colour(d, dimKey(r, 'Colour'));
    const z = size(c, dimKey(r, 'Size'));
    const cb = recQty(r);
    const ob = recOpeningQty(r);
    d.cbs += cb; c.cbs += cb; z.cbs += cb;
    if (ob !== null) {
      d.obs += ob; c.obs += ob; z.obs += ob;
      d.hasOBS = c.hasOBS = z.hasOBS = true;
    }
    meta(d, r);
  });

  const rows = [...designs.values()].map(d => {
    d.colourList = [...d.colours.values()].map(c => {
      c.sizeList = [...c.sizes.values()].sort((a, b) => b.cbs - a.cbs || b.sold - a.sold);
      return Object.assign(c, catalogMetrics(c, days, anchor));
    }).sort((a, b) => b.sold - a.sold || b.cbs - a.cbs);
    return Object.assign(d, catalogMetrics(d, days, anchor));
  });

  return { rows, days, range, anchor };
}

function catalogMetrics(x, days, anchor) {
  const avgDaily = x.sold / days;
  const cover = avgDaily > 0 ? x.cbs / avgDaily : (x.cbs > 0 ? Infinity : 0);
  const opening = x.sold + x.cbs;
  const sellThrough = opening > 0 ? (x.sold / opening) * 100 : 0;
  const daysSince = x.lastSale ? Math.round((anchor - x.lastSale) / 86400000) : null;
  const moved = x.hasOBS ? (x.obs - x.cbs) : null;
  let status;
  if (x.cbs === 0 && x.sold > 0) status = 'Stockout';
  else if (x.sold === 0 && x.cbs > 0) status = 'No sale';
  else if (cover !== Infinity && cover < (CatPrefs.lowCoverDays || 15) && x.sold > 0) status = 'Low cover';
  else if (cover === Infinity || cover > (CatPrefs.overstockDays || 120)) status = 'Overstock';
  else if (x.sold > 0) status = 'Healthy';
  else status = 'Idle';
  return { avgDaily, cover, sellThrough, daysSince, moved, status };
}

function catalogStatusClass(s) {
  return 'cs-' + String(s).toLowerCase().replace(/\s+/g, '-');
}

/* ---- filtering ---- */
function catalogFiltered(all) {
  let rows = all.slice();
  if (Catalog.letter !== 'all') {
    rows = rows.filter(r => String(r.key).trim().charAt(0).toUpperCase() === Catalog.letter);
  }
  if (Catalog.section !== 'all') rows = rows.filter(r => (r.meta.Section || '(blank)') === Catalog.section);
  if (Catalog.sub !== 'all') rows = rows.filter(r => (r.meta['Sub Section'] || '(blank)') === Catalog.sub);

  const f = Catalog.flags;
  if (f.stockout) rows = rows.filter(r => r.status === 'Stockout');
  if (f.nosale) rows = rows.filter(r => r.sold === 0);
  if (f.overstock) rows = rows.filter(r => r.status === 'Overstock');
  if (f.lowcover) rows = rows.filter(r => r.status === 'Low cover');

  const q = Catalog.search.trim().toLowerCase();
  if (q) {
    rows = rows.filter(r =>
      r.key.toLowerCase().indexOf(q) !== -1 ||
      Object.values(r.meta).some(v => String(v).toLowerCase().indexOf(q) !== -1) ||
      [...r.colours.keys()].some(c => String(c).toLowerCase().indexOf(q) !== -1));
  }

  if (Catalog.sortCol) {
    const c = Catalog.sortCol, dir = Catalog.sortDir;
    rows.sort((a, b) => {
      const va = catSortValue(a, c), vb = catSortValue(b, c);
      if (typeof va === 'string' || typeof vb === 'string') return String(va).localeCompare(String(vb)) * dir;
      return ((va || 0) - (vb || 0)) * dir;
    });
    return rows;
  }

  const s = Catalog.sort;
  rows.sort((a, b) => {
    if (s === 'topseller') return b.sold - a.sold;
    if (s === 'moststock') return b.cbs - a.cbs;
    if (s === 'az') return String(a.key).localeCompare(String(b.key));
    // worst cover first: stockouts, then lowest cover, ignoring idle lines
    const av = a.cover === Infinity ? 1e9 : a.cover;
    const bv = b.cover === Infinity ? 1e9 : b.cover;
    return av - bv || b.sold - a.sold;
  });
  return rows;
}

/* ---- rendering ---- */
/* ---- drag the bar under the Catalog table to make it taller or shorter ---- */
function initCatalogHeightGrip() {
  const grip = document.getElementById('cat-height-grip');
  const box = document.getElementById('cat-scroll-box');
  if (!grip || !box) return;

  const saved = parseInt(Store.get('sl_cat_height') || '', 10);
  if (saved > 120) { box.style.flex = '0 0 auto'; box.style.height = saved + 'px'; }

  grip.addEventListener('mousedown', e => {
    e.preventDefault();
    const startY = e.clientY;
    const startH = box.getBoundingClientRect().height;
    box.style.flex = '0 0 auto';
    document.body.classList.add('row-resizing');

    const move = ev => {
      const h = Math.max(140, startH + (ev.clientY - startY));
      box.style.height = h + 'px';
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      document.body.classList.remove('row-resizing');
      Store.set('sl_cat_height', String(Math.round(box.getBoundingClientRect().height)));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  });

  // double-click puts it back to filling the tab
  grip.addEventListener('dblclick', () => {
    box.style.flex = ''; box.style.height = '';
    Store.remove('sl_cat_height');
    toast('Table height reset to fit the window.');
  });
}

function initCatalog() {
  const host = document.getElementById('tab-catalog');
  if (!host) return;
  loadCatalogImages();
  initCatalogHeightGrip();

  document.getElementById('cat-search').addEventListener('input', debounce(e => {
    Catalog.search = e.target.value; renderCatalog();
  }, 200));
  document.getElementById('cat-sort').addEventListener('change', e => {
    Catalog.sort = e.target.value; Catalog.sortCol = null; renderCatalog();
  });
  document.getElementById('cat-expand-all').addEventListener('click', () => {
    const rows = catalogFiltered(buildCatalog().rows).slice(0, 60);
    rows.forEach(r => { Catalog.expanded[r.key] = true; });
    renderCatalog();
  });
  document.getElementById('cat-collapse-all').addEventListener('click', () => {
    Catalog.expanded = {}; renderCatalog();
  });
  document.getElementById('cat-export').addEventListener('click', exportCatalogCSV);
  document.querySelectorAll('#cat-flags .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const k = btn.dataset.flag;
      Catalog.flags[k] = !Catalog.flags[k];
      btn.classList.toggle('active', Catalog.flags[k]);
      renderCatalog();
    });
  });
  const imgInput = document.getElementById('cat-image-input');
  if (imgInput) imgInput.addEventListener('change', onCatalogImagePicked);
}

/** Colour key + the two formulas, shown above the catalog table. */
function replenLegendHtml() {
  const bands = [
    ['0\u201333%', '#ea9999', 'Low stock'],
    ['>33\u201366%', '#ffd966', 'Medium stock'],
    ['>66\u2013100%', '#b6d7a8', 'Good / healthy'],
    ['>100%', '#b4a7d6', 'Overstock']
  ];
  return '<div class="stock-legend"><span class="sl-title">Stock %</span>' +
    bands.map(b => '<span class="sl-item"><span class="sl-sw" style="background:' + b[1] + '"></span>' +
      b[0] + ' \u00b7 ' + b[2] + '</span>').join('') +
    '<span class="sl-note">ML = ADC \u00d7 LT \u00d7 SF \u00b7 Stock % = (Closing + MIT) \u00f7 ML</span></div>';
}

/** Sets one field on every design currently listed. */
function replenBulkApply(field, value) {
  const rows = Catalog.lastRows || [];
  if (!rows.length) { toast('Nothing on screen to apply to.'); return; }
  const v = parseFloat(value);
  if (isNaN(v)) { toast('Enter a number first.'); return; }
  rows.forEach(d => {
    const o = Replen.overrides[d.key] || (Replen.overrides[d.key] = {});
    o[field] = v;
  });
  saveReplen();
  renderCatalog();
  toast(field.toUpperCase() + ' set to ' + v + ' on ' + rows.length + ' designs.');
}

function replenBulkBarHtml() {
  return '<div class="cat-bulk">' +
    '<span class="toolbar-label">Set for all listed designs:</span>' +
    '<label class="toolbar-label">ADC</label><input type="number" id="bulk-adc" class="text-input" min="0" step="0.01" value="' + (Replen.bulk.adc !== undefined ? Replen.bulk.adc : '') + '">' +
    '<label class="toolbar-label">LT</label><input type="number" id="bulk-lt" class="text-input" min="0" step="1" value="' + (Replen.bulk.lt !== undefined ? Replen.bulk.lt : '') + '">' +
    '<label class="toolbar-label">SF</label><input type="number" id="bulk-sf" class="text-input" min="0" step="0.1" value="' + (Replen.bulk.sf !== undefined ? Replen.bulk.sf : '') + '">' +
    '<label class="toolbar-label">MOQ</label><input type="number" id="bulk-moq" class="text-input" min="0" step="1" value="' + (Replen.bulk.moq !== undefined ? Replen.bulk.moq : '') + '">' +
    '<button class="ghost-btn small primary" id="bulk-apply">Apply</button>' +
    '<button class="ghost-btn small" id="bulk-clear">Clear all manual values</button>' +
    '</div>';
}

function wireReplenBulk() {
  const ap = document.getElementById('bulk-apply');
  if (ap) ap.addEventListener('click', () => {
    // Read every box FIRST. Applying one at a time re-drew the bar and wiped
    // the boxes that had not been read yet.
    const picked = [];
    [['adc', 'bulk-adc'], ['lt', 'bulk-lt'], ['sf', 'bulk-sf'], ['moq', 'bulk-moq']].forEach(([f, id]) => {
      const el = document.getElementById(id);
      if (!el) return;
      if (el.value !== '' && !isNaN(parseFloat(el.value))) {
        const v = parseFloat(el.value);
        Replen.bulk[f] = v;
        picked.push([f, v]);
      } else delete Replen.bulk[f];
    });

    if (!picked.length) { saveReplenBulk(); toast('Fill ADC, LT, SF or MOQ first.'); return; }

    const rows = Catalog.lastRows || [];
    if (!rows.length) { toast('Nothing on screen to apply to.'); return; }
    rows.forEach(d => {
      const o = Replen.overrides[d.key] || (Replen.overrides[d.key] = {});
      picked.forEach(([f, v]) => { o[f] = v; });
    });
    saveReplen();
    saveReplenBulk();
    renderCatalog();
    toast(picked.map(p => p[0].toUpperCase() + ' ' + p[1]).join(', ') + ' set on ' + rows.length + ' designs.');
  });
  const cl = document.getElementById('bulk-clear');
  if (cl) cl.addEventListener('click', () => {
    Replen.overrides = {}; Replen.bulk = {};
    saveReplen(); saveReplenBulk(); renderCatalog();
    toast('All manual ADC / LT / SF / MOQ / MIT values cleared.');
  });
}

function renderCatalog() {
  const host = document.getElementById('cat-table');
  if (!host) return;
  if (!App.datasets.length) {
    document.getElementById('cat-filters').innerHTML = '';
    document.getElementById('cat-count').textContent = '';
    host.innerHTML = '<tr><td class="empty-hint">Load your reports on the Import tab first.</td></tr>';
    return;
  }

  const built = buildCatalog();
  _catDays = built.days;
  renderCatalogFilters(built.rows);
  const extra = document.getElementById('cat-extra');
  if (extra) {
    const planOn = ['adc','lt','sf','moq','ml','mit','stockpct','reorder'].some(catColOn);
    extra.innerHTML = planOn ? (replenLegendHtml() + replenBulkBarHtml()) : '';
    if (planOn) wireReplenBulk();
  }
  const rows = catalogFiltered(built.rows);
  Catalog.lastRows = rows;

  document.getElementById('cat-count').textContent =
    rows.length.toLocaleString('en-IN') + ' / ' + built.rows.length.toLocaleString('en-IN') + ' designs' +
    ' \u00b7 ' + built.days + ' days';

  const head = '<thead><tr>' +
    '<th class="cat-c-design" data-sc="key">Design' + catSortArrow('key') + '</th>' +
    (catColOn('category') ? '<th data-sc="category">Category' + catSortArrow('category') + '</th>' : '') +
    (catColOn('colours') ? '<th class="cat-c-colours" data-sc="colours">Colours' + catSortArrow('colours') + '</th>' : '') +
    (catColOn('sold') ? '<th class="num" data-sc="sold">Sold' + catSortArrow('sold') + '</th>' : '') +
    (catColOn('purchased') ? '<th class="num" data-sc="purchased" title="Quantity received in this window">Purchased' + catSortArrow('purchased') + '</th>' : '') +
    (catColOn('opening') ? '<th class="num" data-sc="obs">Opening' + catSortArrow('obs') + '</th>' : '') +
    (catColOn('closing') ? '<th class="num" data-sc="cbs">Closing' + catSortArrow('cbs') + '</th>' : '') +
    (catColOn('moved') ? '<th class="num" data-sc="moved">Moved' + catSortArrow('moved') + '</th>' : '') +
    (catColOn('adc') ? '<th class="num" title="Average Daily Consumption" data-sc="adc">ADC' + catSortArrow('adc') + '</th>' : '') +
    (catColOn('lt') ? '<th class="num" title="Lead Time in days" data-sc="lt">LT' + catSortArrow('lt') + '</th>' : '') +
    (catColOn('sf') ? '<th class="num" title="Safety Factor" data-sc="sf">SF' + catSortArrow('sf') + '</th>' : '') +
    (catColOn('moq') ? '<th class="num" title="Minimum Order Quantity" data-sc="moq">MOQ' + catSortArrow('moq') + '</th>' : '') +
    (catColOn('ml') ? '<th class="num" title="Max Level = ADC x LT x SF" data-sc="ml">ML' + catSortArrow('ml') + '</th>' : '') +
    (catColOn('mit') ? '<th class="num" title="Material In Transit" data-sc="mit">MIT' + catSortArrow('mit') + '</th>' : '') +
    (catColOn('stockpct') ? '<th class="num" title="(Closing + MIT) as a share of Max Level" data-sc="pct">Stock %' + catSortArrow('pct') + '</th>' : '') +
    (catColOn('reorder') ? '<th class="num" title="ML minus what you have, rounded up to the MOQ" data-sc="reorder">Reorder' + catSortArrow('reorder') + '</th>' : '') +
    (catColOn('cover') ? '<th class="num" data-sc="cover">Cover' + catSortArrow('cover') + '</th>' : '') +
    (catColOn('sellthru') ? '<th class="num" data-sc="sellThrough">Sell-thru' + catSortArrow('sellThrough') + '</th>' : '') +
    (catColOn('lastsold') ? '<th data-sc="lastSale">Last sold' + catSortArrow('lastSale') + '</th>' : '') +
    (catColOn('status') ? '<th data-sc="status">Status' + catSortArrow('status') + '</th>' : '') +
    '</tr></thead>';

  const body = rows.slice(0, CatPrefs.maxRows || 300).map(d => catalogDesignRow(d)).join('');

  const tSold = rows.reduce((a, r) => a + r.sold, 0);
  const tPurch = rows.reduce((a, r) => a + r.purchased, 0);
  const tObs = rows.reduce((a, r) => a + r.obs, 0);
  const tCbs = rows.reduce((a, r) => a + r.cbs, 0);
  const tST = (tSold + tCbs) > 0 ? (tSold / (tSold + tCbs)) * 100 : 0;
  const tReorder = rows.reduce((a, r) => a + replenFor(r.key, r.sold, built.days, r.cbs).reorder, 0);
  const fillerCols = (catColOn('category') ? 1 : 0) + (catColOn('colours') ? 1 : 0);
  const foot = '<tfoot><tr>' +
    '<td>Total \u00b7 ' + rows.length.toLocaleString('en-IN') + ' designs</td>' +
    (fillerCols ? '<td colspan="' + fillerCols + '"></td>' : '') +
    (catColOn('sold') ? '<td class="num">' + fmtNum(tSold) + '</td>' : '') +
    (catColOn('purchased') ? '<td class="num">' + fmtNum(tPurch) + '</td>' : '') +
    (catColOn('opening') ? '<td class="num">' + fmtNum(tObs) + '</td>' : '') +
    (catColOn('closing') ? '<td class="num">' + fmtNum(tCbs) + '</td>' : '') +
    (catColOn('moved') ? '<td class="num">' + fmtNum(tObs - tCbs) + '</td>' : '') +
    ['adc','lt','sf','moq','ml','mit','stockpct'].filter(catColOn).map(function () { return '<td></td>'; }).join('') +
    (catColOn('reorder') ? '<td class="num">' + fmtNum(tReorder) + '</td>' : '') +
    (catColOn('cover') ? '<td class="num"></td>' : '') +
    (catColOn('sellthru') ? '<td class="num">' + fmtNum(tST, 1) + '%</td>' : '') +
    (catColOn('lastsold') ? '<td></td>' : '') +
    (catColOn('status') ? '<td></td>' : '') +
    '</tr></tfoot>';

  host.innerHTML = head + '<tbody>' + body + '</tbody>' + foot;
  host.querySelectorAll('thead th[data-sc]').forEach(th => th.addEventListener('click', () => {
    const c = th.dataset.sc;
    if (Catalog.sortCol === c) Catalog.sortDir *= -1;
    else { Catalog.sortCol = c; Catalog.sortDir = (c === 'key' || c === 'category' || c === 'status') ? 1 : -1; }
    renderCatalog();
  }));
  wireCatalogRows();
  wireReplenInputs(host);
  makeTableResizable(host);
}

let _catDays = 30;
function catalogDays() { return _catDays; }

/** Typing in an ADC / LT / SF / MOQ / MIT box saves that value for the row.
 *  Clearing the box goes back to the automatic figure. */
function wireReplenInputs(host) {
  host.querySelectorAll('.cat-inp').forEach(inp => {
    inp.addEventListener('click', e => e.stopPropagation());
    inp.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') e.target.blur(); });
    inp.addEventListener('change', e => {
      e.stopPropagation();
      const key = e.target.dataset.rkey, field = e.target.dataset.rfield;
      const raw = e.target.value;
      const o = Replen.overrides[key] || (Replen.overrides[key] = {});
      if (raw === '' || isNaN(parseFloat(raw))) delete o[field];
      else o[field] = parseFloat(raw);
      if (!Object.keys(o).length) delete Replen.overrides[key];
      saveReplen();
      renderCatalog();
    });
  });
}

/** Little arrow on the header cell that is being sorted. */
function catSortArrow(col) {
  if (Catalog.sortCol !== col) return '';
  return '<span class="sort-arrow">' + (Catalog.sortDir === 1 ? '\u25B2' : '\u25BC') + '</span>';
}

/** Value used when sorting by a clicked column. */
function catSortValue(d, col) {
  if (col === 'key') return String(d.key).toLowerCase();
  if (col === 'category') return String(d.meta['Sub Section'] || d.meta.Section || '').toLowerCase();
  if (col === 'colours') return d.colourList.length;
  if (col === 'obs') return d.obs;
  if (col === 'cbs') return d.cbs;
  if (col === 'moved') return d.moved === null ? -Infinity : d.moved;
  if (col === 'cover') return d.cover === Infinity ? 1e12 : d.cover;
  if (col === 'lastSale') return d.lastSale ? d.lastSale.getTime() : -Infinity;
  if (col === 'status') return String(d.status).toLowerCase();
  if (['adc', 'lt', 'sf', 'moq', 'ml', 'mit', 'pct', 'reorder'].indexOf(col) !== -1) {
    const r = replenFor(d.key, d.sold, catalogDays(), d.cbs);
    return col === 'pct' ? (isFinite(r.pct) ? r.pct : 1e12) : r[col];
  }
  return d[col];
}

function catalogDesignRow(d) {
  const open = !!Catalog.expanded[d.key];
  const img = Catalog.images[d.key];
  const maxDots = CatPrefs.maxDots || 10;
  const dots = d.colourList.slice(0, maxDots).map(c =>
    '<span class="cat-dotwrap">' +
      '<span class="cat-dot" title="' + escapeHtml(c.key + ' \u2014 ' + fmtNum(c.cbs) + ' in stock, ' + fmtNum(c.sold) + ' sold') + '" ' +
      'style="background:' + colourSwatch(c.key) + '"></span>' +
      (CatPrefs.showDotCounts ? '<span class="cat-dotn">' + fmtNum(c.cbs) + '</span>' : '') +
    '</span>').join('') +
    (d.colourList.length > maxDots ? '<span class="cat-more">+' + (d.colourList.length - maxDots) + '</span>' : '');

  let html = '<tr class="cat-row cat-design' + (open ? ' is-open' : '') + '" data-key="' + escapeHtml(d.key) + '">' +
    '<td class="cat-c-design">' +
      '<button class="cat-caret' + (open ? ' open' : '') + '">\u25B8</button>' +
      (CatPrefs.showThumbs
        ? '<button class="cat-thumb" data-img="' + escapeHtml(d.key) + '" title="' +
          (img ? 'Click to view, replace or remove the photo' : 'Click to add a photo') + '">' +
          (img ? '<img src="' + img + '" alt="">' : '<span class="cat-thumb-empty">\uFF0B</span>') +
        '</button>'
        : '') +
      '<span class="cat-name">' + escapeHtml(d.key) + '</span>' +
      (CatPrefs.compactMeta
        ? '<span class="cat-sub">' + d.colourList.length + ' colours \u00b7 ' +
          d.colourList.reduce((a, c) => a + c.sizes.size, 0) + ' sizes</span>'
        : '') +
    '</td>' +
    (catColOn('category') ? '<td class="cat-cat">' + escapeHtml(d.meta['Sub Section'] || d.meta.Section || '\u2014') + '</td>' : '') +
    (catColOn('colours') ? '<td class="cat-c-colours">' + dots + '</td>' : '') +
    (catColOn('sold') ? '<td class="num">' + fmtNum(d.sold) + '</td>' : '') +
    (catColOn('purchased') ? '<td class="num cat-purch">' + fmtNum(d.purchased) + '</td>' : '') +
    (catColOn('opening') ? '<td class="num obs-col">' + (d.hasOBS ? fmtNum(d.obs) : '\u2014') + '</td>' : '') +
    (catColOn('closing') ? '<td class="num cbs-col">' + fmtNum(d.cbs) + '</td>' : '') +
    (catColOn('moved') ? '<td class="num ' + (d.moved > 0 ? 'mv-down' : (d.moved < 0 ? 'mv-up' : '')) + '">' +
      (d.moved === null ? '\u2014' : (d.moved > 0 ? '\u2193 ' : (d.moved < 0 ? '\u2191 ' : '')) + fmtNum(Math.abs(d.moved))) + '</td>' : '') +
    replenCells(d.key, replenFor(d.key, d.sold, catalogDays(), d.cbs)) +
    (catColOn('cover') ? '<td class="num">' + (d.cover === Infinity ? '\u221E' : fmtNum(d.cover, 0) + 'd') + '</td>' : '') +
    (catColOn('sellthru') ? '<td class="num">' + fmtNum(d.sellThrough, 1) + '%</td>' : '') +
    (catColOn('lastsold') ? '<td>' + (d.lastSale ? fmtDate(d.lastSale) : '\u2014') + '</td>' : '') +
    (catColOn('status') ? '<td><span class="status-tag ' + catalogStatusClass(d.status) + '">' + d.status + '</span></td>' : '') +
  '</tr>';

  if (open) {
    d.colourList.forEach(c => { html += catalogColourRow(d, c); });
  }
  return html;
}

/** The colour-coded size bar drawn under a colour row. */
function catalogStripRow(d, c) {
  if (!CatPrefs.showStrip) return '';
  const total = catalogColCount();
  const blocks = c.sizeList.map(z => {
    const zr = replenFor(d.key + '|' + c.key + '|' + z.key, z.sold, catalogDays(), z.cbs);
    const w = Math.max(1, z.cbs);
    return '<span class="cs-block ' + stockPctClass(zr.pct) + '" style="flex:' + w + '" ' +
      'title="' + escapeHtml(z.key + ': ' + fmtNum(z.cbs) + ' in stock, ' + fmtNum(z.sold) + ' sold' +
        (isFinite(zr.pct) ? ', ' + Math.round(zr.pct) + '% of max level' : '')) + '"></span>';
  }).join('');
  if (!blocks) return '';
  return '<tr class="cat-strip-row"><td colspan="' + total + '">' +
    '<span class="cs-bar">' + blocks + '</span></td></tr>';
}

/** How many columns the table currently has, counting colspans. */
function catalogColCount() {
  return 1 +
    (catColOn('category') ? 1 : 0) + (catColOn('colours') ? 1 : 0) +
    ['sold','purchased','opening','closing','moved','adc','lt','sf','moq','ml','mit','stockpct','reorder',
     'cover','sellthru','lastsold','status'].filter(catColOn).length;
}

function catalogColourRow(d, c) {
  const cKey = d.key + '|' + c.key;
  const open = !!Catalog.expanded[cKey];
  const sizes = c.sizeList.slice(0, CatPrefs.maxSizeChips || 24).map(z => {
    // colour every size by its own stock %, using the same four bands
    const zr = replenFor(d.key + '|' + c.key + '|' + z.key, z.sold, catalogDays(), z.cbs);
    return '<span class="cat-size ' + stockPctClass(zr.pct) + '" ' +
    'title="' + escapeHtml(z.key + ': ' + fmtNum(z.cbs) + ' in stock, ' + fmtNum(z.sold) +
      ' sold, ' + (isFinite(zr.pct) ? Math.round(zr.pct) + '% of max level' : 'no demand')) + '">' +
    escapeHtml(truncateLabel(z.key, 4)) + '</span>';
  }).join('');

  let html = '<tr class="cat-row cat-colour' + (open ? ' is-open' : '') + '" data-key="' + escapeHtml(cKey) + '">' +
    '<td class="cat-c-design cat-indent">' +
      '<button class="cat-caret' + (open ? ' open' : '') + '">\u25B8</button>' +
      (CatPrefs.showThumbs
        ? '<button class="cat-thumb cat-thumb-sm" data-img="' + escapeHtml(cKey) + '" title="' +
          (Catalog.images[cKey] ? 'Click to view, replace or remove this colour photo' : 'Click to add a photo for this colour') + '">' +
          (Catalog.images[cKey] ? '<img src="' + Catalog.images[cKey] + '" alt="">' : '<span class="cat-thumb-empty">\uFF0B</span>') +
        '</button>'
        : '') +
      '<span class="cat-swatch" style="background:' + colourSwatch(c.key) + '"></span>' +
      '<span class="cat-name">' + escapeHtml(c.key) + '</span>' +
    '</td>' +
    (catColOn('category') || catColOn('colours')
      ? '<td class="cat-strip" colspan="' + ((catColOn('category') ? 1 : 0) + (catColOn('colours') ? 1 : 0)) + '">' + sizes + '</td>'
      : '') +
    (catColOn('sold') ? '<td class="num">' + fmtNum(c.sold) + '</td>' : '') +
    (catColOn('purchased') ? '<td class="num cat-purch">' + fmtNum(c.purchased) + '</td>' : '') +
    (catColOn('opening') ? '<td class="num obs-col">' + (c.hasOBS ? fmtNum(c.obs) : '\u2014') + '</td>' : '') +
    (catColOn('closing') ? '<td class="num cbs-col">' + fmtNum(c.cbs) + '</td>' : '') +
    (catColOn('moved') ? '<td class="num ' + (c.moved > 0 ? 'mv-down' : (c.moved < 0 ? 'mv-up' : '')) + '">' +
      (c.moved === null ? '\u2014' : (c.moved > 0 ? '\u2193 ' : (c.moved < 0 ? '\u2191 ' : '')) + fmtNum(Math.abs(c.moved))) + '</td>' : '') +
    replenCells(d.key + '|' + c.key, replenFor(d.key + '|' + c.key, c.sold, catalogDays(), c.cbs)) +
    (catColOn('cover') ? '<td class="num">' + (c.cover === Infinity ? '\u221E' : fmtNum(c.cover, 0) + 'd') + '</td>' : '') +
    (catColOn('sellthru') ? '<td class="num">' + fmtNum(c.sellThrough, 1) + '%</td>' : '') +
    (catColOn('lastsold') ? '<td>' + (c.lastSale ? fmtDate(c.lastSale) : '\u2014') + '</td>' : '') +
    (catColOn('status') ? '<td><span class="status-tag ' + catalogStatusClass(c.status) + '">' + c.status + '</span></td>' : '') +
  '</tr>';

  // Full-width bar under each colour: one block per size, coloured by its own
  // stock band, so a broken size run is obvious without opening the row.
  html += catalogStripRow(d, c);

  if (open) {
    c.sizeList.forEach(z => {
      const st = z.cbs === 0 && z.sold > 0 ? 'Stockout' : (z.sold === 0 && z.cbs > 0 ? 'No sale' : (z.sold > 0 ? 'Healthy' : 'Idle'));
      const filler = (catColOn('category') ? 1 : 0) + (catColOn('colours') ? 1 : 0);
      html += '<tr class="cat-row cat-size-row">' +
        '<td class="cat-c-design cat-indent2"><span class="cat-sizekey">' + escapeHtml(z.key) + '</span></td>' +
        (filler ? '<td colspan="' + filler + '"></td>' : '') +
        (catColOn('sold') ? '<td class="num">' + fmtNum(z.sold) + '</td>' : '') +
        (catColOn('purchased') ? '<td class="num cat-purch">' + fmtNum(z.purchased) + '</td>' : '') +
        (catColOn('opening') ? '<td class="num obs-col">' + (z.hasOBS ? fmtNum(z.obs) : '\u2014') + '</td>' : '') +
        (catColOn('closing') ? '<td class="num cbs-col">' + fmtNum(z.cbs) + '</td>' : '') +
        (catColOn('moved') ? '<td class="num">' + (z.hasOBS ? fmtNum(z.obs - z.cbs) : '\u2014') + '</td>' : '') +
        replenCells(d.key + '|' + c.key + '|' + z.key, replenFor(d.key + '|' + c.key + '|' + z.key, z.sold, catalogDays(), z.cbs)) +
        (catColOn('cover') ? '<td class="num"></td>' : '') +
        (catColOn('sellthru') ? '<td class="num"></td>' : '') +
        (catColOn('lastsold') ? '<td></td>' : '') +
        (catColOn('status') ? '<td><span class="status-tag ' + catalogStatusClass(st) + '">' + st + '</span></td>' : '') +
      '</tr>';
    });
  }
  return html;
}

function renderCatalogFilters(all) {
  const wrap = document.getElementById('cat-filters');
  if (!wrap) return;

  const letters = {};
  all.forEach(r => {
    const ch = String(r.key).trim().charAt(0).toUpperCase();
    if (/[A-Z0-9]/.test(ch)) letters[ch] = (letters[ch] || 0) + 1;
  });
  const letterKeys = Object.keys(letters).sort();

  const inLetter = Catalog.letter === 'all' ? all
    : all.filter(r => String(r.key).trim().charAt(0).toUpperCase() === Catalog.letter);

  const sections = {};
  inLetter.forEach(r => { const s = r.meta.Section || '(blank)'; sections[s] = (sections[s] || 0) + 1; });
  const sectionKeys = Object.keys(sections).sort((a, b) => sections[b] - sections[a]).slice(0, 14);

  const inSection = Catalog.section === 'all' ? inLetter
    : inLetter.filter(r => (r.meta.Section || '(blank)') === Catalog.section);
  const subs = {};
  inSection.forEach(r => { const s = r.meta['Sub Section'] || '(blank)'; subs[s] = (subs[s] || 0) + 1; });
  const subKeys = Object.keys(subs).sort((a, b) => subs[b] - subs[a]).slice(0, 18);

  wrap.innerHTML =
    '<div class="cat-filter-row"><span class="cat-flabel">Cat:</span>' +
      '<button class="cat-chip' + (Catalog.letter === 'all' ? ' active' : '') + '" data-letter="all">All</button>' +
      letterKeys.map(k => '<button class="cat-chip' + (Catalog.letter === k ? ' active' : '') + '" data-letter="' + k + '">' +
        k + '<span class="cat-chip-n">' + letters[k] + '</span></button>').join('') +
    '</div>' +
    '<div class="cat-filter-row"><span class="cat-flabel">Section:</span>' +
      '<button class="cat-chip' + (Catalog.section === 'all' ? ' active' : '') + '" data-section="all">All</button>' +
      sectionKeys.map(k => '<button class="cat-chip' + (Catalog.section === k ? ' active' : '') + '" data-section="' + escapeHtml(k) + '">' +
        escapeHtml(truncateLabel(k, 18)) + '<span class="cat-chip-n">' + sections[k] + '</span></button>').join('') +
    '</div>' +
    '<div class="cat-filter-row"><span class="cat-flabel">Sub-cat:</span>' +
      '<button class="cat-chip' + (Catalog.sub === 'all' ? ' active' : '') + '" data-sub="all">All</button>' +
      subKeys.map(k => '<button class="cat-chip' + (Catalog.sub === k ? ' active' : '') + '" data-sub="' + escapeHtml(k) + '">' +
        escapeHtml(truncateLabel(k, 20)) + '<span class="cat-chip-n">' + subs[k] + '</span></button>').join('') +
    '</div>';

  wrap.querySelectorAll('[data-letter]').forEach(b => b.addEventListener('click', () => {
    Catalog.letter = b.dataset.letter; Catalog.section = 'all'; Catalog.sub = 'all'; renderCatalog();
  }));
  wrap.querySelectorAll('[data-section]').forEach(b => b.addEventListener('click', () => {
    Catalog.section = b.dataset.section; Catalog.sub = 'all'; renderCatalog();
  }));
  wrap.querySelectorAll('[data-sub]').forEach(b => b.addEventListener('click', () => {
    Catalog.sub = b.dataset.sub; renderCatalog();
  }));
}

function wireCatalogRows() {
  const host = document.getElementById('cat-table');
  host.querySelectorAll('.cat-caret').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const key = btn.closest('tr').dataset.key;
      Catalog.expanded[key] = !Catalog.expanded[key];
      renderCatalog();
    });
  });
  host.querySelectorAll('.cat-thumb').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      // no photo yet -> pick one; photo present -> open it large
      openCatImage(btn.dataset.img);
    });
  });
  host.querySelectorAll('tr.cat-design, tr.cat-colour').forEach(tr => {
    tr.addEventListener('click', () => {
      const c = tr.querySelector('.cat-caret');
      if (c) c.click();
    });
  });
}

/* ---- photos, kept in this browser ---- */
function loadCatalogImages() {
  try {
    const raw = Store.get('sl_catalog_images');
    if (raw) Catalog.images = JSON.parse(raw) || {};
  } catch (e) { Catalog.images = {}; }
}
function saveCatalogImages() {
  try { Store.set('sl_catalog_images', JSON.stringify(Catalog.images)); }
  catch (e) { toast('Could not save the photo - browser storage is full.'); }
}

function onCatalogImagePicked(e) {
  const file = e.target.files && e.target.files[0];
  const key = Catalog.pendingImage;
  e.target.value = '';
  if (!file || !key) return;
  const reader = new FileReader();
  reader.onload = ev => {
    // shrink it so many photos still fit in browser storage
    const img = new Image();
    img.onload = () => {
      const max = 120;
      const scale = Math.min(1, max / Math.max(img.width, img.height));
      const cv = document.createElement('canvas');
      cv.width = Math.round(img.width * scale);
      cv.height = Math.round(img.height * scale);
      cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
      Catalog.images[key] = cv.toDataURL('image/jpeg', 0.72);
      saveCatalogImages();
      renderCatalog();
      toast('Photo added to ' + key + '.');
    };
    img.onerror = () => toast('That file could not be read as an image.');
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
}

function exportCatalogCSV() {
  if (!Catalog.lastRows || !Catalog.lastRows.length) { toast('Nothing to export yet.'); return; }
  const headers = ['Article No', 'Section', 'Sub Section', 'Brand', 'Supplier', 'Colour', 'Size',
                   'Sold', 'Opening', 'Closing', 'Moved', 'Cover days', 'Sell-through %', 'Status'];
  const out = [];
  Catalog.lastRows.forEach(d => {
    d.colourList.forEach(c => {
      c.sizeList.forEach(z => {
        out.push([d.key, d.meta.Section || '', d.meta['Sub Section'] || '', d.meta.Brand || '',
                  d.meta.Supplier || '', c.key, z.key, z.sold, z.hasOBS ? z.obs : '', z.cbs,
                  z.hasOBS ? (z.obs - z.cbs) : '', '', '', '']);
      });
      out.push([d.key, d.meta.Section || '', d.meta['Sub Section'] || '', d.meta.Brand || '',
                d.meta.Supplier || '', c.key, 'ALL SIZES', c.sold, c.hasOBS ? c.obs : '', c.cbs,
                c.moved === null ? '' : c.moved,
                c.cover === Infinity ? '' : Math.round(c.cover), Number(c.sellThrough.toFixed(1)), c.status]);
    });
    out.push([d.key, d.meta.Section || '', d.meta['Sub Section'] || '', d.meta.Brand || '',
              d.meta.Supplier || '', 'ALL COLOURS', '', d.sold, d.hasOBS ? d.obs : '', d.cbs,
              d.moved === null ? '' : d.moved,
              d.cover === Infinity ? '' : Math.round(d.cover), Number(d.sellThrough.toFixed(1)), d.status]);
  });
  downloadBlob(toCSV(headers, out), 'catalog.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   15b. CATALOG PREFERENCES + IMAGE VIEWER
   --------------------------------------------------------------- */

const CAT_COLUMNS = [
  ['category', 'Category'], ['colours', 'Colours'], ['sold', 'Sold'], ['purchased', 'Purchased'],
  ['opening', 'Opening'], ['closing', 'Closing'], ['moved', 'Moved'],
  ['adc', 'ADC'], ['lt', 'LT'], ['sf', 'SF'], ['moq', 'MOQ'],
  ['ml', 'ML'], ['mit', 'MIT'], ['stockpct', 'Stock %'], ['reorder', 'Reorder'],
  ['cover', 'Cover'], ['sellthru', 'Sell-thru'], ['lastsold', 'Last sold'],
  ['status', 'Status']
];

/* ---------------------------------------------------------------
   Replenishment maths
     ADC = average daily consumption   (auto = sold / days, editable)
     LT  = lead time in days           (default, editable)
     SF  = safety factor               (default, editable)
     MOQ = minimum order quantity      (default, editable)
     ML  = max level  = ADC x LT x SF
     MIT = material in transit         (entered by hand)
     Stock % = (closing + MIT) / ML
     Reorder = ML - (closing + MIT), rounded up to the MOQ
   --------------------------------------------------------------- */
const Replen = { overrides: {}, bulk: {} };   // bulk = last values typed in the "set for all" bar

function loadReplen() {
  try {
    const raw = Store.get('sl_replen');
    if (raw) Replen.overrides = JSON.parse(raw) || {};
  } catch (e) { Replen.overrides = {}; }
  try {
    const b = Store.get('sl_replen_bulk');
    if (b) Replen.bulk = JSON.parse(b) || {};
  } catch (e) { Replen.bulk = {}; }
}
function saveReplen() { Store.set('sl_replen', JSON.stringify(Replen.overrides)); }
function saveReplenBulk() { Store.set('sl_replen_bulk', JSON.stringify(Replen.bulk)); }

function replenFor(key, sold, days, cbs) {
  const o = Replen.overrides[key] || {};
  const autoAdc = days > 0 ? sold / days : 0;
  const adc = (o.adc !== undefined && o.adc !== null) ? o.adc
              : ((CatPrefs.defaultADC > 0) ? CatPrefs.defaultADC : autoAdc);
  const lt = (o.lt !== undefined && o.lt !== null) ? o.lt : (CatPrefs.defaultLT || 15);
  const sf = (o.sf !== undefined && o.sf !== null) ? o.sf : (CatPrefs.defaultSF || 1.5);
  const moq = (o.moq !== undefined && o.moq !== null) ? o.moq : (CatPrefs.defaultMOQ || 12);
  const mit = (o.mit !== undefined && o.mit !== null) ? o.mit : 0;

  const ml = adc * lt * sf;
  const available = cbs + mit;
  let pct;
  if (ml > 0) pct = (available / ml) * 100;
  else pct = available > 0 ? 101 : 0;      // no demand but stock on hand = overstock

  let reorder = 0;
  if (ml > available) {
    const gap = ml - available;
    reorder = moq > 0 ? Math.ceil(gap / moq) * moq : Math.ceil(gap);
  }
  return { adc, lt, sf, moq, ml, mit, available, pct, reorder,
           isAuto: { adc: o.adc === undefined, lt: o.lt === undefined, sf: o.sf === undefined,
                     moq: o.moq === undefined, mit: o.mit === undefined } };
}

/** 0-33 red, 33-66 yellow, 66-100 green, over 100 purple. */
function stockPctClass(pct) {
  if (pct > 100) return 'sp-over';
  if (pct > 66) return 'sp-good';
  if (pct > 33) return 'sp-mid';
  return 'sp-low';
}

function replenCells(key, r) {
  const inp = (field, val, step) =>
    '<input class="cat-inp' + (r.isAuto[field] ? ' is-auto' : ' is-set') + '" type="number" step="' + step + '" ' +
    'value="' + val + '" data-rkey="' + escapeHtml(key) + '" data-rfield="' + field + '" ' +
    'title="' + (r.isAuto[field] ? 'Automatic - type to override' : 'Set by you - clear the box to go back to automatic') + '">';

  return (catColOn('adc') ? '<td class="num cat-edit">' + inp('adc', fmtNum(r.adc, 2), '0.01') + '</td>' : '') +
         (catColOn('lt') ? '<td class="num cat-edit">' + inp('lt', r.lt, '1') + '</td>' : '') +
         (catColOn('sf') ? '<td class="num cat-edit">' + inp('sf', r.sf, '0.1') + '</td>' : '') +
         (catColOn('moq') ? '<td class="num cat-edit">' + inp('moq', r.moq, '1') + '</td>' : '') +
         (catColOn('ml') ? '<td class="num cat-ml">' + fmtNum(r.ml, 0) + '</td>' : '') +
         (catColOn('mit') ? '<td class="num cat-edit">' + inp('mit', r.mit, '1') + '</td>' : '') +
         (catColOn('stockpct') ? '<td class="num stock-pct ' + stockPctClass(r.pct) + '">' + fmtNum(r.pct, 0) + '%</td>' : '') +
         (catColOn('reorder') ? '<td class="num cat-reorder' + (r.reorder > 0 ? ' has' : '') + '">' +
            (r.reorder > 0 ? fmtNum(r.reorder) : '\u2014') + '</td>' : '');
}

const CATPREFS_DEFAULT = {
  // look
  density: 'compact',        // compact | normal | roomy
  fontSize: 11.5,
  headerStyle: 'dark',       // dark | light
  zebra: true,
  gridLines: true,
  accent: '#2C5AA0',         // selected chip colour (the reference UI uses blue)
  hoverColor: '#EAF3FC',
  // thumbnails
  thumbW: 30,
  thumbH: 34,
  showThumbs: true,
  // colour dots
  dotShape: 'square',        // square | circle
  dotSize: 12,
  maxDots: 10,
  showDotCounts: false,
  // size strip
  maxSizeChips: 24,
  showStrip: true,      // colour-coded size bar under each colour row
  lowStockAt: 2,
  // thresholds
  lowCoverDays: 15,
  overstockDays: 120,
  // replenishment defaults
  defaultADC: 0,      // 0 = work it out from sales; anything else is used as-is
  defaultLT: 15,      // lead time, days
  defaultSF: 1.5,     // safety factor
  defaultMOQ: 12,     // minimum order quantity
  // table
  columns: CAT_COLUMNS.map(c => c[0]),
  maxRows: 300,
  showChipCounts: true,
  compactMeta: true
};

const CatPrefs = Object.assign({}, CATPREFS_DEFAULT);

// Columns added in later builds must appear for people who already had
// settings saved, otherwise a new column stays invisible for ever.
const CAT_COLUMNS_ADDED_LATER = ['purchased'];

function loadCatPrefs() {
  try {
    const raw = Store.get('sl_catprefs');
    if (raw) {
      const saved = JSON.parse(raw);
      Object.assign(CatPrefs, CATPREFS_DEFAULT, saved);
      if (Array.isArray(saved.columns)) {
        const known = CAT_COLUMNS.map(c => c[0]);
        CatPrefs.columns = saved.columns.filter(c => known.indexOf(c) !== -1);
        // switch on any column that did not exist when these settings were saved
        CAT_COLUMNS_ADDED_LATER.forEach(c => {
          if (CatPrefs.columns.indexOf(c) === -1) {
            const at = CatPrefs.columns.indexOf('sold');
            if (at >= 0) CatPrefs.columns.splice(at + 1, 0, c);
            else CatPrefs.columns.push(c);
          }
        });
        saveCatPrefsQuiet();
      }
    }
  } catch (e) {}
  applyCatPrefs();
}

function saveCatPrefsQuiet() { Store.set('sl_catprefs', JSON.stringify(CatPrefs)); }
function saveCatPrefs() { Store.set('sl_catprefs', JSON.stringify(CatPrefs)); applyCatPrefs(); }

function applyCatPrefs() {
  const r = document.documentElement;
  r.style.setProperty('--cat-font', CatPrefs.fontSize + 'px');
  r.style.setProperty('--cat-pad', (CatPrefs.density === 'compact' ? 2 : CatPrefs.density === 'roomy' ? 9 : 5) + 'px');
  r.style.setProperty('--cat-accent', CatPrefs.accent);
  r.style.setProperty('--cat-hover', CatPrefs.hoverColor);
  r.style.setProperty('--cat-thumb-w', CatPrefs.thumbW + 'px');
  r.style.setProperty('--cat-thumb-h', CatPrefs.thumbH + 'px');
  r.style.setProperty('--cat-dot', CatPrefs.dotSize + 'px');
  r.style.setProperty('--cat-dot-radius', CatPrefs.dotShape === 'circle' ? '50%' : '2px');
  document.body.classList.toggle('cat-light-head', CatPrefs.headerStyle === 'light');
  document.body.classList.toggle('cat-no-zebra', !CatPrefs.zebra);
  document.body.classList.toggle('cat-no-grid', !CatPrefs.gridLines);
  document.body.classList.toggle('cat-no-thumbs', !CatPrefs.showThumbs);
}

function catColOn(id) { return (CatPrefs.columns || []).indexOf(id) !== -1; }

/* ---- image viewer: click a photo to enlarge, replace or remove ---- */
function ensureCatImageViewer() {
  if (document.getElementById('cat-img-overlay')) return;
  const d = document.createElement('div');
  d.id = 'cat-img-overlay';
  d.className = 'drill-overlay cat-img-overlay';
  d.style.display = 'none';
  d.innerHTML =
    '<div class="cat-img-panel">' +
      '<div class="drill-head"><div>' +
        '<h2 id="cat-img-title">Photo</h2>' +
        '<div class="drill-subtitle" id="cat-img-sub"></div>' +
      '</div><button id="cat-img-close" class="drill-close" title="Close (Esc)">&times;</button></div>' +
      '<div class="cat-img-body"><img id="cat-img-big" alt=""></div>' +
      '<div class="modal-actions">' +
        '<button class="ghost-btn small primary" id="cat-img-replace">Replace photo</button>' +
        '<button class="ghost-btn small" id="cat-img-remove">Remove photo</button>' +
        '<span class="spacer"></span>' +
        '<button class="ghost-btn small" id="cat-img-done">Close</button>' +
      '</div>' +
    '</div>';
  document.body.appendChild(d);

  const close = () => { d.style.display = 'none'; modalClose('cat-img-overlay'); };
  document.getElementById('cat-img-close').addEventListener('click', close);
  document.getElementById('cat-img-done').addEventListener('click', close);
  d.addEventListener('click', e => { if (e.target === d && modalIsTop('cat-img-overlay')) close(); });
  document.getElementById('cat-img-replace').addEventListener('click', () => {
    Catalog.pendingImage = Catalog.viewingImage;
    document.getElementById('cat-image-input').click();
  });
  document.getElementById('cat-img-remove').addEventListener('click', () => {
    const k = Catalog.viewingImage;
    if (!k) return;
    delete Catalog.images[k];
    saveCatalogImages();
    close();
    renderCatalog();
    toast('Photo removed from ' + k + '.');
  });
}

function openCatImage(key) {
  ensureCatImageViewer();
  const src = Catalog.images[key];
  if (!src) { Catalog.pendingImage = key; document.getElementById('cat-image-input').click(); return; }
  Catalog.viewingImage = key;
  document.getElementById('cat-img-title').textContent = key;
  const d = (Catalog.lastRows || []).find(r => r.key === key);
  document.getElementById('cat-img-sub').textContent = d
    ? (d.colourList.length + ' colours \u00b7 ' + fmtNum(d.cbs) + ' in stock \u00b7 ' + fmtNum(d.sold) + ' sold')
    : '';
  document.getElementById('cat-img-big').src = src;
  document.getElementById('cat-img-overlay').style.display = 'flex';
  modalOpen('cat-img-overlay', () => {
    document.getElementById('cat-img-overlay').style.display = 'none';
  });
}

/* ---- the Catalog settings page ---- */
function renderCatalogSettings(wrap) {
  const row = (label, control) =>
    '<div class="settings-row"><label class="toolbar-label">' + label + '</label>' + control + '</div>';

  wrap.innerHTML =
    '<h3 class="snap-set-title">Table look</h3>' +
    row('Row density',
      '<select id="cs-density" class="select">' +
        ['compact', 'normal', 'roomy'].map(d => '<option value="' + d + '"' + (CatPrefs.density === d ? ' selected' : '') + '>' +
          d.charAt(0).toUpperCase() + d.slice(1) + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Font size</label>' +
      '<input type="range" id="cs-font" min="9" max="16" step="0.5" value="' + CatPrefs.fontSize + '">' +
      '<span class="drill-count" id="cs-font-val">' + CatPrefs.fontSize + 'px</span>') +
    row('Header style',
      '<select id="cs-head" class="select">' +
        '<option value="dark"' + (CatPrefs.headerStyle === 'dark' ? ' selected' : '') + '>Dark bar</option>' +
        '<option value="light"' + (CatPrefs.headerStyle === 'light' ? ' selected' : '') + '>Light bar</option></select>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-zebra"' + (CatPrefs.zebra ? ' checked' : '') + '> Row striping</label>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-grid"' + (CatPrefs.gridLines ? ' checked' : '') + '> Grid lines</label>') +
    '<div class="color-row"><label class="toolbar-label">Selected chip colour</label>' +
      '<input type="color" id="cs-accent" value="' + CatPrefs.accent + '">' +
      '<span class="hexcode">' + CatPrefs.accent + '</span>' +
      '<label class="toolbar-label">Row hover</label>' +
      '<input type="color" id="cs-hover" value="' + CatPrefs.hoverColor + '">' +
      '<span class="hexcode">' + CatPrefs.hoverColor + '</span></div>' +

    '<h3 class="snap-set-title">Photos</h3>' +
    row('Show thumbnails',
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-thumbs"' + (CatPrefs.showThumbs ? ' checked' : '') + '> Show a photo column</label>' +
      '<label class="toolbar-label">Width</label>' +
      '<input type="range" id="cs-thumbw" min="18" max="70" value="' + CatPrefs.thumbW + '">' +
      '<span class="drill-count" id="cs-thumbw-val">' + CatPrefs.thumbW + 'px</span>' +
      '<label class="toolbar-label">Height</label>' +
      '<input type="range" id="cs-thumbh" min="18" max="80" value="' + CatPrefs.thumbH + '">' +
      '<span class="drill-count" id="cs-thumbh-val">' + CatPrefs.thumbH + 'px</span>') +
    '<p class="drill-subtitle">Click an empty box to add a photo. Click a photo to open it larger, where you can replace or remove it.</p>' +
    '<div class="settings-row"><span class="drill-count" id="cs-imgcount"></span>' +
      '<button class="ghost-btn small" id="cs-clearimgs">Remove all photos</button></div>' +

    '<h3 class="snap-set-title">Colour dots</h3>' +
    row('Appearance',
      '<select id="cs-dotshape" class="select">' +
        '<option value="square"' + (CatPrefs.dotShape === 'square' ? ' selected' : '') + '>Squares</option>' +
        '<option value="circle"' + (CatPrefs.dotShape === 'circle' ? ' selected' : '') + '>Circles</option></select>' +
      '<label class="toolbar-label">Size</label>' +
      '<input type="range" id="cs-dotsize" min="7" max="20" value="' + CatPrefs.dotSize + '">' +
      '<span class="drill-count" id="cs-dotsize-val">' + CatPrefs.dotSize + 'px</span>' +
      '<label class="toolbar-label">How many</label>' +
      '<input type="number" id="cs-maxdots" class="text-input narrow" min="3" max="40" value="' + CatPrefs.maxDots + '">' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-dotcounts"' + (CatPrefs.showDotCounts ? ' checked' : '') + '> Show stock number on each dot</label>') +

    '<h3 class="snap-set-title">Size run</h3>' +
    row('Chips',
      '<label class="toolbar-label">Max shown</label>' +
      '<input type="number" id="cs-maxsizes" class="text-input narrow" min="4" max="60" value="' + CatPrefs.maxSizeChips + '">' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-strip"' +
        (CatPrefs.showStrip ? ' checked' : '') + '> Colour bar under each colour row</label>' +
      '<label class="toolbar-label">Amber when stock is at or below</label>' +
      '<input type="number" id="cs-lowstock" class="text-input narrow" min="0" max="50" value="' + CatPrefs.lowStockAt + '">') +

    '<h3 class="snap-set-title">Replenishment defaults</h3>' +
    '<p class="drill-subtitle">Starting values for every design. Any figure you type into a row overrides these.</p>' +
    '<div class="settings-row">' +
      '<label class="toolbar-label">ADC</label>' +
      '<input type="number" id="cs-adc" class="text-input narrow" min="0" step="0.01" value="' + (CatPrefs.defaultADC || 0) + '">' +
      '<span class="drill-count">0 = calculate from sales</span>' +
      '<label class="toolbar-label">LT</label>' +
      '<input type="number" id="cs-lt" class="text-input narrow" min="0" step="1" value="' + (CatPrefs.defaultLT || 15) + '"> days' +
      '<label class="toolbar-label">SF</label>' +
      '<input type="number" id="cs-sf" class="text-input narrow" min="0" step="0.1" value="' + (CatPrefs.defaultSF || 1.5) + '">' +
      '<label class="toolbar-label">MOQ</label>' +
      '<input type="number" id="cs-moq" class="text-input narrow" min="0" step="1" value="' + (CatPrefs.defaultMOQ || 12) + '">' +
    '</div>' +

    '<h3 class="snap-set-title">Status thresholds</h3>' +
    row('Cover days',
      '<label class="toolbar-label">Low cover under</label>' +
      '<input type="number" id="cs-lowcover" class="text-input narrow" min="1" max="120" value="' + CatPrefs.lowCoverDays + '"> days' +
      '<label class="toolbar-label">Overstock over</label>' +
      '<input type="number" id="cs-overstock" class="text-input narrow" min="10" max="900" value="' + CatPrefs.overstockDays + '"> days') +

    '<h3 class="snap-set-title">Columns</h3>' +
    '<div class="snap-checklist">' +
      CAT_COLUMNS.map(c => '<label class="snap-check"><input type="checkbox" class="cs-col" value="' + c[0] + '"' +
        (catColOn(c[0]) ? ' checked' : '') + '> ' + c[1] + '</label>').join('') +
    '</div>' +
    row('Rows rendered',
      '<input type="number" id="cs-maxrows" class="text-input narrow" min="50" max="2000" step="50" value="' + CatPrefs.maxRows + '">' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-chipcounts"' + (CatPrefs.showChipCounts ? ' checked' : '') + '> Counts on filter chips</label>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-meta"' + (CatPrefs.compactMeta ? ' checked' : '') + '> "n colours / n sizes" under each design</label>') +

    '<div class="modal-actions" style="margin-top:16px;">' +
      '<button class="ghost-btn small" id="cs-reset">Reset Catalog settings</button>' +
    '</div>';

  const n = Object.keys(Catalog.images || {}).length;
  const ic = wrap.querySelector('#cs-imgcount');
  if (ic) ic.textContent = n ? (n + ' photo' + (n === 1 ? '' : 's') + ' saved') : 'No photos saved yet';

  const bind = (id, fn) => { const el = wrap.querySelector('#' + id); if (el) el.addEventListener('input', fn); };
  const bindC = (id, fn) => { const el = wrap.querySelector('#' + id); if (el) el.addEventListener('change', fn); };

  bindC('cs-density', e => { CatPrefs.density = e.target.value; saveCatPrefs(); renderCatalog(); });
  bind('cs-font', e => {
    CatPrefs.fontSize = parseFloat(e.target.value);
    wrap.querySelector('#cs-font-val').textContent = CatPrefs.fontSize + 'px'; saveCatPrefs();
  });
  bindC('cs-head', e => { CatPrefs.headerStyle = e.target.value; saveCatPrefs(); });
  bindC('cs-zebra', e => { CatPrefs.zebra = e.target.checked; saveCatPrefs(); });
  bindC('cs-grid', e => { CatPrefs.gridLines = e.target.checked; saveCatPrefs(); });
  bind('cs-accent', e => { CatPrefs.accent = e.target.value; saveCatPrefs(); });
  bind('cs-hover', e => { CatPrefs.hoverColor = e.target.value; saveCatPrefs(); });

  bindC('cs-thumbs', e => { CatPrefs.showThumbs = e.target.checked; saveCatPrefs(); renderCatalog(); });
  bind('cs-thumbw', e => {
    CatPrefs.thumbW = parseInt(e.target.value, 10);
    wrap.querySelector('#cs-thumbw-val').textContent = CatPrefs.thumbW + 'px'; saveCatPrefs();
  });
  bind('cs-thumbh', e => {
    CatPrefs.thumbH = parseInt(e.target.value, 10);
    wrap.querySelector('#cs-thumbh-val').textContent = CatPrefs.thumbH + 'px'; saveCatPrefs();
  });
  const ci = wrap.querySelector('#cs-clearimgs');
  if (ci) ci.addEventListener('click', () => {
    Catalog.images = {}; saveCatalogImages(); renderCatalog(); renderSettingsBody();
    toast('All photos removed.');
  });

  bindC('cs-dotshape', e => { CatPrefs.dotShape = e.target.value; saveCatPrefs(); });
  bind('cs-dotsize', e => {
    CatPrefs.dotSize = parseInt(e.target.value, 10);
    wrap.querySelector('#cs-dotsize-val').textContent = CatPrefs.dotSize + 'px'; saveCatPrefs();
  });
  bindC('cs-maxdots', e => { CatPrefs.maxDots = Math.max(3, Math.min(40, parseInt(e.target.value, 10) || 10)); saveCatPrefs(); renderCatalog(); });
  bindC('cs-dotcounts', e => { CatPrefs.showDotCounts = e.target.checked; saveCatPrefs(); renderCatalog(); });

  bindC('cs-maxsizes', e => { CatPrefs.maxSizeChips = Math.max(4, Math.min(60, parseInt(e.target.value, 10) || 24)); saveCatPrefs(); renderCatalog(); });
  bindC('cs-strip', e => { CatPrefs.showStrip = e.target.checked; saveCatPrefs(); renderCatalog(); });
  bindC('cs-lowstock', e => { CatPrefs.lowStockAt = Math.max(0, parseInt(e.target.value, 10) || 0); saveCatPrefs(); renderCatalog(); });
  bindC('cs-lowcover', e => { CatPrefs.lowCoverDays = Math.max(1, parseInt(e.target.value, 10) || 15); saveCatPrefs(); renderCatalog(); });
  bindC('cs-overstock', e => { CatPrefs.overstockDays = Math.max(10, parseInt(e.target.value, 10) || 120); saveCatPrefs(); renderCatalog(); });

  wrap.querySelectorAll('.cs-col').forEach(cb => cb.addEventListener('change', () => {
    CatPrefs.columns = [...wrap.querySelectorAll('.cs-col:checked')].map(x => x.value);
    saveCatPrefs(); renderCatalog();
  }));
  bindC('cs-maxrows', e => { CatPrefs.maxRows = Math.max(50, Math.min(2000, parseInt(e.target.value, 10) || 300)); saveCatPrefs(); renderCatalog(); });
  bindC('cs-chipcounts', e => { CatPrefs.showChipCounts = e.target.checked; saveCatPrefs(); renderCatalog(); });
  bindC('cs-meta', e => { CatPrefs.compactMeta = e.target.checked; saveCatPrefs(); renderCatalog(); });

  bindC('cs-adc', e => { CatPrefs.defaultADC = Math.max(0, parseFloat(e.target.value) || 0); saveCatPrefs(); renderCatalog(); });
  bindC('cs-lt',  e => { CatPrefs.defaultLT  = Math.max(0, parseInt(e.target.value, 10) || 15); saveCatPrefs(); renderCatalog(); });
  bindC('cs-sf',  e => { CatPrefs.defaultSF  = Math.max(0, parseFloat(e.target.value) || 1.5); saveCatPrefs(); renderCatalog(); });
  bindC('cs-moq', e => { CatPrefs.defaultMOQ = Math.max(0, parseInt(e.target.value, 10) || 12); saveCatPrefs(); renderCatalog(); });

  const rs = wrap.querySelector('#cs-reset');
  if (rs) rs.addEventListener('click', () => {
    Object.assign(CatPrefs, CATPREFS_DEFAULT);
    CatPrefs.columns = CAT_COLUMNS.map(c => c[0]);
    saveCatPrefs(); renderCatalog(); renderSettingsBody();
    toast('Catalog settings reset.');
  });
}

/* ---------------------------------------------------------------
   16. DASHBOARD BUILDER — add your own charts, click to cross-filter
   ---------------------------------------------------------------
   Every chart is saved, so the layout you build stays until you
   change it. Clicking a bar or slice filters every other chart on
   the page, and one button clears the lot.
   --------------------------------------------------------------- */

const CHART_TYPES = [
  ['bar', 'Bar (horizontal)'],
  ['column', 'Column (vertical)'],
  ['line', 'Line'],
  ['area', 'Area'],
  ['pie', 'Pie'],
  ['doughnut', 'Doughnut'],
  ['table', 'Table']
];

const CHART_SOURCES = [
  ['sales', 'Sales'],
  ['purchase', 'Purchase'],
  ['stock', 'Stock']
];

const CHART_DIMS = ['Article No', 'Brand', 'Colour', 'Size', 'Style', 'Section',
                    'Sub Section', 'Supplier', 'Item Code', 'Month', 'Week', 'Transaction Date'];

const CHART_MEASURES = [
  ['qty', 'Quantity'],
  ['rows', 'Row count'],
  ['items', 'Distinct items']
];

const DEFAULT_CHARTS = [
  { id: 'c1', title: 'Sales by month',       type: 'column', source: 'sales',    dim: 'Month',       measure: 'qty', topN: 24 },
  { id: 'c2', title: 'Top brands (sold)',    type: 'bar',    source: 'sales',    dim: 'Brand',       measure: 'qty', topN: 10 },
  { id: 'c3', title: 'Sales by section',     type: 'doughnut', source: 'sales',  dim: 'Section',     measure: 'qty', topN: 8 },
  { id: 'c4', title: 'Purchases by month',   type: 'column', source: 'purchase', dim: 'Month',       measure: 'qty', topN: 24 },
  { id: 'c5', title: 'Stock by sub section', type: 'bar',    source: 'stock',    dim: 'Sub Section', measure: 'qty', topN: 10 },
  { id: 'c6', title: 'Top selling articles', type: 'table',  source: 'sales',    dim: 'Article No',  measure: 'qty', topN: 12 }
];

const Dash = {
  charts: [],
  filters: [],          // [{dim, value}] - the Power BI style cross filter
  editing: null,
  instances: {}
};

function loadDashCharts() {
  try {
    const raw = Store.get('sl_dash_charts');
    Dash.charts = raw ? JSON.parse(raw) : DEFAULT_CHARTS.slice();
  } catch (e) { Dash.charts = DEFAULT_CHARTS.slice(); }
  if (!Dash.charts.length) Dash.charts = DEFAULT_CHARTS.slice();
}
function saveDashCharts() { Store.set('sl_dash_charts', JSON.stringify(Dash.charts)); }

function dashRecordsFor(source) {
  const range = periodRange();
  if (source === 'purchase') return purchaseRecords().filter(r => inPeriod(r, range));
  if (source === 'stock') return stockRecords();
  return salesRecords().filter(r => inPeriod(r, range));
}

function dashDimKey(rec, dim) {
  if (dim === 'Month') return rec.Date ? dateKeyForGrain(rec.Date, 'month') : '(blank)';
  if (dim === 'Week') return rec.Date ? weekKeyOf(rec.Date) : '(blank)';
  if (dim === 'Transaction Date') return rec.Date ? fmtDate(rec.Date) : '(blank)';
  return dimKey(rec, dim);
}

/** Applies the cross filter set by clicking other charts. */
function dashPassesFilters(rec) {
  return Dash.filters.every(f => dashDimKey(rec, f.dim) === f.value);
}

function dashAggregate(cfg) {
  let recs = dashRecordsFor(cfg.source).filter(dashPassesFilters);
  const map = new Map();
  const seen = new Map();
  recs.forEach(r => {
    const k = dashDimKey(r, cfg.dim);
    if (cfg.measure === 'rows') map.set(k, (map.get(k) || 0) + 1);
    else if (cfg.measure === 'items') {
      let set = seen.get(k);
      if (!set) { set = new Set(); seen.set(k, set); }
      const code = r['Item Code'] || r['Article No'];
      if (code) set.add(String(code));
      map.set(k, set.size);
    } else map.set(k, (map.get(k) || 0) + recQty(r));
  });

  let entries = [...map.entries()].filter(([k]) => k !== '(blank)');
  const chrono = ['Month', 'Week', 'Transaction Date'].indexOf(cfg.dim) !== -1;
  if (chrono) {
    const g = cfg.dim === 'Month' ? 'month' : 'day';
    entries.sort((a, b) => (cfg.dim === 'Month' ? grainSort(a[0], g) - grainSort(b[0], g)
      : (parseDateLoose(a[0].split(' \u2013 ')[0]) || 0) - (parseDateLoose(b[0].split(' \u2013 ')[0]) || 0)));
    if (entries.length > (cfg.topN || 24)) entries = entries.slice(-(cfg.topN || 24));
  } else {
    entries.sort((a, b) => b[1] - a[1]);
    entries = entries.slice(0, cfg.topN || 10);
  }
  return entries;
}

function initDashboardBuilder() {
  loadDashCharts();
  const add = document.getElementById('dash-add-chart');
  if (add) add.addEventListener('click', () => openChartEditor(null));
  const reset = document.getElementById('dash-reset');
  if (reset) reset.addEventListener('click', () => {
    Dash.charts = DEFAULT_CHARTS.slice();
    Dash.filters = [];
    saveDashCharts();
    renderDashboard();
    toast('Dashboard reset to the standard charts.');
  });
  const clearF = document.getElementById('dash-clear-filters');
  if (clearF) clearF.addEventListener('click', () => { Dash.filters = []; renderDashboard(); });
}

function renderDashFilters() {
  const wrap = document.getElementById('dash-filters');
  if (!wrap) return;
  if (!Dash.filters.length) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = '<span class="toolbar-label">Filtered by:</span>' +
    Dash.filters.map((f, i) =>
      '<span class="filter-chip">' + escapeHtml(f.dim) + ': <strong>' + escapeHtml(f.value) + '</strong>' +
      '<button data-fi="' + i + '">&times;</button></span>').join('') +
    '<button class="ghost-btn small" id="dash-clear-inline">Clear all</button>';
  wrap.querySelectorAll('[data-fi]').forEach(b => b.addEventListener('click', () => {
    Dash.filters.splice(parseInt(b.dataset.fi, 10), 1);
    renderDashboard();
  }));
  const ci = wrap.querySelector('#dash-clear-inline');
  if (ci) ci.addEventListener('click', () => { Dash.filters = []; renderDashboard(); });
}

function renderCustomCharts() {
  const grid = document.getElementById('dash-grid');
  if (!grid) return;
  Object.keys(Dash.instances).forEach(id => {
    if (Dash.instances[id]) { try { Dash.instances[id].destroy(); } catch (e) {} }
  });
  Dash.instances = {};

  if (!App.datasets.length) {
    grid.innerHTML = '<div class="empty-hint big">Load your reports on the Import tab to build a dashboard.</div>';
    return;
  }

  grid.innerHTML = Dash.charts.map(cfg =>
    '<div class="dash-card' + (cfg.wide ? ' wide' : '') + '" data-cid="' + cfg.id + '">' +
      '<div class="dash-card-head">' +
        '<h3>' + escapeHtml(cfg.title) + '</h3>' +
        '<span class="dash-card-meta">' + escapeHtml(cfg.source) + ' \u00b7 ' + escapeHtml(cfg.dim) + '</span>' +
        '<span class="spacer"></span>' +
        '<button class="dash-ico" data-act="wide" title="Make this chart wide or normal">\u2194</button>' +
        '<button class="dash-ico" data-act="edit" title="Edit this chart">\u270E</button>' +
        '<button class="dash-ico" data-act="dup" title="Duplicate">\u29C9</button>' +
        '<button class="dash-ico" data-act="del" title="Remove">\u2715</button>' +
      '</div>' +
      (cfg.type === 'table'
        ? '<div class="dash-table-wrap"><table class="data-table dash-mini" id="dt-' + cfg.id + '"></table></div>'
        : '<div class="chart-box"><canvas id="dc-' + cfg.id + '"></canvas></div>') +
    '</div>').join('');

  Dash.charts.forEach(cfg => drawCustomChart(cfg));

  grid.querySelectorAll('.dash-ico').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const id = btn.closest('.dash-card').dataset.cid;
      const i = Dash.charts.findIndex(c => c.id === id);
      const act = btn.dataset.act;
      if (act === 'del') { Dash.charts.splice(i, 1); saveDashCharts(); renderDashboard(); }
      else if (act === 'edit') openChartEditor(Dash.charts[i]);
      else if (act === 'dup') {
        const copy = Object.assign({}, Dash.charts[i], { id: 'c' + Date.now(), title: Dash.charts[i].title + ' (copy)' });
        Dash.charts.splice(i + 1, 0, copy); saveDashCharts(); renderDashboard();
      } else if (act === 'wide') {
        Dash.charts[i].wide = !Dash.charts[i].wide; saveDashCharts(); renderDashboard();
      }
    });
  });
}

function drawCustomChart(cfg) {
  const data = dashAggregate(cfg);
  const labels = data.map(d => d[0]);
  const values = data.map(d => d[1]);

  if (cfg.type === 'table') {
    const el = document.getElementById('dt-' + cfg.id);
    if (!el) return;
    const total = values.reduce((a, b) => a + b, 0);
    el.innerHTML = '<thead><tr><th>' + escapeHtml(cfg.dim) + '</th><th class="num">Qty</th><th class="num">Share</th></tr></thead>' +
      '<tbody>' + data.map(([k, v]) =>
        '<tr class="dash-trow" data-v="' + escapeHtml(k) + '"><td>' + escapeHtml(k) + '</td>' +
        '<td class="num">' + fmtNum(v) + '</td>' +
        '<td class="num">' + (total ? fmtNum(v / total * 100, 1) + '%' : '\u2014') + '</td></tr>').join('') +
      '</tbody><tfoot><tr><td>Total</td><td class="num">' + fmtNum(total) + '</td><td class="num">100%</td></tr></tfoot>';
    el.querySelectorAll('.dash-trow').forEach(tr => tr.addEventListener('click', () => addDashFilter(cfg.dim, tr.dataset.v)));
    return;
  }

  const canvas = document.getElementById('dc-' + cfg.id);
  if (!canvas || typeof Chart === 'undefined') return;
  const isPie = cfg.type === 'pie' || cfg.type === 'doughnut';
  const chartType = cfg.type === 'column' ? 'bar'
    : cfg.type === 'area' ? 'line'
    : cfg.type === 'bar' ? 'bar' : cfg.type;

  Dash.instances[cfg.id] = makeChart(canvas.getContext('2d'), {
    type: chartType,
    data: {
      labels: labels,
      datasets: [{
        label: cfg.title,
        data: values,
        backgroundColor: isPie ? labels.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) : CHART_COLORS[0],
        borderColor: cfg.type === 'line' || cfg.type === 'area' ? CHART_COLORS[0] : undefined,
        fill: cfg.type === 'area',
        tension: .25
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      indexAxis: cfg.type === 'bar' ? 'y' : 'x',
      onClick: (evt, els) => { if (els && els.length) addDashFilter(cfg.dim, labels[els[0].index]); },
      plugins: {
        legend: { display: isPie, position: 'right', labels: { font: { size: 10 }, boxWidth: 10 } },
        tooltip: { callbacks: { afterLabel: () => 'Click to filter the dashboard' } }
      },
      scales: isPie ? {} : { x: { ticks: { font: { size: 10 } } }, y: { beginAtZero: true } }
    }
  });
}

function addDashFilter(dim, value) {
  const i = Dash.filters.findIndex(f => f.dim === dim);
  if (i >= 0) {
    if (Dash.filters[i].value === value) Dash.filters.splice(i, 1); // click again to clear
    else Dash.filters[i].value = value;
  } else Dash.filters.push({ dim, value });
  renderDashboard();
}

/* ---- add / edit a chart ---- */
function openChartEditor(cfg) {
  const editing = !!cfg;
  const c = cfg || { id: 'c' + Date.now(), title: '', type: 'column', source: 'sales',
                     dim: 'Brand', measure: 'qty', topN: 10 };

  const pop = document.createElement('div');
  pop.className = 'modal-backdrop';
  pop.innerHTML = '<div class="modal-box chart-editor">' +
    '<h3>' + (editing ? 'Edit chart' : 'Add a chart') + '</h3>' +
    '<div class="settings-row"><label class="toolbar-label">Title</label>' +
      '<input type="text" id="ce-title" class="text-input grow" value="' + escapeHtml(c.title) + '" placeholder="Sales by brand"></div>' +
    '<div class="settings-row"><label class="toolbar-label">Chart type</label>' +
      '<select id="ce-type" class="select">' + CHART_TYPES.map(t =>
        '<option value="' + t[0] + '"' + (c.type === t[0] ? ' selected' : '') + '>' + t[1] + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Data</label>' +
      '<select id="ce-source" class="select">' + CHART_SOURCES.map(t =>
        '<option value="' + t[0] + '"' + (c.source === t[0] ? ' selected' : '') + '>' + t[1] + '</option>').join('') + '</select></div>' +
    '<div class="settings-row"><label class="toolbar-label">Group by</label>' +
      '<select id="ce-dim" class="select">' + CHART_DIMS.map(d =>
        '<option value="' + d + '"' + (c.dim === d ? ' selected' : '') + '>' + d + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Measure</label>' +
      '<select id="ce-measure" class="select">' + CHART_MEASURES.map(m =>
        '<option value="' + m[0] + '"' + (c.measure === m[0] ? ' selected' : '') + '>' + m[1] + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Top</label>' +
      '<input type="number" id="ce-topn" class="text-input narrow" min="3" max="50" value="' + (c.topN || 10) + '"></div>' +
    '<label class="toolbar-checkbox"><input type="checkbox" id="ce-wide"' + (c.wide ? ' checked' : '') + '> Full width</label>' +
    '<div class="modal-actions"><button class="ghost-btn primary small" id="ce-save">' +
      (editing ? 'Save chart' : 'Add chart') + '</button>' +
      '<span class="spacer"></span><button class="ghost-btn small" id="ce-cancel">Cancel</button></div>' +
  '</div>';
  document.body.appendChild(pop);

  pop.querySelector('#ce-cancel').onclick = () => pop.remove();
  pop.addEventListener('click', e => { if (e.target === pop) pop.remove(); });
  pop.querySelector('#ce-save').onclick = () => {
    c.title = pop.querySelector('#ce-title').value.trim() ||
      (pop.querySelector('#ce-source').selectedOptions[0].text + ' by ' + pop.querySelector('#ce-dim').value);
    c.type = pop.querySelector('#ce-type').value;
    c.source = pop.querySelector('#ce-source').value;
    c.dim = pop.querySelector('#ce-dim').value;
    c.measure = pop.querySelector('#ce-measure').value;
    c.topN = Math.max(3, Math.min(50, parseInt(pop.querySelector('#ce-topn').value, 10) || 10));
    c.wide = pop.querySelector('#ce-wide').checked;
    if (!editing) Dash.charts.push(c);
    saveDashCharts();
    pop.remove();
    renderDashboard();
    toast(editing ? 'Chart updated.' : 'Chart added.');
  };
}

/* ---------------------------------------------------------------
   16b. CHART BOARDS — move, resize and edit charts
   ---------------------------------------------------------------
   The same board runs on the Dashboard and on Product Performance.
   Cards can be dragged into a new order, resized from the corner,
   edited or removed, and everything is remembered. A lock switch
   in Settings freezes the layout so it cannot be changed by mistake.
   --------------------------------------------------------------- */

const BOARD_DEFS = {
  dash: { key: 'sl_dash_charts', grid: 'dash-grid', filterBar: 'dash-filters' },
  perf: { key: 'sl_perf_charts', grid: 'perf-grid', filterBar: 'perf-chart-filters' }
};

const PERF_DEFAULT_CHARTS = [
  { id: 'p1', title: 'Top sellers',          type: 'bar',      source: 'sales',    dim: 'Article No',  measure: 'qty', topN: 12, w: 520, h: 300 },
  { id: 'p2', title: 'Stock by sub section', type: 'bar',      source: 'stock',    dim: 'Sub Section', measure: 'qty', topN: 12, w: 520, h: 300 },
  { id: 'p3', title: 'Sales by month',       type: 'column',   source: 'sales',    dim: 'Month',       measure: 'qty', topN: 24, w: 520, h: 300 },
  { id: 'p4', title: 'Share by section',     type: 'doughnut', source: 'sales',    dim: 'Section',     measure: 'qty', topN: 8,  w: 520, h: 300 }
];

const Boards = {
  dash: { charts: [], filters: [], instances: {} },
  perf: { charts: [], filters: [], instances: {} },
  locked: false,
  drag: null
};

function loadBoards() {
  try { Boards.locked = Store.get('sl_boards_locked') === '1'; } catch (e) {}
  ['dash', 'perf'].forEach(id => {
    const def = id === 'dash' ? DEFAULT_CHARTS : PERF_DEFAULT_CHARTS;
    try {
      const raw = Store.get(BOARD_DEFS[id].key);
      Boards[id].charts = raw ? JSON.parse(raw) : def.slice();
    } catch (e) { Boards[id].charts = def.slice(); }
    if (!Boards[id].charts.length) Boards[id].charts = def.slice();
  });
}
function saveBoard(id) { Store.set(BOARD_DEFS[id].key, JSON.stringify(Boards[id].charts)); }
function setBoardsLocked(v) { Boards.locked = !!v; Store.set('sl_boards_locked', v ? '1' : '0'); renderAllBoards(); }

function renderAllBoards() {
  renderBoard('dash');
  renderBoard('perf');
}

/* ---- data for one chart, honouring that board's cross filters ---- */
function boardAggregate(boardId, cfg) {
  const filters = Boards[boardId].filters;
  let recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));
  const map = new Map(), seen = new Map();
  recs.forEach(r => {
    const k = dashDimKey(r, cfg.dim);
    if (cfg.measure === 'rows') map.set(k, (map.get(k) || 0) + 1);
    else if (cfg.measure === 'items') {
      let set = seen.get(k);
      if (!set) { set = new Set(); seen.set(k, set); }
      const code = r['Item Code'] || r['Article No'];
      if (code) set.add(String(code));
      map.set(k, set.size);
    } else map.set(k, (map.get(k) || 0) + recQty(r));
  });

  let entries = [...map.entries()].filter(([k]) => k !== '(blank)');
  const chrono = ['Month', 'Week', 'Transaction Date'].indexOf(cfg.dim) !== -1;
  if (chrono) {
    entries.sort((a, b) => cfg.dim === 'Month'
      ? grainSort(a[0], 'month') - grainSort(b[0], 'month')
      : (parseDateLoose(String(a[0]).split(' \u2013 ')[0]) || 0) - (parseDateLoose(String(b[0]).split(' \u2013 ')[0]) || 0));
    if (entries.length > (cfg.topN || 24)) entries = entries.slice(-(cfg.topN || 24));
  } else {
    entries.sort((a, b) => b[1] - a[1]);
    entries = entries.slice(0, cfg.topN || 10);
  }
  return entries;
}

function renderBoard(boardId) {
  const def = BOARD_DEFS[boardId];
  const grid = document.getElementById(def.grid);
  if (!grid) return;
  const B = Boards[boardId];

  Object.keys(B.instances).forEach(k => { try { B.instances[k].destroy(); } catch (e) {} });
  B.instances = {};

  renderBoardFilters(boardId);

  if (!App.datasets.length) {
    grid.innerHTML = '<div class="empty-hint big">Load your reports on the Import tab to build charts.</div>';
    return;
  }

  const locked = Boards.locked;
  grid.className = 'chart-board' + (locked ? ' locked' : '');
  grid.innerHTML = B.charts.map((cfg, i) =>
    '<div class="board-card" data-cid="' + cfg.id + '" data-idx="' + i + '"' +
      (locked ? '' : ' draggable="true"') +
      ' style="width:' + (cfg.w || 520) + 'px;height:' + (cfg.h || 300) + 'px">' +
      '<div class="board-head">' +
        (locked ? '' : '<span class="board-grip" title="Drag to move">\u2630</span>') +
        '<h3>' + escapeHtml(cfg.title) + '</h3>' +
        '<span class="board-meta">' + escapeHtml(cfg.source) + ' \u00b7 ' + escapeHtml(cfg.dim) + '</span>' +
        '<span class="spacer"></span>' +
        (locked ? '<span class="board-lockicon" title="Layout is locked in Settings">\uD83D\uDD12</span>' :
          '<button class="dash-ico" data-act="edit" title="Edit">\u270E</button>' +
          '<button class="dash-ico" data-act="dup" title="Duplicate">\u29C9</button>' +
          '<button class="dash-ico" data-act="del" title="Remove">\u2715</button>') +
      '</div>' +
      (cfg.type === 'table'
        ? '<div class="board-body"><table class="data-table dash-mini" id="bt-' + boardId + '-' + cfg.id + '"></table></div>'
        : '<div class="board-body"><canvas id="bc-' + boardId + '-' + cfg.id + '"></canvas></div>') +
      (locked ? '' : '<span class="board-resize" title="Drag to resize"></span>') +
    '</div>').join('');

  B.charts.forEach(cfg => drawBoardChart(boardId, cfg));
  wireBoardCard(boardId, grid);
}

function drawBoardChart(boardId, cfg) {
  const data = boardAggregate(boardId, cfg);
  const labels = data.map(d => d[0]);
  const values = data.map(d => d[1]);
  const B = Boards[boardId];

  if (cfg.type === 'table') {
    const el = document.getElementById('bt-' + boardId + '-' + cfg.id);
    if (!el) return;
    const total = values.reduce((a, b) => a + b, 0);
    el.innerHTML = '<thead><tr><th>' + escapeHtml(cfg.dim) + '</th><th class="num">Qty</th><th class="num">Share</th></tr></thead>' +
      '<tbody>' + data.map(([k, v]) =>
        '<tr class="dash-trow" data-v="' + escapeHtml(k) + '"><td>' + escapeHtml(k) + '</td>' +
        '<td class="num">' + fmtNum(v) + '</td><td class="num">' +
        (total ? fmtNum(v / total * 100, 1) + '%' : '\u2014') + '</td></tr>').join('') +
      '</tbody><tfoot><tr><td>Total</td><td class="num">' + fmtNum(total) + '</td><td class="num">100%</td></tr></tfoot>';
    el.querySelectorAll('.dash-trow').forEach(tr =>
      tr.addEventListener('click', () => addBoardFilter(boardId, cfg.dim, tr.dataset.v)));
    return;
  }

  const canvas = document.getElementById('bc-' + boardId + '-' + cfg.id);
  if (!canvas || typeof Chart === 'undefined') return;
  const isPie = cfg.type === 'pie' || cfg.type === 'doughnut';
  const chartType = cfg.type === 'column' ? 'bar' : cfg.type === 'area' ? 'line' : cfg.type;

  B.instances[cfg.id] = makeChart(canvas.getContext('2d'), {
    type: chartType,
    data: {
      labels: labels,
      datasets: [{
        label: cfg.title, data: values,
        backgroundColor: isPie ? labels.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) : CHART_COLORS[0],
        borderColor: (cfg.type === 'line' || cfg.type === 'area') ? CHART_COLORS[0] : undefined,
        fill: cfg.type === 'area', tension: .25
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      indexAxis: cfg.type === 'bar' ? 'y' : 'x',
      onClick: (evt, els) => { if (els && els.length) addBoardFilter(boardId, cfg.dim, labels[els[0].index]); },
      plugins: {
        legend: { display: isPie, position: 'right',
          labels: { font: { size: 10 }, boxWidth: 10, color: boardTextColour() } },
        tooltip: { callbacks: { afterLabel: () => 'Click to filter' } }
      },
      scales: isPie ? {} : {
        x: { ticks: { font: { size: 10 }, color: boardTextColour() }, grid: { color: boardGridColour() } },
        y: { beginAtZero: true, ticks: { color: boardTextColour() }, grid: { color: boardGridColour() } }
      }
    }
  });
}

function addBoardFilter(boardId, dim, value) {
  const F = Boards[boardId].filters;
  const i = F.findIndex(f => f.dim === dim);
  if (i >= 0) { if (F[i].value === value) F.splice(i, 1); else F[i].value = value; }
  else F.push({ dim, value });
  renderBoard(boardId);
}

function renderBoardFilters(boardId) {
  const wrap = document.getElementById(BOARD_DEFS[boardId].filterBar);
  if (!wrap) return;
  const F = Boards[boardId].filters;
  if (!F.length) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = '<span class="toolbar-label">Filtered by:</span>' +
    F.map((f, i) => '<span class="filter-chip">' + escapeHtml(f.dim) + ': <strong>' +
      escapeHtml(f.value) + '</strong><button data-fi="' + i + '">&times;</button></span>').join('') +
    '<button class="ghost-btn small" data-clear-all="1">Clear all</button>';
  wrap.querySelectorAll('[data-fi]').forEach(b => b.addEventListener('click', () => {
    F.splice(parseInt(b.dataset.fi, 10), 1); renderBoard(boardId);
  }));
  const ca = wrap.querySelector('[data-clear-all]');
  if (ca) ca.addEventListener('click', () => { Boards[boardId].filters = []; renderBoard(boardId); });
}

/* ---- move, resize, edit ---- */
function wireBoardCard(boardId, grid) {
  const B = Boards[boardId];

  grid.querySelectorAll('.dash-ico').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const card = btn.closest('.board-card');
      const i = B.charts.findIndex(c => c.id === card.dataset.cid);
      const act = btn.dataset.act;
      if (act === 'del') { B.charts.splice(i, 1); saveBoard(boardId); renderBoard(boardId); }
      else if (act === 'edit') openBoardChartEditor(boardId, B.charts[i]);
      else if (act === 'dup') {
        const copy = Object.assign({}, B.charts[i], { id: 'c' + Date.now(), title: B.charts[i].title + ' (copy)' });
        B.charts.splice(i + 1, 0, copy); saveBoard(boardId); renderBoard(boardId);
      }
    });
  });

  if (Boards.locked) return;

  // drag to reorder
  grid.querySelectorAll('.board-card').forEach(card => {
    card.addEventListener('dragstart', e => {
      Boards.drag = { boardId, id: card.dataset.cid };
      card.classList.add('dragging');
      try { e.dataTransfer.setData('text/plain', card.dataset.cid); } catch (err) {}
    });
    card.addEventListener('dragend', () => {
      card.classList.remove('dragging');
      grid.querySelectorAll('.board-card').forEach(c => c.classList.remove('drop-target'));
      Boards.drag = null;
    });
    card.addEventListener('dragover', e => {
      if (!Boards.drag || Boards.drag.boardId !== boardId) return;
      e.preventDefault();
      card.classList.add('drop-target');
    });
    card.addEventListener('dragleave', () => card.classList.remove('drop-target'));
    card.addEventListener('drop', e => {
      e.preventDefault();
      if (!Boards.drag || Boards.drag.boardId !== boardId) return;
      const from = B.charts.findIndex(c => c.id === Boards.drag.id);
      const to = B.charts.findIndex(c => c.id === card.dataset.cid);
      if (from < 0 || to < 0 || from === to) return;
      const [moved] = B.charts.splice(from, 1);
      B.charts.splice(to, 0, moved);
      saveBoard(boardId);
      renderBoard(boardId);
    });
  });

  // corner resize
  grid.querySelectorAll('.board-resize').forEach(handle => {
    handle.addEventListener('mousedown', e => {
      e.preventDefault(); e.stopPropagation();
      const card = handle.closest('.board-card');
      const cfg = B.charts.find(c => c.id === card.dataset.cid);
      const start = { x: e.clientX, y: e.clientY, w: card.offsetWidth, h: card.offsetHeight };

      const move = ev => {
        const w = Math.max(260, start.w + (ev.clientX - start.x));
        const h = Math.max(180, start.h + (ev.clientY - start.y));
        card.style.width = w + 'px';
        card.style.height = h + 'px';
        const inst = B.instances[cfg.id];
        if (inst && inst.resize) { try { inst.resize(); } catch (err) {} }
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        cfg.w = card.offsetWidth; cfg.h = card.offsetHeight;
        saveBoard(boardId);
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    });
  });
}

function openBoardChartEditor(boardId, cfg) {
  const editing = !!cfg;
  const B = Boards[boardId];
  const c = cfg || { id: 'c' + Date.now(), title: '', type: 'column', source: 'sales',
                     dim: 'Brand', measure: 'qty', topN: 10, w: 520, h: 300 };

  const pop = document.createElement('div');
  pop.className = 'modal-backdrop';
  pop.innerHTML = '<div class="modal-box chart-editor">' +
    '<h3>' + (editing ? 'Edit chart' : 'Add a chart') + '</h3>' +
    '<div class="settings-row"><label class="toolbar-label">Title</label>' +
      '<input type="text" id="ce-title" class="text-input grow" value="' + escapeHtml(c.title) + '" placeholder="Sales by brand"></div>' +
    '<div class="settings-row"><label class="toolbar-label">Chart type</label>' +
      '<select id="ce-type" class="select">' + CHART_TYPES.map(t =>
        '<option value="' + t[0] + '"' + (c.type === t[0] ? ' selected' : '') + '>' + t[1] + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Data</label>' +
      '<select id="ce-source" class="select">' + CHART_SOURCES.map(t =>
        '<option value="' + t[0] + '"' + (c.source === t[0] ? ' selected' : '') + '>' + t[1] + '</option>').join('') + '</select></div>' +
    '<div class="settings-row"><label class="toolbar-label">Group by</label>' +
      '<select id="ce-dim" class="select">' + CHART_DIMS.map(d =>
        '<option value="' + d + '"' + (c.dim === d ? ' selected' : '') + '>' + d + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Measure</label>' +
      '<select id="ce-measure" class="select">' + CHART_MEASURES.map(m =>
        '<option value="' + m[0] + '"' + (c.measure === m[0] ? ' selected' : '') + '>' + m[1] + '</option>').join('') + '</select>' +
      '<label class="toolbar-label">Top</label>' +
      '<input type="number" id="ce-topn" class="text-input narrow" min="3" max="50" value="' + (c.topN || 10) + '"></div>' +
    '<div class="settings-row"><label class="toolbar-label">Size</label>' +
      '<input type="number" id="ce-w" class="text-input narrow" min="260" step="20" value="' + (c.w || 520) + '"> \u00d7 ' +
      '<input type="number" id="ce-h" class="text-input narrow" min="180" step="20" value="' + (c.h || 300) + '"> px' +
      '<span class="drill-count">or drag the corner of the card</span></div>' +
    '<div class="modal-actions"><button class="ghost-btn primary small" id="ce-save">' +
      (editing ? 'Save chart' : 'Add chart') + '</button>' +
      '<span class="spacer"></span><button class="ghost-btn small" id="ce-cancel">Cancel</button></div>' +
    '</div>';
  document.body.appendChild(pop);

  pop.querySelector('#ce-cancel').onclick = () => pop.remove();
  pop.addEventListener('click', e => { if (e.target === pop) pop.remove(); });
  pop.querySelector('#ce-save').onclick = () => {
    c.title = pop.querySelector('#ce-title').value.trim() ||
      (pop.querySelector('#ce-source').selectedOptions[0].text + ' by ' + pop.querySelector('#ce-dim').value);
    c.type = pop.querySelector('#ce-type').value;
    c.source = pop.querySelector('#ce-source').value;
    c.dim = pop.querySelector('#ce-dim').value;
    c.measure = pop.querySelector('#ce-measure').value;
    c.topN = Math.max(3, Math.min(50, parseInt(pop.querySelector('#ce-topn').value, 10) || 10));
    c.w = Math.max(260, parseInt(pop.querySelector('#ce-w').value, 10) || 520);
    c.h = Math.max(180, parseInt(pop.querySelector('#ce-h').value, 10) || 300);
    if (!editing) B.charts.push(c);
    saveBoard(boardId);
    pop.remove();
    renderBoard(boardId);
    toast(editing ? 'Chart updated.' : 'Chart added.');
  };
}

function initBoards() {
  loadBoards();
  [['dash', 'dash-add-chart', 'dash-reset', 'dash-clear-filters', DEFAULT_CHARTS],
   ['perf', 'perf-add-chart', 'perf-reset', 'perf-clear-filters', PERF_DEFAULT_CHARTS]]
  .forEach(([id, addId, resetId, clearId, defaults]) => {
    const add = document.getElementById(addId);
    if (add) add.addEventListener('click', () => {
      if (Boards.locked) { toast('The layout is locked - unlock it in Settings.'); return; }
      openBoardChartEditor(id, null);
    });
    const reset = document.getElementById(resetId);
    if (reset) reset.addEventListener('click', () => {
      Boards[id].charts = defaults.slice();
      Boards[id].filters = [];
      saveBoard(id); renderBoard(id);
      toast('Charts reset to the standard set.');
    });
    const clear = document.getElementById(clearId);
    if (clear) clear.addEventListener('click', () => { Boards[id].filters = []; renderBoard(id); });
    const pres = document.getElementById(id + '-present');
    if (pres) pres.addEventListener('click', () => enterPresent(id));
  });
}

/* ---------------------------------------------------------------
   16c. PRESENT MODE + BOARD BACKGROUNDS
   ---------------------------------------------------------------
   Present: the board fills the screen with the sidebar and toolbars
   out of the way, one chart at a time or all together, and clicking
   still cross-filters. Esc leaves.
   Background: pick the canvas colour, the card colour and the text
   colour, or take one of the presets.
   --------------------------------------------------------------- */

const BOARD_PRESETS = [
  { id: 'paper',  name: 'Paper',        bg: '#F6F1E4', card: '#FFFDF8', fg: '#241C14', grid: '#E4DBC6' },
  { id: 'white',  name: 'Clean white',  bg: '#FFFFFF', card: '#FFFFFF', fg: '#1F2933', grid: '#E3E6EA' },
  { id: 'grey',   name: 'Soft grey',    bg: '#EEF1F4', card: '#FFFFFF', fg: '#1F2933', grid: '#D8DEE4' },
  { id: 'navy',   name: 'Dark navy',    bg: '#121B2A', card: '#1B2739', fg: '#E7EDF3', grid: '#2A3A50' },
  { id: 'char',   name: 'Charcoal',     bg: '#1C1C1E', card: '#2A2A2D', fg: '#EDEDEF', grid: '#3A3A3E' },
  { id: 'forest', name: 'Deep green',   bg: '#10231C', card: '#183129', fg: '#E4F0EA', grid: '#27453A' }
];

const BoardTheme = { bg: '#F6F1E4', card: '#FFFDF8', fg: '#241C14', grid: '#E4DBC6', preset: 'paper' };

function loadBoardTheme() {
  try {
    const raw = Store.get('sl_board_theme');
    if (raw) Object.assign(BoardTheme, JSON.parse(raw));
  } catch (e) {}
  applyBoardTheme();
}
function saveBoardTheme() { Store.set('sl_board_theme', JSON.stringify(BoardTheme)); applyBoardTheme(); }

function applyBoardTheme() {
  const r = document.documentElement;
  r.style.setProperty('--board-bg', BoardTheme.bg);
  r.style.setProperty('--board-card', BoardTheme.card);
  r.style.setProperty('--board-fg', BoardTheme.fg);
  r.style.setProperty('--board-grid', BoardTheme.grid || '#E4DBC6');
  // charts need to know whether they are on a dark canvas
  document.body.classList.toggle('board-dark', isDarkColour(BoardTheme.bg));
  ['dash', 'perf'].forEach(id => { if (Boards[id] && Boards[id].charts.length) renderBoard(id); });
}

function isDarkColour(hex) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
  if (!m) return false;
  const lum = (parseInt(m[1], 16) * 299 + parseInt(m[2], 16) * 587 + parseInt(m[3], 16) * 114) / 1000;
  return lum < 128;
}

/** Axis and legend colours follow the background so dark themes stay readable. */
function boardTextColour() { return BoardTheme.fg || '#241C14'; }
function boardGridColour() { return BoardTheme.grid || '#E4DBC6'; }

/* ---- present mode ---- */
const Present = { boardId: null, home: null, index: -1 };

function ensurePresentDom() {
  if (document.getElementById('present-overlay')) return;
  const d = document.createElement('div');
  d.id = 'present-overlay';
  d.className = 'present-overlay';
  d.style.display = 'none';
  d.innerHTML =
    '<div class="present-bar">' +
      '<span class="present-title" id="present-title"></span>' +
      '<span class="present-filters" id="present-filters"></span>' +
      '<span class="spacer"></span>' +
      '<button class="ghost-btn small" id="present-prev" title="Previous chart">\u2039</button>' +
      '<span class="present-count" id="present-count"></span>' +
      '<button class="ghost-btn small" id="present-next" title="Next chart">\u203A</button>' +
      '<button class="ghost-btn small" id="present-all" title="Show every chart">Show all</button>' +
      '<button class="ghost-btn small" id="present-full" title="Full screen">\u26F6</button>' +
      '<button class="ghost-btn small" id="present-exit" title="Leave (Esc)">\u2715 Exit</button>' +
    '</div>' +
    '<div class="present-stage" id="present-stage"></div>';
  document.body.appendChild(d);

  document.getElementById('present-exit').addEventListener('click', exitPresent);
  document.getElementById('present-prev').addEventListener('click', () => stepPresent(-1));
  document.getElementById('present-next').addEventListener('click', () => stepPresent(1));
  document.getElementById('present-all').addEventListener('click', () => { Present.index = -1; layoutPresent(); });
  document.getElementById('present-full').addEventListener('click', togglePresentFullscreen);

  document.addEventListener('keydown', e => {
    if (!Present.boardId) return;
    if (e.key === 'ArrowRight' || e.key === 'PageDown') { e.preventDefault(); stepPresent(1); }
    else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); stepPresent(-1); }
    else if (e.key === 'Home') { e.preventDefault(); Present.index = -1; layoutPresent(); }
  });
}

function enterPresent(boardId) {
  ensurePresentDom();
  const grid = document.getElementById(BOARD_DEFS[boardId].grid);
  if (!grid) return;

  Present.boardId = boardId;
  Present.home = grid.parentNode;
  Present.index = -1;

  document.getElementById('present-stage').appendChild(grid);
  document.getElementById('present-title').textContent =
    (boardId === 'dash' ? 'Dashboard' : 'Product Performance');
  document.getElementById('present-overlay').style.display = 'flex';
  document.body.classList.add('presenting');
  modalOpen('present-overlay', exitPresent);
  layoutPresent();
  togglePresentFullscreen(true);
}

function exitPresent() {
  const boardId = Present.boardId;
  if (!boardId) return;
  const grid = document.getElementById(BOARD_DEFS[boardId].grid);
  if (grid && Present.home) Present.home.appendChild(grid);
  document.getElementById('present-overlay').style.display = 'none';
  document.body.classList.remove('presenting');
  modalClose('present-overlay');
  if (document.fullscreenElement && document.exitFullscreen) {
    try { document.exitFullscreen(); } catch (e) {}
  }
  Present.boardId = null; Present.index = -1;
  if (grid) grid.classList.remove('present-single');
  renderBoard(boardId);
}

function stepPresent(dir) {
  const B = Boards[Present.boardId];
  if (!B || !B.charts.length) return;
  if (Present.index === -1) Present.index = dir > 0 ? 0 : B.charts.length - 1;
  else Present.index = (Present.index + dir + B.charts.length) % B.charts.length;
  layoutPresent();
}

/** Either every chart scaled up, or one chart filling the stage. */
function layoutPresent() {
  const boardId = Present.boardId;
  if (!boardId) return;
  const B = Boards[boardId];
  const grid = document.getElementById(BOARD_DEFS[boardId].grid);
  if (!grid) return;

  const single = Present.index >= 0;
  grid.classList.toggle('present-single', single);
  [...grid.querySelectorAll('.board-card')].forEach((card, i) => {
    card.style.display = (!single || i === Present.index) ? '' : 'none';
    card.style.width = ''; card.style.height = '';   // let CSS size them on stage
  });

  document.getElementById('present-count').textContent = single
    ? (Present.index + 1) + ' / ' + B.charts.length
    : B.charts.length + ' charts';

  const pf = document.getElementById('present-filters');
  pf.innerHTML = B.filters.length
    ? B.filters.map(f => '<span class="filter-chip">' + escapeHtml(f.dim) + ': <strong>' +
        escapeHtml(f.value) + '</strong></span>').join('')
    : '';

  Object.keys(B.instances).forEach(k => {
    const inst = B.instances[k];
    if (inst && inst.resize) { try { inst.resize(); } catch (e) {} }
  });
}

function togglePresentFullscreen(force) {
  const el = document.getElementById('present-overlay');
  if (!el) return;
  const on = !!document.fullscreenElement;
  if (force === true && on) return;
  if (!on && el.requestFullscreen) { try { el.requestFullscreen(); } catch (e) {} }
  else if (on && force !== true && document.exitFullscreen) { try { document.exitFullscreen(); } catch (e) {} }
}

/* ---- background settings UI ---- */
function renderBoardBackgroundSettings(wrap) {
  const html =
    '<h3 class="snap-set-title">Chart background</h3>' +
    '<p class="drill-subtitle">Applies to the Dashboard, the Product Performance charts and Present mode.</p>' +
    '<div class="swatch-grid">' +
      BOARD_PRESETS.map(p =>
        '<button class="swatch bg-preset' + (BoardTheme.preset === p.id ? ' active' : '') + '" data-preset="' + p.id + '">' +
        '<span style="background:' + p.bg + ';box-shadow:inset 0 0 0 1px rgba(0,0,0,.15)"></span>' + p.name + '</button>').join('') +
    '</div>' +
    '<div class="color-row" style="margin-top:10px;">' +
      '<label class="toolbar-label">Canvas</label>' +
      '<input type="color" id="bt-bg" value="' + BoardTheme.bg + '"><span class="hexcode">' + BoardTheme.bg + '</span>' +
      '<label class="toolbar-label">Card</label>' +
      '<input type="color" id="bt-card" value="' + BoardTheme.card + '"><span class="hexcode">' + BoardTheme.card + '</span>' +
    '</div>' +
    '<div class="color-row">' +
      '<label class="toolbar-label">Text</label>' +
      '<input type="color" id="bt-fg" value="' + BoardTheme.fg + '"><span class="hexcode">' + BoardTheme.fg + '</span>' +
      '<label class="toolbar-label">Grid lines</label>' +
      '<input type="color" id="bt-grid" value="' + (BoardTheme.grid || '#E4DBC6') + '"><span class="hexcode">' + (BoardTheme.grid || '#E4DBC6') + '</span>' +
    '</div>';

  const holder = document.createElement('div');
  holder.innerHTML = html;
  wrap.appendChild(holder);

  holder.querySelectorAll('.bg-preset').forEach(b => b.addEventListener('click', () => {
    const p = BOARD_PRESETS.find(x => x.id === b.dataset.preset);
    if (!p) return;
    BoardTheme.bg = p.bg; BoardTheme.card = p.card; BoardTheme.fg = p.fg;
    BoardTheme.grid = p.grid; BoardTheme.preset = p.id;
    saveBoardTheme();
    renderSettingsBody();
  }));
  [['bt-bg', 'bg'], ['bt-card', 'card'], ['bt-fg', 'fg'], ['bt-grid', 'grid']].forEach(([id, field]) => {
    const el = holder.querySelector('#' + id);
    if (el) el.addEventListener('input', e => {
      BoardTheme[field] = e.target.value; BoardTheme.preset = 'custom'; saveBoardTheme();
      const hx = e.target.nextElementSibling;
      if (hx) hx.textContent = e.target.value;
    });
  });
}

/* ---------------------------------------------------------------
   11. INIT
   --------------------------------------------------------------- */
const BUILD_VERSION = 'v31';

/** Ek init fail ho to baaki sab band na ho jaye — har step alag-alag chalta hai.
 *  Pehle ye sab ek hi try-block mein the, to koi ek element missing hone par
 *  uske baad ka saara setup (date range, session) chalta hi nahi tha. */
function safeInit(label, fn) {
  try { fn(); }
  catch (e) { console.error('StockLedger: "' + label + '" setup failed —', e); }
}

document.addEventListener('DOMContentLoaded', function () {
  const vb = document.getElementById('build-badge');
  if (vb) vb.textContent = 'build ' + BUILD_VERSION;

  safeInit('tabs', initTabs);
  safeInit('import', initImport);
  safeInit('sheets', initSheets);
  safeInit('explore', initExplore);
  safeInit('pivot', initPivot);
  safeInit('quick-report', initQuickReport);
  safeInit('insights', initInsights);
  safeInit('performance', initPerformance);
  safeInit('dashboard', initDashboard);
  safeInit('relations', initRelations);
  safeInit('drill', initDrill);
  safeInit('prefs', loadPrefs);
  safeInit('behaviour', loadBehaviour);
  safeInit('modal-stack', initModalStack);
  safeInit('settings', initSettings);
  safeInit('col-resize', initColResize);
  safeInit('catalog-prefs', loadCatPrefs);
  safeInit('catalog', initCatalog);
  safeInit('board-theme', loadBoardTheme);
  safeInit('boards', initBoards);
  safeInit('prefs-apply', applyPrefsToControls);
  safeInit('snapshot', initSnapshot);
  safeInit('session', initSession);
  safeInit('period-selects', wirePeriodSelects);
  safeInit('gs-buttons', updateGsOnlyButtons);
  safeInit('dashboard-render', renderDashboard);
  safeInit('performance-render', renderPerformance);
  safeInit('relations-render', renderRelations);
  safeInit('restore-persisted', restorePersistedDatasets);
});

})();
