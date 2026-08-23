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
  'Quantity', 'Price', 'Amount', 'HSN Code', 'City', 'Discount',
  'Purchase Bill Date'
];

const FIELD_KIND = {
  'Date': 'date', 'Purchase Bill Date': 'date',
  'Quantity': 'number', 'Price': 'number', 'Amount': 'number'
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
  'transactionquantity': 'Quantity', 'quantity': 'Quantity', 'qty': 'Quantity', 'cbsqty': 'Quantity',
  'closingqty': 'Quantity', 'closingstock': 'Quantity', 'stockqty': 'Quantity', 'balanceqty': 'Quantity',
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
  if (headers.some(h => h.includes('cbsqty') || h.includes('closingstock') || h.includes('balanceqty'))) return 'stock';
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
  if (file.size === 0) { toast(file.name + ': file khaali hai (0 bytes) — dobara download karke try karo.'); return; }
  if (file.size > 80 * 1024 * 1024) { toast(file.name + ': file bahut badi hai (' + (file.size / 1024 / 1024).toFixed(0) + ' MB) — browser mein load karna mushkil ho sakta hai.'); }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    // Library abhi tak load na hui ho to pehle usko laate hain, phir parse.
    ensureXLSX().then(function (ok) {
      if (!ok) {
        showLibError('Excel reader library load nahi ho payi. Internet connection, ad-blocker ' +
          'ya office firewall check karo — inke bina file padhna kaam nahi karega.');
        return;
      }
      parseWorkbookBuffer(data, file);
    });
  };
  reader.onerror = () => toast(file.name + ': browser file read nahi kar paya (' + (reader.error ? reader.error.name : 'unknown') + ').');
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
    showLibError('Excel reader library load nahi ho payi. Page refresh karo; phir bhi na chale to ad-blocker/firewall check karo.');
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
      lastErr = new Error('File padhi gayi lekin koi sheet nahi mili.');
    } catch (err) {
      lastErr = err;
    }
  }

  console.error('StockLedger: could not parse', file.name, lastErr);
  toast(file.name + ': read nahi ho payi — ' + (lastErr ? lastErr.message : 'unknown error') +
    '. Excel mein khol kar "Save As → .xlsx" karke dobara try karo.');
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
  renderImportCard(sourceName, sheetName, columns, dataRows, guessedType, origin, headerIdx, droppedCount);
}

function renderImportCard(filename, sheetName, columns, dataRows, guessedType, origin, headerIdx, droppedCount) {
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
      '<span>' + dataRows.length.toLocaleString('en-IN') + ' rows detected' +
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
    confirmImport(card, filename, columns, dataRows, origin, headerIdx);
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

function confirmImport(card, filename, columns, dataRows, origin, headerIdx) {
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
    headerIdx: headerIdx || 0
  };
  App.datasets.push(ds);
  card.remove();
  toast('Added "' + name + '" — ' + records.length.toLocaleString('en-IN') + ' rows.');
  refreshAfterDataChange();
}

function refreshAfterDataChange() {
  clearLookups();
  // dusri file aate hi connections khud detect kar lete hain
  if (App.datasets.length > 1 && !App.relationships.length) autoDetectRelationships(true);
  rescoreRelationships();
  renderSidebarDatasets();
  renderLoadedTable();
  populateDatasetSelects();
  renderExplore();
  renderPivotFieldList();
  computePivot();
  renderDashboard();
  renderInsights();
  renderPerformance();
  renderRelations();
  if (Drill.open) renderDrill();
}

function removeDataset(id) {
  App.datasets = App.datasets.filter(d => d.id !== id);
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
  ['explore-dataset-select', 'pivot-dataset-select'].forEach(id => {
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
  document.getElementById('insights-to-sheet').addEventListener('click', insightsToSheet);
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
        throw new Error('Script ne JSON ke bajaye HTML bheja — deployment "Who has access: Anyone" par set hai kya? URL /exec par khatam hona chahiye.');
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
  if (!url) { setGsStatus('Web app URL daalo.', 'err'); return; }
  if (!/^https:\/\/script\.google\.com\/.*\/exec/.test(url)) {
    setGsStatus('URL https://script.google.com/... /exec jaisa hona chahiye.', 'err');
    return;
  }
  if (!key) { setGsStatus('API key daalo.', 'err'); return; }

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
  if (!meta.sheets.length) { wrap.innerHTML = '<div class="empty-hint">Is spreadsheet mein koi visible sheet nahi mili.</div>'; return; }

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
    if (!all.length) { toast(sheetName + ': koi row nahi mili.'); return; }

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
  if (!grid) { toast('Pehle ek pivot banao.'); return; }
  const sheetName = prompt('Kis sheet mein likhein? (na ho to ban jayegi)', 'StockLedger Pivot');
  if (!sheetName) return;
  toast('Writing to sheet…');
  gsPost({ action: 'write', sheet: sheetName, values: [grid.headers].concat(grid.rows), mode: 'replace' })
    .then(res => toast('Ho gaya — "' + res.sheet + '" mein ' + res.rowsWritten + ' rows likhi gayi.'))
    .catch(err => toast('Write failed: ' + err.message));
}

function insightsToSheet() {
  const grid = insightsToGrid();
  if (!grid) { toast('Abhi kuch export karne ko nahi hai.'); return; }
  const sheetName = prompt('Kis sheet mein likhein? (na ho to ban jayegi)', 'StockLedger Reorder');
  if (!sheetName) return;
  toast('Writing to sheet…');
  gsPost({ action: 'write', sheet: sheetName, values: [grid.headers].concat(grid.rows), mode: 'replace' })
    .then(res => toast('Ho gaya — "' + res.sheet + '" mein ' + res.rowsWritten + ' rows likhi gayi.'))
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
  if (!fields.length) { toast('Pehle koi file load karo.'); return; }
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
      (counts.size > 500 ? '<div class="empty-hint">' + counts.size.toLocaleString('en-IN') + ' unique values — top 500 dikha rahe hain.</div>' : '');
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

  table.innerHTML = thead + tbody;
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

function dataAnchorDate() {
  const dates = salesRecords().concat(purchaseRecords()).map(r => r.Date).filter(Boolean);
  if (!dates.length) return new Date();
  return new Date(Math.max(...dates));
}

function periodRange() {
  const p = App.period;
  const anchor = dataAnchorDate();
  const endOfAnchor = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), anchor.getUTCDate()));
  if (p.mode === 'all') return { from: null, to: null, label: 'All data' };
  if (p.mode === 'custom') {
    return { from: p.from ? parseDateLoose(p.from) : null, to: p.to ? parseDateLoose(p.to) : null, label: 'Custom range' };
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
  const min = new Date(Math.min(...dates)), max = new Date(Math.max(...dates));
  return Math.max(1, Math.round((max - min) / 86400000) + 1);
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
  if (!silent) toast(added ? added + ' connection(s) detect hui.' : 'Koi nayi connection nahi mili.');
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
      s = { key, sold: 0, purchased: 0, stock: 0, saleLines: 0,
            firstSale: null, lastSale: null, oldestStock: null, newestStock: null,
            meta: {} };
      map.set(key, s);
    }
    return s;
  }

  sales.forEach(r => {
    const s = slot(dimKey(r, dim));
    const q = typeof r.Quantity === 'number' ? r.Quantity : 0;
    s.sold += q; s.saleLines++;
    if (r.Date) {
      if (!s.firstSale || r.Date < s.firstSale) s.firstSale = r.Date;
      if (!s.lastSale || r.Date > s.lastSale) s.lastSale = r.Date;
    }
    captureMeta(s, r);
  });

  purchases.forEach(r => {
    const s = slot(dimKey(r, dim));
    s.purchased += typeof r.Quantity === 'number' ? r.Quantity : 0;
    captureMeta(s, r);
  });

  stock.forEach(r => {
    const s = slot(dimKey(r, dim));
    const q = typeof r.Quantity === 'number' ? r.Quantity : 0;
    s.stock += q;
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
    return Object.assign({}, s, { avgDaily, daysCover, suggested, sellThrough, daysSinceLastSale, stockAgeDays, oldestAgeDays });
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
                ['thismonth', 'Latest month'], ['thisyear', 'Latest year']];
  return '<select id="' + id + '" class="select period-select">' +
    opts.map(([v, l]) => '<option value="' + v + '"' + (p === v ? ' selected' : '') + '>' + l + '</option>').join('') +
    '</select>';
}

function wirePeriodSelects() {
  document.querySelectorAll('.period-select').forEach(sel => {
    sel.addEventListener('change', () => {
      App.period.mode = sel.value;
      document.querySelectorAll('.period-select').forEach(s => { s.value = sel.value; });
      renderInsights();
      renderPerformance();
      renderDashboard();
    });
  });
}

/* ---------------------------------------------------------------
   8. REORDER PLANNER
   --------------------------------------------------------------- */
function initInsights() {
  document.getElementById('insights-groupby').addEventListener('change', renderInsights);
  document.getElementById('insights-target-days').addEventListener('input', debounce(renderInsights, 250));
  document.getElementById('insights-only-flagged').addEventListener('change', renderInsights);
  document.getElementById('insights-export').addEventListener('click', exportInsightsCSV);
}

function aggregateByDimension(records, dim) {
  const map = new Map();
  records.forEach(r => {
    const key = dimKey(r, dim);
    map.set(key, (map.get(key) || 0) + (typeof r.Quantity === 'number' ? r.Quantity : 0));
  });
  return map;
}

let lastInsightsRows = null;

function renderInsights() {
  const wrap = document.getElementById('insights-table');
  const kpiWrap = document.getElementById('insights-kpis');
  const note = document.getElementById('insights-missing-note');
  const dim = document.getElementById('insights-groupby').value;
  const targetDays = Math.max(1, parseInt(document.getElementById('insights-target-days').value, 10) || 30);
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

  kpiWrap.innerHTML = [
    ['Analysis window', A.range.label, A.days + ' days · upto ' + fmtDate(A.anchor)],
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

  wrap.innerHTML = head + body;
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
const PerfState = { view: 'all', search: '', sortKey: 'sold', sortDir: -1 };
let perfCharts = {};
let lastPerfRows = null;

function initPerformance() {
  document.getElementById('perf-groupby').addEventListener('change', renderPerformance);
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
  const kpiWrap = document.getElementById('perf-kpis');
  const chartsWrap = document.getElementById('perf-charts');

  if (!App.datasets.length) {
    tableEl.innerHTML = '<tr><td class="empty-hint">Load your data on the Import tab first.</td></tr>';
    kpiWrap.innerHTML = ''; chartsWrap.style.display = 'none'; lastPerfRows = null;
    return;
  }
  chartsWrap.style.display = '';

  const targetDays = Math.max(1, parseInt(document.getElementById('insights-target-days').value, 10) || 30);
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

  kpiWrap.innerHTML = [
    ['Window', A.range.label, A.days + ' days · upto ' + fmtDate(A.anchor)],
    ['Best sellers', best.length.toLocaleString('en-IN'), 'A-class, 80% of sales'],
    ['Dead / non-moving', dead.length.toLocaleString('en-IN'), fmtNum(deadQty) + ' pcs stuck'],
    ['Overstocked', over.length.toLocaleString('en-IN'), fmtNum(excessQty) + ' pcs excess'],
    ['Out of stock', oos.length.toLocaleString('en-IN'), 'sold but zero balance'],
    ['Overall sell-through', fmtNum(overallST, 1) + '%', 'sold ÷ (sold + stock)']
  ].map(([label, value, sub]) => '<div class="kpi-card"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>').join('');

  renderParetoChart(A.rows);
  renderStatusChart(A.rows);
  renderAgeingChart(A.rows);
  renderBottomChart(A.rows);

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

  const cols = [
    ['key', dim, false], ['sold', 'Sold', true], ['stock', 'Stock', true],
    ['sellThrough', 'Sell-through', true], ['daysCover', 'Days cover', true],
    ['lastSale', 'Last sold', false], ['daysSinceLastSale', 'Days since', true],
    ['stockAgeDays', 'Stock age', true], ['excessQty', 'Excess', true],
    ['abc', 'ABC', false], ['status', 'Status', false]
  ];
  const head = '<thead><tr>' + cols.map(([k, label, isNum]) =>
    '<th data-key="' + k + '" class="' + (isNum ? 'num' : '') + '">' + label +
    (sk === k ? '<span class="sort-arrow">' + (sd === 1 ? '▲' : '▼') + '</span>' : '') + '</th>').join('') + '</tr></thead>';

  const body = '<tbody>' + rows.slice(0, 800).map(r =>
    '<tr class="drillable" data-key="' + escapeHtml(r.key) + '">' +
      '<td title="' + escapeHtml(Object.values(r.meta).join(' · ')) + '">' + escapeHtml(r.key) + ' <span class="drill-hint">▸</span></td>' +
      '<td class="num">' + fmtNum(r.sold) + '</td>' +
      '<td class="num">' + fmtNum(r.stock) + '</td>' +
      '<td class="num">' + fmtNum(r.sellThrough, 1) + '%</td>' +
      '<td class="num">' + (r.daysCover === Infinity ? '∞' : fmtNum(r.daysCover, 0)) + '</td>' +
      '<td>' + (r.lastSale ? fmtDate(r.lastSale) : '—') + '</td>' +
      '<td class="num">' + (r.daysSinceLastSale === null ? '—' : r.daysSinceLastSale) + '</td>' +
      '<td class="num">' + (r.stockAgeDays === null ? '—' : r.stockAgeDays) + '</td>' +
      '<td class="num">' + (r.excessQty ? fmtNum(r.excessQty) : '—') + '</td>' +
      '<td><span class="abc-tag abc-' + r.abc + '">' + r.abc + '</span></td>' +
      '<td><span class="status-tag st-' + r.status.replace(/\s+/g, '-').toLowerCase() + '">' + r.status + '</span></td>' +
    '</tr>').join('') + '</tbody>';

  tableEl.innerHTML = head + body;
  tableEl.querySelectorAll('thead th').forEach(th => th.addEventListener('click', () => {
    const k = th.dataset.key;
    if (PerfState.sortKey === k) PerfState.sortDir *= -1;
    else { PerfState.sortKey = k; PerfState.sortDir = (k === 'key' || k === 'status' || k === 'abc') ? 1 : -1; }
    renderPerformance();
  }));
  tableEl.querySelectorAll('tbody tr').forEach(tr => tr.addEventListener('click', () => {
    if (tr.dataset.key) openDrill(dim, tr.dataset.key);
  }));

  document.getElementById('perf-count').textContent =
    rows.length.toLocaleString('en-IN') + ' rows' + (rows.length > 800 ? ' (showing first 800 — export for full list)' : '');
}

function destroyPerfChart(id) { if (perfCharts[id]) { perfCharts[id].destroy(); delete perfCharts[id]; } }

function renderParetoChart(rows) {
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
  document.getElementById('drill-overlay').style.display = 'none';
  if (drillChart) { drillChart.destroy(); drillChart = null; }
}

function initDrill() {
  document.getElementById('drill-close').addEventListener('click', closeDrill);
  document.getElementById('drill-overlay').addEventListener('click', e => {
    if (e.target.id === 'drill-overlay') closeDrill();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && Drill.open) closeDrill(); });
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

  const sumQ = rs => rs.reduce((s, r) => s + (typeof r.Quantity === 'number' ? r.Quantity : 0), 0);
  const sold = sumQ(sales), purchased = sumQ(purch), inStock = sumQ(stock);
  const sellThrough = (sold + inStock) > 0 ? (sold / (sold + inStock)) * 100 : 0;
  const avgDaily = sold / days;
  const daysCover = avgDaily > 0 ? inStock / avgDaily : (inStock > 0 ? Infinity : 0);
  const dates = sales.map(r => r.Date).filter(Boolean);
  const lastSale = dates.length ? new Date(Math.max(...dates)) : null;

  // title + breadcrumb
  const last = Drill.filters[Drill.filters.length - 1];
  document.getElementById('drill-title').textContent = last ? last.value : 'Details';
  document.getElementById('drill-subtitle').textContent =
    Drill.filters.map(f => f.field + ': ' + f.value).join('  ·  ');

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
    m.set(k, (m.get(k) || 0) + (typeof r.Quantity === 'number' ? r.Quantity : 0));
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
    s.sold += typeof r.Quantity === 'number' ? r.Quantity : 0;
    if (r.Date && (!s.lastSale || r.Date > s.lastSale)) s.lastSale = r.Date;
  });
  purch.forEach(r => { slot(keyOf(r)).purchased += typeof r.Quantity === 'number' ? r.Quantity : 0; });
  stock.forEach(r => { slot(keyOf(r)).stock += typeof r.Quantity === 'number' ? r.Quantity : 0; });

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
      '</tr>').join('') + '</tbody>';

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
    noteEl.textContent = 'Dhyan do: is selection ki ' + fmtNum(blankRow.sold / totalSold * 100, 0) +
      '% sale rows mein "' + dim + '" khaali hai, isliye ye breakdown adhoora hai. ' +
      'ERP mein ' + dim + ' bharna shuru karoge to ye analysis kaam karega.';
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
  if (!lastDrillRows || !lastDrillRows.rows.length) { toast('Kuch export karne ko nahi hai.'); return; }
  const { rows, dim } = lastDrillRows;
  const headers = [dim, 'Sold Qty', 'Purchased Qty', 'Stock Qty', 'Sell-through %', 'Days Cover', 'Last Sold'];
  const data = rows.map(r => [r.key, r.sold, r.purchased, r.stock, Number(r.sellThrough.toFixed(1)),
    r.daysCover === Infinity ? '' : Math.round(r.daysCover), r.lastSale ? fmtDate(r.lastSale) : '']);
  const ctx = Drill.filters.map(f => f.field + '-' + f.value).join('_').replace(/[^\w-]+/g, '');
  downloadBlob(toCSV(headers, data), 'drill-' + ctx.slice(0, 60) + '.csv', 'text/csv');
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
  const empty = document.getElementById('dashboard-empty');
  const body = document.getElementById('dashboard-body');
  if (!App.datasets.length) { empty.style.display = ''; body.style.display = 'none'; return; }
  empty.style.display = 'none'; body.style.display = '';

  const range = periodRange();
  const salesRecs = salesRecords().filter(r => inPeriod(r, range));
  const purchaseRecs = purchaseRecords().filter(r => inPeriod(r, range));
  const stockRecs = stockRecords();

  const sumQty = recs => recs.reduce((s, r) => s + (typeof r.Quantity === 'number' ? r.Quantity : 0), 0);
  const distinctItems = recs => new Set(recs.map(r => r['Item Code']).filter(Boolean)).size;

  const soldQty = sumQty(salesRecs), stockQty = sumQty(stockRecs);
  const sellThrough = (soldQty + stockQty) > 0 ? (soldQty / (soldQty + stockQty)) * 100 : 0;

  const A = App.datasets.length ? buildAnalysis('Item Code', 30) : null;
  const deadCount = A ? A.rows.filter(r => r.status === 'Dead stock').length : 0;

  document.getElementById('dashboard-kpis').innerHTML = [
    ['Sold qty', fmtNum(soldQty), salesRecs.length ? salesRecs.length.toLocaleString('en-IN') + ' bill lines' : 'no sales data'],
    ['Purchased qty', fmtNum(sumQty(purchaseRecs)), purchaseRecs.length ? purchaseRecs.length.toLocaleString('en-IN') + ' lines' : 'no purchase data'],
    ['Stock on hand', fmtNum(stockQty), stockRecs.length ? distinctItems(stockRecs).toLocaleString('en-IN') + ' SKUs' : 'no stock data'],
    ['Sell-through', fmtNum(sellThrough, 1) + '%', 'sold ÷ (sold + stock)'],
    ['Non-moving items', deadCount.toLocaleString('en-IN'), 'no sale in 90+ days']
  ].map(([label, value, sub]) => '<div class="kpi-card"><div class="kpi-label">' + label + '</div><div class="kpi-value">' + value + '</div><div class="kpi-sub">' + sub + '</div></div>').join('');

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

function renderTrendChart(salesRecs, purchaseRecs) {
  destroyChart('chart-trend');
  const grain = DashState.grain;
  const byGrain = recs => {
    const m = new Map();
    recs.forEach(r => {
      if (!r.Date) return;
      const key = grainKey(r.Date, grain);
      m.set(key, (m.get(key) || 0) + (typeof r.Quantity === 'number' ? r.Quantity : 0));
    });
    return m;
  };
  const salesM = byGrain(salesRecs), purchM = byGrain(purchaseRecs);
  const allKeys = [...new Set([...salesM.keys(), ...purchM.keys()])]
    .sort((a, b) => grainSort(a, grain) - grainSort(b, grain));

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
  const map = aggregateByDimension(stockRecs, 'Section');
  const top = [...map.entries()].filter(([k]) => k !== '(blank)').sort((a, b) => b[1] - a[1]).slice(0, 8);
  dashCharts['chart-stocksplit'] = makeChart(el.getContext('2d'), {
    type: 'doughnut',
    data: { labels: top.map(t => t[0]), datasets: [{ data: top.map(t => t[1]), backgroundColor: top.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]) }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12 } } } }
  });
}

/* ---------------------------------------------------------------
   9b. CONNECTIONS — Power BI jaisa wire-based data model
   --------------------------------------------------------------- */
const RelUI = { positions: {}, dragNode: null, dragPort: null, tempLine: null, selected: null };

function initRelations() {
  document.getElementById('rel-autodetect').addEventListener('click', () => {
    autoDetectRelationships();
    renderRelations();
    refreshAnalysisViews();
  });
  document.getElementById('rel-clear').addEventListener('click', () => {
    if (!App.relationships.length) { toast('Koi connection nahi hai.'); return; }
    App.relationships = [];
    clearLookups();
    renderRelations();
    refreshAnalysisViews();
    toast('Sab connections hata di gayi.');
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
  const canvas = document.getElementById('rel-canvas');
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
      if (from.dsId === to.dsId) { toast('Ek hi file ke do columns ko jodna kaam nahi karega.'); return; }
      addRelationship(from.dsId, from.field, to.dsId, to.field);
      e.stopPropagation();
    });
  });
}

function addRelationship(fromDsId, fromField, toDsId, toField) {
  const id = relationshipId(fromDsId, fromField, toDsId, toField);
  const rev = relationshipId(toDsId, toField, fromDsId, fromField);
  if (App.relationships.some(r => r.id === id || r.id === rev)) { toast('Ye connection pehle se hai.'); return; }
  const a = App.datasets.find(d => d.id === fromDsId), b = App.datasets.find(d => d.id === toDsId);
  const score = scoreRelationship(a, fromField, b, toField);
  App.relationships.push({ id, fromDsId, fromField, toDsId, toField, enabled: true, score });
  clearLookups();
  renderRelations();
  refreshAnalysisViews();
  toast('Connection bani: ' + fromField + ' ↔ ' + toField + ' (' + fmtNum(score.pct, 0) + '% match)');
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
    wrap.innerHTML = '<div class="empty-hint">Abhi koi connection nahi. "Auto-detect" dabao, ya kisi column ke gol point se drag karke doosri file ke column par chhodo.</div>';
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
            ? '<span class="match-pill ok" title="Ye ek asli key hai — iske through doosri file ke columns bhi use ho sakte hain">Yes</span>'
            : '<span class="match-pill warn" title="Values unique nahi hain, isliye sirf samajhne ke liye — column lookup band hai">No</span>') + '</td>' +
        '<td><input type="checkbox" class="rel-toggle" data-id="' + rel.id + '"' + (rel.enabled ? ' checked' : '') + '></td>' +
        '<td><button class="ghost-btn small rel-del" data-id="' + rel.id + '">Remove</button></td>' +
      '</tr>';
    }).join('') + '</tbody></table>' +
    '<p class="rel-help"><strong>Match %</strong> = pehli file ki kitni rows ko doosri file mein jodne wali value mili. ' +
    '<strong>Unique keys</strong> = target column mein kitni alag values hain. ' +
    '<strong>Can look up</strong> = kya iske through doosri file ka column is file ke analysis mein use ho sakta hai. ' +
    'Ye sirf tab "Yes" hota hai jab values lagbhag unique hon (jaise Item Code / barcode). ' +
    'Section jaisa column 100% match dikhata hai par uske through value uthana galat hoga — ek Section mein hazaron alag items hote hain.</p>';

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
   11. INIT
   --------------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', function () {
  initTabs();
  initImport();
  initSheets();
  initExplore();
  initPivot();
  initInsights();
  initPerformance();
  initDashboard();
  initRelations();
  initDrill();
  initSession();
  wirePeriodSelects();
  updateGsOnlyButtons();
  renderDashboard();
  renderPerformance();
  renderRelations();
});

})();
