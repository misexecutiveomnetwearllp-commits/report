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
    // "Total", "Grand Total", and ERP subtotal labels like "Retail Sales Total"
    // or "Purchase Total" - anything whose label ends in the word Total.
    if (/(^|\s)total$/i.test(s)) hasTotalLabel = true;
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
  set(k, v) {
    let ok = false;
    try { window.localStorage.setItem(k, v); ok = true; } catch (e) {}
    // The sheet keeps the master copy so every browser agrees. Local first,
    // so this never blocks anything; see the settings sync section below.
    try {
      if (typeof SYNC_KEYS !== 'undefined' && SYNC_KEYS.indexOf(k) !== -1 &&
          !Sync.suspend && GS.url && GS.key) pushSettings();
    } catch (e) {}
    return ok;
  },
  remove(k) {
    try { window.localStorage.removeItem(k); } catch (e) {}
    try {
      if (typeof SYNC_KEYS !== 'undefined' && SYNC_KEYS.indexOf(k) !== -1 &&
          !Sync.suspend && GS.url && GS.key) pushSettings();
    } catch (e) {}
  }
};

function debounce(fn, ms) {
  let t;
  return function (...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), ms); };
}


/* ---------------------------------------------------------------
   3b. SETTINGS SYNC — the same setup in every browser
   ---------------------------------------------------------------
   Settings used to live only in localStorage, which is per browser
   and per device. Open the site on a laptop and a phone and you
   got two different dashboards, two different column choices, two
   of everything.

   Now the Google Sheet holds the master copy. The flow is
   local-first, so nothing gets slower and the site still works
   with no connection:

     1. Start from localStorage, exactly as before.
     2. In the background, ask the sheet for its copy.
     3. If the sheet is newer, apply it and redraw.
     4. Whenever a setting changes, push the whole set back
        (debounced, so a slider drag is one save, not fifty).

   The sheet connection details themselves are never synced - they
   are what makes the connection possible, and they are closer to a
   password than a preference.
   --------------------------------------------------------------- */

const SYNC_KEYS = [
  'sl_prefs', 'sl_theme', 'sl_behaviour', 'sl_catprefs', 'sl_colwidths',
  'sl_dash_charts', 'sl_perf_charts', 'sl_board_theme', 'sl_boards_locked',
  'sl_replen', 'sl_replen_bulk', 'sl_snapshot_config', 'sl_cat_height'
];

const Sync = {
  on: false,          // sheet says settings sync is available
  pulling: false,
  lastPush: 0,
  lastPulled: null,   // the stamp the sheet reported
  suspend: false      // true while we are applying a pull, so it cannot echo back
};

/** A friendly name for this browser, so the sheet shows where a change
 *  came from. Nothing identifying, just the browser and platform. */
function deviceLabel() {
  const ua = navigator.userAgent || '';
  const browser = /Edg\//.test(ua) ? 'Edge' : /OPR\//.test(ua) ? 'Opera'
    : /Chrome\//.test(ua) ? 'Chrome' : /Safari\//.test(ua) ? 'Safari'
    : /Firefox\//.test(ua) ? 'Firefox' : 'Browser';
  const plat = /Android/.test(ua) ? 'Android' : /iPhone|iPad/.test(ua) ? 'iOS'
    : /Windows/.test(ua) ? 'Windows' : /Mac OS/.test(ua) ? 'Mac'
    : /Linux/.test(ua) ? 'Linux' : '';
  return (browser + (plat ? ' on ' + plat : '')).trim();
}

/** Everything worth carrying between browsers, as one object. */
function collectSettings() {
  const out = {};
  SYNC_KEYS.forEach(k => {
    const v = Store.get(k);
    if (v !== null && v !== undefined) out[k] = v;
  });
  return out;
}

/** Writes a pulled set into localStorage and reloads everything from it. */
function applySettings(obj) {
  if (!obj || typeof obj !== 'object') return false;
  let changed = 0;
  Sync.suspend = true;
  SYNC_KEYS.forEach(k => {
    if (!(k in obj)) return;
    const next = obj[k];
    if (typeof next !== 'string') return;
    if (Store.get(k) === next) return;
    Store.set(k, next);
    changed++;
  });
  Sync.suspend = false;
  if (!changed) return false;

  // read it all back in and repaint
  try {
    loadPrefs(); loadTheme(); loadBehaviour(); loadCatPrefs(); loadColWidths();
    loadReplen(); loadSnapshotConfig(); loadBoards(); loadBoardTheme();
    applyPrefsToControls();
    renderDashboard(); renderPerformance(); renderCatalog();
  } catch (e) {
    console.error('Could not apply the synced settings', e);
  }
  return true;
}

/** Asks the sheet for its copy. Quiet on failure - the site keeps working
 *  from the local copy, which is the whole point of local-first. */
function pullSettings(announce) {
  if (!GS.url || !GS.key || Sync.pulling) return Promise.resolve(false);
  Sync.pulling = true;
  return gsGet({ action: 'settings' })
    .then(res => {
      Sync.on = true;
      const blob = res.settings && res.settings.all;
      Sync.lastPulled = res.updated || null;
      if (!blob) { setSyncNote('Nothing saved in the sheet yet.'); return false; }
      const applied = applySettings(blob);
      setSyncNote(applied
        ? 'Settings loaded from the sheet' + (res.updated ? ' \u00b7 saved ' + fmtWhen(res.updated) : '') + '.'
        : 'Already up to date with the sheet.');
      if (applied && announce) toast('Settings restored from your Google Sheet.');
      return applied;
    })
    .catch(err => {
      Sync.on = false;
      setSyncNote('Could not read settings from the sheet: ' + err.message);
      return false;
    })
    .then(v => { Sync.pulling = false; return v; });
}

/** Sends the whole set back. Debounced by the caller. */
function pushSettingsNow() {
  if (!GS.url || !GS.key || Sync.suspend) return Promise.resolve(false);
  return gsPost({ action: 'saveSettings', name: 'all', value: collectSettings(), device: deviceLabel() })
    .then(res => {
      Sync.on = true; Sync.lastPush = Date.now();
      setSyncNote('Settings saved to the sheet \u00b7 ' + fmtWhen(res.updated || new Date().toISOString()));
      return true;
    })
    .catch(err => {
      setSyncNote('Could not save settings to the sheet: ' + err.message);
      return false;
    });
}

const pushSettings = debounce(pushSettingsNow, 1500);

function fmtWhen(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
  } catch (e) { return String(iso); }
}

function setSyncNote(msg) {
  const el = document.getElementById('gs-sync-note');
  if (el) el.textContent = msg;
  if (typeof refreshSaveNote === 'function') refreshSaveNote();
}

/** Called once the sheet connection is known to work. */
function startSettingsSync(announce) {
  if (!GS.url || !GS.key) return;
  pullSettings(announce);
}

/* ---- the Save button in the sidebar --------------------------------------
   Settings already save themselves a moment after you change them, but that
   is invisible, and after losing a set of typed-in values once you want to
   see it happen. This is that button: it writes everything to the sheet now
   and says so. ------------------------------------------------------------ */
function initSaveButton() {
  const btn = document.getElementById('save-all');
  if (!btn) return;
  btn.addEventListener('click', () => {
    if (!GS.url || !GS.key) {
      setSaveNote('Connect a Google Sheet first \u2014 see the Google Sheet tab.', 'warn');
      toast('Connect your Google Sheet first, on the Google Sheet tab.');
      return;
    }
    setSaveNote('Saving\u2026', 'busy');
    btn.disabled = true;
    pushSettingsNow().then(okDone => {
      btn.disabled = false;
      if (okDone) {
        setSaveNote('Saved \u00b7 ' + fmtWhen(new Date().toISOString()), 'ok');
        toast('Your setup is saved. Any browser that opens the site will now start from here.');
      } else {
        setSaveNote('Could not save \u2014 see the Google Sheet tab.', 'warn');
      }
    });
  });
  refreshSaveNote();
}

function setSaveNote(msg, cls) {
  const el = document.getElementById('save-all-note');
  if (!el) return;
  el.textContent = msg;
  el.className = 'save-all-note' + (cls ? ' ' + cls : '');
}

/** Keeps the little line under the button honest about where things stand. */
function refreshSaveNote() {
  if (!GS.url || !GS.key) {
    setSaveNote('Saved in this browser only \u2014 connect a sheet to share it.', 'warn');
  } else if (Sync.lastPush) {
    setSaveNote('Saved \u00b7 ' + fmtWhen(new Date(Sync.lastPush).toISOString()), 'ok');
  } else if (Sync.lastPulled) {
    setSaveNote('Loaded from your sheet \u00b7 ' + fmtWhen(Sync.lastPulled), 'ok');
  } else {
    setSaveNote('Connected \u2014 changes save on their own.', 'ok');
  }
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

  // Reconnect in the background so a browser that has never seen this
  // dashboard still opens on the settings you last saved anywhere.
  if (savedUrl && savedKey) {
    GS.url = savedUrl; GS.key = savedKey;
    setGsStatus('Reconnecting\u2026', 'busy');
    gsGet({ action: 'meta' }).then(meta => {
      GS.meta = meta;
      setGsStatus('Connected: ' + meta.spreadsheetName + ' (' + meta.sheets.length + ' sheets)', 'ok');
      renderSheetList(meta);
      updateGsOnlyButtons();
      startSettingsSync(false);
    }).catch(err => {
      GS.meta = null;
      setGsStatus('Saved connection could not be reached: ' + err.message, 'err');
      updateGsOnlyButtons();
    });
  }

  document.getElementById('gs-connect').addEventListener('click', connectSheet);
  document.getElementById('gs-forget').addEventListener('click', () => {
    Store.remove('sl_gs_url'); Store.remove('sl_gs_key');
    urlEl.value = ''; keyEl.value = '';
    GS.url = ''; GS.key = ''; GS.meta = null;
    Sync.on = false;
    setSyncNote('Not connected \u2014 settings are kept in this browser only.');
    document.getElementById('gs-sheet-list').style.display = 'none';
    setGsStatus('Saved details cleared.', '');
    updateGsOnlyButtons();
  });

  const pull = document.getElementById('gs-sync-pull');
  if (pull) pull.addEventListener('click', () => {
    if (!GS.url || !GS.key) { setSyncNote('Connect to the sheet first.'); return; }
    setSyncNote('Reading\u2026');
    pullSettings(true);
  });
  const push = document.getElementById('gs-sync-push');
  if (push) push.addEventListener('click', () => {
    if (!GS.url || !GS.key) { setSyncNote('Connect to the sheet first.'); return; }
    setSyncNote('Saving\u2026');
    pushSettingsNow().then(ok => { if (ok) toast('Settings saved to your Google Sheet.'); });
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
    startSettingsSync(true);         // bring this browser in line with the sheet
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
  openBorderWidth: 2,        // that border's thickness (px)
  // Drill-list level colours. Used by Product Performance AND the Catalog
  // table, so both drill lists always look the same.
  lvlFill: ['#FFFFFF', '#F4F7F5', '#F7F4EC', '#F3F1F7', '#FAF6F2'],
  lvlRail: ['#1F6F5C', '#B9862F', '#4A6FA5', '#8C2E1B']
};

const LVL_FILL_DEFAULT = THEME_DEFAULT.lvlFill.slice();
const LVL_RAIL_DEFAULT = THEME_DEFAULT.lvlRail.slice();

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

const Theme = Object.assign({}, THEME_DEFAULT,
  { lvlFill: LVL_FILL_DEFAULT.slice(), lvlRail: LVL_RAIL_DEFAULT.slice() });

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
  // --lvl-open is what the drill lists actually read, so the "Opened drill row
  // fill" picker now reaches Product Performance and the Catalog alike.
  root.style.setProperty('--lvl-open', Theme.openFill || '#FDF3E7');
  root.style.setProperty('--open-bc', Theme.openBorderColor || '#A6402C');
  root.style.setProperty('--open-bw', (Theme.openBorderWidth === undefined ? 2 : Theme.openBorderWidth) + 'px');

  // Drill-list level colours. Levels past the last one you set simply keep
  // repeating the deepest colour, so nine-level catalogs still look right.
  const fills = (Theme.lvlFill && Theme.lvlFill.length) ? Theme.lvlFill : LVL_FILL_DEFAULT;
  for (let i = 0; i <= 8; i++) {
    root.style.setProperty('--lvl-' + i, fills[Math.min(i, fills.length - 1)]);
  }
  const rails = (Theme.lvlRail && Theme.lvlRail.length) ? Theme.lvlRail : LVL_RAIL_DEFAULT;
  for (let i = 1; i <= 8; i++) {
    root.style.setProperty('--lvl-rail-' + i, rails[Math.min(i - 1, rails.length - 1)]);
  }

  const b = document.body;
  ['density-compact', 'density-normal', 'density-comfortable'].forEach(c => b.classList.remove(c));
  b.classList.add('density-' + Theme.density);
  b.classList.toggle('no-zebra', !Theme.zebra);
  b.classList.toggle('no-gridlines', !Theme.gridLines);
}

/** Current drill-level fill / rail colours, always as a fresh array so the
 *  defaults can never be edited by accident. */
function lvlFills() {
  const a = (Theme.lvlFill && Theme.lvlFill.length) ? Theme.lvlFill : LVL_FILL_DEFAULT;
  return LVL_FILL_DEFAULT.map((d, i) => a[i] || d);
}
function lvlRails() {
  const a = (Theme.lvlRail && Theme.lvlRail.length) ? Theme.lvlRail : LVL_RAIL_DEFAULT;
  return LVL_RAIL_DEFAULT.map((d, i) => a[i] || d);
}

/** The colour-picker rows under Settings -> Look & Feel -> Drill list colours. */
function lvlColorRows() {
  const fills = lvlFills(), rails = lvlRails();
  let out = '';
  for (let i = 0; i < fills.length; i++) {
    out += '<div class="color-row">' +
      '<label class="toolbar-label">' + (i === 0 ? 'Top level (row fill)' : 'Level ' + i + ' fill') + '</label>' +
      '<input type="color" class="lvl-fill" data-i="' + i + '" value="' + fills[i] + '">' +
      '<span class="hexcode" id="lvl-fill-' + i + '-val">' + fills[i] + '</span>' +
      (i > 0 ? '<label class="toolbar-label" style="min-width:96px;">left bar</label>' +
        '<input type="color" class="lvl-rail" data-i="' + (i - 1) + '" value="' + rails[i - 1] + '">' +
        '<span class="hexcode" id="lvl-rail-' + (i - 1) + '-val">' + rails[i - 1] + '</span>' : '') +
    '</div>';
  }
  return out;
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
          '<button class="seg-btn" data-stab="performance">01 Performance</button>' +
          '<button class="seg-btn" data-stab="catalog">02 Catalog</button>' +
          '<button class="seg-btn" data-stab="dashboard">03 Dashboard</button>' +
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
      '<p class="drill-subtitle" style="margin-top:6px;">Applies to both the Dashboard and the Product Performance charts. ' +
        'While it is on, nothing can be moved, resized, added, removed, recoloured or re-pointed at different data \u2014 ' +
        'and the per-card data dropdowns are hidden. Clicking a bar to highlight still works, and so does the enlarge button. ' +
        'Unlock to make changes again.</p>' +
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
      Object.assign(Theme, THEME_DEFAULT,
        { lvlFill: LVL_FILL_DEFAULT.slice(), lvlRail: LVL_RAIL_DEFAULT.slice() });
      saveTheme(); renderSettingsBody(); toast('Look & Feel has been reset.');
    });
    wrap.querySelector('#set-reset-all').addEventListener('click', function () {
      Store.remove('sl_theme'); Store.remove('sl_snapshot_config'); Store.remove('sl_behaviour');
      Store.remove('sl_prefs'); Store.remove('sl_colwidths');
      Object.assign(Theme, THEME_DEFAULT,
        { lvlFill: LVL_FILL_DEFAULT.slice(), lvlRail: LVL_RAIL_DEFAULT.slice() });
      saveTheme();
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

    '<h3 class="snap-set-title">Drill list colours</h3>' +
    '<p class="drill-subtitle">One colour per level, so you can see how deep a row sits. ' +
      'These apply to the drill list on Product Performance and on Catalog together \u2014 ' +
      'both tables always match. Levels below 4 keep the deepest colour.</p>' +
    lvlColorRows() +
    '<div class="settings-row">' +
      '<button class="ghost-btn small" id="set-lvl-reset">Reset drill colours</button>' +
      '<span class="drill-count">back to the built-in shades</span>' +
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
  wrap.querySelectorAll('.lvl-fill').forEach(function (el) {
    el.addEventListener('input', function (e) {
      const i = parseInt(e.target.dataset.i, 10);
      const next = lvlFills(); next[i] = e.target.value;
      Theme.lvlFill = next;
      const lab = wrap.querySelector('#lvl-fill-' + i + '-val');
      if (lab) lab.textContent = e.target.value;
      saveTheme();
    });
  });
  wrap.querySelectorAll('.lvl-rail').forEach(function (el) {
    el.addEventListener('input', function (e) {
      const i = parseInt(e.target.dataset.i, 10);
      const next = lvlRails(); next[i] = e.target.value;
      Theme.lvlRail = next;
      const lab = wrap.querySelector('#lvl-rail-' + i + '-val');
      if (lab) lab.textContent = e.target.value;
      saveTheme();
    });
  });
  const lvlReset = wrap.querySelector('#set-lvl-reset');
  if (lvlReset) lvlReset.addEventListener('click', function () {
    Theme.lvlFill = LVL_FILL_DEFAULT.slice();
    Theme.lvlRail = LVL_RAIL_DEFAULT.slice();
    saveTheme(); renderSettingsBody(); toast('Drill list colours reset.');
  });

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
  flags: { stockout: false, nosale: false, overstock: false, lowstock: false, medium: false, healthy: false },
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

  // Up to nine levels, in the order set in Settings > 02 Catalog.
  const levels = (CatPrefs.levels || []).filter(Boolean);
  if (!levels.length) levels.push('Article No');

  const root = { key: '__root__', path: '', depth: -1, children: new Map(),
                 sold: 0, purchased: 0, obs: 0, cbs: 0, hasOBS: false, lastSale: null,
                 meta: {}, skuMap: null };

  function nodeFor(rec) {
    // walks down one branch, creating the nodes on the way
    const chain = [];
    let node = root;
    for (let i = 0; i < levels.length; i++) {
      const key = dimKey(rec, levels[i]);
      let child = node.children.get(key);
      if (!child) {
        child = { key, dim: levels[i], depth: i,
                  path: (node.path ? node.path + '|' : '') + levels[i] + '=' + key,
                  children: new Map(), sold: 0, purchased: 0, obs: 0, cbs: 0,
                  hasOBS: false, lastSale: null, meta: {},
                  // only the deepest level tracks each unit; parents merge
                  skuMap: (i === levels.length - 1) ? new Map() : null };
        node.children.set(key, child);
      }
      chain.push(child);
      node = child;
    }
    return chain;
  }

  /** The real stock-keeping unit behind a row: one article, in one colour, in
   *  one size. That is the thing you actually reorder, and it is what the Max
   *  Level has to be counted in - see replenFor.
   *
   *  Not Item Code. In this ERP export Item Code is a per-piece lot number:
   *  71,966 different codes across 71,974 stock rows, averaging under one
   *  garment each, with the same article/colour/size carrying dozens of them.
   *  Counting those would inflate the unit count roughly tenfold. Item Code is
   *  only used when the article/colour/size fields are all missing. */
  function skuOf(r) {
    const art = dimKey(r, 'Article No'), col = dimKey(r, 'Colour'), sz = dimKey(r, 'Size');
    if (art !== '(blank)' || col !== '(blank)' || sz !== '(blank)') return art + '|' + col + '|' + sz;
    const code = r['Item Code'];
    return (code === null || code === undefined || code === '') ? '(blank)' : String(code);
  }

  const META_KEEP = ['Section', 'Sub Section', 'Brand', 'Supplier', 'Style', 'Colour', 'Size'];
  function addMeta(n, r) {
    META_KEEP.forEach(f => { if (!n.meta[f] && r[f]) n.meta[f] = String(r[f]); });
  }

  sales.forEach(r => {
    const q = recQty(r), sku = skuOf(r);
    nodeFor(r).forEach(n => {
      n.sold += q;
      if (n.skuMap) n.skuMap.set(sku, (n.skuMap.get(sku) || 0) + q);
      if (r.Date && (!n.lastSale || r.Date > n.lastSale)) n.lastSale = r.Date;
      addMeta(n, r);
    });
  });
  purch.forEach(r => {
    const q = recQty(r), sku = skuOf(r);
    nodeFor(r).forEach(n => {
      n.purchased += q;
      if (n.skuMap && !n.skuMap.has(sku)) n.skuMap.set(sku, 0);   // stocked, never sold
      addMeta(n, r);
    });
  });
  stock.forEach(r => {
    const cb = recQty(r), ob = recOpeningQty(r), sku = skuOf(r);
    nodeFor(r).forEach(n => {
      n.cbs += cb;
      if (n.skuMap && !n.skuMap.has(sku)) n.skuMap.set(sku, 0);
      if (ob !== null) { n.obs += ob; n.hasOBS = true; }
      addMeta(n, r);
    });
  });

  // metrics bottom-up, and a sorted child list on every node
  (function finish(node) {
    node.childList = [...node.children.values()];
    node.childList.forEach(finish);
    node.childList.sort((a, b) => b.sold - a.sold || b.cbs - a.cbs);
    // Every row keeps the sales of each stock-keeping unit beneath it, sorted,
    // with a running total. That is what lets the max level answer "how much of
    // this is fast enough to justify more than the minimum order" for any
    // lead time you type in - see maxLevelFor.
    let arr;
    if (node.skuMap) {
      arr = Float64Array.from(node.skuMap.values());
      node.skuMap = null;
    } else {
      let total = 0;
      node.childList.forEach(c => { total += c.skuSold.length; });
      arr = new Float64Array(total);
      let at = 0;
      node.childList.forEach(c => { arr.set(c.skuSold, at); at += c.skuSold.length; });
    }
    arr.sort();
    const cum = new Float64Array(arr.length + 1);
    for (let i = 0; i < arr.length; i++) cum[i + 1] = cum[i] + arr[i];
    node.skuSold = arr;
    node.skuCum = cum;
    node.skuCount = Math.max(1, arr.length);
    if (node !== root) Object.assign(node, catalogMetrics(node, days, anchor));
  })(root);

  return { rows: root.childList, days, range, anchor, levels };
}

function catalogMetrics(x, days, anchor) {
  const avgDaily = x.sold / days;
  const cover = avgDaily > 0 ? x.cbs / avgDaily : (x.cbs > 0 ? Infinity : 0);
  // Sell-through = how much of what you had actually sold.
  //
  // "What you had" is the opening stock plus everything bought in the window.
  // The old version divided by sold + closing, which quietly reads 100% for
  // ANY row whose closing is zero - including rows with no stock record at
  // all. GARMENTS showed 307 sold, no opening, no purchases and no stock
  // rows whatsoever, yet sat at the top of the column on a confident 100%.
  // With nothing to divide by, the honest answer is that we do not know, so
  // the cell shows a dash.
  const available = (x.hasOBS ? x.obs : 0) + x.purchased;
  let sellThrough;
  if (available > 0) sellThrough = (x.sold / available) * 100;
  else if (x.cbs > 0 || x.hasOBS) sellThrough = (x.sold / (x.sold + x.cbs)) * 100;
  else sellThrough = null;                       // no stock and no purchases
  const daysSince = x.lastSale ? Math.round((anchor - x.lastSale) / 86400000) : null;

  // Status is read off the SAME Stock % that colours the row, so the word and
  // the colour can never disagree.
  //
  // It used to be judged on days of cover instead, which measures something
  // different and punished breadth. "Gharara sait" held 128 pieces spread over
  // 116 article/colour/size combinations - about one piece each, and only 55%
  // of its max level - yet 202 days of cover labelled it "Overstock". Roughly
  // one row in four carried a word that contradicted its own colour.
  const pct = replenFor(x.path, x.sold, days, x.cbs, x).pct;

  let status;
  if (x.cbs === 0 && x.sold > 0) status = 'Stockout';        // sold out entirely
  else if (x.cbs === 0 && x.sold === 0) status = 'Idle';     // nothing either way
  else if (x.sold === 0) status = 'No sale';                 // stock sitting, nothing moved
  else status = stockBandStatus(pct);                        // the band, word for word
  return { avgDaily, cover, sellThrough, daysSince, status, pct };
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
  // one chip per status word, so what you filter is what the column says
  if (f.stockout) rows = rows.filter(r => r.status === 'Stockout');
  if (f.nosale) rows = rows.filter(r => r.status === 'No sale');
  if (f.lowstock) rows = rows.filter(r => r.status === 'Low stock');
  if (f.medium) rows = rows.filter(r => r.status === 'Medium stock');
  if (f.healthy) rows = rows.filter(r => r.status === 'Healthy');
  if (f.overstock) rows = rows.filter(r => r.status === 'Overstock');

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

  // The stylesheet caps this box at 64vh. An inline height alone loses to
  // that cap, so dragging could never make the table taller than about
  // two-thirds of the window. Setting max-height alongside it lifts the cap.
  let currentH = 0;
  function applyHeight(h) {
    currentH = Math.round(h);
    box.style.flex = '0 0 auto';
    box.style.height = currentH + 'px';
    box.style.maxHeight = currentH + 'px';
  }

  const saved = parseInt(Store.get('sl_cat_height') || '', 10);
  if (saved > 120) applyHeight(saved);

  grip.addEventListener('mousedown', e => {
    e.preventDefault();
    const startY = e.clientY;
    const startH = box.getBoundingClientRect().height;
    box.style.flex = '0 0 auto';
    document.body.classList.add('row-resizing');

    const move = ev => {
      // no ceiling: drag it as tall as you like and the page scrolls
      applyHeight(Math.max(140, startH + (ev.clientY - startY)));
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      document.body.classList.remove('row-resizing');
      // Save the height we actually set, not a measured rectangle. Measuring
      // gives 0 whenever the tab is not laid out at that instant, which then
      // saved a zero and lost the setting.
      if (currentH > 120) Store.set('sl_cat_height', String(currentH));
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  });

  // double-click puts it back to filling the tab
  grip.addEventListener('dblclick', () => {
    box.style.flex = ''; box.style.height = ''; box.style.maxHeight = '';
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
  const sortSel = document.getElementById('cat-sort');
  // Point the box at the sort that is actually in use. Without this the
  // browser just selects whichever option happens to be listed first, so
  // reordering the list would have made the box disagree with the table.
  if (sortSel) {
    sortSel.value = Catalog.sort;
    if (!sortSel.value) { sortSel.value = 'worstcover'; Catalog.sort = 'worstcover'; }
    sortSel.addEventListener('change', e => {
      Catalog.sort = e.target.value; Catalog.sortCol = null; renderCatalog();
    });
  }
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
  const b = stockBands();
  const bands = [
    ['0\u2013' + b.low + '%', '#ea9999', 'Low stock'],
    ['>' + b.low + '\u2013' + b.mid + '%', '#ffd966', 'Medium stock'],
    ['>' + b.mid + '\u2013' + b.good + '%', '#b6d7a8', 'Healthy'],
    ['>' + b.good + '%', '#b4a7d6', 'Overstock']
  ];
  return '<div class="stock-legend"><span class="sl-title">Stock %</span>' +
    bands.map(b => '<span class="sl-item"><span class="sl-sw" style="background:' + b[1] + '"></span>' +
      b[0] + ' \u00b7 ' + b[2] + '</span>').join('') +
    '<span class="sl-note">ML = ADC \u00d7 LT \u00d7 SF, at least MOQ per item in the row \u00b7 ' +
    'Stock % = (Closing + MIT) \u00f7 ML \u00b7 the Status column uses these exact bands</span></div>';
}

/** Sets one field on every design currently listed. */
function replenBulkApply(field, value) {
  const rows = Catalog.lastRows || [];
  if (!rows.length) { toast('Nothing on screen to apply to.'); return; }
  const v = parseFloat(value);
  if (isNaN(v)) { toast('Enter a number first.'); return; }
  // Keyed by path, not key. Overrides are stored and read by path everywhere
  // else, so writing them under the plain name meant nothing ever picked them
  // up and the button looked dead.
  rows.forEach(d => {
    const o = Replen.overrides[d.path] || (Replen.overrides[d.path] = {});
    o[field] = v;
  });
  saveReplen();
  renderCatalog();
  toast(field.toUpperCase() + ' set to ' + v + ' on ' + rows.length + ' designs.');
}

/** The row under the catalogue toolbar.
 *
 *  The "Set for all listed designs" boxes were removed in v46 at the user's
 *  request - the same four values already live in Settings > 02 Catalog as
 *  the defaults for every row, and each row can still be typed over
 *  individually. Only the escape hatch stays: one button to drop every
 *  manual value and go back to the calculated numbers. */
function replenBulkBarHtml() {
  return '<div class="cat-bulk">' +
    '<button class="ghost-btn small" id="bulk-clear">Clear all manual values</button>' +
    '<span class="drill-count">puts every ADC, LT, SF and MOQ back to the calculated value</span>' +
    '</div>';
}

function wireReplenBulk() {
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
    // The colour key sits above the table. It is off by default now - the
    // colours are self-explanatory once you know them, and the strip was
    // taking a line of screen for nothing. Turn it back on in
    // Settings > 02 Catalog if you want the reminder.
    extra.innerHTML = planOn
      ? ((CatPrefs.showLegend ? replenLegendHtml() : '') + replenBulkBarHtml())
      : '';
    if (planOn) wireReplenBulk();
  }
  const rows = catalogFiltered(built.rows);
  Catalog.lastRows = rows;
  Catalog.lastBuilt = built;      // kept so a live edit can find its row again

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
    (catColOn('adc') ? '<th class="num" title="Average Daily Consumption" data-sc="adc">ADC' + catSortArrow('adc') + '</th>' : '') +
    (catColOn('lt') ? '<th class="num" title="Lead Time in days" data-sc="lt">LT' + catSortArrow('lt') + '</th>' : '') +
    (catColOn('sf') ? '<th class="num" title="Safety Factor" data-sc="sf">SF' + catSortArrow('sf') + '</th>' : '') +
    (catColOn('moq') ? '<th class="num" title="Minimum Order Quantity" data-sc="moq">MOQ' + catSortArrow('moq') + '</th>' : '') +
    (catColOn('ml') ? '<th class="num" title="Max Level = ADC x LT x SF, and never less than MOQ for each item under the row" data-sc="ml">ML' + catSortArrow('ml') + '</th>' : '') +
    (catColOn('mit') ? '<th class="num" title="Material In Transit" data-sc="mit">MIT' + catSortArrow('mit') + '</th>' : '') +
    (catColOn('stockpct') ? '<th class="num" title="(Closing + MIT) as a share of Max Level" data-sc="pct">Stock %' + catSortArrow('pct') + '</th>' : '') +
    (catColOn('reorder') ? '<th class="num" title="ML minus what you have, rounded up to the MOQ" data-sc="reorder">Reorder' + catSortArrow('reorder') + '</th>' : '') +
    (catColOn('cover') ? '<th class="num" data-sc="cover">Cover' + catSortArrow('cover') + '</th>' : '') +
    (catColOn('sellthru') ? '<th class="num" data-sc="sellThrough">Sell-thru' + catSortArrow('sellThrough') + '</th>' : '') +
    (catColOn('lastsold') ? '<th data-sc="lastSale">Last sold' + catSortArrow('lastSale') + '</th>' : '') +
    (catColOn('status') ? '<th data-sc="status">Status' + catSortArrow('status') + '</th>' : '') +
    '</tr></thead>';

  const body = rows.slice(0, CatPrefs.maxRows || 300).map(d => catalogNodeRow(d, built.days)).join('');

  const tSold = rows.reduce((a, r) => a + r.sold, 0);
  const tPurch = rows.reduce((a, r) => a + r.purchased, 0);
  const tObs = rows.reduce((a, r) => a + r.obs, 0);
  const tCbs = rows.reduce((a, r) => a + r.cbs, 0);
  // same basis as the rows: sold against what was available
  const tAvail = rows.reduce((a, r) => a + (r.hasOBS ? r.obs : 0) + r.purchased, 0);
  const tST = tAvail > 0 ? (tSold / tAvail) * 100
            : (tSold + tCbs) > 0 ? (tSold / (tSold + tCbs)) * 100 : 0;
  // totals now use each row's own path, and add up the max level as well
  const reps = rows.map(r => replenFor(r.path, r.sold, built.days, r.cbs, r));
  const tReorder = reps.reduce((a, x) => a + x.reorder, 0);
  const totMl = reps.reduce((a, x) => a + x.ml, 0);
  const tOnHand = rows.reduce((a, r) => a + (r.cbs || 0), 0);
  const totPct = totMl > 0 ? (tOnHand / totMl) * 100 : 0;
  const fillerCols = (catColOn('category') ? 1 : 0) + (catColOn('colours') ? 1 : 0);
  const foot = '<tfoot><tr>' +
    '<td>Total \u00b7 ' + rows.length.toLocaleString('en-IN') + ' designs</td>' +
    (fillerCols ? '<td colspan="' + fillerCols + '"></td>' : '') +
    (catColOn('sold') ? '<td class="num">' + fmtNum(tSold) + '</td>' : '') +
    (catColOn('purchased') ? '<td class="num">' + fmtNum(tPurch) + '</td>' : '') +
    (catColOn('opening') ? '<td class="num">' + fmtNum(tObs) + '</td>' : '') +
    (catColOn('closing') ? '<td class="num">' + fmtNum(tCbs) + '</td>' : '') +
    ['adc','lt','sf','moq'].filter(catColOn).map(function () { return '<td></td>'; }).join('') +
    (catColOn('ml') ? '<td class="num">' + fmtNum(totMl, 0) + '</td>' : '') +
    (catColOn('mit') ? '<td></td>' : '') +
    (catColOn('stockpct') ? '<td class="num stock-pct ' + stockPctClass(totPct) + '">' +
      (totPct > 999 ? '999%+' : fmtNum(totPct, 0) + '%') + '</td>' : '') +
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

    // As you type: recalculate this row on the spot. Redrawing the whole table
    // on every keystroke would throw away the box you are typing in, so the
    // row's own cells are rewritten in place instead.
    inp.addEventListener('input', e => {
      e.stopPropagation();
      applyReplenEdit(e.target);
      refreshReplenRow(e.target.closest('tr'));
    });

    // On leaving the box: save and redraw properly, so sorting, the totals
    // line and the colour bars all catch up.
    inp.addEventListener('change', e => {
      e.stopPropagation();
      applyReplenEdit(e.target);
      saveReplen();
      renderCatalog();
    });
  });
}

/** Writes one edited box into the overrides. */
function applyReplenEdit(el) {
  const key = el.dataset.rkey, field = el.dataset.rfield, raw = el.value;
  const o = Replen.overrides[key] || (Replen.overrides[key] = {});
  if (raw === '' || isNaN(parseFloat(raw))) delete o[field];
  else o[field] = parseFloat(raw);
  if (!Object.keys(o).length) delete Replen.overrides[key];
}

/** Rewrites the calculated cells of one row from the numbers now in the boxes:
 *  Max Level, Stock % and its colour, Reorder, Cover and Status. Everything
 *  that hangs off ADC / LT / SF / MOQ / MIT moves together, without the table
 *  being rebuilt under the cursor. */
function refreshReplenRow(tr) {
  if (!tr) return;
  const node = catalogNodeByPath(tr.dataset.key);
  if (!node) return;
  const days = catalogDays();
  const r = replenFor(node.path, node.sold, days, node.cbs, node);

  const set = (sel, text, title) => {
    const td = tr.querySelector(sel);
    if (!td) return null;
    td.textContent = text;
    if (title !== undefined) td.title = title;
    return td;
  };

  set('.cat-ml', fmtNum(r.ml, 0),
      r.units > 1
        ? fmtNum(r.units) + ' items under this row \u00d7 MOQ ' + r.moq + ' = ' + fmtNum(r.units * r.moq, 0) +
          '; demand (ADC \u00d7 LT \u00d7 SF) would give ' + fmtNum(r.rawMl, 0) + ' \u2014 the larger wins'
        : 'ADC ' + fmtNum(r.adc, 2) + ' \u00d7 LT ' + r.lt + ' \u00d7 SF ' + r.sf + ', at least MOQ ' + r.moq);

  const pctTd = set('.stock-pct', r.pct > 999 ? '999%+' : fmtNum(r.pct, 0) + '%',
      'Closing ' + fmtNum(r.onHand - r.mit) + (r.mit ? ' + in transit ' + fmtNum(r.mit) : '') +
      ' vs max level ' + fmtNum(r.ml, 0) +
      (r.units > 1 ? ' (' + fmtNum(r.units) + ' items under this row)' : ''));
  if (pctTd) pctTd.className = 'num stock-pct ' + stockPctClass(r.pct);

  const reTd = set('.cat-reorder', r.reorder > 0 ? fmtNum(r.reorder) : '\u2014');
  if (reTd) reTd.className = 'num cat-reorder' + (r.reorder > 0 ? ' has' : '');

  // Status follows the same Stock %, so it has to move with it.
  const stTd = tr.querySelector('.status-tag');
  if (stTd) {
    const m = catalogMetrics(node, days, dataAnchorDate());
    stTd.textContent = m.status;
    stTd.className = 'status-tag ' + catalogStatusClass(m.status);
  }
}

/** Finds a built node again from the path stored on its row. */
function catalogNodeByPath(path) {
  if (!path || !Catalog.lastBuilt) return null;
  let found = null;
  (function walk(list) {
    for (const n of list) {
      if (found) return;
      if (n.path === path) { found = n; return; }
      if (n.childList && n.childList.length) walk(n.childList);
    }
  })(Catalog.lastBuilt.rows || []);
  return found;
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
  if (col === 'colours') return (d.childList || []).length;
  if (col === 'obs') return d.obs;
  if (col === 'cbs') return d.cbs;
  if (col === 'cover') return d.cover === Infinity ? 1e12 : d.cover;
  if (col === 'sellThrough') return d.sellThrough === null ? -Infinity : d.sellThrough;
  if (col === 'lastSale') return d.lastSale ? d.lastSale.getTime() : -Infinity;
  if (col === 'status') return String(d.status).toLowerCase();
  if (['adc', 'lt', 'sf', 'moq', 'ml', 'mit', 'pct', 'reorder'].indexOf(col) !== -1) {
    // path, not key: overrides are stored per path, and the max level needs the
    // row's own unit count. Sorting used to read a different number from the
    // one on screen.
    const r = replenFor(d.path, d.sold, catalogDays(), d.cbs, d);
    return col === 'pct' ? (isFinite(r.pct) ? r.pct : 1e12) : r[col];
  }
  return d[col];
}

/** The columns in the order they are drawn, Design first. */
function catalogVisibleCols() {
  return ['design'].concat(
    ['category','colours','sold','purchased','opening','closing',
     'adc','lt','sf','moq','ml','mit','stockpct','reorder',
     'cover','sellthru','lastsold','status'].filter(catColOn));
}

/** How many columns the colour bar under a row may cover. The setting names
 *  the column to stop after; if that column is switched off, the bar falls
 *  back to the whole width. */
function stripSpanCols() {
  const stop = CatPrefs.stripUntil || 'opening';
  if (stop === 'all') return catalogColCount();
  const cols = catalogVisibleCols();
  const i = cols.indexOf(stop);
  return i === -1 ? catalogColCount() : i + 1;
}

/** How many columns the table has right now, for full-width rows. */
function catalogColCount() {
  return 1 + ['category','colours','sold','purchased','opening','closing',
              'adc','lt','sf','moq','ml','mit','stockpct','reorder',
              'cover','sellthru','lastsold','status'].filter(catColOn).length;
}

/** One row per node, at any depth, plus its children when open. */
function catalogNodeRow(n, days) {
  const open = !!Catalog.expanded[n.path];
  const img = Catalog.images[n.path];
  const kids = n.childList || [];
  const r = replenFor(n.path, n.sold, days, n.cbs, n);
  const band = stockPctClass(r.pct);

  // The colour column shows one block per child, tinted by that child's own
  // stock band - no colour swatches, the colour means stock health.
  const maxDots = CatPrefs.maxDots || 10;
  const dots = kids.slice(0, maxDots).map(k => {
    const kr = replenFor(k.path, k.sold, days, k.cbs, k);
    return '<span class="cat-band ' + stockPctClass(kr.pct) + '" title="' +
      escapeHtml(k.key + ' \u2014 ' + fmtNum(k.cbs) + ' in stock, ' + fmtNum(k.sold) + ' sold, ' +
        (kr.pct > 999 ? '999%+' : Math.round(kr.pct) + '%') + ' of max level') + '"></span>';
  }).join('') + (kids.length > maxDots ? '<span class="cat-more">+' + (kids.length - maxDots) + '</span>' : '');

  let html = '<tr class="cat-row cat-lvl-' + Math.min(n.depth, 8) + (open ? ' is-open' : '') +
      (n.depth === 0 ? ' cat-design' : ' cat-child') + '" data-key="' + escapeHtml(n.path) + '">' +
    '<td class="cat-c-design" style="--indent:' + (10 + n.depth * 18) + 'px">' +
      (kids.length ? '<button class="cat-caret' + (open ? ' open' : '') + '">\u25B8</button>'
                   : '<span class="cat-caret-gap"></span>') +
      (CatPrefs.showThumbs
        ? '<button class="cat-thumb" data-img="' + escapeHtml(n.path) + '" title="' +
          (img ? 'View, replace or remove the photo' : 'Add a photo') + '">' +
          (img ? '<img src="' + img + '" alt="">' : '<span class="cat-thumb-empty">\uFF0B</span>') + '</button>'
        : '') +
      '<span class="cat-name">' + escapeHtml(n.key) + '</span>' +
      (CatPrefs.compactMeta && kids.length
        ? '<span class="cat-sub">' + kids.length + ' ' + escapeHtml(String(kids[0].dim || '')).toLowerCase() + '</span>'
        : '') +
      '<button class="cat-popout" title="Open full details in a separate window">\u29C9</button>' +
    '</td>' +
    (catColOn('category') ? '<td class="cat-cat">' + escapeHtml(n.meta['Sub Section'] || n.meta.Section || '\u2014') + '</td>' : '') +
    (catColOn('colours') ? '<td class="cat-c-colours"><span class="cat-bands">' + dots + '</span></td>' : '') +
    (catColOn('sold') ? '<td class="num">' + fmtNum(n.sold) + '</td>' : '') +
    (catColOn('purchased') ? '<td class="num cat-purch">' + fmtNum(n.purchased) + '</td>' : '') +
    (catColOn('opening') ? '<td class="num obs-col">' + (n.hasOBS ? fmtNum(n.obs) : '\u2014') + '</td>' : '') +
    (catColOn('closing') ? '<td class="num cbs-col">' + fmtNum(n.cbs) + '</td>' : '') +
    replenCells(n.path, r) +
    (catColOn('cover') ? '<td class="num">' + (n.cover === Infinity ? '\u221E' : fmtNum(n.cover, 0) + 'd') + '</td>' : '') +
    (catColOn('sellthru') ? '<td class="num"' +
        (n.sellThrough === null ? ' title="No opening stock, no purchases and no stock rows \u2014 there is nothing to measure the sales against."' : '') +
        '>' + (n.sellThrough === null ? '\u2014' : fmtNum(n.sellThrough, 1) + '%') + '</td>' : '') +
    (catColOn('lastsold') ? '<td>' + (n.lastSale ? fmtDate(n.lastSale) : '\u2014') + '</td>' : '') +
    (catColOn('status') ? '<td><span class="status-tag ' + catalogStatusClass(n.status) + '">' + n.status + '</span></td>' : '') +
  '</tr>';

  // colour-coded bar under the row, one block per child sized by stock
  if (CatPrefs.showStrip && kids.length) html += catalogStripRow(n, days);

  if (open) kids.forEach(k => { html += catalogNodeRow(k, days); });
  return html;
}

/** The bar under a row: one block per child, width by stock, colour by band. */
function catalogStripRow(n, days) {
  const kids = n.childList || [];
  const blocks = kids.map(k => {
    const kr = replenFor(k.path, k.sold, days, k.cbs, k);
    return '<span class="cs-block ' + stockPctClass(kr.pct) + '" style="flex:' + Math.max(1, k.cbs) + '" ' +
      'title="' + escapeHtml(k.key + ': ' + fmtNum(k.cbs) + ' in stock, ' +
        (kr.pct > 999 ? '999%+' : Math.round(kr.pct) + '%') + ' of max level') + '"></span>';
  }).join('');
  if (!blocks) return '';
  // How far across the table the bar is allowed to run. It used to span every
  // column, first to last; now it stops after the column chosen in
  // Settings > 02 Catalog, and the rest of the row is left clear.
  const total = catalogColCount();
  const span = Math.max(1, Math.min(total, stripSpanCols()));
  // The bar carries its parent's level class, so it is tinted with the same
  // drill-list colour as the row it belongs to.
  return '<tr class="cat-strip-row cat-lvl-' + Math.min(n.depth, 8) + '">' +
    '<td colspan="' + span + '" style="--indent:' + (10 + (n.depth + 1) * 18) + 'px">' +
      '<span class="cs-bar">' + blocks + '</span></td>' +
    (span < total ? '<td colspan="' + (total - span) + '" class="cs-pad"></td>' : '') +
    '</tr>';
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
        k + (CatPrefs.showChipCounts ? '<span class="cat-chip-n">' + letters[k] + '</span>' : '') + '</button>').join('') +
    '</div>' +
    '<div class="cat-filter-row"><span class="cat-flabel">Section:</span>' +
      '<button class="cat-chip' + (Catalog.section === 'all' ? ' active' : '') + '" data-section="all">All</button>' +
      sectionKeys.map(k => '<button class="cat-chip' + (Catalog.section === k ? ' active' : '') + '" data-section="' + escapeHtml(k) + '">' +
        escapeHtml(truncateLabel(k, 18)) + (CatPrefs.showChipCounts ? '<span class="cat-chip-n">' + sections[k] + '</span>' : '') + '</button>').join('') +
    '</div>' +
    '<div class="cat-filter-row"><span class="cat-flabel">Sub-cat:</span>' +
      '<button class="cat-chip' + (Catalog.sub === 'all' ? ' active' : '') + '" data-sub="all">All</button>' +
      subKeys.map(k => '<button class="cat-chip' + (Catalog.sub === k ? ' active' : '') + '" data-sub="' + escapeHtml(k) + '">' +
        escapeHtml(truncateLabel(k, 20)) + (CatPrefs.showChipCounts ? '<span class="cat-chip-n">' + subs[k] + '</span>' : '') + '</button>').join('') +
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
  // Full-details window, exactly like the pop-out on Product Performance.
  host.querySelectorAll('.cat-popout').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      openCatalogDrill(btn.closest('tr'));
    });
  });
  // Every row is clickable, at any depth. Previously only the top level and
  // the old "cat-colour" class were wired, so nested rows did nothing.
  host.querySelectorAll('tr.cat-row').forEach(tr => {
    tr.addEventListener('click', () => {
      if (!Behaviour.inlineExpand) { openCatalogDrill(tr); return; }
      const c = tr.querySelector('.cat-caret');
      if (c) c.click();
      else openCatalogDrill(tr);   // deepest level has nothing left to open
    });
  });
}

/** Opens the drill window for a catalog row. The row's data-key is already
 *  stored as "Field=Value|Field=Value|...", the same shape the drill window
 *  wants, so the whole branch carries across instead of just the last level. */
function openCatalogDrill(tr) {
  if (!tr || !tr.dataset.key) return;
  const path = String(tr.dataset.key).split('|').filter(Boolean).map(seg => {
    const i = seg.indexOf('=');
    return { field: seg.slice(0, i), value: seg.slice(i + 1) };
  }).filter(p => p.field);
  if (path.length) openDrillPath(path);
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
  const headers = ['Path', 'Level', 'Section', 'Sub Section', 'Brand', 'Supplier',
                   'Sold', 'Purchased', 'Opening', 'Closing',
                   'ADC', 'LT', 'SF', 'MOQ', 'ML', 'MIT', 'Stock %', 'Reorder', 'Status'];
  const out = [];
  const days = catalogDays();
  (function walk(nodes, trail) {
    nodes.forEach(n => {
      const r = replenFor(n.path, n.sold, days, n.cbs, n);
      const line = trail.concat([n.key]);
      out.push([
        line.join(' > '), n.dim || '', n.meta.Section || '', n.meta['Sub Section'] || '',
        n.meta.Brand || '', n.meta.Supplier || '',
        n.sold, n.purchased, n.hasOBS ? n.obs : '', n.cbs,
        Number(r.adc.toFixed(3)), r.lt, r.sf, r.moq, Math.round(r.ml), r.mit,
        Math.round(r.pct), r.reorder, n.status
      ]);
      if (n.childList && n.childList.length) walk(n.childList, line);
    });
  })(Catalog.lastRows, []);

  downloadBlob(toCSV(headers, out), 'catalog.csv', 'text/csv');
}

/* ---------------------------------------------------------------
   15b. CATALOG PREFERENCES + IMAGE VIEWER
   --------------------------------------------------------------- */

const CATALOG_LEVEL_DIMS = ['Article No', 'Item Code', 'Brand', 'Colour', 'Size', 'Style',
                            'Section', 'Sub Section', 'Supplier'];

const CAT_COLUMNS = [
  ['category', 'Category'], ['colours', 'Colours'], ['sold', 'Sold'], ['purchased', 'Purchased'],
  ['opening', 'Opening'], ['closing', 'Closing'],
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

/** The max level of a row.
 *
 *  A summary row like "T-SHIRT" is not one garment - it covers hundreds of
 *  article/colour/size combinations, and each of those needs its own minimum on
 *  the shelf. So the row's max level is the sum of its units' max levels:
 *
 *      ML = SUM over units of  max( that unit's ADC x LT x SF , MOQ )
 *
 *  Two earlier versions of this got it wrong. Up to v38 the whole row was
 *  floored at a single MOQ, so T-SHIRT came out at 151 against 4,582 in stock,
 *  Stock % read 3,033%, and every row on the page turned the same "Overstock"
 *  colour. v39 floored it at MOQ x unit count, which fixed the percentages but
 *  made the sum jump in one step - the demand side never showed until it beat
 *  the whole floor at once, so typing a new LT or SF moved nothing.
 *
 *  Doing the sum properly fixes both. Fast-selling units pass the minimum one
 *  at a time as the lead time grows, so the total responds to every change,
 *  and a row of slow sellers still sits at its honest floor.
 *
 *  `dist` is the row's sorted unit sales with a running total, built in
 *  buildCatalog. A bare number means "this many units, sales spread evenly";
 *  nothing at all means a single unit, which is the plain textbook formula.
 */
function maxLevelFor(dist, adc, lt, sf, moq, days, soldTotal) {
  const floor = moq > 0 ? moq : 1;
  const sorted = dist && dist.skuSold;
  const units = Math.max(1, sorted ? sorted.length : (typeof dist === 'number' ? dist : (dist && dist.skuCount) || 1));

  // The demand each unit has to cover. An ADC typed on the row overrides what
  // was actually sold, so scale the units to match what was typed.
  const wanted = adc * days;                       // what the row "sells" per period
  if (!sorted || !sorted.length || !(soldTotal > 0)) {
    // no spread to work from: treat the units as identical
    const each = Math.max(0, wanted) / units;
    return { ml: units * Math.max((each / days) * lt * sf, floor), units };
  }
  const scale = wanted / soldTotal;                // 1 unless ADC was overridden
  const c = scale * lt * sf / days;                // unit sales x c = its demand cover
  if (!(c > 0)) return { ml: units * floor, units };

  // Units selling less than this can never beat the minimum order, so they sit
  // at the floor; the rest are driven by their own demand. Binary search finds
  // the split, and the running total sums the fast half in one step.
  const threshold = floor / c;
  const cum = dist.skuCum;
  let lo = 0, hi = sorted.length;
  while (lo < hi) { const mid = (lo + hi) >> 1; if (sorted[mid] < threshold) lo = mid + 1; else hi = mid; }
  const ml = floor * lo + c * (cum[sorted.length] - cum[lo]);
  return { ml, units };
}

function replenFor(key, sold, days, cbs, dist) {
  const o = Replen.overrides[key] || {};
  const autoAdc = days > 0 ? sold / days : 0;
  const adc = (o.adc !== undefined && o.adc !== null) ? o.adc
              : ((CatPrefs.defaultADC > 0) ? CatPrefs.defaultADC : autoAdc);
  const lt = (o.lt !== undefined && o.lt !== null) ? o.lt : (CatPrefs.defaultLT || 15);
  const sf = (o.sf !== undefined && o.sf !== null) ? o.sf : (CatPrefs.defaultSF || 1.5);
  const moq = (o.moq !== undefined && o.moq !== null) ? o.moq : (CatPrefs.defaultMOQ || 12);
  const mit = (o.mit !== undefined && o.mit !== null) ? o.mit : 0;

  const rawMl = adc * lt * sf;
  const lvl = maxLevelFor(dist, adc, lt, sf, moq, days, sold);
  const ml = lvl.ml, units = lvl.units;
  // Stock % = Closing stock / Max Level.
  // CBS from the ERP already equals OBS + purchased - sold, so it IS the stock
  // on hand. The old "CBS + purchased - sold" column counted the same movements
  // a second time and was removed in v38.
  // Anything you type into "In transit" is added on top, because that stock is
  // paid for and on its way; leave that column at 0 and this is exactly
  // Closing / Max Level.
  const onHand = (cbs || 0) + mit;
  let pct = (onHand / ml) * 100;
  if (onHand < 0) pct = 0;
  if (!isFinite(pct)) pct = onHand > 0 ? 999 : 0;

  let reorder = 0;
  if (ml > onHand) {
    const gap = ml - onHand;
    reorder = moq > 0 ? Math.ceil(gap / moq) * moq : Math.ceil(gap);
  }
  return { adc, lt, sf, moq, ml, rawMl, mit, onHand, pct, reorder, units,
           isAuto: { adc: o.adc === undefined, lt: o.lt === undefined, sf: o.sf === undefined,
                     moq: o.moq === undefined, mit: o.mit === undefined } };
}

/** 0-33 red, 33-66 yellow, 66-100 green, over 100 purple. */
/* ---- the Stock % bands -------------------------------------------------
   These four bands are the single definition used by the cell colour, the
   legend under the toolbar AND the Status column. Before v43 the Status
   column had its own thresholds, so a row at 40% was coloured "Medium stock"
   by the legend and simultaneously labelled "Healthy" - 139 rows of your
   catalogue carried a word that contradicted their own colour.
   Edit them in Settings > 02 Catalog and all three move together. ------- */
function stockBands() {
  const p = CatPrefs;
  const low = p.bandLow === undefined ? 33 : p.bandLow;
  const mid = p.bandMid === undefined ? 66 : p.bandMid;
  const good = p.bandGood === undefined ? 100 : p.bandGood;
  return { low, mid: Math.max(low, mid), good: Math.max(low, mid, good) };
}

function stockPctClass(pct) {
  const b = stockBands();
  if (pct > b.good) return 'sp-over';
  if (pct > b.mid) return 'sp-good';
  if (pct > b.low) return 'sp-mid';
  return 'sp-low';
}

/** The word for a band - exactly the same four steps as the colour. */
const STOCK_BAND_NAME = { 'sp-low': 'Low stock', 'sp-mid': 'Medium stock',
                          'sp-good': 'Healthy', 'sp-over': 'Overstock' };
function stockBandStatus(pct) { return STOCK_BAND_NAME[stockPctClass(pct)]; }

function replenCells(key, r) {
  const inp = (field, val, step) =>
    '<input class="cat-inp' + (r.isAuto[field] ? ' is-auto' : ' is-set') + '" type="number" step="' + step + '" ' +
    'value="' + val + '" data-rkey="' + escapeHtml(key) + '" data-rfield="' + field + '" ' +
    'title="' + (r.isAuto[field] ? 'Automatic - type to override' : 'Set by you - clear the box to go back to automatic') + '">';

  return (catColOn('adc') ? '<td class="num cat-edit">' + inp('adc', fmtNum(r.adc, 2), '0.01') + '</td>' : '') +
         (catColOn('lt') ? '<td class="num cat-edit">' + inp('lt', r.lt, '1') + '</td>' : '') +
         (catColOn('sf') ? '<td class="num cat-edit">' + inp('sf', r.sf, '0.1') + '</td>' : '') +
         (catColOn('moq') ? '<td class="num cat-edit">' + inp('moq', r.moq, '1') + '</td>' : '') +
         (catColOn('ml') ? '<td class="num cat-ml" title="' + escapeHtml(
             r.units > 1
               ? fmtNum(r.units) + ' items under this row \u00d7 MOQ ' + r.moq + ' = ' + fmtNum(r.units * r.moq, 0) +
                 '; demand (ADC \u00d7 LT \u00d7 SF) would give ' + fmtNum(r.rawMl, 0) + ' \u2014 the larger wins'
               : 'ADC ' + fmtNum(r.adc, 2) + ' \u00d7 LT ' + r.lt + ' \u00d7 SF ' + r.sf +
                 ', at least MOQ ' + r.moq) + '">' + fmtNum(r.ml, 0) + '</td>' : '') +
         (catColOn('mit') ? '<td class="num cat-edit">' + inp('mit', r.mit, '1') + '</td>' : '') +
         (catColOn('stockpct') ? '<td class="num stock-pct ' + stockPctClass(r.pct) +
           '" title="' + escapeHtml('Closing ' + fmtNum(r.onHand - r.mit) +
             (r.mit ? ' + in transit ' + fmtNum(r.mit) : '') +
             ' vs max level ' + fmtNum(r.ml, 0) +
             (r.units > 1 ? ' (' + fmtNum(r.units) + ' items under this row)' : '')) + '">' +
           (r.pct > 999 ? '999%+' : fmtNum(r.pct, 0) + '%') + '</td>' : '') +
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
  showLegend: false,      // the Stock % colour key above the table
  stripUntil: 'opening',  // colour bar under a row stops after this column
  bandLow: 33,            // up to here: Low stock
  bandMid: 66,            // up to here: Medium stock
  bandGood: 100,          // up to here: Healthy; above it: Overstock
  // replenishment defaults
  defaultADC: 0,      // 0 = work it out from sales; anything else is used as-is
  levels: ['Article No', 'Colour', 'Size'],   // the three drill levels, in order
  defaultLT: 30,      // lead time, days
  defaultSF: 1.5,     // safety factor
  defaultMOQ: 12,     // minimum order quantity (also the floor for the max level)
  // table
  columns: CAT_COLUMNS.map(c => c[0]),
  maxRows: 300,
  showChipCounts: false,
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
        // one-off: the grey counts on filter chips are off by default now
        if (saved.showChipCounts === true && !saved.chipCountsMigrated) {
          CatPrefs.showChipCounts = false;
          CatPrefs.chipCountsMigrated = true;
        }
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
  // row height is shared with every other table (Settings > Look & Feel),
  // so nothing catalog-specific is set here any more
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
    ? ((d.childList || []).length + ' items \u00b7 ' + fmtNum(d.cbs) + ' in stock \u00b7 ' + fmtNum(d.sold) + ' sold')
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
    '<p class="drill-subtitle">Row height and drill-down row height are shared with every other table \u2014 set them in <strong>Look &amp; Feel</strong>.</p>' +
    row('Font size',
      '<label class="toolbar-label"></label>' +
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

    '<h3 class="snap-set-title">Drill sequence</h3>' +
    '<p class="drill-subtitle">Which level opens inside which, top to bottom \u2014 the same idea as the Product Performance expand order.</p>' +
    '<div class="snap-levels">' +
      [0, 1, 2, 3, 4, 5, 6, 7, 8].map(function (i) {
        const cur = (CatPrefs.levels && CatPrefs.levels[i]) || '';
        return '<div><label class="toolbar-label">Level ' + (i + 1) + '</label>' +
          '<select class="select cs-level" data-i="' + i + '">' +
            (i > 0 ? '<option value="">\u2014 none \u2014</option>' : '') +
            CATALOG_LEVEL_DIMS.map(function (d) {
              return '<option value="' + d + '"' + (cur === d ? ' selected' : '') + '>' + d + '</option>';
            }).join('') +
          '</select></div>';
      }).join('') +
    '</div>' +
    '<p class="drill-subtitle">Up to nine levels. Leave the rest as \u201cnone\u201d \u2014 the table opens in exactly this order.</p>' +

    '<h3 class="snap-set-title">Stock %</h3>' +
    '<p class="drill-subtitle">Stock % is Closing stock \u00f7 Max Level. ' +
      'Max Level is ADC \u00d7 LT \u00d7 SF, but never less than MOQ for <em>every</em> item under the row \u2014 ' +
      'a summary row like \u201cT-SHIRT\u201d covers hundreds of article/colour/size combinations, ' +
      'and each one needs its own minimum on the shelf. Before v39 those rows were floored at a single MOQ, ' +
      'which made every one of them read 999%+. ' +
      'The old \u201cAvailable\u201d column (CBS + purchased \u2212 sold) was removed in v38: ' +
      'CBS from your ERP already equals OBS + purchased \u2212 sold, so that column counted ' +
      'the same movements twice. Anything you type in the \u201cIn transit\u201d column is added on top.</p>' +

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

    '<h3 class="snap-set-title">Colour code bar</h3>' +
    row('Bar width',
      '<label class="toolbar-label">Stop after</label>' +
      '<select id="cs-stripuntil" class="select">' +
        [['design', 'Design'], ['category', 'Category'], ['colours', 'Colours'],
         ['sold', 'Sold'], ['purchased', 'Purchased'], ['opening', 'Opening'],
         ['closing', 'Closing'], ['ml', 'ML'], ['stockpct', 'Stock %'],
         ['status', 'Status'], ['all', 'the whole row']].map(function (c) {
          return '<option value="' + c[0] + '"' +
            ((CatPrefs.stripUntil || 'opening') === c[0] ? ' selected' : '') + '>' + c[1] + '</option>';
        }).join('') +
      '</select>') +
    '<p class="drill-subtitle">The coloured bar under an opened row used to run from the first ' +
      'column to the last. Pick where it should stop instead. A column that is switched off ' +
      'is skipped, and the bar falls back to the full width.</p>' +

    '<h3 class="snap-set-title">Stock % bands</h3>' +
    row('Colour key',
      '<label class="toolbar-checkbox"><input type="checkbox" id="cs-legend"' +
        (CatPrefs.showLegend ? ' checked' : '') + '> Show the colour key above the table</label>' +
      '<span class="drill-count">the strip listing what each Stock % colour means</span>') +
    (function () {
      const b = stockBands();
      return row('Bands',
        '<label class="toolbar-label">Low stock up to</label>' +
        '<input type="number" id="cs-band-low" class="text-input narrow" min="1" max="998" value="' + b.low + '">%' +
        '<label class="toolbar-label">Medium up to</label>' +
        '<input type="number" id="cs-band-mid" class="text-input narrow" min="2" max="999" value="' + b.mid + '">%' +
        '<label class="toolbar-label">Healthy up to</label>' +
        '<input type="number" id="cs-band-good" class="text-input narrow" min="3" max="2000" value="' + b.good + '">%' +
        '<span class="drill-count">above that: Overstock</span>');
    })() +
    '<p class="drill-subtitle">These four bands drive the Stock % cell colour, the legend above the ' +
      'table and the Status column \u2014 all three from one place, so the word can never disagree ' +
      'with the colour again. Stockout, No sale and Idle come first, whatever the percentage.</p>' +

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

  bind('cs-font', e => {
    CatPrefs.fontSize = parseFloat(e.target.value);
    wrap.querySelector('#cs-font-val').textContent = CatPrefs.fontSize + 'px'; saveCatPrefs();
  });
  bindC('cs-head', e => { CatPrefs.headerStyle = e.target.value; saveCatPrefs(); });
  bindC('cs-zebra', e => { CatPrefs.zebra = e.target.checked; saveCatPrefs(); });
  bindC('cs-grid', e => { CatPrefs.gridLines = e.target.checked; saveCatPrefs(); });
  bind('cs-accent', e => { CatPrefs.accent = e.target.value; saveCatPrefs(); });
  bind('cs-hover', e => { CatPrefs.hoverColor = e.target.value; saveCatPrefs(); });

  wrap.querySelectorAll('.cs-level').forEach(sel => sel.addEventListener('change', () => {
    const picked = [...wrap.querySelectorAll('.cs-level')].map(x => x.value);
    CatPrefs.levels = picked.filter(v => v).filter((v, i, a) => a.indexOf(v) === i);
    if (!CatPrefs.levels.length) CatPrefs.levels = ['Article No'];
    saveCatPrefs(); Catalog.expanded = {}; renderCatalog();
  }));
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
  bindC('cs-stripuntil', e => { CatPrefs.stripUntil = e.target.value; saveCatPrefs(); renderCatalog(); });
  bindC('cs-legend', e => { CatPrefs.showLegend = e.target.checked; saveCatPrefs(); renderCatalog(); });
  bindC('cs-band-low', e => { CatPrefs.bandLow = Math.max(1, parseInt(e.target.value, 10) || 33); saveCatPrefs(); renderCatalog(); });
  bindC('cs-band-mid', e => { CatPrefs.bandMid = Math.max(2, parseInt(e.target.value, 10) || 66); saveCatPrefs(); renderCatalog(); });
  bindC('cs-band-good', e => { CatPrefs.bandGood = Math.max(3, parseInt(e.target.value, 10) || 100); saveCatPrefs(); renderCatalog(); });

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
  ['doughnut', 'Doughnut / donut'],
  ['stackbar', 'Stacked bar'],
  ['stackcolumn', 'Stacked column'],
  ['groupcolumn', 'Grouped column'],
  ['scatter', 'Scatter plot'],
  ['bubble', 'Bubble chart'],
  ['histogram', 'Histogram'],
  ['heatmap', 'Heatmap'],
  ['treemap', 'Treemap'],
  ['sunburst', 'Sunburst'],
  ['gantt', 'Gantt chart'],
  ['waterfall', 'Waterfall'],
  ['radar', 'Radar / spider'],
  ['polar', 'Polar area'],
  ['sankey', 'Sankey diagram'],
  ['funnel', 'Funnel'],
  ['boxplot', 'Box plot (box & whisker)'],
  ['violin', 'Violin plot'],
  ['candlestick', 'Candlestick'],
  ['map-india', 'Map \u2014 India by state'],
  ['map-world', 'Map \u2014 world by country'],
  ['bullet', 'Bullet graph'],
  ['wordcloud', 'Word cloud'],
  ['table', 'Table'],
  ['kpi', 'KPI tile (one big number)']
];

const CHART_SOURCES = [
  ['sales', 'Sales'],
  ['purchase', 'Purchase'],
  ['stock', 'Stock']
];

const CHART_DIMS = ['Article No', 'Brand', 'Colour', 'Size', 'Style', 'Section',
                    'Sub Section', 'Supplier', 'Item Code', 'City', 'Month', 'Week',
                    'Transaction Date'];

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
   16b. CHART BOARDS — move, resize, recolour and edit charts
   ---------------------------------------------------------------
   The same board runs on the Dashboard and on Product Performance.
   Cards can be dragged into a new order, resized from the corner,
   recoloured, edited or removed, and everything is remembered.
   A lock switch in Settings freezes the lot so nothing can be
   changed by mistake.

   Clicking a bar, slice or table row cross-highlights the board the
   way Power BI does: the thing you clicked stays in full colour, the
   rest of that chart fades, every other card re-reads its numbers —
   and nothing is rebuilt, so no card ever jumps out of its place.
   --------------------------------------------------------------- */

const BOARD_DEFS = {
  dash: { key: 'sl_dash_charts', grid: 'dash-grid', filterBar: 'dash-filters' },
  perf: { key: 'sl_perf_charts', grid: 'perf-grid', filterBar: 'perf-chart-filters' }
};

/* ---- colour palettes ----------------------------------------------------
   Every chart can either follow a named palette or carry its own colours.
   The neon sets come from the reference dashboards. ------------------------ */
const CHART_PALETTES = [
  { id: 'rust',    name: 'Rust (house)',  colors: ['#A6402C', '#1F6F5C', '#B9862F', '#4A6FA5', '#7A4CA0', '#C6784B', '#3E8E7E', '#8C5B3F', '#5B7553', '#9C4F6B'] },
  { id: 'neon',    name: 'Neon',          colors: ['#22D3EE', '#E252C4', '#8B5CF6', '#34D399', '#FBBF24', '#F472B6', '#38BDF8', '#A3E635', '#FB7185', '#C084FC'] },
  { id: 'aurora',  name: 'Aurora',        colors: ['#F0509B', '#7B5CFA', '#2DD4BF', '#FDBA4D', '#4CC9F0', '#B14AED', '#54D98C', '#FF6B6B', '#8AB4F8', '#F9A8D4'] },
  { id: 'ocean',   name: 'Ocean',         colors: ['#1D6FA5', '#2E9CCA', '#25CED1', '#5B8FB9', '#0F4C75', '#3282B8', '#57A0D3', '#7FB3D5', '#154C79', '#48A9A6'] },
  { id: 'berry',   name: 'Berry',         colors: ['#8E2C6B', '#C2296F', '#E4467C', '#F06D9B', '#6A2C70', '#B83B5E', '#F08A5D', '#A64C88', '#D6558F', '#7B3F7A'] },
  { id: 'forest',  name: 'Forest',        colors: ['#1F6F5C', '#2F8F5B', '#6BA368', '#A3C48B', '#3E5E3A', '#57876B', '#89B47C', '#2C5E4F', '#7FA98A', '#456B4E'] },
  { id: 'sunset',  name: 'Sunset',        colors: ['#F76B3C', '#F9A03F', '#FBCB43', '#E14D2A', '#C1462F', '#FF8C61', '#FFB55A', '#D94F2B', '#F4A259', '#B8412B'] },
  { id: 'mono',    name: 'Monochrome',    colors: ['#2B3947', '#455B6E', '#5E7A91', '#7B96AC', '#9BB2C4', '#B7C8D6', '#3A4C5C', '#6B879E', '#8CA5B9', '#A8BECE'] }
];

function paletteById(id) {
  return CHART_PALETTES.find(p => p.id === id) || CHART_PALETTES[0];
}

/** Colours for one chart. A chart with its own `colors` wins; otherwise the
 *  chart's palette; otherwise the board default; otherwise the house colours. */
function chartColours(boardId, cfg) {
  if (cfg.colors && cfg.colors.length) return cfg.colors;
  if (cfg.palette && cfg.palette !== 'board') return paletteById(cfg.palette).colors;
  return paletteById(BoardTheme.palette || 'rust').colors;
}

/** Faded version of a colour, for the bars you did NOT click. */
function fadeColour(hex, alpha) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
  if (!m) return hex;
  return 'rgba(' + parseInt(m[1], 16) + ',' + parseInt(m[2], 16) + ',' +
    parseInt(m[3], 16) + ',' + (alpha === undefined ? 0.22 : alpha) + ')';
}

/* ---- table columns a table card can show ---- */
const BOARD_TABLE_COLS = [
  ['name',  'Name'],
  ['value', 'Value'],
  ['share', 'Share %'],
  ['rank',  'Rank'],
  ['cum',   'Running %']
];
const BOARD_TABLE_COLS_DEFAULT = ['name', 'value', 'share'];

function tableColsOf(cfg) {
  const c = (cfg.tcols && cfg.tcols.length) ? cfg.tcols : BOARD_TABLE_COLS_DEFAULT;
  // "name" always stays, otherwise the rows have nothing to say
  return c.indexOf('name') === -1 ? ['name'].concat(c) : c;
}

const PERF_DEFAULT_CHARTS = [
  { id: 'p1', title: 'Top sellers',          type: 'bar',      source: 'sales', dim: 'Article No',  measure: 'qty', topN: 12, w: 520, h: 300 },
  { id: 'p2', title: 'Stock by sub section', type: 'bar',      source: 'stock', dim: 'Sub Section', measure: 'qty', topN: 12, w: 520, h: 300 },
  { id: 'p3', title: 'Sales by month',       type: 'column',   source: 'sales', dim: 'Month',       measure: 'qty', topN: 24, w: 520, h: 300 },
  { id: 'p4', title: 'Share by section',     type: 'doughnut', source: 'sales', dim: 'Section',     measure: 'qty', topN: 8,  w: 520, h: 300 }
];

const Boards = {
  dash: { charts: [], filters: [], instances: {}, order: {} },
  perf: { charts: [], filters: [], instances: {}, order: {} },
  locked: false,
  drag: null
};

/** Makes one saved chart safe to draw.
 *
 *  A half-written or hand-edited entry used to take the whole dashboard down
 *  with it: boardCardHtml would throw on the bad one, the grid was never
 *  written, and the board sat on the "load your reports" hint until you hit
 *  Reset dashboard. Anything unusable is repaired here instead. */
function sanitizeChart(c, i) {
  if (!c || typeof c !== 'object') return null;
  const known = CHART_TYPES.some(t => t[0] === c.type);
  return {
    id: c.id ? String(c.id) : 'c' + Date.now() + '-' + i,
    title: c.title ? String(c.title) : (c.dim ? 'By ' + c.dim : 'Chart'),
    type: known ? c.type : 'column',
    source: CHART_SOURCES.some(t => t[0] === c.source) ? c.source : 'sales',
    dim: CHART_DIMS.indexOf(c.dim) !== -1 ? c.dim : CHART_DIMS[0],
    measure: CHART_MEASURES.some(m => m[0] === c.measure) ? c.measure : 'qty',
    topN: Math.max(3, Math.min(50, parseInt(c.topN, 10) || 10)),
    split: CHART_DIMS.indexOf(c.split) !== -1 ? c.split : undefined,
    topSplit: Math.max(2, Math.min(12, parseInt(c.topSplit, 10) || 8)),
    palette: c.palette, colors: Array.isArray(c.colors) ? c.colors : undefined,
    multi: !!c.multi,
    tcols: Array.isArray(c.tcols) && c.tcols.length ? c.tcols : undefined,
    w: Math.max(160, parseInt(c.w, 10) || (c.type === 'kpi' ? 240 : 520)),
    h: Math.max(110, parseInt(c.h, 10) || (c.type === 'kpi' ? 150 : 300))
  };
}

function loadBoards() {
  try { Boards.locked = Store.get('sl_boards_locked') === '1'; } catch (e) {}
  ['dash', 'perf'].forEach(id => {
    const def = id === 'dash' ? DEFAULT_CHARTS : PERF_DEFAULT_CHARTS;
    let list;
    try {
      const raw = Store.get(BOARD_DEFS[id].key);
      list = raw ? JSON.parse(raw) : null;
    } catch (e) { list = null; }
    if (!Array.isArray(list)) list = null;
    if (list) {
      const seen = {};
      list = list.map(sanitizeChart).filter(Boolean).map(c => {
        while (seen[c.id]) c.id = c.id + '_';     // duplicate ids collide in the instance map
        seen[c.id] = 1; return c;
      });
    }
    Boards[id].charts = (list && list.length) ? list : def.slice();
  });
}
function saveBoard(id) { Store.set(BOARD_DEFS[id].key, JSON.stringify(Boards[id].charts)); }
function setBoardsLocked(v) { Boards.locked = !!v; Store.set('sl_boards_locked', v ? '1' : '0'); renderAllBoards(); }

/** One place to ask "may I change the board right now?". */
function boardEditable(quiet) {
  if (!Boards.locked) return true;
  if (!quiet) toast('The dashboard is locked \u2014 unlock it in Settings \u203A Charts & layout.');
  return false;
}

function renderAllBoards() {
  renderBoard('dash');
  renderBoard('perf');
}

/* ---- data for one chart -------------------------------------------------
   A chart never filters itself. Clicking "Brand: LEVIS" filters every other
   card, but the Brand chart keeps all its brands on screen and simply
   highlights the one you picked — the Power BI rule. -------------------- */
function boardAggregate(boardId, cfg, opts) {
  const o = opts || {};
  // base = the shape of the data with nothing highlighted at all
  const filters = o.base ? [] : Boards[boardId].filters.filter(f => f.dim !== cfg.dim);
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

/** The categories a chart shows, worked out once with no filters at all and
 *  then remembered. This is what stops bars re-sorting and jumping about
 *  every time you click something. */
function boardOrder(boardId, cfg, force) {
  const B = Boards[boardId];
  const sig = [cfg.source, cfg.dim, cfg.measure, cfg.topN].join('|');
  const cached = B.order[cfg.id];
  if (!force && cached && cached.sig === sig) return cached.keys;
  const keys = boardAggregate(boardId, cfg, { base: true }).map(e => e[0]);
  B.order[cfg.id] = { sig, keys };
  return keys;
}

/** Values for a chart, laid out on its remembered category order. */
function boardSeries(boardId, cfg) {
  const keys = boardOrder(boardId, cfg);
  const map = new Map(boardAggregate(boardId, cfg));
  return { labels: keys, values: keys.map(k => map.get(k) || 0) };
}

/** Which value of this chart's own dimension is currently picked, if any. */
function boardPick(boardId, cfg) {
  const f = Boards[boardId].filters.find(x => x.dim === cfg.dim);
  return f ? f.value : null;
}

/** One number for a KPI tile — the grand total, honouring every filter. */
function boardKpiValue(boardId, cfg) {
  const filters = Boards[boardId].filters;
  const recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));
  if (cfg.measure === 'rows') return recs.length;
  if (cfg.measure === 'items') {
    const s = new Set();
    recs.forEach(r => { const c = r['Item Code'] || r['Article No']; if (c) s.add(String(c)); });
    return s.size;
  }
  return recs.reduce((a, r) => a + recQty(r), 0);
}

/* ---- drawing ------------------------------------------------------------ */

function renderBoard(boardId) {
  const def = BOARD_DEFS[boardId];
  const grid = document.getElementById(def.grid);
  if (!grid) return;
  const B = Boards[boardId];

  // an enlarged card or board lives in the overlay - put it back first
  if (Maxed.boardId === boardId) closeCardMax();

  Object.keys(B.instances).forEach(k => { try { B.instances[k].destroy(); } catch (e) {} });
  B.instances = {};
  B.order = {};        // a full render means the data or the cards changed

  renderBoardFilters(boardId);

  if (!App.datasets.length) {
    grid.innerHTML = '<div class="empty-hint big">Load your reports on the Import tab to build charts.</div>';
    return;
  }

  const locked = Boards.locked;
  grid.className = 'chart-board' + (locked ? ' locked' : '');
  // Built one card at a time on purpose: a single unusable chart shows a small
  // notice in its own tile instead of taking the whole board down.
  grid.innerHTML = B.charts.map((cfg, i) => {
    try { return boardCardHtml(boardId, cfg, i, locked); }
    catch (e) { return boardBrokenCardHtml(cfg, i); }
  }).join('');

  B.charts.forEach(cfg => {
    try { drawBoardChart(boardId, cfg); }
    catch (e) { /* the tile stays, just without its drawing */ }
  });
  wireBoardCard(boardId, grid);
}

function boardCardHtml(boardId, cfg, i, locked) {
  const isKpi = cfg.type === 'kpi';
  return '<div class="board-card' + (isKpi ? ' kpi-card' : '') + '" data-cid="' + cfg.id + '" data-idx="' + i + '"' +
      (locked ? '' : ' draggable="true"') +
      ' style="width:' + (cfg.w || (isKpi ? 240 : 520)) + 'px;height:' + (cfg.h || (isKpi ? 150 : 300)) + 'px">' +
      '<div class="board-head">' +
        (locked ? '' : '<span class="board-grip" title="Drag to move">\u2630</span>') +
        '<h3>' + escapeHtml(cfg.title) + '</h3>' +
        '<span class="board-meta">' + escapeHtml(cfg.source) + ' \u00b7 ' + escapeHtml(cfg.dim) + '</span>' +
        '<span class="spacer"></span>' +
        (locked
          ? '<span class="board-lockicon" title="The dashboard is locked in Settings">\uD83D\uDD12</span>'
          : '<button class="dash-ico" data-act="data" title="Change the data shown">\u25BE</button>' +
            '<button class="dash-ico" data-act="edit" title="Edit">\u270E</button>' +
            '<button class="dash-ico" data-act="dup" title="Duplicate">\u29C9</button>' +
            '<button class="dash-ico" data-act="del" title="Remove">\u2715</button>') +
      '</div>' +
      (locked ? '' : boardQuickBarHtml(cfg)) +
      (cfg.type === 'table'
        ? '<div class="board-body"><table class="data-table dash-mini" id="bt-' + boardId + '-' + cfg.id + '"></table></div>'
        : isKpi
          ? '<div class="board-body kpi-body" id="bk-' + boardId + '-' + cfg.id + '"></div>'
          : isSvgType(cfg.type)
            ? '<div class="board-body svg-body" id="bs-' + boardId + '-' + cfg.id + '"></div>'
            : '<div class="board-body"><canvas id="bc-' + boardId + '-' + cfg.id + '"></canvas></div>') +
      (locked ? '' : '<span class="board-resize" title="Drag to resize"></span>') +
    '</div>';
}

/** Shown in place of a chart that cannot be read at all. */
function boardBrokenCardHtml(cfg, i) {
  return '<div class="board-card" data-cid="broken-' + i + '" style="width:320px;height:150px">' +
    '<div class="board-head"><h3>Chart ' + (i + 1) + '</h3></div>' +
    '<div class="board-body"><div class="sc-empty">This chart could not be read. ' +
      'Edit or remove it, or use Reset dashboard.</div></div></div>';
}

/** The little strip under a card heading: swap the data without opening the
 *  editor. Hidden while the board is locked. */
function boardQuickBarHtml(cfg) {
  const isKpi = cfg.type === 'kpi';
  return '<div class="board-quick" data-open="0">' +
    '<label>Show</label>' +
    '<select class="select mini" data-q="source">' + CHART_SOURCES.map(s =>
      '<option value="' + s[0] + '"' + (cfg.source === s[0] ? ' selected' : '') + '>' + s[1] + '</option>').join('') + '</select>' +
    '<label>' + (isKpi ? 'Of' : 'By') + '</label>' +
    '<select class="select mini" data-q="dim">' + CHART_DIMS.map(d =>
      '<option value="' + d + '"' + (cfg.dim === d ? ' selected' : '') + '>' + d + '</option>').join('') + '</select>' +
    '<select class="select mini" data-q="measure">' + CHART_MEASURES.map(m =>
      '<option value="' + m[0] + '"' + (cfg.measure === m[0] ? ' selected' : '') + '>' + m[1] + '</option>').join('') + '</select>' +
    (isKpi ? '' :
      '<label>Top</label><input type="number" class="text-input narrow" data-q="topN" min="3" max="50" value="' + (cfg.topN || 10) + '">') +
    (needsSplit(cfg.type) || cfg.type === 'sunburst'
      ? '<label>Split</label><select class="select mini" data-q="split">' +
          '<option value="">(none)</option>' +
          CHART_DIMS.map(d => '<option value="' + d + '"' + (cfg.split === d ? ' selected' : '') + '>' + d + '</option>').join('') +
        '</select>'
      : '') +
    '<select class="select mini" data-q="palette" title="Colours">' +
      '<option value="board"' + (!cfg.palette || cfg.palette === 'board' ? ' selected' : '') + '>Board colours</option>' +
      CHART_PALETTES.map(p => '<option value="' + p.id + '"' + (cfg.palette === p.id ? ' selected' : '') + '>' + p.name + '</option>').join('') +
    '</select>' +
  '</div>';
}

function drawBoardChart(boardId, cfg) {
  const B = Boards[boardId];

  if (cfg.type === 'kpi') { drawBoardKpi(boardId, cfg); return; }
  if (isSvgType(cfg.type)) {
    // give the browser a moment to size the box before we measure it
    requestAnimationFrame(() => drawBoardSvgChart(boardId, cfg));
    return;
  }

  const series = boardSeries(boardId, cfg);
  const labels = series.labels, values = series.values;
  const pick = boardPick(boardId, cfg);
  const colours = chartColours(boardId, cfg);

  if (cfg.type === 'table') { drawBoardTable(boardId, cfg, labels, values, pick); return; }

  const canvas = document.getElementById('bc-' + boardId + '-' + cfg.id);
  if (!canvas || typeof Chart === 'undefined') return;

  const isPie = cfg.type === 'pie' || cfg.type === 'doughnut' || cfg.type === 'polar';
  const chartType = CHART_JS_TYPES[cfg.type] || 'bar';
  const stacked = cfg.type === 'stackbar' || cfg.type === 'stackcolumn';
  const grouped = cfg.type === 'groupcolumn';
  const line = cfg.type === 'line' || cfg.type === 'area';
  const perCategory = isPie || cfg.type === 'radar' || cfg.multi;

  let data, indexAxis = cfg.type === 'bar' || cfg.type === 'stackbar' ? 'y' : 'x';
  let clickLabels = labels;

  if (stacked || grouped) {
    // one bar per row, one colour band per value of the second dimension
    const m = boardMatrix(boardId, cfg);
    clickLabels = m.rows;
    data = {
      labels: m.rows,
      datasets: m.cols.map((c, j) => ({
        label: c,
        data: m.rows.map((_, i) => m.grid[i][j]),
        backgroundColor: m.rows.map(rw => scFill(rw, pick, colours[j % colours.length])),
        borderWidth: 0
      }))
    };
  } else if (cfg.type === 'scatter' || cfg.type === 'bubble') {
    // each point is one row: how much was sold against how many rows it took
    const m = boardMatrix(boardId, cfg);
    const pts = m.rows.map((rw, i) => {
      const tot = m.grid[i].reduce((a, b) => a + b, 0);
      const spread = m.grid[i].filter(v => v > 0).length;
      return { x: spread, y: tot, r: cfg.type === 'bubble' ? Math.max(4, Math.sqrt(tot) * 1.1) : 4, _l: rw };
    });
    clickLabels = m.rows;
    data = {
      datasets: [{
        label: cfg.dim,
        data: pts,
        backgroundColor: m.rows.map(rw => scFill(rw, pick, colours[0])),
        borderColor: colours[0], borderWidth: 1
      }]
    };
  } else if (cfg.type === 'histogram') {
    // how the rows themselves are distributed, bucketed
    const vals = boardValuesFor(boardId, cfg);
    const bins = Math.max(5, Math.min(20, Math.round(Math.sqrt(vals.length)) || 8));
    const hi = vals.length ? vals[vals.length - 1] : 1;
    const step = hi / bins || 1;
    const counts = new Array(bins).fill(0);
    vals.forEach(v => { counts[Math.min(bins - 1, Math.floor(v / step))]++; });
    clickLabels = [];
    data = {
      labels: counts.map((_, i) => fmtNum(i * step, 0) + '\u2013' + fmtNum((i + 1) * step, 0)),
      datasets: [{ label: 'How many ' + cfg.dim.toLowerCase() + 's', data: counts,
                   backgroundColor: colours[0], borderWidth: 0 }]
    };
    indexAxis = 'x';
  } else {
    const fill = labels.map((lb, i) => scFill(lb, pick, perCategory ? colours[i % colours.length] : colours[0]));
    data = {
      labels: labels,
      datasets: [{
        label: cfg.title, data: values,
        backgroundColor: line ? fadeColour(colours[0], cfg.type === 'area' ? .22 : .08) : fill,
        borderColor: line || cfg.type === 'radar' ? colours[0] : (pick ? fill : undefined),
        borderWidth: line || cfg.type === 'radar' ? 2 : (pick ? 1.5 : 0),
        pointBackgroundColor: line ? labels.map(lb => (pick && lb !== pick) ? fadeColour(colours[0], .3) : colours[0]) : undefined,
        pointRadius: line ? (pick ? labels.map(lb => lb === pick ? 5 : 2) : 3) : undefined,
        fill: cfg.type === 'area' || cfg.type === 'radar', tension: .25
      }]
    };
  }

  const axis = { ticks: { font: { size: 10 }, color: boardTextColour() }, grid: { color: boardGridColour() } };
  let scales;
  if (isPie) scales = {};
  else if (cfg.type === 'radar') {
    scales = { r: { beginAtZero: true, grid: { color: boardGridColour() },
                    angleLines: { color: boardGridColour() },
                    pointLabels: { color: boardTextColour(), font: { size: 9 } },
                    ticks: { color: boardTextColour(), backdropColor: 'transparent', font: { size: 8 } } } };
  } else if (cfg.type === 'scatter' || cfg.type === 'bubble') {
    scales = {
      x: Object.assign({ title: { display: true, text: 'how many ' + (cfg.split || 'Colour').toLowerCase() + 's it spans',
                                  color: boardTextColour(), font: { size: 9 } } }, axis),
      y: Object.assign({ beginAtZero: true, title: { display: true, text: 'total',
                                  color: boardTextColour(), font: { size: 9 } } }, axis)
    };
  } else {
    scales = { x: Object.assign({ stacked: stacked }, axis),
               y: Object.assign({ beginAtZero: true, stacked: stacked }, axis) };
  }

  B.instances[cfg.id] = makeChart(canvas.getContext('2d'), {
    type: chartType,
    data: data,
    options: {
      responsive: true, maintainAspectRatio: false,
      animation: { duration: 220 },
      indexAxis: indexAxis,
      onClick: (evt, els) => {
        if (!els || !els.length || !clickLabels.length) return;
        const lb = clickLabels[els[0].index];
        if (lb !== undefined) addBoardFilter(boardId, cfg.dim, lb);
      },
      plugins: {
        legend: { display: isPie || stacked || grouped, position: 'right',
          labels: { font: { size: 10 }, boxWidth: 10, color: boardTextColour() } },
        tooltip: {
          callbacks: {
            label: (cfg.type === 'scatter' || cfg.type === 'bubble')
              ? (c => (c.raw && c.raw._l ? c.raw._l + ': ' + fmtNum(c.raw.y) : ''))
              : undefined,
            afterLabel: () => clickLabels.length ? (pick ? 'Click again to clear' : 'Click to highlight') : ''
          }
        }
      },
      scales: scales
    }
  });
}

function drawBoardTable(boardId, cfg, labels, values, pick) {
  const el = document.getElementById('bt-' + boardId + '-' + cfg.id);
  if (!el) return;
  const cols = tableColsOf(cfg);
  const total = values.reduce((a, b) => a + b, 0);
  const head = cols.map(c => {
    if (c === 'name') return '<th>' + escapeHtml(cfg.dim) + '</th>';
    if (c === 'value') return '<th class="num">' + (cfg.measure === 'rows' ? 'Rows' : cfg.measure === 'items' ? 'Items' : 'Qty') + '</th>';
    if (c === 'share') return '<th class="num">Share</th>';
    if (c === 'rank') return '<th class="num">#</th>';
    return '<th class="num">Running</th>';
  }).join('');

  let run = 0;
  const body = labels.map((k, i) => {
    const v = values[i];
    run += v;
    const cells = cols.map(c => {
      if (c === 'name') return '<td>' + escapeHtml(k) + '</td>';
      if (c === 'value') return '<td class="num">' + fmtNum(v) + '</td>';
      if (c === 'share') return '<td class="num">' + (total ? fmtNum(v / total * 100, 1) + '%' : '\u2014') + '</td>';
      if (c === 'rank') return '<td class="num">' + (i + 1) + '</td>';
      return '<td class="num">' + (total ? fmtNum(run / total * 100, 1) + '%' : '\u2014') + '</td>';
    }).join('');
    const cls = 'dash-trow' + (pick ? (k === pick ? ' picked' : ' dimmed') : '');
    return '<tr class="' + cls + '" data-v="' + escapeHtml(k) + '">' + cells + '</tr>';
  }).join('');

  const foot = cols.map((c, i) => {
    if (i === 0) return '<td>Total</td>';
    if (c === 'value') return '<td class="num">' + fmtNum(total) + '</td>';
    if (c === 'share' || c === 'cum') return '<td class="num">100%</td>';
    return '<td></td>';
  }).join('');

  el.innerHTML = '<thead><tr>' + head + '</tr></thead><tbody>' + body +
    '</tbody><tfoot><tr>' + foot + '</tr></tfoot>';
  el.querySelectorAll('.dash-trow').forEach(tr =>
    tr.addEventListener('click', () => addBoardFilter(boardId, cfg.dim, tr.dataset.v)));
}

/** A big-number tile, like the ones across the top of the reference boards. */
function drawBoardKpi(boardId, cfg) {
  const el = document.getElementById('bk-' + boardId + '-' + cfg.id);
  if (!el) return;
  const value = boardKpiValue(boardId, cfg);
  const colours = chartColours(boardId, cfg);
  const series = boardSeries(boardId, cfg);
  const top = series.labels.length
    ? series.labels[0] + ' \u00b7 ' + fmtNum(series.values[0])
    : 'no rows in this window';
  const unit = cfg.measure === 'rows' ? 'rows' : cfg.measure === 'items' ? 'items' : 'pcs';

  el.innerHTML =
    '<div class="kpi-big" style="color:' + colours[0] + '">' + fmtNum(value) + '</div>' +
    '<div class="kpi-unit">' + unit + '</div>' +
    '<div class="kpi-foot">Top ' + escapeHtml(cfg.dim) + ': <strong>' + escapeHtml(top) + '</strong></div>';
}

/* ---- cross-highlighting -------------------------------------------------
   Nothing here rebuilds the grid. The cards keep their exact position and
   size; only the numbers inside them change. --------------------------- */
function addBoardFilter(boardId, dim, value) {
  const F = Boards[boardId].filters;
  const i = F.findIndex(f => f.dim === dim);
  if (i >= 0) { if (F[i].value === value) F.splice(i, 1); else F[i].value = value; }
  else F.push({ dim, value });
  refreshBoardData(boardId);
}

/** Redraws the contents of every card in place. */
function refreshBoardData(boardId) {
  const B = Boards[boardId];
  const grid = document.getElementById(BOARD_DEFS[boardId].grid);
  const live = grid && (grid.querySelector('.board-card') ||
                        Maxed.boardId === boardId || Maxed.whole === boardId);
  if (!live) { renderBoard(boardId); return; }
  renderBoardFilters(boardId);
  B.charts.forEach(cfg => {
    const inst = B.instances[cfg.id];
    if (inst) { try { inst.destroy(); } catch (e) {} delete B.instances[cfg.id]; }
    drawBoardChart(boardId, cfg);
  });
  grid.classList.toggle('has-pick', B.filters.length > 0);
  paintMaxFilters();
  resizeMaxed();
}

function renderBoardFilters(boardId) {
  const wrap = document.getElementById(BOARD_DEFS[boardId].filterBar);
  if (!wrap) return;
  const F = Boards[boardId].filters;
  if (!F.length) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = '<span class="toolbar-label">Highlighting:</span>' +
    F.map((f, i) => '<span class="filter-chip">' + escapeHtml(f.dim) + ': <strong>' +
      escapeHtml(f.value) + '</strong><button data-fi="' + i + '">&times;</button></span>').join('') +
    '<button class="ghost-btn small" data-clear-all="1">Clear all</button>';
  wrap.querySelectorAll('[data-fi]').forEach(b => b.addEventListener('click', () => {
    F.splice(parseInt(b.dataset.fi, 10), 1); refreshBoardData(boardId);
  }));
  const ca = wrap.querySelector('[data-clear-all]');
  if (ca) ca.addEventListener('click', () => { Boards[boardId].filters = []; refreshBoardData(boardId); });
}

/* ---- move, resize, edit, enlarge ---------------------------------------- */
function wireBoardCard(boardId, grid) {
  const B = Boards[boardId];

  grid.querySelectorAll('.dash-ico').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const card = btn.closest('.board-card');
      const i = B.charts.findIndex(c => c.id === card.dataset.cid);
      const act = btn.dataset.act;

      if (act === 'max') { toggleCardMax(boardId, card); return; }
      if (!boardEditable()) return;
      if (act === 'del') { B.charts.splice(i, 1); saveBoard(boardId); renderBoard(boardId); }
      else if (act === 'edit') openBoardChartEditor(boardId, B.charts[i]);
      else if (act === 'data') {
        const q = card.querySelector('.board-quick');
        if (q) q.dataset.open = q.dataset.open === '1' ? '0' : '1';
      } else if (act === 'dup') {
        const copy = Object.assign({}, B.charts[i], { id: 'c' + Date.now(), title: B.charts[i].title + ' (copy)' });
        B.charts.splice(i + 1, 0, copy); saveBoard(boardId); renderBoard(boardId);
      }
    });
  });

  // the quick data / colour strip on each card
  grid.querySelectorAll('.board-quick').forEach(bar => {
    bar.addEventListener('click', e => e.stopPropagation());
    bar.querySelectorAll('[data-q]').forEach(inp => {
      inp.addEventListener('change', e => {
        if (!boardEditable()) return;
        const card = bar.closest('.board-card');
        const cfg = B.charts.find(c => c.id === card.dataset.cid);
        if (!cfg) return;
        const field = e.target.dataset.q;
        if (field === 'topN') cfg.topN = Math.max(3, Math.min(50, parseInt(e.target.value, 10) || 10));
        else cfg[field] = e.target.value;
        delete B.order[cfg.id];          // the categories change, so re-read them
        saveBoard(boardId);
        renderBoard(boardId);
      });
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
        const w = Math.max(200, start.w + (ev.clientX - start.x));
        const h = Math.max(130, start.h + (ev.clientY - start.y));
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

/* ---- enlarge ------------------------------------------------------------
   Fills the browser tab only. No fullscreen call, so the window chrome,
   tab bar and address bar all stay where they are. ---------------------- */
const Maxed = { boardId: null, cardId: null, home: null, next: null,
                whole: null, wholeHome: null, wholeNext: null };

/** A plain fixed panel inside the page. No requestFullscreen anywhere, so the
 *  browser tab bar, address bar and window frame all stay exactly as they are —
 *  the card grows to fill the tab and nothing more. */
function ensureMaxOverlay() {
  let el = document.getElementById('board-max-overlay');
  if (el) return el;
  el = document.createElement('div');
  el.id = 'board-max-overlay';
  el.className = 'board-max-overlay';
  el.style.display = 'none';
  el.innerHTML =
    '<div class="board-max-bar">' +
      '<span class="board-max-title" id="board-max-title"></span>' +
      '<span class="board-max-filters" id="board-max-filters"></span>' +
      '<span class="spacer"></span>' +
      '<button class="ghost-btn small" id="board-max-close" title="Back to the board (Esc)">\u2715 Close</button>' +
    '</div>' +
    '<div class="board-max-stage" id="board-max-stage"></div>';
  document.body.appendChild(el);
  const shut = () => { closeCardMax(); closeBoardMax(); };
  el.querySelector('#board-max-close').addEventListener('click', shut);
  el.addEventListener('click', e => { if (e.target === el) shut(); });
  return el;
}

/** Enlarges the entire grid, not one card. Same fixed panel, same rule: it
 *  fills the browser tab and nothing more, so the tab bar and address bar stay
 *  put. Esc or Close brings it back. */
function toggleBoardMax(boardId) {
  if (Maxed.cardId) closeCardMax();
  if (Maxed.whole === boardId) { closeBoardMax(); return; }
  if (Maxed.whole) closeBoardMax();

  const grid = document.getElementById(BOARD_DEFS[boardId].grid);
  if (!grid) return;
  const ov = ensureMaxOverlay();
  Maxed.whole = boardId;
  Maxed.wholeHome = grid.parentNode;
  Maxed.wholeNext = grid.nextSibling;

  grid.classList.add('board-maxed-grid');
  document.getElementById('board-max-stage').appendChild(grid);
  document.getElementById('board-max-title').textContent =
    (boardId === 'dash' ? 'Dashboard' : 'Product Performance') + ' \u2014 full view';
  ov.style.display = 'flex';
  document.body.classList.add('card-maxed');
  paintMaxFilters();
  setTimeout(() => {
    const B = Boards[boardId];
    Object.keys(B.instances).forEach(k => {
      const inst = B.instances[k];
      if (inst && inst.resize) { try { inst.resize(); } catch (e) {} }
    });
  }, 60);
}

function closeBoardMax() {
  if (!Maxed.whole) return false;
  const grid = document.querySelector('.board-maxed-grid');
  if (grid) {
    grid.classList.remove('board-maxed-grid');
    if (Maxed.wholeHome) Maxed.wholeHome.insertBefore(grid, Maxed.wholeNext);
  }
  const ov = document.getElementById('board-max-overlay');
  if (ov) ov.style.display = 'none';
  document.body.classList.remove('card-maxed');
  const boardId = Maxed.whole;
  Maxed.whole = null; Maxed.wholeHome = null; Maxed.wholeNext = null;
  setTimeout(() => {
    const B = Boards[boardId];
    Object.keys(B.instances).forEach(k => {
      const inst = B.instances[k];
      if (inst && inst.resize) { try { inst.resize(); } catch (e) {} }
    });
  }, 60);
  return true;
}

function toggleCardMax(boardId, card) {
  if (Maxed.cardId === card.dataset.cid) { closeCardMax(); return; }
  if (Maxed.cardId) closeCardMax();
  if (Maxed.whole) closeBoardMax();

  const ov = ensureMaxOverlay();
  const cfg = Boards[boardId].charts.find(c => c.id === card.dataset.cid);
  Maxed.boardId = boardId;
  Maxed.cardId = card.dataset.cid;
  Maxed.home = card.parentNode;
  Maxed.next = card.nextSibling;

  card.classList.add('is-max');
  card.removeAttribute('draggable');
  document.getElementById('board-max-stage').appendChild(card);
  document.getElementById('board-max-title').textContent = cfg ? cfg.title : '';
  ov.style.display = 'flex';
  document.body.classList.add('card-maxed');
  paintMaxFilters();
  resizeMaxed();
}

function closeCardMax() {
  if (!Maxed.cardId) return false;
  const card = document.querySelector('.board-card.is-max');
  if (card) {
    card.classList.remove('is-max');
    if (!Boards.locked) card.setAttribute('draggable', 'true');
    if (Maxed.home) Maxed.home.insertBefore(card, Maxed.next);
  }
  const ov = document.getElementById('board-max-overlay');
  if (ov) ov.style.display = 'none';
  document.body.classList.remove('card-maxed');
  const boardId = Maxed.boardId;
  Maxed.boardId = null; Maxed.cardId = null; Maxed.home = null; Maxed.next = null;
  if (boardId) setTimeout(() => {
    const inst = Boards[boardId].instances[card && card.dataset.cid];
    if (inst && inst.resize) { try { inst.resize(); } catch (e) {} }
  }, 60);
  return true;
}

function resizeMaxed() {
  if (!Maxed.cardId) return;
  // let the layout settle before Chart.js measures the new box
  setTimeout(() => {
    const inst = Boards[Maxed.boardId] && Boards[Maxed.boardId].instances[Maxed.cardId];
    if (inst && inst.resize) { try { inst.resize(); } catch (e) {} }
  }, 60);
}

/** The enlarged card shows the same highlight chips as the board behind it. */
function paintMaxFilters() {
  const el = document.getElementById('board-max-filters');
  if (!el || !Maxed.boardId) return;
  const F = Boards[Maxed.boardId].filters;
  el.innerHTML = F.length
    ? F.map(f => '<span class="filter-chip">' + escapeHtml(f.dim) + ': <strong>' +
        escapeHtml(f.value) + '</strong></span>').join('')
    : '';
}

/* ---- the chart editor --------------------------------------------------- */
function openBoardChartEditor(boardId, cfg) {
  if (!boardEditable()) return;
  const editing = !!cfg;
  const B = Boards[boardId];
  const c = cfg || { id: 'c' + Date.now(), title: '', type: 'column', source: 'sales',
                     dim: 'Brand', measure: 'qty', topN: 10, w: 520, h: 300,
                     palette: 'board', tcols: BOARD_TABLE_COLS_DEFAULT.slice() };
  const cols = tableColsOf(c);
  const own = (c.colors && c.colors.length) ? c.colors : chartColours(boardId, c);

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
    '<div class="settings-row"><label class="toolbar-label">Split by</label>' +
      '<select id="ce-split" class="select">' +
        '<option value="">(none)</option>' +
        CHART_DIMS.map(d => '<option value="' + d + '"' + (c.split === d ? ' selected' : '') + '>' + d + '</option>').join('') +
      '</select>' +
      '<label class="toolbar-label">up to</label>' +
      '<input type="number" id="ce-topsplit" class="text-input narrow" min="2" max="12" value="' + (c.topSplit || 8) + '">' +
      '<span class="drill-count">the second dimension \u2014 stacked, grouped, heatmap, sankey, bubble ' +
        'and sunburst use it; the rest ignore it</span></div>' +

    '<h4 class="ce-sub">Colours</h4>' +
    '<div class="settings-row"><label class="toolbar-label">Palette</label>' +
      '<select id="ce-palette" class="select">' +
        '<option value="board"' + (!c.palette || c.palette === 'board' ? ' selected' : '') + '>Follow the board</option>' +
        CHART_PALETTES.map(p => '<option value="' + p.id + '"' + (c.palette === p.id ? ' selected' : '') + '>' + p.name + '</option>').join('') +
      '</select>' +
      '<label class="toolbar-checkbox"><input type="checkbox" id="ce-multi"' + (c.multi ? ' checked' : '') +
        '> A different colour for every bar</label></div>' +
    '<div class="ce-swatches" id="ce-swatches">' +
      own.slice(0, 6).map((hex, i) =>
        '<span class="ce-sw"><input type="color" data-ci="' + i + '" value="' + hex + '"></span>').join('') +
      '<button class="ghost-btn small" id="ce-colreset">Use the palette</button>' +
    '</div>' +

    '<h4 class="ce-sub">Table columns</h4>' +
    '<div class="settings-row ce-tcols">' + BOARD_TABLE_COLS.map(([k, label]) =>
      '<label class="toolbar-checkbox"><input type="checkbox" data-tc="' + k + '"' +
        (cols.indexOf(k) !== -1 ? ' checked' : '') + (k === 'name' ? ' disabled' : '') + '> ' + label + '</label>').join('') +
      '<span class="drill-count">only used when the type is Table</span></div>' +

    '<div class="settings-row"><label class="toolbar-label">Size</label>' +
      '<input type="number" id="ce-w" class="text-input narrow" min="200" step="20" value="' + (c.w || 520) + '"> \u00d7 ' +
      '<input type="number" id="ce-h" class="text-input narrow" min="130" step="20" value="' + (c.h || 300) + '"> px' +
      '<span class="drill-count">or drag the corner of the card</span></div>' +
    '<div class="modal-actions"><button class="ghost-btn primary small" id="ce-save">' +
      (editing ? 'Save chart' : 'Add chart') + '</button>' +
      '<span class="spacer"></span><button class="ghost-btn small" id="ce-cancel">Cancel</button></div>' +
    '</div>';
  document.body.appendChild(pop);

  let customColors = (c.colors && c.colors.length) ? c.colors.slice() : null;
  pop.querySelectorAll('#ce-swatches input[type="color"]').forEach(inp => {
    inp.addEventListener('input', e => {
      if (!customColors) customColors = own.slice();
      customColors[parseInt(e.target.dataset.ci, 10)] = e.target.value;
    });
  });
  pop.querySelector('#ce-colreset').onclick = () => {
    customColors = null;
    const pal = pop.querySelector('#ce-palette').value;
    const list = pal === 'board' ? paletteById(BoardTheme.palette || 'rust').colors : paletteById(pal).colors;
    pop.querySelectorAll('#ce-swatches input[type="color"]').forEach((inp, i) => { inp.value = list[i] || '#888888'; });
  };
  pop.querySelector('#ce-palette').addEventListener('change', () => {
    if (!customColors) pop.querySelector('#ce-colreset').click();
  });

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
    const sp = pop.querySelector('#ce-split').value;
    if (sp) c.split = sp; else delete c.split;
    c.topSplit = Math.max(2, Math.min(12, parseInt(pop.querySelector('#ce-topsplit').value, 10) || 8));
    c.palette = pop.querySelector('#ce-palette').value;
    c.multi = pop.querySelector('#ce-multi').checked;
    if (customColors) c.colors = customColors; else delete c.colors;
    c.tcols = [...pop.querySelectorAll('[data-tc]')].filter(x => x.checked).map(x => x.dataset.tc);
    if (!c.tcols.length) c.tcols = BOARD_TABLE_COLS_DEFAULT.slice();
    c.w = Math.max(200, parseInt(pop.querySelector('#ce-w').value, 10) || 520);
    c.h = Math.max(130, parseInt(pop.querySelector('#ce-h').value, 10) || 300);
    if (!editing) B.charts.push(c);
    delete B.order[c.id];
    saveBoard(boardId);
    pop.remove();
    renderBoard(boardId);
    toast(editing ? 'Chart updated.' : 'Chart added.');
  };
}

/* ---------------------------------------------------------------
   16c. READY-MADE DASHBOARDS
   ---------------------------------------------------------------
   Six finished layouts you can drop straight onto the Dashboard.
   Applying one replaces the cards and the colour scheme; after that
   every card is a normal card — edit it, recolour it, swap its data,
   add more. Nothing here is fixed.
   --------------------------------------------------------------- */

const DASH_TEMPLATES = [
  {
    id: 'salesstock',
    name: 'Sales & Stock',
    note: 'Big numbers across the top, a sold-vs-purchased trend, and the section split. The closest match to a retail control board.',
    theme: { bg: '#141A2E', card: '#1C2440', fg: '#E8EDF7', grid: '#2C3757', palette: 'neon' },
    charts: [
      { id: 't1a', title: 'Units sold',        type: 'kpi',      source: 'sales',    dim: 'Article No',  measure: 'qty',   w: 250, h: 150 },
      { id: 't1b', title: 'Units purchased',   type: 'kpi',      source: 'purchase', dim: 'Article No',  measure: 'qty',   w: 250, h: 150, palette: 'aurora' },
      { id: 't1c', title: 'Items in stock',    type: 'kpi',      source: 'stock',    dim: 'Item Code',   measure: 'items', w: 250, h: 150, palette: 'ocean' },
      { id: 't1d', title: 'Stock on hand',     type: 'kpi',      source: 'stock',    dim: 'Sub Section', measure: 'qty',   w: 250, h: 150, palette: 'berry' },
      { id: 't1e', title: 'Units sold by month',   type: 'area',     source: 'sales', dim: 'Month',   measure: 'qty', topN: 24, w: 700, h: 330 },
      { id: 't1f', title: 'Sales by product section', type: 'doughnut', source: 'sales', dim: 'Section', measure: 'qty', topN: 8, w: 400, h: 330, multi: true },
      { id: 't1g', title: 'Stock by sub section',  type: 'bar',      source: 'stock', dim: 'Sub Section', measure: 'qty', topN: 10, w: 540, h: 320, palette: 'ocean' },
      { id: 't1h', title: 'Top selling articles',  type: 'table',    source: 'sales', dim: 'Article No',  measure: 'qty', topN: 12, w: 560, h: 320, tcols: ['name', 'value', 'share', 'rank'] }
    ]
  },
  {
    id: 'midnight',
    name: 'Midnight Analytics',
    note: 'Deep purple with ring charts and a wide trend, in the style of the neon analytics boards.',
    theme: { bg: '#1A0F2E', card: '#251643', fg: '#EDE6FA', grid: '#3A2560', palette: 'aurora' },
    charts: [
      { id: 't2a', title: 'Sold',              type: 'kpi',      source: 'sales',    dim: 'Brand',       measure: 'qty', w: 240, h: 145 },
      { id: 't2b', title: 'Purchased',         type: 'kpi',      source: 'purchase', dim: 'Brand',       measure: 'qty', w: 240, h: 145, palette: 'neon' },
      { id: 't2c', title: 'Distinct items',    type: 'kpi',      source: 'sales',    dim: 'Item Code',   measure: 'items', w: 240, h: 145, palette: 'berry' },
      { id: 't2d', title: 'Sales trend',       type: 'line',     source: 'sales',    dim: 'Month',       measure: 'qty', topN: 24, w: 760, h: 340 },
      { id: 't2e', title: 'Share by brand',    type: 'doughnut', source: 'sales',    dim: 'Brand',       measure: 'qty', topN: 8,  w: 360, h: 340, multi: true },
      { id: 't2f', title: 'Sales by colour',   type: 'doughnut', source: 'sales',    dim: 'Colour',      measure: 'qty', topN: 8,  w: 360, h: 300, multi: true, palette: 'neon' },
      { id: 't2g', title: 'Sales by size',     type: 'column',   source: 'sales',    dim: 'Size',        measure: 'qty', topN: 12, w: 520, h: 300, palette: 'berry' },
      { id: 't2h', title: 'Top suppliers',     type: 'bar',      source: 'purchase', dim: 'Supplier',    measure: 'qty', topN: 10, w: 480, h: 300, palette: 'ocean' }
    ]
  },
  {
    id: 'executive',
    name: 'Executive Light',
    note: 'A bright board with one headline chart and a row of smaller ones. Good for printing and for sharing.',
    theme: { bg: '#F4F6FA', card: '#FFFFFF', fg: '#1F2933', grid: '#E1E6EC', palette: 'ocean' },
    charts: [
      { id: 't3a', title: 'Units sold',        type: 'kpi',    source: 'sales',    dim: 'Section',     measure: 'qty', w: 250, h: 145 },
      { id: 't3b', title: 'Units purchased',   type: 'kpi',    source: 'purchase', dim: 'Section',     measure: 'qty', w: 250, h: 145, palette: 'forest' },
      { id: 't3c', title: 'Sales vs month',    type: 'column', source: 'sales',    dim: 'Month',       measure: 'qty', topN: 24, w: 720, h: 330 },
      { id: 't3d', title: 'Purchases by month', type: 'column', source: 'purchase', dim: 'Month',      measure: 'qty', topN: 24, w: 520, h: 300, palette: 'forest' },
      { id: 't3e', title: 'Top brands',        type: 'bar',    source: 'sales',    dim: 'Brand',       measure: 'qty', topN: 10, w: 500, h: 300 },
      { id: 't3f', title: 'Sales by section',  type: 'pie',    source: 'sales',    dim: 'Section',     measure: 'qty', topN: 8,  w: 400, h: 300, multi: true, palette: 'sunset' },
      { id: 't3g', title: 'Stock summary',     type: 'table',  source: 'stock',    dim: 'Sub Section', measure: 'qty', topN: 14, w: 560, h: 320, tcols: ['name', 'value', 'share', 'cum'] }
    ]
  },
  {
    id: 'compact',
    name: 'Small Multiples',
    note: 'Many small panels at once — every dimension on one screen, useful when you want the whole picture in a glance.',
    theme: { bg: '#EDF3FA', card: '#FFFFFF', fg: '#1B2A3A', grid: '#DCE6F0', palette: 'ocean' },
    charts: [
      { id: 't4a', title: 'By brand',       type: 'column', source: 'sales', dim: 'Brand',       measure: 'qty', topN: 8,  w: 330, h: 230 },
      { id: 't4b', title: 'By section',     type: 'column', source: 'sales', dim: 'Section',     measure: 'qty', topN: 8,  w: 330, h: 230 },
      { id: 't4c', title: 'By sub section', type: 'column', source: 'sales', dim: 'Sub Section', measure: 'qty', topN: 8,  w: 330, h: 230 },
      { id: 't4d', title: 'By colour',      type: 'column', source: 'sales', dim: 'Colour',      measure: 'qty', topN: 8,  w: 330, h: 230, multi: true },
      { id: 't4e', title: 'By size',        type: 'column', source: 'sales', dim: 'Size',        measure: 'qty', topN: 10, w: 330, h: 230 },
      { id: 't4f', title: 'By supplier',    type: 'bar',    source: 'sales', dim: 'Supplier',    measure: 'qty', topN: 8,  w: 330, h: 230 },
      { id: 't4g', title: 'Sales by month', type: 'line',   source: 'sales', dim: 'Month',       measure: 'qty', topN: 24, w: 680, h: 260 },
      { id: 't4h', title: 'Stock by section', type: 'doughnut', source: 'stock', dim: 'Section', measure: 'qty', topN: 8, w: 330, h: 260, multi: true }
    ]
  },
  {
    id: 'controlroom',
    name: 'Control Room',
    note: 'Dense and dark, the way a wall display looks. Lots of panels, high-contrast colours, no wasted space.',
    theme: { bg: '#0E1116', card: '#181D25', fg: '#E6EAF0', grid: '#2A313C', palette: 'neon' },
    charts: [
      { id: 't5a', title: 'Sold',          type: 'kpi',    source: 'sales',    dim: 'Brand',      measure: 'qty', w: 210, h: 130 },
      { id: 't5b', title: 'Purchased',     type: 'kpi',    source: 'purchase', dim: 'Brand',      measure: 'qty', w: 210, h: 130, palette: 'aurora' },
      { id: 't5c', title: 'Stock',         type: 'kpi',    source: 'stock',    dim: 'Brand',      measure: 'qty', w: 210, h: 130, palette: 'ocean' },
      { id: 't5d', title: 'Lines sold',    type: 'kpi',    source: 'sales',    dim: 'Item Code',  measure: 'rows', w: 210, h: 130, palette: 'berry' },
      { id: 't5e', title: 'Sales trend',   type: 'area',   source: 'sales',    dim: 'Month',      measure: 'qty', topN: 24, w: 620, h: 280 },
      { id: 't5f', title: 'Purchase trend', type: 'area',  source: 'purchase', dim: 'Month',      measure: 'qty', topN: 24, w: 620, h: 280, palette: 'aurora' },
      { id: 't5g', title: 'Top articles',  type: 'bar',    source: 'sales',    dim: 'Article No', measure: 'qty', topN: 12, w: 410, h: 300 },
      { id: 't5h', title: 'Top colours',   type: 'bar',    source: 'sales',    dim: 'Colour',     measure: 'qty', topN: 12, w: 410, h: 300, multi: true },
      { id: 't5i', title: 'Top sizes',     type: 'bar',    source: 'sales',    dim: 'Size',       measure: 'qty', topN: 12, w: 410, h: 300, palette: 'ocean' },
      { id: 't5j', title: 'Article detail', type: 'table', source: 'sales',    dim: 'Article No', measure: 'qty', topN: 15, w: 640, h: 300, tcols: ['name', 'rank', 'value', 'share', 'cum'] }
    ]
  },
  {
    id: 'classic',
    name: 'Retail Classic',
    note: 'The standard StockLedger board on paper colours — what the Dashboard has always looked like.',
    theme: { bg: '#F6F1E4', card: '#FFFDF8', fg: '#241C14', grid: '#E4DBC6', palette: 'rust' },
    charts: null   // null means "the built-in default set"
  }
];

function applyDashTemplate(tplId) {
  if (!boardEditable()) return;
  const t = DASH_TEMPLATES.find(x => x.id === tplId);
  if (!t) return;
  Boards.dash.charts = t.charts ? JSON.parse(JSON.stringify(t.charts)) : DEFAULT_CHARTS.slice();
  Boards.dash.filters = [];
  Boards.dash.order = {};
  saveBoard('dash');
  if (t.theme) {
    Object.assign(BoardTheme, t.theme, { preset: 'template:' + t.id });
    saveBoardTheme();      // this repaints every board for us
  } else renderBoard('dash');
  toast('"' + t.name + '" applied. Every card is still yours to edit.');
}

/** A tiny CSS mock of each layout, so the gallery shows the shape and the
 *  colours before you commit to it. */
function templateThumb(t) {
  const pal = paletteById((t.theme && t.theme.palette) || 'rust').colors;
  const shapes = [
    'k', 'k', 'k', 'wide', 'half', 'half'
  ];
  return '<span class="tpl-thumb" style="background:' + (t.theme ? t.theme.bg : '#F6F1E4') + '">' +
    shapes.map((s, i) =>
      '<span class="tt ' + s + '" style="background:' + (t.theme ? t.theme.card : '#fff') + '">' +
        '<i style="background:' + pal[i % pal.length] + '"></i></span>').join('') +
  '</span>';
}

function openTemplateGallery() {
  if (!boardEditable()) return;
  const pop = document.createElement('div');
  pop.className = 'modal-backdrop';
  pop.innerHTML = '<div class="modal-box tpl-gallery">' +
    '<h3>Ready-made dashboards</h3>' +
    '<p class="drill-subtitle">Pick a starting point. It sets the cards and the colour scheme in one go \u2014 ' +
      'after that you can edit, recolour, resize or replace anything on it, exactly like a board you built yourself. ' +
      'Your current cards are replaced, so copy anything you want to keep first.</p>' +
    '<div class="tpl-grid">' +
      DASH_TEMPLATES.map(t =>
        '<button class="tpl-card" data-tpl="' + t.id + '">' +
          templateThumb(t) +
          '<strong>' + escapeHtml(t.name) + '</strong>' +
          '<span>' + escapeHtml(t.note) + '</span>' +
          '<span class="tpl-count">' + (t.charts ? t.charts.length : DEFAULT_CHARTS.length) + ' cards</span>' +
        '</button>').join('') +
    '</div>' +
    '<div class="modal-actions"><span class="spacer"></span>' +
      '<button class="ghost-btn small" id="tpl-cancel">Close</button></div>' +
    '</div>';
  document.body.appendChild(pop);

  pop.querySelector('#tpl-cancel').onclick = () => pop.remove();
  pop.addEventListener('click', e => { if (e.target === pop) pop.remove(); });
  pop.querySelectorAll('.tpl-card').forEach(b => b.addEventListener('click', () => {
    applyDashTemplate(b.dataset.tpl);
    pop.remove();
  }));
}

function initBoards() {
  loadBoards();
  [['dash', 'dash-add-chart', 'dash-reset', 'dash-clear-filters', DEFAULT_CHARTS],
   ['perf', 'perf-add-chart', 'perf-reset', 'perf-clear-filters', PERF_DEFAULT_CHARTS]]
  .forEach(([id, addId, resetId, clearId, defaults]) => {
    const add = document.getElementById(addId);
    if (add) add.addEventListener('click', () => openBoardChartEditor(id, null));
    const reset = document.getElementById(resetId);
    if (reset) reset.addEventListener('click', () => {
      if (!boardEditable()) return;
      Boards[id].charts = defaults.slice();
      Boards[id].filters = [];
      Boards[id].order = {};
      saveBoard(id); renderBoard(id);
      toast('Charts reset to the standard set.');
    });
    const clear = document.getElementById(clearId);
    if (clear) clear.addEventListener('click', () => { Boards[id].filters = []; refreshBoardData(id); });
  });

  const tpl = document.getElementById('dash-templates');
  if (tpl) tpl.addEventListener('click', openTemplateGallery);

  ['dash', 'perf'].forEach(id => {
    const b = document.getElementById(id + '-expand');
    if (b) b.addEventListener('click', () => toggleBoardMax(id));
  });

  // Esc leaves an enlarged card or board
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && (Maxed.cardId || Maxed.whole)) {
      e.preventDefault(); closeCardMax(); closeBoardMax();
    }
  });
}

/* ---------------------------------------------------------------
   16f. THE CHART TYPES
   ---------------------------------------------------------------
   Three families, all fed by the same board data and all obeying
   the same cross-highlighting:

     - Chart.js handles bar, column, line, area, pie, doughnut,
       stacked bar/column, scatter, bubble, histogram, radar and
       polar area.
     - Everything else is drawn here as plain SVG, so there are no
       extra libraries to load and the files still work offline.
     - The maps colour the outlines in GEO_INDIA / GEO_WORLD.

   A chart that needs two dimensions (stacked, heatmap, sankey,
   grouped, bubble) reads its second one from cfg.split.
   --------------------------------------------------------------- */

/** Types Chart.js draws for us, mapped to what it calls them. */
const CHART_JS_TYPES = {
  bar: 'bar', column: 'bar', line: 'line', area: 'line',
  pie: 'pie', doughnut: 'doughnut', radar: 'radar', polar: 'polarArea',
  stackbar: 'bar', stackcolumn: 'bar', groupcolumn: 'bar',
  scatter: 'scatter', bubble: 'bubble', histogram: 'bar'
};

/** Types that need a second dimension to mean anything. */
const SPLIT_TYPES = ['stackbar', 'stackcolumn', 'groupcolumn', 'heatmap',
                     'sankey', 'bubble', 'gantt', 'candlestick'];
function needsSplit(type) { return SPLIT_TYPES.indexOf(type) !== -1; }

/** Types drawn by hand, below. */
const SVG_TYPES = ['treemap', 'sunburst', 'funnel', 'waterfall', 'heatmap',
                   'wordcloud', 'bullet', 'boxplot', 'violin', 'sankey',
                   'gantt', 'candlestick', 'map-india', 'map-world'];
function isSvgType(t) { return SVG_TYPES.indexOf(t) !== -1; }

/* ---- two-dimension data -------------------------------------------------
   Rows are the chart's own dimension, columns are cfg.split. Same filter
   rules as everywhere else: a chart never filters itself. -------------- */
function boardMatrix(boardId, cfg) {
  const split = cfg.split || 'Colour';
  const filters = Boards[boardId].filters.filter(f => f.dim !== cfg.dim && f.dim !== split);
  const recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));

  const rowTot = new Map(), colTot = new Map(), cell = new Map();
  recs.forEach(r => {
    const a = dashDimKey(r, cfg.dim), b = dashDimKey(r, split);
    if (a === '(blank)') return;
    const q = cfg.measure === 'rows' ? 1 : recQty(r);
    rowTot.set(a, (rowTot.get(a) || 0) + q);
    colTot.set(b, (colTot.get(b) || 0) + q);
    const k = a + '\u0000' + b;
    cell.set(k, (cell.get(k) || 0) + q);
  });

  const rows = [...rowTot.entries()].sort((x, y) => y[1] - x[1])
    .slice(0, cfg.topN || 10).map(e => e[0]);
  const cols = [...colTot.entries()].sort((x, y) => y[1] - x[1])
    .slice(0, Math.min(cfg.topSplit || 8, 12)).map(e => e[0]);
  const grid = rows.map(a => cols.map(b => cell.get(a + '\u0000' + b) || 0));
  return { rows, cols, grid, split };
}

/* ---- the numbers behind a row, for the distribution charts -------------- */
function boardValuesFor(boardId, cfg) {
  const filters = Boards[boardId].filters.filter(f => f.dim !== cfg.dim);
  const recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));
  const per = new Map();
  recs.forEach(r => {
    const k = dashDimKey(r, cfg.dim);
    if (k === '(blank)') return;
    per.set(k, (per.get(k) || 0) + (cfg.measure === 'rows' ? 1 : recQty(r)));
  });
  return [...per.values()].filter(v => v > 0).sort((a, b) => a - b);
}

/* =========================================================================
   SVG helpers
   ========================================================================= */
function svgOpen(w, h) {
  return '<svg class="svg-chart" viewBox="0 0 ' + w + ' ' + h + '" ' +
    'preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg">';
}
function svgText(x, y, str, cls, anchor, extra) {
  return '<text x="' + x + '" y="' + y + '" class="' + (cls || 'sc-label') + '"' +
    (anchor ? ' text-anchor="' + anchor + '"' : '') + (extra || '') + '>' +
    escapeHtml(String(str)) + '</text>';
}
function scEmpty(msg) { return '<div class="sc-empty">' + escapeHtml(msg) + '</div>'; }

/** Trims a label so it cannot spill out of the shape it sits in. */
function fitLabel(str, px, size) {
  const max = Math.floor(px / (size * 0.56));
  str = String(str);
  return max < 2 ? '' : (str.length <= max ? str : str.slice(0, max - 1) + '\u2026');
}

/** Every SVG chart highlights the same way: picked stays solid, rest fade. */
function scFill(label, pick, colour) { return pick && label !== pick ? fadeColour(colour, .18) : colour; }

/* =========================================================================
   The custom drawings
   ========================================================================= */

/** Treemap - area is the value. Squarified rows, so the tiles stay squarish. */
function svgTreemap(labels, values, colours, pick, W, H) {
  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return scEmpty('Nothing to show yet.');
  const items = labels.map((l, i) => ({ l, v: values[i], c: colours[i % colours.length] }))
    .filter(x => x.v > 0).sort((a, b) => b.v - a.v);

  const tiles = [];
  let x = 0, y = 0, w = W, h = H, i = 0;
  while (i < items.length) {
    const horiz = w >= h;
    const side = horiz ? h : w;
    const left = items.slice(i).reduce((a, b) => a + b.v, 0);
    // grow a row while the aspect ratio keeps improving
    let row = [], best = Infinity;
    for (let j = i; j < items.length; j++) {
      const cand = row.concat([items[j]]);
      const sum = cand.reduce((a, b) => a + b.v, 0);
      const len = (sum / left) * (horiz ? w : h);
      const worst = cand.reduce((m, it) => {
        const thick = (it.v / sum) * side;
        return Math.max(m, Math.max(len / thick, thick / len));
      }, 0);
      if (worst > best && row.length) break;
      best = worst; row = cand;
    }
    const sum = row.reduce((a, b) => a + b.v, 0);
    const len = (sum / left) * (horiz ? w : h);
    let at = 0;
    row.forEach(it => {
      const thick = (it.v / sum) * side;
      tiles.push(horiz ? { x: x, y: y + at, w: len, h: thick, it }
                       : { x: x + at, y: y, w: thick, h: len, it });
      at += thick;
    });
    if (horiz) { x += len; w -= len; } else { y += len; h -= len; }
    i += row.length;
  }

  return svgOpen(W, H) + tiles.map(t => {
    const lab = fitLabel(t.it.l, t.w - 8, 10);
    const val = t.h > 30 ? fitLabel(fmtNum(t.it.v), t.w - 8, 10) : '';
    return '<g class="sc-cell" data-v="' + escapeHtml(t.it.l) + '">' +
      '<title>' + escapeHtml(t.it.l + ': ' + fmtNum(t.it.v) +
        ' (' + fmtNum(t.it.v / total * 100, 1) + '%)') + '</title>' +
      '<rect x="' + t.x.toFixed(1) + '" y="' + t.y.toFixed(1) + '" width="' + Math.max(0, t.w - 1).toFixed(1) +
        '" height="' + Math.max(0, t.h - 1).toFixed(1) + '" fill="' + scFill(t.it.l, pick, t.it.c) + '" rx="2"/>' +
      (lab ? svgText(t.x + 4, t.y + 13, lab, 'sc-label', null, ' fill="#fff"') : '') +
      (val ? svgText(t.x + 4, t.y + 25, val, 'sc-value', null, ' fill="#fff"') : '') +
      '</g>';
  }).join('') + '</svg>';
}

/** Sunburst - one ring here, two when a split dimension is set. */
function svgSunburst(boardId, cfg, labels, values, colours, pick, W, H) {
  const total = values.reduce((a, b) => a + b, 0);
  if (!total) return scEmpty('Nothing to show yet.');
  const cx = W / 2, cy = H / 2, rOuter = Math.min(W, H) / 2 - 6;
  const rInner = rOuter * 0.32;
  const split = cfg.split ? boardMatrix(boardId, cfg) : null;
  const rMid = split ? rInner + (rOuter - rInner) * 0.55 : rOuter;

  const arc = (r0, r1, a0, a1, fill, title, val) => {
    const big = (a1 - a0) > Math.PI ? 1 : 0;
    const p = (r, a) => [(cx + r * Math.cos(a)).toFixed(1), (cy + r * Math.sin(a)).toFixed(1)];
    const [x0, y0] = p(r1, a0), [x1, y1] = p(r1, a1), [x2, y2] = p(r0, a1), [x3, y3] = p(r0, a0);
    return '<g class="sc-cell"' + (val ? ' data-v="' + escapeHtml(val) + '"' : '') + '><title>' + escapeHtml(title) + '</title>' +
      '<path d="M' + x0 + ',' + y0 + 'A' + r1 + ',' + r1 + ' 0 ' + big + ' 1 ' + x1 + ',' + y1 +
      'L' + x2 + ',' + y2 + 'A' + r0 + ',' + r0 + ' 0 ' + big + ' 0 ' + x3 + ',' + y3 + 'Z" fill="' + fill + '"/></g>';
  };

  let a = -Math.PI / 2, out = '';
  labels.forEach((lb, i) => {
    const span = (values[i] / total) * Math.PI * 2;
    if (span <= 0) return;
    const col = colours[i % colours.length];
    out += arc(rInner, rMid, a, a + span, scFill(lb, pick, col),
      lb + ': ' + fmtNum(values[i]) + ' (' + fmtNum(values[i] / total * 100, 1) + '%)', lb);
    if (split) {
      const ri = split.rows.indexOf(lb);
      if (ri !== -1) {
        const rowSum = split.grid[ri].reduce((x, y) => x + y, 0) || 1;
        let b = a;
        split.grid[ri].forEach((v, j) => {
          if (v <= 0) return;
          const sp = span * (v / rowSum);
          out += arc(rMid, rOuter, b, b + sp,
            scFill(lb, pick, shadeColor(col, -18 + (j % 4) * 14)),
            lb + ' \u203a ' + split.cols[j] + ': ' + fmtNum(v), lb);
          b += sp;
        });
      }
    }
    a += span;
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Funnel - stages narrowing down the page, with drop-off between them. */
function svgFunnel(labels, values, colours, pick, W, H) {
  if (!labels.length) return scEmpty('Nothing to show yet.');
  const max = Math.max.apply(null, values) || 1;
  const gap = 4, hh = (H - gap * (labels.length - 1)) / labels.length;
  let y = 0, out = '';
  labels.forEach((lb, i) => {
    const wTop = (values[i] / max) * (W - 90);
    const wBot = (((values[i + 1] !== undefined ? values[i + 1] : values[i])) / max) * (W - 90);
    const cx = (W - 90) / 2 + 88;
    const col = colours[i % colours.length];
    const drop = i > 0 && values[i - 1] > 0 ? (1 - values[i] / values[i - 1]) * 100 : 0;
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ': ' + fmtNum(values[i]) + (i > 0 ? ' (' + fmtNum(drop, 1) + '% down on the step above)' : '')) +
      '</title><path d="M' + (cx - wTop / 2) + ',' + y + 'L' + (cx + wTop / 2) + ',' + y +
      'L' + (cx + wBot / 2) + ',' + (y + hh) + 'L' + (cx - wBot / 2) + ',' + (y + hh) + 'Z" fill="' +
      scFill(lb, pick, col) + '"/></g>' +
      svgText(84, y + hh / 2 + 3, fitLabel(lb, 80, 10), 'sc-label', 'end') +
      svgText(cx, y + hh / 2 + 3, fmtNum(values[i]), 'sc-value', 'middle', ' fill="#fff"');
    y += hh + gap;
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Waterfall - each bar starts where the last one finished. */
function svgWaterfall(labels, values, colours, pick, W, H) {
  if (!labels.length) return scEmpty('Nothing to show yet.');
  const total = values.reduce((a, b) => a + b, 0);
  const padL = 34, padB = 34, padT = 10;
  const cw = (W - padL) / labels.length;
  const scale = (H - padB - padT) / (total || 1);
  let run = 0, out = '<line class="sc-axis" x1="' + padL + '" y1="' + (H - padB) +
    '" x2="' + W + '" y2="' + (H - padB) + '"/>';
  labels.forEach((lb, i) => {
    const y0 = H - padB - run * scale, y1 = H - padB - (run + values[i]) * scale;
    const x = padL + i * cw + cw * 0.15, bw = cw * 0.7;
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ': +' + fmtNum(values[i]) + ', running ' + fmtNum(run + values[i])) + '</title>' +
      '<rect x="' + x.toFixed(1) + '" y="' + Math.min(y0, y1).toFixed(1) + '" width="' + bw.toFixed(1) +
      '" height="' + Math.max(1, Math.abs(y1 - y0)).toFixed(1) + '" fill="' +
      scFill(lb, pick, colours[i % colours.length]) + '" rx="1"/></g>' +
      (i < labels.length - 1 ? '<line class="sc-axis" x1="' + (x + bw) + '" y1="' + y1 +
        '" x2="' + (padL + (i + 1) * cw + cw * 0.15) + '" y2="' + y1 + '" stroke-dasharray="2 2"/>' : '') +
      svgText(x + bw / 2, H - padB + 12, fitLabel(lb, cw, 9), 'sc-tick', 'middle');
    run += values[i];
  });
  out += svgText(padL - 4, padT + 6, fmtNum(total), 'sc-tick', 'end');
  return svgOpen(W, H) + out + '</svg>';
}

/** Heatmap - the chart's dimension down the side, the split across the top. */
function svgHeatmap(m, colours, pick, W, H) {
  if (!m.rows.length || !m.cols.length) return scEmpty('Pick a second dimension to cross this against.');
  const padL = 78, padT = 30;
  const cw = (W - padL) / m.cols.length, ch = (H - padT) / m.rows.length;
  let max = 0;
  m.grid.forEach(r => r.forEach(v => { if (v > max) max = v; }));
  const base = colours[0];
  let out = '';
  m.cols.forEach((c, j) => {
    out += '<g transform="translate(' + (padL + j * cw + cw / 2) + ',' + (padT - 6) + ') rotate(-32)">' +
      svgText(0, 0, fitLabel(c, 62, 9), 'sc-tick', 'start') + '</g>';
  });
  m.rows.forEach((r, i) => {
    out += svgText(padL - 5, padT + i * ch + ch / 2 + 3, fitLabel(r, 72, 9), 'sc-tick', 'end');
    m.cols.forEach((c, j) => {
      const v = m.grid[i][j];
      const t = max ? v / max : 0;
      const fill = v === 0 ? 'transparent' : fadeColour(base, 0.10 + t * 0.9);
      out += '<g class="sc-cell" data-v="' + escapeHtml(r) + '"><title>' +
        escapeHtml(r + ' \u00d7 ' + c + ': ' + fmtNum(v)) + '</title>' +
        '<rect x="' + (padL + j * cw + 1).toFixed(1) + '" y="' + (padT + i * ch + 1).toFixed(1) +
        '" width="' + Math.max(0, cw - 2).toFixed(1) + '" height="' + Math.max(0, ch - 2).toFixed(1) +
        '" fill="' + fill + '" rx="1"' + (pick && r !== pick ? ' class="sc-dim"' : '') + '/></g>';
    });
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Sankey - one dimension flowing into another. */
function svgSankey(m, colours, pick, W, H) {
  if (!m.rows.length || !m.cols.length) return scEmpty('Pick a second dimension to flow into.');
  const padX = 96, nodeW = 11;
  const total = m.grid.reduce((a, r) => a + r.reduce((x, y) => x + y, 0), 0) || 1;
  const gap = 3;
  const lh = (H - gap * (m.rows.length - 1)) / m.rows.length;
  const rh = (H - gap * (m.cols.length - 1)) / m.cols.length;
  const lSum = m.rows.map((_, i) => m.grid[i].reduce((a, b) => a + b, 0));
  const rSum = m.cols.map((_, j) => m.rows.reduce((a, _r, i) => a + m.grid[i][j], 0));

  let out = '', lAt = [], rAt = [];
  let y = 0;
  m.rows.forEach((r, i) => {
    const h = Math.max(2, (lSum[i] / total) * (H - gap * (m.rows.length - 1)));
    lAt[i] = { y: y, cur: y, h: h };
    out += '<rect x="' + padX + '" y="' + y.toFixed(1) + '" width="' + nodeW + '" height="' + h.toFixed(1) +
      '" fill="' + scFill(r, pick, colours[i % colours.length]) + '" rx="2"/>' +
      svgText(padX - 5, y + h / 2 + 3, fitLabel(r, 88, 9), 'sc-tick', 'end');
    y += h + gap;
  });
  y = 0;
  m.cols.forEach((c, j) => {
    const h = Math.max(2, (rSum[j] / total) * (H - gap * (m.cols.length - 1)));
    rAt[j] = { y: y, cur: y, h: h };
    out += '<rect x="' + (W - padX - nodeW) + '" y="' + y.toFixed(1) + '" width="' + nodeW +
      '" height="' + h.toFixed(1) + '" fill="' + shadeColor(colours[(j + 3) % colours.length], -10) + '" rx="2"/>' +
      svgText(W - padX + 5, y + h / 2 + 3, fitLabel(c, 88, 9), 'sc-tick', 'start');
    y += h + gap;
  });

  const x0 = padX + nodeW, x1 = W - padX - nodeW, mid = (x0 + x1) / 2;
  m.rows.forEach((r, i) => {
    m.cols.forEach((c, j) => {
      const v = m.grid[i][j];
      if (v <= 0) return;
      const th = (v / total) * (H - gap * (m.rows.length - 1));
      const a = lAt[i].cur, b = rAt[j].cur;
      lAt[i].cur += th; rAt[j].cur += th;
      out += '<g class="sc-cell" data-v="' + escapeHtml(r) + '"><title>' +
        escapeHtml(r + ' \u2192 ' + c + ': ' + fmtNum(v)) + '</title>' +
        '<path d="M' + x0 + ',' + a.toFixed(1) + 'C' + mid + ',' + a.toFixed(1) + ' ' + mid + ',' + b.toFixed(1) +
        ' ' + x1 + ',' + b.toFixed(1) + 'L' + x1 + ',' + (b + th).toFixed(1) + 'C' + mid + ',' + (b + th).toFixed(1) +
        ' ' + mid + ',' + (a + th).toFixed(1) + ' ' + x0 + ',' + (a + th).toFixed(1) + 'Z" fill="' +
        fadeColour(colours[i % colours.length], pick && r !== pick ? .06 : .40) + '"/></g>';
    });
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Box plot - the spread of the values behind each row. */
function svgBoxplot(vals, labels, colours, pick, W, H, violin) {
  if (!vals.length) return scEmpty('Nothing to show yet.');
  const q = (a, p) => {
    if (!a.length) return 0;
    const i = (a.length - 1) * p, lo = Math.floor(i), hi = Math.ceil(i);
    return a[lo] + (a[hi] - a[lo]) * (i - lo);
  };
  const all = vals.reduce((a, b) => a.concat(b), []);
  const max = Math.max.apply(null, all) || 1;
  const padL = 78, padB = 18, padT = 8;
  const bh = (H - padB - padT) / vals.length;
  const sx = v => padL + (v / max) * (W - padL - 14);
  let out = '';
  vals.forEach((a, i) => {
    const lb = labels[i], col = scFill(lb, pick, colours[i % colours.length]);
    const y = padT + i * bh + bh / 2;
    const q1 = q(a, .25), med = q(a, .5), q3 = q(a, .75);
    const iqr = q3 - q1;
    const lo = Math.max(a[0], q1 - 1.5 * iqr), hi = Math.min(a[a.length - 1], q3 + 1.5 * iqr);
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ' \u2014 min ' + fmtNum(a[0]) + ', lower quarter ' + fmtNum(q1, 1) +
        ', middle ' + fmtNum(med, 1) + ', upper quarter ' + fmtNum(q3, 1) +
        ', max ' + fmtNum(a[a.length - 1]) + ' (' + a.length + ' values)') + '</title>';

    if (violin) {
      // a smoothed outline of how the values bunch up
      const bins = 22, hist = new Array(bins).fill(0);
      a.forEach(v => { hist[Math.min(bins - 1, Math.floor(v / max * bins))]++; });
      const pk = Math.max.apply(null, hist) || 1;
      const step = (W - padL - 14) / bins;
      let up = '', dn = '';
      for (let b = 0; b < bins; b++) {
        const x = padL + b * step + step / 2, t = (hist[b] / pk) * (bh * .40);
        up += (b ? 'L' : 'M') + x.toFixed(1) + ',' + (y - t).toFixed(1);
        dn = 'L' + x.toFixed(1) + ',' + (y + t).toFixed(1) + dn;
      }
      out += '<path d="' + up + dn + 'Z" fill="' + col + '" opacity=".55"/>';
    }
    out += '<line class="sc-axis" x1="' + sx(lo) + '" y1="' + y + '" x2="' + sx(hi) + '" y2="' + y + '"/>' +
      '<rect x="' + sx(q1) + '" y="' + (y - bh * .22) + '" width="' + Math.max(1, sx(q3) - sx(q1)) +
      '" height="' + (bh * .44) + '" fill="' + col + '" rx="1" opacity="' + (violin ? '.85' : '1') + '"/>' +
      '<line x1="' + sx(med) + '" y1="' + (y - bh * .26) + '" x2="' + sx(med) + '" y2="' + (y + bh * .26) +
      '" stroke="#fff" stroke-width="1.6"/></g>' +
      svgText(padL - 5, y + 3, fitLabel(lb, 72, 9), 'sc-tick', 'end');
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Bullet graph - what you have against what the row should hold. */
function svgBullet(labels, values, targets, colours, pick, W, H) {
  if (!labels.length) return scEmpty('Nothing to show yet.');
  const max = Math.max(Math.max.apply(null, values), Math.max.apply(null, targets)) || 1;
  const padL = 84, rowH = Math.min(30, H / labels.length);
  const sx = v => padL + (v / max) * (W - padL - 44);
  let out = '';
  labels.forEach((lb, i) => {
    const y = i * rowH + rowH * 0.18, h = rowH * 0.5;
    const col = scFill(lb, pick, colours[i % colours.length]);
    const hit = targets[i] > 0 ? values[i] / targets[i] * 100 : 0;
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ': ' + fmtNum(values[i]) + ' against a target of ' + fmtNum(targets[i], 0) +
        ' (' + fmtNum(hit, 0) + '%)') + '</title>' +
      '<rect x="' + padL + '" y="' + y + '" width="' + (W - padL - 44) + '" height="' + h +
      '" fill="' + fadeColour(col, .13) + '" rx="2"/>' +
      '<rect x="' + padL + '" y="' + (y + h * .25) + '" width="' + Math.max(1, sx(values[i]) - padL) +
      '" height="' + (h * .5) + '" fill="' + col + '" rx="1"/>' +
      '<line x1="' + sx(targets[i]) + '" y1="' + (y - 2) + '" x2="' + sx(targets[i]) + '" y2="' + (y + h + 2) +
      '" stroke="' + boardTextColour() + '" stroke-width="2"/></g>' +
      svgText(padL - 5, y + h / 2 + 3, fitLabel(lb, 78, 9), 'sc-tick', 'end') +
      svgText(W - 40, y + h / 2 + 3, fmtNum(hit, 0) + '%', 'sc-value', 'start');
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Word cloud - size follows the value, laid out in a simple spiral. */
function svgWordCloud(labels, values, colours, pick, W, H) {
  if (!labels.length) return scEmpty('Nothing to show yet.');
  const max = Math.max.apply(null, values) || 1, min = Math.min.apply(null, values);
  const placed = [];
  let out = '';
  labels.forEach((lb, i) => {
    const t = max === min ? 1 : (values[i] - min) / (max - min);
    const size = 11 + Math.round(Math.pow(t, .6) * 26);
    const wpx = String(lb).length * size * 0.55, hpx = size * 1.15;
    // spiral outwards until it stops overlapping
    let x = W / 2, y = H / 2, ok = false;
    for (let s = 0; s < 900 && !ok; s++) {
      const ang = s * 0.42, rad = s * 0.9;
      x = W / 2 + Math.cos(ang) * rad * (W / H);
      y = H / 2 + Math.sin(ang) * rad;
      if (x - wpx / 2 < 2 || x + wpx / 2 > W - 2 || y - hpx < 2 || y + 2 > H) continue;
      ok = !placed.some(p => Math.abs(p.x - x) * 2 < (p.w + wpx) && Math.abs(p.y - y) * 2 < (p.h + hpx));
    }
    if (!ok) return;
    placed.push({ x, y, w: wpx, h: hpx });
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ': ' + fmtNum(values[i])) + '</title>' +
      '<text x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" text-anchor="middle" ' +
      'font-size="' + size + '" font-weight="' + (size > 22 ? 700 : 500) + '" fill="' +
      scFill(lb, pick, colours[i % colours.length]) + '">' + escapeHtml(String(lb)) + '</text></g>';
  });
  return svgOpen(W, H) + out + '</svg>';
}

/** Gantt - when each item was active, first sale to last. */
function svgGantt(boardId, cfg, colours, pick, W, H) {
  const filters = Boards[boardId].filters.filter(f => f.dim !== cfg.dim);
  const recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));
  const span = new Map();
  recs.forEach(r => {
    if (!r.Date) return;
    const k = dashDimKey(r, cfg.dim);
    if (k === '(blank)') return;
    const t = r.Date.getTime(), e = span.get(k);
    if (!e) span.set(k, { a: t, b: t, q: recQty(r) });
    else { e.a = Math.min(e.a, t); e.b = Math.max(e.b, t); e.q += recQty(r); }
  });
  const rows = [...span.entries()].sort((x, y) => y[1].q - x[1].q).slice(0, cfg.topN || 12);
  if (!rows.length) return scEmpty('This needs dated rows \u2014 try the Sales or Purchase source.');
  const t0 = Math.min.apply(null, rows.map(r => r[1].a));
  const t1 = Math.max.apply(null, rows.map(r => r[1].b)) || t0 + 1;
  const padL = 92, padB = 20, rowH = (H - padB) / rows.length;
  const sx = t => padL + ((t - t0) / (t1 - t0 || 1)) * (W - padL - 10);
  let out = '';
  rows.forEach((e, i) => {
    const [lb, v] = e, y = i * rowH + rowH * .22;
    const days = Math.round((v.b - v.a) / 86400000) + 1;
    out += '<g class="sc-cell" data-v="' + escapeHtml(lb) + '"><title>' +
      escapeHtml(lb + ': ' + fmtDate(new Date(v.a)) + ' to ' + fmtDate(new Date(v.b)) +
        ' \u00b7 ' + days + ' days \u00b7 ' + fmtNum(v.q)) + '</title>' +
      '<rect x="' + sx(v.a).toFixed(1) + '" y="' + y.toFixed(1) + '" width="' +
      Math.max(2, sx(v.b) - sx(v.a)).toFixed(1) + '" height="' + (rowH * .56).toFixed(1) +
      '" fill="' + scFill(lb, pick, colours[i % colours.length]) + '" rx="2"/></g>' +
      svgText(padL - 5, y + rowH * .40, fitLabel(lb, 86, 9), 'sc-tick', 'end');
  });
  out += svgText(padL, H - 5, fmtDate(new Date(t0)), 'sc-tick', 'start') +
         svgText(W - 10, H - 5, fmtDate(new Date(t1)), 'sc-tick', 'end');
  return svgOpen(W, H) + out + '</svg>';
}

/** Candlestick - open, high, low and close of the daily figure, per period. */
function svgCandlestick(boardId, cfg, colours, pick, W, H) {
  const filters = Boards[boardId].filters.filter(f => f.dim !== cfg.dim);
  const recs = dashRecordsFor(cfg.source).filter(r => filters.every(f => dashDimKey(r, f.dim) === f.value));
  const buckets = new Map();
  recs.forEach(r => {
    if (!r.Date) return;
    const k = dashDimKey(r, cfg.dim === 'Month' || cfg.dim === 'Week' ? cfg.dim : 'Month');
    const day = r.Date.toISOString().slice(0, 10);
    let b = buckets.get(k);
    if (!b) { b = new Map(); buckets.set(k, b); }
    b.set(day, (b.get(day) || 0) + recQty(r));
  });
  const keys = [...buckets.keys()].sort((a, b) => grainSort(a, 'month') - grainSort(b, 'month'))
    .slice(-(cfg.topN || 16));
  if (!keys.length) return scEmpty('This needs dated rows \u2014 group by Month or Week.');
  const bars = keys.map(k => {
    const days = [...buckets.get(k).entries()].sort((a, b) => a[0] < b[0] ? -1 : 1).map(e => e[1]);
    return { k, o: days[0], c: days[days.length - 1], h: Math.max.apply(null, days), l: Math.min.apply(null, days) };
  });
  const hi = Math.max.apply(null, bars.map(b => b.h)), lo = Math.min.apply(null, bars.map(b => b.l));
  const padL = 40, padB = 24, padT = 8;
  const cw = (W - padL) / bars.length;
  const sy = v => H - padB - ((v - lo) / ((hi - lo) || 1)) * (H - padB - padT);
  const up = colours[0], down = colours[1] || '#A6402C';
  let out = '<line class="sc-axis" x1="' + padL + '" y1="' + (H - padB) + '" x2="' + W + '" y2="' + (H - padB) + '"/>';
  bars.forEach((b, i) => {
    const x = padL + i * cw + cw / 2, col = scFill(b.k, pick, b.c >= b.o ? up : down);
    out += '<g class="sc-cell" data-v="' + escapeHtml(b.k) + '"><title>' +
      escapeHtml(b.k + ' \u2014 first day ' + fmtNum(b.o) + ', busiest ' + fmtNum(b.h) +
        ', quietest ' + fmtNum(b.l) + ', last day ' + fmtNum(b.c)) + '</title>' +
      '<line x1="' + x + '" y1="' + sy(b.h) + '" x2="' + x + '" y2="' + sy(b.l) + '" stroke="' + col + '"/>' +
      '<rect x="' + (x - cw * .3) + '" y="' + Math.min(sy(b.o), sy(b.c)) + '" width="' + (cw * .6) +
      '" height="' + Math.max(1, Math.abs(sy(b.c) - sy(b.o))) + '" fill="' + col + '"/></g>';
    if (i % Math.ceil(bars.length / 6) === 0) out += svgText(x, H - 8, fitLabel(b.k, cw * 2, 9), 'sc-tick', 'middle');
  });
  out += svgText(padL - 4, padT + 6, fmtNum(hi), 'sc-tick', 'end') +
         svgText(padL - 4, H - padB, fmtNum(lo), 'sc-tick', 'end');
  return svgOpen(W, H) + out + '</svg>';
}

/* =========================================================================
   MAPS
   ========================================================================= */

/** Cities the ERP is likely to name, mapped to their state, so a City column
 *  can drive the India map. Anything unmatched simply is not coloured. */
const CITY_STATE = {
  'new delhi': 'Delhi', 'delhi': 'Delhi', 'noida': 'Uttar Pradesh', 'ghaziabad': 'Uttar Pradesh',
  'gurgaon': 'Haryana', 'gurugram': 'Haryana', 'faridabad': 'Haryana', 'sonipat': 'Haryana',
  'panipat': 'Haryana', 'karnal': 'Haryana', 'ambala': 'Haryana', 'hisar': 'Haryana',
  'mumbai': 'Maharashtra', 'bombay': 'Maharashtra', 'pune': 'Maharashtra', 'nagpur': 'Maharashtra',
  'nashik': 'Maharashtra', 'thane': 'Maharashtra', 'aurangabad': 'Maharashtra',
  'ahmedabad': 'Gujarat', 'surat': 'Gujarat', 'vadodara': 'Gujarat', 'rajkot': 'Gujarat',
  'gandhinagar': 'Gujarat', 'jamnagar': 'Gujarat', 'bhavnagar': 'Gujarat',
  'bangalore': 'Karnataka', 'bengaluru': 'Karnataka', 'mysore': 'Karnataka', 'mysuru': 'Karnataka',
  'hubli': 'Karnataka', 'mangalore': 'Karnataka', 'belgaum': 'Karnataka',
  'chennai': 'Tamil Nadu', 'madras': 'Tamil Nadu', 'coimbatore': 'Tamil Nadu',
  'tirupur': 'Tamil Nadu', 'tiruppur': 'Tamil Nadu', 'madurai': 'Tamil Nadu', 'salem': 'Tamil Nadu',
  'erode': 'Tamil Nadu', 'karur': 'Tamil Nadu',
  'kolkata': 'West Bengal', 'calcutta': 'West Bengal', 'howrah': 'West Bengal', 'siliguri': 'West Bengal',
  'hyderabad': 'Telangana', 'secunderabad': 'Telangana', 'warangal': 'Telangana',
  'visakhapatnam': 'Andhra Pradesh', 'vijayawada': 'Andhra Pradesh', 'guntur': 'Andhra Pradesh',
  'jaipur': 'Rajasthan', 'jodhpur': 'Rajasthan', 'udaipur': 'Rajasthan', 'kota': 'Rajasthan',
  'ajmer': 'Rajasthan', 'bhilwara': 'Rajasthan', 'pali': 'Rajasthan',
  'lucknow': 'Uttar Pradesh', 'kanpur': 'Uttar Pradesh', 'agra': 'Uttar Pradesh',
  'varanasi': 'Uttar Pradesh', 'meerut': 'Uttar Pradesh', 'bareilly': 'Uttar Pradesh',
  'aligarh': 'Uttar Pradesh', 'moradabad': 'Uttar Pradesh', 'gorakhpur': 'Uttar Pradesh',
  'ludhiana': 'Punjab', 'amritsar': 'Punjab', 'jalandhar': 'Punjab', 'patiala': 'Punjab',
  'bathinda': 'Punjab', 'mohali': 'Punjab',
  'indore': 'Madhya Pradesh', 'bhopal': 'Madhya Pradesh', 'gwalior': 'Madhya Pradesh',
  'jabalpur': 'Madhya Pradesh', 'ujjain': 'Madhya Pradesh',
  'patna': 'Bihar', 'gaya': 'Bihar', 'muzaffarpur': 'Bihar', 'bhagalpur': 'Bihar',
  'ranchi': 'Jharkhand', 'jamshedpur': 'Jharkhand', 'dhanbad': 'Jharkhand',
  'bhubaneswar': 'Odisha', 'cuttack': 'Odisha', 'rourkela': 'Odisha',
  'raipur': 'Chhattisgarh', 'bhilai': 'Chhattisgarh', 'bilaspur': 'Chhattisgarh',
  'kochi': 'Kerala', 'cochin': 'Kerala', 'ernakulam': 'Kerala', 'thiruvananthapuram': 'Kerala',
  'trivandrum': 'Kerala', 'kozhikode': 'Kerala', 'calicut': 'Kerala', 'thrissur': 'Kerala',
  'guwahati': 'Assam', 'dibrugarh': 'Assam', 'silchar': 'Assam',
  'chandigarh': 'Punjab', 'shimla': 'Himachal Pradesh', 'dharamshala': 'Himachal Pradesh',
  'baddi': 'Himachal Pradesh', 'solan': 'Himachal Pradesh',
  'dehradun': 'Uttarakhand', 'haridwar': 'Uttarakhand', 'rudrapur': 'Uttarakhand',
  'srinagar': 'Jammu and Kashmir', 'jammu': 'Jammu and Kashmir',
  'panaji': 'Goa', 'margao': 'Goa', 'imphal': 'Manipur', 'shillong': 'Meghalaya',
  'aizawl': 'Mizoram', 'kohima': 'Nagaland', 'agartala': 'Tripura', 'gangtok': 'Sikkim',
  'itanagar': 'Arunachal Pradesh', 'port blair': 'Andaman and Nicobar Islands',
  'daman': 'Dadra and Nagar Haveli and Daman and Diu', 'silvassa': 'Dadra and Nagar Haveli and Daman and Diu',
  'leh': 'Ladakh'
};

/** Turns whatever the row says into an area name the map knows. */
function mapAreaOf(raw, geo) {
  const s = String(raw || '').trim();
  if (!s || s === '(blank)') return null;
  if (geo[s]) return s;
  const low = s.toLowerCase();
  if (geo === GEO_INDIA) {
    if (CITY_STATE[low]) return CITY_STATE[low];
    // "NEW DELHI - 110007" and similar
    const head = low.split(/[,\-(]/)[0].trim();
    if (CITY_STATE[head]) return CITY_STATE[head];
  }
  const hit = Object.keys(geo).find(k => k.toLowerCase() === low);
  return hit || null;
}

/** Choropleth. Areas with no rows stay blank rather than reading as zero. */
function svgMap(geo, labels, values, colours, pick, W, H) {
  const byArea = new Map();
  labels.forEach((lb, i) => {
    const a = mapAreaOf(lb, geo);
    if (!a) return;
    const e = byArea.get(a);
    byArea.set(a, { v: (e ? e.v : 0) + values[i], from: e ? e.from : lb });
  });
  if (!byArea.size) {
    return scEmpty('None of these values match a place on the map. ' +
      'Group the chart by City, State or Country.');
  }
  // longitude / latitude straight onto the box, which is fine at this size
  let minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
  Object.keys(geo).forEach(k => geo[k].forEach(r => {
    for (let i = 0; i < r.length; i += 2) {
      if (r[i] < minX) minX = r[i]; if (r[i] > maxX) maxX = r[i];
      if (r[i + 1] < minY) minY = r[i + 1]; if (r[i + 1] > maxY) maxY = r[i + 1];
    }
  }));
  const pad = 6;
  const sc = Math.min((W - pad * 2) / (maxX - minX), (H - pad * 2) / (maxY - minY));
  const ox = pad + ((W - pad * 2) - (maxX - minX) * sc) / 2;
  const oy = pad + ((H - pad * 2) - (maxY - minY) * sc) / 2;
  const px = x => (ox + (x - minX) * sc).toFixed(1);
  const py = y => (oy + (maxY - y) * sc).toFixed(1);

  let max = 0;
  byArea.forEach(e => { if (e.v > max) max = e.v; });
  const base = colours[0];
  const blank = fadeColour(boardTextColour(), .08);

  let out = '';
  Object.keys(geo).forEach(name => {
    const e = byArea.get(name);
    const t = e && max ? e.v / max : 0;
    const fill = e ? fadeColour(base, .18 + t * .82) : blank;
    const d = geo[name].map(r => {
      let s = '';
      for (let i = 0; i < r.length; i += 2) s += (i ? 'L' : 'M') + px(r[i]) + ',' + py(r[i + 1]);
      return s + 'Z';
    }).join('');
    out += '<g class="sc-cell"' + (e ? ' data-v="' + escapeHtml(e.from) + '"' : '') + '>' +
      '<title>' + escapeHtml(name + (e ? ': ' + fmtNum(e.v) : ' \u2014 no rows')) + '</title>' +
      '<path d="' + d + '" fill="' + fill + '" stroke="' + boardGridColour() +
      '" stroke-width=".5"' + (pick && e && e.from !== pick ? ' class="sc-dim"' : '') + '/></g>';
  });
  return svgOpen(W, H) + out + '</svg>';
}

/* =========================================================================
   The dispatcher
   ========================================================================= */
function drawBoardSvgChart(boardId, cfg) {
  const host = document.getElementById('bs-' + boardId + '-' + cfg.id);
  if (!host) return;
  const W = Math.max(160, host.clientWidth || (cfg.w || 520) - 22);
  const H = Math.max(110, host.clientHeight || (cfg.h || 300) - 66);
  const colours = chartColours(boardId, cfg);
  const pick = boardPick(boardId, cfg);
  const s = boardSeries(boardId, cfg);
  let html = '';

  try {
    switch (cfg.type) {
      case 'treemap':  html = svgTreemap(s.labels, s.values, colours, pick, W, H); break;
      case 'sunburst': html = svgSunburst(boardId, cfg, s.labels, s.values, colours, pick, W, H); break;
      case 'funnel':   html = svgFunnel(s.labels, s.values, colours, pick, W, H); break;
      case 'waterfall':html = svgWaterfall(s.labels, s.values, colours, pick, W, H); break;
      case 'wordcloud':html = svgWordCloud(s.labels, s.values, colours, pick, W, H); break;
      case 'heatmap':  html = svgHeatmap(boardMatrix(boardId, cfg), colours, pick, W, H); break;
      case 'sankey':   html = svgSankey(boardMatrix(boardId, cfg), colours, pick, W, H); break;
      case 'gantt':    html = svgGantt(boardId, cfg, colours, pick, W, H); break;
      case 'candlestick': html = svgCandlestick(boardId, cfg, colours, pick, W, H); break;
      case 'map-india':html = svgMap(GEO_INDIA, s.labels, s.values, colours, pick, W, H); break;
      case 'map-world':html = svgMap(GEO_WORLD, s.labels, s.values, colours, pick, W, H); break;
      case 'bullet': {
        // what is on the shelf against what the row should hold
        const days = catalogDays ? catalogDays() : 30;
        const targets = s.labels.map((lb, i) => Math.max(1, (s.values[i] / (days || 1)) *
          (CatPrefs.defaultLT || 15) * (CatPrefs.defaultSF || 1.5)));
        html = svgBullet(s.labels, s.values, targets, colours, pick, W, H); break;
      }
      case 'boxplot':
      case 'violin': {
        const groups = s.labels.map(lb => {
          const sub = Object.assign({}, cfg, { dim: cfg.split || 'Item Code' });
          const f = { dim: cfg.dim, value: lb };
          Boards[boardId].filters.push(f);
          const v = boardValuesFor(boardId, sub);
          Boards[boardId].filters.pop();
          return v;
        }).filter(a => a.length);
        const keep = s.labels.filter((lb, i) => (i < groups.length));
        html = groups.length
          ? svgBoxplot(groups, keep, colours, pick, W, H, cfg.type === 'violin')
          : scEmpty('Not enough detail behind these rows to show a spread.');
        break;
      }
      default: html = scEmpty('That chart type is not drawn yet.');
    }
  } catch (e) {
    html = scEmpty('This chart could not be drawn from the current data.');
  }

  host.innerHTML = html;
  host.querySelectorAll('.sc-cell[data-v]').forEach(el =>
    el.addEventListener('click', () => addBoardFilter(boardId, cfg.dim, el.dataset.v)));
}

/* ---------------------------------------------------------------
   16e. MAP GEOMETRY
   ---------------------------------------------------------------
   Outlines for the choropleth maps. Both sets are heavily
   simplified - coordinates are rounded to two decimals (about a
   kilometre), which is far more detail than a dashboard tile needs
   and keeps the whole thing under 100 KB.

   India: the 33 states and union territories, built by dissolving
   the 2011 district boundaries.  Source: udit-001/india-maps-data.
   World: 173 countries.  Source: johan/world.geo.json.

   Shape is { "Name": [ [lon,lat,lon,lat,...], ... ] } - one flat
   number list per ring, which is a good deal smaller than nested
   pairs and just as quick to draw.
   --------------------------------------------------------------- */

const GEO_INDIA = {"Mizoram":[[92.6,22.01,92.51,22.74,92.37,22.94,92.39,23.06,92.35,23.24,92.4,23.27,92.26,23.81,92.32,23.91,92.3,24.25,92.42,24.25,92.44,24.15,92.53,24.18,92.63,24.33,92.69,24.35,92.76,24.51,92.8,24.42,93.01,24.41,93.02,24.23,92.98,24.12,93.21,24.05,93.4,23.93,93.42,23.53,93.36,23.33,93.39,23.22,93.36,23.07,93.3,23.0,93.2,23.06,93.12,23.01,93.15,22.93,93.09,22.71,93.13,22.47,93.19,22.42,93.16,22.18,93.05,22.2,93.0,22.05,92.94,22.02,92.71,22.15,92.69,22.04,92.6,22.01]],"Tamil Nadu":[[77.26,8.51,77.27,8.58,77.17,8.74,77.26,8.88,77.15,9.01,77.27,9.15,77.28,9.3,77.4,9.53,77.32,9.6,77.27,9.57,77.17,9.61,77.25,9.8,77.21,9.88,77.29,10.23,77.18,10.36,76.98,10.22,76.89,10.26,76.82,10.42,76.91,10.78,76.83,10.87,76.66,10.94,76.75,11.08,76.69,11.17,76.71,11.24,76.63,11.19,76.44,11.2,76.55,11.36,76.23,11.53,76.51,11.7,76.56,11.62,76.85,11.59,76.91,11.8,77.29,11.81,77.43,11.77,77.49,11.94,77.68,11.95,77.78,12.11,77.75,12.16,77.51,12.2,77.47,12.24,77.62,12.37,77.6,12.67,77.74,12.67,77.83,12.87,77.93,12.89,77.98,12.84,78.06,12.85,78.41,12.63,78.59,12.78,78.62,12.99,78.89,13.1,79.16,13.02,79.23,13.16,79.28,13.12,79.44,13.19,79.41,13.33,79.54,13.34,79.53,13.28,79.68,13.3,79.7,13.22,79.79,13.24,79.81,13.32,79.95,13.36,80.04,13.54,80.08,13.49,80.21,13.49,80.27,13.56,80.34,13.28,80.14,12.41,79.86,12.05,79.84,11.96,79.69,11.88,79.81,11.84,79.76,11.62,79.86,11.18,79.85,10.99,79.75,11.0,79.72,10.93,79.85,10.83,79.86,10.28,79.39,10.32,79.27,10.23,79.23,10.14,79.25,10.05,78.98,9.69,78.9,9.47,79.06,9.3,79.34,9.31,79.3,9.25,79.2,9.28,78.87,9.26,78.38,9.09,78.17,8.88,78.19,8.75,78.14,8.68,78.14,8.5,78.06,8.38,77.84,8.25,77.79,8.16,77.59,8.14,77.56,8.08,77.23,8.18,77.1,8.29,77.26,8.51]],"Madhya Pradesh":[[76.49,21.19,76.38,21.08,76.17,21.08,76.16,21.25,76.1,21.37,75.21,21.41,75.11,21.46,75.05,21.57,74.58,21.66,74.5,21.73,74.52,21.91,74.43,22.03,74.28,21.93,74.15,21.95,74.09,22.02,74.12,22.21,74.05,22.3,74.07,22.36,74.19,22.32,74.22,22.44,74.12,22.42,74.06,22.55,74.15,22.52,74.26,22.64,74.38,22.64,74.47,22.82,74.46,22.91,74.39,22.9,74.32,23.06,74.39,23.11,74.51,23.08,74.73,23.22,74.53,23.32,74.57,23.42,74.85,23.55,74.92,23.67,74.9,23.87,74.97,24.04,74.88,24.17,74.91,24.21,74.75,24.24,74.85,24.42,74.7,24.48,74.79,24.59,74.8,24.74,74.89,24.63,74.99,24.69,74.98,24.76,74.86,24.8,74.83,24.9,74.9,24.93,74.96,24.86,75.04,24.85,75.12,24.88,75.11,24.97,75.19,25.04,75.35,25.02,75.18,24.75,75.61,24.68,75.78,24.76,75.92,24.54,75.9,24.44,75.79,24.45,75.73,24.4,75.81,24.25,75.75,24.14,75.84,24.06,75.7,23.97,75.52,24.03,75.46,23.92,75.69,23.75,75.79,23.87,75.94,23.9,76.04,24.07,76.13,24.09,76.12,24.19,76.32,24.25,76.57,24.21,76.64,24.27,76.68,24.19,76.8,24.12,76.88,24.13,76.92,24.21,76.83,24.36,76.82,24.53,76.86,24.55,76.98,24.46,77.07,24.56,77.03,24.71,76.95,24.76,76.86,24.74,76.8,24.81,76.95,24.87,76.86,24.96,76.88,25.04,77.09,25.05,77.17,25.11,77.31,25.08,77.39,25.12,77.41,25.22,77.37,25.41,77.31,25.43,77.22,25.32,77.08,25.34,76.95,25.28,76.67,25.35,76.52,25.53,76.55,25.84,76.79,25.94,76.9,26.09,77.27,26.27,77.44,26.4,77.82,26.55,77.9,26.66,78.09,26.68,78.11,26.8,78.36,26.87,78.54,26.76,78.78,26.77,79.01,26.67,78.99,26.58,79.14,26.44,79.14,26.34,78.99,26.2,79.01,26.08,78.76,25.72,78.81,25.62,78.66,25.56,78.45,25.57,78.4,25.44,78.34,25.43,78.45,25.12,78.34,25.08,78.35,25.01,78.17,24.88,78.28,24.66,78.27,24.45,78.37,24.37,78.33,24.33,78.39,24.26,78.5,24.39,78.81,24.18,78.97,24.35,78.99,24.44,78.91,24.46,78.95,24.55,78.88,24.64,78.75,24.6,78.77,24.86,78.63,24.96,78.64,25.07,78.51,25.28,78.66,25.38,78.77,25.35,78.73,25.46,78.88,25.52,78.94,25.49,78.88,25.39,78.96,25.34,78.88,25.16,79.06,25.22,79.13,25.11,79.25,25.11,79.32,25.15,79.28,25.32,79.35,25.33,79.44,25.25,79.41,25.11,79.49,25.08,79.57,25.17,79.84,25.1,79.86,25.19,80.03,25.34,80.32,25.39,80.31,25.29,80.4,25.24,80.43,25.17,80.26,25.02,80.34,25.01,80.4,25.07,80.48,24.98,80.5,25.04,80.59,25.09,80.71,25.07,80.72,25.12,80.84,25.19,80.86,24.93,80.95,24.97,81.12,24.9,81.26,25.05,81.28,25.17,81.39,25.13,81.54,25.19,81.63,25.16,81.67,25.06,81.89,24.99,81.94,24.85,82.22,24.79,82.3,24.61,82.42,24.59,82.42,24.65,82.67,24.7,82.71,24.64,82.77,24.64,82.8,24.58,82.71,24.38,82.77,24.37,82.74,24.17,82.66,24.13,82.81,23.96,82.69,23.92,82.64,23.84,82.5,23.78,82.21,23.85,82.08,23.82,81.92,23.87,81.81,23.81,81.67,23.92,81.6,23.88,81.7,23.72,81.61,23.66,81.61,23.51,81.92,23.53,81.99,23.41,82.1,23.4,82.2,23.31,82.12,23.1,81.94,23.08,81.95,22.98,81.86,22.89,81.78,22.87,81.79,22.77,81.62,22.54,81.47,22.49,81.33,22.53,81.25,22.46,81.19,22.49,81.11,22.44,81.12,22.29,81.05,22.24,81.0,22.06,80.91,22.11,80.82,21.76,80.76,21.76,80.71,21.66,80.74,21.47,80.66,21.33,80.42,21.44,80.27,21.62,80.12,21.61,80.07,21.55,79.92,21.52,79.74,21.6,79.54,21.54,79.5,21.67,79.23,21.72,79.22,21.65,78.92,21.59,78.9,21.5,78.58,21.49,78.51,21.53,78.43,21.5,78.42,21.57,78.34,21.59,77.88,21.38,77.49,21.37,77.45,21.55,77.58,21.53,77.54,21.7,77.48,21.77,77.28,21.76,77.21,21.69,77.06,21.71,76.9,21.6,76.86,21.61,76.78,21.56,76.78,21.46,76.62,21.33,76.66,21.28,76.62,21.18,76.49,21.19]],"Maharashtra":[[74.32,16.55,74.32,16.33,74.37,16.26,74.48,16.25,74.48,16.15,74.37,16.08,74.46,16.04,74.35,15.76,74.24,15.74,74.21,15.78,74.12,15.72,74.12,15.65,73.99,15.61,73.94,15.74,73.69,15.73,73.59,15.91,73.51,15.94,73.37,16.36,73.28,16.89,73.29,17.12,73.22,17.27,73.24,17.32,73.18,17.38,73.18,17.58,72.97,18.12,72.97,18.24,73.03,18.26,72.91,18.35,72.89,18.46,72.93,18.54,72.8,18.91,72.84,19.05,72.79,19.13,72.8,19.3,72.65,19.84,72.66,19.93,72.73,19.97,72.71,20.07,72.87,20.23,72.96,20.21,72.96,20.13,73.14,20.09,73.25,20.12,73.3,20.21,73.42,20.2,73.4,20.39,73.47,20.58,73.41,20.62,73.47,20.72,73.65,20.56,73.74,20.56,73.84,20.62,73.87,20.73,73.92,20.72,73.94,20.84,73.89,20.93,73.67,21.15,73.82,21.17,73.82,21.27,73.94,21.3,73.96,21.4,74.05,21.45,74.32,21.5,74.28,21.56,74.16,21.57,73.84,21.52,73.81,21.64,73.88,21.71,73.82,21.84,74.15,21.95,74.28,21.93,74.43,22.03,74.52,21.91,74.5,21.73,74.58,21.66,75.05,21.57,75.11,21.46,75.21,21.41,76.1,21.37,76.16,21.25,76.17,21.08,76.38,21.08,76.49,21.19,76.62,21.18,76.66,21.28,76.62,21.33,76.78,21.46,76.78,21.56,76.86,21.61,76.9,21.6,77.0,21.69,77.21,21.69,77.28,21.76,77.48,21.77,77.54,21.7,77.58,21.53,77.45,21.55,77.49,21.37,77.88,21.38,78.34,21.59,78.42,21.57,78.43,21.5,78.51,21.53,78.58,21.49,78.9,21.5,78.92,21.59,79.22,21.65,79.23,21.72,79.5,21.67,79.54,21.54,79.74,21.6,79.92,21.52,80.07,21.55,80.12,21.61,80.27,21.62,80.42,21.44,80.66,21.33,80.61,21.23,80.46,21.18,80.44,21.1,80.46,20.93,80.55,20.93,80.58,20.67,80.52,20.58,80.62,20.59,80.59,20.39,80.63,20.33,80.4,20.24,80.43,20.14,80.55,20.1,80.47,19.83,80.55,19.82,80.71,19.62,80.9,19.52,80.85,19.36,80.75,19.29,80.61,19.31,80.58,19.4,80.33,19.14,80.27,18.94,80.35,18.85,80.28,18.72,80.12,18.68,79.91,18.83,79.96,18.87,79.94,19.03,79.86,19.1,79.93,19.16,79.98,19.4,79.81,19.58,79.74,19.61,79.47,19.5,79.25,19.59,79.19,19.46,79.1,19.53,78.94,19.55,78.96,19.66,78.86,19.69,78.84,19.76,78.43,19.81,78.31,19.91,78.36,19.78,78.34,19.71,78.28,19.69,78.31,19.48,78.21,19.43,78.17,19.24,77.94,19.34,77.86,19.3,77.86,19.18,77.8,19.07,77.84,18.91,77.95,18.84,77.84,18.81,77.79,18.68,77.73,18.67,77.74,18.56,77.6,18.55,77.52,18.37,77.55,18.29,77.47,18.26,77.41,18.3,77.41,18.39,77.32,18.44,77.06,18.15,76.95,18.19,76.92,17.92,76.74,17.9,76.79,17.83,76.71,17.78,76.69,17.68,76.52,17.76,76.49,17.66,76.33,17.6,76.38,17.31,76.23,17.37,75.93,17.32,75.89,17.4,75.81,17.37,75.63,17.48,75.57,17.38,75.65,17.27,75.63,17.17,75.67,17.12,75.67,16.98,75.28,16.96,75.26,16.86,75.18,16.84,75.09,16.95,74.9,16.86,74.92,16.77,74.7,16.72,74.63,16.58,74.57,16.55,74.54,16.64,74.32,16.55]],"Chhattisgarh":[[81.66,18.35,81.53,18.24,81.48,17.98,81.4,17.89,81.4,17.82,81.17,17.86,81.02,17.79,80.98,18.17,80.87,18.14,80.85,18.21,80.76,18.17,80.71,18.44,80.64,18.52,80.45,18.63,80.37,18.61,80.28,18.72,80.35,18.85,80.27,18.94,80.33,19.14,80.58,19.4,80.61,19.31,80.75,19.29,80.85,19.36,80.9,19.52,80.71,19.62,80.55,19.82,80.47,19.83,80.55,20.1,80.43,20.14,80.4,20.24,80.63,20.33,80.59,20.39,80.62,20.59,80.52,20.58,80.58,20.67,80.55,20.93,80.46,20.93,80.44,21.1,80.46,21.18,80.61,21.23,80.74,21.47,80.71,21.66,80.76,21.76,80.82,21.76,80.91,22.11,81.0,22.06,81.05,22.24,81.12,22.29,81.11,22.44,81.19,22.49,81.25,22.46,81.33,22.53,81.47,22.49,81.62,22.54,81.79,22.77,81.78,22.87,81.86,22.89,81.95,22.98,81.94,23.08,82.12,23.1,82.2,23.31,82.1,23.4,81.99,23.41,81.92,23.53,81.61,23.51,81.61,23.66,81.7,23.72,81.6,23.88,81.67,23.92,81.81,23.81,81.92,23.87,82.08,23.82,82.21,23.85,82.5,23.78,82.81,23.96,82.95,23.87,83.09,23.87,83.21,23.92,83.33,24.1,83.43,24.09,83.57,23.86,83.71,23.82,83.77,23.61,83.94,23.57,84.02,23.63,83.96,23.39,84.07,23.33,84.06,23.11,84.18,22.98,84.37,22.98,84.4,22.93,84.24,22.69,84.01,22.57,84.05,22.47,84.01,22.37,83.86,22.34,83.63,22.2,83.54,22.05,83.59,21.83,83.46,21.7,83.49,21.64,83.38,21.62,83.34,21.5,83.39,21.34,83.27,21.38,83.19,21.14,83.08,21.11,82.97,21.18,82.64,21.15,82.64,21.09,82.46,20.83,82.34,20.87,82.38,20.63,82.33,20.56,82.44,20.43,82.4,20.06,82.72,19.99,82.7,19.83,82.59,19.78,82.53,19.88,82.45,19.91,82.34,19.84,82.27,19.98,81.95,20.11,81.88,20.05,81.87,19.9,82.07,19.79,82.03,19.51,82.1,19.52,82.13,19.43,82.19,19.43,82.16,19.2,82.25,18.92,82.18,18.9,82.16,18.79,82.09,18.74,81.9,18.66,81.91,18.57,81.75,18.35,81.66,18.35]],"Gujarat":[[70.83,20.69,70.45,20.85,70.09,21.11,69.37,21.83,69.4,21.89,69.35,21.84,69.22,21.96,68.94,22.31,69.01,22.44,69.16,22.35,69.18,22.21,69.34,22.31,69.86,22.46,69.98,22.54,70.13,22.55,70.22,22.66,70.22,22.72,70.4,22.9,70.45,22.84,70.72,23.19,70.31,23.22,70.08,22.92,69.65,22.8,69.13,22.87,68.63,23.17,68.53,23.28,68.55,23.36,68.4,23.42,68.42,23.49,68.33,23.5,68.31,23.58,68.14,23.6,68.1,23.68,68.17,23.7,68.21,23.88,68.34,23.97,68.75,23.97,68.76,24.3,68.85,24.24,68.95,24.3,69.0,24.22,69.09,24.27,69.19,24.24,69.59,24.29,69.73,24.17,70.03,24.17,70.11,24.29,70.57,24.42,70.57,24.25,70.81,24.22,70.99,24.37,71.05,24.36,71.1,24.44,71.0,24.44,70.99,24.6,71.11,24.68,71.29,24.62,71.48,24.68,71.66,24.64,72.06,24.71,72.17,24.62,72.35,24.63,72.35,24.56,72.5,24.42,72.54,24.52,72.68,24.46,72.73,24.37,72.96,24.36,73.0,24.48,73.09,24.5,73.08,24.4,73.15,24.35,73.07,24.19,73.24,24.0,73.32,24.05,73.41,24.04,73.37,23.81,73.52,23.61,73.57,23.65,73.66,23.62,73.63,23.45,73.76,23.45,73.89,23.33,73.96,23.38,74.14,23.27,74.13,23.18,74.25,23.18,74.39,22.9,74.46,22.91,74.47,22.82,74.38,22.64,74.26,22.64,74.15,22.52,74.06,22.55,74.12,22.42,74.22,22.44,74.19,22.32,74.07,22.36,74.05,22.3,74.12,22.21,74.09,22.02,74.15,21.95,73.82,21.84,73.88,21.71,73.81,21.64,73.84,21.52,74.16,21.57,74.28,21.56,74.32,21.5,74.05,21.45,73.96,21.4,73.94,21.3,73.82,21.27,73.82,21.17,73.67,21.15,73.89,20.93,73.94,20.84,73.92,20.72,73.87,20.73,73.84,20.62,73.74,20.56,73.65,20.56,73.47,20.72,73.41,20.62,73.47,20.58,73.4,20.39,73.42,20.2,73.3,20.21,73.21,20.12,73.13,20.2,73.17,20.29,73.04,20.33,72.93,20.3,72.96,20.21,72.87,20.23,72.78,20.13,72.74,20.24,72.89,20.46,72.9,20.58,72.72,21.05,72.75,21.06,72.62,21.1,72.6,21.32,72.74,21.47,72.61,21.58,72.7,21.65,72.54,21.73,72.62,21.81,72.51,21.94,72.57,22.18,72.65,22.22,72.75,22.17,72.95,22.29,72.66,22.28,72.52,22.32,72.46,22.26,72.38,22.31,72.2,21.89,72.24,21.82,72.19,21.75,72.31,21.62,72.11,21.3,72.11,21.2,71.51,20.94,71.44,20.87,70.83,20.69]],"Odisha":[[82.35,18.17,82.32,18.04,82.25,17.99,82.03,18.07,81.62,17.82,81.4,17.82,81.4,17.89,81.48,17.98,81.53,18.24,81.66,18.35,81.75,18.35,81.91,18.57,81.9,18.66,82.09,18.74,82.16,18.79,82.18,18.9,82.25,18.92,82.16,19.2,82.19,19.43,82.13,19.43,82.1,19.52,82.03,19.51,82.07,19.79,81.87,19.9,81.88,20.05,81.95,20.11,82.27,19.98,82.34,19.84,82.45,19.91,82.53,19.88,82.59,19.78,82.7,19.83,82.72,19.99,82.4,20.06,82.44,20.43,82.33,20.56,82.38,20.63,82.34,20.87,82.46,20.83,82.64,21.09,82.64,21.15,82.97,21.18,83.08,21.11,83.19,21.14,83.27,21.38,83.39,21.34,83.34,21.5,83.38,21.62,83.49,21.64,83.46,21.7,83.59,21.83,83.54,22.05,83.63,22.2,83.86,22.34,84.01,22.37,84.05,22.47,84.01,22.53,84.29,22.34,84.43,22.35,84.53,22.43,84.75,22.42,85.07,22.49,85.08,22.26,85.03,22.12,85.24,22.01,85.4,22.16,85.68,22.06,85.75,22.07,85.78,21.98,85.9,21.98,86.03,22.19,85.97,22.25,86.03,22.33,85.96,22.49,86.07,22.56,86.12,22.49,86.28,22.45,86.36,22.35,86.43,22.32,86.5,22.35,86.73,22.22,86.84,22.1,86.96,22.09,87.03,22.04,87.0,21.91,87.1,21.86,87.11,21.92,87.24,21.96,87.28,21.8,87.44,21.77,87.48,21.62,87.41,21.54,87.15,21.51,86.95,21.36,86.86,21.21,86.85,21.09,86.98,20.86,86.97,20.79,86.87,20.78,86.81,20.74,86.86,20.78,87.02,20.77,87.06,20.71,86.79,20.53,86.74,20.46,86.78,20.34,86.53,20.18,86.37,19.98,85.35,19.59,84.98,19.32,84.77,19.08,84.7,19.12,84.59,19.02,84.51,19.05,84.32,18.79,84.09,18.76,84.04,18.82,83.89,18.82,83.79,19.02,83.75,18.99,83.76,18.91,83.66,19.1,83.48,19.03,83.46,18.96,83.35,19.02,83.32,18.97,83.41,18.87,83.2,18.75,83.14,18.79,83.0,18.61,83.04,18.55,83.02,18.46,83.08,18.4,82.97,18.36,82.81,18.45,82.77,18.34,82.59,18.28,82.61,18.37,82.49,18.53,82.36,18.42,82.38,18.37,82.31,18.19,82.35,18.17]],"Andhra Pradesh":[[79.16,13.02,78.89,13.1,78.75,13.05,78.62,12.99,78.59,12.78,78.41,12.63,78.3,12.67,78.23,12.76,78.25,12.86,78.47,12.97,78.59,13.28,78.36,13.36,78.4,13.59,78.2,13.57,78.2,13.63,78.12,13.72,78.12,13.86,78.0,13.87,77.97,13.96,77.81,13.92,77.78,13.82,77.46,13.68,77.43,13.84,77.18,13.87,77.06,13.74,76.98,13.83,77.04,13.93,76.93,14.03,76.95,14.1,76.89,14.17,77.0,14.2,77.03,14.06,77.17,14.01,77.32,14.03,77.4,13.91,77.43,13.95,77.39,14.02,77.39,14.17,77.5,14.16,77.48,14.29,77.41,14.34,77.36,14.28,77.14,14.34,77.12,14.23,77.01,14.24,76.88,14.35,76.9,14.49,76.76,14.6,76.87,14.97,76.76,14.99,76.8,15.1,76.99,15.01,77.11,15.03,77.16,15.14,77.15,15.29,77.06,15.33,77.03,15.44,77.03,15.63,77.08,15.66,77.03,15.84,77.07,15.91,77.25,15.96,77.66,15.89,77.89,15.91,78.08,15.84,78.17,15.86,78.26,16.01,78.41,16.07,78.55,16.04,78.6,16.08,78.74,16.01,78.92,16.12,78.9,16.18,79.03,16.24,79.21,16.23,79.22,16.51,79.26,16.57,79.42,16.58,79.79,16.73,79.95,16.64,80.06,16.74,80.06,16.97,80.19,17.02,80.31,16.97,80.32,16.88,80.57,16.77,80.6,16.92,80.36,16.98,80.58,17.15,80.84,17.04,80.94,17.22,80.99,17.19,81.17,17.23,81.19,17.32,81.27,17.32,81.31,17.41,81.02,17.52,80.91,17.64,80.89,17.74,80.97,17.78,81.06,17.73,81.17,17.86,81.27,17.81,81.62,17.82,82.03,18.07,82.25,17.99,82.32,18.04,82.36,18.42,82.49,18.53,82.61,18.37,82.59,18.28,82.77,18.34,82.81,18.45,82.97,18.36,83.08,18.4,83.02,18.46,83.04,18.55,83.0,18.61,83.14,18.79,83.2,18.75,83.41,18.87,83.32,18.97,83.35,19.02,83.46,18.96,83.48,19.03,83.66,19.1,83.76,18.91,83.75,18.99,83.79,19.02,83.89,18.82,84.04,18.82,84.09,18.76,84.32,18.79,84.51,19.05,84.59,19.02,84.7,19.12,84.76,19.07,84.13,18.31,83.66,18.08,83.16,17.56,82.54,17.25,82.32,17.06,82.25,16.94,82.25,16.88,82.36,16.86,82.31,16.56,81.71,16.31,81.39,16.34,81.27,16.28,81.15,15.97,81.01,15.85,81.01,15.75,80.94,15.72,80.83,15.71,80.77,15.87,80.68,15.89,80.32,15.72,80.12,15.37,80.04,15.01,80.2,14.59,80.13,14.23,80.15,14.04,80.25,13.78,80.27,13.56,80.21,13.49,80.08,13.49,80.04,13.54,79.95,13.36,79.81,13.32,79.79,13.24,79.7,13.22,79.68,13.3,79.53,13.28,79.54,13.34,79.41,13.33,79.44,13.19,79.28,13.12,79.23,13.16,79.16,13.02]],"Karnataka":[[75.94,11.94,75.58,12.16,75.49,12.29,75.41,12.29,75.42,12.37,75.37,12.41,75.42,12.5,75.28,12.52,75.11,12.68,74.87,12.76,74.82,12.85,74.58,13.93,74.52,13.99,74.26,14.72,74.12,14.79,74.09,14.9,74.2,14.93,74.3,15.04,74.27,15.1,74.32,15.19,74.26,15.24,74.33,15.28,74.24,15.67,74.12,15.65,74.12,15.72,74.21,15.78,74.24,15.74,74.35,15.76,74.46,16.04,74.37,16.08,74.48,16.15,74.48,16.25,74.37,16.26,74.32,16.33,74.32,16.55,74.54,16.64,74.57,16.55,74.63,16.58,74.7,16.72,74.92,16.77,74.9,16.86,75.09,16.95,75.18,16.84,75.26,16.86,75.28,16.96,75.67,16.98,75.67,17.12,75.63,17.17,75.65,17.27,75.57,17.38,75.63,17.48,75.81,17.37,75.89,17.4,75.93,17.32,76.23,17.37,76.38,17.31,76.33,17.6,76.49,17.66,76.52,17.76,76.69,17.68,76.71,17.78,76.79,17.83,76.74,17.9,76.92,17.92,76.95,18.19,77.06,18.15,77.32,18.44,77.41,18.39,77.41,18.3,77.47,18.26,77.6,18.28,77.56,18.2,77.6,18.09,77.55,18.06,77.66,17.97,77.45,17.69,77.44,17.58,77.69,17.51,77.51,17.43,77.38,17.22,77.36,17.16,77.5,17.04,77.48,16.79,77.43,16.72,77.45,16.62,77.42,16.52,77.24,16.47,77.29,16.41,77.59,16.34,77.5,16.22,77.51,15.93,77.25,15.96,77.07,15.91,77.03,15.84,77.08,15.66,77.03,15.63,77.03,15.44,77.06,15.33,77.15,15.29,77.16,15.14,77.11,15.03,76.99,15.01,76.8,15.1,76.76,14.99,76.87,14.97,76.76,14.6,76.9,14.49,76.88,14.35,77.01,14.24,77.12,14.23,77.14,14.34,77.36,14.28,77.41,14.34,77.48,14.29,77.5,14.16,77.39,14.17,77.39,14.02,77.43,13.95,77.4,13.91,77.32,14.03,77.17,14.01,77.03,14.06,77.0,14.2,76.89,14.17,76.95,14.1,76.93,14.03,77.04,13.93,76.98,13.83,77.06,13.74,77.18,13.87,77.43,13.84,77.46,13.68,77.78,13.82,77.81,13.92,77.97,13.96,78.0,13.87,78.12,13.86,78.12,13.72,78.2,13.63,78.2,13.57,78.4,13.59,78.36,13.36,78.59,13.28,78.47,12.97,78.25,12.86,78.23,12.76,78.06,12.85,77.98,12.84,77.93,12.89,77.83,12.87,77.74,12.67,77.6,12.67,77.62,12.37,77.47,12.24,77.78,12.11,77.68,11.95,77.49,11.94,77.43,11.77,77.29,11.81,77.11,11.77,77.01,11.81,76.91,11.8,76.85,11.59,76.56,11.62,76.51,11.7,76.42,11.67,76.41,11.76,76.32,11.75,76.21,11.86,76.11,11.86,76.11,11.98,75.94,11.94]],"Goa":[[74.28,15.39,74.32,15.37,74.33,15.28,74.26,15.24,74.32,15.19,74.27,15.1,74.3,15.04,74.2,14.93,74.09,14.9,73.98,15.05,73.88,15.35,73.93,15.43,73.77,15.49,73.69,15.73,73.94,15.74,73.99,15.61,74.24,15.67,74.28,15.39]],"Kerala":[[76.9,10.25,76.98,10.22,77.18,10.36,77.29,10.23,77.21,9.88,77.25,9.8,77.17,9.61,77.27,9.57,77.32,9.6,77.4,9.53,77.28,9.3,77.27,9.15,77.15,9.01,77.26,8.88,77.17,8.74,77.27,8.58,77.26,8.51,77.1,8.29,76.55,8.9,76.32,9.48,76.17,10.18,76.21,10.19,76.16,10.18,75.7,11.43,75.62,11.48,75.55,11.7,75.25,12.01,75.2,12.01,74.87,12.76,75.11,12.68,75.28,12.52,75.42,12.5,75.37,12.41,75.42,12.37,75.41,12.29,75.49,12.29,75.58,12.16,75.79,12.06,75.87,11.95,76.0,11.93,76.11,11.98,76.11,11.86,76.21,11.86,76.32,11.75,76.41,11.76,76.38,11.6,76.23,11.53,76.55,11.36,76.44,11.2,76.63,11.19,76.71,11.24,76.69,11.17,76.75,11.08,76.66,10.94,76.83,10.87,76.91,10.78,76.82,10.42,76.9,10.25]],"Telangana":[[77.29,16.41,77.24,16.47,77.42,16.52,77.45,16.62,77.43,16.72,77.48,16.79,77.5,17.04,77.36,17.16,77.38,17.22,77.51,17.43,77.69,17.51,77.44,17.58,77.45,17.69,77.66,17.97,77.55,18.06,77.6,18.09,77.56,18.2,77.6,18.28,77.52,18.37,77.6,18.55,77.74,18.56,77.73,18.67,77.79,18.68,77.84,18.81,77.95,18.84,77.84,18.91,77.8,19.07,77.86,19.18,77.86,19.3,77.94,19.34,78.17,19.24,78.21,19.43,78.31,19.48,78.28,19.69,78.34,19.71,78.36,19.78,78.31,19.91,78.43,19.81,78.84,19.76,78.86,19.69,78.96,19.66,78.94,19.55,79.1,19.53,79.19,19.46,79.25,19.59,79.47,19.5,79.74,19.61,79.94,19.47,79.98,19.4,79.93,19.16,79.86,19.1,79.94,19.03,79.96,18.87,79.91,18.83,80.12,18.68,80.28,18.72,80.37,18.61,80.45,18.63,80.64,18.52,80.71,18.44,80.76,18.17,80.85,18.21,80.87,18.14,80.98,18.17,81.02,17.79,81.08,17.8,81.06,17.73,80.97,17.78,80.89,17.74,80.91,17.64,81.02,17.52,81.31,17.41,81.27,17.32,81.19,17.32,81.17,17.23,80.99,17.19,80.94,17.22,80.84,17.04,80.58,17.15,80.36,16.98,80.6,16.92,80.57,16.77,80.32,16.88,80.31,16.97,80.19,17.02,80.06,16.97,80.06,16.74,79.95,16.64,79.79,16.73,79.42,16.58,79.26,16.57,79.22,16.51,79.21,16.23,79.03,16.24,78.9,16.18,78.92,16.12,78.74,16.01,78.6,16.08,78.55,16.04,78.41,16.07,78.26,16.01,78.17,15.86,77.51,15.93,77.5,16.22,77.59,16.34,77.29,16.41]],"West Bengal":[[88.19,22.1,87.79,21.7,87.48,21.62,87.44,21.77,87.28,21.8,87.24,21.96,87.11,21.92,87.1,21.86,87.0,21.91,87.03,22.04,86.84,22.1,86.73,22.22,86.85,22.4,86.76,22.43,86.8,22.49,86.77,22.58,86.66,22.58,86.64,22.66,86.42,22.79,86.44,22.93,86.5,22.99,86.24,23.0,86.04,23.15,85.89,23.17,85.82,23.27,85.88,23.35,85.89,23.48,86.04,23.5,86.08,23.57,86.16,23.56,86.18,23.47,86.31,23.41,86.37,23.53,86.49,23.64,86.8,23.7,86.8,23.83,86.92,23.88,87.09,23.81,87.23,23.86,87.29,23.9,87.23,24.03,87.34,24.03,87.44,23.98,87.63,24.16,87.64,24.25,87.71,24.26,87.8,24.38,87.79,24.58,87.89,24.57,87.91,24.72,87.84,24.75,87.96,24.92,87.95,24.97,87.79,25.09,87.78,25.25,87.71,25.26,87.76,25.27,87.82,25.24,87.85,25.3,87.77,25.42,87.87,25.5,87.96,25.54,88.07,25.48,88.02,25.6,88.05,25.7,87.97,25.73,87.82,25.94,87.85,26.04,88.28,26.33,88.24,26.46,88.11,26.55,88.19,26.77,88.12,26.95,87.99,27.11,88.04,27.22,88.15,27.11,88.28,27.13,88.42,27.07,88.6,27.19,88.87,27.11,88.87,26.97,88.92,26.99,89.09,26.89,89.13,26.81,89.37,26.85,89.45,26.8,89.55,26.81,89.76,26.7,89.85,26.7,89.87,26.61,89.84,26.38,89.7,26.22,89.73,26.17,89.63,26.23,89.61,26.15,89.65,26.06,89.54,26.01,89.35,26.01,89.15,26.14,89.1,26.39,88.98,26.45,88.91,26.38,88.99,26.35,88.93,26.28,88.76,26.33,88.42,26.56,88.4,26.64,88.33,26.5,88.48,26.46,88.52,26.36,88.4,26.33,88.35,26.22,88.18,26.15,88.09,25.93,88.14,25.79,88.27,25.78,88.45,25.66,88.45,25.6,88.65,25.48,88.81,25.52,88.84,25.36,88.99,25.27,88.91,25.18,88.44,25.2,88.45,25.08,88.39,24.97,88.28,24.89,88.23,24.96,88.14,24.91,88.17,24.86,88.03,24.69,88.13,24.51,88.56,24.31,88.68,24.32,88.74,24.25,88.75,24.15,88.7,24.1,88.74,24.03,88.73,23.91,88.61,23.86,88.56,23.71,88.59,23.62,88.76,23.45,88.71,23.23,88.81,23.26,88.94,23.21,88.86,22.97,88.97,22.84,88.91,22.76,88.96,22.69,88.93,22.65,89.0,22.43,89.0,22.29,89.07,22.21,89.07,21.95,89.01,21.9,88.93,21.94,88.85,21.92,88.76,21.95,88.76,22.02,88.53,21.85,88.44,21.92,88.39,21.9,88.27,21.73,88.3,21.67,88.26,21.56,88.15,21.95,88.22,22.1,88.08,22.26,88.0,22.24,88.02,22.2,88.14,22.19,88.19,22.1]],"Dadra and Nagar Haveli and Daman and Diu":[[73.04,20.33,73.17,20.29,73.13,20.2,73.21,20.12,73.14,20.09,72.96,20.13,72.93,20.3,73.04,20.33]],"Arunachal Pradesh":[[92.67,27.03,92.56,26.96,92.11,26.89,92.1,27.01,92.03,27.06,92.07,27.28,92.01,27.49,91.65,27.48,91.56,27.63,91.63,27.71,91.63,27.8,91.83,27.81,91.88,27.72,92.26,27.86,92.31,27.78,92.56,27.82,92.73,27.98,92.66,28.07,92.69,28.13,93.02,28.3,93.22,28.34,93.18,28.43,93.28,28.5,93.34,28.64,93.96,28.71,94.0,28.84,94.16,28.95,94.29,29.15,94.43,29.23,94.56,29.23,94.68,29.33,94.87,29.18,95.25,29.11,95.27,29.05,95.45,29.03,95.54,29.21,96.09,29.46,96.24,29.23,96.39,29.25,96.35,29.17,96.22,29.14,96.15,29.05,96.25,28.94,96.35,29.03,96.44,29.03,96.61,28.79,96.58,28.72,96.47,28.68,96.49,28.63,96.33,28.51,96.37,28.48,96.31,28.38,96.66,28.46,96.78,28.36,96.9,28.38,96.99,28.31,97.09,28.36,97.37,28.22,97.33,28.06,97.39,28.0,97.38,27.91,97.25,27.9,96.9,27.61,96.86,27.63,96.93,27.51,96.9,27.45,97.16,27.14,97.06,27.09,96.86,27.19,96.88,27.26,96.8,27.35,96.61,27.37,96.51,27.3,96.31,27.29,96.09,27.22,95.95,27.05,95.8,27.02,95.75,26.9,95.68,26.9,95.61,26.81,95.5,26.82,95.43,26.7,95.3,26.65,95.23,26.68,95.2,27.03,95.45,27.13,95.48,27.24,95.53,27.26,95.63,27.23,95.86,27.27,96.02,27.37,95.88,27.44,95.87,27.54,95.78,27.62,95.77,27.73,95.97,27.96,95.82,27.97,95.61,27.96,95.38,27.84,95.31,27.87,95.14,27.82,94.46,27.56,94.24,27.63,94.21,27.61,94.27,27.52,93.99,27.33,93.81,27.15,93.83,27.07,93.67,26.97,93.02,26.91,92.89,27.0,92.67,27.03]],"Assam":[[92.53,24.18,92.44,24.15,92.42,24.25,92.22,24.29,92.26,24.36,92.21,24.51,92.3,24.74,92.23,24.9,92.3,24.92,92.39,24.85,92.52,24.88,92.43,25.03,92.48,25.11,92.62,25.12,92.79,25.28,92.57,25.47,92.66,25.57,92.58,25.55,92.44,25.72,92.23,25.72,92.15,25.67,92.15,25.82,92.2,25.91,92.16,25.97,92.27,26.07,91.87,26.04,91.83,26.12,91.72,26.03,91.72,25.97,91.58,26.03,91.45,25.84,91.34,25.84,91.22,25.72,91.15,25.85,91.01,25.82,91.03,25.89,90.95,25.95,90.63,25.92,90.43,25.99,90.12,25.96,89.9,25.73,90.02,25.61,89.88,25.55,89.85,25.47,89.88,25.62,89.83,25.73,89.89,25.95,89.7,26.22,89.84,26.38,89.86,26.73,90.19,26.77,90.23,26.86,90.42,26.91,90.72,26.77,91.05,26.78,91.1,26.82,91.34,26.78,91.41,26.84,91.5,26.79,91.7,26.81,91.89,26.92,92.06,26.85,92.11,26.89,92.56,26.96,92.67,27.03,92.89,27.0,93.02,26.91,93.67,26.97,93.83,27.07,93.81,27.15,93.99,27.33,94.27,27.52,94.21,27.61,94.24,27.63,94.46,27.56,95.14,27.82,95.51,27.88,95.61,27.96,95.97,27.96,95.77,27.73,95.78,27.62,95.87,27.54,95.88,27.44,96.02,27.37,95.86,27.27,95.48,27.24,95.45,27.13,94.98,26.91,94.92,26.95,94.74,26.75,94.59,26.73,94.4,26.61,94.4,26.53,94.31,26.46,94.28,26.56,94.0,26.17,93.98,25.92,93.88,25.84,93.78,25.85,93.78,25.97,93.68,25.9,93.7,25.85,93.33,25.54,93.46,25.43,93.47,25.31,93.38,25.24,93.17,24.8,93.11,24.81,93.09,24.57,93.01,24.41,92.8,24.42,92.76,24.51,92.69,24.35,92.63,24.33,92.53,24.18]],"Nagaland":[[93.33,25.54,93.7,25.85,93.68,25.9,93.78,25.97,93.78,25.85,93.88,25.84,93.98,25.92,94.0,26.17,94.28,26.56,94.31,26.46,94.4,26.53,94.4,26.61,94.59,26.73,94.74,26.75,94.92,26.95,94.98,26.91,95.2,27.03,95.23,26.68,95.06,26.45,95.12,26.35,95.12,26.11,95.19,26.07,95.01,25.89,95.04,25.74,94.81,25.49,94.67,25.45,94.55,25.5,94.55,25.58,94.49,25.6,94.32,25.5,94.01,25.57,93.76,25.54,93.8,25.46,93.61,25.2,93.5,25.24,93.46,25.43,93.33,25.54]],"Meghalaya":[[91.35,25.23,90.45,25.14,89.97,25.3,89.83,25.3,89.82,25.39,89.88,25.55,90.02,25.61,89.9,25.73,90.12,25.96,90.43,25.99,90.63,25.92,90.95,25.95,91.03,25.89,91.01,25.82,91.15,25.85,91.22,25.72,91.34,25.84,91.45,25.84,91.58,26.03,91.72,25.97,91.72,26.03,91.83,26.12,91.87,26.04,92.27,26.07,92.16,25.97,92.2,25.91,92.15,25.82,92.15,25.67,92.23,25.72,92.44,25.72,92.58,25.55,92.66,25.57,92.57,25.47,92.79,25.28,92.62,25.12,92.48,25.11,92.43,25.03,92.08,25.18,91.75,25.17,91.64,25.12,91.6,25.17,91.49,25.13,91.38,25.18,91.35,25.23]],"Manipur":[[93.5,23.94,93.38,24.09,93.25,24.02,92.98,24.12,93.11,24.81,93.17,24.8,93.38,25.24,93.47,25.31,93.61,25.2,93.8,25.46,93.76,25.54,94.01,25.57,94.11,25.52,94.19,25.55,94.21,25.5,94.42,25.54,94.49,25.6,94.55,25.58,94.55,25.5,94.67,25.45,94.58,25.21,94.65,25.12,94.74,25.11,94.74,25.04,94.63,24.76,94.38,24.49,94.41,24.44,94.16,23.85,94.02,23.93,93.81,23.93,93.74,24.01,93.62,24.01,93.5,23.94]],"Tripura":[[91.76,23.27,91.81,23.06,91.72,22.99,91.58,22.97,91.45,23.26,91.32,23.28,91.26,23.49,91.16,23.59,91.16,23.74,91.23,23.73,91.22,23.88,91.3,24.0,91.63,24.1,91.66,24.19,91.76,24.14,91.75,24.25,91.92,24.11,91.92,24.34,92.14,24.43,92.16,24.53,92.21,24.51,92.26,24.36,92.22,24.29,92.33,24.18,92.32,23.91,92.26,23.81,92.26,23.73,92.12,23.73,92.06,23.64,91.94,23.69,91.97,23.48,91.94,23.43,91.8,23.52,91.84,23.37,91.76,23.27]],"Andaman and Nicobar Islands":[[92.43,10.55,92.38,10.78,92.52,10.9,92.6,10.68,92.53,10.52,92.43,10.55],[92.71,11.97,92.79,11.89,92.75,11.7,92.66,11.65,92.75,11.61,92.67,11.51,92.55,11.72,92.54,11.91,92.6,11.85,92.66,12.21,92.76,12.18,92.73,12.08,92.78,12.05,92.71,11.97],[93.8,6.77,93.65,7.13,93.67,7.19,93.84,7.24,93.91,6.95,93.88,6.79,93.8,6.77],[92.72,12.3,92.73,12.84,92.79,12.86,92.84,13.39,92.99,13.57,93.04,13.53,93.03,13.08,92.96,13.02,92.91,13.05,92.86,12.9,92.92,12.88,92.97,12.54,92.87,12.33,92.72,12.3]],"Uttar Pradesh":[[78.75,24.6,78.88,24.64,78.95,24.55,78.91,24.46,78.99,24.42,78.81,24.18,78.5,24.39,78.39,24.26,78.33,24.33,78.37,24.37,78.27,24.45,78.28,24.66,78.17,24.88,78.35,25.01,78.34,25.08,78.45,25.12,78.34,25.43,78.4,25.44,78.45,25.57,78.66,25.56,78.81,25.62,78.76,25.72,79.01,26.08,78.99,26.2,79.14,26.34,79.14,26.44,78.99,26.58,79.01,26.67,78.78,26.77,78.54,26.76,78.36,26.87,78.22,26.83,78.27,26.92,78.15,26.95,78.11,26.9,77.99,26.89,77.76,26.93,77.45,26.74,77.4,26.83,77.66,26.97,77.55,27.11,77.66,27.17,77.65,27.24,77.34,27.53,77.35,27.65,77.28,27.8,77.53,27.93,77.54,27.99,77.48,28.08,77.51,28.37,77.3,28.57,77.34,28.61,77.21,28.79,77.23,28.92,77.14,29.17,77.11,29.63,77.15,29.69,77.12,29.76,77.23,29.94,77.58,30.31,77.58,30.4,77.66,30.4,77.72,30.33,77.96,30.24,77.8,30.1,77.72,29.89,77.82,29.67,77.93,29.71,78.02,29.58,78.34,29.8,78.51,29.72,78.62,29.56,78.93,29.44,78.76,29.26,78.88,29.21,78.93,29.12,79.04,29.17,79.15,29.12,79.14,29.07,79.26,28.99,79.41,28.95,79.42,28.9,79.58,28.85,79.82,28.89,79.95,28.72,80.03,28.75,80.06,28.84,80.12,28.83,80.28,28.71,80.31,28.63,80.46,28.62,80.55,28.69,80.67,28.64,80.71,28.57,81.21,28.36,81.32,28.2,81.43,28.17,81.48,28.08,81.88,27.86,81.96,27.92,82.07,27.92,82.45,27.68,82.71,27.7,82.75,27.59,82.73,27.49,83.19,27.45,83.29,27.34,83.41,27.41,83.4,27.48,83.61,27.47,83.83,27.37,83.99,27.17,83.96,27.08,84.05,27.04,84.06,26.89,84.26,26.8,84.33,26.68,84.42,26.63,84.28,26.6,84.11,26.63,84.05,26.55,83.92,26.5,83.91,26.45,84.16,26.38,84.18,26.26,84.09,26.22,84.06,26.1,84.62,25.79,84.6,25.73,84.56,25.68,84.34,25.74,84.32,25.67,84.21,25.67,84.14,25.73,83.85,25.43,83.35,25.18,83.4,24.78,83.49,24.74,83.52,24.55,83.4,24.5,83.41,24.27,83.33,24.1,83.21,23.92,83.09,23.87,82.95,23.87,82.66,24.13,82.74,24.17,82.77,24.37,82.71,24.38,82.8,24.58,82.77,24.64,82.71,24.64,82.67,24.7,82.42,24.65,82.42,24.59,82.3,24.61,82.22,24.79,81.94,24.85,81.89,24.99,81.67,25.06,81.63,25.16,81.54,25.19,81.39,25.13,81.28,25.17,81.26,25.05,81.12,24.9,80.95,24.97,80.86,24.93,80.84,25.19,80.72,25.12,80.71,25.07,80.59,25.09,80.5,25.04,80.48,24.98,80.4,25.07,80.34,25.01,80.26,25.02,80.43,25.17,80.4,25.24,80.31,25.29,80.32,25.39,80.03,25.34,79.86,25.19,79.84,25.1,79.57,25.17,79.49,25.08,79.41,25.11,79.44,25.25,79.35,25.33,79.28,25.32,79.32,25.15,79.13,25.11,79.06,25.22,78.88,25.16,78.96,25.34,78.88,25.39,78.94,25.49,78.88,25.52,78.73,25.46,78.77,25.35,78.66,25.38,78.51,25.28,78.64,25.07,78.63,24.96,78.77,24.86,78.75,24.6]],"Rajasthan":[[74.51,23.08,74.39,23.11,74.32,23.06,74.25,23.18,74.13,23.18,74.14,23.27,73.96,23.38,73.89,23.33,73.83,23.43,73.63,23.45,73.66,23.62,73.57,23.65,73.52,23.61,73.37,23.81,73.41,24.04,73.32,24.05,73.24,24.0,73.07,24.19,73.15,24.35,73.08,24.4,73.09,24.5,73.0,24.48,72.96,24.36,72.73,24.37,72.68,24.46,72.52,24.5,72.5,24.42,72.33,24.63,72.17,24.62,72.06,24.71,71.66,24.64,71.48,24.68,71.29,24.62,71.11,24.68,70.94,24.94,70.89,25.15,70.67,25.4,70.67,25.68,70.61,25.71,70.39,25.67,70.27,25.71,70.1,25.94,70.08,26.08,70.17,26.25,70.17,26.55,70.06,26.6,69.79,26.6,69.51,26.74,69.48,26.8,69.59,27.18,70.03,27.56,70.13,27.81,70.37,28.01,70.59,28.01,70.68,27.92,70.76,27.72,70.91,27.71,71.2,27.84,71.67,27.88,71.87,27.96,71.88,27.91,71.93,28.12,72.21,28.39,72.3,28.67,72.4,28.78,72.95,29.03,73.28,29.57,73.4,29.95,73.81,30.07,73.97,30.19,73.98,30.12,73.89,29.97,74.53,29.94,74.56,29.87,74.48,29.81,74.48,29.75,74.62,29.75,74.58,29.56,74.62,29.53,74.6,29.36,74.79,29.36,74.85,29.4,74.94,29.36,74.96,29.28,75.05,29.29,75.08,29.23,75.41,29.26,75.39,29.06,75.53,28.97,75.47,28.93,75.56,28.61,75.89,28.39,76.06,28.22,75.97,27.93,75.93,27.93,75.97,27.85,76.18,27.8,76.23,27.83,76.16,28.0,76.26,28.07,76.32,28.02,76.36,28.12,76.31,28.18,76.54,28.03,76.54,27.97,76.67,28.01,76.69,28.09,76.85,28.22,76.97,28.14,76.89,27.73,76.98,27.65,77.04,27.82,77.28,27.8,77.35,27.65,77.34,27.53,77.65,27.24,77.66,27.17,77.55,27.11,77.66,26.97,77.4,26.83,77.45,26.74,77.76,26.93,77.99,26.89,78.11,26.9,78.15,26.95,78.27,26.92,78.22,26.83,78.11,26.8,78.09,26.68,77.9,26.66,77.82,26.55,77.44,26.4,77.27,26.27,76.9,26.09,76.79,25.94,76.55,25.84,76.52,25.53,76.67,25.35,76.95,25.28,77.08,25.34,77.22,25.32,77.31,25.43,77.37,25.41,77.41,25.22,77.39,25.12,77.31,25.08,77.17,25.11,77.09,25.05,76.88,25.04,76.86,24.96,76.95,24.87,76.8,24.81,76.86,24.74,76.95,24.76,77.03,24.71,77.07,24.56,76.98,24.46,76.86,24.55,76.82,24.53,76.83,24.36,76.92,24.21,76.88,24.13,76.8,24.12,76.68,24.19,76.64,24.27,76.57,24.21,76.32,24.25,76.12,24.19,76.13,24.09,76.04,24.07,75.94,23.9,75.79,23.87,75.69,23.75,75.46,23.92,75.52,24.03,75.7,23.97,75.84,24.06,75.75,24.14,75.81,24.25,75.73,24.4,75.79,24.45,75.9,24.44,75.92,24.54,75.78,24.76,75.61,24.68,75.18,24.75,75.35,25.02,75.19,25.04,75.11,24.97,75.12,24.88,75.04,24.85,74.96,24.86,74.9,24.93,74.83,24.9,74.86,24.8,74.98,24.76,74.99,24.69,74.89,24.63,74.8,24.74,74.79,24.59,74.7,24.48,74.85,24.42,74.75,24.24,74.91,24.21,74.88,24.17,74.97,24.04,74.9,23.87,74.94,23.77,74.85,23.55,74.57,23.42,74.53,23.32,74.73,23.22,74.51,23.08]],"Delhi":[[77.21,28.79,77.34,28.61,77.3,28.57,77.35,28.52,77.17,28.41,77.1,28.51,77.01,28.54,76.88,28.5,76.85,28.55,76.97,28.7,76.97,28.81,77.22,28.85,77.21,28.79]],"Haryana":[[76.06,28.22,75.89,28.39,75.56,28.61,75.47,28.93,75.53,28.97,75.39,29.06,75.41,29.26,75.08,29.23,75.05,29.29,74.96,29.28,74.94,29.36,74.85,29.4,74.79,29.36,74.6,29.36,74.62,29.53,74.58,29.56,74.62,29.75,74.48,29.75,74.48,29.81,74.56,29.87,74.53,29.94,74.65,29.91,74.71,29.97,74.81,29.99,75.0,29.87,75.19,29.84,75.24,29.75,75.17,29.67,75.3,29.56,75.35,29.7,75.46,29.81,75.58,29.74,75.78,29.82,75.95,29.73,76.05,29.75,76.1,29.81,76.21,29.84,76.17,29.93,76.25,30.09,76.41,30.15,76.47,30.1,76.58,30.1,76.64,30.15,76.56,30.24,76.71,30.33,76.75,30.43,76.91,30.43,76.93,30.61,76.83,30.68,76.85,30.83,76.78,30.9,76.92,30.9,77.02,30.75,77.17,30.68,77.17,30.6,77.12,30.56,77.21,30.49,77.41,30.43,77.52,30.44,77.58,30.38,77.58,30.31,77.23,29.94,77.12,29.76,77.14,29.17,77.23,28.92,77.22,28.85,76.94,28.8,76.97,28.7,76.85,28.55,76.88,28.5,77.01,28.54,77.1,28.51,77.17,28.41,77.35,28.52,77.51,28.37,77.48,28.08,77.54,27.99,77.53,27.93,77.28,27.8,77.04,27.82,76.98,27.65,76.89,27.73,76.97,28.14,76.85,28.22,76.69,28.09,76.67,28.01,76.54,27.97,76.54,28.03,76.31,28.18,76.36,28.12,76.32,28.02,76.26,28.07,76.16,28.0,76.23,27.83,76.18,27.8,75.97,27.85,75.93,27.93,75.97,27.93,76.06,28.22]],"Sikkim":[[88.92,27.32,88.72,27.14,88.53,27.18,88.42,27.07,88.28,27.13,88.1,27.13,88.04,27.22,88.04,27.48,88.2,27.79,88.2,27.94,88.62,28.1,88.76,28.08,88.84,27.99,88.89,27.86,88.85,27.67,88.76,27.57,88.79,27.42,88.92,27.32]],"Bihar":[[85.03,24.42,84.9,24.37,84.89,24.46,84.8,24.53,84.59,24.4,84.5,24.29,84.34,24.39,84.33,24.49,84.09,24.53,84.01,24.64,83.87,24.53,83.5,24.53,83.53,24.64,83.49,24.74,83.4,24.78,83.35,25.18,83.85,25.43,84.14,25.73,84.21,25.67,84.32,25.67,84.34,25.74,84.56,25.68,84.62,25.79,84.06,26.1,84.09,26.22,84.18,26.26,84.16,26.38,83.91,26.45,83.92,26.5,84.05,26.55,84.11,26.63,84.28,26.6,84.42,26.63,84.33,26.68,84.26,26.8,84.06,26.89,84.05,27.04,83.96,27.08,83.99,27.17,83.83,27.43,84.02,27.44,84.1,27.52,84.26,27.45,84.28,27.39,84.62,27.34,84.69,27.22,84.64,27.05,84.96,26.96,85.03,26.87,85.19,26.87,85.18,26.8,85.25,26.75,85.33,26.74,85.64,26.85,85.73,26.81,85.72,26.67,85.86,26.57,86.03,26.67,86.54,26.54,86.72,26.43,86.85,26.44,87.04,26.56,87.16,26.4,87.46,26.44,87.59,26.38,87.85,26.44,88.03,26.39,88.1,26.46,88.11,26.55,88.24,26.46,88.28,26.33,87.85,26.04,87.82,25.94,87.97,25.73,88.05,25.7,88.02,25.6,88.07,25.48,87.96,25.54,87.77,25.42,87.85,25.3,87.82,25.24,87.55,25.33,87.48,25.19,87.33,25.22,87.29,25.09,87.15,25.02,87.16,24.88,87.11,24.85,87.05,24.63,86.95,24.63,86.91,24.55,86.75,24.62,86.65,24.57,86.61,24.61,86.5,24.51,86.45,24.37,86.28,24.47,86.3,24.56,86.13,24.6,86.13,24.7,86.05,24.78,85.93,24.74,85.74,24.82,85.68,24.59,85.54,24.53,85.29,24.52,85.09,24.37,85.03,24.42]],"Jharkhand":[[86.03,22.19,85.9,21.98,85.78,21.98,85.75,22.07,85.68,22.06,85.4,22.16,85.24,22.01,85.03,22.12,85.08,22.26,85.07,22.49,84.75,22.42,84.53,22.43,84.43,22.35,84.29,22.34,84.01,22.53,84.08,22.64,84.24,22.69,84.4,22.93,84.37,22.98,84.18,22.98,84.06,23.11,84.07,23.33,83.96,23.39,84.02,23.63,83.94,23.57,83.77,23.61,83.71,23.82,83.57,23.86,83.43,24.09,83.33,24.1,83.41,24.27,83.4,24.5,83.87,24.53,84.01,24.64,84.09,24.53,84.33,24.49,84.34,24.39,84.5,24.29,84.59,24.4,84.8,24.53,84.89,24.46,84.9,24.37,85.03,24.42,85.09,24.37,85.37,24.55,85.54,24.53,85.68,24.59,85.74,24.82,85.93,24.74,86.05,24.78,86.13,24.7,86.13,24.6,86.3,24.56,86.28,24.47,86.45,24.37,86.5,24.51,86.61,24.61,86.65,24.57,86.75,24.62,86.91,24.55,86.95,24.63,87.05,24.63,87.11,24.85,87.16,24.88,87.15,25.02,87.29,25.09,87.33,25.22,87.48,25.19,87.55,25.33,87.69,25.31,87.71,25.26,87.78,25.25,87.79,25.09,87.95,24.97,87.96,24.92,87.84,24.75,87.91,24.72,87.89,24.57,87.79,24.58,87.8,24.38,87.71,24.26,87.64,24.25,87.63,24.16,87.44,23.98,87.34,24.03,87.23,24.03,87.29,23.9,87.23,23.86,87.09,23.81,86.92,23.88,86.8,23.83,86.8,23.7,86.49,23.64,86.37,23.53,86.31,23.41,86.18,23.47,86.16,23.56,85.89,23.48,85.88,23.35,85.82,23.27,85.89,23.17,86.04,23.15,86.24,23.0,86.5,22.99,86.44,22.93,86.42,22.79,86.64,22.66,86.66,22.58,86.77,22.58,86.8,22.49,86.76,22.43,86.85,22.4,86.73,22.22,86.5,22.35,86.43,22.32,86.36,22.35,86.28,22.45,86.12,22.49,86.07,22.56,85.96,22.49,86.03,22.33,85.97,22.25,86.03,22.19]],"Ladakh":[[77.15,32.98,76.93,33.03,76.5,33.52,76.31,33.57,76.26,33.74,76.17,33.8,76.02,34.0,75.86,33.98,75.69,34.16,75.41,34.28,75.32,34.43,75.36,34.53,75.24,34.63,74.73,34.67,74.56,34.77,74.32,34.79,74.57,34.87,74.59,35.01,74.49,35.11,74.25,35.04,73.8,35.24,73.73,35.22,73.7,35.36,73.79,35.46,73.78,35.52,73.63,35.57,73.41,35.56,73.14,35.72,73.15,35.84,73.09,35.88,72.69,35.83,72.53,35.92,72.59,36.03,72.53,36.07,72.59,36.26,72.78,36.31,72.88,36.39,72.88,36.44,72.99,36.46,73.13,36.69,73.44,36.74,73.7,36.68,73.81,36.72,73.86,36.79,73.72,36.83,73.69,36.91,73.88,36.89,74.06,36.81,74.14,36.84,74.17,36.9,74.26,36.89,74.71,37.08,74.73,37.02,74.84,37.02,74.92,36.91,75.02,36.98,75.1,36.94,75.21,37.04,75.34,37.05,75.47,36.81,75.61,36.76,75.75,36.59,76.08,36.47,76.25,36.35,76.35,36.36,76.48,36.21,76.64,36.18,76.8,36.05,76.72,35.95,76.81,35.84,76.93,35.78,77.12,35.8,77.17,35.73,77.43,35.65,77.46,35.57,77.43,35.52,77.61,35.47,77.82,35.51,77.9,35.47,77.97,35.57,78.14,35.54,78.2,35.67,78.42,35.78,78.52,35.76,78.66,35.86,78.96,35.89,79.13,35.84,79.21,35.98,79.38,36.0,79.42,35.9,79.87,35.78,80.0,35.84,80.11,35.68,80.29,35.61,80.33,35.47,80.18,35.17,80.17,35.0,80.07,34.83,80.07,34.71,79.86,34.69,79.77,34.63,79.79,34.48,79.51,34.47,79.59,34.28,79.49,34.13,79.35,34.04,79.02,34.05,78.9,33.98,79.11,33.62,78.91,33.57,78.99,33.33,79.45,33.26,79.38,33.08,79.42,32.96,79.63,32.74,79.45,32.53,79.32,32.59,79.14,32.48,79.11,32.38,78.98,32.34,78.86,32.44,78.76,32.65,78.66,32.65,78.4,32.53,78.31,32.57,78.41,32.62,78.39,32.76,78.02,32.62,77.92,32.69,77.93,32.73,77.74,32.97,77.47,32.86,77.41,32.88,77.34,32.82,77.15,32.98]],"Jammu and Kashmir":[[74.87,32.49,74.68,32.49,74.64,32.61,74.66,32.79,74.38,32.76,73.96,32.99,73.63,33.09,73.66,33.21,73.56,33.34,73.62,33.57,73.56,33.62,73.59,33.88,73.52,33.99,73.49,34.22,73.39,34.37,73.45,34.57,73.62,34.59,73.75,34.79,73.96,34.84,74.07,34.94,74.07,35.05,74.15,35.1,74.25,35.04,74.49,35.11,74.59,35.01,74.57,34.87,74.32,34.79,74.56,34.77,74.73,34.67,75.24,34.63,75.36,34.53,75.32,34.43,75.41,34.28,75.69,34.16,75.86,33.98,76.02,34.0,76.17,33.8,76.26,33.74,76.31,33.57,76.5,33.52,76.77,33.25,76.75,33.18,76.41,33.19,75.95,32.87,75.85,32.93,75.8,32.89,75.92,32.75,75.92,32.63,75.69,32.39,75.51,32.31,75.11,32.46,74.95,32.45,74.89,32.52,74.87,32.49]],"Himachal Pradesh":[[76.92,30.9,76.87,30.87,76.62,31.0,76.65,31.21,76.59,31.27,76.5,31.27,76.39,31.41,76.29,31.31,76.2,31.3,76.02,31.63,75.91,31.95,75.6,32.08,75.66,32.14,75.64,32.25,75.93,32.39,75.84,32.51,75.92,32.63,75.92,32.75,75.8,32.89,75.85,32.93,75.95,32.87,76.41,33.19,76.75,33.18,76.77,33.25,76.93,33.03,77.15,32.98,77.34,32.82,77.41,32.88,77.47,32.86,77.74,32.97,77.93,32.73,77.92,32.69,78.02,32.62,78.39,32.76,78.41,32.62,78.31,32.57,78.53,32.41,78.48,32.28,78.61,32.21,78.59,32.15,78.78,31.99,78.71,31.77,78.85,31.61,78.73,31.54,78.8,31.44,78.75,31.37,78.79,31.3,78.88,31.28,79.01,31.11,78.89,31.1,78.81,31.2,78.61,31.23,78.48,31.2,78.37,31.29,77.97,31.17,77.86,31.11,77.81,30.85,77.7,30.76,77.82,30.51,77.58,30.38,77.52,30.44,77.41,30.43,77.21,30.49,77.12,30.56,77.17,30.6,77.17,30.68,77.02,30.75,76.92,30.9]],"Punjab":[[75.46,29.81,75.35,29.7,75.3,29.56,75.17,29.67,75.24,29.75,75.19,29.84,75.0,29.87,74.81,29.99,74.71,29.97,74.65,29.91,73.89,29.97,73.98,30.12,73.96,30.27,73.88,30.36,74.37,30.89,74.71,31.07,74.69,31.13,74.51,31.13,74.55,31.36,74.65,31.46,74.5,31.74,74.56,31.76,74.57,31.84,74.91,32.04,75.17,32.07,75.38,32.18,75.34,32.35,75.51,32.31,75.84,32.51,75.93,32.39,75.64,32.25,75.66,32.14,75.6,32.08,75.91,31.95,76.02,31.63,76.2,31.3,76.29,31.31,76.39,31.41,76.5,31.27,76.59,31.27,76.65,31.21,76.62,31.0,76.85,30.83,76.84,30.75,76.71,30.76,76.74,30.69,76.93,30.61,76.91,30.43,76.75,30.43,76.71,30.33,76.56,30.24,76.64,30.15,76.58,30.1,76.47,30.1,76.41,30.15,76.25,30.09,76.17,29.93,76.21,29.84,76.1,29.81,76.05,29.75,75.95,29.73,75.78,29.82,75.58,29.74,75.46,29.81]],"Uttarakhand":[[78.62,29.56,78.51,29.72,78.34,29.8,78.02,29.58,77.93,29.71,77.82,29.67,77.72,29.89,77.8,30.1,77.96,30.24,77.72,30.33,77.66,30.4,77.58,30.4,77.82,30.51,77.7,30.76,77.81,30.85,77.86,31.11,78.31,31.28,78.48,31.2,78.61,31.23,78.81,31.2,78.89,31.1,79.01,31.11,78.88,31.28,79.07,31.46,79.14,31.43,79.32,31.14,79.6,30.93,79.86,30.97,80.1,30.79,80.17,30.81,80.25,30.74,80.21,30.58,80.54,30.45,80.61,30.47,81.05,30.21,80.93,30.17,80.68,29.96,80.6,29.96,80.55,29.85,80.37,29.74,80.41,29.6,80.24,29.44,80.32,29.3,80.26,29.19,80.14,29.1,80.03,28.75,79.95,28.72,79.82,28.89,79.58,28.85,79.26,28.99,79.14,29.07,79.15,29.12,79.04,29.17,78.93,29.12,78.88,29.21,78.76,29.26,78.93,29.44,78.62,29.56]]};

const GEO_WORLD = {"Afghanistan":[[61.21,35.65,62.98,35.4,65.75,37.66,69.2,37.15,70.81,38.49,71.84,36.74,73.26,37.5,75.16,37.13,71.85,36.51,71.26,36.07,71.61,35.15,70.88,33.99,69.93,34.02,70.32,33.36,69.26,32.5,69.32,31.9,66.94,31.3,66.35,29.89,64.15,29.34,60.87,29.83,61.78,30.74,60.54,32.98,61.21,35.65]],"Angola":[[16.33,-5.88,17.47,-8.07,19.02,-7.99,20.09,-6.94,21.73,-7.29,22.16,-11.08,23.91,-10.93,24.02,-12.91,21.93,-12.9,21.89,-16.08,23.22,-17.52,21.38,-17.93,18.26,-17.31,14.06,-17.42,13.46,-16.97,11.73,-17.3,12.18,-14.45,13.74,-11.3,12.23,-6.29,16.33,-5.88],[12.44,-5.68,11.91,-5.04,12.62,-4.44,13.0,-4.78,12.44,-5.68]],"Albania":[[20.59,41.86,21.0,40.58,20.15,39.62,19.41,40.25,19.3,42.2,19.74,42.69,20.59,41.86]],"United Arab Emirates":[[51.58,24.25,54.01,24.12,56.07,26.06,56.26,25.71,56.4,24.92,55.89,24.92,55.98,24.13,55.01,22.5,52.0,23.0,51.58,24.25]],"Argentina":[[-65.5,-55.2,-68.63,-54.87,-68.63,-52.64,-67.75,-53.85,-65.05,-54.7,-65.5,-55.2],[-64.96,-22.08,-64.38,-22.8,-63.99,-21.99,-62.85,-22.03,-60.85,-23.88,-57.78,-25.16,-58.62,-27.12,-55.7,-27.39,-54.13,-25.55,-53.65,-26.92,-57.63,-30.22,-58.5,-34.43,-57.23,-35.29,-56.79,-36.9,-57.75,-38.18,-59.23,-38.72,-62.34,-38.83,-62.15,-40.68,-62.75,-41.03,-65.12,-41.06,-64.98,-42.06,-63.76,-42.04,-63.46,-42.56,-65.18,-43.5,-65.57,-45.04,-67.29,-45.55,-67.58,-46.3,-65.64,-47.24,-65.99,-48.13,-69.14,-50.73,-68.15,-52.35,-71.91,-52.01,-72.31,-50.68,-73.33,-50.38,-73.42,-49.32,-72.33,-48.24,-71.66,-44.97,-71.22,-44.78,-72.15,-42.25,-71.41,-38.92,-70.81,-38.55,-71.12,-36.66,-70.36,-36.01,-69.82,-34.19,-70.54,-31.37,-69.66,-28.46,-68.3,-26.9,-68.42,-24.52,-67.33,-24.03,-67.11,-22.74,-66.27,-21.83,-64.96,-22.08]],"Armenia":[[43.58,41.09,45.56,40.81,46.51,38.77,43.66,40.25,43.58,41.09]],"Antarctica":[[-59.57,-80.04,-60.16,-81.0,-64.49,-80.92,-66.29,-80.26,-61.88,-80.39,-60.61,-79.63,-59.57,-80.04],[-159.21,-79.5,-161.13,-79.63,-163.71,-78.6,-161.25,-78.38,-159.21,-79.5],[-45.15,-78.05,-43.92,-78.48,-43.33,-80.03,-50.48,-81.03,-54.16,-80.63,-50.99,-79.61,-48.66,-78.05,-45.15,-78.05],[-121.21,-73.5,-118.72,-73.48,-120.23,-74.09,-122.62,-73.66,-121.21,-73.5],[-125.56,-73.48,-124.03,-73.87,-127.28,-73.46,-125.56,-73.48],[-98.98,-71.93,-96.79,-71.95,-96.2,-72.52,-100.78,-72.5,-102.33,-71.89,-98.98,-71.93],[-68.45,-70.96,-68.78,-72.17,-71.08,-72.5,-72.39,-72.48,-71.9,-72.09,-74.19,-72.37,-75.01,-71.66,-72.07,-71.19,-71.74,-69.51,-70.25,-68.88,-68.45,-70.96],[-58.61,-64.15,-62.02,-64.8,-62.65,-65.48,-62.12,-66.19,-63.75,-66.5,-65.67,-67.95,-63.2,-69.23,-61.81,-70.72,-60.83,-73.7,-64.35,-75.26,-70.6,-76.63,-77.24,-76.71,-73.66,-77.91,-77.93,-78.38,-78.02,-79.18,-75.36,-80.26,-59.69,-82.38,-58.22,-83.22,-49.76,-81.73,-42.81,-82.08,-40.77,-81.36,-28.55,-80.34,-29.69,-79.26,-35.64,-79.46,-35.78,-78.34,-28.88,-76.67,-17.52,-75.13,-15.7,-74.5,-15.41,-74.11,-16.47,-73.87,-15.45,-73.15,-12.29,-72.4,-10.3,-71.27,-7.42,-71.7,-6.87,-70.93,-4.34,-71.46,-0.66,-71.23,-0.23,-71.64,7.74,-69.89,9.53,-70.01,10.82,-70.83,13.42,-69.97,15.13,-70.4,19.26,-69.89,22.57,-70.7,27.09,-70.46,31.99,-69.66,33.87,-68.5,38.65,-69.78,54.53,-65.82,56.36,-65.97,58.74,-67.29,61.43,-67.95,64.05,-67.41,68.89,-67.93,69.67,-69.23,67.81,-70.31,69.07,-70.68,67.95,-71.85,69.87,-72.26,71.02,-72.09,73.86,-69.87,77.64,-69.46,79.11,-68.33,82.78,-67.21,86.75,-67.15,87.99,-66.21,89.67,-67.15,95.78,-67.39,99.72,-67.25,102.83,-65.56,106.18,-66.93,113.6,-65.88,115.6,-66.7,119.83,-67.27,123.22,-66.48,128.8,-66.76,134.76,-66.21,135.07,-65.31,137.46,-66.95,145.49,-66.92,146.65,-67.9,148.84,-68.39,152.5,-68.87,154.28,-68.56,161.57,-70.58,167.31,-70.83,171.21,-71.7,169.29,-73.66,166.09,-74.38,163.57,-76.24,163.49,-77.07,164.74,-78.18,167.0,-78.75,161.77,-79.16,159.79,-80.95,169.4,-83.83,180.0,-84.71,-179.94,-84.72,-179.06,-84.14,-174.38,-84.53,-169.95,-83.88,-158.07,-85.37,-148.53,-85.61,-143.11,-85.04,-142.89,-84.57,-153.59,-83.69,-152.86,-82.04,-156.84,-81.1,-150.65,-81.34,-146.42,-80.34,-149.53,-79.36,-155.33,-79.06,-158.05,-78.03,-158.37,-76.89,-151.33,-77.4,-146.1,-76.48,-146.2,-75.38,-144.91,-75.2,-144.32,-75.54,-135.21,-74.3,-119.7,-74.48,-113.94,-73.71,-112.3,-74.71,-111.26,-74.42,-107.56,-75.18,-100.65,-75.3,-100.12,-74.87,-102.55,-74.11,-103.68,-72.62,-96.34,-73.62,-90.09,-73.32,-89.23,-72.56,-85.19,-73.48,-81.47,-73.85,-80.3,-73.13,-76.22,-73.97,-68.94,-73.01,-67.13,-72.05,-68.54,-69.72,-67.58,-68.54,-67.74,-67.33,-67.25,-66.88,-63.0,-64.64,-58.59,-63.39,-57.22,-63.53,-58.61,-64.15]],"French Southern and Antarctic Lands":[[68.94,-48.62,70.56,-49.26,70.28,-49.71,68.75,-49.77,68.94,-48.62]],"Australia":[[145.4,-40.79,148.29,-40.88,147.91,-43.21,146.05,-43.55,144.72,-41.16,144.74,-40.7,145.4,-40.79],[143.56,-13.76,143.92,-14.55,144.56,-14.17,145.37,-14.98,146.39,-18.96,148.85,-20.39,149.68,-22.34,150.73,-22.4,150.9,-23.46,153.14,-26.07,153.57,-28.11,152.89,-31.64,150.33,-35.67,150.0,-37.43,146.32,-39.04,144.88,-38.42,145.03,-37.9,143.61,-38.81,140.64,-38.02,139.57,-36.14,138.12,-35.61,138.21,-34.38,136.83,-35.26,137.89,-33.64,137.81,-32.9,135.99,-34.89,135.21,-34.48,134.27,-32.62,131.33,-31.5,126.15,-32.22,124.22,-32.96,123.66,-33.89,119.89,-33.98,118.02,-35.06,116.63,-35.03,115.03,-34.2,115.71,-33.26,115.69,-31.61,113.34,-26.12,113.78,-26.55,113.44,-25.62,114.23,-26.3,113.39,-24.38,114.15,-21.76,114.23,-22.52,116.71,-20.7,120.86,-19.68,123.01,-16.41,123.43,-17.27,123.86,-17.07,123.5,-16.6,123.82,-16.11,124.26,-16.33,125.69,-14.23,127.07,-13.82,128.36,-14.87,129.62,-14.97,129.41,-14.42,130.62,-12.54,132.58,-12.11,131.82,-11.27,132.36,-11.13,135.3,-12.25,136.49,-11.86,136.95,-12.35,135.96,-13.32,135.5,-15.0,140.22,-17.71,141.27,-16.39,141.69,-12.41,142.52,-10.67,143.56,-13.76]],"Austria":[[16.98,48.12,16.01,46.68,14.63,46.43,12.15,47.12,11.05,46.75,9.48,47.1,9.9,47.58,12.93,47.47,12.88,48.29,13.6,48.88,16.5,48.79,16.98,48.12]],"Azerbaijan":[[47.37,41.22,48.58,41.81,50.39,40.26,49.57,40.18,48.88,38.32,48.01,38.79,48.06,39.58,46.51,38.77,44.97,41.25,46.5,41.06,46.4,41.86,47.37,41.22]],"Burundi":[[29.34,-4.5,29.02,-2.84,30.47,-2.41,30.75,-3.36,29.34,-4.5]],"Belgium":[[3.31,51.35,4.97,51.48,6.16,50.8,6.04,50.13,5.67,49.53,4.29,49.91,2.66,50.8,2.51,51.15,3.31,51.35]],"Benin":[[2.69,6.26,1.87,6.14,1.66,9.13,0.77,10.47,1.45,11.55,2.85,12.24,3.8,10.73,2.72,8.51,2.69,6.26]],"Burkina Faso":[[-2.83,9.64,-4.78,9.82,-5.4,10.37,-5.22,11.71,-4.01,13.47,-1.07,14.97,0.37,14.93,1.02,12.85,2.18,12.63,2.15,11.94,0.9,11.0,-2.94,10.96,-2.83,9.64]],"Bangladesh":[[92.67,22.04,92.37,20.67,91.42,22.77,90.5,22.81,90.27,21.84,89.03,22.06,88.7,24.23,88.08,24.5,88.93,25.24,88.21,25.77,88.56,26.45,89.83,25.97,89.92,25.27,92.38,24.98,91.16,23.5,91.71,22.99,92.15,23.63,92.67,22.04]],"Bulgaria":[[22.66,44.23,22.94,43.82,27.24,44.18,28.56,43.71,27.67,42.58,28.0,42.01,26.12,41.83,26.11,41.33,22.95,41.34,22.38,42.32,22.99,43.21,22.66,44.23]],"The Bahamas":[[-77.53,23.76,-78.41,24.58,-78.19,25.21,-77.53,23.76]],"Bosnia and Herzegovina":[[19.01,44.86,19.6,44.04,18.56,42.65,15.75,44.82,15.96,45.23,19.01,44.86]],"Belarus":[[23.48,53.91,25.54,54.28,26.49,55.62,28.18,56.17,30.87,55.55,30.76,54.81,32.69,53.35,31.31,53.07,31.79,52.1,30.93,52.04,30.56,51.32,25.33,51.91,23.53,51.58,23.48,53.91]],"Belize":[[-89.14,17.81,-88.11,18.35,-88.93,15.89,-89.14,17.81]],"Bolivia":[[-62.85,-22.03,-63.99,-21.99,-64.38,-22.8,-64.96,-22.08,-66.27,-21.83,-67.83,-22.87,-68.76,-20.37,-68.44,-19.41,-69.59,-17.58,-68.96,-16.5,-69.34,-14.95,-68.67,-12.56,-69.53,-10.95,-68.27,-11.01,-66.65,-9.93,-65.34,-9.76,-65.4,-11.57,-64.32,-12.46,-60.5,-13.78,-60.16,-16.26,-58.24,-16.3,-58.28,-17.27,-57.5,-18.17,-57.85,-19.97,-59.12,-19.36,-61.79,-19.63,-62.85,-22.03]],"Brazil":[[-57.63,-30.22,-53.65,-26.92,-53.63,-26.12,-54.13,-25.55,-54.63,-25.74,-54.29,-24.02,-55.4,-23.96,-55.8,-22.36,-57.94,-22.09,-58.17,-20.18,-57.5,-18.17,-58.28,-17.27,-58.24,-16.3,-60.16,-16.26,-60.5,-13.78,-64.32,-12.46,-65.4,-11.57,-65.34,-9.76,-66.65,-9.93,-68.27,-11.01,-70.55,-11.01,-70.48,-9.49,-72.18,-10.05,-73.23,-9.46,-73.02,-9.03,-73.99,-7.52,-73.12,-6.63,-72.89,-5.27,-69.89,-4.3,-69.42,-1.12,-70.02,0.54,-69.22,0.99,-69.8,1.09,-69.82,1.71,-67.54,2.04,-67.07,1.13,-65.55,0.79,-63.37,2.2,-64.27,2.5,-64.82,4.06,-63.09,3.77,-60.21,5.24,-59.54,3.96,-59.97,2.76,-59.03,1.32,-56.0,1.82,-55.97,2.51,-52.94,2.12,-51.32,4.2,-50.51,1.9,-49.97,1.74,-50.7,0.22,-50.39,-0.08,-48.62,-0.24,-48.58,-1.24,-47.82,-0.58,-44.91,-1.55,-44.58,-2.69,-43.42,-2.38,-39.98,-2.87,-37.22,-4.82,-35.6,-5.15,-34.73,-7.34,-35.13,-9.0,-38.67,-13.06,-39.27,-17.87,-40.94,-21.94,-41.99,-22.97,-44.65,-23.35,-47.65,-24.89,-48.5,-25.88,-48.89,-28.67,-53.37,-33.77,-53.65,-33.2,-53.21,-32.73,-53.79,-32.05,-56.98,-30.11,-57.63,-30.22]],"Brunei":[[114.2,4.53,115.45,5.45,115.35,4.32,114.66,4.01,114.2,4.53]],"Bhutan":[[91.7,27.77,92.03,26.84,88.84,27.1,90.02,28.3,91.7,27.77]],"Botswana":[[25.65,-18.54,27.72,-20.5,28.02,-21.49,29.43,-22.09,27.12,-23.57,25.66,-25.49,23.31,-25.27,20.89,-26.83,19.9,-24.77,19.9,-21.85,20.88,-21.81,20.91,-18.25,23.2,-17.87,23.58,-18.28,25.08,-17.66,25.65,-18.54]],"Central African Republic":[[15.28,7.42,17.96,7.89,18.81,8.98,21.0,9.48,22.86,11.14,23.55,10.09,23.46,8.95,27.37,5.23,24.41,5.11,22.84,4.71,22.41,4.03,19.47,5.03,18.45,3.5,17.13,3.73,16.01,2.27,14.48,4.73,14.54,6.23,15.28,7.42]],"Canada":[[-63.66,46.55,-62.01,46.44,-62.87,45.97,-64.39,46.73,-64.01,47.04,-63.66,46.55],[-123.51,48.51,-125.66,48.83,-128.06,49.99,-128.36,50.77,-125.76,50.3,-123.51,48.51],[-56.13,50.69,-56.8,49.81,-56.14,50.15,-55.47,49.94,-55.82,49.59,-53.48,49.25,-53.79,48.52,-53.09,48.69,-52.65,47.54,-53.07,46.66,-54.18,46.81,-54.24,47.75,-55.4,46.88,-56.0,46.92,-55.29,47.39,-56.25,47.63,-59.27,47.6,-58.8,48.25,-59.23,48.52,-57.36,50.72,-55.41,51.59,-56.13,50.69],[-132.71,54.04,-131.75,54.12,-132.05,52.98,-131.18,52.18,-133.05,53.41,-133.18,54.17,-132.71,54.04],[-81.9,62.71,-83.07,62.16,-83.99,62.45,-83.25,62.91,-81.9,62.71],[-85.16,65.66,-80.1,63.73,-80.99,63.41,-83.11,64.1,-85.52,63.05,-85.87,63.64,-87.22,63.54,-86.35,64.04,-85.88,65.74,-85.16,65.66],[-75.87,67.15,-76.99,67.1,-77.24,67.59,-76.81,68.15,-75.11,68.01,-75.87,67.15],[-95.65,69.11,-96.27,68.76,-99.8,69.4,-98.22,70.14,-95.65,69.11],[-90.55,69.5,-90.55,68.47,-89.22,69.26,-88.02,68.62,-88.32,67.87,-87.35,67.2,-85.58,68.78,-85.52,69.88,-82.62,69.66,-81.28,69.16,-81.96,68.13,-81.26,67.6,-81.39,67.11,-83.34,66.41,-85.77,66.56,-87.32,64.78,-89.91,64.03,-90.7,63.61,-90.77,62.96,-93.16,62.02,-94.24,60.9,-94.68,58.95,-93.22,58.78,-92.3,57.09,-90.9,57.28,-85.01,55.3,-82.27,55.15,-82.13,53.28,-79.91,51.21,-78.6,52.56,-79.83,54.67,-78.23,55.14,-76.54,56.53,-77.3,58.05,-78.52,58.8,-77.34,59.85,-78.11,62.32,-73.84,62.44,-71.37,61.14,-69.59,61.06,-69.29,58.96,-67.65,58.21,-66.2,58.77,-64.58,60.34,-61.4,56.97,-61.8,56.34,-57.33,54.63,-56.94,53.78,-55.76,53.27,-55.68,52.15,-60.03,50.24,-66.4,50.23,-71.1,46.82,-68.65,48.3,-65.06,49.23,-64.17,48.74,-65.12,48.07,-64.47,46.24,-61.52,45.88,-60.52,47.01,-59.8,45.92,-65.36,43.55,-66.12,43.62,-66.16,44.47,-64.43,45.29,-67.14,45.14,-67.79,45.7,-67.79,47.07,-69.24,47.45,-71.51,45.01,-74.87,45.0,-76.82,43.63,-78.72,43.63,-79.17,43.47,-78.94,42.86,-82.69,41.68,-83.12,42.08,-82.14,43.57,-82.55,45.35,-88.38,48.3,-91.64,48.14,-94.33,48.67,-94.82,49.39,-95.16,49.0,-122.97,49.0,-127.44,50.83,-127.85,52.33,-129.13,52.76,-129.31,53.56,-130.51,54.29,-130.01,55.92,-131.71,56.55,-135.48,59.79,-137.45,58.91,-139.04,60.0,-141.0,60.31,-140.99,69.71,-136.5,68.9,-129.79,70.19,-129.11,69.78,-128.14,70.48,-125.76,69.48,-124.42,70.16,-124.29,69.4,-121.47,69.8,-115.25,68.91,-113.9,68.4,-115.3,67.9,-113.5,67.69,-109.95,67.98,-108.88,67.38,-107.79,67.89,-108.81,68.31,-108.17,68.65,-106.15,68.8,-101.45,67.65,-98.44,67.78,-98.56,68.4,-97.67,68.58,-96.12,68.24,-96.13,67.29,-95.49,68.09,-94.69,68.06,-94.23,69.07,-96.47,70.09,-96.39,71.19,-95.21,71.92,-92.88,71.32,-91.52,70.19,-92.41,69.7,-90.55,69.5],[-114.17,73.12,-114.67,72.65,-112.44,72.96,-111.05,72.45,-109.92,72.96,-108.19,71.65,-107.69,72.07,-108.4,73.09,-106.52,73.08,-105.4,72.67,-104.46,70.99,-100.98,70.02,-101.09,69.58,-102.73,69.5,-102.09,69.12,-102.43,68.75,-105.96,69.18,-113.31,68.54,-117.34,69.96,-112.42,70.37,-117.9,70.54,-118.43,70.91,-116.11,71.31,-119.4,71.56,-117.87,72.71,-114.17,73.12],[-104.5,73.42,-105.38,72.76,-106.94,73.46,-104.5,73.42],[-76.34,73.1,-79.49,72.74,-80.88,73.33,-80.35,73.76,-78.06,73.65,-76.34,73.1],[-86.56,73.16,-85.77,72.53,-84.85,73.34,-82.32,73.75,-80.6,72.72,-80.75,72.06,-77.82,72.75,-74.23,71.77,-74.1,71.33,-72.24,71.56,-68.79,70.53,-66.97,69.19,-68.81,68.72,-61.85,66.86,-63.92,65.0,-66.72,66.39,-68.02,66.26,-68.14,65.69,-65.32,64.38,-64.67,63.39,-65.01,62.67,-68.78,63.75,-66.17,61.93,-71.02,62.91,-74.83,64.68,-77.71,64.23,-78.56,64.57,-77.9,65.31,-73.96,65.45,-74.29,65.81,-72.65,67.28,-72.93,67.73,-76.87,68.89,-76.23,69.15,-78.96,70.17,-84.94,69.97,-88.68,70.41,-89.51,70.76,-88.47,71.22,-89.89,71.22,-90.21,72.24,-89.44,73.13,-85.83,73.8,-86.56,73.16],[-100.36,73.84,-97.38,73.76,-97.12,73.47,-98.05,72.99,-96.54,72.56,-96.72,71.66,-99.32,71.36,-102.5,72.51,-100.44,72.71,-101.54,73.36,-100.36,73.84],[-93.2,72.77,-94.27,72.02,-95.41,72.06,-96.02,73.44,-94.5,74.13,-90.51,73.86,-93.2,72.77],[-120.46,71.38,-123.09,70.9,-125.93,71.87,-123.94,73.68,-124.92,74.29,-117.56,74.19,-115.51,73.48,-119.22,72.52,-120.46,71.38],[-93.61,74.98,-94.16,74.59,-96.82,74.93,-94.85,75.65,-93.61,74.98],[-98.5,76.72,-97.74,76.26,-98.16,75.0,-102.5,75.56,-102.57,76.34,-98.5,76.72],[-108.21,76.2,-105.88,75.97,-105.7,75.48,-106.31,75.01,-112.22,74.42,-113.87,74.72,-111.79,75.16,-117.71,75.22,-115.4,76.48,-109.07,75.47,-110.5,76.43,-109.58,76.79,-108.21,76.2],[-94.68,77.1,-91.61,76.78,-90.74,76.45,-90.97,76.07,-89.19,75.61,-81.13,75.71,-79.83,74.92,-81.95,74.44,-89.76,74.52,-92.42,74.84,-92.89,75.88,-93.89,76.32,-97.12,76.75,-96.75,77.16,-94.68,77.1],[-116.2,77.65,-116.34,76.88,-117.11,76.53,-122.85,76.12,-119.1,77.51,-116.2,77.65],[-110.19,77.7,-113.53,77.73,-109.85,78.0,-110.19,77.7],[-109.66,78.6,-112.54,78.41,-111.5,78.85,-109.66,78.6],[-95.83,78.06,-98.12,78.08,-98.63,78.87,-95.56,78.42,-95.83,78.06],[-100.06,78.32,-99.67,77.91,-105.18,78.38,-104.21,78.68,-105.49,79.3,-100.06,78.32],[-87.02,79.66,-85.81,79.34,-90.8,78.22,-92.88,78.34,-93.95,78.75,-93.15,79.38,-94.97,79.37,-96.71,80.16,-94.3,80.98,-94.74,81.21,-92.41,81.26,-87.81,80.32,-87.02,79.66],[-68.5,83.11,-61.85,82.63,-67.66,81.5,-65.48,81.51,-71.18,79.8,-76.91,79.32,-75.53,79.2,-76.22,79.02,-75.39,78.53,-79.76,77.21,-77.89,76.78,-80.56,76.18,-89.49,76.47,-89.62,76.95,-87.77,77.18,-88.26,77.9,-84.98,77.54,-87.96,78.37,-85.09,79.35,-86.93,80.25,-81.85,80.46,-87.6,80.52,-91.59,81.89,-85.5,82.65,-83.18,82.32,-82.42,82.86,-79.31,83.13,-68.5,83.11]],"Switzerland":[[9.59,47.53,10.36,46.48,7.27,45.78,6.02,46.27,6.74,47.54,9.59,47.53]],"Chile":[[-68.63,-52.64,-68.63,-54.87,-66.96,-54.9,-68.15,-55.61,-71.01,-55.05,-74.66,-52.84,-71.11,-54.07,-70.27,-52.93,-68.63,-52.64],[-68.22,-21.49,-67.83,-22.87,-66.99,-22.99,-67.33,-24.03,-68.42,-24.52,-68.3,-26.9,-69.66,-28.46,-70.54,-31.37,-69.82,-34.19,-70.36,-36.01,-71.12,-36.66,-70.81,-38.55,-71.41,-38.92,-72.15,-42.25,-71.22,-44.78,-71.66,-44.97,-72.33,-48.24,-73.42,-49.32,-73.33,-50.38,-72.31,-50.68,-71.91,-52.01,-68.57,-52.3,-70.85,-52.9,-71.43,-53.86,-74.95,-52.26,-75.61,-48.67,-74.13,-46.94,-75.64,-46.65,-74.69,-45.76,-74.35,-44.1,-73.24,-44.45,-72.72,-42.38,-73.39,-42.12,-73.7,-43.37,-74.33,-43.22,-73.22,-39.26,-73.59,-37.16,-73.17,-37.12,-71.44,-32.42,-71.49,-28.86,-70.91,-27.64,-70.09,-21.39,-70.37,-18.35,-69.59,-17.58,-68.44,-19.41,-68.76,-20.37,-68.22,-21.49]],"China":[[110.34,18.68,109.48,18.2,108.66,18.51,108.63,19.37,110.79,20.08,110.34,18.68],[127.66,49.76,129.4,49.44,130.58,48.73,130.99,47.79,135.03,48.48,133.1,45.14,131.03,44.97,131.14,42.93,130.63,42.9,130.64,42.4,129.99,42.99,128.05,41.99,128.21,41.47,126.87,41.82,124.27,39.93,121.05,38.9,122.17,40.42,121.64,40.95,117.53,38.74,119.7,37.16,120.82,37.87,122.36,37.45,122.52,36.93,121.1,36.65,119.15,34.91,120.23,34.36,121.91,31.69,121.89,30.95,121.26,30.68,122.09,29.83,121.68,28.23,121.13,28.14,118.66,24.55,115.89,22.78,110.79,21.4,110.44,20.34,109.89,20.28,109.86,21.4,107.04,21.81,106.73,22.79,105.33,23.35,101.65,22.32,101.8,21.17,101.27,21.2,101.15,21.85,100.42,21.56,99.24,22.12,99.53,22.95,98.9,23.14,98.66,24.06,97.6,23.9,97.72,25.08,98.67,25.92,98.68,27.51,97.91,28.34,96.25,28.41,96.59,28.83,96.12,29.45,94.57,29.28,92.5,27.9,90.02,28.3,88.81,27.3,88.73,28.09,85.82,28.2,79.72,30.88,78.74,31.52,78.46,32.62,79.18,32.48,78.91,34.32,77.84,35.49,76.19,35.9,74.98,37.42,74.86,38.38,73.93,38.51,73.68,39.43,74.78,40.37,76.53,40.43,76.9,41.07,80.12,42.12,80.18,42.92,80.87,43.18,79.97,44.92,82.46,45.54,83.18,47.33,85.16,47.0,85.77,48.46,87.75,49.3,88.01,48.6,90.28,47.69,90.97,46.89,90.59,45.72,90.95,45.29,95.31,44.24,96.35,42.73,100.85,42.66,104.96,41.6,110.41,42.87,111.83,43.74,111.35,44.46,111.87,45.1,113.46,44.81,117.42,46.67,119.66,46.69,118.06,48.07,115.74,47.73,115.49,48.14,116.68,49.89,117.88,49.51,119.29,50.14,120.74,51.96,120.18,52.75,121.0,53.25,123.57,53.46,125.95,52.79,127.66,49.76]],"Ivory Coast":[[-2.86,4.99,-4.65,5.17,-7.71,4.36,-7.57,5.71,-8.6,6.47,-7.83,8.58,-8.23,10.13,-6.21,10.52,-4.33,9.61,-2.83,9.64,-2.56,8.22,-3.24,6.25,-2.86,4.99]],"Cameroon":[[13.08,2.27,9.65,2.28,9.8,3.07,8.5,4.77,9.23,6.44,10.12,7.04,11.06,6.64,11.75,6.98,14.42,11.57,14.5,12.86,15.47,9.98,14.17,10.02,13.95,9.55,15.44,7.69,14.54,6.23,14.48,4.73,15.86,3.01,15.94,1.73,13.08,2.27]],"Democratic Republic of the Congo":[[30.83,3.51,30.77,2.34,31.17,2.2,29.88,0.6,29.02,-2.84,29.42,-5.94,30.74,-8.34,29.0,-8.41,28.45,-9.16,28.37,-11.79,29.62,-12.18,29.7,-13.26,28.93,-13.25,27.16,-11.61,26.55,-11.92,24.26,-10.95,22.16,-11.08,21.73,-7.29,20.09,-6.94,19.02,-7.99,17.47,-8.07,16.33,-5.88,12.32,-6.1,12.63,-4.99,13.6,-4.5,14.58,-4.97,16.01,-3.54,16.41,-1.74,17.64,-0.42,18.54,4.2,19.47,5.03,22.41,4.03,22.84,4.71,24.41,5.11,27.37,5.23,27.98,4.41,29.72,4.6,30.83,3.51]],"Republic of the Congo":[[13.0,-4.78,12.62,-4.44,11.91,-5.04,11.09,-3.98,11.86,-3.43,11.48,-2.77,12.58,-1.95,13.99,-2.47,14.3,-2.0,13.84,0.04,14.28,1.2,13.28,1.31,13.08,2.27,15.94,1.73,17.13,3.73,18.45,3.5,17.64,-0.42,16.41,-1.74,16.01,-3.54,14.58,-4.97,14.14,-4.51,13.0,-4.78]],"Colombia":[[-75.37,-0.15,-77.42,0.4,-78.99,1.69,-77.13,3.85,-77.75,7.71,-77.24,7.94,-77.35,8.67,-75.67,9.44,-74.91,11.08,-73.41,11.23,-71.4,12.38,-71.33,11.78,-72.91,10.45,-73.3,9.15,-72.79,9.09,-71.96,6.99,-70.09,6.96,-69.39,6.1,-67.34,6.1,-67.82,4.5,-67.3,3.32,-67.81,2.82,-66.88,1.25,-67.54,2.04,-69.82,1.71,-69.8,1.09,-69.22,0.99,-70.02,0.54,-69.42,-1.12,-69.89,-4.3,-70.69,-3.74,-70.05,-2.73,-70.81,-2.26,-73.07,-2.31,-73.66,-1.26,-75.37,-0.15]],"Costa Rica":[[-82.97,8.23,-84.98,10.09,-85.11,9.56,-85.66,9.93,-85.94,10.9,-83.66,10.94,-82.55,9.57,-82.97,8.23]],"Cuba":[[-82.27,23.19,-78.35,22.51,-74.18,20.28,-77.76,19.86,-77.09,20.41,-78.14,20.74,-78.72,21.6,-82.17,22.39,-81.8,22.64,-84.97,21.9,-82.27,23.19]],"Cyprus":[[33.97,35.06,32.98,34.57,32.26,35.1,33.97,35.06]],"Czech Republic":[[16.96,48.6,15.25,49.04,14.34,48.56,12.52,49.55,12.24,50.27,15.02,51.11,18.85,49.5,16.96,48.6]],"Germany":[[9.92,54.98,10.94,54.01,12.52,54.47,14.35,53.25,15.02,51.11,12.24,50.27,13.6,48.88,12.88,48.29,12.93,47.47,7.47,47.62,8.1,49.02,6.19,49.46,5.99,51.85,6.84,52.23,6.91,53.48,8.8,54.02,8.53,54.96,9.92,54.98]],"Djibouti":[[43.08,12.7,42.78,10.93,41.76,11.05,42.35,12.54,43.08,12.7]],"Denmark":[[12.69,55.61,12.09,54.8,10.9,55.78,12.37,56.11,12.69,55.61],[10.91,56.46,9.65,55.47,9.92,54.98,8.53,54.96,8.12,55.52,8.54,57.11,10.58,57.73,10.25,56.89,10.91,56.46]],"Dominican Republic":[[-71.71,19.71,-69.95,19.65,-68.32,18.61,-68.69,18.21,-70.67,18.43,-71.4,17.6,-71.95,18.62,-71.71,19.71]],"Algeria":[[12.0,23.47,5.68,19.6,3.16,19.06,3.15,19.69,-8.68,27.4,-8.67,28.84,-5.24,30.0,-3.69,30.9,-3.65,31.64,-1.31,32.26,-2.17,35.17,-1.21,35.71,1.47,36.61,8.42,36.95,8.14,34.66,7.52,34.1,7.61,33.34,9.06,32.1,9.81,29.42,9.32,26.09,10.3,24.38,10.77,24.56,12.0,23.47]],"Ecuador":[[-80.3,-3.4,-79.77,-2.66,-80.97,-2.25,-80.93,-1.06,-80.09,0.77,-78.86,1.38,-75.37,-0.15,-75.54,-1.56,-77.84,-3.0,-79.21,-4.96,-80.44,-4.43,-80.3,-3.4]],"Egypt":[[34.92,29.5,33.92,27.65,32.32,29.76,35.69,23.93,35.53,23.1,36.87,22.0,25.0,22.0,24.7,30.04,25.16,31.57,28.91,30.87,30.98,31.56,31.96,30.93,34.27,31.22,34.92,29.5]],"Eritrea":[[42.35,12.54,40.03,14.52,37.91,14.96,37.59,14.21,36.43,14.42,36.85,16.96,38.41,18.0,39.27,15.92,43.08,12.7,42.35,12.54]],"Spain":[[-9.03,41.88,-9.39,43.03,-7.98,43.75,-1.9,43.42,0.34,42.58,2.99,42.47,3.04,41.89,0.81,41.01,-0.68,37.64,-2.15,36.67,-4.37,36.68,-5.38,35.95,-7.54,37.43,-7.03,38.08,-7.5,39.63,-6.39,41.38,-6.67,41.88,-8.01,41.79,-8.26,42.28,-9.03,41.88]],"Estonia":[[24.31,57.79,24.43,58.38,23.43,58.61,23.34,59.19,25.86,59.61,28.13,59.3,27.42,58.72,27.72,57.79,27.29,57.47,24.31,57.79]],"Ethiopia":[[37.91,14.96,40.03,14.52,41.6,13.45,42.35,12.54,41.76,11.05,42.78,10.93,42.56,10.57,43.68,9.18,47.79,8.0,44.96,5.0,43.66,4.96,41.86,3.92,40.77,4.26,39.56,3.42,38.12,3.6,36.16,4.45,34.71,6.59,32.95,7.78,33.83,8.38,34.26,10.63,35.86,12.58,36.43,14.42,37.59,14.21,37.91,14.96]],"Finland":[[28.59,69.06,28.45,68.36,29.98,67.7,29.05,66.94,30.22,65.81,29.54,64.95,30.44,64.2,30.04,63.55,31.52,62.87,31.14,62.36,28.07,60.5,22.87,59.85,21.32,60.72,21.54,61.71,21.06,62.61,21.54,63.19,25.4,65.11,23.57,66.4,23.54,67.94,20.65,69.11,24.74,68.65,26.18,69.83,27.73,70.16,29.02,69.77,28.59,69.06]],"Fiji":[[178.37,-17.34,178.55,-18.15,177.38,-18.16,177.67,-17.38,178.37,-17.34]],"Falkland Islands":[[-61.2,-51.85,-58.55,-51.1,-57.75,-51.55,-59.4,-52.2,-61.2,-51.85]],"France":[[9.56,42.15,9.23,41.38,8.54,42.26,9.39,43.01,9.56,42.15],[3.59,50.38,8.1,49.02,7.47,47.62,6.04,46.73,6.84,45.99,7.44,43.69,6.53,43.13,3.1,43.08,2.99,42.47,1.83,42.34,-1.5,43.03,-1.9,43.42,-1.19,46.01,-2.96,47.57,-4.49,47.95,-4.59,48.68,-1.62,48.64,-1.93,49.78,-0.99,49.35,1.34,50.13,1.64,50.95,2.51,51.15,3.59,50.38]],"Gabon":[[11.09,-3.98,8.83,-0.78,9.49,1.01,11.29,1.06,11.28,2.26,12.95,2.32,13.28,1.31,14.28,1.2,13.84,0.04,14.3,-2.0,13.99,-2.47,12.58,-1.95,11.48,-2.77,11.86,-3.43,11.09,-3.98]],"United Kingdom":[[-5.66,54.55,-6.2,53.87,-7.57,54.06,-7.57,55.13,-5.66,54.55],[-3.01,58.63,-4.07,57.55,-1.96,57.68,-3.12,55.97,-2.09,55.91,0.47,52.93,1.68,52.74,1.05,51.81,1.45,51.29,-5.25,49.96,-5.78,50.16,-3.41,51.43,-5.27,51.99,-4.22,52.3,-4.77,52.84,-4.58,53.5,-3.09,53.4,-2.95,53.98,-4.84,54.79,-5.05,55.78,-5.59,55.31,-6.15,56.79,-5.01,58.63,-3.01,58.63]],"Georgia":[[41.55,41.54,41.45,42.65,40.08,43.55,45.47,42.5,46.64,41.18,41.55,41.54]],"Ghana":[[1.06,5.93,-1.96,4.71,-2.86,4.99,-3.24,6.25,-2.56,8.22,-2.94,10.96,0.02,11.02,1.06,5.93]],"Guinea":[[-8.44,7.69,-9.21,7.31,-9.76,8.54,-10.51,8.35,-11.12,10.05,-12.43,9.84,-13.25,8.9,-15.13,11.04,-13.74,11.81,-13.7,12.59,-10.17,11.84,-9.13,12.31,-8.03,10.21,-7.83,8.58,-8.44,7.69]],"Gambia":[[-16.84,13.15,-15.4,13.86,-13.84,13.51,-16.84,13.15]],"Guinea Bissau":[[-15.13,11.04,-16.68,12.38,-13.7,12.59,-13.74,11.81,-15.13,11.04]],"Equatorial Guinea":[[9.49,1.01,9.65,2.28,11.28,2.26,11.29,1.06,9.49,1.01]],"Greece":[[23.7,35.71,26.29,35.3,24.72,34.92,23.51,35.28,23.7,35.71],[26.6,41.56,26.06,40.82,23.71,40.69,24.41,40.12,23.9,39.96,22.81,40.48,23.35,39.19,22.97,38.97,24.04,37.66,23.12,37.92,23.41,37.41,22.77,37.31,23.15,36.42,21.67,36.84,20.15,39.62,21.02,40.84,26.6,41.56]],"Greenland":[[-46.76,82.63,-38.62,83.55,-27.1,83.52,-20.85,82.73,-31.9,82.2,-22.07,81.73,-23.17,81.15,-15.77,81.91,-12.21,81.29,-20.05,80.18,-17.73,80.13,-19.7,78.75,-19.67,77.64,-18.47,76.99,-21.68,76.63,-19.83,76.1,-19.6,75.25,-20.67,75.16,-19.37,74.3,-21.59,74.22,-20.43,73.82,-20.76,73.46,-23.57,73.31,-22.3,72.18,-24.79,72.33,-22.13,71.47,-21.75,70.66,-23.54,70.47,-25.54,71.43,-25.2,70.75,-26.36,70.23,-22.35,70.13,-27.75,68.47,-31.78,68.12,-34.2,66.68,-39.81,65.46,-41.19,63.48,-42.82,62.68,-42.42,61.9,-43.38,60.1,-48.26,60.86,-51.63,63.63,-52.28,65.18,-53.66,66.1,-53.3,66.84,-53.97,67.19,-52.98,68.36,-51.48,68.73,-50.87,69.93,-53.46,69.28,-54.68,69.61,-54.36,70.82,-51.39,70.57,-55.83,71.65,-54.72,72.59,-58.59,75.52,-61.27,76.1,-68.5,76.06,-71.4,77.01,-66.76,77.38,-73.3,78.04,-73.16,78.43,-65.71,79.39,-65.32,79.76,-68.02,80.12,-62.23,81.32,-62.65,81.77,-57.21,82.19,-53.04,81.89,-50.39,82.44,-44.52,81.66,-46.9,82.2,-46.76,82.63]],"Guatemala":[[-90.1,13.74,-91.69,14.13,-92.23,15.25,-91.75,16.07,-90.46,16.07,-91.45,17.25,-91.0,17.25,-91.0,17.82,-89.14,17.81,-89.23,15.89,-88.23,15.73,-90.1,13.74]],"French Guiana":[[-52.56,2.5,-53.42,2.05,-54.52,2.31,-54.01,3.62,-54.48,4.9,-53.96,5.76,-51.82,4.57,-52.56,2.5]],"Guyana":[[-59.76,8.37,-57.15,5.97,-58.04,4.06,-56.54,1.9,-58.54,1.27,-59.65,1.79,-59.98,5.01,-61.41,5.96,-59.76,8.37]],"Honduras":[[-87.32,12.98,-87.86,13.89,-89.35,14.42,-88.23,15.73,-84.98,16.0,-83.15,15.0,-84.92,14.79,-87.32,12.98]],"Croatia":[[18.83,45.91,19.39,45.24,19.01,44.86,15.96,45.23,15.75,44.82,18.45,42.48,16.02,43.51,14.9,45.08,13.66,45.14,15.33,45.45,16.56,46.5,18.83,45.91]],"Haiti":[[-73.19,19.92,-71.71,19.71,-71.71,18.04,-74.46,18.34,-72.33,18.67,-73.19,19.92]],"Hungary":[[16.2,46.85,16.98,48.12,17.86,47.76,20.8,48.62,22.71,47.88,21.02,46.32,18.46,45.76,16.2,46.85]],"Indonesia":[[120.72,-10.24,118.97,-9.56,119.9,-9.36,120.72,-10.24],[124.44,-10.14,123.46,-10.24,123.98,-9.29,124.97,-8.89,124.44,-10.14],[117.9,-8.1,119.13,-8.71,116.74,-9.03,117.9,-8.1],[122.9,-8.09,122.76,-8.65,119.92,-8.81,120.72,-8.24,122.9,-8.09],[108.62,-6.78,110.54,-6.88,110.76,-6.47,115.71,-8.37,114.56,-8.75,105.37,-6.85,106.05,-5.9,108.62,-6.78],[130.47,-3.09,130.83,-3.86,127.9,-3.39,128.14,-2.84,130.47,-3.09],[134.14,-1.15,134.42,-2.77,135.46,-3.37,137.44,-1.7,141.0,-2.6,141.03,-9.12,140.14,-8.3,137.61,-8.41,138.67,-7.32,137.93,-5.39,133.66,-3.54,132.98,-4.11,131.99,-2.82,133.7,-2.21,132.23,-2.21,130.52,-0.94,132.38,-0.37,134.14,-1.15],[125.24,1.42,123.69,0.24,120.18,0.24,120.04,-0.52,120.94,-1.41,123.34,-0.62,121.51,-1.9,123.16,-5.34,122.24,-5.28,122.72,-4.46,121.49,-4.57,120.97,-2.63,120.31,-2.93,120.43,-5.53,119.37,-5.38,119.5,-3.49,118.77,-2.8,120.04,0.57,120.89,1.31,122.93,0.88,125.24,1.42],[128.69,1.13,128.1,-0.9,127.4,1.01,127.93,2.17,128.69,1.13],[117.88,1.83,119.0,0.9,117.81,0.78,117.52,-0.8,116.56,-1.49,116.15,-4.01,116.0,-3.66,114.86,-4.11,113.26,-3.12,112.07,-3.48,111.7,-2.99,110.22,-2.93,110.07,-1.59,109.09,-0.46,109.07,1.34,109.66,2.01,110.51,0.77,114.62,1.43,115.87,4.31,117.88,4.14,117.31,3.23,117.88,1.83],[105.82,-5.85,104.71,-5.87,102.58,-4.22,98.6,1.82,95.29,5.48,97.48,5.25,100.64,2.1,101.66,2.08,103.84,0.1,103.44,-0.71,106.11,-3.06,105.82,-5.85]],"India":[[77.84,35.49,78.91,34.32,79.21,32.99,79.18,32.48,78.46,32.62,78.74,31.52,81.11,30.18,80.09,28.79,83.3,27.36,88.06,26.41,88.12,27.88,88.73,28.09,88.84,27.1,89.74,26.72,92.03,26.84,91.7,27.77,94.57,29.28,96.12,29.45,96.59,28.83,96.25,28.41,97.4,27.88,97.13,27.08,96.42,27.26,95.12,26.57,94.11,23.85,93.33,24.08,93.17,22.28,92.67,22.04,92.15,23.63,91.71,22.99,91.16,23.5,92.38,24.98,89.92,25.27,89.83,25.97,88.56,26.45,88.21,25.77,88.93,25.24,88.08,24.5,88.7,24.23,88.89,21.69,86.98,21.5,86.5,20.15,85.06,19.48,82.19,16.56,80.32,15.9,79.86,10.36,77.54,7.97,76.59,8.9,73.53,15.99,72.63,21.36,70.47,20.88,69.16,22.09,69.64,22.45,69.35,22.84,68.18,23.69,68.84,24.36,71.04,24.36,69.51,26.94,70.62,27.99,71.78,27.91,75.26,32.27,74.45,32.76,73.75,34.32,74.24,34.75,76.87,34.65,77.84,35.49]],"Ireland":[[-6.2,53.87,-6.03,53.15,-6.79,52.26,-9.98,51.82,-9.17,52.86,-9.69,53.88,-7.57,55.13,-7.57,54.06,-6.2,53.87]],"Iran":[[53.92,37.2,56.62,38.12,61.12,36.49,60.54,32.98,61.78,30.74,60.87,29.83,62.73,28.26,63.32,26.76,61.87,26.24,61.5,25.08,57.4,25.74,56.49,27.14,54.72,26.48,53.49,26.81,51.52,27.87,50.12,30.15,48.57,29.93,47.33,32.47,45.42,33.97,46.08,35.68,44.23,37.97,44.11,39.43,44.79,39.71,46.14,38.74,48.06,39.58,48.01,38.79,49.2,37.58,50.84,36.87,53.92,37.2]],"Iraq":[[45.42,35.98,46.15,35.09,45.42,33.97,47.33,32.47,48.57,29.93,47.3,30.06,46.57,29.1,44.71,29.18,41.89,31.19,39.2,32.16,38.79,33.38,41.01,34.42,41.29,36.36,42.35,37.23,44.77,37.17,45.42,35.98]],"Iceland":[[-14.51,66.46,-14.74,65.81,-13.61,65.13,-18.66,63.5,-22.76,63.96,-21.78,64.4,-23.96,64.89,-22.23,65.38,-24.33,65.61,-23.65,66.26,-22.13,66.41,-20.58,65.73,-19.06,66.28,-14.51,66.46]],"Israel":[[35.72,32.71,34.97,31.87,35.42,31.1,34.92,29.5,34.27,31.22,35.1,33.08,35.82,33.28,35.72,32.71]],"Italy":[[15.52,38.23,15.1,36.62,12.43,37.61,12.57,38.13,15.52,38.23],[9.21,41.21,9.81,40.5,9.67,39.18,8.81,38.91,8.16,40.95,9.21,41.21],[12.38,46.77,13.81,46.51,13.94,45.59,12.33,45.38,12.59,44.09,15.14,41.96,15.93,41.96,15.89,41.54,18.48,40.17,18.29,39.81,16.87,40.44,16.45,39.8,17.17,39.42,17.05,38.9,16.1,37.99,15.68,37.91,16.11,38.96,15.41,40.05,11.19,42.36,10.2,43.92,8.89,44.37,7.44,43.69,7.55,44.13,7.01,44.25,6.84,45.99,8.97,46.04,10.44,46.89,12.38,46.77]],"Jamaica":[[-77.57,18.49,-76.2,17.89,-78.34,18.23,-77.57,18.49]],"Jordan":[[35.55,32.39,36.83,32.31,38.79,33.38,39.2,32.16,37.0,31.51,38.0,30.51,36.07,29.2,34.92,29.5,35.55,32.39]],"Japan":[[134.64,34.15,134.2,33.2,133.79,33.52,133.01,32.7,132.36,32.99,132.92,34.06,134.64,34.15],[140.98,37.14,140.25,35.14,137.22,34.61,135.79,33.46,135.12,33.85,135.08,34.6,130.99,33.89,132.0,33.15,131.33,31.45,130.69,31.03,130.2,31.42,130.45,32.32,129.41,33.3,132.62,35.43,135.68,35.53,136.72,37.3,137.39,36.83,139.43,38.22,140.05,39.44,139.88,40.56,140.31,41.2,141.37,41.38,141.88,39.18,140.96,38.17,140.98,37.14],[143.91,44.17,145.32,44.38,145.54,43.26,144.06,42.99,143.18,42.0,141.61,42.68,141.07,41.58,139.96,41.57,139.82,42.56,140.31,43.33,141.38,43.39,141.97,45.55,143.91,44.17]],"Kazakhstan":[[70.96,42.27,68.26,40.66,67.99,41.14,66.71,41.17,66.51,41.99,66.02,41.99,66.1,43.0,64.9,43.73,62.01,43.5,58.5,45.59,55.93,45.0,55.97,41.31,54.08,42.32,52.5,41.78,52.5,42.79,51.34,43.13,50.31,44.61,51.28,44.51,51.32,45.25,53.04,45.26,53.04,46.85,51.19,47.05,49.1,46.4,48.06,47.74,46.47,48.39,47.55,50.45,48.58,49.87,48.7,50.61,50.77,51.69,52.33,51.72,55.72,50.62,56.78,51.04,61.34,50.8,61.59,51.27,59.97,51.96,61.7,52.98,60.98,53.66,61.44,54.01,69.07,55.39,70.87,55.17,71.18,54.13,73.51,54.04,73.43,53.49,76.89,54.49,76.53,54.18,80.04,50.86,80.57,51.39,81.95,50.81,83.38,51.07,87.36,49.21,85.77,48.46,85.16,47.0,83.18,47.33,82.46,45.54,79.97,44.92,80.87,43.18,80.18,42.92,80.26,42.35,74.21,43.3,73.49,42.5,71.84,42.85,70.96,42.27]],"Kenya":[[40.99,-0.86,41.59,-1.68,40.26,-2.57,39.2,-4.68,37.77,-3.68,37.7,-3.1,33.9,-0.95,33.89,0.11,35.04,1.91,34.01,4.25,35.3,5.51,36.16,4.45,38.12,3.6,39.56,3.42,40.77,4.26,41.86,3.92,40.98,2.78,40.99,-0.86]],"Kyrgyzstan":[[70.96,42.27,71.84,42.85,73.49,42.5,74.21,43.3,80.26,42.35,76.9,41.07,76.53,40.43,74.78,40.37,73.68,39.43,69.46,39.53,69.56,40.1,71.77,40.15,73.06,40.87,70.42,41.52,70.96,42.27]],"Cambodia":[[103.5,10.63,102.35,13.39,102.99,14.23,106.04,13.88,106.5,14.57,107.38,14.2,107.49,12.34,105.81,11.57,106.25,10.96,103.5,10.63]],"South Korea":[[128.35,38.61,129.46,36.78,129.47,35.63,129.09,35.08,126.49,34.39,126.12,36.73,126.86,36.89,126.17,37.75,128.35,38.61]],"Kosovo":[[20.76,42.05,20.07,42.59,20.81,43.27,21.78,42.68,20.76,42.05]],"Kuwait":[[47.97,29.98,48.42,28.55,46.57,29.1,47.3,30.06,47.97,29.98]],"Laos":[[105.22,14.27,105.59,15.57,103.96,18.24,101.06,17.51,101.28,19.46,100.61,19.51,100.12,20.42,101.18,21.44,101.8,21.17,101.65,22.32,102.17,22.46,103.2,20.77,104.44,20.76,104.82,19.89,103.9,19.27,105.09,18.67,107.31,15.91,107.38,14.2,106.5,14.57,106.04,13.88,105.22,14.27]],"Lebanon":[[35.82,33.28,35.13,33.09,36.0,34.64,36.45,34.59,35.82,33.28]],"Liberia":[[-7.71,4.36,-11.44,6.79,-10.23,8.41,-9.76,8.54,-9.21,7.31,-8.44,7.69,-8.6,6.47,-7.57,5.71,-7.71,4.36]],"Libya":[[14.85,22.86,14.14,22.49,10.77,24.56,10.3,24.38,9.32,26.09,9.86,28.96,9.48,30.31,9.97,30.54,9.95,31.38,11.43,32.37,11.49,33.14,15.25,32.27,15.71,31.38,19.09,30.27,20.05,30.99,19.82,31.75,20.85,32.71,24.92,31.9,25.0,20.0,23.85,20.0,23.84,19.58,15.86,23.41,14.85,22.86]],"Sri Lanka":[[81.79,7.52,81.64,6.48,80.35,5.97,79.7,8.2,80.15,9.82,81.79,7.52]],"Lesotho":[[28.98,-28.96,29.33,-29.26,28.11,-30.55,27.0,-29.88,28.07,-28.85,28.98,-28.96]],"Lithuania":[[22.73,54.33,22.76,54.86,21.27,55.19,21.06,56.03,24.86,56.37,26.49,55.62,25.54,54.28,24.45,53.91,22.73,54.33]],"Latvia":[[21.06,56.03,21.58,57.41,22.52,57.75,23.32,57.01,24.12,57.03,24.31,57.79,25.16,57.97,27.77,57.24,28.18,56.17,26.49,55.62,24.86,56.37,21.06,56.03]],"Morocco":[[-5.19,35.76,-2.17,35.17,-1.31,32.26,-3.65,31.64,-3.69,30.9,-5.24,30.0,-8.67,28.84,-8.79,27.12,-11.39,26.88,-12.5,24.77,-13.89,23.69,-14.75,21.5,-17.02,21.42,-14.44,26.25,-9.56,29.93,-9.81,31.18,-9.3,32.56,-6.91,34.11,-5.93,35.76,-5.19,35.76]],"Moldova":[[26.62,48.22,28.67,48.12,30.02,46.42,28.86,46.44,28.23,45.49,28.13,46.81,26.62,48.22]],"Madagascar":[[49.54,-12.47,50.38,-15.71,49.67,-15.71,49.77,-16.88,47.1,-24.94,45.41,-25.6,44.04,-24.99,43.35,-22.78,43.43,-21.34,44.46,-19.44,43.96,-17.41,44.45,-16.22,46.31,-15.78,47.71,-14.59,49.19,-12.04,49.54,-12.47]],"Mexico":[[-97.14,25.87,-97.7,21.9,-95.9,18.83,-94.43,18.14,-90.77,19.28,-90.28,21.0,-86.81,21.33,-87.84,18.26,-91.0,17.82,-91.0,17.25,-91.45,17.25,-90.46,16.07,-91.75,16.07,-92.23,14.54,-93.88,15.94,-96.56,15.65,-104.99,19.32,-105.73,20.43,-105.27,21.42,-106.03,22.77,-112.23,28.95,-113.15,31.17,-114.78,31.8,-114.67,30.16,-111.62,26.66,-110.66,24.3,-109.41,23.36,-110.03,22.82,-112.18,24.74,-112.3,26.01,-115.06,27.72,-114.16,28.57,-115.52,29.56,-117.13,32.54,-114.72,32.72,-111.02,31.33,-106.51,31.75,-103.94,29.27,-103.11,28.97,-102.48,29.76,-100.96,29.38,-99.02,26.37,-97.14,25.87]],"Macedonia":[[20.59,41.86,22.38,42.32,22.95,41.34,21.02,40.84,20.59,41.86]],"Mali":[[-12.17,14.62,-11.67,15.39,-5.54,15.5,-6.45,24.96,-4.92,24.97,3.15,19.69,3.16,19.06,4.27,19.16,4.27,16.85,3.64,15.57,-1.07,14.97,-4.01,13.47,-5.22,11.71,-5.4,10.37,-6.05,10.1,-6.21,10.52,-8.03,10.21,-9.13,12.31,-10.17,11.84,-11.46,12.08,-12.17,14.62]],"Myanmar":[[99.54,20.19,98.25,19.71,97.38,18.45,98.9,16.18,98.19,15.12,99.59,11.89,98.55,9.93,98.51,13.12,97.16,16.93,95.37,15.71,94.19,16.04,94.32,18.21,93.66,19.73,92.37,20.67,92.3,21.48,93.17,22.28,93.33,24.08,94.11,23.85,95.12,26.57,96.42,27.26,97.13,27.08,97.33,28.26,98.68,27.51,98.67,25.92,97.72,25.08,97.6,23.9,98.66,24.06,98.9,23.14,99.53,22.95,99.24,22.12,100.42,21.56,101.15,21.85,99.54,20.19]],"Montenegro":[[19.8,42.5,19.37,41.88,18.45,42.48,19.22,43.52,20.34,42.9,19.8,42.5]],"Mongolia":[[87.75,49.3,92.23,50.8,97.26,49.73,98.23,50.42,97.83,51.01,98.86,52.05,102.07,51.26,102.26,50.51,103.68,50.09,106.89,50.27,108.48,49.28,110.66,49.13,114.36,50.25,116.68,49.89,115.49,48.14,115.74,47.73,118.06,48.07,119.77,47.05,117.42,46.67,113.46,44.81,111.87,45.1,111.35,44.46,111.83,43.74,110.41,42.87,104.96,41.6,100.85,42.66,96.35,42.73,95.31,44.24,90.95,45.29,90.59,45.72,90.97,46.89,90.28,47.69,88.01,48.6,87.75,49.3]],"Mozambique":[[34.56,-11.52,37.47,-11.57,40.32,-10.32,40.78,-14.69,39.45,-16.72,37.41,-17.59,34.79,-19.78,35.56,-22.09,35.46,-24.12,33.01,-25.36,32.57,-25.73,32.83,-26.74,32.07,-26.73,31.19,-22.25,32.66,-20.3,32.85,-16.71,30.34,-15.88,30.18,-14.8,33.21,-13.97,34.46,-14.61,34.38,-16.18,35.03,-16.8,35.77,-15.9,35.69,-14.61,34.56,-13.58,34.56,-11.52]],"Mauritania":[[-12.17,14.62,-14.58,16.6,-16.46,16.14,-16.28,20.09,-17.06,21.0,-12.93,21.33,-12.87,23.28,-11.94,23.37,-11.97,25.93,-8.69,25.88,-8.68,27.4,-4.92,24.97,-6.45,24.96,-5.32,16.2,-5.54,15.5,-11.67,15.39,-12.17,14.62]],"Malawi":[[34.56,-11.52,34.56,-13.58,35.69,-14.61,35.77,-15.9,35.03,-16.8,34.38,-16.18,34.46,-14.61,32.69,-13.71,33.49,-10.53,32.76,-9.23,33.74,-9.42,34.56,-11.52]],"Malaysia":[[101.08,6.2,101.15,5.69,102.14,6.22,102.96,5.52,104.23,1.29,103.52,1.23,101.39,2.76,100.09,6.46,101.08,6.2],[118.62,4.48,115.87,4.31,114.62,1.43,110.51,0.77,109.83,1.34,109.66,2.01,111.17,1.85,111.37,2.7,113.0,3.1,114.2,4.53,114.66,4.01,115.35,4.32,115.45,5.45,116.73,6.92,119.18,5.41,118.44,4.97,118.62,4.48]],"Namibia":[[16.34,-28.58,15.21,-27.09,14.26,-22.11,11.73,-17.3,13.46,-16.97,14.06,-17.42,18.26,-17.31,21.38,-17.93,24.03,-17.3,25.08,-17.58,23.58,-18.28,23.2,-17.87,20.91,-18.25,20.88,-21.81,19.9,-21.85,19.89,-28.46,18.46,-29.05,16.82,-28.08,16.34,-28.58]],"New Caledonia":[[165.78,-21.08,167.12,-22.16,165.47,-21.68,164.03,-20.11,165.78,-21.08]],"Niger":[[2.15,11.94,2.18,12.63,1.02,12.85,0.37,14.93,3.64,15.57,4.27,16.85,4.27,19.16,12.0,23.47,14.14,22.49,14.85,22.86,15.1,21.31,15.9,20.39,15.25,16.63,13.97,15.68,13.54,14.37,13.95,13.35,14.6,13.33,14.18,12.48,13.08,13.6,12.3,13.04,10.99,13.39,9.01,12.83,5.44,13.87,4.11,13.53,3.61,11.66,2.85,12.24,2.15,11.94]],"Nigeria":[[8.5,4.77,5.9,4.26,4.33,6.27,2.69,6.26,2.72,8.51,3.71,10.06,4.11,13.53,5.44,13.87,9.01,12.83,13.32,13.56,14.58,12.09,14.42,11.57,11.75,6.98,11.06,6.64,10.12,7.04,9.23,6.44,8.5,4.77]],"Nicaragua":[[-85.71,11.09,-87.67,12.91,-84.92,14.79,-83.15,15.0,-83.66,10.94,-85.71,11.09]],"Netherlands":[[6.07,53.51,7.09,53.14,6.59,51.85,5.99,51.85,6.16,50.8,4.97,51.48,3.31,51.35,4.71,53.09,6.07,53.51]],"Norway":[[28.17,71.19,31.29,70.45,30.01,70.19,31.1,69.56,28.59,69.06,29.02,69.77,27.73,70.16,26.18,69.83,24.74,68.65,21.24,69.37,20.03,69.07,19.88,68.41,17.99,68.57,17.73,68.01,16.77,68.01,13.56,64.79,13.92,64.45,13.57,64.05,12.58,64.07,11.93,63.13,11.99,61.8,12.63,61.29,12.3,60.12,11.03,58.86,10.36,59.47,8.38,58.31,7.05,58.08,5.67,58.59,4.99,61.97,10.53,64.49,14.76,67.81,19.18,69.82,23.02,70.2,24.55,71.03,28.17,71.19],[24.72,77.85,20.73,77.68,21.42,77.94,20.81,78.25,22.88,78.45,24.72,77.85],[18.25,79.7,21.54,78.96,19.03,78.56,17.12,76.81,15.91,76.77,13.76,77.38,14.67,77.74,11.22,78.87,10.44,79.65,16.99,80.05,18.25,79.7],[25.45,80.41,27.41,80.06,23.02,79.4,17.37,80.32,25.45,80.41]],"Nepal":[[88.12,27.88,88.06,26.41,87.23,26.4,80.09,28.79,81.53,30.42,85.82,28.2,88.12,27.88]],"New Zealand":[[173.02,-40.92,174.25,-41.35,172.71,-43.37,173.08,-43.85,171.45,-44.24,170.62,-45.91,169.33,-46.64,166.68,-46.22,167.05,-45.11,170.52,-43.03,172.1,-40.96,172.8,-40.49,173.02,-40.92],[174.61,-36.16,175.34,-37.21,175.36,-36.53,176.76,-37.88,178.52,-37.7,177.97,-39.17,177.21,-39.15,176.01,-41.29,175.24,-41.69,174.65,-41.28,175.23,-40.46,174.9,-39.91,173.82,-39.51,174.57,-38.8,174.7,-37.38,172.64,-34.53,174.33,-35.27,174.61,-36.16]],"Oman":[[58.86,21.11,57.83,20.24,57.69,18.94,54.79,16.95,53.11,16.65,52.0,19.0,55.0,20.0,55.89,24.92,59.81,22.53,58.86,21.11]],"Pakistan":[[75.16,37.13,76.19,35.9,77.84,35.49,76.87,34.65,74.24,34.75,73.75,34.32,74.45,32.76,75.26,32.27,71.78,27.91,70.62,27.99,69.51,26.94,71.04,24.36,68.84,24.36,68.18,23.69,66.37,25.43,61.5,25.08,61.87,26.24,63.32,26.76,62.73,28.26,60.87,29.83,62.55,29.32,66.35,29.89,66.94,31.3,69.32,31.9,69.26,32.5,70.32,33.36,69.93,34.02,70.88,33.99,71.61,35.15,71.26,36.07,75.16,37.13]],"Panama":[[-77.88,7.22,-78.18,8.32,-79.12,9.0,-80.38,8.3,-80.0,7.55,-80.42,7.27,-81.72,8.11,-82.85,8.07,-82.93,9.48,-81.44,8.79,-79.57,9.61,-78.06,9.25,-77.24,7.94,-77.88,7.22]],"Peru":[[-69.59,-17.58,-70.37,-18.35,-76.01,-14.65,-79.76,-7.19,-81.25,-6.14,-81.1,-4.04,-80.3,-3.4,-80.44,-4.43,-79.21,-4.96,-77.84,-3.0,-75.54,-1.56,-75.11,-0.06,-73.07,-2.31,-70.81,-2.26,-70.05,-2.73,-70.69,-3.74,-69.89,-4.3,-72.89,-5.27,-73.12,-6.63,-73.99,-7.52,-73.02,-9.03,-73.23,-9.46,-72.18,-10.05,-70.48,-9.49,-70.55,-11.01,-69.53,-10.95,-68.67,-12.56,-69.34,-14.95,-68.96,-16.5,-69.59,-17.58]],"Philippines":[[126.38,8.41,126.54,7.19,126.2,6.27,125.83,7.29,125.36,6.79,125.4,5.58,124.22,6.16,124.24,7.36,123.61,7.83,121.92,7.19,123.49,8.69,123.84,8.24,125.47,8.99,125.41,9.76,126.22,9.29,126.38,8.41],[123.98,10.28,123.0,9.02,122.38,9.71,122.95,10.88,123.5,10.94,123.34,10.27,124.08,11.23,123.98,10.28],[118.5,9.32,117.17,8.37,119.51,11.37,119.69,10.55,118.5,9.32],[121.88,11.89,123.12,11.58,122.0,10.44,121.88,11.89],[125.5,12.16,125.78,11.05,125.01,11.31,125.28,10.36,124.8,10.13,124.3,11.5,124.88,11.79,124.27,12.56,125.5,12.16],[121.53,13.07,121.26,12.21,120.32,13.47,121.53,13.07],[121.32,18.5,122.25,18.48,122.52,17.09,121.66,15.93,121.73,14.33,123.95,13.78,124.08,12.54,122.93,13.55,122.67,13.19,122.03,13.78,120.63,13.86,120.99,14.53,120.07,14.97,119.88,16.36,120.29,16.03,120.72,18.51,121.32,18.5]],"Papua New Guinea":[[155.88,-6.82,155.17,-6.54,154.51,-5.14,155.88,-6.82],[151.98,-5.48,150.24,-6.32,148.32,-5.75,149.85,-5.51,150.14,-5.0,150.24,-5.53,150.81,-5.46,151.65,-4.76,151.54,-4.17,152.34,-4.31,151.98,-5.48],[147.19,-7.39,150.69,-10.58,147.91,-10.13,146.05,-8.07,144.74,-7.63,143.29,-8.25,143.41,-8.98,142.63,-9.33,141.03,-9.12,141.0,-2.6,144.58,-3.86,145.98,-5.47,147.65,-6.08,147.89,-6.61,146.97,-6.72,147.19,-7.39],[153.14,-4.5,152.83,-4.77,152.41,-3.79,150.66,-2.74,152.24,-3.24,153.14,-4.5]],"Poland":[[15.02,51.11,14.12,53.76,17.62,54.85,23.24,54.22,23.8,52.69,23.2,52.49,24.03,50.71,22.52,49.48,22.78,49.03,18.91,49.44,15.02,51.11]],"Puerto Rico":[[-66.28,18.51,-65.59,18.23,-67.18,17.95,-67.1,18.52,-66.28,18.51]],"North Korea":[[130.64,42.4,129.67,41.6,129.71,40.88,127.53,39.76,127.39,39.21,128.21,38.37,125.28,37.67,124.71,38.11,125.39,39.39,124.27,39.93,125.08,40.57,126.87,41.82,128.21,41.47,128.05,41.99,129.99,42.99,130.64,42.4]],"Portugal":[[-9.03,41.88,-8.26,42.28,-6.39,41.38,-7.5,39.63,-7.03,38.08,-7.86,36.84,-8.9,36.87,-8.84,38.27,-9.53,38.74,-8.77,40.76,-9.03,41.88]],"Paraguay":[[-62.69,-22.25,-61.79,-19.63,-59.12,-19.36,-58.18,-19.87,-57.94,-22.09,-55.8,-22.36,-55.4,-23.96,-54.29,-24.02,-54.79,-26.62,-56.49,-27.55,-58.62,-27.12,-57.78,-25.16,-60.85,-23.88,-62.69,-22.25]],"Qatar":[[50.81,24.75,51.29,26.11,51.61,25.22,51.39,24.63,50.81,24.75]],"Romania":[[22.71,47.88,26.92,48.12,28.13,46.81,28.23,45.49,29.63,45.04,28.84,44.91,28.56,43.71,27.24,44.18,22.94,43.82,22.71,44.58,21.56,44.77,20.22,46.13,22.71,47.88]],"Russia":[[143.65,50.75,144.65,48.98,143.17,49.31,142.56,47.86,143.53,46.84,143.51,46.14,142.75,46.74,142.09,45.97,142.18,50.95,141.59,51.94,141.68,53.3,142.61,53.76,142.21,54.23,142.65,54.37,143.65,50.75],[22.73,54.33,19.66,54.43,21.27,55.19,22.76,54.86,22.73,54.33],[-175.01,66.58,-174.34,66.34,-174.57,67.06,-171.86,66.91,-169.9,65.98,-172.53,65.44,-172.96,64.25,-176.21,65.36,-178.36,65.39,-178.9,65.74,-178.69,66.11,-179.88,65.87,-179.43,65.4,-180.0,64.98,-180.0,68.96,-174.93,67.21,-175.01,66.58],[180.0,70.83,178.73,71.1,180.0,71.52,180.0,70.83],[-178.69,70.89,-180.0,70.83,-180.0,71.52,-177.58,71.27,-178.69,70.89],[143.6,73.21,139.86,73.37,142.06,73.86,143.6,73.21],[150.73,75.08,149.58,74.69,146.12,75.17,150.73,75.08],[145.09,75.56,144.3,74.82,138.96,74.61,136.97,75.26,137.51,75.95,138.83,76.14,145.09,75.56],[57.54,70.72,53.68,70.76,51.6,71.47,51.46,72.01,54.43,73.63,53.51,73.75,55.9,74.63,55.63,75.08,61.17,76.25,68.16,76.94,68.85,76.54,58.48,74.31,55.42,72.37,55.62,71.54,57.54,70.72],[106.97,76.97,107.24,76.48,111.08,76.71,114.13,75.85,113.89,75.33,109.4,74.18,113.02,73.98,113.53,73.34,115.57,73.75,123.2,72.97,123.26,73.74,126.98,73.57,128.59,73.04,129.05,72.4,128.46,71.98,131.29,70.79,132.25,71.84,133.86,71.39,139.87,71.49,139.15,72.42,140.47,72.85,149.5,72.2,152.97,70.84,159.0,70.87,159.83,70.45,159.71,69.72,160.94,69.44,167.84,69.58,169.58,68.69,170.82,69.01,170.01,69.65,170.45,70.1,175.72,69.88,180.0,68.96,180.0,64.98,177.41,64.61,179.37,62.98,179.23,62.3,177.36,62.52,173.68,61.65,170.33,59.88,168.9,60.57,166.29,59.79,165.84,60.16,163.54,59.87,162.02,58.24,163.19,57.62,163.06,56.16,162.13,56.12,161.7,55.29,162.12,54.86,160.37,54.34,160.02,53.2,158.53,52.96,158.23,51.94,156.79,51.01,155.43,55.38,155.91,56.77,156.81,57.83,158.36,58.06,163.67,61.14,164.47,62.55,163.26,62.47,162.66,61.64,160.12,60.54,159.3,61.77,156.72,61.43,154.22,59.76,155.04,59.14,151.27,58.78,151.34,59.5,149.78,59.66,148.54,59.16,142.2,59.04,135.13,54.73,136.7,54.6,138.16,53.76,139.9,54.19,141.35,53.09,140.06,48.45,134.87,43.4,133.54,42.81,132.28,43.28,130.78,42.22,131.03,44.97,133.1,45.14,135.03,48.48,130.99,47.79,130.58,48.73,129.4,49.44,127.66,49.76,125.95,52.79,125.07,53.16,121.0,53.25,120.18,52.75,120.74,51.96,119.29,50.14,117.88,49.51,114.36,50.25,110.66,49.13,108.48,49.28,106.89,50.27,103.68,50.09,102.26,50.51,102.07,51.26,98.86,52.05,97.83,51.01,98.23,50.42,97.26,49.73,92.23,50.8,87.36,49.21,83.38,51.07,81.95,50.81,80.57,51.39,80.04,50.86,76.53,54.18,76.89,54.49,73.43,53.49,73.51,54.04,71.18,54.13,70.87,55.17,69.07,55.39,61.44,54.01,60.98,53.66,61.7,52.98,59.97,51.96,61.59,51.27,61.34,50.8,56.78,51.04,55.72,50.62,52.33,51.72,50.77,51.69,48.7,50.61,48.58,49.87,47.55,50.45,46.47,48.39,48.06,47.74,49.1,46.4,46.68,44.61,48.58,41.81,47.82,41.15,45.47,42.5,39.96,43.43,36.68,45.24,38.23,46.24,37.67,46.64,39.15,47.04,38.22,47.1,38.26,47.55,39.74,47.9,40.07,49.6,35.36,50.58,35.02,51.21,34.22,51.26,34.39,51.77,33.75,52.34,31.79,52.1,31.31,53.07,32.69,53.35,30.76,54.81,30.87,55.55,28.18,56.17,27.29,57.47,27.72,57.79,27.42,58.72,29.12,60.03,28.07,60.5,31.52,62.87,30.04,63.55,30.44,64.2,29.54,64.95,30.22,65.81,29.05,66.94,29.98,67.7,28.45,68.36,28.59,69.06,32.13,69.91,41.06,67.46,41.13,66.79,38.38,66.0,33.18,66.63,34.81,65.9,34.94,64.41,37.01,63.85,36.54,64.76,37.18,65.14,39.59,64.52,40.44,64.76,39.76,65.5,42.09,66.48,43.95,66.07,44.53,66.76,43.7,67.35,44.19,67.95,43.45,68.57,46.25,68.25,46.82,67.69,45.56,67.57,45.56,67.01,46.35,66.67,53.72,68.86,54.47,68.81,53.49,68.2,58.8,68.88,59.94,68.28,61.08,68.94,60.03,69.52,60.55,69.85,68.51,68.09,69.18,68.62,66.93,69.45,67.26,69.93,66.69,71.03,69.94,73.04,72.59,72.78,72.8,72.22,71.85,71.41,72.79,70.39,72.56,69.02,73.67,68.41,71.28,66.32,72.42,66.17,75.05,67.76,74.47,68.33,74.94,68.99,73.84,69.07,73.6,69.63,74.4,70.63,73.1,71.45,74.89,72.12,74.66,72.83,75.68,72.3,75.29,71.34,76.36,71.15,75.9,71.87,77.58,72.27,81.5,71.75,80.61,72.58,80.51,73.65,86.82,73.94,86.01,74.46,87.17,75.12,100.76,76.43,101.99,77.29,104.35,77.7,106.07,77.37,104.7,77.13,106.97,76.97],[105.08,78.31,99.44,77.92,102.09,79.35,105.37,78.71,105.08,78.31],[51.14,80.55,47.59,80.01,46.5,80.25,47.07,80.56,44.85,80.59,51.52,80.7,51.14,80.55],[99.94,78.88,94.97,79.04,91.18,80.34,95.94,81.25,100.19,79.78,99.94,78.88]],"Rwanda":[[30.42,-1.13,30.76,-2.29,29.02,-2.84,29.29,-1.62,30.42,-1.13]],"Western Sahara":[[-8.79,27.12,-8.67,27.66,-8.69,25.88,-11.97,25.93,-11.94,23.37,-12.87,23.28,-12.93,21.33,-17.06,21.0,-17.02,21.42,-14.75,21.5,-13.89,23.69,-12.5,24.77,-11.39,26.88,-8.79,27.12]],"Saudi Arabia":[[42.78,16.35,40.94,19.49,39.14,21.29,38.49,23.69,37.48,24.29,35.13,28.06,34.63,28.06,34.96,29.36,36.07,29.2,38.0,30.51,37.0,31.51,39.2,32.16,40.4,31.89,44.71,29.18,48.42,28.55,50.15,26.69,50.24,25.61,51.39,24.63,52.0,23.0,55.21,22.71,55.67,22.0,55.0,20.0,49.12,18.62,47.0,16.95,43.38,17.58,42.78,16.35]],"Sudan":[[33.96,9.46,33.21,10.72,33.21,12.18,32.74,12.25,32.07,11.97,32.4,11.08,31.35,9.81,30.0,10.29,28.97,9.4,26.75,9.47,25.79,10.41,25.07,10.27,24.54,8.92,23.46,8.95,23.55,10.09,21.94,12.59,23.02,15.68,23.89,15.61,23.85,20.0,25.0,20.0,25.0,22.0,36.87,22.0,37.48,18.61,38.41,18.0,36.85,16.96,36.27,13.56,33.96,9.46]],"South Sudan":[[33.96,9.46,33.83,8.38,32.95,7.78,34.08,7.23,35.3,5.51,33.39,3.79,31.88,3.56,30.83,3.51,29.72,4.6,27.98,4.41,23.89,8.62,24.54,8.92,25.07,10.27,25.79,10.41,26.75,9.47,28.97,9.4,30.0,10.29,31.35,9.81,32.4,11.08,32.07,11.97,32.74,12.25,33.21,12.18,33.21,10.72,33.96,9.46]],"Senegal":[[-16.71,13.59,-17.63,14.73,-16.12,16.46,-14.58,16.6,-12.17,14.62,-11.51,12.44,-16.68,12.38,-16.84,13.15,-13.84,13.51,-16.71,13.59]],"Sierra Leone":[[-11.44,6.79,-12.95,7.8,-13.25,8.9,-11.92,10.05,-11.12,10.05,-10.23,8.41,-11.44,6.79]],"El Salvador":[[-87.79,13.38,-90.1,13.74,-89.35,14.42,-87.86,13.89,-87.79,13.38]],"Somaliland":[[48.94,9.45,47.79,8.0,46.95,8.0,43.68,9.18,42.56,10.57,43.15,11.46,44.12,10.45,48.95,11.41,48.94,9.45]],"Somalia":[[49.73,11.58,51.11,12.02,51.05,10.64,48.59,5.34,41.59,-1.68,40.99,-0.86,40.98,2.78,42.13,4.23,44.96,5.0,48.94,9.45,48.94,11.39,49.73,11.58]],"Republic of Serbia":[[20.87,45.42,22.71,44.58,22.41,44.01,22.99,43.21,22.38,42.32,21.58,42.25,21.78,42.68,20.81,43.27,20.26,42.81,19.22,43.52,19.6,44.04,18.83,45.91,20.22,46.13,20.87,45.42]],"Suriname":[[-57.15,5.97,-53.96,5.76,-54.48,4.9,-54.01,3.62,-54.52,2.31,-55.97,2.51,-56.0,1.82,-56.54,1.9,-57.6,3.33,-58.04,4.06,-57.15,5.97]],"Slovakia":[[18.85,49.5,22.56,49.09,21.87,48.32,20.8,48.62,17.86,47.76,16.98,48.12,17.1,48.82,18.85,49.5]],"Slovenia":[[13.81,46.51,16.56,46.5,15.33,45.45,13.72,45.5,13.81,46.51]],"Sweden":[[22.18,65.72,21.21,65.03,21.37,64.41,17.85,62.75,17.12,61.34,18.79,60.08,17.87,58.95,16.83,58.72,15.88,56.1,14.67,56.2,14.1,55.41,12.94,55.36,11.03,58.86,12.3,60.12,12.63,61.29,11.99,61.8,11.93,63.13,12.58,64.07,13.57,64.05,13.92,64.45,13.56,64.79,16.77,68.01,17.73,68.01,17.99,68.57,19.88,68.41,20.03,69.07,23.54,67.94,23.9,66.01,22.18,65.72]],"Swaziland":[[32.07,-26.73,31.28,-27.29,30.69,-26.74,31.04,-25.73,31.84,-25.84,32.07,-26.73]],"Syria":[[38.79,33.38,36.83,32.31,35.7,32.72,36.61,34.2,35.91,35.41,36.74,36.82,42.35,37.23,41.29,36.36,41.01,34.42,38.79,33.38]],"Chad":[[14.5,12.86,13.54,14.37,13.97,15.68,15.25,16.63,15.9,20.39,15.1,21.31,14.85,22.86,15.86,23.41,23.84,19.58,23.89,15.61,23.02,15.68,21.94,12.59,22.86,11.14,21.0,9.48,18.81,8.98,17.96,7.89,15.28,7.42,14.98,8.8,13.95,9.55,14.17,10.02,15.47,9.98,14.5,12.86]],"Togo":[[1.87,6.14,1.06,5.93,0.57,6.91,0.02,11.02,0.9,11.0,1.43,9.83,1.87,6.14]],"Thailand":[[102.58,12.19,100.83,12.63,100.98,13.41,100.1,13.41,99.15,9.96,100.46,7.43,102.14,6.22,101.15,5.69,98.15,8.35,99.59,11.89,98.19,15.12,98.9,16.18,97.38,18.45,98.25,19.71,100.12,20.42,100.61,19.51,101.28,19.46,101.06,17.51,103.2,18.31,104.72,17.43,105.54,14.72,105.22,14.27,102.99,14.23,102.35,13.39,102.58,12.19]],"Tajikistan":[[71.01,40.24,69.56,40.1,69.46,39.53,73.68,39.43,73.93,38.51,74.86,38.38,74.98,37.42,73.26,37.5,71.84,36.74,70.81,38.49,69.2,37.15,67.83,37.14,68.39,38.16,68.18,38.9,67.44,39.14,67.7,39.58,68.54,39.53,69.33,40.73,70.67,40.96,70.46,40.5,71.01,40.24]],"Turkmenistan":[[61.21,35.65,61.12,36.49,57.33,38.03,55.51,37.96,53.92,37.2,53.88,38.95,53.1,39.29,53.36,39.98,52.69,40.03,52.92,40.88,54.74,40.95,53.72,42.12,52.92,41.87,52.81,41.14,52.5,41.78,54.08,42.32,55.46,41.26,57.1,41.32,56.93,41.83,58.63,42.75,59.98,42.22,60.47,41.22,61.88,41.08,62.37,40.05,64.17,38.89,66.55,37.97,66.52,37.36,65.75,37.66,64.75,37.11,64.55,36.31,62.98,35.4,61.21,35.65]],"East Timor":[[124.97,-8.89,127.34,-8.4,125.09,-9.39,124.97,-8.89]],"Trinidad and Tobago":[[-61.68,10.76,-60.9,10.86,-60.94,10.11,-61.95,10.09,-61.68,10.76]],"Tunisia":[[9.48,30.31,9.06,32.1,7.61,33.34,7.52,34.1,8.14,34.66,8.42,36.95,9.51,37.35,10.21,37.23,10.18,36.72,11.03,37.09,10.6,36.41,10.81,34.83,10.15,34.33,11.49,33.14,11.43,32.37,9.95,31.38,9.97,30.54,9.48,30.31]],"Turkey":[[36.91,41.34,40.37,41.01,42.62,41.58,43.58,41.09,43.66,40.25,44.79,39.71,44.11,39.43,44.77,37.17,36.74,36.82,36.15,35.82,35.78,36.27,36.16,36.65,34.71,36.8,34.03,36.22,30.62,36.68,29.7,36.14,27.64,36.66,26.32,38.21,26.8,38.99,26.17,39.46,27.28,40.42,28.82,40.46,29.24,41.22,31.15,41.09,33.51,42.02,36.91,41.34],[27.19,40.69,26.36,40.15,26.06,40.82,26.6,41.56,26.12,41.83,28.0,42.01,28.99,41.3,27.19,40.69]],"Taiwan":[[121.78,24.39,120.75,21.97,120.11,23.56,121.5,25.3,121.78,24.39]],"United Republic of Tanzania":[[33.9,-0.95,37.7,-3.1,37.77,-3.68,39.2,-4.68,38.74,-5.91,39.44,-6.84,39.19,-8.49,40.32,-10.32,39.52,-10.9,36.51,-11.72,34.56,-11.52,33.74,-9.42,30.74,-8.34,29.62,-6.52,29.34,-4.5,30.75,-3.36,30.42,-1.13,33.9,-0.95]],"Uganda":[[31.87,-1.03,29.58,-1.34,29.88,0.6,31.17,2.2,30.77,2.34,30.83,3.51,34.01,4.25,35.04,1.91,33.89,0.11,33.9,-0.95,31.87,-1.03]],"Ukraine":[[31.79,52.1,33.75,52.34,34.39,51.77,34.22,51.26,35.02,51.21,35.36,50.58,40.07,49.6,39.74,47.9,34.96,46.27,35.02,45.65,36.53,45.47,36.33,45.11,33.88,44.36,33.33,44.56,33.55,45.03,32.45,45.33,33.59,45.85,31.68,46.71,30.75,46.58,29.6,45.29,28.23,45.49,28.86,46.44,30.02,46.42,28.67,48.12,27.52,48.47,24.87,47.74,22.71,47.88,22.09,48.42,22.78,49.03,22.52,49.48,23.92,50.42,23.53,51.58,25.33,51.91,30.56,51.32,30.93,52.04,31.79,52.1]],"Uruguay":[[-57.63,-30.22,-56.98,-30.11,-53.79,-32.05,-53.21,-32.73,-53.81,-34.4,-56.22,-34.86,-58.43,-33.91,-57.63,-30.22]],"United States of America":[[-155.54,19.08,-155.94,19.06,-155.86,20.27,-154.81,19.51,-155.54,19.08],[-94.82,49.39,-94.33,48.67,-91.64,48.14,-88.38,48.3,-84.14,46.51,-82.55,45.35,-82.14,43.57,-83.12,42.08,-82.69,41.68,-78.94,42.86,-79.17,43.47,-78.72,43.63,-76.82,43.63,-74.87,45.0,-71.51,45.01,-69.24,47.45,-67.79,47.07,-67.79,45.7,-66.96,44.81,-70.12,43.68,-70.83,42.34,-69.97,41.64,-73.71,40.93,-71.94,40.93,-73.95,40.75,-74.91,38.94,-75.53,39.5,-75.06,38.4,-75.94,37.22,-75.72,37.94,-76.35,39.15,-76.33,38.08,-76.99,38.24,-76.3,37.92,-75.73,35.55,-81.34,31.44,-81.31,30.04,-80.06,26.88,-80.68,25.08,-82.71,27.5,-82.93,29.1,-84.1,30.09,-85.11,29.64,-86.4,30.4,-89.18,30.32,-89.59,30.16,-89.41,29.16,-93.23,29.78,-94.69,29.48,-97.14,27.83,-97.14,25.87,-97.53,25.84,-99.02,26.37,-100.96,29.38,-102.48,29.76,-103.11,28.97,-103.94,29.27,-106.51,31.75,-111.02,31.33,-114.72,32.72,-117.13,32.54,-118.52,34.03,-120.62,34.61,-124.4,40.31,-124.53,42.77,-123.9,45.52,-124.69,48.18,-123.12,48.04,-122.59,47.1,-122.84,49.0,-95.16,49.0,-94.82,49.39],[-153.01,57.12,-154.01,56.73,-154.67,57.46,-153.23,57.97,-152.14,57.59,-153.01,57.12],[-165.58,59.91,-167.46,60.21,-165.67,60.29,-165.58,59.91],[-171.73,63.78,-168.69,63.3,-169.53,62.98,-171.55,63.32,-171.73,63.78],[-155.07,71.15,-154.34,70.7,-140.99,69.71,-141.0,60.31,-139.04,60.0,-137.45,58.91,-135.48,59.79,-131.71,56.55,-130.01,55.92,-129.98,55.28,-130.54,54.8,-131.97,55.5,-134.08,58.12,-136.63,58.21,-142.57,60.08,-147.11,60.88,-148.22,60.67,-148.02,59.98,-151.72,59.16,-151.41,60.73,-150.35,61.03,-150.62,61.28,-154.02,59.35,-153.29,58.86,-154.23,58.15,-158.43,55.99,-164.79,54.4,-157.72,57.57,-157.04,58.92,-159.06,58.42,-160.36,59.07,-161.97,58.67,-161.87,59.63,-162.52,59.99,-163.82,59.8,-165.35,60.51,-166.12,61.5,-165.73,62.07,-164.56,63.15,-160.77,63.77,-161.52,64.4,-160.78,64.79,-164.96,64.45,-168.11,65.67,-164.47,66.58,-163.65,66.58,-163.79,66.08,-161.68,66.12,-166.76,68.36,-166.2,68.88,-163.17,69.37,-161.91,70.33,-156.58,71.36,-155.07,71.15]],"Uzbekistan":[[66.52,37.36,66.55,37.97,64.17,38.89,62.37,40.05,61.88,41.08,60.47,41.22,59.98,42.22,58.63,42.75,56.93,41.83,57.1,41.32,55.97,41.31,55.93,45.0,58.5,45.59,62.01,43.5,64.9,43.73,66.1,43.0,66.02,41.99,66.51,41.99,66.71,41.17,67.99,41.14,68.26,40.66,70.96,42.27,70.42,41.52,73.06,40.87,71.77,40.15,70.6,40.22,70.67,40.96,69.33,40.73,68.54,39.53,67.44,39.14,68.18,38.9,68.39,38.16,67.83,37.14,66.52,37.36]],"Venezuela":[[-71.33,11.78,-71.95,11.42,-71.7,9.07,-71.04,9.86,-71.4,10.97,-70.16,11.38,-69.94,12.16,-68.19,10.55,-66.23,10.65,-64.89,10.08,-64.32,10.64,-61.88,10.72,-62.73,10.42,-59.76,8.37,-61.41,5.96,-60.6,4.92,-63.09,3.77,-64.82,4.06,-64.27,2.5,-63.37,2.2,-66.33,0.72,-67.81,2.82,-67.3,3.32,-67.82,4.5,-67.34,6.1,-69.39,6.1,-70.09,6.96,-71.96,6.99,-72.79,9.09,-73.3,9.15,-72.91,10.45,-71.33,11.78]],"Vietnam":[[108.05,21.55,106.72,20.7,105.66,19.06,108.88,15.28,109.2,11.67,105.16,8.6,105.08,9.92,104.33,10.49,106.25,10.96,105.81,11.57,107.49,12.34,107.56,15.2,105.09,18.67,103.9,19.27,104.82,19.89,104.44,20.76,103.2,20.77,102.17,22.46,105.33,23.35,108.05,21.55]],"Yemen":[[53.11,16.65,52.39,16.38,52.17,15.6,48.68,14.0,44.18,12.59,43.48,12.64,42.6,15.21,43.38,17.58,47.0,16.95,49.12,18.62,52.0,19.0,53.11,16.65]],"South Africa":[[31.52,-29.26,27.46,-33.23,25.78,-33.94,22.57,-33.86,19.62,-34.82,18.38,-34.14,17.93,-32.61,18.22,-31.66,16.34,-28.58,16.82,-28.08,18.46,-29.05,19.89,-28.46,19.9,-24.77,20.89,-26.83,21.61,-26.73,23.31,-25.27,25.66,-25.49,28.02,-22.83,29.84,-22.1,31.19,-22.25,31.93,-24.37,31.84,-25.84,31.04,-25.73,30.69,-26.74,31.28,-27.29,32.83,-26.74,32.46,-28.3,31.52,-29.26]],"Zambia":[[32.76,-9.23,33.49,-10.53,32.69,-13.71,33.21,-13.97,30.18,-14.8,30.27,-15.51,28.95,-16.04,27.04,-17.94,23.22,-17.52,21.89,-16.08,21.93,-12.9,24.02,-12.91,23.91,-10.93,25.75,-11.78,27.16,-11.61,28.93,-13.25,29.7,-13.26,29.62,-12.18,28.37,-11.79,28.73,-8.53,30.35,-8.24,32.76,-9.23]],"Zimbabwe":[[31.19,-22.25,28.02,-21.49,25.26,-17.74,27.04,-17.94,28.95,-16.04,30.27,-15.51,32.85,-16.71,32.66,-20.3,31.19,-22.25]]};

/* ---------------------------------------------------------------
   16d. BOARD BACKGROUNDS
   ---------------------------------------------------------------
   Pick the canvas colour, the card colour, the text colour and the
   default chart palette, or take one of the presets.
   --------------------------------------------------------------- */

const BOARD_PRESETS = [
  { id: 'paper',  name: 'Paper',       bg: '#F6F1E4', card: '#FFFDF8', fg: '#241C14', grid: '#E4DBC6', palette: 'rust' },
  { id: 'white',  name: 'Clean white', bg: '#FFFFFF', card: '#FFFFFF', fg: '#1F2933', grid: '#E3E6EA', palette: 'ocean' },
  { id: 'grey',   name: 'Soft grey',   bg: '#EEF1F4', card: '#FFFFFF', fg: '#1F2933', grid: '#D8DEE4', palette: 'ocean' },
  { id: 'navy',   name: 'Dark navy',   bg: '#121B2A', card: '#1B2739', fg: '#E7EDF3', grid: '#2A3A50', palette: 'neon' },
  { id: 'char',   name: 'Charcoal',    bg: '#1C1C1E', card: '#2A2A2D', fg: '#EDEDEF', grid: '#3A3A3E', palette: 'neon' },
  { id: 'plum',   name: 'Deep plum',   bg: '#1A0F2E', card: '#251643', fg: '#EDE6FA', grid: '#3A2560', palette: 'aurora' },
  { id: 'forest', name: 'Deep green',  bg: '#10231C', card: '#183129', fg: '#E4F0EA', grid: '#27453A', palette: 'forest' }
];

const BoardTheme = { bg: '#F6F1E4', card: '#FFFDF8', fg: '#241C14', grid: '#E4DBC6',
                     preset: 'paper', palette: 'rust' };

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

/* ---- background settings UI ---- */
function renderBoardBackgroundSettings(wrap) {
  const html =
    '<h3 class="snap-set-title">Chart background</h3>' +
    '<p class="drill-subtitle">Applies to the Dashboard and the Product Performance charts. ' +
      'A preset sets the canvas, the cards, the text and the default chart colours together; ' +
      'change any one of them below and the rest stay put.</p>' +
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
    '</div>' +
    '<h3 class="snap-set-title">Default chart colours</h3>' +
    '<p class="drill-subtitle">Every chart set to "Follow the board" uses this palette. ' +
      'A chart with its own palette or its own picked colours keeps them.</p>' +
    '<div class="pal-grid">' +
      CHART_PALETTES.map(p =>
        '<button class="pal-btn' + ((BoardTheme.palette || 'rust') === p.id ? ' active' : '') + '" data-pal="' + p.id + '">' +
          '<span class="pal-strip">' + p.colors.slice(0, 8).map(c =>
            '<i style="background:' + c + '"></i>').join('') + '</span>' + p.name +
        '</button>').join('') +
    '</div>';

  const holder = document.createElement('div');
  holder.innerHTML = html;
  wrap.appendChild(holder);

  holder.querySelectorAll('.bg-preset').forEach(b => b.addEventListener('click', () => {
    const p = BOARD_PRESETS.find(x => x.id === b.dataset.preset);
    if (!p) return;
    BoardTheme.bg = p.bg; BoardTheme.card = p.card; BoardTheme.fg = p.fg;
    BoardTheme.grid = p.grid; BoardTheme.preset = p.id;
    if (p.palette) BoardTheme.palette = p.palette;
    saveBoardTheme();
    renderSettingsBody();
  }));
  holder.querySelectorAll('.pal-btn').forEach(b => b.addEventListener('click', () => {
    BoardTheme.palette = b.dataset.pal;
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
const BUILD_VERSION = 'v47';

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
  // Without this the manual ADC / LT / SF / MOQ / MIT values were written to
  // storage on every edit and then never read again, so they vanished the
  // moment the window was closed.
  safeInit('replen', loadReplen);
  safeInit('catalog', initCatalog);
  safeInit('boards', initBoards);
  safeInit('board-theme', loadBoardTheme);
  safeInit('prefs-apply', applyPrefsToControls);
  safeInit('snapshot', initSnapshot);
  safeInit('session', initSession);
  safeInit('period-selects', wirePeriodSelects);
  safeInit('gs-buttons', updateGsOnlyButtons);
  safeInit('save-button', initSaveButton);
  safeInit('dashboard-render', renderDashboard);
  safeInit('performance-render', renderPerformance);
  safeInit('relations-render', renderRelations);
  safeInit('restore-persisted', restorePersistedDatasets);
});

})();
