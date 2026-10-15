/**
 * SPINE – Google Apps Script backend, connected DIRECTLY to your spreadsheet
 *   SPINE ORDER REVIEW 2026  (ID below)
 *
 *  GET  ?action=all&token=...                      -> every tab as JSON
 *  POST {token, action:"import", tab, mode, rows}  -> write mapped rows into ONE source tab
 *  POST {token, action:"saveOrder", items, vars}   -> write Qty to be ordered + Status into "Transaction Data"
 *
 * Every write is recorded in a "SPINE_Log" tab that this script creates.
 *
 * SETUP (once): paste this file at  script.google.com  (or Sheet > Extensions > Apps Script)
 *   1. Run the function  check  (authorise when asked) – the log shows each tab and its row count
 *   2. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone > Deploy
 *   3. Copy the Web app URL into VITE_API_URL in the .env file
 * Deploy it from the Google account that can EDIT the spreadsheet.
 */
const SHEET_ID  = "YOUR_GOOGLE_SHEET_ID";
const API_TOKEN = "YOUR_SECRET_TOKEN";   // must equal VITE_API_TOKEN in .env  (change both to rotate)
// id -> sheet name, first DATA row (below the header), number of data columns, key column (0-based)
const CFG = {
  TD:  { sheet: 'Transaction Data', start: 2, cols: 5,  key: 3, readOnly: true },
  PO:  { sheet: 'PENDING ORDER',    start: 4, cols: 7,  key: 2 },   // header is on row 3
  MWH: { sheet: 'MWH',              start: 2, cols: 4,  key: 0 },
  DWH: { sheet: 'DWH',              start: 2, cols: 6,  key: 0 },
  BWH: { sheet: 'BWH',              start: 2, cols: 6,  key: 1 },
  LWH: { sheet: 'LWH',              start: 2, cols: 6,  key: 0 },
  UP:  { sheet: 'up usage',         start: 2, cols: 12, key: 8 },
  DEL: { sheet: 'del usage',        start: 2, cols: 10, key: 6 },
  OUT: { sheet: 'outstation',       start: 4, cols: 11, key: 6 },   // header is on row 3
  BLR: { sheet: 'blr usage',        start: 2, cols: 12, key: 8 }
};

function doGet(e) {
  try {
    check_(e.parameter.token);
    if ((e.parameter.action || 'all') !== 'all') return json_({ error: 'Unknown action' });
    const out = {};
    Object.keys(CFG).forEach(id => { out[id] = readTab_(CFG[id]); });
    return json_(out);
  } catch (err) { return json_({ error: String(err.message || err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse(e.postData.contents);
    check_(body.token);
    if (body.action === 'saveOrder') return json_(saveOrder_(body));
    if (body.action !== 'import') throw new Error('Unknown action');
    const cfg = CFG[body.tab];
    if (!cfg) throw new Error('Unknown tab: ' + body.tab);
    if (cfg.readOnly) throw new Error(cfg.sheet + ' is export-only');
    const rows = (body.rows || []).map(r => pad_(r, cfg.cols));
    if (!rows.length) throw new Error('No rows to import');
    const mode = body.mode === 'append' ? 'append' : 'replace';
    writeTab_(cfg, rows, mode);
    log_('import (' + mode + ')', cfg.sheet, rows.length, '');
    return json_({ ok: true, rows: rows.length });
  } catch (err) { return json_({ error: String(err.message || err) }); }
  finally { lock.releaseLock(); }
}

// ---------- helpers ----------
function check_(token) {
  if (!API_TOKEN || token !== API_TOKEN) throw new Error('Unauthorized');
}

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet_(cfg) {
  const sh = ss_().getSheetByName(cfg.sheet);
  if (!sh) throw new Error('Tab not found: ' + cfg.sheet);
  return sh;
}

function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof v === 'string') { v = v.replace(/\u00a0/g, ' ').trim(); return v === '' ? null : v; }
  return v === '' ? null : v;
}

function readTab_(cfg) {
  const sh = sheet_(cfg);
  const last = sh.getLastRow();
  if (last < cfg.start) return [];
  const vals = sh.getRange(cfg.start, 1, last - cfg.start + 1, cfg.cols).getValues();
  const rows = [];
  vals.forEach(r => {
    const c = r.map(clean_);
    if (!c.some(v => v !== null)) return;
    if (cfg.readOnly) { if (c[cfg.key] === null) return; rows.push(c.map(v => v === null ? '' : v)); }
    else rows.push(c);
  });
  return rows;
}

function pad_(r, n) { const o = []; for (let i = 0; i < n; i++) o.push(r[i] === undefined ? '' : (r[i] === null ? '' : r[i])); return o; }

function lastDataRow_(sh, cfg) {
  const last = sh.getLastRow();
  if (last < cfg.start) return cfg.start - 1;
  const keys = sh.getRange(cfg.start, cfg.key + 1, last - cfg.start + 1, 1).getValues();
  for (let i = keys.length - 1; i >= 0; i--) if (keys[i][0] !== '' && keys[i][0] !== null) return cfg.start + i;
  return cfg.start - 1;
}

function writeTab_(cfg, rows, mode) {
  const sh = sheet_(cfg);
  const chk = cfg.cols + 1;                                   // the CHECK column (formulas) sits right after the data
  const oldLast = Math.max(lastDataRow_(sh, cfg), cfg.start);
  const chkFormula = sh.getRange(cfg.start, chk).getFormulaR1C1();
  let first;
  if (mode === 'replace') {
    sh.getRange(cfg.start, 1, Math.max(sh.getMaxRows() - cfg.start + 1, 1), cfg.cols).clearContent();
    if (chkFormula) sh.getRange(cfg.start, chk, Math.max(oldLast - cfg.start + 1, rows.length)).clearContent();
    first = cfg.start;
  } else {
    first = lastDataRow_(sh, cfg) + 1;
  }
  const need = first + rows.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(first, 1, rows.length, cfg.cols).setValues(rows);
  if (chkFormula) sh.getRange(first, chk, rows.length, 1).setFormulaR1C1(chkFormula);   // keep CHECK formulas on every new row
  SpreadsheetApp.flush();
}


// ---------- write back: Qty to be ordered + Status -> Transaction Data ----------
function saveOrder_(body) {
  const items = body.items || [];
  if (!items.length) throw new Error('Nothing to save');
  const sh = ss_().getSheetByName(CFG.TD.sheet);
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).replace(/\s+/g, ' ').trim().toLowerCase());
  const cCat = hdr.indexOf('cat. no.') + 1, cQty = hdr.indexOf('qty to be ordered') + 1, cSt = hdr.indexOf('status') + 1;
  if (!cCat || !cQty || !cSt) throw new Error('Headers "Cat. No.", "Qty to be ordered", "Status" not found in row 1 of Transaction Data');
  const n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('Transaction Data is empty');
  const cats = sh.getRange(2, cCat, n, 1).getValues();
  const qR = sh.getRange(2, cQty, n, 1), sR = sh.getRange(2, cSt, n, 1);
  const qV = qR.getValues(), qF = qR.getFormulas(), sV = sR.getValues(), sF = sR.getFormulas();
  const map = {};
  items.forEach(it => { map[String(it.code).replace(/ /g, ' ').trim().toUpperCase()] = it; });
  const outQ = [], outS = [], before = [];
  let changed = 0;
  for (let i = 0; i < n; i++) {
    const it = map[String(cats[i][0]).replace(/ /g, ' ').trim().toUpperCase()];
    if (it) {
      before.push([cats[i][0], qF[i][0] || qV[i][0], sF[i][0] || sV[i][0]]);
      outQ.push([Number(it.qty) || 0]); outS.push([String(it.status || '')]); changed++;
    } else {            // untouched rows keep their existing formula / value
      outQ.push([qF[i][0] || qV[i][0]]); outS.push([sF[i][0] || sV[i][0]]);
    }
  }
  qR.setValues(outQ); sR.setValues(outS);
  SpreadsheetApp.flush();
  const v = body.vars || {};
  log_('saveOrder', CFG.TD.sheet, changed,
       'consumption = usage / ' + v.dataMonths + ' x ' + v.months + '; shortfall basis ' + v.basis + '. Previous [cat, qty, status]: ' + JSON.stringify(before).slice(0, 45000));
  return { ok: true, rows: changed };
}

function log_(action, tab, rows, details) {
  const ss = ss_();
  let sh = ss.getSheetByName('SPINE_Log');
  if (!sh) { sh = ss.insertSheet('SPINE_Log'); sh.appendRow(['Timestamp', 'User', 'Action', 'Tab', 'Rows', 'Details']); sh.setFrozenRows(1); }
  let user = ''; try { user = Session.getEffectiveUser().getEmail(); } catch (e) {}
  sh.appendRow([new Date(), user, action, tab, rows, details]);
}

// ---------- run this once from the editor to authorise and verify ----------
function check() {
  Object.keys(CFG).forEach(id => {
    const c = CFG[id];
    try {
      const rows = readTab_(c);
      Logger.log(id + '  "' + c.sheet + '"  rows: ' + rows.length + '  first: ' + JSON.stringify(rows[0]));
    } catch (e) { Logger.log(id + '  "' + c.sheet + '"  PROBLEM: ' + e.message); }
  });
}
