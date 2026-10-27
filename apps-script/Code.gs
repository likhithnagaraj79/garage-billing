/**
 * Sri Manjunatha Garage — billing API.
 * Paste this into the Google Sheet: Extensions → Apps Script → Code.gs.
 * Script properties (Project Settings → Script properties):
 *   PIN            – required, the owner's PIN typed in the billing page
 *   START_BILL_NO  – optional, first bill number when the sheet is empty (e.g. 130)
 *
 * Sheet layout (created automatically):
 *   One tab per month, e.g. "Oct-2026" — one row per item.
 *   "Parts" tab — every particular used, with its last rate.
 */

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_TAB = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{4})$/;
const HEADERS = ['Bill No', 'Date', 'M/s', 'Vehicle No', 'Customer Phone', 'Sl.No', 'Particulars',
  'Qty', 'Rate', 'Amount', 'Bill Total', 'Amount in Words', 'Saved On'];
const C = {}; HEADERS.forEach((h, i) => C[h] = i);   // column index by header
const PARTS_TAB = 'Parts';
const PARTS_HEADERS = ['Particulars', 'Last Rate', 'Times Used'];

function doGet() {
  return ContentService.createTextOutput('Garage billing API is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const pin = PropertiesService.getScriptProperties().getProperty('PIN');
    if (!pin) throw new Error('PIN is not set in Apps Script → Project Settings → Script properties.');
    if (String(req.pin || '') !== String(pin)) {
      out = { ok: false, error: 'Wrong PIN', badPin: true };
    } else {
      const handler = { init, save, search, get: getBill }[req.action];
      if (!handler) throw new Error('Unknown action: ' + req.action);
      out = Object.assign({ ok: true }, handler(req));
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- actions ---------- */

function init() {
  return { nextBillNo: nextBillNo(), parts: readParts() };
}

function save(req) {
  const bill = req.bill, lines = req.lines;
  const no = Number(bill.billNo);
  if (!no || !lines || !lines.length) throw new Error('Bill No. and at least one item are required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const drop = [no];
    if (req.replaceNo) drop.push(Number(req.replaceNo));
    monthSheets().forEach(sh => deleteBillRows(sh, drop));

    const sh = monthSheet(bill.date, true);
    const savedOn = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const rows = lines.map((l, i) => [
      no, bill.date, bill.ms || '', bill.vehicleNo || '', bill.phone || '', i + 1, l.particulars,
      Number(l.qty), Number(l.rate), Number(l.amount),
      i === 0 ? Number(bill.total) : '', i === 0 ? bill.words || '' : '', savedOn,
    ]);
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    sortByBillNo(sh);
    updateParts(lines);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { nextBillNo: nextBillNo() };
}

function search(req) {
  const q = String(req.q || '').trim().toLowerCase();
  const from = req.from || '', to = req.to || '';   // ISO yyyy-mm-dd
  const bills = {};
  monthSheets().forEach(sh => {
    const key = tabKey(sh.getName());
    if (from && key < from.slice(0, 7)) return;
    if (to && key > to.slice(0, 7)) return;
    dataRows(sh).forEach(r => {
      const no = Number(r[C['Bill No']]);
      if (!no) return;
      const b = bills[no] || (bills[no] = {
        billNo: no, date: dmy(r[C['Date']]), ms: r[C['M/s']], vehicleNo: r[C['Vehicle No']],
        phone: String(r[C['Customer Phone']]), total: 0,
      });
      b.total += Number(r[C['Amount']]) || 0;
    });
  });
  const list = Object.keys(bills).map(k => bills[k]).filter(b => {
    const iso = toISO(b.date);
    if (from && iso < from) return false;
    if (to && iso > to) return false;
    if (!q) return true;
    return [b.billNo, b.vehicleNo, b.phone, b.ms].some(v => String(v).toLowerCase().indexOf(q) >= 0);
  }).sort((a, b) => b.billNo - a.billNo);
  return { bills: list.slice(0, 200) };
}

function getBill(req) {
  const no = Number(req.no);
  let bill = null;
  const lines = [];
  monthSheets().forEach(sh => dataRows(sh).forEach(r => {
    if (Number(r[C['Bill No']]) !== no) return;
    if (!bill) bill = { billNo: no, date: dmy(r[C['Date']]), ms: r[C['M/s']], vehicleNo: r[C['Vehicle No']], phone: String(r[C['Customer Phone']]) };
    lines.push({ sl: Number(r[C['Sl.No']]), particulars: r[C['Particulars']], qty: Number(r[C['Qty']]), rate: Number(r[C['Rate']]) });
  }));
  if (!bill) throw new Error('Bill No. ' + no + ' not found.');
  lines.sort((a, b) => a.sl - b.sl);
  return { bill: bill, lines: lines };
}

/* ---------- month tabs ---------- */

function tabKey(name) {   // "Oct-2026" -> "2026-10" (sortable)
  const m = MONTH_TAB.exec(name);
  return m ? m[2] + '-' + String(MONTHS.indexOf(m[1]) + 1).padStart(2, '0') : '';
}

function monthSheets() {
  return SpreadsheetApp.getActive().getSheets().filter(sh => MONTH_TAB.test(sh.getName()));
}

function monthSheet(dateDMY, create) {
  const p = String(dateDMY).split('/');
  const name = MONTHS[Number(p[1]) - 1] + '-' + p[2];
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh || !create) return sh;

  // insert before the first later month so tabs stay in date order
  const key = tabKey(name);
  const later = monthSheets().filter(s => tabKey(s.getName()) > key)
    .sort((a, b) => a.getIndex() - b.getIndex())[0];
  sh = later ? ss.insertSheet(name, later.getIndex() - 1) : ss.insertSheet(name, ss.getSheets().length);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#eeeeee');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('@');                       // Date as text DD/MM/YYYY
  sh.getRange('E:E').setNumberFormat('@');                       // Phone as text
  sh.getRange('I:K').setNumberFormat('#,##0.##');                // Rate, Amount, Bill Total
  [70, 90, 160, 110, 110, 50, 220, 50, 80, 90, 90, 380, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return sh;
}

function dataRows(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
}

function deleteBillRows(sh, billNos) {
  const rows = dataRows(sh);
  // delete bottom-up in contiguous runs
  for (let i = rows.length - 1; i >= 0; i--) {
    if (billNos.indexOf(Number(rows[i][0])) < 0) continue;
    let j = i;
    while (j - 1 >= 0 && billNos.indexOf(Number(rows[j - 1][0])) >= 0) j--;
    sh.deleteRows(j + 2, i - j + 1);
    i = j;
  }
}

function sortByBillNo(sh) {
  const n = sh.getLastRow() - 1;
  if (n > 1) sh.getRange(2, 1, n, HEADERS.length).sort([{ column: 1, ascending: true }, { column: 6, ascending: true }]);
}

function nextBillNo() {
  let max = 0;
  monthSheets().forEach(sh => {
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(r => { max = Math.max(max, Number(r[0]) || 0); });
  });
  if (max) return max + 1;
  return Number(PropertiesService.getScriptProperties().getProperty('START_BILL_NO')) || 1;
}

/* ---------- parts ---------- */

function partsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PARTS_TAB);
  if (!sh) {
    sh = ss.insertSheet(PARTS_TAB, 0);
    sh.getRange(1, 1, 1, 3).setValues([PARTS_HEADERS]).setFontWeight('bold').setBackground('#eeeeee');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 240);
  }
  return sh;
}

function readParts() {
  const sh = partsSheet();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), rate: Number(r[1]) || null }));
}

function updateParts(lines) {
  const sh = partsSheet();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, 3).getValues() : [];
  const byName = {};
  rows.forEach(r => { if (String(r[0]).trim()) byName[String(r[0]).trim().toLowerCase()] = r; });
  lines.forEach(l => {
    const name = String(l.particulars).trim();
    const k = name.toLowerCase();
    if (byName[k]) { byName[k][1] = Number(l.rate); byName[k][2] = (Number(byName[k][2]) || 0) + 1; }
    else byName[k] = [name, Number(l.rate), 1];
  });
  const out = Object.keys(byName).map(k => byName[k])
    .sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  if (n > 0) sh.getRange(2, 1, n, 3).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 3).setValues(out);
}

/* ---------- dates ---------- */

function dmy(v) {   // cell value -> "DD/MM/YYYY"
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v);
}

function toISO(dmyStr) {
  const p = String(dmyStr).split('/');
  return p.length === 3 ? p[2] + '-' + p[1].padStart(2, '0') + '-' + p[0].padStart(2, '0') : '';
}
