// NextGen appointment-list upload (Schedule › Upload appt list): parseApptReport, the file read in readApptFile, and
// OfficeStore.replaceAppts against a small in-memory mock of the Office.js range/table API.
//   node tests/appt-upload.test.js
// Checks: the "1 ESK Ordering - B" header row is found within the first 30 rows by its MRN + date headers, under NextGen's short
// names (Md Rc, Pat Name, Appt Dt, Beg Tm, Sts) or the long ones; all six columns are required, Rendering is optional; blank and
// footer lines ("Total Appointments") are skipped; dates become serials. The file read uses the first sheet only. replaceAppts writes
// MRN..Status (+ Rendering when the table has it right after Status) from the top of the table, grows the table carrying the formula
// columns down, clears leftover rows (contents only) and never touches the formula columns or the columns before MRN.
// This is a mock, not Excel: it can't catch Excel-on-the-web quirks. FICTIONAL MRNs, codes and providers only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const fnSrc = name => { const i = src.indexOf(`function ${name}(`); assert(i > 0, name); let d = 0, j = src.indexOf('{', i); for (; j < src.length; j++){ if (src[j] === '{') d++; else if (src[j] === '}' && !--d) break; } return src.slice(i, j + 1); };
// readApptFile's own read line (first sheet, header:1, raw:true, defval:''), run as is against SheetJS from assets/.
const READ = src.match(/const wb = (XLSX\.read\(await f\.arrayBuffer\(\), \{[^}\n]*\}\)); rows = (parseApptReport\(XLSX\.utils\.sheet_to_json\(wb\.Sheets\[wb\.SheetNames\[0\]\], \{[^}\n]*\}\)\));/);
if (!READ) throw new Error('not found: readApptFile read line');
const code = [line(/^const norm = .*$/m), line(/^const isF = .*$/m), line(/^const MS = .*$/m), line(/^const ymd = .*$/m), fnSrc('toSerial'),
  line(/^const APPT_COLS = .*$/m), line(/^const APPT_HDR = .*$/m), fnSrc('parseApptReport'), between('function OfficeStore(){', '// ---------- store: demo'),
  `globalThis.readAppt = async f => { const wb = ${READ[1]}; return ${READ[2]}; };`,
  'globalThis.OfficeStore = OfficeStore; globalThis.parseApptReport = parseApptReport; globalThis.ymd = ymd; globalThis.APPT_COLS = APPT_COLS;'].join('\n');
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };
const eq = (a, b, m) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), b, m);   // arrays built in the vm context have another Array.prototype

// ---- mock Excel: one "Appt List" sheet (cells keyed "r,c"), one table; formulas kept in R1C1 form (same text on every row) ----
const L = [];   // operation log
function Sheet(){ this.cells = new Map(); this.fx = new Map(); this.fmt = new Map(); }
const K = (r, c) => `${r},${c}`;
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; }
const grid = (r, f) => Array.from({ length: r.nr }, (_, i) => Array.from({ length: r.nc }, (_, j) => f(K(r.r0 + i, r.c0 + j))));
Range.prototype.load = function (){ this._v = grid(this, k => this.ws.fx.has(k) ? 'calc' : (this.ws.cells.get(k) ?? '')); this._f = grid(this, k => this.ws.fx.get(k) ?? this.ws.cells.get(k) ?? ''); };
Range.prototype.getCell = function (r, c){ return new Range(this.ws, this.r0 + r, this.c0 + c, 1, 1); };
Range.prototype.getRow = function (r){ return new Range(this.ws, this.r0 + r, this.c0, 1, this.nc); };
Range.prototype.getResizedRange = function (dr, dc){ assert(this.nr + dr > 0 && this.nc + dc > 0, 'empty range'); return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.clear = function (applyTo){ L.push('clear:' + applyTo); grid(this, k => { this.ws.cells.delete(k); this.ws.fx.delete(k); if (applyTo !== 'Contents') this.ws.fmt.delete(k); }); };
const shape = (r, a) => { assert.strictEqual(a.length, r.nr, 'row count'); a.forEach(x => assert.strictEqual(x.length, r.nc, 'col count')); };
Object.defineProperty(Range.prototype, 'values', { get(){ return this._v; }, set(a){ shape(this, a); L.push('values'); a.forEach((row, i) => row.forEach((v, j) => { const k = K(this.r0 + i, this.c0 + j); this.ws.fx.delete(k); this.ws.cells.set(k, v); })); } });
Object.defineProperty(Range.prototype, 'formulas', { get(){ return this._f; } });
Object.defineProperty(Range.prototype, 'formulasR1C1', { get(){ return this._f; }, set(a){ shape(this, a); L.push('formulasR1C1'); a.forEach((row, i) => row.forEach((f, j) => this.ws.fx.set(K(this.r0 + i, this.c0 + j), f))); } });
function Table(ws, r0, c0, headers){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = 1; this.nc = headers.length; headers.forEach((h, j) => ws.cells.set(K(r0, c0 + j), h)); }
Table.prototype.getHeaderRowRange = function (){ return new Range(this.ws, this.r0, this.c0, 1, this.nc); };
Table.prototype.getDataBodyRange = function (){ const r = new Range(this.ws, this.r0 + 1, this.c0, this.nr - 1, this.nc); r.rowCount = this.nr - 1; r.columnCount = this.nc; return r; };
Table.prototype.getRange = function (){ return new Range(this.ws, this.r0, this.c0, this.nr, this.nc); };
Table.prototype.resize = function (rg){ assert(rg.ws === this.ws && rg.r0 === this.r0 && rg.c0 === this.c0 && rg.nc === this.nc, 'resize keeps the anchor and columns'); L.push('resize'); this.nr = rg.nr; };
let T;
const ctx = () => ({ sync: async () => {}, workbook: { tables: { getItem: () => T } } });
const C = { Excel: { run: fn => fn(ctx()), ClearApplyTo: { contents: 'Contents' } }, setTimeout, S: {}, XLSX: require(path.join(__dirname, '..', 'assets', 'xlsx.full.min.js')) };
vm.createContext(C); vm.runInContext(code, C);

// Appt List as in the workbook: sheet row 4 = header, the table starts at L (a hand column before MRN at M), formula columns after Status.
const AP_LEAD = 11, FX = { 'Dose': '=IF(RC[-3]="","",56)', 'Lookup Key': '=RC[-7]&"|"&RC[-6]' };
function apptTable(headers, body){
  const ws = new Sheet(); T = new Table(ws, 3, AP_LEAD, headers); T.nr = 1 + body.length;
  body.forEach((row, i) => headers.forEach((h, j) => { const k = K(4 + i, AP_LEAD + j); if (FX[h]) ws.fx.set(k, FX[h]); else ws.cells.set(k, row[j] ?? ''); if (h === 'Date') ws.fmt.set(k, 'm/d/yyyy'); }));
  return T;
}
const at = (r, h) => { const c = T.c0 + hdr().indexOf(h), k = K(T.r0 + 1 + r, c); return T.ws.fx.has(k) ? { fx: T.ws.fx.get(k) } : (T.ws.cells.get(k) ?? ''); };
const hdr = () => Array.from({ length: T.nc }, (_, j) => T.ws.cells.get(K(T.r0, T.c0 + j)));
const H6 = ['Notes', 'MRN', 'Patient', 'Date', 'Time', 'Event', 'Status'];
const D = (m, d) => C.ymd(2026, m, d);
const old = i => ['keep ' + i, 900300 + i, 'OLDROW', D(10, 1), '9:00 AM', 'ESK Treatment', 'Kept'];
const up = (i, rend = 'Provider A') => [900201 + i, 'ALDBEN', D(10, 19) + (i % 5), '10:30 AM', 'ESK Treatment', 'Scheduled', rend];

// A synthetic "1 ESK Ordering - B" export: title block, header row, data, footer. Columns are in the report's order, not the table's.
const REPORT = (hdrs, extra = []) => [
  ['1 ESK Ordering - B', '', '', '', '', '', '', '', ''], ['Practice: FICTIONAL TEST CLINIC', '', '', '', '', '', '', '', ''], ['Date range: 10/19/2026 - 10/23/2026', '', '', '', '', '', '', '', ''], [],
  hdrs,
  ['Room 2', 900201, 'ALDBEN', '10/19/2026', '9:00 AM', 'ESK Treatment', 'Scheduled', '  Provider A  ', 'x'],
  ['', '0900202', 'BROCAM', 46315, 0.4375, 'ESK Start', 'Checked In', 'Provider B', ''],
  ['', 900203, 'CARDEL', '2026-10-21', '1:15 PM', 'ESK Follow Up', 'Cancelled', '', ''],
  ['', 900204, 'DAVEMI', 'TBD', '2:00 PM', 'ESK Treatment', 'Scheduled', 'Provider A', ''],   // unreadable date: skipped
  ['', '', '', '', '', '', '', '', ''],
  ...extra,
  ['Total Appointments: 4', '', '', '', '', '', '', '', ''],   // footer in the first column
  ['', '', 'Total Appointments', 4, '', '', '', '', ''],   // footer under the patient column
  ['', 'Total Appointments', '', '', '', '', '', '', '']   // footer in the MRN column
];
const SHORT = ['Location', 'Md Rc', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Sts', 'Rendering', 'Resource'];

(async () => {
  // ---------- parseApptReport ----------
  await t('NextGen short headers: header row found below a title block, columns mapped by name', () => {
    const rows = C.parseApptReport(REPORT(SHORT));
    assert(C.S.upRendering === true, 'upRendering');
    eq(rows, [
      [900201, 'ALDBEN', D(10, 19), '9:00 AM', 'ESK Treatment', 'Scheduled', 'Provider A'],
      ['0900202', 'BROCAM', 46315, 0.4375, 'ESK Start', 'Checked In', 'Provider B'],
      [900203, 'CARDEL', D(10, 21), '1:15 PM', 'ESK Follow Up', 'Cancelled', '']
    ]);
  });
  await t('footer ("Total Appointments"), blank lines and unreadable dates are skipped', () => {
    const rows = C.parseApptReport(REPORT(SHORT));
    assert(!rows.some(r => /total/i.test(r.join(' '))), 'no footer row'); assert(!rows.some(r => r[0] === 900204), 'TBD date skipped'); assert(rows.every(r => r[0] !== '' && typeof r[2] === 'number'));
  });
  await t('long header names and the other aliases', () => {
    for (const [hs, rend] of [
      [['Location', 'MRN', 'Patient Name', 'Appt Date', 'Time', 'Event', 'Status', 'Rendering Provider', 'Resource'], true],
      [['Location', 'mrn', 'PATIENT', 'Date', 'Time ', 'event', 'status', 'Provider', 'Resource'], true],
      [['Location', 'Md Rc', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Sts', 'Notes', 'Resource'], false]
    ]){
      const rows = C.parseApptReport(REPORT(hs));
      assert(rows && rows.length === 3, hs.join()); assert.strictEqual(C.S.upRendering, rend, hs.join());
      eq(rows[0].slice(0, 6), [900201, 'ALDBEN', D(10, 19), '9:00 AM', 'ESK Treatment', 'Scheduled']);
      assert.strictEqual(rows[0][6], rend ? 'Provider A' : '', 'Rendering trimmed, blank when the report has none'); assert.strictEqual(rows[0].length, 7);
    }
  });
  await t('not the report: a missing required column, no header, or a header below row 30 → null', () => {
    for (const drop of ['Pat Name', 'Beg Tm', 'Event', 'Sts']) assert.strictEqual(C.parseApptReport(REPORT(SHORT.map(h => h === drop ? 'Other' : h))), null, drop);
    assert.strictEqual(C.parseApptReport(REPORT(SHORT.map(h => h === 'Appt Dt' ? 'Visit' : h))), null, 'no date header');
    assert.strictEqual(C.parseApptReport([['MRN only in a title line'], ['MRN'], ['Appt Dt']]), null, 'MRN and date must share a row');
    assert.strictEqual(C.parseApptReport([]), null);
    const pad = k => [...Array.from({ length: k }, () => ['']), SHORT, ['', 900201, 'ALDBEN', '10/19/2026', '9:00 AM', 'ESK Treatment', 'Scheduled', 'Provider A', '']];
    assert.strictEqual(C.parseApptReport(pad(29)).length, 1, 'header on row 30 is found'); assert.strictEqual(C.parseApptReport(pad(30)), null, 'header on row 31 is not');
  });

  // ---------- the file read (readApptFile's own line, real SheetJS) ----------
  const XL = C.XLSX, file = (aoa, bookType, more) => { const wb = XL.utils.book_new(); XL.utils.book_append_sheet(wb, XL.utils.aoa_to_sheet(aoa), 'Sheet1'); if (more) XL.utils.book_append_sheet(wb, XL.utils.aoa_to_sheet(more), 'Sheet2'); const buf = XL.write(wb, { type: 'array', bookType }); return { arrayBuffer: async () => buf }; };
  const asCsv = aoa => { const s = aoa.map(r => r.map(v => /[",]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v).join(',')).join('\r\n'); const b = Buffer.from(s, 'utf8'); return { arrayBuffer: async () => b.buffer.slice(b.byteOffset, b.byteOffset + b.length) }; };
  await t('xlsx / xls / csv exports read through SheetJS give the same rows', async () => {
    const expect = [[900201, 'ALDBEN', D(10, 19)], ['BROCAM', 46315], [900203, 'CARDEL', D(10, 21)]];
    for (const [kind, f] of [['xlsx', file(REPORT(SHORT), 'xlsx')], ['xls', file(REPORT(SHORT), 'biff8')], ['csv', asCsv(REPORT(SHORT))]]){
      const rows = await C.readAppt(f);
      assert(rows && rows.length === 3, kind); eq([rows[0].slice(0, 3), rows[1].slice(1, 3), rows[2].slice(0, 3)].map(r => r.map(v => typeof v === 'string' && /^\d+$/.test(v) ? +v : v)), expect, kind);
      assert(rows.every(r => r[0] !== '' && Number.isInteger(r[2])), kind + ': serial dates'); assert.strictEqual(rows[0][6], 'Provider A', kind);
    }
  });
  await t('only the first sheet is read', async () => {
    assert.strictEqual((await C.readAppt(file(REPORT(SHORT), 'xlsx', [['junk']]))).length, 3);
    assert.strictEqual(await C.readAppt(file([['Cover sheet'], ['nothing here']], 'xlsx', REPORT(SHORT))), null);
  });

  // ---------- OfficeStore.replaceAppts ----------
  const store = C.OfficeStore();
  await t('shrink, no Rendering column: writes MRN..Status only, clears leftovers (contents only), formula and hand columns untouched', async () => {
    const H = [...H6, 'Dose', 'Lookup Key']; apptTable(H, [0, 1, 2, 3].map(old)); L.length = 0;
    await store.replaceAppts([0, 1].map(i => up(i)));
    assert.strictEqual(T.nr, 5, 'table not resized'); assert(!L.includes('resize'));
    eq(hdr().map((h, j) => at(0, h)), ['keep 0', 900201, 'ALDBEN', D(10, 19), '10:30 AM', 'ESK Treatment', 'Scheduled', { fx: FX.Dose }, { fx: FX['Lookup Key'] }]);
    eq(at(1, 'MRN'), 900202);
    for (const r of [2, 3]){ for (const h of H6.slice(1)) assert.strictEqual(at(r, h), '', `row ${r} ${h} cleared`); assert.strictEqual(at(r, 'Notes'), 'keep ' + r, 'column before MRN kept'); eq(at(r, 'Dose'), { fx: FX.Dose }, 'formula kept'); eq(at(r, 'Lookup Key'), { fx: FX['Lookup Key'] }); }
    eq(L.filter(x => x.startsWith('clear')), ['clear:Contents']);
    for (let r = 0; r < 4; r++) assert.strictEqual(T.ws.fmt.get(K(4 + r, AP_LEAD + 3)), 'm/d/yyyy', 'date format kept');
  });
  for (const rh of ['Rendering', 'Rendering Provider', 'Provider']){
    await t(`grow with a "${rh}" column after Status: table grows, formulas carried down, Rendering written`, async () => {
      const H = [...H6, rh, 'Dose', 'Lookup Key']; apptTable(H, [0, 1].map(old)); L.length = 0;
      await store.replaceAppts([0, 1, 2, 3, 4].map(i => up(i, 'Provider ' + 'ABCDE'[i])));
      assert.strictEqual(T.nr, 6); assert(L.includes('resize')); assert(L.indexOf('resize') < L.indexOf('formulasR1C1') && L.indexOf('formulasR1C1') < L.indexOf('values'), 'grow, fill formulas, then write');
      for (let r = 0; r < 5; r++){ eq(at(r, 'MRN'), 900201 + r); eq(at(r, rh), 'Provider ' + 'ABCDE'[r]); eq(at(r, 'Dose'), { fx: FX.Dose }, `row ${r} Dose formula`); eq(at(r, 'Lookup Key'), { fx: FX['Lookup Key'] }); }
      assert.strictEqual(at(0, 'Notes'), 'keep 0'); assert.strictEqual(at(4, 'Notes'), '', 'new row: hand column left blank'); assert(!L.some(x => x.startsWith('clear')));
    });
  }
  await t('same size: no resize, no clear', async () => {
    apptTable([...H6, 'Rendering', 'Dose'], [0, 1, 2].map(old)); L.length = 0;
    await store.replaceAppts([0, 1, 2].map(i => up(i))); assert(!L.includes('resize') && !L.some(x => x.startsWith('clear'))); eq(at(2, 'Rendering'), 'Provider A'); eq(at(2, 'Dose'), { fx: FX.Dose });
  });
  await t('more than 250 rows (chunked writes) after growing from one row', async () => {
    apptTable([...H6, 'Rendering', 'Dose', 'Lookup Key'], [old(0)]); L.length = 0;
    const rows = Array.from({ length: 300 }, (_, i) => up(i)); await store.replaceAppts(rows);
    assert.strictEqual(T.nr, 301); assert(L.filter(x => x === 'values').length >= 2, 'written in chunks');
    for (const r of [0, 249, 250, 299]){ eq(at(r, 'MRN'), 900201 + r); eq(at(r, 'Lookup Key'), { fx: FX['Lookup Key'] }); }
  });
  await t('end to end: parsed report rows land in the table as parsed (Date as a serial, Time as the report gave it)', async () => {
    const parsed = C.parseApptReport(REPORT(SHORT));
    apptTable([...H6, 'Rendering', 'Dose'], [0, 1, 2, 3, 4].map(old)); await store.replaceAppts(parsed);
    eq(H6.slice(1).concat('Rendering').map(h => at(1, h)), ['0900202', 'BROCAM', 46315, 0.4375, 'ESK Start', 'Checked In', 'Provider B']);
    eq(at(3, 'MRN'), ''); eq(at(3, 'Dose'), { fx: FX.Dose });
    apptTable([...H6, 'Dose'], [0, 1].map(old)); await store.replaceAppts(parsed); eq(at(2, 'Dose'), { fx: FX.Dose }, 'no Rendering column: the 7th value is not written');
  });
  await t('stops when MRN..Status are out of order or not contiguous', async () => {
    for (const H of [['Notes', 'Patient', 'MRN', 'Date', 'Time', 'Event', 'Status'], ['MRN', 'Patient', 'Date', 'Dose', 'Time', 'Event', 'Status']]){
      apptTable(H, [old(0)]); L.length = 0;
      await assert.rejects(() => store.replaceAppts([up(0)]), /should start with MRN, Patient, Date, Time, Event, Status/); assert(!L.includes('values') && !L.includes('resize'), 'nothing written');
    }
  });
  await t('stops when a MRN..Status (or Rendering) cell holds a formula', async () => {
    for (const col of ['Event', 'Rendering']){
      apptTable([...H6, 'Rendering', 'Dose'], [old(0), old(1)]); T.ws.fx.set(K(4, AP_LEAD + hdr().indexOf(col)), '=RC[-1]'); L.length = 0;
      await assert.rejects(() => store.replaceAppts([0, 1, 2].map(i => up(i))), /holds a formula/, col); assert(!L.includes('values') && !L.includes('resize'), col + ': nothing written');
    }
  });
  console.log(`\n${n} passed`);
})().catch(e => { console.error(e); process.exit(1); });
