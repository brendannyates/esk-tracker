// OfficeStore.saveEligibility against a small in-memory mock of the Office.js range/table API (same mock idea as benefits-excel-mock).
//   node tests/elig-feed-excel-mock.test.js
// Checks: the hidden "Eligibility Checks" sheet is created hidden, written as text and replaced on the next pull; tblBenefits gets the
// new columns at the right end, the pull that counts fills the matched row, a tracker MRN with no row is added, BIDF cells stay.
// This is a mock, not Excel: it can't catch Excel-on-the-web quirks. FICTIONAL codes, MRNs and ids only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const norm = .*$/m), line(/^const isF = .*$/m), line(/^const mrnKey = .*$/m), line(/^const same = .*$/m), line(/^const snapMatch = .*$/m),
  between('// ---------- Sohar benefits: parsing', '// ---------- Sohar benefits: end'), between('// ---------- Eligibility feed: pure', '// ---------- Eligibility feed: end'),
  between('function OfficeStore(){', '// ---------- store: demo'), between('const ELIG_FEED_COLS', 'function DemoStore(){'),
  'Object.assign(globalThis, { OfficeStore, BEN_COLS, ELIG_BEN_ADD, ELIG_FEED_COLS, ELIG_SHEET });'].join('\n');

function Sheet(name){ this.name = name; this.cells = new Map(); this.fmt = new Map(); this.visibility = 'Visible'; this.tables = { add: (rg) => { const t = new Table(this, rg.r0, rg.c0, rg.nr, rg.nc); WB.tables.push(t); return t; } }; }
Sheet.prototype.getRange = function (a1){ if (!a1) return new Range(this, 0, 0, 1e6, 16384); const m = a1.match(/^([A-Z]+)(\d+)$/); const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return new Range(this, +m[2] - 1, c, 1, 1); };
Sheet.prototype.getUsedRangeOrNullObject = function (){ return { isNullObject: ![...this.cells.values()].some(v => v !== '') }; };
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.format = { autofitColumns: () => {} }; }
Range.prototype.getResizedRange = function (dr, dc){ return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.getOffsetRange = function (dr, dc){ return new Range(this.ws, this.r0 + dr, this.c0 + dc, this.nr, this.nc); };
Range.prototype.getCell = function (r, c){ return new Range(this.ws, this.r0 + r, this.c0 + c, 1, 1); };
Range.prototype.clear = function (){ for (const k of [...this.ws.cells.keys()]){ const [r, c] = k.split(',').map(Number); if (r >= this.r0 && r < this.r0 + this.nr && c >= this.c0 && c < this.c0 + this.nc) this.ws.cells.delete(k); } };
Range.prototype.load = function (){ const g = (m, d) => Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => { const v = m.get(`${this.r0 + i},${this.c0 + j}`); return v === undefined ? d : v; })); this._v = g(this.ws.cells, ''); };
const shape = (r, a) => { assert.strictEqual(a.length, r.nr, 'row count'); a.forEach(x => assert.strictEqual(x.length, r.nc, 'col count')); };
Object.defineProperty(Range.prototype, 'values', { get(){ return this._v; }, set(a){ shape(this, a); a.forEach((row, i) => row.forEach((v, j) => { const k = `${this.r0 + i},${this.c0 + j}`; const txt = this.ws.fmt.get(k) === '@'; this.ws.cells.set(k, !txt && /^\d+$/.test(v) ? Number(v) : v); })); } });
Object.defineProperty(Range.prototype, 'formulas', { get(){ return this._v; } });
Object.defineProperty(Range.prototype, 'numberFormat', { set(a){ shape(this, a); a.forEach((row, i) => row.forEach((v, j) => this.ws.fmt.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
function Table(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.name = 'Table9'; this.isNullObject = false;
  this.columns = { add: (_i, _v, name) => { this.ws.cells.set(`${this.r0},${this.c0 + this.nc}`, name); this.nc++; } }; }
Table.prototype.getHeaderRowRange = function (){ return new Range(this.ws, this.r0, this.c0, 1, this.nc); };
Table.prototype.getDataBodyRange = function (){ const r = new Range(this.ws, this.r0 + 1, this.c0, this.nr - 1, this.nc); r.rowCount = this.nr - 1; r.columnCount = this.nc; return r; };
Table.prototype.getRange = function (){ return new Range(this.ws, this.r0, this.c0, this.nr, this.nc); };
Table.prototype.resize = function (rg){ assert(rg.nr >= this.nr, 'table never shrinks'); this.nr = rg.nr; };
let WB;
const ctx = () => ({ sync: async () => {}, workbook: {
  tables: { getItemOrNullObject: n => WB.tables.find(t => t.name === n) || { isNullObject: true } },
  worksheets: { getItemOrNullObject: n => WB.sheets.find(s => s.name === n) || { isNullObject: true }, add: n => { const s = new Sheet(n); WB.sheets.push(s); return s; } } } });
const C = { Excel: { run: fn => fn(ctx()), SheetVisibility: { visible: 'Visible', hidden: 'Hidden' }, ClearApplyTo: { contents: 'Contents' } }, setTimeout };
vm.createContext(C); vm.runInContext(code, C);

const H = C.ELIG_FEED_COLS, rowOf = o => H.map(h => String(o[h] ?? ''));
const F = (mrn, at, st, id, copay) => rowOf({ 'Pulled At': at, LASFIR: 'TESABC', MRN: mrn, Kind: 'Verification', 'Sohar Status': st, 'Sohar ID': id, Copay: copay, 'Sohar Link': 'https://app.soharhealth.com/verifications/' + id });
const cellsOf = (ws, r) => { const out = []; for (let c = 0; ws.cells.has(`${r},${c}`); c++) out.push(ws.cells.get(`${r},${c}`)); return out; };
(async () => {
  // 1. a workbook with tblBenefits (BEN_COLS + one BIDF column) and one patient row
  WB = { sheets: [], tables: [] };
  const bs = new Sheet('Benefits'); WB.sheets.push(bs);
  const hdr = [...C.BEN_COLS, 'BIDF Per Treatment'], W = hdr.length;
  hdr.forEach((h, c) => bs.cells.set(`0,${c}`, h));
  const r1 = Array(W).fill(''); r1[0] = '0900101'; r1[1] = 'ALDBEN'; r1[9] = '10'; r1[W - 1] = '$140'; r1.forEach((v, c) => bs.cells.set(`1,${c}`, v));
  const t = new Table(bs, 0, 0, 2, W); t.name = 'tblBenefits'; WB.tables.push(t);
  const store = C.OfficeStore(), patients = [{ mrn: '900101', name: 'ALDBEN' }, { mrn: '0900102', name: 'BROCAM' }];
  const rows = [F('900101', '2026-11-08T15:00:00Z', 'error.member.id', 'e1', ''), F('900101', '2026-11-01T15:00:00Z', 'complete.eligible', 'v1', '30'), F('900102', '2026-11-07T15:00:00Z', 'Pending', 'p1', '')];
  let res = await store.saveEligibility(H, rows, patients, '2026-11-09 07:30');
  assert.strictEqual(JSON.stringify(res), '{"rows":3,"updated":1,"added":1}');
  const es = WB.sheets.find(s => s.name === C.ELIG_SHEET); assert(es, 'Eligibility Checks sheet'); assert.strictEqual(es.visibility, 'Hidden');
  assert.strictEqual(JSON.stringify(cellsOf(es, 0)), JSON.stringify(H)); assert.strictEqual(es.cells.get('1,2'), '900101'); assert.strictEqual(es.fmt.get('1,2'), '@');
  console.log('ok hidden sheet created hidden, header = feed columns, text cells');
  const H2 = cellsOf(bs, 0); assert.strictEqual(JSON.stringify(H2), JSON.stringify([...hdr, ...C.ELIG_BEN_ADD])); assert.strictEqual(t.nr, 3);
  const col = c => H2.indexOf(c);
  assert.strictEqual(bs.cells.get(`1,${col('Copay')}`), '30', 'the eligible verification counts, not the newer error'); assert.strictEqual(bs.cells.get(`1,${col('Sohar ID')}`), 'v1');
  assert.strictEqual(bs.cells.get(`1,${col('BIDF Per Treatment')}`), '$140', 'BIDF kept'); assert.strictEqual(bs.cells.get(`1,${col('Source')}`), 'Eligibility feed');
  assert.strictEqual(bs.cells.get('2,0'), '0900102'); assert.strictEqual(bs.cells.get('2,1'), 'BROCAM'); assert.strictEqual(bs.cells.get(`2,${col('Sohar Status')}`), 'Pending');
  console.log('ok tblBenefits: new columns at the end, matched row filled, missing MRN added, BIDF untouched');
  // 2. the next pull replaces the hidden sheet (fewer rows) and doesn't add the columns again
  res = await store.saveEligibility(H, [F('900102', '2026-11-09T15:00:00Z', 'complete.eligible', 'v2', '0')], patients, 'later');
  assert.strictEqual(JSON.stringify(res), '{"rows":1,"updated":1,"added":0}');
  assert.strictEqual(es.cells.get('1,22'), 'v2'); assert(!es.cells.has('2,0') && !es.cells.has('3,0'), 'old rows cleared');
  assert.strictEqual(JSON.stringify(cellsOf(bs, 0)), JSON.stringify(H2)); assert.strictEqual(bs.cells.get(`2,${col('Copay')}`), '0'); assert.strictEqual(bs.cells.get(`1,${col('Sohar ID')}`), 'v1', 'a patient with no new pull keeps their row');
  console.log('ok next pull replaces the sheet, columns not added twice');
  // 3. no tblBenefits yet: creates it from plain cells
  WB = { sheets: [], tables: [] };
  res = await C.OfficeStore().saveEligibility(H, [F('42', '2026-11-08', 'Pending', 'z', '')], [{ mrn: '0042', name: 'TESABC' }], 'now');
  const nt = WB.tables.find(x => x.name === 'tblBenefits'); assert(nt && nt.nr === 2); assert.strictEqual(res.added, 1);
  assert.strictEqual(WB.sheets.find(s => s.name === 'Benefits').cells.get('1,0'), '0042');
  console.log('ok creates tblBenefits when the workbook has none');
  console.log('\n4 mock tests passed');
})().catch(e => { console.error(e); process.exit(1); });
