// OfficeStore.upsertBenefits against a small in-memory mock of the Office.js range/table API.
//   node tests/benefits-excel-mock.test.js
// Checks: first import creates a visible "Benefits" sheet + tblBenefits from plain cells; MRNs are written as text ('@')
// so leading zeros stay; a second import updates rows in place, keeps others, grows the table, never deletes rows.
// This is a mock, not Excel: it can't catch Excel-on-the-web quirks. FICTIONAL MRNs and names only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const norm = .*$/m), line(/^const isF = .*$/m), line(/^const mrnKey = .*$/m), line(/^const same = .*$/m), line(/^const snapMatch = .*$/m),
  between('// ---------- Sohar benefits: parsing', '// ---------- Sohar benefits: end'), between('function OfficeStore(){', '// ---------- store: demo'),
  'globalThis.OfficeStore = OfficeStore; globalThis.BEN_COLS = BEN_COLS;'].join('\n');

// ---- mock Excel ----
const L = []; // log of operations, to check ordering (numberFormat before values) and that nothing is deleted
function Sheet(name){ this.name = name; this.cells = new Map(); this.fmt = new Map(); this.visibility = 'Visible'; this.tables = { add: (rg, hh) => { const t = new Table(this, rg.r0, rg.c0, rg.nr, rg.nc); WB.tables.push(t); L.push('tables.add'); return t; } }; }
Sheet.prototype.getRange = function (a1){ const m = a1.match(/^([A-Z]+)(\d+)$/); const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return new Range(this, +m[2] - 1, c, 1, 1); };
Sheet.prototype.getUsedRangeOrNullObject = function (){ return { isNullObject: ![...this.cells.values()].some(v => v !== '') }; };
Sheet.prototype.load = () => {};
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.format = { autofitColumns: () => {} }; }
Range.prototype.getResizedRange = function (dr, dc){ return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.getOffsetRange = function (dr, dc){ return new Range(this.ws, this.r0 + dr, this.c0 + dc, this.nr, this.nc); };
Range.prototype.getCell = function (r, c){ return new Range(this.ws, this.r0 + r, this.c0 + c, 1, 1); };
Range.prototype.load = function (){ const g = (m, d) => Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => { const v = m.get(`${this.r0 + i},${this.c0 + j}`); return v === undefined ? d : v; })); this._v = g(this.ws.cells, ''); this._f = g(this.ws.fmt, 'General'); };
const shape = (r, a) => { assert.strictEqual(a.length, r.nr, 'row count'); a.forEach(x => assert.strictEqual(x.length, r.nc, 'col count')); };
Object.defineProperty(Range.prototype, 'values', { get(){ return this._v; }, set(a){ shape(this, a); L.push('values'); a.forEach((row, i) => row.forEach((v, j) => { const k = `${this.r0 + i},${this.c0 + j}`; const txt = this.ws.fmt.get(k) === '@'; this.ws.cells.set(k, !txt && /^\d+$/.test(v) ? Number(v) : v); })); } });   // like Excel: a digit string in a General cell becomes a number
Object.defineProperty(Range.prototype, 'formulas', { get(){ return this._v; } });
Object.defineProperty(Range.prototype, 'numberFormat', { get(){ return this._f; }, set(a){ shape(this, a); L.push('numberFormat'); a.forEach((row, i) => row.forEach((v, j) => this.ws.fmt.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
function Table(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.name = 'Table9'; this.isNullObject = false; this.worksheet = ws; }
Table.prototype.getHeaderRowRange = function (){ return new Range(this.ws, this.r0, this.c0, 1, this.nc); };
Table.prototype.getDataBodyRange = function (){ const r = new Range(this.ws, this.r0 + 1, this.c0, this.nr - 1, this.nc); r.rowCount = this.nr - 1; r.columnCount = this.nc; return r; };
Table.prototype.getRange = function (){ return new Range(this.ws, this.r0, this.c0, this.nr, this.nc); };
Table.prototype.resize = function (rg){ assert(rg.nr >= this.nr, 'table never shrinks'); L.push('resize'); this.nr = rg.nr; };
let WB;
const ctx = () => ({ sync: async () => {}, workbook: {
  tables: { getItemOrNullObject: n => WB.tables.find(t => t.name === n) || { isNullObject: true }, getItem: n => WB.tables.find(t => t.name === n) },
  worksheets: { getItemOrNullObject: n => WB.sheets.find(s => s.name === n) || { isNullObject: true }, add: n => { const s = new Sheet(n); WB.sheets.push(s); L.push('sheet.add'); return s; } } } });
const C = { Excel: { run: fn => fn(ctx()), SheetVisibility: { visible: 'Visible' } }, setTimeout };
vm.createContext(C); vm.runInContext(code, C);

(async () => {
  WB = { sheets: [], tables: [] };
  const store = C.OfficeStore(), W = C.BEN_COLS.length;
  const r = (mrn, name, copay) => { const a = Array(W).fill(''); a[0] = mrn; a[1] = name; a[9] = copay; a[W - 1] = '2026-10-14 09:00'; return a; };
  // 1. first import: creates sheet + table
  let res = await store.upsertBenefits([r('0900101', 'ALDBEN', '30'), r('900102', 'BROCAM', '0')]);
  assert.strictEqual(JSON.stringify(res), '{"updated":0,"added":2}');
  const ws = WB.sheets.find(s => s.name === 'Benefits'); assert(ws, 'Benefits sheet'); assert.strictEqual(ws.visibility, 'Visible');
  const t = WB.tables.find(x => x.name === 'tblBenefits'); assert(t, 'tblBenefits'); assert.strictEqual(t.nr, 3);
  assert.strictEqual(ws.cells.get('0,0'), 'MRN'); assert.strictEqual(ws.cells.get('1,0'), '0900101', 'leading zero kept'); assert.strictEqual(ws.cells.get('2,0'), '900102', 'digits stay text');
  const firstValues = L.indexOf('values'), firstFmt = L.indexOf('numberFormat'); assert(firstFmt >= 0 && firstFmt < firstValues, 'numberFormat set before values');
  assert(L.indexOf('tables.add') > L.lastIndexOf('values'), 'table added after plain cells are written');
  console.log('ok first import creates Benefits / tblBenefits as text');
  // 2. second import: update one in place, add one, keep the other
  L.length = 0;
  res = await store.upsertBenefits([r('900101', 'ALDBEN', '45'), r('900103', 'CARDEL', '10')]);
  assert.strictEqual(JSON.stringify(res), '{"updated":1,"added":1}');
  assert.strictEqual(t.nr, 4); assert.strictEqual(ws.cells.get('1,0'), '900101'); assert.strictEqual(ws.cells.get('1,9'), '45'); assert.strictEqual(ws.cells.get('2,0'), '900102'); assert.strictEqual(ws.cells.get('3,0'), '900103');
  assert(L.includes('resize') && !L.includes('tables.add') && !L.includes('sheet.add'));
  console.log('ok second import updates in place, keeps others, grows the table');
  // 3. re-import the same MRNs: no growth
  L.length = 0; res = await store.upsertBenefits([r('900103', 'CARDEL', '11')]);
  assert.strictEqual(JSON.stringify(res), '{"updated":1,"added":0}'); assert.strictEqual(t.nr, 4); assert(!L.includes('resize'));
  console.log('ok re-import of the same MRN updates without resizing');
  // 4. a Benefits sheet with data but no table → stops
  WB = { sheets: [Object.assign(new Sheet('Benefits'), {})], tables: [] }; WB.sheets[0].cells.set('0,0', 'something');
  await assert.rejects(() => C.OfficeStore().upsertBenefits([r('1', 'X', '')]), /already has data/);
  console.log('ok refuses to overwrite a Benefits sheet that has data but no table');
  // 5. table whose columns were reordered → stops
  WB = { sheets: [], tables: [] }; await C.OfficeStore().upsertBenefits([r('1', 'X', '')]); const s2 = WB.sheets[0]; s2.cells.set('0,1', 'Per Nbr'); s2.cells.set('0,2', 'Patient Name');
  await assert.rejects(() => C.OfficeStore().upsertBenefits([r('2', 'Y', '')]), /should have these columns/);
  console.log('ok refuses when tblBenefits columns are out of order');
  console.log('\n5 mock tests passed');
})().catch(e => { console.error(e); process.exit(1); });
