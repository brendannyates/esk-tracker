// Access tab: Tx Frequency parsing, and OfficeStore.setDropdown (Access › Merge › Write to the tracker) against a small mock of the Office.js API.
//   node tests/access.test.js
// Checks: frequency labels read as visits per week; the clean list goes under its REF header (found, or added after the last header),
// old entries below it are cleared, cells are text, and the Master Patient Tracker column's dropdown points at exactly that list.
// This is a mock, not Excel: it can't catch Excel-on-the-web quirks. FICTIONAL values only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const fnSrc = name => { const i = src.indexOf(`function ${name}(`); assert(i > 0, name); let d = 0, j = src.indexOf('{', i); for (; j < src.length; j++){ if (src[j] === '{') d++; else if (src[j] === '}' && !--d) break; } return src.slice(i, j + 1); };
const code = [line(/^const norm = .*$/m), line(/^const isF = .*$/m), line(/^const same = .*$/m), line(/^const snapMatch = .*$/m), line(/^const colLetter = .*$/m), line(/^const BEN_COLS = .*$/m),
  fnSrc('parseRate'), between('function OfficeStore(){', '// ---------- store: demo'), 'globalThis.OfficeStore = OfficeStore; globalThis.parseRate = parseRate;'].join('\n');
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

// ---- mock Excel: one REF sheet (cells keyed "r,c"), one patients table ----
function Sheet(name, cells){ this.name = name; this.cells = new Map(Object.entries(cells || {})); this.fmt = new Map(); }
Sheet.prototype.getUsedRange = function (){ const ks = [...this.cells.keys()].filter(k => this.cells.get(k) !== '').map(k => k.split(',').map(Number)); const r0 = Math.min(...ks.map(k => k[0])), c0 = Math.min(...ks.map(k => k[1])), r1 = Math.max(...ks.map(k => k[0])), c1 = Math.max(...ks.map(k => k[1])); const r = new Range(this, r0, c0, r1 - r0 + 1, c1 - c0 + 1); r.rowIndex = r0; r.columnIndex = c0; return r; };
Sheet.prototype.getCell = function (r, c){ return new Range(this, r, c, 1, 1); };
Sheet.prototype.getRange = function (a1){ const m = a1.match(/^([A-Z]+)(\d+)$/); const c = m[1].split('').reduce((x, ch) => x * 26 + ch.charCodeAt(0) - 64, 0) - 1; return new Range(this, +m[2] - 1, c, 1, 1); };
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; }
Range.prototype.load = function (){ this._v = Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this.ws.cells.get(`${this.r0 + i},${this.c0 + j}`) ?? '')); };
Range.prototype.getResizedRange = function (dr, dc){ return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.clear = function (){ for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.ws.cells.set(`${this.r0 + i},${this.c0 + j}`, ''); };
Object.defineProperty(Range.prototype, 'values', { get(){ return this._v; }, set(a){ assert.strictEqual(a.length, this.nr); a.forEach((row, i) => row.forEach((v, j) => this.ws.cells.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
Object.defineProperty(Range.prototype, 'numberFormat', { set(a){ assert.strictEqual(a.length, this.nr); a.forEach((row, i) => row.forEach((v, j) => this.ws.fmt.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
let WB;
const DV = () => ({ cleared: 0, clear(){ this.cleared++; this.rule = null; }, rule: null, errorAlert: null });
function PT(headers){ this.headers = headers; this.dv = headers.map(DV); }
PT.prototype.getHeaderRowRange = function (){ const h = this.headers; return { load(){ this.values = [h]; } }; };
Object.defineProperty(PT.prototype, 'columns', { get(){ return { getItemAt: i => ({ getDataBodyRange: () => ({ dataValidation: this.dv[i] }) }) }; } });
const ctx = () => ({ sync: async () => {}, workbook: { tables: { getItem: () => WB.pt }, worksheets: { items: WB.sheets, load(){}, add: nm => { const s = new Sheet(nm); WB.sheets.push(s); return s; } } } });
const C = { Excel: { run: fn => fn(ctx()), ClearApplyTo: { contents: 'Contents' } }, setTimeout, Office: {} };
vm.createContext(C); vm.runInContext(code, C);
const at = (ws, r, c) => ws.cells.get(`${r},${c}`) ?? '';

(async () => {
  await t('frequency labels → visits per week', () => {
    const cases = { '2x per week': 2, '1x per week': 1, 'weekly': 1, '1x per 2 weeks': 0.5, 'EOW': 0.5, 'every other week': 0.5, '1x every 3 weeks': 1 / 3, '1X every 3 weeks ': 1 / 3, '1x per 4 weeks': 0.25, 'Monthly': 0.25, 'q6w': 1 / 6, 'every 6 weeks': 1 / 6, '3x per week': 3 };
    for (const [k, v] of Object.entries(cases)) assert.strictEqual(C.parseRate(k), v, k);
    for (const k of ['Other', '', 'PRN']) assert.strictEqual(C.parseRate(k), null, k);
  });
  // REF like the clinic workbooks: row 1 filler, headers on row 2 from column B, lists below
  const ref = () => new Sheet('REF', { '0,0': 'xxx', '0,1': 'xxx', '1,1': 'Pharmacy', '1,2': 'Dose', '1,3': 'Treatment Frequency', '1,4': 'Status', '2,1': 'Besse', '2,2': 14, '2,3': '2x per week', '2,4': 'Active', '3,3': 'EOW', '4,3': 'Monthly', '5,3': 'Other' });
  await t('existing REF header: list replaced, old rows cleared, dropdown points at it (stop)', async () => {
    const ws = ref(); WB = { sheets: [ws], pt: new PT(['MRN', 'Patient Name', 'Status', 'Rendering Provider', 'Tx Frequency']) };
    const res = await C.OfficeStore().setDropdown('Tx Frequency', 'Treatment Frequency', ['2x per week', '1x per week', '1x every 3 weeks'], true);
    assert.deepStrictEqual([2, 3, 4, 5].map(r => at(ws, r, 3)), ['2x per week', '1x per week', '1x every 3 weeks', '']);
    assert.strictEqual(ws.fmt.get('2,3'), '@'); assert.strictEqual(at(ws, 2, 1), 'Besse', 'other REF columns untouched');
    assert.strictEqual(res.source, "='REF'!$D$3:$D$5");
    const dv = WB.pt.dv[4]; assert.strictEqual(dv.cleared, 1); assert.strictEqual(dv.rule.list.source, res.source); assert.strictEqual(dv.rule.list.inCellDropDown, true); assert.strictEqual(dv.errorAlert.style, 'Stop');
    assert(WB.pt.dv.every((d, i) => i === 4 || d.cleared === 0), 'only that column changed');
  });
  await t('new REF header goes after the last one; warning style for providers', async () => {
    const ws = ref(); WB = { sheets: [ws], pt: new PT(['MRN', 'Patient Name', 'Status', 'Rendering Provider', 'Tx Frequency']) };
    const res = await C.OfficeStore().setDropdown('Rendering Provider', 'Rendering Provider', ['Provider A', 'Provider B'], false);
    assert.strictEqual(at(ws, 1, 5), 'Rendering Provider'); assert.deepStrictEqual([at(ws, 2, 5), at(ws, 3, 5)], ['Provider A', 'Provider B']);
    assert.strictEqual(res.source, "='REF'!$F$3:$F$4"); assert.strictEqual(WB.pt.dv[3].errorAlert.style, 'Warning');
  });
  await t('no REF sheet: one is created', async () => {
    WB = { sheets: [], pt: new PT(['MRN', 'Tx Frequency']) };
    const res = await C.OfficeStore().setDropdown('Tx Frequency', 'Treatment Frequency', ['1x per week'], true);
    const ws = WB.sheets.find(s => s.name === 'REF'); assert(ws); assert.strictEqual(at(ws, 2, 3), '1x per week'); assert.strictEqual(res.source, "='REF'!$D$3:$D$3");
  });
  await t('missing column on the tracker stops', async () => {
    WB = { sheets: [ref()], pt: new PT(['MRN']) };
    await assert.rejects(C.OfficeStore().setDropdown('Tx Frequency', 'Treatment Frequency', ['x'], true), /no Tx Frequency column/);
  });
  console.log(`${n} passed`);
})().catch(e => { console.error(e); process.exit(1); });
