// OfficeStore.writePayerRules against a small in-memory mock of the Office.js range API.
//   node tests/payer-rules-excel-mock.test.js
// Checks: the first import creates a visible "Payer Credential Rules" sheet of plain text cells (no table) with
// PayerRules.HEADERS on row 1; text format goes on before the values; a second import replaces the contents (old rows
// beyond the new ones are cleared) and reads back the same rules. This is a mock, not Excel. MADE-UP payers only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const root = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const esc = [^\n]*$/m), line(/^const norm = [^\n]*$/m), line(/^const isF = [^\n]*$/m), line(/^const mrnKey = [^\n]*$/m), line(/^const same = [^\n]*$/m), line(/^const snapMatch = [^\n]*$/m),
  between('// ---------- Payer credential rules: who may bill', '// ---------- Payer credential rules: end'), between('function OfficeStore(){', '// ---------- store: demo'),
  'globalThis.OfficeStore = OfficeStore;'].join('\n');
const PayerRules = require(path.join(root, 'payer-rules.js'));

// ---- mock Excel ----
const L = [];
function Sheet(name){ this.name = name; this.cells = new Map(); this.fmt = new Map(); this.visibility = 'Hidden'; this.tables = { add: () => { L.push('tables.add'); } }; }
Sheet.prototype.getRange = function (a1){ const m = a1.match(/^([A-Z]+)(\d+)$/); const c = m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; return new Range(this, +m[2] - 1, c, 1, 1); };
Sheet.prototype.getUsedRangeOrNullObject = function (){ const ws = this; const used = [...ws.cells.entries()].filter(([, v]) => v !== '');
  return { isNullObject: !used.length, clear: k => { assert.strictEqual(k, 'Contents'); L.push('clear'); for (const [key] of used) ws.cells.set(key, ''); } }; };
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.format = { autofitColumns: () => L.push('autofit') }; }
Range.prototype.getResizedRange = function (dr, dc){ return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.getOffsetRange = function (dr, dc){ return new Range(this.ws, this.r0 + dr, this.c0 + dc, this.nr, this.nc); };
const shape = (r, a) => { assert.strictEqual(a.length, r.nr, 'row count'); a.forEach(x => assert.strictEqual(x.length, r.nc, 'col count')); };
Object.defineProperty(Range.prototype, 'values', { set(a){ shape(this, a); L.push('values'); a.forEach((row, i) => row.forEach((v, j) => { const k = `${this.r0 + i},${this.c0 + j}`; assert.strictEqual(this.ws.fmt.get(k), '@', 'text format before the value'); assert.strictEqual(typeof v, 'string'); this.ws.cells.set(k, v); })); } });
Object.defineProperty(Range.prototype, 'numberFormat', { set(a){ shape(this, a); L.push('numberFormat'); a.forEach((row, i) => row.forEach((v, j) => this.ws.fmt.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
let WB;
const ctx = () => ({ sync: async () => {}, workbook: { worksheets: { getItemOrNullObject: n => WB.sheets.find(s => s.name === n) || { isNullObject: true }, add: n => { const s = new Sheet(n); WB.sheets.push(s); L.push('sheet.add'); return s; } } } });
const C = { Excel: { run: fn => fn(ctx()), SheetVisibility: { visible: 'Visible' }, ClearApplyTo: { contents: 'Contents' } }, PayerRules, setTimeout };
vm.createContext(C); vm.runInContext(code, C);
const aoa = ws => { const R = Math.max(...[...ws.cells.keys()].map(k => +k.split(',')[0])) + 1, W = Math.max(...[...ws.cells.keys()].map(k => +k.split(',')[1])) + 1;
  return Array.from({ length: R }, (_, r) => Array.from({ length: W }, (_, c) => ws.cells.get(`${r},${c}`) ?? '')).filter(r => r.some(v => v !== '')); };

(async () => {
  WB = { sheets: [] };
  const store = C.OfficeStore();
  const rules = PayerRules.read([PayerRules.HEADERS,
    ['Acme Health', '', 'MD', 'Y', 'N/A', 'Yes', 'Yes', 'Yes', 'acme', ''], ['Acme Health', '', 'NP', 'Y', 'Yes', 'Yes', 'Yes', 'Yes', 'acme', ''],
    ['Orbit Federal', '', 'NP', '', 'No', 'No', 'No', 'Yes', 'orbit; orbit federal', 'NP may do observation only'], ['Zenith Blue CA', 'CA', 'NP', 'N', 'No', 'No', 'No', 'No', 'zenith', '']]).rules;
  // 1. first import: new visible sheet, plain cells, no table
  let res = await store.writePayerRules(PayerRules.toRows(rules));
  assert.strictEqual(res.rules, 4);
  const ws = WB.sheets.find(s => s.name === 'Payer Credential Rules'); assert(ws, 'sheet created'); assert.strictEqual(ws.visibility, 'Visible');
  assert(!L.includes('tables.add'), 'no table'); assert(L.indexOf('numberFormat') < L.indexOf('values'));
  const got = aoa(ws); assert.deepStrictEqual(got[0], PayerRules.HEADERS); assert.strictEqual(got.length, 5);
  assert.deepStrictEqual(got[3], ['Orbit Federal', '', 'NP', '', 'No', 'No', 'No', 'Yes', 'orbit; orbit federal', 'NP may do observation only']);
  assert.deepStrictEqual(PayerRules.read(got).rules, rules, 'reads back the same rules');
  console.log('ok first import writes the Payer Credential Rules sheet as plain text cells');
  // 2. second import with fewer rules: contents replaced, nothing left over, same sheet
  L.length = 0;
  res = await store.writePayerRules(PayerRules.toRows(rules.slice(0, 2)));
  assert.strictEqual(res.rules, 2); assert.strictEqual(WB.sheets.length, 1); assert(L.includes('clear') && !L.includes('sheet.add'));
  assert.strictEqual(aoa(ws).length, 3, 'old rows cleared');
  console.log('ok second import replaces the contents');
  // 3. nothing to write: no Excel call
  L.length = 0; res = await store.writePayerRules([]); assert.strictEqual(res.rules, 0); assert.strictEqual(L.length, 0);
  console.log('ok empty import writes nothing');
  console.log('3 payer rules sheet tests passed');
})().catch(e => { console.error(e); process.exit(1); });
