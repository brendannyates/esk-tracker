// OfficeStore.appendVisitLog / pruneVisitLog against a small in-memory mock of the Office.js range API.
//   node tests/visit-log-excel-mock.test.js
// Checks: the first write creates the hidden "ESK Visit Log" sheet with the header; later writes append under the last row;
// cells are text ('@') so times and MRNs stay as typed; prune removes only the given visits and keeps the header.
// This is a mock, not Excel: it can't catch Excel-on-the-web quirks. FICTIONAL MRNs only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const VLOG_SHEET = .*$/m), line(/^const CH = .*$/m), between('function OfficeStore(){', '// ---------- store: demo'), 'globalThis.OfficeStore = OfficeStore; globalThis.VLOG_COLS = VLOG_COLS;'].join('\n');

function Sheet(name){ this.name = name; this.cells = new Map(); this.fmt = new Map(); this.visibility = 'Visible'; }
const used = ws => { let r = -1, c = -1; for (const [k, v] of ws.cells) if (v !== '') { const [i, j] = k.split(',').map(Number); r = Math.max(r, i); c = Math.max(c, j); } return [r, c]; };
Sheet.prototype.getRange = function (a1){ const m = a1.match(/^([A-Z]+)(\d+)$/); return new Range(this, +m[2] - 1, m[1].charCodeAt(0) - 65, 1, 1); };
Sheet.prototype.getUsedRangeOrNullObject = function (){ const [r, c] = used(this); if (r < 0) return { isNullObject: true, load(){} }; const g = new Range(this, 0, 0, r + 1, c + 1); g.rowCount = r + 1; return g; };
function Range(ws, r0, c0, nr, nc){ this.ws = ws; this.r0 = r0; this.c0 = c0; this.nr = nr; this.nc = nc; this.isNullObject = false; }
Range.prototype.getResizedRange = function (dr, dc){ return new Range(this.ws, this.r0, this.c0, this.nr + dr, this.nc + dc); };
Range.prototype.getOffsetRange = function (dr, dc){ return new Range(this.ws, this.r0 + dr, this.c0 + dc, this.nr, this.nc); };
Range.prototype.load = function (){ this._v = Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this.ws.cells.get(`${this.r0 + i},${this.c0 + j}`) ?? '')); };
Range.prototype.clear = function (){ for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.ws.cells.set(`${this.r0 + i},${this.c0 + j}`, ''); };
const shape = (r, a) => { assert.strictEqual(a.length, r.nr, 'row count'); a.forEach(x => assert.strictEqual(x.length, r.nc, 'col count')); };
Object.defineProperty(Range.prototype, 'values', { get(){ return this._v; }, set(a){ shape(this, a); a.forEach((row, i) => row.forEach((v, j) => { const k = `${this.r0 + i},${this.c0 + j}`; assert.strictEqual(this.ws.fmt.get(k), '@', 'text format before values'); this.ws.cells.set(k, v); })); } });
Object.defineProperty(Range.prototype, 'numberFormat', { set(a){ shape(this, a); a.forEach((row, i) => row.forEach((v, j) => this.ws.fmt.set(`${this.r0 + i},${this.c0 + j}`, v))); } });
let WB;
const ctx = () => ({ sync: async () => {}, workbook: { worksheets: { getItemOrNullObject: n => WB.find(s => s.name === n) || { isNullObject: true }, add: n => { const s = new Sheet(n); WB.push(s); return s; } } } });
const C = { Excel: { run: fn => fn(ctx()), SheetVisibility: { hidden: 'Hidden' }, ClearApplyTo: { contents: 'Contents' } } };
vm.createContext(C); vm.runInContext(code, C);
const W = 12, row = (key, ev) => { const a = Array(W).fill(''); a[0] = '2026-10-20 10:00:00'; a[1] = key; a[2] = key.split('|')[0]; a[4] = ev; a[5] = '09:05'; return a; };
const col = (ws, j) => { const [r] = used(ws); return Array.from({ length: r + 1 }, (_, i) => ws.cells.get(`${i},${j}`)); };

(async () => {
  WB = []; const store = C.OfficeStore();
  await store.appendVisitLog([row('0900101|46315|600', 'Roomed'), row('0900101|46315|600', 'Initial vitals')]);
  const ws = WB.find(s => s.name === 'ESK Visit Log'); assert(ws, 'sheet created'); assert.strictEqual(ws.visibility, 'Hidden');
  assert.strictEqual(col(ws, 0)[0], 'Logged'); assert.strictEqual(col(ws, 4).join(), 'Event,Roomed,Initial vitals'); assert.strictEqual(ws.cells.get('1,2'), '0900101');
  console.log('ok first write creates the hidden sheet with the header');
  await store.appendVisitLog([row('900202|46315|660', 'Roomed')]); await store.appendVisitLog([]);
  assert.strictEqual(col(ws, 4).join(), 'Event,Roomed,Initial vitals,Roomed'); assert.strictEqual(WB.length, 1);
  console.log('ok later writes append under the last row');
  assert.strictEqual(await store.pruneVisitLog(['0900101|46315|600']), 2);
  assert.strictEqual(col(ws, 1).join(), 'Visit,900202|46315|660'); assert.strictEqual(col(ws, 5)[1], '09:05');
  assert.strictEqual(await store.pruneVisitLog(['nothing']), 0);
  await store.appendVisitLog([row('900303|46316|600', 'Roomed')]); assert.strictEqual(col(ws, 1).join(), 'Visit,900202|46315|660,900303|46316|600');
  console.log('ok prune removes only those visits, keeps the header, appends continue after it');
  console.log('3 passed');
})().catch(e => { console.error(e); process.exit(1); });
