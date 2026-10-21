// Bay map on Today: bay matching and the colour each bay gets.
//   node tests/baymap.test.js
// Runs the bay map block of src/taskpane.html in Node with stubbed timers. FICTIONAL codes only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const a = src.indexOf('// ---------- bay map:'), b = src.indexOf("document.addEventListener('toggle', e => { const d = e.target; if (d.dataset?.sec === 'bmap'");
assert(a > 0 && b > a, 'bay map block markers');
const C = { fin: new Set() }; vm.createContext(C);
vm.runInContext(`const obsAt = v => v?.adm ?? null, v40At = v => v.adm + 40, disAt = o => o + 100, fmtT = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');
const vitGet = (k, w) => w === 'fin' ? fin.has(k) : w === 'mid' ? k.endsWith('m') : null, nm = m => m, esc = s => String(s), CFG = { bayMap: true };
` + src.slice(a, b) + '\nObject.assign(globalThis, { bayNo, bayState, bayMapModel, bayMapHTML });', C);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const row = (key, bay, grp, rank, extra = {}) => ({ key, grp, rank, a: { mrn: key.toUpperCase(), time: 600 }, i: { adminKit: {} }, v: { start: 600, adm: 605, bay }, ...extra });

t('bay numbers: "4", "Bay 4", " bay 4 " match; 7 and blank do not', () => {
  assert.strictEqual(C.bayNo('4'), '4'); assert.strictEqual(C.bayNo('Bay 4'), '4'); assert.strictEqual(C.bayNo(' bay 4 '), '4');
  assert.strictEqual(C.bayNo('7'), null); assert.strictEqual(C.bayNo(''), null); assert.strictEqual(C.bayNo(undefined), null);
});
t('states follow the Today ranks', () => {
  assert.strictEqual(C.bayState(row('a', '1', 'need', 0)).st, 'late');
  assert.strictEqual(C.bayState(row('b', '1', 'need', 1)).st, 'soon');
  assert.strictEqual(C.bayState(row('c', '1', 'need', 2, { i: { adminKit: null } })).s, 'Dose not given');
  assert.strictEqual(C.bayState(row('d', '1', 'need', 3)).s, 'Final vitals due');
  C.fin.add('e'); assert.strictEqual(C.bayState(row('e', '1', 'need', 3)).st, 'ready');
  assert.strictEqual(C.bayState(row('fm', '1', 'chair', 4)).s, 'Final vitals 11:45');
  assert.strictEqual(C.bayState(row('g', '1', 'chair', 4)).s, 'Vitals 10:45');
});
t('model: occupied, reserved, open and off-map', () => {
  const w = row('w', 'Bay 5', 'wait', 9); w.a.time = 690; w.v = { bay: 'Bay 5' };
  const m = { act: [1], need: [row('a', '2', 'need', 0), row('x', '9', 'need', 1)], chair: [row('c', '3', 'chair', 4), row('nb', '', 'chair', 4)], wait: [w] };
  const M = C.bayMapModel(m), by = Object.fromEntries(M.bays.map(x => [x.id, x]));
  assert.strictEqual(by['2'].st, 'late'); assert.strictEqual(by['3'].st, 'track'); assert.strictEqual(by['1'].s, 'Nobody booked');
  assert.strictEqual(by['5'].st, 'open'); assert.strictEqual(by['5'].s, 'W at 11:30');
  assert.strictEqual(M.off.map(r => r.key).join(), 'x,nb'); assert.strictEqual(M.used, 2); assert.strictEqual(M.late, 1);
  assert.strictEqual(by['2'].r, 60, 'radius survives (row kept separately)');
  const h = C.bayMapHTML(m, { key: 'a' });
  assert(!/NaN/.test(h), 'no NaN positions'); assert(/data-k="a" data-state="late" aria-current="true"/.test(h));
  assert(/data-k="w" data-state="open"/.test(h), 'reserved bay opens that visit'); assert(/Not on the map:/.test(h));
});
t('no visits today: no map', () => assert.strictEqual(C.bayMapHTML({ act: [], need: [], chair: [], wait: [] }), ''));
console.log(n + ' passed');
