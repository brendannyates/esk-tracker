// Visit log (hidden "ESK Visit Log" sheet): change rows, restore from snapshots, and when rows are cleared.
//   node tests/visit-log.test.js
// Runs the visit log block of src/taskpane.html in Node. FICTIONAL MRNs and staff only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const a = src.indexOf('// ---------- visit log:'), b = src.indexOf('async function vlogWrite');
assert(a > 0 && b > a, 'visit log block markers');
const code = ['const MS = 86400000;', line(/^const sDate = .*$/m), line(/^const isoOf = .*$/m), line(/^const hhmm = .*$/m), 'var CFG = { vitalsMin: 40 };', src.slice(a, b),
  'Object.assign(globalThis, { VLOG_COLS, vlogPick, vlogEvents, vlogRows, parseVlog, vlogRestore, vlogDrop });'].join('\n');
const C = {}; vm.createContext(C); vm.runInContext(code, C);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const D = 46315, K = `900101|${D}|600`, when = new Date(2026, 9, 20, 10, 5, 9);   // 46315 = 2026-10-20
const evs = rows => rows.map(r => r[4]).join(',');

t('header matches the sheet layout', () => assert.strictEqual(C.VLOG_COLS.join('|'), 'Logged|Visit|MRN|Appt date|Event|Time|BP|HR|O2|Detail|By|Snapshot'));
t('rooming with bay and initial vitals: one row per change, with the snapshot', () => {
  const cur = { [K]: { start: 600, bay: '2', vit: { init: { bp: '128/80', hr: '72', o2: '98', at: 600, by: 'Tech A' } }, note: 'not logged' } };
  const R = C.vlogRows({}, cur, when, 'Me', D - 45);
  assert.strictEqual(evs(R), 'Roomed,Bay,Initial vitals');
  assert.strictEqual(R[0][0], '2026-10-20 10:05:09'); assert.strictEqual(R[0][2], '900101'); assert.strictEqual(R[0][3], '2026-10-20');
  assert.strictEqual(R[2].slice(5, 11).join('|'), '10:00|128/80|72|98||Tech A'); assert.strictEqual(R[0][10], 'Me');
  assert(!JSON.parse(R[0][11]).note, 'Tx report text stays out of the snapshot');
});
t('only what changed is logged; no change, no rows', () => {
  const p = { [K]: { start: 600, vit: { init: { bp: '1', at: 600 } } } }, c = JSON.parse(JSON.stringify(p));
  assert.strictEqual(C.vlogRows(p, c, when, 'Me', D - 45).length, 0);
  c[K].vit.mid = { bp: '140/90', hr: '80', o2: '97', at: 645 }; c[K].adm = 605; c[K].note = 'edited';
  assert.strictEqual(evs(C.vlogRows(p, c, when, 'Me', D - 45)), 'Dose time,40-min vitals');
});
t('rechecks, PRN, discharge, move and clears', () => {
  const p = { [K]: { start: 600, rechk: [{ bp: '150/95', at: '10:50', why: 'BP' }], rescue: [] } };
  const c = { [K]: { rechk: [{ bp: '150/95', at: '10:50', why: 'BP' }, { bp: '138/88', at: '11:05', why: 'after nifedipine' }], rescue: [{ med: 'Nifedipine', dose: '10 mg', time: '10:52', reason: 'BP' }], out: { at: 720, by: 'Tech B' }, mv: { time: 630, prov: 'Provider B' } } };
  const R = C.vlogRows(p, c, when, 'Me', D - 45);
  assert.strictEqual(evs(R), 'Room time cleared,Vitals recheck,PRN,Discharged,Moved');
  assert.strictEqual(R[1][9], 'after nifedipine'); assert.strictEqual(R[2][9], 'Nifedipine 10 mg (BP)'); assert.strictEqual(R[3][10], 'Tech B');
});
t('provider dose confirmation and PRN orders: indicated, acknowledged, undone', () => {
  const o = { id: 'o1', med: 'Zofran', dose: '4 mg ODT', reason: 'nausea', by: 'Provider A', at: 620, ack: null };
  const p = { [K]: { start: 600 } }, c = { [K]: { start: 600, pc: { at: 605, sn: '1234', lot: 'AB12', by: 'Provider A' }, prnOrd: [o] } };
  let R = C.vlogRows(p, c, when, 'Me', D - 45);
  assert.strictEqual(evs(R), 'Provider confirmed dose,PRN indicated');
  assert.strictEqual(R[0][9], 'kit …1234 · lot AB12'); assert.strictEqual(R[0][10], 'Provider A');
  assert.strictEqual(R[1].slice(5, 11).join('|'), '10:20||||Zofran 4 mg ODT (nausea)|Provider A');
  const d = JSON.parse(JSON.stringify(c)); d[K].prnOrd[0].ack = { by: 'Tech A', time: '10:24', ts: 1 };
  R = C.vlogRows(c, d, when, 'Me', D - 45); assert.strictEqual(evs(R), 'PRN acknowledged'); assert.strictEqual(R[0][5] + '|' + R[0][10], '10:24|Tech A');
  assert.strictEqual(evs(C.vlogRows(d, c, when, 'Me', D - 45)), 'PRN acknowledgement undone');
  assert.deepStrictEqual(JSON.parse(R[0][11]).prnOrd[0].ack.by, 'Tech A', 'orders are in the snapshot, so a lost visit gets them back');
});
t('visits dropped by the 45-day cut are not logged as removed', () => {
  const old = `900101|${D - 60}|600`;
  assert.strictEqual(C.vlogRows({ [old]: { start: 600 } }, {}, when, 'Me', D - 45).length, 0);
});
t('restore fills blanks from the latest snapshot, never overwrites', () => {
  const rows = C.parseVlog([C.VLOG_COLS,
    ['2026-10-20 10:00:00', K, '900101', '', 'Roomed', '', '', '', '', '', '', JSON.stringify({ start: 600 })],
    ['2026-10-20 10:46:00', K, '900101', '', '40-min vitals', '', '', '', '', '', '', JSON.stringify({ start: 600, adm: 605, vit: { init: { bp: 'a' }, mid: { bp: 'b' } } })]]);
  const v = { [K]: { start: 610, vit: { init: { bp: 'mine' } } } };
  assert.strictEqual(C.vlogRestore(v, rows, D - 45), 1);
  assert.strictEqual(v[K].start, 610); assert.strictEqual(v[K].adm, 605); assert.strictEqual(v[K].vit.init.bp, 'mine'); assert.strictEqual(v[K].vit.mid.bp, 'b');
  const lost = {}; C.vlogRestore(lost, rows, D - 45); assert.strictEqual(lost[K].vit.mid.bp, 'b');
  assert.strictEqual(C.vlogRestore(v, rows, D - 45), 0, 'second pass has nothing to fill');
});
t('a cleared visit stays cleared (latest snapshot is empty)', () => {
  const rows = C.parseVlog([C.VLOG_COLS, ['2026-10-20 10:00:00', K, '900101', '', 'Roomed', '', '', '', '', '', '', '{"start":600}'], ['2026-10-20 10:01:00', K, '900101', '', 'Room time cleared', '', '', '', '', '', '', '{}']]);
  const v = {}; assert.strictEqual(C.vlogRestore(v, rows, D - 45), 0); assert(!v[K]);
});
t('rows stay until the next appointment day comes', () => {
  const rows = [{ key: K, mrn: '900101' }, { key: `900202|${D}|660`, mrn: '900202' }];
  const appts = [{ mrn: '900101', date: D }, { mrn: '900101', date: D + 3 }, { mrn: '900202', date: D + 1, cancelled: false }, { mrn: '900202', date: D + 1 - 1, cancelled: true }];
  assert.strictEqual(C.vlogDrop(rows, appts, D).length, 0, 'next visits are in the future');
  assert.strictEqual(C.vlogDrop(rows, appts, D + 1).join(), `900202|${D}|660`);
  assert.strictEqual(C.vlogDrop(rows, [{ mrn: '900101', date: D + 2, cancelled: true }], D + 5).length, 0, 'a cancelled visit does not clear it');
});
console.log(n + ' passed');
