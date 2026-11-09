// Besse weekly order: which weeks an order date counts, late orders, the history row it logs to, and the weekly reminder.
//   node tests/besse-order.test.js
// Runs the ordering functions of src/taskpane.html in Node. FICTIONAL MRNs only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const html = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const app = html.slice(html.indexOf("<script>\n'use strict';"), html.lastIndexOf('</script>'));
const BOUND = /^(async function |function |const |let |var |class |document\.|window\.|setInterval|\(function|\/\/ -{5,}|\/\/ ={5,}|Object\.assign\()/;
const CH = new Map(); { let cur = null;
  for (const l of app.split('\n')){
    if (BOUND.test(l)){ const m = /^(?:async )?function\s+([A-Za-z0-9_$]+)/.exec(l) || /^(?:const|let|var|class)\s+([A-Za-z0-9_$]+)/.exec(l); cur = m && !CH.has(m[1]) ? (CH.set(m[1], []), CH.get(m[1])) : null; }
    if (cur) cur.push(l);
  } }
const chunk = n => { assert(CH.has(n), 'top-level definition not found: ' + n); return CH.get(n).join('\n'); };
const NEED = ['MS', 'ymd', 'todaySerial', 'sDate', 'dowOf', 'toSerial', 'pharmOf', 'PHX_BLANK', 'PHX_DEF', '_phc', 'phCfg', 'phIsStock',
  'besseOrderDow', 'weeksFor', 'besseNeeds', 'besseModel', 'findHistRow', 'invBesseDone'];

const ser = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 864e5) + 25569, N = d => ser(2026, 11, d);
const iso = s => new Date((s - 25569) * 864e5).toISOString().slice(0, 10);
function make({ today = 9, A = [], K = [], orderDate, besse = null, pharm = null, besseWk = '' } = {}){
  const NOW = Date.UTC(2026, 10, today, 15, 0, 0);   // 15:00 UTC: the same local date from UTC-12 to UTC+8
  class FixedDate extends Date { constructor(...a){ if (a.length) super(...a); else super(NOW); } static now(){ return NOW; } }
  const C = { Date: FixedDate, console };
  vm.createContext(C);
  Object.assign(C, { S: { A, K, orderDate: orderDate ?? N(today), besse, shared: pharm ? { pharm } : {}, ov: { 56: 0, 84: 0 }, besseWk },
    apptInfo: a => ({ pharm: 'BESSE', mg: a.dose }) });
  vm.runInContext(NEED.map(chunk).join('\n') + '\nObject.assign(globalThis, { besseOrderDow, weeksFor, besseModel, findHistRow, invBesseDone });', C);
  return C;
}
const wk = w => [0, 1].map(i => [iso(w[i][0]), iso(w[i][1])]);
const ap = (mrn, d, time = 600, o = {}) => ({ key: `${mrn}|${N(d)}|${time}`, mrn, date: N(d), time, active: true, dose: 84, isStart: false, ...o });
const kit = (mrn, o = {}) => ({ mrn, pharm: 'BESSE', recv: 84, full: true, expired: false, expDate: null, administered: false, ...o });
const withDow = d => ({ list: [{ id: 'besse', name: 'Besse', match: 'besse', rank: 1, orderDow: d }] });
const hist = rows => ({ hist: { hdr: 0, c: 0 }, vals: [['DATE', 56, 84], ...rows] });
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('Friday order (the usual day) counts next Mon–Fri, then the week after', () => {
  const C = make({ today: 6 }); const w = C.weeksFor(N(6));
  assert.deepStrictEqual(wk(w), [['2026-11-09', '2026-11-13'], ['2026-11-16', '2026-11-20']]); assert.strictEqual(w.late, false);
});
t('weekend orders count the coming week', () => {
  const C = make({ today: 7 });
  for (const d of [7, 8]) assert.deepStrictEqual(wk(C.weeksFor(N(d))), [['2026-11-09', '2026-11-13'], ['2026-11-16', '2026-11-20']]);
});
t('late Monday 11/9 order: next week counts from 11/9, 2 weeks out is 11/16–11/20', () => {
  const C = make({ today: 9 }); const w = C.weeksFor(N(9));
  assert.deepStrictEqual(wk(w), [['2026-11-09', '2026-11-13'], ['2026-11-16', '2026-11-20']]);
  assert.strictEqual(w.late, true); assert.strictEqual(iso(w.due), '2026-11-06');
});
t('late Wednesday order counts Wed–Fri', () => {
  const C = make({ today: 11 }); assert.deepStrictEqual(wk(C.weeksFor(N(11))), [['2026-11-11', '2026-11-13'], ['2026-11-16', '2026-11-20']]);
});
t('"It’s for next week" switches a late-looking date back to an early order', () => {
  const C = make({ today: 9, besseWk: 'next' }); const w = C.weeksFor(N(9));
  assert.deepStrictEqual(wk(w), [['2026-11-16', '2026-11-20'], ['2026-11-23', '2026-11-27']]); assert.strictEqual(w.late, false); assert.strictEqual(w.lateable, true);
});
t('Monday order day set in Admin: a Monday order is on time (old behavior)', () => {
  const C = make({ today: 9, pharm: withDow(1) }); const w = C.weeksFor(N(9));
  assert.deepStrictEqual(wk(w), [['2026-11-16', '2026-11-20'], ['2026-11-23', '2026-11-27']]); assert.strictEqual(w.lateable, false);
});
t('Thursday order day: Thursday is on time, Monday is late', () => {
  const C = make({ today: 9, pharm: withDow(4) });
  assert.strictEqual(C.weeksFor(N(5)).late, false); assert.deepStrictEqual(wk(C.weeksFor(N(5)))[0], ['2026-11-09', '2026-11-13']);
  assert.strictEqual(C.weeksFor(N(9)).late, true); assert.strictEqual(iso(C.weeksFor(N(9)).due), '2026-11-05');
});
t('order day: Admin setting, else the history’s usual weekday, else Friday', () => {
  assert.strictEqual(make({ pharm: withDow(2) }).besseOrderDow(), 2);
  assert.strictEqual(make({ besse: hist([[N(5), 1, 2], [ser(2026, 10, 29), 0, 3], [N(6), '', '']]) }).besseOrderDow(), 4);   // Thursdays with quantities
  assert.strictEqual(make({ besse: hist([[N(7), 1, 2], [N(8), 1, 1]]) }).besseOrderDow(), 5);   // weekend labels: fall back
  assert.strictEqual(make().besseOrderDow(), 5);
});
t('late order counts only what is still to come, and orders the same as an on-time order would', () => {
  const A = [ap('9000001', 6), ap('9000002', 9, 480), ap('9000003', 9, 510), ap('9000004', 11), ap('9000005', 16), ap('9000006', 23)];
  // on time, Fri 11/6: 9000003's kit is still on the shelf
  const on = make({ today: 6, A, K: [kit('9000003'), kit('')] }).besseModel();
  // late, Mon 11/9 10:00: 9000003 got that kit at 8:30; 9000002 (8:00) is in the chair, kit not logged yet
  const late = make({ today: 9, A, K: [kit('9000003', { full: false, administered: true, adminDate: N(9) }), kit('')] }).besseModel();
  assert.deepStrictEqual([on.rows[84].w1, on.rows[84].w2, on.rows[84].oh], [3, 1, 2]);
  assert.deepStrictEqual([late.rows[84].w1, late.rows[84].w2, late.rows[84].oh], [2, 1, 1]);
  assert.strictEqual(late.rows[84].order, on.rows[84].order); assert.strictEqual(late.rows[84].order, 2);
  assert.deepStrictEqual([...late.list.map(x => x.a.mrn)], ['9000002', '9000004', '9000005']);
});
t('history row: a late 11/9 order logs on the 11/6 row, not 11/13', () => {
  const besse = hist([[ser(2026, 10, 30), 1, 2], [N(6), '', ''], [N(13), '', '']]);
  assert.strictEqual(iso(make({ today: 9, besse }).findHistRow().d), '2026-11-06');
  assert.strictEqual(iso(make({ today: 6, besse }).findHistRow().d), '2026-11-06');
  assert.strictEqual(iso(make({ today: 7, besse }).findHistRow().d), '2026-11-06');   // Saturday catch-up for the same week
  assert.strictEqual(iso(make({ today: 13, besse }).findHistRow().d), '2026-11-13');
});
t('weekly reminder: a late Monday order counts as done that week, not for Friday’s', () => {
  const besse = hist([[N(6), 1, 3], [N(13), '', '']]); const K = [kit('', { full: false, ordered: N(9), comments: 'Besse 84 mg order' })];
  assert.strictEqual(make({ today: 10, besse, K }).invBesseDone().done, true);
  assert.strictEqual(make({ today: 13, besse, K }).invBesseDone().done, false);
});
console.log(`\n${n} passed`);
