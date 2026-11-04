// Pharmacy transfers and refills: config, order days, transfer states and routing, the order plan, refills, the ESK Pharmacy sheet.
//   node tests/pharmacy.test.js
// Runs the pharmacy block of src/taskpane.html (plus the helpers it calls) in Node. FICTIONAL MRNs and short names only.
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
const a0 = app.indexOf('// ---------- pharmacy transfers and refills'), a1 = app.indexOf('// ---------- end pharmacy logic ----------');
assert(a0 > 0 && a1 > a0, 'pharmacy block markers');
const NEED = ['esc', 'MS', 'ymd', 'todaySerial', 'sDate', 'dowOf', 'DOW', 'fmtD', 'isoOf', 'fromIso', 'fmtT', 'mrnKey', 'pharmOf', 'PH_LABEL', 'PS', 'nowMin', 'pendingAppt',
  'parseRate', 'freqRate', 'fifo', 'genoaModel', 'apptInfo', 'nm'];

// pinned clock: Tue 2026-11-03 15:00 UTC (11/3 local from UTC-12 to UTC+8)
const NOW = Date.UTC(2026, 10, 3, 15, 0, 0);
class FixedDate extends Date { constructor(...a){ if (a.length) super(...a); else super(NOW); } static now(){ return NOW; } }
const ser = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 864e5) + 25569, T = ser(2026, 11, 3), iso = s => new Date((s - 25569) * 864e5).toISOString().slice(0, 10);

const CFG0 = { list: [
  { id: 'besse', name: 'Besse', match: 'besse', rank: 1, orderDow: 4, shipDays: 2, start: '2026-12-01' },
  { id: 'advantage', name: 'Advantage', match: 'advantage, adv rx', rank: 2, orderDow: 2, shipDays: 3, start: '2026-12-01' },
  { id: 'genoa', name: 'Genoa', match: 'genoa', rank: 3 } ] };
function make({ P = [], A = [], K = [], phx = new Map(), cfg = CFG0 } = {}){
  const C = { Date: FixedDate, console };
  vm.createContext(C);
  Object.assign(C, { S: { P, A, K, byMrn: new Map(P.map(p => [p.mrn, p])), phx, shared: { pharm: cfg }, freqRate: {}, win: { from: T, to: T + 30 } } });
  vm.runInContext(NEED.map(chunk).join('\n') + '\n' + app.slice(a0, a1) + '\nObject.assign(globalThis, { phCfg, phById, phMatch, phIsStock, phOrderBy, parsePhxRows, phxRow, mergePhx, phxCandidates, phxState, phxRoute, phxPlan, phxCadence, phxCounts, rxUsed, rxState, rxList, apptInfo, PHX_COLS, PHX_DEF });', C);
  return C;
}
const pt = (mrn, pharmName, o = {}) => ({ mrn, name: 'ZZT' + mrn.slice(-3), status: 'Active', pharm: /besse/i.test(pharmName) ? 'BESSE' : /genoa/i.test(pharmName) ? 'GENOA' : pharmName ? 'OTHER' : '', pharmName, mg: 84, freq: '1x per week', provider: 'Provider Alpha', insurance: 'Plan A', ...o });
const ap = (mrn, date, o = {}) => ({ key: `${mrn}|${date}|600`, mrn, date, time: 600, active: true, dose: 84, isStart: false, ...o });
const kit = (mrn, o = {}) => ({ mrn, pharm: 'GENOA', pharmRaw: 'Genoa', recv: 84, available: false, openOrder: false, administered: false, r: 1, ...o });
const deq = (a, b) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), JSON.parse(JSON.stringify(b)));   // vm arrays come from another realm
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('defaults: Besse, Advantage, Genoa in that order; only Besse is clinic stock', () => {
  const C = make({ cfg: undefined });
  deq(C.phCfg().list.map(p => [p.id, p.rank]), [['besse', 1], ['advantage', 2], ['genoa', 3]]);
  assert.strictEqual(C.phIsStock(C.phById('besse')), true); assert.strictEqual(C.phIsStock(C.phById('advantage')), false);
  assert.strictEqual(C.phCfg().refills, 3);
});
t('match words: comma-separated, case-insensitive, else nothing', () => {
  const C = make();
  assert.strictEqual(C.phMatch('ADV RX Specialty').id, 'advantage'); assert.strictEqual(C.phMatch('Genoa Healthcare').id, 'genoa');
  assert.strictEqual(C.phMatch('Other Specialty'), null); assert.strictEqual(C.phMatch(''), null);
});
t('order-by: back off shipping days, then to the order weekday', () => {
  const C = make(), dos = ser(2026, 12, 1);   // Tue
  assert.strictEqual(iso(C.phOrderBy(dos, C.phById('advantage'))), '2026-11-24');   // Tue order, 3 days → Sat 11/28 → Tue 11/24
  assert.strictEqual(iso(C.phOrderBy(dos, C.phById('besse'))), '2026-11-26');       // Thu order, 2 days → Sun 11/29 → Thu 11/26
  assert.strictEqual(iso(C.phOrderBy(ser(2026, 12, 3), C.phById('advantage'))), '2026-11-24');   // Thu visit: Mon 11/30 → Tue 11/24 (a Tue 12/1 order arrives Fri)
  assert.strictEqual(C.phOrderBy(dos, C.phById('genoa')), dos);   // no order day or shipping set: the visit date
});
t('states: not started → waiting → eligible → transferring / staying / not eligible → moved', () => {
  const p = pt('900201', 'Genoa');
  const st = rec => { const C = make({ P: [p], phx: new Map([['900201', rec]]) }); return C.phxState(p); };
  assert.strictEqual(make({ P: [p] }).phxState(p).st, 'new');
  assert.strictEqual(st({ xf: { resp: { besse: { sub: '2026-11-01', cls: 'pend' } } } }).st, 'wait');
  const e = st({ xf: { resp: { besse: { cls: 'nc' }, advantage: { cls: 'cov' } } } }); assert.strictEqual(e.st, 'decide'); assert.strictEqual(e.best.id, 'advantage');
  assert.strictEqual(st({ xf: { resp: { besse: { cls: 'cov' } }, on: true, to: 'besse' } }).label, 'Transferring to Besse · 12/1');
  assert.strictEqual(st({ xf: { resp: { besse: { cls: 'cov' } }, on: false } }).st, 'stay');
  assert.strictEqual(st({ xf: { resp: { besse: { cls: 'nc' }, advantage: { cls: 'oon' } } } }).st, 'none');
  assert.strictEqual(st({ xf: { on: true, to: 'besse', done: { at: '2026-12-01', to: 'besse' } } }).label, 'Moved to Besse');
});
t('better pharmacies only: an Advantage patient can only move to Besse; Genoa (current) is never asked', () => {
  const C = make({ P: [pt('900202', 'Advantage'), pt('900203', 'Genoa'), pt('900204', 'CVS Specialty')] });
  deq(C.phxState(C.S.P[0]).better.map(x => x.id), ['besse']);
  deq(C.phxState(C.S.P[1]).better.map(x => x.id), ['besse', 'advantage']);
  deq(C.phxState(C.S.P[2]).better.map(x => x.id), ['besse', 'advantage', 'genoa']);   // not on the list: anything ranked
});
t('candidates: specialty patients still in treatment, plus anyone with a record (moved to Besse)', () => {
  const P = [pt('900205', 'Genoa'), pt('900206', 'Besse'), pt('900207', 'Genoa', { status: 'Discontinued' }), pt('900208', 'Besse')];
  const C = make({ P, phx: new Map([['900208', { xf: { done: { at: '2026-12-01', to: 'besse' } } }]]) });
  deq(C.phxCandidates().map(p => p.mrn), ['900205', '900208']);
});
t('routing: from the start date only, Besse target counts as Besse, none once moved', () => {
  const p = pt('900209', 'Genoa'), q = pt('900210', 'Genoa');
  const C = make({ P: [p, q], phx: new Map([['900209', { xf: { on: true, to: 'besse' } }], ['900210', { xf: { on: true, to: 'advantage' } }]]) });
  assert.strictEqual(C.apptInfo(ap('900209', ser(2026, 11, 30))).pharm, 'GENOA');
  const b = C.apptInfo(ap('900209', ser(2026, 12, 1))); assert.strictEqual(b.pharm, 'BESSE'); assert.strictEqual(b.xfer, 'besse');
  const a = C.apptInfo(ap('900210', ser(2026, 12, 8))); assert.strictEqual(a.pharm, 'OTHER'); assert.strictEqual(a.phName, 'Advantage');
  C.S.phx.set('900209', { xf: { on: true, to: 'besse', done: { at: '2026-12-01' } } }); assert.strictEqual(C.apptInfo(ap('900209', ser(2026, 12, 1))).xfer, '');
  C.S.phx.set('900210', { xf: { on: false } }); assert.strictEqual(C.apptInfo(ap('900210', ser(2026, 12, 8))).pharm, 'GENOA');
});
t('order plan: visits on the new pharmacy grouped by order day; kits on hand go to the earliest visits first', () => {
  const p = pt('900211', 'Genoa'), dates = [ser(2026, 11, 24), ser(2026, 12, 1), ser(2026, 12, 3), ser(2026, 12, 8)];
  const C = make({ P: [p], A: dates.map(d => ap('900211', d)), K: [kit('900211', { available: true, full: true, received: T - 2, expDate: T + 40, r: 1 }), kit('900211', { available: true, full: true, received: T - 1, expDate: T + 40, r: 2 })], phx: new Map([['900211', { xf: { on: true, to: 'advantage' } }]]) });
  const plan = C.phxPlan(42); assert.strictEqual(plan.length, 1); assert.strictEqual(plan[0].ph.id, 'advantage');
  const g = plan[0].groups.map(x => [iso(x.date), x.rows.map(r => [iso(r.a.date), r.cover]), x.need.length]);
  // 11/24 is before the start date (Genoa still) and takes the first kit on hand; the second covers the first visit on Advantage
  deq(g, [['2026-11-24', [['2026-12-01', 'onhand'], ['2026-12-03', 'need']], 1], ['2026-12-01', [['2026-12-08', 'need']], 1]]);
  deq(plan[0].groups[0].mg, { 84: 1 });
});
t('order plan: Besse visits are counted as stock, and late order days are grouped as order now', () => {
  const cfg = { list: [{ ...CFG0.list[0], start: '2026-11-03' }, CFG0.list[1], CFG0.list[2]] };
  const C = make({ cfg, P: [pt('900212', 'Genoa')], A: [ap('900212', ser(2026, 11, 4)), ap('900212', ser(2026, 11, 11))], phx: new Map([['900212', { xf: { on: true, to: 'besse' } }]]) });
  const g = C.phxPlan(42)[0].groups;
  assert.strictEqual(g[0].date, T); assert.strictEqual(g[0].late, true); assert.strictEqual(g[0].rows[0].cover, 'stock');
  assert.strictEqual(iso(g[1].date), '2026-11-05');
});
t('open orders: while transferring, an order with the new pharmacy covers its visits, not an earlier one on the old pharmacy', () => {
  const p = pt('900231', 'Genoa'), A = [ap('900231', ser(2026, 11, 26)), ap('900231', ser(2026, 12, 3))];
  const K = [kit('900231', { openOrder: true, ordered: T, pharm: 'OTHER', pharmRaw: 'Advantage' })];
  const C = make({ P: [p], A, K, phx: new Map([['900231', { xf: { on: true, to: 'advantage' } }]]) });
  deq(C.genoaModel(T, T + 40).rows.map(r => [iso(r.a.date), r.cover]), [['2026-11-26', 'need'], ['2026-12-03', 'ordered']]);
  C.S.phx.delete('900231');   // not transferring: oldest first, as before
  deq(C.genoaModel(T, T + 40).rows.map(r => [iso(r.a.date), r.cover]), [['2026-11-26', 'ordered'], ['2026-12-03', 'need']]);
});
t('cadence: kits a week by target and dose; first order from the first booked visit there', () => {
  const P = [pt('900213', 'Genoa', { freq: '2x per week' }), pt('900214', 'Genoa', { freq: '1x every 2 weeks', mg: 56 }), pt('900215', 'Genoa', { freq: 'Other' })];
  const phx = new Map(P.map(p => [p.mrn, { xf: { on: true, to: 'advantage' } }]));
  const C = make({ P, A: [ap('900213', ser(2026, 12, 3))], phx });
  const c = C.phxCadence(); const w = c.weekly[0];
  assert.strictEqual(w.pts, 3); assert.strictEqual(w.mg[84], 2); assert.strictEqual(w.mg[56], 0.5); assert.strictEqual(w.unknown, 1);
  assert.strictEqual(iso(c.rows.find(r => r.p.mrn === '900213').firstOrder), '2026-11-24');
  assert.strictEqual(iso(c.rows.find(r => r.p.mrn === '900214').firstOrder), '2026-11-24');   // not booked: from the start date
});
t('refills: kits ordered since the Rx count against it; another configured pharmacy and Besse do not', () => {
  const p = pt('900216', 'Genoa');
  const K = [kit('900216', { ordered: T - 20 }), kit('900216', { ordered: T - 5 }), kit('900216', { received: T - 2 }), kit('900216', { ordered: T - 3, pharm: 'BESSE', pharmRaw: 'Besse' }), kit('900216', { ordered: T - 3, pharm: 'OTHER', pharmRaw: 'Advantage' })];
  const C = make({ P: [p], K });
  assert.strictEqual(C.rxUsed('900216', { ph: 'genoa', sent: iso(T - 10) }), 2);
  assert.strictEqual(C.rxUsed('900216', { ph: 'genoa', sent: iso(T - 10), adj: 1 }), 3);
  assert.strictEqual(C.rxUsed('900216', { sent: '' }), 0);
});
t('refills: the first visit the Rx cannot cover sets the request date (order day less lead days)', () => {
  const p = pt('900217', 'Genoa', { freq: '' });
  const A = [ser(2026, 11, 10), ser(2026, 11, 17), ser(2026, 11, 24), ser(2026, 12, 1)].map(d => ap('900217', d));
  const K = [kit('900217', { ordered: T - 1, openOrder: true })];   // on order: covers 11/10 and used 1 of 2
  const C = make({ P: [p], A, K, cfg: { list: [CFG0.list[0], CFG0.list[1], { ...CFG0.list[2], orderDow: 1, shipDays: 3 }] }, phx: new Map([['900217', { rx: { ph: 'genoa', sent: iso(T - 2), kits: 2 } }]]) });
  const r = C.rxState(p);
  assert.strictEqual(r.left, 1); assert.strictEqual(iso(r.outAt), '2026-11-24');   // 11/17 takes the last fill
  assert.strictEqual(iso(r.reqBy), '2026-11-09');   // Mon order 11/16 (11/24 − 3 = Sat 11/21 → Mon 11/16), less 7 days
  assert.strictEqual(r.st, 'soon');
});
t('refills: by frequency past the booked visits; request now when due; requested waits; no Rx flagged', () => {
  const p = pt('900218', 'Genoa', { freq: '1x per week' });
  const C = make({ P: [p], A: [ap('900218', T + 7)], phx: new Map([['900218', { rx: { ph: 'genoa', sent: iso(T - 1), kits: 2 } }]]) });
  const r = C.rxState(p); assert.strictEqual(r.projected, true); assert.strictEqual(iso(r.outAt), iso(T + 21)); assert.strictEqual(r.st, 'ok');   // T+7 booked, T+14 by frequency, T+21 has no fill
  C.S.phx.set('900218', { rx: { ph: 'genoa', sent: iso(T - 1), kits: 1 } }); assert.strictEqual(iso(C.rxState(p).reqBy), iso(T + 7)); assert.strictEqual(C.rxState(p).st, 'soon');
  C.S.A.push(ap('900218', T + 2)); assert.strictEqual(C.rxState(p).st, 'now');   // T+2 takes the only fill, T+7 has none: its order is already due
  C.S.phx.set('900218', { rx: { ph: 'genoa', sent: iso(T - 1), kits: 1, req: { at: iso(T), by: 'Tech' } } }); assert.strictEqual(C.rxState(p).st, 'req');
  C.S.phx.set('900218', { rx: { ph: 'genoa', sent: iso(T), kits: 1, req: { at: iso(T - 3) } } }); assert.strictEqual(C.rxState(p).st, 'now');   // request older than the new Rx: done with
  C.S.phx.delete('900218'); const x = C.rxState(p); assert.strictEqual(x.st, 'none'); assert.strictEqual(x.tone, 'warn');
});
t('refill list: per-patient pharmacies only, Besse transfers drop off once started, most urgent first', () => {
  const cfg = { list: [{ ...CFG0.list[0], start: '2026-11-01' }, CFG0.list[1], CFG0.list[2]] };
  const P = [pt('900219', 'Genoa'), pt('900220', 'Besse'), pt('900221', 'Genoa'), pt('900222', 'Advantage')];
  const C = make({ cfg, P, A: [ap('900222', T + 2)], phx: new Map([['900221', { xf: { on: true, to: 'besse' } }], ['900222', { rx: { ph: 'advantage', sent: iso(T - 30), kits: 1 } }]]),
    K: [kit('900222', { ordered: T - 20, pharm: 'OTHER', pharmRaw: 'Advantage' })] });
  deq(C.rxList().map(x => x.p.mrn), ['900222', '900219']);
});
t('refills: moving to another per-patient pharmacy, the old Rx covers only visits before the start; then a new Rx is due', () => {
  const p = pt('900229', 'Genoa');
  const A = [ser(2026, 11, 24), ser(2026, 12, 1), ser(2026, 12, 8)].map(d => ap('900229', d));
  const C = make({ P: [p], A, phx: new Map([['900229', { xf: { on: true, to: 'advantage', resp: { advantage: { cls: 'cov' } } }, rx: { ph: 'genoa', sent: iso(T - 5), kits: 4 } }]]) });
  const r = C.rxState(p); assert.strictEqual(r.st, 'newrx'); assert.strictEqual(r.label, 'Needs an Rx at Advantage');
  assert.strictEqual(iso(r.reqBy), '2026-11-17');   // first Advantage visit 12/1 → order Tue 11/24, less 7 days
  C.S.phx.get('900229').xf.resp.advantage.sub = iso(T); assert.strictEqual(C.rxState(p).tone, 'info');
  C.S.phx.get('900229').rx = { ph: 'advantage', sent: iso(T), kits: 4 }; assert.strictEqual(C.rxState(p).st, 'ok');   // the new Rx is tracked like any other
  C.S.phx.get('900229').rx = { ph: 'genoa', sent: iso(T - 5), kits: 0 }; assert.strictEqual(C.rxState(p).st, 'empty');   // runs out before the move (11/24): request by 11/17
  C.S.A.push(ap('900229', T + 2)); assert.strictEqual(C.rxState(p).st, 'now');
});
t('refills: moving to Besse, no projection past the start date', () => {
  const p = pt('900230', 'Genoa');
  const C = make({ P: [p], A: [ap('900230', ser(2026, 11, 10)), ap('900230', ser(2026, 12, 8))], phx: new Map([['900230', { xf: { on: true, to: 'besse' }, rx: { ph: 'genoa', sent: iso(T - 1), kits: 4 } }]]) });
  const r = C.rxState(p); assert.strictEqual(r.st, 'ok'); assert.strictEqual(r.label, 'Moves to Besse 12/1 · 4 of 4 left'); assert.strictEqual(r.reqBy, null);
  assert.strictEqual(r.visits.length, 1);   // the 12/8 visit comes from Besse stock
});
t('sheet: row text, JSON round trip, upsert by MRN keeps others, header kept', () => {
  const C = make({ P: [pt('900223', 'Genoa')] });
  const rec = { xf: { resp: { besse: { sub: '2026-11-01', cls: 'cov', at: '2026-11-02' } }, on: true, to: 'besse' }, rx: { ph: 'genoa', sent: '2026-10-20', kits: 4 } };
  const row = C.phxRow('900223', rec, { name: 'ZZT223', cur: 'Genoa', left: 2, now: '2026-11-03 09:00', by: 'Tech' });
  assert.strictEqual(row.length, C.PHX_COLS.length);
  deq(row.slice(0, 11), ['900223', 'ZZT223', 'Genoa', 'Besse', 'Besse: Eligible: covered', 'Besse 2026-11-01', 'Yes', '2026-10-20 Genoa', 4, 2, '']);
  let v = C.mergePhx([], '900223', row); v = C.mergePhx(v, '900224', C.phxRow('900224', { rx: { sent: '2026-11-01', kits: 4 } }, {}));
  v = C.mergePhx(v, '900223', C.phxRow('900223', { ...rec, xf: { ...rec.xf, on: false } }, {}));
  assert.strictEqual(v[0][0], 'MRN'); deq(v.slice(1).map(r => r[0]), ['900224', '900223']);
  const m = C.parsePhxRows(v); assert.strictEqual(m.get('900223').xf.on, false); assert.strictEqual(m.get('900224').rx.kits, 4);
  assert.strictEqual(C.parsePhxRows([['MRN', 'Data (JSON)'], ['900225', '{bad']]).size, 0);
});
t('sheet layout: #64 columns first and unchanged, so rows already on the live sheet still parse', () => {
  const C = make(), OLD = ['MRN', 'Patient', 'Current pharmacy', 'Transfer to', 'Benefit results', 'Rx submitted', 'Transfer', 'Rx/refill sent', 'Kits on Rx', 'Kits left', 'Refill requested', 'Updated', 'By', 'Data (JSON)'];
  deq(C.PHX_COLS.slice(0, OLD.length), OLD);
  const old = [OLD, ['900232', 'ZZT232', 'Genoa', '', '', '', '', '', '', '', '', '', '', JSON.stringify({ rx: { sent: '2026-10-01', kits: 4 } })]];
  assert.strictEqual(C.parsePhxRows(old).get('900232').rx.kits, 4);
  const v = C.mergePhx(old, '900233', C.phxRow('900233', { rx: { sent: '2026-11-01', kits: 2 } }, {}));
  deq(v[0], C.PHX_COLS); const m = C.parsePhxRows(v); assert.strictEqual(m.get('900232').rx.kits, 4); assert.strictEqual(m.get('900233').rx.kits, 2);
});
t('counts by benefit: per ranked pharmacy, who could move there and each response', () => {
  const P = [pt('900226', 'Genoa'), pt('900227', 'Genoa'), pt('900228', 'Advantage')];
  const C = make({ P, phx: new Map([['900226', { xf: { resp: { besse: { sub: '2026-11-01', cls: 'cov' } } } }], ['900227', { xf: { resp: { besse: { sub: '2026-11-01' } } } }]]) });
  const b = C.phxCounts().find(x => x.ph.id === 'besse'), a = C.phxCounts().find(x => x.ph.id === 'advantage');
  assert.strictEqual(b.n, 3); assert.strictEqual(b.sub, 2); assert.strictEqual(b.none, 2); assert.strictEqual(b.by.find(x => x.c.id === 'cov').n, 1);
  assert.strictEqual(a.n, 2);   // the Advantage patient isn't counted for Advantage
});
console.log(`${n} passed`);
