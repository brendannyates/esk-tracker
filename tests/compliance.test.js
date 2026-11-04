// REMS enrollment, prescribers and delegates, the weekly REMS call, insurance names, coverage changes and BIDF.
//   node tests/compliance.test.js
// Runs the compliance block of src/taskpane.html (plus the helpers it calls) in Node. FICTIONAL MRNs, names and IDs only.
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
const a0 = app.indexOf('// ---------- REMS enrollment, prescribers, BIDF'), a1 = app.indexOf('// ---------- end compliance logic ----------');
assert(a0 > 0 && a1 > a0, 'compliance block markers');
const NEED = ['esc', 'norm', 'blank', 'MS', 'ymd', 'todaySerial', 'sDate', 'dowOf', 'DOW', 'fmtD', 'isoOf', 'fromIso', 'toSerial', 'mrnKey', 'yes', 'canonProv', 'benKey', 'benOf', 'phxRec', 'Tbl'];

// pinned clock: Tue 2026-11-03 15:00 UTC
const NOW = Date.UTC(2026, 10, 3, 15, 0, 0);
class FixedDate extends Date { constructor(...a){ if (a.length) super(...a); else super(NOW); } static now(){ return NOW; } }
const ser = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 864e5) + 25569, T = ser(2026, 11, 3), iso = s => new Date((s - 25569) * 864e5).toISOString().slice(0, 10);
const REMS0 = { prescribers: [{ id: 'pa', name: 'Provider Alpha', cred: 'MD', delegates: ['Tech One'] }, { id: 'pb', name: 'Provider Beta', cred: 'MD', delegates: ['Tech One', 'Tech Two'] }],
  settings: [{ id: 'h1', name: 'Clinic (Alpha)', dea: 'pa', start: '2025-01-01', end: '2026-12-31' }, { id: 'h2', name: 'Clinic (Beta)', dea: 'pb', start: '2026-11-01' }], callDow: 2 };
function make({ P = [], phx = new Map(), ben = [], rems = REMS0, insNames, ptHeaders } = {}){
  const C = { Date: FixedDate, console }; vm.createContext(C);
  const H = ptHeaders || ['MRN', 'Patient Name', 'Insurance', '2026 BIDF'];
  const rows = P.map(p => H.map(h => ({ mrn: p.mrn, patientname: p.name, insurance: p.insurance, '2026bidf': p.bidf ?? '' })[h.toLowerCase().replace(/[^a-z0-9]/g, '')] ?? ''));
  vm.runInContext(NEED.map(chunk).join('\n') + '\n' + app.slice(a0, a1) + '\nObject.assign(globalThis, { cmCfg, insShort, insSame, sohSer, rxrByName, delegateOf, hcsCurrent, remsPts, remsState, reSet, reStart, reOpen, remsCall, bidfOf, covState, bidfState, bidfYearCol, Tbl, INS_DEF, REMS_STEPS });', C);
  const ptT = new C.Tbl('patients', 'Table1', 'Master Patient Tracker', H, rows, rows.map(r => r.map(() => '')));
  const benH = ['MRN', 'Patient Name', 'Sohar Payer', 'Plan', 'Sohar Date', 'Source', 'BI Completed', 'BI Insurance', 'BIDF Per Treatment'];
  const benM = new Map(ben.map(b => [String(b.mrn), Object.fromEntries(benH.map(h => [h.toLowerCase().replace(/[^a-z0-9]/g, ''), b[h] ?? null]))]));
  C.S = { P: P.map((p, i) => ({ r: i, status: 'Active', provider: 'Provider Alpha', ...p })), phx, ben: benM, tables: { patients: ptT }, shared: { rems, ...(insNames ? { insNames } : {}) } };
  C.S.byMrn = new Map(C.S.P.map(p => [p.mrn, p]));
  return C;
}
const deq = (a, b) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), JSON.parse(JSON.stringify(b)));
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('insurance names: full plan names map to the short insurer, first match wins, unknown stays as typed', () => {
  const C = make();
  assert.strictEqual(C.insShort('Blue Shield Blue Cross Platinum 70 PPO'), 'Blue Shield');
  assert.strictEqual(C.insShort('ANTHEM BLUE CROSS CA'), 'Anthem'); assert.strictEqual(C.insShort('Blue Cross of California'), 'Anthem');
  assert.strictEqual(C.insShort('Humana Medicare Advantage'), 'Humana'); assert.strictEqual(C.insShort('Medi-Cal Managed Care'), 'Medi-Cal');
  assert.strictEqual(C.insShort('  Acme Health  '), 'Acme Health'); assert.strictEqual(C.insShort(''), '');
  assert.strictEqual(C.insSame('UHC Choice Plus', 'United Healthcare'), true);
  const D = make({ insNames: [{ name: 'Acme', match: 'acme' }] }); assert.strictEqual(D.insShort('Acme Health'), 'Acme'); assert.strictEqual(D.insShort('Aetna PPO'), 'Aetna PPO');   // Admin's list replaces the defaults
});
t('prescribers and delegates: matched by tracker name; a technician can be a delegate for several prescribers', () => {
  const C = make();
  assert.strictEqual(C.rxrByName('provider beta').id, 'pb'); assert.strictEqual(C.rxrByName('Nobody'), null);
  deq(C.delegateOf('tech one').map(p => p.id), ['pa', 'pb']); deq(C.delegateOf('Tech Two').map(p => p.id), ['pb']); deq(C.delegateOf(''), []);
});
t('current healthcare setting: started, not ended, latest start', () => {
  assert.strictEqual(make().hcsCurrent().id, 'h2');
  assert.strictEqual(make({ rems: { ...REMS0, settings: [REMS0.settings[0], { ...REMS0.settings[1], start: '2026-12-01' }] } }).hcsCurrent().id, 'h1');   // not started yet
  assert.strictEqual(make({ rems: { ...REMS0, settings: [{ ...REMS0.settings[0], end: '2026-10-01' }] } }).hcsCurrent(), null);
});
t('REMS states: no setting, enrolled under the old one (re-enroll), re-enrolling, enrolled under the current one', () => {
  const P = [{ mrn: '900301' }, { mrn: '900302' }, { mrn: '900303' }, { mrn: '900304' }];
  const C = make({ P, phx: new Map([['900302', { rems: { hcs: 'h1' } }], ['900303', { rems: { hcs: 'h1', re: [{ id: 'x', to: 'h2', opened: '2026-11-02', prepared: '2026-11-02' }] } }], ['900304', { rems: { hcs: 'h2', enrolled: '2026-11-02' } }]]) });
  deq(C.S.P.map(p => C.remsState(p).st), ['unassigned', 'needre', 're', 'ok']);
  assert.strictEqual(C.remsState(C.S.P[2]).label, 'Re-enrolling under Clinic (Beta) · next: sent to patient');
  assert.strictEqual(make({ P, rems: { prescribers: [], settings: [] } }).remsState(C.S.P[0]).st, 'nocfg');
});
t('re-enrollment steps: later steps fill earlier ones; re-enrolled moves the patient; a new start cancels an unfinished one', () => {
  const C = make(); let rec = { rems: { hcs: 'h1' } };
  rec = C.reStart(rec, 'h2', '2026-11-03', 'Tech One'); assert.strictEqual(C.reOpen(rec.rems).to, 'h2'); assert.strictEqual(C.reOpen(rec.rems).from, 'h1');
  assert.strictEqual(C.reStart(rec, 'h2', '2026-11-04').rems.re.length, 1);   // already open for that setting
  C.reSet(rec, 'returned', '2026-11-05'); const x = C.reOpen(rec.rems); deq([x.prepared, x.sent, x.returned, x.submitted], ['2026-11-05', '2026-11-05', '2026-11-05', undefined]);
  C.reSet(rec, 'confirmed', '2026-11-08'); assert.strictEqual(rec.rems.hcs, 'h2'); assert.strictEqual(rec.rems.enrolled, '2026-11-08'); assert.strictEqual(C.reOpen(rec.rems), null);
  let r2 = C.reStart({ rems: { hcs: 'h1' } }, 'h2', '2026-11-03'); r2 = C.reStart(r2, 'h1', '2026-11-04'); assert.strictEqual(r2.rems.re[0].cx, '2026-11-04'); assert.strictEqual(C.reOpen(r2.rems).to, 'h1');
});
t('weekly REMS call: weeks run Sunday to Saturday; due from the call weekday; open actions until done', () => {
  const at = (d, o = {}) => ({ id: d, at: d, by: 'Tech One', outcome: 'none', ...o });
  assert.strictEqual(make().remsCall().st, 'due');   // Tue 11/3, calls on Tuesdays, none logged
  assert.strictEqual(make({ phx: new Map([['_rems', { calls: [at('2026-11-01')] }]]) }).remsCall().st, 'done');   // Sunday counts for this week
  assert.strictEqual(make({ phx: new Map([['_rems', { calls: [at('2026-10-31')] }]]) }).remsCall().st, 'due');    // last Saturday doesn't
  assert.strictEqual(make({ rems: { ...REMS0, callDow: 4 } }).remsCall().st, 'later');
  const k = make({ phx: new Map([['_rems', { calls: [at('2026-10-27', { outcome: 'actions', note: 'x' }), at('2026-10-20', { outcome: 'actions', done: '2026-10-22' })] }]]) }).remsCall();
  assert.strictEqual(k.open.length, 1); assert.strictEqual(k.calls[0].at, '2026-10-27');
});
t('coverage change: tracker insurance not matching Sohar flags it, until the tracker is updated', () => {
  const C = make({ P: [{ mrn: '900305', insurance: 'Aetna' }], ben: [{ mrn: '900305', 'Sohar Payer': 'Blue Shield of California', 'Sohar Date': '2026-10-30', Source: 'Verification' }] });
  const c = C.covState(C.S.P[0]); assert.strictEqual(c.mismatch, true); assert.strictEqual(c.redo, true); assert.strictEqual(c.so, 'Blue Shield');
  deq(c.why, ['Tracker says Aetna, Sohar 10/30 says Blue Shield']);
  C.S.P[0].insurance = 'Blue Shield Blue Cross Platinum 70 PPO'; assert.strictEqual(C.covState(C.S.P[0]).flag, false);   // same insurer, full plan name
});
t('coverage change: any Sohar discovery result means redo the BIDF, cleared by a BI after it', () => {
  const ben = { mrn: '900306', 'Sohar Payer': 'Cigna', 'Sohar Date': '2026-10-30 09:00', Source: 'Discovery' };
  const C = make({ P: [{ mrn: '900306', insurance: 'Cigna' }], ben: [ben] });
  const c = C.covState(C.S.P[0]); assert.strictEqual(c.mismatch, false); assert.strictEqual(c.redo, true); deq(c.why, ['Sohar discovery 10/30 found coverage']);
  const D = make({ P: [{ mrn: '900306', insurance: 'Cigna' }], ben: [{ ...ben, 'BI Completed': '2026-10-29', 'BI Insurance': 'Cigna' }] }); assert.strictEqual(D.covState(D.S.P[0]).redo, true);   // BI before the discovery
  const E = make({ P: [{ mrn: '900306', insurance: 'Cigna' }], ben: [{ ...ben, 'BI Completed': '2026-11-02', 'BI Insurance': 'Cigna' }] }); assert.strictEqual(E.covState(E.S.P[0]).flag, false);
});
t('coverage change: a BI done for another insurer than Sohar shows needs redoing even if it is newer', () => {
  const C = make({ P: [{ mrn: '900307', insurance: 'Anthem' }], ben: [{ mrn: '900307', 'Sohar Payer': 'Anthem Blue Cross', 'Sohar Date': '2026-10-01', Source: 'Verification', 'BI Completed': '2026-10-15', 'BI Insurance': 'Aetna' }] });
  const c = C.covState(C.S.P[0]); assert.strictEqual(c.mismatch, false); assert.strictEqual(c.redo, true); deq(c.why, ['BIDF was done for Aetna']);
  const D = make({ P: [{ mrn: '900308', insurance: 'Aetna' }] }); assert.strictEqual(D.covState(D.S.P[0]).flag, false);   // no Sohar result: nothing to compare
});
t('BIDF for the year: the tracker column says Yes, or a BI completed this year', () => {
  assert.strictEqual(make().bidfYearCol(), '2026 BIDF');
  const C = make({ P: [{ mrn: '900309', bidf: 'Yes' }, { mrn: '900310' }, { mrn: '900311' }], ben: [{ mrn: '900310', 'BI Completed': '2026-02-01', 'BIDF Per Treatment': '45' }, { mrn: '900311', 'BI Completed': '2025-12-15' }] });
  deq(C.S.P.map(p => C.bidfState(p).thisYear), [true, true, false]);
  assert.strictEqual(C.bidfState(C.S.P[1]).bi.perTx, '45'); assert.strictEqual(C.bidfState(C.S.P[0]).col, true);
  const D = make({ P: [{ mrn: '900312' }], ptHeaders: ['MRN', 'Patient Name', 'Insurance'] }); assert.strictEqual(D.bidfState(D.S.P[0]).col, false);
});
t('Sohar dates: serials, ISO with a time, US dates', () => {
  const C = make();
  assert.strictEqual(C.sohSer(ser(2026, 10, 30) + 0.4), ser(2026, 10, 30)); assert.strictEqual(C.sohSer('2026-10-30 09:00'), ser(2026, 10, 30));
  assert.strictEqual(C.sohSer('10/30/2026'), ser(2026, 10, 30)); assert.strictEqual(C.sohSer(''), null);
});
console.log(`${n} passed`);
