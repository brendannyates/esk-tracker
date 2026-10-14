// Sohar benefits import: parsing, latest-per-MRN, matching, merge and card line.
//   node tests/benefits.test.js
// Runs the "Sohar benefits" block of src/taskpane.html in Node. FICTIONAL names and MRNs only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const a = src.indexOf('// ---------- Sohar benefits: parsing'), b = src.indexOf('// ---------- Sohar benefits: end');
assert(a > 0 && b > a, 'benefits block markers');
const code = [line(/^const norm = .*$/m), line(/^const mrnKey = .*$/m), src.slice(a, b),
  'Object.assign(globalThis, { BEN_COLS, BEN_IN, benKey, benElig, parseCSV, sohTime, sohShort, parseBenefits, planBenefits, benRowOut, mergeBenefits, benLine });'].join('\n');
const C = {}; vm.createContext(C); vm.runInContext(code, C);
const eq = (x, y) => assert.strictEqual(JSON.stringify(x), JSON.stringify(y));
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

const HDR = 'Per Nbr,MRN,Last Name,First Name,DOB,Sohar Payer,Plan,Network,Coverage Level,Coverage Start,Coverage End,Copay,Coinsurance,Deductible,Deductible Remaining,Deductible Status,OOP Max,OOP Remaining,OOP Status,Benefit Flag,Sohar Status,Sohar Date,Sohar Ref,Service,Source,Exported';
const row = (mrn, last, date, o = {}) => [o.per || '1', mrn, last, 'Test', '01/01/1980', o.payer || 'Aetna', 'PPO', 'In network', 'Individual', '2026-01-01', '2026-12-31', o.copay ?? '30', '20%', '500', o.ded ?? '100', o.dst || 'Approaching', '4000', o.oop ?? '2500', o.ost || 'Not met', '', o.st || 'complete.eligible', date, 'REF1', o.svc ?? 'ESK', o.src || 'Verification', '2026-10-14 08:00'];
const csv = rows => [HDR, ...rows.map(r => r.map(v => /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v).join(','))].join('\r\n') + '\r\n';
const patients = [{ mrn: '900101', name: 'ALDBEN' }, { mrn: '900102', name: 'BROCAM' }, { mrn: '00900103', name: 'CARDEL' }];

t('CSV parser: BOM, quotes, commas, CRLF', () => {
  const r = C.parseCSV('﻿a,"b, c","d ""q"""\r\n1,,3\n');
  eq(JSON.parse(JSON.stringify(r)), [['a', 'b, c', 'd "q"'], ['1', '', '3']]);
});
t('header found below a title row; ignores case/spaces/punctuation', () => {
  const p = C.parseBenefits([['Sohar export'], [], ['mrn', 'OOP-remaining', 'sohar_status'], ['0123', '50', 'complete.eligible']]);
  assert.strictEqual(p.rows.length, 1); assert.strictEqual(p.rows[0].mrn, '0123'); assert.strictEqual(p.rows[0].oopremaining, '50'); assert.strictEqual(p.hasService, false);
});
t('not the Sohar file → null', () => { assert.strictEqual(C.parseBenefits([['MRN', 'Appt Dt'], ['1', '2']]), null); });
t('benKey ignores case, punctuation, leading zeros', () => {
  assert.strictEqual(C.benKey('0090-0103'), C.benKey(900103)); assert.strictEqual(C.benKey(' ab-12 '), C.benKey('AB12')); assert.strictEqual(C.benKey('900101.0'), '900101');
});
t('latest row per MRN by Sohar Date; ties go to the later line', () => {
  const p = C.parseBenefits(C.parseCSV(csv([row('900101', 'X', '2026-10-01', { copay: 'old' }), row('900101', 'X', '2026-10-10', { copay: 'new' }), row('900101', 'X', '2026-09-01', { copay: 'older' }),
    row('900102', 'Y', '10/5/2026', { copay: 'a' }), row('900102', 'Y', '10/5/2026', { copay: 'b' })])));
  const plan = C.planBenefits(p.rows, patients);
  assert.strictEqual(plan.matched.find(x => x.p.name === 'ALDBEN').row.copay, 'new');
  assert.strictEqual(plan.matched.find(x => x.p.name === 'BROCAM').row.copay, 'b');
  assert.strictEqual(plan.lines, 5); assert.strictEqual(plan.people, 2);
});
t('matching, ESK not on tracker, skipped, not eligible', () => {
  const p = C.parseBenefits(C.parseCSV(csv([
    row('900101', 'X', '2026-10-10'),
    row('900103', 'Z', '2026-10-10', { st: 'complete.ineligible (INACTIVE)' }),   // tracker has 00900103
    row('777001', 'Q', '2026-10-10', { svc: 'TMS; ESK' }),
    row('777002', 'Q', '2026-10-10', { svc: 'TMS' }),
    row('777003', 'Q', '2026-10-10', { svc: '' }),
  ])));
  const plan = C.planBenefits(p.rows, patients);
  eq(plan.matched.map(x => x.p.name), ['ALDBEN', 'CARDEL']);
  eq([...plan.eskMissing], ['777001']);
  assert.strictEqual(plan.skipped, 2);
  eq(plan.notElig.map(x => x.p.name), ['CARDEL']);
  assert.strictEqual(p.hasService, true); assert.strictEqual(p.exported, '2026-10-14 08:00');
});
t('older file without Service/Source/Exported still works', () => {
  const h = HDR.split(',').slice(0, 23), r = row('900102', 'Y', '2026-10-10').slice(0, 23);
  const p = C.parseBenefits([h, r]); const plan = C.planBenefits(p.rows, patients);
  assert.strictEqual(plan.matched.length, 1); assert.strictEqual(plan.matched[0].row.source, ''); assert.strictEqual(p.hasService, false);
});
t('output row: BEN_COLS order, text, tracker name, no file name/DOB', () => {
  const p = C.parseBenefits(C.parseCSV(csv([row('0900101', 'SECRETLAST', '2026-10-10')])));
  const out = C.benRowOut(p.rows[0], { name: 'ALDBEN' }, '2026-10-14 09:00');
  assert.strictEqual(out.length, C.BEN_COLS.length); assert(out.every(v => typeof v === 'string'));
  assert.strictEqual(out[0], '0900101'); assert.strictEqual(out[1], 'ALDBEN'); assert.strictEqual(out[C.BEN_COLS.indexOf('Copay')], '30');
  assert.strictEqual(out[C.BEN_COLS.indexOf('Imported')], '2026-10-14 09:00'); assert(!out.includes('SECRETLAST') && !out.includes('01/01/1980'));
});
t('merge: replace in place, keep others, fill blank rows, then append', () => {
  const W = C.BEN_COLS.length, r = (m, x) => [m, x, ...Array(W - 2).fill('')];
  const ex = [r(900101, 'old'), r('', ''), r('900200', 'keep')].map(v => v.map(String)); ex[1] = Array(W).fill('');
  const m = C.mergeBenefits(ex, [r('0900101', 'new'), r('900300', 'a'), r('900400', 'b')]);
  eq(m.values.map(v => v[1]), ['new', 'a', 'keep', 'b']); assert.strictEqual(m.updated, 1); assert.strictEqual(m.added, 2);
});
t('card line', () => {
  assert.strictEqual(C.benLine({ copay: '30', coinsurance: '20%', deductibleremaining: '100', deductiblestatus: 'Approaching', oopremaining: '2500', oopstatus: 'Not met', sohardate: '2026-10-15' }),
    'Copay 30 · Coins 20% · Ded left 100 (Approaching) · OOP left 2500 · Sohar 10/15');
  assert.strictEqual(C.benLine({ copay: '0', oopremaining: '0', oopstatus: 'Met', sohardate: '10/3/2026 9:00 AM' }), 'Copay 0 · OOP left 0 (Met) · Sohar 10/3');
});
t('eligibility', () => {
  assert(C.benElig('complete.eligible')); assert(!C.benElig('complete.ineligible (INACTIVE)')); assert(!C.benElig('')); assert(!C.benElig('pending'));
});
t('sohTime orders m/d/y with times and ISO', () => {
  assert(C.sohTime('10/5/2026 1:00 PM') > C.sohTime('10/5/2026 11:00 AM')); assert(C.sohTime('2026-10-06') > C.sohTime('10/5/2026 11:59 PM')); assert.strictEqual(C.sohTime('junk'), null);
});
console.log(`\n${n} tests passed`);
