// REMS monitoring forms in the chart: Document Management report parsing, LASFIR, merge, visit matching and owners.
//   node tests/rems.test.js
// Runs the REMS block of src/taskpane.html in Node. FICTIONAL names, MRNs and staff only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const a = src.indexOf('// ---------- REMS monitoring forms in the chart'), b = src.indexOf('async function readRemsFile');
assert(a > 0 && b > a, 'REMS block markers');
const d0 = src.indexOf('const norm = h =>'), d1 = src.indexOf('function toMinutes');
const code = [src.slice(d0, d1), line(/^const mrnKey = .*$/m), 'var S = {};', src.slice(a, b),
  'Object.assign(globalThis, { S, REMS_COLS, lasfirOf, remsWhen, parseRemsReport, parseRemsRows, mergeRems, remsOpps, remsModel, ymd, todaySerial, isoOf });'].join('\n');
const C = {}; vm.createContext(C); vm.runInContext(code, C);
const eq = (x, y) => assert.strictEqual(JSON.stringify(x), JSON.stringify(y));
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

// Same layout as the NextGen export: title rows, a "From … to …" line, headers on row 7, blank columns after.
const HDR = ['Doc Type', 'Bch Nbr', 'Bch Name', 'Bch Status', 'Doc Nbr', 'Pages', 'Pat Name', 'Doc Crt Dt', 'Doc Crt By', 'Scnd Date', 'Scnd By', 'Post Date', 'Pstd By', 'Doc Mod Dt', 'Doc Mod By', 'Sign Off Stat', '', ''];
const top = [['', ' ', 'Clinic'], [], ['', '', 'Patient By Document Type'], ['', '', 'From 10/1/2026 to 10/31/2026'], ['10/20/26 12:22 PM'], []];
const r = (type, nbr, name, crt, scn = crt) => [type, '9001', 'QuickScan', 'Closed', nbr, '2', name, crt, 'Tech, Alpha', scn, 'Tech, Alpha', scn, 'Tech, Alpha', scn, 'Tech, Alpha', '', '', ''];
const MON = 'REMS PT Monitoring Form';

t('LASFIR from "Last, First", "First Last", and a code', () => {
  assert.strictEqual(C.lasfirOf('Testerson, Jane'), 'TESJAN');
  assert.strictEqual(C.lasfirOf('Jane Q Testerson'), 'TESJAN');
  assert.strictEqual(C.lasfirOf('De La Cruz, Ana Maria'), 'DELANA');
  assert.strictEqual(C.lasfirOf('aldben'), 'ALDBEN');
  assert.strictEqual(C.lasfirOf(''), '');
});
t('NextGen times: "12:13 P", "9:05 A", "12:05 A", serials', () => {
  assert.strictEqual(C.remsWhen('10/07/2026 12:13 P'), '2026-10-07 12:13');
  assert.strictEqual(C.remsWhen('10/12/2026 9:05 A'), '2026-10-12 09:05');
  assert.strictEqual(C.remsWhen('10/12/2026 12:05 A'), '2026-10-12 00:05');
  assert.strictEqual(C.remsWhen('10/7/26 4:15 PM'), '2026-10-07 16:15');
  assert.strictEqual(C.remsWhen(C.ymd(2026, 10, 7) + 0.5), '2026-10-07 12:00');
  assert.strictEqual(C.remsWhen('n/a'), '');
});
t('report: header row found by name, range read, footers and repeat headers skipped, no full names kept', () => {
  const aoa = [...top, HDR, r(MON, '1001', 'Testerson, Jane', '10/07/2026 12:13 P'), r('REMS PT Enrollment Form', '1002', '', '10/14/2026 4:45 P'), r(MON, '1003', '', '10/19/2026 3:57 P'), ['', '', ''], HDR, ['Total: 3']];
  const p = C.parseRemsReport(aoa, '2026-10-20 17:00');
  assert.strictEqual(p.n, 3); assert.strictEqual(p.mon, 2); assert.strictEqual(p.named, 1); assert.strictEqual(p.hasMrn, false);
  assert.strictEqual(p.from, C.ymd(2026, 10, 1)); assert.strictEqual(p.to, C.ymd(2026, 10, 31));
  eq([...p.rows[0]], ['1001', MON, '', 'TESJAN', '2026-10-07 12:13', 'Tech, Alpha', '2026-10-07 12:13', 'Tech, Alpha', '2026-10-07 12:13', 'Tech, Alpha', '2026-10-20 17:00']);
  assert(!JSON.stringify(p.rows).includes('Testerson'));
});
t('columns by header, not position; an MRN column wins', () => {
  const H = ['Scnd Date', 'MRN', 'Doc Type', 'Doc Nbr'], aoa = [H, ['10/08/2026 1:00 P', '000123', MON, '77']];
  const p = C.parseRemsReport(aoa, 'now'); assert.strictEqual(p.hasMrn, true); assert.strictEqual(p.rows[0][2], '000123'); assert.strictEqual(p.rows[0][6], '2026-10-08 13:00');
  assert.strictEqual(p.from, C.ymd(2026, 10, 8));   // no "From … to …" line: range = the document dates
});
t('not this report → null', () => { assert.strictEqual(C.parseRemsReport([['MRN', 'Appt Dt'], ['1', '2']], ''), null); assert.strictEqual(C.parseRemsReport([['Doc Type', 'Pages'], [MON, '2']], ''), null); });
t('merge by Doc Nbr; sheet round trip with reordered columns', () => {
  const m = C.mergeRems([['1', MON, '', 'AAABBB', '2026-10-01', '', '', '', '', '', 'x']], [['1', MON, '', 'AAABBB', '2026-10-02', '', '', '', '', '', 'y'], ['2', MON, '', 'CCCDDD', '2026-10-03', '', '', '', '', '', 'y']]);
  assert.strictEqual(m.added, 1); assert.strictEqual(m.updated, 1); assert.strictEqual(m.values[0][4], '2026-10-02');
  const H = [...C.REMS_COLS].reverse(), rows = C.parseRemsRows([H, ...m.values.map(v => [...v].reverse())]);
  assert.strictEqual(rows.length, 2); assert.strictEqual(rows[1].las, 'CCCDDD'); assert.strictEqual(rows[1].day, C.ymd(2026, 10, 3)); assert.strictEqual(rows[1].mon, true);
});

// Visits: kept appointments + administered kits. Owner = Treating Technician on that day's kit.
const D = (d) => C.ymd(2026, 10, d), rec = (nbr, las, day, o = {}) => ({ nbr, type: o.type || MON, mrn: o.mrn || '', las, created: C.isoOf(D(day)) + ' 15:00', scanned: '', day: D(day), mon: !o.type || /monitor/i.test(o.type) });
function world(recs, cov = [{ from: D(1), to: D(31), day: '2026-10-20' }]){
  const P = [{ mrn: '900101', name: 'ALDBEN' }, { mrn: '900102', name: 'BROCAM' }, { mrn: '900103', name: 'CARDEL' }];
  Object.assign(C.S, { byMrn: new Map(P.map(p => [p.mrn, p])), rems: recs, remsCov: cov,
    A: [{ key: 'a1', mrn: '900101', date: D(6), time: 600, status: 'Kept', patient: 'Aldridge, Ben' }, { key: 'a2', mrn: '900101', date: D(9), time: 600, status: 'Kept' },
      { key: 'a3', mrn: '900102', date: D(6), time: 660, status: 'Kept' }, { key: 'a4', mrn: '900103', date: D(6), time: 720, status: 'No Show' },
      { key: 'a5', mrn: '900103', date: D(7), time: 720, status: 'Kept', cancelled: true }, { key: 'a6', mrn: '900102', date: D(20), time: 600, status: 'Expected' }],
    K: [{ mrn: '900101', administered: true, adminDate: D(6), tech: 'Tech One' }, { mrn: '900102', administered: true, adminDate: D(6), tech: 'Tech  Two ' },
      { mrn: '900102', administered: true, adminDate: D(20), tech: 'Tech One' }, { mrn: '900103', administered: false, adminDate: null }] });
}
t('visits: kept + administered today, one per patient-day; no-shows and cancelled skipped; owner from the kit', () => {
  world([]); const o = C.remsOpps(D(1), D(31));
  eq(o.map(x => x.key), ['900101|' + D(6), '900102|' + D(6), '900101|' + D(9), '900102|' + D(20)]);
  eq(o.map(x => x.owner), ['Tech One', 'Tech Two', 'Not recorded', 'Tech One']);
  assert(o[0].las.has('ALDBEN'));
});
t('same day, late, outstanding; a form is used once; earlier forms never count', () => {
  world([rec('1', 'ALDBEN', 6), rec('2', 'ALDBEN', 11), rec('3', 'BROCAM', 5), rec('4', 'ALDBEN', 6, { type: 'REMS PT Enrollment Form' })]);
  const m = C.remsModel(D(1), D(20)), s = Object.fromEntries(m.opps.map(o => [o.key, o.status]));
  assert.strictEqual(s['900101|' + D(6)], 'same'); assert.strictEqual(s['900101|' + D(9)], 'late'); assert.strictEqual(m.opps.find(o => o.status === 'late').lateDays, 2);
  assert.strictEqual(s['900102|' + D(6)], 'missing'); assert.strictEqual(s['900102|' + D(20)], 'missing');
  assert.strictEqual(m.tot.n, 4); assert.strictEqual(m.tot.same, 1); assert.strictEqual(m.tot.missing, 2); assert.strictEqual(m.tot.rate, 0.25);
  const two = m.owners.find(x => x.name === 'Tech Two'); assert.strictEqual(two.c.missing, 1); assert.strictEqual(m.owners[0].c.missing, 1);
});
t('MRN on the form beats the name', () => {
  world([rec('1', 'ZZZZZZ', 6, { mrn: '900102' })]); const m = C.remsModel(D(6), D(6));
  assert.strictEqual(m.opps.find(o => o.mrn === '900102').status, 'same'); assert.strictEqual(m.opps.find(o => o.mrn === '900101').status, 'missing');
});
t('names blank: unnamed forms cover a day by count, a shortfall stays outstanding', () => {
  world([rec('1', '', 6), rec('2', '', 6)]); let m = C.remsModel(D(6), D(6));
  eq(m.opps.map(o => o.status), ['count', 'count']); assert.strictEqual(m.tot.rate, 1);
  world([rec('1', '', 6)]); m = C.remsModel(D(6), D(6));
  eq(m.opps.map(o => o.status), ['missing', 'missing']); assert.strictEqual(m.days[0].short, 1); assert.strictEqual(m.days[0].unnamed, 1);
});
t('days no report covers yet are "not checked", not outstanding', () => {
  world([], [{ from: D(1), to: D(31), day: '2026-10-08' }]); const m = C.remsModel(D(1), D(20)), s = m.opps.map(o => o.status);
  eq(s, ['missing', 'missing', 'unchecked', 'unchecked']); assert.strictEqual(m.tot.n, 2); assert.strictEqual(m.tot.unchecked, 2);
});
console.log(`\n${n} passed`);
