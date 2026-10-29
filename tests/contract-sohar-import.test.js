// Contract test: the Sohar Eligibility add-in's "Tracker file (TMS + ESK)" CSV, read by ESK's importer.
//   node tests/contract-sohar-import.test.js
// Runs only the Node-tested "Sohar benefits" block of src/taskpane.html (plus norm and mrnKey), the same way
// tests/benefits.test.js does. No UI code is loaded. FICTIONAL names, MRNs, dates and amounts only.
//
// If the Sohar add-in folder is on this PC (SOHAR_DIR, or the default OneDrive path below; moved 10/28), the test also checks
// the literal column list against Sohar's live core.js and runs a file built by Sohar's own trackerExport + toCSV.
// If the folder is missing, those checks are skipped and the rest still runs.
const fs = require('fs'), os = require('os'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const a = src.indexOf('// ---------- Sohar benefits: parsing'), b = src.indexOf('// ---------- Sohar benefits: end');
assert(a > 0 && b > a, 'benefits block markers');
const code = [line(/^const norm = [^\n]*$/m), line(/^const mrnKey = [^\n]*$/m), src.slice(a, b),
  'Object.assign(globalThis, { norm, BEN_COLS, BEN_IN, benKey, benElig, parseCSV, sohTime, sohShort, parseBenefits, planBenefits, benRowOut, benLine });'].join('\n');
const E = {}; vm.createContext(E); vm.runInContext(code, E);
const eq = (x, y, msg) => assert.strictEqual(JSON.stringify(x), JSON.stringify(y), msg);
let n = 0, skipped = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

// Sohar core.js (Add-In Manifests/Sohar Eligibility/add-in, v0.9.1, d6d9e91):
//   BENEFITS_EXPORT_COLS (core.js ~line 30, "Treat as a contract: add columns at the end only")
//   TRACKER_EXPORT_COLS = [...BENEFITS_EXPORT_COLS, 'Service', 'Source', 'Exported', 'Secondary Payer', 'Cost-share Flags'] (~line 1471)
// Copied as a literal so this test runs without the Sohar folder. If Sohar appends columns, add them here.
// Sohar v0.19 (10/28) swapped Last Name, First Name and DOB for LASFIR: no names or DOBs in the file at all.
const TRACKER_EXPORT_COLS = ['Per Nbr', 'MRN', 'LASFIR', 'Sohar Payer', 'Plan', 'Network', 'Coverage Level',
  'Coverage Start', 'Coverage End', 'Copay', 'Coinsurance', 'Deductible', 'Deductible Remaining', 'Deductible Status',
  'OOP Max', 'OOP Remaining', 'OOP Status', 'Benefit Flag', 'Sohar Status', 'Sohar Date', 'Sohar Ref',
  'Service', 'Source', 'Exported', 'Secondary Payer', 'Cost-share Flags'];
// Columns ESK must never read: names and DOB by design; the rest are not used.
const IGNORED = ['LASFIR', 'Sohar Ref', 'Secondary Payer', 'Cost-share Flags'];
const PHI_COLS = ['Last Name', 'First Name', 'DOB'];

// Sohar's CSV writer (core.js csvCell/toCSV): header row, quote only when a cell has " , CR or LF, CRLF line ends plus a trailing CRLF, no BOM.
const cell = v => { const s = v == null ? '' : String(v); return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
const toCSV = (cols, rows) => [cols.join(','), ...rows.map(r => cols.map(c => cell(r[c])).join(','))].join('\r\n') + '\r\n';

// FICTIONAL tracker-file rows. Every value is distinct per column so a shifted column shows up.
const EXPORTED = '10/15/2026 9:00 AM';   // pane button format (fmtStamp)
const sohRow = (per, mrn, o = {}) => ({
  'Per Nbr': per, 'MRN': mrn, 'Last Name': o.last || 'Zzfakelast', 'First Name': o.first || 'Zzfakefirst', 'DOB': o.dob || '02/03/1971',
  'Sohar Payer': o.payer || 'Fictional Health', 'Plan': 'PPO, Choice', 'Network': 'In network', 'Coverage Level': 'Individual',
  'Coverage Start': '2026-01-01', 'Coverage End': '2026-12-31', 'Copay': o.copay ?? '30', 'Coinsurance': '20%', 'Deductible': '1500',
  'Deductible Remaining': '100', 'Deductible Status': 'Approaching', 'OOP Max': '4000', 'OOP Remaining': '2500', 'OOP Status': 'Not met',
  'Benefit Flag': 'Prior auth', 'Sohar Status': o.st || 'complete.eligible', 'Sohar Date': o.date || '2026-10-14T01:00:00.000Z',
  'Sohar Ref': 'ref-' + per, 'Service': o.svc ?? 'ESK', 'Source': 'Verification', 'Exported': EXPORTED,
  'Secondary Payer': 'Fictional Secondary', 'Cost-share Flags': 'QMB',
});
const FILE_ROWS = [
  sohRow('5001', '0900101', { last: 'Zzaldlast', first: 'Zzbenfirst', dob: '04/05/1966', date: '2026-10-02T15:30:00.000Z', copay: 'old' }),
  sohRow('5001', '0900101', { last: 'Zzaldlast', first: 'Zzbenfirst', dob: '04/05/1966', date: '2026-10-14T01:00:00.000Z', copay: '30' }),
  sohRow('5002', '900102', { last: 'Zzbrolast', first: 'Zzcamfirst', dob: '06/07/1977', st: 'complete.ineligible (INACTIVE)', svc: 'TMS; ESK', date: '10/13/2026 2:15 PM' }),
  sohRow('5003', '777001', { svc: 'TMS; ESK' }),   // ESK patient not on the tracker
  sohRow('5004', '777002', { svc: 'ESK' }),        // ESK patient not on the tracker
  sohRow('5005', '777003', { svc: 'TMS' }),        // TMS only, not on the tracker
  sohRow('5006', '777004', { svc: '' }),           // no upcoming event
];
const TRACKER = [{ mrn: '900101', name: 'ALDBEN' }, { mrn: '900102', name: 'BROCAM' }, { mrn: '900199', name: 'CARDEL' }];
const NOW = '2026-10-15 09:05';
const csvText = toCSV(TRACKER_EXPORT_COLS, FILE_ROWS);
const parsed = E.parseBenefits(E.parseCSV(csvText));
const plan = parsed && E.planBenefits(parsed.rows, TRACKER);
const fakePHI = FILE_ROWS.flatMap(r => PHI_COLS.map(c => r[c]));

t('every BEN_IN key is a Sohar tracker column; only LASFIR, Sohar Ref and the two newest columns go unread', () => {
  const have = TRACKER_EXPORT_COLS.map(E.norm);
  eq(E.BEN_IN.filter(k => !have.includes(k)), [], 'BEN_IN keys missing from the Sohar file');
  eq(TRACKER_EXPORT_COLS.filter(c => !E.BEN_IN.includes(E.norm(c))), IGNORED);
  assert.strictEqual(E.BEN_IN.length, 22);
});
t('ESK never reads name or DOB columns', () => {
  for (const c of PHI_COLS) assert(!E.BEN_IN.includes(E.norm(c)), c + ' is read');
  for (const c of PHI_COLS) assert(!E.BEN_COLS.map(E.norm).includes(E.norm(c)), c + ' is stored');
});
t('the tracker file is detected as the Sohar file', () => {
  assert(parsed, 'parseBenefits returned null for the Sohar tracker file');
  assert.strictEqual(parsed.rows.length, FILE_ROWS.length); assert.strictEqual(parsed.hasService, true); assert.strictEqual(parsed.exported, EXPORTED);
});
t('every read column lands on the right key (no shift, quoted comma kept)', () => {
  const o = parsed.rows[1], src = FILE_ROWS[1];
  for (const c of TRACKER_EXPORT_COLS) { const k = E.norm(c); if (E.BEN_IN.includes(k)) assert.strictEqual(o[k], src[c], c); }
  assert.strictEqual(o.plan, 'PPO, Choice');
});
t('leading-zero MRN survives the CSV, matches the tracker, and is written as text', () => {
  assert.strictEqual(parsed.rows[0].mrn, '0900101');
  const m = plan.matched.find(x => x.p.name === 'ALDBEN'); assert(m, 'ALDBEN matched');
  assert.strictEqual(m.row.mrn, '0900101');
  const out = E.benRowOut(m.row, m.p, NOW); assert.strictEqual(out[0], '0900101'); assert.strictEqual(typeof out[0], 'string');
});
t('ISO Sohar Date: read, latest per MRN wins, mixed with m/d/y', () => {
  assert.strictEqual(E.sohTime('2026-10-14T01:00:00.000Z'), Date.UTC(2026, 9, 14, 1, 0, 0));
  assert(E.sohTime('2026-10-14T01:00:00.000Z') > E.sohTime('10/13/2026 2:15 PM'));
  assert(E.sohTime('2026-10-02T15:30:00.000Z') < E.sohTime('10/13/2026 2:15 PM'));
  const m = plan.matched.find(x => x.p.name === 'ALDBEN');
  assert.strictEqual(m.row.sohardate, '2026-10-14T01:00:00.000Z'); assert.strictEqual(m.row.copay, '30');
  assert.strictEqual(E.sohShort('2026-10-14T01:00:00.000Z'), '10/14');
  assert(E.benLine(m.row).endsWith('Sohar 10/14'));
  assert.strictEqual(plan.lines, FILE_ROWS.length); assert.strictEqual(plan.people, FILE_ROWS.length - 1);
});
t('ESK tag from Service: off-tracker ESK rows are listed, not imported; others are counted as skipped', () => {
  eq(plan.matched.map(x => x.p.name), ['ALDBEN', 'BROCAM']);
  eq([...plan.eskMissing].sort(), ['777001', '777002']);
  assert.strictEqual(plan.skipped, 2);
  const imported = plan.matched.map(x => E.benRowOut(x.row, x.p, NOW)[0]);
  for (const mrn of plan.eskMissing) assert(!imported.includes(mrn), mrn + ' was imported');
  assert(!/\besk\b/i.test('TMS') && /\besk\b/i.test('TMS; ESK'));
});
t('not eligible status is flagged', () => {
  eq(plan.notElig.map(x => x.p.name), ['BROCAM']);
});
t('tblBenefits row: BEN_COLS order, values from the same-named Sohar column, LASFIR name, no names or DOB', () => {
  for (const m of plan.matched){
    const out = E.benRowOut(m.row, m.p, NOW), src = FILE_ROWS.filter(r => E.benKey(r.MRN) === m.row.key).pop();
    assert.strictEqual(out.length, E.BEN_COLS.length); assert(out.every(v => typeof v === 'string'));
    E.BEN_COLS.forEach((c, i) => {
      if (c === 'Patient Name') assert.strictEqual(out[i], m.p.name);
      else if (c === 'Imported') assert.strictEqual(out[i], NOW);
      else { assert(TRACKER_EXPORT_COLS.includes(c), c + ' is not a Sohar column'); if (c !== 'MRN') assert.strictEqual(out[i], src[c], c); }
    });
    for (const v of fakePHI) assert(!out.includes(v), 'stored ' + v);
    assert(!out.some(v => fakePHI.some(p => v.includes(p))), 'a name or DOB leaked into the row');
  }
});
t('parsed rows hold no name or DOB anywhere', () => {
  for (const o of parsed.rows){
    for (const k of ['lastname', 'firstname', 'dob']) assert(!(k in o), k + ' kept');
    for (const v of Object.values(o)) assert(!fakePHI.includes(v), 'kept ' + v);
  }
});

// ---- optional: Sohar's live core.js ----
const SOHAR = process.env.SOHAR_DIR || path.join(os.homedir(), 'OneDrive - Mindful Health Solutions', 'Projects', 'Clinic Workbook Pilot', 'add-in');
const coreFile = path.join(SOHAR, 'core.js');
if (!fs.existsSync(coreFile)){
  skipped = 3; console.log('skip live Sohar checks (no ' + coreFile + ')');
} else {
  const C = require(coreFile);
  t('live Sohar: TRACKER_EXPORT_COLS starts with this literal (Sohar adds at the end only)', () => {
    assert(Array.isArray(C.TRACKER_EXPORT_COLS), 'core.js no longer exports TRACKER_EXPORT_COLS');
    eq(C.TRACKER_EXPORT_COLS.slice(0, TRACKER_EXPORT_COLS.length), TRACKER_EXPORT_COLS, 'Sohar tracker columns changed');
    const extra = C.TRACKER_EXPORT_COLS.slice(TRACKER_EXPORT_COLS.length);
    if (extra.length) console.log('   note: Sohar added columns ' + JSON.stringify(extra) + '; add them to the literal');
    eq(E.BEN_IN.filter(k => !C.TRACKER_EXPORT_COLS.map(E.norm).includes(k)), []);
  });
  t('live Sohar: serviceLine tags esketamine visits ESK', () => {
    assert.strictEqual(C.serviceLine('10/20 Esketamine Follow Up'), 'ESK');
    assert.strictEqual(C.serviceLine('TMS Treatment 10/20 Spravato'), 'TMS; ESK');
    assert.strictEqual(C.serviceLine('TMS Treatment'), 'TMS');
  });
  t('live Sohar: a file from trackerExport + toCSV imports the same way', () => {
    const res = (per, mrn, o = {}) => ({ 'Sent As': 'Verification', 'Per Nbr': per, 'MRN': mrn, 'Sohar Ref': 'ref-' + per, 'Sohar Date': o.date || '2026-10-14T01:00:00.000Z',
      'Next Event': o.ev || 'Esketamine Follow Up', 'Upcoming Detail': '', 'Last Name': 'Zzfakelast', 'First Name': 'Zzfakefirst', 'DOB': '02/03/1971',
      'Sohar Payer': 'Fictional Health', 'Copay': '25', 'Deductible Remaining': '100', 'OOP Remaining': '2500', 'Sohar Status': o.st || 'complete.eligible' });
    const out = C.trackerExport([res('5001', '0900101'), res('5002', '900102', { st: 'complete.ineligible' }), res('5003', '777001', { ev: 'TMS Treatment; Esketamine Follow Up' }), res('5005', '777003', { ev: 'TMS Treatment' })], EXPORTED, new Map());
    const p = E.parseBenefits(E.parseCSV(C.toCSV(C.TRACKER_EXPORT_COLS, out.rows)));
    assert(p, 'ESK did not detect the Sohar file'); assert.strictEqual(p.hasService, true); assert.strictEqual(p.exported, EXPORTED);
    const pl = E.planBenefits(p.rows, TRACKER);
    eq(pl.matched.map(x => x.p.name), ['ALDBEN', 'BROCAM']); eq([...pl.eskMissing], ['777001']); assert.strictEqual(pl.skipped, 1);
    eq(pl.notElig.map(x => x.p.name), ['BROCAM']);
    const row = E.benRowOut(pl.matched[0].row, pl.matched[0].p, NOW);
    assert.strictEqual(row[0], '0900101'); assert.strictEqual(row[E.BEN_COLS.indexOf('Copay')], '25');
    assert(!row.some(v => /Zzfake|02\/03\/1971/.test(v)), 'name or DOB stored');
  });
}
console.log(`\n${n} passed${skipped ? `, ${skipped} skipped` : ''}`);
