// Clinic Template Tool contract: the access-model export (Access tab > Export for template tool, mhs-access-model v1).
//   node tests/contract-access-export.test.js
// Runs the real ESK code (accessModel -> accessBook -> accXlsx) in a Node vm on a small FICTIONAL clinic with the clock pinned,
// writes the workbook with ESK's own SheetJS, reads it back the way the Clinic Template Tool does, and checks the contract:
// sheet names and order, every header row, the About keys and fixed values, Modality first, Sunday week starts,
// patient rows off by default and the " - PHI" file name rule.
// If the Clinic Template Tool's own copy is on this PC (OneDrive), its SheetJS + engine.js readAccessModel read the file too; skipped otherwise.
// Only protected (non-UI) top-level code is extracted, so the redesign can't break this test without breaking the contract.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');

// ---- top-level chunks of the app script, split the same way the redesign's protected-code check splits them ----
const app = html.slice(html.indexOf("<script>\n'use strict';"), html.lastIndexOf('</script>'));
assert(app.length > 1000, 'app script not found');
const BOUND = /^(async function |function |const |let |var |class |document\.|window\.|setInterval|\(function|\/\/ -{5,}|\/\/ ={5,}|Object\.assign\()/;
const CH = new Map(); { let cur = null;
  for (const l of app.split('\n')){
    if (BOUND.test(l)){ const m = /^(?:async )?function\s+([A-Za-z0-9_$]+)/.exec(l) || /^(?:const|let|var|class)\s+([A-Za-z0-9_$]+)/.exec(l); cur = m && !CH.has(m[1]) ? (CH.set(m[1], []), CH.get(m[1])) : null; }
    if (cur) cur.push(l);
  } }
const chunk = n => { assert(CH.has(n), 'top-level definition not found: ' + n); return CH.get(n).join('\n'); };
const NEED = ['MS', 'ymd', 'todaySerial', 'sDate', 'dowOf', 'DOW', 'fmtD', 'isoOf', 'esc', 'IND', 'suggestDays', 'apptInfo', 'inductionModel', 'credLabel',
  'MODALITY', 'parseRate', 'freqRate', 'RATE_LABEL', 'fmtN', 'stKey', 'accStatuses', 'accDefaultSt', 'CAP0', 'capOf', 'BLOCKS', 'chairBlock', 'provCredOf',
  'tmplOf', 'tmplSum', 'wkStart', 'slotOf', 'modeOf', 'accessModel', 'keptModel', 'DN5', 'ACCESS_SCHEMA', 'accessBook', 'accXlsx'];

// ---- pinned clock: Thu 2026-10-15 17:00 UTC (still 10/15 local anywhere from UTC-12 to UTC+6) ----
const NOW = Date.UTC(2026, 9, 15, 17, 0, 0);
class FixedDate extends Date { constructor(...a){ if (a.length) super(...a); else super(NOW); } static now(){ return NOW; } }

// ---- FICTIONAL clinic: made-up codes and MRNs, no real people ----
const ser = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 864e5) + 25569, T = ser(2026, 10, 15);
const P = [
  { mrn: '900001', name: 'ZZTAAA', status: 'Active', provider: 'Provider Alpha', freq: '2x per week', coverage: 'MD Only', pharm: 'BESSE', mg: 84 },
  { mrn: '900002', name: 'ZZTBBB', status: 'Active', provider: 'Provider Alpha', freq: '1x per week', coverage: 'MD or NP', pharm: 'GENOA', mg: 84 },
  { mrn: '900003', name: 'ZZTCCC', status: 'Active', provider: 'Provider Beta', freq: '1x every 2 weeks', coverage: 'MD Only', pharm: 'GENOA', mg: 84 },
  { mrn: '900004', name: 'ZZTDDD', status: 'Paused', provider: 'Provider Beta', freq: '1x every 2 weeks', coverage: '', pharm: 'OTHER', mg: 56 },
  { mrn: '900005', name: 'ZZTEEE', status: 'Active', provider: 'Provider Beta', freq: '1x every 2 weeks', coverage: 'MD & NP', pharm: 'BESSE', mg: 84 },
  { mrn: '900006', name: 'ZZTFFF', status: 'Active', provider: '', freq: 'Monthly', coverage: '', pharm: '', mg: 84 },
  { mrn: '900007', name: 'ZZTGGG', status: 'Pending Start', provider: 'Provider Alpha', freq: '2x per week', coverage: 'MD Only', pharm: 'BESSE', mg: null },
  { mrn: '900008', name: 'ZZTHHH', status: 'Discontinued', provider: 'Provider Alpha', freq: '1x per week', coverage: 'MD Only', pharm: 'BESSE', mg: 84 },
  { mrn: '900009', name: 'ZZTIII', status: 'Active', provider: 'Provider Alpha', freq: 'PRN', coverage: 'MD Only', pharm: 'GENOA', mg: 84 },
];
const A = [];
const ap = (mrn, date, time, o) => A.push({ mrn, date, time, active: true, isStart: false, dose: null, status: date < T ? 'Kept' : 'Expected', ...o });
for (const w of [-2, -1, 0, 1, 2, 3]){ ap('900001', ser(2026, 10, 12) + w * 7, 540); ap('900001', ser(2026, 10, 15) + w * 7, 600); ap('900002', ser(2026, 10, 13) + w * 7, 780); }
for (const w of [-2, 0, 2]){ ap('900003', ser(2026, 10, 14) + w * 7, 600); ap('900005', ser(2026, 10, 16) + w * 7, 840); }
ap('900004', ser(2026, 10, 7), 600);
ap('900006', ser(2026, 9, 24), 900); ap('900006', ser(2026, 10, 22), 900);
ap('900007', ser(2026, 10, 19), 540, { isStart: true, dose: 56 }); ap('900007', ser(2026, 10, 22), 540, { dose: 84 });
ap('900008', ser(2026, 10, 20), 600);
ap('900003', ser(2026, 10, 21), 600, { active: false });
const PHI = P.flatMap(p => [p.mrn, p.name]);

function run({ clinic = 'TST', checkbox = null, acc } = {}){
  const out = {};
  const C = { console, Date: FixedDate, Math, JSON, setTimeout, clearTimeout, ArrayBuffer, Uint8Array };
  C.window = C; C.self = C; vm.createContext(C);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'assets', 'xlsx.full.min.js'), 'utf8'), C);
  Object.assign(C, {
    S: { P: P.map(p => ({ ...p })), A: A.map(a => ({ ...a })), ref: { status: ['Active', 'Paused', 'Pending Start', 'Discontinued'] }, acc,
      cap: { chairs: 2, visitMin: 120, amMin: 240, pmMin: 240, tmpl: { 'Provider Alpha': { '1|AM': 2, '4|AM': 2, '2|PM': 1 }, 'Provider Beta': { '3|AM': 1, '5|PM': 1 } } },
      provCred: { 'provider alpha': 'MD', 'provider beta': 'NP' }, freqRate: {} },
    CFG: { clinic },
    // UI stubs only: the export checkbox, the lazy loader, toasts, the modal and the browser download
    $: s => (s === '#accXpts' && checkbox != null ? { checked: checkbox } : null),
    loadXLSX: async () => true, toast: (tone, msg) => { out.toast = [tone, msg]; }, closeModal: () => { out.closed = true; },
    downloadText: (name, data, type) => { Object.assign(out, { name, data, type }); return true; },
  });
  C.S.byMrn = new Map(C.S.P.map(p => [p.mrn, p]));
  vm.runInContext(NEED.map(chunk).join('\n') + '\nObject.assign(globalThis, { accessModel, accessBook, accXlsx, ACCESS_SCHEMA, MODALITY });', C);
  return { C, out, done: C.accXlsx().then(() => out) };
}
// Read back exactly like the Clinic Template Tool: XLSX.read(array, cellDates false, raw false), sheet_to_json(header 1, raw true)
function readBack(X, ctx, data){
  const u8 = vm.runInContext(`new Uint8Array(${data.byteLength})`, ctx); u8.set(new Uint8Array(data));
  const wb = X.read(u8, { type: 'array', cellDates: false, raw: false });
  const local = v => JSON.parse(JSON.stringify(v));   // vm arrays have another realm's prototype; values are only strings and numbers
  return { names: local(wb.SheetNames), aoa: n => local(X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' })) };
}

// ---- golden (map-external.md §1b) ----
const SHEETS = ['About', 'Providers', 'Template', 'DayPart', 'Frequencies', 'Cadence', 'Weekly', 'Credential'], PT_SHEETS = ['Moves', 'Patients'];
const HEAD = {
  About: ['Field', 'Value'],
  Providers: ['Modality', 'Provider', 'Credential', 'Patients', 'Need per week', 'MD Only patients', 'MD Only need', 'MD & NP patients', 'MD & NP need', 'Coverage not set patients', 'Coverage not set need', 'Template slots per week', 'Open per week', 'Booked per week', 'MD Only on NP panel'],
  Template: ['Modality', 'Provider', 'Credential', 'Weekday', 'Day part', 'Template slots'],
  DayPart: ['Modality', 'Weekday', 'Day part', 'Need per week', 'Patients usual slot', 'Booked per week', 'Template slots', 'Chair capacity'],
  Frequencies: ['Modality', 'Frequency group', 'Visits per week each', 'Interval weeks', 'Patients', 'Need per week', 'Booked per week', 'Unbooked per week'],
  Cadence: ['Modality', 'Frequency group', 'Interval weeks', 'Patients', 'Cycle week', 'Next week of', 'Patients now', 'Patients evened out', 'Even share'],
  Weekly: ['Modality', 'Week of', 'Series', 'Visits'],
  Credential: ['Modality', 'Metric', 'Visits per week'],
  Moves: ['Modality', 'Short name', 'MRN', 'Provider', 'Frequency group', 'Next visit', 'Shift weeks', 'Move to week of'],
  Patients: ['Modality', 'Short name', 'MRN', 'Status', 'Provider', 'Provider credential', 'Coverage requirement', 'Tx Frequency', 'Frequency group', 'Visits per week', 'Interval weeks', 'Cycle week of', 'Usual weekday', 'Usual day part', 'Next visit', 'Booked per week'],
};
const ABOUT_KEYS = ['Schema', 'Schema version', 'Modality', 'Clinic', 'Generated', 'Source', 'Week runs', 'First week of forecast', 'Forecast weeks', 'Booked window (weeks)', 'Statuses included', 'Pharmacy filter',
  'Chairs', 'Visit length (min)', 'AM block (min)', 'PM block (min)', 'Template slots per week', 'Chair capacity per week', 'Patients', 'Visits needed per week', 'Booked per week', 'Contains patient identifiers'];
const CTT_PROVIDER_COLS = ['Provider', 'Credential', 'Patients', 'Need per week', 'Template slots per week', 'Booked per week'];
const CREDENTIAL_METRICS = 16, BLOCK_COUNT = 10;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };
const sunday = iso => { assert.match(String(iso), /^\d{4}-\d{2}-\d{2}$/, 'ISO date: ' + iso); return new Date(iso + 'T00:00:00Z').getUTCDay() === 0; };

(async () => {
  const base = run(); const R = await base.done, B = readBack(base.C.XLSX, base.C, R.data);
  const about = Object.fromEntries(B.aoa('About').slice(1).map(r => [r[0], r[1]]));

  await t('file name, MIME and the export finishes', () => {
    assert.strictEqual(R.name, 'ESK access model - TST - 2026-10-15.xlsx');
    assert.strictEqual(R.type, XLSX_MIME);
    assert.strictEqual(R.toast[0], 'ok'); assert.strictEqual(R.closed, true);
  });
  await t('schema constant and modality', () => {
    assert.strictEqual(base.C.ACCESS_SCHEMA.name, 'mhs-access-model'); assert.strictEqual(base.C.ACCESS_SCHEMA.version, 1); assert.strictEqual(base.C.MODALITY, 'ESK');
  });
  await t('sheet names and order (patient rows off)', () => assert.deepStrictEqual(B.names, SHEETS));
  await t('every header row', () => { for (const s of SHEETS) assert.deepStrictEqual(B.aoa(s)[0], HEAD[s], s); });
  await t('About keys in order and fixed values', () => {
    assert.deepStrictEqual(B.aoa('About').slice(1).map(r => r[0]), ABOUT_KEYS);
    assert.strictEqual(about.Schema, 'mhs-access-model');
    assert.strictEqual(about['Schema version'], 1);   // a number, not "1"
    assert.strictEqual(about.Modality, 'ESK');
    assert.strictEqual(about.Clinic, 'TST');
    assert.strictEqual(about.Generated, new Date(NOW).toISOString());
    assert.strictEqual(about.Source, 'ESK Tracker add-in · Access tab');
    assert.strictEqual(about['Week runs'], 'Sunday to Saturday');
    assert.strictEqual(about['Pharmacy filter'], 'All');
    assert.strictEqual(about['Contains patient identifiers'], 'No');
    assert.strictEqual(about['Forecast weeks'], 12); assert.strictEqual(about['Booked window (weeks)'], 4);
    assert.strictEqual(about.Chairs, 2); assert.strictEqual(about['Visit length (min)'], 120);
    assert.strictEqual(about['Template slots per week'], 7);
    assert.strictEqual(about['Chair capacity per week'], 40);   // 2 chairs x 2 visits per 240-min block x 10 blocks (Mon-Fri AM/PM)
    assert.strictEqual(about.Patients, 8);                      // every status but Discontinued
    assert.strictEqual(about['Statuses included'], 'Active, Paused, Pending Start');
    for (const k of ['Visits needed per week', 'Booked per week']) assert.strictEqual(typeof about[k], 'number', k);
  });
  await t('Sunday week starts', () => {
    assert.strictEqual(about['First week of forecast'], '2026-10-11'); assert(sunday(about['First week of forecast']));
    const wk = B.aoa('Weekly').slice(1); assert(wk.length > 0);
    for (const r of wk) assert(sunday(r[1]), 'Weekly week of ' + r[1]);
    const cad = B.aoa('Cadence').slice(1); assert(cad.length > 0, 'cadence rows');
    for (const r of cad) assert(sunday(r[5]), 'Cadence next week of ' + r[5]);
  });
  await t('Modality is the first column of every data row', () => {
    for (const s of SHEETS.slice(1)){ const rows = B.aoa(s).slice(1); assert(rows.length > 0, s + ' has rows'); for (const r of rows) assert.strictEqual(r[0], 'ESK', s); }
  });
  await t('table shapes', () => {
    const prov = B.aoa('Providers').slice(1);
    assert.deepStrictEqual(prov.map(r => r[1]), ['Provider Alpha', 'Provider Beta', 'Not set']);
    assert.deepStrictEqual(prov.map(r => r[2]), ['MD', 'NP', '']);
    for (const r of prov) for (const c of [3, 4, 11, 13]) assert.strictEqual(typeof r[c], 'number', `Providers ${r[1]} col ${HEAD.Providers[c]}`);
    assert.strictEqual(B.aoa('Template').length - 1, prov.length * BLOCK_COUNT);
    const dp = B.aoa('DayPart').slice(1); assert.strictEqual(dp.length, BLOCK_COUNT + 1); assert.strictEqual(dp[BLOCK_COUNT][1], 'Unplaced');
    assert.deepStrictEqual(dp.slice(0, 2).map(r => [r[1], r[2]]), [['Mon', 'AM'], ['Mon', 'PM']]);
    assert.strictEqual(B.aoa('Credential').length - 1, CREDENTIAL_METRICS);
    const series = new Set(B.aoa('Weekly').slice(1).map(r => r[2])); for (const s of ['Projected total', 'Booked', 'Template capacity', 'Chair capacity']) assert(series.has(s), s);
    assert.strictEqual(new Set(B.aoa('Weekly').slice(1).map(r => r[1])).size, 12);
  });
  await t('no patient identifiers in the default file', () => {
    const cells = SHEETS.flatMap(s => B.aoa(s).flat().map(String));
    for (const id of PHI) assert(!cells.some(c => c.includes(id)), 'found ' + id);
  });
  await t('patient rows only when ticked, with the PHI suffix', async () => {
    const r = run({ checkbox: true }), o = await r.done, b = readBack(r.C.XLSX, r.C, o.data);
    assert.strictEqual(o.name, 'ESK access model - TST - 2026-10-15 - PHI.xlsx');
    assert.deepStrictEqual(b.names, SHEETS.concat(PT_SHEETS));
    for (const s of PT_SHEETS) assert.deepStrictEqual(b.aoa(s)[0], HEAD[s], s);
    assert.strictEqual(Object.fromEntries(b.aoa('About').slice(1).map(x => [x[0], x[1]]))['Contains patient identifiers'], 'Yes');
    const pts = b.aoa('Patients').slice(1); assert.strictEqual(pts.length, 8); for (const x of pts) assert.strictEqual(x[0], 'ESK');
    assert(!pts.some(x => x[2] === '900008'), 'Discontinued patient left out');
    const off = await run({ checkbox: false }).done; assert.strictEqual(off.name, R.name);
  });
  await t('export follows the Access filters and clinic code', async () => {
    const r = run({ clinic: 'T/S:T', acc: { weeks: 4, ph: 'GENOA', horizon: 12 } }), o = await r.done, b = readBack(r.C.XLSX, r.C, o.data);
    const ab = Object.fromEntries(b.aoa('About').slice(1).map(x => [x[0], x[1]]));
    assert.strictEqual(o.name, 'ESK access model - TST - 2026-10-15.xlsx');   // characters Windows won't take are dropped
    assert.strictEqual(ab.Clinic, 'T/S:T');
    assert.strictEqual(ab['Pharmacy filter'], 'GENOA'); assert.strictEqual(ab.Patients, 3);
    const blank = await run({ clinic: '' }).done; assert.strictEqual(blank.name, 'ESK access model - clinic - 2026-10-15.xlsx');
  });

  // ---- optional: the Clinic Template Tool's own reader (only on a PC with the OneDrive copy) ----
  const CTT = process.env.CTT_DIR || 'C:/Users/byates/OneDrive - Mindful Health Solutions/Add-In Manifests/Clinic Template Tool/add-in';
  if (fs.existsSync(path.join(CTT, 'engine.js')) && fs.existsSync(path.join(CTT, 'assets', 'xlsx.full.min.js'))){
    await t('Clinic Template Tool reads the file (its SheetJS + readAccessModel)', () => {
      const X = {}; X.window = X; X.self = X; vm.createContext(X); vm.runInContext(fs.readFileSync(path.join(CTT, 'assets', 'xlsx.full.min.js'), 'utf8'), X);
      const W = readBack(X.XLSX, X, R.data), sheets = {};
      for (const s of ['About', 'Providers', 'DayPart', 'Frequencies', 'Weekly']) if (W.names.includes(s)) sheets[s] = W.aoa(s);
      const TE = require(path.join(CTT, 'engine.js')), m = TE.readAccessModel(sheets);
      assert(!m.error, m.error); assert.strictEqual(m.version, 1); assert.strictEqual(m.modality, 'ESK'); assert.strictEqual(m.clinic, 'TST');
      assert.strictEqual(String(m.generated).slice(0, 10), '2026-10-15');
      assert.strictEqual(m.providers.length, 3); assert.strictEqual(m.dayPart.length, BLOCK_COUNT + 1); assert(m.frequencies.length > 0 && m.weekly.length > 0);
      for (const p of m.providers) for (const c of CTT_PROVIDER_COLS) assert(c in p, 'CTT column ' + c);
      assert.strictEqual(typeof m.providers[0].Patients, 'number'); assert.strictEqual(m.patients, undefined);
      const bad = { ...sheets, About: sheets.About.map(r => r[0] === 'Schema' ? ['Schema', 'mhs-access-model-v2'] : r) };
      assert(TE.readAccessModel(bad).error, 'a renamed schema is rejected');
    });
  } else console.log('skip Clinic Template Tool reader (no copy at ' + CTT + ')');
  console.log(`${n} passed`);
})().catch(e => { console.error(e); process.exit(1); });
