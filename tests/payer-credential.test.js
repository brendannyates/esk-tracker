// Payer credential rules in the tracker: matching on the tracker's own patient / visit / kit shapes, the NP / MD check per
// ESK step, Besse chips, and "no sheet = nothing".
//   node tests/payer-credential.test.js
// Runs the "Payer credential rules" block of src/taskpane.html in Node with payer-rules.js (tests/payer-rules.test.js
// covers that module). MADE-UP payers, names and MRNs only: the real list lives in each workbook, never in this repo.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const root = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const a = src.indexOf('// ---------- Payer credential rules: who may bill'), b = src.indexOf('// ---------- Payer credential rules: end');
assert(a > 0 && b > a, 'payer block markers');
const code = [line(/^const esc = [^\n]*$/m), line(/^const norm = [^\n]*$/m), line(/^const MS = [^\n]*$/m), line(/^const sDate = [^\n]*$/m), line(/^const DOW = [^\n]*$/m),
  line(/^function fmtD[^\n]*\n[^\n]*$/m), line(/^function fmtT[^\n]*$/m), line(/^function tag[^\n]*$/m), src.slice(a, b),
  'Object.assign(globalThis, { PR, PAYER_SHEET, payerHint, credFor, payerModel, payerChips, payerVisitChip, payerTodoText, payerHTML });'].join('\n');
const PayerRules = require(path.join(root, 'payer-rules.js'));
const C = { PayerRules }; vm.createContext(C); vm.runInContext(code, C);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

// The sheet as imported (PayerRules.HEADERS). Made up.
const SHEET = [PayerRules.HEADERS,
  ['Acme Health', '', 'MD', 'Yes', 'N/A', 'Yes', 'Yes', 'Yes', 'acme; acme health', ''], ['Acme Health', '', 'NP', 'Yes', 'Yes', 'Yes', 'Yes', 'Yes', 'acme; acme health', ''],
  ['Zenith Blue CA', 'CA', 'MD', 'No', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith; zenith blue', ''], ['Zenith Blue CA', 'CA', 'NP', 'No', 'No', 'No', 'No', 'No', 'zenith; zenith blue', ''],
  ['Zenith Blue WA', 'WA', 'MD', 'No', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith; zenith blue', ''], ['Zenith Blue WA', 'WA', 'NP', 'No', 'Yes', 'Yes', 'Yes', 'Yes', 'zenith; zenith blue', ''],
  ['Orbit Federal', '', 'MD', '', 'N/A', 'Yes', 'Yes', 'Yes', 'orbit; orbit federal', ''], ['Orbit Federal', '', 'NP', '', 'No', 'No', 'No', 'Yes', 'orbit; orbit federal', 'NP may do observation only']];
const rules = PayerRules.read(SHEET).rules;
const T = 46300, CRED = { 'mdname, ava': 'MD', 'npname, ben': 'NP' };
// Same fields derive() builds
const pt = (mrn, o = {}) => ({ mrn, name: 'ZZZ' + mrn, status: o.status || 'Active', pharm: o.pharm || 'GENOA', insurance: o.ins ?? '', provider: o.prov ?? 'NPname, Ben' });
const ap = (mrn, date, o = {}) => ({ key: `${mrn}|${date}|600`, mrn, date, time: 600, event: o.ev || 'Esketamine Follow Up', status: o.st || 'Expected', active: o.active ?? true, rendering: o.rend ?? '' });
const kit = (mrn, o = {}) => ({ mrn, pharm: o.pharm || 'GENOA', administered: !!o.given, adminDate: o.given ?? null, disposed: false });
const run = (P, A = [], K = [], o = {}) => C.payerModel({ rules: o.rules === undefined ? rules : o.rules, P, A, K, ben: o.ben || null, provCred: CRED, hint: o.hint ?? 'CA', T });

t('script tag loads the shared file; sheet name matches', () => {
  assert(src.includes('<script src="payer-rules.js"></script>'), 'script tag');
  assert(fs.existsSync(path.join(root, 'payer-rules.js')), 'payer-rules.js at the site root');
  assert.strictEqual(C.PAYER_SHEET, PayerRules.SHEET);
});
t('state hint: Settings, else location or clinic code', () => {
  assert.strictEqual(C.payerHint({ payerState: 'wa', clinic: 'SRL' }, 'San Rafael'), 'WA');
  assert.strictEqual(C.payerHint({ payerState: '', clinic: 'SRL' }, 'San Rafael'), 'CA');
  assert.strictEqual(C.payerHint({ payerState: '', clinic: 'BEL' }, 'Bellevue - WA'), 'WA');
  assert.strictEqual(C.payerHint({ payerState: '', clinic: 'XYZ' }, 'Somewhere'), '');
});
t('credential: Templates & chairs first, else from the name', () => {
  assert.strictEqual(C.credFor('MDname, Ava', CRED), 'MD');
  assert.strictEqual(C.credFor('Lark PMHNP, Jo', CRED), 'NP');
  assert.strictEqual(C.credFor('Provider Q', CRED), '');
  assert.strictEqual(C.credFor('', CRED), '');
});
t('NP blocked, MD fine (tracker Insurance, state from the clinic)', () => {
  const m = run([pt('900201', { ins: 'Zenith Blue PPO' })], [ap('900201', T + 1), ap('900201', T + 8, { rend: 'MDname, Ava' })], [kit('900201', { given: T - 7 })]);
  const f = m.get('900201'); assert(f, 'matched'); assert.strictEqual(f.m.payer, 'Zenith Blue CA');
  assert.strictEqual(f.visits.length, 2);
  assert.strictEqual(f.visits[0].ok, false); assert.strictEqual(f.visits[0].step, 'obs');
  assert.strictEqual(f.visits[0].why, 'Zenith Blue CA doesn’t allow an NP to bill the esketamine observation: reassign to an MD or confirm with billing');
  assert.strictEqual(f.visits[1].ok, true, 'MD on the second visit');
  assert.strictEqual(f.bad.length, 1, 'tracker provider (NP) has the same problem: listed once');
  assert.strictEqual(C.payerVisitChip(f, f.visits[0].a.key), '<span class="chip bad" title="Zenith Blue CA doesn’t allow an NP to bill the esketamine observation: reassign to an MD or confirm with billing">NP can’t bill this observation</span>');
  assert.strictEqual(C.payerVisitChip(f, f.visits[1].a.key), '');
  assert.match(C.payerTodoText(f), /^Zenith Blue CA doesn’t allow an NP .* \(Tue \d+\/\d+ 10:00 AM with NPname, Ben\)$/);
  // the same patient in WA: NP is fine there
  assert.strictEqual(run([pt('900201', { ins: 'Zenith Blue PPO' })], [ap('900201', T + 1)], [], { hint: 'WA' }).get('900201').bad.length, 0);
});
t('observation-only NP: start and first treatment blocked, observation fine', () => {
  const P = [pt('900211', { ins: 'Orbit Federal', status: 'Pending Start' }), pt('900212', { ins: 'ORBIT FEDERAL SERVICES' }), pt('900213', { ins: 'Orbit Federal' })];
  const A = [ap('900211', T, { ev: 'Esketamine Tx Start' }), ap('900211', T + 3),   // start, then observation
    ap('900212', T + 2), ap('900212', T + 9),                                       // nothing given yet: first treatment, then observation
    ap('900213', T - 7, { st: 'Kept', active: false }), ap('900213', T + 1)];        // a kept visit before: observation
  const m = run(P, A);
  const s = m.get('900211').visits; assert.deepStrictEqual(s.map(v => [v.step, v.ok]), [['start', false], ['obs', true]]);
  assert.match(s[0].why, /^Orbit Federal doesn’t allow an NP to bill the esketamine start \(NP may do observation only\): reassign to an MD or confirm with billing$/);
  assert.deepStrictEqual(m.get('900212').visits.map(v => [v.step, v.ok]), [['first', false], ['obs', true]]);
  assert.deepStrictEqual(m.get('900213').visits.map(v => [v.step, v.ok]), [['obs', true]]);
  assert.strictEqual(m.get('900213').bad.length, 0);
  // an MD renders the start: that visit is fine, but the tracker provider (NP) is still flagged for it
  const m2 = run([pt('900214', { ins: 'Orbit Federal', status: 'Pending Start' })], [ap('900214', T, { ev: 'Esketamine Tx Start', rend: 'MDname, Ava' })]);
  const f2 = m2.get('900214'); assert.strictEqual(f2.visits[0].ok, true); assert.strictEqual(f2.bad.length, 1); assert.strictEqual(f2.bad[0].a, undefined);
  assert.match(C.payerTodoText(f2), /\(tracker provider NPname, Ben\)$/);
});
t('who is checked: current patients, booked visits from today; Sohar payer when Insurance is blank', () => {
  const P = [pt('900221', { ins: 'Zenith Blue', status: 'Discontinued' }), pt('900222', { ins: 'Nimbus Care' }), pt('900223', { ins: '' }), pt('900224', { ins: 'Zenith', prov: '' })];
  const A = [ap('900221', T + 1), ap('900222', T + 1), ap('900223', T + 1), ap('900224', T - 1), ap('900224', T + 1, { active: false }), ap('900224', T + 2, { rend: 'Provider Q' })];
  const ben = mrn => (mrn === '900223' ? { soharpayer: 'Zenith Blue Shield' } : null);
  const m = run(P, A, [], { ben });
  assert(!m.has('900221'), 'discontinued'); assert(!m.has('900222'), 'payer not on the list');
  assert.strictEqual(m.get('900223').payer, 'Zenith Blue Shield'); assert.strictEqual(m.get('900223').bad.length, 1);
  assert.strictEqual(m.get('900224').visits.length, 0, 'past, cancelled and unknown-credential visits are not checked'); assert.strictEqual(m.get('900224').bad.length, 0);
});
t('state rules that disagree: muted "check plan state", nothing flagged', () => {
  const f = run([pt('900231', { ins: 'Zenith Blue' })], [ap('900231', T + 1, { rend: 'NPname, Ben' })], [], { hint: '' }).get('900231');
  assert.deepStrictEqual(f.m.ambiguous, ['Zenith Blue CA', 'Zenith Blue WA']); assert.strictEqual(f.unsure, true); assert.strictEqual(f.bad.length, 0);
  assert.match(C.payerChips(f), /^<span class="chip" title="[^"]*Zenith Blue CA and Zenith Blue WA[^"]*">Check plan state<\/span>$/);
  // both states agree for an MD: no note
  const g = run([pt('900232', { ins: 'Zenith Blue', prov: 'MDname, Ava' })], [ap('900232', T + 1)], [], { hint: '' }).get('900232');
  assert.strictEqual(g.unsure, false); assert.strictEqual(C.payerChips(g), '');
});
t('Besse chips: Besse patients or a Besse kit assigned; flag only', () => {
  const P = [pt('900241', { ins: 'Zenith', pharm: 'BESSE', prov: 'MDname, Ava' }), pt('900242', { ins: 'Zenith', prov: 'MDname, Ava' }), pt('900243', { ins: 'Zenith', prov: 'MDname, Ava' }),
    pt('900244', { ins: 'Orbit Federal', pharm: 'BESSE', prov: 'MDname, Ava' }), pt('900245', { ins: 'Acme', pharm: 'BESSE' })];
  const K = [kit('900242', { pharm: 'BESSE' }), kit('900243', { pharm: 'BESSE', given: T - 3 })];
  const m = run(P, [], K);
  assert.strictEqual(m.get('900241').besse, 'not approved');
  assert.match(C.payerChips(m.get('900241')), /^<span class="chip warn" title="[^"]*">Payer isn’t Besse-approved<\/span>$/);
  assert.match(C.payerChips(m.get('900242')), /Payer isn’t Besse-approved/, 'Genoa patient with a Besse kit assigned');
  assert.strictEqual(C.payerChips(m.get('900243')), '', 'only an administered Besse kit: not assigned');
  assert.strictEqual(m.get('900244').besse, 'pending');
  assert.match(C.payerChips(m.get('900244')), /^<span class="chip" title="[^"]*">Besse approval pending<\/span>$/, 'blank = pending, muted');
  assert.strictEqual(C.payerChips(m.get('900245')), '', 'approved');
  assert.match(C.payerHTML(m.get('900244')), /^Orbit Federal <span class="chip"/);
});
t('no sheet, an empty sheet, or no payer-rules.js: nothing', () => {
  const P = [pt('900251', { ins: 'Zenith Blue', pharm: 'BESSE' })], A = [ap('900251', T + 1)];
  assert.strictEqual(run(P, A, [], { rules: null }).size, 0);
  assert.strictEqual(run(P, A, [], { rules: [] }).size, 0);
  assert.strictEqual(C.payerChips(null), ''); assert.strictEqual(C.payerVisitChip(null, 'x'), ''); assert.strictEqual(C.payerTodoText(null), ''); assert.strictEqual(C.payerHTML(null), '');
  const D = {}; vm.createContext(D); vm.runInContext(code, D);   // the script didn't load
  assert.strictEqual(D.PR(), null); assert.strictEqual(D.PAYER_SHEET, 'Payer Credential Rules');
  assert.strictEqual(D.payerModel({ rules, P, A, K: [], provCred: CRED, hint: 'CA', T }).size, 0);
});
t('the reload reads the sheet only through PayerRules.read, and a missing sheet leaves the rules null', () => {
  assert(/S\.payerRules = r\?\.rules \|\| null/.test(src));
  assert(/payerRows: yr && !yr\.isNullObject \? yr\.values : null/.test(src));
});
console.log(`${n} payer credential tests passed`);
