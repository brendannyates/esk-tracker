// Eligibility feed: CSV parse (BOM, quotes, JSON-wrapped answers), MRN match, the pull that counts, tblBenefits plan, Sohar links.
//   node tests/elig-feed.test.js
// Runs the "Eligibility feed" block of src/taskpane.html in Node (with the helpers it uses). FICTIONAL codes, MRNs and ids only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const norm = .*$/m), line(/^const mrnKey = .*$/m), between('// ---------- Sohar benefits: parsing', '// ---------- Sohar benefits: end'),
  between('// ---------- Eligibility feed: pure', '// ---------- Eligibility feed: end'), between('const ELIG_FEED_COLS', 'function DemoStore(){'),
  'Object.assign(globalThis, { BEN_COLS, ELIG_BEN_ADD, ELIG_FEED_COLS, eligKey, feedText, feedTable, parseFeed, feedObj, feedForTracker, eligPick, eligNewest, eligBenVals, planEligBen, soharUrl, soharLinkOf, eligRelated, demoEligCsv });'].join('\n');
const C = {}; vm.createContext(C); vm.runInContext(code, C);
const eq = (x, y) => assert.strictEqual(JSON.stringify(x), JSON.stringify(y));
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

const H = C.ELIG_FEED_COLS;
const q = v => /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
const row = o => H.map(h => String(o[h] ?? ''));
const csv = rows => '﻿' + [H, ...rows].map(r => r.map(q).join(',')).join('\r\n') + '\r\n';
const V = (mrn, at, st, id, o = {}) => row({ 'Pulled At': at, LASFIR: o.code || 'TESABC', MRN: mrn, Kind: o.kind || 'Verification', 'Sohar Status': st, 'Sohar ID': id, Copay: o.copay ?? '30', Plan: o.plan || 'PPO', 'Related Sohar IDs': o.rel || '', 'Sohar Link': o.link || '' });
const objs = rows => rows.map(r => C.feedObj(H, r));

t('CSV: BOM, quoted commas and quotes, CRLF, header first', () => {
  const p = C.parseFeed(csv([V('0900101', '2026-11-08T15:05:00Z', 'complete.eligible', 'a1', { plan: 'PPO, "Choice"' })]));
  eq(p.header, H); assert.strictEqual(p.rows.length, 1);
  const o = C.feedObj(p.header, p.rows[0]); assert.strictEqual(o.plan, 'PPO, "Choice"'); assert.strictEqual(o.mrn, '0900101'); assert.strictEqual(o.pulledat, '2026-11-08T15:05:00Z');
});
t('JSON-wrapped answers: string, {body}, {csv}, {content}', () => {
  const text = csv([V('1', '2026-11-08', 'Pending', 'x')]);
  for (const w of [JSON.stringify(text), JSON.stringify({ body: text }), JSON.stringify({ csv: text }), JSON.stringify({ content: text })]) assert.strictEqual(C.parseFeed(w).rows.length, 1);
  assert.strictEqual(C.feedText(JSON.stringify({ other: 1 })), '');
});
t('not the feed: Sohar final tracker file, HTML error, empty, JSON error', () => {
  assert.strictEqual(C.parseFeed('MRN,Deductible Remaining,OOP Remaining\n1,2,3\n'), null);
  assert.strictEqual(C.parseFeed('<html>error</html>'), null); assert.strictEqual(C.parseFeed(''), null); assert.strictEqual(C.parseFeed(JSON.stringify({ error: 'x' })), null);
});
t('blank lines and trailing blank header cells are dropped', () => {
  const p = C.parseFeed([[...H, '', ''].join(','), '', V('2', '2026-11-08', 'Pending', 'y').join(','), ',,,'].join('\n'));
  assert.strictEqual(p.header.length, H.length); assert.strictEqual(p.rows.length, 1);
});
t('MRN key: digits only, leading zeros stripped', () => {
  assert.strictEqual(C.eligKey('00482913'), '482913'); assert.strictEqual(C.eligKey(482913), '482913'); assert.strictEqual(C.eligKey(' 0900101.0 '), '900101');
  assert.strictEqual(C.eligKey('MRN-0042'), '42'); assert.strictEqual(C.eligKey(''), ''); assert.strictEqual(C.eligKey('000'), '');
});
t('only rows whose MRN is on the tracker are kept', () => {
  const p = C.parseFeed(csv([V('0900101', '2026-11-08', 'complete.eligible', 'a'), V('900102', '2026-11-08', 'Pending', 'b'), V('777001', '2026-11-08', 'complete.eligible', 'c'), V('', '2026-11-08', 'Pending', 'd')]));
  const f = C.feedForTracker(p, [{ mrn: 900101 }, { mrn: '00900102' }, { mrn: '900103' }]);
  eq(f.rows.map(r => r[2]), ['0900101', '900102']); assert.strictEqual(f.dropped, 2); assert.strictEqual(f.people, 2);
});
t('pick: newest eligible verification beats a newer error or discovery', () => {
  const m = C.eligPick(objs([
    V('900101', '2026-11-09T15:00:00Z', 'error.member.id', 'new-error'),
    V('900101', '2026-11-08T15:00:00Z', 'complete.eligible', 'disc', { kind: 'Discovery' }),
    V('900101', '2026-11-01T15:00:00Z', 'complete.eligible', 'old-ok'),
    V('900101', '2026-11-05T15:00:00Z', 'complete.eligible (active)', 'mid-ok'),
  ]));
  assert.strictEqual(m.get('900101').soharid, 'mid-ok');
});
t('pick: no eligible verification, so the newest of any kind; ties keep the earlier line', () => {
  const m = C.eligPick(objs([
    V('900104', '2026-11-07T15:00:00Z', 'complete.ineligible', 'v-old'),
    V('900104', '2026-11-08T15:00:00Z', 'No coverage found', 'd-new', { kind: 'Discovery' }),
    V('900105', '2026-11-08T15:00:00Z', 'Pending', 'first'), V('900105', '2026-11-08T15:00:00Z', 'Pending', 'second'),
    V('00900106', '11/8/2026 9:00 AM', 'complete.ineligible', 'mdy'),
  ]));
  assert.strictEqual(m.get('900104').soharid, 'd-new'); assert.strictEqual(m.get('900105').soharid, 'first'); assert.strictEqual(m.get('900106').soharid, 'mdy');
});
t('newest first, stable', () => {
  const L = C.eligNewest(objs([V('1', '2026-11-01', 'x', 'a'), V('1', '2026-11-09', 'x', 'b'), V('1', '2026-11-05', 'x', 'c'), V('1', '2026-11-05', 'x', 'd')]));
  eq(L.map(o => o.soharid), ['b', 'c', 'd', 'a']);
});
t('benefits values: mapped columns, Sohar Date from Pulled At, Source, blank Per Nbr kept', () => {
  const o = C.feedObj(H, V('900101', '2026-11-08T15:00:00Z', 'complete.eligible', 'id1'));
  const v = C.eligBenVals(o, '2026-11-09 07:30');
  assert.strictEqual(v['Sohar Date'], '2026-11-08T15:00:00Z'); assert.strictEqual(v.Source, 'Eligibility feed'); assert.strictEqual(v.Imported, '2026-11-09 07:30');
  assert.strictEqual(v['Sohar ID'], 'id1'); assert.strictEqual(v.Copay, '30'); assert(!('Per Nbr' in v), 'blank Per Nbr not written'); assert(!('Sent As' in v), 'Sent As only when the feed has it');
  assert(!('LASFIR' in v) && !('Patient Name' in v));
});
t('tblBenefits plan: appends missing columns at the end, fills in place, adds a row, keeps BIDF and other rows', () => {
  const header = [...C.BEN_COLS, 'BIDF Per Treatment', 'Sohar ID'];   // one of the new columns already there
  const W = header.length, r = (mrn, name, copay, bidf) => { const a = Array(W).fill(''); a[0] = mrn; a[1] = name; a[9] = copay; a[C.BEN_COLS.length] = bidf; return a; };
  const body = [r(900101, 'ALDBEN', '10', '$140'), Array(W).fill(''), r('900200', 'KEEPME', '5', '$99')];
  const picks = C.eligPick(objs([V('0900101', '2026-11-08', 'complete.eligible', 'id1', { copay: '30' }), V('900102', '2026-11-08', 'complete.eligible', 'id2', { copay: '0' })]));
  const p = C.planEligBen(header, body, picks, [{ mrn: '900101', name: 'ALDBEN' }, { mrn: '900102', name: 'BROCAM' }, { mrn: '900102', name: 'BROCAM' }, { mrn: '900200', name: 'KEEPME' }], 'now');
  eq(p.header.slice(0, W), header); eq(p.addCols, C.ELIG_BEN_ADD.filter(c => c !== 'Sohar ID')); eq(p.header.slice(W), p.addCols);
  assert.strictEqual(p.updated, 1); assert.strictEqual(p.added, 1); assert.strictEqual(p.values.length, 3);
  const col = c => p.header.indexOf(c);
  assert.strictEqual(p.values[0][col('Copay')], '30'); assert.strictEqual(p.values[0][col('BIDF Per Treatment')], '$140'); assert.strictEqual(p.values[0][col('Source')], 'Eligibility feed');
  assert.strictEqual(p.values[1][0], '900102'); assert.strictEqual(p.values[1][1], 'BROCAM'); assert.strictEqual(p.values[1][col('Sohar ID')], 'id2');   // the blank row is reused
  eq(p.values[2].slice(0, W), body[2]);   // a row with no pull is untouched
  assert(!p.write.includes(col('BIDF Per Treatment')), 'BIDF column never written'); assert(p.write.includes(col('Sohar Link')));
});
t('tblBenefits plan: new table, no rows yet', () => {
  const picks = C.eligPick(objs([V('42', '2026-11-08', 'Pending', 'z')]));
  const p = C.planEligBen([...C.BEN_COLS, ...C.ELIG_BEN_ADD], [], picks, [{ mrn: '0042', name: 'TESABC' }], 'now');
  assert.strictEqual(p.added, 1); assert.strictEqual(p.values[0][0], '0042'); eq(p.addCols, []);
});
t('Sohar links: feed link kept only on app.soharhealth.com; ids cleaned; related by Kind', () => {
  assert.strictEqual(C.soharLinkOf({ soharlink: 'https://app.soharhealth.com/verifications/abc', soharid: 'x' }), 'https://app.soharhealth.com/verifications/abc');
  assert.strictEqual(C.soharLinkOf({ soharlink: 'https://evil.example/x', soharid: 'ab"c', kind: 'Discovery' }), 'https://app.soharhealth.com/discovery/abc');
  assert.strictEqual(C.soharLinkOf({ soharlink: '', soharid: '', kind: 'Verification' }), '');
  eq(C.eligRelated({ kind: 'Verification', relatedsoharids: 'd1; d2' }).map(x => x.url), ['https://app.soharhealth.com/discovery/d1', 'https://app.soharhealth.com/discovery/d2']);
  eq(C.eligRelated({ kind: 'Discovery', relatedsoharids: 'v1' }).map(x => x.url), ['https://app.soharhealth.com/verifications/v1']);
  eq(C.eligRelated({ kind: '', relatedsoharids: 'v9' }).map(x => x.url), ['https://app.soharhealth.com/verifications/v9']);
});
t('demo feed parses, drops the MRN that is not on the demo tracker, has only fake ids', () => {
  const p = C.parseFeed(C.demoEligCsv()); assert(p && p.rows.length === 10);
  const f = C.feedForTracker(p, [900101, 900102, 900104, 900107, 900109].map(mrn => ({ mrn })));
  assert.strictEqual(f.rows.length, 9); assert(objs(f.rows).every(o => /^00000000-0000-4000-8000-\d{12}$/.test(o.soharid)));
  assert(C.eligPick(objs(f.rows)).get('900102').soharid.endsWith('107'), 'the older eligible verification, not today\'s error');
});
console.log(`\n${n} tests passed`);
