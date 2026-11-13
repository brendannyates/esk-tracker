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

// The Eligibility feed flow's { csv, rows } answer: Sohar webhook rows merged into the rotation CSV by Sohar ID. FICTIONAL data.
const H = C.ELIG_FEED_COLS, q = v => /[",\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
const row = o => H.map(h => q(String(o[h] ?? ''))).join(',');
const csv = '﻿' + [H.join(','), row({ 'Pulled At': '2026-11-09T10:00:00Z', MRN: '900', Kind: 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar ID': 'v-1', 'Projected Patient Responsibility': '$20.00' })].join('\r\n');
const rows = [
  { 'Pulled At': '2026-11-09T10:00:05Z', MRN: '900', Kind: 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar ID': 'v-1', Source: 'Sohar webhook' },
  { 'Pulled At': '2026-11-09T16:00:00Z', MRN: '901', Kind: 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar ID': 'v-2', Copay: '15', '@odata.etag': 'x' },
  { 'Pulled At': '2026-11-09T17:00:00Z', MRN: '901', Kind: 'Verification', 'Sohar Status': 'complete.eligible', 'Sohar ID': 'v-2', Copay: '25' },
  { 'Pulled At': '', MRN: '', 'Sohar ID': '' }];
const t = C.parseFeed(JSON.stringify({ csv, rows }));
assert(t && t.rows.length === 2, 'two pulls: v-1 from the CSV, v-2 once');
const objs = t.rows.map(r => C.feedObj(t.header, r));
assert.strictEqual(objs.find(o => o.soharid === 'v-1').projectedpatientresponsibility, '$20.00');
assert.strictEqual(objs.find(o => o.soharid === 'v-2').copay, '25');
assert.strictEqual(C.parseFeed(JSON.stringify({ csv: '', rows })).rows.length, 2);
console.log('ok webhook rows merge into the feed');
