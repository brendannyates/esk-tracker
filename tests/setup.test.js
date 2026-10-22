// ESK Tracker Setup (setup.html): report parsing, the patient-match check, and the full run against a small in-memory mock of
// the Office.js workbook API (sheets, cells, Table1 as columns, protection, linked workbooks).
//   node tests/setup.test.js
// This is a mock, not Excel: formulas aren't evaluated, so it checks what is written where, not what Excel calculates.
// FICTIONAL MRNs, names and providers only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const html = fs.readFileSync(path.join(__dirname, '..', 'setup.html'), 'utf8').replace(/\r\n/g, '\n');
const code = html.match(/<script>\n([\s\S]*?)<\/script>/)[1];

// ---- mock workbook ----
const colNo = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
function parse(addr){
  let sheet = null; const m = addr.match(/^'(.+)'!(.+)$/); if (m){ sheet = m[1]; addr = m[2]; }
  const [a, b = a] = addr.replace(/\$/g, '').split(':'), p = x => { const q = x.match(/^([A-Z]+)(\d*)$/); return [q[2] ? +q[2] - 1 : null, colNo(q[1])]; };
  const [r0, c0] = p(a), [r1, c1] = p(b); return { sheet, r0, c0, r1, c1 };
}
function Book({ prot = null, cols, rows = 3, links = 0, extra = [] }){
  const B = { sheets: [], tables: [], links: Array.from({ length: links }, (_, i) => ({ id: 'L' + i })), log: [] };
  const sheet = name => { const s = { name, position: B.sheets.length, cells: new Map(), merges: [], fmt: {}, prot: { protected: false } }; B.sheets.push(s); return s; };
  const mpt = sheet('Master Patient Tracker'); extra.forEach(sheet);
  if (prot) mpt.prot = { protected: true, pw: prot.pw, options: { allowSort: true, allowAutoFilter: true } };
  B.t1 = { name: 'Table1', sheet: mpt, cols: cols.map((n, i) => ({ name: n, data: Array.from({ length: rows }, (_, r) => n === 'MRN' ? String(900100 + r) : ''), f: null, nf: null })) };
  B.tables.push(B.t1);
  return B;
}
function api(B){
  const guard = s => { if (s.prot.protected) throw new Error('sheet is protected'); };
  const rng = (s, a) => {
    const set = k => v => { guard(s); for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++){ const x = v[r - a.r0]?.[c - a.c0]; const cell = s.cells.get(r + ',' + c) || {}; cell[k] = x; s.cells.set(r + ',' + c, cell); } B.log.push([k, s.name, a]); };
    const R = { load(){}, merge(){ guard(s); s.merges.push(a); }, format: { font: {}, autofitColumns(){ B.log.push(['autofit', s.name, a]); } },
      getSpillingToRangeOrNullObject(){ const t = B.tables.find(t => t.name === 'Table47'); return { load(){}, isNullObject: false, rowCount: t.a.r1 - t.a.r0 } },
      get valueTypes(){ const o = []; for (let r = a.r0; r <= a.r1; r++){ const row = []; for (let c = a.c0; c <= a.c1; c++) row.push(String(s.cells.get(r + ',' + c)?.v ?? '').startsWith('#') ? 'Error' : 'String'); o.push(row); } return o; } };
    for (const k of ['values', 'formulas', 'numberFormat']) Object.defineProperty(R, k, { set: set({ values: 'v', formulas: 'f', numberFormat: 'nf' }[k]) });
    return R;
  };
  const col = (t, c) => ({ load(){}, get name(){ return c.name; }, set name(v){ guard(t.sheet); c.name = v; }, get index(){ return t.cols.indexOf(c); },
    delete(){ guard(t.sheet); t.cols.splice(t.cols.indexOf(c), 1); },
    getDataBodyRange(){ return { load(){}, set formulas(v){ guard(t.sheet); assert.strictEqual(v.length, c.data.length); c.f = v[0][0]; }, set numberFormat(v){ c.nf = v[0][0]; },
      get values(){ return c.data.map(v => [v]); }, get valueTypes(){ return c.data.map(() => ["String"]); } }; } });
  const tbl = t => t.cols ? {
    name: t.name, load(){}, isNullObject: false, worksheet: { load(){}, name: t.sheet.name },
    getHeaderRowRange(){ return { load(){}, values: [t.cols.map(c => c.name)] }; },
    getDataBodyRange(){ return { load(){}, rowCount: t.cols[0].data.length }; },
    columns: { load(){}, get items(){ return t.cols.map(c => col(t, c)); },
      getItem(n){ return col(t, t.cols.find(c => c.name === n)); },
      getItemOrNullObject(n){ const c = t.cols.find(c => c.name === n); return c ? col(t, c) : { load(){}, isNullObject: true }; },
      add(i, v, name){ guard(t.sheet); const c = { name, data: t.cols[0].data.map(() => ''), f: null }; t.cols.splice(i, 0, c); return col(t, c); } }
  } : { load(){}, isNullObject: false, set name(v){ t.name = v; }, getDataBodyRange(){ return rng(t.sheet, { r0: t.a.r0 + 1, c0: t.a.c0, r1: t.a.r1, c1: t.a.c1 }); } };
  const ws = s => ({ load(){}, get name(){ return s.name; }, get position(){ return s.position; }, set position(p){ B.sheets.splice(B.sheets.indexOf(s), 1); B.sheets.splice(p, 0, s); B.sheets.forEach((x, i) => x.position = i); },
    getRange: a => rng(s, parse(a)),
    protection: { load(){}, get protected(){ return s.prot.protected; }, get options(){ return s.prot.options; },
      unprotect(pw){ if (s.prot.pw && pw !== s.prot.pw) throw new Error('bad password'); s.prot.protected = false; },
      protect(o, pw){ s.prot = { protected: true, pw, options: o }; B.log.push(['protect', o, pw]); } } });
  const wb = {
    worksheets: { load(){}, get items(){ return B.sheets.map(ws); }, getItem: n => ws(B.sheets.find(s => s.name === n)),
      add(n){ const s = { name: n, position: B.sheets.length, cells: new Map(), merges: [], prot: { protected: false } }; B.sheets.push(s); return ws(s); } },
    tables: { getItem: n => tbl(B.tables.find(t => t.name === n)), getItemOrNullObject: n => { const t = B.tables.find(t => t.name === n); return t ? tbl(t) : { load(){}, isNullObject: true }; },
      add(addr, h){ const a = parse(addr), t = { name: 'Table' + (B.tables.length + 100), sheet: B.sheets.find(s => s.name === a.sheet), a }; B.tables.push(t); return tbl(t); } },
    linkedWorkbooks: { load(){}, get items(){ return B.links; }, breakAllLinks(){ B.log.push(['breakLinks', B.links.length]); B.links = []; } }
  };
  return { run: async fn => fn({ workbook: wb, sync: async () => {} }) };
}

// ---- mock page ----
function page(B){
  const els = {}, el = id => els[id] || (els[id] = { id, innerHTML: '', textContent: '', hidden: false, value: '', checked: false, disabled: false,
    insertAdjacentHTML(_, h){ this.innerHTML += h; }, click(){} });
  let ready = null;
  const C = { console, document: { getElementById: el }, Excel: api(B),
    Office: { onReady: fn => { ready = fn; }, context: { document: {}, requirements: { isSetSupported: () => true } } } };
  C.window = C; vm.createContext(C); vm.runInContext(code + '\nObject.assign(globalThis, { S, parseReport, overlap, run, ready, F, HDR });', C);
  return { C, els: el, start: () => ready() };
}
const cell = (s, addr) => { const a = parse(addr); return s.cells.get(a.r0 + ',' + a.c0) || {}; };

let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };
(async () => {
  const P = page(Book({ cols: ['MRN'] }));
  const aoa = [['1 ESK Ordering - B'], [], ['', 'Md Rc', 'Pat Name', 'Appt Dt', 'Beg Tm', 'Event', 'Sts', 'Rendering'],
    ['', '900100', 'TEST, ONE', 46316, 0.375, 'ESK Maint', 'Expected', 'DOE, JANE MD'], ['', '900101', 'TEST, TWO', '10/22/2026', '9:30 AM', 'ESK Start', 'Kept', 'DOE, JANE MD'],
    ['', '', '', '', '', '', '', ''], ['Total Appointments: 2']];
  await t('report: header found below the title, short NextGen names, footer skipped, dates to serials', () => {
    const r = P.C.parseReport(aoa); assert.strictEqual(r.rows.length, 2); assert.strictEqual(r.rendering, true);
    assert.deepStrictEqual([...r.rows[1]], ['900101', 'TEST, TWO', 46317, '9:30 AM', 'ESK Start', 'Kept', 'DOE, JANE MD']);
    assert.strictEqual(P.C.parseReport([['x'], ['MRN', 'Patient']]), null);
    const o = P.C.overlap(r.rows.concat([['900999']]), ['900100', ' 900101 ']); assert.strictEqual(o.pts, 3); assert.strictEqual(o.pct, 67);
  });

  await t('full run: protected tracker, report, links, renames and removals', async () => {
    const B = Book({ prot: { pw: 'pw1' }, cols: ['MRN', 'Status', 'Patient Name', 'Order Status', 'Last Tx', 'Next Tx', 'PCN', 'Care Nav', 'Protocol mg'], links: 2, extra: ['Inventory'] });
    const P = page(B); await P.start();
    assert.strictEqual(P.els('form').hidden, false); assert.strictEqual(P.els('pwBox').hidden, false, 'asks for the password');
    P.C.S.rep = P.C.parseReport(aoa); P.els('pw').value = 'pw1'; P.C.ready(); assert.strictEqual(P.els('run').disabled, false);
    await P.C.run();
    const ap = B.sheets.find(s => s.name === 'Appt List');
    assert.deepStrictEqual(B.sheets.map(s => s.name), ['Master Patient Tracker', 'Appt List', 'Inventory'], 'tab right after the tracker');
    assert.deepStrictEqual(ap.merges.map(a => [a.r0, a.c0, a.r1, a.c1]), [[0, 0, 0, 8], [1, 0, 1, 8]], 'rows 1 and 2 merged across A:I');
    assert.deepStrictEqual(Array.from({ length: 9 }, (_, c) => ap.cells.get("3," + c).v), [...P.C.HDR]);
    assert.strictEqual(cell(ap, 'A6').v, '900101'); assert.strictEqual(cell(ap, 'C5').v, 46316); assert.strictEqual(cell(ap, 'G6').v, 'DOE, JANE MD');
    assert.strictEqual(cell(ap, 'H6').f, P.C.F.Dose); assert.strictEqual(cell(ap, 'I5').f, P.C.F.Key); assert.strictEqual(cell(ap, 'AO5').f, P.C.F.AO);
    const t47 = B.tables.find(x => x.name === 'Table47'); assert.deepStrictEqual([t47.a.r0, t47.a.c0, t47.a.r1, t47.a.c1], [3, 0, 5, 8], 'Table47 = A4:I6');
    assert(B.log.some(l => l[0] === 'autofit' && l[1] === 'Appt List'), 'columns autofit');
    assert.deepStrictEqual(B.t1.cols.map(c => c.name), ['MRN', 'Status', 'Patient Name', 'Order Status', 'Last Treatment', 'Next Treatment', 'Scheduled Thru', 'Protocol mg']);
    for (const nm of ['Last Treatment', 'Next Treatment', 'Scheduled Thru']) assert.strictEqual(B.t1.cols.find(c => c.name === nm).f, P.C.F[nm]);
    assert.strictEqual(B.links.length, 0, 'links broken');
    const pr = B.log.find(l => l[0] === 'protect'); assert(pr && pr[2] === 'pw1' && pr[1].allowSort, 'protected again with the same options and password');
    assert(/Done\./.test(P.els('out').innerHTML), P.els('out').innerHTML + P.els('log').innerHTML); assert.strictEqual(P.els('pw').value, '', 'password cleared');
    assert(/already has an Appt List tab/.test(P.els('checks').innerHTML), 'second run is refused');
  });

  await t('template: no report, no password, no Protocol mg', async () => {
    const B = Book({ cols: ['MRN', 'Order Status', 'Notes'], rows: 1 }), P = page(B); await P.start();
    assert.strictEqual(P.els('pwBox').hidden, true); assert(/No Protocol mg column/.test(P.els('checks').innerHTML));
    P.els('none').checked = true; P.C.ready(); assert.strictEqual(P.els('run').disabled, false);
    await P.C.run();
    const ap = B.sheets.find(s => s.name === 'Appt List'), t47 = B.tables.find(x => x.name === 'Table47');
    assert.deepStrictEqual([t47.a.r0, t47.a.r1], [3, 4], 'one empty row'); assert.strictEqual(cell(ap, 'H5').f, P.C.F.DoseNoProtocol);
    assert.deepStrictEqual(B.t1.cols.map(c => c.name), ['MRN', 'Order Status', 'Last Treatment', 'Next Treatment', 'Scheduled Thru', 'Notes']);
    assert(/Done\./.test(P.els('out').innerHTML));
  });

  await t('wrong password: stops before any change', async () => {
    const B = Book({ prot: { pw: 'pw1' }, cols: ['MRN', 'Last Tx'] }), P = page(B); await P.start();
    P.els('none').checked = true; P.els('pw').value = 'nope'; await P.C.run();
    assert.deepStrictEqual(B.sheets.map(s => s.name), ['Master Patient Tracker']); assert.strictEqual(B.t1.cols[1].name, 'Last Tx');
    assert(/didn’t unprotect/.test(P.els('log').innerHTML) && /Not finished/.test(P.els('out').innerHTML));
  });

  await t('low match: needs the "right clinic" box', async () => {
    const B = Book({ cols: ['MRN'], rows: 1 }), P = page(B); await P.start();
    P.C.S.rep = { rows: [['800001'], ['800002']], rendering: true }; P.C.ready();
    assert(/0% of them are on this tracker/.test(P.els('match').innerHTML)); assert.strictEqual(P.els('run').disabled, true);
    P.C.S.lowOk = true; P.C.ready(); assert.strictEqual(P.els('run').disabled, false);
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
