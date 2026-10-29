// Shared state across the two ESK add-ins: ESK Tracker and ESK Bays have different add-in Ids, and Office keeps document settings
// per add-in, so each pane here gets its own settings store, as in Excel. What they share is the workbook: the hidden
// "ESK Shared" sheet. The mock reads cells when a sync runs (not later) and every sync waits a turn, so two panes really interleave.
// This is a mock, not Excel. FICTIONAL MRNs only.
//   node tests/shared-sheet.test.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = ['var S = {}; var CFG = { me: "Tess" };', line(/^const colLetter = .*$/m), line(/^const PRES_QUIET = .*$/m), line(/^const SH_SHEET = .*$/m), line(/^const DOC_KEYS = .*$/m), line(/^const CH = .*$/m),
  between('function OfficeStore(){', '// ---------- store: demo'), 'globalThis.OfficeStore = OfficeStore; globalThis.S = S; globalThis.DOC_KEYS = DOC_KEYS;'].join('\n');

const DOC_KEYS = vm.runInNewContext(line(/^const DOC_KEYS = .*$/m).replace(/^const DOC_KEYS = /, "").replace(/;s*$/, ""));   // the same list the store uses
// the workbook: sheets of cells, shared by every pane
const BOOK = new Map();
const colIdx = l => l.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
const parse = a => { const m = a.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); return { c: colIdx(m[1]), r: +m[2] - 1, c2: colIdx(m[3] || m[1]), r2: +(m[4] || m[2]) - 1 }; };
let wsSeq = 0;
function sheetApi(name, opts){
  const s = BOOK.get(name);
  const range = a => { const p = parse(a), rows = p.r2 - p.r + 1, cols = p.c2 - p.c + 1, grid = () => Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => s.cells.get(`${p.r + i},${p.c + j}`) ?? ''));
    let snap = null;
    return { load(){ if (opts.readFails) throw new Error('An internal error has occurred.'); snap = grid(); }, get values(){ return snap || grid(); },
      set values(v){ v.forEach((row, i) => row.forEach((x, j) => s.cells.set(`${p.r + i},${p.c + j}`, typeof x === 'number' ? x : String(x)))); },
      set numberFormat(v){ assert.strictEqual(v.length, rows); v.forEach(r => assert.strictEqual(r.length, cols, 'numberFormat matches the range')); },
      clear(){ for (let i = p.r; i <= p.r2; i++) for (let j = p.c; j <= p.c2; j++) s.cells.delete(`${i},${j}`); } };
  };
  return { isNullObject: false, get id(){ return s.id; }, load(){}, getRange: range, set visibility(v){ s.vis = v; } };
}
function excel(opts = {}){
  let onSheets = null;
  const ctx = { sync: () => new Promise(r => setTimeout(r, 0)), workbook: { tables: { load(){}, items: [] }, worksheets: { load(){}, items: [],
    getItemOrNullObject: n => { if (opts.broken) throw new Error('protected'); return BOOK.has(n) ? sheetApi(n, opts) : { isNullObject: true, load(){} }; },
    add: n => { if (BOOK.has(n)) throw new Error('ItemAlreadyExists'); BOOK.set(n, { id: 'ws' + (++wsSeq), cells: new Map(), vis: 'Visible' }); return sheetApi(n, opts); },
    onChanged: { add: f => { onSheets = f; } } } } };
  return { Excel: { run: fn => fn(ctx), ClearApplyTo: { contents: 'Contents' }, SheetVisibility: { hidden: 'Hidden' } }, fire: id => onSheets && onSheets({ worksheetId: id, source: 'Remote' }), opts };
}
// one pane = one add-in: its own document settings (seeded), the shared workbook. bays: the ESK Bays view.
function pane(seed = {}, opts = {}){
  const file = JSON.parse(JSON.stringify(seed)), mem = JSON.parse(JSON.stringify(seed));
  const settings = { get: k => mem[k] ?? null, set: (k, v) => { mem[k] = JSON.parse(JSON.stringify(v)); }, saveAsync: cb => { Object.assign(file, JSON.parse(JSON.stringify(mem))); setTimeout(() => cb && cb({ status: 'succeeded' }), 0); },
    refreshAsync: cb => setTimeout(() => cb({ status: 'succeeded' }), 0), addHandlerAsync: () => {} };
  const X = excel(opts);
  const C = { JSON, Object, Promise, setTimeout, Math, Date, String, Array, Number, Map, Set, console: { error(){} },
    Office: { context: { document: { settings }, requirements: { isSetSupported: () => true } }, EventType: { SettingsChanged: 's' } }, Excel: X.Excel };
  vm.createContext(C); vm.runInContext(code, C); C.S.provView = !!opts.bays;
  const store = C.OfficeStore(); let loadErr = null;
  store.load = (orig => async () => { loadErr = null; try { await orig.call(store); } catch (e){ loadErr = e; } })(store.load);   // no tables in this workbook: load() reads none
  let hits = 0; store.watch(() => hits++);
  return { store, mem, file, fire: X.fire, hits: () => hits, opts: X.opts, err: () => loadErr };
}
const rowKeys = () => { const s = BOOK.get('ESK Shared'), out = {}; for (const [k, v] of s.cells) { const [r, c] = k.split(',').map(Number); if (c === 0 && r > 0) out[r + 1] = v; } return out; };
const fixedRows = () => { const keys = rowKeys(); return Object.entries(keys).every(([r, k]) => DOC_KEYS.indexOf(k) + 2 === +r) && new Set(Object.values(keys)).size === Object.values(keys).length; };
const V1 = '900201|46323|600', V2 = '900202|46323|660';
const MAP = { name: 'Test rooms', bays: [{ id: '1' }, { id: '2' }], walls: [] };
let n = 0; const t = async (name, fn) => { BOOK.clear(); await fn(); n++; console.log('ok ' + name); };
const tracker0 = () => ({ eskShared: { bayPlan: MAP, vitRange: { sys: 140 } }, eskVisit: { [V1]: { start: 600, bay: '1', vit: { init: { bp: '120/80' } } } }, eskTasks: { a: 1 }, 'Office.AutoShowTaskpaneWithDocument': true });
const bays0 = () => ({ eskApptList: { sig: 'x' }, eskVisit: { [V1]: { start: 600, pc: { at: 610, sn: '1234' } }, [V2]: { start: 660, bay: '2' } } });

(async () => {
  await t('the bay map the tracker saved reaches ESK Bays, which has its own settings (the original bug)', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true });
    await tr.store.load(); await b.store.load();
    assert(BOOK.get('ESK Shared').vis === 'Hidden', 'sheet hidden');
    assert.deepStrictEqual(b.store.getDoc('eskShared').bayPlan, MAP);
    const v = b.store.getDoc('eskVisit');
    assert.strictEqual(v[V1].vit.init.bp, '120/80', 'tech vitals'); assert.strictEqual(v[V1].pc.sn, '1234', 'provider sign-off'); assert.strictEqual(v[V2].bay, '2', 'visit only ESK Bays had');
    assert(fixedRows(), 'each key on its own row: ' + JSON.stringify(rowKeys()));
  });

  await t('both add-ins copying over at the same moment: no row taken twice, nothing lost (review finding 1)', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true });
    await Promise.all([tr.store.load(), b.store.load()]);
    assert(!tr.err() && !b.err(), 'both loaded: ' + (tr.err() || b.err()));
    assert(fixedRows(), 'fixed rows, no duplicates: ' + JSON.stringify(rowKeys()));
    await tr.store.pollDocs(); await b.store.pollDocs();
    for (const p of [tr, b]){
      assert.deepStrictEqual(p.store.getDoc('eskTasks'), { a: 1 }); assert.deepStrictEqual(p.store.getDoc('eskApptList'), { sig: 'x' });
      const v = p.store.getDoc('eskVisit'); assert.strictEqual(v[V1].pc?.sn, '1234', 'sign-off survives'); assert.strictEqual(v[V1].bay, '1', 'tech bay survives'); assert(v[V2], 'second visit');
    }
  });

  await t('the first version\'s rows and merge are repaired: newest value per key, the tracker\'s tech fields, the provider\'s answers (findings 1, 3)', async () => {
    // what the first version left: ESK Bays copied over first, so its log-restored bay and vitals beat the tracker's on V1; eskTasks on two rows
    BOOK.set('ESK Shared', { id: 'ws0', vis: 'Hidden', cells: new Map() });
    const put = (row, k, stamp, v) => { const c = BOOK.get('ESK Shared').cells; c.set(`${row - 1},0`, k); c.set(`${row - 1},1`, stamp); c.set(`${row - 1},2`, '1'); c.set(`${row - 1},3`, JSON.stringify(v)); };
    const order = { id: 'o1', med: 'Zofran', ts: 1, ack: { by: 'PCT' } }, orderAnswered = { id: 'o1', med: 'Zofran', ts: 1, ok: { by: 'Dr A' } };
    put(2, 'eskVisit', '2026-10-29T06:28:13Z|B|a', { [V1]: { start: 600, bay: '9', vit: { init: { bp: '100/60' } }, prnOrd: [orderAnswered], pc: { sn: '1234' } } });
    put(3, 'eskTasks', '2026-10-29T06:28:14Z|B|b', { old: 1 }); put(4, 'eskTasks', '2026-10-29T06:28:58Z|T|c', { a: 1 });
    const tr = pane({ ...tracker0(), eskVisit: { [V1]: { start: 600, bay: '1', vit: { init: { bp: '120/80' } }, prnOrd: [order] } }, eskSharedSheet: 'x' });
    await tr.store.load(); assert(!tr.err(), String(tr.err()));
    assert(fixedRows(), 'rows put in place: ' + JSON.stringify(rowKeys()));
    assert.deepStrictEqual(tr.store.getDoc('eskTasks'), { a: 1 }, 'newest of the two eskTasks rows');
    const v = tr.store.getDoc('eskVisit')[V1];
    assert.strictEqual(v.bay, '1'); assert.strictEqual(v.vit.init.bp, '120/80'); assert.strictEqual(v.pc.sn, '1234');
    assert(v.prnOrd.length === 1 && v.prnOrd[0].ok.by === 'Dr A' && v.prnOrd[0].ack.by === 'PCT', 'the order keeps both the approval and the PCT\'s ack');
    assert(tr.mem.eskSharedSheet2, 'repair runs once');
  });

  await t('a visits save over a row someone saved since is refused, then saveVisit\'s retry merges and saves (finding 4)', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true });
    await tr.store.load(); await b.store.load(); await tr.store.pollDocs();
    const mine = tr.store.getDoc('eskVisit'); mine[V1].out = 700;                       // the tech ends the visit...
    const theirs = b.store.getDoc('eskVisit'); theirs[V1].prnOrd = [{ id: 'p1', ts: 2 }]; // ...while the provider orders a PRN, first
    assert.strictEqual(await b.store.setDoc('eskVisit', theirs), true);
    assert.strictEqual(await tr.store.setDoc('eskVisit', mine), false, 'refused: it would write over the PRN order');
    const fresh = await tr.store.freshDoc('eskVisit'); assert.strictEqual(fresh[V1].prnOrd[0].id, 'p1', 'the refusal brought the newer row in');
    fresh[V1].out = 700; assert.strictEqual(await tr.store.setDoc('eskVisit', fresh), true, 'merged save goes through');
    await b.store.pollDocs(); const v = b.store.getDoc('eskVisit')[V1]; assert(v.out === 700 && v.prnOrd[0].id === 'p1', 'both changes on the sheet');
  });

  await t('a change on the shared sheet reads at once; a new bay map shows after the next poll and then stays quiet', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true }); await tr.store.load(); await b.store.load();
    const sh = tr.store.getDoc('eskShared'); sh.bayPlan = { ...MAP, name: 'Moved rooms' }; await tr.store.setDoc('eskShared', sh);
    const h = b.hits(); b.fire(BOOK.get("ESK Shared").id); await new Promise(r => setTimeout(r, 200));
    assert(b.hits() > h, 'the sheet edit fires the settings path'); assert.strictEqual(b.store.getDoc('eskShared').bayPlan.name, 'Moved rooms');
    b.store.markDocs(); const h2 = b.hits(); await b.store.pollDocs(); assert.strictEqual(b.hits(), h2, 'no repeat once applied');
  });

  await t('a value past one cell splits into parts and reads back whole; getDoc hands out copies', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true }); await tr.store.load(); await b.store.load();
    const big = { notes: 'x'.repeat(30000 * 2 + 17) }; await tr.store.setDoc('eskNotes', big);
    await b.store.pollDocs(); assert.strictEqual(b.store.getDoc('eskNotes').notes.length, big.notes.length);
    const a = b.store.getDoc('eskNotes'); a.notes = 'changed'; assert.notStrictEqual(b.store.getDoc('eskNotes').notes, 'changed');
  });

  await t('Office\'s own open-with-the-workbook setting stays per add-in', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true }); await tr.store.load(); await b.store.load();
    assert.strictEqual(tr.store.getDoc('Office.AutoShowTaskpaneWithDocument'), true);
    assert.strictEqual(b.store.getDoc('Office.AutoShowTaskpaneWithDocument'), null);
  });

  await t('the merge runs once per add-in: a later open doesn\'t bring old settings back', async () => {
    const tr = pane(tracker0()), b = pane(bays0(), { bays: true }); await tr.store.load(); await b.store.load();
    await b.store.setDoc('eskTasks', { b: 2 }); await tr.store.load();
    assert.deepStrictEqual(tr.store.getDoc('eskTasks'), { b: 2 });
  });

  await t('once the sheet exists, a failing read is a load error, never a quiet switch to own settings (finding 2)', async () => {
    const tr = pane(tracker0()); await tr.store.load();
    const b = pane(bays0(), { bays: true }); b.opts.readFails = true; await b.store.load();
    assert(b.err(), 'load reports the problem');
    assert.notDeepStrictEqual(b.store.getDoc('eskVisit'), bays0().eskVisit, 'not its own stale copy');
    assert.strictEqual(await b.store.setDoc('eskVisit', { x: 1 }), false, 'a save says it didn\'t reach the sheet');
    b.opts.readFails = false; await b.store.load(); assert(!b.err()); assert.deepStrictEqual(b.store.getDoc('eskShared').bayPlan, MAP, 'back on the sheet');
  });

  await t('no sheet and none can be made (protected workbook structure): the add-in\'s own settings, as before', async () => {
    const solo = pane({ eskShared: { bayPlan: MAP } }, { broken: true }); await solo.store.load();
    assert.deepStrictEqual(solo.store.getDoc('eskShared').bayPlan, MAP);
    assert.strictEqual(await solo.store.setDoc('eskShared', { bayPlan: null }), true); assert.strictEqual(solo.mem.eskShared.bayPlan, null);
  });
  console.log(`${n} passed`);
})().catch(e => { console.error(e); process.exit(1); });
