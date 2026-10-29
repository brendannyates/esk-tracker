// Shared state across the two ESK add-ins: ESK Tracker and ESK Bays have different add-in Ids, and Office keeps document settings
// per add-in, so each pane here gets its own settings store, as in Excel. What they share is the workbook: the hidden
// "ESK Shared" sheet. This is a mock, not Excel. FICTIONAL MRNs only.
//   node tests/shared-sheet.test.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = ['var S = {}; var CFG = { me: "Tess" };', line(/^const colLetter = .*$/m), line(/^const PRES_QUIET = .*$/m), line(/^const SH_SHEET = .*$/m), line(/^const DOC_KEYS = .*$/m), line(/^const CH = .*$/m),
  between('function OfficeStore(){', '// ---------- store: demo'), 'globalThis.OfficeStore = OfficeStore; globalThis.S = S; globalThis.SH_PART = SH_PART;'].join('\n');

// the workbook: sheets of cells, shared by every pane
const BOOK = new Map();
const colIdx = l => l.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
const parse = a => { const m = a.replace(/\$/g, '').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); return { c: colIdx(m[1]), r: +m[2] - 1, c2: colIdx(m[3] || m[1]), r2: +(m[4] || m[2]) - 1 }; };
let wsSeq = 0; const changed = [];
function sheetApi(name){
  const s = BOOK.get(name);
  const range = a => { const p = parse(a), rows = p.r2 - p.r + 1, cols = p.c2 - p.c + 1, grid = () => Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => s.cells.get(`${p.r + i},${p.c + j}`) ?? ''));
    return { load(){}, get values(){ return grid(); }, set values(v){ v.forEach((row, i) => row.forEach((x, j) => s.cells.set(`${p.r + i},${p.c + j}`, typeof x === 'number' ? x : String(x)))); changed.push(s.id); },
      set numberFormat(v){ assert.strictEqual(v.length, rows); v.forEach(r => assert.strictEqual(r.length, cols, 'numberFormat matches the range')); }, clear(){ for (let i = p.r; i <= p.r2; i++) for (let j = p.c; j <= p.c2; j++) s.cells.delete(`${i},${j}`); } };
  };
  return { isNullObject: false, get id(){ return s.id; }, load(){}, getRange: range, set visibility(v){ s.vis = v; } };
}
function excel(opts = {}){
  let onSheets = null;
  const ctx = { sync: async () => {}, workbook: { worksheets: {
    getItemOrNullObject: n => { if (opts.broken) throw new Error('protected'); return BOOK.has(n) ? sheetApi(n) : { isNullObject: true, load(){} }; },
    add: n => { BOOK.set(n, { id: 'ws' + (++wsSeq), cells: new Map(), vis: 'Visible' }); return sheetApi(n); },
    onChanged: { add: f => { onSheets = f; } } } } };
  return { Excel: { run: fn => fn(ctx), ClearApplyTo: { contents: 'Contents' }, SheetVisibility: { hidden: 'Hidden' } }, fire: id => onSheets && onSheets({ worksheetId: id, source: 'Remote' }) };
}
// one pane = one add-in: its own document settings (seeded), the shared workbook
function pane(seed = {}, opts){
  const file = JSON.parse(JSON.stringify(seed)), mem = JSON.parse(JSON.stringify(seed));
  const settings = { get: k => mem[k] ?? null, set: (k, v) => { mem[k] = JSON.parse(JSON.stringify(v)); }, saveAsync: cb => { Object.assign(file, JSON.parse(JSON.stringify(mem))); setTimeout(() => cb && cb({ status: 'succeeded' }), 0); },
    refreshAsync: cb => setTimeout(() => cb({ status: 'succeeded' }), 0), addHandlerAsync: () => {} };
  const X = excel(opts);
  const C = { JSON, Object, Promise, setTimeout, Math, Date, String, Array, Number, Map, Set, console: { error(){} },
    Office: { context: { document: { settings }, requirements: { isSetSupported: () => true } }, EventType: { SettingsChanged: 's' } }, Excel: X.Excel };
  vm.createContext(C); vm.runInContext(code, C);
  const store = C.OfficeStore(); store.load = (orig => async () => { try { await orig.call(store); } catch {} })(store.load);   // the tables part of load() isn't mocked
  let hits = 0; store.watch(() => hits++);
  return { store, mem, file, fire: X.fire, hits: () => hits, C };
}
const V1 = '900201|46323|600', V2 = '900202|46323|660';
const MAP = { name: 'Test rooms', bays: [{ id: '1' }, { id: '2' }], walls: [] };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok ' + name); };

(async () => {
  const tracker = pane({ eskShared: { bayPlan: MAP, vitRange: { sys: 140 } }, eskVisit: { [V1]: { start: 600, bay: '1', vit: { init: { bp: '120/80' } } } }, eskTasks: { a: 1 }, 'Office.AutoShowTaskpaneWithDocument': true });
  const bays = pane({ eskVisit: { [V1]: { start: 600, pc: { at: 610, sn: '1234' } }, [V2]: { start: 660, bay: '2' } } });

  await t('the tracker copies its settings to the hidden ESK Shared sheet on its first open', async () => {
    await tracker.store.load();
    assert(BOOK.has('ESK Shared') && BOOK.get('ESK Shared').vis === 'Hidden', 'sheet made, hidden');
    assert.deepStrictEqual(tracker.store.getDoc('eskShared').bayPlan, MAP);
    assert(tracker.mem.eskSharedSheet, 'marked as copied in its own settings');
  });

  await t('ESK Bays, with its own settings, now sees the bay map the tracker saved (the bug)', async () => {
    await bays.store.load();
    assert.deepStrictEqual(bays.store.getDoc('eskShared').bayPlan, MAP, 'bay map reaches the provider add-in');
    assert.strictEqual(bays.store.getDoc('eskShared').vitRange.sys, 140);
  });

  await t('the visits merge field by field: the provider\'s dose sign-off joins the tech\'s visit, nothing is lost', async () => {
    const v = bays.store.getDoc('eskVisit');
    assert.strictEqual(v[V1].vit.init.bp, '120/80', 'tech vitals kept');
    assert.strictEqual(v[V1].pc.sn, '1234', 'provider sign-off kept');
    assert.strictEqual(v[V2].bay, '2', 'visit only ESK Bays had');
    await tracker.store.pollDocs();
    assert.strictEqual(tracker.store.getDoc('eskVisit')[V1].pc.sn, '1234', 'the tracker sees the sign-off after a poll');
  });

  await t('a PRN order the provider saves reaches the tracker; a change to the shared sheet reads at once', async () => {
    const v = bays.store.getDoc('eskVisit'); v[V1].prnOrd = [{ id: 'p1', med: 'ondansetron', ts: 1 }];
    assert.strictEqual(await bays.store.setDoc('eskVisit', v), true, 'saved');
    const before = tracker.hits(); tracker.fire(BOOK.get('ESK Shared').id); await new Promise(r => setTimeout(r, 20));
    assert(tracker.hits() > before, 'the sheet edit fires the settings path, not a table refresh');
    assert.strictEqual(tracker.store.getDoc('eskVisit')[V1].prnOrd[0].med, 'ondansetron');
    assert.strictEqual(tracker.mem.eskVisit[V1].prnOrd, undefined, 'it came through the sheet, not the tracker\'s own settings');
  });

  await t('a new bay map saved in the tracker shows in ESK Bays after its next poll; quiet afterwards', async () => {
    const sh = tracker.store.getDoc('eskShared'); sh.bayPlan = { ...MAP, name: 'Moved rooms' }; await tracker.store.setDoc('eskShared', sh);
    const h = bays.hits(); await bays.store.pollDocs(); assert.strictEqual(bays.hits(), h + 1);
    assert.strictEqual(bays.store.getDoc('eskShared').bayPlan.name, 'Moved rooms');
    bays.store.markDocs(); await bays.store.pollDocs(); assert.strictEqual(bays.hits(), h + 1, 'no repeat once applied');
  });

  await t('a value past one cell splits into parts and reads back whole; getDoc hands out copies', async () => {
    const big = { notes: 'x'.repeat(C_PART() * 2 + 17) }; await tracker.store.setDoc('eskNotes', big);
    await bays.store.pollDocs(); assert.strictEqual(bays.store.getDoc('eskNotes').notes.length, big.notes.length);
    const a = bays.store.getDoc('eskNotes'); a.notes = 'changed'; assert.notStrictEqual(bays.store.getDoc('eskNotes').notes, 'changed');
  });

  await t('Office\'s own setting (open with the workbook) stays per add-in', async () => {
    assert.strictEqual(tracker.store.getDoc('Office.AutoShowTaskpaneWithDocument'), true);
    assert.strictEqual(bays.store.getDoc('Office.AutoShowTaskpaneWithDocument'), null);
  });

  await t('copying over happens once per add-in: a later open doesn\'t bring old settings back', async () => {
    const v = tracker.store.getDoc('eskTasks'); assert.strictEqual(v.a, 1);
    await bays.store.setDoc('eskTasks', { b: 2 }); await tracker.store.load();
    assert.deepStrictEqual(tracker.store.getDoc('eskTasks'), { b: 2 }, 'the tracker reads the sheet, its own old copy isn\'t merged in again');
  });

  await t('without the sheet (protected workbook structure): the add-in\'s own settings, as before', async () => {
    const solo = pane({ eskShared: { bayPlan: MAP } }, { broken: true }); await solo.store.load();
    assert.deepStrictEqual(solo.store.getDoc('eskShared').bayPlan, MAP);
    assert.strictEqual(await solo.store.setDoc('eskShared', { bayPlan: null }), true);
    assert.strictEqual(solo.mem.eskShared.bayPlan, null);
  });
  console.log(`${n} passed`);
})().catch(e => { console.error(e); process.exit(1); });
function C_PART(){ return 30000; }
