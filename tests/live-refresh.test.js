// Live refresh between co-authors: document settings re-read from the file, change detection, and the visit merge on save.
//   node tests/live-refresh.test.js
// Two panes share one mocked "file" of document settings; each keeps its own in-memory copy, as Office does.
// This is a mock, not Excel: it can't show how fast Excel on the web delivers a co-author's save. FICTIONAL MRNs only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const line = re => { const m = src.match(re); if (!m) throw new Error('not found: ' + re); return m[0]; };
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const code = [line(/^const DOC_KEYS = .*$/m), line(/^const CH = .*$/m), between('function OfficeStore(){', '// ---------- store: demo'),
  between('function visitMerge(', '// ---------- visit log'), 'globalThis.OfficeStore = OfficeStore; globalThis.visitMerge = visitMerge;'].join('\n');

const FILE = {};   // what is saved in the workbook
function pane(){
  const mem = JSON.parse(JSON.stringify(FILE)); const handlers = [];
  const settings = { get: k => mem[k] ?? null, set: (k, v) => { mem[k] = JSON.parse(JSON.stringify(v)); },
    saveAsync: cb => { Object.assign(FILE, JSON.parse(JSON.stringify(mem))); setTimeout(cb, 0); },
    refreshAsync: cb => { for (const k of Object.keys(mem)) delete mem[k]; Object.assign(mem, JSON.parse(JSON.stringify(FILE))); setTimeout(cb, 0); },
    addHandlerAsync: (ev, fn) => handlers.push(fn) };
  let onSheets = null;
  const C = { JSON, Object, Promise, setTimeout,
    Office: { context: { document: { settings }, requirements: { isSetSupported: () => true } }, EventType: { SettingsChanged: 'settingsChanged' } },
    Excel: { run: fn => fn({ sync: async () => {}, workbook: { worksheets: { onChanged: { add: f => { onSheets = f; } } } } }) } };
  vm.createContext(C); vm.runInContext(code, C);
  const store = C.OfficeStore(); store.load = (orig => async () => { try { await orig.call(store); } catch {} })(store.load);   // only the settings part of load() runs here
  return { store, C, mem, fireSettings: () => handlers.forEach(f => f()), fireSheet: () => onSheets && onSheets() };
}
const K1 = '900101|46316|600', K2 = '900102|46316|660';

(async () => {
  FILE.eskVisit = { [K1]: { start: 600 } };
  const a = pane(), b = pane(); let hits = 0; b.store.watch(() => hits++);
  await a.store.load(); await b.store.load();
  await a.store.setDoc('eskVisit', { [K1]: { start: 600, vit: { init: { bp: '120/80', at: 602 } } } });
  assert.strictEqual(b.store.getDoc('eskVisit')[K1].vit, undefined, 'b still has the copy from when it opened');
  await b.store.pollDocs(); assert.strictEqual(hits, 1, 'poll sees a\'s save');
  await b.store.pollDocs(); assert.strictEqual(hits, 2, 'keeps firing until b reloads');
  await b.store.load(); assert.strictEqual(b.store.getDoc('eskVisit')[K1].vit.init.bp, '120/80', 'load re-reads the file');
  await b.store.pollDocs(); assert.strictEqual(hits, 2, 'quiet after the reload');
  console.log('ok a co-author\'s vitals reach the other pane: poll fires, load re-reads the file');

  await b.store.setDoc('eskTasks', { x: 1 }); await b.store.pollDocs(); assert.strictEqual(hits, 2, 'own save is not a change');
  b.fireSettings(); b.fireSheet(); assert.strictEqual(hits, 4, 'settings-changed event and any sheet edit both refresh');
  console.log('ok own saves don\'t trigger a refresh; settings and sheet events do');

  // merge: b roomed K2 from a stale copy while a recorded K1 final vitals
  const base = { [K1]: { start: 600, vit: { init: { bp: '120/80' } } } };
  const fresh = { [K1]: { start: 600, vit: { init: { bp: '120/80' }, fin: { bp: '118/76' } } } };
  const mine = JSON.parse(JSON.stringify(base)); mine[K2] = { start: 660 };
  const out = b.C.visitMerge(mine, base, fresh);
  assert.strictEqual(out[K1].vit.fin.bp, '118/76', 'other pane\'s vitals kept'); assert.strictEqual(out[K2].start, 660, 'my new visit kept');
  assert.deepStrictEqual(base[K1], fresh[K1], 'base follows the file, so the visit log only records my change');
  const m2 = { [K1]: { start: 605 } }, b2 = { [K1]: { start: 600 } };
  assert.strictEqual(b.C.visitMerge(m2, b2, { [K1]: { start: 600, bay: '3' } })[K1].start, 605, 'a visit I changed keeps my version');
  const m3 = { [K1]: { start: 600 } }, b3 = { [K1]: { start: 600 } };
  assert.strictEqual(b.C.visitMerge(m3, b3, {})[K1], undefined, 'a visit someone else cleared stays cleared');
  assert.strictEqual(b.C.visitMerge(m3, b3, null), m3, 'no file copy: unchanged');
  console.log('ok saving merges per visit: others\' visits kept, mine win where I changed them');
  console.log('3 passed');
})().catch(e => { console.error(e); process.exit(1); });
