// Settings sheet for Teams notifications (notify-settings.js) against a small mock of Excel's sheet and table API.
//   node tests/notify-settings.test.js
// Checks: tables land where the spec says, nothing existing is overwritten, provider rows add / update / remove. FICTIONAL names only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const code = fs.readFileSync(path.join(__dirname, '..', 'notify-settings.js'), 'utf8');

// ---- mock workbook: sheets of cells; tables = { name, sheet, r0, c0, rows, cols } (rows include the header)
function workbook(){
  const sheets = {}, tables = [];
  const A1 = a => { const m = a.match(/^([A-Z]+)(\d+)$/); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2] - 1, c - 1]; };
  const cells = name => sheets[name] || (sheets[name] = {});
  const get = (sh, r, c) => cells(sh)[r + ',' + c] ?? '', set = (sh, r, c, v) => { cells(sh)[r + ',' + c] = v; };
  const block = (sh, r0, c0, nr, nc) => {
    const o = { _load(){ return o; }, load(){ return o; },
      get values(){ return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => get(sh, r0 + i, c0 + j))); },
      set values(v){ v.forEach((row, i) => row.forEach((x, j) => set(sh, r0 + i, c0 + j, x))); },
      rowCount: nr,
      getCell: (i, j) => block(sh, r0 + i, c0 + j, 1, 1),
      getResizedRange: (dr, dc) => block(sh, r0, c0, nr + dr, nc + dc) };
    return o; };
  const range = (sh, addr) => { const [a, b] = addr.split(':'), [r0, c0] = A1(a), [r1, c1] = A1(b || a); return block(sh, r0, c0, r1 - r0 + 1, c1 - c0 + 1); };
  const tableApi = t => ({
    getDataBodyRange: () => block(t.sheet, t.r0 + 1, t.c0, t.rows - 1, t.cols),
    getRange: () => block(t.sheet, t.r0, t.c0, t.rows, t.cols),
    rows: { add: (_, vals) => { for (const v of vals){ v.forEach((x, j) => set(t.sheet, t.r0 + t.rows, t.c0 + j, x)); t.rows++; } },
      getItemAt: i => ({ delete(){ const r = t.r0 + 1 + i; for (let k = r; k < t.r0 + t.rows - 1; k++) for (let j = 0; j < t.cols; j++) set(t.sheet, k, t.c0 + j, get(t.sheet, k + 1, t.c0 + j));
        for (let j = 0; j < t.cols; j++) set(t.sheet, t.r0 + t.rows - 1, t.c0 + j, ''); t.rows--; } }) },
    set name(n){ t.name = n; }, get name(){ return t.name; } });
  const sheetApi = name => ({ name, getRange: a => range(name, a), getUsedRange: () => ({ format: { autofitColumns(){} } }), set visibility(v){ sheetApi.vis = v; },
    tables: { add: (addr, _h) => { const [a, b] = addr.split(':'), [r0, c0] = A1(a), [r1, c1] = A1(b); const t = { name: 'Table' + (tables.length + 1), sheet: name, r0, c0, rows: r1 - r0 + 1, cols: c1 - c0 + 1 }; tables.push(t); return tableApi(t); } } });
  const ctx = { sync: async () => {}, workbook: {
    worksheets: { getItemOrNullObject: n => sheets[n] ? { ...sheetApi(n), isNullObject: false } : { isNullObject: true }, add: n => { cells(n); return sheetApi(n); } },
    tables: { load(){}, get items(){ return tables.map(t => ({ name: t.name })); },
      getItem: n => { const t = tables.find(x => x.name === n); if (!t) throw new Error('no table ' + n); return tableApi(t); },
      getItemOrNullObject: n => { const t = tables.find(x => x.name === n); return t ? { ...tableApi(t), isNullObject: false } : { isNullObject: true }; } } } };
  return { sheets, tables, get, set, ctx };
}
const run = wb => { const C = { Excel: { run: fn => fn(wb.ctx), SheetVisibility: { hidden: 'Hidden' } } }; vm.createContext(C); vm.runInContext(code + '\nObject.assign(globalThis, { ensureNotifySettings, notifySettingsStatus, readNotifySettings, setNotifyConfig, upsertNotifyRoute, removeNotifyRoute });', C); return C; };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

(async () => {
  await t('new workbook: Settings sheet with the four tables where the spec says', async () => {
    const wb = workbook(), C = run(wb), r = await C.ensureNotifySettings();
    assert.deepStrictEqual([...r.created], ['tblNotifyConfig', 'tblNotifyRoutes', 'tblNotifyTemplates', 'tblNotifyLog']);
    const at = nm => { const x = wb.tables.find(y => y.name === nm); return [x.r0, x.c0, x.rows, x.cols].join(','); };
    assert.strictEqual(at('tblNotifyConfig'), '0,0,12,2'); assert.strictEqual(at('tblNotifyRoutes'), '0,3,2,4'); assert.strictEqual(at('tblNotifyTemplates'), '0,8,3,3'); assert.strictEqual(at('tblNotifyLog'), '0,12,2,10');
    assert.strictEqual(wb.get('Settings', 0, 4), 'FlowUrl'); assert.strictEqual(wb.get('Settings', 0, 5), 'Email'); assert.strictEqual(wb.get('Settings', 1, 0), 'Enabled');
    const again = await C.ensureNotifySettings(); assert.strictEqual(again.created.length, 0); assert.strictEqual(again.existing.length, 4, 'safe to call again');
    const st = await C.notifySettingsStatus(); assert(st.every(x => x.present));
  });
  await t('an existing Settings sheet with something in the way is left alone', async () => {
    const wb = workbook(); wb.set('Settings', 0, 0, 'my notes'); const C = run(wb), r = await C.ensureNotifySettings();
    assert.strictEqual(wb.get('Settings', 0, 0), 'my notes', 'not overwritten');
    assert(r.blocked.some(x => x.startsWith('tblNotifyConfig')) && r.created.includes('tblNotifyRoutes'), 'blocked one, made the rest');
  });
  await t('provider rows: the starter row is reused, then add, update (any case), remove', async () => {
    const wb = workbook(), C = run(wb); await C.ensureNotifySettings();
    await C.upsertNotifyRoute('Dr A', 'https://teams.microsoft.com/l/chat/19:aaa@thread.v2/0', 'dra@example.com', true);
    let d = await C.readNotifySettings(); assert.strictEqual(d.routes.length, 1); assert.strictEqual(d.routes[0].email, 'dra@example.com'); assert(/19:aaa/.test(d.routes[0].flowUrl));
    await C.upsertNotifyRoute('Dr B', '19:bbb@thread.v2', 'drb@example.com', false);
    await C.upsertNotifyRoute('dr a', 'https://teams.microsoft.com/l/chat/19:new@thread.v2/0', 'a2@example.com', true);
    d = await C.readNotifySettings();
    assert.deepStrictEqual([...d.routes.map(r => [r.provider, r.email, r.enabled].join('|'))], ['dr a|a2@example.com|true', 'Dr B|drb@example.com|false']);
    assert.strictEqual(await C.removeNotifyRoute('Dr B'), true); assert.strictEqual(await C.removeNotifyRoute('Dr Z'), false);
    await C.removeNotifyRoute('dr a'); d = await C.readNotifySettings(); assert.strictEqual(d.routes.length, 0, 'last row blanked, table kept');
    assert(wb.tables.find(x => x.name === 'tblNotifyRoutes').rows >= 2);
  });
  await t('config: set a value, add a missing key', async () => {
    const wb = workbook(), C = run(wb); await C.ensureNotifySettings();
    await C.setNotifyConfig('EscalationChat', 'https://example.invalid/flow'); await C.setNotifyConfig('NewKey', 'x');
    const d = await C.readNotifySettings(); assert.strictEqual(d.config.EscalationChat, 'https://example.invalid/flow'); assert.strictEqual(d.config.NewKey, 'x'); assert.strictEqual(d.config.Enabled, 'TRUE');
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
