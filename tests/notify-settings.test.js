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
    columns: { getItemAt: i => ({ set name(n){ set(t.sheet, t.r0, t.c0 + i, n); } }),
      add: (_, __, name) => { for (let i = 0; i < t.rows; i++) if (get(t.sheet, t.r0 + i, t.c0 + t.cols) !== '') throw new Error('cells in the way');
        set(t.sheet, t.r0, t.c0 + t.cols, name); t.cols++; } },
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
const run = wb => { const C = { Excel: { run: fn => fn(wb.ctx), SheetVisibility: { hidden: 'Hidden' } } }; vm.createContext(C); vm.runInContext(code + '\nObject.assign(globalThis, { ensureNotifySettings, notifySettingsStatus, readNotifySettings, setNotifyConfig, upsertNotifyRoute, removeNotifyRoute, nfyChatId });', C); return C; };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

// An older workbook: the one-workflow-per-provider layout the previous add-in made.
function oldLayout(wb){
  const put = (r0, c0, rows) => rows.forEach((row, i) => row.forEach((v, j) => wb.set('Settings', r0 + i, c0 + j, v)));
  const mk = (name, c0, rows) => { put(0, c0, rows); wb.tables.push({ name, sheet: 'Settings', r0: 0, c0, rows: rows.length, cols: rows[0].length }); };
  mk('tblNotifyConfig', 0, [['Key', 'Value'], ['Enabled', 'TRUE'], ['SendGroupChat', 'TRUE'], ['EscalationChat', ''], ['NotifyRecheck', 'FALSE']]);
  mk('tblNotifyRoutes', 3, [['Provider', 'FlowUrl', 'Email', 'Enabled'], ['Dr A', 'https://x.environment.api.powerplatform.com/old?sig=1', 'dra@example.com', 'TRUE']]);
  mk('tblNotifyTemplates', 8, [['Event', 'Template', 'Enabled'], ['standard', '{patient} {stage} vitals {vitals} - Bay {bay}', 'TRUE'],
    ['emergency', '<b>EMERGENCY</b> {patient} {stage} vitals {vitals} - Bay {bay} ({provider})', 'TRUE']]);
  mk('tblNotifyLog', 12, [['Id', 'Provider', 'Patient', 'Bay', 'Stage', 'Vitals', 'Level', 'SentBy', 'SentAt', 'Seen'], ['', '', '', '', '', '', '', '', '', '']]);
}

(async () => {
  await t('new workbook: Settings sheet with the four tables where the spec says', async () => {
    const wb = workbook(), C = run(wb), r = await C.ensureNotifySettings();
    assert.deepStrictEqual([...r.created], ['tblNotifyConfig', 'tblNotifyRoutes', 'tblNotifyTemplates', 'tblNotifyLog']);
    const at = nm => { const x = wb.tables.find(y => y.name === nm); return [x.r0, x.c0, x.rows, x.cols].join(','); };
    assert.strictEqual(at('tblNotifyConfig'), '0,0,10,2'); assert.strictEqual(at('tblNotifyRoutes'), '0,3,2,5'); assert.strictEqual(at('tblNotifyTemplates'), '0,8,3,3'); assert.strictEqual(at('tblNotifyLog'), '0,12,2,10');
    assert.deepStrictEqual([4, 5, 6, 7].map(c => wb.get('Settings', 0, c)), ['ChatLink', 'MentionName', 'MentionId', 'Enabled']);
    assert.strictEqual(wb.get('Settings', 2, 0), 'WebhookUrl'); assert.strictEqual(wb.get('Settings', 3, 0), 'MentionProvider');
    assert.strictEqual(wb.get('Settings', 1, 9), '{patient} is ready in Bay {bay}'); assert.strictEqual(wb.get('Settings', 2, 9), 'EMERGENCY: {patient} in Bay {bay}');
    const again = await C.ensureNotifySettings(); assert.strictEqual(again.created.length, 0); assert.strictEqual(again.existing.length, 4, 'safe to call again'); assert.strictEqual(again.upgraded.length, 0);
    const st = await C.notifySettingsStatus(); assert(st.every(x => x.present)); assert.strictEqual((await C.readNotifySettings()).needsUpgrade, false);
  });
  await t('an existing Settings sheet with something in the way is left alone', async () => {
    const wb = workbook(); wb.set('Settings', 0, 0, 'my notes'); const C = run(wb), r = await C.ensureNotifySettings();
    assert.strictEqual(wb.get('Settings', 0, 0), 'my notes', 'not overwritten');
    assert(r.blocked.some(x => x.startsWith('tblNotifyConfig')) && r.created.includes('tblNotifyRoutes'), 'blocked one, made the rest');
  });
  await t('older workbook: FlowUrl -> ChatLink, Email -> MentionId, MentionName added, new config keys, old default templates replaced', async () => {
    const wb = workbook(); oldLayout(wb); const C = run(wb);
    assert.strictEqual((await C.readNotifySettings()).needsUpgrade, true);
    const r = await C.ensureNotifySettings(); assert.strictEqual(r.blocked.length, 0, r.blocked.join());
    assert.deepStrictEqual([3, 4, 5, 6, 7].map(c => wb.get('Settings', 0, c)), ['Provider', 'ChatLink', 'MentionId', 'Enabled', 'MentionName']);
    assert.strictEqual(wb.get('Settings', 1, 5), 'dra@example.com', 'their email is kept as the mention id');
    const d = await C.readNotifySettings(); assert.strictEqual(d.config.WebhookUrl, ''); assert.strictEqual(d.config.MentionProvider, 'TRUE'); assert.strictEqual(d.config.NotifyRecheck, 'FALSE', 'existing values kept');
    assert.strictEqual(d.routes[0].mentionId, 'dra@example.com'); assert.strictEqual(d.routes[0].email, 'dra@example.com'); assert.strictEqual(d.routes[0].chatId, '', 'an old workflow URL is not a chat');
    assert.strictEqual(wb.get('Settings', 2, 9), 'EMERGENCY: {patient} in Bay {bay}'); assert.strictEqual(d.needsUpgrade, false);
    assert.strictEqual((await C.ensureNotifySettings()).upgraded.length, 0, 'second run changes nothing');
  });
  await t('older workbook with a table right next to the routes: the column is not forced in, and that is reported', async () => {
    const wb = workbook(); oldLayout(wb); wb.set('Settings', 0, 7, 'x'); const C = run(wb);
    const r = await C.ensureNotifySettings(); assert(r.blocked.some(b => /MentionName/.test(b)));
  });
  await t('a custom template is never replaced', async () => {
    const wb = workbook(); oldLayout(wb); wb.set('Settings', 1, 9, 'My own wording {patient}'); const C = run(wb); await C.ensureNotifySettings();
    assert.strictEqual(wb.get('Settings', 1, 9), 'My own wording {patient}');
  });
  await t('provider rows: the starter row is reused, then add, update (any case), remove; chat id pulled from the link', async () => {
    const wb = workbook(), C = run(wb); await C.ensureNotifySettings();
    await C.upsertNotifyRoute({ provider: 'Dr A', chatLink: 'https://teams.microsoft.com/l/chat/19%3Aaaa%40thread.v2/0?tenantId=x', mentionName: 'Alex A', mentionId: 'dra@example.com', enabled: true });
    let d = await C.readNotifySettings(); assert.strictEqual(d.routes.length, 1); assert.strictEqual(d.routes[0].chatId, '19:aaa@thread.v2'); assert.strictEqual(d.routes[0].mentionName, 'Alex A');
    await C.upsertNotifyRoute({ provider: 'Dr B', chatLink: '19:bbb@thread.v2', mentionName: '', mentionId: '00000000-0000-0000-0000-000000000001', enabled: false });
    await C.upsertNotifyRoute({ provider: 'dr a', chatLink: '19:new@thread.v2', mentionName: 'Alex A', mentionId: 'a2@example.com', enabled: true });
    d = await C.readNotifySettings();
    assert.deepStrictEqual([...d.routes.map(r => [r.provider, r.chatId, r.mentionId, r.email, r.enabled].join('|'))],
      ['dr a|19:new@thread.v2|a2@example.com|a2@example.com|true', 'Dr B|19:bbb@thread.v2|00000000-0000-0000-0000-000000000001||false']);
    assert.strictEqual(await C.removeNotifyRoute('Dr B'), true); assert.strictEqual(await C.removeNotifyRoute('Dr Z'), false);
    await C.removeNotifyRoute('dr a'); d = await C.readNotifySettings(); assert.strictEqual(d.routes.length, 0, 'last row blanked, table kept');
    assert(wb.tables.find(x => x.name === 'tblNotifyRoutes').rows >= 2);
  });
  await t('upsert writes by header on an upgraded (reordered) table', async () => {
    const wb = workbook(); oldLayout(wb); const C = run(wb); await C.ensureNotifySettings();
    await C.upsertNotifyRoute({ provider: 'Dr A', chatLink: '19:aaa@thread.v2', mentionName: 'Alex A', mentionId: 'dra@example.com', enabled: true });
    assert.deepStrictEqual([3, 4, 5, 6, 7].map(c => wb.get('Settings', 1, c)), ['Dr A', '19:aaa@thread.v2', 'dra@example.com', 'TRUE', 'Alex A']);
  });
  await t('chat id: from an encoded or plain link, or the id itself; nothing from a workflow URL', async () => {
    const C = run(workbook());
    assert.strictEqual(C.nfyChatId('https://teams.microsoft.com/l/chat/19:abc_def@thread.v2/conversations?context=x'), '19:abc_def@thread.v2');
    assert.strictEqual(C.nfyChatId('19:meeting_x@thread.v2'), '19:meeting_x@thread.v2'); assert.strictEqual(C.nfyChatId('https://x.powerplatform.com/a?sig=1'), '');
  });
  await t('config: set a value, add a missing key', async () => {
    const wb = workbook(), C = run(wb); await C.ensureNotifySettings();
    await C.setNotifyConfig('WebhookUrl', 'https://example.invalid/flow'); await C.setNotifyConfig('NewKey', 'x');
    const d = await C.readNotifySettings(); assert.strictEqual(d.config.WebhookUrl, 'https://example.invalid/flow'); assert.strictEqual(d.config.NewKey, 'x'); assert.strictEqual(d.config.Enabled, 'TRUE');
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
