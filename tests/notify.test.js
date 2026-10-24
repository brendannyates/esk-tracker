// Teams notifications (notify.js) on top of the Settings sheet (notify-settings.js), with the mocked Excel from notify-settings.test.js.
//   node tests/notify.test.js
// Checks the rules (initial always; later only past 140/90 or under SpO2 92; emergency at 180/120), the office toggles, the provider
// table read by header (Email column present), and what's posted to the flow. fetch is captured; nothing leaves the machine.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
const mockSrc = read('tests/notify-settings.test.js').match(/^function workbook\(\)\{[\s\S]*?^\}$/m)[0];
const workbook = vm.runInNewContext('(' + mockSrc + ')');
function setup(){
  const wb = workbook(), posts = [];
  const C = { Excel: { run: fn => fn(wb.ctx), SheetVisibility: {} }, fetch: async (url, o) => { posts.push({ url, o, body: JSON.parse(o.body) }); return {}; }, console, Date, Math, JSON, String, Object, Array, RegExp, decodeURIComponent, window: {} };
  vm.createContext(C); vm.runInContext(read('notify-settings.js') + '\n' + read('notify.js'), C);
  return { wb, C, posts };
}
const CHAT = 'https://teams.microsoft.com/l/chat/19:abc123@thread.v2/conversations?context=%7B%7D';
const base = { provider: 'Dr A', patient: 'SMIJAN', bay: '2', clinic: 'SRL', sentBy: 'Tech T' };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

(async () => {
  const { C, posts, wb } = setup();
  await C.ensureNotifySettings();
  await C.upsertNotifyRoute('Dr A', CHAT, 'dra@example.com', true);
  await C.upsertNotifyRoute('Dr B', '19:zzz@thread.v2', 'drb@example.com', false);
  await C.setNotifyConfig('WebhookUrl', 'https://example.invalid/flow');

  await t('provider table read by header: Email in column 3 does not hide Enabled', async () => {
    const s = await C.loadNotifySettings();
    assert.deepStrictEqual([...s.routes['dr a'].chats], ['19:abc123@thread.v2']); assert.strictEqual(s.routes['dr a'].email, 'dra@example.com');
    assert.strictEqual(s.routes['dr b'], undefined, 'Dr B is off');
  });
  await t('initial vitals always go; the post is a simple text/plain no-cors request with the chat and message', async () => {
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '128/82', hr: '74', spo2: '98%' } });
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'standard');
    const p = posts.at(-1); assert.strictEqual(p.o.mode, 'no-cors'); assert(/text\/plain/.test(p.o.headers['Content-Type']));
    assert.deepStrictEqual([...p.body.chats], ['19:abc123@thread.v2']); assert.strictEqual(p.body.message, 'SMIJAN initial vitals 128/82, HR 74, SpO2 98% - Bay 2');
    assert.strictEqual(p.body.patient, 'SMIJAN');
  });
  await t('later vitals in range: nothing sent; past 140/90 or under SpO2 92: sent', async () => {
    const k = posts.length;
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '138/88', spo2: '96%' } })).sent, false);
    assert.strictEqual(posts.length, k, 'no post');
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '150/92' } })).sent, true);
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'discharge', vitals: { bp: '130/80', spo2: '90%' } })).sent, true);
  });
  await t('emergency at 180/120 goes as emergency, even with that stage turned off', async () => {
    await C.setNotifyConfig('NotifyRecheck', 'FALSE');
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'recheck', vitals: { bp: '150/95' } })).sent, false, 'recheck alerts off');
    const r = await C.onVitalsSaved({ ...base, stage: 'recheck', vitals: { bp: '182/100' } });
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'emergency'); assert(/EMERGENCY/.test(posts.at(-1).body.message));
  });
  await t('every send is logged for the in-app alerts; a disabled provider gets nothing', async () => {
    const log = (await C.readNotifyLogSince('2000')).filter(r => r.provider === 'Dr A'); assert(log.length >= 4);
    const r = await C.onVitalsSaved({ ...base, provider: 'Dr B', stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(r.sent, false); assert(/No Teams group chat/.test(r.reason));
  });
  await t('initial vitals have their own toggle (on by default); an emergency still goes', async () => {
    await C.setNotifyConfig('NotifyInitial', 'FALSE');
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '128/82' } }); assert.strictEqual(r.sent, false); assert(/initial alerts are off/.test(r.reason));
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '186/112' } })).level, 'emergency');
    await C.setNotifyConfig('NotifyInitial', 'TRUE');
  });
  await t('test send: one post, no log row, no patient', async () => {
    const before = (await C.readNotifyLogSince('2000')).length, k = posts.length;
    const r = await C.sendNotifyTest('Dr A', 'Tech T'); assert.strictEqual(r.ok, true); assert.strictEqual(posts.length, k + 1);
    assert.strictEqual(posts.at(-1).body.event, 'test'); assert.strictEqual(posts.at(-1).body.patient, '');
    assert.strictEqual((await C.readNotifyLogSince('2000')).length, before);
    assert(/No Teams group chat ID/.test((await C.sendNotifyTest('Dr Z')).reason));
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
