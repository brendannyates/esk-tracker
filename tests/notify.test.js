// Teams notifications (notify.js) on top of the Settings sheet (notify-settings.js), with the mocked Excel from notify-settings.test.js.
//   node tests/notify.test.js
// Checks the rules (initial always; later only past 140/90 or under SpO2 92; emergency at 180/120), the office toggles, the provider
// table read by header (Email column present), and what's posted to each provider's workflow. fetch is captured; nothing leaves the machine.
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
const FLOW_A = 'https://default00.11.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/aaa/triggers/manual/paths/invoke?api-version=1&sig=TEST';
const FLOW_B = 'https://prod-01.westus.logic.azure.com:443/workflows/bbb/triggers/manual/paths/invoke?sig=TEST';
const FLOW_ESC = 'https://default00.11.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/esc/triggers/manual/paths/invoke?sig=TEST';
const base = { provider: 'Dr A', patient: 'SMIJAN', bay: '2', clinic: 'SRL', sentBy: 'Tech T' };
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

(async () => {
  const { C, posts, wb } = setup();
  await C.ensureNotifySettings();
  await C.upsertNotifyRoute('Dr A', FLOW_A, 'dra@example.com', true);
  await C.upsertNotifyRoute('Dr B', FLOW_B, 'drb@example.com', false);
  await C.upsertNotifyRoute('Dr C', 'https://teams.microsoft.com/l/chat/19:c@thread.v2/0', '', true);

  await t('provider table read by header: Email in column 3 does not hide Enabled', async () => {
    const s = await C.loadNotifySettings();
    assert.strictEqual(s.routes['dr a'].flow, FLOW_A); assert.strictEqual(s.routes['dr a'].email, 'dra@example.com');
    assert.strictEqual(s.routes['dr b'], undefined, 'Dr B is off');
    assert.strictEqual(s.routes['dr c'].flow, '', 'a chat link is not a workflow URL');
    assert.strictEqual(C.nfyFlowUrl(FLOW_B), FLOW_B); assert.strictEqual(C.nfyFlowUrl('https://evil.example/powerplatform.com/x'), '');
  });
  await t('initial vitals go to that provider’s workflow as an Adaptive Card, in a simple text/plain no-cors request', async () => {
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '128/82', hr: '74', spo2: '98%' } });
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'standard');
    const p = posts.at(-1); assert.strictEqual(p.url, FLOW_A); assert.strictEqual(p.o.mode, 'no-cors'); assert(/text\/plain/.test(p.o.headers['Content-Type']));
    assert.strictEqual(p.body.message, 'SMIJAN initial vitals 128/82, HR 74, SpO2 98% - Bay 2');
    assert.strictEqual(p.body.type, 'message'); const card = p.body.attachments[0]; assert.strictEqual(card.contentType, 'application/vnd.microsoft.card.adaptive');
    assert.strictEqual(card.content.type, 'AdaptiveCard'); assert.strictEqual(card.content.body[0].text, p.body.message, 'no red heading for a standard alert');
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
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'emergency'); const b = posts.at(-1).body; assert(/^EMERGENCY SMIJAN/.test(b.message), 'plain text, no tags');
    assert.strictEqual(b.attachments[0].content.body[0].color, 'Attention'); assert(/\*\*EMERGENCY\*\*/.test(b.attachments[0].content.body[1].text), '<b> becomes markdown bold');
  });
  await t('every send is logged for the in-app alerts; a disabled provider gets nothing', async () => {
    const log = (await C.readNotifyLogSince('2000')).filter(r => r.provider === 'Dr A'); assert(log.length >= 4);
    const r = await C.onVitalsSaved({ ...base, provider: 'Dr B', stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(r.sent, false); assert(/No Teams workflow URL/.test(r.reason));
    assert(/No Teams workflow URL/.test((await C.onVitalsSaved({ ...base, provider: 'Dr C', stage: 'initial', vitals: { bp: '120/80' } })).reason), 'chat link only');
  });
  await t('initial vitals have their own toggle (on by default); an emergency still goes', async () => {
    await C.setNotifyConfig('NotifyInitial', 'FALSE');
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '128/82' } }); assert.strictEqual(r.sent, false); assert(/initial alerts are off/.test(r.reason));
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '186/112' } })).level, 'emergency');
    await C.setNotifyConfig('NotifyInitial', 'TRUE');
  });
  await t('escalation: an emergency also goes to the escalation workflow when that is on; standard alerts never do', async () => {
    await C.setNotifyConfig('EscalationChat', FLOW_ESC); await C.setNotifyConfig('EscalateEmergency', 'TRUE');
    let k = posts.length; await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '150/92' } });
    assert.deepStrictEqual(posts.slice(k).map(p => p.url), [FLOW_A]);
    k = posts.length; await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '190/100' } });
    assert.deepStrictEqual(posts.slice(k).map(p => p.url), [FLOW_A, FLOW_ESC]);
    await C.setNotifyConfig('SendGroupChat', 'FALSE'); k = posts.length; await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '190/100' } });
    assert.deepStrictEqual(posts.slice(k).map(p => p.url), [FLOW_ESC], 'group chat off: escalation only');
    await C.setNotifyConfig('SendGroupChat', 'TRUE'); await C.setNotifyConfig('EscalateEmergency', 'FALSE');
  });
  await t('test send: one post, no log row, no patient', async () => {
    const before = (await C.readNotifyLogSince('2000')).length, k = posts.length;
    const r = await C.sendNotifyTest('Dr A', 'Tech T'); assert.strictEqual(r.ok, true); assert.strictEqual(posts.length, k + 1);
    assert.strictEqual(posts.at(-1).url, FLOW_A); assert.strictEqual(posts.at(-1).body.event, 'test'); assert.strictEqual(posts.at(-1).body.patient, ''); assert(posts.at(-1).body.attachments[0].content.body[0].text.startsWith('Test from ESK Tracker'));
    assert.strictEqual((await C.readNotifyLogSince('2000')).length, before);
    assert(/No Teams workflow URL/.test((await C.sendNotifyTest('Dr Z')).reason));
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
