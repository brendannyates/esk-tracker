// Teams notifications (notify.js) on top of the Settings sheet (notify-settings.js), with the mocked Excel from notify-settings.test.js.
//   node tests/notify.test.js
// Checks the rules (initial always; later only past 140/90 or under SpO2 92; emergency at 180/120), the office toggles, the provider
// table read by header, the card (title, Stage, flagged vitals, Sent By, @mention) and the payload posted to the one flow.
// fetch is captured; nothing leaves the machine.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
const mockSrc = read('tests/notify-settings.test.js').match(/^function workbook\(\)\{[\s\S]*?^\}$/m)[0];
const workbook = vm.runInNewContext('(' + mockSrc + ')');
function setup(){
  const wb = workbook(), posts = [];
  const C = { Excel: { run: fn => fn(wb.ctx), SheetVisibility: {} }, fetch: async (url, o) => { posts.push({ url, o, body: JSON.parse(o.body) }); const f = C.__answer; if (f) return f(url, o); return { ok: true, status: 202, text: async () => '' }; }, console, Date, Math, JSON, String, Object, Array, RegExp, decodeURIComponent, window: {} };
  vm.createContext(C); vm.runInContext(read('notify-settings.js') + '\n' + read('notify.js'), C);
  return { wb, C, posts };
}
const FLOW = 'https://default00.11.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/aaa/triggers/manual/paths/invoke?api-version=1&sig=TEST';
const CHAT_A = '19:aaa@thread.v2', CHAT_C = '19:ccc@thread.v2';
const base = { provider: 'Dr A', patient: 'SMIJAN', bay: '2', clinic: 'SRL', sentBy: 'Tech T', sn: '2404' };
const texts = card => card.body.map(b => b.text != null ? b.text : b.inlines.map(r => r.text).join(''));
const flags = card => card.body.flatMap(b => b.inlines || []).filter(r => r.weight === 'Bolder').map(r => r.text + '|' + r.color);
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

(async () => {
  const { C, posts } = setup();
  await C.ensureNotifySettings();
  await t('no flow URL yet: nothing posted, and the reason says where to set it', async () => {
    await C.upsertNotifyRoute({ provider: 'Dr A', chatLink: 'https://teams.microsoft.com/l/chat/19%3Aaaa%40thread.v2/0', mentionName: 'Alex Able', mentionId: 'dra@example.com', enabled: true });
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(r.sent, false); assert(/flow URL isn’t set/.test(r.reason)); assert.strictEqual(posts.length, 0);
  });
  await C.setNotifyConfig('WebhookUrl', FLOW);
  await C.upsertNotifyRoute({ provider: 'Dr B', chatLink: '19:bbb@thread.v2', mentionName: '', mentionId: '', enabled: false });
  await C.upsertNotifyRoute({ provider: 'Dr C', chatLink: CHAT_C, mentionName: '', mentionId: '', enabled: true });
  await C.upsertNotifyRoute({ provider: 'Dr D', chatLink: 'https://x.environment.api.powerplatform.com/old?sig=1', mentionName: '', mentionId: '', enabled: true });

  await t('provider table read by header: chat id from the link, mention fields, Enabled', async () => {
    const s = await C.loadNotifySettings();
    assert.strictEqual(s.webhook, FLOW); assert.strictEqual(s.mention, true);
    assert.deepStrictEqual([...s.routes['dr a'].chats], [CHAT_A]); assert.strictEqual(s.routes['dr a'].mentionId, 'dra@example.com'); assert.strictEqual(s.routes['dr a'].mentionName, 'Alex Able');
    assert.strictEqual(s.routes['dr b'].chats.length, 0, 'Dr B is off'); assert.strictEqual(s.routes['dr d'].chats.length, 0, 'an old workflow URL is not a chat');
  });
  await t('initial vitals: one post to the flow with {event, chats, card} and the card as the build sheet lays it out', async () => {
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '116/70', hr: '69', spo2: '99%' } });
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'standard'); assert.strictEqual(r.message, 'SMIJAN is ready in Bay 2');
    const p = posts.at(-1); assert.strictEqual(p.url, FLOW); assert.strictEqual(p.o.headers['Content-Type'], 'application/json');
    const b = p.body; assert.strictEqual(b.event, 'vitals'); assert.deepStrictEqual([...b.chats], [CHAT_A]); assert.strictEqual(b.mentionId, 'dra@example.com');
    assert.strictEqual(b.card.type, 'AdaptiveCard'); assert.strictEqual(b.card.version, '1.4'); assert.strictEqual(b.card.body[0].type, 'TextBlock'); assert.strictEqual(b.card.body[1].type, 'RichTextBlock'); assert.strictEqual(b.type, undefined, 'not wrapped for the template');
    assert.deepStrictEqual([...texts(b.card)], ['<at>Alex Able</at> SMIJAN is ready in Bay 2', 'Stage: Initial Vitals', 'BP: 116/70, HR: 69, SpO2: 99%', 'SN: 2404 | Sent By: Tech T']);
    assert.strictEqual(b.card.body[0].weight, 'Bolder'); assert.strictEqual(b.card.body[3].inlines[0].isSubtle, true); assert.strictEqual(flags(b.card).length, 0, 'nothing flagged');
    assert.deepStrictEqual([...b.card.body.map(x => x.spacing)], ['Default', 'Medium', 'Medium', 'Medium'], 'space between the title and each line');
    assert.deepStrictEqual([...b.card.body.map(x => !!x.separator)], [false, true, true, true], 'a separator line between sections');
    const e = b.card.msteams.entities[0]; assert.strictEqual(e.text, '<at>Alex Able</at>'); assert.strictEqual(e.mentioned.id, 'dra@example.com'); assert.strictEqual(e.mentioned.name, 'Alex Able');
    assert(!texts(b.card).some(x => /^Provider/.test(x)), 'no Provider row');
    assert.strictEqual(b.message, 'SMIJAN is ready in Bay 2<br>Stage: Initial Vitals<br>BP: 116/70, HR: 69, SpO2: 99%<br>SN: 2404 | Sent By: Tech T'); assert.strictEqual(b.sn, '2404');
    assert.strictEqual(b.patient, 'SMIJAN'); assert.strictEqual(b.source, 'ESK Tracker');
  });
  await t('later vitals in range: nothing sent; HIGH and LOW flags sit on the vital that tripped', async () => {
    const k = posts.length;
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '138/88', spo2: '96%' } })).sent, false);
    assert.strictEqual(posts.length, k, 'no post');
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: '40-min', vitals: { bp: '142/87', hr: '69', spo2: '99%' } })).sent, true);
    assert.strictEqual(texts(posts.at(-1).body.card)[2], 'BP: 142/87 HIGH (over 140/90), HR: 69, SpO2: 99%');
    assert.deepStrictEqual([...flags(posts.at(-1).body.card)], ['HIGH|Attention'], 'HIGH in red');
    assert.strictEqual(texts(posts.at(-1).body.card)[1], 'Stage: 40-min Vitals');
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'discharge', vitals: { bp: '128/82', hr: '80', spo2: '90%' } })).sent, true);
    assert.strictEqual(texts(posts.at(-1).body.card)[2], 'BP: 128/82, HR: 80, SpO2: 90% LOW (under 92%)');
    assert.deepStrictEqual([...flags(posts.at(-1).body.card)], ['LOW|Accent'], 'LOW in blue');
    assert(/<b>LOW<\/b>/.test(posts.at(-1).body.message), 'HTML twin is bold too');
  });
  await t('emergency at 180/120: emergency title, EMERGENCY on the BP line, goes even with that stage turned off', async () => {
    await C.setNotifyConfig('NotifyRecheck', 'FALSE');
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'recheck', vitals: { bp: '150/95' } })).sent, false, 'recheck alerts off');
    const r = await C.onVitalsSaved({ ...base, stage: 'recheck', vitals: { bp: '184/122', hr: '96', spo2: '96%' } });
    assert.strictEqual(r.sent, true); assert.strictEqual(r.level, 'emergency'); const b = posts.at(-1).body;
    assert.strictEqual(b.level, 'emergency'); assert(!b.card.body.some(x => x.type === 'Container'), 'plain card, only the word is colored');
    assert.deepStrictEqual([...texts(b.card)].slice(0, 3), ['<at>Alex Able</at> EMERGENCY: SMIJAN in Bay 2', 'Stage: Recheck Vitals *EMERGENCY* at or above 180/120', 'BP: 184/122, HR: 96, SpO2: 96%']);
    assert.deepStrictEqual([...flags(b.card)], ['*EMERGENCY*|Attention'], 'asterisks shown as written, in red');
    assert(/Stage: Recheck Vitals <b>\*EMERGENCY\*<\/b> at or above 180\/120<br>BP: 184\/122, HR: 96/.test(b.message), b.message);
    await C.setNotifyConfig('NotifyRecheck', 'TRUE');
  });
  await t('MentionProvider off, or no MentionId: no mention, title stays plain', async () => {
    await C.setNotifyConfig('MentionProvider', 'FALSE');
    await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '120/80' } });
    let b = posts.at(-1).body; assert.strictEqual(texts(b.card)[0], 'SMIJAN is ready in Bay 2'); assert.strictEqual(b.card.msteams, undefined); assert.strictEqual(b.mentionId, '');
    await C.setNotifyConfig('MentionProvider', 'TRUE');
    await C.onVitalsSaved({ ...base, provider: 'Dr C', stage: 'initial', vitals: { bp: '120/80' } });
    b = posts.at(-1).body; assert.deepStrictEqual([...b.chats], [CHAT_C]); assert.strictEqual(b.card.msteams, undefined);
  });
  await t('free text cannot format the card: markdown characters in the title are escaped; other lines are plain text; old HTML templates lose their tags', async () => {
    await C.onVitalsSaved({ ...base, patient: 'SMI_JAN*', sentBy: 'T**', stage: 'initial', vitals: { bp: '120/80' } });
    const x = texts(posts.at(-1).body.card); assert(x[0].endsWith('SMI\\_JAN\\* is ready in Bay 2'), x[0]); assert.strictEqual(x.at(-1), 'SN: 2404 | Sent By: T**', 'plain-text runs show it as typed');
    await C.onVitalsSaved({ ...base, sn: '', stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(texts(posts.at(-1).body.card).at(-1), 'Sent By: Tech T', 'no SN known: Sent By alone');
    assert.strictEqual(C.nfyRender('<b>EMERGENCY</b> {patient} - Bay {bay}', { patient: 'SMIJAN', bay: 2 }), 'EMERGENCY SMIJAN - Bay 2');
  });
  await t('every send is logged for the in-app alerts; a disabled provider or one without a chat gets nothing', async () => {
    const log = (await C.readNotifyLogSince('2000')).filter(r => r.provider === 'Dr A'); assert(log.length >= 4); assert.strictEqual(log[0].vitals, '116/70, HR 69, SpO2 99%');
    let r = await C.onVitalsSaved({ ...base, provider: 'Dr B', stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(r.sent, false); assert(/turned off/.test(r.reason));
    r = await C.onVitalsSaved({ ...base, provider: 'Dr D', stage: 'initial', vitals: { bp: '120/80' } }); assert(/No group chat link/.test(r.reason), r.reason);
    r = await C.onVitalsSaved({ ...base, provider: 'Dr Z', stage: 'initial', vitals: { bp: '120/80' } }); assert(/No group chat set/.test(r.reason), r.reason);
  });
  await t('initial vitals have their own toggle (on by default); an emergency still goes', async () => {
    await C.setNotifyConfig('NotifyInitial', 'FALSE');
    const r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '128/82' } }); assert.strictEqual(r.sent, false); assert(/initial alerts are off/.test(r.reason));
    assert.strictEqual((await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '186/112' } })).level, 'emergency');
    await C.setNotifyConfig('NotifyInitial', 'TRUE');
  });
  await t('the flow’s answer comes back: 404 and a disabled flow are reported; a blocked request falls back to no-cors', async () => {
    C.__answer = async () => ({ ok: false, status: 404, text: async () => '{"error":{"code":"WorkflowNotFound","message":"x"}}' });
    let r = await C.sendNotifyTest('Dr A', 'T'); assert.strictEqual(r.ok, false); assert(/answered 404 \(WorkflowNotFound\): the URL is wrong/.test(r.reason), r.reason);
    C.__answer = async () => ({ ok: false, status: 400, text: async () => '{"error":{"code":"WorkflowDisabled","message":"The workflow is disabled."}}' });
    r = await C.onVitalsSaved({ ...base, stage: 'initial', vitals: { bp: '120/80' } }); assert.strictEqual(r.sent, false); assert(/turned off/.test(r.reason), r.reason);
    let k = 0; C.__answer = async (u, o) => { if (o.mode !== 'no-cors') throw new TypeError('Failed to fetch'); k++; return {}; };
    r = await C.sendNotifyTest('Dr A', 'T'); assert.strictEqual(r.ok, true); assert.strictEqual(r.unconfirmed, true); assert.strictEqual(k, 1);
    assert(/text\/plain/.test(posts.at(-1).o.headers['Content-Type']));
    C.__answer = null; r = await C.sendNotifyTest('Dr A', 'T'); assert.strictEqual(r.status, 202); assert(!r.unconfirmed);
  });
  await t('test send: one post to that provider’s chat with their mention, no log row, no patient', async () => {
    const before = (await C.readNotifyLogSince('2000')).length, k = posts.length;
    const r = await C.sendNotifyTest('Dr A', 'Tech T'); assert.strictEqual(r.ok, true); assert.strictEqual(r.mentioned, true); assert.strictEqual(posts.length, k + 1);
    const b = posts.at(-1).body; assert.strictEqual(posts.at(-1).url, FLOW); assert.strictEqual(b.event, 'test'); assert.strictEqual(b.patient, ''); assert.deepStrictEqual([...b.chats], [CHAT_A]);
    assert(texts(b.card)[0].startsWith('<at>Alex Able</at> Test from ESK Tracker'));
    assert.strictEqual((await C.readNotifyLogSince('2000')).length, before);
    assert(/No group chat set/.test((await C.sendNotifyTest('Dr Z')).reason));
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
