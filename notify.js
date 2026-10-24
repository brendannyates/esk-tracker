// notify.js  (ESK Tracker add-in)
// onVitalsSaved(data)           call when vitals are saved: picks the stage from the visit, applies the rules, sends when warranted
// sendToProviderChatButton()    the "Send to Provider Chat" button: always sends, still marks an emergency when the numbers say so
// startProviderAlerts()         provider view: flags the bay map and plays a sound on new log rows for that provider
// Each provider has their own Teams workflow ("Send webhook alerts to a chat"): its URL sits in tblNotifyRoutes, and the post is an
// Adaptive Card that workflow can show as it is. The same post also carries the plain fields (message, level, vitals...) for a custom flow.
// Tables are found by name anywhere in the workbook, so they can live on the Settings sheet.
var NOTIFY_TABLES = { config: 'tblNotifyConfig', routes: 'tblNotifyRoutes', templates: 'tblNotifyTemplates', log: 'tblNotifyLog' };
var NOTIFY_VITALS = [["bp",""],["hr","HR"],["spo2","SpO2"],["temp","Temp"]];   // [field, label] in print order; blank label prints the value alone. Keys bp and spo2 feed the rules.
var NOTIFY_THRESHOLDS = {"emergency":{"sys":180,"dia":120},"alert":{"sys":140,"dia":90},"spo2Min":92};   // hardcoded on purpose: change here, not in the workbook
var NOTIFY_POST_DOSE_MIN = 60;   // minutes after administration that still count as the 40-min check

function nfyTruthy(v){ return v === true || /^(true|yes|y|1)$/i.test(String(v == null ? '' : v).trim()); }
function nfyNorm(s){ return String(s == null ? '' : s).trim().toLowerCase(); }
// A Power Automate trigger URL (new powerplatform.com form or the older logic.azure.com one), else ''.
function nfyFlowUrl(v){
  var s = String(v == null ? '' : v).trim();
  return /^https:\/\/[^\/\s?#]+\.(?:powerplatform\.com|logic\.azure\.com)(?::\d+)?\/\S+$/i.test(s) ? s : '';
}
// Plain text for the card and the toast: <b>x</b> becomes **x** (Adaptive Card markdown), other tags go.
function nfyText(tpl, d){
  return String(tpl).replace(/\{(\w+)\}/g, function(_, k){ return d[k] == null ? '' : String(d[k]); })
    .replace(/<\/?b>/gi, '**').replace(/<[^>]*>/g, '').replace(/\s{2,}/g, ' ').replace(/\s+([,.)])/g, '$1').trim();
}
// What the Teams workflow template expects: { type: 'message', attachments: [ Adaptive Card ] }. Emergencies get a red heading.
function nfyCard(text, facts, emergency){
  var body = [];
  if (emergency) body.push({ type: 'TextBlock', text: 'EMERGENCY', weight: 'Bolder', size: 'Large', color: 'Attention' });
  body.push({ type: 'TextBlock', text: text, wrap: true });
  var f = (facts || []).filter(function(x){ return x[1]; }).map(function(x){ return { title: x[0], value: String(x[1]) }; });
  if (f.length) body.push({ type: 'FactSet', facts: f });
  return { type: 'message', attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', contentUrl: null,
    content: { $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4', body: body } }] };
}
// Real JSON with CORS: the trigger answers the preflight (Allow-Origin *, content-type allowed), so the workflow reads the body as-is
// and the add-in sees its answer: 202 = accepted. Returns { ok, status } or { ok: false, status, reason }.
// If the browser blocks the request anyway, it falls back to a text/plain no-cors post, which can't be confirmed ({ ok, unconfirmed }).
var NFY_HTTP_HINT = { 401: 'the URL is incomplete or out of date (paste it again from the trigger card), or the workflow isn’t set to Anyone', 403: 'the workflow isn’t set to Anyone, or the tenant blocks it', 404: 'the URL is wrong or the workflow was deleted', 409: 'the workflow is turned off', 429: 'too many requests, try again in a minute' };
async function nfyPost(url, payload){
  var r;
  try { r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }); }
  catch (e) {
    await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(payload) });
    return { ok: true, unconfirmed: true };
  }
  if (r.ok) return { ok: true, status: r.status };
  var t = ''; try { t = await r.text(); } catch (e) {}
  var code = (t.match(/"code"\s*:\s*"([^"]+)"/) || [])[1], hint = NFY_HTTP_HINT[r.status] || (/disabled/i.test(t) ? 'the workflow is turned off' : '');
  return { ok: false, status: r.status, reason: 'The Teams workflow answered ' + r.status + (code ? ' (' + code + ')' : '') + (hint ? ': ' + hint : '') };
}
function nfyId(){ return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
function nfyParseBp(v){ var m = String(v == null ? '' : v).match(/(\d{2,3})\s*\/\s*(\d{2,3})/); return m ? { sys: +m[1], dia: +m[2] } : null; }
function nfyParseSpo2(v){ var m = String(v == null ? '' : v).match(/(\d{2,3})/); return m ? +m[1] : null; }
// Prints only the fields that were taken: { bp:'120/80', hr:'72' } -> "120/80, HR 72"
function nfyVitals(v){
  if (v == null) return '';
  if (typeof v !== 'object') return String(v).trim();
  return NOTIFY_VITALS.map(function(f){ var val = String(v[f[0]] == null ? '' : v[f[0]]).trim(); return val ? (f[1] ? f[1] + ' ' + val : val) : ''; })
    .filter(Boolean).join(', ');
}

// Stage from where the patient is in the visit. visit = { vitalsCount, adminTime }  (map these two from your row)
//   initial    no vitals recorded yet for this visit
//   recheck    vitals taken, not dosed yet
//   40-min     under NOTIFY_POST_DOSE_MIN minutes after administration
//   discharge  after that
function nfyStage(visit){
  if (!visit || !(visit.vitalsCount > 0)) return 'initial';
  if (!visit.adminTime) return 'recheck';
  var mins = (Date.now() - new Date(visit.adminTime).getTime()) / 60000;
  return mins < NOTIFY_POST_DOSE_MIN ? '40-min' : 'discharge';
}

// The rules. Initial vitals always go. Later vitals go only on a threshold:
//   emergency  BP at or above 180/120 (either number)
//   standard   BP over 140/90 (either number), or SpO2 under 92%
function nfyEvaluate(stage, vitals){
  var T = NOTIFY_THRESHOLDS, v = vitals || {};
  var bp = nfyParseBp(v.bp), sp = nfyParseSpo2(v.spo2), reasons = [], level = 'standard', note = '';
  if (bp && (bp.sys >= T.emergency.sys || bp.dia >= T.emergency.dia)) { level = 'emergency'; reasons.push('BP ' + bp.sys + '/' + bp.dia + ' at or above ' + T.emergency.sys + '/' + T.emergency.dia); }
  else if (bp && (bp.sys > T.alert.sys || bp.dia > T.alert.dia)) reasons.push('BP ' + bp.sys + '/' + bp.dia + ' over ' + T.alert.sys + '/' + T.alert.dia);
  if (!bp && String(v.bp == null ? '' : v.bp).trim()) note = 'BP "' + v.bp + '" is not in 120/80 form';
  if (sp != null && sp < T.spo2Min) reasons.push('SpO2 ' + sp + '% under ' + T.spo2Min + '%');
  var send = stage === 'initial' || reasons.length > 0;
  if (stage === 'initial' && !reasons.length) reasons.push('first vitals of the visit');
  return { send: send, level: level, reasons: reasons, note: note };
}

async function loadNotifySettings(){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables;
    var cfg = t.getItem(NOTIFY_TABLES.config).getDataBodyRange().load('values');
    var rts = t.getItem(NOTIFY_TABLES.routes).getRange().load('values');   // with the header row: columns are found by name
    var tpl = t.getItem(NOTIFY_TABLES.templates).getDataBodyRange().load('values');
    await ctx.sync();
    var config = {};    cfg.values.forEach(function(r){ config[String(r[0]).trim()] = r[1]; });
    function flag(key, dflt){ var v = config[key]; return (v == null || String(v).trim() === '') ? dflt : nfyTruthy(v); }
    // Office toggles. The add-in applies the stage ones; the channel ones ride in the payload and the flow applies them.
    var toggles = {
      group: flag('SendGroupChat', true), escalate: flag('EscalateEmergency', false), urgent: flag('UrgentEmergencies', true),
      stages: { initial: flag('NotifyInitial', true), recheck: flag('NotifyRecheck', true), '40-min': flag('Notify40Min', true), discharge: flag('NotifyDischarge', true) }
    };
    var esc = nfyFlowUrl(config.EscalationChat);   // EscalationChat now holds the escalation chat's workflow URL
    // Provider | FlowUrl | Email | Enabled, found by header (older sheets call the second column ChatLink)
    var hd = rts.values[0].map(nfyNorm), col = function(n, d){ var i = hd.indexOf(n); return i < 0 ? d : i; };
    var cP = col('provider', 0), cC = col('flowurl', col('chatlink', 1)), cE = col('email', -1), cOn = col('enabled', hd.length - 1);
    var routes = {};    rts.values.slice(1).forEach(function(r){
      if (!nfyTruthy(r[cOn]) || !String(r[cP]).trim()) return;
      routes[nfyNorm(r[cP])] = { flow: nfyFlowUrl(r[cC]), email: cE < 0 ? '' : String(r[cE] == null ? '' : r[cE]).trim() };
    });
    var templates = {}; tpl.values.forEach(function(r){ if (nfyTruthy(r[2])) templates[String(r[0]).trim()] = String(r[1]); });
    return { config: config, routes: routes, templates: templates, toggles: toggles, escalation: esc };
  });
}

// data: { provider, patient (LASFIR), bay, clinic, sentBy,
//         vitals: { bp:'120/80', hr:'72', spo2:'98%', temp:'98.2' },   whatever was taken
//         visit:  { vitalsCount: 0, adminTime: null },                  where the patient is in the visit
//         stage:  optional override of the computed stage }

// Call when vitals are saved. Returns { sent, stage, level, message } or { sent:false, stage, reason }.
async function onVitalsSaved(data){
  var stage = data.stage || nfyStage(data.visit);
  var d = nfyEvaluate(stage, data.vitals);
  if (!d.send) return { ok: true, sent: false, stage: stage, reason: d.note ? d.note + '; nothing sent' : stage + ' vitals within limits, nothing sent' };
  if (d.level !== 'emergency') {                       // office stage toggles gate standard alerts; emergencies always go
    var s0 = await loadNotifySettings();
    if (s0.toggles.stages[stage] === false) return { ok: true, sent: false, stage: stage, reason: stage + ' alerts are off for this office (tblNotifyConfig), nothing sent' };
  }
  var r = await sendToProviderChat(Object.assign({}, data, { stage: stage, level: d.level, reasons: d.reasons }));
  r.sent = !!r.ok; r.stage = stage; r.level = d.level;
  return r;
}

// The "Send to Provider Chat" button. Always sends; the message still says EMERGENCY when the numbers say so.
async function sendToProviderChatButton(data){
  var stage = data.stage || nfyStage(data.visit);
  var d = nfyEvaluate(stage, data.vitals);
  var r = await sendToProviderChat(Object.assign({}, data, { stage: stage, level: d.level, reasons: d.reasons.length ? d.reasons : ['sent by ' + (data.sentBy || 'staff')] }));
  r.sent = !!r.ok; r.stage = stage; r.level = d.level;
  return r;
}

// Core send: log row, then the provider's workflow (and the escalation workflow for an emergency when that's on).
async function sendToProviderChat(data){
  var s = await loadNotifySettings();
  if (!nfyTruthy(s.config.Enabled)) return { ok: false, reason: 'Notifications are off (tblNotifyConfig > Enabled)' };
  var route = s.routes[nfyNorm(data.provider)];
  if (!route || !route.flow)
    return { ok: false, reason: 'No Teams workflow URL set for ' + data.provider + ' (Settings › Teams notifications)' };
  var level = data.level === 'emergency' ? 'emergency' : 'standard';
  var T = s.toggles, willEscalate = level === 'emergency' && T.escalate && !!s.escalation;
  if (!T.group && !willEscalate)
    return { ok: false, reason: 'Group chat posting is off for this office (SendGroupChat in tblNotifyConfig)' };
  var tpl = s.templates[level] || s.templates.standard;
  if (!tpl) return { ok: false, reason: 'No "' + level + '" row in tblNotifyTemplates' };
  var vitals = nfyVitals(data.vitals);
  if (!vitals) return { ok: false, reason: 'No vitals entered' };

  var now = new Date();
  var d = Object.assign({ time: now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) }, data,
                        { vitals: vitals, reasons: (data.reasons || []).join('; ') });
  var message = nfyText(tpl, d);
  var payload = Object.assign(nfyCard(message, [['Provider', d.provider], ['Stage', d.stage], ['Why', d.reasons], ['Sent by', d.sentBy]], level === 'emergency'), {
    event: 'vitals', message: message.replace(/\*\*/g, ''),
    level: level, stage: d.stage || '', vitals: vitals, reasons: data.reasons || [],
    toggles: s.toggles,
    provider: d.provider || '', patient: d.patient || '', bay: String(d.bay || ''), clinic: d.clinic || '',
    sentBy: d.sentBy || '', sentAt: now.toISOString(), source: 'ESK Tracker'
  });

  // 1) Log row: drives the in-app flag + sound for provider view, and is the audit trail.
  try { await appendNotifyLog(payload); } catch (e) { console.warn('notify log', e); }

  // 2) Teams, through the provider's workflow (and the escalation one).
  try {
    var a = T.group ? await nfyPost(route.flow, payload) : { ok: true };
    if (!a.ok) return { ok: false, reason: a.reason };
    var b = willEscalate ? await nfyPost(s.escalation, payload) : { ok: true };
    if (!b.ok) return { ok: false, reason: 'Sent to ' + data.provider + ', but the escalation workflow failed. ' + b.reason };
    return { ok: true, message: payload.message, unconfirmed: !!(a.unconfirmed || b.unconfirmed) };
  } catch (e) {
    return { ok: false, reason: 'Network error: ' + e.message };
  }
}

// Settings › Teams notifications › "Send test": one test card through a provider's workflow, no log row, no patient.
// { ok, status } means the workflow accepted it; whether the card then posted shows in the chat and the run history.
async function sendNotifyTest(provider, sentBy){
  var s = await loadNotifySettings();
  var route = s.routes[nfyNorm(provider)];
  if (!route || !route.flow) return { ok: false, reason: 'No Teams workflow URL saved for ' + provider };
  var now = new Date(), msg = 'Test from ESK Tracker' + (sentBy ? ' (' + sentBy + ')' : '') + ' - ' + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  var payload = Object.assign(nfyCard(msg), { event: 'test', message: msg, level: 'standard', stage: 'test', vitals: '', reasons: ['test'],
    toggles: s.toggles, provider: provider, patient: '', bay: '', clinic: '', sentBy: sentBy || '', sentAt: now.toISOString(), source: 'ESK Tracker' });
  try { var r = await nfyPost(route.flow, payload); return r.ok ? { ok: true, message: msg, status: r.status, unconfirmed: !!r.unconfirmed } : { ok: false, reason: r.reason }; }
  catch (e) { return { ok: false, reason: 'Network error: ' + e.message }; }
}

async function appendNotifyLog(p){        // Id | Provider | Patient | Bay | Stage | Vitals | Level | SentBy | SentAt | Seen
  return Excel.run(async function(ctx){
    ctx.workbook.tables.getItem(NOTIFY_TABLES.log).rows.add(null,
      [[ nfyId(), p.provider, p.patient, p.bay, p.stage, p.vitals, p.level, p.sentBy, p.sentAt, 'FALSE' ]]);
    await ctx.sync();
  });
}

async function readNotifyLogSince(iso){
  return Excel.run(async function(ctx){
    var r = ctx.workbook.tables.getItem(NOTIFY_TABLES.log).getDataBodyRange().load('values');
    await ctx.sync();
    return r.values
      .filter(function(v){ return v[0] && String(v[8]) > iso; })
      .map(function(v){ return { id: v[0], provider: v[1], patient: v[2], bay: String(v[3]), stage: v[4], vitals: v[5], level: v[6], sentBy: v[7], sentAt: v[8] }; });
  });
}

// ---- Provider view: call startProviderAlerts(myProvider, onAlert) when the user picks "I'm this provider",
// ---- and stopProviderAlerts() when they leave provider view. onAlert gets one object per new log row:
// ---- { id, provider, patient, bay, stage, vitals, level ('standard' | 'emergency'), sentBy, sentAt }.
var nfyAlerts = { provider: '', lastSeen: '', timer: null, handler: null, sound: true, enabled: true };
var nfyAudio = null;

async function startProviderAlerts(myProvider, onAlert){
  stopProviderAlerts();
  var s = await loadNotifySettings();
  nfyAlerts.enabled = nfyTruthy(s.config.InAppAlerts);
  nfyAlerts.sound   = nfyTruthy(s.config.AlertSound);
  if (!nfyAlerts.enabled) return;
  nfyAlerts.provider = nfyNorm(myProvider);
  nfyAlerts.lastSeen = new Date().toISOString();   // only rows sent after the view opened
  // Create the audio context inside this click so the browser lets it play later.
  try { nfyAudio = nfyAudio || new (window.AudioContext || window.webkitAudioContext)(); if (nfyAudio.state === 'suspended') nfyAudio.resume(); } catch (e) {}

  async function check(){
    try {
      var rows = (await readNotifyLogSince(nfyAlerts.lastSeen)).filter(function(r){ return nfyNorm(r.provider) === nfyAlerts.provider; });
      if (!rows.length) return;
      var worst = 'standard';
      rows.forEach(function(r){ if (r.sentAt > nfyAlerts.lastSeen) nfyAlerts.lastSeen = r.sentAt; if (r.level === 'emergency') worst = 'emergency'; });
      rows.forEach(onAlert);
      if (nfyAlerts.sound) nfyBeep(worst);
    } catch (e) { console.warn('provider alerts', e); }
  }
  // Excel fires onChanged for other people's edits while co-authoring; the 30 s poll is the backstop.
  await Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem(NOTIFY_TABLES.log);
    nfyAlerts.handler = t.onChanged.add(function(){ check(); return Promise.resolve(); });
    await ctx.sync();
  });
  nfyAlerts.timer = setInterval(check, 30000);
}

function stopProviderAlerts(){
  if (nfyAlerts.timer) { clearInterval(nfyAlerts.timer); nfyAlerts.timer = null; }
  if (nfyAlerts.handler) {
    var h = nfyAlerts.handler; nfyAlerts.handler = null;
    Excel.run(h.context, async function(ctx){ h.remove(); await ctx.sync(); }).catch(function(){});
  }
}

function nfyBeep(level){                  // two short tones; three for an emergency. No audio file needed.
  try {
    if (!nfyAudio) return;
    var times = level === 'emergency' ? [0, 0.22, 0.44] : [0, 0.22];
    times.forEach(function(t){
      var o = nfyAudio.createOscillator(), g = nfyAudio.createGain(), at = nfyAudio.currentTime + t;
      o.connect(g); g.connect(nfyAudio.destination); o.frequency.value = level === 'emergency' ? 988 : 880;
      g.gain.setValueAtTime(0.0001, at); g.gain.exponentialRampToValueAtTime(0.2, at + 0.01); g.gain.exponentialRampToValueAtTime(0.0001, at + 0.18);
      o.start(at); o.stop(at + 0.2);
    });
  } catch (e) {}
}

// ---- Wiring examples
// When vitals are saved for a bay (this is the automatic path):
//   var r = await onVitalsSaved({ provider: row.provider, patient: row.lasfir, bay: row.bay, clinic: settings.clinic, sentBy: user.initials,
//                                 vitals: { bp: ui.bp, hr: ui.hr, spo2: ui.spo2, temp: ui.temp },
//                                 visit:  { vitalsCount: row.vitalsCount, adminTime: row.adminTime } });
//   showToast(r.sent ? 'Sent to provider chat: ' + r.message : r.reason);
//     -> "Sent to provider chat: SMIJAN initial vitals 120/80 - Bay 2"
//     -> "40-min vitals within limits, nothing sent"
//
// "Send to Provider Chat" button, same data:
//   var r = await sendToProviderChatButton(data);
//
// Provider view toggle:
//   startProviderAlerts(settings.myProvider, function(a){ flagBay(a.bay, a.level, a.patient + ' ' + a.stage + ' vitals ' + a.vitals); });
//   ...and stopProviderAlerts() when the user switches back to the tech view.
