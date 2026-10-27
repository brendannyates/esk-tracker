// notify.js  (ESK Tracker add-in)
// onVitalsSaved(data)           call when vitals are saved: picks the stage from the visit, applies the rules, sends when warranted
// sendToProviderChatButton()    the "Send to Provider Chat" button: always sends, still marks an emergency when the numbers say so
// startProviderAlerts()         provider view: flags the bay map and plays a sound on new log rows for that provider
// One Power Automate flow for every office ("ESK Tracker to Teams", built from blank, not the Teams template): its trigger URL is
// tblNotifyConfig › WebhookUrl. The add-in decides whether to send, builds the Adaptive Card (title, Stage, flagged vitals, Sent By,
// @mention of the provider) and names the group chat(s); the flow only forwards the card to each chat as the Flow bot.
// Tables are found by name anywhere in the workbook, so they can live on the Settings sheet.
var NOTIFY_TABLES = { config: 'tblNotifyConfig', routes: 'tblNotifyRoutes', templates: 'tblNotifyTemplates', log: 'tblNotifyLog' };
var NOTIFY_VITALS = [["bp",""],["hr","HR"],["spo2","SpO2"],["temp","Temp"]];   // [field, label] for the log row and toasts. Keys bp and spo2 feed the rules.
var NOTIFY_THRESHOLDS = {"emergency":{"sys":180,"dia":120},"alert":{"sys":140,"dia":90},"spo2Min":92};   // hardcoded on purpose: change here, not in the workbook
var NFY_STAGE_LABEL = {"initial":"Initial Vitals","recheck":"Recheck Vitals","40-min":"40-min Vitals","discharge":"Discharge Vitals"};
var NOTIFY_POST_DOSE_MIN = 60;   // minutes after administration that still count as the 40-min check

function nfyTruthy(v){ return v === true || /^(true|yes|y|1)$/i.test(String(v == null ? '' : v).trim()); }
function nfyNorm(s){ return String(s == null ? '' : s).trim().toLowerCase(); }
// A Power Automate trigger URL (new powerplatform.com form or the older logic.azure.com one), else ''.
function nfyFlowUrl(v){
  var s = String(v == null ? '' : v).trim();
  return /^https:\/\/[^\/\s?#]+\.(?:powerplatform\.com|logic\.azure\.com)(?::\d+)?\/\S+$/i.test(s) ? s : '';
}
function nfyParseBp(v){ var m = String(v == null ? '' : v).match(/(\d{2,3})\s*\/\s*(\d{2,3})/); return m ? { sys: +m[1], dia: +m[2] } : null; }
function nfyParseSpo2(v){ var m = String(v == null ? '' : v).match(/(\d{2,3})/); return m ? +m[1] : null; }
// Prints only the fields that were taken: { bp:'120/80', hr:'72' } -> "120/80, HR 72"   (log row, toasts, a {vitals} token)
function nfyVitals(v){
  if (v == null) return '';
  if (typeof v !== 'object') return String(v).trim();
  return NOTIFY_VITALS.map(function(f){ var val = String(v[f[0]] == null ? '' : v[f[0]]).trim(); return val ? (f[1] ? f[1] + ' ' + val : val) : ''; })
    .filter(Boolean).join(', ');
}

// ---- The card (nfy-card logic from the build sheet). Reads as:
//   <at>Provider Name</at> GRECAR is ready in Bay 2       title: the tblNotifyTemplates row, @mention when MentionProvider is on
//   Stage: Initial Vitals
//   BP: 142/87 **HIGH** (over 140/90)                    flag on the vital that tripped: HIGH over 140/90, *EMERGENCY* at or above 180/120, LOW under 92% SpO2
//                                                        (an emergency card sits on Teams' red background; the rest stay plain)
//   HR: 69, SpO2: 99%
//   Sent By: John C.
// Lines for vitals that were not taken are left out. There is no Provider row; the card lands in the provider's own chat.

// Cards render markdown, not HTML. Neutralise the characters that would format free text.
function nfyMd(s){ return String(s == null ? '' : s).replace(/[\\*_]/g, '\\$&'); }
// Title template -> plain text; tokens are markdown-escaped. Any HTML left in an old template row is dropped.
function nfyRender(tpl, d){
  return String(tpl).replace(/<[^>]*>/g, '').replace(/\{(\w+)\}/g, function(_, k){ return nfyMd(d[k] == null ? '' : d[k]); })
    .replace(/\s{2,}/g, ' ').replace(/\s+([,.)])/g, '$1').replace(/^[\s:,-]+/, '').trim();
}
function nfyBpLine(raw){
  var T = NOTIFY_THRESHOLDS, bp = nfyParseBp(raw), txt = String(raw == null ? '' : raw).trim();
  if (!txt) return '';
  if (!bp) return 'BP: ' + nfyMd(txt);
  var shown = 'BP: ' + bp.sys + '/' + bp.dia;
  if (bp.sys >= T.emergency.sys || bp.dia >= T.emergency.dia) return shown + ' \\***EMERGENCY**\\* (at or above ' + T.emergency.sys + '/' + T.emergency.dia + ')';   // reads *EMERGENCY*, asterisks shown
  if (bp.sys > T.alert.sys || bp.dia > T.alert.dia) return shown + ' **HIGH** (over ' + T.alert.sys + '/' + T.alert.dia + ')';
  return shown;
}
function nfyHrSpLine(v){
  var T = NOTIFY_THRESHOLDS, parts = [];
  if (String(v.hr == null ? '' : v.hr).trim()) parts.push('HR: ' + nfyMd(v.hr));
  var sp = nfyParseSpo2(v.spo2);
  if (sp != null) parts.push('SpO2: ' + sp + '%' + (sp < T.spo2Min ? ' **LOW** (under ' + T.spo2Min + '%)' : ''));
  else if (String(v.spo2 == null ? '' : v.spo2).trim()) parts.push('SpO2: ' + nfyMd(v.spo2));
  if (String(v.temp == null ? '' : v.temp).trim()) parts.push('Temp: ' + nfyMd(v.temp));
  return parts.join(', ');
}
// d: { title, stage, vitals: { bp, hr, spo2, temp }, sentBy }  ->  markdown lines, title first
function nfyLines(d){
  var v = d.vitals || {}, out = [d.title];
  if (d.stage) out.push('Stage: ' + (NFY_STAGE_LABEL[d.stage] || nfyMd(d.stage)));
  var bp = nfyBpLine(v.bp); if (bp) out.push(bp);
  var hs = nfyHrSpLine(v);  if (hs) out.push(hs);
  if (d.sentBy) out.push('Sent By: ' + nfyMd(d.sentBy));
  return out;
}
// mention: { name, id } or null. id is the provider's email (UPN); if Teams shows the name as plain text, use the Entra object ID instead.
// The <at> text in the title must match the entity text exactly, which is why both come from the same string.
// emergency: true puts the whole card on Teams' red "attention" background (a theme color, so it reads in light and dark mode);
// every other card stays plain.
function nfyCard(lines, mention, emergency){
  var body = lines.map(function(t, i){
    var b = { type: 'TextBlock', text: t, wrap: true, spacing: i === 0 ? 'Default' : 'Medium' };   // room between the title and each line
    if (i === 0) b.weight = 'Bolder';
    if (i === lines.length - 1 && /^Sent By: /.test(t)) b.isSubtle = true;
    return b;
  });
  var card = { $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4', body: body };
  if (mention && mention.id) {
    var at = '<at>' + mention.name + '</at>';
    body[0].text = at + ' ' + body[0].text;
    card.msteams = { entities: [{ type: 'mention', text: at, mentioned: { id: mention.id, name: mention.name } }] };
  }
  if (emergency) card.body = [{ type: 'Container', style: 'attention', bleed: true, items: body }];
  return card;
}
// HTML twin of the same lines, for a fallback flow that posts a plain message instead of the card.
function nfyHtml(lines){
  return lines.map(function(t){
    // escaped characters (\* \_ \\) are set aside first so they never count as bold markers
    return t.replace(/[&<>]/g, function(c){ return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; })
            .replace(/\\([\\*_])/g, function(_, c){ return '\u0000' + c.charCodeAt(0) + '\u0000'; })
            .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
            .replace(/\u0000(\d+)\u0000/g, function(_, n){ return String.fromCharCode(+n); });
  }).join('<br>');
}
// Markdown line -> plain text, for toasts.
function nfyPlain(t){ return nfyHtml([String(t == null ? '' : t)]).replace(/<\/?b>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'); }

// Real JSON with CORS: the trigger answers the preflight (Allow-Origin *, content-type allowed), so the flow reads the body as-is
// and the add-in sees its answer: 202 = accepted. Returns { ok, status } or { ok: false, status, reason }.
// If the browser blocks the request anyway, it falls back to a text/plain no-cors post, which can't be confirmed ({ ok, unconfirmed }).
var NFY_HTTP_HINT = { 401: 'the URL is incomplete or out of date (paste it again from the flow’s trigger), or the flow isn’t set to Anyone', 403: 'the flow isn’t set to Anyone, or the tenant blocks it', 404: 'the URL is wrong or the flow was deleted', 409: 'the flow is turned off', 429: 'too many requests, try again in a minute' };
async function nfyPost(url, payload){
  var r;
  try { r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) }); }
  catch (e) {
    await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(payload) });
    return { ok: true, unconfirmed: true };
  }
  if (r.ok) return { ok: true, status: r.status };
  var t = ''; try { t = await r.text(); } catch (e) {}
  var code = (t.match(/"code"\s*:\s*"([^"]+)"/) || [])[1], hint = NFY_HTTP_HINT[r.status] || (/disabled/i.test(t) ? 'the flow is turned off' : '');
  return { ok: false, status: r.status, reason: 'The Teams flow answered ' + r.status + (code ? ' (' + code + ')' : '') + (hint ? ': ' + hint : '') };
}
function nfyId(){ return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

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
    var toggles = {   // office stage toggles, applied here; emergencies ignore them
      stages: { initial: flag('NotifyInitial', true), recheck: flag('NotifyRecheck', true), '40-min': flag('Notify40Min', true), discharge: flag('NotifyDischarge', true) }
    };
    // Provider | ChatLink | MentionName | MentionId | Enabled, found by header (older sheets: FlowUrl / Email)
    var col = {}; rts.values[0].forEach(function(h, i){ col[nfyNorm(h)] = i; });
    function cell(r, name){ var i = col[name]; return i == null ? '' : String(r[i] == null ? '' : r[i]).trim(); }
    var routes = {};    rts.values.slice(1).forEach(function(r){
      var prov = cell(r, 'provider'); if (!prov) return;
      var chat = nfyChatId(cell(r, 'chatlink') || cell(r, 'flowurl'));
      routes[nfyNorm(prov)] = { chats: ((col.enabled == null || nfyTruthy(cell(r, 'enabled'))) && chat) ? [chat] : [], enabled: col.enabled == null || nfyTruthy(cell(r, 'enabled')),
                                mentionName: cell(r, 'mentionname'), mentionId: cell(r, 'mentionid') || cell(r, 'email') };
    });
    var templates = {}; tpl.values.forEach(function(r){ if (nfyTruthy(r[2])) templates[String(r[0]).trim()] = String(r[1]); });
    return { config: config, routes: routes, templates: templates, toggles: toggles, webhook: nfyFlowUrl(config.WebhookUrl), mention: flag('MentionProvider', true) };
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

// The "Send to Provider Chat" button. Always sends; the card still flags EMERGENCY when the numbers say so.
async function sendToProviderChatButton(data){
  var stage = data.stage || nfyStage(data.visit);
  var d = nfyEvaluate(stage, data.vitals);
  var r = await sendToProviderChat(Object.assign({}, data, { stage: stage, level: d.level, reasons: d.reasons.length ? d.reasons : ['sent by ' + (data.sentBy || 'staff')] }));
  r.sent = !!r.ok; r.stage = stage; r.level = d.level;
  return r;
}

// Who to @mention for a route, or null (MentionProvider off, or no MentionId on the provider's row).
function nfyMention(s, route, provider){
  return (s.mention && route && route.mentionId) ? { name: route.mentionName || provider, id: route.mentionId } : null;
}
// Why a provider can't be posted to, or '' when they can.
function nfyRouteProblem(s, provider){
  if (!s.webhook) return 'The Teams flow URL isn’t set (Settings › Teams notifications)';
  var route = s.routes[nfyNorm(provider)];
  if (!route) return 'No group chat set for ' + provider + ' (Settings › Teams notifications)';
  if (!route.enabled) return provider + ' is turned off in Settings › Teams notifications';
  if (!route.chats.length) return 'No group chat link for ' + provider + ' (paste the chat’s Teams web address in Settings › Teams notifications)';
  return '';
}

// Core send: log row, then the one flow, which posts the card to the provider's group chat.
async function sendToProviderChat(data){
  var s = await loadNotifySettings();
  if (!nfyTruthy(s.config.Enabled)) return { ok: false, reason: 'Notifications are off (tblNotifyConfig > Enabled)' };
  var problem = nfyRouteProblem(s, data.provider); if (problem) return { ok: false, reason: problem };
  var route = s.routes[nfyNorm(data.provider)];
  var level = data.level === 'emergency' ? 'emergency' : 'standard';
  var tpl = s.templates[level] || s.templates.standard;
  if (!tpl) return { ok: false, reason: 'No "' + level + '" row in tblNotifyTemplates' };
  var vitalsText = nfyVitals(data.vitals);
  if (!vitalsText) return { ok: false, reason: 'No vitals entered' };

  var now = new Date();
  var title = nfyRender(tpl, Object.assign({ time: now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) }, data, { vitals: vitalsText, reasons: (data.reasons || []).join('; ') }));
  var lines = nfyLines({ title: title, stage: data.stage, vitals: typeof data.vitals === 'object' ? data.vitals : {}, sentBy: data.sentBy });
  var mention = nfyMention(s, route, data.provider);
  var payload = { event: 'vitals', chats: route.chats, card: nfyCard(lines, mention, level === 'emergency'), message: nfyHtml(lines), mentionId: mention ? mention.id : '',
                  level: level, stage: data.stage || '', patient: data.patient || '', bay: String(data.bay || ''), clinic: data.clinic || '',
                  sentBy: data.sentBy || '', sentAt: now.toISOString(), source: 'ESK Tracker' };

  // 1) Log row: drives the in-app flag + sound for provider view, and is the audit trail.
  try { await appendNotifyLog({ provider: data.provider || '', patient: payload.patient, bay: payload.bay, stage: payload.stage, vitals: vitalsText, level: level, sentBy: payload.sentBy, sentAt: payload.sentAt }); }
  catch (e) { console.warn('notify log', e); }

  // 2) Teams, through the flow.
  try {
    var a = await nfyPost(s.webhook, payload);
    if (!a.ok) return { ok: false, reason: a.reason };
    return { ok: true, message: nfyPlain(title), unconfirmed: !!a.unconfirmed };
  } catch (e) {
    return { ok: false, reason: 'Network error: ' + e.message };
  }
}

// Settings › Teams notifications › "Send test": one test card to a provider's chat (with their @mention when that's on), no log row, no patient.
// { ok, status } means the flow accepted it; whether the card then posted shows in the chat and the flow's run history.
async function sendNotifyTest(provider, sentBy){
  var s = await loadNotifySettings();
  var problem = nfyRouteProblem(s, provider); if (problem) return { ok: false, reason: problem };
  var route = s.routes[nfyNorm(provider)], now = new Date();
  var title = 'Test from ESK Tracker - ' + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  var lines = nfyLines({ title: title, stage: '', vitals: { bp: '142/87', hr: '69', spo2: '99%' }, sentBy: sentBy });
  var mention = nfyMention(s, route, provider);
  var payload = { event: 'test', chats: route.chats, card: nfyCard(lines, mention), message: nfyHtml(lines), mentionId: mention ? mention.id : '',
                  level: 'standard', stage: 'test', patient: '', bay: '', clinic: '', sentBy: sentBy || '', sentAt: now.toISOString(), source: 'ESK Tracker' };
  try { var r = await nfyPost(s.webhook, payload); return r.ok ? { ok: true, message: title, status: r.status, unconfirmed: !!r.unconfirmed, mentioned: !!mention } : { ok: false, reason: r.reason }; }
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
//     -> "Sent to provider chat: SMIJAN is ready in Bay 2"
//     -> "40-min vitals within limits, nothing sent"
//
// "Send to Provider Chat" button, same data:
//   var r = await sendToProviderChatButton(data);
//
// Provider view toggle:
//   startProviderAlerts(settings.myProvider, function(a){ flagBay(a.bay, a.level, a.patient + ' ' + a.stage + ' vitals ' + a.vitals); });
//   ...and stopProviderAlerts() when the user switches back to the tech view.
