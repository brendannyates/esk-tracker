// notify-settings.js  (ESK Tracker add-in)
// Creates the Settings sheet and the four notification tables if they are missing. Safe to call every time;
// it never overwrites a table that already exists, and stops rather than write over cells that already hold something.
// notify.js finds the tables by name, so the sheet can be renamed later.
//
// Sheet: Settings            (tables sit side by side so tblNotifyLog can grow downward without hitting anything)
//
//   tblNotifyConfig     A1   Key | Value
//                            WebhookUrl        ''        HTTP URL from the flow's trigger card, pasted once per office
//                            Enabled           TRUE      master switch
//                            SendGroupChat     TRUE      office toggle, applied by the flow
//                            UrgentEmergencies TRUE      office toggle, applied by the flow (Teams priority notification)
//                            EscalateEmergency FALSE     office toggle, applied by the flow
//                            EscalationChat    ''        group chat URL or ID for emergencies (charge RN, clinic lead)
//                            NotifyInitial     TRUE      stage toggles, applied by the add-in;
//                            NotifyRecheck     TRUE        emergencies ignore them
//                            Notify40Min       TRUE
//                            NotifyDischarge   TRUE
//                            InAppAlerts       TRUE      provider view flag
//                            AlertSound        TRUE      provider view sound
//
//   tblNotifyRoutes     D1   Provider | ChatLink | Email | Enabled   one row per provider; ChatLink = Teams URL or 19:…@thread.v2
//   tblNotifyTemplates  I1   Event | Template | Enabled             rows: standard, emergency
//   tblNotifyLog        M1   Id | Provider | Patient | Bay | Stage | Vitals | Level | SentBy | SentAt | Seen   (written by notify.js)
//
// Thresholds are deliberately not here; they are constants in notify.js.

var NOTIFY_SETTINGS_SHEET = 'Settings';

var NOTIFY_SETTINGS_SPEC = [
  { name: 'tblNotifyConfig', anchor: 'A1', columns: ['Key', 'Value'], rows: [
      ['WebhookUrl', ''],
      ['Enabled', 'TRUE'],
      ['SendGroupChat', 'TRUE'],
      ['UrgentEmergencies', 'TRUE'],
      ['EscalateEmergency', 'FALSE'],
      ['EscalationChat', ''],
      ['NotifyInitial', 'TRUE'],
      ['NotifyRecheck', 'TRUE'],
      ['Notify40Min', 'TRUE'],
      ['NotifyDischarge', 'TRUE'],
      ['InAppAlerts', 'TRUE'],
      ['AlertSound', 'TRUE']
  ] },
  { name: 'tblNotifyRoutes', anchor: 'D1', columns: ['Provider', 'ChatLink', 'Email', 'Enabled'], rows: [
      ['', '', '', 'TRUE']                               // blank starter row; notify.js skips rows with no Provider
  ] },
  { name: 'tblNotifyTemplates', anchor: 'I1', columns: ['Event', 'Template', 'Enabled'], rows: [
      ['standard',  '{patient} {stage} vitals {vitals} - Bay {bay}', 'TRUE'],
      ['emergency', '<b>EMERGENCY</b> {patient} {stage} vitals {vitals} - Bay {bay} ({provider})', 'TRUE']
  ] },
  { name: 'tblNotifyLog', anchor: 'M1', columns: ['Id', 'Provider', 'Patient', 'Bay', 'Stage', 'Vitals', 'Level', 'SentBy', 'SentAt', 'Seen'], rows: [
      ['', '', '', '', '', '', '', '', '', '']           // blank starter row; notify.js skips rows with no Id
  ] }
];

function nfyColLetter(n){ var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function nfyRangeFrom(anchor, rowCount, colCount){
  var m = anchor.match(/^([A-Z]+)(\d+)$/), col = m[1], row = +m[2], colIdx = 0;
  for (var i = 0; i < col.length; i++) colIdx = colIdx * 26 + (col.charCodeAt(i) - 64);
  return col + row + ':' + nfyColLetter(colIdx + colCount - 1) + (row + rowCount - 1);
}

// Creates whatever is missing. Returns { sheet, created: [...], existing: [...], blocked: [...] }.
// blocked = tables not created because their cells on an existing Settings sheet already hold something (nothing is overwritten).
// options.hide = true hides the sheet afterwards (the add-in can still read hidden sheets).
async function ensureNotifySettings(options){
  options = options || {};
  return Excel.run(async function(ctx){
    var sheet = ctx.workbook.worksheets.getItemOrNullObject(NOTIFY_SETTINGS_SHEET);
    var tables = ctx.workbook.tables; tables.load('items/name');
    await ctx.sync();
    var isNew = sheet.isNullObject;
    if (isNew) sheet = ctx.workbook.worksheets.add(NOTIFY_SETTINGS_SHEET);

    var have = {}; tables.items.forEach(function(t){ have[t.name] = true; });
    var created = [], existing = [], blocked = [], todo = [];
    NOTIFY_SETTINGS_SPEC.forEach(function(spec){
      if (have[spec.name]) { existing.push(spec.name); return; }
      var values  = [spec.columns].concat(spec.rows);
      var address = nfyRangeFrom(spec.anchor, values.length, spec.columns.length);
      var target = sheet.getRange(address); if (!isNew) target.load('values');
      todo.push({ spec: spec, values: values, address: address, target: target });
    });
    if (!isNew && todo.length) await ctx.sync();
    todo.forEach(function(t){
      if (!isNew && t.target.values.some(function(r){ return r.some(function(v){ return v !== '' && v != null; }); })) { blocked.push(t.spec.name + ' (' + t.address + ')'); return; }
      t.target.values = t.values;
      var table = sheet.tables.add(t.address, true);
      table.name = t.spec.name;
      created.push(t.spec.name);
    });
    if (created.length) sheet.getUsedRange().format.autofitColumns();
    if (options.hide) sheet.visibility = Excel.SheetVisibility.hidden;
    await ctx.sync();
    return { sheet: NOTIFY_SETTINGS_SHEET, created: created, existing: existing, blocked: blocked };
  });
}

// Which tables exist. Use it to decide whether to show a "Set up Teams notifications" button.
async function notifySettingsStatus(){
  return Excel.run(async function(ctx){
    var tables = ctx.workbook.tables; tables.load('items/name');
    await ctx.sync();
    var have = {}; tables.items.forEach(function(t){ have[t.name] = true; });
    return NOTIFY_SETTINGS_SPEC.map(function(s){ return { table: s.name, present: !!have[s.name] }; });
  });
}

// Everything the add-in's Settings panel shows: { config: { Key: value }, routes: [{ provider, chatLink, email, enabled }] }.
// Columns are found by header, so a column moved by hand still reads right. Missing tables give empty results.
async function readNotifySettings(){
  return Excel.run(async function(ctx){
    var c = ctx.workbook.tables.getItemOrNullObject('tblNotifyConfig'), r = ctx.workbook.tables.getItemOrNullObject('tblNotifyRoutes');
    await ctx.sync();
    var cv = c.isNullObject ? null : c.getRange().load('values'), rv = r.isNullObject ? null : r.getRange().load('values');
    await ctx.sync();
    var config = {}; if (cv) cv.values.slice(1).forEach(function(row){ var k = String(row[0]).trim(); if (k) config[k] = row[1]; });
    var routes = [];
    if (rv){ var h = rv.values[0].map(function(x){ return String(x).trim().toLowerCase(); }), at = function(n){ return h.indexOf(n); };
      rv.values.slice(1).forEach(function(row){ var p = String(row[at('provider')] || '').trim(); if (!p) return;
        routes.push({ provider: p, chatLink: String(row[at('chatlink')] ?? '').trim(), email: at('email') < 0 ? '' : String(row[at('email')] ?? '').trim(), enabled: String(row[at('enabled')]).toUpperCase() !== 'FALSE' }); }); }
    return { config: config, routes: routes, hasConfig: !!cv, hasRoutes: !!rv };
  });
}

// Write one config value from the add-in's Settings panel (so staff never open the sheet). Adds the row if it is missing.
async function setNotifyConfig(key, value){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem('tblNotifyConfig');
    var body = t.getDataBodyRange().load('values');
    await ctx.sync();
    var idx = body.values.findIndex(function(r){ return String(r[0]).trim() === key; });
    if (idx < 0) t.rows.add(null, [[key, value]]);
    else body.getCell(idx, 1).values = [[value]];
    await ctx.sync();
  });
}

// Add or update a provider's group chat and email from the add-in's Settings panel (matched by provider name, any case).
async function upsertNotifyRoute(provider, chatLink, email, enabled){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem('tblNotifyRoutes');
    var body = t.getDataBodyRange().load('values');
    await ctx.sync();
    var key = String(provider).trim().toLowerCase();
    var idx = body.values.findIndex(function(r){ return String(r[0]).trim().toLowerCase() === key; });
    var row = [String(provider).trim(), String(chatLink || '').trim(), String(email || '').trim(), enabled === false ? 'FALSE' : 'TRUE'];
    if (idx < 0) {
      var blank = body.values.findIndex(function(r){ return !String(r[0]).trim(); });   // reuse the starter row if it is still empty
      if (blank < 0) t.rows.add(null, [row]);
      else body.getCell(blank, 0).getResizedRange(0, 3).values = [row];
    } else {
      body.getCell(idx, 0).getResizedRange(0, 3).values = [row];
    }
    await ctx.sync();
  });
}

// Remove a provider's row (Settings panel "Remove"). The last row is blanked instead, since a table can't have no rows.
async function removeNotifyRoute(provider){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem('tblNotifyRoutes');
    var body = t.getDataBodyRange().load('values,rowCount');
    await ctx.sync();
    var key = String(provider).trim().toLowerCase();
    var idx = body.values.findIndex(function(r){ return String(r[0]).trim().toLowerCase() === key; });
    if (idx < 0) return false;
    if (body.rowCount === 1) body.getCell(0, 0).getResizedRange(0, 3).values = [['', '', '', 'TRUE']];
    else t.rows.getItemAt(idx).delete();
    await ctx.sync();
    return true;
  });
}

// ---- Wiring (taskpane.html, Settings › Teams notifications)
//   "Set up Teams notifications" → ensureNotifySettings();  webhook field → setNotifyConfig('WebhookUrl', url)
//   provider rows → upsertNotifyRoute(provider, chatLink, email, enabled) / removeNotifyRoute(provider)
