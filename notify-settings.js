// notify-settings.js  (ESK Tracker add-in)
// Creates the Settings sheet and the four notification tables if they are missing, and brings tables from the older
// one-workflow-per-provider setup up to date. Safe to call every time: it never overwrites a value someone typed,
// and stops rather than write over cells that already hold something. notify.js finds the tables by name and the
// columns by header, so the sheet can be renamed and columns moved.
//
// Sheet: Settings            (tables sit side by side so tblNotifyLog can grow downward without hitting anything)
//
//   tblNotifyConfig     A1   Key | Value
//                            Enabled           TRUE      master switch
//                            WebhookUrl        ''        HTTP URL of the "ESK Tracker to Teams" flow's trigger (one flow for every office)
//                            MentionProvider   TRUE      @mention the provider at the start of the card (needs MentionId on their row)
//                            NotifyInitial     TRUE      stage toggles, applied by the add-in;
//                            NotifyRecheck     TRUE        emergencies ignore them
//                            Notify40Min       TRUE
//                            NotifyDischarge   TRUE
//                            InAppAlerts       TRUE      provider view flag
//                            AlertSound        TRUE      provider view sound
//
//   tblNotifyRoutes     D1   Provider | ChatLink | MentionName | MentionId | Enabled
//                            ChatLink    the provider's group chat: its Teams web address, or the 19:…@thread.v2 id itself
//                            MentionName the provider's name as Teams shows it;  MentionId their email (UPN) or Entra object ID
//   tblNotifyTemplates  I1   Event | Template | Enabled             rows: standard, emergency (the card's title line, plain text)
//   tblNotifyLog        M1   Id | Provider | Patient | Bay | Stage | Vitals | Level | SentBy | SentAt | Seen   (written by notify.js)
//
// Thresholds are deliberately not here; they are constants in notify.js.

var NOTIFY_SETTINGS_SHEET = 'Settings';
var NOTIFY_ROUTE_COLUMNS = ['Provider', 'ChatLink', 'MentionName', 'MentionId', 'Enabled'];
var NOTIFY_TEMPLATE_DEFAULTS = { standard: '{patient} is ready in Bay {bay}', emergency: 'EMERGENCY: {patient} in Bay {bay}' };
// Template rows the older add-in wrote; ensureNotifySettings() swaps these (and only these) for the new wording.
var NOTIFY_TEMPLATE_OLD = { standard: ['{patient} {stage} vitals {vitals} - Bay {bay}'],
                            emergency: ['<b>EMERGENCY</b> {patient} {stage} vitals {vitals} - Bay {bay} ({provider})'] };

var NOTIFY_SETTINGS_SPEC = [
  { name: 'tblNotifyConfig', anchor: 'A1', columns: ['Key', 'Value'], rows: [
      ['Enabled', 'TRUE'],
      ['WebhookUrl', ''],
      ['MentionProvider', 'TRUE'],
      ['NotifyInitial', 'TRUE'],
      ['NotifyRecheck', 'TRUE'],
      ['Notify40Min', 'TRUE'],
      ['NotifyDischarge', 'TRUE'],
      ['InAppAlerts', 'TRUE'],
      ['AlertSound', 'TRUE']
  ] },
  { name: 'tblNotifyRoutes', anchor: 'D1', columns: NOTIFY_ROUTE_COLUMNS, rows: [
      ['', '', '', '', 'TRUE']                           // blank starter row; notify.js skips rows with no Provider
  ] },
  { name: 'tblNotifyTemplates', anchor: 'I1', columns: ['Event', 'Template', 'Enabled'], rows: [
      ['standard',  NOTIFY_TEMPLATE_DEFAULTS.standard,  'TRUE'],
      ['emergency', NOTIFY_TEMPLATE_DEFAULTS.emergency, 'TRUE']
  ] },
  { name: 'tblNotifyLog', anchor: 'M1', columns: ['Id', 'Provider', 'Patient', 'Bay', 'Stage', 'Vitals', 'Level', 'SentBy', 'SentAt', 'Seen'], rows: [
      ['', '', '', '', '', '', '', '', '', '']           // blank starter row; notify.js skips rows with no Id
  ] }
];

// The group chat id out of whatever was pasted: the chat's Teams web address (encoded or not) or the id itself. '' if none.
function nfyChatId(v){
  var s = String(v == null ? '' : v).trim();
  try { s = decodeURIComponent(s); } catch (e) {}
  var m = s.match(/19:[^\s\/?#&"'<>]+?@(?:thread\.v2|unq\.gbl\.spaces)/i);
  return m ? m[0] : '';
}
function nfyIsEmail(s){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s == null ? '' : s).trim()); }
function nfyIsGuid(s){ return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(s == null ? '' : s).trim()); }

function nfyColLetter(n){ var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function nfyRangeFrom(anchor, rowCount, colCount){
  var m = anchor.match(/^([A-Z]+)(\d+)$/), col = m[1], row = +m[2], colIdx = 0;
  for (var i = 0; i < col.length; i++) colIdx = colIdx * 26 + (col.charCodeAt(i) - 64);
  return col + row + ':' + nfyColLetter(colIdx + colCount - 1) + (row + rowCount - 1);
}
function nfyHeaderIndex(header){ var col = {}; header.forEach(function(h, i){ col[String(h).trim().toLowerCase()] = i; }); return col; }

// What an existing workbook still needs from the older setup. Pure: takes the three tables' values (header row included).
//   routes:    FlowUrl -> ChatLink, Email -> MentionId (their email is the mention id), MentionName added
//   config:    WebhookUrl and MentionProvider rows added
//   templates: the old default wording (with HTML and {vitals}) replaced by the new title lines
function nfyUpgradePlan(routes, config, templates){
  var plan = { rename: [], addColumns: [], addConfig: [], templates: [] };
  if (routes) {
    var h = routes[0].map(function(x){ return String(x).trim().toLowerCase(); });
    if (h.indexOf('chatlink') < 0 && h.indexOf('flowurl') >= 0) plan.rename.push({ index: h.indexOf('flowurl'), to: 'ChatLink' });
    if (h.indexOf('mentionid') < 0 && h.indexOf('email') >= 0) plan.rename.push({ index: h.indexOf('email'), to: 'MentionId' });
    var have = h.slice(); plan.rename.forEach(function(r){ have[r.index] = r.to.toLowerCase(); });
    ['ChatLink', 'MentionName', 'MentionId', 'Enabled'].forEach(function(c){ if (have.indexOf(c.toLowerCase()) < 0) plan.addColumns.push(c); });
  }
  if (config) {
    var keys = config.slice(1).map(function(r){ return String(r[0]).trim(); });
    [['WebhookUrl', ''], ['MentionProvider', 'TRUE']].forEach(function(kv){ if (keys.indexOf(kv[0]) < 0) plan.addConfig.push(kv); });
  }
  if (templates) templates.slice(1).forEach(function(r, i){
    var ev = String(r[0]).trim(), old = NOTIFY_TEMPLATE_OLD[ev] || [];
    if (old.indexOf(String(r[1]).trim()) >= 0) plan.templates.push({ row: i, value: NOTIFY_TEMPLATE_DEFAULTS[ev] });
  });
  plan.needed = !!(plan.rename.length || plan.addColumns.length || plan.addConfig.length || plan.templates.length);
  return plan;
}

// Creates whatever is missing and upgrades what's there. Returns { sheet, created, existing, blocked, upgraded }.
// blocked = tables not created (or columns not added) because cells already hold something (nothing is overwritten).
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
    var created = [], existing = [], blocked = [], upgraded = [], todo = [];
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

    // Upgrade the tables that were already there.
    var T = {}; ['tblNotifyRoutes', 'tblNotifyConfig', 'tblNotifyTemplates'].forEach(function(n){ if (have[n]) T[n] = ctx.workbook.tables.getItem(n).getRange().load('values'); });
    if (T.tblNotifyRoutes || T.tblNotifyConfig || T.tblNotifyTemplates) {
      await ctx.sync();
      var plan = nfyUpgradePlan(T.tblNotifyRoutes && T.tblNotifyRoutes.values, T.tblNotifyConfig && T.tblNotifyConfig.values, T.tblNotifyTemplates && T.tblNotifyTemplates.values);
      if (plan.rename.length || plan.addColumns.length) {
        var rt = ctx.workbook.tables.getItem('tblNotifyRoutes');
        plan.rename.forEach(function(r){ rt.columns.getItemAt(r.index).name = r.to; upgraded.push('tblNotifyRoutes: column ' + r.to); });
        await ctx.sync();
        for (var i = 0; i < plan.addColumns.length; i++) {
          var c = plan.addColumns[i];
          try { rt.columns.add(null, null, c); await ctx.sync(); upgraded.push('tblNotifyRoutes: column ' + c + ' added'); }
          catch (e) { blocked.push('tblNotifyRoutes column ' + c + ' (the cells to the right of the table are in use)'); }
        }
      }
      if (plan.addConfig.length) {
        ctx.workbook.tables.getItem('tblNotifyConfig').rows.add(null, plan.addConfig);
        plan.addConfig.forEach(function(kv){ upgraded.push('tblNotifyConfig: ' + kv[0]); });
      }
      if (plan.templates.length) {
        var tb = ctx.workbook.tables.getItem('tblNotifyTemplates').getDataBodyRange();
        plan.templates.forEach(function(t){ tb.getCell(t.row, 1).values = [[t.value]]; upgraded.push('tblNotifyTemplates: ' + t.value); });
      }
      await ctx.sync();
    }
    return { sheet: NOTIFY_SETTINGS_SHEET, created: created, existing: existing, blocked: blocked, upgraded: upgraded };
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

// Everything the add-in's Settings panel shows:
//   { config: { Key: value }, routes: [{ provider, chatLink, chatId, mentionName, mentionId, email, enabled }], needsUpgrade }
// email = mentionId when it's an email address (the add-in also uses it to address MADRS reminders).
async function readNotifySettings(){
  return Excel.run(async function(ctx){
    var c = ctx.workbook.tables.getItemOrNullObject('tblNotifyConfig'), r = ctx.workbook.tables.getItemOrNullObject('tblNotifyRoutes'),
        p = ctx.workbook.tables.getItemOrNullObject('tblNotifyTemplates');
    await ctx.sync();
    var cv = c.isNullObject ? null : c.getRange().load('values'), rv = r.isNullObject ? null : r.getRange().load('values'),
        pv = p.isNullObject ? null : p.getRange().load('values');
    await ctx.sync();
    var config = {}; if (cv) cv.values.slice(1).forEach(function(row){ var k = String(row[0]).trim(); if (k) config[k] = row[1]; });
    var routes = [];
    if (rv) {
      var col = nfyHeaderIndex(rv.values[0]);
      var cell = function(row, name){ var i = col[name]; return i == null ? '' : String(row[i] == null ? '' : row[i]).trim(); };
      rv.values.slice(1).forEach(function(row){
        var prov = cell(row, 'provider'); if (!prov) return;
        var link = cell(row, 'chatlink') || cell(row, 'flowurl'), id = cell(row, 'mentionid') || cell(row, 'email');
        routes.push({ provider: prov, chatLink: link, chatId: nfyChatId(link), mentionName: cell(row, 'mentionname'), mentionId: id,
                      email: nfyIsEmail(id) ? id : '', enabled: col.enabled == null || String(row[col.enabled]).toUpperCase() !== 'FALSE' });
      });
    }
    var needsUpgrade = nfyUpgradePlan(rv && rv.values, cv && cv.values, pv && pv.values).needed;
    return { config: config, routes: routes, hasConfig: !!cv, hasRoutes: !!rv, needsUpgrade: needsUpgrade };
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

// One provider row in the table's own column order (columns found by header; unknown columns are left blank).
function nfyRouteRow(header, r){
  var v = { provider: String(r.provider || '').trim(), chatlink: String(r.chatLink || '').trim(), flowurl: String(r.chatLink || '').trim(),
            mentionname: String(r.mentionName || '').trim(), mentionid: String(r.mentionId || '').trim(), email: String(r.mentionId || '').trim(),
            enabled: r.enabled === false ? 'FALSE' : 'TRUE' };
  return header.map(function(h){ var k = String(h).trim().toLowerCase(); return v[k] == null ? '' : v[k]; });
}

// Add or update a provider's row from the add-in's Settings panel (matched by provider name, any case).
// r = { provider, chatLink, mentionName, mentionId, enabled }
async function upsertNotifyRoute(r){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem('tblNotifyRoutes');
    var all = t.getRange().load('values');
    await ctx.sync();
    var header = all.values[0], body = all.values.slice(1), pi = Math.max(0, nfyHeaderIndex(header).provider || 0);
    var key = String(r.provider).trim().toLowerCase();
    var idx = body.findIndex(function(x){ return String(x[pi]).trim().toLowerCase() === key; });
    if (idx < 0) idx = body.findIndex(function(x){ return !String(x[pi]).trim(); });   // reuse the starter row if it is still empty
    var row = nfyRouteRow(header, r);
    if (idx < 0) t.rows.add(null, [row]);
    else t.getDataBodyRange().getCell(idx, 0).getResizedRange(0, header.length - 1).values = [row];
    await ctx.sync();
  });
}

// Remove a provider's row (Settings panel "Remove"). The last row is blanked instead, since a table can't have no rows.
async function removeNotifyRoute(provider){
  return Excel.run(async function(ctx){
    var t = ctx.workbook.tables.getItem('tblNotifyRoutes');
    var all = t.getRange().load('values');
    await ctx.sync();
    var header = all.values[0], body = all.values.slice(1), pi = Math.max(0, nfyHeaderIndex(header).provider || 0);
    var key = String(provider).trim().toLowerCase();
    var idx = body.findIndex(function(x){ return String(x[pi]).trim().toLowerCase() === key; });
    if (idx < 0) return false;
    if (body.length === 1) t.getDataBodyRange().getCell(0, 0).getResizedRange(0, header.length - 1).values = [nfyRouteRow(header, { provider: '' })];
    else t.rows.getItemAt(idx).delete();
    await ctx.sync();
    return true;
  });
}

// ---- Wiring (taskpane.html, Settings › Teams notifications)
//   "Set up Teams notifications" / "Update for the new Teams flow" → ensureNotifySettings();  flow URL → setNotifyConfig('WebhookUrl', url)
//   provider rows → upsertNotifyRoute({ provider, chatLink, mentionName, mentionId, enabled }) / removeNotifyRoute(provider)
