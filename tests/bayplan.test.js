// Bay map plans: the old built-in drawing, saved as the workbook's map, draws as before; bpCheck guards every saved map; a clinic's
// own map draws and matches bays; no saved map (or a turned-down one) means no map, with bays shown as typed; Settings › Bay map
// loads, downloads or removes the map; the demo seeds one. The Bay Map Builder is its own page: the tracker only takes the file
// it downloads. FICTIONAL codes only.
//   node tests/bayplan.test.js
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert'), { execSync } = require('child_process');
const root = path.join(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
let old = null; try { old = execSync('git show main:src/taskpane.html', { cwd: root, encoding: 'utf8', maxBuffer: 64 << 20, stdio: ['ignore', 'pipe', 'ignore'] }).replace(/\r\n/g, '\n'); } catch {}

// Runs the bay map block of a taskpane.html in its own context, with the same stubs as baymap.test.js plus what Settings needs.
function load(s){
  const a = s.indexOf('// ---------- bay map:'), b = s.indexOf("document.addEventListener('toggle', e => { const d = e.target; if (d.dataset?.sec === 'bmap'");
  const pa = s.indexOf('// Vitals ranges (Settings'), pb = s.indexOf('function prnStart(');
  const fa = s.indexOf('function factsDL('), fb = s.indexOf('\n}\n', fa) + 3;
  assert(a > 0 && b > a && pa > 0 && pb > pa && fa > 0, 'block markers');
  const neu = s.includes('function bayPlan('), oa = s.indexOf('function bayClashes('), ob = s.indexOf('\n}\n', s.indexOf('function openBays(')) + 3;   // Plan bays
  assert(!neu || (oa > 0 && ob > oa), 'Plan bays markers');
  const C = { now: 650, today: [], saved: [], toasts: [] }; vm.createContext(C);
  vm.runInContext(`const obsAt = v => v?.adm ?? null, v40At = v => v.adm + 40, disAt = o => o + 100, fmtT = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');
const nowMin = () => now, S = { store: { setDoc: async (k, v) => { saved.push([k, v]); return true; } } }, tdType = () => 'Follow-up', tdLead = r => '<p class="lead">lead ' + r.key + '</p>', trackHTML = () => '<div class="track"></div>', stepsHTML = () => '<ol class="steps"></ol>';
const tdPrimary = r => ({ act: 'vit', html: '<button data-act="vit" data-k="' + r.key + '">Record</button>' }), vitGet = () => null;
const canonProv = s => String(s).trim(), PS = p => p !== 'BESSE', nm = m => m, CFG = { bayMap: true, me: 'Pat T', vitSysMax: 140, vitDiaMax: 90, vitHrMax: 110, vitHrMin: 50, vitO2Min: 92, vitEmSys: 180, vitEmDia: 120, vitalsMin: 40 };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fromHHMM = t => { const m = /^(\\d+):(\\d+)$/.exec(t || ''); return m ? +m[1] * 60 + +m[2] : null; }, rkText = () => '';
const todaySerial = () => 1, todayModel = () => today.map(([key, bay]) => ({ a: { key, active: true }, bay })), visitGet = k => ({ bay: today.find(x => x[0] === k)?.[1] });
const render = () => {}, toast = (k, h) => toasts.push([k, h]), downloads = [], downloadText = (n, t) => (downloads.push([n, t]), true);
let modal = null; const openModal = h => { modal = h; }, closeModal = () => { modal = null; }, getModal = () => modal;
const DAYNAME = ['Sunday'], dowOf = () => 0, fmtD = () => 'Oct 25', datalist = (id, L) => '<datalist id="' + id + '">' + L.map(v => '<option value="' + esc(v) + '">').join('') + '</datalist>';
` + s.slice(fa, fb) + s.slice(pa, pb) + s.slice(a, b) + (neu ? s.slice(oa, ob) : '') + '\nObject.assign(globalThis, { S, bayMapHTML, bayMapModel, bayPick, bayNo, ' + (neu ? 'bayPlan, bpCheck, BP_DEFAULT, bayMapSetHTML, bpLost, bpOpenHTML, BP_MAX, bayLabel, bayNextHTML, bpDownload, downloads, bpReadFile, getModal, bpRemoveAsk, openBays, openBayMove' : 'bayState') + ' });', C);
  return C;
}
const N = load(src);
let n = 0; const t = (name, fn) => Promise.resolve().then(fn).then(() => { n++; console.log('ok', name); });
const row = (key, bay, grp, rank) => ({ key, grp, rank, a: { mrn: key.toUpperCase(), time: 600, active: true, rendering: key.startsWith('o') ? 'Dr B' : 'Dr A' }, i: { adminKit: {}, inf: { mg: 84 } }, v: { start: 600, adm: 605, bay } });
const norm = h => h.replace(/ style="aspect-ratio:[^"]*"/, '').replace(/(<div class="room" style=")([^"]*)"/g, (_, p, st) => p + st.replace(/[\d.]+%/g, x => (+x.slice(0, -1)).toFixed(1) + '%') + '"').replace(/>\s+</g, '><').replace(/\s+/g, ' ').trim();
const mk = (...L) => ({ act: [1], need: L.filter(r => r.grp === 'need'), chair: L.filter(r => r.grp === 'chair'), wait: L.filter(r => r.grp === 'wait') });
const plan = () => ({ schema: 'mhs-bay-map', v: 1, name: 'Clinic A rooms', clinic: 'ZZ1', savedAt: '2026-10-24T15:00:00.000Z', savedBy: 'Sam U', w: 600, h: 300,
  bays: [{ id: 'A', cx: 100, cy: 150, r: 56, tag: [150, 150], lamp: [60, 110] }, { id: '2B', cx: 300, cy: 150, r: 56, tag: [350, 150], lamp: null }, { id: '10', cx: 480, cy: 150, r: 56, tag: [530, 150], lamp: [440, 110] }],
  curtains: [{ d: 'M200 20C210 120 210 180 200 280', bays: ['a', '2b'] }],
  walls: [{ d: 'M10 10H590V290H10Z' }, { d: 'M400 10A120 120 0 0 1 400 290', thin: true }], doors: [{ d: 'M60 290V242A48 48 0 0 0 12 290' }], windows: [{ d: 'M250 10H330' }],
  tables: [{ c: [60, 110, 22] }, { rect: [520, 40, 60, 24, 4], a: 90 }], chairs: [[100, 170, 0], [300, 170, 15], [480, 170, -30]],
  rooms: [{ t: 'Treatment', x: 300, y: 30 }, { t: 'Nook', x: 520, y: 270, small: true }], edit: { v: 1, items: [] } });

(async () => {
await t('the old built-in drawing, saved as the map: same drawing as main (bar whitespace, the aspect-ratio style and the room label %)', () => {
  if (!old){ console.log('  (skipped: no main branch here)'); return; }
  const O = load(old); assert(!O.bayPlan, 'main has no bayPlan');
  N.S.shared = { bayPlan: N.BP_DEFAULT };
  const cases = [
    () => mk(row('a', '2', 'need', 0), row('x', '9', 'need', 1), row('c', '3', 'chair', 4), row('nb', '', 'chair', 4)),
    () => mk(row('a', 'Bay 2', 'need', 0), row('b', '4', 'need', 1), row('c', '5', 'chair', 4), row('d', '1', 'need', 3)),
    () => mk(row('w', '5', 'wait', 9)),
    () => mk(row('pa', '2', 'need', 1), row('ob', '3', 'need', 0), row('pn', '5', 'chair', 4)),
  ];
  for (const [i, c] of cases.entries()) for (const prov of [undefined, 'Dr A']) for (const sel of [null, { key: 'c' }]){
    const a = O.bayMapHTML(c(), sel, prov), b = N.bayMapHTML(c(), sel, prov);
    assert(/aspect-ratio:808\/392/.test(b), 'aspect ratio from the plan');
    assert.strictEqual(norm(b), norm(a), `case ${i} prov ${prov} sel ${!!sel}`);
  }
});
await t('bpCheck: the built-in map passes; wrong schema, injection, non-finite, duplicates and oversize are turned down', () => {
  const ok = x => { const r = N.bpCheck(x); assert(r.plan, r.err); return r.plan; }, bad = (x, re) => { const r = N.bpCheck(x); assert(r.err && !r.plan, 'should fail: ' + re); assert(re.test(r.err), r.err); };
  ok(N.BP_DEFAULT); ok(plan());
  bad(null, /isn’t a bay map/); bad({ ...plan(), schema: 'other' }, /isn’t a bay map/); bad({ ...plan(), v: 2 }, /isn’t a bay map/);
  bad({ ...plan(), walls: [{ d: 'M0 0"/><script>alert(1)</script>' }] }, /wall is malformed/);
  bad({ ...plan(), doors: [{ d: 'javascript:alert(1)' }] }, /door is malformed/);
  bad({ ...plan(), curtains: [{ d: 'M0 0L10 10', bays: 'A' }] }, /curtain/);
  bad({ ...plan(), bays: [{ ...plan().bays[0], id: '<img src=x>' }] }, /can’t be used/);
  bad({ ...plan(), bays: [{ ...plan().bays[0], id: 'Bay4' }] }, /can’t be used/);
  bad({ ...plan(), bays: [{ ...plan().bays[0], cx: NaN }] }, /bad position/); bad({ ...plan(), bays: [{ ...plan().bays[0], r: Infinity }] }, /bad position/);
  bad({ ...plan(), bays: [{ ...plan().bays[0], cy: '150' }] }, /bad position/); bad({ ...plan(), chairs: [[1, 2, NaN]] }, /chair/);
  bad({ ...plan(), tables: [{ rect: [0, 0, -5, 5] }] }, /table/); bad({ ...plan(), w: Infinity }, /size/); bad({ ...plan(), w: 7000 }, /size/);
  bad({ ...plan(), bays: [plan().bays[1], { ...plan().bays[0], id: '2b' }] }, /Bay 2b is on the map twice/);
  bad({ ...plan(), bays: [] }, /no bays/);
  bad({ ...plan(), bays: Array.from({ length: N.BP_MAX.bays + 1 }, (_, i) => ({ ...plan().bays[0], id: String(i + 1) })) }, /more than 40 bays/);
  bad({ ...plan(), edit: { pad: 'x'.repeat(400001) } }, /too large/);
  const p = ok({ ...plan(), rooms: [{ t: '<script>x</script>', x: 10, y: 10 }], extra: 'dropped' });
  assert(!('extra' in p) && p.edit, 'only what the renderer reads, plus the edit model');
  assert.deepStrictEqual([...p.curtains[0].bays], ['A', '2B'], 'curtain bays follow the plan ids');
});
await t('custom map: letter ids, curved and thin walls, windows draw; bays match any case and "Bay"; sorted 2B before 10', () => {
  N.S.shared = { bayPlan: plan() }; const P = N.bayPlan();
  assert.strictEqual(P.name, 'Clinic A rooms'); assert.strictEqual(N.bayPlan(), P, 'cached while the stored map is the same object');
  assert.strictEqual(N.bayNo('bay 2b'), '2B'); assert.strictEqual(N.bayNo(' a '), 'A'); assert.strictEqual(N.bayNo('10'), '10'); assert.strictEqual(N.bayNo('4'), null);
  const m = mk(row('a', 'Bay 2b', 'need', 0), row('c', 'a', 'chair', 4), row('x', '4', 'need', 1));
  const M = N.bayMapModel(m); assert.deepStrictEqual([...M.bays.map(b => b.id)], ['A', '2B', '10']); assert.strictEqual(M.off.map(r => r.key).join(), 'x');
  const h = N.bayMapHTML(m, null);
  assert(h.includes('<div class="bplan" style="aspect-ratio:600/300">') && h.includes('viewBox="0 0 600 300"') && !/NaN|undefined/.test(h));
  assert(/data-b="2B" data-k="a" data-state="late"/.test(h) && /data-b="A" data-k="c" data-state="track"/.test(h) && /data-b="10" data-state="open"/.test(h));
  assert(h.includes('<g class="wall thin"><path d="M400 10A120 120 0 0 1 400 290"/></g>'), 'curved thin wall');
  assert(h.includes('<g class="window"><path class="wf" d="M250 10H330"/><path class="wg" d="M250 10H330"/></g>'), 'window');
  assert(h.includes('<div class="room small" style="left:86.667%;top:90.000%">Nook</div>') && h.includes('transform="rotate(90 550 52)"'));
  assert(/<g class="curtain"><path class="crail" d="M200 20C210 120 210 180 200 280"\/><path class="fabric"/.test(h), 'curtain drawn: A and 2B in use');
  assert.strictEqual((h.match(/class="lantern/g) || []).length, 2, 'bay 2B has no lantern'); assert(/Not on the map:.*X \(Bay 4\)/.test(h));
  assert.strictEqual(N.bayPick(N.bayMapModel(mk(row('p', '10', 'need', 1), row('q', '2B', 'need', 1))), null).id, '2B', 'numeric-aware order');
  const o = N.bpOpenHTML(P); assert.strictEqual((o.match(/class="bay" data-state="open"/g) || []).length, 3); assert(!/lantern on/.test(o) && !/fabric/.test(o));
});
await t('a saved map bpCheck turns down: nothing is drawn, bays show as typed, and Settings says why', () => {
  N.S.shared = { bayPlan: { ...plan(), walls: [{ d: '<svg onload=x>' }] } };
  assert.strictEqual(N.bayPlan(), null); assert.strictEqual(N.bayNo('A'), null); assert.strictEqual(N.bayLabel('Bay A'), 'A');
  const m = mk(row('a', 'A', 'need', 0), row('c', '2B', 'chair', 4));
  assert.strictEqual(N.bayMapHTML(m, null), '', 'no Bay map card on Today'); assert(!/class="bplan"/.test(N.bayMapHTML(m, null, 'Dr A')), 'ESK Bays: no map drawn');
  const h = N.bayMapSetHTML();
  assert(h.includes('<dd>Saved, but it can’t be used</dd>') && h.includes('The saved bay map can’t be used, so no map shows: A wall is malformed.'), h);
  assert(/data-act="bmDown">Download this map/.test(h) && /data-act="bmRemove">Remove the bay map/.test(h) && /data-act="bmFile">Load a map file…/.test(h));
  assert(!/built-in|bmReset/i.test(h), 'no built-in map to go back to');
  N.S.shared = { bayPlan: { schema: 'mhs-bay-map', v: 1, w: 1, h: 1 } }; assert.strictEqual(N.bayPlan(), null, 'never throws');
});
await t('Settings › Bay map: no map yet, or the saved map with who saved it', () => {
  N.S.shared = {}; let h = N.bayMapSetHTML();
  assert(h.includes('<dt>Bay map</dt><dd>No bay map yet</dd>') && !/msg warn/.test(h), h);
  assert(/<button class="btn primary" data-act="bmFile">Load a map file…<\/button><\/div>/.test(h), 'Load a map file is the one, primary, button');
  assert(!/bmDown|bmRemove|bmReset|bmEdit|built-in/i.test(h), 'nothing to download or remove');
  assert(h.includes('Draw your rooms, doors and bays in the <a href="https://brendannyates.github.io/bay-map-builder/" target="_blank" rel="noopener">Bay Map Builder ↗</a>, click Download file, then Load a map file here.'), h);
  N.S.shared = { bayPlan: { ...plan(), name: 'Rooms <b>A</b>' } }; h = N.bayMapSetHTML();
  assert(h.includes('<dd>Rooms &lt;b&gt;A&lt;/b&gt; · 3 bays · saved by Sam U Oct 24, 2026</dd>'), h);
  assert(/data-act="bmFile">Load a map file…/.test(h) && /data-act="bmDown">Download this map/.test(h) && /data-act="bmRemove">Remove the bay map/.test(h) && !/built-in|bmReset/i.test(h));
  assert(/A new map file replaces it for everyone/.test(h));
  assert(/\$\{bayMapSetHTML\(\)\}/.test(src.slice(src.indexOf('function settingsHTML('), src.indexOf('const imports = ', src.indexOf('function settingsHTML(')))), 'on the For everyone card');
});
await t('Plan bays: the map’s bay numbers to pick from; no map: Bay # is typed, no list', () => {
  N.today = [['k1', 'Bay 4'], ['k2', '']]; N.S.day = 1;
  N.S.shared = { bayPlan: plan() }; N.openBays(); let h = N.getModal();
  assert(h.includes('<datalist id="dlBayIds"><option value="A"><option value="2B"><option value="10"></datalist>') && /data-bay="k1" list="dlBayIds" value="Bay 4"/.test(h), h);
  N.S.shared = {}; N.openBays(); h = N.getModal();
  assert(!/dlBayIds/.test(h) && /<input class="fld" data-bay="k1" value="Bay 4"/.test(h) && /data-bay="k2" value=""/.test(h) && /data-act="baysSave"/.test(h), h);
  N.openBayMove('k1'); assert(/<h2>Plan bays/.test(N.getModal()), 'Change bay with no map opens Plan bays, not the map’s bay picker');
  vm.runInContext('closeModal()', N); N.today = [];
});
// The ACT handlers for the preview's Save and for Remove the bay map, run in the same context.
vm.runInContext('const ACT = {' + src.split('\n').filter(l => /^  bm(RemoveGo|Use): /.test(l)).join('\n') + '};\nglobalThis.ACT = ACT;', N);
const file = (x, size) => ({ size: size ?? 100, text: async () => typeof x === 'string' ? x : JSON.stringify(x) });
await t('Load a map file: bad files never reach the preview; a good one previews, warns about lost bays, saves stamped on Save', async () => {
  N.S.shared = { vitRange: { vitSysMax: 150 } }; N.today = [['k1', 'Bay 4'], ['k2', '7'], ['k3', ''], ['k4', '4']]; N.saved.length = 0; N.toasts.length = 0;
  assert(N.ACT.bmUse && N.ACT.bmRemoveGo, 'handlers found');
  await N.bpReadFile(file('{not json')); assert(/isn’t a bay map\. Choose the \.json file the Bay Map Builder downloaded/.test(N.toasts.at(-1)[1]));
  await N.bpReadFile(file(plan(), 3e6)); assert(/too large/.test(N.toasts.at(-1)[1]));
  await N.bpReadFile(file({ ...plan(), bays: [] })); assert.deepStrictEqual([...N.toasts.at(-1)], ['bad', 'The map has no bays.']);
  await N.bpReadFile(file({ ...plan(), walls: [{ d: 'M0 0"/><script>' }] })); assert(/wall is malformed/.test(N.toasts.at(-1)[1]));
  assert.strictEqual(N.getModal(), null, 'no preview for a bad file'); assert.strictEqual(N.S.bpNew, undefined); assert.strictEqual(N.saved.length, 0);
  await N.bpReadFile(file(plan())); const m = N.getModal();
  assert(m.includes('<h2>Use this bay map?</h2>') && m.includes('Clinic A rooms · 3 bays · ZZ1') && m.includes('data-act="bmUse">Save<') && m.includes('data-act="close">Cancel<'), m);
  assert(m.includes('sees it on Today and in ESK Bays.</div>'), 'no map yet: nothing to replace');
  assert(/Today’s visits use bays 4, 7, which aren’t on this map/.test(m), 'lost bays warned'); assert.strictEqual((m.match(/class="bay" data-state="open"/g) || []).length, 3);
  assert.strictEqual(N.S.bpNew.name, 'Clinic A rooms'); assert.strictEqual(N.saved.length, 0, 'nothing saved until Save');
  await N.ACT.bmUse(); assert.strictEqual(N.getModal(), null); assert.strictEqual(N.S.bpNew, null);
  const [k, v] = N.saved[0]; assert.strictEqual(k, 'eskShared'); assert.strictEqual(v.vitRange.vitSysMax, 150, 'other shared settings kept');
  assert.strictEqual(v.bayPlan.savedBy, 'Pat T'); assert(!isNaN(new Date(v.bayPlan.savedAt))); assert.strictEqual(N.bayPlan().name, 'Clinic A rooms');
  assert.deepStrictEqual([...N.toasts.at(-1)], ['ok', 'Bay map saved for everyone in this workbook.']);
  await N.ACT.bmUse(); assert.strictEqual(N.saved.length, 1, 'Save twice saves once');
  N.today = [['k1', 'A']]; await N.bpReadFile(file(plan())); assert(!/msg warn/.test(N.getModal()), 'every bay on the map: no warning');
  assert(N.getModal().includes('sees it on Today and in ESK Bays, in place of the map saved now.</div>'), 'a map is saved: this one replaces it');
  const set = N.S.store.setDoc; N.S.store.setDoc = async () => false; await N.ACT.bmUse(); N.S.store.setDoc = set;
  assert(N.toasts.at(-1)[0] === 'bad' && /didn’t save/.test(N.toasts.at(-1)[1])); assert.strictEqual(N.saved.length, 1);
});
await t('Remove the bay map: asks first, then deletes only bayPlan from eskShared; no map after', async () => {
  const keep = { vitRange: { vitSysMax: 150 }, medInvUrl: 'https://example.invalid/inv', soundLinks: { bing: 'https://example.invalid/b' } };
  N.S.shared = { ...keep, bayPlan: plan() }; N.saved.length = 0;
  assert(/data-act="bmRemove">Remove the bay map/.test(N.bayMapSetHTML()));
  N.bpRemoveAsk(); const m = N.getModal();
  assert(m.includes('<h2>Remove the bay map?</h2>') && /stop showing a map for everyone/.test(m) && m.includes('data-act="bmRemoveGo">Remove the bay map<') && m.includes('data-act="close">Cancel<'), m);
  assert.strictEqual(N.saved.length, 0, 'nothing removed until confirmed');
  await N.ACT.bmRemoveGo(); assert.strictEqual(N.getModal(), null);
  const [k, v] = N.saved[0]; assert.strictEqual(k, 'eskShared'); assert.deepStrictEqual(JSON.parse(JSON.stringify(v)), keep, 'every other shared key kept');
  assert.strictEqual(N.bayPlan(), null); assert.deepStrictEqual([...N.toasts.at(-1)], ['ok', 'Bay map removed. Bays show as typed.']);
  assert(N.bayMapSetHTML().includes('<dd>No bay map yet</dd>') && !/bmRemove/.test(N.bayMapSetHTML()), 'nothing to remove now');
  assert.strictEqual(N.bayMapHTML(mk(row('a', '2', 'need', 0)), null), '', 'Today: no Bay map card');
});
await t('a bay that isn’t on the map still shows where the patient is: emergency banner and alert, provider feed, Still to come', () => {
  N.S.shared = { bayPlan: plan() };
  assert.strictEqual(N.bayNo('5'), null); assert.strictEqual(N.bayLabel('5'), '5'); assert.strictEqual(N.bayLabel(' Bay 5 '), '5');
  assert.strictEqual(N.bayLabel('bay 2b'), '2B'); assert.strictEqual(N.bayLabel(''), ''); assert.strictEqual(N.bayLabel(undefined), '');
  const fa = src.indexOf('function emBannerHTML(){'), fb = src.indexOf('\n}\n', fa) + 3; assert(fa > 0);
  vm.runInContext('const emActive = () => emList;' + src.slice(fa, fb) + '\nglobalThis.emBannerHTML = emBannerHTML;', N);
  N.emList = [{ key: 'k|1', rid: 'r1', reading: '182/124', what: 'Initial vitals', at: 610, open: true, acks: [], a: { mrn: 'ZZTEST' }, v: { bay: '5' } }];
  const h = N.emBannerHTML(); assert(h.includes('EMERGENCY · 182/124</b> · ZZTEST · Bay 5 · Initial vitals'), h);
  N.emList[0].v = {}; assert(!/Bay/.test(N.emBannerHTML()), 'no bay typed: none shown');
  assert(src.includes("body: bayLabel(e.v.bay) ? `Bay ${bayLabel(e.v.bay)}` : 'ESK'"), 'emergency pop-up body');
  assert(src.includes("ev = { b: bayLabel(v.bay) ? `Bay ${bayLabel(v.bay)}` : 'No bay'"), 'provider event feed');
  assert(src.includes('bay: bayLabel(v.bay), clinic:'), 'Teams notify: the bay as people read it');
  const w = N.bayNextHTML(mk(row('w', '5', 'wait', 9)));
  assert(w.includes('<span class="where">Bay 5</span>') && w.includes('>Change bay</button>'), w);
});
await t('no bay map at all: alarms, emergency banner and Still to come keep the bay as typed; no map-only actions', () => {
  N.S.shared = {};
  assert.strictEqual(N.bayPlan(), null); assert.strictEqual(N.bayNo('4'), null); assert.strictEqual(N.bayNo('Bay 4'), null);
  assert.strictEqual(N.bayLabel('Bay 4'), '4'); assert.strictEqual(N.bayLabel(' 2a '), '2a'); assert.strictEqual(N.bayLabel(''), '');
  N.emList = [{ key: 'k|1', rid: 'r1', reading: '182/124', what: 'Initial vitals', at: 610, open: true, acks: [], a: { mrn: 'ZZTEST' }, v: { bay: 'Bay 4' } }];
  assert(N.emBannerHTML().includes('EMERGENCY · 182/124</b> · ZZTEST · Bay 4 · Initial vitals'));
  const m = mk(row('a', '2', 'need', 0), row('w', 'Bay 4', 'wait', 9), row('x', '', 'wait', 9));
  assert.strictEqual(N.bayMapHTML(m, null), '', 'Today: no Bay map card');
  const h = N.bayNextHTML(m);
  assert(h.includes('<span class="where">Bay 4</span><span></span>') && h.includes('<span class="where none">No bay yet</span><span></span>') && !/bayMove|Change bay|Choose a bay/.test(h), h);
  N.downloads.length = 0; N.toasts.length = 0; N.bpDownload();
  assert.strictEqual(N.downloads.length, 0); assert.deepStrictEqual([...N.toasts.at(-1)], ['warn', 'There’s no bay map to download.']);
});
await t('a turned-down saved map: Download keeps the stored file; Settings says fix it in the Bay Map Builder when it has an edit model, else load the file again',() => {
  const edit = { v: 1, name: 'Clinic A rooms', items: [{ id: 'b1', k: 'bay', no: 'A', x: 100, y: 150, r: 58 }] };
  N.S.shared = { bayPlan: { ...plan(), walls: [{ d: '<svg onload=x>' }], edit } }; N.today = [];
  const h = N.bayMapSetHTML();
  assert(h.includes('A wall is malformed. Download this map, open it in the <a href="https://brendannyates.github.io/bay-map-builder/" target="_blank" rel="noopener">Bay Map Builder ↗</a> with Open a map file, fix it there and load it here again.</div>'), h);
  assert(/data-act="bmDown">Download this map/.test(h) && /data-act="bmFile">Load a map file…/.test(h) && /data-act="bmRemove"/.test(h));
  N.downloads.length = 0; N.toasts.length = 0; N.bpDownload();
  const [fn, txt] = N.downloads[0]; assert.strictEqual(fn, 'Bay map - Clinic A rooms (can’t be used).json');
  assert.deepStrictEqual(JSON.parse(txt), JSON.parse(JSON.stringify(N.S.shared.bayPlan)), 'the stored map as it is, edit model and all');
  assert(/downloaded as it is/.test(N.toasts.at(-1)[1]));
  N.S.shared = { bayPlan: { ...plan(), name: 'a/b', walls: [{ d: 'x' }], edit: [1] } };
  const h2 = N.bayMapSetHTML(); assert(h2.includes('A wall is malformed. Load the file again.</div>') && !/with Open a map file/.test(h2), 'edit model isn’t an object: the builder can’t open it either');
  N.S.shared = { bayPlan: { ...plan(), walls: [{ d: 'x' }], edit: undefined } }; assert(N.bayMapSetHTML().includes('A wall is malformed. Load the file again.</div>'), 'no edit model');
  N.S.shared = { bayPlan: { ...plan(), name: 'a/b', walls: [{ d: 'x' }], edit: [1] } };
  N.downloads.length = 0; N.bpDownload(); assert.strictEqual(N.downloads[0][0], 'Bay map - a-b (can’t be used).json');
  N.S.shared = { bayPlan: plan() }; N.downloads.length = 0; N.bpDownload();
  assert.strictEqual(N.downloads[0][0], 'Bay map - Clinic A rooms.json'); assert.strictEqual(JSON.parse(N.downloads[0][1]).name, 'Clinic A rooms', 'a good saved map: the map in use');
  N.S.shared = {}; N.downloads.length = 0; N.bpDownload(); assert.strictEqual(N.downloads.length, 0, 'no saved map: nothing to download');
});
await t('nothing falls back to the built-in drawing: BP_DEFAULT is used only by the shared block and the demo', () => {
  const a = src.indexOf('// ==== bay plan: shared'), b = src.indexOf('// ==== end bay plan ====');
  const out = (src.slice(0, a) + src.slice(b)).split('\n').filter(l => l.includes('BP_DEFAULT') && !/^\s*\/\//.test(l));
  assert.strictEqual(out.length, 1, out.join('\n')); assert(/^  const docs = \{ eskShared: \{ bayPlan: \{ \.\.\.JSON\.parse\(JSON\.stringify\(BP_DEFAULT\)\)/.test(out[0]), out[0]);
  const ds = src.indexOf('function DemoStore(){'), de = src.indexOf('\nfunction ', ds + 10), at = src.indexOf(out[0]); assert(at > ds && at < de, 'inside DemoStore');
});
await t('the demo seeds the built-in drawing as its saved map, so the Tracker and ESK Bays demos show the map', () => {
  const line = src.split('\n').find(l => l.startsWith('  const docs = { eskShared: { bayPlan:'));
  vm.runInContext(line.replace(/^  const docs =/, 'globalThis.demoDocs =').replace(/;\s*\/\/.*$/, ';'), N);
  const P = N.demoDocs.eskShared.bayPlan; assert(N.bpCheck(P).plan, 'passes bpCheck');
  assert.strictEqual(P.name, 'ESK rooms (demo)'); assert.strictEqual(P.savedBy, 'Taylor R.'); assert.strictEqual(P.bays.map(b => b.id).join(), '1,2,3,4,5');
  assert.notStrictEqual(P.bays, N.BP_DEFAULT.bays, 'a copy: the demo can remove it without touching BP_DEFAULT');
  const sd = src.slice(src.indexOf('async function seedDemo(){'), src.indexOf('\n}\n', src.indexOf('async function seedDemo(){')));
  for (const id of [...sd.matchAll(/bay: '(\w+)'/g)].map(x => x[1])) assert(P.bays.some(b => b.id === id), 'seeded visit bay ' + id + ' is on the demo map');
  N.S.shared = N.demoDocs.eskShared;
  const m = mk(row('a', '1', 'need', 0), row('ob', '4', 'need', 1)); assert(/aspect-ratio:808\/392/.test(N.bayMapHTML(m, null)) && /aspect-ratio:808\/392/.test(N.bayMapHTML(m, null, 'Dr A')), 'Tracker and ESK Bays draw it');
  assert(N.bayMapSetHTML().includes('<dd>ESK rooms (demo) · 5 bays · saved by Taylor R. '), N.bayMapSetHTML());
});
await t('the Bay Map Builder is its own page: no opener, no messages, no frame CSS, in src and root', () => {
  for (const [f, s] of [['src', src], ['root', fs.readFileSync(path.join(root, 'taskpane.html'), 'utf8')]]){
    for (const w of ['bmHandle', 'bmReady', 'bmLoad', 'bmSave', 'bmClose', 'bmEdit', 'BAY_BUILDER_URL', 'openBayBuilder', 'bpBroken', '#cfloat.bm', "className = 'bm'", 'from=esk', 'Edit in the Bay Map Builder'])
      assert(!s.includes(w), f + ' still has ' + w);
    assert.strictEqual((s.match(/github\.io\/bay-map-builder\//g) || []).length, 1, f + ': one plain link to the builder');
    assert(!/displayDialogAsync\([^)]*bay-map/.test(s) && !/<iframe[^>]*Bay Map/.test(s), f);
  }
});
await t('narrow maps: tags are one line no wider than 0.19 of the map, the size the builder plans for', () => {
  const r = /@container \(max-width:560px\)\{([^\n]*)\}\n/.exec(src)?.[1] || '';
  assert(/\.bplan \.btag\{display:flex;align-items:center;[^}]*max-width:19cqw\}/.test(r), r);
  assert(/\.bplan \.btag \.s\{order:-1;[^}]*margin-top:0\}/.test(r) && /\.bplan \.btag \.tx\{display:none\}/.test(r), 'the dot moves in front of the number; the status text hides');
});
console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
