// Bay map plans: the built-in map draws as before, bpCheck guards every saved map, a clinic's own map draws and matches bays,
// and Settings › Bay map shows which map is in use. FICTIONAL codes only.
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
  const C = { now: 650, today: [], saved: [], toasts: [] }; vm.createContext(C);
  vm.runInContext(`const obsAt = v => v?.adm ?? null, v40At = v => v.adm + 40, disAt = o => o + 100, fmtT = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');
const nowMin = () => now, S = { store: { setDoc: async (k, v) => { saved.push([k, v]); return true; } } }, tdType = () => 'Follow-up', tdLead = r => '<p class="lead">lead ' + r.key + '</p>', trackHTML = () => '<div class="track"></div>', stepsHTML = () => '<ol class="steps"></ol>';
const tdPrimary = r => ({ act: 'vit', html: '<button data-act="vit" data-k="' + r.key + '">Record</button>' }), vitGet = () => null;
const canonProv = s => String(s).trim(), PS = p => p !== 'BESSE', nm = m => m, CFG = { bayMap: true, me: 'Pat T', vitSysMax: 140, vitDiaMax: 90, vitHrMax: 110, vitHrMin: 50, vitO2Min: 92, vitEmSys: 180, vitEmDia: 120, vitalsMin: 40 };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fromHHMM = t => { const m = /^(\\d+):(\\d+)$/.exec(t || ''); return m ? +m[1] * 60 + +m[2] : null; }, rkText = () => '';
const todaySerial = () => 1, todayModel = () => today.map(([key, bay]) => ({ a: { key, active: true }, bay })), visitGet = k => ({ bay: today.find(x => x[0] === k)?.[1] });
const render = () => {}, toast = (k, h) => toasts.push([k, h]), downloads = [], downloadText = (n, t) => (downloads.push([n, t]), true);
` + s.slice(fa, fb) + s.slice(pa, pb) + s.slice(a, b) + '\nObject.assign(globalThis, { S, bayMapHTML, bayMapModel, bayPick, bayNo, ' + (s.includes('function bayPlan(') ? 'bayPlan, bpCheck, BP_DEFAULT, bayMapSetHTML, bmHandle, bpLost, bpOpenHTML, BP_MAX, bayLabel, bayNextHTML, bpDownload, downloads' : 'bayState') + ' });', C);
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
await t('built-in map: same drawing as main (bar whitespace, the aspect-ratio style and the room label %)', () => {
  if (!old){ console.log('  (skipped: no main branch here)'); return; }
  const O = load(old); assert(!O.bayPlan, 'main has no bayPlan');
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
await t('a saved map bpCheck turns down: the built-in map draws and Settings says why', () => {
  N.S.shared = { bayPlan: { ...plan(), walls: [{ d: '<svg onload=x>' }] } };
  assert.strictEqual(N.bayPlan(), N.BP_DEFAULT); assert.strictEqual(N.bayNo('4'), '4');
  const h = N.bayMapSetHTML(); assert(/Built-in \(5 bays\)/.test(h) && /can’t be used.*A wall is malformed/.test(h) && /data-act="bmReset"/.test(h));
  N.S.shared = { bayPlan: { schema: 'mhs-bay-map', v: 1, w: 1, h: 1 } }; assert.strictEqual(N.bayPlan(), N.BP_DEFAULT, 'never throws');
});
await t('Settings › Bay map: built-in or the saved map with who saved it', () => {
  N.S.shared = {}; let h = N.bayMapSetHTML();
  assert(h.includes('<dt>Bay map</dt><dd>Built-in (5 bays)</dd>') && !/bmReset/.test(h) && !/msg warn/.test(h));
  assert(/data-act="bmEdit">Edit in the Bay Map Builder/.test(h) && /data-act="bmFile">Load a map file…/.test(h) && /data-act="bmDown">Download this map/.test(h));
  N.S.shared = { bayPlan: { ...plan(), name: 'Rooms <b>A</b>' } }; h = N.bayMapSetHTML();
  assert(h.includes('<dd>Rooms &lt;b&gt;A&lt;/b&gt; · 3 bays · saved by Sam U Oct 24, 2026</dd>'), h); assert(/data-act="bmReset">Use the built-in map/.test(h));
  assert(/\$\{bayMapSetHTML\(\)\}/.test(src.slice(src.indexOf('function settingsHTML('), src.indexOf('const imports = ', src.indexOf('function settingsHTML(')))), 'on the For everyone card');
  assert(/list="dlBayIds"/.test(src) && /datalist\('dlBayIds', bayPlan\(\)\.bays\.map\(b => b\.id\)\)/.test(src), 'Plan bays offers the map’s bay numbers');
});
await t('builder messages: bmReady gets the map and today’s bays; bmSave checks, warns about lost bays, saves stamped; bmClose closes', async () => {
  N.S.shared = { vitRange: { vitSysMax: 150 } }; N.today = [['k1', 'Bay 4'], ['k2', '7'], ['k3', ''], ['k4', '4']]; N.saved.length = 0;
  const out = []; let closed = 0; const send = o => out.push(JSON.parse(JSON.stringify(o))), close = () => closed++;
  await N.bmHandle({ ev: 'bmReady' }, send, close);
  assert.deepStrictEqual(out[0], { ev: 'bmLoad', plan: null, bays: ['4', '7'] }, 'built-in goes as null');
  await N.bmHandle({ ev: 'bmSave', plan: { ...plan(), bays: [] } }, send, close);
  assert.deepStrictEqual(out[1], { ev: 'bmSaved', ok: false, err: 'The map has no bays.' }); assert.strictEqual(N.saved.length, 0);
  await N.bmHandle({ ev: 'bmSave', plan: plan() }, send, close);
  assert.strictEqual(out[2].ok, true); assert(/Today’s visits use bays 4, 7, which aren’t on this map/.test(out[2].warn), out[2].warn);
  const [k, v] = N.saved[0]; assert.strictEqual(k, 'eskShared'); assert.strictEqual(v.vitRange.vitSysMax, 150, 'other shared settings kept');
  assert.strictEqual(v.bayPlan.savedBy, 'Pat T'); assert(!isNaN(new Date(v.bayPlan.savedAt))); assert.strictEqual(N.S.shared.bayPlan.name, 'Clinic A rooms');
  await N.bmHandle({ ev: 'bmReady' }, send, close); assert.strictEqual(out[3].plan.name, 'Clinic A rooms', 'the saved map goes back to the builder');
  assert.deepStrictEqual(out[3].bays, ['4', '7']);
  N.S.store.setDoc = async () => false; await N.bmHandle({ ev: 'bmSave', plan: plan() }, send, close); assert(out[4].ok === false && /didn’t save/.test(out[4].err));
  await N.bmHandle({ ev: 'bmClose' }, send, close); await N.bmHandle({ ev: 'other' }, send, close); assert.strictEqual(closed, 1); assert.strictEqual(out.length, 5);
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
  const w = N.bayNextHTML(mk(row('w', '5', 'wait', 9)));
  assert(w.includes('<span class="where">Bay 5</span>') && w.includes('>Change bay</button>'), w);
});
await t('a turned-down saved map: the builder gets its edit model, Download keeps the stored file, Settings says what works', async () => {
  const edit = { v: 1, name: 'Clinic A rooms', items: [{ id: 'b1', k: 'bay', no: 'A', x: 100, y: 150, r: 58 }] };
  N.S.shared = { bayPlan: { ...plan(), walls: [{ d: '<svg onload=x>' }], edit } }; N.today = [];
  const out = [], send = o => out.push(JSON.parse(JSON.stringify(o)));
  await N.bmHandle({ ev: 'bmReady' }, send, () => {});
  assert.deepStrictEqual(out[0], { ev: 'bmLoad', plan: { schema: 'mhs-bay-map', v: 1, edit }, bays: [] });
  assert(/Open it in the Bay Map Builder to fix it, or load the file again\./.test(N.bayMapSetHTML()));
  N.downloads.length = 0; N.toasts.length = 0; N.bpDownload();
  const [fn, txt] = N.downloads[0]; assert.strictEqual(fn, 'Bay map - Clinic A rooms (can’t be used).json');
  assert.deepStrictEqual(JSON.parse(txt), JSON.parse(JSON.stringify(N.S.shared.bayPlan)), 'the stored map, as it is');
  assert(/downloaded as it is/.test(N.toasts.at(-1)[1]));
  N.S.shared = { bayPlan: { ...plan(), name: 'a/b', walls: [{ d: 'x' }], edit: [1] } }; out.length = 0;
  await N.bmHandle({ ev: 'bmReady' }, send, () => {}); assert.strictEqual(out[0].plan, null, 'no usable edit model: the builder starts from the built-in');
  const h = N.bayMapSetHTML(); assert(/Load the file again\.<\/div>/.test(h) && !/Bay Map Builder to fix/.test(h));
  N.downloads.length = 0; N.bpDownload(); assert.strictEqual(N.downloads[0][0], 'Bay map - a-b (can’t be used).json');
  N.S.shared = {}; N.downloads.length = 0; N.bpDownload(); assert.strictEqual(N.downloads[0][0], 'Bay map - ESK rooms (built-in).json', 'no saved map: the built-in');
});
await t('narrow maps: tags are one line no wider than 0.19 of the map, the size the builder plans for', () => {
  const r = /@container \(max-width:560px\)\{([^\n]*)\}\n/.exec(src)?.[1] || '';
  assert(/\.bplan \.btag\{display:flex;align-items:center;[^}]*max-width:19cqw\}/.test(r), r);
  assert(/\.bplan \.btag \.s\{order:-1;[^}]*margin-top:0\}/.test(r) && /\.bplan \.btag \.tx\{display:none\}/.test(r), 'the dot moves in front of the number; the status text hides');
});
console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
