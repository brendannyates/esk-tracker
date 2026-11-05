// Who is using the tracker, per workbook: the workbook key, the clinic from the workbook's name, the daily pick's order, what a
// profile row and this computer keep (never the clinic's settings), signature files, the admin code, and that nothing carries over
// from one workbook to another. Pulled from src/taskpane.html.   node tests/users.test.js   FICTIONAL names only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const C = { TextDecoder, TextEncoder, crypto: globalThis.crypto }; vm.createContext(C);
vm.runInContext(between('// ---------- who is using it: pure helpers', '// ---------- who is using it: end')
  + '\nObject.assign(globalThis, { wbHash, wbKeyFor, clinicFromName, pickOrder, PERSONAL_CFG, profCfg, CLINIC_OWN, localOnly, sigFromFile, htmlText });', C);
vm.runInContext(between('async function pinHash(', 'function adminCardHTML') .replace(/^/, '') + '\n'
  + between('const ADMIN_CODE =', '\n', ) + '\n' + between('async function adminCodeOk(', '\n') + '\nObject.assign(globalThis, { pinHash, adminCodeOk, ADMIN_CODE });', C);
let n = 0; const tests = [];
const eq = (a, b, m) => assert.strictEqual(JSON.stringify(a), JSON.stringify(b), m);   // values from the vm context are another realm
const t = (name, fn) => tests.push([name, fn]);

t('workbook key: the address (hashed, case and query ignored) comes first, then the saved id; nothing gives none', () => {
  const a = C.wbKeyFor('https://example.sharepoint.com/sites/x/Shared Documents/Castro ESK Tracker.xlsx?web=1', 'abc');
  assert.strictEqual(a, C.wbKeyFor('HTTPS://example.sharepoint.com/sites/x/Shared Documents/Castro ESK Tracker.xlsx', 'zzz'));
  assert.notStrictEqual(a, C.wbKeyFor('https://example.sharepoint.com/sites/x/Shared Documents/Oakland ESK Tracker.xlsx', 'abc'), 'another workbook, another key');
  assert(/^u[0-9a-z]+$/.test(a) && !a.includes('sharepoint'), 'hashed, not the address');
  const doc = g => 'https://example.sharepoint.com/:x:/r/personal/x/_layouts/15/Doc.aspx?sourcedoc=%7B' + g + '%7D&action=edit';
  assert.notStrictEqual(C.wbKeyFor(doc('AAAA'), ''), C.wbKeyFor(doc('BBBB'), ''), 'a page address keeps the query that names the file');
  assert.strictEqual(C.wbKeyFor('', 'k1'), 'ik1'); assert.strictEqual(C.wbKeyFor('', ''), '');
});
t('clinic from the workbook name: before the modality word, or after a dash; codes for the three known ones; else null', () => {
  const f = x => { const r = C.clinicFromName(x); return r ? r.name + '|' + r.code : null; };
  assert.strictEqual(f('San Rafael ESK Tracker.xlsx'), 'San Rafael|SRL');
  assert.strictEqual(f('Castro MagV Tracker'), 'Castro|CAS');
  assert.strictEqual(f('Walnut Creek Esketamine Master Tracker 2026.xlsx'), 'Walnut Creek|Walnut Creek');
  assert.strictEqual(f('Esketamine Master Tracker - Oakland.xlsx'), 'Oakland|Oakland');
  assert.strictEqual(f('SRL ESK Tracker'), 'San Rafael|SRL');
  assert.strictEqual(f('Copy of Castro ESK Tracker (2).xlsx'), 'Castro|CAS');
  for (const x of ['Esketamine Master Tracker.xlsx', 'ESK Tracker', 'Tracker', '', null, 'Copy of Esketamine Master Tracker', 'Test ESK Tracker', 'Template ESK Tracker'])
    assert.strictEqual(C.clinicFromName(x), null, String(x));
});
t('daily pick: only registered people with a name, this computer\'s first ("On this computer"), each group by name, none picked', () => {
  const P = [{ key: 'name:zed q.', name: 'Zed Q.' }, { key: 'name:amy b.', name: 'Amy B.' }, { key: 'name:kim l.', name: 'Kim L.' }, { key: 'x', name: '  ' }, { key: '', name: 'No Key' }];
  const L = C.pickOrder(P, ['name:zed q.']);
  eq(L.map(p => p.name + (p.here ? '*' : '')), ['Zed Q.*', 'Amy B.', 'Kim L.']);
  eq(C.pickOrder(P, []).map(p => p.here), [false, false, false]);
  eq(C.pickOrder(null, null), []);
});
t('a profile row keeps only a person\'s own settings: an old row\'s clinic, cc, practice manager and emails are dropped', () => {
  const old = { me: 'Amy B.', sig: 'Amy', theme: 'dark', vitNotify: true, sup: 'Dr Example', prov: 'Provider A', clinic: 'SRL', pm: 'Pat M.', cc: 'lead@example.com', helpEmail: 'help@example.com', leadEmail: 'x@example.com', provEmails: 'A = a@example.com', email: 'o@example.com', vitalsMin: 30, profKey: 'k', profAt: 't' };
  const p = C.profCfg(old);
  for (const k of ['clinic', 'pm', 'cc', 'helpEmail', 'leadEmail', 'provEmails', 'email', 'vitalsMin', 'profKey', 'profAt']) assert(!(k in p), k);
  eq(Object.keys(p).sort(), ['me', 'prov', 'sig', 'sup', 'theme', 'vitNotify']);
  for (const k of Object.keys(p)) assert(C.PERSONAL_CFG.includes(k));
  for (const k of ['clinic', 'locName', 'pm', 'cc', 'helpEmail', 'leadEmail', 'genoaEmail', 'provEmails']) assert(!C.PERSONAL_CFG.includes(k), k + ' is never personal');
});
t('this computer\'s copy for a person leaves out every clinic setting and the slot key', () => {
  const keys = ['email', 'genoaEmail', 'cc', 'helpEmail', 'leadEmail', 'provEmails', 'vitalsMin', 'disAfter', 'schedDays', 'genoaDays', 'vitSysMax', ...C.CLINIC_OWN];
  const o = C.localOnly({ me: 'Amy B.', schedOpen: true, profAt: 't', profKey: 'k', clinic: 'SRL', pm: 'Pat M.', cc: 'c@example.com', locName: 'San Rafael', vitSysMax: 150 }, keys);
  eq(Object.keys(o).sort(), ['me', 'profAt', 'schedOpen']);
});
t('signature files: .txt as is, .htm as plain text, UTF-16 and Windows-1252 read, blank runs collapsed, 12 lines at most', () => {
  const enc = s => new TextEncoder().encode(s);
  assert.strictEqual(C.sigFromFile('Work.txt', enc('\r\n  Amy B.\r\nPCT\r\n\r\n\r\nMindful Health Solutions  \r\n\r\n')), 'Amy B.\nPCT\n\nMindful Health Solutions');
  const htm = '<html><head><style>p{}</style><title>x</title></head><body><p class=MsoNormal>Amy B.<o:p></o:p></p><p>PCT &amp; tech<br>Phone 555-0100</p><!-- c --><script>alert(1)</script></body></html>';
  assert.strictEqual(C.sigFromFile('Work.htm', enc(htm)), 'Amy B.\nPCT & tech\nPhone 555-0100');
  const u16 = Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from('Kim L.\r\nRN', 'utf16le')]);
  assert.strictEqual(C.sigFromFile('Sig.txt', new Uint8Array(u16)), 'Kim L.\nRN');
  assert.strictEqual(C.sigFromFile('Sig.txt', new Uint8Array([0x4A, 0xF6, 0x72, 0x67])), 'Jörg', 'Windows-1252 when not UTF-8');
  assert.strictEqual(C.sigFromFile('Long.txt', enc(Array.from({ length: 20 }, (_, i) => 'line ' + i).join('\n'))).split('\n').length, 12);
  assert(!/[<>]/.test(C.sigFromFile('x.html', enc('<b>Bold</b> <a href="javascript:x">link</a>'))));
});
t('admin code: 0582 unlocks (also with spaces), nothing else does; the code itself is not in the file', async () => {
  assert.strictEqual(await C.adminCodeOk('0582'), true); assert.strictEqual(await C.adminCodeOk(' 0582 '), true);
  for (const v of ['', '582', '05820', '1234', '0000']) assert.strictEqual(await C.adminCodeOk(v), false, v);
  const block = between('const ADMIN_CODE =', 'function adminCardHTML'); assert(!block.includes("'0582'") && !/\b0582\b/.test(block));
});
t('nothing carries over between workbooks: no SRL / San Rafael defaults, no seeding from the latest profile, no legacy reads', () => {
  const def = between('const DEF = {', '};'); assert(!/SRL|San Rafael/.test(def), 'DEF has no clinic'); assert(/clinic:''/.test(def) && /pm:''/.test(def) && /cc:''/.test(def) && /helpEmail:''/.test(def));
  assert(!src.includes('S._seeded'), 'no "new person seeded from the latest profile"');
  const sl = between('function saveLocal(){', '\n}\n'); assert(!sl.includes("'eskCfg'") && !sl.includes('LEGACY_CFG'), 'saves never go to the shared eskCfg');
  const reads = [...src.matchAll(/localStorage\.getItem\(([^)]*)\)/g)].map(m => m[1].trim());
  assert(reads.every(r => r !== "'eskCfg'" && r !== 'CFG_KEY'), reads.join());
  assert.strictEqual(reads.filter(r => r === 'LEGACY_CFG').length, 2, 'eskCfg is read only for the theme fallback and Settings › Use the values this computer had before');
  assert(!/return t\.some\(Boolean\) \? t : \[CFG\.me/.test(src), 'no "signature blank → your name + clinic"');
  assert(/sentBy: CFG\.me \|\| ''/.test(src), 'Teams Sent By is the person picked today, blank otherwise');
});
t('the user picker only lists the user sheet: no datalist of names from the Inventory, no name typed into Settings', () => {
  assert(!src.includes("datalist('dlMe'")); assert(!/f\('sMe'/.test(src));
  const ph = between('function pickHTML(){', '\n}\n'); assert(ph.includes('pickOrder(S.profiles'), 'the pick reads the profile rows only');
});

t('this computer, per workbook: a person picked in one workbook is nobody in another, and clinic settings never come from the computer', () => {
  const store = {}, ls = { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: k => { delete store[k]; } };
  const D = { localStorage: ls, location: { search: '' } }; vm.createContext(D);
  vm.runInContext(`const S = { shared: {}, wbName: '' }, applyTheme = () => {}, todaySerial = () => 46000, isoOf = s => 'day' + s;
const VIT_RANGE = ['vitSysMax', 'vitDiaMax', 'vitHrMax', 'vitHrMin', 'vitO2Min', 'vitEmSys', 'vitEmDia'];\n`
    + between('// ---------- settings (no patient data)', '// ---------- user profiles:')
    + between('const CLINIC_SHARED =', '// Settings › Clinic: what this computer had before')
    + `\nObject.defineProperty(globalThis, 'CFG', { get(){ return CFG; }, set(v){ CFG = v; } });   // the script's own let binding
Object.assign(globalThis, { S, loadCfg, saveLocal, applySharedCfg, setWb: k => { WB_KEY = k; }, wbStoreKey });`, D);
  // the old one-for-every-workbook copy from before this change: Brendan-style clinic values on a shared computer
  store.eskCfg = JSON.stringify({ me: 'Old Person', clinic: 'SRL', pm: 'Old Person', cc: 'old@example.com', helpEmail: 'help-old@example.com', theme: 'dark' });
  D.setWb('uA'); store[D.wbStoreKey()] = JSON.stringify({ pick: { day: 'day46000', key: 'name:amy b.' }, used: ['name:amy b.'], users: { 'name:amy b.': { me: 'Amy B.', sig: 'Amy', clinic: 'SRL', cc: 'leak@example.com' } } });
  D.CFG = D.loadCfg(); D.applySharedCfg();
  assert.strictEqual(D.CFG.me, 'Amy B.'); assert.strictEqual(D.CFG.profKey, 'name:amy b.'); assert.strictEqual(D.CFG.theme, 'dark', 'theme is the only thing from the old copy');
  assert.strictEqual(D.CFG.clinic, '', 'clinic never from this computer'); assert.strictEqual(D.CFG.cc, ''); assert.strictEqual(D.CFG.pm, ''); assert.strictEqual(D.CFG.helpEmail, '');
  D.saveLocal(); assert(!('cc' in JSON.parse(store[D.wbStoreKey()]).users['name:amy b.']), 'saving drops clinic fields from this computer\'s copy');
  D.setWb('uB'); D.CFG = D.loadCfg(); D.applySharedCfg();
  assert.strictEqual(D.CFG.me, ''); assert.strictEqual(D.CFG.profKey, undefined, 'another workbook: nobody picked'); assert.strictEqual(D.CFG.sig, '');
  D.S.wbName = 'Castro ESK Tracker.xlsx'; D.applySharedCfg(); assert.strictEqual(D.CFG.clinic, 'CAS'); assert.strictEqual(D.CFG.locName, 'Castro');
  D.S.shared = { clinicCfg: { clinic: 'OAK', cc: 'desk@example.com' } }; D.applySharedCfg(); assert.strictEqual(D.CFG.clinic, 'OAK', 'the workbook wins over its name'); assert.strictEqual(D.CFG.cc, 'desk@example.com');
  D.setWb('uA'); store[D.wbStoreKey()] = JSON.stringify({ ...JSON.parse(store[D.wbStoreKey()]), pick: { day: 'day45999', key: 'name:amy b.' } });
  D.CFG = D.loadCfg(); assert.strictEqual(D.CFG.me, '', 'yesterday\'s pick: pick again today');
});

(async () => { for (const [name, fn] of tests){ await fn(); n++; console.log('ok', name); } console.log(n + ' passed'); })().catch(e => { console.error(e); process.exit(1); });
