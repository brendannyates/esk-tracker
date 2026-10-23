// Bay map on Today: bay matching and the colour each bay gets.
//   node tests/baymap.test.js
// Runs the bay map block of src/taskpane.html in Node with stubbed timers. FICTIONAL codes only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const a = src.indexOf('// ---------- bay map:'), b = src.indexOf("document.addEventListener('toggle', e => { const d = e.target; if (d.dataset?.sec === 'bmap'");
assert(a > 0 && b > a, 'bay map block markers');
const C = { fin: new Set() }; vm.createContext(C);
vm.runInContext(`const obsAt = v => v?.adm ?? null, v40At = v => v.adm + 40, disAt = o => o + 100, fmtT = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');
const S = {}, tdType = () => 'Follow-up', tdLead = r => '<p class="lead">lead ' + r.key + '</p>', trackHTML = () => '<div class="track"></div>', stepsHTML = () => '<ol class="steps"></ol>';
const tdPrimary = r => ({ act: 'vit', html: '<button data-act="vit" data-k="' + r.key + '">Record</button>' });
const vitGet = (k, w) => w === 'fin' ? fin.has(k) : w === 'mid' ? k.endsWith('m') : null, nm = m => m, esc = s => String(s), CFG = { bayMap: true };
` + src.slice(a, b) + '\nObject.assign(globalThis, { S, bayNo, bayState, bayMapModel, bayMapHTML, bayPick, bayPanelHTML, bayNextHTML });', C);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const row = (key, bay, grp, rank, extra = {}) => ({ key, grp, rank, a: { mrn: key.toUpperCase(), time: 600, active: true }, i: { adminKit: {}, inf: { mg: 84 } }, v: { start: 600, adm: 605, bay }, ...extra });

t('bay numbers: "4", "Bay 4", " bay 4 " match; 7 and blank do not', () => {
  assert.strictEqual(C.bayNo('4'), '4'); assert.strictEqual(C.bayNo('Bay 4'), '4'); assert.strictEqual(C.bayNo(' bay 4 '), '4');
  assert.strictEqual(C.bayNo('7'), null); assert.strictEqual(C.bayNo(''), null); assert.strictEqual(C.bayNo(undefined), null);
});
t('states follow the Today ranks', () => {
  assert.strictEqual(C.bayState(row('a', '1', 'need', 0)).st, 'late');
  assert.strictEqual(C.bayState(row('b', '1', 'need', 1)).st, 'soon');
  assert.strictEqual(C.bayState(row('c', '1', 'need', 2, { i: { adminKit: null } })).s, 'Dose not given');
  assert.strictEqual(C.bayState(row('d', '1', 'need', 3)).s, 'Final vitals due');
  C.fin.add('e'); assert.strictEqual(C.bayState(row('e', '1', 'need', 3)).st, 'ready');
  assert.strictEqual(C.bayState(row('fm', '1', 'chair', 4)).s, 'Final vitals 11:45');
  assert.strictEqual(C.bayState(row('g', '1', 'chair', 4)).s, 'Vitals 10:45');
});
t('model: occupied, reserved, open and off-map', () => {
  const w = row('w', 'Bay 5', 'wait', 9); w.a.time = 690; w.v = { bay: 'Bay 5' };
  const m = { act: [1], need: [row('a', '2', 'need', 0), row('x', '9', 'need', 1)], chair: [row('c', '3', 'chair', 4), row('nb', '', 'chair', 4)], wait: [w] };
  const M = C.bayMapModel(m), by = Object.fromEntries(M.bays.map(x => [x.id, x]));
  assert.strictEqual(by['2'].st, 'late'); assert.strictEqual(by['3'].st, 'track'); assert.strictEqual(by['1'].s, 'Nobody booked');
  assert.strictEqual(by['5'].st, 'open'); assert.strictEqual(by['5'].s, 'W at 11:30');
  assert.strictEqual(M.off.map(r => r.key).join(), 'x,nb'); assert.strictEqual(M.used, 2); assert.strictEqual(M.late, 1);
  assert.strictEqual(by['2'].r, 60, 'radius survives (row kept separately)');
  const h = C.bayMapHTML(m, { key: 'a' });
  assert(!/NaN/.test(h), 'no NaN positions'); assert(/data-k="a" data-state="late" aria-current="true"/.test(h));
  assert(/data-k="w" data-state="open"/.test(h), 'reserved bay opens that visit'); assert(/Not on the map:/.test(h));
});
t('selected bay: tapped bay, else the Today visit, else most urgent; panel beside the map', () => {
  const w = row('w', '5', 'wait', 9); w.a.time = 690; w.v = { bay: '5' };
  const m = { act: [1], need: [row('b', '4', 'need', 1), row('a', '2', 'need', 0)], chair: [row('c', '3', 'chair', 4)], wait: [w, row('n', '', 'wait', 9, { v: {} })] };
  const M = C.bayMapModel(m);
  assert.strictEqual(C.bayPick(M, null).id, '2', 'overdue bay first');
  assert.strictEqual(C.bayPick(M, { key: 'c' }).id, '3', 'follows the visit picked on Today');
  C.S.bSel = '5'; assert.strictEqual(C.bayPick(M, { key: 'c' }).id, '5', 'a tapped bay wins'); C.S.bSel = null;
  let h = C.bayMapHTML(m, null);
  assert(/class="bpanel" aria-label="Bay 2"/.test(h) && /lead a/.test(h) && /data-act="vit" data-k="a"/.test(h) && /data-act="bayMove" data-k="a"/.test(h) && /data-act="bayNote" data-k="a"/.test(h), 'visit panel: lead, next step, move, note');
  assert(/data-b="2" data-k="a" data-state="late" aria-current="true"/.test(h), 'map marks the selected bay');
  C.S.bSel = '5'; h = C.bayMapHTML(m, null);
  assert(/W is booked here at 11:30/.test(h) && /data-act="baySeat" data-k="w" data-b="5">Seat W in Bay 5/.test(h) && /data-mode="book">Book someone else here/.test(h), 'reserved open bay: seat or book someone else');
  C.S.bSel = '1'; h = C.bayMapHTML(m, null); assert(/Nobody is booked here yet/.test(h) && /data-mode="seat">Seat a patient here/.test(h), 'empty bay: seat a patient'); C.S.bSel = null;
  assert(/Still to come/.test(h) && /Bay 5<\/span><button[^>]*data-k="w">Change bay/.test(h) && /No bay yet<\/span><button[^>]*data-k="n">Choose a bay/.test(h), 'still to come with change / choose');
});
t('curtains draw for occupied bays they screen; lanterns lit only where someone is in the chair', () => {
  const w = row('w', '5', 'wait', 9); w.v = { bay: '5' };
  const h = C.bayMapHTML({ act: [1], need: [row('a', '2', 'need', 0)], chair: [], wait: [w] }, null);
  const cur = [...h.matchAll(/<g class="curtain">.*?<\/g>/g)].map(x => /class="fabric"/.test(x[0]) ? 'drawn' : 'bunched');
  assert.deepStrictEqual(cur, ['drawn', 'drawn', 'bunched', 'bunched'], 'bay 2 draws tracks 1-2 and 2-3; a booking alone draws nothing');
  assert.strictEqual((h.match(/class="lantern on"/g) || []).length, 1); assert.strictEqual((h.match(/class="lantern"/g) || []).length, 4);
  assert(/class="lantern on" transform="translate\(150 84\)"/.test(h) && /class="glow" cx="150" cy="84"/.test(h), 'bay 2 lantern lit with its glow');
});
t('no visits today: no map', () => assert.strictEqual(C.bayMapHTML({ act: [], need: [], chair: [], wait: [] }), ''));
console.log(n + ' passed');
