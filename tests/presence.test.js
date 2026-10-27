// "Who's active": the pure parts of the presence list (initials, colors, who shows, idle, one entry per name), pulled from
// src/taskpane.html. The sheet I/O itself needs Excel.   node tests/presence.test.js   FICTIONAL names only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const between = (a, b) => { const i = src.indexOf(a), j = src.indexOf(b, i); assert(i > 0 && j > i, a); return src.slice(i, j); };
const C = {}; vm.createContext(C);
vm.runInContext(between('const PRES_BEAT_MS', 'function presWhereNow') + '\nObject.assign(globalThis, { presInitials, presColor, presAgo, presList, PRES_COLORS });', C);
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };
const T0 = Date.parse('2026-10-27T16:00:00Z'), iso = minAgo => new Date(T0 - minAgo * 60000).toISOString();
const row = (session, name, seenMin, inputMin, view) => ({ session, name, view: view || 'Today', where: 'Excel pane', seen: iso(seenMin), input: iso(inputMin) });

t('initials: first + last, titles and credentials dropped, one word gives two letters, no name gives ?', () => {
  assert.strictEqual(C.presInitials('Taylor R.'), 'TR'); assert.strictEqual(C.presInitials('Ritu Aujla'), 'RA');
  assert.strictEqual(C.presInitials('Dr. Pat Example, PMHNP'), 'PE'); assert.strictEqual(C.presInitials('casey'), 'CA');
  assert.strictEqual(C.presInitials(''), '?'); assert.strictEqual(C.presInitials('  '), '?');
});
t('colors: the same name always gets the same color, from the palette', () => {
  assert.strictEqual(C.presColor('Taylor R.'), C.presColor('taylor r.')); assert(C.PRES_COLORS.includes(C.presColor('Jordan P.')));
});
t('who shows: not me, only panels seen in the last 3½ min, active before idle, then by name', () => {
  const L = C.presList([row('me', 'Me Myself', 0, 0), row('a', 'Taylor R.', 0.5, 0.5), row('b', 'Jordan P.', 1, 2, 'Meds'), row('c', 'Casey L.', 1, 7),
                        row('d', 'Gone G.', 5, 5)], T0, 'me');
  assert.deepStrictEqual([...L.map(p => p.name)], ['Jordan P.', 'Taylor R.', 'Casey L.']);
  assert.deepStrictEqual([...L.map(p => p.idle)], [false, false, true]);
  assert.deepStrictEqual([...L.map(p => p.ago)], ['2 min ago', 'now', 'idle 7 min']);
  assert.strictEqual(L[0].view, 'Meds');
});
t('the same person on two computers shows once, with their latest activity; a blank name still shows', () => {
  const L = C.presList([row('a1', 'Taylor R.', 1, 4), row('a2', 'taylor r.', 0.2, 0.2), row('x', '', 0.5, 0.5)], T0, 'me');
  assert.strictEqual(L.length, 2); const tr = L.find(p => p.ini === 'TR'); assert.strictEqual(tr.ago, 'now');
  assert(L.some(p => p.ini === '?' && /no name/.test(p.name)));
});
t('bad or missing times never show someone, a missing input time falls back to the check-in', () => {
  const L = C.presList([{ session: 'z', name: 'No Time', seen: '', input: '' }, { session: 'y', name: 'Old Input', seen: iso(1), input: '' }], T0, 'me');
  assert.deepStrictEqual([...L.map(p => p.name + '|' + p.ago)], ['Old Input|1 min ago']);
});
console.log(n + ' passed');
