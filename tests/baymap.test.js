// Bay map on Today: bay matching and the colour each bay gets. Runs with the built-in drawing saved as the workbook's map
// (there is no fallback: no saved map = no map, tested at the end).
//   node tests/baymap.test.js
// Runs the bay map block of src/taskpane.html in Node with stubbed timers. FICTIONAL codes only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const a = src.indexOf('// ---------- bay map:'), b = src.indexOf("document.addEventListener('toggle', e => { const d = e.target; if (d.dataset?.sec === 'bmap'");
assert(a > 0 && b > a, 'bay map block markers');
const pa = src.indexOf('// Vitals ranges (Settings'), pb = src.indexOf('function prnStart(');   // vitOut, prnPend, prnAsk, prnAckOrders, techDue
assert(pa > 0 && pb > pa, 'PRN helper markers');
const C = { fin: new Set(), vit: {}, now: 650 }; vm.createContext(C);
vm.runInContext(`const obsAt = v => v?.adm ?? null, v40At = v => v.adm + 40, disAt = o => o + 100, fmtT = m => Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0');
const nowMin = () => now, S = {}, tdType = () => 'Follow-up', tdLead = r => '<p class="lead">lead ' + r.key + '</p>', trackHTML = () => '<div class="track"></div>', stepsHTML = () => '<ol class="steps"></ol>';
const tdPrimary = r => ({ act: 'vit', html: '<button data-act="vit" data-k="' + r.key + '">Record</button>' });
const vitGet = (k, w) => (vit[k] || {})[w] || (w === 'fin' ? (fin.has(k) ? { bp: '120/80', at: 700 } : null) : w === 'mid' ? (k.endsWith('m') ? { bp: '130/80', at: 650 } : null) : null);
const canonProv = s => String(s).trim(), PS = p => p !== 'BESSE', nm = m => m, esc = s => String(s), CFG = { bayMap: true, vitSysMax: 140, vitDiaMax: 90, vitHrMax: 110, vitHrMin: 50, vitO2Min: 92, vitEmSys: 180, vitEmDia: 120, vitalsMin: 40 };
const fromHHMM = t => { const m = /^(\\d+):(\\d+)$/.exec(t || ''); return m ? +m[1] * 60 + +m[2] : null; }, rkText = r => [r.bp && 'BP ' + r.bp, r.hr && 'HR ' + r.hr].filter(Boolean).join(' · ');
` + src.slice(pa, pb) + src.slice(a, b) + '\nObject.assign(globalThis, { S, bayNo, bayState, bayMapModel, bayMapHTML, bayPick, bayPanelHTML, bayNextHTML, bayProvPanelHTML, bayLabel, BP_DEFAULT, pvMatch, pvKits, prnPend, prnAsk, prnAckOrders, techDue, bayUntil, vitOut, vitEm, emReadings });', C);
C.S.shared = { bayPlan: C.BP_DEFAULT };
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
t('provider view: only their patients named; other bays "In use"; tag shows initial vitals or ready', () => {
  const mine = (key, bay, grp, rank) => { const r = row(key, bay, grp, rank); r.a.rendering = 'Dr A'; return r; };
  const other = (key, bay, grp, rank) => { const r = row(key, bay, grp, rank); r.a.rendering = 'Dr B'; return r; };
  C.vit.pa = { init: { bp: '132/84', hr: 80, at: 600 } }; C.fin.add('pr');
  const wo = other('wo', '1', 'wait', 9); wo.v = { bay: '1' }; const wm = mine('wm', '', 'wait', 9); wm.v = {};
  const m = { act: [1], need: [mine('pa', '2', 'need', 1), other('ob', '3', 'need', 0), mine('pr', '4', 'need', 3)], chair: [mine('pn', '5', 'chair', 4)], wait: [wo, wm] };
  const by = Object.fromEntries(C.bayMapModel(m, 'Dr A').bays.map(x => [x.id, x]));
  assert.strictEqual(by['3'].st, 'busy'); assert.strictEqual(by['3'].title, 'Another provider'); assert.strictEqual(by['3'].s, 'In use until ~11:45', 'until their final vitals (dose 10:05 + 100 min)'); assert.strictEqual(by['3'].row, null);
  assert.strictEqual(by['2'].s, 'Initial 132/84 · HR 80'); assert.strictEqual(by['2'].st, 'soon', 'colour still follows urgency');
  assert.strictEqual(by['4'].st, 'ready'); assert.strictEqual(by['4'].s, 'Ready to discharge');
  assert.strictEqual(by['5'].s, 'No initial vitals yet');
  const pc = mine('pc', '1', 'need', 2); pc.v.pc = { at: 605, sn: '1234' };
  assert.strictEqual(C.bayMapModel({ ...m, need: [pc] }, 'Dr A').bays[0].ts, 'Administered, the tech signs it');
  pc.v.prnOrd = [{ id: 'o1', med: 'Zofran', dose: '4 mg ODT', at: 610, ack: null }];
  assert.strictEqual(C.bayMapModel({ ...m, need: [pc] }, 'Dr A').bays[0].ts, 'PRN waiting for the PCT');
  assert.strictEqual(C.bayState(pc).s, 'PRN ordered · Zofran', 'tech map shows the open PRN');
  pc.v.prnOrd[0].ack = { by: 'T', time: '10:12' }; assert.notStrictEqual(C.bayState(pc).s, 'PRN ordered · Zofran', 'acknowledged: gone');
  assert.strictEqual(by['1'].res, null, 'another provider’s booking is not shown'); assert.strictEqual(by['1'].s, 'Open');
  const h = C.bayMapHTML(m, null, 'Dr A');
  assert(!/OB\b/.test(h) && !/WO\b/.test(h), 'no names of other providers’ patients anywhere');
  assert(/data-b="3" data-state="busy"/.test(h) && !/data-b="3" data-k=/.test(h), 'busy bay carries no visit key');
  assert(/<details class="card bmap noprint" data-sec="bmap" open>/.test(h) && /3 of your patients in a bay · 1 ready to discharge/.test(h));
  assert(/Still to come/.test(h) && /WM/.test(h) && !/Choose a bay/.test(h), 'their arrivals, nothing to change');
  assert.strictEqual((h.match(/class="lantern on"/g) || []).length, 4, 'busy bay lantern lit too'); C.fin.delete('pr');
});
t('provider panel: vitals lines, then Confirm dose / confirmed / administered', () => {
  const r = row('pa', '2', 'need', 1); r.a.rendering = 'Dr A'; r.i.adminKit = null; r.v.pc = null;
  const b = { id: '2', row: r, st: 'soon', s: 'Initial 132/84 · HR 80', ts: 'Vitals due 10:45' };
  let h = C.bayProvPanelHTML(b);
  assert(/Initial<\/span><span><b>132\/84 · HR 80<\/b>/.test(h) && /40-min<\/span><span class="muted">Due 10:45/.test(h) && /Final<\/span><span class="muted">Due 11:45/.test(h));
  assert(/<p class="lead soon">Vitals due 10:45/.test(h) && /data-act="pvDose" data-k="pa">Confirm dose/.test(h));
  r.v.pc = { at: 605, sn: '1234' }; h = C.bayProvPanelHTML(b); assert(/You confirmed the dose at 10:05 · kit …1234/.test(h) && !/pvDose/.test(h));
  r.v.pc.lot = 'LOT9876';
  r.i.adminKit = { serial: 'ABC1234' }; h = C.bayProvPanelHTML(b); assert(/Administered at 10:05 · kit …1234 · lot LOT9876. The tech adds their name to finish it./.test(h), 'provider wrote it; no tech yet');
  r.i.adminKit.tech = 'Pat T'; h = C.bayProvPanelHTML(b); assert(/Administered at 10:05 · kit …1234 · you confirmed it · Pat T/.test(h));
  assert(/data-act="pvPrn" data-k="pa">Indicate PRN/.test(h), 'roomed, not discharged: can indicate a PRN');
  r.v.prnOrd = [{ id: 'o1', med: 'Zofran', dose: '4 mg ODT', reason: 'nausea', at: 620, ack: null }, { id: 'o2', med: 'Nifedipine', dose: '10 mg PO', at: 625, ack: { by: 'Pat T', time: '10:30' } }];
  r.v.rechk = [{ bp: '150/95', at: '10:50', why: 'after nifedipine' }];
  h = C.bayProvPanelHTML(b);
  assert(/<b>Zofran 4 mg ODT<\/b> <span class="muted">nausea · ordered 10:20 · waiting for the PCT/.test(h));
  assert(/<b>Nifedipine 10 mg PO<\/b> <span class="muted">given 10:30 by Pat T/.test(h));
  assert(/Recheck<\/span><span><b>BP 150\/95<\/b> <span class="muted">10:50 · after nifedipine/.test(h), 'rechecks show to the provider');
  r.v.out = { at: 720 }; h = C.bayProvPanelHTML(b); assert(!/pvPrn/.test(h) && /Zofran/.test(h), 'discharged: list only');
  delete r.v.out; delete r.v.prnOrd; delete r.v.rechk;
  h = C.bayProvPanelHTML({ id: '3', row: null, busy: true }); assert(/in use/.test(h) && /Another provider’s patient is in this bay./.test(h) && !/pvDose/.test(h));
  h = C.bayProvPanelHTML({ id: '3', row: null, busy: true, until: 705 }); assert(h.includes('in this bay until about 11:45 (their final vitals).'));
  delete C.vit.pa;
});
t('dose confirmation: last 4 of the serial must match exactly one kit this patient can get, and the lot must match it', () => {
  C.S.K = [{ id: 'k1', serial: 'SN0001234', lot: 'AB12345', available: true, pharm: 'BESSE', mrn: '' }, { id: 'k2', serial: 'SN0005678', lot: 'X1', available: true, pharm: 'BESSE', mrn: 'P' },
    { id: 'k3', serial: 'SN9995678', lot: 'X1', available: true, pharm: 'BESSE', mrn: '' }, { id: 'k4', serial: 'SN0004321', lot: 'X1', available: true, pharm: 'BESSE', mrn: 'Q' }, { id: 'k5', serial: 'SN0009999', lot: 'X1', available: false, pharm: 'BESSE', mrn: '' },
    { id: 'k6', serial: 'SN0002468', lot: '', available: true, pharm: 'BESSE', mrn: 'P' }];
  const it = { a: { mrn: 'P' }, inf: { pharm: 'BESSE' }, adminKit: null, pulledKit: null, suggest: null };
  assert.strictEqual(C.pvMatch(it, '1234', 'AB12345').k.id, 'k1', 'full lot'); assert.strictEqual(C.pvMatch(it, ' 1234 ', '2345').k.id, 'k1', 'last 4 of the lot');
  assert.strictEqual(C.pvMatch(it, '1234', 'ab12345').k.id, 'k1', 'lot case does not matter');
  assert(/lot doesn’t match kit …1234/.test(C.pvMatch(it, '1234', '9999').err), 'wrong lot');
  assert(/Enter the lot number/.test(C.pvMatch(it, '1234', '').err) && /Enter the lot number/.test(C.pvMatch(it, '1234', '345').err), 'lot required, 4+');
  assert(/no lot number on the Inventory/.test(C.pvMatch(it, '2468', 'ABCD').err), 'kit row with no lot');
  assert(/2 kits end in 5678/.test(C.pvMatch(it, '5678', 'AAAA').err), 'ambiguous');
  assert(/No kit for P ends in 4321/.test(C.pvMatch(it, '4321', 'AAAA').err), 'labeled for someone else');
  assert(/No kit/.test(C.pvMatch(it, '9999', 'AAAA').err), 'not available'); assert(/last 4/.test(C.pvMatch(it, '12', 'AAAA').err));
  const gen = { a: { mrn: 'P' }, inf: { pharm: 'GENOA' }, adminKit: null, pulledKit: null, suggest: null };
  assert(/No kit/.test(C.pvMatch(gen, '1234', 'AB12345').err), 'specialty patient: Besse stock does not count');
  assert(/doesn’t match/.test(C.pvMatch({ ...it, adminKit: { serial: 'SN0001111', lot: 'AB12345' } }, '1234', '2345').err));
  assert.strictEqual(C.pvMatch({ ...it, adminKit: C.S.K[0] }, '1234', '2345').k.id, 'k1', 'tech already recorded it: still checks serial and lot');
  C.S.K = [];
});
t('PRN orders: open until a logged PRN row carries the order id; removing that row reopens it', () => {
  const O = [{ id: 'o1', med: 'Zofran', dose: '4 mg', ack: null }, { id: 'o2', med: 'Nifedipine', dose: '10 mg', ack: null }];
  assert.strictEqual(C.prnPend({ prnOrd: O }).length, 2); assert.strictEqual(C.prnPend({ prnOrd: O, out: { at: 700 } }).length, 0, 'discharged: nothing pending');
  let A = C.prnAckOrders(O, [{ med: 'Zofran', dose: '4 mg', time: '10:15', ord: 'o1' }, { med: 'Captopril', dose: '25 mg', time: '10:16' }], 'Pat T', 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(A[0].ack)), { by: 'Pat T', time: '10:15', ts: 1 }); assert.strictEqual(A[1].ack, null);
  A = C.prnAckOrders(A, [{ med: 'Zofran', dose: '4 mg', time: '10:20', ord: 'o1' }], 'Sam U', 2);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(A[0].ack)), { by: 'Pat T', time: '10:20', ts: 1 }, 'edited time follows; who acknowledged stays');
  assert.strictEqual(C.prnAckOrders(A, [], 'Sam U', 3)[0].ack, null, 'row removed: open again');
  assert.strictEqual(C.techDue({ tech: '' }), true); assert.strictEqual(C.techDue({ tech: 'Pat T' }), false); assert.strictEqual(C.techDue(null), false);
});
t('in use until: final vitals from the dose, else from room time; past due says so', () => {
  assert.strictEqual(C.bayUntil({ start: 600, adm: 605 }), 705); assert.strictEqual(C.bayUntil({ start: 600 }), 700); assert.strictEqual(C.bayUntil({ start: 600, adm: 605, out: { at: 700 } }), null);
  const o = row('ob', '3', 'need', 0); o.a.rendering = 'Dr B'; C.now = 720;
  assert.strictEqual(C.bayMapModel({ act: [1], need: [o], chair: [], wait: [] }, 'Dr A').bays[2].s, 'In use · discharge due'); C.now = 650;
});
t('vitals ranges: BP, pulse and O2 past the limits; fine readings and blanks are not flagged', () => {
  assert.deepStrictEqual([...C.vitOut({ bp: '132/84', hr: '78', o2: '98' })], []);
  assert.deepStrictEqual([...C.vitOut({ bp: '152/84', hr: '78', o2: '98' })], ['BP 152/84'], 'systolic');
  assert.deepStrictEqual([...C.vitOut({ bp: '130 / 95' })], ['BP 130/95'], 'diastolic, spaces');
  assert.deepStrictEqual([...C.vitOut({ bp: '140/90', hr: '110', o2: '92' })], [], 'at the limit is fine');
  assert.deepStrictEqual([...C.vitOut({ hr: '44', o2: '89' })], ['HR 44', 'O2 89%']);
  assert.deepStrictEqual([...C.vitOut({ bp: '', hr: '', o2: '' })], []); assert.deepStrictEqual([...C.vitOut(null)], []);
});
t('PRN requests: wait for the provider, then go to the PCT once approved; declined never does', () => {
  const r = { id: 'r1', med: 'Zofran', dose: '4 mg ODT', req: { by: 'Pat T', at: 620, ts: 1 }, ok: null, no: null, ack: null };
  assert.strictEqual(C.prnAsk({ prnOrd: [r] }).length, 1); assert.strictEqual(C.prnPend({ prnOrd: [r] }).length, 0, 'not the PCT’s yet');
  const ok = { ...r, ok: { by: 'Dr A', at: 622, ts: 2 } };
  assert.strictEqual(C.prnAsk({ prnOrd: [ok] }).length, 0); assert.strictEqual(C.prnPend({ prnOrd: [ok] }).length, 1, 'approved: PCT gives and logs');
  const no = { ...r, no: { by: 'Dr A', at: 622, ts: 2, why: 'recheck first' } };
  assert.strictEqual(C.prnAsk({ prnOrd: [no] }).length, 0); assert.strictEqual(C.prnPend({ prnOrd: [no] }).length, 0);
  assert.strictEqual(C.prnPend({ prnOrd: [{ id: 'o', med: 'Zofran', ack: null }] }).length, 1, 'provider-indicated: straight to the PCT');
});
t('provider panel: a PCT request shows Approve / Decline; the tag asks for approval first', () => {
  const p = row('pq', '2', 'need', 3); p.a.rendering = 'Dr A';
  p.v.prnOrd = [{ id: 'r1', med: 'Zofran', dose: '4 mg ODT', reason: 'nausea', by: 'Dr A', at: 620, ts: Date.now(), req: { by: 'Pat T', at: 620, ts: Date.now() }, ok: null, no: null, ack: null }];
  assert.strictEqual(C.bayMapModel({ act: [1], need: [p], chair: [], wait: [] }, 'Dr A').bays[1].ts, 'PRN approval needed');
  let h = C.bayProvPanelHTML({ id: '2', row: p, st: 'soon', s: '' });
  assert(/asked by Pat T at 10:20/.test(h) && /data-act="pvOk" data-k="pq" data-o="r1">Approve/.test(h) && /data-act="pvNo"/.test(h));
  p.v.prnOrd[0].no = { by: 'Dr A', at: 622, why: 'recheck first' }; h = C.bayProvPanelHTML({ id: '2', row: p, st: 'soon', s: '' });
  assert(/declined: recheck first/.test(h) && !/pvOk/.test(h));
});
t('emergency BP: at or above 180 systolic or 120 diastolic; each reading and value is its own alarm', () => {
  assert.strictEqual(C.vitEm({ bp: '179/119' }), ''); assert.strictEqual(C.vitEm({ bp: '180/100' }), 'BP 180/100'); assert.strictEqual(C.vitEm({ bp: '150/120' }), 'BP 150/120');
  assert.strictEqual(C.vitEm({ bp: '' }), ''); assert.strictEqual(C.vitEm(null), '');
  const L = C.emReadings('k', { vit: { init: { bp: '132/84', at: 600 }, mid: { bp: '186/112', at: 645 } }, rechk: [{ bp: '170/100', at: '10:50' }, { bp: '182/121', at: '11:00' }] });
  assert.deepStrictEqual([...L.map(e => e.rid + ' ' + e.reading)], ['mid|186/112 BP 186/112', 'rk1|182/121 BP 182/121']);
  assert.strictEqual(L[1].at, 660, 'recheck time read from HH:MM');
});
t('no visits today: no map', () => assert.strictEqual(C.bayMapHTML({ act: [], need: [], chair: [], wait: [] }), ''));
t('no bay map saved: no card on Today; bays show as typed; Still to come has nothing that needs the map', () => {
  C.S.shared = {};
  const w = row('w', 'Bay 5', 'wait', 9); w.v = { bay: 'Bay 5' }; const n0 = row('n', '', 'wait', 9, { v: {} });
  const m = { act: [1], need: [row('a', '2', 'need', 0)], chair: [row('c', 'Bay 3', 'chair', 4)], wait: [w, n0] };
  assert.strictEqual(C.bayMapHTML(m, null), '', 'no Bay map card'); assert.strictEqual(C.bayMapHTML(m, { key: 'a' }), '');
  assert.strictEqual(C.bayNo('2'), null); assert.strictEqual(C.bayLabel('Bay 3'), '3'); assert.strictEqual(C.bayLabel('2'), '2');
  assert.strictEqual(C.bayMapModel(m).bays.length, 0); assert.strictEqual(C.bayPick(C.bayMapModel(m), null), null);
  const h = C.bayNextHTML(m);
  assert(/Bay 5<\/span><span><\/span>/.test(h) && /No bay yet<\/span><span><\/span>/.test(h), h);
  assert(!/bayMove|Change bay|Choose a bay/.test(h), 'no map bay picker');
  C.S.shared = { bayPlan: C.BP_DEFAULT }; assert(/data-act="bayMove"/.test(C.bayNextHTML(m)), 'with a map the picker is back');
});
t('ESK Bays with no bay map: a notice, the provider’s own roomed patients as a list, and the same panel as on the map', () => {
  C.S.shared = {}; C.S.tSel = null;
  const mine = (key, bay, grp, rank) => { const r = row(key, bay, grp, rank); r.a.rendering = 'Dr A'; return r; };
  const other = (key, bay, grp, rank) => { const r = row(key, bay, grp, rank); r.a.rendering = 'Dr B'; return r; };
  C.vit.pa = { init: { bp: '132/84', hr: 80, at: 600 } }; C.fin.add('pr');
  const pa = mine('pa', 'Bay 4', 'need', 1); pa.i.adminKit = null; pa.v.pc = null;
  const pq = mine('pq', '', 'chair', 4); pq.v.prnOrd = [{ id: 'r1', med: 'Zofran', dose: '4 mg ODT', reason: 'nausea', at: 620, ts: 1, req: { by: 'Pat T', at: 620, ts: 1 }, ok: null, no: null, ack: null }];
  const wm = mine('wm', '7', 'wait', 9); wm.v = { bay: '7' };
  const m = { act: [1], need: [pa, other('ob', '3', 'need', 0), mine('pr', '2', 'need', 3)], chair: [pq], wait: [other('wo', '1', 'wait', 9), wm] };
  let h = C.bayMapHTML(m, null, 'Dr A');
  assert(h.includes('<div class="msg info">No bay map in this workbook yet. A practice lead can load one in Settings › Bay map.</div>'), h);
  assert(/<summary>Your patients <span class="muted">3 of your patients roomed · 1 ready to discharge<\/span>/.test(h) && !/class="bplan"/.test(h));
  const items = [...h.matchAll(/<button class="bprow" data-act="bSel" data-k="(\w+)" data-state="(\w+)" aria-current="(\w+)"><span class="p">(\w+)<\/span><span class="where( none)?">([^<]*)<\/span><span class="s"><span class="dot"><\/span>([^<]*)<\/span>/g)].map(x => x.slice(1).map(y => y ?? '').join('|'));
  assert.deepStrictEqual(items, ['pa|soon|true|PA||Bay 4|Initial 132/84 · HR 80', 'pr|ready|false|PR||Bay 2|Ready to discharge', 'pq|track|false|PQ| none|No bay|No initial vitals yet'], 'urgent first; bay as typed; the tag text');
  assert(!/OB\b|WO\b/.test(h), 'other providers’ patients are not named');
  assert(/class="bpanel" aria-label="Bay 4"/.test(h) && /data-act="pvDose" data-k="pa">Confirm dose/.test(h) && /Initial<\/span><span><b>132\/84 · HR 80<\/b>/.test(h), 'same provider panel: vitals and Confirm dose');
  assert(/Still to come/.test(h) && /WM/.test(h) && /Bay 7/.test(h) && !/bayMove/.test(h));
  C.S.tSel = 'pq'; h = C.bayMapHTML(m, null, 'Dr A');
  assert(/data-k="pq" data-state="track" aria-current="true"/.test(h) && /class="bpanel" aria-label="No bay"/.test(h), 'picked from the list');
  assert(/data-act="pvOk" data-k="pq" data-o="r1">Approve/.test(h) && /data-act="pvNo" data-k="pq"/.test(h) && /data-act="pvPrn" data-k="pq">Indicate PRN/.test(h), 'PRN approve / decline and indicate');
  h = C.bayMapHTML({ act: [1], need: [other('ob', '3', 'need', 0)], chair: [], wait: [] }, null, 'Dr A');
  assert(/None of your patients are roomed yet/.test(h) && !/bpanel/.test(h) && !/OB\b/.test(h));
  assert.strictEqual(C.bayMapHTML(m, null), '', 'Today: still no card');
  C.S.tSel = null; C.fin.delete('pr'); delete C.vit.pa; C.S.shared = { bayPlan: C.BP_DEFAULT };
});
console.log(n + ' passed');
