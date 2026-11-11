// tblOrders (Appt List): Genoa / specialty "Mark ordered" logs the order so the Master tracker's Order Status formula shows it;
// receiving or clearing the kit's Inventory line updates the same row.
//   node tests/orders.test.js
// Runs the tblOrders helpers of src/taskpane.html in Node against a fake store. FICTIONAL MRNs and short names only.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const html = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8').replace(/\r\n/g, '\n');
const app = html.slice(html.indexOf("<script>\n'use strict';"), html.lastIndexOf('</script>'));
const BOUND = /^(async function |function |const |let |var |class |document\.|window\.|setInterval|\(function|\/\/ -{5,}|\/\/ ={5,}|Object\.assign\()/;
const CH = new Map(); { let cur = null;
  for (const l of app.split('\n')){
    if (BOUND.test(l)){ const m = /^(?:async )?function\s+([A-Za-z0-9_$]+)/.exec(l) || /^(?:const|let|var|class)\s+([A-Za-z0-9_$]+)/.exec(l); cur = m && !CH.has(m[1]) ? (CH.set(m[1], []), CH.get(m[1])) : null; }
    if (cur) cur.push(l);
  } }
const chunk = n => { assert(CH.has(n), 'top-level definition not found: ' + n); return CH.get(n).join('\n'); };
const NEED = ['norm', 'isF', 'blank', 'toSerial', 'mrnKey', 'num', 'Tbl', 'changesFor', 'mrnVal', 'lineMg',
  'ORD_OPEN', 'apptTimeVal', 'sameTimeVal', 'orderRows', 'ordersMarked', 'orderRowFor', 'orderRowSet', 'ordersReceived', 'ordersCleared', 'ordersSafe'];

const ser = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 864e5) + 25569, T = ser(2026, 11, 10);
const OH = ['MRN', 'Patient', 'Pharmacy', 'Dose mg', 'Needed By Date', 'Needed By Time', 'Order Date', 'Received Date', 'Status', 'Notes'];
const AH = ['MRN', 'Patient', 'Date', 'Time', 'Event', 'Status', 'Rendering', 'Dose', 'Lookup Key'];
function make({ orders = [], appts = [], withOrders = true, failOrders = false } = {}){
  const C = { console: failOrders ? { ...console, error(){} } : console }; vm.createContext(C);
  vm.runInContext(NEED.map(chunk).join('\n') + '\nObject.assign(globalThis, { Tbl, orderRows, ordersMarked, ordersReceived, ordersCleared, ordersSafe });', C);
  const tbl = (k, h, rows) => new C.Tbl(k, k, 'Appt List', h, rows, rows.map(r => r.map(() => '')));
  const tables = { appts: tbl('appts', AH, appts) }; if (withOrders) tables.orders = tbl('orders', OH, orders);
  const writes = [], toasts = [];
  const apply = (k, r, changes) => { const v = C.S.tables[k].values; if (r == null){ r = v.length; v.push(Array(C.S.tables[k].headers.length).fill('')); } for (const [c, x] of changes) v[r][c] = x && typeof x === 'object' ? x.date : x; return r; };
  const store = {
    async addMany(k, list){ if (failOrders && k === 'orders') throw new Error('locked'); for (const ch of list){ writes.push(['add', k, apply(k, null, ch)]); } },
    async updateMany(k, items){ for (const x of items) writes.push(['update', k, apply(k, x.r, x.changes)]); return items; },
    async update(k, r, snap, changes){ if (failOrders && k === 'orders') throw new Error('locked'); writes.push(['update', k, apply(k, r, changes)]); },
  };
  Object.assign(C, { S: { tables, store, byMrn: new Map([['900501', { mg: 84 }], ['900502', { mg: 56 }]]) }, CFG: { me: 'Tech A' },
    nm: m => 'ZZT' + String(m).slice(-3), fmtD: s => { const d = new Date((s - 25569) * 864e5); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`; },
    todaySerial: () => T, esc: s => String(s), toast: (k, m) => toasts.push([k, m]) });
  return { C, writes, toasts, row: r => Object.fromEntries(OH.map((h, i) => [h, C.S.tables.orders.values[r][i]])) };
}
const visit = (mrn, date, min, r) => ({ a: { mrn, date, time: min, r, key: `${mrn}|${date}|${min}` }, mg: 84, pharm: 'Genoa' });
let n = 0; const t = async (name, fn) => { await fn(); n++; console.log('ok', name); };

(async () => {
  await t('mark ordered adds an In Transit row with the order date and the visit date / time from Table47', async () => {
    const tv = 570 / 1440, x = make({ appts: [['900501', 'ZZT501', ser(2026, 11, 13), tv, 'Esketamine Follow Up', 'Expected', '', 84, '']] });
    assert.strictEqual(await x.C.ordersMarked([visit('900501', ser(2026, 11, 13), 570, 0)], T), 1);
    const o = x.row(0);
    assert.strictEqual(o.MRN, 900501); assert.strictEqual(o.Pharmacy, 'Genoa'); assert.strictEqual(o['Dose mg'], 84);
    assert.strictEqual(o['Needed By Date'], ser(2026, 11, 13)); assert.strictEqual(o['Needed By Time'], tv);   // same value the J formula's MINIFS returns
    assert.strictEqual(o['Order Date'], T); assert.strictEqual(o.Status, 'In Transit'); assert.match(o.Notes, /Marked ordered 11\/10 by Tech A/);
  });
  await t('a Pending / Cancelled row for the same visit is reused, not duplicated', async () => {
    const d = ser(2026, 11, 13), tv = 600 / 1440;
    const x = make({ orders: [[900501, 'ZZT501', 'Genoa', 84, d, tv, '', '', 'Pending', '']], appts: [['900501', '', d, tv, '', 'Expected', '', 84, '']] });
    await x.C.ordersMarked([visit('900501', d, 600, 0)], T);
    assert.strictEqual(x.C.S.tables.orders.values.length, 1); assert.strictEqual(x.row(0).Status, 'In Transit'); assert.strictEqual(x.row(0)['Order Date'], T);
  });
  await t('receiving the kit marks the row Received with the received date', async () => {
    const d = ser(2026, 11, 13);
    const x = make({ orders: [[900501, '', 'Genoa', 84, d, 0.4, T - 3, '', 'In Transit', ''], [900502, '', 'Genoa', 56, d, 0.4, T - 3, '', 'In Transit', '']] });
    assert.strictEqual(await x.C.ordersReceived({ mrn: '900501', ordered: T - 3, recv: 84 }, T), true);
    assert.strictEqual(x.row(0).Status, 'Received'); assert.strictEqual(x.row(0)['Received Date'], T); assert.strictEqual(x.row(1).Status, 'In Transit');
  });
  await t('two open orders: the one placed the same day as the Inventory line wins, else the visit needed soonest', async () => {
    const x = make({ orders: [[900501, '', 'Genoa', 84, ser(2026, 11, 20), 0.4, T - 1, '', 'In Transit', ''], [900501, '', 'Genoa', 84, ser(2026, 11, 13), 0.4, T - 5, '', 'Delayed', '']] });
    await x.C.ordersReceived({ mrn: '900501', ordered: T - 1, recv: 84 }, T);
    assert.strictEqual(x.row(0).Status, 'Received'); assert.strictEqual(x.row(1).Status, 'Delayed');
    await x.C.ordersReceived({ mrn: '900501', ordered: null, recv: 84 }, T);
    assert.strictEqual(x.row(1).Status, 'Received');
  });
  await t('clearing an order line marks its row Cancelled so Order Status falls back to Order Needed', async () => {
    const x = make({ orders: [[900501, '', 'Genoa', 84, ser(2026, 11, 13), 0.4, T - 2, '', 'In Transit', '']] });
    assert.strictEqual(await x.C.ordersCleared({ mrn: '900501', ordered: T - 2 }), true); assert.strictEqual(x.row(0).Status, 'Cancelled');
  });
  await t('no tblOrders in the workbook: nothing written, no error', async () => {
    const x = make({ withOrders: false });
    assert.strictEqual(await x.C.ordersMarked([visit('900501', T + 3, 600, 0)], T), 0);
    assert.strictEqual(await x.C.ordersReceived({ mrn: '900501', ordered: T }, T), false); assert.strictEqual(x.writes.length, 0);
  });
  await t('a failed tblOrders write warns instead of failing the Inventory action', async () => {
    const x = make({ failOrders: true });
    await x.C.ordersSafe(() => x.C.ordersMarked([visit('900501', T + 3, 600, 0)], T));
    assert.strictEqual(x.toasts.length, 1); assert.strictEqual(x.toasts[0][0], 'warn'); assert.match(x.toasts[0][1], /Inventory was updated/);
  });
  console.log(n + ' passed');
})().catch(e => { console.error(e); process.exit(1); });
