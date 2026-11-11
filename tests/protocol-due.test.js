// Protocol + consent renewal window (needsProto / protoDueFrom in src/taskpane.html).
//   node tests/protocol-due.test.js
// Due once expired, or from whichever comes first: 7 days before expiry, or the patient's last booked visit before expiry.
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane.html'), 'utf8');
const code = src.match(/const PROTO_DAYS = [\s\S]*?\nfunction needsProto\(p\)\{[^\n]*\n/)[0];
const EXP = 46376;   // a protocol exp date (serial); visits are given as days before it
function run(today, visits, p = {}){
  const A = visits.map(v => typeof v === 'object' ? { mrn: '1', active: true, ...v, date: EXP + v.d } : { mrn: '1', active: true, date: EXP + v });
  const C = { S: { A }, todaySerial: () => EXP + today, Math };
  vm.createContext(C); vm.runInContext(code + '\nthis.np = needsProto; this.from = protoDueFrom;', C);
  return C.np({ mrn: '1', status: 'Active', protoExp: EXP, ...p });
}
let n = 0; const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('last visit 10 days before expiry: due from that visit, not before', () => {
  assert.strictEqual(run(-11, [-10, 4]), false); assert.strictEqual(run(-10, [-10, 4]), true); assert.strictEqual(run(-9, [-10, 4]), true);
});
t('visits every few days: due from 7 days out, not 30', () => {
  assert.strictEqual(run(-20, [-7, -4, -2, 2]), false); assert.strictEqual(run(-8, [-7, -4, -2, 2]), false); assert.strictEqual(run(-7, [-7, -4, -2, 2]), true);
});
t('no visit booked before expiry: due from 7 days out', () => {
  assert.strictEqual(run(-8, [3]), false); assert.strictEqual(run(-7, [3]), true); assert.strictEqual(run(-8, []), false);
});
t('expired, or no protocol date: always due', () => {
  assert.strictEqual(run(1, []), true); assert.strictEqual(run(-100, [], { protoExp: null }), true);
});
t('a cancelled visit or another patient’s visit does not count', () => {
  assert.strictEqual(run(-12, [{ d: -12, active: false }]), false); assert.strictEqual(run(-12, [{ d: -12, mrn: '2' }]), false);
});
t('only active or pending patients', () => { assert.strictEqual(run(1, [], { status: 'Discharged' }), false); });
console.log(n + ' passed');
