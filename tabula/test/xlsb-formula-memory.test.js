import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeFormula, convertXlsb } from '../src/xlsb.js';

const u32 = (...values) => { const b = Buffer.alloc(values.length * 4); values.forEach((n, i) => b.writeUInt32LE(n >>> 0, i * 4)); return b; };
const u16 = n => { const b = Buffer.alloc(2); b.writeUInt16LE(n); return b; };
const f64 = n => { const b = Buffer.alloc(8); b.writeDoubleLE(n); return b; };
const wide = s => Buffer.concat([u32(s.length), Buffer.from(s, 'utf16le')]);
const variable = n => { const out = []; do { const b = n & 127; n >>>= 7; out.push(b | (n ? 128 : 0)); } while (n); return Buffer.from(out); };
const rec = (n, b = Buffer.alloc(0)) => Buffer.concat([variable(n), variable(b.length), b]);
const formula = (tokens, extra = Buffer.alloc(0)) => Buffer.concat([u32(tokens.length), tokens, u32(extra.length), extra]);
const ref = r => formula(Buffer.concat([Buffer.from([0x24]), u32(r), u16(0)]));
const consume = it => { for (;;) { const next = it.next(); if (next.done) return next.value; } };

test('distinct filled-down XLSB formula memo stays bounded without changing decoded values', () => {
  const memo = new Map(), env = { memo }, base = { r: 0, c: 0 };
  for (let r = 0; r < 100000; r++) assert.equal(decodeFormula(ref(r), 0, env, base).text, '$A$' + (r + 1));
  assert.ok(memo.size > 0 && memo.size <= 16384);
  for (const r of [0, 12345, 65535, 99999]) assert.deepEqual(decodeFormula(ref(r), 0, env, base), decodeFormula(ref(r), 0, {}, base));
  memo.clear(); assert.equal(decodeFormula(ref(100001), 0, env, base).text, '$A$100002'); assert.equal(memo.size, 1);
});

test('long memo keys obey the text budget as well as the entry budget', () => {
  const memo = new Map();
  for (let i = 0; i < 10000; i++) {
    const text = String(i).padStart(6, '0') + '가'.repeat(210);
    const source = formula(Buffer.concat([Buffer.from([0x17]), u16(text.length), Buffer.from(text, 'utf16le')]));
    assert.equal(decodeFormula(source, 0, { memo }, { r: 0, c: 0 }).text, '"' + text + '"');
  }
  let chars = 0; for (const [key, text] of memo) chars += key.length + text.length;
  assert.ok(chars <= 2 << 20); assert.ok(memo.size > 1000 && memo.size < 10000);
});

function sheet(shared = false) {
  const entries = [rec(129), rec(148, u32(0, 1, 0, 1)), rec(145)];
  for (let r = 0; r < 2; r++) {
    entries.push(rec(0, Buffer.concat([u32(r, 0), Buffer.from([44, 1, 0, 0])])));
    const f = shared ? formula(Buffer.concat([Buffer.from([1]), u32(0)]), u32(1)) : ref(r);
    entries.push(rec(9, Buffer.concat([u32(1, 0), f64(r + 1), u16(0), f])));
    if (shared && r === 0) entries.push(rec(427, Buffer.concat([u32(0, 1, 1, 1), formula(Buffer.concat([Buffer.from([0x2c]), u32(0), u16(0xffff)]))])));
  }
  entries.push(rec(146), rec(130)); return Buffer.concat(entries);
}
function workbook(shared = false) {
  const files = {
    'xl/workbook.bin': Buffer.concat([rec(156, Buffer.concat([u32(0, 1), wide('r1'), wide('First')])), rec(156, Buffer.concat([u32(0, 2), wide('r2'), wide('Second')]))]),
    'xl/_rels/workbook.bin.rels': new TextEncoder().encode('<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.bin"/></Relationships>'),
    'xl/worksheets/sheet1.bin': sheet(shared), 'xl/worksheets/sheet2.bin': sheet(shared)
  };
  consume(convertXlsb(files, { lazySheets: true })); return files;
}

test('completed or cancelled sheet iterators release their separate decode caches', () => {
  const files = workbook(), set = Map.prototype.set, memoMaps = new Set();
  Map.prototype.set = function(key, value) {
    if (typeof key === 'string' && key.charCodeAt(0) === 7 && typeof value === 'string' && value.startsWith('$A$')) memoMaps.add(this);
    return set.call(this, key, value);
  };
  try {
    for (let i = 1; i <= 2; i++) {
      const path = 'xl/worksheets/sheet' + i + '.bin'; files.__xlsb.prepareSheet(path);
      const iterator = files.__xlsb.rows.get(path)();
      const first = iterator.next(); assert.equal(first.value.cells[0].f, '$A$1');
      assert.ok([...memoMaps].some(m => m.size > 0));
      if (i === 1) { const second = iterator.next(); assert.equal(second.value.cells[0].f, '$A$2'); assert.equal(iterator.next().done, true); }
      else iterator.return();
      assert.ok([...memoMaps].every(m => m.size === 0));
    }
  } finally { Map.prototype.set = set; }
  assert.equal(memoMaps.size, 2);
});

test('shared XLSB references remain relative to each consuming cell across sheets', () => {
  const files = workbook(true);
  for (let i = 1; i <= 2; i++) {
    const path = 'xl/worksheets/sheet' + i + '.bin'; files.__xlsb.prepareSheet(path);
    const rows = [...files.__xlsb.rows.get(path)()];
    assert.deepEqual(rows.map(row => row.cells[0].f), ['A1', 'A2']);
  }
});
