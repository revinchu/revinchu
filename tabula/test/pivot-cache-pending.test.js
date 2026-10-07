import test from 'node:test';
import assert from 'node:assert/strict';
import { PivotSnapshotBuilder, pivotSnapshotValue, restorePivotSnapshot } from '../src/pivot-cache-data.js';
import { jsonReplacer } from '../src/block.js';
const values = s => Array.from({ length: s.n }, (_, r) => s.header.map((_, c) => pivotSnapshotValue(s, r, c)));

test('59 numeric columns and a million uniform rows never allocate typed arrays beyond the small initial capacity', () => {
  const Original = globalThis.Float64Array, lengths = [];
  globalThis.Float64Array = new Proxy(Original, { construct(Target, args, NewTarget) { assert.ok(args[0] <= 1024, 'unexpected full-column allocation: ' + args[0]); lengths.push(args[0]); return Reflect.construct(Target, args, NewTarget); } });
  try {
    const n = 1000000, row = new Array(59).fill(0), b = new PivotSnapshotBuilder(row.map((_, c) => 'F' + c), n);
    for (let r = 0; r < n; r++) b.add(row);
    assert.equal(b.n, n); assert.equal(lengths.length, 59);
    assert.equal(b.columns.reduce((bytes, col) => bytes + col.num.byteLength, 0), 59 * 1024 * 8);
    const snapshot = b.finish(); assert.equal(snapshot.n, n);
    for (let c = 0; c < 59; c++) { assert.deepEqual(snapshot.columns[c].tail, { start: 0, value: 0 }); assert.equal(snapshot.columns[c].num.length, 0); assert.equal(pivotSnapshotValue(snapshot, n - 1, c), 0); }
    assert.equal(lengths.filter(n => n > 1024).length, 0);
  } finally { globalThis.Float64Array = Original; }
});

test('255/256/257 uniform values switch only at the threshold and preserve every logical value', () => {
  for (const n of [255, 256, 257]) {
    const b = new PivotSnapshotBuilder(['F'], 16); for (let r = 0; r < n; r++) b.add([0]);
    assert.equal(b.columns[0].pending, n >= 256);
    const s = b.finish(); assert.equal(s.n, n); assert.equal(s.columns[0].tail?.start, n >= 256 ? 0 : undefined);
    for (const r of [0, n - 1]) assert.equal(pivotSnapshotValue(s, r, 0), 0);
  }
});

test('growing changing columns does not expand uniform or missing columns', () => {
  const n = 9000, b = new PivotSnapshotBuilder(['Changing', 'Zero', 'Missing', 'Empty'], n);
  for (let r = 0; r < n; r++) b.add([r, 0, null, '']);
  assert.ok(b.columns[0].num.length >= n); assert.equal(b.columns[1].num.length, 1024); assert.equal(b.columns[2].num, null); assert.equal(b.columns[3].str.length, 1024);
  const s = b.finish(); assert.equal(s.columns[0].tail, undefined);
  for (const r of [0, 1023, 1024, n - 1]) assert.deepEqual(s.header.map((_, c) => pivotSnapshotValue(s, r, c)), [r, 0, null, '']);
});

test('pending numeric/missing/empty/boolean/error tails materialize on each later type change', () => {
  const b = new PivotSnapshotBuilder(['F'], 16), input = [];
  for (const value of [0, null, '', false, true, { error: '#N/A' }, 5]) for (let r = 0; r < 300; r++) { b.add([value]); input.push([value]); }
  const s = b.finish(); assert.deepEqual(s.columns[0].tail, { start: 1800, value: 5 }); assert.equal(s.columns[0].num.length, 1800); assert.equal(s.columns[0].str.length, 1800);
  assert.deepEqual(values(s), input);
  const copy = restorePivotSnapshot(JSON.parse(JSON.stringify(s, jsonReplacer))); assert.deepEqual(values(copy), input);
});

test('a different last value expands a million-row pending tail exactly rather than dropping earlier rows', () => {
  const n = 1000000, b = new PivotSnapshotBuilder(['F'], n);
  for (let r = 0; r < n - 1; r++) b.add([0]);
  assert.equal(b.columns[0].num.length, 1024); b.add([7]);
  const s = b.finish(); assert.equal(s.n, n); assert.equal(s.columns[0].tail, undefined); assert.equal(s.columns[0].num.length, n);
  for (const r of [0, 254, 255, 1024, n - 2]) assert.equal(pivotSnapshotValue(s, r, 0), 0);
  assert.equal(pivotSnapshotValue(s, n - 1, 0), 7);
});

test('a million missing rows allocate nothing until the late numeric value, preserving missing versus zero', () => {
  const n = 1000000, b = new PivotSnapshotBuilder(['F'], n);
  for (let r = 0; r < n - 1; r++) b.add([null]);
  assert.equal(b.columns[0].num, null); assert.equal(b.columns[0].str, null); b.add([7]);
  const s = b.finish(); assert.equal(s.n, n); assert.equal(s.columns[0].num.length, n);
  for (const r of [0, 1024, n - 2]) { assert.equal(pivotSnapshotValue(s, r, 0), null); assert.ok(Number.isNaN(s.columns[0].num[r])); }
  assert.equal(pivotSnapshotValue(s, n - 1, 0), 7);
});

test('short missing suffixes still pad dense arrays to n and long missing suffixes stay compressed', () => {
  for (const missing of [15, 255, 256, 3000]) {
    const b = new PivotSnapshotBuilder(['F'], 16); b.add([8]); for (let r = 0; r < missing; r++) b.add([undefined]);
    const s = b.finish(); assert.equal(s.n, missing + 1); assert.equal(pivotSnapshotValue(s, 0, 0), 8); assert.equal(pivotSnapshotValue(s, missing, 0), null);
    assert.equal(s.columns[0].num.length, missing >= 256 ? 1 : missing + 1); assert.equal(s.columns[0].tail?.value, missing >= 256 ? null : undefined);
  }
});

test('negative zero remains dense and transitions to positive zero without coercion; nonfinite errors do not flush valid pending values', () => {
  const b = new PivotSnapshotBuilder(['F'], 16);
  for (let r = 0; r < 300; r++) b.add([-0]); assert.equal(b.columns[0].pending, false);
  for (let r = 0; r < 300; r++) b.add([0]); assert.equal(b.columns[0].pending, true);
  for (const value of [NaN, Infinity, -Infinity]) { assert.throws(() => b.add([value]), /올바르지 않은 숫자/); assert.equal(b.n, 600); assert.equal(b.columns[0].pending, true); }
  const s = b.finish(); assert.deepEqual(s.columns[0].tail, { start: 300, value: 0 }); assert.ok(Object.is(pivotSnapshotValue(s, 299, 0), -0)); assert.ok(Object.is(pivotSnapshotValue(s, 300, 0), 0));
});

test('seeded changing runs across growth boundaries agree with retained input rows and JSON restoration', () => {
  let seed = 943723; const next = () => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return seed >>> 0; };
  const pool = [0, null, '', false, true, 'text', 1.25, { error: '#DIV/0!' }];
  for (let sample = 0; sample < 30; sample++) {
    const input = [], b = new PivotSnapshotBuilder(['A', 'B', 'C'], 16);
    for (let run = 0; run < 12; run++) { const row = Array.from({ length: 3 }, () => pool[next() % pool.length]), count = 1 + next() % 400;
      for (let r = 0; r < count; r++) { input.push(row); b.add(row); }
    }
    const s = b.finish(); assert.deepEqual(values(s), input); assert.deepEqual(values(restorePivotSnapshot(JSON.parse(JSON.stringify(s, jsonReplacer)))), input);
  }
});
