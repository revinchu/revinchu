import test from 'node:test';
import assert from 'node:assert/strict';
import { RunColumn } from '../src/run-column.js';
import { columnKeysInRange } from '../src/column-keys-in-range.js';

const blank = Object.freeze({ raw: '', style: Object.freeze({ bold: true }) });
const isBlank = value => value === blank;
function counted(column) {
  const calls = { has: 0, keys: 0, visited: 0 };
  return { calls, get size() { return column.size; }, has(r) { calls.has++; return column.has(r); },
    *keys() { calls.keys++; for (const r of column.keys()) { calls.visited++; yield r; } } };
}
const sorted = rows => [...rows].sort((a, b) => a - b);

test('Map short area probes only inclusive occupied rows, irrespective of insertion order', () => {
  const map = new Map([[20, 'x'], [8, 'a'], [3, 'b'], [6, 'c'], [7, 'd'], [4, 'e'], [5, 'f']]);
  const col = counted(map);
  assert.deepEqual([...columnKeysInRange(col, 3, 8)], [3, 4, 5, 6, 7, 8]);
  assert.deepEqual(col.calls, { has: 6, keys: 0, visited: 0 });
  assert.deepEqual(sorted(columnKeysInRange(map, 3, 8)), sorted([...map.keys()].filter(r => r >= 3 && r <= 8)));
});

test('long sparse Map area visits only stored keys and retains insertion order', () => {
  const col = counted(new Map([[900000, 'a'], [17, 'b'], [3, 'c'], [1000001, 'd']]));
  assert.deepEqual([...columnKeysInRange(col, 0, 1000000)], [900000, 17, 3]);
  assert.deepEqual(col.calls, { has: 0, keys: 1, visited: 4 });
});

test('one-million-row compressed blank column checks five rows without detaching shared data', () => {
  const column = RunColumn.fromSortedStorage(isBlank, [[0, blank, 1000000]]);
  const snapshot = column.shareData(), key = column.dataKey;
  const col = counted(column);
  assert.deepEqual([...columnKeysInRange(col, 456789, 456793)], [456789, 456790, 456791, 456792, 456793]);
  assert.deepEqual(col.calls, { has: 5, keys: 0, visited: 0 });
  assert.equal(column.dataKey, key);
  assert.equal(snapshot.dataKey, key);
  assert.equal(column.runs.length, 1);
  assert.equal(column.size, 1000000);
});

test('short compressed range includes points, gaps and run endpoints without visiting outside', () => {
  const column = RunColumn.fromSortedStorage(isBlank, [[0, blank, 128], [140, { raw: 'value' }], [200, blank, 128]]);
  column.delete(126);
  const col = counted(column);
  assert.deepEqual([...columnKeysInRange(col, 125, 142)], [125, 127, 140]);
  assert.equal(col.calls.has, 18); assert.equal(col.calls.keys, 0);
});

test('long sparse compressed range filters occupied coordinates exactly', () => {
  const column = RunColumn.fromSortedStorage(isBlank, [[30, blank, 3], [500000, { raw: '=1', formula: true }], [1000001, blank]]);
  const col = counted(column);
  assert.deepEqual([...columnKeysInRange(col, 31, 1000000)], [31, 32, 500000]);
  assert.deepEqual(col.calls, { has: 0, keys: 1, visited: 5 });
});

test('area probes observe deletion and future insertion while preserving COW snapshots', () => {
  const column = RunColumn.fromSortedStorage(isBlank, [[0, blank, 128]]);
  const snapshot = column.shareData(), iterator = columnKeysInRange(column, 50, 53);
  assert.equal(iterator.next().value, 50);
  column.delete(51); column.set(52, { raw: 'edited' }); column.set(130, blank);
  assert.deepEqual([...iterator], [52, 53]);
  assert.equal(snapshot.has(51), true); assert.equal(snapshot.get(52), blank); assert.equal(snapshot.has(130), false);
  assert.equal(column.get(52).raw, 'edited');
});

test('sparse fallback retains live Map deletion and appended-key behavior', () => {
  const column = new Map([[1, 'a'], [2, 'b'], [30, 'outside']]);
  const iterator = columnKeysInRange(column, 0, 20);
  assert.equal(iterator.next().value, 1);
  column.delete(2); column.set(4, 'c'); column.set(40, 'outside');
  assert.deepEqual([...iterator], [4]);
});

test('compressed fallback with a live iterator and shareData sees append but snapshot stays fixed', () => {
  const column = RunColumn.fromSortedStorage(isBlank, [[0, blank, 128]]);
  const iterator = columnKeysInRange(column, 0, 500);
  assert.equal(iterator.next().value, 0);
  const snapshot = column.shareData(); column.set(128, blank);
  const rest = [...iterator];
  assert.equal(rest.length, 128); assert.equal(rest.at(-1), 128);
  assert.equal(snapshot.has(128), false); assert.equal(column.size, 129);
});

test('invalid/reversed areas and missing columns perform no probes', () => {
  const col = counted(new Map([[1, 'a']]));
  for (const [a, b] of [[5, 4], [NaN, 4], [0, Infinity], [0.5, 4]]) assert.deepEqual([...columnKeysInRange(col, a, b)], []);
  assert.deepEqual([...columnKeysInRange(null, 0, 4)], []);
  assert.deepEqual(col.calls, { has: 0, keys: 0, visited: 0 });
});

test('gone-cell removal is independent of enumeration order and leaves all other cells intact', () => {
  const source = [[9, 'outside'], [7, 'old-role'], [3, 'old-role'], [5, 'next-role'], [4, 'user-gap'], [6, 'old-role'], [2, 'outside']];
  const oldOrder = new Map(source), rangeOrder = new Map(source);
  const remove = (column, rows) => { for (const r of rows) if ([3, 6, 7].includes(r)) column.delete(r); };
  remove(oldOrder, [...oldOrder.keys()].filter(r => r >= 3 && r <= 7));
  remove(rangeOrder, [...columnKeysInRange(rangeOrder, 3, 7)]);
  assert.deepEqual([...rangeOrder], [...oldOrder]);
  assert.equal(rangeOrder.get(4), 'user-gap'); assert.equal(rangeOrder.get(5), 'next-role');
});
