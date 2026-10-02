import test from 'node:test';
import assert from 'node:assert/strict';
import { Axis } from '../src/axis.js';
import { nextVisibleAxisIndex, visibleAxisIndices } from '../src/axis-window.js';

const bits = (start, flags) => ({ start, __bits: Uint8Array.from(flags), count: flags.reduce((a, b) => a + b, 0) });
function scan(axis, first, last) { const out = []; for (let i = Math.max(0, Math.ceil(first)); i <= Math.min(axis.max - 1, Math.floor(last)); i++) if (axis.size(i)) out.push(i); return out; }
function next(axis, i, dir) { let j = i; while (j >= 0 && j < axis.max && axis.size(j) === 0) j += dir; return j < 0 || j >= axis.max ? i : j; }
function verify(axis) {
  for (let i = -1; i <= axis.max; i++) for (const dir of [-1, 1]) assert.equal(nextVisibleAxisIndex(axis, i, dir), next(axis, i, dir), 'next ' + i + '/' + dir);
  for (let first = -3; first < axis.max + 3; first += 7) for (const width of [-1, 0, 1, 15, 80]) assert.deepEqual(visibleAxisIndices(axis, first, first + width), scan(axis, first, first + width));
}

test('visible window preserves custom sizes, sparse hidden rows and boundaries', () => {
  verify(new Axis(23, { 0: 0, 1: 0, 8: 42.5, 11: 0, 12: 0, 31: 0 }, [{ 5: true, 6: true, 7: true, 30: true }], 32));
  assert.deepEqual(visibleAxisIndices(new Axis(20, {}, [], 0), 0, 10), []);
});

test('bitmap runs cross 256-entry prefix boundaries without skipping visible rows', () => {
  const flags = Array(800).fill(1); for (const i of [0, 255, 256, 511, 799]) flags[i] = 0;
  const axis = new Axis(20, { 4: 25 }, [bits(5, flags)], 810);
  verify(axis);
  assert.deepEqual(visibleAxisIndices(axis, 6, 803), [260, 261, 516]);
});

test('overlapping and interleaved hidden sources form a union for enumeration', () => {
  const a = Array.from({ length: 600 }, (_, i) => +(i % 7 !== 0));
  const b = Array.from({ length: 590 }, (_, i) => +(i % 11 !== 0));
  verify(new Axis(18, { 9: 0, 595: 0 }, [bits(0, a), bits(8, b), { 0: true, 14: true, 599: true }], 605));
});

test('all hidden, hidden tails and zero default dimensions stop without a loop', () => {
  verify(new Axis(20, {}, [bits(0, Array(90).fill(1))], 90));
  verify(new Axis(0, { 0: 20, 17: 25, 35: 30, 63: 20 }, [{ 35: true }], 64));
  verify(new Axis(0, {}, [], 40));
});

test('seeded randomized sparse sizes and multiple bitmaps match the linear reference', () => {
  let seed = 98271;
  const random = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
  for (let round = 0; round < 20; round++) {
    const size = {}, hidden = {}, n = 90;
    for (let i = 0; i < n; i++) { if (random() < .12) size[i] = random() < .6 ? 0 : 31.25; if (random() < .3) hidden[i] = true; }
    const a = bits(3, Array.from({ length: 82 }, () => +(random() < .65)));
    const b = bits(18, Array.from({ length: 55 }, () => +(random() < .35)));
    verify(new Axis(23, size, [hidden, a, b], n));
  }
});

test('million-row hidden run uses bounded size lookups and does not mutate sources', () => {
  const data = new Uint8Array(1_000_000); data.fill(1); data[0] = data[999999] = 0;
  const hidden = { start: 0, __bits: data, count: 999998 }, axis = new Axis(23, {}, [hidden], data.length);
  let calls = 0; const size = axis.size.bind(axis); axis.size = i => { calls++; return size(i); };
  assert.deepEqual(visibleAxisIndices(axis, 0, 999999), [0, 999999]);
  assert.ok(calls < 20, 'size lookups: ' + calls);
  assert.equal(nextVisibleAxisIndex(axis, 999998, -1), 0);
  assert.equal(data[123], 1); assert.equal(hidden.count, 999998);
});

test('new axes observe edited dimensions and hidden sources after layout or Undo', () => {
  const sizes = { 1: 0, 2: 0 }, hidden = { 5: true }, data = new Uint8Array(12); data.fill(1, 8, 11);
  const source = { start: 0, __bits: data, count: 3 };
  const first = new Axis(20, sizes, [hidden, source], 12);
  assert.deepEqual(visibleAxisIndices(first, 0, 11), [0, 3, 4, 6, 7, 11]);
  sizes[1] = 30; delete hidden[5]; data[9] = 0; source.count--;
  const second = new Axis(20, sizes, [hidden, source], 12);
  assert.deepEqual(visibleAxisIndices(second, 0, 11), [0, 1, 3, 4, 5, 6, 7, 9, 11]);
  const restored = new Axis(20, { 1: 0, 2: 0 }, [{ 5: true }, bits(8, [1, 1, 1])], 12);
  assert.deepEqual(visibleAxisIndices(restored, 0, 11), [0, 3, 4, 6, 7, 11]);
});

function verifyCoordinates(axis) {
  let y = 0;
  for (let i = 0; i < axis.max; i++) {
    assert.equal(axis.pos(i), y, 'position ' + i);
    const size = axis.size(i);
    if (size > 0) { assert.equal(axis.indexAt(y), i, 'edge hit ' + i); assert.equal(axis.indexAt(y + size / 2), i, 'middle hit ' + i); }
    y += size;
  }
  assert.equal(axis.pos(axis.max), y, 'end position');
}

test('overlapping filters subtract each hidden row once from position and hit-testing', () => {
  const a = bits(1, [1, 1]), b = bits(2, [1, 1]);
  const axis = new Axis(20, {}, [a, b], 8);
  assert.deepEqual(Array.from({ length: 9 }, (_, i) => axis.pos(i)), [0, 20, 20, 20, 20, 40, 60, 80, 100]);
  assert.equal(axis.indexAt(20), 4); assert.equal(axis.bits.length, 1);
  assert.deepEqual([...a.__bits], [1, 1]); assert.deepEqual([...b.__bits], [1, 1]);
  assert.deepEqual([a.start, a.count, b.start, b.count], [1, 2, 2, 2]);
  verifyCoordinates(axis);
});

test('overlap union preserves custom sizes, sparse hidden rows and input order', () => {
  const inputs = [bits(8, [1, 0, 1, 1]), { 4: true, 15: true }, bits(2, [1, 1, 0, 1, 0, 1, 1]), bits(1, [0, 1, 0, 1])];
  const before = structuredClone(inputs);
  const axis = new Axis(20, { 0: 33, 2: 100, 6: 17.25, 12: 0, 14: 31.5 }, inputs, 20);
  verifyCoordinates(axis); verify(axis); assert.deepEqual(inputs, before);
  assert.equal(axis.pos(6), 53); assert.equal(axis.size(2), 0);
});

test('disjoint sparse bitmaps never allocate the gap and keep their byte arrays', () => {
  const first = bits(1, [1, 1]), far = bits(100000000, [1, 1, 0]), close = bits(2, [1, 1]);
  const axis = new Axis(20, {}, [far, first, close], 100000010);
  assert.equal(axis.bits.length, 2);
  assert.equal(axis.bits.reduce((n, b) => n + b.data.length, 0), 6);
  assert.equal(axis.bits[1].data, far.__bits);
  assert.equal(axis.pos(100000003), (100000003 - 5) * 20);
  const touching = new Axis(20, {}, [bits(1, [1, 1]), bits(3, [1, 1])], 8);
  assert.equal(touching.bits.length, 2); verifyCoordinates(touching);
});

test('random overlapping filters agree with independent union coordinates and hit-tests', () => {
  let seed = 881783;
  const random = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
  for (let round = 0; round < 35; round++) {
    const n = 620, sources = [], sizes = {}, sparse = {};
    for (let k = 0; k < 4; k++) { const start = Math.floor(random() * 280), len = 40 + Math.floor(random() * 330); sources.push(bits(start, Array.from({ length: len }, () => +(random() < .58)))); }
    for (let i = 0; i < n; i++) { if (random() < .05) sizes[i] = random() < .4 ? 0 : 31.25; if (random() < .04) sparse[i] = true; }
    const axis = new Axis(20, sizes, [...sources, sparse], n); let y = 0;
    for (let i = 0; i < n; i++) {
      const hidden = !!sparse[i] || sources.some(b => i >= b.start && i < b.start + b.__bits.length && b.__bits[i-b.start]);
      const size = hidden ? 0 : sizes[i] ?? 20;
      assert.equal(axis.size(i), size); assert.equal(axis.pos(i), y);
      if (size > 0) assert.equal(axis.indexAt(y + size / 2), i);
      y += size;
    }
    assert.equal(axis.pos(n), y);
  }
});
