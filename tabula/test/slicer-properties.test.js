import test from 'node:test';
import assert from 'node:assert/strict';
import { slicerSizePatch, slicerSourceKey } from '../src/slicer-properties.js';
const sl = { x: 96, y: 48, w: 192, h: 240, noMove: true };
const values = { x: '2.54', y: '1.27', w: '5.08', h: '6.35', placement: 'absolute', locked: false, noMove: false };
test('슬라이서 cm 치수와 위치 속성은 정확히 변환하고 원본을 보존한다', () => {
  const saved = structuredClone(sl), patch = slicerSizePatch(sl, values);
  assert.deepEqual(patch, { x: 96, y: 48, w: 192, h: 240, placement: 'absolute', locked: false, noMove: undefined }); assert.deepEqual(sl, saved);
  assert.equal(slicerSizePatch(sl, { ...values, w: '5.0801' }).w, 192);
});
test('잘못된 치수와 잠긴 이동은 조용히 보정하지 않고 거부한다', () => {
  for (const w of ['', 'NaN', '-1', '100000', '1']) assert.throws(() => slicerSizePatch(sl, { ...values, w }));
  assert.throws(() => slicerSizePatch(sl, { ...values, x: '5', noMove: true }), /해제/);
  assert.throws(() => slicerSizePatch(sl, { ...values, placement: 'unknown' }), /속성/);
  assert.equal(slicerSizePatch(sl, { ...values, noMove: true }).noMove, true);
});
test('보고서 연결 원본은 시트·표·범위가 일치해야 하며 rows 지연 getter를 읽지 않는다', () => {
  const source = { si: 0, ref: { r1: 0, c1: 0, r2: 3, c2: 2 }, get rows() { throw new Error('large allocation'); } };
  assert.equal(slicerSourceKey(source), slicerSourceKey({ si: 0, ref: { ...source.ref } }));
  assert.notEqual(slicerSourceKey(source), slicerSourceKey({ si: 1, ref: source.ref }));
  assert.notEqual(slicerSourceKey(source), slicerSourceKey({ si: 0, table: '표1', ref: source.ref }));
  assert.equal(slicerSourceKey({ si: -1, ref: source.ref }), null);
});
