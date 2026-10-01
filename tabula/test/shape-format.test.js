import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shapeSizePatch, shapeArrowEnd, shapeGradientStops, shapeGradientStopPatch, addShapeGradientStop } from '../src/shape-format.js';
import { Workbook } from '../src/workbook.js';

test('비율 고정 치수는 비율을 유지하고 최대 크기에서도 유한한 값으로 제한', () => {
  assert.deepEqual(shapeSizePatch({ w: 240, h: 120, lockAspect: true }, 'w', 360), { w: 360, h: 180 });
  assert.deepEqual(shapeSizePatch({ w: 240, h: 120, lockAspect: true }, 'h', 6000), { h: 5000, w: 10000 });
  assert.deepEqual(shapeSizePatch({ w: 240, h: 120 }, 'w', 360), { w: 360 });
  assert.deepEqual(shapeSizePatch({ w: 240, h: 120, noMove: true }, 'w', 360), {});
  assert.deepEqual(shapeSizePatch({}, 'w', Infinity), {});
});
test('높이/너비 0인 선은 비율 고정 시 방향을 유지하고 0 나눗셈하지 않음', () => {
  assert.deepEqual(shapeSizePatch({ w: 240, h: 0, lockAspect: true }, 'w', 360), { w: 360, h: 0 });
  assert.deepEqual(shapeSizePatch({ w: 240, h: 0, lockAspect: true }, 'h', 10), { h: 0 });
  assert.deepEqual(shapeSizePatch({ w: 0, h: 240, lockAspect: true }, 'h', 360), { h: 360, w: 0 });
  assert.deepEqual(shapeSizePatch({ w: 0, h: 240, lockAspect: true }, 'w', 10), { w: 0 });
  assert.deepEqual(shapeSizePatch({ w: 240, h: 0, lockAspect: false }, 'h', 10), { h: 10 });
});
test('기존 양끝 화살표와 명시적인 시작/끝 없음 설정을 구별', () => {
  assert.equal(shapeArrowEnd({ arrow: 'both' }, 'headEnd').type, 'triangle');
  assert.equal(shapeArrowEnd({ arrow: 'both' }, 'tailEnd').type, 'triangle');
  assert.equal(shapeArrowEnd({ arrow: 'end' }, 'headEnd').type, 'none');
  assert.deepEqual(shapeArrowEnd({ arrow: 'both', headEnd: { type: 'none', w: 'lg', len: 'sm' } }, 'headEnd'), { type: 'none', w: 'lg', len: 'sm' });
});
test('그라데이션 중지점 편집은 주변 순서/색/투명도를 유지하고 원본을 바꾸지 않음', () => {
  const shape = { grad: { ang: 30, stops: [[0, '#000000', .1], [.4, '#123456', .5], [1, '#ffffff', .9]] } };
  const saved = structuredClone(shape);
  const changed = shapeGradientStopPatch(shape, 1, { position: .7, color: '#abcdef' });
  assert.deepEqual(changed.grad.stops[1], [.7, '#abcdef', .5]); assert.equal(changed.grad.ang, 30);
  assert.equal(shapeGradientStopPatch(shape, 1, { position: -1 }).grad.stops[1][0], 0);
  assert.equal(shapeGradientStopPatch(shape, 1, { position: 10 }).grad.stops[1][0], 1);
  assert.equal(shapeGradientStopPatch(shape, 1, { opacity: 0 }).grad.stops[1][2], 0);
  assert.deepEqual(shape, saved);
});
test('새 중지점은 가장 넓은 간격에 색/투명도를 보간하고 최대 16개로 제한', () => {
  const shape = { fill: '#000000', grad: { ang: 80, stops: [[0, '#000000', .2], [1, '#ffffff', .8]] } };
  let next = addShapeGradientStop(shape); assert.equal(next.index, 1); assert.deepEqual(next.grad.stops[1], [.5, '#808080', .5]);
  assert.equal(shape.grad.stops.length, 2); assert.equal(next.grad.ang, 80);
  for (let i = 0; i < 20; i++) next = addShapeGradientStop({ grad: next.grad });
  assert.equal(next.grad.stops.length, 16);
  const stops = shapeGradientStops(shape); stops[0][1] = '#ff0000'; assert.equal(shape.grad.stops[0][1], '#000000');
});
test('크기+비율과 선 끝 상세 패치를 한 실행 취소로 되돌리고 다시 적용', () => {
  const wb = new Workbook(); const original = { id: 'line', kind: 'line', x: 10, y: 20, w: 200, h: 0, lockAspect: true, stroke: '#4472c4', arrow: 'end' };
  wb.transact(() => wb.setSheetProp(0, 'shapes', [original]));
  wb.transact(() => wb.setSheetProp(0, 'shapes', [{ ...original, ...shapeSizePatch(original, 'w', 400), headEnd: { type: 'diamond', w: 'lg', len: 'sm' } }]));
  assert.equal(wb.sheets[0].shapes[0].w, 400); assert.equal(wb.sheets[0].shapes[0].h, 0);
  wb.undo(); assert.deepEqual(wb.sheets[0].shapes[0], original);
  wb.redo(); assert.equal(wb.sheets[0].shapes[0].headEnd.type, 'diamond'); assert.equal(wb.sheets[0].shapes[0].w, 400);
});
