import { test } from 'node:test';
import assert from 'node:assert/strict';
import { addPrintAreas, resizePrintArea } from '../src/print-area-edit.js';
const a = { r1: 0, c1: 0, r2: 9, c2: 3 };
test('인쇄영역에 이웃 사각형 추가 시 확장하고 떨어진 영역은 별도로 유지', () => {
  assert.deepEqual(addPrintAreas([a], { r1: 10, c1: 0, r2: 19, c2: 3 }), [{ ...a, r2: 19 }]);
  assert.deepEqual(addPrintAreas([a], { r1: 0, c1: 4, r2: 9, c2: 7 }), [{ ...a, c2: 7 }]);
  const distant = { r1: 0, c1: 8, r2: 9, c2: 9 };
  assert.deepEqual(addPrintAreas([a], distant), [a, distant]);
  assert.deepEqual(addPrintAreas([a], { r1: 2, c1: 1, r2: 3, c2: 2 }), [a]);
});
test('부분 접촉한 L자 영역은 빈 사각형을 추가로 인쇄하지 않음', () => {
  const half = { r1: 10, c1: 0, r2: 19, c2: 1 };
  assert.deepEqual(addPrintAreas([a], half), [a, half]);
});
test('영역 경계는 셀 사이에 붙고 반대편을 지나도 최소 한 셀 보존', () => {
  assert.deepEqual(resizePrintArea(a, 'bottom', 15), { ...a, r2: 14 });
  assert.deepEqual(resizePrintArea(a, 'right', 0), { ...a, c2: 0 });
  assert.deepEqual(resizePrintArea(a, 'top', 20), { ...a, r1: 9 });
  assert.deepEqual(resizePrintArea(a, 'left', 20), { ...a, c1: 3 });
  assert.deepEqual(a, { r1: 0, c1: 0, r2: 9, c2: 3 });
});
