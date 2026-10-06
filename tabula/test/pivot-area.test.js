import test from 'node:test';
import assert from 'node:assert/strict';
import { pivotOutputArea, pivotAreaConflict } from '../src/pivot-area.js';

const point = (r, c) => ({ r1: r, c1: c, r2: r, c2: c });

test('피벗 결과 범위에는 필터와 빈 구분 행 및 가장 긴 행이 포함된다', () => {
  const grid = [[{}, {}], [], [{}, {}, {}, {}], [{}, {}]];
  assert.deepEqual(pivotOutputArea(grid, { top: 3, left: 5 }), { r1: 3, c1: 5, r2: 6, c2: 8 });
  assert.deepEqual(pivotOutputArea([[{}], [{}, {}]], {}), { r1: 0, c1: 0, r2: 1, c2: 1 });
});

test('빈 피벗 결과도 시작 셀을 예약한다', () => {
  for (const grid of [[], [[]], [[], []]]) {
    assert.deepEqual(pivotOutputArea(grid, { top: 10, left: 7 }), { r1: 10, c1: 7, r2: 10 + Math.max(0, grid.length - 1), c2: 7 });
  }
});

test('피벗의 경계 셀과 포함 관계는 충돌하며 인접 셀은 충돌하지 않는다', () => {
  const def = { name: '다른 피벗', area: { r1: 2, c1: 3, r2: 6, c2: 5 } };
  for (const area of [point(2, 3), point(6, 5), point(4, 4), { r1: 0, c1: 0, r2: 10, c2: 10 }, { r1: 6, c1: 5, r2: 9, c2: 8 }]) {
    assert.equal(pivotAreaConflict(area, [def]), def);
  }
  for (const area of [point(1, 4), point(7, 4), point(4, 2), point(4, 6), { r1: 0, c1: 6, r2: 10, c2: 9 }]) {
    assert.equal(pivotAreaConflict(area, [def]), null);
  }
});

test('교체 정의는 객체 동일성으로만 제외하며 같은 이름의 다른 피벗은 보호한다', () => {
  const old = { name: '피벗1', area: point(2, 3) }, copy = structuredClone(old);
  assert.equal(pivotAreaConflict(point(2, 3), [old], old), null);
  assert.equal(pivotAreaConflict(point(2, 3), [old, copy], old), copy);
  assert.equal(pivotAreaConflict(point(2, 3), [old], copy), old);
});

test('출력 영역이 없는 피벗은 명시된 유효한 시작 셀만 예약한다', () => {
  const pending = { name: '새 피벗', top: 4, left: 6 };
  assert.equal(pivotAreaConflict(point(4, 6), [pending]), pending);
  assert.equal(pivotAreaConflict(point(4, 7), [pending]), null);
  for (const def of [null, {}, { top: 4 }, { top: -1, left: 6 }, { top: NaN, left: 6 }, { top: 4, left: .5 }, { area: { r1: 2, c1: 3, r2: 1, c2: 4 } }]) {
    assert.equal(pivotAreaConflict(point(0, 0), [def]), null);
  }
  const damaged = { top: 4, left: 6, area: { r1: 4, c1: 6, r2: NaN, c2: 6 } };
  assert.equal(pivotAreaConflict(point(4, 6), [damaged]), damaged);
});

test('검사할 영역이 잘못되면 충돌을 보고하지 않는다', () => {
  for (const area of [null, {}, { ...point(1, 1), r2: 0 }, { ...point(1, 1), c1: -1 }, { ...point(1, 1), r2: Infinity }]) {
    assert.equal(pivotAreaConflict(area, [{ area: point(1, 1) }]), null);
  }
});

test('일괄 사전 검사는 각 피벗의 새 범위끼리도 비교할 수 있다', () => {
  const oldA = { name: 'A', area: point(0, 0) }, oldB = { name: 'B', area: point(0, 3) };
  const nextA = { ...oldA, area: pivotOutputArea([[{}, {}, {}, {}]], oldA) };
  const nextB = { ...oldB, area: pivotOutputArea([[{}]], { top: 0, left: 3 }) };
  assert.equal(pivotAreaConflict(nextA.area, [nextA, nextB], nextA), nextB);
  assert.equal(pivotAreaConflict(nextA.area, [nextA, { ...nextB, area: point(0, 4) }], nextA), null);
});

test('범위 계산과 충돌 검사는 입력 정의와 결과를 변경하지 않는다', () => {
  const grid = Object.freeze([Object.freeze([{}, {}]), Object.freeze([])]);
  const def = Object.freeze({ top: 8, left: 4, area: Object.freeze(point(8, 4)) });
  const before = structuredClone({ grid, def });
  const area = pivotOutputArea(grid, def);
  assert.equal(pivotAreaConflict(Object.freeze(area), Object.freeze([def])), def);
  assert.deepEqual({ grid, def }, before);
});
