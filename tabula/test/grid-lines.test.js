import { test } from 'node:test';
import assert from 'node:assert/strict';
import { gridLineWidth, resolveGridBorders } from '../src/grid-lines.js';
const edge = (overrides = {}) => ({ vertical: false, at: 30, start: 0, end: 60, width: 1, pattern: 'solid', color: '#000', ...overrides });

test('축소/DPR마다 테두리는 정수 device pixel 폭이며 CSS 최소폭으로 과장되지 않음', () => {
  for (const zoom of [.25, .33, .5, .75, 1, 1.25]) for (const dpr of [1, 1.25, 2]) for (const width of [1, 2, 3]) {
    const px = gridLineWidth(width, zoom, dpr) * zoom * dpr;
    assert.ok(Math.abs(px - Math.max(1, Math.round(width * zoom * dpr))) < 1e-8);
  }
  assert.equal(gridLineWidth(1, .5, 2), 1, '50% DPR2의 가는 선은 실제 1 device pixel');
  assert.equal(gridLineWidth(3, .5, 2), 3, '굵은 하단 선의 차이를 유지');
  assert.equal(gridLineWidth(3, .25, 1, 'double'), 12, '이중선은 선·간격·선의 최소 3 device pixel 유지');
});

test('서로 맞닿은 두 셀의 같은 경계는 중복 없이 한 번만 렌더', () => {
  const line = edge();
  assert.deepEqual(resolveGridBorders([line, { ...line }]), [line]);
  assert.deepEqual(line, edge(), '입력 경계 객체를 변경하지 않음');
});

test('병합 셀의 긴 경계와 일부 구간의 굵은 선을 분리하고 겹침 제거', () => {
  const plain = edge({ end: 120 });
  const thick = edge({ start: 40, end: 80, width: 3, color: '#f00' });
  assert.deepEqual(resolveGridBorders([plain, thick]), [edge({ end: 40 }), thick, edge({ start: 80, end: 120 })]);
});

test('이중선의 빈 간격이 이웃의 가는 선으로 채워지지 않음', () => {
  const double = edge({ pattern: 'double', width: 3 });
  assert.deepEqual(resolveGridBorders([edge(), double]), [double]);
  assert.deepEqual(resolveGridBorders([edge({ width: 3 }), double]), [double]);
});

test('축이 다른 경계·숨김으로 폭이 없는 경계·연속 같은 선을 구분', () => {
  const a = edge({ end: 30 }), b = edge({ start: 30 });
  const vertical = edge({ vertical: true });
  assert.deepEqual(resolveGridBorders([a, b, vertical, edge({ start: 5, end: 5, width: 3 })]), [edge(), vertical]);
});
