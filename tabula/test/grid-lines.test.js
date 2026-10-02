import { test } from 'node:test';
import assert from 'node:assert/strict';
import { gridLineWidth, resolveGridBorders, gridBorderPaintOrder } from '../src/grid-lines.js';
const edge = (overrides = {}) => ({ vertical: false, at: 30, start: 0, end: 60, width: 1, pattern: 'solid', color: '#000', ...overrides });

test('thin은 배율·DPR와 무관하게 화면 1픽셀이고 굵은 선의 두께 차이는 유지', () => {
  for (const zoom of [.25, .33, .5, .75, .8, 1, 1.25]) for (const dpr of [1, 1.25, 1.5, 2]) for (const width of [1, 2, 3]) {
    const px = gridLineWidth(width, zoom, dpr) * zoom * dpr;
    assert.ok(Math.abs(px - (width === 1 ? 1 : Math.max(1, Math.round(width * zoom * dpr)))) < 1e-8);
  }
  assert.equal(gridLineWidth(1, 1, 2), .5, 'DPR2/100%에서도 thin을 2픽셀로 늘리지 않음');
  assert.equal(gridLineWidth(1, 1, 1.5), 2 / 3, '150% 화면 배율에서도 1픽셀');
  assert.equal(gridLineWidth(1, .5, 2), 1, '50% DPR2의 가는 선은 실제 1 device pixel');
  assert.equal(gridLineWidth(3, .5, 2), 3, '굵은 하단 선의 차이를 유지');
  assert.equal(gridLineWidth(3, .25, 1, 'double'), 12, '이중선은 선·간격·선의 최소 3 device pixel 유지');
});

test('다른 축의 같은 두께 선은 밝은 선부터 그려 검정 교차점을 연속으로 유지', () => {
  const vertical = edge({ vertical: true, color: '#000000' });
  const horizontal = edge({ color: '#D0CECE' });
  for (const input of [[vertical, horizontal], [horizontal, vertical]]) {
    const resolved = resolveGridBorders(input);
    assert.deepEqual(gridBorderPaintOrder(resolved), [horizontal, vertical]);
  }
});

test('도장 순서는 두께, 같은 두께의 이중선, 색 밝기 순이며 원본 경계는 불변', () => {
  const thin = edge(), thick = edge({ width: 3, color: '#fff' });
  const double = edge({ width: 3, pattern: 'double', color: '#eee' });
  const blackThick = edge({ width: 3 });
  const input = Object.freeze([Object.freeze(double), Object.freeze(blackThick), Object.freeze(thick), Object.freeze(thin)]);
  const out = gridBorderPaintOrder(input);
  assert.deepEqual(out, [thin, thick, blackThick, double]);
  assert.notEqual(out, input); assert.equal(out[0], thin, '좌표나 색을 바꾸지 않고 참조만 정렬');
  assert.deepEqual(input, [double, blackThick, thick, thin]);
});

test('공유 경계의 같은 두께 충돌 선택 규칙을 도장 정렬이 변경하지 않음', () => {
  const gray = edge({ color: '#ddd' }), black = edge();
  assert.deepEqual(gridBorderPaintOrder(resolveGridBorders([gray, black])), [gray]);
  assert.deepEqual(gridBorderPaintOrder(resolveGridBorders([black, gray])), [black]);
});

test('CSS 색 표기·투명도·해석 불가 값도 결정적인 도장 순서를 가짐', () => {
  const colors = ['black', '#000', '#000000', 'rgb(0, 0, 0)', 'rgb(0% 0% 0% / 100%)'];
  for (const color of colors) {
    const dark = edge({ color }), light = edge({ color: '#d0cece' });
    assert.deepEqual(gridBorderPaintOrder([dark, light]), [light, dark]);
  }
  const black = edge(), half = edge({ color: 'rgba(0,0,0,.5)' }), clear = edge({ color: '#0000' });
  assert.deepEqual(gridBorderPaintOrder([black, half, clear]), [clear, half, black]);
  const unknown = [undefined, 'var(--border)', 'currentColor', 'rgb(bad)', '#12345'].map(color => edge({ color }));
  assert.deepEqual(gridBorderPaintOrder([black, ...unknown]), [...unknown, black]);
  assert.deepEqual(gridBorderPaintOrder([]), []);
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


test('이중선은 DPR·배율마다 동일한 한 픽셀 두 획과 한 픽셀 간격', () => {
  for (const zoom of [.5, .75, .8, 1, 1.25]) for (const dpr of [1, 1.25, 1.5, 2]) {
    assert.ok(Math.abs(gridLineWidth(3, zoom, dpr, 'double') * zoom * dpr - 3) < 1e-8);
  }
});
