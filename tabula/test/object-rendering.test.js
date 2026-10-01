import { test } from 'node:test';
import assert from 'node:assert/strict';
import { objectIntersectsWindow } from '../src/view.js';

const rect = { x1: 100, y1: 200, x2: 800, y2: 900 };
const obj = (x, y, w = 100, h = 100, extra = {}) => ({ x, y, w, h, ...extra });
test('개체 생성 전 숨김·화면 밖을 제외하고 부분 교차는 유지', () => {
  assert.equal(objectIntersectsWindow(obj(300, 400), rect), true);
  assert.equal(objectIntersectsWindow(obj(300, 400, 100, 100, { hidden: true }), rect), false);
  for (const [x, y] of [[-500, 400], [1000, 400], [300, -200], [300, 1100]]) assert.equal(objectIntersectsWindow(obj(x, y), rect), false);
  assert.equal(objectIntersectsWindow(obj(780, 250), rect), true);
  assert.equal(objectIntersectsWindow(obj(300, 890), rect), true);
});
test('회전·그림자·선과 선택 단추 여유가 렌더 창에 닿으면 유지', () => {
  assert.equal(objectIntersectsWindow(obj(880, 350, 20, 240), rect), false);
  assert.equal(objectIntersectsWindow(obj(880, 350, 20, 240, { rot: 90 }), rect), true);
  assert.equal(objectIntersectsWindow(obj(900, 350, 100, 100, { shadow: { dx: -100, blur: 10 } }), rect), true);
  assert.equal(objectIntersectsWindow(obj(825, 350, 20, 0), rect), true);
  assert.equal(objectIntersectsWindow(obj(920, 350, 20, 20, { glow: { size: 50 } }), rect), true);
});
test('경계를 가로지르는 같은 개체는 고정·스크롤 창 양쪽에 존재', () => {
  const item = obj(80, 80, 100, 100);
  for (const r of [{ x1: 0, y1: 0, x2: 100, y2: 100 }, { x1: 100, y1: 100, x2: 500, y2: 500 }]) assert.equal(objectIntersectsWindow(item, r), true);
});
