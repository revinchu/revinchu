import { test } from 'node:test';
import assert from 'node:assert/strict';
import { menuNavigationTarget } from '../src/menu-navigation.js';

const command = top => ({ group: null, rect: { left: 0, top, width: 160, height: 24 } });
const grid = (cols, count, { group = 'colors', left = 10, top = 30, width = 16, height = 16, gap = 3 } = {}) => Array.from({ length: count }, (_, i) => ({ group, rect: { left: left + (i % cols) * (width + gap), top: top + Math.floor(i / cols) * (height + gap), width, height } }));

test('자동 명령 다음 방향키는 첫 색에 진입한다', () => {
  const items = [command(0), ...grid(10, 20), command(90)];
  assert.equal(menuNavigationTarget(items, 0, 'ArrowDown'), 1);
  assert.equal(menuNavigationTarget(items, 0, 'ArrowRight'), 1);
  assert.equal(menuNavigationTarget(items, 1, 'ArrowUp'), 0);
});

test('색은 같은 열의 위아래 행으로 이동한다', () => {
  const items = [command(0), ...grid(10, 30), command(100)];
  assert.equal(menuNavigationTarget(items, 4, 'ArrowDown'), 14);
  assert.equal(menuNavigationTarget(items, 14, 'ArrowDown'), 24);
  assert.equal(menuNavigationTarget(items, 24, 'ArrowUp'), 14);
});

test('모바일에서 바뀐 열 수를 실제 표시 위치로 따른다', () => {
  const desktop = grid(8, 24), mobile = grid(3, 24);
  assert.equal(menuNavigationTarget(desktop, 1, 'ArrowDown'), 9);
  assert.equal(menuNavigationTarget(mobile, 1, 'ArrowDown'), 4);
  assert.equal(menuNavigationTarget(mobile, 4, 'ArrowUp'), 1);
});

test('일부 견본이 비활성인 행에서는 남은 가장 가까운 열을 고른다', () => {
  const items = grid(3, 9).filter((_, i) => i !== 4);
  assert.equal(menuNavigationTarget(items, 1, 'ArrowDown'), 3);
  assert.equal(menuNavigationTarget(items, 4, 'ArrowDown'), 7);
});

test('마지막 짧은 행의 빈 열은 가장 가까운 견본으로 이동한다', () => {
  const items = grid(4, 10);
  assert.equal(menuNavigationTarget(items, 7, 'ArrowDown'), 9);
  assert.equal(menuNavigationTarget(items, 9, 'ArrowUp'), 5);
});

test('색 그리드 위아래 경계에서 일반 명령과 연결한다', () => {
  const items = [command(0), ...grid(3, 6), command(90)];
  assert.equal(menuNavigationTarget(items, 2, 'ArrowUp'), 0);
  assert.equal(menuNavigationTarget(items, 5, 'ArrowDown'), 7);
  assert.equal(menuNavigationTarget(items, 7, 'ArrowUp'), 6);
});

test('상단 명령 없는 채우기 팔레트 첫 행 위는 첫 견본에 머문다', () => {
  const items = [...grid(10, 20), command(90)];
  assert.equal(menuNavigationTarget(items, 0, 'ArrowUp'), 0);
  assert.equal(menuNavigationTarget(items, 5, 'ArrowUp'), 0);
});

test('갤러리 그룹 사이도 빠짐 없이 탐색한다', () => {
  const items = [...grid(3, 6), ...grid(2, 4, { group: 'standards', top: 110 }), command(170)];
  assert.equal(menuNavigationTarget(items, 4, 'ArrowDown'), 6);
  assert.equal(menuNavigationTarget(items, 6, 'ArrowUp'), 5);
  assert.equal(menuNavigationTarget(items, 9, 'ArrowDown'), 10);
});

test('좌우는 DOM 순서의 견본으로 이동하고 바깥에서 멈춘다', () => {
  const items = grid(3, 6);
  assert.equal(menuNavigationTarget(items, 2, 'ArrowRight'), 3);
  assert.equal(menuNavigationTarget(items, 3, 'ArrowLeft'), 2);
  assert.equal(menuNavigationTarget(items, 0, 'ArrowLeft'), 0);
  assert.equal(menuNavigationTarget(items, 5, 'ArrowRight'), 5);
});

test('Home과 End는 현재 견본 행, Ctrl 조합은 전체 메뉴 끝으로 간다', () => {
  const items = [command(0), ...grid(4, 10), command(150)];
  assert.equal(menuNavigationTarget(items, 6, 'Home'), 5);
  assert.equal(menuNavigationTarget(items, 6, 'End'), 8);
  assert.equal(menuNavigationTarget(items, 10, 'End'), 10);
  assert.equal(menuNavigationTarget(items, 6, 'Home', { whole: true }), 0);
  assert.equal(menuNavigationTarget(items, 6, 'End', { whole: true }), 11);
});

test('일반 세로 메뉴는 이전과 같이 순환하고 Home/End를 지원한다', () => {
  const items = [command(0), command(25), command(50)];
  assert.equal(menuNavigationTarget(items, 2, 'ArrowDown'), 0);
  assert.equal(menuNavigationTarget(items, 0, 'ArrowUp'), 2);
  assert.equal(menuNavigationTarget(items, 1, 'Home'), 0);
  assert.equal(menuNavigationTarget(items, 1, 'End'), 2);
});

test('스크롤 위치가 음수여도 행·열 관계는 보존한다', () => {
  const items = grid(4, 12, { top: -400 });
  assert.equal(menuNavigationTarget(items, 2, 'ArrowDown'), 6);
  assert.equal(menuNavigationTarget(items, 6, 'ArrowDown'), 10);
});

test('빈 메뉴와 현재 초점 없는 메뉴도 안전하다', () => {
  assert.equal(menuNavigationTarget([], -1, 'ArrowDown'), -1);
  const items = grid(2, 4);
  assert.equal(menuNavigationTarget(items, -1, 'ArrowDown'), 0);
  assert.equal(menuNavigationTarget(items, -1, 'ArrowUp'), 3);
  assert.equal(menuNavigationTarget(items, 10, 'Home'), 0);
});
