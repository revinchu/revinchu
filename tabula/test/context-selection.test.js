import { test } from 'node:test';
import assert from 'node:assert/strict';
import { contextMenuKind, contextContains, selectionAxisRanges, selectionAxisTargets } from '../src/context-selection.js';

const sel = { r1: 2, c1: 1, r2: 8, c2: 6 };
const cells = [[8, 6], [2, 1], [2, 6], [7, 1]];

test('본문 우클릭은 전체 행·열 선택의 메뉴를 유지하고 헤더는 해당 축을 사용한다', () => {
  assert.equal(contextMenuKind('rows'), 'row');
  assert.equal(contextMenuKind('cols'), 'col');
  assert.equal(contextMenuKind('cells'), 'cell');
  assert.equal(contextMenuKind('all'), 'cell');
  assert.equal(contextMenuKind('rows', 'col'), 'col');
  assert.equal(contextMenuKind('cols', 'row'), 'row');
});

test('희소 선택 사이의 미선택 셀을 선택 안으로 오인하지 않는다', () => {
  assert.equal(contextContains(sel, cells, 3, 3), false);
  assert.equal(contextContains(sel, cells, 8, 6), true);
  assert.equal(contextContains(sel, null, 3, 3), true);
  assert.equal(contextContains(sel, null, 9, 3), false);
  assert.equal(contextContains(sel, [], 2, 1), false);
  assert.equal(contextContains(sel, null, 2.5, 1), false);
});

test('비연속 행·열 작업은 실제 축 합집합만 처리하고 원본 선택을 변경하지 않는다', () => {
  const before = structuredClone(cells);
  assert.deepEqual(selectionAxisTargets(sel, 'row', { cells }), [2, 7, 8]);
  assert.deepEqual(selectionAxisTargets(sel, 'col', { cells }), [1, 6]);
  assert.deepEqual(selectionAxisRanges(sel, 'row', { cells }), [[2, 2], [7, 8]]);
  assert.deepEqual(cells, before);
  assert.deepEqual(selectionAxisTargets(sel, 'row', { cells: [] }), []);
});

test('처리 한도를 넘은 선택은 절대 일부만 처리하지 않는다', () => {
  assert.deepEqual(selectionAxisTargets(sel, 'row', { cells, limit: 3 }), [2, 7, 8]);
  assert.throws(() => selectionAxisTargets(sel, 'row', { cells, limit: 2 }), (e) => e.code === 'AXIS_SELECTION_LIMIT' && e.count === 3 && e.limit === 2);
  assert.throws(() => selectionAxisTargets({ r1: 0, r2: 1048575 }, 'row'), (e) => e.code === 'AXIS_SELECTION_LIMIT' && e.count === 1048576);
  assert.deepEqual(selectionAxisRanges({ r1: 0, r2: 1048575 }, 'row'), [[0, 1048575]]);
});

test('잘못된 축·좌표·한도를 거부한다', () => {
  for (const fn of [() => selectionAxisTargets(sel, 'rows'), () => selectionAxisTargets({ r1: 5, r2: 2 }, 'row'), () => selectionAxisTargets(sel, 'row', { cells: [[NaN, 1]] }), () => selectionAxisTargets(sel, 'col', { limit: -1 })]) assert.throws(fn, (e) => e.code === 'INVALID_SELECTION');
});
