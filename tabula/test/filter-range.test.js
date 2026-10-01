import { test } from 'node:test';
import assert from 'node:assert/strict';
import { expandedFilterEnd } from '../src/filter-range.js';
import { Workbook } from '../src/workbook.js';

test('명시된 100만 행 필터는 기존 범위 셀을 다시 읽지 않는다', () => {
  const wb = { usedRange: () => ({ rows: 1048576, cols: 18 }), getCell: () => { throw new Error('기존 범위 재조회'); } };
  assert.equal(expandedFilterEnd(wb, 0, { r1: 1, r2: 1048575, c1: 0, c2: 17 }), 1048575);
});

test('필터 바로 아래에 붙인 행만 확장하고 빈 행과 옆의 별도 표를 넘지 않는다', () => {
  const wb = new Workbook();
  wb.setInput(0, 5, 1, '0');
  wb.setInput(0, 6, 2, '새 행');
  wb.setInput(0, 7, 3, '옆 표');
  wb.setInput(0, 8, 1, '분리된 표');
  assert.equal(expandedFilterEnd(wb, 0, { r1: 1, r2: 4, c1: 1, c2: 2 }), 6);
});

test('삭제된 데이터 때문에 저장된 필터 범위를 줄이지 않고 수식 셀도 확장한다', () => {
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '헤더');
  assert.equal(expandedFilterEnd(wb, 0, { r1: 0, r2: 40, c1: 0, c2: 0 }), 40);
  wb.setInput(0, 1, 0, '=IF(TRUE,"",1)');
  assert.equal(expandedFilterEnd(wb, 0, { r1: 0, r2: 0, c1: 0, c2: 0 }), 1);
});

test('새 행 확인은 사용된 열까지만 읽고 기존 행은 읽지 않는다', () => {
  const visited = [];
  const wb = { usedRange: () => ({ rows: 60003, cols: 2 }), getCell: (si, r, c) => { visited.push([r, c]); return { raw: c === 1 ? '값' : '' }; } };
  assert.equal(expandedFilterEnd(wb, 0, { r1: 1, r2: 59999, c1: 0, c2: 16383 }), 60002);
  assert.deepEqual(visited, [[60000, 0], [60000, 1], [60001, 0], [60001, 1], [60002, 0], [60002, 1]]);
});
