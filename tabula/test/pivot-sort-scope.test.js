import test from 'node:test';
import assert from 'node:assert/strict';
import { pivotSortScope, pivotSortIntersections } from '../src/pivot-sort-scope.js';
import { Workbook } from '../src/workbook.js';

const area = { r1: 2, c1: 3, r2: 6, c2: 5 };
const pivot = { name: '합성피벗', area, rows: ['상품'], cols: [], values: [{ field: '매출', agg: 'sum' }] };
const point = (r, c) => ({ r1: r, c1: c, r2: r, c2: c });

test('pivot sort routing uses owned bounds, including headers/totals and extra pivots', () => {
  const extra = { ...pivot, area: { r1: 10, c1: 3, r2: 14, c2: 5 } };
  const sheet = { pivot, pivotsExtra: [null, extra] };
  for (const [r, c, def] of [[2, 3, pivot], [6, 5, pivot], [12, 4, extra]]) {
    assert.deepEqual(pivotSortScope(sheet, point(r, c), { r, c }), { kind: 'pivot', def });
  }
  assert.deepEqual(pivotSortScope(sheet, area, { r: 3, c: 4 }), { kind: 'pivot', def: pivot });
  for (const [r, c] of [[1, 4], [7, 4], [3, 2], [3, 6]]) assert.equal(pivotSortScope(sheet, point(r, c), { r, c }).kind, 'none');
});

test('mixed ordinary/pivot and multi-pivot selections are never silently narrowed', () => {
  const sheet = { pivot, pivotsExtra: [{ ...pivot, area: { ...area, r1: 8, r2: 12 } }] };
  for (const range of [{ ...area, r1: 1 }, { ...area, c2: 6 }, { ...area, r2: 12 }, { r1: 0, c1: 0, r2: 1048575, c2: 16383 }]) {
    assert.equal(pivotSortScope(sheet, range, { r: 3, c: 4 }).kind, 'mixed');
  }
  assert.equal(pivotSortScope(sheet, area, { r: 0, c: 0 }).kind, 'mixed');
});

test('unmaterialized/invalid pivot bounds do not claim unrelated cells', () => {
  for (const def of [{ rows: ['商品'] }, { area: { ...area, r2: -1 } }, { area: { ...area, r2: NaN } }]) {
    assert.equal(pivotSortIntersections({ pivot: def }, area).length, 0);
  }
});

function fixture(extra = false) {
  const wb = new Workbook();
  [['B', 20], ['A', 10], ['총합계', 30]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, 3 + r, 3 + c, String(v))));
  wb.sheets[0][extra ? 'pivotsExtra' : 'pivot'] = extra ? [pivot] : pivot;
  return wb;
}
const cells = wb => JSON.stringify(wb.serialize().sheets[0].cells);

test('all ordinary sorting entry points reject pivot intersections before changing data or history', () => {
  const actions = [
    wb => wb.sortRange(0, 3, 3, 5, 4, 4, false),
    wb => wb.sortMulti(0, 3, 3, 5, 4, [{ at: 4, asc: false }]),
    wb => wb.sortMulti(0, 3, 3, 5, 4, [{ at: 3, on: 'fill', color: '#fff' }]),
    wb => wb.sortMulti(0, 3, 3, 5, 4, [{ at: 3 }], { byCols: true }),
    wb => wb.sortBlock(0, 0, 3, 2000, 4, 4, true),
  ];
  for (const extra of [false, true]) for (const action of actions) {
    const wb = fixture(extra), before = cells(wb), history = wb.undoStack.length;
    assert.throws(() => wb.transact(() => action(wb)), { code: 'PIVOT_SORT_RANGE' });
    assert.equal(cells(wb), before); assert.equal(wb.undoStack.length, history);
  }
});

test('raw source beside a pivot remains sortable, with undo/redo and literal grand-total item intact', () => {
  const wb = fixture(), before = cells(wb);
  [['총합계', 5], ['Z', 3], ['A', 1]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
  const snapshot = cells(wb);
  wb.transact(() => wb.sortRange(0, 0, 0, 2, 1, 1));
  assert.deepEqual([0, 1, 2].map(r => wb.getValue(0, r, 0)), ['A', 'Z', '총합계']);
  assert.equal(wb.getValue(0, 5, 3), '총합계');assert.equal(wb.getValue(0, 5, 4), 30);
  wb.undo();assert.equal(cells(wb), snapshot);wb.redo();assert.equal(wb.getValue(0, 0, 0), 'A');
  assert.equal(wb.sheets[0].pivot, pivot);assert.ok(before);
});
