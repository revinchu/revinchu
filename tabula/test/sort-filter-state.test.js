import test from 'node:test';
import assert from 'node:assert/strict';
import { sortScope, refreshSortedFilters } from '../src/sort-filter-state.js';
import { Workbook } from '../src/workbook.js';
import { tableFilterRange } from '../src/tables.js';
import { expandedFilterEnd } from '../src/filter-range.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const range = { r1: 0, c1: 0, r2: 3, c2: 1 };
const key = { at: 1, on: 'value', asc: true };
function book(rows = [['Item', 'Value'], ['A', 3], ['B', 1], ['A', 2]]) {
  const wb = new Workbook();
  rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
  wb.sheets[0].filter = { ...range, criteria: { 0: ['A'] }, hidden: { 2: true }, hiddenButtons: { 1: true } };
  return wb;
}
function visible(wb, filter = wb.sheets[0].filter) {
  return Array.from({ length: filter.r2 - filter.r1 }, (_, i) => i + filter.r1 + 1)
    .filter(r => !filter.hidden[r]).map(r => [wb.getValue(0, r, 0), wb.getValue(0, r, 1)]);
}
function refreshed(wb, range, filters, keys, predicate, extra = {}) {
  return refreshSortedFilters({ range, filters, keys, ...extra, recompute: (f, id) => ({ ...f, hidden: Object.fromEntries(
    Array.from({ length: f.r2 - f.r1 }, (_, i) => i + f.r1 + 1).filter(r => !predicate(r, id)).map(r => [r, true])) }) });
}

test('text-only table keeps its explicit header and excludes the total row despite heuristic=false', () => {
  const table = { ...range, r2: 4, header: true, totals: true };
  assert.deepEqual(sortScope({ range: { ...range, r2: 4 }, activeOnly: true, table, header: false }), { range, header: true });
});

test('selecting a complete table also excludes its total row; partial selection stays explicit', () => {
  const table = { ...range, r2: 4, header: true, totals: true };
  assert.deepEqual(sortScope({ range: table, table }), { range, header: true });
  const partial = { r1: 2, c1: 0, r2: 4, c2: 1 };
  assert.deepEqual(sortScope({ range: partial, table, header: false }), { range: partial, header: false });
  const noHeader = { ...table, header: false };
  assert.deepEqual(sortScope({ range: noHeader, table: noHeader, activeOnly: true, header: true }), { range, header: false });
});

test('active cell inside a range filter uses its full explicit data boundary and header', () => {
  const fallback = { r1: 1, c1: 0, r2: 1, c2: 0 };
  assert.deepEqual(sortScope({ range: fallback, activeOnly: true, active: { r: 2, c: 1 }, filter: range }), { range, header: true });
  assert.deepEqual(sortScope({ range: fallback, activeOnly: true, active: { r: 8, c: 1 }, filter: range }), { range: fallback, header: false });
});

test('custom sort re-evaluates checklist visibility at the new row positions in the same Undo entry', () => {
  const wb = book(), before = wb.sheets[0].filter, moved = { ...range, r1: 1 };
  assert.deepEqual(visible(wb), [['A', 3], ['A', 2]]);
  wb.transact(() => {
    wb.sortMulti(0, 1, 0, 3, 1, [key], {});
    for (const [, f] of refreshed(wb, moved, [['', wb.sheets[0].filter]], [key], r => wb.getValue(0, r, 0) === 'A')) wb.setSheetProp(0, 'filter', f);
  });
  assert.deepEqual(visible(wb), [['A', 2], ['A', 3]]);
  assert.deepEqual(wb.sheets[0].filter.hidden, { 1: true });
  assert.deepEqual(wb.sheets[0].filter.sort, { col: 1, asc: true });
  assert.deepEqual(wb.sheets[0].filter.hiddenButtons, { 1: true });
  assert.deepEqual(before.hidden, { 2: true }, 'saved filter object is immutable');
  wb.undo();assert.deepEqual(visible(wb), [['A', 3], ['A', 2]]);assert.deepEqual(wb.sheets[0].filter.hidden, { 2: true });
  wb.redo();assert.deepEqual(visible(wb), [['A', 2], ['A', 3]]);
});

test('numeric predicate uses freshly sorted cells rather than the previous hidden bitmap', () => {
  const wb = book([['Item', 'Value'], ['C', 10], ['A', 20], ['B', 30]]), moved = { ...range, r1: 1 };
  const f = { ...range, criteria: { 1: { type: 'custom', op1: 'gt', v1: '15' } }, hidden: { 1: true } };
  wb.sortMulti(0, 1, 0, 3, 1, [{ at: 0, on: 'value', asc: true }], {});
  const [[, next]] = refreshed(wb, moved, [['table', f]], [{ at: 0, on: 'value', asc: true }], r => wb.getValue(0, r, 1) > 15);
  assert.deepEqual(visible(wb, next), [['A', 20], ['B', 30]]);assert.deepEqual(next.hidden, { 3: true });
  assert.deepEqual(next.criteria, f.criteria);
});

test('every overlapping filter receives its own key and evaluator, while unrelated tables stay untouched', () => {
  const a = { ...range }, b = { ...range, c1: 2, c2: 3 }, far = { ...range, r1: 9, r2: 12 }, calls = [];
  const result = refreshSortedFilters({ range: { r1: 1, c1: 1, r2: 3, c2: 2 }, filters: [['', a], ['table', b], ['far', far]], keys: [key], recompute: (f, id) => { calls.push(id);return { ...f, hidden: { 2: true } }; } });
  assert.deepEqual(calls, ['', 'table']);assert.deepEqual(result.map(([id]) => id), ['', 'table']);
  assert.deepEqual(result.map(([, f]) => f.hidden), [{ 2: true }, { 2: true }]);
});

test('header-only and empty sort ranges do not recompute untouched data filters', () => {
  const recompute = () => assert.fail('header does not move data');
  for (const moved of [{ ...range, r2: 0 }, { ...range, r1: 4 }]) assert.deepEqual(refreshSortedFilters({ range: moved, filters: [['', range]], keys: [key], recompute }), []);
});

test('complex, partial and horizontal sorts clear stale single-column metadata but preserve criteria', () => {
  const f = { ...range, sort: { col: 0, asc: false }, criteria: { 0: ['A'] } };
  const cases = [
    { keys: [key, { at: 0, asc: false }] }, { keys: [{ ...key, on: 'fill', color: '#ff0000' }] },
    { keys: [{ ...key, list: ['A', 'B'] }] }, { keys: [key], byCols: true },
    { keys: [key], range: { ...range, r1: 2 } }, { keys: [key], options: { caseSensitive: true } }, { keys: [key], options: { natural: true } },
  ];
  for (const extra of cases) {
    const [[, next]] = refreshSortedFilters({ range: { ...range, r1: 1 }, filters: [['', f]], recompute: x => x, ...extra });
    assert.equal(next.sort, undefined);assert.deepEqual(next.criteria, { 0: ['A'] });
  }
  assert.deepEqual(f.sort, { col: 0, asc: false });
});

test('multikey table sort keeps header and total values fixed and recalculates visibility', () => {
  const wb = book([['Item', 'Value'], ['A', 3], ['B', 1], ['A', 2], ['Total', 6]]);
  const table = { id: 't', name: 'Table1', r1: 0, c1: 0, r2: 4, c2: 1, header: true, totals: true, filter: { criteria: { 0: ['A'] }, hidden: { 2: true } } };
  const scope = sortScope({ range: table, activeOnly: true, table, header: false }), keys = [{ at: 0, on: 'value', asc: true }, key];
  const moved = { ...scope.range, r1: scope.range.r1 + 1 };
  wb.sortMulti(0, moved.r1, moved.c1, moved.r2, moved.c2, keys, {});
  const [[, next]] = refreshed(wb, moved, [['t', tableFilterRange(table)]], keys, r => wb.getValue(0, r, 0) === 'A');
  assert.deepEqual(visible(wb, next), [['A', 2], ['A', 3]]);
  assert.deepEqual([wb.getValue(0, 0, 0), wb.getValue(0, 4, 0), wb.getValue(0, 4, 1)], ['Item', 'Total', 6]);
  assert.equal(next.sort, undefined);
});

test('refreshed range filter survives native XLSX save/read with the sorted data', () => {
  const wb = book(), moved = { ...range, r1: 1 };
  wb.transact(() => {
    wb.sortMulti(0, 1, 0, 3, 1, [key], {});
    const [[, filter]] = refreshed(wb, moved, [['', wb.sheets[0].filter]], [key], r => wb.getValue(0, r, 0) === 'A');
    wb.setSheetProp(0, 'filter', filter);
  });
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(visible(back), [['A', 2], ['A', 3]]);
  assert.deepEqual(back.sheets[0].filter.criteria, { 0: ['A'] });
  assert.deepEqual(back.sheets[0].filter.sort, { col: 1, asc: true });
});


test('new contiguous rows stay in the normalized filter sort scope and its saved sort state', () => {
  const wb = book(), old = { ...wb.sheets[0].filter, r2: 2 };
  const expanded = { ...old, r2: expandedFilterEnd(wb, 0, old) };
  assert.equal(expanded.r2, 3);
  const scope = sortScope({ range, activeOnly: true, active: { r: 1, c: 1 }, filter: expanded, header: false });
  assert.deepEqual(scope, { range, header: true });
  const moved = { ...scope.range, r1: 1 };
  wb.sortMulti(0, moved.r1, moved.c1, moved.r2, moved.c2, [key], {});
  const [[, next]] = refreshed(wb, moved, [['', expanded]], [key], r => wb.getValue(0, r, 0) === 'A');
  assert.deepEqual(visible(wb, next), [['A', 2], ['A', 3]]);
  assert.deepEqual(next.sort, { col: 1, asc: true });assert.equal(next.r2, 3);
});
