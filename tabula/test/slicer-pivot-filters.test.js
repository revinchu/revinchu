import test from 'node:test';
import assert from 'node:assert/strict';
import { cubeFromRows, itemStats, itemText } from '../src/cube.js';
import { resolvePivot } from '../src/pivot.js';
import { slicerPivotFilters, slicerPivotFilterKey } from '../src/slicer-pivot-filters.js';

const available = (source, def, field = 'Item') => {
  const index = source.cube.header.findIndex(h => h.toLowerCase() === field.toLowerCase());
  const stats = itemStats(source.cube, index, slicerPivotFilters(source, def, field));
  return stats.keys.filter((_, i) => stats.has[i]).map(itemText);
};
const sourceOf = (rows, date1904 = false) => ({ cube: cubeFromRows(rows), date1904 });
const dateRows = [['Item', 'Date', 'Value'], ['A', 45352, 1], ['B', 45353, 2], ['C', 45354, 3], ['D', 45383, 4], ['E', 44986, 5]];
const cases = [
  ['derived month', { groups: { Month: { base: 'Date', by: 'months' } }, filters: { Month: ['3월'] } }, dateRows, ['A', 'B', 'C', 'E']],
  ['group on source field', { groups: { Date: { by: 'months' } }, filters: { Date: ['3월'] } }, dateRows, ['A', 'B', 'C', 'E']],
  ['year AND month on one base', { groups: { Month: { base: 'Date', by: 'months' }, Year: { base: 'Date', by: 'years' } }, filters: { Month: ['3월'], Year: ['2024'] } }, dateRows, ['A', 'B', 'C']],
  ['number interval', { groups: { Bucket: { base: 'Value', by: 'number', start: 0, size: 10 } }, filters: { Bucket: ['10-19'] } }, [['Item', 'Value'], ['A', 10], ['B', 19], ['C', 20]], ['A', 'B']],
  ['manual item group', { groups: { Team: { base: 'Campaign', by: 'items', map: { North: 'One', South: 'One', West: 'Two' } } }, filters: { Team: ['One'] } }, [['Item', 'Campaign', 'Value'], ['A', 'North', 1], ['B', 'South', 2], ['C', 'West', 3]], ['A', 'B']],
  ['1904 month', { date1904: true, groups: { Month: { base: 'Date', by: 'months' } }, filters: { Month: ['3월'] } }, [['Item', 'Date', 'Value'], ['A', 43890, 1], ['B', 43891, 2], ['C', 43921, 3]], ['A', 'B']],
];
for (const [name, config, rows, expected] of cases) test(`${name}: cross-filter availability matches the pivot engine`, () => {
  const source = sourceOf(rows, !!config.date1904), def = { ...config, rows: ['Item'], values: [{ field: 'Value', agg: 'sum' }] }, before = structuredClone(def);
  assert.deepEqual(available(source, def), expected);
  assert.deepEqual(resolvePivot(source, def).rows.slice(1).map(r => r[0]), expected);
  assert.deepEqual(def, before);
});

test('only its own field is excluded, retaining other ordinary and grouped filters', () => {
  const source = sourceOf(dateRows), def = { groups: { Month: { base: 'Date', by: 'months' } }, filters: { ITEM: ['not-current'], Month: ['3월'], Value: ['2', '4'] } };
  assert.deepEqual(available(source, def, 'Item'), ['B']);
  assert.deepEqual(slicerPivotFilters(source, def, 'item').map(([i]) => i), [1, 2]);
});

test('case-insensitive field aliases and manual group selections follow item identity semantics', () => {
  const source = sourceOf([['Item', 'Campaign'], ['A', 'North'], ['B', 'West']]);
  const def = { groups: { TEAM: { base: 'campaign', by: 'items', map: { north: 'One', west: 'Two' } } }, filters: { team: ['ONE'] } };
  assert.deepEqual(available(source, def), ['A']);
});

test('empty grouped selection has zero matches rather than dropping the cross-filter', () => {
  const source = sourceOf(dateRows), def = { groups: { Month: { base: 'Date', by: 'months' } }, filters: { Month: [] } };
  const filters = slicerPivotFilters(source, def, 'Item');
  assert.equal(filters.length, 1);assert.equal(filters[0][1].size, 0);assert.deepEqual(available(source, def), []);
});

test('blank date group keeps blank rows and does not coerce blank into a date', () => {
  const source = sourceOf([['Item', 'Date', 'Value'], ['A', null, 1], ['B', 45352, 2]]);
  const def = { groups: { Month: { base: 'Date', by: 'months' } }, filters: { Month: ['(비어 있음)'] }, rows: ['Item'], values: [{ field: 'Value', agg: 'sum' }] };
  assert.deepEqual(available(source, def), ['A']);assert.deepEqual(resolvePivot(source, def).rows.slice(1).map(r => r[0]), ['A']);
});

test('source date system takes priority over stale group metadata', () => {
  const source = sourceOf([['Item', 'Date'], ['A', 43890], ['B', 43921]], true);
  const def = { date1904: false, groups: { Month: { base: 'Date', by: 'months', date1904: false } }, filters: { Month: ['3월'] } };
  assert.deepEqual(available(source, def), ['A']);
});

test('memo key ignores own selection but includes other filters, group definition, field and date system', () => {
  const source = sourceOf(dateRows), def = { groups: { Bucket: { base: 'Value', by: 'number', size: 10 } }, filters: { Item: ['A'], Bucket: ['0-9'] } };
  const key = slicerPivotFilterKey(source, def, 'item');
  assert.equal(key, slicerPivotFilterKey(source, { ...def, filters: { ...def.filters, Item: ['B'] } }, 'ITEM'));
  assert.notEqual(key, slicerPivotFilterKey(source, { ...def, filters: { ...def.filters, Bucket: ['10-19'] } }, 'Item'));
  assert.notEqual(key, slicerPivotFilterKey(source, { ...def, groups: { Bucket: { ...def.groups.Bucket, size: 5 } } }, 'Item'));
  assert.notEqual(key, slicerPivotFilterKey({ ...source, date1904: true }, def, 'Item'));
  assert.notEqual(key, slicerPivotFilterKey(source, def, 'Value'));
});

test('one million rows require only distinct keys, never rows, values or a new cube', () => {
  let calls = 0;
  const keys = [10, 15, 20], cube = { n: 1_000_000, header: ['Item', 'Value'], col(index) { assert.equal(index, 1);calls++;return { dim: () => ({ keys }), get() { throw Error('row value read'); } }; }, row() { throw Error('row materialization'); } };
  const source = { cube }, def = { groups: { Bucket: { base: 'Value', by: 'number', size: 10 } }, filters: { Bucket: ['10-19'] } }, before = structuredClone(def);
  const filters = slicerPivotFilters(source, def, 'Item');
  assert.equal(source.cube, cube);assert.equal(calls, 1);assert.deepEqual([...filters[0][1]], ['10', '15']);assert.deepEqual(keys, [10, 15, 20]);assert.deepEqual(def, before);
});

test('returned ordinary and grouped filter sets do not share mutable selection state', () => {
  const source = sourceOf(dateRows), def = { groups: { Month: { base: 'Date', by: 'months' } }, filters: { Value: ['1'], Month: ['3월'] } }, before = structuredClone(def);
  const first = slicerPivotFilters(source, def, 'Item');first[0][1].clear();first[1][1].clear();
  assert.deepEqual(def, before);assert.deepEqual(available(source, def), ['A']);
});

test('missing unrelated fields do not break valid cross-filters', () => {
  const source = sourceOf(dateRows), def = { filters: { DeletedField: ['X'], Value: ['2'] } };
  assert.deepEqual(available(source, def), ['B']);assert.deepEqual(slicerPivotFilters(null, def, 'Item'), []);
});
