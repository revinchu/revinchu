import test from 'node:test';
import assert from 'node:assert/strict';
import { pivotSortPatch } from '../src/pivot-sort-state.js';
import { sortKeys, EMPTY, EMPTY_TEXT } from '../src/cube.js';
import { computePivot, resolvePivot, normalizeDef, pivotSourceData } from '../src/pivot.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml } from '../src/xml.js';

const path = 'xl/pivotTables/pivotTable1.xml';
const base = { rows: ['Item'], values: [{ field: 'Value', agg: 'sum' }], layout: 'tabular', sort: { Item: { dir: 'asc' } } };
function grid(rows, def) { const res = resolvePivot(rows, def);return computePivot(res, res.def).grid.map(row => row.map(cell => cell.raw)); }
function labels(rows, def) { return grid(rows, def).slice(1, -1).map(row => row[0].replace(/^'/, '')); }
function fixture(options = {}, values = ['A', 'B', 'C']) {
  const wb = new Workbook();wb.sheets[0].name = 'Source';
  [['Item', 'Value'], ...values.map((v, i) => [v, i + 1])].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
  const si = wb.addSheet('Report');
  wb.sheets[si].pivot = { ...base, name: 'ReportPivot', source: 'Source', range: { r1: 0, c1: 0, r2: values.length, c2: 1 }, top: 0, left: 0, ...options };
  return wb;
}
function bookLabels(wb) {
  const def = wb.sheets[1].pivot, source = pivotSourceData(wb, def);
  return labels(source, def);
}

// The switch governs field-caption sorting, including initial field ordering and
// user-requested sorting: https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.sortusingcustomlists
test('caption ordering honors customListSort true/false for months in both directions', () => {
  const rows = [['Item', 'Value'], ['10월', 10], ['2월', 2], ['1월', 1]];
  for (const [customListSort, expected] of [[true, ['1월', '2월', '10월']], [false, ['10월', '1월', '2월']]]) {
    assert.deepEqual(labels(rows, { ...base, customListSort }), expected);
    assert.deepEqual(labels(rows, { ...base, customListSort, sort: { Item: { dir: 'desc' } } }), [...expected].reverse());
    assert.deepEqual(labels(rows, { ...base, customListSort, sort: {} }), expected, 'initial/default caption sort follows the setting');
  }
});

test('weekday list can be disabled and toggled repeatedly without reusing stale sorted results', () => {
  const rows = [['Item', 'Value'], ...['월요일', '화요일', '수요일', '목요일', '금요일'].map(v => [v, 1])];
  const normal = ['금요일', '목요일', '수요일', '월요일', '화요일'];
  for (const enabled of [true, false, true, false]) {
    assert.deepEqual(labels(rows, { ...base, customListSort: enabled }), enabled ? rows.slice(1).map(row => row[0]) : normal);
  }
});

test('numeric, literal empty string and blank placement keep existing ordering with lists disabled', () => {
  const keys = [EMPTY, 'Tue', 10, EMPTY_TEXT, 2, 'Mon', 'Zulu'];
  assert.deepEqual(sortKeys([...keys], { customList: false }), [2, 10, EMPTY_TEXT, 'Mon', 'Tue', 'Zulu', EMPTY]);
  assert.deepEqual(sortKeys(['Fri', 'Mon', 'Tue']), ['Mon', 'Tue', 'Fri']);
  assert.deepEqual(sortKeys(['Fri', 'Mon', 'Tue'], { customList: false }), ['Fri', 'Mon', 'Tue']);
});

test('normalization retains the switch with the historical default enabled', () => {
  const header = ['Item', 'Value'];
  assert.equal(normalizeDef(base, header).customListSort, true);
  const def = normalizeDef({ ...base, customListSort: false }, header);
  assert.equal(def.customListSort, false);assert.equal(normalizeDef(def, header).customListSort, false);
});

test('value-based sorting is independent of the custom caption list switch', () => {
  const rows = [['Item', 'Value'], ['Mon', 20], ['Fri', 10], ['Tue', 30]];
  for (const customListSort of [true, false]) assert.deepEqual(labels(rows, { ...base, customListSort, sort: { Item: { dir: 'desc', by: 0 } } }), ['Tue', 'Mon', 'Fri']);
});

test('initial imported order remains unchanged until an explicit sort is requested', () => {
  const rows = [['Item', 'Value'], ['A', 1], ['B', 2], ['C', 3]];
  const def = { ...base, order: { Item: ['B', 'A', 'C'], Other: ['y', 'x'] } }, before = structuredClone(def);
  assert.deepEqual(labels(rows, def), ['B', 'A', 'C']);
  const patch = pivotSortPatch(def, { field: 'Item', sort: { dir: 'desc' } });
  assert.deepEqual(labels(rows, { ...def, ...patch }), ['C', 'B', 'A']);
  assert.deepEqual(patch.order, { Other: ['y', 'x'] });assert.deepEqual(def, before);
});

test('same-direction explicit sort also discards the old file order for only the requested field', () => {
  const def = { ...base, cols: ['Other'], sort: { Item: { dir: 'asc' }, Other: { dir: 'desc' } }, order: { Item: ['B', 'A'], Other: ['y', 'x'] } };
  const patch = pivotSortPatch(def, { field: 'Item', sort: { dir: 'asc' } });
  assert.deepEqual(patch.sort, def.sort);assert.deepEqual(patch.order, { Other: ['y', 'x'] });
});

test('manual mode keeps the saved order, other field sorts and value tie information', () => {
  const def = { ...base, sort: { Item: { dir: 'asc' }, Other: { dir: 'desc', by: 0 } }, order: { Item: ['B', 'A', 'C'] }, tieOrder: { Other: ['X', 'Y'] } };
  const next = { ...def, ...pivotSortPatch(def, { field: 'Item', sort: null }) };
  assert.deepEqual(next.sort, { Other: { dir: 'desc', by: 0 } });assert.equal(next.order, def.order);assert.equal(next.tieOrder, def.tieOrder);
  assert.deepEqual(labels([['Item', 'Value'], ['A', 1], ['B', 2], ['C', 3]], next), ['B', 'A', 'C']);
});

test('custom-list option changes invalidate automatic caption orders while keeping manual and value orders', () => {
  const def = { ...base, sort: { Item: { dir: 'asc' }, Amount: { dir: 'desc', by: 0 } }, order: { Item: ['Tue', 'Mon'], Manual: ['b', 'a'], Amount: ['x', 'y'] } };
  const patch = pivotSortPatch(def, { customListSort: false });
  assert.deepEqual(patch, { customListSort: false, order: { Manual: ['b', 'a'], Amount: ['x', 'y'] } });
  assert.deepEqual(pivotSortPatch(def, { customListSort: true }), {}, 'unchanged checkbox does not invalidate file order');
  const next = { ...def, ...patch }, enabled = pivotSortPatch(next, { customListSort: true });
  assert.ok(Object.hasOwn(enabled, 'customListSort'));assert.equal(enabled.customListSort, undefined);
});

test('combined dialog patch preserves both explicit sort and checkbox order invalidation', () => {
  const def = { ...base, rows: ['Item', 'Other'], sort: { Item: { dir: 'asc' }, Other: { dir: 'asc' } }, order: { Item: ['A', 'B'], Other: ['Tue', 'Mon'], Manual: ['y', 'x'] } };
  const next = { ...def, ...pivotSortPatch(def, { field: 'item', sort: { dir: 'desc', by: 0 }, customListSort: false }) };
  assert.deepEqual(next.sort, { Item: { dir: 'desc', by: 0 }, Other: { dir: 'asc' } });
  assert.deepEqual(next.order, { Manual: ['y', 'x'] });assert.equal(next.customListSort, false);
});

// Native attribute, independent of any private Wixel extension:
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivottabledefinition.customlistsort
test('customListSort is saved as native XML and false survives save/read without private extensions', () => {
  for (const enabled of [false, true]) {
    const wb = fixture({ customListSort: enabled }, ['Mon', 'Fri', 'Tue']), files = unzip(writeXlsx(wb));
    const xml = textOf(files[path]), attrs = parseXml(xml).attrs;
    assert.equal(attrs.customListSort === '0', !enabled);
    files[path] = xml.replace(/<extLst>[\s\S]*?<\/extLst>/g, '');
    const back = new Workbook(readXlsx(zip(files)).data);
    assert.equal(back.sheets[1].pivot.customListSort !== false, enabled);
    assert.deepEqual(bookLabels(back), enabled ? ['Mon', 'Tue', 'Fri'] : ['Fri', 'Mon', 'Tue']);
  }
});

test('external native XML boolean spellings and omitted default are imported correctly', () => {
  for (const [attribute, expected] of [['0', false], ['false', false], ['1', true], ['true', true], [null, true]]) {
    const files = unzip(writeXlsx(fixture({ customListSort: false })));
    files[path] = textOf(files[path]).replace(' customListSort="0"', attribute === null ? '' : ` customListSort="${attribute}"`);
    const back = new Workbook(readXlsx(zip(files)).data);
    assert.equal(back.sheets[1].pivot.customListSort !== false, expected);
  }
});

test('native imported ascending order responds to descending and remains descending after XLSX roundtrip', () => {
  const wb = new Workbook(readXlsx(writeXlsx(fixture())).data), before = wb.sheets[1].pivot;
  assert.deepEqual(before.order.Item, ['A', 'B', 'C']);assert.deepEqual(bookLabels(wb), ['A', 'B', 'C']);
  wb.transact(() => wb.setSheetProp(1, 'pivot', { ...before, ...pivotSortPatch(before, { field: 'Item', sort: { dir: 'desc' } }) }));
  assert.deepEqual(bookLabels(wb), ['C', 'B', 'A']);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[1].pivot.sort.Item.dir, 'desc');assert.deepEqual(bookLabels(back), ['C', 'B', 'A']);
  wb.undo();assert.deepEqual(bookLabels(wb), ['A', 'B', 'C']);assert.deepEqual(wb.sheets[1].pivot.order.Item, ['A', 'B', 'C']);
  wb.redo();assert.deepEqual(bookLabels(wb), ['C', 'B', 'A']);
});

test('custom-list setting and imported order are both restored by Undo/Redo', () => {
  const wb = new Workbook(readXlsx(writeXlsx(fixture({}, ['Mon', 'Fri', 'Tue']))).data), before = wb.sheets[1].pivot;
  wb.transact(() => wb.setSheetProp(1, 'pivot', { ...before, ...pivotSortPatch(before, { customListSort: false }) }));
  assert.deepEqual(bookLabels(wb), ['Fri', 'Mon', 'Tue']);
  wb.undo();assert.deepEqual(bookLabels(wb), ['Mon', 'Tue', 'Fri']);assert.deepEqual(wb.sheets[1].pivot.order.Item, before.order.Item);
  wb.redo();assert.deepEqual(bookLabels(wb), ['Fri', 'Mon', 'Tue']);assert.equal(wb.sheets[1].pivot.customListSort, false);
});


test('choosing manual mode and toggling custom lists together keeps that manual order', () => {
  const def = { ...base, order: { Item: ['B', 'A', 'C'] } };
  const patch = pivotSortPatch(def, { field: 'Item', sort: null, customListSort: false });
  assert.deepEqual(patch.sort, {});assert.equal(patch.order, undefined);assert.equal(patch.customListSort, false);
  assert.deepEqual(labels([['Item', 'Value'], ['A', 1], ['B', 2], ['C', 3]], { ...def, ...patch }), ['B', 'A', 'C']);
});


test('new source items make incomplete imported automatic order fall back to a fresh sort across save/read', () => {
  const wb = fixture({ order: { Item: ['B', 'A'] } });
  assert.deepEqual(bookLabels(wb), ['A', 'B', 'C']);
  assert.deepEqual(bookLabels(new Workbook(readXlsx(writeXlsx(wb)).data)), ['A', 'B', 'C']);
});
