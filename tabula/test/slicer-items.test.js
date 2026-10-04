import test from 'node:test';
import assert from 'node:assert/strict';
import { slicerDisplayModel } from '../src/slicer-items.js';
import { slicerPickValues, pivotFieldItemModel, pivotItemSelection } from '../src/pivot-field-items.js';
import { cachedSlicerItems, applyCachedSlicerSelection } from '../src/slicer-cache.js';
import { Workbook } from '../src/workbook.js';
import { cubeFromRows, itemStats } from '../src/cube.js';
import { pivotSourceData } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';

const names = ['A', 'B', 'C', 'D', 'E', 'F'];
const rows = [['Item', 'Campaign', 'Value'], ...names.map((name, i) => [name, i < 3 ? 'keep' : 'other', i + 1])];
const raw = (patch = {}) => ({ filtered: false, items: names.map((key, i) => ({ key, text: key, v: key, hasData: i < 3, selected: true })), ...patch });
const visible = m => m.items.map(i => i.text);
const selected = m => m.allItems.filter(i => i.selected).map(i => i.text);
const display = (sl, model) => slicerDisplayModel(sl, model, []);
const sl = kind => ({ id: 's', caption: 'Item', source: { kind }, hideNoData: true, x: 10, y: 100, w: 140, h: 200, style: 'SlicerStyleLight1' });

for (const kind of ['pivot', 'table', 'cache']) test(`${kind}: Ctrl/multi first pick preserves 3 visible items and hidden selections`, () => {
  const settings = sl(kind), original = raw(), before = structuredClone(original);
  const first = display(settings, original);
  assert.deepEqual(visible(first), ['A', 'B', 'C']);
  const choice = slicerPickValues(first, 'A', true);
  assert.deepEqual(choice, ['B', 'C', 'D', 'E', 'F']);
  const next = raw({ filtered: true, items: original.items.map(i => ({ ...i, selected: choice.includes(i.key) })) });
  const shown = display(settings, next);
  assert.deepEqual(visible(shown), ['A', 'B', 'C']);
  assert.deepEqual(selected(shown), ['B', 'C', 'D', 'E', 'F']);
  assert.deepEqual(original, before);
});

test('changing another slicer reveals only newly available items with their retained selection', () => {
  const first = raw({ filtered: true });first.items[0].selected = false;
  const shown = display(sl('pivot'), first);
  const switched = { ...first, items: first.items.map(i => ({ ...i, hasData: ['A', 'D', 'E'].includes(i.key) })) };
  const next = display(sl('pivot'), switched);
  assert.deepEqual(visible(shown), ['A', 'B', 'C']);
  assert.deepEqual(visible(next), ['A', 'D', 'E']);
  assert.deepEqual(next.items.map(i => i.selected), [false, true, true]);
  assert.deepEqual(selected(next), ['B', 'C', 'D', 'E', 'F']);
});

test('no matching data means no visible items even when all values remain explicitly selected', () => {
  const source = raw({ filtered: true });source.items = source.items.map(i => ({ ...i, hasData: false }));
  const shown = display(sl('pivot'), source);
  assert.deepEqual(shown.items, []);assert.equal(shown.allItems.length, 6);assert.equal(selected(shown).length, 6);
  assert.deepEqual(slicerPickValues(shown, 'D', true), ['A', 'B', 'C', 'E', 'F']);
});

test('hide-no-data wins over markNoData=false; the full universe retains real availability', () => {
  const source = raw({ filtered: true }), before = structuredClone(source);
  const shown = display({ ...sl('pivot'), markNoData: false }, source);
  assert.deepEqual(visible(shown), ['A', 'B', 'C']);
  assert.equal(shown.allItems.find(i => i.key === 'D').hasData, false);
  assert.deepEqual(source, before);
});

test('show-deleted and hide-no-data are independent and neither drops retained selections', () => {
  const source = raw({ filtered: true });source.items[0].deleted = true;source.items[3].deleted = true;
  const hidden = display(sl('pivot'), source);
  assert.deepEqual(visible(hidden), ['B', 'C']);assert.equal(hidden.allItems.length, 6);
  const history = display({ ...sl('pivot'), showDeleted: true }, source);
  assert.deepEqual(visible(history), ['A', 'B', 'C']);
  const all = display({ ...sl('pivot'), showDeleted: true, hideNoData: false }, source);
  assert.deepEqual(visible(all), names);assert.deepEqual(selected(all), names);
});

test('custom-list order, descending, blank-last, and no-data-last affect display only', () => {
  const source = { filtered: true, items: ['Mon', 'Tue', 'Wed', ''].map((key, i) => ({ key, text: key, v: key, selected: true, hasData: i !== 1 })) };
  const before = structuredClone(source);
  const settings = { source: { kind: 'pivot' }, hideNoData: false, sort: 'desc', showDeleted: true };
  const normal = slicerDisplayModel(settings, source, [['Wed', 'Mon', 'Tue']]);
  assert.deepEqual(visible(normal), ['Mon', 'Wed', '', 'Tue']);
  const noLast = slicerDisplayModel({ ...settings, noDataLast: false }, source, [['Wed', 'Mon', 'Tue']]);
  assert.deepEqual(visible(noLast), ['Tue', 'Mon', 'Wed', '']);
  const noCustom = slicerDisplayModel({ ...settings, customList: false, noDataLast: false }, source, [['Wed', 'Mon', 'Tue']]);
  assert.deepEqual(visible(noCustom), ['Wed', 'Tue', 'Mon', '']);
  assert.deepEqual(source, before);assert.deepEqual(normal.allItems, source.items);
});

test('two slicer views over one raw model cannot mutate one another', () => {
  const source = raw({ filtered: true }), before = structuredClone(source);
  const compact = display({ ...sl('pivot'), markNoData: false, sort: 'desc' }, source);
  const full = display({ ...sl('pivot'), hideNoData: false, noDataLast: false }, source);
  assert.deepEqual(visible(compact), ['C', 'B', 'A']);assert.deepEqual(visible(full), names);
  assert.equal(full.items.find(i => i.key === 'D').hasData, false);
  assert.deepEqual(source, before);assert.equal(compact.allItems.length, full.allItems.length);
});

test('single-click selection and same-single-click clearing keep their prior meaning', () => {
  const source = raw({ filtered: true });source.items = source.items.map(i => ({ ...i, selected: i.key === 'B' }));
  const shown = display(sl('table'), source);
  assert.deepEqual(slicerPickValues(shown, 'A', false), ['A']);
  assert.equal(slicerPickValues(shown, 'B', false), null);
});

function book(kind) {
  const wb = new Workbook();wb.sheets[0].name = 'Source';
  rows.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
  const settings = sl(kind);
  if (kind === 'table') {
    wb.sheets[0].tables = [{ id: 't', name: 'Data', r1: 0, c1: 0, r2: 6, c2: 2, header: true, style: 'TableStyleMedium2', filter: { criteria: { 0: names.slice(1), 1: ['keep'] }, hidden: {} } }];
    settings.source = { kind, table: 'Data', column: 'Item' };wb.sheets[0].slicers = [settings];
  } else if (kind === 'pivot') {
    const report = wb.addSheet('Report');
    wb.sheets[report].pivot = { name: 'ReportPivot', source: 'Source', range: { r1: 0, c1: 0, r2: 6, c2: 2 }, rows: ['Item'], pages: ['Campaign'], values: [{ field: 'Value', agg: 'sum' }], filters: { Campaign: ['keep'], Item: names.slice(1) }, top: 3, left: 0 };
    settings.source = { kind, field: 'Item', pivots: [{ sheet: 'Report', name: 'ReportPivot' }] };wb.sheets[report].slicers = [settings];
  } else {
    settings.source = { kind, field: 'Item', cacheKey: 'shared-items', values: names, items: names.map((_, index) => ({ index, hasData: index < 3 })) };
    settings.cacheSelection = ['1', '2', '3', '4', '5'];wb.sheets[0].slicers = [settings, { ...structuredClone(settings), id: 's-copy', x: 200 }];
  }
  return wb;
}
function model(wb, settings) {
  if (settings.source.kind === 'cache') return display(settings, { items: cachedSlicerItems(settings), filtered: Array.isArray(settings.cacheSelection) });
  let cube, filters;
  if (settings.source.kind === 'pivot') {
    const def = wb.sheets.find(s => s.pivot)?.pivot, source = pivotSourceData(wb, def);cube = source.cube;filters = def.filters;
    const other = Object.entries(filters).filter(([name]) => name.toLowerCase() !== 'item').map(([name, values]) => [cube.header.findIndex(h => h.toLowerCase() === name.toLowerCase()), new Set(values)]);
    const stats = itemStats(cube, 0, other), field = pivotFieldItemModel(wb, def, 'Item', { source, currentItems: stats.keys.map((v, i) => ({ v, hasData: !!stats.has[i] })) });
    return display(settings, { items: field.items.map(i => ({ ...i, key: i.id })), filtered: field.filtered, preserveSelection: true });
  }
  const table = wb.sheets[0].tables[0];filters = table.filter.criteria;
  cube = cubeFromRows(Array.from({ length: 7 }, (_, r) => Array.from({ length: 3 }, (_, c) => wb.getValue(0, r, c))));
  const stats = itemStats(cube, 0, [[1, new Set(filters[1])]]), chosen = new Set(filters[0]);
  return display(settings, { filtered: true, items: stats.keys.map((v, i) => ({ key: stats.texts[i], text: stats.texts[i], v, selected: chosen.has(stats.texts[i]), hasData: !!stats.has[i] })) });
}
for (const kind of ['table', 'pivot', 'cache']) test(`${kind}: XLSX native settings retain hidden selections without expanding visible items`, () => {
  const wb = book(kind), original = wb.sheets.flatMap(s => s.slicers ?? [])[0], bytes = writeXlsx(wb), files = unzip(bytes);
  const cacheXml = Object.keys(files).filter(p => /^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(p)).map(p => textOf(files[p])).join('');
  assert.match(cacheXml, /slicerCacheHideItemsWithNoData/);
  assert.deepEqual(visible(model(wb, original)), ['A', 'B', 'C']);
  const restored = new Workbook(readXlsx(bytes).data), settings = restored.sheets.flatMap(s => s.slicers ?? [])[0], shown = model(restored, settings);
  assert.equal(settings.hideNoData, true);assert.equal(settings.source.kind, kind);
  assert.deepEqual(visible(shown), ['A', 'B', 'C']);assert.deepEqual(selected(shown), ['B', 'C', 'D', 'E', 'F']);
  assert.equal(shown.allItems.length, 6);
  if (kind === 'table') assert.deepEqual(restored.sheets[0].tables[0].filter.criteria[1], ['keep']);
  if (kind === 'pivot') assert.deepEqual(restored.sheets.find(s => s.pivot).pivot.filters.Campaign, ['keep']);
  if (kind === 'cache') {
    applyCachedSlicerSelection(restored, settings.source, ['1', '2', '4']);
    const copies = restored.sheets[0].slicers;assert.equal(copies.length, 2);
    for (const copy of copies) { assert.deepEqual(copy.cacheSelection, ['1', '2', '4']);assert.deepEqual(visible(model(restored, copy)), ['A', 'B', 'C']); }
    restored.undo();for (const copy of restored.sheets[0].slicers) assert.deepEqual(copy.cacheSelection, ['1', '2', '3', '4', '5']);
  }
});

test('real pivot item identities preserve the other filter and hidden members after additive choice', () => {
  const wb = book('pivot'), def = wb.sheets[1].pivot, settings = wb.sheets[1].slicers[0];delete def.filters.Item;
  const source = pivotSourceData(wb, def), stats = itemStats(source.cube, 0, [[1, new Set(['keep'])]]);
  const field = pivotFieldItemModel(wb, def, 'Item', { source, currentItems: stats.keys.map((v, i) => ({ v, hasData: !!stats.has[i] })) });
  const first = model(wb, settings), choice = pivotItemSelection(field, slicerPickValues(first, first.items[0].key, true));
  def.filters = { ...def.filters, Item: choice.values };
  assert.deepEqual(def.filters, { Campaign: ['keep'], Item: ['B', 'C', 'D', 'E', 'F'] });
  assert.deepEqual(visible(model(wb, settings)), ['A', 'B', 'C']);assert.equal(model(wb, settings).allItems.length, 6);
});
