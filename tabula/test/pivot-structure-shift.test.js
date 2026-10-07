import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
const rg = (r1, c1, r2, c2) => ({ r1, c1, r2, c2 });
function definition(name, top, left = 2) {
  return {
    name, source: 'Source', range: rg(0, 0, 3, 1), rows: ['Category'],
    values: [{ field: 'Amount', agg: 'sum' }, { field: 'Amount', agg: 'count', name: 'Count' }], top, left, area: rg(top, left, top + 4, left + 2),
    cellFmt: { rowHead: { bold: true }, 'data:0': { numFmt: '0.00' } },
    buttons: [
      { r: top, c: left, kind: 'page', field: 'Category' },
      { r: top + 2, c: left, kind: 'rows', field: 'Category' },
      { r: top + 3, c: left + 1, kind: 'toggle', field: 'Category', item: 'Group', collapsed: true },
    ],
  };
}
function fixture() {
  const pivot = definition('First', 10), extra = definition('Second', 25), cells = {};
  for (const def of [pivot, extra]) for (let r = def.area.r1; r <= def.area.r2; r++)
    for (let c = def.area.c1; c <= def.area.c2; c++) cells[`${r},${c}`] = { raw: `${def.name}:${r - def.top}:${c - def.left}` };
  return new Workbook({ sheets: [
    { name: 'Output', cells, pivot, pivotsExtra: [extra] },
    { name: 'Source', cells: { '0,0': { raw: 'Category' }, '0,1': { raw: 'Amount' }, '1,0': { raw: 'Group' }, '1,1': { raw: '3' }, '2,0': { raw: 'Other' }, '2,1': { raw: '7' }, '3,0': { raw: 'Last' }, '3,1': { raw: '2' } }, pivot: definition('Local', 12, 4) },
    { name: 'Other', cells: {}, pivot: definition('Remote', 30, 5) },
  ] });
}
const defs = w => [w.sheets[0].pivot, ...w.sheets[0].pivotsExtra].filter(Boolean);
const snap = w => JSON.parse(JSON.stringify(w.serialize()));
function assertOutput(w) {
  for (const def of defs(w)) {
    assert.equal(def.top, def.area.r1); assert.equal(def.left, def.area.c1);
    assert.equal(w.getValue(0, def.top, def.left), `${def.name}:0:0`);
    assert.equal(def.buttons[0].r, def.top); assert.equal(def.buttons[0].c, def.left);
    assert.equal(def.buttons[1].r, def.top + 2); assert.equal(def.buttons[2].r, def.top + 3); assert.equal(def.buttons[2].c, def.left + 1);
    assert.equal(new Set(def.buttons.map(b => `${b.r},${b.c}`)).size, def.buttons.length);
    assert.deepEqual(def.cellFmt, definition('Expected', 0).cellFmt);
  }
}
for (const [method, index, amount, expected] of [
  ['insertRows', 0, 3, [13, 28]], ['insertRows', 10, 2, [12, 27]],
  ['insertRows', 20, 4, [10, 29]], ['deleteRows', 0, 3, [7, 22]],
]) test(`${method} at ${index} moves primary and extra pivot cells, areas and buttons together`, () => {
  const w = fixture(), before = snap(w), remote = w.sheets[2].pivot;
  w.transact(() => w[method](0, index, amount));
  assert.deepEqual(defs(w).map(d => d.top), expected); assertOutput(w); assert.equal(w.sheets[2].pivot, remote);
  for (const def of defs(w)) assert.deepEqual(def.range, rg(0, 0, 3, 1));
  w.undo(); assert.deepEqual(snap(w), before); w.redo(); assertOutput(w);
});
for (const method of ['insertCols', 'deleteCols']) test(`${method} moves absolute header and toggle columns with output`, () => {
  const w = fixture(); w.transact(() => w[method](0, 0, 1));
  assert.deepEqual(defs(w).map(d => d.left), method === 'insertCols' ? [3, 3] : [1, 1]); assertOutput(w);
  w.undo(); assertOutput(w); w.redo(); assertOutput(w);
});
test('20 repeated inserts preserve one set of buttons and survive Undo/Redo and saved metadata', async () => {
  const w = fixture(), before = snap(w);
  for (let i = 0; i < 20; i++) w.transact(() => w.insertRows(0, 4, 1));
  assert.deepEqual(defs(w).map(d => d.top), [30, 45]); assertOutput(w);
  for (const copy of [new Workbook(w.serialize()), new Workbook(JSON.parse(await w.serializeBlob().text()))]) {
    assertOutput(copy); assert.deepEqual(defs(copy).map(d => d.area), defs(w).map(d => d.area));
  }
  // Native XLSX headers are reconstructed by the app; its stored output and location must still match.
  const native = new Workbook(readXlsx(writeXlsx(w)).data);
  for (const [i, def] of defs(native).entries()) {
    assert.deepEqual(def.area, defs(w)[i].area); assert.equal(def.top, defs(w)[i].top); assert.equal(def.left, defs(w)[i].left);
    assert.equal(native.getValue(0, def.top, def.left), `${def.name}:0:0`);
  }
  for (let i = 0; i < 20; i++) w.undo(); assert.deepEqual(snap(w), before);
  for (let i = 0; i < 20; i++) w.redo(); assertOutput(w);
});
test('source-sheet structural changes move local output and adjust all consumer source ranges', () => {
  const w = fixture(), before = snap(w); w.transact(() => w.insertRows(1, 2, 2));
  for (const sheet of w.sheets) for (const def of [sheet.pivot, ...(sheet.pivotsExtra ?? [])].filter(Boolean)) assert.deepEqual(def.range, rg(0, 0, 5, 1));
  assert.deepEqual(defs(w).map(d => d.top), [10, 25]); assert.equal(w.sheets[1].pivot.top, 14); assert.equal(w.sheets[1].pivot.buttons[1].r, 16);
  assert.equal(w.sheets[2].pivot.top, 30); assert.equal(w.sheets[2].pivot.buttons[1].r, 32); w.undo(); assert.deepEqual(snap(w), before);
});
test('direct model insertion inside output expands area and moves only trailing buttons', () => {
  const w = fixture(); w.transact(() => w.insertRows(0, 12, 2));
  const d = w.sheets[0].pivot; assert.equal(d.top, 10); assert.deepEqual(d.area, rg(10, 2, 16, 4));
  assert.deepEqual(d.buttons.map(b => b.r), [10, 14, 15]); assert.equal(w.getValue(0, 14, 2), 'First:2:0'); assert.equal(w.sheets[0].pivotsExtra[0].top, 27);
  w.undo(); assertOutput(w); w.redo(); assert.deepEqual(w.sheets[0].pivot.buttons.map(b => b.r), [10, 14, 15]);
});
test('direct model deletion removes buttons in deleted cells rather than duplicating them at the boundary', () => {
  const w = fixture(); w.transact(() => w.deleteRows(0, 12, 1));
  const d = w.sheets[0].pivot; assert.deepEqual(d.area, rg(10, 2, 13, 4));
  assert.deepEqual(d.buttons.map(b => [b.kind, b.r]), [['page', 10], ['toggle', 12]]); assert.equal(w.getValue(0, 12, 3), 'First:3:1'); w.undo(); assertOutput(w);
});
for (const [method, index, count] of [['deleteRows', 10, 5], ['deleteCols', 2, 3]]) test(`direct ${method} of a complete output removes the pivot definition and orphan buttons`, () => {
  const w = fixture(), before = snap(w); w.transact(() => w[method](0, index, count)); assert.equal(w.sheets[0].pivot, null);
  if (method === 'deleteRows') { assert.equal(w.sheets[0].pivotsExtra[0].top, 20); assertOutput(w); } else assert.deepEqual(w.sheets[0].pivotsExtra, []);
  w.undo(); assert.deepEqual(snap(w), before); w.redo(); assert.equal(w.sheets[0].pivot, null);
});
test('definitions without area still move anchors and buttons without changing role-based formats', () => {
  const w = fixture(); delete w.sheets[0].pivot.area; w.transact(() => w.insertRows(0, 2, 1));
  assert.equal(w.sheets[0].pivot.top, 11); assert.deepEqual(w.sheets[0].pivot.buttons.map(b => b.r), [11, 13, 14]); assert.equal(w.sheets[0].pivot.area, undefined);
  assert.deepEqual(w.sheets[0].pivot.cellFmt, definition('Expected', 0).cellFmt);
});
test('changes below all output keep existing output definitions and coordinates', () => {
  const w = fixture(), original = defs(w); w.transact(() => w.insertRows(0, 40, 2));
  assert.equal(w.sheets[0].pivot, original[0]); assert.equal(w.sheets[0].pivotsExtra[0], original[1]); assertOutput(w);
});

test('insertion immediately after an output moves only later pivots', () => {
  const w = fixture(), first = w.sheets[0].pivot; w.transact(() => w.insertRows(0, 15, 2));
  assert.equal(w.sheets[0].pivot, first); assert.equal(w.sheets[0].pivotsExtra[0].top, 27); assertOutput(w);
});

test('deleting a gap between outputs preserves the first pivot and shifts later output only', () => {
  const w = fixture(), first = w.sheets[0].pivot; w.transact(() => w.deleteRows(0, 18, 3));
  assert.equal(w.sheets[0].pivot, first); assert.equal(w.sheets[0].pivotsExtra[0].top, 22); assertOutput(w);
  w.undo(); assertOutput(w); w.redo(); assertOutput(w);
});

test('direct model deletion crossing the leading output edge clips area and drops deleted header buttons', () => {
  const w = fixture(); w.transact(() => w.deleteRows(0, 8, 4));
  const d = w.sheets[0].pivot; assert.equal(d.top, 8); assert.deepEqual(d.area, rg(8, 2, 10, 4));
  assert.deepEqual(d.buttons.map(b => [b.kind, b.r]), [['rows', 8], ['toggle', 9]]);
  assert.equal(w.getValue(0, 8, 2), 'First:2:0'); w.undo(); assertOutput(w);
});

test('direct column deletion drops deleted buttons and clips remaining output columns', () => {
  const w = fixture(); w.transact(() => w.deleteCols(0, 3, 1));
  const d = w.sheets[0].pivot; assert.equal(d.left, 2); assert.deepEqual(d.area, rg(10, 2, 14, 3));
  assert.deepEqual(d.buttons.map(b => b.kind), ['page', 'rows']);
  assert.equal(w.getValue(0, 13, 3), 'First:3:2'); w.undo(); assertOutput(w);
});

test('removing an extra output keeps primary and following extra definitions with Undo/Redo', () => {
  const w = fixture(), before = definition('Third', 40); w.sheets[0].pivotsExtra.push(before);
  w.transact(() => w.deleteRows(0, 25, 5));
  assert.equal(w.sheets[0].pivot.name, 'First'); assert.deepEqual(w.sheets[0].pivotsExtra.map(d => [d.name, d.top]), [['Third', 35]]);
  assert.deepEqual(w.sheets[0].pivotsExtra[0].buttons.map(b => b.r), [35, 37, 38]);
  w.undo(); assert.deepEqual(w.sheets[0].pivotsExtra.map(d => [d.name, d.top]), [['Second', 25], ['Third', 40]]);
  w.redo(); assert.deepEqual(w.sheets[0].pivotsExtra.map(d => [d.name, d.top]), [['Third', 35]]);
});

for (const anchor of ['top', 'left']) for (const missing of [undefined, null]) test(`${anchor} default zero follows structural insertion when ${missing === null ? 'null' : 'absent'}`, () => {
  const d = definition('Implicit', anchor === 'top' ? 0 : 5, anchor === 'left' ? 0 : 2);
  d[anchor] = missing;
  const initialRow = d.top ?? 0, initialCol = d.left ?? 0;
  const w = new Workbook({ sheets: [{ name: 'Output', cells: { [`${initialRow},${initialCol}`]: { raw: 'Header' } }, pivot: d }] });
  const method = anchor === 'top' ? 'insertRows' : 'insertCols';
  w.transact(() => w[method](0, 0, 2));
  const next = w.sheets[0].pivot; assert.equal(next[anchor], 2);
  assert.equal(anchor === 'top' ? next.area.r1 : next.area.c1, 2);
  assert.equal(anchor === 'top' ? next.buttons[0].r : next.buttons[0].c, 2);
  assert.equal(w.getValue(0, anchor === 'top' ? 2 : initialRow, anchor === 'left' ? 2 : initialCol), 'Header');
  w.undo(); assert.equal(w.sheets[0].pivot[anchor], missing); w.redo(); assert.equal(w.sheets[0].pivot[anchor], 2);
});

test('untouched implicit zero anchors are not unnecessarily materialized', () => {
  const d = definition('Implicit', 0, 0); delete d.top; delete d.left;
  const w = new Workbook({ sheets: [{ name: 'Output', cells: {}, pivot: d }] });
  w.transact(() => w.insertRows(0, 10, 2));
  assert.equal(w.sheets[0].pivot.top, undefined); assert.equal(w.sheets[0].pivot.left, undefined);
});
