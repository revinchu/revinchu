import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { filterButtonVisible, filterButtonsVisible, filterWithButtons } from '../src/filter-display.js';

const table = () => ({ id: 't', name: 'Report', r1: 1, c1: 2, r2: 5, c2: 4, header: true,
  filter: { criteria: { 2: ['서울'], 4: { type: 'custom', op1: 'ge', v1: '20' } }, hidden: { 2: true, 4: true }, sort: { col: 4, asc: false } } });
const book = () => new Workbook({ sheets: [{ name: '자료', fileValues: true,
  cells: { '0,0': { raw: '=2+2', cached: 42 } }, tables: [table()] }] });

test('button visibility changes no criteria, hidden rows or sort and does not mutate source', () => {
  const original = table().filter, copy = structuredClone(original);
  const off = filterWithButtons(original, 2, 4, false);
  assert.equal(filterButtonsVisible(off, 2, 4), false);
  assert.equal(filterButtonVisible(off, 3), false);
  assert.deepEqual(original, copy);
  assert.deepEqual(filterWithButtons(off, 2, 4, true), original);
  assert.equal(filterButtonsVisible({ ...off, hiddenButtons: { 2: true, 4: true } }, 2, 4), true);
});

test('table arrow toggle and undo redo preserve imported caches and exact filter state', () => {
  const wb = book(), before = structuredClone(wb.sheets[0].tables[0].filter);
  wb.transact(() => wb.setTableFilterButtons(0, 't', false));
  assert.equal(wb.getValue(0, 0, 0), 42);
  assert.deepEqual(wb.sheets[0].tables[0].filter.criteria, before.criteria);
  assert.deepEqual(wb.sheets[0].tables[0].filter.hidden, before.hidden);
  assert.deepEqual(wb.sheets[0].tables[0].filter.sort, before.sort);
  const next = new Workbook(wb.serialize());
  assert.equal(filterButtonsVisible(next.sheets[0].tables[0].filter, 2, 4), false);
  wb.undo(); assert.deepEqual(wb.sheets[0].tables[0].filter, before); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.redo(); assert.equal(filterButtonsVisible(wb.sheets[0].tables[0].filter, 2, 4), false); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.transact(() => wb.setTableFilterButtons(0, 't', false)); assert.equal(wb.undoStack.length, 1);
});

test('structural changes in either transaction order must still invalidate calculations', () => {
  for (const first of [true, false]) {
    const wb = book();
    const structural = () => wb.setSheetProp(0, 'tables', wb.sheets[0].tables.map(t => ({ ...t, r2: 6 })));
    wb.transact(() => { if (first) structural(); wb.setTableFilterButtons(0, 't', false); if (!first) structural(); });
    assert.equal(wb.getValue(0, 0, 0), 4);
    assert.equal(wb.undoStack.at(-1).entries.find(e => e.prop === 'tables').calcNeutral, undefined);
  }
});

test('all hidden arrows, criteria and sort follow inserted and deleted table columns', () => {
  const wb = book();
  wb.transact(() => wb.setTableFilterButtons(0, 't', false));
  wb.transact(() => wb.insertCols(0, 3));
  let t = wb.sheets[0].tables[0];
  assert.deepEqual(t.filter.hiddenButtons, { 2: true, 3: true, 4: true, 5: true });
  assert.deepEqual(t.filter.criteria[5], { type: 'custom', op1: 'ge', v1: '20' });
  assert.deepEqual(t.filter.sort, { col: 5, asc: false });
  wb.transact(() => wb.deleteCols(0, 3)); t = wb.sheets[0].tables[0];
  assert.deepEqual(t.filter.hiddenButtons, { 2: true, 3: true, 4: true });
  assert.deepEqual(t.filter.sort, { col: 4, asc: false });
  wb.transact(() => wb.deleteCols(0, 4)); assert.equal(wb.sheets[0].tables[0].filter.sort, undefined);
});

test('sheet filters preserve mixed arrow visibility and imported multiple-sort refs through structure changes', () => {
  const wb = book(), sort = { col: 4, asc: true, xlsx: { name: 'sortState', attrs: { ref: 'C3:E6' }, children: [
    { name: 'sortCondition', attrs: { ref: 'E3:E6' }, children: [] }, { name: 'sortCondition', attrs: { ref: 'C3:C6' }, children: [] } ] } };
  wb.sheets[0].filter = { r1: 1, c1: 2, r2: 5, c2: 4, criteria: {}, hidden: {}, hiddenButtons: { 2: true }, sort };
  wb.transact(() => wb.insertCols(0, 3));
  let f = wb.sheets[0].filter;
  assert.deepEqual(f.hiddenButtons, { 2: true }); assert.equal(f.sort.col, 5);
  assert.equal(f.sort.xlsx.attrs.ref, 'C3:F6'); assert.equal(f.sort.xlsx.children[0].attrs.ref, 'F3:F6');
  wb.transact(() => wb.insertRows(0, 3)); f = wb.sheets[0].filter;
  assert.equal(f.sort.xlsx.attrs.ref, 'C3:F7');
});

test('saved total formulas track off-sheet edits, moves, renames and undo even with no live formula dependencies', () => {
  const wb = new Workbook({ sheets: [ { name: '원본', cells: { '2,0': { raw: '5' } } },
    { name: '표', cells: {}, tables: [{ ...table(), totals: false, totalsFns: { 4: 'custom' },
      totalsCells: { 2: { raw: '내 합계' }, 4: { raw: "='원본'!A3+7", cached: 12 } } }] } ] });
  wb.transact(() => wb.insertRows(0, 1));
  assert.equal(wb.sheets[1].tables[0].totalsCells[4].raw, "='원본'!A4+7");
  assert.equal(wb.sheets[1].tables[0].totalsCells[4].cached, undefined);
  wb.undo(); assert.equal(wb.sheets[1].tables[0].totalsCells[4].raw, "='원본'!A3+7");
  wb.transact(() => wb.renameSheet(0, '바뀐 원본'));
  assert.equal(wb.sheets[1].tables[0].totalsCells[4].raw, "='바뀐 원본'!A3+7");
  wb.transact(() => wb.moveRange(0, { r1: 2, c1: 0, r2: 2, c2: 0 }, 2, 1));
  assert.equal(wb.sheets[1].tables[0].totalsCells[4].raw, "='바뀐 원본'!B5+7");
  wb.transact(() => wb.insertCols(1, 3));
  assert.equal(wb.sheets[1].tables[0].totalsCells[5].raw, "='바뀐 원본'!B5+7");
  assert.equal(wb.sheets[1].tables[0].totalsFns[5], 'custom');
});

test('deferred totals honor literal text formatting and explicitly forced formulas', () => {
  const wb = book(); wb.sheets[0].tables[0].totalsCells = {
    2: { raw: '=資料!A1', style: { numFmt: 'text' } },
    3: { raw: '=資料!A1', inputType: 'text', fx: true },
    4: { raw: '=資料!A1', inputType: 'value', style: { numFmt: 'text' } },
  };
  wb.transact(() => wb.rewriteTableTotals(raw => raw.replace('資料', '新資料')));
  const cells = wb.sheets[0].tables[0].totalsCells;
  assert.equal(cells[2].raw, '=資料!A1');
  assert.equal(cells[3].raw, '=新資料!A1'); assert.equal(cells[4].raw, '=新資料!A1');
});
