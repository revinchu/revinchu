import test from 'node:test';
import assert from 'node:assert/strict';
import { cellDropdownTarget, altArrowDownKey } from '../src/cell-dropdown-target.js';
import { Workbook } from '../src/workbook.js';

const filterTarget = (column, key = '') => ({ kind: 'filter', column, key });
const pivotTarget = (pivotIndex, buttonKind, field) => ({ kind: 'pivot', pivotIndex, buttonKind, field });
const table = (id = 'first') => ({ id, name: 'SyntheticTable', r1: 2, c1: 3, r2: 8, c2: 5, header: true, filter: {} });
const pivot = () => ({ rows: ['Group'], cols: ['Period'], pages: ['Region'], buttons: [
  { r: 30, c: 4, kind: 'rows' }, { r: 30, c: 5, kind: 'cols', field: 'Period' }, { r: 28, c: 5, kind: 'page', field: 'Region' },
] });

test('ordinary filters target only visible header buttons inside the exact column bounds', () => {
  const sheet = { filter: { r1: 3, c1: 5, r2: 12, c2: 7, criteria: { 5: ['A'] }, hidden: { 4: true }, hiddenButtons: { 6: true } } };
  assert.deepEqual(cellDropdownTarget(sheet, 3, 5), filterTarget(5));
  assert.deepEqual(cellDropdownTarget(sheet, 3, 7), filterTarget(7));
  for (const [r, c] of [[3, 4], [3, 8], [4, 5], [12, 5], [3, 6]]) assert.equal(cellDropdownTarget(sheet, r, c), null);
});

test('table headers use tableAt first-match precedence over overlapping tables and sheet filters', () => {
  const first = table(), second = table('second'), sheet = { tables: [first, second], filter: { r1: 2, c1: 0, r2: 12, c2: 8 } };
  assert.deepEqual(cellDropdownTarget(sheet, 2, 3), filterTarget(3, 'first'));
  assert.deepEqual(cellDropdownTarget(sheet, 2, 1), filterTarget(1));
  first.filter.hiddenButtons = { 3: true };
  assert.equal(cellDropdownTarget(sheet, 2, 3), null);
  first.filter = null;
  assert.deepEqual(cellDropdownTarget(sheet, 2, 3), filterTarget(3));
  delete sheet.filter;
  assert.equal(cellDropdownTarget(sheet, 2, 3), null);
});

test('table body, totals and headerless tables do not become filter targets', () => {
  for (const row of [3, 8]) assert.equal(cellDropdownTarget({ tables: [table()] }, row, 3), null);
  assert.equal(cellDropdownTarget({ tables: [{ ...table(), header: false }] }, 2, 3), null);
  assert.equal(cellDropdownTarget({ tables: [{ ...table(), filter: null }] }, 2, 3), null);
  assert.deepEqual(cellDropdownTarget({ tables: [{ ...table(), filter: { hiddenButtons: { 0: true } } }] }, 2, 3), filterTarget(3, 'first'));
});

test('primary pivot row, column and page button cells preserve the mouse menu field choices', () => {
  const sheet = { pivot: pivot() };
  assert.deepEqual(cellDropdownTarget(sheet, 30, 4), pivotTarget(0, 'rows'));
  assert.deepEqual(cellDropdownTarget(sheet, 30, 5), pivotTarget(0, 'cols', 'Period'));
  assert.deepEqual(cellDropdownTarget(sheet, 28, 5), pivotTarget(0, 'page', 'Region'));
  for (const [r, c] of [[31, 4], [30, 3], [29, 5]]) assert.equal(cellDropdownTarget(sheet, r, c), null);
});

test('extra pivot indices match rendered buttons with absent or sparse primary definitions', () => {
  const extra = pivot(), first = { rows: ['Earlier'], buttons: [{ r: 1, c: 1, kind: 'rows' }] };
  assert.deepEqual(cellDropdownTarget({ pivotsExtra: [null, first, undefined, extra] }, 30, 4), pivotTarget(1, 'rows'));
  assert.deepEqual(cellDropdownTarget({ pivot: first, pivotsExtra: [null, extra] }, 30, 4), pivotTarget(1, 'rows'));
  assert.deepEqual(cellDropdownTarget({ pivot: null, pivotsExtra: [extra] }, 30, 4), pivotTarget(0, 'rows'));
});

test('hidden pivot headers exclude row and column menus while retaining report page filters', () => {
  for (const patch of [{ showHeaders: false }, { fieldCaptions: false }, { showHeaders: false, fieldCaptions: false }]) {
    const sheet = { pivot: { ...pivot(), ...patch } };
    assert.equal(cellDropdownTarget(sheet, 30, 4), null);
    assert.equal(cellDropdownTarget(sheet, 30, 5), null);
    assert.deepEqual(cellDropdownTarget(sheet, 28, 5), pivotTarget(0, 'page', 'Region'));
  }
});

test('sigma buttons require real axis fields; toggle, values and unbound page buttons are not dropdowns', () => {
  for (const kind of ['rows', 'cols']) {
    const def = { [kind]: ['Dimension'], buttons: [{ r: 0, c: 0, kind, sigma: true }] };
    assert.deepEqual(cellDropdownTarget({ pivot: def }, 0, 0), pivotTarget(0, kind));
    def[kind] = [];
    assert.equal(cellDropdownTarget({ pivot: def }, 0, 0), null);
  }
  for (const button of [{ kind: 'toggle', field: 'Group' }, { kind: 'values', field: 'Measure' }, { kind: 'page' }, { kind: 'rows' }])
    assert.equal(cellDropdownTarget({ pivot: { rows: [], pages: ['Region'], buttons: [{ r: 0, c: 0, ...button }] } }, 0, 0), null);
});

test('virtualized targets require no DOM, cell reads, pivot computation or model mutation', () => {
  const button = Object.freeze({ r: 900, c: 140, kind: 'cols', field: 'Period' });
  const def = Object.freeze({ cols: Object.freeze(['Period']), buttons: Object.freeze([button]) });
  const sheet = Object.freeze({ pivot: def, get cells() { throw new Error('Cell reads are unnecessary'); } });
  assert.deepEqual(cellDropdownTarget(sheet, 900, 140), pivotTarget(0, 'cols', 'Period'));
  for (const [r, c] of [[-1, 0], [0, -1], [0.5, 0], [0, NaN], [Infinity, 0]]) assert.equal(cellDropdownTarget(sheet, r, c), null);
  assert.equal(cellDropdownTarget(null, 0, 0), null);
});

test('model targets follow row and column insertion and undo redo without remembered DOM coordinates', () => {
  const wb = new Workbook({ sheets: [{ name: 'Synthetic', cells: {}, filter: { r1: 2, c1: 1, r2: 9, c2: 2 },
    tables: [{ ...table(), r1: 20, r2: 25, c1: 10, c2: 11 }],
    pivot: { ...pivot(), top: 30, left: 4, area: { r1: 28, c1: 4, r2: 35, c2: 6 } },
  }] });
  const check = (dr, dc) => {
    const sheet = wb.sheets[0];
    assert.deepEqual(cellDropdownTarget(sheet, 2 + dr, 1 + dc), filterTarget(1 + dc));
    assert.deepEqual(cellDropdownTarget(sheet, 20 + dr, 10 + dc), filterTarget(10 + dc, 'first'));
    assert.deepEqual(cellDropdownTarget(sheet, 30 + dr, 4 + dc), pivotTarget(0, 'rows'));
  };
  check(0, 0);
  wb.transact(() => { wb.insertRows(0, 1, 3); wb.insertCols(0, 1, 2); });
  check(3, 2);
  assert.equal(cellDropdownTarget(wb.sheets[0], 30, 4), null);
  wb.undo(); check(0, 0);
  wb.redo(); check(3, 2);
});

test('Alt ArrowDown uses the physical key even when an idle IME reports Process or 229', () => {
  for (const key of ['ArrowDown', 'Process', 'Unidentified']) for (const keyCode of [40, 229])
    assert.equal(altArrowDownKey({ altKey: true, code: 'ArrowDown', key, keyCode }), true);
  assert.equal(altArrowDownKey({ altKey: true, shiftKey: true, code: 'ArrowDown', key: 'ArrowDown' }), true);
});

test('Alt ArrowDown never owns Ctrl, Meta, AltGr, active composition or another physical key', () => {
  for (const extra of [{ ctrlKey: true }, { metaKey: true }, { isComposing: true }, { altKey: false }, { getModifierState: key => key === 'AltGraph' }, { code: 'ArrowUp' }])
    assert.equal(altArrowDownKey({ altKey: true, code: 'ArrowDown', key: 'ArrowDown', ...extra }), false);
  assert.equal(altArrowDownKey(null), false);
});

test('legacy ArrowDown fallback does not infer an IME key without a physical code', () => {
  for (const code of [undefined, '', 'Unidentified']) {
    assert.equal(altArrowDownKey({ altKey: true, code, key: 'ArrowDown', keyCode: 40 }), true);
    assert.equal(altArrowDownKey({ altKey: true, code, key: 'ArrowDown', keyCode: 229 }), false);
    for (const key of ['Process', 'Unidentified', 'Dead', 'ArrowUp']) assert.equal(altArrowDownKey({ altKey: true, code, key, keyCode: 40 }), false);
  }
});
