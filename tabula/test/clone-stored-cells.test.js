import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap, getSharedBlankCell } from '../src/cellmap.js';
import { RunColumn } from '../src/run-column.js';
import { cloneStoredCells } from '../src/clone-stored-cells.js';
import { createImportedBlankRuns } from '../src/import-blank-runs.js';
import { shiftStoredCells } from '../src/cell-transforms.js';
import { Workbook, makeCellRC, cellData } from '../src/workbook.js';

const styles = [{ bold: true, fill: '#aabbcc' }, { italic: true, fill: '#112233' }];
function fixture(rows = 128, cols = 32) {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) collector.add(r, c, styles[r % 2]);
  collector.finish(); return cells;
}
const expanded = cells => [...cells];
function oracleShift(cells, axis, index, count, band = null) {
  const out = new CellMap();
  for (const [key, cell] of cells) {
    const [r, c] = key.split(',').map(Number), affected = !band || (axis === 'row' ? c : r) >= band[0] && (axis === 'row' ? c : r) <= band[1];
    const at = axis === 'row' ? r : c;
    if (affected && count < 0 && at >= index && at < index - count) continue;
    const moved = affected && at >= index ? at + count : at;
    out.setRC(axis === 'row' ? moved : r, axis === 'col' ? moved : c, cell);
  }
  return out;
}

function forbidExpanded(cells) {
  for (const key of ['entries', 'keys', 'values', Symbol.iterator]) cells[key] = () => { throw Error('expanded CellMap iterator forbidden'); };
  for (const column of cells.cols.values()) for (const key of ['entries', 'keys', 'values', Symbol.iterator]) column[key] = () => { throw Error('expanded RunColumn iterator forbidden'); };
}

test('one million alternating styled blanks clone without logical-cell enumeration and retain COW', () => {
  const cells = fixture(1024, 1024); forbidExpanded(cells);
  let calls = 0;
  const copy = cloneStoredCells(cells, () => { throw Error('blank-only patterns have no point'); }, cell => { calls++; return cell; });
  assert.equal(copy.size, 1_048_576); assert.equal(calls, 2); assert.equal(copy.col(0).runs.length, 1024); assert.equal(copy.col(0).runs, copy.col(1023).runs);
  copy.setRC(100, 0, { raw: '7' }); assert.equal(cells.getRC(100, 0).raw, ''); assert.equal(copy.getRC(100, 1023).raw, '');
  const constructed = new CellMap(cells); assert.equal(constructed.size, cells.size); assert.equal(constructed.col(0).runs, cells.col(0).runs);
  constructed.deleteRC(100, 2); assert.ok(cells.hasRC(100, 2)); assert.ok(constructed.hasRC(100, 3));
});

test('clone point mappers receive original coordinates and create independent candidate records', () => {
  const cells = fixture(8, 4), formula = makeCellRC({ raw: '=A1+1', cached: 4 }, 2, 1), link = { raw: 'link', link: '#A1' };
  cells.setRC(2, 1, formula); cells.setRC(6, 3, link); const visits = [];
  const copy = cloneStoredCells(cells, (cell, r, c) => { visits.push([r, c]); return { ...cell }; });
  assert.deepEqual(visits, [[2, 1], [6, 3]]); assert.notEqual(copy.getRC(2, 1), formula); assert.notEqual(copy.getRC(6, 3), link);
  assert.equal(copy.formulaAsts.get(formula.ast), 1); copy.deleteRC(2, 1); assert.equal(cells.formulaAsts.get(formula.ast), 1);
  assert.equal([...cells.formulaEntries()].length, 1); assert.equal([...copy.formulaEntries()].length, 0);
});

test('internal encoded snapshots preserve dirty cached trust, annotations and formula hydration', () => {
  const cells = fixture(12, 6), formula = makeCellRC({ raw: '=A1+2', cached: 9, comment: 'note', link: '#A1', style: { numFmt: 'text' }, fx: true }, 4, 2);
  formula.dirty = true; cells.setRC(4, 2, formula);
  const encoded = cloneStoredCells(cells, cell => cellData(cell), cell => cell);
  const saved = encoded.getRC(4, 2); assert.equal(saved.staleCached, 9); assert.equal(saved.cached, undefined); assert.equal(saved.fx, true); assert.equal(saved.comment, 'note'); assert.equal(saved.link, '#A1');
  const hydrated = cloneStoredCells(encoded, (cell, r, c) => makeCellRC(cell, r, c), cell => makeCellRC(cell));
  assert.equal(hydrated.getRC(4, 2).dirty, true); assert.equal(hydrated.getRC(4, 2).cached, 9); assert.equal(hydrated.getRC(4, 2).formula, true);
  assert.notEqual(hydrated.getRC(4, 2), formula); assert.equal([...hydrated.formulaEntries()].length, 1); assert.equal(hydrated.col(0).runs, hydrated.col(5).runs);
});

test('clone blank removal and canonical new styles preserve the source and share only immutable values', () => {
  const cells = fixture(16, 8), style = { underline: true };
  const copy = cloneStoredCells(cells, cell => ({ ...cell }), cell => ({ raw: '', style }));
  assert.ok(Object.isFrozen(copy.getRC(2, 5))); assert.equal(copy.getRC(2, 5).style, style); assert.equal(cells.getRC(2, 5).style, styles[0]);
  assert.equal(copy.col(0).runs, copy.col(7).runs); assert.equal(copy.col(0).runs.length, 1);
  const removed = cloneStoredCells(cells, cell => cell, () => null); assert.equal(removed.size, 0); assert.equal(removed.cols.size, 0); assert.equal(cells.size, 128);
  assert.throws(() => cloneStoredCells(cells, cell => cell, () => ({ raw: 'mutable nonblank' })), /불변 빈 셀/);
});

test('whole row insertion/deletion transforms one shared pattern and preserves all blank positions', () => {
  const cells = fixture(1024, 1024); forbidExpanded(cells);
  const representative = cells.col(0), originalStorage = representative.storageEntries;
  let reads = 0; representative.storageEntries = function* (...args) { reads++; yield* originalStorage.apply(this, args); };
  for (const column of [...cells.cols.values()].slice(1)) column.storageEntries = () => { throw Error('repeated pattern traversal forbidden'); };
  const inserted = shiftStoredCells(cells, 'row', 400, 3);
  assert.equal(reads, 1); assert.equal(inserted.size, cells.size); assert.equal(inserted.col(0).runs, inserted.col(1023).runs);
  assert.equal(inserted.hasRC(400, 0), false); assert.equal(inserted.getRC(403, 1023).style, styles[0]); assert.equal(cells.getRC(400, 0).style, styles[0]);
  const deleted = shiftStoredCells(inserted, 'row', 403, -7);
  assert.equal(deleted.size, 1_048_576 - 7 * 1024); assert.equal(deleted.col(0).runs, deleted.col(1023).runs);
  deleted.setRC(410, 2, { raw: 'edited' }); assert.equal(deleted.getRC(410, 3).raw, ''); assert.equal(inserted.getRC(410, 2).raw, '');
});

test('whole column insertion/deletion reuses untouched row backing without pattern traversal', () => {
  const cells = fixture(256, 96); forbidExpanded(cells);
  for (const column of cells.cols.values()) column.storageEntries = () => { throw Error('whole column shift must not traverse blank runs'); };
  const inserted = shiftStoredCells(cells, 'col', 30, 4); assert.equal(inserted.size, cells.size); assert.equal(inserted.hasRC(0, 30), false); assert.equal(inserted.getRC(100, 34).style, styles[0]);
  assert.equal(inserted.col(0).runs, cells.col(0).runs); assert.equal(inserted.col(99).runs, cells.col(95).runs);
  const deleted = shiftStoredCells(inserted, 'col', 29, -9); assert.equal(deleted.size, cells.size - 5 * 256); assert.equal(deleted.getRC(0, 29).style, styles[0]);
});

test('whole shifts and partial bands preserve native insertion order and mixed-cell candidate indexes', () => {
  const cells = fixture(22, 8); cells.deleteRC(3, 0); cells.setRC(3, 0, getSharedBlankCell(styles[1]));
  const formula = makeCellRC({ raw: '=A1' }, 10, 4); cells.setRC(10, 4, formula); cells.setRC(8, 2, { raw: 'link', link: '#A1' });
  for (const [axis, index, count, band] of [['row',5,3,null],['row',8,-4,null],['col',3,2,null],['col',4,-2,null],['row',6,2,[2,5]],['col',4,-1,[6,15]]]) {
    const actual = shiftStoredCells(cells, axis, index, count, band), expected = oracleShift(cells, axis, index, count, band);
    assert.deepEqual(expanded(actual), expanded(expected)); assert.equal(actual.size, expected.size); assert.deepEqual([...actual.formulaEntries()], [...expected.formulaEntries()]);
  }
});

test('plain Map inputs and independent default point cloning retain original stored order', () => {
  const raw = new Map([['8,2',{raw:'8'}],['2,0',{raw:'2'}],['1,2',{raw:'1'}]]), visits = [];
  const copy = cloneStoredCells(raw, (cell,r,c) => { visits.push([r,c]); return { ...cell }; });
  assert.deepEqual(visits, [[8,2],[2,0],[1,2]]); assert.deepEqual([...copy].map(([k]) => k), ['8,2','1,2','2,0']);
  copy.getRC(8,2).raw = 'different'; assert.equal(raw.get('8,2').raw, '8');
});
test('large worksheet structure history avoids public JSON serialization and expanded iterators through Undo/Redo', () => {
  const wb = new Workbook({ sheets: [{ name: 'Synthetic', cells: fixture(1024, 1024) }] });
  const originalCellEntries = CellMap.prototype.entries, originalColumnEntries = RunColumn.prototype.entries;
  const originalSerializeSheet = wb.serializeSheet;
  CellMap.prototype.entries = function* () { throw Error('history expanded CellMap entries'); };
  RunColumn.prototype.entries = function* () { throw Error('history expanded RunColumn entries'); };
  wb.serializeSheet = () => { throw Error('structure history used public JSON serializer'); };
  try {
    const check = () => { assert.equal(wb.sheets[0].cells.size, 1_048_576); assert.equal(wb.sheets[0].cells.col(0).runs, wb.sheets[0].cells.col(1023).runs); };
    for (let i = 0; i < 3; i++) {
      wb.transact(() => wb.insertRows(0, 400, 3)); check();
      const entry = wb.undoStack.at(-1).entries.find(entry => entry.t === 'sheet');
      assert.ok(entry.before.cells instanceof CellMap); assert.ok(entry.after.cells instanceof CellMap);
      assert.notEqual(entry.after.cells, wb.sheets[0].cells); assert.equal(entry.before.cells.size, 1_048_576); assert.equal(entry.after.cells.size, 1_048_576);
      assert.equal(entry.after.cells.col(0).runs, entry.after.cells.col(1023).runs);
    }
    assert.equal(wb.undoStack.length, 3); assert.equal(wb.getCell(0, 400, 0), undefined); assert.equal(wb.styleAt(0, 409, 1023).fill, styles[0].fill);
    for (let i = 0; i < 3; i++) { wb.undo(); check(); }
    assert.equal(wb.styleAt(0, 400, 1023).fill, styles[0].fill); assert.equal(wb.redoStack.length, 3);
    for (let i = 0; i < 3; i++) { wb.redo(); check(); }
    const snapshot = wb.undoStack.at(-1).entries.find(entry => entry.t === 'sheet').after.cells;
    wb.setInput(0, 409, 0, '77'); assert.equal(wb.getValue(0, 409, 0), 77); assert.equal(snapshot.getRC(409, 0).raw, ''); assert.equal(wb.getValue(0, 409, 1023), null);
  } finally { CellMap.prototype.entries = originalCellEntries; RunColumn.prototype.entries = originalColumnEntries; wb.serializeSheet = originalSerializeSheet; }
});

test('structure memory history retains dirty formula trust, annotations and later independent style edits', () => {
  const cells = fixture(32, 8), formula = makeCellRC({ raw: '=A1+2', cached: 9, comment: 'note', link: '#A1', style: { bold: true, numFmt: 'text' }, fx: true }, 10, 0);
  formula.dirty = true; cells.setRC(10, 0, formula);
  const wb = new Workbook({ sheets: [{ name: 'Synthetic', cells }] }); wb.getCell(0, 10, 0).dirty = true; wb.transact(() => wb.insertRows(0, 5, 2));
  const entry = wb.undoStack.at(-1).entries.find(entry => entry.t === 'sheet'), stored = entry.before.cells.getRC(10, 0);
  assert.equal(stored.staleCached, 9); assert.equal(stored.cached, undefined); assert.equal(stored.comment, 'note'); assert.equal(stored.link, '#A1'); assert.equal(stored.fx, true);
  wb.undo(); const restored = wb.getCell(0, 10, 0); assert.equal(restored.formula, true); assert.equal(restored.dirty, true); assert.equal(restored.cached, 9); assert.equal(restored.comment, 'note');
  assert.equal(restored.link, '#A1'); assert.equal(restored.fx, true); assert.notEqual(restored, stored);
  wb.redo(); assert.equal(wb.getCell(0, 12, 0).comment, 'note'); assert.equal(wb.getCell(0, 12, 0).dirty, true);
  wb.transact(() => wb.setInput(0, 20, 3, '5')); wb.undo(); assert.equal(wb.getValue(0, 20, 3), null); assert.equal(entry.after.cells.getRC(20, 3).raw, '');
});

test('column structure history remains exact across Undo/Redo with shared blank backing', () => {
  const wb = new Workbook({ sheets: [{ name: 'Synthetic', cells: fixture(64, 80) }] });
  wb.transact(() => wb.deleteCols(0, 20, 5)); assert.equal(wb.sheets[0].cells.size, 64 * 75); assert.equal(wb.sheets[0].cells.col(0).runs, wb.sheets[0].cells.col(74).runs);
  wb.undo(); assert.equal(wb.sheets[0].cells.size, 64 * 80); assert.equal(wb.styleAt(0, 63, 79).fill, styles[1].fill);
  wb.redo(); assert.equal(wb.sheets[0].cells.size, 64 * 75); assert.equal(wb.sheets[0].cells.col(0).runs, wb.sheets[0].cells.col(74).runs);
  wb.transact(() => wb.insertCols(0, 10, 3)); assert.equal(wb.getCell(0, 0, 10), undefined); assert.equal(wb.styleAt(0, 20, 13).fill, styles[0].fill);
  wb.undo(); assert.equal(wb.styleAt(0, 20, 10).fill, styles[0].fill); wb.redo(); assert.equal(wb.getCell(0, 0, 10), undefined);
});

test('even frozen annotated/formula objects cannot be introduced through the blank-only mapper', () => {
  const cells = fixture(4, 2); cells.setRC(1, 0, { raw: 'point' });
  for (const invalid of [Object.freeze({ raw: '=A1', style: styles[0] }), Object.freeze({ raw: '', style: styles[0], link: '#A1' })]) {
    assert.throws(() => cloneStoredCells(cells, cell => ({ ...cell }), () => invalid), /불변 빈 셀/);
    assert.equal(cells.getRC(0, 0).raw, ''); assert.equal(cells.getRC(1, 0).raw, 'point'); assert.equal(cells.size, 8);
  }
});
