import { test } from 'node:test';
import assert from 'node:assert/strict';
import { CellMap, getSharedBlankCell } from '../src/cellmap.js';
import { Workbook, makeCellRC } from '../src/workbook.js';
import { shiftStoredCells, moveStoredCells } from '../src/cell-transforms.js';
import { markPreparedWorkbook } from '../src/prepared-sheet-data.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const formula = (raw = '=Source!A1', r = 0, c = 0) => makeCellRC({ raw }, r, c);
const asts = cells => { const result = []; cells.forEachFormulaAST(ast => result.push(ast)); return result; };
function matchesStorage(cells) {
  const expected = new Map();
  for (const [, , cell] of cells.storageEntries()) if (cell.formula) {
    const ast = cell.ast || null; expected.set(ast, (expected.get(ast) ?? 0) + 1);
  }
  assert.equal(cells.formulaAsts.size, expected.size);
  for (const [ast, count] of expected) assert.equal(cells.formulaAsts.get(ast), count);
  assert.deepEqual(new Set(asts(cells)), new Set(expected.keys()));
}
const dependencyShape = deps => deps.map(d => ({ all: d.all, sheets: [...d.sheets].sort((a, b) => a - b) }));
function matchesDependencies(wb) {
  const actual = wb.sheetDeps();
  const scanned = wb.sheets.map((sheet, si) => {
    const d = { all: false, sheets: new Set() }, seen = new Set();
    sheet.cells.forEachFormulaRC(cell => {
      if (!cell.ast) d.all = true;
      else if (!seen.has(cell.ast)) { seen.add(cell.ast); wb.addDeps(d, si, cell.ast); }
    });
    matchesStorage(sheet.cells); return d;
  });
  assert.deepEqual(dependencyShape(actual), dependencyShape(scanned));
}
function fixture() {
  return new Workbook({ sheets: [
    { name: 'Source', cells: { '4,0': { raw: '7' } } },
    { name: 'Report', cells: { '0,0': { raw: '=Source!A5' }, '1,0': { raw: '=Source!A6' }, '2,1': { raw: 'link', link: '#Source!A5' } } },
    { name: 'Other', cells: { '0,0': { raw: '8' } } },
  ] });
}

test('shared AST summary retains duplicates until the final coordinate is removed', () => {
  const cells = new CellMap(), a = formula(), b = formula('=Other!A1');
  cells.setRC(0, 0, a); cells.setRC(1, 0, a); cells.setRC(0, 2, { ...a, link: '#Source!A1' });
  assert.deepEqual(asts(cells), [a.ast]); assert.equal(cells.formulaAsts.get(a.ast), 3);
  cells.setRC(1, 0, b); cells.deleteRC(0, 0); matchesStorage(cells);
  assert.equal(cells.formulaAsts.get(a.ast), 1);
  cells.setRC(0, 2, { raw: 'link', link: '#Source!A1' });
  assert.deepEqual(asts(cells), [b.ast]); matchesStorage(cells);
  cells.deleteRC(1, 0); assert.deepEqual(asts(cells), []);
});

test('raw imports, text formulas, forced formulas, links and parse failures keep distinct semantics', () => {
  const cells = new CellMap({
    '0,0': { raw: '=Source!A1' }, '1,0': { raw: '=Source!A1', inputType: 'text' },
    '2,0': { raw: '=Source!A1', style: { numFmt: 'text' }, fx: true },
    '3,0': { raw: '=SUM(' }, '4,0': { raw: 'link', link: '#Source!A1' },
    '5,0': { raw: "'=Source!A1" },
  });
  assert.deepEqual(asts(cells), []);
  for (const count of cells.mapValues((v, r, c) => makeCellRC(v, r, c), v => makeCellRC(v))) assert.ok(count > 0);
  matchesStorage(cells); assert.equal(cells.formulaAsts.get(null), 1);
  assert.equal([...cells.formulaEntries()].length, 3); assert.equal(asts(cells).length, 2);
  cells.deleteRC(3, 0); assert.ok(!asts(cells).includes(null)); matchesStorage(cells);
});

test('same-object replacement and in-place normalization subtract the recorded previous AST', () => {
  const cells = new CellMap(), a = formula(), b = formula('=Other!A1');
  cells.setRC(0, 0, a); a.ast = b.ast; cells.setRC(0, 0, a);
  assert.deepEqual(asts(cells), [b.ast]); matchesStorage(cells);
  cells.setRC(1, 0, a);
  const replacement = formula('=Third!A1').ast;
  for (const count of cells.mapValues(cell => { cell.ast = replacement; return cell; }, v => v)) assert.ok(count > 0);
  assert.equal(cells.formulaAsts.get(replacement), 2); matchesStorage(cells);
  for (const count of cells.mapValues(cell => { delete cell.formula; delete cell.ast; cell.raw = 'plain'; return cell; }, v => v)) assert.ok(count > 0);
  assert.deepEqual(asts(cells), []); matchesStorage(cells);
});

test('compact blank overwrite, normalization removal, rejected runs and clear release AST references', () => {
  const cells = new CellMap(), blank = getSharedBlankCell({ bold: true }), a = formula(), b = formula('=Other!A1');
  cells.setRunRC(0, 0, 200000, blank); cells.setRC(300, 0, a); cells.setRC(500, 0, a); cells.setRC(700, 0, b);
  assert.throws(() => cells.setRunRC(500, 0, 2, a), /빈 셀 범위/); matchesStorage(cells);
  cells.setRunRC(250, 0, 100, blank); assert.equal(cells.formulaAsts.get(a.ast), 1);
  cells.setRunRC(450, 0, 100, blank); assert.ok(!cells.formulaAsts.has(a.ast)); matchesStorage(cells);
  for (const count of cells.mapValues(() => null, () => null)) assert.ok(count > 0);
  assert.equal(cells.size, 0); assert.deepEqual(asts(cells), []);
  cells.setRC(0, 0, b); cells.clear(); assert.equal(cells.formulaAsts.size, 0);
});

test('random writes, deletions, run overwrites and re-normalization match a complete scan', () => {
  const cells = new CellMap(), blank = getSharedBlankCell({ italic: true }), formulas = [formula(), formula('=Other!A1'), formula('=SUM(')];
  let seed = 507; const random = n => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) % n);
  for (let i = 0; i < 600; i++) {
    const r = random(80), c = random(5), op = random(6);
    if (!op) cells.deleteRC(r, c);
    else if (op === 1) cells.setRunRC(r, c, random(15) + 1, blank);
    else cells.setRC(r, c, op < 5 ? formulas[op - 2] : { raw: 'link', link: '#Source!A1' });
    if (i % 20 === 0) matchesStorage(cells);
  }
  for (const count of cells.mapValues((v, r, c) => makeCellRC(v, r, c), v => v)) assert.ok(count > 0);
  matchesStorage(cells);
});

test('row/column shifts, band moves and copies rebuild shared-AST counts', () => {
  const cells = new CellMap(), a = formula(); cells.setRC(5, 0, a); cells.setRC(7, 1, a); cells.setRC(9, 1, formula('=Other!A1'));
  for (const [axis, at, count] of [['row', 4, 3], ['row', 5, -3], ['col', 1, 2], ['col', 0, -1]]) {
    const shifted = shiftStoredCells(cells, axis, at, count); matchesStorage(shifted); matchesStorage(new CellMap(shifted));
  }
  matchesStorage(moveStoredCells(cells, { r1: 5, c1: 0, r2: 7, c2: 1 }, 3, 4)); matchesStorage(cells);
});

test('cold dependencies inspect one shared AST for one million formulas without per-cell iteration', () => {
  const cells = new CellMap(), cell = formula();
  for (let r = 0; r < 1000000; r++) cells.setRC(r, 0, cell);
  const data = { sheets: [{ name: 'Source', cells: new CellMap(), blocks: [] }, { name: 'Report', cells, blocks: [] }] };
  markPreparedWorkbook(data, [false, false]); const wb = new Workbook(data);
  assert.equal(cells.size, 1000000); assert.equal(cells.formulaAsts.get(cell.ast), 1000000);
  cells.formulaEntries = () => { throw Error('per-cell dependency iterator'); };
  cells.forEachFormulaRC = () => { throw Error('per-cell dependency callback'); };
  let visits = 0; const original = cells.forEachFormulaAST.bind(cells);
  cells.forEachFormulaAST = fn => original(ast => { visits++; fn(ast); });
  const first = wb.sheetDeps(); assert.deepEqual([...first[1].sheets], [0]); assert.equal(first[1].all, false);
  assert.equal(visits, 1); assert.equal(wb.sheetDeps(), first); assert.equal(visits, 1);
  wb.invalidateStructure(); assert.deepEqual([...wb.sheetDeps()[1].sheets], [0]); assert.equal(visits, 2);
  cells.clear();
});

test('formula edit, removal and Undo/Redo maintain cold dependencies', () => {
  const wb = fixture(); matchesDependencies(wb);
  wb.transact(() => wb.setInput(1, 0, 0, '=Other!A1')); wb.invalidateStructure(); matchesDependencies(wb);
  assert.deepEqual([...wb.sheetDeps()[1].sheets].sort(), [0, 2]);
  wb.undo(); wb.invalidateStructure(); matchesDependencies(wb); assert.deepEqual([...wb.sheetDeps()[1].sheets], [0]);
  wb.redo(); wb.invalidateStructure(); matchesDependencies(wb); assert.deepEqual([...wb.sheetDeps()[1].sheets].sort(), [0, 2]);
  wb.transact(() => wb.setInput(1, 1, 0, 'plain')); wb.invalidateStructure(); matchesDependencies(wb);
  assert.deepEqual([...wb.sheetDeps()[1].sheets], [2]);
});

test('row insertion, column deletion and Undo/Redo preserve formula coordinates and conservative dependencies', () => {
  const wb = fixture(); wb.transact(() => wb.insertRows(0, 3, 2)); matchesDependencies(wb);
  assert.equal(wb.getCell(1, 0, 0).raw, '=Source!A7'); assert.equal(wb.getCell(1, 2, 1).link, '#Source!A7');
  wb.undo(); matchesDependencies(wb); assert.equal(wb.getCell(1, 0, 0).raw, '=Source!A5');
  wb.redo(); matchesDependencies(wb); wb.transact(() => wb.deleteCols(0, 0, 1)); matchesDependencies(wb);
  assert.match(wb.getCell(1, 0, 0).raw, /#REF!/); wb.undo(); matchesDependencies(wb);
});

test('JSON, raw CellMap, prepared and native XLSX restore reconstruct the AST summary', () => {
  const original = fixture();
  for (const restored of [new Workbook(JSON.parse(JSON.stringify(original.serialize()))), new Workbook(readXlsx(writeXlsx(original)).data)]) matchesDependencies(restored);
  const cells = new CellMap({ '0,0': { raw: '=Source!A1' }, '1,0': { raw: '=SUM(' } });
  const restored = new Workbook({ sheets: [{ name: 'Source', cells: {} }, { name: 'Report', cells }] });
  matchesDependencies(restored); assert.equal(restored.sheetDeps()[1].all, true);
  restored.transact(() => restored.replaceSheet(1, { name: 'Report', cells: { '0,0': { raw: '=Source!A1' } } }));
  matchesDependencies(restored); assert.equal(restored.sheetDeps()[1].all, false);
  restored.undo(); matchesDependencies(restored); assert.equal(restored.sheetDeps()[1].all, true);
});

test('table ownership, defined names, dynamic refs and sheet rename rebuild the original dependency rules', () => {
  const table = { name: 'ValuesTable', r1: 0, c1: 0, r2: 1, c2: 0, columns: ['Amount'], header: true };
  const wb = new Workbook({ names: [{ name: 'Chosen', ref: '=Source!A1' }], sheets: [
    { name: 'Source', cells: { '0,0': { raw: 'Amount' }, '1,0': { raw: '5' } }, tables: [table] },
    { name: 'Report', cells: { '0,0': { raw: '=Chosen' }, '1,0': { raw: '=SUM(ValuesTable[Amount])' }, '2,0': { raw: '=INDIRECT("Source!A2")' }, '3,0': { raw: '=MissingSheet!A1' }, '4,0': { raw: '=Source!A2' } } },
    { name: 'Other', cells: { '0,0': { raw: 'Amount' }, '1,0': { raw: '6' } } },
  ] });
  matchesDependencies(wb); assert.equal(wb.sheetDeps()[1].all, true);
  wb.transact(() => { wb.setSheetProp(0, 'tables', []); wb.setSheetProp(2, 'tables', [table]); });
  matchesDependencies(wb); assert.deepEqual([...wb.sheetDeps()[1].sheets].sort(), [0, 2]);
  wb.undo(); matchesDependencies(wb); assert.deepEqual([...wb.sheetDeps()[1].sheets], [0]);
  wb.transact(() => wb.setNames([{ name: 'Chosen', ref: '=Other!A2' }])); matchesDependencies(wb);
  wb.transact(() => wb.renameSheet(0, 'Renamed')); matchesDependencies(wb); assert.ok(wb.getCell(1, 4, 0).raw.includes('Renamed')); assert.equal(wb.getCell(1, 2, 0).raw, '=INDIRECT("Source!A2")');
  wb.undo(); matchesDependencies(wb); wb.redo(); matchesDependencies(wb);
});

test('pivot filter changes recalculate GETPIVOTDATA and ordinary pivot-cell refs with Undo/Redo', () => {
  const pivot = { name: 'Summary', source: 'Source', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['Kind'], values: [{ field: 'Amount', name: 'Total', agg: 'sum' }], top: 0, left: 0, area: { r1: 0, c1: 0, r2: 3, c2: 1 } };
  const wb = new Workbook({ sheets: [
    { name: 'Source', cells: { '0,0': { raw: 'Kind' }, '0,1': { raw: 'Amount' }, '1,0': { raw: 'A' }, '1,1': { raw: '10' }, '2,0': { raw: 'B' }, '2,1': { raw: '20' } } },
    { name: 'Pivot', pivot, cells: { '3,1': { raw: '30' } } },
    { name: 'Report', fileValues: true, cells: { '0,0': { raw: '=GETPIVOTDATA("Total",Pivot!$A$1)', cached: 30 }, '1,0': { raw: '=Pivot!B4', cached: 30 } } },
  ] });
  assert.equal(wb.getValue(2, 0, 0), 30); assert.equal(wb.getValue(2, 1, 0), 30);
  wb.transact(() => { wb.setSheetProp(1, 'pivot', { ...wb.sheets[1].pivot, filters: { Kind: ['A'] } }); wb.setInput(1, 3, 1, '10'); });
  assert.equal(wb.getValue(2, 0, 0), 10); assert.equal(wb.getValue(2, 1, 0), 10); matchesDependencies(wb);
  wb.undo(); assert.equal(wb.getValue(2, 0, 0), 30); assert.equal(wb.getValue(2, 1, 0), 30); matchesDependencies(wb);
  wb.redo(); assert.equal(wb.getValue(2, 0, 0), 10); assert.equal(wb.getValue(2, 1, 0), 10); matchesDependencies(wb);
});
