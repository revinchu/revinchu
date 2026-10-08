import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { child, kids, parseXml, esc } from '../src/xml.js';
import { functionNameRepairPlan, applyFunctionNameRepairs } from '../src/formula-name-repair.js';

function fixture() {
  const wb = new Workbook();
  wb.transact(() => {
    for (const [c, value] of [[2, 'Amount'], [3, 'Double'], [4, 'Other']]) wb.setInput(0, 2, c, value);
    for (const [r, value] of [[3, '2'], [4, '3'], [5, '4']]) wb.setInput(0, r, 2, value);
    wb.setInput(0, 3, 3, '=C4*2');
    wb.setInput(0, 4, 3, '=C5*3');
    wb.setInput(0, 5, 3, '9');
    wb.setSheetProp(0, 'tables', [{ id: 'demo', name: 'Demo', r1: 2, c1: 2, r2: 5, c2: 4, header: true, totals: false, totalsFns: {}, style: 'None' }]);
  });
  return wb;
}
function withImportedTemplate(wb, column, raw) {
  const files = unzip(writeXlsx(wb));
  const path = 'xl/tables/table1.xml';
  const xml = textOf(files[path]);
  const pattern = new RegExp(`(<tableColumn\\b[^>]*\\bid="${column + 1}"[^>]*)(/>)`);
  assert.match(xml, pattern);
  files[path] = xml.replace(pattern, (_match, open) => `${open}><calculatedColumnFormula>${esc(raw)}</calculatedColumnFormula></tableColumn>`);
  return new Workbook(readXlsx(zip(files)).data);
}
function templatesInXml(wb) {
  const files = unzip(writeXlsx(wb));
  const root = parseXml(textOf(files['xl/tables/table1.xml']));
  return kids(child(root, 'tableColumns'), 'tableColumn').map(column => child(column, 'calculatedColumnFormula')?.text ?? null);
}

test('independent table XML reads calculated templates with absolute column keys and a single leading equals', () => {
  const wb = withImportedTemplate(fixture(), 1, '=_xlfn.IFS(TRUE,C4*2)');
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=IFS(TRUE,C4*2)' });
  assert.equal(wb.getCell(0, 4, 3).raw, '=C5*3');
  assert.equal(wb.getValue(0, 5, 3), 9);
  const restored = new Workbook(wb.serialize());
  assert.deepEqual(restored.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=IFS(TRUE,C4*2)' });
  const again = new Workbook(readXlsx(writeXlsx(restored)).data);
  assert.deepEqual(again.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=IFS(TRUE,C4*2)' });
});

test('a saved template survives constant, blank and per-row formula exceptions without filling data cells', () => {
  const wb = withImportedTemplate(fixture(), 1, 'C4*2');
  wb.transact(() => wb.setInput(0, 3, 3, ''));
  const before = wb.serialize();
  assert.deepEqual(templatesInXml(wb), [null, 'C4*2', null]);
  assert.deepEqual(wb.serialize(), before);
  const again = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(again.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C4*2' });
  assert.equal(again.getValue(0, 3, 3), null);
  assert.equal(again.getCell(0, 4, 3).raw, '=C5*3');
  assert.equal(again.getValue(0, 5, 3), 9);
});

test('current uniform structured formulas supersede an imported template while export keeps the model intact', () => {
  const wb = withImportedTemplate(fixture(), 1, 'C4*2');
  wb.transact(() => { for (let r = 3; r <= 5; r++) wb.setInput(0, r, 3, '=Demo[@Amount]*3'); });
  const before = wb.serialize();
  assert.deepEqual(templatesInXml(wb), [null, 'Demo[[#This Row],[Amount]]*3', null]);
  assert.deepEqual(wb.serialize(), before);
  const again = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(again.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=Demo[[#This Row],[Amount]]*3' });
  assert.equal(again.getValue(0, 3, 3), 6);
});

test('existing inference still writes a calculated column for a new uniform structured formula', () => {
  const wb = fixture();
  wb.transact(() => { for (let r = 3; r <= 5; r++) wb.setInput(0, r, 3, '=Demo[@Amount]*2'); });
  assert.equal(wb.sheets[0].tables[0].calculatedColumnFormulas, undefined);
  assert.deepEqual(templatesInXml(wb), [null, 'Demo[[#This Row],[Amount]]*2', null]);
});

test('row insertion and deletion update template references with undo and redo', () => {
  const wb = withImportedTemplate(fixture(), 1, 'C4*2+$A$1');
  wb.transact(() => wb.insertRows(0, 1, 2));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C6*2+$A$1' });
  wb.undo();
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C4*2+$A$1' });
  wb.redo();
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C6*2+$A$1' });
  wb.transact(() => wb.deleteRows(0, 5, 1));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=#REF!*2+$A$1' });
  assert.deepEqual(readXlsx(writeXlsx(wb)).data.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=#REF!*2+$A$1' });
});

test('column insertion shifts template keys and references, and deletion drops only removed columns', () => {
  const wb = withImportedTemplate(fixture(), 1, 'C4*2');
  wb.sheets[0].tables[0].calculatedColumnFormulas[4] = '=D4+1';
  wb.transact(() => wb.insertCols(0, 3, 1));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 4: '=C4*2', 5: '=E4+1' });
  wb.undo();
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C4*2', 4: '=D4+1' });
  wb.redo();
  wb.transact(() => wb.deleteCols(0, 4, 1));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 4: '=#REF!+1' });
  assert.deepEqual(readXlsx(writeXlsx(wb)).data.sheets[0].tables[0].calculatedColumnFormulas, { 4: '=#REF!+1' });
});

test('moving a whole table through a cell insertion shifts absolute template keys and banded references', () => {
  const wb = withImportedTemplate(fixture(), 1, 'C4*2+$A$1');
  wb.transact(() => assert.equal(wb.shiftCells(0, { r1: 2, r2: 5, c1: 1, c2: 1 }, 'right'), null));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 4: '=D4*2+$A$1' });
  assert.equal(wb.sheets[0].tables[0].c1, 3);
  wb.undo();
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: '=C4*2+$A$1' });
});

test('templates on another sheet follow referenced sheet structure and renaming', () => {
  const wb = withImportedTemplate(fixture(), 1, "'Inputs'!$B$2+C4");
  wb.transact(() => wb.addSheet('Inputs'));
  wb.transact(() => wb.insertRows(1, 0, 2));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: "='Inputs'!$B$4+C4" });
  wb.transact(() => wb.insertCols(1, 0, 1));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: "='Inputs'!$C$4+C4" });
  wb.transact(() => wb.renameSheet(1, 'New Inputs'));
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: "='New Inputs'!$C$4+C4" });
  wb.undo();
  assert.deepEqual(wb.sheets[0].tables[0].calculatedColumnFormulas, { 3: "='Inputs'!$C$4+C4" });
});

test('explicit function repair updates imported cells and their retained template before XLSX roundtrip', () => {
  const initial = fixture();
  const raw = '=IFERROR(SUMIFSㅋ(Demo[Amount],Demo[Double],"red"),0)';
  initial.transact(() => {
    for (const [r, text] of [[3, 'red'], [4, 'blue'], [5, 'red']]) initial.setInput(0, r, 3, text);
    initial.setInput(0, 3, 4, raw);
    initial.setInput(0, 4, 4, raw);
    initial.setInput(0, 5, 4, '11');
  });
  const wb = withImportedTemplate(initial, 2, raw.slice(1));
  const plan = functionNameRepairPlan(wb);
  assert.equal(plan.cellCount, 2);
  assert.equal(plan.templateCount, 1);
  assert.deepEqual(applyFunctionNameRepairs(wb, plan), { cellCount: 2, templateCount: 1, total: 3 });
  assert.equal(wb.getValue(0, 3, 4), 6);
  assert.equal(wb.getValue(0, 5, 4), 11);
  assert.doesNotMatch(templatesInXml(wb)[2], /SUMIFSㅋ/);
  const again = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(again.getValue(0, 3, 4), 6);
  assert.equal(again.getValue(0, 5, 4), 11);
  assert.equal(again.sheets[0].tables[0].calculatedColumnFormulas[4], '=IFERROR(SUMIFS(Demo[Amount],Demo[Double],"red"),0)');
  wb.undo();
  assert.match(wb.sheets[0].tables[0].calculatedColumnFormulas[4], /SUMIFSㅋ/);
  assert.equal(wb.getValue(0, 3, 4), 0);
});
