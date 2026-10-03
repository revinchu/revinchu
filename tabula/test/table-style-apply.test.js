import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { blockValue } from '../src/block.js';
import { tableCellDisplayStyle } from '../src/table-format.js';

const table = (patch = {}) => ({ id: 't', name: '표1', r1: 1, c1: 1, r2: 3, c2: 2, header: true, banded: true, style: 'TableStyleMedium2', ...patch });
const apply = wb => wb.transact(() => { wb.clearTableVisualFormatting(0, 't'); wb.setTableStyle(0, 't', { style: 'TableStyleMedium4' }); });
const visual = { fill: '#ff0000', gradient: { stops: [[0, '#ff0000'], [1, '#000000']] }, pattern: 'darkGrid', patternColor: '#ffff00', color: '#00ff00', bold: true, italic: true, underline: true, strike: true, bt: true, bts: 'double', btc: '#112233', bb: true, bbs: 'thin', bbc: '#445566', bl: true, br: true, dd: true, dds: 'dashed', ddc: '#778899', du: true };

function fixture() {
  return new Workbook({ sheets: [{ name: '자료', fileValues: true, tables: [table()], cells: {
    '0,0': { raw: '=NOSUCHFUNCTION(1)', cached: 99, style: { fill: '#112233' } },
    '1,1': { raw: '분류', style: { ...visual, font: 'Arial', size: 13 } },
    '1,2': { raw: '금액' },
    '2,1': { raw: '001', style: { ...visual, numFmt: 'text' }, comment: '메모', link: 'https://example.com' },
    '2,2': { raw: '=NOSUCHFUNCTION(2)', cached: 42, style: { ...visual, numFmt: 'custom', code: '#,##0.00', align: 'right', valign: 'bottom', wrap: true, locked: false, hideFormula: true } },
    '3,1': { raw: '2026-10-03', style: { ...visual, numFmt: 'custom', code: 'yyyy-mm-dd' } },
    '3,2': { raw: '2', style: { fill: '#ff0000', numFmt: 'percent', decimals: 2 } },
    '4,1': { raw: '바깥', style: { ...visual } },
  } }] });
}

test('빠른 표 스타일은 직접 색·선·강조만 지우고 값·수식·숫자 형식·보호를 보존한다', () => {
  const wb = fixture(), outside = cellData(wb.getCell(0, 4, 1));
  const originalDate = wb.getValue(0, 3, 1);
  apply(wb);
  const st = wb.getCell(0, 2, 2).style;
  for (const key of Object.keys(visual)) assert.equal(st[key], undefined, key);
  assert.equal(st.numFmt, 'custom'); assert.equal(st.code, '#,##0.00'); assert.equal(st.align, 'right');
  assert.equal(st.valign, 'bottom'); assert.equal(st.wrap, true); assert.equal(st.locked, false); assert.equal(st.hideFormula, true);
  assert.equal(wb.getCell(0, 1, 1).style.font, 'Arial'); assert.equal(wb.getCell(0, 1, 1).style.size, 13);
  assert.equal(wb.getValue(0, 2, 1), '001'); assert.equal(wb.getCell(0, 2, 1).comment, '메모'); assert.equal(wb.getCell(0, 2, 1).link, 'https://example.com');
  assert.equal(wb.getValue(0, 2, 2), 42); assert.equal(wb.getCell(0, 2, 2).raw, '=NOSUCHFUNCTION(2)');
  assert.equal(wb.getValue(0, 0, 0), 99); assert.equal(wb.getValue(0, 3, 1), originalDate);
  assert.deepEqual(cellData(wb.getCell(0, 4, 1)), outside);
});

test('빠른 스타일과 직접 서식 초기화는 한 번에 실행 취소·다시 실행된다', () => {
  const wb = fixture(), before = wb.serialize();
  apply(wb); const after = wb.serialize();
  assert.equal(wb.undoStack.length, 1); assert.equal(wb.sheets[0].fileValues, true);
  wb.undo(); assert.deepEqual(wb.serialize(), before); assert.equal(wb.getValue(0, 2, 2), 42);
  wb.redo(); assert.deepEqual(wb.serialize(), after); assert.equal(wb.getValue(0, 0, 0), 99);
  const restored = new Workbook(wb.serialize());
  assert.equal(restored.getCell(0, 2, 2).style.code, '#,##0.00'); assert.equal(restored.getValue(0, 2, 2), 42);
});

test('행·열·시트에서 상속한 색은 표 내부에서만 차단하고 이후 직접 칠하기를 허용한다', () => {
  const wb = fixture(), sh = wb.sheets[0];
  sh.allStyle = { fill: '#ff00ff', bold: true, align: 'center' };
  sh.colStyles[1] = { color: '#778899', italic: true, numFmt: 'text' };
  sh.rowStyles[2] = { fill: '#012345', bt: true, btc: '#ff0000', size: 16 };
  const layers = structuredClone([sh.allStyle, sh.colStyles, sh.rowStyles]);
  const outside = structuredClone(wb.styleAt(0, 4, 1));
  apply(wb);
  const st = wb.styleAt(0, 2, 1);
  assert.notEqual(st.fill, '#012345'); assert.notEqual(st.color, '#778899'); assert.equal(st.bold, false); assert.equal(st.italic, false);
  assert.equal(st.numFmt, 'text'); assert.equal(st.size, 16); assert.equal(st.align, 'center');
  assert.ok(Object.hasOwn(st.tableStyleInherit, 'fill')); assert.ok(Object.hasOwn(st.tableStyleInherit, 'color'));
  assert.deepEqual([sh.allStyle, sh.colStyles, sh.rowStyles], layers); assert.deepEqual(wb.styleAt(0, 4, 1), outside);
  wb.transact(() => wb.setStyle(0, 2, 1, { fill: '#abcdef', color: '#123456', bold: true }));
  const edited = wb.styleAt(0, 2, 1);
  assert.equal(edited.fill, '#abcdef'); assert.equal(edited.color, '#123456'); assert.equal(edited.bold, true);
  for (const key of ['fill', 'color', 'bold']) assert.equal(edited.tableStyleInherit?.[key], undefined);
});

test('같은 빠른 스타일을 다시 적용해도 새로 칠한 색을 초기화한다', () => {
  const wb = fixture(); apply(wb);
  wb.transact(() => wb.setStyle(0, 3, 2, { fill: '#ff3333' }));
  apply(wb);
  assert.equal(wb.getCell(0, 3, 2).style.fill, undefined);
  assert.equal(wb.getCell(0, 3, 2).style.numFmt, 'percent'); assert.equal(wb.getCell(0, 3, 2).style.decimals, 2);
  wb.undo(); assert.equal(wb.getCell(0, 3, 2).style.fill, '#ff3333');
});

test('표 서식 모델만 바꾸는 기존 API는 직접 서식을 지우지 않는다', () => {
  const wb = fixture();
  wb.transact(() => wb.setTableStyle(0, 't', { style: 'None' }));
  assert.equal(wb.getCell(0, 2, 2).style.fill, '#ff0000'); assert.equal(wb.getValue(0, 2, 2), 42);
  assert.equal(wb.clearTableVisualFormatting(0, 'missing'), false);
});

test('표가 10만 행 블록 열 전체를 포함하면 값 배열과 순열을 유지하고 다른 열은 그대로 둔다', () => {
  const wb = new Workbook(), n = 100000, values = new Float64Array(n).fill(12), other = new Float64Array(n).fill(34);
  const perm = Int32Array.from({ length: n }, (_, i) => n - i - 1);
  const block = { r0: 0, c0: 0, n, ver: 0, dver: 0, perm, cols: [
    { num: values, str: null, dict: [], fmt: { ...visual, numFmt: 'custom', code: '#,##0.00' } },
    { num: other, str: null, dict: [], fmt: { fill: '#123456' } },
  ] };
  wb.sheets[0].blocks.push(block); wb.sheets[0].tables.push(table({ r1: 0, c1: 0, r2: n - 1, c2: 0 }));
  const outside = structuredClone(block.cols[1].fmt);
  apply(wb);
  assert.equal(wb.sheets[0].cells.size, 0); assert.equal(block.cols[0].num, values); assert.equal(block.perm, perm); assert.equal(block.dver, 0);
  assert.equal(block.cols[0].fmt.fill, undefined); assert.equal(block.cols[0].fmt.code, '#,##0.00');
  assert.equal(wb.getValue(0, n - 1, 0), 12); assert.deepEqual(block.cols[1].fmt, outside);
  assert.equal(wb.undoStack[0].entries.length, 2);
  wb.undo(); assert.equal(block.cols[0].fmt.fill, '#ff0000'); assert.equal(block.cols[0].num, values);
  wb.redo(); assert.equal(block.cols[0].fmt.fill, undefined); assert.equal(blockValue(block, 99999, 0), 12);
});

test('표가 블록 일부만 포함하면 바깥 행의 직접 서식과 값을 유지한다', () => {
  const wb = new Workbook(), values = Float64Array.from([1, 2, 3, 4, 5, 6, 7, 8]);
  const block = { r0: 0, c0: 0, n: 8, ver: 0, cols: [{ num: values, str: null, dict: [], fmt: { fill: '#ff0000', numFmt: 'custom', code: '0.000' } }] };
  wb.sheets[0].blocks.push(block); wb.sheets[0].tables.push(table({ r1: 2, c1: 0, r2: 4, c2: 0 }));
  apply(wb);
  assert.equal(block.cols[0].fmt.fill, '#ff0000'); assert.equal(wb.styleAt(0, 1, 0).fill, '#ff0000'); assert.equal(wb.styleAt(0, 5, 0).fill, '#ff0000');
  for (let r = 2; r <= 4; r++) { assert.notEqual(tableCellDisplayStyle(wb, 0, r, 0).fill, '#ff0000'); assert.equal(wb.styleAt(0, r, 0).code, '0.000'); assert.equal(wb.getValue(0, r, 0), r + 1); }
  assert.equal(wb.sheets[0].cells.size, 3);
  wb.undo(); assert.equal(wb.sheets[0].cells.size, 0); assert.deepEqual(Array.from(values), [1, 2, 3, 4, 5, 6, 7, 8]);
  wb.redo(); assert.equal(wb.getValue(0, 4, 0), 5);
});

test('희소 표는 100만 빈 행을 순회하거나 실체화하지 않는다', () => {
  const wb = new Workbook({ sheets: [{ name: '큰 표', tables: [table({ r1: 0, c1: 0, r2: 1048575, c2: 0 })], allStyle: { fill: '#fedcba' }, cells: { '0,0': { raw: '제목', style: { fill: '#123456' } }, '1,0': { raw: '값' } } }] });
  let reads = 0; const original = wb.getCell.bind(wb); wb.getCell = (...args) => { reads++; return original(...args); };
  apply(wb);
  assert.ok(reads < 20); assert.equal(wb.sheets[0].cells.size, 2); assert.equal(wb.getValue(0, 1, 0), '값');
});

test('빈 셀의 상속색 차단과 직접 서식 삭제도 가져온 수식 계산값을 무효화하지 않는다', () => {
  const wb = fixture(), sh = wb.sheets[0];
  wb.transact(() => wb.setCellData(0, 3, 1, { raw: '', style: { fill: '#cccccc' } }));
  sh.fileValues = true; wb.getCell(0, 0, 0).dirty = false; wb.getCell(0, 2, 2).dirty = false;
  sh.allStyle = { fill: '#abcdef' }; sh.cells.deleteRC(3, 2);
  apply(wb);
  assert.equal(wb.getValue(0, 0, 0), 99); assert.equal(wb.getValue(0, 2, 2), 42);
  assert.equal(wb.getValue(0, 3, 2), null); assert.ok(wb.getCell(0, 3, 2).style.tableStyleInherit);
  wb.undo(); assert.equal(wb.getValue(0, 2, 2), 42);
  wb.redo(); assert.equal(wb.getValue(0, 2, 2), 42);
});


test('부분 블록 서식에 색만 있을 때도 지운 색이 다시 상속되지 않는다', () => {
  const wb = new Workbook(), num = Float64Array.from([10, 20, 30, 40]);
  const block = { r0: 0, c0: 0, n: 4, cols: [{ num, str: null, dict: [], fmt: { fill: '#ff0000' } }] };
  wb.sheets[0].blocks.push(block); wb.sheets[0].tables.push(table({ r1: 1, c1: 0, r2: 2, c2: 0 }));
  apply(wb);
  assert.notEqual(tableCellDisplayStyle(wb, 0, 1, 0).fill, '#ff0000'); assert.notEqual(tableCellDisplayStyle(wb, 0, 2, 0).fill, '#ff0000');
  assert.equal(wb.styleAt(0, 0, 0).fill, '#ff0000'); assert.equal(wb.styleAt(0, 3, 0).fill, '#ff0000');
  assert.equal(wb.getValue(0, 1, 0), 20); assert.equal(wb.getValue(0, 2, 0), 30);
  wb.undo(); assert.equal(wb.sheets[0].cells.size, 0); assert.deepEqual(Array.from(num), [10, 20, 30, 40]);
});


test('값이 없는 표의 마지막 행·열도 상속색 대신 선택한 표 스타일을 표시한다', () => {
  const wb = new Workbook({ sheets: [{ name: '자료', allStyle: { fill: '#ff0000', color: '#00ff00' }, colStyles: { 2: { fill: '#00ffff' } },
    tables: [table({ r1: 0, c1: 0, r2: 5, c2: 2 })], cells: { '0,0': { raw: '제목' }, '1,0': { raw: '자료' } } }] });
  const outside = structuredClone(wb.styleAt(0, 6, 2));
  apply(wb);
  assert.notEqual(tableCellDisplayStyle(wb, 0, 5, 2).fill, '#00ffff'); assert.notEqual(tableCellDisplayStyle(wb, 0, 5, 1).fill, '#ff0000');
  assert.notEqual(tableCellDisplayStyle(wb, 0, 5, 1).color, '#00ff00'); assert.equal(wb.getValue(0, 5, 2), null);
  assert.deepEqual(wb.styleAt(0, 6, 2), outside);
  wb.undo(); assert.equal(wb.sheets[0].cells.size, 2); assert.equal(wb.styleAt(0, 5, 2).fill, '#00ffff');
});


test('기본 흰색이 있는 표에서 숫자 형식만 수정해도 표 색을 가리지 않는다', () => {
  const wb = new Workbook({ baseStyle: { fill: '#ffffff', color: '#000000', font: 'Arial', size: 12 }, sheets: [{ name: '자료', tables: [table({ r1: 0, c1: 0, r2: 2, c2: 1 })], cells: { '0,0': { raw: '제목' }, '1,0': { raw: '42' } } }] });
  const before = tableCellDisplayStyle(wb, 0, 1, 0);
  wb.transact(() => wb.setStyle(0, 1, 0, { numFmt: 'custom', code: '0.00' }));
  assert.notEqual(wb.getCell(0, 1, 0).style.fill, '#ffffff'); assert.notEqual(wb.getCell(0, 1, 0).style.color, '#000000');
  assert.equal(wb.getCell(0, 1, 0).style.font, 'Arial'); assert.equal(wb.getCell(0, 1, 0).style.size, 12);
  assert.equal(tableCellDisplayStyle(wb, 0, 1, 0).fill, before.fill); assert.equal(tableCellDisplayStyle(wb, 0, 1, 0).color, before.color);
  wb.transact(() => wb.setStyle(0, 2, 0, { fill: '#123456' })); assert.equal(tableCellDisplayStyle(wb, 0, 2, 0).fill, '#123456');
  wb.transact(() => wb.setStyle(0, 3, 0, { numFmt: 'percent' })); assert.equal(wb.styleAt(0, 3, 0).fill, '#ffffff');
  wb.sheets[0].tables[0].style = 'None';
  wb.transact(() => wb.setStyle(0, 2, 1, { numFmt: 'percent' })); assert.equal(wb.styleAt(0, 2, 1).fill, '#ffffff');
});
