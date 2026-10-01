import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { computePivot, pivotLookup, pivotSourceData, resolvePivot } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

// 기대값은 Excel 16.0.14334의 ErrorString 설정 후 Range.Value2와
// GETPIVOTDATA 수식, '-'와 '' PDF 출력으로 독립 확인했다.
const oracle = [
  ['', null], ['-', 0], [' ', 0], [' - ', 0], ['+', '+'], ['0', 0], ['00123', 123],
  ['1.5', '1.5'], ['12.0', '12.0'], ['10%', '10%'], ['(2)', '(2)'], ['1,234', '1,234'],
  ['1E3', '1E3'], ['-12', -12], ['+12', '+12'], [' 12 ', 12], ['\t12', '\t12'],
  ['32767', 32767], ['32768', '32768'], ['-32765', -32765], ['-32766', '-32766'],
  ['-32767', '-32767'], ['-32768', '-32768'], ['-32769', '-32769'], ['0000000001', 1],
  ['2147483648', '2147483648'], ['오류', '오류'], ['#N/A', '#N/A'], ['=1+1', '=1+1'],
  ['－', 0], ['１２', 12], ['－１２', -12], ['−', '−'],
];
function fixture(caption, enabled = true) {
  const wb = new Workbook();
  wb.transact(() => [['지역', '비용', '클릭'], ['오류행', 10, 0], ['정상행', 30, 3]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  wb.transact(() => wb.setSheetProp(0, 'pivot', {
    name: '오류문구', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 2 },
    rows: ['지역'], cols: [], values: [{ field: 'CPC', name: '클릭단가', agg: 'sum' }],
    calcFields: [{ name: 'CPC', formula: '비용/클릭' }], layout: 'tabular',
    top: 0, left: 5, area: { r1: 0, c1: 5, r2: 3, c2: 6 }, errorShow: enabled, errorCaption: caption,
  }));
  return wb;
}
function check(wb, caption, expected) {
  const def = wb.sheets[0].pivot, source = pivotSourceData(wb, def), res = resolvePivot(source, def);
  assert.equal(def.errorCaption, caption, '저장된 대체문구 자체는 변경하지 않는다.');
  const grid = computePivot(res, res.def).grid;
  const raw = grid.find((r) => r[0]?.raw === '오류행')[1].raw;
  assert.equal(raw, expected === null ? '' : typeof expected === 'number' ? String(expected) : "'" + expected, '화면 셀의 타입과 내용');
  assert.equal(pivotLookup(source, def, '클릭단가', [['지역', '오류행']], res), expected ?? 0, '빈 셀 조회는 0, 문자열 조회는 문자열');
  wb.transact(() => {
    grid.forEach((row, r) => row.forEach((cell, c) => wb.setCellData(0, r, 5 + c, { raw: cell.raw, style: cell.style })));
    wb.setInput(0, 0, 9, '=GETPIVOTDATA("클릭단가",F1,"지역","오류행")');
  });
  assert.equal(wb.getValue(0, 1, 6), expected, '실제 워크북에 기록한 셀 타입');
  assert.equal(wb.getValue(0, 0, 9), expected ?? 0, '실제 GETPIVOTDATA 수식');
}
test('Excel 오류 대체문구 33종: 표시 타입·GETPIVOTDATA·XLSX 왕복', () => {
  for (const [caption, expected] of oracle) {
    const wb = fixture(caption); check(wb, caption, expected);
    check(new Workbook(readXlsx(writeXlsx(wb)).data), caption, expected);
  }
});
test('오류 표시 false는 대체문구가 있어도 셀과 GETPIVOTDATA에 원래 오류 보존', () => {
  for (const caption of ['', '-', '0', '오류']) {
    const wb = fixture(caption, false), def = wb.sheets[0].pivot, source = pivotSourceData(wb, def);
    const res = resolvePivot(source, def), grid = computePivot(res, res.def).grid;
    assert.equal(grid.find((r) => r[0]?.raw === '오류행')[1].raw, '#DIV/0!');
    assert.equal(pivotLookup(source, def, '클릭단가', [['지역', '오류행']], res).code, '#DIV/0!');
    wb.transact(() => wb.setInput(0, 0, 9, '=GETPIVOTDATA("클릭단가",F1,"지역","오류행")'));
    assert.equal(wb.getValue(0, 0, 9).code, '#DIV/0!');
  }
});
