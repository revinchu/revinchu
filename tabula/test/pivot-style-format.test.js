import { test } from 'node:test';
import assert from 'node:assert/strict';
import { capturePivotCellFormat } from '../src/pivot-style-format.js';
import { Workbook } from '../src/workbook.js';
import { formatValue } from '../src/format.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { resolvePivot, computePivot, pivotSourceData } from '../src/pivot.js';

test('가져온 기존 피벗 색과 굵기는 다시 고른 스타일을 가리지 않는다', () => {
  const old = { fill: '#fed7aa', color: '#123456', bold: true };
  const capture = capturePivotCellFormat({ ...old, numFmt: 'general', align: 'center' }, old);
  assert.deepEqual(capture, { numFmt: 'general', align: 'center' });
  assert.equal({ ...old, fill: '#c0ffee', ...capture }.fill, '#c0ffee');
});
test('사용자 색·명시 false·숫자 코드·맞춤과 보호는 초기 캡처에서 보존된다', () => {
  const own = { fill: '#aabbcc', bold: false, color: '#000000', numFmt: 'custom', code: '0.000%', decimals: 3, align: 'right', locked: false, wrap: true };
  assert.deepEqual(capturePivotCellFormat(own, { fill: '#fed7aa', bold: true, color: '#123456', numFmt: 'comma' }), own);
});
test('문서 기본 글꼴만 상속하고 색이나 굵기를 바꾼 테두리는 통째로 보존한다', () => {
  const own = { font: 'Arial', size: 12, bt: true, bts: 'double', btc: '#ff0000', bb: false };
  assert.deepEqual(capturePivotCellFormat(own, { bt: true, bts: 'thin', btc: '#ff0000', bb: true }, { font: 'Arial', size: 12 }), { bt: true, bts: 'double', btc: '#ff0000', bb: false });
  assert.equal(own.font, 'Arial');
  assert.deepEqual(capturePivotCellFormat({ bt: true, bts: 'thin', btc: '#000000' }, { bt: true }), {});
});
test('같은 역할의 줄무늬는 각 셀의 생성 서식과 비교하고 일반 숫자 형식은 보존한다', () => {
  const own = { fill: '#ffffff' };
  assert.deepEqual(capturePivotCellFormat(own, { fill: '#ffffff' }), {});
  assert.deepEqual(capturePivotCellFormat(own, { fill: '#abcdef' }), own);
  assert.deepEqual(capturePivotCellFormat({}, { numFmt: 'comma' }), { numFmt: 'general' });
});
test('합성 XLSX 재열기 뒤 스타일 교체는 생성 색만 바꾸고 직접 표시 형식을 보존한다', () => {
  const elements = [{ type: 'wholeTable', style: { color: '#123456' } }, { type: 'headerRow', style: { fill: '#fed7aa', bold: true } }];
  const wb = new Workbook({ sheets: [{ name: '자료', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '매출' }, '1,0': { raw: '서울' }, '1,1': { raw: '123.45' } } }, { name: '피벗', cells: {}, pivot: { name: '피벗1', source: '자료', range: { r1: 0, c1: 0, r2: 1, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 0, left: 0, style: '사용자', styleElements: elements } }], objectStyles: { tables: [{ name: '사용자', table: false, pivot: true, elements }], slicers: [] } });
  const draw = (w, d) => { const res = resolvePivot(pivotSourceData(w, d), d); return computePivot(res, res.def).grid; };
  const first = draw(wb, wb.sheets[1].pivot);
  first.forEach((row, r) => row.forEach((cd, c) => { if (cd) wb.setCellData(1, r, c, { raw: cd.raw, style: { ...cd.style, ...(r === 1 && c === 1 ? { numFmt: 'custom', code: '0.000', align: 'right' } : {}) } }); }));
  wb.sheets[1].pivot.area = { r1: 0, c1: 0, r2: first.length - 1, c2: 1 };
  const read = new Workbook(readXlsx(writeXlsx(wb)).data), def = read.sheets[1].pivot, base = draw(read, def);
  const head = capturePivotCellFormat(read.getCell(1, 0, 0).style, base[0][0].style);
  const body = capturePivotCellFormat(read.getCell(1, 1, 1).style, base[1][1].style);
  def.styleElements.find(e => e.type === 'headerRow').style.fill = '#c0ffee';
  const changed = draw(read, def);
  assert.equal({ ...changed[0][0].style, ...head }.fill, '#c0ffee');
  assert.equal(formatValue(123.45, { ...changed[1][1].style, ...body }).text, '123.450');
  assert.equal({ ...changed[1][1].style, ...body }.decimals, 3);
  assert.equal({ ...changed[1][1].style, ...body }.align, 'right');
});
