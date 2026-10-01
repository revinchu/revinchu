import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { cellStylePatch, cellStyleUpdatePatch, importCellStyleList, validCellStyleName } from '../src/cell-style.js';

test('이름 스타일: 포함한 요소만 초기화하고 행 서식의 굵기·채우기도 제거', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setLineStyle(0, 'row', 0, { bold: true, fill: '#ff0000', gradient: { stops: [[0, '#ff0000'], [1, '#0000ff']] }, align: 'right', numFmt: 'currency' });
    wb.setStyle(0, 0, 0, cellStylePatch({ name: '글꼴만', style: {}, include: { number: false, alignment: false, border: false, fill: false, protection: false } }));
  });
  let st = wb.styleAt(0, 0, 0);
  assert.equal(st.bold, false); assert.equal(st.fill, '#ff0000'); assert.equal(st.align, 'right'); assert.equal(st.numFmt, 'currency');
  wb.transact(() => wb.setStyle(0, 0, 0, cellStylePatch({ name: '표준', style: {} })));
  st = wb.styleAt(0, 0, 0);
  assert.equal(st.fill, '#ffffff'); assert.equal(st.gradient, false); assert.equal(st.align, 'general'); assert.equal(st.numFmt, 'general'); assert.equal(st.cellStyleName, undefined);
});

test('스타일 수정·삭제는 직접 추가한 서식과 포함 제외한 요소를 유지', () => {
  const before = { name: '매출', style: { bold: true, color: '#ff0000' }, include: { number: false, alignment: false, border: false, fill: false, protection: false } };
  const current = { ...cellStylePatch(before), size: 20, numFmt: 'currency', fill: '#eeeeee' };
  const next = { ...before, name: '수정', style: { bold: false, color: '#008800', size: 12 } };
  const patch = cellStyleUpdatePatch(current, before, next);
  assert.equal(patch.bold, false); assert.equal(patch.color, '#008800'); assert.equal(patch.cellStyleName, '수정');
  assert.equal('size' in patch, false); assert.equal('numFmt' in patch, false); assert.equal('fill' in patch, false);
  const deleted = cellStyleUpdatePatch(current, before, null);
  assert.equal(deleted.bold, false); assert.equal(deleted.color, '#000000'); assert.equal(deleted.cellStyleName, undefined); assert.equal('size' in deleted, false);
});

test('스타일 병합 입력은 목록만 분리하고 대소문자 중복·예약명·셀 메타데이터 제외', () => {
  assert.equal(validCellStyleName('Revenue 2026'), true); assert.equal(validCellStyleName('bad\u0001name'), false);
  const input = [{ name: '  매출  ', style: { bold: true, cellStyleName: '기존', link: 'https://example.com' }, include: { fill: false } },
    { name: '매출', style: { bold: false } }, { name: 'Normal', style: {} }, { name: '표준', style: {} }, { name: 'STYLE', style: {} }, { name: 'style', style: {} }];
  const result = importCellStyleList(input);
  assert.deepEqual(result.map((s) => s.name), ['매출', 'STYLE']); assert.deepEqual(result[0].style, { bold: true });
  assert.equal(result[0].include.fill, false); assert.equal(result[0].include.font, true);
  result[0].style.bold = false; assert.equal(input[0].style.bold, true);
});

test('대각선의 방향별 색·굵기를 적용/삭제하며 표시형식 적용은 오래된 QUERY 표시를 제거', () => {
  const def = { name: '대각선', style: { du: true, duc: '#112233', dus: 'thick', dd: true, ddc: '#334455', dds: 'double' } };
  const patch = cellStylePatch(def);
  assert.equal(patch.duc, '#112233'); assert.equal(patch.dus, 'thick'); assert.equal(patch.ddc, '#334455'); assert.equal(patch.dds, 'double');
  const reset = cellStyleUpdatePatch(patch, def, null);
  assert.equal(reset.du, false); assert.equal(reset.dd, false); assert.equal(reset.duc, '#000000'); assert.equal(reset.dds, 'thin');
  const wb = new Workbook(); wb.transact(() => { wb.setStyle(0, 0, 0, { queryFormat: '0.00' }); wb.setStyle(0, 0, 0, cellStylePatch({ name: '정수', style: { numFmt: 'number', decimals: 0 } })); });
  assert.equal(wb.styleAt(0, 0, 0).queryFormat, undefined);
  const imported = importCellStyleList([{ name: '원본', style: { queryFormat: '0.00', numFmt: 'custom', code: '0.00' } }]);
  assert.equal(imported[0].style.queryFormat, undefined); assert.equal(imported[0].style.code, '0.00');
});
