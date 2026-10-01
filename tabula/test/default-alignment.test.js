import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

test('새 문서·새 시트는 가운데 맞춤을 기본으로 하고 개별 맞춤이 우선한다', () => {
  const wb = new Workbook();
  assert.equal(wb.styleAt(0, 0, 0).align, 'center');
  wb.transact(() => {
    wb.setInput(0, 0, 0, '새 텍스트');
    wb.setStyle(0, 0, 0, { bold: true, fill: '#ffff00' });
    wb.setInput(0, 1, 0, '123');
    wb.setStyle(0, 1, 0, { align: 'right' });
    wb.addSheet('추가');
  });
  assert.equal(wb.styleAt(0, 0, 0).align, 'center');
  assert.equal(wb.styleAt(0, 1, 0).align, 'right');
  assert.equal(wb.styleAt(1, 50, 20).align, 'center');
  const saved = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(saved.styleAt(0, 0, 0).align, 'center');
  assert.equal(saved.styleAt(0, 1, 0).align, 'right');
  assert.equal(new Workbook(wb.serialize()).styleAt(1, 50, 20).align, 'center');
});

test('가져온 문서의 일반·명시 맞춤은 바꾸지 않고 새 시트만 가운데 맞춤', () => {
  const wb = new Workbook({ sheets: [{ name: '기존', cells: { '0,0': { raw: '일반' }, '0,1': { raw: '왼쪽', style: { align: 'left' } }, '0,2': { raw: '오른쪽', style: { align: 'right' } } } }] });
  assert.equal(wb.styleAt(0, 0, 0).align, undefined);
  assert.equal(wb.styleAt(0, 0, 1).align, 'left');
  assert.equal(wb.styleAt(0, 0, 2).align, 'right');
  wb.transact(() => wb.addSheet('새 시트'));
  assert.equal(wb.styleAt(1, 0, 0).align, 'center');
});

test('명시한 일반 맞춤은 시트 가운데 맞춤을 덮어쓰고 XLSX 왕복에서도 보존된다', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, '123');
    wb.setStyle(0, 0, 0, { align: 'general' });
  });
  assert.equal(wb.styleAt(0, 0, 0).align, 'general');
  const saved = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(saved.styleAt(0, 0, 0).align, 'general');
});
