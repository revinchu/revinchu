import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { excelHash, protectXml, protectFromAttrs } from '../src/protect.js';
import { sparkValues, sparkSvg, sparkItems, sparkDefaults } from '../src/sparkline.js';
import { unzip, textOf } from '../src/zip.js';

test('시트 보호: 엑셀 암호 해시 · 허용 항목 · 셀 잠금 xlsx 왕복', () => {
  assert.equal(excelHash('test'), 'CBEB');
  assert.equal(excelHash('password'), '83AF');
  const xml = protectXml({ on: true, hash: 'CBEB', allow: { selectLocked: true, selectUnlocked: true, sort: true, formatCells: true } });
  assert.match(xml, /password="CBEB" sheet="1"/);
  assert.match(xml, /formatCells="0"/);
  assert.match(xml, /sort="0"/);
  const back = protectFromAttrs({ sheet: '1', password: 'CBEB', formatCells: '0', sort: '0' });
  assert.equal(back.allow.sort, true);
  assert.equal(back.allow.insertRows, false);
  assert.equal(back.allow.selectLocked, true);
  const wb = new Workbook();
  wb.transact(() => {
    wb.setCellData(0, 0, 0, { raw: '1', style: { locked: false } });
    wb.setCellData(0, 1, 0, { raw: '=A1*2', style: { hideFormula: true } });
    wb.setSheetProp(0, 'protect', { on: true, hash: excelHash('abc'), allow: { selectLocked: true, selectUnlocked: true } });
  });
  const files = unzip(writeXlsx(wb));
  assert.match(textOf(files['xl/styles.xml']), /<protection locked="0"\/>/);
  const s = readXlsx(writeXlsx(wb)).data.sheets[0];
  assert.equal(s.protect.hash, excelHash('abc'));
  assert.equal(s.cells.get('0,0').style.locked, false);
  assert.equal(s.cells.get('1,0').style.hideFormula, true);
});

test('스파크라인: 값 · SVG · 항목 나누기 · 행 삽입 · xlsx 왕복', () => {
  const wb = new Workbook();
  wb.transact(() => { [[1, 3, -2, 5], [4, 4, 4, 4]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))); });
  assert.deepEqual(sparkValues(wb, 0, 'A1:D1'), [1, 3, -2, 5]);
  const items = sparkItems({ r1: 0, c1: 0, r2: 1, c2: 3 }, { r1: 0, c1: 4, r2: 1, c2: 4 });
  assert.deepEqual(items, [{ r: 0, c: 4, ref: 'A1:D1' }, { r: 1, c: 4, ref: 'A2:D2' }]);
  const g = { id: 'sp1', ...sparkDefaults('column'), high: true, items };
  assert.match(sparkSvg([1, 3, -2, 5], g, 60, 20), /<rect/);
  assert.match(sparkSvg([1, 3, -2, 5], { ...g, type: 'line' }, 60, 20), /<polyline/);
  wb.transact(() => wb.setSheetProp(0, 'sparklines', [g]));
  wb.transact(() => wb.insertRows(0, 1, 1));
  assert.deepEqual(wb.sheets[0].sparklines[0].items.map((it) => [it.r, it.c, it.ref]), [[0, 4, 'A1:D1'], [2, 4, 'A3:D3']]);
  wb.undo();
  assert.equal(wb.sheets[0].sparklines[0].items[1].ref, 'A2:D2');
  const s = readXlsx(writeXlsx(wb)).data.sheets[0];
  assert.equal(s.sparklines.length, 1);
  assert.equal(s.sparklines[0].type, 'column');
  assert.equal(s.sparklines[0].high, true);
  assert.deepEqual(s.sparklines[0].items.map((it) => [it.r, it.c, it.ref.replace(/^.*!/, '')]), [[0, 4, 'A1:D1'], [1, 4, 'A2:D2']]);
});

test('페이지 설정: 방향 · 용지 · 여백 · 맞춤 · 머리글 · 인쇄 영역 · 인쇄 제목 xlsx 왕복', async () => {
  const { headerParts } = await import('../src/page.js');
  assert.deepEqual(headerParts('&L&A&C&P / &N쪽&R&F', { page: 2, pages: 5, sheet: '매출', file: '보고서' }), { left: '매출', center: '2 / 5쪽', right: '보고서' });
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, 'x');
    wb.setSheetProp(0, 'page', { orientation: 'landscape', paper: 8, margins: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      scale: 100, fitW: 1, fitH: 0, gridlines: true, header: '&C&A', footer: '&R&P', area: { r1: 0, c1: 0, r2: 20, c2: 5 }, titleRows: [0, 1] });
  });
  wb.transact(() => wb.insertRows(0, 0, 1));
  assert.deepEqual(wb.sheets[0].page.area, { r1: 1, c1: 0, r2: 21, c2: 5 });
  assert.deepEqual(wb.sheets[0].page.titleRows, [1, 2]);
  const p = readXlsx(writeXlsx(wb)).data.sheets[0].page;
  assert.equal(p.orientation, 'landscape');
  assert.equal(p.paper, 8);
  assert.equal(p.fitW, 1);
  assert.equal(p.fitH, 0);
  assert.equal(p.gridlines, true);
  assert.equal(p.margins.left, 0.25);
  assert.equal(p.header, '&C&A');
  assert.deepEqual(p.area, { r1: 1, c1: 0, r2: 21, c2: 5 });
  assert.deepEqual(p.titleRows, [1, 2]);
});
