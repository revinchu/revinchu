import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { digitWidth, width2pxM, pt2px, readXlsx, writeXlsx } from '../src/xlsx.js';
import { fontDigitWidth, rowPointsToPixels, rowPixelsToPoints, columnCharsToPixels, pixelsToColumnChars } from '../src/dimension.js';

test('열 문자 단위는 Excel COM 합성 문서의 Calibri/맑은 고딕 결과와 일치한다', () => {
  const chars = [0, .1, .5, .99, 1, 2, 8.43, 10, 20, 255];
  const oracle = [
    ['Calibri', [0, 1, 6, 12, 12, 19, 64, 75, 145, 1790], [0, .08, .5, 1, 1, 2, 8.43, 10, 20, 255]],
    ['맑은 고딕', [0, 1, 7, 13, 13, 21, 72, 85, 165, 2045], [0, .08, .54, 1, 1, 2, 8.38, 10, 20, 255]],
  ];
  for (const [name, pixels, display] of oracle) {
    const font = { name, size: 11 };
    assert.deepEqual(chars.map((n) => columnCharsToPixels(n, font)), pixels);
    assert.deepEqual(pixels.map((n) => pixelsToColumnChars(n, font)), display);
  }
});

test('기존 XLSX MDW/행 높이 변환과 동일하며 UI 열 문자와 파일 width를 혼동하지 않는다', () => {
  for (const name of ['Calibri', 'Calibri Light', '맑은 고딕', 'Malgun Gothic', 'Arial', '굴림', 'Gulim', '굴림체', '돋움', 'Dotum', '돋움체', '바탕', 'Batang', '나눔고딕', 'NanumGothic', 'Nanum Gothic', 'Times New Roman', 'Cambria', 'Segoe UI', 'Verdana', 'Tahoma', 'Meiryo UI', 'MS Gothic', 'SimSun', 'Unknown']) for (const size of [9, 11, 18]) assert.equal(fontDigitWidth({ name, size }), digitWidth({ name, size }));
  for (const points of [0, .1, 1, 15, 15.2, 20, 409, 409.5]) assert.equal(rowPointsToPixels(points), pt2px(points));
  assert.equal(rowPixelsToPoints(20), 15);
  assert.equal(rowPointsToPixels(409.5), 546);
  assert.equal(width2pxM(10.7109375, 7), columnCharsToPixels(10, 7));
  assert.notEqual(width2pxM(10, 7), columnCharsToPixels(10, 7));
});

test('행 pt/열 문자 입력은 Excel 최대값과 유한 숫자만 허용한다', () => {
  for (const n of [-1, 409.51, NaN, Infinity, '15']) assert.throws(() => rowPointsToPixels(n), (e) => e.code === 'INVALID_DIMENSION');
  for (const n of [-1, 255.01, NaN, Infinity, '10']) assert.throws(() => columnCharsToPixels(n), (e) => e.code === 'INVALID_DIMENSION');
});

test('행·열 UI 크기 변경은 기존 모델의 Undo와 XLSX 저장 왕복에서 픽셀 치수를 보존한다', () => {
  for (const name of ['Calibri', '맑은 고딕']) {
    const wb = new Workbook(), font = { name, size: 11 };
    wb.defaultFont = font;
    const before = { row: wb.rowHeight(0, 0), col: wb.colWidth(0, 0) };
    wb.transact(() => { wb.setRowHeight(0, 0, rowPointsToPixels(30)); wb.setColWidth(0, 0, columnCharsToPixels(10, font)); });
    const round = new Workbook(readXlsx(writeXlsx(wb)).data);
    assert.equal(round.rowHeight(0, 0), 40);
    assert.equal(pixelsToColumnChars(round.colWidth(0, 0), round.defaultFont), 10);
    wb.undo();
    assert.equal(wb.rowHeight(0, 0), before.row);
    assert.equal(wb.colWidth(0, 0), before.col);
  }
});
