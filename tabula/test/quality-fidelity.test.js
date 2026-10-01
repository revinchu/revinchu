import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { Workbook } from '../src/workbook.js';
import { parseInput, formatValue } from '../src/format.js';
import { moveRefsInFormula } from '../src/formula.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';

const roundtrip = (wb) => new Workbook(readXlsx(writeXlsx(wb)).data);

test('분산 탐색이 끝난 뒤 추가한 배열 수식도 앵커 조회 없이 새 분산 영역을 표시', () => {
  const wb = new Workbook();
  wb.transact(() => { wb.addSheet('원본'); wb.addSheet('대상'); });
  assert.deepEqual(wb.spillsOf(0), []);
  wb.transact(() => wb.setInput(2, 4, 7, '=SEQUENCE(10,3)'));
  const shape = () => wb.spillsOf(2).map(({ r, c, h, w }) => ({ r, c, h, w }));
  assert.deepEqual(shape(), [{ r: 4, c: 7, h: 10, w: 3 }]);
  assert.equal(wb.getValue(2, 5, 8), 5);
  wb.transact(() => wb.setInput(2, 4, 7, '=SEQUENCE(2,2)'));
  assert.deepEqual(shape(), [{ r: 4, c: 7, h: 2, w: 2 }]);
  wb.undo();
  assert.deepEqual(shape(), [{ r: 4, c: 7, h: 10, w: 3 }]);
  wb.redo();
  assert.deepEqual(shape(), [{ r: 4, c: 7, h: 2, w: 2 }]);
  wb.transact(() => wb.setInput(2, 4, 7, '일반 값'));
  assert.deepEqual(shape(), []);
});

function tableBook() {
  const wb = new Workbook();
  wb.transact(() => {
    [['서울', '10', '20'], ['부산', '30', '40']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c + 3, v)));
    wb.setSheetProp(0, 'tables', [{ id: 't1', name: '판매', r1: 0, c1: 3, r2: 1, c2: 5, header: false, totals: false, columns: ['지역', '비용', '매출'] }]);
    wb.addSheet('계산');
    wb.setInput(1, 0, 0, '=SUM(판매[매출])');
  });
  return wb;
}

test('표 바깥과 경계의 열 삽입·삭제는 열 이름/구조적 참조를 보존', () => {
  for (const [index, count] of [[0, 1], [3, 1], [6, 1], [0, -1], [6, -1]]) {
    const wb = tableBook();
    wb.transact(() => wb.shiftAxis(0, 'col', index, count));
    assert.deepEqual(wb.sheets[0].tables[0].columns, ['지역', '비용', '매출'], `index=${index} count=${count}`);
    assert.equal(wb.getValue(1, 0, 0), 60);
    const saved = roundtrip(wb);
    assert.deepEqual(saved.sheets[0].tables[0].columns, ['지역', '비용', '매출']);
    saved.invalidate();
    assert.equal(saved.getValue(1, 0, 0), 60);
    wb.undo();
    assert.deepEqual(wb.sheets[0].tables[0].columns, ['지역', '비용', '매출']);
    wb.redo();
    assert.equal(wb.getValue(1, 0, 0), 60);
  }
});

test('표와 일부만 겹치는 삭제는 실제로 지워진 열 이름만 제거', () => {
  const wb = tableBook();
  wb.transact(() => wb.deleteCols(0, 1, 3)); // B:D 삭제: 표 D:F 중 D만 삭제
  assert.deepEqual(wb.sheets[0].tables[0].columns, ['비용', '매출']);
  assert.equal(wb.getValue(1, 0, 0), 60);
  wb.undo();
  assert.deepEqual(wb.sheets[0].tables[0].columns, ['지역', '비용', '매출']);
  wb.transact(() => wb.insertCols(0, 4));
  assert.deepEqual(wb.sheets[0].tables[0].columns, ['지역', '', '비용', '매출']);
  assert.equal(wb.getValue(1, 0, 0), 60);
});

test('xlsx 왕복: 기본 높이 수동 지정과 자동 맞춤 행 높이를 구분해 보존', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setRowHeight(0, 0, 20, true);
    wb.setRowHeight(0, 1, 40, false);
    wb.setRowHeight(0, 2, 40, true);
    wb.setInput(0, 1, 0, '첫 줄\n둘째 줄');
    wb.setStyle(0, 1, 0, { wrap: true });
  });
  const bytes = writeXlsx(wb);
  const xml = textOf(unzip(bytes)['xl/worksheets/sheet1.xml']);
  assert.match(xml, /<row r="1" ht="15" customHeight="1"/);
  assert.doesNotMatch(xml.match(/<row r="2"[^>]*>/)[0], /customHeight="1"/);
  const s = new Workbook(readXlsx(bytes).data).sheets[0];
  assert.deepEqual(s.rowHeights, { 0: 20, 1: 40, 2: 40 });
  assert.deepEqual(s.rowManual, { 0: true, 2: true });
});

test('xlsx 왕복: 이스케이프 모양 텍스트·CR·제어 문자·메모를 손실 없이 보존', () => {
  const wb = new Workbook();
  const values = ['_x0041_', '_x005F_x0041_', '앞\r\n뒤', '앞\r뒤', '앞\u0001뒤', '한글 & < > "\t끝'];
  wb.transact(() => {
    values.forEach((v, r) => wb.setInput(0, r, 0, v));
    wb.setComment(0, 0, 0, '_x0041_\r메모\u0001');
    wb.setInput(0, 6, 0, '="_x0041_"');
  });
  const bytes = writeXlsx(wb);
  const strings = textOf(unzip(bytes)['xl/sharedStrings.xml']);
  assert.match(strings, /_x005F_x0041_/);
  assert.match(strings, /_x000D_/);
  assert.match(strings, /_x0001_/);
  assert.doesNotMatch(strings, /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/);
  const restored = new Workbook(readXlsx(bytes).data);
  values.forEach((v, r) => assert.equal(restored.getValue(0, r, 0), v));
  assert.equal(restored.getCell(0, 0, 0).comment, '_x0041_\r메모\u0001');
  assert.equal(restored.getCell(0, 6, 0).raw, '="_x0041_"');
  assert.equal(restored.getValue(0, 6, 0), '_x0041_');
  restored.invalidate();
  assert.equal(restored.getValue(0, 6, 0), '_x0041_');
});

test('직접 날짜 입력: 없는 날짜를 다음 달 날짜로 바꾸지 않음', () => {
  for (const value of ['2025-02-29', '2026-04-31', '2026년 2월 30일', '2025-02-29 12:34', '2026-04-31 01:02:03']) {
    assert.deepEqual(parseInput(value), { value });
  }
  for (const value of ['2024-02-29', '2000-02-29', '1900-02-29']) {
    const parsed = parseInput(value);
    assert.equal(parsed.numFmt, 'date');
    assert.equal(formatValue(parsed.value, { numFmt: 'custom', code: 'yyyy-mm-dd' }).text, value);
  }
  const timestamp = parseInput('2024-02-29 12:34:56');
  assert.equal(formatValue(timestamp.value, { numFmt: 'custom', code: 'yyyy-mm-dd hh:mm:ss' }).text, '2024-02-29 12:34:56');
});

test('범위 이동: 다른 시트로 옮길 때 외부 참조·GETPIVOTDATA·절대 참조를 보존', () => {
  const options = { targetSheet: '원본', hostSheet: '원본', src: { r1: 0, c1: 0, r2: 2, c2: 2 }, dr: 4, dc: 3 };
  assert.equal(moveRefsInFormula('=A1+$B$2+원본!C3+다른!A1+D4', options), '=D5+$E$6+원본!F7+다른!A1+D4');
  const cross = { ...options, destinationSheet: "새 '보고서" };
  const prefix = "'새 ''보고서'!";
  assert.equal(moveRefsInFormula('=A1+$B$2+원본!C3+다른!A1+D4', cross), `=${prefix}D5+${prefix}$E$6+${prefix}F7+다른!A1+D4`);
  assert.equal(moveRefsInFormula('=GETPIVOTDATA("합계",원본!$A$1)', { ...cross, hostSheet: '요약' }), `=GETPIVOTDATA("합계",${prefix}$D$5)`);
  assert.equal(moveRefsInFormula('=SUM(원본!A1:C3)+SUM(원본!A1:D4)+원본!A:A', cross), `=SUM(${prefix}D5:F7)+SUM(원본!A1:D4)+원본!A:A`);
  assert.equal(moveRefsInFormula('=원본!A1+A1', { ...cross, hostSheet: '요약', dr: 0, dc: 0 }), `=${prefix}A1+A1`);
});

test('xlsx: 1904 날짜 체계 플래그는 불리언 값으로 가져오고 불필요한 미지원 경고를 내지 않는다', () => {
  const bytes = writeXlsx(new Workbook());
  assert.ok(!readXlsx(bytes).warnings.some((s) => s.includes('1904')));
  for (const flag of ['1', 'true', '0', 'false']) {
    const files = unzip(bytes);
    files['xl/workbook.xml'] = textOf(files['xl/workbook.xml']).replace('<bookViews>', `<workbookPr date1904="${flag}"/><bookViews>`);
    const result = readXlsx(zip(files));
    assert.equal(!!result.data.date1904, flag === '1' || flag === 'true');
    assert.equal(result.warnings.some((s) => s.includes('1904')), false);
  }
});
