import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';

test('셀 스타일 목록 변경은 한 번에 실행 취소·다시 실행하고 원본 값을 보존한다', () => {
  const wb = new Workbook({ sheets: [{ name: '보고서', fileValues: true, cells: { '0,0': { raw: '=NOTAFUNC(1)', cached: 42 } } }] });
  let events = 0; wb.onChange(() => events++);
  const list = [{ name: '강조', style: { fill: '#123456', gradient: { stops: [[0, '#ffffff']] } }, include: { fill: true } }];
  const cells = wb.sheets[0].cells;
  wb.transact(() => { wb.setCellStyles(list); wb.setCellStyles([...list, { name: '제목', style: { bold: true } }]); }, { action: '스타일 병합' });
  list[0].style.gradient.stops[0][1] = '#000000';
  assert.equal(wb.cellStyles[0].style.gradient.stops[0][1], '#ffffff');
  assert.equal(wb.cellStyles.length, 2);
  assert.equal(wb.getValue(0, 0, 0), 42);
  assert.equal(wb.undoStack.length, 1);
  assert.equal(wb.undoStack[0].entries.length, 1, '목록만 기록하며 전체 통합 문서를 복사하지 않음');
  wb.undo(); assert.equal(wb.cellStyles, null); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.redo(); assert.equal(wb.cellStyles.length, 2); assert.equal(wb.getValue(0, 0, 0), 42);
  assert.equal(wb.sheets[0].cells, cells);
  assert.equal(wb.sheets[0].fileValues, true);
  assert.equal(events, 3, '자동 저장 UI가 변경·실행 취소·다시 실행을 감지');
});

test('셀 스타일의 저장·복원 결과를 바꿔도 통합 문서의 중첩 속성은 바뀌지 않는다', () => {
  const data = { sheets: [{ name: 'S', cells: {} }], cellStyles: [{ name: '사용자', include: { fill: true }, style: { gradient: { stops: [[0, '#112233']] } } }] };
  const wb = new Workbook(data);
  data.cellStyles[0].include.fill = false;
  const snap = wb.serialize(); snap.cellStyles[0].style.gradient.stops[0][1] = '#ffffff';
  assert.equal(wb.cellStyles[0].include.fill, true);
  assert.equal(wb.cellStyles[0].style.gradient.stops[0][1], '#112233');
});

test('10만 행 블록의 스타일 수정은 값 배열을 복사하지 않고 실행 취소한다', () => {
  const wb = new Workbook(), num = new Float64Array(100000).fill(7);
  const block = { r0: 0, c0: 0, n: num.length, ver: 0, cols: [{ num, str: null, dict: [], fmt: { fill: '#ffffff', cellStyleName: '사용자' } }] };
  wb.sheets[0].blocks.push(block);
  const originalFormat = structuredClone(block.cols[0].fmt);
  const next = { fill: '#112233', cellStyleName: '사용자' };
  wb.transact(() => { wb.setBlockStyle(0, 0, 0, { fill: '#abcdef' }); wb.setBlockStyle(0, 0, 0, next); });
  next.fill = '#000000';
  assert.equal(wb.styleAt(0, 99999, 0).fill, '#112233');
  assert.equal(wb.getValue(0, 99999, 0), 7);
  assert.equal(wb.undoStack[0].entries.length, 1);
  assert.equal(block.cols[0].num, num);
  wb.undo(); assert.deepEqual(block.cols[0].fmt, originalFormat);
  wb.redo(); assert.equal(block.cols[0].fmt.fill, '#112233');
  assert.equal(block.cols[0].num, num); assert.equal(num[99999], 7);
  assert.equal(wb.setBlockStyle(0, 9, 0, {}), false);
});

test('표준 스타일 수정은 실행 취소되며 정의의 중첩 속성도 독립적이다', () => {
  const wb = new Workbook();
  const original = { font: '맑은 고딕', align: 'center' };
  wb.baseStyle = structuredClone(original);
  const next = { font: 'Arial', fill: '#123456', gradient: { stops: [[0, '#ffffff']] } };
  wb.transact(() => wb.setBaseStyle(next));
  next.gradient.stops[0][1] = '#000000';
  assert.equal(wb.baseStyle.gradient.stops[0][1], '#ffffff');
  wb.undo(); assert.deepEqual(wb.baseStyle, original);
  wb.redo(); assert.equal(wb.baseStyle.font, 'Arial');
});

test('서식·메모 편집과 실행 취소는 파일 계산값을 바꾸지 않으며 텍스트 서식도 수식을 유지한다', () => {
  const wb = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: { '0,0': { raw: '=1+1', cached: 42 }, '0,1': { raw: '=A1*2', cached: 84 } } }] });
  wb.transact(() => wb.setStyle(0, 0, 0, { bold: true, numFmt: 'text' }));
  assert.equal(wb.getCell(0, 0, 0).formula, true);
  assert.equal(wb.getValue(0, 0, 0), 42);
  assert.equal(wb.getValue(0, 0, 1), 84);
  wb.undo(); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.redo(); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.transact(() => wb.setComment(0, 0, 0, '메모'));
  assert.equal(wb.serialize().sheets[0].cells['0,0'].cached, 42);
  wb.undo(); wb.redo(); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.transact(() => wb.setInput(0, 0, 0, '7'));
  assert.equal(wb.getValue(0, 0, 0), '7', '사용자 입력은 텍스트 표시 형식을 따른다');
  assert.equal(wb.getValue(0, 0, 1), 14, '실제 값 편집은 의존 수식을 다시 계산한다');
});
