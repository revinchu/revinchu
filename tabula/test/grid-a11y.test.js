import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { MAX_ROWS, MAX_COLS } from '../src/formula.js';
import { selectionDescription, gridCoordinates, gridMerges, gridCellInfo } from '../src/grid-a11y.js';

const state = (cells = {}, extra = {}) => ({ wb: new Workbook({ sheets: [{ name: '접근성 검사', cells, ...extra }] }), si: 0 });
test('접근성 셀 안내는 표시값·수식·오류·메모와 링크를 구분한다', () => {
  const s = state({ '0,0': { raw: '=1+1', comment: '합성 메모', link: 'https://example.com' }, '0,1': { raw: '=1/0' }, '0,2': { raw: '0.25', style: { numFmt: 'percent', decimals: 0 } } });
  assert.equal(gridCellInfo(s, 0, 0).label, 'A1, 2, 수식 =1+1, 메모 있음, 하이퍼링크 있음');
  assert.match(gridCellInfo(s, 0, 1).label, /^B1, 오류 #DIV\/0!, 수식 =1\/0$/);
  assert.equal(gridCellInfo(s, 0, 2).label, 'C1, 25%');
  assert.equal(gridCellInfo(s, 1, 0).label, 'A2, 빈 셀');
});
test('보호한 숨긴 수식은 발표하지 않으며 잠금 해제 셀·공개 보기의 읽기 전용을 구별한다', () => {
  const s = state({ '0,0': { raw: '=40+2', style: { hideFormula: true } }, '0,1': { raw: '7', style: { locked: false } } }, { protect: { on: true } });
  assert.equal(gridCellInfo(s, 0, 0).label, 'A1, 42, 읽기 전용');
  assert.equal(gridCellInfo(s, 0, 1).readonly, false);
  assert.equal(gridCellInfo({ ...s, readonly: true }, 0, 1).readonly, true);
  s.wb.props.markedFinal = true; assert.equal(gridCellInfo(s, 0, 1).readonly, true);
});
test('병합 셀은 원점·행열 범위로 표현하고 후보 인덱스는 겹치는 네 창을 통합한다', () => {
  const merge = { r1: 1, c1: 1, r2: 3, c2: 4 }, s = state({ '1,1': { raw: '병합 제목' } });
  assert.deepEqual(gridCellInfo(s, 2, 3, merge), { r: 1, c: 1, label: 'B2, 병합 제목, 병합 B2:E4', readonly: false, rowSpan: 3, colSpan: 4 });
  const candidates = [{ r: 0, c: 0 }, { r: 1, c: 1 }, { r: 2, c: 3 }, { r: 4, c: 3 }];
  const found = gridMerges(candidates, [merge]);
  assert.equal(found.size, 2); assert.equal(found.get('2,3'), merge);
});
test('좌표 수와 숨긴 거대 영역 방문 횟수를 제한하고 활성 셀은 항상 포함한다', () => {
  let calls = 0; const visible = { size: () => { calls++; return 20; } };
  const window = { r1: 0, c1: 0, r2: 5000000, c2: 10000 };
  const coords = gridCoordinates({ r: MAX_ROWS - 1, c: MAX_COLS - 1 }, [window, window, window, window], visible, visible);
  assert.equal(coords.length, 400); assert.deepEqual(coords[0], { r: MAX_ROWS - 1, c: MAX_COLS - 1 });
  assert.equal(new Set(coords.map(p => `${p.r},${p.c}`)).size, coords.length); assert.ok(calls < 4000);
  calls = 0; const hidden = { size: () => { calls++; return 0; } };
  assert.deepEqual(gridCoordinates({ r: 99, c: 99 }, [window], hidden, hidden), [{ r: 99, c: 99 }]); assert.ok(calls < 1700);
});
test('범위·행·열·전체 및 비연속 선택을 한국어로 안내한다', () => {
  const s = { sel: { r1: 1, c1: 2, r2: 4, c2: 3 }, selKind: 'cells' };
  assert.equal(selectionDescription(s), 'C2:D5 범위 선택');
  assert.equal(selectionDescription({ ...s, selKind: 'rows' }), '2행부터 5행까지 선택');
  assert.equal(selectionDescription({ ...s, selKind: 'cols' }), 'C열부터 D열까지 선택');
  assert.equal(selectionDescription({ ...s, selKind: 'all' }), '워크시트 전체 선택');
  assert.equal(selectionDescription({ ...s, special: [[1,2], [4,3]] }), '비연속 2개 셀 선택');
});
test('긴 문자열·확인란 및 1904년 날짜 표시를 안전한 일반 텍스트로 전달한다', () => {
  const s = state({ '0,0': { raw: '<img onerror=alert(1)>' + '가'.repeat(500) }, '0,1': { raw: 'TRUE', style: { checkbox: true } }, '0,2': { raw: '1', style: { numFmt: 'custom', code: 'yyyy-mm-dd' } } });
  assert.match(gridCellInfo(s, 0, 0).label, /^A1, <img onerror=alert\(1\)>/); assert.match(gridCellInfo(s, 0, 0).label, /일부 표시/);
  assert.equal(gridCellInfo(s, 0, 1).label, 'B1, 확인란 선택됨');
  s.wb.date1904 = true; assert.equal(gridCellInfo(s, 0, 2).label, 'C1, 1904-01-02');
});
