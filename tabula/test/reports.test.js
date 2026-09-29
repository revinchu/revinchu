// 실제 보고서 파일에서 찾은 엑셀 동작: 1900 날짜 체계 · 빈 인수 · 배열 IF · *IFS 색인 · 피벗 빈 글자 항목 · 계산 필드 총합계 · 표시 글자
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { computePivot, resolvePivot, pivotSourceData, pivotLookup } from '../src/pivot.js';
import { formatValue } from '../src/format.js';

const put = (wb, rows) => wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== undefined) wb.setInput(0, r, c, v); })));
const calc = (wb, f, r = 300, c = 7) => { wb.setInput(0, r, c, f); return wb.getValue(0, r, c); };

test('1900 날짜 체계: 일련번호 0 = 1900-01-00, 60 = 1900-02-29', () => {
  const wb = new Workbook();
  assert.equal(calc(wb, '=YEAR(0)'), 1900);
  assert.equal(calc(wb, '=MONTH(0)'), 1);
  assert.equal(calc(wb, '=DAY(0)'), 0);
  assert.equal(calc(wb, '=WEEKNUM(0,12)'), 1);
  assert.equal(calc(wb, '=WEEKDAY(0)'), 7);
  assert.equal(calc(wb, '=WEEKDAY(1)'), 1);
  assert.equal(calc(wb, '=DAY(60)'), 29);
  assert.equal(calc(wb, '=DATE(1900,1,1)'), 1);
  assert.equal(calc(wb, '=DATE(1900,3,1)'), 61);
  assert.equal(calc(wb, '=DATEVALUE("1900-01-05")'), 5);
  assert.equal(calc(wb, '=TEXT(0,"yyyy-mm-dd")'), '1900-01-00');
  assert.equal(calc(wb, '=TEXT(61,"yyyy-mm-dd ddd")'), '1900-03-01 Thu');
  assert.equal(formatValue(45000, { numFmt: 'date' }).text, '2023-03-15');
});

test('빈 인수 · 배열 IF 의 빈 칸 · *IFS 해시 색인', () => {
  const wb = new Workbook();
  const rows = [['k', 'v', 'd']];
  for (let i = 0; i < 200; i++) rows.push([i % 2 ? 'b' : 'a', String(i), i % 5 ? String(45000 + i) : undefined]);
  put(wb, rows);
  // 정렬되지 않은 범위: VLOOKUP(…,) · MATCH(…,) 는 정확히 일치 (빈 인수 = FALSE/0)
  put(wb, [[], [], [], [], [], [], [], [], [], [], [undefined, undefined, undefined, 'z', '9'], [undefined, undefined, undefined, 'a', '1'], [undefined, undefined, undefined, 'y', '5']]);
  assert.equal(calc(wb, '=VLOOKUP("a",D11:E13,2,)'), 1);
  assert.equal(calc(wb, '=MATCH("a",D11:D13,)'), 2);
  // 배열 IF 에서 빈 칸은 0 → MIN 이 0
  assert.equal(calc(wb, '=MIN(IF(A2:A201="a",C2:C201))'), 0);
  // *IFS: 같음 조건은 색인 경로, 결과는 느린 경로와 같아야 함
  assert.equal(calc(wb, '=SUMIFS(B2:B201,A2:A201,"a")'), Array.from({ length: 100 }, (_, i) => i * 2).reduce((s, x) => s + x, 0));
  assert.equal(calc(wb, '=COUNTIFS(A2:A201,"B",B2:B201,"7")'), 1);
  assert.equal(calc(wb, '=MAXIFS(C2:C201,A2:A201,"b")'), 45199);
  assert.equal(calc(wb, '=MINIFS(C2:C201,A2:A201,"a")'), 45002);
  assert.equal(calc(wb, '=AVERAGEIFS(B2:B201,A2:A201,"=b")'), 100);
  assert.equal(calc(wb, '=COUNTIFS(A2:A201,"a",C2:C201,"")'), 20);
  assert.equal(calc(wb, '=MINIFS(C2:C201,A2:A201,"없음")'), 0);
  // 크기 비교 · 와일드카드는 원래 경로
  assert.equal(calc(wb, '=COUNTIFS(A2:A201,"a*",B2:B201,">=190")'), 5);
});

function pivotOf(wb, def) {
  const res = resolvePivot(pivotSourceData(wb, def), def);
  return { res, grid: computePivot(res, res.def).grid.map((r) => r.map((c) => (c ? String(c.raw ?? '').replace(/^'/, '') : null))) };
}

test('피벗: 빈 글자 항목은 빈 이름 · 데이터 없는 총합계의 계산 필드 · 숫자 오류 표시 · 필드 이름 머리글', () => {
  const wb = new Workbook();
  put(wb, [['매체', '비용', '매출'], ["'", '10', '30'], ['네이버', '20', '10'], [undefined, '5', '0']]);
  const base = { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, cols: [], pages: [], layout: 'tabular', top: 0, left: 5 };
  // 빈 글자("")는 '(비어 있음)'(빈 칸)과 다른 항목, 이름은 빈칸
  const { grid } = pivotOf(wb, { ...base, rows: ['매체'], values: [{ field: '비용', agg: 'sum' }], filters: {} });
  assert.deepEqual(grid.slice(1, 4).map((r) => r[0]), ['', '네이버', '(비어 있음)']);
  // 걸러서 데이터가 없는 총합계: 계산 필드(차이)는 0, 나눗셈은 오류, 일반 합계는 빈칸
  const calcFields = [{ name: '차이', formula: '매출-비용' }, { name: 'ROAS', formula: '매출/비용' }];
  const def = { ...base, rows: ['매체'], values: [{ field: '비용', agg: 'sum' }, { field: '차이', agg: 'sum' }, { field: 'ROAS', agg: 'sum' }], filters: { 매체: ['없는 항목'] }, calcFields };
  const empty = pivotOf(wb, def).grid;
  assert.deepEqual(empty[empty.length - 1].slice(1), ['', '0', '#DIV/0!']);
  // 오류 표시 글자가 숫자 모양이면 숫자로 씀 (엑셀은 셀에 숫자 0 저장)
  const zero = pivotOf(wb, { ...def, errorCaption: '0' }).grid;
  assert.equal(zero[zero.length - 1][3], '0');
  // GETPIVOTDATA: 오류 표시 글자가 빈칸이면 0
  const res = resolvePivot(pivotSourceData(wb, { ...def, errorCaption: '' }), { ...def, errorCaption: '' });
  assert.equal(pivotLookup(null, { ...def, errorCaption: '' }, 'ROAS', [], res), 0);
  // 테이블 형식 + 열 필드: 머리글 첫 행에 열 필드 이름 (압축 형식은 '열 레이블')
  const t = pivotOf(wb, { ...base, rows: ['비용'], cols: ['매체'], values: [{ field: '매출', agg: 'sum' }], filters: {} }).grid;
  assert.equal(t[0][1], '매체');
  const c = pivotOf(wb, { ...base, layout: 'compact', rows: ['비용'], cols: ['매체'], values: [{ field: '매출', agg: 'sum' }], filters: {} }).grid;
  assert.equal(c[0][1], '열 레이블');
});

test('피벗 보고서 필터: 원본에 있는 항목만 세어 하나면 그 이름', () => {
  const wb = new Workbook();
  put(wb, [['매체', '비용'], ['네이버', '1'], ['구글', '2']]);
  const base = { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: [], cols: [], values: [{ field: '비용', agg: 'sum' }], pages: ['매체'], layout: 'tabular', top: 0, left: 4 };
  // '카카오' 는 원본에 없는 항목 (엑셀의 m="1") → 보이는 항목은 네이버 하나
  assert.equal(pivotOf(wb, { ...base, filters: { 매체: ['네이버', '카카오'] } }).grid[0][1], '네이버');
  assert.equal(pivotOf(wb, { ...base, filters: { 매체: ['카카오'] } }).grid[0][1], '(다중 항목)');
  assert.equal(pivotOf(wb, { ...base, filters: { 매체: ['네이버', '구글'] } }).grid[0][1], '(다중 항목)');
});
