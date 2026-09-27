import { test } from 'node:test';
import assert from 'node:assert/strict';
import { formatCode, styleForCode, adjustCodeDecimals, formatValue, parseInput } from '../src/format.js';
import { buildCode, describeCode, DATE_TYPES, TIME_TYPES, CUSTOM_LIST, SPECIAL_TYPES, FRACTION_TYPES } from '../src/fmtpresets.js';
import { splitDelimited, splitFixed, suggestBreaks, parseDateOrder, convertPart } from '../src/textsplit.js';
import {
  parseSpec, resolveStructRef, canonicalRef, expansionFor, uniqueNames, validTableName, tableCellStyle, normalizeStyleName, TABLE_STYLES,
} from '../src/tables.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { evaluateFormula, parse } from '../src/formula.js';

const fc = (v, code) => formatCode(v, code).text;

test('사용자 지정 서식: 숫자', () => {
  assert.equal(fc(1234.567, '#,##0.00'), '1,234.57');
  assert.equal(fc(-1234.567, '#,##0.00;[빨강](#,##0.00)'), '(1,234.57)');
  assert.equal(formatCode(-1, '0;[Red]-0').color, '#ff0000');
  assert.equal(fc(0, '#,##0;-#,##0;"-"'), '-');
  assert.equal(fc(0.1234, '0.0%'), '12.3%');
  assert.equal(fc(1234567, '#,##0,"천원"'), '1,235천원');
  assert.equal(fc(1234567, '0.0,,"M"'), '1.2M');
  assert.equal(fc(12345, '0.00E+00'), '1.23E+04');
  assert.equal(fc(0.5, '# ?/?'), ' 1/2');
  assert.equal(fc(1.25, '# ?/4'), '1 1/4');
  assert.equal(fc(1012345678, '000-0000-0000'), '010-1234-5678');
  assert.equal(fc(123, '00000'), '00123');
  assert.equal(fc(3.14159, '0.##'), '3.14');
  assert.equal(fc(0, '#,###'), '');
  assert.equal(fc(-5, '"₩"#,##0'), '-₩5');
  assert.equal(fc(8, '0"개"'), '8개');
  assert.equal(fc(12345.678, 'General'), '12345.678');
});

test('사용자 지정 서식: 조건 · 텍스트 · 날짜', () => {
  assert.deepEqual(formatCode(150, '[>100][빨강]0;[파랑]0'), { text: '150', color: '#ff0000' });
  assert.deepEqual(formatCode(50, '[>100][빨강]0;[파랑]0'), { text: '50', color: '#0000ff' });
  assert.equal(fc('abc', '@"님"'), 'abc님');
  assert.equal(fc('abc', '0;-0;0;"["@"]"'), '[abc]');
  assert.equal(fc('abc', '#,##0'), 'abc');
  assert.equal(fc(45366.5625, 'yyyy-mm-dd hh:mm'), '2024-03-15 13:30');
  assert.equal(fc(45366, 'yyyy"년" m"월" d"일" aaaa'), '2024년 3월 15일 금요일');
  assert.equal(fc(45366.5625, '[$-412]AM/PM h:mm:ss'), '오후 1:30:00');
  assert.equal(fc(0.75, 'h:mm AM/PM'), '6:00 PM');
  assert.equal(fc(1.5625, '[h]:mm:ss'), '37:30:00');
  assert.equal(fc(45366, 'mmm d, yyyy ddd'), 'Mar 15, 2024 Fri');
  assert.equal(fc(45366, 'yy.mm.dd'), '24.03.15');
  assert.equal(fc(0.5, 'hh:mm:ss.00'), '12:00:00.00');
});

test('서식 코드 → 스타일, 자릿수 조정, 셀 표시', () => {
  assert.deepEqual(styleForCode('#,##0'), { numFmt: 'comma', decimals: undefined, code: undefined });
  assert.deepEqual(styleForCode('yyyy-mm-dd'), { numFmt: 'date', decimals: undefined, code: undefined });
  assert.deepEqual(styleForCode('0"개"'), { numFmt: 'custom', code: '0"개"', decimals: undefined });
  assert.equal(styleForCode('G/표준').numFmt, undefined);
  assert.equal(adjustCodeDecimals('#,##0', 1), '#,##0.0');
  assert.equal(adjustCodeDecimals('#,##0.0', -1), '#,##0');
  assert.equal(adjustCodeDecimals('"₩"#,##0;[빨강]-"₩"#,##0', 1), '"₩"#,##0.0;[빨강]-"₩"#,##0.0');
  assert.deepEqual(formatValue(-3, { numFmt: 'custom', code: '0;[빨강]-0' }), { text: '-3', align: 'right', color: '#ff0000' });
  assert.deepEqual(formatValue(5, null), { text: '5', align: 'right' });
  assert.deepEqual(parseInput('2024-03-15 13:30'), { value: 45366.5625, numFmt: 'datetime' });
});

test('셀 서식 대화상자의 범주 코드', () => {
  assert.equal(buildCode('number', { decimals: 1, thousands: true, negative: 'paren' }), '#,##0.0_);(#,##0.0)');
  assert.equal(buildCode('currency', { decimals: 0, symbol: '₩', negative: 'minus' }), '"₩"#,##0');
  assert.deepEqual(describeCode('#,##0.0_);(#,##0.0)'), { cat: 'number', decimals: 1, thousands: true, negative: 'paren' });
  assert.deepEqual(describeCode('0.00%'), { cat: 'percent', decimals: 2 });
  assert.equal(describeCode('yyyy-mm-dd').cat, 'date');
  assert.equal(describeCode('0"점"').cat, 'custom');
  // 목록에 있는 코드는 모두 해석 가능해야 함
  for (const code of [...DATE_TYPES, ...TIME_TYPES, ...CUSTOM_LIST, ...SPECIAL_TYPES.map((x) => x.code), ...FRACTION_TYPES.map((x) => x.code)]) {
    const c = code === 'G/표준' ? 'General' : code;
    assert.equal(typeof formatCode(1234.5, c).text, 'string', code);
    assert.equal(typeof formatCode(-3, c).text, 'string', code);
  }
});

test('텍스트 나누기', () => {
  assert.deepEqual(splitDelimited('a,b,,c', { comma: true }), ['a', 'b', '', 'c']);
  assert.deepEqual(splitDelimited('a  b   c', { space: true, consecutive: true }), ['a', 'b', 'c']);
  assert.deepEqual(splitDelimited('"x,y",z', { comma: true }), ['x,y', 'z']);
  assert.deepEqual(splitDelimited('a|b', { other: '|' }), ['a', 'b']);
  assert.deepEqual(splitDelimited('a\tb;c', { tab: true, semicolon: true }), ['a', 'b', 'c']);
  assert.deepEqual(splitFixed('ABC  DEF  12', [5, 10]), ['ABC', 'DEF', '12']);
  assert.deepEqual(suggestBreaks(['홍길동  서울  30', '김철수  부산  25']), [5, 9]);
  assert.equal(parseDateOrder('20240315'), '2024-03-15');
  assert.equal(parseDateOrder('03/15/2024', 'MDY'), '2024-03-15');
  assert.equal(parseDateOrder('15.03.24', 'DMY'), '2024-03-15');
  assert.equal(parseDateOrder('2024년 3월 15일'), '2024-03-15');
  assert.equal(parseDateOrder('15-Mar-2024', 'DMY'), '2024-03-15');
  assert.equal(parseDateOrder('2024-02-30'), null);
  assert.equal(convertPart('1234-'), '-1234');
  assert.equal(convertPart('=SUM(A1)'), "'=SUM(A1)");
  assert.equal(convertPart('007', 'text'), "'007");
  assert.equal(convertPart('20240315', 'date', 'YMD'), '2024-03-15');
  assert.equal(convertPart('x', 'skip'), null);
});

function tableBook() {
  const wb = new Workbook();
  const rows = [['제품', '지역', '금액'], ['사과', '서울', 100], ['배', '부산', 200], ['감', '서울', 300]];
  wb.transact(() => rows.forEach((r, i) => r.forEach((v, j) => wb.setInput(0, i, j, String(v)))));
  wb.sheets[0].tables = [{ id: 't1', name: '표1', r1: 0, c1: 0, r2: 3, c2: 2, header: true, totals: false, style: 'TableStyleMedium2', filter: { criteria: {}, hidden: {} } }];
  wb.invalidate();
  return wb;
}

test('표: 구조적 참조 · SUBTOTAL', () => {
  assert.deepEqual(parseSpec('금액'), { areas: ['data'], c1: '금액', c2: null });
  assert.deepEqual(parseSpec('[#Headers],[지역]'), { areas: ['headers'], c1: '지역', c2: null });
  assert.deepEqual(parseSpec('@금액'), { areas: ['thisrow'], c1: '금액', c2: null });
  assert.deepEqual(parseSpec('[제품]:[금액]'), { areas: ['data'], c1: '제품', c2: '금액' });
  const wb = tableBook();
  assert.deepEqual(resolveStructRef(wb, '표1', '#모두', null), { si: 0, r1: 0, c1: 0, r2: 3, c2: 2 });
  assert.deepEqual(resolveStructRef(wb, null, '@금액', { si: 0, r: 2, c: 1 }), { si: 0, r1: 2, c1: 2, r2: 2, c2: 2 });
  assert.equal(resolveStructRef(wb, '표1', '없는열', null), null);
  assert.equal(canonicalRef(null, '@금액', '표1'), '표1[[#This Row],[금액]]');
  assert.equal(canonicalRef('표1', '금액'), '표1[금액]');
  wb.transact(() => {
    wb.setInput(0, 0, 5, '=SUM(표1[금액])');
    wb.setInput(0, 1, 5, '=SUBTOTAL(109,표1[금액])');
    wb.setInput(0, 2, 5, '=표1[[#Headers],[지역]]');
    wb.setInput(0, 3, 5, '=STDEV(1,2,3,4)');
  });
  assert.equal(wb.getValue(0, 0, 5), 600);
  assert.equal(wb.getValue(0, 2, 5), '지역');
  assert.ok(Math.abs(wb.getValue(0, 3, 5) - 1.2909944487) < 1e-9);
  wb.sheets[0].tables[0].filter.hidden = { 2: true };
  wb.invalidate();
  assert.equal(wb.getValue(0, 1, 5), 400); // 필터로 숨긴 행 제외
  assert.equal(wb.getValue(0, 0, 5), 600);
  // 행 삽입 시 표 범위도 이동
  wb.transact(() => wb.insertRows(0, 0, 2));
  assert.deepEqual([wb.sheets[0].tables[0].r1, wb.sheets[0].tables[0].r2], [2, 5]);
  assert.equal(evaluateFormula(parse('SUM(표1[금액])'), wb.ctxFor(0)), 600);
});

test('표: 자동 확장 · 이름 · 스타일', () => {
  const s = tableBook().sheets[0];
  assert.deepEqual(expansionFor(s, { r1: 4, c1: 1, r2: 4, c2: 1 }), [{ id: 't1', r2: 4 }]);
  assert.deepEqual(expansionFor(s, { r1: 2, c1: 3, r2: 2, c2: 3 }), [{ id: 't1', c2: 3 }]);
  assert.deepEqual(expansionFor(s, { r1: 6, c1: 0, r2: 6, c2: 0 }), []);
  assert.deepEqual(uniqueNames(['a', '', 'A', 'b']), ['a', '열2', 'A2', 'b']);
  assert.equal(validTableName('표1'), true);
  assert.equal(validTableName('A1'), false);
  assert.equal(validTableName('판매 목록'), false);
  assert.equal(TABLE_STYLES.length, 21);
  assert.equal(normalizeStyleName('TableStyleMedium9'), 'TableStyleMedium2');
  const t = s.tables[0];
  assert.equal(tableCellStyle(t, 0, 0).fill, '#4472c4');
  assert.ok(tableCellStyle(t, 1, 0).fill);
  assert.equal(tableCellStyle(t, 2, 0).fill, undefined);
});

test('표 xlsx 왕복', () => {
  const wb = tableBook();
  wb.sheets[0].tables[0] = { ...wb.sheets[0].tables[0], r2: 4, totals: true, totalsFns: { 2: 'sum' }, style: 'TableStyleLight3' };
  wb.transact(() => {
    wb.setInput(0, 4, 0, '요약');
    wb.setInput(0, 4, 2, '=SUBTOTAL(109,표1[금액])');
    wb.setInput(0, 1, 3, '=[@금액]*2');
  });
  wb.sheets[0].tables[0].c2 = 3;
  wb.transact(() => wb.setInput(0, 0, 3, '두배'));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  const t = back.sheets[0].tables[0];
  assert.deepEqual([t.name, t.r1, t.c1, t.r2, t.c2, t.totals, t.style, t.totalsFns[2]], ['표1', 0, 0, 4, 3, true, 'TableStyleLight3', 'sum']);
  assert.equal(back.getValue(0, 4, 2), 600);
  assert.equal(back.getRaw(0, 1, 3), '=표1[[#This Row],[금액]]*2');
  assert.equal(back.getValue(0, 1, 3), 200);
});
