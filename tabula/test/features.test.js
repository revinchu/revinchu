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

test('조건부 서식: 규칙 종류 · 우선순위 · 중지', async () => {
  const { prepareCond, condFormatAt, describeCond, periodRange } = await import('../src/condfmt.js');
  const wb = new Workbook();
  wb.transact(() => [5, 50, 95, 120, 30].forEach((v, i) => wb.setInput(0, i, 0, String(v))));
  wb.transact(() => { wb.setInput(0, 0, 1, '사과나무'); wb.setInput(0, 1, 1, '배'); wb.setInput(0, 2, 1, '=1/0'); });
  const s = wb.sheets[0];
  const fmt = (r, c) => condFormatAt(prepareCond(wb, 0), wb, 0, r, c, wb.getValue(0, r, c));
  s.cond = [
    { r1: 0, c1: 0, r2: 4, c2: 0, type: 'formula', formula: '=A1>100', style: { fill: '#ff0000' } },
    { r1: 0, c1: 0, r2: 4, c2: 0, type: 'gt', v1: '90', style: { fill: '#00ff00', bold: true } },
  ];
  assert.deepEqual(fmt(3, 0).style, { fill: '#ff0000', bold: true }); // 앞 규칙이 이기고, 겹치지 않는 서식은 합쳐짐
  assert.deepEqual(fmt(2, 0).style, { fill: '#00ff00', bold: true });
  s.cond[0].stopIfTrue = true;
  assert.deepEqual(fmt(3, 0).style, { fill: '#ff0000' });
  s.cond = [
    { r1: 0, c1: 0, r2: 4, c2: 0, type: 'bottom', v1: '2', style: { color: '#111111' } },
    { r1: 0, c1: 0, r2: 4, c2: 0, type: 'ge', v1: '=A$2', style: { italic: true } },
    { r1: 0, c1: 1, r2: 4, c2: 1, type: 'begins', v1: '사과', style: { bold: true } },
    { r1: 0, c1: 1, r2: 4, c2: 1, type: 'errors', style: { strike: true } },
    { r1: 0, c1: 1, r2: 4, c2: 1, type: 'blank', style: { fill: '#eeeeee' } },
  ];
  assert.deepEqual(fmt(0, 0).style, { color: '#111111' });
  assert.deepEqual(fmt(1, 0).style, { italic: true });
  assert.deepEqual(fmt(4, 0).style, { color: '#111111' });
  assert.deepEqual(fmt(0, 1).style, { bold: true });
  assert.deepEqual(fmt(2, 1).style, { strike: true });
  assert.deepEqual(fmt(4, 1).style, { fill: '#eeeeee' });
  s.cond = [{ r1: 0, c1: 0, r2: 4, c2: 0, type: 'icons', icons: '3Arrows' }, { r1: 0, c1: 0, r2: 4, c2: 0, type: 'bar', color: '#638ec6' }];
  assert.equal(fmt(3, 0).icon, 'arrowUpGreen');
  assert.equal(fmt(0, 0).icon, 'arrowDownRed');
  assert.ok(Math.abs(fmt(3, 0).bar.pct - 100) < 1e-9);
  assert.equal(describeCond({ type: 'between', v1: '1', v2: '9' }), '셀 값 다음 값 사이 1 및 9');
  const [a, b] = periodRange('thisMonth', new Date(2024, 2, 15));
  assert.deepEqual([a, b], [45352, 45383]);
});

test('조건부 서식 xlsx 왕복 (새 규칙 종류)', () => {
  const wb = new Workbook();
  wb.transact(() => wb.setInput(0, 0, 0, '1'));
  wb.sheets[0].cond = [
    { r1: 0, c1: 0, r2: 9, c2: 0, type: 'formula', formula: '=$A1>5', style: { fill: '#ffc7ce', underline: true }, stopIfTrue: true },
    { r1: 0, c1: 0, r2: 9, c2: 0, type: 'notBetween', v1: '1', v2: '=$B$1', style: { color: '#9c0006', bt: true, bb: true, bl: true, br: true } },
    { r1: 0, c1: 1, r2: 9, c2: 1, type: 'ends', v1: '다', style: { bold: true } },
    { r1: 0, c1: 1, r2: 9, c2: 1, type: 'date', period: 'lastWeek', style: { fill: '#c6efce' } },
    { r1: 0, c1: 2, r2: 9, c2: 2, type: 'top', v1: '10', percent: true, style: { fill: '#ffeb9c' } },
    { r1: 0, c1: 2, r2: 9, c2: 2, type: 'icons', icons: '3TrafficLights1', reverse: true, iconOnly: true },
    { r1: 0, c1: 3, r2: 9, c2: 3, type: 'noErrors', style: { italic: true } },
  ];
  const back = readXlsx(writeXlsx(wb)).data.sheets[0].cond;
  assert.deepEqual(back.map((r) => r.type), ['formula', 'notBetween', 'ends', 'date', 'top', 'icons', 'noErrors']);
  assert.equal(back[0].formula, '=$A1>5');
  assert.equal(back[0].stopIfTrue, true);
  assert.deepEqual(back[0].style, { fill: '#ffc7ce', underline: true });
  assert.deepEqual([back[1].v1, back[1].v2], ['1', '=$B$1']);
  assert.equal(back[1].style.bt, true);
  assert.equal(back[3].period, 'lastWeek');
  assert.equal(back[4].percent, true);
  assert.deepEqual([back[5].icons, back[5].reverse, back[5].iconOnly], ['3TrafficLights1', true, true]);
});

test('슬라이서 · 피벗 테이블 xlsx 왕복 (엑셀 형식)', async () => {
  const { buildPivot, pivotSourceData, resolvePivot } = await import('../src/pivot.js');
  const { unzip } = await import('../src/zip.js');
  const wb = new Workbook();
  const data = [['지역', '제품', '수량'], ['서울', '사과', 10], ['부산', '배', 5], ['서울', '감', 8], ['대구', '사과', 3]];
  wb.transact(() => data.forEach((row, i) => row.forEach((v, j) => wb.setInput(0, i, j, String(v)))));
  wb.sheets[0].tables = [{ id: 't1', name: '표1', r1: 0, c1: 0, r2: 4, c2: 2, header: true, totals: false, style: 'TableStyleMedium2', filter: { criteria: { 1: ['사과'] }, hidden: { 2: true, 3: true } } }];
  wb.sheets[0].slicers = [{ id: 's1', caption: '제품', source: { kind: 'table', table: '표1', column: '제품' }, columns: 2, color: 'green', x: 250, y: 10, w: 180, h: 120 }];
  const at = wb.addSheet('피벗1');
  const def = { table: '표1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 4, c2: 2 }, rowField: 0, colField: null, valueField: 2, agg: 'sum', fieldNames: { row: '지역', col: null, val: '수량' }, filters: { 제품: ['사과', '감'] } };
  const { def: d, rows } = resolvePivot(pivotSourceData(wb, def).rows, def);
  wb.transact(() => buildPivot(rows, d).forEach((row, r) => row.forEach((cd, c) => { if (cd) wb.setCellData(at, r, c, cd); })));
  wb.sheets[at].pivot = def;
  wb.sheets[at].slicers = [{ id: 's2', caption: '제품', source: { kind: 'pivot', self: true, field: '제품' }, x: 200, y: 0, w: 180, h: 120 }];
  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  const text = (p) => new TextDecoder().decode(files[p]);
  assert.match(text('xl/workbook.xml'), /x15:slicerCaches/);
  assert.match(text('xl/workbook.xml'), /<pivotCaches><pivotCache cacheId="1"/);
  assert.match(text('xl/pivotTables/pivotTable1.xml'), /<item h="1" x="1"\/>/); // 배(숨김)
  assert.match(text('xl/drawings/drawing1.xml'), /Requires="sle15"/);
  const back = readXlsx(bytes);
  assert.deepEqual(back.warnings, []);
  const [s1, s2] = back.data.sheets;
  assert.deepEqual(s1.slicers.map((x) => [x.caption, x.source, x.columns, x.style]), [['제품', { kind: 'table', table: '표1', column: '제품' }, 2, 'SlicerStyleLight6']]);
  assert.deepEqual([s1.slicers[0].x, s1.slicers[0].y, s1.slicers[0].w, s1.slicers[0].h], [250, 10, 180, 120]);
  assert.deepEqual(s1.tables[0].filter.criteria, { 1: ['사과'] });
  assert.deepEqual(s2.slicers[0].source, { kind: 'pivot', field: '제품', pivots: [{ sheet: '피벗1', name: '피벗 테이블1' }] });
  assert.equal(s2.pivot.table, '표1');
  assert.deepEqual(s2.pivot.filters, { 제품: ['감', '사과'] });
  assert.deepEqual([s2.pivot.rows, s2.pivot.cols, s2.pivot.values], [['지역'], [], [{ field: '수량', agg: 'sum' }]]);
});

test('셀 그림 · IMAGE 함수 · xlsx 왕복', () => {
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const wb = new Workbook();
  wb.transact(() => {
    wb.setCellData(0, 0, 0, { raw: '', image: { src: png, alt: '로고' } });
    wb.setStyle(0, 0, 0, { bold: true });
    wb.setInput(0, 1, 0, '=IMAGE("https://example.com/a.png","대체",1)');
    wb.setInput(0, 2, 0, '=ISERROR(A1)');
  });
  assert.equal(wb.getCell(0, 0, 0).image.alt, '로고'); // 서식을 바꿔도 그림 유지
  assert.equal(wb.getValue(0, 0, 0).type, 'image');
  assert.equal(wb.getValue(0, 1, 0).src, 'https://example.com/a.png');
  assert.equal(wb.getValue(0, 1, 0).alt, '대체');
  const bytes = writeXlsx(wb);
  const { data } = readXlsx(bytes);
  const cells = data.sheets[0].cells;
  assert.equal(cells['0,0'].image.src, png);
  assert.equal(cells['0,0'].image.alt, '로고');
  assert.equal(cells['0,0'].style.bold, true);
  assert.equal(cells['1,0'].raw, '=IMAGE("https://example.com/a.png","대체",1)');
});
