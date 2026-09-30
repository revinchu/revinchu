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
  assert.equal(pivotOf(wb, { ...base, filters: { 매체: ['네이버', '구글'] } }).grid[0][1], '(모두)'); // 원본 항목을 모두 고름
});

test('다른 시트 범위의 암시적 교차 · 오류 조건 · 범위 자리의 오류 · 순환 참조의 파일 값', () => {
  const wb = new Workbook();
  wb.addSheet('S2');
  wb.setInput(1, 1, 2, '7'); wb.setInput(1, 1, 3, '8');
  wb.setInput(0, 4, 2, '=@S2!C2:D2');
  wb.setInput(0, 4, 3, '=@S2!C2:D2');
  assert.equal(wb.getValue(0, 4, 2), 7);
  assert.equal(wb.getValue(0, 4, 3), 8);
  wb.setInput(0, 0, 0, '1'); wb.setInput(0, 1, 0, '=NA()'); wb.setInput(0, 0, 1, '5'); wb.setInput(0, 1, 1, '7');
  assert.equal(calc(wb, '=SUMIFS(B1:B2,A1:A2,NA())'), 7); // 오류 조건은 같은 오류 칸과 맞음
  assert.equal(calc(wb, '=SUMIFS(B1:B2,A1:A2,1/0)'), 0);
  assert.equal(calc(wb, '=COUNTIF(A1:A2,NA())'), 1);
  assert.equal(calc(wb, '=SUMIFS(#REF!,#REF!,1)')?.code, '#REF!');
});

test('외부 통합 문서 참조: [1]시트!A1 은 저장된 외부 값, 닫힌 파일 범위의 SUMIF 는 #VALUE!, 저장하면 연결 보존', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const { zip, unzip, textOf } = await import('../src/zip.js');
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const bytes = zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`,
    'xl/externalLinks/externalLink1.xml': `<?xml version="1.0"?><externalLink ${NS}><externalBook r:id="rId1"><sheetNames><sheetName val="원본"/><sheetName val="다른 시트"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="5"><cell r="D5"><v>467</v></cell><cell r="E5" t="s"><v>글자</v></cell></row></sheetData><sheetData sheetId="1"><row r="1"><cell r="A1"><v>3</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`,
    'xl/externalLinks/_rels/externalLink1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="file:///C:/raw.xlsx" TargetMode="External"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${NS}><sheetData><row r="1"><c r="A1"><f>[1]원본!D5</f><v>467</v></c><c r="B1" t="str"><f>[1]원본!E5</f><v>글자</v></c><c r="C1"><f>'[1]다른 시트'!A1*2</f><v>6</v></c><c r="D1" t="e"><f>SUMIF([1]원본!D1:D9,1,[1]원본!D1:D9)</f><v>#VALUE!</v></c></row></sheetData></worksheet>`,
  });
  const wb = new Workbook(readXlsx(bytes).data);
  wb.invalidate();
  assert.equal(wb.getValue(0, 0, 0), 467);
  assert.equal(wb.getValue(0, 0, 1), '글자');
  assert.equal(wb.getValue(0, 0, 2), 6);
  assert.equal(wb.getValue(0, 0, 3)?.code, '#VALUE!');
  assert.equal(wb.ownSheetCount(), 1);
  assert.equal(wb.addSheet('새 시트'), 1); // 새 시트는 외부 값 시트 앞에
  const files = unzip(writeXlsx(wb));
  const wx = textOf(files['xl/workbook.xml']);
  assert.equal((wx.match(/<sheet /g) || []).length, 2);
  assert.match(wx, /<externalReferences><externalReference r:id="rId\d+"\/><\/externalReferences>/);
  assert.ok(files['xl/externalLinks/externalLink1.xml']);
  assert.match(textOf(files['xl/worksheets/sheet1.xml']), /<f>\[1\]원본!D5<\/f>/);
});

test('피벗: 필드 머리글 숨기기 · 값만 있는 피벗 · 열 항목 값 기준 정렬(빈 값 = 0)', () => {
  const wb = new Workbook();
  put(wb, [['매체', '월', '비용'], ['a', '11', '5'], ['b', '11', '9'], ['b', '12', '1'], ['c', '12', '4'], ['d', '11', '2']]);
  const base = { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 5, c2: 2 }, pages: [], filters: {}, layout: 'tabular', top: 0, left: 5 };
  const only = pivotOf(wb, { ...base, rows: [], cols: [], values: [{ field: '비용', agg: 'sum' }, { field: '비용', agg: 'count', name: '건수' }] }).grid;
  assert.deepEqual(only, [['합계 : 비용', '건수'], ['21', '5']]);
  const hid = pivotOf(wb, { ...base, rows: ['매체'], cols: [], values: [{ field: '비용', agg: 'sum' }], showHeaders: false }).grid;
  assert.equal(hid[0][0], '');
  // 12월 열 값 기준 내림차순: c(4) > b(1) > a · d(빈 값 = 0, 저장 순서)
  const s = pivotOf(wb, { ...base, rows: ['매체'], cols: ['월'], values: [{ field: '비용', agg: 'sum' }], sort: { 매체: { dir: 'desc', by: 0, at: [['월', '12']] } } }).grid;
  assert.deepEqual(s.slice(2, 6).map((r) => r[0]), ['c', 'b', 'a', 'd']);
});

test('시간 기본 형식 (h:mm:ss 는 24시간) · 표 참조 표기 표1[#All]', async () => {
  const { BUILTIN_FMT } = await import('../src/xlsx.js');
  const { canonicalRef } = await import('../src/tables.js');
  assert.equal(formatValue(0.65625, BUILTIN_FMT[21]).text, '15:45:00');
  assert.equal(formatValue(0.65625, BUILTIN_FMT[18]).text, '오후 3:45');
  assert.equal(canonicalRef('k', '#All'), 'k[#All]');
});

test('피벗: Σ 값을 행 영역에 (dataOnRows) · 같은 필드를 행/열에 둔 두 피벗', () => {
  const wb = new Workbook();
  put(wb, [['매체', '비용', '클릭'], ['a', '10', '1'], ['b', '20', '2'], ['a', '5', '3']]);
  const base = { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, pages: [], filters: {}, layout: 'tabular', top: 0, left: 5, grandRows: false, grandCols: false };
  const values = [{ field: '비용', agg: 'sum', name: '광고비' }, { field: '클릭', agg: 'sum', name: '클릭수' }];
  // 열 필드 = 매체, 값은 행으로: 값 이름이 행 머리글 ('값' 열)
  const g = pivotOf(wb, { ...base, rows: [], cols: ['매체'], values, valuesOnRows: true }).grid;
  assert.deepEqual(g, [['', '매체', ''], ['값', 'a', 'b'], ['광고비', '15', '20'], ['클릭수', '4', '2']]);
  // 같은 원본에서 매체를 행에 둔 피벗 (묶음 결과를 공유해도 행 · 열이 바뀌면 안 됨)
  const r = pivotOf(wb, { ...base, rows: ['매체'], cols: [], values }).grid;
  assert.deepEqual(r.slice(1).map((x) => x[0]), ['a', 'b']);
  // 행 필드 + 값 행: 항목마다 값 필드 수만큼 행, 총합계는 '전체 …'
  const t = pivotOf(wb, { ...base, rows: ['매체'], cols: [], values, valuesOnRows: true, grandRows: true }).grid;
  assert.deepEqual(t, [['매체', '값', ''], ['a', '광고비', '15'], ['', '클릭수', '4'], ['b', '광고비', '20'], ['', '클릭수', '2'], ['전체 광고비', '', '35'], ['전체 클릭수', '', '6']]);
});

test('순환 참조: 고리 전체가 파일에 저장된 값을 유지 · 숫자 비교는 유효 숫자 15자리', () => {
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '=B1+1');
  wb.setInput(0, 0, 1, '=A1*2');
  wb.getCell(0, 0, 0).cached = 5;
  wb.getCell(0, 0, 1).cached = 10;
  wb.invalidate(0);
  assert.equal(wb.getValue(0, 0, 1), 10);
  assert.equal(wb.getValue(0, 0, 0), 5);
  wb.setInput(0, 2, 0, '3977975.925');
  wb.setInput(0, 2, 1, '=3977975.9249999993');
  assert.equal(calc(wb, '=A3=B3'), true);
  assert.equal(calc(wb, '=MAX(A3:B3)>B3'), false);
});

test('열 전체 참조(B:B) 수식도 모양이 같으면 AST 공유 · 파일의 시트!#REF! 수식 모양 재사용', async () => {
  const { formulaShifter } = await import('../src/workbook.js');
  assert.equal(formulaShifter('=SUMIFS(K:K,$E:$E,D6)+SUM(3:3)')(1, 1), '=SUMIFS(L:L,$E:$E,E7)+SUM(4:4)');
  const wb = new Workbook();
  put(wb, [['a', '1', 'x'], ['b', '2', 'y'], ['a', '3', 'x']]);
  wb.transact(() => { for (let r = 0; r < 3; r++) wb.setInput(0, r, 4, `=SUMIFS(B:B,A:A,A${r + 1})`); });
  assert.deepEqual([0, 1, 2].map((r) => wb.getValue(0, r, 4)), [4, 2, 4]);
  assert.equal(wb.getCell(0, 0, 4).ast, wb.getCell(0, 2, 4).ast);
  const { readXlsx } = await import('../src/xlsx.js');
  const { zip } = await import('../src/zip.js');
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rows = [1, 2, 3].map((r) => `<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}" t="e"><f>IF(A${r}=1,SUM(S!#REF!),A${r}*2)</f><v>#REF!</v></c></row>`).join('');
  const bytes = zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${NS}><sheetData>${rows}</sheetData></worksheet>`,
  });
  const s = readXlsx(bytes).data.sheets[0];
  const first = s.cells.get('0,1').raw;
  assert.equal(s.cells.get('2,1').raw, first.replace(/A1/g, 'A3'));
});

test('한국어 서식: [$-412]ddd = 금 · [$-F800] 시스템 긴 날짜 · 1900년 3월 전 날짜 저장', async () => {
  const { formatCode } = await import('../src/format.js');
  const { numberRaw, isoSerial } = await import('../src/xlsx.js');
  assert.equal(formatCode(44228, '[$-412]ddd').text, '월');
  assert.equal(formatCode(44228, '[$-ko-KR]dddd').text, '월요일');
  assert.equal(formatCode(44228, '[$-412]mmm').text, '2월');
  assert.equal(formatCode(44228, 'ddd').text, 'Mon');
  assert.equal(formatCode(44228, '[$-F800]dddd, mmmm dd, yyyy').text, '2021년 2월 1일 월요일');
  const wb = new Workbook();
  assert.equal(calc(wb, '=TEXT(44228,"[$-412]ddd")'), '월');
  assert.equal(numberRaw(28, { numFmt: 'date' }), '1900-01-28');
  assert.equal(isoSerial('1900-01-28'), 28);
  assert.equal(isoSerial('2021-02-01T12:00:00'), 44228.5);
});

test('SUM(A1): 참조한 칸의 글자는 무시 (직접 인수 "" 만 #VALUE!) · INDIRECT 배열 수식이 자기 원본을 읽을 때 순환으로 보지 않음', () => {
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '=""'); wb.setInput(0, 0, 1, '5'); wb.setInput(0, 0, 2, 'abc');
  assert.equal(calc(wb, '=SUM(A1,B1)'), 5);
  assert.equal(calc(wb, '=SUM(A1)'), 0);
  assert.equal(calc(wb, '=AVERAGE(C1,B1)'), 5);
  assert.equal(calc(wb, '=SUM("",B1)')?.code, '#VALUE!');
  assert.equal(calc(wb, '=A1+B1')?.code, '#VALUE!');
  const w2 = new Workbook();
  put(w2, [['1', '10'], ['2', '20']]);
  w2.setInput(0, 5, 0, 'A1:B2');
  w2.setInput(0, 10, 0, '=INDIRECT(A6)');
  assert.equal(w2.getValue(0, 10, 0), 1);
  assert.equal(calc(w2, '=SUM(B11:B12)'), 30);
});

test('INDIRECT(ADDRESS()) 는 한 칸 · 보고서 필터만 있는 피벗 · 표 밖 [#This Row] 는 #VALUE! · 오류 값 입력', async () => {
  const { parse, mayReturnArray } = await import('../src/formula.js');
  assert.equal(mayReturnArray(parse('SUMIFS(A:A,B:B,INDIRECT(ADDRESS(1,COLUMN(),1,1)))')), false);
  assert.equal(mayReturnArray(parse('INDIRECT("A1:B2")')), true);
  const wb = new Workbook();
  put(wb, [['월', '비용'], ['7월', '1'], ['8월', '2']]);
  const g = pivotOf(wb, { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: [], cols: [], values: [], pages: ['월'], filters: { 월: ['7월'] }, layout: 'compact', top: 0, left: 4 }).grid;
  assert.deepEqual(g, [['월', '7월']]);
  wb.setInput(0, 10, 0, '#n/a');
  assert.equal(calc(wb, '=ISNA(A11)'), true);
});

test('find ignores full/half width unless 전자/반자 구분', async () => {
  const { findRegex, replaceText, foldWidth } = await import('../src/find.js');
  assert.equal(foldWidth('ＡＢＣ１２３　'), 'ABC123 ');
  assert.ok(findRegex(foldWidth('ＡＢ'), {}).test(foldWidth('xabx')));
  assert.equal(replaceText('ＡＢ가AB', { text: 'ab' }, 'Z'), 'Z가Z');
  assert.equal(replaceText('ＡＢ가AB', { text: 'ab', matchByte: true }, 'Z'), 'ＡＢ가Z');
});

test('named cell styles round-trip through xlsx', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.cellStyles = [{ name: '20% - 강조색1 2', style: { fill: '#dae3f3', color: '#000000', bold: true } }, { name: '백분율 2', style: { numFmt: 'percent' }, builtinId: 5 }];
  wb.setCellData(0, 0, 0, { raw: '1' });
  const back = readXlsx(writeXlsx(wb)).data;
  assert.deepEqual(back.cellStyles.map((c) => c.name), ['20% - 강조색1 2', '백분율 2']);
  assert.equal(back.cellStyles[0].style.fill, '#dae3f3');
  assert.equal(back.cellStyles[0].style.bold, true);
  assert.equal(back.cellStyles[1].builtinId, 5);
});

test('text wizard advanced number recognition and hierarchical treemap', async () => {
  const { convertPart } = await import('../src/textsplit.js');
  assert.equal(convertPart('1.234,5', 'general', 'YMD', { decimal: ',', thousand: '.', trailingMinus: true }), '1234.5');
  assert.equal(convertPart('12-', 'general'), '-12');
  const { resolveChart } = await import('../src/chart.js');
  const rows = [['상위', '구분', '값'], ['2026.04', '4월 집행', 30], [null, '4월 예상', 20], ['Total', '전년', 50], [null, '4월', 10]];
  const d = resolveChart({ type: 'treemap', range: {} }, { range: () => rows, values: () => [] });
  assert.deepEqual(d.catLevels[0].map((g) => g.text), ['2026.04', 'Total']);
  assert.deepEqual(d.series[0].values, [30, 20, 50, 10]);
});

test('range chart with a merged-style outer label column uses multi-level categories', async () => {
  const { chartData } = await import('../src/chart.js');
  const d = chartData([['상위', '구분', '값'], [2026.04, '4월 집행', 30], [null, '4월 예상', 20], ['Total', '전년', 50], [null, '4월', 10]]);
  assert.deepEqual(d.categories, ['4월 집행', '4월 예상', '전년', '4월']);
  assert.equal(d.series.length, 1);
  assert.deepEqual(d.catLevels[0].map((g) => g.text), ['2026.04', 'Total']);
});

test('workbook theme is written to theme1.xml and read back', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.theme = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D'];
  wb.setCellData(0, 0, 0, { raw: '1' });
  const a = readXlsx(writeXlsx(wb)).data;
  assert.deepEqual(a.theme, wb.theme);
  // 파일의 테마 XML 을 유지하면서 색만 바꿈
  const wb2 = new Workbook();
  wb2.load(a);
  wb2.theme = [...wb.theme];
  wb2.theme[4] = 'FF0000';
  const b = readXlsx(writeXlsx(wb2)).data;
  assert.equal(b.theme[4], 'FF0000');
  assert.equal(b.theme[5], 'E97132');
});

test('pivot date filters (dynamic periods and custom dates) match Excel', async () => {
  const { dateFilterMatch } = await import('../src/pivot.js');
  const { serialOf } = await import('../src/format.js');
  const today = serialOf(2026, 4, 15); // 수요일
  assert.ok(dateFilterMatch('today', today, null, null, today));
  assert.ok(dateFilterMatch('yesterday', today - 1, null, null, today));
  assert.ok(dateFilterMatch('thisWeek', serialOf(2026, 4, 12), null, null, today)); // 일요일 시작
  assert.ok(!dateFilterMatch('thisWeek', serialOf(2026, 4, 11), null, null, today));
  assert.ok(dateFilterMatch('lastMonth', serialOf(2026, 3, 31), null, null, today));
  assert.ok(dateFilterMatch('thisQuarter', serialOf(2026, 6, 30), null, null, today));
  assert.ok(dateFilterMatch('lastYear', serialOf(2025, 1, 1), null, null, today));
  assert.ok(dateFilterMatch('yearToDate', serialOf(2026, 1, 2), null, null, today));
  assert.ok(!dateFilterMatch('yearToDate', serialOf(2026, 5, 2), null, null, today));
  assert.ok(dateFilterMatch('Q2', serialOf(2024, 5, 1)));
  assert.ok(dateFilterMatch('M12', serialOf(2023, 12, 25)));
  assert.ok(dateFilterMatch('dateBetween', serialOf(2026, 4, 3), serialOf(2026, 4, 1), serialOf(2026, 4, 5)));
  assert.ok(!dateFilterMatch('today', '글자', null, null, today));
});

test('xlsx: 셀 요소의 서식은 행 서식과 섞이지 않음(엑셀), 읽을 수 없는 수식은 저장된 값', async () => {
  const { readXlsx } = await import('../src/xlsx.js');
  const { zip } = await import('../src/zip.js');
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const bytes = zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<?xml version="1.0"?><styleSheet ${NS}><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="3"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf fontId="1" fillId="0" borderId="0" applyFont="1"><alignment vertical="center"/></xf></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${NS}><sheetData><row r="1" s="1" customFormat="1"><c r="A1" s="2" t="inlineStr"><is><t>제목</t></is></c><c r="B1"><f>[2]!표[[#This Row],[열]]</f><v>7</v></c></row></sheetData></worksheet>`,
  });
  const wb = new Workbook(readXlsx(bytes).data);
  assert.equal(!!wb.styleAt(0, 0, 0).wrap, false); // 행의 줄 바꿈이 셀로 새지 않음
  assert.equal(wb.styleAt(0, 0, 0).bold, true);
  assert.equal(!!wb.styleAt(0, 0, 5).wrap, true); // 빈 칸은 행 서식
  wb.invalidate();
  assert.equal(wb.getValue(0, 0, 1), 7);
});

test('고급 필터: 조건 줄 안은 AND, 줄끼리 OR, 글자 조건은 시작 일치, 고유 레코드', async () => {
  const { advancedFilter } = await import('../src/analysis.js');
  const head = ['이름', '지역', '매출'];
  const rows = [['김', '서울', 100], ['이', '부산', 200], ['박', '서울', 300], ['김', '서울', 100]];
  assert.deepEqual(advancedFilter(head, rows, ['지역', '매출'], [['서울', '>150']]), [2]);
  assert.deepEqual(advancedFilter(head, rows, ['지역'], [['서울'], ['부산']]), [0, 1, 2, 3]);
  assert.deepEqual(advancedFilter(head, rows, ['지역'], [['서']], true), [0, 2]);
  assert.deepEqual(advancedFilter(head, rows, ['매출'], [[200]]), [1]);
});

test('아이콘(SVG 그림): PNG 대체 그림 + svgBlip 원본으로 저장하고 다시 읽음', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const { unzip, textOf } = await import('../src/zip.js');
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96" fill="#ff0000"><path d="M0 0h96v96z"/></svg>';
  const b64 = (s) => Buffer.from(s).toString('base64');
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const wb = new Workbook();
  wb.setSheetProp(0, 'images', [{ id: 'im1', name: '아이콘', x: 10, y: 10, w: 96, h: 96, src: `data:image/svg+xml;base64,${b64(svg)}`, png, icon: { vb: '0 0 96 96', body: '<path d="M0 0h96v96z"/>', fill: '#ff0000' } }]);
  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  assert.ok(Object.keys(files).some((f) => f.endsWith('.svg')));
  assert.ok(Object.keys(files).some((f) => f.endsWith('.png')));
  assert.match(textOf(files['xl/drawings/drawing1.xml']), /svgBlip/);
  const back = new Workbook(readXlsx(bytes).data).sheets[0].images[0];
  assert.match(back.src, /^data:image\/svg\+xml/);
  assert.match(back.png, /^data:image\/png/);
  assert.equal(back.icon.fill, '#ff0000');
});

test('셀 삽입 · 삭제 (밀기): 선택 열만 이동, 수식 참조도 띠 안에서만 조정', () => {
  const wb = new Workbook();
  wb.transact(() => {
    for (let r = 0; r < 5; r++) { wb.setCellData(0, r, 0, { raw: String(r + 1) }); wb.setCellData(0, r, 1, { raw: String((r + 1) * 10) }); }
    wb.setCellData(0, 0, 3, { raw: '=A3' });
    wb.setCellData(0, 1, 3, { raw: '=B3' });
    wb.setCellData(0, 2, 3, { raw: '=SUM(A1:A5)' });
  });
  // A2:A3 에 셀 삽입, 아래로 밀기 → A 열만 2칸 내려감, B 열은 그대로
  assert.equal(wb.shiftCells(0, { r1: 1, c1: 0, r2: 2, c2: 0 }, 'down'), null);
  assert.equal(wb.getValue(0, 0, 0), 1);
  assert.equal(wb.getCell(0, 1, 0), undefined);
  assert.equal(wb.getValue(0, 3, 0), 2);
  assert.equal(wb.getValue(0, 6, 0), 5);
  assert.equal(wb.getValue(0, 2, 1), 30);
  assert.equal(wb.getCell(0, 0, 3).raw, '=A5'); // A3 → A5
  assert.equal(wb.getCell(0, 1, 3).raw, '=B3'); // B 열은 그대로
  assert.equal(wb.getCell(0, 2, 3).raw, '=SUM(A1:A7)');
  // 다시 삭제, 위로 밀기
  assert.equal(wb.shiftCells(0, { r1: 1, c1: 0, r2: 2, c2: 0 }, 'up'), null);
  assert.equal(wb.getValue(0, 1, 0), 2);
  assert.equal(wb.getCell(0, 0, 3).raw, '=A3');
  // 병합 셀 일부는 막음
  wb.transact(() => wb.setSheetProp(0, 'merges', [{ r1: 0, c1: 0, r2: 0, c2: 1 }]));
  assert.match(wb.shiftCells(0, { r1: 0, c1: 0, r2: 0, c2: 0 }, 'down'), /병합/);
});

test('문서 속성 · 통합 문서 구조 보호 · 읽기 전용 권장 · 최종본: xlsx 왕복', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.props = { title: '월간 보고서', tags: '광고;리포트', category: '마케팅', creator: '홍길동', lockStructure: true, readOnlyRecommended: true, markedFinal: true };
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.props.title, '월간 보고서');
  assert.equal(back.props.tags, '광고;리포트');
  assert.equal(back.props.category, '마케팅');
  assert.equal(back.props.creator, '홍길동');
  assert.equal(back.props.lockStructure, true);
  assert.equal(back.props.readOnlyRecommended, true);
  assert.equal(back.props.markedFinal, true);
});

test('0 값 숨기기 (showZeros="0") · 새 문서 기본 글꼴: xlsx 왕복', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook({ sheets: [{ name: 'S', cells: { A1: 0 } }], defaultFont: { name: '굴림', size: 10 } });
  wb.transact(() => wb.setSheetProp(0, 'noZeros', true));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].noZeros, true);
  assert.equal(back.defaultFont?.name, '굴림');
});
