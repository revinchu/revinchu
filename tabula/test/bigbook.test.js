// 큰 통합 문서용 동작: 파일 계산 결과 재사용, 시트별 다시 계산, 가벼운 실행 취소, 빠른 시트 읽기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zip, unzip, textOf } from '../src/zip.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
import { ColBuilder, blockValue } from '../src/block.js';
import { CsvBlockReader } from '../src/csv.js';
import { Axis, hid, hidCount } from '../src/axis.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';

const book = () => new Workbook({
  sheets: [
    // 파일에서 연 것처럼: 수식 셀에 엑셀이 저장한 결과(cached) — 일부러 계산 결과와 다르게
    { name: 'raw', fileValues: true, cells: { '0,0': { raw: '10' }, '0,1': { raw: '=A1*2', cached: 99 } } },
    { name: 'dash', fileValues: true, cells: { '0,0': { raw: '=raw!B1+1', cached: 77 }, '1,0': { raw: '5' }, '1,1': { raw: '=A2*3', cached: 55 } } },
    { name: 'other', fileValues: true, cells: { '0,0': { raw: '=1+1', cached: 42 } } },
  ],
});

test('파일 계산 결과는 바뀌기 전까지 그대로 쓰고, 바뀐 시트와 참조하는 시트만 다시 계산', () => {
  const wb = book();
  assert.equal(wb.getValue(0, 0, 1), 99);
  assert.equal(wb.getValue(1, 0, 0), 77);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // dash 만 바꿈 → raw · other 는 그대로
  wb.transact(() => wb.setInput(1, 1, 0, '6'));
  assert.equal(wb.getValue(1, 1, 1), 18);
  assert.equal(wb.getValue(0, 0, 1), 99);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // raw 를 바꾸면 raw 를 참조하는 dash 도 다시 계산
  wb.transact(() => wb.setInput(0, 0, 0, '1'));
  assert.equal(wb.getValue(0, 0, 1), 2);
  assert.equal(wb.getValue(1, 0, 0), 3);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // 실행 취소해도 다시 계산된 값이 맞음
  wb.undo();
  assert.equal(wb.getValue(0, 0, 1), 20);
  assert.equal(wb.getValue(1, 0, 0), 21);
});

test('행 삽입 · 시트 추가 · 이름 바꾸기 실행 취소 (시트 단위 기록)', () => {
  const wb = book();
  wb.transact(() => wb.addSheet('new', 3));
  wb.transact(() => wb.insertRows(0, 0, 1));
  assert.equal(wb.getRaw(0, 1, 0), '10');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B2+1'); // 다른 시트의 참조도 조정
  assert.equal(wb.getValue(1, 0, 0), 21);
  assert.equal(wb.getValue(2, 0, 0), 42); // 참조하지 않는 시트는 파일 값 유지
  wb.undo();
  assert.equal(wb.getRaw(0, 0, 0), '10');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B1+1');
  wb.undo();
  assert.deepEqual(wb.sheets.map((s) => s.name), ['raw', 'dash', 'other']);
  assert.equal(wb.getRaw(0, 0, 0), '10');
  wb.redo();
  wb.redo();
  assert.equal(wb.sheets.length, 4);
  assert.equal(wb.getRaw(0, 1, 0), '10');

  wb.transact(() => wb.renameSheet(0, '원본'));
  assert.equal(wb.getRaw(1, 0, 0), '=원본!B2+1');
  assert.equal(wb.getValue(1, 0, 0), 21);
  wb.undo();
  assert.equal(wb.sheets[0].name, 'raw');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B2+1');
});

test('이름 정의 변경은 이름을 쓰는 시트만 다시 계산', () => {
  const wb = new Workbook({
    sheets: [
      { name: 'a', fileValues: true, cells: { '0,0': { raw: '=세율*2', cached: 1 } } },
      { name: 'b', fileValues: true, cells: { '0,0': { raw: '=3+4', cached: 8 } } },
    ],
    names: [{ name: '세율', ref: '=10', sheet: null }],
  });
  assert.equal(wb.getValue(1, 0, 0), 8);
  wb.transact(() => wb.setNames([{ name: '세율', ref: '=5' }]));
  assert.equal(wb.getValue(0, 0, 0), 10);
  assert.equal(wb.getValue(1, 0, 0), 8);
  wb.undo();
  assert.equal(wb.getValue(0, 0, 0), 20);
});

/** 최소 xlsx (직접 만든 시트 XML) */
function makeXlsx(sheetXml) {
  return zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>가</t></si><si><t>a&lt;b</t></si></sst>',
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetData>${sheetXml}</sheetData></worksheet>`,
  });
}

test('빠른 시트 읽기: 공유 문자열 · 인라인 문자열 · 빈 행 · 공유 수식 · 엔티티 · 눈금선', async () => {
  const bytes = makeXlsx(
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>x&amp;y</t></is></c></row>'
    + '<row r="2" ht="30" customHeight="1"/>'
    + '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" ref="B3:B4" si="0">A3&amp;"원"</f><v>2원</v></c><c r="C3" s="0"/></row>'
    + '<row r="4"><c r="A4"><v>3</v></c><c r="B4" t="str"><f t="shared" si="0"/><v>3원</v></c><c r="C4" t="b"><v>1</v></c><c r="D4" t="e"><v>#N/A</v></c></row>',
  );
  for (const read of [readXlsx, (b) => readXlsxAsync(b)]) {
    const { data } = await read(bytes);
    const s = data.sheets[0];
    const cells = Object.fromEntries(s.cells);
    assert.equal(cells['0,0'].raw, '가');
    assert.equal(cells['0,1'].raw, 'a<b');
    assert.equal(cells['0,2'].raw, 'x&y');
    assert.equal(cells['2,1'].raw, '=A3&"원"');
    assert.equal(cells['3,1'].raw, '=A4&"원"');
    assert.equal(cells['3,1'].cached, '3원');
    assert.equal(cells['3,2'].raw, 'TRUE');
    assert.equal(cells['3,3'].raw, '#N/A');
    assert.equal(s.rowHeights[1], 40);
    assert.equal(s.noGrid, true);
    assert.equal(s.fileValues, true);
    const wb = new Workbook();
    wb.load(data);
    assert.equal(wb.getValue(0, 3, 1), '3원');
    // 눈금선 숨김은 저장해도 유지
    const again = readXlsx(writeXlsx(wb)).data.sheets[0];
    assert.equal(again.noGrid, true);
  }
});

test('피벗 오류 값 표시 옵션 (showError) 과 사용자 지정 피벗 스타일', async () => {
  const rows = [['캠페인', '비용', '전환'], ['가', 100, 0], ['나', 50, 5]];
  const def = { rows: ['캠페인'], values: [{ field: 'CPA', agg: 'sum' }], calcFields: [{ name: 'CPA', formula: '비용/전환' }], layout: 'tabular' };
  const cellOf = (d, label) => {
    const res = resolvePivot(rows, d);
    const { grid } = computePivot(res, res.def);
    const row = grid.find((g) => g[0]?.raw === label);
    return row[1];
  };
  assert.equal(cellOf(def, '가').raw, '#DIV/0!');
  assert.equal(cellOf({ ...def, errorCaption: '' }, '가').raw, '');
  assert.equal(cellOf({ ...def, errorCaption: '-' }, '가').raw, "'-");
  assert.equal(cellOf({ ...def, errorCaption: '' }, '나').raw, '10');
  // 사용자 지정 스타일: 머리글 흰 글씨
  const custom = { header: { color: '#ffffff', fill: '#2f5597', bold: true } };
  const res = resolvePivot(rows, { ...def, style: '내 스타일', styleDef: custom });
  const head = computePivot(res, res.def).grid.flat().find((c) => c?.role?.startsWith('valueHead') || c?.role === 'rowHead');
  assert.equal(head.style.color, '#ffffff');
  assert.equal(head.style.fill, '#2f5597');
});

// ───────────── 열 블록 · 빅데이터 ─────────────

function blockBook(n, fill) {
  const cols = [new ColBuilder(n), new ColBuilder(n), new ColBuilder(n)];
  for (let i = 0; i < n; i++) fill(i, (j, v) => cols[j].set(i, v));
  return new Workbook({
    sheets: [{
      name: 'D',
      cells: new Map([['0,0', { raw: '지역' }], ['0,1', { raw: '금액' }], ['0,2', { raw: '날짜' }]]),
      blocks: [{ r0: 1, c0: 0, n, ver: 0, cols: [cols[0].finish(n), cols[1].finish(n), cols[2].finish(n, { numFmt: 'date' })] }],
    }],
  });
}

test('열 블록: 값 · 입력 글자 · 편집 · 수식 덮어쓰기 · 실행 취소 · 행 삽입', () => {
  const wb = blockBook(5, (i, set) => { set(0, ['서울', '부산'][i % 2]); set(1, i * 10); set(2, 46204 + i); });
  assert.equal(wb.getValue(0, 1, 0), '서울');
  assert.equal(wb.getValue(0, 3, 1), 20);
  assert.equal(wb.getRaw(0, 2, 2), '2026-07-02');
  assert.equal(wb.styleAt(0, 2, 2).numFmt, 'date');
  assert.deepEqual(wb.usedRange(0), { rows: 6, cols: 3 });
  wb.transact(() => wb.setInput(0, 2, 1, '777'));
  assert.equal(wb.getValue(0, 2, 1), 777);
  assert.equal(wb.sheets[0].cells.has('2,1'), false); // 블록 안에 저장
  wb.transact(() => wb.setInput(0, 3, 1, '=B2+1'));
  assert.equal(wb.getValue(0, 3, 1), 1); // B2 = 0
  assert.equal(wb.sheets[0].cells.has('3,1'), true); // 수식은 일반 셀
  wb.undo();
  assert.equal(wb.getValue(0, 3, 1), 20);
  wb.undo();
  assert.equal(wb.getValue(0, 2, 1), 10);
  wb.transact(() => wb.insertRows(0, 2, 2));
  assert.equal(wb.getValue(0, 1, 1), 0);
  assert.equal(wb.getValue(0, 2, 1), null);
  assert.equal(wb.getValue(0, 4, 1), 10);
  assert.equal(wb.usedRange(0).rows, 8);
  wb.undo();
  assert.equal(wb.getValue(0, 2, 1), 10);
});

test('열 블록 정렬 (기수 · 계수 정렬, 빈 칸 끝, 실행 취소)', () => {
  const vals = [5, -2, null, 3.5, 0, -2, 100, null, 7];
  const n = 1200;
  const wb = blockBook(n, (i, set) => { const v = vals[i % vals.length]; if (v !== null) set(1, v + i / 1e6); set(0, `k${(i * 7) % 13}`); });
  const before = Array.from({ length: n }, (_, i) => wb.getValue(0, i + 1, 1));
  wb.transact(() => wb.sortRange(0, 1, 0, n, 2, 1, true));
  const after = Array.from({ length: n }, (_, i) => wb.getValue(0, i + 1, 1));
  const nums = before.filter((v) => v !== null).sort((a, b) => a - b);
  assert.deepEqual(after.slice(0, nums.length), nums);
  assert.ok(after.slice(nums.length).every((v) => v === null));
  wb.transact(() => wb.sortRange(0, 1, 0, n, 2, 0, false));
  const txt = Array.from({ length: n }, (_, i) => wb.getValue(0, i + 1, 0));
  assert.deepEqual(txt, [...txt].sort((a, b) => (a < b ? 1 : a > b ? -1 : 0)));
  wb.undo();
  wb.undo();
  assert.deepEqual(Array.from({ length: n }, (_, i) => wb.getValue(0, i + 1, 1)), before);
});

test('큰 CSV 스트리밍 → 블록 (조각 경계 · 따옴표 · 날짜)', () => {
  const text = '이름,금액,날짜\r\n"홍, 길동",1200,2026-07-01\r\n"줄\n바꿈",-3.5,x\r\n일반,,2026-12-31\r\n"따옴표""",007,\r\n';
  for (const size of [1, 3, 7, 1000]) {
    const rd = new CsvBlockReader(',', 4);
    for (let i = 0; i < text.length; i += size) rd.push(text.slice(i, i + size));
    const { header, block } = rd.finish();
    assert.deepEqual(header, ['이름', '금액', '날짜']);
    assert.equal(block.n, 4);
    const v = (r, c) => blockValue(block, r + 1, c);
    assert.equal(v(0, 0), '홍, 길동');
    assert.equal(v(1, 0), '줄\n바꿈');
    assert.equal(v(3, 0), '따옴표"');
    assert.equal(v(0, 1), 1200);
    assert.equal(v(1, 1), -3.5);
    assert.equal(v(2, 1), null);
    assert.equal(v(3, 1), '007'); // 앞의 0 은 글자로
    assert.equal(v(0, 2), 46204);
    assert.equal(block.cols[2].fmt.numFmt, 'date');
    assert.equal(v(1, 2), 'x');
  }
});

test('숨긴 행 비트맵과 위치 계산', () => {
  const bits = new Uint8Array(1000);
  for (let i = 0; i < 1000; i += 3) bits[i] = 1;
  const h = { __bits: bits, start: 10, count: 334 };
  const ax = new Axis(20, { 5: 40 }, [h, { 2000: true }], 1e6);
  const naive = (i) => { let p = 0; for (let k = 0; k < i; k++) p += ax.size(k); return p; };
  for (const i of [0, 6, 10, 11, 13, 500, 1009, 1010, 1500, 2001, 5000]) assert.equal(ax.pos(i), naive(i));
  assert.equal(hid(h, 10), true);
  assert.equal(hid(h, 11), false);
  assert.equal(hidCount(h), 334);
  assert.equal(ax.indexAt(ax.pos(1501)), 1501);
});

test('피벗: 열 블록 원본 · 요약 캐시(롤업) 결과가 행 배열 원본과 같음', () => {
  const n = 250000;
  const regions = ['서울', '부산', '대구', '광주'];
  const wb = blockBook(n, (i, set) => { set(0, regions[(i * 31) % 4]); if (i % 11) set(1, (i * 7919) % 1000); set(2, 46023 + (i % 365)); });
  const rows = [['지역', '금액', '날짜']];
  for (let i = 0; i < n; i++) rows.push([wb.getValue(0, i + 1, 0), wb.getValue(0, i + 1, 1), wb.getValue(0, i + 1, 2)]);
  const def = { source: 'D', range: { r1: 0, c1: 0, r2: n, c2: 2 }, rows: ['날짜'], cols: ['지역'], values: [{ field: '금액', agg: 'sum' }, { field: '금액', agg: 'count' }, { field: '금액', agg: 'max' }], groups: { 날짜: { by: 'quarters' } }, filters: { 지역: ['서울', '대구'] } };
  const src = pivotSourceData(wb, def);
  assert.equal(src.cube.n, n);
  const a = computePivotGrid(src, def);
  const b = computePivotGrid(rows, def);
  assert.deepEqual(a, b);
  assert.equal(a[0].length > 3, true);
});

function computePivotGrid(input, def) {
  const res = resolvePivot(input, def);
  return computePivot(res, res.def).grid.map((r) => r.map((c) => c?.raw ?? ''));
}

test('피벗 값 표시 형식: 누계 · 차이 · 기준값 · 순위 · 상위 합계 비율 + xlsx 왕복', async () => {
  const { buildPivot } = await import('../src/pivot.js');
  const rows = [['지역', '월', '매출'], ['A', 1, 10], ['A', 2, 30], ['A', 3, 20], ['B', 1, 5], ['B', 2, 5], ['B', 3, 40]];
  const col = (as, extra = {}) => buildPivot(rows, { rows: ['지역', '월'], cols: [], values: [{ field: '매출', agg: 'sum', showAs: as, ...extra }], layout: 'tabular' })
    .slice(1).map((r) => r[2].raw);
  assert.deepEqual(col('runTotal', { baseField: '월' }).slice(0, 3), ['10', '40', '60']);
  assert.deepEqual(col('difference', { baseField: '월', basePos: 'prev' }).slice(0, 3), ['', '20', '-10']);
  assert.deepEqual(col('percent', { baseField: '월', baseItem: '1' }).slice(4, 7), ['1', '1', '8']);
  assert.deepEqual(col('rankDescending', { baseField: '월' }).slice(4, 7), ['2', '2', '1']);
  assert.equal(col('percentDiff', { baseField: '지역', baseItem: 'A' })[7], String(Number((-10 / 60).toPrecision(15))));
  assert.deepEqual(col('percentOfParentRow').slice(-1), ['1']);
  // xlsx: 기본 이름은 showDataAs, 확장 이름은 x14 pivotShowAs, 이전 항목은 특수 번호
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  const at = wb.transact(() => wb.addSheet('피벗'));
  const def = {
    source: 'Sheet1', range: { r1: 0, c1: 0, r2: 6, c2: 2 }, rows: ['지역', '월'], cols: [], layout: 'tabular', top: 0, left: 0,
    values: [{ field: '매출', agg: 'sum', showAs: 'difference', baseField: '월', basePos: 'prev', name: '차이' },
      { field: '매출', agg: 'sum', showAs: 'rankDescending', baseField: '월', name: '순위' },
      { field: '매출', agg: 'sum', showAs: 'percent', baseField: '지역', baseItem: 'B', name: '기준' }],
  };
  wb.transact(() => wb.setSheetProp(at, 'pivot', { ...def, area: { r1: 0, c1: 0, r2: 12, c2: 4 } }));
  const files = unzip(writeXlsx(wb));
  const pt = textOf(files['xl/pivotTables/pivotTable1.xml']);
  assert.match(pt, /showDataAs="difference" baseField="1" baseItem="1048828"/);
  assert.match(pt, /pivotShowAs="rankDescending"/);
  const back = readXlsx(zip(files)).data.sheets[1].pivot;
  assert.deepEqual(back.values.map((v) => [v.showAs, v.baseField, v.baseItem ?? v.basePos]),
    [['difference', '월', 'prev'], ['rankDescending', '월', undefined], ['percent', '지역', 'B']]);
});

test('피벗 펼치기 · 축소, 부분합 아래, 빈 줄, 레이블 반복, 빈 셀 표시, 세부 정보 + xlsx 왕복', async () => {
  const { buildPivot, pivotDetail } = await import('../src/pivot.js');
  const rows = [['지역', '구', '분기', '매출'], ['서울', '강남', 'Q1', 100], ['서울', '강북', 'Q1', 50], ['부산', '해운대', 'Q2', 30], ['서울', '강남', 'Q2', 20]];
  const base = { rows: ['지역', '구'], cols: [], values: [{ field: '매출', agg: 'sum' }] };
  const col0 = (def) => buildPivot(rows, def).map((r) => r.map((c) => c.raw).join('|'));
  // 축소: 서울의 하위 항목이 사라지고 서울 행에 합계
  assert.deepEqual(col0({ ...base, collapsed: { 지역: ['서울'] } }), ['행 레이블|합계 : 매출', '부산|30', '해운대|30', '서울|170', '총합계|200']);
  const g = buildPivot(rows, { ...base, collapsed: { 지역: ['서울'] } });
  assert.deepEqual(g[3][0].toggle, { axis: 'r', field: '지역', item: '서울', collapsed: true });
  // 부분합을 아래에 + 빈 줄 + 빈 셀 표시
  assert.deepEqual(col0({ ...base, layout: 'outline', subtotalTop: false, blankRows: true }),
    ['지역|구|합계 : 매출', '부산||', '|해운대|30', '부산 요약||30', '||', '서울||', '|강남|120', '|강북|50', '서울 요약||170', '||', '총합계||200']);
  assert.equal(col0({ rows: ['지역'], cols: ['분기'], values: [{ field: '매출', agg: 'sum' }], missingCaption: '-' })[2], "부산|'-|30|30");
  // 레이블 반복 (테이블 형식)
  assert.deepEqual(col0({ ...base, layout: 'tabular', repeatLabels: true, subtotals: false }), ['지역|구|합계 : 매출', '부산|해운대|30', '서울|강남|120', '서울|강북|50', '총합계||200']);
  // 세부 정보: 서울 / 강남 → 원본 2행
  const det = pivotDetail(rows, base, 4, 1);
  assert.deepEqual([...det.idx].map((i) => det.cube.row(i)), [['서울', '강남', 'Q1', 100], ['서울', '강남', 'Q2', 20]]);
  assert.equal(pivotDetail(rows, base, 0, 1), null);
  // xlsx: sd="0", subtotalTop="0", insertBlankRow, fillDownLabels, missingCaption
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  const at = wb.transact(() => wb.addSheet('피벗'));
  const def = { source: 'Sheet1', range: { r1: 0, c1: 0, r2: 4, c2: 3 }, ...base, layout: 'outline', top: 0, left: 0,
    collapsed: { 지역: ['부산'] }, subtotalTop: false, blankRows: true, repeatLabels: true, missingCaption: '-' };
  wb.transact(() => wb.setSheetProp(at, 'pivot', { ...def, area: { r1: 0, c1: 0, r2: 10, c2: 2 } }));
  const files = unzip(writeXlsx(wb));
  const pt = textOf(files['xl/pivotTables/pivotTable1.xml']);
  assert.match(pt, /sd="0"/);
  assert.match(pt, /missingCaption="-"/);
  assert.match(pt, /<i t="blank">/);
  const back = readXlsx(zip(files)).data.sheets[1].pivot;
  assert.deepEqual([back.collapsed, back.subtotalTop, back.blankRows, back.repeatLabels, back.missingCaption], [{ 지역: ['부산'] }, false, true, true, '-']);
});
