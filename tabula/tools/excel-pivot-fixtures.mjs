// 합성 피벗만 생성합니다. 실제 Excel 열기·RefreshTable·SaveAs는 excel-interop.ps1로 별도 실행합니다.
// node tools/excel-pivot-fixtures.mjs D:/Codex/Temp/wixel-pivot-excel [--verify 저장폴더]
import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { pivotSourceData, resolvePivot, computePivot, pivotLookup, pivotErrorDisplay } from '../src/pivot.js';

const out = process.argv[2];
if (!out) throw new Error('사용법: node tools/excel-pivot-fixtures.mjs D:/Codex/Temp/wixel-pivot-excel [--verify 저장폴더]');
const at = process.argv.indexOf('--verify'), verifyRoot = at >= 0 ? process.argv[at + 1] : null;
if (at >= 0 && !verifyRoot) throw new Error('--verify 뒤에 Excel 재저장 폴더가 필요합니다.');
const baseRows = [
  ['지역', '제품', '분기', '경로', '매출', '클릭'],
  ['서울', 'A', '1분기', '검색', 120, 6], ['서울', 'A', '1분기', '영상', 80, 4],
  ['서울', 'B', '2분기', '검색', 60, 3], ['부산', 'A', '1분기', '검색', 90, 3],
  ['부산', 'B', '2분기', '영상', 150, 5], ['부산', 'B', '2분기', '검색', 100, 4],
];
const sales = { field: '매출', name: '총매출', agg: 'sum' }, clicks = { field: '클릭', name: '총클릭', agg: 'sum' };
const q = (value, ...pairs) => ({ dataField: '총매출', pairs, value });
const baseQueries = [q(600), q(260, '지역', '서울'), q(340, '지역', '부산')];
const nestedQueries = [...baseQueries, q(200, '지역', '서울', '제품', 'A'), q(250, '지역', '부산', '제품', 'B'), { dataField: '총클릭', pairs: [], value: 25 }];
const errorRows = [['지역', '비용', '클릭'], ['오류행', 10, 0], ['정상행', 30, 3]];
const ratio = { rows: ['지역'], values: [{ field: 'CPC', name: '평균클릭단가', agg: 'sum' }], calcFields: [{ name: 'CPC', formula: '비용/클릭' }] };
const dateSerial = (s) => Date.parse(s + 'T00:00:00Z') / 86400000 + 25569;
const dateRows = [['날짜', '지역', '매출'], ['2024-01-02', '서울', 10], ['2024-01-20', '부산', 90], ['2024-02-03', '서울', 20], ['2024-02-20', '서울', 30], ['2024-03-01', '부산', 70]].map((r, i) => i ? [dateSerial(r[0]), ...r.slice(1)] : r);
const cases = [
  { label: '테이블 형식', rows: baseRows, def: { layout: 'tabular', rows: ['지역'] }, checks: baseQueries },
  { label: '압축 형식·다중 행·다중 값', rows: baseRows, def: { layout: 'compact', rows: ['지역', '제품'], values: [sales, clicks] }, checks: nestedQueries },
  { label: '개요 형식·반복 레이블', rows: baseRows, def: { layout: 'outline', rows: ['지역', '제품'], values: [sales, clicks], repeatLabels: true, subtotalTop: false }, checks: nestedQueries },
  { label: '다중 행·다중 열', rows: baseRows, def: { layout: 'tabular', rows: ['지역', '제품'], cols: ['분기', '경로'], values: [sales, clicks] }, checks: [q(600), q(120, '지역', '서울', '제품', 'A', '분기', '1분기', '경로', '검색'), q(150, '지역', '부산', '제품', 'B', '분기', '2분기', '경로', '영상'), { dataField: '총클릭', pairs: ['지역', '서울'], value: 13 }] },
  { label: '계산 필드·집계 후 나눗셈', rows: baseRows, def: { layout: 'tabular', rows: ['지역'], values: [{ field: 'CPC', name: '평균클릭단가', agg: 'sum' }], calcFields: [{ name: 'CPC', formula: '매출/클릭' }] }, checks: [{ dataField: '평균클릭단가', pairs: [], value: 24 }, { dataField: '평균클릭단가', pairs: ['지역', '서울'], value: 20 }, { dataField: '평균클릭단가', pairs: ['지역', '부산'], value: 340 / 12 }] },
  // 실제 Excel COM·PDF에서 ErrorString='-'는 숫자 0이다. 문구 자체는 '-'로 보존한다.
  ...[false, true].map((show) => ({ label: '오류 표시 ' + show, rows: errorRows, def: { layout: 'tabular', ...ratio, errorShow: show, errorCaption: '-' }, checks: [{ dataField: '평균클릭단가', pairs: ['지역', '오류행'], text: show ? '0' : '#DIV/0!' }, { dataField: '평균클릭단가', pairs: ['지역', '정상행'], value: 10 }] })),
  { label: '월 그룹·보고서 필터·슬라이서', rows: dateRows, dateColumn: 0, slicer: '지역', def: { layout: 'tabular', rows: ['날짜'], pages: ['지역'], filters: { 지역: ['서울'] }, groups: { 날짜: { by: 'months', start: dateSerial('2024-01-01'), end: dateSerial('2024-12-31') } } }, checks: [q(60), q(10, '날짜', 1), q(50, '날짜', 2)] },
];

function renderPivot(wb, si, def) {
  const res = resolvePivot(pivotSourceData(wb, def), def), { grid } = computePivot(res, res.def);
  let width = 0;
  wb.transact(() => grid.forEach((row, r) => {
    width = Math.max(width, row.length);
    row.forEach((cd, c) => { if (cd) wb.setCellData(si, def.top + r, def.left + c, { raw: cd.raw, style: cd.style }); });
  }));
  def.area = { r1: def.top, c1: def.left, r2: def.top + grid.length - 1, c2: def.left + width - 1 };
  wb.transact(() => wb.setSheetProp(si, 'pivot', def));
}
function readChecks(wb, fixture) {
  const si = wb.sheets.findIndex((s) => s.name === fixture.pivotSheet), sheet = wb.sheets[si];
  assert.ok(sheet, '피벗 시트');
  const def = sheet.pivot;
  assert.ok(def, '피벗 정의가 유지되어야 합니다.');
  assert.equal(def.name, fixture.pivotName, '피벗 이름');
  assert.equal(def.layout, fixture.expected.layout, '보고서 레이아웃');
  for (const key of ['rows', 'cols', 'pages']) assert.deepEqual(def[key] ?? [], fixture.expected[key] ?? [], key + ' 필드');
  assert.deepEqual(def.values.map((v) => [v.field, v.name, v.agg]), fixture.expected.values.map((v) => [v.field, v.name, v.agg]), '값 필드·이름·집계');
  if (fixture.expected.errorShow !== undefined) {
    assert.equal(pivotErrorDisplay(def), fixture.expected.errorShow, '오류 표시 사용 여부');
    assert.equal(def.errorCaption, fixture.expected.errorCaption, '오류 표시 문자열');
  }
  for (const [field, group] of Object.entries(fixture.expected.groups ?? {})) assert.equal(def.groups?.[field]?.by, group.by, '날짜 그룹 단위');
  for (const [field, items] of Object.entries(fixture.expected.filters ?? {})) assert.deepEqual(def.filters?.[field], items, '필터 선택');
  assert.equal((sheet.slicers ?? []).length, fixture.slicers ?? 0, '슬라이서 개수');
  const input = pivotSourceData(wb, def), res = resolvePivot(input, def);
  for (const check of fixture.pivotChecks) {
    const pairs = []; for (let i = 0; i < check.pairs.length; i += 2) pairs.push(check.pairs.slice(i, i + 2));
    const value = pivotLookup(input, def, check.dataField, pairs, res);
    if ('text' in check) assert.equal(value?.code ?? String(value), check.text, '재계산 표시: ' + JSON.stringify(check));
    else { assert.equal(typeof value, 'number'); assert.ok(Math.abs(value - check.value) < 1e-9, '재계산 값: ' + JSON.stringify({ check, value })); }
  }
  const source = wb.sheets.findIndex((s) => s.name === '원본');
  for (let r = 0; r < fixture.sourceRows.length; r++) for (let c = 0; c < fixture.sourceRows[r].length; c++) assert.equal(wb.getValue(source, r, c), fixture.sourceRows[r][c], `원본 ${r},${c}`);
  return { layout: def.layout, values: def.values.length, queries: fixture.pivotChecks.length, slicers: (sheet.slicers ?? []).length };
}

let failures = 0;
if (verifyRoot) {
  const manifest = JSON.parse(readFileSync(join(out, 'synthetic-fixtures.json'), 'utf8')), results = [];
  assert.equal(manifest.format, 'wixel-excel-interop'); assert.equal(manifest.version, 1);
  for (const fixture of manifest.fixtures) {
    assert.match(fixture.file, /^[a-zA-Z0-9_-]+\.xlsx$/);
    try { results.push({ file: fixture.file, label: fixture.label, ok: true, ...readChecks(new Workbook(readXlsx(readFileSync(join(verifyRoot, fixture.file))).data), fixture) }); }
    catch (e) { failures++; results.push({ file: fixture.file, label: fixture.label, ok: false, error: e.message }); }
  }
  writeFileSync(join(out, 'pivot-reread-results.json'), JSON.stringify({ total: results.length, failures, results }, null, 2));
  results.forEach((r) => console.log(JSON.stringify(r)));
  console.log(JSON.stringify({ mode: 'verify', total: results.length, failures }));
} else {
  mkdirSync(out, { recursive: true }); const fixtures = [], results = [];
  for (const [i, entry] of cases.entries()) {
    const wb = new Workbook(), name = `검증피벗${i + 1}`, file = `pivot-${String(i + 1).padStart(2, '0')}.xlsx`;
    wb.sheets[0].name = '원본';
    wb.transact(() => {
      entry.rows.forEach((row, r) => row.forEach((v, c) => { wb.setInput(0, r, c, String(v)); if (r && c === entry.dateColumn) wb.setStyle(0, r, c, { numFmt: 'date', code: 'yyyy-mm-dd' }); }));
      wb.addSheet('피벗');
    });
    const def = { name, source: '원본', range: { r1: 0, c1: 0, r2: entry.rows.length - 1, c2: entry.rows[0].length - 1 }, top: 0, left: 0, cols: [], pages: [], values: [sales], subtotals: true, grandRows: true, grandCols: true, style: 'PivotStyleMedium2', ...structuredClone(entry.def) };
    renderPivot(wb, 1, def);
    if (entry.slicer) wb.transact(() => wb.setSheetProp(1, 'slicers', [{ id: '검증슬라이서', caption: entry.slicer, source: { kind: 'pivot', field: entry.slicer, pivots: [{ sheet: '피벗', name }] }, x: 400, y: 20, w: 180, h: 200 }]));
    const fixture = { file, label: entry.label, chartTypes: [], pivots: 1, pivotSheet: '피벗', pivotName: name, refreshPivots: true, pivotChecks: entry.checks.map((c) => ({ name, ...c })), slicers: entry.slicer ? 1 : 0, expected: def, sourceRows: entry.rows, cells: [{ address: 'A1', value: entry.rows[0][0] }] };
    const bytes = writeXlsx(wb); writeFileSync(join(out, file), bytes); fixtures.push(fixture);
    try { results.push({ file, label: entry.label, ok: true, ...readChecks(new Workbook(readXlsx(bytes).data), fixture) }); }
    catch (e) { failures++; results.push({ file, label: entry.label, ok: false, error: e.message }); }
  }
  writeFileSync(join(out, 'synthetic-fixtures.json'), JSON.stringify({ format: 'wixel-excel-interop', version: 1, purpose: '합성 피벗 Excel 열기·새로 고침·재저장·재읽기 검증. 실제 업무 파일이나 Excel 모든 버전의 호환성을 보증하지 않습니다.', refreshSource: 'https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.refreshtable', lookupSource: 'https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.getpivotdata', fixtures }, null, 2));
  writeFileSync(join(out, 'pivot-baseline-results.json'), JSON.stringify({ total: results.length, failures, results }, null, 2));
  results.forEach((r) => console.log(JSON.stringify(r)));
  console.log(JSON.stringify({ mode: 'generate', path: resolve(out), total: fixtures.length, failures }));
}
if (failures) process.exitCode = 1;
