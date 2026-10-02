import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { computePivot, normalizeDef, pivotPageMulti, pivotSourceData, resolvePivot } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const path = 'xl/pivotTables/pivotTable1.xml';
const header = ['Region', 'Product', 'Amount'];
function fixture(options = {}) {
  const wb = new Workbook();
  wb.transact(() => {
    [header, ['Seoul', 'A', 10], ['Busan', 'A', 20], ['Daegu', 'B', 30]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.setSheetProp(0, 'pivot', {
      name: 'ReportFilter', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 2 },
      rows: ['Product'], cols: [], pages: ['Region'], values: [{ field: 'Amount', agg: 'sum' }],
      top: 0, left: 5, area: { r1: 0, c1: 5, r2: 5, c2: 6 }, ...options,
    });
  });
  return wb;
}
function result(wb) {
  const def = wb.sheets[0].pivot, resolved = resolvePivot(pivotSourceData(wb, def), def);
  return computePivot(resolved, resolved.def).grid.map((r) => r.map((c) => c.raw));
}
function xmlFields(bytes) {
  const root = parseXml(textOf(unzip(bytes)[path]));
  return { root, field: kids(child(root, 'pivotFields'), 'pivotField')[0], page: kids(child(root, 'pageFields'), 'pageField')[0] };
}

for (const [multi, selected] of [[true, ['Seoul']], [true, ['Seoul', 'Busan']], [true, null], [false, ['Seoul']], [false, null]]) {
  test(`보고서 필터 다중 모드 ${multi}, 선택 ${selected?.length ?? '전체'}: 표준 XLSX·집계 왕복`, () => {
    const wb = fixture({ pageMulti: { Region: multi }, ...(selected ? { filters: { Region: selected } } : {}) });
    const before = wb.serialize(), grid = result(wb), bytes = writeXlsx(wb), { field, page } = xmlFields(bytes);
    assert.equal(field.attrs.multipleItemSelectionAllowed, multi ? '1' : '0');
    assert.equal(page.attrs.item !== undefined, !multi && selected?.length === 1, '다중 모드에는 pageField@item을 쓰지 않음');
    if (selected) assert.equal(kids(child(field, 'items'), 'item').filter((it) => it.attrs.h === '1').length, 3 - selected.length);
    const back = new Workbook(readXlsx(bytes).data), def = back.sheets[0].pivot;
    assert.equal(def.pageMulti.Region, multi);
    assert.equal(pivotPageMulti(def, 'Region'), multi, '항목 하나 또는 전체 선택이어도 다중 모드를 보존');
    assert.deepEqual(def.filters?.Region ? new Set(def.filters.Region) : undefined, selected ? new Set(selected) : undefined);
    assert.deepEqual(result(back), grid, '모드 저장으로 표시 값이 바뀌지 않음');
    assert.deepEqual(wb.serialize(), before, '내보내기는 원본 정의와 셀을 변경하지 않음');
  });
}

test('외부 XLSX bool 0/1/false/true 및 생략을 선택 수와 분리해서 읽음', () => {
  for (const [value, multi] of [[null, false], ['0', false], ['false', false], ['1', true], ['true', true]]) {
    const files = unzip(writeXlsx(fixture({ pageMulti: { Region: false }, filters: { Region: ['Seoul'] } })));
    files[path] = textOf(files[path]).replace(/ multipleItemSelectionAllowed="[^"]*"/, value === null ? '' : ` multipleItemSelectionAllowed="${value}"`);
    const back = new Workbook(readXlsx(zip(files)).data);
    assert.equal(back.sheets[0].pivot.pageMulti.Region, multi, String(value));
    assert.deepEqual(back.sheets[0].pivot.filters.Region, ['Seoul']);
    assert.equal(xmlFields(writeXlsx(back)).page.attrs.item !== undefined, !multi);
  }
});

test('다중 필터는 외부 파일에 잘못 남은 pageField@item보다 항목 숨김을 사용', () => {
  const files = unzip(writeXlsx(fixture({ pageMulti: { Region: true }, filters: { Region: ['Seoul', 'Busan'] } })));
  files[path] = textOf(files[path]).replace('<pageField fld="0"', '<pageField item="2" fld="0"');
  const back = new Workbook(readXlsx(zip(files)).data);
  assert.deepEqual(new Set(back.sheets[0].pivot.filters.Region), new Set(['Seoul', 'Busan']));
  assert.equal(xmlFields(writeXlsx(back)).page.attrs.item, undefined);
});

test('옛 JSON만 선택 개수 fallback, 명시 모드·정규화·실행 취소를 보존', () => {
  const wb = fixture({ filters: { Region: ['Seoul', 'Busan'] } }), original = structuredClone(wb.sheets[0].pivot);
  assert.equal(pivotPageMulti(original, 'Region'), true);
  assert.equal(pivotPageMulti({ filters: { Region: ['Seoul'] } }, 'Region'), false);
  assert.equal(pivotPageMulti({ ...original, pageMulti: { Region: false } }, 'Region'), false);
  const normalized = normalizeDef({ ...original, pageMulti: { region: true, Amount: false, Product: 'true', Missing: true } }, header);
  assert.deepEqual(normalized.pageMulti, { Region: true, Amount: false });
  assert.deepEqual(normalizeDef(normalized, header).pageMulti, normalized.pageMulti);
  assert.equal(xmlFields(writeXlsx(wb)).field.attrs.multipleItemSelectionAllowed, '1', '기존 다중 선택 내보내기도 다중 모드');
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...original, filters: { Region: ['Seoul'] }, pageMulti: { Region: true } }));
  const changed = structuredClone(wb.sheets[0].pivot);
  wb.undo(); assert.deepEqual(wb.sheets[0].pivot, original);
  wb.redo(); assert.deepEqual(wb.sheets[0].pivot, changed);
  assert.equal(new Workbook(wb.serialize()).sheets[0].pivot.pageMulti.Region, true);
});

test('보고서 필터가 여러 개여도 필드 순서와 각 체크 목록 모드를 독립 보존', () => {
  const wb = fixture({ rows: [], pages: ['Product', 'Region'], pageMulti: { Product: false, Region: true }, filters: { Product: ['A'], Region: ['Seoul'] } });
  const before = result(wb), back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(back.sheets[0].pivot.pages, ['Product', 'Region']);
  assert.deepEqual(back.sheets[0].pivot.pageMulti, { Region: true, Product: false });
  assert.deepEqual(back.sheets[0].pivot.filters, { Region: ['Seoul'], Product: ['A'] });
  assert.deepEqual(result(back), before);
});
