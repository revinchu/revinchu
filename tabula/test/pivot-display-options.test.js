import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { computePivot, normalizeDef, pivotDisplayOptions, pivotSourceData, resolvePivot } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const header = ['지역', '품목', '매출', '비용'];
const path = 'xl/pivotTables/pivotTable1.xml';
const displayUri = '{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}';
function fixture(options = {}) {
  const wb = new Workbook();
  wb.transact(() => {
    [header, ['서울', 'A', 10, 1], ['서울', 'B', 20, 2], ['부산', 'A', 30, 3]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.setSheetProp(0, 'pivot', {
      name: '표시옵션', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 3 },
      rows: ['지역', '품목'], cols: [], values: [{ field: '매출', name: '총매출' }, { field: '비용', name: '총비용' }],
      layout: 'compact', top: 0, left: 6, area: { r1: 0, c1: 6, r2: 9, c2: 9 }, ...options,
    });
  });
  return wb;
}
function rendered(wb) {
  const def = wb.sheets[0].pivot;
  const res = resolvePivot(pivotSourceData(wb, def), def);
  return computePivot(res, res.def);
}

// Excel COM로 compact/tabular × InGridDropZones × ShowValuesRow 8종을 생성해
// TableRange1.Text와 저장 XML을 확인했다. classic=true는 ShowValuesRow=false여도
// 값 행을 표시하지만, 두 속성 및 compact/tabular 레이아웃은 독립적으로 보존한다.
for (const layout of ['compact', 'tabular']) for (const classic of [false, true]) for (const showValuesRow of [false, true]) {
  test(`피벗 표시 ${layout}, classic=${classic}, showValuesRow=${showValuesRow}: 헤더·집계·표준 XLSX 왕복`, () => {
    const wb = fixture({ layout, classic, showValuesRow }), before = rendered(wb);
    const expectedHeaders = classic || showValuesRow ? 2 : 1;
    assert.equal(before.meta.headerRows, expectedHeaders);
    const caps = layout === 'compact' ? ['행 레이블', '총매출', '총비용'] : ['지역', '품목', '총매출', '총비용'];
    assert.deepEqual(before.grid[expectedHeaders - 1].map((c) => c.raw), caps);
    if (expectedHeaders === 2) assert.deepEqual(before.grid[0].map((c) => c.raw), layout === 'compact' ? ['', '값', ''] : ['', '', '값', '']);
    assert.deepEqual(before.grid.at(-1).slice(-2).map((c) => c.raw), ['60', '6'], '표시 옵션은 집계를 바꾸지 않음');
    const bytes = writeXlsx(wb), xml = parseXml(textOf(unzip(bytes)[path]));
    assert.equal(xml.attrs.gridDropZones === '1', classic, 'tabular라는 이유만으로 classic을 켜면 안 됨');
    const ext = kids(child(xml, 'extLst'), 'ext').find((e) => e.attrs.uri === displayUri);
    assert.equal(child(ext, 'pivotTableDefinition').attrs.hideValuesRow, showValuesRow ? '0' : '1');
    assert.equal(Number(child(xml, 'location').attrs.firstDataRow), expectedHeaders);
    const back = new Workbook(readXlsx(bytes).data), def = back.sheets[0].pivot;
    assert.equal(def.layout, layout); assert.equal(def.classic, classic); assert.equal(def.showValuesRow, showValuesRow);
    assert.deepEqual(rendered(back).grid.map((r) => r.map((c) => c.raw)), before.grid.map((r) => r.map((c) => c.raw)));
  });
}

test('구 JSON의 valuesHeadRow 별칭, 명시적 false, 정규화 재실행을 구분', () => {
  assert.deepEqual(pivotDisplayOptions({}), { classic: false, showValuesRow: false, valuesHeadRow: false });
  assert.equal(pivotDisplayOptions({ valuesHeadRow: true }).showValuesRow, true);
  assert.equal(pivotDisplayOptions({ valuesHeadRow: true, showValuesRow: false }).valuesHeadRow, false);
  const d = { ...fixture().sheets[0].pivot, classic: true, showValuesRow: false };
  const normalized = normalizeDef(d, header);
  assert.deepEqual(pivotDisplayOptions(normalized), { classic: true, showValuesRow: false, valuesHeadRow: true });
  assert.equal(normalizeDef({ ...normalized, classic: false }, header).valuesHeadRow, false);
  assert.equal(d.valuesHeadRow, undefined, '원본 정의 변경 금지');
});

test('외부 XLSX의 값 행 기본값·x14 bool·클래식 bool을 정확히 읽음', () => {
  for (const [hide, expected] of [[null, true], ['0', true], ['false', true], ['1', false], ['true', false]]) {
    const files = unzip(writeXlsx(fixture({ classic: true, showValuesRow: false })));
    files[path] = textOf(files[path]).replace('gridDropZones="1"', 'gridDropZones="true"').replace(' hideValuesRow="1"', hide === null ? '' : ` hideValuesRow="${hide}"`);
    const wb = new Workbook(readXlsx(zip(files)).data);
    assert.equal(wb.sheets[0].pivot.showValuesRow, expected);
    assert.equal(wb.sheets[0].pivot.classic, true);
    assert.equal(rendered(wb).meta.headerRows, 2, 'classic 상태의 헤더 위치로 체크 상태를 추론하지 않음');
  }
  const files = unzip(writeXlsx(fixture({ showValuesRow: false })));
  files[path] = textOf(files[path]).replace(/<extLst>[\s\S]*?<\/extLst>/, '');
  const wb = new Workbook(readXlsx(zip(files)).data);
  assert.equal(wb.sheets[0].pivot.showValuesRow, true, '확장 자체가 없어도 표준 기본값은 표시');
});

test('값 하나·값을 행에 배치·열 필드가 있는 경우에는 별도 값 윗행을 추가하지 않음', () => {
  for (const options of [{ values: [{ field: '매출' }] }, { valuesOnRows: true }, { rows: ['지역'], cols: ['품목'] }]) {
    const off = rendered(fixture({ ...options, showValuesRow: false }));
    const on = rendered(fixture({ ...options, showValuesRow: true }));
    assert.deepEqual(on.grid, off.grid);
    assert.equal(on.meta.headerRows, off.meta.headerRows);
  }
});

test('한 extLst 안에 값 행 표준 확장과 계산 항목 확장을 함께 보존', () => {
  const calcItems = { 품목: [{ name: '묶음', formula: 'A+B' }] };
  const bytes = writeXlsx(fixture({ showValuesRow: true, calcItems }));
  const xml = parseXml(textOf(unzip(bytes)[path]));
  assert.equal(kids(xml, 'extLst').length, 1);
  assert.equal(kids(child(xml, 'extLst'), 'ext').length, 2);
  const wb = new Workbook(readXlsx(bytes).data);
  assert.equal(wb.sheets[0].pivot.showValuesRow, true);
  assert.deepEqual(wb.sheets[0].pivot.calcItems, calcItems);
});

test('표시 설정 변경·해제는 실행 취소/다시 실행에서 원본 정의와 표시 행을 복원', () => {
  const wb = fixture({ showValuesRow: false, classic: false });
  const before = JSON.stringify(wb.sheets[0].pivot);
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...wb.sheets[0].pivot, classic: true, layout: 'tabular' }));
  assert.equal(rendered(wb).meta.headerRows, 2);
  wb.undo();
  assert.equal(JSON.stringify(wb.sheets[0].pivot), before);
  assert.equal(rendered(wb).meta.headerRows, 1);
  wb.redo();
  assert.equal(rendered(wb).meta.headerRows, 2);
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...wb.sheets[0].pivot, classic: false }));
  assert.equal(rendered(wb).meta.headerRows, 1);
  assert.equal(wb.sheets[0].pivot.layout, 'tabular');
});
