import { test } from 'node:test';
import { deepEqual, equal, ok } from 'node:assert/strict';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';
import { pivotSortPatch } from '../src/pivot-sort-state.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxBlobAsync } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { ERR } from '../src/fxcore.js';

const values = [{ field: '매출', agg: 'sum' }, { field: '수량', agg: 'sum' }];
const source = [
  ['지역', '상품', '연도', '분기', '매출', '수량'],
  ['서울', 'A', 2025, '1분기', 3, 2],
  ['서울', 'B', 2025, '2분기', 20, 3],
  ['부산', 'A', 2026, '1분기', 50, 7],
  ['부산', 'B', 2026, '2분기', 5, 5],
  ['총합계', 'A', 2025, '1분기', 1, 1],
];
const result = (rows, def) => {
  const res = resolvePivot(rows, def);
  return computePivot(res, res.def);
};
const childKeys = node => node.children.map(child => child.key);
const plain = raw => String(raw).replace(/^'/, '');

test('피벗 정렬은 총합계 이름의 실제 항목도 정렬하고 구조상 총합계만 마지막에 둔다', () => {
  const rows = [['지역', '매출'], ['총합계', 5], ['서울', 20], ['부산', -3]];
  for (const layout of ['compact', 'outline', 'tabular']) {
    for (const dir of ['asc', 'desc']) for (const grandCaption of ['총합계', '전체 보고서', '']) {
      const out = result(rows, { rows: ['지역'], values: [values[0]], layout, grandCaption, sort: { 지역: { dir, by: 0 } } });
      deepEqual(childKeys(out.meta.rowTree), dir === 'asc' ? ['부산', '총합계', '서울'] : ['서울', '총합계', '부산']);
      deepEqual(out.meta.rowItems.map(item => item.kind), ['item', 'item', 'item', 'grand']);
      equal(out.grid.at(-1)[0].role, 'grandLabel');
      equal(plain(out.grid.at(-1)[0].raw), grandCaption);
      equal(out.grid.at(-1)[1].raw, '22');
    }
  }
});

test('여러 수준의 행·열 값 정렬은 부모별 항목·부분합·총합계 연결을 유지한다', () => {
  for (const layout of ['compact', 'outline', 'tabular']) for (const subtotalTop of [true, false]) {
    const out = result(source, {
      rows: ['지역', '상품'], cols: ['연도', '분기'], values, layout, subtotalTop,
      sort: Object.fromEntries(['지역', '상품', '연도', '분기'].map(field => [field, { dir: 'desc', by: 0 }])),
    });
    deepEqual(childKeys(out.meta.rowTree), ['부산', '서울', '총합계']);
    deepEqual(out.meta.rowTree.children.map(childKeys), [['A', 'B'], ['B', 'A'], ['A']]);
    deepEqual(childKeys(out.meta.colTree), [2026, 2025]);
    deepEqual(out.meta.colTree.children.map(childKeys), [['1분기', '2분기'], ['2분기', '1분기']]);
    equal(out.meta.rowItems.at(-1).kind, 'grand');
    deepEqual(out.meta.colLeaves.slice(-2).map(leaf => [leaf.kind, leaf.vi]), [['grand', 0], ['grand', 1]]);
    deepEqual(out.grid.at(-1).slice(-2).map(cell => cell.raw), ['79', '18']);
    // Each column subtotal follows its own children, even after descending sorts.
    deepEqual(out.meta.colLeaves.filter(leaf => leaf.vi === 0).map(leaf => [leaf.kind, leaf.node?.key]), [
      ['item', '1분기'], ['item', '2분기'], ['sub', 2026],
      ['item', '2분기'], ['item', '1분기'], ['sub', 2025], ['grand', null],
    ]);
    if (!subtotalTop) {
      deepEqual(out.meta.rowItems.filter(item => item.kind === 'sub').map(item => item.node.key), ['부산', '서울', '총합계']);
    }
  }
});

test('선택한 반대 축 항목의 값 정렬은 전체 합계 정렬과 구별되고 총합계를 이동하지 않는다', () => {
  const rows = [['지역', '상품', '매출'], ['서울', 'A', 1], ['서울', 'B', 90], ['부산', 'A', 20], ['부산', 'B', 2]];
  const def = { rows: ['지역'], cols: ['상품'], values: [values[0]], layout: 'tabular' };
  const scopedRows = result(rows, { ...def, sort: { 지역: { dir: 'desc', by: 0, at: [['상품', 'A']] } } });
  const totalRows = result(rows, { ...def, sort: { 지역: { dir: 'desc', by: 0 } } });
  deepEqual(childKeys(scopedRows.meta.rowTree), ['부산', '서울']);
  deepEqual(childKeys(totalRows.meta.rowTree), ['서울', '부산']);
  const scopedCols = result(rows, { ...def, sort: { 상품: { dir: 'desc', by: 0, at: [['지역', '부산']] } } });
  const totalCols = result(rows, { ...def, sort: { 상품: { dir: 'desc', by: 0 } } });
  deepEqual(childKeys(scopedCols.meta.colTree), ['A', 'B']);
  deepEqual(childKeys(totalCols.meta.colTree), ['B', 'A']);
  for (const out of [scopedRows, totalRows, scopedCols, totalCols]) {
    equal(out.meta.rowItems.at(-1).kind, 'grand');
    equal(out.meta.colLeaves.at(-1).kind, 'grand');
    equal(out.grid.at(-1).at(-1).raw, '113');
  }
});

test('행에 배치한 여러 값과 축소 항목도 정렬 뒤 총합계 묶음과 부분합을 보존한다', () => {
  for (const layout of ['compact', 'outline', 'tabular']) {
    const out = result(source, {
      rows: ['지역', '상품'], cols: ['연도'], values, valuesOnRows: true, layout, subtotalTop: false,
      collapsed: { 지역: ['부산'] }, sort: { 지역: { dir: 'desc', by: 0 }, 연도: { dir: 'desc', by: 1 } },
    });
    deepEqual(childKeys(out.meta.rowTree), ['부산', '서울', '총합계']);
    deepEqual(out.meta.rowItems.slice(-2).map(item => [item.kind, item.vi]), [['grand', 0], ['grand', 1]]);
    equal(out.meta.colLeaves.at(-1).kind, 'grand');
    deepEqual(out.grid.slice(-2).map(row => row.at(-1).raw), ['79', '18']);
    const collapsed = out.meta.rowItems.filter(item => item.node?.key === '부산' && item.vi !== undefined);
    deepEqual(collapsed.map(item => item.vi), [0, 1]);
    ok(collapsed.every(item => item.coll));
  }
});

// Native Excel 16.0 build 14334 AutoSort, synthetic pivot, 2026-10-05:
// both ascending and descending keep #DIV/0! and #N/A after numeric/blank values.
test('피벗 값 정렬은 오름·내림차순 모두 오류 항목을 정상값 뒤에 둔다', () => {
  const rows = [['지역', '매출'], ['오류1', ERR.DIV0], ['양수', 10], ['빈값', null], ['음수', -2], ['영', 0], ['오류2', ERR.NA]];
  for (const axis of ['rows', 'cols']) for (const dir of ['asc', 'desc']) for (const errorShow of [false, true]) {
    const out = result(rows, {
      rows: [], cols: [], [axis]: ['지역'], values: [values[0]], layout: 'tabular', errorShow, errorCaption: '999',
      sort: { 지역: { dir, by: 0 } }, tieOrder: { 지역: ['영', '빈값', '오류2', '오류1'] },
    });
    const keys = childKeys(axis === 'rows' ? out.meta.rowTree : out.meta.colTree);
    deepEqual(keys.slice(0, 4), dir === 'asc' ? ['음수', '영', '빈값', '양수'] : ['양수', '영', '빈값', '음수']);
    deepEqual(keys.slice(4), ['오류2', '오류1']);
    if (axis === 'rows') equal(out.meta.rowItems.at(-1).kind, 'grand');
    else equal(out.meta.colLeaves.at(-1).kind, 'grand');
  }
});

test('다단계 선택 열 값 정렬은 XLSX 왕복과 실행 취소 뒤에도 원본·부분합·총합계를 보존한다', () => {
  const wb = new Workbook();
  wb.sheets[0].name = '원본';
  source.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
  const si = wb.addSheet('피벗');
  wb.sheets[si].pivot = {
    name: '정렬검증', source: '원본', range: { r1: 0, c1: 0, r2: source.length - 1, c2: source[0].length - 1 },
    top: 2, left: 1, rows: ['지역', '상품'], cols: ['연도', '분기'], values,
    layout: 'tabular', subtotalTop: false, grandCaption: '보고서 합계',
    sort: { 지역: { dir: 'asc' }, 연도: { dir: 'asc' } },
  };
  const imported = new Workbook(readXlsx(writeXlsx(wb)).data), before = imported.sheets[si].pivot;
  const snapshot = source.map((row, r) => row.map((_, c) => imported.getValue(0, r, c)));
  const query = book => result(pivotSourceData(book, book.sheets[si].pivot), book.sheets[si].pivot);
  const oldGrid = query(imported).grid.map(row => row.map(cell => cell.raw));
  let next = { ...before, ...pivotSortPatch(before, { field: '지역', sort: { dir: 'desc', by: 1, at: [['연도', '2025'], ['분기', '2분기']] } }) };
  next = { ...next, ...pivotSortPatch(next, { field: '연도', sort: { dir: 'desc', by: 0, at: [['지역', '부산']] } }) };
  imported.transact(() => imported.setSheetProp(si, 'pivot', next));
  const changed = query(imported), expected = changed.grid.map(row => row.map(cell => cell.raw));
  deepEqual(childKeys(changed.meta.rowTree), ['서울', '부산', '총합계']);
  deepEqual(childKeys(changed.meta.colTree), [2026, 2025]);
  equal(changed.meta.rowItems.at(-1).kind, 'grand');
  deepEqual(changed.grid.at(-1).slice(-2).map(cell => cell.raw), ['79', '18']);
  const reopened = new Workbook(readXlsx(writeXlsx(imported)).data);
  deepEqual(reopened.sheets[si].pivot.sort.지역, next.sort.지역);
  deepEqual(reopened.sheets[si].pivot.sort.연도, next.sort.연도);
  deepEqual(query(reopened).grid.map(row => row.map(cell => cell.raw)), expected);
  deepEqual(source.map((row, r) => row.map((_, c) => reopened.getValue(0, r, c))), snapshot);
  imported.undo(); deepEqual(query(imported).grid.map(row => row.map(cell => cell.raw)), oldGrid);
  imported.redo(); deepEqual(query(imported).grid.map(row => row.map(cell => cell.raw)), expected);
});

for (const mode of ['sync', 'blob']) test('선택 항목과 두 번째 값 정렬을 표준 XLSX 참조로 기록: ' + mode, async () => {
  const rows = [['Row', 'Col', 'First', 'Second'], ['X', 'A', 1, 20], ['X', 'C', 3, 4], ['Y', 'B', 5, 100], ['Y', 'A', 40, 1]];
  const wb = new Workbook(); wb.sheets[0].name = 'Source';
  rows.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
  const si = wb.addSheet('Report');
  const def = {
    source: 'Source', range: { r1: 0, c1: 0, r2: 4, c2: 3 }, top: 0, left: 0,
    rows: ['Row'], cols: ['Col'], values: [{ field: 'First', agg: 'sum' }, { field: 'Second', agg: 'sum', name: 'Chosen value' }],
    sort: { Row: { dir: 'desc', by: 'Chosen value', at: [['col', 'a']], importedOption: { keep: true } }, Col: { dir: 'desc' } },
  };
  wb.sheets[si].pivot = def;
  const before = JSON.stringify(def);
  const bytes = mode === 'sync' ? writeXlsx(wb) : new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
  const files = unzip(bytes), path = 'xl/pivotTables/pivotTable1.xml';
  const xml = textOf(files[path]), fields = kids(child(parseXml(xml), 'pivotFields'), 'pivotField');
  deepEqual(descendants(child(fields[0], 'autoSortScope'), 'reference').map(ref => [ref.attrs.field, Number(child(ref, 'x').attrs.v)]), [['4294967294', 1], ['1', 2]]);
  // C,B,A are pivot item positions; their shared cache indexes stay 2,1,0.
  deepEqual(kids(child(fields[1], 'items'), 'item').filter(item => item.attrs.x !== undefined).map(item => Number(item.attrs.x)), [2, 1, 0]);
  equal(JSON.stringify(def), before);
  files[path] = xml.replace(/<extLst>[\s\S]*?<\/extLst>/g, '');
  const reopened = new Workbook(readXlsx(zip(files)).data), restored = reopened.sheets[si].pivot;
  deepEqual(restored.sort.Row, { dir: 'desc', by: 1, at: [['Col', 'A']] });
  const out = result(pivotSourceData(reopened, restored), restored);
  deepEqual(childKeys(out.meta.rowTree), ['X', 'Y']);
  deepEqual(childKeys(out.meta.colTree), ['C', 'B', 'A']);
  equal(out.meta.rowItems.at(-1).kind, 'grand');
  equal(out.meta.colLeaves.at(-1).kind, 'grand');
});
