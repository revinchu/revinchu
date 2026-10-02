import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { pivotSourceData, resolvePivot, computePivot, pivotPageLayout } from '../src/pivot.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child } from '../src/xml.js';

const part = 'xl/pivotTables/pivotTable1.xml';
function fixture(body, pageOrder, pageWrap) {
  const wb = new Workbook(), rows = [
    ['지역', '상품', '연도', '경로', '팀', '매출'],
    ['서울', 'A', '2025', '웹', '1팀', 10], ['부산', 'B', '2026', '매장', '2팀', 20],
  ];
  wb.sheets[0].name = '원본';
  wb.transact(() => {
    rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.addSheet('보고서');
  });
  const def = { name: '보고서필터', source: '원본', range: { r1: 0, c1: 0, r2: 2, c2: 5 },
    top: 4, left: 1, pages: rows[0].slice(0, 4), pageOrder, pageWrap,
    rows: body ? ['팀'] : [], cols: [], values: body ? [{ field: '매출', agg: 'sum' }] : [], style: 'PivotStyleLight16',
    fieldCaptions: { 지역: '지역 이름' }, filters: { 지역: ['서울'] }, pageMulti: { 지역: false } };
  const resolved = resolvePivot(pivotSourceData(wb, def), def), result = computePivot(resolved, resolved.def);
  let width = 0;
  wb.transact(() => {
    result.grid.forEach((row, r) => { width = Math.max(width, row.length); row.forEach((cd, c) => {
      if (cd) wb.setCellData(1, def.top + r, def.left + c, { raw: cd.raw, style: cd.style });
    }); });
    def.area = { r1: def.top, c1: def.left, r2: def.top + result.grid.length - 1, c2: def.left + width - 1 };
    wb.setSheetProp(1, 'pivot', def);
  });
  return { wb, def, result };
}

for (const body of [false, true]) for (const order of ['down', 'over']) for (const wrap of [0, 2]) {
  test(`보고서 필터 XLSX 표준 왕복: body=${body}, ${order}, wrap=${wrap}`, () => {
    const { wb, def, result } = fixture(body, order, wrap), before = wb.serialize();
    const bytes = writeXlsx(wb), files = unzip(bytes);
    assert.ok(files[part], '필터 전용 피벗도 표준 pivotTable 파트가 있어야 함');
    const table = parseXml(textOf(files[part])), loc = child(table, 'location'), layout = pivotPageLayout(def);
    assert.equal(table.attrs.pageOverThenDown ?? '0', order === 'over' ? '1' : '0');
    assert.equal(Number(table.attrs.pageWrap ?? 0), wrap);
    assert.equal(Number(loc.attrs.rowPageCount), layout.height);
    assert.equal(Number(loc.attrs.colPageCount), (layout.width + 1) / 3);
    const bodyStart = def.top + layout.height + 1;
    if (!body) {
      assert.equal(loc.attrs.ref, `B${bodyStart + 1}`);
      assert.equal(loc.attrs.firstDataRow, '0'); assert.equal(loc.attrs.firstDataCol, '0');
      assert.equal(child(table, 'rowItems'), null); assert.equal(child(table, 'colItems'), null);
      assert.equal(child(table, 'dataFields'), null, '빈 dataFields는 Excel에서 정상 열기를 거부함');
    } else assert.match(loc.attrs.ref, new RegExp(`^B${bodyStart + 1}:C`), '본문 폭은 넓은 필터 폭과 분리');
    const imported = new Workbook(readXlsx(bytes).data), actual = imported.sheets[1].pivot;
    assert.ok(actual); assert.deepEqual(actual.pages, def.pages);
    assert.equal(actual.pageOrder, order); assert.equal(actual.pageWrap, wrap);
    assert.equal(actual.top, def.top); assert.equal(actual.left, def.left);
    assert.deepEqual(actual.area, def.area, '필터+본문 실제 출력 범위, 전용일 때 빈 앵커는 제외');
    assert.deepEqual(actual.filters.지역, ['서울']); assert.equal(actual.pageMulti.지역, false);
    assert.equal(actual.fieldCaptions.지역, '지역 이름');
    const res = resolvePivot(pivotSourceData(imported, actual), actual), again = computePivot(res, res.def);
    assert.deepEqual(again.meta.pageFields, result.meta.pageFields);
    for (const p of again.meta.pageFields) {
      assert.equal(again.grid[p.r][p.c].raw, result.grid[p.r][p.c].raw);
      assert.equal(again.grid[p.r][p.c + 1].raw, result.grid[p.r][p.c + 1].raw);
    }
    const reread = readXlsx(writeXlsx(imported)).data.sheets[1].pivot;
    assert.deepEqual(reread.area, actual.area); assert.equal(reread.pageOrder, order); assert.equal(reread.pageWrap, wrap);
    assert.deepEqual(wb.serialize(), before, '내보내기로 원본 통합 문서를 변경하지 않음');
  });
}

test('Excel 필터 전용 단일 셀 location과 true 속성은 가로 필터 출력 영역으로 읽음', () => {
  const files = unzip(writeXlsx(fixture(false, 'over', 2).wb));
  files[part] = textOf(files[part]).replace(/pageOverThenDown="1"/, 'pageOverThenDown="true"')
    .replace(/<location[^>]+\/>/, '<location ref="B8" firstHeaderRow="0" firstDataRow="0" firstDataCol="0" rowPageCount="2" colPageCount="2"/>');
  const def = readXlsx(zip(files)).data.sheets[1].pivot;
  assert.equal(def.top, 4); assert.equal(def.left, 1);
  assert.deepEqual(def.area, { r1: 4, c1: 1, r2: 5, c2: 5 });
});

test('보고서 필터가 없고 필드도 없는 피벗 저장 정책은 기존대로 유지', () => {
  const { wb } = fixture(false, 'down', 0);
  wb.sheets[1].pivot.pages = [];
  assert.equal(unzip(writeXlsx(wb))[part], undefined);
});
