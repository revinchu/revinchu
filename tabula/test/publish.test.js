import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { publishedWorkbook } from '../src/publish.js';
import { chartModelData } from '../src/chart.js';
import { ColBuilder } from '../src/block.js';

const secret = 'PRIVATE_ONLY_SENTINEL';
function fixture() {
  const wb = new Workbook();
  wb.load({ version: 1, vba: secret, props: { creator: secret }, names: [{ name: 'private_name', ref: '비공개!A1', note: secret }],
    externals: [{ name: secret, cached: secret }], themeXml: secret, themeName: secret,
    sheets: [
      { name: '공개', cells: { '0,0': { raw: "='비공개'!A1*2", comment: secret, link: '#비공개!A1', cached: -999 }, '1,0': { raw: '공개 항목' } },
        scenarios: [{ name: secret }], external: { cached: secret }, validations: [{ list: [secret] }],
        pivot: { name: 'private_cache', source: '비공개', cached: [secret] }, pivotsExtra: [{ cache: secret }], slicers: [{ cache: secret }],
        tables: [{ name: secret, r1: 10, c1: 0, r2: 11, c2: 1, formula: secret }],
        charts: [{ id: 'c1', type: 'column', title: '공개 차트', x: 0, y: 0, w: 400, h: 200,
          series: [{ name: { text: '공개 값' }, val: { sheet: '비공개', r1: 0, c1: 0, r2: 0, c2: 0 } }, { name: { text: secret }, cache: [999] }],
          hiddenSeries: [1], seriesFmt: [{ color: '#ff0000', marker: 'circle' }, {}], external: secret, rawXml: secret }] },
      { name: '비공개', state: 'hidden', cells: { '0,0': { raw: '21' }, '1,0': { raw: secret } }, images: [{ src: secret }] },
    ] });
  wb.pivotSnapshots = new Map([['private_cache', { rows: [[secret]], ver: undefined }]]);
  return wb;
}
test('개별 시트 공유: 다른 시트·수식·전역/피벗/외부 캐시를 제거하고 값/차트를 유지', () => {
  const wb = fixture(), before = wb.serialize();
  const exported = publishedWorkbook(wb, '0'), json = JSON.stringify(exported);
  assert.equal(exported.sheets.length, 1);
  assert.equal(exported.sheets[0].state, 'visible');
  assert.equal(json.includes(secret), false);
  assert.equal(json.includes('비공개'), false);
  for (const prop of ['vba', 'names', 'externals', 'props', 'themeXml', 'themeName', 'pivotSnapshots']) assert.equal(exported[prop], undefined);
  for (const prop of ['pivot', 'pivotsExtra', 'external', 'validations', 'scenarios', 'slicers', 'tables']) assert.equal(exported.sheets[0][prop], undefined);
  assert.equal(exported.sheets[0].cells['0,0'].raw, '42');
  const back = new Workbook(); back.load(exported);
  assert.equal(back.getValue(0, 0, 0), 42);
  const chart = back.sheets[0].charts[0], data = chartModelData(back, 0, chart);
  assert.deepEqual(data.categories, ['1']);
  assert.deepEqual(data.series[0].values, [21]);
  assert.equal(data.series[0].color, '#ff0000');
  assert.equal(data.series[0].marker, 'circle');
  assert.equal(data.series.length, 1);
  for (const prop of ['series', 'range', 'sheet', 'pivot', 'seriesFmt', 'hiddenSeries', 'external', 'rawXml']) assert.equal(chart[prop], undefined);
  assert.deepEqual(wb.serialize(), before, '원본 셀·시트·메타데이터가 바뀌면 안 됨');
  assert.equal(wb.pivotSnapshots.get('private_cache').rows[0][0], secret);
});
test('개별 시트 공유: 분산 배열과 문자열/논리/오류 값 타입을 고정', () => {
  const wb = new Workbook(); wb.load({ sheets: [{ name: '값', cells: {
    '0,0': { raw: '=SEQUENCE(2,2)' }, '3,0': { raw: '="001"' }, '4,0': { raw: '="=1+1"' },
    '5,0': { raw: '=TRUE()' }, '6,0': { raw: '=1/0' }, '7,0': { raw: "''앞", style: { numFmt: 'text' } },
  } }] });
  const exported = publishedWorkbook(wb, 0), back = new Workbook(); back.load(exported);
  assert.equal(back.getValue(0, 0, 0), 1); assert.equal(back.getValue(0, 0, 1), 2);
  assert.equal(back.getValue(0, 1, 0), 3); assert.equal(back.getValue(0, 1, 1), 4);
  assert.equal(back.getValue(0, 3, 0), '001'); assert.equal(back.getValue(0, 4, 0), '=1+1');
  assert.equal(back.getValue(0, 5, 0), true); assert.equal(back.getValue(0, 6, 0).code, '#DIV/0!');
  assert.equal(back.getValue(0, 7, 0), "''앞");
  for (const cell of back.sheets[0].cells.values()) assert.equal(cell.formula, undefined);
});
test('개별 시트 공유: 블록 값만 풀고 사용하지 않는 사전/캐시를 버림', () => {
  const builder = new ColBuilder(); builder.set(0, 'A'); builder.set(1, 'B'); builder.set(2, 7);
  const column = builder.finish(3); column.dict.push(secret);
  const wb = new Workbook(); wb.sheets[0].blocks = [{ r0: 2, c0: 1, n: 3, cols: [column], perm: new Uint32Array([2, 0, 1]) }];
  const before = wb.serialize(), exported = publishedWorkbook(wb, 0);
  assert.equal(JSON.stringify(exported).includes(secret), false); assert.equal(exported.sheets[0].blocks, undefined);
  const back = new Workbook(); back.load(exported);
  for (let r = 2; r < 5; r++) assert.equal(back.getValue(0, r, 1), wb.getValue(0, r, 1));
  assert.deepEqual(wb.serialize(), before);
});
test('개별 시트 공유: 조건부 서식 색은 고정하고 조건의 비공개 수식은 제거', () => {
  const wb = new Workbook(); wb.load({ sheets: [{ name: '색', cells: { '0,0': { raw: '5' } }, cond: [
    { type: 'gt', v1: '2', r1: 0, c1: 0, r2: 0, c2: 0, style: { fill: '#112233' }, privateCache: secret },
  ] }] });
  const exported = publishedWorkbook(wb, 0);
  assert.equal(exported.sheets[0].cells['0,0'].style.fill, '#112233');
  assert.equal(exported.sheets[0].cond, undefined); assert.equal(JSON.stringify(exported).includes(secret), false);
});
test('개별 시트 공유: 희소 시트는 전체 사각형을 순회하지 않고 실제 셀 수로 제한', () => {
  const wb = new Workbook(); wb.load({ sheets: [{ name: '희소', cells: { '0,0': { raw: '1' }, '999999,16000': { raw: '2' } } }] });
  assert.equal(Object.keys(publishedWorkbook(wb, 0, { maxCells: 2 }).sheets[0].cells).length, 2);
  assert.throws(() => publishedWorkbook(wb, 0, { maxCells: 1 }), /게시/);
  assert.throws(() => publishedWorkbook(wb, -1), /시트/);
});
test('전체 통합 문서 공유는 숨긴 시트를 포함한 원본 전체를 보존', () => {
  const wb = fixture();
  assert.deepEqual(publishedWorkbook(wb, 'all'), wb.serialize());
  assert.equal(publishedWorkbook(wb, 'all').sheets[1].state, 'hidden');
});
