import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';

const pivot = (name = 'Summary', left = 6) => ({
  name, source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 3 },
  rows: ['Item', 'Region'], cols: [], pages: [], filters: {},
  values: [{ field: 'Amount', agg: 'sum', name: 'Revenue' }, { field: 'Count', agg: 'sum', name: 'Units' }],
  top: 0, left, area: { r1: 0, c1: left, r2: 6, c2: left + 3 }, layout: 'tabular',
});
function book() {
  const wb = new Workbook();
  const rows = [['Item', 'Region', 'Amount', 'Count'], ['A', 'North', 10, 1], ['A', 'South', 20, 2], ['B', 'North', 30, 3]];
  wb.transact(() => {
    rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.setSheetProp(0, 'pivot', pivot());
  });
  return wb;
}
const rule = (scope = 'selection', extra = {}) => ({
  r1: 1, c1: 9, r2: 2, c2: 9, type: 'formula', formula: '=J2>0', stopIfTrue: true,
  style: { color: '#123456', fill: '#fedcba' },
  pivot: { name: 'Summary', scope, value: 'Units' }, ...extra,
});
const setRules = (wb, rules) => wb.transact(() => wb.setSheetProp(0, 'cond', rules));
const worksheetFormats = parts => kids(parseXml(textOf(parts['xl/worksheets/sheet1.xml'])), 'conditionalFormatting');
function tableFormats(parts, name = 'Summary') {
  const path = Object.keys(parts).find(path => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(path) && parseXml(textOf(parts[path])).attrs.name === name);
  assert.ok(path, '피벗 정의가 저장되어야 한다');
  return kids(child(parseXml(textOf(parts[path])), 'conditionalFormats'), 'conditionalFormat');
}
const metadata = rules => rules.map(({ r1, c1, r2, c2, more, pivot }) => ({ r1, c1, r2, c2, ...(more ? { more } : {}), ...(pivot ? { pivot } : {}) }));
const digest = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');

test('선택 영역 피벗 CF는 worksheet pivot 표시와 selection scope·두 번째 값 참조를 저장한다', () => {
  const wb = book(), selected = rule('selection', { more: [{ r1: 4, c1: 9, r2: 4, c2: 9 }] });
  setRules(wb, [selected]);
  const before = structuredClone(wb.sheets[0].cond), parts = unzip(writeXlsx(wb));
  const [worksheet] = worksheetFormats(parts), [cf] = tableFormats(parts);
  assert.equal(worksheet.attrs.pivot, '1');
  assert.equal(worksheet.attrs.sqref, 'J2:J3 J5');
  assert.equal(child(worksheet, 'cfRule').attrs.priority, '1');
  assert.equal(child(worksheet, 'cfRule').attrs.stopIfTrue, '1');
  assert.equal(child(child(worksheet, 'cfRule'), 'formula').text, 'J2>0');
  assert.deepEqual(cf.attrs, { scope: 'selection', priority: '1' });
  const refs = descendants(cf, 'reference');
  assert.equal(refs.length, 1);
  assert.deepEqual(refs[0].attrs, { field: '4294967294', count: '1', selected: '0' });
  assert.equal(child(refs[0], 'x').attrs.v, '1');
  assert.deepEqual(wb.sheets[0].cond, before, '저장할 때 원래 규칙을 변경하지 않는다');
  const back = readXlsx(writeXlsx(wb)).data.sheets[0].cond[0];
  assert.deepEqual(metadata([back]), metadata([selected]));
  assert.equal(back.formula, selected.formula);
  assert.equal(back.style.color, selected.style.color);
  assert.equal(back.style.fill, selected.style.fill);
});

test('선택 영역의 기존 행 필드 참조·이름·범위는 두 차례 왕복에도 보존된다', () => {
  const wb = book(), selected = rule('selection', { pivot: { name: 'Summary', scope: 'selection', value: 'Units', rowField: 'Region' } });
  setRules(wb, [selected]);
  const expected = metadata(wb.sheets[0].cond), expectedHash = digest(expected);
  let current = wb;
  for (let cycle = 0; cycle < 2; cycle++) {
    const bytes = writeXlsx(current), refs = descendants(tableFormats(unzip(bytes))[0], 'reference');
    assert.deepEqual(refs.map(ref => ref.attrs), [
      { field: '4294967294', count: '1', selected: '0' }, { field: '1', count: '0', selected: '0' },
    ]);
    current = new Workbook(readXlsx(bytes).data);
    assert.deepEqual(metadata(current.sheets[0].cond), expected);
    assert.equal(digest(metadata(current.sheets[0].cond)), expectedHash);
  }
});

test('일반 CF는 pivot 표시 없이 유지되고 data·field·selection의 전역 우선순위가 같다', () => {
  const wb = book();
  const ordinary = { r1: 1, c1: 2, r2: 3, c2: 2, type: 'gt', v1: '15', style: { fill: '#112233' } };
  const data = rule('data', { r1: 1, c1: 8, r2: 5, c2: 8, type: 'bar', color: '#638ec6', pivot: { name: 'Summary', scope: 'data', value: 'Revenue' } });
  const field = rule('field', { r2: 5, pivot: { name: 'Summary', scope: 'field', value: 'Units', rowField: 'Region' } });
  setRules(wb, [ordinary, data, field, rule()]);
  const parts = unzip(writeXlsx(wb)), formats = worksheetFormats(parts), pivotFormats = tableFormats(parts);
  assert.equal(formats.length, 4);
  assert.equal(formats[0].attrs.pivot, undefined);
  assert.deepEqual(formats.slice(1).map(cf => cf.attrs.pivot), ['1', '1', '1']);
  assert.deepEqual(formats.map(cf => child(cf, 'cfRule').attrs.priority), ['1', '2', '3', '4']);
  assert.deepEqual(pivotFormats.map(cf => cf.attrs), [
    { scope: 'data', priority: '2' }, { scope: 'field', priority: '3' }, { scope: 'selection', priority: '4' },
  ]);
  assert.deepEqual(pivotFormats.map(cf => child(descendants(cf, 'reference')[0], 'x').attrs.v), ['0', '1', '1']);
  assert.equal(descendants(pivotFormats[0], 'reference').length, 1);
  assert.deepEqual(descendants(pivotFormats[1], 'reference')[1].attrs, { field: '1', count: '0', selected: '0' });
  const back = readXlsx(writeXlsx(wb)).data.sheets[0].cond;
  assert.equal(back[0].pivot, undefined);
  assert.equal(back[0].type, 'gt');
  assert.equal(back[0].v1, '15');
  assert.deepEqual(metadata(back), metadata(wb.sheets[0].cond));
});

test('같은 시트의 서로 다른 피벗 selection은 이름별 소유와 전역 우선순위를 유지한다', () => {
  const wb = book();
  wb.transact(() => wb.setSheetProp(0, 'pivotsExtra', [pivot('OtherSummary', 12)]));
  setRules(wb, [rule(), rule('selection', { r1: 1, c1: 15, r2: 3, c2: 15, formula: '=P2>0', pivot: { name: 'OtherSummary', scope: 'selection', value: 'Revenue' } })]);
  const bytes = writeXlsx(wb), parts = unzip(bytes);
  assert.deepEqual(tableFormats(parts).map(cf => cf.attrs), [{ scope: 'selection', priority: '1' }]);
  assert.deepEqual(tableFormats(parts, 'OtherSummary').map(cf => cf.attrs), [{ scope: 'selection', priority: '2' }]);
  const back = readXlsx(bytes).data.sheets[0];
  assert.equal(back.pivot.name, 'Summary');
  assert.equal(back.pivotsExtra[0].name, 'OtherSummary');
  assert.deepEqual(metadata(back.cond), metadata(wb.sheets[0].cond));
});

test('scope가 생략된 기존 XML의 선택 연결도 다시 저장할 때 명시하여 유지한다', () => {
  const wb = book();
  setRules(wb, [rule()]);
  const parts = unzip(writeXlsx(wb)), path = 'xl/pivotTables/pivotTable1.xml';
  parts[path] = textOf(parts[path]).replace('<conditionalFormat scope="selection"', '<conditionalFormat');
  const first = new Workbook(readXlsx(zip(parts)).data);
  assert.deepEqual(first.sheets[0].cond[0].pivot, { name: 'Summary', scope: 'selection', value: 'Units' });
  const saved = writeXlsx(first), cf = tableFormats(unzip(saved))[0];
  assert.equal(cf.attrs.scope, 'selection');
  assert.equal(worksheetFormats(unzip(saved))[0].attrs.pivot, '1');
  assert.deepEqual(metadata(readXlsx(saved).data.sheets[0].cond), metadata(first.sheets[0].cond));
});

test('앱 방식으로 수동 선택 규칙을 추가·Undo·Redo한 뒤 저장해도 범위와 연결을 보존한다', () => {
  const wb = book(), selected = rule('selection', { pivot: { name: 'Summary', scope: 'selection', value: 'Revenue', rowField: 'Item' } });
  setRules(wb, [selected]);
  wb.undo();
  assert.deepEqual(wb.sheets[0].cond, []);
  wb.redo();
  assert.deepEqual(wb.sheets[0].cond, [selected]);
  const before = digest(metadata(wb.sheets[0].cond)), back = readXlsx(writeXlsx(wb)).data.sheets[0].cond;
  assert.equal(digest(metadata(back)), before);
  assert.deepEqual(back[0].pivot, selected.pivot);
});

test('selection 연결은 동기·stream bytes·stream Blob 가져오기 모두에서 동일하다', async () => {
  const wb = book();
  setRules(wb, [rule('selection', { pivot: { name: 'Summary', scope: 'selection', value: 'Units', rowField: 'Region' } })]);
  const bytes = writeXlsx(wb), expected = metadata(wb.sheets[0].cond);
  const results = [readXlsx(bytes), await readXlsxAsync(bytes, undefined, { streamThreshold: 0 }), await readXlsxAsync(new Blob([bytes]), undefined, { streamThreshold: 0 })];
  for (const result of results) assert.deepEqual(metadata(result.data.sheets[0].cond), expected);
});
