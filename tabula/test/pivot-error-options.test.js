import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { computePivot, resolvePivot, pivotSourceData, pivotErrorDisplay } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';

function fixture() {
  const wb = new Workbook();
  const data = [['지역', '비용', '클릭'], ['오류행', '10', '0'], ['정상행', '30', '3']];
  wb.transact(() => data.forEach((row, r) => row.forEach((raw, c) => wb.setInput(0, r, c, raw))));
  wb.transact(() => wb.setSheetProp(0, 'pivot', {
    name: '오류옵션', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: 'CPC', agg: 'sum' }],
    calcFields: [{ name: 'CPC', formula: '비용/클릭' }], layout: 'tabular', top: 0, left: 5, area: { r1: 0, c1: 5, r2: 3, c2: 6 },
  }));
  return wb;
}
function result(wb) {
  const def = wb.sheets[0].pivot;
  const res = resolvePivot(pivotSourceData(wb, def), def);
  const grid = computePivot(res, res.def).grid;
  return { error: grid.find((r) => r[0]?.raw === '오류행')[1].raw, valid: grid.find((r) => r[0]?.raw === '정상행')[1].raw };
}
for (const [flag, caption, enabled] of [[null, null, false], [null, '-', false], ['0', '', false], ['0', '-', false], ['false', '-', false], ['1', null, true], ['1', '', true], ['1', '-', true], ['true', '0', true]]) {
  test('피벗 오류 옵션 XLSX 가져오기·계산·왕복: showError=' + flag + ', caption=' + JSON.stringify(caption), () => {
    const files = unzip(writeXlsx(fixture()));
    const path = 'xl/pivotTables/pivotTable1.xml';
    files[path] = textOf(files[path]).replace(/ showError="[^"]*"| errorCaption="[^"]*"/g, '').replace('<pivotTableDefinition ', '<pivotTableDefinition ' + (flag === null ? '' : 'showError="' + flag + '" ') + (caption === null ? '' : 'errorCaption="' + caption + '" '));
    const imported = new Workbook(readXlsx(zip(files)).data);
    const def = imported.sheets[0].pivot;
    assert.equal(def.errorShow, enabled);
    assert.equal(pivotErrorDisplay(def), enabled, '옵션 UI와 같은 판정');
    assert.equal(def.errorCaption, caption ?? (enabled ? '' : undefined));
    const error = !enabled ? '#DIV/0!' : caption === '-' ? '0' : caption ?? '';
    assert.deepEqual(result(imported), { error, valid: '10' });
    const exported = writeXlsx(imported);
    const xml = textOf(unzip(exported)[path]);
    assert.ok(xml.includes('showError="' + (enabled ? '1' : '0') + '"'));
    const back = new Workbook(readXlsx(exported).data);
    assert.equal(pivotErrorDisplay(back.sheets[0].pivot), enabled);
    assert.equal(back.sheets[0].pivot.errorCaption, def.errorCaption);
    assert.deepEqual(result(back), { error, valid: '10' });
  });
}
test('새 피벗 빈 오류 기본과 명시적 true/false 및 이전 JSON 정의를 구분', () => {
  const wb = fixture(), def = wb.sheets[0].pivot;
  assert.equal(pivotErrorDisplay(def), false);
  def.errorCaption = '';
  assert.equal(pivotErrorDisplay(def), true);
  assert.equal(result(wb).error, '');
  def.errorShow = false;
  assert.equal(pivotErrorDisplay(def), false);
  assert.equal(result(wb).error, '#DIV/0!');
  def.errorShow = true; delete def.errorCaption;
  assert.equal(pivotErrorDisplay(def), true);
  assert.equal(result(wb).error, '');
});
