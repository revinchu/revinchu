import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child } from '../src/xml.js';

// All fixtures are synthetic. Coordinates model a nonzero frozen origin without user data.
function fixture(views) {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, '합성');
    wb.setInput(0, 60, 7, '=2+3');
    wb.setRowHeight(0, 31, 29);
  });
  wb.sheets[0].hiddenRows = { 29: true, 30: true, 33: true };
  const files = unzip(writeXlsx(wb));
  files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml'])
    .replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, `<sheetViews>${views}</sheetViews>`);
  return new Workbook(readXlsx(zip(files)).data);
}
function saved(wb) {
  const bytes = writeXlsx(wb);
  const root = parseXml(textOf(unzip(bytes)['xl/worksheets/sheet1.xml']));
  const view = child(child(root, 'sheetViews'), 'sheetView');
  return { wb: new Workbook(readXlsx(bytes).data), view, pane: child(view, 'pane'), selection: child(view, 'selection') };
}

test('행 고정 원점·본문 스크롤·활성 pane 셀을 분리하고 숨김 행과 값을 보존한다', () => {
  const wb = fixture('<sheetView workbookViewId="0" topLeftCell="D30" showGridLines="0" zoomScale="330"><pane ySplit="5" topLeftCell="D80" state="frozen" activePane="bottomLeft"/><selection activeCell="F12" sqref="F12"/><selection pane="bottomLeft" activeCell="E82" sqref="E82"/></sheetView>');
  const s = wb.sheets[0];
  assert.deepEqual(s.freeze, { rows: 5, cols: 0, top: 29 });
  assert.deepEqual(s.view, { top: 79, left: 3, r: 81, c: 4, activePane: 'bottomLeft' });
  const before = wb.serialize(), back = saved(wb);
  assert.equal(back.view.attrs.topLeftCell, 'D30');
  assert.equal(back.pane.attrs.topLeftCell, 'D80');
  assert.equal(back.pane.attrs.ySplit, '5');
  assert.equal(back.selection.attrs.activeCell, 'E82');
  assert.equal(back.selection.attrs.pane, 'bottomLeft');
  assert.deepEqual(back.wb.sheets[0].freeze, s.freeze);
  assert.deepEqual(back.wb.sheets[0].view, s.view);
  assert.deepEqual(back.wb.sheets[0].hiddenRows, { 29: true, 30: true, 33: true });
  assert.equal(back.wb.sheets[0].rowHeights[31], 29);
  assert.equal(back.wb.sheets[0].zoom, 330);
  assert.equal(back.wb.sheets[0].noGrid, true);
  assert.equal(back.wb.getValue(0, 60, 7), 5);
  assert.equal(back.wb.getRaw(0, 0, 0), '합성');
  assert.deepEqual(wb.serialize(), before, '저장은 원본 모델을 바꾸지 않는다');
});

test('열만 고정하면 세로 스크롤을 잃지 않고 원점은 고정축에만 둔다', () => {
  const wb = fixture('<sheetView workbookViewId="0" topLeftCell="G21"><pane xSplit="3" topLeftCell="M41" state="frozen" activePane="topRight"/><selection pane="topRight" activeCell="O45"/></sheetView>');
  assert.deepEqual(wb.sheets[0].freeze, { rows: 0, cols: 3, left: 6 });
  assert.deepEqual(wb.sheets[0].view, { top: 40, left: 12, r: 44, c: 14, activePane: 'topRight' });
  const back = saved(wb);
  assert.equal(back.view.attrs.topLeftCell, 'G41');
  assert.equal(back.pane.attrs.topLeftCell, 'M41');
  assert.deepEqual(back.wb.sheets[0].view, wb.sheets[0].view);
  assert.deepEqual(back.wb.sheets[0].freeze, wb.sheets[0].freeze);
});

test('양쪽 고정과 고정창 자체의 활성 선택도 표준 XML로 왕복한다', () => {
  const wb = fixture('<sheetView workbookViewId="0" topLeftCell="D11"><pane xSplit="2" ySplit="4" topLeftCell="H31" state="frozenSplit" activePane="topLeft"/><selection pane="bottomRight" activeCell="H31"/><selection pane="topLeft" activeCell="E12"/></sheetView>');
  assert.deepEqual(wb.sheets[0].freeze, { rows: 4, cols: 2, top: 10, left: 3 });
  assert.deepEqual(wb.sheets[0].view, { top: 30, left: 7, r: 11, c: 4, activePane: 'topLeft' });
  const back = saved(wb);
  assert.equal(back.view.attrs.topLeftCell, 'D11');
  assert.equal(back.pane.attrs.topLeftCell, 'H31');
  assert.equal(back.pane.attrs.activePane, 'topLeft');
  assert.equal(back.selection.attrs.activeCell, 'E12');
  assert.deepEqual(back.wb.sheets[0].view, wb.sheets[0].view);
});

test('pane 본문 좌표가 없으면 origin + split, 고정되지 않은 축은 sheetView 좌표다', () => {
  for (const [attrs, freeze, view] of [
    ['ySplit="5"', { rows: 5, cols: 0, top: 29 }, { top: 34, left: 3, activePane: 'bottomLeft' }],
    ['xSplit="2"', { rows: 0, cols: 2, left: 3 }, { top: 29, left: 5, activePane: 'topRight' }],
    ['ySplit="5" xSplit="2"', { rows: 5, cols: 2, top: 29, left: 3 }, { top: 34, left: 5, activePane: 'bottomRight' }],
  ]) {
    const wb = fixture(`<sheetView workbookViewId="0" topLeftCell="D30"><pane ${attrs} state="frozen"/></sheetView>`);
    assert.deepEqual(wb.sheets[0].freeze, freeze);
    assert.deepEqual(wb.sheets[0].view, view);
    assert.deepEqual(saved(wb).wb.sheets[0].view, view);
  }
});

test('본문 좌표가 고정 경계보다 앞이면 경계로 제한하며 모델은 변형하지 않는다', () => {
  const wb = fixture('<sheetView workbookViewId="0" topLeftCell="D30"><pane xSplit="2" ySplit="5" topLeftCell="A1" state="frozen"/></sheetView>');
  assert.deepEqual(wb.sheets[0].view, { top: 34, left: 5, activePane: 'bottomRight' });
  wb.sheets[0].view = { top: 0, left: 0, r: 35, c: 6, activePane: 'bottomRight' };
  const before = structuredClone(wb.sheets[0].view), back = saved(wb);
  assert.equal(back.pane.attrs.topLeftCell, 'F35');
  assert.deepEqual(wb.sheets[0].view, before);
});

test('workbookViewId 0의 pane와 선택만 읽고 다른 통합문서 창은 섞지 않는다', () => {
  const wb = fixture('<sheetView workbookViewId="1" topLeftCell="Z100" zoomScale="25"><pane xSplit="12" ySplit="20" state="frozen"/><selection activeCell="Z100"/></sheetView><sheetView workbookViewId="0" topLeftCell="C8" zoomScale="115"><pane ySplit="2" topLeftCell="C40" state="frozen" activePane="bottomLeft"/><selection activeCell="D41"/></sheetView>');
  assert.deepEqual(wb.sheets[0].freeze, { rows: 2, cols: 0, top: 7 });
  assert.deepEqual(wb.sheets[0].view, { top: 39, left: 2, r: 40, c: 3, activePane: 'bottomLeft' });
  assert.equal(wb.sheets[0].zoom, 115);
});

test('일반 분할의 point 단위를 고정 행 수로 오인하지 않는다', () => {
  const wb = fixture('<sheetView workbookViewId="0" topLeftCell="C12"><pane xSplit="2160" ySplit="3600" topLeftCell="J50" state="split" activePane="bottomRight"/><selection pane="bottomRight" activeCell="K51"/></sheetView>');
  assert.deepEqual(wb.sheets[0].freeze, { rows: 0, cols: 0 });
  assert.deepEqual(wb.sheets[0].view, { top: 11, left: 2, r: 50, c: 10 });
  assert.equal(saved(wb).pane, null);
});

test('틀 고정 없는 저장 화면과 A1을 포함한 활성 셀도 왕복한다', () => {
  for (const active of ['A1', 'G91']) {
    const wb = fixture(`<sheetView workbookViewId="0" topLeftCell="F70"><selection activeCell="${active}"/></sheetView>`);
    const back = saved(wb);
    assert.equal(back.view.attrs.topLeftCell, 'F70');
    assert.equal(back.selection.attrs.activeCell, active);
    assert.deepEqual(back.wb.sheets[0].view, wb.sheets[0].view);
  }
});

test('기존 원점 없는 틀 고정 모델은 기존 split 크기를 유지한다', () => {
  const wb = new Workbook();
  wb.sheets[0].freeze = { rows: 2, cols: 1 };
  const back = saved(wb);
  assert.deepEqual(back.wb.sheets[0].freeze, { rows: 2, cols: 1 });
  assert.equal(back.view.attrs.topLeftCell, undefined);
  assert.equal(back.pane.attrs.topLeftCell, 'B3');
  assert.deepEqual(back.wb.sheets[0].view, { top: 2, left: 1, activePane: 'bottomRight' });
});

test('행 삽입은 고정 원점/구간/본문/활성 셀을 이동하며 끝 경계는 본문이다', () => {
  for (const [index, expectedFreeze, expectedTop] of [
    [5, { rows: 5, cols: 2, top: 12, left: 3 }, 32],
    [10, { rows: 5, cols: 2, top: 12, left: 3 }, 32],
    [12, { rows: 7, cols: 2, top: 10, left: 3 }, 32],
    [15, { rows: 5, cols: 2, top: 10, left: 3 }, 32],
    [35, { rows: 5, cols: 2, top: 10, left: 3 }, 30],
  ]) {
    const wb = new Workbook(), s = wb.sheets[0];
    s.freeze = { rows: 5, cols: 2, top: 10, left: 3 };
    s.view = { top: 30, left: 8, r: 40, c: 9, activePane: 'bottomRight' };
    wb.transact(() => wb.shiftAxis(0, 'row', index, 2));
    assert.deepEqual(wb.sheets[0].freeze, expectedFreeze, `index=${index}`);
    assert.deepEqual(wb.sheets[0].view, { top: expectedTop, left: 8, r: 42, c: 9, activePane: 'bottomRight' });
    wb.undo();
    assert.deepEqual(wb.sheets[0].freeze, { rows: 5, cols: 2, top: 10, left: 3 });
    assert.equal(wb.sheets[0].view.top, 30);
    wb.redo();
    assert.deepEqual(wb.sheets[0].freeze, expectedFreeze);
  }
});

test('행 삭제는 고정 구간 경계를 접고 사라진 pane을 정리한다', () => {
  for (const [index, count, freeze, pane] of [
    [3, -2, { rows: 5, cols: 2, top: 8, left: 3 }, 'bottomRight'],
    [8, -4, { rows: 3, cols: 2, top: 8, left: 3 }, 'bottomRight'],
    [12, -2, { rows: 3, cols: 2, top: 10, left: 3 }, 'bottomRight'],
    [12, -10, { rows: 2, cols: 2, top: 10, left: 3 }, 'bottomRight'],
    [9, -10, { rows: 0, cols: 2, left: 3 }, 'topRight'],
  ]) {
    const wb = new Workbook();
    wb.sheets[0].freeze = { rows: 5, cols: 2, top: 10, left: 3 };
    wb.sheets[0].view = { top: 30, left: 8, r: 31, c: 9, activePane: 'bottomRight' };
    wb.transact(() => wb.shiftAxis(0, 'row', index, count));
    assert.deepEqual(wb.sheets[0].freeze, freeze);
    assert.deepEqual(wb.sheets[0].view, { top: 30 + count, left: 8, r: 31 + count, c: 9, activePane: pane });
  }
});

test('열 이동도 같은 규칙을 쓰며 없는 view를 만들거나 비고정축을 경계로 밀지 않는다', () => {
  const wb = new Workbook();
  wb.sheets[0].freeze = { rows: 2, cols: 3, top: 7, left: 4 };
  wb.transact(() => wb.shiftAxis(0, 'column', 5, 2));
  assert.deepEqual(wb.sheets[0].freeze, { rows: 2, cols: 5, top: 7, left: 4 });
  assert.equal(wb.sheets[0].view, undefined);
  wb.sheets[0].view = { top: 1, left: 6, r: 1, c: 6, activePane: 'bottomRight' };
  wb.transact(() => wb.shiftAxis(0, 'column', 3, -10));
  assert.deepEqual(wb.sheets[0].freeze, { rows: 2, cols: 0, top: 7 });
  assert.deepEqual(wb.sheets[0].view, { top: 1, left: 3, r: 1, c: 3, activePane: 'bottomLeft' });
  wb.transact(() => wb.shiftAxis(0, 'row', 7, -2));
  assert.deepEqual(wb.sheets[0].freeze, { rows: 0, cols: 0 });
  assert.deepEqual(wb.sheets[0].view, { top: 1, left: 3, r: 1, c: 3 });
});
