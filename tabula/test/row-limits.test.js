import test from 'node:test';
import assert from 'node:assert/strict';
import { MAX_ROWS, EXCEL_MAX_ROWS, MAX_COLS } from '../src/formula.js';
import { worksheetRowLimit, normalizeWorksheetRowLimit } from '../src/row-limits.js';
import { Workbook } from '../src/workbook.js';
import { Axis } from '../src/axis.js';
import { GridView } from '../src/view.js';
import { GridAccessibility, gridCoordinates, selectionDescription } from '../src/grid-a11y.js';

test('확장 행은 strict true일 때만 켜지며 엔진의 2천만 행은 유지된다', () => {
  assert.equal(MAX_ROWS, 20_000_000);
  assert.equal(EXCEL_MAX_ROWS, 1_048_576);
  for (const options of [undefined, null, {}, { extendedRows: false }, { extendedRows: 1 }, { extendedRows: 'true' }, { extendedRows: [] }]) assert.equal(worksheetRowLimit(options), EXCEL_MAX_ROWS);
  assert.equal(worksheetRowLimit({ extendedRows: true }), MAX_ROWS);
  for (const value of [undefined, null, 0, -1, 123, Infinity, NaN, String(MAX_ROWS), EXCEL_MAX_ROWS]) assert.equal(normalizeWorksheetRowLimit(value), EXCEL_MAX_ROWS);
  assert.equal(normalizeWorksheetRowLimit(MAX_ROWS), MAX_ROWS);
});

function grid(rowLimit = EXCEL_MAX_ROWS, extra = {}) {
  const wb = new Workbook({ sheets: [{ name: '행 한도 검사', cells: { '0,0': { raw: '첫 행' }, [`${EXCEL_MAX_ROWS + 5},0`]: { raw: '보존할 확장 데이터' } }, ...extra }] });
  const state = { wb, si: 0, rowLimit, showHeaders: true, active: { r: 0, c: 0 }, sel: { r1: 0, c1: 0, r2: 0, c2: 0 }, selKind: 'cells' };
  const view = Object.create(GridView.prototype), counts = [];
  Object.assign(view, { host: { state: () => state }, panes: [], z: 1, sx: 0, sy: 0, extR: 100, extC: 26, hw: 34, hh: 20,
    scroll: { clientWidth: 800, clientHeight: 500, scrollWidth: 2000, scrollHeight: 15000000, scrollLeft: 0, scrollTop: 0 },
    sizer: { style: {} }, viewEl: { style: { setProperty() {} } }, freezeV: { style: {} }, freezeH: { style: {} },
    a11y: { setRowLimit: n => counts.push(n), update() {} }, renderHeaders() {}, renderAll() {} });
  view.refreshAxes();
  return { view, state, wb, counts };
}

test('화면 축은 기본 Excel 한도이며 확장 모드에서만 끝 행에 도달한다', () => {
  const { view, state, wb } = grid();
  assert.equal(view.rowLimit, EXCEL_MAX_ROWS); assert.equal(view.rows.max, EXCEL_MAX_ROWS);
  view.ensureExtentFor(MAX_ROWS - 1, 0);
  assert.equal(view.extR, EXCEL_MAX_ROWS);
  assert.equal(view.rows.indexAt(Infinity), EXCEL_MAX_ROWS - 1);
  assert.equal(view.rows.keys.length, 0); assert.equal(view.rows.cum.length, 1);
  state.rowLimit = MAX_ROWS; view.refreshAxes(); view.ensureExtentFor(MAX_ROWS - 1, 0);
  assert.equal(view.rows.max, MAX_ROWS); assert.equal(view.extR, MAX_ROWS);
  assert.equal(view.rows.indexAt(Infinity), MAX_ROWS - 1);
  assert.equal(wb.sheets[0].cells.size, 2);
});

test('한도를 줄이면 기존 extR·스크롤·가상 창과 접근성 행 수를 함께 갱신한다', () => {
  const { view, state, wb, counts } = grid(MAX_ROWS);
  const original = JSON.stringify(wb.serialize());
  view.extR = MAX_ROWS; view.sy = 1e9; view.scroll.scrollTop = 15000000;
  state.rowLimit = EXCEL_MAX_ROWS;
  view.update();
  assert.equal(view.rows.max, EXCEL_MAX_ROWS); assert.equal(view.extR, EXCEL_MAX_ROWS);
  assert.ok(view.sy <= view.maxScroll().y);
  assert.equal(counts.at(-1), EXCEL_MAX_ROWS);
  assert.deepEqual(view.visibleRange({ scrollX: true, scrollY: true }, { w: 300, h: 500 }).r2, EXCEL_MAX_ROWS - 1);
  assert.equal(JSON.stringify(wb.serialize()), original);
  state.rowLimit = MAX_ROWS; view.refreshAxes();
  view.panes.push({ win: { r1: MAX_ROWS - 50, r2: MAX_ROWS - 1 } });
  state.rowLimit = EXCEL_MAX_ROWS; view.refreshAxes();
  assert.equal(view.panes[0].win, null);
});

test('사용 범위가 커도 extent 축소가 작동하고 자동 확장은 마지막 표시행에서 멈춘다', () => {
  const { view, state } = grid(MAX_ROWS);
  view.extR = MAX_ROWS;
  state.rowLimit = EXCEL_MAX_ROWS;
  assert.equal(view.ensureExtentFor(0, 0), true);
  assert.equal(view.extR, EXCEL_MAX_ROWS);
  view.refreshAxes(); view.extC = MAX_COLS;
  view.readScroll = () => { view.sy = view.maxScroll().y; }; view.update = () => {};
  view.onScroll();
  assert.equal(view.extR, EXCEL_MAX_ROWS);
});

test('범위 좌표·강제 표시와 저장된 틀 고정은 현재 행 한도를 넘지 않는다', () => {
  const { view, wb } = grid(EXCEL_MAX_ROWS, { freeze: { top: MAX_ROWS - 10, rows: 5 }, rowHeights: { [EXCEL_MAX_ROWS + 2]: 55 } });
  assert.ok(view.fr < EXCEL_MAX_ROWS); assert.ok(view.frozenTop < EXCEL_MAX_ROWS);
  assert.equal(wb.sheets[0].freeze.top, MAX_ROWS - 10);
  const outside = view.sheetRect({ r1: EXCEL_MAX_ROWS + 5, r2: MAX_ROWS - 1, c1: 0, c2: 0 });
  assert.equal(outside.h, 0);
  assert.equal(outside.y, view.rows.pos(EXCEL_MAX_ROWS));
  const crossing = view.sheetRect({ r1: EXCEL_MAX_ROWS - 1, r2: MAX_ROWS - 1, c1: 0, c2: 0 });
  assert.equal(crossing.h, view.rows.size(EXCEL_MAX_ROWS - 1));
  view.ensureVisible(MAX_ROWS - 1, 0);
  assert.equal(view.extR, EXCEL_MAX_ROWS); assert.ok(view.sy <= view.maxScroll().y);
});

test('표시 한도를 줄여도 확장 영역의 원래 인쇄영역과 행 높이를 보존한다', () => {
  const row = EXCEL_MAX_ROWS + 4;
  const page = { area: { r1: row, c1: 0, r2: row + 1, c2: 1 } };
  const { view, wb, state } = grid(MAX_ROWS, { page, rowHeights: { [row]: 55 } });
  const before = view.printLayout(), original = JSON.stringify(wb.sheets[0].page);
  state.rowLimit = EXCEL_MAX_ROWS; view.refreshAxes();
  const after = view.printLayout();
  assert.equal(before.error, undefined); assert.equal(after.error, undefined);
  assert.deepEqual(after.areas.map(x => x.area), before.areas.map(x => x.area));
  assert.equal(after.areas[0].area.r1, row);
  assert.equal(after.areas[0].h, before.areas[0].h);
  assert.equal(JSON.stringify(wb.sheets[0].page), original);
});

test('접근성 좌표는 현재 모드와 실제 Axis.max 양쪽 한도를 지키며 거대 영역을 순회하지 않는다', () => {
  let calls = 0;
  const rows = { max: MAX_ROWS, size: () => { calls++; return 20; } }, cols = { max: MAX_COLS, size: () => { calls++; return 80; } };
  const windows = [{ r1: EXCEL_MAX_ROWS - 2, r2: MAX_ROWS - 1, c1: 0, c2: MAX_COLS - 1 }];
  const points = gridCoordinates({ r: MAX_ROWS - 1, c: 0 }, windows, rows, cols, 400, EXCEL_MAX_ROWS);
  assert.equal(points[0].r, EXCEL_MAX_ROWS - 1); assert.ok(points.every(p => p.r < EXCEL_MAX_ROWS)); assert.ok(calls < 4000);
  rows.max = 9; cols.max = 4;
  const small = gridCoordinates({ r: 100, c: 100 }, [{ r1: 0, r2: MAX_ROWS, c1: 0, c2: MAX_COLS }], rows, cols, 400, MAX_ROWS);
  assert.ok(small.every(p => p.r < 9 && p.c < 4)); assert.equal(small.length, 36);
  assert.equal(selectionDescription({ rowLimit: EXCEL_MAX_ROWS, sel: { r1: 0, c1: 0, r2: EXCEL_MAX_ROWS - 1, c2: MAX_COLS - 1 } }), '워크시트 전체 선택');
});

function documentStub() {
  const doc = { createElement: () => ({ ownerDocument: doc, attrs: new Map(), dataset: {}, children: [], textContent: '',
    setAttribute(k, v) { this.attrs.set(k, v); }, getAttribute(k) { return this.attrs.get(k) ?? null; }, removeAttribute(k) { this.attrs.delete(k); },
    addEventListener() {}, append(...nodes) { this.children.push(...nodes); }, replaceChildren(...nodes) { this.children = nodes; } }) };
  return doc;
}

test('접근성 rowcount·노드·병합 rowspan은 옵션 전환만으로 즉시 갱신되고 모델은 유지된다', () => {
  const merge = { r1: EXCEL_MAX_ROWS - 2, c1: 0, r2: EXCEL_MAX_ROWS + 3, c2: 0 };
  const { view, state, wb } = grid(MAX_ROWS, { merges: [merge] });
  const doc = documentStub(), editor = doc.createElement('textarea'); view.wrap = doc.createElement('div');
  state.active = { r: EXCEL_MAX_ROWS - 2, c: 0 }; state.sel = { ...merge };
  view.panes = [{ win: { r1: EXCEL_MAX_ROWS - 2, r2: EXCEL_MAX_ROWS + 3, c1: 0, c2: 1 } }];
  const a11y = new GridAccessibility(view, editor); view.a11y = a11y; a11y.update();
  const key = `${merge.r1},0`;
  assert.equal(a11y.grid.getAttribute('aria-rowcount'), String(MAX_ROWS));
  assert.equal(a11y.nodes.get(key).getAttribute('aria-rowspan'), '6');
  state.rowLimit = EXCEL_MAX_ROWS; view.refreshAxes();
  assert.equal(a11y.grid.getAttribute('aria-rowcount'), String(EXCEL_MAX_ROWS));
  a11y.update();
  assert.equal(a11y.nodes.get(key).getAttribute('aria-rowspan'), '2');
  assert.ok([...a11y.nodes.values()].every(n => Number(n.dataset.r) < EXCEL_MAX_ROWS));
  assert.deepEqual(wb.sheets[0].merges[0], merge);
  state.rowLimit = MAX_ROWS; a11y.update();
  assert.equal(a11y.nodes.get(key).getAttribute('aria-rowspan'), '2');
  view.refreshAxes(); a11y.update();
  assert.equal(a11y.grid.getAttribute('aria-rowcount'), String(MAX_ROWS));
  assert.equal(a11y.nodes.get(key).getAttribute('aria-rowspan'), '6');
});
