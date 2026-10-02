import test from 'node:test';
import assert from 'node:assert/strict';
import { Axis } from '../src/axis.js';
import { Workbook } from '../src/workbook.js';
import { computePrintLayout, normalizePrintAreas, normalizePageBreaks, printPageGeometry } from '../src/print-layout.js';
import { pageXml, pageFromXml } from '../src/page.js';
import { parseXml, descendants } from '../src/xml.js';

const page = { paper: 1, margins: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 }, scale: 100 };
const area = { r1: 0, c1: 0, r2: 21, c2: 17 };
const layout = (extra = {}, axes = {}) => computePrintLayout({ page: { ...page, area, ...extra }, rows: axes.rows ?? new Axis(100, {}, [], 100), cols: axes.cols ?? new Axis(100, {}, [], 100), ...axes });

test('print areas preserve ordered finite ranges and clip only full axes to used bounds', () => {
  const a = { r1: 3, c1: 2, r2: 8, c2: 6 }, b = { r1: 20, c1: 10, r2: 29, c2: 15 };
  assert.deepEqual(normalizePrintAreas({ area: b, areas: [a, b, { ...a }] }, a), [a, b]);
  assert.deepEqual(normalizePrintAreas({ area: { r1: 0, c1: 2, r2: 1048575, c2: 4 } }, a), [{ r1: 0, c1: 2, r2: 8, c2: 4 }]);
  assert.deepEqual(normalizePrintAreas({ area: { r1: 5, c1: 0, r2: 6, c2: 16383 } }, a), [{ r1: 5, c1: 0, r2: 6, c2: 6 }]);
  assert.deepEqual(normalizePrintAreas({}, a), [a]);
  assert.throws(() => normalizePrintAreas({ areas: [{ r1: NaN }] }, a), /범위/);
  assert.throws(() => normalizePrintAreas({ areas: Array(129).fill(a) }, a), /128/);
  assert.deepEqual(normalizePageBreaks([0, 2, 2, 7, -1, 1.5, 10], 10), [2, 7]);
});

test('paper margins reserve header/footer inside margins, not twice from the body', () => {
  const g = printPageGeometry({ ...page, margins: { top: 1, bottom: .5, left: .25, right: .75 } });
  assert.equal(g.paper.w, 816); assert.equal(g.paper.h, 1056);
  assert.equal(g.content.w, 720); assert.equal(g.content.h, 912);
  const p = layout({ area: { r1: 0, c1: 0, r2: 9, c2: 7 } });
  assert.equal(p.error, undefined); assert.equal(p.pages.length, 1);
});

test('automatic page rectangles and page order match known physical row/column capacity', () => {
  const p = layout(); assert.equal(p.pages.length, 9);
  assert.deepEqual(p.areas[0].rowBands.map(b => [b.start, b.end]), [[0, 9], [10, 19], [20, 21]]);
  assert.deepEqual(p.areas[0].colBands.map(b => [b.start, b.end]), [[0, 7], [8, 15], [16, 17]]);
  assert.deepEqual(p.pages.slice(0, 4).map(p => [p.r1, p.c1]), [[0, 0], [10, 0], [20, 0], [0, 8]]);
  assert.deepEqual(layout({ order: 'overThenDown' }).pages.slice(0, 4).map(p => [p.r1, p.c1]), [[0, 0], [0, 8], [0, 16], [10, 0]]);
  assert.ok(p.breaks.every(b => !b.manual));
});

test('manual boundaries include hidden runs without duplicates; custom sizes affect automatic breaks', () => {
  const rows = new Axis(100, { 0: 200 }, [{ 5: true, 6: true, 7: true }], 100);
  const p = layout({ rowBreaks: [5, 15], colBreaks: [3] }, { rows });
  assert.deepEqual(p.areas[0].rowBands.map(b => [b.start, b.end, b.bodyStart, b.manual]), [[0, 4, 0, false], [5, 14, 8, true], [15, 21, 15, true]]);
  assert.deepEqual(p.breaks.filter(b => b.axis === 'row').map(b => [b.index, b.position, b.manual]), [[5, 600, true], [15, 1300, true]]);
  assert.equal(layout({}, { rows: new Axis(100, { 0: 200 }, [], 100) }).areas[0].rowBands[0].end, 8);
});

test('repeat rows and columns reserve physical space and never duplicate body cells', () => {
  const p = layout({ titleRows: [0, 0], titleCols: [0, 0] });
  assert.deepEqual(p.areas[0].rowBands.map(b => [b.bodyStart, b.bodyEnd]), [[1, 9], [10, 18], [19, 21]]);
  assert.deepEqual(p.areas[0].colBands.map(b => [b.bodyStart, b.bodyEnd]), [[1, 7], [8, 14], [15, 17]]);
  const outside = layout({ area: { r1: 10, c1: 10, r2: 25, c2: 20 }, titleRows: [0, 1], titleCols: [0, 1] });
  assert.equal(outside.pages.length, 4); assert.equal(outside.pages[0].bodyR2, 17); assert.equal(outside.pages[0].bodyC2, 15);
  assert.equal(layout({ area: { r1: 0, c1: 0, r2: 0, c2: 0 }, titleRows: [0, 0], titleCols: [0, 0] }).pages.length, 1);
});

test('fit modes ignore manual breaks and preserve same scale across independent areas', () => {
  const p = layout({ fitW: 1, fitH: 1, rowBreaks: [5], colBreaks: [3] });
  assert.equal(p.manualIgnored, true); assert.equal(p.pages.length, 1); assert.ok(p.scale < 1);
  const multi = layout({ areas: [area, { r1: 30, c1: 30, r2: 31, c2: 31 }], fitW: 1, fitH: 1 });
  assert.equal(multi.pages.length, 2); assert.deepEqual(multi.pages.map(p => p.areaIndex), [0, 1]);
  assert.equal(multi.areas[0].scale, multi.areas[1].scale);
  const repeats = layout({ fitW: 2, fitH: 2, titleRows: [0, 3], titleCols: [0, 2] });
  assert.ok(repeats.areas[0].rowBands.length <= 2); assert.ok(repeats.areas[0].colBands.length <= 2);
});

test('hidden dimensions have no print pages, and huge areas fail explicitly without full-axis iteration', () => {
  const hidden = layout({}, { rows: new Axis(0, {}, [], 100) }); assert.equal(hidden.pages.length, 0);
  const rows = new Axis(22, {}, [], 20000000), original = rows.pos.bind(rows); let calls = 0;
  rows.pos = n => { calls++; return original(n); };
  const p = computePrintLayout({ page: { ...page, area: { r1: 0, c1: 0, r2: 19999999, c2: 1 } }, rows, cols: new Axis(64, {}, [], 16384), maxPages: 10 });
  assert.match(p.error, /10쪽/); assert.equal(p.pages.length, 0); assert.equal(p.breaks.length, 0);
  assert.equal(p.areas.length, 1); assert.equal(p.areas[0].h, 440000000); assert.ok(calls < 1500, String(calls));
});

test('row/column inserts shift print areas and manual breaks atomically with Undo/Redo', () => {
  const original = { areas: [{ r1: 2, c1: 2, r2: 8, c2: 8 }, { r1: 12, c1: 10, r2: 15, c2: 13 }], rowBreaks: [5, 12], colBreaks: [4, 10], titleRows: [0, 1], titleCols: [0, 1] };
  original.area = { ...original.areas[0] };
  const wb = new Workbook({ sheets: [{ name: '출력', cells: {}, page: original, view: { mode: 'pageBreakPreview' } }] });
  wb.transact(() => { wb.shiftAxis(0, 'row', 4, 2); wb.shiftAxis(0, 'column', 3, 1); });
  const pg = wb.sheets[0].page;
  assert.deepEqual(pg.areas, [{ r1: 2, c1: 2, r2: 10, c2: 9 }, { r1: 14, c1: 11, r2: 17, c2: 14 }]);
  assert.deepEqual(pg.area, pg.areas[0]); assert.deepEqual(pg.rowBreaks, [7, 14]); assert.deepEqual(pg.colBreaks, [5, 11]);
  const after = structuredClone(pg); wb.undo(); assert.deepEqual(wb.sheets[0].page, original); wb.redo(); assert.deepEqual(wb.sheets[0].page, after);
  const restored = new Workbook(wb.serialize()); assert.deepEqual(restored.sheets[0].page, after); assert.equal(restored.sheets[0].view.mode, 'pageBreakPreview');
  wb.transact(() => wb.shiftAxis(0, 'row', 6, -9)); assert.deepEqual(wb.sheets[0].page.rowBreaks, [6]);
});

test('page order round-trips standard pageSetup XML', () => {
  const raw = pageXml({ order: 'overThenDown' }, s => s), root = parseXml(`<root>${raw.setup}</root>`);
  assert.equal(pageFromXml({ pageSetup: descendants(root, 'pageSetup')[0] }).order, 'overThenDown');
});
