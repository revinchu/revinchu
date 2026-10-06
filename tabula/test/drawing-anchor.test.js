import test from 'node:test';
import assert from 'node:assert/strict';
import { Axis } from '../src/axis.js';
import { captureDrawingAnchors, reflowDrawingAnchors } from '../src/drawing-anchor.js';
import { Workbook } from '../src/workbook.js';
const cols = (sizes, hidden = {}) => new Axis(80, sizes, [hidden], 100);
const rows = (sizes, hidden = {}) => new Axis(20, sizes, [hidden], 100);
const object = (id, placement) => ({ id, x: 160.125, y: 40.375, w: 160.25, h: 40.5, placement });
const sheet = () => ({ slicers: [object('move'), object('resize', 'twoCell'), object('fixed', 'absolute')] });
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-7, `${a} ≈ ${b}`);

test('Excel placement follows fitted rows: move, move and size, or absolute', () => {
  const s = sheet(), anchors = captureDrawingAnchors(s, cols(), rows());
  const out = reflowDrawingAnchors(s, anchors, cols(), rows({ 0: 35, 3: 30 })).slicers;
  near(out[0].y, 55.375); near(out[0].h, 40.5);
  near(out[1].y, 55.375); near(out[1].h, 50.5);
  assert.equal(out[2], s.slicers[2]);
  assert.equal(s.slicers[0].y, 40.375);
});

test('cell anchors preserve the clearance from a neighboring table when columns change', () => {
  const s = sheet(), anchors = captureDrawingAnchors(s, cols(), rows());
  const nextCols = cols({ 0: 110, 3: 125 });
  const out = reflowDrawingAnchors(s, anchors, nextCols, rows()).slicers;
  near(out[0].x - nextCols.pos(2), .125); near(out[0].w, 160.25);
  near(out[1].w, 205.25); assert.equal(out[2].x, 160.125);
});

test('fractional EMU coordinates survive unrelated resizing and repeated round trips', () => {
  const s = sheet(), original = structuredClone(s);
  for (let i = 0; i < 25; i++) {
    let anchors = captureDrawingAnchors(s, cols(), rows());
    Object.assign(s, reflowDrawingAnchors(s, anchors, cols({ 50: 200 }), rows({ 50: 50 })));
    assert.deepEqual(s, original);
    anchors = captureDrawingAnchors(s, cols(), rows());
    Object.assign(s, reflowDrawingAnchors(s, anchors, cols({ 0: 120 }), rows({ 0: 35 })));
    anchors = captureDrawingAnchors(s, cols({ 0: 120 }), rows({ 0: 35 }));
    Object.assign(s, reflowDrawingAnchors(s, anchors, cols(), rows()));
    assert.deepEqual(s, original);
  }
});

test('insert and delete before anchored slicers retain offsets and placement', () => {
  const s = sheet(), anchors = captureDrawingAnchors(s, cols(), rows());
  const inserted = reflowDrawingAnchors(s, anchors, cols(), rows(), { shift: { axis: 'row', index: 1, count: 2 } }).slicers;
  near(inserted[0].y, 80.375); near(inserted[1].h, 40.5);
  assert.equal(inserted[2], s.slicers[2]);
  const next = { slicers: inserted };
  const restored = reflowDrawingAnchors(next, captureDrawingAnchors(next, cols(), rows()), cols(), rows(), { shift: { axis: 'row', index: 1, count: -2 } });
  assert.deepEqual(restored.slicers, s.slicers);
});

test('hidden rows above the drawing move it with the visible cells', () => {
  const s = sheet(), anchors = captureDrawingAnchors(s, cols(), rows());
  const out = reflowDrawingAnchors(s, anchors, cols(), rows({}, { 0: true })).slicers;
  near(out[0].y, 20.375); near(out[1].y, 20.375);
  assert.equal(out[2], s.slicers[2]);
});

test('dimension and drawing edits share one undo/redo transaction', () => {
  const wb = new Workbook(), s = wb.sheets[0];
  s.slicers = sheet().slicers;
  const original = structuredClone(s.slicers);
  wb.transact(() => {
    const anchors = captureDrawingAnchors(s, cols(), rows());
    wb.setSheetProp(0, 'rowHeights', { 0: 35 });
    for (const [prop, value] of Object.entries(reflowDrawingAnchors(s, anchors, cols(), rows(s.rowHeights)))) wb.setSheetProp(0, prop, value);
  });
  assert.equal(wb.undoStack.length, 1); near(s.slicers[0].y, 55.375);
  wb.undo(); assert.deepEqual(s.slicers, original); assert.deepEqual(s.rowHeights, {});
  wb.redo(); near(s.slicers[0].y, 55.375);
});
const reflow = (s, beforeCols, beforeRows, afterCols, afterRows, options = {}) => {
  Object.assign(s, reflowDrawingAnchors(s, captureDrawingAnchors(s, beforeCols, beforeRows), afterCols, afterRows, options));
};

test('hiding and showing the anchor cell restores fractional slicer bounds repeatedly', () => {
  const s = sheet(), original = JSON.parse(JSON.stringify(s)), hiddenCols = cols({}, { 2: true }), hiddenRows = rows({}, { 2: true });
  for (let i = 0; i < 25; i++) {
    reflow(s, cols(), rows(), hiddenCols, hiddenRows, { moveOnly: true });
    for (const obj of s.slicers.slice(0, 2)) {
      assert.deepEqual(obj._hiddenCellAnchor.markers.c1, [2, .125]);
      assert.deepEqual(obj._hiddenCellAnchor.markers.r1, [2, .375]);
    }
    // Plain metadata must retain its meaning without object identity or a cache.
    Object.assign(s, JSON.parse(JSON.stringify(s)));
    reflow(s, hiddenCols, hiddenRows, cols(), rows(), { moveOnly: true });
    assert.deepEqual(s, original);
  }
});

test('two-cell bottom-right markers survive hiding their own row and column', () => {
  const s = { shapes: [object('two', 'twoCell')] }, original = structuredClone(s);
  const hiddenCols = cols({}, { 4: true }), hiddenRows = rows({}, { 4: true });
  reflow(s, cols(), rows(), hiddenCols, hiddenRows);
  assert.deepEqual(s.shapes[0]._hiddenCellAnchor.markers.c2, [4, .375]);
  assert.deepEqual(s.shapes[0]._hiddenCellAnchor.markers.r2, [4, .875]);
  reflow(s, hiddenCols, hiddenRows, cols(), rows());
  assert.deepEqual(s, original);
});

test('hidden anchor metadata follows insert/delete remapping before showing the cell', () => {
  const s = { slicers: [object('move')] }, original = structuredClone(s);
  const hidden = rows({}, { 2: true }), insertedHidden = rows({}, { 4: true });
  reflow(s, cols(), rows(), cols(), hidden, { moveOnly: true });
  reflow(s, cols(), hidden, cols(), insertedHidden, { shift: { axis: 'row', index: 1, count: 2 } });
  assert.deepEqual(s.slicers[0]._hiddenCellAnchor.markers.r1, [4, .375]);
  near(s.slicers[0].y, 80.375);
  reflow(s, cols(), insertedHidden, cols(), hidden, { shift: { axis: 'row', index: 1, count: -2 } });
  assert.deepEqual(s.slicers[0]._hiddenCellAnchor.markers.r1, [2, .375]);
  reflow(s, cols(), hidden, cols(), rows(), { moveOnly: true });
  assert.deepEqual(s, original);
});

test('moving a drawing while its old anchor cell is hidden invalidates saved markers', () => {
  const s = { slicers: [object('move')] }, hidden = rows({}, { 2: true });
  reflow(s, cols(), rows(), cols(), hidden, { moveOnly: true });
  s.slicers = s.slicers.map(obj => ({ ...obj, y: obj.y + 40 }));
  const captured = captureDrawingAnchors(s, cols(), hidden);
  assert.deepEqual(captured.get('slicers:move').r1, [5, .375]);
  reflow(s, cols(), hidden, cols(), rows(), { moveOnly: true });
  near(s.slicers[0].y, 100.375);
  assert.equal(s.slicers[0]._hiddenCellAnchor, undefined);
});

test('resizing a drawing invalidates stale hidden geometry instead of reusing markers', () => {
  const s = { slicers: [object('move')] }, hidden = rows({}, { 2: true });
  reflow(s, cols(), rows(), cols(), hidden, { moveOnly: true });
  s.slicers = s.slicers.map(obj => ({ ...obj, h: obj.h - 10 }));
  const captured = captureDrawingAnchors(s, cols(), hidden);
  assert.deepEqual(captured.get('slicers:move').r1, [3, .375]);
  reflow(s, cols(), hidden, cols(), rows(), { moveOnly: true });
  near(s.slicers[0].h, 30.5);
  assert.equal(s.slicers[0]._hiddenCellAnchor, undefined);
});

test('hide undo redo show restores bounds despite structuredClone snapshots', () => {
  const wb = new Workbook(), s = wb.sheets[0];
  s.slicers = sheet().slicers;
  const original = structuredClone(s.slicers);
  const changeHidden = hidden => wb.transact(() => {
    const beforeRows = rows({}, s.hiddenRows), anchors = captureDrawingAnchors(s, cols(), beforeRows);
    wb.setSheetProp(0, 'hiddenRows', hidden);
    const changed = reflowDrawingAnchors(s, anchors, cols(), rows({}, s.hiddenRows), { moveOnly: true });
    for (const [prop, value] of Object.entries(changed)) wb.setSheetProp(0, prop, value);
  });
  changeHidden({ 2: true });
  assert.equal(wb.undoStack.length, 1);
  assert.deepEqual(s.slicers[0]._hiddenCellAnchor.markers.r1, [2, .375]);
  wb.undo(); assert.deepEqual(s.slicers, original);
  wb.redo(); assert.deepEqual(s.slicers[0]._hiddenCellAnchor.markers.r1, [2, .375]);
  changeHidden({});
  assert.deepEqual(s.slicers, original);
  wb.undo(); assert.deepEqual(s.slicers[0]._hiddenCellAnchor.markers.r1, [2, .375]);
  wb.redo(); assert.deepEqual(s.slicers, original);
});

test('hidden metadata changes are immutable even when drawing bounds do not change', () => {
  const s = { slicers: [object('move')] }, old = s.slicers[0], hidden = rows({}, { 2: true });
  reflow(s, cols(), rows(), cols(), hidden, { moveOnly: true });
  assert.notEqual(s.slicers[0], old);
  assert.equal(old._hiddenCellAnchor, undefined);
  assert.deepEqual({ x: s.slicers[0].x, y: s.slicers[0].y, w: s.slicers[0].w, h: s.slicers[0].h },
    { x: old.x, y: old.y, w: old.w, h: old.h });
  const once = s.slicers[0];
  reflow(s, cols(), hidden, cols(), hidden, { moveOnly: true });
  assert.equal(s.slicers[0], once);
});
