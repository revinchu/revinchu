import test from 'node:test';
import assert from 'node:assert/strict';
import { GridView } from '../src/view.js';

function harness(previous, current, si = 0) {
  const oldSheet = previous.sheets[0];
  const view = {
    host: { state: () => ({ wb: current, si }) }, z: 1,
    _objectRenderState: { wb: previous, si: 0, sheet: oldSheet, version: previous.version },
    _objectRenderCache: new Map([[oldSheet, { html: '<svg></svg>' }]]),
    _slicerVirtualModels: new Map([['previous', { items: ['old'] }]]),
  };
  const pane = { win: null, objHtml: new Map(), objects: { replaceChildren() { this.cleared = true; } } };
  return { view, pane };
}
const sheet = (shapes = []) => ({ charts: [], shapes });

test('renderObjects releases old document caches even when the next sheet has no objects', () => {
  const previous = { version: 0, sheets: [sheet([{ id: 'shape' }])] };
  const current = { version: 0, sheets: [sheet()] };
  const { view, pane } = harness(previous, current);
  GridView.prototype.renderObjects.call(view, pane);
  assert.equal(view._objectRenderState.wb, current);
  assert.equal(view._objectRenderState.sheet, current.sheets[0]);
  assert.equal(view._objectRenderCache.size, 0);
  assert.equal(view._slicerVirtualModels.size, 0);
  assert.equal(pane.objHtml, null);
  assert.equal(pane.objects.cleared, true);
});

test('renderObjects clears old sheet caches when switching to an empty sheet in the same document', () => {
  const book = { version: 2, sheets: [sheet([{ id: 'shape' }]), sheet()] };
  const { view, pane } = harness(book, book, 1);
  GridView.prototype.renderObjects.call(view, pane);
  assert.equal(view._objectRenderState.si, 1);
  assert.equal(view._objectRenderState.sheet, book.sheets[1]);
  assert.equal(view._objectRenderCache.size, 0);
  assert.equal(view._slicerVirtualModels.size, 0);
});

test('renderObjects releases the old document before returning for an unmeasured pane', () => {
  const previous = { version: 1, sheets: [sheet([{ id: 'old' }])] };
  const current = { version: 1, sheets: [sheet([{ id: 'new' }])] };
  const { view, pane } = harness(previous, current);
  GridView.prototype.renderObjects.call(view, pane);
  assert.equal(view._objectRenderState.wb, current);
  assert.equal(view._objectRenderCache.size, 0);
  assert.equal(view._slicerVirtualModels.size, 0);
});
