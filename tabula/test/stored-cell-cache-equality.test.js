import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';

const raw = '=UNKNOWN_SYNTHETIC()';
const fixture = data => new Workbook({calculation:{mode:'manual'},sheets:[{name:'Synthetic',cells:{'0,0':{raw,...data}}}]});
const stored = w => cellData(w.getCell(0,0,0));
function change(w, data) { w.transact(() => w.setCellData(0,0,0,{raw,...data})); }
function roundtripHistory(w, before, after) {
  assert.equal(w.undoStack.length,1);
  assert.deepEqual(stored(w),after);
  w.undo(); assert.deepEqual(stored(w),before);
  w.redo(); assert.deepEqual(stored(w),after);
}

test('changing retained stale scalar cache records the new value and exact Undo/Redo', () => {
  const w=fixture({staleCached:7}), before=stored(w);
  change(w,{staleCached:8});
  roundtripHistory(w,before,{raw,staleCached:8});
  assert.equal(stored(new Workbook(w.serialize())).staleCached,8);
});

for (const trustedFirst of [true,false]) test('scalar cache trust transition '+(trustedFirst?'valid to stale':'stale to valid')+' is stored', () => {
  const first=trustedFirst?'cached':'staleCached', last=trustedFirst?'staleCached':'cached';
  const w=fixture({[first]:null}), before=stored(w);
  w.transact(()=>{w.setCellData(0,0,0,{raw,[last]:null}); assert.deepEqual(stored(w),{raw,[last]:null});});
  assert.equal(w.undoStack.length,1);
  assert.deepEqual(w.undoStack[0].entries[0].before,before);
  assert.deepEqual(w.undoStack[0].entries[0].after,{raw,[last]:null});
  // End-of-transaction calculation invalidation may mark the changed formula stale.
  assert.equal(stored(w).staleCached,null);
  w.undo(); assert.equal(w.getCell(0,0,0).cached,null);
  w.redo(); assert.equal(w.getCell(0,0,0).cached,null);
  assert.equal(stored(new Workbook(w.serialize())).staleCached,null);
});

test('null retained cache differs from absence but undefined remains absent', () => {
  const w=fixture({staleCached:null});
  change(w,{}); roundtripHistory(w,{raw,staleCached:null},{raw});
  const same=fixture({cached:undefined}); change(same,{staleCached:undefined});
  assert.equal(same.undoStack.length,0); assert.deepEqual(stored(same),{raw});
});

test('scalar caches preserve signed zero and distinguish NaN from null', () => {
  const zero=fixture({staleCached:-0}); change(zero,{staleCached:0});
  assert.equal(zero.undoStack.length,1); assert.ok(Object.is(stored(zero).staleCached,0));
  zero.undo(); assert.ok(Object.is(stored(zero).staleCached,-0));
  zero.redo(); assert.ok(Object.is(stored(zero).staleCached,0));
  const nan=fixture({staleCached:NaN}); change(nan,{staleCached:null});
  assert.equal(nan.undoStack.length,1); assert.equal(stored(nan).staleCached,null);
  nan.undo(); assert.ok(Number.isNaN(stored(nan).staleCached));
  const same=fixture({staleCached:NaN}); change(same,{staleCached:NaN});
  assert.equal(same.undoStack.length,0);
});

test('equal error cache copies are a no-op and changed error codes are retained', () => {
  const w=fixture({staleCached:{error:'#N/A'}});
  change(w,{staleCached:{error:'#N/A'}}); assert.equal(w.undoStack.length,0);
  change(w,{staleCached:{error:'#VALUE!'}});
  roundtripHistory(w,{raw,staleCached:{error:'#N/A'}},{raw,staleCached:{error:'#VALUE!'}});
});

const array = value => ({h:1,w:2,values:[0,0,value,0,1,6]});
for (const trustedFirst of [true,false]) test('array-only cache trust transition '+(trustedFirst?'valid to stale':'stale to valid')+' is stored', () => {
  const first=trustedFirst?'cachedArray':'staleCachedArray', last=trustedFirst?'staleCachedArray':'cachedArray';
  const data=array(5), w=fixture({[first]:data}), before=stored(w);
  w.transact(()=>{w.setCellData(0,0,0,{raw,[last]:array(5)}); assert.deepEqual(stored(w),{raw,[last]:data});});
  assert.equal(w.undoStack.length,1);
  assert.deepEqual(w.undoStack[0].entries[0].before,before);
  assert.deepEqual(w.undoStack[0].entries[0].after,{raw,[last]:data});
  assert.deepEqual(w.getCell(0,0,0).cachedArray,data);
  w.undo(); assert.deepEqual(w.getCell(0,0,0).cachedArray,data);
  w.redo(); assert.deepEqual(w.getCell(0,0,0).cachedArray,data);
  assert.ok(stored(new Workbook(w.serialize())).staleCachedArray);
});

test('array cache value updates and signed zero remain distinct', () => {
  const w=fixture({staleCachedArray:array(5)}), before=stored(w);
  change(w,{staleCachedArray:array(8)});
  roundtripHistory(w,before,{raw,staleCachedArray:array(8)});
  const zero=fixture({staleCachedArray:array(-0)}); change(zero,{staleCachedArray:array(0)});
  assert.equal(zero.undoStack.length,1); zero.undo(); assert.ok(Object.is(stored(zero).staleCachedArray.values[2],-0));
});

test('formula dirty flag without retained scalar or array cache remains a no-op', () => {
  const w=fixture({}); w.getCell(0,0,0).dirty=true;
  change(w,{}); assert.equal(w.undoStack.length,0); assert.deepEqual(stored(w),{raw});
});

test('format change preserves stale cache trust and equal cache copies avoid extra history', () => {
  const w=fixture({staleCached:7,staleCachedArray:array(5)});
  w.transact(()=>w.setStyle(0,0,0,{bold:true}));
  assert.equal(w.undoStack.length,1); assert.equal(stored(w).staleCached,7);
  assert.ok(stored(w).staleCachedArray); assert.ok(!Object.hasOwn(stored(w),'cachedArray'));
  change(w,{style:{bold:true},staleCached:7,staleCachedArray:array(5)});
  assert.equal(w.undoStack.length,1);
});
