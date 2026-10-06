import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotFieldItemModel } from '../src/pivot-field-items.js';

const saved = [['항목', '금액'], ['A', 1]];
function fixture() {
  const def = { source: '원본', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, snapshotId: 'saved', cacheItemsId: 'items', missingItems: 'none', rows: ['항목'], values: [{ field: '금액', agg: 'sum' }], top: 8, left: 4 };
  const wb = new Workbook({ pivotSnapshots: { saved }, pivotCacheItems: { items: { missingItemsLimit: 0, fields: [{ name: '항목', shared: ['A', 'B', '과거'], sharedTypes: 'sss' }] } }, sheets: [{
    name: '원본', cells: { '0,0': { raw: '항목' }, '0,1': { raw: '금액' }, '1,0': { raw: 'A' }, '1,1': { raw: '10' }, '2,0': { raw: 'B' }, '2,1': { raw: '20' } }, pivot: def,
  }] });
  return { wb, def: wb.sheets[0].pivot };
}
const state = wb => ({ undo: wb.undoStack.length, redo: wb.redoStack.length, version: wb.version });

test('필터 준비 목록의 과거 항목 판정도 최초 snapshot 메타데이터를 변경하지 않는다', () => {
  const { wb, def } = fixture(), snap = Object.freeze({ rows: saved });
  wb.pivotSnapshots.set('saved', snap);
  const before = state(wb), original = structuredClone(def);
  const model = pivotFieldItemModel(wb, def, '항목', { preserveSnapshot: true });
  assert.deepEqual(model.items.map(item => item.key), ['A', 'B', '과거']);
  assert.equal(model.items.find(item => item.key === 'B').deleted, true);
  assert.equal(wb.pivotSnapshots.get('saved'), snap);
  assert.deepEqual(Object.keys(snap), ['rows']);
  assert.deepEqual(def, original);
  assert.deepEqual(state(wb), before);
  // 일반 표시 조회는 이전 확장 API의 최초 변경 추적 등록을 그대로 수행한다.
  const ordinary = { rows: saved };
  wb.pivotSnapshots.set('saved', ordinary);
  assert.deepEqual(pivotFieldItemModel(wb, def, '항목').items.map(item => item.key), ['A', 'B', '과거']);
  assert.ok(ordinary.sourceWatch);
  assert.equal(ordinary.ver, ordinary.sourceWatch.version);
});

test('필터 준비 목록은 stale 캐시를 지우지 않고 현재 원본 항목으로 판단한다', () => {
  const { wb, def } = fixture(), snap = wb.pivotSnapshots.get('saved');
  wb.transact(() => wb.setInput(0, 1, 0, '새 항목'));
  const before = state(wb), original = structuredClone(def), version = snap.ver, watch = snap.sourceWatch;
  const model = pivotFieldItemModel(wb, def, '항목', { preserveSnapshot: true });
  assert.deepEqual(model.items.map(item => item.key), ['B', '새 항목']);
  assert.equal(wb.pivotSnapshots.get('saved'), snap);
  assert.equal(snap.ver, version);
  assert.equal(snap.sourceWatch, watch);
  assert.deepEqual(def, original);
  assert.deepEqual(state(wb), before);
  // 일반 표시 조회에서는 기존 stale snapshot 제거 동작을 유지한다.
  assert.deepEqual(pivotFieldItemModel(wb, def, '항목').items.map(item => item.key), ['B', '새 항목']);
  assert.equal(wb.pivotSnapshots.has('saved'), false);
});
