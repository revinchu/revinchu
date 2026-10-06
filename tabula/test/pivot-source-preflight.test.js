import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData } from '../src/pivot.js';

const saved = [['항목', '금액'], ['A', 1]];
const live = [['항목', '금액'], ['A', 10], ['B', 20]];
function fixture() {
  const def = { name: '보고서', source: '원본', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, snapshotId: 'saved', rows: ['항목'], cols: [], values: [{ field: '금액', agg: 'sum' }], top: 8, left: 4 };
  const wb = new Workbook({ pivotSnapshots: { saved }, sheets: [{
    name: '원본', cells: { '0,0': { raw: '항목' }, '0,1': { raw: '금액' }, '1,0': { raw: 'A' }, '1,1': { raw: '10' }, '2,0': { raw: 'B' }, '2,1': { raw: '20' } }, pivot: def,
  }] });
  return { wb, def: wb.sheets[0].pivot };
}
const history = wb => ({ undo: wb.undoStack.length, redo: wb.redoStack.length, version: wb.version });

test('사전 조회는 유효한 저장 원본을 쓰면서 최초 조회 메타데이터를 변경하지 않는다', () => {
  const { wb, def } = fixture();
  // 이전 확장 API처럼 아직 변경 추적이 등록되지 않은 저장 캐시.
  const snap = Object.freeze({ rows: saved });
  wb.pivotSnapshots.set('saved', snap);
  const before = history(wb), originalDef = structuredClone(def);
  for (let i = 0; i < 2; i++) assert.deepEqual(pivotSourceData(wb, def, { preserveSnapshot: true }).rows, saved);
  assert.equal(wb.pivotSnapshots.get('saved'), snap);
  assert.deepEqual(Object.keys(snap), ['rows']);
  assert.deepEqual(def, originalDef);
  assert.deepEqual(history(wb), before);
  // 옵션을 쓰지 않는 기존 조회는 변경 추적 등록 동작을 유지한다.
  const ordinary = { rows: saved };
  wb.pivotSnapshots.set('saved', ordinary);
  assert.deepEqual(pivotSourceData(wb, def).rows, saved);
  assert.ok(ordinary.sourceWatch);
  assert.equal(ordinary.sourceSheet, wb.sheets[0]);
  assert.equal(ordinary.ver, ordinary.sourceWatch.version);
});

test('사전 조회는 오래된 캐시 대신 수정한 원본을 읽되 캐시와 Undo 상태를 보존한다', () => {
  const { wb, def } = fixture(), snap = wb.pivotSnapshots.get('saved');
  wb.transact(() => wb.setInput(0, 1, 1, '15'));
  const before = history(wb), originalDef = structuredClone(def), snapshotKeys = Object.keys(snap), version = snap.ver, watch = snap.sourceWatch;
  const expected = [['항목', '금액'], ['A', 15], ['B', 20]];
  for (let i = 0; i < 2; i++) assert.deepEqual(pivotSourceData(wb, def, { preserveSnapshot: true }).rows, expected);
  assert.equal(wb.pivotSnapshots.get('saved'), snap);
  assert.deepEqual(Object.keys(snap), snapshotKeys);
  assert.equal(snap.rows, saved);
  assert.equal(snap.ver, version);
  assert.equal(snap.sourceWatch, watch);
  assert.deepEqual(def, originalDef);
  assert.deepEqual(history(wb), before);
  // 실제 조회의 기본 동작은 오래된 캐시를 계속 제거한다.
  assert.deepEqual(pivotSourceData(wb, def).rows, expected);
  assert.equal(wb.pivotSnapshots.has('saved'), false);
});

test('변경 추적 API가 없는 이전 원본 제공자도 사전 조회 시 캐시 버전과 등록을 보존한다', () => {
  const { wb, def } = fixture();
  wb.pivotSnapshotCurrent = undefined;
  const snap = Object.freeze({ rows: saved });
  wb.pivotSnapshots.set('saved', snap);
  assert.deepEqual(pivotSourceData(wb, def, { preserveSnapshot: true }).rows, saved);
  assert.equal(Object.hasOwn(snap, 'ver'), false);
  const stale = Object.freeze({ rows: saved, ver: -1 });
  wb.pivotSnapshots.set('saved', stale);
  assert.deepEqual(pivotSourceData(wb, def, { preserveSnapshot: true }).rows, live);
  assert.equal(wb.pivotSnapshots.get('saved'), stale);
  assert.equal(stale.ver, -1);
  assert.deepEqual(pivotSourceData(wb, def).rows, live);
  assert.equal(wb.pivotSnapshots.has('saved'), false);
});
