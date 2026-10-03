import test from 'node:test';
import assert from 'node:assert/strict';

// Small deterministic IDB fault harness; real IDB is also checked by
// tools/storage-connection.mjs in Chromium/WebKit isolated browser profiles.
function databaseHarness() {
  const data = new Map(), connections = [];
  let opens = 0, nextOpenError = null, syncOpenError = null, abortNext = null;
  const indexedDB = { open() {
    opens++;
    if (syncOpenError) { const error = syncOpenError; syncOpenError = null; throw error; }
    const request = {};
    queueMicrotask(() => {
      if (nextOpenError) { request.error = nextOpenError; nextOpenError = null; request.onerror?.(); return; }
      const db = { closed: false, transactions: 0, requests: 0, failTransaction: null,
        close() { this.closed = true; },
        transaction() {
          this.transactions++;
          if (this.closed) throw new DOMException('closed connection', 'InvalidStateError');
          if (this.failTransaction) throw this.failTransaction;
          const writes = new Map(), deletes = new Set();
          let pending = 0, finished = false;
          const finish = () => setImmediate(() => {
            if (finished || pending) return;
            if (abortNext) { tx.error = abortNext; abortNext = null; tx.abort(); return; }
            finished = true;
            for (const key of deletes) data.delete(key);
            for (const [key, value] of writes) data.set(key, value);
            tx.oncomplete?.();
          });
          const requestFor = action => {
            const request = {}; pending++; db.requests++;
            queueMicrotask(() => {
              if (finished) { request.error = new DOMException('aborted', 'AbortError'); request.onerror?.(); pending--; return; }
              request.result = action(); request.onsuccess?.(); pending--; finish();
            });
            return request;
          };
          const tx = { error: null,
            abort() { if (finished) return; finished = true; queueMicrotask(() => tx.onabort?.()); },
            objectStore() { return {
              get: key => requestFor(() => writes.has(key) ? writes.get(key) : deletes.has(key) ? undefined : data.get(key)),
              put: (value, key) => requestFor(() => { writes.set(key, structuredClone(value)); deletes.delete(key); return key; }),
              delete: key => requestFor(() => { writes.delete(key); deletes.add(key); }),
              getAllKeys: range => requestFor(() => [...data.keys()].filter(key => key >= range.lower && key <= range.upper)),
            }; },
          };
          finish(); return tx;
        },
      };
      connections.push(db); request.result = db; request.onsuccess?.();
    });
    return request;
  } };
  return { indexedDB, data, connections, get opens() { return opens; },
    failOpen(error, sync = false) { if (sync) syncOpenError = error; else nextOpenError = error; },
    abort(error) { abortNext = error; },
  };
}
let importId = 0;
async function fixture(fn) {
  const previous = globalThis.indexedDB, previousRange = globalThis.IDBKeyRange, harness = databaseHarness();
  globalThis.indexedDB = harness.indexedDB;
  globalThis.IDBKeyRange = { bound: (lower, upper) => ({ lower, upper }) };
  try {
    const storage = await import(`../src/storage.js?connection-test=${++importId}`);
    await fn(storage, harness);
  } finally { globalThis.indexedDB = previous; globalThis.IDBKeyRange = previousRange; }
}

test('IndexedDB 일시 열기 실패 뒤 다음 저장은 새 연결로 성공한다', () => fixture(async (s, h) => {
  h.failOpen(new DOMException('temporary open failure', 'UnknownError'));
  await assert.rejects(s.idbSet('doc', 'new'), { name: 'UnknownError' });
  await s.idbSet('doc', 'new'); assert.equal(await s.idbGet('doc'), 'new'); assert.equal(h.opens, 2);
}));

test('IndexedDB open의 동기 예외도 다음 저장을 영구 차단하지 않는다', () => fixture(async (s, h) => {
  h.failOpen(new DOMException('temporary unavailable', 'InvalidStateError'), true);
  await assert.rejects(s.idbGet('doc'), { name: 'InvalidStateError' });
  await s.idbSet('doc', 2); assert.equal(await s.idbGet('doc'), 2); assert.equal(h.opens, 2);
}));

test('닫힌 연결의 동시 저장은 한 번 재연결하고 옛 close 이벤트가 새 연결을 버리지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('old', 1); const old = h.connections[0]; old.close();
  await Promise.all(Array.from({ length: 20 }, (_, i) => s.idbSet(`new:${i}`, i)));
  assert.equal(h.opens, 2); assert.equal(h.connections[1].requests, 20);
  old.onclose?.(); assert.equal(await s.idbGet('new:19'), 19); assert.equal(h.opens, 2);
}));

test('강제 close와 versionchange 이후 읽기·저장이 새 연결을 연다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1);
  h.connections[0].closed = true; h.connections[0].onclose?.();
  assert.equal(await s.idbGet('doc'), 1); assert.equal(h.opens, 2);
  h.connections[1].onversionchange?.(); assert.equal(h.connections[1].closed, true);
  await s.idbSet('doc', 2); assert.equal(await s.idbGet('doc'), 2); assert.equal(h.opens, 3);
}));

test('진행 중 quota 중단은 자동 재실행하지 않고 기존 값을 유지한다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 'old'); const db = h.connections[0], before = db.requests;
  h.abort(new DOMException('quota', 'QuotaExceededError'));
  await assert.rejects(s.idbSet('doc', 'new'), { name: 'QuotaExceededError' });
  assert.equal(db.requests, before + 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 'old');
  await s.idbSet('doc', 'new'); assert.equal(await s.idbGet('doc'), 'new');
}));

test('연결 종료가 아닌 트랜잭션 생성 오류는 자동 재시도하지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1); const db = h.connections[0], before = db.transactions;
  db.failTransaction = new DOMException('missing store', 'NotFoundError');
  await assert.rejects(s.idbSet('doc', 2), { name: 'NotFoundError' });
  assert.equal(db.transactions, before + 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 1);
}));

test('이미 시작한 저장 콜백의 InvalidStateError는 재연결·중복 실행하지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1); let calls = 0;
  await assert.rejects(s.idbUpdate(['doc'], () => { calls++; throw new DOMException('callback error', 'InvalidStateError'); }), { name: 'InvalidStateError' });
  assert.equal(calls, 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 1);
}));

test('각 저장 API가 종료된 연결을 복구하고 CAS 충돌 검사는 그대로 유지한다', () => fixture(async (s, h) => {
  const close = () => h.connections.at(-1).close();
  await s.idbSet('a', 1); close(); assert.equal(await s.idbGet('a'), 1);
  close(); await s.idbUpdate(['a'], values => ({ set: [['b', values.get('a') + 1]], result: true }));
  close(); assert.deepEqual((await s.idbKeys('')).sort(), ['a', 'b']);
  close(); await s.idbDel('a'); assert.equal(h.data.has('a'), false);
  close(); await s.idbCompareAndSet('b', 2, 3); assert.equal(h.data.get('b'), 3);
  close(); await assert.rejects(s.idbCompareAndSet('b', 2, 4), { code: 'IDB_CONFLICT' }); assert.equal(h.data.get('b'), 3);
  close(); await s.idbDeleteMany(['b']); assert.equal(h.data.size, 0); assert.equal(h.opens, 8);
}));

test('대형 문서의 열린 세대 검증은 뒤늦은 덮어쓰기를 청크 저장 전에 차단한다', () => fixture(async (_s, h) => {
  const { saveLargeWorkbook, loadLargeWorkbook } = await import('../src/big-storage.js');
  const { Workbook } = await import('../src/workbook.js');
  const key = 'synthetic:observed-generation', original = new Workbook({ sheets: [{ name: '동시 편집', cells: { '0,0': { raw: '1' } } }] });
  const first = await saveLargeWorkbook(key, original, {}, { expectedGeneration: null });
  const current = new Workbook((await loadLargeWorkbook(key)).workbook), stale = new Workbook((await loadLargeWorkbook(key)).workbook);
  current.transact(() => current.setInput(0, 0, 0, '101')); stale.transact(() => stale.setInput(0, 0, 0, '202'));
  const second = await saveLargeWorkbook(key, current, {}, { expectedGeneration: first.manifest.generation });
  const requests = h.connections[0].requests, keys = [...h.data.keys()];
  await assert.rejects(saveLargeWorkbook(key, stale, {}, { expectedGeneration: first.manifest.generation }), { code: 'IDB_CONFLICT' });
  assert.equal(h.connections[0].requests, requests + 1); // Only the manifest read; no staged records or GC.
  assert.deepEqual([...h.data.keys()], keys); assert.equal(h.data.get(key).generation, second.manifest.generation);
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 101);
  assert.equal(stale.getValue(0, 0, 0), 202);
  await assert.rejects(saveLargeWorkbook(key, stale, {}, { expectedGeneration: null }), { code: 'IDB_CONFLICT' });
  const forkKey = key + ':fork'; await saveLargeWorkbook(forkKey, stale, {}, { expectedGeneration: null });
  assert.equal(new Workbook((await loadLargeWorkbook(forkKey)).workbook).getValue(0, 0, 0), 202);
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 101);
  // Callers that have not opted into observed generations retain their prior API.
  stale.transact(() => stale.setInput(0, 0, 0, '203'));
  await saveLargeWorkbook(key, stale, {});
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 203);
}));
