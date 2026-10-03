import test from 'node:test';
import assert from 'node:assert/strict';
import { createDocumentOpenGate, isDocumentOpenCancelled } from '../src/document-session.js';

function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function openAfter(gate, pending, applied) {
  const token = gate.begin();
  return pending.then(value => {
    token.assertCurrent();
    applied.push(value);
    return value;
  });
}

test('a new open permanently supersedes the previous token', () => {
  const gate = createDocumentOpenGate();
  const first = gate.begin();
  assert.equal(first.isCurrent(), true);
  assert.doesNotThrow(() => first.assertCurrent());
  const second = gate.begin();
  assert.equal(first.isCurrent(), false);
  assert.equal(second.isCurrent(), true);
  assert.throws(() => first.assertCurrent(), error => {
    assert.equal(error.code, 'DOCUMENT_OPEN_CANCELLED');
    assert.match(error.message, /문서.*취소/);
    assert.equal(isDocumentOpenCancelled(error), true);
    return true;
  });
  assert.doesNotThrow(() => second.assertCurrent());
});

test('a slow first document cannot overwrite a faster second document', async () => {
  const gate = createDocumentOpenGate(), first = deferred(), second = deferred(), applied = [];
  const firstResult = openAfter(gate, first.promise, applied);
  const firstCancelled = assert.rejects(firstResult, isDocumentOpenCancelled);
  const secondResult = openAfter(gate, second.promise, applied);
  second.resolve('두 번째 문서');
  assert.equal(await secondResult, '두 번째 문서');
  first.resolve('첫 번째 문서');
  await firstCancelled;
  assert.deepEqual(applied, ['두 번째 문서']);
});

test('a failed newer request does not reactivate an older pending document', async () => {
  const gate = createDocumentOpenGate(), first = deferred(), second = deferred(), applied = [];
  const firstResult = openAfter(gate, first.promise, applied);
  const firstCancelled = assert.rejects(firstResult, isDocumentOpenCancelled);
  const secondResult = openAfter(gate, second.promise, applied);
  const failure = new Error('파일을 읽을 수 없습니다.');
  const secondFailed = assert.rejects(secondResult, error => error === failure);
  second.reject(failure);
  await secondFailed;
  first.resolve('첫 번째 문서');
  await firstCancelled;
  assert.deepEqual(applied, []);
});

test('invalidation cancels pending work and allows a fresh document request', async () => {
  const gate = createDocumentOpenGate(), pending = deferred(), applied = [];
  const result = openAfter(gate, pending.promise, applied);
  const cancelled = assert.rejects(result, isDocumentOpenCancelled);
  gate.invalidate();
  gate.invalidate();
  pending.resolve('닫힌 문서');
  await cancelled;
  assert.deepEqual(applied, []);
  const next = gate.begin();
  assert.equal(next.isCurrent(), true);
  assert.doesNotThrow(() => next.assertCurrent());
});

test('each async stage rechecks the same token after a newer request begins', async () => {
  const gate = createDocumentOpenGate(), parsed = deferred(), applied = [];
  const token = gate.begin();
  const work = (async () => {
    await Promise.resolve();
    token.assertCurrent();
    await parsed.promise;
    token.assertCurrent();
    applied.push('이전 문서');
  })();
  const cancelled = assert.rejects(work, isDocumentOpenCancelled);
  await Promise.resolve();
  assert.equal(token.isCurrent(), true);
  const next = gate.begin();
  parsed.resolve();
  await cancelled;
  assert.equal(token.isCurrent(), false);
  assert.equal(next.isCurrent(), true);
  assert.deepEqual(applied, []);
});

test('separate document gates do not cancel each other', () => {
  const firstGate = createDocumentOpenGate(), secondGate = createDocumentOpenGate();
  const first = firstGate.begin(), second = secondGate.begin();
  firstGate.invalidate();
  assert.equal(first.isCurrent(), false);
  assert.equal(second.isCurrent(), true);
  secondGate.begin();
  assert.equal(second.isCurrent(), false);
});

test('only the explicit cancellation code is classified as a superseded open', () => {
  for (const error of [null, undefined, false, 'DOCUMENT_OPEN_CANCELLED', new Error('취소'), { code: 'ABORT_ERR' }]) {
    assert.equal(isDocumentOpenCancelled(error), false);
  }
  assert.equal(isDocumentOpenCancelled({ code: 'DOCUMENT_OPEN_CANCELLED' }), true);
});
