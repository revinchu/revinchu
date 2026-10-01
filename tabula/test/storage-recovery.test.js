import test from 'node:test';
import assert from 'node:assert/strict';
import { prepareBackupImport, server } from '../src/storage.js';

const backup = (documents) => ({ format: 'wixel-vault-backup', version: 1, documents });
test('백업 복원 계획은 대상 버전만 사용하고 중복·누락·초과 입력을 거부한다', () => {
  const input = backup([{ name: '합성', revision: 900, data: { sheets: [] } }]);
  assert.equal(prepareBackupImport(input).documents[0].expectedRevision, 0);
  assert.equal(prepareBackupImport(input, { overwrite: true, expectedRevisions: [{ name: '합성', revision: 8 }] }).documents[0].expectedRevision, 8);
  for (const bad of [{}, backup([]), backup([{ name: 'a', data: {} }, { name: ' a ', data: {} }]), backup([{ name: 'bad\u0000', data: {} }]), backup([{ name: 'a' }])]) assert.throws(() => prepareBackupImport(bad));
  assert.throws(() => prepareBackupImport(input, {}, { maxDocumentBytes: 2 }), /한도/);
  const hostile = JSON.parse('{"format":"wixel-vault-backup","version":1,"documents":[{"name":"__proto__","data":{"__proto__":{"polluted":true}}}]}');
  assert.equal(prepareBackupImport(hostile).documents[0].name, '__proto__'); assert.equal({}.polluted, undefined);
});

function connected() {
  server.vault = true; server.available = true; server.capabilities = { versionHistory: true, backupImport: true };
  server.connect('A'.repeat(43));
}
const json = (data, status = 200, headers = {}) => new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json', ...headers } });
test('온라인 과거본 조회는 현재 CAS를 유지하고 명시한 현재 버전으로 복원한다', async () => {
  const original = globalThis.fetch; connected(); server.restoreRevision('files/a', 8);
  globalThis.fetch = async (path, opts) => {
    if (path.startsWith('api/versions?')) return json({ currentRevision: 8, versions: [{ revision: 5 }] });
    if (opts.method === 'POST') { assert.equal(opts.headers['If-Match'], '"8"'); return json({ ok: true, revision: 9 }); }
    return json({ older: true }, 200, { ETag: '"5"', 'X-Wixel-Revision': '5' });
  };
  try {
    assert.equal((await server.versions('a')).currentRevision, 8);
    assert.deepEqual(await server.loadVersion('a', 5), { older: true }); assert.equal(server.revision('files/a'), '8');
    await server.restoreVersion('a', 5, 8); assert.equal(server.revision('files/a'), '9');
  } finally { globalThis.fetch = original; server.disconnect(); }
});

test('백업 클라이언트는 전체 전송 이후 commit하며 취소·오류는 staging을 정리한다', async () => {
  const original = globalThis.fetch, requests = []; connected();
  const input = backup([{ name: 'a', data: { a: 1 } }, { name: 'b', data: { b: 2 } }]);
  globalThis.fetch = async (path, opts) => {
    requests.push([path, opts]);
    if (path === 'api/imports') {
      const plan = JSON.parse(opts.body); assert.equal(plan.mode, 'replace');
      assert.deepEqual(plan.documents, [{ name: 'a', expectedRevision: 2 }, { name: 'b', expectedRevision: 0 }]);
      return json({ id: 'i', expires: 9 }, 201);
    }
    if (path.endsWith('/commit')) return json({ ok: true, documents: [{ name: 'a', revision: 4 }, { name: 'b', revision: 5 }] });
    assert.equal(opts.headers['If-Match'], path.endsWith('/a') ? '"2"' : '"0"'); return json({ ok: true });
  };
  try {
    const progress = [];
    await server.importBackup(input, { overwrite: true, expectedRevisions: [{ name: 'a', revision: 2 }], onProgress: (...n) => progress.push(n) });
    assert.deepEqual(requests.map(([p]) => p), ['api/imports', 'api/imports/i/files/a', 'api/imports/i/files/b', 'api/imports/i/commit']);
    assert.deepEqual(progress, [[1, 2], [2, 2]]); assert.equal(server.revision('files/a'), '4');
    requests.length = 0;
    globalThis.fetch = async (path, opts) => { requests.push([path, opts]); if (path === 'api/imports') return json({ id: 'failed' }); if (opts.method === 'DELETE') return json({ ok: true }); return json({ error: '잘못된 JSON', code: 'INVALID_JSON' }, 400); };
    await assert.rejects(server.importBackup(input), { code: 'INVALID_JSON' });
    assert.deepEqual(requests.map(([p]) => p), ['api/imports', 'api/imports/failed/files/a', 'api/imports/failed']);
  } finally { globalThis.fetch = original; server.disconnect(); }
});

test('복원 중 연결 전환과 미지원 Node 서버는 새 보관함 쓰기를 발생시키지 않는다', async () => {
  const original = globalThis.fetch, paths = []; connected();
  globalThis.fetch = async (path) => { paths.push(path); if (path === 'api/imports') return json({ id: 'i' }); server.connect('B'.repeat(42) + 'A'); return json({ ok: true }); };
  try {
    await assert.rejects(server.importBackup(backup([{ name: 'a', data: 1 }, { name: 'b', data: 2 }])), { code: 'CONNECTION_CHANGED' });
    assert.deepEqual(paths, ['api/imports', 'api/imports/i/files/a']);
    server.vault = false; server.capabilities = {}; paths.length = 0;
    assert.throws(() => server.versions('a'), { code: 'UNSUPPORTED_CAPABILITY' });
    await assert.rejects(server.importBackup(backup([{ name: 'a', data: 1 }])), { code: 'UNSUPPORTED_CAPABILITY' }); assert.deepEqual(paths, []);
  } finally { globalThis.fetch = original; server.disconnect(); }
});
