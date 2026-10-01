import { test } from 'node:test';
import assert from 'node:assert/strict';
import { server, createVaultKey, validVaultKey } from '../src/storage.js';

test('개인 보관함 클라이언트: 명시적 연결·버전 조건·충돌·연결 변경 격리', async () => {
  const realFetch = globalThis.fetch;
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const settings = new Map(), requests = [];
  Object.defineProperty(globalThis, 'localStorage', { configurable: true, value: { getItem: (k) => settings.get(k) ?? null, setItem: (k, v) => settings.set(k, v) } });
  let rev = 0, waitForSave = null;
  globalThis.fetch = async (url, options = {}) => {
    requests.push({ url, ...options });
    if (url === 'api/health') return Response.json({ ok: true, vault: true, auth: false });
    if (waitForSave) await waitForSave;
    if (options.method === 'PUT') {
      if (options.headers['If-Match'] !== `"${rev}"`) return Response.json({ error: '다른 기기에서 수정됨', code: 'REVISION_CONFLICT', currentRevision: rev }, { status: 412 });
      return Response.json({ ok: true, revision: ++rev, modified: 123 });
    }
    if (options.method === 'DELETE') return Response.json({ ok: true });
    return Response.json({ docName: '테스트', workbook: { sheets: [] } }, { headers: { 'ETag': `"${rev}"`, 'X-Wixel-Revision': String(rev) } });
  };
  try {
    server.disconnect(); await server.init();
    assert.equal(server.connected, false); assert.equal(requests.length, 1);
    const key = createVaultKey(); assert.equal(key.length, 43); assert.equal(validVaultKey(key), true); assert.equal(validVaultKey('잘못된 키'), false);
    assert.equal(validVaultKey(`${'A'.repeat(42)}B`), false, '비정규 base64url 마지막 문자는 거부');
    server.connect(key); assert.equal(server.connected, true); assert.equal(requests.length, 1, '연결은 문서를 업로드하지 않음');
    await server.save('테스트', {}); assert.equal(requests.at(-1).headers['If-Match'], '"0"'); assert.equal(requests.at(-1).headers['X-Wixel-Vault'], key);
    await server.save('테스트', {}); assert.equal(requests.at(-1).headers['If-Match'], '"1"');
    rev = 3; await assert.rejects(server.save('테스트', {}), { status: 412, code: 'REVISION_CONFLICT', currentRevision: 3 });
    assert.equal(server.revision('files/%ED%85%8C%EC%8A%A4%ED%8A%B8'), '2', '충돌 시 새 버전을 수락해 덮어쓰지 않음');
    await server.load('테스트'); assert.equal(server.revision('files/%ED%85%8C%EC%8A%A4%ED%8A%B8'), '3');
    await server.remove('테스트', 3); assert.equal(requests.at(-1).headers['If-Match'], '"3"');
    let release; waitForSave = new Promise((resolve) => { release = resolve; });
    const saving = server.save('테스트', {});
    server.connect(createVaultKey()); release();
    await assert.rejects(saving, { code: 'CONNECTION_CHANGED' });
    assert.equal(server.revision('files/%ED%85%8C%EC%8A%A4%ED%8A%B8'), '0');
    server.disconnect(); assert.equal(server.recoveryKey(), null); assert.equal(server.connected, false);
  } finally {
    globalThis.fetch = realFetch; server.disconnect(); server.available = false; server.vault = false;
    if (descriptor) Object.defineProperty(globalThis, 'localStorage', descriptor); else delete globalThis.localStorage;
  }
});
