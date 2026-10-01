// Local workerd only: synthetic vaults, no existing user documents or recovery keys.
import assert from 'node:assert/strict';
import { randomBytes } from 'node:crypto';
const base = process.env.WIXEL_WORKER_URL || 'http://127.0.0.1:8787';
assert.ok(['127.0.0.1', 'localhost'].includes(new URL(base).hostname), '복원 회귀 검사는 로컬 workerd에서만 실행합니다.');
const key = randomBytes(32).toString('base64url'), other = randomBytes(32).toString('base64url');
let requests = 0, checks = 0;
const sessions = new Set();
async function call(path, { method = 'GET', vault = key, revision, data } = {}, expected = 200) {
  const headers = {}; if (vault) headers['X-Wixel-Vault'] = vault;
  if (revision !== undefined) headers['If-Match'] = `"${revision}"`;
  if (data !== undefined) headers['Content-Type'] = 'application/json';
  requests++;
  const res = await fetch(base + path, { method, headers, body: data, signal: AbortSignal.timeout(30000) });
  assert.equal(res.status, expected, `${method} ${path}: ${res.status}`); checks++;
  return res;
}
const plan = (a, mode = 'replace') => JSON.stringify({ format: 'wixel-vault-backup', version: 1, mode, documents: [{ name: 'a', expectedRevision: a }, { name: 'b', expectedRevision: 0 }] });
async function begin(revision) { const session = await (await call('/api/imports', { method: 'POST', data: plan(revision) }, 201)).json(); sessions.add(session.id); return session.id; }
try {
  const health = await (await call('/api/health', { vault: null })).json(); assert.equal(health.versionHistory, true); assert.equal(health.backupImport, true);
  const first = await (await call('/api/files/a', { method: 'PUT', revision: 0, data: '{"v":1}' })).json();
  const second = await (await call('/api/files/a', { method: 'PUT', revision: first.revision, data: '{"v":2}' })).json();
  const history = await (await call('/api/versions?name=a')).json(); assert.equal(history.versions[0].revision, first.revision);
  const old = await call(`/api/version?name=a&revision=${first.revision}`); assert.equal(old.headers.get('X-Wixel-Revision'), String(first.revision)); assert.deepEqual(await old.json(), { v: 1 });
  await call('/api/versions?name=a', { vault: other }, 404);
  await call(`/api/version?name=a&revision=${first.revision}`, { method: 'POST' }, 428);
  await call(`/api/version?name=a&revision=${first.revision}`, { method: 'POST', revision: first.revision }, 412);
  const restored = await (await call(`/api/version?name=a&revision=${first.revision}`, { method: 'POST', revision: second.revision })).json(); assert.ok(restored.revision > second.revision);
  assert.deepEqual(await (await call('/api/files/a')).json(), { v: 1 });
  await call('/api/imports', { method: 'POST', data: plan(0, 'create') }, 412);
  await call('/api/imports', { method: 'POST', data: '{"broken":' }, 400);
  const id = await begin(restored.revision);
  await call(`/api/imports/${id}/files/a`, { method: 'PUT', vault: other, revision: restored.revision, data: '{}' }, 404);
  await call(`/api/imports/${id}/files/a`, { method: 'PUT', revision: restored.revision, data: '{"bad":' }, 400);
  await call(`/api/imports/${id}/commit`, { method: 'POST' }, 409);
  await call(`/api/imports/${id}/files/a`, { method: 'PUT', revision: restored.revision, data: '{"v":5}' });
  await call(`/api/imports/${id}/files/b`, { method: 'PUT', revision: 0, data: '[1,2,3]' });
  assert.deepEqual(await (await call('/api/files/a')).json(), { v: 1 });
  const external = await (await call('/api/files/a', { method: 'PUT', revision: restored.revision, data: '{"external":true}' })).json();
  await call(`/api/imports/${id}/commit`, { method: 'POST' }, 412);
  assert.deepEqual(await (await call('/api/files/a')).json(), { external: true }); await call('/api/files/b', {}, 404);
  await call(`/api/imports/${id}`, { method: 'DELETE' }); sessions.delete(id);
  const next = await begin(external.revision), large = JSON.stringify({ value: '한글😀'.repeat(250000) });
  await call(`/api/imports/${next}/files/a`, { method: 'PUT', revision: external.revision, data: '{"restored":true}' });
  await call(`/api/imports/${next}/files/b`, { method: 'PUT', revision: 0, data: large });
  const committed = await (await call(`/api/imports/${next}/commit`, { method: 'POST' })).json(); sessions.delete(next); assert.equal(committed.documents.length, 2);
  assert.deepEqual(await (await call('/api/files/a')).json(), { restored: true }); assert.equal(await (await call('/api/files/b')).text(), large);
  const exported = await (await call('/api/backup')).json(); assert.equal(exported.documents.length, 2); assert.equal(exported.documents.find(x => x.name === 'a').data.restored, true);
  assert.ok((await (await call('/api/versions?name=a')).json()).versions.some(v => v.revision === external.revision));
} finally {
  for (const id of sessions) await call(`/api/imports/${id}`, { method: 'DELETE' });
  const files = await (await call('/api/files')).json();
  for (const file of files) await call(`/api/files/${encodeURIComponent(file.name)}`, { method: 'DELETE', revision: file.revision });
  assert.deepEqual(await (await call('/api/files')).json(), []); await call('/api/versions?name=a', {}, 404);
}
console.log(JSON.stringify({ ok: true, requests, httpChecks: checks, currentDocumentsAndHistoryRemoved: true, syntheticVaultsOnly: true }));
