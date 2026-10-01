// 큰 합성 문서의 정확한 20MiB 경계 및 초과 거부를 실제 workerd에서 확인합니다.
import assert from 'node:assert/strict';
import { randomBytes, createHash } from 'node:crypto';
const base = process.env.WIXEL_WORKER_URL || 'http://127.0.0.1:8787';
const key = randomBytes(32).toString('base64url'), max = 20 * 1024 * 1024;
const headers = { 'X-Wixel-Vault': key, 'Content-Type': 'application/json', 'If-Match': '"0"' };
const body = '"' + 'x'.repeat(max - 2) + '"';
let revision;
try {
  const put = await fetch(base + '/api/files/limit', { method: 'PUT', headers, body });
  assert.equal(put.status, 200); revision = (await put.json()).revision;
  const got = await fetch(base + '/api/files/limit', { headers: { 'X-Wixel-Vault': key } });
  assert.equal(got.status, 200);
  const hash = createHash('sha256'); let size = 0;
  for await (const chunk of got.body) { hash.update(chunk); size += chunk.length; }
  assert.equal(size, max);
  assert.equal(hash.digest('hex'), createHash('sha256').update(body).digest('hex'));
  const rejected = await fetch(base + '/api/files/oversized', { method: 'PUT', headers, body: body + ' ' });
  assert.equal(rejected.status, 413); await rejected.arrayBuffer();
  console.log(JSON.stringify({ exactLimitBytes: size, roundTrip: true, overLimitStatus: rejected.status }));
} finally {
  if (revision) {
    const removed = await fetch(base + '/api/files/limit', { method: 'DELETE', headers: { 'X-Wixel-Vault': key, 'If-Match': '"' + revision + '"' } });
    assert.equal(removed.status, 200); await removed.arrayBuffer();
  }
}
