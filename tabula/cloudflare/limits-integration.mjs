// 로컬 workerd에서 실제 크기의 합성 packed JSON 및 정확한 32MiB 경계를 검증합니다.
import assert from 'node:assert/strict';
import { randomBytes } from 'node:crypto';
import { LIMITS } from './shared.js';
import { packedPayload, streamDigest } from './large-payload-fixture.mjs';
const base = process.env.WIXEL_WORKER_URL || 'http://127.0.0.1:8787';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(base).hostname)) throw new Error('이 검사는 로컬 workerd에서만 실행할 수 있습니다.');
const key = randomBytes(32).toString('base64url'), max = LIMITS.maxDocumentBytes;
const headers = { 'X-Wixel-Vault': key, 'Content-Type': 'application/json' };
const name = 'packed-limit-synthetic'; let revision;
const request = (payload, expected) => fetch(`${base}/api/files/${name}`, { method:'PUT', headers:{...headers,'If-Match':`"${expected}"`}, body:payload.stream(), duplex:'half' });
async function verify(payload) {
  const got=await fetch(`${base}/api/files/${name}`,{headers:{'X-Wixel-Vault':key}});
  assert.equal(got.status,200);
  assert.deepEqual(await streamDigest(got.body),{bytes:payload.bytes,sha256:payload.sha256});
}
try {
  const health=await (await fetch(base+'/api/health')).json();
  assert.equal(health.maxDocumentBytes,max);assert.equal(health.maxVaultBytes,100*1024*1024);assert.equal(health.maxHistoryBytes,100*1024*1024);
  const packed=packedPayload(27_536_555);
  const saved=await request(packed,0);assert.equal(saved.status,200);revision=(await saved.json()).revision;
  await verify(packed);
  const rejected=await request(packedPayload(max+1),revision);assert.equal(rejected.status,413);assert.equal((await rejected.json()).code,'DOCUMENT_TOO_LARGE');
  await verify(packed);
  const list=await (await fetch(base+'/api/files',{headers:{'X-Wixel-Vault':key}})).json();assert.equal(list[0].revision,revision);
  const exact=packedPayload(max), boundary=await request(exact,revision);assert.equal(boundary.status,200);revision=(await boundary.json()).revision;
  await verify(exact);
  console.log(JSON.stringify({packedBytes:packed.bytes,packedRoundTrip:true,overLimitStatus:413,rollback:true,exactLimitBytes:max,exactRoundTrip:true}));
} finally {
  if(revision){const removed=await fetch(`${base}/api/files/${name}`,{method:'DELETE',headers:{'X-Wixel-Vault':key,'If-Match':`"${revision}"`}});assert.equal(removed.status,200);await removed.arrayBuffer();}
}
