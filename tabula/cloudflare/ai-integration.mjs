// Synthetic online document only. Never reads the browser's or user's recovery key.
import assert from 'node:assert/strict';
import { randomBytes, createHash } from 'node:crypto';
import { Workbook } from '../src/workbook.js';
const base = process.env.WIXEL_WORKER_URL || 'http://127.0.0.1:8792';
const key = randomBytes(32).toString('base64url'), other = randomBytes(32).toString('base64url');
const name = 'AI-MCP-검증-' + Date.now(), grantIds = [];
let revision = 0, assertions = 0;
function check(value, expected, label) { assert.equal(value, expected, label); assertions++; }
async function call(path, { method = 'GET', data, headers = {}, auth = true, redirect = 'follow' } = {}) {
  return fetch(base + path, { method, headers: { ...(auth ? { 'X-Wixel-Vault': key } : {}), ...(data === undefined ? {} : { 'Content-Type': 'application/json' }), ...headers }, body: data === undefined ? undefined : JSON.stringify(data), redirect });
}
const docPath = '/api/files/' + encodeURIComponent(name);
async function rpc(token, method, params = {}) {
  const r = await call('/api/mcp', { method: 'POST', auth: false, headers: { Authorization: 'Bearer ' + token, Accept: 'application/json, text/event-stream', 'MCP-Protocol-Version': '2025-11-25' }, data: { jsonrpc: '2.0', id: assertions + 1, method, params } });
  check(r.status, 200, 'MCP HTTP'); return r.json();
}
try {
  const health = await (await call('/api/health', { auth: false })).json(); check(health.mcp, true); check(health.ai, true);
  const unauth = await call('/api/mcp', { auth: false }); check(unauth.status, 401); check(unauth.headers.get('WWW-Authenticate').includes('oauth-protected-resource'), true);
  check((await call('/api/mcp', { auth: false, headers: { Origin: 'https://evil.example' } })).status, 403);
  check((await call('/api/files', { headers: { Origin: 'https://chatgpt.com' } })).status, 403);
  const metadata = await (await call('/.well-known/oauth-protected-resource/api/mcp', { auth: false })).json(); check(metadata.resource, base + '/api/mcp');
  const book = new Workbook(); book.transact(() => { book.setInput(0, 0, 0, '100'); book.setInput(0, 0, 1, '=A1*2'); book.setStyle(0, 0, 0, { fill: '#AABBCC' }); });
  const saved = await call(docPath, { method: 'PUT', data: { workbook: book.serialize(), docName: name }, headers: { 'If-Match': '0' } }); check(saved.status, 200); revision = (await saved.json()).revision;
  let r = await call('/api/ai/connections', { method: 'POST', data: { label: '검증 읽기', document: name, write: false } }); check(r.status, 201);
  const read = await r.json(); grantIds.push(read.id);
  const init = await rpc(read.token, 'initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'wixel-test', version: '1' } }); check(init.result.serverInfo.name, 'wixel');
  const list = await rpc(read.token, 'tools/list'); check(list.result.tools.length, 3);
  const files = await rpc(read.token, 'tools/call', { name: 'list_documents', arguments: {} }); check(files.result.structuredContent.documents.length, 1);
  const cells = await rpc(read.token, 'tools/call', { name: 'read_range', arguments: { sheet: 'Sheet1', range: 'A1:B1' } }); check(cells.result.structuredContent.cells[0].input, '100');
  const forbidden = await rpc(read.token, 'tools/call', { name: 'write_cells', arguments: { sheet: 'Sheet1', expectedRevision: revision, changes: [{ cell: 'A1', input: '200' }] } }); check(forbidden.error.code, -32602);
  // DCR → consent → PKCE → real Durable Object token exchange.
  const redirect = 'https://client.example/callback';
  r = await call('/api/oauth/register', { auth: false, method: 'POST', data: { client_name: '위셀 통합 검증', redirect_uris: [redirect], token_endpoint_auth_method: 'none' } }); check(r.status, 201);
  const client = await r.json(), verifier = randomBytes(32).toString('base64url');
  const query = new URLSearchParams({ client_id: client.client_id, redirect_uri: redirect, response_type: 'code', state: 'fixture-state', resource: base + '/api/mcp', code_challenge_method: 'S256', code_challenge: createHash('sha256').update(verifier).digest('base64url'), scope: 'document:read document:write' });
  r = await call('/api/oauth/authorize?' + query, { auth: false, redirect: 'manual' }); check(r.status, 302);
  const cookie = r.headers.get('Set-Cookie').split(';')[0], request = new URL(r.headers.get('Location'), base).searchParams.get('request');
  check((await call('/api/oauth/pending?request=' + request, { auth: false })).status, 403);
  check((await call('/api/oauth/pending?request=' + request, { auth: false, headers: { Cookie: cookie } })).status, 200);
  r = await call('/api/ai/consent', { method: 'POST', headers: { Cookie: cookie }, data: { request, document: name, write: true } }); check(r.status, 200);
  const callback = new URL((await r.json()).redirect); check(callback.searchParams.get('state'), 'fixture-state');
  const form = { grant_type: 'authorization_code', code: callback.searchParams.get('code'), client_id: client.client_id, redirect_uri: redirect, resource: base + '/api/mcp', code_verifier: verifier };
  const exchange = data => fetch(base + '/api/oauth/token', { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(data) });
  r = await exchange(form); check(r.status, 200); const token = await r.json(); check(typeof token.access_token, 'string'); grantIds.push(token.access_token.split('.')[0]);
  check((await exchange(form)).status, 400);
  const change = { name: 'write_cells', arguments: { sheet: 'Sheet1', expectedRevision: revision, changes: [{ cell: 'A1', input: '250' }] } };
  const edited = await rpc(token.access_token, 'tools/call', change); check(edited.result.isError, false); revision = edited.result.structuredContent.revision;
  const conflict = await rpc(token.access_token, 'tools/call', change); check(conflict.result.isError, true);
  const restored = await rpc(token.access_token, 'tools/call', { name: 'read_range', arguments: { sheet: 'Sheet1', range: 'A1:B1' } }); check(restored.result.structuredContent.cells[0].input, '250');
  r = await call('/api/ai/connections/' + grantIds.at(-1), { method: 'DELETE', headers: { 'X-Wixel-Vault': other } }); check(r.status, 200);
  check((await rpc(token.access_token, 'ping')).result !== undefined, true, 'other vault cannot revoke');
  const rotated = await (await exchange({ grant_type: 'refresh_token', refresh_token: token.refresh_token, client_id: client.client_id, resource: base + '/api/mcp' })).json(); check(typeof rotated.access_token, 'string');
  check((await call('/api/mcp', { auth: false, headers: { Authorization: 'Bearer ' + token.access_token } })).status, 401);
  await call('/api/ai/connections/' + grantIds.at(-1), { method: 'DELETE' });
  check((await call('/api/mcp', { auth: false, headers: { Authorization: 'Bearer ' + rotated.access_token } })).status, 401);
  check((await call('/api/ai/chat', { method: 'POST', data: { provider: 'openai', model: 'x', prompt: 'test' } })).status, 400);
  console.log(JSON.stringify({ ok: true, base, checks: assertions, syntheticDocumentOnly: true }));
} finally {
  for (const id of grantIds) await call('/api/ai/connections/' + id, { method: 'DELETE' });
  const doc = await call(docPath); if (doc.ok) { const latest = doc.headers.get('X-Wixel-Revision'); await doc.body.cancel(); await call(docPath, { method: 'DELETE', headers: { 'If-Match': latest } }); }
}
