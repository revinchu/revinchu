import { DurableObject } from 'cloudflare:workers';
import { VaultStore } from './store.js';
import { ApiError, LIMITS, documentHeaders, errorResponse, json, parseRevision, randomId, sameHash, securityHeaders, sha256, validPublicId, validateName, validateOrigin, vaultKey } from './shared.js';
import { boundedBytes, fetchPublicText, naverKeywords } from './proxy.js';

function storeFor(object, create = false, limits = LIMITS) {
  if (object._store) return object._store;
  const exists = object._storage.sql.exec("SELECT name FROM sqlite_master WHERE type='table' AND name='meta'").toArray().length > 0;
  if (!exists && !create) return null;
  object._store = new VaultStore(object._storage, object._limits ?? limits); return object._store;
}
function requireStore(object) {
  const store = storeFor(object);
  if (!store) throw new ApiError(404, '문서가 없습니다.', 'NOT_FOUND');
  return store;
}
// 존재하지 않는 보관함/게시물의 조회는 DDL을 실행하거나 빈 테이블을 저장하지 않습니다.
// RPC는 Response/ReadableStream을 전달하므로 20MiB 문서를 한 RPC 값으로 복제하지 않습니다.
export class UserVault extends DurableObject {
  constructor(ctx, env) { super(ctx, env); this._storage = ctx.storage; this._store = null; }
  list() { return json((storeFor(this)?.list() ?? []).map(({ chunks, ...item }) => item)); }
  usage() { return json(storeFor(this)?.usage() ?? { documents: 0, bytes: 0, publications: 0, ...LIMITS }); }
  read(name) { try { const { meta, body } = requireStore(this).read(name); return new Response(body, { headers: documentHeaders(meta) }); } catch (e) { return errorResponse(e); } }
  async put(name, revision, body) { try { return json(await storeFor(this, true).put(name, revision, body)); } catch (e) { return errorResponse(e); } }
  remove(name, revision) { try { return json(requireStore(this).remove(name, revision)); } catch (e) { return errorResponse(e); } }
  versions(name) { try { return json(requireStore(this).versions(name)); } catch (e) { return errorResponse(e); } }
  readVersion(name, revision) { try { const { meta, body } = requireStore(this).readVersion(name, revision); return new Response(body, { headers: documentHeaders(meta) }); } catch (e) { return errorResponse(e); } }
  restore(name, revision, expected) { try { return json(requireStore(this).restore(name, revision, expected)); } catch (e) { return errorResponse(e); } }
  async beginImport(manifest) {
    const store = storeFor(this, true); let session;
    try { session = store.beginImport(manifest); await this._storage.setAlarm(session.expires); return json(session, 201); }
    catch (e) { if (session) store.cancelImport(session.id); return errorResponse(e); }
  }
  alarm() { storeFor(this)?.expireImports(); }
  async stageImport(id, name, expected, body) { try { return json(await requireStore(this).stageImport(id, name, expected, body)); } catch (e) { return errorResponse(e); } }
  commitImport(id) { try { return json(requireStore(this).commitImport(id)); } catch (e) { return errorResponse(e); } }
  cancelImport(id) { try { return json(requireStore(this).cancelImport(id)); } catch (e) { return errorResponse(e); } }
  reserve(id) { try { return json(storeFor(this, true).reservePublication(id)); } catch (e) { return errorResponse(e); } }
  release(id) { return json(storeFor(this)?.releasePublication(id) ?? { ok: true }); }
  publications() { return json(storeFor(this)?.publications() ?? []); }
  hasPublication(id) { return storeFor(this)?.publications().some(item => item.id === id) ?? false; }
  backup() {
    const store = storeFor(this);
    const body = store ? store.backup() : JSON.stringify({ format: 'wixel-vault-backup', version: 1, exportedAt: new Date().toISOString(), documents: [] });
    return new Response(body, { headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'Content-Disposition': 'attachment; filename="wixel-vault-backup.json"' } });
  }
}
export class PublishedDocument extends DurableObject {
  constructor(ctx, env) { super(ctx, env); this._storage = ctx.storage; this._store = null; this._limits = { ...LIMITS, maxDocuments: 1, maxVaultBytes: LIMITS.maxDocumentBytes, maxVersions: 0, maxHistoryBytes: 0 }; }
  authorize(owner) {
    const store = requireStore(this);
    if (!sameHash(store.meta('owner'), owner)) throw new ApiError(403, '이 게시물을 변경할 권한이 없습니다.', 'PUBLISH_OWNER');
    return store;
  }
  async create(owner, body) {
    try {
      const store = storeFor(this, true, { ...LIMITS, maxDocuments: 1, maxVaultBytes: LIMITS.maxDocumentBytes });
      if (store.meta('owner')) throw new ApiError(409, '이미 사용한 게시 주소입니다.', 'PUBLISH_EXISTS');
      store.setMeta('owner', owner);
      return json(await store.put('document', 0, body));
    } catch (e) { return errorResponse(e); }
  }
  read() { try { const { meta, body } = requireStore(this).read('document'); return new Response(body, { headers: documentHeaders(meta) }); } catch (e) { return errorResponse(e); } }
  async update(owner, revision, body) {
    try { const store = this.authorize(owner); if (!store.find('document')) throw new ApiError(404, '게시가 중지된 문서입니다.', 'NOT_FOUND');
      return json(await store.put('document', revision, body)); } catch (e) { return errorResponse(e); }
  }
  remove(owner, revision) {
    try { return json(this.authorize(owner).remove('document', revision)); } catch (e) { return errorResponse(e); }
  }
  // 생성 요청이 실패했을 때만 내부 Worker가 호출합니다. 외부 HTTP에는 노출하지 않습니다.
  abandon(owner) {
    try { const store = this.authorize(owner), current = store.find('document'); if (current) store.remove('document', current.revision); return json({ ok: true }); } catch (e) { return errorResponse(e); }
  }
}

async function limited(binding, key) {
  if (!binding) throw new ApiError(503, '요청 제한 설정을 확인하는 중입니다.', 'RATE_UNAVAILABLE');
  if (!(await binding.limit({ key })).success) throw new ApiError(429, '요청이 많습니다. 잠시 후 다시 시도하세요.', 'RATE_LIMIT');
}
function nameFrom(path, prefix) {
  try { return validateName(decodeURIComponent(path.slice(prefix.length))); } catch (e) {
    if (e instanceof ApiError) throw e; throw new ApiError(400, '문서 이름의 주소 형식이 올바르지 않습니다.', 'INVALID_NAME');
  }
}
function checkBody(request) {
  const length = request.headers.get('Content-Length');
  if (length !== null && Number(length) > LIMITS.maxDocumentBytes) throw new ApiError(413, '문서 하나는 최대 20MiB까지 저장할 수 있습니다.', 'DOCUMENT_TOO_LARGE');
  const type = request.headers.get('Content-Type') ?? '';
  if (!/^application\/json(?:\s*;|$)/i.test(type)) throw new ApiError(415, 'JSON 형식으로 문서를 전송하세요.', 'CONTENT_TYPE');
  if (request.headers.has('Content-Encoding')) throw new ApiError(415, '압축하지 않은 JSON 문서를 전송하세요.', 'CONTENT_ENCODING');
}
async function importManifest(request) {
  checkBody(request);
  const control = new AbortController(), timer = setTimeout(() => control.abort(), 10000);
  try {
    const bytes = await boundedBytes(new Response(request.body, { headers: request.headers }), 32768, control.signal);
    try { return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)); }
    catch { throw new ApiError(400, '백업 복원 목록이 올바른 UTF-8 JSON이 아닙니다.', 'INVALID_BACKUP'); }
  } finally { clearTimeout(timer); }
}
function sourceRevision(value) {
  if (!/^[1-9]\d*$/.test(value ?? '') || !Number.isSafeInteger(Number(value))) throw new ApiError(400, '복원할 버전 번호가 올바르지 않습니다.', 'INVALID_REVISION');
  return Number(value);
}
async function smallResponse(response) { return new Response(await response.text(), response); }
export async function route(request, env) {
  const url = new URL(request.url), path = url.pathname, method = request.method;
  if (!path.startsWith('/api/')) return env.ASSETS.fetch(request);
  validateOrigin(request);
  const ip = request.headers.get('CF-Connecting-IP') || 'local';
  await limited(env.API_RATE, ip);
  if (path === '/api/health' && method === 'GET') return json({ ok: true, auth: false, vault: true, publish: true, versionHistory: true, backupImport: true, ...LIMITS });
  if (path === '/api/fetch' && method === 'GET') {
    await limited(env.FETCH_RATE, ip);
    return fetchPublicText(url.searchParams.get('url') || '', url.hostname);
  }
  if (path.startsWith('/api/published/') && method === 'GET') {
    const id = path.slice('/api/published/'.length);
    if (!validPublicId(id)) throw new ApiError(404, '게시물이 없습니다.', 'NOT_FOUND');
    return env.PUBLISHED.getByName(id).read();
  }
  const owner = await sha256(vaultKey(request));
  await limited(env.VAULT_RATE, owner);
  const vault = env.VAULTS.getByName(owner);
  if (path === '/api/files' && method === 'GET') return vault.list();
  if (path === '/api/usage' && method === 'GET') return vault.usage();
  if (path === '/api/backup' && method === 'GET') return vault.backup();
  if (path === '/api/versions' && method === 'GET') return vault.versions(validateName(url.searchParams.get('name')));
  if (path === '/api/version' && ['GET', 'POST'].includes(method)) {
    const name = validateName(url.searchParams.get('name')), revision = sourceRevision(url.searchParams.get('revision'));
    return method === 'GET' ? vault.readVersion(name, revision) : smallResponse(await vault.restore(name, revision, parseRevision(request)));
  }
  if (path === '/api/imports' && method === 'POST') return smallResponse(await vault.beginImport(await importManifest(request)));
  const importing = /^\/api\/imports\/([A-Za-z0-9_-]{43})(?:\/(commit|files\/(.+)))?$/.exec(path);
  if (importing) {
    const id = importing[1];
    if (!importing[2] && method === 'DELETE') return smallResponse(await vault.cancelImport(id));
    if (importing[2] === 'commit' && method === 'POST') return smallResponse(await vault.commitImport(id));
    if (importing[3] && method === 'PUT') { checkBody(request); return smallResponse(await vault.stageImport(id, nameFrom(path, `/api/imports/${id}/files/`), parseRevision(request), request.body)); }
  }
  if (path === '/api/publications' && method === 'GET') return vault.publications();
  if (path.startsWith('/api/files/')) {
    const name = nameFrom(path, '/api/files/');
    if (method === 'GET') return vault.read(name);
    if (method === 'PUT') { checkBody(request); return vault.put(name, parseRevision(request), request.body); }
    if (method === 'DELETE') {
      const removed = await vault.remove(name, parseRevision(request));
      // 작은 삭제 응답을 현재 요청 안에서 읽어 RPC 스트림이 먼저 끊기지 않게 합니다.
      return new Response(await removed.text(), removed);
    }
  }
  if (path === '/api/publish' && method === 'POST') {
    checkBody(request);
    const id = randomId(), reserved = await vault.reserve(id);
    if (!reserved.ok) return reserved;
    const publication = env.PUBLISHED.getByName(id);
    try {
      const created = await publication.create(owner, request.body);
      if (!created.ok) {
        await publication.abandon(owner); await vault.release(id); return created;
      }
      return json({ id, ...(await created.json()) }, 201);
    } catch (error) {
      // 통신 결과가 불확실할 때 게시 링크를 보관함 목록에 남겨 사용자가 중지할 수 있게 합니다.
      throw error;
    }
  }
  if (path.startsWith('/api/published/') && ['PUT', 'DELETE'].includes(method)) {
    const id = path.slice('/api/published/'.length);
    if (!validPublicId(id)) throw new ApiError(404, '게시물이 없습니다.', 'NOT_FOUND');
    if (!await vault.hasPublication(id)) throw new ApiError(403, '이 게시물을 변경할 권한이 없습니다.', 'PUBLISH_OWNER');
    const publication = env.PUBLISHED.getByName(id), revision = parseRevision(request);
    if (method === 'PUT') { checkBody(request); return publication.update(owner, revision, request.body); }
    const removed = await publication.remove(owner, revision);
    if (removed.ok || removed.status === 404) { await vault.release(id); return json({ ok: true }); }
    return removed;
  }
  if (path === '/api/naver/keywordstool' && method === 'GET') { await limited(env.FETCH_RATE, ip); return naverKeywords(request); }
  throw new ApiError(404, '요청한 API가 없습니다.', 'NOT_FOUND');
}
export default {
  async fetch(request, env) {
    try { return securityHeaders(await route(request, env)); } catch (error) { return securityHeaders(errorResponse(error)); }
  },
};
