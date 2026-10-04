import { unpackPublishedBlob } from './publish.js';
import { readWixelFile } from './wixel-file.js';

async function decodeStoredDocument(data, options = {}) {
  return data?.format === 'wixel-packed' ? readWixelFile(unpackPublishedBlob(data), options) : data;
}
const connectionChanged = () => Object.assign(new Error('보관함 연결이 바뀌었습니다. 현재 연결에서 다시 시도하세요.'), { code:'CONNECTION_CHANGED' });
const documentChanged = () => Object.assign(new Error('문서가 바뀌어 이전 온라인 요청을 중단했습니다.'), { code:'DOCUMENT_OPEN_CANCELLED' });
const requestBody = data => data instanceof Blob ? data : JSON.stringify(data);
// 서버 저장소 (server.js 의 /api) 클라이언트. 정적 호스팅이면 available = false.
const TOKEN_KEY = 'tabula.serverToken';
const CONNECTION_KEY = 'wixel.connection.v3';
const REVISION_KEY = 'wixel.revisions.v3';
function readSetting(key, fallback) { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } }
function writeSetting(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* 세션 동안 연결 유지 */ } }
let connection = readSetting(CONNECTION_KEY, null);
let revisions = readSetting(REVISION_KEY, {});
let connectionVersion = 0;
export function createVaultKey() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return btoa(String.fromCharCode(...bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
export function validVaultKey(key) { return /^[A-Za-z0-9_-]{42}[AEIMQUYcgkosw048]$/.test(String(key ?? '')); }
/** Backup metadata is input, never authority: destination revisions come from the current vault. */
export function prepareBackupImport(backup, { overwrite = false, expectedRevisions = [] } = {}, limits = {}) {
  const bad = () => { throw new Error('WIXEL 보관함 백업 파일의 형식이나 문서 목록을 확인하세요.'); };
  if (!backup || backup.format !== 'wixel-vault-backup' || backup.version !== 1 || !Array.isArray(backup.documents) || !backup.documents.length || backup.documents.length > (limits.maxDocuments ?? 50)) bad();
  if (!Array.isArray(expectedRevisions)) bad();
  const expected = new Map(expectedRevisions.map((item) => [item.name, item.revision])), names = new Set();
  let bytes = 0;
  const documents = backup.documents.map((item) => {
    const name = typeof item?.name === 'string' ? item.name.trim() : '';
    if (!name || item.name.length > 120 || /[\u0000-\u001f\u007f]/.test(name) || names.has(name) || !Object.hasOwn(item, 'data')) bad();
    names.add(name);
    const revision = overwrite ? (expected.get(name) ?? 0) : 0;
    if (!Number.isSafeInteger(revision) || revision < 0) bad();
    const text = JSON.stringify(item.data);
    if (text === undefined) bad();
    const size = new TextEncoder().encode(text).byteLength; bytes += size;
    if (size > (limits.maxDocumentBytes ?? 20971520) || bytes > (limits.maxVaultBytes ?? 104857600)) throw new Error('백업의 문서 크기 또는 전체 용량이 보관함 한도를 초과합니다.');
    // Keep only the existing parsed data; do not retain a second full JSON string per document.
    return { name, expectedRevision: revision, data: item.data };
  });
  return { format: backup.format, version: 1, mode: overwrite ? 'replace' : 'create', documents };
}
function revision(path, value) {
  if (value !== undefined) { revisions[path] = String(value); writeSetting(REVISION_KEY, revisions); }
  return revisions[path] ?? '0';
}

// 큰 통합 문서는 localStorage(약 5MB, JSON 문자열 변환 필요) 대신 IndexedDB 에 객체 그대로 저장
let dbPromise = null, dbConnection = null;
function releaseDb(db) {
  // A late close event from an old connection must not clear a newer open.
  if (dbConnection === db) { dbConnection = null; dbPromise = null; }
  try { db.close(); } catch { /* Already closed. */ }
}
function openDb() {
  if (!globalThis.indexedDB) return Promise.reject(new Error('IndexedDB 없음'));
  if (!dbPromise) {
    const pending = new Promise((resolve, reject) => {
      const req = indexedDB.open('tabula', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('docs');
      req.onsuccess = () => {
        const db = req.result;
        dbConnection = db;
        db.onversionchange = () => releaseDb(db);
        db.onclose = () => releaseDb(db);
        resolve(db);
      };
      req.onerror = () => reject(req.error);
    });
    dbPromise = pending;
    // A transient open failure must not poison every later autosave in this tab.
    pending.catch(() => { if (dbPromise === pending) dbPromise = null; });
  }
  return dbPromise;
}
async function withDbTransaction(mode, run) {
  let db = await openDb(), tx;
  try { tx = db.transaction('docs', mode); }
  catch (error) {
    if (error?.name !== 'InvalidStateError') throw error;
    // No transaction exists yet, so reopening cannot repeat a committed write.
    // Never retry a request, callback, or an in-flight/aborted transaction.
    releaseDb(db);
    db = await openDb();
    tx = db.transaction('docs', mode);
  }
  // Queue requests synchronously: do not await after creating the transaction.
  return run(tx);
}
export async function idbSet(key, value) {
  return withDbTransaction('readwrite', tx => new Promise((resolve, reject) => {
    tx.objectStore('docs').put(value, key);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error ?? new DOMException('저장 트랜잭션이 중단되었습니다.', 'AbortError'));
  }));
}
export async function idbDel(key) {
  return withDbTransaction('readwrite', tx => new Promise((resolve, reject) => {
    tx.objectStore('docs').delete(key);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error ?? new DOMException('삭제 트랜잭션이 중단되었습니다.', 'AbortError'));
  }));
}
export async function idbGet(key) {
  return withDbTransaction('readonly', tx => new Promise((resolve, reject) => {
    const req = tx.objectStore('docs').get(key);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
  }));
}

/** Only the small generation manifest changes in this atomic transaction. */
export async function idbCompareAndSet(key, expected, value, validate = () => {}) {
  return withDbTransaction('readwrite', tx => new Promise((resolve, reject) => {
    const store = tx.objectStore('docs');
    let failure = null;
    const fail = (error) => { failure = error; try { tx.abort(); } catch { reject(error); } };
    const token = (item) => [3,4].includes(item?.v) ? item.generation : JSON.stringify(item ?? null);
    const req = store.get(key);
    req.onsuccess = () => {
      try {
        validate();
        if (token(req.result) !== token(expected)) throw Object.assign(new Error('다른 탭에서 저장한 문서가 바뀌었습니다. 다시 저장하세요.'), { code: 'IDB_CONFLICT' });
        const put = store.put(value, key);
        put.onsuccess = () => { try { validate(); } catch (error) { fail(error); } };
      } catch (error) { fail(error); }
    };
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(failure ?? tx.error);
    tx.onabort = () => reject(failure ?? tx.error ?? new DOMException('저장 트랜잭션이 중단되었습니다.', 'AbortError'));
  }));
}

export async function idbKeys(prefix) {
  return withDbTransaction('readonly', tx => new Promise((resolve, reject) => {
    const req = tx.objectStore('docs').getAllKeys(IDBKeyRange.bound(prefix, prefix + '\uffff'));
    req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error);
  }));
}

export async function idbDeleteMany(keys) {
  if (!keys.length) return;
  return withDbTransaction('readwrite', tx => new Promise((resolve, reject) => {
    const store = tx.objectStore('docs');
    for (const key of keys) store.delete(key);
    tx.oncomplete = () => resolve(true); tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error ?? new DOMException('삭제 트랜잭션이 중단되었습니다.', 'AbortError'));
  }));
}

/** Atomically read/update related index, document and version records.
 * update(Map) is synchronous: {set:[[key,value]], delete:[key], result}.
 * Prepare compression/network work before entering the transaction.
 */
export async function idbUpdate(readKeys, update) {
  return withDbTransaction('readwrite', tx => new Promise((resolve, reject) => {
    const store = tx.objectStore('docs');
    const keys = [...new Set(readKeys)], values = new Map();
    let result, failure, remaining = keys.length;
    const apply = () => {
      try {
        const change = update(values);
        if (change && typeof change.then === 'function') { Promise.resolve(change).catch(() => {}); throw new TypeError('저장 트랜잭션 콜백은 동기 변경 내용을 반환해야 합니다.'); }
        if (!change) throw new TypeError('저장 트랜잭션 변경 내용이 없습니다.');
        for (const [itemKey,value] of change.set ?? []) store.put(value, itemKey);
        for (const itemKey of change.delete ?? []) store.delete(itemKey);
        result = change.result;
      } catch (error) { failure = error; try { tx.abort(); } catch { reject(error); } }
    };
    tx.oncomplete = () => resolve(result);
    tx.onerror = () => reject(failure ?? tx.error);
    tx.onabort = () => reject(failure ?? tx.error ?? new DOMException('저장 트랜잭션이 중단되었습니다.', 'AbortError'));
    for (const itemKey of keys) {
      const req = store.get(itemKey);
      req.onsuccess = () => { values.set(itemKey, req.result ?? null); if (--remaining === 0) apply(); };
    }
    if (!remaining) apply();
  }));
}

function token() {
  try { return localStorage.getItem(TOKEN_KEY) ?? ''; } catch { return ''; }
}

export const server = {
  available: false,
  needsToken: false,
  vault: false,
  capabilities: {},
  get connected() { return this.available && !!connection && (this.vault ? validVaultKey(connection.key) : connection.kind === 'node'); },
  get label() { return this.vault ? '개인 보관함' : '연결 서버'; },
  get connectionVersion() { return connectionVersion; },
  connect(key = null) {
    if (this.vault && !validVaultKey(key)) throw new Error('복구키는 43자의 영문·숫자·밑줄·하이픈입니다.');
    if (connection?.key !== key) { revisions = {}; writeSetting(REVISION_KEY, revisions); }
    connection = this.vault ? { kind: 'vault', key } : { kind: 'node' };
    connectionVersion++;
    writeSetting(CONNECTION_KEY, connection);
  },
  disconnect() { connection = null; connectionVersion++; revisions = {}; writeSetting(CONNECTION_KEY, null); writeSetting(REVISION_KEY, revisions); },
  recoveryKey() { return this.vault && this.connected ? connection.key : null; },
  headers() { return { 'X-Tabula-Token': token(), ...(this.vault && this.connected ? { 'X-Wixel-Vault': connection.key } : {}) }; },
  revision(path) { return revision(path); },
  restoreRevision(path, value) { revision(path, value); },

  setToken(t) {
    try { localStorage.setItem(TOKEN_KEY, t); } catch { /* 무시 */ }
  },

  async init() {
    this.available = false;
    this.needsToken = false;
    this.vault = false;
    this.capabilities = {};
    // 한 파일 배포본(build.mjs)이나 파일로 직접 연 경우에는 서버가 없음
    if (globalThis.TABULA_STATIC || globalThis.location?.protocol === 'file:') return false;
    try {
      const res = await fetch('api/health', { cache: 'no-store' });
      const data = res.ok ? await res.json() : null;
      this.available = !!data?.ok;
      this.needsToken = !!data?.auth;
      this.vault = !!data?.vault;
      this.capabilities = data ?? {};
    } catch {
      this.available = false;
    }
    return this.available;
  },

  async request(path, opts = {}) {
    const sentConnection = connectionVersion, { isCurrent, ...fetchOptions } = opts;
    const current = () => {
      if (sentConnection !== connectionVersion) throw connectionChanged();
      if (isCurrent?.() === false) throw documentChanged();
      return true;
    };
    current();
    const res = await fetch(`api/${path}`, {
      ...fetchOptions,
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json', ...this.headers(), ...(opts.headers ?? {}) },
    });
    try { current(); } catch (error) { await res.body?.cancel().catch(() => {}); throw error; }
    if (!res.ok) {
      let msg = `서버 오류 (${res.status})`;
      let data = {};
      try { data = await res.json(); msg = data.error ?? msg; } catch { /* 무시 */ }
      current();
      throw Object.assign(new Error(msg), { status: res.status, code: data.code, currentRevision: data.currentRevision });
    }
    const parsed = await res.json(); current();
    const data = await decodeStoredDocument(parsed, { isCurrent:current, signal:opts.signal });
    current();
    if (/^(files|published)\//.test(path)) {
      const rev = res.headers.get('X-Wixel-Revision') ?? res.headers.get('ETag')?.replace(/^"|"$/g, '') ?? data?.revision;
      if (rev != null) revision(path, rev);
    }
    return data;
  },

  /** 웹 가져오기 함수 중계 (/api/fetch): 텍스트 */
  async fetchText(url) {
    const res = await fetch(`api/fetch?url=${encodeURIComponent(url)}`, { cache: 'no-store', headers: this.headers() });
    const text = await res.text();
    if (!res.ok) { let msg = text; try { msg = JSON.parse(text).error ?? text; } catch { /* 글자 */ } throw Object.assign(new Error(msg || `오류 (${res.status})`), { status: res.status }); }
    return text;
  },
  list() { return this.request('files'); },
  async publish(data, id = null, options = {}) {
    const sentConnection = connectionVersion;
    const res = id ? await this.mutate(`published/${id}`, 'PUT', data, options) : await this.request('publish', { ...options, method: 'POST', body: requestBody(data) });
    if (sentConnection !== connectionVersion) throw connectionChanged();
    if (options.isCurrent?.() === false) throw documentChanged();
    if (res.revision != null) revision(`published/${id ?? res.id}`, res.revision);
    return { ...res, id: id ?? res.id };
  },
  unpublish(id) { return this.mutate(`published/${id}`, 'DELETE'); },
  async published(id, options = {}) {
    const current = () => { if (options.isCurrent?.() === false) throw documentChanged(); return true; };
    current();
    const res = await fetch(`api/published/${id}`, { cache: 'no-store', ...(options.signal?{signal:options.signal}:{}) });
    try { current(); } catch (error) { await res.body?.cancel().catch(() => {}); throw error; }
    if (!res.ok) throw Object.assign(new Error((await res.json().catch(() => ({}))).error ?? `오류 (${res.status})`), { status: res.status });
    const parsed = await res.json(); current();
    const data = await decodeStoredDocument(parsed, { isCurrent:current, signal:options.signal });
    current(); return { data, modified: Number(res.headers.get('X-Modified')) || 0 };
  },
  load(name, options = {}) { return this.request(`files/${encodeURIComponent(name)}`, options); },
  mutate(path, method, data, options = {}) {
    return this.request(path, { ...options, method, ...(data === undefined ? {} : { body: requestBody(data) }), headers: { ...options.headers, ...(this.vault ? { 'If-Match': `"${revision(path)}"` } : {}) } });
  },
  save(name, data, options = {}) { return this.mutate(`files/${encodeURIComponent(name)}`, 'PUT', data, options); },
  remove(name, expected) {
    const path = `files/${encodeURIComponent(name)}`;
    return this.vault && expected != null ? this.request(path, { method: 'DELETE', headers: { 'If-Match': `"${expected}"` } }) : this.mutate(path, 'DELETE');
  },
  backup() { return this.request('backup'); },
  requireCapability(name) {
    if (!this.vault || !this.capabilities[name]) throw Object.assign(new Error('이 서버는 온라인 버전 기록·백업 복원을 지원하지 않습니다. 파일 또는 브라우저 보관함을 사용하세요.'), { code: 'UNSUPPORTED_CAPABILITY' });
  },
  versions(name) { this.requireCapability('versionHistory'); return this.request(`versions?name=${encodeURIComponent(name)}`); },
  loadVersion(name, value, options = {}) { this.requireCapability('versionHistory'); return this.request(`version?name=${encodeURIComponent(name)}&revision=${encodeURIComponent(value)}`, options); },
  async restoreVersion(name, value, expected) {
    this.requireCapability('versionHistory');
    if (!Number.isSafeInteger(Number(expected)) || Number(expected) < 1) throw new Error('현재 문서 버전을 먼저 확인하세요.');
    const result = await this.request(`version?name=${encodeURIComponent(name)}&revision=${encodeURIComponent(value)}`, { method: 'POST', headers: { 'If-Match': `"${expected}"` } });
    revision(`files/${encodeURIComponent(name)}`, result.revision); return result;
  },
  async importBackup(backup, options = {}) {
    this.requireCapability('backupImport');
    const plan = prepareBackupImport(backup, options, this.capabilities), sentConnection = connectionVersion;
    const sameConnection = () => { if (connectionVersion !== sentConnection) throw Object.assign(new Error('보관함 연결이 바뀌어 복원을 중단했습니다.'), { code: 'CONNECTION_CHANGED' }); };
    const session = await this.request('imports', { method: 'POST', signal: options.signal, body: JSON.stringify({ ...plan, documents: plan.documents.map(({ name, expectedRevision }) => ({ name, expectedRevision })) }) });
    let committing = false;
    try {
      for (let i = 0; i < plan.documents.length; i++) {
        sameConnection(); const item = plan.documents[i];
        await this.request(`imports/${session.id}/files/${encodeURIComponent(item.name)}`, { method: 'PUT', signal: options.signal, headers: { 'If-Match': `"${item.expectedRevision}"` }, body: JSON.stringify(item.data) });
        options.onProgress?.(i + 1, plan.documents.length);
      }
      sameConnection(); committing = true;
      const result = await this.request(`imports/${session.id}/commit`, { method: 'POST', signal: options.signal });
      for (const item of result.documents) revision(`files/${encodeURIComponent(item.name)}`, item.revision);
      return result;
    } catch (error) {
      if (connectionVersion === sentConnection) await this.request(`imports/${session.id}`, { method: 'DELETE' }).catch(() => {});
      if (committing && !error.status && error.code !== 'CONNECTION_CHANGED') {
        error.code = 'IMPORT_RESULT_UNKNOWN'; error.message = '복원 결과를 확인하지 못했습니다. 자동으로 다시 복원하지 말고 온라인 문서 목록을 확인하세요.';
      }
      throw error;
    }
  },
  publications() { return this.request('publications'); },
  async publicationRevision(id) {
    const sentConnection = connectionVersion;
    const res = await fetch(`api/published/${encodeURIComponent(id)}`, { cache: 'no-store' });
    if (!res.ok) throw Object.assign(new Error('게시 문서를 찾을 수 없습니다.'), { status: res.status });
    const rev = res.headers.get('X-Wixel-Revision') ?? res.headers.get('ETag')?.replace(/^"|"$/g, '');
    await res.body?.cancel();
    if (sentConnection !== connectionVersion) throw connectionChanged();
    if (rev == null) throw new Error('게시 버전을 확인할 수 없습니다. 다시 시도하세요.');
    revision(`published/${id}`, rev);
    return rev;
  },
};
