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
function revision(path, value) {
  if (value !== undefined) { revisions[path] = String(value); writeSetting(REVISION_KEY, revisions); }
  return revisions[path] ?? '0';
}

// 큰 통합 문서는 localStorage(약 5MB, JSON 문자열 변환 필요) 대신 IndexedDB 에 객체 그대로 저장
let dbPromise = null;
function openDb() {
  if (!globalThis.indexedDB) return Promise.reject(new Error('IndexedDB 없음'));
  dbPromise ??= new Promise((resolve, reject) => {
    const req = indexedDB.open('tabula', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('docs');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}
export async function idbSet(key, value) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('docs', 'readwrite');
    tx.objectStore('docs').put(value, key);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}
export async function idbDel(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('docs', 'readwrite');
    tx.objectStore('docs').delete(key);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}
export async function idbGet(key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction('docs', 'readonly').objectStore('docs').get(key);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
  });
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
    const sentConnection = connectionVersion;
    const res = await fetch(`api/${path}`, {
      ...opts,
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json', ...this.headers(), ...(opts.headers ?? {}) },
    });
    if (!res.ok) {
      let msg = `서버 오류 (${res.status})`;
      let data = {};
      try { data = await res.json(); msg = data.error ?? msg; } catch { /* 무시 */ }
      throw Object.assign(new Error(msg), { status: res.status, code: data.code, currentRevision: data.currentRevision });
    }
    const data = await res.json();
    if (sentConnection !== connectionVersion) throw Object.assign(new Error('보관함 연결이 바뀌었습니다. 현재 연결에서 다시 시도하세요.'), { code: 'CONNECTION_CHANGED' });
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
  async publish(data, id = null) {
    const res = id ? await this.mutate(`published/${id}`, 'PUT', data) : await this.request('publish', { method: 'POST', body: JSON.stringify(data) });
    if (res.revision != null) revision(`published/${id ?? res.id}`, res.revision);
    return { ...res, id: id ?? res.id };
  },
  unpublish(id) { return this.mutate(`published/${id}`, 'DELETE'); },
  async published(id) {
    const res = await fetch(`api/published/${id}`, { cache: 'no-store' });
    if (!res.ok) throw Object.assign(new Error((await res.json().catch(() => ({}))).error ?? `오류 (${res.status})`), { status: res.status });
    return { data: await res.json(), modified: Number(res.headers.get('X-Modified')) || 0 };
  },
  load(name) { return this.request(`files/${encodeURIComponent(name)}`); },
  mutate(path, method, data) {
    return this.request(path, { method, ...(data === undefined ? {} : { body: JSON.stringify(data) }), headers: this.vault ? { 'If-Match': `"${revision(path)}"` } : {} });
  },
  save(name, data) { return this.mutate(`files/${encodeURIComponent(name)}`, 'PUT', data); },
  remove(name, expected) {
    const path = `files/${encodeURIComponent(name)}`;
    return this.vault && expected != null ? this.request(path, { method: 'DELETE', headers: { 'If-Match': `"${expected}"` } }) : this.mutate(path, 'DELETE');
  },
  backup() { return this.request('backup'); },
  publications() { return this.request('publications'); },
  async publicationRevision(id) {
    const res = await fetch(`api/published/${encodeURIComponent(id)}`, { cache: 'no-store' });
    if (!res.ok) throw Object.assign(new Error('게시 문서를 찾을 수 없습니다.'), { status: res.status });
    const rev = res.headers.get('X-Wixel-Revision') ?? res.headers.get('ETag')?.replace(/^"|"$/g, '');
    await res.body?.cancel();
    if (rev == null) throw new Error('게시 버전을 확인할 수 없습니다. 다시 시도하세요.');
    revision(`published/${id}`, rev);
    return rev;
  },
};
