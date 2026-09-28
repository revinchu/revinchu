// 서버 저장소 (server.js 의 /api) 클라이언트. 정적 호스팅이면 available = false.
const TOKEN_KEY = 'tabula.serverToken';

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

  setToken(t) {
    try { localStorage.setItem(TOKEN_KEY, t); } catch { /* 무시 */ }
  },

  async init() {
    // 한 파일 배포본(build.mjs)이나 파일로 직접 연 경우에는 서버가 없음
    if (globalThis.TABULA_STATIC || globalThis.location?.protocol === 'file:') return false;
    try {
      const res = await fetch('api/health', { cache: 'no-store' });
      const data = res.ok ? await res.json() : null;
      this.available = !!data?.ok;
      this.needsToken = !!data?.auth;
    } catch {
      this.available = false;
    }
    return this.available;
  },

  async request(path, opts = {}) {
    const res = await fetch(`api/${path}`, {
      ...opts,
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json', 'X-Tabula-Token': token(), ...(opts.headers ?? {}) },
    });
    if (res.status === 401) throw Object.assign(new Error('서버 암호가 필요합니다'), { status: 401 });
    if (!res.ok) {
      let msg = `서버 오류 (${res.status})`;
      try { msg = (await res.json()).error ?? msg; } catch { /* 무시 */ }
      throw new Error(msg);
    }
    return res.json();
  },

  list() { return this.request('files'); },
  load(name) { return this.request(`files/${encodeURIComponent(name)}`); },
  save(name, data) { return this.request(`files/${encodeURIComponent(name)}`, { method: 'PUT', body: JSON.stringify(data) }); },
  remove(name) { return this.request(`files/${encodeURIComponent(name)}`, { method: 'DELETE' }); },
};
