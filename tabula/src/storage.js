// 서버 저장소 (server.js 의 /api) 클라이언트. 정적 호스팅이면 available = false.
const TOKEN_KEY = 'tabula.serverToken';

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
