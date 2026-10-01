import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { server } from '../src/storage.js';

test('서버 클라이언트: 웹 가져오기 인증·재시도와 HTTP 오류 상태 보존', async () => {
  const values = new Map();
  const storage = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  const realFetch = globalThis.fetch;
  const http = createServer((req, res) => {
    res.setHeader('Content-Type', 'application/json');
    if (req.url === '/api/health') return res.end(JSON.stringify({ ok: true, auth: true }));
    if (req.url.startsWith('/api/published/')) { res.statusCode = 404; return res.end(JSON.stringify({ error: '게시 문서가 없습니다' })); }
    if (req.headers['x-tabula-token'] !== 'test-token') {
      res.statusCode = 401; return res.end(JSON.stringify({ error: '서버 암호가 필요합니다' }));
    }
    if (req.url.startsWith('/api/fetch?')) { res.setHeader('Content-Type', 'text/plain'); return res.end('이름,합계\n합성,15'); }
    res.statusCode = 413; res.end(JSON.stringify({ error: '문서가 너무 큽니다' }));
  });
  await new Promise(resolve => http.listen(0, '127.0.0.1', resolve));
  const base = `http://127.0.0.1:${http.address().port}/`;
  Object.defineProperty(globalThis, 'localStorage', { configurable: true, value: {
    getItem: key => values.get(key) ?? null,
    setItem: (key, value) => values.set(key, value),
  } });
  globalThis.fetch = (url, options) => realFetch(new URL(url, base), options);
  try {
    assert.equal(await server.init(), true);
    assert.equal(server.needsToken, true);
    await assert.rejects(server.fetchText('https://example.com/data.csv'), { status: 401, message: '서버 암호가 필요합니다' });
    server.setToken('test-token');
    assert.equal(await server.fetchText('https://example.com/data.csv'), '이름,합계\n합성,15');
    await assert.rejects(server.save('합성 문서', {}), { status: 413, message: '문서가 너무 큽니다' });
    await assert.rejects(server.published('missing'), { status: 404, message: '게시 문서가 없습니다' });
    globalThis.fetch = async () => { throw new Error('오프라인'); };
    assert.equal(await server.init(), false);
    assert.equal(server.needsToken, false);
  } finally {
    globalThis.fetch = realFetch;
    if (storage) Object.defineProperty(globalThis, 'localStorage', storage); else delete globalThis.localStorage;
    server.available = false; server.needsToken = false;
    http.closeAllConnections();
    await new Promise(resolve => http.close(resolve));
  }
});
