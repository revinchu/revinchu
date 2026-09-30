// WIXEL 서버: 정적 파일 + 통합 문서 저장 API
//   node server.js                → http://localhost:5178 (같은 네트워크의 다른 기기에서도 접속 가능)
//   PORT=8080 HOST=127.0.0.1      → 포트/주소 변경 (HOST=127.0.0.1 이면 이 컴퓨터에서만 접속)
//   TABULA_DATA=/경로             → 저장 폴더 (기본: ./data)
//   TABULA_TOKEN=비밀값           → 설정하면 저장 API 사용 시 이 값이 필요
import { createServer } from 'node:http';
import { readFile, writeFile, readdir, stat, unlink, mkdir, rename } from 'node:fs/promises';
import { extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { networkInterfaces } from 'node:os';
import { timingSafeEqual, randomBytes, createHmac } from 'node:crypto';

const ROOT = fileURLToPath(new URL('.', import.meta.url));
const PORT = Number(process.env.PORT) || 5178;
const HOST = process.env.HOST || '0.0.0.0';
const DATA = resolve(process.env.TABULA_DATA || join(ROOT, 'data'));
const TOKEN = process.env.TABULA_TOKEN || '';
const MAX_BODY = 50 * 1024 * 1024;
const TYPES = {
  '.gz': 'application/gzip',
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.json': 'application/json; charset=utf-8',
};
const SUFFIX = '.tabula.json';

/** 문서 이름 → 안전한 파일 이름 (경로 문자 차단) */
function fileFor(name) {
  const n = String(name ?? '').trim();
  if (!n || n.length > 120 || /[\u0000-\u001f]/.test(n)) return null;
  return join(DATA, encodeURIComponent(n) + SUFFIX);
}

function send(res, status, body, type = 'application/json; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
}

function authorized(req) {
  if (!TOKEN) return true;
  const given = Buffer.from(String(req.headers['x-tabula-token'] ?? ''));
  const want = Buffer.from(TOKEN);
  return given.length === want.length && timingSafeEqual(given, want);
}

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_BODY) throw Object.assign(new Error('too large'), { status: 413 });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf8');
}

const PUB = join(DATA, 'published');
const pubFile = (id) => (/^[a-z0-9]{8,32}$/.test(id) ? join(PUB, `${id}.json`) : null);

async function api(req, res, path) {
  if (path === '/api/health') return send(res, 200, { ok: true, auth: !!TOKEN, publish: true });
  // 웹에 게시한 문서: 링크를 아는 사람은 누구나 읽기 전용으로 봄 (암호 없이 GET)
  const pm = /^\/api\/published\/([a-z0-9]+)$/.exec(path);
  if (pm && req.method === 'GET') {
    const f = pubFile(pm[1]);
    try {
      const [body, st] = await Promise.all([readFile(f), stat(f)]);
      res.setHeader('X-Modified', String(st.mtimeMs));
      return send(res, 200, body);
    } catch { return send(res, 404, { error: '게시가 중지되었거나 없는 문서입니다' }); }
  }
  // 네이버 검색광고 API 중계 (연관검색어 템플릿): 브라우저는 CORS 때문에 직접 부를 수 없어 서버가 서명 · 요청
  //   GET /api/naver/keywordstool?hintKeywords=…&month=&event=  헤더 X-Customer / X-API-KEY / X-Secret
  if (path === '/api/naver/keywordstool' && req.method === 'GET') {
    const q = new URL(req.url, 'http://x').searchParams;
    const customer = req.headers['x-customer'];
    const key = req.headers['x-api-key'];
    const secret = req.headers['x-secret'];
    if (!customer || !key || !secret || !q.get('hintKeywords')) return send(res, 400, { error: '계정 ID · 액세스 라이선스 · 비밀 키 · 키워드를 모두 입력하세요' });
    const ts = String(Date.now());
    const sig = createHmac('sha256', String(secret)).update(`${ts}.GET./keywordstool`).digest('base64');
    const url = new URL('https://api.searchad.naver.com/keywordstool');
    url.searchParams.set('hintKeywords', q.get('hintKeywords'));
    url.searchParams.set('showDetail', '1');
    for (const k of ['month', 'event']) if (q.get(k)) url.searchParams.set(k, q.get(k));
    try {
      const r = await fetch(url, { headers: { 'X-Timestamp': ts, 'X-API-KEY': String(key), 'X-Customer': String(customer), 'X-Signature': sig } });
      const text = await r.text();
      res.writeHead(r.status, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(text);
    } catch (e) { return send(res, 502, { error: `네이버 API 에 연결하지 못했습니다: ${e.message}` }); }
    return undefined;
  }
  if (!authorized(req)) return send(res, 401, { error: '인증이 필요합니다' });
  // 웹 가져오기 함수 중계 (IMPORTXML · IMPORTHTML · IMPORTDATA · IMPORTFEED · GOOGLEFINANCE · GOOGLETRANSLATE …)
  //   GET /api/fetch?url=…  — 공개 http(s) 주소만 (내부망 · localhost 막음), 15초 · 5MB 제한
  if (path === '/api/fetch' && req.method === 'GET') {
    const target = new URL(req.url, 'http://x').searchParams.get('url') ?? '';
    let u;
    try { u = new URL(target); } catch { return send(res, 400, { error: '잘못된 주소입니다' }); }
    if (!/^https?:$/.test(u.protocol) || privateHost(u.hostname)) return send(res, 400, { error: '공개 웹 주소(http · https)만 가져올 수 있습니다' });
    try {
      const r = await fetch(u, { redirect: 'follow', signal: AbortSignal.timeout(15000), headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36 WIXEL', 'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8' } });
      if (privateHost(new URL(r.url).hostname)) return send(res, 400, { error: '내부 주소로 이동하는 링크입니다' });
      const buf = Buffer.from(await r.arrayBuffer());
      if (buf.length > 5 * 1024 * 1024) return send(res, 413, { error: '5MB 보다 큰 응답은 가져오지 않습니다' });
      const type = r.headers.get('content-type') ?? '';
      const cs = /charset=([\w-]+)/i.exec(type)?.[1] ?? /<meta[^>]+charset=["']?([\w-]+)/i.exec(buf.subarray(0, 4096).toString('latin1'))?.[1] ?? 'utf-8';
      let text;
      try { text = new TextDecoder(cs.toLowerCase() === 'ks_c_5601-1987' ? 'euc-kr' : cs).decode(buf); } catch { text = buf.toString('utf8'); }
      res.writeHead(r.ok ? 200 : 502, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' });
      res.end(r.ok ? text : `HTTP ${r.status}`);
    } catch (e) { return send(res, 502, { error: `가져오지 못했습니다: ${e.message}` }); }
    return undefined;
  }
  if (path === '/api/publish' && req.method === 'POST') {
    await mkdir(PUB, { recursive: true });
    const body = await readBody(req);
    try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
    const id = randomBytes(9).toString('hex');
    await writeFile(pubFile(id), body);
    return send(res, 200, { ok: true, id });
  }
  if (pm && (req.method === 'PUT' || req.method === 'DELETE')) {
    const f = pubFile(pm[1]);
    if (!f) return send(res, 400, { error: '잘못된 게시 id' });
    if (req.method === 'DELETE') { try { await unlink(f); } catch { /* 없음 */ } return send(res, 200, { ok: true }); }
    const body = await readBody(req);
    try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
    await mkdir(PUB, { recursive: true });
    await writeFile(`${f}.tmp`, body);
    await rename(`${f}.tmp`, f);
    return send(res, 200, { ok: true });
  }
  if (!authorized(req)) return send(res, 401, { error: '인증이 필요합니다' });
  await mkdir(DATA, { recursive: true });
  if (path === '/api/files' && req.method === 'GET') {
    const names = (await readdir(DATA)).filter((f) => f.endsWith(SUFFIX));
    const list = await Promise.all(names.map(async (f) => {
      const st = await stat(join(DATA, f));
      return { name: decodeURIComponent(f.slice(0, -SUFFIX.length)), size: st.size, modified: st.mtimeMs };
    }));
    list.sort((a, b) => b.modified - a.modified);
    return send(res, 200, list);
  }
  const m = /^\/api\/files\/(.+)$/.exec(path);
  if (!m) return send(res, 404, { error: 'not found' });
  const file = fileFor(decodeURIComponent(m[1]));
  if (!file) return send(res, 400, { error: '파일 이름이 올바르지 않습니다' });
  if (req.method === 'GET') {
    try {
      const [body, st] = await Promise.all([readFile(file), stat(file)]);
      res.setHeader('X-Modified', String(st.mtimeMs));
      return send(res, 200, body);
    } catch {
      return send(res, 404, { error: '파일이 없습니다' });
    }
  }
  if (req.method === 'PUT') {
    const body = await readBody(req);
    try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
    const tmp = `${file}.${process.pid}.tmp`;
    await writeFile(tmp, body);
    await rename(tmp, file);
    const st = await stat(file);
    return send(res, 200, { ok: true, modified: st.mtimeMs });
  }
  if (req.method === 'DELETE') {
    try { await unlink(file); } catch { /* 이미 없음 */ }
    return send(res, 200, { ok: true });
  }
  return send(res, 405, { error: 'method not allowed' });
}

/** 내부망 · 자기 자신 주소 (웹 가져오기 중계에서 막음) */
function privateHost(h) {
  const x = h.toLowerCase().replace(/^\[|\]$/g, '');
  if (x === 'localhost' || x.endsWith('.localhost') || x.endsWith('.local') || x.endsWith('.internal') || x === '0.0.0.0' || x === '::' || x === '::1') return true;
  const m = /^(\d+)\.(\d+)\.(\d+)\.(\d+)$/.exec(x);
  if (m) {
    const [a, b] = [+m[1], +m[2]];
    return a === 10 || a === 127 || a === 0 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || (a === 100 && b >= 64 && b <= 127);
  }
  return /^(fc|fd|fe80)/.test(x);
}

createServer(async (req, res) => {
  try {
    const rawPath = new URL(req.url, 'http://x').pathname;
    if (rawPath.startsWith('/api/')) { await api(req, res, rawPath); return; }
    const path = decodeURIComponent(rawPath);
    const file = normalize(join(ROOT, path === '/' ? 'index.html' : path));
    if (!file.startsWith(ROOT) || file.startsWith(DATA)) { res.writeHead(403).end(); return; }
    const body = await readFile(file);
    res.writeHead(200, { 'Content-Type': TYPES[extname(file)] || 'application/octet-stream' });
    res.end(body);
  } catch (e) {
    if (res.headersSent) return;
    if (e.status) send(res, e.status, { error: e.message });
    else if (e.code === 'ENOENT' || e.code === 'EISDIR') { res.writeHead(404).end('Not found'); }
    else send(res, 500, { error: '서버 오류' });
  }
}).listen(PORT, HOST, () => {
  console.log(`WIXEL: http://localhost:${PORT}`);
  if (HOST === '0.0.0.0') {
    for (const addrs of Object.values(networkInterfaces())) {
      for (const a of addrs ?? []) if (a.family === 'IPv4' && !a.internal) console.log(`  다른 기기에서: http://${a.address}:${PORT}`);
    }
  }
  console.log(`  저장 폴더: ${DATA}${TOKEN ? ' (토큰 필요)' : ''}`);
});
