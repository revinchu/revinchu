// Node 서버 전용 HTTP 앱. import 시에는 포트를 열지 않습니다.
import { createServer } from 'node:http';
import { readFile, writeFile, readdir, stat, unlink, mkdir, rename, realpath } from 'node:fs/promises';
import { extname, join, resolve, relative, isAbsolute, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { timingSafeEqual, randomBytes, createHmac } from 'node:crypto';
import { fetchPublicResource, isLoopbackAddress, ProxyError } from './network.js';

const ROOT = fileURLToPath(new URL('../', import.meta.url));
const SUFFIX = '.tabula.json';
const TYPES = {
  '.gz': 'application/gzip', '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.json': 'application/json; charset=utf-8',
};

export function serverConfig(env = process.env) {
  let host = String(env.HOST || '127.0.0.1').trim();
  if (host.toLowerCase() === 'localhost') host = '127.0.0.1';
  host = host.replace(/^\[|\]$/g, '');
  const token = String(env.TABULA_TOKEN || '');
  const port = env.PORT === undefined || env.PORT === '' ? 5178 : Number(env.PORT);
  if (!Number.isInteger(port) || port < 0 || port > 65535) throw new Error('PORT는 0~65535의 정수여야 합니다');
  if (!isLoopbackAddress(host) && !token.trim()) throw new Error('외부 접속 주소에서는 TABULA_TOKEN을 설정해야 서버를 시작할 수 있습니다');
  const proxyRateLimit = env.TABULA_PROXY_RATE_LIMIT === undefined ? 120 : Number(env.TABULA_PROXY_RATE_LIMIT);
  if (!Number.isInteger(proxyRateLimit) || proxyRateLimit < 1) throw new Error('TABULA_PROXY_RATE_LIMIT는 1 이상의 정수여야 합니다');
  return { host, port, token, data: resolve(env.TABULA_DATA || join(ROOT, 'data')), proxyRateLimit };
}

function within(base, file) {
  const rel = relative(base, file);
  return rel === '' || (!isAbsolute(rel) && rel !== '..' && !rel.startsWith('..' + sep));
}

function allowedAsset(asset) {
  return !asset.split('/').some((part) => part.startsWith('.')) &&
    /^(index\.html|styles\.css|src\/.+\.js|assets\/.+|dist\/(index\.html|assets\/.+))$/.test(asset);
}

function send(res, status, body, type = 'application/json; charset=utf-8') {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
}

function sameOrigin(req) {
  if (!req.headers.origin) return true;
  try {
    const origin = new URL(req.headers.origin);
    return ['http:', 'https:'].includes(origin.protocol) && origin.host.toLowerCase() === String(req.headers.host).toLowerCase();
  } catch { return false; }
}

function localRequest(req) {
  try {
    const host = new URL('http://' + req.headers.host).hostname;
    return isLoopbackAddress(req.socket.remoteAddress) && (host === 'localhost' || isLoopbackAddress(host));
  } catch { return false; }
}

async function atomicWrite(file, body) {
  const temp = `${file}.${process.pid}.${randomBytes(8).toString('hex')}.tmp`;
  try { await writeFile(temp, body, { mode: 0o600 }); await rename(temp, file); }
  finally { await unlink(temp).catch(() => {}); }
}

export function createWixelServer(options = {}) {
  const { host = '127.0.0.1', token = '', data = join(ROOT, 'data'), root = ROOT,
    fetchResource = fetchPublicResource, proxyRateLimit = 120 } = options;
  if (!isLoopbackAddress(host) && !String(token).trim()) throw new Error('외부 접속 주소에서는 TABULA_TOKEN을 설정해야 서버를 시작할 수 있습니다');
  const DATA = resolve(data), PUB = join(DATA, 'published'), staticRoot = resolve(root);
  const wantedToken = Buffer.from(token);
  const rateBuckets = new Map();
  const pendingWrites = new Map();

  function writeDocument(file, body) {
    // Windows에서 같은 대상의 동시 rename이 EPERM을 내지 않도록 문서별 저장을 직렬화합니다.
    const previous = pendingWrites.get(file) ?? Promise.resolve();
    const next = previous.catch(() => {}).then(() => atomicWrite(file, body));
    pendingWrites.set(file, next);
    return next.finally(() => { if (pendingWrites.get(file) === next) pendingWrites.delete(file); });
  }

  function authorized(req) {
    if (!wantedToken.length) return localRequest(req);
    const given = Buffer.from(String(req.headers['x-tabula-token'] ?? ''));
    return given.length === wantedToken.length && timingSafeEqual(given, wantedToken);
  }

  function allowProxy(req, res) {
    const now = Date.now(), key = req.socket.remoteAddress ?? 'unknown';
    let bucket = rateBuckets.get(key);
    if (!bucket || bucket.until <= now) {
      if (rateBuckets.size >= 10000) {
        for (const [address, item] of rateBuckets) if (item.until <= now) rateBuckets.delete(address);
        if (rateBuckets.size >= 10000 && !rateBuckets.has(key)) {
          res.setHeader('Retry-After', '60'); send(res, 429, { error: '요청이 많습니다. 잠시 후 다시 시도하세요' }); return false;
        }
      }
      bucket = { count: 0, until: now + 60000 }; rateBuckets.set(key, bucket);
    }
    if (++bucket.count <= proxyRateLimit) return true;
    res.setHeader('Retry-After', String(Math.ceil((bucket.until - now) / 1000)));
    send(res, 429, { error: '웹 중계 요청 한도를 초과했습니다. 잠시 후 다시 시도하세요' }); return false;
  }

  function fileFor(name) {
    const n = String(name ?? '').trim();
    if (!n || n.length > 120 || /[\u0000-\u001f]/.test(n) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(n)) return null;
    return join(DATA, encodeURIComponent(n) + SUFFIX);
  }
  const pubFile = (id) => (/^[a-z0-9]{8,32}$/.test(id) ? join(PUB, `${id}.json`) : null);

  async function readBody(req) {
    const chunks = []; let size = 0;
    for await (const chunk of req) {
      size += chunk.length;
      if (size > 50 * 1024 * 1024) throw new ProxyError('문서는 50MB 이하만 저장할 수 있습니다', 413);
      chunks.push(chunk);
    }
    return Buffer.concat(chunks).toString('utf8');
  }

  async function api(req, res, path) {
    if (path === '/api/health' && req.method === 'GET') return send(res, 200, { ok: true, auth: !!wantedToken.length, publish: true });
    const pm = /^\/api\/published\/([a-z0-9]+)$/.exec(path);
    // 사용자가 게시한 링크만 의도적으로 공개합니다. 원본 문서/게시 수정은 아래 인증을 거칩니다.
    if (pm && req.method === 'GET') {
      const file = pubFile(pm[1]);
      if (!file) return send(res, 404, { error: '게시가 중지되었거나 없는 문서입니다' });
      try {
        const [body, info] = await Promise.all([readFile(file), stat(file)]);
        res.setHeader('X-Modified', String(info.mtimeMs)); return send(res, 200, body);
      } catch { return send(res, 404, { error: '게시가 중지되었거나 없는 문서입니다' }); }
    }
    if (!authorized(req)) return send(res, 401, { error: '인증이 필요합니다' });
    if (!sameOrigin(req)) return send(res, 403, { error: '다른 웹사이트에서 보낸 요청은 허용하지 않습니다' });

    if (path === '/api/naver/keywordstool' && req.method === 'GET') {
      if (!allowProxy(req, res)) return;
      const query = new URL(req.url, 'http://x').searchParams;
      const customer = req.headers['x-customer'], key = req.headers['x-api-key'], secret = req.headers['x-secret'];
      if (!customer || !key || !secret || !query.get('hintKeywords')) return send(res, 400, { error: '계정 ID · 액세스 라이선스 · 비밀 키 · 키워드를 모두 입력하세요' });
      const ts = String(Date.now());
      const sig = createHmac('sha256', String(secret)).update(`${ts}.GET./keywordstool`).digest('base64');
      const url = new URL('https://api.searchad.naver.com/keywordstool');
      url.searchParams.set('hintKeywords', query.get('hintKeywords')); url.searchParams.set('showDetail', '1');
      for (const key of ['month', 'event']) if (query.get(key)) url.searchParams.set(key, query.get(key));
      const upstream = await fetchResource(url, {
        maxRedirects: 0, headers: { 'X-Timestamp': ts, 'X-API-KEY': String(key), 'X-Customer': String(customer), 'X-Signature': sig },
      });
      return send(res, upstream.status, upstream.body);
    }
    if (path === '/api/fetch' && req.method === 'GET') {
      if (!allowProxy(req, res)) return;
      const target = new URL(req.url, 'http://x').searchParams.get('url') ?? '';
      const upstream = await fetchResource(target);
      if (upstream.status < 200 || upstream.status >= 300) return send(res, 502, `HTTP ${upstream.status}`, 'text/plain; charset=utf-8');
      const buf = upstream.body, type = upstream.headers['content-type'] ?? '';
      const charset = /charset=["']?([\w-]+)/i.exec(type)?.[1] ?? /<meta[^>]+charset=["']?([\w-]+)/i.exec(buf.subarray(0, 4096).toString('latin1'))?.[1] ?? 'utf-8';
      let text;
      try { text = new TextDecoder(charset.toLowerCase() === 'ks_c_5601-1987' ? 'euc-kr' : charset).decode(buf); }
      catch { text = buf.toString('utf8'); }
      return send(res, 200, text, 'text/plain; charset=utf-8');
    }
    if (path === '/api/publish' && req.method === 'POST') {
      const body = await readBody(req);
      try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
      await mkdir(PUB, { recursive: true });
      const id = randomBytes(9).toString('hex'); await writeDocument(pubFile(id), body);
      return send(res, 200, { ok: true, id });
    }
    if (pm && (req.method === 'PUT' || req.method === 'DELETE')) {
      const file = pubFile(pm[1]);
      if (!file) return send(res, 400, { error: '잘못된 게시 ID입니다' });
      if (req.method === 'DELETE') { await unlink(file).catch((e) => { if (e.code !== 'ENOENT') throw e; }); return send(res, 200, { ok: true }); }
      const body = await readBody(req);
      try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
      await mkdir(PUB, { recursive: true }); await writeDocument(file, body);
      return send(res, 200, { ok: true });
    }
    if (path === '/api/files' && req.method === 'GET') {
      await mkdir(DATA, { recursive: true });
      const names = (await readdir(DATA)).filter((file) => file.endsWith(SUFFIX));
      const list = (await Promise.all(names.map(async (file) => {
        try {
          const info = await stat(join(DATA, file));
          return { name: decodeURIComponent(file.slice(0, -SUFFIX.length)), size: info.size, modified: info.mtimeMs };
        } catch { return null; }
      }))).filter(Boolean);
      list.sort((a, b) => b.modified - a.modified); return send(res, 200, list);
    }
    const match = /^\/api\/files\/(.+)$/.exec(path);
    if (!match) return send(res, 404, { error: '요청한 API가 없습니다' });
    let name;
    try { name = decodeURIComponent(match[1]); } catch { return send(res, 400, { error: '파일 이름이 올바르지 않습니다' }); }
    const file = fileFor(name);
    if (!file) return send(res, 400, { error: '파일 이름이 올바르지 않습니다' });
    if (req.method === 'GET') {
      try {
        const [body, info] = await Promise.all([readFile(file), stat(file)]);
        res.setHeader('X-Modified', String(info.mtimeMs)); return send(res, 200, body);
      } catch { return send(res, 404, { error: '파일이 없습니다' }); }
    }
    if (req.method === 'PUT') {
      const body = await readBody(req);
      try { JSON.parse(body); } catch { return send(res, 400, { error: 'JSON 형식이 아닙니다' }); }
      await mkdir(DATA, { recursive: true }); await writeDocument(file, body);
      const info = await stat(file); return send(res, 200, { ok: true, modified: info.mtimeMs });
    }
    if (req.method === 'DELETE') {
      await unlink(file).catch((e) => { if (e.code !== 'ENOENT') throw e; });
      return send(res, 200, { ok: true });
    }
    return send(res, 405, { error: '허용되지 않는 요청 방식입니다' });
  }

  const server = createServer(async (req, res) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Referrer-Policy', 'same-origin');
    res.setHeader('X-Frame-Options', 'SAMEORIGIN');
    res.setHeader('Content-Security-Policy', "frame-ancestors 'self'; object-src 'none'; base-uri 'self'");
    try {
      const path = new URL(req.url, 'http://x').pathname;
      if (path.startsWith('/api/')) return await api(req, res, path);
      if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, { error: '허용되지 않는 요청 방식입니다' });
      let decoded;
      try { decoded = decodeURIComponent(path).replace(/\\/g, '/'); } catch { return send(res, 400, { error: '잘못된 주소입니다' }); }
      const asset = decoded === '/' ? 'index.html' : decoded.slice(1);
      if (!allowedAsset(asset)) {
        return send(res, 403, { error: '접근할 수 없는 경로입니다' });
      }
      const file = await realpath(resolve(staticRoot, asset));
      const [realRoot, realData] = await Promise.all([realpath(staticRoot), realpath(DATA).catch(() => DATA)]);
      if (!within(realRoot, file) || within(realData, file) || !allowedAsset(relative(realRoot, file).replace(/\\/g, '/'))) {
        return send(res, 403, { error: '접근할 수 없는 경로입니다' });
      }
      const body = await readFile(file);
      res.writeHead(200, { 'Content-Type': TYPES[extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
      res.end(req.method === 'HEAD' ? undefined : body);
    } catch (error) {
      if (res.headersSent) { res.destroy(); return; }
      if (error.status) send(res, error.status, { error: error.message });
      else if (error.code === 'ENOENT' || error.code === 'EISDIR') send(res, 404, { error: '파일이 없습니다' });
      else send(res, 500, { error: '서버 오류' });
    }
  });
  server.requestTimeout = 30000;
  server.headersTimeout = 15000;
  server.keepAliveTimeout = 5000;
  return server;
}
