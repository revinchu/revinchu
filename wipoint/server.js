// 개발용 정적 서버 (의존성 없음): node server.js → http://localhost:5179
// 문서는 브라우저(IndexedDB)에 저장되므로 서버 저장 API 는 없습니다.
// 공동 편집 중계: /api/collab/<방>/events (SSE) · POST /api/collab/<방> — 메시지를 같은 방의 다른 사람에게 그대로 전달
//   (마지막 'state' 이후의 patch 를 메모리에 두어 새로 들어온 사람에게 다시 보냄). 다른 기기에서 쓰려면 HOST=0.0.0.0
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT ?? 5179);
const HOST = process.env.HOST ?? '127.0.0.1';
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.gz': 'application/gzip', '.svg': 'image/svg+xml', '.png': 'image/png', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };

const rooms = new Map(); // 방 → { clients: Map(id → res), log: [메시지 문자열] }
const room = (id) => { if (!rooms.has(id)) rooms.set(id, { clients: new Map(), log: [] }); return rooms.get(id); };
function collab(req, res, url) {
  const m = /^\/api\/collab\/([\w-]{4,64})(\/events)?$/.exec(url.pathname);
  if (url.pathname === '/api/collab/ping') { res.writeHead(200, { 'Content-Type': 'application/json' }).end('{"ok":true}'); return true; }
  if (!m) return false;
  const r = room(m[1]);
  const client = url.searchParams.get('c') ?? '';
  if (m[2] && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
    res.write(': hi\n\n');
    for (const msg of r.log) res.write(`data: ${msg}\n\n`);
    r.clients.set(client, res);
    const ping = setInterval(() => res.write(': ping\n\n'), 20000);
    req.on('close', () => {
      clearInterval(ping);
      if (r.clients.get(client) === res) r.clients.delete(client);
      for (const [, other] of r.clients) other.write(`data: ${JSON.stringify({ type: 'bye', from: client })}\n\n`);
      if (!r.clients.size) setTimeout(() => { if (!r.clients.size) rooms.delete(m[1]); }, 10 * 60 * 1000);
    });
    return true;
  }
  if (!m[2] && req.method === 'POST') {
    let body = '';
    let size = 0;
    req.on('data', (d) => { size += d.length; if (size < 64 * 1024 * 1024) body += d; });
    req.on('end', () => {
      let msg;
      try { msg = JSON.parse(body); } catch { res.writeHead(400).end(); return; }
      const line = JSON.stringify(msg);
      if (msg.type === 'state') r.log = [line];
      else if (msg.type === 'patch') { r.log.push(line); if (r.log.length > 2000) r.log.splice(1, r.log.length - 2000); }
      for (const [id, other] of r.clients) if (id !== msg.from) other.write(`data: ${line}\n\n`);
      res.writeHead(204).end();
    });
    return true;
  }
  res.writeHead(405).end();
  return true;
}

createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://x');
    if (url.pathname.startsWith('/api/collab/') && collab(req, res, url)) return;
    let path = decodeURIComponent(url.pathname);
    if (path === '/') path = '/index.html';
    let file = normalize(join(root, path));
    if (!file.startsWith(root)) { res.writeHead(403).end(); return; }
    // 아이콘 모음은 WIXEL 폴더의 것을 함께 씀
    if (path.startsWith('/assets/iconlib') && !(await stat(file).catch(() => null))) file = join(root, '..', 'tabula', path);
    const body = await readFile(file);
    res.writeHead(200, { 'Content-Type': TYPES[extname(file)] ?? 'application/octet-stream', 'Cache-Control': 'no-cache' });
    res.end(body);
  } catch {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('없음');
  }
}).listen(PORT, HOST, () => console.log(`WIPOINT: http://${HOST === '0.0.0.0' ? 'localhost' : HOST}:${PORT}`));
