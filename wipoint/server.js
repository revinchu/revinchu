// 개발용 정적 서버 (의존성 없음): node server.js → http://localhost:5179
// 문서는 브라우저(IndexedDB)에 저장되므로 서버 저장 API 는 없습니다.
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT ?? 5179);
const HOST = process.env.HOST ?? '127.0.0.1';
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.gz': 'application/gzip', '.svg': 'image/svg+xml', '.png': 'image/png', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };

createServer(async (req, res) => {
  try {
    let path = decodeURIComponent(new URL(req.url, 'http://x').pathname);
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
