// 설치용 아이콘·매니페스트를 격리 Chromium에서 검사한다. 실제 OS 홈 화면 설치는 수행하지 않는다.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const origin = new URL(url).origin;
const assets = [
  { file: 'favicon.ico', mime: /^image\/(?:vnd\.microsoft\.icon|x-icon)$/ },
  { file: 'favicon-32x32.png', mime: /^image\/png$/, size: 32 },
  { file: 'apple-touch-icon.png', mime: /^image\/png$/, size: 180 },
  { file: 'icons/wixel-192.png', mime: /^image\/png$/, size: 192 },
  { file: 'icons/wixel-512.png', mime: /^image\/png$/, size: 512 },
  { file: 'icons/wixel-maskable-512.png', mime: /^image\/png$/, size: 512 },
];
const local = async file => readFile(new URL('../' + file, import.meta.url));
const sha256 = bytes => createHash('sha256').update(bytes).digest('hex');
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 390, height: 844 }, hasTouch: true });
const page = await context.newPage(), errors = [], writes = [], blocked = [], results = [], http = [], images = [];
let manifest, dom, parsed, installability;
page.on('pageerror', error => errors.push(error.message));
page.setDefaultTimeout(15000);
await context.route('**/*', route => {
  const request = route.request(), u = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push({ method: request.method(), url: u.href }); return route.abort(); }
  if (u.origin !== origin || u.pathname.startsWith('/api/')) { blocked.push(u.href); return route.abort(); }
  return route.continue();
});
await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
async function check(name, fn) {
  try { await fn(); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (error) { results.push({ name, ok: false, error: error.message }); console.error('NG ' + name + ': ' + error.stack); }
}
async function asset(file, mime) {
  const target = new URL(file, url); assert.equal(target.origin, origin); assert.ok(!target.pathname.startsWith('/api/'));
  const response = await context.request.get(target.href, { timeout: 30000 });
  assert.equal(response.status(), 200, file + ' HTTP');
  const type = response.headers()['content-type']?.split(';')[0].trim().toLowerCase();
  assert.match(type || '', mime, file + ' MIME');
  const bytes = await response.body(), expected = await local(file);
  assert.equal(sha256(bytes), sha256(expected), file + ' 배포 바이트가 로컬 원본과 같아야 함');
  http.push({ file, status: response.status(), contentType: type, bytes: bytes.length, sha256: sha256(bytes) });
  return bytes;
}
async function imageStats(src, maskable = false) {
  return page.evaluate(async ({ src, maskable }) => {
    const image = new Image(); image.src = src; await image.decode();
    const width = image.naturalWidth, height = image.naturalHeight;
    const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
    const ctx = canvas.getContext('2d', { willReadFrequently: true }); ctx.drawImage(image, 0, 0);
    const pixels = ctx.getImageData(0, 0, width, height).data;
    let transparent = 0, mark = 0, white = 0, mint = 0, outside = 0, maxRadius = 0, sx = 0, sy = 0;
    const radius = Math.min(width, height) * 0.4;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i + 3] !== 255) transparent++;
      if (!maskable) continue;
      const r = pixels[i], g = pixels[i + 1], b = pixels[i + 2];
      // WIXEL의 흰색/민트 전경을 짙은 녹색 배경에서 분리한다. 일반 이미지 의미 인식 검사가 아니다.
      if (r > 90 && g > 140 && b > 110 && (r + g + b) / 3 > 150) {
        const x = (i / 4) % width + 0.5, y = Math.floor(i / 4 / width) + 0.5;
        const distance = Math.hypot(x - width / 2, y - height / 2);
        mark++; sx += x; sy += y; maxRadius = Math.max(maxRadius, distance);
        if (distance > radius) outside++;
        if (Math.min(r, g, b) > 225) white++;
        if (g - r > 15 && b - r > 5 && g > 180) mint++;
      }
    }
    return { width, height, transparent, ...(maskable ? { mark, white, mint, outside, radius, maxRadius, centroid: { x: sx / mark, y: sy / mark } } : {}) };
  }, { src, maskable });
}
try {
  await check('앱 부팅과 DOM favicon·Apple·manifest 링크', async () => {
    const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); assert.equal(response.status(), 200);
    await page.waitForFunction(() => !!window.tabula?.wb());
    dom = await page.evaluate(() => ({
      icons: [...document.querySelectorAll('link[rel="icon"]')].map(e => ({ href: e.href, sizes: e.sizes.value, type: e.type })),
      apple: [...document.querySelectorAll('link[rel="apple-touch-icon"]')].map(e => ({ href: e.href, sizes: e.sizes.value })),
      manifests: [...document.querySelectorAll('link[rel="manifest"]')].map(e => e.href),
      meta: Object.fromEntries([...document.querySelectorAll('meta[name]')].map(e => [e.name, e.content])),
    }));
    assert.equal(dom.icons.length, 2); assert.equal(dom.apple.length, 1); assert.equal(dom.manifests.length, 1);
    assert.ok(dom.icons.some(x => x.type === 'image/png' && x.sizes === '32x32'));
    assert.ok(dom.icons.some(x => /icon$/.test(x.type) && x.sizes === '16x16 32x32 48x48'));
    assert.equal(dom.apple[0].sizes, '180x180');
    assert.equal(dom.apple[0].href, new URL('apple-touch-icon.png', url).href);
    assert.equal(dom.manifests[0], new URL('manifest.webmanifest', url).href);
  });
  await check('Apple 홈 화면 앱 이름·독립 실행·테마 메타', async () => {
    assert.equal(dom.meta['apple-mobile-web-app-capable'], 'yes');
    assert.equal(dom.meta['apple-mobile-web-app-title'], 'WIXEL');
    assert.equal(dom.meta['apple-mobile-web-app-status-bar-style'], 'default');
    assert.equal(dom.meta['theme-color'], '#185b46');
  });
  await check('manifest HTTP MIME·원본 해시', async () => {
    const bytes = await asset('manifest.webmanifest', /^application\/manifest\+json$/); manifest = JSON.parse(bytes.toString('utf8'));
  });
  await check('Chromium Page.getAppManifest 파싱 오류 없음', async () => {
    const cdp = await context.newCDPSession(page); await cdp.send('Page.enable');
    const response = await cdp.send('Page.getAppManifest');
    assert.deepEqual(response.errors, []); assert.equal(response.url, new URL('manifest.webmanifest', url).href);
    assert.equal(JSON.parse(response.data).short_name, 'WIXEL');
    parsed = { url: response.url, errors: response.errors, manifest: response.manifest || response.parsed || null };
    installability = await cdp.send('Page.getInstallabilityErrors').catch(error => ({ unavailable: error.message }));
    await cdp.detach();
  });
  await check('앱 식별자·시작 URL·scope·standalone 범위', async () => {
    assert.equal(manifest.short_name, 'WIXEL'); assert.match(manifest.name, /WIXEL/); assert.equal(manifest.lang, 'ko-KR');
    assert.equal(manifest.display, 'standalone'); assert.equal(manifest.theme_color, dom.meta['theme-color']);
    const scope = new URL(manifest.scope, dom.manifests[0]), start = new URL(manifest.start_url, dom.manifests[0]), id = new URL(manifest.id, dom.manifests[0]);
    for (const target of [scope, start, id]) { assert.equal(target.origin, origin); assert.equal(target.search, ''); assert.equal(target.hash, ''); }
    assert.ok(start.pathname.startsWith(scope.pathname)); assert.equal(id.href, scope.href);
  });
  await check('192·512 일반 아이콘과 전용 maskable 선언', async () => {
    assert.equal(manifest.icons.length, 3);
    for (const [src, sizes, purpose] of [['icons/wixel-192.png', '192x192', 'any'], ['icons/wixel-512.png', '512x512', 'any'], ['icons/wixel-maskable-512.png', '512x512', 'maskable']]) {
      const icon = manifest.icons.find(x => x.src === src); assert.ok(icon); assert.equal(icon.sizes, sizes); assert.equal(icon.type, 'image/png'); assert.equal(icon.purpose, purpose);
    }
  });
  await check('아이콘 6개 HTTP MIME·SHA256 원본 일치', async () => { for (const a of assets) await asset(a.file, a.mime); });
  await check('PNG 5개 실제 디코딩 크기·완전 불투명 배경', async () => {
    for (const a of assets.filter(a => a.size)) {
      const stats = await imageStats(new URL(a.file, url).href, a.file.includes('maskable'));
      images.push({ file: a.file, ...stats }); assert.equal(stats.width, a.size); assert.equal(stats.height, a.size); assert.equal(stats.transparent, 0, a.file + ' 투명 픽셀');
    }
  });
  await check('maskable 흰색·민트 핵심 마크가 반지름40% 안전원 안에 있음', async () => {
    const stats = images.find(x => x.file.includes('maskable')); assert.ok(stats);
    assert.ok(stats.white > 1000 && stats.mint > 1000 && stats.mark > 10000, '밝은 전경이 존재해야 함');
    assert.equal(stats.outside, 0); assert.ok(stats.maxRadius <= stats.radius);
    assert.ok(Math.abs(stats.centroid.x - 256) < 40 && Math.abs(stats.centroid.y - 256) < 40, '마크 중심');
  });
  await check('ICO 16·32·48 프레임과 브라우저 favicon 디코딩', async () => {
    const bytes = await local('favicon.ico'); assert.equal(bytes.readUInt16LE(0), 0); assert.equal(bytes.readUInt16LE(2), 1); assert.equal(bytes.readUInt16LE(4), 3);
    const dimensions = [];
    for (let n = 0; n < 3; n++) {
      const offset = 6 + n * 16, width = bytes[offset] || 256, height = bytes[offset + 1] || 256;
      const length = bytes.readUInt32LE(offset + 8), start = bytes.readUInt32LE(offset + 12);
      assert.equal(width, height); assert.ok(length > 0 && start >= 54 && start + length <= bytes.length); dimensions.push(width);
    }
    assert.deepEqual(dimensions.sort((a, b) => a - b), [16, 32, 48]);
    for (const icon of dom.icons) {
      // Playwright request routing suppresses the bare /favicon.ico request internally.
      // The original URL's MIME/hash were checked above; an explicit query keeps API blocking active.
      const src = new URL(icon.href); if (src.protocol !== 'data:' && src.pathname.endsWith('/favicon.ico')) src.searchParams.set('wixel-icon-check', '1');
      const stats = await imageStats(src.href); assert.ok([16, 32, 48].includes(stats.width)); assert.equal(stats.width, stats.height);
    }
  });
  await check('오류·원격 문서 쓰기 없음', async () => { assert.deepEqual(errors, []); assert.deepEqual(writes, []); });
} finally {
  await browser.close();
  console.log(JSON.stringify({ url, total: results.length, good: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok).length,
    pageErrors: errors, remoteWrites: writes.length, blockedRequests: [...new Set(blocked)], http, images, parsed, installability,
    limits: '실제 OS 설치/홈 화면 추가는 미실행. maskable은 흰색·민트 전경 색상 분리 기반 정적 안전원 검사.', results }, null, 2));
  if (results.some(x => !x.ok)) process.exitCode = 1;
}
