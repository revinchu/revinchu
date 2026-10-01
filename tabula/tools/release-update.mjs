// 새 배포 안내 회귀. 격리 컨텍스트·합성 문서·합성 manifest만 사용하며 파일/서버에 쓰지 않는다.
// 소스 서버에서는 빈 해시 모듈을 주입해 배포 표식을 재현한다. 번들 서버에서는 실제 표식을 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 서버에서만 실행하세요.');
const fakeAsset = 'wixel-0123456789abcdef.js', nextAsset = 'wixel-fedcba9876543210.js';
const browser = await chromium.launch(), results = [], errors = [], writes = [], requests = [];
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
const page = await context.newPage(); page.setDefaultTimeout(15000);
let currentAsset, mode = 'same', source = false, navigations = 0;
page.on('pageerror', e => errors.push(e.message));
page.on('framenavigated', frame => { if (frame === page.mainFrame()) navigations++; });
await context.route('**/*', async route => {
  const req = route.request(), requestUrl = new URL(req.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.url()); return route.abort(); }
  if (requestUrl.pathname.endsWith('/version.json')) {
    requests.push({ method: req.method(), url: req.url(), headers: req.headers(), body: req.postData() });
    if (mode === 'offline') return route.abort();
    if (mode === 'invalid') return route.fulfill({ contentType: 'application/json', body: JSON.stringify({ asset: '../outside.js' }) });
    return route.fulfill({ contentType: 'application/json', body: JSON.stringify({ asset: mode === 'new' ? nextAsset : currentAsset }) });
  }
  if (requestUrl.pathname.endsWith('/' + fakeAsset)) return route.fulfill({ contentType: 'text/javascript', body: '// 합성 배포 표식' });
  if (req.isNavigationRequest() && req.frame() === page.mainFrame()) {
    const response = await route.fetch(); let body = await response.text();
    source = body.includes('src="src/app.js"');
    if (source) body = body.replace('</head>', `<script type="module" src="${fakeAsset}"></script></head>`);
    return route.fulfill({ response, body });
  }
  return route.continue();
});
try {
  await page.addInitScript(() => {
    window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; window.__releaseSaveCalls = [];
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: async () => {
      window.__releaseSaveCalls.push({ active: navigator.userActivation.isActive });
      throw new DOMException('합성 선택기 취소', 'AbortError');
    } });
  });
  await page.clock.install();
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.wb());
  currentAsset = await page.evaluate(() => [...document.querySelectorAll('script[type="module"][src]')].map(s => s.src.split('/').at(-1)).find(s => /^wixel-[a-f0-9]{16}\.js$/.test(s)));
  assert.ok(currentAsset);
  if (await page.locator('#autosaveToggle').getAttribute('aria-checked') === 'true') await page.locator('#autosaveToggle').click();
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setInput(0, 0, 0, '879')); window.__releaseBook = w; });
  await page.clock.fastForward(14000);
  await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  assert.equal(requests.length, 0, '최초 15초 전 focus는 확인하지 않음');
  const firstResponse = page.waitForResponse(response => response.url().endsWith('/version.json'));
  await page.clock.fastForward(1000);
  await firstResponse;
  assert.equal(requests.length, 1); assert.equal(await page.locator('#releaseUpdate').count(), 0);
  results.push('15초 지연·같은 버전 숨김');
  await page.evaluate(() => window.dispatchEvent(new Event('focus'))); assert.equal(requests.length, 1);
  await page.clock.fastForward(299000); await page.evaluate(() => window.dispatchEvent(new Event('focus'))); assert.equal(requests.length, 1);
  results.push('5분 이내 focus 중복 확인 차단');
  mode = 'offline'; await page.clock.fastForward(1000); await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.waitForTimeout(100); assert.equal(requests.length, 2); assert.equal(await page.locator('#releaseUpdate').count(), 0);
  mode = 'invalid'; await page.clock.fastForward(300000); await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.waitForTimeout(100); assert.equal(requests.length, 3); assert.equal(await page.locator('#releaseUpdate').count(), 0);
  results.push('오프라인·잘못된 manifest 조용히 무시');
  mode = 'new'; await page.clock.fastForward(300000); await page.evaluate(() => window.dispatchEvent(new Event('focus')));
  await page.locator('#releaseUpdate').waitFor({ state: 'visible' });
  assert.equal(requests.length, 4); assert.equal(await page.locator('.dialog').count(), 0);
  assert.equal(await page.evaluate(() => window.__releaseSaveCalls.length), 0); assert.equal(navigations, 1);
  assert.equal(await page.evaluate(() => window.__releaseBook === tabula.wb() && tabula.wb().getValue(0, 0, 0) === 879), true);
  results.push('새 버전은 작은 버튼만 표시·문서와 포커스 작업 유지');
  await page.locator('#releaseUpdate').click();
  const dialog = page.getByRole('dialog', { name: '새 버전 안내' });
  assert.match(await dialog.innerText(), /현재 작업을 파일로 저장한 뒤 새로고침하면 적용됩니다/);
  await dialog.getByRole('button', { name: '닫기', exact: true }).last().click();
  assert.equal(await page.locator('#releaseUpdate').count(), 1); assert.equal(navigations, 1);
  results.push('안내창 닫기는 알림·작업을 유지');
  await page.locator('#releaseUpdate').click();
  await page.getByRole('dialog', { name: '새 버전 안내' }).getByRole('button', { name: '파일로 저장', exact: true }).click();
  await page.waitForFunction(() => window.__releaseSaveCalls.length === 1);
  assert.equal(await page.evaluate(() => window.__releaseSaveCalls[0].active), true);
  assert.equal(await page.evaluate(() => window.__releaseBook === tabula.wb() && tabula.wb().getValue(0, 0, 0) === 879), true);
  assert.equal(navigations, 1); results.push('명시적 저장만 파일 선택기 호출·취소해도 새로고침 없음');
  await page.clock.fastForward(300000); await page.evaluate(() => window.dispatchEvent(new Event('focus'))); assert.equal(requests.length, 4);
  for (const request of requests) {
    assert.equal(request.method, 'GET'); assert.equal(request.body, null); assert.equal(request.headers.cookie, undefined); assert.equal(request.headers.referer, undefined);
    assert.equal(new URL(request.url).origin, new URL(url).origin);
  }
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
  results.push('동일 출처 GET만·문서/쿠키/referrer 전송 없음·오류 없음');
  console.log(JSON.stringify({ mode: source ? '소스 + 합성 배포 표식' : '실제 해시 번들', cases: results.length, results, manifestRequests: requests.length, errors, writes }, null, 2));
} catch (error) {
  console.error(JSON.stringify({ source, currentAsset, mode, requests, errors, writes }));
  throw error;
} finally { await context.close(); await browser.close(); }
