// 합성 검색 결과만 사용하는 사진 미리보기 크기/동작 회귀.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 서버에서만 실행하세요.');
const out = process.env.WIXEL_GALLERY_OUT || 'D:/Codex/Temp/wixel-online-gallery/' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [];
const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="600" height="400"><rect width="600" height="400" fill="#8ad4e8"/><circle cx="450" cy="95" r="50" fill="#ffc94d"/><path d="M0 400V300L170 100 380 300 480 220 600 340V400Z" fill="#268463"/></svg>';
try {
  for (const width of [320, 390, 768, 1024, 1366]) for (const mode of ['off', 'on']) {
    const height = ({ 320: 640, 390: 844, 768: 1024, 1024: 768, 1366: 900 })[width];
    const context = await browser.newContext({ viewport: { width, height }, serviceWorkers: 'block' });
    const p = await context.newPage(), errors = [], writes = [];
    p.on('pageerror', e => errors.push(e.message)); p.setDefaultTimeout(10000);
    await context.addInitScript(mode => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel.mobile-work.v1', mode); }, mode);
    await context.route('**/*', r => {
      const q = r.request(), u = new URL(q.url());
      if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push(q.method() + ' ' + u.pathname); return r.abort(); }
      if (u.hostname === 'api.openverse.org') return r.fulfill({ json: { page_count: 1, results: Array.from({ length: 12 }, (_, i) => ({ id: i, title: '사진 미리보기 ' + i, url: 'https://gallery.invalid/' + i + '.svg', thumbnail: 'https://gallery.invalid/' + i + '.svg', creator: '합성 작가', license: 'by', license_url: 'https://creativecommons.org/licenses/by/4.0/', foreign_landing_url: 'https://gallery.invalid/source/' + i })) }, headers: { 'access-control-allow-origin': '*' } });
      if (u.hostname === 'gallery.invalid') return r.fulfill({ contentType: 'image/svg+xml', body: svg, headers: { 'access-control-allow-origin': '*' } });
      return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? r.continue() : r.abort();
    });
    try {
      await p.goto(url, { waitUntil: 'domcontentloaded' }); await p.waitForFunction(() => window.tabula?.gv());
      await p.evaluate(() => tabula.openNamedMenu('picture', { x: 20, y: 180 }));
      await p.getByRole('menuitem', { name: '온라인 그림...', exact: true }).click();
      const dlg = p.getByRole('dialog', { name: '온라인 그림', exact: true });
      await dlg.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('openverse');
      await dlg.getByRole('searchbox', { name: '그림 검색어' }).fill('사진');
      await dlg.getByRole('button', { name: '검색', exact: true }).click();
      await p.waitForFunction(() => document.querySelectorAll('.online-card').length === 12 && !document.querySelector('.online-grid').hasAttribute('aria-busy'));
      await dlg.locator('.online-card').first().scrollIntoViewIfNeeded();
      const dimensions = await dlg.evaluate(d => {
        const g = d.querySelector('.online-grid'), card = g.querySelector('.online-card'), img = card.querySelector('img'), r = d.getBoundingClientRect(), cr = card.getBoundingClientRect(), ir = img.getBoundingClientRect();
        return { dialog: { left: r.left, right: r.right, top: r.top, bottom: r.bottom }, viewport: innerWidth, overflow: d.scrollWidth - d.clientWidth, grid: g.clientWidth, card: card.clientWidth, cardHeight: cr.height, imageVisibleHeight: Math.max(0, Math.min(ir.bottom, cr.bottom) - Math.max(ir.top, cr.top)), image: { width: img.clientWidth, height: img.clientHeight, fit: getComputedStyle(img).objectFit }, columns: getComputedStyle(g).gridTemplateColumns };
      });
      assert.ok(dimensions.card >= Math.min(178, dimensions.grid - 2), JSON.stringify(dimensions));
      assert.ok(dimensions.image.height >= 120, JSON.stringify(dimensions));
      assert.ok(dimensions.imageVisibleHeight >= dimensions.image.height - 1 && dimensions.cardHeight > dimensions.image.height + 50, JSON.stringify(dimensions));
      assert.equal(dimensions.image.fit, 'contain'); assert.ok(dimensions.overflow <= 1, JSON.stringify(dimensions));
      assert.ok(dimensions.dialog.left >= -1 && dimensions.dialog.right <= width + 1 && dimensions.dialog.top >= -1 && dimensions.dialog.bottom <= height + 1, JSON.stringify(dimensions));
      await dlg.locator('.online-item').nth(0).click(); await dlg.locator('.online-item').nth(1).click();
      assert.equal(await dlg.locator('.online-item[aria-pressed="true"]').count(), 2);
      await dlg.locator('.online-card-actions button').first().click();
      const preview = p.getByRole('dialog', { name: '그림 미리보기', exact: true });
      await preview.locator('img').waitFor(); const box = await preview.locator('img').boundingBox();
      assert.ok(box.width >= Math.min(260, width - 60), JSON.stringify(box));
      await preview.getByRole('button', { name: '닫기', exact: true }).last().click();
      assert.equal(await dlg.locator('.online-item[aria-pressed="true"]').count(), 2);
      await dlg.locator('.online-card').first().scrollIntoViewIfNeeded();
      await p.screenshot({ path: `${out}/${width}-${mode}.png` });
      assert.deepEqual(errors, []); assert.deepEqual(writes, []);
      results.push({ width, mode, ok: true, dimensions, assets: await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src'))) });
      console.log('OK', width, mode, dimensions.card, dimensions.image.height);
    } catch (e) { results.push({ width, mode, ok: false, error: e.stack, errors, writes }); console.error('NG', width, mode, e.message); }
    finally { await context.close(); }
  }
} finally {
  await browser.close(); await writeFile(out + '/result.json', JSON.stringify({ engine, results }, null, 2));
  console.log(JSON.stringify({ engine, cases: results.length, passed: results.filter(r => r.ok).length }));
  if (results.some(r => !r.ok)) process.exitCode = 1;
}
