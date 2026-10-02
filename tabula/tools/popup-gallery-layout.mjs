// 합성 문서 전용: 비차트 갤러리의 접근키/미리보기/이름/화면 경계. 실제 모바일 OS 검사는 아님.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_POPUP_OUT || 'D:/Codex/Temp/wixel-popup-gallery';
const origin = new URL(url).origin;
const galleries = [
  { name: 'table-styles', command: 'tableStyles', selector: '.tstyle', min: 60 },
  { name: 'cell-styles', command: 'cellStyles', selector: '.style-chip[data-cell-style]', min: 50 },
  { name: 'smartart', command: 'insertSmartArt', selector: '.sa-layout', min: 10 },
  { name: 'icons', command: 'insertIcons', selector: '.ic-cell', min: 100 },
  { name: 'picture-styles', command: 'pictureStyles', selector: '.picture-style-swatch', min: 10 },
];
const results = [], errors = [], writes = [], blockedReads = [];
let checks = 0;
const ok = (value, message) => { checks++; assert.ok(value, message); };
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const browser = await chromium.launch();
await mkdir(out, { recursive: true });

async function fixture(page, picture = false) {
  for (let i = 0; i < 3; i++) await page.keyboard.press('Escape');
  await page.evaluate(picture => {
    const t = window.tabula, w = t.wb();
    w.restore({ sheets: [{ name: '갤러리 합성', cells: { '0,0': { raw: '분류' }, '0,1': { raw: '금액' }, '1,0': { raw: 'A' }, '1,1': { raw: '100' } } }] });
    t.switchSheet(0); t.selectRange({ r1: 0, c1: 0, r2: 1, c2: 1 });
    if (picture) {
      const c = document.createElement('canvas'); c.width = 100; c.height = 60;
      const ctx = c.getContext('2d'); ctx.fillStyle = '#123456'; ctx.fillRect(0, 0, 100, 60);
      w.transact(() => w.setSheetProp(0, 'images', [{ id: 'gallery-picture', src: c.toDataURL(), x: 30, y: 30, w: 100, h: 60 }]));
    }
    t.gv().layout(); t.gv().renderAll();
  }, picture);
  if (picture) await page.locator('.obj.pic[data-id="gallery-picture"]').first().click({ position: { x: 10, y: 10 } });
}

async function measure(page, selector) {
  return page.evaluate(selector => {
    const scopes = [...document.querySelectorAll('#dialogLayer .dialog,#menuLayer>.menu')].filter(e => e.getBoundingClientRect().width);
    const scope = scopes.at(-1); if (!scope) throw new Error('갤러리의 창 또는 메뉴가 없습니다.');
    const box = e => { const r = e.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height }; };
    const visible = e => e.getClientRects().length && getComputedStyle(e).visibility !== 'hidden';
    const chips = [...scope.querySelectorAll(selector)], hints = [...scope.querySelectorAll('.access-key-hint')];
    return {
      scope: box(scope), scopeScroll: { client: scope.clientWidth, scroll: scope.scrollWidth }, chipCount: chips.length,
      previewHints: hints.filter(e => {
        const p = e.parentElement, clone = p.cloneNode(true);
        clone.querySelectorAll('.access-key-hint,svg,img,canvas').forEach(x => x.remove());
        return p.querySelector('svg,img,canvas') && !clone.textContent.trim();
      }).map(e => ({ parent: e.parentElement.className, caption: e.dataset.accessCaption })),
      chips: chips.map(e => ({ title: e.title, aria: e.getAttribute('aria-label'), text: e.textContent.slice(0, 80), client: e.clientWidth, scroll: e.scrollWidth, rect: box(e) })),
      offControls: [...scope.querySelectorAll('input,select,button')].filter(visible).filter(e => { const r = e.getBoundingClientRect(); return r.x < -1 || r.right > innerWidth + 1; }).map(e => ({ class: e.className, title: e.title, text: e.textContent.slice(0, 35), rect: box(e) })),
    };
  }, selector);
}

try {
  for (const [width, height] of [[320, 844], [390, 844], [1024, 768]]) {
    const context = await browser.newContext({ viewport: { width, height } }), page = await context.newPage();
    page.setDefaultTimeout(10000); page.on('pageerror', e => errors.push({ width, message: e.message }));
    await context.addInitScript(() => {
      window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { write: async () => {}, read: async () => [], writeText: async () => {}, readText: async () => '' } });
    });
    await context.route('**/*', route => {
      const request = route.request(), target = new URL(request.url());
      if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push({ method: request.method(), path: target.pathname }); return route.abort(); }
      if (target.origin !== origin || target.pathname.startsWith('/api/')) { blockedReads.push(target.pathname); return route.abort(); }
      return route.continue();
    });
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForFunction(() => window.tabula?.wb()); await page.evaluate(() => document.fonts.ready);
      for (const item of galleries) {
        const errorStart = errors.length;
        try {
          await fixture(page, item.name === 'picture-styles');
          await page.evaluate(({ command, name }) => name.endsWith('-styles') ? window.tabula.openNamedMenu(command, { x: 12, y: 140 }) : window.tabula.run(command), item);
          await page.locator(item.selector).first().waitFor();
          await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
          const m = await measure(page, item.selector), r = m.scope;
          ok(m.chipCount >= item.min, item.name + ' 실제 갤러리 항목');
          ok(r.x >= -1 && r.y >= -1 && r.x + r.w <= width + 1 && r.y + r.h <= height + 1, item.name + ' 창 경계 ' + JSON.stringify(r));
          ok(m.scopeScroll.scroll <= m.scopeScroll.client + 2, item.name + ' 메뉴 가로 넘침');
          eq(m.previewHints, [], item.name + ' SVG/그림 미리보기 안에 접근키를 삽입하지 않음');
          eq(m.offControls, [], item.name + ' 입력·버튼의 가로 화면 경계');
          if (item.name === 'cell-styles') eq(m.chips.filter(c => c.scroll > c.client + 2), [], '셀 스타일의 전체 이름과 상시 접근키가 잘리지 않음');
          await page.screenshot({ path: path.join(out, `${width}-${item.name}.png`) });
          await page.keyboard.down('Alt');
          const alt = await page.locator('.access-key-badge').count(); ok(alt > 0, item.name + ' Alt 접근키 레이어 유지');
          await page.screenshot({ path: path.join(out, `${width}-${item.name}-alt.png`) }); await page.keyboard.up('Alt');
          eq(errors.slice(errorStart), [], item.name + ' 페이지 오류');
          results.push({ width, name: item.name, ok: true, alt, ...m }); console.log(`OK ${width} ${item.name}: ${m.chipCount}개`);
        } catch (error) {
          results.push({ width, name: item.name, ok: false, error: error.message }); console.error(`NG ${width} ${item.name}: ${error.stack}`);
          await page.keyboard.up('Alt').catch(() => {}); await page.screenshot({ path: path.join(out, `${width}-${item.name}-failure.png`) }).catch(() => {});
        }
      }
    } finally { await context.close(); }
  }
  eq(results.length, 15, '3개 화면 너비 × 5개 갤러리'); eq(errors, [], '전체 페이지 오류'); eq(writes, [], '원격 쓰기 요청');
} finally {
  await browser.close();
  const summary = { url, tests: results.length, good: results.filter(r => r.ok).length, bad: results.filter(r => !r.ok).length, checks, pageErrors: errors, remoteWrites: writes, blockedReads, results };
  await writeFile(path.join(out, 'popup-gallery-layout.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ ...summary, results: results.map(({ width, name, ok, error }) => ({ width, name, ok, ...(error ? { error } : {}) })) }, null, 2));
  if (summary.bad || errors.length || writes.length) process.exitCode = 1;
}
