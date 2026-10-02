// 합성 도형의 텍스트 축소/회전/여백을 실제 DOM 크기로 확인한다.
// WIXEL_URL은 소스 또는 배포 번들, WIXEL_SHAPE_TEXT_SCREENSHOT은 선택 PNG 경로.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
const page = await context.newPage(), errors = [], writes = [], results = [];
page.on('pageerror', error => errors.push(error.message));
await context.route('**/*', route => {
  const request = route.request(), target = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
  return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
});
await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
try {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.wb());
  for (const rotation of [0, 90, 270]) for (const rich of [false, true]) for (const nowrap of [false, true]) {
    await page.evaluate(({ rotation, rich, nowrap }) => {
      const t = window.tabula, w = t.wb(), text = ('매출 보고서와 긴 텍스트 서식 확인 2026\n').repeat(8);
      const shape = { id: 'text-fit', kind: 'rect', x: 100, y: 80, w: 320, h: 190, fill: '#ffffff', stroke: '#4472c4',
        text, size: 26, pad: [13, 19, 23, 29], valign: 'middle', textFit: 'shrink', textRot: rotation, nowrap };
      if (rich) shape.paras = [{ runs: [{ t: text.slice(0, 100), sz: 36, b: true, color: '#123456' }, { t: text.slice(100), sz: 20, i: true }] }];
      w.restore({ sheets: [{ name: '합성 텍스트', cells: { '0,0': { raw: '보존' } }, shapes: [shape] }] });
      t.gv().layout(); t.gv().renderAll();
    }, { rotation, rich, nowrap });
    for (const zoom of [50, 100, 150]) {
      const result = await page.evaluate(zoom => {
        const view = window.tabula.gv(); view.setZoom(zoom); view.layout(); view.renderAll();
        const box = document.querySelector('.obj[data-id="text-fit"] .sh-text'), content = box.firstElementChild, css = getComputedStyle(box);
        const range = document.createRange(); range.selectNodeContents(content);
        const rect = element => { const r = element.getBoundingClientRect(); return { x: r.x, y: r.y, right: r.right, bottom: r.bottom }; };
        const scale = Number(content.style.zoom || 1);
        return { scale, available: [box.clientWidth - parseFloat(css.paddingLeft) - parseFloat(css.paddingRight), box.clientHeight - parseFloat(css.paddingTop) - parseFloat(css.paddingBottom)],
          scroll: [content.scrollWidth * scale, content.scrollHeight * scale], box: rect(box), actual: rect(range) };
      }, zoom);
      const measured = { rotation, rich, nowrap, zoom, ...result }, detail = JSON.stringify(measured);
      assert.ok(result.scale > 0 && result.scale < 1, '긴 텍스트를 실제 축소: ' + detail);
      assert.ok(result.scroll[0] <= result.available[0] + 1 && result.scroll[1] <= result.available[1] + 1, '내부 여백을 제외한 공간에 맞춤: ' + detail);
      assert.ok(result.actual.x >= result.box.x - 1 && result.actual.y >= result.box.y - 1 && result.actual.right <= result.box.right + 1 && result.actual.bottom <= result.box.bottom + 1, '회전 후 실제 글자 경계에 잘림 없음: ' + detail);
      results.push(measured);
    }
  }
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb();
    w.restore({ sheets: [{ name: '패널 확인', cells: { '0,0': { raw: '보존' } }, shapes: [{ id: 'text-fit', kind: 'rect', x: 100, y: 80, w: 320, h: 190, fill: '#ffffff', stroke: '#4472c4', text: '텍스트 '.repeat(60), size: 32, pad: [13, 19, 23, 29] }] }] });
    t.gv().setZoom(100); t.gv().layout(); t.gv().renderAll();
  });
  await page.locator('.obj[data-id="text-fit"]').dblclick({ position: { x: 30, y: 30 } });
  const pane = page.locator('.shape-format-pane');
  await pane.getByRole('tab', { name: '텍스트 옵션', exact: true }).click();
  await pane.getByRole('tab', { name: '텍스트 상자', exact: true }).click();
  await pane.getByLabel('텍스트 자동 맞춤', { exact: true }).selectOption('shrink');
  for (const rotation of [90, 270]) {
    await pane.getByLabel('텍스트 회전', { exact: true }).selectOption(String(rotation));
    assert.deepEqual(await page.evaluate(() => { const shape = window.tabula.wb().sheets[0].shapes[0]; return [shape.textFit, shape.textRot]; }), ['shrink', rotation]);
  }
  if (process.env.WIXEL_SHAPE_TEXT_SCREENSHOT) await page.screenshot({ path: process.env.WIXEL_SHAPE_TEXT_SCREENSHOT });
  assert.equal(await page.evaluate(() => window.tabula.wb().getRaw(0, 0, 0)), '보존');
  assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기');
  console.log(JSON.stringify({ ok: true, geometryCases: results.length, ui: true, pageErrors: errors, blockedWrites: writes }));
} finally { await context.close(); await browser.close(); }
