// 합성 차트 삽입 갤러리: 소스/배포 번들 공용. 외부/API/쓰기 차단, 사용자 문서 미사용.
// 확대 모의는 CSS viewport/DPR 조합이며 실제 OS 브라우저 확대나 모바일 기기 검사가 아니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_CHART_GALLERY_OUT || 'D:/Codex/Temp/wixel-chart-gallery/source';
const filter = process.env.WIXEL_CHART_GALLERY_FILTER || '';
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = [], measurements = [], pageErrors = [], remoteWrites = [], blockedRequests = [], assets = new Set();
let checks = 0;
const eq = (a, b, m) => { checks++; assert.deepEqual(a, b, m); };
const ok = (a, m) => { checks++; assert.ok(a, m); };
const run = (p, command) => p.evaluate(command => window.tabula.run(command), command);
const dlg = p => p.getByRole('dialog', { name: '차트 삽입', exact: true });
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const model = p => p.evaluate(() => { const w = window.tabula.wb(), s = w.serialize().sheets[0]; return { cells: s.cells, charts: s.charts || [], undo: w.undoStack.length }; });
async function fixture(p, kind = 'long') {
  await p.evaluate(kind => {
    const t = window.tabula, w = t.wb(), cells = {};
    const rows = kind === 'empty' ? [] : kind === 'text' ? [['분류', '내용'], ['문자 항목 하나', '값 없음'], ['문자 항목 둘', '숫자가 아닌 문자열']] :
      [['광고 캠페인', '서울특별시 광고 전환 매출 긴 계열 이름', '부산광역시 광고 전환 매출 긴 계열 이름'], ['아주 긴 한글 범주 이름 캠페인 첫 번째', '120', '200'], ['아주 긴 한글 범주 이름 캠페인 두 번째', '180', '280'], ['아주 긴 한글 범주 이름 캠페인 세 번째', '240', '160']];
    rows.forEach((row, r) => row.forEach((raw, c) => { cells[r + ',' + c] = { raw }; }));
    w.restore({ sheets: [{ name: '갤러리 합성', cells }] });
    t.gv().setZoom(100); t.gv().layout(); t.gv().renderAll();
    t.selectRange({ r1: 0, c1: 0, r2: kind === 'empty' ? 2 : rows.length - 1, c2: kind === 'text' ? 1 : 2 });
    w.undoStack = []; w.redoStack = [];
  }, kind);
  await raf(p);
}
async function open(p, all = false) {
  await run(p, all ? 'insertChartCatalog' : 'insertChartAll');
  const d = dlg(p); await d.waitFor(); await raf(p); return d;
}
async function close(p) { await p.keyboard.press('Escape'); await dlg(p).waitFor({ state: 'detached' }); }
async function usable(p, node, label) {
  await node.scrollIntoViewIfNeeded();
  const result = await node.evaluate(e => { const r = e.getBoundingClientRect(), hit = document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2); return { x: r.x, y: r.y, w: r.width, h: r.height, hit: hit === e || e.contains(hit) }; });
  const v = p.viewportSize();
  ok(result.x >= -1 && result.y >= -1 && result.x + result.w <= v.width + 1 && result.y + result.h <= v.height + 1 && result.hit, label + ' 화면 접근 ' + JSON.stringify(result));
}
async function audit(p, d, name, { numeric = true } = {}) {
  await raf(p);
  await usable(p, d.locator('.dialog-head button'), name + ' 닫기');
  for (const b of await d.locator('.dialog-foot button').all()) await usable(p, b, name + ' 하단');
  const info = await d.evaluate(e => {
    const rect = n => { const r = n.getBoundingClientRect(); return { x: r.x, y: r.y, w: r.width, h: r.height, right: r.right, bottom: r.bottom }; };
    const visible = n => n.getClientRects().length && getComputedStyle(n).visibility !== 'hidden';
    const within = (a, b, tol = 1) => a.x >= b.x - tol && a.y >= b.y - tol && a.right <= b.right + tol && a.bottom <= b.bottom + tol;
    const body = e.querySelector('.dialog-body');
    const buttons = [...e.querySelectorAll('.cr-thumb,.cg-sub')].filter(visible).map(b => {
      const box = rect(b), caption = b.querySelector('.cr-thumb-name,.cr-name,.cg-sub-name,[data-access-caption-host]');
      const texts = [], walker = document.createTreeWalker(caption || b, NodeFilter.SHOW_TEXT);
      let n; while ((n = walker.nextNode())) if (n.textContent.trim() && !n.parentElement.closest('svg,.access-key-hint')) { const range = document.createRange(); range.selectNodeContents(n); for (const r of range.getClientRects()) texts.push({ x: r.x, y: r.y, w: r.width, h: r.height, right: r.right, bottom: r.bottom }); }
      const hints = [...b.querySelectorAll('.access-key-hint')].filter(visible).map(h => ({ ...rect(h), inSvg: !!h.closest('svg'), caption: h.dataset.accessCaption }));
      const overlap = hints.some(h => texts.some(t => Math.min(h.right, t.right) - Math.max(h.x, t.x) > .75 && Math.min(h.bottom, t.bottom) - Math.max(h.y, t.y) > .75));
      return { title: b.title, aria: b.getAttribute('aria-label'), box, caption: caption?.textContent || '', texts, hints, overlap, textOutside: texts.some(r => !within(r, box)), hintOutside: hints.some(r => r.h < 5 || !within(r, box)), svgHint: hints.some(h => h.inSvg), marks: [...b.querySelectorAll('svg [data-s]')].filter(n => { const r = rect(n); return r.w > 2 && r.h > 2; }).length };
    });
    const previews = [...e.querySelectorAll('.cg-prev svg,.cr-prev svg')].filter(visible).map(svg => {
      const box = rect(svg), parent = rect(svg.parentElement), matrix = svg.getScreenCTM();
      const texts = [...svg.querySelectorAll('text')].filter(visible).filter(n => n.textContent.trim()).map(n => ({ text: n.textContent, box: rect(n), size: parseFloat(getComputedStyle(n).fontSize) * Math.hypot(matrix.a, matrix.b) }));
      return { box, parent, source: svg.parentElement.dataset.previewSource, outside: !within(box, parent, 2), textOutside: texts.filter(t => !within(t.box, box, 2)), tinyText: texts.filter(t => t.size < 7.5), texts: texts.length, invalid: /NaN|Infinity|undefined|�/.test(svg.outerHTML) };
    });
    return { bounds: rect(e), body: { scroll: body.scrollWidth, client: body.clientWidth }, buttons, previews };
  });
  measurements.push({ name, viewport: p.viewportSize(), ...info });
  const v = p.viewportSize();
  ok(info.bounds.x >= -1 && info.bounds.y >= -1 && info.bounds.right <= v.width + 1 && info.bounds.bottom <= v.height + 1, name + ' 창 크기');
  ok(info.body.scroll <= info.body.client + 2, name + ' 본문 가로 넘침');
  ok(info.buttons.length > 0, name + ' 갤러리 항목');
  for (const b of info.buttons) {
    eq(b.aria, b.title, name + ' 정확한 유형 이름');
    ok(b.caption.trim().length > 0, name + ' 유형 HTML 캡션 ' + b.title);
    ok(!b.textOutside && !b.hintOutside && !b.svgHint && !b.overlap, name + ' 캡션·접근키 배치 ' + JSON.stringify(b));
    if (numeric) ok(b.marks > 0, name + ' 실제 데이터 도형 ' + b.title);
  }
  ok(info.previews.length > 0, name + ' 큰 미리 보기');
  for (const preview of info.previews) eq(preview.source, numeric ? 'selection' : 'sample', name + ' 선택 데이터와 예시 구분');
  for (const s of info.previews) {
    ok(!s.outside && !s.invalid, name + ' 미리 보기 SVG 범위/유효값');
    eq(s.textOutside.map(t => t.text), [], name + ' SVG 글자 잘림');
    eq(s.tinyText.map(t => ({ text: t.text, size: t.size })), [], name + ' 미리 보기 글자 7.5CSSpx 미만');
  }
  return info;
}
async function shot(p, name) { await p.screenshot({ path: out + '/' + name + '.png' }); }
async function test(name, viewport, action, scale = 1) {
  if (filter && !name.includes(filter)) return;
  const context = await browser.newContext({ viewport, deviceScaleFactor: scale }), p = await context.newPage();
  const localErrors = [], localWrites = [];
  p.setDefaultTimeout(10000);
  p.on('pageerror', e => { localErrors.push(e.message); pageErrors.push({ name, message: e.message }); });
  p.on('dialog', d => d.type() === 'beforeunload' ? d.accept() : d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', read: async () => [], writeText: async () => {}, write: async () => {} } }); });
  await context.route('**/*', route => {
    const q = route.request(), target = new URL(q.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { const item = { name, method: q.method(), path: target.pathname }; localWrites.push(item); remoteWrites.push(item); return route.abort(); }
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) { blockedRequests.push(target.origin + target.pathname); return route.abort(); }
    return route.continue();
  });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await p.waitForFunction(() => window.tabula?.wb());
    for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await p.evaluate(width => { if ((width <= 600) !== !!window.tabula.mobile().active) window.tabula.run('mobileWorkMode'); }, viewport.width);
    await fixture(p); await action(p); eq(localErrors, [], '페이지 오류'); eq(localWrites, [], '원격 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message }); console.error('NG ' + name + ': ' + error.stack); await shot(p, 'failure-' + results.length).catch(() => {}); }
  finally { await context.close(); }
}
try {
  for (const [width, height] of [[1440, 900], [1024, 700], [390, 844], [320, 640]]) {
    const viewport = { width, height };
    await test(width + 'px 추천 차트: 긴 한글·캡션·키보드 선택·취소', viewport, async p => {
      const before = await model(p), d = await open(p);
      await audit(p, d, width + '-recommended');
      const first = d.locator('.cr-thumb').first(); await first.focus(); await p.keyboard.press('ArrowDown');
      eq(await d.locator('.cr-thumb.on').getAttribute('title'), await d.locator('.cr-thumb').nth(1).getAttribute('title'), '추천 목록 방향키 선택');
      await first.scrollIntoViewIfNeeded(); const key = await first.getAttribute('data-resolved-access-key'); ok(key, '추천 항목 접근키'); await p.keyboard.press('Alt'); await p.keyboard.press(key); eq(await first.getAttribute('aria-pressed'), 'true', '추천 항목 Alt 접근키 선택'); await shot(p, width + '-recommended');
      await close(p); eq(await model(p), before, '추천 선택 취소는 원본·Undo 불변');
    });
    await test(width + 'px 모든 차트: 100% 누적·검색·클릭·삽입·Undo', viewport, async p => {
      const before = await model(p), d = await open(p, true), search = d.getByRole('searchbox', { name: '차트 종류 검색' });
      await search.fill('100%'); await d.locator('.cg-sub:visible').first().click();
      await audit(p, d, width + '-all-percent'); await shot(p, width + '-all-percent');
      await usable(p, d.getByRole('button', { name: '확인', exact: true }), '삽입 확인');
      await d.getByRole('button', { name: '확인', exact: true }).click(); await d.waitFor({ state: 'detached' });
      const after = await model(p); eq(after.cells, before.cells, '원본 데이터 불변'); eq(after.charts.length, 1); eq(after.charts[0].grouping, 'percentStacked'); eq(after.undo, before.undo + 1, '삽입 Undo 하나');
      await run(p, 'undo'); eq((await model(p)).cells, before.cells); eq((await model(p)).charts, before.charts);
    });
    await test(width + 'px 빈 셀·문자만 선택: 예시 안내·실제 삽입 차단·취소', viewport, async p => {
      for (const kind of ['empty', 'text']) {
        await fixture(p, kind); const before = await model(p), d = await open(p);
        await audit(p, d, width + '-' + kind, { numeric: false });
        ok(/데이터|수치|숫자|선택|미리/.test(await d.innerText()), '숫자 없는 선택 안내');
        await shot(p, width + '-' + kind); await close(p); eq(await model(p), before);
        const reopened = await open(p); await reopened.getByRole('button', { name: '확인', exact: true }).click();
        await p.waitForFunction(() => /숫자가 들어 있는 데이터 범위/.test(document.querySelector('#dialogLayer')?.textContent || ''));
        eq(await model(p), before, '실제 빈 차트·예시 숫자를 생성하지 않음');
        for (let i = 0; i < 3; i++) await p.keyboard.press('Escape');
      }
    });
    await test(width + 'px 시트 400% 확대: 갤러리 자체는 확대되지 않고 조작 가능', viewport, async p => {
      await p.locator('#zoomSlider').evaluate(slider => { slider.value = '400'; slider.dispatchEvent(new Event('input', { bubbles: true })); slider.dispatchEvent(new Event('change', { bubbles: true })); });
      const before = await model(p), d = await open(p, true); await d.getByRole('searchbox', { name: '차트 종류 검색' }).fill('누적');
      await audit(p, d, width + '-sheet400'); await shot(p, width + '-sheet400');
      eq(await p.evaluate(() => Math.round(window.tabula.gv().z * 100)), 400); await close(p); eq(await model(p), before);
    });
  }
  await test('검색 결과 없음·검색 지우기·하위 유형 방향키·Enter 확정', { width: 1440, height: 900 }, async p => {
    const before = await model(p), d = await open(p, true), search = d.getByRole('searchbox', { name: '차트 종류 검색' });
    await search.fill('없는차트종류999'); eq(await d.locator('.cg-sub:visible').count(), 0); ok(/일치하는 차트가 없습니다/.test(await d.innerText()));
    const confirm = d.getByRole('button', { name: '확인', exact: true }); if (await confirm.isEnabled()) await confirm.click(); eq(await d.count(), 1); eq(await model(p), before);
    await search.fill(''); const cat = d.locator('.cg-cat').filter({ hasText: /^세로 막대형$/ }); await cat.focus(); await p.keyboard.press('ArrowDown'); eq(await d.locator('.cg-cat.on').innerText(), '꺾은선형'); await p.keyboard.press('Home');
    await d.locator('.cg-sub').first().focus(); await p.keyboard.press('ArrowRight'); eq(await d.locator('.cg-sub[aria-pressed=true]').getAttribute('title'), '누적 세로 막대형'); await p.keyboard.press('Enter');
    // 갤러리 Enter는 해당 항목 선택, 확인 단추 Enter는 확정이다.
    if (await d.count()) { await confirm.focus(); await p.keyboard.press('Enter'); }
    await d.waitFor({ state: 'detached' }); eq((await model(p)).charts[0].grouping, 'stacked'); eq((await model(p)).cells, before.cells);
  });
  for (const scale of [1.25, 2]) await test('브라우저 확대 ' + scale * 100 + '% 모의: CSS viewport·DPR 조합', { width: Math.round(1440 / scale), height: Math.round(900 / scale) }, async p => {
    const before = await model(p), d = await open(p); await audit(p, d, 'browser-scale-' + scale); await shot(p, 'browser-scale-' + scale); await close(p); eq(await model(p), before);
  }, scale);
} finally {
  await browser.close(); const summary = { url, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, pageErrors, remoteWrites, blockedRequests, results };
  await writeFile(out + '/result.json', JSON.stringify(summary, null, 2)); await writeFile(out + '/measurements.json', JSON.stringify(measurements, null, 2));
  console.log(JSON.stringify(summary)); if (results.some(r => !r.ok) || pageErrors.length || remoteWrites.length) process.exitCode = 1;
}
