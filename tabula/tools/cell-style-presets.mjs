// 새 보고서 스타일과 갤러리: 합성 문서·격리 컨텍스트만 사용, 원격 쓰기는 차단.
import assert from 'node:assert/strict';
import { REPORT_CELL_STYLE_SECTIONS } from '../src/cell-style-presets.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 서버에서만 실행하세요.');
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
const page = await context.newPage(), errors = [], writes = [], results = [];
page.setDefaultTimeout(20000); page.on('pageerror', error => errors.push(error.message));
await context.route('**/*', route => { if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); } return route.continue(); });
const menu = async () => { await page.evaluate(() => tabula.openNamedMenu('cellStyles', { x: 150, y: 100 })); };
const search = () => page.getByRole('searchbox', { name: '셀 스타일 검색' });
const cell = () => page.evaluate(() => ({ style: tabula.wb().styleAt(0, 0, 0), value: tabula.wb().getValue(0, 0, 0), names: tabula.wb().cellStyles?.map(s => s.name) ?? [] }));
try {
  await page.addInitScript(() => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => window.tabula?.wb());
  if (await page.locator('#autosaveToggle').getAttribute('aria-checked') === 'true') await page.locator('#autosaveToggle').click();
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => { w.setInput(0, 0, 0, '0.175'); w.setStyle(0, 0, 0, { numFmt: 'custom', code: '0.000%', align: 'right', locked: false }); }); tabula.selectCell(0, 0); });
  await menu(); assert.equal(await page.locator('.style-chip').count(), 71);
  const names = await page.locator('.style-chip').evaluateAll(nodes => nodes.map(n => n.dataset.cellStyle));
  for (const old of ['표준', '좋음', '입력', '제목', '제목 4', '20% - 강조색1', '60% - 강조색6', '강조색6', '통화 [0]']) assert.ok(names.includes(old), old);
  for (const s of REPORT_CELL_STYLE_SECTIONS.flatMap(([, list]) => list)) assert.ok(names.includes(s.name), s.name);
  assert.equal(await search().evaluate(node => node === document.activeElement), true);
  results.push('기존 47개와 보고서 24개·검색 초기 포커스');
  await search().fill('KPI 블루'); await page.keyboard.press('ArrowDown');
  assert.equal(await page.evaluate(() => document.activeElement.dataset.cellStyle), 'KPI 블루');
  await page.keyboard.press('Enter'); let state = await cell();
  assert.equal(state.style.cellStyleName, 'KPI 블루'); assert.equal(state.style.fill, '#eff6ff'); assert.equal(state.style.code, '0.000%');
  assert.equal(state.style.align, 'right'); assert.equal(state.style.locked, false); assert.equal(state.value, .175);
  await page.evaluate(() => tabula.run('undo')); assert.equal((await cell()).style.cellStyleName, undefined);
  await page.evaluate(() => tabula.run('redo')); assert.equal((await cell()).style.cellStyleName, 'KPI 블루');
  results.push('검색·방향키·Enter 적용·숫자 표시/맞춤/잠금 유지·한 번 Undo/Redo');
  await menu(); await search().fill('일치하지 않는 합성 이름'); assert.equal(await page.locator('.style-chip:visible').count(), 0);
  await page.keyboard.press('Enter'); assert.equal((await cell()).style.cellStyleName, 'KPI 블루');
  await search().fill(''); assert.equal(await page.locator('.style-chip:not([hidden])').count(), 71);
  await page.keyboard.press('ArrowDown'); await page.keyboard.press('ArrowRight'); assert.equal(await page.evaluate(() => document.activeElement.dataset.cellStyle), '나쁨');
  await page.keyboard.press('Escape'); results.push('검색 결과 없음은 적용하지 않고 검색 초기화·방향키 유지');
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setCellStyles([{ name: 'KPI 블루', style: { fill: '#ffe4e6', color: '#881337' }, include: { number: false, alignment: false, font: true, border: true, fill: true, protection: false } }, { name: '내 기존 스타일', style: { color: '#123456' } }])); });
  await menu(); assert.equal(await page.locator('[data-cell-style="KPI 블루"]').count(), 1); assert.equal(await page.locator('[data-cell-style="내 기존 스타일"]').count(), 1);
  await search().fill('KPI 블루'); assert.equal(await page.locator('[data-cell-style="KPI 블루"]').evaluate(n => getComputedStyle(n).backgroundColor), 'rgb(255, 228, 230)');
  await page.keyboard.press('Enter'); state = await cell(); assert.equal(state.style.fill, '#ffe4e6'); assert.equal(state.style.code, '0.000%');
  results.push('기존 사용자 정의와 동일 이름의 저장된 프리셋 재정의를 우선');
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setCellStyles([])); });
  for (const viewport of [{ width: 320, height: 480 }, { width: 760, height: 520 }]) {
    await page.setViewportSize(viewport); await menu();
    const geom = await page.locator('.menu').evaluate(node => { const r = node.getBoundingClientRect(); const grid = node.querySelector('.style-grid'); return { x: r.x, right: r.right, top: r.top, bottom: r.bottom, scrollWidth: node.scrollWidth, width: node.clientWidth, columns: getComputedStyle(grid).gridTemplateColumns.split(' ').length }; });
    assert.ok(geom.x >= 0 && geom.right <= viewport.width && geom.top >= 0 && geom.bottom <= viewport.height, JSON.stringify(geom));
    assert.ok(geom.scrollWidth <= geom.width + 1, JSON.stringify(geom));
    await search().fill('검토'); await page.keyboard.press('ArrowDown');
    const before = await page.locator('.style-chip:visible').evaluateAll(nodes => nodes.map(n => n.dataset.cellStyle));
    await page.keyboard.press('ArrowDown'); assert.equal(await page.evaluate(() => document.activeElement.dataset.cellStyle), before[Math.min(geom.columns, before.length - 1)]);
    await page.keyboard.press('Escape');
  }
  results.push('320/760px 화면 안 배치·가로 스크롤 없음·실제 열 수로 위아래 이동');
  await page.setViewportSize({ width: 1440, height: 1000 }); await menu(); await search().fill('보고서');
  await page.screenshot({ path: process.env.WIXEL_STYLE_PRESETS_SCREENSHOT || 'D:/Codex/Temp/wixel-cell-style-presets.png' });
  await page.keyboard.press('Escape');
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
  console.log(JSON.stringify({ cases: results.length, results, pageErrors: errors, blockedWrites: writes }, null, 2));
} catch (error) { console.error(JSON.stringify({ errors, writes, menu: await page.locator('#menuLayer').innerText() })); throw error; }
finally { await context.close(); await browser.close(); }
