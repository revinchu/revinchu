// 새 브라우저 프로필의 옵션만 사용한다. 사용자 설정·문서·서버 보관함은 건드리지 않는다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const p = await context.newPage(); p.setDefaultTimeout(10000); const errors = [], writes = [];
  p.on('pageerror', (e) => errors.push(e.stack || e.message));
  await context.route('**/*', (route) => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  try {
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/'); await p.waitForFunction(() => window.tabula?.wb());
    await run(p); assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '서버 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
const open = async (p) => { await p.evaluate(() => window.tabula.run('options')); return p.getByRole('dialog', { name: 'WIXEL 옵션', exact: true }); };
const saved = (p) => p.evaluate(() => localStorage.getItem('wixel.options'));
const qat = (p) => p.locator('#quickAccess [data-qat-cmd]').evaluateAll((nodes) => nodes.map((n) => n.dataset.qatCmd));
const listValues = (locator) => locator.locator('option').evaluateAll((options) => options.map((o) => o.value));
try {
  await test('설정 검색: 눈금선 범주·강조·0결과에서 복원·취소 무변경', async (p) => {
    const original = await saved(p), d = await open(p), search = d.getByRole('searchbox', { name: '설정 검색', exact: true });
    await search.fill('눈금선'); assert.deepEqual(await d.getByRole('tab').allTextContents(), ['고급', '빠른 실행 도구 모음']);
    assert.equal(await d.getByRole('checkbox', { name: '눈금선 표시', exact: true }).isVisible(), true);
    assert.ok(await d.locator('.opt-match').count() >= 2);
    await d.getByRole('checkbox', { name: '눈금선 표시', exact: true }).uncheck();
    await search.fill('없는설정999'); assert.equal(await d.getByRole('tab').count(), 0); assert.equal(await d.locator('.opt-box > *').count(), 0); assert.match(await d.locator('.opt-search-status').innerText(), /일치하는 설정이 없습니다/);
    await search.fill(''); assert.equal(await d.getByRole('tab').count(), 9); assert.equal(await d.getByRole('checkbox', { name: '눈금선 표시', exact: true }).isChecked(), false);
    await d.getByRole('button', { name: '취소', exact: true }).click();
    assert.equal(await saved(p), original); assert.equal(await p.evaluate(() => !!window.tabula.wb().sheets[0].noGrid), false);
  });
  await test('옵션 범주 방향키·Home/End와 검색된 범주 사이 이동', async (p) => {
    const d = await open(p); await d.getByRole('tab', { name: '일반', exact: true }).focus(); await p.keyboard.press('ArrowDown');
    assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '수식'); await p.keyboard.press('ArrowUp'); assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '일반');
    await p.keyboard.press('End'); assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '빠른 실행 도구 모음');
    await p.keyboard.press('Home'); assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '일반');
    await d.getByRole('searchbox', { name: '설정 검색', exact: true }).fill('눈금선');
    await d.getByRole('tab', { name: '고급', exact: true }).focus(); await p.keyboard.press('ArrowDown'); assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '빠른 실행 도구 모음');
    await p.keyboard.press('ArrowUp'); assert.equal(await d.getByRole('tab', { selected: true }).innerText(), '고급');
    assert.equal(await d.getByRole('tabpanel', { name: '고급', exact: true }).isVisible(), true);
    await p.screenshot({ path: process.env.WIXEL_SETTINGS_SCREENSHOT || 'D:/Codex/Temp/wixel3-settings-search.png' });
  });
  await test('빠른 실행 명령 검색: 순서·선택 보존·취소·추가/이동 저장·새로고침 복원', async (p) => {
    const original = await qat(p); let d = await open(p); await d.getByRole('tab', { name: '빠른 실행 도구 모음', exact: true }).click();
    let all = d.getByRole('listbox', { name: '사용 가능한 명령', exact: true }), cur = d.getByRole('listbox', { name: '현재 도구 모음 순서', exact: true }), search = d.getByRole('searchbox', { name: '빠른 실행 명령 검색', exact: true });
    await cur.selectOption(original[1]); await search.fill('굵게'); assert.deepEqual(await listValues(all), ['bold', 'slicerBold']);
    assert.deepEqual(await listValues(cur), original); assert.equal(await cur.inputValue(), original[1]);
    await search.fill('없는명령999'); assert.deepEqual(await listValues(all), []); assert.deepEqual(await listValues(cur), original); assert.equal(await cur.inputValue(), original[1]);
    await search.fill('굵게'); await all.selectOption('bold'); await d.getByRole('button', { name: '추가', exact: true }).click();
    assert.deepEqual(await listValues(cur), [...original, 'bold']); await d.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await qat(p), original);
    d = await open(p); await d.getByRole('tab', { name: '빠른 실행 도구 모음', exact: true }).click();
    all = d.getByRole('listbox', { name: '사용 가능한 명령', exact: true }); cur = d.getByRole('listbox', { name: '현재 도구 모음 순서', exact: true }); search = d.getByRole('searchbox', { name: '빠른 실행 명령 검색', exact: true });
    await search.fill('굵게'); await all.selectOption('bold'); await d.getByRole('button', { name: '추가', exact: true }).click();
    await cur.selectOption(original[1]); await d.getByRole('button', { name: '위로', exact: true }).click();
    const expected = [original[1], original[0], ...original.slice(2), 'bold']; assert.deepEqual(await listValues(cur), expected);
    await d.getByRole('button', { name: '확인', exact: true }).click(); assert.deepEqual(await qat(p), expected);
    await p.reload(); await p.waitForFunction(() => window.tabula?.wb()); assert.deepEqual(await qat(p), expected);
  });
  await test('고급 소수 자릿수 0: 저장·재열기·실제 입력·취소·새로고침 보존', async (p) => {
    let d = await open(p); await d.getByRole('tab', { name: '고급', exact: true }).click();
    await d.getByRole('checkbox', { name: '소수점 자동 삽입', exact: true }).check();
    let decimals = d.getByRole('spinbutton', { name: '소수 자릿수', exact: true }); await decimals.fill('0'); await decimals.press('Tab');
    await d.getByRole('button', { name: '확인', exact: true }).click(); assert.equal(JSON.parse(await saved(p)).decimalPlaces, 0);
    d = await open(p); await d.getByRole('tab', { name: '고급', exact: true }).click(); decimals = d.getByRole('spinbutton', { name: '소수 자릿수', exact: true }); assert.equal(await decimals.inputValue(), '0');
    await decimals.fill('-3'); await decimals.press('Tab'); await d.getByRole('button', { name: '취소', exact: true }).click(); assert.equal(JSON.parse(await saved(p)).decimalPlaces, 0);
    await p.evaluate(() => window.tabula.selectCell(0, 0)); await p.keyboard.type('123'); await p.keyboard.press('Enter'); assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 123);
    await p.reload(); await p.waitForFunction(() => window.tabula?.wb()); d = await open(p); await d.getByRole('tab', { name: '고급', exact: true }).click(); assert.equal(await d.getByRole('spinbutton', { name: '소수 자릿수', exact: true }).inputValue(), '0');
  });
  const failed = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ tests: results.length, failed, results }, null, 2)); if (failed) process.exitCode = 1;
} finally { await browser.close(); }
