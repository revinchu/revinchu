// 실제 앱의 찾기/바꾸기 접근키 회귀. 새 컨텍스트의 합성 문서만 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
async function test(name, check) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  try {
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await p.evaluate(() => {
      const t = window.tabula, w = t.wb();
      w.restore({ sheets: [{ name: '찾기 합성 문서', cells: {
        '0,0': { raw: 'Alpha' }, '1,0': { raw: 'alpha' }, '2,0': { raw: 'ALPHA' }, '3,0': { raw: 'Beta' }, '0,1': { raw: '=SUM(2,3)' },
      } }] }); w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
    });
    await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+h');
    const d = p.getByRole('dialog', { name: '찾기 및 바꾸기', exact: true }); await d.waitFor();
    await check(p, d); assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
const values = p => p.evaluate(() => {
  const w = window.tabula.wb(); return { values: [0, 1, 2, 3].map(r => w.getRaw(0, r, 0)), formula: w.getRaw(0, 0, 1), undo: w.undoStack.length };
});
const alt = (p, key) => p.keyboard.press('Alt+' + key);
const focusIs = (p, label) => p.evaluate(label => document.activeElement?.getAttribute('aria-label') === label, label);
const fill = async (p, d, text, replacement = '') => {
  await alt(p, 'n'); assert.ok(await focusIs(p, '찾을 내용')); await d.getByRole('textbox', { name: '찾을 내용', exact: true }).fill(text);
  await alt(p, 'e'); assert.ok(await focusIs(p, '바꿀 내용')); await d.getByRole('textbox', { name: '바꿀 내용', exact: true }).fill(replacement);
};
try {
  await test('Ctrl+H 초기 포커스·Alt+N/E 입력란 전환·Alt+D 닫기·문서 무변경', async (p, d) => {
    assert.ok(await focusIs(p, '찾을 내용'), 'Ctrl+H 직후 찾을 내용에 포커스'); const before = await values(p);
    await fill(p, d, 'alpha', '교체 예정');
    await alt(p, 'n'); assert.ok(await focusIs(p, '찾을 내용')); await alt(p, 'd'); await d.waitFor({ state: 'hidden' });
    assert.deepEqual(await values(p), before); assert.equal(await p.locator('#cellEditor').inputValue(), '');
  });
  await test('Alt+A 모두 바꾸기는 세 셀을 한 번에 변경하고 Ctrl+Z 한 번으로 복원', async (p, d) => {
    const before = await values(p); await fill(p, d, 'alpha', '교체'); await alt(p, 'a');
    assert.deepEqual((await values(p)).values, ['교체', '교체', '교체', 'Beta']);
    assert.equal((await values(p)).formula, '=SUM(2,3)'); assert.equal((await values(p)).undo, before.undo + 1);
    assert.match(await d.locator('.find-status').innerText(), /3개/);
    await alt(p, 'd'); await d.waitFor({ state: 'hidden' }); await p.keyboard.press('Control+z');
    assert.deepEqual(await values(p), before);
  });
  await test('Alt+T 옵션·Alt+C 대소문자 구분은 실제 검색·교체 범위에 반영', async (p, d) => {
    await fill(p, d, 'alpha', '소문자만'); await alt(p, 't');
    const check = d.getByRole('checkbox', { name: '대/소문자 구분', exact: true }); await check.waitFor({ state: 'visible' });
    assert.equal(await check.isChecked(), false); await alt(p, 'c'); assert.equal(await check.isChecked(), true);
    await alt(p, 'a'); assert.deepEqual((await values(p)).values, ['Alpha', '소문자만', 'ALPHA', 'Beta']);
    await alt(p, 't'); await check.waitFor({ state: 'hidden' }); await alt(p, 't'); await check.waitFor({ state: 'visible' });
    assert.equal(await check.isChecked(), true); await alt(p, 'c'); assert.equal(await check.isChecked(), false);
  });
  await test('Alt+F 다음 찾기·Alt+I 모두 찾기는 실제 위치와 세 결과를 표시', async (p, d) => {
    const before = await values(p); await fill(p, d, 'alpha'); await alt(p, 'f');
    assert.deepEqual(await p.evaluate(() => window.tabula.active), { r: 1, c: 0 });
    await alt(p, 'i'); assert.equal(await d.locator('.find-results tbody tr').count(), 3);
    assert.deepEqual(await d.locator('.find-results tbody tr td:nth-child(2)').allTextContents(), ['$A$1', '$A$2', '$A$3']);
    assert.match(await d.locator('.find-status').innerText(), /3개 셀/); assert.deepEqual(await values(p), before);
  });
  await test('Alt+M 중첩 서식 창에서는 Alt+A가 부모 모두 바꾸기를 실행하지 않음', async (p, d) => {
    const before = await values(p); await fill(p, d, 'alpha', '실행되면 안 됨'); await alt(p, 't'); await alt(p, 'm');
    const nested = p.getByRole('dialog', { name: '찾을 서식', exact: true }); await nested.waitFor();
    assert.ok(await nested.evaluate(e => e.contains(document.activeElement)));
    await alt(p, 'a'); assert.deepEqual(await values(p), before);
    if (await nested.count()) { await p.keyboard.press('Escape'); await nested.waitFor({ state: 'hidden' }); }
    assert.ok(await d.evaluate(e => e.contains(document.activeElement)), '중첩 창 종료 뒤 부모 포커스 복원');
    await alt(p, 'm'); await nested.waitFor(); await alt(p, 'r'); await nested.waitFor({ state: 'hidden' });
    assert.ok(await d.evaluate(e => e.contains(document.activeElement)), '서식 지우기 뒤 부모 포커스 복원');
    assert.deepEqual(await values(p), before);
    await alt(p, 'd'); await d.waitFor({ state: 'hidden' }); assert.deepEqual(await values(p), before);
  });
  await test('Alt 배지는 명시 접근키와 연결되고 창 안에서 리본 키팁을 띄우지 않음', async (p, d) => {
    await p.keyboard.press('Alt');
    for (const key of ['n', 'e', 'a', 'd', 't', 'f', 'i']) {
      assert.equal(await d.locator(`[data-access-key="${key}"]`).getAttribute('aria-keyshortcuts'), 'Alt+' + key.toUpperCase());
      assert.ok(await p.locator(`.access-key-badge[data-key="${key}"]`).count(), key + ' 배지');
    }
    assert.equal(await p.evaluate(() => document.body.classList.contains('keytips')), false);
    await p.keyboard.press('Escape'); assert.equal(await p.locator('.access-key-layer').count(), 0); await d.waitFor();
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok) }, null, 2));
if (results.some(x => !x.ok)) process.exitCode = 1;
