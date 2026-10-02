// 텍스트 나누기 단계별 기본 동작·Excel 접근키·변환/실행 취소 회귀.
// 합성 문서와 격리 브라우저만 사용한다. 실제 Windows IME를 구동하는 검사는 아니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5178/';
if (!['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 서버에서만 실행하세요.');
const browser = await chromium.launch(), results = [];
const dlg = p => p.getByRole('dialog', { name: '텍스트 마법사', exact: true });
const stage = p => dlg(p).locator('.ttc-step').textContent();
const focus = p => p.evaluate(() => document.activeElement?.dataset.accessKey ?? document.activeElement?.dataset.ttcFocus);
const cell = (p, r, c) => p.evaluate(([r, c]) => ({ raw: tabula.wb().getRaw(0, r, c), value: tabula.wb().getValue(0, r, c), style: tabula.wb().styleAt(0, r, c) }), [r, c]);
const snap = p => p.evaluate(() => {
 const ordered = v => Array.isArray(v) ? v.map(ordered) : v && typeof v === 'object' ? Object.fromEntries(Object.keys(v).sort().map(k => [k, ordered(v[k])])) : v;
 return JSON.stringify(ordered(tabula.wb().serialize()));
});
const next = p => dlg(p).getByRole('button', { name: '다음(N) >', exact: true }).click();
const finish = p => dlg(p).getByRole('button', { name: '마침(F)', exact: true }).click();
async function setup(p, lines, extra = {}) {
 await p.evaluate(({ lines, extra }) => {
  const cells = { ...extra }; lines.forEach((raw, r) => { cells[`${r},0`] = { raw }; });
  const t = tabula, w = t.wb(); w.restore({ sheets: [{ name: '텍스트 합성', cells }] }); w.undoStack = []; w.redoStack = [];
  t.gv().layout(); t.gv().renderAll(); t.selectRange({ r1: 0, c1: 0, r2: lines.length - 1, c2: 0 });
 }, { lines, extra });
}
async function open(p) { await p.keyboard.press('Alt'); await p.keyboard.press('a'); await p.keyboard.press('e'); await dlg(p).waitFor(); }
async function test(name, fn) {
 const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }), p = await context.newPage(), errors = [], writes = [];
 p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message));
 await context.route('**/*', r => { const q = r.request(); if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push(q.url()); return r.abort(); } if (new URL(q.url()).origin !== new URL(url).origin || new URL(q.url()).pathname.startsWith('/api/')) return r.abort(); return r.continue(); });
 try {
  await p.addInitScript(() => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; });
  await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
  await fn(p); assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기'); results.push({ name, ok: true }); console.log(`OK ${name}`);
 } catch (e) { results.push({ name, ok: false, error: e.message, errors, writes }); console.error(`NG ${name}: ${e.stack}`); }
 finally { await context.close(); }
}
try {
 await test('Alt A E → Enter 3회: 다음/다음/마침·셀 오염 없음·한 번 Undo/Redo', async p => {
  await setup(p, ['12\t34', '56\t78']); const before = await snap(p); await open(p);
  assert.equal(await focus(p), 'n'); assert.match(await stage(p), /^1\/3/);
  await p.keyboard.press('Enter'); assert.match(await stage(p), /^2\/3/); assert.equal(await focus(p), 'n'); assert.equal(await snap(p), before);
  assert.equal(await dlg(p).getByRole('checkbox', { name: '탭(T)', exact: true }).isChecked(), true);
  await p.keyboard.press('Enter'); assert.match(await stage(p), /^3\/3/); assert.equal(await focus(p), 'f'); assert.equal(await dlg(p).locator('input[value="general"]').isChecked(), true); assert.equal(await snap(p), before);
  await p.keyboard.press('Enter'); assert.equal(await dlg(p).count(), 0); assert.equal((await cell(p, 0, 0)).value, 12); assert.equal((await cell(p, 0, 1)).value, 34); assert.equal((await cell(p, 1, 1)).value, 78);
  assert.equal(await p.locator('#cellEditor').evaluate(n => n.classList.contains('idle') && n.value === ''), true);
  await p.evaluate(() => tabula.run('undo')); assert.equal(await snap(p), before); await p.evaluate(() => tabula.run('redo')); assert.equal((await cell(p, 1, 1)).value, 78);
 });
 await test('탭만 기본 선택·쉼표/공백 추측 제거·라디오 Enter는 다음 단계', async p => {
  await setup(p, ['앞,뒤 공백']); await open(p); await p.keyboard.press('Alt+d'); assert.equal(await dlg(p).locator('input[value="delimited"]').isChecked(), true);
  await p.keyboard.press('Enter'); assert.match(await stage(p), /^2\/3/);
  for (const key of ['m', 'c', 's', 'o', 'r']) assert.equal(await dlg(p).locator(`input[data-access-key="${key}"]`).isChecked(), false);
  await p.keyboard.press('Enter'); await p.keyboard.press('Enter'); assert.equal((await cell(p, 0, 0)).value, '앞,뒤 공백'); assert.equal((await cell(p, 0, 1)).raw, '');
 });
 await test('구분 기호 T/M/C/S/O/R/Q 명시키·재렌더 후 포커스·뒤로 상태 유지', async p => {
  await setup(p, ['a,b;c d|e']); await open(p); await p.keyboard.press('Alt+n');
  for (const key of ['t', 'm', 'c', 's', 'o', 'r']) { const input = dlg(p).locator(`input[data-access-key="${key}"]`), before = await input.isChecked(); await p.keyboard.press(`Alt+${key}`); assert.equal(await input.isChecked(), !before, key); assert.equal(await focus(p), key); }
  const other = dlg(p).getByRole('textbox', { name: '기타 구분 기호', exact: true }); await other.fill('|'); assert.equal(await other.evaluate(n => n === document.activeElement), true);
  await p.keyboard.press('Alt+q'); assert.equal(await focus(p), 'q'); await dlg(p).getByRole('combobox', { name: '텍스트 한정자', exact: true }).selectOption(''); assert.equal(await focus(p), 'q');
  await p.keyboard.press('Alt+b'); assert.match(await stage(p), /^1\/3/); assert.equal(await focus(p), 'n'); await p.keyboard.press('Alt+n'); assert.match(await stage(p), /^2\/3/); assert.equal(await other.inputValue(), '|');
  await next(p); await finish(p); assert.deepEqual(await p.evaluate(() => [0, 1, 2, 3, 4].map(c => tabula.wb().getValue(0, 0, c))), ['a', 'b', 'c', 'd', 'e']);
 });
 await test('원본 형식 D/W 명시키·고정 너비 유효성 실패 시 2단계 유지', async p => {
  await setup(p, ['abcdef']); await open(p); await p.keyboard.press('Alt+w'); assert.equal(await dlg(p).locator('input[value="fixed"]').isChecked(), true); assert.equal(await focus(p), 'w'); await p.keyboard.press('Enter'); assert.match(await stage(p), /열 구분선/);
  await next(p); assert.match(await stage(p), /^2\/3/); assert.match(await p.locator('#toast').textContent(), /구분선을 하나 이상/);
  await dlg(p).locator('.ttc-line.scale [data-i="3"]').click(); await p.keyboard.press('Alt+n'); assert.match(await stage(p), /^3\/3/); await finish(p); assert.equal((await cell(p, 0, 0)).value, 'abc'); assert.equal((await cell(p, 0, 1)).value, 'def');
 });
 await test('3단계 G/T/D/I·열 머리글 방향키·텍스트/날짜/건너뜀과 대상 보존', async p => {
  await setup(p, ['0012\t2024-03-15\t생략']); const before = await snap(p); await open(p); await next(p); await next(p);
  await p.keyboard.press('Alt+t'); assert.equal(await dlg(p).locator('input[value="text"]').isChecked(), true); assert.equal(await focus(p), 't');
  await dlg(p).locator('th[data-ttc-focus="column-0"]').focus(); await p.keyboard.press('ArrowRight'); assert.match(await dlg(p).locator('.fc-title').first().textContent(), /2번째/); await p.keyboard.press('Alt+d'); assert.equal(await dlg(p).locator('input[value="date"]').isChecked(), true); await dlg(p).getByRole('combobox', { name: '날짜 순서', exact: true }).selectOption('YMD');
  await dlg(p).locator('th[data-ttc-focus="column-1"]').focus(); await p.keyboard.press('End'); await p.keyboard.press('Alt+i'); assert.equal(await dlg(p).locator('input[value="skip"]').isChecked(), true); await p.keyboard.press('Alt+g'); assert.equal(await dlg(p).locator('input[value="general"]').isChecked(), true); await p.keyboard.press('Alt+i');
  await p.keyboard.press('Alt+e'); const dest = dlg(p).getByRole('textbox', { name: '대상', exact: true }); await dest.fill('$D$1'); await p.keyboard.press('Enter');
  assert.equal((await cell(p, 0, 0)).raw, '0012\t2024-03-15\t생략'); assert.equal((await cell(p, 0, 3)).value, '0012'); assert.equal((await cell(p, 0, 3)).style.numFmt, 'text'); assert.equal((await cell(p, 0, 4)).value, 45366); assert.equal((await cell(p, 0, 4)).style.numFmt, 'date'); assert.equal((await cell(p, 0, 5)).raw, ''); await p.evaluate(() => tabula.run('undo')); assert.equal(await snap(p), before);
 });
 await test('고급 D/T/M/R·잘못된 구분기호 차단·원래대로·확정·포커스 복귀', async p => {
  await setup(p, ['1.234,50-\t9']); await open(p); await next(p); await next(p); await p.keyboard.press('Alt+a');
  const adv = p.getByRole('dialog', { name: '텍스트 가져오기 고급 설정', exact: true }); assert.equal(await focus(p), 'd');
  await p.keyboard.press('Alt+d'); await adv.getByRole('combobox', { name: '소수 구분 기호', exact: true }).selectOption(','); await p.keyboard.press('Enter'); assert.equal(await adv.count(), 1); assert.match(await p.locator('#toast').textContent(), /달라야/);
  await p.keyboard.press('Alt+t'); await adv.getByRole('combobox', { name: '1000 단위 구분 기호', exact: true }).selectOption('.'); await p.keyboard.press('Alt+m'); assert.equal(await adv.getByRole('checkbox').isChecked(), false); await p.keyboard.press('Alt+r'); assert.equal(await adv.getByRole('combobox', { name: '소수 구분 기호', exact: true }).inputValue(), '.'); assert.equal(await adv.getByRole('combobox', { name: '1000 단위 구분 기호', exact: true }).inputValue(), ','); assert.equal(await adv.getByRole('checkbox').isChecked(), true);
  await adv.getByRole('combobox', { name: '소수 구분 기호', exact: true }).selectOption(','); await adv.getByRole('combobox', { name: '1000 단위 구분 기호', exact: true }).selectOption('.'); await adv.getByRole('button', { name: '확인', exact: true }).click(); assert.equal(await focus(p), 'a'); await p.keyboard.press('Alt+f'); assert.equal((await cell(p, 0, 0)).value, -1234.5); assert.equal((await cell(p, 0, 1)).value, 9);
 });
 await test('고급 취소는 숫자 옵션 보존·마법사 취소는 모든 셀 불변', async p => {
  await setup(p, ['1,234.50-\t9']); const before = await snap(p); await open(p); await next(p); await next(p); await p.keyboard.press('Alt+a'); let adv = p.getByRole('dialog', { name: '텍스트 가져오기 고급 설정', exact: true }); await adv.getByRole('combobox', { name: '소수 구분 기호', exact: true }).selectOption(','); await adv.getByRole('checkbox').uncheck(); await adv.getByRole('button', { name: '취소', exact: true }).click(); await p.keyboard.press('Alt+a'); adv = p.getByRole('dialog', { name: '텍스트 가져오기 고급 설정', exact: true }); assert.equal(await adv.getByRole('combobox', { name: '소수 구분 기호', exact: true }).inputValue(), '.'); assert.equal(await adv.getByRole('checkbox').isChecked(), true); await p.keyboard.press('Escape'); await p.keyboard.press('Escape'); assert.equal(await snap(p), before);
 });
 await test('덮어쓰기 취소는 원본 보존·확정은 한 번 Undo', async p => {
  await setup(p, ['12\t34'], { '0,1': { raw: '기존 데이터' } }); const before = await snap(p); await open(p); await next(p); await next(p); await finish(p); let overwrite = p.getByRole('dialog', { name: 'WIXEL', exact: true }); await overwrite.getByRole('button', { name: '취소', exact: true }).click(); assert.equal(await snap(p), before);
  await open(p); await next(p); await next(p); await finish(p); overwrite = p.getByRole('dialog', { name: 'WIXEL', exact: true }); await overwrite.getByRole('button', { name: '확인', exact: true }).click(); assert.equal((await cell(p, 0, 1)).value, 34); await p.evaluate(() => tabula.run('undo')); assert.equal(await snap(p), before);
 });
 await test('한글 물리코드 Alt A E 및 늦은 조합 뒤 단계 Enter는 셀로 새지 않음', async p => {
  await setup(p, ['12\t34']); const before = await snap(p); await p.keyboard.press('Alt');
  for (const [code, key] of [['KeyA', 'ㅁ'], ['KeyE', 'ㄷ']]) assert.equal(await p.evaluate(({ code, key }) => { const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, code, keyCode: 229, isComposing: true }); document.activeElement.dispatchEvent(e); return e.defaultPrevented; }, { code, key }), true);
  await dlg(p).waitFor(); assert.equal(await focus(p), 'n'); await p.evaluate(() => { const input = document.getElementById('cellEditor'); input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })); input.value = 'ㅁㄷ'; input.dispatchEvent(new InputEvent('input', { bubbles: true, data: 'ㅁㄷ', inputType: 'insertCompositionText', isComposing: true })); input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'ㅁㄷ' })); });
  assert.equal(await snap(p), before); assert.equal(await p.locator('#cellEditor').inputValue(), ''); await p.keyboard.press('Enter'); assert.match(await stage(p), /^2\/3/); assert.equal(await snap(p), before); await p.keyboard.press('Escape');
 });
 await test('320/390 화면의 3단계 탐색·footer 접근·Esc 취소', async p => {
  for (const width of [320, 390]) { await p.setViewportSize({ width, height: 640 }); await setup(p, ['1\t2']); await open(p); for (let step = 1; step <= 3; step++) { const rect = await dlg(p).boundingBox(); assert.ok(rect.x >= 0 && rect.x + rect.width <= width + 1); for (const button of await dlg(p).locator('.dialog-foot button').all()) { const r = await button.boundingBox(); assert.ok(r.x >= 0 && r.x + r.width <= width + 1 && r.y >= 0 && r.y + r.height <= 641, JSON.stringify(r)); } if (step < 3) await p.keyboard.press('Enter'); } await p.keyboard.press('Escape'); assert.equal((await cell(p, 0, 0)).raw, '1\t2'); }
 });
} finally { await browser.close(); }
console.log(JSON.stringify({ cases: results.length, passed: results.filter(x => x.ok).length, results }, null, 2));
if (results.some(x => !x.ok)) process.exitCode = 1;
