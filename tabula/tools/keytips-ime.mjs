// 물리 키 code + 한글 IME 지연 이벤트 회귀. 실제 Windows IME 자체를 구동하는 검사는 아니다.
// 새 컨텍스트의 합성 문서만 사용하며 소스/최종 번들 공통이다.
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
      w.transact(() => { w.setInput(0, 1, 1, '보존할 값'); w.setInput(0, 2, 2, '=SUM(2,3)'); });
      t.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
      window.__imeKey = (type, code, key = 'Process', extra = {}) => {
        const e = new KeyboardEvent(type, { bubbles: true, cancelable: true, code, key, keyCode: 229, isComposing: true, ...extra });
        (document.activeElement || document.getElementById('cellEditor')).dispatchEvent(e); return e.defaultPrevented;
      };
      window.__imeText = (text, { start = true, end = true, cancelable = false, target = '#cellEditor' } = {}) => {
        const e = document.querySelector(target);
        if (start) e.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
        e.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: text }));
        const before = new InputEvent('beforeinput', { bubbles: true, cancelable, inputType: 'insertCompositionText', data: text, isComposing: true });
        e.dispatchEvent(before);
        if (!before.defaultPrevented) { e.value = text; e.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: text, isComposing: true })); }
        if (end) {
          e.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: text }));
          const final = new InputEvent('beforeinput', { bubbles: true, cancelable, inputType: 'insertText', data: text, isComposing: false });
          e.dispatchEvent(final);
          if (!final.defaultPrevented) { e.value = text; e.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: text, isComposing: false })); }
        }
        return before.defaultPrevented;
      };
    });
    await p.locator('#cellEditor').focus(); await check(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
const state = p => p.evaluate(() => ({
  raw: window.tabula.wb().getRaw(0, 1, 1), formula: window.tabula.wb().getRaw(0, 2, 2), sel: window.tabula.sel,
  editor: document.getElementById('cellEditor').value, editing: !document.getElementById('cellEditor').classList.contains('idle'),
  noGrid: !!window.tabula.wb().sheets[0].noGrid, keytips: document.body.classList.contains('keytips'), seq: document.body.dataset.keytipSequence,
}));
const key = (p, code, text, extra) => p.evaluate(({ code, text, extra }) => window.__imeKey('keydown', code, text, extra), { code, text, extra });
const up = (p, code, text, extra) => p.evaluate(({ code, text, extra }) => window.__imeKey('keyup', code, text, extra), { code, text, extra });
const ime = (p, text = 'ㅊㅎ', opts = {}) => p.evaluate(({ text, opts }) => window.__imeText(text, opts), { text, opts });
const settle = p => p.evaluate(() => new Promise(resolve => setTimeout(resolve, 60)));
const unchanged = async (p, before, hidden = !before.noGrid) => {
  const s = await state(p); assert.equal(s.raw, before.raw); assert.equal(s.formula, before.formula); assert.deepEqual(s.sel, before.sel);
  assert.equal(s.noGrid, hidden); assert.equal(s.editing, false, '키팁의 후속 IME 이벤트가 셀 편집에 진입함'); assert.equal(s.editor, '', '키팁 조합 문자가 편집기에 남음');
};
async function wvg(p, { held = false, prefixes = false } = {}) {
  if (held) await p.keyboard.down('Alt'); else await p.keyboard.press('Alt');
  for (const [code, text] of [['KeyW', 'ㅈ'], ['KeyV', 'ㅍ'], ['KeyG', 'ㅎ']]) {
    assert.equal(await key(p, code, text, { altKey: held }), true, code + ' 키팁 처리');
    if (prefixes && code !== 'KeyG') await ime(p, text, { end: false });
  }
}
async function normalHangul(p, text = '한글 정상 입력') {
  await key(p, 'KeyG', 'Process'); await ime(p, text); await up(p, 'KeyG', 'Process');
  await p.keyboard.press('Enter');
}
try {
  await test('한글 W/V 조합 중 입력과 마지막 G 후 비취소 input이 눈금선 명령에 섞이지 않음', async p => {
    const before = await state(p); await wvg(p, { prefixes: true });
    await ime(p, 'ㅊㅎ', { start: false }); await up(p, 'KeyG', 'ㅎ'); await settle(p); await unchanged(p, before);
  });
  await test('명령 종료 뒤 compositionstart가 늦게 도착해도 셀 값·선택 범위 보존', async p => {
    const before = await state(p); await wvg(p); await ime(p); await up(p, 'KeyG', 'ㅎ'); await settle(p); await unchanged(p, before);
  });
  await test('취소 가능한 beforeinput은 차단되고 지연 keyup 뒤 반복 눈금선 전환 가능', async p => {
    const before = await state(p); await wvg(p);
    assert.equal(await ime(p, 'ㅊㅎ', { cancelable: true }), true); await up(p, 'KeyG', 'ㅎ'); await settle(p); await unchanged(p, before);
    await wvg(p); await ime(p, 'ㅊㅎ', { cancelable: false }); await up(p, 'KeyG', 'ㅎ'); await unchanged(p, before, before.noGrid);
  });
  await test('Alt를 누른 채 순서 입력 후 Alt keyup·조합 종료 지연도 셀 오염 없음', async p => {
    const before = await state(p); await wvg(p, { held: true, prefixes: true });
    await up(p, 'KeyG', 'ㅎ', { altKey: true }); await p.keyboard.up('Alt'); await ime(p, 'ㅊㅎ'); await settle(p); await unchanged(p, before);
  });
  await test('키팁 잔여 입력을 버린 뒤 다음 정상 한글 조합은 그대로 확정', async p => {
    const before = await state(p); await wvg(p); await ime(p); await up(p, 'KeyG', 'ㅎ'); await unchanged(p, before);
    await normalHangul(p); assert.equal((await state(p)).raw, '한글 정상 입력'); assert.equal((await state(p)).formula, before.formula);
  });
  await test('키팁 종료 뒤 새 영문 키 입력도 버리지 않고 셀에 정상 입력', async p => {
    const before = await state(p); await wvg(p); await ime(p); await unchanged(p, before);
    await p.keyboard.type('normal'); await p.keyboard.press('Enter'); assert.equal((await state(p)).raw, 'normal');
  });
  await test('키팁 하위 테두리 메뉴 포커스·Escape 복귀 중 늦은 IME 입력 차단', async p => {
    const before = await state(p); await p.keyboard.press('Alt'); await key(p, 'KeyH', 'ㅗ'); await key(p, 'KeyB', 'ㅠ');
    await p.locator('.keytip-command-menu').waitFor(); await ime(p, 'ㅗㅠ');
    await unchanged(p, before, before.noGrid); await p.keyboard.press('Escape'); assert.equal((await state(p)).seq, 'h');
    await p.keyboard.press('Escape'); await p.keyboard.press('Escape'); await p.locator('#cellEditor').focus(); await normalHangul(p, '메뉴 뒤 한글');
    assert.equal((await state(p)).raw, '메뉴 뒤 한글');
  });
  await test('키팁으로 연 팝업에서 이전 격자의 지연 입력만 차단하고 팝업 한글은 보존', async p => {
    const before = await state(p); await p.keyboard.press('Alt'); await key(p, 'KeyN', 'ㅜ'); await key(p, 'KeyI', 'ㅑ');
    const dialog = p.getByRole('dialog', { name: '하이퍼링크 삽입', exact: true }); await dialog.waitFor();
    await ime(p, 'ㅜㅑ'); await unchanged(p, before, before.noGrid);
    const input = dialog.getByLabel('표시할 텍스트', { exact: true }); await input.focus();
    await input.evaluate(e => { e.dataset.imeTest = 'dialog-input'; });
    await key(p, 'KeyG', 'Process'); await ime(p, '팝업 한글', { target: '[data-ime-test="dialog-input"]' }); await up(p, 'KeyG', 'Process');
    assert.equal(await input.inputValue(), '팝업 한글'); await dialog.getByRole('button', { name: '취소', exact: true }).click();
    assert.equal((await state(p)).raw, before.raw);
  });
  await test('찾기 창 Alt+D 닫기 뒤 지연 한글 입력 차단·다음 정상 입력 보존', async p => {
    const before = await state(p); await p.keyboard.press('Control+f');
    const dialog = p.getByRole('dialog', { name: '찾기 및 바꾸기', exact: true }); await dialog.waitFor();
    assert.equal(await key(p, 'KeyD', 'ㅇ', { altKey: true }), true);
    await dialog.waitFor({ state: 'hidden' }); await ime(p, 'ㅇ'); await up(p, 'KeyD', 'ㅇ'); await settle(p);
    await unchanged(p, before, before.noGrid); await p.locator('#cellEditor').focus(); await normalHangul(p, '닫은 뒤 한글');
    assert.equal((await state(p)).raw, '닫은 뒤 한글');
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok), note: 'Chromium 합성 composition/beforeinput/input 및 물리 code 검사. 실제 Windows 한국어 IME는 별도 검증 필요.' }, null, 2));
if (results.some(x => !x.ok)) process.exitCode = 1;
