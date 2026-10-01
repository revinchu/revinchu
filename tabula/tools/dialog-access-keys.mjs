// 공통 UI 모듈을 실제 소스 앱에서 호출하는 합성 접근키 회귀. 소스 서버 전용(/src/ui.js 필요).
// 모든 컨텍스트를 격리하고 외부/서버 쓰기를 차단한다. Excel 전체 키 배치 동등성 검사는 아니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
assert.ok(['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname), '로컬 합성 검사만 허용');
const browser = await chromium.launch(), results = [];
const openUi = (page, fn) => page.evaluate(async (source) => {
  const ui = await import('/src/ui.js');
  return new Function('ui', `return (${source})(ui)`)(ui);
}, fn.toString());
const calls = (page) => page.evaluate(() => window.accessCalls ?? []);
async function pressAccess(page, target) {
  if (await page.locator('.access-key-layer').count()) await page.keyboard.press('Escape');
  await page.keyboard.press('Alt');
  const key = await target.getAttribute('data-resolved-access-key'); assert.ok(key, '표시 중 컨트롤에 접근키가 없음');
  await page.keyboard.press(key);
}
async function test(name, action) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage(), errors = [], blockedWrites = [];
  page.setDefaultTimeout(10000); page.on('pageerror', (e) => errors.push(e.message));
  await context.route('**/*', (route) => {
    const req = route.request();
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method()) || new URL(req.url()).origin !== new URL(url).origin) {
      blockedWrites.push(req.method() + ' ' + req.url()); return route.abort();
    }
    return route.continue();
  });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.accessCalls = []; });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => !!window.tabula?.wb());
    await action(page);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(blockedWrites, [], '외부 요청/서버 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors, blockedWrites }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
try {
  await test('공통 확인 O·취소 C·닫기 D 기본값과 명시 키 우선', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '기본 단추', body: el('input'), buttons: [{ label: '확인', action: () => { window.accessCalls.push('ok'); return false; } }, { label: '취소', action: () => { window.accessCalls.push('cancel'); return false; } }] }));
    await p.keyboard.press('Alt+o'); await p.keyboard.press('Alt+c'); assert.deepEqual(await calls(p), ['ok', 'cancel']);
    await p.keyboard.press('Alt+d'); assert.equal(await p.getByRole('dialog').count(), 0);
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '명시 우선', body: el('input'), buttons: [{ label: '확인', accessKey: 'x', action: () => window.accessCalls.push('explicit') }, { label: '닫기' }] }));
    assert.equal(await p.locator('.dialog-head button').getAttribute('data-access-key'), 'none');
    assert.equal(await p.locator('.dialog-foot button').first().getAttribute('aria-keyshortcuts'), 'Alt+X');
    await p.keyboard.press('Alt+x'); assert.deepEqual(await calls(p), ['ok', 'cancel', 'explicit']);
  });
  await test('body의 닫기 D와 header X는 동일 동작 중복 키로 순환하지 않음', async (p) => {
    await openUi(p, ({ openDialog, el }) => { let d; d = openDialog({ title: '본문 닫기', body: el('button', { onclick: () => d.close() }, '닫기') }); });
    assert.equal(await p.locator('.dialog-head button').getAttribute('data-access-key'), 'none');
    assert.equal(await p.locator('.dialog-head button').getAttribute('aria-keyshortcuts'), null);
    await p.keyboard.press('Alt+d'); assert.equal(await p.getByRole('dialog').count(), 0);
    await openUi(p, ({ openDialog, el }) => { let d; d = openDialog({ title: '동적 닫기', body: el('button', { id: 'dynamicClose', onclick: () => d.close() }, '닫기') }); document.querySelector('#dynamicClose').hidden = true; });
    await p.keyboard.press('Alt'); assert.equal(await p.locator('.dialog-head button').getAttribute('aria-keyshortcuts'), 'Alt+D');
    await p.keyboard.press('d'); assert.equal(await p.getByRole('dialog').count(), 0);
  });
  await test('Alt 배지와 순차 키: 지정 명령 한 번 실행, 리본 키팁·레이아웃 불변', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '접근키 합성', body: el('input', { id: 'accessText', accessKey: 'n', value: '원문' }), buttons: [{ label: '실행', accessKey: 'a', action: () => { window.accessCalls.push('run'); return false; } }] }));
    const before = await p.getByRole('dialog').boundingBox();
    await p.keyboard.press('Alt');
    assert.ok(await p.locator('.access-key-badge').count() >= 3);
    assert.equal(await p.locator('.keytip-panel').count(), 0);
    assert.deepEqual(await p.getByRole('dialog').boundingBox(), before);
    await p.keyboard.press('a'); assert.deepEqual(await calls(p), ['run']);
    assert.equal(await p.locator('#accessText').inputValue(), '원문');
    assert.equal(await p.locator('.access-key-layer').count(), 0);
  });
  await test('formDialog 레이블·명시 키: 입력 선택, 확인란 1회, select 포커스', async (p) => {
    await openUi(p, ({ formDialog }) => formDialog('필드 접근키', [{ name: 'n', label: '찾을 항목(&N)', value: '원문' }, { name: 'c', label: '검사(K)', type: 'checkbox', value: false }, { name: 't', label: '대상', accessKey: 't', type: 'select', value: '1', options: [{ value: '1', label: '하나' }, { value: '2', label: '둘' }] }], (v) => { window.accessCalls.push(v); return false; }));
    await p.keyboard.press('Alt+n'); assert.equal(await p.evaluate(() => document.activeElement.value), '원문');
    assert.deepEqual(await p.evaluate(() => [document.activeElement.selectionStart, document.activeElement.selectionEnd]), [0, 2]);
    await p.keyboard.press('Alt+k'); assert.equal(await p.getByRole('checkbox').isChecked(), true);
    await p.keyboard.press('Alt+t'); assert.equal(await p.evaluate(() => document.activeElement.tagName), 'SELECT');
    await p.keyboard.press('ArrowDown'); assert.equal(await p.getByRole('dialog').getByRole('combobox').inputValue(), '2');
  });
  await test('중복 명시 키는 포커스 순환, Enter에서 선택 항목만 실행', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '중복 키', body: el('input', { accessKey: 'n' }), buttons: ['첫째', '둘째'].map((label) => ({ label, accessKey: 'a', action: () => { window.accessCalls.push(label); return false; } })) }));
    await p.keyboard.press('Alt+a'); assert.equal(await p.evaluate(() => document.activeElement.textContent), '첫째');
    await p.keyboard.press('Alt+a'); assert.equal(await p.evaluate(() => document.activeElement.textContent), '둘째');
    assert.deepEqual(await calls(p), []); await p.keyboard.press('Enter'); assert.deepEqual(await calls(p), ['둘째']);
  });
  await test('숨김·disabled·disabled fieldset 컨트롤은 키 대상에서 제외', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '비활성 제외', body: el('div', {}, el('button', { hidden: true, accessKey: 'a', onclick: () => window.accessCalls.push('hidden') }, '숨김'), el('fieldset', { disabled: true }, el('button', { accessKey: 'a', onclick: () => window.accessCalls.push('fieldset') }, '비활성')), el('button', { disabled: true, accessKey: 'a', onclick: () => window.accessCalls.push('disabled') }, '비활성2')), buttons: [{ label: '활성', accessKey: 'a', action: () => { window.accessCalls.push('active'); return false; } }] }));
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['active']);
  });
  await test('최상위 modal만 처리하고 닫은 뒤 하위 창 접근키 복귀', async (p) => {
    await openUi(p, ({ openDialog, el }) => { openDialog({ title: '하위 창', body: el('div'), buttons: [{ label: '하위 실행', accessKey: 'a', action: () => { window.accessCalls.push('lower'); return false; } }] }); openDialog({ title: '상위 창', body: el('div'), buttons: [{ label: '상위 실행', accessKey: 'a', action: () => { window.accessCalls.push('upper'); } }] }); });
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['upper']);
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['upper', 'lower']);
  });
  await test('modeless 창은 내부 포커스에서만 처리, 격자 Alt는 리본으로', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: 'modeless 합성', modeless: true, body: el('input', { id: 'modelessInput', accessKey: 'n' }), buttons: [{ label: '동작', accessKey: 'a', action: () => { window.accessCalls.push('modeless'); return false; } }] }));
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['modeless']);
    await p.evaluate(() => window.tabula.selectCell(0, 0)); await p.locator('#cellEditor').focus();
    await p.keyboard.press('Alt'); assert.equal(await p.locator('.access-key-layer').count(), 0); assert.equal(await p.locator('body.keytips').count(), 1);
  });
  await test('하위 메뉴 접근키와 3단계 계층, ArrowLeft 부모 복귀', async (p) => {
    await openUi(p, ({ openMenu }) => openMenu({ x: 300, y: 260 }, [{ label: '부모', accessKey: 'p', submenu: [{ label: '중간', accessKey: 'm', submenu: [{ label: '끝 실행', accessKey: 'r', action: () => window.accessCalls.push('leaf') }] }, { label: '중간 동작', accessKey: 'r', action: () => window.accessCalls.push('middle') }] }]));
    await p.keyboard.press('Alt+p'); assert.equal(await p.locator('#menuLayer .menu').count(), 2);
    await p.keyboard.press('m'); assert.equal(await p.locator('#menuLayer .menu').count(), 3);
    await p.keyboard.press('ArrowLeft'); assert.equal(await p.locator('#menuLayer .menu').count(), 2);
    await p.keyboard.press('r'); assert.deepEqual(await calls(p), ['middle']); assert.equal(await p.locator('#menuLayer .menu').count(), 0);
  });
  await test('새 modal은 이전 메뉴의 접근키보다 우선', async (p) => {
    await openUi(p, ({ openMenu, openDialog, el }) => { openMenu({ x: 40, y: 300 }, [{ label: '배경 메뉴', accessKey: 'a', action: () => window.accessCalls.push('menu') }]); openDialog({ title: '메뉴 위 modal', body: el('div'), buttons: [{ label: '상위 명령', accessKey: 'a', action: () => { window.accessCalls.push('dialog'); return false; } }] }); });
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['dialog']);
  });
  await test('대화상자 내부에서 연 메뉴는 대화상자보다 우선', async (p) => {
    await openUi(p, ({ openMenu, openDialog, el }) => openDialog({ title: '메뉴 포함 창', body: el('button', { accessKey: 'p', onclick: (event) => openMenu(event.currentTarget, [{ label: '메뉴 실행', accessKey: 'a', action: () => window.accessCalls.push('menu') }]) }, '메뉴'), buttons: [{ label: '창 명령', accessKey: 'a', action: () => { window.accessCalls.push('dialog'); return false; } }] }));
    await p.keyboard.press('Alt+p'); await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['menu']);
  });
  await test('커스텀 node 메뉴 입력·옵션·단추도 처리하며 입력 글자는 명령으로 실행하지 않음', async (p) => {
    await openUi(p, ({ openMenu, el }) => openMenu({ x: 350, y: 280 }, [{ node: el('div', {}, el('label', {}, '검색(&N)', el('input', { id: 'nodeInput' })), el('div', { role: 'option', tabindex: 0, accessKey: 's', onclick: () => window.accessCalls.push('option') }, '목록 선택'), el('button', { accessKey: 'r', onclick: () => window.accessCalls.push('button') }, '적용')) }]));
    await p.keyboard.press('Alt+n'); await p.keyboard.type('sample'); assert.equal(await p.locator('#nodeInput').inputValue(), 'sample'); assert.deepEqual(await calls(p), []);
    await p.keyboard.press('Alt+s'); await p.keyboard.press('Alt+r'); assert.deepEqual(await calls(p), ['option', 'button']);
  });
  await test('한글 physical code+IME229 Alt 명령 및 AltGr 제외', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '한글 접근키', body: el('input', { id: 'imeInput', value: '유지', accessKey: 'n' }), buttons: [{ label: '실행', accessKey: 'a', action: () => { window.accessCalls.push('ime'); return false; } }] }));
    const consumed = await p.evaluate(() => { const e = new KeyboardEvent('keydown', { key: 'ㅁ', code: 'KeyA', altKey: true, isComposing: true, keyCode: 229, bubbles: true, cancelable: true }); document.activeElement.dispatchEvent(e); return e.defaultPrevented; });
    assert.equal(consumed, true); assert.deepEqual(await calls(p), ['ime']);
    const altGr = await p.evaluate(() => { const e = new KeyboardEvent('keydown', { key: 'a', code: 'KeyA', ctrlKey: true, altKey: true, bubbles: true, cancelable: true }); document.activeElement.dispatchEvent(e); return e.defaultPrevented; });
    assert.equal(altGr, false); assert.deepEqual(await calls(p), ['ime']); assert.equal(await p.locator('#imeInput').inputValue(), '유지');
  });
  await test('Escape는 배지만 취소 후 창 닫기, 일반 입력은 보존', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '취소 단계', body: el('input', { id: 'normalInput', accessKey: 'n' }) }));
    await p.keyboard.press('Alt'); await p.keyboard.press('Escape'); assert.equal(await p.getByRole('dialog').count(), 1); assert.equal(await p.locator('.access-key-layer').count(), 0);
    await p.keyboard.type('hello'); assert.equal(await p.locator('#normalInput').inputValue(), 'hello');
    await p.keyboard.press('Escape'); assert.equal(await p.getByRole('dialog').count(), 0);
  });
  await test('별칭도 같은 명령을 실행하고 닫힌 뒤 keyup은 상위로 새지 않음', async (p) => {
    await openUi(p, ({ openDialog, el }) => { window.accessUp = 0; document.addEventListener('keyup', () => window.accessUp++); openDialog({ title: '닫기 별칭', body: el('input'), buttons: [{ label: '닫기', accessKey: 'd', accessAliases: 'c', action: () => window.accessCalls.push('close') }] }); });
    assert.equal(await p.getByRole('button', { name: '닫기', exact: true }).last().getAttribute('aria-keyshortcuts'), 'Alt+D Alt+C');
    await p.keyboard.press('Alt+c'); assert.deepEqual(await calls(p), ['close']); assert.equal(await p.getByRole('dialog').count(), 0); assert.equal(await p.evaluate(() => window.accessUp), 0);
  });
  await test('비동기 확인 중 접근키 중복 호출을 막고 완료 후 정상 복귀', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '비동기 접근키', body: el('input'), buttons: [{ label: '작업', accessKey: 'a', action: () => { window.accessCalls.push('start'); return new Promise((resolve) => { window.accessRelease = () => resolve(false); }); } }] }));
    await p.keyboard.press('Alt+a'); await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['start']);
    assert.equal(await p.getByRole('dialog').getAttribute('aria-busy'), 'true'); await p.evaluate(() => window.accessRelease());
    await p.waitForFunction(() => !document.querySelector('.dialog').hasAttribute('aria-busy'));
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['start', 'start']); await p.evaluate(() => window.accessRelease());
  });
  await test('한국어 Excel 레이블 추천과 자동 키는 출처 구분·명시 키 우선', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '키 추천', body: el('div', {}, el('label', {}, el('span', {}, '암호'), el('input', { id: 'nativePassword', type: 'password' })), el('label', {}, '글꼴 크기', el('select', { id: 'nativeSize' }, el('option', {}, '작게'), el('option', {}, '크게'))), el('button', { id: 'nativeTitle', title: '계열' }, '◎'), el('button', { accessKey: 'f', onclick: () => window.accessCalls.push('explicit') }, '명시 명령'), el('button', { id: 'fontHint', onclick: () => window.accessCalls.push('font') }, '글꼴'), el('button', { id: 'autoHint' }, '합성 신규 명령')) }));
    assert.equal(await p.locator('#nativePassword').getAttribute('data-access-key-source'), 'excel');
    assert.equal(await p.locator('#fontHint').getAttribute('data-access-key-source'), 'wixel');
    assert.equal(await p.locator('#autoHint').getAttribute('data-access-key-source'), 'wixel');
    assert.equal(await p.locator('#nativeSize').getAttribute('data-resolved-access-key'), 'z');
    assert.equal(await p.locator('#nativeTitle').getAttribute('data-resolved-access-key'), 's');
    await p.keyboard.press('Alt+p'); assert.equal(await p.evaluate(() => document.activeElement.id), 'nativePassword');
    await p.keyboard.press('Alt+f'); assert.deepEqual(await calls(p), ['explicit']);
  });
  await test('스크롤 밖·숨긴 탭 제외 후 새로 표시한 동적 필드 접근 가능', async (p) => {
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '동적 표시', body: el('div', {}, el('button', { id: 'dynamicTarget', hidden: true, accessKey: 'z', onclick: () => window.accessCalls.push('dynamic') }, '동적 대상'), el('div', { style: { height: '50px', overflow: 'auto' } }, el('div', { style: { height: '120px' } }), el('button', { accessKey: 'z', onclick: () => window.accessCalls.push('outside') }, '스크롤 밖'))), buttons: [{ label: '보이기', accessKey: 'v', action: () => { document.querySelector('#dynamicTarget').hidden = false; return false; } }] }));
    await p.keyboard.press('Alt+z'); assert.deepEqual(await calls(p), []);
    await p.keyboard.press('Alt+v'); await p.keyboard.press('Alt+z'); assert.deepEqual(await calls(p), ['dynamic']);
  });
  await test('실제 셀 서식 창에서 Alt 배지 표시 및 폼 필드 포커스', async (p) => {
    await p.keyboard.press('Control+1'); await p.getByRole('dialog').waitFor(); await p.getByRole('dialog').getByRole('tab', { name: '맞춤', exact: true }).click(); await p.keyboard.press('Alt');
    const data = await p.locator('.access-key-badge').evaluateAll((nodes) => nodes.map((n) => ({ key: n.dataset.key, r: n.getBoundingClientRect().toJSON() })));
    assert.ok(data.length >= 5); const viewport = p.viewportSize(); assert.ok(data.every(({ r }) => r.left >= 0 && r.top >= 0 && r.right <= viewport.width && r.bottom <= viewport.height));
    const uniqueInput = await p.getByRole('dialog').evaluate((d) => { const controls = [...d.querySelectorAll('[data-resolved-access-key]')].filter((n) => n.getClientRects().length && !n.disabled); return controls.find((n) => n.matches('input,select,textarea') && controls.filter((x) => x.dataset.resolvedAccessKey === n.dataset.resolvedAccessKey).length === 1)?.dataset.resolvedAccessKey; });
    assert.ok(uniqueInput); await p.keyboard.press(uniqueInput);
    assert.ok(['INPUT', 'SELECT', 'TEXTAREA'].includes(await p.evaluate(() => document.activeElement.tagName)));
    await p.keyboard.press('Alt');
    if (process.env.WIXEL_ACCESS_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_ACCESS_SCREENSHOT });
  });
  await test('빠른 분석 독립 팝업: Alt 범주 전환과 합계가 실제 셀에 계산', async (p) => {
    await p.evaluate(() => { const t = window.tabula, wb = t.wb(); wb.transact(() => [['항목', '값'], ['가', '10'], ['나', '20']].forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, value)))); t.selectRange({ r1: 0, c1: 0, r2: 2, c2: 1 }); t.run('quickAnalysis'); });
    const pop = p.locator('.qa-pop'); await pop.waitFor();
    await pressAccess(p, pop.locator('.qa-tab').filter({ hasText: /^합계$/ }));
    await pressAccess(p, pop.locator('.qa-tile[title="합계"]'));
    assert.equal(await pop.count(), 0); assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 3, 1)), 30);
  });
  await test('셀 무늬 독립 팝업: modal 소유 범위와 Alt 선택·셀 반영', async (p) => {
    await p.keyboard.press('Control+1'); const dialog = p.getByRole('dialog', { name: '셀 서식', exact: true }); await dialog.getByRole('tab', { name: '채우기', exact: true }).click();
    await dialog.locator('.pat-drop').click(); const popup = p.locator('.pat-pop'); await popup.waitFor();
    await pressAccess(p, popup.locator('.pat-cell').nth(1));
    assert.equal(await popup.count(), 0); assert.equal(await dialog.count(), 1);
    await dialog.locator('.dialog-foot .primary').click();
    assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).pattern), 'darkGray');
  });
  await test('파일 허브 독립 화면: Alt 탐색·검색 입력 유지·중첩 modal 우선·Esc 닫기', async (p) => {
    await p.evaluate(() => window.tabula.run('saveLocations')); const hub = p.locator('.backstage'); await hub.waitFor();
    await pressAccess(p, hub.locator('button[data-page="templates"]'));
    const search = hub.locator('input[type="search"]'); await search.fill('보고서'); assert.equal(await search.inputValue(), '보고서');
    await openUi(p, ({ openDialog, el }) => openDialog({ title: '허브 위 modal', body: el('div'), buttons: [{ label: '작업', accessKey: 'a', action: () => window.accessCalls.push('modal') }] }));
    await p.keyboard.press('Alt+a'); assert.deepEqual(await calls(p), ['modal']); assert.equal(await hub.count(), 1);
    if (await p.locator('.access-key-layer').count()) { await p.keyboard.press('Escape'); assert.equal(await hub.count(), 1); }
    await p.keyboard.press('Escape'); assert.equal(await hub.count(), 0);
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ suite: 'dialog-access-keys', sourceOnly: true, passed: results.filter((r) => r.ok).length, total: results.length, results }, null, 2));
if (results.some((r) => !r.ok)) process.exitCode = 1;
