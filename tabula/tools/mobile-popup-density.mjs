// 합성 문서·격리 Chromium. 실제 iPhone Safari/OS 확대를 인증하는 도구는 아니다.
// WIXEL_MOBILE_DENSITY_OUT / WIXEL_MOBILE_DENSITY_FILTER로 출력 및 부분 실행을 선택한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
assert.ok(['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname), '로컬 합성 검사만 허용');
const origin = new URL(url).origin;
const out = process.env.WIXEL_MOBILE_DENSITY_OUT || 'D:/Codex/Temp/wixel-mobile-popups-dense';
const only = process.env.WIXEL_MOBILE_DENSITY_FILTER || '';
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = [], measurements = [], assets = new Set();
let checks = 0;
const eq = (actual, expected, message) => { checks++; assert.deepEqual(actual, expected, message); };
const ok = (actual, message) => { checks++; assert.ok(actual, message); };
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const snapshot = p => p.evaluate(() => JSON.stringify(window.tabula.wb().serialize()));
const menu = p => p.locator('#menuLayer > .menu[data-level="0"]');
const child = p => p.locator('#menuLayer > .menu[data-level="1"]');
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
const shot = (p, name) => p.screenshot({ path: out + '/' + name + '.png' });
async function inside(p, target, name) {
  await target.waitFor(); const b = await target.boundingBox(), v = await p.evaluate(() => ({ w: innerWidth, h: innerHeight }));
  ok(b && b.width > 0 && b.height > 0 && b.x >= -1 && b.y >= -1 && b.x + b.width <= v.w + 1 && b.y + b.height <= v.h + 1, name + ': ' + JSON.stringify({ b, v }));
  return b;
}
async function usable(p, target, name) {
  await target.scrollIntoViewIfNeeded(); await inside(p, target, name); eq(await target.isEnabled(), true, name + ' 활성');
  ok(await target.evaluate(n => { const r = n.getBoundingClientRect(), hit = document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2); return hit === n || n.contains(hit); }), name + ' 클릭점 가림 없음');
}
async function mobile(p, wanted) {
  const toggle = p.locator('#mobileModeToggle');
  if ((await toggle.getAttribute('aria-pressed') === 'true') !== wanted) await toggle.click();
  await raf(p); eq(await toggle.getAttribute('aria-pressed'), String(wanted), '실제 모바일 전환 버튼');
  eq(await p.evaluate(() => window.tabula.mobile().active), wanted, '화면 모드와 버튼 일치');
}
async function openCellMenu(p, row = 0, keyboard = false) {
  await p.evaluate(row => window.tabula.selectCell(row, 0), row);
  if (keyboard) { await p.locator('#gridView').focus(); await p.keyboard.press('Shift+F10'); }
  else {
    const b = await p.evaluate(row => { const r = window.tabula.gv().clientRect({ r1: row, c1: 0, r2: row, c2: 0 }); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }, row);
    await p.mouse.click(b.x, b.y, { button: 'right' });
  }
  await menu(p).waitFor(); await raf(p);
}
async function hiddenItems(p, hidden) {
  for (const name of [/^스마트 조회/, /^윗주 필드 표시/, /^윗주 편집/]) eq(await menu(p).getByRole('menuitem', { name }).count(), hidden ? 0 : 1, String(name));
}
async function rowDialog(p, keyboard = false) {
  const parent = menu(p).getByRole('menuitem', { name: /^행·열 크기 및 숨기기/ });
  if (keyboard) {
    await p.keyboard.press('End');
    eq(await parent.evaluate(n => n === document.activeElement), true, 'End로 실제 마지막 하위 메뉴 선택');
    await p.keyboard.press('ArrowRight');
  } else { await parent.scrollIntoViewIfNeeded(); await parent.click(); }
  await child(p).waitFor(); await inside(p, child(p), '행·열 하위 메뉴');
  const row = child(p).getByRole('menuitem', { name: /^행 높이\(R\)/ });
  if (keyboard) { await p.keyboard.press('Home'); await p.keyboard.press('Enter'); }
  else { await usable(p, row, '행 높이 메뉴'); await row.click(); }
  const d = dialog(p, '행 높이'); await d.waitFor(); return d;
}
async function cancel(p, d) { const b = d.getByRole('button', { name: '취소', exact: true }); await usable(p, b, '취소'); await b.click(); await d.waitFor({ state: 'detached' }); }
const metrics = target => target.evaluate(n => {
  const s = getComputedStyle(n), b = n.getBoundingClientRect();
  return { font: parseFloat(s.fontSize), height: b.height, width: b.width, top: parseFloat(s.paddingTop), bottom: parseFloat(s.paddingBottom), left: parseFloat(s.paddingLeft), right: parseFloat(s.paddingRight), transform: s.transform, zoom: s.zoom };
});
async function densitySample(p, name, isMobile) {
  await openCellMenu(p); await hiddenItems(p, isMobile);
  const item = await metrics(menu(p).getByRole('menuitem', { name: /^복사\(C\)/ }));
  await inside(p, menu(p), name + ' 메뉴'); await inside(p, p.locator('.context-mini-toolbar'), name + ' 미니');
  await shot(p, name + '-menu'); const d = await rowDialog(p);
  const label = await metrics(d.locator('.form-dialog-label').first()), body = await metrics(d.locator('.dialog-body'));
  const input = await metrics(d.getByRole('spinbutton')), button = await metrics(d.getByRole('button', { name: '취소', exact: true }));
  await inside(p, d, name + ' 폼'); await shot(p, name + '-form'); await cancel(p, d);
  const result = { name, item, label, body, input, button }; measurements.push(result); return result;
}
async function test(name, size, fn, touch = true) {
  if (only && !name.includes(only)) return;
  const context = await browser.newContext({ viewport: { width: size[0], height: size[1] }, hasTouch: touch, isMobile: touch });
  const p = await context.newPage(), errors = [], writes = [], blocked = [];
  p.setDefaultTimeout(15000); p.on('pageerror', e => errors.push(e.message));
  p.on('dialog', d => d.type() === 'beforeunload' ? d.accept() : d.dismiss());
  p.on('response', r => { if (/wixel-[a-f\d]+\.js/.test(r.url())) assets.add(r.url().split('/').pop()); });
  await context.addInitScript(() => {
    window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel:version', '3.0.0');
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {}, read: async () => [], write: async () => {} } });
  });
  await context.route('**/*', route => {
    const req = route.request(), u = new URL(req.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname); return route.abort(); }
    if (u.origin !== origin || u.pathname.startsWith('/api/')) { blocked.push(u.pathname); return route.abort(); }
    return route.continue();
  });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 }); await p.waitForFunction(() => !!window.tabula?.wb(), null, { timeout: 60000 });
    await p.evaluate(() => {
      const t = window.tabula; t.wb().restore({ sheets: [
        { name: '합성 검증', zoom: 125, cells: { '0,0': { raw: '합성 텍스트' }, '0,1': { raw: '원본 숫자' }, '1,0': { raw: '123.456' }, '1,1': { raw: '=1+1' }, '3,0': { raw: '합성 글꼴', style: { font: '합성 문서 글꼴' } } } },
        { name: '준비', cells: {} }
      ] }); t.switchSheet(1); t.switchSheet(0); t.selectCell(0, 0);
    });
    await raf(p); await fn(p); eq(errors, [], '페이지 오류 없음'); eq(writes, [], '원격 쓰기 없음');
    results.push({ name, size, ok: true, errors, writes, blocked }); console.log('OK ' + name);
  } catch (error) {
    results.push({ name, size, ok: false, error: error.message, errors, writes, blocked }); console.error('NG ' + name + ': ' + error.stack);
    await shot(p, 'failure-' + results.length).catch(() => {});
  } finally { await context.close(); }
}
try {
  await test('1366 실제 모드 OFF→ON→OFF 밀도 및 데스크톱 복원', [1366, 900], async p => {
    await mobile(p, false); const before = await snapshot(p), desktop = await densitySample(p, '1366-desktop', false);
    await mobile(p, true); const compact = await densitySample(p, '1366-mobile', true);
    ok(compact.item.font < desktop.item.font, '같은 메뉴 글자 축소'); ok(compact.item.height < desktop.item.height, '같은 메뉴 행 축소');
    ok(compact.label.font < desktop.label.font, '폼 라벨 글자 축소'); ok(compact.body.top + compact.body.bottom < desktop.body.top + desktop.body.bottom, '폼 내용 여백 축소');
    ok(compact.button.height < desktop.button.height, '폼 버튼 높이 축소'); ok(compact.input.font >= 16, '입력 원래 CSS 글꼴은 16px 이상');
    eq(compact.body.transform, 'none', '전체 팝업 transform 축소 없음'); eq(compact.body.zoom, '1', '전체 팝업 zoom 축소 없음');
    await mobile(p, false); const restored = await densitySample(p, '1366-restored', false);
    for (const part of ['item', 'label', 'body', 'input', 'button']) for (const key of ['font', 'height', 'top', 'bottom', 'left', 'right']) ok(Math.abs(restored[part][key] - desktop[part][key]) <= 0.5, '복원 ' + part + '.' + key);
    eq(await snapshot(p), before, '토글·메뉴·취소는 문서와 저장 배율 불변');
  }, false);
  for (const size of [[320, 740], [844, 390]]) {
    const width = size[0];
    await test(width + ' 빈칸·텍스트 메뉴 숨김, 마우스·키보드 하위 메뉴', size, async p => {
      await mobile(p, false); await mobile(p, true); const before = await snapshot(p);
      for (const row of [2, 0]) {
        await openCellMenu(p, row); await hiddenItems(p, true); const m = await inside(p, menu(p), '셀 메뉴'), b = await inside(p, p.locator('.context-mini-toolbar'), '미니');
        ok(b.y + b.height <= m.y + 1 || m.y + m.height <= b.y + 1, '미니와 메뉴 겹침 없음');
        eq(await p.locator('.context-mini-toolbar .access-key-hint').count(), 0, '미니 상시 접근키 숨김 유지');
        await shot(p, width + '-context-' + (row ? 'blank' : 'text')); await p.keyboard.press('Escape');
      }
      for (const keyboard of [false, true]) {
        await openCellMenu(p, 0, keyboard); const d = await rowDialog(p, keyboard); await inside(p, d, '행 높이 폼');
        const input = d.getByRole('spinbutton'); await usable(p, input, '행 높이 입력'); ok((await metrics(input)).font >= 16, '입력 확대 방지 원래 글꼴');
        await input.fill('37.5'); await shot(p, width + '-row-' + (keyboard ? 'keyboard' : 'mouse')); await cancel(p, d);
      }
      eq(await snapshot(p), before, '메뉴 탐색과 폼 취소는 문서 불변');
      await mobile(p, false); await openCellMenu(p); await hiddenItems(p, false); await p.keyboard.press('Escape');
      eq(await snapshot(p), before, '모바일 해제 후 메뉴 복원 및 문서 불변');
    });
    await test(width + ' 전체 글꼴·크기 목록과 명시 입력·Undo', size, async p => {
      await mobile(p, true); const before = await snapshot(p);
      await openCellMenu(p); await p.locator('.mini-font-picker button').click();
      eq(await p.evaluate(() => document.activeElement.matches('input,textarea,select')), false, '목록 열 때 입력 초점 없음');
      await inside(p, menu(p), '글꼴 목록'); const search = p.locator('.font-search'); ok((await metrics(search)).font >= 16, '글꼴 검색 원래 글꼴');
      for (const font of ['맑은 고딕', '굴림', '돋움', '바탕', '궁서', 'Arial', 'Calibri', 'Consolas', 'Times New Roman', 'Verdana', '합성 문서 글꼴']) ok(await p.locator('.font-item').filter({ hasText: new RegExp('^' + font + '$') }).count() >= 1, font + ' 선택 가능');
      await shot(p, width + '-font-list'); const verdana = p.locator('.font-item').filter({ hasText: /^Verdana$/ }).last(); await usable(p, verdana, '아래쪽 글꼴'); await verdana.click();
      eq(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).font), 'Verdana', '글꼴 실제 적용'); await p.evaluate(() => window.tabula.run('undo'));
      await openCellMenu(p); await p.locator('.mini-size-picker button').click(); await inside(p, menu(p), '크기 목록');
      eq(await p.evaluate(() => document.activeElement.matches('input,textarea,select')), false, '크기 목록 열 때 입력 초점 없음');
      for (const n of [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]) eq(await menu(p).getByRole('menuitem', { name: String(n), exact: true }).count(), 1, n + ' 크기');
      const large = menu(p).getByRole('menuitem', { name: '72', exact: true }); await usable(p, large, '마지막 크기'); await shot(p, width + '-size-list'); await large.click();
      eq(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).size), 72, '크기 실제 적용'); await p.evaluate(() => window.tabula.run('undo'));
      await openCellMenu(p); await p.locator('.mini-size-picker button').click(); const direct = menu(p).getByRole('menuitem', { name: '직접 입력…', exact: true }); await direct.scrollIntoViewIfNeeded(); await direct.click();
      const d = dialog(p, '글꼴 크기 직접 입력'); await inside(p, d, '직접 입력창'); const input = d.getByRole('textbox', { name: '글꼴 크기', exact: true });
      ok((await metrics(input)).font >= 16, '직접 입력 원래 CSS 16px 유지'); await input.fill('12.5'); await cancel(p, d);
      eq(await snapshot(p), before, '두 적용 Undo 및 직접입력 취소는 문서 불변');
    });
    await test(width + ' 셀서식 6탭·찾기 확장·입력과 닫기 접근', size, async p => {
      await mobile(p, true); const before = await snapshot(p);
      await openCellMenu(p); await menu(p).getByRole('menuitem', { name: /^셀 서식/ }).click(); const d = dialog(p, '셀 서식'); await inside(p, d, '셀 서식');
      for (const name of ['표시 형식', '맞춤', '글꼴', '테두리', '채우기', '보호']) {
        const tab = d.getByRole('tab', { name, exact: true }); await usable(p, tab, name + ' 탭'); await tab.click();
        const inputs = d.locator('.fc-page input:visible:not([type=checkbox]):not([type=radio]):not([type=range]):not([type=color]),.fc-page select:visible');
        if (await inputs.count()) { const input = inputs.first(); await usable(p, input, name + ' 입력'); ok((await metrics(input)).font >= 16, name + ' 입력 CSS16px'); }
      }
      await shot(p, width + '-format'); await cancel(p, d);
      await p.evaluate(() => window.tabula.run('replace')); const find = dialog(p, '찾기 및 바꾸기'); await inside(p, find, '찾기');
      const search = find.getByRole('textbox', { name: '찾을 내용', exact: true }); await usable(p, search, '찾기 입력'); ok((await metrics(search)).font >= 16, '찾기 입력 원래 글꼴'); await search.fill('합성');
      await find.getByRole('button', { name: '옵션 >>', exact: true }).click(); await inside(p, find, '확장 찾기');
      await usable(p, find.getByRole('checkbox', { name: '대/소문자 구분', exact: true }), '확장 옵션');
      const close = find.getByRole('button', { name: '닫기', exact: true }).last(); await usable(p, close, '찾기 닫기'); await shot(p, width + '-find'); await close.click();
      eq(await snapshot(p), before, '서식/찾기 열기·취소와 입력만으로 셀 불변');
    });
    await test(width + ' 빠른 분석 5탭·무늬 18개·그리기 팔레트 배치', size, async p => {
      await mobile(p, true); const before = await snapshot(p);
      await openCellMenu(p); await menu(p).getByRole('menuitem', { name: /^빠른 분석/ }).click();
      const quick = p.locator('.qa-pop'); await inside(p, quick, '빠른 분석');
      for (const name of ['서식', '차트', '합계', '표', '스파크라인']) {
        const tab = quick.locator('.qa-tab').filter({ hasText: new RegExp('^' + name) }); await usable(p, tab, name + ' 빠른 분석 탭'); await tab.click(); await raf(p);
        await inside(p, quick, name + ' 탭 전환 뒤 팝업'); const tiles = quick.locator('.qa-tile'); ok(await tiles.count() > 0, name + ' 동작 단추');
        for (const tile of await tiles.all()) await usable(p, tile, name + ' 빠른 분석 단추');
      }
      await shot(p, width + '-quick-analysis'); await p.keyboard.press('Escape'); eq(await quick.count(), 0, '빠른 분석 Escape 종료');
      await openCellMenu(p); await menu(p).getByRole('menuitem', { name: /^셀 서식/ }).click(); const d = dialog(p, '셀 서식');
      await d.getByRole('tab', { name: '채우기', exact: true }).click(); await d.locator('.pat-drop').click(); const pattern = p.locator('.pat-pop');
      await inside(p, pattern, '무늬 목록'); eq(await pattern.locator('.pat-cell').count(), 18, '무늬 없음 포함 18개');
      for (const tile of await pattern.locator('.pat-cell').all()) await usable(p, tile, '무늬 선택');
      await shot(p, width + '-patterns'); await p.keyboard.press('Escape'); eq(await pattern.count(), 0, '무늬 Escape 종료'); eq(await d.count(), 1, '부모 셀 서식 유지');
      await d.locator('.pat-drop').click(); await p.locator('.pat-pop .pat-cell').last().click(); eq(await p.locator('.pat-pop').count(), 0, '무늬 초안 선택 종료'); await cancel(p, d);
      await p.evaluate(() => window.tabula.run('drawingPalette')); const draw = p.locator('.drawing-palette'); await inside(p, draw, '그리기 팔레트');
      for (const name of ['선택', '펜', '형광펜', '선', '사각형', '원', '화살표', '획 지우개']) await usable(p, draw.getByRole('button', { name, exact: true }), name + ' 도구');
      const close = draw.getByRole('button', { name: '그리기 팔레트 닫기', exact: true }); await usable(p, close, '팔레트 닫기'); await shot(p, width + '-drawing'); await close.click(); eq(await draw.count(), 0);
      eq(await snapshot(p), before, '팝업 탐색·무늬 초안 취소·그리기 열기 닫기는 문서 불변');
    });
  }
  await test('390 모바일 격자 Ctrl+1·ShiftF10·한 칸 이동·서식 Undo', [390, 844], async p => {
    await mobile(p, true); await p.evaluate(() => window.tabula.selectCell(0, 0)); await p.locator('#gridView').focus();
    const before = await snapshot(p), history = await p.evaluate(() => window.tabula.wb().undoStack.length);
    eq(await p.evaluate(() => document.activeElement.id), 'gridView', '실제 모바일 격자 초점');
    await p.keyboard.press('Control+1'); const d = dialog(p, '셀 서식'); await inside(p, d, 'Ctrl+1 셀 서식');
    eq(await p.getByRole('dialog').count(), 1, '대화상자는 한 번만 생성'); await cancel(p, d);
    eq(await p.evaluate(() => document.activeElement.id), 'gridView', '취소 후 모바일 격자 초점 복귀');
    eq(await snapshot(p), before, 'Ctrl+1 취소 후 문서 불변');
    await p.keyboard.press('ArrowDown');
    eq(await p.evaluate(() => ({ ...window.tabula.active })), { r: 1, c: 0 }, '아래 키 한 번은 정확히 한 칸');
    eq(await p.evaluate(() => ({ ...window.tabula.sel })), { r1: 1, c1: 0, r2: 1, c2: 0 }, '선택 범위도 한 칸');
    eq(await p.evaluate(() => window.tabula.wb().undoStack.length), history, '이동은 Undo 항목을 만들지 않음');
    await p.keyboard.press('Shift+F10'); await inside(p, menu(p), 'Shift+F10 메뉴');
    eq(await menu(p).count(), 1, '키보드 메뉴 한 개'); await hiddenItems(p, true); await p.keyboard.press('Escape');
    eq(await menu(p).count(), 0, 'Escape로 우클릭 메뉴 종료');
    eq(await p.evaluate(() => document.activeElement.id), 'gridView', '메뉴 종료 뒤 격자 초점');
    await p.keyboard.press('ArrowUp'); eq(await p.evaluate(() => ({ ...window.tabula.active })), { r: 0, c: 0 }, '위 키도 한 번만 이동');
    await p.keyboard.press('Control+b'); eq(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).bold), true, '중복 처리 없이 굵게 한 번 적용');
    eq(await p.evaluate(() => window.tabula.wb().undoStack.length), history + 1, '서식 변경 한 번의 Undo');
    await p.keyboard.press('Control+z'); eq(await snapshot(p), before, '실제 Ctrl+Z로 원문·수식·서식 복원');
    eq(await p.evaluate(() => window.tabula.wb().undoStack.length), history, '실행 취소 이력 복원');
  });
} finally {
  await browser.close(); const result = { url, cases: results.length, passed: results.filter(r => r.ok).length, checks, pageErrors: results.flatMap(r => r.errors).length, writes: results.flatMap(r => r.writes).length, assets: [...assets], measurements, results, limits: '합성 Chromium 터치/키보드 검사. 실제 iPhone Safari 자동 확대나 OS 키보드 검증은 아님. 대표 팝업 표본이며 모든 팝업 전수 검사는 아님.' };
  await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify(result)); if (results.some(r => !r.ok)) process.exitCode = 1;
}
