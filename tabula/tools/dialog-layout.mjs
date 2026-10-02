// 합성 문서의 대표 팝업 레이아웃/키보드/확정·취소 회귀. 실제 모바일 OS·IME 검증은 아님.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5178/';
const out = process.env.WIXEL_DIALOG_OUT || 'D:/Codex/Temp/wixel-dialogs/current';
const filter = process.env.WIXEL_DIALOG_FILTER || '';
const widths = process.env.WIXEL_DIALOG_WIDTHS?.trim() ? process.env.WIXEL_DIALOG_WIDTHS.split(',').map(Number) : null;
const sizes = [[1366, 900], [1024, 700], [390, 844], [320, 640]].filter(([w]) => !widths || widths.includes(w));
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = [], measurements = [], errors = [], writes = [], reads = [];
let checks = 0;
const eq = (a, b, label) => { checks++; assert.deepEqual(a, b, label); };
const ok = (a, label) => { checks++; assert.ok(a, label); };
const run = (p, cmd) => p.evaluate(cmd => window.tabula.run(cmd), cmd);
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
async function raf(p) { await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))); }
async function closeAll(p) { for (let i = 0; i < 4; i++) await p.keyboard.press('Escape'); }
async function reset(p) {
  await closeAll(p);
  await p.evaluate(() => {
    const t = window.tabula, w = t.wb(); if (t.si !== 0) t.switchSheet(0);
    const cells = {}; [['지역', '부서', '매출'], ['서울', '영업', '100'], ['부산', '개발', '200'], ['대구', '영업', '300']].forEach((row, r) => row.forEach((raw, c) => cells[r + ',' + c] = { raw }));
    w.restore({ sheets: [{ name: '팝업 합성', cells }] }); w.undoStack = []; w.redoStack = [];
    t.gv().setZoom(100); t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
  }); await raf(p);
}
async function inside(p, loc, label) {
  await loc.waitFor(); const b = await loc.boundingBox(), v = p.viewportSize();
  ok(b && b.x >= -1 && b.y >= -1 && b.x + b.width <= v.width + 1 && b.y + b.height <= v.height + 1, label + ' 화면 범위 ' + JSON.stringify({ b, v }));
}
async function usable(p, loc, label) {
  await loc.scrollIntoViewIfNeeded(); await inside(p, loc, label);
  const hit = await loc.evaluate(e => { const r = e.getBoundingClientRect(), h = document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2); return { ok: !!h && (e === h || e.contains(h)), target: [e.tagName, e.getAttribute('aria-label')], rect: { x: r.x, y: r.y, w: r.width, h: r.height }, hit: h?.outerHTML.slice(0, 260) }; });
  ok(hit.ok, label + ' 중심이 다른 요소에 가리지 않음 ' + JSON.stringify(hit));
}
async function audit(p, d, label) {
  await inside(p, d, label); await usable(p, d.locator('.dialog-head button'), label + ' 닫기');
  await inside(p, d.locator('.dialog-head'), label + ' 제목');
  if (await d.locator('.dialog-foot').count()) {
    await inside(p, d.locator('.dialog-foot'), label + ' 하단');
    for (const button of await d.locator('.dialog-foot button:visible').all()) await usable(p, button, label + ' 하단 버튼');
  }
  const m = await d.evaluate(e => {
    const body = e.querySelector('.dialog-body'), visible = n => n.getClientRects().length && getComputedStyle(n).visibility !== 'hidden';
    const b = body.getBoundingClientRect(), controls = [...body.querySelectorAll('input,select,textarea,button')].filter(visible);
    return { title: e.getAttribute('aria-label'), body: { scroll: body.scrollWidth, client: body.clientWidth },
      offX: controls.filter(n => { const r = n.getBoundingClientRect(); return r.left < b.left - 2 || r.right > b.right + 2; }).filter(n => !n.closest('.opt-tabs,.dlg-tabs,.format-pane-tabs')).map(n => ({ tag: n.tagName, name: n.getAttribute('aria-label') || n.textContent?.slice(0, 80), width: n.getBoundingClientRect().width })),
      clippedText: [...body.querySelectorAll('label,legend,summary,.form-dialog-label')].filter(visible).filter(n => { const s = getComputedStyle(n); return ['hidden', 'clip'].includes(s.overflowX) && n.scrollWidth > n.clientWidth + 2; }).map(n => n.textContent.slice(0, 100)),
      checkRows: [...body.querySelectorAll('.form-dialog-check')].filter(visible).map(n => { const i = n.querySelector('input').getBoundingClientRect(), s = n.querySelector('span').getBoundingClientRect(); return { text: n.textContent, control: { x: i.x, y: i.y, w: i.width, h: i.height }, caption: { x: s.x, y: s.y, w: s.width, h: s.height } }; }),
      lists: [...body.querySelectorAll('select[size]')].filter(visible).map(n => ({ label: n.getAttribute('aria-label'), width: n.clientWidth })) };
  }); measurements.push({ width: p.viewportSize().width, label, ...m });
  ok(m.body.scroll <= m.body.client + 2, label + ' 본문 가로 넘침 ' + JSON.stringify(m.body));
  eq(m.offX, [], label + ' 입력 가로 범위'); eq(m.clippedText, [], label + ' 설명 잘림');
  for (const row of m.checkRows) { ok(row.control.x + row.control.w <= row.caption.x + 1, label + ' 체크박스가 설명 앞'); ok(row.control.w >= 15 && row.control.h >= 15, label + ' 체크박스 크기'); }
  const controls = d.locator('.dialog-body input:visible:not(:disabled),.dialog-body select:visible:not(:disabled),.dialog-body textarea:visible:not(:disabled)');
  if (await controls.count()) { await usable(p, controls.first(), label + ' 첫 입력'); await usable(p, controls.last(), label + ' 마지막 입력'); }
}
async function capture(p, key) { await p.screenshot({ path: path.join(out, p.viewportSize().width + '-' + key + '.png') }); }
async function fixtureObject(p, kind) {
  await p.evaluate(kind => {
    const t = window.tabula, w = t.wb(), canvas = document.createElement('canvas'); canvas.width = 140; canvas.height = 70; const ctx = canvas.getContext('2d'); ctx.fillStyle = '#2f855a'; ctx.fillRect(0, 0, 140, 70);
    w.transact(() => w.setSheetProp(0, kind === 'shape' ? 'shapes' : 'images', [{ id: 'dialog-' + kind, name: '합성 개체', kind: 'roundRect', src: canvas.toDataURL(), x: 30, y: 30, w: 180, h: 100, text: '합성 도형', fill: '#4472c4', stroke: '#2f528f' }])); t.gv().renderObjectsAll();
  }, kind);
  const before = await p.evaluate(kind => window.tabula.wb().sheets[0][kind === 'shape' ? 'shapes' : 'images'][0], kind);
  await p.locator('.obj[data-id="dialog-' + kind + '"]').first().dblclick();
  return before;
}
async function chart(p) { await p.evaluate(() => { const t = window.tabula; t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 }); t.run('chartColumn'); }); }
async function pivot(p) { await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => { w.addSheet('피벗 합성'); w.setSheetProp(1, 'pivot', { name: '팝업피벗', source: '팝업 합성', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 0, left: 0 }); }); t.switchSheet(1); t.run('pivotRefresh'); t.selectCell(1, 0); }); }
async function test(p, width, name, fn) {
  if (filter && !name.includes(filter)) return;
  const beforeErrors = errors.length;
  try { await reset(p); await fn(p); eq(errors.slice(beforeErrors), [], '페이지 오류'); results.push({ width, name, ok: true }); console.log('OK ' + width + ' ' + name); }
  catch (e) { results.push({ width, name, ok: false, error: e.message }); console.error('NG ' + width + ' ' + name + ': ' + e.stack); await capture(p, 'failure-' + results.length).catch(() => {}); }
  finally { await closeAll(p); }
}
try {
  for (const [width, height] of sizes) {
    const context = await browser.newContext({ viewport: { width, height } }), p = await context.newPage(); p.setDefaultTimeout(8000);
    p.on('pageerror', e => errors.push(e.stack || e.message)); p.on('dialog', d => d.dismiss());
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; let clip = ''; Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => clip, writeText: async x => { clip = x; } } }); });
    const origin = new URL(url).origin;
    await context.route('**/*', r => { const q = r.request(), u = new URL(q.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push({ method: q.method(), path: u.pathname }); return r.abort(); } if (u.origin !== origin || u.pathname.startsWith('/api/')) { reads.push(u.pathname); return r.abort(); } return r.continue(); });
    try {
      await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb()); await p.evaluate(() => document.fonts.ready); await raf(p);
      await test(p, width, '이동 옵션: 라디오·종류 체크·Tab 순환·확정', async p => {
        await run(p, 'gotoSpecial'); let d = dialog(p, '이동 옵션'); await audit(p, d, '이동 옵션');
        eq(await d.getByRole('checkbox', { name: '숫자', exact: true }).isEnabled(), false, '빈 셀에는 데이터 종류 비활성');
        await d.getByRole('radio', { name: '상수', exact: true }).check(); eq(await d.getByRole('checkbox', { name: '숫자', exact: true }).isEnabled(), true);
        for (const name of ['텍스트', '논리값', '오류']) await d.getByRole('checkbox', { name, exact: true }).uncheck();
        await capture(p, 'goto-special'); const cancel = d.getByRole('button', { name: '취소', exact: true }); await cancel.focus(); await p.keyboard.press('Tab');
        ok(await d.locator('.dialog-head button').evaluate(e => e === document.activeElement), '마지막 Tab이 닫기로 순환'); await p.keyboard.press('Shift+Tab'); ok(await cancel.evaluate(e => e === document.activeElement), '역방향 순환');
        await d.getByRole('button', { name: '확인', exact: true }).click(); await d.waitFor({ state: 'detached' }); eq(await p.evaluate(() => window.tabula.sel), { r1: 1, c1: 2, r2: 3, c2: 2 }, '숫자 상수만 선택');
      });
      await test(p, width, '긴 설명 폼: 시트 보호 체크 배치·취소·확인·Undo', async p => {
        await run(p, 'gotoSpecial'); const prior = dialog(p, '이동 옵션'); await prior.getByRole('radio', { name: '상수', exact: true }).check(); await prior.getByRole('button', { name: '확인', exact: true }).click(); await prior.waitFor({ state: 'detached' });
        eq(await p.locator('#toast').evaluate(e => e.classList.contains('show')), true, '이동 옵션 완료 알림 사전 조건');
        await run(p, 'protectSheet'); let d = dialog(p, '시트 보호'); await d.waitFor(); eq(await p.locator('#toast').evaluate(e => e.classList.contains('show')), false, '새 대화상자에서 이전 완료 알림 해제'); await audit(p, d, '시트 보호'); await capture(p, 'long-form');
        const last = d.getByRole('checkbox').last(), checked = await last.isChecked(); await last.locator('..').click(); eq(await last.isChecked(), !checked, '설명 행 클릭');
        await p.keyboard.press('Escape'); eq(await p.evaluate(() => !!window.tabula.wb().sheets[0].protect?.on), false, '취소 무변경');
        await run(p, 'protectSheet'); d = dialog(p, '시트 보호'); await d.getByRole('button', { name: '확인', exact: true }).click(); eq(await p.evaluate(() => !!window.tabula.wb().sheets[0].protect?.on), true); await run(p, 'undo'); eq(await p.evaluate(() => !!window.tabula.wb().sheets[0].protect?.on), false);
      });
      await test(p, width, '셀 서식: 6개 탭·맞춤 초안 취소·Escape', async p => {
        const before = await p.evaluate(() => window.tabula.wb().getCell(0, 0, 0)); await run(p, 'formatCells'); const d = dialog(p, '셀 서식');
        for (const name of ['표시 형식', '맞춤', '글꼴', '테두리', '채우기', '보호']) { const tab = d.getByRole('tab', { name, exact: true }); await usable(p, tab, name + ' 탭'); await tab.click(); await audit(p, d, '셀 서식 ' + name); }
        await d.getByRole('tab', { name: '맞춤', exact: true }).click(); await capture(p, 'cell-format');
        const wrap = d.getByRole('checkbox', { name: /텍스트 줄 바꿈|자동 줄 바꿈/ }).first(); await wrap.setChecked(!await wrap.isChecked());
        await p.keyboard.press('Escape'); await d.waitFor({ state: 'detached' }); eq(await p.evaluate(() => window.tabula.wb().getCell(0, 0, 0)), before, '서식 취소 무변경');
      });
      await test(p, width, 'WIXEL 옵션: 9개 범주·QAT·취소와 0 설정 저장', async p => {
        const before = await p.evaluate(() => localStorage.getItem('wixel.options')); await run(p, 'options'); let d = dialog(p, 'WIXEL 옵션');
        const tabs = await d.getByRole('tab').allTextContents(); eq(tabs.length, 9);
        for (const name of tabs) { const tab = d.getByRole('tab', { name, exact: true }); await usable(p, tab, name + ' 범주'); await tab.click(); await audit(p, d, '옵션 ' + name); }
        const list = d.getByRole('listbox', { name: '사용 가능한 명령', exact: true }); await list.selectOption({ index: 0 });
        eq(await d.locator('.qat-selection').first().innerText(), await list.locator('option:checked').innerText(), '잘린 목록의 전체 선택 이름');
        if (width < 660) { const boxes = await d.locator('.qat-list').evaluateAll(ns => ns.map(n => { const r = n.getBoundingClientRect(); return { y: r.y, bottom: r.bottom, width: r.width }; })); ok(boxes[1].y >= boxes[0].bottom, '모바일 QAT 목록은 세로 배치'); ok(boxes.every(x => x.width >= width - 90), '모바일 QAT 목록 폭'); }
        await capture(p, 'options-qat'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await p.evaluate(() => localStorage.getItem('wixel.options')), before);
        await run(p, 'options'); d = dialog(p, 'WIXEL 옵션'); await d.getByRole('tab', { name: '고급', exact: true }).click(); await d.getByRole('spinbutton', { name: '소수 자릿수', exact: true }).fill('0'); await d.getByRole('spinbutton', { name: '소수 자릿수', exact: true }).press('Tab'); await d.getByRole('button', { name: '확인', exact: true }).click(); eq(await p.evaluate(() => JSON.parse(localStorage.getItem('wixel.options')).decimalPlaces), 0);
      });
      await test(p, width, '피벗 옵션: 6개 범주·긴 설명·취소 무변경', async p => {
        await pivot(p); const before = await p.evaluate(() => window.tabula.wb().sheets[1].pivot); await run(p, 'pivotOptions'); const d = dialog(p, '피벗 테이블 옵션');
        const tabs = await d.locator('.opt-tabs button').allTextContents(); eq(tabs.length, 6);
        for (const name of tabs) { const tab = d.locator('.opt-tabs').getByRole('button', { name, exact: true }); await usable(p, tab, name + ' 피벗 범주'); await tab.click(); await audit(p, d, '피벗 ' + name); }
        await capture(p, 'pivot-options'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await p.evaluate(() => window.tabula.wb().sheets[1].pivot), before);
      });
      await test(p, width, '차트 종류: 검색·미리 보기·선택 취소', async p => {
        await chart(p); const before = await p.evaluate(() => window.tabula.wb().sheets[0].charts[0]); await run(p, 'chartChangeType'); const d = dialog(p, '차트 종류 변경'); await audit(p, d, '차트 종류');
        const search = d.getByRole('searchbox', { name: '차트 종류 검색' }); await search.fill('꺾은선'); await d.locator('.cg-sub:visible').first().click(); await audit(p, d, '차트 검색 결과'); await capture(p, 'chart-type'); await p.keyboard.press('Escape'); eq(await p.evaluate(() => window.tabula.wb().sheets[0].charts[0]), before);
      });
      await test(p, width, '차트 서식: 범주·요소 선택·닫기 접근', async p => {
        await chart(p); await run(p, 'chartFormat'); const d = dialog(p, '차트 서식'); const tabs = await d.getByRole('tab').allTextContents();
        for (const name of tabs) { const tab = d.getByRole('tab', { name, exact: true }); await usable(p, tab, name + ' 차트 서식 탭'); await tab.click(); await audit(p, d, '차트 서식 ' + name); }
        const picker = d.getByRole('combobox', { name: '서식을 지정할 차트 요소' }); await usable(p, picker, '차트 요소'); await picker.selectOption({ label: '차트 제목' }); await audit(p, d, '차트 제목 서식'); await capture(p, 'chart-format');
        if (width >= 1024) { const head = await d.locator('.dialog-head').boundingBox(); await p.mouse.move(head.x + 60, head.y + head.height / 2); await p.mouse.down(); await p.mouse.move(width - 2, height - 2, { steps: 5 }); await p.mouse.up(); await raf(p); await inside(p, d, '드래그한 차트 창'); await p.setViewportSize({ width: 900, height: 600 }); await raf(p); await audit(p, d, '창 축소 후 차트 서식'); await p.setViewportSize({ width, height }); await raf(p); }
        await d.locator('.dialog-head button').click(); await d.waitFor({ state: 'detached' });
      });
      await test(p, width, '도형 서식: 도형·텍스트 5개 범주·입력칸 접근·모델리스 닫기', async p => {
        await fixtureObject(p, 'shape'); const d = dialog(p, '도형 서식');
        const groups = [['도형 옵션', '도형 서식 범주', ['채우기 및 선', '효과', '크기 및 속성']], ['텍스트 옵션', '텍스트 서식 범주', ['텍스트 및 글꼴', '텍스트 상자']]];
        for (const [group, label, names] of groups) {
          const toggle = d.getByRole('tab', { name: group, exact: true }); await usable(p, toggle, group); await toggle.click();
          eq(await d.getByRole('tablist', { name: label, exact: true }).getByRole('tab').allTextContents(), names, group + ' 범주');
          for (const name of names) { const tab = d.getByRole('tab', { name, exact: true }); await usable(p, tab, name + ' 도형 탭'); await tab.click(); await audit(p, d, '도형 ' + name); }
        }
        await capture(p, 'shape-format'); await d.locator('.dialog-head button').click(); await d.waitFor({ state: 'detached' });
      });
      await test(p, width, '그림 서식: 긴 세로 내용·치수 초안·취소 무변경', async p => {
        const before = await fixtureObject(p, 'picture'); const d = dialog(p, '그림 서식'); await audit(p, d, '그림 서식');
        const rotate = d.getByLabel('회전(°)', { exact: true }); await usable(p, rotate, '그림 회전'); await rotate.fill('30'); const desc = d.getByLabel('그림 설명', { exact: true }); await usable(p, desc, '그림 설명'); await desc.fill('합성 그림 설명'); await capture(p, 'picture-format'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await p.evaluate(() => window.tabula.wb().sheets[0].images[0]), before);
      });
      await test(p, width, '찾기·바꾸기: Alt 입력/옵션/닫기와 셀 무변경', async p => {
        await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+h'); const d = dialog(p, '찾기 및 바꾸기'); await audit(p, d, '찾기 바꾸기');
        await p.keyboard.press('Alt+n'); ok(await d.getByRole('textbox', { name: '찾을 내용', exact: true }).evaluate(e => e === document.activeElement)); await d.getByRole('textbox', { name: '찾을 내용', exact: true }).fill('서울');
        await p.keyboard.press('Alt+e'); ok(await d.getByRole('textbox', { name: '바꿀 내용', exact: true }).evaluate(e => e === document.activeElement)); await p.keyboard.press('Alt+t'); await audit(p, d, '찾기 바꾸기 상세'); await capture(p, 'find-replace'); await p.keyboard.press('Alt+d'); await d.waitFor({ state: 'detached' }); eq(await p.evaluate(() => window.tabula.wb().getRaw(0, 1, 0)), '서울');
      });
      await test(p, width, '선택하여 붙여넣기: 옵션·연산·빈 셀·취소', async p => {
        await p.evaluate(() => { const t = window.tabula; t.selectRange({ r1: 1, c1: 0, r2: 2, c2: 2 }); t.run('copy'); t.selectCell(6, 0); }); await run(p, 'pasteSpecial'); const d = dialog(p, '선택하여 붙여넣기'); await audit(p, d, '선택하여 붙여넣기');
        const options = d.locator('input[name=psWhat]'); eq(await options.count(), 12); await usable(p, options.last(), '마지막 붙여넣기'); await options.last().check(); for (const check of await d.getByRole('checkbox').all()) { await usable(p, check, '붙여넣기 체크'); await check.setChecked(!await check.isChecked()); }
        await capture(p, 'paste-special'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await p.evaluate(() => window.tabula.wb().getValue(0, 6, 0)), null);
      });
    } finally { await context.close(); }
  }
  ok(results.length > 0, '실행할 검사 조건이 없음'); eq(errors, [], '전체 페이지 오류'); eq(writes, [], '원격 쓰기 요청');
} finally {
  await browser.close(); const summary = { url, simulation: 'Chromium viewport; 실제 모바일 OS 및 화면 읽기 프로그램 미검증', total: results.length, good: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok).length, checks, pageErrors: errors, remoteWrites: writes, blockedReads: reads, results };
  await writeFile(path.join(out, 'dialog-layout.json'), JSON.stringify({ ...summary, measurements }, null, 2)); console.log(JSON.stringify(summary, null, 2)); if (summary.bad || errors.length || writes.length) process.exitCode = 1;
}
