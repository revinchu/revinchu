// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
import { MAX_ROWS, MAX_COLS } from '../src/formula.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
try {
const page = await browser.newPage({ viewport: { width: 1400, height: 820 } });
page.setDefaultTimeout(10000);
const errs = []; page.on('pageerror', (e) => errs.push(e.message));
// 서버 문서의 자동 복원·저장이 키보드 검증용 통합 문서와 섞이지 않도록 합니다.
await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/'), { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
const ev = (f, a) => page.evaluate(f, a);
const reset = async () => {
  for (let i = 0; i < 6; i++) await page.keyboard.press('Escape');
  await ev(() => { document.querySelectorAll('.dialog-backdrop, .dlg-wrap, .modal, .qa-pop, .menu').forEach((d) => d.remove()); });
  await ev(() => {
    const t = window.tabula; const wb = t.wb(); t.switchSheet(0);
    const cells = {};
    [['이름', '수량', '단가'], ['A', '1', '100'], ['B', '2', '200'], ['C', '3', '300']].forEach((row, r) => row.forEach((v, c) => { cells[`${r},${c}`] = { raw: v }; }));
    wb.restore({ sheets: [{ name: '단축키 검사', cells, allStyle: { align: 'center' } }] });
    wb.undoStack = []; wb.redoStack = []; t.gv().layout(); t.gv().renderAll(); t.selectCell(1, 1);
  });
  await page.locator('#cellEditor').focus();
};
const val = (r, c) => ev(([r, c]) => window.tabula.wb().getValue(0, r, c), [r, c]);
const raw = (r, c) => ev(([r, c]) => window.tabula.wb().getCell(0, r, c)?.raw, [r, c]);
const st = (r, c) => ev(([r, c]) => window.tabula.wb().styleAt(0, r, c), [r, c]);
const sel = () => ev(() => ({ ...window.tabula.sel, a: window.tabula.active }));
const dlg = () => ev(() => document.querySelector('.dialog .dialog-title, .dialog h2, .dlg-title')?.textContent ?? document.querySelector('.dialog')?.textContent?.slice(0, 20) ?? null);
const results = [];
const T = async (name, fn) => {
  let result; const startErrors = errs.length;
  try { await reset(); result = await fn(); if (errs.length > startErrors) result = { pageErrors: errs.slice(startErrors) }; }
  catch (e) { result = '오류 ' + e.message.slice(0, 500); }
  const ok = result === true; results.push({ name, ok, detail: result });
  console.log(`${ok ? 'OK' : 'NG'} ${name}${ok ? '' : ` → ${JSON.stringify(result)}`}`);
};
const k = (x) => page.keyboard.press(x);
await T('Ctrl+; 오늘 날짜', async () => { const expected = await ev(() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }); await k('Control+;'); assert.equal(await raw(1, 1), expected); return true; });
await T('Ctrl+Shift+; 현재 시간', async () => { const minutes = () => ev(() => { const d = new Date(); return d.getHours() * 60 + d.getMinutes(); }); const before = await minutes(); await k('Control+Shift+;'); const after = await minutes(); const v = await val(1, 1); assert.equal(typeof v, 'number'); assert.ok([before, after].includes(Math.round(v * 1440))); assert.notEqual(await raw(1, 1), '1'); return true; });
await T('Ctrl+D 아래로 채우기', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 1 }, 'cells', { r: 1, c: 1 })); await k('Control+d'); return (await val(3, 1)) === 1 || await val(3, 1); });
await T('Ctrl+R 오른쪽 채우기', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 1, c2: 3 }, 'cells', { r: 1, c: 1 })); await k('Control+r'); return (await val(1, 3)) === 1 || await val(1, 3); });
await T('Ctrl+Enter 여러 셀 동시 입력', async () => { await ev(() => window.tabula.selectRange({ r1: 5, c1: 0, r2: 7, c2: 0 }, 'cells', { r: 5, c: 0 })); await page.keyboard.type('x'); await k('Control+Enter'); return (await val(7, 0)) === 'x' || await val(7, 0); });
await T('Alt+Enter 셀 안 줄바꿈', async () => { await page.keyboard.type('가'); await k('Alt+Enter'); await page.keyboard.type('나'); await k('Enter'); return (await raw(1, 1)) === '가\n나' || await raw(1, 1); });
await T('Ctrl+→ 끝으로 이동', async () => { await ev(() => window.tabula.selectCell(1, 0)); await k('Control+ArrowRight'); return (await sel()).a.c === 2 || await sel(); });
await T('Ctrl+Shift+↓ 끝까지 선택', async () => { await k('Control+Shift+ArrowDown'); return (await sel()).r2 === 3 || await sel(); });
await T('Alt+= 자동 합계', async () => { await ev(() => window.tabula.selectCell(4, 1)); await k('Alt+='); await k('Enter'); return (await val(4, 1)) === 6 || [await raw(4, 1), await val(4, 1)]; });
await T('Ctrl+Shift+1 숫자 서식 (#,##0)', async () => { await k('Control+Shift+!'); const s = await st(1, 1); return (s.numFmt === 'comma' && s.decimals === 0) || s; });
await T('Ctrl+Shift+4 통화', async () => { await k('Control+Shift+$'); return /currency|accounting/.test((await st(1, 1)).numFmt) || (await st(1, 1)).numFmt; });
await T('Ctrl+Shift+5 백분율', async () => { await k('Control+Shift+%'); return (await st(1, 1)).numFmt === 'percent' || (await st(1, 1)).numFmt; });
await T('Ctrl+Shift+3 날짜', async () => { await k('Control+Shift+#'); return /date/.test((await st(1, 1)).numFmt) || (await st(1, 1)).numFmt; });
await T('Ctrl+Shift+~ 일반', async () => { await k('Control+Shift+%'); await k('Control+Shift+~'); const f = (await st(1, 1)).numFmt; return !f || f === 'general' || f; });
await T('Ctrl+B 굵게', async () => { await k('Control+b'); return (await st(1, 1)).bold === true; });
await T('Ctrl+5 취소선', async () => { await k('Control+5'); return (await st(1, 1)).strike === true; });
await T('Ctrl+Shift+& 바깥 테두리', async () => { await k('Control+Shift+&'); const s = await st(1, 1); return !!(s.bt && s.bb && s.bl && s.br) || s; });
await T('Ctrl+Shift+_ 테두리 제거', async () => { await k('Control+Shift+&'); await k('Control+Shift+_'); const s = await st(1, 1); return !s.bt && !s.bb || s; });
await T('Ctrl+9 행 숨기기', async () => { await k('Control+9'); return ev(() => !!window.tabula.wb().sheets[0].hiddenRows?.[1]); });
await T('Ctrl+Shift+9 행 숨기기 취소', async () => { await ev(() => window.tabula.selectRange({ r1: 0, c1: 0, r2: 2, c2: 0 }, 'cells', { r: 0, c: 0 })); await k('Control+9'); await k('Control+Shift+('); return ev(() => !Object.keys(window.tabula.wb().sheets[0].hiddenRows ?? {}).length); });
await T('Ctrl+Space 활성 열만 전체 선택', async () => { await k('Control+Space'); assert.deepEqual(await sel(), { r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 1, a: { r: 1, c: 1 } }); return true; });
await T('Shift+Space 활성 행만 전체 선택', async () => { await k('Shift+Space'); assert.deepEqual(await sel(), { r1: 1, c1: 0, r2: 1, c2: MAX_COLS - 1, a: { r: 1, c: 1 } }); return true; });
await T('Ctrl+Shift+= 삽입', async () => { await k('Shift+Space'); await k('Control+Shift+='); return (await val(2, 1)) === 1 || await val(2, 1); });
await T('Ctrl+- 삭제', async () => { await k('Shift+Space'); await k('Control+-'); return (await val(1, 1)) === 2 || await val(1, 1); });
await T('Shift+F11 새 시트', async () => { await k('Shift+F11'); return ev(() => window.tabula.wb().sheets.length === 2); });
await T('Ctrl+PageDown 다음 시트', async () => { await ev(() => window.tabula.run('addSheet')); await ev(() => window.tabula.switchSheet(0)); await k('Control+PageDown'); return ev(() => window.tabula.si === 1); });
await T('Ctrl+` 수식 표시 켜기/끄기', async () => { await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 1, 1, '=1+1')); }); const cell = page.locator('.c[data-r="1"][data-c="1"]').first(); assert.equal(await cell.innerText(), '2'); await k('Control+`'); assert.equal(await cell.innerText(), '=1+1'); await k('Control+`'); assert.equal(await cell.innerText(), '2'); return true; });
await T('F4 참조 전환 (편집 중)', async () => { await ev(() => window.tabula.selectCell(6, 3)); await page.keyboard.type('=A1'); await k('F4'); await k('Enter'); return (await raw(6, 3)) === '=$A$1' || await raw(6, 3); });
await T("Ctrl+' 위 셀 수식", async () => { await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(0, 6, 3, '=B2*2')); window.tabula.selectCell(7, 3); }); await k("Control+'"); await k('Enter'); return (await raw(7, 3)) === '=B2*2' || await raw(7, 3); });
await T('Ctrl+Shift+L 필터', async () => { await k('Control+Shift+l'); return ev(() => !!window.tabula.wb().sheets[0].filter); });
await T('Ctrl+T 표 만들기', async () => { await k('Control+t'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /표/.test(d) || d.slice(0, 40); });
await T('Ctrl+1 셀 서식', async () => { await k('Control+1'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /셀 서식/.test(d) || d.slice(0, 40); });
await T('Ctrl+K 하이퍼링크', async () => { await k('Control+k'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /하이퍼링크/.test(d) || d.slice(0, 40); });
await T('F5 이동', async () => { await k('F5'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이동/.test(d) || d.slice(0, 40); });
await T('Ctrl+F 찾기', async () => { await k('Control+f'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /찾기/.test(d) || d.slice(0, 40); });
await T('Ctrl+H 바꾸기', async () => { await k('Control+h'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /바꾸기/.test(d) || d.slice(0, 40); });
await T('Ctrl+Alt+V 선택하여 붙여넣기', async () => { await k('Control+c'); await k('Control+Alt+v'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /붙여넣기/.test(d) || d.slice(0, 40); });
await T('Ctrl+Shift+V 값 붙여넣기', async () => { await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(0, 6, 5, '=1+1')); window.tabula.selectCell(6, 5); }); await k('Control+c'); await ev(() => window.tabula.selectCell(7, 5)); await k('Control+Shift+v'); return (await raw(7, 5)) === '2' || await raw(7, 5); });
await T('Ctrl+E 빠른 채우기', async () => { await ev(() => { const wb = window.tabula.wb(); wb.transact(() => { wb.setInput(0, 1, 4, 'A-1'); }); window.tabula.selectCell(2, 4); }); await k('Control+e'); return (await val(2, 4)) === 'B-2' || await val(2, 4); });
await T('Alt+Shift+→ 그룹', async () => { await k('Shift+Space'); await k('Alt+Shift+ArrowRight'); return ev(() => !!window.tabula.wb().sheets[0].outline); });
await T('Ctrl+[ 참조 셀로', async () => { await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(0, 8, 0, '=C3')); window.tabula.selectCell(8, 0); }); await k('Control+['); const s = await sel(); return (s.a.r === 2 && s.a.c === 2) || s.a; });
await T('Ctrl+Shift+O 메모 셀', async () => {
  await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setComment(0, 3, 2, '검증 메모')); });
  await k('Control+Shift+o');
  const s = await sel();
  return (s.r1 === 3 && s.r2 === 3 && s.c1 === 2 && s.c2 === 2 && s.a.r === 3 && s.a.c === 2) || s;
});
await T('Shift+F3 함수 삽입', async () => { await k('Shift+F3'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /함수/.test(d) || d.slice(0, 40); });
await T('Alt+F1 차트', async () => { await ev(() => window.tabula.selectCell(1, 1)); await k('Alt+F1'); return ev(() => window.tabula.wb().sheets[0].charts.length > 0); });
await T('F11 차트 시트', async () => { await k('F11'); await page.waitForTimeout(150); return ev(() => window.tabula.wb().sheets.some((x) => /^Chart/.test(x.name) && x.charts?.length > 0) && window.tabula.wb().sheets[window.tabula.si].name.startsWith('Chart')); });
await T('Ctrl+Z / Ctrl+Y', async () => { await page.keyboard.type('zz'); await k('Enter'); await k('Control+z'); const a = await val(1, 1); await k('Control+y'); const b = await val(1, 1); return (a === 1 && b === 'zz') || [a, b]; });
await T('Ctrl+Alt+F9 파일 계산 캐시 대신 실제 전체 계산', async () => { await ev(() => { const t = window.tabula; t.wb().restore({ sheets: [{ name: '계산', fileValues: true, cells: { '1,1': { raw: '=1+1', cached: 42 } } }] }); t.selectCell(1, 1); }); assert.equal(await val(1, 1), 42); await k('Control+Alt+F9'); assert.equal(await val(1, 1), 2); return true; });
await T('Ctrl+Shift+U 수식 입력줄', async () => { const a = await ev(() => document.getElementById('formulaRow').className); await k('Control+Shift+u'); const b = await ev(() => document.getElementById('formulaRow').className); await k('Control+Shift+u'); return a !== b || [a, b]; });
await T('Ctrl+Shift+Enter 배열', async () => { await ev(() => window.tabula.selectCell(9, 5)); await page.keyboard.type('=SUM(B2:B4*C2:C4)'); await k('Control+Shift+Enter'); return (await val(9, 5)) === 1400 || await val(9, 5); });
await T('Ctrl+A 현재 영역 → 전체 행/열', async () => { await k('Control+a'); const a = await sel(); assert.deepEqual([a.r1, a.c1, a.r2, a.c2], [0, 0, 3, 2]); await k('Control+a'); const b = await sel(); assert.deepEqual([b.r1, b.c1, b.r2, b.c2], [0, 0, MAX_ROWS - 1, MAX_COLS - 1]); return true; });
await T('Ctrl+Home', async () => { await k('Control+Home'); const s = await sel(); return (s.a.r === 0 && s.a.c === 0) || s.a; });
await T('Ctrl+End', async () => { await k('Control+End'); const s = await sel(); return (s.a.r === 3 && s.a.c === 2) || s.a; });
await T('Ctrl+Backspace 스크롤을 활성 셀로 복원', async () => { await page.locator('#gridScroll').evaluate((s) => { s.scrollTop = 700; }); await page.waitForFunction(() => document.getElementById('gridScroll').scrollTop > 100 && window.tabula.gv().sy > 100); await k('Control+Backspace'); await page.waitForFunction(() => document.getElementById('gridScroll').scrollTop < 50 && window.tabula.gv().sy < 50); assert.deepEqual((await sel()).a, { r: 1, c: 1 }); return true; });
await T('F2 편집', async () => { await k('F2'); const e = await ev(() => !document.getElementById('cellEditor').classList.contains('idle')); await k('Escape'); return e; });
await T('Shift+F2 메모', async () => { await k('Shift+F2'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? document.querySelector('.cm-edit, .comment-edit')?.textContent ?? ''); await k('Escape'); return d.length > 0 || d; });
await T('Ctrl+Q 빠른 분석', async () => { await k('Control+q'); const d = await ev(() => !!document.querySelector('.qa-pop')); await k('Escape'); return d; });
await T('Alt+; 보이는 행만 선택 후 Delete', async () => { await ev(() => { const t = window.tabula; t.wb().transact(() => t.wb().setSheetProp(0, 'hiddenRows', { 2: true })); t.gv().layout(); t.selectRange({ r1: 1, c1: 1, r2: 3, c2: 1 }); }); await k('Alt+;'); await k('Delete'); assert.deepEqual([await val(1, 1), await val(2, 1), await val(3, 1)], [null, 2, null]); return true; });
await T('Ctrl+\\ 각 행의 다른 값만 선택 후 Delete', async () => { await ev(() => { const t = window.tabula, w = t.wb(); w.transact(() => [[1, 1, 9], [2, 3, 2], [3, 3, 5]].forEach((row, r) => row.forEach((v, c) => w.setInput(0, r + 1, c + 1, String(v))))); t.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 }, 'cells', { r: 1, c: 1 }); }); await k('Control+\\'); await k('Delete'); const values = await ev(() => [1, 2, 3].map((r) => [1, 2, 3].map((c) => window.tabula.wb().getValue(0, r, c)))); assert.deepEqual(values, [[1, 1, null], [2, null, 2], [3, 3, null]]); return true; });
await T('Alt+↓ 목록 선택', async () => {
  await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setSheetProp(0, 'validations', [{ r1: 1, c1: 1, r2: 1, c2: 1, type: 'list', f1: '"사과,배"' }])); });
  try {
    await k('Alt+ArrowDown');
    const items = await page.locator('.dv-list .menu-item').allTextContents();
    if (items.length !== 2 || items[0].trim() !== '사과' || items[1].trim() !== '배') return items;
    await page.locator('.dv-list .menu-item').nth(1).click();
    return (await val(1, 1)) === '배' || await val(1, 1);
  } finally {
    await k('Escape');
    await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setSheetProp(0, 'validations', [])); });
  }
});
await T('Ctrl+F1 리본 접기', async () => { await k('Control+F1'); const c = await ev(() => document.getElementById('ribbon').classList.contains('collapsed')); await k('Control+F1'); return c; });
await T('Alt → H 키팁 실제 표시·단계 전환·닫기', async () => { assert.equal(await ev(() => document.body.classList.contains('keytips')), false); await k('Alt'); assert.equal(await ev(() => document.body.classList.contains('keytips')), true); assert.equal(await page.locator('.keytip-panel').count(), 0); await k('h'); assert.equal(await ev(() => document.body.dataset.keytipSequence), 'h'); const borderBadge = page.locator('.keytip-badge[data-keytip-path="hb"]'); assert.equal(await borderBadge.isVisible(), true); assert.equal(await borderBadge.innerText(), 'B'); await k('Alt'); assert.equal(await ev(() => document.body.classList.contains('keytips')), false); assert.equal(await page.locator('.keytip-badge').count(), 0); return true; });
await T('F12 다른 이름으로 저장', async () => { await k('F12'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /다른 이름/.test(d) || d.slice(0, 40); });
await T('Ctrl+Shift+F3 이름 만들기', async () => { await k('Control+Shift+F3'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이름/.test(d) || d.slice(0, 40); });
await T('Ctrl+F3 이름 관리자', async () => { await k('Control+F3'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이름/.test(d) || d.slice(0, 40); });
await T('F7 맞춤법 검사', async () => { await k('F7'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /맞춤법/.test(d) || d.slice(0, 40); });
await T('Ctrl+Shift+A 인수 이름', async () => { await page.keyboard.type('=SUM'); await k('Control+Shift+a'); const t = await ev(() => document.getElementById('cellEditor').value); await k('Escape'); return /SUM\(number1/.test(t) || t; });
// 메뉴가 열렸다는 사실뿐 아니라 선택·값·서식의 정확한 변화를 검증합니다.
await T('행 전체 선택 → Ctrl+Space 활성 열만', async () => { await k('Shift+Space'); await k('Control+Space'); assert.deepEqual(await sel(), { r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 1, a: { r: 1, c: 1 } }); return true; });
await T('여러 셀 → Ctrl+Space 해당 열들만', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 }, 'cells', { r: 2, c: 2 })); await k('Control+Space'); assert.deepEqual(await sel(), { r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 3, a: { r: 2, c: 2 } }); return true; });
await T('여러 셀 → Shift+Space 해당 행들만', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 }, 'cells', { r: 2, c: 2 })); await k('Shift+Space'); assert.deepEqual(await sel(), { r1: 1, c1: 0, r2: 3, c2: MAX_COLS - 1, a: { r: 2, c: 2 } }); return true; });
await T('Ctrl+Shift+Space 전체 → Ctrl+Space 활성 열', async () => { await k('Control+Shift+Space'); const s = await sel(); assert.deepEqual([s.r1, s.c1, s.r2, s.c2], [0, 0, MAX_ROWS - 1, MAX_COLS - 1]); await k('Control+Space'); assert.deepEqual(await sel(), { r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 1, a: { r: 1, c: 1 } }); return true; });
for (const [key, prop] of [['Control+i', 'italic'], ['Control+u', 'underline'], ['Control+2', 'bold'], ['Control+3', 'italic'], ['Control+4', 'underline']]) {
  await T(`${key} 서식 적용·해제`, async () => { await k(key); assert.equal((await st(1, 1))[prop], true); await k(key); assert.equal(!!(await st(1, 1))[prop], false); return true; });
}
await T('Ctrl+Shift+2 시간 서식', async () => { await k('Control+Shift+@'); assert.equal((await st(1, 1)).numFmt, 'time'); return true; });
await T('Ctrl+Shift+6 지수 서식', async () => { await k('Control+Shift+^'); assert.equal((await st(1, 1)).numFmt, 'scientific'); return true; });
await T('Ctrl+Shift+8 현재 영역 선택', async () => { await k('Control+Shift+*'); const s = await sel(); assert.deepEqual([s.r1, s.c1, s.r2, s.c2], [0, 0, 3, 2]); return true; });
await T('Ctrl+0 / Ctrl+Shift+0 열 숨기기·취소', async () => { await k('Control+0'); assert.equal(await ev(() => !!window.tabula.wb().sheets[0].hiddenCols?.[1]), true); await ev(() => window.tabula.selectRange({ r1: 0, c1: 0, r2: 0, c2: 2 })); await k('Control+Shift+)'); assert.equal(await ev(() => Object.keys(window.tabula.wb().sheets[0].hiddenCols ?? {}).length), 0); return true; });
await T('Tab / Shift+Tab / Enter / Shift+Enter 이동', async () => { for (const [key, expected] of [['Tab', { r: 1, c: 2 }], ['Shift+Tab', { r: 1, c: 1 }], ['Enter', { r: 2, c: 1 }], ['Shift+Enter', { r: 1, c: 1 }]]) { await k(key); assert.deepEqual((await sel()).a, expected); } return true; });
await T('Shift+방향키 확장과 Shift+Backspace 축소', async () => { await k('Shift+ArrowRight'); await k('Shift+ArrowDown'); assert.deepEqual(await sel(), { r1: 1, c1: 1, r2: 2, c2: 2, a: { r: 1, c: 1 } }); await k('Shift+Backspace'); assert.deepEqual(await sel(), { r1: 1, c1: 1, r2: 1, c2: 1, a: { r: 1, c: 1 } }); return true; });
await T('F8 선택 확장 → Escape 해제', async () => { await k('F8'); await k('ArrowRight'); assert.deepEqual(await sel(), { r1: 1, c1: 1, r2: 1, c2: 2, a: { r: 1, c: 1 } }); await k('Escape'); await k('ArrowDown'); const s = await sel(); assert.equal(s.r1, s.r2); assert.equal(s.c1, s.c2); return true; });
await T('Ctrl+Shift+Home / End 활성 셀을 보존하며 확장', async () => { await k('Control+Shift+Home'); assert.deepEqual(await sel(), { r1: 0, c1: 0, r2: 1, c2: 1, a: { r: 1, c: 1 } }); await ev(() => window.tabula.selectCell(1, 1)); await k('Control+Shift+End'); assert.deepEqual(await sel(), { r1: 1, c1: 1, r2: 3, c2: 2, a: { r: 1, c: 1 } }); return true; });
await T('Ctrl+PageUp 이전 시트', async () => { await ev(() => window.tabula.run('addSheet')); assert.equal(await ev(() => window.tabula.si), 1); await k('Control+PageUp'); assert.equal(await ev(() => window.tabula.si), 0); return true; });
await T('Ctrl+. 선택 영역 모서리 네 번 순환', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 }, 'cells', { r: 1, c: 1 })); for (const [r, c] of [[1, 3], [3, 3], [3, 1], [1, 1]]) { await k('Control+.'); assert.deepEqual(await sel(), { r1: 1, c1: 1, r2: 3, c2: 3, a: { r, c } }); } return true; });
await T('Ctrl+Shift+따옴표 위 셀의 계산 결과만 입력', async () => { await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 0, 1, '=2+3')); }); await k('Control+Shift+"'); await k('Enter'); assert.equal(await raw(1, 1), '5'); assert.equal(await val(1, 1), 5); return true; });
await T('Ctrl+] 종속 셀 선택', async () => { await ev(() => { const t = window.tabula; t.wb().transact(() => t.wb().setInput(0, 6, 3, '=B2*2')); }); await k('Control+]'); assert.deepEqual((await sel()).a, { r: 6, c: 3 }); return true; });
await T('Ctrl+Shift+\\ 각 열의 다른 값만 삭제', async () => { await ev(() => { const t = window.tabula, w = t.wb(); w.transact(() => [[1, 1, 9], [2, 3, 9], [1, 3, 5]].forEach((row, r) => row.forEach((v, c) => w.setInput(0, r + 1, c + 1, String(v))))); t.selectRange({ r1: 1, c1: 1, r2: 3, c2: 3 }, 'cells', { r: 1, c: 1 }); }); await k('Control+Shift+|'); await k('Delete'); assert.deepEqual(await ev(() => [1, 2, 3].map((r) => [1, 2, 3].map((c) => window.tabula.wb().getValue(0, r, c)))), [[1, 1, 9], [null, null, 9], [1, null, null]]); return true; });
await T('Alt+; 숨긴 열을 제외한 셀만 삭제', async () => { await ev(() => { const t = window.tabula; t.wb().transact(() => t.wb().setSheetProp(0, 'hiddenCols', { 1: true })); t.gv().layout(); t.selectRange({ r1: 1, c1: 0, r2: 1, c2: 2 }); }); await k('Alt+;'); await k('Delete'); assert.deepEqual([await val(1, 0), await val(1, 1), await val(1, 2)], [null, 1, null]); return true; });
await T('Alt+Shift+왼쪽 그룹 해제', async () => { await k('Shift+Space'); await k('Alt+Shift+ArrowRight'); assert.equal(await ev(() => window.tabula.wb().sheets[0].outline?.rows?.[1]), 1); await k('Alt+Shift+ArrowLeft'); assert.equal(await ev(() => !!window.tabula.wb().sheets[0].outline?.rows?.[1]), false); return true; });
await T('Ctrl+F8 값 강조 켜기·끄기', async () => { await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 1, 2, '=1+1')); }); const colors = () => ev(() => [0, 1, 2].map((c) => getComputedStyle(document.querySelector(`.c[data-r="1"][data-c="${c}"]`)).color)); const before = await colors(); await k('Control+F8'); assert.deepEqual(await colors(), ['rgb(0, 0, 0)', 'rgb(0, 0, 255)', 'rgb(0, 128, 0)']); await k('Control+F8'); assert.deepEqual(await colors(), before); return true; });
await T('Shift+F9 현재 시트만 파일 캐시 재계산', async () => { await ev(() => { const t = window.tabula; t.wb().restore({ sheets: [{ name: '현재', fileValues: true, cells: { '1,1': { raw: '=1+1', cached: 42 } } }, { name: '다음', fileValues: true, cells: { '1,1': { raw: '=2+2', cached: 84 } } }] }); t.selectCell(1, 1); }); assert.equal(await val(1, 1), 42); await k('Shift+F9'); assert.equal(await val(1, 1), 2); assert.equal(await ev(() => window.tabula.wb().getValue(1, 1, 1)), 84); return true; });
await T('F4 / Ctrl+Y 마지막 서식 반복', async () => { await k('Control+b'); await ev(() => window.tabula.selectCell(2, 1)); await k('F4'); assert.equal((await st(2, 1)).bold, true); await ev(() => window.tabula.selectCell(3, 1)); await k('Control+y'); assert.equal((await st(3, 1)).bold, true); return true; });
await T('F4 편집 중 상대·절대 참조 네 단계', async () => { await page.keyboard.type('=A1'); for (const expected of ['=$A$1', '=A$1', '=$A1', '=A1']) { await k('F4'); assert.equal(await page.locator('#cellEditor').inputValue(), expected); } await k('Escape'); assert.equal(await val(1, 1), 1); return true; });
await T('한글 자판 Ctrl+B/I/U/A 물리 키 코드 (합성 이벤트)', async () => { for (const [key, code, prop] of [['ㅠ', 'KeyB', 'bold'], ['ㅑ', 'KeyI', 'italic'], ['ㅕ', 'KeyU', 'underline']]) { await page.locator('#cellEditor').dispatchEvent('keydown', { key, code, ctrlKey: true, bubbles: true }); assert.equal((await st(1, 1))[prop], true); } await page.locator('#cellEditor').dispatchEvent('keydown', { key: 'ㅁ', code: 'KeyA', ctrlKey: true, bubbles: true }); const s = await sel(); assert.deepEqual([s.r1, s.c1, s.r2, s.c2], [0, 0, 3, 2]); return true; });
await T('등록되지 않은 Ctrl+Alt+B/I/U는 서식을 바꾸지 않음', async () => { const before = await st(1, 1); for (const key of ['Control+Alt+b', 'Control+Alt+i', 'Control+Alt+u']) await k(key); assert.deepEqual(await st(1, 1), before); return true; });
await T('Shift+F10 우클릭 메뉴 키보드 열기·닫기', async () => { await k('Shift+F10'); await page.getByRole('menu').first().waitFor(); assert.ok(await page.getByRole('menuitem').count() > 0); await k('Escape'); assert.equal(await page.getByRole('menu').count(), 0); return true; });
await T('Ctrl+C / Ctrl+V 수식 상대 참조와 서식 복사', async () => {
  await ev(() => { const t = window.tabula, w = t.wb(); w.transact(() => { w.setInput(0, 1, 1, '=C2*2'); w.setStyle(0, 1, 1, { bold: true }); }); });
  await k('Control+c'); await ev(() => window.tabula.selectCell(2, 1)); await k('Control+v');
  await page.waitForFunction(() => window.tabula.wb().getRaw(0, 2, 1) === '=C3*2');
  assert.equal(await val(2, 1), 400); assert.equal((await st(2, 1)).bold, true); assert.equal(await raw(1, 1), '=C2*2'); return true;
});
await T('Ctrl+X / Ctrl+V 잘라내기와 Ctrl+Z 복원', async () => {
  await k('Control+x'); await ev(() => window.tabula.selectCell(5, 4)); await k('Control+v');
  await page.waitForFunction(() => window.tabula.wb().getValue(0, 5, 4) === 1);
  assert.equal(await val(1, 1), null); await k('Control+z'); assert.equal(await val(1, 1), 1); assert.equal(await val(5, 4), null); return true;
});
await T('Ctrl+T 실제 표 생성 / Ctrl+Shift+T 합계 행 켜기·끄기', async () => {
  await k('Control+t'); await page.getByRole('dialog', { name: '표 만들기', exact: true }).getByRole('button', { name: '확인', exact: true }).click();
  assert.equal(await ev(() => window.tabula.wb().sheets[0].tables.length), 1);
  await page.locator('#cellEditor').focus(); await k('Control+Shift+t'); assert.equal(await ev(() => window.tabula.wb().sheets[0].tables[0].totals), true);
  await k('Control+Shift+t'); assert.equal(await ev(() => !!window.tabula.wb().sheets[0].tables[0].totals), false); return true;
});
await T('빠른 실행 기본 12개 순서·번호·리본 아래 위치', async () => {
  assert.deepEqual(await page.locator('#quickAccess [data-qat-cmd]').evaluateAll((buttons) => buttons.map((b) => [b.dataset.qatCmd, b.dataset.qatKey])), [
    ['painter', '1'], ['mergeCenter', '2'], ['alignCenter', '3'], ['incDecimal', '4'], ['autosum', '5'], ['calcField', '6'],
    ['toggleGrid', '7'], ['condColorScale', '8'], ['condDataBar', '9'], ['refreshAll', '09'], ['textToColumns', '08'], ['replace', '07'],
  ]);
  assert.match(await page.locator('#quickAccess').getAttribute('class'), /below/); return true;
});
await T('기본 Alt+1 서식 복사 후 대상 셀 클릭', async () => {
  await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setStyle(0, 1, 1, { bold: true, fill: '#ff0000' })); });
  await k('Alt+1'); await page.locator('.c[data-r="2"][data-c="1"]').first().click();
  assert.equal((await st(2, 1)).bold, true); assert.equal((await st(2, 1)).fill, '#ff0000'); return true;
});
await T('기본 Alt+2 병합하고 가운데 맞춤', async () => {
  await ev(() => window.tabula.selectRange({ r1: 5, c1: 4, r2: 5, c2: 5 })); await k('Alt+2');
  assert.equal(await ev(() => window.tabula.wb().sheets[0].merges.some((m) => m.r1 === 5 && m.r2 === 5 && m.c1 === 4 && m.c2 === 5)), true);
  assert.equal((await st(5, 4)).align, 'center'); return true;
});
await T('기본 Alt+3 가운데 맞춤', async () => { await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setStyle(0, 1, 1, { align: 'right' })); }); await k('Alt+3'); assert.equal((await st(1, 1)).align, 'center'); return true; });
await T('기본 Alt+4 소수 자릿수 늘리기', async () => { await k('Alt+4'); assert.equal((await st(1, 1)).decimals, 1); return true; });
await T('기본 Alt+5 자동 합계', async () => { await ev(() => window.tabula.selectCell(4, 1)); await k('Alt+5'); await k('Enter'); assert.equal(await val(4, 1), 6); return true; });
await T('기본 Alt+7 눈금선 켜기·끄기', async () => { await k('Alt+7'); assert.equal(await ev(() => window.tabula.wb().sheets[0].noGrid), true); await k('Alt+7'); assert.equal(await ev(() => !!window.tabula.wb().sheets[0].noGrid), false); return true; });
for (const [key, type, name] of [['8', 'scale', '색조'], ['9', 'bar', '데이터 막대']]) {
  await T(`기본 Alt+${key} 조건부 서식 ${name} 견본 적용`, async () => {
    await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 1 })); await k(`Alt+${key}`);
    await page.locator('.menu .cf-swatch').first().click();
    assert.deepEqual(await ev(() => { const r = window.tabula.wb().sheets[0].cond[0]; return [r.type, r.r1, r.c1, r.r2, r.c2]; }), [type, 1, 1, 3, 1]); return true;
  });
}
await T('기본 Alt→0→9 피벗 새로 고침 / Alt+6 계산 필드', async () => {
  await ev(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'pivot', { name: '키 검사 피벗', source: w.sheets[0].name, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, rows: ['이름'], cols: [], values: [{ field: '수량', agg: 'sum' }], top: 5, left: 4 })); t.selectCell(5, 4); t.run('refreshAll'); });
  assert.equal(await val(6, 5), 1);
  await ev(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 1, 1, '100')); });
  assert.equal(await val(6, 5), 1); await page.locator('#cellEditor').focus(); await k('Alt'); await k('0'); await k('9'); assert.equal(await val(6, 5), 100);
  await k('Alt+6'); assert.match(await page.locator('.dialog').innerText(), /계산 필드/); await k('Escape'); return true;
});
await T('기본 Alt→0→8 텍스트 나누기 / Alt→0→7 바꾸기', async () => {
  await k('Alt'); await k('0'); await k('8'); assert.equal(await page.getByRole('dialog', { name: '텍스트 마법사', exact: true }).isVisible(), true); await k('Escape');
  await page.locator('#cellEditor').focus(); await k('Alt'); await k('0'); await k('7'); assert.match(await page.locator('.dialog').innerText(), /바꾸기/); await k('Escape'); return true;
});
await T('빠른 실행 명령 순서·리본 아래·Alt+1·설정 복원', async () => {
  await ev(() => window.tabula.run('options'));
  await page.getByRole('tab', { name: '빠른 실행 도구 모음', exact: true }).click();
  await page.getByLabel('표시 위치', { exact: true }).selectOption('below');
  await page.getByLabel('사용 가능한 명령').selectOption('bold'); await page.getByRole('button', { name: '추가(A) >>', exact: true }).click();
  const moveCount = await page.getByLabel('현재 도구 모음 순서').evaluate((s) => s.selectedIndex);
  for (let i = 0; i < moveCount; i++) await page.getByRole('button', { name: '위로', exact: true }).click();
  await page.getByRole('button', { name: '확인', exact: true }).click();
  assert.match(await page.locator('#quickAccess').getAttribute('class'), /below/);
  assert.equal(await page.locator('[data-qat-key="1"]').getAttribute('data-qat-cmd'), 'bold');
  await page.locator('#cellEditor').focus(); await k('Alt+1'); assert.equal((await st(1, 1)).bold, true);
  await page.reload({ waitUntil: 'domcontentloaded' }); await page.waitForFunction(() => !!window.tabula?.wb(), null, { timeout: 60000 });
  assert.match(await page.locator('#quickAccess').getAttribute('class'), /below/);
  assert.equal(await page.locator('[data-qat-key="1"]').getAttribute('data-qat-cmd'), 'bold');
  await reset(); await k('Alt'); await k('1'); assert.equal((await st(1, 1)).bold, true); return true;
});
await T('빠른 실행 리본 위 복원·10번째 Alt→0→9', async () => {
  await ev(() => localStorage.setItem('wixel.options', JSON.stringify({ qatPosition: 'above', qatOrder: ['save', 'undo', 'redo', 'italic', 'underline', 'strike', 'alignLeft', 'alignRight', 'alignCenter', 'bold'] })));
  await page.reload({ waitUntil: 'domcontentloaded' }); await page.waitForFunction(() => !!window.tabula?.wb(), null, { timeout: 60000 }); await reset();
  assert.doesNotMatch(await page.locator('#quickAccess').getAttribute('class'), /below/);
  assert.equal(await page.locator('[data-qat-key="09"]').getAttribute('data-qat-cmd'), 'bold');
  await k('Alt'); await k('0'); await k('9'); assert.equal((await st(1, 1)).bold, true); return true;
});
await T('공개 읽기 전용: Ctrl+B / Delete / F2 / 입력 차단, 선택 허용', async () => {
  const fixture = { docName: '키보드 공개 합성 문서', workbook: { sheets: [{ name: '공개', cells: { '1,1': { raw: '123' } } }] } };
  const encoded = gzipSync(JSON.stringify(fixture)).toString('base64url');
  await page.goto('about:blank'); // 같은 페이지의 해시 이동은 앱을 다시 초기화하지 않습니다.
  await page.goto(`${process.env.WIXEL_URL || 'http://localhost:5178/'}#view=${encoded}`, { waitUntil: 'domcontentloaded' });
  await page.locator('.view-bar').waitFor({ timeout: 60000 }); await ev(() => window.tabula.selectCell(1, 1)); await page.locator('#cellEditor').focus();
  const before = await st(1, 1); await k('Control+b'); await k('Delete'); await k('F2'); await page.keyboard.type('456'); await k('Enter');
  assert.equal(await val(1, 1), 123); assert.deepEqual(await st(1, 1), before);
  await ev(() => window.tabula.selectCell(1, 1)); await k('Control+Space'); const s = await sel(); assert.deepEqual([s.r1, s.c1, s.r2, s.c2], [0, 1, MAX_ROWS - 1, 1]); return true;
});
if (errs.length) console.error('페이지 오류:\n' + errs.join('\n'));
const bad = results.filter((r) => !r.ok).length;
console.log(JSON.stringify({ total: results.length, ok: results.length - bad, bad, pageErrors: errs.length }));
if (bad || errs.length) process.exitCode = 1;
} finally {
  await browser.close();
}
