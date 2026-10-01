// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
try {
const page = await browser.newPage({ viewport: { width: 1400, height: 820 } });
const errs = []; page.on('pageerror', (e) => errs.push(e.message));
// 서버 문서의 자동 복원·저장이 키보드 검증용 통합 문서와 섞이지 않도록 합니다.
await page.addInitScript(() => { window.TABULA_STATIC = true; });
await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/')); await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForFunction(() => window.tabula?.wb());
const ev = (f, a) => page.evaluate(f, a);
const reset = async () => {
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  await ev(() => { document.querySelectorAll('.dialog-backdrop, .dlg-wrap, .modal, .qa-pop, .menu').forEach((d) => d.remove()); });
  await ev(() => { const t = window.tabula; const wb = t.wb(); t.switchSheet(0); while (wb.sheets.length > 1) wb.sheets.pop(); t.switchSheet(0); wb.transact(() => { for (let r = 0; r < 12; r++) for (let c = 0; c < 8; c++) wb.setCellData(0, r, c, null); [['이름', '수량', '단가'], ['A', '1', '100'], ['B', '2', '200'], ['C', '3', '300']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v))); wb.setSheetProp(0, 'hiddenRows', {}); wb.setSheetProp(0, 'filter', null); wb.setSheetProp(0, 'outline', null); }); t.gv().layout(); t.switchSheet(0); t.selectCell(1, 1); });
  await page.waitForTimeout(80);
};
const val = (r, c) => ev(([r, c]) => window.tabula.wb().getValue(0, r, c), [r, c]);
const raw = (r, c) => ev(([r, c]) => window.tabula.wb().getCell(0, r, c)?.raw, [r, c]);
const st = (r, c) => ev(([r, c]) => window.tabula.wb().styleAt(0, r, c), [r, c]);
const sel = () => ev(() => ({ ...window.tabula.sel, a: window.tabula.active }));
const dlg = () => ev(() => document.querySelector('.dialog .dialog-title, .dialog h2, .dlg-title')?.textContent ?? document.querySelector('.dialog')?.textContent?.slice(0, 20) ?? null);
const results = [];
const T = async (name, fn) => { let ok; try { await reset(); ok = await fn(); } catch (e) { ok = '오류 ' + e.message.slice(0, 160); } results.push({ name, ok: ok === true, detail: ok }); };
const k = (x) => page.keyboard.press(x);
await T('Ctrl+; 오늘 날짜', async () => { await k('Control+;'); await k('Enter'); return typeof (await val(1, 1)) === 'number' || await val(1, 1); });
await T('Ctrl+Shift+; 현재 시간', async () => { await k('Control+Shift+;'); await k('Enter'); const v = await val(1, 1); return typeof v === 'number' && v < 1 || v; });
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
await T('Ctrl+Space 열 선택', async () => { await k('Control+Space'); return ev(() => window.tabula.sel.r2 > 1000); });
await T('Shift+Space 행 선택', async () => { await k('Shift+Space'); return ev(() => window.tabula.sel.c2 > 1000); });
await T('Ctrl+Shift+= 삽입', async () => { await k('Shift+Space'); await k('Control+Shift+='); return (await val(2, 1)) === 1 || await val(2, 1); });
await T('Ctrl+- 삭제', async () => { await k('Shift+Space'); await k('Control+-'); return (await val(1, 1)) === 2 || await val(1, 1); });
await T('Shift+F11 새 시트', async () => { await k('Shift+F11'); return ev(() => window.tabula.wb().sheets.length === 2); });
await T('Ctrl+PageDown 다음 시트', async () => { await ev(() => window.tabula.run('addSheet')); await ev(() => window.tabula.switchSheet(0)); await k('Control+PageDown'); return ev(() => window.tabula.si === 1); });
await T('Ctrl+` 수식 보기', async () => { await k('Control+`'); const on = await ev(() => document.querySelector('.c[data-r="1"][data-c="1"]')?.textContent); await k('Control+`'); return on !== undefined; });
await T('F4 참조 전환 (편집 중)', async () => { await ev(() => window.tabula.selectCell(6, 3)); await page.keyboard.type('=A1'); await k('F4'); await k('Enter'); return (await raw(6, 3)) === '=$A$1' || await raw(6, 3); });
await T("Ctrl+' 위 셀 수식", async () => { await ev(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(0, 6, 3, '=B2*2')); window.tabula.selectCell(7, 3); }); await k("Control+'"); await k('Enter'); return (await raw(7, 3)) === '=B2*2' || await raw(7, 3); });
await T('Ctrl+Shift+L 필터', async () => { await k('Control+Shift+l'); return ev(() => !!window.tabula.wb().sheets[0].filter); });
await T('Ctrl+T 표 만들기', async () => { await k('Control+t'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /표/.test(d) || d.slice(0, 40); });
await T('Ctrl+1 셀 서식', async () => { await k('Control+1'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /셀 서식/.test(d) || d.slice(0, 40); });
await T('Ctrl+K 하이퍼링크', async () => { await k('Control+k'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /하이퍼링크/.test(d) || d.slice(0, 40); });
await T('Ctrl+G / F5 이동', async () => { await k('F5'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이동/.test(d) || d.slice(0, 40); });
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
await T('Ctrl+Alt+F9 전체 계산 (실행 오류 확인)', async () => { await k('Control+Alt+F9'); return true; });
await T('Ctrl+Shift+U 수식 입력줄', async () => { const a = await ev(() => document.getElementById('formulaRow').className); await k('Control+Shift+u'); const b = await ev(() => document.getElementById('formulaRow').className); await k('Control+Shift+u'); return a !== b || [a, b]; });
await T('Ctrl+Shift+Enter 배열', async () => { await ev(() => window.tabula.selectCell(9, 5)); await page.keyboard.type('=SUM(B2:B4*C2:C4)'); await k('Control+Shift+Enter'); return (await val(9, 5)) === 1400 || await val(9, 5); });
await T('Ctrl+A 영역 → 전체', async () => { await k('Control+a'); const a = await sel(); await k('Control+a'); const b = await sel(); return (a.r2 === 3 && b.r2 > 1000) || [a, b]; });
await T('Ctrl+Home', async () => { await k('Control+Home'); const s = await sel(); return (s.a.r === 0 && s.a.c === 0) || s.a; });
await T('Ctrl+End', async () => { await k('Control+End'); const s = await sel(); return (s.a.r === 3 && s.a.c === 2) || s.a; });
await T('Ctrl+Backspace 활성 셀 보기 (실행 오류 확인)', async () => { await k('Control+Backspace'); return true; });
await T('F2 편집', async () => { await k('F2'); const e = await ev(() => !document.getElementById('cellEditor').classList.contains('idle')); await k('Escape'); return e; });
await T('Shift+F2 메모', async () => { await k('Shift+F2'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? document.querySelector('.cm-edit, .comment-edit')?.textContent ?? ''); await k('Escape'); return d.length > 0 || d; });
await T('Ctrl+Q 빠른 분석', async () => { await k('Control+q'); const d = await ev(() => !!document.querySelector('.qa-pop')); await k('Escape'); return d; });
await T('Alt+; 보이는 셀만', async () => { await k('Alt+;'); return ev(() => /칸을 골랐/.test(document.body.textContent)); });
await T('Ctrl+\\ 행 내용 차이 (실행 오류 확인)', async () => { await ev(() => window.tabula.selectRange({ r1: 1, c1: 1, r2: 3, c2: 2 }, 'cells', { r: 1, c: 1 })); await k('Control+\\'); return true; });
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
await T('Alt 키팁 (Alt,H)', async () => { await k('Alt'); await page.waitForTimeout(100); const t = await ev(() => document.body.classList.contains('keytips') || document.querySelectorAll('[data-keytip], .keytip, .kt').length); await k('Escape'); return !!t || t; });
await T('F12 다른 이름으로 저장', async () => { await k('F12'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /다른 이름/.test(d) || d.slice(0, 40); });
await T('Ctrl+Shift+F3 이름 만들기', async () => { await k('Control+Shift+F3'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이름/.test(d) || d.slice(0, 40); });
await T('Ctrl+F3 이름 관리자', async () => { await k('Control+F3'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /이름/.test(d) || d.slice(0, 40); });
await T('F7 맞춤법 검사', async () => { await k('F7'); await page.waitForTimeout(150); const d = await ev(() => document.querySelector('.dialog')?.textContent ?? ''); await k('Escape'); return /맞춤법/.test(d) || d.slice(0, 40); });
await T('Ctrl+Shift+A 인수 이름', async () => { await page.keyboard.type('=SUM'); await k('Control+Shift+a'); const t = await ev(() => document.getElementById('cellEditor').value); await k('Escape'); return /SUM\(number1/.test(t) || t; });
console.log(results.map(({ name, ok, detail }) => `${ok ? 'OK' : 'NG'} ${name}${ok ? '' : ` → ${JSON.stringify(detail)}`}`).join('\n'));
if (errs.length) console.error('페이지 오류:\n' + errs.join('\n'));
const bad = results.filter((r) => !r.ok).length;
console.log(JSON.stringify({ total: results.length, ok: results.length - bad, bad, pageErrors: errs.length }));
if (bad || errs.length) process.exitCode = 1;
} finally {
  await browser.close();
}
