// WIXEL 3 소스 서버용 회귀(/src 모듈 상태 검사 포함). 새 컨텍스트/가상 API만 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
const test = async (name, fn, { cloud = false, home = false } = {}) => {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const errors = [];
  page.setDefaultTimeout(7000);
  page.on('pageerror', (e) => errors.push(e.message));
  try {
    await page.addInitScript(({ cloud, home }) => { window.WIXEL_SKIP_START = !home; if (!cloud) window.TABULA_STATIC = true; }, { cloud, home });
    if (cloud) await page.route('**/api/health', (route) => route.fulfill({ json: { ok: true, vault: true, auth: false, publish: true } }));
    await fn(page, async () => { await page.goto(process.env.WIXEL_URL || 'http://localhost:5180/'); await page.waitForFunction(() => !!window.tabula); });
    assert.deepEqual(errors, []);
    console.log(`OK ${name}`); results.push({ name, ok: true });
  } catch (e) { console.error(`NG ${name}: ${e.stack}`); results.push({ name, ok: false }); }
  finally { await page.close(); }
};
const grid = async (page) => { await page.locator('#cellEditor').focus(); };
try {
  await test('첫 화면·4개 저장 방식·파일 다운로드', async (p, start) => {
    await p.addInitScript(() => { Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined }); });
    await start(); await p.waitForSelector('.wixel-hub');
    assert.match(await p.locator('h1').innerText(), /새로운 작업 공간/);
    await p.locator('[data-page=storage]').click(); assert.equal(await p.locator('.storage-cards .hub-card').count(), 4);
    assert.doesNotMatch(await p.locator('.backstage-main').innerText(), /\b(?:null|undefined)\b/);
    await p.getByRole('button', { name: /Excel 파일 다운로드/ }).click();
    const save = p.getByRole('dialog', { name: '파일로 저장', exact: true });
    await save.getByRole('textbox', { name: '파일 이름', exact: true }).fill('위셀-다운로드-회귀.xlsx');
    assert.match(await save.innerText(), /폴더|저장 위치/);
    const download = p.waitForEvent('download', { timeout: 30000 });
    await save.getByRole('button', { name: '다운로드', exact: true }).click();
    assert.equal((await download).suggestedFilename(), '위셀-다운로드-회귀.xlsx');
    await p.screenshot({ path: process.env.WIXEL_STORAGE_SCREENSHOT || 'D:/Codex/Temp/wixel3-storage.png' });
    await p.locator('[data-page=new]').click(); await p.screenshot({ path: process.env.WIXEL_HOME_SCREENSHOT || 'D:/Codex/Temp/wixel3-home.png' });
  }, { home: true });
  await test('Ctrl+Space 행선택→활성열·여러열·전체선택', async (p, start) => {
    await start(); await p.evaluate(() => window.tabula.selectCell(4, 3)); await grid(p);
    await p.keyboard.press('Shift+Space'); await p.keyboard.press('Control+Space');
    assert.deepEqual(await p.evaluate(() => [window.tabula.sel.c1, window.tabula.sel.c2, window.tabula.sel.r1]), [3, 3, 0]);
    await p.evaluate(() => window.tabula.selectRange({ r1: 3, c1: 2, r2: 5, c2: 4 })); await grid(p); await p.keyboard.press('Control+Space');
    assert.deepEqual(await p.evaluate(() => [window.tabula.sel.c1, window.tabula.sel.c2]), [2, 4]);
    await p.keyboard.press('Control+Shift+Space'); assert.equal(await p.evaluate(() => window.tabula.sel.c2), 16383);
  });
  await test('빠른 실행 순서·리본 아래·Alt 키·설정 복원', async (p, start) => {
    await start(); await p.evaluate(() => window.tabula.run('options'));
    await p.getByRole('tab', { name: '빠른 실행 도구 모음', exact: true }).click();
    await p.getByLabel('표시 위치', { exact: true }).selectOption('below');
    await p.getByLabel('사용 가능한 명령').selectOption('bold'); await p.getByRole('tabpanel', { name: '빠른 실행 도구 모음', exact: true }).getByRole('button', { name: '추가', exact: true }).click();
    const moveCount = await p.getByLabel('현재 도구 모음 순서').evaluate((select) => select.selectedIndex);
    for (let i = 0; i < moveCount; i++) await p.getByRole('button', { name: '위로', exact: true }).click();
    await p.getByRole('button', { name: '확인', exact: true }).click();
    assert.match(await p.locator('#quickAccess').getAttribute('class'), /below/);
    assert.equal(await p.locator('[data-qat-key="1"]').getAttribute('data-qat-cmd'), 'bold');
    await p.evaluate(() => window.tabula.selectCell(0, 0)); await grid(p); await p.keyboard.press('Alt+1');
    assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).bold), true);
    await p.reload(); await p.waitForFunction(() => !!window.tabula);
    assert.match(await p.locator('#quickAccess').getAttribute('class'), /below/);
    assert.equal(await p.locator('[data-qat-key="1"]').getAttribute('data-qat-cmd'), 'bold');
    await grid(p); await p.keyboard.press('Alt'); await p.keyboard.press('h'); await p.keyboard.press('a'); await p.keyboard.press('l');
    assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).align), 'left');
  });
  await test('빠른 실행 10번째 Alt→0→9', async (p, start) => {
    await p.addInitScript(() => localStorage.setItem('wixel.options', JSON.stringify({ qatOrder: ['save', 'undo', 'redo', 'italic', 'underline', 'strike', 'alignLeft', 'alignRight', 'alignCenter', 'bold'] })));
    await start(); assert.equal(await p.locator('[data-qat-key="09"]').getAttribute('data-qat-cmd'), 'bold');
    await grid(p); await p.keyboard.press('Alt'); await p.keyboard.press('0'); await p.keyboard.press('9');
    assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).bold), true);
  });
  await test('개인 보관함 연결만으로 업로드하지 않음·저장 충돌 보호', async (p, start) => {
    let puts = 0, revision = 0; const seen = [];
    await p.route('**/api/files**', async (route) => {
      const req = route.request();
      if (req.method() === 'GET') return route.fulfill({ json: [] });
      puts++; seen.push(req.headers());
      if (req.headers()['if-match'] !== `"${revision}"`) return route.fulfill({ status: 412, json: { error: '다른 기기에서 바뀐 문서입니다', code: 'REVISION_CONFLICT', currentRevision: revision } });
      revision++; return route.fulfill({ json: { ok: true, revision, modified: Date.now() } });
    });
    await start(); await p.waitForFunction(async () => (await import('/src/storage.js')).server.available);
    await p.evaluate(() => window.tabula.run('saveLocations'));
    await p.getByRole('button', { name: '다른 복구키로 연결', exact: true }).click();
    await p.getByRole('textbox', { name: '개인 보관함 복구키' }).fill('A'.repeat(43));
    await p.getByRole('button', { name: '기존 복구키로 연결', exact: true }).click();
    assert.equal(puts, 0);
    await p.getByRole('button', { name: /온라인 개인 보관함/ }).click();
    await p.getByRole('dialog', { name: '서버에 저장 (다른 기기에서 열기)', exact: true }).getByRole('button', { name: '저장', exact: true }).click();
    await p.waitForFunction(() => document.querySelector('#saveState').textContent.includes('저장됨'));
    assert.equal(puts, 1); assert.equal(seen[0]['x-wixel-vault'], 'A'.repeat(43)); assert.equal(seen[0]['if-match'], '"0"');
    await p.keyboard.press('Escape'); revision = 2;
    await p.evaluate(() => window.tabula.wb().transact(() => window.tabula.wb().setInput(0, 0, 0, '91')));
    await p.getByRole('dialog', { name: '저장 충돌 — 사본을 유지했습니다', exact: true }).waitFor();
    assert.equal(puts, 2); assert.equal(seen[1]['if-match'], '"1"');
    await p.getByRole('button', { name: '브라우저에 계속 작업', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 91);
  }, { cloud: true });
  await test('Google Sheets 값 가져오기와 수식 주입 방지', async (p, start) => {
    const reply = (route) => route.fulfill({ contentType: 'text/csv', body: '부서,금액\n서울,120\n부산,=1+1' });
    await p.route('https://docs.google.com/**', reply);
    await p.route('**/api/fetch?**', (route) => {
      assert.equal(new URL(new URL(route.request().url()).searchParams.get('url')).hostname, 'docs.google.com');
      return reply(route);
    });
    await start(); await p.evaluate(() => window.tabula.run('googleSheets'));
    await p.getByLabel('Google Sheets 공유 주소').fill('https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwx/edit#gid=0');
    await p.getByRole('button', { name: '가져오기', exact: true }).click();
    await p.getByRole('dialog', { name: 'Google Sheets 가져오기', exact: true }).waitFor({ state: 'hidden' });
    assert.deepEqual(await p.evaluate(() => [window.tabula.wb().getValue(1, 1, 1), window.tabula.wb().getValue(1, 2, 1)]), [120, '=1+1']);
  });
  await test('피벗 필드 키보드 이동·값 설정·실행 취소', async (p, start) => {
    await start(); await p.evaluate(() => {
      const t = window.tabula, w = t.wb(); w.transact(() => {
        [['부서', '매출'], ['서울', '100'], ['부산', '200']].forEach((row, r) => row.forEach((v, c) => w.setInput(0, r, c, v)));
        w.addSheet('보고서'); w.setSheetProp(1, 'pivot', { name: '검증', source: w.sheets[0].name, range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['부서'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 2, left: 1 });
      }); t.switchSheet(1); t.run('pivotRefresh'); t.selectCell(2, 1);
    });
    const row = p.locator('[data-pivot-area=rows] .pp-item').first(); await row.focus(); await p.keyboard.press('Alt+ArrowDown');
    await p.getByRole('menuitem', { name: 'Σ 값(으)로 이동', exact: true }).click();
    assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.rows), []);
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.values.length), 2);
    await p.locator('[data-pivot-area=values] .pp-item').first().focus(); await p.keyboard.press('Enter');
    const dialog = p.getByRole('dialog', { name: '값 필드 설정', exact: true });
    await dialog.locator('.vf-list').first().selectOption('average'); await dialog.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.values[0].agg), 'average');
    await p.evaluate(() => window.tabula.run('undo')); assert.equal(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.values[0].agg), 'sum');
  });
  await test('셀 서식 키보드 탭·맞춤 미리보기·보호 적용', async (p, start) => {
    await start(); await p.evaluate(() => window.tabula.run('formatCells'));
    await p.getByRole('tab', { name: '표시 형식', exact: true }).focus(); await p.keyboard.press('ArrowRight');
    assert.equal(await p.getByRole('tab', { name: '맞춤', exact: true }).getAttribute('aria-selected'), 'true');
    await p.getByLabel('가로', { exact: true }).selectOption('right');
    assert.equal(await p.locator('.fc-align-preview').evaluate((el) => el.style.justifyContent), 'flex-end');
    await p.getByRole('tab', { name: '보호', exact: true }).click(); await p.getByLabel('잠금', { exact: true }).uncheck(); await p.getByLabel('숨김', { exact: true }).check();
    await p.getByRole('button', { name: '확인', exact: true }).click();
    assert.deepEqual(await p.evaluate(() => { const s = window.tabula.wb().styleAt(0, 0, 0); return [s.align, s.locked, s.hideFormula]; }), ['right', false, true]);
  });
  await test('다른 WIXEL 문서 참조 기본 차단·현재 문서만 허용·새 문서 권한 초기화', async (p, start) => {
    let reads = 0;
    await p.addInitScript(() => localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'vault', key: 'A'.repeat(43) })));
    await p.route('**/api/files**', (route) => {
      if (route.request().url().endsWith('/api/files')) return route.fulfill({ json: [] });
      reads++; return route.fulfill({ json: { docName: '개인 자료', workbook: { sheets: [{ name: 'Sheet1', cells: { '0,0': { raw: '987' } } }] } } });
    });
    await start(); await p.waitForFunction(async () => (await import('/src/storage.js')).server.connected);
    await p.evaluate(() => { const w = window.tabula.wb(); w.setInput(0, 0, 0, '=IMPORTRANGE("개인 자료","A1")'); w.getValue(0, 0, 0); });
    await p.waitForFunction(async () => [...(await import('/src/fx-web.js')).NET.cache.values()].some((v) => v.state === 'error'));
    assert.equal(reads, 0);
    await p.evaluate(() => window.tabula.run('privateImportPermission'));
    await p.getByRole('button', { name: '현재 문서에서 참조 허용', exact: true }).click();
    await p.waitForFunction(() => window.tabula.wb().getValue(0, 0, 0) === 987); assert.equal(reads, 1);
    await p.evaluate(async () => { await window.tabula.newWorkbook(); });
    await p.waitForFunction(() => window.tabula.wb().getRaw(0, 0, 0) === '');
    await p.evaluate(() => { const w = window.tabula.wb(); w.setInput(0, 0, 0, '=IMPORTRANGE("개인 자료","A1")'); w.getValue(0, 0, 0); });
    await p.waitForFunction(async () => [...(await import('/src/fx-web.js')).NET.cache.values()].some((v) => v.state === 'error'));
    assert.equal(reads, 1);
  }, { cloud: true });
  await test('공개 문서와 편집용 사본은 개인 보관함 참조·자동 저장 차단', async (p) => {
    let privateRequests = 0;
    await p.addInitScript(() => localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'vault', key: 'A'.repeat(43) })));
    await p.route('**/api/files**', (route) => { privateRequests++; return route.fulfill({ json: [] }); });
    await p.route('**/api/published/**', (route) => route.fulfill({ json: { docName: '외부 공유 문서', workbook: { sheets: [{ name: '공유', cells: { '0,0': { raw: '=IMPORTRANGE("개인 자료","A1")' } } }] } } }));
    await p.goto(`${process.env.WIXEL_URL || 'http://localhost:5180/'}?view=public-test`);
    await p.waitForSelector('.view-bar');
    await p.waitForFunction(async () => [...(await import('/src/fx-web.js')).NET.cache.values()].some((v) => v.state === 'error'));
    assert.equal(privateRequests, 0);
    await p.evaluate(() => window.tabula.run('privateImportPermission'));
    const dialog = p.getByRole('dialog', { name: '문서 참조 권한', exact: true }); assert.match(await dialog.innerText(), /참조할 수 없습니다/);
    await dialog.getByRole('button', { name: '확인', exact: true }).click();
    await p.evaluate(async () => { const { newShape } = await import('/src/shapes.js'); window.tabula.wb().setSheetProp(0, 'shapes', [{ ...newShape('rect', { x: 400, y: 100, w: 180, h: 60 }), id: 'public-macro', text: '외부 버튼', macro: 'GetNaverAdKeyword' }]); window.tabula.gv().renderAll(); });
    await p.locator('[data-id="public-macro"]').first().click();
    assert.match(await p.locator('#toast').innerText(), /매크로를 실행할 수 없습니다/);
    await p.getByRole('button', { name: '편집용 사본 만들기', exact: true }).click();
    await p.waitForFunction(async () => [...(await import('/src/fx-web.js')).NET.cache.values()].some((v) => v.state === 'error'));
    assert.equal(privateRequests, 0);
  }, { cloud: true });
  const bad = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ total: results.length, ok: results.length - bad, bad }));
  if (bad) process.exitCode = 1;
} finally { await browser.close(); }
