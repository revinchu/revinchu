import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1280, height: 900 } }), page = await context.newPage();
const errors = [], writes = [], results = [];
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET','HEAD','OPTIONS'].includes(r.method())) { writes.push(r.url()); return route.abort(); } return u.origin === new URL(base).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
const current = () => page.evaluate(() => tabula.wb().sheets[tabula.si].page);
const open = async () => { await page.evaluate(() => tabula.run('pageSetup')); return page.locator('.dialog').last(); };
async function test(name, fn) { try { await fn(); results.push({name,ok:true}); console.log('PASS', name); } catch(e) { results.push({name,ok:false,error:e.message}); console.log('FAIL', name, e.message); } }
try {
  await page.goto(base); await page.waitForFunction(() => window.tabula?.wb());
  await page.evaluate(() => { const w=tabula.wb(); w.restore({sheets:[{name:'초기',cells:{}},{name:"쉼표,따옴'표",cells:{'0,0':{raw:'합성'}},view:{mode:'normal',top:2,left:1,r:2,c:1}}]});tabula.switchSheet(1); });
  await test('복수 영역·반복행열·인쇄 순서 대화상자 저장과 재열기', async () => {
    const d=await open(); await d.locator('[name=area]').fill("'쉼표,따옴''표'!$A$1:$C$12,'쉼표,따옴''표'!$G$1:$H$5");
    await d.locator('[name=titleRows]').fill('1:2'); await d.locator('[name=titleCols]').fill('A:B'); await d.locator('[name=order]').selectOption('overThenDown'); await d.getByRole('button',{name:'확인',exact:true}).click();
    const pg=await current(); assert.equal(pg.areas.length,2); assert.deepEqual(pg.titleRows,[0,1]);assert.deepEqual(pg.titleCols,[0,1]);assert.equal(pg.order,'overThenDown');
    const reopened=await open();assert.equal(await reopened.locator('[name=area]').inputValue(),'A1:C12,G1:H5');await reopened.getByRole('button',{name:'취소',exact:true}).click();
  });
  await test('다른 시트의 인쇄 영역 입력은 오류를 알리고 기존 설정 보존', async () => {
    const before=await current(),d=await open();await d.locator('[name=area]').fill("'다른 시트'!A1:C5");await d.getByRole('button',{name:'확인',exact:true}).click();
    assert.ok(await page.locator('.dialog').last().innerText().then(s=>s.includes('올바르지')));assert.deepEqual(await current(),before);
    await page.keyboard.press('Escape');await page.keyboard.press('Escape');
  });
  await test('유효하지 않은 반복 행은 설정 변경 없이 거절', async () => {
    const before=await current(),d=await open();await d.locator('[name=titleRows]').fill('9:2');await d.getByRole('button',{name:'확인',exact:true}).click();assert.deepEqual(await current(),before);await page.keyboard.press('Escape');await page.keyboard.press('Escape');
  });
  await test('빈 영역 입력은 기존 복수 영역을 모두 지움', async () => {
    const d=await open();await d.locator('[name=area]').fill('');await d.getByRole('button',{name:'확인',exact:true}).click();const pg=await current();assert.equal(pg.area,null);assert.deepEqual(pg.areas,[]);
  });
  await test('전체열 입력은 저장 참조 보존·화면 계획은 사용 범위 제한', async () => {
    const d=await open();await d.locator('[name=area]').fill('$A:$C');await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal((await current()).area.r2,1048575);const plan=await page.evaluate(()=>tabula.gv().printLayout());assert.equal(plan.areas[0].area.r2,0);assert.equal(plan.error,undefined);
  });
  await test('보기 전환은 저장된 스크롤/선택 좌표 유지', async () => {
    const before=await page.evaluate(()=>({...tabula.wb().sheets[tabula.si].view}));await page.evaluate(()=>tabula.run('viewPageBreakPreview'));const after=await page.evaluate(()=>tabula.wb().sheets[tabula.si].view);assert.deepEqual(after,{...before,mode:'pageBreakPreview'});await page.evaluate(()=>tabula.run('viewNormal'));assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[tabula.si].view),{...before,mode:'normal'});
  });
  assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);console.log(JSON.stringify({tests:results.length,passed:results.filter(r=>r.ok).length,errors,writes}));if(results.some(r=>!r.ok))process.exitCode=1;
} finally { await browser.close(); }
