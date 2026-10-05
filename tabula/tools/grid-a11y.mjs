// 합성 문서만 검사한다. 실제 보관함 요청/쓰기 없음. NVDA 실기기 검사를 대체하지 않는다.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const base = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const fixture = { sheets: [{ name: '접근성 합성 검사', cells: { '0,0': { raw: '제목' }, '0,1': { raw: '=1+1', comment: '메모' }, '0,2': { raw: '=1/0' }, '1,1': { raw: '7' } } }] };
const publishedFixture = { workbook: fixture, docName: '접근성 게시 합성 문서', si: 0, view: { grid: true, headers: true } };
async function test(name, check, view = false) {
  const context = await browser.newContext({ viewport: { width: 1380, height: 900 } });
  const page = await context.newPage(), errors = [], writes = [];
  page.setDefaultTimeout(12000); page.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => { if (!['GET','HEAD','OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); } return route.continue(); });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(base + (view ? '#view=' + gzipSync(JSON.stringify(publishedFixture)).toString('base64url') : ''), { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => window.tabula?.wb() && document.querySelector('#accessibleGrid [role=gridcell]'));
    if (!view) await page.evaluate(f => { const t = window.tabula; t.wb().restore(f); t.gv().layout(); t.gv().renderAll(); t.selectCell(0,0); }, fixture);
    // #view 문서는 시작 격자를 만든 뒤 비동기로 설치된다. 빈 Sheet1을 검사하지 않는다.
    // readonly 자체는 여기서 기다리지 않고 아래 검사에서 독립적으로 검증한다.
    await page.waitForFunction(f => {
      const t = window.tabula, grid = document.querySelector('#accessibleGrid');
      return t?.wb().sheets[0]?.name === f.sheets[0].name && t.wb().getValue(0,0,0) === f.sheets[0].cells['0,0'].raw
        && grid?.getAttribute('aria-label') === f.sheets[0].name + ' 워크시트'
        && !document.querySelector('.load-progress');
    }, fixture);
    await page.locator('#cellEditor').focus(); await check(page, context);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '서버 쓰기');
    results.push({ name, ok:true }); console.log('OK ' + name);
  } catch(e) { results.push({ name, ok:false, error:e.message, pageErrors:errors }); console.error('NG ' + name + ': ' + e.stack); }
  finally { await context.close(); }
}
const active = p => p.evaluate(() => { const e = document.querySelector('#cellEditor'), n = document.getElementById(e.getAttribute('aria-activedescendant')); return { focus:document.activeElement.id, id:n?.id, text:n?.textContent, row:n?.getAttribute('aria-rowindex'), col:n?.getAttribute('aria-colindex'), selected:n?.getAttribute('aria-selected') }; });
const ax = async (p, context) => { const session = await context.newCDPSession(p); const tree = await session.send('Accessibility.getFullAXTree'); await session.detach(); return tree.nodes.filter(n => !n.ignored); };
try {
  await test('Chromium 접근성 트리: 단일 격자·행·셀과 활성 표시값/수식/오류', async (p, context) => {
    const nodes = await ax(p,context), grids = nodes.filter(n => n.role?.value === 'grid');
    assert.equal(grids.length,1); assert.equal(grids[0].name.value,'접근성 합성 검사 워크시트');
    assert.ok(nodes.some(n => n.role?.value === 'row'));
    assert.ok(nodes.some(n => n.role?.value === 'gridcell' && /B1, 2, 수식 =1\+1, 메모 있음/.test(n.name?.value)));
    assert.ok(nodes.some(n => n.role?.value === 'gridcell' && /C1, 오류 #DIV\/0!/.test(n.name?.value)));
    const editor = nodes.find(n => n.role?.value === 'textbox' && n.name?.value === '워크시트 셀 탐색');
    assert.ok(editor); assert.ok(editor.properties.some(v => v.name === 'activedescendant' && v.value.relatedNodes?.length));
    assert.equal((await active(p)).focus,'cellEditor');
    assert.equal(await p.locator('#accessibleGrid').getAttribute('aria-rowcount'),'1048576');
  });
  await test('키보드 이동·범위·행/열 선택의 활성 후손과 선택 안내', async p => {
    await p.evaluate(() => { window.__a11yRow = document.querySelector('#accessibleGrid [role=row]'); });
    await p.keyboard.press('ArrowRight'); assert.match((await active(p)).text,/^B1, 2/);
    assert.equal(await p.evaluate(() => window.__a11yRow === document.querySelector('#accessibleGrid [role=row]')), true, '선택만 바뀔 때 같은 접근성 행 노드 유지');
    await p.keyboard.press('Shift+ArrowDown'); assert.match((await active(p)).text,/B1:B2 범위 선택/);
    await p.keyboard.press('Control+Space'); assert.match((await active(p)).text,/B열부터 B열까지 선택/);
    await p.evaluate(() => window.tabula.selectCell(0,1));
    await p.keyboard.press('Shift+Space'); assert.match((await active(p)).text,/1행부터 1행까지 선택/);
    const a = await active(p); assert.equal(a.focus,'cellEditor'); assert.equal(a.selected,'true');
  });
  await test('F2·취소·입력 확정 중 textarea 포커스/캐럿 유지', async p => {
    await p.keyboard.press('ArrowRight'); await p.keyboard.press('F2');
    const editor = p.locator('#cellEditor'); assert.equal(await editor.getAttribute('aria-activedescendant'),null);
    assert.equal(await editor.inputValue(),'=1+1'); assert.equal(await editor.getAttribute('aria-label'),'B1 셀 편집');
    await p.keyboard.press('End'); await p.keyboard.type('+4'); await p.keyboard.press('Escape');
    assert.match((await active(p)).text,/B1, 2/);
    await p.keyboard.press('F2'); await p.keyboard.press('Control+a'); await p.keyboard.type('9'); await p.keyboard.press('Enter');
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,1)),9); assert.equal((await active(p)).focus,'cellEditor');
  });
  await test('한글 composition 이벤트가 편집 상태와 확정값을 보존한다', async p => {
    await p.evaluate(() => { const e = document.querySelector('#cellEditor'); e.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:''})); e.value='한글 입력'; e.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertCompositionText',data:'한글 입력',isComposing:true})); e.dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,data:'한글 입력'})); });
    assert.equal(await p.locator('#cellEditor').getAttribute('aria-activedescendant'),null);
    assert.equal(await p.locator('#cellEditor').inputValue(),'한글 입력'); await p.keyboard.press('Enter');
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,0)),'한글 입력'); assert.ok((await active(p)).id);
  });
  await test('틀 고정·병합·숨김 셀의 중복 제거와 AT 셀 포커스 연결', async (p,context) => {
    await p.evaluate(() => { const t=window.tabula,w=t.wb(); w.transact(()=>{w.setInput(0,0,0,'고유병합제목');w.setSheetProp(0,'merges',[{r1:0,c1:0,r2:2,c2:2}]);w.setSheetProp(0,'freeze',{rows:1,cols:1});w.setSheetProp(0,'hiddenRows',{5:true});w.setSheetProp(0,'hiddenCols',{5:true});}); t.gv().layout();t.gv().renderAll();t.selectCell(1,1); });
    assert.equal(await p.locator('#accessibleGrid [role=gridcell]').filter({hasText:'고유병합제목'}).count(),1);
    assert.equal(await p.locator('#a11y-cell-0-0-0').getAttribute('aria-rowspan'),'3'); assert.equal(await p.locator('#a11y-cell-0-0-0').getAttribute('aria-colspan'),'3');
    assert.equal(await p.locator('#accessibleGrid [aria-colindex="6"]').count(),0); assert.equal(await p.locator('#accessibleGrid [role=gridcell][aria-rowindex="6"]').count(),0);
    const nodes=await ax(p,context); assert.equal(nodes.filter(n=>n.role?.value==='gridcell'&&n.name?.value.includes('고유병합제목')).length,1);
    assert.equal(await p.locator('.pane .cells:not([aria-hidden=true])').count(),0);
    await p.locator('#a11y-cell-0-3-0').focus(); assert.equal((await active(p)).focus,'cellEditor'); assert.equal((await active(p)).row,'4');
  });
  await test('시트 보호와 숨긴 수식·잠금 해제 셀의 접근성 상태', async (p,context) => {
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>{w.setStyle(0,0,1,{hideFormula:true});w.setStyle(0,1,1,{locked:false});w.setSheetProp(0,'protect',{on:true});});t.gv().renderAll();t.selectCell(0,1);});
    assert.match((await active(p)).text,/B1, 2, 읽기 전용/); assert.doesNotMatch((await active(p)).text,/수식/);
    const nodes=await ax(p,context); assert.ok(nodes.filter(n=>n.role?.value==='gridcell').every(n=>!n.name?.value.includes('=1+1')));
    assert.equal(await p.locator('#cellEditor').getAttribute('aria-readonly'),'true');
    await p.evaluate(()=>window.tabula.selectCell(1,1)); assert.equal(await p.locator('#cellEditor').getAttribute('aria-readonly'),'false');
  });
  await test('먼 셀에서도 접근성 노드 수 제한·문자열 HTML 비실행', async p => {
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(0,900000,100,'<img src=x onerror="window.a11yInjected=true">'));t.selectCell(900000,100);});
    const a=await active(p); assert.equal(a.row,'900001');assert.equal(a.col,'101');assert.match(a.text,/<img/);
    assert.ok(await p.locator('#accessibleGrid [role=gridcell]').count()<=400);assert.equal(await p.locator('#accessibleGrid img').count(),0);
    assert.equal(await p.evaluate(()=>!!window.a11yInjected),false);
  });
  await test('공개 읽기 전용 문서의 셀·격자·편집기 상태와 편집 차단', async p => {
    assert.equal(await p.locator('#accessibleGrid').getAttribute('aria-readonly'),'true');
    assert.match((await active(p)).text,/읽기 전용/); assert.equal(await p.locator('#cellEditor').getAttribute('aria-readonly'),'true');
    await p.keyboard.press('F2'); assert.ok((await active(p)).id); assert.equal(await p.locator('#cellEditor').inputValue(),'');
  },true);
} finally { await browser.close(); }
console.log(JSON.stringify({ total:results.length,ok:results.filter(r=>r.ok).length,bad:results.filter(r=>!r.ok), note:'Chromium AX/합성 composition 검사. 실제 NVDA·OS IME는 미검증.' },null,2));
if(results.some(r=>!r.ok))process.exitCode=1;
