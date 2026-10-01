// Fresh browser contexts and synthetic sheets; never accesses a user's open workbook.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5180/', browser=await chromium.launch(), results=[];
async function test(name, width, height, fn, touch=true) {
 const c=await browser.newContext({viewport:{width,height},hasTouch:touch,isMobile:touch}),p=await c.newPage(),errors=[],writes=[];
 p.on('pageerror',e=>errors.push(e.message));p.setDefaultTimeout(8000);
 await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
 await c.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>'',writeText:async()=>{}}});});
 try{await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());await p.evaluate(()=>{const t=window.tabula;t.wb().restore({sheets:[{name:'모바일 검사',zoom:125,cells:{'0,0':{raw:'보존'},'1,0':{raw:'42'}}}]});t.switchSheet(0);});await fn(p,c);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name,ok:true});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+': '+e.stack);}finally{await c.close();}
}
const tools=p=>p.locator('#mobileTools').click();
const drawer=p=>p.getByRole('dialog',{name:'모바일 작업 도구',exact:true});
try{
 await test('자동 화면 맞춤·명시적 해제·새로고침 기억·문서 배율 보존',390,844,async p=>{
  assert.equal(await p.evaluate(()=>window.tabula.mobile().active),true);
  assert.equal(await p.evaluate(()=>window.tabula.gv().z),.85);
  assert.equal(await p.evaluate(()=>window.tabula.wb().sheets[0].zoom),125);
  await p.locator('#mobileModeToggle').click();assert.equal(await p.evaluate(()=>window.tabula.mobile().active),false);assert.equal(await p.evaluate(()=>window.tabula.gv().z),1.25);
  await p.reload();await p.waitForFunction(()=>window.tabula?.wb());assert.equal(await p.evaluate(()=>window.tabula.mobile().active),false);
  await p.locator('#mobileModeToggle').click();assert.equal(await p.evaluate(()=>window.tabula.mobile().preference),'on');
 });
 await test('모바일 도구는 키보드 없이 열리고 모든 일반 탭 이동',320,568,async p=>{
  await tools(p);assert.equal(await p.evaluate(()=>document.activeElement?.matches('input,textarea')),false);
  const ids=await drawer(p).locator('[data-mobile-tab]').evaluateAll(xs=>xs.map(x=>x.dataset.mobileTab));assert.ok(ids.includes('layout')&&ids.includes('review')&&ids.includes('data'));
  for(const id of ids.filter(x=>x!=='file')){await drawer(p).locator(`[data-mobile-tab="${id}"]`).click();assert.equal(await p.evaluate(()=>window.tabula.keytipRegistry().currentTab),id);await tools(p);}
  await drawer(p).getByRole('button',{name:'닫기',exact:true}).last().click();
 });
 await test('도구 검색·서식 실행·배율 변경·원본 유지',390,844,async p=>{
  await tools(p);await drawer(p).getByRole('searchbox').fill('굵게');await drawer(p).locator('.mobile-tool-results [data-mobile-command="bold"]').click();assert.equal(await p.evaluate(()=>window.tabula.wb().styleAt(0,0,0).bold),true);
  await tools(p);await drawer(p).getByRole('button',{name:/배율 .* 변경/}).click();const d=p.getByRole('dialog',{name:'화면 배율',exact:true});await d.locator('input').fill('110');await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal(await p.evaluate(()=>window.tabula.gv().z),1.1);assert.equal(await p.evaluate(()=>window.tabula.wb().sheets[0].zoom),125);
  await tools(p);await drawer(p).getByRole('button',{name:'화면에 맞추기',exact:true}).click();assert.equal(await p.evaluate(()=>window.tabula.gv().z),.85);
 });
 await test('도구 편집·확인 후 값 반영·키보드 숨김',390,844,async p=>{
  await p.evaluate(()=>window.tabula.selectCell(2,1));await tools(p);await drawer(p).getByRole('button',{name:'선택 셀 편집',exact:true}).click();await p.locator('#cellEditor').fill('모바일 입력');await p.locator('#fxEnter').click();assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,2,1)),'모바일 입력');assert.equal(await p.evaluate(()=>document.activeElement.id),'gridView');
 });
 await test('키보드 visualViewport 축소·복원',390,844,async p=>{
  await p.locator('#formulaInput').focus();await p.evaluate(()=>{Object.defineProperty(visualViewport,'height',{configurable:true,get:()=>360});visualViewport.dispatchEvent(new Event('resize'));});await p.waitForTimeout(100);assert.equal(await p.locator('body').evaluate(x=>x.classList.contains('mobile-keyboard')),true);assert.equal(await p.locator('#ribbon').isVisible(),false);const b=await p.locator('#gridWrap').boundingBox();assert.ok(b.height>=100);
  await p.evaluate(()=>{document.activeElement.blur();Object.defineProperty(visualViewport,'height',{configurable:true,get:()=>844});visualViewport.dispatchEvent(new Event('resize'));});await p.waitForTimeout(100);assert.equal(await p.locator('body').evaluate(x=>x.classList.contains('mobile-keyboard')),false);assert.equal(await p.locator('#ribbon').isVisible(),true);
 });
 await test('가로 회전·태블릿·데스크톱 크기 이동',390,844,async p=>{
  for(const size of [{width:844,height:390},{width:768,height:1024},{width:390,height:844}]){await p.setViewportSize(size);await p.waitForTimeout(100);assert.equal(await p.evaluate(()=>window.tabula.mobile().active),true);assert.ok(await p.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));assert.ok((await p.locator('#gridWrap').boundingBox()).height>90);}
 });
 await test('데스크톱 기본 유지·수동 모바일 모드와 복귀',1365,900,async p=>{
  assert.equal(await p.evaluate(()=>window.tabula.mobile().active),false);assert.equal(await p.locator('#mobileTools').isVisible(),false);await p.locator('#mobileModeToggle').click();assert.equal(await p.locator('#mobileTools').isVisible(),true);await tools(p);await drawer(p).getByRole('button',{name:'데스크톱 화면',exact:true}).click();assert.equal(await p.evaluate(()=>window.tabula.mobile().active),false);assert.equal(await p.evaluate(()=>window.tabula.wb().sheets[0].zoom),125);
 },false);
}finally{await browser.close();console.log(JSON.stringify({checks:results.length,bad:results.filter(x=>!x.ok).length,results}));if(results.some(x=>!x.ok))process.exitCode=1;}
