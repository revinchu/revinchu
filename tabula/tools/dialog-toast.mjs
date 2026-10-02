// 공통 팝업의 이전 토스트 정리 및 새 검증 안내 회귀. /src/ui.js가 있는 로컬 소스 서버 전용.
// 합성 문서와 새 격리 컨텍스트만 사용하며 외부 요청과 API/원격 쓰기를 차단한다.
import assert from 'node:assert/strict';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL || 'http://127.0.0.1:5178/';
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname)) throw Error('로컬 소스 서버에서만 실행하세요.');
const browser=await chromium.launch(),context=await browser.newContext({viewport:{width:390,height:700}}),page=await context.newPage();
const errors=[],writes=[],results=[];
page.on('pageerror',e=>errors.push(e.message));
await context.route('**/*',r=>{if(!['GET','HEAD','OPTIONS'].includes(r.request().method())){writes.push(r.request().url());return r.abort();}if(new URL(r.request().url()).origin!==new URL(url).origin || new URL(r.request().url()).pathname.startsWith('/api/'))return r.abort();return r.continue();});
const shown=()=>page.locator('#toast').evaluate(n=>n.classList.contains('show'));
async function initialNotice(mode) {
 const c=await browser.newContext({viewport:{width:390,height:700}}),p=await c.newPage();
 p.on('pageerror',e=>errors.push(e.message));
 await c.route('**/*',r=>{if(!['GET','HEAD','OPTIONS'].includes(r.request().method())){writes.push(r.request().url());return r.abort();}if(new URL(r.request().url()).origin!==new URL(url).origin||new URL(r.request().url()).pathname.startsWith('/api/'))return r.abort();return r.continue();});
 try {
  await p.clock.install();
  await p.addInitScript(()=>{window.WIXEL_SKIP_START=true;window.TABULA_STATIC=true;});
  await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula);
  if(mode==='modal') await p.evaluate(()=>tabula.run('protectSheet'));
  if(mode==='modeless') await p.evaluate(async()=>{const ui=await import('/src/ui.js');ui.openDialog({title:'초기 모델리스 검사',body:ui.el('p',{},'합성'),modeless:true,buttons:[{label:'닫기'}]});});
  assert.equal(await p.locator('#toast').evaluate(n=>n.classList.contains('show')),false,'초기 타이머 도착 전');
  await p.clock.fastForward(1300);
  assert.equal(await p.locator('#toast').evaluate(n=>n.classList.contains('show')),!mode,'팝업 유무에 따른 초기 소개 알림');
  if(mode) assert.equal(await p.getByRole('dialog').count(),1,'소개 알림 때문에 현재 창이 닫히지 않음');
  else assert.match(await p.locator('#toast').textContent(),/새로운 기능은/);
 } finally { await c.close(); }
}
try{
 await page.addInitScript(()=>{window.WIXEL_SKIP_START=true;window.TABULA_STATIC=true;});
 await page.goto(url,{waitUntil:'domcontentloaded'});await page.waitForFunction(()=>window.tabula);
 await page.evaluate(()=>{const t=tabula,w=t.wb();w.transact(()=>{w.setInput(0,0,0,'1');w.setComment(0,0,0,'합성 메모');});t.selectCell(0,0);t.run('selectComments');});assert.equal(await shown(),true);
 await page.evaluate(()=>tabula.run('protectSheet'));assert.equal(await shown(),false);assert.equal(await page.getByRole('dialog',{name:'시트 보호',exact:true}).count(),1);results.push('실제 셀 선택 안내 뒤 시트 보호를 열면 이전 토스트 제거');await page.keyboard.press('Escape');
 await page.clock.install();
 await page.evaluate(async()=>{const ui=await import('/src/ui.js');window.__popupUi=ui;ui.toast('이전 토스트');window.__popup=ui.openDialog({title:'합성 검증',body:ui.el('p',{},'확인하면 새 오류 안내'),onOpen:()=>ui.toast('창을 연 뒤 새 안내'),buttons:[{label:'확인',primary:true,action:()=>{ui.toast('유효한 값을 입력하세요.');return false;}},{label:'취소'}]});});
 assert.equal(await shown(),true);assert.equal(await page.locator('#toast').textContent(),'창을 연 뒤 새 안내');results.push('onOpen 이후 새 안내 토스트는 표시');
 await page.getByRole('dialog',{name:'합성 검증',exact:true}).getByRole('button',{name:'확인',exact:true}).click();assert.equal(await shown(),true);assert.equal(await page.locator('#toast').textContent(),'유효한 값을 입력하세요.');assert.equal(await page.getByRole('dialog',{name:'합성 검증',exact:true}).count(),1);results.push('창 안 확인 검증 토스트 유지·false 반환 시 창 유지');
 await page.clock.fastForward(1900);assert.equal(await shown(),false);results.push('새 검증 토스트의 기존 자동 숨김 유지');
 await page.evaluate(()=>{const ui=window.__popupUi;ui.toast('부모 창 안내');window.__child=ui.openDialog({title:'하위 합성 창',body:ui.el('p',{},'하위 창'),buttons:[{label:'닫기'}]});});assert.equal(await shown(),false);assert.equal(await page.getByRole('dialog').count(),2);await page.keyboard.press('Escape');assert.equal(await page.getByRole('dialog').count(),1);results.push('하위 창에서도 이전 안내만 제거·부모 창 복귀 유지');
 await page.keyboard.press('Escape');await page.evaluate(()=>{const ui=window.__popupUi;ui.toast('다음 창 이전 안내');window.__modeless=ui.openDialog({title:'모델리스 합성 창',body:ui.el('p',{},'합성'),modeless:true,buttons:[{label:'닫기'}]});});assert.equal(await shown(),false);await page.keyboard.press('Escape');results.push('모델리스 창도 이전 안내 제거·닫기 유지');
 for(const mode of ['modal','modeless']) await initialNotice(mode);
 results.push('초기 1.2초 소개 알림은 모달·모델리스 창이 열려 있으면 생략');
 await initialNotice(null);results.push('열린 창이 없으면 초기 소개 알림 정상 표시');
 assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);const output={cases:results.length,results,pageErrors:errors,blockedWrites:writes};console.log(JSON.stringify(output,null,2));
}finally{await context.close();await browser.close();}
