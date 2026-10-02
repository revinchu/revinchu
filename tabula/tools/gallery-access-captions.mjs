// SVG 견본과 상시 접근키 이름을 분리하는 공통 UI 회귀. 소스 서버/합성 문서 전용.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/', baseline=process.env.GALLERY_CAPTION_BASELINE;
const baselineUi=baseline?execFileSync('git',['-c','safe.directory=D:/위셀/wixel-3','show',`${baseline}:tabula/src/ui.js`],{cwd:process.cwd(),encoding:'utf8'}):null;
const browser=await chromium.launch(),results=[];let checks=0;
const eq=(a,b,m)=>{assert.deepEqual(a,b,m);checks++;};
const ok=(a,m)=>{assert.ok(a,m);checks++;};
const raf=p=>p.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
try { for(const width of [1440,1024,390,320]) {
 const context=await browser.newContext({viewport:{width,height:900}}),p=await context.newPage(),errors=[],writes=[];
 p.on('pageerror',e=>errors.push(e.message));
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}if(u.origin!==new URL(url).origin)return r.abort();if(u.pathname==='/api/health')return r.fulfill({json:{ok:true,vault:false,auth:false,publish:false}});if(u.pathname.startsWith('/api/'))return r.abort();if(baselineUi&&u.pathname==='/src/ui.js')return r.fulfill({contentType:'application/javascript',body:baselineUi});return r.continue();});
 try {
  await p.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.captionCalls=[];});await p.goto(url);await p.waitForFunction(()=>window.tabula?.wb);await p.evaluate(()=>{document.querySelectorAll('.dialog-backdrop').forEach(n=>n.remove());});
  await p.evaluate(async()=>{
   const {el,openDialog}=await import('/src/ui.js');
   const svg='<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80"><text x="2" y="20">표시할 숫자 데이터가 없습니다</text></svg>';
   const picture=()=>el('span',{class:'test-preview',style:{display:'block',height:'80px',lineHeight:'0'},html:svg});
   const body=el('div',{style:{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(120px,1fr))',gap:'14px'}},
    el('button',{id:'svgCaption',accessKey:'e',onclick:()=>captionCalls.push('caption'),style:{width:'120px',padding:'0',fontSize:'12px',whiteSpace:'normal'}},picture(),el('span',{id:'realCaption',style:{display:'block',lineHeight:'1.5',padding:'8px'}},'100% 기준 누적 세로 막대형')),
    el('button',{id:'svgOnly','aria-label':'아이콘 삽입',accessKey:'i',onclick:()=>captionCalls.push('icon')},picture()),
    el('button',{id:'explicitCaption',accessKey:'x'},el('span',{'data-access-preview':true},'값 123 견본'),el('span',{'data-access-caption-host':true,id:'explicitLabel'},'견본 선택')),
    el('label',{},picture(),el('span',{id:'inputCaption'},'범위'),el('input',{id:'captionInput',accessKey:'r'})),
    el('button',{id:'plainButton',accessKey:'p'},'일반 명령'));
   window.captionDialog=openDialog({title:'갤러리 이름 점검',width:650,body,buttons:[{label:'취소'}]});
  });await raf(p);
  eq(await p.locator('#realCaption>.access-key-hint').count(),1,'키는 미리보기 아닌 실제 이름 안에 표시');
  eq(await p.locator('.test-preview .access-key-hint').count(),0,'그림의 텍스트를 접근키 표시 위치로 사용하지 않음');
  eq(await p.locator('#svgOnly .access-key-hint').count(),0,'이름 없는 그림 단추에 문자를 덧그리지 않음');
  eq(await p.locator('#svgOnly').getAttribute('aria-keyshortcuts'),'Alt+I','아이콘 키보드 접근 보존');
  eq(await p.locator('#explicitLabel>.access-key-hint').count(),1,'명시한 이름 영역을 우선');
  eq(await p.locator('#inputCaption>.access-key-hint').count(),1,'입력 레이블도 견본을 제외');
  eq(await p.locator('#plainButton>.access-key-hint').count(),1,'일반 명령의 상시 표시 보존');
  const pos=await p.locator('#svgCaption').evaluate(n=>{const a=n.querySelector('.test-preview').getBoundingClientRect(),b=n.querySelector('#realCaption').getBoundingClientRect(),h=n.querySelector('.access-key-hint').getBoundingClientRect();return {previewBottom:a.bottom,labelTop:b.top,hintTop:h.top,hintBottom:h.bottom,labelBottom:b.bottom};});
  ok(pos.labelTop>=pos.previewBottom-.5 && pos.hintTop>=pos.labelTop-.5 && pos.hintBottom<=pos.labelBottom+.5,'미리보기·이름·키의 세로영역 겹침 없음');
  await p.locator('#svgCaption').scrollIntoViewIfNeeded();await p.locator('#svgCaption').focus();await p.keyboard.press('Alt+e');await p.locator('#svgOnly').scrollIntoViewIfNeeded();await p.keyboard.press('Alt+i');eq(await p.evaluate(()=>captionCalls),['caption','icon']);
  await p.evaluate(()=>{document.querySelector('#svgCaption .test-preview').innerHTML='<svg xmlns="http://www.w3.org/2000/svg"><text>동적 차트 제목</text></svg>';document.querySelector('#realCaption').textContent='변경된 차트 이름';});await raf(p);
  eq(await p.locator('#realCaption>.access-key-hint').count(),1,'동적 이름 갱신 뒤 중복 없음');eq(await p.locator('.test-preview .access-key-hint').count(),0);
  await p.evaluate(()=>{document.querySelector('#svgCaption').disabled=true;});await raf(p);eq(await p.locator('#realCaption>.access-key-hint').getAttribute('data-access-caption'),'(E)');
  await p.keyboard.press('Alt+e');eq(await p.evaluate(()=>captionCalls),['caption','icon'],'비활성 키 실행 차단');
  const mutations=await p.evaluate(()=>new Promise(resolve=>{let count=0;const o=new MutationObserver(rs=>count+=rs.length);o.observe(document.querySelector('.dialog'),{subtree:true,childList:true,attributes:true});setTimeout(()=>{o.disconnect();resolve(count);},150);}));eq(mutations,0,'정지 후 무한 관찰 갱신 없음');eq(errors,[]);eq(writes,[]);
  results.push({width,ok:true});console.log('OK 공통 갤러리 이름·접근키 '+width);
 }catch(e){const controls=await p.locator('.dialog [data-resolved-access-key]').evaluateAll(ns=>ns.map(n=>({tag:n.tagName,id:n.id,key:n.dataset.resolvedAccessKey,label:n.getAttribute('aria-label'),text:n.textContent.slice(-80)})));results.push({width,ok:false,error:e.stack,controls});console.error('FAIL '+width+' '+e.message);}finally{await context.close();}
 }}finally{await browser.close();}
console.log(JSON.stringify({baseline:baseline||null,cases:results.length,passed:results.filter(r=>r.ok).length,checks,results}));if(results.some(r=>!r.ok))process.exitCode=1;
