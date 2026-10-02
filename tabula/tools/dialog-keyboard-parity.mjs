// 공통 UI 모듈을 불러오는 소스 서버 전용 합성 회귀. 사용자 문서/원격 쓰기를 사용하지 않는다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
assert.ok(['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname));
const browser = await chromium.launch(), results = [];
let assertions = 0;
const eq = (a,b,m) => { assertions++; assert.deepEqual(a,b,m); };
const ok = (a,m) => { assertions++; assert.ok(a,m); };
const ui = (p,fn) => p.evaluate(async source => { const module = await import('/src/ui.js'); return new Function('ui', `return (${source})(ui)`)(module); }, fn.toString());
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function test(name, action) {
 const context=await browser.newContext({viewport:{width:1280,height:900}}), p=await context.newPage(), errors=[],blocked=[];
 p.setDefaultTimeout(10000); p.on('pageerror',e=>errors.push(e.message));
 await context.route('**/*',r=>{const q=r.request(); if(!['GET','HEAD','OPTIONS'].includes(q.method()) || new URL(q.url()).origin!==new URL(url).origin){blocked.push(q.method());return r.abort();}return r.continue();});
 try {
  await p.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.parityCalls=[];});
  await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());
  await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.restore({sheets:[{name:'합성',cells:{'0,0':{raw:'원문'}}}]});t.gv().layout();t.gv().renderAll();t.selectCell(0,0);window.gridKeys=0;document.querySelector('#cellEditor').addEventListener('keydown',()=>window.gridKeys++);});
  await action(p);eq(errors,[],'페이지 오류');eq(blocked,[],'외부 요청/쓰기');results.push({name,ok:true});console.log('OK '+name);
 } catch(error){results.push({name,ok:false,error:error.message});console.error('NG '+name+': '+error.stack);}finally{await context.close();}
}
try {
 await test('지연 격자 focus와 Enter가 모달 기본 동작으로만 전달',async p=>{
  await ui(p,({openDialog,el})=>{window.parityDialog=openDialog({title:'누수 검사',body:el('input',{id:'insideInput',value:'입력'}),buttons:[{label:'확인',primary:true,action:()=>parityCalls.push('ok')}]});setTimeout(()=>document.querySelector('#cellEditor').focus(),0);});
  await raf(p);eq(await p.evaluate(()=>document.activeElement.id),'insideInput');await p.keyboard.press('Enter');
  eq(await p.evaluate(()=>parityCalls),['ok']);eq(await p.getByRole('dialog').count(),0);eq(await p.evaluate(()=>gridKeys),0);eq(await p.evaluate(()=>window.tabula.wb().getRaw(0,0,0)),'원문');eq(await p.evaluate(()=>window.tabula.active),{r:0,c:0});
 });
 await test('본문 재작성 후 body 대상으로 온 Enter도 셀로 새지 않음',async p=>{
  await ui(p,({openDialog,el})=>{const body=el('div',{},el('input',{id:'oldInput'}));window.parityDialog=openDialog({title:'본문 교체',body,buttons:[{label:'확인',primary:true,action:()=>{parityCalls.push('ok');return false;}}]});body.replaceChildren(el('input',{id:'newInput'}));document.body.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',bubbles:true,cancelable:true}));});
  eq(await p.evaluate(()=>parityCalls),['ok']);eq(await p.evaluate(()=>document.activeElement.id),'newInput');eq(await p.evaluate(()=>gridKeys),0);eq(await p.getByRole('dialog').count(),1);
 });
 await test('initialFocus와 단계별 defaultAction·native button Enter',async p=>{
  await ui(p,({openDialog,el})=>{const next=el('button',{id:'nextNative',onclick:()=>parityCalls.push('button')},'다음');window.parityDialog=openDialog({title:'단계별 동작',body:el('div',{},el('input',{id:'stepInput'}),next),initialFocus:()=>next,defaultAction:()=>parityCalls.push('default'),buttons:[{label:'마침',primary:true,action:()=>parityCalls.push('finish')}]});});
  eq(await p.evaluate(()=>document.activeElement.id),'nextNative');await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['button']);
  await p.locator('#stepInput').focus();await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['button','default']);eq(await p.getByRole('dialog').count(),1);
 });
 await test('textarea·IME·자체 Enter 처리·비동기 중복 방지',async p=>{
  await ui(p,({openDialog,el})=>{window.parityDialog=openDialog({title:'입력 동작',body:el('div',{},el('textarea',{id:'multiline'},'첫'),el('input',{id:'imeField'}),el('input',{id:'handledEnter',onkeydown:e=>{if(e.key==='Enter'){e.preventDefault();parityCalls.push('handled');}}})),defaultAction:()=>{parityCalls.push('async');return new Promise(r=>window.parityRelease=r);}});});
  await p.locator('#multiline').fill('첫');await p.locator('#multiline').press('End');await p.keyboard.press('Enter');eq(await p.locator('#multiline').inputValue(),'첫\n');eq(await p.evaluate(()=>parityCalls),[]);await p.keyboard.down('Enter');await p.keyboard.down('Enter');await p.keyboard.up('Enter');eq(await p.locator('#multiline').inputValue(),'첫\n\n\n','textarea Enter 반복은 줄 바꿈 유지');
  await p.locator('#imeField').focus();await p.locator('#imeField').evaluate(e=>e.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',isComposing:true,keyCode:229,bubbles:true,cancelable:true})));eq(await p.evaluate(()=>parityCalls),[]);
  await p.locator('#handledEnter').press('Enter');eq(await p.evaluate(()=>parityCalls),['handled']);
  await p.locator('#imeField').press('Enter');await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['handled','async']);await p.evaluate(()=>parityRelease());await raf(p);eq(await p.getByRole('dialog').count(),1);
 });
 await test('중첩 모달과 소유 메뉴는 올바른 범위만 실행',async p=>{
  await ui(p,({openDialog,openMenu,el})=>{openDialog({title:'부모',body:el('button',{id:'openOwnedMenu',onclick:e=>openMenu(e.currentTarget,[{label:'메뉴 동작',accessKey:'a',action:()=>parityCalls.push('menu')}])},'메뉴'),buttons:[{label:'확인',primary:true,action:()=>{parityCalls.push('parent');return false;}}]});openDialog({title:'자식',body:el('input'),buttons:[{label:'확인',primary:true,action:()=>parityCalls.push('child')}]});});
  await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['child']);await p.locator('#openOwnedMenu').click();await p.keyboard.press('Alt+a');eq(await p.evaluate(()=>parityCalls),['child','menu']);
  await p.getByRole('dialog',{name:'부모',exact:true}).getByRole('button',{name:'확인',exact:true}).focus();await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['child','menu','parent']);
 });
 await test('모델리스에서 격자 탐색 및 부모 모달 소유 모델리스 허용',async p=>{
  await ui(p,({openDialog,el})=>{window.parityDialog=openDialog({title:'모델리스',modeless:true,body:el('input'),buttons:[{label:'확인',primary:true,action:()=>parityCalls.push('modeless')}]});});
  await p.locator('#cellEditor').focus();await p.keyboard.press('ArrowDown');eq(await p.evaluate(()=>window.tabula.active),{r:1,c:0});eq(await p.evaluate(()=>parityCalls),[]);await p.evaluate(()=>parityDialog.close());
  await ui(p,({openDialog,el})=>{openDialog({title:'모달 부모',body:el('input')});openDialog({title:'소유 모델리스',modeless:true,body:el('input',{id:'ownedInput'}),buttons:[{label:'작업',accessKey:'a',action:()=>parityCalls.push('owned')}]});});
  eq(await p.evaluate(()=>document.activeElement.id),'ownedInput');await p.keyboard.press('Alt+a');eq(await p.evaluate(()=>parityCalls),['owned']);
 });
 await test('Alt 없이 항상 표시·기존 괄호 중복 없음·accessible name 보존',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'상시 표시',body:el('div',{},el('label',{},el('span',{},'찾을 내용'),el('input',{id:'captionInput',accessKey:'n'})),el('button',{id:'disabledNext',disabled:true,accessKey:'w'},'다음'),el('button',{id:'legacyLabel',accessKey:'f'},'마침(F)')),buttons:[{label:'취소'}]}));await raf(p);
  eq(await p.locator('.access-key-layer').count(),0);eq(await p.locator('#disabledNext .access-key-hint').getAttribute('data-access-caption'),'(W)');eq(await p.locator('#legacyLabel .access-key-hint').count(),0);eq(await p.locator('.dialog-close .access-key-hint').count(),0);
  eq(await p.getByRole('button',{name:'다음',exact:true}).count(),1);eq(await p.getByRole('textbox',{name:'찾을 내용',exact:true}).count(),1);
  const hint=p.locator('#captionInput').locator('..').locator('.access-key-hint');eq(await hint.getAttribute('data-access-caption'),'(N)');ok(await hint.evaluate(e=>e.getBoundingClientRect().width>0 && getComputedStyle(e,'::after').content.includes('(N)')),'실제 CSS 표시');
 });
 await test('disabled 전환과 동적 항목 추가에서 실제 키·문자 일치',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'동적 키',body:el('div',{id:'dynamicBody'},el('button',{id:'stableButton',accessKey:'p',onclick:()=>parityCalls.push('stable')},'동작'),el('button',{id:'toggleButton',accessKey:'n',disabled:true,onclick:()=>parityCalls.push('next')},'다음'))}));await raf(p);
  await p.keyboard.press('Alt+n');eq(await p.evaluate(()=>parityCalls),[]);
  await p.evaluate(()=>{document.querySelector('#toggleButton').disabled=false;document.querySelector('#dynamicBody').append(Object.assign(document.createElement('button'),{textContent:'새 옵션'}));});await raf(p);
  eq(await p.locator('#toggleButton').getAttribute('data-resolved-access-key'),'n');eq(await p.locator('#toggleButton .access-key-hint').getAttribute('data-access-caption'),'(N)');await p.keyboard.press('Alt+n');eq(await p.evaluate(()=>parityCalls),['next']);
  await p.evaluate(()=>document.querySelector('#toggleButton').disabled=true);await raf(p);eq(await p.locator('#toggleButton').getAttribute('data-resolved-access-key'),null);eq(await p.locator('#toggleButton .access-key-hint').getAttribute('data-access-caption'),'(N)');eq(await p.locator('#stableButton').getAttribute('data-resolved-access-key'),'p');
 });
 await test('스크롤 전후 키 안정성·항상 표기와 Alt 실행 일치',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'스크롤 키',body:el('div',{id:'keyScroller',style:{height:'100px',overflow:'auto'}},Array.from({length:18},(_,i)=>el('button',{id:'scrollKey'+i,style:{display:'block',height:'30px'},onclick:()=>parityCalls.push(i)},'합성 명령 '+i)))}));await raf(p);
  const keys=await p.locator('#keyScroller button').evaluateAll(ns=>ns.map(n=>n.dataset.resolvedAccessKey));ok(keys.every(Boolean));await p.locator('#scrollKey17').scrollIntoViewIfNeeded();await raf(p);
  eq(await p.locator('#keyScroller button').evaluateAll(ns=>ns.map(n=>n.dataset.resolvedAccessKey)),keys);eq(await p.locator('#scrollKey17 .access-key-hint').getAttribute('data-access-caption'),'('+keys[17].toUpperCase()+')');await p.keyboard.press('Alt+'+keys[17]);eq(await p.evaluate(()=>parityCalls),[17]);
 });
 await test('동일 명시 키는 순환·동적 표시 뒤 표기 중복/관찰자 반복 없음',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'순환 키',body:el('div',{id:'cycleBody'},el('button',{accessKey:'a',onclick:()=>parityCalls.push('one')},'첫째'),el('button',{accessKey:'a',onclick:()=>parityCalls.push('two')},'둘째'),el('button',{id:'showLater',hidden:true,accessKey:'b'},'나중'))}));await raf(p);
  await p.keyboard.press('Alt+a');eq(await p.evaluate(()=>document.activeElement.textContent),'첫째');await p.keyboard.press('Alt+a');eq(await p.evaluate(()=>document.activeElement.textContent),'둘째');await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['two']);
  await p.evaluate(()=>document.querySelector('#showLater').hidden=false);await raf(p);eq(await p.locator('#showLater .access-key-hint').count(),1);
  const mutations=await p.evaluate(()=>new Promise(resolve=>{let count=0;const observer=new MutationObserver(rs=>count+=rs.length);observer.observe(document.querySelector('.dialog'),{subtree:true,childList:true,attributes:true});setTimeout(()=>{observer.disconnect();resolve(count);},150);}));eq(mutations,0,'정지 상태에서 힌트 갱신 무한반복 없음');
 });
 await test('본문 닫기가 스크롤 밖이면 헤더 닫기 Alt D 유지',async p=>{
  await ui(p,({openDialog,el})=>{window.parityDialog=openDialog({title:'스크롤 닫기',body:el('div',{style:{height:'100px',overflow:'auto'}},el('button',{onclick:()=>parityDialog.close()},'닫기'),el('div',{style:{height:'300px'}}),el('input',{id:'scrollCloseEnd'}))});});
  await p.locator('#scrollCloseEnd').scrollIntoViewIfNeeded();await p.keyboard.press('Alt');eq(await p.locator('.dialog-close').getAttribute('data-resolved-access-key'),'d');await p.keyboard.press('d');eq(await p.getByRole('dialog').count(),0);
 });
 await test('비활성 기본 버튼은 Enter로 실행되지 않음',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'기본 버튼 비활성',body:el('input'),buttons:[{label:'확인',primary:true,action:()=>parityCalls.push('ok')}],onOpen:d=>d.querySelector('.btn.primary').disabled=true}));
  await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),[]);eq(await p.getByRole('dialog').count(),1);eq(await p.evaluate(()=>gridKeys),0);
  await p.evaluate(()=>document.querySelector('.btn.primary').disabled=false);await p.keyboard.press('Enter');eq(await p.evaluate(()=>parityCalls),['ok']);eq(await p.getByRole('dialog').count(),0);
 });
} finally { await browser.close(); }
console.log(JSON.stringify({url,total:results.length,passed:results.filter(r=>r.ok).length,failed:results.filter(r=>!r.ok).length,assertions,results},null,2));
if(results.some(r=>!r.ok))process.exitCode=1;
