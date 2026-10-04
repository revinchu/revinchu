// iPad 입력 경로 회귀: 개체 선택은 입력기를 깨우지 않는다. 실제 iPadOS 키보드 표시는 별도 기기 확인이 필요하다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
const engines=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'webkit',url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
const out=process.env.WIXEL_SLICER_FOCUS_OUT||'D:/Codex/Temp/wixel-slicer-input-focus/'+engine;
await mkdir(out,{recursive:true});
const browser=await engines[engine].launch(),results=[],errors=[],writes=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);};
const ok=(a,m)=>{checks++;assert.ok(a,m);};
const frame=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const object=(p,id)=>p.locator('.obj[data-id="'+id+'"]:visible').first();
const item=(p,id,label)=>object(p,id).locator('.sl-item').filter({hasText:new RegExp('^'+label+'$')}).first();
const state=p=>p.evaluate(()=>({focus:{tag:document.activeElement?.tagName,id:document.activeElement?.id,editable:!!document.activeElement?.matches('input:not([type=button]):not([type=checkbox]):not([type=radio]),textarea,[contenteditable=true]')},trace:window.__focusTrace,scroll:{x:tabula.gv().scroll.scrollLeft,y:tabula.gv().scroll.scrollTop,windowX:window.scrollX,windowY:window.scrollY},table:tabula.wb().sheets[0].tables[0].filter?.criteria?.[0]??null,pivot:tabula.wb().sheets[1].pivot.filters?.지역??null,active:{...tabula.active},cell:tabula.wb().getRaw(0,250,3),undo:tabula.wb().undoStack.length}));
async function fixture(p){await p.evaluate(()=>{const w=tabula.wb();w.restore({sheets:[{name:'자료',cells:{'0,0':{raw:'지역'},'0,1':{raw:'매출'},'1,0':{raw:'가'},'1,1':{raw:'10'},'2,0':{raw:'나'},'2,1':{raw:'20'},'3,0':{raw:'다'},'3,1':{raw:'30'},'0,2':{raw:'날짜'},'1,2':{raw:'2026-10-01'},'2,2':{raw:'2026-11-01'},'3,2':{raw:'2026-12-01'},'250,3':{raw:'먼 곳 선택 셀'}},tables:[{id:'t1',name:'Table1',r1:0,c1:0,r2:3,c2:2,header:true,filter:{criteria:{0:['가']},hidden:{2:true,3:true}}}],slicers:[{id:'table-sl',caption:'표 지역',source:{kind:'table',table:'Table1',column:'지역'},x:330,y:40,w:210,h:230},{id:'pivot-sl',caption:'피벗 지역',source:{kind:'pivot',field:'지역',pivots:[{sheet:'보고서',name:'피벗1'}]},x:600,y:40,w:210,h:230},{id:'timeline',caption:'날짜',timeline:true,source:{kind:'table',table:'Table1',column:'날짜'},x:330,y:310,w:400,h:140}],shapes:[{id:'rect',kind:'rect',x:870,y:50,w:100,h:100,fill:'#00aabb'}]},{name:'보고서',cells:{},pivot:{name:'피벗1',source:'자료',range:{r1:0,c1:0,r2:3,c2:1},rows:['지역'],cols:[],pages:[],values:[{field:'매출',name:'총매출'}],filters:{지역:['가']},top:0,left:0,area:{r1:0,c1:0,r2:4,c2:1}}}]});tabula.switchSheet(0);tabula.selectCell(250,3);tabula.gv().layout();tabula.gv().setScroll(0,0);tabula.gv().renderAll();document.getElementById('gridView').tabIndex=-1;document.getElementById('gridView').focus({preventScroll:true});w.undoStack=[];w.redoStack=[];window.__focusTrace=[];});await frame(p);}
async function action(p,label,fn,expected){await p.evaluate(()=>window.__focusTrace=[]);const before=await state(p);await fn();await frame(p);const after=await state(p);const failures=[];for(const [message,check] of [['입력 요소에 초점을 보내지 않음',()=>{eq(after.focus.editable,false);eq(after.trace.filter(x=>x.editable),[]);} ],['격자와 문서 스크롤 유지',()=>eq(after.scroll,before.scroll)],['선택 셀 내용 유지',()=>eq(after.cell,before.cell)],['필터 적용',()=>expected?.(after)]]){try{check();}catch(e){failures.push(message+': '+e.message);}}return {label,ok:!failures.length,failures,before,after};}
try{
for(const mobile of [true,false])for(const preference of ['auto','hardware','screen'])for(const input of ['mouse','touch']){
 const name=`iPad ${mobile?'모바일 ON':'모바일 OFF'} ${preference} ${input}`;if(process.env.WIXEL_SLICER_FOCUS_FILTER&&!name.includes(process.env.WIXEL_SLICER_FOCUS_FILTER))continue;
 const c=await browser.newContext({viewport:{width:1100,height:900},hasTouch:true,isMobile:true,serviceWorkers:'block',userAgent:'Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Version/18.0 Mobile/15E148 Safari/604.1'}),p=await c.newPage(),pe=[],steps=[];
 p.setDefaultTimeout(10000);p.on('pageerror',e=>{pe.push(e.message);errors.push({name,message:e.message});});
 await c.addInitScript(({mobile,preference})=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');localStorage.setItem('wixel.mobile-work.v1',mobile?'on':'off');localStorage.setItem('wixel.mobile-keyboard.v1',preference);Object.defineProperty(navigator,'platform',{value:'MacIntel'});Object.defineProperty(navigator,'maxTouchPoints',{value:5});Object.defineProperty(navigator,'clipboard',{configurable:true,value:{writeText:async()=>{},readText:async()=>''}});window.__focusTrace=[];const f=HTMLElement.prototype.focus;HTMLElement.prototype.focus=function(...args){window.__focusTrace.push({tag:this.tagName,id:this.id,mode:this.inputMode,editable:this.matches('input:not([type=button]):not([type=checkbox]):not([type=radio]),textarea,[contenteditable=true]'),preventScroll:args[0]?.preventScroll===true});return f.apply(this,args);};},{mobile,preference});
 await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return r.abort();return r.continue();});
 const press=async locator=>input==='touch'?locator.tap():locator.click();
 try{
  await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>!!window.tabula?.wb());await fixture(p);
  await p.locator('.font-box .format-picker-toggle').click();await frame(p);
  for(const id of ['table-sl','pivot-sl']){
   steps.push(await action(p,id+' 머리글 선택',()=>press(object(p,id).locator('.sl-head'))));
   steps.push(await action(p,id+' 항목 선택',()=>press(item(p,id,'나')),s=>eq(s[id==='table-sl'?'table':'pivot'],['나'])));
   steps.push(await action(p,id+' 다중 선택 켜기',()=>press(object(p,id).locator('.sl-multi'))));
   steps.push(await action(p,id+' 항목 추가',()=>press(item(p,id,'다')),s=>eq(s[id==='table-sl'?'table':'pivot'],['나','다'])));
   steps.push(await action(p,id+' 필터 지우기',()=>press(object(p,id).locator('.sl-clear')),s=>eq(s[id==='table-sl'?'table':'pivot'],null)));
  }
  steps.push(await action(p,'시간 표시 막대 기간 선택',()=>press(object(p,'timeline').locator('.tl-cell').first())));
  steps.push(await action(p,'시간 표시 막대 수준 메뉴',()=>press(object(p,'timeline').locator('.tl-level'))));
  steps.push(await action(p,'시간 표시 막대 수준 메뉴 닫기',()=>p.keyboard.press('Escape')));
  steps.push(await action(p,'슬라이서 항목 두 번 클릭',()=>item(p,'table-sl','가').dblclick()));
  eq(await p.locator('[role=dialog]').count(),0,'항목 두 번 클릭은 설정을 열지 않음');
  await object(p,'table-sl').locator('.sl-head').dblclick();await p.getByRole('dialog',{name:'슬라이서 설정',exact:true}).waitFor();await frame(p);
  steps.push(await action(p,'머리글 두 번 클릭 설정 닫기',()=>p.keyboard.press('Escape')));
  steps.push(await action(p,'슬라이서 우클릭 메뉴 열기',()=>object(p,'table-sl').locator('.sl-head').click({button:'right'})));
  await p.getByRole('menuitem',{name:/^슬라이서 설정/}).waitFor();
  steps.push(await action(p,'슬라이서 메뉴 Escape 닫기',()=>p.keyboard.press('Escape')));
  await p.evaluate(()=>tabula.run('slicerSettings'));await p.getByRole('dialog',{name:'슬라이서 설정',exact:true}).waitFor();await frame(p);
  steps.push(await action(p,'슬라이서 설정 취소',()=>press(p.getByRole('dialog',{name:'슬라이서 설정',exact:true}).getByRole('button',{name:'취소',exact:true}))));
  const oldX=await p.evaluate(()=>tabula.wb().sheets[0].slicers.find(s=>s.id==='table-sl').x);
  steps.push(await action(p,'슬라이서 첫 물리 화살표 키',()=>p.keyboard.press('ArrowRight')));
  eq(await p.evaluate(()=>tabula.wb().sheets[0].slicers.find(s=>s.id==='table-sl').x),oldX+1,'슬라이서 키보드 이동');
  steps.push(await action(p,'슬라이서 복사·붙여넣기',async()=>{await p.keyboard.press('Control+c');await p.evaluate(()=>{const data=new DataTransfer();data.setData('text/plain','');document.activeElement.dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));});}));
  eq(await p.evaluate(()=>tabula.wb().sheets[0].slicers.length),4,'키보드 복사·붙여넣기 이벤트 유지');
  await p.keyboard.press('Delete');await frame(p);
  eq(await p.evaluate(()=>tabula.wb().sheets[0].slicers.length),3,'붙인 개체만 삭제');
  await press(object(p,'rect'));await frame(p);
  const shapeEditor=!(mobile&&preference==='screen');
  eq(await p.evaluate(()=>document.activeElement.id),shapeEditor?'cellEditor':'gridView','도형은 기존 즉시 입력 정책 유지');
  if(shapeEditor){
   await p.evaluate(()=>{const ed=document.getElementById('cellEditor');ed.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:''}));ed.value='한글';ed.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertCompositionText',data:'한글',isComposing:true}));ed.dispatchEvent(new CompositionEvent('compositionend',{bubbles:true,data:'한글'}));});
   const shapeDialog=p.getByRole('dialog',{name:'도형 서식',exact:true});await shapeDialog.waitFor();
   eq(await p.evaluate(()=>tabula.wb().sheets[0].shapes.find(s=>s.id==='rect').text),'한글','한글 조합 종료 후 도형 텍스트 보존');
   eq(await shapeDialog.locator('textarea').inputValue(),'한글','도형 편집 창에서 조합 결과 유지');
   await shapeDialog.getByRole('button',{name:'닫기',exact:true}).click();await frame(p);
  }
  // 필터로 이동한 개체가 격자 좌표를 가릴 수 있으므로 실제 빈 셀을 hit-test 합니다.
  const point=await p.evaluate(()=>{for(const r of [16,18,20,12,10])for(const c of [1,2,11,12,13,0]){const rect=tabula.gv().clientRect({r1:r,c1:c,r2:r,c2:c}),x=rect.left+rect.width/2,y=rect.top+rect.height/2,node=document.elementFromPoint(x,y),hit=tabula.gv().hitTest(x,y);if(node?.closest('#gridView')&&!node.closest('.obj')&&hit.zone==='cell'&&hit.r===r&&hit.c===c&&tabula.wb().getRaw(0,r,c)==='')return{x,y,r,c};}throw Error('검사용 빈 격자 칸이 보이지 않습니다.');});
  await p.mouse.click(point.x,point.y);eq(await p.evaluate(()=>tabula.chartSel),null,'실제 셀 클릭으로 개체 선택 해제');
  await p.keyboard.press('F2');eq(await p.locator('#cellEditor').evaluate(n=>n.classList.contains('idle')),false,'셀 편집 시작');await p.locator('#cellEditor').fill('137');await frame(p);
  steps.push(await action(p,'셀 편집 확정 후 슬라이서 항목 선택',()=>press(item(p,'table-sl','나'))));
  eq(await p.evaluate(({r,c})=>tabula.wb().getRaw(0,r,c),point),'137','슬라이서 전환에서 진행 중 셀 편집 보존');
  await p.keyboard.press('Escape');eq(await p.evaluate(()=>tabula.chartSel),null,'Escape 개체 선택 해제');if(mobile&&preference==='screen'){await p.keyboard.press('F2');await p.locator('#cellEditor').fill('');}await p.keyboard.type('42');await p.keyboard.press('Enter');
  eq(await p.evaluate(({r,c})=>tabula.wb().getRaw(0,r,c),point),'42','선택 해제 후 첫 셀 문자 보존');
  await press(object(p,'rect'));await p.keyboard.press('Delete');eq(await p.evaluate(()=>tabula.wb().sheets[0].shapes.length),0,'도형 삭제');
  const target=await p.evaluate(()=>({...tabula.active}));if(mobile&&preference==='screen')await p.keyboard.press('F2');await p.keyboard.type('49');await p.keyboard.press('Enter');
  eq(await p.evaluate(({r,c})=>tabula.wb().getRaw(0,r,c),target),'49','개체 삭제 후 첫 셀 문자 보존');
  eq(pe,[],'페이지 오류 없음');ok(steps.every(s=>s.ok),'모든 선택에서 비입력 포커스·스크롤·필터 유지');
  results.push({name,ok:true,steps});console.log('OK '+name);
 }catch(e){results.push({name,ok:false,error:e.message,steps,state:await state(p).catch(()=>null)});console.error('NG '+name+': '+e.message);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}
 finally{await c.close();}
}
}finally{await browser.close();const summary={url,engine,cases:results.length,passed:results.filter(r=>r.ok).length,checks,pageErrors:errors,remoteWrites:writes,scope:'Synthetic table and pivot slicers; Playwright trusted mouse/touch with iPad UA. Every focus call and browser/grid scroll observed. Shape IME uses a synthetic composition sequence to check the existing native editor route. Paste uses an explicitly dispatched ClipboardEvent with synthetic data; copy and arrow/typing use trusted keyboard events. Actual iPadOS virtual keyboard presentation and Bluetooth device behavior are not measured.',results};await writeFile(out+'/result.json',JSON.stringify(summary,null,2));console.log(JSON.stringify({...summary,results:results.map(({name,ok,error,steps})=>({name,ok,error,steps:steps.map(({label,ok,failures})=>({label,ok,failures}))}))}));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;}
