// 합성 피벗의 실제 마우스 및 모의 Pointer/Touch 이동 검사. 물리 iPhone 검사가 아닙니다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const repo=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');

const engines=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const browserName=process.env.WIXEL_BROWSER||'chromium';
assert.ok(['chromium','webkit'].includes(browserName));
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
const out=process.env.WIXEL_PIVOT_DRAG_OUT||'D:/Codex/Temp/wixel-pivot-field-drag/source';
const only=process.env.WIXEL_PIVOT_DRAG_FILTER||'';

const parsed=new URL(url),local=['localhost','127.0.0.1','[::1]'].includes(parsed.hostname);
if(!local&&process.env.WIXEL_ALLOWED_TEST_URL!==url)throw new Error('Public synthetic tests require an exact WIXEL_ALLOWED_TEST_URL.');
if(!/^[dD]:/.test(path.resolve(out)))throw new Error('Test outputs must be on D:.');
const baseline=process.env.WIXEL_PIVOT_DRAG_BASELINE||'',baselineSources=new Map();
if(baseline) {
  if(!local)throw new Error('Git baseline tests require a local source server.');
  for(const name of ['src/app.js','src/pivot-field-drag.js','styles.css'])baselineSources.set('/'+name,
    execFileSync('git',['-c','safe.directory='+path.resolve(repo,'..').replaceAll(String.fromCharCode(92),'/'),'-C',repo,'show',baseline+':tabula/'+name],{encoding:'utf8',maxBuffer:32*1024*1024}));
}

await mkdir(out,{recursive:true});
const browser=await engines[browserName].launch(),results=[],skipped=[],errors=[],writes=[],blocked=[],assets=new Set();let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(a,m)=>{checks++;assert.ok(a,m);};
const raf=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const run=(p,id)=>p.evaluate(id=>window.tabula.run(id),id);
const field=(p,name)=>p.locator('#pivotPane .pp-field').filter({has:p.locator('[data-pivot-field="'+name+'"]')}).locator('.pp-fname');
const area=(p,name)=>p.locator('#pivotPane [data-pivot-area="'+name+'"]');
const definition=p=>p.evaluate(()=>{const t=window.tabula,d=t.wb().sheets[t.si].pivot;return{rows:d.rows||[],cols:d.cols||[],pages:d.pages||[],values:d.values||[],valuesOnRows:!!d.valuesOnRows,valuesPos:d.valuesPos??null};});
const model=p=>p.evaluate(()=>{const t=window.tabula,w=t.wb();return{source:JSON.stringify([...w.sheets[0].cells.entries()]),undo:w.undoStack.length,redo:w.redoStack.length};});
async function fixture(p){await p.evaluate(()=>{const t=window.tabula,w=t.wb(),cells={};[['지역','채널','상품','연도','매출','비용','수량'],['서울','온라인','A',2025,100,40,1],['부산','매장','B',2026,200,80,2],['서울','매장','B',2026,300,120,3],['대구','온라인','A',2025,400,160,4]].forEach((row,r)=>row.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));w.restore({sheets:[{name:'원본 합성',cells}]});t.switchSheet(0);t.gv().setZoom(100);t.gv().renderAll();t.selectRange({r1:0,c1:0,r2:4,c2:6});t.run('insertPivot');});await p.getByRole('dialog',{name:'피벗 테이블 만들기',exact:true}).getByRole('button',{name:'확인',exact:true}).click();await p.locator('#pivotPane .pp-fields').waitFor();await p.evaluate(()=>{const w=window.tabula.wb();w.undoStack=[];w.redoStack=[];window.__pivotDragTrace=[];for(const type of ['pointerdown','pointermove','pointerup','pointercancel','gotpointercapture','lostpointercapture','dragstart','dragover','drop','dragend'])document.addEventListener(type,e=>{if(type==='pointermove'&&!e.buttons)return;window.__pivotDragTrace.push({type,pointerType:e.pointerType,buttons:e.buttons,tag:e.target?.className,area:e.target?.closest?.('[data-pivot-area]')?.dataset.pivotArea});},true);});await raf(p);}
async function mouseDrag(p,source,target,{position='end'}={}){await source.scrollIntoViewIfNeeded();const a=await source.boundingBox();await target.scrollIntoViewIfNeeded();const b=await target.boundingBox();ok(a&&b,'드래그 원본과 대상이 보임');const x=a.x+Math.min(30,a.width/2),y=a.y+a.height/2,tx=b.x+b.width/2,ty=position==='start'?b.y+3:position==='middle'?b.y+b.height/2:b.y+b.height-5;await p.mouse.move(x,y);await p.mouse.down();await p.mouse.move(x+9,y+2,{steps:3});await p.mouse.move(tx,ty,{steps:12});await p.mouse.move(tx,ty+1);await p.mouse.up();await raf(p);}
async function pointerDrag(p,source,target,{pointerType='touch',position='end',finish=true}={}){await source.scrollIntoViewIfNeeded();await target.scrollIntoViewIfNeeded();const a=await source.boundingBox(),b=await target.boundingBox();ok(a&&b,'Pointer 원본과 대상이 보임');const from={x:a.x+Math.min(30,a.width/2),y:a.y+a.height/2},to={x:b.x+b.width/2,y:position==='start'?b.y+3:b.y+b.height-5};await source.evaluate((node,{from})=>{window.__pivotPointerSource=node;node.dispatchEvent(new PointerEvent('pointerdown',{bubbles:true,cancelable:true,pointerType:'touch',pointerId:51,isPrimary:true,button:0,buttons:1,clientX:from.x,clientY:from.y}));},{from});for(let i=1;i<=12;i++)await p.evaluate(({from,to,i,pointerType})=>window.__pivotPointerSource.dispatchEvent(new PointerEvent('pointermove',{bubbles:true,cancelable:true,pointerType,pointerId:51,isPrimary:true,button:-1,buttons:1,clientX:from.x+(to.x-from.x)*i/12,clientY:from.y+(to.y-from.y)*i/12})),{from,to,i,pointerType});if(finish)await p.evaluate(({to,pointerType})=>window.__pivotPointerSource.dispatchEvent(new PointerEvent('pointerup',{bubbles:true,cancelable:true,pointerType,pointerId:51,isPrimary:true,button:0,buttons:0,clientX:to.x,clientY:to.y})),{to,pointerType});await raf(p);return to;}
async function add(p,name,dest){await field(p,name).click({button:'right'});await p.getByRole('menuitem',{name:{pages:'보고서 필터에 추가',rows:'행 레이블에 추가',cols:'열 레이블에 추가',values:'값에 추가'}[dest],exact:true}).click();}
const item=(p,dest,name)=>area(p,dest).locator('.pp-item').filter({has:p.locator('.pp-label').filter({hasText:new RegExp('^'+name+'$')})});
async function pointerUp(p,to,type='pointerup'){await p.evaluate(({to,type})=>(window.__pivotPointerSource?.isConnected?window.__pivotPointerSource:document).dispatchEvent(new PointerEvent(type,{bubbles:true,cancelable:true,pointerType:'touch',pointerId:51,isPrimary:true,button:0,buttons:0,clientX:to.x,clientY:to.y})),{to,type});await raf(p);}
async function configure(p,patch){await p.evaluate(patch=>{const t=window.tabula,w=t.wb(),d=w.sheets[t.si].pivot;w.transact(()=>w.setSheetProp(t.si,'pivot',{...d,...patch}));t.run('pivotRefresh');w.undoStack=[];w.redoStack=[];},patch);await raf(p);}
async function changedOnce(p,before,original){const after=await model(p);eq(after.source,before.source,'원본 셀 불변');eq(after.undo,before.undo+1,'한 번의 드래그 한 번의 Undo');await run(p,'undo');eq(await definition(p),original,'Undo 배치 복원');eq((await model(p)).source,before.source);}
async function unchanged(p,before,original){eq(await definition(p),original,'피벗 정의 불변');eq(await model(p),before,'원본과 Undo/Redo 불변');eq(await p.locator('.pivot-field-ghost,.pivot-field-marker').count(),0,'드래그 보조 요소 정리');}

const fields = p => p.locator('#pivotPane .pp-fields');
const fullBook = p => p.evaluate(() => tabula.wb().serialize().sheets);
const checkField = (p, name) => p.locator('#pivotPane [data-pivot-field="'+name+'"]');
const removalLayout = { rows:['지역'], cols:['채널'], pages:['상품'], values:[{field:'매출',agg:'sum',name:'총매출'}], layout:'tabular', autofit:false, autoRefresh:false };
async function nativeDrag(p,source,target,{finish=true}={}) {
  await source.scrollIntoViewIfNeeded();await target.scrollIntoViewIfNeeded();
  const a=await source.boundingBox(),b=await target.boundingBox();ok(a&&b,'HTML drag 원본과 대상이 보임');
  const from={x:a.x+a.width/2,y:a.y+a.height/2},to={x:b.x+b.width/2,y:b.y+b.height/2};
  await source.evaluate((node,from)=>{const row=node.closest('.pp-item,.pp-field')??node;
    window.__pivotNativeSource=row;window.__pivotNativeTransfer=new DataTransfer();
    row.dispatchEvent(new DragEvent('dragstart',{bubbles:true,cancelable:true,dataTransfer:window.__pivotNativeTransfer,clientX:from.x,clientY:from.y}));
  },from);
  await target.evaluate((node,to)=>{const dt=window.__pivotNativeTransfer;
    node.dispatchEvent(new DragEvent('dragenter',{bubbles:true,cancelable:true,dataTransfer:dt,clientX:to.x,clientY:to.y}));
    const over=new DragEvent('dragover',{bubbles:true,cancelable:true,dataTransfer:dt,clientX:to.x,clientY:to.y});
    node.dispatchEvent(over);window.__pivotNativeAccepted=over.defaultPrevented;
  },to);
  if(finish) {
    await target.evaluate((node,to)=>node.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true,dataTransfer:window.__pivotNativeTransfer,clientX:to.x,clientY:to.y})),to);
    await p.evaluate(()=>window.__pivotNativeSource.dispatchEvent(new DragEvent('dragend',{bubbles:true,cancelable:true,dataTransfer:window.__pivotNativeTransfer})));
  }
  await raf(p);return to;
}
async function assertRemoval(p,before,bookBefore,expected,fieldName,checked=false) {
  eq(await definition(p),expected,'필드 목록 drop은 선택한 배치 슬롯만 제거');
  eq(await checkField(p,fieldName).isChecked(),checked,'목록 체크는 남은 배치 사용 여부와 일치');
  const after=await model(p),bookAfter=await fullBook(p);eq(after.source,before.source,'목록 제거가 원본 데이터에 영향을 주지 않음');
  eq(after.undo,before.undo+1,'목록 제거는 Undo 한 번');eq(after.redo,0,'새 제거는 이전 Redo를 비움');
  eq(await p.locator('.pivot-field-ghost,.pivot-field-marker').count(),0,'제거 후 드래그 보조 요소 정리');
  await run(p,'undo');await raf(p);eq(await fullBook(p),bookBefore,'Undo는 피벗 셀·서식·너비·배치를 함께 복원');
  await run(p,'redo');await raf(p);eq(await fullBook(p),bookAfter,'Redo는 같은 슬롯 제거를 정확히 복원');
  eq(await definition(p),expected,'Redo 뒤 배치 유지');eq(await checkField(p,fieldName).isChecked(),checked,'Redo 뒤 목록 체크 유지');
}
const clearTrace=p=>p.evaluate(()=>window.__pivotDragTrace=[]);

async function test(name,fn,{cdpOnly=false}={}){if(only&&!name.includes(only))return;if(cdpOnly&&browserName!=='chromium'){skipped.push({name,reason:'Chromium CDP touch injection only'});console.log('SKIP '+name);return;}const context=await browser.newContext({viewport:{width:1440,height:1000},hasTouch:true,serviceWorkers:'block'}),p=await context.newPage(),pe=[],pw=[];p.setDefaultTimeout(10000);p.on('pageerror',e=>{pe.push(e.message);errors.push({name,error:e.message});});p.on('dialog',d=>d.dismiss());await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');localStorage.setItem('wixel.mobile-work.v1','off');Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>'',writeText:async()=>{}}});});await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method()+' '+u.pathname);pw.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')){blocked.push(u.pathname);return r.abort();}if(baselineSources.has(u.pathname))return r.fulfill({status:200,contentType:u.pathname.endsWith('.css')?'text/css; charset=utf-8':'text/javascript; charset=utf-8',body:baselineSources.get(u.pathname)});return r.continue();});try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());for(const s of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(s);await fixture(p);await fn(p,context);eq(pe,[]);eq(pw,[]);results.push({name,ok:true,trace:await p.evaluate(()=>window.__pivotDragTrace)});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.message,trace:await p.evaluate(()=>window.__pivotDragTrace).catch(()=>[]),definition:await definition(p).catch(()=>null)});console.error('NG '+name+': '+e.message.split('\n')[0]);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}finally{await context.close();}}
try{

for(const method of ['mouse','touch','native'])for(const [dest,name] of [['pages','상품'],['rows','지역'],['cols','채널'],['values','매출']])
await test('목록으로 제거 '+method+' '+dest,async p=>{
  await configure(p,removalLayout);await clearTrace(p);
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  const expected={...original,[dest]:[]},source=area(p,dest).locator('.pp-item:not(.sigma) .pp-label').first();
  if(method==='mouse')await mouseDrag(p,source,fields(p),{position:'middle'});
  else if(method==='touch')await pointerDrag(p,source,fields(p),{position:'middle'});
  else await nativeDrag(p,source,fields(p));
  await assertRemoval(p,before,bookBefore,expected,name);
  const trace=await p.evaluate(()=>window.__pivotDragTrace);
  if(method==='native'){ok(trace.some(e=>e.type==='dragstart')&&trace.some(e=>e.type==='drop'),'HTML dragstart/drop fallback 경로 실행');eq(trace.some(e=>e.type==='pointerdown'),false,'HTML 검사는 Pointer 세션과 분리');}
  else ok(trace.some(e=>e.type==='pointerdown')&&trace.some(e=>e.type==='pointerup'),'실제 마우스 또는 Touch Pointer 경로 실행');
});

for(const method of ['mouse','native'])await test('목록으로 제거 '+method+' 중복 값 중 선택한 한 개 제거',async p=>{
  const values=[{field:'매출',agg:'sum',name:'총매출',numFmt:{numFmt:'currency',decimals:0}},{field:'매출',agg:'average',name:'평균매출',numFmt:{numFmt:'number',decimals:2}},{field:'비용',agg:'max',name:'최대비용'}];
  await configure(p,{rows:['상품'],values,autofit:false,autoRefresh:false});await clearTrace(p);
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  const source=area(p,'values').locator('.pp-item:not(.sigma) .pp-label').nth(1);
  if(method==='mouse')await mouseDrag(p,source,fields(p),{position:'middle'});else await nativeDrag(p,source,fields(p));
  await assertRemoval(p,before,bookBefore,{...original,values:[values[0],values[2]]},'매출',true);
  eq(await checkField(p,'비용').isChecked(),true,'다른 남은 값 필드도 체크 유지');
});

await test('목록으로 제거 다른 영역에 같은 필드가 남으면 체크 유지',async p=>{
  await configure(p,{rows:['매출'],values:[{field:'매출',agg:'sum',name:'총매출'}],autofit:false,autoRefresh:false});
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,item(p,'rows','매출').locator('.pp-label'),fields(p),{position:'middle'});
  await assertRemoval(p,before,bookBefore,{...original,rows:[]},'매출',true);
});

await test('목록으로 제거 체크박스 위 drop은 클릭으로 재추가하지 않음',async p=>{
  await configure(p,{rows:['지역'],values:[{field:'매출',agg:'sum'}],autofit:false,autoRefresh:false});
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,item(p,'rows','지역').locator('.pp-label'),checkField(p,'지역'),{position:'middle'});
  await assertRemoval(p,before,bookBefore,{...original,rows:[]},'지역');
});

await test('목록으로 제거 안내와 Escape 취소는 배치 유지',async p=>{
  await configure(p,removalLayout);
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  const to=await pointerDrag(p,item(p,'rows','지역').locator('.pp-label'),fields(p),{position:'middle',finish:false});
  eq(await fields(p).evaluate(n=>n.classList.contains('pivot-field-drop-target')),true,'목록 제거 대상 강조');
  ok((await p.locator('.pivot-field-ghost').textContent()).includes('필드 제거'),'ghost에 제거 동작 안내');
  eq(await p.locator('.pivot-field-marker').isVisible(),false,'제거 대상에는 삽입선이 없음');
  await p.screenshot({path:out+'/field-list-remove-hover.png'});
  await p.keyboard.press('Escape');await pointerUp(p,to);await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'Escape는 전체 피벗 상태 유지');
});

await test('목록으로 제거 영역 밖 drop은 배치 유지',async p=>{
  await configure(p,removalLayout);const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await pointerDrag(p,item(p,'rows','지역').locator('.pp-label'),p.locator('.formula-row'),{position:'middle'});
  await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'목록 밖 drop은 전체 셀 상태 유지');
});

await test('목록으로 제거 드래그 중 보호되면 제거하지 않음',async p=>{
  await configure(p,removalLayout);const original=await definition(p);
  const to=await pointerDrag(p,item(p,'rows','지역').locator('.pp-label'),fields(p),{position:'middle',finish:false});
  await p.evaluate(()=>{const t=tabula;t.wb().sheets[t.si].protect={on:true,allow:{selectLocked:true}};});
  const before=await model(p),bookBefore=await fullBook(p);await pointerUp(p,to);await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'보호가 적용된 뒤 전체 상태 유지');
});

await test('목록으로 제거 드래그 중 버전이 바뀌면 제거하지 않음',async p=>{
  await configure(p,removalLayout);const original=await definition(p);
  const to=await pointerDrag(p,item(p,'rows','지역').locator('.pp-label'),fields(p),{position:'middle',finish:false});
  await p.evaluate(()=>{const t=tabula,w=t.wb();w.transact(()=>w.setInput(t.si,30,10,'드래그 중 별도 편집'));});
  const before=await model(p),bookBefore=await fullBook(p);await pointerUp(p,to);await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'지연 드롭은 새 편집과 Undo를 보존');
});

for(const method of ['mouse','native'])await test('목록으로 제거 '+method+' Σ 값은 배치 유지',async p=>{
  await configure(p,{rows:['상품'],cols:['채널'],values:[{field:'매출',agg:'sum'},{field:'비용',agg:'sum'}],autofit:false,autoRefresh:false});
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p),source=area(p,'cols').locator('.sigma .pp-label');
  if(method==='mouse')await mouseDrag(p,source,fields(p),{position:'middle'});else await nativeDrag(p,source,fields(p));
  await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'Σ 값 drop으로 값 배열을 비우지 않음');
});

await test('목록으로 제거 미배치 필드를 목록에 놓으면 배치 유지',async p=>{
  await configure(p,removalLayout);const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,field(p,'연도'),fields(p),{position:'middle'});await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'미배치 필드 목록 drop은 Undo 없이 유지');
});

await test('목록으로 제거 인접 피벗 충돌은 두 피벗과 Undo 보존',async p=>{
  await configure(p,{rows:['지역'],pages:['상품'],cols:[],values:[{field:'매출',agg:'sum'}],layout:'tabular',autofit:false,autoRefresh:false,subtotals:'none'});
  await p.evaluate(()=>{const t=tabula,w=t.wb(),d=w.sheets[t.si].pivot;
    const other={name:'인접 피벗',source:d.source,range:structuredClone(d.range),top:d.area.r2+1,left:d.left??0,
      rows:['상품'],cols:[],values:[{field:'비용',agg:'sum'}],layout:'tabular',autofit:false,autoRefresh:false};
    w.transact(()=>w.setSheetProp(t.si,'pivotsExtra',[other]));t.run('pivotRefresh');
    w.undoStack=[];w.redoStack=[];w.transact(()=>w.setInput(0,1,0,'인천'));
  });await raf(p);
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,item(p,'pages','상품').locator('.pp-label'),fields(p),{position:'middle'});
  await unchanged(p,before,original);eq(await fullBook(p),bookBefore,'제거 후 새 원본 계산이 충돌하면 두 피벗 셀·정의·너비 모두 보존');
  ok(/겹|중첩|overlap/i.test(await p.locator('#toast').textContent()),'제거 거절 이유를 충돌 경고로 표시');
});

const selection=p=>p.evaluate(()=>({si:tabula.si,active:{...tabula.active},sel:{...tabula.sel}}));
const pivotArea=p=>p.evaluate(()=>({...tabula.wb().sheets[tabula.si].pivot.area}));
const inside=(a,pos)=>pos.r>=a.r1&&pos.r<=a.r2&&pos.c>=a.c1&&pos.c<=a.c2;

await test('목록으로 제거 오른쪽 값 셀을 없애도 Redo 초점과 필드 창 유지',async p=>{
  await configure(p,{rows:['상품'],cols:['채널'],pages:[],values:[{field:'매출',agg:'sum',name:'총매출'},{field:'비용',agg:'sum',name:'총비용'}],layout:'tabular',autofit:false,autoRefresh:false});
  const oldArea=await pivotArea(p),selected={r:oldArea.r2-1,c:oldArea.c2};
  await p.evaluate(pos=>tabula.selectCell(pos.r,pos.c),selected);await raf(p);
  eq(await p.locator('#pivotPane').isVisible(),true,'오른쪽 끝 값 셀 선택으로 필드 창이 열림');
  const original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,area(p,'values').locator('.pp-item:not(.sigma) .pp-label').nth(1),fields(p),{position:'middle'});
  const bookAfter=await fullBook(p),newArea=await pivotArea(p);
  ok(newArea.c2<oldArea.c2&&!inside(newArea,selected),'제거한 값 때문에 이전 활성 열이 실제 피벗 범위 밖으로 줄어듦');
  eq(await definition(p),{...original,values:[original.values[0]]},'오른쪽 값 슬롯 제거');
  eq((await model(p)).undo,before.undo+1,'선택 보정은 별도 Undo를 만들지 않음');
  await run(p,'undo');await raf(p);eq(await fullBook(p),bookBefore,'Undo는 전체 원래 피벗 복원');
  eq((await selection(p)).active,selected,'Undo에서 복원된 원래 오른쪽 값 셀 선택 유지');
  eq(await p.locator('#pivotPane').isVisible(),true,'Undo 뒤 필드 창 유지');
  await run(p,'redo');await raf(p);eq(await fullBook(p),bookAfter,'Redo는 동일한 값 제거 복원');
  const restored=await selection(p);eq(restored.active,{r:newArea.r1,c:newArea.c1},'Redo에서 사라진 셀 대신 현재 피벗 시작 셀 선택');
  eq(restored.sel,{r1:newArea.r1,c1:newArea.c1,r2:newArea.r1,c2:newArea.c1},'범위 밖 활성셀 보정은 단일 셀 선택');
  eq(await p.locator('#pivotPane').isVisible(),true,'Redo 뒤 필드 창이 닫히지 않음');
  eq(await checkField(p,'비용').isChecked(),false,'Redo 뒤 보이는 필드 창은 제거 체크 상태 반영');
  eq((await model(p)).source,before.source,'초점 복원은 원본 데이터 유지');
});

await test('목록으로 제거 범위 안 다중 셀 선택은 Undo Redo 뒤 그대로 유지',async p=>{
  await configure(p,{rows:['상품'],cols:[],pages:[],values:[{field:'매출',agg:'sum',name:'총매출'},{field:'비용',agg:'sum',name:'총비용'}],layout:'tabular',autofit:false,autoRefresh:false});
  const a=await pivotArea(p),rg={r1:a.r1+1,c1:a.c1,r2:a.r1+2,c2:a.c1+1},act={r:rg.r2,c:rg.c2};
  await p.evaluate(({rg,act})=>tabula.selectRange(rg,'cells',act),{rg,act});await raf(p);
  const selected=await selection(p),original=await definition(p),before=await model(p),bookBefore=await fullBook(p);
  await mouseDrag(p,area(p,'values').locator('.pp-item:not(.sigma) .pp-label').nth(1),fields(p),{position:'middle'});
  const next=await pivotArea(p),bookAfter=await fullBook(p);
  ok(inside(next,{r:rg.r1,c:rg.c1})&&inside(next,{r:rg.r2,c:rg.c2}),'다중 선택 전체가 축소된 피벗 안에 있음');
  eq(await definition(p),{...original,values:[original.values[0]]});
  eq(await p.evaluate(()=>tabula.wb().undoStack.at(-1).meta.selKind),'cells','피벗 작업은 원래 선택 종류 메타 유지');
  await run(p,'undo');await raf(p);eq(await fullBook(p),bookBefore);eq(await selection(p),selected,'Undo는 다중 선택 범위와 활성셀 그대로 유지');
  await run(p,'redo');await raf(p);eq(await fullBook(p),bookAfter);eq(await selection(p),selected,'Redo도 범위 안 다중 선택을 시작 셀로 바꾸지 않음');
  eq(await p.locator('#pivotPane').isVisible(),true,'범위 안 선택의 필드 창 유지');
  eq((await model(p)).undo,before.undo+1,'다중 선택 복원은 하나의 작업 이력');
});

await test('목록으로 제거 회귀 일반 F2 편집 Undo 선택은 피벗으로 보정하지 않음',async p=>{
  await configure(p,removalLayout);const def=await definition(p),outside={r:30,c:10};
  await p.evaluate(pos=>tabula.selectCell(pos.r,pos.c),outside);await raf(p);
  const selected=await selection(p),oldRaw=await p.evaluate(pos=>tabula.wb().getRaw(tabula.si,pos.r,pos.c),outside);
  eq(await p.locator('#pivotPane').isVisible(),false,'일반 셀 선택에서 필드 창은 닫힘');
  await p.locator('#cellEditor').focus();await p.keyboard.press('F2');await p.locator('#cellEditor').fill('일반 셀 편집 선택 보존');await p.keyboard.press('Enter');await raf(p);
  eq(await p.evaluate(pos=>tabula.wb().getRaw(tabula.si,pos.r,pos.c),outside),'일반 셀 편집 선택 보존','실제 F2 입력 확정');
  eq(await p.evaluate(()=>tabula.wb().undoStack.at(-1).meta.pivotFocus??null),null,'일반 셀 작업에 피벗 초점 힌트를 붙이지 않음');
  await run(p,'undo');await raf(p);eq(await selection(p),selected,'일반 편집 Undo는 원래 일반 셀 선택 복원');
  eq(await p.evaluate(pos=>tabula.wb().getRaw(tabula.si,pos.r,pos.c),outside),oldRaw);eq(await p.locator('#pivotPane').isVisible(),false);
  await run(p,'redo');await raf(p);eq(await selection(p),selected,'일반 편집 Redo도 피벗 시작 셀로 강제 이동하지 않음');
  eq(await p.evaluate(pos=>tabula.wb().getRaw(tabula.si,pos.r,pos.c),outside),'일반 셀 편집 선택 보존');eq(await definition(p),def,'일반 선택 보정 검사는 피벗 배치 불변');
});
for(const [dest,name] of [['pages','지역'],['rows','상품'],['cols','채널'],['values','매출']])await test('마우스로 필드 목록에서 '+dest+' 영역에 추가',async p=>{const before=await model(p),original=await definition(p);await mouseDrag(p,field(p,name),area(p,dest));const d=await definition(p);eq(dest==='values'?d.values.map(v=>v.field):d[dest],[name],'드래그한 필드가 대상 영역에 배치');if(dest==='values')eq(d.values[0].agg,'sum');const after=await model(p);eq(after.source,before.source,'원본 셀 불변');eq(after.undo,before.undo+1,'드래그 한 번은 Undo 한 번');await run(p,'undo');eq(await definition(p),original,'Undo로 이전 피벗 배치 복원');eq((await model(p)).source,before.source);});
for(const [dest,name] of [['pages','지역'],['rows','상품'],['cols','채널'],['values','매출']])await test('터치 Pointer만으로 필드 목록에서 '+dest+' 영역에 추가',async p=>{const before=await model(p),original=await definition(p);await pointerDrag(p,field(p,name),area(p,dest));const d=await definition(p);eq(dest==='values'?d.values.map(v=>v.field):d[dest],[name],'HTML DragEvent 없는 터치 배치');eq((await model(p)).source,before.source);eq((await model(p)).undo,before.undo+1);await run(p,'undo');eq(await definition(p),original);});
await test('행 영역 첫 필드를 아래쪽 두 항목 사이로 재정렬',async p=>{for(const f of ['지역','채널','상품'])await add(p,f,'rows');const before=await model(p),original=await definition(p);await mouseDrag(p,item(p,'rows','지역').locator('.pp-label'),item(p,'rows','상품'),{position:'start'});eq((await definition(p)).rows,['채널','지역','상품']);eq((await model(p)).source,before.source);eq((await model(p)).undo,before.undo+1);await run(p,'undo');eq(await definition(p),original);});

for(const dest of ['pages','cols'])await test(dest+' 영역 내부에서 아래로 한 칸 이동',async p=>{for(const f of ['지역','채널','상품'])await add(p,f,dest);const before=await model(p),original=await definition(p);await mouseDrag(p,item(p,dest,'지역').locator('.pp-label'),item(p,dest,'상품'),{position:'start'});eq((await definition(p))[dest],['채널','지역','상품']);await changedOnce(p,before,original);});
await test('값 영역의 순서를 바꾸면 집계·표시 형식 유지',async p=>{const values=[{field:'매출',agg:'sum',name:'총매출',numFmt:{numFmt:'currency',decimals:0}},{field:'비용',agg:'average',name:'평균비용',numFmt:{numFmt:'number',decimals:2}},{field:'수량',agg:'max',name:'최대수량'}];await configure(p,{values});const before=await model(p),original=await definition(p);await mouseDrag(p,area(p,'values').locator('.pp-item .pp-label').first(),area(p,'values').locator('.pp-item').nth(2),{position:'start'});eq((await definition(p)).values,[values[1],values[0],values[2]]);await changedOnce(p,before,original);});
await test('행 필드를 열로 옮겨도 같은 원본의 값 필드 복사본 유지',async p=>{await configure(p,{rows:['매출'],values:[{field:'매출',agg:'sum',name:'총매출'},{field:'매출',agg:'average',name:'평균매출'}]});const before=await model(p),original=await definition(p);await mouseDrag(p,item(p,'rows','매출').locator('.pp-label'),area(p,'cols'));const d=await definition(p);eq(d.rows,[]);eq(d.cols,['매출']);eq(d.values,original.values);await changedOnce(p,before,original);});
await test('중복 값 중 선택한 한 개만 행으로 옮김',async p=>{await configure(p,{values:[{field:'매출',agg:'sum',name:'총매출'},{field:'매출',agg:'average',name:'평균매출'}]});const before=await model(p),original=await definition(p);await mouseDrag(p,area(p,'values').locator('.pp-item .pp-label').nth(1),area(p,'rows'));const d=await definition(p);eq(d.rows,['매출']);eq(d.values,[original.values[0]]);await changedOnce(p,before,original);});
await test('Σ 값의 열 앞 위치는 값 정의를 바꾸지 않음',async p=>{await configure(p,{cols:['지역','채널'],values:[{field:'매출',agg:'sum'},{field:'비용',agg:'sum'}]});const before=await model(p),original=await definition(p);await mouseDrag(p,area(p,'cols').locator('.sigma .pp-label'),item(p,'cols','지역'),{position:'start'});const d=await definition(p);eq(d.valuesPos,0);eq(d.valuesOnRows,false);eq(d.cols,original.cols);eq(d.values,original.values);await changedOnce(p,before,original);});
await test('Σ 값을 행으로 옮겼다가 열로 다시 이동',async p=>{await configure(p,{rows:['상품'],cols:['지역'],values:[{field:'매출',agg:'sum'},{field:'비용',agg:'sum'}]});for(const [from,to] of [['cols','rows'],['rows','cols']]){const before=await model(p),original=await definition(p);await mouseDrag(p,area(p,from).locator('.sigma .pp-label'),area(p,to));eq((await definition(p)).valuesOnRows,to==='rows');eq((await definition(p)).values,original.values);eq((await model(p)).undo,before.undo+1);eq((await model(p)).source,before.source);}await run(p,'undo');eq((await definition(p)).valuesOnRows,true);await run(p,'undo');eq((await definition(p)).valuesOnRows,false);});
await test('같은 슬롯에 다시 놓으면 Undo 항목을 만들지 않음',async p=>{for(const f of ['지역','채널'])await add(p,f,'rows');const before=await model(p),original=await definition(p);await mouseDrag(p,item(p,'rows','지역').locator('.pp-label'),item(p,'rows','지역'),{position:'start'});await unchanged(p,before,original);});
await test('Escape는 진행 중인 필드 드래그를 취소',async p=>{const before=await model(p),original=await definition(p),to=await pointerDrag(p,field(p,'지역'),area(p,'rows'),{finish:false});eq(await p.locator('.pivot-field-ghost').isVisible(),true);await p.keyboard.press('Escape');await pointerUp(p,to);await unchanged(p,before,original);});
await test('영역 밖에 놓으면 제거하거나 추가하지 않음',async p=>{await add(p,'지역','rows');const before=await model(p),original=await definition(p);await pointerDrag(p,item(p,'rows','지역').locator('.pp-label'),p.locator('.formula-row'));await unchanged(p,before,original);});
await test('드래그 도중 시트 보호를 켜면 적용하지 않음',async p=>{const original=await definition(p),to=await pointerDrag(p,field(p,'지역'),area(p,'rows'),{finish:false});await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.sheets[t.si].protect={on:true,allow:{selectLocked:true}};});const before=await model(p);await pointerUp(p,to);await unchanged(p,before,original);});
await test('드래그 도중 통합 문서 버전이 바뀌면 적용하지 않음',async p=>{const original=await definition(p),to=await pointerDrag(p,field(p,'지역'),area(p,'rows'),{finish:false});await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(t.si,30,10,'별도 편집'));});const before=await model(p);await pointerUp(p,to);await unchanged(p,before,original);});
await test('드래그 도중 다른 시트로 이동해도 원본은 바뀌지 않음',async p=>{const before=await model(p),original=await definition(p),to=await pointerDrag(p,field(p,'지역'),area(p,'rows'),{finish:false});await p.evaluate(()=>window.tabula.switchSheet(0));await pointerUp(p,to);eq(await p.evaluate(()=>window.tabula.si),0);eq((await model(p)).source,before.source);eq((await model(p)).undo,before.undo);await p.evaluate(()=>window.tabula.switchSheet(1));eq(await definition(p),original);eq(await p.locator('.pivot-field-ghost,.pivot-field-marker').count(),0);});
await test('체크박스와 이동 메뉴 버튼은 드래그가 아닌 기존 클릭 실행',async p=>{await p.locator('#pivotPane [data-pivot-field="지역"]').check();eq((await definition(p)).rows,['지역']);eq(await p.locator('.pivot-field-ghost').count(),0);await area(p,'rows').locator('.pp-menu').click();await p.getByRole('menuitem',{name:'열(으)로 이동',exact:true}).click();eq((await definition(p)).cols,['지역']);eq((await definition(p)).rows,[]);eq(await p.locator('.pivot-field-ghost').count(),0);});
await test('Chromium 실제 터치 스트림으로 필드 목록에서 필터에 추가',async(p,c)=>{const before=await model(p),original=await definition(p),a=await field(p,'지역').boundingBox(),b=await area(p,'pages').boundingBox(),cdp=await c.newCDPSession(p),from={x:a.x+20,y:a.y+a.height/2},to={x:b.x+b.width/2,y:b.y+b.height-5};await cdp.send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{...from,id:0}]});for(let i=1;i<=12;i++)await cdp.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[{x:from.x+(to.x-from.x)*i/12,y:from.y+(to.y-from.y)*i/12,id:0}]});await cdp.send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});await raf(p);eq((await definition(p)).pages,['지역']);await changedOnce(p,before,original);},{cdpOnly:true});
for(const width of [320,390])await test(width+'px 모바일에서 터치 드래그와 대상 표시 접근',async p=>{await p.setViewportSize({width,height:844});await p.locator('#mobileModeToggle').click();await raf(p);const before=await model(p),original=await definition(p),to=await pointerDrag(p,field(p,'지역'),area(p,'pages'),{finish:false});eq(await p.locator('.pivot-field-ghost').isVisible(),true);eq(await p.locator('.pivot-field-marker').isVisible(),true);await p.screenshot({path:out+'/mobile-'+width+'-drag.png'});await pointerUp(p,to);eq((await definition(p)).pages,['지역']);await changedOnce(p,before,original);});
await test('긴 값 영역 가장자리 드래그는 자동 스크롤 후 정확히 한 번 추가',async p=>{const values=Array.from({length:20},(_,i)=>({field:'매출',agg:i%2?'average':'sum',name:'지표'+i}));await configure(p,{values});const before=await model(p),original=await definition(p),box=area(p,'values');eq(await box.evaluate(n=>n.scrollTop),0);const to=await pointerDrag(p,field(p,'비용'),box,{finish:false});await p.waitForFunction(()=>document.querySelector('#pivotPane [data-pivot-area="values"]').scrollTop>30);ok(await box.evaluate(n=>n.scrollTop)>30,'실제 앱 값 목록 자동 스크롤');await p.evaluate(()=>window.addEventListener('pointerup',()=>{window.__pivotDropScroll=document.querySelector('#pivotPane [data-pivot-area="values"]').scrollTop;},{capture:true,once:true}));await pointerUp(p,to);const scroll=await p.evaluate(()=>({before:window.__pivotDropScroll,after:document.querySelector('#pivotPane [data-pivot-area="values"]').scrollTop}));ok(scroll.before>0,'드롭 직전 스크롤이 실제 이동됨');eq(scroll.after,scroll.before,'드롭 뒤 값 영역 스크롤 위치 보존');const after=await definition(p);eq(after.values.length,21);eq(after.values.filter(v=>v.field==='비용').length,1);eq(after.values.filter(v=>v.field==='매출'),values);await changedOnce(p,before,original);});
await test('필드 창을 연 뒤 일반 셀을 편집해도 새 드래그는 정상 시작',async p=>{await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(t.si,30,10,'드래그 전에 마친 편집'));});await raf(p);const before=await model(p),original=await definition(p);await mouseDrag(p,field(p,'지역'),area(p,'rows'));eq((await definition(p)).rows,['지역']);await changedOnce(p,before,original);eq(await p.evaluate(()=>{const t=window.tabula;return t.wb().getRaw(t.si,30,10);}), '드래그 전에 마친 편집');});
}finally{await browser.close();const summary={url,baseline:baseline||null,browser:browserName,browserVersion:browser.version(),simulation:'Isolated synthetic workbook; real Playwright mouse input, simulated Pointer/Touch and DataTransfer/HTML DragEvent paths; no physical-device verification',assets:[...assets],skipped,cases:results.length,passed:results.filter(r=>r.ok).length,checks,pageErrors:errors,remoteWrites:writes,blockedRequests:blocked,results};await writeFile(out+'/result.json',JSON.stringify(summary,null,2));console.log(JSON.stringify({url,baseline:baseline||null,browser:browserName,cases:summary.cases,passed:summary.passed,checks,assets:summary.assets,out,failures:results.filter(r=>!r.ok).map(r=>({name:r.name,error:r.error.split('\n')[0]}))}));if(!results.length||results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;}
