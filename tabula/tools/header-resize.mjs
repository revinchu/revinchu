// 합성 시트에서 실제 마우스 경계 hover·끌기·더블 클릭과 실행 취소를 검증합니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 문서 전용 검사입니다.');
const engines = (process.env.WIXEL_BROWSER || 'chromium,webkit').split(',');
const out = process.env.WIXEL_HEADER_RESIZE_OUT || 'D:/Codex/Temp/wixel-header-resize/source';
const only = process.env.WIXEL_HEADER_RESIZE_FILTER || '';
await mkdir(out, { recursive: true });
const results = [], assets = new Set(); let checks = 0;
const eq = (a,b,m) => { checks++; assert.deepEqual(a,b,m); };
const ok = (a,m) => { checks++; assert.ok(a,m); };
const near = (a,b,m) => ok(Math.abs(a-b)<=1, m+': '+a+' / '+b);
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
function fixture(freeze=false) {
  const cells = {}, colWidths = {}, rowHeights = {};
  for(let c=0;c<20;c++) { colWidths[c]=100; cells['0,'+c]={raw:'열 '+c}; cells['2,'+c]={raw:'열 너비 자동 맞춤 확인을 위한 충분히 긴 문자열 '+c}; }
  for(let r=0;r<50;r++) { rowHeights[r]=30; cells[r+',0']={raw:'첫 번째 줄\n두 번째 줄\n세 번째 줄',style:{size:12,wrap:true}}; }
  cells['10,10']={raw:'=1+2',type:'n',value:3};
  return {sheets:[{name:'행열 경계 합성 검사',cells,colWidths,rowHeights,...(freeze?{freeze:{rows:2,cols:2}}:{})}]};
}
async function seed(p,t={}) {
  await p.evaluate(({data,zoom,scroll})=>{ const a=tabula,w=a.wb();w.restore(data);w.setSnapshots(data);a.switchSheet(0);a.selectCell(1,1,{scroll:false});a.gv().setZoom(zoom);a.gv().setScroll(scroll?200:0,scroll?90:0);a.gv().layout();a.gv().renderAll();w.undoStack=[];w.redoStack=[]; },{data:fixture(t.freeze),zoom:t.zoom||100,scroll:t.scroll});
  await raf(p);
}
async function point(p,axis,index,edge=true,offset=0) {
  return p.evaluate(({axis,index,edge,offset})=>{const g=tabula.gv(),v=g.viewEl.getBoundingClientRect(),b=g.clientRect({r1:axis==='row'?index:0,r2:axis==='row'?index:0,c1:axis==='col'?index:0,c2:axis==='col'?index:0});return axis==='col'?{x:(edge?b.right:(b.left+b.right)/2)+offset,y:v.top+(g.hh+(g.olh||0))*g.z/2}:{x:v.left+(g.hw+(g.olw||0))*g.z/2,y:(edge?b.bottom:(b.top+b.bottom)/2)+offset};},{axis,index,edge,offset});
}
const size = (p,axis,i) => p.evaluate(({axis,i})=>axis==='col'?tabula.wb().colWidth(0,i):tabula.gv().rows.size(i),{axis,i});
const sizes = (p,axis,indices) => p.evaluate(({axis,indices})=>indices.map(i=>axis==='col'?tabula.wb().colWidth(0,i):tabula.gv().rows.size(i)),{axis,indices});
const state=p=>p.evaluate(()=>{const w=tabula.wb(),s=w.sheets[0];return{colWidths:structuredClone(s.colWidths),rowHeights:structuredClone(s.rowHeights),rowManual:structuredClone(s.rowManual),cells:Object.fromEntries(s.cells.entries()),undo:w.undoStack.length,sel:structuredClone(tabula.sel)};});
async function hover(p,axis,i,offset=0) {
  const pt=await point(p,axis,i,true,offset);await p.mouse.move(pt.x,pt.y);await raf(p);
  const v=await p.evaluate(pt=>{const target=document.elementFromPoint(pt.x,pt.y);return{cursor:target&&getComputedStyle(target).cursor,hit:tabula.gv().hitTest(pt.x,pt.y),target:target?.className};},pt);
  eq(v.hit.zone,axis==='col'?'colHeader':'rowHeader','머리글 경계 영역');eq(v.hit[axis==='col'?'edgeCol':'edgeRow'],i,'경계 대상 정확');
  ok(v.cursor?.includes(axis==='col'?'col-resize':'row-resize'),'경계 크기 조절 커서: '+JSON.stringify(v));return pt;
}
async function dragBy(p,axis,i,delta){const pt=await hover(p,axis,i);await p.mouse.down();await p.mouse.move(pt.x+(axis==='col'?delta:0),pt.y+(axis==='row'?delta:0),{steps:8});ok(await p.locator('#headerResizeHint').isVisible(),'끌기 중 크기 안내 표시');ok(await p.locator('#headerResizeGuide').isVisible(),'끌기 중 경계 안내선 표시');ok((await p.locator('#headerResizeHint').textContent()).includes(axis==='col'?'너비':'높이'),'축에 맞는 크기 안내');await p.mouse.up();await raf(p);ok(!(await p.locator('#headerResizeGuide').isVisible()),'끌기 종료 후 안내선 제거');}
async function double(p,axis,i){const pt=await point(p,axis,i);await p.mouse.dblclick(pt.x,pt.y,{delay:70});await raf(p);}
async function selectHeaders(p,axis,a,b){const x=await point(p,axis,a,false),y=await point(p,axis,b,false);await p.mouse.click(x.x,x.y);if(a!==b){await p.keyboard.down('Shift');await p.mouse.click(y.x,y.y);await p.keyboard.up('Shift');}await raf(p);}
const tests=[];const test=(name,options,fn)=>tests.push({...options,name,fn});
for(const profile of [{name:'desktop',mobile:false,ipad:false},{name:'mobile',mobile:true,ipad:false},{name:'ipad',mobile:false,ipad:true}])for(const zoom of[55,100,150])test(profile.name+'-'+zoom,{...profile,zoom},async(p,t,label)=>{
  const before=await state(p);for(const axis of['col','row']){
    const i=axis==='col'?(t.mobile?1:2):3,old=await size(p,axis,i),delta=axis==='col'?(t.mobile?22:44):33;
    await hover(p,axis,i);await dragBy(p,axis,i,delta);near(await size(p,axis,i),old+delta/(t.zoom/100),'배율을 반영한 '+axis+' 크기');
    await p.evaluate(()=>tabula.run('undo'));await raf(p);near(await size(p,axis,i),old,'끌기 실행 취소');
    await double(p,axis,i);const fitted=await size(p,axis,i);ok(fitted>old+10,'내용 기준 자동 맞춤 '+axis+': '+fitted);
    await p.evaluate(()=>tabula.run('undo'));await raf(p);near(await size(p,axis,i),old,'자동 맞춤 실행 취소');
  }eq((await state(p)).cells,before.cells,'값·수식·서식 보존');await p.screenshot({path:out+'/'+label+'.png'});
});
for(const config of[{name:'frozen',freeze:true,scroll:false},{name:'frozen-scroll',freeze:true,scroll:true},{name:'scroll',freeze:false,scroll:true}])test(config.name,{...config,zoom:100},async(p,t)=>{
  for(const axis of['col','row'])for(const i of[t.scroll?7:4,...(t.freeze?[1]:[])]){const before=await size(p,axis,i);await dragBy(p,axis,i,21);near(await size(p,axis,i),before+21,'틀 고정·스크롤 '+axis+' '+i);await double(p,axis,i);ok((await size(p,axis,i))>before,'틀 고정·스크롤 자동 맞춤 '+axis+' '+i);}
});
for(const axis of['col','row'])test('multi-'+axis,{zoom:100},async p=>{const indices=[1,2,3],before=await sizes(p,axis,indices);await selectHeaders(p,axis,1,3);await dragBy(p,axis,3,25);eq(await sizes(p,axis,indices),before.map(v=>v+25),'선택한 여러 '+axis+' 끌기');await p.evaluate(()=>tabula.run('undo'));await raf(p);eq(await sizes(p,axis,indices),before,'다중 크기 실행 취소');await double(p,axis,3);const fit=await sizes(p,axis,indices);ok(fit.every((n,i)=>n>before[i]+10),'선택한 모든 '+axis+' 자동 맞춤: '+fit);await p.evaluate(()=>tabula.run('undo'));await raf(p);eq(await sizes(p,axis,indices),before,'다중 자동 맞춤 한 번 실행 취소');});
for(const axis of['col','row'])test('small-zoom-hit-'+axis,{zoom:55},async p=>{await hover(p,axis,axis==='col'?2:3,-3);await hover(p,axis,axis==='col'?2:3,3);});

for(const axis of ['col','row']) {
  test('cancel-'+axis,{zoom:100},async p=>{const i=axis==='col'?2:3,before=await state(p);for(const mode of ['escape','blur']){const pt=await hover(p,axis,i);await p.mouse.down();await p.mouse.move(pt.x+(axis==='col'?35:0),pt.y+(axis==='row'?35:0),{steps:5});if(mode==='escape')await p.keyboard.press('Escape');else await p.evaluate(()=>window.dispatchEvent(new Event('blur')));await p.mouse.up();await raf(p);const after=await state(p);eq(after.colWidths,before.colWidths,mode+' 취소 열 너비 복원');eq(after.rowHeights,before.rowHeights,mode+' 취소 행 높이 복원');eq(after.rowManual,before.rowManual,mode+' 취소 수동 높이 복원');eq(after.undo,0,mode+' 취소 Undo 없음');}const pt=await point(p,axis,i);await p.mouse.click(pt.x,pt.y);await raf(p);eq((await state(p)).undo,0,'경계 단일 클릭은 Undo 없음');});
  test('protected-'+axis,{zoom:100},async p=>{const i=axis==='col'?2:3;await p.evaluate(()=>{tabula.wb().sheets[0].protect={on:true,allow:{selectLocked:true,selectUnlocked:true}};});const before=await state(p),pt=await point(p,axis,i);await p.mouse.move(pt.x,pt.y);await p.mouse.down();await p.mouse.move(pt.x+(axis==='col'?30:0),pt.y+(axis==='row'?30:0),{steps:5});await p.mouse.up();await double(p,axis,i);const after=await state(p);eq(after.colWidths,before.colWidths,'보호 열 너비 변경 차단');eq(after.rowHeights,before.rowHeights,'보호 행 높이 변경 차단');eq(after.undo,0,'보호 변경 Undo 없음');await p.evaluate(axis=>{tabula.wb().sheets[0].protect.allow[axis==='col'?'formatColumns':'formatRows']=true;},axis);await dragBy(p,axis,i,20);near(await size(p,axis,i),(axis==='col'?100:30)+20,'보호 중 명시적 서식 허용');});
  test('marked-final-'+axis,{zoom:100},async p=>{const i=axis==='col'?2:3;await p.evaluate(()=>{tabula.wb().props.markedFinal=true;});const before=await state(p),pt=await point(p,axis,i);await p.mouse.move(pt.x,pt.y);await p.mouse.down();await p.mouse.move(pt.x+(axis==='col'?30:0),pt.y+(axis==='row'?30:0),{steps:4});await p.mouse.up();await double(p,axis,i);const after=await state(p);eq(after.colWidths,before.colWidths,'최종본 열 너비 변경 차단');eq(after.rowHeights,before.rowHeights,'최종본 행 높이 변경 차단');eq(after.undo,0,'최종본 변경 Undo 없음');});
}
test('row-manual-height',{zoom:100},async p=>{await dragBy(p,'row',3,25);eq((await state(p)).rowManual[3],true,'끌기는 수동 높이로 표시');eq((await state(p)).undo,1,'한 번 끌기 Undo 1회');await double(p,'row',3);eq(!!(await state(p)).rowManual[3],false,'자동 맞춤은 수동 높이 해제');eq((await state(p)).undo,2,'자동 맞춤 Undo 1회');});


test('xlsx-roundtrip',{zoom:100},async(p,t,label)=>{
  await dragBy(p,'col',2,37);await dragBy(p,'row',3,28);await double(p,'col',4);await double(p,'row',4);
  const before={cols:await sizes(p,'col',[2,4]),rows:await sizes(p,'row',[3,4]),manual:(await state(p)).rowManual};
  await p.evaluate(()=>{window.__headerFile=null;window.showSaveFilePicker=async()=>({name:'행열크기.xlsx',createWritable:async()=>({write:async value=>{const b=value instanceof Blob?value:new Blob([value]);window.__headerFile=Array.from(new Uint8Array(await b.arrayBuffer()));},close:async()=>{},abort:async()=>{}})});});
  await p.evaluate(()=>tabula.exportXlsx('행열크기','xlsx'));await p.waitForFunction(()=>window.__headerFile?.length>0);const bytes=Buffer.from(await p.evaluate(()=>window.__headerFile));await writeFile(out+'/'+label+'.xlsx',bytes);
  await p.evaluate(()=>{window.__headerOldBook=tabula.wb();});await p.locator('#fileInput').setInputFiles({name:'행열크기.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes});await p.waitForFunction(()=>tabula.wb()!==window.__headerOldBook&&tabula.wb().sheets[0].name==='행열 경계 합성 검사');await raf(p);
  const after={cols:await sizes(p,'col',[2,4]),rows:await sizes(p,'row',[3,4]),manual:(await state(p)).rowManual};
  for(let i=0;i<2;i++){near(after.cols[i],before.cols[i],'XLSX 열 너비 유지');near(after.rows[i],before.rows[i],'XLSX 행 높이 유지');}
  eq(!!after.manual[3],true,'XLSX 수동 행 높이 유지');eq(!!after.manual[4],false,'XLSX 자동 행 높이 유지');eq(await p.evaluate(()=>tabula.wb().getCell(0,10,10).raw),'=1+2','XLSX 원본 수식 보존');
});


test('all-selection-zero-size',{zoom:100},async p=>{
  const corner=await p.evaluate(()=>{const g=tabula.gv(),b=g.viewEl.getBoundingClientRect();return{x:b.left+g.hw*g.z/2,y:b.top+g.hh*g.z/2};});await p.mouse.click(corner.x,corner.y);await raf(p);
  const before=await state(p),defaults=await p.evaluate(()=>({col:tabula.wb().sheets[0].defColW,row:tabula.wb().sheets[0].defRowH}));
  for(const axis of ['col','row']){await dragBy(p,axis,axis==='col'?2:3,axis==='col'?-140:-55);const after=await state(p);eq(after.colWidths,before.colWidths,'전체 선택 너비 0 방지');eq(after.rowHeights,before.rowHeights,'전체 선택 높이 0 방지');eq(after.undo,0,'전체 숨김 차단 Undo 없음');eq(await p.evaluate(()=>({col:tabula.wb().sheets[0].defColW,row:tabula.wb().sheets[0].defRowH})),defaults,'전체 선택 기본 크기 보존');}
});

for(const engine of engines){const browser=await pw[engine].launch();try{for(const t of tests){const label=engine+'-'+t.name;if(only&&!label.includes(only))continue;const context=await browser.newContext({viewport:{width:t.mobile?390:1280,height:900},isMobile:!!t.ipad,hasTouch:!!t.ipad,serviceWorkers:'block',...(t.ipad?{userAgent:'Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1'}:{})}),p=await context.newPage(),errors=[],writes=[];p.setDefaultTimeout(10000);
  await context.addInitScript(mobile=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1',mobile?'on':'off');},!!t.mobile);
  await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD'].includes(q.method())){writes.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return r.abort();return r.continue();});p.on('pageerror',e=>errors.push(e.message));const start=checks;
  try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());assets.add(await p.evaluate(()=>[...document.scripts].map(s=>s.getAttribute('src')).find(s=>s?.includes('wixel-'))||'source'));await seed(p,t);await t.fn(p,t,label);eq(errors,[],'실행 오류 없음');eq(writes,[],'외부 쓰기 없음');results.push({label,ok:true,checks:checks-start});console.log('OK '+label);}catch(e){results.push({label,ok:false,error:e.message,checks:checks-start,errors,writes,state:await state(p).catch(()=>null)});console.error('NG '+label+': '+e.message);await p.screenshot({path:out+'/'+label+'-failure.png'}).catch(()=>{});}finally{await context.close();}
}}finally{await browser.close();}}
const result={url,assets:[...assets],cases:results.length,passed:results.filter(x=>x.ok).length,checks,results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify({...result,results:results.map(({label,ok,error,checks})=>({label,ok,error,checks}))}));if(result.passed!==result.cases)process.exitCode=1;
