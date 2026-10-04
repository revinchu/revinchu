// Synthetic, isolated browser compatibility regressions. No user workbooks or OS clipboard.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { readXlsx } from '../src/xlsx.js';
import { readWixelFile } from '../src/wixel-file.js';
import { Workbook } from '../src/workbook.js';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const names=(process.env.WIXEL_BROWSERS||'chromium,webkit,firefox').split(',');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5195/';
const out=process.env.WIXEL_BROWSER_QUALITY_OUT||'D:/Codex/Temp/wixel-browser-quality';
const only=process.env.WIXEL_BROWSER_QUALITY_FILTER||'',exclude=process.env.WIXEL_BROWSER_QUALITY_EXCLUDE||'';
await mkdir(out,{recursive:true});
let checks=0;const results=[],errors=[],writes=[],blocked=[],versions={},capabilities=[];
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(x,m)=>{checks++;assert.ok(x,m);};
const frame=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const fixture={sheets:[{name:'합성 자료',cells:{'0,0':{raw:'12',style:{bold:true,fill:'#ffeeaa'},comment:'메모'},'0,1':{raw:'=A1*2'},'1,0':{raw:'한글🙂'},'3,0':{raw:'대상'},'3,1':{raw:'보존'}}},{name:'다른 시트',cells:{}}]};
for(const engine of names){
 const browser=await engines[engine].launch({headless:true});versions[engine]=browser.version();
 async function test(label,fn,{mobile=false,missing=false}={}){
  const name=engine+'/'+label;if(only&&!name.includes(only)||exclude&&name.includes(exclude))return;
  const ctx=await browser.newContext({viewport:{width:mobile?390:1366,height:mobile?844:960},hasTouch:mobile,...(engine==='firefox'?{}:{isMobile:mobile}),acceptDownloads:true,serviceWorkers:'block'});
  const p=await ctx.newPage();p.setDefaultTimeout(10000);const pe=[],pw=[];
  p.on('pageerror',e=>{pe.push(e.message);errors.push({name,message:e.message});});
  ctx.on('page',q=>q.on('pageerror',e=>{pe.push(e.message);errors.push({name,message:e.message});}));
  p.on('dialog',d=>d.type()==='beforeunload'?d.accept():d.dismiss());
  await ctx.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){pw.push(u.pathname);writes.push({name,path:u.pathname});return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')){blocked.push({name,path:u.pathname});return r.abort();}return r.continue();});
  await ctx.addInitScript(({mobile,missing})=>{
   window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;try{localStorage.setItem('wixel.mobile-work.v1',mobile?'on':'off');localStorage.setItem('wixel.mobile-keyboard.v1','hardware');}catch{/* The init script also runs in the new page's opaque about:blank document. */}
   window.__idbErrors=[];document.addEventListener('error',e=>{if(e.target?.error)window.__idbErrors.push({name:e.target.error.name,message:e.target.error.message});},true);
   const txOrig=IDBDatabase.prototype.transaction;IDBDatabase.prototype.transaction=function(...args){const tx=txOrig.apply(this,args);tx.addEventListener('error',e=>window.__idbErrors.push({name:e.target?.error?.name,message:e.target?.error?.message,txName:tx.error?.name}));tx.addEventListener('abort',()=>window.__idbErrors.push({stage:'abort',name:tx.error?.name,message:tx.error?.message}));return tx;};
   window.__clipboard={reads:0,writes:0,deferred:false};
   Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>{__clipboard.reads++;if(__clipboard.deferred)return new Promise(r=>__clipboard.resolve=r);throw new DOMException('합성 권한 거절','NotAllowedError');},writeText:async()=>{__clipboard.writes++;throw new DOMException('합성 권한 거절','NotAllowedError');}}});
   window.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&/^Key[CVX]$/.test(e.code))e.preventDefault();});
   Object.defineProperty(window,'showSaveFilePicker',{configurable:true,value:undefined});
   Object.defineProperty(window,'showOpenFilePicker',{configurable:true,value:undefined});
   if(missing){for(const k of ['CompressionStream','DecompressionStream'])Object.defineProperty(window,k,{configurable:true,value:undefined});Object.defineProperty(navigator,'locks',{configurable:true,value:undefined});}
  },{mobile,missing});
  const start=checks;
  try{
   await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
   capabilities.push({name,assets:await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))),...await p.evaluate(()=>({compression:typeof CompressionStream,locks:!!navigator.locks,filePicker:typeof showSaveFilePicker,clipboard:typeof navigator.clipboard?.readText,ua:navigator.userAgent}))});
   if(await p.locator('#autosaveToggle').getAttribute('aria-checked')==='true')await p.locator('#autosaveToggle').evaluate(n=>n.click());
   await p.evaluate(data=>{const t=tabula,w=t.wb();w.restore(data);t.switchSheet(1);t.switchSheet(0);t.selectCell(0,0);w.undoStack=[];w.redoStack=[];},fixture);
   await fn(p,ctx);eq(pe,[],'페이지 오류 없음');eq(pw,[],'원격 쓰기 없음');results.push({name,ok:true,checks:checks-start,handledStorageErrors:await p.evaluate(()=>window.__idbErrors)});console.log('OK '+name);
  }catch(e){results.push({name,ok:false,checks:checks-start,error:e.stack});console.error('NG '+name+' '+e.stack);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}
  finally{await ctx.close();}
 }
 try{
  await test('지연 클립보드 읽기 중 같은 시작 셀의 선택 범위 확장 취소',async p=>{
   await p.evaluate(()=>{__clipboard.deferred=true;tabula.selectCell(3,0);tabula.run('paste');});await p.waitForFunction(()=>typeof __clipboard.resolve==='function');
   await p.evaluate(()=>{tabula.selectRange({r1:3,c1:0,r2:3,c2:1},'cells',{r:3,c:0});__clipboard.resolve('늦게 받은 외부 값');});await frame(p);
   eq(await p.evaluate(()=>[tabula.wb().getRaw(0,3,0),tabula.wb().getRaw(0,3,1),tabula.wb().undoStack.length]),['대상','보존',0],'이전 비동기 붙여넣기가 새 범위를 덮어쓰지 않음');
  });
  await test('셀 입력·수식·실행 취소 및 다시 실행',async p=>{
   await p.keyboard.type('=SUM(4,8)');await p.keyboard.press('Enter');eq(await p.evaluate(()=>tabula.wb().getValue(0,0,0)),12);eq(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'=SUM(4,8)');
   await p.keyboard.press('Control+z');eq(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'12');await p.keyboard.press('Control+y');eq(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'=SUM(4,8)');
  });
  await test('모바일 Ctrl·Meta 내부 복사 붙여넣기 권한 거절',async p=>{
   for(const mod of ['Control','Meta']){await p.evaluate(()=>tabula.selectRange({r1:0,c1:0,r2:1,c2:1}));await p.keyboard.press(mod+'+c');await p.evaluate(()=>tabula.selectCell(3,0));await p.keyboard.press(mod+'+v');await frame(p);
    eq(await p.evaluate(()=>[tabula.wb().getRaw(0,3,0),tabula.wb().getRaw(0,3,1),tabula.wb().getRaw(0,4,0)]),['12','=A4*2','한글🙂']);eq(await p.evaluate(()=>tabula.wb().styleAt(0,3,0).bold),true);eq(await p.evaluate(()=>__clipboard.reads),0);await p.evaluate(()=>tabula.run('undo'));eq(await p.evaluate(()=>tabula.wb().getRaw(0,3,0)),'대상');}
  },{mobile:true});
  await test('셀과 슬라이서 우클릭 메뉴 분리·다중 선택·편집 초점 없음',async p=>{
   await p.evaluate(()=>{const t=tabula,w=t.wb();w.restore({sheets:[{name:'자료',cells:{'0,0':{raw:'항목'},'0,1':{raw:'값'},'1,0':{raw:'가'},'1,1':{raw:'1'},'2,0':{raw:'나'},'2,1':{raw:'2'},'3,0':{raw:'다'},'3,1':{raw:'3'}},slicers:[{id:'compat-sl',caption:'항목',source:{kind:'pivot',field:'항목',pivots:[{sheet:'결과',name:'P'}]},x:25,y:50,w:200,h:230}]},{name:'결과',cells:{},pivot:{name:'P',source:'자료',range:{r1:0,c1:0,r2:3,c2:1},rows:['항목'],cols:[],values:[{field:'값',agg:'sum'}],top:0,left:0}}]});t.switchSheet(1);t.run('pivotRefresh');t.switchSheet(0);t.selectCell(0,0);});
   const sl=p.locator('.obj.slicer[data-id="compat-sl"]:visible').first();await sl.locator('.sl-item').filter({hasText:/^가$/}).click();await sl.locator('.sl-item').filter({hasText:/^나$/}).click({modifiers:['Control']});
   eq(await p.evaluate(()=>tabula.wb().sheets[1].pivot.filters.항목),['가','나']);eq(await p.evaluate(()=>document.activeElement.matches('input,textarea:not(.idle),[contenteditable=true]')),false);
   await sl.locator('.sl-cap').click({button:'right'});const text=await p.locator('#menuLayer').innerText();ok(text.includes('슬라이서 설정'),'슬라이서 전용 메뉴');ok(!text.includes('셀 서식'),'개체 메뉴에 셀 서식 없음');await p.keyboard.press('Escape');
   await p.evaluate(()=>tabula.selectCell(8,5));const pos=await p.evaluate(()=>{const r=tabula.gv().clientRect({r1:8,c1:5,r2:8,c2:5});return{x:(r.left+r.right)/2,y:(r.top+r.bottom)/2};});await p.mouse.click(pos.x,pos.y,{button:'right'});ok((await p.locator('#menuLayer').innerText()).includes('셀 서식'),'셀 우클릭은 셀 서식');
  });
  await test('좁은 화면 리본 드래그 양방향·손바닥 두 축 이동',async p=>{
   const before=await p.evaluate(()=>JSON.stringify(tabula.wb().serialize()));const tabs=p.locator('#ribbonTabs');
   const point=async()=>tabs.evaluate(n=>{const b=n.getBoundingClientRect();return{x:b.right-25,y:b.y+b.height/2};});let a=await point();await p.mouse.move(a.x,a.y);await p.mouse.down();await p.mouse.move(a.x-110,a.y,{steps:10});await p.mouse.up();await frame(p);const right=await tabs.evaluate(n=>n.scrollLeft);ok(right>20,'리본 오른쪽 보이기');
   a=await point();await p.mouse.move(a.x-130,a.y);await p.mouse.down();await p.mouse.move(a.x-20,a.y,{steps:10});await p.mouse.up();await frame(p);ok(await tabs.evaluate(n=>n.scrollLeft)<right-20,'리본 왼쪽 복귀');
   await p.locator('#mobileHandPan').click();const vp=await p.locator('#gridView').boundingBox();const x=vp.x+vp.width*.6,y=vp.y+vp.height*.6;const old=await p.evaluate(()=>({x:tabula.gv().sx,y:tabula.gv().sy}));await p.mouse.move(x,y);await p.mouse.down();await p.mouse.move(x-70,y-60,{steps:10});await frame(p);ok(await p.locator('#gridPanCursor').isVisible(),'손 커서 표시');await p.mouse.up();await frame(p);const after=await p.evaluate(()=>({x:tabula.gv().sx,y:tabula.gv().sy}));ok(after.x>old.x+30&&after.y>old.y+30,'본문 두 축 이동');eq(await p.evaluate(()=>JSON.stringify(tabula.wb().serialize())),before,'탐색은 문서를 변경하지 않음');await p.keyboard.press('Escape');
  },{mobile:true});
  for(const missing of [false,true])await test('IndexedDB v5 저장·새 페이지 복원'+(missing?' 압축·Web Locks 없음':''),async(p,ctx)=>{
   const expected=await p.evaluate(async()=>{const {saveLargeWorkbook}=await import('/src/big-storage.js');const t=tabula,w=t.wb();w.transact(()=>{for(let r=0;r<15000;r++){w.setInput(0,r+10,0,String(r));w.setInput(0,r+10,1,'=A'+(r+11)+'*2');w.setInput(0,r+10,2,'한글-'+r);w.setInput(0,r+10,3,String(r%11));}w.setStyle(0,15009,2,{bold:true,fill:'#abc123'});});try{await saveLargeWorkbook('compat-test',w,{docName:'합성 호환성',si:0});}catch(e){throw new Error('saveLargeWorkbook failed: '+JSON.stringify({name:e?.name,message:e?.message,error:String(e),diagnostic:window.__idbErrors}));}return{cellCount:w.sheets[0].cells.size,last:w.getRaw(0,15009,2),formula:w.getRaw(0,15009,1)};});
   const next=await ctx.newPage();await next.goto(url,{waitUntil:'domcontentloaded'});await next.waitForFunction(()=>window.tabula?.wb());const actual=await next.evaluate(async()=>{const {loadLargeWorkbook}=await import('/src/big-storage.js'),{idbGet}=await import('/src/storage.js'),{Workbook}=await import('/src/workbook.js');const data=await loadLargeWorkbook('compat-test',null,{prepareCells:true}),w=new Workbook(data.workbook);return{cellCount:w.sheets[0].cells.size,last:w.getRaw(0,15009,2),formula:w.getRaw(0,15009,1),value:w.getValue(0,15009,1),style:w.styleAt(0,15009,2),version:(await idbGet('compat-test')).v};});
   eq(actual.cellCount,expected.cellCount);eq(actual.last,expected.last);eq(actual.formula,expected.formula);eq(actual.value,29998);eq(actual.style.fill,'#abc123');eq(actual.style.bold,true);eq(actual.version,5);await next.close();
  },{missing});
  for(const missing of [false,true])await test('IndexedDB 보관함 현재 문서·과거 버전 새 페이지 복원'+(missing?' 압축·Web Locks 없음':''),async(p,ctx)=>{
   const version=await p.evaluate(async()=>{const {libSave}=await import('/src/library.js'),{librarySnapshot}=await import('/src/snapshot-blob.js');const w=tabula.wb();w.transact(()=>w.setInput(0,7,0,'한글🙂'.repeat(180000)));const one=librarySnapshot(w,{docName:'보관 검사',docId:'library-compat'});await libSave('library-compat','보관 검사',one.blob,{version:{label:'첫 기록'}});w.transact(()=>w.setInput(0,0,0,'29'));const two=librarySnapshot(w,{docName:'보관 검사',docId:'library-compat'});const entry=await libSave('library-compat','보관 검사',two.blob,{version:{label:'둘째 기록'}});return entry.versions[0].ts;});
   const next=await ctx.newPage();await next.goto(url,{waitUntil:'domcontentloaded'});await next.waitForFunction(()=>window.tabula?.wb());const state=await next.evaluate(async ts=>{const {libLoad,libLoadVersion,libList}=await import('/src/library.js'),{Workbook}=await import('/src/workbook.js');const now=new Workbook((await libLoad('library-compat')).workbook),old=new Workbook((await libLoadVersion('library-compat',ts)).workbook),entry=(await libList()).find(x=>x.id==='library-compat');return{now:now.getRaw(0,0,0),old:old.getRaw(0,0,0),longText:now.getRaw(0,7,0)==='한글🙂'.repeat(180000),versions:entry.versions.map(x=>x.label),bold:now.styleAt(0,0,0).bold};},version);eq(state,{now:'29',old:'12',longText:true,versions:['첫 기록','둘째 기록'],bold:true});await next.close();
  },{missing});
  for(const kind of ['xlsx','wixel'])for(const missing of [false,true])await test('실제 Blob 다운로드와 재읽기 '+kind+(missing?' 압축 API 없음':''),async p=>{
   await p.evaluate(()=>{const real=window.setTimeout;window.setTimeout=(fn,ms,...args)=>real(fn,ms===0?5:ms,...args);tabula.run('saveAs');});
   const d=p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});await d.locator('select').selectOption(kind);await d.getByLabel('파일 이름',{exact:true}).fill('합성 호환성');await d.getByRole('button',{name:'저장',exact:true}).click();
   const fallback=p.getByRole('dialog',{name:'파일로 저장',exact:true});await fallback.waitFor();const event=p.waitForEvent('download',{timeout:30000});await fallback.getByRole('button',{name:'다운로드',exact:true}).click();const download=await event;eq(download.suggestedFilename(),'합성 호환성.'+kind);const parts=[];for await(const part of await download.createReadStream())parts.push(part);const bytes=Buffer.concat(parts);ok(bytes.length>500,'실제 파일 바이트');const data=kind==='xlsx'?readXlsx(bytes).data:(await readWixelFile(new Blob([bytes]))).workbook;const w=new Workbook(data);eq(w.getRaw(0,0,0),'12');eq(w.getRaw(0,0,1),'=A1*2');eq(w.getValue(0,0,1),24);eq(w.styleAt(0,0,0).bold,true);eq(w.getRaw(0,1,0),'한글🙂');await p.evaluate(()=>{const w=tabula.wb();w.transact(()=>w.setInput(0,0,0,'재열기 전 변경'));});await p.locator('#fileInput').setInputFiles({name:'합성 호환성.'+kind,mimeType:'application/octet-stream',buffer:bytes});await p.waitForFunction(()=>tabula.wb().getRaw(0,0,0)==='12'&&!document.querySelector('.load-progress'));eq(await p.evaluate(()=>[tabula.wb().getRaw(0,0,1),tabula.wb().getValue(0,0,1),tabula.wb().getRaw(0,1,0)]),['=A1*2',24,'한글🙂'],'브라우저 파일 열기에서 저장본 수식·값 복원');await download.delete();
  },{missing});
 }finally{await browser.close();await writeFile(out+'/result.json',JSON.stringify({url,filter:only,exclude,versions,checks,cases:results.length,passed:results.filter(r=>r.ok).length,errors,writes,blocked,capabilities,results,scope:'Fresh Chromium/WebKit/Firefox browser engines; small synthetic workbooks. Browser clipboard denied mock, missing APIs simulated, real anchor downloads independently parsed. Narrow touch viewport is emulation; no physical iPad/iPhone/Android proof.'},null,2));}
}
console.log(JSON.stringify({checks,cases:results.length,passed:results.filter(r=>r.ok).length,errors,writes,failures:results.filter(r=>!r.ok)}));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;
