// Read-only originals; browser edits stay in an isolated local browser profile.
import assert from 'node:assert/strict';
import { readFile,writeFile,mkdir,stat } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname))throw Error('로컬 검사 서버가 필요합니다.');
const selected=(process.env.WIXEL_REAL_AUDIT_IDS||'').split(',').filter(Boolean);
const files=JSON.parse((await readFile(process.env.WIXEL_REAL_AUDIT_MANIFEST,'utf8')).replace(/^\uFEFF/,'')).filter(f=>!selected.length||selected.includes(f.id));
const out=process.env.WIXEL_ROLLBACK_OUT||'D:/Codex/Temp/wixel-xlsb-rollback/stress';
await mkdir(out,{recursive:true});
const engines=(process.env.WIXEL_BROWSER||'chromium,webkit').split(','),timeout=Number(process.env.WIXEL_ROLLBACK_TIMEOUT||240000);
const isolated=process.env.WIXEL_ISOLATED_FILES==='1';
const runs=engines.flatMap(engine=>isolated?files.map(file=>({engine,inputs:[file]})):[{engine,inputs:files}]);
const heapMb=Number(process.env.WIXEL_CHROMIUM_HEAP_MB||0);
if(heapMb && (!Number.isInteger(heapMb)||heapMb<1024))throw Error('WIXEL_CHROMIUM_HEAP_MB must be >= 1024');
const results=[],errors=[],writes=[],assets=new Set();let activeStage='',checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);};
const save=()=>writeFile(out+'/result.json',JSON.stringify({checks,results,errors,writes,assets:[...assets],activeStage,chromiumHeapMb:heapMb||null,isolatedFiles:isolated},null,2));
for(const {engine,inputs} of runs){
 let browser,context,p,minFree=os.freemem(),memoryTimer;
 try{
  // A fresh D: persistent profile also supports Windows WebKit's native IDB Blob backing.
  const profile=path.join(out,'profiles',engine+'-'+Date.now());
  context=await pw[engine].launchPersistentContext(profile,{headless:true,viewport:{width:1180,height:820},serviceWorkers:'block',args:engine==='chromium'?['--enable-precise-memory-info',...(heapMb?['--js-flags=--max-old-space-size='+heapMb]:[])]:[]});
  browser=context.browser();
  await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
  await context.route('**/*',route=>{const q=route.request(),u=new URL(q.url());if(!['GET','HEAD'].includes(q.method())){writes.push({method:q.method(),path:u.pathname});return route.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return route.abort();return route.continue();});
  p=context.pages()[0]||await context.newPage();p.setDefaultTimeout(timeout);p.on('dialog',d=>d.accept());p.on('pageerror',e=>errors.push({engine,message:e.message,stage:activeStage}));
  memoryTimer=setInterval(()=>{minFree=Math.min(minFree,os.freemem());if(minFree<1024**3){activeStage+='-host-memory-guard';context.close().catch(()=>{});}},500);
  await p.goto(url);await p.waitForFunction(()=>!!window.tabula?.wb());
  assets.add(await p.evaluate(()=>[...document.scripts].map(x=>x.getAttribute('src')).find(x=>x?.includes('wixel-'))||'source'));
  for(const file of inputs){
   const record={engine,id:file.id,ok:false,timings:{},checks:0};results.push(record);const start=Date.now(),checkAt=checks;
   const before=await stat(file.path);activeStage=engine+'-'+file.id+'-open';await save();
   await p.evaluate(()=>{window.__priorDoc=tabula.wb();});
   // Windows WebKit persistent profiles cannot always read disk-backed File objects.
   const input=engine==='webkit'?{name:path.basename(file.path),mimeType:'application/vnd.ms-excel.sheet.binary.macroEnabled.12',buffer:await readFile(file.path)}:file.path;
   await p.locator('#fileInput').setInputFiles(input);
   await p.waitForFunction(()=>!document.querySelector('.load-progress')&&(tabula.wb()!==window.__priorDoc||document.querySelector('#dialogLayer .dialog')));
   if(await p.evaluate(()=>tabula.wb()===window.__priorDoc))throw Error(await p.locator('#dialogLayer .dialog').innerText());
   await p.evaluate(()=>{window.__priorDoc=null;window.__stressBook=tabula.wb();});
   for(let n=0;n<8&&await p.locator('#dialogLayer .dialog').count();n++)await p.keyboard.press('Escape');
   record.timings.openMs=Date.now()-start;eq((await p.title()).includes(path.basename(file.path).replace(/\.[^.]+$/,'')),true,'문서 제목');
   const identity=await p.evaluate(()=>({count:tabula.wb().sheets.length,formulas:tabula.wb().sheets.reduce((n,s)=>{s.cells.forEachStoredRC(cell=>{if(cell.formula)n++;});return n;},0),heap:performance.memory?.usedJSHeapSize??null}));record.identity=identity;
   const checkpoint=async()=>{const deadline=Date.now()+timeout;let poll=0;while(Date.now()<deadline){const complete=await p.evaluate(async()=>{const pointer=JSON.parse(sessionStorage.getItem('tabula.workbook.v1')||localStorage.getItem('tabula.workbook.v1')||'null');if(!pointer?.idb)return false;const manifest=await new Promise((resolve,reject)=>{const q=indexedDB.open('tabula',1);q.onerror=()=>reject(q.error);q.onsuccess=()=>{const db=q.result,r=db.transaction('docs').objectStore('docs').get(pointer.storageKey||'tabula.workbook.v1');r.onsuccess=()=>{db.close();resolve(r.result);};r.onerror=()=>{db.close();reject(r.error);};};});const marker=JSON.parse(sessionStorage.getItem('wixel.document-recovery.v1')||'null');return manifest?.docName===document.title.replace(/ - WIXEL$/,'')&&pointer.docId===manifest.docId&&(!marker||marker.phase==='saved');});if(complete)return;if(poll++%20===0){record.checkpointState=await p.evaluate(()=>({phase:JSON.parse(sessionStorage.getItem('wixel.document-recovery.v1')||'null')?.phase,status:document.querySelector('#saveState')?.textContent}));await save();if(record.checkpointState.phase==='failed')throw Error(record.checkpointState.status||'브라우저 저장 실패');}await p.waitForTimeout(500);}throw Error('완료 저장 checkpoint 대기 시간 초과');};
   activeStage=engine+'-'+file.id+'-initial-checkpoint';await save();await checkpoint();record.timings.initialCheckpointMs=Date.now()-start;
   const edit=await p.evaluate(()=>{const w=tabula.wb();const si=w.sheets.findIndex(s=>s.state!=='hidden'&&s.state!=='veryHidden'&&!s.protect?.on);if(si<0)throw Error('편집 가능한 시트 없음');const u=w.usedRange(si),r=Math.min(1048500,u.rows+3),c=0;tabula.switchSheet(si);tabula.selectCell(r,c);return{si,r,c};});
   const value='WIXEL_RECOVERY_'+file.id;
   activeStage=engine+'-'+file.id+'-edit';await p.locator('#cellEditor').focus();await p.keyboard.press('F2');await p.locator('#cellEditor').fill(value);await p.keyboard.press('Enter');
   eq(await p.evaluate(e=>tabula.wb().getRaw(e.si,e.r,e.c),edit),value,'편집 유지');
   eq(await p.evaluate(()=>tabula.wb()===window.__stressBook),true,'편집 후 같은 문서');
   activeStage=engine+'-'+file.id+'-edited-checkpoint';await save();await checkpoint();record.timings.editedCheckpointMs=Date.now()-start;
   const savedId=await p.evaluate(()=>JSON.parse(sessionStorage.getItem('tabula.workbook.v1')||localStorage.getItem('tabula.workbook.v1')).docId);
   activeStage=engine+'-'+file.id+'-reload';await save();await p.reload({waitUntil:'domcontentloaded'});await p.waitForFunction(()=>!!window.tabula?.wb());
   eq(await p.evaluate(e=>tabula.wb().getRaw(e.si,e.r,e.c),edit),value,'새로고침 후 최신 편집 유지');
   eq(await p.evaluate(()=>JSON.parse(sessionStorage.getItem('tabula.workbook.v1')||localStorage.getItem('tabula.workbook.v1')).docId),savedId,'새로고침 후 같은 저장 문서');
   eq(await p.evaluate(()=>tabula.wb().sheets.length),identity.count,'모든 시트 유지');
   eq(await p.evaluate(()=>tabula.wb().sheets.reduce((n,s)=>{s.cells.forEachStoredRC(cell=>{if(cell.formula)n++;});return n;},0)),identity.formulas,'모든 수식 셀 개수 유지');
   eq((await p.title()).includes(path.basename(file.path).replace(/\.[^.]+$/,'')),true,'새로고침 후 다른 파일로 복귀하지 않음');
   const after=await stat(file.path);eq([after.size,after.mtimeMs],[before.size,before.mtimeMs],'입력 파일 불변');
   record.ok=true;record.timings.totalMs=Date.now()-start;record.checks=checks-checkAt;record.minimumFree=minFree;
   console.log(JSON.stringify({engine,id:file.id,ok:true,timing:record.timings}));await save();
  }
 }catch(e){results.push({engine,ok:false,stage:activeStage,error:e.message,minimumFree:minFree});console.error(JSON.stringify({engine,stage:activeStage,error:e.message}));await p?.screenshot({path:out+'/'+engine+'-failed.png',timeout:10000}).catch(()=>{});}
 finally{clearInterval(memoryTimer);await context?.close().catch(()=>{});await browser?.close().catch(()=>{});await save();}
}
if(results.some(x=>!x.ok)||errors.length||writes.length)process.exitCode=1;
