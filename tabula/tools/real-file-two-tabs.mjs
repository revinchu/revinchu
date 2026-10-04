// 격리된 동일 origin 두 탭에서 실파일 A/B가 서로 덮어쓰이지 않는지 검사한다.
// file input으로 로컬에서만 읽으며 원본 저장·내보내기·API/외부 요청은 금지한다.
import assert from 'node:assert/strict';
import os from 'node:os';
import path from 'node:path';
import { readFile, writeFile, mkdir, stat } from 'node:fs/promises';
import { createHash } from 'node:crypto';
const { chromium }=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
if(!['localhost','127.0.0.1'].includes(new URL(url).hostname))throw Error('로컬 서버만 검사합니다.');
const out=process.env.WIXEL_REAL_TWO_TABS_OUT||'D:/Codex/Temp/wixel-final-audit/browser-two-tabs';
const manifest=JSON.parse((await readFile(process.env.WIXEL_REAL_AUDIT_MANIFEST||'D:/Codex/Temp/wixel-final-audit/manifest.json','utf8')).replace(/^\uFEFF/,''));
const ids=(process.env.WIXEL_REAL_TWO_TABS_IDS||'F20,F19').split(',');
if(ids.length!==2||new Set(ids).size!==2)throw Error('서로 다른 파일 ID 두 개가 필요합니다.');
const requireIdb=process.env.WIXEL_REAL_TWO_TABS_REQUIRE_IDB==='1';
const files=ids.map(id=>{const item=manifest.find(f=>f.id===id);if(!item)throw Error('알 수 없는 파일 ID');return item;});
await mkdir(out,{recursive:true});
const report={started:new Date().toISOString(),url,ids,checks:0,stages:[],errors:[],blockedRequests:[],scriptHashes:{},memory:{freeBefore:os.freemem(),minFree:os.freemem()}};
const eq=(a,b,message)=>{report.checks++;assert.deepEqual(a,b,message);};
const raf=p=>p.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
const nameOf=item=>path.basename(item.path).replace(/\.[^.]+$/,'');
async function savedState(p,item){
 await p.waitForFunction(name=>{
  const marker=JSON.parse(sessionStorage.getItem('wixel.document-recovery.v1')||'null');
  const pointer=JSON.parse(sessionStorage.getItem('tabula.workbook.v1')||'null');
  return marker?.phase==='saved'&&marker.docName===name&&marker.version>=tabula.wb().version&&marker.docId===pointer?.docId&&marker.generation===pointer?.generation;
 },nameOf(item),{timeout:180000});
 const saved=await p.evaluate(async()=>{
  const marker=JSON.parse(sessionStorage.getItem('wixel.document-recovery.v1'));
  const pointer=JSON.parse(sessionStorage.getItem('tabula.workbook.v1'));
  let stored=pointer;
  if(pointer.idb)stored=await new Promise((resolve,reject)=>{
   const open=indexedDB.open('tabula',1);open.onerror=()=>reject(open.error);
   open.onsuccess=()=>{const db=open.result,tx=db.transaction('docs','readonly'),request=tx.objectStore('docs').get(pointer.storageKey);let value;
    request.onsuccess=()=>{value=request.result;};tx.oncomplete=()=>{db.close();resolve(value);};tx.onerror=()=>{db.close();reject(tx.error);};};
  });
  return{docId:marker.docId,generation:marker.generation,version:marker.version,idb:!!pointer.idb,storedId:stored?.docId,storedGeneration:stored?.generation,sheets:stored?.sheets?.length??stored?.workbook?.sheets?.length};
 });
 eq(saved.storedId,saved.docId,item.id+' committed 저장 문서 식별자');
 eq(saved.storedGeneration,saved.generation,item.id+' committed 저장 세대');
 if(requireIdb)eq(saved.idb,true,item.id+' 실제 IndexedDB 저장');
 eq(saved.sheets,await p.evaluate(()=>tabula.wb().sheets.length),item.id+' 저장 시트 수');
 return saved;
}
let browser,memoryAbort=false;
const guard=setInterval(()=>{report.memory.minFree=Math.min(report.memory.minFree,os.freemem());if(report.memory.minFree<1024**3&&!memoryAbort){memoryAbort=true;browser?.close().catch(()=>{});}},1000);
try{
 browser=await chromium.launch({args:['--enable-precise-memory-info']});
 const context=await browser.newContext({viewport:{width:1440,height:920},serviceWorkers:'block'});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1','off');});
 await context.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(!['GET','HEAD'].includes(req.method())||u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')){report.blockedRequests.push({method:req.method(),kind:u.pathname.startsWith('/api/')?'api':'external'});return route.abort();}return route.continue();});
 const tabs=[],originals=[],scriptReads=[],positions=[],saved=[];
 for(const item of files){
  originals.push(await stat(item.path));const p=await context.newPage();tabs.push(p);p.setDefaultTimeout(600000);
  p.on('pageerror',e=>report.errors.push({id:item.id,error:e.message}));
  p.on('response',res=>{if(res.request().resourceType()==='script')scriptReads.push(res.body().then(bytes=>{report.scriptHashes[new URL(res.url()).pathname]=createHash('sha256').update(bytes).digest('hex');}).catch(()=>{}));});
  await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());
  await p.evaluate(()=>window.__oldAuditBook=tabula.wb());const start=Date.now();
  await p.locator('#fileInput').setInputFiles(item.path);
  await p.waitForFunction(()=>!document.querySelector('.load-progress')&&tabula.wb()!==window.__oldAuditBook);
  for(let i=0;i<8;i++){const d=p.locator('#dialogLayer .dialog').last();if(!await d.count())break;const close=d.locator('[data-dialog-close-head]');if(await close.count())await close.click();else await p.keyboard.press('Escape');}
  eq((await p.title()).includes(path.basename(item.path).replace(/\.[^.]+$/,'')),true,item.id+' 제목');
  eq(await p.evaluate(()=>tabula.wb().undoStack.length),0,item.id+' 이전 문서 undo 없음');
  const position=await p.evaluate(()=>{const w=tabula.wb();window.__auditBook=w;window.__oldAuditBook=null;const i=w.sheets.findIndex(s=>s.state!=='hidden'&&s.state!=='veryHidden'&&!s.protect?.on);if(i<0)throw Error('편집 가능한 시트 없음');tabula.switchSheet(i);const u=w.usedRange(i);for(let n=0;n<30;n++){const r=Math.min(1048500,u.rows+2+n),c=0;if(!w.getRaw(i,r,c)&&!w.mergeAt(i,r,c)){window.__auditPosition={i,r,c};tabula.selectCell(r,c);return{i,r,c};}}throw Error('안전한 빈 셀 없음');});
  positions.push(position);
  await p.locator('#cellEditor').focus();await p.keyboard.press('F2');await p.locator('#cellEditor').fill('WIXEL_TWO_TABS_'+item.id);await p.keyboard.press('Enter');await raf(p);
  eq(await p.evaluate(({i,r,c})=>tabula.wb().getRaw(i,r,c),position),'WIXEL_TWO_TABS_'+item.id,item.id+' 입력');
  report.stages.push({id:item.id,stage:'opened-edited',ms:Date.now()-start});console.log(JSON.stringify(report.stages.at(-1)));
 }
 for(let round=0;round<3;round++)for(let i=0;i<tabs.length;i++){
  const p=tabs[i],item=files[i];await p.bringToFront();await raf(p);
  eq(await p.evaluate(()=>tabula.wb()===window.__auditBook),true,item.id+' 탭 전환 후 문서 동일');
  eq((await p.title()).includes(path.basename(item.path).replace(/\.[^.]+$/,'')),true,item.id+' 탭 전환 후 제목');
  eq(await p.evaluate(()=>{const {i,r,c}=window.__auditPosition;return tabula.wb().getRaw(i,r,c);}),'WIXEL_TWO_TABS_'+item.id,item.id+' 탭 전환 후 값');
 }
 for(let i=0;i<tabs.length;i++){const start=Date.now();saved.push(await savedState(tabs[i],files[i]));report.stages.push({id:files[i].id,stage:'autosave-committed',ms:Date.now()-start,idb:saved[i].idb});}
 for(let i=0;i<tabs.length;i++){
  const p=tabs[i],item=files[i];await p.bringToFront();await p.evaluate(()=>tabula.run('undo'));await raf(p);
  eq(await p.evaluate(()=>{const {i,r,c}=window.__auditPosition;return tabula.wb().getRaw(i,r,c);}), '',item.id+' 고유 undo');
  await p.evaluate(()=>tabula.run('redo'));await raf(p);
  eq(await p.evaluate(()=>{const {i,r,c}=window.__auditPosition;return tabula.wb().getRaw(i,r,c);}), 'WIXEL_TWO_TABS_'+item.id,item.id+' 고유 redo');
  eq(await p.evaluate(()=>tabula.wb()===window.__auditBook),true,item.id+' 자동 저장 대기 후 문서 동일');
  const committed=await savedState(p,item),start=Date.now();
  eq(committed.docId,saved[i].docId,item.id+' undo/redo 후 저장 문서 식별자');
  await p.reload({waitUntil:'domcontentloaded'});
  await p.waitForFunction(name=>window.tabula?.wb()&&!document.querySelector('.load-progress')&&document.title.includes(name),nameOf(item));
  await raf(p);
  eq(await p.locator('[data-document-recovery]').count(),0,item.id+' 새로고침 복구 오류 없음');
  eq(await p.evaluate(({i,r,c})=>tabula.wb().getRaw(i,r,c),positions[i]),'WIXEL_TWO_TABS_'+item.id,item.id+' 새로고침 후 자기 문서 편집값 복원');
  const restored=await p.evaluate(()=>JSON.parse(sessionStorage.getItem('wixel.document-recovery.v1')));
  eq(restored.docId,committed.docId,item.id+' 새로고침 후 자기 문서 식별자');
  eq(await p.evaluate(()=>tabula.wb().undoStack.length),0,item.id+' 새로고침 후 다른 문서 undo 없음');
  report.stages.push({id:item.id,stage:'reloaded-own-snapshot',ms:Date.now()-start,idb:committed.idb});console.log(JSON.stringify(report.stages.at(-1)));
  const after=await stat(item.path);eq(after.size,originals[i].size,item.id+' 원본 크기 보존');eq(after.mtimeMs,originals[i].mtimeMs,item.id+' 원본 시각 보존');
 }
 eq(report.errors,[],'페이지 실행 오류 없음');await Promise.all(scriptReads);report.ok=true;
}catch(error){report.ok=false;report.resourceInconclusive=memoryAbort;report.error=error.message;process.exitCode=1;}
finally{clearInterval(guard);await browser?.close().catch(()=>{});report.memory.freeAfter=os.freemem();report.finished=new Date().toISOString();await writeFile(out+'/result.json',JSON.stringify(report,null,2));console.log(JSON.stringify({ok:report.ok,checks:report.checks,resourceInconclusive:report.resourceInconclusive,error:report.error,out}));}
