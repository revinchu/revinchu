// Synthetic documents only. Real app/writer/storage code; HTTP revision/response timing is controlled.
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createWixelServer } from '../server/app.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const out = process.env.WIXEL_SAVE_GUARDS_OUT || 'D:/Codex/Temp/wixel-export-repair/online-save-guards';
await mkdir(out, {recursive:true});
const backend = createWixelServer({host:'127.0.0.1',data:await mkdtemp(join(out,'server-'))});
await new Promise(resolve => backend.listen(0,'127.0.0.1',resolve));
const origin = `http://127.0.0.1:${backend.address().port}`;
const hooks = `
window.__saveGuard={server,saveNow,autoRepublish,
 state:()=>({name:docName,id:docId,value:wb.getValue(0,0,0),dirty,error:serverState.error,saving:serverState.saving,pub:pubInfo()}),
 local:(name='Race',value='initial')=>{loadWorkbook({sheets:[{name:'Sheet',cells:{'0,0':{raw:value},'1,0':{raw:'=1+2'}}}]},name,0);remoteDoc=true;autosave=true;dirty=true;serverState.error=null;setPubInfo(null);},
 edit:value=>wb.transact(()=>wb.setInput(0,0,0,value)),
 pub:id=>setPubInfo({id,auto:true,opt:{sheet:'all',grid:true}}),
 pubAuto:enabled=>setPubInfo({...pubInfo(),auto:enabled}),
 clearError:()=>{serverState.error=null;},
 stop:()=>{clearTimeout(saveTimer);clearTimeout(serverTimer);clearTimeout(libTimer);clearInterval(collabTimer);stopPublishedWatch();}
};
window.__snapshotRace=book=>{const action=window.__race;if(!action)return;window.__race=null;queueMicrotask(()=>{if(action==='switch')__saveGuard.local('Other','new-document');else book.transact(()=>book.setInput(0,0,0,action));});};`;
let checks=0;const results=[];
const eq=(a,b,label)=>{checks++;assert.deepEqual(a,b,label);};
try {
 for(const engine of (process.env.WIXEL_BROWSER||'chromium,webkit').split(',')) {
  const browser=await pw[engine].launch(),context=await browser.newContext({serviceWorkers:'block'}),page=await context.newPage(),errors=[];
  page.on('pageerror',error=>errors.push(error.message));
  const stored=new Map(),puts=[],revisions=new Map();let delayNext=null,ready,release,failNext=null;
  await context.addInitScript(()=>{const nativeFetch=window.fetch;window.__requestBodies={};window.fetch=async(input,init)=>{if(init?.body instanceof Blob)window.__requestBodies[new URL(String(input),location.href).pathname]=await init.body.text();return nativeFetch(input,init);};window.WIXEL_SKIP_START=true;window.__retryCount=0;const timer=window.setTimeout;window.setTimeout=function(fn,ms,...args){if(ms===1500){window.__retry=()=>fn(...args);window.__retryCount++;return timer.call(this,fn,600000,...args);}return timer.call(this,fn,ms,...args);};});
  await context.route('**/*',async route=>{
   const request=route.request(),url=new URL(request.url()),path=url.pathname;
   if(url.origin!==origin)return route.abort();
   if(path==='/src/app.js'){const response=await route.fetch();return route.fulfill({response,body:await response.text()+hooks});}
   if(path==='/src/wixel-file.js'){const response=await route.fetch(),source=await response.text(),needle='const version=book.version, sheets=';assert.ok(source.includes(needle));return route.fulfill({response,body:source.replace(needle,'window.__snapshotRace?.(book); '+needle)});}
   if(!path.startsWith('/api/'))return route.continue();
   if(path==='/api/health')return route.fulfill({json:{ok:true,vault:true,auth:false,publish:true,maxDocumentBytes:32*1024*1024}});
   const key=path==='/api/publish'?'/api/published/synthetic-publication':path;
   if(request.method()==='PUT'||request.method()==='POST') {
    if(failNext===key){failNext=null;return route.fulfill({status:503,json:{error:'synthetic network failure'}});}
    const expected=String(revisions.get(key)||0),match=request.headers()['if-match'];
    if(request.method()==='PUT'&&match!==`"${expected}"`)return route.fulfill({status:412,json:{error:'synthetic revision conflict',code:'REVISION_CONFLICT'}});
    const revision=Number(expected)+1;revisions.set(key,revision);stored.set(key,JSON.parse(request.postData()??await page.evaluate(path=>window.__requestBodies[path],path)));puts.push({key,revision,match});
    if(delayNext===key){delayNext=null;ready();await new Promise(resolve=>release=resolve);}
    return route.fulfill({headers:{'X-Wixel-Revision':String(revision)},json:{ok:true,revision,modified:Date.now(),...(path==='/api/publish'?{id:'synthetic-publication'}:{})}});
   }
   if(path==='/api/files')return route.fulfill({json:[]});
   if(stored.has(key))return route.fulfill({headers:{'X-Wixel-Revision':String(revisions.get(key)),'X-Modified':String(revisions.get(key))},json:stored.get(key)});
   return route.fulfill({status:404,json:{error:'synthetic not found'}});
  });
  const state=()=>page.evaluate(()=>__saveGuard.state()),beginSave=()=>page.evaluate(()=>{window.__done=false;__saveGuard.saveNow(false).finally(()=>window.__done=true);});
  const delayed=path=>{delayNext=path;return new Promise(resolve=>ready=resolve);};
  const start=checks;
  try {
   await page.goto(origin);await page.waitForFunction(()=>window.__saveGuard&&window.tabula);
   await page.evaluate(async()=>{await __saveGuard.server.init();__saveGuard.server.connect('A'.repeat(43));__saveGuard.local();window.__race='edited-during-snapshot';});
   await page.evaluate(()=>__saveGuard.saveNow(false));let s=await state();
   eq([s.dirty,s.error,s.saving],[true,null,false],'Snapshot edit cancellation leaves autosave enabled and dirty');eq(puts.length,0,'Cancelled snapshot never uploads');
   const retries=await page.evaluate(()=>__retryCount);assert.ok(retries>0&&retries<10);checks++;
   await page.evaluate(()=>__retry());s=await state();eq([s.dirty,s.error],[false,null],'Debounced retry completes normally');
   eq(await page.evaluate(async()=>(await __saveGuard.server.load('Race')).workbook.sheets[0].cells.get('0,0').raw),'edited-during-snapshot','Retry stores latest edit');

   await page.evaluate(()=>__saveGuard.edit('sent-before-response'));const uploadReady=delayed('/api/files/Race');await beginSave();await uploadReady;
   await page.evaluate(()=>__saveGuard.edit('edited-after-server-commit'));release();await page.waitForFunction(()=>__done);s=await state();
   eq([s.dirty,s.error],[true,null],'Edit after server commit stays dirty without losing success');eq(await page.evaluate(()=>__saveGuard.server.revision('files/Race')),'2','Committed upload revision is received after an edit');
   await page.evaluate(()=>__retry());eq(puts.at(-1).match,'"2"','Next upload uses committed revision');eq((await state()).error,null,'Next upload has no revision conflict');

   const pub=await page.evaluate(async()=>__saveGuard.server.publish({workbook:{sheets:[]}}));await page.evaluate(id=>{__saveGuard.pub(id);window.__race='edited-during-publish';},pub.id);
   await page.evaluate(()=>__saveGuard.autoRepublish());eq((await state()).pub.auto,true,'Snapshot edit never disables automatic publication');eq(revisions.get('/api/published/'+pub.id),1,'Cancelled publication never uploads');
   await page.evaluate(()=>__saveGuard.autoRepublish());eq(revisions.get('/api/published/'+pub.id),2,'Next stable publication succeeds');
   const pubReady=delayed('/api/published/'+pub.id);await page.evaluate(()=>{window.__done=false;__saveGuard.autoRepublish().finally(()=>window.__done=true);});await pubReady;
   await page.evaluate(()=>__saveGuard.edit('edited-after-publication-commit'));release();await page.waitForFunction(()=>__done);
   eq(await page.evaluate(id=>__saveGuard.server.revision('published/'+id),pub.id),'3','Committed publication revision survives edit during response');eq((await state()).pub.auto,true,'Upload-time edit preserves automatic publication');
   await page.evaluate(()=>__saveGuard.autoRepublish());eq(puts.at(-1).match,'"3"','Next publication uses acknowledged revision');

   await page.evaluate(()=>__saveGuard.edit('first-queued-publish'));const batchReady=delayed('/api/published/'+pub.id),beforeBatch=puts.length;
   await page.evaluate(()=>{window.__batchDone=false;__saveGuard.autoRepublish().finally(()=>window.__batchDone=true);});await batchReady;
   await page.evaluate(()=>{__saveGuard.edit('latest-queued-publish');window.__queuedDone=Promise.all([__saveGuard.autoRepublish(),__saveGuard.autoRepublish(),__saveGuard.autoRepublish()]);});
   await page.waitForTimeout(100);eq((await state()).pub.auto,true,'Overlapping requests cannot disable automatic publication');eq(puts.length,beforeBatch+1,'One publication PUT is in flight');
   release();await page.waitForFunction(()=>__batchDone);await page.evaluate(()=>__queuedDone);
   eq(puts.length,beforeBatch+2,'Overlapping requests coalesce into exactly one latest snapshot');eq(revisions.get('/api/published/'+pub.id),6,'Queued publication acknowledges each revision once');
   eq(await page.evaluate(async id=>(await __saveGuard.server.published(id)).data.workbook.sheets[0].cells.get('0,0').raw,pub.id),'latest-queued-publish','Queued publication stores the latest edit');

   const disableReady=delayed('/api/published/'+pub.id),beforeDisable=puts.length;
   await page.evaluate(()=>{window.__disabledDone=false;__saveGuard.autoRepublish().finally(()=>window.__disabledDone=true);});await disableReady;
   await page.evaluate(()=>{__saveGuard.autoRepublish();__saveGuard.pubAuto(false);});release();await page.waitForFunction(()=>__disabledDone);
   eq(puts.length,beforeDisable+1,'Disabling automatic publication cancels its queued request');eq(await page.evaluate(id=>__saveGuard.server.revision('published/'+id),pub.id),'7','Disabling keeps the already committed server revision');
   await page.evaluate(id=>__saveGuard.pub(id),pub.id);await page.evaluate(()=>__saveGuard.autoRepublish());eq((await state()).pub.auto,true,'Re-enabling publication does not cause a stale revision conflict');

   failNext='/api/files/Race';await page.evaluate(()=>__saveGuard.saveNow(false));eq((await state()).error,'synthetic network failure','Real upload failure remains visible');await page.evaluate(()=>__saveGuard.clearError());
   failNext='/api/published/'+pub.id;await page.evaluate(()=>__saveGuard.autoRepublish());eq((await state()).pub.auto,false,'Real publication failure retains existing failure policy');
   await page.evaluate(id=>__saveGuard.pub(id),pub.id);const switchReady=delayed('/api/published/'+pub.id),beforeSwitch=puts.length;
   await page.evaluate(()=>{window.__pubSwitchDone=false;__saveGuard.autoRepublish().finally(()=>window.__pubSwitchDone=true);});await switchReady;
   await page.evaluate(()=>{__saveGuard.autoRepublish();__saveGuard.local('PublicationSwitch','unrelated');});release();await page.waitForFunction(()=>__pubSwitchDone);
   eq(puts.length,beforeSwitch+1,'Document replacement cancels the old queued publication');eq([(await state()).name,(await state()).pub],['PublicationSwitch',null],'Old publication leaves the new document unchanged');
   await page.evaluate(()=>{__saveGuard.local('Race','before-switch');window.__race='switch';});await page.evaluate(()=>__saveGuard.saveNow(false));s=await state();eq([s.name,s.value,s.error],['Other','new-document',null],'Old snapshot cancellation never changes the new document');
   await page.evaluate(()=>__saveGuard.local('OldUpload','old'));const oldReady=delayed('/api/files/OldUpload');await page.evaluate(()=>{window.__oldDone=false;__saveGuard.saveNow(false).finally(()=>window.__oldDone=true);});await oldReady;const releaseOld=release;
   await page.evaluate(()=>__saveGuard.local('NewUpload','new'));const newReady=delayed('/api/files/NewUpload');await beginSave();await newReady;const releaseNew=release;
   releaseOld();await page.waitForFunction(()=>__oldDone);s=await state();eq([s.name,s.saving,s.error],['NewUpload',true,null],'Old upload completion cannot unlock the new document save');
   const sent=puts.length;await page.evaluate(()=>__saveGuard.saveNow(false));eq(puts.length,sent,'New document save remains single-flight');
   releaseNew();await page.waitForFunction(()=>__done);s=await state();eq([s.name,s.saving,s.dirty,s.error],['NewUpload',false,false,null],'Only the new save can complete its status');
   eq(errors,[],'No page errors');results.push({engine,ok:true,checks:checks-start});
  }catch(error){results.push({engine,ok:false,error:error.stack,checks:checks-start,state:await state().catch(()=>null),errors});}
  finally{release?.();await page.evaluate(()=>window.__saveGuard?.stop()).catch(()=>{});await browser.close();}
 }
}finally{backend.closeAllConnections?.();await new Promise(resolve=>backend.close(resolve));}
const report={ok:results.every(item=>item.ok),checks,results};await writeFile(join(out,'result.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report));if(!report.ok)process.exitCode=1;
