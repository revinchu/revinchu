// Synthetic documents only. Product Node HTTP/storage is real; hooks only delay or fail preparation.
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createWixelServer } from '../server/app.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const out = process.env.WIXEL_OPEN_GUARDS_OUT || 'D:/Codex/Temp/wixel-export-repair/open-guards';
await mkdir(out, { recursive:true });
const data = await mkdtemp(join(out, 'server-'));
const server = createWixelServer({ host:'127.0.0.1', data });
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
const fixture = (name, value, revision = '2') => ({ docName:name, revision, workbook:{ sheets:[{ name, cells:{ '0,0':{ raw:value } } }] } });
const put = async (name, value, revision = '2') => {
  const res = await fetch(`${origin}/api/files/${name}`, { method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify(fixture(name, value, revision)) });
  assert.equal(res.status, 200);
};
const instrumentation = `
window.__openGuard = {
 server, openFromServer, startCollabWatch,
 state:() => ({docName, docId, viewOnly, autosave, remoteDoc, dirty, version:wb.version, value:wb.getValue(0,0,0), bar:document.querySelector('.view-bar')?.textContent ?? null, headers:view.showHeaders, grid:view.showGrid}),
 local:(name='Local', value='local') => { stopPublishedWatch(); clearInterval(collabTimer); collabTimer=null; viewOnly=false; autosave=true; document.body.classList.remove('view-mode'); document.querySelector('.view-bar')?.remove(); loadWorkbook({sheets:[{name,cells:{'0,0':{raw:value}}}]}, name, 0); },
 enter:async data => { const request=beginDocumentOpen(); try { return await enterViewMode(data,{request}); } finally { finishDocumentOpen(request); } },
 edit:value => wb.transact(()=>wb.setInput(0,0,0,value)),
 pollReady:() => {dirty=false;serverState.savedAt=1;startCollabWatch();},
 stop:() => {clearInterval(collabTimer);clearTimeout(saveTimer);clearTimeout(serverTimer);clearTimeout(libTimer);stopPublishedWatch();}
};`;
const results=[];
let checks=0;
const eq=(actual, expected, label)=>{checks++;assert.deepEqual(actual,expected,label);};
try {
  for (const engine of (process.env.WIXEL_BROWSER || 'chromium,webkit').split(',')) {
    const browser=await pw[engine].launch(), context=await browser.newContext({serviceWorkers:'block'}), page=await context.newPage(), errors=[];
    page.on('pageerror', e=>errors.push(e.message));
    await context.addInitScript(()=>{
      window.WIXEL_SKIP_START=true;
      const interval=window.setInterval;
      window.setInterval=function(fn,ms,...args){if(ms===15000)window.__pollTick=()=>fn(...args);return interval.call(this,fn,ms===15000?600000:ms,...args);};
    });
    // The app module is extended with test accessors; API requests use the unmodified real backend.
    await context.route('**/*', async route=>{
      const u=new URL(route.request().url());
      if(u.origin!==origin)return route.abort();
      if(u.pathname==='/src/app.js') { const response=await route.fetch();return route.fulfill({response,body:await response.text()+instrumentation}); }
      return route.continue();
    });
    const state=()=>page.evaluate(()=>__openGuard.state());
    const gate=async(name, fail=false)=>page.evaluate(({name,fail})=>{
      const proto=tabula.wb().constructor.prototype, original=proto.loadAsync;
      window.__prep={entered:false,calls:0};
      proto.loadAsync=async function(data,...args){
        window.__prep.calls++;
        if(data.sheets?.[0]?.name===name){
          window.__prep.entered=true;
          if(fail)throw Error('synthetic preparation failure');
          await new Promise(resolve=>window.__releasePreparation=resolve);
        }
        return original.call(this,data,...args);
      };
      window.__restorePreparation=()=>{proto.loadAsync=original;};
    },{name,fail});
    const release=async()=>{await page.evaluate(()=>__releasePreparation());await page.waitForFunction(()=>window.__done);await page.evaluate(()=>__restorePreparation());};
    const begin=async(name, mode='server')=>page.evaluate(({name,mode})=>{
      window.__done=false;
      const work=mode==='server'?__openGuard.openFromServer(name):__openGuard.server.load(name).then(data=>__openGuard.enter({...data,view:{title:name,headers:false,grid:false}}));
      work.then(value=>{window.__outcome={value};},error=>{window.__outcome={error:error.message,code:error.code};}).finally(()=>window.__done=true);
    },{name,mode});
    const start=checks;
    try {
      await page.goto(origin);await page.waitForFunction(()=>window.__openGuard && window.tabula);
      await page.evaluate(async()=>{await __openGuard.server.init();__openGuard.server.connect();__openGuard.local();});
      await put('Broken','broken');await gate('Broken',true);
      const before=await state();await begin('Broken','view');await page.waitForFunction(()=>window.__done);await page.evaluate(()=>__restorePreparation());
      eq(await state(),before,'Failed published preparation preserves old document, edit mode and autosave');
      eq((await page.evaluate(()=>__outcome)).error,'synthetic preparation failure','Preparation error remains visible to caller');

      await put('View','published');await gate('View');await begin('View','view');await page.waitForFunction(()=>__prep.entered);
      eq(await state(),before,'Published mode is not applied while preparation is pending');
      await release();let s=await state();eq([s.docName,s.value,s.viewOnly,s.autosave,s.grid,s.headers],['View','published',true,false,false,false],'Successful published install commits view settings');
      await gate('Broken',true);const viewBefore=await state();await begin('Broken','view');await page.waitForFunction(()=>__done);await page.evaluate(()=>__restorePreparation());
      eq(await state(),viewBefore,'Failed replacement preserves prior published bar and view');

      await page.evaluate(()=>__openGuard.local());await gate('View');await begin('View','view');await page.waitForFunction(()=>__prep.entered);
      await page.evaluate(()=>__openGuard.local('NewLocal','new-local'));await release();s=await state();
      eq([s.docName,s.value,s.viewOnly,s.autosave,s.bar],['NewLocal','new-local',false,true,null],'Cancelled published read never changes newer document mode');

      await put('Remote','remote','2');await page.evaluate(()=>{__openGuard.local();__openGuard.server.restoreRevision('files/Remote','1');});
      await gate('Remote');await begin('Remote');await page.waitForFunction(()=>__prep.entered);
      eq((await state()).value,'local','Server document waits for asynchronous preparation');
      await page.evaluate(()=>__openGuard.edit('local-edit'));await release();s=await state();
      eq([s.docName,s.value,s.remoteDoc],['Local','local-edit',false],'Edit during server preparation cancels incoming replacement');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Remote')),'1','Aborted preparation restores only its uncommitted revision');
      eq((await page.evaluate(()=>__outcome)).value,false,'Cancelled server preparation returns false');

      await gate('Remote');await begin('Remote');await page.waitForFunction(()=>__prep.entered);await release();s=await state();
      eq([s.docName,s.value,s.remoteDoc],['Remote','remote',true],'Successful server load commits asynchronous document');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Remote')),'2','Successful server load retains incoming revision');

      // A real response is delayed before delivery; its JSON/body is never mocked or changed.
      await put('Delayed','delayed','9');await page.evaluate(()=>{
        __openGuard.local();__openGuard.server.restoreRevision('files/Delayed','8');
        const fetchOriginal=window.fetch;window.__responseReady=false;
        window.fetch=async(...args)=>{const response=await fetchOriginal(...args);if(String(args[0]).includes('files/Delayed')){window.__responseReady=true;await new Promise(resolve=>window.__releaseResponse=resolve);}return response;};
        window.__restoreFetch=()=>window.fetch=fetchOriginal;
      });
      await begin('Delayed');await page.waitForFunction(()=>__responseReady);await page.evaluate(()=>{__openGuard.edit('pending-edit');__releaseResponse();});await page.waitForFunction(()=>window.__done);await page.evaluate(()=>__restoreFetch());
      eq((await state()).value,'pending-edit','Edit during real server response is kept');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Delayed')),'8','Late response does not advance revision metadata');

      await begin('Remote');await page.waitForFunction(()=>window.__done);await put('Remote','poll-update','3');
      await gate('Remote');await page.evaluate(()=>{__openGuard.pollReady();window.__pollDone=false;__pollTick().finally(()=>window.__pollDone=true);});await page.waitForFunction(()=>__prep.entered);
      await page.evaluate(()=>__pollTick());eq(await page.evaluate(()=>__prep.calls),1,'Background watcher has at most one preparation in flight');
      await page.evaluate(()=>{__openGuard.edit('poll-local-edit');__releasePreparation();});await page.waitForFunction(()=>window.__pollDone);await page.evaluate(()=>__restorePreparation());
      eq((await state()).value,'poll-local-edit','Background preparation cannot replace local edits');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Remote')),'2','Cancelled background update preserves old save revision');

      const docId=(await state()).docId;await page.evaluate(()=>{__openGuard.pollReady();window.__pollDone=false;__pollTick().finally(()=>window.__pollDone=true);});await page.waitForFunction(()=>window.__pollDone);s=await state();
      eq([s.value,s.docId,s.remoteDoc],['poll-update',docId,true],'Successful background async reload preserves document identity');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Remote')),'3','Background success commits revision');

      await put('Remote','stale-poll','4');await gate('Remote');await page.evaluate(()=>{__openGuard.pollReady();window.__pollDone=false;__pollTick().finally(()=>window.__pollDone=true);});await page.waitForFunction(()=>__prep.entered);
      await page.evaluate(()=>{__openGuard.local('Latest','latest');__openGuard.server.restoreRevision('files/Remote','99');__releasePreparation();});await page.waitForFunction(()=>window.__pollDone);await page.evaluate(()=>__restorePreparation());
      eq((await state()).value,'latest','Old background preparation never replaces newer document');
      eq(await page.evaluate(()=>__openGuard.server.revision('files/Remote')),'99','Old watcher never rolls back newer revision');
      eq(errors,[],'No page errors');
      results.push({engine,ok:true,checks:checks-start});
    } catch(error) {results.push({engine,ok:false,error:error.stack,checks:checks-start,errors,state:await state().catch(()=>null)});}
    finally {await page.evaluate(()=>window.__openGuard?.stop()).catch(()=>{});await browser.close();}
  }
} finally {server.closeAllConnections?.();await new Promise(resolve=>server.close(resolve));}
const report={ok:results.every(x=>x.ok),checks,realLocalBackend:true,results};
await writeFile(join(out,'result.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report));
if(!report.ok)process.exitCode=1;
