// Synthetic IDB restore -> document replacement retention diagnostic. No original files or remote writes.
// Forced Chromium GC is a diagnostic only; it does not certify iPad memory behavior.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5191/');
if (!['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)) throw Error('Local source or compiled server required.');
const out = process.env.WIXEL_RELOADED_RELEASE_OUT || 'D:/Codex/Temp/wixel-reloaded-document-release';
const sheetCount = Math.max(2, Math.min(15, Number(process.env.WIXEL_RELOADED_RELEASE_SHEETS || 15)));
const repeat = Math.max(1, Math.min(5, Number(process.env.WIXEL_RELOADED_RELEASE_REPEAT || 1)));
const KEY = 'tabula.workbook.v1', JOURNAL = 'wixel.document-recovery.v1';
await mkdir(out, { recursive:true });
let checks=0, stage='start';
const eq=(a,b,message)=>{checks++;assert.deepEqual(a,b,message);};
const errors=[], writes=[], results=[];
const browser=await chromium.launch({headless:true,args:['--enable-precise-memory-info']});
const context=await browser.newContext({viewport:{width:1440,height:1000},serviceWorkers:'block'});
const page=await context.newPage();
page.setDefaultTimeout(30000);
page.on('pageerror',e=>errors.push({stage,message:e.message}));
page.on('dialog',d=>d.accept());
await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
await context.route('**/*',route=>{
  const request=route.request(), target=new URL(request.url());
  if(!['GET','HEAD'].includes(request.method())){writes.push(request.method()+' '+target.pathname);return route.abort();}
  if(target.origin!==url.origin||target.pathname.startsWith('/api/'))return route.abort();
  return route.continue();
});
function fixture(name){
  return {name:name+'.wixel',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({docName:name,workbook:{
    calculation:{mode:'auto'},
    sheets:Array.from({length:sheetCount},(_,i)=>({name:name+' '+(i+1),
      cells:{'0,0':{raw:'Region'},'0,1':{raw:'Sales'},'1,0':{raw:'Seoul'},'1,1':{raw:'100'},'2,0':{raw:'Busan'},'2,1':{raw:'200'},'3,1':{raw:'=SUM(B2:B3)'}},
      ...(i===0?{cellRuns:[[4,0,60001,{raw:'',style:{numFmt:'number'}}]]}:{}),
      charts:[{id:name+'-chart-'+i,type:'column',x:300,y:30,w:400,h:280,range:{r1:0,c1:0,r2:2,c2:1},title:name+' chart '+(i+1)}],
    })),
  }}))};
}
// Return a primitive: an undisposed Playwright JSHandle to Workbook would itself keep the restored book alive.
async function ready(){await page.waitForFunction(()=>!!window.tabula?.wb());}
async function open(name){
  await page.locator('#fileInput').setInputFiles(fixture(name));
  await page.waitForFunction(name=>window.tabula?.wb().sheets[0].name===name+' 1'&&!document.querySelector('.load-progress'),name);
}
async function saved(name){
  const deadline=Date.now()+30000;
  while(Date.now()<deadline){
    const complete=await page.evaluate(async({KEY,JOURNAL,name})=>{
    const marker=JSON.parse(sessionStorage.getItem(JOURNAL)||'null');
    const pointer=JSON.parse(sessionStorage.getItem(KEY)||'null');
    if(marker?.phase!=='saved'||marker.docName!==name||!pointer?.idb||pointer.docId!==marker.docId||pointer.generation!==marker.generation)return false;
    const manifest=await new Promise((resolve,reject)=>{const request=indexedDB.open('tabula',1);request.onerror=()=>reject(request.error);request.onsuccess=()=>{const db=request.result,tx=db.transaction('docs'),read=tx.objectStore('docs').get(pointer.storageKey||KEY);read.onsuccess=()=>{db.close();resolve(read.result);};read.onerror=()=>{db.close();reject(read.error);};};});
    return manifest?.docId===pointer.docId&&manifest.generation===marker.generation;
    },{KEY,JOURNAL,name});
    if(complete)return;
    await page.waitForTimeout(250);
  }
  throw Error('Complete IDB checkpoint timeout: '+name);
}
async function diagnostics(){return page.evaluate(({KEY,JOURNAL})=>{
  const book=window.tabula?.wb(),view=window.tabula?.gv(),previous=window.__previousBookWeak?.deref();
  const directViewReferences=[];
  if(previous&&view)for(const key of Object.keys(view)){const value=view[key];if(value===previous||value?.wb===previous||value?.book===previous)directViewReferences.push(key);}
  return {title:document.title,sheets:book?.sheets.length,cells:book?.sheets.reduce((n,s)=>n+s.cells.size,0),
    value:book?.getRaw(0,0,0),sum:book?.getValue(0,3,1),heap:performance.memory?.usedJSHeapSize??null,
    previousRetained:!!previous,directViewReferences,
    viewCacheCurrent:view?._objectRenderState?.wb===book,
    marker:JSON.parse(sessionStorage.getItem(JOURNAL)||'null'),pointer:JSON.parse(sessionStorage.getItem(KEY)||'null'),
    status:document.querySelector('#saveState')?.textContent,dialogs:[...document.querySelectorAll('#dialogLayer .dialog')].map(e=>e.textContent),
  };
},{KEY,JOURNAL});}
let cdp;
try{
  await page.goto(url.href,{waitUntil:'domcontentloaded'});await ready();
  cdp=await context.newCDPSession(page);
  const asset=await page.evaluate(()=>[...document.scripts].map(s=>s.getAttribute('src')).find(s=>s?.includes('wixel-'))||'source');
  for(let run=0;run<repeat;run++){
    const a='Reload source A '+run,b='Destination B '+run,record={run,asset,ok:false};results.push(record);
    stage='A-open';await open(a);
    stage='A-first-save';await saved(a);
    const aPointer=await page.evaluate(KEY=>JSON.parse(sessionStorage.getItem(KEY)),KEY);
    eq(aPointer.idb,true,'Source uses real large-workbook IDB storage');
    stage='A-reload';await page.reload({waitUntil:'domcontentloaded'});await ready();
    eq(await page.evaluate(()=>tabula.wb().sheets.length),sheetCount,'All source sheets restored');
    eq(await page.evaluate(()=>tabula.wb().sheets[0].name),a+' 1','Source restored through init');
    eq(await page.evaluate(()=>!!document.querySelector('[data-document-recovery]')),false,'Complete source restoration has no recovery prompt');
    eq(await page.evaluate(KEY=>JSON.parse(sessionStorage.getItem(KEY)).docId,KEY),aPointer.docId,'Restored source keeps document identity');
    await page.evaluate(marker=>{
      const book=tabula.wb();book.__releaseAudit=marker;window.__previousBookWeak=new WeakRef(book);
    },'reloaded-source-workbook-'+run);
    stage='B-open';await open(b);
    eq(await page.evaluate(()=>tabula.wb().sheets.length),sheetCount,'Destination has the same sheet count');
    eq(await page.evaluate(()=>tabula.wb().sheets.every(s=>s.charts.length===1)),true,'Every destination sheet has a chart');
    stage='B-first-save';await saved(b);
    const bPointer=await page.evaluate(KEY=>JSON.parse(sessionStorage.getItem(KEY)),KEY);
    eq(bPointer.docId!==aPointer.docId,true,'Source and destination have separate document IDs');
    record.beforeGC=await diagnostics();
    // The document-open graph timer runs at 1200 ms. Keep finite async work out of the retention diagnosis.
    await page.waitForTimeout(1800);
    stage='B-GC';
    for(let n=0;n<3;n++){
      await cdp.send('HeapProfiler.collectGarbage');
      await page.evaluate(()=>new Promise(resolve=>setTimeout(resolve,0)));
    }
    record.afterGC=await diagnostics();
    if(record.afterGC.previousRetained&&process.env.WIXEL_RELOADED_RELEASE_HEAP==='1'){
      // diagnostics() dereferences WeakRef only inside its completed evaluation job.
      await page.evaluate(()=>new Promise(resolve=>setTimeout(resolve,0)));
      await cdp.send('HeapProfiler.collectGarbage');
      const chunks=[],listener=event=>chunks.push(event.chunk);
      cdp.on('HeapProfiler.addHeapSnapshotChunk',listener);
      try{await cdp.send('HeapProfiler.takeHeapSnapshot',{reportProgress:false});}
      finally{cdp.off('HeapProfiler.addHeapSnapshotChunk',listener);}
      record.snapshot=out+'/reloaded-source-'+run+'.heapsnapshot';await writeFile(record.snapshot,chunks.join(''));
    }
    eq(record.afterGC.previousRetained,false,'A restored through init is collectible after B is opened and saved');
    stage='B-edit';await page.evaluate(()=>tabula.wb().transact(()=>tabula.wb().setInput(0,1,1,'250')));
    eq(await page.evaluate(()=>tabula.wb().getRaw(0,1,1)),'250','Destination remains editable');
    eq(errors,[],'No application errors');eq(writes,[],'No remote writes');record.ok=true;
  }
}catch(error){results.push({ok:false,stage,error:error.stack,state:await diagnostics().catch(()=>null)});}
finally{await cdp?.detach().catch(()=>{});await context.close();await browser.close();}
const result={url:url.href,checks,stage,results,errors,writes};
await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify(result,null,2));
if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;
