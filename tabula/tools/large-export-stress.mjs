// Private real-file save/reopen regression. Only localhost GETs; originals read-only.
// Required: WIXEL_EXPORT_INPUT=absolute local source path. Optional:
// WIXEL_EXPORT_SHEET=sheet to edit (default: first sheet with pivot + slicer),
// WIXEL_EXPORT_OUT, WIXEL_EXPORT_TYPES=wixel,html,xlsx, WIXEL_BROWSER,
// WIXEL_EXPORT_MODE=export|integrity|reopen, WIXEL_EXPORT_BASELINE (reopen),
// WIXEL_URL (localhost only). Performance runs hash full content only after exports.
import assert from 'node:assert/strict';
import {mkdir,readFile,writeFile,stat} from 'node:fs/promises';
import {openAsBlob} from 'node:fs';
import {packPublishedBlob,packedPublishedSize,assertPublishLinkSize} from '../src/publish.js';
import path from 'node:path';
import os from 'node:os';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
if(!['localhost','127.0.0.1'].includes(new URL(url).hostname))throw Error('Local host required');
const source=process.env.WIXEL_EXPORT_INPUT;if(!source)throw Error('WIXEL_EXPORT_INPUT must name a private local workbook');
const editSheet=process.env.WIXEL_EXPORT_SHEET||null;
const diagnostics=process.env.WIXEL_EXPORT_DIAG==='1',dropGraph=process.env.WIXEL_EXPORT_DROP_GRAPH==='1';
const out=process.env.WIXEL_EXPORT_OUT||'D:/Codex/Temp/wixel-export-repair/chromium';
const engine=process.env.WIXEL_BROWSER||'chromium';const mode=process.env.WIXEL_EXPORT_MODE||'export';
await mkdir(out,{recursive:true});
const result={engine,mode,diagnostics,dropGraph,fingerprintVersion:3,stage:'start',errors:[],exports:[],initialFree:os.freemem(),minimumFree:os.freemem()};
let reportWrites=Promise.resolve();
const save=()=>{const text=JSON.stringify(result,null,2),write=()=>writeFile(out+'/'+mode+'.json',text);reportWrites=reportWrites.then(write,write);return reportWrites;};
let context,timer,page,probeTimer,probing=false;
try{
 if(process.env.WIXEL_EXPORT_MIN_START_FREE&&os.freemem()<Number(process.env.WIXEL_EXPORT_MIN_START_FREE))throw Error('Insufficient host RAM to start a meaningful stress measurement');
 context=await pw[engine].launchPersistentContext(out+'/profile-'+mode+'-'+Date.now(),{headless:true,viewport:{width:1280,height:850},serviceWorkers:'block',args:engine==='chromium'?['--enable-precise-memory-info']:[]});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.__writes=[];const fail=ReadableStreamDefaultController.prototype.error;ReadableStreamDefaultController.prototype.error=function(e){window.__streamError=e?.stack??String(e);return fail.call(this,e);};window.showSaveFilePicker=async options=>({name:options.suggestedName,createWritable:async()=>{let file;return{write:async blob=>{file=blob;},close:async()=>{window.__savedBlob=file;window.__savedName=options.suggestedName;window.__writes.push({name:options.suggestedName,size:file.size});file=null;},abort:async()=>{file=null;}};}});});
 await context.route('**/*',async r=>{const u=new URL(r.request().url());if(['blob:','data:'].includes(u.protocol))return r.continue();if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')||!['GET','HEAD'].includes(r.request().method()))return r.abort();
  if(diagnostics&&u.pathname==='/src/app.js'){
   const response=await r.fetch();let text=await response.text();
   text+='\n;globalThis.__exportDiagnostics=()=>{const g=wb.graph;let buckets=0,runLinks=0;for(const cols of g?.cols??[])if(cols)for(const b of cols.values()){buckets++;runLinks+=b.runs.length;}return{bigSaveRun:!!bigSaveRun,exportBusy,graph:!!g,graphRun:!!wb.graphRun,graphBuckets:buckets,graphRunLinks:runLinks,graphDynamic:g?.dyn?.size??0,pending:wb.pending?.length??0,formulaCaches:(wb.caches??[]).reduce((n,c)=>n+(c?.size??0),0),background:globalThis.__exportBg??null}};';
   return r.fulfill({response,body:text});
  }
  if(diagnostics&&u.pathname==='/src/big-storage.js'){
   const response=await r.fetch();let text=await response.text();
   text=text.replace('const chunks=[], blocks=[], partKeys=[];','const chunks=[], blocks=[], partKeys=[];globalThis.__exportBg={sheet:i,chunks:0,bytes:0,metadataChars:sheetTags.reduce((n,s)=>n+s.length,0)};');
   text=text.replace('chunks.push(await bigPack(JSON.stringify(chunk),gz));','const packed=await bigPack(JSON.stringify(chunk),gz);chunks.push(packed);globalThis.__exportBg.chunks++;globalThis.__exportBg.bytes+=packed.size;');
   return r.fulfill({response,body:text});
  }
  return r.continue();});
 page=context.pages()[0];page.setDefaultTimeout(300000);page.on('pageerror',e=>result.errors.push({stage:result.stage,message:e.message}));page.on('dialog',d=>d.accept());
 timer=setInterval(()=>{result.minimumFree=Math.min(result.minimumFree,os.freemem());if(result.minimumFree<1024**3){result.memoryGuard=true;context.close().catch(()=>{});}},500);
 await page.goto(url);await page.waitForFunction(()=>!!window.tabula?.wb());
 probeTimer=setInterval(async()=>{if(probing)return;probing=true;try{result.samples??=[];result.samples.push({time:Date.now(),stage:result.stage,free:os.freemem(),...await page.evaluate(()=>({heap:performance.memory?.usedJSHeapSize,progress:document.querySelector('.lp-bar')?.style.width,message:document.querySelector('.lp-msg')?.textContent,harnessBlobBytes:window.__savedBlob?.size??0,harnessBlobUrl:!!window.__saveObjectUrl,versionChangeCount:window.__versionChangeCount??0,product:window.__exportDiagnostics?.()}))});await save();}catch{}finally{probing=false;}},3000);
 const open=async file=>{result.stage='open-'+path.extname(file);await save();const started=Date.now();await page.evaluate(()=>{window.__before=tabula.wb();});await page.locator('#fileInput').setInputFiles(engine==='webkit'?{name:path.basename(file),mimeType:'application/octet-stream',buffer:await readFile(file)}:file);await page.waitForFunction(()=>!document.querySelector('.load-progress')&&(tabula.wb()!==window.__before||document.querySelector('#dialogLayer .dialog')));assert.equal(await page.evaluate(()=>tabula.wb()!==window.__before),true,await page.locator('#dialogLayer').innerText());await page.evaluate(()=>window.__before=null);for(let n=0;n<8&&await page.locator('#dialogLayer .dialog').count();n++)await page.keyboard.press('Escape');return Date.now()-started;};
 const fingerprint=(content=false)=>page.evaluate(async includeContent=>{
  const {slicerColors}=await import('/src/slicerstyle.js'),{createCellFingerprint}=await import('/tools/export-fingerprint.mjs'),w=tabula.wb();
  return w.sheets.map((s,si)=>{
   const cells=createCellFingerprint(includeContent);
   s.cells.forEachStoredRC((v,r,c)=>cells.add(v,r,c));
   return {name:s.name,used:w.usedRange(si),cells:s.cells.size,...cells.result(),pivots:[s.pivot,...(s.pivotsExtra??[])].filter(Boolean).map(p=>({id:p.id,style:p.style})),slicers:(s.slicers??[]).map(s=>({style:s.style,caption:s.caption,colors:Object.fromEntries(['frame','border','head','selFill','selText','item','itemText','noData'].map(k=>[k,String(slicerColors(s)[k]??'').toLowerCase()]))})),merges:s.merges?.length??0,charts:s.charts?.length??0,images:s.images?.length??0,shapes:s.shapes?.length??0};
  });
 },content);
 result.openMs=await open(source);result.before=await fingerprint();await save();
 if(mode==='export'||mode==='integrity'){
  result.edit=await page.evaluate(sheetName=>{const w=tabula.wb();window.__versionChanges=[];window.__versionChangeCount=0;let version=w.version;Object.defineProperty(w,'version',{configurable:true,get:()=>version,set:v=>{window.__versionChangeCount++;if(window.__versionChanges.length<64)window.__versionChanges.push({from:version,to:v,stack:new Error().stack});version=v;}});const si=w.sheets.findIndex(s=>sheetName?s.name===sheetName:s.pivot&&s.slicers?.length),s=w.sheets[si];if(si<0||!s.pivot||!s.slicers?.length)throw Error('실제 피벗/슬라이서 없음');tabula.switchSheet(si);w.transact(()=>{w.setSheetProp(si,'slicers',s.slicers.map((o,i)=>i===0?{...o,style:'SlicerStyleDark6',custom:{...o.custom,selFill:'#123abc',selText:'#ffffff'}}:o));w.setSheetProp(si,'pivot',{...s.pivot,style:'PivotStyleDark9'});});return{si,slicerId:s.slicers[0].id,pivotId:s.pivot.id};},editSheet);
  result.edited=await fingerprint();await save();
  const waitWritten=async count=>{const until=Date.now()+300000;while(Date.now()<until){if(await page.evaluate(n=>window.__writes.length>n,count))return;const d=page.locator('#dialogLayer .dialog').last();if(await d.count()){const text=await d.innerText();if(text.includes('Excel 호환성 확인'))await d.locator('.dialog-foot .btn.primary').click();else if(text.includes('저장하지 못했습니다')||text.includes('초과'))throw Error(text+'\n'+await page.evaluate(()=>JSON.stringify({stream:window.__streamError,versions:window.__versionChanges,si:tabula.si,title:document.title})));}await page.waitForTimeout(300);}throw Error('파일 쓰기 완료 제한시간 초과');};
  for(const type of (mode==='export'?(process.env.WIXEL_EXPORT_TYPES||'wixel,html,xlsx').split(','):[])){
   if(dropGraph){result.graphReleases??=[];result.graphReleases.push({type,...await page.evaluate(()=>{const w=tabula.wb();if(w.pending?.length||w.holdDirty)throw Error('Cannot release graph with pending edits');const before=window.__exportDiagnostics?.();w.dropGraph();return{before,after:window.__exportDiagnostics?.()};})});await save();}
   result.stage='save-'+type;await save();const started=Date.now(),count=await page.evaluate(()=>window.__writes.length);
   await page.evaluate(()=>tabula.run('saveAs'));await page.locator('.form-dialog input[name="name"]').fill('export-roundtrip');await page.locator('.form-dialog select[name="type"]').selectOption(type);await page.locator('.form-dialog .dialog-foot .btn.primary').click();await waitWritten(count);await page.waitForFunction(()=>!document.querySelector('.load-progress'));
   const target=out+'/export-roundtrip.'+type;
   // The mock transfers ownership only until the browser download reaches disk.
   // No previous Blob, temporary download, or delayed object URL survives into
   // the next save; real FileSystemWritableFileStream does not retain these.
   const download=page.waitForEvent('download');let fileDownload;
   try {
    await page.evaluate(()=>{const a=document.createElement('a');window.__saveObjectUrl=URL.createObjectURL(window.__savedBlob);a.href=window.__saveObjectUrl;a.download=window.__savedName;a.click();});
    fileDownload=await download;await fileDownload.saveAs(target);
   } finally {
    await fileDownload?.delete().catch(()=>{});
    await page.evaluate(()=>{if(window.__saveObjectUrl)URL.revokeObjectURL(window.__saveObjectUrl);window.__saveObjectUrl=null;window.__savedBlob=null;window.__savedName=null;}).catch(()=>{});
   }
   const item={type,bytes:(await stat(target)).size,ms:Date.now()-started};result.exports.push(item);assert.ok(item.bytes>0);await save();
   if(type==='wixel') {result.packedEstimate={bytes:packedPublishedSize(item.bytes),linkFits:Math.ceil(item.bytes/3)*4<=1500000};await save();}
   assert.deepEqual(await page.evaluate(()=>({blob:!!window.__savedBlob,url:!!window.__saveObjectUrl})),{blob:false,url:false});
  }
  result.after=await fingerprint();assert.deepEqual(result.after,result.edited);result.stage='integrity-after-exports';await save();result.integrity=await fingerprint(true);result.ok=true;
 }else{
  const baseline=JSON.parse(await readFile(process.env.WIXEL_EXPORT_BASELINE||out+'/export.json','utf8'));
  assert.equal(baseline.fingerprintVersion,3,'Recreate the integrity baseline: V3 separates formula source, typed scalar values and raw representation');
  const expected=baseline.integrity;assert.ok(expected?.length,'Run WIXEL_EXPORT_MODE=integrity on the source first');
  result.after=await fingerprint(true);
  result.comparisons=result.after.map((s,i)=>({name:s.name,formulas:s.formulas===expected[i].formulas,nonempty:s.nonempty===expected[i].nonempty,hash:s.hash===expected[i].hash&&s.xor===expected[i].xor,formulaText:s.formulaHash===expected[i].formulaHash&&s.formulaXor===expected[i].formulaXor,values:s.valueHash===expected[i].valueHash&&s.valueXor===expected[i].valueXor,caches:s.cached===expected[i].cached&&s.cacheHash===expected[i].cacheHash&&s.cacheXor===expected[i].cacheXor,pivots:JSON.stringify(s.pivots)===JSON.stringify(expected[i].pivots),slicers:JSON.stringify(s.slicers.map(({style,...sl})=>sl))===JSON.stringify(expected[i].slicers.map(({style,...sl})=>sl))}));
  result.ok=result.comparisons.every(s=>s.formulas&&s.nonempty&&s.formulaText&&s.values&&s.pivots&&s.slicers&&(!/\.wixel$/i.test(source)||(s.hash&&s.caches)));
 }
 result.stage='done';await save();
}catch(e){result.error=e.stack;console.error(e.stack);await save();}finally{clearInterval(timer);clearInterval(probeTimer);await context?.close().catch(()=>{});await save();}
// Test the actual online envelope separately after the browser releases its
// live workbook. Building it between file saves artificially adds a second
// independent export workload (and base64 temporaries) to the XLSX test.
if(mode==='export'&&result.exports.some(item=>item.type==='wixel')) {
 try {
  const blob=await openAsBlob(out+'/export-roundtrip.wixel'),body=await packPublishedBlob(blob);
  let linkFits=true;try{assertPublishLinkSize(blob);}catch(error){if(error.code!=='PUBLISH_LINK_TOO_LARGE')throw error;linkFits=false;}
  result.packed={bytes:body.size,linkFits,verified:'local-file-after-browser-close'};
 } catch(error) {result.ok=false;result.packedError=error.stack;}
 await save();
}
console.log(JSON.stringify({ok:result.ok,stage:result.stage,exports:result.exports,packed:result.packed,errors:result.errors,minimumFree:result.minimumFree}));if(!result.ok||result.errors.length)process.exitCode=1;
