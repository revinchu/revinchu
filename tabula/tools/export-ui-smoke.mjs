// Synthetic local-only export UI check. File picker and every API are mocked;
// it does not touch user files or transmit workbook data outside localhost.
import assert from 'node:assert/strict';
import { mkdir,writeFile } from 'node:fs/promises';
import { gzipSync } from 'node:zlib';
import { Workbook } from '../src/workbook.js';
import { readXlsx } from '../src/xlsx.js';
import { readWixelFile } from '../src/wixel-file.js';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',origin=new URL(url).origin;
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname))throw Error('로컬 합성 검사 전용');
const out=process.env.WIXEL_EXPORT_SMOKE_OUT||'D:/Codex/Temp/wixel-export-repair/smoke';await mkdir(out,{recursive:true});
const engines=(process.env.WIXEL_BROWSER||'chromium,webkit').split(','),results=[];
const fixture={app:'wixel',docName:'합성 저장',si:0,workbook:{sheets:[{name:'데이터',cells:{'0,0':{raw:'분류'},'0,1':{raw:'금액'},'1,0':{raw:'가'},'1,1':{raw:'=SUM(4,6)',cached:10,style:{fill:'#ffeecc',bold:true}},'2,0':{raw:'나'},'2,1':{raw:'20'}},page:{printArea:'A1:XFD1048576'},shapes:[{id:'shape1',kind:'rect',x:250,y:30,w:80,h:40,fill:'#008844',text:'보존'}]},{name:'숨김',state:'hidden',cells:{'0,0':{raw:'숨긴 원본'}}}]}};
const bytes=Buffer.from(JSON.stringify(fixture));
function init(){
  window.WIXEL_SKIP_START=true;try{localStorage.setItem('wixel.connection.v3',JSON.stringify({kind:'node'}));}catch{}
  window.__apiBodies=[];const nativeFetch=window.fetch;window.fetch=async(input,options={})=>{if(options.body instanceof Blob)window.__apiBodies.push({url:String(input),body:await options.body.text()});return nativeFetch(input,options);};
  window.__fileMock={calls:[],saved:[],cancel:false};
  Object.defineProperty(window,'showSaveFilePicker',{configurable:true,value:async options=>{
    const m=window.__fileMock;m.calls.push(options.suggestedName);
    if(m.cancel){m.cancel=false;throw new DOMException('합성 취소','AbortError');}
    return {name:options.suggestedName,kind:'file',async createWritable(){let buffer;return {async write(blob){buffer=Array.from(new Uint8Array(await blob.arrayBuffer()));},async close(){m.saved.push({name:options.suggestedName,bytes:buffer});},async abort(){}};}};
  }});
}
async function setup(browser,options={}){
  const context=await browser.newContext({viewport:{width:1280,height:900},acceptDownloads:true}),page=await context.newPage();
  const state={writes:[],published:null,errors:[],blocked:[]};page.setDefaultTimeout(18000);page.on('pageerror',e=>state.errors.push(e.message));
  await context.addInitScript(init);
  await context.route('**/*',async route=>{
    const req=route.request(),u=new URL(req.url());
    if(u.origin!==origin){state.blocked.push(req.url());return route.abort();}
    if(!u.pathname.startsWith('/api/'))return route.continue();
    let data={};
    if(u.pathname==='/api/health')data={ok:true,vault:false,maxDocumentBytes:50*1024*1024};
    else if(u.pathname==='/api/files')data=[];
    else if(u.pathname==='/api/publish'&&req.method()==='POST'){const body=req.postData()??await page.evaluate(()=>window.__apiBodies.at(-1)?.body);state.writes.push(body);state.published=JSON.parse(body);data={id:'synthetic-pub',revision:1};}
    else if(u.pathname.startsWith('/api/published/'))data=state.published??options.published??{};
    else if(!['GET','HEAD'].includes(req.method())){state.writes.push(req.postData());data={ok:true,revision:1};}
    await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(data),headers:{'X-Modified':'1'}});
  });
  await page.goto(url,{waitUntil:'domcontentloaded'});await page.waitForFunction(()=>!!window.tabula?.wb());
  if(await page.locator('#autosaveToggle').getAttribute('aria-checked')==='true')await page.locator('#autosaveToggle').click();
  return {context,page,state};
}
async function load(page,name='fixture.wixel',buffer=bytes){await page.locator('#fileInput').setInputFiles({name,mimeType:'application/octet-stream',buffer});await page.waitForFunction(()=>window.tabula?.wb()?.sheets[0]?.name==='데이터');}
async function saveAs(page,type,name){
  const previous=await page.evaluate(()=>window.__fileMock.saved.length);await page.evaluate(()=>tabula.run('saveAs'));
  const dialog=page.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});await dialog.getByLabel('파일 이름',{exact:true}).fill(name);await dialog.locator('select').selectOption(type);await dialog.getByRole('button',{name:'저장',exact:true}).click();
  await page.waitForFunction(n=>window.__fileMock.saved.length>n,previous);await dialog.waitFor({state:'hidden'});
  return page.evaluate(()=>window.__fileMock.saved.at(-1));
}
const assertBook=book=>{assert.equal(book.getValue(0,1,1),10);assert.equal(book.getRaw(0,1,1),'=SUM(4,6)');assert.equal(book.sheets[1].state,'hidden');assert.equal(book.styleAt(0,1,1).fill,'#ffeecc');};
async function run(engine,name,action){const browser=await pw[engine].launch();let env;try{env=await setup(browser);await action(env);assert.deepEqual(env.state.errors,[]);assert.deepEqual(env.state.blocked,[]);results.push({engine,name,ok:true});console.log('PASS',engine,name);}catch(error){results.push({engine,name,ok:false,error:error.stack,errors:env?.state.errors});if(env)await env.page.screenshot({path:out+'/'+engine+'-'+name+'.png'}).catch(()=>{});console.log('FAIL',engine,name,error.message);}finally{await browser.close();await writeFile(out+'/result.json',JSON.stringify({results},null,2));}}
for(const engine of engines){
  await run(engine,'save-formats-and-cancel',async({page,state})=>{
    await load(page);
    const wixel=await saveAs(page,'wixel','합성 WIXEL');assertBook(new Workbook((await readWixelFile(new Blob([Uint8Array.from(wixel.bytes)]))).workbook));
    await load(page,'roundtrip.wixel',Buffer.from(wixel.bytes));assert.equal(await page.evaluate(()=>tabula.wb().getValue(0,1,1)),10);
    const xlsx=await saveAs(page,'xlsx','합성 XLSX');assertBook(new Workbook(readXlsx(Uint8Array.from(xlsx.bytes)).data));
    const html=await saveAs(page,'html','합성 HTML'),text=Buffer.from(html.bytes).toString('utf8');assert.match(text,/<html/i);assert.match(text,/금액/);assert.match(text,/10/);assert.doesNotMatch(text,/페이지가 2,000쪽/);
    const before=await page.evaluate(()=>({title:document.querySelector('#docTitle').textContent,value:tabula.wb().getValue(0,1,1),saved:__fileMock.saved.length,calls:__fileMock.calls.length}));
    await page.evaluate(()=>{__fileMock.cancel=true;tabula.run('saveAs');});const d=page.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});await d.getByLabel('파일 이름',{exact:true}).fill('취소 이름');await d.locator('select').selectOption('wixel');await d.getByRole('button',{name:'저장',exact:true}).click();await d.waitFor({state:'hidden'});
    const after=await page.evaluate(()=>({title:document.querySelector('#docTitle').textContent,value:tabula.wb().getValue(0,1,1),saved:__fileMock.saved.length,calls:__fileMock.calls.length}));assert.deepEqual({...after,calls:before.calls},before);assert.equal(after.calls,before.calls+1);assert.equal(state.writes.length,0);
  });
  await run(engine,'publish-and-link',async({page,state,context})=>{
    await load(page);await page.evaluate(()=>tabula.run('publish'));let d=page.getByRole('dialog',{name:'공유 · 웹에 게시',exact:true});
    await d.getByRole('button',{name:/링크 만들기/}).click();await d.locator('.pub-link input').waitFor();const link=await d.locator('.pub-link input').inputValue();assert.ok(link.includes('#view='));
    await d.getByRole('button',{name:/온라인에 게시/}).click();await page.waitForFunction(()=>document.querySelector('.pub-link input')?.value.includes('?view='));assert.equal(state.writes.length,1);assert.equal(state.published.format,'wixel-packed');
    const p2=await context.newPage();await p2.goto(link);await p2.waitForFunction(()=>document.body.classList.contains('view-mode')&&window.tabula?.wb()?.sheets[0]?.name==='데이터');assert.equal(await p2.evaluate(()=>tabula.wb().getValue(0,1,1)),10);await p2.close();
    const legacy=origin+'/#view='+gzipSync(bytes).toString('base64url'),p3=await context.newPage();await p3.goto(legacy);await p3.waitForFunction(()=>document.body.classList.contains('view-mode')&&window.tabula?.wb()?.sheets[0]?.name==='데이터');assert.equal(await p3.evaluate(()=>tabula.wb().getValue(0,1,1)),10);await p3.close();
  });
  await run(engine,'switch-during-publish',async({page,state})=>{
    await load(page);await page.evaluate(()=>{const original=Response.prototype.blob;window.__holdExport=true;Response.prototype.blob=async function(...args){const value=await original.apply(this,args);if(window.__holdExport){window.__holdExport=false;window.__exportPaused=true;await new Promise(resolve=>window.__resumeExport=resolve);}return value;};tabula.run('publish');});
    await page.getByRole('dialog',{name:'공유 · 웹에 게시',exact:true}).getByRole('button',{name:/온라인에 게시/}).click();await page.waitForFunction(()=>window.__exportPaused===true);
    const other={docName:'새 문서',workbook:{sheets:[{name:'새 문서',cells:{'0,0':{raw:'현재 내용'}}}]}};
    await page.locator('#fileInput').setInputFiles({name:'new.wixel',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(other))});
    await page.waitForFunction(()=>tabula.wb().sheets[0].name==='새 문서');await page.evaluate(()=>window.__resumeExport());
    await page.waitForTimeout(150);assert.equal(state.writes.length,0);assert.equal(await page.evaluate(()=>tabula.wb().getValue(0,0,0)),'현재 내용');
  });
}
console.log(JSON.stringify({total:results.length,passed:results.filter(r=>r.ok).length,failed:results.filter(r=>!r.ok)},null,2));if(results.some(r=>!r.ok))process.exitCode=1;
