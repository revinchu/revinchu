// Only synthetic workbooks in a fresh local browser context; no user files or remote APIs.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:5191/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('Local source server required');
const out=process.env.WIXEL_ARRAY_STORAGE_OUT||'D:/Codex/Temp/wixel-final-audit/array-storage';
await mkdir(out,{recursive:true});
const browser=await chromium.launch(),context=await browser.newContext(),page=await context.newPage(),errors=[];
page.on('pageerror',e=>errors.push(e.message));
await context.route('**/*',route=>{
 const u=new URL(route.request().url());
 if(u.origin!==url.origin||u.pathname.startsWith('/api/'))return route.abort();
 if(u.pathname==='/__array_storage')return route.fulfill({contentType:'text/html',body:'<!doctype html><title>Array cache storage test</title>'});
 return route.continue();
});
let report;
try{
 await page.goto(new URL('/__array_storage',url).href);
 report=await page.evaluate(async()=>{
  const {Workbook}=await import('/src/workbook.js'),{saveLargeWorkbook,loadLargeWorkbook}=await import('/src/big-storage.js');
  let checks=0;const eq=(a,b,m)=>{checks++;if(JSON.stringify(a)!==JSON.stringify(b))throw Error(m+': '+JSON.stringify({actual:a,expected:b}));};
  const key='synthetic:imported-array';
  const w=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'2'}}},{name:'Results',cells:{'0,2':{raw:'=SEQUENCE(Source!A1,1,10,10)',cached:10,cachedArray:{h:3,w:1,values:[0,0,10,1,0,20,2,0,30]}}}}]});
  const values=book=>[0,1,2].map(r=>book.getValue(1,r,2));
  const save=()=>saveLargeWorkbook(key,w,{docName:'Synthetic array'}),load=async()=>new Workbook((await loadLargeWorkbook(key)).workbook);
  eq(values(w),[10,20,30],'Initial saved values differ from current source');
  const first=(await save()).manifest;
  eq(values(await load()),[10,20,30],'IndexedDB retains all cached results');
  const again=(await save()).manifest;
  eq(again.sheets.map(s=>s.key),first.sheets.map(s=>s.key),'Unchanged sheet records reused');
  w.transact(()=>w.setStyle(1,0,2,{bold:true}));await save();
  let back=await load();eq(values(back),[10,20,30],'Formatting retains saved results');eq(back.styleAt(1,0,2).bold,true,'Formatting retained');
  w.transact(()=>w.setInput(0,0,0,'1'));
  // Save before reading the dependent array, to exercise dirty sheet persistence.
  await save();back=await load();eq(values(back),[10,null,null],'Source edit never restores trusted stale array');
  eq(back.getCell(1,0,2).dirty,true,'Dependent array remains dirty after load');
  w.undo();await save();back=await load();eq(values(back),[10,20,30],'Undo restores original cached results in IndexedDB');
  w.redo();await save();back=await load();eq(values(back),[10,null,null],'Redo invalidates cache after resaving');
  const scalarKey='synthetic:imported-scalar';
  const scalar=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'2'}}},{name:'Results',fileValues:true,cells:{'0,0':{raw:'=Source!A1*10',cached:99}}},{name:'Untouched',cells:{'0,0':{raw:'42'}}}]});
  const saveScalar=()=>saveLargeWorkbook(scalarKey,scalar,{}),loadScalar=async()=>new Workbook((await loadLargeWorkbook(scalarKey)).workbook);
  const scalarFirst=(await saveScalar()).manifest;
  eq((await loadScalar()).getValue(1,0,0),99,'Scalar cached result initially retained');
  scalar.transact(()=>scalar.setInput(0,0,0,'1'));
  const scalarEdited=(await saveScalar()).manifest,scalarBack=await loadScalar();
  eq(scalarBack.getValue(1,0,0),10,'Dirty scalar cache is not trusted after load');
  eq(scalarBack.getCell(1,0,0).dirty,true,'Scalar remains dirty after load');
  eq(scalarEdited.sheets[1].key===scalarFirst.sheets[1].key,false,'Dependent scalar chunk is rewritten');
  eq(scalarEdited.sheets[2].key,scalarFirst.sheets[2].key,'Unrelated sheet still reused');
  scalar.undo();await saveScalar();eq((await loadScalar()).getValue(1,0,0),20,'Scalar Undo uses restored source');
  scalar.redo();await saveScalar();eq((await loadScalar()).getValue(1,0,0),10,'Scalar Redo never restores stale saved value');

  return{checks,ok:true};
 });
 assert.deepEqual(errors,[]);report.pageErrors=errors;
}finally{await browser.close();}
await writeFile(out+'/result.json',JSON.stringify(report,null,2));console.log(JSON.stringify(report));
