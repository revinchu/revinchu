// 실제 IndexedDB에 합성 데이터만 쓰는 저장 엔진 검사. iPhone OS 종료 재현 검사는 아니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:5191/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('로컬 소스 서버만 지원합니다.');
const out=process.env.WIXEL_SHEET_STORAGE_OUT||'D:/Codex/Temp/wixel-sheet-state-storage';await mkdir(out,{recursive:true});
const browser=await chromium.launch(),results=[],pageErrors=[],blocked=[];
const names=['300000 cells: hide, Undo, Redo reuse payload','cell edit prevents payload reuse','other metadata prevents payload reuse','16MB parts survive state record replacement and GC','state mutation during save aborts','record abort preserves previous generation','manifest CAS collision preserves other writer','without Web Locks retain shared old parts'];
try{for(const name of names){
 const context=await browser.newContext(),page=await context.newPage();page.on('pageerror',e=>pageErrors.push(e.message));
 await context.route('**/*',r=>{const u=new URL(r.request().url());if(u.origin!==url.origin||u.pathname.startsWith('/api/')){blocked.push(u.origin+u.pathname);return r.abort();}if(u.pathname==='/__sheet_state_storage')return r.fulfill({contentType:'text/html',body:'<!doctype html><title>합성 시트 저장 검사</title>'});return r.continue();});
 try{
  await page.goto(new URL('/__sheet_state_storage',url).href,{waitUntil:'domcontentloaded'});
  const result=await page.evaluate(async name=>{
   const {Workbook}=await import('/src/workbook.js'),{saveLargeWorkbook,loadLargeWorkbook}=await import('/src/big-storage.js'),{idbGet,idbKeys}=await import('/src/storage.js');
   let checks=0;const eq=(a,b,why)=>{checks++;if(JSON.stringify(a)!==JSON.stringify(b))throw Error(why+': '+JSON.stringify(a)+' != '+JSON.stringify(b));};const ok=(v,why)=>{checks++;if(!v)throw Error(why);};
   const key='synthetic:sheet-state',n=name.startsWith('300000')?300000:1000;
   const w=new Workbook({sheets:[{name:'합성 원본',cells:{}},{name:'합성 다음',cells:{}}]});
   const style={numFmt:'number',decimals:2,fontFamily:'Arial',fontSize:11,fill:'#ffffff',color:'#222222'};
   for(let i=0;i<n;i++)w.sheets[0].cells.setRC(Math.floor(i/20),i%20,{raw:String(i/100),v:i/100,formula:false,style});
   w.sheets[1].cells.setRC(0,0,{raw:'다음',v:'다음',formula:false});
   const withParts=name.startsWith('16MB')||name.startsWith('without');
   if(withParts){const num=new Float64Array(2100000);num[0]=123;num[num.length-1]=987;w.sheets[0].blocks=[{r0:50000,c0:0,n:num.length,ver:0,cols:[{num,str:null,dict:[]}]}];}
   let counts={};const chunks=w.cellRunChunks.bind(w),put=IDBObjectStore.prototype.put;
   w.cellRunChunks=function*(i,...args){for(const chunk of chunks(i,...args)){counts.rows=(counts.rows??0)+chunk.length;counts.chunks=(counts.chunks??0)+1;yield chunk;}};
   const save=async()=>{counts={};const t=performance.now(),saved=await saveLargeWorkbook(key,w,{docName:'합성 저장',si:1});return{manifest:saved.manifest,rows:counts.rows??0,chunks:counts.chunks??0,ms:+(performance.now()-t).toFixed(2)};};
   const load=async()=>new Workbook((await loadLargeWorkbook(key)).workbook);
   const first=await save(),prior=first.manifest,record=await idbGet(prior.sheets[0].key);let metrics;
   if(name.startsWith('300000')){
    const idle=await save();eq(idle.rows,0,'무변경 저장');
    w.transact(()=>w.setSheetProp(0,'state','hidden'));const hidden=await save();eq(hidden.rows,0,'숨김 셀 순회 없음');eq(hidden.chunks,0,'숨김 gzip 없음');eq(hidden.manifest.sheets[1].key,prior.sheets[1].key,'다른 시트 재사용');ok(hidden.manifest.sheets[0].key!==prior.sheets[0].key,'새 메타 레코드');
    let loaded=await load();eq(loaded.sheets[0].state,'hidden','숨김 복구');eq(loaded.getValue(0,14999,19),2999.99,'마지막 값');eq(loaded.sheets[0].cells.size,n,'셀 수 유지');
    w.undo();const undo=await save();eq(undo.rows,0,'Undo 재순회 없음');eq((await load()).sheets[0].state,undefined,'Undo 복구');
    w.redo();const redo=await save();eq(redo.rows,0,'Redo 재순회 없음');eq((await load()).sheets[0].state,'hidden','Redo 복구');metrics={first:first.ms,idle:idle.ms,hidden:hidden.ms,undo:undo.ms,redo:redo.ms};
   }else if(name.startsWith('cell')){
    w.transact(()=>{w.setSheetProp(0,'state','hidden');w.setInput(0,0,0,'999');});const next=await save();eq(next.rows,n,'변경한 데이터 재저장');const loaded=await load();eq(loaded.getValue(0,0,0),999,'새 값 보존');eq(loaded.sheets[0].state,'hidden','상태 보존');
   }else if(name.startsWith('other')){
    w.transact(()=>w.setSheetProp(0,'state','hidden'));w.sheets[0].zoom=175;const next=await save();eq(next.rows,n,'state 이외 메타는 기존 경로');const loaded=await load();eq(loaded.sheets[0].zoom,175,'확대율 보존');eq(loaded.sheets[0].state,'hidden','상태 보존');
   }else if(withParts){
    eq(record.partKeys.length,2,'16MB 분할 생성');if(name.startsWith('without'))Object.defineProperty(navigator,'locks',{value:undefined,configurable:true});
    w.transact(()=>w.setSheetProp(0,'state','veryHidden'));const next=await save();eq(next.rows,0,'분할 블록 변경 없음');const now=await idbGet(next.manifest.sheets[0].key);eq(now.partKeys,record.partKeys,'기존 분할 키 재사용');for(const k of record.partKeys)ok(await idbGet(k),'참조 중 분할 보존');
    eq(!!(await idbGet(prior.sheets[0].key)),name.startsWith('without'),'잠금 있으면 이전 메타만 GC');const loaded=await load();eq(loaded.sheets[0].state,'veryHidden','매우 숨김 보존');eq(loaded.sheets[0].blocks[0].cols[0].num[0],123,'첫 블록 값');eq(loaded.sheets[0].blocks[0].cols[0].num.at(-1),987,'마지막 블록 값');
   }else{
    w.transact(()=>w.setSheetProp(0,'state','hidden'));let hit=false;
    IDBObjectStore.prototype.put=function(value,k){const req=put.call(this,value,k);if(!hit&&value?.meta?.state==='hidden'){hit=true;
     if(name.startsWith('record'))this.transaction.abort();
     else if(name.startsWith('state'))req.addEventListener('success',()=>w.undo());
     else req.addEventListener('success',()=>put.call(this,{...prior,generation:'synthetic-other-tab'},key));
    }return req;};
    let code=null;try{await save();}catch(e){code=e?.code||e?.name||'REJECTED_NULL';}
    ok(hit,'재사용 메타 저장 중 고장 주입');ok(code,'저장 실패가 보고됨');if(name.startsWith('state'))eq(code,'BIG_SAVE_ABORT','state 변경 감지');if(name.startsWith('manifest'))eq(code,'IDB_CONFLICT','CAS 충돌 감지');
    const current=await idbGet(key);eq(current.generation,name.startsWith('manifest')?'synthetic-other-tab':prior.generation,'이전 또는 다른 writer manifest 보존');eq((await load()).sheets[0].state,undefined,'부분 숨김이 복구에 섞이지 않음');eq((await load()).getValue(0,0,0),0,'이전 셀 보존');
    eq((await idbKeys(key+'#g#')).sort(),prior.sheets.map(x=>x.key).sort(),'미완료 메타 정리');
   }
   return{name,checks,metrics};
  },name);
  results.push({ok:true,...result});console.log('OK '+name+' '+result.checks+' checks'+(result.metrics?' '+JSON.stringify(result.metrics):''));
 }catch(e){results.push({name,ok:false,error:e.stack});console.error('FAIL '+name+' '+e.stack);}finally{await context.close();}
}}
finally{await browser.close();}
const summary={cases:results.length,passed:results.filter(x=>x.ok).length,checks:results.reduce((n,x)=>n+(x.checks||0),0),pageErrors,blocked,results};await writeFile(out+'/results.json',JSON.stringify(summary,null,2));console.log(JSON.stringify(summary,null,2));assert.equal(summary.passed,summary.cases);assert.deepEqual(pageErrors,[]);assert.deepEqual(blocked,[]);
