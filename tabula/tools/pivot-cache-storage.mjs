// 합성 피벗 캐시만 IndexedDB에 저장한다. 실제 업무 파일과 원격 API는 사용하지 않는다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:5191/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('로컬 소스 서버만 검사할 수 있습니다.');
const out=process.env.WIXEL_CACHE_STORAGE_OUT||'D:/Codex/Temp/wixel-final-audit/cache-storage';
await mkdir(out,{recursive:true});
const browser=await chromium.launch(),context=await browser.newContext(),page=await context.newPage(),errors=[];
page.on('pageerror',e=>errors.push(e.message));
await context.route('**/*',r=>{const u=new URL(r.request().url());if(u.origin!==url.origin||u.pathname.startsWith('/api/'))return r.abort();if(u.pathname==='/__pivot_cache_storage')return r.fulfill({contentType:'text/html',body:'<!doctype html><title>합성 피벗 캐시 복원 검사</title>'});return r.continue();});
let report;
try{
 await page.goto(new URL('/__pivot_cache_storage',url).href);
 report=await page.evaluate(async()=>{
  const {Workbook}=await import('/src/workbook.js'),{PivotSnapshotBuilder}=await import('/src/pivot-cache-data.js'),{pivotSourceData}=await import('/src/pivot.js');
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('/src/big-storage.js'),{idbGet,idbSet,idbDeleteMany,idbKeys}=await import('/src/storage.js');
  let checks=0;const eq=(a,b,m)=>{checks++;if(JSON.stringify(a)!==JSON.stringify(b))throw Error(m);},ok=(a,m)=>{checks++;if(!a)throw Error(m);};
  const key='synthetic:pivot-cache',builder=new PivotSnapshotBuilder(['분류','금액'],10000);
  for(let i=0;i<10000;i++)builder.add([i%2?'가':'',i%7?i:null]);
  const w=new Workbook({pivotSnapshots:{saved:builder.finish()},sheets:[{name:'원본',cells:{'0,0':{raw:'분류'},'0,1':{raw:'금액'},'1,0':{raw:'가'},'1,1':{raw:'5'}}},{name:'보고서',cells:{},pivot:{source:'원본',range:{r1:0,c1:0,r2:1,c2:1},snapshotId:'saved'}}]});
  w.pivotCacheItems={saved:{fields:[{name:'분류',shared:Array.from({length:4100},(_,i)=>i===4099?{error:'#N/A'}:'과거 항목 '+i)}]}};w.sheets[1].pivot.cacheItemsId='saved';
  const source=b=>pivotSourceData(b,b.sheets[1].pivot),save=()=>saveLargeWorkbook(key,w,{docName:'합성 캐시'}),load=async()=>new Workbook((await loadLargeWorkbook(key)).workbook);
  w.sheets[1].cells.setRunRC(10,3,1_000_000,{raw:'',style:{fill:'#aabbcc',bold:true}});
  const expected=source(w).rows,first=(await save()).manifest;
  eq(first.v,5,'압축 저장 형식 버전');
  ok(first.snapshotKey,'캐시 별도 레코드 저장');ok(!first.book.pivotSnapshots,'메타 JSON에서 대형 배열 제외');
  ok(first.cacheItemsKey,'과거 항목 별도 레코드 저장');ok(!first.book.pivotCacheItems,'메타에서 과거 항목 배열 제외');
  const history=await idbGet(first.cacheItemsKey);ok(history.partKeys.length>=3,'과거 항목을 여러 조각으로 기록');
  let restored=await load();eq(restored.pivotCacheItems,w.pivotCacheItems,'실제 IDB의 과거 항목·오류 형식 복원');eq(source(restored).rows,expected,'모든 캐시 값 복원');eq(restored.sheets[1].cells.size,1_000_000,'백만 빈 셀 좌표 보존');eq(restored.styleAt(1,1_000_009,3).fill,'#aabbcc','마지막 빈 셀 서식');ok([...restored.sheets[1].cells.storageEntries()].length<4,'복원 때 범위를 펼치지 않음');eq(restored.getValue(0,1,1),5,'원본 값과 구별');
  const again=(await save()).manifest;eq(again.snapshotKey,first.snapshotKey,'무변경 캐시 재사용');
  w.transact(()=>w.setSheetProp(1,'state','hidden'));const hidden=(await save()).manifest;eq(hidden.snapshotKey,first.snapshotKey,'시트 숨김도 캐시 재사용');
  eq(source(await load()).rows,expected,'숨김 저장 후 캐시 복원');ok(await idbGet(first.snapshotKey),'참조된 캐시 GC 보존');
  const expectedStoredKeys=(await idbKeys(key+'#g#')).sort(),storedSnapshot=await idbGet(first.snapshotKey);
  const oldCache=w.pivotSnapshots.get('saved');w.pivotSnapshots.set('saved',{rows:[['분류','금액'],['가',777]],ver:0});
  const put=IDBObjectStore.prototype.put;let injected=false;IDBObjectStore.prototype.put=function(value,k){const result=put.call(this,value,k);if(String(k).endsWith('#pivot-cache')){injected=true;this.transaction.abort();}return result;};
  let aborted=false;try{await save();}catch{aborted=true;}finally{IDBObjectStore.prototype.put=put;w.pivotSnapshots.set('saved',oldCache);}
  ok(injected&&aborted,'캐시 저장 실패 주입');eq((await idbGet(key)).generation,hidden.generation,'실패 시 기존 manifest 유지');eq(source(await load()).rows,expected,'실패 후 이전 완전본 복원');
  eq((await idbKeys(key+'#g#')).sort(),expectedStoredKeys,'실패 캐시 조각 제거');
  await idbDeleteMany([first.snapshotKey]);let missing=false;try{await load();}catch(e){missing=e.message.includes('피벗 캐시');}ok(missing,'누락 캐시는 원본으로 조용히 대체하지 않음');await idbSet(first.snapshotKey,storedSnapshot);
  w.transact(()=>w.setInput(0,1,1,'456'));const edited=(await save()).manifest;ok(!edited.snapshotKey,'원본 변경 뒤 오래된 캐시 제거');
  eq(source(await load()).rows[1][1],456,'편집된 원본 기준 복원');ok(!(await idbGet(first.snapshotKey)),'더 이상 참조하지 않는 캐시 GC');
  // 과거 v3의 일반 셀 청크는 새 형식에서 읽고 무변경 저장으로 안전하게 재사용한다.
  const legacyKey='synthetic:pivot-cache-legacy',legacyRecord=legacyKey+'#g#old#s0';
  await idbSet(legacyRecord,{meta:{name:'이전'},chunks:[new Blob([JSON.stringify([['1,2',{raw:'17',style:{bold:true}}]])])],gz:false,blocks:[],partKeys:[]});
  await idbSet(legacyKey,{v:3,generation:'old',book:{},sheets:[{id:'legacy-sheet',ev:0,key:legacyRecord}]});
  const legacy=new Workbook((await loadLargeWorkbook(legacyKey)).workbook);eq(legacy.getValue(0,1,2),17,'v3 이전 셀 청크 복원');
  const upgraded=await saveLargeWorkbook(legacyKey,legacy,{docName:'이전 보관본'});eq(upgraded.manifest.v,5,'v3에서v5 메타 전환');
  const legacyReload=new Workbook((await loadLargeWorkbook(legacyKey)).workbook);eq(legacyReload.getValue(0,1,2),17,'이전 청크 재사용 후 값 유지');ok(legacyReload.styleAt(0,1,2).bold,'이전 서식 유지');
  return{checks,ok:true};
 });
 assert.deepEqual(errors,[]);report.pageErrors=errors;
}finally{await browser.close();}
await writeFile(out+'/result.json',JSON.stringify(report,null,2));console.log(JSON.stringify(report));
