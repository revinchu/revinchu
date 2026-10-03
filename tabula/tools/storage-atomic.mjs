// 실제 IndexedDB, 별도 브라우저 컨텍스트와 합성 문서만 사용한다. 소스/번들 공통 검사.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results=[];
const url=process.env.WIXEL_URL||'http://127.0.0.1:5180/', key='tabula.workbook.v1';
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname))throw new Error('실제 IndexedDB 고장 주입은 로컬 서버에서만 실행합니다.');
async function newPage(context) {
  const p=await context.newPage();p.setDefaultTimeout(30000);
  await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());return p;
}
async function test(name,fn) {
  const context=await browser.newContext(),errors=[],writes=[];
  context.on('page',p=>p.on('pageerror',e=>errors.push(e.message)));
  await context.route('**/*',route=>{if(!['GET','HEAD','OPTIONS'].includes(route.request().method())){writes.push(route.request().url());return route.abort();}return route.continue();});
  await context.addInitScript(()=>{
    window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;try{localStorage.setItem('wixel.options',JSON.stringify({saveConfirm:false}));}catch{}
    window.__idb = async (mode,key,value) => new Promise((resolve,reject)=>{
      const open=indexedDB.open('tabula',1);open.onupgradeneeded=()=>open.result.createObjectStore('docs');open.onerror=()=>reject(open.error);
      open.onsuccess=()=>{const db=open.result,tx=db.transaction('docs',mode==='get'||mode==='keys'?'readonly':'readwrite'),store=tx.objectStore('docs');const req=mode==='get'?store.get(key):mode==='keys'?store.getAllKeys():store.put(value,key);let result;req.onsuccess=()=>{result=req.result;};tx.oncomplete=()=>{db.close();resolve(result??null);};tx.onerror=()=>{db.close();reject(tx.error);};tx.onabort=()=>{db.close();reject(tx.error??new Error('abort'));};};
    });
  });
  try{const p=await newPage(context);if(await p.locator('#autosaveToggle').getAttribute('aria-checked')==='true')await p.locator('#autosaveToggle').click();await fn(p,context);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name,ok:true});console.log('OK '+name);}
  catch(e){results.push({name,ok:false,error:e.message,pageErrors:errors});console.error('NG '+name+': '+e.stack);}
  finally{await context.close();}
}
const read=(p,k=key)=>p.evaluate(k=>window.__idb('get',k),k);
const pointer=p=>p.evaluate(k=>localStorage.getItem(k),key);
const idle=p=>p.evaluate(async()=>{if(!navigator.locks)return;const end=performance.now()+30000;while(performance.now()<end){if(!(await navigator.locks.query()).held.some(x=>x.name==='wixel-large-save:tabula.workbook.v1'))return;await new Promise(r=>setTimeout(r,20));}throw new Error('저장 잠금 해제 시간 초과');});
async function fixture(p,block=false) {
  await p.evaluate(block=>{
    const sheets=[];for(let i=0;i<2;i++){const cells={};for(let r=0;r<30000;r++)cells[`${r},0`]={raw:r===0?'OLD-'+(i+1):String(r+i*100000),...(r===1?{style:{numFmt:'custom',code:'yyyy-mm-dd'}}:{})};sheets.push({name:i?'두 번째':'첫 번째',cells});}
    if(block){const num=new Float64Array(2100000);num[0]=123;num[num.length-1]=987;sheets[0].blocks=[{r0:40000,c0:0,n:num.length,ver:0,cols:[{num,str:null,dict:[]}]}];}
    const t=window.tabula;t.wb().restore({date1904:true,props:{title:'원본 메타'},sheets});t.gv().layout();t.gv().renderAll();t.selectCell(0,0);
  },block);
}
// Ctrl+S는 파일 저장이다. 브라우저 사본은 사용자가 고르는 저장 위치 카드로 저장한다.
async function saveBrowserCopy(p) {
  await p.evaluate(()=>window.tabula.run('saveLocations'));
  const hub=p.locator('.wixel-hub');
  await hub.waitFor({state:'visible'});
  await hub.getByRole('button',{name:/^이 브라우저/}).click();
  // 카드를 다시 그리면서 초점이 사라질 수 있어 허브 안에서 Escape를 누른다.
  await hub.locator('button.back').press('Escape');
  await hub.waitFor({state:'detached'});
}
async function save(p,previous=null,trigger=true) {
  if(trigger)await saveBrowserCopy(p);
  await p.evaluate(async ({key,previous})=>{const end=performance.now()+30000;while(performance.now()<end){const idx=await window.__idb('get',key),ptr=JSON.parse(localStorage.getItem(key)||'null');if([3,4].includes(idx?.v)&&idx.generation!==previous&&ptr?.generation===idx.generation)return;await new Promise(r=>setTimeout(r,20));}throw new Error('manifest 저장 시간 초과: '+document.body.innerText.slice(-900));},{key,previous});
  await idle(p);return read(p);
}
async function change(p){await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>{w.setInput(0,0,0,'NEW-1');w.setInput(1,0,0,'NEW-2');});});}
async function fault(p,mode){await p.evaluate(mode=>{
  window.__faultMode=mode;window.__faultHit=false;
  const original=IDBObjectStore.prototype.put;
  IDBObjectStore.prototype.put=function(value,key){
    const selected=!window.__faultHit&&typeof key==='string'&&key.startsWith('tabula.workbook.v1')&&((mode==='second'&&value?.meta?.name==='두 번째')||(mode==='manifest'&&[3,4].includes(value?.v))||(mode==='edit'&&value?.meta?.name==='첫 번째')||(mode==='cancel'&&value?.meta?.name==='첫 번째')||(mode==='close'&&value?.meta?.name==='두 번째'));
    const req=original.call(this,value,key);if(!selected)return req;window.__faultHit=true;
    if(mode==='second'||mode==='manifest'){this.transaction.abort();return req;}
    if(mode==='edit'||mode==='cancel')req.addEventListener('success',()=>{const t=window.tabula,w=t.wb();if(mode==='edit')w.transact(()=>w.setInput(0,2,0,'편집 중 변경'));else{w.restore({sheets:[{name:'취소 후 문서',cells:{'0,0':{raw:'다른 문서'}}}]});t.gv().layout();t.gv().renderAll();}});
    if(mode==='close'){const store=this;req.addEventListener('success',()=>{window.__saveStalled=true;const keep=()=>{const r=store.get('__synthetic_keepalive__');r.onsuccess=keep;};keep();});}
    return req;
  };
},mode);}
async function reopen(p) { const context=p.context();await p.close({runBeforeUnload:false});return newPage(context); }
async function reloadAndCheck(p,expected='OLD',block=false){
  p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb()?.sheets.length===2);
  const value=await p.evaluate(()=>{const w=window.tabula.wb();return {rows:w.sheets.map(s=>s.cells.size),names:w.sheets.map(s=>s.name),values:[w.getValue(0,0,0),w.getValue(1,0,0)],date1904:w.date1904,title:w.props.title,last:[w.getValue(0,29999,0),w.getValue(1,29999,0)],block:w.sheets[0].blocks.length?[w.getValue(0,40000,0),w.getValue(0,2139999,0)]:null};});
  assert.deepEqual(value,{rows:[30000,30000],names:['첫 번째','두 번째'],values:[expected+'-1',expected+'-2'],date1904:true,title:'원본 메타',last:[29999,129999],block:block?[123,987]:null});return p;
}
try {
  await test('두 번째 시트 저장 중 실제 IDB abort: 이전 시트·16MB 분할 청크·1904 날짜 복구',async p=>{
    await fixture(p,true);const prior=await save(p),first=await read(p,prior.sheets[0].key);assert.equal(first.partKeys.length,2);
    const before=await pointer(p);await change(p);await fault(p,'second');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__faultHit);await idle(p);
    assert.equal((await read(p)).generation,prior.generation);assert.equal(await pointer(p),before);for(const k of first.partKeys)assert.ok(await read(p,k));
    await reloadAndCheck(p,'OLD',true);
  });
  await test('최종 manifest 트랜잭션 abort: 완전한 이전 두 시트 유지',async p=>{
    await fixture(p);const prior=await save(p),before=await pointer(p);await change(p);await fault(p,'manifest');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__faultHit);await idle(p);
    assert.equal((await read(p)).generation,prior.generation);assert.equal(await pointer(p),before);await reloadAndCheck(p);
  });
  await test('저장 도중 셀 편집: 혼합 스냅샷 거부 후 명시 재저장으로 최신값 복구',async p=>{
    await fixture(p);const prior=await save(p);await change(p);await fault(p,'edit');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__faultHit);await idle(p);
    assert.equal((await read(p)).generation,prior.generation);await save(p,prior.generation);
    p=await reloadAndCheck(p,'NEW');assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,0)),'편집 중 변경');
  });
  await test('문서 교체로 저장 취소: 이미 쓴 새 시트가 이전 manifest에 노출되지 않음',async p=>{
    await fixture(p);const prior=await save(p);await change(p);await fault(p,'cancel');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__faultHit);await idle(p);
    assert.equal((await read(p)).generation,prior.generation);await reloadAndCheck(p);
  });
  await test('두 번째 시트 트랜잭션 중 탭 종료: 다음 탭이 이전 완전 문서 복구',async(p,context)=>{
    await fixture(p);const prior=await save(p);await change(p);await fault(p,'close');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__saveStalled);await p.close();
    const next=await newPage(context);assert.equal((await read(next)).generation,prior.generation);await reloadAndCheck(next);
  });
  await test('완료된 새 세대 복구·변경 없는 시트 재사용·이전 세대 정리',async p=>{
    await fixture(p);const prior=await save(p);await p.evaluate(()=>window.tabula.wb().transact(()=>window.tabula.wb().setInput(0,0,0,'NEW-1')));const next=await save(p,prior.generation);
    assert.notEqual(next.sheets[0].key,prior.sheets[0].key);assert.equal(next.sheets[1].key,prior.sheets[1].key);assert.equal(await read(p,prior.sheets[0].key),null);
    p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb()?.sheets.length===2);
    assert.deepEqual(await p.evaluate(()=>[window.tabula.wb().getValue(0,0,0),window.tabula.wb().getValue(1,0,0),window.tabula.wb().date1904]),['NEW-1','OLD-2',true]);
  });
  await test('첫 대용량 저장 실패는 기존 localStorage 작은 문서 포인터를 보존',async p=>{
    await p.evaluate(()=>{const t=window.tabula;t.wb().restore({sheets:[{name:'작은 원본',cells:{'0,0':{raw:'작은 원본 값'}}}]});t.gv().layout();t.gv().renderAll();});await saveBrowserCopy(p);
    const before=await pointer(p);assert.match(before,/작은 원본 값/);await fixture(p);await fault(p,'second');await saveBrowserCopy(p);await p.waitForFunction(()=>window.__faultHit);await idle(p);assert.equal(await pointer(p),before);
    p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb());assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),'작은 원본 값');
  });
  await test('IDB commit 성공 뒤 localStorage 실패: 거짓 성공 없이 새 완전 세대 복구',async p=>{
    await fixture(p);const prior=await save(p),before=await pointer(p);await change(p);
    await p.evaluate(()=>{const original=Storage.prototype.setItem;Storage.prototype.setItem=function(key,value){if(!window.__lsFault&&key==='tabula.workbook.v1'&&JSON.parse(value).idb){window.__lsFault=true;throw new DOMException('합성 quota','QuotaExceededError');}return original.call(this,key,value);};});
    await saveBrowserCopy(p);await p.waitForFunction(()=>window.__lsFault);await idle(p);
    assert.notEqual((await read(p)).generation,prior.generation);assert.equal(await pointer(p),before);
    assert.match(await p.locator('#saveState').innerText(),/저장 안 됨/);await reloadAndCheck(p,'NEW');
  });
  await test('_ev가 바뀌지 않은 시트 확대율도 새 시트 레코드로 저장한다',async p=>{
    await fixture(p);const prior=await save(p);
    await p.evaluate(()=>{window.tabula.wb().sheets[0].zoom=175;});const next=await save(p,prior.generation);
    assert.equal(next.sheets[0].ev,prior.sheets[0].ev);assert.notEqual(next.sheets[0].key,prior.sheets[0].key);assert.equal(next.sheets[1].key,prior.sheets[1].key);
    p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb());assert.equal(await p.evaluate(()=>window.tabula.wb().sheets[0].zoom),175);
  });
  await test('CAS 후 정리 중 편집: 저장 완료 표시를 보류하고 다음 저장에 최신값 반영',async p=>{
    await fixture(p);const prior=await save(p),before=await pointer(p);await change(p);
    await p.evaluate(()=>{const original=IDBObjectStore.prototype.delete;IDBObjectStore.prototype.delete=function(key){const req=original.call(this,key);if(!window.__postCommitEdit&&typeof key==='string'&&key.includes('#g#')){window.__postCommitEdit=true;req.addEventListener('success',()=>window.tabula.wb().transact(()=>window.tabula.wb().setInput(0,2,0,'CAS 이후 편집')));}return req;};});
    await saveBrowserCopy(p);await p.waitForFunction(()=>window.__postCommitEdit);await idle(p);const staged=await read(p);
    assert.notEqual(staged.generation,prior.generation);assert.equal(await pointer(p),before);assert.match(await p.locator('#saveState').innerText(),/저장 안 됨/);
    await save(p,staged.generation);p=await reloadAndCheck(p,'NEW');assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,0)),'CAS 이후 편집');
  });
  await test('Web Locks 미지원 시에도 이전 세대 청크를 삭제하지 않고 새 문서를 완성',async p=>{
    await fixture(p);const prior=await save(p);await p.evaluate(()=>Object.defineProperty(navigator,'locks',{value:undefined,configurable:true}));await change(p);const next=await save(p,prior.generation);
    assert.notEqual(next.generation,prior.generation);assert.ok(await read(p,prior.sheets[0].key));await reloadAndCheck(p,'NEW');
  });
  await test('자동 저장 중 편집 취소는 다음 예약 저장에서 완전한 최신 문서로 재시도',async p=>{
    await fixture(p);const prior=await save(p);await change(p);await fault(p,'edit');
    await p.locator('#autosaveToggle').click();await p.waitForFunction(()=>window.__faultHit);
    const next=await save(p,prior.generation,false);assert.equal(next.autosave,true);
    p=await reloadAndCheck(p,'NEW');assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,0)),'편집 중 변경');
  });
  await test('저장 중 외부 manifest 교체는 CAS 충돌로 보존하고 덮어쓰지 않는다',async p=>{
    await fixture(p);const prior=await save(p);await change(p);
    await p.evaluate(prior=>{const original=IDBObjectStore.prototype.put;IDBObjectStore.prototype.put=function(value,key){const req=original.call(this,value,key);if(!window.__competing&&value?.meta?.name==='첫 번째'){window.__competing=true;req.addEventListener('success',()=>original.call(this,{...prior,generation:'synthetic-other-tab'},'tabula.workbook.v1'));}return req;};},prior);
    await saveBrowserCopy(p);await p.waitForFunction(()=>window.__competing);await idle(p);
    assert.equal((await read(p)).generation,'synthetic-other-tab');await reloadAndCheck(p);
  });
  await test('v2→v4 변환은 구버전 탭의 시트·분할 청크·별도 저장 레코드를 보존한다',async p=>{
    await fixture(p);
    const legacy=await p.evaluate(async key=>{
      const w=window.tabula.wb(),sheets=[],partKey=key+'#legacy-0#p0',otherKey=key+'#legacy-other-tab';
      await window.__idb('set',partKey,new Float64Array([123,987]));
      for(let i=0;i<w.sheets.length;i++) {
        const id='legacy-'+i,record={meta:w.sheetMeta(i),chunks:[...w.cellChunks(i)].map(chunk=>JSON.stringify(chunk)),blocks:[],partKeys:[]};
        if(i===0){record.blocks=[{r0:40000,c0:0,n:2,ver:0,cols:[{num:null,str:null,dict:[],numParts:{parts:[partKey],len:2,kind:'Float64Array'}}]}];record.partKeys=[partKey];}
        await window.__idb('set',key+'#'+id,record);sheets.push({id,ev:0});
      }
      const manifest={v:2,docName:'이전 형식 원본',docId:'legacy-migration',autosave:false,book:w.bookMeta(),sheets,names:[],vba:null};
      await window.__idb('set',key,manifest);await window.__idb('set',otherKey,{writer:'구버전 별도 탭',value:456});
      localStorage.setItem(key,JSON.stringify({idb:true,docId:'legacy-migration',remoteDoc:false}));
      return {manifest,partKey,otherKey};
    },key);
    p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb()?.sheets.length===2);
    const next=await save(p);assert.equal(next.v,4);assert.ok(next.sheets.every(sheet=>sheet.key.startsWith(key+'#g#')));
    assert.deepEqual(await p.evaluate(async({key,legacy})=>{
      const first=await window.__idb('get',key+'#'+legacy.manifest.sheets[0].id),second=await window.__idb('get',key+'#'+legacy.manifest.sheets[1].id),part=await window.__idb('get',legacy.partKey),other=await window.__idb('get',legacy.otherKey);
      return {values:[first&&JSON.parse(first.chunks[0])[0][1].raw,second&&JSON.parse(second.chunks[0])[0][1].raw],part:part?[...part]:null,other};
    },{key,legacy}),{values:['OLD-1','OLD-2'],part:[123,987],other:{writer:'구버전 별도 탭',value:456}});
    p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb()?.sheets.length===2);
    assert.deepEqual(await p.evaluate(()=>{const w=window.tabula.wb();return [w.getValue(0,0,0),w.getValue(1,0,0),w.getValue(0,40000,0),w.getValue(0,40001,0),w.date1904];}),['OLD-1','OLD-2',123,987,true]);
  });
  for(const version of [1,2])await test(`기존 v${version} IndexedDB 저장본·날짜 체계 읽기 하위 호환`,async p=>{
    await p.evaluate(async({key,version})=>{const workbook={date1904:true,props:{title:'옛 저장본'},sheets:[{name:'이전 형식',cells:{'0,0':{raw:'이전 값'},'1,0':{raw:'1',style:{numFmt:'custom',code:'yyyy-mm-dd'}}}}]};
      if(version===1)await window.__idb('set',key,{docName:'v1 문서',autosave:false,workbook});
      else{await window.__idb('set',key+'#old',{meta:{name:'이전 형식'},chunks:[JSON.stringify(Object.entries(workbook.sheets[0].cells))],blocks:[]});await window.__idb('set',key,{v:2,docName:'v2 문서',autosave:false,book:{date1904:true,props:{title:'옛 저장본'}},sheets:[{id:'old',ev:0}],names:[],vba:null});}
      localStorage.setItem(key,JSON.stringify({idb:true,docId:'legacy-id',remoteDoc:false}));
    },{key,version});p=await reopen(p);await p.waitForFunction(()=>window.tabula?.wb());assert.deepEqual(await p.evaluate(()=>{const w=window.tabula.wb();return [w.getValue(0,0,0),w.date1904,w.props.title];}),['이전 값',true,'옛 저장본']);
  });
} finally {await browser.close();}
console.log(JSON.stringify({total:results.length,ok:results.filter(x=>x.ok).length,bad:results.filter(x=>!x.ok)},null,2));
if(results.some(x=>!x.ok))process.exitCode=1;
