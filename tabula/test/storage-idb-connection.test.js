import test from 'node:test';
import assert from 'node:assert/strict';

// Small deterministic IDB fault harness; real IDB is also checked by
// tools/storage-connection.mjs in Chromium/WebKit isolated browser profiles.
function databaseHarness() {
  const data = new Map(), connections = [];
  let opens = 0, nextOpenError = null, syncOpenError = null, abortNext = null;
  const indexedDB = { open() {
    opens++;
    if (syncOpenError) { const error = syncOpenError; syncOpenError = null; throw error; }
    const request = {};
    queueMicrotask(() => {
      if (nextOpenError) { request.error = nextOpenError; nextOpenError = null; request.onerror?.(); return; }
      const db = { closed: false, transactions: 0, requests: 0, failTransaction: null,
        close() { this.closed = true; },
        transaction() {
          this.transactions++;
          if (this.closed) throw new DOMException('closed connection', 'InvalidStateError');
          if (this.failTransaction) throw this.failTransaction;
          const writes = new Map(), deletes = new Set();
          let pending = 0, finished = false;
          const finish = () => setImmediate(() => {
            if (finished || pending) return;
            if (abortNext) { tx.error = abortNext; abortNext = null; tx.abort(); return; }
            finished = true;
            for (const key of deletes) data.delete(key);
            for (const [key, value] of writes) data.set(key, value);
            tx.oncomplete?.();
          });
          const requestFor = action => {
            const request = {}; pending++; db.requests++;
            queueMicrotask(() => {
              if (finished) { request.error = new DOMException('aborted', 'AbortError'); request.onerror?.(); pending--; return; }
              request.result = action(); request.onsuccess?.(); pending--; finish();
            });
            return request;
          };
          const tx = { error: null,
            abort() { if (finished) return; finished = true; queueMicrotask(() => tx.onabort?.()); },
            objectStore() { return {
              get: key => requestFor(() => structuredClone(writes.has(key) ? writes.get(key) : deletes.has(key) ? undefined : data.get(key))),
              put: (value, key) => requestFor(() => { writes.set(key, structuredClone(value)); deletes.delete(key); return key; }),
              delete: key => requestFor(() => { writes.delete(key); deletes.add(key); }),
              getAllKeys: range => requestFor(() => [...data.keys()].filter(key => key >= range.lower && key <= range.upper)),
            }; },
          };
          finish(); return tx;
        },
      };
      connections.push(db); request.result = db; request.onsuccess?.();
    });
    return request;
  } };
  return { indexedDB, data, connections, get opens() { return opens; },
    failOpen(error, sync = false) { if (sync) syncOpenError = error; else nextOpenError = error; },
    abort(error) { abortNext = error; },
  };
}
let importId = 0;
async function fixture(fn) {
  const previous = globalThis.indexedDB, previousRange = globalThis.IDBKeyRange, harness = databaseHarness();
  globalThis.indexedDB = harness.indexedDB;
  globalThis.IDBKeyRange = { bound: (lower, upper) => ({ lower, upper }) };
  try {
    const storage = await import(`../src/storage.js?connection-test=${++importId}`);
    await fn(storage, harness);
  } finally { for (const db of harness.connections) db.onversionchange?.(); globalThis.indexedDB = previous; globalThis.IDBKeyRange = previousRange; }
}

test('IndexedDB 일시 열기 실패 뒤 다음 저장은 새 연결로 성공한다', () => fixture(async (s, h) => {
  h.failOpen(new DOMException('temporary open failure', 'UnknownError'));
  await assert.rejects(s.idbSet('doc', 'new'), { name: 'UnknownError' });
  await s.idbSet('doc', 'new'); assert.equal(await s.idbGet('doc'), 'new'); assert.equal(h.opens, 2);
}));

test('IndexedDB open의 동기 예외도 다음 저장을 영구 차단하지 않는다', () => fixture(async (s, h) => {
  h.failOpen(new DOMException('temporary unavailable', 'InvalidStateError'), true);
  await assert.rejects(s.idbGet('doc'), { name: 'InvalidStateError' });
  await s.idbSet('doc', 2); assert.equal(await s.idbGet('doc'), 2); assert.equal(h.opens, 2);
}));

test('닫힌 연결의 동시 저장은 한 번 재연결하고 옛 close 이벤트가 새 연결을 버리지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('old', 1); const old = h.connections[0]; old.close();
  await Promise.all(Array.from({ length: 20 }, (_, i) => s.idbSet(`new:${i}`, i)));
  assert.equal(h.opens, 2); assert.equal(h.connections[1].requests, 20);
  old.onclose?.(); assert.equal(await s.idbGet('new:19'), 19); assert.equal(h.opens, 2);
}));

test('강제 close와 versionchange 이후 읽기·저장이 새 연결을 연다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1);
  h.connections[0].closed = true; h.connections[0].onclose?.();
  assert.equal(await s.idbGet('doc'), 1); assert.equal(h.opens, 2);
  h.connections[1].onversionchange?.(); assert.equal(h.connections[1].closed, true);
  await s.idbSet('doc', 2); assert.equal(await s.idbGet('doc'), 2); assert.equal(h.opens, 3);
}));

test('진행 중 quota 중단은 자동 재실행하지 않고 기존 값을 유지한다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 'old'); const db = h.connections[0], before = db.requests;
  h.abort(new DOMException('quota', 'QuotaExceededError'));
  await assert.rejects(s.idbSet('doc', 'new'), { name: 'QuotaExceededError' });
  assert.equal(db.requests, before + 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 'old');
  await s.idbSet('doc', 'new'); assert.equal(await s.idbGet('doc'), 'new');
}));

test('연결 종료가 아닌 트랜잭션 생성 오류는 자동 재시도하지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1); const db = h.connections[0], before = db.transactions;
  db.failTransaction = new DOMException('missing store', 'NotFoundError');
  await assert.rejects(s.idbSet('doc', 2), { name: 'NotFoundError' });
  assert.equal(db.transactions, before + 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 1);
}));

test('이미 시작한 저장 콜백의 InvalidStateError는 재연결·중복 실행하지 않는다', () => fixture(async (s, h) => {
  await s.idbSet('doc', 1); let calls = 0;
  await assert.rejects(s.idbUpdate(['doc'], () => { calls++; throw new DOMException('callback error', 'InvalidStateError'); }), { name: 'InvalidStateError' });
  assert.equal(calls, 1); assert.equal(h.opens, 1); assert.equal(h.data.get('doc'), 1);
}));

test('각 저장 API가 종료된 연결을 복구하고 CAS 충돌 검사는 그대로 유지한다', () => fixture(async (s, h) => {
  const close = () => h.connections.at(-1).close();
  await s.idbSet('a', 1); close(); assert.equal(await s.idbGet('a'), 1);
  close(); await s.idbUpdate(['a'], values => ({ set: [['b', values.get('a') + 1]], result: true }));
  close(); assert.deepEqual((await s.idbKeys('')).sort(), ['a', 'b']);
  close(); await s.idbDel('a'); assert.equal(h.data.has('a'), false);
  close(); await s.idbCompareAndSet('b', 2, 3); assert.equal(h.data.get('b'), 3);
  close(); await assert.rejects(s.idbCompareAndSet('b', 2, 4), { code: 'IDB_CONFLICT' }); assert.equal(h.data.get('b'), 3);
  close(); await s.idbDeleteMany(['b']); assert.equal(h.data.size, 0); assert.equal(h.opens, 8);
}));

test('대형 문서의 열린 세대 검증은 뒤늦은 덮어쓰기를 청크 저장 전에 차단한다', () => fixture(async (_s, h) => {
  const { saveLargeWorkbook, loadLargeWorkbook } = await import('../src/big-storage.js');
  const { Workbook } = await import('../src/workbook.js');
  const key = 'synthetic:observed-generation', original = new Workbook({ sheets: [{ name: '동시 편집', cells: { '0,0': { raw: '1' } } }] });
  const first = await saveLargeWorkbook(key, original, {}, { expectedGeneration: null });
  const current = new Workbook((await loadLargeWorkbook(key)).workbook), stale = new Workbook((await loadLargeWorkbook(key)).workbook);
  current.transact(() => current.setInput(0, 0, 0, '101')); stale.transact(() => stale.setInput(0, 0, 0, '202'));
  const second = await saveLargeWorkbook(key, current, {}, { expectedGeneration: first.manifest.generation });
  const requests = h.connections[0].requests, keys = [...h.data.keys()];
  await assert.rejects(saveLargeWorkbook(key, stale, {}, { expectedGeneration: first.manifest.generation }), { code: 'IDB_CONFLICT' });
  assert.equal(h.connections[0].requests, requests + 1); // Only the manifest read; no staged records or GC.
  assert.deepEqual([...h.data.keys()], keys); assert.equal(h.data.get(key).generation, second.manifest.generation);
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 101);
  assert.equal(stale.getValue(0, 0, 0), 202);
  await assert.rejects(saveLargeWorkbook(key, stale, {}, { expectedGeneration: null }), { code: 'IDB_CONFLICT' });
  const forkKey = key + ':fork'; await saveLargeWorkbook(forkKey, stale, {}, { expectedGeneration: null });
  assert.equal(new Workbook((await loadLargeWorkbook(forkKey)).workbook).getValue(0, 0, 0), 202);
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 101);
  // Callers that have not opted into observed generations retain their prior API.
  stale.transact(() => stale.setInput(0, 0, 0, '203'));
  await saveLargeWorkbook(key, stale, {});
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getValue(0, 0, 0), 203);
}));


test('v5 셀/블록/캐시 저장은 큰 배열을 메타에 남기지 않고 부분별로 왕복한다', () => fixture(async (_s, h) => {
  const { saveLargeWorkbook, loadLargeWorkbook } = await import('../src/big-storage.js');
  const { Workbook } = await import('../src/workbook.js');
  const { PivotSnapshotBuilder } = await import('../src/pivot-cache-data.js');
  const w = new Workbook({sheets:[{name:'조각 검증',cells:{'0,0':{raw:'=1+2',cached:3,style:{bold:true},comment:'검사',link:'https://example.com'},'1,0':{raw:'한글😀'}}}]});
  w.sheets[0].cells.setRunRC(20,4,15,{raw:'',style:{fill:'#aabbcc'}});
  const n=140000, num=new Float64Array(n).fill(NaN), str=new Int32Array(n).fill(-1), dict=['한글😀',true,false,{error:'#N/A'},'',...Array.from({length:4100},(_,i)=>'항목'+i)];
  num[1]=123.5;str[2]=0;str[3]=1;str[4]=2;str[5]=3;str[6]=4;
  w.sheets[0].blocks=[{r0:100,c0:2,n,ver:0,cols:[{num,str,dict,fmt:{numFmt:'0.00'}}],perm:Int32Array.from({length:n},(_,i)=>i)}];
  const builder=new PivotSnapshotBuilder(['분류','값']);for(const row of [['가',1],['나',2],['',{error:'#N/A'}]])builder.add(row);
  const snapshots={compact:builder.finish(),legacy:[['분류','값'],['한글😀',42],['',null]]};
  w.setSnapshots({pivotSnapshots:snapshots});
  const shared=['',true,23,{error:'#N/A'},...Array.from({length:4100},(_,i)=>'과거'+i)];
  const types=['m','b','n','e',...Array.from({length:4100},()=> 's')];
  for(const i of[2047,2048,4095,4096]){shared[i]=45000+i/100;types[i]='d';}
  w.pivotCacheItems={cache:{missingItemsLimit:1048576,fields:[{name:'과거항목',shared,sharedTypes:types.join(''),format:{numFmt:'custom',code:'yyyy-mm-dd hh:mm:ss.000'}}]}};
  const saved=await saveLargeWorkbook('v5',w,{docName:'합성'});assert.equal(saved.manifest.v,5);
  const rec=h.data.get(saved.manifest.sheets[0].key);assert.ok(rec.cellPartKeys.length);assert.equal(rec.chunks,undefined);
  assert.equal(rec.blocks[0].cols[0].num,null);assert.equal(rec.blocks[0].cols[0].dict,null);assert.ok(rec.blocks[0].cols[0].numParts.parts.length>1);
  assert.equal(saved.manifest.book.pivotSnapshots,undefined);assert.equal(saved.manifest.book.pivotCacheItems,undefined);
  for(const [key,value]of h.data)if(ArrayBuffer.isView(value))assert.ok(value.byteLength<=1<<20,key);
  for(const key of rec.cellPartKeys)assert.ok(h.data.get(key)instanceof Blob);
  const loaded=await loadLargeWorkbook('v5'),out=new Workbook(loaded.workbook);
  for(const [r,c]of[[0,0],[1,0],[20,4],[34,4],[101,2],[102,2],[103,2],[104,2],[105,2],[106,2]]){assert.deepEqual(out.getValue(0,r,c),w.getValue(0,r,c));assert.deepEqual(out.styleAt(0,r,c),w.styleAt(0,r,c));}
  assert.equal(out.getCell(0,0,0).raw,'=1+2');assert.equal(out.getCell(0,0,0).comment,'검사');
  assert.deepEqual(loaded.workbook.pivotSnapshots,snapshots);assert.deepEqual(loaded.workbook.pivotCacheItems,w.pivotCacheItems);
  const restoredField=out.pivotCacheItems.cache.fields[0];
  assert.equal(restoredField.sharedTypes.length,restoredField.shared.length);
  for(const i of[2047,2048,4095,4096]){assert.equal(restoredField.sharedTypes[i],'d');assert.equal(restoredField.shared[i],45000+i/100);}
  assert.deepEqual(restoredField.format,{numFmt:'custom',code:'yyyy-mm-dd hh:mm:ss.000'});
  assert.equal(out.pivotCacheItems.cache.missingItemsLimit,1048576);
  const partKey=rec.cellPartKeys[0],part=h.data.get(partKey);h.data.delete(partKey);
  await assert.rejects(loadLargeWorkbook('v5'),/셀 조각/);h.data.set(partKey,part);
  const cache=h.data.get(saved.manifest.snapshotKey),cachePart=cache.partKeys[0],prior=h.data.get(cachePart);h.data.delete(cachePart);
  await assert.rejects(loadLargeWorkbook('v5'),/배열 조각/);h.data.set(cachePart,prior);
  assert.deepEqual((await loadLargeWorkbook('v5')).workbook.pivotSnapshots,snapshots);
}));

test('v5 변경 취소와 quota 실패는 새 셀 청크만 지우고 기존 세대를 보존한다', () => fixture(async (_s,h) => {
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');
  const {Workbook}=await import('../src/workbook.js');
  const w=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'11'}}}]});
  const first=await saveLargeWorkbook('atomic',w,{}),keys=[...h.data.keys()].sort();
  w.transact(()=>w.setInput(0,0,0,'22'));
  const original=w.cellRunChunks;
  w.cellRunChunks=function*(){yield [[0,0,1,{raw:'22'}]];w.version++;yield [[1,0,1,{raw:'33'}]];};
  await assert.rejects(saveLargeWorkbook('atomic',w,{}),{code:'BIG_SAVE_ABORT'});
  assert.deepEqual([...h.data.keys()].sort(),keys);assert.equal(h.data.get('atomic').generation,first.manifest.generation);
  assert.equal(new Workbook((await loadLargeWorkbook('atomic')).workbook).getValue(0,0,0),11);
  w.cellRunChunks=function*(){yield [[0,0,1,{raw:'22'}]];h.abort(new DOMException('quota','QuotaExceededError'));yield [[1,0,1,{raw:'33'}]];};
  await assert.rejects(saveLargeWorkbook('atomic',w,{}),{name:'QuotaExceededError'});
  assert.deepEqual([...h.data.keys()].sort(),keys);assert.equal(h.data.get('atomic').generation,first.manifest.generation);
  w.cellRunChunks=original;
  await saveLargeWorkbook('atomic',w,{});assert.equal(new Workbook((await loadLargeWorkbook('atomic')).workbook).getValue(0,0,0),22);
}));

test('v5 시트 상태만 바꾸면 청크를 재사용하고 잠금 GC는 참조된 캐시 조각을 보존한다', () => fixture(async (_s,h) => {
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const {Workbook}=await import('../src/workbook.js');
  const previousNavigator=Object.getOwnPropertyDescriptor(globalThis,'navigator');
  Object.defineProperty(globalThis,'navigator',{value:{locks:{request:async(_name,_options,fn)=>fn()}},configurable:true});
  try{
    const w=new Workbook({sheets:[{name:'data',cells:{'0,0':{raw:'7'}}},{name:'other',cells:{}}]});
    w.setSnapshots({pivotSnapshots:{saved:[['키','값'],['가',99]]}});w.pivotCacheItems={x:{fields:[{name:'키',shared:['가','과거']}]}};
    const first=await saveLargeWorkbook('gc',w,{}),before=h.data.get(first.manifest.sheets[0].key);
    w.transact(()=>w.setSheetProp(0,'state','hidden'));
    const original=w.cellRunChunks;w.cellRunChunks=function*(){assert.fail('state-only save must not reread cells');};
    const second=await saveLargeWorkbook('gc',w,{});w.cellRunChunks=original;
    const after=h.data.get(second.manifest.sheets[0].key);assert.deepEqual(after.cellPartKeys,before.cellPartKeys);
    assert.equal(second.manifest.snapshotKey,first.manifest.snapshotKey);assert.equal(second.manifest.cacheItemsKey,first.manifest.cacheItemsKey);
    for(const key of [...after.partKeys,...h.data.get(second.manifest.snapshotKey).partKeys,...h.data.get(second.manifest.cacheItemsKey).partKeys])assert.ok(h.data.has(key),key);
    assert.equal(h.data.has(first.manifest.sheets[0].key),false);assert.equal(new Workbook((await loadLargeWorkbook('gc')).workbook).getValue(0,0,0),7);
    const oldCache=second.manifest.cacheItemsKey;w.pivotCacheItems={x:{fields:[{name:'키',shared:['새항목']}]}};
    const third=await saveLargeWorkbook('gc',w,{});assert.notEqual(third.manifest.cacheItemsKey,oldCache);assert.equal(h.data.has(oldCache),false);
    assert.deepEqual((await loadLargeWorkbook('gc')).workbook.pivotCacheItems,w.pivotCacheItems);
  }finally{if(previousNavigator)Object.defineProperty(globalThis,'navigator',previousNavigator);else delete globalThis.navigator;}
}));

test('v2/v3/v4 inline 저장본을 읽고 v5로 저장해도 값과 빈 셀 run을 보존한다', () => fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  for(const v of[2,3,4]){
    const key='legacy'+v,recordKey=key+'#sheet',runs=v===4,entries=runs?[[2,3,1,{raw:'17'}],[4,3,2,{raw:'',style:{bold:true}}]]:[['2,3',{raw:'17'}]];
    h.data.set(key,{v,generation:'old',book:{date1904:true},sheets:[{id:'sheet',key:recordKey,ev:0}]});
    h.data.set(recordKey,{meta:{name:'이전'},chunks:[new Blob([JSON.stringify(entries)])],cellEncoding:runs?'runs-v1':undefined,gz:false,blocks:[],partKeys:[]});
    const loaded=await loadLargeWorkbook(key),book=new Workbook(loaded.workbook);assert.equal(book.getValue(0,2,3),17);assert.equal(book.date1904,true);
    const saved=await saveLargeWorkbook(key,book,{});assert.equal(saved.manifest.v,5);const roundtrip=new Workbook((await loadLargeWorkbook(key)).workbook);assert.equal(roundtrip.getValue(0,2,3),17);if(runs)assert.equal(roundtrip.styleAt(0,5,3).bold,true);
  }
}));


test('v5 autosave shares immutable styles while capturing identical JSON and rejecting later style edits',()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const style={font:'Arial',size:11,bold:true,fill:'#e8f2ff',color:'#16395e',numFmt:'custom',code:'#,##0.00',border:{bottom:{color:'#117744',style:'thin'}}};
  const cells={};for(let r=0;r<2000;r++)cells[r+',0']={raw:r%2?'=ROW()':String(r),style};
  const w=new Workbook({sheets:[{name:'Synthetic',cells}]});
  const original=w.cellRunChunks,expected=Array.from(original.call(w,0),chunk=>JSON.stringify(chunk));let part=0,shared=0;
  w.cellRunChunks=function*(si,size,options){for(const chunk of original.call(this,si,size,options)){assert.equal(JSON.stringify(chunk),expected[part++]);for(const[r,c,_n,data]of chunk){if(data.style){assert.equal(data.style,this.sheets[si].cells.getRC(r,c).style);shared++;}}yield chunk;}};
  const first=await saveLargeWorkbook('shared-style',w,{});assert.equal(part,expected.length);assert.equal(shared,2000);w.cellRunChunks=original;
  const restored=new Workbook((await loadLargeWorkbook('shared-style')).workbook);assert.equal(restored.getRaw(0,0,0),'0');assert.equal(restored.getRaw(0,1,0),'=ROW()');assert.deepEqual(restored.styleAt(0,0,0),w.styleAt(0,0,0));
  w.transact(()=>w.setInput(0,0,0,'9'));let changed=false;
  await assert.rejects(saveLargeWorkbook('shared-style',w,{}, {expectedGeneration:first.manifest.generation,waitForIdle:()=>{if(!changed){changed=true;w.transact(()=>w.setStyle(0,0,0,{bold:false,fill:'#ffeecc'}));}}}),{code:'BIG_SAVE_ABORT'});
  assert.equal(h.data.get('shared-style').generation,first.manifest.generation);
  const previous=new Workbook((await loadLargeWorkbook('shared-style')).workbook);assert.equal(previous.getRaw(0,0,0),'0');assert.equal(previous.styleAt(0,0,0).bold,true);assert.equal(previous.styleAt(0,0,0).fill,'#e8f2ff');
}));


for(const compressed of[true,false])test('v5 UTF8 compression/fallback preserves empty, Korean, emoji and one oversized cell: '+compressed,()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const previous=globalThis.CompressionStream;
  try {
    if(!compressed)globalThis.CompressionStream=undefined;
    const large='한글😀\n'.repeat(220000),w=new Workbook({sheets:[{name:'Unicode',cells:{'0,0':{raw:'',style:{fill:'#123456'}},'1,0':{raw:'한글😀'},'2,0':{raw:large},'3,0':{raw:'=1+2',cached:3}}}]});
    const saved=await saveLargeWorkbook('unicode-'+compressed,w,{}),record=h.data.get(saved.manifest.sheets[0].key);
    assert.equal(record.gz,compressed);assert.ok(record.cellPartKeys.length>=2);
    const out=new Workbook((await loadLargeWorkbook('unicode-'+compressed)).workbook);
    for(let r=0;r<4;r++){assert.equal(out.getRaw(0,r,0),w.getRaw(0,r,0));assert.deepEqual(out.styleAt(0,r,0),w.styleAt(0,r,0));}
  }finally{globalThis.CompressionStream=previous;}
}));

for(const failure of['write','read'])test('v5 compression '+failure+' failure settles, cleans staged parts and preserves the previous generation',()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const w=new Workbook({sheets:[{name:'Original',cells:{'0,0':{raw:'17'}}}]}),key='codec-'+failure;
  const first=await saveLargeWorkbook(key,w,{}),keys=[...h.data.keys()].sort(),previous=globalThis.CompressionStream;
  w.transact(()=>w.setInput(0,0,0,'18'));const error=new Error('injected-'+failure);let released=0,aborted=0;
  try {
    if(failure==='write')globalThis.CompressionStream=class{constructor(){return new TransformStream({transform(){throw error;}});}};
    else globalThis.CompressionStream=class{constructor(){let rejectWrite;this.readable=new ReadableStream({start(c){queueMicrotask(()=>c.error(error));}});this.writable={getWriter:()=>({write:()=>new Promise((_resolve,reject)=>{rejectWrite=reject;}),close:async()=>{},abort:async e=>{aborted++;rejectWrite?.(e);},releaseLock:()=>{released++;}})};}};
    await assert.rejects(saveLargeWorkbook(key,w,{}),e=>e===error);
    if(failure==='read'){assert.ok(aborted>0);assert.equal(released,1);}
    assert.deepEqual([...h.data.keys()].sort(),keys);assert.equal(h.data.get(key).generation,first.manifest.generation);
  }finally{globalThis.CompressionStream=previous;}
  assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getRaw(0,0,0),'17');
}));


test('v5 captures sheet metadata once and compares without repeated clone/string allocation',()=>fixture(async(_s,h)=>{
  const{saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const w=new Workbook({sheets:[{name:'Metadata',cells:{'0,0':{raw:'1'}},images:[{src:'data:image/png;base64,'+'x'.repeat(200000)}],page:{orientation:'portrait',margins:{top:0.5}},hiddenRows:{__bits:new Uint8Array([0,1,0]),start:1}}]});
  const original=w.sheetMeta;let captures=0;w.sheetMeta=function(si){captures++;return original.call(this,si);};
  const first=await saveLargeWorkbook('metadata-once',w,{});assert.equal(captures,1);assert.equal(first.isCurrent(),true);assert.equal(captures,1);
  const restored=new Workbook((await loadLargeWorkbook('metadata-once')).workbook);assert.deepEqual(restored.sheetMeta(0),original.call(w,0));
  w.cellRunChunks=function*(){assert.fail('unchanged metadata must reuse cells');};
  const second=await saveLargeWorkbook('metadata-once',w,{});assert.equal(second.manifest.sheets[0].key,first.manifest.sheets[0].key);assert.equal(captures,2);
}));

test('v5 unversioned nested metadata changes abort and keep previous generation',()=>fixture(async(_s,h)=>{
  const{saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  for(const kind of['page','image','typed','fileValues']) {
    const w=new Workbook({sheets:[{name:'Metadata',cells:{'0,0':{raw:'1'}},page:{orientation:'portrait',margins:{top:0.5}},images:[{src:'data:test'}],hiddenRows:{__bits:new Uint8Array([0,1,0]),start:1}}]});
    const key='metadata-'+kind,first=await saveLargeWorkbook(key,w,{}),keys=[...h.data.keys()].sort();
    w.transact(()=>w.setInput(0,0,0,'2'));let once=false;
    await assert.rejects(saveLargeWorkbook(key,w,{}, {waitForIdle:()=>{if(once)return;once=true;const s=w.sheets[0];if(kind==='page')s.page.margins.top=2;else if(kind==='image')s.images[0].src='data:changed';else if(kind==='typed')s.hiddenRows.__bits[1]=0;else s.fileValues=!s.fileValues;}}),{code:'BIG_SAVE_ABORT'});
    assert.equal(h.data.get(key).generation,first.manifest.generation);assert.deepEqual([...h.data.keys()].sort(),keys);assert.equal(new Workbook((await loadLargeWorkbook(key)).workbook).getRaw(0,0,0),'1');
  }
}));


for(const date1904 of [false,true])test('prepared IDB restore adopts cells once without changing stored semantics: '+date1904,()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');
  const {Workbook}=await import('../src/workbook.js');
  const key='prepared-'+date1904,style={font:'Arial',size:10,bold:true,numFmt:'custom',code:'0.00'};
  const input={date1904,sheets:[{name:'Prepared',fileValues:true,cells:{
    '0,0':{raw:'17',style},'1,0':{raw:'TRUE'},'2,0':{raw:'2024-02-29'},'3,0':{raw:'0017',inputType:'text'},
    '4,0':{raw:'=A1+1',cached:18,comment:'note',link:'#Prepared!A1'},
    '5,0':{raw:'=A1+2',staleCached:19},'6,0':{raw:'=SEQUENCE(2)',cached:1,cachedArray:{h:2,w:1,values:[0,0,1,1,0,2]}},
    '8,0':{raw:'=SEQUENCE(2)',staleCachedArray:{h:2,w:1,values:[0,0,3,1,0,4]}},
    '10,0':{raw:'#N/A'},'11,0':{raw:'한글😀',phonetic:{runs:[{text:'reading',start:0,end:2}]}},
    '12,0':{raw:'',image:{src:'data:image/png;base64,YQ=='}},'13,0':{raw:'=1+2',fx:true,style:{numFmt:'text'}},
    '14,0':{raw:'',style:{fill:'#aabbcc'}}
  },images:[{src:'data:image/png;base64,YQ==',x:12,y:24,w:32,h:48}],page:{orientation:'landscape',margins:{top:0.5}},freeze:{rows:1,cols:2},hiddenRows:{__bits:new Uint8Array([0,1,0]),start:1}},
  {name:'Other',cells:{'0,0':{raw:'=Prepared!A1',cached:17}}}]};
  const original=new Workbook(input);original.sheets[0].cells.setRunRC(30,1,1200,{raw:'',style:{fill:'#abcdef'}});
  original.sheets[0].blocks=[{r0:2000,c0:2,n:3,cols:[{num:new Float64Array([3,NaN,NaN]),str:new Int32Array([-1,0,1]),dict:['text',true],fmt:{numFmt:'number'}}],perm:new Int32Array([2,0,1])}];
  original.pivotCacheItems={cache:{fields:[{name:'Date',shared:[45351,7],sharedTypes:'dn',format:{numFmt:'date'}},{name:'EmptyDate',shared:[],dateOnly:true}]}};
  const saved=await saveLargeWorkbook(key,original,{}),stored=h.data.get(saved.manifest.sheets[0].key),storedPage=structuredClone(stored.meta.page);
  const normal=new Workbook((await loadLargeWorkbook(key)).workbook),loaded=await loadLargeWorkbook(key,null,{prepareCells:true}),data=loaded.workbook;
  const cells=data.sheets[0].cells,formula=cells.getRC(4,0),page=data.sheets[0].page,blocks=data.sheets[0].blocks;
  assert.equal(formula.formula,true);assert.equal(cells.getRC(5,0).dirty,true);assert.equal(cells.getRC(8,0).dirty,true);
  cells.mapValues=function*(){assert.fail('prepared cells must not be normalized twice');};
  const actual=new Workbook(data);
  assert.equal(actual.sheets[0].cells,cells);assert.equal(actual.getCell(0,4,0),formula);assert.equal(actual.sheets[0].page,page);assert.equal(actual.sheets[0].blocks,blocks);
  assert.equal(data.sheets[0].cells,null);assert.equal(data.sheets[0].blocks,null);assert.equal(data.sheets[0].page,undefined);assert.equal(actual.sheets[0]._hasCachedArrays,true);
  assert.deepEqual(actual.serialize(),normal.serialize());
  for(const [r,c]of[[0,0],[1,0],[2,0],[3,0],[4,0],[5,0],[10,0],[13,0],[1229,1],[2000,2],[2001,2],[2002,2]])assert.deepEqual(actual.getValue(0,r,c),normal.getValue(0,r,c));
  actual.transact(()=>{actual.setInput(0,0,0,'42');actual.setSheetProp(0,'page',{orientation:'portrait'});});
  assert.equal(actual.getValue(0,0,0),42);actual.undo();assert.equal(actual.getValue(0,0,0),17);actual.redo();assert.equal(actual.getValue(0,0,0),42);
  assert.deepEqual(stored.meta.page,storedPage);assert.equal(normal.getValue(0,0,0),17);
  await saveLargeWorkbook(key,actual,{});const reopened=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);
  assert.deepEqual(reopened.serialize(),actual.serialize());
}));

test('prepared restore rejects copied owners, consumed input, replaced maps and date changes before consuming any sheet',()=>fixture(async()=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const {Workbook}=await import('../src/workbook.js');const {CellMap}=await import('../src/cellmap.js');
  const key='prepared-owner',original=new Workbook({date1904:true,sheets:[{name:'One',cells:{'0,0':{raw:'17'}}},{name:'Two',cells:{'0,0':{raw:'2024-02-29'}}}]});await saveLargeWorkbook(key,original,{});
  const current=new Workbook({sheets:[{name:'Current',cells:{'0,0':{raw:'999'}}}]}),expected=current.serialize();
  for(const kind of['copy-book','copy-sheet','map','date','order']){
    const data=(await loadLargeWorkbook(key,null,{prepareCells:true})).workbook,first=data.sheets[0].cells,second=data.sheets[1].cells;
    let candidate=data;
    if(kind==='copy-book')candidate={...data};
    if(kind==='copy-sheet')candidate={sheets:[{...data.sheets[0]}]};
    if(kind==='map')data.sheets[1].cells=new CellMap();
    if(kind==='date')data.date1904=false;
    if(kind==='order')data.sheets.reverse();
    assert.throws(()=>current.restore(candidate),{code:'PREPARED_RESTORE_INVALID'});assert.deepEqual(current.serialize(),expected);
    if(kind==='map')data.sheets[1].cells=second;if(kind==='date')data.date1904=true;if(kind==='order')data.sheets.reverse();
    assert.equal(data.sheets[0].cells,first);assert.equal(data.sheets[1].cells,second);
    const restored=new Workbook(data);assert.equal(restored.getValue(0,0,0),17);
    assert.throws(()=>current.restore(data),{code:'PREPARED_RESTORE_INVALID'});assert.deepEqual(current.serialize(),expected);
    assert.throws(()=>new Workbook({sheets:[{name:'Alias',cells:first}]}),{code:'PREPARED_RESTORE_INVALID'});assert.equal(restored.getValue(0,0,0),17);
  }
}));

test('two fresh prepared IDB reads create independent editable workbooks and ignore serialized preparation flags',()=>fixture(async()=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const {Workbook}=await import('../src/workbook.js');
  const key='prepared-independent',original=new Workbook({sheets:[{name:'Separate',cells:{'0,0':{raw:'17',style:{fill:'#abcdef'}},'1,0':{raw:'=A1+1',cached:18}},page:{margins:{top:0.5}}}]});await saveLargeWorkbook(key,original,{});
  const first=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook),second=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);
  assert.notEqual(first.sheets[0].cells,second.sheets[0].cells);assert.notEqual(first.sheets[0].page,second.sheets[0].page);assert.notEqual(first.getCell(0,1,0),second.getCell(0,1,0));
  first.transact(()=>{first.setInput(0,0,0,'99');first.setStyle(0,0,0,{fill:'#112233'});first.setSheetProp(0,'page',{margins:{top:2}});});
  assert.equal(second.getValue(0,0,0),17);assert.equal(second.styleAt(0,0,0).fill,'#abcdef');assert.equal(second.sheets[0].page.margins.top,0.5);
  const fake=new Workbook({prepared:true,date1904:false,sheets:[{name:'Fake',prepared:true,cells:{'0,0':{raw:'=1+2',cached:3,prepared:true}}}]});assert.equal(fake.getValue(0,0,0),3);
}));

for(const v of[2,3,4])test('prepared loader preserves legacy inline cells and runs: v'+v,()=>fixture(async(_s,h)=>{
  const {loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const key='prepared-legacy'+v,recordKey=key+'#sheet',runs=v===4,entries=runs?[[2,3,1,{raw:'=1+2',staleCached:3}],[4,3,200,{raw:'',style:{bold:true}}]]:[['2,3',{raw:'=1+2',staleCached:3}],['3,3',{raw:'2024-02-29'}]];
  h.data.set(key,{v,generation:'old',book:{date1904:true},sheets:[{id:'sheet',key:recordKey,ev:0}]});
  h.data.set(recordKey,{meta:{name:'Legacy',page:{orientation:'landscape'}},chunks:[new Blob([JSON.stringify(entries)])],cellEncoding:runs?'runs-v1':undefined,gz:false,blocks:[],partKeys:[]});
  const normal=new Workbook((await loadLargeWorkbook(key)).workbook),prepared=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);
  assert.deepEqual(prepared.serialize(),normal.serialize());assert.equal(prepared.getCell(0,2,3).dirty,true);assert.equal(prepared.date1904,true);
}));

test('failed prepared IDB read leaves its committed generation and current workbook intact',()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const key='prepared-failure',original=new Workbook({sheets:[{name:'One',cells:{'0,0':{raw:'17'}}},{name:'Two',cells:{'0,0':{raw:'19'}}}]}),saved=await saveLargeWorkbook(key,original,{}),expected=original.serialize();
  const record=h.data.get(saved.manifest.sheets[1].key),partKey=record.cellPartKeys[0],part=h.data.get(partKey);h.data.delete(partKey);
  await assert.rejects(loadLargeWorkbook(key,null,{prepareCells:true}),/셀 조각/);assert.equal(h.data.get(key).generation,saved.manifest.generation);assert.deepEqual(original.serialize(),expected);
  h.data.set(partKey,part);const out=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);assert.deepEqual(out.serialize(),expected);
}));


test('IDB captures one workbook metadata snapshot and validates live views without repeated style clones',()=>fixture(async(_s,h)=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const w=new Workbook({cellStyles:Array.from({length:1200},(_,i)=>({name:'S'+i,style:{fill:'#abcdef',font:'Arial',size:11}})),objectStyles:{table:[]},themeEffects:{shadow:{color:'#123456'}},names:[{name:'Named',ref:'=A1'}],sheets:[{name:'Meta',cells:{'0,0':{raw:'17'}}}]});
  w.names[0]._ast={cached:true};w.names[0]._text='cache';
  const snapshot=w.bookMeta(false,false),view=w.bookMeta(false,false,{shareData:true});
  assert.deepEqual(view,snapshot);assert.equal(view.cellStyles,w.cellStyles);assert.equal(view.objectStyles,w.objectStyles);assert.equal(view.themeEffects,w.themeEffects);
  assert.notEqual(snapshot.cellStyles,w.cellStyles);assert.notEqual(snapshot.cellStyles[0].style,w.cellStyles[0].style);assert.equal(view.names[0]._ast,undefined);assert.equal(view.names[0]._text,undefined);
  const clone=globalThis.structuredClone;let repeatedStyleClone=0;
  globalThis.structuredClone=value=>{if(value===w.cellStyles||value===w.objectStyles||value===w.themeEffects)repeatedStyleClone++;return clone(value);};
  try{const first=await saveLargeWorkbook('book-meta-view',w,{});assert.equal(first.isCurrent(),true);assert.equal(first.isCurrent(),true);assert.equal(repeatedStyleClone,0);}finally{globalThis.structuredClone=clone;}
  const restored=new Workbook((await loadLargeWorkbook('book-meta-view')).workbook);assert.deepEqual(restored.bookMeta(false,false),snapshot);
  for(const kind of['cellStyle','effect','name']){
    const key='book-meta-change-'+kind,first=await saveLargeWorkbook(key,w,{}),prior=w.bookMeta(false,false);w.transact(()=>w.setInput(0,0,0,String(20+w.version)));let once=false;
    await assert.rejects(saveLargeWorkbook(key,w,{}, {waitForIdle:()=>{if(once)return;once=true;if(kind==='cellStyle')w.cellStyles[0].style.fill='#445566';else if(kind==='effect')w.themeEffects.shadow.color='#654321';else w.names[0].ref='=B2';}}),{code:'BIG_SAVE_ABORT'});
    assert.equal(h.data.get(key).generation,first.manifest.generation);assert.deepEqual(new Workbook((await loadLargeWorkbook(key)).workbook).bookMeta(false,false),prior);
  }
}));


test('prepared exact formula text sharing keeps cell caches, notes and edits independent',()=>fixture(async()=>{
  const {saveLargeWorkbook,loadLargeWorkbook}=await import('../src/big-storage.js');const{Workbook}=await import('../src/workbook.js');
  const raw='=IF("한글😀"="한글😀",ROW()+1,0)',cells={};
  for(let r=0;r<450;r++)cells[r+',0']={raw,comment:'note'+r,link:'#Repeated!A'+(r+1),cached:r+2,style:{fill:r%2?'#112233':'#445566'}};
  cells['0,1']={raw,style:{numFmt:'text'},inputType:'text'};
  const w=new Workbook({sheets:[{name:'Repeated',cells,fileValues:true}]}),key='exact-formula-raw';
  await saveLargeWorkbook(key,w,{});
  const normal=new Workbook((await loadLargeWorkbook(key)).workbook),prepared=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);
  assert.deepEqual(prepared.serialize(),normal.serialize());
  const first=prepared.getCell(0,0,0),second=prepared.getCell(0,1,0);
  assert.notEqual(first,second);assert.equal(first.raw,raw);assert.equal(second.raw,raw);assert.equal(first.cached,2);assert.equal(second.cached,3);assert.equal(first.comment,'note0');assert.equal(second.comment,'note1');
  assert.equal(prepared.getValue(0,0,1),raw);assert.equal(prepared.getCell(0,0,1).formula,undefined);
  prepared.transact(()=>prepared.setInput(0,0,0,'=123'));
  assert.equal(prepared.getValue(0,0,0),123);assert.equal(prepared.getRaw(0,1,0),raw);assert.equal(prepared.getCell(0,1,0).comment,'note1');
  prepared.undo();assert.equal(prepared.getRaw(0,0,0),raw);prepared.redo();assert.equal(prepared.getRaw(0,0,0),'=123');
  await saveLargeWorkbook(key,prepared,{});const out=new Workbook((await loadLargeWorkbook(key,null,{prepareCells:true})).workbook);
  assert.deepEqual(out.serialize(),prepared.serialize());assert.equal(out.getRaw(0,449,0),raw);assert.equal(out.getCell(0,449,0).comment,'note449');
}));
