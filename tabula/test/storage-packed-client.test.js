import test from 'node:test';
import assert from 'node:assert/strict';
import { server } from '../src/storage.js';
import { Workbook } from '../src/workbook.js';
import { writeWixelFile } from '../src/wixel-file.js';
import { packPublishedBlob } from '../src/publish.js';

const KEY_A='A'.repeat(43),KEY_B='B'.repeat(42)+'A';
function fixture(){return new Workbook({date1904:true,sheets:[{name:'합성',fileValues:true,cells:{'0,0':{raw:'=UNSUPPORTED(7)',cached:42},'1,0':{raw:'001',inputType:'text',style:{bold:true}},'2,0':{raw:'오류 아님'}},slicers:[{name:'분류',selected:['가'],style:'사용자색'}],pivot:{name:'피벗',rows:['분류']}}]});}
async function payload(gzip=true) {return packPublishedBlob(await writeWixelFile(fixture(),{app:'wixel',docId:'fixture',docName:'합성',si:0},{gzip}));}
function response(data,headers={},status=200){return new Response(JSON.stringify(data),{status,headers:{'Content-Type':'application/json',...headers}});}
async function isolated(run) {
  const originalFetch=globalThis.fetch,setting=Object.getOwnPropertyDescriptor(globalThis,'localStorage'),decompress=globalThis.DecompressionStream;
  const values=new Map();Object.defineProperty(globalThis,'localStorage',{configurable:true,value:{getItem:k=>values.get(k)??null,setItem:(k,v)=>values.set(k,v)}});
  server.disconnect();server.available=true;server.vault=true;server.capabilities={versionHistory:true};server.connect(KEY_A);
  try{return await run();}finally{server.disconnect();server.available=false;server.vault=false;server.capabilities={};globalThis.fetch=originalFetch;globalThis.DecompressionStream=decompress;if(setting)Object.defineProperty(globalThis,'localStorage',setting);else delete globalThis.localStorage;}
}

test('압축 온라인 저장: Blob을 JSON.stringify하지 않고 인증·CAS와 함께 보낸다',()=>isolated(async()=>{
  const blob=await payload(),calls=[];
  server.restoreRevision('files/%ED%95%A9%EC%84%B1','8');
  globalThis.fetch=async(url,options)=>{calls.push({url,options});return response({ok:true,revision:9},{'X-Wixel-Revision':'9'});};
  await server.save('합성',blob);
  assert.equal(calls.length,1);assert.equal(calls[0].options.body,blob);assert.equal(calls[0].options.method,'PUT');
  assert.equal(calls[0].options.headers['X-Wixel-Vault'],KEY_A);assert.equal(calls[0].options.headers['If-Match'],'"8"');
  assert.equal(calls[0].options.headers['Content-Type'],'application/json');assert.equal(server.revision('files/%ED%95%A9%EC%84%B1'),'9');
}));
test('압축 게시: 새 링크·같은 링크 재게시도 Blob과 현재 버전을 보존한다',()=>isolated(async()=>{
  const blob=await payload(),calls=[];
  globalThis.fetch=async(url,options)=>{calls.push({url,options});return response({id:'pub-fixture',revision:calls.length+2});};
  const first=await server.publish(blob);assert.equal(first.id,'pub-fixture');assert.equal(server.revision('published/pub-fixture'),'3');
  await server.publish(blob,first.id);assert.equal(calls[0].options.method,'POST');assert.equal(calls[1].options.method,'PUT');
  for(const call of calls)assert.equal(call.options.body,blob);assert.equal(calls[1].options.headers['If-Match'],'"3"');
}));
test('압축 온라인 문서·버전·게시 읽기: 압축 여부와 무관하게 값·서식·수식·슬라이서 복원',()=>isolated(async()=>{
  const expected=JSON.stringify(fixture().serialize());
  for(const gzip of [true,false]) {
    const packed=JSON.parse(await(await payload(gzip)).text());
    globalThis.fetch=async()=>response(packed,{'X-Wixel-Revision':'7','X-Modified':'1234'});
    for(const data of [await server.load('합성'),await server.loadVersion('합성',6),(await server.published('pub-fixture')).data]) {
      assert.equal(data.docId,'fixture');const wb=new Workbook(data.workbook);
      assert.equal(JSON.stringify(wb.serialize()),expected);assert.equal(wb.getValue(0,0,0),42);
      assert.equal(wb.sheets[0].slicers[0].style,'사용자색');
    }
    assert.equal((await server.published('pub-fixture')).modified,1234);
  }
}));
test('이전 비압축 온라인 JSON과 오류 응답 상태는 그대로 유지한다',()=>isolated(async()=>{
  const data={docName:'이전 형식',workbook:fixture().serialize()};globalThis.fetch=async()=>response(data);
  assert.deepEqual(await server.load('기존'),JSON.parse(JSON.stringify(data)));
  globalThis.fetch=async()=>response({error:'다른 탭의 최신 버전',code:'REVISION_CONFLICT',currentRevision:12},{},412);
  await assert.rejects(server.save('합성',await payload()),{status:412,code:'REVISION_CONFLICT',currentRevision:12,message:'다른 탭의 최신 버전'});
}));
test('응답 JSON을 읽는 동안 보관함이 바뀌면 압축 해제 전 중단하고 새 보관함 버전을 덮지 않는다',()=>isolated(async()=>{
  const packed=JSON.parse(await(await payload()).text());let decompressCount=0;
  const Real=globalThis.DecompressionStream;globalThis.DecompressionStream=class{constructor(kind){decompressCount++;return new Real(kind);}};
  globalThis.fetch=async()=>({ok:true,headers:new Headers({'X-Wixel-Revision':'77'}),json:async()=>{server.connect(KEY_B);return packed;}});
  await assert.rejects(server.load('합성'),{code:'CONNECTION_CHANGED'});
  assert.equal(decompressCount,0);assert.equal(server.revision('files/%ED%95%A9%EC%84%B1'),'0');
}));
test('압축 해제 도중 보관함 변경도 복원을 중단하고 잘못된 버전을 채택하지 않는다',()=>isolated(async()=>{
  const packed=JSON.parse(await(await payload()).text());const Real=globalThis.DecompressionStream;
  globalThis.DecompressionStream=class{constructor(kind){server.connect(KEY_B);return new Real(kind);}};
  globalThis.fetch=async()=>response(packed,{'X-Wixel-Revision':'77'});
  await assert.rejects(server.load('합성'),{code:'CONNECTION_CHANGED'});assert.equal(server.revision('files/%ED%95%A9%EC%84%B1'),'0');
}));
test('문서 전환 guard는 보내기 전과 읽기 도중 모두 중단하며 fetch로 전달되지 않는다',()=>isolated(async()=>{
  let calls=0;globalThis.fetch=async()=>{calls++;throw new Error('호출되면 안 됨');};
  await assert.rejects(server.save('합성',await payload(),{isCurrent:()=>false}),{code:'DOCUMENT_OPEN_CANCELLED'});assert.equal(calls,0);
  const packed=JSON.parse(await(await payload()).text());let current=true;
  globalThis.fetch=async(url,options)=>{assert.equal(Object.hasOwn(options,'isCurrent'),false);return {ok:true,headers:new Headers(),json:async()=>{current=false;return packed;}};};
  await assert.rejects(server.load('합성',{isCurrent:()=>current}),{code:'DOCUMENT_OPEN_CANCELLED'});
  current=true;await assert.rejects(server.published('pub-fixture',{isCurrent:()=>current}),{code:'DOCUMENT_OPEN_CANCELLED'});
}));
test('공개 게시 응답은 보관함 전환과 무관하지만 문서 전환 중에는 중단한다',()=>isolated(async()=>{
  const packed=JSON.parse(await(await payload()).text());
  globalThis.fetch=async()=>({ok:true,headers:new Headers({'X-Modified':'10'}),json:async()=>{server.connect(KEY_B);return packed;}});
  assert.equal((await server.published('pub-fixture')).data.docName,'합성');
  let current=true;const Real=globalThis.DecompressionStream;globalThis.DecompressionStream=class{constructor(kind){current=false;return new Real(kind);}};
  await assert.rejects(server.published('pub-fixture',{isCurrent:()=>current}),{code:'DOCUMENT_OPEN_CANCELLED'});
}));
test('손상된 압축 게시 조각을 받으면 문서·버전을 성공으로 채택하지 않는다',()=>isolated(async()=>{
  globalThis.fetch=async()=>response({format:'wixel-packed',version:1,compression:'gzip',chunks:['bad!']},{'X-Wixel-Revision':'77'});
  await assert.rejects(server.load('합성'),/손상/);assert.equal(server.revision('files/%ED%95%A9%EC%84%B1'),'0');
}));
test('게시 버전 헤더를 읽는 중 보관함이 바뀌면 새 보관함에 과거 버전을 기록하지 않는다',()=>isolated(async()=>{
  globalThis.fetch=async()=>({ok:true,headers:new Headers({'X-Wixel-Revision':'77'}),body:{cancel:async()=>{server.connect(KEY_B);}}});
  await assert.rejects(server.publicationRevision('pub-fixture'),{code:'CONNECTION_CHANGED'});assert.equal(server.revision('published/pub-fixture'),'0');
}));
