import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { jsonPartsBlob, librarySnapshot, equalByteStreams } from '../src/snapshot-blob.js';
import { packText, unpackText } from '../src/library.js';

const meta = { app: 'wixel', docName: '복원 "문서" 😀', docId: 'fixture-original', si: 1 };
function fixture() {
  const w = new Workbook({ date1904: true, calculation: { mode: 'manual', iterate: true }, props: { title: '제목', lockStructure: true },
    names: [{ name: '목록', ref: '원본!$A$1:$A$2' }], objectStyles: { table: [{ name: '사용자' }] },
    sheets: [{ name: '원본', cells: {
      '0,0': { raw: '한글😀\ud800\n"\\' }, '1,0': { raw: '=UNSUPPORTED(1)', cached: 42 },
      '2,0': { raw: '=A1', staleCached: '이전', style: { bold: true, numFmt: 'text' }, fx: true },
      '3,0': { raw: '001', inputType: 'text', comment: '설명', link: 'https://example.org/', image: {src:'data:image/png;base64,AAAA'}, phonetic:{text:'윗주'} },
    }, state: 'hidden', page: { printArea: 'A1:C10' }, view: { mode: 'pageBreakPreview' },
    charts: [{type:'column',title:'합성'}], images:[{src:'data:image/png;base64,'+'A'.repeat(150000)}],
    filter: {r1:0,c1:0,r2:5,c2:2,cols:{0:{values:['한글']}}}, tables:[{name:'표',r1:0,c1:0,r2:5,c2:2}],
    pivot:{rows:['이름']}, slicers:[{field:'이름'}], validations:[{type:'list',f1:'"가,나"'}],
    blocks:[{r0:20,c0:0,n:3,ver:2,perm:new Int32Array([2,0,1]),cols:[
      {num:new Float64Array([3,NaN,0]),str:new Int32Array([-1,0,-1]),dict:['문자'],fmt:{numFmt:'comma'}},
      {num:null,str:new Int32Array([0,1,2]),dict:[true,false,{error:'#N/A'}],fmt:null}
    ]}] }, {name:'두 번째',cells:{'0,0':{raw:'빈 시트 아님'}}}] });
  w.sheets[0].fileValues = true;
  return w;
}

test('Blob 직렬화는 기존 전체 직렬화와 byte 동일하고 모델·Undo를 변경하지 않는다', async () => {
  const w=fixture(), expected=JSON.stringify({ ...meta, workbook:w.serialize() }), version=w.version, undo=w.undoStack.length;
  assert.equal(await librarySnapshot(w,meta).blob.text(),expected);
  assert.equal(w.version,version);assert.equal(w.undoStack.length,undo);
  assert.equal(JSON.stringify({...meta,workbook:w.serialize()}),expected);
});

test('셀 청크 경계·모든 셀과 시트 순서를 유지하고 serialize를 호출하지 않는다',async()=>{
  const w=new Workbook();for(let i=0;i<5003;i++)w.setInput(0,i%1000,Math.floor(i/1000),String(i));
  const expected=JSON.stringify(w.serialize());w.serialize=()=>{throw new Error('전체 복제 금지');};
  assert.equal(await w.serializeBlob().text(),expected);
});

test('동기 스냅숏 뒤 원본 교체·메타 변경·충돌 ID 교체에도 원래 내용 유지',async()=>{
  const w=fixture(), header={...meta}, expected=JSON.stringify({...header,workbook:w.serialize()}), snapshot=librarySnapshot(w,header);
  w.restore({sheets:[{name:'새 문서',cells:{'0,0':{raw:'나중'}}}]});header.docName='나중 이름';
  assert.equal(await snapshot.blob.text(),expected);
  assert.equal(await snapshot.withDocId('copy').text(),JSON.stringify({...JSON.parse(expected),docId:'copy'}));
});

test('기존 JSON 블록 저장본 복원: 정렬·0·빈칸·문자·논리값·오류',async()=>{
  const w=fixture(), expected=[];for(let r=20;r<23;r++)for(let c=0;c<2;c++)expected.push(w.getValue(0,r,c));
  for(const payload of [JSON.stringify(w.serialize()),await w.serializeBlob().text()]){
    const restored=new Workbook(JSON.parse(payload)), actual=[];
    for(let r=20;r<23;r++)for(let c=0;c<2;c++)actual.push(restored.getValue(0,r,c));
    assert.deepEqual(actual,expected);assert.ok(Number.isNaN(restored.sheets[0].blocks[0].cols[0].num[1]));
    assert.equal(restored.date1904,true);
    restored.transact(()=>restored.setInput(0,20,0,'99'));assert.equal(restored.getValue(0,20,0),99);
    restored.undo();assert.equal(restored.getValue(0,20,0),w.getValue(0,20,0));
  }
});

test('Blob 청크는 한글·서로게이트 쌍을 경계에서 손상하지 않는다',async()=>{
  const text=JSON.stringify({s:'😀a한글'.repeat(1000),u:'\ud800'});
  for(const size of [2,3,7,65536])for(const split of [1,2,7,63]){
    const parts=[];for(let i=0;i<text.length;i+=split)parts.push(text.slice(i,i+split));
    assert.equal(await jsonPartsBlob(parts,size).text(),text);
  }
  assert.throws(()=>jsonPartsBlob(['x'],1),RangeError);
});

test('gzip와 압축 미지원의 Blob/문자열은 기존 JSON으로 풀린다',async()=>{
  const value=await fixture().serializeBlob().text(), gzip=globalThis.CompressionStream;
  for(const supported of [true,false])try{
    if(!supported)globalThis.CompressionStream=undefined;
    for(const input of [value,new Blob([value])])assert.equal(await unpackText(await packText(input)),value);
  }finally{globalThis.CompressionStream=gzip;}
});

function bytes(parts,fail=false,onCancel=()=>{}){
  let i=0;return new ReadableStream({pull(c){if(i<parts.length)c.enqueue(Uint8Array.from(parts[i++]));else if(fail)c.error(new Error('읽기 실패'));else c.close();},cancel:onCancel});
}
test('스트림 비교: 다른 조각 크기·빈 조각·prefix·불일치·오류',async()=>{
  assert.equal(await equalByteStreams(bytes([[],[1,2],[3],[]]),bytes([[1],[2,3]])),true);
  assert.equal(await equalByteStreams(bytes([]),bytes([[]])),true);
  for(const pair of [[[1,2],[1]],[[1],[1,2]],[[1,2],[1,3]],[[],[1]]])assert.equal(await equalByteStreams(bytes([pair[0]]),bytes([pair[1]])),false);
  await assert.rejects(equalByteStreams(bytes([],true),bytes([[1]])),/읽기 실패/);
});
