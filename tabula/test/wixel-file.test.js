import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { writeWixelFile, readWixelFile, WIXEL_FILE_MAGIC } from '../src/wixel-file.js';
import { PivotSnapshotBuilder, pivotSnapshotValue } from '../src/pivot-cache-data.js';
import { publishedWixelFile, packPublishedBlob, unpackPublishedBlob, packedPublishedSize, assertPublishLinkSize } from '../src/publish.js';
import { jsonValueParts } from '../src/json-parts.js';
import { jsonReplacer } from '../src/block.js';

function fixture() {
  const book=new Workbook({date1904:true,props:{title:'합성 문서'},objectStyles:{slicer:[{name:'내 스타일',fill:'#cc1122'}]},
    names:[{name:'목록',ref:'원본!A1:A2'}],sheets:[{name:'원본',fileValues:true,cells:{
      '0,0':{raw:'한글😀\ud800'},'1,0':{raw:'=UNSUPPORTED(1)',cached:23,style:{bold:true}},
      '2,0':{raw:'001',inputType:'text',comment:'설명',link:'https://example.org/'}},
      pivot:{name:'피벗',source:'원본',range:{r1:0,c1:0,r2:2,c2:1},snapshotId:'cache',rows:['항목']},
      slicers:[{id:'slicer',caption:'선택',selected:['가'],style:'내 스타일'}],
      page:{printArea:'A1:C20'},cond:[{type:'gt',v1:10,style:{fill:'#ff0000'}}],
      images:[{src:'data:image/png;base64,'+'x'.repeat(70000)}],
      blocks:[{r0:20,c0:0,n:3,ver:2,perm:new Int32Array([2,0,1]),cols:[{num:new Float64Array([0,NaN,7]),str:new Int32Array([-1,0,-1]),dict:['문자'],fmt:{bold:true}}]}]},
      {name:'숨김',state:'hidden',cells:{'0,0':{raw:'비공개'}}}]});
  book.sheets[0].cells.setRunRC(50,2,1000,Object.freeze({raw:'',style:Object.freeze({fill:'#ccff22'})}));
  const b=new PivotSnapshotBuilder(['항목','값']);b.add(['가',12]);b.add(['나',null]);b.add(['다',0]);
  book.pivotSnapshots=new Map([['cache',{rows:b.finish(),ver:undefined}]]);return book;
}
const meta={app:'wixel',docName:'보존',docId:'synthetic-id',si:1,view:{headers:false,grid:true}};

test('WIXEL v2: 압축/비압축 왕복은 모든 셀·서식·피벗·슬라이서·숨김·블록·캐시를 유지',async()=>{
  for(const gzip of [true,false]) {
    const book=fixture(),before=JSON.stringify(book.serialize()),version=book.version;
    const blob=await writeWixelFile(book,meta,{gzip}),data=await readWixelFile(blob),back=new Workbook(data.workbook);
    assert.equal(JSON.stringify(back.serialize()),before);assert.equal(book.version,version);
    assert.equal(data.docName,meta.docName);assert.deepEqual(data.view,meta.view);
    assert.equal(back.getValue(0,20,0),7);assert.equal(back.getValue(0,21,0),0);assert.equal(back.getValue(0,22,0),'문자');
    assert.equal(back.sheets[0].slicers[0].style,'내 스타일');assert.equal(back.sheets[1].state,'hidden');
    assert.equal(pivotSnapshotValue(back.pivotSnapshots.get('cache').rows,1,1),null);
    assert.equal(pivotSnapshotValue(back.pivotSnapshots.get('cache').rows,2,1),0);
  }
});
test('WIXEL v2: 전체 serialize/serializeBlob 복제를 호출하지 않고 셀 청크 경계를 읽는다',async()=>{
  const book=new Workbook();for(let r=0;r<7000;r++)book.setInput(0,r,0,String(r));
  book.serialize=book.serializeBlob=()=>{throw new Error('전체 복제 금지');};
  const data=await readWixelFile(await writeWixelFile(book,meta));const back=new Workbook(data.workbook);
  assert.equal(back.sheets[0].cells.size,7000);for(const r of [0,2047,2048,4095,6999])assert.equal(back.getValue(0,r,0),r);
});
test('WIXEL v2: 큰 단일 셀과 UTF-16 경계도 정확히 복원한다',async()=>{
  const book=new Workbook(),text=('가😀\ud800').repeat(80000);book.setInput(0,0,0,text);
  const blob=await writeWixelFile(book,meta,{gzip:false}),lines=(await blob.text()).split('\n');
  assert.ok(lines.every(line=>line.length<250000));
  const back=new Workbook((await readWixelFile(blob)).workbook);assert.equal(back.getCell(0,0,0).raw,text);
});
test('WIXEL v2: 중간 편집/취소는 성공한 저장으로 돌려주지 않는다',async()=>{
  const book=fixture();await assert.rejects(writeWixelFile(book,meta,{onProgress(){book.version++;}}),e=>e.code==='WIXEL_FILE_ABORT');
  await assert.rejects(writeWixelFile(fixture(),meta,{isCurrent:()=>false}),e=>e.code==='WIXEL_FILE_ABORT');
  const blob=await writeWixelFile(fixture(),meta);await assert.rejects(readWixelFile(blob,{isCurrent:()=>false}),e=>e.code==='WIXEL_FILE_ABORT');
});
test('WIXEL v2: 잘린 파일·완료 개수 불일치·추가 내용·중복 키를 거부한다',async()=>{
  const text=await(await writeWixelFile(fixture(),meta,{gzip:false})).text();
  for(const changed of [text.slice(0,-20),text.replace(/\["finish",\d+\]/,'["finish",999999]'),text+'[]\n',text.replace('["o","workbook"]','["v","docName","중복"]\n["o","workbook"]')])await assert.rejects(readWixelFile(new Blob([changed])));
});
test('WIXEL v2: 기존 JSON .wixel/.json 통합 문서를 계속 읽는다',async()=>{
  const book=fixture(),data={...meta,workbook:book.serialize()};
  const result=await readWixelFile(new Blob([JSON.stringify(data)]));assert.deepEqual(result,JSON.parse(JSON.stringify(data)));
});
test('게시 압축은 모든 시트 기능을 유지하고 서버 JSON 조각으로 왕복한다',async()=>{
  const book=fixture(),blob=await publishedWixelFile(book,'all',meta),packed=await packPublishedBlob(blob),json=JSON.parse(await packed.text());
  assert.ok(packed.size<=packedPublishedSize(blob.size));
  assert.equal(json.format,'wixel-packed');assert.ok(json.chunks.every(v=>v.length<=262144));
  const back=new Workbook((await readWixelFile(unpackPublishedBlob(json))).workbook);
  assert.equal(JSON.stringify(back.serialize()),JSON.stringify(book.serialize()));
});
test('개별 시트 게시 압축은 기존의 값만 공유하는 정책을 유지한다',async()=>{
  const book=new Workbook({sheets:[{name:'공개',cells:{'0,0':{raw:"='숨김'!A1*2"}}},{name:'숨김',state:'hidden',cells:{'0,0':{raw:'21'}}}]});
  const data=await readWixelFile(await publishedWixelFile(book,'0',meta));const back=new Workbook(data.workbook);
  assert.equal(back.sheets.length,1);assert.equal(back.getValue(0,0,0),42);assert.equal(back.getCell(0,0,0).formula,undefined);
});
test('서버와 서버 없는 링크 크기는 인코딩·업로드 전에 판정하고 내용은 버리지 않는다',async()=>{
  const blob=new Blob(['x'.repeat(1200000)]);assert.throws(()=>assertPublishLinkSize(blob),e=>e.code==='PUBLISH_LINK_TOO_LARGE');
  await assert.rejects(packPublishedBlob(blob,{maxBytes:500000}),e=>e.code==='PUBLISH_TOO_LARGE');
  assert.equal(assertPublishLinkSize(new Blob(['abc'])),4);
  assert.throws(()=>unpackPublishedBlob({format:'wixel-packed',version:1,compression:'gzip',chunks:['bad!']}));
});
test('피벗 배열 JSON 분할 인코딩은 기존 typed-array 표현과 바이트 동일하다',()=>{
  const value={num:new Float64Array([NaN,0,12]),str:new Int32Array([0,-1,1]),dict:['가😀','나']};
  assert.equal([...jsonValueParts(value,true)].join(''),JSON.stringify(value,jsonReplacer));
  const big={num:new Float64Array(100000),dict:['합성']};let max=0,count=0;for(const part of jsonValueParts(big,true)){max=Math.max(max,part.length);count++;}
  assert.ok(count>20);assert.ok(max<=32768);
});

test('WIXEL v2: 압축 기능이 없는 브라우저의 게시 포맷도 왕복한다',async()=>{
  const blob=await writeWixelFile(fixture(),meta,{gzip:false});
  const data=JSON.parse(await(await packPublishedBlob(blob)).text());assert.equal(data.compression,'none');
  const read=await readWixelFile(unpackPublishedBlob(data));assert.equal(read.docName,meta.docName);
});
test('WIXEL v2: 스트림 청크가 한글·레코드 가운데에서 잘려도 읽는다',async()=>{
  const original=await writeWixelFile(fixture(),meta,{gzip:false}),bytes=new Uint8Array(await original.arrayBuffer());
  const wrapped={size:original.size,slice:(a,b)=>original.slice(a,b),stream(){let i=0;return new ReadableStream({pull(c){if(i===bytes.length){c.close();return;}const end=Math.min(bytes.length,i+17);c.enqueue(bytes.subarray(i,end));i=end;}});}};
  const read=await readWixelFile(wrapped);assert.equal(read.docName,meta.docName);
  assert.equal(new Workbook(read.workbook).getCell(0,0,0).raw,'한글😀\ud800');
});
test('WIXEL v2: 메타데이터 __proto__는 프로토타입을 변경하지 않는다',async()=>{
  const header=JSON.parse('{"app":"wixel","__proto__":{"polluted":true}}');
  const data=await readWixelFile(await writeWixelFile(new Workbook(),header));
  assert.equal(Object.getPrototypeOf(data),Object.prototype);assert.equal(Object.hasOwn(data,'__proto__'),true);
  assert.equal({}.polluted,undefined);
});

test('기존 #view= 링크의 gzip JSON 형식도 계속 읽는다',async()=>{
  const original={...meta,workbook:fixture().serialize()};
  const compressed=await new Response(new Blob([JSON.stringify(original)]).stream().pipeThrough(new CompressionStream('gzip'))).blob();
  const read=await readWixelFile(compressed);assert.deepEqual(read,JSON.parse(JSON.stringify(original)));
});

test('WIXEL v2: 배열 수식 캐시는 0·FALSE·빈 문자열·빈칸·오류를 구별한다',async()=>{
  const cachedArray={h:6,w:1,values:[0,0,10,1,0,0,2,0,false,3,0,'',4,0,null,5,0,{error:'#N/A'}]};
  const book=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'2'}}},{name:'배열',fileValues:true,cells:{'0,2':{raw:'=NOSUCHARRAY(원본!A1)',cached:10,cachedArray}}}]});
  const before=[0,1,2,3,4,5].map(r=>book.getValue(1,r,2));
  const back=new Workbook((await readWixelFile(await writeWixelFile(book,meta))).workbook);
  assert.deepEqual([0,1,2,3,4,5].map(r=>back.getValue(1,r,2)),before);
  assert.equal(back.getValue(1,1,2),0);assert.equal(back.getValue(1,2,2),false);assert.equal(back.getValue(1,3,2),'');assert.equal(back.getValue(1,4,2),null);
  book.transact(()=>book.setInput(0,0,0,'1'));
  const dirty=new Workbook((await readWixelFile(await writeWixelFile(book,meta))).workbook);
  assert.equal(dirty.getValue(1,0,2).code,'#NAME?');assert.equal(dirty.getValue(1,1,2),null);
  book.undo();const undone=new Workbook((await readWixelFile(await writeWixelFile(book,meta))).workbook);
  assert.deepEqual([0,1,2,3,4,5].map(r=>undone.getValue(1,r,2)),before);
});

test('WIXEL v2: 날짜 메타데이터는 기존 JSON과 같은 ISO 문자열로 보존한다',async()=>{
  const book=new Workbook({props:{created:new Date('2026-01-02T03:04:05Z')},sheets:[{name:'날짜',cells:{}}]});
  const read=await readWixelFile(await writeWixelFile(book,meta));assert.equal(read.workbook.props.created,'2026-01-02T03:04:05.000Z');
  assert.equal([...jsonValueParts({date:new Date('2026-01-02T03:04:05Z')},true)].join(''),JSON.stringify({date:new Date('2026-01-02T03:04:05Z')},jsonReplacer));
});


test('WIXEL v2: shared style references remove repeated JSON without changing the source',async()=>{
  const book=new Workbook(),style=Object.freeze({font:'맑은 고딕',size:11,bold:true,fill:'#ffeeaa',color:'#223344',border:{bottom:{style:'thin',color:'#778899'}},numFmt:'#,##0.00'});
  for(let r=0;r<5000;r++)book.sheets[0].cells.setRC(r,0,{raw:String(r),style});
  const source=book.sheets[0].cells.getRC(0,0),normal=[...book.cellRunChunks(0,2048)][0][0][3],shared=[...book.cellRunChunks(0,2048,{shareStyle:true})][0][0][3];
  assert.notEqual(normal.style,style);assert.equal(shared.style,style);
  assert.equal(cellData(source,null).style,undefined);assert.notEqual(cellData(source).style,style);
  const text=await(await writeWixelFile(book,meta,{gzip:false})).text(),records=text.trim().split('\n').slice(1).map(line=>JSON.parse(line));
  assert.equal(records.filter(r=>r[0]==='d').length,1);
  assert.equal(text.split('맑은 고딕').length-1,1);assert.ok(text.length<260000);
  assert.equal(source.style,style);assert.deepEqual(shared.style,style);
  const read=await readWixelFile(new Blob([text])),cells=read.workbook.sheets[0].cells;
  assert.equal(cells.getRC(0,0).style,cells.getRC(4999,0).style);
  const back=new Workbook(read.workbook);assert.deepEqual(back.getCell(0,4999,0).style,style);
  back.transact(()=>back.setStyle(0,0,0,{bold:false}));assert.equal(back.getCell(0,4999,0).style.bold,true);
});

test('WIXEL v2: bounded style dictionaries fall back losslessly for unique and oversized styles',async()=>{
  const book=new Workbook();for(let r=0;r<4200;r++)book.sheets[0].cells.setRC(r,0,{raw:String(r),style:{fill:'#'+r.toString(16).padStart(6,'0')}});
  const huge='큰 서식😀'.repeat(16000);book.sheets[0].cells.setRC(4200,0,{raw:'last',style:{font:huge}});
  const blob=await writeWixelFile(book,meta,{gzip:false}),text=await blob.text(),records=text.trim().split('\n').slice(1).map(line=>JSON.parse(line));
  assert.equal(records.filter(r=>r[0]==='d').length,4096);
  assert.ok(records.every(r=>JSON.stringify(r).length<250000));
  const back=new Workbook((await readWixelFile(blob)).workbook);
  for(const r of [0,4095,4096,4199])assert.equal(back.getCell(0,r,0).style.fill,'#'+r.toString(16).padStart(6,'0'));
  assert.equal(back.getCell(0,4200,0).style.font,huge);
});

test('WIXEL v2: existing inline-style records remain readable and invalid style references fail',async()=>{
  const records=[['o',null],['o','workbook'],['v','version',1],['a','sheets'],['o',null],['v','name','Legacy'],['c','cells'],['r',[[0,0,1,{raw:'old',style:{bold:true}}]]],['e'],['e'],['e'],['e'],['e']];
  const encode=rows=>new Blob([WIXEL_FILE_MAGIC+'\n'+rows.map(r=>JSON.stringify(r)+'\n').join('')+JSON.stringify(['finish',rows.length])+'\n']);
  assert.equal(new Workbook((await readWixelFile(encode(records))).workbook).getCell(0,0,0).style.bold,true);
  const missing=structuredClone(records);missing[7]=['r',[[0,0,1,{raw:'bad'},0]]];await assert.rejects(readWixelFile(encode(missing)));
  const duplicate=structuredClone(records);duplicate.splice(7,0,['d',0,{bold:true}],['d',0,{bold:false}]);await assert.rejects(readWixelFile(encode(duplicate)));
});


test('WIXEL stream batches many small metadata records without changing framing or cancellation',async()=>{
  const book=new Workbook();for(let r=0;r<5000;r++){book.sheets[0].rowHeights[r]=20;book.sheets[0].rowStyles[r]={bold:true,color:'#223344'};}
  const before=JSON.stringify(book.serialize()),encode=TextEncoder.prototype.encode;let calls=0,blob;
  try {TextEncoder.prototype.encode=function(...args){calls++;return encode.apply(this,args);};blob=await writeWixelFile(book,meta);}finally{TextEncoder.prototype.encode=encode;}
  assert.ok(calls<2000,'Tiny metadata records must share stream buffers');
  const back=new Workbook((await readWixelFile(blob)).workbook);assert.equal(JSON.stringify(back.serialize()),before);
  let checks=0;await assert.rejects(writeWixelFile(book,meta,{isCurrent:()=>++checks<5}),error=>error.code==='WIXEL_FILE_ABORT');
});

test('reader bounds compressed Blob input independently of native stream chunk sizes',async()=>{
  let state=17,parts=[];for(let i=0;i<180000;i++){state=(Math.imul(state,1664525)+1013904223)>>>0;parts.push(String.fromCharCode(32+state%94));}
  const raw=parts.join('')+'한글😊'.repeat(20000),w=new Workbook({sheets:[{name:'청크',cells:{'0,0':{raw,inputType:'text'}}}]});
  const file=await writeWixelFile(w);assert.ok(file.size>65536);
  let reads=0,maxRead=0;
  class CheckedBlob extends Blob {
    stream(){assert.fail('native Blob.stream chunk sizes must not control decompression');}
    slice(...args){const part=super.slice(...args),read=part.arrayBuffer.bind(part);part.arrayBuffer=()=>{reads++;maxRead=Math.max(maxRead,part.size);assert.ok(part.size<=65536);return read();};return part;}
  }
  const data=await readWixelFile(new CheckedBlob([file]));
  assert.equal(new Workbook(data.workbook).sheets[0].cells.getRC(0,0).raw,raw);
  assert.ok(reads>=3);assert.equal(maxRead,65536);
});


for(const gzip of [false,true])test('WIXEL exact formula memo preserves independent cells and oversized text formulas: '+gzip,async()=>{
  const raw='=IF("한글😀"="한글😀",ROW()+1,0)',huge='="'+'한글😀'.repeat(530000)+'"',cells={};
  for(let r=0;r<300;r++)cells[r+',0']={raw,cached:r+2,comment:'note'+r,style:{fill:r%2?'#223344':'#556677'}};
  cells['0,1']={raw,style:{numFmt:'text'},inputType:'text'};cells['1,1']={raw:huge,inputType:'text'};
  const book=new Workbook({sheets:[{name:'Repeat',fileValues:true,cells},{name:'Next',fileValues:true,cells:{'0,0':{raw,cached:99,comment:'separate'}}}]}),original=book.serialize();
  const file=await writeWixelFile(book,meta,{gzip}),loaded=await readWixelFile(file),first=loaded.workbook.sheets[0].cells.getRC(0,0),second=loaded.workbook.sheets[0].cells.getRC(1,0);
  assert.notEqual(first,second);assert.equal(first.raw,raw);assert.equal(second.raw,raw);assert.equal(first.cached,2);assert.equal(second.cached,3);assert.equal(first.comment,'note0');assert.equal(second.comment,'note1');
  assert.equal(loaded.workbook.sheets[0].cells.getRC(1,1).raw,huge);assert.equal(loaded.workbook.sheets[1].cells.getRC(0,0).cached,99);
  const back=new Workbook(loaded.workbook);assert.deepEqual(back.serialize(),original);assert.equal(back.getValue(0,0,1),raw);assert.equal(back.getValue(0,1,1),huge);
  back.transact(()=>back.setInput(0,0,0,'=123'));assert.equal(back.getRaw(0,1,0),raw);assert.equal(back.getCell(0,1,0).comment,'note1');back.undo();assert.equal(back.getRaw(0,0,0),raw);back.redo();assert.equal(back.getRaw(0,0,0),'=123');
  const again=new Workbook((await readWixelFile(await writeWixelFile(back,meta,{gzip}))).workbook);assert.deepEqual(again.serialize(),back.serialize());assert.deepEqual(book.serialize(),original);
});
