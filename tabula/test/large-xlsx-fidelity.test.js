import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { dirname, resolve, join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';
import { createMemoryMetrics, canonicalHash, digestWorkbook, compareDigests } from '../tools/large-xlsx-fidelity.mjs';

const fixture=(cells=[],extra={})=>new Workbook({sheets:[{name:'검수표',cells:new Map(cells.map(([r,c,data])=>[r+','+c,data])),...extra}]});
const info=async book=>digestWorkbook(book,{maxSamples:0});

test('canonical 값은 키 순서와 긴 UTF-16 문자열 chunk 경계에 의존하지 않는다',()=>{
  assert.equal(canonicalHash({b:[true,null,3],a:'검수'}),canonicalHash({a:'검수',b:[true,null,3]}));
  const value='a'.repeat(32767)+'😀'+'한글'.repeat(18000);
  const expected=createHash('sha256').update('s'+value.length+':').update(value).update(';').digest('hex');
  assert.equal(canonicalHash(value),expected);
  assert.notEqual(canonicalHash('1'),canonicalHash(1));
  assert.throws(()=>{const a={};a.a=a;canonicalHash(a);},/순환/);
});

test('모든 block 값·논리 순서와 일반 셀 override를 한 번만 검사한다',async()=>{
  const fmt={numFmt:'number',color:'#123456',bt:{color:'#abcdef',style:'thin'}};
  const block={r0:7,c0:3,n:3,perm:Int32Array.from([2,0,1]),cols:[{num:Float64Array.from([11,22,33]),str:null,dict:[],fmt}]};
  const a=fixture([[8,3,{raw:'99',style:fmt}]],{blocks:[block]});
  const b=fixture([[7,3,{raw:'33',style:fmt}],[8,3,{raw:'99',style:fmt}],[9,3,{raw:'22',style:fmt}]]);
  const ad=await info(a),bd=await info(b);
  assert.equal(ad.sheets[0].counts.logicalCells,3);
  assert.equal(ad.sheets[0].counts.blockValues,2);
  assert.equal(compareDigests(ad,bd).equal,true);
  a.sheets[0].blocks[0].cols[0].num[1]=23;
  assert.equal(compareDigests(ad,await info(a)).equal,false);
});

test('수백만 서식 전용 빈 셀은 구간으로 검사하고 행별 상속 변경을 반영한다',async()=>{
  const book=fixture([],{rowStyles:{5:{bold:true}}}),style={fill:'#eeddaa'};
  book.sheets[0].cells.setRunRC(2,1,1000000,{raw:'',v:null,style});
  let calls=0;const styleAt=book.styleAt.bind(book);
  book.styleAt=(...args)=>{calls++;return styleAt(...args);};
  const result=await info(book);
  assert.equal(result.sheets[0].counts.logicalCells,1000000);
  assert.equal(result.sheets[0].counts.blankCells,1000000);
  assert.ok(calls<=4,'논리 빈 행 전체를 펼치지 않아야 한다');
  assert.equal(result.sheets[0].counts.canonicalRuns,4);
});

test('sparse 좌표 양방향과 추가 셀을 검출하며 used rectangle을 훑지 않는다',async()=>{
  const a=fixture([[0,0,{raw:'처음'}],[1048575,16383,{raw:'마지막'}]]);
  let calls=0;const styleAt=a.styleAt.bind(a);a.styleAt=(...args)=>{calls++;return styleAt(...args);};
  const before=await info(a);assert.equal(calls,2);assert.equal(before.sheets[0].counts.logicalCells,2);
  a.sheets[0].cells.setRC(900001,900,{raw:'추가',v:'추가'});
  const diff=compareDigests(before,await info(a));
  assert.ok(diff.differences.some(x=>x.category==='contents'));
  assert.ok(diff.differences.some(x=>x.category==='count.logicalCells'));
});

test('압축 구간과 삽입 순서가 다른 point 표현도 같은 logical hash다',async()=>{
  const a=fixture(),b=fixture(),style={color:'#abcdef',bb:{color:'#ffeedd',style:'double'}};
  a.sheets[0].cells.setRunRC(0,3,140,{raw:'',v:null,style});
  for(let r=139;r>=0;r--)b.sheets[0].cells.setRC(r,3,{raw:'',v:null,style});
  assert.equal(compareDigests(await info(a),await info(b)).equal,true);
});

test('일반 셀이 겹친 블록을 우선하며 empty 첫 블록도 이후 block을 가린다',async()=>{
  const first={r0:2,c0:0,n:3,cols:[{num:null,str:null,dict:[],fmt:null}]};
  const second={r0:2,c0:0,n:3,cols:[{num:Float64Array.from([10,20,30]),str:null,dict:[],fmt:{bold:true}}]};
  const a=fixture([[3,0,{raw:'100'}]],{blocks:[first,second]}),b=fixture([[3,0,{raw:'100'}]]);
  assert.equal(compareDigests(await info(a),await info(b)).equal,true);
});

test('수식·cached error·주석·링크는 재계산 없이 검사한다',async()=>{
  const a=fixture([[2,1,{raw:'=1+1',cached:2}],[4,1,{raw:'#N/A',comment:{text:'원문'},link:'https://example.test/'}]]);
  a.getValue=()=>{throw new Error('재계산 금지');};
  const before=await info(a);assert.equal(before.sheets[0].counts.formulas,1);
  a.sheets[0].cells.getRC(2,1).cached=3;
  assert.ok(compareDigests(before,await info(a)).differences.some(x=>x.category==='contents'));
});

test('행열 크기·병합·모든 drawing 속성을 해시하고 무작위 ID만 정규화한다',async()=>{
  const extra=id=>({rowHeights:{2:31},colWidths:{1:93},merges:[{r1:1,c1:1,r2:2,c2:3}],shapes:[{id,name:'상자',x:31,y:21,w:40,h:60,anchor:{kind:'twoCell',from:{col:1,row:1}},line:{color:'#123456',width:2},children:[{id:id+'child',name:'자식',fill:'#aabbcc'}]}],slicers:[{id:id+'sl',caption:'캠페인',source:{field:'원본 필드'},x:100,y:10,w:150,h:120}]});
  const a=fixture([],extra('sh-random-a')),b=fixture([],extra('sh-random-b'));
  const before=await info(a);assert.equal(compareDigests(before,await info(b)).equal,true);
  b.sheets[0].shapes[0].line.color='#654321';
  assert.ok(compareDigests(before,await info(b)).differences.some(x=>x.category==='drawings'));
  b.sheets[0].rowHeights[2]=32;
  assert.ok(compareDigests(before,await info(b)).differences.some(x=>x.category==='metadata'));
  b.sheets[0].slicers[0].source.field='다른 필드';
  assert.notEqual((await info(b)).sheets[0].hashes.drawings,before.sheets[0].hashes.drawings);
});

test('서식의 테두리·색·숫자형식·폰트 차이를 구별한다',async()=>{
  const a=fixture([[1,1,{raw:'123',style:{font:'맑은 고딕',size:11,color:'#112233',fill:'#ffffff',code:'0.00',numFmt:'custom',bl:{style:'thin',color:'#aabbcc'}}}]]);
  const before=await info(a),previous=a.sheets[0].cells.getRC(1,1);a.sheets[0].cells.setRC(1,1,{...previous,style:{...previous.style,bl:{style:'double',color:'#aabbcc'}}});
  const diff=compareDigests(before,await info(a));
  assert.ok(diff.differences.some(x=>x.category==='styles'));
  assert.ok(!diff.differences.some(x=>x.category==='contents'));
});

test('CLI는 별도 순차 process로 작은 합성 파일을 왕복하고 원본 hash를 보존한다',async()=>{
  const root=resolve(dirname(fileURLToPath(import.meta.url)),'../.local');await mkdir(root,{recursive:true});
  const folder=await mkdtemp(join(root,'fidelity-test-'));
  try {
    const input=join(folder,'input.xlsx'),output=join(folder,'saved.xlsx'),reportDir=join(folder,'report'),nativeTargets=join(folder,'targets.json');
    await writeFile(nativeTargets,JSON.stringify({sheets:[{index:0,name:'검수표',cells:['A1','B2'],rows:[1,2],columns:['A','B']}]}));
    await writeFile(input,writeXlsx(fixture([[0,0,{raw:'값'}],[0,1,{raw:'123'}],[1,1,{raw:'=B1+1',cached:124}]])));
    const before=createHash('sha256').update(await readFile(input)).digest('hex');
    const child=spawn(process.execPath,[resolve(dirname(fileURLToPath(import.meta.url)),'../tools/large-xlsx-fidelity.mjs'),input,'--out',reportDir,'--roundtrip',output,'--native-targets',nativeTargets,'--source',resolve(dirname(fileURLToPath(import.meta.url)),'../src')],{stdio:['ignore','pipe','pipe']});
    let logs='';child.stdout.on('data',chunk=>{logs+=chunk;});child.stderr.on('data',chunk=>{logs+=chunk;});
    const [code]=await once(child,'exit');const cliReport=JSON.parse(await readFile(join(reportDir,'report.json'),'utf8'));assert.equal(code,0,logs+' '+JSON.stringify(cliReport.comparison));
    const report=JSON.parse(await readFile(join(reportDir,'report.json'),'utf8'));
    assert.equal(report.pass,true);assert.equal(report.sameSourceBetweenPhases,true);assert.equal(report.sameNativeTargets,true);
    assert.equal(report.original.nativeModelSamples.cells,2);assert.equal(report.reimport.nativeModelSamples.cells,2);
    const samples=JSON.parse(await readFile(join(reportDir,'original','native-model-samples.json'),'utf8'));assert.equal(samples.sheets[0].cells.length,2);assert.ok(!JSON.stringify(samples).includes('cached'));
    assert.equal(report.original.inputBefore.sha256,before);assert.equal(report.original.inputUnchanged,true);
    assert.equal(report.reimport.inputUnchanged,true);assert.equal(report.original.digest.scope.visualRender,false);
    assert.equal(report.original.diagnosticBandRows,4096);
    for(const phase of [report.original,report.reimport]) {
      assert.ok(phase.importMs>=0);assert.ok(phase.digestMs>=0);assert.equal(phase.modelGc.performed,true);
      assert.ok(phase.memory.modelAfterGc.heapUsed>0);assert.ok(phase.memory.import.samples>=2);assert.ok(phase.memory.import.highWater.rss.bytes>0);
    }
    assert.ok(report.original.exportMs>=0);assert.ok(report.original.memory.afterExport.rss>0);assert.equal(report.reimport.exportMs,undefined);
    for(const line of (await readFile(join(reportDir,'original','contents.ndjson'),'utf8')).trim().split('\n'))assert.ok(JSON.parse(line).hash);
    const after=createHash('sha256').update(await readFile(input)).digest('hex');assert.equal(after,before);
    const refused=spawn(process.execPath,[resolve(dirname(fileURLToPath(import.meta.url)),'../tools/large-xlsx-fidelity.mjs'),input,'--out',reportDir,'--roundtrip',input],{stdio:'ignore'});
    const [refusedCode]=await once(refused,'exit');assert.equal(refusedCode,2);
  } finally {
    assert.ok(resolve(folder).startsWith(root+sep));await rm(folder,{recursive:true,force:true});
  }
});
test('CLI 실패도 검수 미완료 보고를 남기고 기존 결과 폴더를 덮어쓰지 않는다',async()=>{
  const root=resolve(dirname(fileURLToPath(import.meta.url)),'../.local');await mkdir(root,{recursive:true});
  const folder=await mkdtemp(join(root,'fidelity-test-'));
  try {
    const input=join(folder,'broken.xlsx'),out=join(folder,'report'),tool=resolve(dirname(fileURLToPath(import.meta.url)),'../tools/large-xlsx-fidelity.mjs');
    await writeFile(input,Uint8Array.of(1,2,3,4));
    const child=spawn(process.execPath,[tool,input,'--out',out],{stdio:'ignore'});
    const [code]=await once(child,'exit');assert.equal(code,2);
    const reportText=await readFile(join(out,'report.json'),'utf8'),report=JSON.parse(reportText);
    assert.equal(report.pass,false);assert.equal(report.original.completed,false);assert.ok(report.error.message);
    const repeat=spawn(process.execPath,[tool,input,'--out',out],{stdio:'ignore'});
    const [repeatCode]=await once(repeat,'exit');assert.equal(repeatCode,2);
    assert.equal(await readFile(join(out,'report.json'),'utf8'),reportText);
  } finally {assert.ok(resolve(folder).startsWith(root+sep));await rm(folder,{recursive:true,force:true});}
});

test('피벗 캐시의 legacy rows·typed prefix·uniform tail은 같은 논리 값이며 조회 상태를 변경하지 않는다',async()=>{
  const header=['구분','값'],rows=[['',0],[null,true],[{error:'#N/A'},-2],['끝',false],...Array.from({length:300},()=>['',0])];
  const builder=new PivotSnapshotBuilder(header);for(const row of rows)builder.add(row);const packed=builder.finish();
  assert.ok(packed.columns.every(c=>c.tail));
  const old=fixture(),fresh=fixture();
  const oldSnap={rows:[header,...rows],ver:7,sourceWatch:{version:7}},freshSnap={rows:packed,ver:7,sourceWatch:{version:7}};
  old.pivotSnapshots=new Map([['old/path',{...oldSnap}]]);fresh.pivotSnapshots=new Map([['new/path',freshSnap]]);
  old.sheets[0].pivot={name:'검수피벗',snapshotId:'old/path',cacheItemsId:'old/path'};fresh.sheets[0].pivot={name:'검수피벗',snapshotId:'new/path',cacheItemsId:'new/path'};
  const retained={missingItemsLimit:1000,fields:[{name:'구분',shared:['',null,'과거 항목'],sharedTypes:'sms',format:{code:'0.00'}}]};
  old.pivotCacheItems={'old/path':retained};fresh.pivotCacheItems={'new/path':structuredClone(retained)};
  const before=await info(old),after=await info(fresh);
  assert.equal(compareDigests(before,after).equal,true);assert.equal(before.caches.snapshots.logicalValues,rows.length*header.length);
  assert.equal(fresh.pivotSnapshots.get('new/path'),freshSnap);assert.equal(freshSnap.ver,7);assert.equal(freshSnap.sourceWatch.version,7);
  freshSnap.rows.columns[0].tail.value=null;
  assert.ok(compareDigests(before,await info(fresh)).differences.some(d=>d.category==='caches.snapshots'));
});

test('피벗 캐시 누락·추가·행수 및 과거 항목/서식 손실을 양방향으로 검출한다',async()=>{
  const book=fixture(),snap={kind:'pivot-cache',version:1,header:['값'],n:1000000,columns:[{num:null,str:null,dict:[],tail:{start:0,value:''}}]};
  book.pivotSnapshots=new Map([['cache',{rows:snap}]]);book.pivotCacheItems={cache:{fields:[{name:'값',shared:['과거'],format:{numFmt:'custom',code:'0.0'}}]}};
  const before=await info(book);assert.equal(before.caches.snapshots.runs,1);assert.equal(before.caches.snapshots.logicalValues,1000000);
  snap.n--;assert.ok(compareDigests(before,await info(book)).differences.some(d=>d.category==='caches.snapshots'));snap.n++;
  book.pivotCacheItems.cache.fields[0].format.code='0.00';assert.ok(compareDigests(before,await info(book)).differences.some(d=>d.category==='caches.items'));
  delete book.pivotCacheItems.cache;assert.ok(compareDigests(before,await info(book)).differences.some(d=>d.category==='caches.items'));
  book.pivotSnapshots.delete('cache');const empty=await info(book);assert.ok(compareDigests(before,empty).differences.some(d=>d.category==='caches.snapshots'));assert.equal(compareDigests(empty,before).equal,false);
  assert.notEqual(canonicalHash(-0),canonicalHash(0));
});


test('메모리 관측값은 지표별 최고 단계와 시점을 보존하고 보고 사본을 분리한다',()=>{
  let t=0,i=0;const usage=[{rss:100,heapUsed:80,arrayBuffers:60},{rss:90,heapUsed:95,arrayBuffers:40},{rss:120,heapUsed:70,arrayBuffers:65}];
  const metrics=createMemoryMetrics({now:()=>t,readMemory:()=>usage[i++]});
  t=1;metrics.sample('parse');t=2;metrics.sample('hydrate');const first=metrics.report();t=3;metrics.sample('complete');
  assert.deepEqual(first.highWater.rss,{bytes:100,stage:'parse',elapsedMs:1});assert.equal(first.highWater.heapUsed.stage,'hydrate');
  assert.equal(metrics.report().highWater.rss.stage,'complete');assert.equal(metrics.report().highWater.arrayBuffers.bytes,65);assert.equal(first.samples,2);
});

test('메모리 계측은 동일 Workbook의 논리 hash와 값을 변경하지 않는다',async()=>{
  const book=fixture([[1,2,{raw:'값',style:{bold:true}}]]),before=await info(book);let ticks=0;
  const metrics=createMemoryMetrics({now:()=>ticks++,readMemory:()=>({rss:10,heapUsed:5,arrayBuffers:2})});metrics.sample('import-complete');
  const after=await digestWorkbook(book,{onProgress:data=>metrics.sample(data.stage)});metrics.sample('digest-complete');
  assert.equal(compareDigests(before,after).equal,true);assert.equal(book.sheets[0].cells.getRC(1,2).raw,'값');assert.equal(metrics.report().samples,2);
});


const groupFixture = prefix => fixture([], { shapes: [{
  id: prefix + '-parent', kind: 'group', x: 10, y: 20, w: 240, h: 140,
  groupItems: [{ id: prefix + '-nested', kind: 'group', x: 2, y: 3, w: 80, h: 50,
    groupId: prefix + '-parent', groupItems: [
      { id: prefix + '-leaf', kind: 'rect', x: 5, y: 6, w: 30, h: 20, text: '합성 자식', ownerId: prefix + '-nested' },
      { id: prefix + '-other', kind: 'rect', x: 45, y: 6, w: 30, h: 20, text: '합성 비교', ownerId: prefix + '-nested' },
    ],
  }, { id: prefix + '-connector', kind: 'line', x: 0, y: 0, w: 70, h: 40,
    start: { shapeId: prefix + '-parent' }, end: { targetId: prefix + '-leaf' },
  }],
}] });

test('groupItems의 무작위 부모·중첩 자식 ID만 다른 그룹은 동일한 도면 hash다', async () => {
  const a = groupFixture('random-one'), b = groupFixture('random-two');
  const beforeA = structuredClone(a.sheets[0].shapes), beforeB = structuredClone(b.sheets[0].shapes);
  const ad = await info(a), bd = await info(b);
  assert.equal(compareDigests(ad, bd).equal, true);
  assert.equal(ad.sheets[0].hashes.drawings, bd.sheets[0].hashes.drawings);
  assert.deepEqual(a.sheets[0].shapes, beforeA, '정규화는 실제 모델 ID를 바꾸지 않는다');
  assert.deepEqual(b.sheets[0].shapes, beforeB);
});

test('그룹 내부 자식의 실제 위치·텍스트 변화는 ID 정규화 후에도 검출한다', async () => {
  const a = groupFixture('same-a'), before = await info(a);
  for (const [key, value] of [['x', 6], ['text', '합성 다른 내용']]) {
    const b = groupFixture('same-b'); b.sheets[0].shapes[0].groupItems[0].groupItems[0][key] = value;
    const diff = compareDigests(before, await info(b));
    assert.ok(diff.differences.some(d => d.category === 'drawings'), key + '는 실제 변경이다');
  }
});

test('그룹 부모·자식 참조는 같은 대상을 정규화하고 다른 자식으로 연결하면 차이를 유지한다', async () => {
  const a = groupFixture('before'), b = groupFixture('after'), ad = await info(a);
  assert.equal(compareDigests(ad, await info(b)).equal, true);
  b.sheets[0].shapes[0].groupItems[1].end.targetId = 'after-other';
  assert.ok(compareDigests(ad, await info(b)).differences.some(d => d.category === 'drawings'));
  const c = groupFixture('changed-owner');
  c.sheets[0].shapes[0].groupItems[0].groupItems[0].ownerId = 'changed-owner-parent';
  assert.ok(compareDigests(ad, await info(c)).differences.some(d => d.category === 'drawings'));
});


const importedCacheReferenceFixture = prefix => {
  const book=fixture(), range={r1:0,c1:0,r2:2,c2:1};
  const definition=(name,id) => ({name,source:'SyntheticExternal',range:{...range},snapshotId:id,cacheItemsId:id,
    sourceReference:{kind:'imported-pivot-cache',ref:'A1:B3',sheet:'SyntheticExternal',name:null,
      external:{type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath',target:'synthetic-origin.xlsx',mode:'External'},
      expectedCacheItemsId:id,binding:{source:'syntheticexternal',table:null,range:{...range}}}});
  const ids=[prefix+'/first.xml',prefix+'/second.xml'];
  book.sheets[0].pivot=definition('First',ids[0]); book.sheets[0].pivotsExtra=[definition('Second',ids[1])];
  book.pivotSnapshots=new Map(ids.map((id,i)=>[id,{rows:[['Value'],[i+1]]}]));
  book.pivotCacheItems=Object.fromEntries(ids.map((id,i)=>[id,{fields:[{name:'Value',shared:[i+1]}]}]));
  return book;
};

test('sourceReference expected cache path is normalized with its existing snapshot/cache relationship', async () => {
  const a=importedCacheReferenceFixture('original-path'), b=importedCacheReferenceFixture('saved-path');
  const original=structuredClone(b.sheets[0].pivot.sourceReference);
  assert.equal(compareDigests(await info(a),await info(b)).equal,true);
  assert.deepEqual(b.sheets[0].pivot.sourceReference,original,'감사는 실제 참조를 변경하지 않는다');
});

test('sourceReference different cache targets and meaningful source/binding changes remain differences', async () => {
  const ad=await info(importedCacheReferenceFixture('before'));
  const wrong=importedCacheReferenceFixture('after');
  wrong.sheets[0].pivot.sourceReference.expectedCacheItemsId=wrong.sheets[0].pivotsExtra[0].cacheItemsId;
  assert.ok(compareDigests(ad,await info(wrong)).differences.some(d=>d.category==='metadata'));
  for(const mutate of [r=>r.ref='A1:C3',r=>r.sheet='OtherSynthetic',r=>r.external.target='other-origin.xlsx',r=>r.binding.source='othersynthetic',r=>r.binding.range.r2=3]) {
    const b=importedCacheReferenceFixture('after'); mutate(b.sheets[0].pivot.sourceReference);
    assert.ok(compareDigests(ad,await info(b)).differences.some(d=>d.category==='metadata'));
  }
});
