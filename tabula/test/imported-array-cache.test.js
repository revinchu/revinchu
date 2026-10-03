import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { ERR } from '../src/formula.js';
import { arrayCache, arrayCacheValue } from '../src/array-cache.js';

function fixture(formula = 'SEQUENCE(A1,1,10,10)', ref = 'C1:C3', tails = '<row r="2"><c r="C2"><v>20</v></c></row><row r="3"><c r="C3"><v>30</v></c></row>') {
  const w = new Workbook(), files = unzip(writeXlsx(w));
  files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="C1"><f t="array" ref="${ref}">${formula}</f><v>10</v></c></row>${tails}</sheetData></worksheet>`);
  return readXlsx(zip(files)).data;
}
const values = w => [w.getValue(0,0,2), w.getValue(0,1,2), w.getValue(0,2,2)];

test('배열 저장 캐시: 현재 계산과 다른 앵커·follower 결과 보존 및 XLSX 왕복', () => {
  const w = new Workbook(fixture());
  assert.deepEqual(values(w), [10,20,30]);
  assert.equal(w.getRaw(0,1,2), '');
  assert.deepEqual(w.rangeRead(0,0,2,2,2), [[10],[20],[30]]);
  assert.equal(w.getValue(0,0,2),10);
  w.transact(() => w.setInput(0,0,4,'=SUM(C1#)'));
  // 실제 입력이 바뀌면 시트 캐시를 무효화하여 현재 source에 맞춰 계산한다.
  assert.deepEqual(values(w),[10,20,null]);
  w.undo(); assert.deepEqual(values(w),[10,20,30]);
  const bytes=writeXlsx(w), xml=textOf(unzip(bytes)['xl/worksheets/sheet1.xml']);
  assert.match(xml,/<c r="C3"[^>]*><v>30<\/v>/);
  assert.match(xml,/<f t="array" ref="C1:C3"/);
  assert.deepEqual(values(new Workbook(readXlsx(bytes).data)),[10,20,30]);
});

for (const readFirst of [false,true]) test(`원본 편집→무효화→Undo/Redo (첫 조회 ${readFirst})`, () => {
  const w=new Workbook(fixture());
  if(readFirst)assert.deepEqual(values(w),[10,20,30]);
  w.transact(()=>w.setInput(0,0,0,'1'));
  assert.deepEqual(values(w),[10,null,null]);
  w.undo(); assert.deepEqual(values(w),[10,20,30]);
  w.redo(); assert.deepEqual(values(w),[10,null,null]);
  const back=new Workbook(w.serialize());
  assert.deepEqual(values(back),[10,null,null]);
});

test('서식 변경은 배열 캐시와 모든 follower를 유지하고 지우기는 캐시를 제거', () => {
  const w=new Workbook(fixture());
  w.transact(()=>w.setStyle(0,0,2,{bold:true}));
  assert.deepEqual(values(w),[10,20,30]);
  w.transact(()=>w.clearRange(0,0,2,0,2,'contents'));
  assert.deepEqual(values(w),[null,null,null]);
  w.undo(); assert.deepEqual(values(w),[10,20,30]);
});

test('지원하지 않는 배열 함수도 저장 캐시 사용·의존 데이터 변경 시 안전 무효화·Undo', () => {
  const w=new Workbook(fixture('NOSUCHARRAY(A1)'));
  assert.deepEqual(values(w),[10,20,30]);
  w.transact(()=>w.setInput(0,0,0,'1'));
  assert.deepEqual(values(w),[ERR.NAME,null,null]);
  w.undo();assert.deepEqual(values(w),[10,20,30]);
});

test('리터럴 장애물은 원본 배열 캐시를 무효화하고 SPILL을 발생시킨다', () => {
  const w=new Workbook(fixture());
  w.transact(()=>w.setInput(0,1,2,'blocked'));
  assert.deepEqual(values(w),[ERR.SPILL,'blocked',null]);
  w.undo();assert.deepEqual(values(w),[10,20,30]);
});

test('JSON·Blob·셀 저장 레코드에서 배열 캐시와 stale 상태 유지', async () => {
  const w=new Workbook(fixture());
  for(const back of [new Workbook(JSON.parse(JSON.stringify(w.serialize()))),new Workbook(JSON.parse(await w.serializeBlob().text()))])assert.deepEqual(values(back),[10,20,30]);
  assert.equal(cellData(w.getCell(0,0,2)).cachedArray.values.length,9);
  w.transact(()=>w.setInput(0,0,0,'1'));
  assert.ok(cellData(w.getCell(0,0,2)).staleCachedArray);
  assert.deepEqual(values(new Workbook(JSON.parse(await w.serializeBlob().text()))),[10,null,null]);
});

test('큰 희소 배열은 실제 저장된 좌표만 만들고 범위 전체를 펼치지 않는다', () => {
  const w=new Workbook(fixture('NOSUCHARRAY(A1)','C1:XFD1048576','<row r="1048576"><c r="XFD1048576"><v>99</v></c></row>'));
  assert.equal(w.getValue(0,1048575,16383),99);
  assert.equal(w.spillOwner.size,1);
  assert.deepEqual(w.spillAnchorOf(0,1000,1000),{r:0,c:2});
  const bytes=writeXlsx(w),xml=textOf(unzip(bytes)['xl/worksheets/sheet1.xml']);
  assert.ok(xml.length<10000);assert.match(xml,/r="XFD1048576"[^>]*><v>99<\/v>/);
});

test('캐시 검증은 범위 밖·중복·잘못된 자료형을 거부한다', () => {
  assert.equal(arrayCache({h:2,w:1,values:[2,0,3]}),null);
  assert.equal(arrayCache({h:2,w:1,values:[0,0,3,0,0,4]}),null);
  assert.equal(arrayCache({h:2,w:1,values:[0,0,Infinity]}),null);
  const a=arrayCache({h:2,w:1,values:[0,0,false,1,0,'']});
  assert.equal(arrayCacheValue(a,0,0),false);assert.equal(arrayCacheValue(a,1,0),'');
});

test('다른 시트 원본 변경도 첫 조회 전부터 무효화하고 Undo/Redo로 복원', () => {
  const data=fixture('SEQUENCE(Input!A1,1,10,10)');
  data.sheets.push({name:'Input',cells:{'0,0':{raw:'2'}}});
  const w=new Workbook(data);
  w.transact(()=>w.setInput(1,0,0,'1'));
  assert.deepEqual(values(w),[10,null,null]);
  w.undo();assert.deepEqual(values(w),[10,20,30]);
  w.redo();assert.deepEqual(values(w),[10,null,null]);
});

test('오류·0·FALSE·빈문자열·명시적인 빈 캐시를 구분해 왕복', () => {
  const tails='<row r="2"><c r="C2" t="e"><v>#N/A</v></c></row><row r="3"><c r="C3"><v>0</v></c></row><row r="4"><c r="C4" t="b"><v>0</v></c></row><row r="5"><c r="C5" t="str"><v></v></c></row><row r="6"><c r="C6"><v/></c></row>';
  const w=new Workbook(fixture('NOSUCHARRAY(A1)','C1:C7',tails));
  const expected=[10,ERR.NA,0,false,'',null,null];
  assert.deepEqual(Array.from({length:7},(_,r)=>w.getValue(0,r,2)),expected);
  const bytes=writeXlsx(w),xml=textOf(unzip(bytes)['xl/worksheets/sheet1.xml']);
  assert.match(xml,/<c r="C6"[^>]*><v\/><\/c>/);assert.doesNotMatch(xml,/<c r="C7"/);
  const back=new Workbook(readXlsx(bytes).data);
  assert.deepEqual(Array.from({length:7},(_,r)=>back.getValue(0,r,2)),expected);
  assert.deepEqual(back.getCell(0,0,2).cachedArray.values,w.getCell(0,0,2).cachedArray.values);
});

test('저장 결과 없는 array ref만으로 계산을 차단하지 않는다', () => {
  const data=fixture();
  delete data.sheets[0].cells.getRC(0,2).cached;
  data.sheets[0].cells.getRC(0,2).cachedArray={h:3,w:1,values:[]};
  const w=new Workbook(data);assert.deepEqual(values(w),[10,20,null]);
});

test('행 구조 변경은 오래된 위치의 결과를 제거하고 Undo는 원래 캐시 복원', () => {
  const w=new Workbook(fixture());
  w.transact(()=>w.insertRows(0,0,1));
  assert.equal(w.getRaw(0,1,2),'=SEQUENCE(A2,1,10,10)');
  assert.equal(w.getValue(0,0,2),null);
  assert.equal(w.getValue(0,3,2),null);
  w.undo(); assert.deepEqual(values(w),[10,20,30]);
  w.redo();assert.equal(w.getValue(0,3,2),null);
});

test('복사한 배열 수식은 목적지의 참조로 다시 계산하며 옛 캐시를 복제하지 않는다', () => {
  const w=new Workbook(fixture());
  const data=cellData(w.getCell(0,0,2));
  w.transact(()=>w.setCellData(0,5,2,{...data,raw:'=SEQUENCE(1,1,99)'}));
  assert.equal(w.getValue(0,5,2),99);assert.equal(w.getValue(0,6,2),null);
  w.undo();assert.deepEqual(values(w),[10,20,30]);
});

for (const styleFirst of [true,false]) test(`서식·원본 편집을 함께 Redo해도 오래된 배열 캐시를 살리지 않음 (${styleFirst})`, () => {
 const w=new Workbook(fixture());
 w.transact(()=>{if(styleFirst)w.setStyle(0,0,2,{bold:true});w.setInput(0,0,0,'1');if(!styleFirst)w.setStyle(0,0,2,{bold:true});});
 assert.deepEqual(values(w),[10,null,null]);w.undo();assert.deepEqual(values(w),[10,20,30]);w.redo();assert.deepEqual(values(w),[10,null,null]);
});

test('배열 참조 및 합계는 첫 불러온 저장 결과를 사용한다', () => {
 const data=fixture();data.sheets[0].cells.setRC(0,4,{raw:'=SUM(C1#)'});
 const w=new Workbook(data);assert.equal(w.getValue(0,0,4),60);
});
