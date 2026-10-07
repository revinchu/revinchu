import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { unzip, zip } from '../src/zip.js';
import { slicerStartItem, slicerStartScrollTop, slicerWindow } from '../src/slicer-window.js';
import { applyCachedSlicerSelection } from '../src/slicer-cache.js';

const slicer = (startItem = 7) => ({ id:'synthetic-start',caption:'항목',x:20,y:20,w:220,h:120,columns:2,buttonHeight:24,startItem,
  source:{kind:'cache',field:'항목',cacheKey:'synthetic-start-cache',values:Array.from({length:20},(_,i)=>'항목 '+String(i).padStart(2,'0')),items:Array.from({length:20},(_,index)=>({index,hasData:index<18})),format:{}},cacheSelection:['7','11'],style:'SlicerStyleLight1' });
const book = startItem => { const wb=new Workbook();wb.sheets[0].slicers=[slicer(startItem)];return wb; };

test('startItem의 unsignedInt 기본값·무효값과 표시 항목 범위를 안전하게 처리한다',()=>{
 for(const value of [undefined,null,'',NaN,Infinity,-1,1.5,4294967296,'invalid','1e3','0x10','11.0',true,false,{}])assert.equal(slicerStartItem(value),0);
 for(const value of [0,1,9,4294967295,'11'])assert.equal(slicerStartItem(value),Number(value));
 assert.equal(slicerStartItem(11,5),4);assert.equal(slicerStartItem(11,0),0);
});

test('단일열·다중열의 초기 스크롤은 항목의 행을 사용하고 source 인덱스를 바꾸지 않는다',()=>{
 assert.equal(slicerStartScrollTop(20,1,24,3,100,7),189);
 assert.equal(slicerStartScrollTop(20,2,24,3,100,7),81);
 assert.equal(slicerStartScrollTop(20,2,24,3,100,6),81);
 assert.equal(slicerStartScrollTop(0,2,24,3,100,7),0);
 assert.equal(slicerStartScrollTop(20,2,24,3,100,-1),0);
 const end=slicerWindow(20,2,24,3,100,Infinity);
 assert.equal(slicerStartScrollTop(20,2,24,3,100,999),end.max);
});

test('8M 픽셀보다 큰 가상 슬라이서의 초기 항목은 논리 스크롤 좌표로 변환한다',()=>{
 const count=1000000,columns=2,start=900001;
 const top=slicerStartScrollTop(count,columns,24,3,240,start),win=slicerWindow(count,columns,24,3,240,top);
 assert.ok(win.physical<=8000000);assert.ok(Math.abs(win.logical-Math.floor(start/columns)*27)<1e-6);
 assert.ok(win.first<=start&&win.last>start);assert.ok(win.last-win.first<50);
});

test('XLSX startItem은 일반·stream bytes·stream Blob 가져오기 및 왕복에 보존된다',async()=>{
 const wb=book(7),before=structuredClone(wb.sheets[0].slicers[0]),bytes=writeXlsx(wb);
 assert.deepEqual(wb.sheets[0].slicers[0],before,'내보내기는 기존 source/selection/위치를 바꾸지 않는다');
 const parts=unzip(bytes);assert.match(new TextDecoder().decode(parts['xl/slicers/slicer1.xml']),/ startItem="7"/);
 for(const result of [readXlsx(bytes),await readXlsxAsync(bytes,undefined,{streamThreshold:0}),await readXlsxAsync(new Blob([bytes]),undefined,{streamThreshold:0})]){
  const sl=result.data.sheets[0].slicers[0];assert.equal(sl.startItem,7);assert.equal(sl.columns,2);assert.deepEqual(sl.source.values,before.source.values);assert.deepEqual(sl.source.items,before.source.items);assert.deepEqual(sl.cacheSelection,before.cacheSelection);
  const again=readXlsx(writeXlsx(new Workbook(result.data))).data.sheets[0].slicers[0];assert.equal(again.startItem,7);assert.deepEqual(again.source.values,before.source.values);assert.deepEqual(again.cacheSelection,before.cacheSelection);
 }
});

test('XLSX 생략·무효 startItem은 0이며 유효한 큰 값은 모델에 보존하고 표시 시만 제한한다',()=>{
 const parts=unzip(writeXlsx(book(7))),path='xl/slicers/slicer1.xml',base=new TextDecoder().decode(parts[path]);
 for(const [attribute,expected]of [[undefined,0],['0',0],['-1',0],['1.5',0],['invalid',0],['4294967296',0],['4294967295',4294967295]]){
  const updated=base.replace(/ startItem="[^"]*"/,attribute===undefined?'':` startItem="${attribute}"`);
  const sl=readXlsx(zip({...parts,[path]:new TextEncoder().encode(updated)})).data.sheets[0].slicers[0];assert.equal(sl.startItem??0,expected);assert.equal(slicerStartItem(sl.startItem,sl.source.items.length),expected?19:0);
 }
 const xml=new TextDecoder().decode(unzip(writeXlsx(book(0)))[path]);assert.doesNotMatch(xml,/ startItem=/);
});

test('필터 선택과 Undo/Redo는 저장된 최초 표시 항목·원본 값·위치를 변경하지 않는다',()=>{
 const wb=book(7),sl=wb.sheets[0].slicers[0],source=structuredClone(sl.source);
 applyCachedSlicerSelection(wb,sl.source,['3']);assert.equal(wb.sheets[0].slicers[0].startItem,7);assert.deepEqual(wb.sheets[0].slicers[0].source,source);
 wb.undo();assert.equal(wb.sheets[0].slicers[0].startItem,7);assert.deepEqual(wb.sheets[0].slicers[0].cacheSelection,['7','11']);
 wb.redo();assert.equal(wb.sheets[0].slicers[0].startItem,7);assert.deepEqual(wb.sheets[0].slicers[0].source,source);assert.deepEqual([wb.sheets[0].slicers[0].x,wb.sheets[0].slicers[0].y],[20,20]);
});
