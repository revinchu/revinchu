import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
function fixture() {
  return new Workbook({sheets:[
    {name:'Source',cells:{'0,0':{raw:'2'}}},
    {name:'Array',cells:{'0,2':{raw:'=SEQUENCE(Source!A1,1,10,10)',cached:10,cachedArray:{h:3,w:1,values:[0,0,10,1,0,20,2,0,30]}}}},
    {name:'Scalar',cells:{'0,0':{raw:'=NOSUCHFUNCTION(Source!A1)',cached:99}}},
    {name:'Unrelated',cells:{'0,0':{raw:'42'}}},
  ]});
}
const generation = w => w.sheets.map(s=>s._ev??0);
const arrayValues = w => [0,1,2].map(r=>w.getValue(1,r,2));
test('의존 수식의 배열·스칼라 캐시 신뢰 변경은 저장 세대를 올리고 무관한 시트는 재사용', () => {
  const w=fixture(), before=generation(w);
  assert.deepEqual(arrayValues(w),[10,20,30]); assert.equal(w.getValue(2,0,0),99);
  w.transact(()=>w.setInput(0,0,0,'1'));
  const changed=generation(w);
  assert.ok(changed[1]>before[1]);assert.ok(changed[2]>before[2]);assert.equal(changed[3],before[3]);
  assert.ok(cellData(w.getCell(1,0,2)).staleCachedArray);assert.equal(cellData(w.getCell(2,0,0)).staleCached,99);
  const after=new Workbook(w.serialize());assert.deepEqual(arrayValues(after),[10,null,null]);assert.equal(after.getValue(2,0,0).code,'#NAME?');
  w.undo();const undone=generation(w);assert.ok(undone[1]>changed[1]);assert.ok(undone[2]>changed[2]);assert.equal(undone[3],before[3]);
  assert.deepEqual(arrayValues(w),[10,20,30]);assert.equal(w.getValue(2,0,0),99);
  w.redo();const redone=generation(w);assert.ok(redone[1]>undone[1]);assert.ok(redone[2]>undone[2]);assert.equal(redone[3],before[3]);
  assert.deepEqual(arrayValues(new Workbook(w.serialize())),[10,null,null]);
});
test('이미 무효한 캐시와 저장 캐시 없는 수식은 추가 저장 세대를 만들지 않음', () => {
  const w=fixture();w.transact(()=>w.setInput(0,0,0,'1'));const before=generation(w);
  for(let si=1;si<=2;si++)w.markFormulaDirty(si,0,si===1?2:0,w.getCell(si,0,si===1?2:0));
  assert.deepEqual(generation(w),before);
  w.transact(()=>w.setInput(3,1,0,'=A1+1'));const plain=generation(w);w.markFormulaDirty(3,1,0,w.getCell(3,1,0));assert.deepEqual(generation(w),plain);
});
test('수동 계산은 F9까지 저장 캐시 세대를 유지하고 계산 시에만 의존 시트 갱신', () => {
  const w=fixture();w.manualCalc=true;const before=generation(w);w.transact(()=>w.setInput(0,0,0,'1'));
  assert.equal(generation(w)[1],before[1]);assert.equal(generation(w)[2],before[2]);assert.deepEqual(arrayValues(w),[10,20,30]);assert.equal(w.getValue(2,0,0),99);
  w.calculateNow();assert.ok(generation(w)[1]>before[1]);assert.ok(generation(w)[2]>before[2]);assert.deepEqual(arrayValues(w),[10,null,null]);assert.equal(w.getValue(2,0,0).code,'#NAME?');
});

test('지원되는 일반 스칼라 수식도 첫 의존 편집 전에 캐시 신뢰 변경을 저장', () => {
  const w=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'2'}}},{name:'Result',fileValues:true,cells:{'0,0':{raw:'=Source!A1*10',cached:99}}}]});
  assert.equal(w.getValue(1,0,0),99);const before=w.sheets[1]._ev??0;
  w.transact(()=>w.setInput(0,0,0,'1'));
  assert.ok((w.sheets[1]._ev??0)>before);assert.equal(cellData(w.getCell(1,0,0)).staleCached,99);
  assert.equal(new Workbook(w.serialize()).getValue(1,0,0),10);
  w.undo();assert.equal(new Workbook(w.serialize()).getValue(1,0,0),20);
});

for (const clear of ['input', 'range']) test(`수동 계산에서 서식 빈 셀 ${clear} 삭제와 실행 취소는 불변 셀을 바꾸지 않음`, () => {
  const w=new Workbook({sheets:[{name:'S',fileValues:true,cells:{
    '0,0':{raw:'2',style:{fill:'#ffff00'}},
    '0,1':{raw:'=A1+1',cached:99},
  }}]});
  w.manualCalc=true;
  const erase=()=>clear==='input'?w.setInput(0,0,0,''):w.clearRange(0,0,0,0,0,'contents');
  assert.doesNotThrow(()=>w.transact(erase));
  const blank=w.getCell(0,0,0);
  assert.ok(Object.isFrozen(blank));assert.equal(blank.raw,'');assert.equal(blank.style.fill,'#ffff00');
  assert.equal(Object.hasOwn(blank,'dirty'),false);assert.equal(w.getValue(0,0,1),99);
  assert.doesNotThrow(()=>w.undo());assert.equal(w.getValue(0,0,0),2);
  assert.doesNotThrow(()=>w.redo());assert.equal(w.getValue(0,0,0),null);assert.equal(w.getValue(0,0,1),99);
  assert.doesNotThrow(()=>w.calculateNow());assert.equal(w.getValue(0,0,1),1);
  assert.equal(cellData(w.getCell(0,0,1)).staleCached,99);
  assert.equal(new Workbook(w.serialize()).getValue(0,0,1),1);
});
