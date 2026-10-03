import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData } from '../src/pivot.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';

function fixture(compact = false) {
  const rows = [['분류','값'], ['가',99], ['나',null], ['',0], [true,{error:'#N/A'}]];
  let snapshot = rows;
  if (compact) { const b = new PivotSnapshotBuilder(rows[0], rows.length-1); for (const row of rows.slice(1)) b.add(row); snapshot = b.finish(); }
  const def = {source:'원본',range:{r1:0,c1:0,r2:4,c2:1},snapshotId:'saved-cache',rows:['분류'],values:[{field:'값',agg:'sum'}]};
  return new Workbook({pivotSnapshots:{'saved-cache':snapshot},sheets:[
    {name:'원본',cells:{'0,0':{raw:'분류'},'0,1':{raw:'값'},'1,0':{raw:'가'},'1,1':{raw:'5'}}},
    {name:'보고서',cells:{},pivot:def}
  ]});
}
const source = w => pivotSourceData(w,w.sheets[1].pivot);
for (const compact of [false,true]) test(`피벗 캐시는 일반 JSON/Blob/메타 복원에서 원본과 다른 저장값을 유지: compact=${compact}`,async()=>{
  const w=fixture(compact),expected=source(w).rows;
  assert.equal(expected[1][1],99);assert.equal(w.getValue(0,1,1),5);
  const data=JSON.parse(JSON.stringify(w.serialize()));
  const copies=[new Workbook(data),new Workbook(JSON.parse(await w.serializeBlob().text())),new Workbook({...w.bookMeta(),sheets:w.serialize().sheets})];
  for(const copy of copies){assert.deepEqual(source(copy).rows,expected);assert.equal(copy.getValue(0,1,1),5);}
  assert.equal(w.bookMeta(false).pivotSnapshots,undefined);
});
test('원본 변경 후 아직 피벗을 조회하지 않았어도 오래된 저장 캐시를 다시 저장하지 않는다',async()=>{
  const w=fixture(true);source(w);
  w.transact(()=>w.setInput(0,1,1,'123'));
  assert.deepEqual(w.snapshotData(),{});
  const copy=new Workbook(JSON.parse(await w.serializeBlob().text()));
  assert.equal(source(copy).rows[1][1],123);
});
test('보고서 상태 변경·Undo·Redo는 캐시 내용을 보존한다',async()=>{
  const w=fixture(true),before=source(w).rows;
  w.transact(()=>w.setSheetProp(1,'state','hidden'));
  for(let i=0;i<3;i++){const copy=new Workbook(JSON.parse(await w.serializeBlob().text()));assert.deepEqual(source(copy).rows,before);if(!i)w.undo();else w.redo();}
});

test('첫 피벗 조회 전에 원본을 편집해도 오래된 캐시가 자동 저장본에서 되살아나지 않는다',async()=>{
  const w=fixture(true);
  w.transact(()=>w.setInput(0,1,1,'123'));
  assert.deepEqual(w.snapshotData(),{});
  const copies=[new Workbook(JSON.parse(JSON.stringify(w.serialize()))),new Workbook(JSON.parse(await w.serializeBlob().text()))];
  for(const copy of copies)assert.equal(source(copy).rows[1][1],123);
});


test('원본 서식·메모·크기 변경과 Undo/Redo는 피벗 저장값을 버리지 않는다',async()=>{
  const w=fixture(true),before=source(w).rows,epoch=w.sourceVersion(0);
  w.transact(()=>{
    w.setStyle(0,1,1,{bold:true,color:'#f00',numFmt:'0.00'});
    const cell=w.getCell(0,1,1);
    w.setCellData(0,1,1,{raw:cell.raw,style:cell.style,comment:'메모'});
    w.setColWidth(0,1,160);w.setRowHeight(0,1,48);
    w.setSheetProp(0,'allStyle',{font:'Arial'});
    w.setSheetProp(0,'zoom',70);
  });
  for(let i=0;i<3;i++){
    assert.equal(w.sourceVersion(0),epoch);assert.deepEqual(source(w).rows,before);
    assert.deepEqual(source(new Workbook(JSON.parse(await w.serializeBlob().text()))).rows,before);
    if(!i)w.undo();else w.redo();
  }
  w.transact(()=>w.setInput(0,1,1,'123'));
  assert.equal(source(w).rows[1][1],123);
});

test('표의 색과 필터 버튼만 바꾸면 저장 캐시를 유지하고 범위 변경은 무효화한다',()=>{
  const w=fixture(true),t={id:'t',name:'Table1',r1:0,c1:0,r2:4,c2:1,header:true,totals:false};
  w.sheets[0].tables=[t];w.sheets[1].pivot.table='Table1';
  const before=source(w).rows,epoch=w.sourceVersion(0);
  w.transact(()=>w.setSheetProp(0,'tables',[{...t,style:'TableStyleMedium2',filter:{hiddenButtons:{1:true}}}]));
  assert.equal(w.sourceVersion(0),epoch);assert.deepEqual(source(w).rows,before);
  w.undo();assert.deepEqual(source(w).rows,before);w.redo();assert.deepEqual(source(w).rows,before);
  w.transact(()=>w.setSheetProp(0,'tables',[{...t,r2:5}]));
  assert.ok(w.sourceVersion(0)>epoch);assert.equal(source(w).rows[1][1],5);
});

test('블록 열 서식은 캐시 유지, 블록 값 변경과 정렬은 캐시 무효화',async()=>{
  const {ColBuilder}=await import('../src/block.js');
  const a=new ColBuilder(1002),b=new ColBuilder(1002);
  for(let i=0;i<1002;i++){a.set(i,'가');b.set(i,1002-i);}
  const w=fixture(true);w.sheets[0].cells.deleteRC(1,0);w.sheets[0].cells.deleteRC(1,1);
  w.sheets[0].blocks=[{r0:1,c0:0,n:1002,ver:0,cols:[a.finish(),b.finish()]}];
  const before=source(w).rows,epoch=w.sourceVersion(0);
  w.transact(()=>w.setBlockStyle(0,0,1,{bold:true,numFmt:'0.0'}));
  assert.equal(w.sourceVersion(0),epoch);assert.deepEqual(source(w).rows,before);
  w.undo();assert.deepEqual(source(w).rows,before);w.redo();assert.deepEqual(source(w).rows,before);
  w.transact(()=>assert.equal(w.sortBlock(0,1,0,1002,1,1,true),true));
  assert.ok(w.sourceVersion(0)>epoch);assert.equal(source(w).rows[1][1],1);
  w.undo();assert.equal(w.getValue(0,1,1),1002);assert.equal(source(w).rows[1][1],99);w.redo();assert.equal(source(w).rows[1][1],1);
});


test('원본의 다른 시트 참조 수식이 바뀌면 캐시 무효화, Undo/Redo는 저장본과 편집본을 복원',()=>{
 const w=fixture(true);w.sheets.push(new Workbook({sheets:[{name:'입력',cells:{'0,0':{raw:'5'}}}]}).sheets[0]);
 w.setCellData(0,1,1,{raw:'=입력!A1',cached:5});
 w.setSnapshots({pivotSnapshots:{'saved-cache':[['분류','값'],['가',99]]}});
 w.transact(()=>w.setInput(2,0,0,'20'));
 assert.equal(w.getValue(0,1,1),20);assert.deepEqual(w.snapshotData(),{});assert.equal(source(w).rows[1][1],20);
 w.undo();assert.equal(w.getValue(0,1,1),5);assert.equal(source(w).rows[1][1],99);
 w.redo();assert.equal(w.getValue(0,1,1),20);assert.equal(source(w).rows[1][1],20);
});

test('날짜 기준 변경 전의 피벗 저장 캐시를 Undo에서 복원',()=>{
 const w=fixture(true);assert.equal(source(w).rows[1][1],99);
 w.transact(()=>w.setDate1904(true));assert.equal(w.date1904,true);assert.equal(source(w).rows[1][1],5);
 w.undo();assert.equal(w.date1904,false);assert.equal(source(w).rows[1][1],99);
 w.redo();assert.equal(w.date1904,true);assert.equal(source(w).rows[1][1],5);
});

test('원본 값 변경을 되돌리면 원본과 달랐던 피벗 저장값까지 복원',()=>{
 const w=fixture(true);w.transact(()=>w.setInput(0,1,1,'20'));assert.equal(source(w).rows[1][1],20);
 w.undo();assert.equal(w.getValue(0,1,1),5);assert.equal(source(w).rows[1][1],99);
 w.redo();assert.equal(w.getValue(0,1,1),20);assert.equal(source(w).rows[1][1],20);
});


test('같은 편집 버전의 다른 원본 시트나 범위로 이전 캐시를 재사용하지 않는다',()=>{
 for(const change of ['sheet','range']){
  const w=fixture(true);const before=source(w).rows;assert.equal(before[1][1],99);
  if(change==='sheet'){
   w.sheets.push(new Workbook({sheets:[{name:'새 원본',cells:{'0,0':{raw:'분류'},'0,1':{raw:'값'},'1,0':{raw:'나'},'1,1':{raw:'123'}}}]}).sheets[0]);
   w.sheets[1].pivot.source='새 원본';
   assert.deepEqual(w.snapshotData(),{});assert.equal(source(w).rows[1][1],123);
  }else{
   w.sheets[1].pivot.range={r1:0,c1:0,r2:1,c2:1};
   assert.deepEqual(w.snapshotData(),{});assert.equal(source(w).rows[1][1],5);
  }
 }
});
