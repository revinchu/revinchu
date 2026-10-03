import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData } from '../src/pivot.js';

const fixture=()=>new Workbook({sheets:[{name:'원본',fileValues:true,cells:{'0,0':{raw:'분류'},'0,1':{raw:'값'},'1,0':{raw:'가'},'1,1':{raw:'=2+3',cached:99}}},{name:'다음',cells:{}}]});

test('시트 표시 상태는 Undo/Redo와 저장 메타를 유지하며 셀 데이터 버전을 바꾸지 않는다',()=>{
 const w=fixture(),s=w.sheets[0],cells=s.cells,ev=s._ev??0,version=w.version;
 const before=w.serialize();let serial=0;w.serialize=()=>{serial++;throw Error('전체 복사 금지');};w.serializeSheet=()=>{serial++;throw Error('시트 복사 금지');};
 w.transact(()=>w.setSheetProp(0,'state','hidden'));
 assert.equal(s._ev??0,ev);assert.ok(w.version>version);assert.equal(w.sheetMeta(0).state,'hidden');assert.equal(s.cells,cells);
 assert.deepEqual(w.undoStack[0].entries,[{t:'prop',si:0,prop:'state',before:undefined,after:'hidden'}]);
 w.undo();assert.equal(s.state,undefined);assert.equal(s._ev??0,ev);
 w.redo();assert.equal(s.state,'hidden');assert.equal(s._ev??0,ev);assert.equal(serial,0);
 const restored=new Workbook({...before,sheets:[{...before.sheets[0],...w.sheetMeta(0)},before.sheets[1]]});
 assert.equal(restored.sheets[0].state,'hidden');assert.equal(restored.getValue(0,1,1),99);
});

test('숨김 변경은 파일 계산값과 피벗 원본 스냅샷을 버리지 않는다',()=>{
 const w=fixture(),def={source:'원본',range:{r1:0,c1:0,r2:1,c2:1},snapshotId:'cache'};
 w.pivotSnapshots=new Map([['cache',{rows:[['분류','값'],['가',99]],ver:w.sheets[0]._ev??0}]]);
 const snapshot=w.pivotSnapshots.get('cache'),beforeVersion=w.sheetVersion(0);assert.equal(w.getValue(0,1,1),99);
 w.transact(()=>w.setSheetProp(0,'state','veryHidden'));
 assert.equal(w.sheetVersion(0),beforeVersion);assert.equal(w.getValue(0,1,1),99);
 assert.equal(pivotSourceData(w,def).rows[1][1],99);assert.equal(w.pivotSnapshots.get('cache'),snapshot);
 w.undo();assert.equal(w.getValue(0,1,1),99);assert.equal(w.pivotSnapshots.get('cache'),snapshot);
});

test('숨김과 함께 바뀐 셀 또는 다른 메타는 데이터 변경 표시를 유지한다',()=>{
 const w=fixture(),s=w.sheets[0];let ev=s._ev??0;
 w.transact(()=>{w.setSheetProp(0,'state','hidden');w.setInput(0,1,0,'나');});assert.ok(s._ev>ev);ev=s._ev;
 w.undo();assert.equal(w.getRaw(0,1,0),'가');assert.ok(s._ev>ev);ev=s._ev;
 w.transact(()=>{w.setSheetProp(0,'state','hidden');w.setSheetProp(0,'zoom',150);});assert.ok(s._ev>ev);
});
