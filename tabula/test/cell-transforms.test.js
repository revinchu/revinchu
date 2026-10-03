import test from 'node:test';
import assert from 'node:assert/strict';
import {CellMap} from '../src/cellmap.js';
import {shiftStoredCells,moveStoredCells} from '../src/cell-transforms.js';
import {Workbook} from '../src/workbook.js';

const dump = cells => [...cells].sort(([a],[b])=>a.localeCompare(b));
function fixture(){const c=new CellMap();for(let col=0;col<4;col++)c.setRunRC(3,col,400,{raw:'',style:{fill:'#abcdef'}});c.setRC(8,1,{raw:'5'});c.setRC(350,2,{raw:'=A1'});return c;}

test('압축된 행/열 삽입·삭제와 부분 셀 밀기는 각 좌표를 이동한 기준과 일치',()=>{
 for(const axis of ['row','col'])for(const index of [0,3,4,8,200,403,500])for(const count of [-5,-1,1,6])for(const band of [null,[0,0],[1,2],[7,9],[180,280],[399,410]]){
  const c=fixture(),expected=new CellMap(),isRow=axis==='row';
  c.forEachRC((cell,r,col)=>{const p=isRow?r:col,o=isRow?col:r;if(band&&(o<band[0]||o>band[1])){expected.setRC(r,col,cell);return;}if(count<0&&p>=index&&p<index-count)return;const next=p>=index?p+count:p;expected.setRC(isRow?next:r,isRow?col:next,cell);});
  assert.deepEqual(dump(shiftStoredCells(c,axis,index,count,band)),dump(expected),JSON.stringify({axis,index,count,band}));
 }
});

test('겹치는 셀 이동은 대상 덮어쓰기·부분 서식 범위를 보존',()=>{
 for(const src of [{r1:0,c1:0,r2:9,c2:2},{r1:7,c1:1,r2:360,c2:2},{r1:300,c1:0,r2:405,c2:3}])for(const dr of [-3,0,2,200])for(const dc of [-1,0,1,4]){
  if(src.r1+dr<0||src.c1+dc<0)continue;
  const c=fixture(),expected=new CellMap(),moved=[];
  c.forEachRC((cell,r,col)=>{if(r>=src.r1&&r<=src.r2&&col>=src.c1&&col<=src.c2){moved.push([r+dr,col+dc,cell]);return;}if(r>=src.r1+dr&&r<=src.r2+dr&&col>=src.c1+dc&&col<=src.c2+dc)return;expected.setRC(r,col,cell);});
  for(const [r,col,cell]of moved)expected.setRC(r,col,cell);
  assert.deepEqual(dump(moveStoredCells(c,src,dr,dc)),dump(expected));
 }
});

test('백만 서식 셀에서 작은 지우기·행 삽입·삭제·이동·Undo는 압축을 유지',()=>{
 const w=new Workbook({sheets:[{name:'서식',cellRuns:[[0,0,1_000_000,{raw:'',style:{fill:'#abcdef'}}]],cells:{'3,2':{raw:'=A11'}}}]});
 w.setInput(0,10,0,'20');
 const compressed=()=>assert.ok([...w.sheets[0].cells.storageEntries()].length<20);
 w.transact(()=>w.clearRange(0,8,0,12,0,'contents'));assert.equal(w.getRaw(0,10,0),'');assert.equal(w.styleAt(0,10,0).fill,'#abcdef');compressed();
 w.undo();assert.equal(w.getRaw(0,10,0),'20');compressed();
 w.transact(()=>w.insertRows(0,10,3));assert.equal(w.getRaw(0,13,0),'20');assert.equal(w.getRaw(0,3,2),'=A14');compressed();
 w.undo();assert.equal(w.getRaw(0,10,0),'20');compressed();
 w.transact(()=>w.deleteRows(0,9,3));assert.equal(w.getRaw(0,3,2),'=#REF!');compressed();w.undo();assert.equal(w.getRaw(0,10,0),'20');
 w.transact(()=>assert.equal(w.moveRange(0,{r1:8,c1:0,r2:12,c2:0},20,1),null));assert.equal(w.getRaw(0,30,1),'20');assert.equal(w.getRaw(0,3,2),'=B31');compressed();
 w.undo();assert.equal(w.getRaw(0,10,0),'20');assert.equal(w.getRaw(0,3,2),'=A11');compressed();
});
