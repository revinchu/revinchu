import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap } from '../src/cellmap.js';
import { xlsxRowPlan } from '../src/xlsx-row-stream.js';
import { xlsxBoundedRowPlan } from '../src/xlsx-compact-rows.js';
const ordered = plan => ({ maxR:plan.maxR,maxC:plan.maxC,rows:[...plan.rows].map(([r,cells])=>[r,cells.sort((a,b)=>a[0]-b[0])]) });

test('numeric-key planner matches existing planner for 250 randomized mixed sparse/run/block/spill cases',()=>{
  let seed = 190003; const rand = n => { seed = (Math.imul(seed,1664525)+1013904223)>>>0; return seed%n; };
  for(let trial=0;trial<250;trial++) {
    const cells=new CellMap();
    for(let i=0;i<9;i++) cells.setRunRC(rand(350),rand(16),128+rand(150),{raw:'',style:{fill:'#abcdef',bold:!!rand(2)}});
    for(let i=0;i<500;i++) {const r=rand(600),c=rand(20);if(rand(6)===0)cells.deleteRC(r,c);else cells.setRC(r,c,{raw:String(i)});}
    const before=[...cells.storageEntries()];
    const options={rowLimit:450+rand(150),colLimit:16+rand(4),extraRows:[599,rand(600),rand(600),'123',-1],blocks:[{r0:rand(590),c0:rand(18),n:20,cols:[{},{}]},{r0:400,c0:0,n:0,cols:[{}]}],spills:[{r:rand(600),c:rand(18),h:4,w:4}]};
    assert.deepEqual(ordered(xlsxBoundedRowPlan(cells,options)),ordered(xlsxRowPlan(cells,options)),`trial ${trial}`);
    assert.deepEqual([...cells.storageEntries()],before,'neither planner changes source');
  }
});

test('numeric planner keeps million-cell blank runs compressed and does not create per-cell coordinate copies while planning',()=>{
  const cells=new CellMap();cells.setRunRC(4,3,1000000,{raw:'',style:{bold:true}});
  for(let r=0;r<10000;r++)cells.setRC(r,0,{raw:String(r)});
  cells.setRC(999998,3,{raw:'tail'});cells.setRC(1048578,1,{raw:'outside'});
  const push=Array.prototype.push;let plan;
  try {
    Array.prototype.push=function(...values){for(const value of values)if(Array.isArray(value)&&value.length===2&&typeof value[1]==='object')throw Error('per-cell coordinate clone');return push.apply(this,values);};
    plan=xlsxBoundedRowPlan(cells,{rowLimit:1048576,colLimit:16384});
  } finally {Array.prototype.push=push;}
  assert.equal(plan.maxR,1000003);assert.equal(plan.maxC,3);
  assert.deepEqual(plan.rows.next().value,[0,[[0,cells.getRC(0,0)]]]);plan.rows.return();
});

test('numeric planner supports generic Map fallback and sparse last Excel row without a full-grid scan',()=>{
  const options={rowLimit:1048576,colLimit:16384};
  const generic=new Map([['1048575,16383',{raw:'last'}],['3,2',{raw:'early'}]]);
  assert.deepEqual(ordered(xlsxBoundedRowPlan(generic,options)),ordered(xlsxRowPlan(generic,options)));
  const cells=new CellMap(generic);assert.deepEqual(ordered(xlsxBoundedRowPlan(cells,options)),ordered(xlsxRowPlan(generic,options)));
});
