import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

function fixture({table=false, second=false}={}) {
  const cells={'0,0':{raw:'Kind'},'0,1':{raw:'Amount'}};
  for(let i=0;i<6;i++){cells[`${i+1},0`]={raw:String.fromCharCode(65+i)};cells[`${i+1},1`]={raw:String((i+1)*10)};}
  const pivot={name:'SameSheet',source:'Data',range:{r1:0,c1:0,r2:6,c2:1},snapshotId:'saved',rows:['Kind'],values:[{field:'Amount',agg:'sum'}],top:10,left:4,...(table?{table:'SourceTable'}:{})};
  const saved=[['Kind','Amount'],['A',1],['B',2],['C',3]], data={pivotSnapshots:{saved},sheets:[{name:'Data',cells,pivot,...(table?{tables:[{id:'source',name:'SourceTable',r1:0,c1:0,r2:6,c2:1,header:true}]}:{})}]};
  if(second){data.pivotSnapshots.output=[['Kind','Amount'],['old output',999]];data.sheets[0].pivotsExtra=[{...pivot,name:'OutputConsumer',snapshotId:'output',range:{r1:10,c1:4,r2:14,c2:5},top:20,left:8}];}
  return new Workbook(data);
}
const rows=w=>pivotSourceData(w,w.sheets[0].pivot).rows;
function writePivot(w) {
  const def=w.sheets[0].pivot, resolved=resolvePivot(pivotSourceData(w,def),def), layout=computePivot(resolved,resolved.def);
  for(let r=0;r<layout.grid.length;r++)for(let c=0;c<layout.grid[r].length;c++)w.setCellData(0,def.top+r,def.left+c,layout.grid[r][c]);
}

test('same-sheet pivot output keeps three saved source items instead of switching to six live items',()=>{
  const w=fixture(),snap=w.pivotSnapshots.get('saved'),before=rows(w),epoch=w.sourceVersion(0);
  w.transact(()=>{w.setSheetProp(0,'pivot',{...w.sheets[0].pivot,filters:{Kind:['A']}});writePivot(w);});
  assert.deepEqual(rows(w),before);assert.equal(w.pivotSnapshots.get('saved'),snap);assert.ok(w.sourceVersion(0)>epoch);
  assert.equal(w.snapshotData().saved,snap.rows);
});

test('another pivot whose source overlaps the output is invalidated independently',()=>{
  const w=fixture({second:true}),before=rows(w);
  w.transact(()=>writePivot(w));
  assert.deepEqual(rows(w),before);assert.equal(w.snapshotData().output,undefined);
  assert.ok(!pivotSourceData(w,w.sheets[0].pivotsExtra[0]).rows.some(row=>row.includes('old output')));
});

for(const [r,c,raw] of [[1,1,'123'],[0,0,'Changed header'],[6,0,'Last row']])test(`actual source edit at ${r},${c} invalidates cache with correct Undo/Redo`,()=>{
  const w=fixture(),before=rows(w),saved=w.pivotSnapshots.get('saved').rows;
  w.transact(()=>w.setInput(0,r,c,raw));assert.equal(w.snapshotData().saved,undefined);assert.notDeepEqual(rows(w),before);
  w.undo();assert.deepEqual(rows(w),before);assert.equal(w.pivotSnapshots.get('saved').rows,saved);
  w.redo();assert.equal(w.snapshotData().saved,undefined);assert.notDeepEqual(rows(w),before);
});

test('outside edit, formatting, comments, Undo/Redo and serialization retain saved rows',async()=>{
  const w=fixture(),before=rows(w),saved=w.pivotSnapshots.get('saved').rows;
  w.transact(()=>{w.setInput(0,40,10,'outside');w.setStyle(0,1,1,{bold:true});w.setComment(0,2,1,'note');});
  for(const action of [null,'undo','redo']){
    if(action)w[action]();assert.deepEqual(rows(w),before);assert.equal(w.snapshotData().saved,saved);
    for(const copy of [new Workbook(w.serialize()),new Workbook(JSON.parse(await w.serializeBlob().text())),new Workbook(readXlsx(writeXlsx(w)).data)])assert.deepEqual(rows(copy),before);
  }
});

for(const method of ['insertRows','insertCols','deleteRows','deleteCols'])test(`${method} invalidates source structure and Undo restores cached source`,()=>{
  const w=fixture(),before=rows(w);w.transact(()=>w[method](0,1,1));assert.equal(w.snapshotData().saved,undefined);
  w.undo();assert.deepEqual(rows(w),before);w.redo();assert.equal(w.snapshotData().saved,undefined);
});

test('table resize invalidates the bound cache while table style does not',()=>{
  const w=fixture({table:true}),before=rows(w),t=w.sheets[0].tables[0];
  w.transact(()=>w.setSheetProp(0,'tables',[{...t,style:'TableStyleMedium2'}]));assert.deepEqual(rows(w),before);
  w.transact(()=>w.setSheetProp(0,'tables',[{...t,r2:7}]));assert.equal(w.snapshotData().saved,undefined);
  w.undo();assert.deepEqual(rows(w),before);w.redo();assert.equal(w.snapshotData().saved,undefined);
});

for(const crossSheet of [false,true])test(`source formula depending on outside input invalidates cache (${crossSheet?'cross':'same'} sheet)`,()=>{
  const w=fixture();
  if(crossSheet)w.sheets.push(new Workbook({sheets:[{name:'Input',cells:{'0,0':{raw:'10'}}}]}).sheets[0]);
  else w.setInput(0,30,10,'10');
  w.setCellData(0,1,1,{raw:crossSheet?'=Input!A1':'=K31',cached:10});w.setSnapshots({pivotSnapshots:{saved:[['Kind','Amount'],['A',99]]}});
  const before=rows(w);w.transact(()=>w.setInput(crossSheet?1:0,crossSheet?0:30,crossSheet?0:10,'20'));
  assert.equal(w.snapshotData().saved,undefined);assert.equal(rows(w)[1][1],20);
  w.undo();assert.deepEqual(rows(w),before);w.redo();assert.equal(rows(w)[1][1],20);
});

test('source identity and range changes cannot rebind stale saved cache',()=>{
  for(const variant of ['range','sheet']){
    const w=fixture();if(variant==='range')w.sheets[0].pivot.range={r1:0,c1:0,r2:2,c2:1};
    else{w.sheets.push(new Workbook({sheets:[{name:'Other',cells:{'0,0':{raw:'Kind'},'0,1':{raw:'Amount'}}}]}).sheets[0]);w.sheets[0].pivot.source='Other';}
    assert.equal(w.snapshotData().saved,undefined);
  }
});

test('cell touch hot path never scans all pivot definitions after snapshots are registered',()=>{
  const w=fixture(),before=rows(w),lookup=w.pivotSnapshotSource;
  w.pivotSnapshotSource=()=>{throw Error('per-cell pivot scan');};
  for(let i=0;i<10000;i++)w.touchSource(0,30+i,10);
  w.pivotSnapshotSource=lookup;assert.deepEqual(rows(w),before);
});


test('same source ranges share one watch and ordinary workbooks allocate no watches',()=>{
  const w=fixture(),def=w.sheets[0].pivot,saved=rows(w);
  w.sheets[0].pivotsExtra=[{...def,name:'Shared',snapshotId:'shared',top:30}];
  w.setSnapshots({pivotSnapshots:{saved,shared:saved}});
  assert.equal(w.pivotSnapshots.get('saved').sourceWatch,w.pivotSnapshots.get('shared').sourceWatch);
  assert.equal(w.pivotSourceRanges.get(w.sheets[0]).size,1);
  const plain=new Workbook();assert.equal(plain.pivotSourceRanges,undefined);
  plain.transact(()=>plain.setInput(0,0,0,'42'));assert.equal(plain.pivotSourceRanges,undefined);
  const restored=new Workbook({sheets:[{name:'NoCache',cells:{}}]});assert.equal(restored.pivotSourceRanges,null);
  restored.transact(()=>restored.setInput(0,0,0,'42'));assert.equal(restored.pivotSourceRanges,null);
});

for(const all of [false,true])test(`explicit ${all?'all':'selected'} cache refresh is independently undoable without cell changes`,async()=>{
  const w=fixture({second:true}),before=rows(w),snap=w.pivotSnapshots.get('saved').rows;
  w.transact(()=>assert.equal(w.clearPivotSnapshots(all?null:['saved']),true));
  assert.equal(w.snapshotData().saved,undefined);assert.equal(w.undoStack.length,1);assert.equal(w.pivotSnapshots.has('output'),!all);
  assert.equal(rows(w).length,7);
  const fresh=new Workbook(JSON.parse(await w.serializeBlob().text()));assert.equal(rows(fresh).length,7);
  w.undo();assert.deepEqual(rows(w),before);assert.equal(w.pivotSnapshots.get('saved').rows,snap);assert.ok(w.snapshotData().output);
  w.redo();assert.equal(rows(w).length,7);assert.equal(w.snapshotData().saved,undefined);
});

test('refresh does not revive an already-invalid other cache while rebuilding the watch index',()=>{
  const w=fixture({second:true});w.setInput(0,10,4,'changed output source');
  assert.equal(w.snapshotData().output,undefined);
  w.transact(()=>w.clearPivotSnapshots(['saved']));assert.equal(w.snapshotData().output,undefined);
  w.undo();assert.equal(w.snapshotData().output,undefined);assert.equal(rows(w).length,4);
});

test('unknown dependency fallback invalidates rather than reviving a potentially changed source formula',()=>{
  const w=fixture();w.setInput(0,30,10,'10');w.setCellData(0,1,1,{raw:'=K31',cached:10});
  w.setSnapshots({pivotSnapshots:{saved:[['Kind','Amount'],['A',99]]}});
  // An existing spill takes the conservative sheet recalculation branch.
  w.spills.set('synthetic',{si:0,r:40,c:10,h:1,w:1,values:[[1]]});
  w.transact(()=>w.setInput(0,30,10,'20'));
  assert.equal(w.snapshotData().saved,undefined);assert.equal(rows(w)[1][1],20);
  w.undo();assert.equal(rows(w)[1][1],99);
});

for (const crossSheet of [false, true]) test(`spill anchor outside the watched source invalidates spilled source values (${crossSheet?'cross':'same'} sheet input)`,()=>{
  const pivot={name:'SpillSource',source:'Source',range:{r1:0,c1:1,r2:1,c2:2},snapshotId:'saved',rows:['Kind'],values:[{field:'Amount',agg:'sum'}],top:10,left:4};
  const w=new Workbook({sheets:[{name:'Input',cells:{'0,0':{raw:'1'}}},{name:'Source',cells:{'0,1':{raw:'Kind'},'0,2':{raw:'Amount'},'1,0':{raw:crossSheet?'=SEQUENCE(1,3,Input!A1)':'=SEQUENCE(1,3,E1)'},'0,4':{raw:'1'}},pivot}]});
  assert.deepEqual([0,1,2].map(c=>w.getValue(1,1,c)),[1,2,3]);
  const saved=[['Kind','Amount'],['saved',99]];
  w.setSnapshots({pivotSnapshots:{saved}});
  w.transact(()=>w.setInput(crossSheet?0:1,0,crossSheet?0:4,'10'));
  assert.deepEqual([0,1,2].map(c=>w.getValue(1,1,c)),[10,11,12]);
  assert.equal(w.snapshotData().saved,undefined);
  assert.deepEqual(pivotSourceData(w,w.sheets[1].pivot).rows,[['Kind','Amount'],[11,12]]);
  w.undo();assert.equal(w.snapshotData().saved,saved);
  assert.deepEqual(pivotSourceData(w,w.sheets[1].pivot).rows,saved);
  w.redo();assert.equal(w.snapshotData().saved,undefined);
  assert.deepEqual(pivotSourceData(w,w.sheets[1].pivot).rows,[['Kind','Amount'],[11,12]]);
});
