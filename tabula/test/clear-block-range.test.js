import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { ColBuilder } from '../src/block.js';
import { pivotSourceData } from '../src/pivot.js';
function fixture(n=8,permuted=true){
 const cols=[0,1,2].map(c=>{const b=new ColBuilder(n);for(let r=0;r<n;r++)b.set(r,c===1?['text',true,{error:'#N/A'},0][r%4]:r*10+c);return b.finish(n,{fill:'#ff0000',bold:true});});
 return new Workbook({sheets:[{name:'Source',cells:{},blocks:[{r0:2,c0:1,n,cols,...(permuted?{perm:Uint32Array.from({length:n},(_,i)=>n-1-i)}:{})}]}]});
}
function values(w,r1=2,r2=9){return Array.from({length:r2-r1+1},(_,i)=>[1,2,3].map(c=>w.getRaw(0,r1+i,c)));}
for(const what of ['contents','all','formats'])test(`partial block ${what} preserves neighbors and Undo/Redo`,()=>{
 const w=fixture(),before=values(w);w.transact(()=>w.clearRange(0,4,2,6,2,what));
 for(let r=2;r<=9;r++)for(let c=1;c<=3;c++){
  const selected=r>=4&&r<=6&&c===2;
  assert.equal(w.getRaw(0,r,c),selected&&what!=='formats'?'':before[r-2][c-1]);
  assert.equal(w.getCell(0,r,c)?.style?.fill,selected&&what!=='contents'?undefined:'#ff0000');
 }
 const after=values(w);w.undo();assert.deepEqual(values(w),before);assert.deepEqual(Array.from(w.sheets[0].blocks[0].perm),[7,6,5,4,3,2,1,0]);
 w.redo();assert.deepEqual(values(w),after);assert.ok(w.undoStack.at(-1).entries.length<8);
});
test('block clear preserves overridden comments, links and blank cells without revealing hidden values',()=>{
 const w=fixture();w.transact(()=>{w.setCellData(0,3,1,{raw:'42',style:{fill:'#00ff00'},comment:'keep',link:'https://example.test'});w.setCellData(0,4,1,{raw:'',style:{fill:'#0000ff'}});});
 // A blank CellMap override must win even if an imported block contains a hidden value.
 w.sheets[0].blocks[0].cols[0].num[w.sheets[0].blocks[0].perm[2]]=999;
 w.transact(()=>w.clearRange(0,3,1,4,1,'contents'));
 assert.equal(w.getRaw(0,3,1),'');assert.equal(w.getCell(0,3,1).comment,'keep');assert.equal(w.getCell(0,3,1).link,undefined);assert.equal(w.getCell(0,3,1).style.fill,'#00ff00');assert.equal(w.getRaw(0,4,1),'');
 w.undo();w.transact(()=>w.clearRange(0,3,1,4,1,'formats'));
 assert.equal(w.getRaw(0,3,1),'42');assert.equal(w.getCell(0,3,1).comment,'keep');assert.equal(w.getCell(0,3,1).link,'https://example.test');assert.equal(w.getRaw(0,4,1),'');assert.equal(w.getCell(0,3,1).style,undefined);
 w.undo();assert.equal(w.getRaw(0,4,1),'');assert.equal(w.getCell(0,4,1).style.fill,'#0000ff');
});
for(const what of ['contents','all','formats'])test(`large block ${what} creates a compact Undo entry`,()=>{
 const w=fixture(50010),before=[w.getRaw(0,2,2),w.getRaw(0,50011,2)];
 w.transact(()=>w.clearRange(0,2,2,50011,2,what));
 assert.equal(w.undoStack.length,1);assert.equal(w.undoStack[0].entries.filter(e=>e.t==='cell').length,0);assert.ok(w.sheets[0].cells.size<5);
 assert.equal(w.getRaw(0,2,2),what==='formats'?before[0]:'');assert.equal(w.getRaw(0,2,1),'500090');
 w.undo();assert.deepEqual([w.getRaw(0,2,2),w.getRaw(0,50011,2)],before);
 w.redo();assert.equal(w.getRaw(0,50011,2),what==='formats'?before[1]:'');
});
test('million styled blanks clear without expansion and restore as compact runs',()=>{
 const w=new Workbook({sheets:[{name:'Blank',cellRuns:[[0,0,1000000,{raw:'',style:{fill:'#ff0000'}}]],cells:{}}]});
 w.transact(()=>w.clearRange(0,100,0,900000,0,'all'));
 assert.equal(w.sheets[0].cells.size,100099);assert.equal(w.sheets[0].cells.getRC(500000,0),undefined);assert.equal([...w.sheets[0].cells.storageEntries()].length,2);
 assert.equal(w.undoStack[0].entries.length,1);w.undo();assert.equal(w.sheets[0].cells.size,1000000);assert.equal([...w.sheets[0].cells.storageEntries()].length,1);
});
test('bulk formats/comments/hyperlinks keep saved pivot cache and typed values through Undo',()=>{
 for(const what of ['formats','comments','hyperlinks']){
  const w=fixture(50010),def={name:'P',source:'Source',snapshotId:'cache',range:{r1:2,c1:1,r2:3,c2:2},rows:['Kind'],cols:[],values:[{field:'Value',agg:'sum'}]};
  w.sheets.push(new Workbook().sheets[0]);w.sheets[1].name='Pivot';w.sheets[1].pivot=def;w.setSnapshots({pivotSnapshots:{cache:[['Kind','Value'],['saved',10]]}});
  w.transact(()=>w.clearRange(0,2,1,50011,2,what));assert.equal(pivotSourceData(w,def).cube.row(0)[1],10);
  w.undo();assert.equal(pivotSourceData(w,w.sheets[1].pivot).cube.row(0)[1],10);
 }
});
for(const transactional of [false,true])test(`bulk content clear invalidates a dependent pivot source (${transactional?'transaction':'direct'})`,()=>{
 const base=fixture(50010).serialize();base.sheets[0].name='Input';
 base.sheets.push({name:'Dependent',cells:{'0,0':{raw:'Kind'},'0,1':{raw:'Value'},'1,0':{raw:'row'},'1,1':{raw:'=Input!B3',cached:10}}});
 base.sheets.push({name:'Pivot',cells:{},pivot:{source:'Dependent',snapshotId:'cache',range:{r1:0,c1:0,r2:1,c2:1},rows:['Kind'],cols:[],values:[{field:'Value',agg:'sum'}]}});
 base.pivotSnapshots={cache:[['Kind','Value'],['row',10]]};const w=new Workbook(base),def=w.sheets[2].pivot;
 assert.equal(pivotSourceData(w,def).cube.row(0)[1],10);
 const clear=()=>w.clearRange(0,2,1,50011,1,'contents');if(transactional)w.transact(clear);else clear();
 assert.equal(w.getValue(1,1,1),0);assert.equal(pivotSourceData(w,def).cube.row(0)[1],0);
 if(transactional){w.undo();assert.equal(pivotSourceData(w,w.sheets[2].pivot).cube.row(0)[1],10);w.redo();assert.equal(pivotSourceData(w,w.sheets[2].pivot).cube.row(0)[1],0);}
});
test('clear formats preserves existing inherited row and column formatting semantics',()=>{
 const w=fixture();w.transact(()=>{w.setLineStyle(0,'row',4,{fill:'#00ff00'});w.setLineStyle(0,'col',2,{italic:true});w.clearRange(0,4,2,4,2,'formats');});
 assert.equal(w.styleAt(0,4,2).fill,'#00ff00');assert.equal(w.styleAt(0,4,2).italic,true);assert.equal(w.getCell(0,4,2)?.style,undefined);
});

for(const what of ['contents','all','formats','comments','hyperlinks'])test(`small ${what} clear never enumerates the whole sheet`,()=>{
 const w=new Workbook({sheets:[{name:'Small',cells:{'0,0':{raw:'선택',comment:'메모',link:'#B1',style:{fill:'#123456'}},'0,1':{raw:'이웃'}},cellRuns:[[100,2,1_000_000,{raw:'',style:{fill:'#abcdef'}}]]}]});
 // Prepare dependency metadata before forbidding unrelated whole-sheet reads.
 w.sheetDeps();w.prepareGraph();
 const cells=w.sheets[0].cells,original=cells.storageEntries;
 cells.storageEntries=()=>{throw new Error('small clear enumerated all stored cells');};
 try{w.transact(()=>w.clearRange(0,0,0,0,0,what));assert.equal(w.getRaw(0,0,1),'이웃');}
 finally{cells.storageEntries=original;}
 assert.equal(w.getRaw(0,0,0),['contents','all'].includes(what)?'':'선택');
 w.undo();assert.equal(w.getRaw(0,0,0),'선택');assert.equal(w.getCell(0,0,0).comment,'메모');
 w.redo();assert.equal(w.getRaw(0,0,0),['contents','all'].includes(what)?'':'선택');
});
test('small clear splits a blank run and keeps covered block values hidden through Undo',()=>{
 const w=fixture(200,false);w.sheets[0].cells.setRunRC(2,1,200,{raw:'',style:{fill:'#abcdef'}});
 w.sheetDeps();w.prepareGraph();const cells=w.sheets[0].cells,original=cells.storageEntries;
 cells.storageEntries=()=>{throw new Error('small clear enumerated all stored cells');};
 try{w.transact(()=>w.clearRange(0,70,1,72,1,'contents'));assert.equal(w.getRaw(0,71,1),'');}
 finally{cells.storageEntries=original;}
 assert.equal(w.getRaw(0,69,1),'');assert.equal(w.getRaw(0,73,1),'');
 w.transact(()=>w.clearRange(0,70,1,72,1,'all'));
 assert.equal(w.getRaw(0,71,1),'');assert.equal(w.styleAt(0,69,1).fill,'#abcdef');
 w.undo();assert.equal(w.styleAt(0,71,1).fill,'#abcdef');assert.equal(w.getRaw(0,71,1),'');
});
