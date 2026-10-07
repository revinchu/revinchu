import {test} from 'node:test';
import assert from 'node:assert/strict';
import {CellMap,getSharedBlankCell} from '../src/cellmap.js';
import {Workbook,makeCellRC} from '../src/workbook.js';
import {shiftStoredCells,moveStoredCells} from '../src/cell-transforms.js';
import {DepGraph} from '../src/depgraph.js';
import {markPreparedWorkbook} from '../src/prepared-sheet-data.js';
import {readXlsx,writeXlsx} from '../src/xlsx.js';
const blank=getSharedBlankCell({fill:'#aabbcc'});
const formula=(raw='=A1')=>makeCellRC({raw},0,0);
const links=cells=>{const out=[];cells.forEachLinkRC((v,r,c)=>out.push([r,c,v.link]));return out.sort((a,b)=>a[1]-b[1]||a[0]-b[0]);};
const formulas=cells=>[...cells.formulaEntries()].map(([r,c,v])=>[r,c,v.raw]).sort((a,b)=>a[1]-b[1]||a[0]-b[0]);
const matchesStorage=cells=>{
 const fs=[],ls=[];for(const[r,c,v]of cells.storageEntries()){if(v.formula)fs.push([r,c,v.raw]);if(v.link)ls.push([r,c,v.link]);}
 const sort=a=>a.sort((x,y)=>x[1]-y[1]||x[0]-y[0]);assert.deepEqual(formulas(cells),sort(fs));assert.deepEqual(links(cells),sort(ls));
};
test('candidate index updates formula/link transitions, overwrites, deletion and clear',()=>{
 const cells=new CellMap();cells.setRC(9,1,{raw:'1'});cells.setRC(9,1,formula());cells.setRC(3,2,{raw:'linked',link:'#A5'});cells.setRC(9,1,{...formula(),link:'#B2'});matchesStorage(cells);
 assert.equal([...cells.formulaEntries()].length,1);assert.equal(links(cells).length,2);
 cells.setRC(9,1,{raw:'7'});assert.equal(formulas(cells).length,0);assert.equal(links(cells).length,1);
 assert.equal(cells.deleteRC(3,2),true);assert.equal(cells.deleteRC(3,2),false);assert.equal(links(cells).length,0);
 cells.set('1,4',formula());cells.clear();assert.equal(cells.size,0);matchesStorage(cells);assert.equal(cells.candidates.size,0);
});
test('raw import candidates normalize into formulas, forced formulas and text without stale values',()=>{
 const cells=new CellMap(new Map([['0,0',{raw:'=A1'}],['1,0',{raw:'=A1',inputType:'text'}],['2,0',{raw:'=A1',style:{numFmt:'text'},fx:true}],['3,0',{raw:'linked',link:'#A5'}]]));
 assert.equal(formulas(cells).length,0);
 for(const n of cells.mapValues((v,r,c)=>makeCellRC(v,r,c),v=>makeCellRC(v)))assert.ok(n>0);
 matchesStorage(cells);assert.deepEqual(formulas(cells).map(x=>x.slice(0,2)),[[0,0],[2,0]]);
 for(const n of cells.mapValues((v,r,c)=>r===0?null:r===3?{raw:'plain'}:v,v=>v))assert.ok(n>0);
 matchesStorage(cells);assert.equal(formulas(cells).length,1);assert.equal(links(cells).length,0);
});
test('compressed blank overwrites remove only candidates within the covered interval',()=>{
 const cells=new CellMap();cells.setRunRC(0,0,200000,blank);cells.setRC(300,0,formula());cells.setRC(400,0,{raw:'linked',link:'#A5'});cells.setRC(700,0,formula());
 cells.setRunRC(250,0,200,blank);matchesStorage(cells);assert.deepEqual(formulas(cells).map(x=>x[0]),[700]);assert.equal(links(cells).length,0);assert.equal(cells.size,200000);
 assert.ok([...cells.storageEntries()].length<10);assert.equal(cells.candidates.get(0).size,1);
});
test('multi-row nonblank runs keep their existing rejection contract and candidate contents',()=>{
 const cells=new CellMap();cells.setRC(0,0,formula());const before=formulas(cells);
 assert.throws(()=>cells.setRunRC(0,0,2,formula()),/빈 셀 범위/);assert.deepEqual(formulas(cells),before);
 assert.throws(()=>cells.setRunRC(0,0,2,{raw:'linked',link:'#A5'}),/빈 셀 범위/);assert.deepEqual(formulas(cells),before);
});
test('formula/link indexes follow rows, columns, partial-band moves and CellMap copies',()=>{
 const cells=new CellMap();cells.setRunRC(0,0,1000,blank);cells.setRC(100,0,formula());cells.setRC(101,1,{raw:'linked',link:'#A5'});
 for(const[axis,index,count,band]of [['row',50,3,null],['row',99,-5,null],['col',0,2,null],['col',0,-1,null],['row',50,2,[0,0]]]){const next=shiftStoredCells(cells,axis,index,count,band);matchesStorage(next);matchesStorage(new CellMap(next));}
 const moved=moveStoredCells(cells,{r1:100,c1:0,r2:101,c2:1},4,3);matchesStorage(moved);assert.deepEqual(formulas(moved).map(x=>x.slice(0,2)),[[104,3]]);assert.equal(links(moved)[0][0],105);
});
test('candidate visitation stays correct while callbacks replace or remove indexed cells',()=>{
 const cells=new CellMap();for(let r=0;r<20;r++)cells.setRC(r,0,{...formula(),link:'#A5'});
 let seen=0;cells.forEachLinkRC((v,r,c)=>{seen++;cells.setRC(r,c,{raw:'plain'});});assert.equal(seen,20);assert.equal(formulas(cells).length,0);assert.equal(links(cells).length,0);
});
test('random writes, deletes, compact runs and normalization match a complete stored-cell scan',()=>{
 const cells=new CellMap();let seed=901;const random=n=>((seed=(Math.imul(seed,1664525)+1013904223)>>>0)%n);
 for(let i=0;i<1000;i++){const r=random(100),c=random(5),op=random(6);if(op===0)cells.deleteRC(r,c);else if(op===1)cells.setRunRC(r,c,random(20)+1,blank);else cells.setRC(r,c,op===2?formula():op===3?{raw:'link',link:'#A5'}:{raw:String(i)});if(i%20===0)matchesStorage(cells);}matchesStorage(cells);
});
function book(){return new Workbook({sheets:[{name:'Source',cells:{'4,0':{raw:'9'},'1,2':{raw:'=A5'}},cellRuns:[[10,7,20000,{raw:'',style:{italic:true}}]]},{name:'Report',fileValues:true,cells:{'0,0':{raw:'=Source!A5',cached:9},'2,1':{raw:'link',link:'#Source!A5'},'3,1':{raw:'web',link:'https://example.invalid/#Source!A5'}}},{name:'Values',cells:{'0,0':{raw:'2'}},cellRuns:[[1,0,20000,{raw:'',style:{bold:true}}]]}]});}
test('row insertion touches formula/link candidates without scanning unrelated stored values',()=>{
 const wb=book(),values=wb.sheets[2].cells;values.storageEntries=()=>{throw Error('unrelated full scan');};values.forEachStoredRC=()=>{throw Error('unrelated full scan');};
 wb.transact(()=>wb.insertRows(0,3,2));assert.equal(wb.getCell(1,0,0).raw,'=Source!A7');assert.equal(wb.getCell(1,2,1).link,'#Source!A7');assert.equal(wb.getCell(1,3,1).link,'https://example.invalid/#Source!A5');
 assert.deepEqual([...wb.affected(0)].sort(),[0,1]);assert.equal(wb.getValue(1,0,0),9);assert.deepEqual(wb.arrayCandidates(),[]);new DepGraph(wb);
});
test('insert, delete, undo and redo rebuild target coordinates and preserve foreign formula/link indexes',()=>{
 const wb=book();wb.invalidate(undefined,false);const before=wb.serialize();wb.transact(()=>wb.insertRows(0,3,2));const after=wb.serialize();wb.sheets.forEach(s=>matchesStorage(s.cells));
 wb.undo();assert.deepEqual(wb.serialize(),before);wb.sheets.forEach(s=>matchesStorage(s.cells));wb.redo();assert.deepEqual(wb.serialize(),after);wb.sheets.forEach(s=>matchesStorage(s.cells));
 wb.transact(()=>wb.deleteCols(0,0,1));assert.match(wb.getCell(1,0,0).raw,/#REF!/);wb.sheets.forEach(s=>matchesStorage(s.cells));wb.undo();assert.deepEqual(wb.serialize(),after);wb.sheets.forEach(s=>matchesStorage(s.cells));
});
test('ordinary formula/link edits and clearing update future dependency and link rewrites',()=>{
 const wb=book();wb.transact(()=>wb.setCellData(2,0,0,{raw:'=Source!A5',link:'#Source!A5'}));assert.ok(wb.affected(0).has(2));wb.transact(()=>wb.insertRows(0,3,1));assert.equal(wb.getCell(2,0,0).raw,'=Source!A6');assert.equal(wb.getCell(2,0,0).link,'#Source!A6');
 wb.transact(()=>wb.setCellData(2,0,0,{raw:'plain'}));wb.invalidateStructure();assert.ok(!wb.affected(0).has(2));assert.equal(links(wb.sheets[2].cells).length,0);
});
test('defined names, table references, dynamic and malformed formulas keep conservative dependencies',()=>{
 const wb=new Workbook({names:[{name:'Chosen',ref:'=Source!A1'}],sheets:[{name:'Source',cells:{'0,0':{raw:'5'},'1,0':{raw:'6'}},tables:[{name:'DataTable',r1:0,c1:0,r2:1,c2:0,columns:['Amount'],header:true}]},{name:'Report',cells:{'0,0':{raw:'=Chosen'},'1,0':{raw:'=SUM(DataTable[Amount])'},'2,0':{raw:'=INDIRECT("Source!A1")'},'3,0':{raw:'=SUM('}}}]});
 assert.equal(wb.sheetDeps()[1].all,true);assert.ok(wb.sheetDeps()[1].sheets.has(0));wb.transact(()=>wb.setNames([{name:'Chosen',ref:'=Source!A2'}]));assert.equal(wb.sheetDeps()[1].all,true);wb.undo();assert.equal(wb.names[0].ref,'=Source!A1');
 wb.transact(()=>wb.renameSheet(0,'Renamed'));assert.equal(wb.getCell(1,1,0).raw,'=SUM(DataTable[Amount])');assert.equal(wb.sheetDeps()[1].all,true);assert.ok(wb.sheetDeps()[1].sheets.has(0));wb.undo();wb.sheets.forEach(s=>matchesStorage(s.cells));
});
test('JSON, owned CellMap and native XLSX hydration reconstruct formula/link candidates',()=>{
 const original=book();original.sheets[0].cells.deleteRC(1,2);for(const restored of [new Workbook(JSON.parse(JSON.stringify(original.serialize()))),new Workbook(readXlsx(writeXlsx(original)).data)]){restored.sheets.forEach(s=>matchesStorage(s.cells));restored.transact(()=>restored.insertRows(0,3,1));assert.equal(restored.getCell(1,0,0).raw,'=Source!A6');assert.equal(restored.getCell(1,2,1).link,'#Source!A6');}
 const cells=new CellMap(new Map([['0,0',{raw:'=Source!A5'}],['1,0',{raw:'link',link:'#Source!A5'}]]));const restored=new Workbook({sheets:[{name:'Source',cells:{'4,0':{raw:'2'}}},{name:'Owned',cells}]});matchesStorage(restored.sheets[1].cells);restored.transact(()=>restored.insertRows(0,3,1));assert.equal(restored.getCell(1,0,0).raw,'=Source!A6');assert.equal(restored.getCell(1,1,0).link,'#Source!A6');
});

test('moving a table between sheets and Undo/Redo rebuild context against indexed formulas',()=>{
 const table={name:'DataTable',r1:0,c1:0,r2:1,c2:0,columns:['Amount'],header:true};
 const wb=new Workbook({sheets:[{name:'First',tables:[table],cells:{'0,0':{raw:'Amount'},'1,0':{raw:'5'}}},{name:'Report',cells:{'0,0':{raw:'=SUM(DataTable[Amount])'}}},{name:'Second',cells:{'0,0':{raw:'Amount'},'1,0':{raw:'7'}}}]});
 assert.deepEqual([...wb.sheetDeps()[1].sheets],[0]);wb.transact(()=>{wb.setSheetProp(0,'tables',[]);wb.setSheetProp(2,'tables',[table]);});
 assert.deepEqual([...wb.sheetDeps()[1].sheets],[2]);wb.undo();assert.deepEqual([...wb.sheetDeps()[1].sheets],[0]);wb.redo();assert.deepEqual([...wb.sheetDeps()[1].sheets],[2]);wb.sheets.forEach(s=>matchesStorage(s.cells));
});
test('prepared exclusive restore adopts already indexed formula/link cells without a second scan',()=>{
 const cells=new CellMap();cells.setRC(0,0,makeCellRC({raw:'=A2',cached:7},0,0));cells.setRC(1,0,makeCellRC({raw:'7',link:'#A2'},1,0));
 const data={sheets:[{name:'Prepared',cells,blocks:[]}]};markPreparedWorkbook(data,[false]);const wb=new Workbook(data);assert.equal(wb.sheets[0].cells,cells);matchesStorage(cells);
 wb.transact(()=>wb.insertRows(0,1,1));assert.equal(wb.getCell(0,0,0).raw,'=A3');assert.equal(wb.getCell(0,2,0).link,'#A3');assert.equal(wb.getValue(0,0,0),7);wb.undo();wb.sheets.forEach(s=>matchesStorage(s.cells));
});
