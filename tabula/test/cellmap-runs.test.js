import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap, getSharedBlankCell } from '../src/cellmap.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const style = { fill: '#aabbcc', bold: true };
const blank = getSharedBlankCell(style);
function pair(n = 200) { const cells = new CellMap(), map = new Map(); for (let r = 0; r < n; r++) { cells.setRC(r, 0, blank); map.set(r, blank); } return [cells.col(0), map]; }
function same(a, b) { assert.equal(a.size, b.size); assert.deepEqual([...a], [...b]); for (let r = 0; r < 400; r++) { assert.equal(a.has(r), b.has(r)); assert.equal(a.get(r), b.get(r)); } }

test('million styled blank coordinates occupy one reversible run with unchanged logical size', () => {
  const cells = new CellMap(); cells.setRunRC(0, 3, 1_000_000, blank);
  assert.equal(cells.size, 1_000_000);
  assert.deepEqual([...cells.storageEntries()], [[0, 3, blank, 1_000_000]]);
  for (const r of [0, 127, 500_000, 999_999]) { assert.equal(cells.getRC(r, 3), blank); assert.ok(cells.hasRC(r, 3)); }
  assert.equal(cells.hasRC(1_000_000, 3), false);
  cells.setRC(500_000, 3, { raw: 'edited' }); cells.deleteRC(127, 3);
  assert.equal(cells.size, 999_999); assert.equal(cells.getRC(500_000, 3).raw, 'edited');
  assert.equal(cells.getRC(500_001, 3), blank); assert.equal(cells.getRC(127, 3), undefined);
  const restored = new CellMap(); for (const [r,c,v,count] of cells.storageEntries()) restored.setRunRC(r,c,count,v);
  assert.deepEqual([...restored.storageEntries()], [...cells.storageEntries()]);
});

test('run column updates, middle deletions and reinsertions match native Map order', () => {
  const [a,b] = pair();
  let seed = 71; const random = n => ((seed = (Math.imul(seed,1664525)+1013904223)>>>0) % n);
  for (let i=0;i<2500;i++) {
    const r=random(400), action=random(5);
    if (action===0) assert.equal(a.delete(r),b.delete(r));
    else { const value = action===1 ? blank : {raw:String(i)}; a.set(r,value);b.set(r,value); }
    if (i%100===0) same(a,b);
  }
  same(a,b); assert.deepEqual([...a.keys()], [...b.keys()]); assert.deepEqual([...a.values()], [...b.values()]);
});

test('live run iterators skip deletions, observe edits and appended or reinserted cells', () => {
  const [a,b] = pair(), x=a.entries(), y=b.entries();
  assert.deepEqual(x.next(),y.next());
  const edit={raw:'change'};
  for (const target of [a,b]) { target.delete(5);target.delete(0);target.set(0,blank);target.set(6,edit);target.set(300,blank); }
  // Mutable values differ only by identity in the two maps.
  assert.deepEqual([...x], [...y]); same(a,b);
});

test('clear during run iteration exposes only subsequent insertion', () => {
  const [a,b] = pair(), x=a.entries(),y=b.entries(); assert.deepEqual(x.next(),y.next());
  a.clear();b.clear(); for (const target of [a,b]) {target.set(200,blank);target.set(0,blank);}
  assert.deepEqual([...x],[...y]);same(a,b);
});

test('conversion waits for native iterator so newly inserted cells remain visible', () => {
  const [a,b] = pair(20), x=a.entries(),y=b.entries();assert.deepEqual(x.next(),y.next());
  for(let r=20;r<200;r++){a.set(r,blank);b.set(r,blank);}
  assert.deepEqual([...x],[...y]);assert.equal([...a.storageEntries()].length,1);same(a,b);
});

test('overlapping setRun keeps existing insertion order and adds only absent rows', () => {
  const [a,b] = pair();a.delete(20);b.delete(20);
  const other=getSharedBlankCell({italic:true});a.setRun(10,250,other);for(let r=10;r<260;r++)b.set(r,other);
  same(a,b);
});

test('bulk normalization processes shared blanks once and preserves explicit-cell coordinates', () => {
  const cells=new CellMap();cells.setRunRC(4,2,200_000,blank);cells.setRC(10,2,{raw:'=A1'});
  let sharedCalls=0;const visited=[];
  let n=0;for(const count of cells.mapValues((v,r,c)=>{visited.push([r,c]);return {...v,v:1};},v=>{sharedCalls++;return getSharedBlankCell(v.style,true);}))n+=count;
  assert.equal(n,200_000);assert.equal(sharedCalls,2);assert.deepEqual(visited,[[10,2]]);
  assert.equal(cells.size,200_000);assert.equal(cells.getRC(4,2).v,null);assert.equal(cells.getRC(10,2).v,1);
});

test('workbook compact restore preserves formatting, edits and undo without expanding blank runs', async () => {
  const wb=new Workbook({sheets:[{name:'Sheet1',cells:{'0,0':{raw:'7'}},cellRuns:[[1,0,1000,{raw:'',style}]]}]});
  assert.equal(wb.sheets[0].cells.size,1001);assert.equal(wb.getValue(0,0,0),7);
  assert.equal(wb.sheets[0].cells.getRC(900,0).style.fill,style.fill);
  assert.ok([...wb.sheets[0].cells.storageEntries()].length<5);
  wb.transact(()=>wb.setInput(0,500,0,'99'));assert.equal(wb.getValue(0,500,0),99);
  wb.undo();assert.equal(wb.getValue(0,500,0),null);assert.equal(wb.styleAt(0,500,0).fill,style.fill);
  wb.transact(()=>wb.insertRows(0,500,2));assert.equal(wb.getCell(0,500,0),undefined);assert.equal(wb.getCell(0,1002,0).style.fill,style.fill);
  wb.undo();assert.equal(wb.sheets[0].cells.size,1001);
  const saved=wb.serialize();assert.equal(await wb.serializeBlob().text(),JSON.stringify(saved));
  for(const next of [new Workbook(saved),new Workbook(readXlsx(writeXlsx(wb)).data)]) {
    assert.equal(next.sheets[0].cells.size,1001);assert.equal(next.getValue(0,0,0),7);
    for(const r of [1,499,500,501,1000])assert.equal(next.styleAt(0,r,0).fill,style.fill);
  }
});


test('bounded save iterator exactly matches legacy seq sort after mutations and overlapping runs',()=>{
  const [column]=pair(),other=getSharedBlankCell({italic:true});
  const sameRecords=()=>assert.deepEqual([...column.storageEntries({bounded:true})],[...column.storageEntries()]);
  sameRecords();let seed=517;
  const random=n=>((seed=(Math.imul(seed,1664525)+1013904223)>>>0)%n);
  for(let i=0;i<4000;i++) {
    const r=random(450),op=random(9);
    if(op<2)column.delete(r);
    else if(op===2)column.setRun(r,random(30)+1,other);
    else if(op===3)column.set(r,blank);
    else column.set(r,{raw:String(i)});
    if(i%40===0)sameRecords();
  }
  sameRecords();column.clear();column.setRun(10,500,blank);column.set(0,{raw:'first'});
  column.delete(10);column.set(10,blank);column.set(509,other);column.set(200,{raw:'middle'});sameRecords();
  column.set(200,blank);sameRecords();
});

test('bounded save skips expansion and points enumeration; early return releases reader guard',()=>{
  const cells=new CellMap();cells.setRunRC(0,0,1_000_000,blank);const col=cells.col(0);
  col.set(2,{raw:'note'});col.set(999997,{raw:'=1'});
  const expected=[...col.storageEntries()];
  const original=col.points[Symbol.iterator];col.points[Symbol.iterator]=()=>{throw Error('must not copy the entire points map');};
  try {assert.deepEqual([...col.storageEntries({bounded:true})],expected);const it=col.storageEntries({bounded:true});it.next();assert.equal(col.readers,1);it.return();assert.equal(col.readers,0);}
  finally{col.points[Symbol.iterator]=original;}
});

test('bounded workbook JSON chunks keep legacy byte sequence for plain and compact columns',()=>{
  const w=new Workbook({sheets:[{name:'Synthetic',cells:{'0,1':{raw:'1'},'1,1':{raw:'=A1',cached:3,comment:'note',link:'#A1'}}}]});
  w.sheets[0].cells.setRunRC(0,0,10000,{raw:'',style});
  for(let r=0;r<10000;r+=5)w.setInput(0,r,0,r%2?'=1+2':String(r));
  w.sheets[0].cells.deleteRC(127,0);w.sheets[0].cells.setRC(127,0,{raw:'reinserted'});
  const a=Array.from(w.cellRunChunks(0,211,{shareStyle:true}),JSON.stringify),b=Array.from(w.cellRunChunks(0,211,{shareStyle:true,bounded:true}),JSON.stringify);
  assert.deepEqual(b,a);
});
