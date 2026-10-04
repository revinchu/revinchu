import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { CellMap } from '../src/cellmap.js';
import { storedCellChunks, CELL_CHUNK_CHAR_LIMIT } from '../src/cell-storage.js';

function restoreChunks(chunks) {
  const cells=new CellMap();
  for(const chunk of chunks)for(const [r,c,count,data] of JSON.parse(JSON.stringify(chunk))) {
    if(count===1)cells.setRC(r,c,data);else cells.setRunRC(r,c,count,data);
  }
  return new Workbook({sheets:[{name:'복원',cells}]});
}
function assertBounded(chunks) {
  for(const chunk of chunks)assert.ok(chunk.length===1||JSON.stringify(chunk).length<=CELL_CHUNK_CHAR_LIMIT);
}

test('long text, notes, escaped characters and shared styles respect the text budget and preserve storage order',()=>{
  const cells=new CellMap(),style={font:'공유 서식'.repeat(1000),fill:'#abcdef'};
  for(let r=0;r<96;r++)cells.setRC(r%12,Math.floor(r/12),{raw:'가'.repeat(32767),comment:'\u0000'.repeat(100),style});
  const original=[...cells.storageEntries()].map(([r,c])=>[r,c]);
  const chunks=[...storedCellChunks(cells,cellData)];assert.ok(chunks.length>3);assertBounded(chunks);
  assert.deepEqual(chunks.flat().map(([r,c])=>[r,c]),original);
  const restored=restoreChunks(chunks);assert.equal(restored.sheets[0].cells.size,96);
  for(let r=0;r<12;r++)for(let c=0;c<8;c++){
    assert.equal(restored.getRaw(0,r,c),'가'.repeat(32767));
    assert.equal(restored.getCell(0,r,c).comment,'\u0000'.repeat(100));
    assert.equal(restored.styleAt(0,r,c).font,style.font);
  }
});
test('item count cap remains active and compact runs are never expanded',()=>{
  const cells=new CellMap();for(let i=0;i<18;i++)cells.setRC(i,0,{raw:String(i)});
  cells.setRunRC(100,1,1_000_000,{raw:'',style:{fill:'#123456'}});
  cells.entries=()=>{throw new Error('logical range expansion');};
  const chunks=[...storedCellChunks(cells,cellData,7)];assert.deepEqual(chunks.map(c=>c.length),[7,7,5]);
  assert.equal(chunks.flat().find(e=>e[3].raw==='')[2],1_000_000);
  const restored=restoreChunks(chunks);assert.equal(restored.sheets[0].cells.size,1_000_018);assert.equal(restored.styleAt(0,1_000_099,1).fill,'#123456');
});
test('one oversized cell is retained as one chunk between ordinary cells',()=>{
  const cells=new CellMap(),large='🙂'.repeat(CELL_CHUNK_CHAR_LIMIT);
  cells.setRC(0,0,{raw:'앞'});cells.setRC(1,0,{raw:large,comment:'큰 셀'});cells.setRC(2,0,{raw:'뒤'});
  const chunks=[...storedCellChunks(cells,cellData)];assert.deepEqual(chunks.map(c=>c.length),[1,1,1]);
  assert.equal(restoreChunks(chunks).getRaw(0,1,0),large);
});
test('cancellation closes the source iterator without reading further cells or changing the previous snapshot',async()=>{
  const old=new Workbook({sheets:[{name:'이전',cells:{'0,0':{raw:'이전 값'}}}]});
  const snapshot=old.serializeBlob(),large='x'.repeat(CELL_CHUNK_CHAR_LIMIT+1);let visited=0,closed=false;
  const cells={*storageEntries(){try{for(let r=0;r<100;r++){visited++;yield[r,0,{raw:large},1];}}finally{closed=true;}}};
  const steps=storedCellChunks(cells,cellData);assert.equal(steps.next().value.length,1);steps.return();
  assert.equal(visited,1);assert.equal(closed,true);assert.equal(steps.next().done,true);
  assert.equal(new Workbook(JSON.parse(await snapshot.text())).getRaw(0,0,0),'이전 값');
});
test('chunk size estimation neither stringifies each cell nor touches ASTs, and memoizes repeated styles',()=>{
  const cells=new CellMap();let styleReads=0;
  const style={get font(){styleReads++;return '공유'.repeat(2000);}};
  for(let r=0;r<40;r++)cells.setRC(r,0,{raw:'문자',style,get ast(){throw new Error('AST read');}});
  const stringify=JSON.stringify;let chunks;
  try{JSON.stringify=()=>{throw new Error('double serialization');};chunks=[...storedCellChunks(cells,cellData)];}finally{JSON.stringify=stringify;}
  assert.ok(styleReads<=45);assertBounded(chunks);assert.equal(chunks.flat().length,40);
});
test('Workbook chunks survive Blob/gzip round trips with formulas, comments, links and compact ranges',async()=>{
  const w=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'=UNSUPPORTED(1)',cached:42},'1,0':{raw:'원본',comment:'메모',link:'#A1'}},cellRuns:[[100,2,100000,{raw:'',style:{bold:true}}]]}]});
  for(let r=2;r<70;r++)w.setInput(0,r,0,'본문'.repeat(10000));
  const chunks=[...w.cellRunChunks(0)];assertBounded(chunks);assert.ok(chunks.length>1);
  const back=[];
  for(const chunk of chunks){const blob=new Blob([JSON.stringify(chunk)]);const packed=await new Response(blob.stream().pipeThrough(new CompressionStream('gzip'))).blob();back.push(JSON.parse(await new Response(packed.stream().pipeThrough(new DecompressionStream('gzip'))).text()));}
  const restored=restoreChunks(back),blobRestored=new Workbook(JSON.parse(await restored.serializeBlob().text()));
  assert.equal(blobRestored.getRaw(0,0,0),'=UNSUPPORTED(1)');assert.equal(blobRestored.getCell(0,0,0).cached,42);
  assert.equal(blobRestored.getCell(0,1,0).comment,'메모');assert.equal(blobRestored.getCell(0,1,0).link,'#A1');
  assert.equal(blobRestored.getRaw(0,69,0),'본문'.repeat(10000));assert.equal(blobRestored.styleAt(0,100099,2).bold,true);
});


test('bounded chunk string sizing avoids scanning source text and remains conservative',async()=>{
  const {createJsonSizer}=await import('../src/json-size.js');
  const measure=createJsonSizer(100000,{conservativeStrings:true}).size;
  const samples=['plain','a\"b\\c\n\u0000','가😀\ud800'.repeat(1000)];
  const testPattern=RegExp.prototype.test;let scans=0;
  try {RegExp.prototype.test=function(...args){scans++;return testPattern.apply(this,args);};for(const text of samples)assert.ok(measure(text)>=JSON.stringify(text).length);}
  finally {RegExp.prototype.test=testPattern;}
  assert.equal(scans,0);
  assert.equal(createJsonSizer(10,{conservativeStrings:true}).size('abc'),11);
  assert.equal(createJsonSizer(100000).size('plain'),JSON.stringify('plain').length);
});
