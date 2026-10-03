import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pictureExportBounds, preparePictureExport } from '../src/picture-export.js';
import { Workbook } from '../src/workbook.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';
import { pivotSourceData } from '../src/pivot.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
test('picture export bounds preserve rotation and both directions of reflection', () => {
  const p = { w:320, h:160, rot:90 };
  const b = pictureExportBounds(p); assert.ok(Math.abs(b.w-160)<1e-9); assert.ok(Math.abs(b.h-320)<1e-9);
  const r = { w:320,h:160,reflection:{size:.5,gap:10,opacity:.5} };
  assert.deepEqual(pictureExportBounds(r), { x:0,y:0,w:320,h:250 });
  assert.deepEqual(pictureExportBounds({...r,flipV:true}), { x:0,y:-90,w:320,h:250 });
  const soft = pictureExportBounds({w:320,h:160,artistic:{type:'blur',amount:1}}); assert.ok(soft.w>350); assert.ok(soft.h>190);
});
test('picture export copy never mutates document objects, cells or undo', async () => {
  const book = new Workbook({sheets:[{name:'합성',cells:{'0,0':{raw:'123'}},images:[{id:'im',w:10,h:20,x:0,y:0,src:'data:image/png;base64,AA==',effectPng:'stale'}],shapes:[{id:'group',kind:'group',groupItems:[{kind:'picture',id:'child',w:10,h:20,src:'data:image/png;base64,AA==',effectPng:'stale'}]}]}]});
  const before = JSON.stringify(book.serialize()), undo = book.undoStack.length;
  const copy = await preparePictureExport(book);
  assert.equal(JSON.stringify(book.serialize()),before); assert.equal(book.undoStack.length,undo);
  assert.equal(copy.sheets[0].cells,book.sheets[0].cells); assert.equal(copy.getValue(0,0,0),123);
  assert.equal(copy.sheets[0].images[0].effectPng,undefined); assert.equal(copy.sheets[0].shapes[0].groupItems[0].effectPng,undefined);
  assert.equal(book.sheets[0].images[0].effectPng,'stale'); assert.notEqual(copy.sheets[0].images[0],book.sheets[0].images[0]);
});

test('picture export refuses a mixed snapshot when the workbook changes during async preparation', async () => {
  const book = new Workbook({sheets:[{name:'합성',cells:{},images:[{id:'im',w:10,h:20,src:'data:image/png;base64,AA=='}]}]});
  const exporting = preparePictureExport(book);
  book.transact(() => book.setCellData(0,0,0,{raw:'changed'}));
  await assert.rejects(exporting, /저장 중 문서가 변경/);
  assert.equal(book.getCell(0,0,0).raw,'changed');
});

function savedPivotBook(compact, table = false) {
  let rows = [['분류','값'],['가',10]];
  if (compact) { const b = new PivotSnapshotBuilder(rows[0]); b.add(rows[1]); rows=b.finish(); }
  const source = table ? {table:'원본표'} : {source:'원본',range:{r1:0,c1:0,r2:1,c2:1}};
  return new Workbook({pivotSnapshots:{saved:rows},sheets:[
    {name:'원본',cells:{'0,0':{raw:'분류'},'0,1':{raw:'값'},'1,0':{raw:'가'},'1,1':{raw:'999'}},...(table?{tables:[{id:'t',name:'원본표',r1:0,c1:0,r2:1,c2:1,header:true,totals:false}]}:{})},
    {name:'보고서',cells:{},pivot:{name:'저장피벗',...source,snapshotId:'saved',rows:['분류'],values:[{field:'값',agg:'sum'}],top:0,left:0}}
  ]});
}
for (const compact of [false,true]) for (const table of [false,true]) test(`picture export retains saved pivot rows without evicting live cache, compact=${compact}, table=${table}`, async () => {
  const book=savedPivotBook(compact,table),original=book.pivotSnapshots.get('saved'),version=book.version;
  const copy=await preparePictureExport(book);
  assert.notEqual(copy.pivotSnapshots,book.pivotSnapshots);
  assert.notEqual(copy.pivotSnapshots.get('saved'),original);
  assert.equal(copy.pivotSnapshots.get('saved').rows,original.rows);
  assert.equal(copy.pivotSnapshots.get('saved').sourceSheet,copy.sheets[0]);
  assert.equal(original.sourceSheet,book.sheets[0]);
  assert.equal(pivotSourceData(copy,copy.sheets[1].pivot).rows[1][1],10);
  const bytes=writeXlsx(copy),parts=unzip(bytes),reopened=new Workbook(readXlsx(bytes).data);
  assert.match(textOf(parts['xl/pivotCache/pivotCacheDefinition1.xml']),/saveData="1"/);
  assert.ok(parts['xl/pivotCache/pivotCacheRecords1.xml']);
  assert.equal(pivotSourceData(reopened,reopened.sheets[1].pivot).rows[1][1],10);
  assert.equal(reopened.getValue(0,1,1),999);
  assert.equal(book.pivotSnapshots.get('saved'),original);
  assert.equal(pivotSourceData(book,book.sheets[1].pivot).rows[1][1],10);
  assert.equal(book.version,version);assert.equal(book.undoStack.length,0);
});
test('picture export does not revive a saved pivot cache after a real source edit', async () => {
  const book=savedPivotBook(true,true);
  book.transact(()=>book.setInput(0,1,1,'123'));
  const original=book.pivotSnapshots.get('saved'),copy=await preparePictureExport(book);
  assert.equal(copy.pivotSnapshots.size,0);
  assert.equal(pivotSourceData(copy,copy.sheets[1].pivot).rows[1][1],123);
  assert.equal(book.pivotSnapshots.get('saved'),original);
  assert.equal(book.undoStack.length,1);
});