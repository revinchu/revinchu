import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pictureExportBounds, preparePictureExport } from '../src/picture-export.js';
import { Workbook } from '../src/workbook.js';
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
