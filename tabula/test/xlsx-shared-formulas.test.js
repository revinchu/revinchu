import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxBlobAsync, writeXlsxToSink } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';
import { createXlsxSharedFormulas } from '../src/xlsx-shared-formulas.js';

const formulas = bytes => descendants(parseXml(textOf(unzip(bytes)['xl/worksheets/sheet1.xml'])), 'f');
function fixture() {
  const wb = new Workbook();
  for (let r = 0; r < 10; r++) {
    wb.setInput(0, r, 0, String(r + 1));
    wb.setInput(0, r, 1, `=A${r + 1}+$A$1+A$1+$A${r + 1}`);
    wb.setInput(0, r, 2, `=IF(A${r + 1}>4,"A1<&","A2")`);
  }
  return wb;
}
for (const mode of ['sync', 'blob', 'sink']) test(`shared formulas ${mode}: mixed references and cached values survive two roundtrips`, async () => {
  const wb = fixture();
  let bytes;
  if (mode === 'sync') bytes = writeXlsx(wb);
  else if (mode === 'blob') bytes = new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
  else {
    const chunks = []; await writeXlsxToSink(wb, {}, { write: bytes => chunks.push(bytes.slice()), close() {} });
    bytes = Buffer.concat(chunks);
  }
  const fs = formulas(bytes), anchors = fs.filter(f => f.attrs.ref);
  assert.equal(anchors.length, 2); assert.equal(fs.filter(f => !f.text).length, 18);
  assert.deepEqual(anchors.map(f => f.attrs.ref), ['B1:B10', 'C1:C10']);
  for (let cycle = 0; cycle < 2; cycle++) {
    const back = new Workbook(readXlsx(bytes).data);
    for (let r = 0; r < 10; r++) for (let c = 1; c < 3; c++) {
      assert.equal(back.getCell(0, r, c).raw, wb.getCell(0, r, c).raw);
      assert.equal(back.getValue(0, r, c), wb.getValue(0, r, c));
    }
    bytes = writeXlsx(back);
  }
});
test('gaps, changed formulas and identical ASTs with different relative origins break groups', () => {
  const wb = fixture();
  wb.setInput(0, 3, 1, '0'); wb.setInput(0, 6, 1, '=A2+$A$1+A$1+$A2');
  const bytes = writeXlsx(wb), back = new Workbook(readXlsx(bytes).data);
  assert.deepEqual(formulas(bytes).filter(f => f.attrs.ref?.startsWith('B')).map(f => f.attrs.ref), ['B1:B3', 'B5:B6', 'B8:B10']);
  for (let r = 0; r < 10; r++) assert.equal(back.getCell(0,r,1).raw,wb.getCell(0,r,1).raw);
});
test('array and structured formulas keep their individual export path', () => {
  const wb = fixture();
  for (let r=0;r<10;r++) wb.getCell(0,r,1).maybeArray=true;
  const select=createXlsxSharedFormulas(wb.sheets[0].cells,1048576,16384);
  assert.equal(select(0,1,wb.getCell(0,0,1)),null);
  assert.equal(select(0,2,{formula:true,ast:{},raw:'=Table1[@Value]'}),null);
});
test('lookahead is bounded, respects Excel row limits and leaves cells unchanged', () => {
  const ast={},cell=r=>({formula:true,ast,raw:`=A${r+1}`,dr:r});let reads=0;
  const select=createXlsxSharedFormulas({getRC(r){reads++;return cell(r);}},5000,16384);
  const first=select(0,1,cell(0));assert.deepEqual(first,{r1:0,r2:4095,c:1,si:0});
  assert.equal(reads,4095);assert.equal(select(4095,1,cell(4095)),first);
  assert.deepEqual(select(4096,1,cell(4096)),{r1:4096,r2:4999,c:1,si:1});
  assert.equal(select(5000,1,cell(5000)),null);
});
test('table formulas remain explicit and preserve their formula context and values', async () => {
  const wb=new Workbook();
  wb.sheets[0].tables=[{name:'Data',r1:0,c1:0,r2:5,c2:2,header:true,columns:['Amount','Double','Mixed'],totals:true,style:'TableStyleMedium2'}];
  ['Amount','Double','Mixed'].forEach((v,c)=>wb.setInput(0,0,c,v));
  for(let r=1;r<=6;r++){
    wb.setInput(0,r,0,String(r));
    wb.setInput(0,r,1,'=Data[[#This Row],[Amount]]*2');
    wb.setInput(0,r,2,'=Data[[#This Row],[Amount]]+A2');
  }
  const bytes=new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
  const anchors=formulas(bytes).filter(f=>f.attrs.ref);
  assert.deepEqual(anchors.map(f=>f.attrs.ref),[]);
  const back=new Workbook(readXlsx(bytes).data);
  for(let r=1;r<5;r++)for(let c=1;c<=2;c++){
    assert.equal(back.getCell(0,r,c).raw,wb.getCell(0,r,c).raw);
    assert.equal(back.getValue(0,r,c),c===1?r*2:r+1);
  }
});
