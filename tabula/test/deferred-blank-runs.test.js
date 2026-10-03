import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx,readXlsxAsync,writeXlsx } from '../src/xlsx.js';
import { unzip,zip,textOf } from '../src/zip.js';
function fixture(rows) {
 const files=unzip(writeXlsx(new Workbook({sheets:[{name:'Data',cells:{'0,0':{raw:'head',style:{fill:'#ffeeaa'}},'0,1':{raw:'other',style:{fill:'#aaffee'}}}}]})));
 const xml=textOf(files['xl/worksheets/sheet1.xml']);
 const first=/<c\b[^>]*r="A1"[^>]*s="(\d+)"/.exec(xml)?.[1],other=/<c\b[^>]*r="B1"[^>]*s="(\d+)"/.exec(xml)?.[1];
 assert.ok(first&&other);
 files['xl/worksheets/sheet1.xml']='<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C100001"/><sheetData>'+rows(first,other)+'</sheetData></worksheet>';
 return zip(files);
}
test('deferred styled blanks preserve all tail coordinates while retaining block overrides',async()=>{
 const bytes=fixture((s,t)=>{
  const rows=['<row r="1"><c r="A1" s="'+s+'" t="inlineStr"><is><t>head</t></is></c></row>'];
  for(let r=2;r<=10000;r++)rows.push(`<row r="${r}"><c r="A${r}" s="${s}">${r===2||r===5000?'<v>7</v>':''}</c><c r="B${r}" s="${s}"/><c r="C${r}" s="${r===500?t:s}">${r===500?'<v>123</v>':''}</c></row>`);
  return rows.join('');
 });
 for(const data of [readXlsx(bytes).data,(await readXlsxAsync(bytes)).data]){
  const wb=new Workbook(data),sheet=wb.sheets[0];
  assert.equal(wb.getValue(0,1,0),7);assert.equal(wb.getValue(0,4999,0),7);assert.equal(wb.getValue(0,499,2),123);
  assert.equal(sheet.cells.size,15002);
  assert.ok([...sheet.cells.storageEntries()].length<12);
  for(const r of [5000,5001,9999])for(let c=0;c<3;c++){assert.equal(wb.getValue(0,r,c),null);assert.equal(wb.styleAt(0,r,c).fill,'#ffeeaa');assert.ok(sheet.cells.hasRC(r,c));}
 }
});
test('out-of-order deferred blank rows never overwrite explicit cells or lose coordinates',()=>{
 const bytes=fixture((s,t)=>`<row r="1"><c r="A1" s="${s}" t="inlineStr"><is><t>head</t></is></c></row><row r="10"><c r="A10" s="${s}"/></row><row r="2"><c r="A2" s="${s}"><v>7</v></c></row><row r="8"><c r="A8" s="${t}"/></row><row r="10"><c r="A10" s="${t}"><v>99</v></c></row><row r="7"><c r="A7" s="${s}"/></row>`);
 const wb=new Workbook(readXlsx(bytes).data);
 assert.equal(wb.getValue(0,9,0),99);assert.equal(wb.styleAt(0,9,0).fill,'#aaffee');
 assert.equal(wb.getValue(0,6,0),null);assert.equal(wb.styleAt(0,6,0).fill,'#ffeeaa');
 assert.equal(wb.styleAt(0,7,0).fill,'#aaffee');
});
