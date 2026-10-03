import test from 'node:test';
import assert from 'node:assert/strict';
import { decodeFormula } from '../src/xlsb.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { zip } from '../src/zip.js';
const u32=(...xs)=>{const b=Buffer.alloc(xs.length*4);xs.forEach((x,i)=>b.writeUInt32LE(x>>>0,i*4));return b;};
const wide=s=>Buffer.concat([u32(s.length),Buffer.from(s,'utf16le')]);
const variable=x=>{const a=[];do{const b=x&127;x>>>=7;a.push(b|(x?128:0));}while(x);return Buffer.from(a);};
const rec=(id,b=Buffer.alloc(0))=>Buffer.concat([variable(id),variable(b.length),b]);
const formula=bytes=>Buffer.concat([u32(bytes.length),Buffer.from(bytes),u32(0)]);
// MS-XLSB 2.5.98.2 BErr: 0x2B is GETTING_DATA, not SPILL.
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/c3d25119-1da4-44dc-bdb5-19b8ba2ddf90
function fixture(){return zip({
 '_rels/.rels':'<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>',
 'xl/workbook.bin':rec(156,Buffer.concat([u32(0,1),wide('r1'),wide('Sheet1')])),
 'xl/_rels/workbook.bin.rels':'<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>',
 'xl/worksheets/sheet1.bin':Buffer.concat([rec(129),rec(148,u32(0,0,0,1)),rec(145),rec(0,Buffer.concat([u32(0,0),Buffer.from([44,1,0,0])])),rec(3,Buffer.concat([u32(0,0),Buffer.from([43])])),rec(11,Buffer.concat([u32(1,0),Buffer.from([43,0,0]),formula([30,1,0])])),rec(146),rec(130)])
});}
test('XLSB PtgErr retains the defined GETTING_DATA error',()=>assert.equal(decodeFormula(formula([28,43]),0,{},{r:0,c:0}).text,'#GETTING_DATA'));
for(const asyncRead of [false,true])test(`XLSB ${asyncRead?'async':'sync'} literals and cached formula errors preserve GETTING_DATA through XLSX`,async()=>{
 const data=asyncRead?(await readXlsxAsync(fixture())).data:readXlsx(fixture()).data,w=new Workbook(data);
 assert.equal(w.getValue(0,0,0).code,'#GETTING_DATA');assert.equal(w.getValue(0,0,1).code,'#GETTING_DATA');
 const back=new Workbook(readXlsx(writeXlsx(w)).data);
 assert.equal(back.getValue(0,0,0).code,'#GETTING_DATA');assert.equal(back.getValue(0,0,1).code,'#GETTING_DATA');
});
test('GETTING_DATA stays an error literal and formula, with Excel ERROR.TYPE 8',()=>{
 const w=new Workbook();w.setInput(0,0,0,'#GETTING_DATA');w.setInput(0,0,1,'=#GETTING_DATA');w.setInput(0,0,2,'=ERROR.TYPE(A1)');w.setInput(0,0,3,'=ISERROR(A1)');w.setInput(0,0,4,'#SPILL!');
 assert.equal(w.getValue(0,0,0).code,'#GETTING_DATA');assert.equal(w.getValue(0,0,1).code,'#GETTING_DATA');assert.equal(w.getValue(0,0,2),8);assert.equal(w.getValue(0,0,3),true);assert.equal(w.getValue(0,0,4).code,'#SPILL!');
 const back=new Workbook(w.serialize());assert.equal(back.getValue(0,0,0).code,'#GETTING_DATA');assert.equal(back.getValue(0,0,1).code,'#GETTING_DATA');
});
