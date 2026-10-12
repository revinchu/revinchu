import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { convertXlsb } from '../src/xlsb.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, child, descendants } from '../src/xml.js';
const u32=(...xs)=>{const b=Buffer.alloc(xs.length*4);xs.forEach((x,i)=>b.writeUInt32LE(x>>>0,i*4));return b;};
const wide=s=>Buffer.concat([u32(s.length),Buffer.from(s,'utf16le')]);
const variable=x=>{const a=[];do{const b=x&127;x>>>=7;a.push(b|(x?128:0));}while(x);return Buffer.from(a);};
const rec=(id,b=Buffer.alloc(0))=>Buffer.concat([variable(id),variable(b.length),b]);
const color=rgb=>rec(564,Buffer.from([5,255,0,0,...rgb,255]));
const vo=(type,v=0)=>{const b=Buffer.alloc(24);b.writeUInt32LE(type);b.writeDoubleLE(v,4);return rec(471,b);};
function converted(scope=2,columns=false) {
 const head=Buffer.alloc(32);head.writeUInt32LE(6,0);head.writeUInt32LE(0x80000|0x80,4);head.writeUInt32LE(0xc0,8);head.writeUInt32LE(1,28);head[12]=2;
 const pt=[rec(280,Buffer.concat([head,wide('Report'),wide('Values')]))];
 for(let f=0;f<4;f++){pt.push(rec(285,u32(f===0?1:f===1&&columns?2:f>=2?8:0,0,0x100,0,0)),rec(286));}
 pt.push(rec(309,u32(1,0)));if(columns)pt.push(rec(311,u32(2,1,-2)));
 for(let f=2;f<4;f++)pt.push(rec(293,Buffer.concat([u32(f,0,0,0,0,0),Buffer.from([1]),wide(f===2?'Sales':'ROAS')])));
 pt.push(rec(314,u32(132,140,6,9,132,133,7,0,0)));
 for(const priority of [7,2]) {
  pt.push(rec(558,u32(scope,0,priority)),rec(641,u32(1)),rec(247,Buffer.concat([u32(-1),Buffer.from([1,129,0,0])])),rec(249,u32(2)));
  pt.push(rec(251,Buffer.concat([u32(-2,1),Buffer.from([1,0,0])])),rec(382,u32(1)),rec(383),rec(252));
  pt.push(rec(251,Buffer.concat([u32(columns?1:0,0),Buffer.from([1,0,0])])),rec(252),rec(250),rec(248),rec(642),rec(559));
 }
 const sheet=[];
 for(const [r,priority] of [[165,7],[133,2]]) {
  const rule=Buffer.alloc(46);rule.writeInt32LE(3);rule.writeInt32LE(2,4);rule.writeInt32LE(-1,8);rule.writeInt32LE(priority,12);rule.writeInt32LE(-1,42);
  sheet.push(rec(461,u32(1,1,1,r,r+6,9,9)),rec(463,rule),rec(469),vo(2),vo(5,50),vo(3),color([99,190,123]),color([255,235,132]),color([248,105,107]),rec(470),rec(464),rec(462));
 }
 const files={'xl/workbook.bin':rec(156,Buffer.concat([u32(0,1),wide('r1'),wide('Sheet1')])),
  'xl/_rels/workbook.bin.rels':Buffer.from('<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>'),
  'xl/pivotTables/pivotTable1.bin':Buffer.concat(pt),'xl/worksheets/sheet1.bin':Buffer.concat(sheet)};
 for(const _ of convertXlsb(files)){}
 return {sheet:textOf(files['xl/worksheets/sheet1.bin']),pivot:textOf(files['xl/pivotTables/pivotTable1.bin'])};
}
function fixture(scope,columns) {
 const wb=new Workbook();
 [['Region','Channel','Amount','Rate'],['A','X',10,1],['B','Y',20,2]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
 wb.sheets[0].pivot={name:'Report',source:'Sheet1',range:{r1:0,c1:0,r2:2,c2:3},rows:['Region'],cols:columns?['Channel']:[],values:[{field:'Amount',name:'Sales'},{field:'Rate',name:'ROAS'}],top:132,left:6};
 const files=unzip(writeXlsx(wb)),convertedParts=converted(scope,columns);
 files['xl/pivotTables/pivotTable1.xml']=convertedParts.pivot;
 files['xl/worksheets/sheet1.xml']=textOf(files['xl/worksheets/sheet1.xml']).replace('</worksheet>',convertedParts.sheet.match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g).join('')+'</worksheet>');
 return zip(files);
}
test('XLSB preserves pivot marker, priorities, rule areas and value-field reference',()=>{
 const {sheet,pivot}=converted();
 const wsRules=descendants(parseXml(sheet),'conditionalFormatting');assert.equal(wsRules.length,2);assert.equal(wsRules.every(x=>x.attrs.pivot==='1'),true);
 const cf=descendants(parseXml(pivot),'conditionalFormat');
 assert.deepEqual(cf.map(x=>x.attrs.priority),['7','2']);
 assert.equal(cf[0].attrs.scope,'field');
 const refs=descendants(cf[0],'reference');assert.equal(refs[0].attrs.field,'4294967294');assert.equal(child(refs[0],'x').attrs.v,'1');
 assert.equal(refs[1].attrs.field,'0');assert.equal(descendants(parseXml(pivot),'autoSortScope').length,0);
});
for(const [scope,columns] of [[0,false],[1,false],[2,false],[2,true]])test(`XLSB pivot CF scope ${scope}, columns ${columns} survives XLSX roundtrip`,async()=>{
 const bytes=fixture(scope,columns),res=await readXlsxAsync(new Blob([bytes]));
 const wb=new Workbook(res.data),rules=wb.sheets[0].cond;
 assert.equal(rules.length,2);assert.deepEqual(rules.map(r=>r.r1),[133,165]);
 assert.deepEqual(rules[0].colors,['#63be7b','#ffeb84','#f8696b']);
 assert.deepEqual(rules[0].pivot,{name:'Report',scope:['selection','data','field'][scope],value:'ROAS',...(columns?{colField:'Channel'}:{rowField:'Region'})});
 // Data scope is field-wide; row/column intersections are only meaningful in field/selection scope.
 if(scope===1)delete rules[0].pivot.rowField,delete rules[1].pivot.rowField;
 const out=writeXlsx(wb),back=readXlsx(out).data.sheets[0].cond;
 assert.deepEqual(back,rules);
 assert.equal(descendants(parseXml(textOf(unzip(out)['xl/worksheets/sheet1.xml'])),'conditionalFormatting').every(x=>x.attrs.pivot==='1'),true);
});
