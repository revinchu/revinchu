import test from 'node:test';
import assert from 'node:assert/strict';
import { convertXlsb } from '../src/xlsb.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { computePivot, pivotSourceData, resolvePivot } from '../src/pivot.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child } from '../src/xml.js';
const u32=(...xs)=>{const b=Buffer.alloc(xs.length*4);xs.forEach((x,i)=>b.writeUInt32LE(x>>>0,i*4));return b;};
const wide=s=>Buffer.concat([u32(s.length),Buffer.from(s,'utf16le')]);
const variable=x=>{const a=[];do{const b=x&127;x>>>=7;a.push(b|(x?128:0));}while(x);return Buffer.from(a);};
const rec=(id,b)=>Buffer.concat([variable(id),variable(b.length),b]);
function binaryDefinition({showValuesRow,classic,multi,repeatLabels,twoRows}) {
 const fixed=Buffer.alloc(32);fixed.writeUInt32LE(6,0);fixed.writeUInt32LE(0x80000|0x80,4);fixed.writeUInt32LE(0xc0|(classic?0:0x10),8);fixed[12]=2;fixed.writeInt32LE(-1,16);
 const chunks=[rec(280,Buffer.concat([fixed,wide('Display'),wide('Values')]))];
 for(let f=0;f<4;f++){chunks.push(rec(285,u32(f===0||twoRows&&f===1?1:f>=2?8:0,0,0x100,0,0)));if(repeatLabels!==undefined)chunks.push(rec(1061,u32(0,repeatLabels?1:0,0)));chunks.push(rec(286,Buffer.alloc(0)));}
 chunks.push(rec(309,twoRows?u32(2,0,1):u32(1,0)),rec(311,u32(1,-2)));
 for(const f of (multi?[2,3]:[2]))chunks.push(rec(293,Buffer.concat([u32(f,0,0,0,0,0),Buffer.from([1]),wide(f===2?'Revenue':'Cost')])));
 chunks.push(rec(314,u32(2,6,6,multi?8:7,2,3,7,0,0)));
 if(showValuesRow!==undefined)chunks.push(rec(1062,Buffer.concat([u32(0),Buffer.from([showValuesRow?0x10:0]),u32(0,-1,-1,-1)])));
 const files={'xl/workbook.bin':new Uint8Array(),'xl/pivotTables/pivotTable1.bin':Buffer.concat(chunks)};
 for(const _ of convertXlsb(files)){}
 return textOf(files['xl/pivotTables/pivotTable1.bin']);
}
function fixture(options) {
 const w=new Workbook();[['Region','Item','Revenue','Cost'],['A','x',10,1],[options.twoRows?'A':'B','y',20,2]].forEach((row,r)=>row.forEach((v,c)=>w.setInput(0,r,c,String(v))));
 w.sheets[0].pivot={name:'Display',source:'Sheet1',range:{r1:0,c1:0,r2:2,c2:3},rows:options.twoRows?['Region','Item']:['Region'],cols:[],values:(options.multi?['Revenue','Cost']:['Revenue']).map(field=>({field,name:field})),layout:'tabular',top:2,left:6,area:{r1:2,c1:6,r2:6,c2:8},grandRows:false,grandCols:false,...options};
 const files=unzip(writeXlsx(w));files['xl/pivotTables/pivotTable1.xml']=binaryDefinition(options);
 return new Workbook(readXlsx(zip(files)).data);
}
const render=w=>{const d=w.sheets[0].pivot,r=resolvePivot(pivotSourceData(w,d),d);return computePivot(r,r.def);};
for(const showValuesRow of [false,true])for(const classic of [false,true])for(const multi of [false,true])test(`XLSB pivot values row ${showValuesRow}, classic ${classic}, multiple values ${multi}`,()=>{
 const w=fixture({showValuesRow,classic,multi}),d=w.sheets[0].pivot,b=render(w),headers=multi&&(classic||showValuesRow)?2:1;
 assert.equal(d.showValuesRow,showValuesRow);assert.equal(d.classic,classic);assert.equal(b.meta.headerRows,headers);
 assert.equal(d.top,2);assert.equal(d.left,6);
 assert.deepEqual(b.grid[headers-1].map(c=>c.raw),multi?['Region','Revenue','Cost']:['Region','Revenue']);
 assert.equal(b.grid[headers][1].raw,'10','first data stays at expected worksheet row');
 if(!showValuesRow&&!classic)assert.equal(b.grid[0][1].raw,'Revenue','H3 caption must not be replaced with an inserted Values row');
 const back=new Workbook(readXlsx(writeXlsx(w)).data);assert.equal(back.sheets[0].pivot.showValuesRow,showValuesRow);assert.equal(back.sheets[0].pivot.classic,classic);
 assert.deepEqual(render(back).grid.map(r=>r.map(c=>c.raw)),b.grid.map(r=>r.map(c=>c.raw)));
});
test('XLSB absent 2010 display extension preserves OOXML default',()=>{const x=parseXml(binaryDefinition({multi:true,classic:false}));assert.equal(child(x,'extLst'),null);assert.equal(fixture({multi:true,classic:false}).sheets[0].pivot.showValuesRow,true);});

for(const repeatLabels of [false,true])test(`XLSB repeated outer row labels ${repeatLabels} retain coordinates and XLSX export`,()=>{
 const w=fixture({showValuesRow:false,classic:false,multi:true,twoRows:true,repeatLabels}),b=render(w);
 assert.equal(!!w.sheets[0].pivot.repeatLabels,repeatLabels);
 assert.equal(b.grid[1][0].raw,'A');assert.equal(b.grid[2][0].raw,repeatLabels?'A':'');
 assert.equal(b.grid[2][1].raw,'y');assert.equal(b.grid[2][2].raw,'20');
 const back=new Workbook(readXlsx(writeXlsx(w)).data);assert.equal(!!back.sheets[0].pivot.repeatLabels,repeatLabels);
 assert.deepEqual(render(back).grid.map(r=>r.map(c=>c.raw)),b.grid.map(r=>r.map(c=>c.raw)));
});
