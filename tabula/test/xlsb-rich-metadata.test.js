import test from 'node:test';
import assert from 'node:assert/strict';
import { xlsbRichMetadataXml } from '../src/xlsb-rich-metadata.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
const u32=(...xs)=>{const b=Buffer.alloc(xs.length*4);xs.forEach((x,i)=>b.writeUInt32LE(x>>>0,i*4));return b;};
const wide=s=>Buffer.concat([u32(s.length),Buffer.from(s,'utf16le')]);
const variable=x=>{const a=[];do{const b=x&127;x>>>=7;a.push(b|(x?128:0));}while(x);return Buffer.from(a);};
const rec=(id,b=Buffer.alloc(0))=>Buffer.concat([variable(id),variable(b.length),b]);
function metadata({endPayload=false,blocks=[1,0,null],refs=[[2,1],[2,0],[1,99],[2,2]],pairs}={}) {
 const chunks=[rec(332),rec(334,u32(2)),rec(335,Buffer.concat([u32(0,120000),wide('UNSUPPORTED')])),rec(335,Buffer.concat([u32(0,120000),wide('XLRICHVALUE')])),rec(336),rec(339,Buffer.concat([u32(1),wide('UNSUPPORTED')])),rec(52),rec(9999,u32(123)),rec(53),rec(340),rec(339,Buffer.concat([u32(blocks.length),wide('XLRICHVALUE')]))];
 for(const i of blocks){chunks.push(rec(52));if(i!==null)chunks.push(...(pairs??[rec(5002,endPayload?Buffer.alloc(0):u32(0,i)),rec(5003,endPayload?u32(0,i):Buffer.alloc(0))]));chunks.push(rec(53));}
 chunks.push(rec(340),rec(337,u32(1,1)),rec(51,u32(1,1,0)),rec(338),rec(337,u32(refs.length,0)));
 for(const ref of refs)chunks.push(rec(51,u32(ref.length/2,...ref)));
 chunks.push(rec(338),rec(333));return Buffer.concat(chunks);
}
const xml=opts=>parseXml(xlsbRichMetadataXml(metadata(opts)));
for(const endPayload of [false,true])test(`XLSB metadata maps rich values by indexes, payload on ${endPayload?'native End':'specified Begin'}`,()=>{
 const tree=xml({endPayload}),future=child(tree,'futureMetadata'),values=child(tree,'valueMetadata');
 assert.deepEqual(kids(future,'bk').map(b=>descendants(b,'rvb')[0]?.attrs.i??null),['1','0',null]);
 assert.deepEqual(kids(values,'bk').map(b=>child(b,'rc')?.attrs??null),[{t:'2',v:'1'},{t:'2',v:'0'},null,{t:'2',v:'2'}]);
 assert.equal(kids(child(tree,'metadataTypes'),'metadataType')[0].attrs.name,'UNSUPPORTED');
});
test('unknown future metadata and missing picture blocks preserve fallback error values',()=>{
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
 const w=new Workbook();for(let r=0;r<2;r++)w.setCellData(0,r,0,{raw:'',image:{src:png,alt:r?'second':'first'}});
 w.setInput(0,2,0,'#VALUE!');w.setInput(0,3,0,'#N/A');
 const files=unzip(writeXlsx(w));files['xl/metadata.xml']=xlsbRichMetadataXml(metadata());
 let sheet=textOf(files['xl/worksheets/sheet1.xml']);sheet=sheet.replace('<c r="A3"','<c vm="3" r="A3"').replace('<c r="A4"','<c vm="4" r="A4"');files['xl/worksheets/sheet1.xml']=sheet;
 const back=new Workbook(readXlsx(zip(files)).data);
 assert.equal(back.getCell(0,0,0).image.alt,'first');assert.equal(back.getCell(0,1,0).image.alt,'second');
 assert.equal(back.getValue(0,2,0).code,'#VALUE!');assert.equal(back.getValue(0,3,0).code,'#N/A');
 assert.equal(back.getCell(0,2,0).image,undefined);assert.equal(back.getCell(0,3,0).image,undefined);
 const twice=new Workbook(readXlsx(writeXlsx(back)).data);assert.deepEqual(twice.getCell(0,0,0).image,back.getCell(0,0,0).image);
});
for(const [name,pairs] of [
 ['duplicate payload',[rec(5002,u32(0,0)),rec(5003,u32(0,1))]],
 ['duplicate pair',[rec(5002,u32(0,0)),rec(5003),rec(5002,u32(0,1)),rec(5003)]],
 ['truncated payload',[rec(5002,Buffer.alloc(7)),rec(5003)]],
 ['nonblank FRT header',[rec(5002,u32(1,0)),rec(5003)]],
 ['missing start',[rec(5003,u32(0,0))]],
])test(`invalid rich metadata cannot select an unrelated image: ${name}`,()=>assert.throws(()=>xlsbRichMetadataXml(metadata({blocks:[0],refs:[[2,0]],pairs})),/메타데이터/));
test('out-of-range value metadata fails explicitly',()=>assert.throws(()=>xml({refs:[[2,999]]}),/메타데이터/));
test('truncated binary metadata cannot read past record bounds',()=>assert.throws(()=>xlsbRichMetadataXml(metadata().subarray(0,-1)),/메타데이터/));

function binaryFixture() {
 const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
 const w=new Workbook();for(let r=0;r<2;r++)w.setCellData(0,r,0,{raw:'',image:{src:png,alt:r?'second':'first'}});
 const files=unzip(writeXlsx(w));
 files['xl/workbook.bin']=rec(156,Buffer.concat([u32(0,1),wide('rId1'),wide('Sheet1')]));delete files['xl/workbook.xml'];
 files['_rels/.rels']=textOf(files['_rels/.rels']).replaceAll('workbook.xml','workbook.bin');
 files['xl/_rels/workbook.bin.rels']=textOf(files['xl/_rels/workbook.xml.rels']).replaceAll('sheet1.xml','sheet1.bin').replaceAll('metadata.xml','metadata.bin');delete files['xl/_rels/workbook.xml.rels'];
 delete files['xl/styles.xml'];delete files['xl/sharedStrings.xml'];
 files['xl/metadata.bin']=metadata({endPayload:true});delete files['xl/metadata.xml'];
 const error=c=>rec(3,Buffer.concat([u32(c,0),Buffer.from([15])]));
 const row=r=>rec(0,Buffer.concat([u32(r,0),Buffer.from([44,1,0,0])]));
 files['xl/worksheets/sheet1.bin']=Buffer.concat([rec(129),rec(148,u32(0,1,0,6)),rec(145),row(0),rec(50,u32(1)),error(0),rec(14,Buffer.concat([u32(0),Buffer.from([15])])),rec(50,u32(2)),error(2),rec(50,u32(3)),error(3),rec(50,u32(4)),error(4),rec(50,u32(1)),rec(50,u32(0)),error(5),rec(50,u32(1)),rec(49,u32(1)),error(6),rec(50,u32(2)),row(1),error(0),rec(146),rec(130)]);
 delete files['xl/worksheets/sheet1.xml'];return zip(files);
}
for(const asyncRead of [false,true])test(`complete XLSB ${asyncRead?'async':'sync'} import maps only the intended picture cells and round-trips`,async()=>{
 const data=asyncRead?(await readXlsxAsync(binaryFixture())).data:readXlsx(binaryFixture()).data;
 const w=new Workbook(data);
 for(const [c,alt] of [[0,'first'],[2,'second'],[6,'first']]){assert.equal(w.getCell(0,0,c).image.alt,alt);assert.equal(w.getValue(0,0,c).type,'image');}
 for(const [r,c] of [[0,1],[0,3],[0,4],[0,5],[1,0]]){assert.equal(w.getCell(0,r,c).image,undefined);assert.equal(w.getValue(0,r,c).code,'#VALUE!');}
 const back=new Workbook(readXlsx(writeXlsx(w)).data);
 for(const c of [0,2,6])assert.deepEqual(back.getCell(0,0,c).image,w.getCell(0,0,c).image);
 for(const [r,c] of [[0,1],[0,3],[0,4],[0,5],[1,0]])assert.equal(back.getValue(0,r,c).code,'#VALUE!');
});
