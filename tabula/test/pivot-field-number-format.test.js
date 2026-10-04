import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx, writeXlsxBlobAsync } from '../src/xlsx.js';
import { writeWixelFile, readWixelFile } from '../src/wixel-file.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';
const table=n=>`xl/pivotTables/pivotTable${n}.xml`;
const code='yyyy-mm-dd';
function source({overrideBuiltin=false}={}) {
 const wb=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Date'},'0,1':{raw:'Amount'},'0,2':{raw:'Label'},'0,3':{raw:'Unformatted'},'1,0':{raw:'45352'},'1,1':{raw:'10'},'1,2':{raw:'A'},'1,3':{raw:'x'}}},{name:'Report',cells:{}}]});
 const def={source:'Source',range:{r1:0,c1:0,r2:1,c2:3},rows:['Date'],values:[{field:'Amount',agg:'sum'}],layout:'tabular'};
 wb.sheets[1].pivot={...def,name:'Built-in',top:0};wb.sheets[1].pivotsExtra=[{...def,name:'Custom',top:8}];
 const files=unzip(writeXlsx(wb));
 assert.equal(Object.keys(files).filter(p=>/pivotCacheDefinition\d+\.xml$/.test(p)).length,1);
 files['xl/styles.xml']=textOf(files['xl/styles.xml']).replace('<fonts ',`<numFmts count="${overrideBuiltin?2:1}"><numFmt numFmtId="5000" formatCode="${code}"/>${overrideBuiltin?'<numFmt numFmtId="14" formatCode="yyyy/mm/dd"/>':''}</numFmts><fonts `);
 for(const n of[1,2]){
  let i=0;
  files[table(n)]=textOf(files[table(n)]).replace(/<pivotField\b/g,()=>`<pivotField${i++===0?' numFmtId="'+(n===1?14:5000)+'"':i===3?' numFmtId="0"':''}`);
 }
 return new Workbook(readXlsx(zip(files)).data);
}
function fields(files,n){return kids(child(parseXml(textOf(files[table(n)])),'pivotFields'),'pivotField');}
function assertFormats(files){
 assert.equal(fields(files,1)[0].attrs.numFmtId,'14');
 assert.equal(fields(files,1)[2].attrs.numFmtId,'0');assert.equal(fields(files,1)[3].attrs.numFmtId,undefined);
 const id=fields(files,2)[0].attrs.numFmtId;assert.notEqual(id,'5000');
 assert.equal(kids(child(parseXml(textOf(files['xl/styles.xml'])),'numFmts'),'numFmt').find(n=>n.attrs.numFmtId===id)?.attrs.formatCode,code);
 assert.equal(Object.keys(files).filter(p=>/pivotCacheDefinition\d+\.xml$/.test(p)).length,1);
}
for(const mode of['sync','blob'])test('공유 캐시의 피벗별 날짜 서식·General·서식 미지정은 독립 보존: '+mode,async()=>{
 const wb=source(),a=wb.sheets[1].pivot,b=wb.sheets[1].pivotsExtra[0];
 assert.equal(a.cacheItemsId,b.cacheItemsId);
 assert.deepEqual(a.fieldNumberFormats,{Date:{numFmt:'date',xlsxBuiltinId:14},Label:{numFmt:'general',xlsxBuiltinId:0}});
 assert.deepEqual(b.fieldNumberFormats.Date,{numFmt:'custom',code});
 const before=JSON.stringify([a,b,wb.pivotCacheItems]);
 const bytes=mode==='sync'?writeXlsx(wb):new Uint8Array(await(await writeXlsxBlobAsync(wb)).arrayBuffer());
 const files=unzip(bytes);assertFormats(files);
 assert.equal(JSON.stringify([a,b,wb.pivotCacheItems]),before);
 const back=new Workbook(readXlsx(bytes).data);
 assert.deepEqual(back.sheets[1].pivot.fieldNumberFormats,a.fieldNumberFormats);
 assert.deepEqual(back.sheets[1].pivotsExtra[0].fieldNumberFormats,b.fieldNumberFormats);
 assert.deepEqual(back.sheets[1].pivot.values,a.values);
});
test('피벗 필드 서식은 WIXEL·JSON 보관에도 남고 집계 값 서식을 바꾸지 않음',async()=>{
 const wb=source(),before=wb.sheets[1].pivot;
 for(const back of[new Workbook(JSON.parse(await wb.serializeBlob().text())),new Workbook((await readWixelFile(await writeWixelFile(wb))).workbook)]){
  assert.deepEqual(back.sheets[1].pivot.fieldNumberFormats,before.fieldNumberFormats);
  assert.deepEqual(back.sheets[1].pivot.values,before.values);assertFormats(unzip(writeXlsx(back)));
 }
});
test('내장 ID와 사용자 코드 중 명시한 필드 설정만 기록하고 계산 필드에도 보존',()=>{
 const wb=source(),p=wb.sheets[1].pivot;
 p.calcFields=[{name:'Calculated',formula:'Amount*2'}];p.values.push({field:'Calculated',agg:'sum'});
 p.fieldNumberFormats={...p.fieldNumberFormats,Date:{numFmt:'custom',code:'0.0000'},Calculated:{numFmt:'percent',decimals:2}};
 const files=unzip(writeXlsx(wb)),f=fields(files,1);
 const formats=Object.fromEntries(kids(child(parseXml(textOf(files['xl/styles.xml'])),'numFmts'),'numFmt').map(x=>[x.attrs.numFmtId,x.attrs.formatCode]));
 assert.equal(formats[f[0].attrs.numFmtId],'0.0000');assert.equal(formats[f.at(-1).attrs.numFmtId],'0.00%');
 assert.equal(f[3].attrs.numFmtId,undefined);
});

test('내장 번호에 명시된 원본 사용자 서식은 기본 로캘 서식보다 우선한다',()=>{
 const wb=source({overrideBuiltin:true}),format=wb.sheets[1].pivot.fieldNumberFormats.Date;
 assert.deepEqual(format,{numFmt:'custom',code:'yyyy/mm/dd'});
 const files=unzip(writeXlsx(wb)),id=fields(files,1)[0].attrs.numFmtId;
 assert.notEqual(id,'14');
 assert.equal(kids(child(parseXml(textOf(files['xl/styles.xml'])),'numFmts'),'numFmt').find(n=>n.attrs.numFmtId===id)?.attrs.formatCode,'yyyy/mm/dd');
 assert.deepEqual(new Workbook(readXlsx(zip(files)).data).sheets[1].pivot.fieldNumberFormats.Date,format);
});
