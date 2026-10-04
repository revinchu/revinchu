import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, isoSerial } from '../src/xlsx.js';
import { readWixelFile, writeWixelFile } from '../src/wixel-file.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const definition='xl/pivotCache/pivotCacheDefinition1.xml',table='xl/pivotTables/pivotTable1.xml',records='xl/pivotCache/pivotCacheRecords1.xml';
const dates=['2024-03-01T00:00:00','2024-12-17T12:34:56.789','2024-12-31T00:00:00'];
const dateCode='yyyy"년" m"월" d"일" hh:mm:ss.000';
function fixture({date1904=false,mixed=false}={}) {
 const wb=new Workbook();wb.date1904=date1904;wb.sheets[0].name='Source';
 const current=[isoSerial(dates[1],date1904),mixed?7:isoSerial(dates[2],date1904)];
 wb.transact(()=>{
  [['Date','Value'],[current[0],10],[current[1],20]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
  wb.setStyle(0,1,0,{numFmt:'custom',code:dateCode});wb.addSheet('Report');
  wb.setSheetProp(1,'pivot',{name:'Report',source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Date'],values:[{field:'Value',agg:'sum'}],filters:{Date:[current[0]]},top:0,left:0});
  wb.setSheetProp(1,'slicers',[{id:'date',caption:'Date',source:{kind:'pivot',field:'Date',self:true},x:0,y:100,w:150,h:220,showDeleted:true}]);
 });
 const files=unzip(writeXlsx(wb)),styles=parseXml(textOf(files['xl/styles.xml']));
 const fmt=kids(child(styles,'numFmts'),'numFmt').find(n=>n.attrs.formatCode===dateCode)?.attrs.numFmtId;assert.ok(fmt);
 const values=['<d v="'+dates[0]+'"/>','<d v="'+dates[1]+'"/>',mixed?'<n v="7"/>':'<d v="'+dates[2]+'"/>'];
 files[definition]=textOf(files[definition]).replace(/<cacheField\b[^>]*>[\s\S]*?<\/cacheField>/,xml=>xml.replace(/numFmtId="\d+"/,'numFmtId="'+fmt+'"').replace(/<sharedItems\b[^>]*>[\s\S]*?<\/sharedItems>/,'<sharedItems count="3" containsDate="1">'+values.join('')+'</sharedItems>'));
 files[table]=textOf(files[table]).replace(/<pivotField\b[^>]*>[\s\S]*?<\/pivotField>/,xml=>xml.replace(/<items\b[^>]*>[\s\S]*?<\/items>/,'<items count="3"><item x="0"/><item x="1"/><item x="2" h="1"/></items>'));
 let r=0;files[records]=textOf(files[records]).replace(/<r><n v="[^"]*"\/>/g,()=>'<r>'+values[++r]);
 return new Workbook(readXlsx(zip(files)).data);
}
const cacheFields=files=>kids(child(parseXml(textOf(files[definition])),'cacheFields'),'cacheField');
const shared=files=>child(cacheFields(files)[0],'sharedItems');
const fieldMeta=wb=>wb.pivotCacheItems[wb.sheets[1].pivot.cacheItemsId].fields[0];
function assertDateFormat(files) {
 const fmt=cacheFields(files)[0].attrs.numFmtId,styles=parseXml(textOf(files['xl/styles.xml']));
 assert.equal(kids(child(styles,'numFmts'),'numFmt').find(n=>n.attrs.numFmtId===fmt)?.attrs.formatCode,dateCode);
}
for(const date1904 of[false,true])test('날짜 캐시의 과거 선택·타입·사용자 서식과 밀리초 보존: date1904='+date1904,()=>{
 const wb=fixture({date1904}),before=JSON.stringify(wb.pivotCacheItems),originalFilters=wb.sheets[1].pivot.filters.Date;
 assert.equal(fieldMeta(wb).sharedTypes,'ddd');assert.equal(fieldMeta(wb).format.code,dateCode);
 wb.transact(()=>wb.setSheetProp(1,'pivot',{...wb.sheets[1].pivot,style:'PivotStyleMedium9'}));
 const files=unzip(writeXlsx(wb)),items=shared(files),back=new Workbook(readXlsx(zip(files)).data);
 assert.deepEqual(items.children.map(n=>n.name),['d','d','d']);assert.deepEqual(items.children.map(n=>n.attrs.v),dates);
 assert.equal(items.attrs.containsDate,'1');assert.equal(items.attrs.containsNonDate,'0');assert.equal(items.attrs.containsNumber,undefined);
 assert.equal(items.attrs.minValue,undefined);assert.equal(items.attrs.maxValue,undefined);assertDateFormat(files);
 assert.deepEqual(kids(parseXml(textOf(files[records])),'r').map(row=>row.children[0].name),['d','d']);
 assert.deepEqual(back.sheets[1].pivot.filters.Date,originalFilters);assert.deepEqual(fieldMeta(back),fieldMeta(wb));
 assert.equal(JSON.stringify(wb.pivotCacheItems),before);
});
test('날짜와 숫자가 섞인 캐시는 숫자를 날짜로 바꾸지 않음',()=>{
 const wb=fixture({mixed:true}),files=unzip(writeXlsx(wb)),items=shared(files);
 assert.deepEqual(items.children.map(n=>n.name),['d','d','n']);assert.equal(items.children[2].attrs.v,'7');
 assert.equal(items.attrs.containsDate,'1');assert.equal(items.attrs.containsNumber,'1');assert.equal(items.attrs.containsMixedTypes,'1');
 assert.equal(items.attrs.minValue,'7');assert.equal(items.attrs.maxValue,'7');assert.equal(items.attrs.containsNonDate,undefined);
 assert.deepEqual(kids(parseXml(textOf(files[records])),'r').map(row=>row.children[0].name),['d','n']);
});
test('원본 날짜 추가와 누락 항목 제거도 날짜 타입·서식 유지',()=>{
 const wb=fixture(),newDate='2025-01-05T00:00:00';
 wb.transact(()=>wb.setInput(0,2,0,String(isoSerial(newDate))));
 let files=unzip(writeXlsx(wb));assert.deepEqual(shared(files).children.map(n=>n.attrs.v),[...dates,newDate]);assertDateFormat(files);
 wb.sheets[1].pivot={...wb.sheets[1].pivot,missingItems:'none'};
 files=unzip(writeXlsx(wb));assert.deepEqual(new Set(shared(files).children.map(n=>n.attrs.v)),new Set([dates[1],newDate]));
 assert.deepEqual(shared(files).children.map(n=>n.name),['d','d']);assertDateFormat(files);
});
test('WIXEL과 Blob 복원은 날짜 타입·서식 메타를 공유 캐시와 함께 보존',async()=>{
 const wb=fixture();
 for(const restored of[new Workbook(JSON.parse(await wb.serializeBlob().text())),new Workbook((await readWixelFile(await writeWixelFile(wb))).workbook)]){
  assert.deepEqual(restored.pivotCacheItems,JSON.parse(JSON.stringify(wb.pivotCacheItems)));
  const files=unzip(writeXlsx(restored));assert.deepEqual(shared(files).children.map(n=>n.attrs.v),dates);assertDateFormat(files);
 }
});
test('명시 연결 캐시 union도 추가 과거 날짜의 타입 보존',()=>{
 const wb=fixture(),first=wb.sheets[1].pivot,original=wb.pivotCacheItems[first.cacheItemsId],extra='2023-02-01T00:00:00';
 wb.pivotCacheItems={...wb.pivotCacheItems,other:{fields:original.fields.map((f,i)=>i?f:{...f,shared:[...f.shared,isoSerial(extra)],sharedTypes:'dddd'})}};
 wb.sheets[1].pivotsExtra=[{...structuredClone(first),name:'Second',top:15,cacheItemsId:'other'}];
 wb.sheets[1].slicers[0].source={kind:'pivot',field:'Date',pivots:[{sheet:'Report',name:'Report'},{sheet:'Report',name:'Second'}]};
 const before=JSON.stringify(wb.pivotCacheItems),files=unzip(writeXlsx(wb));
 assert.equal(Object.keys(files).filter(k=>/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(k)).length,1);
 assert.deepEqual(shared(files).children.map(n=>n.name),['d','d','d','d']);assert.deepEqual(shared(files).children.map(n=>n.attrs.v),[...dates,extra]);
 assert.equal(JSON.stringify(wb.pivotCacheItems),before);assertDateFormat(files);
});

function recordsOnlyFixture(booleanWords=false) {
 const wb=new Workbook();wb.sheets[0].name='Source';
 wb.transact(()=>{
  [['Date','Group','Value'],[45352,'A',10],[45353,'B',20]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
  wb.addSheet('Report');wb.setSheetProp(1,'pivot',{name:'Report',source:'Source',range:{r1:0,c1:0,r2:2,c2:2},rows:['Group'],values:[{field:'Value',agg:'sum'}]});
 });
 const files=unzip(writeXlsx(wb));
 files[definition]=textOf(files[definition]).replace(/<sharedItems\b[^>]*\/>/,'<sharedItems containsDate="'+(booleanWords?'true':'1')+'" containsNonDate="'+(booleanWords?'false':'0')+'" containsString="0"/>');
 let i=0;files[records]=textOf(files[records]).replace(/<r><n v="[^"]*"\/>/g,()=>'<r><d v="2024-03-0'+(++i)+'T00:00:00"/>');
 return new Workbook(readXlsx(zip(files)).data);
}
for(const booleanWords of[false,true])test('목록 없는 날짜 캐시도 records 날짜 타입과 빈 sharedItems 속성 유지: '+booleanWords,()=>{
 const wb=recordsOnlyFixture(booleanWords),before=JSON.stringify(wb.pivotCacheItems);
 assert.deepEqual(fieldMeta(wb),{name:'Date',shared:[],dateOnly:true});
 const files=unzip(writeXlsx(wb));assert.equal(shared(files).attrs.containsDate,'1');assert.equal(shared(files).attrs.containsNonDate,'0');
 assert.equal(shared(files).children.length,0);assert.deepEqual(kids(parseXml(textOf(files[records])),'r').map(r=>r.children[0].name),['d','d']);
 const back=new Workbook(readXlsx(zip(files)).data);assert.equal(fieldMeta(back).dateOnly,true);assert.equal(JSON.stringify(wb.pivotCacheItems),before);
 wb.sheets[1].pivot={...wb.sheets[1].pivot,rows:['Date']};
 const active=unzip(writeXlsx(wb));assert.deepEqual(shared(active).children.map(n=>n.name),['d','d']);
});
test('목록 없는 날짜 캐시 dateOnly 메타는 WIXEL과 Blob에 보존',async()=>{
 const wb=recordsOnlyFixture();
 for(const restored of[new Workbook(JSON.parse(await wb.serializeBlob().text())),new Workbook((await readWixelFile(await writeWixelFile(wb))).workbook)]){
  assert.equal(fieldMeta(restored).dateOnly,true);
  const files=unzip(writeXlsx(restored));assert.deepEqual(kids(parseXml(textOf(files[records])),'r').map(r=>r.children[0].name),['d','d']);
 }
});
