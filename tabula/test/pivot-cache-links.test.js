import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxToSink } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';

// Microsoft OOXML missingItemsLimit and modern Excel maximum:
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.pivotcachedefinition
// https://learn.microsoft.com/en-us/office/vba/api/excel.xlpivottablemissingitems
const definition='xl/pivotCache/pivotCacheDefinition1.xml';
function fixture() {
 const wb=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Item'},'0,1':{raw:'Amount'},'1,0':{raw:'Now'},'1,1':{raw:'10'},'2,0':{raw:'Hidden'},'2,1':{raw:'20'}}},{name:'Report',cells:{}}]});
 wb.pivotCacheItems={original:{fields:[{name:'Item',shared:['Past','Now','Hidden']},{name:'Amount',shared:[]}]}};
 wb.sheets[1].pivot={name:'First',source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Item'],values:[{field:'Amount',agg:'sum'}],filters:{Item:['Past','Now']},top:0,left:0,cacheItemsId:'original'};
 return wb;
}
const cacheRoots=parts=>Object.keys(parts).filter(k=>/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(k)).map(k=>parseXml(textOf(parts[k])));
const history=root=>kids(child(root,'cacheFields'),'cacheField').map(f=>({name:f.attrs.name,items:child(f,'sharedItems')?.children.map(i=>i.attrs.v??null)??[]}));
const reopened=wb=>new Workbook(readXlsx(writeXlsx(wb)).data);
function withPolicy(limit) {
 const files=unzip(writeXlsx(fixture()));
 if(limit!==undefined)files[definition]=textOf(files[definition]).replace('<pivotCacheDefinition ',`<pivotCacheDefinition missingItemsLimit="${limit}" `);
 return new Workbook(readXlsx(zip(files)).data);
}
for(const [limit,mode]of [[undefined,undefined],[0,'none'],[1048576,'max'],[32500,32500],[125,125]])test(`missing items policy reads and writes exact OOXML threshold: ${limit}`,()=>{
 const wb=withPolicy(limit);assert.equal(wb.sheets[1].pivot.missingItems,mode);
 const before=JSON.stringify(wb.pivotCacheItems),parts=unzip(writeXlsx(wb));
 assert.equal(cacheRoots(parts)[0].attrs.missingItemsLimit,limit===undefined?undefined:String(limit));
 assert.deepEqual(history(cacheRoots(parts)[0])[0].items,['Past','Now','Hidden'],'style-only export keeps the saved cache, even with a loaded none policy');
 assert.equal(JSON.stringify(wb.pivotCacheItems),before);
 assert.deepEqual(reopened(wb).sheets[1].pivot.filters.Item,['Past','Now']);
});
test('switching to none removes missing items but preserves present values and aggregates',()=>{
 const wb=withPolicy(undefined);wb.sheets[1].pivot={...wb.sheets[1].pivot,missingItems:'none'};
 const parts=unzip(writeXlsx(wb)),root=cacheRoots(parts)[0];
 assert.equal(root.attrs.missingItemsLimit,'0');assert.deepEqual(new Set(history(root)[0].items),new Set(['Now','Hidden']));
 const back=new Workbook(readXlsx(zip(parts)).data);assert.deepEqual(back.sheets[1].pivot.filters.Item,['Now']);assert.equal(back.getValue(0,1,1),10);
});
test('none policy prunes historical items after the saved snapshot is refreshed or invalidated',()=>{
 const wb=withPolicy(0);wb.pivotSnapshots=null;
 const root=cacheRoots(unzip(writeXlsx(wb)))[0];assert.deepEqual(new Set(history(root)[0].items),new Set(['Now','Hidden']));
 const max=withPolicy(1048576);max.pivotSnapshots=null;assert.deepEqual(history(cacheRoots(unzip(writeXlsx(max)))[0])[0].items,['Past','Now','Hidden']);
});
function linkedBook({newPivot=false,unlinked=false}={}) {
 const wb=withPolicy(undefined),first=wb.sheets[1].pivot;
 wb.pivotCacheItems={...wb.pivotCacheItems,second:{fields:[{name:'Item',shared:['Other past','Hidden','Now']}]}};
 const second={...structuredClone(first),name:'Second',top:10,cacheItemsId:'second'};
 if(newPivot){delete second.cacheItemsId;delete second.snapshotId;}
 wb.sheets[1].pivotsExtra=[second];
 if(unlinked)wb.sheets[1].pivotsExtra.push({...structuredClone(first),name:'Independent',top:20,cacheItemsId:'independent'});
 wb.pivotCacheItems.independent={fields:[{name:'Item',shared:['Independent past','Now','Hidden']}]};
 wb.sheets[1].slicers=[{id:'s1',caption:'Item',source:{kind:'pivot',field:'Item',pivots:[{sheet:'Report',name:'First'},{sheet:'Report',name:'Second'}]},x:0,y:100,w:140,h:200,showDeleted:true}];
 return wb;
}
for(const newPivot of[false,true])test(`explicit slicer links retain both reports across independent/imported-new caches: new=${newPivot}`,()=>{
 const wb=linkedBook({newPivot,unlinked:true}),before=JSON.stringify(wb.pivotCacheItems),parts=unzip(writeXlsx(wb)),roots=cacheRoots(parts);
 assert.equal(roots.length,2,'only explicit connections merge; independent third cache stays separate');
 assert.deepEqual(history(roots[0])[0].items,newPivot?['Past','Now','Hidden']:['Past','Now','Hidden','Other past']);
 assert.deepEqual(history(roots[1])[0].items,['Independent past','Now','Hidden']);
 const sc=Object.keys(parts).find(k=>/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(k));
 assert.deepEqual(descendants(parseXml(textOf(parts[sc])),'pivotTable').map(t=>t.attrs.name),['First','Second']);
 const back=new Workbook(readXlsx(zip(parts)).data);assert.equal(back.sheets[1].pivot.cacheItemsId,back.sheets[1].pivotsExtra[0].cacheItemsId);
 assert.notEqual(back.sheets[1].pivot.cacheItemsId,back.sheets[1].pivotsExtra[1].cacheItemsId);
 assert.equal(back.sheets[1].slicers[0].source.pivots.length,2);
 assert.deepEqual(back.sheets[1].pivot.filters.Item,['Past','Now']);assert.deepEqual(back.sheets[1].pivotsExtra[0].filters.Item,['Past','Now']);
 assert.equal(JSON.stringify(wb.pivotCacheItems),before,'export never mutates shared historical metadata');
});
test('explicit connections do not silently replace different saved cache values',()=>{
 const wb=linkedBook(),second=wb.sheets[1].pivotsExtra[0];
 wb.pivotSnapshots.set('different',{rows:[['Item','Amount'],['Now',999],['Hidden',20]],ver:undefined});second.snapshotId='different';
 assert.throws(()=>writeXlsx(wb),/저장된 데이터가 다릅니다/);
});
test('disk cancellation reaches ZIP input even before another compressed write',async()=>{
 const wb=fixture(),controller=new AbortController();let writes=0;
 await assert.rejects(writeXlsxToSink(wb,{signal:controller.signal},{async write(){writes++;controller.abort(new Error('cancel input'));}}),/cancel input/);
 assert.equal(writes,1);
});
