import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx,writeXlsx,writeXlsxAsync } from '../src/xlsx.js';
import { unzip,textOf } from '../src/zip.js';
import { parseXml,kids,child } from '../src/xml.js';

function fixture(){
 const cells={'0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'A'},'1,1':{raw:'10'},'2,0':{raw:'B'},'2,1':{raw:'20'}};
 const sheets=[{name:'Source',cells,tables:[{name:'Data',r1:0,c1:0,r2:2,c2:1,header:true,columns:[{name:'Group'},{name:'Value'}]}]}];
 for(let i=1;i<=3;i++)sheets.push({name:'Report'+i,cells:{},pivot:{name:'Pivot'+i,source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Group'],values:[{field:'Value',agg:'sum'}],top:0,left:0},slicers:[]});
 return new Workbook({sheets});
}
const refs=ids=>ids.map(i=>({sheet:'Report'+i,name:'Pivot'+i}));
const slicer=(id,connections,x=20,field='Group')=>({id,caption:id,source:{kind:'pivot',field,pivots:refs(connections)},x,y:20,w:160,h:200,style:'SlicerStyleLight1'});
function parts(bytes){const zip=unzip(bytes);return{zip,caches:Object.keys(zip).filter(p=>/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(p)).map(p=>parseXml(textOf(zip[p]))),slicers:Object.keys(zip).filter(p=>/^xl\/slicers\/slicer\d+\.xml$/.test(p)).flatMap(p=>kids(parseXml(textOf(zip[p])),'slicer'))};}

for(const async of [false,true])test(`same logical pivot cache is shared across sheets and reordered connections, async=${async}`,async()=>{
 const wb=fixture();wb.sheets[1].slicers=[slicer('First',[1,2])];wb.sheets[2].slicers=[{...slicer('Second',[2,1,1],220),style:'SlicerStyleDark2',showHeader:false}];wb.sheets[3].slicers=[slicer('Independent',[3])];
 const before=JSON.stringify(wb.sheets.map(s=>s.slicers));for(const s of wb.sheets)for(const sl of s.slicers??[])Object.freeze(sl);
 const bytes=async?await writeXlsxAsync(wb):writeXlsx(wb),p=parts(bytes);
 assert.equal(p.caches.length,2);assert.equal(p.slicers.length,3);assert.equal(p.slicers[0].attrs.cache,p.slicers[1].attrs.cache);assert.notEqual(p.slicers[0].attrs.cache,p.slicers[2].attrs.cache);
 assert.equal(kids(child(p.caches[0],'pivotTables'),'pivotTable').length,2);assert.equal(JSON.stringify(wb.sheets.map(s=>s.slicers)),before);
 const after=readXlsx(bytes).data.sheets;assert.equal(after[1].slicers[0].caption,'First');assert.equal(after[2].slicers[0].caption,'Second');assert.equal(after[2].slicers[0].x,220);assert.equal(after[2].slicers[0].style,'SlicerStyleDark2');assert.equal(after[2].slicers[0].showHeader,false);assert.equal(after[2].slicers[0].source.pivots.length,2);
});
test('distinct fields and table/standalone sources never alias a pivot slicer cache',()=>{
 const wb=fixture();wb.sheets[1].slicers=[slicer('PivotGroup',[1]),slicer('PivotValue',[1],220,'Value'),{...slicer('Table',[1],420),source:{kind:'table',table:'Data',column:'Group'}},{...slicer('Standalone',[1],620),source:{kind:'cache',cacheKey:'saved',field:'Group',cacheSource:{name:'Data'},values:['A','B']}}];
 const p=parts(writeXlsx(wb));assert.equal(p.caches.length,4);assert.equal(p.slicers.length,4);assert.equal(new Set(p.slicers.map(s=>s.attrs.cache)).size,4);
});
test('cache-level options do not split identical logical caches or discard physical slicer formatting',()=>{
 const wb=fixture();wb.sheets[1].slicers=[{...slicer('First',[1]),sort:'desc',hideNoData:true},{...slicer('Second',[1],220),sort:'asc',style:'SlicerStyleDark2'}];
 const p=parts(writeXlsx(wb));assert.equal(p.caches.length,1);assert.equal(p.slicers.length,2);assert.equal(child(child(p.caches[0],'data'),'tabular').attrs.sortOrder,'descending');assert.equal(p.slicers[1].attrs.style,'SlicerStyleDark2');
});
