import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';
import { readXlsx,writeXlsx } from '../src/xlsx.js';
import { pivotSourceData,resolvePivot,computePivot } from '../src/pivot.js';
import { unzip,textOf } from '../src/zip.js';
function make(rows,compact=false) {
 let saved=rows;
 if(compact){const b=new PivotSnapshotBuilder(rows[0]);rows.slice(1).forEach(row=>b.add(row));saved=b.finish();}
 return new Workbook({pivotSnapshots:{saved},sheets:[{name:'원본',cells:{'0,0':{raw:'분류'},'0,1':{raw:'값'},'1,0':{raw:'가'},'1,1':{raw:'999'}}},{name:'보고서',cells:{},pivot:{name:'저장피벗',source:'원본',range:{r1:0,c1:0,r2:1,c2:1},snapshotId:'saved',rows:['분류'],values:[{field:'값',agg:'sum'}],top:0,left:0}}]});
}
const source=(wb,si=1)=>pivotSourceData(wb,wb.sheets[si].pivot).rows;
for(const compact of [false,true])test(`XLSX saved cache survives export without refreshing from changed source, compact=${compact}`,()=>{
 const rows=[['분류','값'],['가',10],['',0],[null,null],[true,{error:'#N/A'}],['줄\n바꿈_x000A_',false]];
 const wb=make(rows,compact),bytes=writeXlsx(wb),parts=unzip(bytes),after=new Workbook(readXlsx(bytes).data);
 assert.deepEqual(source(after),rows);assert.equal(after.getValue(0,1,1),999);
 const def=textOf(parts['xl/pivotCache/pivotCacheDefinition1.xml']);
 assert.match(def,/saveData="1"/);assert.match(def,/refreshOnLoad="0"/);assert.match(def,/r:id="rCacheRecords"/);
 assert.match(textOf(parts['[Content_Types].xml']),/pivotCacheRecords\+xml/);
 assert.match(textOf(parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels']),/pivotCacheRecords1.xml/);
});
test('separate saved caches on the same worksheet range are not merged during XLSX export',()=>{
 const wb=make([['분류','값'],['가',10]],true);
 const second={...wb.sheets[1].pivot,name:'별도피벗',snapshotId:'saved2'};
 const data=wb.serialize();data.sheets.push({name:'두번째',cells:{},pivot:second});data.pivotSnapshots.saved2=[['분류','값'],['가',20]];
 const before=new Workbook(data),after=new Workbook(readXlsx(writeXlsx(before)).data);
 assert.equal(source(after,1)[1][1],10);assert.equal(source(after,2)[1][1],20);
 assert.equal(after.pivotSnapshots.size,2);
});
test('source edit discards stale cache before XLSX export',()=>{
 const wb=make([['분류','값'],['가',10]],true);wb.transact(()=>wb.setInput(0,1,1,'777'));
 const after=new Workbook(readXlsx(writeXlsx(wb)).data);assert.equal(source(after)[1][1],777);
});
test('date grouped saved cache retains its date and time during XLSX export',()=>{
 const wb=make([['분류','값'],[45292.50000142361,10]],true);
 wb.sheets[1].pivot.groups={분류:{by:'months'}};
 const after=new Workbook(readXlsx(writeXlsx(wb)).data);
 assert.ok(Math.abs(source(after)[1][0]-45292.50000142361)<1e-9);assert.equal(source(after)[1][1],10);
});

test('empty text, blank and boolean pivot item labels and ordering survive cached export',()=>{
 const wb=make([['분류','값'],['가',10],['',0],[null,null],[true,20],['줄\n바꿈',30]],true);
 const grid=w=>{const r=resolvePivot(pivotSourceData(w,w.sheets[1].pivot),w.sheets[1].pivot);return computePivot(r,r.def).grid.map(row=>row.map(cell=>cell.raw));};
 const before=grid(wb),after=new Workbook(readXlsx(writeXlsx(wb)).data);assert.deepEqual(grid(after),before);
 wb.sheets[1].pivot.filters={분류:['']};const filtered=new Workbook(readXlsx(writeXlsx(wb)).data);assert.deepEqual(filtered.sheets[1].pivot.filters.분류,['']);
});
