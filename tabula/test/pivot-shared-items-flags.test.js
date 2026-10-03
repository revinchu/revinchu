import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, writeXlsxAsync, readXlsx } from '../src/xlsx.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';
import { pivotSourceData } from '../src/pivot.js';
import { pivotValueStats } from '../src/pivot-export-data.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

function fixture(values) {
 const b=new PivotSnapshotBuilder(['Group','Value']);for(let i=0;i<values.length;i++)b.add([values[i],i+1]);
 return new Workbook({pivotSnapshots:{saved:b.finish()},sheets:[{name:'Source',cells:{'0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'different'},'1,1':{raw:'100'}}},{name:'Report',cells:{},pivot:{name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:1,c2:1},snapshotId:'saved',rows:['Group'],values:[{field:'Value',agg:'sum'}],top:0,left:0}}]});
}
function fields(bytes){return kids(child(parseXml(textOf(unzip(bytes)['xl/pivotCache/pivotCacheDefinition1.xml'])),'cacheFields'),'cacheField');}

test('pivot statistics distinguish missing cells from empty text without changing date-group blank counts',()=>{
 const s=pivotValueStats([null,'',2,'a']);assert.equal(s.blanks,2);assert.equal(s.missing,1);assert.equal(s.emptyStrings,1);assert.equal(s.hasString,true);
});
for(const async of [false,true])test(`empty text cache items never claim a missing item, async=${async}`,async()=>{
 const wb=fixture(['A','','']),bytes=async?await writeXlsxAsync(wb):writeXlsx(wb),f=fields(bytes),shared=child(f[0],'sharedItems');
 assert.equal(shared.attrs.containsBlank,undefined);assert.notEqual(shared.attrs.containsString,'0');assert.equal(kids(shared,'s').filter(x=>x.attrs.v==='').length,1);assert.equal(kids(shared,'m').length,0);
 const result=new Workbook(readXlsx(bytes).data),src=pivotSourceData(result,result.sheets[1].pivot);assert.equal(src.cube.n,3);assert.equal(src.cube.col(0).get(1),'');assert.equal(src.cube.col(1).get(2),3);
 // Office requires a numeric child before min/max bounds; value-only fields have none.
 const numbers=child(f[1],'sharedItems');assert.equal(numbers.attrs.minValue,undefined);assert.equal(numbers.attrs.maxValue,undefined);
});
test('empty-text-only, missing, mixed numeric and long text shared items retain exact types',()=>{
 for(const values of [[''],[null],[2,null],[2,''],['x'.repeat(256)]]){
  const bytes=writeXlsx(fixture(values)),shared=child(fields(bytes)[0],'sharedItems');
  assert.equal(shared.attrs.containsBlank,values.includes(null)?'1':undefined);
  if(values.includes(null)||values.some(v=>typeof v==='string'))assert.notEqual(shared.attrs.containsSemiMixedTypes,'0');
  if(values.some(v=>typeof v==='string'))assert.notEqual(shared.attrs.containsString,'0');
  if(values.some(v=>typeof v==='number')){assert.equal(shared.attrs.minValue,'2');assert.equal(kids(shared,'n').length,1);}
  if(values[0]?.length>255)assert.equal(shared.attrs.longText,'1');
  const back=new Workbook(readXlsx(bytes).data),src=pivotSourceData(back,back.sheets[1].pivot);assert.deepEqual(values.map((_,r)=>src.cube.col(0).get(r)),values);
 }
});

test('date-group shared metadata excludes empty-text dates and has no bounds without date children',()=>{
 for(const values of [[46000,null],[46000,'']]){
  const wb=fixture(values);wb.sheets[1].pivot.groups={Group:{by:'months'}};
  const bytes=writeXlsx(wb),shared=child(fields(bytes)[0],'sharedItems');assert.equal(shared.attrs.minDate,undefined);assert.equal(shared.attrs.maxDate,undefined);
  assert.notEqual(shared.attrs.containsSemiMixedTypes,'0');
  if(values[1]===''){assert.notEqual(shared.attrs.containsString,'0');assert.equal(shared.attrs.containsBlank,undefined);}else assert.equal(shared.attrs.containsBlank,'1');
  const after=new Workbook(readXlsx(bytes).data);assert.equal(after.sheets[1].pivot.groups.Group.by,'months');
 }
});
