import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxAsync } from '../src/xlsx.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

function fixture(values = ['Alpha','ALPHA','Beta'], saved = false) {
  const cells = {'0,0':{raw:'Group'},'0,1':{raw:'Value'}};
  values.forEach((v,i)=>{if(v!==null)cells[`${i+1},0`]={raw:typeof v==='string'?"'"+v:String(v)};cells[`${i+1},1`]={raw:String((i+1)*10)};});
  return new Workbook({...(saved?{pivotSnapshots:{saved:[['Group','Value'],...values.map((v,i)=>[v,(i+1)*10])]}}:{}),sheets:[
    {name:'Source',cells},
    {name:'Report',cells:{},pivot:{name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:values.length,c2:1},...(saved?{snapshotId:'saved'}:{}),rows:['Group'],values:[{field:'Value',agg:'sum'}],top:0,left:0},
      slicers:[{id:'Slicer',caption:'Group',source:{kind:'pivot',field:'Group',self:true},x:240,y:20,w:160,h:180}]},
  ]});
}
const xml=(files,path)=>parseXml(textOf(files[path]));
const field=(files)=>kids(child(xml(files,'xl/pivotCache/pivotCacheDefinition1.xml'),'cacheFields'),'cacheField')[0];
const grid=wb=>{const d=wb.sheets[1].pivot,r=resolvePivot(pivotSourceData(wb,d),d);return computePivot(r,r.def).grid.map(row=>row.map(c=>c.raw));};

for(const async of [false,true])for(const saved of [false,true])test(`case aliases share pivot and slicer items but preserve cell/record spelling, saved=${saved}, async=${async}`,async()=>{
  const wb=fixture(undefined,saved),before=grid(wb),bytes=async?await writeXlsxAsync(wb):writeXlsx(wb),files=unzip(bytes);
  const shared=child(field(files),'sharedItems');assert.equal(shared.attrs.count,'2');assert.deepEqual(kids(shared,'s').map(n=>n.attrs.v),['Alpha','Beta']);
  const rec=xml(files,'xl/pivotCache/pivotCacheRecords1.xml');assert.deepEqual(kids(rec,'r').map(r=>r.children[0].attrs.v),['Alpha','ALPHA','Beta']);
  const pivot=xml(files,'xl/pivotTables/pivotTable1.xml'),items=child(kids(child(pivot,'pivotFields'),'pivotField')[0],'items');
  assert.equal(kids(items,'item').filter(i=>i.attrs.x!==undefined).length,2);
  const tabular=child(child(xml(files,'xl/slicerCaches/slicerCache1.xml'),'data'),'tabular');assert.equal(child(tabular,'items').attrs.count,'2');
  const after=new Workbook(readXlsx(bytes).data);assert.deepEqual(grid(after),before);assert.deepEqual(before.slice(1,3),[['Alpha','30'],['Beta','30']]);
  assert.deepEqual([1,2,3].map(r=>after.getValue(0,r,0)),['Alpha','ALPHA','Beta']);
  const src=pivotSourceData(after,after.sheets[1].pivot).cube;assert.deepEqual([0,1,2].map(r=>src.col(0).get(r)),['Alpha','ALPHA','Beta']);
});

test('alias selections, manual order, captions and collapsed state survive case canonicalization',()=>{
  const wb=fixture();Object.assign(wb.sheets[1].pivot,{filters:{Group:['aLpHa']},order:{Group:['beta','alpha']},itemCaptions:{Group:{ALPHA:'Caption'}},collapsed:{Group:['aLPhA']}});
  const bytes=writeXlsx(wb),files=unzip(bytes),pivot=xml(files,'xl/pivotTables/pivotTable1.xml');
  const items=kids(child(kids(child(pivot,'pivotFields'),'pivotField')[0],'items'),'item').filter(i=>i.attrs.x!==undefined);
  assert.deepEqual(items.map(i=>i.attrs.x),['1','0']);assert.equal(items[0].attrs.h,'1');assert.equal(items[1].attrs.h,undefined);assert.equal(items[1].attrs.sd,'0');assert.equal(items[1].attrs.n,'Caption');
  const sl=kids(child(child(child(xml(files,'xl/slicerCaches/slicerCache1.xml'),'data'),'tabular'),'items'),'i');
  assert.equal(sl[0].attrs.s,'1');assert.equal(sl[1].attrs.s,undefined);
  const after=new Workbook(readXlsx(bytes).data);assert.deepEqual(grid(after),grid(wb));
});

test('single-page selection can use a different case from the canonical shared item',()=>{
  const wb=fixture();Object.assign(wb.sheets[1].pivot,{rows:[],pages:['Group'],filters:{Group:['aLpHa']},pageMulti:{Group:false}});
  const files=unzip(writeXlsx(wb)),pivot=xml(files,'xl/pivotTables/pivotTable1.xml');assert.equal(kids(child(pivot,'pageFields'),'pageField')[0].attrs.item,'0');
});

test('typed and Unicode values remain separate except ordinary case aliases',()=>{
  const values=[1,'1',null,'','한글','한글','É','é','e','Ｅ','E'];
  const wb=fixture(values),bytes=writeXlsx(wb),files=unzip(bytes),shared=child(field(files),'sharedItems');
  assert.equal(shared.attrs.count,'8');assert.equal(kids(shared,'n').length,1);assert.equal(kids(shared,'m').length,1);
  assert.deepEqual(kids(shared,'s').map(n=>n.attrs.v).sort(),['','1','한글','É','e','Ｅ'].sort());
  const after=new Workbook(readXlsx(bytes).data),src=pivotSourceData(after,after.sheets[1].pivot).cube;assert.deepEqual(values.map((_,r)=>src.col(0).get(r)),values);
});
