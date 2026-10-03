import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { Cube } from '../src/cube.js';
import { readXlsx, writeXlsx, writeXlsxAsync } from '../src/xlsx.js';
import { pivotSourceData } from '../src/pivot.js';
import { unzip, textOf, zip } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

function fixture(saved = true) {
  return new Workbook({ ...(saved ? { pivotSnapshots: { old: [['Group','Value'],['A',10]] } } : {}), sheets: [
    { name:'Source', cells:{
      '0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'A'},'1,1':{raw:'999'},
      '2,0':{raw:'B'},'2,1':{raw:'2'},'4,0':{raw:"'"},'4,1':{raw:'FALSE'},
      '5,0':{raw:'TRUE'},'5,1':{raw:'#N/A'},
    } },
    { name:'Report', cells:{}, pivot:{name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:5,c2:1},
      ...(saved ? {snapshotId:'old'} : {}),rows:['Group'],values:[{field:'Value',agg:'sum'}],top:0,left:0} },
  ] });
}
const definition = files => parseXml(textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']));
const records = files => parseXml(textOf(files['xl/pivotCache/pivotCacheRecords1.xml']));
const sourceValues = wb => { const cube=pivotSourceData(wb,wb.sheets[1].pivot).cube;return Array.from({length:cube.n},(_,r)=>cube.header.map((_,c)=>cube.col(c).get(r))); };

for(const async of [false,true]) test(`edited source exports current records instead of stale imported values, async=${async}`,async()=>{
  const wb=fixture();assert.deepEqual(sourceValues(wb),[['A',10]]);
  wb.transact(()=>wb.setInput(0,1,1,'777'));assert.deepEqual(wb.snapshotData(),{});
  const row=Cube.prototype.row;let bytes;
  try { Cube.prototype.row=()=>{throw Error('row-array materialization')};bytes=async?await writeXlsxAsync(wb):writeXlsx(wb); }
  finally { Cube.prototype.row=row; }
  const files=unzip(bytes),def=definition(files),rec=records(files);
  assert.equal(def.attrs.saveData,'1');assert.equal(def.attrs.refreshOnLoad,'0');assert.equal(def.attrs.recordCount,'5');
  assert.equal(def.attrs['r:id'],'rCacheRecords');assert.equal(rec.attrs.count,'5');assert.equal(kids(rec,'r').length,5);
  assert.match(textOf(files['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels']),/Target="pivotCacheRecords1.xml"/);
  assert.match(textOf(files['[Content_Types].xml']),/PartName="\/xl\/pivotCache\/pivotCacheRecords1.xml"/);
  const after=new Workbook(readXlsx(bytes).data);
  assert.deepEqual(sourceValues(after),[['A',777],['B',2],[null,null],['',false],[true,{error:'#N/A'}]]);
  assert.equal(after.getValue(0,1,1),777);assert.ok(after.sheets[1].pivot.snapshotId);
});

test('a newly created pivot saves all current rows, including an entirely blank row',()=>{
  const wb=fixture(false),after=new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(sourceValues(after),[['A',999],['B',2],[null,null],['',false],[true,{error:'#N/A'}]]);
});

for(const saveData of [false,true])for(const refreshOnOpen of [false,true])test(`cache options honor saveData=${saveData} and refreshOnOpen=${refreshOnOpen}`,()=>{
  const wb=fixture(false);Object.assign(wb.sheets[1].pivot,{saveData,refreshOnOpen});
  const files=unzip(writeXlsx(wb)),def=definition(files);
  assert.equal(def.attrs.saveData,saveData?'1':'0');assert.equal(def.attrs.refreshOnLoad,refreshOnOpen?'1':'0');
  assert.equal(def.attrs.recordCount,saveData?'5':'0');
  assert.equal(Object.hasOwn(files,'xl/pivotCache/pivotCacheRecords1.xml'),saveData);
  assert.equal(Object.hasOwn(files,'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'),saveData);
  assert.equal(!!def.attrs['r:id'],saveData);
  const after=new Workbook(readXlsx(zip(files)).data),p=after.sheets[1].pivot;
  assert.equal(p.saveData!==false,saveData);assert.equal(!!p.refreshOnOpen,refreshOnOpen);
  assert.equal(!!p.snapshotId,saveData&&!refreshOnOpen);
});

test('shared cache conflicts honor any explicit storage exclusion and any refresh request',()=>{
  const wb=fixture(false);wb.sheets[1].pivotsExtra=[{...wb.sheets[1].pivot,name:'Second',top:10,saveData:false,refreshOnOpen:true}];
  const files=unzip(writeXlsx(wb)),def=definition(files);
  assert.equal(def.attrs.saveData,'0');assert.equal(def.attrs.refreshOnLoad,'1');assert.equal(def.attrs.recordCount,'0');
  assert.equal(Object.keys(files).filter(p=>/pivotCacheDefinition\d+\.xml$/.test(p)).length,1);
  const after=readXlsx(zip(files)).data.sheets[1];
  for(const p of [after.pivot,...after.pivotsExtra]){assert.equal(p.saveData,false);assert.equal(p.refreshOnOpen,true);}
});

test('cache reader accepts true/false XML booleans and their default settings',()=>{
  const files=unzip(writeXlsx(fixture(false))),path='xl/pivotCache/pivotCacheDefinition1.xml',xml=textOf(files[path]);
  files[path]=xml.replace('saveData="1"','saveData="false"').replace('refreshOnLoad="0"','refreshOnLoad="true"');
  let after=readXlsx(zip(files)).data.sheets[1].pivot;assert.equal(after.saveData,false);assert.equal(after.refreshOnOpen,true);assert.equal(after.snapshotId,undefined);
  files[path]=xml.replace(' saveData="1"','').replace(' refreshOnLoad="0"','');
  after=readXlsx(zip(files)).data.sheets[1].pivot;assert.equal(after.saveData,undefined);assert.equal(after.refreshOnOpen,undefined);assert.ok(after.snapshotId);
});
