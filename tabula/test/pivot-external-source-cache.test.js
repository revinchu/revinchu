import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { PivotSnapshotBuilder, pivotSnapshotValue } from '../src/pivot-cache-data.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
import { makeImportedPivotSourceReference, pivotSourceReferenceCurrent } from '../src/pivot-source-reference.js';
import { writeXlsx, readXlsx, readXlsxAsync } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child } from '../src/xml.js';
const rows = [['Group','Amount','Flag'],['A',10,true],['B',20,false],['',0,null]];
function fixture({external=true, table=false, local=false, saveData=true, count=1, compact=true}={}) {
  const b = new PivotSnapshotBuilder(rows[0]); rows.slice(1).forEach(row=>b.add(row));
  const saved=compact?b.finish():rows;
  const source = table ? {name:'MissingTable',ref:null,sheet:null} : {name:null,ref:'A1:C1048576',sheet:'Source'};
  if(external)source.external='file:///D:/Synthetic/Source%20%26%20Data.xlsx';
  const def={name:'ExternalPivot',...(table?{table:source.name}:{source:'Source',range:{r1:0,c1:0,r2:1048575,c2:2}}),
    rows:['Group'],cols:[],pages:['Flag'],values:[{field:'Amount',agg:'sum'}],top:3,left:4,area:{r1:3,c1:4,r2:10,c2:5},
    cacheItemsId:'saved',snapshotId:'saved',...(saveData?{}:{saveData:false})};
  def.sourceReference=makeImportedPivotSourceReference(source,def);
  const report={name:'Report',cells:{'3,4':{raw:'Saved label'},'9,5':{raw:'30',style:{bold:true}}},pivot:def,pivotsExtra:[]};
  for(let i=1;i<count;i++)report.pivotsExtra.push({...def,name:'ExternalPivot'+i,top:20*i,area:{r1:20*i,c1:4,r2:20*i+7,c2:5}});
  const sourceSheet={name:'Source',cells:{'0,0':{raw:'Group'},'0,1':{raw:'Amount'},'0,2':{raw:'Flag'},'1,0':{raw:'Local'},'1,1':{raw:'999'},'1,2':{raw:'FALSE'}}};
  if(table)sourceSheet.tables=[{name:'MissingTable',r1:0,c1:0,r2:1,c2:2,header:true,columns:['Group','Amount','Flag']}];
  return new Workbook({pivotSnapshots:{saved},pivotCacheItems:{saved:{fields:rows[0].map((name,c)=>({name,shared:rows.slice(1).map(r=>r[c])}))}},sheets:[...(local?[sourceSheet]:[]),report]});
}
const report = wb => wb.sheets.find(s=>s.name==='Report');
const source = wb => pivotSourceData(wb,report(wb).pivot);
const rt = wb => new Workbook(readXlsx(writeXlsx(wb)).data);
const tableCount = files => Object.keys(files).filter(n=>/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n)).length;
const grid = wb => {const d=report(wb).pivot,s=source(wb),r=resolvePivot(s,d);return computePivot(r,r.def).grid.map(row=>row.map(c=>c.raw));};
for(const compact of [false,true])test('missing external source uses imported saved records and preserves XML external relationship, compact='+compact,()=>{
  const wb=fixture({compact}),s=source(wb); assert.equal(s.cacheOnly,true);assert.equal(s.ref,null);assert.equal(s.si,undefined);assert.deepEqual(s.rows,rows);
  const bytes=writeXlsx(wb),files=unzip(bytes),after=new Workbook(readXlsx(bytes).data);
  assert.equal(tableCount(files),1);assert.deepEqual(source(after).rows,rows);assert.equal(report(after).pivot.sourceReference.external,report(wb).pivot.sourceReference.external);
  const rel=textOf(files['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels']);assert.match(rel,/rCacheRecords/);assert.match(rel,/rExternalSource/);assert.match(rel,/externalLinkPath/);assert.match(rel,/TargetMode="External"/);
  assert.match(textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']),/ref="A1:C1048576" sheet="Source" r:id="rExternalSource"/);
});
test('missing worksheet without external link and missing named table preserve saved records',()=>{
  for(const options of [{external:false},{external:false,table:true},{external:true,table:true}]){
    const wb=fixture(options);assert.deepEqual(source(wb).rows,rows);const after=rt(wb);assert.deepEqual(source(after).rows,rows);assert.equal(source(after).cacheOnly,true);
    if(options.table)assert.equal(report(after).pivot.table,'MissingTable');
  }
});
test('an external source never reads or changes its same-name local worksheet or table',()=>{
  for(const table of [false,true]){
    const wb=fixture({local:true,table}),def=report(wb).pivot,before=structuredClone(def.sourceReference);assert.deepEqual(source(wb).rows,rows);assert.equal(wb.pivotSnapshotSource(def),null);
    wb.transact(()=>wb.setInput(0,1,1,'888'));assert.deepEqual(source(wb).rows,rows);
    wb.transact(()=>wb.insertRows(0,0,1));assert.deepEqual(report(wb).pivot.sourceReference,before);assert.deepEqual(source(wb).rows,rows);
    wb.transact(()=>wb.renameSheet(0,'Renamed'));assert.deepEqual(report(wb).pivot.sourceReference,before);assert.deepEqual(source(wb).rows,rows);
  }
});
test('source/table/range or imported cache identity changes reject unrelated saved records',()=>{
  const patches=[{source:'Changed'},{range:{r1:1,c1:0,r2:20,c2:2}},{table:'ChangedTable'},{cacheItemsId:'other'},{snapshotId:'other'}];
  for(const patch of patches){const wb=fixture(),def={...report(wb).pivot,...patch};wb.pivotSnapshots.set('other',wb.pivotSnapshots.get('saved'));assert.equal(pivotSourceReferenceCurrent(def),false);assert.equal(pivotSourceData(wb,def,{preserveSnapshot:true}),null);}
});
test('a changed source with a real local worksheet uses live values, not the old external snapshot',()=>{
  const wb=fixture({local:true}),def={...report(wb).pivot,source:'Live'};wb.transact(()=>wb.renameSheet(0,'Live'));
  const src=pivotSourceData(wb,def);assert.equal(src.cacheOnly,undefined);assert.equal(src.rows[1][1],999);assert.equal(wb.pivotSnapshots.has('saved'),true);assert.equal(source(wb).rows[1][1],10);
});
test('when an originally missing local source becomes available its current values win',()=>{
  const wb=fixture({external:false,local:true}),s=source(wb);assert.equal(s.cacheOnly,undefined);assert.equal(s.rows[1][1],999);assert.equal(wb.pivotSnapshots.has('saved'),true);assert.equal(Object.hasOwn(wb.snapshotData(),'saved'),false);
});
test('normal local cache remains unmarked and uses its existing source-edit watch',()=>{
  const wb=fixture({external:false,local:true});delete report(wb).pivot.sourceReference;wb.setSnapshots({pivotSnapshots:{saved:rows}});assert.equal(source(wb).rows[1][1],10);
  const imported=rt(wb);assert.equal(report(imported).pivot.sourceReference,undefined);
  wb.transact(()=>wb.setInput(0,1,1,'777'));assert.equal(source(wb).rows[1][1],777);
});
test('1m-row 59/35-column compact tails are read without row materialization',()=>{
  for(const columns of [59,35]){
    const wb=fixture(),head=Array.from({length:columns},(_,i)=>'F'+i),snapshot={kind:'pivot-cache',version:1,header:head,n:1048575,columns:head.map((_,i)=>({num:null,str:null,dict:[],tail:{start:0,value:i}}))};
    wb.setSnapshots({pivotSnapshots:{saved:snapshot}});const src=source(wb);assert.equal(src.cube.n,1048575);assert.deepEqual(src.cube.header,head);
    assert.equal(src.cube.col(columns-1).get(1048574),columns-1);assert.equal(pivotSnapshotValue(wb.pivotSnapshots.get('saved').rows,1048574,columns-1),columns-1);
  }
});
test('shared cached pivots, filters, calculated fields and Undo/Redo survive roundtrip',()=>{
  const wb=fixture({count:20}),sh=report(wb);sh.pivot={...sh.pivot,calcFields:[{name:'Double',formula:'Amount*2'}],values:[{field:'Amount',agg:'sum'},{field:'Double',agg:'sum'}]};
  const initial=grid(wb);wb.transact(()=>wb.setSheetProp(wb.sheets.indexOf(sh),'pivot',{...sh.pivot,filters:{Group:['A']}}));const filtered=grid(wb);assert.notDeepEqual(filtered,initial);
  wb.undo();assert.deepEqual(grid(wb),initial);wb.redo();assert.deepEqual(grid(wb),filtered);
  const files=unzip(writeXlsx(wb)),after=rt(wb);assert.equal(tableCount(files),20);assert.equal(after.pivotSnapshots.size,1);assert.deepEqual(grid(after),filtered);assert.deepEqual(report(after).pivot.filters.Group,['A']);
});
test('different external file bindings remain separate even when their values and source names match',()=>{
  const wb=fixture(),sh=report(wb),other={...sh.pivot,name:'Other',cacheItemsId:'other',snapshotId:'other',top:30};
  other.sourceReference=makeImportedPivotSourceReference({...sh.pivot.sourceReference,external:'file:///D:/Synthetic/Other.xlsx'},other);sh.pivotsExtra=[other];
  wb.pivotSnapshots.set('other',{rows:wb.pivotSnapshots.get('saved').rows});wb.pivotCacheItems.other=wb.pivotCacheItems.saved;
  const after=rt(wb);assert.equal(after.pivotSnapshots.size,2);assert.equal(report(after).pivotsExtra[0].sourceReference.external,'file:///D:/Synthetic/Other.xlsx');
});
test('saveData=false second roundtrip preserves definitions, layout, cached output and chart data without UI empty calculation',()=>{
  const wb=fixture({saveData:false});report(wb).charts=[{id:'chart',type:'column',x:300,y:20,w:300,h:200,range:{r1:3,c1:4,r2:9,c2:5},series:[{name:'Saved',values:[10,20,0],categories:['A','B','']}]}];
  const first=writeXlsx(wb),a=new Workbook(readXlsx(first).data),src=source(a);assert.equal(src,null);assert.equal(a.pivotSnapshots?.size??0,0);
  const e=pivotSourceData(a,report(a).pivot,{allowCacheMetadata:true});assert.equal(e.metadataOnly,true);assert.equal(e.cube.n,0);assert.deepEqual(e.cube.header,rows[0]);
  const second=writeXlsx(a),b=new Workbook(readXlsx(second).data),af=unzip(first),bf=unzip(second),pa=parseXml(textOf(af['xl/pivotTables/pivotTable1.xml'])),pb=parseXml(textOf(bf['xl/pivotTables/pivotTable1.xml']));
  assert.equal(tableCount(bf),1);assert.equal(Object.keys(bf).some(n=>/pivotCacheRecords\d+\.xml$/.test(n)),false);
  for(const name of ['location','rowItems','colItems'])assert.deepEqual(child(pb,name),child(pa,name));
  assert.equal(b.getRaw(0,3,4),'Saved label');assert.equal(b.getRaw(0,9,5),'30');assert.deepEqual(report(b).charts[0].series,report(a).charts[0].series);
  assert.equal(source(b),null);assert.equal(report(b).pivot.sourceReference.external,report(a).pivot.sourceReference.external);
});
test('metadata-only layout is translated on output movement but field/filter changes fail before saving',()=>{
  const wb=rt(fixture({saveData:false})),sh=report(wb),location=sh.pivot.sourceReference.pivotLayout.location.ref;
  sh.pivot={...sh.pivot,top:sh.pivot.top+5,left:sh.pivot.left+2};const moved=rt(wb);assert.notEqual(report(moved).pivot.sourceReference.pivotLayout.location.ref,location);
  for(const patch of [{rows:['Flag']},{filters:{Group:['A']}},{calcFields:[{name:'Double',formula:'Amount*2'}]}]){const changed=rt(fixture({saveData:false}));report(changed).pivot={...report(changed).pivot,...patch};assert.throws(()=>writeXlsx(changed),/原本|원본.*레코드/);}
});
test('unavailable imported cache and incompatible cache headers cause useful errors instead of dropping pivots',()=>{
  const wb=fixture();wb.pivotSnapshots.clear();assert.throws(()=>writeXlsx(wb),/원본.*저장 캐시/);
  const other=fixture();other.setSnapshots({pivotSnapshots:{saved:[['Different'],[10]]}});assert.throws(()=>writeXlsx(other),/캐시 필드.*배치 필드/);
});
test('streamed XLSX import retains cached external source origin and records',async()=>{
  const bytes=writeXlsx(fixture()),loaded=await readXlsxAsync(bytes,undefined,{streamThreshold:0}),wb=new Workbook(loaded.data);assert.equal(source(wb).cacheOnly,true);assert.deepEqual(source(wb).rows,rows);assert.equal(pivotSourceReferenceCurrent(report(wb).pivot),true);
});

test('unrelated cache identity cannot fall through to a same-name local external source',()=>{
  for(const patch of [{cacheItemsId:'other'},{snapshotId:'other'}]){
    const wb=fixture({local:true}),def={...report(wb).pivot,...patch};wb.pivotSnapshots.set('other',wb.pivotSnapshots.get('saved'));
    assert.equal(pivotSourceData(wb,def,{preserveSnapshot:true}),null);assert.equal(wb.pivotSnapshotSource(def),null);
    assert.equal(pivotSourceData(wb,def),null);assert.equal(wb.pivotSnapshots.has('other'),true);assert.equal(wb.pivotSnapshots.has('saved'),true);
  }
});
test('metadata-only layout binding uses final normalized off-axis filters and slicer links',()=>{
  for(const linked of [false,true]){
    const wb=fixture({saveData:false}),sh=report(wb);sh.pivot.filters={Amount:['10']};
    if(linked)sh.slicers=[{id:'amount',caption:'Amount',source:{kind:'pivot',field:'Amount',self:true},x:300,y:0,w:160,h:200,showDeleted:true}];
    const loaded=rt(wb);assert.equal(!!report(loaded).pivot.filters?.Amount,linked);assert.doesNotThrow(()=>writeXlsx(loaded));
  }
});


test('changing one imported source preserves other shared owners through filters, JSON, save and Undo',()=>{
  const wb=fixture({count:2,local:true}),sh=report(wb),si=wb.sheets.indexOf(sh),saved=wb.pivotSnapshots.get('saved');
  const cached = book => pivotSourceData(book,report(book).pivotsExtra[0]);
  wb.transact(()=>wb.setSheetProp(si,'pivot',{...sh.pivot,source:'Live'}));
  wb.transact(()=>wb.renameSheet(0,'Live'));
  assert.equal(source(wb).rows[1][1],999);assert.deepEqual(cached(wb).rows,rows);assert.equal(wb.pivotSnapshots.get('saved'),saved);
  assert.equal(wb.snapshotData().saved,saved.rows);
  wb.transact(()=>wb.setSheetProp(si,'pivotsExtra',[{...sh.pivotsExtra[0],filters:{Group:['A']}}]));
  const filtered=book=>{const d=report(book).pivotsExtra[0],r=resolvePivot(cached(book),d);return computePivot(r,r.def).grid.map(row=>row.map(c=>c.raw));};
  const expected=filtered(wb);assert.ok(expected.some(row=>row.some(value=>value==='10')));
  const json=new Workbook(JSON.parse(JSON.stringify(wb.serialize())));assert.equal(source(json).rows[1][1],999);assert.deepEqual(cached(json).rows,rows);assert.deepEqual(filtered(json),expected);
  const after=rt(wb);assert.equal(source(after).rows[1][1],999);assert.deepEqual(cached(after).rows,rows);assert.deepEqual(filtered(after),expected);assert.equal(tableCount(unzip(writeXlsx(wb))),2);
  wb.undo();assert.equal(sh.pivotsExtra[0].filters,undefined);assert.deepEqual(cached(wb).rows,rows);
  wb.undo();wb.undo();assert.equal(sh.pivot.source,'Source');assert.deepEqual(source(wb).rows,rows);assert.deepEqual(cached(wb).rows,rows);
  wb.redo();wb.redo();wb.redo();assert.equal(source(wb).rows[1][1],999);assert.deepEqual(cached(wb).rows,rows);assert.deepEqual(filtered(wb),expected);
});

test('shared imported ownership does not override ordinary local edit invalidation',()=>{
  const wb=fixture({count:2,local:true}),sh=report(wb);
  sh.pivot={...sh.pivot};delete sh.pivot.sourceReference;
  wb.setSnapshots({pivotSnapshots:{saved:rows}});assert.equal(source(wb).rows[1][1],10);assert.deepEqual(pivotSourceData(wb,sh.pivotsExtra[0]).rows,rows);
  wb.transact(()=>wb.setInput(0,1,1,'777'));assert.equal(Object.hasOwn(wb.snapshotData(),'saved'),false);
  assert.equal(source(wb).rows[1][1],777);assert.equal(wb.pivotSnapshots.has('saved'),false);
});


test('metadata-only structural and missing-item policy changes fail before saving old item indices',()=>{
  for(const patch of [{classic:true},{blankRows:true},{subtotalTop:false},{repeatLabels:true},{missingItems:'none'}]){
    const wb=rt(fixture({saveData:false})),sh=report(wb),cells=JSON.stringify(wb.serialize().sheets[0].cells),reference=structuredClone(sh.pivot.sourceReference);
    sh.pivot={...sh.pivot,...patch};assert.throws(()=>writeXlsx(wb),/원본.*레코드.*배치/);
    assert.equal(JSON.stringify(wb.serialize().sheets[0].cells),cells);assert.deepEqual(sh.pivot.sourceReference,reference);
  }
});

test('external saved records survive refresh-on-open without fetching their source',async()=>{
  const wb=fixture();report(wb).pivot={...report(wb).pivot,refreshOnOpen:true};
  const bytes=writeXlsx(wb);
  for(const loaded of [readXlsx(bytes),await readXlsxAsync(bytes,undefined,{streamThreshold:0})]){
    const after=new Workbook(loaded.data);assert.equal(report(after).pivot.refreshOnOpen,true);assert.deepEqual(source(after).rows,rows);assert.equal(source(after).cacheOnly,true);
    assert.equal(after.pivotSnapshotSource(report(after).pivot),null);const twice=rt(after);assert.deepEqual(source(twice).rows,rows);assert.equal(report(twice).pivot.refreshOnOpen,true);
  }
});

test('normal local refresh-on-open and external saveData=false policies retain their prior meaning',()=>{
  const wb=fixture({external:false,local:true});delete report(wb).pivot.sourceReference;report(wb).pivot={...report(wb).pivot,refreshOnOpen:true};
  wb.setSnapshots({pivotSnapshots:{saved:rows}});assert.equal(source(wb).rows[1][1],10);
  const local=rt(wb);assert.equal(local.pivotSnapshots?.size??0,0);assert.equal(source(local).rows[1][1],999);assert.equal(report(local).pivot.sourceReference,undefined);
  const external=fixture({saveData:false});report(external).pivot={...report(external).pivot,refreshOnOpen:true};const after=rt(external);
  assert.equal(after.pivotSnapshots?.size??0,0);assert.equal(source(after),null);assert.equal(report(rt(after)).pivot.saveData,false);
  const missing=fixture({external:false});report(missing).pivot={...report(missing).pivot,refreshOnOpen:true};const missingAfter=rt(missing);
  assert.equal(missingAfter.pivotSnapshots?.size??0,0);assert.equal(source(missingAfter),null);
});
