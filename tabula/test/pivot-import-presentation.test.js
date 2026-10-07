import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';
import { makeImportedPivotSourceReference } from '../src/pivot-source-reference.js';
import { pivotImportedPresentation, pivotImportedPresentationCurrent, importedPivotPresentationCurrent, pivotImportedOwnsCell, pivotImportedClearAreas, pivotImportedButtons, pivotCellButtons, captureImportedPivotFormats, pivotItemOrderSignature } from '../src/pivot-import-presentation.js';
import { CellMap } from '../src/cellmap.js';
import { ColBuilder } from '../src/block.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child } from '../src/xml.js';
const te=new TextEncoder();
const monthBytes=(external=false)=>{
  const def={name:'Saved month',source:'Source',range:{r1:0,c1:0,r2:1,c2:1},rows:['Date'],cols:[],pages:[],values:[{field:'Amount',agg:'sum'}],top:4,left:5,layout:'tabular',grandRows:false,grandCols:false,subtotals:false,autofit:false};
  const data={sheets:[{name:'Source',cells:{'0,0':{raw:'Date'},'0,1':{raw:'Amount'},'1,0':{raw:'44136'},'1,1':{raw:'134'}}},{name:'Report',cells:{'4,5':{raw:'Date'},'4,6':{raw:'Sum'},'5,5':{raw:'44136',style:{numFmt:'date'}},'5,6':{raw:'134',style:{bold:true,color:'#123456'}}},pivot:def,colWidths:{5:173,6:99},rowHeights:{5:31},merges:[],cond:[{r1:5,c1:6,r2:5,c2:6,type:'cell',op:'greater',v1:'100',style:{fill:'#00ff00'}}] }]};
  if(external){def.cacheItemsId='saved';def.snapshotId='saved';def.sourceReference=makeImportedPivotSourceReference({ref:'A1:B2',sheet:'Source',external:'file:///D:/Synthetic/Missing.xlsx'},def);data.pivotSnapshots={saved:[['Date','Amount'],[44136,134]]};}
  const files=unzip(writeXlsx(new Workbook(data))),p='xl/pivotTables/pivotTable1.xml';
  files[p]=te.encode(textOf(files[p]).replace('</pivotTableDefinition>','<filters count="1"><filter fld="0" type="thisMonth" id="0"><autoFilter ref="A1"><filterColumn colId="0"><dynamicFilter type="thisMonth" val="44136"/></filterColumn></autoFilter></filter></filters></pivotTableDefinition>'));
  return zip(files);
};
const load=bytes=>new Workbook(readXlsx(bytes).data);
const defOf=w=>w.sheets[1].pivot;
const layouts=bytes=>{const node=parseXml(textOf(unzip(bytes)['xl/pivotTables/pivotTable1.xml']));return Object.fromEntries(['location','rowItems','colItems'].map(k=>[k,child(node,k)]));};
for(const external of [false,true])test('이전달 thisMonth 저장 결과·위치·항목은 두 번 저장해도 원자적으로 보존된다, external='+external,()=>{
  const bytes=monthBytes(external),w=load(bytes),d=defOf(w),before=JSON.stringify(w.sheets[1].cells);
  const prepared=pivotImportedPresentation(w,1,d);assert.ok(prepared);assert.equal(prepared.plan.grid.length,1);assert.equal(prepared.plan.grid.some(row=>row.some(cell=>cell?.raw==='134')),false);
  const s=pivotSourceData(w,d),res=resolvePivot(s,d),current=computePivot(res,res.def);assert.equal(current.grid.some(row=>row.some(c=>c?.raw==='134')),false,'현재 월 계산은 저장된 이전달 값과 다르다');
  assert.equal(JSON.stringify(w.sheets[1].cells),before,'초기 준비는 원본 셀을 쓰지 않는다');
  assert.equal(pivotImportedButtons(w.sheets[1],d,prepared.plan).some(b=>b.kind==='rows'&&b.r===4&&b.c===5),true);
  const second=writeXlsx(w),a=load(second),third=writeXlsx(a),b=load(third);assert.deepEqual(layouts(second),layouts(bytes));assert.deepEqual(layouts(third),layouts(bytes));
  for(const wb of[a,b]){assert.equal(wb.getValue(1,5,5),44136);assert.equal(wb.getValue(1,5,6),134);assert.deepEqual(defOf(wb).area,d.area);assert.equal(wb.sheets[1].colWidths[5],w.sheets[1].colWidths[5]);assert.equal(wb.sheets[1].rowHeights[5],31);}
  if(!external)assert.equal(d.sourceReference,undefined,'로컬 캐시는 외부 원본으로 표시하지 않는다');
});
test('동기/stream/Blob reader는 초기 query·원본 location·캐시 식별 메타를 같은 의미로 준비한다',async()=>{
  const bytes=monthBytes();const a=load(bytes),b=new Workbook((await readXlsxAsync(bytes,undefined,{streamThreshold:0})).data),c=new Workbook((await readXlsxAsync(new Blob([bytes]),undefined,{streamThreshold:0})).data);
  for(const w of[b,c]){assert.deepEqual(defOf(w).importedPresentation,defOf(a).importedPresentation);assert.ok(pivotImportedPresentation(w,1,defOf(w)));}
});
test('새로고침/rebuild/원본 변경/캐시 교체/최초 query 변경은 원본 화면 보존을 거절한다',()=>{
  for(const patch of[{refreshOnOpen:true},{needsRender:true},{captureFmt:false},{filters:{Date:['44136']}},{snapshotId:'other'},{cacheItemsId:'other'},{source:'Changed'},{values:[{field:'Amount',agg:'count'}]}]){
    const w=load(monthBytes()),d=defOf(w);w.pivotSnapshots.set('other',w.pivotSnapshots.get(d.snapshotId));assert.equal(pivotImportedPresentation(w,1,{...d,...patch}),null,JSON.stringify(patch));
  }
  const w=load(monthBytes()),d=defOf(w);w.transact(()=>w.setInput(0,1,1,'999'));assert.equal(pivotImportedPresentation(w,1,d),null);assert.equal(pivotSourceData(w,d).cube.col(1).get(0),999);
  const empty=load(monthBytes());empty.pivotSnapshots.clear();assert.equal(pivotImportedPresentation(empty,1,defOf(empty)),null);
});
test('원본 본문 없는 새 피벗은 초기 보존 대상이 아니다',()=>{
  const w=load(monthBytes()),d=defOf(w);w.sheets[1].cells.deleteRC(5,5);w.sheets[1].cells.deleteRC(5,6);assert.equal(pivotImportedPresentation(w,1,d),null);
});
test('전체 피벗 이동은 상대 배치·버튼을 유지하지만 본문/필터 사이 삽입은 오래된 배치를 거절한다',()=>{
  const w=load(monthBytes()),d=defOf(w);w.transact(()=>w.insertRows(1,0,2));const moved=defOf(w),p=pivotImportedPresentation(w,1,moved);assert.ok(p);assert.equal(pivotImportedButtons(w.sheets[1],moved,p.plan)[0].r,6);
  const bytes=writeXlsx(w);assert.equal(layouts(bytes).location.attrs.ref,'F7:G8');assert.equal(layouts(bytes).rowItems.attrs.count,'1');
  const invalid={...moved,area:{...moved.area,r2:moved.area.r2+1}};assert.equal(importedPivotPresentationCurrent(invalid),false);assert.equal(pivotImportedPresentation(w,1,invalid),null);
});
test('원본 마커는 JSON/Undo clone과 좌표 이동을 허용하고 바뀐 query/area/source는 거절한다',()=>{
  const w=load(monthBytes()),d=defOf(w),marker=pivotImportedPresentation(w,1,d).marker;
  const copy=JSON.parse(JSON.stringify(d));assert.equal(pivotImportedPresentationCurrent(copy,marker),true);
  assert.equal(pivotImportedPresentationCurrent({...copy,top:copy.top+2,area:{...copy.area,r1:copy.area.r1+2,r2:copy.area.r2+2}},marker),true);
  assert.equal(pivotImportedPresentationCurrent({...copy,filters:{Date:['44136']}},marker),false);
  assert.equal(pivotImportedPresentationCurrent({...copy,area:{...copy.area,r2:copy.area.r2+1}},marker),false);
  assert.equal(pivotImportedPresentationCurrent({...copy,source:'Different'},marker),false);
  assert.equal(pivotImportedPresentationCurrent({...copy,captureFmt:false},marker),false,'정상 쓰기가 완료된 새 정의는 역할 캐시를 다시 사용할 수 있다');
  assert.equal(pivotImportedPresentationCurrent(JSON.parse(JSON.stringify(d)),marker),true,'첫 query 뒤 삭제하지 않는 마커는 Undo 후 다시 원본을 인식한다');
});
test('원본 소유권은 본문과 page 이름/값만 포함하고 gap 사용자 셀은 제외한다',()=>{
  const d={top:3,left:4,area:{r1:3,c1:4,r2:12,c2:9}},m={bodyRow:4,bodyColumnStart:0,bodyColumnEnd:1,pageFields:[{r:0,c:0},{r:0,c:3},{r:1,c:0}]};
  for(const p of[[3,4],[3,5],[3,7],[3,8],[7,4],[12,5]])assert.equal(pivotImportedOwnsCell(d,m,...p),true);
  for(const p of[[3,6],[4,7],[6,4],[9,8],[13,4]])assert.equal(pivotImportedOwnsCell(d,m,...p),false);
});
test('서식 투표는 원본 bounds와 머리글 구간에 한정하고 외부 강조 서식이 번지지 않는다',()=>{
  const cells=new CellMap();cells.setRC(4,5,{raw:'Header',style:{color:'#123456'}});cells.setRC(5,5,{raw:'1',style:{numFmt:'number',decimals:2}});cells.setRC(6,5,{raw:'9',style:{numFmt:'percent',bold:true}});
  const plan={top:4,left:5,pm:{pageRows:0,headerRows:1},grid:[[{raw:'Header',role:'rowHead:0',style:{}}],[{raw:'1',role:'data:0',style:{numFmt:'comma'}}],[{raw:'9',role:'data:0',style:{numFmt:'comma'}}]]},d={};
  const head=captureImportedPivotFormats({cells},d,plan,{}, {area:{r1:4,c1:5,r2:5,c2:5},headersOnly:true});assert.equal(head['rowHead:0'].color,'#123456');assert.equal(head['data:0'],undefined);
  const all=captureImportedPivotFormats({cells},d,plan,{}, {area:{r1:4,c1:5,r2:5,c2:5}});assert.equal(all['data:0'].numFmt,'number');assert.equal(all['data:0'].decimals,2);assert.equal(all['data:0'].bold,undefined);
});
test('기존 헤더/page/Σ/toggle 버튼 판정은 절대 좌표에서도 동일하다',()=>{
  const d={rows:['Group','Detail'],cols:['Region','Device'],pages:['Page'],values:[{field:'A'},{field:'B'}],layout:'tabular',valuesPos:1};const pm={top:4,left:5,labelCols:2};
  assert.deepEqual(pivotCellButtons({role:'rowHead:1'},8,6,pm,d),[{r:8,c:6,kind:'rows',field:'Detail'}]);
  assert.deepEqual(pivotCellButtons({role:'colHead',raw:'Values'},4,8,pm,d),[{r:4,c:8,kind:'cols',sigma:true}]);
  assert.deepEqual(pivotCellButtons({role:'colHead',raw:'Device'},4,9,pm,d),[{r:4,c:9,kind:'cols',field:'Device'}]);
  assert.deepEqual(pivotCellButtons({role:'pageValue',field:'Page'},4,6,pm,d),[{r:4,c:6,kind:'page',field:'Page'}]);
  assert.equal(pivotCellButtons({role:'rowItem:0',toggle:{field:'Group',item:'A'}},8,5,pm,d).length,1);
  assert.equal(pivotCellButtons({role:'rowItem:0',toggle:{field:'Group',item:'A'}},8,5,pm,d,{includeToggle:false}).length,0);
});
test('필드 항목 순번 서명은 타입/빈 값/특수항목/순서를 모두 구분한다',()=>{
  const sig=pivotItemOrderSignature;
  for(const pair of[[['v:number:1'],['v:string:1']],[['v:string:(비어 있음)'],['v:string:\u0000""']],[['v:string:a','t:default','v:string:b'],['v:string:a','v:string:b','t:default']],[['v:string:a','v:string:b'],['v:string:b','v:string:a']]])assert.notEqual(sig(pair[0]),sig(pair[1]));
  assert.equal(sig(['v:number:44136']),sig(['v:number:44136']),'n 날짜와 d에서 읽은 일련번호는 같은 의미이다');
});
test('items 또는 축 순번이 달라지면 location과 rowItems/colItems를 함께 현재 계산으로 쓴다',()=>{
  for(const change of [p=>p.itemOrder[0].signature='different',p=>p.axes.rows=['@values','Date'],p=>p.axes.cols=['@values']]){
    const w=load(monthBytes()),d=defOf(w);change(d.importedPresentation);const result=layouts(writeXlsx(w));assert.equal(result.location.attrs.ref,'F5:G5');assert.equal(result.rowItems,null);
  }
});

test('블록에만 저장된 원본 본문/머리글도 계산 없이 준비하고 첫 query tail을 Undo로 정확히 복원한다',()=>{
  const w=load(monthBytes()),d=defOf(w),sh=w.sheets[1],a=new ColBuilder(),b=new ColBuilder();
  a.set(0,'Date');a.set(1,44136);b.set(0,'Sum');b.set(1,134);
  sh.blocks=[{r0:4,c0:5,n:2,cols:[a.finish(2,{numFmt:'date'}),b.finish(2,{bold:true,color:'#123456'})]}];sh.cells.clear();
  const initial=pivotImportedPresentation(w,1,d);assert.ok(initial);assert.deepEqual(pivotImportedButtons(sh,d,initial.plan),[{r:4,c:5,kind:'rows',field:'Date'}]);
  assert.equal(initial.plan.grid[0][0].raw,'Date');assert.equal(w.getValue(1,5,6),134);assert.equal(sh.cells.size,0);
  const src=pivotSourceData(w,d),resolved=resolvePivot(src,d),computed=computePivot(resolved,resolved.def);
  const nextArea={r1:4,c1:5,r2:4,c2:6},plan={...computed,pm:computed.meta,top:4,left:5,bodyColsN:2,area:nextArea};
  const areas=pivotImportedClearAreas(d,initial.marker,plan);assert.deepEqual(areas,[{r1:5,c1:5,r2:5,c2:6}]);
  const run=()=>w.transact(()=>{for(const a of areas)w.clearRange(1,a.r1,a.c1,a.r2,a.c2,'contents');w.setSheetProp(1,'pivot',{...defOf(w),area:nextArea,captureFmt:false});});
  run();assert.equal(w.getValue(1,5,5),null);assert.equal(w.getValue(1,5,6),null);assert.equal(w.styleAt(1,5,6).color,'#123456');
  w.undo();assert.equal(w.getValue(1,5,5),44136);assert.equal(w.getValue(1,5,6),134);assert.equal(pivotImportedPresentationCurrent(defOf(w),initial.marker),true);assert.ok(pivotImportedPresentation(w,1,defOf(w)));
  run();assert.equal(w.getValue(1,5,6),null);w.undo();assert.equal(w.getValue(1,5,6),134);w.redo();assert.equal(w.getValue(1,5,6),null);
});
test('원본 body/page slots에서 새 결과가 쓰지 않는 영역만 정리하며 gap sentinel은 제외한다',()=>{
  const d={top:3,left:4,area:{r1:3,c1:4,r2:12,c2:8}},m={bodyRow:4,bodyColumnStart:0,bodyColumnEnd:1,pageFields:[{r:0,c:0},{r:0,c:3}]};
  const plan={top:3,left:4,grid:Array.from({length:7},()=>[]),pm:{pageRows:4,pageFields:[{r:0,c:0}]},bodyColsN:2,area:{r1:3,c1:4,r2:9,c2:5}};
  const areas=pivotImportedClearAreas(d,m,plan);assert.deepEqual(areas,[{r1:10,c1:4,r2:12,c2:5},{r1:3,c1:7,r2:3,c2:8}]);
  assert.equal(areas.some(a=>6>=a.r1&&6<=a.r2&&4>=a.c1&&4<=a.c2),false);assert.equal(areas.some(a=>3>=a.r1&&3<=a.r2&&6>=a.c1&&6<=a.c2),false);
});

test('실제 XML의 n 날짜→d 날짜 표현 변경은 항목 의미·원문 배치 보존을 허용한다',()=>{
  const files=unzip(monthBytes()),path='xl/pivotCache/pivotCacheDefinition1.xml';
  const xml=textOf(files[path]);assert.match(xml,/<n v="44136"\/>/);
  files[path]=te.encode(xml.replace('<n v="44136"/>','<d v="2020-11-01T00:00:00"/>'));
  const bytes=zip(files),w=load(bytes);assert.equal(pivotSourceData(w,defOf(w)).cube.col(0).get(0),44136);
  assert.deepEqual(layouts(writeXlsx(w)),layouts(bytes));
});
test('공유 cache x가 재번호되어도 pivotField item 의미/순번이 같으면 원문 rowItems를 보존한다',()=>{
  const files=unzip(monthBytes()),cache='xl/pivotCache/pivotCacheDefinition1.xml',pivot='xl/pivotTables/pivotTable1.xml',records='xl/pivotCache/pivotCacheRecords1.xml';
  let xml=textOf(files[cache]);xml=xml.replace(/<sharedItems([^>]*)count="1"([^>]*)><n v="44136"\/><\/sharedItems>/,'<sharedItems$1count="2"$2><n v="999"/><n v="44136"/></sharedItems>');
  assert.match(xml,/<n v="999"\/><n v="44136"\/>/);files[cache]=te.encode(xml);
  files[pivot]=te.encode(textOf(files[pivot]).replace(/<items count="1"><item x="0"\/><\/items>/,'<items count="2"><item x="1"/><item x="0"/></items>'));
  files[records]=te.encode(textOf(files[records]).replace('<r><x v="0"/>','<r><x v="1"/>'));
  const bytes=zip(files),w=load(bytes);assert.equal(pivotSourceData(w,defOf(w)).cube.col(0).get(0),44136);
  assert.deepEqual(layouts(writeXlsx(w)),layouts(bytes));
});


test('원문 sparse 토글은 reader·JSON·전체 이동에서 유지하고 수동 레이블 변경은 그 단추만 제외한다',()=>{
  const rows=[['Group','Detail','Amount'],['A','x',1],['A','y',2],['B','x',3]];
  const def={name:'Sparse buttons',source:'Source',range:{r1:0,c1:0,r2:3,c2:2},rows:['Group','Detail'],cols:[],pages:[],values:[{field:'Amount',agg:'sum'}],layout:'tabular',subtotals:false,grandRows:false,grandCols:false,top:3,left:4};
  const computed=computePivot(rows,resolvePivot(rows,def).def),source={},report={};
  rows.forEach((row,r)=>row.forEach((v,c)=>source[r+','+c]={raw:String(v)}));
  computed.grid.forEach((row,r)=>row.forEach((cell,c)=>{if(cell)report[(def.top+r)+','+(def.left+c)]={raw:String(cell.raw??''),style:cell.style};}));
  const bytes=writeXlsx(new Workbook({sheets:[{name:'Source',cells:source},{name:'Report',cells:report,pivot:def}]}));
  const wb=load(bytes),d=defOf(wb),prepared=pivotImportedPresentation(wb,1,d);
  assert.ok(prepared);assert.equal(d.importedPresentation.toggles.length,2);
  const onlyToggles=book=>{const d=defOf(book),p=pivotImportedPresentation(book,1,d);assert.ok(p);return pivotImportedButtons(book.sheets[1],d,p.plan).filter(b=>b.kind==='toggle');};
  const initial=onlyToggles(wb);assert.deepEqual(initial.map(b=>b.item),['A','B']);
  const json=new Workbook(JSON.parse(JSON.stringify(wb.serialize())));assert.deepEqual(onlyToggles(json),initial);
  wb.transact(()=>wb.insertRows(1,0,2));const shifted=onlyToggles(wb);assert.deepEqual(shifted,initial.map(b=>({...b,r:b.r+2})));
  const changed=shifted[0];wb.transact(()=>wb.setInput(1,changed.r,changed.c,'User label'));
  assert.deepEqual(onlyToggles(wb),[shifted[1]],'원문과 달라진 레이블에 오래된 그룹 단추를 붙이지 않는다');
  wb.undo();assert.deepEqual(onlyToggles(wb),shifted);wb.undo();assert.deepEqual(onlyToggles(wb),initial);
});
