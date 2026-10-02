import { test } from 'node:test';
import assert from 'node:assert/strict';
import { computePivot, resolvePivot, pivotDetail } from '../src/pivot.js';
import { pivotContextTarget, pivotValueDef, pivotRemoveContextField } from '../src/pivot-context.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
const source=[['지역','상품','매출','수량'],['서울','A',10,2],['서울','B',30,3],['부산','A',20,4]];
const base={rows:['지역'],cols:[],values:[{field:'매출',agg:'sum'},{field:'매출',agg:'average'},{field:'수량',agg:'sum'}],layout:'tabular'};
const result=def=>{const r=resolvePivot(source,def);return {def:r.def,out:computePivot(r,r.def)};};
test('값·부분합·총합계·값머리글 역할은 선택한 인덱스를 유지한다',()=>{
  for(const role of ['data','subData','groupData','grandData','grandColData','colSubData','valueHead','grandHead']){const c=pivotContextTarget(base,{grid:[[{role:role+':1'}]],meta:{}},0,0);assert.equal(c.valueIndex,1);assert.equal(c.field,'매출');assert.equal(c.kind,'value');}
});
test('값을 행에 배치한 레이블도 해당 값 필드만 선택한다',()=>{
  const {def,out}=result({...base,valuesOnRows:true});let found=0;
  out.grid.forEach((row,r)=>row.forEach((cell,c)=>{if(cell.role==='rowItem:1'){const t=pivotContextTarget(def,out,r,c);assert.equal(t.area,'values');assert.equal(t.valueIndex,out.meta.rowItems[r-out.meta.pageRows-out.meta.headerRows].vi);found++;}}));assert.ok(found>=6);
});
test('행·열·보고서 필터 머리글은 값 필드로 오인하지 않는다',()=>{
  const def={...base,cols:['상품'],pages:['수량']};for(const [role,area,field]of[['rowItem:0','rows','지역'],['rowHead:0','rows','지역'],['colItem:0','cols','상품'],['pageValue','pages','수량']]){const t=pivotContextTarget(def,{grid:[[{role}]],meta:{}},0,0);assert.equal(t.area,area);assert.equal(t.field,field);assert.equal(t.valueIndex,null);}
});
test('값 필드 변경은 알 수 없는 기존 속성·다른 값 서식을 보존한다',()=>{
  const d={...base,values:[{...base.values[0],name:'고유 이름',nativeOption:7,numFmt:{code:'#,##0'}},base.values[1]],cellFmt:{'data:0':{numFmt:'number',decimals:2,fill:'#abc123'},'data:1':{numFmt:'number',decimals:1}}};
  const n=pivotValueDef(d,0,{agg:'count',showAs:undefined});assert.equal(n.values[0].nativeOption,7);assert.equal(n.values[0].name,'고유 이름');assert.deepEqual(n.cellFmt,d.cellFmt);
  const f=pivotValueDef(d,0,{numFmt:{code:'0.00%'}});assert.deepEqual(f.cellFmt['data:0'],{fill:'#abc123'});assert.deepEqual(f.cellFmt['data:1'],d.cellFmt['data:1']);assert.equal(d.values[0].agg,'sum');
});
test('중복 값 필드 제거는 1개만 제거하고 정렬·필터·역할 인덱스를 이동한다',()=>{
  const d={...base,sort:{지역:{dir:'desc',by:2},상품:{dir:'asc',by:1}},fieldFilters:{지역:{type:'value',by:1},상품:{type:'top',by:2}},cellFmt:{'data:0':{a:0},'data:1':{a:1},'data:2':{a:2},'rowItem:0':{keep:true}}};
  const n=pivotRemoveContextField(d,{area:'values',valueIndex:1});assert.deepEqual(n.values,[base.values[0],base.values[2]]);assert.equal(n.sort.지역.by,1);assert.equal(n.sort.상품.by,undefined);assert.equal(n.fieldFilters.지역,undefined);assert.equal(n.fieldFilters.상품.by,1);assert.deepEqual(n.cellFmt['data:1'],{a:2});assert.equal(d.values.length,3);
});
test('레이블 제거는 해당 영역과 종속 필터만 정리한다',()=>{
  const d={...base,cols:['상품','수량'],valuesPos:1,filters:{지역:['서울'],상품:['A']},sort:{지역:{dir:'asc'}}};const n=pivotRemoveContextField(d,{area:'rows',field:'지역'});assert.deepEqual(n.rows,[]);assert.deepEqual(n.filters,{상품:['A']});assert.equal(n.sort.지역,undefined);assert.equal(d.rows.length,1);
});
test('행 배치 다중값 자세한 정보는 선택값 이름과 원본 행을 반환한다',()=>{
  for(const layout of ['compact','outline','tabular'])for(const rows of [['지역'],[]]){const {def,out}=result({...base,rows,valuesOnRows:true,layout});let found=0;out.grid.forEach((row,r)=>row.forEach((cell,c)=>{if(/^data:1$/.test(cell.role)){const d=pivotDetail(source,def,r,c);assert.ok(d);assert.equal(d.valueField,'평균 : 매출');assert.ok(d.idx.length);found++;}}));assert.ok(found);assert.equal(pivotDetail(source,def,999,1),null);}
});
test('값 행 배치의 XLSX 왕복 뒤에도 2번째 값 자세한 정보가 일치한다',()=>{
  const cells={};source.forEach((row,r)=>row.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));const d={...base,valuesOnRows:true,source:'원본',range:{r1:0,c1:0,r2:3,c2:3},name:'합성피벗',top:0,left:0};const w=new Workbook({sheets:[{name:'원본',cells},{name:'피벗',cells:{},pivot:d}]});const restored=new Workbook(readXlsx(writeXlsx(w)).data);const after=restored.sheets[1].pivot;assert.equal(after.valuesOnRows,true);assert.deepEqual(after.values.map(v=>[v.field,v.agg]),d.values.map(v=>[v.field,v.agg]));const {def,out}=result(after);let found=0;out.grid.forEach((row,r)=>row.forEach((cell,c)=>{if(cell.role==='data:1'){assert.equal(pivotDetail(source,def,r,c).valueField,'평균 : 매출');found++;}}));assert.ok(found);
});
