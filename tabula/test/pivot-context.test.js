import { test } from 'node:test';
import assert from 'node:assert/strict';
import { computePivot, resolvePivot, pivotDetail, pivotChartData } from '../src/pivot.js';
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


test('가로 보고서 필터 문맥은 행 번호 대신 셀의 실제 필드를 사용한다',()=>{
  const d={...base,pages:['지역','상품','수량'],pageOrder:'over',pageWrap:2};
  const {def,out}=result(d);
  for(const point of out.meta.pageFields){
    for(const c of [point.c,point.c+1]){
      const t=pivotContextTarget(def,out,point.r,c);
      assert.equal(t.area,'pages'); assert.equal(t.field,point.field);
      const next=pivotRemoveContextField(d,t);
      assert.deepEqual(next.pages,d.pages.filter(f=>f!==point.field));
      assert.deepEqual(next.values,d.values);
    }
  }
});

const targetCells = ({def, out}, role) => out.grid.flatMap((row,r)=>row.flatMap((cell,c)=>role.test(cell.role)?[{r,c,cell,target:pivotContextTarget(def,out,r,c)}]:[]));
const sortFromTarget = (def, target, dir='desc') => ({...def,sort:{...def.sort,[target.sortField]:{dir,by:target.valueIndex,...(target.sortAt?{at:target.sortAt}:{})}}});

test('값을 고른 열의 실제 항목 경로로 정렬하며 총합계와 피벗 차트가 일치한다',()=>{
  const data=[['품목','채널','매출'],['A','X',90],['A','Y',1],['B','X',10],['B','Y',200]];
  const d={rows:['품목'],cols:['채널'],values:[{field:'매출',agg:'sum'}],layout:'tabular',itemCaptions:{채널:{X:'선택 열'}}};
  const resolved=resolvePivot(data,d), out=computePivot(resolved,resolved.def);
  const col=out.meta.labelCols+out.meta.colLeaves.findIndex(x=>x.node.key==='X');
  const r=out.meta.pageRows+out.meta.headerRows;
  const target=pivotContextTarget(resolved.def,out,r,col);
  assert.equal(target.sortField,'품목');assert.equal(target.sortAxis,'rows');assert.deepEqual(target.sortAt,[['채널','X']]);
  const sorted=sortFromTarget(d,target), res=resolvePivot(data,sorted), after=computePivot(res,res.def);
  assert.deepEqual(after.meta.rowItems.filter(x=>x.kind==='item').map(x=>x.node.key),['A','B']);
  assert.equal(after.meta.rowItems.at(-1).kind,'grand');assert.equal(after.grid.at(-1).at(-1).raw,'301');
  const chart=pivotChartData(data,sorted);assert.deepEqual(chart.categories,['A','B']);assert.deepEqual(chart.series.map(x=>x.values),[[90,10],[1,200]]);
  const totalTarget=pivotContextTarget(resolved.def,out,r,out.meta.labelCols+out.meta.colLeaves.findIndex(x=>x.kind==='grand'));
  assert.equal(totalTarget.sortField,'품목');assert.equal(totalTarget.sortAt,undefined);
  const totalRes=resolvePivot(data,sortFromTarget(d,totalTarget));
  assert.deepEqual(computePivot(totalRes,totalRes.def).meta.rowItems.filter(x=>x.kind==='item').map(x=>x.node.key),['B','A']);
});

test('총합계 행의 값은 열 항목을 정렬하고 합계 교차점은 정렬 대상으로 삼지 않는다',()=>{
  const {def,out}=result({...base,cols:['상품'],values:[base.values[0]]});
  const r=out.grid.length-1, first=out.meta.labelCols;
  const t=pivotContextTarget(def,out,r,first);
  assert.equal(t.sortAxis,'cols');assert.equal(t.sortField,'상품');assert.equal(t.valueIndex,0);assert.equal(t.sortAt,undefined);
  const corner=pivotContextTarget(def,out,r,out.meta.width-1);
  assert.equal(corner.sortField,null);assert.equal(corner.sortAxis,null);
  const rowOnly=result({...base,values:[base.values[0]]});
  assert.equal(pivotContextTarget(rowOnly.def,rowOnly.out,rowOnly.out.grid.length-1,rowOnly.out.meta.width-1).sortField,null);
});

test('부분합과 축소 그룹 값은 선택 수준과 반대 축 전체 경로를 유지한다',()=>{
  const data=[['지역','상품','연도','채널','매출'],['서울','A',2025,'X',90],['서울','B',2025,'Y',1],['부산','A',2026,'X',10],['부산','B',2026,'Y',200]];
  for(const layout of ['compact','outline','tabular'])for(const collapsed of [{},{지역:['서울']}]){
    const d={rows:['지역','상품'],cols:['연도','채널'],values:[{field:'매출',agg:'sum'}],layout,subtotals:true,subtotalTop:false,collapsed};
    const res=resolvePivot(data,d), out=computePivot(res,res.def);
    const r=out.meta.pageRows+out.meta.headerRows+out.meta.rowItems.findIndex(x=>x.node?.key==='서울'&&(x.kind==='sub'||x.coll));
    const c=out.meta.labelCols+out.meta.colLeaves.findIndex(x=>x.node.key==='X'&&x.node.parent.key===2025);
    const target=pivotContextTarget(res.def,out,r,c);
    assert.equal(target.sortField,'지역');assert.deepEqual(target.sortAt,[['연도','2025'],['채널','X']]);
    const subcol=out.meta.labelCols+out.meta.colLeaves.findIndex(x=>x.kind==='sub'&&x.node.key===2025);
    assert.deepEqual(pivotContextTarget(res.def,out,r,subcol).sortAt,[['연도','2025']]);
    const grandRow=out.grid.length-1;
    assert.equal(pivotContextTarget(res.def,out,grandRow,subcol).sortField,'연도');
    const subheads=targetCells({def:res.def,out},/^colSubHead$/);
    assert.ok(subheads.length);assert.ok(subheads.every(x=>x.target.sortAxis==='cols'&&x.target.sortField==='연도'));
  }
});

test('보고서 필터와 빈 줄 및 총합계 레이블에서 다른 필드 정렬을 추측하지 않는다',()=>{
  const computed=result({...base,rows:['지역','상품'],pages:['수량'],blankRows:true,subtotals:true});
  for(const {target}of targetCells(computed,/^(pageLabel|pageValue|blank|grandLabel)$/)){
    assert.equal(target.sortField,null);assert.equal(target.sortAxis,null);assert.equal(target.sortAt,undefined);
  }
  const outside=pivotContextTarget(computed.def,computed.out,999,999);assert.equal(outside.sortField,null);
});

test('값 행 배치에서 선택한 지표 인덱스와 열 항목별 정렬 범위를 보존한다',()=>{
  for(const layout of ['compact','outline','tabular']){
    const computed=result({...base,cols:['상품'],valuesOnRows:true,layout});
    const values=targetCells(computed,/^data:1$/);assert.ok(values.length);
    for(const {target,c}of values){
      assert.equal(target.valueIndex,1);assert.equal(target.sortField,'지역');assert.equal(target.sortAxis,'rows');
      assert.deepEqual(target.sortAt,[['상품',computed.out.meta.colLeaves[c-computed.out.meta.labelCols].node.key]]);
    }
    const grand=targetCells(computed,/^grandData:1$/).find(x=>computed.out.meta.colLeaves[x.c-computed.out.meta.labelCols].kind==='item');
    assert.equal(grand.target.sortAxis,'cols');assert.equal(grand.target.sortField,'상품');assert.equal(grand.target.valueIndex,1);
  }
});

test('선택 열 값 정렬의 원본 경로는 XLSX 저장·재열기 뒤에도 같은 순서를 낸다',()=>{
  const cells={};source.forEach((row,r)=>row.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));
  const d={...base,cols:['상품'],values:[base.values[0]],source:'원본',range:{r1:0,c1:0,r2:3,c2:3},name:'정렬피벗',top:0,left:0};
  const computed=result(d), c=computed.out.meta.labelCols+computed.out.meta.colLeaves.findIndex(x=>x.node.key==='A');
  const target=pivotContextTarget(computed.def,computed.out,computed.out.meta.pageRows+computed.out.meta.headerRows,c);
  const sorted=sortFromTarget(d,target), wb=new Workbook({sheets:[{name:'원본',cells},{name:'피벗',cells:{},pivot:sorted}]});
  const restored=new Workbook(readXlsx(writeXlsx(wb)).data).sheets[1].pivot;
  assert.deepEqual(restored.sort.지역.at,[['상품','A']]);
  assert.deepEqual(result(restored).out.meta.rowItems.filter(x=>x.kind==='item').map(x=>x.node.key),['부산','서울']);
});
