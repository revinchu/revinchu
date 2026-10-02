import {test} from 'node:test';
import assert from 'node:assert/strict';
import {computePivot,normalizeDef,pivotPageLayout,resolvePivot} from '../src/pivot.js';
import {Workbook} from '../src/workbook.js';

const source=[['Region','Product','Year','Channel','Owner','Amount'],['East','A','2025','Web','Kim',10],['West','B','2026','App','Lee',20],['East','B','2026','Web','Lee',30]];
const fields=source[0].slice(0,5);
const base={rows:[],cols:[],values:[],pages:['Region'],filters:{},style:'PivotStyleLight16'};
const calculate=(patch={})=>{const def={...base,...patch},resolved=resolvePivot(source,def);return computePivot(resolved,resolved.def);};
const labels=(result)=>result.meta.pageFields.map(({field,r,c})=>({field,r,c,label:result.grid[r][c].raw,value:result.grid[r][c+1].raw}));

test('보고서 필터만/본문 있는 피벗은 같은 표시 값과 기본 제공 스타일을 사용',()=>{
 for(const style of ['PivotStyleLight16','PivotStyleMedium9','PivotStyleDark1'])for(const selected of [undefined,[],['East'],['East','West'],['East','Missing'],['Missing']]){
  const patch={style,filters:selected?{Region:selected}:{}};
  const only=calculate(patch),withValues=calculate({...patch,values:[{field:'Amount',agg:'sum'}]});
  assert.deepEqual(only.grid[0],withValues.grid[0],style+' / '+JSON.stringify(selected));
  assert.ok(Object.keys(only.grid[0][0].style).length>0,'필터만 있어도 표준 스타일을 적용');
  assert.equal(only.meta.pageRows,1);assert.equal(only.meta.bodyRows,0);assert.equal(only.meta.empty,true);
  assert.equal(withValues.meta.pageRows,2);assert.deepEqual(withValues.grid[1],[],'필터 뒤 한 줄 여백');
 }
 assert.equal(calculate({filters:{Region:['East','West']}}).grid[0][1].raw,'(모두)');
 assert.equal(calculate({filters:{Region:[]}}).grid[0][1].raw,'(다중 항목)','빈 선택을 전체로 바꾸지 않음');
 assert.equal(calculate({filters:{Region:['East','Missing']}}).grid[0][1].raw,'East');
});

test('필터의 실제 필드 이름과 사용자 캡션은 분리, 대소문자별 필터 키도 동일',()=>{
 const g=calculate({pages:['region'],filters:{region:['East','Missing']},fieldCaptions:{Region:'지역'},itemCaptions:{Region:{East:'동부'}}});
 assert.deepEqual(g.meta.pageFields,[{field:'Region',r:0,c:0}]);
 assert.deepEqual(g.grid[0].map(c=>[c.raw,c.field,c.role]),[['지역','Region','pageLabel'],['동부','Region','pageValue']]);
});

test('사용자 지정 보고서 필터 스타일도 본문 유무와 관계없이 유지',()=>{
 const patch={style:'Custom',styleDef:{page:{fill:'#112233',color:'#ffffff',bold:true}},pages:['Region','Product']};
 assert.deepEqual(calculate(patch).grid,calculate({...patch,values:[{field:'Amount',agg:'sum'}]}).grid.slice(0,2));
 assert.equal(calculate(patch).grid[1][1].style.fill,'#112233');
});

test('행 우선/열 우선 및 필드 수 배치를 희소 셀 좌표로 계산',()=>{
 const expected={
  'down:0':[[0,0],[1,0],[2,0],[3,0],[4,0]],
  'over:0':[[0,0],[0,3],[0,6],[0,9],[0,12]],
  'down:2':[[0,0],[1,0],[0,3],[1,3],[0,6]],
  'over:2':[[0,0],[0,3],[1,0],[1,3],[2,0]],
 };
 for(const [key,positions] of Object.entries(expected)){
  const [pageOrder,wrap]=key.split(':'),patch={pages:fields,pageOrder,pageWrap:+wrap};
  const plan=pivotPageLayout(patch),g=calculate(patch);
  assert.deepEqual(plan.fields.map(({r,c})=>[r,c]),positions,key);
  assert.deepEqual(g.meta.pageFields,plan.fields);assert.equal(g.meta.pageRows,plan.height);assert.equal(g.meta.pageWidth,plan.width);
  assert.deepEqual(labels(g).map(x=>x.label),fields);
  for(const row of g.grid)for(let c=2;c<row.length;c+=3)assert.equal(c in row,false,'필터 사이 사용자 칸은 희소로 남김');
  for(const {field,r,c} of plan.fields){assert.equal(g.grid[r][c].field,field);assert.equal(g.grid[r][c+1].field,field);}
 }
});

test('보고서 필터가 본문보다 넓어도 본문 폭·총합계·스타일을 늘리지 않음',()=>{
 const patch={pages:fields,pageOrder:'over',pageWrap:2,values:[{field:'Amount',agg:'sum'}]};
 const g=calculate(patch),normal=calculate({values:patch.values});
 assert.equal(g.meta.width,1);assert.equal(g.meta.pageWidth,5);assert.equal(g.meta.pageRows,4);
 assert.deepEqual(g.grid.slice(4),normal.grid.slice(2));
 assert.equal(g.grid.at(-1)[0].raw,'60');
 assert.equal(g.grid[2].length,2,'마지막 줄 미배치 필터 자리는 생성하지 않음');
 assert.deepEqual(g.grid[3],[]);
});

test('보고서 필터 스타일은 표 배치·머리글 표시 옵션에 따라 소실되지 않음',()=>{
 const before=calculate().grid[0];
 for(const layout of ['compact','outline','tabular'])for(const showHeaders of [true,false]){
  const g=calculate({rows:['Product'],values:[{field:'Amount',agg:'sum'}],layout,showHeaders});
  assert.deepEqual(g.grid[0],before);
 }
});

test('배치 옵션 정규화·원본 불변·직렬화 및 Undo/Redo 보존',()=>{
 const def={...base,pages:fields,pageOrder:'over',pageWrap:2},before=structuredClone(def),data=structuredClone(source);
 const n=normalizeDef(def,source[0]);assert.equal(n.pageOrder,'over');assert.equal(n.pageWrap,2);assert.deepEqual(normalizeDef(n,source[0]),n);
 calculate(def);assert.deepEqual(def,before);assert.deepEqual(source,data);
 for(const pageWrap of [undefined,-1,NaN,Infinity,'bad'])assert.equal(normalizeDef({...def,pageWrap},source[0]).pageWrap,0);
 assert.equal(normalizeDef({...def,pageOrder:'bad'},source[0]).pageOrder,'down');
 const wb=new Workbook();wb.sheets[0].pivot=structuredClone(def);
 wb.transact(()=>wb.setSheetProp(0,'pivot',{...def,pageOrder:'down',pageWrap:3}));
 wb.undo();assert.deepEqual(wb.sheets[0].pivot,def);wb.redo();assert.equal(wb.sheets[0].pivot.pageWrap,3);
 const restored=new Workbook(wb.serialize());assert.equal(restored.sheets[0].pivot.pageOrder,'down');assert.equal(restored.sheets[0].pivot.pageWrap,3);
});

test('새 시트의 가운데 정렬을 상속하지 않고 표준 일반 정렬, 명시 스타일은 보존',()=>{
 const g=calculate({rows:['Product'],values:[{field:'Amount',agg:'sum'}]});
 for(const row of g.grid)for(const cell of row)if(cell)assert.equal(cell.style.align,'general');
 const wb=new Workbook();wb.setSheetProp(0,'allStyle',{align:'center'});
 for(let r=0;r<g.grid.length;r++)g.grid[r].forEach((cell,c)=>{if(cell)wb.setCellData(0,r,c,{raw:cell.raw,style:cell.style});});
 assert.equal(wb.styleAt(0,0,0).align,'general');assert.equal(wb.styleAt(0,g.grid.length-1,1).align,'general');
 const explicit=calculate({rows:['Product'],values:[{field:'Amount',agg:'sum'}],style:'Custom',styleDef:{page:{align:'right'},header:{align:'center'},body:{align:'left'},grand:{align:'right'}}});
 assert.equal(explicit.grid[0][0].style.align,'right');assert.equal(explicit.grid[2][0].style.align,'center');assert.equal(explicit.grid[3][1].style.align,'left');assert.equal(explicit.grid.at(-1)[1].style.align,'right');
});

test('옛 fieldCaptions=false는 행·열 머리글만 끄고 보고서 필터는 유지',()=>{
 const patch={rows:['Product'],values:[{field:'Amount',agg:'sum'}],fieldCaptions:false};
 assert.equal(normalizeDef({...base,...patch},source[0]).showHeaders,false);
 const hidden=calculate(patch),visible=calculate({...patch,fieldCaptions:{Region:'지역'}});
 assert.equal(hidden.grid[2][0].raw,'');assert.equal(visible.grid[2][0].raw,'행 레이블');
 assert.equal(hidden.grid[0][0].raw,'Region');assert.equal(hidden.grid[0][1].raw,'(모두)');assert.equal(visible.grid[0][0].raw,'지역');
});

test('실제 Excel DisplayFormat 기준 보고서 스타일의 채움·하단선·글자 강조 보존',()=>{
 // Excel Range.Interior는 기본 셀 서식을, DisplayFormat은 피벗 스타일까지 포함한 표시를 반환합니다.
 // 합성 Light16/Medium9/Dark1에서 필터만/본문 있음의 DisplayFormat은 동일했습니다.
 // 채움 hex는 기존 WIXEL tint 기준입니다(Excel Light16 #d9e1f2와 1~2 RGB 차이, 픽셀 동등 검사는 아님).
 const expected={
  PivotStyleLight16:{fill:'#dae2f3',bottom:true,boldLabel:false,color:undefined},
  PivotStyleMedium9:{fill:'#dae2f3',bottom:false,boldLabel:false,color:'#000000'},
  PivotStyleDark1:{fill:'#808080',bottom:false,boldLabel:true,color:'#ffffff'},
 };
 for(const [style,want] of Object.entries(expected))for(const values of [[],[{field:'Amount',agg:'sum'}]]){
  const g=calculate({style,values});
  for(const {r,c} of g.meta.pageFields){
   const label=g.grid[r][c].style,value=g.grid[r][c+1].style;
   for(const cell of [label,value]){
    assert.equal(cell.fill,want.fill);assert.equal(!!cell.bb,want.bottom);
    assert.equal(cell.color,want.color);assert.equal(cell.align,'general');
   }
   assert.equal(!!label.bold,want.boldLabel);assert.equal(!!value.bold,false);
  }
 }
});
