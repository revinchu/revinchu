import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Column, Cube, EMPTY, EMPTY_TEXT, IMG_KEY, itemIdentity, itemProperty, keyOf, cubeFromRows, blockColumn, filterRows, groupAggregate, groupedColumn, itemStats, planRollup, aggregateQuery } from '../src/cube.js';
import { computePivot, resolvePivot, pivotLookup, pivotFilterKey } from '../src/pivot.js';

const rows = [['Label','Period','Amount'],['Alpha','X',1],['alpha','x',2],['Beta','X',4],['BETA','Y',8]];
const base = {rows:['Label'],cols:['Period'],values:[{field:'Amount',agg:'sum'}],layout:'tabular'};
const grid = (def=base,data=rows) => {const r=resolvePivot(data,def);return computePivot(r,r.def).grid.map(row=>row.map(c=>c.raw));};

test('항목 비교만 case-fold하고 타입·빈값·악센트·너비·그림 주소 및 원문은 보존',()=>{
 assert.equal(itemIdentity('Alpha'),itemIdentity('ALPHA'));
 for(const [a,b] of [[1,'1'],[true,'true'],[EMPTY,EMPTY_TEXT],['é','e'],['Ａ','A'],['ß','ss'],[IMG_KEY+'data:A',IMG_KEY+'data:a']])assert.notEqual(itemIdentity(a),itemIdentity(b));
 assert.equal(keyOf('Alpha'),'Alpha');assert.equal(keyOf(''),EMPTY_TEXT);assert.equal(keyOf(null),EMPTY);
 assert.equal(itemProperty({ALPHA:'caption'},'Alpha'),'caption');
});

test('일반 차원은 첫 표시문자를 보존하고 alias행을 같은 코드로 집계',()=>{
 const values=['Alpha','alpha','ALPHA',1,'1',true,'true',null,''];
 const col=new Column(values.length,i=>values[i]), d=col.dim();
 assert.deepEqual([...d.codes].slice(0,3),[0,0,0]);assert.equal(d.keys[0],'Alpha');assert.equal(d.keys.length,7);
 assert.deepEqual(values.map((_,i)=>col.get(i)),values);
});

test('열 블록의 사전 경로도 일반 차원과 같은 항목 코드 및 원문을 유지',()=>{
 const values=['Alpha','ALPHA','é','É','e'];
 const bc={dict:values,str:Int32Array.from(values.map((_,i)=>i)),num:new Float64Array(values.length).fill(NaN)};
 const col=blockColumn(bc,0,values.length);
 assert.deepEqual([...col.dim().codes],[0,0,1,1,2]);assert.deepEqual(col.dim().keys,['Alpha','é','e']);
 assert.deepEqual(values.map((_,i)=>col.get(i)),values);assert.deepEqual(bc.dict,values);
});

test('행·열 항목과 양방향 총합계가 alias의 모든 값을 포함',()=>{
 assert.deepEqual(grid(),[['합계 : Amount','Period','',''],['Label','X','Y','총합계'],['Alpha','3','','3'],['Beta','4','8','12'],['총합계','7','8','15']]);
 assert.deepEqual(rows[2],['alpha','x',2]);
});

test('항목 필터는 어떤 대소문자 alias를 선택해도 모든 원본 행을 포함',()=>{
 const cube=cubeFromRows(rows);
 for(const chosen of ['Alpha','alpha','ALPHA'])assert.deepEqual([...filterRows(cube,[[0,new Set([chosen])]])],[0,1]);
 const r=resolvePivot(rows,{...base,filters:{Label:['aLpHa']}});assert.equal(pivotLookup(rows,r.def,'Amount',[],r),3);
 assert.equal(pivotFilterKey({Label:['Alpha','ALPHA']}),pivotFilterKey({label:['alpha']}));
});

test('값 필터와 상위 항목 필터는 통합 집계 결과를 기준으로 적용',()=>{
 const value=resolvePivot(rows,{...base,fieldFilters:{Label:{type:'value',op:'gt',v1:2,v2:'',by:0}}});
 assert.deepEqual([...new Set(value.groups.map(g=>g.r[0]))],['Alpha','Beta']);
 const top=resolvePivot(rows,{...base,fieldFilters:{Label:{type:'top',n:1,top:true,by:0}}});
 assert.deepEqual([...new Set(top.groups.map(g=>g.r[0]))],['Beta']);
});

test('슬라이서 항목은 통합하고 다른 alias 필터의 데이터 있음 상태 유지',()=>{
 const cube=cubeFromRows(rows), stats=itemStats(cube,0,[[1,new Set(['y'])]]);
 assert.deepEqual(stats.keys,['Alpha','Beta']);assert.deepEqual([...stats.has],[0,1]);
 const selected=new Set(['ALPHA'].map(itemIdentity));
 assert.deepEqual(stats.texts.map(t=>selected.has(itemIdentity(t))),[true,false]);
});

test('선택 항목 그룹 이름과 그룹 필터도 alias에 일관되게 적용',()=>{
 const cube=cubeFromRows(rows),spec={by:'items',map:{ALPHA:'Group',beta:'group'}};
 assert.deepEqual(groupedColumn(cube,0,spec).dim().keys,['Group']);
 const r=resolvePivot(rows,{...base,cols:[],groups:{Label:spec},filters:{Label:['GROUP']}});
 assert.equal(pivotLookup(rows,r.def,'Amount',[],r),15);
});

test('저장된 수동순서·사용자 캡션과 기준 항목 계산은 alias를 찾음',()=>{
 const def={...base,cols:[],order:{Label:['BETA','ALPHA']},itemCaptions:{Label:{ALPHA:'Display Alpha'}}};
 const g=grid(def);assert.equal(g[1][0],'Beta');assert.equal(g[2][0],'Display Alpha');
 const diff=grid({...base,cols:[],values:[{field:'Amount',agg:'sum',showAs:'difference',baseField:'Label',baseItem:'ALPHA'}]});
 assert.equal(diff.find(r=>r[0]==='Beta')[1],'9');
 assert.equal(pivotLookup(rows,base,'Amount',[['Label','ALPHA']]),3);
});

test('축소 상태와 GETPIVOTDATA 항목 조회가 대소문자 alias에서 유지',()=>{
 const data=[['Outer','Inner','Amount'],['Alpha','one',1],['alpha','TWO',2],['Beta','one',4]];
 const def={rows:['Outer','Inner'],cols:[],values:[{field:'Amount',agg:'sum'}],layout:'tabular',collapsed:{Outer:['ALPHA']}};
 const g=grid(def,data);assert.equal(g.some(r=>r.includes('TWO')),false);
 assert.equal(pivotLookup(data,def,'Amount',[['Outer','ALPHA']]),3);
 assert.equal(pivotLookup(data,def,'Amount',[['Outer','alpha'],['Inner','two']]),null);
});

test('20만행 롤업의 집계·필터·슬라이서 빠른 경로도 alias를 일관되게 비교',()=>{
 const n=200000, labels=['Alpha','ALPHA','Beta','BETA'];
 const cube=new Cube(n,['Label','Period','Amount'],j=>new Column(n,i=>j===0?labels[i%4]:j===1?(i%2?'north':'NORTH'):1),i=>[labels[i%4],i%2?'north':'NORTH',1]);
 const measures=[{col:2,need:{}}];planRollup(cube,[0,1],measures);
 assert.ok(cube.rollup&&!cube.rollup.useless);
 const agg=aggregateQuery(cube,[[0,new Set(['alpha'])]],[1],measures);
 assert.equal(agg.G,1);assert.equal(agg.stats[0].sum[0],100000);
 assert.deepEqual([...itemStats(cube,0,[[1,new Set(['NoRtH'])]]).has],[1,1]);
 assert.deepEqual([...itemStats(cube,1,[[0,new Set(['aLpHa'])]]).has],[1]);
 assert.equal(cube.row(1)[0],'ALPHA');
});
