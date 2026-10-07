import test from 'node:test';
import assert from 'node:assert/strict';
import { importedPivotToggleButtons } from '../src/pivot-import-toggles.js';
import { parseXml, child, kids } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
import { computePivot, resolvePivot } from '../src/pivot.js';
import { unzip, textOf } from '../src/zip.js';
import { parseRangeName } from '../src/formula.js';
const refToRange = ref => parseRangeName(ref);

function fixture({ fields = [['Group', ['A', 'B']], ['Detail', ['x', 'y']]], rows = [0, 1], cols = [], rowItems = '', colItems = '',
  firstDataRow = 1, firstDataCol = 2, firstHeaderRow = 0, layout = 'tabular', pitems = null, values = [{field:'Amount'}], cells = {},
  body = {r1:10,c1:4,r2:40,c2:30}, patch = {}, date1904 = false } = {}) {
  const cache={date1904,fields:fields.map(([name,items])=>({name,items}))};
  const pf=fields.map(([,items],i)=>'<pivotField><items>'+ (pitems?.[i] ?? items.map((_,x)=>'<item x="'+x+'"/>').join(''))+'</items></pivotField>').join('');
  const fieldXml=a=>a.map(x=>'<field x="'+x+'"/>').join('');
  const root=parseXml('<pivotTableDefinition><location firstDataRow="'+firstDataRow+'" firstHeaderRow="'+firstHeaderRow+'" firstDataCol="'+firstDataCol+'"/>'
    +'<pivotFields>'+pf+'</pivotFields><rowFields>'+fieldXml(rows)+'</rowFields><colFields>'+fieldXml(cols)+'</colFields>'
    +'<rowItems>'+rowItems+'</rowItems><colItems>'+colItems+'</colItems></pivotTableDefinition>');
  const def={rows:rows.filter(f=>f>=0).map(f=>fields[f][0]),cols:cols.filter(f=>f>=0).map(f=>fields[f][0]),values,layout,...patch};
  const calls=[];
  const cellAt=(r,c)=>{calls.push([r,c]);return cells[r+','+c];};
  return {root,cache,def,body,cellAt,calls,cells};
}
const button=(r,c,field,item,collapsed=false)=>({r,c,kind:'toggle',field,item,collapsed});
const raw=s=>({raw:s});
const triple=[['Group',['A','B']],['Middle',['m','n']],['Detail',['x','y']]];

test('tabular 상속 prefix는 첫 그룹만 복원하고 반복 레이블/총합계를 중복하지 않는다',()=>{
  const f=fixture({rowItems:'<i><x/><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i><i><x v="1"/><x/></i><i t="grand"><x/></i>',
    cells:{'11,4':raw('A'),'12,4':raw('A'),'13,4':raw('A'),'14,4':raw('B'),'15,4':raw('B')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(14,4,'Group','B')]);
  assert.deepEqual(f.calls,[[11,4],[14,4]],'반복 레이블 셀이나 전체 본문을 조회하지 않는다');
});

test('compact는 실제 표시 수준만 복원하고 subtotal/blank/grand를 포함한 물리 행 순번을 유지한다',()=>{
  const f=fixture({fields:triple,rows:[0,1,2],layout:'compact',firstDataCol:1,
    rowItems:'<i><x/></i><i r="1"><x/></i><i r="2"><x/></i><i t="default" r="1"><x/></i><i t="blank" r="1"><x/></i><i><x v="1"/></i><i r="1"><x v="1"/></i><i t="grand"><x/></i><i r="1"><x/></i>',
    cells:{'11,4':raw('A'),'12,4':raw('m'),'13,4':raw('x'),'14,4':raw('m'),'16,4':raw('B'),'17,4':raw('n'),'19,4':raw('m')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(12,4,'Middle','m'),button(16,4,'Group','B'),button(17,4,'Middle','n')]);
  assert.deepEqual(f.calls,[[11,4],[12,4],[16,4],[17,4]]);
});

test('outline 반복 상위 레이블에는 버튼을 늘리지 않고 고유 중간 그룹은 부모별로 복원한다',()=>{
  const f=fixture({fields:triple,rows:[0,1,2],layout:'outline',firstDataCol:3,
    rowItems:'<i><x/></i><i r="1"><x/></i><i r="2"><x/></i><i><x v="1"/></i><i r="1"><x/></i>',
    cells:{'11,4':raw('A'),'12,4':raw('A'),'12,5':raw('m'),'13,4':raw('A'),'13,5':raw('m'),'14,4':raw('B'),'15,5':raw('m')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(12,5,'Middle','m'),button(14,4,'Group','B'),button(15,5,'Middle','m')]);
});

test('짧은 새 경로 뒤의 옛 suffix를 제거하여 잘못된 r 상속을 거절한다',()=>{
  const f=fixture({fields:triple,rows:[0,1,2],firstDataCol:3,rowItems:'<i><x/><x/><x/></i><i><x v="1"/></i><i r="2"><x v="1"/></i><i r="1"><x v="1"/><x/></i>',
    cells:{'11,4':raw('A'),'11,5':raw('m'),'12,4':raw('B'),'13,5':raw('m'),'14,5':raw('n')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(11,5,'Middle','m'),button(12,4,'Group','B')]);
  assert.equal(f.calls.some(([r])=>r===13||r===14),false,'불완전한 상속 경로는 실제 셀까지 조회하지 않는다');
});

test('값 Σ는 cache item을 조회하거나 토글하지 않고 값 행도 그룹 레이블로 오인하지 않는다',()=>{
  const f=fixture({rows:[0,1,-2],firstDataCol:3,values:[{field:'A'},{field:'B'}],
    rowItems:'<i><x/><x/><x/></i><i r="2" i="1"><x v="1"/></i><i><x v="1"/><x/><x/></i>',
    cells:{'11,4':raw('A'),'12,4':raw('A'),'13,4':raw('B')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(13,4,'Group','B')]);
  assert.deepEqual(f.calls,[[11,4],[13,4]]);
  const compact=fixture({fields:triple,rows:[0,1,2,-2],layout:'compact',firstDataCol:1,values:[{field:'A'},{field:'B'}],
    rowItems:'<i><x/></i><i r="1"><x/></i><i r="2"><x/></i><i r="3" i="1"><x v="1"/></i>',cells:{'11,4':raw('A'),'12,4':raw('m'),'14,4':raw('m')}});
  assert.deepEqual(importedPivotToggleButtons(compact),[button(11,4,'Group','A'),button(12,4,'Middle','m')]);
});

test('열 축은 firstHeaderRow/firstDataCol과 subtotal/grand의 실제 열 순번을 보존한다',()=>{
  const f=fixture({rows:[],cols:[0,1],firstHeaderRow:1,firstDataRow:3,firstDataCol:1,
    colItems:'<i><x/><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i><i><x v="1"/><x/></i><i t="grand"><x/></i>',
    cells:{'11,5':raw('A'),'11,6':raw('A'),'11,7':raw('A'),'11,8':raw('B'),'11,9':raw('B')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,5,'Group','A'),button(11,8,'Group','B')]);
  assert.deepEqual(f.calls,[[11,5],[11,8]]);
});

test('Σ가 열 축의 앞/가운데 있어도 부모 수준과 prefix identity를 분리한다',()=>{
  const f=fixture({rows:[],cols:[-2,0,1],values:[{field:'A'},{field:'B'}],firstHeaderRow:1,firstDataRow:4,firstDataCol:1,
    colItems:'<i><x/><x/><x/></i><i r="2"><x v="1"/></i><i i="1"><x v="1"/><x/><x/></i>',
    cells:{'12,5':raw('A'),'12,6':raw('A'),'12,7':raw('A')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(12,5,'Group','A'),button(12,7,'Group','A')]);
  const middle=fixture({rows:[],cols:[0,-2,1],values:[{field:'A'},{field:'B'}],firstHeaderRow:1,firstDataRow:4,firstDataCol:1,
    colItems:'<i><x/><x/><x/></i><i r="1" i="1"><x v="1"/><x/></i>',cells:{'11,5':raw('A'),'11,6':raw('A')}});
  assert.deepEqual(importedPivotToggleButtons(middle),[button(11,5,'Group','A')]);
});

test('pivotField.items 순서로 cache index를 해석하고 sd/정규화 collapsed를 보존한다',()=>{
  const f=fixture({pitems:['<item x="1" sd="0"/><item t="default"/><item x="0"/>'],
    rowItems:'<i><x/><x/></i><i><x v="2"/><x/></i><i><x v="1"/><x/></i>',
    patch:{collapsed:{Group:['a']}},cells:{'11,4':raw('B'),'12,4':raw('A'),'13,4':raw('A')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','B',true),button(12,4,'Group','A',true)]);
});

test('raw caption 검사는 사용자 캡션·인용 문자·숫자 레이블·날짜를 허용하고 잘못된 원문은 제외한다',()=>{
  const cases=[
    ['A',raw('changed'),{},false], ['A',raw('내 캡션'),{itemCaptions:{Group:{A:'내 캡션'}}},true],
    ["'A",raw("''A"),{},true], ['=A',raw("'=A"),{},true], ['=A',raw('=A'),{},false],
    [1,raw('1'),{},true], ['1',raw('1'),{},true], [1,raw('2'),{},false], [true,raw('TRUE'),{},true],
    [true,raw("'TRUE"),{},true], [44136,{raw:'2020-11-01',style:{numFmt:'date'}},{},true],
    [0,{raw:'1904-01-01',style:{numFmt:'date'}},{date1904:true},true],
    [null,raw('(비어 있음)'),{},true], ['',raw("'"),{},true],
    ['A',{raw:'A',formula:true},{},false], ['A',{raw:'A',fx:true},{},false],
  ];
  for(const [value,cell,patch,expected] of cases){
    const f=fixture({fields:[['Group',[value]],['Detail',['x']]],rowItems:'<i><x/><x/></i>',cells:{'11,4':cell},patch,date1904:patch.date1904??false});
    const out=importedPivotToggleButtons(f);assert.equal(out.length,expected?1:0,JSON.stringify([value,cell,patch]));
    if(expected)assert.equal(out[0].field,'Group');
  }
});

test('알 수 없는 타입/범위/순번/축과 비정상 header 위치는 추측하지 않는다',()=>{
  for(const rowItems of['<i r="-1"><x/></i>','<i r="1.2"><x/></i>','<i r="1"><x/></i>','<i><x v="99"/><x/></i>',
    '<i><x v="-1"/><x/></i>','<i><x v="1.2"/><x/></i>','<i><x/><x/><x/></i>','<i t="unknown"><x/><x/></i>']){
    const f=fixture({rowItems,cells:{'11,4':raw('A')}});assert.deepEqual(importedPivotToggleButtons(f),[],rowItems);
  }
  for(const change of[f=>f.def.rows.reverse(),f=>child(f.root,'rowFields').children[0].attrs.x='99',f=>child(f.root,'location').attrs.firstDataRow='-1',
    f=>f.body.r2=10,f=>f.def.showExpand=false,f=>f.cellAt=null]){
    const f=fixture({rowItems:'<i><x/><x/></i>',cells:{'11,4':raw('A')}});change(f);assert.deepEqual(importedPivotToggleButtons(f),[]);
  }
  const f=fixture({rows:[],cols:[0,1],firstHeaderRow:1,firstDataRow:1,colItems:'<i><x/><x/></i>',cells:{'11,6':raw('A')}});
  assert.deepEqual(importedPivotToggleButtons(f),[]);
  assert.deepEqual(importedPivotToggleButtons(),[]);
});

test('원문 범위 밖을 읽거나 본문/서식/캐시/정의를 변경하지 않는다',()=>{
  const f=fixture({rowItems:'<i><x/><x/></i><i><x v="1"/><x/></i>',body:{r1:10,c1:4,r2:11,c2:5},cells:{'11,4':{raw:'A',style:{bold:true}},'12,4':raw('B')}});
  const before=JSON.stringify({root:f.root,cache:f.cache,def:f.def,body:f.body,cells:f.cells});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A')]);
  assert.deepEqual(f.calls,[[11,4]]);
  assert.equal(JSON.stringify({root:f.root,cache:f.cache,def:f.def,body:f.body,cells:f.cells}),before);
});

test('숫자/문자/빈값은 같은 표시여도 typed cache 순번을 구분하고 대량 반복은 sparse 버튼만 만든다',()=>{
  const f=fixture({fields:[['Group',[1,'1',null,'']],['Detail',['x']]],
    rowItems:'<i><x/><x/></i><i><x v="1"/><x/></i><i><x v="2"/><x/></i><i><x v="3"/><x/></i>',
    cells:{'11,4':raw('1'),'12,4':raw('1'),'13,4':raw('(비어 있음)'),'14,4':raw("'")}});
  assert.equal(importedPivotToggleButtons(f).length,4,'항목 순번의 타입을 표시문구로 합치지 않는다');
  const many=fixture({rowItems:'<i><x/><x/></i>'+'<i r="1"><x/></i>'.repeat(20000),body:{r1:10,c1:4,r2:20020,c2:5},cells:{'11,4':raw('A')}});
  assert.deepEqual(importedPivotToggleButtons(many),[button(11,4,'Group','A')]);
  assert.equal(many.calls.length,1,'2만 반복 행의 셀/역할/토글 배열을 만들지 않는다');
});


test('축소한 compact 그룹 아래의 두 번째 값 행은 r이 중간 필드 위치여도 토글 대상이 아니다',()=>{
  const f=fixture({fields:triple,rows:[0,1,2,-2],layout:'compact',firstDataCol:1,values:[{field:'A'},{field:'B',name:'n'}],
    rowItems:'<i><x/></i><i r="1" i="1"><x v="1"/></i><i r="1"><x/></i>',
    cells:{'11,4':raw('A'),'12,4':raw('n'),'13,4':raw('m')}});
  assert.deepEqual(importedPivotToggleButtons(f),[button(11,4,'Group','A'),button(13,4,'Middle','m')]);
  assert.deepEqual(f.calls,[[11,4],[13,4]]);
});


// 현재 계산기가 만든 작은 피벗을 실제 XLSX writer로 저장합니다. XML 항목
// 구조와 raw 셀을 따로 읽어 재구성한 sparse 버튼이 계산기의 toggle과 같아야 합니다.
for(const layout of ['tabular','compact','outline'])for(const onRows of [false,true])test('작은 XLSX 출력의 행/열 토글 oracle: '+layout+', valuesOnRows='+onRows,()=>{
  const header=['Group','Middle','Detail','Region','Device','Amount','Other'];
  const rows=[header,['A','m','x','East','mobile',1,2],['A','m','y','East','desktop',3,4],['A','n','x','West','mobile',2,5],['B','m','x','East','mobile',7,8],['B','n','y','West','desktop',4,9]];
  const def={name:'Tiny toggles',source:'Source',range:{r1:0,c1:0,r2:rows.length-1,c2:header.length-1},rows:['Group','Middle','Detail'],cols:['Region','Device'],
    values:Array.from({length:13},(_,i)=>({field:i%2?'Other':'Amount',agg:'sum',name:'Value '+i})),layout,valuesOnRows:onRows,
    collapsed:{Group:['A'],Region:['East']},subtotals:true,blankRows:true,grandRows:true,grandCols:true,showExpand:true,top:4,left:3};
  const computed=computePivot(rows,resolvePivot(rows,def).def),source={};
  rows.forEach((row,r)=>row.forEach((v,c)=>source[r+','+c]={raw:String(v)}));
  const w=new Workbook({sheets:[{name:'Source',cells:source},{name:'Report',cells:{},pivot:def}]}),expected=[];
  computed.grid.forEach((row,r)=>row.forEach((cell,c)=>{w.setCellData(1,def.top+r,def.left+c,{raw:String(cell.raw),style:cell.style});if(cell.toggle)expected.push(button(def.top+r,def.left+c,cell.toggle.field,cell.toggle.item,cell.toggle.collapsed));}));
  const files=unzip(writeXlsx(w)),root=parseXml(textOf(files['xl/pivotTables/pivotTable1.xml'])),cacheRoot=parseXml(textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']));
  const cache={fields:kids(child(cacheRoot,'cacheFields'),'cacheField').map(f=>({name:f.attrs.name,items:(child(f,'sharedItems')?.children??[]).map(item=>item.name==='n'?Number(item.attrs.v):item.name==='b'?item.attrs.v==='1':item.name==='m'?null:item.attrs.v??'')}))};
  const actual=importedPivotToggleButtons({root,cache,def,body:refToRange(child(root,'location').attrs.ref),cellAt:(r,c)=>w.sheets[1].cells.getRC(r,c)});
  const ordered=list=>list.sort((a,b)=>a.r-b.r||a.c-b.c||a.field.localeCompare(b.field));
  assert.deepEqual(ordered(actual),ordered(expected));
});
