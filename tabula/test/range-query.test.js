import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { serialOf } from '../src/format.js';
import { readRangeQuerySource, rangeQueryHeaders, transformRangeQuery } from '../src/range-query.js';
const area = (r1, c1, r2 = r1, c2 = c1) => ({r1,c1,r2,c2});
const source = (matrix, extra = {}) => ({matrix,width:matrix[0].length,header:true,range:area(0,0,matrix.length-1,matrix[0].length-1),...extra});

test('범위는 원본 값·수식·캐시를 바꾸지 않고 계산값 사본을 만든다', () => {
  const w = new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'이름'},'0,1':{raw:'금액'},'1,0':{raw:'001',inputType:'text'},'1,1':{raw:'=1+1',cached:42},'2,0':{raw:'=1+1',inputType:'text'},'2,1':{raw:'FALSE'}}}]});
  const before = JSON.stringify(w.serialize()), version = w.version, undo = w.undoStack.length;
  const s = readRangeQuerySource(w,0,area(0,0,2,1));
  assert.deepEqual(s.matrix,[['이름','금액'],['001',2],['=1+1',false]]);
  s.matrix[1][0]='다른값'; assert.equal(w.getValue(0,1,0),'001');
  assert.equal(JSON.stringify(w.serialize()),before); assert.equal(w.version,version); assert.equal(w.undoStack.length,undo);
});
test('전체 행·열만 사용 범위로 제한하고 초과 범위는 읽기 전에 차단한다', () => {
  let calls=0; const w={sheets:[{name:'합성'}],usedRange:()=>({rows:4,cols:3}),rangeRead:(_s,r1,c1,r2,c2)=>{calls++;return Array.from({length:r2-r1+1},()=>Array(c2-c1+1).fill(null));}};
  const s=readRangeQuerySource(w,0,area(0,0,1048575,16383));assert.deepEqual(s.range,area(0,0,3,2));assert.equal(s.clipped,true);assert.equal(calls,1);
  assert.throws(()=>readRangeQuerySource(w,0,area(0,0,100000,10)),/범위를 줄여/);assert.equal(calls,1);
  assert.throws(()=>readRangeQuerySource(w,0,area(5,0,1048575,1)),/데이터가 없습니다/);assert.equal(calls,1);
  assert.throws(()=>readRangeQuerySource(w,0,area(-1,0,3,1)),/올바르지/);
});
test('표의 머리글 설정과 합계 제외·숨긴 행 포함을 보존한다',()=>{
  const w=new Workbook({sheets:[{name:'표 시트',cells:{'2,1':{raw:'제목'},'3,1':{raw:'5'},'4,1':{raw:'7'},'5,1':{raw:'12'}},hiddenRows:[3]}]});
  const table={name:'표1',r1:2,c1:1,r2:5,c2:1,totals:true};
  const s=readRangeQuerySource(w,0,area(3,1),{table});assert.deepEqual(s.matrix,[['제목'],[5],[7]]);assert.equal(s.header,true);assert.match(s.label,/표1/);
  const no=readRangeQuerySource(w,0,area(3,1,4,1),{header:false});assert.deepEqual(transformRangeQuery(no).matrix,[['열1'],[5],[7]]);
});
test('머리글의 빈 값·대소문자 중복·생성 이름 충돌을 정리한다',()=>{
  const s=source([['매출','매출','매출_2','','Revenue','revenue'],[1,2,3,4,5,6]]);
  assert.deepEqual(rangeQueryHeaders(s),['매출','매출_2','매출_2_2','열4','Revenue','revenue_2']);
  assert.deepEqual(transformRangeQuery(s,{header:false}).matrix[1],['매출','매출','매출_2',null,'Revenue','revenue']);
});
test('열 선택·유형·다중 조건·안정 정렬·빈 행·중복 제거가 함께 작동한다',()=>{
  const s=source([['품목','금액','버림'],['가','1,200','a'],['나','900','b'],['가','1,200','c'],['','',null],['다','1,300','d'],['라','1,200','e']]);
  const old=JSON.stringify(s),r=transformRangeQuery(s,{columns:[1,0],types:{1:'number'},filters:[{column:1,op:'ge',value:'1000'},{column:0,op:'notContains',value:'다'}],removeBlankRows:true,removeDuplicates:true,sorts:[{column:1,direction:'desc'}]});
  assert.deepEqual(r.matrix,[['금액','품목'],[1200,'가'],[1200,'라']]);assert.equal(r.stats.duplicateRows,1);assert.equal(r.stats.filteredRows,3);assert.equal(JSON.stringify(s),old);
});
test('빈 행 제거는 0·FALSE·공백 문자를 제거하지 않고 중복은 유형을 구별한다',()=>{
  const s=source([['열'],[null],[''],[0],[false],[' '],[1],['1'],[1],[{code:'#N/A'}]]);
  const r=transformRangeQuery(s,{removeBlankRows:true,removeDuplicates:true});assert.deepEqual(r.matrix,[['열'],[0],[false],[' '],[1],['1'],[{code:'#N/A'}]]);assert.equal(r.stats.blankRows,2);assert.equal(r.stats.duplicateRows,1);
});
test('숫자는 쉼표·백분율·지수를 변환하고 부정확한 큰 정수·비정상 그룹은 오류다',()=>{
  const s=source([['값'],['1,234.50'],['12.5%'],['-2e3'],['9007199254740993'],['12,34'],['abc'],[''],[true]]);
  const r=transformRangeQuery(s,{types:['number']});assert.deepEqual(r.matrix.slice(1,4),[[1234.5],[0.125],[-2000]]);assert.equal(r.stats.errorCount,4);assert.equal(r.errors[0].row,5);assert.deepEqual(r.columnFormats,[{numFmt:'general'}]);
});
test('정수·논리값 변환은 잘못된 값을 조용히 반올림하지 않는다',()=>{
  const r=transformRangeQuery(source([['정수','참거짓'],['2','TRUE'],['2.5','거짓'],['','1'],['-3','틀림']]),{types:['integer','boolean']});
  assert.deepEqual(r.matrix[1],[2,true]);assert.deepEqual(r.matrix[2],[null,false]);assert.equal(r.stats.errorCount,2);assert.deepEqual(r.columnFormats[0],{numFmt:'number',decimals:0});
});
test('날짜는 윤년과 1904 통합 문서를 따르고 날짜 필터는 ISO 날짜를 비교한다',()=>{
  for(const date1904 of [false,true]){const s=source([['날짜'],['2024-02-29'],['2024-03-01'],['2023-02-29']],{date1904});const r=transformRangeQuery(s,{types:['date'],filters:[{column:0,op:'ge',value:'2024-03-01'}]});assert.deepEqual(r.matrix,[['날짜'],[serialOf(2024,3,1,date1904)]]);assert.equal(r.stats.errorCount,1);assert.deepEqual(r.columnFormats,[{numFmt:'date'}]);}
});
test('텍스트 변환은 수식처럼 보이는 문자열·오류 코드를 실행하지 않는다',()=>{
  const r=transformRangeQuery(source([['값'],['=HYPERLINK("x")'],[0],[false],[{code:'#DIV/0!'}]]),{types:['text']});assert.deepEqual(r.matrix,[['값'],['=HYPERLINK("x")'],['0'],['FALSE'],['#DIV/0!']]);assert.deepEqual(r.columnFormats,[{numFmt:'text'}]);
});
test('제외한 열의 오류는 로드를 막지 않으며 필터·정렬 열은 선택 밖에서도 동작한다',()=>{
  const s=source([['이름','번호','미사용'],['가','2','bad'],['나','1','bad']]);const r=transformRangeQuery(s,{columns:[0],types:['keep','number','integer'],sorts:[{column:1,direction:'asc'}]});assert.equal(r.stats.errorCount,0);assert.deepEqual(r.matrix,[['이름'],['나'],['가']]);
});
test('빈 결과·잘못된 열·숫자 필터는 명확한 결과 또는 오류를 반환한다',()=>{
  const s=source([['제목'],[1]]);assert.deepEqual(transformRangeQuery(s,{filters:[{column:0,op:'gt',value:'2'}]}).matrix,[['제목']]);assert.throws(()=>transformRangeQuery(s,{columns:[]}),/하나 이상/);assert.throws(()=>transformRangeQuery(s,{columns:[0,0]}),/하나 이상/);assert.throws(()=>transformRangeQuery(s,{filters:[{column:0,op:'gt',value:'oops'}]}),/올바른 숫자/);assert.throws(()=>transformRangeQuery(s,{sorts:[{column:7,direction:'asc'}]}),/정렬/);
});
test('10만 행 변환은 스택 확장 없이 입력을 보존한다',()=>{
  const matrix=[['번호','분류']];for(let i=0;i<99999;i++)matrix.push([String(i),i%3]);const s=source(matrix);const r=transformRangeQuery(s,{types:['integer','keep'],filters:[{column:1,op:'eq',value:'2'}],sorts:[{column:0,direction:'desc'}]});assert.equal(r.stats.outputRows,33333);assert.deepEqual(r.matrix[1],[99998,2]);assert.equal(s.matrix[1][0],'0');
});
