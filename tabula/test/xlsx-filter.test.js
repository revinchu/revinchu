import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { readAutoFilter, autoFilterXml } from '../src/xlsx-filter.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const range={r1:2,c1:2,r2:7,c2:5};
function fixture(table=true){
  const wb=new Workbook();
  for(let c=2;c<=5;c++)wb.setInput(0,2,c,`열${c}`);
  for(let r=3;r<=7;r++)for(let c=2;c<=5;c++)wb.setInput(0,r,c,c===2?(r%2?'A':'B'):String(r*10+c));
  const f={criteria:{2:['A'],3:{type:'custom',op1:'ge',v1:'30',join:'and',op2:'le',v2:'70'}},hidden:{4:true,6:true},hiddenButtons:{2:true,4:true},sort:{col:3,asc:false}};
  if(table)wb.sheets[0].tables=[{id:'t1',name:'자료표',...range,header:true,totals:false,filter:f}];
  else wb.sheets[0].filter={...range,...f};
  return wb;
}
const after=wb=>readXlsx(writeXlsx(wb)).data;
const model=(data,table=true)=>table?data.sheets[0].tables[0].filter:data.sheets[0].filter;
function replaceFilter(wb,body,attrs='',sort=''){
  const files=unzip(writeXlsx(wb));const path='xl/tables/table1.xml';
  files[path]=textOf(files[path]).replace(/<autoFilter[\s\S]*?<\/autoFilter>/,`<autoFilter ref="C3:F8" ${attrs}>${body}</autoFilter>`).replace(/<sortState[\s\S]*?<\/sortState>/,sort);
  return readXlsx(zip(files)).data;
}

test('표/범위의 단추 표시만 바뀌어도 조건·숨김·정렬·원본셀은 왕복에서 유지된다',()=>{
  for(const table of [true,false]){
    const wb=fixture(table),before=wb.serialize(),out=after(wb),f=model(out,table);
    assert.deepEqual(f.criteria,{2:['A'],3:{type:'custom',op1:'ge',v1:'30',join:'and',op2:'le',v2:'70'}});
    assert.deepEqual(f.hidden,{4:true,6:true});assert.deepEqual(f.hiddenButtons,{2:true,4:true});assert.deepEqual(f.sort,{col:3,asc:false});
    assert.deepEqual(wb.serialize(),before);
    const book=new Workbook(out);assert.equal(book.getValue(0,3,3),33);
    const fm=table?book.sheets[0].tables[0].filter:book.sheets[0].filter;
    delete fm.hiddenButtons;const shown=model(after(book),table);assert.equal(shown.hiddenButtons,undefined);assert.deepEqual(shown.criteria,f.criteria);assert.deepEqual(shown.hidden,f.hidden);
  }
});

test('조건 없는 hiddenButton/showButton도 읽고 상대 colId를 절대열로 옮긴다',()=>{
  const f=readAutoFilter(parseXml('<autoFilter><filterColumn colId="0" hiddenButton="true"/><filterColumn colId="1" showButton="false"/><filterColumn colId="2" hiddenButton="0" showButton="1"/><filterColumn colId="-1" hiddenButton="1"/><filterColumn colId="999" hiddenButton="1"/></autoFilter>'),range);
  assert.deepEqual(f,{criteria:{},hidden:{},hiddenButtons:{2:true,3:true}});
  const xml=autoFilterXml(f,range);assert.match(xml,/colId="0" hiddenButton="1"/);assert.doesNotMatch(xml,/colId="2"/);
});

test('필터를 해제해도 표의 형제 sortState는 남고 요약행을 정렬범위에 넣지 않는다',()=>{
  const wb=fixture(),t=wb.sheets[0].tables[0];t.sort=t.filter.sort;t.filter=null;t.totals=true;t.r2++;
  const files=unzip(writeXlsx(wb)),xml=textOf(files['xl/tables/table1.xml']);
  assert.doesNotMatch(xml,/<autoFilter/);assert.match(xml,/<sortState ref="C4:F8">/);assert.match(xml,/ref="D4:D8"/);
  const t2=after(wb).sheets[0].tables[0];assert.equal(t2.filter,null);assert.deepEqual(t2.sort,t.sort);
});

test('모든 표준 비교와 AND/OR 조건을 이름·숫자 값 그대로 보존한다',()=>{
  for(const [op,native] of Object.entries({eq:'equal',ne:'notEqual',gt:'greaterThan',ge:'greaterThanOrEqual',lt:'lessThan',le:'lessThanOrEqual'})){
    const f={criteria:{2:{type:'custom',op1:op,v1:'20',op2:'ne',v2:'A&"B',join:'or'}}};
    const x=autoFilterXml(f,range);assert.match(x,new RegExp(`operator="${native}"`));
    assert.deepEqual(readAutoFilter(parseXml(x),range).criteria,f.criteria);
  }
});

test('시작/끝은 리터럴 wildcard를 이스케이프하고 포함은 wildcard 의미를 유지한다',()=>{
  for(const [op,val,expected] of [['begins','~A*?','~~A~*~?*'],['ends','A?','*A~?'],['contains','~*A?','*~*A?*'],['notBegins','A','A*'],['notEnds','A','*A'],['notContains','A','*A*']]){
    const x=parseXml(autoFilterXml({criteria:{2:{type:'custom',op1:op,v1:val}}},range));
    const c=child(child(child(x,'filterColumn'),'customFilters'),'customFilter');assert.equal(c.attrs.val,expected);assert.equal(c.attrs.operator,op.startsWith('not')?'notEqual':'equal');
  }
});

test('네이티브 빈문자 custom·날짜그룹·아이콘·동적날짜는 삭제하지 않고 원형으로 보존한다',()=>{
  const fragments=['<customFilters><customFilter operator="equal" val=""/></customFilters>','<filters blank="1"><dateGroupItem year="2026" month="10" dateTimeGrouping="month"/></filters>','<iconFilter iconSet="3Arrows" iconId="1"/>','<dynamicFilter type="thisMonth" val="46000" maxVal="46030"/>'];
  for(const raw of fragments){
    const data=replaceFilter(fixture(),`<filterColumn colId="0" hiddenButton="1">${raw}</filterColumn>`),f=model(data);
    assert.equal(f.criteria[2].type,'xlsx');assert.deepEqual(f.hiddenButtons,{2:true});
    const out=after(new Workbook(data));assert.deepEqual(model(out).criteria,f.criteria);assert.deepEqual(model(out).hidden,f.hidden);
  }
});

test('상위/하위 항목·백분율·평균 조건의 기준을 표와 범위에서 보존한다',()=>{
  for(const table of [true,false])for(const cr of [{type:'top',n:3,bottom:false,percent:false},{type:'top',n:25,bottom:true,percent:true},{type:'avg',above:true},{type:'avg',above:false}]){
    const wb=fixture(table);(table?wb.sheets[0].tables[0].filter:wb.sheets[0].filter).criteria={3:cr};assert.deepEqual(model(after(wb),table).criteria[3],cr);
  }
});

test('채움/글꼴색 DXF를 새 풀에 등록하고 다른 서식과 섞여도 같은 색을 읽는다',()=>{
  for(const table of [true,false])for(const type of ['fill','font']){
    const wb=fixture(table),f=table?wb.sheets[0].tables[0].filter:wb.sheets[0].filter;f.criteria={3:{type,value:'#f0a020'}};
    wb.sheets[0].cond=[{r1:3,c1:3,r2:7,c2:3,type:'cell',op:'gt',v1:'5',style:{fill:'#00ff00'}}];
    const bytes=writeXlsx(wb),files=unzip(bytes),xml=textOf(files[table?'xl/tables/table1.xml':'xl/worksheets/sheet1.xml']);
    assert.match(xml,new RegExp(`cellColor="${type==='fill'?1:0}"`));assert.deepEqual(model(readXlsx(bytes).data,table).criteria[3],{type,value:'#f0a020'});
  }
});

test('Excel 글꼴색 필터의 fill 전경색은 흰 배경색으로 바뀌지 않는다',()=>{
  const wb=fixture(),files=unzip(writeXlsx(wb));
  files['xl/styles.xml']=textOf(files['xl/styles.xml']).replace(/<dxfs[^>]*>[\s\S]*?<\/dxfs>/,'<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill></dxf></dxfs>');
  files['xl/tables/table1.xml']=textOf(files['xl/tables/table1.xml']).replace(/<autoFilter[\s\S]*?<\/autoFilter>/,'<autoFilter ref="C3:F8"><filterColumn colId="1"><colorFilter dxfId="0" cellColor="0"/></filterColumn></autoFilter>');
  const data=readXlsx(zip(files)).data;assert.deepEqual(model(data).criteria[3],{type:'font',value:'#ff0000'});assert.deepEqual(model(after(new Workbook(data))).criteria[3],model(data).criteria[3]);
});

test('복합정렬의 순서·대소문자·원형 속성을 남기고 새 단일정렬은 이를 대체한다',()=>{
  const data=replaceFilter(fixture(),'', '', '<sortState ref="C4:F8" caseSensitive="1"><sortCondition descending="1" ref="D4:D8"/><sortCondition ref="F4:F8" customList="A,B,C"/></sortState>');
  const f=model(data);assert.equal(f.sort.col,3);assert.equal(f.sort.xlsx.children.length,2);assert.deepEqual(model(after(new Workbook(data))).sort,f.sort);
  f.sort={col:5,asc:true};const root=parseXml(textOf(unzip(writeXlsx(new Workbook(data)))['xl/tables/table1.xml']));assert.equal(kids(child(root,'sortState'),'sortCondition').length,1);assert.equal(child(root,'sortState').attrs.caseSensitive,undefined);
});

test('표준 표현 없는 정규식은 조건을 조용히 버리는 파일을 만들지 않는다',()=>{
  const wb=fixture();wb.sheets[0].tables[0].filter.criteria={2:{type:'custom',op1:'regex',v1:'^A'}};const before=wb.serialize();assert.throws(()=>writeXlsx(wb),/정규식/);assert.deepEqual(wb.serialize(),before);
});

test('숨긴 요약행의 표준 함수·사용자 수식·문자 레이블은 저장 뒤 다시 복원 가능하다',()=>{
  const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsFns={3:'average'};t.totalsCells={2:{raw:'=표시문자',inputType:'text'},4:{raw:'=SUM(D4:D8)+1'},5:null};
  const bytes=writeXlsx(wb),xml=textOf(unzip(bytes)['xl/tables/table1.xml']);assert.match(xml,/totalsRowFunction="average"/);assert.match(xml,/totalsRowLabel="=표시문자"/);assert.match(xml,/<totalsRowFormula>SUM\(D4:D8\)\+1<\/totalsRowFormula>/);
  const out=readXlsx(bytes).data.sheets[0].tables[0];assert.equal(out.totals,false);assert.deepEqual(out.totalsFns,{3:'average'});assert.equal(out.totalsCells[2].inputType,'text');assert.equal(out.totalsCells[4].raw,'=SUM(D4:D8)+1');
});

test('custom 플래그와 오래된 sum 플래그보다 실제 요약 수식/문자/빈칸을 우선한다',()=>{
 for(const visible of [false,true])for(const fn of ['sum','custom']){
  const wb=fixture(),t=wb.sheets[0].tables[0];t.totals=visible;t.totalsFns={3:fn,4:'sum',5:'sum'};
  const cells={3:{raw:'=SUM(D4:D8)+7'},4:{raw:'직접 입력',inputType:'text'},5:null};
  if(visible){t.r2++;for(const [c,data] of Object.entries(cells))wb.setCellData(0,t.r2,Number(c),data);}else t.totalsCells=cells;
  const bytes=writeXlsx(wb),root=parseXml(textOf(unzip(bytes)['xl/tables/table1.xml'])),cols=kids(child(root,'tableColumns'),'tableColumn');
  assert.equal(cols[1].attrs.totalsRowFunction,'custom');assert.equal(child(cols[1],'totalsRowFormula').text,'SUM(D4:D8)+7');assert.equal(cols[2].attrs.totalsRowFunction,undefined);assert.equal(cols[2].attrs.totalsRowLabel,'직접 입력');assert.equal(cols[3].attrs.totalsRowFunction,undefined);
  const actual=readXlsx(bytes).data.sheets[0].tables[0];assert.deepEqual(actual.totalsFns,{});if(!visible)assert.equal(actual.totalsCells[3].raw,'=SUM(D4:D8)+7');
 }
});

test('표준 SUBTOTAL 함수와 실제 참조가 일치할 때만 메뉴 함수로 저장한다',()=>{
 for(const [fn,code]of [['average',101],['sum',109],['countNums',102]]){
  const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsFns={3:fn};t.totalsCells={3:{raw:`=SUBTOTAL(${code},자료표[열3])`}};
  const col=kids(child(parseXml(textOf(unzip(writeXlsx(wb))['xl/tables/table1.xml'])),'tableColumns'),'tableColumn')[1];assert.equal(col.attrs.totalsRowFunction,fn);assert.equal(child(col,'totalsRowFormula'),null);
  t.totalsCells[3].raw=`=SUBTOTAL(${code},자료표[열4])`;const other=kids(child(parseXml(textOf(unzip(writeXlsx(wb))['xl/tables/table1.xml'])),'tableColumns'),'tableColumn')[1];assert.equal(other.attrs.totalsRowFunction,'custom');assert.ok(child(other,'totalsRowFormula'));
 }
});

test('명시적으로 모두 비운 숨긴 요약행은 처음부터 없는 요약행과 구별된다',()=>{
 const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsFns={};t.totalsCells={2:null,3:null,4:{raw:''},5:null};
 const out=after(wb).sheets[0].tables[0];assert.deepEqual(Object.keys(out.totalsCells),['2','3','4','5']);assert.ok(Object.values(out.totalsCells).every(c=>c.raw===''));
 const fresh=after(fixture()).sheets[0].tables[0];assert.equal(fresh.totalsCells,undefined);
});

test('머리글 없는 표의 저장된 정렬은 첫 데이터 행을 건너뛰지 않는다',()=>{
 const wb=fixture(),t=wb.sheets[0].tables[0];t.header=false;t.r1++;t.columns=['열2','열3','열4','열5'];t.sort=t.filter.sort;t.filter=null;
 const root=parseXml(textOf(unzip(writeXlsx(wb))['xl/tables/table1.xml']));assert.equal(child(root,'sortState').attrs.ref,'C4:F8');assert.equal(child(child(root,'sortState'),'sortCondition').attrs.ref,'D4:D8');
});

test('숨긴 요약셀의 명시 수식 fx/inputType는 텍스트 서식보다 우선한다',()=>{
 for(const data of [{raw:'=1+2',inputType:'text',fx:true},{raw:'=1+2',inputType:'value',style:{numFmt:'text'}}]){
  const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsCells={3:data};
  const bytes=writeXlsx(wb),cols=kids(child(parseXml(textOf(unzip(bytes)['xl/tables/table1.xml'])),'tableColumns'),'tableColumn');assert.equal(cols[1].attrs.totalsRowFunction,'custom');assert.equal(child(cols[1],'totalsRowFormula').text,'1+2');
  const restored=readXlsx(bytes).data.sheets[0].tables[0].totalsCells[3];assert.equal(restored.raw,'=1+2');assert.equal(restored.inputType,undefined);
 }
});

test('숨긴 요약 라벨·평균 형식·빈칸 서식은 totalsRowDxfId로 보존된다',()=>{
 const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsFns={3:'average'};
 t.totalsCells={2:{raw:'직접 요약',inputType:'text',style:{bold:false,color:'#123456',fill:'#abcdef'}},3:{raw:'=SUBTOTAL(101,자료표[열3])',style:{numFmt:'custom',code:'#,##0.0000;[Red](#,##0.0000)',font:'Arial',size:14}},4:{raw:'',style:{fill:'#fedcba',bb:true,bbs:'double',bbc:'#112233'}}};
 const bytes=writeXlsx(wb),root=parseXml(textOf(unzip(bytes)['xl/tables/table1.xml']));assert.equal(kids(child(root,'tableColumns'),'tableColumn').filter(c=>c.attrs.totalsRowDxfId!==undefined).length,3);
 const out=readXlsx(bytes).data.sheets[0].tables[0];
 for(const c of [2,3,4]){assert.equal(out.totalsCells[c].raw,t.totalsCells[c].raw);for(const [key,val]of Object.entries(t.totalsCells[c].style))assert.deepEqual(out.totalsCells[c].style[key],val,`${c}:${key}`);}
 assert.equal(out.totalsFns[3],'average');
 const second=after(new Workbook(readXlsx(bytes).data)).sheets[0].tables[0];assert.deepEqual(second.totalsCells,out.totalsCells);
});

test('이전에 표시한 총계행 플래그는 현재 숨김과 독립이며 빈 이력도 보존한다',()=>{
 const wb=fixture(),t=wb.sheets[0].tables[0];t.totalsCells={};
 let bytes=writeXlsx(wb);assert.match(textOf(unzip(bytes)['xl/tables/table1.xml']),/totalsRowShown="1"/);assert.deepEqual(readXlsx(bytes).data.sheets[0].tables[0].totalsCells,{});
 const files=unzip(bytes);files['xl/tables/table1.xml']=textOf(files['xl/tables/table1.xml']).replace(' totalsRowShown="1"','');assert.deepEqual(readXlsx(zip(files)).data.sheets[0].tables[0].totalsCells,{});
 delete t.totalsCells;bytes=writeXlsx(wb);assert.match(textOf(unzip(bytes)['xl/tables/table1.xml']),/totalsRowShown="0"/);assert.equal(readXlsx(bytes).data.sheets[0].tables[0].totalsCells,undefined);
});
