import {test} from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/workbook.js';
import {ERR, Range} from '../src/fxcore.js';
import {parseInput, dateParts, serialOf, formatValue, formatCode, formatQuery} from '../src/format.js';
import {readXlsx, writeXlsx, isoSerial} from '../src/xlsx.js';
import {unzip} from '../src/zip.js';
import {resolvePivot, pivotSourceData, computePivot, dateFilterMatch, pivotChartData} from '../src/pivot.js';
import {groupKey} from '../src/cube.js';
import {chartModelData, renderChartSvg} from '../src/chart.js';
import {checkValidation} from '../src/validation.js';
import {condMatch, prepareCond, periodRange} from '../src/condfmt.js';
import {publishedWorkbook} from '../src/publish.js';
import {WEB, NET, historyUrl, parseTicker} from '../src/fx-web.js';
const book=(date1904=false,cells={})=>new Workbook({date1904,sheets:[{name:'날짜',cells}]});
const calc=(wb,formula)=>{wb.setInput(0,100,20,formula);return wb.getValue(0,100,20);};
const xml=(files,name)=>new TextDecoder().decode(files[name]);

test('1904 날짜 체계: 0·윤년·시각·일반 숫자·경과시간은 독립된 서식으로 표시한다',()=>{
 assert.equal(serialOf(1904,1,1,true),0);assert.equal(serialOf(2024,2,29,true),43889);
 assert.deepEqual([dateParts(0,1,true).y,dateParts(0,1,true).m,dateParts(0,1,true).d,dateParts(0,1,true).dow],[1904,1,1,5]);
 for(const [n,text] of [[0,'1904-01-01'],[59,'1904-02-29'],[60,'1904-03-01'],[45322,'2028-02-01']])assert.equal(formatValue(n,{numFmt:'date'},true).text,text);
 assert.equal(formatCode(45322.5,'yyyy-mm-dd hh:mm:ss',true).text,'2028-02-01 12:00:00');
 assert.equal(formatCode(45322.5,'[h]:mm',true).text,'1087740:00');assert.equal(formatCode(45322.5,'0.0',true).text,'45322.5');
 assert.equal(formatQuery(0,'yyyy-MM-dd',true),'1904-01-01');
 assert.equal(parseInput('2024-02-29',true).value,43889);assert.equal(parseInput('2024-02-29 12:00',true).value,43889.5);
 assert.equal(parseInput('45322',true).value,45322);assert.equal(parseInput('12:00',true).value,.5);
});
test('두 통합 문서의 중첩 날짜·텍스트·배열·이름 LAMBDA가 서로 날짜 체계를 오염시키지 않는다',()=>{
 const a=book(),b=book(true);
 const cases=[['=DATE(2024,2,29)',45351,43889],['=DATE(1900,2,29)',60,ERR.NUM],['=YEAR(0)',1900,1904],['=DAY(0)',0,1],['=WEEKDAY(0)',7,6],['=TEXT(0,"yyyy-mm-dd")','1900-01-00','1904-01-01'],['=DATEVALUE("2024-02-29")',45351,43889],['=VALUE("2024-02-29")',45351,43889],['=EDATE(DATE(2024,1,31),1)',45351,43889],['=EOMONTH(DATE(2024,1,2),1)',45351,43889],['=WORKDAY(DATE(2024,2,28),2)',45352,43890],['=NETWORKDAYS(DATE(2024,2,28),DATE(2024,3,4))',4,4],['=DAYS360(DATE(2024,1,31),DATE(2024,3,31))',60,60],['=COUPNCD(DATE(2024,2,29),DATE(2028,12,31),2,0)',45473,44011],['=HOUR(0.5)',12,12],['=TIME(12,0,0)',.5,.5],['=YEAR(-1)',ERR.NUM,ERR.NUM]];
 for(const [formula,v0,v1]of cases){assert.equal(calc(b,formula),v1,formula+' 1904');assert.equal(calc(a,formula),v0,formula+' 1900');assert.equal(calc(b,formula),v1,formula+' 1904 repeat');}
 assert.equal(calc(b,'=LET(d,DATE(2024,2,29),YEAR(d))'),2024);
 assert.equal(calc(b,'=LAMBDA(d,TEXT(d,"yyyy-mm-dd"))(0)'),'1904-01-01');
 assert.equal(calc(b,'=YEAR({0,59})'),1904);assert.equal(b.getValue(0,100,21),1904);
});
test('휴일 배열과 재무 함수는 날짜 인수만 이동하고 일수·금액·금리는 그대로 둔다',()=>{
 for(const fn of ['=WORKDAY(DATE(2024,2,28),2,DATE(2024,2,29))','=WORKDAY.INTL(DATE(2024,2,28),2,1,DATE(2024,2,29))'])assert.equal(calc(book(true),fn),calc(book(),fn)-1462);
 for(const fn of ['=NETWORKDAYS(DATE(2024,2,28),DATE(2024,3,4),DATE(2024,2,29))','=NETWORKDAYS.INTL(DATE(2024,2,28),DATE(2024,3,4),1,DATE(2024,2,29))','=PRICE(DATE(2024,2,29),DATE(2028,12,31),0.05,0.06,100,2,0)','=YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),1)','=XNPV(0.1,VSTACK(-100,120),VSTACK(DATE(2024,1,1),DATE(2025,1,1)))']){const x=calc(book(),fn),y=calc(book(true),fn);assert.equal(typeof x,'number',fn);assert.ok(Math.abs(x-y)<1e-9,fn);}
});
test('날짜 체계 변경·Undo·Redo·청크/JSON 복원은 숫자 원본과 계산값을 올바르게 보존한다',()=>{
 const wb=book(false,{'0,0':{raw:'2024-02-29'},'0,1':{raw:'=YEAR(A1)'},'0,2':{raw:'45351'},'0,3':{raw:'=DATE(2024,2,29)'}});
 assert.equal(wb.getValue(0,0,0),45351);wb.transact(()=>wb.setDate1904(true));assert.equal(wb.date1904,true);assert.equal(wb.getValue(0,0,0),45351);assert.equal(wb.getValue(0,0,2),45351);assert.equal(wb.getValue(0,0,3),43889);assert.equal(wb.getValue(0,0,1),2028);
 wb.undo();assert.equal(wb.date1904,false);assert.equal(wb.getValue(0,0,1),2024);wb.redo();assert.equal(wb.date1904,true);assert.equal(wb.getValue(0,0,1),2028);
 const restored=new Workbook(JSON.parse(JSON.stringify(wb.serialize())));assert.equal(restored.date1904,true);assert.equal(restored.getValue(0,0,0),45351);
 const chunks=[...wb.cellChunks(0,2)].flat();const chunked=new Workbook({...wb.bookMeta(),sheets:[{name:'날짜',cells:new Map(chunks)}]});assert.equal(chunked.date1904,true);assert.equal(chunked.getValue(0,0,3),43889);
 wb.transact(()=>wb.setInput(0,1,0,'2024-02-29'));assert.equal(wb.getValue(0,1,0),43889);
});
test('XLSX 저장·가져오기에서 1904 메타·0 날짜·일반 숫자·수식 재계산이 왕복된다',()=>{
 const wb=book(true,{'0,0':{raw:'0',style:{numFmt:'date'}},'1,0':{raw:'43889',style:{numFmt:'date'}},'2,0':{raw:'43889'},'3,0':{raw:'=DATE(2024,2,29)',style:{numFmt:'date'}},'4,0':{raw:'1.5',style:{numFmt:'custom',code:'[h]:mm'}}});
 const bytes=writeXlsx(wb),files=unzip(bytes);assert.match(xml(files,'xl/workbook.xml'),/<workbookPr[^>]*date1904="1"/);
 const read=readXlsx(bytes);assert.equal(read.data.date1904,true);assert.equal(read.warnings.some(x=>/1904/.test(x)),false);
 const copy=new Workbook(read.data);copy.invalidate(0);for(let r=0;r<5;r++)assert.equal(copy.getValue(0,r,0),wb.getValue(0,r,0));assert.equal(formatValue(copy.getValue(0,0,0),copy.styleAt(0,0,0),copy.date1904).text,'1904-01-01');
 assert.equal(isoSerial('1904-01-01T12:00:00',true),.5);assert.equal(new Workbook(readXlsx(writeXlsx(copy)).data).date1904,true);
});
test('피벗 날짜 그룹·날짜 필터·차트와 XLSX 그룹 경계는 1904 연도를 사용한다',()=>{
 const wb=book(true,{'0,0':{raw:'날짜'},'0,1':{raw:'값'},'1,0':{raw:'0',style:{numFmt:'date'}},'1,1':{raw:'2'},'2,0':{raw:'59',style:{numFmt:'date'}},'2,1':{raw:'3'}});
 const def={source:'날짜',range:{r1:0,c1:0,r2:2,c2:1},top:5,left:0,rows:['날짜'],cols:[],values:[{field:'값',agg:'sum'}],groups:{날짜:{by:'years',start:0,end:59}}};
 const src=pivotSourceData(wb,def),res=resolvePivot(src,def);assert.equal(res.def.date1904,true);assert.equal(res.def.groups.날짜.date1904,true);
 assert.ok(JSON.stringify(computePivot(res,res.def).grid).includes('1904'));assert.equal(groupKey(59,{by:'days',date1904:true}),'1904-02-29');
 assert.equal(dateFilterMatch('thisMonth',59,null,null,31,true),true);assert.equal(dateFilterMatch('thisWeek',2,null,null,3,true),true);
 const data=pivotChartData(src,{...def,groups:{}},()=>({numFmt:'date'}));assert.deepEqual(data.categories,['1904-01-01','1904-02-29']);
 wb.sheets[0].pivot=def;const files=unzip(writeXlsx(wb));const cache=Object.keys(files).find(x=>/pivotCacheDefinition\d+\.xml$/.test(x));assert.ok(cache);assert.match(xml(files,cache),/startDate="1904-01-01T00:00:00"/);assert.match(xml(files,cache),/endDate="1904-02-29T00:00:00"/);
 const again=new Workbook(readXlsx(writeXlsx(wb)).data);assert.equal(again.sheets[0].pivot.groups.날짜.start,0);
});
test('QUERY의 날짜 리터럴·날짜 함수·no_values 표시에도 Workbook context가 전달된다',()=>{
 const wb=book(true,{'0,0':{raw:'0'},'1,0':{raw:'59'}});
 assert.equal(calc(wb,'=QUERY(A1:A2,"select year(A) where A >= date \'1904-02-01\' label year(A) \'\'",0)'),1904);
 assert.equal(calc(wb,'=QUERY(A1:A2,"select A format A \'yyyy-MM-dd\' options no_values",0)'),'1904-01-01');
 assert.equal(calc(wb,'=EPOCHTODATE(0)')+1462,calc(book(),'=EPOCHTODATE(0)'));
});
test('검증·조건부 서식·게시 스냅숏이 1904 날짜 숫자를 다른 날짜로 바꾸지 않는다',()=>{
 const wb=book(true,{'0,0':{raw:'0',style:{numFmt:'date'}}});
 const rule={r1:0,c1:0,r2:0,c2:0,type:'date',op:'equal',f1:'1904-01-01'};assert.equal(checkValidation(wb,0,rule,0,0,'0'),true);assert.equal(checkValidation(wb,0,rule,0,0,'1904-01-01'),true);
 wb.sheets[0].cond=[{r1:0,c1:0,r2:0,c2:0,type:'eq',v1:'1904-01-01',style:{fill:'#ff0000'}}];assert.equal(condMatch(prepareCond(wb,0)[0],0,wb),true);
 assert.equal(periodRange('today',new Date(2024,1,29),true)[0],43889);
 const shared=new Workbook(publishedWorkbook(wb,0));assert.equal(shared.date1904,true);assert.equal(shared.getValue(0,0,0),0);assert.equal(formatValue(0,shared.styleAt(0,0,0),shared.date1904).text,'1904-01-01');
});
test('차트 항목·날짜 축·레이블과 저장 차트의 날짜 체계가 일치한다',()=>{
 const wb=book(true,{'0,0':{raw:'날짜'},'0,1':{raw:'값'},'1,0':{raw:'0',style:{numFmt:'date'}},'1,1':{raw:'59'}});
 const ch={type:'column',range:{r1:0,c1:0,r2:1,c2:1},w:500,h:300,labels:true,axes:{y:{numFmt:'yyyy-mm-dd'}},seriesFmt:{0:{numFmt:'yyyy-mm-dd'}}};
 const data=chartModelData(wb,0,ch);assert.equal(data.date1904,true);assert.deepEqual(data.categories,['1904-01-01']);assert.equal(data.series.length,1);const svg=renderChartSvg(ch,data);assert.match(svg,/1904-02-29/);assert.match(svg,/1904-/);
 wb.sheets[0].charts=[ch];const files=unzip(writeXlsx(wb));assert.match(xml(files,'xl/charts/chart1.xml'),/<c:date1904 val="1"/);
});

test('빈 휴일 참조는 1904-01-01을 휴일로 만들지 않으며 상한 날짜는 오류로 막는다',()=>{
 const wb=book(true);assert.equal(calc(wb,'=NETWORKDAYS(0,0,A1:A2)'),1);assert.equal(calc(wb,'=WORKDAY(0,1,A1:A2)'),3);
 assert.equal(calc(wb,'=DATE(9999,12,31)'),2957003);assert.equal(calc(wb,'=DATE(9999,12,32)'),ERR.NUM);assert.equal(calc(wb,'=YEAR(2957004)'),ERR.NUM);
});
test('시세 캐시를 공유해도 날짜 열만 각 통합 문서의 체계로 변환하고 가격은 유지한다',()=>{
 const tk=parseTicker('AAPL'),url=historyUrl(tk,45351,45351);const previous=NET.cache.get(url),auth=NET.authorize;
 NET.authorize=null;NET.cache.set(url,{state:'ok',t:Date.now(),sheets:new Set(),data:'Date,Open,High,Low,Close,Volume\n2024-02-29,10,12,9,11,100'});
 try{const a=WEB.STOCKHISTORY(['AAPL',45351,45351,0,0],{date1904:false}),b=WEB.STOCKHISTORY(['AAPL',43889,43889,0,0],{date1904:true});assert.ok(a instanceof Range);assert.ok(b instanceof Range);assert.deepEqual(a.rows,[[45351,11]]);assert.deepEqual(b.rows,[[43889,11]]);const g=WEB.GOOGLEFINANCE(['AAPL','close',43889,43889],{date1904:true});assert.deepEqual(g.rows,[['Date','Close'],[43889,11]]);}finally{NET.authorize=auth;if(previous)NET.cache.set(url,previous);else NET.cache.delete(url);}
});

test('Excel 경계 oracle: 1904 음수 날짜·시간 표시와 함수별 음수 결과 정책을 구분한다',()=>{
 const wb=book(true);assert.equal(formatCode(-.5,'hh:mm:ss',true).text,'-12:00:00');assert.equal(formatCode(-.5,'[h]:mm',true).text,'-12:00');assert.equal(formatCode(-1,'yyyy-mm-dd hh:mm:ss',true).text,'-1904-01-02 00:00:00');assert.equal(formatCode(-.5,'[h]:mm').text,'########');assert.equal(formatCode(2957004,'[h]:mm',true).text,'########');
 assert.equal(formatValue(-1,{numFmt:'date'},true).text,'-1904-01-02');assert.equal(formatValue(2957004,{numFmt:'date'},true).text,'########');
 assert.equal(calc(wb,'=DATEVALUE("1903-12-31")'),ERR.VALUE);assert.equal(calc(wb,'=DATEVALUE("12:00")'),0);assert.equal(calc(wb,'=EDATE(0,-1)'),-31);assert.equal(calc(wb,'=EOMONTH(0,-1)'),-1);assert.equal(calc(wb,'=COUPPCD(0,365,2)'),-1);assert.equal(calc(wb,'=WORKDAY(0,-1)'),ERR.NUM);assert.equal(calc(book(),'=EDATE(0,-1)'),ERR.NUM);assert.equal(calc(book(),'=EOMONTH(0,-1)'),ERR.NUM);
});
