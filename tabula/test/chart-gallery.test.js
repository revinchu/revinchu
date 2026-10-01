import test from 'node:test';
import assert from 'node:assert/strict';
import { chartDataGuide, chartPresetMatches } from '../src/chart-ui.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { chartModelData } from '../src/chart.js';
import { unzip, zip, textOf } from '../src/zip.js';
function book(patch, rows) {
 const wb = new Workbook();
 wb.transact(() => { rows.forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))); wb.setSheetProp(0,'charts',[{id:'c',type:'column',x:10,y:10,w:500,h:320,range:{r1:0,c1:0,r2:rows.length-1,c2:rows[0].length-1},...patch}]); }); return wb;
}
function load(bytes) { const wb=new Workbook(); wb.restore(readXlsx(bytes).data); return wb; }
function noExtension(bytes) {
 const files=unzip(bytes), name=Object.keys(files).find(n=>/^xl\/charts\/chart\d+\.xml$/.test(n));
 assert.ok(name); const xml=textOf(files[name]); files[name]=xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,''); return {xml,bytes:zip(files)};
}
test('차트 자료 안내와 하위 유형 선택을 구분',()=>{
 for(const [p,count] of [[{},3],[{ohlc:true},4],[{volume:true},4],[{volume:true,ohlc:true},5]]) {
  assert.equal(chartDataGuide({type:'stock',...p},{categories:['A'],series:Array(count).fill({})}).error,'');
  assert.match(chartDataGuide({type:'stock',...p},{series:[]}).error,new RegExp(`${count}개`));
 }
 assert.ok(chartDataGuide({type:'surface'},{categories:['A'],series:[{}]}).error);
 assert.ok(chartDataGuide({type:'pieOfPie'},{categories:['A','B'],series:[{}]}).error);
 assert.ok(chartPresetMatches({type:'stock',volume:true,ohlc:true},{type:'stock',threeD:false,volume:true,ohlc:true}));
 assert.ok(!chartPresetMatches({type:'surface',surfaceStyle:'contour'},{type:'surface',surfaceStyle:'surface'}));
});
test('보조 원형·막대형: 자체 확장 없이 표준 OOXML 종류·분할 복원',()=>{
 for(const type of ['pieOfPie','barOfPie']) {
  const wb=book({type,splitType:'custom',splitPoints:[2,3],splitPos:2,secondSize:90,splitGap:120},[['항목','값'],['A',30],['B',25],['C',5],['D',10]]);
  const {xml,bytes}=noExtension(writeXlsx(wb)); assert.match(xml,/<c:ofPieChart>/); assert.match(xml,/<c:custSplit>/);
  const chart=load(bytes).sheets[0].charts[0]; assert.equal(chart.type,type);assert.equal(chart.splitType,'custom');assert.deepEqual(chart.splitPoints,[2,3]);assert.equal(chart.secondSize,90);assert.equal(chart.splitGap,120);
 }
});
test('표면형4종: 표준 유형·축·행 방향 참조 보존',()=>{
 for(const surfaceStyle of ['surface','wireframe','contour','wireframeContour']) {
  const threeD=['surface','wireframe'].includes(surfaceStyle),wb=book({type:'surface',surfaceStyle,threeD},[['온도',10,20,30],['고도1',5,20,9],['고도2',10,4,30]]);
  const before=chartModelData(wb,0,wb.sheets[0].charts[0]),{xml,bytes}=noExtension(writeXlsx(wb));assert.match(xml,threeD?/<c:surface3DChart>/ : /<c:surfaceChart>/);assert.match(xml,/<c:serAx>/);
  const round=load(bytes),chart=round.sheets[0].charts[0],after=chartModelData(round,0,chart);assert.equal(chart.type,'surface');assert.equal(chart.surfaceStyle,surfaceStyle);assert.deepEqual(after.categories,before.categories);assert.deepEqual(after.series.map(s=>s.values),before.series.map(s=>s.values));
 }
});
test('거래량주식: 표준 거래량 막대·가격 차트의 두 축 보존',()=>{
 for(const ohlc of [false,true]) {
  const rows=ohlc?[['일','거래량','시가','고가','저가','종가'],['월',1000,20,30,10,25],['화',2000,25,35,15,30]]:[['일','거래량','고가','저가','종가'],['월',1000,30,10,25],['화',2000,35,15,30]];
  const wb=book({type:'stock',volume:true,ohlc},rows),{xml,bytes}=noExtension(writeXlsx(wb));assert.match(xml,/<c:barChart>/);assert.match(xml,/<c:stockChart>/);assert.match(xml,/<c:axId val="444444444"/);
  const round=load(bytes),chart=round.sheets[0].charts[0];assert.equal(chart.type,'stock');assert.equal(chart.volume,true);assert.equal(chart.ohlc,ohlc);assert.deepEqual(chartModelData(round,0,chart).series.map(s=>s.values),chartModelData(wb,0,wb.sheets[0].charts[0]).series.map(s=>s.values));
 }
});


test('표식 크기·분산 선 유형·방사형 옵션은 표준 XML에서 유지',()=>{
 for(const patch of [
  {type:'scatter',scatterStyle:'smoothMarker',seriesFmt:[{marker:'diamond',markerSize:12}]},
  {type:'radar',radarStyle:'marker',seriesFmt:[{marker:'circle',markerSize:9}]},
  {type:'radar',radarStyle:'filled'},
 ]) {
  const wb=book(patch,[['항목','값'],['A',3],['B',5],['C',8]]);
  const {xml,bytes}=noExtension(writeXlsx(wb)),chart=load(bytes).sheets[0].charts[0];
  assert.equal(chart.type,patch.type);
  if(patch.scatterStyle) assert.equal(chart.scatterStyle,patch.scatterStyle);
  if(patch.radarStyle) assert.equal(chart.radarStyle,patch.radarStyle);
  if(patch.seriesFmt) { assert.match(xml,new RegExp(`<c:size val="${patch.seriesFmt[0].markerSize}"`)); assert.equal(chart.seriesFmt[0].markerSize,patch.seriesFmt[0].markerSize); }
 }
});


test('보조 차트 네 분할 기준은 표준 enum으로 쓰고 모델로 복원',()=>{
 for(const [splitType,xmlValue] of [['position','pos'],['value','val'],['percent','percent'],['custom','cust']]) {
  const wb=book({type:'pieOfPie',splitType,splitPoints:[2]},[['항목','값'],['A',30],['B',25],['C',5]]);
  const {xml,bytes}=noExtension(writeXlsx(wb)); assert.match(xml,new RegExp(`<c:splitType val="${xmlValue}"`));assert.equal(load(bytes).sheets[0].charts[0].splitType,splitType);
 }
});
