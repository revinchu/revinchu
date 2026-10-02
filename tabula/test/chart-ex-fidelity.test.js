import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { chartModelData } from '../src/chart.js';
import { readChartEx, writeChartEx } from '../src/chart-ex.js';

const txt = (value) => typeof value === 'string' ? value : textOf(value);
const hierarchy = [['상위','중간','항목','금액'],['가','공통','첫째',10],['가','공통','둘째',20],['나','공통','첫째',30],['나','','',7]];
const normal = [['분류','첫째','둘째','셋째'],['가',10,2,100],['나',20,3,200],['다',30,4,300]];
function book(type, props = {}, rows = hierarchy) {
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row,r) => row.forEach((v,c) => wb.setInput(0,r,c,String(v)))));
  wb.sheets[0].charts=[{id:'fidelity',type,range:{r1:0,c1:0,r2:rows.length-1,c2:rows[0].length-1},x:0,y:0,w:640,h:400,...props}];
  return wb;
}
function standard(wb) {
  const files=unzip(writeXlsx(wb));
  for(const key of Object.keys(files)) if (/^xl\/(drawings|charts)\/[^/]+\.xml$/.test(key)) files[key]=txt(files[key]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'').replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
  return files;
}
function read(files) {const wb=new Workbook(readXlsx(zip(files)).data);return {wb,chart:wb.sheets[0].charts[0]};}

test('계층 ChartEx는 자체 메타 없이 레이블 세 종류·위치·글꼴·표시 형식을 보존',()=>{
  for(const type of ['treemap','sunburst']) {
    const sf={labels:true,catName:false,serName:true,labelPos:'insideEnd',labelSize:15,labelColor:'#123456',labelBold:false,numFmt:'#,##0.00',pointColors:{1:'#abcdef'}};
    const before=book(type,{seriesFmt:[sf],legend:'r',legendSize:12,legendColor:'#456789',legendBold:true});
    const files=standard(before), {wb,chart}=read(files);
    for(const [key,value] of Object.entries(sf)) assert.deepEqual(chart.seriesFmt[0][key],value,key);
    assert.equal(chart.legend,'r'); assert.equal(chart.legendSize,12); assert.equal(chart.legendColor,'#456789'); assert.equal(chart.legendBold,true);
    const data=chartModelData(wb,0,chart), original=chartModelData(before,0,before.sheets[0].charts[0]);
    assert.deepEqual(data.categories,original.categories); assert.deepEqual(data.catLevels,original.catLevels);
    assert.deepEqual(data.series[0].values,[10,20,30,7]);
    assert.equal(child(kids(descendants(parseXml(txt(files['xl/charts/chart1.xml'])),'plotAreaRegion')[0],'series')[0],'spPr'),null,'암묵 단일색이 계층 팔레트를 덮지 않음');
  }
});

test('트리맵 상위 레이블 배너·겹침·없음은 표준 parentLabelLayout으로 왕복',()=>{
  for(const treemapLabelLayout of ['banner','overlapping','none']) {
    const files=standard(book('treemap',{treemapLabelLayout}));
    assert.equal(read(files).chart.treemapLabelLayout,treemapLabelLayout);
    assert.match(txt(files['xl/charts/chart1.xml']),new RegExp(`parentLabelLayout val="${treemapLabelLayout}"`));
  }
});

test('레이블을 모두 끄는 명시 false와 범주만 표시하는 기본값을 구분',()=>{
  const data={categories:['가'],series:[{name:'금액',values:[10],labels:false,catName:false,serName:false}]};
  const off=readChartEx(parseXml(writeChartEx({type:'sunburst'},data)));
  assert.equal(off.seriesFmt[0].labels,false); assert.equal(off.seriesFmt[0].catName,false);
  const defaults=readChartEx(parseXml(writeChartEx({type:'sunburst'},{categories:['가'],series:[{name:'금액',values:[10]}]})));
  assert.equal(defaults.seriesFmt[0].labels,false); assert.equal(defaults.seriesFmt[0].catName,true);
});

test('ChartEx 색상 관계의 실제 색·테마색을 팔레트로 읽고 다시 저장',()=>{
  const files=standard(book('treemap'));
  files['xl/charts/colors1.xml']='<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:srgbClr val="112233"/><a:schemeClr val="accent2"/><cs:variation/></cs:colorStyle>';
  const first=read(files); assert.equal(first.chart.palette[0],'#112233');
  assert.equal(first.chart.palette[1],'#ed7d31');
  assert.deepEqual(read(standard(first.wb)).chart.palette,first.chart.palette);
});

test('native 레이블·팔레트 수정이 오래된 보조 옵션보다 우선, 부모 색은 보조 보존',()=>{
  const key=JSON.stringify(['가']);
  const before=book('treemap',{palette:['#112233','#445566'],seriesFmt:[{catName:true,labelSize:10,hierarchyColors:{[key]:'#ff8800'}}]});
  const files=unzip(writeXlsx(before));
  files['xl/charts/chart1.xml']=txt(files['xl/charts/chart1.xml']).replace('categoryName="1"','categoryName="0"').replace('sz="1000"','sz="1700"');
  files['xl/charts/colors1.xml']=txt(files['xl/charts/colors1.xml']).replace('112233','AABBCC');
  const {chart}=read(files); assert.equal(chart.seriesFmt[0].catName,false); assert.equal(chart.seriesFmt[0].labelSize,17); assert.equal(chart.palette[0],'#aabbcc'); assert.deepEqual(chart.seriesFmt[0].hierarchyColors,{[key]:'#ff8800'});
});

test('3D 막대 모양·깊이 계열 grouping은 표준 XML/계열별 덮어쓰기로 보존',()=>{
  for(const barShape of ['box','cylinder','cone','pyramid']) {
    const files=standard(book('column',{threeD:true,grouping:'standard',barShape,seriesFmt:[{barShape:'cone'}]},normal));
    const xml=parseXml(txt(files['xl/charts/chart1.xml'])); const group=descendants(xml,'bar3DChart')[0];
    assert.equal(child(group,'shape').attrs.val,barShape); assert.equal(child(kids(group,'ser')[0],'shape').attrs.val,'cone');
    assert.equal(kids(group,'axId').length,3); assert.equal(descendants(xml,'serAx').length,1);
    const ch=read(files).chart; assert.equal(ch.barShape,barShape); assert.equal(ch.grouping,'standard'); assert.equal(ch.seriesFmt[0].barShape,'cone');
  }
});

test('같은 축·종류의 서로 다른 누적 그룹은 합치지 않고 각 계열 모드·값 유지',()=>{
  const before=book('combo',{seriesFmt:[{type:'column',axis:0,grouping:'stacked'},{type:'column',axis:0,grouping:'clustered'},{type:'line',axis:1,grouping:'clustered'}]},normal);
  const files=standard(before), xml=parseXml(txt(files['xl/charts/chart1.xml']));
  assert.deepEqual(descendants(xml,'barChart').map(g=>child(g,'grouping').attrs.val),['stacked','clustered']);
  const {wb,chart}=read(files); assert.deepEqual(chart.seriesFmt.map(s=>s.grouping),['stacked','clustered','clustered']);
  assert.deepEqual(chartModelData(wb,0,chart).series.map(s=>s.values),[[10,20,30],[2,3,4],[100,200,300]]);
});


test('고전 차트 표준 레이블은 값 false에도 범주·계열명·백분율과 글꼴을 보존',()=>{
  for(const type of ['column','line','pie','doughnut']) {
    const sf={labels:false,catName:true,serName:true,pct:true,labelSize:13,labelBold:false,labelColor:'#234567',labelPos:'center',numFmt:'0.00'};
    const files=standard(book(type,{seriesFmt:[sf]},normal));
    const xml=parseXml(txt(files['xl/charts/chart1.xml'])), labels=descendants(xml,'dLbls')[0];
    for(const flag of ['showCatName','showSerName','showPercent'])assert.equal(child(labels,flag).attrs.val,'1');
    assert.equal(child(labels,'showVal').attrs.val,'0');assert.equal(descendants(labels,'defRPr')[0].attrs.b,'0');
    const ch=read(files).chart;for(const key of ['labels','catName','serName','pct','labelSize','labelBold','labelColor','numFmt'])assert.equal(ch.seriesFmt[0][key],sf[key],type+' '+key);
  }
});

test('WIXEL 프리셋 이름은 native 색이 같은 경우만 복원하며 Excel 색 변경은 유지',()=>{
  const wb=book('treemap',{palette:'modern'}), files=unzip(writeXlsx(wb));
  assert.equal(read(files).chart.palette,'modern');
  assert.ok(Array.isArray(read(standard(wb)).chart.palette),'자체 정보가 없는 외부 파일은 실제 색 목록');
  files['xl/charts/colors1.xml']=txt(files['xl/charts/colors1.xml']).replace(/<a:srgbClr val="[A-F0-9]+"\/>/,'<a:srgbClr val="ABCDEF"/>');
  const edited=read(files).chart;assert.ok(Array.isArray(edited.palette));assert.equal(edited.palette[0],'#abcdef','오래된 프리셋 이름으로 Excel 색 수정 내용을 덮지 않는다');
});
