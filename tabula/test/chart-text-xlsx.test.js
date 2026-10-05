import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { chartTextStyle } from '../src/chart-text-format.js';
import { readChartTextFont } from '../src/chart-text-xml.js';

const xmlText = v => typeof v === 'string' ? v : textOf(v);
const fields = ['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike'];
const prefixed = (style, prefix) => Object.fromEntries(Object.entries(style).map(([key, value]) => [prefix + key[0].toUpperCase() + key.slice(1), value]));
const fonts = [
  {font:'Noto Sans KR',size:12.5,color:'#123456',bold:true,italic:false,underline:true,strike:false},
  {font:'나눔명조',size:24,color:'#aabbcc',bold:false,italic:true,underline:false,strike:true},
  {font:'Arial',size:15,color:'#6b214f',bold:true,italic:true,underline:true,strike:true},
  {font:'Pretendard',size:9,color:'#010203',bold:false,italic:false,underline:false,strike:false},
];
function book(type='combo') {
  const wb=new Workbook();
  const rows=[['항목','광고비','매출'],['가',10,100],['나',20,200],['다',30,300]];
  wb.transact(()=>rows.forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))));
  const chart={id:'text',type,x:0,y:0,w:600,h:400,title:'텍스트 서식',legend:'b',labels:true,
    range:{r1:0,c1:0,r2:3,c2:2},...prefixed(fonts[1],'title'),...prefixed(fonts[2],'legend'),
    font:fonts[0].font,size:fonts[0].size,textColor:fonts[0].color,bold:fonts[0].bold,italic:fonts[0].italic,underline:fonts[0].underline,strike:fonts[0].strike,
    axes:Object.fromEntries(['x','y','y2'].map((key,i)=>[key,{...fonts[i+1],title:'축 '+key,...prefixed(fonts[(i+2)%4],'title')}])),
    seriesFmt:[{...prefixed(fonts[2],'label'),pointLabelStyles:{1:fonts[3]},type:'column'},{...prefixed(fonts[3],'label'),type:'line',axis:'secondary'}],
    dataTable:true,dataTableText:fonts[1]};
  if(type!=='combo')for(const sf of chart.seriesFmt){delete sf.type;delete sf.axis;}
  wb.sheets[0].charts=[chart];return wb;
}
function exported(wb, strip=true) {
  const files=unzip(writeXlsx(wb));
  if(strip)for(const key of Object.keys(files))if(/^xl\/(?:charts|drawings)\/[^/]+\.xml$/.test(key))files[key]=xmlText(files[key]).replace(/<(?:\w+:)?extLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?extLst>/g,'');
  return files;
}
const chartOf = files => readXlsx(zip(files)).data.sheets[0].charts[0];
const expectStyle = (chart,part,expected) => { const actual=chartTextStyle(chart,part); for(const field of fields)assert.equal(actual[field],expected[field],`${part?.kind??'chart'} ${field}`); };

for(const type of ['combo','bar','scatter','pie','doughnut','radar','sunburst','treemap','waterfall','pareto','boxWhisker'])test(`${type}: 차트 텍스트 서식은 WIXEL 확장 없이 Excel 표준 XML로 왕복`,()=>{
  const wb=book(type), before=wb.serialize(), files=exported(wb), chart=chartOf(files);
  expectStyle(chart,null,fonts[0]);expectStyle(chart,{kind:'title'},fonts[1]);expectStyle(chart,{kind:'legend'},fonts[2]);
  expectStyle(chart,{kind:'label',s:0},fonts[2]);expectStyle(chart,{kind:'label',s:0,p:1},fonts[3]);
  if(!['pie','doughnut','sunburst','treemap'].includes(type))for(const [i,key] of ['x','y'].entries()) {
    expectStyle(chart,{kind:'axis-'+key},fonts[i+1]);expectStyle(chart,{kind:'axis-title-'+key},fonts[(i+2)%4]);
  }
  if(type==='combo'||type==='pareto') {expectStyle(chart,{kind:'axis-y2'},fonts[3]);expectStyle(chart,{kind:'axis-title-y2'},fonts[0]);}
  if(!['sunburst','treemap','waterfall','pareto','boxWhisker','pie','doughnut'].includes(type))expectStyle(chart,{kind:'dataTable'},fonts[1]);
  const root=parseXml(xmlText(files['xl/charts/chart1.xml']));
  assert.equal(descendants(root,'props').length,0);
  assert.ok(descendants(root,'latin').some(n=>n.attrs.typeface==='Noto Sans KR'));
  assert.ok(descendants(root,'ea').some(n=>n.attrs.typeface==='나눔명조'));
  assert.deepEqual(wb.serialize(),before,'저장이 원본 통합 문서를 바꾸지 않음');
});

test('외부 Excel 문서의 단락 서식과 run의 명시 false·글꼴을 함께 읽음',()=>{
  const node=parseXml('<c:tx><c:rich><a:p><a:pPr><a:defRPr sz="1350" b="1" i="1" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:r><a:rPr b="false" i="0" u="none" strike="noStrike"><a:ea typeface="나눔고딕"/></a:rPr><a:t>축 제목</a:t></a:r></a:p></c:rich></c:tx>');
  const style=readChartTextFont(node,n=>child(n,'srgbClr')?.attrs.val?'#'+child(n,'srgbClr').attrs.val:null);
  assert.deepEqual(style,{size:13.5,bold:false,italic:false,underline:false,strike:false,color:'#112233',font:'나눔고딕'});
});

test('Excel이 바꾼 native 제목 크기와 색이 이전 WIXEL 확장보다 우선',()=>{
  for(const type of ['column','sunburst']) {
    const files=exported(book(type),false), key='xl/charts/chart1.xml';
    files[key]=xmlText(files[key]).replaceAll('sz="2400"','sz="3200"').replaceAll('val="AABBCC"','val="FEDCBA"');
    const chart=chartOf(files);assert.equal(chart.titleSize,32);assert.equal(chart.titleColor,'#fedcba');
  }
});

test('숨긴 범주 원본 참조와 계열 숨김 방식에 맞춰 개별 레이블 번호를 보존',()=>{
  for(const type of ['column','waterfall']) {
    const wb=book(type), chart=wb.sheets[0].charts[0];chart.hiddenSeries=[0];chart.hiddenCats=[0];
    chart.seriesFmt[1].pointLabelStyles={1:fonts[1],2:fonts[2]};
    const back=chartOf(exported(wb));
    const index=type==='waterfall'?1:0;
    expectStyle(back,{kind:'label',s:index,p:1},fonts[1]);expectStyle(back,{kind:'label',s:index,p:2},fonts[2]);
    assert.equal(back.seriesFmt[index].pointLabelStyles[0],undefined);
  }
});

test('축 레이블 글꼴과 회전은 하나의 txPr 안에서 함께 보존',()=>{
  const wb=book('column');wb.sheets[0].charts[0].axes.x.labelRotation=-45;
  const files=exported(wb), root=parseXml(xmlText(files['xl/charts/chart1.xml']));
  const axis=descendants(root,'catAx')[0];assert.equal(kids(axis,'txPr').length,1);assert.equal(child(child(axis,'txPr'),'bodyPr').attrs.rot,'-2700000');
  const back=chartOf(files);assert.equal(back.axes.x.labelRotation,-45);assert.equal(back.axes.x.font,fonts[1].font);
});

test('단일 레이블에서 일부 속성만 바꿀 때 Excel에도 계열의 나머지 서식을 명시한다',()=>{
  for(const type of ['column','sunburst']) {
    const wb=book(type), source=wb.sheets[0].charts[0];
    source.seriesFmt[0].pointLabelStyles={1:{font:'Courier New',italic:false}};
    const files=exported(wb), root=parseXml(xmlText(files['xl/charts/chart1.xml']));
    const point=descendants(root,type==='sunburst'?'dataLabel':'dLbl')[0], properties=descendants(child(point,'txPr'),'defRPr')[0];
    assert.equal(properties.attrs.b,'1','Excel 개별 레이블이 기본 글꼴로 돌아가지 않음');assert.equal(properties.attrs.i,'0');
    expectStyle(chartOf(files),{kind:'label',s:0,p:1},{...fonts[2],font:'Courier New',italic:false});
  }
});

test('글꼴 이름의 XML 특수문자는 실행 가능한 마크업 없이 그대로 왕복',()=>{
  const name='Custom "A&B" <font>';
  const wb=book('column');wb.sheets[0].charts[0].titleFont=name;
  const files=exported(wb), xml=xmlText(files['xl/charts/chart1.xml']);
  assert.match(xml,/Custom &quot;A&amp;B&quot; &lt;font&gt;/);assert.equal(chartOf(files).titleFont,name);
  assert.equal(descendants(parseXml(xml),'font').length,0);
});

test('하나의 축만 바꾼 크기가 저장 후 다른 축에 전파되지 않는다',()=>{
  for(const type of ['column','scatter','waterfall','pareto'])for(const axis of ['x','y']) {
    const wb=book(type), chart=wb.sheets[0].charts[0];
    delete chart.size;delete chart.axisSize;chart.axes={[axis]:{size:27,font:'Courier New'}};
    const back=chartOf(exported(wb));
    assert.equal(chartTextStyle(back,{kind:'axis-'+axis}).size,27);
    assert.equal(chartTextStyle(back,{kind:'axis-'+(axis==='x'?'y':'x')}).size,9);
    assert.equal(back.axisSize,undefined);
  }
});

test('숨긴 제목·범례·축 제목·레이블·데이터표의 서식은 데이터 없는 보조 정보로 유지',()=>{
  for(const type of ['column','waterfall']) {
    const wb=book(type), source=wb.sheets[0].charts[0];
    source.title='';source.legend='none';source.dataTable=false;delete source.labels;
    for(const axis of Object.values(source.axes))delete axis.title;
    const back=chartOf(exported(wb,false));
    expectStyle(back,{kind:'title'},fonts[1]);expectStyle(back,{kind:'legend'},fonts[2]);
    expectStyle(back,{kind:'axis-title-x'},fonts[2]);expectStyle(back,{kind:'dataTable'},fonts[1]);
    expectStyle(back,{kind:'label',s:0},fonts[2]);expectStyle(back,{kind:'label',s:0,p:1},fonts[3]);
  }
});

test('표시 중인 개별 레이블을 Excel에서 초기화하면 이전 보조 서식이 부활하지 않는다',()=>{
  for(const type of ['column','sunburst']) {
    const files=exported(book(type),false),key='xl/charts/chart1.xml';
    files[key]=xmlText(files[key]).replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/g,'').replace(/<cx:dataLabel\b[^>]*>[\s\S]*?<\/cx:dataLabel>/g,'');
    const back=chartOf(files);assert.equal(back.seriesFmt[0].pointLabelStyles,undefined);
    expectStyle(back,{kind:'label',s:0,p:1},fonts[2]);
  }
});
