import test from 'node:test';
import assert from 'node:assert/strict';
import { CHART_GALLERY, chartComboDefaults, chartStackValues, resolveChart, renderChartSvg, chartDataLabel, paletteOf } from '../src/chart.js';
import { barSolid3D } from '../src/chart-3d.js';
import { parseXml, descendants } from '../src/xml.js';

const data = { categories: ['가','나','다'], series: [{name:'매출',values:[10,-5,0]},{name:'비용',values:[20,8,5]},{name:'비율',values:[3,7,4]}] };
const base = {type:'column',w:540,h:360,legend:'none'};
const all = svg => descendants(parseXml(svg),'g');
const elements = (svg, name, series) => descendants(parseXml(svg),name).filter(n=>n.attrs['data-s']===String(series));

test('86개 갤러리: 깊이축과 원통·원뿔·피라미드가 독립된 실제 SVG다',()=>{
  const variants=CHART_GALLERY.flatMap(([,entries])=>entries);
  assert.equal(variants.length,86);
  const geometries=new Set();
  for(const shape of ['box','cylinder','cone','pyramid']){
    const svg=renderChartSvg({...base,threeD:true,barShape:shape},data);
    assert.doesNotMatch(svg,/NaN|Infinity|undefined/);geometries.add(svg);
    if(shape!=='box')assert.ok(all(svg).some(g=>g.attrs['data-3d']===shape));
  }
  assert.equal(geometries.size,4);
  for(const shape of ['cylinder','cone','pyramid'])assert.equal(variants.filter(([,p])=>p.barShape===shape).length,7);
  assert.equal(variants.filter(([,p])=>p.grouping==='standard'&&p.type==='column').length,4);
});

test('3D 모양은 방향·음수·영값·회전의 경계에서 값을 변경하지 않는다',()=>{
  for(const shape of ['cylinder','cone','pyramid'])for(const horizontal of [false,true])for(const to of [0,50,100])for(const [dx,dy]of [[0,0],[18,-12],[-18,12]]){
    const svg=barSolid3D({x:20,y:Math.min(50,to),w:30,h:Math.abs(50-to),from:50,to,horizontal,dx,dy,shape,fill:'#336699',attrs:' data-s="7" data-p="2"'});
    assert.doesNotMatch(svg,/NaN|Infinity|undefined/);assert.match(svg,/data-s="7" data-p="2"/);assert.ok(descendants(parseXml(svg),'polygon').length>1);
  }
  const before=structuredClone(data);renderChartSvg({...base,threeD:true,barShape:'cone'},data);assert.deepEqual(data,before);
});

test('깊이축 막대는 계열 순서에 따른 깊이 배치와 선택 번호를 유지한다',()=>{
  const clustered=renderChartSvg({...base,threeD:true},data),deep=renderChartSvg({...base,threeD:true,grouping:'standard'},data);
  assert.notEqual(deep,clustered);
  for(let i=0;i<3;i++)assert.match(deep,new RegExp(`data-s="${i}" data-p="0" data-depth="${i}"`));
  const points=[0,1,2].map(i=>elements(deep,'polygon',i)[0].attrs.points);assert.equal(new Set(points).size,3);
});

test('콤보 프리셋은 타입·축·누적을 실제 계열에 적용한다',()=>{
  const rows=[['분기','온라인','오프라인','목표'],['1분기',10,20,80],['2분기',15,30,90],['3분기',20,25,70],['4분기',15,10,65]];
  for(const [layout,types,grouping]of [['areaColumn',['area','area','column'],['stacked','stacked','clustered']],['stackedColumnLine',['column','column','line'],['stacked','stacked','clustered']]]){
    const chart={...base,type:'combo',comboLayout:layout,comboAxis:'secondary'},resolved=resolveChart(chart,{range:()=>rows});
    assert.deepEqual(resolved.series.map(s=>s.type),types);assert.deepEqual(resolved.series.map(s=>s.grouping),grouping);
    assert.equal(resolved.series.at(-1).axis,layout==='areaColumn'?0:1);
    assert.deepEqual(chartComboDefaults(chart,0,3),{type:types[0],axis:0,grouping:'stacked'});
    const svg=renderChartSvg(chart,resolved);assert.doesNotMatch(svg,/NaN|Infinity/);
    if(layout==='areaColumn')assert.ok(svg.indexOf('fill-opacity="0.9" data-s="0"')<svg.indexOf('data-s="2" data-p="0"'),'면적을 막대보다 먼저 그려 막대를 가리지 않음');
  }
});

test('같은 축·종류라도 계열별 누적과 묶음을 섞어 독립 계산한다',()=>{
  const series=[{type:'column',grouping:'stacked',values:[10,-2]},{type:'column',grouping:'stacked',values:[30,-4]},{type:'column',grouping:'clustered',values:[100,40]},{type:'column',grouping:'percentStacked',values:[2,1]},{type:'column',grouping:'percentStacked',values:[6,3]}];
  const result=chartStackValues(series,'clustered',2);
  assert.deepEqual(result[1]._upper,[40,-6]);assert.deepEqual(result[2]._upper,[100,40]);assert.equal(result[2]._stacked,undefined);assert.deepEqual(result[3]._upper,[.25,.25]);assert.deepEqual(result[4]._upper,[1,1]);
  const svg=renderChartSvg(base,{categories:['가','나'],series}),rects=[0,1,2,3,4].map(i=>elements(svg,'rect',i)[0].attrs);
  assert.equal(rects[0].x,rects[1].x);assert.equal(rects[3].x,rects[4].x);assert.notEqual(rects[0].x,rects[2].x);assert.notEqual(rects[2].x,rects[3].x);
});

test('x/y/y2 축 그룹은 레이블과 제목만 묶고 숨긴 축은 노출하지 않는다',()=>{
  const ch={...base,type:'combo',seriesFmt:[{axis:0},{axis:1}],axes:{x:{title:'분기'},y:{title:'금액'},y2:{title:'비율'}}};
  const d={categories:['A','B'],series:[{name:'주',type:'column',axis:0,values:[10,20]},{name:'보조',type:'line',axis:1,values:[1,2]}]};
  const svg=renderChartSvg(ch,d);
  for(const axis of ['x','y','y2']){const group=all(svg).filter(g=>g.attrs['data-el']==='axis-'+axis);assert.equal(group.length,1);assert.ok(descendants(group[0],'text').length>0);assert.equal(descendants(group[0],'line').length,0);}
  const hidden=renderChartSvg({...ch,axes:{x:{hide:true},y:{hide:true},y2:{hide:true}}},d);assert.ok(!all(hidden).some(g=>String(g.attrs['data-el']).startsWith('axis-')));
});

test('분산·거품 x축 최소/최대/단위/반전/표시형식을 실제 좌표에 반영한다',()=>{
  const d={categories:['A','B'],series:[{name:'값',values:[10,20],x:[2,8],size:[1,2]}]};
  for(const type of ['scatter','bubble']){
    const ch={...base,type,axes:{x:{min:0,max:10,major:5,numFmt:'0.0',title:'측정 X'},y:{min:0,max:30}}},normal=renderChartSvg(ch,d),reverse=renderChartSvg({...ch,axes:{...ch.axes,x:{...ch.axes.x,reverse:true}}},d);
    const groups=all(normal),xt=descendants(groups.find(g=>g.attrs['data-el']==='axis-x'),'text').map(n=>n.text);assert.deepEqual(xt,['0.0','5.0','10.0','측정 X']);
    const a=elements(normal,'circle',0),b=elements(reverse,'circle',0);assert.ok(+a[0].attrs.cx<+a[1].attrs.cx);assert.ok(+b[0].attrs.cx>+b[1].attrs.cx);
    assert.match(normal,/clipPath/);assert.doesNotMatch(normal,/NaN|Infinity/);
  }
});

test('범주명·계열명·값·백분율 레이블과 가져온 팔레트를 실제 렌더한다',()=>{
  const ch={...base,labels:false},s={name:'매출',values:[10,30],catName:true,serName:true,labels:true,pct:true};
  assert.equal(chartDataLabel(ch,s,['1월','2월'],0,10),'매출, 1월, 10, 25%');
  const svg=renderChartSvg(ch,{categories:['1월','2월'],series:[s]});assert.match(svg,/매출, 1월, 10, 25%/);
  assert.deepEqual(paletteOf({palette:['#112233','invalid','#AbCdEf']}),['#112233','#AbCdEf']);
  assert.match(renderChartSvg({...base,palette:['#123456']},data),/fill="#123456"/);
});

test('범주 역순은 막대·선·다중수준 레이블과 원본 선택번호를 함께 유지한다',()=>{
  const d={categories:['A','B','C'],catLevels:[[{text:'앞',start:0,end:1},{text:'뒤',start:2,end:2}]],series:[{name:'값',values:[10,20,30]}]};
  for(const type of ['column','bar','line']){
    const normal=renderChartSvg({...base,type},d),reverse=renderChartSvg({...base,type,axes:{x:{reverse:true}}},d);
    const tag=type==='line'?'circle':'rect',a=elements(normal,tag,0),b=elements(reverse,tag,0),key=type==='bar'?'y':type==='line'?'cx':'x';
    assert.ok(+a[0].attrs[key]<+a[2].attrs[key]);assert.ok(+b[0].attrs[key]>+b[2].attrs[key]);assert.deepEqual(b.map(n=>n.attrs['data-p']),['0','1','2']);
    const labels=descendants(all(reverse).find(g=>g.attrs['data-el']==='axis-x'),'text');assert.ok(labels.some(n=>n.text==='A'));
  }
});

test('원형은 값 없이도 계열/범주/백분율을 선택해 표시하고 글꼴 서식을 적용한다',()=>{
  const d={categories:['A','B'],series:[{name:'매출',values:[1,3],labels:false,serName:true,catName:true,pct:true,labelPos:'center',labelColor:'#123456',labelSize:15,labelBold:false}]};
  const svg=renderChartSvg({...base,type:'pie'}, d);
  const labels=descendants(parseXml(svg),'text').filter(n=>n.attrs['data-el']==='label');
  assert.ok(labels.some(n=>n.text==='매출, A, 25%'));assert.ok(labels.every(n=>n.attrs.fill==='#123456'&&n.attrs['font-size']==='20'&&n.attrs['font-weight']==='400'));
  d.series[0].labelPos='outEnd';
  const outside=descendants(parseXml(renderChartSvg({...base,type:'pie'},d)),'text').filter(n=>n.attrs['data-el']==='label');
  assert.deepEqual(outside.map(n=>n.text),['매출','A','25%','매출','B','75%']);assert.ok(outside.every(n=>n.attrs.fill==='#123456'&&n.attrs['font-size']==='20'));
});
