import { test } from 'node:test';
import { strict as assert } from 'node:assert';
import { normalizeChartAreaFormat, chartAreaFormat, chartAreaFormatPatch, chartAreaSvg } from '../src/chart-area-format.js';
import { chartAreaFormatXml, readChartAreaFormat } from '../src/chart-area-drawingml.js';
import { renderChartSvg } from '../src/chart.js';
import { parseXml, child, descendants } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';

const round = (f, opts={}) => readChartAreaFormat(parseXml(chartAreaFormatXml(f,opts)),opts);
const near=(a,b,eps=1e-4)=>assert.ok(Math.abs(a-b)<eps,`${a} != ${b}`);
const data={categories:['A','B'],series:[{name:'값',values:[2,4]}]};

test('legacy chart/plot backgrounds are preserved without changing input models',()=>{
  const ch={fill:'#abc',border:'#123456',plotFill:'transparent'},before=JSON.stringify(ch);
  assert.equal(chartAreaFormat(ch).fill,'#aabbcc');assert.equal(chartAreaFormat(ch).lineMode,'solid');
  assert.equal(chartAreaFormat(ch,'plot').fillMode,'none');assert.equal(JSON.stringify(ch),before);
  assert.equal(chartAreaFormat({}).fillMode,'auto');assert.equal(chartAreaFormat({},'plot').lineMode,'auto');
});
test('unsafe strings, NaN, unsupported modes and unbounded geometry are rejected',()=>{
  const f=normalizeChartAreaFormat({fillMode:'<script>',fill:'" onload=evil',strokeWidth:Infinity,picture:{src:'javascript:alert(1)'},soft:900,grad:{stops:[[2,'red'],[-1,'#fff'],[NaN,'#abc']]}});
  assert.equal(f.fillMode,'auto');assert.equal(f.strokeWidth,1);assert.equal(f.picture.src,'');assert.equal(f.soft,50);assert.deepEqual(f.grad.stops.map(s=>s[0]),[0,1]);
  const svg=chartAreaSvg(f,{x:0,y:0,w:100,h:100},{id:'bad"><'});assert.doesNotMatch(svg,/onload|javascript:|NaN|Infinity/);
});
test('area patch targets only selected area and deeply isolates stop arrays',()=>{
  const ch={fill:'#123456',plotFill:'#abcdef'},p=chartAreaFormatPatch(ch,'plot',{fillMode:'gradient',grad:{stops:[[0,'#000000'],[1,'#ffffff']]}});
  assert.deepEqual(Object.keys(p),['plotAreaFormat']);p.plotAreaFormat.grad.stops[0][1]='#777777';assert.equal(ch.plotFill,'#abcdef');assert.equal(ch.fill,'#123456');
});
test('solid fill and line alpha/width/compound/cap/join round trip in native XML',()=>{
  const f={fillMode:'solid',fill:'#123456',fillOpacity:.3,lineMode:'solid',stroke:'#abcdef',strokeOpacity:.6,strokeWidth:8,compound:'thickThin',dash:'lgDashDotDot',lineCap:'sq',lineJoin:'miter'};
  const b=round(f);for(const k of Object.keys(f))assert.equal(b[k],f[k]);
});
test('native noFill remains explicit for both fill and border; auto omits them',()=>{
  const f=round({fillMode:'none',lineMode:'none'});assert.equal(f.fillMode,'none');assert.equal(f.lineMode,'none');
  const auto=chartAreaFormatXml({fillMode:'auto',lineMode:'auto'});assert.equal(auto,'<c:spPr></c:spPr>');assert.equal(readChartAreaFormat(parseXml(auto)),undefined);
});
test('linear and radial gradients retain positions and multiplied alpha',()=>{
  for(const type of ['linear','radial']){const f={fillMode:'gradient',fillOpacity:.5,grad:{type,ang:35,stops:[[0,'#000000',.4],[.6,'#336699',.7],[1,'#ffffff',1]]}};const b=round(f);assert.equal(b.grad.type,type);assert.deepEqual(b.grad.stops,[[0,'#000000',.2],[.6,'#336699',.35],[1,'#ffffff',.5]]);if(type==='linear')assert.equal(b.grad.ang,35);}
});
test('gradient outline and pattern background produce distinct native XML',()=>{
  const f={fillMode:'pattern',pattern:{preset:'diagCross',fg:'#112233',bg:'#ddeeff'},lineMode:'gradient',strokeGrad:{ang:180,stops:[[0,'#ff0000'],[1,'#0000ff']]}};
  const xml=chartAreaFormatXml(f),b=round(f);assert.match(xml,/<a:pattFill/);assert.equal(b.pattern.preset,'diagCross');assert.equal(b.strokeGrad.ang,180);assert.equal(b.lineMode,'gradient');
});
test('image fill uses explicit media callback, with tile scale and transparency',()=>{
  const src='data:image/png;base64,AAAA',seen=[];
  const f={fillMode:'picture',fillOpacity:.5,picture:{src,mode:'tile',scale:1.5,opacity:.8}};
  const xml=chartAreaFormatXml(f,{imageRel:s=>{seen.push(s);return 'rId9';}});assert.deepEqual(seen,[src]);assert.match(xml,/r:embed="rId9"/);
  const b=readChartAreaFormat(parseXml(xml),{imageSource:(id,external)=>{assert.equal(id,'rId9');assert.equal(external,false);return src;}});assert.equal(b.picture.src,src);assert.equal(b.picture.scale,1.5);assert.equal(b.picture.opacity,.4);assert.equal(b.picture.mode,'tile');
  assert.doesNotMatch(chartAreaFormatXml(f),/<a:blip/);
});
test('theme color callback receives its color container and alpha transforms apply',()=>{
  const xml='<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:alpha val="80000"/><a:alphaMod val="50000"/></a:schemeClr></a:solidFill></c:spPr>';
  const f=readChartAreaFormat(parseXml(xml),{readColor:node=>{assert.equal(node.name,'solidFill');return '#123456';}});assert.equal(f.fill,'#123456');assert.equal(f.fillOpacity,.4);
});
test('outer shadow angle/distance/size glow and soft edges use standard units',()=>{
  const f={shadow:{dx:-7,dy:4,blur:9,color:'#234567',opacity:.45,scale:1.15},glow:{color:'#ff6600',size:7,opacity:.7},soft:2};
  const xml=chartAreaFormatXml(f),b=round(f);near(b.shadow.dx,-7);near(b.shadow.dy,4);assert.equal(b.shadow.scale,1.15);assert.equal(b.shadow.blur,9);assert.deepEqual(b.glow,f.glow);assert.equal(b.soft,2);
  assert.ok(xml.indexOf('<a:glow')<xml.indexOf('<a:outerShdw'));assert.ok(xml.indexOf('<a:outerShdw')<xml.indexOf('<a:softEdge'));
});
test('bevel top/bottom depth contour material and lighting use native sp3d',()=>{
  const d={bevelTop:{type:'angle',w:8,h:4},bevelBottom:{type:'circle',w:5,h:3},depth:9,contourWidth:2,contourColor:'#008899',material:'metal',lightRig:'balanced',lightAngle:120};
  const b=round({threeD:d});assert.deepEqual(b.threeD,d);assert.match(chartAreaFormatXml({threeD:d}),/<a:camera prst="orthographicFront"/);
});
test('transparent compound outlines are separate strokes instead of opaque gap masks',()=>{
  const xml=parseXml(chartAreaSvg({fillMode:'none',lineMode:'solid',strokeWidth:9,compound:'dbl'},{x:0,y:0,w:100,h:80}));
  const lines=descendants(xml,'rect').filter(n=>n.attrs.stroke);assert.equal(lines.length,2);assert.deepEqual(lines.map(n=>+n.attrs['stroke-width']),[3,3]);assert.deepEqual(lines.map(n=>+n.attrs.x),[1.5,7.5]);assert.ok(lines.every(n=>n.attrs.fill==='none'));
});
test('picture stretch, tile and scale change actual SVG placement',()=>{
  const src='data:image/png;base64,AAAA',f={fillMode:'picture',picture:{src,mode:'stretch'}};
  const one=chartAreaSvg(f,{x:10,y:20,w:200,h:80}),two=chartAreaSvg({...f,picture:{src,mode:'tile',scale:2}},{x:10,y:20,w:200,h:80});
  assert.match(one,/preserveAspectRatio="none"/);assert.match(two,/<pattern[^>]+width="256"/);assert.match(one,/x="10" y="20" width="200" height="80"/);
});
test('area-only effects do not filter chart data or title',()=>{
  const svg=renderChartSvg({type:'column',w:500,h:300,title:'검증',chartAreaFormat:{fillMode:'solid',shadow:{dx:3,dy:5,scale:1.1}},plotAreaFormat:{fillMode:'pattern'}},data),xml=parseXml(svg);
  assert.equal(descendants(xml,'g').filter(n=>n.attrs['data-area-format']).length,2);
  assert.ok(descendants(xml,'g').some(n=>n.attrs['data-el']==='title'));
  assert.ok(descendants(xml,'rect').some(n=>n.attrs['data-s']!==undefined));
  assert.doesNotMatch(svg,/NaN|Infinity/);
});
test('plot format uses cartesian data area and works for empty/special charts',()=>{
  const chart={type:'column',w:500,h:300,plotAreaFormat:{fillMode:'solid',fill:'#ff0000'},legend:'none'};
  const parsed=parseXml(renderChartSvg(chart,data)),plot=descendants(parsed,'g').find(n=>n.attrs['data-el']==='plot');assert.ok(+child(plot,'rect').attrs.x>10);
  for(const type of ['pie','treemap','waterfall'])assert.match(renderChartSvg({...chart,type},data),/data-area-format="1"/);
  assert.match(renderChartSvg(chart,{categories:[],series:[]}),/data-el="plot"/);
});
test('3D options change visible geometry or shade and remain bounded',()=>{
  const rect={x:10,y:10,w:200,h:100},base={fillMode:'solid',threeD:{depth:10,bevelTop:{type:'angle',w:8,h:4}}};
  const a=chartAreaSvg(base,rect);
  for(const patch of [{depth:20},{bevelTop:{type:'circle',w:8,h:4}},{bevelBottom:{type:'angle',w:8,h:4}},{material:'metal'},{lightRig:'soft'},{lightAngle:120},{contourWidth:2}])assert.notEqual(chartAreaSvg({...base,threeD:{...base.threeD,...patch}},rect),a);
  assert.doesNotMatch(a,/NaN|Infinity/);
});

test('XLSX native chart/plot styles survive without WIXEL extensions',()=>{
  const wb=new Workbook();wb.transact(()=>[['항목','값'],['가',30],['나',70]].forEach((r,y)=>r.forEach((v,x)=>wb.setInput(0,y,x,String(v)))));
  wb.sheets[0].charts=[{id:'area',type:'column',w:600,h:360,x:0,y:0,range:{r1:0,c1:0,r2:2,c2:1},chartAreaFormat:{fillMode:'pattern',pattern:{preset:'cross',fg:'#112233',bg:'#ddeeff'},lineMode:'solid',stroke:'#445566',strokeWidth:4},plotAreaFormat:{fillMode:'gradient',grad:{ang:40,stops:[[0,'#123456',.5],[1,'#abcdef',1]]},lineMode:'gradient',strokeGrad:{ang:180,stops:[[0,'#ff0000'],[1,'#0000ff']]},shadow:{dx:3,dy:4,blur:5,scale:1.1,color:'#000000',opacity:.4},glow:{size:3,color:'#ffaa00',opacity:.5},soft:1,threeD:{bevelTop:{type:'angle',w:6,h:3},depth:4}}}];
  const files=unzip(writeXlsx(wb));for(const k of Object.keys(files))if(k.endsWith('.xml'))files[k]=textOf(files[k]).replace(/<(?:\w+:)?extLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?extLst>/g,'');
  const back=readXlsx(zip(files)).data.sheets[0].charts[0];assert.equal(back.chartAreaFormat.fillMode,'pattern');assert.equal(back.chartAreaFormat.pattern.fg,'#112233');assert.equal(back.chartAreaFormat.strokeWidth,4);assert.equal(back.plotAreaFormat.grad.ang,40);assert.equal(back.plotAreaFormat.grad.stops[0][2],.5);assert.equal(back.plotAreaFormat.lineMode,'gradient');assert.equal(back.plotAreaFormat.shadow.scale,1.1);near(back.plotAreaFormat.shadow.dy,4);assert.equal(back.plotAreaFormat.threeD.depth,4);
});
