import test from 'node:test';
import assert from 'node:assert/strict';
import {renderChartSvg,resolveChart,chartModelData} from '../src/chart.js';
import {parseXml,descendants} from '../src/xml.js';
const source=()=>({categories:['A','B','C'],series:[{name:'매출',_fi:0,values:[10,20,30]},{name:'방문',_fi:1,values:[30,10,20]}]});
const nodes=xml=>['path','rect','circle'].flatMap(tag=>descendants(xml,tag));
const render=(extra={},series={})=>parseXml(renderChartSvg({type:'radar',w:600,h:400,...extra},{...source(),series:[{...source().series[0],...series}]}));
test('방사형 선의 굵기와 파선이 실제 경로에 적용된다',()=>{const path=nodes(render({},{lineWidth:8,dash:'dash'})).find(n=>n.attrs['data-s']==='0'&&n.attrs['data-p']===undefined);assert.equal(path.attrs['stroke-width'],'8');assert.ok(path.attrs['stroke-dasharray']);});
test('방사형 각 표식은 크기·개별색·원본 인덱스를 유지한다',()=>{for(const [marker,tag]of [['circle','circle'],['square','rect'],['diamond','path'],['triangle','path']]){const points=nodes(render({radarStyle:'marker'},{marker,markerSize:12,_pi:[2,4,6],pointColors:{4:'#123456'}})).filter(n=>n.attrs['data-p']!==undefined);assert.equal(points.length,3,marker);assert.deepEqual(points.map(n=>n.name),[tag,tag,tag],marker);assert.deepEqual(points.map(n=>n.attrs['data-p']),['2','4','6']);assert.equal(points[1].attrs.fill,'#123456');}});
test('방사형 자동 표식과 명시적인 표식 없음·표식 선택을 구분한다',()=>{const count=(ch,s)=>nodes(render(ch,s)).filter(n=>n.attrs['data-p']!==undefined).length;assert.equal(count({radarStyle:'marker'},{}),3);assert.equal(count({},{}),0);assert.equal(count({radarStyle:'filled'},{}),0);assert.equal(count({radarStyle:'marker'},{marker:'none'}),0);assert.equal(count({radarStyle:'marker'},{marker:false}),0);assert.equal(count({radarStyle:'filled'},{marker:'square'}),3);});
test('스냅샷 필터는 계열·범주를 제외하고 원본 인덱스와 서식을 보존한다',()=>{const snapshotData=source();snapshotData.series[0]._fi=2;snapshotData.series[1]._fi=4;const before=structuredClone(snapshotData),seriesFmt=[];seriesFmt[4]={color:'#123456'};const result=resolveChart({snapshotData,hiddenSeries:[2],hiddenCats:[1],seriesFmt},{});assert.equal(result.series.length,1);assert.equal(result.series[0]._fi,4);assert.equal(result.series[0].color,'#123456');assert.deepEqual(result.series[0].values,[30,20]);assert.deepEqual(result.series[0]._pi,[0,2]);assert.deepEqual(result.categories,['A','C']);assert.deepEqual(snapshotData,before);});
test('인덱스가 없는 스냅샷도 숨김/복원 시 일관된 원본 인덱스를 갖는다',()=>{const snapshotData=source();snapshotData.series.forEach(s=>delete s._fi);const filtered=resolveChart({snapshotData,hiddenSeries:[0]},{});assert.equal(filtered.series[0]._fi,1);const original=resolveChart({snapshotData},{});assert.deepEqual(original.series.map(s=>s._fi),[0,1]);assert.equal(original.series.length,2);});

import {chartPalettePatch} from '../src/chart-edit.js';
import {Workbook} from '../src/workbook.js';
import {writeXlsx,readXlsx} from '../src/xlsx.js';
import {unzip,zip,textOf} from '../src/zip.js';
const piePoints=(type,fmt,values=[10,20,30],chart={})=>nodes(parseXml(renderChartSvg({type,w:500,h:320,...chart}, {categories:['A','B','C'],series:[{name:'매출',_fi:0,values,...fmt}]}))).filter(n=>n.attrs['data-p']!==undefined);
for(const type of ['pie','doughnut']){
 test(type+': 계열 테두리가 모든 조각·전체 원에 적용되고 기본 경계는 유지된다',()=>{
  for(const values of [[10,20,30],[10,0,0]])for(const [fmt,chart,expected]of [[{outline:'#123456'},{},'#123456'],[{outline:'#123456'},{fill:'#cccccc'},'#123456'],[{}, {}, '#fff']]){const pts=piePoints(type,fmt,values,chart);assert.equal(pts.length,values.filter(Boolean).length);assert.ok(pts.every(n=>n.attrs.stroke===expected));}
 });
 test(type+': 팔레트 변경은 계열 테두리 직접색도 해제해 기본값으로 복원한다',()=>{
  const chart={type,w:500,h:320,seriesFmt:[{outline:'#123456'}],snapshotData:source()},next={...chart,...chartPalettePatch(chart,'modern')};assert.equal(next.seriesFmt[0].outline,undefined);const points=nodes(parseXml(renderChartSvg(next,resolveChart(next,{})))).filter(n=>n.attrs['data-p']!==undefined);assert.ok(points.length);assert.ok(points.every(n=>n.attrs.stroke==='#fff'));
 });
 test(type+': 테두리 색은 WIXEL 확장 없이 표준 XLSX에서도 왕복된다',()=>{
  const wb=new Workbook();[['항목','값'],['A',10],['B',20],['C',30]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));wb.sheets[0].charts=[{id:'outline',type,x:10,y:20,w:500,h:320,range:{r1:0,c1:0,r2:3,c2:1},seriesFmt:[{outline:'#123456'}]}];
  const files=unzip(writeXlsx(wb));for(const path of Object.keys(files).filter(p=>/^xl\/(charts|drawings)\/[^/]+\.xml$/.test(p))){const text=typeof files[path]==='string'?files[path]:textOf(files[path]);files[path]=text.replace(/<(?:c|cx|a):extLst>[\s\S]*?<\/(?:c|cx|a):extLst>/g,'');}
  const back=new Workbook(readXlsx(zip(files)).data),chart=back.sheets[0].charts[0];assert.equal(chart.seriesFmt[0].outline.toLowerCase(),'#123456');const pts=nodes(parseXml(renderChartSvg(chart,chartModelData(back,0,chart)))).filter(n=>n.attrs['data-p']!==undefined);assert.equal(pts.length,3);assert.ok(pts.every(n=>n.attrs.stroke.toLowerCase()==='#123456'));
 });
}

test('방사형 선 굵기·파선·표식은 확장 정보 없는 표준 XLSX에서 복원된다',()=>{
 const wb=new Workbook();[['항목','값'],['A',10],['B',20],['C',30]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));wb.sheets[0].charts=[{id:'radar-format',type:'radar',radarStyle:'marker',x:10,y:20,w:500,h:320,range:{r1:0,c1:0,r2:3,c2:1},seriesFmt:[{lineWidth:8,dash:'dash',marker:'square',markerSize:12}]}];
 const files=unzip(writeXlsx(wb));for(const path of Object.keys(files).filter(p=>/^xl\/(charts|drawings)\/[^/]+\.xml$/.test(p))){const text=typeof files[path]==='string'?files[path]:textOf(files[path]);files[path]=text.replace(/<(?:c|cx|a):extLst>[\s\S]*?<\/(?:c|cx|a):extLst>/g,'');}
 const back=new Workbook(readXlsx(zip(files)).data),chart=back.sheets[0].charts[0],fmt=chart.seriesFmt[0];assert.equal(fmt.lineWidth,8);assert.equal(fmt.dash,'dash');assert.equal(fmt.marker,'square');assert.equal(fmt.markerSize,12);const ns=nodes(parseXml(renderChartSvg(chart,chartModelData(back,0,chart)))),line=ns.find(n=>n.attrs['data-s']==='0'&&n.attrs['data-p']===undefined);assert.equal(line.attrs['stroke-width'],'8');assert.ok(line.attrs['stroke-dasharray']);const markers=ns.filter(n=>n.attrs['data-p']!==undefined);assert.equal(markers.length,3);assert.ok(markers.every(n=>n.name==='rect'&&n.attrs.width==='16'));
});
