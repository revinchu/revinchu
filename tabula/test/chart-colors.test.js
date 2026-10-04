import test from 'node:test';
import assert from 'node:assert/strict';
import { chartPalettePatch, chartSeriesColorPatch, chartStylePatch, chartPointColorPatch } from '../src/chart-edit.js';
import { paletteOf, resolveChart, renderChartSvg, chartModelData } from '../src/chart.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const colorKeys=['color','grad','colors','pointColors','hierarchyColors','markerColor','trendColor','outline'];
const overrides=()=>({color:'#ff0000',grad:{ang:90,stops:[[0,'#ff0000'],[1,'#990000']]},colors:['#00aa00','#0000aa','#aaaa00'],pointColors:{0:'#00aa00',1:'#0000aa',2:'#aaaa00'},hierarchyColors:{'["A"]':'#ee00ee'},markerColor:'#abcdef',trendColor:'#fedcba',outline:'#123456'});
const chart=(type='column')=>({id:'color-test',type,x:10,y:20,w:500,h:340,legend:'none',series:[{name:{text:'매출'},catCache:['A','B','C'],cache:[10,20,30]}],seriesFmt:[{...overrides(),lineWidth:3,pointExplosion:{1:30},labels:false,trend:'linear',trendForward:2}],axes:{y:{min:0,max:80}},hiddenCats:[]});
const apply=(c,p)=>({...c,...p});
const draw=c=>parseXml(renderChartSvg(c,resolveChart(c,{})));
const shapes=root=>['path','rect','circle','ellipse','polygon'].flatMap(k=>descendants(root,k));
const points=root=>shapes(root).filter(n=>n.attrs['data-p']!==undefined);
const fills=c=>points(draw(c)).map(n=>n.attrs.fill).filter(Boolean);
const lower=a=>a.map(s=>s.toLowerCase());
const native=wb=>{
  const files=unzip(writeXlsx(wb));
  for(const path of Object.keys(files).filter(p=>/^xl\/charts\/chart\d+\.xml$/.test(p))){const xml=typeof files[path]==='string'?files[path]:textOf(files[path]);files[path]=xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');}
  return new Workbook(readXlsx(zip(files)).data);
};
const book=(type='column')=>{
  const wb=new Workbook();
  [['항목','매출'],['A',10],['B',20],['C',30]].forEach((r,ri)=>r.forEach((v,ci)=>wb.setInput(0,ri,ci,String(v))));
  const c=chart(type);delete c.series;c.range={r1:0,c1:0,r2:3,c2:1};wb.sheets[0].charts=[c];return wb;
};
const renderedBook=wb=>{const c=wb.sheets[0].charts[0];return parseXml(renderChartSvg(c,chartModelData(wb,0,c)));};

test('전체 팔레트는 데이터색 오버라이드만 교체하고 원본 참조·서식·옵션은 보존한다',()=>{
  const c=chart();Object.assign(c,{upColor:'#010101',downColor:'#020202',totalColor:'#030303',mapLowColor:'#040404',mapMidColor:'#050505',mapHighColor:'#060606',otherColor:'#070707'});
  const before=structuredClone(c),next=apply(c,chartPalettePatch(c,'vivid'));
  assert.deepEqual(c,before);assert.deepEqual(paletteOf(next),paletteOf({palette:'vivid'}));
  for(const k of colorKeys)assert.equal(next.seriesFmt[0][k],undefined,k);
  for(const k of ['upColor','downColor','totalColor','mapLowColor','mapMidColor','mapHighColor','otherColor'])assert.equal(next[k],undefined,k);
  for(const k of ['lineWidth','pointExplosion','labels','trend','trendForward'])assert.deepEqual(next.seriesFmt[0][k],c.seriesFmt[0][k]);
  assert.deepEqual(next.series,c.series);assert.deepEqual(next.axes,c.axes);assert.deepEqual(next.hiddenCats,c.hiddenCats);
  assert.equal(next.x,c.x);assert.equal(next.y,c.y);assert.equal(next.w,c.w);assert.equal(next.h,c.h);
});

test('표준 팔레트로 복원하면 실제 데이터가 Office 색으로 다시 그려진다',()=>{
  const c=apply(chart(),chartPalettePatch(chart(),'vivid')),next=apply(c,chartPalettePatch(c,'office'));
  assert.deepEqual(paletteOf(next),paletteOf({}));
  assert.ok(fills(next).every(c=>c===paletteOf({})[0]));
});

test('가져온 그라데이션과 포인트별색이 있는 막대도 선택한 팔레트가 보인다',()=>{
  const c=chart(),before=fills(c),next=apply(c,chartPalettePatch(c,'vivid'));
  assert.notDeepEqual(fills(next),before);
  assert.deepEqual(fills(next),Array(3).fill(paletteOf(next)[0]));
  assert.equal(descendants(draw(next),'linearGradient').length,0);
});

for(const type of ['pie','doughnut','pieOfPie','barOfPie'])test(`${type}: 전체 색 구성은 가져온 조각별 색보다 우선하여 재적용된다`,()=>{
  const c=chart(type),next=apply(c,chartPalettePatch(c,'vivid')),pal=paletteOf(next);
  const ns=points(draw(next));assert.ok(ns.length>=3);
  for(const n of ns)assert.equal(n.attrs.fill,pal[Number(n.attrs['data-p'])%pal.length]);
});

test('숨긴 계열·항목의 원본 번호를 지키며 새 팔레트 색을 고른다',()=>{
  const c=chart();c.series.push({name:{text:'이익'},catCache:['A','B','C'],cache:[2,4,6]});c.seriesFmt.push({...overrides()});c.hiddenSeries=[0];c.hiddenCats=[0];
  const next=apply(c,chartPalettePatch(c,'vivid')),ns=points(draw(next));
  assert.deepEqual([...new Set(ns.map(n=>n.attrs['data-s']))],['1']);
  assert.deepEqual(ns.map(n=>n.attrs['data-p']),['1','2']);
  assert.ok(ns.every(n=>n.attrs.fill===paletteOf(next)[1]));
  const pie=chart('pie');pie.hiddenCats=[0];const colored=apply(pie,chartPalettePatch(pie,'vivid'));
  assert.deepEqual(points(draw(colored)).map(n=>[n.attrs['data-p'],n.attrs.fill]),[['1',paletteOf(colored)[1]],['2',paletteOf(colored)[2]]]);
});

test('계열 단색은 해당 계열의 오래된 채우기만 치우고 다른 계열은 유지한다',()=>{
  const c=chart();c.series.push({name:{text:'이익'},catCache:['A','B','C'],cache:[2,4,6]});c.seriesFmt.push({...overrides(),lineWidth:6});
  const before=structuredClone(c),next=apply(c,chartSeriesColorPatch(c,0,'#654321'));
  assert.deepEqual(c,before);assert.deepEqual(next.seriesFmt[1],c.seriesFmt[1]);
  assert.equal(next.seriesFmt[0].color,'#654321');for(const k of colorKeys.filter(k=>k!=='color'))assert.equal(next.seriesFmt[0][k],undefined,k);
  assert.equal(next.seriesFmt[0].lineWidth,3);assert.deepEqual(next.seriesFmt[0].pointExplosion,{1:30});
  const ns=points(draw(next));assert.ok(ns.filter(n=>n.attrs['data-s']==='0').every(n=>n.attrs.fill==='#654321'));
  assert.deepEqual(ns.filter(n=>n.attrs['data-s']==='1').map(n=>n.attrs.fill),['#00aa00','#0000aa','#aaaa00']);
});

test('계열 색 자동 선택은 원래 그라데이션이 재등장하지 않고 팔레트로 돌아간다',()=>{
  const c=chart(),next=apply(c,chartSeriesColorPatch(c,0,undefined));
  assert.ok(fills(next).every(v=>v===paletteOf(next)[0]));
  assert.equal(descendants(draw(next),'linearGradient').length,0);
});

test('스냅샷 차트의 캐시 색과 새 계열색·전체 팔레트는 같은 우선순위를 쓴다',()=>{
  const c={...chart(),series:undefined,seriesFmt:[{...overrides(),labels:false}],snapshotData:{categories:['A','B','C'],series:[{name:'매출',values:[10,20,30],...overrides()}]}};
  const before=structuredClone(c),palette=apply(c,chartPalettePatch(c,'vivid'));
  assert.ok(fills(palette).every(v=>v===paletteOf(palette)[0]));
  const solid=apply(c,chartSeriesColorPatch(c,0,'#654321'));
  assert.ok(fills(solid).every(v=>v==='#654321'));
  assert.deepEqual(c,before);assert.deepEqual(solid.snapshotData.series[0].values,[10,20,30]);
});

test('스타일의 배경·팔레트는 이전 영역서식보다 우선하고 원본 데이터는 보존된다',()=>{
  const c=chart();Object.assign(c,{fill:'#010101',plotFill:'#020202',titleColor:'#030303',chartAreaFormat:{fillMode:'solid',fill:'#040404'},plotAreaFormat:{fillMode:'solid',fill:'#050505'}});
  const before=structuredClone(c),preset={fill:'#0f172a',plotFill:'#111827',palette:'vivid',titleBold:true};
  const next=apply(c,chartStylePatch(c,preset,9));
  assert.equal(next.chartAreaFormat,undefined);assert.equal(next.plotAreaFormat,undefined);
  assert.equal(next.fill,preset.fill);assert.equal(next.plotFill,preset.plotFill);assert.equal(next.chartStyle,9);
  assert.ok(fills(next).every(v=>v===paletteOf(next)[0]));
  assert.equal(draw(next).children.find(n=>n.name==='rect').attrs.fill,preset.fill);
  assert.deepEqual(c,before);assert.deepEqual(next.series,c.series);
  const backgroundOnly=apply(c,chartStylePatch(c,{fill:'#ffffff'},0));
  assert.deepEqual(backgroundOnly.seriesFmt,c.seriesFmt,'팔레트 없는 배경 스타일은 개별 데이터색 보존');
});

test('폭포·계층형 차트도 예전 사용자 색 때문에 전체 색 변경이 막히지 않는다',()=>{
  const wf=chart('waterfall');wf.upColor='#010101';wf.downColor='#020202';wf.totalColor='#030303';wf.series[0].cache=[10,-3,12];
  const w=apply(wf,chartPalettePatch(wf,'vivid'));assert.ok(fills(w).every(v=>paletteOf(w).includes(v)));
  for(const type of ['treemap','sunburst']){
    const c=chart(type),next=apply(c,chartPalettePatch(c,'vivid'));const ns=shapes(draw(next)).filter(n=>n.attrs['data-node']);
    assert.ok(ns.length);assert.ok(ns.every(n=>paletteOf(next).includes(n.attrs.fill)));
  }
});

test('전체 색 변경은 Undo·Redo 한 번으로 복원되고 셀 수식은 변경하지 않는다',()=>{
  const wb=new Workbook({sheets:[{name:'검증',cells:{'0,0':{raw:'=1+1',cached:2}},charts:[chart()]}]}),before=structuredClone(wb.sheets[0].charts[0]);
  wb.transact(()=>wb.setSheetProp(0,'charts',[apply(before,chartPalettePatch(before,'vivid'))]));
  const after=structuredClone(wb.sheets[0].charts[0]);assert.notDeepEqual(after,before);
  wb.undo();assert.deepEqual(wb.sheets[0].charts[0],before);
  wb.redo();assert.deepEqual(wb.sheets[0].charts[0],after);
  assert.equal(wb.getRaw(0,0,0),'=1+1');
});

test('새 팔레트 막대색은 WIXEL 확장 없이 표준 XLSX로도 왕복된다',()=>{
  const wb=book(),c=wb.sheets[0].charts[0];wb.sheets[0].charts=[apply(c,chartPalettePatch(c,'vivid'))];
  const before=points(renderedBook(wb)).map(n=>n.attrs.fill),back=native(wb);
  assert.deepEqual(lower(points(renderedBook(back)).map(n=>n.attrs.fill)),lower(before));
});

test('원형 계열 단색과 그 후 지정한 한 조각색이 표준 XLSX의 실제 색에 저장된다',()=>{
  for(const type of ['pie','doughnut','pieOfPie','barOfPie']){
    const wb=book(type),c=wb.sheets[0].charts[0];let next=apply(c,chartSeriesColorPatch(c,0,'#654321'));
    next=apply(next,chartPointColorPatch(next,0,1,'#abcdef'));wb.sheets[0].charts=[next];
    const before=points(renderedBook(wb)).map(n=>[n.attrs['data-p'],n.attrs.fill.toLowerCase()]);
    assert.ok(before.some(([p,v])=>p==='1'&&v==='#abcdef'),type);
    assert.ok(before.filter(([p])=>p!=='1').every(([,v])=>v==='#654321'),type);
    const back=native(wb),after=points(renderedBook(back)).map(n=>[n.attrs['data-p'],n.attrs.fill.toLowerCase()]);
    assert.deepEqual(after,before,type);
  }
});


test('요소마다 다른 색을 켠 막대에서도 명시 계열색이 화면과 표준 XLSX에 우선한다',()=>{
  for(const type of ['column','bar']){
    const wb=book(type),c=wb.sheets[0].charts[0];c.varyColors=true;
    let next=apply(c,chartSeriesColorPatch(c,0,'#654321'));wb.sheets[0].charts=[next];
    assert.ok(points(renderedBook(wb)).every(n=>n.attrs.fill==='#654321'),type);
    assert.ok(points(renderedBook(native(wb))).every(n=>n.attrs.fill.toLowerCase()==='#654321'),type);
    next=apply(next,chartPalettePatch(next,'vivid'));wb.sheets[0].charts=[next];
    assert.deepEqual(lower(points(renderedBook(wb)).map(n=>n.attrs.fill)),lower(paletteOf(next).slice(0,3)),type);
    assert.deepEqual(lower(points(renderedBook(native(wb))).map(n=>n.attrs.fill)),lower(paletteOf(next).slice(0,3)),type);
  }
});

test('개별 요소 색 자동복원은 가져온 중복 colors 캐시를 제거하되 이웃 색은 유지한다',()=>{
  const c=chart('pie');delete c.seriesFmt[0].color;const before=structuredClone(c);
  const next=apply(c,chartPointColorPatch(c,0,1,undefined));
  const ns=points(draw(next));
  assert.deepEqual(ns.map(n=>n.attrs.fill),['#00aa00',paletteOf(next)[1],'#aaaa00']);
  assert.deepEqual(c,before);
  assert.deepEqual(next.seriesFmt[0].pointExplosion,{1:30});
});

test('필터된 스냅샷의 원본 계열 번호가 전체 팔레트 변경 후에도 유지된다',()=>{
  const c={...chart(),series:undefined,seriesFmt:[{},{},{},{...overrides()}],snapshotData:{categories:['A','B','C'],series:[{name:'원래 넷째 계열',_fi:3,values:[10,20,30],...overrides()}]}};
  const next=apply(c,chartPalettePatch(c,'vivid'));
  const ns=points(draw(next));assert.ok(ns.length);
  assert.ok(ns.every(n=>n.attrs['data-s']==='3'));
  assert.ok(ns.every(n=>n.attrs.fill===paletteOf(next)[3]));
});
