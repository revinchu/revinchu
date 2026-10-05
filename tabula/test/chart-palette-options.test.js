import test from 'node:test';
import assert from 'node:assert/strict';
import { CHART_PALETTES, paletteOf, renderChartSvg, chartModelData } from '../src/chart.js';
import { CHART_PALETTE_GROUPS, EXTRA_CHART_PALETTES, normalizeChartPalette, chartPaletteOptions } from '../src/chart-palette-options.js';
import { chartPalettePatch } from '../src/chart-edit.js';
import { withThemeColors, DEFAULT_THEME } from '../src/stylepresets.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const originalIds = ['office','colorful2','colorful3','colorful4', ...Array.from({length:13},(_,i)=>`mono${i+1}`), 'modern','pastel','slate','vivid'];
function book(type, palette) {
  const wb = new Workbook(), rows = type === 'column' ? [['항목','가','나','다','라','마','바'],['A',10,20,30,40,50,60],['B',15,25,35,45,55,65]] : [['항목','금액'],['A',10],['B',type==='waterfall'?-3:3],['합계',7]];
  wb.transact(()=>rows.forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))));
  wb.sheets[0].charts = [{id:'palette',type,palette,range:{r1:0,c1:0,r2:rows.length-1,c2:rows[0].length-1},x:0,y:0,w:520,h:340,legend:'b'}];
  if(type==='column')wb.sheets[0].charts[0].series=Array.from({length:6},(_,i)=>({name:{text:rows[0][i+1]},cat:{sheet:wb.sheets[0].name,r1:1,c1:0,r2:2,c2:0},val:{sheet:wb.sheets[0].name,r1:1,c1:i+1,r2:2,c2:i+1}}));
  return wb;
}
const xmlText = value => typeof value==='string' ? value : textOf(value);
const svg = wb => {const chart=wb.sheets[0].charts[0];return renderChartSvg(chart,chartModelData(wb,0,chart));};
const pointFills = wb => descendants(parseXml(svg(wb)),'rect').filter(n=>n.attrs['data-p']!==undefined).map(n=>n.attrs.fill.toUpperCase());
function standardReopen(wb) {
  const files = unzip(writeXlsx(wb));
  for(const path of Object.keys(files).filter(p=>/^xl\/(charts|drawings)\/[^/]+\.xml$/.test(p))) {
    files[path] = xmlText(files[path]).replace(/<(?:c|cx|a):extLst>[\s\S]*?<\/(?:c|cx|a):extLst>/g,'');
  }
  return {wb:new Workbook(readXlsx(zip(files)).data),files};
}

test('기존 21개 팔레트 ID를 유지하고 6분류 18개 색 구성을 추가한다',()=>{
  assert.deepEqual(Object.keys(CHART_PALETTES).slice(0,21),originalIds);
  assert.equal(Object.keys(CHART_PALETTES).length,39);
  assert.equal(Object.keys(EXTRA_CHART_PALETTES).length,18);
  for(const group of ['report','vivid','pastel','dark','mono','contrast'])assert.equal(Object.values(EXTRA_CHART_PALETTES).filter(p=>p.group===group).length,3);
  for(const p of Object.values(CHART_PALETTES))assert.ok(p.colors.length && p.colors.every(c=>/^#[0-9A-F]{6}$/i.test(c)));
  assert.equal(new Set(Object.values(EXTRA_CHART_PALETTES).map(p=>p.colors.join())).size,18);
});

test('그룹·이름·색 검색은 현재 테마를 읽고 반환 배열은 원본과 분리된다',()=>{
  const all=chartPaletteOptions(CHART_PALETTES);
  assert.equal(all.length,39);assert.equal(chartPaletteOptions(CHART_PALETTES,'','theme').length,4);
  assert.deepEqual(CHART_PALETTE_GROUPS.map(g=>g.id),['theme','report','vivid','pastel','dark','mono','contrast']);
  assert.equal(chartPaletteOptions(CHART_PALETTES,'대비 주황')[0].id,'contrastBlueOrange');
  assert.equal(chartPaletteOptions(CHART_PALETTES,'REPORTBLUE')[0].id,'reportBlue');
  assert.equal(chartPaletteOptions(CHART_PALETTES,'#285a8e')[0].id,'reportBlue');
  assert.equal(chartPaletteOptions(CHART_PALETTES,'없는 검색어').length,0);
  assert.equal(chartPaletteOptions(CHART_PALETTES,'','unknown').length,0);
  const old=CHART_PALETTES.reportBlue.colors[0];all.find(p=>p.id==='reportBlue').colors[0]='#000000';assert.equal(CHART_PALETTES.reportBlue.colors[0],old);
  const colors=[...DEFAULT_THEME];colors[4]='123456';
  withThemeColors(colors,()=>{
    assert.equal(paletteOf({palette:'office'})[0],'#123456');
    assert.equal(chartPaletteOptions(CHART_PALETTES,'','theme')[0].colors[0],'#123456');
    assert.deepEqual(paletteOf({palette:'reportBlue'}),EXTRA_CHART_PALETTES.reportBlue.colors);
  });
});

test('사용자 팔레트는 HEX만 정규화하며 순서·반복색과 입력 원본을 보존한다',()=>{
  const input=[' #abc ','#aBcDeF','#abc'],before=[...input];
  assert.deepEqual(normalizeChartPalette(input),{colors:['#AABBCC','#ABCDEF','#AABBCC'],error:null});
  assert.deepEqual(input,before);
  assert.deepEqual(normalizeChartPalette('#123, #abcdef\n#000; #fff').colors,['#112233','#ABCDEF','#000000','#FFFFFF']);
  assert.deepEqual(normalizeChartPalette(['#012345']).colors,['#012345']);
  assert.equal(normalizeChartPalette(Array(32).fill('#123456')).colors.length,32);
  for(const invalid of [[],null,'',Array(33).fill('#123456'),['#123456','red'],['rgb(0,0,0)'],['#1234'],['#11223344'],['#12GG56'],['#123456',{}],['#123456',null],['#123456" onload="x']]) {
    const result=normalizeChartPalette(invalid);assert.equal(result.colors,null);assert.ok(result.error);
  }
  assert.match(normalizeChartPalette(['#123456','bad']).error,/2번째/);
});

test('가져온 팔레트는 기존 유효 색과 32색 초과 배열을 자르지 않는다',()=>{
  assert.deepEqual(paletteOf({palette:['#112233','invalid','#AbCdEf']}),['#112233','#AbCdEf']);
  const many=Array.from({length:40},(_,i)=>'#'+i.toString(16).padStart(6,'0'));
  assert.deepEqual(paletteOf({palette:many}),many);
  assert.deepEqual(normalizeChartPalette(many,{maxColors:Infinity}).colors,many.map(c=>c.toUpperCase()));
  assert.deepEqual(paletteOf({palette:[' #abc ']}),['#AABBCC']);
  for(const palette of [[],['bad'],'missing'])assert.deepEqual(paletteOf({palette}),paletteOf({}));
});

for(const [id,palette] of Object.entries(EXTRA_CHART_PALETTES))test(`${palette.label}: 실제 막대색과 표준 XLSX 재열기에서 6색 유지`,()=>{
  const wb=book('column',id),before=pointFills(wb),{wb:loaded,files}=standardReopen(wb);
  assert.equal(before.length,12);assert.deepEqual([...new Set(before)],palette.colors);
  assert.deepEqual(pointFills(loaded),before);
  const xml=xmlText(files['xl/charts/chart1.xml']);
  for(const color of palette.colors)assert.ok(xml.includes(`val="${color.slice(1)}"`));
  assert.deepEqual(chartModelData(loaded,0,loaded.sheets[0].charts[0]).series.map(s=>s.values),[[10,15],[20,25],[30,35],[40,45],[50,55],[60,65]]);
});

for(const type of ['waterfall','pareto'])for(const palette of [['#123456'],['#123456','#ABCDEF']])test(`${type}: ${palette.length}색도 모든 기본색을 순환하며 표준 XLSX 왕복`,()=>{
  const wb=book(type,palette),before=svg(wb),{wb:loaded,files}=standardReopen(wb),after=svg(loaded);
  assert.doesNotMatch(before,/undefined|NaN/);assert.doesNotMatch(after,/undefined|NaN/);
  assert.deepEqual(paletteOf(loaded.sheets[0].charts[0]).map(c=>c.toUpperCase()),palette);
  const colors=descendants(parseXml(xmlText(files['xl/charts/colors1.xml'])),'srgbClr').map(n=>'#'+n.attrs.val);
  assert.deepEqual(colors,palette);
  if(type==='waterfall')assert.deepEqual(pointFills(wb),['#123456',palette[1%palette.length],palette[2%palette.length]]);
  else assert.match(before,new RegExp(`data-pareto="cumulative"[^>]+stroke="${palette[1%palette.length]}"`));
});

test('사용자 색 적용은 한 번의 Undo/Redo로 복원되고 원본 셀·계열 값은 유지된다',()=>{
  const wb=book('column','office'),ch=wb.sheets[0].charts[0],before=structuredClone(ch),colors=normalizeChartPalette('#123 #456 #789').colors;
  wb.transact(()=>wb.setSheetProp(0,'charts',[{...ch,...chartPalettePatch(ch,colors)}]));
  const after=structuredClone(wb.sheets[0].charts[0]);assert.deepEqual([...new Set(pointFills(wb))],colors);
  wb.undo();assert.deepEqual(wb.sheets[0].charts[0],before);
  wb.redo();assert.deepEqual(wb.sheets[0].charts[0],after);
  assert.equal(wb.getValue(0,1,1),10);assert.equal(wb.getValue(0,2,6),65);
  assert.deepEqual(pointFills(standardReopen(wb).wb),pointFills(wb));
});
