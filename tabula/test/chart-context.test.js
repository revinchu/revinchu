import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartResetFormattingPatch, chartTemplateFormat, applyChartTemplatePatch } from '../src/chart-context.js';
import { writeChartTemplate, readChartTemplate } from '../src/chart-template.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip } from '../src/zip.js';
import { chartModelData } from '../src/chart.js';

test('선택한 차트 요소만 초기화하고 데이터와 축 척도를 유지', () => {
 const chart={id:'c',range:{r1:0,c1:0,r2:3,c2:1},plotFill:'#ff0000',fill:'#123456',axes:{y:{min:3}},hiddenSeries:[0],seriesFmt:[{type:'line',secondary:true,color:'#ff0000',pointColors:{0:'#000000',1:'#ffffff'}}]};
 const plot={...chart,...chartResetFormattingPatch(chart,{kind:'plot'},{plotFill:'#eeeeee'})};
 assert.equal(plot.plotFill,'#eeeeee');assert.equal(plot.fill,chart.fill);assert.deepEqual(plot.range,chart.range);assert.deepEqual(plot.axes,chart.axes);
 const point=chartResetFormattingPatch(chart,{kind:'point',s:0,p:1});assert.deepEqual(point.seriesFmt[0].pointColors,{0:'#000000'});assert.equal(point.seriesFmt[0].type,'line');
 const series=chartResetFormattingPatch(chart,{kind:'series',s:0});assert.equal(series.seriesFmt[0].color,undefined);assert.equal(series.seriesFmt[0].secondary,true);
});
test('CRTX는 실제 문서·원본값을 포함하지 않고 스타일을 왕복한다', () => {
 const chart={id:'secret-id',sheet:'기밀시트',title:'기밀제목',range:{r1:0,c1:0,r2:2,c2:1},type:'column',fill:'#123456',legend:'b',series:[{name:'비밀계열',cache:[918273645],val:{sheet:'기밀시트',r1:1,c1:1,r2:1,c2:1}}],seriesFmt:[{color:'#aabbcc'}]};
 const bytes=writeChartTemplate(chart),files=unzip(bytes);const xml=Object.values(files).map(b=>new TextDecoder().decode(b)).join('');
 assert.doesNotMatch(xml,/기밀시트|기밀제목|비밀계열|918273645|secret-id/);assert.ok(files['_rels/.rels']);assert.ok(files['xl/charts/chart1.xml']);
 const read=readChartTemplate(bytes);assert.equal(read.type,'column');assert.equal(read.fill,'#123456');assert.equal(read.seriesFmt[0].color,'#aabbcc');
 const patch=applyChartTemplatePatch(chart,read);assert.equal(patch.range,undefined);assert.equal(patch.series,undefined);assert.equal(patch.title,undefined);assert.equal(chart.title,'기밀제목');
});
test('서식 입력은 원본 참조·본문·임의 속성을 제거한다',()=>{
 assert.deepEqual(chartTemplateFormat({type:'pie',sheet:'secret',range:{},series:[{}],evil:'x'}),{type:'pie'});
 assert.throws(()=>readChartTemplate(new Uint8Array()),/ZIP|zip|서식|압축|파일/);
});

test('슬라이서 대체 텍스트·매크로·인쇄/보호 속성의 표준 XLSX 왕복',()=>{
 const wb=new Workbook({sheets:[{name:'원본',cells:{'0,0':{raw:'항목'},'1,0':{raw:'가'}},tables:[{id:'t',name:'표1',r1:0,c1:0,r2:1,c2:0,header:true}],slicers:[{id:'sl',caption:'항목',source:{kind:'table',table:'표1',column:'항목'},x:10,y:10,w:160,h:180,alt:'항목을 선택하는 필터',macro:'MyFilter',noPrint:true,locked:false}]}]});
 const read=readXlsx(writeXlsx(wb)).data.sheets[0].slicers[0];assert.ok(read);assert.equal(read.alt,'항목을 선택하는 필터');assert.equal(read.macro,'MyFilter');assert.equal(read.noPrint,true);assert.equal(read.locked,false);
});

 test('3D cylinder chart template retains column shape', () => {
  const back = readChartTemplate(writeChartTemplate({type:'column',threeD:true,barShape:'cylinder'}));
  assert.equal(back.seriesFmt[0].barShape, 'cylinder');
});

test('계열 템플릿은 지원 서식을 보존하고 값·참조·내부 인덱스를 제거한다', () => {
  const style = { type: 'line', axis: 1, grouping: 'stacked', barShape: 'cylinder', color: '#123456', outline: '#654321', lineWidth: 2.5, dash: 'dash', marker: 'diamond', markerSize: 9, markerColor: '#abcdef', smooth: true, labels: false, catName: true, serName: true, pct: true, numFmt: '0.0%', percentFmt: '0%', labelPos: 'insideEnd', labelSeparator: '\n', labelColor: '#ffffff', labelSize: 13, labelBold: true, explode: 12, trend: 'movingAvg', trendColor: '#ff0000', trendPeriod: 3, trendForward: 2, colors: ['#000000', null, '#ffffff'], pointColors: { 1: '#112233' }, pointExplosion: { 2: 25 }, hierarchyColors: { '["분류"]': '#778899' }, grad: { ang: 40, stops: [[0, '#123456'], [1, '#ffffff']] }, shadow: true };
  const hostile = { ...structuredClone(style), values: [998877], name: '비공개 계열', val: { sheet: '비공개시트', r1: 1 }, cat: {}, cache: [998877], catCache: ['비공개'], catLevels: [['비공개']], x: [998877], size: [998877], xCache: [998877], sizeCache: [998877], _fi: 99, _pi: [99], _upper: [998877], _lower: [998877], arbitrary: '비공개' };
  const before = structuredClone(hostile), out = chartTemplateFormat({ type: 'column', seriesFmt: [hostile] });
  assert.deepEqual(out.seriesFmt, [style]); assert.deepEqual(hostile, before);
  out.seriesFmt[0].grad.stops[0][1] = '#000000'; out.seriesFmt[0].pointColors[1] = '#ffffff';
  assert.equal(hostile.grad.stops[0][1], '#123456'); assert.equal(hostile.pointColors[1], '#112233');
  assert.deepEqual(chartTemplateFormat({ seriesFmt: { values: [1] } }), {});
  assert.deepEqual(chartTemplateFormat({ seriesFmt: [null, { color: { values: [1] }, grad: { stops: [] }, pointColors: { bad: '#fff', 0: { cache: [1] } } }] }).seriesFmt, [{}, { pointColors: {} }]);
});

test('가져온 CRTX 확장의 계열 데이터로 현재 차트 값을 덮지 않는다', () => {
  const files = unzip(writeChartTemplate({ type: 'column', fill: '#fff000' }));
  const payload = { type: 'column', seriesFmt: [{ values: [999], name: '바뀐 이름', cache: [999], x: [999], size: [999], _fi: 999, _pi: [999], color: '#123456' }] };
  const json = JSON.stringify(payload).replaceAll('&', '&amp;').replaceAll('"', '&quot;');
  const path = 'xl/charts/chart1.xml', xml = new TextDecoder().decode(files[path]);
  assert.match(xml, /<tb:props json="/);
  files[path] = xml.replace(/<tb:props json="[^"]*"\/>/, `<tb:props json="${json}"/>`);
  const fmt = readChartTemplate(zip(files)); assert.deepEqual(fmt.seriesFmt, [{ color: '#123456' }]);
  const wb = new Workbook({ sheets: [{ name: '자료', cells: { '0,0': { raw: '원래 이름' }, '1,0': { raw: '10' } } }] });
  const chart = { type: 'column', series: [{ name: { ref: { sheet: '자료', r1: 0, c1: 0, r2: 0, c2: 0 } }, val: { sheet: '자료', r1: 1, c1: 0, r2: 1, c2: 0 }, catCache: ['범주'] }] };
  const next = { ...chart, ...applyChartTemplatePatch(chart, fmt) }, data = chartModelData(wb, 0, next);
  assert.equal(next.series, chart.series); assert.equal(data.series[0].name, '원래 이름'); assert.deepEqual(data.series[0].values, [10]); assert.equal(data.series[0].color, '#123456'); assert.equal(data.series[0]._fi, 0);
  wb.setInput(0, 1, 0, '20'); assert.deepEqual(chartModelData(wb, 0, next).series[0].values, [20]);
});

test('CRTX 저장 시 계열 서식에 끼어든 값과 이름은 예시 데이터로 유출되지 않는다', () => {
  const chart = { type: 'column', seriesFmt: [{ values: [918273645], name: '민감계열표식', cache: [918273645], color: '#123456', grad: { stops: [[0, '#123456', '민감계열표식'], [1, '#ffffff']], source: '민감계열표식' }, shadow: { source: '민감계열표식' } }] };
  const xml = Object.values(unzip(writeChartTemplate(chart))).map(bytes => new TextDecoder().decode(bytes)).join('');
  assert.doesNotMatch(xml, /민감계열표식|918273645/);
  assert.equal(readChartTemplate(writeChartTemplate(chart)).seriesFmt[0].color, '#123456');
});

test('텍스트 초기화는 새 글꼴 속성 모두와 대상만 복원하고 계열 채우기는 보존한다', () => {
  const chart = { font: 'Arial', size: 12, titleFont: 'Georgia', titleItalic: true, titleUnderline: true, legendFont: 'Verdana', legendStrike: true, axes: { x: { title: '기간', font: 'Arial', size: 14, min: 1, titleFont: 'Georgia', titleSize: 20 }, y: { title: '매출', max: 100 } }, seriesFmt: [{ color: '#ff0000', lineWidth: 3, labelFont: 'Georgia', labelItalic: true, pointLabelStyles: { 0: { font: 'Verdana' }, 1: { size: 16 } } }], dataTableText: { font: 'Georgia' } };
  const before = structuredClone(chart);
  const title = { ...chart, ...chartResetFormattingPatch(chart, { kind: 'title' }) };
  assert.equal(title.titleFont, undefined); assert.equal(title.titleItalic, undefined); assert.equal(title.titleUnderline, undefined); assert.equal(title.legendFont, 'Verdana');
  const axis = { ...chart, ...chartResetFormattingPatch(chart, { kind: 'axis-title-x' }) };
  assert.deepEqual(axis.axes.x, { title: '기간', font: 'Arial', size: 14, min: 1 }); assert.deepEqual(axis.axes.y, chart.axes.y);
  const tick = { ...chart, ...chartResetFormattingPatch(chart, { kind: 'axis-x' }) };
  assert.deepEqual(tick.axes.x, { title: '기간', min: 1, titleFont: 'Georgia', titleSize: 20 });
  const point = chartResetFormattingPatch(chart, { kind: 'label', s: 0, p: 0 }).seriesFmt[0];
  assert.deepEqual(point.pointLabelStyles, { 1: { size: 16 } }); assert.equal(point.labelFont, 'Georgia'); assert.equal(point.color, '#ff0000');
  const labels = chartResetFormattingPatch(chart, { kind: 'label', s: 0 }).seriesFmt[0];
  assert.deepEqual(labels, { color: '#ff0000', lineWidth: 3 });
  const whole = { ...chart, ...chartResetFormattingPatch(chart) };
  assert.equal(whole.font, undefined); assert.equal(whole.legendFont, undefined); assert.equal(whole.dataTableText, undefined);
  assert.deepEqual(whole.axes, { x: { title: '기간', min: 1 }, y: { title: '매출', max: 100 } }); assert.deepEqual(whole.seriesFmt[0], { color: '#ff0000', lineWidth: 3 });
  assert.deepEqual(chart, before);
});

test('새 차트 글꼴 템플릿은 축 문구·범위·민감 데이터 없이 텍스트 서식만 보존한다', () => {
  const chart = { type: 'column', font: 'Arial', italic: true, titleFont: 'Georgia', titleUnderline: true, legendFont: 'Verdana', legendStrike: true, axes: { x: { title: '비공개기간문구', font: 'Arial', min: 918273645, max: 987654321, titleFont: 'Georgia', labelRotation: 45 }, y: { title: '비공개금액문구', size: 14, titleItalic: true, metadata: '유출금지표식' } }, dataTableText: { font: 'Georgia', color: '#123456', values: ['유출금지표식'] }, seriesFmt: [{ labelFont: 'Georgia', labelItalic: true, pointLabelStyles: { 1: { font: 'Verdana', size: 16, color: '#aa3344', values: [918273645], name: '유출금지표식' }, invalid: { font: '민감' } } }] };
  const before = structuredClone(chart), format = chartTemplateFormat(chart);
  assert.deepEqual(format.axes, { x: { font: 'Arial', titleFont: 'Georgia' }, y: { size: 14, titleItalic: true } });
  assert.deepEqual(format.dataTableText, { font: 'Georgia', color: '#123456' });
  assert.deepEqual(format.seriesFmt[0].pointLabelStyles, { 1: { font: 'Verdana', size: 16, color: '#aa3344' } });
  const bytes = writeChartTemplate(chart), xml = Object.values(unzip(bytes)).map(b => new TextDecoder().decode(b)).join('');
  assert.doesNotMatch(xml, /비공개기간문구|비공개금액문구|918273645|987654321|유출금지표식/);
  const back = readChartTemplate(bytes);
  for (const key of ['font', 'italic', 'titleFont', 'titleUnderline', 'legendFont', 'legendStrike', 'dataTableText']) assert.deepEqual(back[key], format[key], key);
  assert.equal(back.axes.x.font, 'Arial'); assert.equal(back.axes.x.titleFont, 'Georgia'); assert.equal(back.axes.y.titleItalic, true);
  assert.deepEqual(back.seriesFmt[0].pointLabelStyles, format.seriesFmt[0].pointLabelStyles);
  const current = { title: '현재 차트', axes: { x: { title: '현재 기간', font: '이전글꼴', min: 0, max: 15, labelRotation: 30 }, y: { title: '현재 금액', numFmt: '0.0%', color: '#ffffff' } } };
  const next = { ...current, ...applyChartTemplatePatch(current, format) };
  assert.equal(next.title, '현재 차트'); assert.equal(next.axes.x.title, '현재 기간'); assert.equal(next.axes.x.max, 15); assert.equal(next.axes.x.labelRotation, 30); assert.equal(next.axes.x.font, 'Arial');
  assert.equal(next.axes.y.numFmt, '0.0%'); assert.equal(next.axes.y.title, '현재 금액'); assert.equal(next.axes.y.color, undefined); assert.equal(next.axes.y.size, 14);
  assert.deepEqual(chart, before);
});

test('원본 계열 안에 저장된 레이블 글꼴도 초기화하지만 원본 값·참조는 보존한다', () => {
  const ref = { sheet: '자료', r1: 1, c1: 1, r2: 3, c2: 1 }, cache = [10, 20, 30];
  const chart = { series: [{ name: { text: '매출' }, val: ref, cache, labelFont: 'Georgia', labelItalic: true, pointLabelStyles: { 0: { size: 15 }, 1: { font: 'Verdana' } } }] };
  const point = chartResetFormattingPatch(chart, { kind: 'label', s: 0, p: 0 });
  assert.deepEqual(point.series[0].pointLabelStyles, { 1: { font: 'Verdana' } }); assert.equal(point.series[0].labelFont, 'Georgia');
  for (const patch of [chartResetFormattingPatch(chart), chartResetFormattingPatch(chart, { kind: 'label', s: 0 }), applyChartTemplatePatch(chart, { font: 'Arial' })]) {
    assert.equal(patch.series[0].labelFont, undefined); assert.equal(patch.series[0].pointLabelStyles, undefined); assert.equal(patch.series[0].cache, cache); assert.equal(patch.series[0].val, ref); assert.deepEqual(patch.series[0].name, chart.series[0].name);
  }
  assert.equal(chart.series[0].labelFont, 'Georgia');
});
