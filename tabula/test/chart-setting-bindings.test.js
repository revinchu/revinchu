import { test } from 'node:test';
import assert from 'node:assert/strict';
import { renderChartSvg, filterChart } from '../src/chart.js';
import { parseXml, descendants } from '../src/xml.js';

const data = fmt => ({ categories: ['첫째', '둘째', '셋째'], series: [{ name: '값', values: [10, 40, 20], x: [1, 2, 3], ...fmt }] });
const draw = (chart, source) => parseXml(renderChartSvg({ w: 500, h: 320, legend: 'none', ...chart }, source));
const lines = svg => descendants(svg, 'path').filter(n => n.attrs['data-s'] === '0' && n.attrs['data-p'] === undefined);
const points = svg => ['rect', 'circle', 'path'].flatMap(tag => descendants(svg, tag)).filter(n => n.attrs['data-p'] !== undefined);
const labels = svg => descendants(svg, 'text').filter(n => n.attrs['data-el'] === 'label');

test('분산형의 명시적 부드러운 선은 갤러리 기본보다 우선하며 파선·굵기를 렌더', () => {
  const smooth = lines(draw({ type: 'scatter', scatterStyle: 'lineMarker' }, data({ smooth: true, dash: 'dashDot', lineWidth: 3 })))[0];
  assert.match(smooth.attrs.d, /C/); assert.equal(smooth.attrs['stroke-dasharray'], '12 6 3 6'); assert.equal(smooth.attrs['stroke-width'], '3');
  const straight = lines(draw({ type: 'scatter', scatterStyle: 'smoothMarker' }, data({ smooth: false })))[0];
  assert.doesNotMatch(straight.attrs.d, /C/); assert.match(straight.attrs.d, /L/);
  assert.match(lines(draw({ type: 'scatter', scatterStyle: 'smoothMarker' }, data({})))[0].attrs.d, /C/);
});
test('분산형 표식 없음과 명시 표식은 갤러리 기본보다 우선', () => {
  assert.equal(points(draw({ type: 'scatter', scatterStyle: 'lineMarker' }, data({ marker: 'none' }))).length, 0);
  assert.equal(points(draw({ type: 'scatter', scatterStyle: 'lineMarker' }, data({ marker: false }))).length, 0);
  assert.equal(points(draw({ type: 'scatter', scatterStyle: 'line' }, data({ marker: 'square' }))).length, 3);
  assert.equal(points(draw({ type: 'scatter', scatterStyle: 'line' }, data({}))).length, 0);
});
test('깔때기형 레이블은 명시 숨김·숫자 형식·이름·글꼴 서식을 반영', () => {
  assert.equal(labels(draw({ type: 'funnel' }, data({}))).length, 3, '기존 기본값 유지');
  assert.equal(labels(draw({ type: 'funnel' }, data({ labels: false }))).length, 0);
  assert.equal(labels(draw({ type: 'funnel', labels: false }, data({}))).length, 0);
  const list = labels(draw({ type: 'funnel' }, data({ labels: false, catName: true, serName: true, labelColor: '#ff0000', labelSize: 15, labelBold: false })));
  assert.equal(list[0].text, '값, 첫째');
  assert.equal(list[0].attrs.fill, '#ff0000'); assert.equal(list[0].attrs['font-size'], '20'); assert.equal(list[0].attrs['font-weight'], undefined);
  assert.equal(labels(draw({ type: 'funnel' }, data({ labels: true, numFmt: '0.00' })))[0].text, '10.00');
});
test('깔때기형 개별 색·선택 태그는 필터 전 원본 항목 번호를 보존', () => {
  const source = data({ pointColors: { 2: '#fe2301' } }), before = structuredClone(source);
  const filtered = filterChart(source, { hiddenCats: [1] });
  const svg = draw({ type: 'funnel' }, filtered), p = points(svg);
  assert.deepEqual(p.map(n => n.attrs['data-p']), ['0', '2']);
  assert.ok(p.every(n => n.attrs['data-s'] === '0'));
  assert.equal(p[1].attrs.fill, '#fe2301');
  assert.deepEqual(labels(svg).map(n => n.attrs['data-p']), ['0', '2']);
  assert.deepEqual(source, before);
});


