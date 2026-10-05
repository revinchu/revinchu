import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartData, renderChartSvg } from '../src/chart.js';
import { parseXml, descendants, child } from '../src/xml.js';

const rows = [['광고세트', '클릭수'], ['광고 세트 A', 879], ['광고 세트 B', 5379]];
const data = chartData(rows, 'column');
const draw = (chart, source = data) => parseXml(renderChartSvg({ w: 480, h: 288, type: 'column', ...chart }, source));
const legend = (svg) => descendants(svg, 'g').find((g) => g.attrs['data-el'] === 'legend');
const labels = (svg) => descendants(legend(svg), 'text');

test('단일 계열도 기본·명시한 네 위치에서 실제 범례를 표시한다', () => {
  for (const type of ['column', 'bar', 'line', 'area', 'scatter', 'radar']) {
    for (const position of [undefined, 'b', 't', 'r', 'l']) {
      const svg = draw({ type, legend: position });
      assert.ok(legend(svg), `${type}/${position ?? 'default'}`);
      assert.deepEqual(labels(svg).map((node) => node.text), ['클릭수']);
    }
  }
});

test('명시적 범례 없음은 단일 계열·원형·여러 계열 모두 유지한다', () => {
  for (const type of ['column', 'line', 'pie', 'doughnut']) {
    assert.equal(legend(draw({ type, legend: 'none' })), undefined);
    assert.equal(legend(draw({ type, legend: 'none' }, { ...data, series: [...data.series, { name: '전환수', values: [8, 53] }] })), undefined);
  }
});

test('범례를 기본으로 숨기는 특수 차트는 기본 정책과 명시 표시를 구분한다', () => {
  for (const type of ['histogram', 'funnel']) {
    assert.equal(legend(draw({ type })), undefined, `${type} 기본 숨김`);
    assert.deepEqual(labels(draw({ type, legend: 'b' })).map((node) => node.text), ['클릭수']);
    assert.equal(legend(draw({ type, legend: 'none' })), undefined);
  }
  assert.equal(legend(draw({}, { categories: [], series: [] })), undefined, '데이터 없는 차트에는 빈 범례를 만들지 않는다');
});

test('3D 원형 범례는 한글 항목명과 항목별 색을 보존한다', () => {
  const source = chartData(rows, 'pie');
  source.series[0].pointColors = { 0: '#123456', 1: '#654321' };
  const before = structuredClone(source), svg = draw({ type: 'pie', threeD: true, legend: 'r' }, source);
  assert.deepEqual(labels(svg).map((node) => node.text), ['광고 세트 A', '광고 세트 B']);
  assert.deepEqual(descendants(legend(svg), 'rect').map((node) => node.attrs.fill), ['#123456', '#654321']);
  assert.deepEqual(source, before, '범례를 그릴 때 항목명·값을 바꾸지 않는다');
});

test('명시 범례 글자색은 배경과 같아도 사용자가 지정한 정확한 색을 유지한다', () => {
  for (const [fill, legendColor, expected] of [
    ['#ffffff', '#ffffff', '#ffffff'], ['#000000', '#000000', '#000000'],
    ['#fff', '#ffffff', '#ffffff'], ['#000', '#000', '#000'],
    ['#ffffff', '#123456', '#123456'],
  ]) assert.equal(labels(draw({ fill, legendColor }))[0].attrs.fill, expected);
});

test('가로·세로 범례의 긴 이름은 줄임표와 안전하게 이스케이프한 전체 이름 title을 함께 제공한다', () => {
  const full = '광고세트 <전체 이름> & "원본" 😀 여름 프로모션 방문자를 위한 아주 긴 항목명';
  for (const position of ['b', 't', 'r', 'l']) {
    for (const type of ['column', 'pie']) {
      const source = type === 'pie' ? { categories: [full, '짧은 이름'], series: [{ name: '클릭수', values: [879, 5379] }] }
        : { categories: ['A', 'B'], series: [{ name: full, values: [879, 5379] }] };
      const svg = draw({ type, legend: position }, source), label = labels(svg)[0];
      assert.match(label.text, /…$/);
      assert.equal(child(label, 'title')?.text, full);
      assert.equal(descendants(svg, '전체').length, 0, '이름 문자열을 SVG 요소로 해석하지 않는다');
      if (type === 'pie') assert.equal(child(labels(svg)[1], 'title'), null, '짧은 이름은 중복 제목이 불필요하다');
    }
  }
});
