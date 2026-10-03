import test from 'node:test';
import assert from 'node:assert/strict';
import { axisTextWidth, categoryAxisLayout, categoryAxisLabelSvg } from '../src/chart-axis-labels.js';
import { renderChartSvg } from '../src/chart.js';
import { parseXml, descendants } from '../src/xml.js';

const names = ['브랜드검색 시드니면세점 PC', '브랜드검색 시드니면세점 모바일', '쇼핑검색 브레인올로지 PC', '쇼핑검색 브레인올로지 모바일', '파워링크 신규 방문 고객', '파워링크 재방문 고객', '일반키워드 어린이 건강식품', '브랜드키워드 성인 건강식품', '일반키워드 주말 할인행사', '브랜드키워드 평일 할인행사', '신규 캠페인 검색 광고', '재방문 캠페인 배너 광고'];
const source = (type, categories = names) => ({ categories: [...categories], series: type === 'boxWhisker' ? categories.map((name, i) => ({ name, values: [10+i, 20+i, 25+i, 32+i, 39+i] })) : [{ name: '광고비', values: categories.map((_, i) => 130+i*17) }] });
const labels = svg => descendants(parseXml(svg), 'text').filter(n => n.attrs['data-axis-label'] !== undefined);
const visibleText = node => node.text + node.children.map(visibleText).join('');
const noWhitespace = s => s.replace(/\s/g, '');

for (const type of ['column', 'line', 'area', 'combo', 'bar', 'waterfall', 'boxWhisker']) {
  test(`${type}: 일반 크기에서 긴 한글 항목을 하나만 남기지 않고 전부 보존`, () => {
    for (const [w, count] of [[420, 6], [640, 8], [900, 12]]) {
      const data = source(type, names.slice(0, count));
      if (type === 'combo') data.series.push({ name:'전환율', values: data.categories.map((_, i) => i+1), type:'line', axis:1 });
      const chart = { type, w, h:380, legend:'none', title:'광고 성과' };
      const beforeData = structuredClone(data), beforeChart = structuredClone(chart);
      const svg = renderChartSvg(chart, data), text = labels(svg);
      assert.equal(text.length, count, `${w}px의 ${count}항목`);
      assert.deepEqual(text.map(n => n.attrs['aria-label']), data.categories);
      assert.deepEqual(text.map(n => noWhitespace(visibleText(n))), data.categories.map(noWhitespace));
      assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
      assert.deepEqual(data, beforeData); assert.deepEqual(chart, beforeChart);
    }
  });
}

test('간격 1은 밀집해도 모든 항목, 간격 3은 해당 원본 인덱스만 표시', () => {
  for (const type of ['column', 'bar', 'waterfall', 'boxWhisker']) {
    const data = source(type);
    for (const interval of [1, 3]) {
      const result = labels(renderChartSvg({ type, w:420, h:380, legend:'none', axes:{ x:{ labelInterval:interval } } }, data));
      const expected = data.categories.map((_, i) => i).filter(i => i % interval === 0);
      assert.deepEqual(result.map(n => Number(n.attrs['data-axis-label'])), expected, type);
      assert.deepEqual(result.map(n => n.attrs['aria-label']), expected.map(i => data.categories[i]));
    }
  }
});

test('짧은 항목은 수평 쓰기를 유지하고, 긴 항목 하나가 나머지 표시 개수를 바꾸지 않는다', () => {
  const short = ['서울', '부산', '대전', '광주', '제주', '인천'];
  const a = categoryAxisLayout(short, { length:360, depth:150 });
  const b = categoryAxisLayout([...short.slice(0, 5), names[0]], { length:360, depth:150 });
  assert.equal(a.angle, 0); assert.equal(a.labels.length, short.length);
  assert.deepEqual(b.labels.map(l => l.index), a.labels.map(l => l.index));
});

test('밀집한 범주의 자동 회전은 45도 줄 높이가 겹치면 90도를 선택한다', () => {
  const categories = Array.from({length:20}, (_, i) => i < 2 ? '가' : '가나다라마바');
  const plan = categoryAxisLayout(categories, { length:400, depth:150, font:12, startRoom:100 });
  assert.equal(plan.every, 1); assert.equal(plan.labels.length, 20);
  assert.equal(plan.angle, -90);
  assert.ok(plan.labels.every(l => l.complete));
});

test('명시 회전은 자동으로 덮어쓰지 않으며 양수·음수·0을 구분한다', () => {
  for (const horizontal of [false, true]) for (const rotation of [-90, -45, 0, 45, 90]) {
    const plan = categoryAxisLayout(['가', '나', '다'], { length:300, depth:150, rotation, horizontal });
    assert.equal(plan.angle, rotation);
    assert.equal(plan.labels.length, 3);
    const svg = categoryAxisLabelSvg(plan, plan.labels[0], 50, 200);
    if (rotation) assert.ok(svg.includes(`rotate(${rotation})`));
    else assert.doesNotMatch(svg, /rotate/);
    assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
  }
});

test('역순 축은 데이터 순서를 보존하면서 화면 위치만 반전한다', () => {
  for (const type of ['column', 'bar', 'waterfall']) {
    const data = source(type, ['서울', '부산', '대전', '광주']);
    const text = labels(renderChartSvg({ type, w:640, h:400, legend:'none', axes:{ x:{ reverse:true, labelInterval:1 } } }, data));
    assert.deepEqual(text.map(n => n.attrs['aria-label']), data.categories);
    const locations = text.map(n => n.attrs.transform.match(/translate\(([-\d.]+) ([-\d.]+)\)/).slice(1).map(Number));
    const axis = type === 'bar' ? 1 : 0;
    assert.ok(locations.every((p, i) => !i || p[axis] < locations[i-1][axis]), type);
  }
});

test('축 숨김·다단계 항목·데이터 표는 중복 레이블을 만들지 않는다', () => {
  const data = source('column', ['서울', '부산', '대전', '광주']);
  data.catLevels = [[{start:0,end:1,text:'앞 묶음'}, {start:2,end:3,text:'뒤 묶음'}]];
  const base = { type:'column', w:640, h:400, legend:'none' };
  const hidden = renderChartSvg({ ...base, axes:{x:{hide:true}} }, data);
  assert.equal(labels(hidden).length, 0); assert.doesNotMatch(hidden, /앞 묶음|뒤 묶음/);
  const multi = renderChartSvg(base, data);
  assert.equal(labels(multi).length, 4); assert.match(multi, /앞 묶음/); assert.match(multi, /뒤 묶음/);
  const table = renderChartSvg({ ...base, dataTable:true }, data);
  assert.equal(labels(table).length, 0, '항목 이름은 데이터 표에 한 번 표시');
  for (const name of data.categories) assert.match(table, new RegExp(name));
});

test('3D 세로·가로 막대도 수동 간격과 범주 이름을 보존한다', () => {
  for (const type of ['column', 'bar']) {
    const data = source(type, names.slice(0, 6));
    const text = labels(renderChartSvg({ type, threeD:true, w:760, h:480, legend:'none', view3D:{rotX:35,rotY:270,depthPercent:150}, axes:{x:{labelInterval:2}} }, data));
    assert.deepEqual(text.map(n => n.attrs['aria-label']), [names[0], names[2], names[4]]);
    assert.deepEqual(text.map(n => Number(n.attrs['data-axis-label'])), [0, 2, 4]);
  }
});

test('Unicode·줄바꿈·XML 문자와 전체 원문은 안전하게 보존된다', () => {
  const input = ['가😀𠮷나', 'Café & <광고> "A"', '첫 줄\n둘째 줄'];
  const before = structuredClone(input), plan = categoryAxisLayout(input, { length:720, depth:150 });
  for (const label of plan.labels) {
    const node = parseXml(categoryAxisLabelSvg(plan, label, 100, 100));
    assert.equal(node.attrs['aria-label'], input[label.index]);
    assert.equal(noWhitespace(visibleText(node)), noWhitespace(input[label.index]));
    assert.equal(node.attrs.onload, undefined);
    for (const line of label.lines) assert.doesNotMatch(line, /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/);
  }
  assert.deepEqual(input, before);
  assert.equal(axisTextWidth('é'), axisTextWidth('e'), '결합 부호는 별도 글자 너비가 아님');
  assert.ok(axisTextWidth('한') > axisTextWidth('i'));
});

test('아주 긴 항목은 명시 제한 안에서 말줄임하되 접근 가능한 원문을 유지한다', () => {
  const name = '대한민국광고성과'.repeat(20);
  const plan = categoryAxisLayout([name], { length:30, depth:40, rotation:0, interval:1 });
  assert.equal(plan.labels.length, 1);
  assert.ok(plan.labels[0].lines.at(-1).endsWith('…'));
  const node = parseXml(categoryAxisLabelSvg(plan, plan.labels[0], 15, 10));
  assert.equal(node.attrs['aria-label'], name);
  assert.ok(visibleText(node).length < name.length);
});


test('줄 너비에 꽉 찬 단어 사이의 공백이 빈 줄을 추가하지 않는다', () => {
  const original = '가나다라 가나다라 가나다라';
  const plan = categoryAxisLayout([original], { length:56, depth:120, rotation:0 });
  assert.deepEqual(plan.labels[0].lines, ['가나다라', '가나다라', '가나다라']);
  assert.ok(plan.labels[0].complete);
  assert.equal(plan.labels[0].text, original);
});


test('가로 막대 항목 축 제목은 왼쪽, 값 축 제목은 아래쪽에 배치한다', () => {
  for (const threeD of [false, true]) {
    const data = source('bar', ['서울', '부산', '대전', '광주']);
    const svg = parseXml(renderChartSvg({ type:'bar', threeD, w:640, h:400, legend:'none', axes:{x:{title:'지역 구분'}, y:{title:'광고비(원)'}} }, data));
    const groups = descendants(svg, 'g');
    const category = descendants(groups.find(n => n.attrs['data-el'] === 'axis-x'), 'text');
    const value = descendants(groups.find(n => n.attrs['data-el'] === 'axis-y'), 'text');
    const categoryTitle = category.find(n => n.text === '지역 구분');
    const valueTitle = value.find(n => n.text === '광고비(원)');
    assert.ok(categoryTitle); assert.ok(valueTitle);
    assert.match(categoryTitle.attrs.transform, /^rotate\(-90 /);
    assert.equal(valueTitle.attrs.transform, undefined);
    const positions = category.filter(n => n.attrs['data-axis-label'] !== undefined).map(n => n.attrs.transform.match(/translate\(([-\d.]+) ([-\d.]+)\)/).slice(1).map(Number));
    assert.ok(positions.every(([x]) => Number(categoryTitle.attrs.x) < x), '항목 제목이 항목 레이블 왼쪽');
    assert.ok(Number(categoryTitle.attrs.y) > Math.min(...positions.map(p => p[1])));
    assert.ok(Number(categoryTitle.attrs.y) < Math.max(...positions.map(p => p[1])));
    assert.ok(Number(valueTitle.attrs.y) > Math.max(...positions.map(p => p[1])), '값 제목이 항목축보다 아래');
    assert.ok(Number(valueTitle.attrs.x) > Number(categoryTitle.attrs.x));
  }
});
