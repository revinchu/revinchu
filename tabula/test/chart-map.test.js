import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MAP_COUNTRIES, MAP_DATA_SOURCES } from '../src/chart-map-data.js';
import { mapCountry, mapValues, mapColor, MAP_CHARTS } from '../src/chart-map.js';
const draw = (data, chart = {}) => { const parts = [], defs = []; MAP_CHARTS.map.draw({ chart, data, series: data.series.map((s, i) => ({ ...s, _fi: i })), plot: { x: 10, y: 30, w: 620, h: 330 }, parts, defs, uid: 'test', TXT: '#444' }); return '<svg><defs>' + defs.join('') + '</defs>' + parts.join('') + '</svg>'; };
test('고정된 공식지도는205국가/지역·177경계·37작은지역점을 포함한다', () => {
  assert.equal(MAP_COUNTRIES.length, 205); assert.equal(MAP_COUNTRIES.filter(c => c.path).length, 177); assert.equal(MAP_COUNTRIES.reduce((n, c) => n + c.points.length, 0), 37);
  assert.equal(new Set(MAP_COUNTRIES.map(c => c.id)).size, 205);
  assert.equal(MAP_DATA_SOURCES.length, 2); assert.ok(MAP_DATA_SOURCES.every(s => /^[a-f\d]{64}$/.test(s.sha256) && s.url.includes('/v5.1.2/')));
  assert.ok(MAP_COUNTRIES.every(c => !c.path || /^[MLZ\d.,-]+$/.test(c.path)));
});
test('한국어·영문·ISO2·ISO3 및 원본ISO누락국가를 정확히 연결한다', () => {
  for (const key of ['한국', '대한민국', 'KR', 'KOR', 'South Korea', 'Republic of Korea']) assert.equal(mapCountry(key)?.id, 'KOR', key);
  for (const key of ['미국', 'US', 'USA', 'United States', 'U.S.A.']) assert.equal(mapCountry(key)?.id, 'USA', key);
  for (const key of ['영국', 'UK', 'GB', 'GBR']) assert.equal(mapCountry(key)?.id, 'GBR', key);
  assert.equal(mapCountry('FR')?.id, 'FRA'); assert.equal(mapCountry('NO')?.id, 'NOR'); assert.equal(mapCountry('북한')?.id, 'PRK');
  assert.equal(mapCountry('싱가포르')?.id, 'SGP'); assert.equal(mapCountry('SH')?.id, 'SHN'); assert.equal(mapCountry('서울'), null);
  assert.equal(mapCountry('Côte d’Ivoire')?.id, 'CIV');
  for (const country of MAP_COUNTRIES) for (const code of [country.id, country.iso2, country.iso3]) if (code) assert.equal(mapCountry(code)?.id, country.id, code);
});
test('중복국가 합계·음수·0을 유지하고 빈값·미지원지역·오류를0으로위조하지않는다', () => {
  const result = mapValues({ categories: ['한국', 'KR', '미국', '영국', '일본', '서울', '독일'], series: [{ values: [20, 5, -10, 0, null, 50, '#DIV/0!'] }] });
  assert.deepEqual(result.items.map(x => [x.country.id, x.value, x.count]), [['KOR', 25, 2], ['USA', -10, 1], ['GBR', 0, 1]]);
  assert.equal(result.min, -10); assert.equal(result.max, 25); assert.deepEqual(result.unknown.map(x => x.label), ['서울']); assert.deepEqual(result.invalid.map(x => x.label), ['일본', '독일']);
});
test('합계overflow는부분합색칠로숨기지않고유효하지않은국가로제외한다', () => {
  const result = mapValues({ categories: ['US', 'USA'], series: [{ values: [1e308, 1e308] }] });
  assert.deepEqual(result.items, []); assert.equal(result.invalid[0].reason, '합계 범위 초과');
});
test('값0은결측색과다르고음수/양수혼합은0중심발산색을사용한다', () => {
  assert.equal(mapColor(0, -10, 20), '#ffffff'); assert.notEqual(mapColor(0, 0, 10), mapColor(null, 0, 10));
  assert.equal(mapColor(-10, -10, 20, '#ff0000', '#0000ff'), '#ff0000'); assert.equal(mapColor(20, -10, 20, '#ff0000', '#0000ff'), '#0000ff');
  assert.match(mapColor(5, 5, 5), /^#[a-f\d]{6}$/); assert.match(mapColor(1, 0, 10, '" onload="evil'), /^#[a-f\d]{6}$/);
});
test('실제국가경계·싱가포르점·툴팁·연속범례·unknown안내를렌더한다', () => {
  const svg = draw({ categories: ['한국', '싱가포르', '미국', '<script>서울</script>'], series: [{ name: '매출', values: [100, 0, -50, 25] }] });
  assert.match(svg, /<path[^>]+data-country="KOR"[^>]+data-value="100"/); assert.match(svg, /<circle[^>]+data-country="SGP"[^>]+data-value="0"/);
  assert.match(svg, /data-el="legend"/); assert.match(svg, /인식 못한 지역 1개/); assert.match(svg, /&lt;script&gt;서울&lt;\/script&gt;/); assert.doesNotMatch(svg, /<script>|NaN|Infinity/);
  assert.doesNotMatch(draw({ categories: ['KR'], series: [{ values: [10], color: '" onload="evil' }] }), /onload=/);
  assert.doesNotMatch(draw({ categories: ['KR', 'US'], series: [{ values: [-1e308, 1e308] }] }), /NaN|Infinity/);
  assert.match(svg, /대한민국 \(KR\): 100/); assert.match(svg, /작은 국가\/지역 위치 표시/);
});
test('범례숨김·첫계열만사용안내·동일값범례·빈데이터안내', () => {
  const data = { categories: ['한국', '미국'], series: [{ values: [4, 4] }, { values: [500, 700] }] };
  assert.doesNotMatch(draw(data, { legend: 'none' }), /data-el="legend"/); assert.match(draw(data), /첫 번째 값 계열 사용/);
  assert.match(draw({ categories: ['서울'], series: [{ values: [10] }] }), /국가\/지역 이름과 숫자 값을 확인하세요/);
  assert.match(draw({ categories: ['KR'], series: [{ values: [1234], numFmt: '#,##0' }] }), /1,234/);
});
