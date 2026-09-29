// 데이터 분석 도구 · FORECAST.ETS · 정규식 찾기 · 시나리오 xlsx 왕복 · 해 찾기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  splitGroups, descriptive, matrixTool, regression, histogram, tTest, fTest, anova1, movingAverage, expSmoothing, solveMin,
} from '../src/analysis.js';
import { timeAxis } from '../src/ets.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const near = (a, b, eps = 1e-9) => assert.ok(Math.abs(a - b) <= eps * Math.max(1, Math.abs(b)), `${a} ≉ ${b}`);
const x = [1, 2, 3, 4, 5, 6, 7, 8];
const y = [2.1, 3.9, 6.2, 7.8, 10.1, 12.2, 13.8, 16.1];
const z = [5, 3, 4, 8, 6, 2, 9, 1];

function sheetWith(cols) {
  const wb = new Workbook();
  wb.transact(() => cols.forEach((col, c) => col.forEach((v, r) => wb.setInput(0, r, c, String(v)))));
  return wb;
}
const evalIn = (wb, f) => { wb.transact(() => wb.setInput(0, 30, 10, f)); return wb.getValue(0, 30, 10); };

test('기술 통계 · 상관 · 공분산이 엑셀 함수와 같음', () => {
  const wb = sheetWith([x, y, z]);
  const g = splitGroups(x.map((v, i) => [v, y[i], z[i]]));
  const d = descriptive([g[1]], { kth: 2, kthSmall: 2 }).rows;
  const get = (label) => d.find((r) => r[0] === label)[1];
  near(get('평균'), evalIn(wb, '=AVERAGE(B1:B8)'));
  near(get('표준 편차'), evalIn(wb, '=STDEV.S(B1:B8)'));
  near(get('첨도'), evalIn(wb, '=KURT(B1:B8)'));
  near(get('왜도'), evalIn(wb, '=SKEW(B1:B8)'));
  near(get('신뢰 수준(95.0%)'), evalIn(wb, '=CONFIDENCE.T(0.05,STDEV.S(B1:B8),8)'), 1e-7);
  assert.equal(get('최대값(2)'), 13.8);
  const cor = matrixTool(g, 'correl').rows;
  near(cor[3][1], evalIn(wb, '=CORREL(A1:A8,C1:C8)'));
  assert.equal(cor[1][2], null); // 위쪽 삼각형은 비움 (엑셀과 같음)
  near(matrixTool(g, 'covar').rows[3][1], evalIn(wb, '=COVARIANCE.P(A1:A8,C1:C8)'));
});

test('회귀 분석: 계수 · 결정계수 · F · P-값', () => {
  const wb = sheetWith([x, y, z]);
  assert.equal(regression(y, [x, z]).coef.length, 3);
  // 단순 회귀는 SLOPE · INTERCEPT · RSQ 와 같아야 함
  const s = regression(y, [x]);
  near(s.coef[0], evalIn(wb, '=INTERCEPT(B1:B8,A1:A8)'));
  near(s.coef[1], evalIn(wb, '=SLOPE(B1:B8,A1:A8)'));
  near(s.rows.find((q) => q[0] === '결정계수')[1], evalIn(wb, '=RSQ(B1:B8,A1:A8)'));
  const coefRow = s.rows.find((q) => q[0] === 'X 1');
  near(coefRow[2], evalIn(wb, '=INDEX(LINEST(B1:B8,A1:A8,TRUE,TRUE),2,1)'), 1e-7);
  assert.throws(() => regression([1, 2], [[1, 2]]), /관측수/);
});

test('t · F-검정 · 분산 분석 · 히스토그램', () => {
  const wb = sheetWith([x, y, z]);
  const tt = (kind) => tTest(y, z, { kind }).rows.find((q) => q[0] === 'P(T<=t) 양측 검정')[1];
  near(tt('equal'), evalIn(wb, '=T.TEST(B1:B8,C1:C8,2,2)'), 1e-8);
  near(tt('paired'), evalIn(wb, '=T.TEST(B1:B8,C1:C8,2,1)'), 1e-8);
  const f = fTest(y, z).rows.find((q) => q[0] === 'P(F<=f) 단측 검정')[1];
  near(f * 2, evalIn(wb, '=F.TEST(B1:B8,C1:C8)'), 1e-8);
  const a = anova1(splitGroups(x.map((v, i) => [v, y[i], z[i]])));
  assert.ok(a.F > 0);
  const h = histogram(z, [2, 5, 8], { cumulative: true });
  assert.deepEqual(h.rows.map((q) => q.slice(0, 2)), [['계급', '빈도수'], [2, 2], [5, 3], [8, 2], ['기타', 1]]);
  assert.equal(h.rows.at(-1)[2], 1);
});

test('이동 평균 · 지수 평활은 엑셀처럼 수식으로 출력', () => {
  const inRef = (i) => `B${i + 2}`;
  const outRef = (i) => `D${i + 2}`;
  const ma = movingAverage(5, inRef, outRef, { interval: 3 }).rows;
  assert.deepEqual(ma.map((q) => q[0].f ?? q[0].err), ['#N/A', '#N/A', '=AVERAGE(B2:B4)', '=AVERAGE(B3:B5)', '=AVERAGE(B4:B6)']);
  const es = expSmoothing(3, inRef, outRef, { damping: 0.3 }).rows;
  assert.deepEqual(es.map((q) => q[0].f ?? q[0].err), ['#N/A', '=B2', '=0.7*B3+0.3*D3']);
});

test('FORECAST.ETS: 계절성 찾기 · 예측 · 신뢰 구간 · 월 단위 시간 표시줄', () => {
  const vals = Array.from({ length: 36 }, (_, i) => 100 + 2 * i + 10 * Math.sin((2 * Math.PI * i) / 12) + ((Math.sin(i * 12.9898) * 43758.5453) % 1) * 3);
  // 매월 1일 (간격이 28~31일로 달라도 월 단위로 인식)
  const dates = Array.from({ length: 36 }, (_, i) => (Date.UTC(2024, i, 1) - Date.UTC(1899, 11, 30)) / 86400000);
  const wb = sheetWith([dates, vals]);
  assert.equal(evalIn(wb, '=FORECAST.ETS.SEASONALITY(B1:B36,A1:A36)'), 12);
  const next = (Date.UTC(2027, 0, 1) - Date.UTC(1899, 11, 30)) / 86400000;
  const f = evalIn(wb, `=FORECAST.ETS(${next},B1:B36,A1:A36)`);
  assert.ok(Math.abs(f - (100 + 72)) < 8, `예측 ${f}`);
  const ci = evalIn(wb, `=FORECAST.ETS.CONFINT(${next},B1:B36,A1:A36,0.95)`);
  assert.ok(ci > 0 && ci < 20);
  assert.equal(evalIn(wb, '=FORECAST.ETS.STAT(B1:B36,A1:A36,8)'), 1); // 단계: 1개월
  const ax = timeAxis(dates);
  assert.equal(ax.months, 1);
  assert.equal(ax.at(36), next);
  assert.equal(timeAxis([1, 2, 4]).step, 1);
  assert.equal(timeAxis([1, 2.5, 4.1]), null);
});

test('정규식 찾기: XLOOKUP/XMATCH match_mode 3 · REGEXVLOOKUP · REGEXMATCHPOS · REGEXSEARCH', () => {
  const wb = new Workbook();
  const data = [['브랜드_PC', 100], ['일반_MO_2024', 200], ['Brand_MO', 300]];
  wb.transact(() => data.forEach((r, i) => r.forEach((v, c) => wb.setInput(0, i, c, String(v)))));
  assert.equal(evalIn(wb, '=XLOOKUP("^일반_.*MO",A1:A3,B1:B3,"없음",3)'), 200);
  assert.equal(evalIn(wb, '=XLOOKUP("^없는",A1:A3,B1:B3,"없음",3)'), '없음');
  assert.equal(evalIn(wb, '=XMATCH("MO$",A1:A3,3)'), 3);
  assert.equal(evalIn(wb, '=XMATCH("_MO",A1:A3,3,-1)'), 3);
  assert.equal(evalIn(wb, '=INDEX(B1:B3,REGEXMATCHPOS("^brand",A1:A3,1))'), 300);
  assert.equal(evalIn(wb, '=REGEXVLOOKUP("\\d{4}$",A1:B3,2)'), 200);
  assert.equal(evalIn(wb, '=REGEXSEARCH("[0-9]+","abc_2024x")'), 5);
  assert.equal(evalIn(wb, '=INDEX(B1:B3,MATCH(TRUE,REGEXTEST(A1:A3,"_PC$"),0))'), 100);
  assert.equal(evalIn(wb, '=XLOOKUP("[",A1:A3,B1:B3,,3)').code, '#VALUE!');
});

test('시나리오: xlsx 왕복 · 행 삽입 시 이동', () => {
  const wb = sheetWith([[1, 2, 3]]);
  wb.transact(() => wb.setSheetProp(0, 'scenarios', [{ name: '낙관', comment: '메모', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['10', '20'] }]));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(back.sheets[0].scenarios.map(({ name, comment, cells, values }) => ({ name, comment, cells, values })), [{ name: '낙관', comment: '메모', cells: [{ r: 0, c: 0 }, { r: 1, c: 0 }], values: ['10', '20'] }]);
  wb.transact(() => wb.insertRows(0, 1, 2));
  assert.deepEqual(wb.sheets[0].scenarios[0].cells, [{ r: 0, c: 0 }, { r: 3, c: 0 }]);
});

test('해 찾기: 선형 · 정수 · 비선형', () => {
  const lp = solveMin(([a, b]) => ({ f: -(3 * a + 2 * b), pen: Math.max(0, a + b - 4) + Math.max(0, a + 3 * b - 6) + Math.max(0, -a) + Math.max(0, -b) }), [0, 0]);
  near(lp.x[0], 4, 1e-6);
  near(-lp.f, 12, 1e-6);
  const ip = solveMin(([a, b]) => ({ f: (a - 1.3) ** 2 + (b - 2.7) ** 2, pen: 0 }), [0, 0], { ints: [0, 1] });
  assert.deepEqual(ip.x, [1, 3]);
  const rb = solveMin(([a, b]) => (1 - a) ** 2 + 100 * (b - a * a) ** 2, [-1, 1]);
  near(rb.x[0], 1, 1e-4);
});

test('피벗 조건부 서식: 값 필드 전체 규칙이 xlsx <conditionalFormats> 로 왕복', async () => {
  const { unzip, textOf } = await import('../src/zip.js');
  const wb = sheetWith([['주차', '1주', '1주', '2주'], ['비용', 10, 20, 30]]);
  wb.transact(() => wb.setSheetProp(0, 'pivot', { name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 1 }, rows: ['주차'], cols: [], pages: [], filters: {}, values: [{ field: '비용', agg: 'sum' }], top: 0, left: 5, area: { r1: 0, c1: 5, r2: 3, c2: 6 } }));
  wb.transact(() => wb.setSheetProp(0, 'cond', [{ r1: 1, c1: 6, r2: 3, c2: 6, type: 'bar', color: '#638ec6', pivot: { name: '피벗1', scope: 'data', value: '합계 : 비용' } }]));
  const bytes = writeXlsx(wb);
  assert.match(textOf(unzip(bytes)['xl/pivotTables/pivotTable1.xml']), /<conditionalFormat scope="data" priority="1">/);
  const back = new Workbook(readXlsx(bytes).data);
  assert.deepEqual(back.sheets[0].cond[0].pivot, { name: '피벗1', scope: 'data', value: '합계 : 비용' });
});

test('위셀 피벗 차트: 고른 지표만 · 엑셀에는 피벗 범위 참조 일반 차트 · 다시 열면 피벗 차트', async () => {
  const { unzip, textOf } = await import('../src/zip.js');
  const { chartModelData } = await import('../src/chart.js');
  const wb = sheetWith([['주차', '1주', '1주', '2주'], ['비용', 10, 20, 30], ['클릭', 1, 2, 3], ['전환', 0, 1, 1]]);
  const values = [{ field: '비용', agg: 'sum' }, { field: '클릭', agg: 'sum' }, { field: '전환', agg: 'sum' }];
  wb.transact(() => wb.setSheetProp(0, 'pivot', { name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 3 }, rows: ['주차'], cols: [], pages: [], filters: {}, values, top: 0, left: 6, area: { r1: 0, c1: 6, r2: 3, c2: 9 } }));
  const chart = { id: 'c1', type: 'column', title: 't', pivot: { sheet: 'Sheet1', name: '피벗1', values: ['합계 : 전환', '합계 : 비용'] }, x: 0, y: 0, w: 400, h: 300 };
  wb.transact(() => wb.setSheetProp(0, 'charts', [chart]));
  const data = chartModelData(wb, 0, chart);
  assert.deepEqual(data.series.map((s) => s.name), ['합계 : 전환', '합계 : 비용']);
  const bytes = writeXlsx(wb);
  const xml = textOf(unzip(bytes)['xl/charts/chart1.xml']);
  assert.doesNotMatch(xml, /<c:pivotSource>/);
  assert.equal((xml.match(/<c:ser>/g) ?? []).length, 2);
  const back = new Workbook(readXlsx(bytes).data);
  assert.deepEqual(back.sheets[0].charts[0].pivot, chart.pivot);
});

test('차트 서식: 데이터 표 · 추세선 · 선 종류 · 요소 색 · 겹치기 · 값 축 거꾸로 xlsx 왕복', async () => {
  const { renderChartSvg } = await import('../src/chart.js');
  const wb = sheetWith([['월', '1월', '2월', '3월', '4월'], ['비용', 10, 20, 30, 25], ['ROAS', 3, 4, 2, 5]]);
  const chart = {
    id: 'c', type: 'combo', title: 't', x: 0, y: 0, w: 500, h: 320, dataTable: true, overlap: 20, varyColors: false, axes: { y: { reverse: true } },
    series: [{ name: { text: '비용' }, cat: { r1: 1, c1: 0, r2: 4, c2: 0 }, val: { r1: 1, c1: 1, r2: 4, c2: 1 } }, { name: { text: 'ROAS' }, cat: { r1: 1, c1: 0, r2: 4, c2: 0 }, val: { r1: 1, c1: 2, r2: 4, c2: 2 } }],
    seriesFmt: [{ type: 'column', trend: 'linear', pointColors: { 2: '#ff0000' }, labels: true, labelPos: 'center' }, { type: 'line', axis: 1, dash: 'dash', trend: 'movingAvg', trendPeriod: 2 }],
  };
  wb.transact(() => wb.setSheetProp(0, 'charts', [chart]));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].charts[0];
  assert.equal(back.dataTable, true);
  assert.equal(back.overlap, 20);
  assert.equal(back.axes.y.reverse, true);
  assert.equal(back.seriesFmt[0].trend, 'linear');
  assert.equal(back.seriesFmt[0].pointColors[2], '#ff0000');
  assert.equal(back.seriesFmt[0].labelPos, 'center');
  assert.equal(back.seriesFmt[1].dash, 'dash');
  assert.equal(back.seriesFmt[1].trend, 'movingAvg');
  assert.equal(back.seriesFmt[1].trendPeriod, 2);
  const svg = renderChartSvg({ ...chart, labels: true }, { categories: ['1월', '2월', '3월', '4월'], series: [{ name: '비용', values: [10, 20, 30, 25], type: 'column', trend: 'linear', pointColors: { 2: '#ff0000' } }, { name: 'ROAS', values: [3, 4, 2, 5], type: 'line', axis: 1, dash: 'dash' }] });
  assert.match(svg, /fill="#ff0000"/);
  assert.match(svg, /stroke-dasharray/);
  assert.match(svg, />1월</);
});
