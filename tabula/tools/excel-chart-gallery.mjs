// 실제 Excel용 합성 차트 갤러리. 업무 파일·실행 중인 Excel은 접근하지 않습니다.
// Enum: https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype
import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { CHART_GALLERY, chartModelData, renderChartSvg } from '../src/chart.js';

const out = process.argv[2];
if (!out) throw new Error('사용법: node tools/excel-chart-gallery.mjs D:/Codex/Temp/wixel-gallery-excel [--verify 저장폴더]');
const verifyIndex = process.argv.indexOf('--verify'), verifyRoot = verifyIndex >= 0 ? process.argv[verifyIndex + 1] : null;
if (verifyIndex >= 0 && !verifyRoot) throw new Error('--verify 뒤에 Excel 재저장 폴더가 필요합니다.');
const enumSource = 'https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype';

function excelType(p) {
  const g = ['clustered', 'stacked', 'percentStacked'].indexOf(p.grouping ?? 'clustered');
  switch (p.type) {
    case 'column': return (p.threeD ? [54, 55, 56] : [51, 52, 53])[g];
    case 'bar': return (p.threeD ? [60, 61, 62] : [57, 58, 59])[g];
    case 'area': return (p.threeD ? [-4098, 78, 79] : [1, 76, 77])[g];
    case 'line': return p.threeD ? -4101 : (p.marker && p.marker !== 'none' ? [65, 66, 67] : [4, 63, 64])[g];
    case 'pie': return p.threeD ? p.explode ? 70 : -4102 : p.explode ? 69 : 5;
    case 'doughnut': return p.explode ? 80 : -4120;
    case 'pieOfPie': return 68;
    case 'barOfPie': return 71;
    case 'scatter': return ({ marker: -4169, smoothMarker: 72, smooth: 73, lineMarker: 74, line: 75 })[p.scatterStyle ?? 'marker'];
    case 'bubble': return p.threeD ? 87 : 15;
    case 'stock': return p.volume ? p.ohlc ? 91 : 90 : p.ohlc ? 89 : 88;
    case 'radar': return ({ standard: -4151, marker: 81, filled: 82 })[p.radarStyle ?? 'standard'];
    case 'surface': return ({ surface: 83, wireframe: 84, contour: 85, wireframeContour: 86 })[p.surfaceStyle];
    case 'combo': return -4111; // Actual getter: Excel.Constants.xlCombination; verify each type/axis.
    default: return ({ treemap: 117, histogram: 118, waterfall: 119, sunburst: 120, boxWhisker: 121, pareto: 122, funnel: 123, map: 140 })[p.type];
  }
}

function inputRows(p) {
  if (p.type === 'stock') {
    const header = ['일', ...(p.volume ? ['거래량'] : []), ...(p.ohlc ? ['시가'] : []), '고가', '저가', '종가'];
    return [header, ...['월', '화', '수', '목', '금', '토'].map((d, i) => [d, ...(p.volume ? [1000 + i * 500] : []), ...(p.ohlc ? [20 + i] : []), 30 + i, 10 + i, 25 + i])];
  }
  if (p.type === 'surface') return [['Y / X', 'A', 'B', 'C'], ['행1', 5, 20, 9], ['행2', 10, 4, 30], ['행3', 2, 14, 18]];
  if (['treemap', 'sunburst'].includes(p.type)) return [['대륙', '국가', '금액'], ['아시아', '한국', 10], ['아시아', '일본', 5], ['미주', '미국', 7], ['미주', '캐나다', 4]];
  if (p.type === 'map') return [['국가', '금액'], ['South Korea', 10], ['Japan', 5], ['United States', 7], ['Canada', 4]];
  if (p.type === 'histogram') return [['측정값'], ...[1, 1.5, 2, 2.5, 3, 3, 4, 6, 8, 10].map((v) => [v])];
  if (p.type === 'boxWhisker') return [['관측', '집단 A', '집단 B'], ...[1, 2, 3, 4, 5, 7, 20].map((v, i) => [`표본 ${i + 1}`, v, 2 * v + 3])];
  if (p.type === 'scatter') return [['X', 'Y A', 'Y B'], ...[1, 2, 4, 7, 11, 16].map((v) => [v, v * 2, v + 5])];
  if (p.type === 'bubble') return [['X', 'Y A', '크기 A', 'Y B', '크기 B'], ...[1, 2, 4, 7, 11, 16].map((v) => [v, v * 2, 10 + v, v + 5, 30 + v * 3])];
  if (p.type === 'waterfall') return [['항목', '금액'], ['시작', 10], ['증가', 8], ['감소', -4], ['합계', 14], ['추가', 3], ['총계', 17]];
  const rows = [['분류', '매출', '수량', '이익'], ['가', 10, 20, 3], ['나', 6, 8, 2], ['다', 18, 12, 5], ['라', 9, 16, 4], ['마', 15, 25, 6], ['바', 12, 14, 3]];
  if (['pie', 'pieOfPie', 'barOfPie', 'funnel', 'pareto'].includes(p.type)) return rows.map((r) => r.slice(0, 2));
  return rows;
}

function snapshot(wb, chart) {
  const d = chartModelData(wb, 0, chart);
  return { categories: d.categories, catLevels: d.catLevels ?? null, series: d.series.map((s) => ({ name: s.name, values: s.values, x: s.x ?? null, size: s.size ?? null, type: s.type, axis: s.axis })) };
}
function inspect(bytes, fixture) {
  const wb = new Workbook(readXlsx(bytes).data), chart = wb.sheets[0].charts[0];
  assert.ok(chart, '차트가 유지되어야 합니다.');
  assert.equal(chart.type, fixture.chart.type, '차트 종류');
  const p = fixture.chart;
  assert.equal(!!chart.threeD, !!p.threeD, '3차원 변형');
  if (p.grouping) assert.equal(chart.grouping ?? 'clustered', p.grouping, '누적/100% 변형');
  if (p.explode) assert.equal(chart.explode, p.explode, '쪼개진 원형/도넛 분리량');
  if (p.type === 'scatter') assert.equal(chart.scatterStyle ?? 'marker', p.scatterStyle ?? 'marker', '분산형 선/곡선/표식');
  if (p.type === 'radar') assert.equal(chart.radarStyle ?? 'standard', p.radarStyle ?? 'standard', '방사형 변형');
  if (p.type === 'surface') assert.equal(chart.surfaceStyle, p.surfaceStyle, '표면/등고선/골격 변형');
  if (p.type === 'stock') { assert.equal(!!chart.ohlc, !!p.ohlc, '주식 시가'); assert.equal(!!chart.volume, !!p.volume, '주식 거래량'); }
  const got = snapshot(wb, chart), expected = fixture.data;
  assert.deepEqual(got.series.map((s) => s.values), expected.series.map((s) => s.values), '계열 수치');
  assert.deepEqual(got.series.map((s) => s.name), expected.series.map((s) => s.name), '계열 이름');
  if (!['histogram', 'scatter', 'bubble'].includes(chart.type)) assert.deepEqual(got.categories, expected.categories, '범주');
  if (expected.catLevels) assert.deepEqual(got.catLevels, expected.catLevels, '계층 범주');
  if (['scatter', 'bubble'].includes(chart.type)) assert.deepEqual(got.series.map((s) => s.x), expected.series.map((s) => s.x), 'X 수치');
  if (chart.type === 'bubble') assert.deepEqual(got.series.map((s) => s.size), expected.series.map((s) => s.size), '거품 크기');
  if (chart.type === 'combo') for (const key of ['type', 'axis']) assert.deepEqual(got.series.map((s) => s[key]), expected.series.map((s) => s[key]), `콤보 ${key}`);
  const svg = renderChartSvg(chart, chartModelData(wb, 0, chart));
  assert.doesNotMatch(svg, /NaN|Infinity/, '유한 렌더 기하'); assert.ok(svg.includes('<svg'), '실제 SVG');
  return { type: chart.type, series: got.series.length };
}

let failures = 0;
if (verifyRoot) {
  const manifest = JSON.parse(readFileSync(join(out, 'synthetic-fixtures.json'), 'utf8')), results = [];
  assert.equal(manifest.format, 'wixel-excel-interop'); assert.equal(manifest.version, 1);
  for (const fixture of manifest.fixtures) {
    try { results.push({ file: fixture.file, label: fixture.label, ok: true, ...inspect(readFileSync(join(verifyRoot, fixture.file)), fixture) }); }
    catch (e) { failures++; results.push({ file: fixture.file, label: fixture.label, ok: false, error: e.message }); }
  }
  writeFileSync(join(out, 'gallery-reread-results.json'), JSON.stringify({ total: results.length, failures, results }, null, 2));
  for (const result of results) console.log(JSON.stringify(result));
  console.log(JSON.stringify({ mode: 'verify', total: results.length, failures }));
} else {
  mkdirSync(out, { recursive: true }); const fixtures = [], baseline = [];
  for (const [group, entries] of CHART_GALLERY) for (const [label, options] of entries) {
    const index = fixtures.length + 1, file = `gallery-${String(index).padStart(2, '0')}.xlsx`, rows = inputRows(options), wb = new Workbook();
    wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
    const chart = { id: `gallery${index}`, title: label, x: 20, y: 30, w: 640, h: 420, range: { r1: 0, c1: 0, r2: rows.length - 1, c2: rows[0].length - 1 }, ...options };
    wb.sheets[0].charts = [chart]; const data = snapshot(wb, chart), type = excelType(options);
    assert.ok(Number.isInteger(type), `${label}: XlChartType 기대값 누락`);
    const fixture = { file, group, label, chart: options, data, chartTypes: [type], seriesCounts: [data.series.length + (options.type === 'pareto' ? 1 : 0)], shape: { width: 480, height: 315 } };
    if (options.type === 'combo') { fixture.seriesTypes = [data.series.map((s) => s.type === 'line' ? 65 : 51)]; fixture.axisGroups = [data.series.map((s) => s.axis + 1)]; }
    const bytes = writeXlsx(wb); writeFileSync(join(out, file), bytes); fixtures.push(fixture);
    try { baseline.push({ file, label, ok: true, ...inspect(bytes, fixture) }); }
    catch (e) { failures++; baseline.push({ file, label, ok: false, error: e.message }); }
  }
  writeFileSync(join(out, 'synthetic-fixtures.json'), JSON.stringify({ format: 'wixel-excel-interop', version: 1, enumSource, combinationSource: 'https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.constants', fixtures }, null, 2));
  writeFileSync(join(out, 'gallery-baseline-results.json'), JSON.stringify({ total: baseline.length, failures, results: baseline }, null, 2));
  for (const item of baseline.filter((r) => !r.ok)) console.error(JSON.stringify(item));
  console.log(JSON.stringify({ mode: 'generate', path: resolve(out), total: fixtures.length, failures }));
}
if (failures) process.exitCode = 1;
