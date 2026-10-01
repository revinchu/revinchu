import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { chartData, renderChartSvg, CHART_GALLERY } from '../src/chart.js';
import { chartView3D, chartDepth } from '../src/chart-3d.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, descendants, child } from '../src/xml.js';
const rows = [['분기', '온라인', '오프라인'], ['1분기', 120, 80], ['2분기', -40, 100], ['3분기', 0, 75], ['4분기', 180, 140]];
test('3D 갤러리: 기존 12개 투영 변형을 유지하고 표면·거품을 별도 지원', () => {
  const variants = CHART_GALLERY.flatMap(([, entries]) => entries);
  assert.equal(variants.filter(([, p]) => p.threeD && ['column', 'bar', 'pie', 'area', 'line'].includes(p.type)).length, 12);
  assert.ok(variants.some(([, p]) => p.threeD && p.type === 'surface'));
  assert.ok(variants.some(([, p]) => p.threeD && p.type === 'bubble'));
  assert.ok(variants.filter(([, p]) => p.type === 'pieOfPie' || p.type === 'barOfPie').every(([, p]) => !p.threeD));
  assert.ok(variants.every(([, p]) => typeof p.threeD === 'boolean'));
});
test('3D 투영: 모든 유형과 누적·음수·영 값이 유효 SVG로 렌더링', () => {
  for (const [, patch] of CHART_GALLERY.flatMap(([, entries]) => entries).filter(([, p]) => p.threeD)) {
    for (const rotY of [0, 30, 180, 270, 360]) {
      const svg = renderChartSvg({ ...patch, w: 500, h: 340, view3D: { rotX: 35, rotY, depthPercent: 150 }, title: '분기 매출' }, chartData(rows, patch.type));
      assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
      assert.match(svg, /data-3d=/);
      assert.equal(parseXml(svg).name, 'svg');
    }
  }
  for (const values of [[0, 0], [1, 0], [null, null]]) {
    const svg = renderChartSvg({ type: 'pie', threeD: true, w: 400, h: 260 }, { categories: ['A', 'B'], series: [{ name: '값', values }] });
    assert.doesNotMatch(svg, /NaN|Infinity/);
  }
});
test('3D 회전·깊이·원근감은 기하를 바꾸고 잘못된 입력은 제한', () => {
  const ch = { type: 'column', threeD: true, w: 500, h: 320 };
  assert.notDeepEqual(chartDepth(ch, 300, 200), chartDepth({ ...ch, view3D: { rotY: 270 } }, 300, 200));
  assert.notDeepEqual(chartDepth(ch, 300, 200), chartDepth({ ...ch, view3D: { depthPercent: 200 } }, 300, 200));
  assert.notDeepEqual(chartDepth(ch, 300, 200), chartDepth({ ...ch, view3D: { rAngAx: false, perspective: 180 } }, 300, 200));
  const v = chartView3D({ view3D: { rotX: Infinity, rotY: -400, depthPercent: -1, perspective: 500 } });
  assert.deepEqual(v, { rotX: 20, rotY: 0, depthPercent: 20, perspective: 240, rAngAx: true });
});
test('3D 차트는 WIXEL 확장 없이도 표준 XLSX 종류·회전·누적·색 복원', () => {
  const wb = new Workbook();
  rows.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
  const types = ['column', 'bar', 'pie', 'area', 'line'];
  const view3D = { rotX: 40, rotY: 70, depthPercent: 160, perspective: 60, rAngAx: false };
  wb.sheets[0].charts = types.map((type, i) => ({ id: 'native' + i, type, threeD: true, view3D, grouping: 'stacked', title: type, x: 20, y: i * 300, w: 440, h: 280, gap: 0, range: { r1: 0, c1: 0, r2: 4, c2: 2 }, seriesFmt: [{ pointColors: { 0: '#ff0000' } }] }));
  const files = unzip(writeXlsx(wb));
  for (let i = 0; i < types.length; i++) {
    const path = 'xl/charts/chart' + (i + 1) + '.xml';
    const xml = textOf(files[path]);
    const root = parseXml(xml);
    const kind = types[i] === 'column' || types[i] === 'bar' ? 'bar3DChart' : types[i] + '3DChart';
    const group = descendants(root, kind)[0];
    assert.ok(group, kind);
    assert.ok(descendants(root, 'view3D').length);
    assert.equal(child(group, 'overlap'), null);
    assert.equal(child(group, 'firstSliceAng'), null);
    files[path] = xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  }
  const back = readXlsx(zip(files)).data.sheets[0].charts;
  for (let i = 0; i < types.length; i++) {
    assert.equal(back[i].type, types[i]);
    assert.equal(back[i].threeD, true);
    assert.deepEqual(back[i].view3D, view3D);
    if (types[i] !== 'pie') assert.equal(back[i].grouping, 'stacked');
  }
  assert.equal(back[0].gap, 0);
  assert.equal(back[2].seriesFmt[0].colors[0].toLowerCase(), '#ff0000');
});


test('100% 누적 차트: 단일 계열과 음수의 값 축은 비율 기준', () => {
  const settings = { type: 'column', threeD: true, grouping: 'percentStacked', w: 440, h: 280 };
  const single = renderChartSvg(settings, chartData([['항목', '값'], ['A', 80], ['B', 120]], 'column'));
  assert.match(single, />100%</);
  assert.doesNotMatch(single, /[0-9],?000%/);
  const negative = renderChartSvg(settings, chartData([['항목', '값1', '값2'], ['A', -10, 30], ['B', 10, -20]], 'column'));
  assert.match(negative, />-100%</);
});
