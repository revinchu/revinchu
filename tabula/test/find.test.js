// 찾기 및 바꾸기: 와일드카드 · 대/소문자 · 전체 일치 · 찾는 위치 · 순서 · 통합 문서 · 서식
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findRegex, replaceText, findMatches, nextMatch, formatMatches } from '../src/find.js';
import { Workbook } from '../src/workbook.js';

test('와일드카드: * ? ~ 와 대/소문자 · 전체 셀 내용 일치', () => {
  assert.ok(findRegex('검색*광고').test('검색 파워링크 광고'));
  assert.ok(findRegex('a?c').test('xabcx'));
  assert.ok(!findRegex('a?c', { whole: true }).test('xabcx'));
  assert.ok(findRegex('50~*', {}).test('50*2'));
  assert.ok(!findRegex('50~*', {}).test('500'));
  assert.ok(findRegex('ABC').test('abc'));
  assert.ok(!findRegex('ABC', { matchCase: true }).test('abc'));
  assert.equal(replaceText('파워링크 광고 광고', { text: '광고' }, 'AD'), '파워링크 AD AD');
  assert.equal(replaceText('abc', { text: 'ABC', whole: true }, 'X'), 'X');
  assert.equal(replaceText('abcd', { text: 'ABC', whole: true }, 'X'), 'abcd');
});

test('찾는 위치 · 검색 순서 · 서식 · 다음/이전', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 1, '사과');
    wb.setInput(0, 1, 0, '사과나무');
    wb.setInput(0, 2, 2, '=A2&"!"');
    wb.setInput(0, 3, 0, '0.5');
    wb.setStyle(0, 1, 0, { bold: true });
  });
  const rows = findMatches(wb, { text: '사과', sheets: [0] });
  assert.deepEqual(rows.map((m) => [m.r, m.c]), [[0, 1], [1, 0]]);
  const cols = findMatches(wb, { text: '사과', sheets: [0], byCols: true });
  assert.deepEqual(cols.map((m) => [m.r, m.c]), [[1, 0], [0, 1]]);
  // 값에서 찾으면 수식 결과도 찾음
  const vals = findMatches(wb, { text: '사과', sheets: [0], lookIn: 'values' });
  assert.equal(vals.length, 3);
  // 수식에서 찾기: 수식 글자
  assert.equal(findMatches(wb, { text: 'A2', sheets: [0] }).length, 1);
  assert.equal(findMatches(wb, { text: 'A2', sheets: [0], lookIn: 'values' }).length, 0);
  // 서식으로 찾기
  const bold = findMatches(wb, { text: '사과', sheets: [0], format: { bold: true } });
  assert.deepEqual(bold.map((m) => [m.r, m.c]), [[1, 0]]);
  assert.ok(formatMatches({ bold: true, fill: '#FFFF00' }, { fill: '#ffff00' }));
  assert.ok(!formatMatches({}, { bold: true }));
  assert.ok(formatMatches({}, { bold: false }));
  // 다음 / 이전 찾기 (끝에서 처음으로)
  assert.deepEqual(nextMatch(rows, { si: 0, r: 0, c: 1 }), rows[1]);
  assert.deepEqual(nextMatch(rows, { si: 0, r: 1, c: 0 }), rows[0]);
  assert.deepEqual(nextMatch(rows, { si: 0, r: 0, c: 1 }, { back: true }), rows[1]);
});

test('시트 탭 색: xlsx 왕복', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '1'); wb.setSheetProp(0, 'tabColor', '#c00000'); });
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].tabColor, '#c00000');
  // 자동 저장용 통합 문서 속성 (테마 · 기본 서식)
  wb.theme = ['FFFFFF', '000000'];
  wb.baseStyle = { valign: 'middle' };
  const meta = wb.bookMeta();
  assert.deepEqual(meta.theme, ['FFFFFF', '000000']);
  assert.deepEqual(meta.baseStyle, { valign: 'middle' });
});

test('도형: 엑셀 기본 도형 이름 그대로 그리기 · xlsx 왕복 (회전 · 화살표 · 대시)', async () => {
  const { SHAPE_KINDS, newShape, shapeSvg, GEOM } = await import('../src/shapes.js');
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  assert.ok(SHAPE_KINDS.length >= 120);
  for (const k of SHAPE_KINDS) assert.doesNotMatch(shapeSvg(newShape(k.id, { x: 0, y: 0, w: 80, h: 50 })), /NaN|undefined/, k.id);
  const wb = new Workbook();
  const shapes = [
    { ...newShape('star5', { x: 10, y: 10, w: 80, h: 80 }), rot: 30 },
    { ...newShape('flowChartDecision', { x: 100, y: 10, w: 80, h: 60 }), text: '판단' },
    { ...newShape('lineArrow', { x: 10, y: 120, w: 100, h: 40 }), dash: 'dash' },
    newShape('lineDblArrow', { x: 10, y: 170, w: 100, h: 0 }),
  ];
  wb.transact(() => { wb.setInput(0, 0, 0, 'x'); wb.setSheetProp(0, 'shapes', shapes); });
  const back = new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].shapes;
  assert.equal(back[0].kind, 'star5');
  assert.equal(back[0].rot, 30);
  assert.equal(back[1].kind, 'flowChartDecision');
  assert.equal(back[2].kind, 'line');
  assert.equal(back[2].arrow, 'end');
  assert.equal(back[2].dash, 'dash');
  assert.equal(back[3].arrow, 'both');
  assert.ok(GEOM.heart);
});

test('차트: 새 종류(방사형 · 폭포 · 거품형 · 주식형 · 누적 꺾은선) 그리기와 xlsx 왕복', async () => {
  const { renderChartSvg, chartData, CHART_GALLERY, histogramBins } = await import('../src/chart.js');
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const rows = [['월', '매출', '비용'], ['1월', 10, 4], ['2월', 20, 8], ['3월', 15, 6], ['합계', 0, 0]];
  for (const [, list] of CHART_GALLERY) for (const [label, patch] of list) {
    const svg = renderChartSvg({ w: 400, h: 260, title: label, ...patch }, chartData(rows, patch.type === 'combo' ? 'column' : patch.type));
    assert.doesNotMatch(svg, /NaN|undefined/, label);
  }
  const bins = histogramBins([1, 2, 2, 3, 3, 3, 4, 10]);
  assert.equal(bins.reduce((a, b) => a + b.count, 0), 8);
  const wb = new Workbook();
  wb.transact(() => {
    rows.forEach((r, i) => r.forEach((v, c) => wb.setInput(0, i, c, String(v))));
    wb.setSheetProp(0, 'charts', [
      { id: 'a', type: 'radar', radarStyle: 'filled', title: 'R', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, x: 0, y: 0, w: 300, h: 200 },
      { id: 'b', type: 'waterfall', palette: 'modern', title: 'W', range: { r1: 0, c1: 0, r2: 4, c2: 1 }, x: 0, y: 220, w: 300, h: 200 },
      { id: 'c', type: 'line', grouping: 'stacked', gridY: false, title: 'L', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, x: 320, y: 0, w: 300, h: 200 },
    ]);
  });
  const ch = new Workbook(readXlsx(writeXlsx(wb)).data).sheets[0].charts;
  assert.equal(ch[0].type, 'radar');
  assert.equal(ch[0].radarStyle, 'filled');
  assert.equal(ch[1].type, 'waterfall');
  assert.equal(ch[1].palette, 'modern');
  assert.equal(ch[2].grouping, 'stacked');
  assert.equal(ch[2].gridY, false);
});

test('테두리: 대각선 · 선 스타일 · 색 xlsx 왕복', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, 'x');
    wb.setStyle(0, 0, 0, { bb: true, bbs: 'double', bbc: '#c00000', dd: true, dds: 'dashed', ddc: '#0070c0' });
  });
  const st = new Workbook(readXlsx(writeXlsx(wb)).data).styleAt(0, 0, 0);
  assert.equal(st.bbs, 'double');
  assert.equal(st.bbc, '#c00000');
  assert.equal(st.dd, true);
  assert.equal(st.dds, 'dashed');
  assert.equal(st.ddc, '#0070c0');
});
