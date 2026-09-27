import { test } from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync } from 'node:zlib';
import { inflate, zip, unzip, crc32 } from '../src/zip.js';
import { parseXml, allText } from '../src/xml.js';
import { readXlsx, writeXlsx, fmtFromCode } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
import { Axis } from '../src/axis.js';
import { chartData, renderChartSvg, niceScale } from '../src/chart.js';
import { buildPivot } from '../src/pivot.js';

test('inflate: zlib 로 압축한 데이터 복원', () => {
  const text = `${'가나다라마바사 '.repeat(3000)}${JSON.stringify([...Array(3000)].map((_, i) => i * 7))}`;
  const src = new TextEncoder().encode(text);
  for (const level of [0, 1, 9]) {
    assert.equal(new TextDecoder().decode(inflate(deflateRawSync(src, { level }))), text);
  }
});

test('zip 쓰기/읽기와 CRC', () => {
  const z = zip({ 'a.txt': 'hello', 'dir/b.xml': '<x>한글</x>' });
  const files = unzip(z);
  assert.equal(new TextDecoder().decode(files['dir/b.xml']), '<x>한글</x>');
  assert.equal(crc32(new TextEncoder().encode('hello')), 0x3610a686);
});

test('XML 파서', () => {
  const root = parseXml('<?xml version="1.0"?><x:a b="1 &amp; 2"><x:si><x:r><x:t>안</x:t></x:r><x:r><x:t xml:space="preserve"> 녕</x:t></x:r></x:si><c/><![CDATA[<raw>]]></x:a>');
  assert.equal(root.name, 'a');
  assert.equal(root.attrs.b, '1 & 2');
  assert.equal(allText(root.children[0]), '안 녕');
  assert.equal(root.text, '<raw>');
});

test('표시 형식 코드 해석', () => {
  assert.deepEqual(fmtFromCode('#,##0'), { numFmt: 'comma' });
  assert.deepEqual(fmtFromCode('#,##0.00'), { numFmt: 'number', decimals: 2 });
  assert.deepEqual(fmtFromCode('0.0%'), { numFmt: 'percent', decimals: 1 });
  assert.deepEqual(fmtFromCode('yyyy-mm-dd'), { numFmt: 'date' });
  assert.deepEqual(fmtFromCode('h:mm:ss'), { numFmt: 'time' });
  assert.deepEqual(fmtFromCode('[$₩-412]#,##0'), { numFmt: 'currency', decimals: undefined });
  assert.deepEqual(fmtFromCode('@'), { numFmt: 'text' });
  assert.deepEqual(fmtFromCode('0.00E+00'), { numFmt: 'scientific', decimals: 2 });
});

function sampleBook() {
  const wb = new Workbook();
  wb.transact(() => {
    const rows = [['품목', '1월', '2월'], ['사과', '100', '120'], ['배', '80', '95'], ['포도', '60', '70']];
    rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v)));
    wb.setInput(0, 4, 0, '합계');
    wb.setInput(0, 4, 1, '=SUM(B2:B4)');
    wb.setInput(0, 4, 2, '=IFS(C4>50,"큼",TRUE,"작음")');
    wb.setInput(0, 5, 0, '2024-03-15');
    wb.setInput(0, 5, 1, '15%');
    wb.setInput(0, 5, 2, "'007");
    wb.setInput(0, 6, 0, 'TRUE');
    wb.setStyle(0, 0, 0, { bold: true, fill: '#217346', color: '#ffffff', align: 'center', bb: true });
    wb.setStyle(0, 1, 1, { numFmt: 'comma' });
    wb.setComment(0, 1, 0, '빨간 사과');
    wb.setColWidth(0, 0, 120);
    wb.setRowHeight(0, 2, 30);
    const s = wb.sheets[0];
    s.merges.push({ r1: 7, c1: 0, r2: 7, c2: 2 });
    s.cond.push({ r1: 1, c1: 1, r2: 3, c2: 1, type: 'gt', v1: '90', style: { fill: '#ffc7ce', color: '#9c0006' } });
    s.cond.push({ r1: 1, c1: 2, r2: 3, c2: 2, type: 'bar', color: '#638ec6' });
    s.freeze = { rows: 1, cols: 1 };
    s.hiddenCols[5] = true;
    s.charts.push({ id: 'c1', type: 'column', title: '월별 판매', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, x: 300, y: 20, w: 400, h: 250 });
    wb.addSheet('데이터 2');
    wb.setInput(1, 0, 0, "='Sheet1'!B5*2");
  });
  return wb;
}

test('xlsx 저장 → 다시 열기 (왕복)', () => {
  const wb = sampleBook();
  const bytes = writeXlsx(wb);
  const { data, warnings } = readXlsx(bytes);
  assert.deepEqual(warnings, []);
  const wb2 = new Workbook(data);
  const v = (si, r, c) => wb2.getValue(si, r, c);
  assert.equal(wb2.sheets.length, 2);
  assert.equal(wb2.sheets[1].name, '데이터 2');
  assert.equal(v(0, 4, 1), 240);
  assert.equal(wb2.getRaw(0, 4, 2), '=IFS(C4>50,"큼",TRUE,"작음")');
  assert.equal(v(0, 4, 2), '큼');
  assert.equal(v(1, 0, 0), 480);
  assert.equal(v(0, 5, 0), 45366);
  assert.equal(wb2.getRaw(0, 5, 0), '2024-03-15');
  assert.equal(wb2.styleAt(0, 5, 0).numFmt, 'date');
  assert.equal(v(0, 5, 1), 0.15);
  assert.equal(wb2.styleAt(0, 5, 1).numFmt, 'percent');
  assert.equal(v(0, 5, 2), '007');
  assert.equal(v(0, 6, 0), true);
  const h = wb2.styleAt(0, 0, 0);
  assert.equal(h.bold, true);
  assert.equal(h.fill, '#217346');
  assert.equal(h.color, '#ffffff');
  assert.equal(h.align, 'center');
  assert.equal(h.bb, true);
  assert.equal(wb2.styleAt(0, 1, 1).numFmt, 'comma');
  assert.equal(wb2.getCell(0, 1, 0).comment, '빨간 사과');
  const s = wb2.sheets[0];
  assert.equal(s.colWidths[0], 120);
  assert.equal(s.rowHeights[2], 30);
  assert.deepEqual(s.merges, [{ r1: 7, c1: 0, r2: 7, c2: 2 }]);
  assert.deepEqual(s.freeze, { rows: 1, cols: 1 });
  assert.equal(s.hiddenCols[5], true);
  assert.equal(s.cond.length, 2);
  assert.equal(s.cond[0].type, 'gt');
  assert.equal(s.cond[0].v1, '90');
  assert.equal(s.cond[0].style.fill, '#ffc7ce');
  assert.equal(s.cond[1].type, 'bar');
  assert.equal(s.charts.length, 1);
  assert.equal(s.charts[0].type, 'column');
  assert.equal(s.charts[0].title, '월별 판매');
  assert.deepEqual(s.charts[0].range, { r1: 0, c1: 0, r2: 3, c2: 2 });
  assert.ok(Math.abs(s.charts[0].x - 300) <= 1 && Math.abs(s.charts[0].w - 400) <= 2);
});

test('긴 참조 사슬 계산 (스택 넘침 없음)', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, '1');
    for (let r = 1; r < 20000; r++) wb.sheets[0].cells.set(`${r},0`, { raw: `=A${r}+1`, formula: true, ast: { type: 'bin', op: '+', a: { type: 'ref', ref: { sheet: null, r1: r - 1, c1: 0, r2: r - 1, c2: 0, range: false } }, b: { type: 'num', v: 1 } } });
  });
  assert.equal(wb.getValue(0, 19999, 0), 20000);
});

test('Axis: 사용자 지정 크기와 숨김', () => {
  const ax = new Axis(20, { 2: 40, 5: 0 }, [{ 3: true }], 1048576);
  assert.equal(ax.pos(0), 0);
  assert.equal(ax.pos(3), 80);
  assert.equal(ax.pos(4), 80);
  assert.equal(ax.pos(6), 100);
  assert.equal(ax.indexAt(79), 2);
  assert.equal(ax.indexAt(80), 4);
  assert.equal(ax.indexAt(99.5), 4);
  assert.equal(ax.indexAt(100), 6);
  assert.equal(ax.pos(1048576), 1048576 * 20 + 20 - 20 - 20);
  assert.equal(ax.indexAt(1e12), 1048575);
});

test('차트 데이터 해석과 SVG', () => {
  const rows = [[null, '1월', '2월'], ['사과', 100, 120], ['배', 80, 95]];
  const d = chartData(rows, 'column');
  assert.deepEqual(d.categories, ['사과', '배']);
  assert.deepEqual(d.series.map((s) => s.name), ['1월', '2월']);
  assert.deepEqual(d.series[1].values, [120, 95]);
  const svg = renderChartSvg({ type: 'column', title: '판매', w: 400, h: 250 }, d);
  assert.ok(svg.startsWith('<svg') && svg.includes('판매') && (svg.match(/<rect /g) ?? []).length >= 5);
  const mixed = chartData([['지역', '분기', '매출'], ['서울', 'Q1', 100], ['부산', 'Q2', 50]], 'column');
  assert.deepEqual(mixed.series.map((x) => x.name), ['매출']);
  const pie = renderChartSvg({ type: 'pie', w: 300, h: 200 }, chartData([['a', 1], ['b', 3]], 'pie'));
  assert.ok(pie.includes('75%'));
  assert.deepEqual(niceScale(0, 97), { min: 0, max: 100, step: 20 });
});

test('피벗 테이블', () => {
  const rows = [['지역', '분기', '매출'], ['서울', 'Q1', 100], ['부산', 'Q1', 50], ['서울', 'Q2', 30], ['서울', 'Q1', 20]];
  const p = buildPivot(rows, { rowField: 0, colField: 1, valueField: 2, agg: 'sum' });
  const raw = p.map((r) => r.map((c) => c?.raw ?? ''));
  assert.deepEqual(raw, [
    ['합계 : 매출', '열 레이블', '', ''],
    ['행 레이블', 'Q1', 'Q2', '총합계'],
    ['부산', '50', '', '50'],
    ['서울', '120', '30', '150'],
    ['총합계', '170', '30', '200'],
  ]);
  const cnt = buildPivot(rows, { rowField: 0, colField: null, valueField: 2, agg: 'count' });
  assert.deepEqual(cnt.map((r) => r.map((c) => c.raw)), [['행 레이블', '개수 : 매출'], ['부산', '1'], ['서울', '3'], ['총합계', '4']]);
});
