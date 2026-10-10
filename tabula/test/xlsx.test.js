import { test } from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync } from 'node:zlib';
import { inflate, zip, unzip, crc32 } from '../src/zip.js';
import { parseXml, allText } from '../src/xml.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { fmtFromCode } from '../src/format.js';
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

test('그림·도형·데이터 유효성 검사·매크로 왕복', async () => {
  const { readFileSync } = await import('node:fs');
  const { extractVbaModules, fromBase64 } = await import('../src/vba.js');
  const bin = new Uint8Array(readFileSync(new URL('./fixtures/vbaProject.bin', import.meta.url)));
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '사과');
  wb.setInput(0, 1, 0, '배');
  const s = wb.sheets[0];
  s.images.push({ id: 'im1', name: '로고', x: 100, y: 50, w: 64, h: 32, src: `data:image/png;base64,${png}` });
  s.shapes.push({ id: 'sh1', kind: 'roundRect', x: 200, y: 100, w: 150, h: 60, fill: '#4472c4', stroke: '#2f528f', text: '안녕\n하세요', color: '#ffffff', align: 'center', bold: true, size: 14 });
  s.shapes.push({ id: 'sh2', kind: 'textbox', x: 20, y: 300, w: 120, h: 40, fill: null, stroke: '#000000', text: '메모' });
  s.shapes.push({ id: 'sh3', kind: 'line', x: 10, y: 10, w: 100, h: 0, stroke: '#ff0000' });
  s.validations.push({ r1: 0, c1: 2, r2: 9, c2: 2, type: 'list', f1: '예,아니오', errorStyle: 'stop', error: '목록에서 고르세요' });
  s.validations.push({ r1: 0, c1: 3, r2: 9, c2: 3, type: 'whole', op: 'between', f1: '1', f2: '10', prompt: '1~10' });
  s.validations.push({ r1: 0, c1: 4, r2: 9, c2: 4, type: 'list', f1: '=$A$1:$A$2' });
  wb.vba = { bin: (await import('../src/vba.js')).toBase64(bin), codeName: 'ThisWorkbook', sheetCodes: { Sheet1: 'Sheet1' } };

  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  assert.ok(files['xl/vbaProject.bin']);
  assert.match(new TextDecoder().decode(files['[Content_Types].xml']), /macroEnabled\.main\+xml/);
  assert.match(new TextDecoder().decode(files['[Content_Types].xml']), /Extension="png"/);
  const { data, warnings } = readXlsx(bytes);
  assert.deepEqual(warnings, []);
  const wb2 = new Workbook(data);
  const s2 = wb2.sheets[0];
  assert.equal(s2.images.length, 1);
  assert.equal(s2.images[0].name, '로고');
  assert.equal(s2.images[0].src, `data:image/png;base64,${png}`);
  assert.deepEqual([s2.images[0].x, s2.images[0].y, s2.images[0].w, s2.images[0].h], [100, 50, 64, 32]);
  assert.equal(s2.shapes.length, 3);
  const [a, b, c] = s2.shapes;
  assert.equal(a.kind, 'roundRect');
  assert.equal(a.text, '안녕\n하세요');
  assert.equal(a.fill, '#4472c4');
  assert.equal(a.stroke, '#2f528f');
  assert.equal(a.bold, true);
  assert.equal(a.size, 14);
  assert.deepEqual([a.x, a.y, a.w, a.h], [200, 100, 150, 60]);
  assert.equal(b.kind, 'textbox');
  assert.equal(b.fill, null);
  assert.equal(c.kind, 'line');
  assert.equal(c.stroke, '#ff0000');
  assert.equal(s2.validations.length, 3);
  const [v1, v2, v3] = s2.validations;
  assert.equal(v1.type, 'list');
  assert.equal(v1.f1, '"예,아니오"');
  assert.equal(v1.error, '목록에서 고르세요');
  assert.deepEqual([v1.r1, v1.c1, v1.r2, v1.c2], [0, 2, 9, 2]);
  assert.equal(v2.type, 'whole');
  assert.equal(v2.f2, '10');
  assert.equal(v2.prompt, '1~10');
  assert.equal(v3.f1, '$A$1:$A$2');
  assert.ok(wb2.vba?.bin);
  const mods = extractVbaModules(fromBase64(wb2.vba.bin));
  assert.ok(mods.some((m) => m.name === 'Module1' && m.code.includes('MsgBox')));
});

test('데이터 유효성 검사 판정', async () => {
  const { checkValidation, listItems, validationAt } = await import('../src/validation.js');
  const wb = new Workbook();
  wb.setInput(0, 0, 0, '서울');
  wb.setInput(0, 1, 0, '부산');
  const s = wb.sheets[0];
  const list = { r1: 0, c1: 1, r2: 5, c2: 1, type: 'list', f1: '$A$1:$A$2' };
  const lit = { r1: 0, c1: 2, r2: 5, c2: 2, type: 'list', f1: '"예,아니오"' };
  const whole = { r1: 0, c1: 3, r2: 5, c2: 3, type: 'whole', op: 'between', f1: '1', f2: '10' };
  const len = { r1: 0, c1: 4, r2: 5, c2: 4, type: 'textLength', op: 'lessThanOrEqual', f1: '3' };
  const custom = { r1: 0, c1: 5, r2: 5, c2: 5, type: 'custom', f1: 'F1>A1', allowBlank: true };
  s.validations.push(list, lit, whole, len);
  assert.equal(validationAt(s, 3, 1), list);
  assert.equal(validationAt(s, 6, 1), null);
  assert.deepEqual(listItems(wb, 0, list), ['서울', '부산']);
  assert.deepEqual(listItems(wb, 0, lit), ['예', '아니오']);
  assert.equal(checkValidation(wb, 0, list, 0, 1, '부산'), true);
  assert.equal(checkValidation(wb, 0, list, 0, 1, '대구'), false);
  assert.equal(checkValidation(wb, 0, lit, 0, 2, '예'), true);
  assert.equal(checkValidation(wb, 0, whole, 0, 3, '5'), true);
  assert.equal(checkValidation(wb, 0, whole, 0, 3, '5.5'), false);
  assert.equal(checkValidation(wb, 0, whole, 0, 3, '11'), false);
  assert.equal(checkValidation(wb, 0, whole, 0, 3, ''), true);
  assert.equal(checkValidation(wb, 0, len, 0, 4, '가나다'), true);
  assert.equal(checkValidation(wb, 0, len, 0, 4, '가나다라'), false);
  wb.setInput(0, 0, 0, '5');
  assert.equal(checkValidation(wb, 0, custom, 0, 5, '7'), true);
  assert.equal(checkValidation(wb, 0, custom, 0, 5, '3'), false);
  assert.equal(wb.getValue(0, 0, 5), null);
});

test('유효성 검사 범위 빼기 · 잘못된 데이터', async () => {
  const { subtractRange, invalidCells } = await import('../src/validation.js');
  const rule = { r1: 0, c1: 0, r2: 9, c2: 3, type: 'whole', op: 'greaterThan', f1: '0' };
  const parts = subtractRange(rule, { r1: 2, c1: 1, r2: 4, c2: 2 });
  const count = parts.reduce((n, p) => n + (p.r2 - p.r1 + 1) * (p.c2 - p.c1 + 1), 0);
  assert.equal(count, 40 - 6);
  assert.deepEqual(subtractRange(rule, { r1: 20, c1: 0, r2: 30, c2: 0 }), [rule]);
  assert.deepEqual(subtractRange(rule, { r1: 0, c1: 0, r2: 100, c2: 100 }), []);
  const wb = new Workbook();
  wb.sheets[0].validations.push(rule);
  wb.setInput(0, 0, 0, '5');
  wb.setInput(0, 1, 0, '-1');
  wb.setInput(0, 2, 0, '=A1-10');
  wb.setInput(0, 20, 0, '-5');
  assert.deepEqual(invalidCells(wb, 0), [{ r: 1, c: 0 }, { r: 2, c: 0 }]);
});
