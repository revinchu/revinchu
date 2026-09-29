// 엑셀 화면 동등성: 기본 셀 서식 · 기본 제공 표시 형식 · 자동 행 높이 · 표/피벗 기본 제공 스타일 · 색 밝기 · 원형 차트 레이블
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zip } from '../src/zip.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { Workbook } from '../src/workbook.js';
import { applyTint, presetStyle, tablePresetCell, PRESET_STYLES } from '../src/stylepresets.js';
import { computePivot } from '../src/pivot.js';
import { renderChartSvg } from '../src/chart.js';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
function miniBook(sheetXml, stylesXml) {
  return zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': `<?xml version="1.0"?><workbook ${NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet ${NS}><sheetData>${sheetXml}</sheetData></worksheet>`,
    'xl/styles.xml': `<?xml version="1.0"?><styleSheet ${NS}>${stylesXml}</styleSheet>`,
  });
}

const STYLES = '<fonts count="2"><font><sz val="9"/><name val="맑은 고딕"/></font><font><sz val="14"/><name val="맑은 고딕"/></font></fonts>'
  + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>'
  + '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>'
  + '<xf numFmtId="40" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
  + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
  + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="45" shrinkToFit="1"/></xf></cellXfs>';

test('기본 셀 서식(xf 0) · 기본 제공 형식 40 · 회전/축소 · 자동 높이 대상 행', () => {
  const bytes = miniBook('<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>-1234.5</v></c></row>'
    + '<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>큰 글자</t></is></c></row>'
    + '<row r="3" ht="12" customHeight="1"><c r="A3" s="2" t="inlineStr"><is><t>고정</t></is></c><c r="B3" s="3" t="inlineStr"><is><t>회전</t></is></c></row>', STYLES);
  const { data } = readXlsx(bytes);
  const wb = new Workbook(data);
  // s 없는 셀도 엑셀의 표준 스타일(세로 가운데)
  assert.equal(wb.styleAt(0, 0, 0).valign, 'middle');
  assert.equal(wb.styleAt(0, 5, 5).valign, 'middle');
  // 형식 40 = #,##0.00_);[Red](#,##0.00)
  assert.equal(wb.styleAt(0, 0, 1).code, '#,##0.00_);[Red](#,##0.00)');
  assert.deepEqual([wb.styleAt(0, 2, 1).rotate, wb.styleAt(0, 2, 1).shrink], [45, true]);
  // 높이가 저장되지 않은 2행만 자동 높이 대상 (3행은 높이 지정)
  assert.deepEqual(wb.fitRows, [[1]]);
  // 서식을 새로 입혀도 기본 맞춤이 유지됨
  wb.transact(() => wb.setStyle(0, 4, 4, { bold: true }));
  assert.deepEqual({ ...wb.getCell(0, 4, 4).style }, { valign: 'middle', bold: true });
  // 저장하면 xf 0 에 기본 맞춤, 회전 · 축소도 그대로
  const styles = textOf(unzip(writeXlsx(wb))['xl/styles.xml']);
  assert.match(styles, /<cellXfs count="\d+"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"\/><\/xf>/);
  assert.match(styles, /textRotation="45" shrinkToFit="1"/);
});

test('엑셀 색 밝기(tint): 윈도 정수 HLS', () => {
  assert.equal(applyTint('FFFFFF', -0.1499984740745262), 'D9D9D9');
  assert.equal(applyTint('4472C4', 0.3999755851924192), '8EA9DB');
  assert.equal(applyTint('4472C4', -0.249977111117893), '305496');
  assert.equal(applyTint('000000', 0.499984740745262), '808080');
});

test('기본 제공 표 · 피벗 스타일 144개 정의', () => {
  assert.equal(Object.keys(PRESET_STYLES).filter((k) => !k.startsWith('Wixel')).length, 144);
  assert.equal(Object.keys(PRESET_STYLES).filter((k) => k.startsWith('Wixel')).length, 40);
  const t = { r1: 0, c1: 0, r2: 4, c2: 2, header: true, banded: true, firstCol: true };
  // TableStyleLight9: 머리글 강조1 + 흰 굵은 글자, 첫 열 굵게, 줄무늬 행 위쪽 선
  const h = tablePresetCell('TableStyleLight9', t, 0, 1);
  assert.deepEqual([h.fill, h.color, h.bold], ['#4472c4', '#ffffff', true]);
  assert.equal(tablePresetCell('TableStyleLight9', t, 2, 0).bold, true);
  assert.equal(tablePresetCell('TableStyleLight9', t, 2, 1).bt, true);
  // 어둡게 스타일은 흰 글자 머리글 + 채운 본문
  assert.ok(tablePresetCell('TableStyleDark2', t, 1, 1).fill);
  assert.ok(presetStyle('PivotStyleLight16').hr.fill);
});

test('피벗: 기본 제공 스타일을 표 모양대로 칠함 (PivotStyleMedium2)', () => {
  const rows = [['지역', '매출'], ['서울', '10'], ['부산', '20']];
  const d = { rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], pages: [], filters: {}, layout: 'compact', grandRows: true, grandCols: true, style: 'PivotStyleMedium2', styleOpts: { rowHeaders: true, colHeaders: true } };
  const { grid } = computePivot(rows, d);
  assert.equal(grid[0][0].style.fill, '#305496'); // 머리글 행: 강조1 25% 어둡게
  assert.equal(grid[0][0].style.color, '#ffffff');
  assert.equal(grid[0][0].style.bold, true); // 첫 머리글 셀 굵게
  const last = grid[grid.length - 1];
  assert.equal(last[0].role, 'grandLabel');
  assert.equal(last[1].style.bts, 'double'); // 총합계 행 위 이중선
  assert.equal(last[1].style.numFmt, 'comma'); // 표시 형식은 셀 것이 우선
});

test('원형 차트 레이블: 파일에 없으면 표시하지 않음, 백분율 지정 · 기본값', () => {
  const data = { categories: ['가', '나'], series: [{ name: 'A', values: [30, 70] }] };
  const svg = (fmt) => renderChartSvg({ type: 'pie', w: 300, h: 200, seriesFmt: [fmt] }, { ...data, series: [{ ...data.series[0], ...fmt }] });
  assert.doesNotMatch(svg({ labels: false }), />70%</);
  assert.match(svg({ pct: true }), />70%</);
  assert.match(svg({}), />70%</);
  assert.match(svg({ labels: true }), />70</);
});

test('차트: 그라데이션 채우기 · 그림자 · 파일의 글꼴 크기', () => {
  const s = { name: 'A', values: [1, 2], type: 'column', grad: { stops: [[0, '#4472c4'], [1, '#2f5597']], ang: 90 }, shadow: true, labels: true, labelSize: 9, labelColor: '#404040' };
  const out = renderChartSvg({ type: 'column', w: 400, h: 300, axisSize: 9, title: 'T', titleSize: 14 }, { categories: ['x', 'y'], series: [s] });
  assert.match(out, /<linearGradient id="[^"]+"/);
  assert.match(out, /fill="url\(#[^)]+\)" filter="url\(#[^)]+\)"/);
  assert.match(out, /font-size="18.7"[^>]*>T</);
  assert.match(out, /font-size="12" fill="#404040"/);
});

test('WIXEL 모던 스타일: 표 · 피벗 · 슬라이서가 xlsx 사용자 지정 스타일로 왕복', async () => {
  const { Workbook } = await import('../src/workbook.js');
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const { unzip, textOf } = await import('../src/zip.js');
  const { tablePresetCell } = await import('../src/stylepresets.js');
  const wb = new Workbook();
  wb.transact(() => {
    [['지역', '매출'], ['서울', '10'], ['부산', '20']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v)));
    wb.setSheetProp(0, 'tables', [{ id: 't1', name: '표1', r1: 0, c1: 0, r2: 2, c2: 1, header: true, totals: false, style: 'WixelTableModern2', banded: true, filter: null, columns: ['지역', '매출'] }]);
    wb.setSheetProp(0, 'slicers', [{ id: 's1', caption: '지역', source: { kind: 'table', table: '표1', column: '지역' }, style: 'WixelSlicerSolid3', x: 200, y: 10, w: 150, h: 150 }]);
  });
  const head = tablePresetCell('WixelTableModern2', { r1: 0, c1: 0, r2: 2, c2: 1, header: true, banded: true }, 0, 0);
  assert.equal(head.fill, '#312e81');
  assert.equal(head.color, '#ffffff');
  const styles = textOf(unzip(writeXlsx(wb))['xl/styles.xml']);
  assert.match(styles, /<tableStyle name="WixelTableModern2" pivot="0"/);
  assert.match(styles, /<x14:slicerStyle name="WixelSlicerSolid3">/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].tables[0].style, 'WixelTableModern2');
  assert.equal(back.sheets[0].slicers[0].style, 'WixelSlicerSolid3');
});

test('콤보 차트: 다단계 항목 축 · 여러 칸 계열 이름 · 파일의 축 배정 · 조건 범위 COUNTIF', async () => {
  const { resolveChart, chartModelData } = await import('../src/chart.js');
  const wb = new Workbook();
  const put = (r, c, v) => wb.setInput(0, r, c, v);
  put(0, 2, '매출'); put(1, 2, '진척');           // 이름 두 칸: 위는 병합 바깥쪽 빈 칸처럼 값 하나만
  put(1, 3, '수익');
  put(2, 0, '1본부'); put(2, 1, '1월'); put(3, 1, '2월'); put(4, 0, '2본부'); put(4, 1, '1월');
  [[2, 10, 100], [3, 20, 200], [4, 30, 300]].forEach(([r, a, b]) => { put(r, 2, String(a)); put(r, 3, String(b)); });
  const ref = (r1, c1, r2, c2) => ({ sheet: 'Sheet1', r1, c1, r2, c2 });
  const ch = {
    type: 'combo', w: 400, h: 300,
    series: [
      { name: { ref: ref(0, 2, 1, 2) }, cat: ref(2, 0, 4, 1), val: ref(2, 2, 4, 2) },
      { name: { ref: ref(0, 3, 1, 3) }, cat: ref(2, 0, 4, 1), val: ref(2, 3, 4, 3) },
    ],
    seriesFmt: [{ type: 'column', axis: 1 }, { type: 'line' }],
  };
  wb.sheets[0].charts = [ch];
  const d = chartModelData(wb, 0, ch);
  assert.deepEqual(d.categories, ['1월', '2월', '1월']);
  assert.deepEqual(d.catLevels, [[{ text: '1본부', start: 0, end: 1 }, { text: '2본부', start: 2, end: 2 }]]);
  assert.deepEqual(d.series.map((s) => s.name), ['매출 진척', '수익']);
  assert.deepEqual(d.series.map((s) => s.axis), [1, 0]); // 마지막 계열도 파일대로 기본 축
  const svg = renderChartSvg(ch, d);
  assert.match(svg, />1본부</);
  assert.match(svg, />2본부</);
  assert.equal(typeof resolveChart, 'function');

  // 조건 자리에 범위 → 조건마다 개수 (고유 개수 공식)
  put(6, 0, 'a'); put(6, 1, 'b'); put(6, 2, 'a');
  put(7, 0, '=SUMPRODUCT((A7:D7<>"")/(1-(A7:D7<>"")+COUNTIF(A7:D7,A7:D7)))');
  assert.equal(wb.getValue(0, 7, 0), 2);
  put(8, 0, '=SUMPRODUCT(COUNTIFS(A7:C7,{"a","b"}))');
  assert.equal(wb.getValue(0, 8, 0), 3);
});
