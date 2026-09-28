// 큰 통합 문서용 동작: 파일 계산 결과 재사용, 시트별 다시 계산, 가벼운 실행 취소, 빠른 시트 읽기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { zip } from '../src/zip.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';

const book = () => new Workbook({
  sheets: [
    // 파일에서 연 것처럼: 수식 셀에 엑셀이 저장한 결과(cached) — 일부러 계산 결과와 다르게
    { name: 'raw', fileValues: true, cells: { '0,0': { raw: '10' }, '0,1': { raw: '=A1*2', cached: 99 } } },
    { name: 'dash', fileValues: true, cells: { '0,0': { raw: '=raw!B1+1', cached: 77 }, '1,0': { raw: '5' }, '1,1': { raw: '=A2*3', cached: 55 } } },
    { name: 'other', fileValues: true, cells: { '0,0': { raw: '=1+1', cached: 42 } } },
  ],
});

test('파일 계산 결과는 바뀌기 전까지 그대로 쓰고, 바뀐 시트와 참조하는 시트만 다시 계산', () => {
  const wb = book();
  assert.equal(wb.getValue(0, 0, 1), 99);
  assert.equal(wb.getValue(1, 0, 0), 77);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // dash 만 바꿈 → raw · other 는 그대로
  wb.transact(() => wb.setInput(1, 1, 0, '6'));
  assert.equal(wb.getValue(1, 1, 1), 18);
  assert.equal(wb.getValue(0, 0, 1), 99);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // raw 를 바꾸면 raw 를 참조하는 dash 도 다시 계산
  wb.transact(() => wb.setInput(0, 0, 0, '1'));
  assert.equal(wb.getValue(0, 0, 1), 2);
  assert.equal(wb.getValue(1, 0, 0), 3);
  assert.equal(wb.getValue(2, 0, 0), 42);
  // 실행 취소해도 다시 계산된 값이 맞음
  wb.undo();
  assert.equal(wb.getValue(0, 0, 1), 20);
  assert.equal(wb.getValue(1, 0, 0), 21);
});

test('행 삽입 · 시트 추가 · 이름 바꾸기 실행 취소 (시트 단위 기록)', () => {
  const wb = book();
  wb.transact(() => wb.addSheet('new', 3));
  wb.transact(() => wb.insertRows(0, 0, 1));
  assert.equal(wb.getRaw(0, 1, 0), '10');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B2+1'); // 다른 시트의 참조도 조정
  assert.equal(wb.getValue(1, 0, 0), 21);
  assert.equal(wb.getValue(2, 0, 0), 42); // 참조하지 않는 시트는 파일 값 유지
  wb.undo();
  assert.equal(wb.getRaw(0, 0, 0), '10');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B1+1');
  wb.undo();
  assert.deepEqual(wb.sheets.map((s) => s.name), ['raw', 'dash', 'other']);
  assert.equal(wb.getRaw(0, 0, 0), '10');
  wb.redo();
  wb.redo();
  assert.equal(wb.sheets.length, 4);
  assert.equal(wb.getRaw(0, 1, 0), '10');

  wb.transact(() => wb.renameSheet(0, '원본'));
  assert.equal(wb.getRaw(1, 0, 0), '=원본!B2+1');
  assert.equal(wb.getValue(1, 0, 0), 21);
  wb.undo();
  assert.equal(wb.sheets[0].name, 'raw');
  assert.equal(wb.getRaw(1, 0, 0), '=raw!B2+1');
});

test('이름 정의 변경은 이름을 쓰는 시트만 다시 계산', () => {
  const wb = new Workbook({
    sheets: [
      { name: 'a', fileValues: true, cells: { '0,0': { raw: '=세율*2', cached: 1 } } },
      { name: 'b', fileValues: true, cells: { '0,0': { raw: '=3+4', cached: 8 } } },
    ],
    names: [{ name: '세율', ref: '=10', sheet: null }],
  });
  assert.equal(wb.getValue(1, 0, 0), 8);
  wb.transact(() => wb.setNames([{ name: '세율', ref: '=5' }]));
  assert.equal(wb.getValue(0, 0, 0), 10);
  assert.equal(wb.getValue(1, 0, 0), 8);
  wb.undo();
  assert.equal(wb.getValue(0, 0, 0), 20);
});

/** 최소 xlsx (직접 만든 시트 XML) */
function makeXlsx(sheetXml) {
  return zip({
    '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>가</t></si><si><t>a&lt;b</t></si></sst>',
    'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews><sheetData>${sheetXml}</sheetData></worksheet>`,
  });
}

test('빠른 시트 읽기: 공유 문자열 · 인라인 문자열 · 빈 행 · 공유 수식 · 엔티티 · 눈금선', async () => {
  const bytes = makeXlsx(
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>x&amp;y</t></is></c></row>'
    + '<row r="2" ht="30" customHeight="1"/>'
    + '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" ref="B3:B4" si="0">A3&amp;"원"</f><v>2원</v></c><c r="C3" s="0"/></row>'
    + '<row r="4"><c r="A4"><v>3</v></c><c r="B4" t="str"><f t="shared" si="0"/><v>3원</v></c><c r="C4" t="b"><v>1</v></c><c r="D4" t="e"><v>#N/A</v></c></row>',
  );
  for (const read of [readXlsx, (b) => readXlsxAsync(b)]) {
    const { data } = await read(bytes);
    const s = data.sheets[0];
    const cells = Object.fromEntries(s.cells);
    assert.equal(cells['0,0'].raw, '가');
    assert.equal(cells['0,1'].raw, 'a<b');
    assert.equal(cells['0,2'].raw, 'x&y');
    assert.equal(cells['2,1'].raw, '=A3&"원"');
    assert.equal(cells['3,1'].raw, '=A4&"원"');
    assert.equal(cells['3,1'].cached, '3원');
    assert.equal(cells['3,2'].raw, 'TRUE');
    assert.equal(cells['3,3'].raw, '#N/A');
    assert.equal(s.rowHeights[1], 40);
    assert.equal(s.noGrid, true);
    assert.equal(s.fileValues, true);
    const wb = new Workbook();
    wb.load(data);
    assert.equal(wb.getValue(0, 3, 1), '3원');
    // 눈금선 숨김은 저장해도 유지
    const again = readXlsx(writeXlsx(wb)).data.sheets[0];
    assert.equal(again.noGrid, true);
  }
});

test('피벗 오류 값 표시 옵션 (showError) 과 사용자 지정 피벗 스타일', async () => {
  const { computePivot, resolvePivot } = await import('../src/pivot.js');
  const rows = [['캠페인', '비용', '전환'], ['가', 100, 0], ['나', 50, 5]];
  const def = { rows: ['캠페인'], values: [{ field: 'CPA', agg: 'sum' }], calcFields: [{ name: 'CPA', formula: '비용/전환' }], layout: 'tabular' };
  const cellOf = (d, label) => {
    const res = resolvePivot(rows, d);
    const { grid } = computePivot(res, res.def);
    const row = grid.find((g) => g[0]?.raw === label);
    return row[1];
  };
  assert.equal(cellOf(def, '가').raw, '#DIV/0!');
  assert.equal(cellOf({ ...def, errorCaption: '' }, '가').raw, '');
  assert.equal(cellOf({ ...def, errorCaption: '-' }, '가').raw, "'-");
  assert.equal(cellOf({ ...def, errorCaption: '' }, '나').raw, '10');
  // 사용자 지정 스타일: 머리글 흰 글씨
  const custom = { header: { color: '#ffffff', fill: '#2f5597', bold: true } };
  const res = resolvePivot(rows, { ...def, style: '내 스타일', styleDef: custom });
  const head = computePivot(res, res.def).grid.flat().find((c) => c?.role?.startsWith('valueHead') || c?.role === 'rowHead');
  assert.equal(head.style.color, '#ffffff');
  assert.equal(head.style.fill, '#2f5597');
});
