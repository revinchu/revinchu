import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { tableCellStyle } from '../src/tables.js';
import { fmtCode } from '../src/format.js';
import { tableCellDisplayStyle } from '../src/table-format.js';

const URI = '{3720B142-2CF8-4B6E-95A7-574958454C54}';
const NS = 'https://wixel.app/table-style/1';
const roundtrip = wb => new Workbook(readXlsx(writeXlsx(wb)).data);
function fixture() {
  return new Workbook({ sheets: [{ name: '자료', fileValues: true,
    allStyle: { italic: true, color: '#990099' },
    colStyles: { 0: { fill: '#ffdd00', bold: true, numFmt: 'custom', code: '0.00', font: 'Arial', size: 13 }, 1: { fill: '#aabbcc', bb: true, bbc: '#cc0000' } },
    rowStyles: { 0: { fill: '#111111', color: '#eeeeee', bold: true }, 2: { fill: '#abcdef' } },
    cells: {
      '0,0': { raw: '구분', style: { fill: '#ff0000', color: '#00ff00', cellStyleName: '기존 강조' } },
      '0,1': { raw: '금액' },
      '1,0': { raw: '001', inputType: 'text', comment: '메모 유지', link: 'https://example.com/table' },
      '1,1': { raw: '=2+2', cached: 42, style: { numFmt: 'custom', code: '#,##0.00', align: 'right', locked: false } },
      '2,0': { raw: '3' }, '2,1': { raw: '5' },
      '4,0': { raw: '표 밖', style: { fill: '#445566' } },
      '5,5': { raw: '외부 수정 색', style: { fill: '#123456', color: '#dd0022', bold: true } },
    },
    tables: [{ id: 't', name: '자료표', r1: 0, c1: 0, r2: 2, c2: 1, header: true, banded: true, style: 'TableStyleMedium2', filter: { criteria: {}, hidden: {} } }],
  }] });
}
function applyStyle(wb) {
  wb.transact(() => { wb.clearTableVisualFormatting(0, 't'); wb.setTableStyle(0, 't', { style: 'TableStyleMedium7' }); });
}
function assertTableColor(wb) {
  const table = wb.sheets[0].tables[0];
  for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) {
    const expected = tableCellStyle(table, r, c), actual = tableCellDisplayStyle(wb, 0, r, c);
    assert.equal(actual.fill, expected.fill, `fill ${r},${c}`);
    assert.equal(actual.color, expected.color, `color ${r},${c}`);
    assert.equal(actual.bold, expected.bold, `bold ${r},${c}`);
    assert.equal(actual.italic, expected.italic, `italic ${r},${c}`);
  }
}

test('빠른 표 스타일은 기존 행/열/전체 색을 XLSX 재열기 후에도 다시 상속하지 않는다', () => {
  const wb = fixture(); applyStyle(wb); assertTableColor(wb);
  const back = roundtrip(wb); assertTableColor(back);
  assertTableColor(roundtrip(back));
  assert.equal(back.sheets[0].tables[0].style, 'TableStyleMedium7');
  assert.equal(back.sheets[0].colStyles[0].fill, '#ffdd00');
  assert.equal(back.sheets[0].rowStyles[0].fill, '#111111');
  assert.equal(back.getCell(0, 4, 0).style.fill, '#445566');
  assert.equal(back.getCell(0, 0, 0).style.cellStyleName, undefined);
});

test('빠른 표 스타일 XLSX 저장은 값/수식캐시/숫자형식/메모/링크/맞춤/보호를 보존한다', () => {
  const wb = fixture(); applyStyle(wb); const back = roundtrip(wb);
  assert.equal(back.getValue(0, 1, 0), '001');
  assert.equal(back.getRaw(0, 1, 1), '=2+2'); assert.equal(back.getValue(0, 1, 1), 42);
  assert.equal(back.getCell(0, 1, 0).comment, '메모 유지');
  assert.equal(back.getCell(0, 1, 0).link, 'https://example.com/table');
  assert.equal(back.styleAt(0, 1, 0).font, 'Arial'); assert.equal(back.styleAt(0, 1, 0).size, 13);
  assert.equal(fmtCode(back.styleAt(0, 1, 0)), '0.00');
  assert.equal(fmtCode(back.styleAt(0, 1, 1)), '#,##0.00');
  assert.equal(back.styleAt(0, 1, 1).align, 'right'); assert.equal(back.styleAt(0, 1, 1).locked, false);
});

test('표 스타일을 바꾼 뒤 직접 고친 굵게/글자색/채우기는 XLSX에서 우선한다', () => {
  const wb = fixture(); applyStyle(wb);
  wb.transact(() => wb.setStyle(0, 1, 0, { bold: true, color: '#5500aa', fill: '#00ffff' }));
  const back = roundtrip(wb), actual = tableCellDisplayStyle(back, 0, 1, 0);
  assert.equal(actual.bold, true); assert.equal(actual.color, '#5500aa'); assert.equal(actual.fill, '#00ffff');
  const marker = back.getCell(0, 1, 0).style.tableStyleInherit;
  assert.equal(marker.bold, undefined); assert.equal(marker.color, undefined); assert.equal(marker.fill, undefined);
  assert.equal(tableCellDisplayStyle(back, 0, 0, 0).fill, tableCellStyle(back.sheets[0].tables[0], 0, 0).fill);
});

test('XLSX native XF는 이전 채우기를 제거하고 tableStyleInfo와 확인란 확장을 함께 보존한다', () => {
  const wb = fixture(); applyStyle(wb); wb.transact(() => wb.setStyle(0, 1, 0, { checkbox: true }));
  const files = unzip(writeXlsx(wb)), styles = parseXml(textOf(files['xl/styles.xml']));
  const xfs = kids(child(styles, 'cellXfs'), 'xf');
  const marked = xfs.filter(x => descendants(x, 'tableStyleInherit').length);
  assert.ok(marked.length >= 2);
  for (const xf of marked) assert.equal(xf.attrs.fillId, '0');
  assert.ok(marked.some(x => descendants(x, 'xfComplement').length));
  assert.match(textOf(files['xl/tables/table1.xml']), /tableStyleInfo name="TableStyleMedium7"/);
  const back = new Workbook(readXlsx(zip(files)).data);
  assert.equal(back.styleAt(0, 1, 0).checkbox, true); assertTableColor(back);
});

test('표 스타일 상속 확장이 없는 XLSX 직접 채우기는 변경하지 않는다', () => {
  const wb = fixture(), bytes = writeXlsx(wb), xml = textOf(unzip(bytes)['xl/styles.xml']);
  assert.ok(!xml.includes('tableStyleInherit'));
  const back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.getCell(0, 0, 0).style.fill, '#ff0000');
  assert.equal(tableCellDisplayStyle(back, 0, 0, 0).fill, '#ff0000');
  assert.equal(back.sheets[0].tables[0].style, 'TableStyleMedium2');
});

test('다른 편집기의 native 채우기 변경이 오래된 상속 확장보다 우선한다', () => {
  const wb = fixture(); applyStyle(wb); const files = unzip(writeXlsx(wb));
  const styles = parseXml(textOf(files['xl/styles.xml'])), xfs = kids(child(styles, 'cellXfs'), 'xf');
  const native = xfs.find(x => x.attrs.fillId !== '0' && x.attrs.fontId !== '0' && !descendants(x, 'tableStyleInherit').length);
  assert.ok(native);
  const xml = textOf(files['xl/styles.xml']).replace(/<xf\b[^>]*?(?:\/>|>[^]*?<\/xf>)/g, xf => xf.includes(`uri="${URI}"`) ? xf.replace(/fillId="\d+"/, `fillId="${native.attrs.fillId}"`) : xf);
  files['xl/styles.xml'] = xml;
  const back = new Workbook(readXlsx(zip(files)).data);
  const style = back.getCell(0, 0, 0).style;
  assert.notEqual(style.fill, ''); assert.notEqual(style.fill, undefined);
  assert.equal(style.tableStyleInherit.fill, undefined);
  assert.equal(tableCellDisplayStyle(back, 0, 0, 0).fill, style.fill);
});

test('잘못된 표 스타일 확장은 무시하고 숫자형식/보호 같은 다른 속성을 주입하지 않는다', () => {
  const wb = fixture(); applyStyle(wb); const files = unzip(writeXlsx(wb));
  const source = textOf(files['xl/styles.xml']);
  const inject = json => source.replace(/(<wx:tableStyleInherit\b[^>]*json=")[^"]*(")/g, `$1${json.replaceAll('"', '&quot;')}$2`);
  files['xl/styles.xml'] = inject('{"fill":"","numFmt":"","locked":false,"bold":false}');
  let back = new Workbook(readXlsx(zip(files)).data), marker = back.getCell(0, 0, 0).style.tableStyleInherit;
  assert.equal(marker.fill, ''); assert.equal(marker.bold, false);
  assert.equal(marker.numFmt, undefined); assert.equal(marker.locked, undefined);
  // Import may add safe neutral masks for row/column defaults; the extension
  // itself must never inject number formats or protection properties.
  files['xl/styles.xml'] = inject('invalid-json');
  back = new Workbook(readXlsx(zip(files)).data); assert.equal(back.getCell(0, 0, 0).style.tableStyleInherit, undefined);
  files['xl/styles.xml'] = source.replaceAll(NS, 'https://example.com/not-wixel');
  back = new Workbook(readXlsx(zip(files)).data); assert.equal(back.getCell(0, 0, 0).style.tableStyleInherit, undefined);
});


test('기본 XF의 흰 채우기/검정 글자는 빠른 스타일을 XLSX 재열기 후 덮지 않는다', () => {
  const wb = new Workbook({ baseStyle: { fill: '#ffffff', color: '#000000', font: 'Arial', size: 13, valign: 'middle' },
    sheets: [{ name: '자료', cells: { '0,0': { raw: '제목', style: { fill: '#ff0000' } }, '1,0': { raw: '본문' } },
      tables: [{ id: 't', name: '표1', r1: 0, c1: 0, r2: 2, c2: 0, header: true, banded: true, style: 'TableStyleMedium2' }] }] });
  applyStyle(wb);
  for (const book of [wb, roundtrip(wb), roundtrip(roundtrip(wb))]) {
    for (const r of [0, 1]) {
      const actual = tableCellDisplayStyle(book, 0, r, 0), expected = tableCellStyle(book.sheets[0].tables[0], r, 0);
      assert.equal(actual.fill, expected.fill); assert.equal(actual.color, expected.color);
      assert.equal(actual.font, 'Arial'); assert.equal(actual.size, 13); assert.equal(actual.valign, 'middle');
    }
    assert.equal(book.getValue(0, 0, 0), '제목'); assert.equal(book.getValue(0, 1, 0), '본문');
  }
});


test('기본 흰색인 표에서 숫자형식만 고쳐도 XLSX에 직접 흰색을 복제하지 않는다', () => {
  const wb = new Workbook({ baseStyle: { fill: '#ffffff', color: '#000000', font: 'Arial', size: 13 },
    sheets: [{ name: '자료', cells: { '0,0': { raw: '제목' }, '1,0': { raw: '0.125' } },
      tables: [{ id: 't', name: '표1', r1: 0, c1: 0, r2: 2, c2: 0, header: true, banded: true, style: 'TableStyleMedium3' }] }] });
  wb.transact(() => wb.setStyle(0, 1, 0, { numFmt: 'percent', decimals: 2 }));
  const back = roundtrip(wb), actual = tableCellDisplayStyle(back, 0, 1, 0);
  assert.equal(actual.fill, tableCellStyle(back.sheets[0].tables[0], 1, 0).fill);
  assert.equal(actual.color, tableCellStyle(back.sheets[0].tables[0], 1, 0).color);
  assert.equal(fmtCode(back.styleAt(0, 1, 0)), '0.00%'); assert.equal(back.getValue(0, 1, 0), 0.125);
});


test('새 행 기본값은 다시 연 표에서 다음 굵게 스타일을 가로막지 않는다', () => {
  const wb = fixture(); applyStyle(wb); const back = roundtrip(wb);
  const t = back.sheets[0].tables[0];
  t.styleElements = [{ type: 'wholeTable', style: { bold: true, italic: true, fill: '#123456', color: '#ffffff' } }];
  const actual = tableCellDisplayStyle(back, 0, 2, 1);
  assert.equal(actual.bold, true); assert.equal(actual.italic, true);
  assert.equal(actual.fill, '#123456'); assert.equal(actual.color, '#ffffff');
  back.transact(() => back.setStyle(0, 2, 1, { bold: false, fill: '#fedcba' }));
  const explicit = tableCellDisplayStyle(back, 0, 2, 1);
  assert.equal(explicit.bold, false); assert.equal(explicit.fill, '#fedcba');
  assert.equal(tableCellDisplayStyle(roundtrip(back), 0, 2, 1).bold, false);
});
