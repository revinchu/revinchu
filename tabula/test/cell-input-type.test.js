import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';

const values = (wb, n) => Array.from({ length: n }, (_, r) => wb.getValue(0, r, 0));
const strings = ['00123', '=1+1', 'TRUE', '#N/A', '1/1', "'앞 작은따옴표"];
const textBook = () => new Workbook({ sheets: [{ name: 'S', cells: Object.fromEntries(strings.map((raw, r) => [`${r},0`, { raw, link: 'https://example.com', style: { numFmt: 'text', bold: true } }])) }] });

test('서식 지우기는 텍스트의 원본 raw·자료형을 실행 취소와 JSON 저장 후에도 유지한다', () => {
  const wb = textBook();
  wb.transact(() => wb.clearRange(0, 0, 0, strings.length - 1, 0, 'formats'));
  assert.deepEqual(values(wb, strings.length), strings);
  assert.equal(wb.getCell(0, 1, 0).formula, undefined, '문자 수식을 실행하지 않음');
  for (let r = 0; r < strings.length; r++) {
    assert.equal(wb.getCell(0, r, 0).raw, strings[r]);
    assert.equal(wb.getCell(0, r, 0).style, undefined);
  }
  wb.undo(); assert.deepEqual(values(wb, strings.length), strings);
  wb.redo(); assert.deepEqual(values(wb, strings.length), strings);
  assert.deepEqual(values(new Workbook(JSON.parse(JSON.stringify(wb.serialize()))), strings.length), strings);
  wb.transact(() => wb.setComment(0, 1, 0, '메모'));
  wb.undo(); wb.redo(); assert.equal(wb.getValue(0, 1, 0), '=1+1');
});

test('링크 제거의 표준 서식 교체도 문자 내용을 보존하고 새 입력만 다시 해석한다', () => {
  const wb = textBook();
  wb.transact(() => {
    for (let r = 0; r < strings.length; r++) {
      const data = cellData(wb.getCell(0, r, 0), { numFmt: 'general', bold: false });
      delete data.link; wb.setCellData(0, r, 0, data);
    }
  });
  assert.deepEqual(values(wb, strings.length), strings);
  wb.undo(); assert.equal(wb.getCell(0, 0, 0).link, 'https://example.com');
  wb.redo(); assert.equal(wb.getCell(0, 0, 0).link, undefined);
  wb.transact(() => wb.setInput(0, 0, 0, '00123'));
  assert.equal(wb.getValue(0, 0, 0), 123);
  wb.transact(() => wb.setInput(0, 1, 0, '=1+1'));
  assert.equal(wb.getValue(0, 1, 0), 2);
});

test('텍스트 표시 형식은 숫자·논리·오류·인용 문자의 자료형을 바꾸지 않는다', () => {
  const wb = new Workbook({ sheets: [{ name: 'S', cells: {
    '0,0': { raw: '123' }, '1,0': { raw: 'TRUE' }, '2,0': { raw: '#N/A' }, '3,0': { raw: "'00123" },
  } }] });
  const before = values(wb, 4);
  wb.transact(() => { for (let r = 0; r < 4; r++) wb.setStyle(0, r, 0, { numFmt: 'text' }); });
  assert.deepEqual(values(wb, 4), before);
  wb.undo(); assert.deepEqual(values(wb, 4), before); wb.redo(); assert.deepEqual(values(wb, 4), before);
  assert.deepEqual(values(new Workbook(wb.serialize()), 4), before);
  wb.transact(() => wb.setLineStyle(0, 'col', 0, { numFmt: 'general' }));
  assert.deepEqual(values(wb, 4), before);
  wb.transact(() => wb.setInput(0, 0, 0, '123'));
  assert.equal(wb.getValue(0, 0, 0), '123', '직접 다시 입력하면 현재 텍스트 서식을 따름');
  wb.transact(() => wb.clearRange(0, 0, 0, 3, 0, 'contents'));
  assert.deepEqual(values(wb, 4), [null, null, null, null]);
});

test('블록의 텍스트 표시 형식을 가진 숫자를 일반 셀로 꺼내도 숫자로 유지한다', () => {
  const wb = new Workbook();
  wb.sheets[0].blocks.push({ r0: 0, c0: 0, n: 2, ver: 0, cols: [{ num: new Float64Array([123, 7]), str: null, dict: [], fmt: { numFmt: 'text' } }] });
  wb.transact(() => wb.setComment(0, 0, 0, '메모'));
  assert.equal(wb.getValue(0, 0, 0), 123);
  wb.transact(() => wb.setStyle(0, 1, 0, { bold: true }));
  assert.equal(wb.getValue(0, 1, 0), 7);
  const saved = new Workbook(wb.serialize());
  assert.equal(saved.getValue(0, 0, 0), 123); assert.equal(saved.getValue(0, 1, 0), 7);
});

test('표준 XLSX 자료형은 @ 서식과 독립적이며 숫자·논리·문자 수식의 왕복을 유지한다', () => {
  const base = new Workbook();
  base.transact(() => base.setStyle(0, 0, 0, { numFmt: 'text' }));
  const files = unzip(writeXlsx(base));
  const xml = textOf(files['xl/worksheets/sheet1.xml']);
  const style = xml.match(/<c\b[^>]*\bs="(\d+)"/)[1];
  // 앱 자체 확장/힌트 없이 OOXML의 t 속성만으로 원본 파일을 구성한다.
  const cells = [
    '<c r="A1" t="n"><v>123</v></c>', '<c r="A2" t="b"><v>1</v></c>', '<c r="A3" t="e"><v>#N/A</v></c>',
    '<c r="A4" t="inlineStr"><is><t>00123</t></is></c>', '<c r="A5" t="inlineStr"><is><t>=1+1</t></is></c>', '<c r="A6" t="n"><v>0</v></c>',
  ];
  files['xl/worksheets/sheet1.xml'] = xml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${cells.map((c, i) => `<row r="${i + 1}">${c.replace('<c ', `<c s="${style}" `)}</row>`).join('')}</sheetData>`);
  const wb = new Workbook(readXlsx(zip(files)).data);
  assert.equal(wb.getValue(0, 0, 0), 123); assert.equal(wb.getValue(0, 1, 0), true); assert.equal(wb.getValue(0, 2, 0).code, '#N/A');
  assert.equal(wb.getValue(0, 3, 0), '00123'); assert.equal(wb.getValue(0, 4, 0), '=1+1'); assert.equal(wb.getValue(0, 5, 0), 0);
  const exported = writeXlsx(wb), sheet = textOf(unzip(exported)['xl/worksheets/sheet1.xml']);
  assert.match(sheet, /<c\b[^>]*r="A1"[^>]*><v>123<\/v><\/c>/);
  assert.match(sheet, /<c\b[^>]*r="A2"[^>]*t="b"[^>]*><v>1<\/v><\/c>/);
  assert.match(sheet, /<c\b[^>]*r="A4"[^>]*t="s"/);
  assert.doesNotMatch(sheet, /<f\b/);
  assert.deepEqual(values(new Workbook(readXlsx(exported).data), 6), values(wb, 6));
  wb.transact(() => wb.clearRange(0, 0, 0, 5, 0, 'formats'));
  assert.deepEqual(values(new Workbook(readXlsx(writeXlsx(wb)).data), 6), values(wb, 6));
});
