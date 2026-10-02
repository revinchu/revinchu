import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizePhonetic, phoneticText, phoneticHtml } from '../src/phonetic.js';
import { Workbook, cellData } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { ColBuilder } from '../src/block.js';

const meta = { runs: [{ sb: 0, eb: 2, text: 'とうきょう' }, { sb: 3, eb: 5, text: 'かんこく' }], visible: true, type: 'Hiragana', alignment: 'center', font: { font: '맑은 고딕', size: 8, color: '#336699', bold: true } };
const fixture = () => new Workbook({ sheets: [{ name: '합성', cells: { '0,0': { raw: '東京 韓國', phonetic: meta }, '0,1': { raw: '=PHONETIC(A1)' }, '1,0': { raw: '東京 韓國' } } }] });

test('윗주 구간은 원문과 분리하여 읽고 표시와 XML 안전성을 지킨다', () => {
  const p = normalizePhonetic(meta, '東京 韓國', true); p.runs[0].text = 'changed'; assert.equal(meta.runs[0].text, 'とうきょう');
  assert.equal(phoneticText('東京 韓國', meta), 'とうきょう かんこく');
  assert.equal(phoneticHtml('東京 韓國', { ...meta, visible: false }), null);
  assert.match(phoneticHtml('東京 韓國', meta), /phonetic-guide/);
  const safe = phoneticHtml('<a>', { runs: [{ sb: 0, eb: 3, text: '<img onerror="x">' }], visible: true });
  assert.doesNotMatch(safe, /<img|<a>/); assert.match(safe, /&lt;img/);
  assert.throws(() => normalizePhonetic({ runs: [{ sb: 0, eb: 3, text: '가' }, { sb: 2, eb: 4, text: '나' }] }, 'abcd', true), /겹치지/);
});
test('윗주 편집·표시·실행 취소와 직렬화 및 PHONETIC 의존 결과가 함께 보존된다', () => {
  const wb = fixture(); assert.equal(wb.getValue(0, 0, 1), 'とうきょう かんこく');
  const updated = structuredClone(meta); updated.runs[0].text = '도쿄';
  wb.transact(() => wb.setPhonetic(0, 0, 0, updated));
  assert.equal(wb.getValue(0, 0, 0), '東京 韓國'); assert.equal(wb.getValue(0, 0, 1), '도쿄 かんこく');
  wb.undo(); assert.equal(wb.getValue(0, 0, 1), 'とうきょう かんこく'); wb.redo();
  wb.transact(() => wb.setPhoneticVisible(0, 0, 0, false)); assert.equal(wb.getCell(0, 0, 0).phonetic.visible, false); assert.equal(wb.getValue(0, 0, 1), '도쿄 かんこく');
  const restored = new Workbook(wb.serialize()); assert.deepEqual(restored.getCell(0, 0, 0).phonetic, wb.getCell(0, 0, 0).phonetic);
  const copied = cellData(wb.getCell(0, 0, 0)); copied.phonetic.runs[0].text = '외부'; assert.equal(wb.getCell(0, 0, 0).phonetic.runs[0].text, '도쿄');
  wb.transact(() => wb.setStyle(0, 0, 0, { bold: true })); assert.equal(wb.getCell(0, 0, 0).phonetic.runs[0].text, '도쿄');
  wb.transact(() => wb.setInput(0, 0, 0, '새 원문')); assert.equal(wb.getCell(0, 0, 0).phonetic, undefined); assert.equal(wb.getValue(0, 0, 1), '새 원문');
  assert.throws(() => wb.setPhonetic(0, 0, 1, meta), /문자열/);
});
test('XLSX 표준 inlineStr rPh/phoneticPr/c@ph 왕복은 같은 원문·다른 윗주를 구별한다', () => {
  const wb = fixture(), bytes = writeXlsx(wb), files = unzip(bytes), xml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.match(xml, /t="inlineStr" ph="1"/); assert.match(xml, /<rPh sb="0" eb="2">/); assert.match(xml, /phoneticPr fontId="\d+" type="Hiragana" alignment="center"/);
  const round = new Workbook(readXlsx(bytes).data); assert.deepEqual(round.getCell(0, 0, 0).phonetic, meta); assert.equal(round.getCell(0, 1, 0).phonetic, undefined); assert.equal(round.getValue(0, 0, 0), '東京 韓國');
  round.transact(() => round.setPhoneticVisible(0, 0, 0, false)); const twice = new Workbook(readXlsx(writeXlsx(round)).data); assert.equal(twice.getCell(0, 0, 0).phonetic.visible, false); assert.equal(twice.getValue(0, 0, 1), 'とうきょう かんこく');
});
test('외부 sharedStrings와 rich runs의 윗주는 표시 생략 기본 false와 전용 글꼴을 읽는다', () => {
  const wb = new Workbook({ sheets: [{ name: '합성', cells: { '0,0': { raw: '東京' } } }] }), files = unzip(writeXlsx(wb));
  files['xl/sharedStrings.xml'] = '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><t>東</t></r><r><t>京</t></r><rPh sb="0" eb="2"><t>とうきょう</t></rPh><phoneticPr fontId="0" type="Hiragana" alignment="distributed"/></si></sst>';
  const read = new Workbook(readXlsx(zip(files)).data), cell = read.getCell(0, 0, 0);
  assert.equal(cell.v, '東京'); assert.equal(cell.phonetic.visible, false); assert.equal(cell.phonetic.runs[0].text, 'とうきょう'); assert.equal(cell.phonetic.alignment, 'distributed'); assert.ok(cell.phonetic.font.font);
  assert.equal(new Workbook(readXlsx(writeXlsx(read)).data).getCell(0, 0, 0).phonetic.runs[0].text, 'とうきょう');
});
test('윗주 원문 편집은 숫자나 수식처럼 보이는 텍스트도 문자로 보존한다', () => {
  const wb = fixture(); wb.transact(() => wb.setPhonetic(0, 0, 0, { runs: [{ sb: 0, eb: 4, text: '수식 아님' }], visible: true }, '=1+1'));
  assert.equal(wb.getValue(0, 0, 0), '=1+1'); assert.equal(wb.getCell(0, 0, 0).formula, undefined);
  const rt = new Workbook(readXlsx(writeXlsx(wb)).data); assert.equal(rt.getValue(0, 0, 0), '=1+1'); assert.equal(rt.getCell(0, 0, 0).phonetic.runs[0].text, '수식 아님');
  wb.transact(() => wb.clearRange(0, 0, 0, 0, 0)); assert.equal(wb.getCell(0, 0, 0)?.phonetic, undefined); wb.undo(); assert.equal(wb.getValue(0, 0, 0), '=1+1');
});
test('블록 문자열의 윗주는 별도 셀로 보존되고 행 이동·Undo에도 구간이 유지된다', () => {
  const col = new ColBuilder(); col.set(0, '東京 韓國'); col.set(1, '다음');
  const wb = new Workbook({ sheets: [{ name: '합성', cells: {}, blocks: [{ r0: 0, c0: 0, n: 2, ver: 0, cols: [col.finish(2)] }] }] });
  wb.transact(() => wb.setPhonetic(0, 0, 0, meta));
  assert.equal(wb.sheets[0].cells.size, 1); assert.deepEqual(wb.getCell(0, 0, 0).phonetic, meta); assert.equal(wb.getValue(0, 1, 0), '다음');
  wb.transact(() => wb.shiftAxis(0, 'row', 0, 1)); assert.equal(wb.getCell(0, 1, 0).phonetic.runs[0].text, 'とうきょう'); wb.undo();
  assert.deepEqual(new Workbook(wb.serialize()).getCell(0, 0, 0).phonetic, meta);
  assert.equal(new Workbook(readXlsx(writeXlsx(wb)).data).getCell(0, 0, 0).phonetic.runs[1].text, 'かんこく');
  wb.undo(); assert.equal(wb.getCell(0, 0, 0).phonetic, undefined); assert.equal(wb.getValue(0, 0, 0), '東京 韓國');
});
test('윗주 이스케이프 문자열은 한 번만 해석하며 서로게이트 문자 중간을 나누지 않는다', () => {
  const wb = fixture(), hint = '_x000A_ & <한글>\r';
  wb.transact(() => wb.setPhonetic(0, 0, 0, { runs: [{ sb: 0, eb: 2, text: hint }], visible: true }));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data); assert.equal(back.getCell(0, 0, 0).phonetic.runs[0].text, hint);
  assert.throws(() => normalizePhonetic({ runs: [{ sb: 0, eb: 1, text: '중간' }] }, '😀', true), /구간/);
  wb.transact(() => { wb.setInput(0, 1, 0, '={"가";"나"}'); }); wb.getValue(0, 1, 0);
  assert.throws(() => wb.setPhonetic(0, 2, 0, meta), /문자열/);
});
