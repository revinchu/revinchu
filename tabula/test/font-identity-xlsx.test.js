import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { writeOds, readOds } from '../src/ods.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { fontDesktopStyle } from '../src/font-identity.js';

const desktop = (font, style = {}) => ({ ...style, ...fontDesktopStyle(font, style) });
const fontNode = n => ({ font: child(n, 'name')?.attrs.val, bold: child(n, 'b') ? !['0', 'false'].includes(child(n, 'b').attrs.val) : false });
const textFont = n => ({ font: child(n, 'ea')?.attrs.typeface, bold: n.attrs.b === '1' });
const expected = (font, style = {}) => { const d = desktop(font, style); return { font: d.font, bold: !!d.bold }; };
const filesOf = wb => unzip(writeXlsx(wb));
const xmlText = value => typeof value === 'string' ? value : textOf(value);
const xml = (files, name) => parseXml(xmlText(files[name]));
const stripExtensions = files => { for (const key of Object.keys(files)) if (/^xl\/(?:charts|drawings)\/[^/]+\.xml$/.test(key)) files[key] = textOf(files[key]).replace(/<(?:\w+:)?extLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?extLst>/g, ''); return files; };

test('설치용 글꼴 이름은 기본·셀·행·열·이름서식·윗주 XLSX에 기록하고 원본은 유지한다', () => {
  const wb = new Workbook(); wb.defaultFont = { name: 'G마켓 산스', size: 11 };
  for (let r = 0; r < 4; r++) wb.setInput(0, r, 0, r ? String(r) : '위셀');
  wb.setStyle(0, 1, 0, { bold: true }); wb.setStyle(0, 2, 0, { font: '프리텐다드', bold: false, italic: true });
  const unknown = 'Custom "A&B" <font>'; wb.setStyle(0, 3, 0, { font: unknown, bold: true });
  wb.sheets[0].rowStyles[5] = { font: 'G마켓 산스', bold: true };
  wb.sheets[0].colStyles[2] = { font: '프리텐다드', bold: false };
  wb.cellStyles = [{ name: '제목', style: { font: 'G마켓 산스', bold: true } }];
  wb.sheets[0].cells.getRC(0, 0).phonetic = { visible: true, runs: [{ sb: 0, eb: 2, text: 'Wixel' }], font: { font: '프리텐다드', size: 7 } };
  const before = wb.serialize(), files = filesOf(wb), styles = xml(files, 'xl/styles.xml');
  const fonts = kids(child(styles, 'fonts'), 'font').map(fontNode), sheet = xml(files, 'xl/worksheets/sheet1.xml');
  const xfs = kids(child(styles, 'cellXfs'), 'xf'), styleFont = id => fonts[Number(xfs[Number(id)].attrs.fontId)];
  assert.deepEqual(fonts[0], expected('G마켓 산스'));
  const c = descendants(sheet, 'c');
  assert.deepEqual(styleFont(c.find(n => n.attrs.r === 'A2').attrs.s), expected('G마켓 산스', { bold: true }));
  assert.deepEqual(styleFont(c.find(n => n.attrs.r === 'A3').attrs.s), expected('프리텐다드', { bold: false }));
  assert.deepEqual(styleFont(c.find(n => n.attrs.r === 'A4').attrs.s), { font: unknown, bold: true });
  assert.deepEqual(styleFont(descendants(sheet, 'row').find(n => n.attrs.r === '6').attrs.s), expected('G마켓 산스', { bold: true }));
  assert.deepEqual(styleFont(descendants(sheet, 'col').find(n => n.attrs.min === '3').attrs.style), expected('프리텐다드', { bold: false }));
  const named = kids(child(styles, 'cellStyleXfs'), 'xf')[1]; assert.deepEqual(fonts[Number(named.attrs.fontId)], expected('G마켓 산스', { bold: true }));
  assert.equal(fonts[Number(descendants(sheet, 'phoneticPr')[0].attrs.fontId)].font, 'Pretendard');
  const reread = new Workbook(readXlsx(zip(files)).data);
  assert.equal(reread.defaultFont.name, expected('G마켓 산스').font);
  assert.equal(reread.styleAt(0, 1, 0).font, expected('G마켓 산스', { bold: true }).font);
  assert.equal(reread.styleAt(0, 3, 0).font, unknown);
  assert.deepEqual(wb.serialize(), before);
});

test('조건부·표·피벗·슬라이서 DXF의 이름과 명시적 굵게 해제를 함께 보존한다', () => {
  const wb = new Workbook(); wb.setInput(0, 0, 0, '1');
  const style = { font: 'G마켓 산스', bold: true, size: 13, italic: false };
  wb.sheets[0].cond = [{ r1: 0, c1: 0, r2: 0, c2: 0, type: 'gt', v1: '0', style }];
  wb.setObjectStyles({ tables: [{ name: '표 별칭', table: true, pivot: false, elements: [{ type: 'wholeTable', style }] }, { name: '피벗 별칭', table: false, pivot: true, elements: [{ type: 'wholeTable', style: { font: '프리텐다드', bold: false } }] }], slicers: [{ name: '슬라이서 별칭', elements: [{ type: 'wholeTable', style }, { type: 'selectedItemWithData', style }] }] });
  const before = wb.serialize(), files = filesOf(wb), styles = xml(files, 'xl/styles.xml');
  const dxfs = descendants(styles, 'dxf').map(n => fontNode(child(n, 'font')));
  assert.equal(dxfs.length, 5); assert.equal(dxfs.filter(n => n.font === expected(style.font, style).font).length, 4);
  for (const dxf of dxfs) assert.equal(dxf.bold, false);
  const data = readXlsx(zip(files)).data;
  assert.equal(data.sheets[0].cond[0].style.font, expected(style.font, style).font);
  assert.equal(data.objectStyles.slicers[0].elements[1].style.font, expected(style.font, style).font);
  assert.deepEqual(wb.serialize(), before);
});

test('미편집 원본 DXF는 별칭을 고쳐도 미지원 서식과 XML 속성을 잃지 않는다', () => {
  const wb = new Workbook(); wb.setInput(0, 0, 0, '1');
  wb.setObjectStyles({ tables: [{ name: '원본', table: true, elements: [{ type: 'wholeTable', style: { font: 'Arial', bold: true } }] }] });
  const files = filesOf(wb); files['xl/styles.xml'] = textOf(files['xl/styles.xml']).replace('<name val="Arial"/>', '<name val="G마켓 산스"/>').replace('</dxf>', '<protection locked="0"/><alignment readingOrder="2"/></dxf>');
  const imported = new Workbook(readXlsx(zip(files)).data), before = imported.serialize();
  const out = filesOf(imported), dxf = descendants(xml(out, 'xl/styles.xml'), 'dxf')[0];
  assert.deepEqual(fontNode(child(dxf, 'font')), expected('G마켓 산스', { bold: true }));
  assert.equal(child(dxf, 'protection').attrs.locked, '0'); assert.equal(child(dxf, 'alignment').attrs.readingOrder, '2');
  assert.deepEqual(imported.serialize(), before);
});

test('도형 부분 서식과 빈 문단은 글꼴 이름·상속 굵기를 표준 DrawingML로 기록한다', () => {
  const wb = new Workbook(); wb.sheets[0].shapes = [{ id: 'font', kind: 'textbox', x: 0, y: 0, w: 200, h: 80, text: '굵게보통\n', font: 'G마켓 산스', bold: true, paras: [{ runs: [{ t: '굵게' }, { t: '보통', b: false }] }, { runs: [] }] }];
  const before = wb.serialize(), files = stripExtensions(filesOf(wb)), drawing = xml(files, 'xl/drawings/drawing1.xml');
  const runs = descendants(drawing, 'rPr');
  assert.deepEqual(textFont(runs[0]), expected('G마켓 산스', { bold: true }));
  assert.deepEqual(textFont(runs[1]), expected('G마켓 산스', { bold: false }));
  assert.deepEqual(textFont(descendants(drawing, 'endParaRPr')[0]), expected('G마켓 산스', { bold: true }));
  const shape = readXlsx(zip(files)).data.sheets[0].shapes[0]; assert.equal(shape.paras[0].runs[1].font, expected('G마켓 산스').font);
  assert.deepEqual(wb.serialize(), before);
});

for (const type of ['column', 'sunburst']) test(`${type}: 차트·제목·범례·축·레이블이 확장 없이 설치용 글꼴로 왕복된다`, () => {
  const wb = new Workbook(); for (const [r, row] of [['항목', '값'], ['가', 10], ['나', 20]].entries()) for (const [c, v] of row.entries()) wb.setInput(0, r, c, String(v));
  const ch = { id: 'font', type, x: 0, y: 0, w: 400, h: 260, range: { r1: 0, c1: 0, r2: 2, c2: 1 }, title: '글꼴', font: 'G마켓 산스', bold: true, titleFont: '프리텐다드', titleBold: false, legend: 'b', legendFont: 'G마켓 산스', legendBold: false, axes: { x: { font: '프리텐다드', bold: false, title: '축', titleFont: 'G마켓 산스', titleBold: true }, y: { font: 'G마켓 산스', bold: true } }, labels: true, seriesFmt: [{ labelFont: 'G마켓 산스', labelBold: true, pointLabelStyles: { 1: { font: '프리텐다드', bold: false } } }] };
  wb.sheets[0].charts = [ch]; const before = wb.serialize(), files = stripExtensions(filesOf(wb));
  const saved = xmlText(files['xl/charts/chart1.xml']); assert.doesNotMatch(saved, /typeface="(?:G마켓 산스|프리텐다드)"/);
  const result = readXlsx(zip(files)).data.sheets[0].charts[0];
  assert.equal(result.font, expected(ch.font, ch).font); assert.equal(result.bold, false); assert.equal(result.titleFont, 'Pretendard'); assert.equal(result.titleBold, false);
  assert.equal(result.legendFont, expected('G마켓 산스').font);
  assert.equal(result.seriesFmt[0].labelFont, expected('G마켓 산스', { bold: true }).font); assert.equal(result.seriesFmt[0].pointLabelStyles[1].font, 'Pretendard');
  if (type === 'column') { assert.equal(result.axes.x.font, 'Pretendard'); assert.equal(result.axes.x.titleFont, expected('G마켓 산스', { bold: true }).font); }
  assert.deepEqual(wb.serialize(), before);
});

test('테마와 머리글의 알려진 별칭만 바꾸고 필드 코드·미지 글꼴은 그대로 둔다', () => {
  const wb = new Workbook(); wb.setInput(0, 0, 0, '1');
  wb.themeFonts = { name: '사용자', major: 'G마켓 산스', minor: '프리텐다드', majorLatin: 'UserFont', minorLatin: 'Pretendard' };
  const header = '&L&"G마켓 산스,Bold Italic"표제 &P &R&&"G마켓 산스" &"UserFont,Bold"보존';
  wb.sheets[0].page = { header, footer: '&C&"프리텐다드,Regular"&N' };
  const before = wb.serialize(), out = readXlsx(writeXlsx(wb)).data;
  assert.equal(out.themeFonts.major, expected('G마켓 산스').font); assert.equal(out.themeFonts.minor, 'Pretendard'); assert.equal(out.themeFonts.majorLatin, 'UserFont');
  assert.equal(out.sheets[0].page.header, '&L&"' + expected('G마켓 산스', { bold: true }).font + ',Italic"표제 &P &R&&"G마켓 산스" &"UserFont,Bold"보존');
  assert.equal(out.sheets[0].page.footer, '&C&"Pretendard,Regular"&N'); assert.deepEqual(wb.serialize(), before);
});

test('ODS도 동일한 설치용 글꼴 이름과 분리된 굵기 가족을 기록한다', () => {
  const wb = new Workbook(); wb.setInput(0, 0, 0, '값'); wb.setStyle(0, 0, 0, { font: 'G마켓 산스', bold: true });
  const before = wb.serialize();
  const bytes = writeOds([{ si: 0, name: 'Sheet1' }], { raw: (s, r, c) => wb.getRaw(s, r, c), value: (s, r, c) => wb.getValue(s, r, c), style: (s, r, c) => wb.styleAt(s, r, c), used: s => wb.usedRange(s), colWidth: (s, c) => wb.colWidth(s, c), rowHeight: () => null });
  const actual = new Workbook(readOds(bytes).data).styleAt(0, 0, 0);
  assert.equal(actual.font, expected('G마켓 산스', { bold: true }).font); assert.equal(!!actual.bold, false); assert.deepEqual(wb.serialize(), before);
});
