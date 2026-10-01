import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const all = { number: true, alignment: true, font: true, border: true, fill: true, protection: true };
const parts = { font: 'Arial', size: 15, bold: true, italic: true, underline: true, strike: true,
  color: '#123456', fill: '#abcdef', pattern: 'darkGrid', patternColor: '#fedcba',
  bt: true, bts: 'double', btc: '#112233', dd: true, dds: 'dashed', ddc: '#332211', du: true, dus: 'dashed', duc: '#332211',
  numFmt: 'custom', code: '0.0000" kg"', align: 'center', valign: 'bottom', wrap: true,
  indent: 2, rotate: -30, shrink: true, locked: false, hideFormula: true };
const parsedStyles = (bytes) => parseXml(textOf(unzip(bytes)['xl/styles.xml']));
const roundtrip = (wb) => new Workbook(readXlsx(writeXlsx(wb)).data);

test('이름 있는 셀 스타일: 구성요소·숨김·참조를 XLSX 왕복으로 보존', () => {
  const wb = new Workbook();
  wb.cellStyles = [{ name: '내 스타일', style: parts, include: all, builtinId: 20, customBuiltin: true, hidden: true, iLevel: 2 }];
  wb.transact(() => { wb.setInput(0, 0, 0, '12.5'); wb.setStyle(0, 0, 0, { ...parts, cellStyleName: '내 스타일' }); });
  const saved = writeXlsx(wb), back = new Workbook(readXlsx(saved).data), named = back.cellStyles[0];
  assert.equal(named.name, '내 스타일'); assert.deepEqual(named.include, all);
  for (const [key, value] of Object.entries(parts)) assert.deepEqual(named.style[key], value, key);
  assert.equal(named.builtinId, 20); assert.equal(named.customBuiltin, true); assert.equal(named.hidden, true); assert.equal(named.iLevel, 2);
  assert.equal(back.styleAt(0, 0, 0).cellStyleName, '내 스타일');
  assert.equal(back.styleAt(0, 0, 0).fill, parts.fill); assert.equal(back.getValue(0, 0, 0), 12.5);
  const xml = parsedStyles(saved), xfs = kids(child(xml, 'cellXfs'), 'xf');
  const linked = xfs.find(xf => xf.attrs.xfId === '1');
  assert.ok(linked); for (const name of ['Font', 'Fill', 'Border', 'NumberFormat', 'Alignment', 'Protection']) assert.equal(linked.attrs['apply' + name], '0', name);
});

test('셀 스타일 읽기: 부모의 여섯 구성요소 상속과 직접 서식 우선', () => {
  const wb = new Workbook(); for (let c = 0; c < 3; c++) wb.setInput(0, 0, c, '1');
  const files = unzip(writeXlsx(wb));
  files['xl/styles.xml'] = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
   <fonts count="2"><font><name val="Arial"/><sz val="11"/><b val="false"/><i val="0"/><strike val="false"/></font><font><name val="Arial"/><sz val="11"/><b/></font></fonts>
   <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill></fill></fills>
   <borders count="2"><border/><border><bottom style="double"><color rgb="FF445566"/></bottom></border></borders>
   <cellStyleXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="2" borderId="1" numFmtId="4"><alignment horizontal="right" vertical="center"/><protection locked="0" hidden="1"/></xf></cellStyleXfs>
   <cellXfs count="4"><xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="0"/>
    <xf fontId="0" fillId="0" borderId="0" numFmtId="0" xfId="1" applyFont="false" applyFill="0" applyBorder="0" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/>
    <xf fontId="0" xfId="1" applyFont="1"/><xf xfId="1"/></cellXfs>
   <cellStyles count="2"><cellStyle name="Normal" builtinId="0" xfId="0"/><cellStyle name="상속" xfId="1"/></cellStyles>
  </styleSheet>`;
  files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replace(/<c r="([ABC])1"[^>]*>/g, (m, c) => `<c r="${c}1" s="${c.charCodeAt(0) - 64}">`);
  const back = new Workbook(readXlsx(zip(files)).data);
  for (let c = 0; c < 3; c++) {
    const style = back.styleAt(0, 0, c);
    assert.equal(style.cellStyleName, '상속'); assert.equal(style.fill, '#112233'); assert.equal(style.bbs, 'double');
    assert.equal(style.align, 'right'); assert.equal(style.valign, 'middle'); assert.equal(style.locked, false); assert.equal(style.hideFormula, true);
    assert.equal(style.numFmt, 'number'); assert.equal(style.decimals, 2);
    assert.equal(!!style.bold, c !== 1); assert.equal(!!style.italic, false); assert.equal(!!style.strike, false);
  }
  const again = roundtrip(back);
  for (let c = 0; c < 3; c++) assert.deepEqual(again.styleAt(0, 0, c), back.styleAt(0, 0, c));
});

test('셀 스타일 구성요소 선택을 보존하고 직접 고친 서식은 부모보다 우선', () => {
  const wb = new Workbook(), include = { ...all, number: false, border: false, protection: false };
  wb.cellStyles = [{ name: '강조', style: parts, include }];
  wb.setInput(0, 0, 0, '17'); wb.setStyle(0, 0, 0, { bold: false, color: '#00ff00', fill: parts.fill, cellStyleName: '강조', numFmt: 'percent', decimals: 0 });
  const back = roundtrip(wb);
  assert.deepEqual(back.cellStyles[0].include, include);
  const style = back.styleAt(0, 0, 0);
  assert.equal(!!style.bold, false); assert.equal(style.color, '#00ff00'); assert.equal(style.fill, parts.fill);
  assert.equal(style.numFmt, 'percent'); assert.equal(!!style.bt, false); assert.equal(style.locked, undefined);
  assert.equal(style.cellStyleName, '강조');
});

test('셀 스타일 행·열 참조 및 표준 스타일의 기본 서식 보존', () => {
  const wb = new Workbook(); wb.defaultFont = { name: 'Arial', size: 12 };
  wb.baseStyle = { valign: 'middle', fill: '#ddeeff', locked: false };
  wb.cellStyles = [{ name: '강조', style: { bold: true } }];
  wb.sheets[0].rowStyles[1] = { bold: true, cellStyleName: '강조' };
  wb.sheets[0].colStyles[2] = { bold: true, cellStyleName: '강조' };
  wb.setInput(0, 0, 0, '1'); wb.setInput(0, 1, 0, '2');
  const bytes = writeXlsx(wb), back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.sheets[0].rowStyles[1].cellStyleName, '강조'); assert.equal(back.sheets[0].colStyles[2].cellStyleName, '강조');
  assert.equal(back.baseStyle.valign, 'middle'); assert.equal(back.baseStyle.fill, '#ddeeff'); assert.equal(back.baseStyle.locked, false);
  const normal = kids(child(parsedStyles(bytes), 'cellStyleXfs'), 'xf')[0];
  assert.equal(child(normal, 'alignment').attrs.vertical, 'center'); assert.notEqual(normal.attrs.fillId, '0');
  assert.equal(child(normal, 'protection').attrs.locked, '0'); assert.deepEqual(back.defaultFont, wb.defaultFont);
});

test('셀 스타일 이름의 대소문자 중복·예약 이름·OpenXML 문자열을 정규화', () => {
  const wb = new Workbook(), name = '보고서 _x0041_ & "한글"';
  wb.cellStyles = [
    { name: 'Report', style: { bold: true } }, { name: ' report ', style: { italic: true } },
    { name: 'Normal', style: { bold: true } }, { name: '표준', style: { bold: true } }, { name: '   ', style: {} },
    { name, style: {} },
  ];
  wb.setInput(0, 0, 0, '1'); wb.setStyle(0, 0, 0, { bold: true, cellStyleName: 'REPORT' });
  const back = roundtrip(wb);
  assert.deepEqual(back.cellStyles.map(style => style.name), ['Report', name]);
  assert.equal(back.cellStyles[0].style.bold, true); assert.equal(back.cellStyles[0].style.italic, undefined);
  assert.equal(back.styleAt(0, 0, 0).cellStyleName, 'Report');
});

test('셀 스타일 1만 개: 선형 이름 매핑·공통 구성요소 재사용·끝 항목 참조 왕복', () => {
  const wb = new Workbook();
  wb.cellStyles = Array.from({ length: 10000 }, (_, i) => ({ name: `스타일 ${i}`, style: { bold: !!(i % 2), fill: i % 2 ? '#123456' : '#abcdef' } }));
  wb.setInput(0, 0, 0, '1'); wb.setStyle(0, 0, 0, { ...wb.cellStyles[9999].style, cellStyleName: '스타일 9999' });
  const bytes = writeXlsx(wb), xml = parsedStyles(bytes), back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.cellStyles.length, 10000); assert.equal(back.cellStyles[9999].name, '스타일 9999');
  assert.equal(back.styleAt(0, 0, 0).cellStyleName, '스타일 9999');
  assert.equal(child(xml, 'cellStyleXfs').attrs.count, '10001'); assert.ok(kids(child(xml, 'fonts'), 'font').length <= 2);
  assert.ok(kids(child(xml, 'fills'), 'fill').length <= 4); assert.ok(kids(child(xml, 'cellXfs'), 'xf').length <= 2);
});

test('표준 스타일이 있어도 명시적 셀 맞춤·보호 기본값은 부모에서 다시 새지 않음', () => {
  const wb = new Workbook(); wb.baseStyle = { valign: 'middle', wrap: true, locked: false, hideFormula: true, bold: true };
  wb.setInput(0, 0, 0, '1');
  wb.sheets[0].cells.getRC(0, 0).style = { numFmt: 'general' };
  const back = roundtrip(wb), style = back.styleAt(0, 0, 0);
  assert.equal(style.valign, undefined); assert.equal(style.wrap, undefined); assert.equal(style.bold, undefined);
  assert.equal(style.locked, undefined); assert.equal(style.hideFormula, undefined);
});


test('셀 스타일 가져오기: 중복 이름은 첫 정의 유지, 숨김 및 그라데이션 보존', () => {
  const wb = new Workbook();
  wb.cellStyles = [
    { name: '첫째', style: { gradient: { deg: 45, stops: [[0, '#112233'], [1, '#ddeeff']] } } },
    { name: '둘째', style: { bold: true }, hidden: true },
  ];
  wb.setInput(0, 0, 0, '1'); wb.setStyle(0, 0, 0, { bold: true, cellStyleName: '둘째' });
  const files = unzip(writeXlsx(wb));
  files['xl/styles.xml'] = textOf(files['xl/styles.xml']).replace('name="둘째"', 'name=" 첫째 "');
  const back = new Workbook(readXlsx(zip(files)).data);
  assert.equal(back.cellStyles.length, 1); assert.equal(back.cellStyles[0].name, '첫째');
  assert.deepEqual(back.cellStyles[0].style.gradient, wb.cellStyles[0].style.gradient);
  // 같은 이름의 두 번째 스타일 서식은 직접 서식으로 남아 셀 모양을 바꾸지 않습니다.
  assert.equal(back.styleAt(0, 0, 0).cellStyleName, '첫째'); assert.equal(back.styleAt(0, 0, 0).bold, true);
  assert.equal(back.styleAt(0, 0, 0).gradient, undefined);
  const again = roundtrip(back); assert.equal(again.styleAt(0, 0, 0).bold, true); assert.equal(again.styleAt(0, 0, 0).gradient, undefined);
});
