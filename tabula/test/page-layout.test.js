import test from 'node:test';
import assert from 'node:assert/strict';
import { normPage, printScale, pageScalePatch, pageXml, pageFromXml } from '../src/page.js';
import { parseXml, descendants } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { THEME_FONTS, THEME_EFFECTS, readThemeOptions, applyThemeOptionsXml } from '../src/theme-options.js';

test('인쇄 자동너비/높이와 수동 배율 모드 전환·기본 여백·잘못된 숫자 정규화', () => {
  assert.deepEqual(pageScalePatch('fitW', '1'), { fitW: 1 });
  assert.deepEqual(pageScalePatch('fitH', 0), { fitH: 0 });
  assert.deepEqual(pageScalePatch('scale', 75), { scale: 75, fitW: 0, fitH: 0 });
  assert.throws(() => pageScalePatch('scale', 'abc'), /숫자/);
  const pg = normPage({ margins: { left: 1 }, fitW: 1.8, fitH: Infinity, scale: 999 });
  assert.equal(pg.margins.top, 0.75); assert.equal(pg.margins.left, 1); assert.equal(pg.fitW, 1); assert.equal(pg.fitH, 32767); assert.equal(pg.scale, 400);
  assert.equal(printScale({ scale: 75 }, 100, 100), 0.75);
  assert.ok(printScale({ fitW: 1, fitH: 0 }, 1400, 100) < 0.5);
  const raw = pageXml({ fitW: 1, fitH: 0, scale: 75, titleRows: [0, 1] }, x => x);
  const root = parseXml(`<root>${raw.setup}${raw.margins}</root>`);
  const reread = pageFromXml({ pageSetup: descendants(root, 'pageSetup')[0], pageMargins: descendants(root, 'pageMargins')[0], fitToPage: true });
  assert.equal(reread.fitW, 1); assert.equal(reread.fitH, 0);
});

test('통합 문서 테마/문서속성은 한 번 Undo/Redo, 원본 객체 별칭·미지원 수식 캐시 불변', () => {
  const wb = new Workbook({ sheets: [{ name: '시트', cells: { '0,0': { raw: '=UNSUPPORTED_FUNCTION()', cached: 42 } } }] });
  const fonts = structuredClone(THEME_FONTS[4]);
  wb.transact(() => {
    wb.setBookProp('themeFonts', fonts); wb.setBookProp('defaultFont', { name: fonts.minor, size: 11 });
    wb.setBookProp('themeEffects', THEME_EFFECTS[2]); wb.setBookProp('props', { lockStructure: true });
  });
  fonts.minor = '외부 변경'; assert.equal(wb.themeFonts.minor, 'Arial'); assert.equal(wb.getValue(0, 0, 0), 42);
  assert.equal(wb.undoStack.length, 1); wb.undo(); assert.equal(wb.themeFonts, null); assert.equal(wb.props.lockStructure, undefined); assert.equal(wb.getValue(0, 0, 0), 42);
  wb.redo(); assert.equal(wb.themeFonts.minor, 'Arial'); assert.equal(wb.props.lockStructure, true); assert.equal(wb.getValue(0, 0, 0), 42);
  assert.throws(() => wb.setBookProp('sheets', []), /지원하지 않는/);
  const restored = new Workbook(wb.serialize()); assert.deepEqual(restored.themeFonts, wb.themeFonts); assert.deepEqual(restored.themeEffects, wb.themeEffects);
});

test('표준 XLSX 테마의 제목/본문 글꼴·효과·색과 기본 글꼴 왕복', () => {
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '한국어 테마'); wb.setBookProp('themeFonts', THEME_FONTS[5]); wb.setBookProp('defaultFont', { name: 'Arial', size: 11 }); wb.setBookProp('themeEffects', THEME_EFFECTS[2]); });
  const read = readXlsx(writeXlsx(wb)).data;
  assert.equal(read.defaultFont.name, 'Arial'); assert.equal(read.themeFonts.major, 'Georgia'); assert.equal(read.themeFonts.minor, 'Arial');
  assert.deepEqual(read.themeEffects, THEME_EFFECTS[2]);
  assert.match(read.themeXml, /a:outerShdw/); assert.match(read.themeXml, /a:alpha val="26000"/);
  const again = readXlsx(writeXlsx(new Workbook(read))).data; assert.deepEqual(again.themeFonts, read.themeFonts); assert.deepEqual(again.themeEffects, read.themeEffects);
});

test('선택하지 않은 원본 테마 글꼴 메타·다른 스크립트·색·채우기와 선 XML을 보존', () => {
  const original = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="original"><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme><a:fontScheme name="old"><a:majorFont><a:latin typeface="Georgia" panose="123"/><a:ea typeface=""/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Jpan" typeface="游ゴシック"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:font script="Hang" typeface="맑은 고딕"/></a:minorFont></a:fontScheme><a:fmtScheme name="Original"><a:fillStyleLst><a:solidFill/></a:fillStyleLst><a:lnStyleLst><a:ln w="1234"/></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>';
  assert.equal(applyThemeOptionsXml(original, readThemeOptions(original)), original);
  const updated = applyThemeOptionsXml(original, { themeFonts: THEME_FONTS[4], themeEffects: THEME_EFFECTS[0] });
  assert.match(updated, /script="Jpan" typeface="游ゴシック"/); assert.match(updated, /a:accent1><a:srgbClr val="123456"/);
  assert.match(updated, /a:ln w="1234"/); assert.match(updated, /a:fillStyleLst><a:solidFill\//);
  assert.equal(readThemeOptions(updated).themeFonts.minor, 'Arial'); assert.equal(readThemeOptions(updated).themeEffects.id, 'none');
});
