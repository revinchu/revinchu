import { test } from 'node:test';
import assert from 'node:assert/strict';
import { REPORT_CELL_STYLE_SECTIONS } from '../src/cell-style-presets.js';
import { cellStylePatch, importCellStyleList } from '../src/cell-style.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { codeOfStyle } from '../src/format.js';

const styles = REPORT_CELL_STYLE_SECTIONS.flatMap(([, list]) => list);
const luminance = hex => {
  const rgb = hex.match(/[a-f0-9]{2}/gi).map(n => parseInt(n, 16) / 255).map(v => v <= .04045 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4);
  return .2126 * rgb[0] + .7152 * rgb[1] + .0722 * rgb[2];
};
test('보고서 스타일 24개는 고유한 이름·RGB 색과 읽을 수 있는 글자 대비를 가짐', () => {
  assert.equal(REPORT_CELL_STYLE_SECTIONS.length, 4); assert.equal(styles.length, 24);
  assert.equal(new Set(styles.map(s => s.name)).size, 24);
  for (const s of styles) {
    assert.match(s.style.fill, /^#[0-9a-f]{6}$/); assert.match(s.style.color, /^#[0-9a-f]{6}$/);
    const a = luminance(s.style.fill), b = luminance(s.style.color);
    assert.ok((Math.max(a, b) + .05) / (Math.min(a, b) + .05) >= 4.5, s.name);
    const patch = cellStylePatch(s);
    for (const key of ['numFmt', 'code', 'decimals', 'queryFormat', 'align', 'wrap', 'locked', 'hideFormula']) assert.equal(key in patch, false, `${s.name}: ${key}`);
  }
});
test('새 프리셋 적용·실행 취소·다시 실행은 값/숫자 형식/맞춤/보호를 유지', () => {
  for (const def of styles) {
    const wb = new Workbook();
    wb.transact(() => { wb.setInput(0, 0, 0, '0.175'); wb.setStyle(0, 0, 0, { numFmt: 'custom', code: '0.000%', align: 'right', locked: false, hideFormula: true }); });
    wb.transact(() => { wb.setCellStyles([def]); wb.setStyle(0, 0, 0, cellStylePatch(def)); });
    const st = wb.styleAt(0, 0, 0);
    assert.equal(st.cellStyleName, def.name); assert.equal(st.fill, def.style.fill); assert.equal(st.code, '0.000%');
    assert.equal(st.align, 'right'); assert.equal(st.locked, false); assert.equal(st.hideFormula, true); assert.equal(wb.getValue(0, 0, 0), .175);
    wb.undo(); assert.equal(wb.styleAt(0, 0, 0).cellStyleName, undefined); assert.equal(wb.styleAt(0, 0, 0).code, '0.000%');
    wb.redo(); assert.equal(wb.styleAt(0, 0, 0).cellStyleName, def.name); assert.equal(wb.getValue(0, 0, 0), .175);
  }
});
test('24개 이름·포함 항목·직접 숫자 형식은 XLSX 저장 후 재열기와 병합 입력에서 보존', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setCellStyles(styles);
    styles.forEach((def, r) => { wb.setInput(0, r, 0, '0.175'); wb.setStyle(0, r, 0, { numFmt: 'custom', code: '0.000%', align: 'right', locked: false, ...cellStylePatch(def) }); });
  });
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(back.cellStyles.map(s => s.name), styles.map(s => s.name));
  const merged = importCellStyleList(back.cellStyles); assert.equal(merged.length, 24);
  styles.forEach((def, r) => {
    const st = back.styleAt(0, r, 0), named = merged[r];
    assert.equal(st.cellStyleName, def.name); assert.equal(st.fill, def.style.fill); assert.equal(st.color, def.style.color);
    assert.equal(codeOfStyle(st), '0.000%'); assert.equal(st.align, 'right'); assert.equal(st.locked, false); assert.equal(back.getValue(0, r, 0), .175);
    assert.equal(named.include.number, false); assert.equal(named.include.alignment, false); assert.equal(named.include.protection, false);
  });
});
