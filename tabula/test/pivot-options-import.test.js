import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { convertXlsb } from '../src/xlsb.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml } from '../src/xml.js';

// MS-XLSB 2.4.278 BrtBeginSXView: fAutoFormat(bit 8), fSingleFilterPerField(bit 9), fDontUseCustomLists(bit 14).
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/f6cbfd44-775f-44f6-8ba8-fac46d7a445c
const part = 'xl/pivotTables/pivotTable1.xml';
const u32 = (...xs) => { const b = Buffer.alloc(xs.length * 4); xs.forEach((x, i) => b.writeUInt32LE(x >>> 0, i * 4)); return b; };
const wide = s => Buffer.concat([u32(s.length), Buffer.from(s, 'utf16le')]);
const variable = x => { const out = []; do { const n = x & 127; x >>>= 7; out.push(n | (x ? 128 : 0)); } while (x); return Buffer.from(out); };
const record = (id, data) => Buffer.concat([variable(id), variable(data.length), data]);
function binaryOptions(f1 = 6, f2 = 0, f3 = 0) {
  const fixed = Buffer.alloc(32);
  fixed.writeUInt32LE(f1 >>> 0, 0); fixed.writeUInt32LE((f2 | 0x80000) >>> 0, 4); fixed.writeUInt32LE((f3 | 0xc0) >>> 0, 8);
  fixed[12] = 2; fixed[14] = 6; fixed[15] = 3; fixed.writeInt32LE(-1, 16); fixed.writeUInt32LE(1, 28);
  const files = { 'xl/workbook.bin': new Uint8Array(), 'xl/pivotTables/pivotTable1.bin': record(280, Buffer.concat([fixed, wide('옵션피벗'), wide('값')])) };
  for (const _ of convertXlsb(files)) { /* 변환 완료 */ }
  return parseXml(textOf(files['xl/pivotTables/pivotTable1.bin'])).attrs;
}
function fixture(options = {}) {
  return new Workbook({ sheets: [
    { name: '원본', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '값' }, '1,0': { raw: '서울' }, '1,1': { raw: '10' }, '2,0': { raw: '부산' }, '2,1': { raw: '20' } } },
    { name: '보고서', cells: {}, colWidths: { 1: 143, 2: 59 }, pivot: { name: '옵션피벗', source: '원본', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, top: 4, left: 1, rows: ['지역'], cols: [], values: [{ field: '값', agg: 'sum' }], ...options } },
  ] });
}
const attrNames = ['useAutoFormatting', 'preserveFormatting', 'enableDrill', 'showHeaders', 'showDrill', 'mergeItem', 'multipleFieldFilters', 'customListSort'];
function externalOptions(attrs) {
  const files = unzip(writeXlsx(fixture()));
  let xml = textOf(files[part]);
  for (const key of attrNames) xml = xml.replace(new RegExp(` ${key}="[^"]*"`, 'g'), '');
  const attributes = Object.entries(attrs).filter(([key]) => attrNames.includes(key)).map(([key, value]) => ` ${key}="${value}"`).join('');
  files[part] = xml.replace('<pivotTableDefinition ', '<pivotTableDefinition' + attributes + ' ');
  return new Workbook(readXlsx(zip(files)).data);
}
function effective(def) {
  return { autofit: def.autofit !== false, preserveFormat: def.preserveFormat !== false, enableDrill: def.enableDrill !== false,
    showHeaders: def.showHeaders !== false, showExpand: def.showExpand !== false, mergeLabels: !!def.mergeLabels, multiFilters: !!def.multiFilters, customListSort: def.customListSort !== false };
}

for (const autofit of [false, true]) test(`XLSB 자동 맞춤 ${autofit}는 생략 없이 명시하고 인접 비트와 분리`, () => {
  const attrs = binaryOptions(6, (autofit ? 0x100 : 0) | 0x80 | 0x20 | 0x200 | 0x400);
  assert.equal(attrs.useAutoFormatting, autofit ? '1' : '0');
  assert.equal(attrs.preserveFormatting ?? '1', '1'); assert.equal(attrs.enableDrill ?? '1', '1');
  assert.equal(attrs.name, '옵션피벗'); assert.equal(attrs.cacheId, '1');
});
for (const enabled of [false, true]) test(`XLSB 여러 필터·사용자 지정 목록 옵션 ${enabled}의 반전 비트를 보존`, () => {
  const attrs = binaryOptions(6, 0x100, enabled ? 0 : 0x200 | 0x4000);
  assert.equal(attrs.multipleFieldFilters ?? '0', enabled ? '1' : '0');
  assert.equal(attrs.customListSort ?? '1', enabled ? '1' : '0');
  assert.equal(attrs.useAutoFormatting, '1');
});
for (const enabled of [false, true]) test(`XLSB 옵션 변환 → XLSX 모델 → 표준 XLSX 재열기 ${enabled}`, () => {
  const attrs = binaryOptions(enabled ? 6 : 6 | 0x80000000 | 0x100000,
    enabled ? 0x100 | 0x80 | 0x20 | 0x40000 : 0, enabled ? 0 : 0x200 | 0x4000);
  const wb = externalOptions(attrs), expected = Object.fromEntries(Object.keys(effective({})).map(key => [key, enabled]));
  assert.deepEqual(effective(wb.sheets[1].pivot), expected);
  const before = wb.serialize(), bytes = writeXlsx(wb), back = new Workbook(readXlsx(bytes).data), written = parseXml(textOf(unzip(bytes)[part])).attrs;
  assert.deepEqual(effective(back.sheets[1].pivot), expected);
  assert.equal(written.useAutoFormatting, enabled ? '1' : '0'); assert.equal(written.multipleFieldFilters, enabled ? '1' : '0');
  assert.deepEqual(back.sheets[1].colWidths, wb.sheets[1].colWidths, '직접 지정한 열 너비를 저장 후에도 유지');
  assert.deepEqual(wb.serialize(), before, '저장은 원본 모델을 변경하지 않음');
});
for (const lexical of ['numeric', 'word']) test(`표준 XLSX 피벗 체크 옵션은 ${lexical} boolean 표기를 읽고 왕복 보존`, () => {
  const no = lexical === 'word' ? 'false' : '0', yes = lexical === 'word' ? 'true' : '1';
  const wb = externalOptions({ useAutoFormatting: no, preserveFormatting: no, enableDrill: no, showHeaders: no, showDrill: no, mergeItem: yes, multipleFieldFilters: yes, customListSort: no });
  const expected = { autofit: false, preserveFormat: false, enableDrill: false, showHeaders: false, showExpand: false, mergeLabels: true, multiFilters: true, customListSort: false };
  assert.deepEqual(effective(wb.sheets[1].pivot), expected);
  assert.deepEqual(effective(readXlsx(writeXlsx(wb)).data.sheets[1].pivot), expected);
});
test('XLSX 생략된 피벗 체크 옵션은 기존 기본값을 유지', () => {
  assert.deepEqual(effective(externalOptions({}).sheets[1].pivot), { autofit: true, preserveFormat: true, enableDrill: true, showHeaders: true, showExpand: true, mergeLabels: false, multiFilters: false, customListSort: true });
});
