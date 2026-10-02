import { test } from 'node:test';
import assert from 'node:assert/strict';
import { convertXlsb } from '../src/xlsb.js';
import { parseXml, child } from '../src/xml.js';

// MS-XLSB 2.4.278 BrtBeginSXView:
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/f6cbfd44-775f-44f6-8ba8-fac46d7a445c
// 사용자 파일 없이 고정 필드와 인접 플래그를 직접 구성한다.
const u32 = (...xs) => { const b = Buffer.alloc(xs.length * 4); xs.forEach((x, i) => b.writeUInt32LE(x >>> 0, i * 4)); return b; };
const wide = s => Buffer.concat([u32(s.length), Buffer.from(s, 'utf16le')]);
const variable = x => { const out = []; do { const n = x & 127; x >>>= 7; out.push(n | (x ? 128 : 0)); } while (x); return Buffer.from(out); };
const record = (id, data) => Buffer.concat([variable(id), variable(data.length), data]);
function table({ over, wrap, flags = 0 }) {
  const fixed = Buffer.alloc(32);
  fixed.writeUInt32LE(6, 0);
  fixed.writeUInt32LE(0x80000 | flags | (over ? 0x800 : 0), 4);
  fixed.writeUInt32LE(0xc0, 8);
  fixed[12] = 2; fixed[13] = wrap; fixed[14] = 7; fixed[15] = 3;
  fixed.writeInt32LE(-1, 16); fixed.writeUInt16LE(1, 20); fixed.writeUInt32LE(19, 24); fixed.writeUInt32LE(42, 28);
  const files = { 'xl/workbook.bin': new Uint8Array(), 'xl/pivotTables/pivotTable1.bin': Buffer.concat([
    record(280, Buffer.concat([fixed, wide('합성보고서'), wide('집계 값')])),
    record(314, u32(7, 10, 1, 2, 8, 8, 2, 2, 2)),
  ]) };
  for (const _ of convertXlsb(files)) { /* 변환 완료 */ }
  return parseXml(new TextDecoder().decode(files['xl/pivotTables/pivotTable1.bin']));
}

for (const over of [false, true]) for (const wrap of [0, 2, 255]) {
  test(`XLSB 보고서 필터 순서·줄바꿈: over=${over}, wrap=${wrap}`, () => {
    const root = table({ over, wrap });
    assert.equal(root.attrs.pageOverThenDown ?? '0', over ? '1' : '0');
    assert.equal(Number(root.attrs.pageWrap ?? 0), wrap);
    assert.equal(root.attrs.name, '합성보고서'); assert.equal(root.attrs.dataCaption, '집계 값');
    assert.equal(root.attrs.cacheId, '42', '바이트 위치가 뒤쪽 캐시와 문자열을 밀지 않음');
    assert.deepEqual(child(root, 'location').attrs, { ref: 'B8:C11', firstHeaderRow: '1', firstDataRow: '1', firstDataCol: '1', rowPageCount: '2', colPageCount: '2' });
  });
}
test('XLSB 인접 플래그는 보고서 필터 가로 배치를 잘못 켜지 않음', () => {
  const root = table({ over: false, wrap: 2, flags: 0x400 | 0x1000 });
  assert.equal(root.attrs.pageOverThenDown, undefined);
  assert.equal(root.attrs.pageWrap, '2');
});
