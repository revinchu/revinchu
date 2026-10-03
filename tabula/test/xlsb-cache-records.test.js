import test from 'node:test';
import assert from 'node:assert/strict';
import { convertXlsb, readPivotSnapshotBinary } from '../src/xlsb.js';
import { readPivotSnapshotXml, pivotSnapshotValue } from '../src/pivot-cache-data.js';
import { textOf } from '../src/zip.js';

// MS-XLSB 2.4.742, synthetic data only:
// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/0b06f09d-9dad-49f3-b96d-523363ba81b9
const u32 = x => { const b = Buffer.alloc(4); b.writeUInt32LE(x); return b; };
const u16 = x => { const b = Buffer.alloc(2); b.writeUInt16LE(x); return b; };
const f64 = x => { const b = Buffer.alloc(8); b.writeDoubleLE(x); return b; };
const wide = s => Buffer.concat([u32(s.length), Buffer.from(s, 'utf16le')]);
const variable = x => { const out = []; do { const n = x & 127; x >>>= 7; out.push(n | (x ? 128 : 0)); } while (x); return Buffer.from(out); };
const record = (id, bytes) => Buffer.concat([variable(id), variable(bytes.length), bytes]);
const consume = steps => { for (;;) { const step = steps.next(); if (step.done) return step.value; } };
const values = snapshot => Array.from({ length: snapshot.n }, (_, r) => snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c)));
const configs = [
  { flags: 0x12f, value: wide('2026-10-04 또는 미정'), expected: '2026-10-04 또는 미정' },
  { flags: 0x13f, value: wide(''), expected: '' },
  { flags: 0x44, value: f64(123.5), expected: 123.5 },
  { flags: 0x104, value: Buffer.concat([u16(1900), u16(2), Buffer.from([28, 12, 0, 0])]), expected: 59.5 },
  { flags: 0x11, value: wide(''), expected: '' },
  { flags: 0, value: wide('나머지 문자열'), expected: '나머지 문자열' },
  { flags: 0x14f, items: ['공유 항목'], value: u32(0), expected: '공유 항목' },
];
const fields = configs.map((f, i) => ({ name: `열${i}`, db: true, shared: f.items ?? [] }));
const metadata = { fields: configs.map(f => ({ database: true, items: f.items ?? [], sflags: f.flags })), recordCount: 1 };
const bytes = record(33, Buffer.concat(configs.map(f => f.value)));
test('packed XLSB caches distinguish mixed strings, numeric/date precedence, dates, blank strings and shared indexes', () => {
  assert.deepEqual(values(consume(readPivotSnapshotBinary(bytes, fields, metadata))), [configs.map(f => f.expected)]);
});
test('packed date-only fields preserve the workbook date system', () => {
  const f = [{ name: '날짜', db: true, shared: [] }], m = { fields: [{ database: true, items: [], sflags: 4 }], recordCount: 1 };
  const data = record(33, Buffer.concat([u16(1904), u16(1), Buffer.from([1, 0, 0, 0])]));
  assert.deepEqual(values(consume(readPivotSnapshotBinary(data, f, m, true))), [[0]]);
  assert.deepEqual(values(consume(readPivotSnapshotBinary(data, f, m, false))), [[1462]]);
});
test('synchronous XLSB conversion preserves the same packed cache values as the asynchronous reader', () => {
  const header = Buffer.alloc(21); header.writeInt32LE(1, 17);
  const definition = [record(179, header)];
  configs.forEach((f, i) => {
    const fixed = Buffer.alloc(20); fixed.writeUInt16LE(4, 0);
    definition.push(record(183, Buffer.concat([fixed, wide(`열${i}`)])), record(189, Buffer.concat([u16(f.flags), u32(f.items?.length ?? 0)])));
    for (const item of f.items ?? []) definition.push(record(24, wide(item)));
    definition.push(record(190, Buffer.alloc(0)), record(184, Buffer.alloc(0)));
  });
  const files = {
    'xl/workbook.bin': new Uint8Array(),
    'xl/pivotCache/pivotCacheDefinition1.bin': Buffer.concat(definition),
    'xl/pivotCache/pivotCacheRecords1.bin': bytes,
    'xl/pivotCache/_rels/pivotCacheDefinition1.bin.rels': new TextEncoder().encode('<Relationships><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.bin"/></Relationships>'),
  };
  consume(convertXlsb(files));
  assert.ok(files['xl/pivotCache/pivotCacheRecords1.bin']);
  assert.ok(!textOf(files['xl/pivotCache/pivotCacheDefinition1.bin']).includes('refreshOnLoad="1"'));
  assert.deepEqual(values(consume(readPivotSnapshotXml(files['xl/pivotCache/pivotCacheRecords1.bin'], fields))), [configs.map(f => f.expected)]);
});
for (const [label, flags, body, items] of [
  ['string length', 0x12f, Buffer.alloc(3), []],
  ['string characters', 0x12f, Buffer.concat([u32(3), Buffer.from('가', 'utf16le')]), []],
  ['number', 0x44, Buffer.alloc(7), []],
  ['date', 4, Buffer.alloc(7), []],
  ['shared index', 0x40, Buffer.alloc(3), ['항목']],
]) test(`truncated packed ${label} cannot read into the next BIFF record`, () => {
  const m = { fields: [{ database: true, items, sflags: flags }], recordCount: 1 };
  const f = [{ name: '열', db: true, shared: items }];
  const data = Buffer.concat([record(33, body), record(999, Buffer.alloc(32))]);
  assert.throws(() => consume(readPivotSnapshotBinary(data, f, m)), /끝까지 저장되지/);
});
test('invalid cache metadata and invalid string lengths fail explicitly', () => {
  const f = [{ name: '열', db: true, shared: [] }];
  assert.throws(() => consume(readPivotSnapshotBinary(record(33, wide('')), f, { fields: [{ database: true, items: [], sflags: NaN }], recordCount: 1 })), /형식 정보/);
  for (const length of [32768, 0xffffffff]) assert.throws(() => consume(readPivotSnapshotBinary(record(33, u32(length)), f, { fields: [{ database: true, items: [], sflags: 0x12f }], recordCount: 1 })), /문자열의 길이/);
});
