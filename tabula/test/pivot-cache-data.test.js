import test from 'node:test';
import assert from 'node:assert/strict';
import { PivotSnapshotBuilder, readPivotSnapshotXml, restorePivotSnapshot, pivotSnapshotValue, cubeFromPivotSnapshot } from '../src/pivot-cache-data.js';
import { readPivotSnapshotBinary } from '../src/xlsb.js';
import { jsonReplacer } from '../src/block.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
const enc = new TextEncoder();
function consume(steps) { for (;;) { const value = steps.next(); if (value.done) return value.value; } }
const fields = [{ name: '분류', db: true, shared: ['서울', '', null, true, { error: '#N/A' }] }, { name: '값', db: true, shared: [] }];
const rows = [['서울', 10], ['', 0], [null, -2], [true, 3.5], [{ error: '#N/A' }, null], ['서울', false]];
function pack(data = rows) { const builder = new PivotSnapshotBuilder(fields.map(f => f.name), 16); data.forEach(row => builder.add(row)); return builder.finish(); }
const values = snapshot => Array.from({ length: snapshot.n }, (_, r) => snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c)));
test('compact cache preserves blank text, missing values, zero, booleans and errors', () => {
  const snapshot = pack(); assert.deepEqual(values(snapshot), rows);
  const cube = cubeFromPivotSnapshot(snapshot); assert.equal(cubeFromPivotSnapshot(snapshot), cube); assert.equal(cube.n, rows.length);
  assert.deepEqual(Array.from({ length: cube.n }, (_, r) => cube.row(r)), rows);
  assert.equal(cube.col(0).dim().keys.length, 5); assert.deepEqual(Array.from(cube.col(1).num()), [10, 0, -2, 3.5, NaN, NaN]);
});
test('compact cache restores JSON numeric-key arrays and base64 typed wrappers without copying live arrays', () => {
  const snapshot = pack(); assert.equal(restorePivotSnapshot(snapshot), snapshot);
  assert.deepEqual(values(restorePivotSnapshot(JSON.parse(JSON.stringify(snapshot)))), rows);
  assert.deepEqual(values(restorePivotSnapshot(JSON.parse(JSON.stringify(snapshot, jsonReplacer)))), rows);
  const old = [fields.map(f => f.name), ...rows]; assert.equal(restorePivotSnapshot(old), old);
});
test('snapshot builder growth preserves every row and dictionaries', () => {
  const data = Array.from({ length: 2050 }, (_, i) => [i % 3 ? '' : '항목', i]);
  assert.deepEqual(values(pack(data)), data);
});
test('XML snapshot records retain shared items, namespace prefixes, escaped strings and dates', () => {
  const xml = '<p:pivotCacheRecords xmlns:p="urn:test"><p:r><p:x v="0"/><p:n v="10"/></p:r><p:r><p:s v=""/><p:b v="false"/></p:r><p:r><p:m/><p:e v="#DIV/0!"/></p:r><p:r><p:s v="&amp;_x000A_"/><p:d v="1900-02-28T12:00:00"/></p:r></p:pivotCacheRecords>';
  const snapshot = consume(readPivotSnapshotXml(enc.encode(xml), fields));
  assert.deepEqual(values(snapshot), [['서울', 10], ['', false], [null, { error: '#DIV/0!' }], ['&\n', 59.5]]);
});
test('XML cached records larger than the former 40 MB cutoff retain all rows', () => {
  const head = enc.encode('<pivotCacheRecords>'), tail = enc.encode('<r><x v="0"/><n v="10"/></r><r><s v=""/><n v="20"/></r></pivotCacheRecords>');
  const bytes = new Uint8Array((40 << 20) + 512).fill(32); bytes.set(head); bytes.set(tail, bytes.length - tail.length);
  assert.deepEqual(values(consume(readPivotSnapshotXml(bytes, fields))), [['서울', 10], ['', 20]]);
});
const u32 = (...values) => { const b = Buffer.alloc(values.length * 4); values.forEach((v, i) => b.writeUInt32LE(v, i * 4)); return b; };
const f64 = value => { const b = Buffer.alloc(8); b.writeDoubleLE(value); return b; };
const variable = value => { const out = []; do { const n = value & 127; value >>>= 7; out.push(n | (value ? 128 : 0)); } while (value); return Buffer.from(out); };
const record = (type, bytes) => Buffer.concat([variable(type), variable(bytes.length), bytes]);
const binaryFields = [{ database: true, items: ['present'], sflags: 8 }, { database: true, items: [], sflags: 0x40 }];
test('binary snapshots preserve packed and per-item records without XML expansion', () => {
  const binary = Buffer.concat([record(33, Buffer.concat([u32(0), f64(10)])), record(34, Buffer.alloc(0)), record(26, u32(1)), record(22, Buffer.from([1])), record(34, Buffer.alloc(0)), record(20, Buffer.alloc(0)), record(23, Buffer.from([7]))]);
  assert.deepEqual(values(consume(readPivotSnapshotBinary(binary, fields, { fields: binaryFields, recordCount: 3 }))), [['서울', 10], ['', true], [null, { error: '#DIV/0!' }]]);
});
test('binary cached records beyond the former 24 MB cutoff retain the saved rows', () => {
  const padding = record(999, Buffer.alloc((24 << 20) + 64));
  const bytes = Buffer.concat([padding, record(33, Buffer.concat([u32(0), f64(10)]))]);
  assert.deepEqual(values(consume(readPivotSnapshotBinary(bytes, fields, { fields: binaryFields, recordCount: 1 }))), [['서울', 10]]);
});
test('imported cached values override changed source values until an explicit source edit', async () => {
  const wb = new Workbook({ sheets: [{ name: '원본', cells: { '0,0': { raw: '분류' }, '0,1': { raw: '값' }, '1,0': { raw: '서울' }, '1,1': { raw: '999' } } }, { name: '피벗', cells: {}, pivot: { name: '피벗1', source: '원본', range: { r1: 0, c1: 0, r2: 1, c2: 1 }, rows: ['분류'], cols: [], values: [{ field: '값', agg: 'sum' }], top: 0, left: 0 } }] });
  const parts = unzip(writeXlsx(wb)), path = 'xl/pivotCache/pivotCacheDefinition1.xml';
  parts[path] = textOf(parts[path]).replace('refreshOnLoad="1"', 'refreshOnLoad="0"');
  parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] = '<Relationships><Relationship Id="rCache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>';
  parts['xl/pivotCache/pivotCacheRecords1.xml'] = '<pivotCacheRecords count="1"><r><s v="서울"/><n v="10"/></r></pivotCacheRecords>';
  const bytes = zip(parts);
  for (const reader of [readXlsx, readXlsxAsync]) {
    const imported = new Workbook((await reader(bytes)).data), def = imported.sheets[1].pivot;
    assert.ok(def.snapshotId); assert.equal(imported.pivotSnapshots.get(def.snapshotId).rows.kind, 'pivot-cache');
    let source = pivotSourceData(imported, def); assert.equal(source.cube.row(0)[1], 10);
    const resolved = resolvePivot(source, def); assert.ok(computePivot(resolved, resolved.def).grid.some(row => row.some(cell => cell?.raw === '10')));
    imported.transact(() => imported.setInput(0, 1, 1, '888'));
    source = pivotSourceData(imported, def); assert.equal(source.cube.row(0)[1], 888);
  }
});
