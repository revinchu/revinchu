import test from 'node:test';
import assert from 'node:assert/strict';
import { PivotSnapshotBuilder, pivotSnapshotValue, restorePivotSnapshot, cubeFromPivotSnapshot } from '../src/pivot-cache-data.js';
import { jsonReplacer, jsonReviver } from '../src/block.js';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
import { pivotExportData } from '../src/pivot-export-data.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { saveLargeWorkbook, loadLargeWorkbook } from '../src/big-storage.js';
import { writeWixelFile, readWixelFile } from '../src/wixel-file.js';

function pack(rows, expected = 0) {
  const b = new PivotSnapshotBuilder(rows[0].map((_, i) => 'F' + i), expected);
  for (const row of rows) b.add(row);
  return b.finish();
}
const values = snapshot => Array.from({ length: snapshot.n }, (_, r) => snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c)));
function workbook(snapshot) {
  const n = Array.isArray(snapshot) ? snapshot.length - 1 : snapshot.n, header = Array.isArray(snapshot) ? snapshot[0] : snapshot.header;
  return new Workbook({ pivotSnapshots: { saved: snapshot }, sheets: [
    { name: 'Source', cells: { '0,0': { raw: 'F0' }, '0,1': { raw: 'F1' }, '1,0': { raw: 'different' }, '1,1': { raw: '999' } } },
    { name: 'Report', cells: {}, pivot: { name: 'SavedPivot', source: 'Source', range: { r1: 0, c1: 0, r2: n, c2: header.length - 1 }, snapshotId: 'saved', rows: ['F0'], values: [{ field: 'F1', agg: 'sum' }], top: 0, left: 0 } },
  ] });
}
const saved = w => w.snapshotData().saved;
const source = w => pivotSourceData(w, w.sheets[1].pivot);
function grid(w) { const resolved = resolvePivot(source(w), w.sheets[1].pivot); return computePivot(resolved, resolved.def).grid.map(row => row.map(cell => cell.raw)); }

// Compare decoded values with independently retained input rows, not compressed metadata.
test('uniform tails keep row count/order and distinguish null, zero, empty text, booleans and errors', () => {
  const tail = [null, 0, '', false, true, { error: '#N/A' }];
  const rows = [[12, 'prefix', 3, 5, null, 'ok'], ...Array.from({ length: 256 }, () => tail)];
  const snapshot = pack(rows, 16);
  assert.equal(snapshot.n, 257); assert.deepEqual(values(snapshot), rows);
  for (let c = 0; c < tail.length; c++) {
    const col = snapshot.columns[c]; assert.deepEqual(col.tail, { start: 1, value: tail[c] });
    if (col.num) assert.equal(col.num.length, 1);
    if (col.str) assert.equal(col.str.length, 1);
    assert.equal(pivotSnapshotValue(snapshot, snapshot.n, c), null);
  }
  assert.equal(pivotSnapshotValue(snapshot, -1, 0), null);
  assert.equal(pivotSnapshotValue(snapshot, 10, 99), null);
});

test('the 256-row threshold and suffix boundary do not merge unlike values', () => {
  const short = pack([[7], ...Array.from({ length: 255 }, () => [0])]);
  assert.equal(short.columns[0].tail, undefined); assert.equal(short.columns[0].num.length, 256);
  const rows = [[7], [0], [null], [''], ...Array.from({ length: 256 }, () => [0])];
  const snapshot = pack(rows);
  assert.deepEqual(snapshot.columns[0].tail, { start: 4, value: 0 });
  assert.equal(snapshot.columns[0].num.length, 4); assert.deepEqual(values(snapshot), rows);
  const alternating = pack(Array.from({ length: 1000 }, (_, i) => [i % 2 ? null : 0]));
  assert.equal(alternating.columns[0].tail, undefined); assert.equal(alternating.columns[0].num.length, 1000);
});

test('entire constant columns allow zero-length prefixes; missing-only columns allocate no arrays', () => {
  const rows = Array.from({ length: 300 }, () => [null, '', 0, false]);
  const snapshot = pack(rows);
  assert.deepEqual(values(snapshot), rows);
  assert.equal(snapshot.columns[0].num, null); assert.equal(snapshot.columns[0].str, null);
  for (const col of snapshot.columns.slice(1)) { assert.equal(col.tail.start, 0); assert.equal(col.num?.length ?? 0, 0); assert.equal(col.str?.length ?? 0, 0); }
  const empty = new PivotSnapshotBuilder(['Empty']).finish();
  assert.equal(empty.n, 0); assert.equal(empty.columns[0].tail, undefined); assert.equal(restorePivotSnapshot(empty), empty);
});

test('a million saved rows retain exact random access while final typed storage keeps only the prefix', () => {
  const n = 1000000, prefix = 17, b = new PivotSnapshotBuilder(['Number', 'Text'], n);
  for (let r = 0; r < n; r++) b.add(r < prefix ? [r + 1, 'Item' + r] : [0, null]);
  const snapshot = b.finish();
  assert.equal(snapshot.n, n);
  assert.deepEqual(snapshot.columns.map(col => col.tail), [{ start: prefix, value: 0 }, { start: prefix, value: null }]);
  assert.equal(snapshot.columns[0].num.byteLength, prefix * 8); assert.equal(snapshot.columns[1].str.byteLength, prefix * 4);
  for (const r of [0, prefix - 1, prefix, 500000, n - 1]) {
    assert.deepEqual(snapshot.header.map((_, c) => pivotSnapshotValue(snapshot, r, c)), r < prefix ? [r + 1, 'Item' + r] : [0, null]);
  }
  for (const col of b.columns) { assert.equal(col.num, null); assert.equal(col.str, null); assert.equal(col.index.size, 0); }
});

test('tail snapshots roundtrip plain JSON and typed JSON without expanding prefix arrays', () => {
  const rows = [[3, 'start', false], ...Array.from({ length: 300 }, () => [null, '', { code: '#DIV/0!' }])];
  const snapshot = pack(rows);
  assert.equal(restorePivotSnapshot(snapshot), snapshot);
  for (const copy of [JSON.parse(JSON.stringify(snapshot)), JSON.parse(JSON.stringify(snapshot, jsonReplacer)), JSON.parse(JSON.stringify(snapshot, jsonReplacer), jsonReviver)]) {
    const restored = restorePivotSnapshot(copy);
    assert.deepEqual(values(restored), rows); assert.equal(restored.n, rows.length);
    assert.equal(restored.columns[0].num.length, 1); assert.equal(restored.columns[1].str.length, 1);
    assert.deepEqual(restored.columns.map(col => col.tail), snapshot.columns.map(col => col.tail));
  }
});

test('negative zero stays typed so typed JSON preserves its sign', () => {
  const snapshot = pack(Array.from({ length: 300 }, () => [-0]));
  assert.equal(snapshot.columns[0].tail, undefined);
  const copy = restorePivotSnapshot(JSON.parse(JSON.stringify(snapshot, jsonReplacer)));
  assert.ok(Object.is(pivotSnapshotValue(copy, 299, 0), -0));
  const mixed = pack([[-0], ...Array.from({ length: 256 }, () => [0])]);
  assert.equal(mixed.columns[0].tail.start, 1);
  assert.ok(Object.is(pivotSnapshotValue(mixed, 0, 0), -0)); assert.ok(Object.is(pivotSnapshotValue(mixed, 1, 0), 0));
});

test('v1 restore rejects corrupt tail metadata and full/short prefix mismatches', () => {
  const column = { num: new Float64Array([9]), str: null, dict: [], tail: { start: 1, value: 0 } };
  const snapshot = { kind: 'pivot-cache', version: 1, n: 300, header: ['F0'], columns: [column] };
  const withColumn = col => ({ ...snapshot, columns: [col] });
  for (const tail of [null, [], {}, { start: -1, value: 0 }, { start: 300, value: 0 }, { start: 1.5, value: 0 }, { start: 1 }, { start: 1, value: undefined }, { start: 1, value: NaN }, { start: 1, value: Infinity }, { start: 1, value: {} }, { start: 1, value: [] }]) {
    assert.throws(() => restorePivotSnapshot(withColumn({ ...column, tail })), /꼬리 값/);
  }
  for (const num of [new Float64Array(300), new Float64Array(0), [9, 10], { 0: 9, 1: 10 }, new Int32Array([9])]) {
    assert.throws(() => restorePivotSnapshot(withColumn({ ...column, num })), /길이/);
  }
  assert.throws(() => restorePivotSnapshot(withColumn({ ...column, str: new Int32Array(300) })), /길이/);
  for (const value of [null, 0, '', false, true, 'tail', { error: '#N/A' }, { code: '#REF!' }]) {
    const restored = restorePivotSnapshot(withColumn({ ...column, tail: { start: 1, value } }));
    assert.deepEqual(pivotSnapshotValue(restored, 299, 0), value);
  }
});

test('legacy v1 full columns and row arrays remain compatible and invalid full typed lengths remain rejected', () => {
  const snapshot = { kind: 'pivot-cache', version: 1, n: 3, header: ['F0'], columns: [{ num: new Float64Array([2, NaN, 0]), str: null, dict: [] }] };
  assert.equal(restorePivotSnapshot(snapshot), snapshot);
  assert.deepEqual(values(restorePivotSnapshot(JSON.parse(JSON.stringify(snapshot)))), [[2], [null], [0]]);
  assert.throws(() => restorePivotSnapshot({ ...snapshot, columns: [{ ...snapshot.columns[0], num: new Float64Array(2) }] }), /길이/);
  assert.throws(() => restorePivotSnapshot({ ...snapshot, columns: [{ ...snapshot.columns[0], str: new Int32Array(2) }] }), /길이/);
  const legacy = [['F0'], [2], [null], [0]]; assert.equal(restorePivotSnapshot(legacy), legacy);
});

test('Cube rebuilds full numeric/count/dimension/error arrays through getters, never a short pre.num', () => {
  const rows = [[3, 'prefix', true, 5], ...Array.from({ length: 300 }, () => [0, '', false, { error: '#N/A' }])];
  const snapshot = pack(rows), cube = cubeFromPivotSnapshot(snapshot);
  assert.equal(cube.n, rows.length); assert.equal(cubeFromPivotSnapshot(snapshot), cube);
  assert.equal(cube.col(0)._num, null);
  assert.equal(cube.col(0).num().length, rows.length); assert.equal(cube.col(0).num()[300], 0);
  assert.equal(cube.col(1).ne()[300], 1); assert.equal(cube.col(2).ne()[300], 1);
  assert.equal(cube.col(3).num().length, rows.length); assert.ok(Number.isNaN(cube.col(3).num()[300]));
  assert.deepEqual(cube.col(3).errs().get(300), { code: '#N/A' });
  assert.deepEqual(cube.row(300), rows[300]); assert.equal(cube.col(1).dim().codes.length, rows.length);
});

test('pivot aggregates, filters and indexed export match uncompressed saved records including the full tail', () => {
  for (const tail of [0, null, '', false, { error: '#N/A' }]) {
    const rows = [['prefix', 7], ...Array.from({ length: 300 }, () => ['same', tail])];
    const compact = workbook(pack(rows)), legacy = workbook([rows[0].map((_, i) => 'F' + i), ...rows]);
    for (const agg of ['sum', 'count', 'countNums', 'average', 'min', 'max']) {
      compact.sheets[1].pivot.values = [{ field: 'F1', agg }]; legacy.sheets[1].pivot.values = [{ field: 'F1', agg }];
      assert.deepEqual(grid(compact), grid(legacy), String(tail) + ':' + agg);
    }
    compact.sheets[1].pivot.filters = { F0: ['same'] }; legacy.sheets[1].pivot.filters = { F0: ['same'] };
    assert.deepEqual(grid(compact), grid(legacy));
    const exported = pivotExportData(source(compact).cube, true); assert.equal(exported.length, rows.length);
    assert.deepEqual(exported.value(300, 1), tail); assert.deepEqual([...exported.values(1)], rows.map(row => row[1]));
  }
});

test('Workbook JSON/Blob and Undo/Redo preserve tail metadata and saved values over different live source values', async () => {
  const rows = [['prefix', 17], ...Array.from({ length: 300 }, () => ['', 0])], snapshot = pack(rows), w = workbook(snapshot);
  for (const data of [JSON.parse(JSON.stringify(w.serialize())), JSON.parse(await w.serializeBlob().text()), { ...w.bookMeta(), sheets: w.serialize().sheets }]) {
    const copy = new Workbook(data); assert.deepEqual(values(saved(copy)), rows);
    assert.equal(saved(copy).columns[0].str.length, 1); assert.equal(source(copy).cube.n, rows.length); assert.equal(copy.getValue(0, 1, 1), 999);
  }
  w.transact(() => w.setSheetProp(1, 'state', 'hidden'));
  for (const command of ['undo', 'redo']) { w[command](); assert.deepEqual(values(saved(w)), rows); assert.equal(saved(w).columns[1].num.length, 1); }
});

test('native WIXEL framed files preserve prefix arrays and all tail scalars with and without gzip', async () => {
  const rows = [['prefix', 7, false], ...Array.from({ length: 300 }, () => ['', 0, { code: '#REF!' }])], w = workbook(pack(rows));
  for (const gzip of [false, true]) {
    const file = await writeWixelFile(w, { docName: 'Synthetic tail' }, { gzip }), data = await readWixelFile(file), copy = new Workbook(data.workbook);
    assert.deepEqual(values(saved(copy)), rows); assert.deepEqual(saved(copy).columns.map(col => col.tail), saved(w).columns.map(col => col.tail));
    assert.equal(saved(copy).columns[1].num.length, 1); assert.deepEqual(grid(copy), grid(w));
  }
});

test('XLSX saved-cache export/reopen keeps all trailing records and type distinctions', () => {
  const rows = [['prefix', 7, 9, false], ...Array.from({ length: 300 }, () => [null, 0, '', { error: '#N/A' }])], w = workbook(pack(rows));
  const bytes = writeXlsx(w), files = unzip(bytes);
  assert.match(textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']), /recordCount="301"/);
  const records = textOf(files['xl/pivotCache/pivotCacheRecords1.xml']); assert.equal((records.match(/<r>/g) ?? []).length, 301);
  const copy = new Workbook(readXlsx(bytes).data), cube = source(copy).cube;
  assert.equal(cube.n, rows.length); assert.deepEqual(Array.from({ length: cube.n }, (_, r) => cube.row(r)), rows);
  assert.deepEqual(grid(copy), grid(w)); assert.equal(copy.getValue(0, 1, 1), 999);
});

// A small in-memory IndexedDB implementation exercises actual saveLargeWorkbook/loadLargeWorkbook.
function memoryIDB() {
  const data = new Map(); let connection;
  return { data, get connection() { return connection; }, indexedDB: { open() {
    const req = {}; queueMicrotask(() => {
      connection = { close() {}, transaction() {
        let pending = 0, finished = false; const writes = new Map(), deletes = new Set();
        const finish = () => setImmediate(() => { if (finished || pending) return; finished = true; for (const key of deletes) data.delete(key); for (const [key, value] of writes) data.set(key, value); tx.oncomplete?.(); });
        const request = fn => { const r = {}; pending++; queueMicrotask(() => { r.result = fn(); r.onsuccess?.(); pending--; finish(); }); return r; };
        const tx = { abort() { finished = true; queueMicrotask(() => tx.onabort?.()); }, objectStore() { return {
          get: key => request(() => structuredClone(writes.has(key) ? writes.get(key) : deletes.has(key) ? undefined : data.get(key))),
          put: (value, key) => { const copy = structuredClone(value); return request(() => { writes.set(key, copy); return key; }); },
          delete: key => request(() => deletes.add(key)),
          getAllKeys: range => request(() => [...data.keys()].filter(key => key >= range.lower && key <= range.upper)),
        }; } }; finish(); return tx;
      } }; req.result = connection; req.onsuccess?.();
    }); return req;
  } } };
}

test('large-storage column splitting/restoration keeps tail metadata with prefix-only typed parts', async () => {
  const h = memoryIDB(), prior = globalThis.indexedDB, priorRange = globalThis.IDBKeyRange;
  globalThis.indexedDB = h.indexedDB; globalThis.IDBKeyRange = { bound: (lower, upper) => ({ lower, upper }) };
  try {
    const rows = [['prefix', 17, true, 'initial'], ...Array.from({ length: 300 }, () => ['', 0, null, { error: '#N/A' }])], w = workbook(pack(rows));
    const result = await saveLargeWorkbook('synthetic:tail-cache', w, {});
    const record = h.data.get(result.manifest.snapshotKey), columns = record.data.saved.columns;
    assert.deepEqual(columns.map(col => col.tail), saved(w).columns.map(col => col.tail));
    for (const col of columns) for (const name of ['num', 'str']) if (col[name + 'Parts']) assert.equal(col[name + 'Parts'].len, 1);
    const loaded = await loadLargeWorkbook('synthetic:tail-cache'), copy = new Workbook(loaded.workbook);
    assert.deepEqual(values(saved(copy)), rows); assert.deepEqual(grid(copy), grid(w));
    assert.equal(saved(copy).columns[1].num.length, 1); assert.equal(source(copy).cube.n, rows.length);
    const constantRows = Array.from({ length: 300 }, () => ['', 0]), constant = workbook(pack(constantRows));
    const constantResult = await saveLargeWorkbook('synthetic:constant-tail', constant, {});
    const constantParts = h.data.get(constantResult.manifest.snapshotKey).data.saved.columns;
    assert.deepEqual(constantParts[0].strParts, { parts: [], len: 0, kind: 'Int32Array' });
    assert.deepEqual(constantParts[1].numParts, { parts: [], len: 0, kind: 'Float64Array' });
    const constantCopy = new Workbook((await loadLargeWorkbook('synthetic:constant-tail')).workbook);
    assert.deepEqual(values(saved(constantCopy)), constantRows); assert.equal(saved(constantCopy).columns[1].num.length, 0);
  } finally { h.connection?.onversionchange?.(); globalThis.indexedDB = prior; globalThis.IDBKeyRange = priorRange; }
});
