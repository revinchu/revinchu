import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap, getSharedBlankCell } from '../src/cellmap.js';
import { RunColumn } from '../src/run-column.js';
import { createImportedBlankRuns } from '../src/import-blank-runs.js';
import { Workbook, makeCellRC } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const styles = [{ bold: true, fill: '#112233', border: { bottom: { color: '#aabbcc', style: 'thin' } } }, { italic: true, fill: '#445566' }];
function fixture(rows = 40, cols = 80) {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) collector.add(r, c, styles[r % 2]);
  const stats = collector.finish(); return { cells, stats };
}
function snapshots(column) { return [...column].map(([r, v]) => [r, v]); }
function checkSame(column, expected) { assert.equal(column.size, expected.size); assert.deepEqual(snapshots(column), [...expected]); }

test('horizontal imported blanks preserve every coordinate, styles and logical size with shared column backing', () => {
  const { cells, stats } = fixture(64, 4096);
  assert.equal(cells.size, 64 * 4096); assert.equal(stats.horizontalSpans, 64); assert.equal(stats.uniquePatterns, 1);
  assert.equal(stats.sharedColumns, 4096); assert.equal(cells.col(0).runs.length, 64); assert.equal(cells.col(0).order.length, 1);
  assert.equal(cells.col(0).runs, cells.col(4095).runs); assert.equal(cells.col(0).points, cells.col(4095).points);
  for (const c of [0, 73, 4095]) for (const r of [0, 17, 63]) { assert.ok(cells.hasRC(r, c)); assert.equal(cells.getRC(r, c).style, styles[r % 2]); }
  assert.equal(cells.hasRC(64, 4095), false); assert.equal(cells.hasRC(0, 4096), false);
});

test('column encounter order and point/formula candidates survive deferred blank insertion', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells), formula = makeCellRC({ raw: '=A1' }, 1, 5);
  collector.add(0, 5, styles[0]); cells.setRC(0, 7, { raw: 'header' });
  collector.add(1, 2, styles[1]); cells.setRC(1, 5, formula); collector.add(2, 5, styles[1]);
  assert.equal(cells.size, 5); collector.finish();
  assert.deepEqual([...cells.cols.keys()], [5, 7, 2]); assert.deepEqual([...cells.col(5).keys()], [0, 1, 2]);
  assert.equal(cells.getRC(1, 5), formula); assert.deepEqual([...cells.formulaEntries()], [[1, 5, formula]]); assert.equal(cells.formulaAsts.get(formula.ast), 1);
});

test('mixed-column merge retains annotated cells and holes, with point priority', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  collector.add(0, 0, styles[0]); collector.add(0, 1, styles[0]);
  const noted = { raw: '3', style: styles[1], comment: 'retained', link: '#A1' };
  cells.setRC(1, 0, noted); collector.add(1, 1, styles[1]); collector.add(3, 0, styles[0]); collector.add(3, 1, styles[0]);
  collector.finish(); assert.equal(cells.size, 6); assert.equal(cells.getRC(1, 0), noted); assert.equal(cells.getRC(2, 0), undefined);
  assert.deepEqual([...cells.col(0).keys()], [0, 1, 3]); assert.deepEqual([...cells.col(1).keys()], [0, 1, 3]);
  const links = []; cells.forEachLinkRC((value, r, c) => links.push([r, c, value])); assert.deepEqual(links, [[1, 0, noted]]);
});

test('shared imported columns detach for direct set, setRun, delete and clear', () => {
  for (const mutate of [column => column.set(17, { raw: 'edited' }), column => column.setRun(9, 4, getSharedBlankCell({ underline: true })), column => column.delete(13), column => column.clear()]) {
    const { cells } = fixture(), source = cells.col(1), expected = new Map(source), target = cells.col(0);
    mutate(target); checkSame(source, expected); assert.notEqual(target.dataKey, source.dataKey);
  }
});

test('live compressed and plain shared iterators observe only their own subsequent writes', () => {
  for (const compressed of [false, true]) {
    const a = compressed ? RunColumn.fromSortedStorage(value => value === getSharedBlankCell(styles[0]), [[0, getSharedBlankCell(styles[0]), 20]]) : new RunColumn(() => false);
    if (!compressed) for (let r = 0; r < 20; r++) a.set(r, { raw: String(r) });
    const b = a.shareData(), before = new Map(a), expected = new Map(b), actualIt = b.entries(), expectedIt = expected.entries();
    assert.deepEqual(actualIt.next(), expectedIt.next());
    const value = { raw: 'changed' }; b.delete(5); expected.delete(5); b.set(2, value); expected.set(2, value); b.set(21, value); expected.set(21, value);
    assert.deepEqual([...actualIt], [...expectedIt]); checkSame(a, before); checkSame(b, expected);
  }
});

test('shared mapValues normalizes one pattern once and retains backing across all wrappers', () => {
  const { cells } = fixture(12, 160); let calls = 0, visited = 0;
  for (const count of cells.mapValues(() => { throw Error('blank-only column has no point'); }, value => { calls++; return getSharedBlankCell(value.style, true); })) visited += count;
  assert.equal(visited, 12 * 160); assert.equal(cells.size, visited); assert.equal(calls, 12);
  assert.equal(cells.col(0).runs, cells.col(159).runs); assert.equal(cells.getRC(9, 150).v, null);
  cells.setRC(9, 0, { raw: '19' }); assert.equal(cells.getRC(9, 150).raw, '');
});

test('shared mapValues can remove a blank pattern without retaining logical count', () => {
  const { cells } = fixture(11, 35); let count = 0;
  for (const n of cells.mapValues(() => null, () => null)) count += n;
  assert.equal(count, 11 * 35); assert.equal(cells.size, 0); assert.equal(cells.cols.size, 0);
});

test('unordered and duplicate source positions flush once and follow ordinary Map semantics', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells), expected = new Map();
  const add = (r, style) => { collector.add(r, 0, style); expected.set(r, getSharedBlankCell(style)); };
  add(0, styles[0]); add(3, styles[0]); add(1, styles[1]); add(3, styles[1]); add(7, styles[0]); collector.finish();
  assert.equal(collector.disabled, true); assert.equal(cells.size, expected.size); checkSame(cells.col(0), expected);
  const second = new CellMap(), duplicate = createImportedBlankRuns(second); duplicate.add(0, 0, styles[0]); duplicate.add(0, 0, styles[1]);
  assert.equal(second.size, 1); assert.equal(second.getRC(0, 0).style, styles[1]); assert.equal(duplicate.disabled, true);
});

test('explicit flush leaves subsequent nonmonotonic point insertions in native order', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  collector.add(2, 0, styles[0]); collector.add(4, 0, styles[0]); collector.flush();
  cells.setRC(0, 0, { raw: 'after flush' }); collector.add(1, 0, styles[1]);
  assert.deepEqual([...cells.col(0).keys()], [2, 4, 0, 1]); assert.equal(cells.size, 4);
});

test('nonadjacent identical blank patterns share backing while a different pattern remains independent', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  for (let r = 0; r < 14; r++) for (let c = 0; c < 3; c++) collector.add(r, c, c === 1 ? styles[1] : styles[0]);
  collector.finish(); assert.equal(cells.col(0).runs, cells.col(2).runs); assert.notEqual(cells.col(0).runs, cells.col(1).runs);
  cells.col(0).set(4, { raw: 'new' }); assert.equal(cells.getRC(4, 2).raw, ''); assert.equal(cells.getRC(4, 1).style, styles[1]);
});

test('workbook hydration, ordinary edit Undo/Redo and row structure preserve independent shared columns', () => {
  const { cells } = fixture(24, 8), wb = new Workbook({ sheets: [{ name: 'Synthetic', cells }] });
  assert.equal(wb.sheets[0].cells.size, 24 * 8); assert.equal(wb.sheets[0].cells.col(0).runs, wb.sheets[0].cells.col(7).runs);
  wb.transact(() => wb.setInput(0, 8, 0, '19')); assert.equal(wb.getValue(0, 8, 0), 19); assert.equal(wb.getValue(0, 8, 7), null);
  wb.undo(); assert.equal(wb.getValue(0, 8, 0), null); assert.equal(wb.styleAt(0, 8, 0).fill, styles[0].fill);
  wb.redo(); assert.equal(wb.getValue(0, 8, 0), 19); assert.equal(wb.getValue(0, 8, 7), null);
  wb.transact(() => wb.insertRows(0, 3, 2)); assert.equal(wb.getCell(0, 3, 7), undefined); assert.equal(wb.styleAt(0, 5, 7).fill, styles[1].fill);
  wb.undo(); assert.equal(wb.styleAt(0, 3, 7).fill, styles[1].fill); assert.equal(wb.getValue(0, 8, 0), 19);
  wb.redo(); assert.equal(wb.getValue(0, 10, 0), 19); assert.equal(wb.getValue(0, 10, 7), null);
});

test('compact JSON and XLSX round trips preserve all shared blank coordinates, styles and values', () => {
  const { cells } = fixture(18, 6); cells.setRC(5, 0, { raw: '23', style: styles[1] });
  const wb = new Workbook({ sheets: [{ name: 'Synthetic', cells }] });
  for (const restored of [new Workbook(wb.serialize()), new Workbook(readXlsx(writeXlsx(wb)).data)]) {
    assert.equal(restored.sheets[0].cells.size, 18 * 6); assert.equal(restored.getValue(0, 5, 0), 23);
    for (const r of [0, 5, 17]) for (const c of [1, 5]) { assert.equal(restored.getValue(0, r, c), null); assert.equal(restored.styleAt(0, r, c).fill, styles[r % 2].fill); }
    assert.equal(restored.styleAt(0, 0, 1).bold, true);
  }
});

test('shareData called during a native Map iterator does not hide later original mutations', () => {
  const a = new RunColumn(() => false), expected = new Map();
  for (let r = 0; r < 9; r++) { const cell = { raw: String(r) }; a.set(r, cell); expected.set(r, cell); }
  const it = a.entries(), oracle = expected.entries(); assert.deepEqual(it.next(), oracle.next());
  const copied = a.shareData(), before = new Map(copied), cell = { raw: 'updated' };
  a.delete(4); expected.delete(4); a.set(2, cell); expected.set(2, cell); a.set(15, cell); expected.set(15, cell);
  assert.deepEqual([...it], [...oracle]); checkSame(copied, before); checkSame(a, expected);
});

test('sharing an active compressed iterator retains later end extension and independent snapshot data', () => {
  const value=getSharedBlankCell(styles[0]);
  for (const bounded of [false,true]) for (const useRun of [false,true]) {
    const column=RunColumn.fromSortedStorage(cell=>cell===value,[[0,value,128]]), expected=new Map(column);
    const iterator=bounded ? column.storageEntries({bounded:true}) : column.entries(), oracle=expected.entries();
    const first=iterator.next();
    if(bounded) { assert.deepEqual(first.value,[0,value,128]); for(let i=0;i<128;i++)oracle.next(); }
    else assert.deepEqual(first,oracle.next());
    const snapshot=column.shareData();
    if(useRun)column.setRun(128,3,value);else column.set(128,value);
    for(let r=128;r<(useRun?131:129);r++)expected.set(r,value);
    const remaining=[...iterator].flatMap(([r,cell,count=1])=>Array.from({length:count},(_,offset)=>[r+offset,cell]));
    assert.deepEqual(remaining,[...oracle]); assert.equal(snapshot.size,128); assert.equal(snapshot.has(128),false);
    checkSame(column,expected); assert.equal(column.readers,0);
  }
});

test('iterator beginning on shared compressed backing isolates its later end extension from other wrappers', () => {
  const value=getSharedBlankCell(styles[0]);
  for(const bounded of [false,true]) {
    const column=RunColumn.fromSortedStorage(cell=>cell===value,[[0,value,128]]), snapshot=column.shareData();
    const iterator=bounded ? column.storageEntries({bounded:true}) : column.entries();
    assert.equal(iterator.next().done,false); column.set(128,value);
    const remaining=[...iterator].flatMap(([r,cell,count=1])=>Array.from({length:count},(_,offset)=>[r+offset,cell]));
    assert.equal(remaining.at(-1)?.[0],128); assert.equal(snapshot.has(128),false); assert.equal(snapshot.size,128);
    assert.notEqual(column.dataKey,snapshot.dataKey); assert.equal(column.readers,0);
  }
});

test('COW clear and reinsertion during a live compressed iterator preserves native visibility', () => {
  const { cells } = fixture(30, 3), a = cells.col(0), before = new Map(cells.col(1)), expected = new Map(a);
  const it = a.entries(), oracle = expected.entries(); assert.deepEqual(it.next(), oracle.next());
  a.clear(); expected.clear(); const value = { raw: 'new' };
  a.set(40, value); expected.set(40, value); a.set(0, value); expected.set(0, value);
  assert.deepEqual([...it], [...oracle]); checkSame(cells.col(1), before);
});

test('random mixed source and later edits match ordinary CellMap coordinates, ordering and candidate indexes', () => {
  const cells = new CellMap(), ordinary = new CellMap(), collector = createImportedBlankRuns(cells);
  let seed = 733; const random = n => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) % n);
  for (let r = 0; r < 48; r++) for (let c = 0; c < 64; c++) {
    const choice = random(10), style = styles[random(2)];
    if (choice < 7) { collector.add(r, c, style); ordinary.setRC(r, c, getSharedBlankCell(style)); }
    else { const value = choice === 7 ? makeCellRC({ raw: '=A1', style }, r, c) : { raw: String(r + c), style }; cells.setRC(r, c, value); ordinary.setRC(r, c, value); }
  }
  collector.finish();
  const same = () => { assert.equal(cells.size, ordinary.size); assert.deepEqual([...cells], [...ordinary]); assert.deepEqual([...cells.formulaEntries()], [...ordinary.formulaEntries()]); };
  same();
  for (let i = 0; i < 300; i++) {
    const r = random(55), c = random(64);
    if (random(4) === 0) { cells.deleteRC(r, c); ordinary.deleteRC(r, c); }
    else { const value = random(3) ? getSharedBlankCell(styles[random(2)]) : { raw: 'edit' + i }; cells.setRC(r, c, value); ordinary.setRC(r, c, value); }
    if (i % 30 === 0) same();
  }
  same();
});


test('one million rows across three styled blank columns remain one rectangle and one shared run', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  for (let r = 0; r < 1_000_000; r++) for (let c = 0; c < 3; c++) collector.add(r, c, styles[0]);
  assert.equal(cells.size, 3_000_000); const stats = collector.finish();
  assert.equal(stats.horizontalSpans, 1_000_000); assert.equal(stats.rectangles, 1); assert.equal(stats.uniquePatterns, 1);
  assert.equal(cells.col(0).runs.length, 1); assert.equal(cells.col(0).runs, cells.col(2).runs);
  assert.equal(cells.size, 3_000_000); assert.equal(cells.getRC(999999, 2).style, styles[0]); assert.equal(cells.hasRC(1000000, 2), false);
  cells.setRC(999998, 0, { raw: 'edit' }); assert.equal(cells.getRC(999998, 2).raw, ''); assert.equal(cells.getRC(999998, 0).raw, 'edit');
});

test('rectangle boundaries preserve style changes, skipped rows, changing widths and point overrides', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells), ordinary = new CellMap();
  for (let r = 0; r < 28; r++) for (let c = 0; c < 12; c++) {
    if (r === 12 || r === 16 && c > 5) continue;
    if (r === 9 && c === 3 || r === 22 && c === 7) { const cell = { raw: String(r), comment: 'point' }; cells.setRC(r, c, cell); ordinary.setRC(r, c, cell); }
    else { const style = styles[r < 8 || r >= 20 ? 0 : 1]; collector.add(r, c, style); ordinary.setRC(r, c, getSharedBlankCell(style)); }
  }
  const stats = collector.finish(); assert.ok(stats.rectangles < stats.horizontalSpans); assert.deepEqual([...cells], [...ordinary]); assert.equal(cells.size, ordinary.size);
  assert.equal(cells.hasRC(12, 5), false); assert.equal(cells.hasRC(16, 9), false); assert.equal(cells.getRC(9, 3).comment, 'point');
});

test('rectangle flush before nonmonotonic XML retains established row order and subsequent duplicate writes', () => {
  const cells = new CellMap(), collector = createImportedBlankRuns(cells);
  for (let r = 3; r < 13; r++) for (let c = 0; c < 4; c++) collector.add(r, c, styles[0]);
  collector.flush(); cells.setRC(0, 0, { raw: 'unordered' }); collector.add(4, 0, styles[1]);
  assert.equal(cells.size, 41); assert.deepEqual([...cells.col(0).keys()], [3,4,5,6,7,8,9,10,11,12,0]); assert.equal(cells.getRC(4,0).style, styles[1]);
});
