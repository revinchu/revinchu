import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { CellMap } from '../src/cellmap.js';
import { DepGraph } from '../src/depgraph.js';

function tableBook(n, formula, extra = {}) {
  const cells = new CellMap();
  cells.setRC(0, 0, { raw: 'Value' }); cells.setRC(0, 1, { raw: 'Result' });
  for (let r = 1; r <= n; r++) {
    cells.setRC(r, 0, { raw: String(r) });
    cells.setRC(r, 1, { raw: typeof formula === 'function' ? formula(r) : formula });
  }
  return new Workbook({ ...extra, sheets: [{ name: 'Data', cells,
    tables: [{ name: 'T', r1: 0, c1: 0, r2: n, c2: 1, header: true }] }] });
}
function runs(graph) {
  const result = new Set();
  for (const cols of graph.cols) if (cols) for (const bucket of cols.values()) for (const run of bucket.runs) result.add(run);
  for (const list of graph.wide) if (list) for (const run of list) result.add(run);
  return [...result];
}
const dependents = (graph, r, c = 0) => graph.dependentsOf(0, r, c).map(p => `${p.r},${p.c}`).sort();

test('four structured row references share four graph runs across 10,000 fixed-dr formulas', async () => {
  const n = 10000, wb = tableBook(n, '=IFS(AND(T[@Value]>=1,T[@Value]<4),1,AND(T[@Value]>=4,T[@Value]<8),2,TRUE,3)');
  assert.equal(wb.getCell(0, 1, 1).ast, wb.getCell(0, n, 1).ast);
  assert.equal(wb.getCell(0, n, 1).dr ?? 0, 0);
  await wb.prepareGraph();
  assert.equal(runs(wb.graph).length, 4);
  for (const run of runs(wb.graph)) { assert.equal(run.n, n); assert.equal(run.drs, 0); }
  for (const r of [1, 5000, n]) assert.deepEqual(dependents(wb.graph, r), [`${r},1`]);
});

test('A1 fills and mixed A1/structured references retain exact dependencies in growing-dr runs', () => {
  for (const formula of [r => `=A${r + 1}*2`, r => `=A${r + 1}+[@Value]`]) {
    const wb = tableBook(50, formula), graph = new DepGraph(wb);
    assert.ok(runs(graph).length <= 2);
    for (const run of runs(graph)) { assert.equal(run.n, 50); assert.equal(run.drs, 1); }
    for (const r of [1, 25, 50]) assert.deepEqual(dependents(graph, r), [`${r},1`]);
  }
});

test('partial formula replacement, literal replacement and undo do not reuse stale run dependencies', async () => {
  const wb = tableBook(20, '=[@Value]*2'); await wb.prepareGraph();
  assert.equal(wb.getValue(0, 10, 1), 20); assert.equal(wb.getValue(0, 11, 1), 22);
  wb.transact(() => wb.setInput(0, 10, 1, '=A2*3'));
  assert.deepEqual(dependents(wb.graph, 10), []);
  assert.deepEqual(dependents(wb.graph, 1), ['1,1', '10,1']);
  wb.transact(() => wb.setInput(0, 1, 0, '100'));
  assert.equal(wb.getValue(0, 10, 1), 300); assert.equal(wb.getValue(0, 11, 1), 22);
  wb.undo(); wb.undo();
  assert.deepEqual(dependents(wb.graph, 10), ['10,1']); assert.equal(wb.getValue(0, 10, 1), 20);
  wb.transact(() => wb.setInput(0, 10, 1, 'literal'));
  assert.deepEqual(dependents(wb.graph, 10), []);
  wb.undo(); assert.deepEqual(dependents(wb.graph, 10), ['10,1']);
});

test('fixed names and structured references merge without losing fixed-range dependents', () => {
  const wb = tableBook(30, '=SUM(Fixed)+[@Value]', { names: [{ name: 'Fixed', ref: '=Data!$D$1:$D$3', sheet: null }] });
  const graph = new DepGraph(wb);
  assert.equal(runs(graph).length, 2);
  assert.deepEqual(dependents(graph, 12), ['12,1']);
  assert.equal(dependents(graph, 0, 3).length, 30);
  assert.deepEqual(dependents(graph, 3, 3), []);
});

test('wide fixed ranges retain every dependent alongside structured current-row references', () => {
  const wb = tableBook(30, '=SUM($D$1:$AZ$1)+[@Value]'), graph = new DepGraph(wb);
  assert.equal(graph.wide[0].length, 1); assert.equal(graph.wide[0][0].n, 30);
  assert.equal(dependents(graph, 0, 40).length, 30);
  assert.deepEqual(dependents(graph, 1, 40), []);
  assert.deepEqual(dependents(graph, 15), ['15,1']);
});

test('run verification still rejects same-AST cells whose exact dr or dc changed', () => {
  const wb = tableBook(10, '=[@Value]*2'), graph = new DepGraph(wb), cells = wb.sheets[0].cells;
  const original = cells.getRC(5, 1);
  for (const offset of [{ dr: 1 }, { dc: 1 }]) {
    cells.setRC(5, 1, { ...original, ...offset });
    assert.deepEqual(dependents(graph, 5), []);
    assert.deepEqual(dependents(graph, 6), ['6,1']);
  }
  cells.setRC(5, 1, original); assert.deepEqual(dependents(graph, 5), ['5,1']);
});

test('changes in offset slope split runs rather than accepting unrelated same-AST formulas', () => {
  const wb = tableBook(8, '=A1*2'), cells = wb.sheets[0].cells;
  const ast = cells.getRC(1, 1).ast;
  const offsets = [0, 0, 2, 3, 4, 6, 6, 6];
  for (let i = 0; i < offsets.length; i++) cells.setRC(i + 1, 1, { ...cells.getRC(i + 1, 1), ast, dr: offsets[i], dc: 0 });
  const graph = new DepGraph(wb);
  assert.equal(runs(graph).length, 3);
  for (let source = 0; source < 8; source++) {
    const expected = offsets.flatMap((dr, i) => dr === source ? [`${i + 1},1`] : []).sort();
    assert.deepEqual(dependents(graph, source), expected);
  }
});
