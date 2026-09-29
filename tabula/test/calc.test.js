import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { CellMap } from '../src/cellmap.js';

const fill = (wb, si, rows) => wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== null) wb.setInput(si, r, c, String(v)); })));

test('CellMap: Map 과 같은 사용법 + 행·열 빠른 접근', () => {
  const m = new CellMap();
  m.set('3,1', 'a');
  m.setRC(10, 1, 'b');
  m.set('2,5', 'c');
  assert.equal(m.size, 3);
  assert.equal(m.get('10,1'), 'b');
  assert.equal(m.getRC(3, 1), 'a');
  assert.ok(m.has('2,5'));
  m.delete('3,1');
  assert.equal(m.size, 2);
  assert.deepEqual([...m.keys()].sort(), ['10,1', '2,5']);
  assert.deepEqual(new Map(m), new Map([['10,1', 'b'], ['2,5', 'c']]));
});

test('공유 수식: 아래로 채운 수식은 AST 하나를 공유하고 위치별로 계산', () => {
  const wb = new Workbook();
  const rows = [['a', 'b', 'c']];
  for (let i = 1; i <= 50; i++) rows.push([i, i * 2, `=A${i + 1}*B${i + 1}+$A$2`]);
  fill(wb, 0, rows);
  const c2 = wb.getCell(0, 1, 2);
  const c50 = wb.getCell(0, 50, 2);
  assert.equal(c2.ast, c50.ast); // 같은 AST
  assert.equal(wb.getValue(0, 1, 2), 1 * 2 + 1);
  assert.equal(wb.getValue(0, 50, 2), 50 * 100 + 1);
  // 열 전체 · 한글 이름이 섞여도 잘못 공유하지 않음
  wb.transact(() => { wb.setInput(0, 0, 5, '=SUM(A:A)'); wb.setInput(0, 1, 5, '=SUM(B:B)'); });
  assert.equal(wb.getValue(0, 0, 5), 1275);
  assert.equal(wb.getValue(0, 1, 5), 2550);
  // 행 삽입 뒤에도 옮겨진 수식이 맞는 칸을 참조
  wb.transact(() => wb.insertRows(0, 10, 2));
  assert.equal(wb.getRaw(0, 52, 2), '=A53*B53+$A$2');
  assert.equal(wb.getValue(0, 52, 2), 50 * 100 + 1);
  assert.equal(wb.getValue(0, 5, 2), 5 * 10 + 1);
});

test('셀 단위 재계산: 바뀐 칸을 참조하는 수식만 다시 계산 (사슬 · 범위 · 다른 시트)', () => {
  const wb = new Workbook();
  const rows = [['x', 'run', 'sq']];
  for (let i = 1; i <= 200; i++) rows.push([i, i === 1 ? '=A2' : `=B${i}+A${i + 1}`, `=A${i + 1}^2`]);
  fill(wb, 0, rows);
  const at = wb.transact(() => wb.addSheet('요약'));
  wb.transact(() => { wb.setInput(at, 0, 0, '=SUM(Sheet1!C2:C201)'); wb.setInput(at, 1, 0, '=Sheet1!B201'); wb.setInput(at, 2, 0, '=Sheet1!A2*100'); });
  assert.equal(wb.getValue(0, 200, 1), 20100);
  assert.equal(wb.getValue(at, 0, 0), (200 * 201 * 401) / 6);
  assert.equal(wb.getValue(at, 2, 0), 100);
  const sq5 = wb.getValue(0, 5, 2);
  // A101 을 바꾸면: 누계(B101…B201), C101, 요약의 합계 · 마지막 누계가 바뀜. C5 는 그대로(캐시 유지)
  wb.transact(() => wb.setInput(0, 100, 0, '1000'));
  assert.ok(wb.caches[0].hasRC(5, 2), '관계없는 수식의 계산 결과는 유지');
  assert.equal(wb.getValue(0, 5, 2), sq5);
  assert.equal(wb.getValue(0, 200, 1), 20100 - 100 + 1000);
  assert.equal(wb.getValue(at, 0, 0), (200 * 201 * 401) / 6 - 100 * 100 + 1000 * 1000);
  assert.equal(wb.getValue(at, 1, 0), 20100 - 100 + 1000);
  // 실행 취소도 같은 방식
  wb.undo();
  assert.equal(wb.getValue(at, 1, 0), 20100);
  assert.equal(wb.getValue(0, 100, 2), 10000);
  // 수식을 바꾸면 새 참조로 이어짐
  wb.transact(() => wb.setInput(at, 2, 0, '=Sheet1!A3*100'));
  wb.transact(() => wb.setInput(0, 2, 0, '7'));
  assert.equal(wb.getValue(at, 2, 0), 700);
});

test('큰 범위 캐시: 다른 열을 고치면 유지, 같은 열을 고치면 다시 읽음', () => {
  const wb = new Workbook();
  const rows = [['v', 'w']];
  for (let i = 1; i <= 6000; i++) rows.push([i, 1]);
  fill(wb, 0, rows);
  wb.transact(() => { wb.setInput(0, 0, 4, '=SUM(A2:A6001)'); wb.setInput(0, 1, 4, '=COUNTIF(A2:A6001,">3000")'); });
  assert.equal(wb.getValue(0, 0, 4), 6000 * 6001 / 2);
  assert.equal(wb.getValue(0, 1, 4), 3000);
  const cached = [...wb.rangeMemo.values()][0].rows;
  wb.transact(() => wb.setInput(0, 7, 1, '5'));
  wb.transact(() => wb.setInput(0, 2, 4, '=MAX(A2:A6001)'));
  assert.equal(wb.getValue(0, 2, 4), 6000);
  assert.equal([...wb.rangeMemo.values()][0].rows, cached, '다른 열 편집 후에도 같은 배열');
  wb.transact(() => wb.setInput(0, 7, 0, '100000'));
  assert.equal(wb.getValue(0, 0, 4), 6000 * 6001 / 2 - 7 + 100000);
  assert.equal(wb.getValue(0, 2, 4), 100000);
});

test('파일 계산 결과: 입력이 바뀐 수식만 다시 계산하고 나머지는 파일 값 유지', () => {
  const wb = new Workbook({ sheets: [{ name: 'S', fileValues: true, cells: {
    '0,0': { raw: '1' }, '1,0': { raw: '=A1+1', cached: 2 }, '2,0': { raw: '=NOTAFUNC(1)', cached: 42 }, '3,0': { raw: '=A2*10', cached: 20 },
  } }] });
  assert.equal(wb.getValue(0, 2, 0), 42);
  wb.transact(() => wb.setInput(0, 0, 0, '5'));
  assert.equal(wb.getValue(0, 1, 0), 6);
  assert.equal(wb.getValue(0, 3, 0), 60);
  assert.equal(wb.getValue(0, 2, 0), 42, '관계없는 (지원하지 않는 함수) 수식은 파일 값 그대로');
});

test('표 계산 열([@열])과 이름 범위도 셀 단위로 따라감', () => {
  const wb = new Workbook();
  fill(wb, 0, [['비용', '클릭', 'CPC'], [100, 10, null], [300, 20, null], [50, 0, null]]);
  wb.transact(() => {
    wb.setSheetProp(0, 'tables', [{ id: 't1', name: '광고', r1: 0, c1: 0, r2: 3, c2: 2, header: true, totals: false, style: 'TableStyleMedium2', filter: null }]);
    for (let r = 1; r <= 3; r++) wb.setInput(0, r, 2, '=IFERROR([@비용]/[@클릭],0)');
    wb.setNames([{ name: '총비용', ref: '=Sheet1!$A$2:$A$4', sheet: null }]);
    wb.setInput(0, 5, 0, '=SUM(총비용)');
    wb.setInput(0, 6, 0, '=SUM(광고[CPC])');
  });
  assert.equal(wb.getValue(0, 5, 0), 450);
  assert.equal(wb.getValue(0, 6, 0), 10 + 15);
  wb.transact(() => wb.setInput(0, 3, 1, '5'));
  assert.equal(wb.getValue(0, 3, 2), 10);
  assert.equal(wb.getValue(0, 6, 0), 35);
  wb.transact(() => wb.setInput(0, 1, 0, '200'));
  assert.equal(wb.getValue(0, 5, 0), 550);
  assert.equal(wb.getValue(0, 6, 0), 45);
});

test('동적 수식(INDIRECT · OFFSET)은 무엇이 바뀌어도 다시 계산', () => {
  const wb = new Workbook();
  fill(wb, 0, [[1], [2], [3], ['=SUM(INDIRECT("A1:A3"))', '=SUM(OFFSET(A1,0,0,3,1))']]);
  assert.equal(wb.getValue(0, 3, 0), 6);
  wb.transact(() => wb.setInput(0, 1, 0, '20'));
  assert.equal(wb.getValue(0, 3, 0), 24);
  assert.equal(wb.getValue(0, 3, 1), 24);
});

test('GETPIVOTDATA: 피벗이 있는 시트가 바뀔 때만 다시 계산 (다른 칸 편집은 무관)', async () => {
  const { DepGraph } = await import('../src/depgraph.js');
  const wb = new Workbook();
  wb.transact(() => wb.addSheet('피벗'));
  fill(wb, 0, [[1, '=IFERROR(GETPIVOTDATA("매출",피벗!$A$1,"지역",A1),0)'], [2, '=IFERROR(GETPIVOTDATA("매출",피벗!$A$1,"지역",A2),0)']]);
  const g = new DepGraph(wb);
  assert.equal(g.dyn.size, 0);
  const pts = (arr) => { const out = g.propagate(arr) ?? []; const s = new Set(); for (let i = 0; i < out.length; i += 3) s.add(`${out[i]}:${out[i + 1]},${out[i + 2]}`); return s; };
  assert.deepEqual(pts([0, 5, 5]), new Set()); // 관계없는 칸
  assert.deepEqual(pts([1, 40, 7]), new Set(['0:0,1', '0:1,1'])); // 피벗 시트의 아무 칸
  assert.equal(g.propagate([1, 40, 7], 1), null); // 한도를 넘으면 시트 단위로
});
