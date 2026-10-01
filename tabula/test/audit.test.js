import { test } from 'node:test';
import assert from 'node:assert/strict';
import { evalSteps, goalSeek, dataTable, specialCells } from '../src/audit.js';
import { parse } from '../src/formula.js';
import { Workbook } from '../src/workbook.js';

test('수식 계산 단계: 안쪽부터 값으로 바뀜', () => {
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '3'); wb.setInput(0, 0, 1, '4'); });
  const src = 'IF(A1>2,SUM(A1,B1)*2,0)';
  const steps = evalSteps(src, parse(src), wb.ctxFor(0, 5, 5));
  const texts = steps.map((s) => s.text);
  assert.equal(texts[0], '=IF(A1>2,SUM(A1,B1)*2,0)');
  assert.deepEqual(steps[0].mark, [4, 6]); // A1
  assert.ok(texts.includes('=IF(TRUE,SUM(3,4)*2,0)'));
  assert.ok(texts.includes('=IF(TRUE,14,0)'));
  assert.equal(texts.at(-1), '14');
});

test('목표값 찾기: 비선형 식도 수렴', () => {
  const r = goalSeek((x) => x * x * x - 2 * x, 1, 20);
  assert.ok(r.ok);
  assert.ok(Math.abs(r.x ** 3 - 2 * r.x - 20) < 0.001);
  // 워크북: 대출 월 상환액 = -1000 이 되는 원금
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '10000'); wb.setInput(0, 1, 0, '=PMT(0.05/12,36,A1)'); });
  const res = goalSeek((x) => { wb.setCellData(0, 0, 0, { raw: String(x) }); return wb.getValue(0, 1, 0); }, 10000, -1000);
  assert.ok(res.ok);
  assert.ok(Math.abs(wb.getValue(0, 1, 0) + 1000) < 0.01);
});

test('데이터 표 · 이동 옵션', () => {
  const vals = new Map([['0,0', null], ['0,1', 1], ['0,2', 2], ['1,0', 10], ['2,0', 20]]);
  let x = 0;
  let y = 0;
  const out = dataTable({ r1: 0, c1: 0, r2: 2, c2: 2 }, 'x', 'y', (k, v) => { if (k === 'x') x = v; else y = v; },
    (r, c) => (r === 0 && c === 0 ? x * y : vals.get(`${r},${c}`)));
  assert.deepEqual(out, [[1, 1, 10], [1, 2, 20], [2, 1, 20], [2, 2, 40]]);
  const cells = { '0,0': { raw: '1' }, '0,1': { raw: '=A1', formula: true }, '1,1': { raw: '', comment: 'x' } };
  const found = specialCells('blanks', { r1: 0, c1: 0, r2: 1, c2: 1 }, { cellAt: (r, c) => cells[`${r},${c}`], valueAt: () => 1 });
  assert.deepEqual(found, [[1, 0], [1, 1]]);
});


test('보이는 셀 선택: 숨긴 행과 열을 모두 제외하고 빈 셀은 포함', () => {
  const rg = { r1: 2, c1: 4, r2: 4, c2: 7 };
  const cells = { '2,4': { raw: '공개' }, '2,5': { raw: '숨긴 열' }, '3,4': { raw: '숨긴 행' } };
  const opts = { cellAt: (r, c) => cells[r + ',' + c], valueAt: () => 1, hidden: (r, c) => r === 3 || c === 5 || c === 7 };
  assert.deepEqual(specialCells('visible', rg, opts), [[2, 4], [2, 6], [4, 4], [4, 6]]);
  assert.deepEqual(specialCells('visible', rg, { ...opts, limit: 3 }), [[2, 4], [2, 6], [4, 4]], '제한은 숨긴 셀을 제외한 선택 개수에 적용');
  assert.deepEqual(specialCells('visible', rg, { ...opts, hidden: () => true }), [], '범위 전체가 숨겨진 경우 선택 없음');
  assert.deepEqual(specialCells('visible', { r1: 2, c1: 5, r2: 4, c2: 5 }, opts), [], '숨긴 열 하나만 선택해도 포함하지 않음');
});

test('보이는 셀 이외의 이동 옵션은 숨김 여부와 관계없이 검색', () => {
  const rg = { r1: 2, c1: 4, r2: 3, c2: 5 };
  const cells = { '2,4': { raw: '공개' }, '2,5': { raw: '숨긴 열' }, '3,4': { raw: '숨긴 행' } };
  const opts = { cellAt: (r, c) => cells[r + ',' + c], valueAt: (r, c) => cells[r + ',' + c]?.raw, hidden: () => true };
  assert.deepEqual(specialCells('constants', rg, opts), [[2, 4], [2, 5], [3, 4]]);
  assert.deepEqual(specialCells('blanks', rg, opts), [[3, 5]]);
  assert.deepEqual(specialCells('visible', rg, { ...opts, hidden: undefined }), [[2, 4], [2, 5], [3, 4], [3, 5]], '숨김 콜백이 없으면 모든 셀 포함');
});
