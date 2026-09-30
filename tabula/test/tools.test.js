// 데이터 도구: 통합 (Consolidate) · 다단계 정렬 (셀 색 · 사용자 지정 목록 · 자연 정렬 · 대/소문자 · 왼쪽→오른쪽)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { consolidate } from '../src/analysis.js';

const put = (wb, rows) => wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== undefined && v !== null) wb.setInput(0, r, c, String(v)); })));
const col = (wb, c, r1, r2) => Array.from({ length: r2 - r1 + 1 }, (_, i) => wb.getValue(0, r1 + i, c));

test('통합: 첫 행 · 왼쪽 열 이름으로 합계 · 평균 · 개수', () => {
  const a = [[null, '1월', '2월'], ['서울', 10, 20], ['부산', 5, 5]];
  const b = [[null, '2월', '3월'], ['부산', 1, 2], ['대구', 7, null], ['서울', 3, 4]];
  const s = consolidate([a, b], { fn: 'sum', topRow: true, leftCol: true });
  assert.deepEqual(s.rows, [[null, '1월', '2월', '3월'], ['서울', 10, 23, 4], ['부산', 5, 6, 2], ['대구', null, 7, null]]);
  const avg = consolidate([a, b], { fn: 'average', topRow: true, leftCol: true });
  assert.equal(avg.rows[1][2], 11.5);
  const cnt = consolidate([a, b], { fn: 'count', topRow: true, leftCol: true });
  assert.equal(cnt.rows[1][2], 2);
  // 위치로 통합
  const p = consolidate([[[1, 2], [3, 4]], [[10, 20]]], { fn: 'sum' });
  assert.deepEqual(p.rows, [[11, 22], [3, 4]]);
  // 이름은 대/소문자 무시
  assert.deepEqual(consolidate([[['A', 1]], [['a', 2]]], { leftCol: true }).rows, [['A', 3]]);
});

test('다단계 정렬: 두 기준 · 사용자 지정 목록 · 셀 색', () => {
  const wb = new Workbook();
  put(wb, [['화', 2], ['월', 1], ['수', 3], ['월', 0], ['일', 5]]);
  wb.transact(() => wb.sortMulti(0, 0, 0, 4, 1, [{ at: 0, asc: true, list: ['일', '월', '화', '수', '목', '금', '토'] }, { at: 1, asc: false }]));
  assert.deepEqual(col(wb, 0, 0, 4), ['일', '월', '월', '화', '수']);
  assert.deepEqual(col(wb, 1, 0, 4), [5, 1, 0, 2, 3]);
  const wb2 = new Workbook();
  put(wb2, [['a'], ['b'], ['c']]);
  wb2.transact(() => wb2.setCellData(0, 2, 0, { raw: 'c', style: { fill: '#FFFF00' } }));
  wb2.transact(() => wb2.sortMulti(0, 0, 0, 2, 0, [{ at: 0, on: 'fill', color: '#ffff00', asc: true }]));
  assert.deepEqual(col(wb2, 0, 0, 2), ['c', 'a', 'b']);
});

test('정렬 옵션: 자연 정렬 · 대/소문자 구분 · 왼쪽에서 오른쪽 · 수식 이동', () => {
  const wb = new Workbook();
  put(wb, [['항목10'], ['항목2'], ['항목1']]);
  wb.transact(() => wb.sortMulti(0, 0, 0, 2, 0, [{ at: 0, asc: true }], { natural: true }));
  assert.deepEqual(col(wb, 0, 0, 2), ['항목1', '항목2', '항목10']);
  const wb2 = new Workbook();
  put(wb2, [['B'], ['a'], ['A'], ['b']]);
  wb2.transact(() => wb2.sortMulti(0, 0, 0, 3, 0, [{ at: 0, asc: true }], { caseSensitive: true }));
  assert.deepEqual(col(wb2, 0, 0, 3), ['a', 'A', 'b', 'B']);
  const wb3 = new Workbook();
  put(wb3, [[3, 1, 2], ['=A1*10', '=B1*10', '=C1*10']]);
  wb3.transact(() => wb3.sortMulti(0, 0, 0, 1, 2, [{ at: 0, asc: true }], { byCols: true }));
  assert.deepEqual([0, 1, 2].map((c) => wb3.getValue(0, 0, c)), [1, 2, 3]);
  assert.deepEqual([0, 1, 2].map((c) => wb3.getValue(0, 1, c)), [10, 20, 30]);
  assert.equal(wb3.getCell(0, 1, 0).raw, '=A1*10');
});

test('연결된 그림: xlsx 로 저장하고 다시 열어도 원본 범위 유지 · 시트 이름 바꾸기 따라감', async () => {
  const { writeXlsx, readXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  put(wb, [['a', 1], ['b', 2]]);
  wb.transact(() => wb.setSheetProp(0, 'images', [{ id: 'im1', name: '연결된 그림', x: 10, y: 10, w: 120, h: 40, linked: { sheet: wb.sheets[0].name, r1: 0, c1: 0, r2: 1, c2: 1 } }]));
  wb.transact(() => wb.renameSheet(0, '원본 표'));
  assert.equal(wb.sheets[0].images[0].linked.sheet, '원본 표');
  const { data } = readXlsx(writeXlsx(wb));
  const im = data.sheets[0].images[0];
  assert.deepEqual(im.linked, { sheet: '원본 표', r1: 0, c1: 0, r2: 1, c2: 1 });
  assert.ok(im.src.startsWith('data:image/png'));
});

test('연결된 그림: 원본 시트에 행을 넣으면 범위가 따라 내려감', () => {
  const wb = new Workbook();
  put(wb, [['a', 1], ['b', 2]]);
  wb.transact(() => wb.setSheetProp(0, 'images', [{ id: 'im1', x: 10, y: 10, w: 120, h: 40, linked: { sheet: wb.sheets[0].name, r1: 0, c1: 0, r2: 1, c2: 1 } }]));
  wb.transact(() => wb.insertRows(0, 0, 2));
  assert.deepEqual({ ...wb.sheets[0].images[0].linked, sheet: '' }, { sheet: '', r1: 2, c1: 0, r2: 3, c2: 1 });
});
