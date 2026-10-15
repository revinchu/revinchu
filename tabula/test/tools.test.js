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

test('추천 피벗 테이블: 범주 × 숫자 합계 · 교차 · 개수 후보', async () => {
  const { recommendPivots } = await import('../src/pivot.js');
  const recs = recommendPivots(['지역', '채널', '주문번호', '매출', '수량'], [
    ['서울', '부산', '서울', '대구'], ['검색', '영상', '검색', '검색'], [1, 2, 3, 4], [100, 50, 30, 20], [1, 2, 1, 3],
  ]);
  assert.ok(recs.length >= 4);
  assert.deepEqual(recs[0].rows.length, 1);
  assert.equal(recs[0].values[0].agg, 'sum');
  assert.ok(!recs.some((r) => r.values.some((v) => v.field === '주문번호')), 'ID 같은 열은 합계하지 않음');
  assert.ok(recs.some((r) => r.cols.length === 1), '두 범주 교차');
  assert.deepEqual(recommendPivots(['이름'], [['a', 'b', 'a']])[0].values[0], { field: '이름', agg: 'count' });
});

test('피벗 계산 항목: 수도권 = 서울 + 인천 (다른 필드 · 총합계에도 반영), 수식 검사', async () => {
  const { resolvePivot, computePivot, parseCalcItem } = await import('../src/pivot.js');
  const rows = [['지역', '채널', '비용', '연도'], ['서울', '검색', 100, 2026], ['인천', '검색', 50, 2026], ['부산', '영상', 30, 2026], ['서울', '영상', 10, 2026]];
  const def = { rows: ['지역'], cols: [], values: [{ field: '비용', agg: 'sum' }], calcItems: { 지역: [{ name: '수도권', formula: "서울+'인천'" }, { name: '서울 절반', formula: '서울*0.5' }] } };
  const res = resolvePivot(rows, def);
  const { grid } = computePivot(res, res.def);
  const map = Object.fromEntries(grid.filter((r) => /^(rowItem|grandLabel)/.test(r[0]?.role ?? '')).map((r) => [String(r[0].raw).replace(/^'/, ''), Number(r[r.length - 1].raw)]));
  assert.equal(map['수도권'], 160);
  assert.equal(map['서울 절반'], 55);
  assert.equal(map['서울'], 110);
  assert.deepEqual(parseCalcItem("서울 - '부산'/2"), [{ item: '서울', k: 1 }, { item: '부산', k: -0.5 }]);
  assert.throws(() => parseCalcItem('서울*부산'), /항목끼리/);
  assert.throws(() => parseCalcItem('100'), /상수/);
  // 연도(숫자 차원)는 계수를 곱하지 않음
  const res2 = resolvePivot(rows, { ...def, rows: ['연도'], calcItems: { 지역: [{ name: '서울 절반', formula: '서울*0.5' }] } });
  const g2 = computePivot(res2, res2.def).grid;
  assert.ok(g2.some((r) => String(r[0]?.raw) === '2026'));
});

test('피벗 계산 항목: xlsx 저장 후 다시 열어도 유지', async () => {
  const { writeXlsx, readXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  put(wb, [['지역', '비용'], ['서울', 10], ['인천', 5]]);
  const def = { name: '피벗 테이블1', source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '비용', agg: 'sum' }], top: 0, left: 4, calcItems: { 지역: [{ name: '수도권 "합"', formula: '서울+인천' }] } };
  wb.transact(() => wb.setSheetProp(0, 'pivot', def));
  const { data } = readXlsx(writeXlsx(wb));
  const back = data.sheets.flatMap((s) => [s.pivot, ...(s.pivotsExtra ?? [])]).find(Boolean);
  assert.deepEqual(back.calcItems, def.calcItems);
});

test('이상치 찾기 · 광고 지표 열 찾기', async () => {
  const { detectAnomalies, marketingMetrics } = await import('../src/analysis.js');
  const v = [100, 104, 98, 101, 97, 103, 99, 450, 102, 20, 'x', null];
  for (const method of ['mad', 'zscore', 'iqr']) {
    const r = detectAnomalies(v, { method, threshold: method === 'zscore' ? 2 : undefined });
    assert.ok(r.hits.some((h) => h.i === 7 && h.dir === 'up'), method);
  }
  assert.ok(detectAnomalies(v, { method: 'mad' }).hits.some((h) => h.i === 9 && h.dir === 'down'));
  assert.equal(detectAnomalies([1, 2, 3], {}).hits.length, 0);
  const m = marketingMetrics(['날짜', '노출수', '클릭수', '비용', '전환수', '전환매출', 'CTR']);
  assert.deepEqual(m.list.map((x) => x.name), ['CPC', 'CPM', 'CVR', 'CPA', 'ROAS']);
  assert.equal(m.list[0].formula((k) => ({ cost: 'D2', click: 'C2' })[k]), '=IFERROR(D2/C2,"")');
});

test('이상치 조건부 서식 수식이 JS 판정과 같음', async () => {
  const { detectAnomalies, anomalyFormula } = await import('../src/analysis.js');
  const v = [100, 104, 98, 101, 97, 103, 99, 450, 102, 20];
  for (const [method, t] of [['mad', 3.5], ['zscore', 2], ['iqr', 1.5]]) {
    const wb = new Workbook();
    put(wb, v.map((x) => [x]));
    const f = anomalyFormula(method, t, 'A1', '$A$1:$A$10');
    wb.transact(() => v.forEach((_, i) => wb.setInput(0, i, 1, f.replace(/\bA1\b(?!:)/g, `A${i + 1}`))));
    const got = v.map((_, i) => wb.getValue(0, i, 1));
    const want = v.map((_, i) => detectAnomalies(v, { method, threshold: t }).hits.some((h) => h.i === i));
    assert.deepEqual(got, want, method);
  }
});

test('피벗 선택 항목 그룹화: 계산 · xlsx 왕복 (discretePr)', async () => {
  const { writeXlsx, readXlsx } = await import('../src/xlsx.js');
  const { resolvePivot, computePivot, pivotSourceData } = await import('../src/pivot.js');
  const wb = new Workbook();
  put(wb, [['상품', '비용'], ['A', 10], ['B', 20], ['C', 5], ['D', 1]]);
  const def = { name: '피벗 테이블1', source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 4, c2: 1 }, rows: ['상품2', '상품'], cols: [], values: [{ field: '비용', agg: 'sum' }], top: 0, left: 4, groups: { 상품2: { by: 'items', base: '상품', map: { A: '그룹1', B: '그룹1' } } } };
  const grid = (w, d) => { const res = resolvePivot(pivotSourceData(w, d), d); return computePivot(res, res.def).grid.map((r) => [String(r[0]?.raw ?? '').replace(/^'/, ''), r[r.length - 1]?.raw]); };
  const g1 = grid(wb, def);
  assert.ok(g1.some(([a, b]) => a === '그룹1' && Number(b) === 30));
  wb.transact(() => wb.setSheetProp(0, 'pivot', def));
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  const d2 = back.sheets[0].pivot;
  assert.deepEqual(d2.groups.상품2, def.groups.상품2);
  const sorted = (g) => g.map((r) => r.join('=')).sort();
  assert.deepEqual(sorted(grid(back, d2)), sorted(g1));
});

test('계산 필드: 후위 % · 숫자 글자 결과는 숫자, GETPIVOTDATA 월 그룹에 숫자 항목', async () => {
  const { resolvePivot, computePivot, pivotLookup, parseCalc } = await import('../src/pivot.js');
  assert.ok(parseCalc('IFERROR(클릭/노출,"0"%)'));
  const rows = [['일자', '노출', '클릭'], [46000, 100, 5], [46001, 0, 0], [46040, 50, 1]];
  const def = { rows: ['일자'], cols: [], values: [{ field: 'CTR', agg: 'sum' }, { field: '클릭', agg: 'sum' }], calcFields: [{ name: 'CTR', formula: 'IFERROR(클릭/노출,"0"%)' }, { name: 'Z', formula: 'IFERROR(클릭/0,"0")' }], groups: { 일자: { by: 'months' } } };
  const res = resolvePivot(rows, def);
  const g = computePivot(res, res.def).grid;
  assert.ok(g.every((r) => !r.some((x) => x?.raw === '#NAME?')));
  const z = resolvePivot(rows, { ...def, values: [{ field: 'Z', agg: 'sum' }] });
  assert.ok(computePivot(z, z.def).grid.some((r) => r[r.length - 1]?.raw === '0'));
  const month = new Date(Date.UTC(1899, 11, 30) + 46000 * 864e5).getUTCMonth() + 1;
  assert.equal(pivotLookup(rows, def, '클릭', [['일자', month]], res), 5);
});
