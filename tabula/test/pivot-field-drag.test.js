import test from 'node:test';
import assert from 'node:assert/strict';
import { pivotFieldRemovePatch } from '../src/pivot-field-drag.js';

test('행·열·값·보고서 필터에서 끌어온 해당 배치 슬롯만 제거한다', () => {
  for (const from of ['rows', 'cols', 'values', 'pages']) {
    const fields = from === 'values' ? [{ field: 'A', agg: 'sum' }, { field: 'B', agg: 'average' }, { field: 'C', agg: 'count' }] : ['A', 'B', 'C'];
    const def = { [from]: fields, filters: { B: ['선택'] }, top: 3, left: 5 };
    assert.deepEqual(pivotFieldRemovePatch(def, { name: 'B', from, index: 1 }), { [from]: [fields[0], fields[2]] });
    assert.equal(def[from], fields);
    assert.equal(fields.length, 3);
  }
});

test('같은 원본 값 필드가 반복되어도 끌어온 집계 한 개만 제거한다', () => {
  const sum = { field: '매출', agg: 'sum', name: '합계' }, average = { field: '매출', agg: 'average', name: '평균' }, max = { field: '매출', agg: 'max', name: '최댓값' };
  const def = { values: [sum, average, max] };
  const patch = pivotFieldRemovePatch(def, { name: '매출', from: 'values', index: 1 });
  assert.deepEqual(patch, { values: [sum, max] });
  assert.equal(patch.values[0], sum);
  assert.equal(patch.values[1], max);
  assert.deepEqual(pivotFieldRemovePatch(def, { name: '매출', from: 'values', index: 2 }), { values: [sum, average] });
});

test('마지막 배치 필드를 제거하면 해당 영역만 빈 배열로 반환한다', () => {
  const def = { rows: ['상품'], cols: ['지역'], values: [{ field: '금액' }], pages: ['기간'] };
  assert.deepEqual(pivotFieldRemovePatch(def, { name: '상품', from: 'rows', index: 0 }), { rows: [] });
  assert.deepEqual(def.rows, ['상품']);
  assert.deepEqual(def.cols, ['지역']);
});

test('전체 필드 목록에서 끌어온 필드는 이미 사용 중이어도 제거 대상이 아니다', () => {
  const def = { rows: ['상품'], values: [{ field: '금액' }] };
  for (const dragged of [null, {}, { name: '상품' }, { name: '금액' }, { name: '상품', index: 0 }, { name: '상품', from: 'fields', index: 0 }]) {
    assert.equal(pivotFieldRemovePatch(def, dragged), null);
  }
});

test('Σ 값 가상 필드는 제거하지 않으며 같은 이름의 실제 필드와 구분한다', () => {
  const def = { rows: ['Σ 값'], values: [{ field: '금액' }] };
  for (const from of ['rows', 'cols', 'values', 'pages']) {
    assert.equal(pivotFieldRemovePatch(def, { name: 'Σ 값', from, index: 0, sigma: true }), null);
  }
  assert.deepEqual(pivotFieldRemovePatch(def, { name: 'Σ 값', from: 'rows', index: 0 }), { rows: [] });
});

test('변경되었거나 유효하지 않은 배치 슬롯은 제거하지 않는다', () => {
  const def = { rows: ['상품'], values: [{ field: '금액' }] };
  for (const dragged of [
    { name: '기간', from: 'rows', index: 0 }, { name: '상품', from: 'rows', index: 1 },
    { name: '상품', from: 'rows', index: -1 }, { name: '상품', from: 'rows', index: .5 },
    { name: '상품', from: 'rows', index: NaN }, { name: '상품', from: 'rows' },
    { name: '금액', from: 'values', index: '0' }, { name: '상품', from: 'unknown', index: 0 },
  ]) assert.equal(pivotFieldRemovePatch(def, dragged), null);
  assert.equal(pivotFieldRemovePatch(null, { name: '상품', from: 'rows', index: 0 }), null);
  assert.equal(pivotFieldRemovePatch({ rows: {} }, { name: '상품', from: 'rows', index: 0 }), null);
});

test('동결된 원본 정의·값 서식·끌기 정보를 변경하지 않고 제거 계획만 만든다', () => {
  const value = Object.freeze({ field: '금액', agg: 'sum', numFmt: Object.freeze({ numFmt: 'number', decimals: 2 }) });
  const def = Object.freeze({ values: Object.freeze([value]), pages: Object.freeze(['기간']), filters: Object.freeze({ 기간: ['이번 주'] }) });
  const dragged = Object.freeze({ name: '금액', from: 'values', index: 0 });
  assert.deepEqual(pivotFieldRemovePatch(def, dragged), { values: [] });
  assert.equal(def.values[0], value);
  assert.deepEqual(def.pages, ['기간']);
});
