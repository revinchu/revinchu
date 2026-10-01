import { test } from 'node:test';
import assert from 'node:assert/strict';
import { checkCalc, computePivot, resolvePivot, pivotLookup, itemText, keyOf } from '../src/pivot.js';

const base = { rows: ['지역'], cols: [], pages: [], values: [{ field: '결과', agg: 'sum' }], layout: 'tabular', errorShow: false };
function calculated(formula, options = {}, input = -1.5) {
  const def = { ...base, ...options, calcFields: [{ name: '결과', formula }] };
  const resolved = resolvePivot([['지역', '금액'], ['서울', input]], def);
  const grid = computePivot(resolved, resolved.def).grid;
  return { raw: grid.find((r) => r[0]?.raw === '서울')[1].raw, resolved, def };
}

test('계산 필드: 음수 반올림은 0에서 먼 방향, 소수 자릿수는 정수로 처리', () => {
  for (const [formula, expected] of [
    ['ROUND(금액,0)', '-2'], ['ROUND(-15,-1)', '-20'], ['ROUND(금액,0.9)', '-2'],
    ['ROUNDUP(금액,0.9)', '-2'], ['ROUNDDOWN(금액,0.9)', '-1'], ['TRUNC(금액,0.9)', '-1'],
  ]) assert.equal(calculated(formula).raw, expected, formula);
});

test('계산 필드: 논리 함수는 결정된 참/거짓 뒤에 있는 오류도 전파', () => {
  for (const formula of ['AND(금액/0,1)', 'AND(0,금액/0)', 'OR(1,금액/0)', 'NOT(금액/0)']) {
    assert.equal(calculated(formula).raw, '#DIV/0!', formula);
    assert.equal(calculated(`IFERROR(${formula},99)`).raw, '99', formula);
  }
  assert.equal(calculated('AND(1,2)').raw, 'TRUE');
  assert.equal(calculated('OR(0,0)').raw, 'FALSE');
  assert.equal(calculated('NOT(0)').raw, 'TRUE');
});

test('계산 필드: 모든 반올림 자릿수 오류와 비정상 숫자는 IFERROR 및 오류 표시 옵션을 따름', () => {
  for (const fn of ['ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TRUNC']) assert.equal(calculated(`IFERROR(${fn}(금액,1/0),99)`).raw, '99', fn);
  for (const f of ['EXP(1000)', 'POWER(-1,0.5)', '1e300*1e300']) {
    assert.equal(calculated(f).raw, '#NUM!', f);
    assert.equal(calculated(`IFERROR(${f},99)`).raw, '99', f);
  }
  assert.equal(calculated('AND(금액/0,1)', { errorShow: true, errorCaption: '' }).raw, '');
  assert.equal(calculated('AND(금액/0,1)', { errorShow: true, errorCaption: '-' }).raw, '0');
  const p = calculated('AND(금액/0,1)', { errorShow: false, errorCaption: '-' });
  assert.equal(p.raw, '#DIV/0!');
  assert.equal(pivotLookup(null, p.def, '결과', [['지역', '서울']], p.resolved).code, '#DIV/0!');
});

test('계산 필드: 인수 개수 오류는 입력 단계에서 안내하고 가져온 수식도 예외 없이 오류를 표시', () => {
  for (const f of ['ROUND()', 'ROUND(금액)', 'IFERROR(금액)', 'DIVIDE(금액)', 'NOT(1,2)', 'PI(1)']) {
    assert.match(checkCalc(f, ['금액']).error, /인수 개수/, f);
    assert.equal(calculated(f).raw, '#VALUE!', f);
  }
  assert.equal(checkCalc('IFERROR(금액/0,)', ['금액']).ok, true);
  assert.equal(calculated('IFERROR(금액/0,)').raw, '0');
  assert.equal(calculated('IF(1,금액,1/0)').raw, '-1.5');
});

test('가져온 피벗 캐시: 다른 오류 항목이 합쳐지지 않고 항목 선택·슬라이서 필터가 각각 적용', () => {
  const rows = [['분류', '금액'], [{ error: '#N/A' }, 10], [{ error: '#DIV/0!' }, 20], [{ code: '#N/A' }, 5], ['정상', 30]];
  const def = { rows: ['분류'], cols: [], pages: [], values: [{ field: '금액', agg: 'sum' }], layout: 'tabular' };
  const run = (filters) => { const r = resolvePivot(rows, { ...def, filters }); return computePivot(r, r.def).grid.map((row) => row.map((c) => c.raw)); };
  assert.equal(keyOf(rows[1][0]), '#N/A');
  assert.equal(itemText(rows[2][0]), '#DIV/0!');
  assert.deepEqual(run({}).slice(1), [['#DIV/0!', '20'], ['#N/A', '15'], ['정상', '30'], ['총합계', '65']]);
  assert.deepEqual(run({ 분류: ['#N/A'] }).slice(1), [['#N/A', '15'], ['총합계', '15']]);
  assert.deepEqual(run({ 분류: ['#DIV/0!'] }).slice(1), [['#DIV/0!', '20'], ['총합계', '20']]);
  assert.deepEqual(run({ 분류: [] }).slice(1), [['총합계', '']]);
});
