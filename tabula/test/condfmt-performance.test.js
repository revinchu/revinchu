import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { prepareCond, condFormatAt, condMatch } from '../src/condfmt.js';

const area = { r1: 0, c1: 0, r2: 99999, c2: 0 };
test('셀 비교·수식 조건부 서식은 적용 범위 전체를 조회하지 않는다', () => {
  let reads = 0;
  const wb = { sheets: [{ cond: ['gt', 'eq', 'between', 'formula', 'date', 'text', 'blank', 'errors'].map(type => ({ ...area, type, v1: '3', formula: '=A1>3' })) }], usedRange: () => ({ rows: 100000, cols: 1 }), getValue: () => { reads++; return 4; } };
  const preps = prepareCond(wb, 0);
  assert.equal(reads, 0);
  assert.equal(condMatch(preps[0], 4), true);
});

test('숫자 순위·평균·백분위수와 중복 서식 결과를 유지한다', () => {
  const values = [10, 2, 2, 8, 4, 'text', 'text'];
  const rules = [
    { type: 'top', v1: 2 }, { type: 'bottom', v1: 2 },
    { type: 'aboveAvg', stdDev: 1 }, { type: 'belowAvg', stdDev: 1 },
    { type: 'dup' }, { type: 'unique' },
    { type: 'scale', colors: ['#000000', '#ffffff'], cfvo: [{ type: 'percentile', v: 25 }, { type: 'percentile', v: 75 }] },
    { type: 'bar' }, { type: 'icons' },
  ].map(rule => ({ r1: 0, c1: 0, r2: values.length - 1, c2: 0, ...rule }));
  const wb = { sheets: [{ cond: rules }], usedRange: () => ({ rows: values.length, cols: 1 }), getValue: (si, r) => values[r] };
  const p = prepareCond(wb, 0);
  assert.equal(p[0].topCut, 8); assert.equal(p[1].bottomCut, 2);
  assert.equal(p[2].avg, 5.2); assert.ok(Math.abs(p[2].sd - Math.sqrt(13.2)) < 1e-12);
  assert.equal(condMatch(p[2], 10), true); assert.equal(condMatch(p[3], 2), false);
  assert.equal(condMatch(p[4], 2), true); assert.equal(condMatch(p[4], 'text'), true);
  assert.equal(condMatch(p[5], 4), true);
  assert.deepEqual(p[6].points, [2, 8]);
  assert.equal(p[7].min, 2); assert.equal(p[7].max, 10);
  assert.equal(p[8].min, 2); assert.equal(p[8].max, 10);
});

test('수식 조건부 서식은 화면 셀의 상대 참조와 편집 결과를 계속 반영한다', () => {
  const wb = new Workbook({ sheets: [{ name: 'Sheet1', cells: { '0,0': { raw: '2' }, '1,0': { raw: '8' } }, cond: [{ r1: 0, c1: 0, r2: 999999, c2: 0, type: 'formula', formula: '=A1>3', style: { fill: '#ff0000' } }] }] });
  let p = prepareCond(wb, 0);
  assert.equal(condFormatAt(p, wb, 0, 0, 0, 2).style, null);
  assert.equal(condFormatAt(p, wb, 0, 1, 0, 8).style.fill, '#ff0000');
  wb.transact(() => wb.setInput(0, 0, 0, '9'));
  p = prepareCond(wb, 0);
  assert.equal(condFormatAt(p, wb, 0, 0, 0, 9).style.fill, '#ff0000');
});
