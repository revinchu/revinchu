import test from 'node:test';
import assert from 'node:assert/strict';
import { filterSelection } from '../src/filter-selection.js';
import { parseTableRange, tableRangeProblem } from '../src/table-ux.js';
import { Workbook } from '../src/workbook.js';

test('filter search restores the draft selection and adds results to that draft', () => {
  const s = filterSelection(['서울', '서울2', '부산', '대구'], ['서울']);
  s.toggle('대구', true); s.search('부');
  assert.deepEqual(s.result(), ['부산']);
  assert.deepEqual(s.result(true), ['서울', '대구', '부산']);
  s.search(''); assert.deepEqual(s.result(), ['서울', '대구']);
  s.search('서울'); s.toggle('서울2', false); s.search('서울2');
  assert.deepEqual(s.result(), []);
  s.search('없음'); assert.deepEqual(s.visible(), []);
  assert.deepEqual(s.result(true), ['서울', '대구']);
});
test('filter search matches labels, keeps blanks and selects only visible results', () => {
  const s = filterSelection(['', 'a', 'b'], null, (v) => v === 'a' ? '첫째' : v);
  s.search('첫'); assert.deepEqual(s.visible(), ['a']);
  s.selectVisible(false); assert.deepEqual(s.result(), []);
  s.search(''); assert.deepEqual(s.result(), ['', 'a', 'b']);
  s.toggle('', false); assert.deepEqual(s.result(), ['a', 'b']);
});
test('table references accept exact current sheet and reject other sheets and invalid ranges', () => {
  assert.deepEqual(parseTableRange("='월별 ''현황'!$A$1:$C$9", "월별 '현황"), { r1: 0, c1: 0, r2: 8, c2: 2 });
  for (const ref of ['다른시트!A1:B3', '[Book]Sheet1!A1:B3', 'A0:B2', 'A:A', '1:4', 'A1, B3']) assert.equal(parseTableRange(ref, 'Sheet1'), null, ref);
});
test('table resize rejects merged, pivot and spilled ranges without changing workbook', () => {
  const wb = new Workbook(), rg = { r1: 0, c1: 0, r2: 4, c2: 2 };
  assert.equal(tableRangeProblem(wb, 0, rg), '');
  wb.transact(() => wb.merge(0, 2, 1, 2, 2)); assert.match(tableRangeProblem(wb, 0, rg), /병합/);
  wb.undo(); wb.transact(() => wb.setInput(0, 1, 0, '=SEQUENCE(3,2)'));
  const before = wb.serialize(); assert.match(tableRangeProblem(wb, 0, rg), /분산/); assert.deepEqual(wb.serialize(), before);
  wb.undo(); wb.transact(() => wb.setSheetProp(0, 'pivot', { area: { r1: 2, c1: 1, r2: 3, c2: 2 } }));
  assert.match(tableRangeProblem(wb, 0, rg), /피벗/);
});
