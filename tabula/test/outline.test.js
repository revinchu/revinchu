import { test } from 'node:test';
import assert from 'node:assert/strict';
import { groupsOf, changeLevels, toggleGroup, showLevel, groupAt, planSubtotals, normOutline } from '../src/outline.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

test('개요: 그룹 목록 · 접기 · 펼치기 · 수준 단추', () => {
  // 행 1~3 수준 2, 4 수준 1(요약), 5~6 수준 2, 7 수준 1(요약)
  let lv = changeLevels({}, 1, 7, 1);
  lv = changeLevels(lv, 1, 3, 1);
  lv = changeLevels(lv, 5, 6, 1);
  assert.deepEqual(groupsOf(lv), [{ a: 1, b: 7, level: 1 }, { a: 1, b: 3, level: 2 }, { a: 5, b: 6, level: 2 }]);
  assert.deepEqual(groupAt(lv, 4, true), { a: 1, b: 3, level: 2 }); // 요약 행은 그 그룹
  // 안쪽 그룹 접기 → 1~3 숨김, 요약 행 4 에 표시
  let st = toggleGroup(lv, {}, {}, { a: 1, b: 3, level: 2 }, true, true);
  assert.deepEqual(Object.keys(st.hidden).map(Number), [1, 2, 3]);
  assert.deepEqual(st.coll, { 4: true });
  // 바깥 그룹 접고 펼치면 안쪽 접힌 그룹은 그대로 숨김
  st = toggleGroup(lv, st.hidden, st.coll, { a: 1, b: 7, level: 1 }, true, true);
  assert.equal(Object.keys(st.hidden).length, 7);
  st = toggleGroup(lv, st.hidden, st.coll, { a: 1, b: 7, level: 1 }, true, false);
  assert.deepEqual(Object.keys(st.hidden).map(Number), [1, 2, 3]);
  // 수준 2 단추: 수준 2 행만 숨김
  const s2 = showLevel(lv, {}, 2, true);
  assert.deepEqual(Object.keys(s2.hidden).map(Number), [1, 2, 3, 5, 6]);
  assert.deepEqual(s2.coll, { 4: true, 7: true });
  const s3 = showLevel(lv, s2.hidden, 3, true);
  assert.deepEqual(s3.hidden, {});
  assert.deepEqual(planSubtotals(0, 5, (r) => ['', 'a', 'a', 'b', 'c', 'c'][r]).groups.map((g) => [g.key, g.a, g.b]), [['a', 1, 2], ['b', 3, 3], ['c', 4, 5]]);
});

test('개요: 행 삽입 시 수준 따라감 + xlsx 왕복(outlineLevel · collapsed · summaryBelow)', () => {
  const wb = new Workbook();
  wb.transact(() => {
    for (let r = 0; r < 8; r++) wb.setInput(0, r, 0, String(r));
    wb.setSheetProp(0, 'outline', { ...normOutline(), rows: { 1: 1, 2: 1, 3: 1 }, rowsColl: { 4: true }, cols: { 2: 1 }, below: false });
    wb.setSheetProp(0, 'hiddenRows', { 1: true, 2: true, 3: true });
  });
  wb.transact(() => wb.insertRows(0, 2, 1));
  assert.deepEqual(wb.sheets[0].outline.rows, { 1: 1, 2: 1, 3: 1, 4: 1 });
  assert.deepEqual(wb.sheets[0].outline.rowsColl, { 5: true });
  const back = readXlsx(writeXlsx(wb)).data.sheets[0];
  assert.deepEqual(back.outline.rows, { 1: 1, 2: 1, 3: 1, 4: 1 });
  assert.deepEqual(back.outline.rowsColl, { 5: true });
  assert.deepEqual(back.outline.cols, { 2: 1 });
  assert.equal(back.outline.below, false);
  assert.ok(back.hiddenRows[1]);
});
