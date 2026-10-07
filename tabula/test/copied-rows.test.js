import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';
import { MAX_ROWS, MAX_COLS } from '../src/formula.js';
import { hid } from '../src/axis.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { snapshotPasteSource, pasteSpecialRange, preflightPasteSpecial, applyPasteSpecial, remapPasteSource } from '../src/paste-special.js';
const box = (r1, c1 = 0, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const book = (cells = {}, target = {}) => new Workbook({ sheets: [{ name: 'Source', cells }, { name: 'Target', cells: target }] });
function copyRows(w, r1, count, width = 1, si = 0, included) {
  const rows = included ?? Array.from({ length: count }, (_, i) => r1 + i);
  const data = rows.map(r => Array.from({ length: width }, (_, c) => cellData(w.getCell(si, r, c))));
  const values = rows.map(r => Array.from({ length: width }, (_, c) => w.getValue(si, r, c)));
  return snapshotPasteSource(w, { kind: 'rows', si, r1, r2: r1 + rows.length - 1, c1: 0, c2: width - 1, rows, data, values });
}
const paste = (w, src, target, opts = {}, si = 1) => w.transact(() => applyPasteSpecial(w, si, src, target, opts));
function insert(w, src, si, index) {
  const target = box(index), opts = { insertRows: true };
  const plan = preflightPasteSpecial(w, si, src, target, opts);
  let shifted = remapPasteSource(w, src, { si, axis: 'row', index, count: src.data.length });
  w.transact(() => { w.insertRows(si, index, src.data.length); shifted = { ...shifted, sourceSheet: w.sheets[shifted.si] }; applyPasteSpecial(w, si, shifted, target, {}, plan); });
  return shifted;
}
const state = w => JSON.stringify(w.serialize());

test('전체 행 snapshot은 사용 범위 밖 빈 10행과 복사 시점 높이·서식·숨김을 보존한다', () => {
  const w = book({ '1,0': { raw: '7' } }), s = w.sheets[0];
  s.rowHeights = { 1: 27, 2: 35, 10: 44 }; s.rowManual = { 1: true, 10: true }; s.rowStyles = { 10: { italic: true, fill: '#abcdef' } }; s.hiddenRows = { 2: true };
  const src = copyRows(w, 1, 10);
  s.rowHeights[10] = 80; s.rowStyles[10].italic = false; delete s.hiddenRows[2]; w.setCellData(0, 1, 0, { raw: '99' });
  assert.equal(src.data.length, 10); assert.equal(src.data[9][0].raw, ''); assert.equal(src.values[0][0], 7);
  assert.equal(src.rowMeta[9].height, 44); assert.equal(src.rowMeta[9].manual, true); assert.equal(src.rowMeta[9].style.italic, true); assert.equal(!!src.rowMeta[1].hidden, true);
});

test('Ctrl+V 전체 행은 정확히 10행 덮어쓰고 먼 열을 지우며 아래 행을 이동하지 않는다', () => {
  const w = book({ '1,0': { raw: '11' }, '2,0': { raw: '=A2+1' } }, { '20,0': { raw: 'old' }, '29,12': { raw: 'far', style: { bold: true } }, '30,0': { raw: 'below' } });
  w.sheets[0].rowHeights = { 1: 27, 10: 44 }; w.sheets[0].rowManual = { 10: true }; w.sheets[0].rowStyles = { 10: { italic: true } };
  w.sheets[1].colStyles[12] = { bold: true }; w.sheets[1].rowManual[20] = true; w.sheets[1].hiddenRows[20] = true;
  const src = copyRows(w, 1, 10), before = state(w), area = paste(w, src, box(20, 7));
  assert.deepEqual([area.r1, area.r2, area.c1, area.c2], [20, 29, 0, MAX_COLS - 1]);
  assert.equal(w.getRaw(1, 20, 0), '11'); assert.equal(w.getRaw(1, 21, 0), '=A21+1'); assert.equal(w.getRaw(1, 29, 12), ''); assert.equal(w.getRaw(1, 30, 0), 'below');
  assert.equal(w.rowHeight(1, 20), 27); assert.equal(w.rowHeight(1, 29), 44); assert.equal(w.sheets[1].rowManual[20], undefined); assert.equal(w.sheets[1].rowManual[29], true);
  assert.equal(w.styleAt(1, 29, 12).italic, true); assert.equal(w.styleAt(1, 29, 12).bold, false); assert.equal(!!hid(w.sheets[1].hiddenRows, 20), false);
  assert.ok(w.sheets[1].cells.size < 50); assert.equal(w.undoStack.length, 1); const after = state(w); w.undo(); assert.equal(state(w), before); w.redo(); assert.equal(state(w), after);
});

test('복사한 행 삽입은 10행을 만들고 기존 값·높이·병합·피벗을 아래로 옮기며 한 번에 Undo한다', () => {
  const w = book({ '1,0': { raw: '5' } }, { '20,0': { raw: 'old' }, '30,0': { raw: 'bottom' } });
  const s = w.sheets[1]; s.rowHeights[20] = 60; s.rowManual[20] = true; s.merges = [box(30, 2, 31, 3)];
  s.pivot = { name: 'SyntheticPivot', top: 30, left: 8, area: box(30, 8, 34, 9), buttons: [{ r: 30, c: 8 }] };
  w.sheets[0].rowHeights[10] = 44; w.sheets[0].rowManual[10] = true;
  const src = copyRows(w, 1, 10), before = state(w), frozen = JSON.stringify(src.data); insert(w, src, 1, 20);
  assert.equal(w.getRaw(1, 20, 0), '5'); assert.equal(w.getRaw(1, 29, 0), ''); assert.equal(w.getRaw(1, 30, 0), 'old'); assert.equal(w.getRaw(1, 40, 0), 'bottom');
  assert.equal(w.rowHeight(1, 30), 60); assert.equal(w.rowHeight(1, 29), 44); assert.equal(w.sheets[1].pivot.top, 40); assert.equal(w.sheets[1].pivot.buttons[0].r, 40); assert.equal(w.sheets[1].merges[0].r1, 40);
  assert.equal(JSON.stringify(src.data), frozen); assert.equal(w.undoStack.length, 1); const after = state(w); w.undo(); assert.equal(state(w), before); w.redo(); assert.equal(state(w), after);
});

test('전부 빈 10행도 기존 10행을 밀어낸다', () => {
  const w = book({}, { '5,0': { raw: 'marker' } }), src = copyRows(w, 100, 10);
  assert.equal(src.data.length, 10); insert(w, src, 1, 5); assert.equal(w.getRaw(1, 5, 0), ''); assert.equal(w.getRaw(1, 15, 0), 'marker');
});

test('행 서식 붙여넣기는 높이도 복사하지만 값·수식·빈셀 건너뛰기는 대상 높이를 유지한다', () => {
  for (const opts of [{ what: 'formats' }, { what: 'values' }, { what: 'formulas' }, { skipBlanks: true }]) {
    const w = book({ '1,0': { raw: '5', style: { fill: '#abcdef' } } }, { '8,0': { raw: 'old' } });
    w.sheets[0].rowHeights[1] = 27; w.sheets[0].rowManual[1] = true; w.sheets[1].rowHeights[8] = 60; w.sheets[1].rowManual[8] = true;
    paste(w, copyRows(w, 1, 1), box(8), opts); assert.equal(w.rowHeight(1, 8), opts.what === 'formats' ? 27 : 60);
    if (opts.what === 'formats') { assert.equal(w.getRaw(1, 8, 0), 'old'); assert.equal(w.styleAt(1, 8, 12).fill, '#ffffff'); }
  }
});

test('전체 행 복사도 열 너비·연결·전치는 조밀한 복사 폭을 유지한다', () => {
  const w = book({ '1,0': { raw: '1' }, '1,1': { raw: '2' } }); w.sheets[0].colWidths = { 0: 111, 1: 222 }; w.sheets[1].colWidths[12] = 333;
  const src = copyRows(w, 1, 1, 2); const a = paste(w, src, box(10, 3), { what: 'colWidths' });
  assert.equal(a.wholeRows, undefined); assert.equal(a.c2, 4); assert.equal(w.colWidth(1, 3), 111); assert.equal(w.colWidth(1, 4), 222); assert.equal(w.colWidth(1, 12), 333);
  paste(w, src, box(20, 4), { what: 'link' }); assert.equal(w.getRaw(1, 20, 4), '=Source!$A$2');
  const t = paste(w, src, box(30, 4), { transpose: true }); assert.equal(t.c2, 4); assert.equal(t.r2, 31); assert.equal(w.getRaw(1, 31, 4), '2');
});

test('전체 행 snapshot은 먼 셀·열 서식·병합·조건부서식·유효성도 복사한다', () => {
  const w = book({ '1,0': { raw: '1' }, '1,12': { raw: '=A2', style: { italic: true }, comment: 'synthetic' } });
  w.sheets[0].colStyles[20] = { fill: '#abc123' }; w.sheets[0].merges = [box(1, 12, 1, 13)];
  w.sheets[0].cond = [{ ...box(1, 12), type: 'formula', formula: '=M2>0', style: { bold: true } }];
  w.sheets[0].validations = [{ ...box(1, 20), type: 'custom', f1: 'U2>0' }];
  const src = copyRows(w, 1, 1); w.sheets[0].colStyles[20].fill = '#ffffff'; paste(w, src, box(10));
  assert.equal(w.getRaw(1, 10, 12), '=A11'); assert.equal(w.getCell(1, 10, 12).comment, 'synthetic'); assert.equal(w.styleAt(1, 10, 20).fill, '#abc123');
  assert.deepEqual(w.sheets[1].merges, [box(10, 12, 10, 13)]); assert.equal(w.sheets[1].cond[0].formula, '=M11>0'); assert.equal(w.sheets[1].validations[0].f1, 'U11>0');
});

test('빈셀 건너뛰기는 먼 tail 값의 규칙을 복사하고 빈 tail의 대상 규칙을 유지한다', () => {
  const w = book({ '1,0': { raw: '1' }, '1,12': { raw: '9' } }, { '10,13': { raw: 'keep' } });
  w.sheets[0].cond = [{ ...box(1, 12, 1, 13), type: 'formula', formula: '=M2>0' }];
  w.sheets[0].validations = [{ ...box(1, 12, 1, 13), type: 'custom', f1: 'M2>0' }];
  w.sheets[1].cond = [{ ...box(10, 12, 10, 13), type: 'formula', formula: '=M11<0' }];
  w.sheets[1].validations = [{ ...box(10, 12, 10, 13), type: 'whole', f1: '7' }];
  paste(w, copyRows(w, 1, 1), box(10), { skipBlanks: true });
  assert.equal(w.getRaw(1, 10, 12), '9'); assert.equal(w.getRaw(1, 10, 13), 'keep');
  assert.ok(w.sheets[1].cond.some(r => r.c1 === 12 && r.c2 === 12 && r.formula === '=M11>0')); assert.ok(w.sheets[1].cond.some(r => r.c1 === 13 && r.formula === '=N11<0'));
  assert.ok(w.sheets[1].validations.some(r => r.c1 === 12 && r.f1 === 'M11>0')); assert.ok(w.sheets[1].validations.some(r => r.c1 === 13 && r.f1 === '7'));
});

test('가시 행 복사는 실제 포함된 행 개수·높이·원본 수식 주소를 유지한다', () => {
  const w = book({ '1,0': { raw: '=B2' }, '3,0': { raw: '=B4' } }); w.sheets[0].rowHeights = { 1: 27, 3: 35 };
  const src = copyRows(w, 1, 4, 1, 0, [1, 3]); assert.equal(src.data.length, 2); insert(w, src, 1, 10);
  assert.equal(w.getRaw(1, 10, 0), '=B11'); assert.equal(w.getRaw(1, 11, 0), '=B12'); assert.equal(w.rowHeight(1, 10), 27); assert.equal(w.rowHeight(1, 11), 35);
});

test('같은 시트 복사 삽입은 절대 참조부터 구조 보정하고 원본 clip을 변경하지 않는다', () => {
  const w = book({ '1,0': { raw: '=$A$10+B2+$B2+B$2', link: '#Source!A10' }, '9,0': { raw: '3' } });
  const src = copyRows(w, 1, 1), original = JSON.stringify(src.data); const shifted = insert(w, src, 0, 3);
  assert.equal(w.getRaw(0, 3, 0), '=$A$11+B4+$B4+B$2'); assert.equal(shifted.data[0][0].raw, '=$A$11+B2+$B2+B$2'); assert.equal(shifted.data[0][0].link, '#Source!A11');
  assert.equal(JSON.stringify(src.data), original); const repeated = insert(w, shifted, 0, 3); assert.equal(w.getRaw(0, 3, 0), '=$A$12+B4+$B4+B$2'); assert.equal(repeated.data[0][0].raw, '=$A$12+B2+$B2+B$2');
});

test('복사 위치 앞 삽입은 source rows를 이동하고 내부 삽입도 복사 snapshot 높이를 유지한다', () => {
  const w = book({ '5,0': { raw: '=B6+$A$10' }, '6,0': { raw: '2' } }), src = copyRows(w, 5, 2);
  const shifted = insert(w, src, 0, 2); assert.deepEqual(shifted.rows, [7, 8]); assert.equal(w.getRaw(0, 2, 0), '=B3+$A$12'); assert.equal(w.getRaw(0, 7, 0), '=B8+$A$12');
  const across = remapPasteSource(w, shifted, { si: 0, axis: 'row', index: 8, count: 2 }); assert.deepEqual(across.rows, [7, 10]); assert.equal(across.data.length, 2); assert.equal(across.values.length, 2);
});

test('다른 원본 시트 clip은 대상 시트 참조만 구조 보정하고 source 좌표를 유지한다', () => {
  const w = book({ '1,0': { raw: '=Target!$A$10+$A$10', link: '#Target!A10' }, '1,12': { raw: '=Target!$A$10' } });
  w.sheets[0].cond = [{ ...box(1, 12), type: 'formula', formula: '=Target!$A$10>0' }]; w.sheets[0].validations = [{ ...box(1, 12), type: 'custom', f1: 'Target!$A$10>0' }];
  const src = copyRows(w, 1, 1), shifted = insert(w, src, 1, 3);
  assert.deepEqual(shifted.rows, [1]); assert.equal(w.getRaw(1, 3, 0), '=Target!$A$11+$A$10'); assert.equal(w.getRaw(1, 3, 12), '=Target!$A$11'); assert.equal(shifted.data[0][0].link, '#Target!A11');
  assert.equal(w.sheets[1].cond[0].formula, '=Target!$A$11>0'); assert.equal(w.sheets[1].validations[0].f1, 'Target!$A$11>0');
});

test('숨김 bitset과 보호된 행의 잠금 속성은 대상 보호 계약을 보존한다', () => {
  const w = book({ '1,0': { raw: '1' } }); w.sheets[0].hiddenRows = { 1: true }; w.sheets[0].rowStyles[1] = { locked: false, hideFormula: false, italic: true };
  w.sheets[1].hiddenRows = { start: 8, __bits: new Uint8Array([0, 1]), count: 1 }; w.sheets[1].protect = { on: true }; w.sheets[1].rowStyles[8] = { locked: true, hideFormula: true };
  paste(w, copyRows(w, 1, 2), box(8)); assert.equal(!!hid(w.sheets[1].hiddenRows, 8), true); assert.equal(!!hid(w.sheets[1].hiddenRows, 9), false); assert.equal(w.sheets[1].hiddenRows.count, 1);
  assert.equal(w.styleAt(1, 8, 12).locked, true); assert.equal(w.styleAt(1, 8, 12).hideFormula, true); assert.equal(w.styleAt(1, 8, 12).italic, true);
});

test('행 마지막 경계와 옵션 오류는 구조 변경 전 거절한다', () => {
  const w = book({ '1,0': { raw: '1' } }), src = copyRows(w, 1, 10), before = state(w);
  assert.throws(() => preflightPasteSpecial(w, 1, src, box(MAX_ROWS - 9), { insertRows: true }), /마지막 행/);
  assert.equal(pasteSpecialRange(src, box(MAX_ROWS - 10)).r2, MAX_ROWS - 1);
  assert.throws(() => preflightPasteSpecial(w, 1, src, box(10), { what: 'invalid' }), /지원하지/); assert.equal(state(w), before); assert.equal(w.undoStack.length, 0);
});

test('희소 target 면적은 실제 대상 행만 검사하고 삽입에는 밀릴 기존 열을 포함하지 않는다', () => {
  const w = book({ '1,0': { raw: '1' } }), src = copyRows(w, 1, 128);
  for (let c = 1; c < MAX_COLS; c++) w.setCellData(1, 1000, c, { raw: '1' });
  assert.equal(preflightPasteSpecial(w, 1, src, box(20)).tailColumns.size, 0);
  assert.throws(() => preflightPasteSpecial(w, 1, src, box(1000)), /200만/); assert.equal(w.undoStack.length, 0);
  const plan = preflightPasteSpecial(w, 1, src, box(1000), { insertRows: true }); assert.equal(plan.tailColumns.size, 0);
  insert(w, src, 1, 1000); assert.equal(w.getRaw(1, 1128, MAX_COLS - 1), '1'); assert.equal(w.getRaw(1, 1000, MAX_COLS - 1), '');
});

test('전체 행 logical area는 먼 피벗과 보호 범위까지 포함하고 조밀한 source 크기는 작다', () => {
  const w = book({ '1,0': { raw: '1' } }), src = copyRows(w, 1, 10), area = pasteSpecialRange(src, box(20, 2));
  assert.equal(src.data[0].length, 1); assert.equal(src.c2, 0); assert.equal(area.c1, 0); assert.equal(area.c2, MAX_COLS - 1); assert.equal(area.r2, 29);
});


test('전체 행 덮어쓰기가 병합의 일부와 교차하면 변경·Undo 없이 거절한다', () => {
  const w = book({ '1,0': { raw: '1' } }), src = copyRows(w, 1, 10); w.sheets[1].merges = [box(19, 12, 21, 13)]; const before = state(w);
  assert.throws(() => paste(w, src, box(20)), /병합된 셀의 일부/); assert.equal(state(w), before); assert.equal(w.undoStack.length, 0);
  assert.doesNotThrow(() => preflightPasteSpecial(w, 1, src, box(20), { insertRows: true }));
});

test('복사한 뒤 이름 변경 및 Undo/Redo 후에도 같은 시트 수식 구조 보정이 유지된다', () => {
  const w = book({ '1,0': { raw: '=$A$10' } }), src = copyRows(w, 1, 1);
  w.renameSheet(0, 'Renamed'); const afterName = remapPasteSource(w, src, { si: 0, axis: 'row', index: 3, count: 1 });
  assert.equal(afterName.sourceName, 'Renamed'); assert.equal(afterName.data[0][0].raw, '=$A$11');
  const shifted = insert(w, src, 0, 3); w.undo();
  const afterUndo = remapPasteSource(w, { ...src, sourceName: 'Renamed', sourceSheet: { ...src.sourceSheet } }, { si: 0, axis: 'row', index: 3, count: 1 });
  assert.equal(afterUndo.data[0][0].raw, '=$A$11'); assert.equal(afterUndo.sourceSheet, w.sheets[0]);
  w.redo(); const afterRedo = remapPasteSource(w, shifted, { si: 0, axis: 'row', index: 3, count: 1 }); assert.equal(afterRedo.data[0][0].raw, '=$A$12');
});

test('복사 삽입 저장 왕복은 수식·빈 꼬리 행 높이·행 서식·병합·CF·DV를 보존한다', () => {
  const w = book({ '1,0': { raw: '5' }, '2,0': { raw: '=A2*2' }, '1,12': { raw: '=A2' } }, { '20,0': { raw: 'below' } });
  w.sheets[0].rowHeights[10] = 44; w.sheets[0].rowManual[10] = true; w.sheets[0].rowStyles[10] = { italic: true, fill: '#abcdef' }; w.sheets[0].hiddenRows[10] = true;
  w.sheets[0].merges = [box(1, 12, 1, 13)]; w.sheets[0].cond = [{ ...box(1, 12), type: 'formula', formula: '=M2>0', style: { bold: true } }];
  w.sheets[0].validations = [{ ...box(1, 12), type: 'custom', f1: 'M2>0' }]; insert(w, copyRows(w, 1, 10), 1, 20);
  const reopened = new Workbook(readXlsx(writeXlsx(w)).data);
  assert.equal(reopened.getRaw(1, 21, 0), '=A21*2'); assert.equal(reopened.getRaw(1, 20, 12), '=A21'); assert.equal(reopened.getRaw(1, 30, 0), 'below');
  assert.equal(reopened.rowHeight(1, 29), 44); assert.equal(reopened.sheets[1].rowManual[29], true); assert.equal(reopened.styleAt(1, 29, 100).italic, true); assert.equal(reopened.styleAt(1, 29, 100).fill, '#abcdef'); assert.equal(!!hid(reopened.sheets[1].hiddenRows, 29), true);
  assert.deepEqual(reopened.sheets[1].merges, [box(20, 12, 20, 13)]); assert.equal(reopened.sheets[1].cond[0].formula, '=M21>0'); assert.equal(reopened.sheets[1].validations[0].f1, 'M21>0');
});
