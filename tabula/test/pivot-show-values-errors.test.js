import { test } from 'node:test';
import assert from 'node:assert/strict';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

// Independent Excel 16.0 COM oracle: SUM of A/B/C, item B as base except baseNumeric=A.
// DisplayErrorString=false; normal/cumulative, baseline comparisons, ratios, dense ranks.
const inputs = {
  baseNumeric: [10, { code: '#N/A' }, 30], firstError: [{ code: '#DIV/0!' }, 20, 30],
  lastError: [10, 20, { code: '#N/A' }], baseError: [10, { code: '#DIV/0!' }, 30],
  numeric: [10, 20, 30], blank: [10, null, 30], zero: [10, 0, 30],
  twoErrors: [{ code: '#N/A' }, { code: '#DIV/0!' }, 30],
};
const oracle = [
  ["baseNumeric","normal",[10.0,"#N/A",30.0,"#N/A"]],
  ["baseNumeric","difference",["","#N/A",20.0,""]],
  ["baseNumeric","percent",[1.0,"#N/A",3.0,""]],
  ["baseNumeric","percentDiff",["","#N/A",2.0,""]],
  ["baseNumeric","runTotal",[10.0,"#N/A","#N/A",""]],
  ["baseNumeric","percentOfRow",[1.0,"#N/A",1.0,"#N/A"]],
  ["baseNumeric","percentOfCol",["#N/A","#N/A","#N/A","#N/A"]],
  ["baseNumeric","percentOfTotal",["#N/A","#N/A","#N/A","#N/A"]],
  ["baseNumeric","index",["#N/A","#N/A","#N/A","#N/A"]],
  ["baseNumeric","percentOfParentRow",["","","",""]],
  ["baseNumeric","percentOfParentCol",["","","",""]],
  ["baseNumeric","percentOfParent",[1.0,"",1.0,""]],
  ["baseNumeric","percentOfRunningTotal",["#DIV/0!","#N/A","#N/A","#N/A"]],
  ["baseNumeric","rankAscending",[1.0,"",2.0,""]],
  ["baseNumeric","rankDescending",[2.0,"",1.0,""]],
  ["firstError","normal",["#DIV/0!",20.0,30.0,"#DIV/0!"]],
  ["firstError","difference",["#DIV/0!","",10.0,""]],
  ["firstError","percent",["#DIV/0!",1.0,1.5,""]],
  ["firstError","percentDiff",["#DIV/0!","",0.5,""]],
  ["firstError","runTotal",["#DIV/0!","#DIV/0!","#DIV/0!",""]],
  ["firstError","percentOfRow",["#DIV/0!",1.0,1.0,"#DIV/0!"]],
  ["firstError","percentOfCol",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["firstError","percentOfTotal",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["firstError","index",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["firstError","percentOfParentRow",["","","",""]],
  ["firstError","percentOfParentCol",["","","",""]],
  ["firstError","percentOfParent",["",1.0,1.0,""]],
  ["firstError","percentOfRunningTotal",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["firstError","rankAscending",["",1.0,2.0,""]],
  ["firstError","rankDescending",["",2.0,1.0,""]],
  ["lastError","normal",[10.0,20.0,"#N/A","#N/A"]],
  ["lastError","difference",[-10.0,"","#N/A",""]],
  ["lastError","percent",[0.5,1.0,"#N/A",""]],
  ["lastError","percentDiff",[-0.5,"","#N/A",""]],
  ["lastError","runTotal",[10.0,30.0,"#N/A",""]],
  ["lastError","percentOfRow",[1.0,1.0,"#N/A","#N/A"]],
  ["lastError","percentOfCol",["#N/A","#N/A","#N/A","#N/A"]],
  ["lastError","percentOfTotal",["#N/A","#N/A","#N/A","#N/A"]],
  ["lastError","index",["#N/A","#N/A","#N/A","#N/A"]],
  ["lastError","percentOfParentRow",["","","",""]],
  ["lastError","percentOfParentCol",["","","",""]],
  ["lastError","percentOfParent",[1.0,1.0,"",""]],
  ["lastError","percentOfRunningTotal",["#DIV/0!","#DIV/0!","#N/A","#N/A"]],
  ["lastError","rankAscending",[1.0,2.0,"",""]],
  ["lastError","rankDescending",[2.0,1.0,"",""]],
  ["baseError","normal",[10.0,"#DIV/0!",30.0,"#DIV/0!"]],
  ["baseError","difference",["#DIV/0!","","#DIV/0!",""]],
  ["baseError","percent",["","","",""]],
  ["baseError","percentDiff",["","","",""]],
  ["baseError","runTotal",[10.0,"#DIV/0!","#DIV/0!",""]],
  ["baseError","percentOfRow",[1.0,"#DIV/0!",1.0,"#DIV/0!"]],
  ["baseError","percentOfCol",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["baseError","percentOfTotal",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["baseError","index",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["baseError","percentOfParentRow",["","","",""]],
  ["baseError","percentOfParentCol",["","","",""]],
  ["baseError","percentOfParent",[1.0,"",1.0,""]],
  ["baseError","percentOfRunningTotal",["#DIV/0!","#DIV/0!","#DIV/0!","#DIV/0!"]],
  ["baseError","rankAscending",[1.0,"",2.0,""]],
  ["baseError","rankDescending",[2.0,"",1.0,""]],
  ["numeric","normal",[10.0,20.0,30.0,60.0]],
  ["numeric","difference",[-10.0,"",10.0,""]],
  ["numeric","percent",[0.5,1.0,1.5,""]],
  ["numeric","percentDiff",[-0.5,"",0.5,""]],
  ["numeric","runTotal",[10.0,30.0,60.0,""]],
  ["numeric","percentOfRow",[1.0,1.0,1.0,1.0]],
  ["numeric","percentOfCol",[0.16666666666666666,0.3333333333333333,0.5,1.0]],
  ["numeric","percentOfTotal",[0.16666666666666666,0.3333333333333333,0.5,1.0]],
  ["numeric","index",[1.0,1.0,1.0,1.0]],
  ["numeric","percentOfParentRow",[0.16666666666666666,0.3333333333333333,0.5,1.0]],
  ["numeric","percentOfParentCol",["","","",""]],
  ["numeric","percentOfParent",[1.0,1.0,1.0,""]],
  ["numeric","percentOfRunningTotal",[0.16666666666666666,0.5,1.0,""]],
  ["numeric","rankAscending",[1.0,2.0,3.0,""]],
  ["numeric","rankDescending",[3.0,2.0,1.0,""]],
  ["blank","normal",[10.0,"",30.0,40.0]],
  ["blank","difference",[10.0,"",30.0,""]],
  ["blank","percent",["","","",""]],
  ["blank","percentDiff",["","","",""]],
  ["blank","runTotal",[10.0,10.0,40.0,""]],
  ["blank","percentOfRow",[1.0,"#DIV/0!",1.0,1.0]],
  ["blank","percentOfCol",[0.25,0.0,0.75,1.0]],
  ["blank","percentOfTotal",[0.25,0.0,0.75,1.0]],
  ["blank","index",[1.0,"#DIV/0!",1.0,1.0]],
  ["blank","percentOfParentRow",[0.25,0.0,0.75,1.0]],
  ["blank","percentOfParentCol",["","","",""]],
  ["blank","percentOfParent",[1.0,"",1.0,""]],
  ["blank","percentOfRunningTotal",[0.25,0.25,1.0,""]],
  ["blank","rankAscending",[1.0,"",2.0,""]],
  ["blank","rankDescending",[2.0,"",1.0,""]],
  ["zero","normal",[10.0,0.0,30.0,40.0]],
  ["zero","difference",[10.0,"",30.0,""]],
  ["zero","percent",["#DIV/0!","#DIV/0!","#DIV/0!",""]],
  ["zero","percentDiff",["#DIV/0!","","#DIV/0!",""]],
  ["zero","runTotal",[10.0,10.0,40.0,""]],
  ["zero","percentOfRow",[1.0,"#DIV/0!",1.0,1.0]],
  ["zero","percentOfCol",[0.25,0.0,0.75,1.0]],
  ["zero","percentOfTotal",[0.25,0.0,0.75,1.0]],
  ["zero","index",[1.0,"#DIV/0!",1.0,1.0]],
  ["zero","percentOfParentRow",[0.25,0.0,0.75,1.0]],
  ["zero","percentOfParentCol",["","","",""]],
  ["zero","percentOfParent",[1.0,"",1.0,""]],
  ["zero","percentOfRunningTotal",[0.25,0.25,1.0,""]],
  ["zero","rankAscending",[2.0,1.0,3.0,""]],
  ["zero","rankDescending",[2.0,3.0,1.0,""]],
  ["twoErrors","normal",["#N/A","#DIV/0!",30.0,"#N/A"]],
  ["twoErrors","difference",["#N/A","","#DIV/0!",""]],
  ["twoErrors","percent",["#N/A","","",""]],
  ["twoErrors","percentDiff",["#N/A","","",""]],
  ["twoErrors","runTotal",["#N/A","#DIV/0!","#DIV/0!",""]],
  ["twoErrors","percentOfRow",["#N/A","#DIV/0!",1.0,"#N/A"]],
  ["twoErrors","percentOfCol",["#N/A","#DIV/0!","#N/A","#N/A"]],
  ["twoErrors","percentOfTotal",["#N/A","#DIV/0!","#N/A","#N/A"]],
  ["twoErrors","index",["#N/A","#DIV/0!","#N/A","#N/A"]],
  ["twoErrors","percentOfParentRow",["","","",""]],
  ["twoErrors","percentOfParentCol",["","","",""]],
  ["twoErrors","percentOfParent",["","",1.0,""]],
  ["twoErrors","percentOfRunningTotal",["#N/A","#DIV/0!","#DIV/0!","#N/A"]],
  ["twoErrors","rankAscending",["","",1.0,""]],
  ["twoErrors","rankDescending",["","",1.0,""]],
];
function result(values, showAs, options = {}, patch = {}) {
  const input = [['Item', 'Amount'], ...values.map((v, i) => [String.fromCharCode(65 + i), v])];
  const original = structuredClone(input);
  const def = { rows: ['Item'], cols: [], pages: [], values: [{ field: 'Amount', agg: 'sum', showAs, baseField: 'Item', baseItem: 'B', ...options }], layout: 'tabular', errorShow: false, ...patch };
  const saved = structuredClone(def), resolved = resolvePivot(input, def);
  const grid = computePivot(resolved, resolved.def).grid;
  assert.deepEqual(input, original, 'source unchanged');
  assert.deepEqual(def, saved, 'definition unchanged');
  return grid.slice(1).map(row => row[1]?.raw);
}
function same(actual, expected, label) {
  assert.equal(actual.length, expected.length, label);
  expected.forEach((x, i) => {
    if (typeof x === 'number') {
      assert.notEqual(actual[i], '', label + ' nonblank');
      assert.ok(Math.abs(Number(actual[i]) - x) < 1e-12, `${label}[${i}]: ${actual[i]} != ${x}`);
    } else assert.equal(actual[i], x, label + '[' + i + ']');
  });
}
for (const [name, values] of Object.entries(inputs)) {
  test('값 표시 형식 Excel 합성 비교: ' + name, () => {
    for (const [, mode, expected] of oracle.filter(x => x[0] === name)) same(result(values, mode, { baseItem: name === 'baseNumeric' ? 'A' : 'B' }), expected, name + '/' + mode);
  });
}
test('누계는 열 축·형제 그룹에서도 오류를 유지하고 다음 그룹에는 누출하지 않음', () => {
  const input = [['Group', 'Item', 'Amount'], ['X', 'A', 10], ['X', 'B', { code: '#N/A' }], ['X', 'C', 30], ['Y', 'A', 1], ['Y', 'B', 2], ['Y', 'C', 3]];
  const def = { rows: ['Group'], cols: ['Item'], values: [{ field: 'Amount', agg: 'sum', showAs: 'runTotal', baseField: 'Item' }], layout: 'tabular', errorShow: false };
  const resolved = resolvePivot(input, def), grid = computePivot(resolved, resolved.def).grid;
  assert.deepEqual(grid.find(r => r[0]?.raw === 'X').slice(1, 4).map(c => c.raw), ['10', '#N/A', '#N/A']);
  assert.deepEqual(grid.find(r => r[0]?.raw === 'Y').slice(1, 4).map(c => c.raw), ['1', '3', '6']);
});
test('표시 형식 오류는 오류 캡션 옵션을 따르고 순위는 오류를 빈 항목으로 처리', () => {
  const values = inputs.baseNumeric;
  assert.deepEqual(result(values, 'runTotal', {}, { errorShow: true, errorCaption: '' }), ['10', '', '', '']);
  assert.deepEqual(result(values, 'difference', {}, { errorShow: true, errorCaption: '확인' }), ["'확인", '', "'확인", '']);
  assert.deepEqual(result(values, 'rankAscending'), ['1', '', '2', '']);
});

test('15개 값 표시 설정은 XLSX 저장·재열기 후에도 기준 항목과 오류 결과를 보존', () => {
  const wb = new Workbook();
  wb.transact(() => {
    [['Item', 'Amount'], ['A', 10], ['B', '=NA()'], ['C', 30]].forEach((row, r) => row.forEach((x, c) => wb.setInput(0, r, c, String(x))));
    for (const [, mode] of oracle.filter(x => x[0] === 'baseNumeric')) {
      const si = wb.addSheet(mode);
      wb.setSheetProp(si, 'pivot', { source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, top: 0, left: 0, area: { r1: 0, c1: 0, r2: 4, c2: 1 }, rows: ['Item'], cols: [], pages: [], layout: 'tabular', errorShow: false, values: [{ field: 'Amount', agg: 'sum', showAs: mode, baseField: 'Item', baseItem: 'A' }] });
    }
  });
  const before = wb.serialize();
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(wb.serialize(), before, 'export source unchanged');
  for (const sheet of back.sheets.slice(1)) {
    const d = sheet.pivot, r = resolvePivot(pivotSourceData(back, d), d);
    const expected = oracle.find(x => x[0] === 'baseNumeric' && x[1] === sheet.name)[2];
    same(computePivot(r, r.def).grid.slice(1).map(row => row[1].raw), expected, 'roundtrip/' + sheet.name);
  }
});
test('원본 오류를 고친 뒤 재집계 및 Undo는 누계와 순위를 다시 계산', () => {
  const wb = new Workbook();
  wb.transact(() => [['Item', 'Amount'], ['A', 10], ['B', '=NA()'], ['C', 30]].forEach((row, r) => row.forEach((x, c) => wb.setInput(0, r, c, String(x)))));
  const run = showAs => {
    const d = { source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, rows: ['Item'], cols: [], values: [{ field: 'Amount', agg: 'sum', showAs, baseField: 'Item' }], layout: 'tabular', errorShow: false };
    const r = resolvePivot(pivotSourceData(wb, d), d);
    return computePivot(r, r.def).grid.slice(1).map(row => row[1].raw);
  };
  assert.deepEqual(run('runTotal'), ['10', '#N/A', '#N/A', '']);
  wb.transact(() => wb.setInput(0, 2, 1, '20'));
  assert.deepEqual(run('runTotal'), ['10', '30', '60', '']);
  assert.deepEqual(run('rankAscending'), ['1', '2', '3', '']);
  wb.undo();
  assert.deepEqual(run('runTotal'), ['10', '#N/A', '#N/A', '']);
  assert.deepEqual(run('rankAscending'), ['1', '', '2', '']);
});
