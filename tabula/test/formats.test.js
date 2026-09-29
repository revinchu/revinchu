// XLSB 수식 해석 · ODS 읽기/쓰기 · 피벗 'Σ 값' 위치
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { decodeFormula } from '../src/xlsb.js';
import { readOds, writeOds, toOdfFormula, fromOdfFormula } from '../src/ods.js';
import { computePivot } from '../src/pivot.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';

/** rgce 바이트 배열 → CellParsedFormula (cce + rgce + cb=0) */
const fmla = (bytes, extra = []) => new Uint8Array([bytes.length, 0, 0, 0, ...bytes, extra.length, 0, 0, 0, ...extra]);
const i32 = (n) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
const u16 = (n) => [n & 255, (n >> 8) & 255];
const env = { sheets: ['Sheet1', '데이터 시트'], names: [{ name: '_xlfn.XLOOKUP' }, { name: '세율' }], xti: [{ self: true, first: 1, last: 1 }, { self: true, first: -2, last: -2 }], tables: new Map([[1, { name: '표1', cols: ['날짜', '비용'] }]]) };

test('XLSB 수식: 참조 · 연산자 · 함수 · 문자열 · 3D 참조 · 이름 · 미래 함수 · 표 참조', () => {
  // =SUM(A1:B2)*2 → PtgArea(0x25) · PtgInt · PtgMul · PtgFuncVar SUM(4) 순서: area, SUM(1개), int, mul
  const area = [0x25, ...i32(0), ...i32(1), ...u16(0xC000), ...u16(0xC001)];
  const sum = [0x22, 1, ...u16(4)];
  const t1 = decodeFormula(fmla([...area, ...sum, 0x1E, ...u16(2), 0x05]), 0, env, { r: 5, c: 5 });
  assert.equal(t1.text, 'SUM(A1:B2)*2');
  // 절대 참조 $C$3 & "원"
  const ref = [0x24, ...i32(2), ...u16(2)];
  const str = [0x17, ...u16(1), 0xD0, 0xC6];
  assert.equal(decodeFormula(fmla([...ref, ...str, 0x08]), 0, env, { r: 0, c: 0 }).text, '$C$3&"원"');
  // 3D: '데이터 시트'!A1
  assert.equal(decodeFormula(fmla([0x3A, ...u16(0), ...i32(0), ...u16(0xC000)]), 0, env, { r: 0, c: 0 }).text, "'데이터 시트'!A1");
  // 이름 · 미래 함수 (PtgNameX + FuncVar 255)
  assert.equal(decodeFormula(fmla([0x23, ...i32(2)]), 0, env, { r: 0, c: 0 }).text, '세율');
  const xl = [0x39, ...u16(1), ...i32(1), ...ref, ...ref, ...ref, 0x22, 4, ...u16(255)];
  assert.equal(decodeFormula(fmla(xl), 0, env, { r: 0, c: 0 }).text, '_xlfn.XLOOKUP($C$3,$C$3,$C$3)');
  // 표 참조: 표1[[#This Row],[비용]]
  const list = [0x18, 0x19, 0, 0, ...u16(0x41), ...u16(1), 0, 0, ...u16(1), ...u16(1)];
  assert.equal(decodeFormula(fmla(list), 0, env, { r: 3, c: 3 }).text, '표1[[#This Row],[비용]]');
  // 공유 수식의 상대 참조(PtgRefN): 한 행 위
  const refN = [0x2C, ...i32(0xFFFFF), ...u16(0xC000)];
  assert.equal(decodeFormula(fmla(refN), 0, env, { r: 9, c: 2 }).text, 'C9');
  // 모르는 토큰 → null (값만 가져옴)
  assert.equal(decodeFormula(fmla([0x02]), 0, env, { r: 0, c: 0 }).text, null);
});

test('ODS: 수식 변환 · 값 · 서식 · 병합 왕복', () => {
  const f = "=SUM(A1:B2)+'시트 2'!C3*IF(A1>0,\"a,b\",Sheet1!$D$4)";
  assert.equal(toOdfFormula(f), 'of:=SUM([.A1:.B2])+[$\'시트 2\'.C3]*IF([.A1]>0;"a,b";[$Sheet1.$D$4])');
  assert.equal(`=${fromOdfFormula(toOdfFormula(f))}`, f);
  const wb = new Workbook();
  wb.transact(() => {
    [['이름', '값'], ['A', '12.5'], ['', '=B2*2'], ['', '2024-03-15'], ['', '15%']].forEach((row, r) => row.forEach((v, c) => { if (v) wb.setInput(0, r, c, v); }));
    for (const c of [0, 1]) wb.setStyle(0, 0, c, { bold: true, fill: '#ffee00' });
  });
  const bytes = writeOds([{ si: 0, name: 'Sheet1', merges: [{ r1: 5, c1: 0, r2: 5, c2: 1 }] }], {
    raw: (s, r, c) => wb.getRaw(s, r, c), value: (s, r, c) => wb.getValue(s, r, c), style: (s, r, c) => wb.styleAt(s, r, c),
    used: (s) => wb.usedRange(s), colWidth: (s, c) => wb.colWidth(s, c), rowHeight: () => null,
  });
  assert.equal(textOf(unzip(bytes).mimetype), 'application/vnd.oasis.opendocument.spreadsheet');
  const back = new Workbook(readOds(bytes).data);
  assert.equal(back.getRaw(0, 2, 1), '=B2*2');
  assert.equal(back.getValue(0, 2, 1), 25);
  assert.equal(back.getValue(0, 3, 1), 45366);
  assert.equal(back.styleAt(0, 4, 1).numFmt, 'percent');
  assert.equal(back.styleAt(0, 0, 0).bold, true);
  assert.equal(back.styleAt(0, 0, 0).fill, '#ffee00');
  assert.deepEqual(back.sheets[0].merges, [{ r1: 5, c1: 0, r2: 5, c2: 1 }]);
});

test("피벗 'Σ 값' 위치: 맨 안쪽(월별) / 맨 바깥(지표별로 4월·5월 나란히) + xlsx 왕복", () => {
  const rows = [['월', '매체', '비용', '클릭'], ['4월', 'A', 10, 1], ['5월', 'A', 20, 2], ['4월', 'B', 30, 3], ['5월', 'B', 40, 4]];
  const base = { rows: ['매체'], cols: ['월'], pages: [], filters: {}, layout: 'tabular', grandRows: true, grandCols: false, style: 'None', values: [{ field: '비용', agg: 'sum' }, { field: '클릭', agg: 'sum' }] };
  const txt = (grid, r) => grid[r].map((c) => String(c?.raw ?? '').replace(/^'/, ''));
  const inner = computePivot(rows, base).grid;
  assert.deepEqual(txt(inner, 3).slice(1), ['10', '1', '20', '2']);
  const outer = computePivot(rows, { ...base, valuesPos: 0 }).grid;
  assert.deepEqual(txt(outer, 1).slice(1), ['합계 : 비용', '', '합계 : 클릭', '']);
  assert.deepEqual(txt(outer, 2).slice(1), ['4월', '5월', '4월', '5월']);
  assert.deepEqual(txt(outer, 3).slice(1), ['10', '20', '1', '2']);
  // xlsx: colFields 에 -2 가 앞에
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  wb.transact(() => wb.setSheetProp(0, 'pivot', { ...base, name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 4, c2: 3 }, valuesPos: 0, top: 0, left: 6, area: { r1: 0, c1: 6, r2: 5, c2: 10 } }));
  const xml = textOf(unzip(writeXlsx(wb))['xl/pivotTables/pivotTable1.xml']);
  assert.match(xml, /<colFields count="2"><field x="-2"\/><field x="0"\/><\/colFields>/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].pivot.valuesPos, 0);
});
