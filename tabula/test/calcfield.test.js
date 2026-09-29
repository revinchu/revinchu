// 계산 필드(검사 · 이름 바꾸기 · 엑셀 수식 · ROWS/DIVIDE) · COUNTIF 색인 · 숫자→글자 · 의존 전파
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { checkCalc, renameCalcRefs, excelCalcFormula, computePivot } from '../src/pivot.js';
import { fastCount, numberText, Range } from '../src/fxcore.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';

test('계산 필드 검사: 없는 필드 · 함수 · 순환 · 구문', () => {
  const fields = ['비용', '클릭수', '전환 매출'];
  assert.equal(checkCalc('비용/클릭수', fields).ok, true);
  assert.deepEqual(checkCalc("'전환 매출'/비용", fields).refs, ['전환 매출', '비용']);
  assert.match(checkCalc('비용/노출수', fields).error, /없는 필드: 노출수/);
  assert.match(checkCalc('VLOOKUP(비용)', fields).error, /지원하지 않는 함수/);
  assert.match(checkCalc('비용/(클릭수', fields).error, /구문/);
  const calcs = [{ name: 'A', formula: 'B*2' }, { name: 'B', formula: 'A+1' }];
  assert.equal(checkCalc('B*2', fields, calcs, 'A').cycle, true);
  assert.equal(checkCalc('DIVIDE(비용, ROWS())', fields).ok, true);
});

test('계산 필드 이름 바꾸기: 다른 수식의 참조도 바뀜 (따옴표 · 함수 이름은 그대로)', () => {
  assert.equal(renameCalcRefs('CPC*2+IF(CPC>0,1,0)', 'CPC', '클릭당 비용'), "'클릭당 비용'*2+IF('클릭당 비용'>0,1,0)");
  assert.equal(renameCalcRefs("'전환 매출'/비용", '전환 매출', '매출'), '매출/비용');
  assert.equal(renameCalcRefs('ROUND(X, 2)', 'ROUND', 'Z'), 'ROUND(X, 2)');
});

test('엑셀 파일용 수식: DIVIDE → IF, ROWS → 0, 그 밖은 그대로', () => {
  assert.equal(excelCalcFormula('비용/클릭수'), '비용/클릭수');
  assert.equal(excelCalcFormula('DIVIDE(비용, 클릭수)'), 'IF(클릭수=0,0,비용/클릭수)');
  assert.equal(excelCalcFormula("DIVIDE('전환 매출', 비용+1, -1)*100"), "IF(비용+1=0,-1,'전환 매출'/(비용+1))*100");
  assert.equal(excelCalcFormula('비용/ROWS()'), '비용/0');
});

test('피벗: ROWS() 는 그룹의 원본 행 수, DIVIDE 는 0 으로 나눠도 오류 없음', () => {
  const rows = [['지역', '비용', '클릭'], ['서울', 100, 10], ['서울', 50, 0], ['부산', 30, 0]];
  const d = {
    rows: ['지역'], cols: [], pages: [], filters: {}, layout: 'compact', grandRows: true, grandCols: true, style: 'None',
    values: [{ field: '평균비용', agg: 'sum' }, { field: 'CPC', agg: 'sum' }],
    calcFields: [{ name: '평균비용', formula: '비용/ROWS()' }, { name: 'CPC', formula: 'DIVIDE(비용, 클릭, -1)' }],
  };
  const { grid } = computePivot(rows, d);
  const at = (label) => grid.find((r) => String(r[0]?.raw).replace(/^'/, '') === label);
  assert.equal(Number(at('서울')[1].raw), 75);
  assert.equal(Number(at('부산')[1].raw), 30);
  assert.equal(Number(at('총합계')[1].raw), 60);
  assert.equal(Number(at('부산')[2].raw), -1);
  assert.equal(Number(at('서울')[2].raw), 15);
});

test('계산 필드 xlsx 왕복: 엑셀용 수식 + tb:formula 로 원래 수식 보존', () => {
  const wb = new Workbook();
  const data = [['지역', '비용', '클릭'], ['서울', '100', '10'], ['부산', '30', '0']];
  wb.transact(() => data.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v))));
  const def = {
    name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: 'CPC', agg: 'sum' }],
    calcFields: [{ name: 'CPC', formula: 'DIVIDE(비용, 클릭)' }], top: 0, left: 5, area: { r1: 0, c1: 5, r2: 3, c2: 6 },
  };
  wb.transact(() => wb.setSheetProp(0, 'pivot', def));
  const files = unzip(writeXlsx(wb));
  const cache = textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']);
  assert.match(cache, /mc:Ignorable="tb"/);
  assert.match(cache, /formula="IF\(클릭=0,0,비용\/클릭\)" tb:formula="DIVIDE\(비용, 클릭\)"/);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.deepEqual(back.sheets[0].pivot.calcFields, [{ name: 'CPC', formula: 'DIVIDE(비용, 클릭)' }]);
});

test('COUNTIF 색인: 같음 · 다름 · 크기 비교 결과가 일반 계산과 같음', () => {
  const vals = [];
  for (let i = 0; i < 200; i++) vals.push([i % 7 === 0 ? String(i % 13) : i % 5 === 0 ? null : (i % 13) + (i % 3) / 3]);
  vals.push(['abc'], ['ABC'], [''], [true], [1 / 3]);
  const r = new Range(vals);
  const brute = (test) => vals.flat().filter(test).length;
  assert.equal(fastCount(r, 5), brute((v) => v === 5 || v === '5'));
  assert.equal(fastCount(r, '<>5'), vals.length - brute((v) => v === 5 || v === '5'));
  assert.equal(fastCount(r, 'abc'), 2);
  assert.equal(fastCount(r, '>10'), brute((v) => typeof v === 'number' && v > 10));
  assert.equal(fastCount(r, '<=3'), brute((v) => typeof v === 'number' && v <= 3));
  assert.equal(fastCount(r, ''), brute((v) => v === null || v === ''));
  assert.equal(fastCount(r, 'TRUE'), 1);
  assert.equal(fastCount(r, 'a*'), null); // 와일드카드는 일반 계산으로
  // 15자리로 비교: 1/3 을 글자로 만든 조건과 같음
  assert.equal(fastCount(r, `>=${numberText(1 / 3)}`), brute((v) => typeof v === 'number' && Number(v.toPrecision(15)) >= 0.333333333333333));
});

test('숫자 → 글자: 엑셀처럼 유효 숫자 15자리', () => {
  assert.equal(numberText(1 / 3), '0.333333333333333');
  assert.equal(numberText(0.1 + 0.2), '0.3');
  assert.equal(numberText(1e20), '1E+20');
  assert.equal(numberText(123456789012345678), '1.23456789012346E+17');
  assert.equal(numberText(1.23e-10), '1.23E-10');
  assert.equal(numberText(0.000000123), '0.000000123');
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '=1/3'); wb.setInput(0, 0, 1, '=">"&A1'); wb.setInput(0, 0, 2, '=LEN(A1&"")'); });
  assert.equal(wb.getValue(0, 0, 1), '>0.333333333333333');
  assert.equal(wb.getValue(0, 0, 2), 17);
});

test('의존 전파: 고정 범위를 참조하는 수식 묶음은 바뀐 칸이 많아도 한 번만 훑음', () => {
  const wb = new Workbook();
  const N = 3000;
  wb.transact(() => {
    for (let r = 0; r < N; r++) wb.setInput(0, r, 0, String(r % 50));
    for (let r = 0; r < N; r++) wb.setInput(0, r, 1, `=COUNTIF($A$1:$A$${N},">"&A${r + 1})+COUNTIF($A$1:A${r + 1},A${r + 1})`);
  });
  assert.equal(wb.getValue(0, 0, 1), 2941);
  const t = Date.now();
  // 한 번에 3000칸을 바꿔도 (칸마다 3000개 수식을 훑으면 9백만 번) 빠르게
  wb.transact(() => { for (let r = 0; r < N; r++) wb.setInput(0, r, 0, String((r * 7) % 50)); });
  const took = Date.now() - t;
  assert.ok(took < 3000, `전파가 느림: ${took}ms`);
  let above = 0;
  for (let r = 0; r < N; r++) if ((r * 7) % 50 > 0) above++;
  assert.equal(wb.getValue(0, 0, 1), above + 1);
});
