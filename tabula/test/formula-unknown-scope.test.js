import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unknownFunctions, parse } from '../src/formula.js';
import { formulaSupportIssue } from '../src/calculation-state.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const unknown = (formula, isName) => unknownFunctions(formula.replace(/^=/, ''), isName);
const hasUnsupportedWarning = result => result.warnings.some(w => /지원하지 않는 함수/.test(w));

test('미지원 함수 검사는 기본 함수·문자열을 제외하고 실제 누락을 중복 없이 반환한다', () => {
  assert.deepEqual(unknown('SUM(A1:A5)+IF(TRUE,1,2)'), []);
  assert.deepEqual(unknown('IF(TRUE,"CUSTOM_FN(1)","LET(f,1,f(2))")'), []);
  assert.deepEqual(unknown('NO_SUCH_FN(1)+no_such_fn(2)+SECOND_FN(3)'), ['NO_SUCH_FN', 'SECOND_FN']);
});

test('LET의 지역 LAMBDA와 함수 값 바인딩은 미지원 함수가 아니다', () => {
  for (const formula of [
    'LET(f,LAMBDA(x,x+1),f(2))',
    'LET(F,LAMBDA(x,x+1),f(2))',
    'LET(sumfn,SUM,sumfn(1,2))',
    'LET(f,LAMBDA(x,x+1),g,LAMBDA(x,f(x)*2),g(3))',
  ]) assert.deepEqual(unknown(formula), [], formula);
});

test('LET 이름은 이전 바인딩이나 자기 값에는 보이지 않고 다음 바인딩부터 보인다', () => {
  assert.deepEqual(unknown('LET(value,f(1),f,LAMBDA(x,x+1),value)'), ['F']);
  assert.deepEqual(unknown('LET(f,f(1),f)'), ['F']);
  assert.deepEqual(unknown('LET(f,LAMBDA(x,x+1),value,f(1),value)'), []);
});

test('중첩 지역 범위는 부모 함수에 접근하되 형제 수식과 바깥으로 누출되지 않는다', () => {
  assert.deepEqual(unknown('LET(f,LAMBDA(x,x+1),LET(g,LAMBDA(x,f(x)*2),g(3)))'), []);
  assert.deepEqual(unknown('SUM(LET(f,LAMBDA(x,x+1),f(2)),f(3))'), ['F']);
  assert.deepEqual(unknown('LET(value,LET(f,LAMBDA(x,x+1),f(2)),f(value))'), ['F']);
});

test('LAMBDA 함수 매개변수는 본문에만 적용하고 호출 인수의 누락 함수는 보존한다', () => {
  assert.deepEqual(unknown('LAMBDA(fn,fn(2))(LAMBDA(x,x+1))'), []);
  assert.deepEqual(unknown('LAMBDA(fn,LAMBDA(value,fn(value)))(LAMBDA(x,x+1))(2)'), []);
  assert.deepEqual(unknown('LAMBDA(fn,fn(2))(fn(1))'), ['FN']);
  assert.deepEqual(unknown('LAMBDA(fn,fn(NO_SUCH_FN(2)))(LAMBDA(x,x))'), ['NO_SUCH_FN']);
});

test('Excel 접두사와 통합 문서 정의 함수도 지역 함수 검사와 함께 처리한다', () => {
  assert.deepEqual(unknown('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),f(2))'), []);
  const isName = name => name.toLowerCase() === 'twice';
  assert.deepEqual(unknown('LET(f,LAMBDA(x,twice(x)),f(2))', isName), []);
  assert.deepEqual(unknown('LET(f,LAMBDA(x,MISSING_FN(x)),f(2))+twice(1)', isName), ['MISSING_FN']);
});

test('AST를 읽을 수 없는 수식은 기존 미지원 후보를 보수적으로 유지한다', () => {
  assert.deepEqual(unknown('LET(f,LAMBDA(x,x),f(2),NO_SUCH_FN(1)'), ['F', 'NO_SUCH_FN']);
});

test('가져오기 함수 판별과 계산 상태의 지역 범위 판별이 일치한다', () => {
  for (const formula of [
    'LET(f,LAMBDA(x,x+1),f(2))',
    'LET(value,f(1),f,LAMBDA(x,x),value)',
    'SUM(LET(f,LAMBDA(x,x+1),f(2)),f(3))',
    'LAMBDA(fn,fn(2))(fn(1))',
    'LET(f,LAMBDA(x,NO_SUCH_FN(x)),f(2))',
  ]) {
    const issue = formulaSupportIssue({ formula: true, ast: parse(formula) });
    assert.deepEqual(unknown(formula), issue?.functions ?? [], formula);
  }
});

test('지역 함수 XLSX 왕복은 경고 없이 입력 변경·Undo·Redo 결과를 다시 계산한다', () => {
  const source = new Workbook({ names: [{ name: 'twice', ref: '=LAMBDA(x,x*2)' }], sheets: [{ name: '입력', cells: {
    '0,0': { raw: '3' },
    '0,1': { raw: '=LET(f,LAMBDA(x,x+1),f(A1))' },
    '1,1': { raw: '=LAMBDA(fn,fn(A1))(LAMBDA(x,x+2))' },
    '2,1': { raw: '=LET(sumfn,SUM,sumfn(A1,2))' },
    '3,1': { raw: '=LET(f,LAMBDA(x,twice(x)),f(A1))' },
  } }] });
  const initial = readXlsx(writeXlsx(source));
  assert.equal(hasUnsupportedWarning(initial), false);
  const wb = new Workbook(initial.data);
  const values = () => [0, 1, 2, 3].map(r => wb.getValue(0, r, 1));
  assert.deepEqual(values(), [4, 5, 5, 6]);
  assert.equal(wb.calculationIssues().total, 0);
  wb.transact(() => wb.setInput(0, 0, 0, '5'));
  assert.deepEqual(values(), [6, 7, 7, 10]);
  assert.equal(wb.calculationIssues().total, 0);
  wb.undo(); assert.deepEqual(values(), [4, 5, 5, 6]);
  wb.redo(); assert.deepEqual(values(), [6, 7, 7, 10]);
  const saved = readXlsx(writeXlsx(wb));
  assert.equal(hasUnsupportedWarning(saved), false);
  const reopened = new Workbook(saved.data);
  assert.deepEqual([0, 1, 2, 3].map(r => reopened.getValue(0, r, 1)), [6, 7, 7, 10]);
  assert.equal(reopened.calculationIssues().total, 0);
});

test('지역 함수 본문 안의 진짜 미지원 함수와 파일 저장값 경고는 유지한다', () => {
  const wb = new Workbook({ sheets: [{ name: '입력', cells: {
    '0,0': { raw: '2' },
    '0,1': { raw: '=LET(f,LAMBDA(x,NO_SUCH_FN(x)),f(A1))', cached: 20 },
  } }] });
  const result = readXlsx(writeXlsx(wb));
  assert.equal(hasUnsupportedWarning(result), true);
  const reopened = new Workbook(result.data);
  assert.deepEqual(reopened.getCalculationStatus(0, 0, 1).functions, ['NO_SUCH_FN']);
  assert.equal(reopened.getValue(0, 0, 1), 20);
});
