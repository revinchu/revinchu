import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  colToName, nameToCol, parse, shiftFormula, adjustFormulaForStructure,
  renameSheetInFormula, ERR,
} from '../src/formula.js';
import { formatGeneral, formatNumber, parseInput, formatValue } from '../src/format.js';
import { Workbook } from '../src/workbook.js';

function book(cells) {
  const wb = new Workbook();
  wb.transact(() => {
    for (const [addr, raw] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      wb.setInput(0, +m[2] - 1, nameToCol(m[1]), raw);
    }
  });
  return wb;
}
const val = (wb, addr, si = 0) => {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  return wb.getValue(si, +m[2] - 1, nameToCol(m[1]));
};

test('열 이름 변환', () => {
  assert.equal(colToName(0), 'A');
  assert.equal(colToName(25), 'Z');
  assert.equal(colToName(26), 'AA');
  assert.equal(colToName(701), 'ZZ');
  assert.equal(colToName(702), 'AAA');
  assert.equal(nameToCol('AA'), 26);
  assert.equal(nameToCol('xfd'), 16383);
});

test('사칙연산과 우선순위', () => {
  const wb = book({
    A1: '=1+2*3', A2: '=(1+2)*3', A3: '=-2^2', A4: '=2^3^2', A5: '=10/4', A6: '=50%*10',
    A7: '="a"&1&TRUE', A8: '=1/0', A9: '=1<2', A10: '="abc"="ABC"',
  });
  assert.equal(val(wb, 'A1'), 7);
  assert.equal(val(wb, 'A2'), 9);
  assert.equal(val(wb, 'A3'), 4);
  assert.equal(val(wb, 'A4'), 64);
  assert.equal(val(wb, 'A5'), 2.5);
  assert.equal(val(wb, 'A6'), 5);
  assert.equal(val(wb, 'A7'), 'a1TRUE');
  assert.equal(val(wb, 'A8'), ERR.DIV0);
  assert.equal(val(wb, 'A9'), true);
  assert.equal(val(wb, 'A10'), true);
});

test('참조와 재계산', () => {
  const wb = book({ A1: '10', A2: '20', A3: '=A1+A2', B1: '=A3*2' });
  assert.equal(val(wb, 'B1'), 60);
  wb.transact(() => wb.setInput(0, 0, 0, '5'));
  assert.equal(val(wb, 'A3'), 25);
  assert.equal(val(wb, 'B1'), 50);
  wb.undo();
  assert.equal(val(wb, 'B1'), 60);
  wb.redo();
  assert.equal(val(wb, 'B1'), 50);
});

test('빈 셀 참조는 0, 순환 참조는 오류', () => {
  const wb = book({ A1: '=B1', C1: '=C2', C2: '=C1' });
  assert.equal(val(wb, 'A1'), 0);
  assert.equal(val(wb, 'C1'), ERR.CIRC);
});

test('집계 함수', () => {
  const wb = book({
    A1: '1', A2: '2', A3: 'x', A4: '4', A5: '',
    B1: '=SUM(A1:A5)', B2: '=AVERAGE(A1:A4)', B3: '=COUNT(A1:A5)', B4: '=COUNTA(A1:A5)',
    B5: '=MAX(A1:A4)', B6: '=MIN(A1:A4)', B7: '=SUM(A:A)', B8: '=MEDIAN(A1:A4)',
    B9: '=SUM(1,"2",TRUE)', B10: '=PRODUCT(A1:A4)', B11: '=COUNTBLANK(A1:A5)',
  });
  assert.equal(val(wb, 'B1'), 7);
  assert.ok(Math.abs(val(wb, 'B2') - 7 / 3) < 1e-12);
  assert.equal(val(wb, 'B3'), 3);
  assert.equal(val(wb, 'B4'), 4);
  assert.equal(val(wb, 'B5'), 4);
  assert.equal(val(wb, 'B6'), 1);
  assert.equal(val(wb, 'B7'), 7);
  assert.equal(val(wb, 'B8'), 2);
  assert.equal(val(wb, 'B9'), 4);
  assert.equal(val(wb, 'B10'), 8);
  assert.equal(val(wb, 'B11'), 1);
});

test('논리/오류 처리 함수', () => {
  const wb = book({
    A1: '5',
    B1: '=IF(A1>3,"큼","작음")', B2: '=IF(A1>9,1/0,"ok")', B3: '=IFERROR(1/0,"오류")',
    B4: '=AND(TRUE,A1>1)', B5: '=OR(FALSE,A1>10)', B6: '=NOT(A1=5)', B7: '=ISBLANK(C1)',
    B8: '=ISERROR(1/0)', B9: '=IFS(A1<3,"a",A1<6,"b")', B10: '=FOO(1)',
  });
  assert.equal(val(wb, 'B1'), '큼');
  assert.equal(val(wb, 'B2'), 'ok');
  assert.equal(val(wb, 'B3'), '오류');
  assert.equal(val(wb, 'B4'), true);
  assert.equal(val(wb, 'B5'), false);
  assert.equal(val(wb, 'B6'), false);
  assert.equal(val(wb, 'B7'), true);
  assert.equal(val(wb, 'B8'), true);
  assert.equal(val(wb, 'B9'), 'b');
  assert.equal(val(wb, 'B10'), ERR.NAME);
});

test('텍스트 함수', () => {
  const wb = book({
    A1: '  하늘  마음 ',
    B1: '=TRIM(A1)', B2: '=LEN("가나다")', B3: '=LEFT("abcdef",3)', B4: '=RIGHT("abcdef",2)',
    B5: '=MID("abcdef",2,3)', B6: '=UPPER("abc")', B7: '=CONCATENATE("a","b",1)',
    B8: '=SUBSTITUTE("a-b-c","-","+")', B9: '=FIND("c","abc")', B10: '=TEXT(1234.5,"#,##0.00")',
    B11: '=TEXT(0.256,"0.0%")', B12: '=VALUE("12.5")', B13: '=REPT("ab",3)',
  });
  assert.equal(val(wb, 'B1'), '하늘 마음');
  assert.equal(val(wb, 'B2'), 3);
  assert.equal(val(wb, 'B3'), 'abc');
  assert.equal(val(wb, 'B4'), 'ef');
  assert.equal(val(wb, 'B5'), 'bcd');
  assert.equal(val(wb, 'B6'), 'ABC');
  assert.equal(val(wb, 'B7'), 'ab1');
  assert.equal(val(wb, 'B8'), 'a+b+c');
  assert.equal(val(wb, 'B9'), 3);
  assert.equal(val(wb, 'B10'), '1,234.50');
  assert.equal(val(wb, 'B11'), '25.6%');
  assert.equal(val(wb, 'B12'), 12.5);
  assert.equal(val(wb, 'B13'), 'ababab');
});

test('조건부 집계와 찾기 함수', () => {
  const wb = book({
    A1: '사과', B1: '100', A2: '배', B2: '200', A3: '사과', B3: '300', A4: '포도', B4: '50',
    D1: '=SUMIF(A1:A4,"사과",B1:B4)', D2: '=COUNTIF(B1:B4,">=100")', D3: '=AVERAGEIF(A1:A4,"사*",B1:B4)',
    D4: '=VLOOKUP("배",A1:B4,2,FALSE)', D5: '=VLOOKUP("수박",A1:B4,2,FALSE)', D6: '=MATCH("포도",A1:A4,0)',
    D7: '=INDEX(A1:B4,3,2)', D8: '=SUMIFS(B1:B4,A1:A4,"사과",B1:B4,">150")', D9: '=COUNTIFS(A1:A4,"<>배")',
    D10: '=XLOOKUP("포도",A1:A4,B1:B4)',
  });
  assert.equal(val(wb, 'D1'), 400);
  assert.equal(val(wb, 'D2'), 3);
  assert.equal(val(wb, 'D3'), 200);
  assert.equal(val(wb, 'D4'), 200);
  assert.equal(val(wb, 'D5'), ERR.NA);
  assert.equal(val(wb, 'D6'), 4);
  assert.equal(val(wb, 'D7'), 300);
  assert.equal(val(wb, 'D8'), 300);
  assert.equal(val(wb, 'D9'), 3);
  assert.equal(val(wb, 'D10'), 50);
});

test('근사 VLOOKUP', () => {
  const wb = book({ A1: '0', B1: 'F', A2: '60', B2: 'D', A3: '80', B3: 'B', C1: '=VLOOKUP(75,A1:B3,2)' });
  assert.equal(val(wb, 'C1'), 'D');
});

test('날짜 함수와 자동 서식', () => {
  const wb = book({ A1: '=DATE(2024,3,15)', A2: '=YEAR(A1)', A3: '=MONTH(A1)', A4: '=DAY(A1+20)', A5: '2024-01-05' });
  assert.equal(val(wb, 'A1'), 45366);
  assert.equal(val(wb, 'A2'), 2024);
  assert.equal(val(wb, 'A3'), 3);
  assert.equal(val(wb, 'A4'), 4);
  assert.equal(wb.getCell(0, 0, 0).style.numFmt, 'date');
  assert.equal(formatValue(val(wb, 'A1'), wb.getCell(0, 0, 0).style).text, '2024-03-15');
  assert.equal(val(wb, 'A5'), 45296);
  const wb2 = book({ A1: '=TEXT(DATE(2024,3,5)+0.5,"yyyy-mm-dd hh:mm")', A2: '=TEXT(DATE(2024,3,5),"yyyy년 m월 d일")' });
  assert.equal(val(wb2, 'A1'), '2024-03-05 12:00');
  assert.equal(val(wb2, 'A2'), '2024년 3월 5일');
});

test('다른 시트 참조와 시트 이름 변경', () => {
  const wb = book({ A1: '=Sheet2!B2*2', A2: "='내 시트'!A1" });
  wb.transact(() => {
    wb.addSheet('Sheet2');
    wb.addSheet('내 시트');
    wb.setInput(1, 1, 1, '21');
    wb.setInput(2, 0, 0, 'hi');
  });
  assert.equal(val(wb, 'A1'), 42);
  assert.equal(val(wb, 'A2'), 'hi');
  wb.transact(() => wb.renameSheet(1, '데이터'));
  assert.equal(wb.getRaw(0, 0, 0), '=데이터!B2*2');
  assert.equal(val(wb, 'A1'), 42);
  assert.equal(renameSheetInFormula("='a b'!A1", 'a b', 'x y'), "='x y'!A1");
});

test('수식 복사 시 상대/절대 참조 이동', () => {
  assert.equal(shiftFormula('=A1+$B$2+C$3+$D4', 1, 1), '=B2+$B$2+D$3+$D5');
  assert.equal(shiftFormula('=SUM(A1:B2)', 2, 0), '=SUM(A3:B4)');
  assert.equal(shiftFormula('=A1', -1, 0), '=#REF!');
  assert.equal(shiftFormula('=SUM(A:A)', 5, 1), '=SUM(B:B)');
  assert.equal(shiftFormula('="A1"&A1', 1, 0), '="A1"&A2');
  assert.equal(shiftFormula('=LOG10(A1)', 1, 0), '=LOG10(A2)');
});

test('행/열 삽입·삭제 시 참조 조정', () => {
  const o = { targetSheet: 'Sheet1', hostSheet: 'Sheet1', axis: 'row' };
  assert.equal(adjustFormulaForStructure('=A1+A5', { ...o, index: 2, count: 2 }), '=A1+A7');
  assert.equal(adjustFormulaForStructure('=SUM(A1:A5)', { ...o, index: 1, count: -2 }), '=SUM(A1:A3)');
  assert.equal(adjustFormulaForStructure('=A3', { ...o, index: 1, count: -3 }), '=#REF!');
  assert.equal(adjustFormulaForStructure('=Other!A5', { ...o, index: 0, count: 1 }), '=Other!A5');

  const wb = book({ A1: '1', A2: '2', A3: '=A1+A2', B3: '=SUM(A1:A2)' });
  wb.transact(() => wb.insertRows(0, 1, 1));
  assert.equal(wb.getRaw(0, 3, 0), '=A1+A3');
  assert.equal(val(wb, 'A4'), 3);
  assert.equal(wb.getRaw(0, 3, 1), '=SUM(A1:A3)');
  wb.transact(() => wb.insertCols(0, 0, 1));
  assert.equal(wb.getRaw(0, 3, 1), '=B1+B3');
  wb.undo();
  wb.undo();
  assert.equal(wb.getRaw(0, 2, 0), '=A1+A2');
});

test('정렬', () => {
  const wb = book({ A1: '3', B1: '=A1*10', A2: '1', B2: '=A2*10', A3: '', A4: '2', B4: '=A4*10' });
  wb.transact(() => wb.sortRange(0, 0, 0, 3, 1, 0, true));
  assert.deepEqual([val(wb, 'A1'), val(wb, 'A2'), val(wb, 'A3')], [1, 2, 3]);
  assert.deepEqual([val(wb, 'B1'), val(wb, 'B2'), val(wb, 'B3')], [10, 20, 30]);
  assert.equal(val(wb, 'A4'), null);
});

test('수식 구문 오류', () => {
  assert.throws(() => parse('1+'));
  assert.throws(() => parse('SUM(1,2'));
  const wb = book({ A1: '=1+' });
  assert.equal(val(wb, 'A1'), ERR.NAME);
});

test('숫자 표시 형식', () => {
  assert.equal(formatGeneral(0.1 + 0.2), '0.3');
  assert.equal(formatGeneral(1 / 3), '0.333333333');
  assert.equal(formatGeneral(123456789012), '1.23457E+11');
  assert.equal(formatGeneral(-5), '-5');
  assert.equal(formatNumber(1234567.891, 'number'), '1,234,567.89');
  assert.equal(formatNumber(-1234, 'currency'), '-₩1,234');
  assert.equal(formatNumber(0.1234, 'percent', 1), '12.3%');
  assert.equal(formatNumber(1.5, 'fraction'), '1 1/2');
  assert.equal(formatNumber(45366, 'longdate'), '2024년 3월 15일 금요일');
});

test('입력 해석', () => {
  assert.deepEqual(parseInput('42'), { value: 42 });
  assert.deepEqual(parseInput('1,234'), { value: 1234, numFmt: 'comma' });
  assert.deepEqual(parseInput('15%'), { value: 0.15, numFmt: 'percent', decimals: undefined });
  assert.deepEqual(parseInput('₩5,000'), { value: 5000, numFmt: 'currency', decimals: undefined });
  assert.deepEqual(parseInput("'007"), { value: '007' });
  assert.deepEqual(parseInput('true'), { value: true });
  assert.equal(parseInput('12:30').numFmt, 'time');
  assert.deepEqual(parseInput('하늘마음'), { value: '하늘마음' });
});
