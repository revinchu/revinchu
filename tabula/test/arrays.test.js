import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { toFileFormula, fromFileFormula } from '../src/xlfn.js';
import { ERR } from '../src/formula.js';

function book(cells) {
  const wb = new Workbook();
  wb.transact(() => {
    for (const [a, raw] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      wb.setInput(0, Number(m[2]) - 1, c, raw);
    }
  });
  return wb;
}
const v = (wb, a, si = 0) => {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return wb.getValue(si, Number(m[2]) - 1, c);
};

test('동적 배열: 분산 · #SPILL! · A1# 참조', () => {
  const wb = book({ A1: '=SEQUENCE(3)', B1: '=A1#*10', C1: '=SUM(A1#)' });
  assert.equal(v(wb, 'A3'), 3);
  assert.equal(v(wb, 'B2'), 20);
  assert.equal(v(wb, 'C1'), 6);
  assert.deepEqual(wb.spillRange(0, 0, 0), { r1: 0, c1: 0, r2: 2, c2: 0 });
  assert.deepEqual(wb.spillAnchorOf(0, 2, 1), { r: 0, c: 1 });
  wb.transact(() => wb.setInput(0, 1, 0, 'x'));
  assert.equal(v(wb, 'A1'), ERR.SPILL);
  assert.equal(v(wb, 'A2'), 'x');
  assert.equal(v(wb, 'C1'), ERR.REF);
});

test('배열 함수: FILTER · SORT · UNIQUE · XLOOKUP · LET · LAMBDA', () => {
  const wb = book({
    A1: '서울', A2: '부산', A3: '서울', A4: '대구',
    B1: '10', B2: '20', B3: '30', B4: '5',
    D1: '=UNIQUE(A1:A4)', E1: '=SORT(B1:B4,,-1)', F1: '=FILTER(A1:A4,B1:B4>=20)',
    G1: '=XLOOKUP("부산",A1:A4,B1:B4)', H1: '=LET(x,SUM(B1:B4),x/4)', I1: '=MAP(B1:B2,LAMBDA(v,v*2))',
    J1: '=SUMPRODUCT((A1:A4="서울")*B1:B4)', K1: '=TEXTJOIN("-",TRUE,UNIQUE(A1:A4))',
  });
  assert.deepEqual([v(wb, 'D1'), v(wb, 'D2'), v(wb, 'D3')], ['서울', '부산', '대구']);
  assert.deepEqual([v(wb, 'E1'), v(wb, 'E4')], [30, 5]);
  assert.deepEqual([v(wb, 'F1'), v(wb, 'F2')], ['부산', '서울']);
  assert.equal(v(wb, 'G1'), 20);
  assert.equal(v(wb, 'H1'), 16.25);
  assert.equal(v(wb, 'I2'), 40);
  assert.equal(v(wb, 'J1'), 40);
  assert.equal(v(wb, 'K1'), '서울-부산-대구');
});

test('이름 정의: 범위 · 상수 · LAMBDA · 시트 범위', () => {
  const wb = book({ A1: '1', A2: '2', A3: '3', B1: '=SUM(값)', B2: '=세율*100', B3: '=두배(21)', B4: '=ROWS(값)' });
  wb.transact(() => wb.setNames([
    { name: '값', ref: '=Sheet1!$A$1:$A$3' },
    { name: '세율', ref: '=0.1' },
    { name: '두배', ref: '=LAMBDA(x,x*2)' },
  ]));
  assert.equal(v(wb, 'B1'), 6);
  assert.equal(v(wb, 'B2'), 10);
  assert.equal(v(wb, 'B3'), 42);
  assert.equal(v(wb, 'B4'), 3);
  // 행 삽입 시 이름 참조도 이동
  wb.transact(() => wb.insertRows(0, 0, 1));
  assert.equal(wb.names[0].ref, '=Sheet1!$A$2:$A$4');
  assert.equal(v(wb, 'B2'), 6);
});

test('파일 수식 변환: 접두사 · SINGLE · ANCHORARRAY · 옛 형식 암시적 교차', () => {
  assert.equal(toFileFormula('=XLOOKUP(A1,B:B,C:C)'), '_xlfn.XLOOKUP(A1,B:B,C:C)');
  assert.equal(toFileFormula('=SORT(FILTER(A1:A9,B1:B9>0))'), '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))');
  assert.equal(toFileFormula('=LET(x,2,x*x)'), '_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlpm.x)');
  assert.equal(toFileFormula('=SUM(A1#)'), 'SUM(_xlfn.ANCHORARRAY(A1))');
  assert.equal(toFileFormula('=@A1:A9', { dynamic: true }), '_xlfn.SINGLE(A1:A9)');
  assert.equal(toFileFormula('=GROUPBY(A1:A9,B1:B9,SUM)'), '_xlfn.GROUPBY(A1:A9,B1:B9,_xleta.SUM)');
  assert.equal(fromFileFormula('_xlfn.LET(_xlpm.x,2,_xlpm.x*_xlpm.x)'), 'LET(x,2,x*x)');
  assert.equal(fromFileFormula('SUM(_xlfn.ANCHORARRAY(A1))'), 'SUM(A1#)');
  assert.equal(fromFileFormula('A1:A10*2', { legacy: true }), '@A1:A10*2');
  assert.equal(fromFileFormula('SUM(A1:A10)', { legacy: true }), 'SUM(A1:A10)');
});

test('xlsx 왕복: 동적 배열 · 이름 · 지원하지 않는 함수의 저장된 값', () => {
  const wb = book({ A1: '=SEQUENCE(2,2)', D1: '=SUM(A1#)', E1: '=금액*2', F1: '5' });
  wb.transact(() => wb.setNames([{ name: '금액', ref: '=Sheet1!$F$1' }]));
  const bytes = writeXlsx(wb);
  const files = unzip(bytes);
  const sheetXml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.match(sheetXml, /<c r="A1" cm="1"><f t="array" ref="A1:B2" aca="false">_xlfn\.SEQUENCE\(2,2\)<\/f><v>1<\/v><\/c>/);
  assert.match(sheetXml, /<c r="B2"><v>4<\/v><\/c>/);
  assert.match(sheetXml, /_xlfn\.ANCHORARRAY\(A1\)/);
  assert.ok(files['xl/metadata.xml']);
  assert.match(textOf(files['xl/workbook.xml']), /<definedName name="금액">Sheet1!\$F\$1<\/definedName>/);

  const { data } = readXlsx(bytes);
  const back = new Workbook(data);
  assert.equal(back.getRaw(0, 0, 0), '=SEQUENCE(2,2)');
  assert.equal(back.getRaw(0, 1, 1), ''); // 분산 값은 다시 계산됨
  assert.equal(back.getValue(0, 1, 1), 4);
  assert.equal(back.getValue(0, 0, 3), 10);
  assert.equal(back.getValue(0, 0, 4), 10);
  assert.equal(back.names[0].name, '금액');

  // 지원하지 않는 함수: 수식 유지 + 파일의 계산 값 표시
  const xml = sheetXml.replace('<c r="F1"><v>5</v></c>', '<c r="F1"><f>_xlfn.NOSUCHFN(1)</f><v>77</v></c>');
  const patched = { ...files, 'xl/worksheets/sheet1.xml': new TextEncoder().encode(xml) };
  const res = readXlsx(zip(patched));
  const wb2 = new Workbook(res.data);
  assert.equal(wb2.getRaw(0, 0, 5), '=NOSUCHFN(1)');
  assert.equal(wb2.getValue(0, 0, 5), 77);
  assert.ok(res.warnings.some((w) => w.includes('저장된 계산 결과')));
});
