import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { Range, ERR } from '../src/fxcore.js';
import { WEB, NET, LOADING, netClear, importRangeSource, parseImportRange, parseQuery, runQuery } from '../src/fx-web.js';
import { formatValue, formatQuery } from '../src/format.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const ID = '1XWJLkAwch5GXAt_7zOFDcg8Wm8Xv29_8PWuuW15qmAE'; // Google Charts 공식 공개 예제
const query = (rows, text, headers = 0) => WEB.QUERY([new Range(rows), text, headers]);
const execute = (rows, text) => runQuery(parseQuery(text), rows, ['A', 'B', 'C'], ['', '', '']);

test('IMPORTRANGE: 공유 주소·ID·gid·한글/따옴표 시트 이름과 열린 범위', () => {
  const a = new URL(importRangeSource(`https://docs.google.com/spreadsheets/d/${ID}/edit#gid=7`, "'매출''분석'!$A$1:$C").url);
  assert.equal(a.hostname, 'docs.google.com');
  assert.equal(a.searchParams.get('headers'), '0');
  assert.equal(a.searchParams.get('sheet'), "매출'분석");
  assert.equal(a.searchParams.get('range'), 'A1:C');
  assert.equal(a.searchParams.get('gid'), null);
  const b = new URL(importRangeSource(`https://docs.google.com/spreadsheets/d/${ID}/edit?access_token=secret#gid=7`, 'A:C').url);
  assert.equal(b.searchParams.get('gid'), '7');
  assert.equal(b.searchParams.get('access_token'), null);
  assert.equal(importRangeSource(ID, 'A1:C').url, importRangeSource(`https://docs.google.com/spreadsheets/d/${ID}`, 'A1:C').url);
  assert.match(importRangeSource('월별 보고서', 'Sheet1!A1:C10').url, /^wixel-doc:/);
});

test('IMPORTRANGE: 호스트 위장·자격 증명·포트·잘못된 범위를 요청 전에 거부', () => {
  for (const src of [`https://evil.example/spreadsheets/d/${ID}`, `https://docs.google.com.evil.example/spreadsheets/d/${ID}`, `https://user:pass@docs.google.com/spreadsheets/d/${ID}`, `https://docs.google.com:8443/spreadsheets/d/${ID}`, 'javascript:alert(1)']) {
    assert.throws(() => importRangeSource(src, 'A1:C'));
  }
  for (const spec of ['A0:C4', 'A1:C&headers=1', '이름정의', "'깨진 시트!A1:C", '']) assert.throws(() => importRangeSource(ID, spec));
});

test('IMPORTRANGE: 게시 CSV는 gid를 유지하고 범위를 로컬에서 자른다', () => {
  const s = importRangeSource('https://docs.google.com/spreadsheets/d/e/2PACX-public/pubhtml?gid=10', 'B2:C');
  assert.equal(new URL(s.url).searchParams.get('output'), 'csv');
  assert.equal(new URL(s.url).searchParams.get('gid'), '10');
  assert.deepEqual(parseImportRange('a,b,c\r\n1,"x,y",3\r\n2,"line\n2",4\r\n', s.crop).rows, [['x,y', 3], ['line\n2', 4]]);
  assert.throws(() => importRangeSource('https://docs.google.com/spreadsheets/d/e/2PACX-public/pubhtml', '이름!A1:C'), /gid/);
});

test('IMPORTRANGE: 공식 공개 표 CSV 첫 행·빈 열 유지, HTML/JSONP 오류 차단', () => {
  // https://developers.google.com/chart/interactive/docs/spreadsheets 의 공개 시트 A1:C3 응답.
  const csv = '"LinearOptimizationSheet","","Variable Name"\n"","","Type"\n"","","Lower Bound"';
  assert.deepEqual(parseImportRange(csv).rows, [['LinearOptimizationSheet', '', 'Variable Name'], ['', '', 'Type'], ['', '', 'Lower Bound']]);
  assert.deepEqual(parseImportRange('\ufeff" code ",001,"x""y"\n').rows, [[' code ', '001', 'x"y']]);
  for (const html of ['<!DOCTYPE html><html>로그인</html>', '<html><body>Access denied</body></html>', '/*O_o*/\ngoogle.visualization.Query.setResponse({"status":"error"});']) assert.throws(() => parseImportRange(html), /공개|보기 허용/);
});

test('IMPORTRANGE + QUERY: 비동기 로딩·캐시 공유·시트 무효화·F9 새로 고침', async () => {
  const old = { fetcher: NET.fetcher, onDone: NET.onDone, maxAge: NET.maxAge };
  NET.cache.clear();
  let calls = 0, amount = 20;
  NET.fetcher = async () => { calls++; return `매체,매출\n검색,${amount}\n영상,10`; };
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 0, `=QUERY(IMPORTRANGE("${ID}","A1:B"),"select Col1, sum(Col2) group by Col1 order by sum(Col2) desc",1)`);
    wb.addSheet('다른 시트');
    wb.setInput(1, 0, 0, `=IMPORTRANGE("${ID}","A1:B")`);
  });
  NET.onDone = (sheets) => sheets.forEach((si) => wb.invalidate(si));
  try {
    assert.equal(wb.getValue(0, 0, 0), LOADING);
    assert.equal(wb.getValue(1, 0, 0), LOADING);
    await new Promise((r) => setImmediate(r));
    assert.equal(calls, 1);
    assert.equal(wb.getValue(0, 1, 1), 20);
    assert.equal(wb.getValue(1, 0, 0), '매체');
    amount = 30;
    const dirty = netClear();
    assert.deepEqual(dirty.sort(), [0, 1]);
    dirty.forEach((si) => wb.invalidate(si));
    assert.equal(wb.getValue(0, 0, 0), LOADING);
    await new Promise((r) => setImmediate(r));
    assert.equal(wb.getValue(0, 1, 1), 30);
    assert.equal(calls, 2);
    NET.cache.values().next().value.t = 0;
    wb.invalidate(1);
    assert.equal(wb.getValue(1, 0, 0), LOADING);
    await new Promise((r) => setImmediate(r));
    assert.equal(calls, 3);
  } finally { Object.assign(NET, old); NET.cache.clear(); }
});

test('QUERY: 소수 자료형·빈 값은 null, NOT/부등식은 null 행을 선택하지 않음', () => {
  const rows = [['검색', 20], ['영상', '문자'], ['검색', 10], ['검색', null]];
  assert.deepEqual(query(rows, 'select Col1, Col2 where Col2 != 20').rows, [['검색', 10]]);
  assert.deepEqual(query(rows, 'select Col1 where not Col2 = 20').rows, [['검색']]);
  assert.deepEqual(query(rows, 'select Col1 where Col2 is null').rows, [['영상'], ['검색']]);
  assert.deepEqual(query([['A', 1], ['a', 2]], "select Col2 where Col1 = 'a'").rows, [[2]]);
});

test('QUERY: 여러 머리글·생략 추측·빈 label·Col 표기·빈 결과', () => {
  assert.deepEqual(query([['매체', '성과'], ['이름', '매출'], ['검색', 4]], 'select Col1,Col2', 2).rows, [['매체 이름', '성과 매출'], ['검색', 4]]);
  assert.deepEqual(query([['매체', '매출'], ['검색', 4]], 'select Col1,Col2', -1).rows, [['매체', '매출'], ['검색', 4]]);
  assert.deepEqual(query([['매출'], [4], [8]], "select sum(Col1) label sum(Col1) ''", 1).rows, [[12]]);
  assert.equal(query([['매체'], ['검색']], "select Col1 where Col1 = '없는 값'", 1), ERR.NA);
  assert.equal(query([[1]], 'select *', -2), ERR.VALUE);
  assert.equal(query([[1]], 'select *', 0.5), ERR.VALUE);
  assert.deepEqual(WEB.QUERY([new Range([[2, 3]], { c1: 1 }), 'select Col2', 0]).rows, [[3]]);
});

test('QUERY: 잘못된 구문·절 순서·집계/그룹 조합은 무음 실행하지 않음', () => {
  for (const text of ['select Col1 @ garbage', 'select Col1 limit nope', 'select Col1 limit -1', 'select Col1 offset 1.5', 'select Col1 label Col1 unquoted', 'select Col1 format Col1', 'select Col1 select Col2', 'where Col1>0 select Col1', 'select sum(Col1+Col2)', 'select sum(Col1),Col2', 'select Col2 group by Col1', 'select Col1 where sum(Col2)>0', 'select sum(Col2) pivot Col1 order by sum(Col2)', 'select Col1 options unknown', 'select Col1 label Col2 "x"']) assert.equal(query([[1, 2], [1, 3]], text), ERR.VALUE, text);
  assert.equal(query([], 'select Missing'), ERR.VALUE);
});

test('QUERY: 여러 피벗 필드·집계 열 순서·명시 label·다중 정렬·offset/limit', () => {
  const rows = [['a', 2, 10], ['a', 1, 20], ['b', 1, 30], ['a', 2, 40]];
  const result = execute(rows, "select sum(C),count(C) pivot A,B label sum(C) '합계',count(C) '수'");
  assert.deepEqual(result.head, ['a,1 합계', 'a,2 합계', 'b,1 합계', 'a,1 수', 'a,2 수', 'b,1 수']);
  assert.deepEqual(result.rows, [[20, 50, 30, 1, 2, 1]]);
  assert.deepEqual(execute(rows, 'select A,B,C order by A asc,C desc limit 2 offset 1').rows, [['a', 1, 20], ['a', 2, 10]]);
});

test('QUERY format: 숫자·날짜·불리언 표시, no_format/no_values, 스필과 xlsx 왕복', () => {
  const rows = [[1234.5, 0.125, true]];
  const result = query(rows, "select Col1,Col2,Col3 format Col1 '#,##0.00',Col2 '0.0%',Col3 '예:아니오'");
  assert.deepEqual(result.rows, rows);
  assert.equal(formatQuery(1234.5, result.formats[0]), '1,234.50');
  assert.equal(formatQuery(0.125, result.formats[1]), '12.5%');
  assert.equal(formatQuery(true, result.formats[2]), '예');
  assert.equal(formatQuery(45292, 'yyyy-MM-dd'), '2024-01-01');
  assert.deepEqual(query(rows, "select Col1 format Col1 '0.00' options no_values").rows, [['1234.50']]);
  assert.equal(query(rows, "select Col1 format Col1 '0.00' options no_format").formats, undefined);
  const wb = new Workbook();
  wb.transact(() => wb.setInput(0, 0, 0, '=QUERY({1234.5;2},"select Col1 format Col1 \'#,##0.00\'",0)'));
  assert.equal(wb.getValue(0, 0, 0), 1234.5);
  assert.equal(formatValue(wb.getValue(0, 1, 0), wb.styleAt(0, 1, 0)).text, '2.00');
  wb.transact(() => wb.setInput(0, 0, 2, '=SUM(A1:A2)'));
  assert.equal(wb.getValue(0, 0, 2), 1236.5);
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.getValue(0, 0, 0), 1234.5);
  assert.equal(formatValue(back.getValue(0, 1, 0), back.styleAt(0, 1, 0)).text, '2.00');
  const single = new Workbook();
  single.transact(() => single.setInput(0, 0, 0, '=QUERY({1.5},"select Col1 format Col1 \'0.00\'",0)'));
  assert.equal(single.getValue(0, 0, 0), 1.5);
  assert.equal(formatValue(single.getValue(0, 0, 0), single.styleAt(0, 0, 0)).text, '1.50');
});

test('QUERY: 날짜·시간 리터럴 검사와 .5 산술', () => {
  assert.deepEqual(query([[1]], "select Col1+.5 where date '2024-02-29' > date '2024-01-01'").rows, [['+0.5'], [1.5]]);
  assert.equal(query([[1]], "select Col1 where date '2024-02-30' > date '2024-01-01'"), ERR.VALUE);
  assert.equal(query([[1]], "select Col1 where timeofday '25:00:00' > timeofday '00:00:00'"), ERR.VALUE);
});
