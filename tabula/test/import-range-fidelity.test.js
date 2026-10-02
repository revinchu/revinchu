import { test } from 'node:test';
import assert from 'node:assert/strict';
import { importRangeSource, parseImportRange } from '../src/fx-web.js';

const ID = 'synthetic_mixed_sheet_123456';
test('Google 일반 범위는 열 타입 추론이 없는 CSV export를 사용한다', () => {
  const source = importRangeSource(`https://docs.google.com/spreadsheets/d/${ID}/edit?gid=42#gid=99`, 'A1:Z1000');
  const url = new URL(source.url);
  assert.equal(url.pathname, `/spreadsheets/d/${ID}/export`);
  assert.deepEqual(Object.fromEntries(url.searchParams), { format: 'csv', range: 'A1:Z1000' });
  assert.equal(source.crop, undefined, 'export 범위를 다시 로컬에서 잘라 위치를 잃지 않는다');
});
test('혼합 자료형·헤더·완전 빈 행·빈 열의 원래 좌표를 유지한다', () => {
  const csv = '\r\n구분,1월,,합계\r\n목표,"123,456",,9\r\n,,,\r\n,,,\r\n상세,,,\r\n1,12.5%,,=1+1\r\n2,문자,,001\r\n';
  const out = parseImportRange(csv);
  assert.equal(out.height, 8); assert.equal(out.width, 4);
  assert.deepEqual(out.rows, [['','','',''],['구분','1월','','합계'],['목표',123456,'',9],['','','',''],['','','',''],['상세','','',''],[1,0.125,'','=1+1'],[2,'문자','','001']]);
  assert.equal(out.cellFormats[2][1].code, '#,##0');
  assert.equal(out.cellFormats[6][1].code, '0.0%');
  assert.equal(out.cellFormats[6][1].shrink, true);
});
test('CSV 인용부호·줄바꿈·공백·한글·0·false를 보존한다', () => {
  const out = parseImportRange('" 앞뒤 ","줄1\n줄2","따옴표""글",0,false\r\n');
  assert.deepEqual(out.rows, [[' 앞뒤 ', '줄1\n줄2', '따옴표"글', 0, false]]);
});
test('게시 CSV offset 범위와 표시 힌트는 함께 같은 위치로 자른다', () => {
  const out = parseImportRange('A,B,C,D\n가,1,"9,876.50",10\n나,2,2.00%,11\n다,3,끝,12', 'B2:C3');
  assert.deepEqual(out.rows, [[1,9876.5],[2,0.02]]);
  assert.equal(out.cellFormats[0][1].code, '#,##0.00');
  assert.equal(out.cellFormats[1][1].code, '0.00%');
});
test('원본 CSV의 맨 끝 개행은 가짜 데이터 행을 만들지 않는다', () => {
  assert.deepEqual(parseImportRange('A,B\r\n1,2\r\n').rows, [['A','B'],[1,2]]);
  assert.deepEqual(parseImportRange('A,B\r\n,\r\n').rows, [['A','B'],['','']]);
});

test('끝의 빈 행·열도 명시한 유한 범위 크기로 유지한다', () => {
  const source = importRangeSource(ID, 'B3:D8');
  assert.deepEqual(source.shape, { height: 6, width: 3 });
  const out = parseImportRange('제목,값\n항목,12.5%', null, source.shape);
  assert.equal(out.height, 6); assert.equal(out.width, 3);
  assert.deepEqual(out.rows[0], ['제목','값','']);
  assert.deepEqual(out.rows[1], ['항목',0.125,'']);
  assert.deepEqual(out.rows[5], ['','','']);
});
test('열린 범위는 끝없는 행을 만들지 않고 닫힌 축만 채운다', () => {
  const source = importRangeSource(ID, 'B3:D');
  assert.deepEqual(source.shape, { height: null, width: 3 });
  assert.deepEqual(parseImportRange('X,4', null, source.shape).rows, [['X',4,'']]);
  assert.deepEqual(parseImportRange('', null, { height: 2, width: 2 }).rows, [['',''],['','']]);
  for (const area of ['C4:A2','A100:C2','A1:Z1000000']) assert.throws(() => importRangeSource(ID, area));
});
