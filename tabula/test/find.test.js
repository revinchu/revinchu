// 찾기 및 바꾸기: 와일드카드 · 대/소문자 · 전체 일치 · 찾는 위치 · 순서 · 통합 문서 · 서식
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findRegex, replaceText, findMatches, nextMatch, formatMatches } from '../src/find.js';
import { Workbook } from '../src/workbook.js';

test('와일드카드: * ? ~ 와 대/소문자 · 전체 셀 내용 일치', () => {
  assert.ok(findRegex('검색*광고').test('검색 파워링크 광고'));
  assert.ok(findRegex('a?c').test('xabcx'));
  assert.ok(!findRegex('a?c', { whole: true }).test('xabcx'));
  assert.ok(findRegex('50~*', {}).test('50*2'));
  assert.ok(!findRegex('50~*', {}).test('500'));
  assert.ok(findRegex('ABC').test('abc'));
  assert.ok(!findRegex('ABC', { matchCase: true }).test('abc'));
  assert.equal(replaceText('파워링크 광고 광고', { text: '광고' }, 'AD'), '파워링크 AD AD');
  assert.equal(replaceText('abc', { text: 'ABC', whole: true }, 'X'), 'X');
  assert.equal(replaceText('abcd', { text: 'ABC', whole: true }, 'X'), 'abcd');
});

test('찾는 위치 · 검색 순서 · 서식 · 다음/이전', () => {
  const wb = new Workbook();
  wb.transact(() => {
    wb.setInput(0, 0, 1, '사과');
    wb.setInput(0, 1, 0, '사과나무');
    wb.setInput(0, 2, 2, '=A2&"!"');
    wb.setInput(0, 3, 0, '0.5');
    wb.setStyle(0, 1, 0, { bold: true });
  });
  const rows = findMatches(wb, { text: '사과', sheets: [0] });
  assert.deepEqual(rows.map((m) => [m.r, m.c]), [[0, 1], [1, 0]]);
  const cols = findMatches(wb, { text: '사과', sheets: [0], byCols: true });
  assert.deepEqual(cols.map((m) => [m.r, m.c]), [[1, 0], [0, 1]]);
  // 값에서 찾으면 수식 결과도 찾음
  const vals = findMatches(wb, { text: '사과', sheets: [0], lookIn: 'values' });
  assert.equal(vals.length, 3);
  // 수식에서 찾기: 수식 글자
  assert.equal(findMatches(wb, { text: 'A2', sheets: [0] }).length, 1);
  assert.equal(findMatches(wb, { text: 'A2', sheets: [0], lookIn: 'values' }).length, 0);
  // 서식으로 찾기
  const bold = findMatches(wb, { text: '사과', sheets: [0], format: { bold: true } });
  assert.deepEqual(bold.map((m) => [m.r, m.c]), [[1, 0]]);
  assert.ok(formatMatches({ bold: true, fill: '#FFFF00' }, { fill: '#ffff00' }));
  assert.ok(!formatMatches({}, { bold: true }));
  assert.ok(formatMatches({}, { bold: false }));
  // 다음 / 이전 찾기 (끝에서 처음으로)
  assert.deepEqual(nextMatch(rows, { si: 0, r: 0, c: 1 }), rows[1]);
  assert.deepEqual(nextMatch(rows, { si: 0, r: 1, c: 0 }), rows[0]);
  assert.deepEqual(nextMatch(rows, { si: 0, r: 0, c: 1 }, { back: true }), rows[1]);
});

test('시트 탭 색: xlsx 왕복', async () => {
  const { readXlsx, writeXlsx } = await import('../src/xlsx.js');
  const wb = new Workbook();
  wb.transact(() => { wb.setInput(0, 0, 0, '1'); wb.setSheetProp(0, 'tabColor', '#c00000'); });
  const back = new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[0].tabColor, '#c00000');
  // 자동 저장용 통합 문서 속성 (테마 · 기본 서식)
  wb.theme = ['FFFFFF', '000000'];
  wb.baseStyle = { valign: 'middle' };
  const meta = wb.bookMeta();
  assert.deepEqual(meta.theme, ['FFFFFF', '000000']);
  assert.deepEqual(meta.baseStyle, { valign: 'middle' });
});
