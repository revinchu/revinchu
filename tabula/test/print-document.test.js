import test from 'node:test';
import assert from 'node:assert/strict';
import { splitPrintIndexes, imagePagesPdf, printMergeMap } from '../src/print-document.js';

test('인쇄 행·열은 경계에서 나누고 반복 제목은 모든 페이지에 한 번만 배치', () => {
  assert.deepEqual(splitPrintIndexes([20, 30, 30, 30, 30], 80, [0]), [[0, 1, 2], [0, 3, 4]]);
  assert.deepEqual(splitPrintIndexes([10, 25, 30, 20], 60), [[0, 1], [2, 3]]);
  assert.deepEqual(splitPrintIndexes([100, 10], 50), [[0], [1]]);
  assert.deepEqual(splitPrintIndexes([], 100), [[]]);
  assert.deepEqual(splitPrintIndexes([10, 20], 100, [0, 0, -1, 8]), [[0, 1]]);
});

test('PDF 페이지수·용지치수·JPEG 스트림·한글 제목 및 xref 실제바이트 위치', () => {
  const jpeg = Uint8Array.of(255, 216, 255, 224, 0, 16, 255, 217);
  const bytes = imagePagesPdf([{ jpeg, width: 800, height: 1200, paperWidth: 595.44, paperHeight: 841.68 }, { jpeg, width: 1200, height: 800, paperWidth: 841.68, paperHeight: 595.44 }], { title: '한글 보고서' });
  const text = new TextDecoder('latin1').decode(bytes);
  assert.ok(text.startsWith('%PDF-1.4')); assert.match(text, /\/Count 2/);
  assert.match(text, /\/MediaBox \[0 0 595.44 841.68\]/); assert.match(text, /\/MediaBox \[0 0 841.68 595.44\]/);
  assert.equal((text.match(/\/Filter \/DCTDecode/g) ?? []).length, 2);
  assert.match(text, /\/Title <FEFFd55cae00/);
  const xref = Number(/startxref\n(\d+)/.exec(text)[1]); assert.equal(new TextDecoder().decode(bytes.slice(xref, xref + 4)), 'xref');
  const offsets = text.slice(text.indexOf('xref\n')).split('\n').slice(3, 12);
  offsets.forEach((line, i) => assert.ok(new TextDecoder().decode(bytes.slice(Number(line.slice(0, 10)), Number(line.slice(0, 10)) + 12)).startsWith(`${i + 1} 0 obj`)));
  assert.throws(() => imagePagesPdf([]), /페이지가 없습니다/);
  assert.throws(() => imagePagesPdf([{ jpeg: Uint8Array.of(1, 2), width: 1, height: 1, paperWidth: 1, paperHeight: 1 }]), /올바르지/);
});


test('병합 인쇄는 숨긴 행·열과 영역 밖 셀을 제외하고 원래 앵커 값·서식을 참조', () => {
  const cells = printMergeMap([{r1:0,c1:0,r2:3,c2:3}], [1,3,4], [1,3,4]);
  assert.deepEqual(cells.get('1,1'), {r:0,c:0,rowSpan:2,colSpan:2});
  assert.equal(cells.get('1,3'), null); assert.equal(cells.get('3,1'), null); assert.equal(cells.get('3,3'), null);
  assert.equal(cells.has('4,4'), false); assert.equal(cells.size, 4);
  assert.equal(printMergeMap([{r1:0,c1:0,r2:3,c2:3}], [8], [1]).size, 0);
});
