import test from 'node:test';
import assert from 'node:assert/strict';
import { createStoredStyleMemo } from '../src/cell-storage.js';
import { CellMap } from '../src/cellmap.js';
import { Workbook } from '../src/workbook.js';

const original = () => ({ fontFamily: '맑은 고딕', fontSize: 11, bold: true, fill: '#ddeeff', numFmt: 'number', decimals: 2, borders: { left: { color: '#112233', width: 2 }, right: { color: '#334455', width: 1 } } });
const decoded = style => JSON.parse(JSON.stringify(style));

test('5만 개 JSON 셀 서식을 하나로 공유하고 수식·저장값은 독립적으로 유지한다', () => {
  const memo = createStoredStyleMemo(), cells = new CellMap(), style = original();
  for (let i = 0; i < 50000; i++) cells.setRC(i, 0, { raw: '=1+1', cached: 2, style: memo(decoded(style)) });
  const styles = new Set(); for (const [, cell] of cells) styles.add(cell.style);
  assert.equal(styles.size, 1); assert.deepEqual([...styles][0], style);
  const book = new Workbook({ sheets: [{ name: '복구', cells }] });
  assert.equal(book.sheets[0].cells.size, 50000); assert.equal(book.getValue(0, 49999, 0), 2);
  assert.equal(book.getCell(0, 0, 0).style, book.getCell(0, 49999, 0).style);
  assert.notEqual(book.getCell(0, 0, 0), book.getCell(0, 1, 0));
});

test('중첩 테두리 공유 후 단일 셀 서식 편집·실행 취소가 이웃에 번지지 않는다', () => {
  const memo = createStoredStyleMemo(), cells = new CellMap(), before = original();
  cells.setRC(0, 0, { raw: '1', style: memo(decoded(before)) }); cells.setRC(1, 0, { raw: '2', style: memo(decoded(before)) });
  const book = new Workbook({ sheets: [{ name: '복구', cells }] });
  book.transact(() => book.setStyle(0, 0, 0, { fill: '#ff0000', borders: { left: { color: '#abcdef', width: 4 } } }));
  assert.equal(book.styleAt(0, 0, 0).fill, '#ff0000');
  assert.deepEqual(book.getCell(0, 1, 0).style, before);
  assert.equal(book.getCell(0, 1, 0).style.borders.left.width, 2);
  book.undo(); assert.deepEqual(book.getCell(0, 0, 0).style, before); assert.deepEqual(book.getCell(0, 1, 0).style, before);
  book.redo(); assert.equal(book.styleAt(0, 0, 0).fill, '#ff0000'); assert.deepEqual(book.getCell(0, 1, 0).style, before);
});

test('서식 메모 개수 한도 뒤에도 기존 서식만 공유하고 새 서식은 보존한다', () => {
  const memo = createStoredStyleMemo({ limit: 2 }), a = memo({ fill: 'a' }), b = memo({ fill: 'b' });
  const c = memo({ fill: 'c' }); assert.notEqual(memo({ fill: 'c' }), c);
  assert.equal(memo({ fill: 'a' }), a); assert.equal(memo({ fill: 'b' }), b);
});

test('긴 서식과 많은 고유 서식이 문자 한도를 넘겨 메모리에 쌓이지 않는다', () => {
  const memo = createStoredStyleMemo({ charLimit: 30 }), a = memo({ fill: 'a' });
  const large = memo({ code: '0'.repeat(1000) }); assert.notEqual(memo(decoded(large)), large);
  const b = memo({ fontFamily: 'x'.repeat(25) }); assert.notEqual(memo(decoded(b)), b);
  assert.equal(memo({ fill: 'a' }), a);
});

test('다른 중첩 값·자료형·명시 false와 빠진 필드를 합치지 않는다', () => {
  const memo = createStoredStyleMemo(), variants = [{ bold: false }, {}, { width: 1 }, { width: '1' }, { border: { color: '#a' } }, { border: { color: '#b' } }];
  const styles = variants.map(s => memo(decoded(s))); assert.equal(new Set(styles).size, variants.length);
  for (let i = 0; i < variants.length; i++) assert.equal(memo(decoded(variants[i])), styles[i]);
  assert.equal(memo(null), null); assert.equal(memo(undefined), undefined);
});
