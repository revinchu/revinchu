import { test } from 'node:test';
import assert from 'node:assert/strict';
import { CellMap, shareBlankCell } from '../src/cellmap.js';
import { Workbook, makeCellRC } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

test('대량 서식 빈 셀: 좌표와 논리 셀 수를 유지하고 불변 값만 공유', () => {
  const cells = new CellMap(), styles = [{ fill: '#ff0000', bb: true, bbs: 'thin', bbc: '#112233' }, { italic: true, font: '맑은 고딕', size: 13 }];
  for (let r = 0; r < 100000; r++) cells.setRC(r, r % 2, { raw: '', style: styles[r % 2] });
  assert.equal(cells.size, 100000);
  assert.equal(cells.getRC(0, 0), cells.getRC(99998, 0));
  assert.notEqual(cells.getRC(0, 0), cells.getRC(1, 1));
  assert.ok(Object.isFrozen(cells.getRC(0, 0)));
  const wb = new Workbook({ sheets: [{ name: '서식', cells }] });
  assert.equal(wb.sheets[0].cells.size, 100000);
  assert.equal(wb.getCell(0, 0, 0), wb.getCell(0, 99998, 0));
  assert.equal(wb.getValue(0, 99998, 0), null);
  wb.transact(() => wb.setInput(0, 0, 0, '123'));
  assert.equal(wb.getValue(0, 0, 0), 123);
  assert.equal(wb.getValue(0, 2, 0), null);
  wb.undo(); assert.equal(wb.getValue(0, 0, 0), null);
  wb.redo(); assert.equal(wb.getValue(0, 0, 0), 123);
  wb.transact(() => wb.setStyle(0, 2, 0, { fill: '#00ff00' }));
  assert.equal(wb.styleAt(0, 2, 0).fill, '#00ff00');
  assert.equal(wb.styleAt(0, 4, 0).fill, '#ff0000');
  wb.undo(); assert.equal(wb.styleAt(0, 2, 0).fill, '#ff0000');
});

test('빈 셀의 메모·링크·그림·입력 힌트·저장값은 공유하지 않음', () => {
  const style = { bold: true }, plain = { raw: '', style };
  for (const props of [{ comment: '메모' }, { link: '#시트!A1' }, { image: { src: 'data:image/png;base64,AA==' } }, { inputType: 'text' }, { cached: 0 }, { staleCached: 12 }, { fx: true }, { phonetic: {} }]) {
    const value = { ...plain, ...props };
    assert.equal(shareBlankCell(value), value);
    assert.notEqual(makeCellRC(value, 0, 0), makeCellRC(plain, 0, 0));
  }
  assert.notEqual(shareBlankCell(plain), shareBlankCell({ ...plain, v: null }));
});

test('공유 빈 서식 셀: 구조 편집·Undo·JSON/Blob/XLSX 왕복 보존', async () => {
  const cells = new CellMap(), style = { fill: '#fff2cc', bold: true, italic: true, font: 'Arial', size: 12, br: true, brs: 'double', brc: '#123456' };
  for (let r = 0; r < 32; r++) cells.setRC(r, 2, { raw: '', style });
  const wb = new Workbook({ sheets: [{ name: '서식', cells }] });
  const before = JSON.stringify(wb.serialize());
  assert.equal(await wb.serializeBlob().text(), before);
  wb.transact(() => wb.insertRows(0, 4, 2));
  assert.equal(wb.styleAt(0, 33, 2).fill, '#fff2cc');
  assert.equal(wb.getCell(0, 4, 2), undefined);
  wb.undo(); assert.equal(JSON.stringify(wb.serialize()), before);
  for (const restored of [new Workbook(JSON.parse(before)), new Workbook(readXlsx(writeXlsx(wb)).data)]) {
    assert.equal(restored.sheets[0].cells.size, 32);
    for (let r = 0; r < 32; r++) assert.deepEqual(restored.styleAt(0, r, 2), style);
  }
});
