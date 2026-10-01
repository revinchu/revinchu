import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook, cellData } from '../src/workbook.js';

const fixture = () => new Workbook({ sheets: [{ name: 'Sheet1', fileValues: true, cells: {
  '0,0': { raw: '=NOSUCHFUNCTION(1)', cached: 77, fx: true, style: { numFmt: 'text', fill: '#ffeecc' }, comment: '메모', link: 'https://example.com' },
  '1,0': { raw: '=HYPERLINK("https://example.com","수식 링크")', cached: '수식 링크', fx: true, style: { color: '#0563c1', underline: true } },
  '2,0': { raw: '', image: { src: 'data:image/png;base64,AA==', alt: '그림' }, style: { fill: '#eeeeee' }, link: 'https://example.com/image', comment: '그림 메모' },
} }] });

test('하이퍼링크 지우기는 링크 속성만 삭제하고 수식·서식·그림·캐시를 보존한다', () => {
  const wb = fixture();
  const before = [0, 1, 2].map((r) => cellData(wb.getCell(0, r, 0)));
  wb.transact(() => wb.clearRange(0, 0, 0, 2, 0, 'hyperlinks'));
  for (let r = 0; r < 3; r++) {
    const expected = { ...before[r] }; delete expected.link;
    assert.deepEqual(cellData(wb.getCell(0, r, 0)), expected);
  }
  assert.equal(wb.getValue(0, 0, 0), 77);
  assert.equal(wb.getCell(0, 1, 0).raw, before[1].raw);
  wb.undo();
  assert.deepEqual([0, 1, 2].map((r) => cellData(wb.getCell(0, r, 0))), before);
  wb.redo();
  assert.equal(wb.getCell(0, 0, 0).link, undefined);
});

test('서식·메모 지우기는 파일의 미지원 수식 저장값과 fx를 보존한다', () => {
  for (const mode of ['formats', 'comments']) {
    const wb = fixture();
    wb.transact(() => wb.clearRange(0, 0, 0, 2, 0, mode));
    assert.equal(wb.getCell(0, 0, 0).cached, 77, mode);
    assert.equal(wb.getCell(0, 0, 0).fx, true, mode);
    assert.equal(wb.getValue(0, 0, 0), 77, mode);
    assert.equal(wb.getCell(0, 0, 0).link, 'https://example.com', mode);
    assert.equal(wb.getCell(0, 2, 0).image.alt, '그림', mode);
    assert.equal(wb.getCell(0, 0, 0)[mode === 'formats' ? 'style' : 'comment'], undefined, mode);
    wb.undo(); assert.equal(wb.getCell(0, 0, 0).comment, '메모');
  }
});

test('내용·모두 지우기는 값·수식·링크를 지우고 다른 범위를 보존한다', () => {
  for (const mode of ['contents', 'all']) {
    const wb = fixture();
    const outside = cellData(wb.getCell(0, 1, 0));
    wb.transact(() => wb.clearRange(0, 0, 0, 0, 0, mode));
    if (mode === 'all') assert.equal(wb.getCell(0, 0, 0) ?? null, null);
    else assert.deepEqual(cellData(wb.getCell(0, 0, 0)), { raw: '', style: { numFmt: 'text', fill: '#ffeecc' }, comment: '메모' });
    assert.deepEqual(cellData(wb.getCell(0, 1, 0)), outside);
    wb.undo(); assert.equal(wb.getValue(0, 0, 0), 77);
  }
});

test('서식 지우기는 선택한 조건부 서식만 빼고 다른 범위와 상대 수식을 유지한다', () => {
  const wb = fixture();
  const rule = { r1: 0, c1: 0, r2: 2, c2: 2, more: [{ r1: 6, c1: 0, r2: 7, c2: 0 }], type: 'formula', formula: '=A1>0', style: { fill: '#ff0000' } };
  wb.transact(() => wb.setSheetProp(0, 'cond', [rule]));
  wb.transact(() => wb.clearRange(0, 0, 0, 0, 2, 'formats'));
  const next = wb.sheets[0].cond[0];
  assert.deepEqual([next.r1, next.c1, next.r2, next.c2], [1, 0, 2, 2]);
  assert.equal(next.formula, '=A2>0');
  assert.deepEqual(next.more, [{ r1: 6, c1: 0, r2: 7, c2: 0 }]);
  wb.undo(); assert.deepEqual(wb.sheets[0].cond, [rule]);
  wb.transact(() => wb.clearCondRules(0, { r1: 6, c1: 0, r2: 7, c2: 0 }));
  assert.equal(wb.sheets[0].cond[0].more, undefined);
  assert.equal(wb.sheets[0].cond[0].formula, '=A1>0');
  wb.transact(() => wb.clearCondRules(0, { r1: 1, c1: 1, r2: 1, c2: 1 }));
  const parts = [wb.sheets[0].cond[0], ...wb.sheets[0].cond[0].more];
  assert.equal(parts.reduce((n, g) => n + (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1), 0), 8);
  assert.ok(parts.every((g) => !(g.r1 <= 1 && g.r2 >= 1 && g.c1 <= 1 && g.c2 >= 1)));
});
