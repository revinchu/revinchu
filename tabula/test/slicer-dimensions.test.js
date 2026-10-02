import { test } from 'node:test';
import assert from 'node:assert/strict';
import { slicerDimensions, slicerDimensionPatch, SLICER_COLUMNS_MAX } from '../src/slicer-properties.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';

const cm = px => px * 2.54 / 96;
const close = (actual, expected) => assert.ok(Math.abs(actual - expected) <= 1 / 9525 + 1e-8, `${actual} ≈ ${expected}`);
const sl = { id: 's1', x: 200.125, y: 10.375, w: 192, h: 240, source: { kind: 'table', table: 'DataTable', column: 'Region' }, caption: '지역' };
function book(slicer) {
  const wb = new Workbook();
  wb.transact(() => {
    [['Region', 'Value'], ['A', 1], ['B', 2], ['C', 3]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.setSheetProp(0, 'tables', [{ id: 't1', name: 'DataTable', r1: 0, c1: 0, r2: 3, c2: 1, header: true, totals: false, columns: ['Region', 'Value'] }]);
    wb.setSheetProp(0, 'slicers', [slicer]);
  });
  return wb;
}

test('Excel 합성 실측과 같은 단추 폭: 열 변경은 전체 폭 유지, 단추 폭 변경은 전체 폭 조정', () => {
  const original = structuredClone(sl);
  assert.equal(slicerDimensions(sl).buttonWidth, 174);
  const two = { ...sl, ...slicerDimensionPatch(sl, 'columns', 2) };
  assert.equal(two.w, 192); assert.equal(slicerDimensions(two).buttonWidth, 85.5);
  const wide = { ...two, ...slicerDimensionPatch(two, 'buttonWidth', cm(96)) };
  assert.equal(wide.w, 213); assert.equal(slicerDimensions(wide).buttonWidth, 96);
  const three = { ...wide, ...slicerDimensionPatch(wide, 'columns', 3) };
  assert.equal(three.w, 213); assert.equal(slicerDimensions(three).buttonWidth, 63);
  assert.deepEqual(sl, original);
});

test('cm 입력은 소수 치수와 표시값 무변경을 구분하고 전체 높이를 임의로 바꾸지 않음', () => {
  assert.deepEqual(slicerDimensionPatch(sl, 'buttonHeight', cm(26.125)), { buttonHeight: 26.125 });
  assert.equal(slicerDimensions({ ...sl, buttonHeight: 26.125 }).h, 240);
  assert.deepEqual(slicerDimensionPatch({ ...sl, buttonHeight: 26.125 }, 'buttonHeight', cm(26.125).toFixed(2)), {});
  assert.deepEqual(slicerDimensionPatch(sl, 'h', cm(241.125)), { h: 241.125 });
  assert.deepEqual(slicerDimensionPatch(sl, 'w', cm(193.375)), { w: 193.375, buttonWidth: undefined });
});

test('기존 고정 단추 폭은 새 너비·열 변경 때만 해제하고 간격을 계산에 반영', () => {
  const legacy = { ...sl, buttonWidth: 55, columns: 2, gap: 5 };
  assert.equal(slicerDimensions(legacy).buttonWidth, 55);
  assert.deepEqual(slicerDimensionPatch(legacy, 'buttonHeight', cm(30)), { buttonHeight: 30 });
  assert.deepEqual(slicerDimensionPatch(legacy, 'buttonWidth', cm(60)), { w: 143, buttonWidth: undefined });
  assert.deepEqual(slicerDimensionPatch(legacy, 'columns', 3), { columns: 3, buttonWidth: undefined });
});

test('열 수 20000까지 상수 시간 계산, 음수 폭·잘못된 크기·잠긴 개체 변경 거부', () => {
  assert.equal(SLICER_COLUMNS_MAX, 20000);
  assert.equal(slicerDimensions({ ...sl, columns: 20000 }).buttonWidth, 0);
  assert.deepEqual(slicerDimensionPatch(sl, 'columns', 20000), { columns: 20000, buttonWidth: undefined });
  for (const value of ['', 0, -1, 1.5, 20001, NaN, Infinity]) assert.throws(() => slicerDimensionPatch(sl, 'columns', value));
  for (const kind of ['w', 'h', 'buttonWidth', 'buttonHeight']) for (const value of ['', -1, NaN, Infinity, 1000]) assert.throws(() => slicerDimensionPatch(sl, kind, value));
  for (const [kind, value] of [['w', 10], ['h', 10], ['buttonWidth', 10], ['buttonHeight', 1], ['columns', 2]]) assert.throws(() => slicerDimensionPatch({ ...sl, noMove: true }, kind, value), /해제/);
  assert.deepEqual(slicerDimensionPatch({ ...sl, noMove: true }, 'w', cm(sl.w)), {});
  assert.throws(() => slicerDimensionPatch(sl, 'unknown', 1), /항목/);
});

for (const height of [10, 24, 25.3, 26.125, 40.5]) {
  test(`슬라이서 단추 높이 ${height}px와 소수 전체 치수의 표준 XLSX 왕복`, () => {
    const wb = book({ ...sl, buttonHeight: height, columns: 3, w: 300.125, h: 241.375 }), before = wb.serialize();
    const bytes = writeXlsx(wb), files = unzip(bytes), xml = textOf(files['xl/slicers/slicer1.xml']);
    assert.match(xml, new RegExp(`rowHeight="${Math.round(height * 9525)}"`));
    const back = new Workbook(readXlsx(bytes).data).sheets[0].slicers[0];
    close(back.buttonHeight, height); close(back.w, 300.125); close(back.h, 241.375); close(back.x, sl.x); close(back.y, sl.y);
    assert.equal(back.columns, 3); close(slicerDimensions(back).buttonWidth, slicerDimensions(wb.sheets[0].slicers[0]).buttonWidth);
    assert.deepEqual(back.source, sl.source); assert.deepEqual(wb.serialize(), before);
  });
}

test('기본 24px 화면 높이도 파일에서 유지하고 표준 boolean 문자열을 읽음', () => {
  const files = unzip(writeXlsx(book(sl)));
  files['xl/slicers/slicer1.xml'] = textOf(files['xl/slicers/slicer1.xml']).replace('rowHeight=', 'showCaption="false" lockedPosition="true" rowHeight=');
  const back = new Workbook(readXlsx(zip(files)).data).sheets[0].slicers[0];
  assert.equal(back.buttonHeight, 24); assert.equal(back.showHeader, false); assert.equal(back.noMove, true);
});

test('순수 슬라이서 그룹은 표준 grpSp와 독립 필터 연결·소수 좌표·숨김·쌓기 순서를 보존', () => {
  const one = { ...sl, objectGroup: 'group1', z: 5, hidden: true, alt: '첫 슬라이서', macro: 'ShowRegion', columns: 2, buttonHeight: 26.125 };
  const two = { ...sl, id: 's2', objectGroup: 'group1', z: 4, caption: '두 번째', x: 430.375, y: 40.125, w: 211.25, h: 189.5 };
  const wb = book(one); wb.transact(() => wb.setSheetProp(0, 'slicers', [one, two]));
  const before = wb.serialize(), bytes = writeXlsx(wb), drawing = parseXml(textOf(unzip(bytes)['xl/drawings/drawing1.xml']));
  assert.equal(kids(drawing, 'twoCellAnchor').length, 1);
  const group = child(drawing.children[0], 'grpSp'); assert.ok(group);
  assert.equal(kids(group, 'AlternateContent').length, 2);
  assert.equal(descendants(group, 'slicer').length, 2);
  const back = new Workbook(readXlsx(bytes).data).sheets[0].slicers;
  assert.equal(back.length, 2); assert.ok(back[0].objectGroup); assert.equal(back[0].objectGroup, back[1].objectGroup);
  for (const original of [one, two]) {
    const got = back.find(s => s.caption === original.caption);
    for (const key of ['x', 'y', 'w', 'h']) close(got[key], original[key]);
    assert.deepEqual(got.source, original.source);
  }
  const first = back.find(s => s.caption === one.caption), second = back.find(s => s.caption === two.caption);
  assert.equal(first.hidden, true); assert.equal(first.alt, one.alt); assert.equal(first.macro, one.macro); assert.ok(first.z > second.z);
  assert.deepEqual(wb.serialize(), before);
});

test('외부 그룹의 chOff·chExt 좌표 변환과 부모 숨김은 flat 슬라이서에 반영', () => {
  const wb = book(sl), two = { ...sl, id: 's2', caption: '두 번째', x: 430, y: 40 };
  wb.transact(() => wb.setSheetProp(0, 'slicers', [{ ...sl, objectGroup: 'g' }, { ...two, objectGroup: 'g' }]));
  const files = unzip(writeXlsx(wb));
  // Excel의 그룹 예처럼 자식 원점과 chOff가 같은 nonzero 좌표계를 사용한다.
  let xml = textOf(files['xl/drawings/drawing1.xml']);
  xml = xml.replace('<a:chOff x="0" y="0"/>', '<a:chOff x="952500" y="476250"/>')
    .replace(/(<xdr:xfrm><a:off x=")(\d+)(" y=")(\d+)("\/>)/g, (_, a, x, b, y, c) => `${a}${Number(x) + 952500}${b}${Number(y) + 476250}${c}`)
    .replace(/(<xdr:nvGrpSpPr><xdr:cNvPr [^>]*?)(\/?>)/, '$1 hidden="1"$2');
  files['xl/drawings/drawing1.xml'] = xml;
  const back = new Workbook(readXlsx(zip(files)).data).sheets[0].slicers;
  close(back[0].x, sl.x); close(back[0].y, sl.y); close(back[1].x, two.x); close(back[1].y, two.y);
  assert.equal(back[0].hidden, true); assert.equal(back[1].hidden, true); assert.equal(back[0].objectGroup, back[1].objectGroup);
});

test('그룹 멤버가 한 개 남거나 해제하면 단독 네이티브 슬라이서로 저장', () => {
  for (const objectGroup of [undefined, 'staleGroup']) {
    const bytes = writeXlsx(book({ ...sl, objectGroup }));
    assert.doesNotMatch(textOf(unzip(bytes)['xl/drawings/drawing1.xml']), /<xdr:grpSp>/);
    assert.equal(new Workbook(readXlsx(bytes).data).sheets[0].slicers[0].objectGroup, undefined);
  }
});
