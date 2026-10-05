import test, { after } from 'node:test';
import assert from 'node:assert/strict';

// 네트워크·실제 DOM 없이 폰트가 늦게 준비되는 캔버스 측정을 재현한다.
const previousDocument = globalThis.document;
let loaded = false, created = 0, measured = 0;
const canvasContext = {
  font: '',
  measureText() {
    measured++;
    const chosen = /Font Metrics Fixture|Malgun Gothic|맑은 고딕/.test(this.font);
    return { width: chosen && loaded ? 42 : 10, fontBoundingBoxAscent: 80, fontBoundingBoxDescent: 20,
      actualBoundingBoxAscent: loaded ? 50 : 70, actualBoundingBoxDescent: 10 };
  },
};
globalThis.document = { createElement(tag) { created++; assert.equal(tag, 'canvas'); return { getContext: () => canvasContext }; } };
const { fontStack, measureText, fontMissing, glyphShift, clearFontMetrics, GridView, shapeTextHtml } = await import('../src/view.js');
after(() => { if (previousDocument === undefined) delete globalThis.document; else globalThis.document = previousDocument; });

test('폰트 스택 생성은 순수하며 기존 한글·영문 별칭을 보존한다', () => {
  const before = created;
  assert.match(fontStack('맑은 고딕'), /'맑은 고딕'.*'Malgun Gothic'/);
  for (let i = 0; i < 2000; i++) fontStack('목록 미리보기 ' + i);
  assert.equal(created, before, '목록 생성은 stylesheet나canvas를 추가하지 않는다');
  assert.doesNotMatch(fontStack('잘못된"<글꼴>\n'), /["<>\n]/);
});

test('늦게 준비된 폰트는 폭·미설치·세로 정렬 캐시를 함께 재측정한다', () => {
  loaded = false; clearFontMetrics();
  const font = 'Font Metrics Fixture', style = { font, size: 11 };
  const before = { width: measureText('예제', style), missing: fontMissing(font), shift: glyphShift(fontStack(font)) };
  assert.equal(before.width, 10); assert.equal(before.missing, true);
  loaded = true;
  const cachedCount = measured;
  assert.equal(measureText('예제', style), before.width);
  assert.equal(fontMissing(font), true); assert.equal(glyphShift(fontStack(font)), before.shift);
  assert.equal(measured, cachedCount, '준비 완료 전에는 같은 계산을 반복하지 않는다');
  clearFontMetrics();
  assert.equal(measureText('예제', style), 42); assert.equal(fontMissing(font), false);
  assert.notEqual(glyphShift(fontStack(font)), before.shift);
});

test('미설치 판정도 CSS에서 사용하는 한글·영문 후보를 검사한다', () => {
  loaded = true; clearFontMetrics();
  assert.equal(fontMissing('맑은 고딕'), false);
});

test('폰트 준비는 객체 측정 캐시만 무효화하고 문서·원본 개체를 바꾸지 않는다', () => {
  const object = { id: '텍스트', kind: 'textbox', text: '유지', font: 'Font Metrics Fixture', x: 0, y: 0, w: 90, h: 50 };
  const sheet = { charts: [], shapes: [object] }, book = { version: 4, sheets: [sheet] };
  const view = { host: { state: () => ({ wb: book, si: 0 }) }, z: 1 };
  const pane = { win: null };
  GridView.prototype.renderObjects.call(view, pane);
  const firstAppearance = view._objectRenderState.appearance;
  view._objectRenderCache.set(object, { html: '이전 글꼴 크기로 만든 렌더' });
  const source = JSON.stringify(book);
  clearFontMetrics(); GridView.prototype.renderObjects.call(view, pane);
  assert.notEqual(view._objectRenderState.appearance, firstAppearance);
  assert.equal(view._objectRenderCache.size, 0); assert.equal(JSON.stringify(book), source);
  assert.equal(book.version, 4); assert.equal(sheet.shapes[0], object);
});

test('도형의 원본 글꼴 이름과 부분 서식을 유지하며 CSS 별칭을 사용한다', () => {
  const shape = { kind: 'textbox', font: '맑은 고딕', bold: true, paras: [{ runs: [
    { t: '본문', font: 'Malgun Gothic', b: false }, { t: '강조', font: 'Font Metrics Fixture', i: true },
  ] }] };
  const before = JSON.stringify(shape), html = shapeTextHtml(shape);
  assert.match(html, /font-family:'맑은 고딕', 'Malgun Gothic'/);
  assert.match(html, /font-weight:400/); assert.match(html, /font-style:italic/);
  assert.equal(JSON.stringify(shape), before);
});
