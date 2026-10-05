import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { shapeTextHtml } from '../src/view.js';
import { shapeTextFormatPatch } from '../src/shape-text-format.js';

const shape = () => ({ id: 'rich-shape', kind: 'rect', x: 20, y: 40, w: 320, h: 220, text: '강조보통\n\n마지막',
  bold: true, italic: true, underline: true, strike: true, size: 11, font: 'Arial', color: '#123456',
  paras: [{ align: 'center', runs: [{ t: '강조', sz: 20, color: '#ff0000' }, { t: '보통', b: false, i: false, u: false, s: false, font: '맑은 고딕', sz: 12 }] },
    { align: 'right', runs: [], sz: 25 }, { align: 'left', runs: [{ t: '마지막', i: false }] }] });
const bytes = sh => writeXlsx(new Workbook({ sheets: [{ name: '텍스트', cells: {}, shapes: [sh] }] }));
const back = sh => readXlsx(bytes(sh)).data.sheets[0].shapes[0];

test('전체 글꼴·색 변경은 다른 부분 서식과 문단 구조를 보존하고 원본을 바꾸지 않는다', () => {
  const s = shape(), before = structuredClone(s);
  const patch = shapeTextFormatPatch(s, { font: 'Georgia', color: '#abcdef' });
  assert.equal(patch.font, 'Georgia'); assert.equal(patch.color, '#abcdef');
  for (const p of patch.paras) for (const r of p.runs) { assert.equal(r.font, 'Georgia'); assert.equal(r.color, '#abcdef'); }
  assert.equal(patch.paras[0].runs[0].sz, 20); assert.equal(patch.paras[0].runs[1].sz, 12);
  assert.equal(patch.paras[0].runs[1].b, false); assert.equal(patch.paras[0].runs[1].u, false);
  assert.deepEqual(patch.paras.map(p => p.align), ['center', 'right', 'left']);
  assert.deepEqual(s, before);
});

test('굵게 해제·크기·맞춤은 대상 속성만 모든 조각과 빈 문단에 적용한다', () => {
  const s = shape();
  const patch = shapeTextFormatPatch(s, { bold: undefined, size: 16, align: 'justify' });
  assert.equal(patch.paras[0].runs[0].b, false); assert.equal(patch.paras[0].runs[1].b, false);
  for (const p of patch.paras) { assert.equal(p.align, 'justify'); for (const r of p.runs) assert.equal(r.sz, 16); }
  assert.equal(patch.paras[1].sz, 16);
  assert.equal(patch.paras[0].runs[0].color, '#ff0000'); assert.equal(patch.paras[0].runs[1].i, false);
  assert.equal(patch.paras[0].runs[1].font, '맑은 고딕');
});

test('같은 텍스트 change는 부분 서식을 없애지 않고 실제 본문 교체만 평문으로 바꾼다', () => {
  const s = shape();
  assert.deepEqual(shapeTextFormatPatch(s, { text: s.text }), {});
  assert.equal(shapeTextFormatPatch(s, { text: '새 본문' }).paras, undefined);
  assert.equal(shapeTextFormatPatch(s, { text: s.text, align: 'left' }).paras.length, 3);
  const noPlain = { paras: [{ runs: [{ t: '첫째' }] }, { runs: [{ t: '둘째' }] }] };
  assert.deepEqual(shapeTextFormatPatch(noPlain, { text: '첫째\n둘째' }), {});
});

test('rich 렌더링은 전역 장식 대신 각 조각의 명시 false와 상속 값을 적용한다', () => {
  const html = shapeTextHtml(shape());
  const outer = html.slice(0, html.indexOf('><div class="sh-text-content">'));
  assert.doesNotMatch(outer, /font-weight:700|font-style:italic|text-decoration:/);
  assert.match(html, /font-weight:700;font-style:italic;text-decoration:underline line-through[^>]*>강조/);
  assert.match(html, /font-weight:400;font-style:normal;text-decoration:none[^>]*>보통/);
  assert.match(html, /font-weight:700;font-style:normal;text-decoration:underline line-through[^>]*>마지막/);
  const plain = shapeTextHtml({ text: '평문', bold: true, italic: true, underline: true, strike: true });
  assert.match(plain, /font-weight:700/); assert.match(plain, /text-decoration:underline line-through/);
});

test('혼합 강조·글꼴·크기·색·빈 문단 맞춤은 표준 DrawingML만으로 왕복한다', () => {
  const files = unzip(bytes(shape())), xml = textOf(files['xl/drawings/drawing1.xml']);
  assert.doesNotMatch(xml, /wx:|tabula|wixel/i);
  assert.match(xml, /b="0" i="0" u="none" strike="noStrike"/);
  assert.match(xml, /xml:space="preserve"/);
  const first = readXlsx(zip(files)).data.sheets[0].shapes[0], second = back(first);
  assert.equal(first.text, shape().text); assert.deepEqual(second.paras, first.paras);
  const [strong, normal] = first.paras[0].runs;
  assert.deepEqual([strong.b, strong.i, strong.u, strong.s, strong.sz, strong.font, strong.color], [true, true, true, true, 20, 'Arial', '#ff0000']);
  assert.deepEqual([normal.b, normal.i, normal.u, normal.s, normal.sz, normal.font, normal.color], [false, false, false, false, 12, '맑은 고딕', '#123456']);
  assert.deepEqual(first.paras.map(p => p.align), ['center', 'right', 'left']); assert.equal(first.paras[1].sz, 25);
});

test('run 안의 연속 줄바꿈은 a:br로 저장하고 편집용 평문 및 재저장에 유지한다', () => {
  const sh = { ...shape(), text: '첫째\n\n둘째', paras: [{ align: 'left', runs: [{ t: '첫째\n\n둘째', b: false, i: true, sz: 14, color: '#abcdef' }] }] };
  const files = unzip(bytes(sh)); const xml = textOf(files['xl/drawings/drawing1.xml']);
  assert.equal((xml.match(/<a:br>/g) ?? []).length, 2);
  const first = readXlsx(zip(files)).data.sheets[0].shapes[0];
  assert.equal(first.text, sh.text);
  assert.deepEqual(first.paras[0].runs.map(r => r.t), ['첫째', '\n', '\n', '둘째']);
  for (const r of first.paras[0].runs) { assert.equal(r.b, false); assert.equal(r.i, true); assert.equal(r.sz, 14); }
  const second = back(first); assert.equal(second.text, sh.text); assert.deepEqual(second.paras, first.paras);
});

test('Excel의 문단 기본 서식과 목록 단계 서식은 조각별 명시 해제보다 우선하지 않는다', () => {
  const files = unzip(bytes({ ...shape(), paras: undefined }));
  let xml = textOf(files['xl/drawings/drawing1.xml']);
  const native = `<xdr:txBody><a:bodyPr/><a:lstStyle><a:defPPr><a:defRPr sz="1100"><a:latin typeface="Arial"/></a:defRPr></a:defPPr><a:lvl1pPr algn="ctr"><a:defRPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:defRPr></a:lvl1pPr><a:lvl2pPr algn="r"><a:defRPr sz="2400"/></a:lvl2pPr></a:lstStyle><a:p><a:pPr><a:defRPr sz="2000"><a:ea typeface="맑은 고딕"/></a:defRPr></a:pPr><a:r><a:t>상속</a:t></a:r><a:r><a:rPr b="0" i="0" u="none" strike="noStrike" sz="1300"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>해제</a:t></a:r><a:br/><a:r><a:t>다음</a:t></a:r></a:p><a:p><a:pPr lvl="1"/><a:r><a:t>단계</a:t></a:r></a:p></xdr:txBody>`;
  files['xl/drawings/drawing1.xml'] = new TextEncoder().encode(xml.replace(/<xdr:txBody>[\s\S]*?<\/xdr:txBody>/, native));
  const sh = readXlsx(zip(files)).data.sheets[0].shapes[0];
  assert.equal(sh.text, '상속해제\n다음\n단계'); assert.equal(sh.align, 'center');
  const [inherited, disabled, br] = sh.paras[0].runs;
  assert.deepEqual([inherited.b, inherited.i, inherited.u, inherited.s, inherited.sz, inherited.font, inherited.color], [true, true, true, true, 20, '맑은 고딕', '#123456']);
  assert.deepEqual([disabled.b, disabled.i, disabled.u, disabled.s, disabled.sz, disabled.font, disabled.color], [false, false, false, false, 13, 'Georgia', '#ff0000']);
  assert.equal(br.t, '\n'); assert.equal(br.sz, 20);
  assert.equal(sh.paras[1].align, 'right'); assert.equal(sh.paras[1].runs[0].sz, 24);
  const twice = back(sh); assert.equal(twice.text, sh.text);
  const displayed = ps => ps.map(p => ({ ...p, runs: p.runs.map(r => ({ ...r, color: r.color ?? '#000000' })) }));
  assert.deepEqual(displayed(twice.paras), displayed(sh.paras));
});

test('기본 가로 맞춤은 도형 가운데·텍스트 상자 왼쪽으로 화면과 XLSX가 같다', () => {
  for (const kind of ['rect', 'textbox']) {
    const sh = back({ id: 'align', kind, x: 0, y: 0, w: 100, h: 60, text: '본문' });
    assert.equal(sh.align, kind === 'rect' ? 'center' : 'left');
  }
});
