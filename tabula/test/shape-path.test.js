import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shapeFromPoints, normalizedPathBounds, validShapePath, shapePathParts, customGeometryXml, readCustomGeometry } from '../src/shape-path.js';
import { newShape, shapeSvg, isShapeLine, shapeLineEnds, SHAPE_PATTERN_PRESETS } from '../src/shapes.js';
import { parseXml } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';

const close = (a, b, epsilon = 1e-6) => assert.ok(Math.abs(a - b) <= epsilon, `${a} != ${b}`);
const make = (kind, closed = false) => newShape(kind, shapeFromPoints(kind, [[20, 40], [90, 10], [160, 90], [220, 35]], { closed }));
const save = shapes => {
  const wb = new Workbook(); wb.transact(() => wb.setSheetProp(0, 'shapes', shapes));
  const bytes = writeXlsx(wb), files = unzip(bytes), xml = textOf(files['xl/drawings/drawing1.xml']);
  return { wb, files, xml, back: readXlsx(bytes).data.sheets[0].shapes };
};

test('자유곡선·자유형은 모든 점을 정규화하고 닫힌 도형만 채운다', () => {
  for (const [kind, closed] of [['scribble', false], ['freeform', false], ['freeform', true]]) {
    const sh = make(kind, closed), p = sh.path.paths[0];
    assert.equal(validShapePath(sh.path), true); assert.equal(isShapeLine(sh), !closed);
    assert.equal(p.fill, closed); assert.equal(p.commands.at(-1)[0] === 'Z', closed);
    assert.deepEqual([sh.x, sh.y, sh.w, sh.h], [20, 10, 200, 80]);
    const absolute = p.commands.filter(c => c.length > 1).map(c => [sh.x + c[1] * sh.w, sh.y + c[2] * sh.h]);
    assert.deepEqual(absolute, [[20, 40], [90, 10], [160, 90], [220, 35]]);
  }
});

test('곡선 경계는 제어점 대신 실제 cubic/quadratic 극값을 사용한다', () => {
  const c = { paths: [{ commands: [['M', 0, 0], ['C', 0, 2, 1, 2, 1, 0]], fill: false }] };
  assert.deepEqual(normalizedPathBounds(c), { minX: 0, minY: 0, maxX: 1, maxY: 1.5 });
  const q = { paths: [{ commands: [['M', 0, 0], ['Q', .5, 2, 1, 0]] }] };
  assert.deepEqual(normalizedPathBounds(q), { minX: 0, minY: 0, maxX: 1, maxY: 1 });
  const sh = make('curve'), b = normalizedPathBounds(sh.path);
  close(b.minX, 0); close(b.minY, 0); close(b.maxX, 1); close(b.maxY, 1);
  const p = sh.path.paths[0]; assert.equal(p.commands[1][0], 'C');
  for (const [i, point] of [[1, [90, 10]], [2, [160, 90]], [3, [220, 35]]]) {
    close(sh.x + p.commands[i][5] * sh.w, point[0]); close(sh.y + p.commands[i][6] * sh.h, point[1]);
  }
});

test('점 입력 한도·비수치 경로를 명시적으로 거절하고 원본 배열을 바꾸지 않는다', () => {
  const pts = [[1, 2], [1, 2], [5, 7]], before = structuredClone(pts);
  assert.equal(shapeFromPoints('scribble', pts).path.paths[0].commands.length, 2); assert.deepEqual(pts, before);
  assert.throws(() => shapeFromPoints('curve', [[0, 0], [NaN, 1]]), /좌표/);
  assert.throws(() => shapeFromPoints('freeform', [[1, 1], [1, 1]]), /서로 다른/);
  assert.throws(() => shapeFromPoints('scribble', Array.from({ length: 10000 }, (_, i) => [i, i])), /9,998/);
  assert.equal(validShapePath({ paths: [{ commands: [['M', 0, 0], ['L', '" onload="x', 1]] }] }), false);
  assert.equal(normalizedPathBounds(null), null);
});

test('자유 경로 SVG는 resize/flip과 점편집 좌표가 일치하고 선에도 효과를 적용한다', () => {
  const sh = { ...make('curve'), w: 300, h: 140, flip: true, flipV: true, shadow: true, glow: { size: 3 }, soft: 1, lineCap: 'rnd', lineJoin: 'bevel', dash: 'lgDashDotDot', arrow: 'both' };
  const original = structuredClone(sh), svg = shapeSvg(sh);
  assert.match(svg, /translate\(0,0\) translate\(300,140\) scale\(-1,-1\)/);
  assert.ok(svg.includes(shapePathParts(sh.path, 300, 140)[0].d));
  assert.match(svg, /stroke-linecap="round"/); assert.match(svg, /stroke-linejoin="bevel"/);
  assert.match(svg, /feDropShadow/); assert.match(svg, /feMorphology/); assert.match(svg, /feGaussianBlur/);
  assert.match(svg, /marker-start=/); assert.match(svg, /marker-end=/); assert.deepEqual(sh, original);
  assert.doesNotMatch(shapeSvg({ ...sh, stroke: null }), /stroke="#000/);
});

test('양끝 화살표 종류·너비·길이, 이중·삼중선과 패턴은 실제 SVG 정의를 만든다', () => {
  const sh = { ...make('scribble'), strokeWidth: 6, headEnd: { type: 'oval', w: 'sm', len: 'lg' }, tailEnd: { type: 'arrow', w: 'lg', len: 'sm' } };
  assert.deepEqual(shapeLineEnds(sh).headEnd, sh.headEnd);
  const svg = shapeSvg(sh); assert.match(svg, /<ellipse/); assert.match(svg, /markerWidth="7" markerHeight="3"/);
  assert.match(svg, /markerWidth="3" markerHeight="7"/);
  for (const compound of ['dbl', 'tri']) {
    assert.match(shapeSvg({ ...sh, compound }), /<mask/);
    // A zero-height/width line has no objectBoundingBox area: use sheet units.
    for (const box of [{ w: 200, h: 0 }, { w: 0, h: 200 }]) {
      const line = shapeSvg({ ...newShape('line', box), compound, strokeWidth: 6, shadow: true });
      assert.match(line, /maskUnits="userSpaceOnUse"/); assert.match(line, /filterUnits="userSpaceOnUse"/);
      assert.ok(line.includes(`M0,0 L${box.w},${box.h}`));
    }
  }
  for (const preset of SHAPE_PATTERN_PRESETS) {
    const text = shapeSvg({ ...make('freeform', true), pattern: { preset, fg: '#112233', bg: '#aabbcc' } });
    assert.match(text, /<pattern/); assert.match(text, /fill="url\(#p/);
  }
});

test('표준 custGeom M/L/C/Q/Z·복수 경로를 확장 없이 저장하고 다시 읽는다', () => {
  const sh = make('freeform', true);
  sh.path.paths.push({ commands: [['M', .1, .2], ['Q', .7, -.2, .9, .8]], fill: false, stroke: false });
  const source = [make('scribble'), make('curve'), sh].map(s => ({ ...s, rot: 35, flip: true, flipV: true }));
  const { xml, back } = save(source);
  assert.equal((xml.match(/<a:custGeom>/g) ?? []).length, 3); assert.doesNotMatch(xml, /wx:|wixel/i);
  assert.match(xml, /<a:cubicBezTo>/); assert.match(xml, /<a:quadBezTo>/); assert.match(xml, /<a:close\/>/);
  for (let i = 0; i < source.length; i++) {
    assert.equal(back[i].rot, 35); assert.equal(back[i].flip, true); assert.equal(back[i].flipV, true);
    const a = source[i].path.paths, b = back[i].path.paths; assert.equal(a.length, b.length);
    a.forEach((p, k) => { assert.equal(p.fill, b[k].fill); assert.equal(p.stroke, b[k].stroke); p.commands.forEach((c, j) => { assert.equal(c[0], b[k].commands[j][0]); c.slice(1).forEach((n, x) => close(n, b[k].commands[j][x + 1])); }); });
  }
});

test('선 끝·대시·cap/join/compound·패턴·텍스트/비율잠금을 표준 XML에 보존한다', () => {
  const shape = { ...make('freeform', true), headEnd: { type: 'stealth', w: 'lg', len: 'sm' }, tailEnd: { type: 'diamond', w: 'sm', len: 'lg' }, dash: 'lgDashDotDot', lineCap: 'sq', lineJoin: 'miter', compound: 'tri', pattern: { preset: 'diagCross', fg: '#112233', bg: '#abcdef' }, fillOpacity: .35, text: '합성', textFit: 'shrink', textRot: 270, lockAspect: true, alt: '도형 설명' };
  const { xml, back: [back] } = save([shape]);
  for (const k of ['headEnd', 'tailEnd', 'dash', 'lineCap', 'lineJoin', 'compound', 'pattern', 'fillOpacity', 'textFit', 'textRot', 'lockAspect', 'alt']) assert.deepEqual(back[k], shape[k], k);
  assert.match(xml, /<a:pattFill/); assert.match(xml, /<a:spLocks noChangeAspect="1"/); assert.match(xml, /<a:normAutofit\/>/);
  const start = { ...newShape('line', { x: 10, y: 10, w: 120, h: 40 }), arrow: 'start' };
  const result = save([start]).back[0]; assert.equal(result.arrow, 'start'); assert.equal(!!result.flip, false); assert.equal(!!result.flipV, false);
});

test('가져온 미지원 arc/guide 경로는 경고하고 원래 표준 geometry를 재저장한다', () => {
  const { files, xml } = save([make('freeform', true)]);
  const unsupported = '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="10800000" swAng="5400000"/></a:path></a:pathLst></a:custGeom>';
  files['xl/drawings/drawing1.xml'] = new TextEncoder().encode(xml.replace(/<a:custGeom>[\s\S]*?<\/a:custGeom>/, unsupported));
  const read = readXlsx(zip(files)); assert.ok(read.warnings.some(w => /경로 수식/.test(w)));
  const s = read.data.sheets[0].shapes[0]; assert.ok(s.customGeometry); assert.equal(s.path, undefined);
  assert.match(save([s]).xml, /<a:arcTo wR="50" hR="50" stAng="10800000" swAng="5400000">/);
});

test('작은 표준 path의 guide 좌표를 읽고 Workbook 실행 취소/직렬화가 자유 경로를 보존한다', () => {
  const xml = '<a:custGeom><a:gdLst><a:gd name="p" fmla="val 25"/></a:gdLst><a:pathLst><a:path w="100" h="200" fill="none"><a:moveTo><a:pt x="p" y="0"/></a:moveTo><a:lnTo><a:pt x="w" y="h"/></a:lnTo></a:path></a:pathLst></a:custGeom>';
  const path = readCustomGeometry(parseXml(xml), 100, 100); assert.deepEqual(path.paths[0].commands, [['M', .25, 0], ['L', 1, 1]]);
  assert.ok(customGeometryXml(path).includes('x="250000"'));
  const sh = make('curve'), wb = save([sh]).wb, before = wb.serialize();
  wb.transact(() => wb.setSheetProp(0, 'shapes', [{ ...sh, w: sh.w * 2 }])); wb.undo(); assert.deepEqual(wb.serialize(), before);
  wb.redo(); const restored = new Workbook(wb.serialize()); assert.deepEqual(restored.sheets[0].shapes[0].path, sh.path); assert.equal(restored.sheets[0].shapes[0].w, sh.w * 2);
});
