// PowerPoint 호환성: 미리 설정 도형 기하 · 조정 핸들 · 그룹 회전 · 글자 서식 왕복 · 레거시 그림 형식
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PRESETS } from '../src/prstgeom.js';
import { presetGeom, presetHandles, adjustFromPoint, evalFormula, placeInGroups, arcSvg, PRESET_NAMES } from '../src/geom.js';
import { shapePath, shapePaths, textRect, SHAPE_GALLERY } from '../src/shapes.js';
import { writePptx, readPptx } from '../src/pptx.js';
import { newPresentation, newSlide, newTextBox, newShape, para, run } from '../src/model.js';
import { wmfToSvg, isWmf } from '../src/wmf.js';
import { decodeTiff, isTiff } from '../src/tiff.js';
import { symbolText } from '../src/symfont.js';
import { PATTERN_NAMES, patternPath } from '../src/patterns.js';
import { isCfb } from '../src/cfb.js';

test('미리 설정 도형 187종: 모두 경로 · 글 영역이 계산됨 (NaN 없음)', () => {
  assert.equal(PRESET_NAMES.length, 187);
  for (const name of PRESET_NAMES) {
    const g = presetGeom(name, 200, 120);
    assert.ok(g.paths.length > 0, name);
    for (const p of g.paths) assert.ok(p.d.length > 0 && !/NaN|undefined|Infinity/.test(p.d), `${name}: ${p.d.slice(0, 80)}`);
    assert.ok(g.rect.every(Number.isFinite), name);
  }
  // 갤러리 도형은 모두 표준 도형 (텍스트 상자 제외)
  for (const [, list] of SHAPE_GALLERY) for (const [k] of list) if (k !== 'textBox') assert.ok(PRESETS[k], k);
});

test('안내선 식: ECMA-376 연산자', () => {
  const v = { w: 100, h: 50, a: 3, b: 4 };
  assert.equal(evalFormula('*/ w 3 4', v), 75);
  assert.equal(evalFormula('+- w h 10', v), 140);
  assert.equal(evalFormula('+/ w h 3', v), 50);
  assert.equal(evalFormula('?: -1 5 6', v), 6);
  assert.equal(evalFormula('mod a b 0', v), 5);
  assert.equal(evalFormula('pin 0 120 100', v), 100);
  assert.equal(Math.round(evalFormula('at2 0 10', v)), 5400000);
  assert.equal(Math.round(evalFormula('cos 10 5400000', v) * 1000) / 1000, 0);
  assert.equal(evalFormula('sqrt 16', v), 4);
});

test('도형 모양: 직사각형 · 원통(채우기 + 윤곽 경로) · 글 영역', () => {
  assert.equal(shapePath('rect', 10, 10).replace(/\s+/g, ' '), 'M0 0 L10 0 L10 10 L0 10 Z');
  const can = shapePaths('can', 100, 100);
  assert.ok(can.some((p) => p.fill === 'lighten') && can.some((p) => p.fill === 'none'));
  assert.deepEqual(textRect('rect', 10, 10), [0, 0, 10, 10]);
  const [l, t, r, b] = textRect('ellipse', 100, 100);
  assert.ok(l > 14 && l < 15 && r > 85 && Math.abs(t - l) < 1e-9 && Math.abs(b - r) < 1e-9);
});

test('arcTo: 보이는 각도 기준 (타원 0°→90° 는 오른쪽 끝에서 아래 끝으로)', () => {
  const a = arcSvg(200, 50, 100, 50, 0, 5400000);
  assert.ok(Math.abs(a.x - 100) < 1e-6 && Math.abs(a.y - 100) < 1e-6);
});

test('조정 핸들: 위치 계산과 끌기 (오른쪽 화살표 머리 · 말풍선 꼬리)', () => {
  const hs = presetHandles('rightArrow', 200, 100);
  assert.deepEqual(hs.map((p) => [Math.round(p.x), Math.round(p.y)]), [[0, 25], [150, 0]]);
  const a = adjustFromPoint('rightArrow', 200, 100, null, 1, 120, 0);
  assert.equal(a.adj2, 80000);
  const c = adjustFromPoint('wedgeRectCallout', 200, 100, null, 0, 250, 180);
  assert.deepEqual(c, { adj1: 75000, adj2: 130000 });
  // 범위 밖은 최대값으로 (rightArrow adj1 ≤ 100000)
  assert.equal(adjustFromPoint('rightArrow', 200, 100, null, 0, 0, -500).adj1, 100000);
});

test('그룹 회전 · 대칭: 자식 위치 · 회전 · 대칭 합성', () => {
  const g = { x: 0, y: 0, w: 200, h: 100, chx: 0, chy: 0, chw: 200, chh: 100, rot: 90 };
  // 그룹 왼쪽 위 칸 (0,0,100,50) → 그룹 중심 (100,50) 기준 시계 방향 90° 돌면 칸 중심 (50,25) 이 (125,0) 으로
  const r = placeInGroups(g, { x: 0, y: 0, w: 100, h: 50, rot: 0 });
  assert.equal(Math.round(r.x + r.w / 2), 125);
  assert.equal(Math.round(r.y + r.h / 2), 0);
  assert.equal(r.rot, 90);
  const fl = placeInGroups({ ...g, rot: 0, flipH: true }, { x: 0, y: 0, w: 100, h: 50, rot: 30 });
  assert.equal(Math.round(fl.x), 100);
  assert.equal(fl.rot, 330);
  assert.equal(fl.flipH, true);
  // 바깥 그룹이 크기를 반으로
  const nested = placeInGroups({ x: 10, y: 10, w: 100, h: 50, chx: 0, chy: 0, chw: 200, chh: 100, up: { x: 0, y: 0, w: 1000, h: 1000, chx: 0, chy: 0, chw: 1000, chh: 1000 } }, { x: 100, y: 50, w: 100, h: 50 });
  assert.deepEqual([nested.x, nested.y, nested.w, nested.h], [60, 35, 50, 25]);
});

test('pptx 왕복: 밑줄 종류 · 이중 취소선 · 그라데이션 글자 · 네온 · 그림자 · 필드 · 단 · WordArt 변환 · 조정값', () => {
  const p = newPresentation({ theme: 'office' });
  const s = newSlide(p, 'blank'); s.objects = []; p.slides = [s];
  const tb = newTextBox({ x: 10, y: 10, w: 600, h: 200 }, '');
  tb.text.cols = 2; tb.text.colSpc = 24; tb.text.warp = { prst: 'textArchUp' };
  tb.text.paras = [para('')];
  tb.text.paras[0].runs = [
    run('A', { u: 'dbl', uColor: '#FF0000' }), run('B', { u: 'wavyHeavy' }), run('C', { s: 'dbl' }),
    run('D', { fill: { type: 'gradient', stops: [[0, '#FF0000'], [1, '#0000FF']], angle: 0 }, color: '#FF0000' }),
    run('E', { glow: { rad: 6, color: '#00FF00' } }), run('F', { shadow: { blur: 4, dist: 5, dir: 135, color: '#000000:a50' } }),
    { t: '‹#›', size: 18, fld: { type: 'slidenum' } },
  ];
  s.objects.push(tb);
  const sh = newShape('wedgeRectCallout', { x: 700, y: 10, w: 200, h: 100 }); sh.adj = { adj1: 75000, adj2: 130000 };
  s.objects.push(sh);
  const back = readPptx(writePptx(p)).pres;
  const t2 = back.slides[0].objects.find((o) => o.text?.warp);
  assert.ok(t2, 'warp');
  assert.equal(t2.text.cols, 2);
  assert.equal(Math.round(t2.text.colSpc), 24);
  const rs = t2.text.paras[0].runs;
  assert.equal(rs[0].u, 'dbl'); assert.equal(rs[0].uColor, '#FF0000');
  assert.equal(rs[1].u, 'wavyHeavy');
  assert.equal(rs[2].s, 'dbl');
  assert.equal(rs[3].fill.type, 'gradient'); assert.equal(rs[3].fill.stops.length, 2);
  assert.equal(rs[4].glow.rad, 6);
  assert.equal(rs[5].shadow.dist, 5); assert.equal(rs[5].shadow.dir, 135);
  assert.equal(rs[6].fld.type, 'slidenum');
  const s2 = back.slides[0].objects.find((o) => o.shape === 'wedgeRectCallout');
  assert.deepEqual(s2.adj, { adj1: 75000, adj2: 130000 });
});

test('WMF: 사각형 · 기호 글꼴 글자 → SVG', () => {
  // 위치 지정 머리말 (22바이트) + 표준 머리말 (18바이트) + 레코드
  const recs = [];
  const rec = (fn, ...words) => { recs.push([3 + words.length, fn, words]); };
  rec(0x020b, 0, 0); // SetWindowOrg
  rec(0x020c, 100, 100); // SetWindowExt
  rec(0x041b, 90, 90, 10, 10); // Rectangle (b, r, t, l)
  rec(0, ); // EOF
  const body = [];
  for (const [size, fn, words] of recs) { body.push(size & 0xffff, size >>> 16, fn, ...words); }
  const words = [0xcdd7, 0x9ac6, 0, 0, 0, 100, 100, 1440, 0, 0, 0, /* std */ 1, 9, 0x300, 0, 0, 0, 0, 0, 0, ...body];
  const bytes = new Uint8Array(words.length * 2);
  words.forEach((w, i) => { bytes[i * 2] = w & 255; bytes[i * 2 + 1] = (w >> 8) & 255; });
  assert.ok(isWmf(bytes));
  const svg = wmfToSvg(bytes);
  assert.match(svg, /<svg[^>]*viewBox/);
  assert.match(svg, /<(rect|path)/);
});

test('TIFF: 무압축 RGB 2×1', () => {
  const b = new Uint8Array(8 + 2 + 12 * 8 + 4 + 6);
  const dv = new DataView(b.buffer);
  b.set([0x49, 0x49, 42, 0]); dv.setUint32(4, 8, true);
  dv.setUint16(8, 8, true);
  const dataOff = 8 + 2 + 12 * 8 + 4;
  const tags = [[256, 3, 1, 2], [257, 3, 1, 1], [258, 3, 1, 8], [259, 3, 1, 1], [262, 3, 1, 2], [273, 4, 1, dataOff], [277, 3, 1, 3], [279, 4, 1, 6]];
  tags.forEach(([t, ty, n, v], i) => { const o = 10 + i * 12; dv.setUint16(o, t, true); dv.setUint16(o + 2, ty, true); dv.setUint32(o + 4, n, true); if (ty === 3) dv.setUint16(o + 8, v, true); else dv.setUint32(o + 8, v, true); });
  b.set([255, 0, 0, 0, 0, 255], dataOff);
  assert.ok(isTiff(b));
  const img = decodeTiff(b);
  assert.equal(img.w, 2); assert.equal(img.h, 1);
  assert.deepEqual([...img.rgba], [255, 0, 0, 255, 0, 0, 255, 255]);
});

test('기호 글꼴 (Symbol · Wingdings) → 유니코드, 무늬 채우기 54종, OLE 서명', () => {
  assert.equal(symbolText('Symbol', 'abt'), 'αβτ');
  assert.equal(symbolText('Symbol', ''), 'α');
  assert.equal(PATTERN_NAMES.length, 54);
  assert.ok(PATTERN_NAMES.every((n) => patternPath(n).startsWith('M')));
  assert.ok(isCfb(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Array(600).fill(0)])));
});
