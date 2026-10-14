// 1.9 호환성: .ppt 애니메이션 (PowerPoint 2002+ 시간 노드 · 97 빌드 정보) · 번호 글머리 (PP9) · WordArt 변형 40종 · 글꼴 파일 · 무료 글꼴 이름
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { timeNodeAnims, legacyAnims, parseStyle9 } from '../src/ppt.js';
import { WARP_PATHS, WARP_ENVELOPES } from '../src/render.js';
import { fontKind, sfntInfo, makeEot, decodeEmbeddedFont, embeddable } from '../src/fonts.js';
import { webFont, WEB_FONTS } from '../src/webfonts.js';

// 레코드 만들기: [type, inst, 내용(Uint8Array) | 자식 레코드 배열]
function rec(type, inst, body) {
  const kids = Array.isArray(body);
  const data = kids ? concat(body) : body;
  const h = new Uint8Array(8);
  const dv = new DataView(h.buffer);
  dv.setUint16(0, (inst << 4) | (kids ? 0xf : 0), true); dv.setUint16(2, type, true); dv.setUint32(4, data.length, true);
  return concat([h, data]);
}
const concat = (parts) => { const n = parts.reduce((s, p) => s + p.length, 0); const o = new Uint8Array(n); let p = 0; for (const x of parts) { o.set(x, p); p += x.length; } return o; };
const i32s = (...v) => { const b = new Uint8Array(v.length * 4); const dv = new DataView(b.buffer); v.forEach((x, i) => dv.setInt32(i * 4, x, true)); return b; };
const prop = (id, v) => rec(0xf142, id, concat([new Uint8Array([1]), i32s(v)]));
const node = (dur, ...rest) => rec(0xf144, 1, [rec(0xf127, 0, i32s(0, 0, 0, 0, 0, 0, dur, 0)), ...rest]);
const props = (m) => rec(0xf13d, 0, Object.entries(m).map(([k, v]) => prop(+k, v)));
const begin = (delay, target) => rec(0xf125, 1, [rec(0xf128, 0, i32s(0, target ? 5 : 0, 0, delay)), ...(target ? [rec(0xf13c, 0, [rec(0x2afb, 0, i32s(0, 1, target, -1, -1))])] : [])]);
const target = (spid, mode = 0) => rec(0xf12b, 0, [rec(0xf12a, 0, [rec(0xf13c, 0, [rec(0x2afb, 0, i32s(mode, 1, spid, -1, -1))])])]);
const effect = (cls, pid, sub, nodeType, spid, dur = 500, delay = 0) => node(0, props({ 20: nodeType, 9: pid, 10: sub, 11: cls }), begin(delay), node(dur, target(spid)));

test('.ppt 시간 노드 (PowerPoint 2002+): 효과 · 방향 · 시작 · 지연 · 길이 · 트리거', () => {
  const main = node(0, props({ 20: 4 }), effect(1, 2, 8, 1, 1025, 750), effect(3, 6, 0, 2, 1026, 2000, 500), effect(2, 10, 0, 3, 1025));
  const inter = node(0, props({ 20: 5 }), begin(0, 1026), effect(1, 10, 0, 1, 1025));
  const root = node(0, props({ 20: 9 }), node(0, main, inter));
  const slide = { kids: [{ type: 0x138b, data: root }] };
  const spMap = new Map([[1025, ['a']], [1026, ['b']]]);
  const anims = timeNodeAnims(slide, spMap, [{ id: 'a' }, { id: 'b' }]);
  assert.deepEqual(anims.map((a) => [a.obj, a.cls, a.effect, a.start, a.dir, a.trigger ?? null]), [
    ['a', 'entr', 'fly', 'click', 'l', null], ['b', 'emph', 'growShrink', 'with', 'b', null], ['a', 'exit', 'fade', 'after', 'b', null], ['a', 'entr', 'fade', 'click', 'b', 'b'],
  ]);
  assert.equal(anims[0].dur, 0.75);
  assert.equal(anims[1].delay, 0.5);
  // 시간 노드가 없으면 null (97 빌드 정보로 넘어감)
  assert.equal(timeNodeAnims({ kids: [] }, spMap, []), null);
});

test('.ppt 빌드 정보 (PowerPoint 97): 순서 · 날아오기 방향 · 자동 · 단락별', () => {
  const atom = (order, eff, dir, auto, build = 1) => { const d = new Uint8Array(28); const dv = new DataView(d.buffer); dv.setUint16(4, auto ? 4 : 0, true); dv.setInt32(12, 1000, true); dv.setUint16(16, order, true); d[20] = build; d[21] = eff; d[22] = dir; return d; };
  const anims = legacyAnims([{ obj: 'b', d: atom(2, 12, 2, true) }, { obj: 'a', d: atom(1, 2, 0, false, 2) }, { obj: 'c', d: atom(0, 6, 0, false) }], [{ id: 'a' }, { id: 'b' }, { id: 'c' }]);
  assert.deepEqual(anims.map((a) => [a.obj, a.effect, a.start, a.dir, !!a.byPara]), [['a', 'blinds', 'click', 'l', true], ['b', 'fly', 'after', 'r', false]]);
  assert.equal(anims[1].delay, 1);
});

test('.ppt 번호 글머리 (StyleTextProp9Atom): 번호 종류 · 시작 번호', () => {
  const d = new Uint8Array(20); const dv = new DataView(d.buffer);
  dv.setUint32(0, 0x3000000, true); dv.setUint16(4, 1, true); dv.setUint16(6, 3, true); dv.setInt16(8, 5, true); // pf9: 자동 번호 · 종류 3 · 시작 5
  dv.setUint32(10, 0, true); dv.setUint32(14, 0x2, true); dv.setUint16(18, 0x412, true); // cf9 없음 · si: lang
  assert.deepEqual(parseStyle9(d), [{ auto: true, scheme: 3, start: 5 }]);
});

test('WordArt 변형 40종: 경로 또는 봉투로 모두 그림', () => {
  const all = ['textNoShape', 'textPlain', 'textStop', 'textTriangle', 'textTriangleInverted', 'textChevron', 'textChevronInverted', 'textRingInside', 'textRingOutside', 'textArchUp', 'textArchDown', 'textCircle', 'textButton', 'textArchUpPour', 'textArchDownPour', 'textCirclePour', 'textButtonPour', 'textCurveUp', 'textCurveDown', 'textCanUp', 'textCanDown', 'textWave1', 'textWave2', 'textDoubleWave1', 'textWave4', 'textInflate', 'textDeflate', 'textInflateBottom', 'textDeflateBottom', 'textInflateTop', 'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate', 'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown', 'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown'];
  const missing = all.filter((k) => k !== 'textNoShape' && !WARP_PATHS[k] && !WARP_ENVELOPES[k]);
  assert.deepEqual(missing, []);
});

// 최소 TrueType: name (1 = 글꼴 이름, 2 = 스타일) + OS/2 (굵기 · fsType)
function tinyTtf(name, fsType = 0) {
  const be16 = (v) => [v >> 8 & 255, v & 255];
  const s = [...name].flatMap((c) => be16(c.charCodeAt(0)));
  const st = [...'Bold'].flatMap((c) => be16(c.charCodeAt(0)));
  const nameT = [...be16(0), ...be16(2), ...be16(6 + 24), ...be16(3), ...be16(1), ...be16(0x412), ...be16(1), ...be16(s.length), ...be16(0), ...be16(3), ...be16(1), ...be16(0x409), ...be16(2), ...be16(st.length), ...be16(s.length), ...s, ...st];
  const os2 = new Array(78).fill(0); os2[4] = 700 >> 8; os2[5] = 700 & 255; os2[8] = fsType >> 8; os2[9] = fsType & 255;
  const tables = [['OS/2', os2], ['name', nameT]];
  let off = 12 + tables.length * 16;
  const head = [0, 1, 0, 0, ...be16(tables.length), 0, 0, 0, 0, 0, 0];
  const dirs = []; const body = [];
  for (const [tag, t] of tables) { dirs.push(...[...tag].map((c) => c.charCodeAt(0)), 0, 0, 0, 0, off >>> 24, off >> 16 & 255, off >> 8 & 255, off & 255, 0, 0, t.length >> 8, t.length & 255); while (t.length % 4) t.push(0); body.push(...t); off += t.length; }
  return new Uint8Array([...head, ...dirs, ...body]);
}

test('글꼴 파일: 종류 · 이름 · 굵기 · 포함 허용 · EOT 왕복', () => {
  const ttf = tinyTtf('위포인트체');
  assert.equal(fontKind(ttf), 'ttf');
  const info = sfntInfo(ttf);
  assert.equal(info.family, '위포인트체');
  assert.equal(info.style, 'Bold');
  assert.equal(info.weight, 700);
  assert.ok(embeddable(info));
  assert.ok(!embeddable(sfntInfo(tinyTtf('잠금체', 0x0002))));
  const eot = makeEot(ttf, info);
  const back = decodeEmbeddedFont(eot);
  assert.deepEqual([...(back.data ?? back)].slice(0, ttf.length), [...ttf]);
});

test('무료 한글 글꼴: 목록 · 별칭 · 띄어쓰기 무시', () => {
  assert.ok(WEB_FONTS.length > 500);
  assert.ok(webFont('Pretendard'));
  assert.ok(webFont('배달의민족 주아'));
  assert.ok(webFont('나눔스퀘어 네오'));
  assert.ok(webFont('Nanum Gothic'));
  assert.equal(webFont('Nanum Gothic').family, webFont('NanumGothic')?.family);
});
