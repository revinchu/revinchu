// ECMA-376 미리 설정 도형 정의 (presetShapeDefinitions.xml, Apache POI 배포본) → src/prstgeom.js
// 사용: node tools/gen-prstgeom.mjs path/to/presetShapeDefinitions.xml
import { readFileSync, writeFileSync } from 'node:fs';
import { parseXml } from '../src/xml.js';

const src = readFileSync(process.argv[2], 'utf8');
const root = parseXml(src);
const kids = (el) => (el?.children ?? []).filter((c) => c && typeof c === 'object' && c.name);
const loc = (n) => n.replace(/^.*:/, '');
const child = (el, n) => kids(el).find((c) => loc(c.name) === n);
const num = (v) => (/^-?\d+$/.test(v) ? Number(v) : v);
const out = {};
for (const shp of kids(root)) {
  const name = loc(shp.name);
  const gd = (list) => kids(list).map((g) => [g.attrs.name, g.attrs.fmla]);
  const av = gd(child(shp, 'avLst'));
  const gl = gd(child(shp, 'gdLst'));
  const r = child(shp, 'rect');
  const rect = r ? [r.attrs.l, r.attrs.t, r.attrs.r, r.attrs.b].map(num) : null;
  const paths = kids(child(shp, 'pathLst')).map((p) => {
    const a = p.attrs ?? {};
    const cmds = [];
    for (const c of kids(p)) {
      const k = loc(c.name);
      const pts = kids(c).map((pt) => [num(pt.attrs.x), num(pt.attrs.y)]);
      if (k === 'moveTo') cmds.push(['M', ...pts[0]]);
      else if (k === 'lnTo') cmds.push(['L', ...pts[0]]);
      else if (k === 'cubicBezTo') cmds.push(['C', ...pts.flat()]);
      else if (k === 'quadBezTo') cmds.push(['Q', ...pts.flat()]);
      else if (k === 'arcTo') cmds.push(['A', num(c.attrs.wR), num(c.attrs.hR), num(c.attrs.stAng), num(c.attrs.swAng)]);
      else if (k === 'close') cmds.push(['Z']);
    }
    const o = { c: cmds };
    if (a.w) o.w = Number(a.w);
    if (a.h) o.h = Number(a.h);
    if (a.fill && a.fill !== 'norm') o.f = a.fill;
    if (a.stroke === 'false' || a.stroke === '0') o.s = 0;
    return o;
  });
  // 조정 핸들 (노란 점): XY = { x: 조정값, minX, maxX, y, minY, maxY, p: [x, y] }, 극좌표 = { r, minR, maxR, a, minAng, maxAng, p }
  const hs = kids(child(shp, 'ahLst')).map((h) => {
    const at = h.attrs ?? {};
    const pos = child(h, 'pos');
    const o = { p: [num(pos.attrs.x), num(pos.attrs.y)] };
    if (loc(h.name) === 'ahXY') {
      if (at.gdRefX) Object.assign(o, { x: at.gdRefX, minX: num(at.minX ?? '0'), maxX: num(at.maxX ?? '0') });
      if (at.gdRefY) Object.assign(o, { y: at.gdRefY, minY: num(at.minY ?? '0'), maxY: num(at.maxY ?? '0') });
    } else {
      if (at.gdRefR) Object.assign(o, { r: at.gdRefR, minR: num(at.minR ?? '0'), maxR: num(at.maxR ?? '0') });
      if (at.gdRefAng) Object.assign(o, { a: at.gdRefAng, minAng: num(at.minAng ?? '0'), maxAng: num(at.maxAng ?? '0') });
    }
    return o;
  });
  out[name] = hs.length ? [av, gl, rect, paths, hs] : [av, gl, rect, paths];
}
const body = Object.entries(out).map(([k, v]) => `  ${k}: ${JSON.stringify(v)},`).join('\n');
writeFileSync(new URL('../src/prstgeom.js', import.meta.url), `// ECMA-376 Part 1 미리 설정 도형 (presetShapeDefinitions) — tools/gen-prstgeom.mjs 가 만든 파일, 손으로 고치지 않음
// 도형 이름 → [조정값 [이름, 식], 안내선 [이름, 식], 글 영역 [l, t, r, b], 경로 [{ w, h, f: 채우기 모드, s: 0 = 선 없음, c: 명령 }], 조정 핸들?]
export const PRESETS = {
${body}
};
`);
console.log(Object.keys(out).length, 'presets');
