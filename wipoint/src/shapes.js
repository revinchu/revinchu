// 도형 모양(OOXML prstGeom 이름) → SVG 경로 — DOM 없음
// shapePath(kind, w, h, adj) 는 (0,0)-(w,h) 상자 안의 경로 문자열을 돌려줌

const f = (n) => Math.round(n * 100) / 100;
const P = (pts) => `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' L')} Z`;

function star(n, w, h, inner = 0.38) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    const r = i % 2 ? inner : 0.5;
    pts.push([w / 2 + Math.cos(a) * r * w, h / 2 + Math.sin(a) * r * h]);
  }
  return P(pts);
}
function poly(n, w, h, rot = -Math.PI / 2) {
  const pts = [];
  for (let i = 0; i < n; i++) {
    const a = rot + (i * 2 * Math.PI) / n;
    pts.push([w / 2 + Math.cos(a) * w / 2, h / 2 + Math.sin(a) * h / 2]);
  }
  return P(pts);
}
const ellipse = (cx, cy, rx, ry) => `M${f(cx - rx)} ${f(cy)} A${f(rx)} ${f(ry)} 0 1 0 ${f(cx + rx)} ${f(cy)} A${f(rx)} ${f(ry)} 0 1 0 ${f(cx - rx)} ${f(cy)} Z`;

/** 기본 조정값 (OOXML adj, 0~100000 비율) */
export const DEFAULT_ADJ = {
  roundRect: 16667, rightArrow: 50000, leftArrow: 50000, upArrow: 50000, downArrow: 50000, leftRightArrow: 50000,
  chevron: 50000, homePlate: 50000, parallelogram: 25000, trapezoid: 25000, plus: 25000, donut: 25000, wedgeRectCallout: 0,
  round2SameRect: 16667, snip1Rect: 16667, frame: 12500, can: 25000, cube: 25000, hexagon: 25000, octagon: 29289,
};

export function shapePath(kind, w, h, adj = {}) {
  const a = (k = 'adj') => (adj[k] ?? DEFAULT_ADJ[kind] ?? 0) / 100000;
  const m = Math.min(w, h);
  switch (kind) {
    case 'ellipse': return ellipse(w / 2, h / 2, w / 2, h / 2);
    case 'roundRect': {
      const r = Math.min(a() * m, w / 2, h / 2);
      return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`;
    }
    case 'round2SameRect': {
      const r = Math.min(a() * m, w / 2, h / 2);
      return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h)} H0 V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`;
    }
    case 'snip1Rect': { const s = a() * m; return P([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]]); }
    case 'triangle': return P([[w * (adj.adj ?? 50000) / 100000, 0], [w, h], [0, h]]);
    case 'rtTriangle': return P([[0, 0], [w, h], [0, h]]);
    case 'diamond': return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'parallelogram': { const s = a() * m; return P([[s, 0], [w, 0], [w - s, h], [0, h]]); }
    case 'trapezoid': { const s = a() * m; return P([[s, 0], [w - s, 0], [w, h], [0, h]]); }
    case 'pentagon': return poly(5, w, h);
    case 'hexagon': { const s = a() * m; return P([[s, 0], [w - s, 0], [w, h / 2], [w - s, h], [s, h], [0, h / 2]]); }
    case 'octagon': { const s = a() * m; return P([[s, 0], [w - s, 0], [w, s], [w, h - s], [w - s, h], [s, h], [0, h - s], [0, s]]); }
    case 'star4': return star(4, w, h, 0.19);
    case 'star5': return star(5, w, h, 0.19);
    case 'star6': return star(6, w, h, 0.29);
    case 'star8': return star(8, w, h, 0.38);
    case 'rightArrow': { const t = h * (1 - a('adj1')) / 2; const hd = Math.min(w, h * a('adj2')); return P([[0, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [0, h - t]]); }
    case 'leftArrow': { const t = h * 0.25; const hd = Math.min(w, h / 2); return P([[w, t], [hd, t], [hd, 0], [0, h / 2], [hd, h], [hd, h - t], [w, h - t]]); }
    case 'upArrow': { const t = w * 0.25; const hd = Math.min(h, w / 2); return P([[t, h], [t, hd], [0, hd], [w / 2, 0], [w, hd], [w - t, hd], [w - t, h]]); }
    case 'downArrow': { const t = w * 0.25; const hd = Math.min(h, w / 2); return P([[t, 0], [w - t, 0], [w - t, h - hd], [w, h - hd], [w / 2, h], [0, h - hd], [t, h - hd]]); }
    case 'leftRightArrow': { const t = h * 0.25; const hd = Math.min(w / 2, h / 2); return P([[0, h / 2], [hd, 0], [hd, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [hd, h - t], [hd, h]]); }
    case 'chevron': { const s = Math.min(a() * m, w); return P([[0, 0], [w - s, 0], [w, h / 2], [w - s, h], [0, h], [s, h / 2]]); }
    case 'homePlate': { const s = Math.min(a() * m, w); return P([[0, 0], [w - s, 0], [w, h / 2], [w - s, h], [0, h]]); }
    case 'plus': { const s = a() * m; return P([[s, 0], [w - s, 0], [w - s, s], [w, s], [w, h - s], [w - s, h - s], [w - s, h], [s, h], [s, h - s], [0, h - s], [0, s], [s, s]]); }
    case 'donut': { const s = a() * m; return `${ellipse(w / 2, h / 2, w / 2, h / 2)} ${ellipse(w / 2, h / 2, Math.max(1, w / 2 - s), Math.max(1, h / 2 - s))}`; }
    case 'frame': { const s = a() * m; return `M0 0 H${f(w)} V${f(h)} H0 Z M${f(s)} ${f(s)} V${f(h - s)} H${f(w - s)} V${f(s)} Z`; }
    case 'heart': return `M${f(w / 2)} ${f(h * 0.25)} C${f(w * 0.5)} ${f(-h * 0.05)} ${f(-w * 0.05)} ${f(-h * 0.02)} ${f(w * 0.02)} ${f(h * 0.35)} C${f(w * 0.07)} ${f(h * 0.6)} ${f(w * 0.35)} ${f(h * 0.75)} ${f(w / 2)} ${f(h)} C${f(w * 0.65)} ${f(h * 0.75)} ${f(w * 0.93)} ${f(h * 0.6)} ${f(w * 0.98)} ${f(h * 0.35)} C${f(w * 1.05)} ${f(-h * 0.02)} ${f(w * 0.5)} ${f(-h * 0.05)} ${f(w / 2)} ${f(h * 0.25)} Z`;
    case 'cloud': return `M${f(w * 0.2)} ${f(h * 0.85)} C${f(w * 0.02)} ${f(h * 0.85)} ${f(-w * 0.02)} ${f(h * 0.55)} ${f(w * 0.14)} ${f(h * 0.48)} C${f(w * 0.08)} ${f(h * 0.2)} ${f(w * 0.38)} ${f(h * 0.08)} ${f(w * 0.45)} ${f(h * 0.22)} C${f(w * 0.55)} ${f(-h * 0.02)} ${f(w * 0.85)} ${f(h * 0.02)} ${f(w * 0.82)} ${f(h * 0.3)} C${f(w * 1.02)} ${f(h * 0.32)} ${f(w * 1.02)} ${f(h * 0.7)} ${f(w * 0.85)} ${f(h * 0.75)} C${f(w * 0.85)} ${f(h * 0.98)} ${f(w * 0.55)} ${f(h * 1.02)} ${f(w * 0.48)} ${f(h * 0.85)} C${f(w * 0.4)} ${f(h * 0.98)} ${f(w * 0.22)} ${f(h * 0.98)} ${f(w * 0.2)} ${f(h * 0.85)} Z`;
    case 'wedgeRectCallout': return P([[0, 0], [w, 0], [w, h * 0.8], [w * 0.4, h * 0.8], [w * 0.2, h], [w * 0.25, h * 0.8], [0, h * 0.8]]);
    case 'wedgeRoundRectCallout': { const r = m * 0.12; const b = h * 0.8; return `M${f(r)} 0 H${f(w - r)} Q${f(w)} 0 ${f(w)} ${f(r)} V${f(b - r)} Q${f(w)} ${f(b)} ${f(w - r)} ${f(b)} H${f(w * 0.4)} L${f(w * 0.2)} ${f(h)} L${f(w * 0.25)} ${f(b)} H${f(r)} Q0 ${f(b)} 0 ${f(b - r)} V${f(r)} Q0 0 ${f(r)} 0 Z`; }
    case 'wedgeEllipseCallout': return `M${f(w * 0.22)} ${f(h * 0.78)} A${f(w / 2)} ${f(h * 0.4)} 0 1 1 ${f(w * 0.35)} ${f(h * 0.79)} L${f(w * 0.15)} ${f(h)} Z`;
    case 'can': { const e = a() * m / 2; return `M0 ${f(e)} A${f(w / 2)} ${f(e)} 0 0 1 ${f(w)} ${f(e)} V${f(h - e)} A${f(w / 2)} ${f(e)} 0 0 1 0 ${f(h - e)} Z M0 ${f(e)} A${f(w / 2)} ${f(e)} 0 0 0 ${f(w)} ${f(e)}`; }
    case 'cube': { const d = a() * m; return `M0 ${f(d)} L${f(d)} 0 H${f(w)} V${f(h - d)} L${f(w - d)} ${f(h)} H0 Z M0 ${f(d)} H${f(w - d)} L${f(w)} 0 M${f(w - d)} ${f(d)} V${f(h)}`; }
    case 'flowChartTerminator': { const r = h / 2; return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`; }
    case 'flowChartDecision': return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case 'flowChartDocument': return `M0 0 H${f(w)} V${f(h * 0.83)} C${f(w * 0.7)} ${f(h * 0.7)} ${f(w * 0.35)} ${f(h * 1.1)} 0 ${f(h * 0.9)} Z`;
    case 'flowChartInputOutput': return P([[w * 0.2, 0], [w, 0], [w * 0.8, h], [0, h]]);
    case 'line': case 'straightConnector1': return `M0 0 L${f(w)} ${f(h)}`;
    case 'arc': return `M${f(w / 2)} 0 A${f(w / 2)} ${f(h / 2)} 0 0 1 ${f(w)} ${f(h / 2)}`;
    case 'blockArc': return `M0 ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 0 1 ${f(w)} ${f(h / 2)} H${f(w * 0.75)} A${f(w / 4)} ${f(h / 4)} 0 0 0 ${f(w / 4)} ${f(h / 2)} Z`;
    case 'pie': return `M${f(w / 2)} ${f(h / 2)} L${f(w)} ${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w / 2)} 0 Z`;
    case 'moon': return `M${f(w)} 0 A${f(w)} ${f(h / 2)} 0 0 0 ${f(w)} ${f(h)} A${f(w * 0.5)} ${f(h / 2)} 0 0 1 ${f(w)} 0 Z`;
    case 'smileyFace': return `${ellipse(w / 2, h / 2, w / 2, h / 2)} ${ellipse(w * 0.33, h * 0.38, w * 0.06, h * 0.06)} ${ellipse(w * 0.67, h * 0.38, w * 0.06, h * 0.06)} M${f(w * 0.28)} ${f(h * 0.65)} Q${f(w / 2)} ${f(h * 0.85)} ${f(w * 0.72)} ${f(h * 0.65)}`;
    case 'lightningBolt': return P([[w * 0.39, 0], [w * 0.7, h * 0.3], [w * 0.57, h * 0.37], [w * 0.88, h * 0.62], [w * 0.77, h * 0.67], [w, h], [w * 0.5, h * 0.72], [w * 0.62, h * 0.66], [w * 0.2, h * 0.42], [w * 0.36, h * 0.35], [0, h * 0.1]]);
    case 'teardrop': return `M${f(w)} 0 V${f(h / 2)} A${f(w / 2)} ${f(h / 2)} 0 1 1 ${f(w / 2)} 0 Z`;
    case 'notchedRightArrow': { const t = h * 0.25; const hd = Math.min(w, h / 2); return P([[0, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [0, h - t], [hd / 2, h / 2]]); }
    case 'bentArrow': return P([[0, h], [0, h * 0.45], [w * 0.65, h * 0.45], [w * 0.65, h * 0.2], [w, h * 0.55], [w * 0.65, h * 0.9], [w * 0.65, h * 0.65], [w * 0.25, h * 0.65], [w * 0.25, h]]);
    case 'uturnArrow': return `M0 ${f(h)} V${f(h * 0.4)} A${f(w * 0.35)} ${f(h * 0.4)} 0 0 1 ${f(w * 0.7)} ${f(h * 0.4)} V${f(h * 0.55)} H${f(w * 0.85)} L${f(w * 0.6)} ${f(h * 0.85)} L${f(w * 0.35)} ${f(h * 0.55)} H${f(w * 0.5)} V${f(h * 0.4)} A${f(w * 0.15)} ${f(h * 0.2)} 0 0 0 ${f(w * 0.2)} ${f(h * 0.4)} V${f(h)} Z`;
    case 'curvedRightArrow': return `M0 ${f(h * 0.2)} Q${f(w * 0.7)} ${f(h * 0.2)} ${f(w * 0.6)} ${f(h * 0.7)} H${f(w * 0.4)} L${f(w * 0.75)} ${f(h)} L${f(w)} ${f(h * 0.7)} H${f(w * 0.85)} Q${f(w * 0.85)} 0 0 0 Z`;
    case 'bracketPair': return `M${f(w * 0.12)} 0 Q0 0 0 ${f(h * 0.12)} V${f(h * 0.88)} Q0 ${f(h)} ${f(w * 0.12)} ${f(h)} M${f(w * 0.88)} 0 Q${f(w)} 0 ${f(w)} ${f(h * 0.12)} V${f(h * 0.88)} Q${f(w)} ${f(h)} ${f(w * 0.88)} ${f(h)}`;
    case 'bracePair': return `M${f(w * 0.15)} 0 Q${f(w * 0.05)} 0 ${f(w * 0.05)} ${f(h * 0.1)} V${f(h * 0.4)} Q${f(w * 0.05)} ${f(h / 2)} 0 ${f(h / 2)} Q${f(w * 0.05)} ${f(h / 2)} ${f(w * 0.05)} ${f(h * 0.6)} V${f(h * 0.9)} Q${f(w * 0.05)} ${f(h)} ${f(w * 0.15)} ${f(h)} M${f(w * 0.85)} 0 Q${f(w * 0.95)} 0 ${f(w * 0.95)} ${f(h * 0.1)} V${f(h * 0.4)} Q${f(w * 0.95)} ${f(h / 2)} ${f(w)} ${f(h / 2)} Q${f(w * 0.95)} ${f(h / 2)} ${f(w * 0.95)} ${f(h * 0.6)} V${f(h * 0.9)} Q${f(w * 0.95)} ${f(h)} ${f(w * 0.85)} ${f(h)}`;
    case 'rect': default: return `M0 0 H${f(w)} V${f(h)} H0 Z`;
  }
}

/** 채우기가 없는 열린 선 모양 */
export const OPEN_SHAPES = new Set(['line', 'straightConnector1', 'arc', 'bracketPair', 'bracePair']);
export const isLineShape = (kind) => kind === 'line' || kind === 'straightConnector1';

/** 글자 영역 (도형 안쪽 비율): [l, t, r, b] */
export function textRect(kind, w, h) {
  switch (kind) {
    case 'ellipse': case 'wedgeEllipseCallout': return [w * 0.146, h * 0.146, w * 0.854, h * 0.854];
    case 'diamond': case 'flowChartDecision': return [w / 4, h / 4, w * 0.75, h * 0.75];
    case 'triangle': return [w / 4, h / 2, w * 0.75, h];
    case 'rtTriangle': return [w / 12, h * 0.42, w * 0.58, h * 11 / 12];
    case 'wedgeRectCallout': case 'wedgeRoundRectCallout': return [0, 0, w, h * 0.8];
    case 'star5': case 'star4': case 'star6': case 'star8': return [w * 0.3, h * 0.3, w * 0.7, h * 0.75];
    case 'rightArrow': case 'leftArrow': case 'leftRightArrow': case 'notchedRightArrow': return [0, h / 4, w, h * 0.75];
    case 'chevron': case 'homePlate': return [w * 0.1, 0, w * 0.85, h];
    case 'parallelogram': case 'trapezoid': case 'flowChartInputOutput': return [w * 0.15, 0, w * 0.85, h];
    case 'hexagon': case 'octagon': return [w * 0.12, h * 0.12, w * 0.88, h * 0.88];
    case 'can': return [0, h * 0.25, w, h];
    case 'cube': return [0, h * 0.25, w * 0.75, h];
    default: return [0, 0, w, h];
  }
}

/** 도형 갤러리 (PowerPoint 삽입 › 도형 분류) */
export const SHAPE_GALLERY = [
  ['선', [['line', '선'], ['straightConnector1', '화살표'], ['arc', '호']]],
  ['사각형', [['rect', '직사각형'], ['roundRect', '모서리가 둥근 직사각형'], ['snip1Rect', '한쪽 모서리가 잘린 사각형'], ['round2SameRect', '위쪽 모서리가 둥근 사각형']]],
  ['기본 도형', [['ellipse', '타원'], ['triangle', '이등변 삼각형'], ['rtTriangle', '직각 삼각형'], ['parallelogram', '평행 사변형'], ['trapezoid', '사다리꼴'], ['diamond', '다이아몬드'], ['pentagon', '정오각형'], ['hexagon', '육각형'], ['octagon', '팔각형'], ['donut', '도넛'], ['frame', '액자'], ['can', '원통형'], ['cube', '정육면체'], ['heart', '하트'], ['lightningBolt', '번개'], ['moon', '달'], ['cloud', '구름'], ['smileyFace', '웃는 얼굴'], ['teardrop', '눈물 방울'], ['pie', '원형'], ['blockArc', '막힌 원호'], ['plus', '더하기'], ['bracketPair', '양쪽 대괄호'], ['bracePair', '양쪽 중괄호']]],
  ['블록 화살표', [['rightArrow', '오른쪽 화살표'], ['leftArrow', '왼쪽 화살표'], ['upArrow', '위쪽 화살표'], ['downArrow', '아래쪽 화살표'], ['leftRightArrow', '왼쪽/오른쪽 화살표'], ['notchedRightArrow', '톱니 모양의 오른쪽 화살표'], ['bentArrow', '굽은 화살표'], ['uturnArrow', 'U자형 화살표'], ['curvedRightArrow', '오른쪽으로 구부러진 화살표'], ['chevron', '갈매기형 수장'], ['homePlate', '오각형']]],
  ['순서도', [['rect', '순서도: 처리'], ['flowChartDecision', '순서도: 판단'], ['flowChartTerminator', '순서도: 수행의 시작/종료'], ['flowChartDocument', '순서도: 문서'], ['flowChartInputOutput', '순서도: 데이터']]],
  ['별 및 현수막', [['star4', '포인트가 4개인 별'], ['star5', '포인트가 5개인 별'], ['star6', '포인트가 6개인 별'], ['star8', '포인트가 8개인 별']]],
  ['설명선', [['wedgeRectCallout', '말풍선: 사각형'], ['wedgeRoundRectCallout', '말풍선: 모서리가 둥근 사각형'], ['wedgeEllipseCallout', '말풍선: 타원형']]],
];
export const SHAPE_LABEL = Object.fromEntries(SHAPE_GALLERY.flatMap(([, list]) => list));
export const KNOWN_SHAPES = new Set(Object.keys(SHAPE_LABEL));

/** custGeom 경로 목록 → SVG 경로 (pptx 사용자 지정 도형) — paths: [{w, h, cmds: [[op, ...pts]]}] */
export function customPath(paths, w, h) {
  const out = [];
  for (const p of paths ?? []) {
    const sx = p.w ? w / p.w : 1;
    const sy = p.h ? h / p.h : 1;
    let cx = 0;
    let cy = 0;
    for (const [op, ...v] of p.cmds) {
      if (op === 'M') { cx = v[0]; cy = v[1]; out.push(`M${f(v[0] * sx)} ${f(v[1] * sy)}`); }
      else if (op === 'L') { cx = v[0]; cy = v[1]; out.push(`L${f(v[0] * sx)} ${f(v[1] * sy)}`); }
      else if (op === 'C') { out.push(`C${f(v[0] * sx)} ${f(v[1] * sy)} ${f(v[2] * sx)} ${f(v[3] * sy)} ${f(v[4] * sx)} ${f(v[5] * sy)}`); cx = v[4]; cy = v[5]; }
      else if (op === 'Q') { out.push(`Q${f(v[0] * sx)} ${f(v[1] * sy)} ${f(v[2] * sx)} ${f(v[3] * sy)}`); cx = v[2]; cy = v[3]; }
      else if (op === 'A') {
        // arcTo: wR hR stAng swAng (각도는 60000 분의 1도) — 현재 점에서 시작
        const [wr, hr, st, sw] = v;
        const a1 = (st / 60000) * Math.PI / 180;
        const a2 = ((st + sw) / 60000) * Math.PI / 180;
        const ox = cx - wr * Math.cos(a1);
        const oy = cy - hr * Math.sin(a1);
        const ex = ox + wr * Math.cos(a2);
        const ey = oy + hr * Math.sin(a2);
        const large = Math.abs(sw / 60000) > 180 ? 1 : 0;
        const sweep = sw > 0 ? 1 : 0;
        out.push(`A${f(wr * sx)} ${f(hr * sy)} 0 ${large} ${sweep} ${f(ex * sx)} ${f(ey * sy)}`);
        cx = ex; cy = ey;
      } else if (op === 'Z') out.push('Z');
    }
  }
  return out.join(' ');
}
