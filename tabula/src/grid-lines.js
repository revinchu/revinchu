// 셀 공유 경계를 한 번만 그립니다. 숨긴 행/열은 화면 좌표가 같아도 별도 경계를 만들지 않습니다.
export function gridLineWidth(width, zoom = 1, dpr = 1, pattern = 'solid') {
  const scale = Number.isFinite(zoom * dpr) && zoom * dpr > 0 ? zoom * dpr : 1;
  // 가는 선은 1장치 픽셀, 이중선은 선/간격/선 각각 1픽셀로 유지한다.
  // DPI 반올림 때문에 가는 선이 두꺼워지거나 두 획이 비대칭이 되는 것을 막는다.
  return (pattern === 'double' ? 3 : width <= 1 ? 1 : Math.max(1, Math.round(width * scale))) / scale;
}

// 파일의 테두리 색은 보통 #RRGGBB이며, 직접 입력한 CSS 기본 색도 처리한다.
// DOM에 의존하는 currentColor/var() 등은 밝은 쪽으로 놓고 원래 순서를 유지한다.
function borderBrightness(color) {
  const names = { black: '#000', white: '#fff', gray: '#808080', grey: '#808080', silver: '#c0c0c0', red: '#f00', green: '#008000', blue: '#00f', yellow: '#ff0', lime: '#0f0', aqua: '#0ff', cyan: '#0ff', fuchsia: '#f0f', magenta: '#f0f', maroon: '#800000', olive: '#808000', navy: '#000080', teal: '#008080', purple: '#800080', transparent: '#0000' };
  let text = typeof color === 'string' ? color.trim().toLowerCase() : '';
  text = names[text] ?? text;
  let rgb, alpha = 1;
  if (/^#[\da-f]{3,4}$/.test(text)) text = '#' + [...text.slice(1)].map(c => c + c).join('');
  if (/^#[\da-f]{6}(?:[\da-f]{2})?$/.test(text)) {
    rgb = [1, 3, 5].map(i => parseInt(text.slice(i, i + 2), 16) / 255);
    if (text.length === 9) alpha = parseInt(text.slice(7, 9), 16) / 255;
  } else {
    const match = /^rgba?\(([^)]+)\)$/.exec(text);
    if (!match) return 1;
    const fields = match[1].trim().split(/[\s,/]+/);
    if (fields.length < 3 || fields.length > 4 || fields.some(s => !/^[+-]?(?:\d+\.?\d*|\.\d+)%?$/.test(s))) return 1;
    const value = (s, base) => Math.min(1, Math.max(0, parseFloat(s) / (s.endsWith('%') ? 100 : base)));
    rgb = fields.slice(0, 3).map(s => value(s, 255));
    if (fields.length === 4) alpha = value(fields[3], 1);
  }
  // 투명 선은 흰 바탕에 합성한 밝기로 비교한다. 색 자체는 변경하지 않는다.
  return 1 - alpha + alpha * (rgb[0] * .2126 + rgb[1] * .7152 + rgb[2] * .0722);
}

/** 공유 경계 결정 후의 그리기 순서. 강한 선을 나중에 그려 교차점 단절을 막는다. */
export function gridBorderPaintOrder(edges) {
  const colors = new Map();
  return edges.map((edge, index) => {
    if (!colors.has(edge.color)) colors.set(edge.color, borderBrightness(edge.color));
    return { edge, index, width: Number.isFinite(edge.width) ? edge.width : 0, brightness: colors.get(edge.color) };
  }).sort((a, b) => a.width - b.width
    || Number(a.edge.pattern === 'double') - Number(b.edge.pattern === 'double')
    || b.brightness - a.brightness
    || a.index - b.index).map(item => item.edge);
}

/** {vertical, at, start, end, width, pattern, color} → 겹치는 구간마다 우선하는 선 한 개. */
export function resolveGridBorders(edges) {
  const lines = new Map();
  for (const e of edges) {
    if (!(e.end > e.start)) continue;
    const key = (e.vertical ? 'v' : 'h') + ':' + e.at;
    if (!lines.has(key)) lines.set(key, []);
    lines.get(key).push(e);
  }
  const out = [];
  for (const segments of lines.values()) {
    const events = new Map();
    const event = (at) => { if (!events.has(at)) events.set(at, { add: [], remove: [] }); return events.get(at); };
    for (const e of segments) { event(e.start).add.push(e); event(e.end).remove.push(e); }
    const points = [...events.keys()].sort((a, b) => a - b), active = new Set();
    let prev = null;
    for (let i = 0; i + 1 < points.length; i++) {
      const start = points[i], end = points[i + 1], ev = events.get(start);
      for (const e of ev.remove) active.delete(e);
      for (const e of ev.add) active.add(e);
      let best = null;
      for (const e of active) if (!best || e.width > best.width || e.width === best.width && e.pattern === 'double' && best.pattern !== 'double') best = e;
      if (!best) { prev = null; continue; }
      if (prev && prev.end === start && prev.width === best.width && prev.pattern === best.pattern && prev.color === best.color) prev.end = end;
      else { prev = { ...best, start, end }; out.push(prev); }
    }
  }
  return out;
}
