// 셀 공유 경계를 한 번만 그립니다. 숨긴 행/열은 화면 좌표가 같아도 별도 경계를 만들지 않습니다.
export function gridLineWidth(width, zoom = 1, dpr = 1, pattern = 'solid') {
  const scale = Number.isFinite(zoom * dpr) && zoom * dpr > 0 ? zoom * dpr : 1;
  return Math.max(pattern === 'double' ? 3 : 1, Math.round(width * scale)) / scale;
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
