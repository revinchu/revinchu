// 도형 서식 패널의 수치/선 끝 정규화. DOM이나 통합 문서를 변경하지 않는다.
export const SHAPE_DASH_OPTIONS = [['', '실선'], ['dot', '점선'], ['dash', '파선'], ['lgDash', '긴 파선'], ['dashDot', '일점 쇄선'], ['lgDashDot', '긴 일점 쇄선'], ['lgDashDotDot', '긴 이점 쇄선'], ['sysDash', '짧은 파선'], ['sysDot', '가는 점선'], ['sysDashDot', '짧은 일점 쇄선'], ['sysDashDotDot', '짧은 이점 쇄선']];
export const SHAPE_ARROW_OPTIONS = [['none', '없음'], ['triangle', '삼각형'], ['stealth', '오목한 화살표'], ['diamond', '마름모'], ['oval', '원형'], ['arrow', '열린 화살표']];
export const SHAPE_PATTERN_OPTIONS = [['pct5', '5% 점'], ['pct25', '25% 점'], ['pct50', '50% 점'], ['horz', '가로선'], ['vert', '세로선'], ['cross', '격자'], ['dnDiag', '오른쪽 아래 대각선'], ['upDiag', '오른쪽 위 대각선'], ['diagCross', '대각선 격자']];

export function shapeSizePatch(shape, dimension, value) {
  if (!['w', 'h'].includes(dimension) || !Number.isFinite(value) || value < 0 || shape.noMove) return {};
  const size = Math.min(10000, value), patch = { [dimension]: size };
  const other = dimension === 'w' ? 'h' : 'w';
  const from = Number(shape[dimension]), to = Number(shape[other]);
  // 수평/수직 선(한쪽 0)은 그 방향을 유지하며 0/0으로 Infinity를 만들지 않는다.
  if (shape.lockAspect && from === 0 && to > 0) return { [dimension]: 0 };
  if (shape.lockAspect && from > 0 && Number.isFinite(to) && to >= 0) {
    const scaled = size * to / from;
    if (scaled > 10000) { patch[other] = 10000; patch[dimension] = 10000 * from / to; }
    else patch[other] = scaled;
  }
  return patch;
}

export function shapeArrowEnd(shape, end) {
  const legacy = end === 'headEnd' ? ['start', 'both'].includes(shape.arrow) : ['end', 'both'].includes(shape.arrow);
  return { type: legacy ? 'triangle' : 'none', w: 'med', len: 'med', ...(shape[end] ?? {}) };
}

export function shapeGradientStops(shape) {
  const source = shape.grad?.stops;
  return Array.isArray(source) && source.length >= 2 ? source.map(stop => [...stop]) : [[0, shape.fill || '#4472c4'], [1, '#ffffff']];
}

export function shapeGradientStopPatch(shape, index, patch) {
  const stops = shapeGradientStops(shape);
  if (!stops[index]) return {};
  const [offset, color, opacity] = stops[index];
  const position = Number.isFinite(patch.position) ? Math.max(stops[index - 1]?.[0] ?? 0, Math.min(stops[index + 1]?.[0] ?? 1, patch.position)) : offset;
  const alpha = Number.isFinite(patch.opacity) ? Math.max(0, Math.min(1, patch.opacity)) : opacity;
  stops[index] = [position, patch.color ?? color, ...(alpha === undefined ? [] : [alpha])];
  return { grad: { ...shape.grad, stops } };
}

export function addShapeGradientStop(shape) {
  const stops = shapeGradientStops(shape);
  if (stops.length >= 16) return { grad: { ...shape.grad, stops }, index: 0 };
  let left = 0;
  for (let i = 1; i < stops.length - 1; i++) if (stops[i + 1][0] - stops[i][0] > stops[left + 1][0] - stops[left][0]) left = i;
  const a = stops[left], b = stops[left + 1];
  const rgb = /^#[0-9a-f]{6}$/i.test(a[1]) && /^#[0-9a-f]{6}$/i.test(b[1]);
  const color = rgb ? '#' + [1, 3, 5].map(i => Math.round((parseInt(a[1].slice(i, i + 2), 16) + parseInt(b[1].slice(i, i + 2), 16)) / 2).toString(16).padStart(2, '0')).join('') : a[1];
  stops.splice(left + 1, 0, [(a[0] + b[0]) / 2, color, ((a[2] ?? 1) + (b[2] ?? 1)) / 2]);
  return { grad: { ...shape.grad, stops }, index: left + 1 };
}
