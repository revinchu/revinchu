import { axisTextWidth } from './chart-axis-labels.js';

// 자동 데이터 레이블만 가까운 빈 위치로 옮긴다. 값/축 범위/원본 서식은 변경하지 않는다.
export function chartLabelPlacer(bounds) {
  const buckets = new Map(), step = 64, gap = 3;
  const keys = b => {
    const result = [];
    for (let y = Math.floor((b.y - gap) / step); y <= Math.floor((b.y + b.h + gap) / step); y++)
      for (let x = Math.floor((b.x - gap) / step); x <= Math.floor((b.x + b.w + gap) / step); x++) result.push(x + ',' + y);
    return result;
  };
  const collision = b => keys(b).some(k => buckets.get(k)?.some(a => b.x < a.x + a.w + gap && b.x + b.w + gap > a.x && b.y < a.y + a.h + gap && b.y + b.h + gap > a.y));
  return (x, y, text, font, anchor = 'middle', automatic = true) => {
    const w = axisTextWidth(String(text).trim(), font), h = font * 1.3;
    const box = (px, py) => ({ x: px - (anchor === 'middle' ? w / 2 : anchor === 'end' ? w : 0), y: py - font, w, h });
    let b = box(x, y);
    if (automatic && collision(b)) {
      const candidates = [[0, -h - gap], [w + gap, 0], [-w - gap, 0], [0, h + gap], [w * 2 + gap, 0], [-w * 2 - gap, 0]];
      for (const [dx, dy] of candidates) {
        const next = box(x + dx, y + dy);
        if (next.x < bounds.x || next.y < bounds.y || next.x + w > bounds.x + bounds.w || next.y + h > bounds.y + bounds.h || collision(next)) continue;
        x += dx; y += dy; b = next; break;
      }
    }
    for (const k of keys(b)) { if (!buckets.has(k)) buckets.set(k, []); buckets.get(k).push(b); }
    return { x, y };
  };
}
