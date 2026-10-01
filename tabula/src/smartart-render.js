import { smartArtParts } from './smartart.js';
import { shapeSvg } from './shapes.js';
import { esc } from './xml.js';

function label(shape) {
  if (!shape.text) return '';
  const fs = (shape.size ?? 11) * 4 / 3, width = Math.max(1, shape.w - 12), lines = [];
  for (const paragraph of String(shape.text).split('\n')) {
    let row = '', used = 0;
    for (const c of paragraph) { const cw = /[^\u0000-\u00ff]/.test(c) ? fs : fs * .56; if (row && used + cw > width) { lines.push(row); row = ''; used = 0; } row += c; used += cw; }
    lines.push(row);
  }
  const lineH = fs * 1.15, start = (shape.h - lines.length * lineH) / 2 + fs * .9;
  return `<text fill="${esc(shape.color ?? '#000000')}" font-family="${esc(shape.font ?? '맑은 고딕')}" font-size="${fs}" text-anchor="${shape.align === 'left' ? 'start' : 'middle'}">${lines.map((line, i) => `<tspan x="${shape.align === 'left' ? 6 : shape.w / 2}" y="${start + i * lineH}">${esc(line)}</tspan>`).join('')}</text>`;
}
export function smartArtSvg(shape) {
  let parts; try { parts = smartArtParts(shape); } catch { return '<svg xmlns="http://www.w3.org/2000/svg"><text x="8" y="20">SmartArt 형식을 확인하세요.</text></svg>'; }
  const transform = `translate(${shape.flip ? shape.w : 0},${shape.flipV ? shape.h : 0}) scale(${shape.flip ? -1 : 1},${shape.flipV ? -1 : 1})`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${shape.w} ${shape.h}" role="img" aria-label="SmartArt" style="overflow:visible"><g transform="${transform}">${parts.map(p => `<g transform="translate(${p.x},${p.y})"${p.smartArtNode ? ` data-smartart-node="${esc(p.smartArtNode)}"` : ''}>${p.kind === 'picture' ? `<image href="${esc(p.src)}" width="${p.w}" height="${p.h}" preserveAspectRatio="none"/>` : shapeSvg(p) + label(p)}</g>`).join('')}</g></svg>`;
}
