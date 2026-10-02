import { smartArtParts } from './smartart.js';
import { shapeSvg } from './shapes.js';
import { esc } from './xml.js';

function label(shape) {
  if (!shape.text) return '';
  const px = Math.min(6, shape.w * .08), py = Math.min(4, shape.h * .08), width = Math.max(.01, shape.w - px * 2), height = Math.max(.01, shape.h - py * 2);
  const wrap = fs => {
    const lines = [];
    for (const paragraph of String(shape.text).split('\n')) {
      let row = '', used = 0;
      for (const c of paragraph) { const cw = /[^\u0000-\u00ff]/.test(c) ? fs : fs * .56; if (row && used + cw > width) { lines.push(row); row = ''; used = 0; } row += c; used += cw; }
      lines.push(row);
    }
    return lines;
  };
  let fs = Math.min((shape.size ?? 11) * 4 / 3, width), lines = wrap(fs);
  // 매우 많은 항목이나 긴 문장도 이웃 항목에 넘치지 않도록 실제 줄 수로 추가 축소한다.
  for (let i = 0; i < 12 && lines.length * fs * 1.15 > height; i++) { fs *= Math.max(.2, Math.min(.9, height / (lines.length * fs * 1.15))); lines = wrap(fs); }
  const lineH = fs * 1.15, start = (shape.h - lines.length * lineH) / 2 + fs * .9;
  const x = shape.align === 'left' ? px : shape.align === 'right' ? shape.w - px : shape.w / 2, anchor = shape.align === 'left' ? 'start' : shape.align === 'right' ? 'end' : 'middle';
  return `<text fill="${esc(shape.color ?? '#000000')}" font-family="${esc(shape.font ?? '맑은 고딕')}" font-size="${fs}" text-anchor="${anchor}">${lines.map((line, i) => `<tspan x="${x}" y="${start + i * lineH}">${esc(line)}</tspan>`).join('')}</text>`;
}
export function smartArtSvg(shape) {
  let parts; try { parts = smartArtParts(shape); } catch { return '<svg xmlns="http://www.w3.org/2000/svg"><text x="8" y="20">SmartArt 형식을 확인하세요.</text></svg>'; }
  const transform = `translate(${shape.flip ? shape.w : 0},${shape.flipV ? shape.h : 0}) scale(${shape.flip ? -1 : 1},${shape.flipV ? -1 : 1})`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${shape.w} ${shape.h}" role="img" aria-label="SmartArt" style="overflow:visible"><g transform="${transform}">${parts.map(p => `<g transform="translate(${p.x},${p.y})"${p.smartArtNode ? ` data-smartart-node="${esc(p.smartArtNode)}"` : ''}><title>${esc(p.text ?? '')}</title>${p.kind === 'picture' ? `<image href="${esc(p.src)}" width="${p.w}" height="${p.h}" preserveAspectRatio="none"/>` : shapeSvg(p) + label(p)}</g>`).join('')}</g></svg>`;
}
