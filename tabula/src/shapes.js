// 그림 개체(차트·그림·도형) 공통 도우미와 도형 SVG (DOM 없음)

export const OBJECT_PROPS = ['charts', 'images', 'shapes'];
export const OBJECT_LABEL = { charts: '차트', images: '그림', shapes: '도형' };

export const SHAPE_KINDS = [
  { id: 'rect', label: '사각형' },
  { id: 'roundRect', label: '둥근 사각형' },
  { id: 'ellipse', label: '타원' },
  { id: 'triangle', label: '삼각형' },
  { id: 'arrow', label: '화살표' },
  { id: 'line', label: '선' },
  { id: 'textbox', label: '텍스트 상자' },
];

/** 새 도형 기본값 */
export function newShape(kind, box) {
  const id = `sh${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  if (kind === 'textbox') return { id, kind, ...box, fill: '#ffffff', stroke: '#000000', text: '', color: '#000000', size: 11, align: 'left' };
  if (kind === 'line') return { id, kind, ...box, fill: null, stroke: '#4472c4', strokeWidth: 1.5, text: '' };
  return { id, kind, ...box, fill: '#4472c4', stroke: '#2f528f', text: '', color: '#ffffff', size: 11, align: 'center' };
}

/** 시트에서 id 로 개체 찾기 → { prop, obj } | null */
export function findObject(sheet, id) {
  if (!id) return null;
  for (const prop of OBJECT_PROPS) {
    const obj = (sheet[prop] ?? []).find((o) => o.id === id);
    if (obj) return { prop, obj };
  }
  return null;
}

const attr = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** 도형 모양 SVG (글자는 따로 HTML 로 그림). 선 굵기만큼 안쪽으로 그려서 잘리지 않게 함 */
export function shapeSvg(sh) {
  const w = Math.max(1, sh.w);
  const h = Math.max(1, sh.h);
  const sw = sh.stroke ? (sh.strokeWidth ?? 1) : 0;
  const i = sw / 2;
  const fill = sh.fill ? attr(sh.fill) : 'none';
  const stroke = sh.stroke ? `stroke="${attr(sh.stroke)}" stroke-width="${sw}"` : 'stroke="none"';
  let body;
  switch (sh.kind) {
    case 'ellipse':
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${Math.max(0, w / 2 - i)}" ry="${Math.max(0, h / 2 - i)}" fill="${fill}" ${stroke}/>`;
      break;
    case 'roundRect': {
      const r = Math.min(w, h) * 0.1667;
      body = `<rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, h - sw)}" rx="${r}" fill="${fill}" ${stroke}/>`;
      break;
    }
    case 'triangle':
      body = `<polygon points="${w / 2},${i} ${w - i},${h - i} ${i},${h - i}" fill="${fill}" ${stroke} stroke-linejoin="round"/>`;
      break;
    case 'arrow': {
      const head = Math.min(w * 0.5, h);
      const t1 = h * 0.25;
      const t2 = h * 0.75;
      let pts = [[i, t1], [w - head, t1], [w - head, i], [w - i, h / 2], [w - head, h - i], [w - head, t2], [i, t2]];
      if (sh.flip) pts = pts.map(([x, y]) => [w - x, y]);
      body = `<polygon points="${pts.map((p) => p.join(',')).join(' ')}" fill="${fill}" ${stroke} stroke-linejoin="round"/>`;
      break;
    }
    case 'line': {
      const [x1, x2] = sh.flip ? [w, 0] : [0, w];
      const [y1, y2] = sh.flipV ? [h, 0] : [0, h];
      body = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${attr(sh.stroke ?? '#000000')}" stroke-width="${Math.max(1, sw)}"/>`;
      break;
    }
    default:
      body = `<rect x="${i}" y="${i}" width="${Math.max(0, w - sw)}" height="${Math.max(0, h - sw)}" fill="${fill}" ${stroke}/>`;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible">${body}</svg>`;
}
