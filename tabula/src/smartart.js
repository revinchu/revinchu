// Editable diagram model and layout. Excel export uses ordinary DrawingML shapes.
export const SMARTART_CATEGORIES = ['목록', '프로세스', '순환', '계층', '관계', '행렬', '피라미드', '그림'];
export const SMARTART_LAYOUTS = [
  ['list', '세로 목록', '목록'], ['cards', '블록 목록', '목록'], ['alternating', '교대 목록', '목록'],
  ['process', '기본 프로세스', '프로세스'], ['chevrons', '갈매기형 프로세스', '프로세스'], ['verticalProcess', '세로 프로세스', '프로세스'], ['steps', '단계 프로세스', '프로세스'],
  ['cycle', '기본 순환', '순환'], ['radialCycle', '방사형 순환', '순환'],
  ['hierarchy', '조직도', '계층'], ['horizontalHierarchy', '가로 계층', '계층'], ['treeList', '계층 목록', '계층'],
  ['radial', '기본 방사형', '관계'], ['venn', '겹치는 관계', '관계'], ['balance', '관계 비교', '관계'],
  ['matrix', '기본 행렬', '행렬'], ['titledMatrix', '제목 있는 행렬', '행렬'], ['pyramid', '분할 피라미드', '피라미드'],
  ['pictureCards', '그림 설명 목록', '그림'], ['pictureProcess', '그림 프로세스', '그림'],
].map(([id, name, category]) => ({ id, name, category }));
export const SMARTART_PALETTES = [
  ['Office', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'],
  ['푸른색', '#17365d', '#2f5597', '#4472c4', '#5b9bd5', '#1f4e79', '#366092'],
  ['초록색', '#375623', '#548235', '#70ad47', '#447c69', '#228b7c', '#306b34'],
  ['보라색', '#7030a0', '#8064a2', '#5f497a', '#a64d79', '#674ea7', '#741b47'],
];
export const SMARTART_MAX_NODES = 60;
const id = () => `sa${Date.now().toString(36)}${Math.random().toString(36).slice(2, 9)}`;
const color = s => typeof s === 'string' && /^#[\da-f]{6}$/i.test(s);
export function isSmartArt(o) { return !!o && o.kind === 'smartart' && o.smartArt?.version === 1; }
export function smartArtPicture(src) { return typeof src === 'string' && /^data:image\/(?:png|jpeg);base64,[A-Za-z0-9+/=\r\n]+$/.test(src) && src.length <= 2800000; }
export function normalizeSmartArt(value) {
  if (!value || value.version !== 1 || !SMARTART_LAYOUTS.some(x => x.id === value.layout)) throw new Error('지원하지 않는 SmartArt 형식입니다.');
  if (!Array.isArray(value.nodes) || !value.nodes.length || value.nodes.length > SMARTART_MAX_NODES) throw new Error(`SmartArt 항목은 1~${SMARTART_MAX_NODES}개까지 사용할 수 있습니다.`);
  const used = new Set(); let previous = 0, pictures = 0;
  const nodes = value.nodes.map((n, i) => {
    const key = typeof n.id === 'string' && n.id.length < 100 && !used.has(n.id) ? n.id : id(); used.add(key);
    const text = String(n.text ?? ''); if (text.length > 500) throw new Error('항목의 텍스트는 500자 이내로 입력하세요.');
    const level = Math.max(0, Math.min(5, i ? previous + 1 : 0, Math.floor(Number(n.level) || 0))); previous = level;
    if (n.picture && !smartArtPicture(n.picture)) throw new Error('SmartArt 그림은 2MB 이하 PNG/JPEG를 사용하세요.');
    pictures += n.picture?.length ?? 0; if (pictures > 14000000) throw new Error('SmartArt 그림의 합계는 10MB 이내로 제한합니다.');
    return { id: key, text, level, ...(n.picture ? { picture: n.picture } : {}) };
  });
  const palette = Array.isArray(value.palette) ? value.palette.filter(color).slice(0, 12) : [];
  return { version: 1, layout: value.layout, nodes, palette: palette.length ? palette : SMARTART_PALETTES[0].slice(1), style: ['flat', 'outline', 'shadow'].includes(value.style) ? value.style : 'flat', font: String(value.font || '맑은 고딕').slice(0, 100) };
}
export function newSmartArt(layout = 'process', box = {}, theme) {
  const palette = Array.isArray(theme) ? theme.slice(4, 10).map(c => String(c).startsWith('#') ? c : `#${c}`).filter(color) : null;
  return { id: id(), kind: 'smartart', name: 'SmartArt', x: 0, y: 0, w: 600, h: 340, ...box, smartArt: normalizeSmartArt({ version: 1, layout, nodes: Array.from({ length: 3 }, (_, i) => ({ id: id(), text: `항목 ${i + 1}`, level: layout.includes('Hierarchy') || layout === 'hierarchy' ? (i ? 1 : 0) : 0 })), palette, style: 'flat' }) };
}
function endOf(nodes, at) { let end = at + 1; while (end < nodes.length && nodes[end].level > nodes[at].level) end++; return end; }
export function editSmartArt(value, nodeId, action, text) {
  const model = normalizeSmartArt(value), nodes = model.nodes, at = nodes.findIndex(n => n.id === nodeId);
  if (at < 0) return model;
  const end = endOf(nodes, at), level = nodes[at].level;
  if (action === 'text') nodes[at].text = String(text ?? '');
  else if (action === 'add') { if (nodes.length >= SMARTART_MAX_NODES) throw new Error(`항목은 ${SMARTART_MAX_NODES}개까지 추가할 수 있습니다.`); nodes.splice(end, 0, { id: id(), text: '새 항목', level }); }
  else if (action === 'delete') { if (nodes.length === 1) throw new Error('항목은 한 개 이상 필요합니다.'); for (let i = at + 1; i < end; i++) nodes[i].level--; nodes.splice(at, 1); }
  else if (action === 'promote' && level > 0) { for (let i = at; i < end; i++) nodes[i].level--; }
  else if (action === 'demote' && at > 0 && nodes[at - 1].level >= level && nodes.slice(at, end).every(n => n.level < 5)) { for (let i = at; i < end; i++) nodes[i].level++; }
  else if (action === 'up') { let p = at - 1; while (p >= 0 && nodes[p].level > level) p--; if (p >= 0 && nodes[p].level === level) { const moved = nodes.splice(at, end - at); nodes.splice(p, 0, ...moved); } }
  else if (action === 'down' && end < nodes.length && nodes[end].level === level) { const nextEnd = endOf(nodes, end), moved = nodes.splice(at, end - at); nodes.splice(nextEnd - moved.length, 0, ...moved); }
  return normalizeSmartArt(model);
}

/** Local pixel geometry. All nodes remain present; long text shrinks consistently. */
export function smartArtParts(shape) {
  const m = normalizeSmartArt(shape.smartArt), nodes = m.nodes, n = nodes.length, W = Math.max(60, Number(shape.w) || 600), H = Math.max(40, Number(shape.h) || 340), pad = Math.min(16, W / 30, H / 20), gap = Math.min(16, W / (n * 4), H / (n * 4)), w = W - pad * 2, h = H - pad * 2, out = [], lines = [];
  const positions = [];
  const base = (i, kind, x, y, bw, bh, extra = {}) => { const fill = m.palette[i % m.palette.length]; return { id: `${shape.id}-${nodes[i].id}-${out.length}`, kind, x, y, w: Math.max(1, bw), h: Math.max(1, bh), fill: m.style === 'outline' ? '#ffffff' : fill, stroke: fill, strokeWidth: 1.2, text: nodes[i].text, color: m.style === 'outline' ? fill : '#ffffff', size: 12, font: m.font, align: 'center', valign: 'middle', pad: [4, 6, 4, 6], ...(m.style === 'shadow' ? { shadow: { dx: 2, dy: 2, blur: 2, opacity: .22 } } : {}), smartArtNode: nodes[i].id, ...extra }; };
  const add = (i, kind, x, y, bw, bh, extra) => { const s = base(i, kind, x, y, bw, bh, extra); out.push(s); positions[i] = s; return s; };
  const line = (x1, y1, x2, y2, arrow = true) => { const left = Math.min(x1, x2), top = Math.min(y1, y2); lines.push({ id: `${shape.id}-ln${lines.length}`, kind: 'line', x: left, y: top, w: Math.max(1, Math.abs(x2 - x1)), h: Math.max(1, Math.abs(y2 - y1)), path: { paths: [{ fill: false, commands: [['M', x1 > x2 ? 1 : 0, y1 > y2 ? 1 : 0], ['L', x1 > x2 ? 0 : 1, y1 > y2 ? 0 : 1]] }] }, stroke: '#7f8c99', strokeWidth: 1.5, ...(arrow ? { tailEnd: { type: 'triangle' } } : {}) }); };
  const grid = (count = n, start = 0, y = pad, height = h) => { const cols = Math.min(count, Math.max(1, Math.ceil(Math.sqrt(count * w / Math.max(1, height))))), rows = Math.ceil(count / cols), bw = (w - gap * (cols - 1)) / cols, bh = (height - gap * (rows - 1)) / rows; for (let j = 0; j < count; j++) add(j + start, 'roundRect', pad + (j % cols) * (bw + gap), y + Math.floor(j / cols) * (bh + gap), bw, bh); };
  if (['process', 'chevrons', 'verticalProcess', 'steps'].includes(m.layout)) {
    const vertical = m.layout === 'verticalProcess', count = n, bw = vertical ? w : (w - gap * (count - 1)) / count, bh = vertical ? (h - gap * (count - 1)) / count : m.layout === 'steps' ? h / Math.max(2, count) : h * .5;
    for (let i = 0; i < n; i++) add(i, m.layout === 'chevrons' ? 'chevron' : 'roundRect', pad + (vertical ? 0 : i * (bw + gap)), pad + (vertical ? i * (bh + gap) : m.layout === 'steps' ? i * (h - bh) / Math.max(1, n - 1) : (h - bh) / 2), bw, bh);
    for (let i = 1; i < n; i++) { const a = positions[i - 1], b = positions[i]; if (vertical) line(a.x + a.w / 2, a.y + a.h, b.x + b.w / 2, b.y); else line(a.x + a.w, a.y + a.h / 2, b.x, b.y + b.h / 2); }
  } else if (['cycle', 'radialCycle', 'radial'].includes(m.layout)) {
    const center = m.layout !== 'cycle', count = center ? n - 1 : n, bw = Math.min(w * .26, w / Math.max(3, Math.sqrt(n) * 2)), bh = Math.min(h * .25, h / Math.max(3, Math.sqrt(n) * 1.5));
    if (center) add(0, 'ellipse', W / 2 - bw / 2, H / 2 - bh / 2, bw, bh);
    for (let i = 0; i < count; i++) { const angle = -Math.PI / 2 + i * Math.PI * 2 / Math.max(1, count); add(i + (center ? 1 : 0), 'ellipse', W / 2 + Math.cos(angle) * (w - bw) / 2 - bw / 2, H / 2 + Math.sin(angle) * (h - bh) / 2 - bh / 2, bw, bh); }
    if (m.layout === 'radial') for (let i = 1; i < n; i++) { const a = positions[0], b = positions[i]; line(a.x + a.w / 2, a.y + a.h / 2, b.x + b.w / 2, b.y + b.h / 2, false); }
    else if (count > 1) for (let j = 0; j < count; j++) { const a = positions[j + (center ? 1 : 0)], b = positions[(j + 1) % count + (center ? 1 : 0)], ax = a.x + a.w / 2, ay = a.y + a.h / 2, bx = b.x + b.w / 2, by = b.y + b.h / 2, dx = bx - ax, dy = by - ay, start = 1 / Math.hypot(dx / (a.w / 2), dy / (a.h / 2)), end = 1 / Math.hypot(dx / (b.w / 2), dy / (b.h / 2)); line(ax + dx * start, ay + dy * start, bx - dx * end, by - dy * end); }
  } else if (['hierarchy', 'horizontalHierarchy'].includes(m.layout)) {
    const horiz = m.layout === 'horizontalHierarchy', maxLevel = Math.max(...nodes.map(x => x.level)), levels = Array.from({ length: maxLevel + 1 }, () => []), stack = [];
    nodes.forEach((x, i) => levels[x.level].push(i));
    for (let lv = 0; lv < levels.length; lv++) { const arr = levels[lv], bw = (w - gap * ((horiz ? levels.length : arr.length) - 1)) / (horiz ? levels.length : arr.length), bh = (h - gap * ((horiz ? arr.length : levels.length) - 1)) / (horiz ? arr.length : levels.length); arr.forEach((i, j) => add(i, 'roundRect', pad + (horiz ? lv : j) * (bw + gap), pad + (horiz ? j : lv) * (bh + gap), bw, bh)); }
    nodes.forEach((node, i) => { if (node.level && stack[node.level - 1] !== undefined) { const a = positions[stack[node.level - 1]], b = positions[i]; if (horiz) line(a.x + a.w, a.y + a.h / 2, b.x, b.y + b.h / 2, false); else line(a.x + a.w / 2, a.y + a.h, b.x + b.w / 2, b.y, false); } stack[node.level] = i; stack.length = node.level + 1; });
  } else if (m.layout === 'pyramid') {
    const bh = h / n; for (let i = 0; i < n; i++) { const top = i / n, bottom = (i + 1) / n, bw = w * bottom; add(i, 'freeform', W / 2 - bw / 2, pad + i * bh, bw, bh, { path: { paths: [{ commands: [['M', (1 - top / bottom) / 2, 0], ['L', (1 + top / bottom) / 2, 0], ['L', 1, .95], ['L', 0, .95], ['Z']], fill: true }] } }); }
  } else if (m.layout === 'venn') {
    const bw = w / Math.max(1.5, n * .6), bh = Math.min(h, bw); for (let i = 0; i < n; i++) add(i, 'ellipse', pad + i * (w - bw) / Math.max(1, n - 1), (H - bh) / 2, bw, bh, { fillOpacity: .7 });
  } else if (['list', 'alternating', 'treeList'].includes(m.layout)) {
    const bh = (h - gap * (n - 1)) / n; for (let i = 0; i < n; i++) { const indent = m.layout === 'treeList' ? nodes[i].level * w * .06 : m.layout === 'alternating' ? (i % 2) * w * .15 : 0; add(i, 'roundRect', pad + indent, pad + i * (bh + gap), w - (m.layout === 'alternating' ? w * .15 : indent), bh, { align: 'left' }); }
  } else if (m.layout === 'titledMatrix' && n > 1) { add(0, 'roundRect', pad, pad, w, h * .2); grid(n - 1, 1, pad + h * .2 + gap, h * .8 - gap); }
  else { grid(); if (m.layout === 'balance') { const mid = W / 2; line(mid, pad, mid, H - pad, false); } }
  if (m.layout.startsWith('picture')) {
    for (let i = 0; i < n; i++) { const s = positions[i], bh = s.h * .66; s.y += bh; s.h -= bh; const picture = nodes[i].picture; out.unshift({ id: `${shape.id}-pic-${nodes[i].id}`, kind: picture ? 'picture' : 'rect', x: s.x, y: s.y - bh, w: s.w, h: bh, ...(picture ? { src: picture } : { fill: '#eef2f6', stroke: '#8899aa', text: '그림 선택', align: 'center', valign: 'middle', color: '#52677c', font: m.font, size: 11 }), smartArtNode: nodes[i].id }); }
    if (m.layout === 'pictureProcess') for (let i = 1; i < n; i++) { const a = positions[i - 1], b = positions[i]; if (Math.abs(a.y - b.y) < 1) line(a.x + a.w, a.y + a.h / 2, b.x, b.y + b.h / 2); }
  }
  let size = 12;
  for (const s of out) if (s.text) { const len = Array.from(s.text).length, width = Math.max(1, s.w - 12), height = Math.max(1, s.h - 8); size = Math.min(size, Math.sqrt(width * height / Math.max(1, len) / .9) * .75, height * .6); }
  for (const s of out) if (s.text) s.size = Math.max(3, size);
  return [...lines, ...out];
}

export function splitSmartArt(shape) {
  return smartArtParts(shape).map((p, i) => ({ ...p, id: id(), x: shape.x + p.x, y: shape.y + p.y, z: (shape.z ?? 0) + i / 1000, placement: shape.placement, locked: shape.locked }));
}
