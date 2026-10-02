// Editable diagram model and layout. Excel export uses ordinary DrawingML shapes.
export const SMARTART_CATEGORIES = ['목록', '프로세스', '순환', '계층', '관계', '행렬', '피라미드', '그림'];
export const SMARTART_LAYOUTS = [
  ['list','세로 목록','목록','동등한 항목을 위에서 아래로 정리합니다.'],
  ['cards','블록 목록','목록','항목마다 넓은 카드로 핵심 내용을 나란히 설명합니다.'],
  ['alternating','교대 목록','목록','들여쓰기 위치를 번갈아 바꿔 목록의 흐름을 강조합니다.'],
  ['numberedList','번호 강조 목록','목록','큰 번호와 설명을 짝지어 순서 있는 안내를 만듭니다.'],
  ['linedList','세로 강조 목록','목록','왼쪽 색 띠와 구분선으로 각 항목을 구별합니다.'],
  ['groupedList','묶음 목록','목록','같은 수준의 정보를 두 열의 교차 묶음으로 정리합니다.'],
  ['horizontalList','가로 목록','목록','가로 방향의 제목 카드와 아래 설명을 한 줄로 비교합니다.'],
  ['bracketList','대괄호 목록','목록','공통 묶음 안에 포함되는 항목들을 표시합니다.'],
  ['process','기본 프로세스','프로세스','작업이나 사건을 왼쪽에서 오른쪽 순서로 보여 줍니다.'],
  ['chevrons','갈매기형 프로세스','프로세스','방향성이 있는 단계들을 갈매기형으로 연결합니다.'],
  ['verticalProcess','세로 프로세스','프로세스','위에서 아래로 진행하는 처리 순서를 보여 줍니다.'],
  ['steps','단계 프로세스','프로세스','높이가 달라지는 계단으로 점진적인 변화를 나타냅니다.'],
  ['timeline','기본 시간 표시 막대','프로세스','중앙 시간선의 위아래에 사건을 교대로 배치합니다.'],
  ['alternatingProcess','교대 프로세스','프로세스','단계를 높낮이로 번갈아 놓고 화살표로 연결합니다.'],
  ['snakeProcess','여러 줄 프로세스','프로세스','많은 단계를 지그재그로 이어 여러 줄에 배치합니다.'],
  ['arrowProcess','연속 화살표 프로세스','프로세스','화살표 안에 텍스트를 넣어 연속 진행을 강조합니다.'],
  ['circleProcess','원형 단계 프로세스','프로세스','원형 단계와 연결선으로 간단한 진행 순서를 나타냅니다.'],
  ['funnelProcess','범위 축소 프로세스','프로세스','높이가 점차 줄어드는 단계로 정보가 정제되는 흐름을 나타냅니다.'],
  ['cycle','기본 순환','순환','반복되는 과정의 단계들을 원형으로 연결합니다.'],
  ['radialCycle','중앙 개념 순환','순환','첫 항목을 중심에 두고 나머지 단계가 주위를 순환합니다.'],
  ['segmentedCycle','분할 고리 순환','순환','고리를 구성하는 조각으로 반복 과정의 각 부분을 나타냅니다.'],
  ['orbitCycle','이중 궤도 순환','순환','안쪽과 바깥쪽 궤도를 번갈아 이동하는 단계를 연결합니다.'],
  ['spiralCycle','나선형 진행','순환','중심에서 바깥으로 뻗는 나선으로 반복과 확장을 나타냅니다.'],
  ['hexagonCycle','육각형 순환','순환','육각형 항목을 고리로 배치하여 순환 관계를 연결합니다.'],
  ['hierarchy','조직도','계층','텍스트 창의 수준에 따라 상위 항목과 하위 항목을 연결합니다.'],
  ['horizontalHierarchy','가로 계층','계층','왼쪽 상위 항목에서 오른쪽 하위 항목으로 연결합니다.'],
  ['treeList','계층 목록','계층','텍스트 수준을 들여쓰기로 나타내며 모든 항목을 세로로 나열합니다.'],
  ['tableHierarchy','표 계층','계층','깊이별 열과 항목별 행으로 구조를 표처럼 나타냅니다.'],
  ['labeledHierarchy','머리글 계층','계층','첫 항목을 큰 머리글로 두고 하위 수준을 여러 행으로 배치합니다.'],
  ['radialHierarchy','방사형 계층','계층','상위 항목을 안쪽에 두고 깊이에 따라 바깥 고리로 배치합니다.'],
  ['compactHierarchy','압축 조직도','계층','상위 항목 아래로 연결선을 꺾어 공간을 절약하며 계층을 나타냅니다.'],
  ['radial','기본 방사형','관계','첫 항목과 그 주위의 항목 간 관계를 보여 줍니다.'],
  ['venn','겹치는 관계','관계','겹치는 원으로 공통 영역이 있는 관계를 보여 줍니다.'],
  ['balance','관계 비교','관계','중앙 기준선 양쪽에 항목을 배치해 비교합니다.'],
  ['opposing','대립 관계','관계','중앙을 향한 양쪽 화살표로 마주 보는 생각을 나타냅니다.'],
  ['converging','집중 관계','관계','여러 항목이 첫 항목으로 모이는 관계를 보여 줍니다.'],
  ['diverging','분산 관계','관계','첫 항목에서 여러 결과로 퍼지는 관계를 보여 줍니다.'],
  ['nestedCircles','포함 관계','관계','크기가 다른 원을 안쪽으로 포개어 포함 관계를 나타냅니다.'],
  ['target','목표 관계','관계','동심원과 옆의 설명을 연결해 중심 목표와 주변 범위를 보여 줍니다.'],
  ['plusRelation','합성 관계','관계','항목 사이의 더하기 표시로 구성 요소들의 결합을 나타냅니다.'],
  ['hubList','중앙 연결 목록','관계','중앙 개념의 양쪽에 관련 설명을 목록으로 배치합니다.'],
  ['matrix','기본 행렬','행렬','항목을 같은 크기의 사각형으로 나누어 비교합니다.'],
  ['titledMatrix','제목 있는 행렬','행렬','첫 항목을 제목으로 두고 나머지를 행렬로 정리합니다.'],
  ['quadrantMatrix','사분면 행렬','행렬','가운데 축으로 나뉜 네 영역에 항목을 배치합니다.'],
  ['triangularMatrix','삼각 행렬','행렬','행마다 항목 수가 늘어나는 삼각 배열로 관계를 나타냅니다.'],
  ['gridMatrix','강조 행렬','행렬','첫 행과 첫 열을 구분하는 표 모양으로 항목을 정리합니다.'],
  ['pyramid','분할 피라미드','피라미드','위가 좁고 아래가 넓은 층으로 단계나 기반을 나타냅니다.'],
  ['invertedPyramid','역 피라미드','피라미드','위가 넓고 아래가 좁은 층으로 집중되는 구조를 나타냅니다.'],
  ['pyramidList','피라미드 설명 목록','피라미드','작은 피라미드 층 옆에 각 단계의 설명을 배치합니다.'],
  ['stackedPyramid','계단형 피라미드','피라미드','너비가 점차 늘어나는 사각 층으로 단계를 비교합니다.'],
  ['funnelPyramid','깔때기','피라미드','넓은 입구와 좁은 결과로 내려가는 필터링을 나타냅니다.'],
  ['triangleGrid','삼각형 묶음','피라미드','작은 삼각형을 피라미드 배열로 쌓아 구성 요소를 보여 줍니다.'],
  ['pictureCards','그림 설명 목록','그림','그림 아래에 설명이 있는 카드를 배열합니다.'],
  ['pictureProcess','그림 프로세스','그림','그림과 설명을 가로로 놓고 순서대로 연결합니다.'],
  ['pictureList','그림 옆 설명 목록','그림','각 행의 왼쪽 그림과 오른쪽 설명을 짝지어 표시합니다.'],
  ['pictureAlternating','교대 그림 목록','그림','그림과 설명의 좌우 위치를 교대로 배치합니다.'],
  ['pictureGrid','그림 격자','그림','그림을 강조한 정사각 카드와 작은 캡션을 배열합니다.'],
  ['pictureHierarchy','그림 조직도','그림','그림이 포함된 항목을 수준에 따라 조직도로 연결합니다.'],
  ['pictureRadial','방사형 그림 관계','그림','중앙 그림 항목과 주변 그림 항목의 관계를 보여 줍니다.'],
  ['pictureCaption','그림 위 제목 목록','그림','위쪽 제목과 아래쪽 그림으로 발표용 카드를 배열합니다.'],
].map(([id, name, category, description]) => ({ id, name, category, description }));
export const SMARTART_STYLES = [['flat','단색'], ['outline','윤곽선'], ['shadow','그림자'], ['gradient','그라데이션']];
export const SMARTART_PALETTES = [
  ['Office', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'],
  ['푸른색', '#17365d', '#2f5597', '#4472c4', '#5b9bd5', '#1f4e79', '#366092'],
  ['초록색', '#375623', '#548235', '#70ad47', '#447c69', '#228b7c', '#306b34'],
  ['보라색', '#7030a0', '#8064a2', '#5f497a', '#a64d79', '#674ea7', '#741b47'],
  ['따뜻한 색', '#c55a11', '#ed7d31', '#bf9000', '#a61c00', '#d5a64a', '#843c0c'],
  ['청록색', '#006b76', '#008c95', '#159e9c', '#306b34', '#2f5597', '#548235'],
  ['차분한 색', '#475569', '#64748b', '#466c73', '#766b80', '#807260', '#5c7160'],
  ['단일 강조', '#4472c4', '#4472c4', '#4472c4', '#4472c4', '#4472c4', '#4472c4'],
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
  return { version: 1, layout: value.layout, nodes, palette: palette.length ? palette : SMARTART_PALETTES[0].slice(1), style: SMARTART_STYLES.some(([key]) => key === value.style) ? value.style : 'flat', font: String(value.font || '맑은 고딕').slice(0, 100) };
}
export function newSmartArt(layout = 'process', box = {}, theme) {
  const palette = Array.isArray(theme) ? theme.slice(4, 10).map(c => String(c).startsWith('#') ? c : `#${c}`).filter(color) : null;
  return { id: id(), kind: 'smartart', name: 'SmartArt', x: 0, y: 0, w: 600, h: 340, ...box, smartArt: normalizeSmartArt({ version: 1, layout, nodes: Array.from({ length: 3 }, (_, i) => ({ id: id(), text: `항목 ${i + 1}`, level: SMARTART_LAYOUTS.find(x => x.id === layout)?.category === '계층' || layout === 'pictureHierarchy' ? (i ? 1 : 0) : 0 })), palette, style: 'flat' }) };
}
function endOf(nodes, at) { let end = at + 1; while (end < nodes.length && nodes[end].level > nodes[at].level) end++; return end; }
export function editSmartArt(value, nodeId, action, text) {
  const model = normalizeSmartArt(value), nodes = model.nodes, at = nodes.findIndex(n => n.id === nodeId);
  if (at < 0) return model;
  const end = endOf(nodes, at), level = nodes[at].level;
  if (action === 'text') nodes[at].text = String(text ?? '');
  else if (action === 'add') { if (nodes.length >= SMARTART_MAX_NODES) throw new Error(`항목은 ${SMARTART_MAX_NODES}개까지 추가할 수 있습니다.`); nodes.splice(end, 0, { id: id(), text: '새 항목', level }); }
  else if (action === 'before') { if (nodes.length >= SMARTART_MAX_NODES) throw new Error(`항목은 ${SMARTART_MAX_NODES}개까지 추가할 수 있습니다.`); nodes.splice(at, 0, { id: id(), text: '새 항목', level }); }
  else if (action === 'child') { if (nodes.length >= SMARTART_MAX_NODES) throw new Error(`항목은 ${SMARTART_MAX_NODES}개까지 추가할 수 있습니다.`); if (level >= 5) throw new Error('수준은 6단계까지 사용할 수 있습니다.'); nodes.splice(end, 0, { id: id(), text: '새 하위 항목', level: level + 1 }); }
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
  const base = (i, kind, x, y, bw, bh, extra = {}) => { const fill = m.palette[i % m.palette.length]; return { id: `${shape.id}-${nodes[i].id}-${out.length}`, kind, x, y, w: Math.max(1, bw), h: Math.max(1, bh), fill: m.style === 'outline' ? '#ffffff' : fill, stroke: fill, strokeWidth: 1.2, text: nodes[i].text, color: m.style === 'outline' ? fill : '#ffffff', size: 12, font: m.font, align: 'center', valign: 'middle', pad: [4, 6, 4, 6], ...(m.style === 'gradient' ? { grad: { ang: 90, stops: [[0, fill], [1, '#ffffff']] }, color: '#172b4d' } : {}), ...(m.style === 'shadow' ? { shadow: { dx: 2, dy: 2, blur: 2, opacity: .22 } } : {}), smartArtNode: nodes[i].id, ...extra }; };
  const add = (i, kind, x, y, bw, bh, extra) => { const s = base(i, kind, x, y, bw, bh, extra); out.push(s); positions[i] = s; return s; };
  const line = (x1, y1, x2, y2, arrow = true) => { const left = Math.min(x1, x2), top = Math.min(y1, y2); lines.push({ id: `${shape.id}-ln${lines.length}`, kind: 'line', x: left, y: top, w: Math.max(1, Math.abs(x2 - x1)), h: Math.max(1, Math.abs(y2 - y1)), path: { paths: [{ fill: false, commands: [['M', x1 > x2 ? 1 : 0, y1 > y2 ? 1 : 0], ['L', x1 > x2 ? 0 : 1, y1 > y2 ? 0 : 1]] }] }, stroke: '#7f8c99', strokeWidth: 1.5, ...(arrow ? { tailEnd: { type: 'triangle' } } : {}) }); };
  const grid = (count = n, start = 0, y = pad, height = h) => { const cols = Math.min(count, Math.max(1, Math.ceil(Math.sqrt(count * w / Math.max(1, height))))), rows = Math.ceil(count / cols), bw = (w - gap * (cols - 1)) / cols, bh = (height - gap * (rows - 1)) / rows; for (let j = 0; j < count; j++) add(j + start, 'roundRect', pad + (j % cols) * (bw + gap), y + Math.floor(j / cols) * (bh + gap), bw, bh); };
  // 레이아웃들은 같은 미리 보기의 이름만 바꾸지 않고 배치·도형·연결 구조를 각각 계산한다.
  const decor = (kind, x, y, bw, bh, extra = {}) => { const p = { id: `${shape.id}-dec${out.length}`, kind, x, y, w: Math.max(.05, bw), h: Math.max(.05, bh), fill: '#e7edf5', stroke: '#bcc9d9', strokeWidth: 1, ...extra }; out.push(p); return p; };
  const tiles = (cols, kind = 'roundRect', x = pad, y = pad, width = w, height = h, start = 0, count = n) => {
    cols = Math.max(1, Math.min(cols, count)); const rows = Math.ceil(count / cols), bw = (width - gap * (cols - 1)) / cols, bh = (height - gap * (rows - 1)) / rows;
    for (let j = 0; j < count; j++) add(start + j, kind, x + j % cols * (bw + gap), y + Math.floor(j / cols) * (bh + gap), bw, bh);
  };
  const join = (a, b, arrow = true) => {
    const ax = a.x + a.w / 2, ay = a.y + a.h / 2, bx = b.x + b.w / 2, by = b.y + b.h / 2, dx = bx - ax, dy = by - ay;
    const start = Math.min(.45, 1 / Math.max(Math.abs(dx) / (a.w / 2), Math.abs(dy) / (a.h / 2), 1));
    const end = Math.min(.45, 1 / Math.max(Math.abs(dx) / (b.w / 2), Math.abs(dy) / (b.h / 2), 1));
    line(ax + dx * start, ay + dy * start, bx - dx * end, by - dy * end, arrow);
  };
  const chain = (closed = false) => { for (let i = 1; i < n; i++) join(positions[i - 1], positions[i]); if (closed && n > 2) join(positions[n - 1], positions[0]); };
  const levels = () => {
    const max = nodes.reduce((v, a) => Math.max(v, a.level), 0), groups = Array.from({ length: max + 1 }, () => []);
    nodes.forEach((v, i) => groups[v.level].push(i)); return groups;
  };
  const parents = () => { const stack = [], pairs = []; nodes.forEach((v, i) => { if (v.level && stack[v.level - 1] !== undefined) pairs.push([stack[v.level - 1], i]); stack[v.level] = i; stack.length = v.level + 1; }); return pairs; };
  const layout = m.layout; let custom = true;
  if (['numberedList', 'linedList', 'bracketList'].includes(layout)) {
    const rh = h / n, badge = Math.min(w * .15, rh * .75);
    if (layout === 'bracketList') { line(pad + w * .08, pad, pad + w * .08, pad + h, false); line(pad + w * .08, pad, pad + w * .18, pad, false); line(pad + w * .08, pad + h - 1, pad + w * .18, pad + h - 1, false); }
    for (let i = 0; i < n; i++) {
      const inset = layout === 'numberedList' ? badge + gap : layout === 'bracketList' ? w * .2 : w * .035;
      add(i, layout === 'linedList' ? 'rect' : 'roundRect', pad + inset, pad + i * rh, w - inset, rh - gap, { align: 'left', ...(layout === 'linedList' ? { fill: '#f3f6fb', color: m.palette[i % m.palette.length], stroke: undefined } : {}) });
      if (layout === 'numberedList') decor('ellipse', pad, pad + i * rh + (rh - gap - badge) / 2, badge, badge, { text: String(i + 1), fill: m.palette[i % m.palette.length], color: '#ffffff', size: Math.min(12, badge / 3), align: 'center', valign: 'middle' });
      if (layout === 'linedList') decor('rect', pad, pad + i * rh, w * .02, rh - gap, { fill: m.palette[i % m.palette.length], stroke: undefined });
    }
  } else if (layout === 'groupedList') {
    const rows = Math.ceil(n / 2), bh = (h - gap * (rows - 1)) / rows;
    for (let i = 0; i < n; i++) { const row = Math.floor(i / 2), col = i % 2; add(i, 'roundRect', pad + col * (w * .51), pad + row * (bh + gap) + (col ? bh * .15 : 0), w * .49, bh * .82, { align: 'left' }); if (col) join(positions[i - 1], positions[i], false); }
  } else if (layout === 'horizontalList') {
    const bw = (w - gap * (n - 1)) / n;
    for (let i = 0; i < n; i++) { decor('rect', pad + i * (bw + gap), pad, bw, h * .14, { fill: m.palette[i % m.palette.length] }); add(i, 'rect', pad + i * (bw + gap), pad + h * .18, bw, h * .82, { fill: '#f3f6fb', color: m.palette[i % m.palette.length], align: 'left' }); }
  } else if (layout === 'cards') {
    tiles(Math.ceil(Math.sqrt(n))); for (const p of positions) { p.y += p.h * .12; p.h *= .76; }
  } else if (['timeline','alternatingProcess','arrowProcess','circleProcess','funnelProcess'].includes(layout)) {
    const bw = (w - gap * (n - 1)) / n;
    if (layout === 'timeline') line(pad, H / 2, W - pad - 1, H / 2, true);
    for (let i = 0; i < n; i++) {
      const bh = layout === 'circleProcess' ? Math.min(bw, h * .65) : layout === 'funnelProcess' ? h * (1 - .6 * i / Math.max(1, n - 1)) : h * .34;
      const y = layout === 'timeline' ? (i % 2 ? H / 2 + gap * 2 : H / 2 - bh - gap * 2) : layout === 'alternatingProcess' ? pad + (i % 2 ? h - bh : 0) : pad + (h - bh) / 2;
      add(i, layout === 'arrowProcess' ? 'rightArrow' : layout === 'circleProcess' ? 'ellipse' : 'roundRect', pad + i * (bw + gap), y, bw, bh);
      if (layout === 'timeline') { const p = positions[i]; line(p.x + p.w / 2, H / 2, p.x + p.w / 2, i % 2 ? p.y : p.y + p.h, false); decor('ellipse', p.x + p.w / 2 - Math.min(4, bw / 8), H / 2 - Math.min(4, bw / 8), Math.min(8, bw / 4), Math.min(8, bw / 4), { fill: m.palette[i % m.palette.length] }); }
    }
    if (!['timeline','arrowProcess'].includes(layout)) chain();
  } else if (layout === 'snakeProcess') {
    const cols = Math.min(n, Math.max(2, Math.ceil(Math.sqrt(n * 1.5)))), rows = Math.ceil(n / cols), bw = (w - gap * (cols - 1)) / cols, bh = (h - gap * (rows - 1)) / rows;
    for (let i = 0; i < n; i++) { const r = Math.floor(i / cols), c = r % 2 ? cols - 1 - i % cols : i % cols; add(i, 'roundRect', pad + c * (bw + gap), pad + r * (bh + gap), bw, bh * .7); } chain();
  } else if (['segmentedCycle','orbitCycle','spiralCycle','hexagonCycle'].includes(layout)) {
    const bw = w / Math.max(5, Math.sqrt(n) * 2.4), bh = h / Math.max(5, Math.sqrt(n) * 2.4);
    for (let i = 0; i < n; i++) {
      const angle = -Math.PI / 2 + i * Math.PI * 2 / Math.max(1, n) * (layout === 'spiralCycle' ? 1.7 : 1), ratio = layout === 'orbitCycle' ? (i % 2 ? .52 : 1) : layout === 'spiralCycle' ? .15 + .85 * (i + 1) / n : 1;
      if (layout === 'segmentedCycle') {
        const start = -Math.PI / 2 + i * Math.PI * 2 / n, end = start + Math.PI * 2 / n * .91, pts = [];
        for (let j = 0; j <= 12; j++) { const a = start + (end - start) * j / 12; pts.push([.5 + .48 * Math.cos(a), .5 + .48 * Math.sin(a)]); }
        for (let j = 12; j >= 0; j--) { const a = start + (end - start) * j / 12; pts.push([.5 + .28 * Math.cos(a), .5 + .28 * Math.sin(a)]); }
        const x0 = Math.min(...pts.map(p => p[0])), y0 = Math.min(...pts.map(p => p[1])), x1 = Math.max(...pts.map(p => p[0])), y1 = Math.max(...pts.map(p => p[1]));
        add(i, 'freeform', pad + x0 * w, pad + y0 * h, (x1 - x0) * w, (y1 - y0) * h, { path: { paths: [{ fill: true, commands: pts.map((p,j) => [j ? 'L' : 'M', (p[0]-x0)/(x1-x0), (p[1]-y0)/(y1-y0)]).concat([['Z']]) }] } });
      } else add(i, layout === 'hexagonCycle' ? 'hexagon' : 'ellipse', W / 2 + Math.cos(angle) * (w - bw) / 2 * ratio - bw / 2, H / 2 + Math.sin(angle) * (h - bh) / 2 * ratio - bh / 2, bw, bh);
    }
    if (layout !== 'segmentedCycle') chain(layout !== 'spiralCycle');
  } else if (['tableHierarchy','labeledHierarchy','radialHierarchy','compactHierarchy'].includes(layout)) {
    const groups = levels();
    if (layout === 'tableHierarchy') { const rh = h / n, cw = w / Math.max(2, groups.length); for (let i = 0; i < n; i++) add(i, 'rect', pad + nodes[i].level * cw, pad + i * rh, w - nodes[i].level * cw, rh - gap, { align: 'left' }); }
    else if (layout === 'labeledHierarchy') {
      add(0, 'rect', pad, pad, w, h * .2);
      if (n > 1) { tiles(Math.min(3, n - 1), 'roundRect', pad, pad + h * .25, w, h * .75, 1, n - 1); for (let i = 1; i < n; i++) positions[i].x += Math.min(positions[i].w * .22, nodes[i].level * 4); for (let i = 1; i < n; i++) positions[i].w -= Math.min(positions[i].w * .22, nodes[i].level * 4); }
    } else if (layout === 'radialHierarchy') {
      const bw = w / Math.max(5, Math.sqrt(n) * 2.5), bh = h / Math.max(5, Math.sqrt(n) * 2.5);
      groups.forEach((arr,lv) => arr.forEach((i,j) => { const a = -Math.PI / 2 + 2 * Math.PI * j / arr.length, radius = groups.length === 1 ? .75 : lv === 0 && arr.length > 1 ? .2 : lv / Math.max(1,groups.length - 1); add(i,'ellipse',W/2+Math.cos(a)*(w-bw)/2*radius-bw/2,H/2+Math.sin(a)*(h-bh)/2*radius-bh/2,bw,bh); }));
    } else { const rh=h/n, cw=w/Math.max(2,groups.length+1); for(let i=0;i<n;i++)add(i,'roundRect',pad+nodes[i].level*cw*.8,pad+i*rh,cw*1.15,rh-gap); }
    for (const [ai,bi] of parents()) { const a=positions[ai],b=positions[bi]; if (layout==='compactHierarchy'||layout==='tableHierarchy') { line(a.x+a.w/2,a.y+a.h/2,a.x+a.w/2,b.y+b.h/2,false);line(a.x+a.w/2,b.y+b.h/2,b.x,b.y+b.h/2,false); } else join(a,b,false); }
  } else if (['opposing','converging','diverging','hubList'].includes(layout)) {
    if(layout==='opposing') { const rh=h/Math.ceil(n/2); for(let i=0;i<n;i++)add(i,i%2?'leftArrow':'rightArrow',pad+(i%2)*w*.55,pad+Math.floor(i/2)*rh,w*.45,rh-gap); }
    else { const center=layout==='hubList', right=layout==='converging';add(0,'ellipse',center?W/2-w*.13:right?pad+w*.72:pad,H/2-h*.2,w*.26,h*.4);const count=n-1,rh=h/Math.max(1,center?Math.ceil(count/2):count);
      for(let i=1;i<n;i++){const side=center?(i-1)%2:right?0:1,row=center?Math.floor((i-1)/2):i-1;add(i,'roundRect',pad+side*w*.74,pad+row*rh,w*.26,rh-gap);if(right)join(positions[i],positions[0]);else join(positions[0],positions[i],!center);}
    }
  } else if (layout==='nestedCircles') {
    for(let i=0;i<n;i++){const bw=w*(1-.8*i/n),bh=h*(1-.8*i/n);add(i,'ellipse',pad,pad+(h-bh)/2,bw,bh,{fillOpacity:.35,align:'right'});}
  } else if (layout==='target') {
    for(let i=0;i<n;i++){const d=h*(1-.8*i/n);decor('ellipse',pad+(w*.48-d)/2,pad+(h-d)/2,d,d,{fill:'none',stroke:m.palette[i%m.palette.length],strokeWidth:2});const p=add(i,'rect',pad+w*.57,pad+i*h/n,w*.43,h/n-gap,{align:'left'});line(pad+w*.24+d/2,pad+h/2,p.x,p.y+p.h/2,false);}
  } else if(layout==='plusRelation') {
    const bw=w/Math.max(1,n*1.4-.4),bh=Math.min(h*.6,bw);for(let i=0;i<n;i++){add(i,'roundRect',pad+i*bw*1.4,(H-bh)/2,bw,bh);if(i)decor('plus',pad+(i*1.4-.32)*bw,H/2-bh*.15,bw*.24,bh*.3,{fill:'#7f8c99',stroke:undefined});}
  } else if(['quadrantMatrix','gridMatrix','triangularMatrix','triangleGrid'].includes(layout)) {
    if(layout==='quadrantMatrix'){const cw=w/2,ch=h/2;for(let i=0;i<n;i++){const q=i%4,count=Math.ceil((n-q)/4),index=Math.floor(i/4),bh=(ch-gap*2)/count;add(i,'rect',pad+(q%2)*cw+gap,pad+Math.floor(q/2)*ch+gap+index*bh,cw-gap*2,bh-gap);}line(W/2,pad,W/2,H-pad,false);line(pad,H/2,W-pad,H/2,false);}
    else if(layout==='gridMatrix'){tiles(Math.ceil(Math.sqrt(n)),'rect');const cols=Math.ceil(Math.sqrt(n));positions.forEach((p,i)=>{if(i<cols||i%cols===0){p.strokeWidth=3;p.fillOpacity=.65;}});}
    else {let rows=1;while(rows*(rows+1)/2<n)rows++;const bh=h/rows,bw=w/rows;let i=0;for(let r=0;r<rows;r++)for(let c=0;c<=r&&i<n;c++,i++)add(i,layout==='triangleGrid'?'triangle':'rect',pad+(w-(r+1)*bw)/2+c*bw,pad+r*bh,bw-gap,bh-gap);}
  } else if(['invertedPyramid','pyramidList','stackedPyramid','funnelPyramid'].includes(layout)) {
    const bh=h/n;
    for(let i=0;i<n;i++){const invert=layout==='invertedPyramid'||layout==='funnelPyramid',top=invert?1-i/n:i/n,bottom=invert?1-(i+1)/n:(i+1)/n,bw=w*Math.max(top,bottom),x=pad+(w-bw)/2;
      if(layout==='pyramidList'){const s=add(i,'rect',pad+w*.4,pad+i*bh,w*.6,bh-gap,{align:'left'});decor('freeform',pad+w*.18-w*.18*(i+1)/n,pad+i*bh,w*.36*(i+1)/n,bh-gap,{fill:m.palette[i%m.palette.length],path:{paths:[{fill:true,commands:[['M',.5/(i+1),0],['L',1-.5/(i+1),0],['L',1,1],['L',0,1],['Z']]}]}});}
      else if(layout==='stackedPyramid')add(i,'rect',x,pad+i*bh,bw,bh-gap);
      else {const high=layout==='funnelPyramid'?Math.max(.12,top):top,low=layout==='funnelPyramid'?Math.max(.12,bottom):bottom,max=Math.max(high,low);add(i,'freeform',pad+(w-w*max)/2,pad+i*bh,w*max,bh-gap,{path:{paths:[{fill:true,commands:[['M',(1-high/max)/2,0],['L',(1+high/max)/2,0],['L',(1+low/max)/2,1],['L',(1-low/max)/2,1],['Z']]}]}});}
    }
  } else if(['pictureList','pictureAlternating'].includes(layout)) {
    const bh=(h-gap*(n-1))/n;for(let i=0;i<n;i++)add(i,'rect',pad,pad+i*(bh+gap),w,bh,{align:'left'});
  } else if(layout==='pictureGrid')tiles(Math.ceil(Math.sqrt(n)),'rect');
  else if(layout==='pictureCaption')tiles(Math.min(3,n),'roundRect');
  else if(layout==='pictureProcess'){tiles(n);for(let i=1;i<n;i++)join(positions[i-1],positions[i]);}
  else custom=false;

  if (custom) { /* 위에서 계산한 배치를 사용한다. */ } else if (['process', 'chevrons', 'verticalProcess', 'steps'].includes(m.layout)) {
    const vertical = m.layout === 'verticalProcess', count = n, bw = vertical ? w : (w - gap * (count - 1)) / count, bh = vertical ? (h - gap * (count - 1)) / count : m.layout === 'steps' ? h / Math.max(2, count) : h * .5;
    for (let i = 0; i < n; i++) add(i, m.layout === 'chevrons' ? 'chevron' : 'roundRect', pad + (vertical ? 0 : i * (bw + gap)), pad + (vertical ? i * (bh + gap) : m.layout === 'steps' ? i * (h - bh) / Math.max(1, n - 1) : (h - bh) / 2), bw, bh);
    for (let i = 1; i < n; i++) { const a = positions[i - 1], b = positions[i]; if (vertical) line(a.x + a.w / 2, a.y + a.h, b.x + b.w / 2, b.y); else line(a.x + a.w, a.y + a.h / 2, b.x, b.y + b.h / 2); }
  } else if (['cycle', 'radialCycle', 'radial', 'pictureRadial'].includes(m.layout)) {
    const center = m.layout !== 'cycle', count = center ? n - 1 : n, bw = Math.min(w * .26, w / Math.max(3, Math.sqrt(n) * 2)), bh = Math.min(h * .25, h / Math.max(3, Math.sqrt(n) * 1.5));
    if (center) add(0, 'ellipse', W / 2 - bw / 2, H / 2 - bh / 2, bw, bh);
    for (let i = 0; i < count; i++) { const angle = -Math.PI / 2 + i * Math.PI * 2 / Math.max(1, count); add(i + (center ? 1 : 0), 'ellipse', W / 2 + Math.cos(angle) * (w - bw) / 2 - bw / 2, H / 2 + Math.sin(angle) * (h - bh) / 2 - bh / 2, bw, bh); }
    if (m.layout === 'radial' || m.layout === 'pictureRadial') for (let i = 1; i < n; i++) { const a = positions[0], b = positions[i]; line(a.x + a.w / 2, a.y + a.h / 2, b.x + b.w / 2, b.y + b.h / 2, false); }
    else if (count > 1) for (let j = 0; j < count; j++) { const a = positions[j + (center ? 1 : 0)], b = positions[(j + 1) % count + (center ? 1 : 0)], ax = a.x + a.w / 2, ay = a.y + a.h / 2, bx = b.x + b.w / 2, by = b.y + b.h / 2, dx = bx - ax, dy = by - ay, start = 1 / Math.hypot(dx / (a.w / 2), dy / (a.h / 2)), end = 1 / Math.hypot(dx / (b.w / 2), dy / (b.h / 2)); line(ax + dx * start, ay + dy * start, bx - dx * end, by - dy * end); }
  } else if (['hierarchy', 'horizontalHierarchy', 'pictureHierarchy'].includes(m.layout)) {
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
    for (let i = 0; i < n; i++) {
      const s = positions[i], old = { x:s.x,y:s.y,w:s.w,h:s.h }, side = ['pictureList','pictureAlternating'].includes(layout), reverse = layout === 'pictureAlternating' && i % 2;
      let box;
      if(side){const pw=Math.min(old.w*.3,old.h*1.2);box={x:reverse?old.x+old.w-pw:old.x,y:old.y,w:pw,h:old.h};s.x+=reverse?0:pw+gap;s.w-=pw+gap;}
      else {const bh=old.h*(layout==='pictureGrid'?.8:.66),above=layout==='pictureCaption';box={x:old.x,y:above?old.y+old.h-bh:old.y,w:old.w,h:bh};s.y+=above?0:bh;s.h-=bh;}
      const picture = nodes[i].picture; out.unshift({ id: `${shape.id}-pic-${nodes[i].id}`, kind: picture ? 'picture' : 'rect', ...box,
        ...(picture ? { src: picture } : { fill:'#eef2f6',stroke:'#8899aa',text:'그림',align:'center',valign:'middle',color:'#52677c',font:m.font,size:11 }),smartArtNode:nodes[i].id });
    }
  }

  for (const p of [...lines, ...out]) { p.x = Math.max(0, Math.min(W - .05, p.x)); p.y = Math.max(0, Math.min(H - .05, p.y)); p.w = Math.max(.05, Math.min(p.w, W - p.x)); p.h = Math.max(.05, Math.min(p.h, H - p.y)); }
  let size = 12;
  for (const s of out) if (s.text) { const len = Array.from(s.text).length, width = Math.max(1, s.w - 12), height = Math.max(1, s.h - 8); size = Math.min(size, Math.sqrt(width * height / Math.max(1, len) / .9) * .75, height * .6); }
  for (const s of out) if (s.text) s.size = Math.max(3, size);
  return [...lines, ...out];
}

export function splitSmartArt(shape) {
  return smartArtParts(shape).map((p, i) => ({ ...p, id: id(), x: shape.x + p.x, y: shape.y + p.y, z: (shape.z ?? 0) + i / 1000, placement: shape.placement, locked: shape.locked }));
}
