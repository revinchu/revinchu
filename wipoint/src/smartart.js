// SmartArt 도식 (도형 묶음으로 생성 → 일반 도형처럼 편집 · pptx 그룹으로 저장) — DOM 없음
//   항목 글: 한 줄에 하나, 앞에 Tab · 공백 2칸 · '-' 를 붙이면 하위 항목 (조직도 · 글머리 목록 등)
//   opts: { multicolor | colors: 'multi'|'accent1'…'accent6'|'gradient1'…, style: 'flat'|'subtle'|'moderate'|'intense'|'outline' }
import { newShape, newImage, uid, para } from './model.js';

/** [키, 이름, 범주] — PowerPoint SmartArt 범주 순서 */
export const SMART_KINDS = [
  // 목록형
  ['blocks', '기본 블록 목록', '목록형'], ['list', '세로 번호 목록', '목록형'], ['vBullet', '세로 글머리 기호 목록', '목록형'], ['hBullet', '가로 글머리 기호 목록', '목록형'],
  ['trapezoidList', '사다리꼴 목록', '목록형'], ['tableList', '표 목록', '목록형'], ['hexList', '육각형 목록', '목록형'], ['boxList', '세로 상자 목록', '목록형'], ['stackedList', '누적 목록', '목록형'],
  // 프로세스형
  ['process', '기본 프로세스', '프로세스형'], ['chevron', '갈매기형 프로세스', '프로세스형'], ['arrowProcess', '연속 블록 프로세스', '프로세스형'], ['stepUp', '단계 상승 프로세스', '프로세스형'],
  ['vProcess', '세로 프로세스', '프로세스형'], ['bentProcess', '굽은 프로세스', '프로세스형'], ['circleProcess', '원형 프로세스', '프로세스형'], ['vChevron', '세로 갈매기형 목록', '프로세스형'],
  ['timeline', '기본 타임라인', '프로세스형'], ['funnel', '깔때기형', '프로세스형'],
  // 주기형
  ['cycle', '기본 주기형', '주기형'], ['textCycle', '텍스트 주기형', '주기형'], ['segmentedCycle', '분할 주기형', '주기형'], ['radialCycle', '방사형 주기', '주기형'], ['continuousCycle', '연속 주기형', '주기형'], ['blockCycle', '블록 주기형', '주기형'],
  // 계층 구조형
  ['hierarchy', '조직도', '계층 구조형'], ['hHierarchy', '가로 계층 구조', '계층 구조형'], ['tableHierarchy', '표 계층 구조', '계층 구조형'], ['nameHierarchy', '이름 및 직위 조직도', '계층 구조형'],
  // 관계형
  ['venn', '기본 벤형', '관계형'], ['targetVenn', '대상 벤형', '관계형'], ['stackedVenn', '누적 벤형', '관계형'], ['radialVenn', '방사형 벤형', '관계형'], ['converging', '수렴 방사형', '관계형'],
  ['counterArrows', '반대 방향 화살표', '관계형'], ['equation', '수식', '관계형'], ['balance', '균형', '관계형'],
  // 행렬형
  ['matrix', '기본 행렬', '행렬형'], ['titledMatrix', '제목 있는 행렬', '행렬형'], ['cycleMatrix', '주기 행렬', '행렬형'],
  // 피라미드형
  ['pyramid', '기본 피라미드', '피라미드형'], ['invPyramid', '역피라미드', '피라미드형'], ['pyramidList', '피라미드 목록', '피라미드형'], ['segPyramid', '분할 피라미드', '피라미드형'],
  // 그림
  ['picAccent', '그림 강조 목록', '그림'], ['picGrid', '그림 격자', '그림'], ['picCircles', '원형 그림 캡션', '그림'], ['picTitles', '제목 있는 그림 블록', '그림'],
];
export const SMART_CATS = ['목록형', '프로세스형', '주기형', '계층 구조형', '관계형', '행렬형', '피라미드형', '그림'];
export const SMART_LABEL = Object.fromEntries(SMART_KINDS.map(([k, l]) => [k, l]));
/** 색 변경 (SmartArt 디자인 › 색 변경) */
export const SMART_COLORS = [
  ['multi', '색상형 - 강조색'], ['multi2', '색상형 범위 - 강조색 2/3'], ['accent1', '강조 1'], ['accent2', '강조 2'], ['accent3', '강조 3'], ['accent4', '강조 4'], ['accent5', '강조 5'], ['accent6', '강조 6'],
  ['gradient1', '그라데이션 범위 - 강조 1'], ['gradient2', '그라데이션 범위 - 강조 2'], ['dark', '어둡게 2 채우기'],
];
/** 스타일 (SmartArt 스타일) */
export const SMART_STYLES = [['flat', '단순 채우기'], ['outline', '흰색 윤곽선'], ['subtle', '미세 효과'], ['moderate', '보통 효과'], ['intense', '강한 효과'], ['inset', '삽입'], ['cartoon', '만화']];

const ACC = ['@accent1', '@accent2', '@accent3', '@accent4', '@accent5', '@accent6'];
const GRAD_STEPS = ['', ':lm75', ':lm60:lo40', ':lm40:lo60', ':lm50', ':lm80:lo20'];

/** 줄 글 → 나무 [{ t, kids: [] }] (들여쓰기 · '-' = 하위) */
export function parseItems(items) {
  const lines = (Array.isArray(items) ? items : [items]).flatMap((s) => String(s ?? '').replace(/\r/g, '').split('\n')).filter((s) => s.trim());
  const root = { t: '', kids: [] };
  const stack = [{ lvl: -1, node: root }];
  for (const line of lines) {
    const m = /^([\t ]*)(-*)\s*(.*)$/s.exec(line);
    const lvl = (m[1].replace(/ {2}/g, '\t').match(/\t/g)?.length ?? 0) + m[2].length;
    const node = { t: m[3].trim(), kids: [] };
    while (stack.length > 1 && stack[stack.length - 1].lvl >= lvl) stack.pop();
    stack[stack.length - 1].node.kids.push(node);
    stack.push({ lvl, node });
  }
  if (!root.kids.length) root.kids.push({ t: '텍스트', kids: [] });
  return root.kids;
}
const flatText = (n) => [n.t, ...n.kids.flatMap((k) => flatText(k).map((x) => `\t${x}`))];

export function smartArt(kind, items, rect, opts = {}) {
  const tree = parseItems(items);
  const top = tree.map((n) => n.t);
  const n = top.length;
  const colors = opts.colors ?? (opts.multicolor === false ? 'accent1' : 'multi');
  const style = opts.style ?? 'flat';
  const col = (i, total = n) => {
    if (colors === 'multi') return ACC[i % ACC.length];
    if (colors === 'multi2') return ACC[(i % 2) + 1];
    if (colors === 'dark') return '@dk2';
    const g = /^gradient(\d)$/.exec(colors);
    if (g) { const k = total > 1 ? i / (total - 1) : 0; return `@accent${g[1]}${k < 0.2 ? '' : k < 0.45 ? ':lm75' : k < 0.7 ? ':lm60:lo40' : ':lm40:lo60'}`; }
    return `@${colors.startsWith('accent') ? colors : 'accent1'}${GRAD_STEPS[0]}`;
  };
  const light = (i) => `${col(i)}:lm20:lo80`;
  const { x, y, w, h } = rect;
  const fs = Math.max(10, Math.min(24, Math.round(Math.min(w / Math.max(1, n), h) / 7)));
  const out = [];
  const add = (o) => { out.push(o); return o; };
  const lineColor = style === 'outline' || style === 'flat' ? { color: '@lt1', width: 1.33 } : null;
  function box(shape, bx, by, bw, bh, text, color, o = {}) {
    const s = newShape(shape, { x: bx, y: by, w: Math.max(4, bw), h: Math.max(4, bh) });
    s.fill = color ? { type: 'solid', color } : null;
    s.line = o.line === undefined ? lineColor : o.line;
    const lines = Array.isArray(text) ? text : [text];
    s.text.paras = lines.map((t, k) => para(String(t ?? ''), { align: o.align ?? 'ctr', ...(o.bullets && k > 0 ? { bullet: { type: 'char', char: '•' }, marL: 14, indent: -10 } : {}) }, { size: k === 0 ? o.size ?? fs : Math.max(9, (o.size ?? fs) - 4), ...(o.b && k === 0 ? { b: true } : {}) }));
    s.text.anchor = o.anchor ?? 'ctr';
    s.text.defColor = o.textColor ?? (color && !/lo[6-9]\d/.test(color) ? '@lt1' : '@tx1');
    if (o.rot) s.rot = o.rot;
    if (o.adj) s.adj = o.adj;
    return add(s);
  }
  const arrow = (shape, ax, ay, aw, ah, rot = 0) => { const a = newShape(shape, { x: ax, y: ay, w: aw, h: ah }); a.fill = { type: 'solid', color: '@tx1:lm25:lo75' }; a.line = null; a.text = null; if (rot) a.rot = rot; return add(a); };
  const ln = (x1, y1, x2, y2, c = '@tx1:lm50:lo50') => { const l = newShape('line', { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) }); if ((x2 < x1) !== (y2 < y1) && x1 !== x2 && y1 !== y2) l.flipV = true; l.line = { color: c, width: 1.5 }; l.text = null; return add(l); };
  const pic = (px, py, pw, ph, shape = 'rect') => { const im = newImage(null, { x: px, y: py, w: pw, h: ph }); im.picSlot = true; if (shape !== 'rect') im.shape = shape; return add(im); };
  /** 부채꼴 (중심 cx, cy, 반지름 r, 시작 각 st · 크기 sw, 도) → 자유형 */
  const wedge = (cx, cy, r, st, sw, color) => {
    const cmds = [['M', r, r]];
    const segs = Math.max(1, Math.ceil(sw / 90));
    const pt = (a) => [r + r * Math.cos((a * Math.PI) / 180), r + r * Math.sin((a * Math.PI) / 180)];
    cmds.push(['L', ...pt(st)]);
    for (let k = 0; k < segs; k++) {
      const a0 = st + (sw * k) / segs; const a1 = st + (sw * (k + 1)) / segs;
      const t = (4 / 3) * Math.tan(((a1 - a0) * Math.PI) / 720) * r;
      const [x0, y0] = pt(a0); const [x3, y3] = pt(a1);
      const r0 = (a0 * Math.PI) / 180; const r1 = (a1 * Math.PI) / 180;
      cmds.push(['C', x0 - t * Math.sin(r0), y0 + t * Math.cos(r0), x3 + t * Math.sin(r1), y3 - t * Math.cos(r1), x3, y3]);
    }
    cmds.push(['Z']);
    const o = newShape('rect', { x: cx - r, y: cy - r, w: r * 2, h: r * 2 });
    o.path = [{ w: r * 2, h: r * 2, cmds }];
    o.fill = { type: 'solid', color }; o.line = { color: '@lt1', width: 2 }; o.text = null;
    return add(o);
  };
  const kidsText = (node) => node.kids.map((k) => k.t);

  switch (kind) {
    // ───── 목록형 ─────
    case 'blocks': {
      const cols = n <= 3 ? n : n === 4 ? 2 : 3;
      const rows = Math.ceil(n / cols);
      const gx = w * 0.03; const gy = h * 0.05;
      const bw = (w - gx * (cols - 1)) / cols;
      const bh = Math.min((h - gy * (rows - 1)) / rows, bw * 0.75);
      const oy = y + (h - (bh * rows + gy * (rows - 1))) / 2;
      tree.forEach((nd, i) => box('rect', x + (i % cols) * (bw + gx), oy + Math.floor(i / cols) * (bh + gy), bw, bh, [nd.t, ...kidsText(nd)], col(i), { bullets: true }));
      break;
    }
    case 'list': {
      const gap = h * 0.03; const bh = (h - gap * (n - 1)) / n;
      tree.forEach((nd, i) => {
        box('roundRect', x, y + i * (bh + gap), w * 0.18, bh, String(i + 1), col(i), { size: fs + 4, b: true });
        box('rect', x + w * 0.19, y + i * (bh + gap), w * 0.81, bh, [nd.t, ...kidsText(nd)], light(i), { textColor: '@tx1', line: null, align: 'l', bullets: true });
      });
      break;
    }
    case 'vBullet': {
      const per = h / n;
      tree.forEach((nd, i) => {
        const hh = per * (nd.kids.length ? 0.4 : 0.8);
        box('roundRect', x, y + i * per, w, hh, nd.t, col(i), { align: 'l', b: true });
        if (nd.kids.length) box('rect', x + w * 0.04, y + i * per + hh, w * 0.96, per * 0.55, kidsText(nd).map((t, k) => (k === 0 ? t : t)), light(i), { textColor: '@tx1', line: null, align: 'l', anchor: 't', size: fs - 2 });
      });
      break;
    }
    case 'hBullet': {
      const gap = w * 0.02; const bw = (w - gap * (n - 1)) / n;
      tree.forEach((nd, i) => {
        const bx = x + i * (bw + gap);
        box('rect', bx, y, bw, h * 0.22, nd.t, col(i), { b: true });
        box('rect', bx, y + h * 0.22, bw, h * 0.78, kidsText(nd).length ? kidsText(nd) : [''], light(i), { textColor: '@tx1', line: null, align: 'l', anchor: 't', size: fs - 2 });
      });
      break;
    }
    case 'trapezoidList': {
      const gap = w * 0.015; const bw = (w - gap * (n - 1)) / n; const bh = Math.min(h, bw * 1.1);
      tree.forEach((nd, i) => box('trapezoid', x + i * (bw + gap), y + (h - bh) / 2, bw, bh, [nd.t, ...kidsText(nd)], col(i), { rot: 180, adj: { adj: 12000 }, bullets: true }));
      // 사다리꼴을 뒤집었으므로 글자를 바로 세움
      for (const o of out) if (o.shape === 'trapezoid') { o.rot = 0; o.flipV = true; }
      break;
    }
    case 'tableList': {
      const [head, ...rest] = tree;
      box('rect', x, y, w, h * 0.3, head.t, col(0), { b: true, size: fs + 2 });
      const m = Math.max(1, rest.length); const bw = w / m;
      (rest.length ? rest : [{ t: '', kids: [] }]).forEach((nd, i) => box('rect', x + i * bw, y + h * 0.32, bw - 3, h * 0.6, [nd.t, ...kidsText(nd)], col(i + 1), { bullets: true }));
      box('rect', x, y + h * 0.94, w, h * 0.06, '', '@tx1:lm50:lo50', { line: null }).text = null;
      break;
    }
    case 'hexList': {
      const d = Math.min(h * 0.8, (w / (n * 0.8 + 0.25)));
      const total = d * (0.8 * (n - 1) + 1);
      tree.forEach((nd, i) => box('hexagon', x + (w - total) / 2 + i * d * 0.8, y + (h - d * 0.87) / 2 + (i % 2 ? d * 0.25 : -d * 0.05), d, d * 0.87, nd.t, col(i), { adj: { adj: 25000, vf: 115470 } }));
      break;
    }
    case 'boxList': {
      const gap = h * 0.04; const bh = (h - gap * (n - 1)) / n;
      tree.forEach((nd, i) => {
        box('rect', x + w * 0.06, y + i * (bh + gap) + bh * 0.3, w * 0.94, bh * 0.7, kidsText(nd).length ? kidsText(nd) : [''], '@lt1', { line: { color: col(i), width: 1.5 }, textColor: '@tx1', align: 'l', size: fs - 2 });
        box('roundRect', x, y + i * (bh + gap), w * 0.45, bh * 0.45, nd.t, col(i), { b: true });
      });
      break;
    }
    case 'stackedList': {
      const gap = w * 0.03; const bw = (w - gap * (n - 1)) / n;
      tree.forEach((nd, i) => {
        const bx = x + i * (bw + gap);
        box('roundRect', bx + bw * 0.12, y + bw * 0.45, bw * 0.88, h - bw * 0.45, kidsText(nd).length ? kidsText(nd) : [''], light(i), { textColor: '@tx1', line: null, anchor: 'b', size: fs - 2 });
        box('ellipse', bx, y, bw * 0.75, bw * 0.75, nd.t, col(i), { b: true });
      });
      break;
    }
    // ───── 프로세스형 ─────
    case 'chevron': {
      const gap = w * 0.01; const bw = (w - gap * (n - 1)) / n; const bh = Math.min(h, bw * 0.55);
      tree.forEach((nd, i) => box(i === 0 ? 'homePlate' : 'chevron', x + i * (bw + gap), y + (h - bh) / 2, bw, bh, nd.t, col(i), { line: null }));
      break;
    }
    case 'arrowProcess': {
      arrow('rightArrow', x, y + h * 0.1, w, h * 0.8).fill = { type: 'solid', color: `${col(0)}:lm20:lo80` };
      const gap = w * 0.03; const bw = (w * 0.86 - gap * (n - 1)) / n; const bh = h * 0.42;
      tree.forEach((nd, i) => box('roundRect', x + w * 0.03 + i * (bw + gap), y + (h - bh) / 2, bw, bh, nd.t, col(i)));
      break;
    }
    case 'stepUp': {
      const bw = w / n; const step = (h * 0.6) / Math.max(1, n - 1);
      tree.forEach((nd, i) => {
        const by = y + h * 0.6 - i * step;
        box('rect', x + i * bw, by, bw - 4, h - (by - y), [nd.t, ...kidsText(nd)], col(i), { anchor: 't', bullets: true });
      });
      break;
    }
    case 'vProcess': {
      const ah = h * 0.06; const bh = (h - ah * 1.5 * (n - 1)) / n; const bw = Math.min(w, bh * 5);
      tree.forEach((nd, i) => {
        const by = y + i * (bh + ah * 1.5);
        box('roundRect', x + (w - bw) / 2, by, bw, bh, nd.t, col(i));
        if (i < n - 1) arrow('downArrow', x + w / 2 - ah * 0.6, by + bh + ah * 0.25, ah * 1.2, ah);
      });
      break;
    }
    case 'bentProcess': {
      const cols = Math.min(3, n); const rows = Math.ceil(n / cols);
      const aw = w * 0.06; const bw = (w - aw * 1.5 * (cols - 1)) / cols; const bh = Math.min((h - h * 0.1 * (rows - 1)) / rows, bw * 0.6);
      tree.forEach((nd, i) => {
        const r = Math.floor(i / cols); const c0 = i % cols; const c = r % 2 ? cols - 1 - c0 : c0;
        const bx = x + c * (bw + aw * 1.5); const by = y + r * (bh + h * 0.1);
        box('roundRect', bx, by, bw, bh, nd.t, col(i));
        if (i < n - 1) {
          if (c0 < cols - 1) arrow('rightArrow', r % 2 ? bx - aw * 1.25 : bx + bw + aw * 0.25, by + bh / 2 - aw * 0.5, aw, aw, r % 2 ? 180 : 0);
          else arrow('downArrow', bx + bw / 2 - aw * 0.5, by + bh + h * 0.015, aw, h * 0.07);
        }
      });
      break;
    }
    case 'circleProcess': {
      const d = Math.min(h * 0.7, (w / n) * 0.75); const step = (w - d) / Math.max(1, n - 1);
      if (n > 1) ln(x + d / 2, y + h / 2, x + w - d / 2, y + h / 2).line.width = 4;
      tree.forEach((nd, i) => box('ellipse', x + i * step, y + (h - d) / 2, d, d, nd.t, col(i), { line: { color: '@lt1', width: 3 } }));
      break;
    }
    case 'vChevron': {
      const bh = h / n;
      const cw = Math.min(w * 0.2, bh * 1.1);
      tree.forEach((nd, i) => {
        box('chevron', x + (cw - bh * 0.95) / 2, y + i * bh + (bh * 0.95 - cw) / 2, bh * 0.95, cw, '', col(i), { rot: 90, line: null }).text = null;
        box('rect', x + cw + 8, y + i * bh + bh * 0.08, w - cw - 8, bh * 0.8, [nd.t, ...kidsText(nd)], '@lt1', { line: { color: col(i), width: 1.5 }, textColor: '@tx1', align: 'l', size: fs - 2, b: true, bullets: true });
      });
      break;
    }
    case 'funnel': {
      const fw = Math.min(w * 0.6, h * 1.1);
      const fx = x + (w - fw) / 2;
      const bh = (h * 0.8) / n;
      tree.forEach((nd, i) => { const tw = fw * (1 - i * (0.55 / n)); box('trapezoid', fx + (fw - tw) / 2, y + i * bh, tw, bh - 3, nd.t, col(i), { line: null, adj: { adj: 15000 } }).flipV = true; });
      arrow('downArrow', x + w / 2 - fw * 0.08, y + h * 0.82, fw * 0.16, h * 0.17);
      break;
    }
    case 'timeline': {
      const ly = y + h / 2;
      const l = newShape('straightConnector1', { x, y: ly, w, h: 0 });
      l.line = { color: '@tx1:lm50:lo50', width: 3, tail: 'triangle' };
      l.text = null;
      add(l);
      const step = w / (n + 0.5);
      tree.forEach((nd, i) => {
        const cx = x + step * (i + 0.6);
        box('ellipse', cx - 12, ly - 12, 24, 24, '', col(i), { line: { color: '@lt1', width: 3 } }).text = null;
        const up = i % 2 === 0;
        const tb = box('rect', cx - step * 0.48, up ? ly - h * 0.42 : ly + h * 0.1, step * 0.96, h * 0.32, [nd.t, ...kidsText(nd)], null, { textColor: '@tx1', line: null, anchor: up ? 'b' : 't' });
        tb.fill = null;
      });
      break;
    }
    // ───── 주기형 ─────
    case 'cycle':
    case 'textCycle':
    case 'blockCycle': {
      const r = Math.min(w, h) / 2; const cx = x + w / 2; const cy = y + h / 2;
      const d = Math.min(r * 0.8, (Math.PI * 2 * r * 0.75) / n / 1.15);
      const ring = r - d / 2;
      tree.forEach((nd, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        const shape = kind === 'blockCycle' ? 'roundRect' : 'ellipse';
        box(shape, cx + Math.cos(a) * ring - d / 2, cy + Math.sin(a) * ring - d * (kind === 'blockCycle' ? 0.35 : 0.5), d, d * (kind === 'blockCycle' ? 0.7 : 1), nd.t, kind === 'textCycle' ? null : col(i), { size: Math.max(10, fs - 2), textColor: kind === 'textCycle' ? '@tx1' : undefined, line: kind === 'textCycle' ? null : undefined });
        const am = a + Math.PI / n;
        arrow(kind === 'textCycle' ? 'curvedRightArrow' : 'rightArrow', cx + Math.cos(am) * ring - d * 0.18, cy + Math.sin(am) * ring - d * 0.12, d * 0.36, d * 0.24, Math.round(((am * 180) / Math.PI + 90 + 360) % 360)).fill = { type: 'solid', color: kind === 'textCycle' ? col(i) : '@tx1:lm25:lo75' };
      });
      break;
    }
    case 'segmentedCycle': {
      const r = Math.min(w, h) / 2; const cx = x + w / 2; const cy = y + h / 2;
      tree.forEach((nd, i) => {
        const st = (i * 360) / n - 90; const sw = 360 / n;
        wedge(cx, cy, r, st, sw, col(i));
        const am = ((st + sw / 2) * Math.PI) / 180;
        const tb = box('rect', cx + Math.cos(am) * r * 0.55 - r * 0.32, cy + Math.sin(am) * r * 0.55 - r * 0.16, r * 0.64, r * 0.32, nd.t, null, { line: null, textColor: '@lt1', size: Math.max(10, fs - 2) });
        tb.fill = null;
      });
      break;
    }
    case 'radialCycle':
    case 'converging': {
      const [center, ...rest] = kind === 'radialCycle' ? tree : [tree[tree.length - 1], ...tree.slice(0, -1)];
      const r = Math.min(w, h) / 2; const cx = x + w / 2; const cy = y + h / 2;
      const m = Math.max(1, rest.length);
      const d = Math.min(r * 0.62, (Math.PI * 2 * r * 0.6) / m);
      const cd = r * 0.7;
      rest.forEach((nd, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / m;
        const px = cx + Math.cos(a) * (r - d / 2); const py = cy + Math.sin(a) * (r - d / 2);
        if (kind === 'radialCycle') ln(cx + Math.cos(a) * cd / 2, cy + Math.sin(a) * cd / 2, px - Math.cos(a) * d / 2, py - Math.sin(a) * d / 2);
        else arrow('rightArrow', cx + Math.cos(a) * (cd / 2 + (r - d - cd / 2) / 2) - d * 0.2, cy + Math.sin(a) * (cd / 2 + (r - d - cd / 2) / 2) - d * 0.12, d * 0.4, d * 0.24, Math.round(((a * 180) / Math.PI + 180 + 360) % 360));
        box(kind === 'radialCycle' ? 'ellipse' : 'roundRect', px - d / 2, py - d / 2, d, d * (kind === 'radialCycle' ? 1 : 0.6), nd.t, col(i + 1), { size: Math.max(10, fs - 3) });
      });
      box('ellipse', cx - cd / 2, cy - cd / 2, cd, cd, center?.t ?? '', col(0), { b: true });
      break;
    }
    case 'continuousCycle': {
      const r = Math.min(w, h) / 2; const cx = x + w / 2; const cy = y + h / 2;
      const ring = newShape('donut', { x: cx - r * 0.8, y: cy - r * 0.8, w: r * 1.6, h: r * 1.6 });
      ring.fill = { type: 'solid', color: `${col(0)}:lm40:lo60` }; ring.line = null; ring.text = null; ring.adj = { adj: 12000 };
      add(ring);
      const bw = Math.min(r * 0.9, (Math.PI * 2 * r * 0.8) / n * 0.9); const bh = bw * 0.45;
      tree.forEach((nd, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / n; box('roundRect', cx + Math.cos(a) * r * 0.8 - bw / 2, cy + Math.sin(a) * r * 0.8 - bh / 2, bw, bh, nd.t, col(i), { size: Math.max(10, fs - 2) }); });
      break;
    }
    // ───── 계층 구조형 ─────
    case 'hierarchy':
    case 'nameHierarchy': {
      // 첫 줄이 맨 위, 하위 수준은 들여쓰기 (들여쓰기가 없으면 나머지 줄이 둘째 수준)
      const rootNode = tree.length === 1 ? tree[0] : { t: tree[0].t, kids: [...tree[0].kids, ...tree.slice(1)] };
      orgChart(rootNode, x, y, w, h);
      break;
    }
    case 'hHierarchy': {
      const rootNode = tree.length === 1 ? tree[0] : { t: tree[0].t, kids: [...tree[0].kids, ...tree.slice(1)] };
      const depth = (nd) => 1 + Math.max(0, ...nd.kids.map(depth));
      const leaves = (nd) => (nd.kids.length ? nd.kids.reduce((s, k) => s + leaves(k), 0) : 1);
      const D = depth(rootNode); const L = leaves(rootNode);
      const cw = w / D; const rh = h / L;
      const place = (nd, lvl, row) => {
        const span = leaves(nd);
        const by = y + (row + span / 2) * rh - Math.min(rh * 0.8, 70) / 2;
        const bh = Math.min(rh * 0.8, 70);
        box('roundRect', x + lvl * cw, by, cw * 0.78, bh, nd.t, col(lvl, D), { size: Math.max(10, fs - lvl * 2) });
        let r = row;
        for (const k of nd.kids) {
          const kspan = leaves(k);
          const ky = y + (r + kspan / 2) * rh;
          ln(x + lvl * cw + cw * 0.78, by + bh / 2, x + (lvl + 1) * cw, ky);
          place(k, lvl + 1, r);
          r += kspan;
        }
      };
      place(rootNode, 0, 0);
      break;
    }
    case 'tableHierarchy': {
      const rootNode = tree.length === 1 ? tree[0] : { t: tree[0].t, kids: [...tree[0].kids, ...tree.slice(1)] };
      const depth = (nd) => 1 + Math.max(0, ...nd.kids.map(depth));
      const leaves = (nd) => (nd.kids.length ? nd.kids.reduce((s, k) => s + leaves(k), 0) : 1);
      const D = depth(rootNode); const L = leaves(rootNode);
      const rh = h / D; const lw = w / L;
      const place = (nd, lvl, c0) => { box('rect', x + c0 * lw + 2, y + lvl * rh + 2, leaves(nd) * lw - 4, rh - 4, nd.t, col(lvl, D)); let c = c0; for (const k of nd.kids) { place(k, lvl + 1, c); c += leaves(k); } };
      place(rootNode, 0, 0);
      break;
    }
    // ───── 관계형 ─────
    case 'venn': {
      const d = Math.min(h, w / (n * 0.7 + 0.3)); const total = d * (0.7 * (n - 1) + 1);
      tree.forEach((nd, i) => box('ellipse', x + (w - total) / 2 + i * d * 0.7, y + (h - d) / 2, d, d, nd.t, `${col(i)}:a55`, { textColor: '@tx1', line: null }));
      break;
    }
    case 'targetVenn': {
      const R = Math.min(h, w * 0.55) / 2; const cx = x + R; const cy = y + h / 2;
      tree.forEach((nd, i) => {
        const r = R * (1 - i / n);
        box('ellipse', cx - r, cy - r, r * 2, r * 2, '', col(i), { line: { color: '@lt1', width: 2 } }).text = null;
        const ty = cy - r + (R / n) * 0.5;
        ln(cx, ty, x + R * 2 + w * 0.04, ty);
        box('rect', x + R * 2 + w * 0.05, ty - h / n / 2, w - R * 2 - w * 0.05, h / n, [nd.t, ...kidsText(nd)], null, { textColor: '@tx1', line: null, align: 'l' }).fill = null;
      });
      break;
    }
    case 'stackedVenn': {
      const R = Math.min(h, w) / 2; const cx = x + w / 2; const by = y + h;
      tree.forEach((nd, i) => {
        const r = R * (1 - i / (n + 0.5));
        box('ellipse', cx - r, by - r * 2, r * 2, r * 2, '', `${col(i)}:a70`, { line: { color: '@lt1', width: 2 } }).text = null;
        box('rect', cx - r * 0.6, by - r * 2 + r * 0.12, r * 1.2, R / n * 0.7, nd.t, null, { textColor: '@tx1', line: null, size: fs - 1 }).fill = null;
      });
      break;
    }
    case 'radialVenn': {
      const [center, ...rest] = tree;
      const r = Math.min(w, h) / 2; const cx = x + w / 2; const cy = y + h / 2;
      const cd = r * 0.85; const m = Math.max(1, rest.length); const d = r * 0.75;
      box('ellipse', cx - cd / 2, cy - cd / 2, cd, cd, center.t, `${col(0)}:a70`, { textColor: '@tx1', line: null, b: true });
      rest.forEach((nd, i) => { const a = -Math.PI / 2 + (i * 2 * Math.PI) / m; box('ellipse', cx + Math.cos(a) * (r - d / 2) - d / 2, cy + Math.sin(a) * (r - d / 2) - d / 2, d, d, nd.t, `${col(i + 1)}:a55`, { textColor: '@tx1', line: null, size: fs - 2 }); });
      break;
    }
    case 'counterArrows': {
      const [a1, a2 = { t: '', kids: [] }] = tree;
      box('rightArrow', x + w * 0.05, y + h * 0.05, w * 0.9, h * 0.42, [a1.t, ...kidsText(a1)], col(0), { line: null });
      box('leftArrow', x + w * 0.05, y + h * 0.53, w * 0.9, h * 0.42, [a2.t, ...kidsText(a2)], col(1), { line: null });
      break;
    }
    case 'equation': {
      const parts = n > 1 ? n : 2; const d = Math.min(h * 0.8, w / (parts * 1.5 + 0.3)); const sw = d * 0.5;
      let cxp = x + (w - (d * parts + sw * parts)) / 2;
      tree.forEach((nd, i) => {
        if (i === n - 1 && n > 1) { box('rect', cxp, y + h / 2 - sw * 0.3, sw * 0.8, sw * 0.6, '=', null, { line: null, textColor: '@tx1:lm50:lo50', size: fs * 2 }).fill = null; cxp += sw; }
        box('ellipse', cxp, y + (h - d) / 2, d, d, nd.t, col(i));
        cxp += d;
        if (i < n - 2) { box('plus', cxp + sw * 0.15, y + h / 2 - sw * 0.35, sw * 0.7, sw * 0.7, '', '@tx1:lm50:lo50', { line: null }).text = null; cxp += sw; }
      });
      break;
    }
    case 'balance': {
      const [l1, r1 = { t: '', kids: [] }] = tree;
      const tri = box('triangle', x + w / 2 - h * 0.12, y + h * 0.7, h * 0.24, h * 0.3, '', '@tx1:lm50:lo50', { line: null });
      tri.text = null;
      const beam = box('rect', x + w * 0.05, y + h * 0.64, w * 0.9, h * 0.06, '', '@tx1:lm65:lo35', { line: null, rot: -6 });
      beam.text = null;
      box('roundRect', x + w * 0.06, y + h * 0.2, w * 0.36, h * 0.38, [l1.t, ...kidsText(l1)], col(0), { bullets: true });
      box('roundRect', x + w * 0.58, y + h * 0.12, w * 0.36, h * 0.38, [r1.t, ...kidsText(r1)], col(1), { bullets: true });
      break;
    }
    // ───── 행렬형 ─────
    case 'matrix':
    case 'titledMatrix':
    case 'cycleMatrix': {
      const s = Math.min(w, h); const ox = x + (w - s) / 2; const oy = y + (h - s) / 2;
      const [title, ...cells] = kind === 'titledMatrix' ? tree : [null, ...tree];
      const four = [...cells.slice(0, 4)];
      while (four.length < 4) four.push({ t: '', kids: [] });
      four.forEach((nd, i) => {
        const qx = ox + (i % 2) * s / 2; const qy = oy + Math.floor(i / 2) * s / 2;
        if (kind === 'cycleMatrix') {
          wedge(ox + s / 2, oy + s / 2, s / 2, [180, 270, 90, 0][i], 90, col(i));
          box('rect', qx + s * 0.04, qy + s * 0.12, s * 0.42, s * 0.26, nd.t, null, { line: null, textColor: '@lt1' }).fill = null;
        } else box(kind === 'titledMatrix' ? 'roundRect' : 'rect', qx + 2, qy + 2, s / 2 - 4, s / 2 - 4, [nd.t, ...kidsText(nd)], col(i), { bullets: true });
      });
      if (title) box('ellipse', ox + s * 0.32, oy + s * 0.32, s * 0.36, s * 0.36, title.t, '@lt1', { textColor: '@tx1', line: { color: col(0), width: 2 }, b: true });
      break;
    }
    // ───── 피라미드형 ─────
    case 'pyramid':
    case 'invPyramid': {
      const bh = h / n;
      tree.forEach((nd, i) => {
        const k = kind === 'pyramid' ? i : n - 1 - i;
        const tw = w * ((k + 1) / n);
        const o = box(k === 0 ? 'triangle' : 'trapezoid', x + (w - tw) / 2, y + i * bh, tw, bh, nd.t, col(i));
        if (k > 0) o.adj = { adj: Math.round(((w / n / 2) / Math.min(tw, bh)) * 100000) };
        if (kind === 'invPyramid') o.flipV = true;
      });
      break;
    }
    case 'pyramidList': {
      const tw = Math.min(w * 0.6, h * 1.15);
      const tri = box('triangle', x, y, tw, h, '', `${col(0)}:lm40:lo60`, { line: null });
      tri.text = null;
      const bh = (h * 0.8) / n;
      tree.forEach((nd, i) => box('roundRect', x + tw * 0.5, y + h * 0.15 + i * bh, w - tw * 0.5, bh * 0.85, [nd.t, ...kidsText(nd)], '@lt1', { line: { color: col(i), width: 1.5 }, textColor: '@tx1' }));
      break;
    }
    case 'segPyramid': {
      // 삼각형 4개로 나눈 피라미드 (항목이 더 있으면 줄을 늘림)
      const rows = Math.max(2, Math.ceil(Math.sqrt(n)));
      const th = h / rows; const tw = (Math.min(w, h * 1.15)) / rows; const ox = x + (w - tw * rows) / 2;
      let k = 0;
      for (let r = 0; r < rows; r++) for (let c = 0; c < 2 * r + 1; c++) {
        const nd = tree[k] ?? { t: '', kids: [] };
        const up = c % 2 === 0;
        const t = box('triangle', ox + (rows - 1 - r) * tw / 2 + (c * tw) / 2, y + r * th, tw, th, nd.t, col(k), { size: Math.max(9, fs - 4) });
        if (!up) t.flipV = true;
        k++;
      }
      break;
    }
    // ───── 그림 ─────
    case 'picAccent':
    case 'picTitles': {
      const gap = w * 0.03; const bw = (w - gap * (n - 1)) / n;
      tree.forEach((nd, i) => {
        const bx = x + i * (bw + gap);
        pic(bx, y, bw, h * 0.65);
        box(kind === 'picTitles' ? 'rect' : 'roundRect', bx + (kind === 'picTitles' ? 0 : bw * 0.08), y + h * (kind === 'picTitles' ? 0.65 : 0.58), bw * (kind === 'picTitles' ? 1 : 0.84), h * (kind === 'picTitles' ? 0.35 : 0.3), [nd.t, ...kidsText(nd)], col(i));
      });
      break;
    }
    case 'picGrid': {
      const cols = n <= 2 ? n : n <= 4 ? 2 : 3; const rows = Math.ceil(n / cols);
      const g = 6; const bw = (w - g * (cols - 1)) / cols; const bh = (h - g * (rows - 1)) / rows;
      tree.forEach((nd, i) => {
        const bx = x + (i % cols) * (bw + g); const by = y + Math.floor(i / cols) * (bh + g);
        pic(bx, by, bw, bh);
        box('rect', bx, by + bh * 0.78, bw, bh * 0.22, nd.t, `${col(i)}:a80`, { line: null });
      });
      break;
    }
    case 'picCircles': {
      const gap = w * 0.04; const d = Math.min((w - gap * (n - 1)) / n, h * 0.65);
      const total = d * n + gap * (n - 1);
      tree.forEach((nd, i) => {
        const bx = x + (w - total) / 2 + i * (d + gap);
        pic(bx, y, d, d, 'ellipse');
        box('rect', bx - gap / 2, y + d + 6, d + gap, h - d - 6, [nd.t, ...kidsText(nd)], null, { textColor: '@tx1', line: null, anchor: 't' }).fill = null;
      });
      break;
    }
    default: {
      // 기본 프로세스: 상자 → 화살표 → 상자
      const aw = w * 0.06; const bw = (w - aw * (n - 1) * 1.6) / n; const bh = Math.min(h, bw * 0.65);
      tree.forEach((nd, i) => {
        const bx = x + i * (bw + aw * 1.6);
        box('roundRect', bx, y + (h - bh) / 2, bw, bh, [nd.t, ...kidsText(nd)], col(i), { bullets: true });
        if (i < n - 1) arrow('rightArrow', bx + bw + aw * 0.3, y + h / 2 - aw * 0.6, aw, aw * 1.2);
      });
    }
  }

  /** 조직도: 위에서 아래로, 수준마다 같은 높이, 잎 수만큼 너비 */
  function orgChart(rootNode, ox, oy, ow, oh) {
    const depth = (nd) => 1 + Math.max(0, ...nd.kids.map(depth));
    const leaves = (nd) => (nd.kids.length ? nd.kids.reduce((s, k) => s + leaves(k), 0) : 1);
    const D = depth(rootNode); const L = leaves(rootNode);
    const lh = oh / D; const bh = Math.min(lh * 0.62, 90); const lw = ow / L;
    const place = (nd, lvl, c0) => {
      const span = leaves(nd);
      const cx = ox + (c0 + span / 2) * lw;
      const bw = Math.min(span * lw * 0.9, lw * 0.9 * Math.max(1, Math.min(span, 1.6)), 260);
      const by = oy + lvl * lh;
      const lines = kind === 'nameHierarchy' && nd.kids.length === 0 && nd.t.includes(',') ? nd.t.split(',').map((s) => s.trim()) : [nd.t];
      box('roundRect', cx - bw / 2, by, bw, bh, lines, lvl === 0 ? col(0) : colors === 'multi' ? ACC[(lvl) % 6] : col(lvl, D), { size: Math.max(10, fs - lvl * 2), b: lvl === 0 });
      if (nd.kids.length) {
        const busY = by + bh + (lh - bh) / 2;
        ln(cx, by + bh, cx, busY);
        let c = c0;
        const xs = [];
        for (const k of nd.kids) { const ks = leaves(k); xs.push(ox + (c + ks / 2) * lw); place(k, lvl + 1, c); c += ks; }
        if (xs.length > 1) ln(xs[0], busY, xs[xs.length - 1], busY);
        for (const kx of xs) ln(kx, busY, kx, oy + (lvl + 1) * lh);
      }
    };
    place(rootNode, 0, 0);
  }

  // 스타일 (SmartArt 스타일): 그림자 · 입체 · 윤곽선
  for (const o of out) {
    if (o.type !== 'shape' || !o.fill || o.shape === 'line') continue;
    if (style === 'subtle') o.fx = { shadow: { dist: 1.5, dir: 90, blur: 3, color: '#000000', alpha: 0.25 } };
    else if (style === 'moderate') o.fx = { shadow: { dist: 3, dir: 90, blur: 6, color: '#000000', alpha: 0.35 } };
    else if (style === 'intense') o.fx = { shadow: { dist: 4, dir: 90, blur: 8, color: '#000000', alpha: 0.4 }, bevel: { prst: 'circle', w: 6, h: 6 } };
    else if (style === 'inset') o.fx = { shadow: { inner: true, dist: 2, dir: 225, blur: 6, color: '#000000', alpha: 0.45 } };
    else if (style === 'cartoon') { o.line = { color: '@tx1', width: 2 }; o.fx = { shadow: { dist: 4, dir: 45, blur: 0, color: '#000000', alpha: 0.6 } }; }
    if (style === 'outline' && o.line) o.line = { color: '@lt1', width: 3 };
  }
  const g = uid('g');
  for (const o of out) { o.grp = g; o.smart = kind; }
  if (out[0]) { out[0].smartItems = tree.flatMap(flatText); out[0].smartMulti = colors === 'multi'; out[0].smartColors = colors; out[0].smartStyle = style; }
  return out;
}
