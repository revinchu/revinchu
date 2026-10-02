import { bindGridTouch } from './mobile-grid.js';
import { noteVisible } from './review-state.js';
// 가상 스크롤 그리드: 화면에 보이는 행/열만 그림 (20,000,000행 × 16,384열 지원)
// 틀 고정은 4개 창(TL/TR/BL/BR)으로, 각 창은 시트 좌표계 콘텐츠를 transform 으로 이동시켜 표시.
import { Axis } from './axis.js';
import { visibleAxisIndices } from './axis-window.js';
import { GridAccessibility } from './grid-a11y.js';
import { gridLineWidth, resolveGridBorders } from './grid-lines.js';
import { pictureCropStyle, pictureTransform, pictureEffects, pictureShadowStyle } from './picture.js';
import { sanitizeHtml, setSafeHtml } from './safe-html.js';
import { colToName, cellName, MAX_ROWS, MAX_COLS } from './formula.js';
import { formatValue, formatGeneral } from './format.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { renderChartSvg, chartModelData } from './chart.js';
import { shapeSvg, LINE_KINDS, isShapeLine } from './shapes.js';
import { isSmartArt } from './smartart.js';
import { smartArtSvg } from './smartart-render.js';
import { shapePathHandles } from './shape-edit.js';
import { validationAt } from './validation.js';
import { prepareCond, condFormatAt, ICON_SVG, EMPTY_MATCH_TYPES, ruleRanges, inRule } from './condfmt.js';
import { tableAt, tableCellStyle, tableFilterRange, styleByName } from './tables.js';
import { slicerCssVars } from './slicerstyle.js';
import { THEME } from './stylepresets.js';
import { fontAlias } from './fonts.js';
import { maxLevel, groupsOf } from './outline.js';
import { sparkValues, sparkSvg } from './sparkline.js';

const sparkCache = new WeakMap(); // 스파크라인 항목 → { key, svg }

// 통합 문서 기본 글꼴 (파일의 표준 스타일, 예: 맑은 고딕 9pt). setBaseFont 로 바꾸고 BASE_FONT 를 읽음
export const BASE_FONT = { name: '맑은 고딕', size: 11 };
export function setBaseFont(f) {
  BASE_FONT.name = f?.name || '맑은 고딕';
  BASE_FONT.size = f?.size || 11;
}
export const fontStack = (f) => {
  const name = String(f).replace(/'/g, '');
  const alias = fontAlias(name);
  return `'${name}'${alias ? `, '${alias}'` : ''}, 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif`;
};
const HEAD_H = 20;
const MAX_PX = 15_000_000; // 스크롤 영역 최대 픽셀 (브라우저 한계 회피)
const OVER_R = 12;
const OVER_C = 4;

/** 개체 본문·회전·효과가 렌더 여유 영역과 만나는지. 원본/내보내기에는 관여하지 않는다. */
export function objectIntersectsWindow(o, rect) {
  if (o.hidden) return false;
  const { x, y, w, h } = o;
  if (![x, y, w, h].every(Number.isFinite)) return true; // 불완전한 모델은 기존 렌더가 처리
  const angle = (Number(o.rot) || 0) * Math.PI / 180;
  const width = Math.abs(w), height = Math.max(1, Math.abs(h));
  const rx = (Math.abs(Math.cos(angle)) * width + Math.abs(Math.sin(angle)) * height) / 2;
  const ry = (Math.abs(Math.sin(angle)) * width + Math.abs(Math.cos(angle)) * height) / 2;
  const shadow = o.shadow && typeof o.shadow === 'object' ? o.shadow : {};
  const num = (v, fallback = 0) => Number.isFinite(Number(v)) ? Math.abs(Number(v)) : fallback;
  // 선택 손잡이/차트 옆 단추와 흐림 효과까지 보수적으로 남긴다.
  const pad = 40 + (o.shadow ? Math.max(num(shadow.dx, 3), num(shadow.dy, 3)) + num(shadow.blur, 8) * 2 : 0)
    + num(o.glow?.size) * 2 + num(o.soft) * 2 + num(o.borderW) + num(o.strokeWidth);
  const cx = x + w / 2, cy = y + h / 2;
  return cx + rx + pad >= rect.x1 && cx - rx - pad <= rect.x2 && cy + ry + pad >= rect.y1 && cy - ry - pad <= rect.y2;
}

const measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;

// ───── 시간 표시 막대 기간 나누기 ─────
const TL_EPOCH = Date.UTC(1899, 11, 30);
const tlDate = (v) => new Date(TL_EPOCH + Math.floor(v) * 86400000);
/** 날짜 항목 → 기간 칸 [{ key, short, long, group, items }] (항목이 없는 사이 기간도 칸으로) */
export function timelinePeriods(items, level = 'M') {
  const dated = items.filter((it) => typeof it.v === 'number' && it.v > 0);
  if (!dated.length) return [];
  const keyOf = (d) => {
    const y = d.getUTCFullYear();
    const mo = d.getUTCMonth();
    if (level === 'Y') return y * 10000;
    if (level === 'Q') return y * 10000 + Math.floor(mo / 3) * 100;
    if (level === 'M') return y * 10000 + mo * 100;
    return y * 10000 + mo * 100 + d.getUTCDate();
  };
  const buckets = new Map();
  let lo = Infinity;
  let hi = -Infinity;
  for (const it of dated) {
    const d = tlDate(it.v);
    const k = keyOf(d);
    if (!buckets.has(k)) buckets.set(k, []);
    buckets.get(k).push(it);
    lo = Math.min(lo, it.v);
    hi = Math.max(hi, it.v);
  }
  const out = [];
  const start = tlDate(lo);
  const end = tlDate(hi);
  let cur = new Date(Date.UTC(start.getUTCFullYear(), level === 'Y' ? 0 : level === 'Q' ? Math.floor(start.getUTCMonth() / 3) * 3 : start.getUTCMonth(), level === 'D' ? start.getUTCDate() : 1));
  for (let guard = 0; cur <= end && guard < 4000; guard++) {
    const y = cur.getUTCFullYear();
    const mo = cur.getUTCMonth();
    const k = keyOf(cur);
    const short = level === 'Y' ? `${y}` : level === 'Q' ? `${Math.floor(mo / 3) + 1}분기` : level === 'M' ? `${mo + 1}월` : `${cur.getUTCDate()}`;
    const long = level === 'Y' ? `${y}년` : level === 'Q' ? `${y}년 ${Math.floor(mo / 3) + 1}분기` : level === 'M' ? `${y}년 ${mo + 1}월` : `${y}년 ${mo + 1}월 ${cur.getUTCDate()}일`;
    const group = level === 'Y' ? '' : level === 'D' ? `${y}년 ${mo + 1}월` : `${y}년`;
    out.push({ key: k, short, long, group, items: buckets.get(k) ?? [] });
    if (level === 'Y') cur = new Date(Date.UTC(y + 1, 0, 1));
    else if (level === 'Q') cur = new Date(Date.UTC(y, mo + 3, 1));
    else if (level === 'M') cur = new Date(Date.UTC(y, mo + 1, 1));
    else cur = new Date(cur.getTime() + 86400000);
  }
  return out;
}

/**
 * 글자 세로 보정 (em): 한글 글꼴(맑은 고딕 등)은 아래 여백(descent)이 커서 글자가 줄 상자 위쪽에 붙어 보임.
 * 실제 설치된 글꼴로 한글 · 숫자의 잉크 영역을 재어, 줄 상자(1.2em) 가운데에 오도록 내릴 양을 구함
 */
export function glyphShift(family) {
  if (!measureCtx) return 0;
  measureCtx.font = `100px ${family}`;
  const m = measureCtx.measureText('가나다0123ABC');
  if (!m.fontBoundingBoxAscent) return 0;
  const fa = m.fontBoundingBoxAscent; const fd = m.fontBoundingBoxDescent;
  const L = 120;
  const base = (L - (fa + fd)) / 2 + fa;
  const glyphMid = base + (m.actualBoundingBoxDescent - m.actualBoundingBoxAscent) / 2;
  const dy = (L / 2 - glyphMid) / 100;
  return Math.max(-0.15, Math.min(0.15, Math.round(dy * 1000) / 1000));
}
const measureCache = new Map();
export function fontCss(st = {}) {
  return `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${st.size || BASE_FONT.size}pt ${fontStack(st.font || BASE_FONT.name)}`;
}
export function measureText(text, st) {
  const font = fontCss(st);
  const k = `${font}\u0001${text}`;
  let w = measureCache.get(k);
  if (w === undefined) {
    measureCtx.font = font;
    w = measureCtx.measureText(text).width;
    if (measureCache.size > 20000) measureCache.clear();
    measureCache.set(k, w);
  }
  return w;
}

/** 글꼴이 이 컴퓨터에 없는지 (없으면 대체 글꼴로 그려져 너비가 파일을 만든 엑셀과 다름) */
const missingFont = new Map();
export function fontMissing(name) {
  if (!name || typeof document === 'undefined') return false;
  let v = missingFont.get(name);
  if (v === undefined) {
    const probe = '0123456789,.원 %Wm';
    v = ['monospace', 'serif'].every((fb) => {
      measureCtx.font = `40px ${fb}`;
      const a = measureCtx.measureText(probe).width;
      measureCtx.font = `40px "${name}", ${fb}`;
      return measureCtx.measureText(probe).width === a;
    });
    missingFont.set(name, v);
  }
  return v;
}

/** 열이 좁을 때: 일반 서식은 소수 자릿수를 줄이거나 지수로, 그 밖에는 ### (엑셀과 동일) */
function fitNumber(v, maxW, style) {
  const general = (!style.numFmt || style.numFmt === 'general') && style.decimals === undefined;
  if (general) {
    const abs = Math.abs(v);
    if (abs >= 1e-4 && abs < 1e11 && !Number.isInteger(v)) {
      for (let d = 9; d >= 0; d--) {
        const t = formatGeneral(Number(v.toFixed(d)));
        if (measureText(t, style) <= maxW && t !== '0' && t !== '-0') return t;
      }
    }
    for (let d = 4; d >= 0; d--) {
      const [m, e] = v.toExponential(d).split('e');
      const t = `${m.includes('.') ? m.replace(/\.?0+$/, '') : m}E${e[0] === '-' ? '-' : '+'}${e.replace(/^[+-]/, '').padStart(2, '0')}`;
      if (measureText(t, style) <= maxW) return t;
    }
  }
  return '#'.repeat(Math.max(1, Math.floor(maxW / Math.max(1, measureText('#', style)))));
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

// 엑셀 테두리 선 종류 → CSS (굵기 · 모양)
const BORDER_CSS = {
  thin: [1, 'solid'], hair: [1, 'dotted'], dotted: [1, 'dotted'], dashed: [1, 'dashed'], dashDot: [1, 'dashed'], dashDotDot: [1, 'dashed'],
  medium: [2, 'solid'], mediumDashed: [2, 'dashed'], mediumDashDot: [2, 'dashed'], mediumDashDotDot: [2, 'dashed'], slantDashDot: [2, 'dashed'],
  thick: [3, 'solid'], double: [3, 'double'],
};
/**
 * 병합 셀의 테두리: 엑셀은 병합 영역 가장자리 셀들의 선을 그림 (아래 선 = 맨 아래 행 셀, 오른쪽 선 = 맨 오른쪽 열 셀).
 * 왼쪽 위 셀 서식만 보면 파일의 아래 · 오른쪽 선이나 병합 셀에 그린 테두리가 사라짐.
 */
export function mergeEdgeBorders(wb, si, m, style) {
  const out = { ...style };
  const pick = (k, cells) => {
    for (const [rr, cc] of cells) {
      const s = wb.styleAt(si, rr, cc);
      if (s?.[k]) { out[k] = true; out[`${k}s`] = s[`${k}s`]; out[`${k}c`] = s[`${k}c`]; return; }
    }
  };
  const span = (a, b) => { const o = []; for (let i = a; i <= Math.min(b, a + 200); i++) o.push(i); return o; };
  if (m.r2 > m.r1) pick('bb', span(m.c1, m.c2).map((cc) => [m.r2, cc]));
  if (m.c2 > m.c1) pick('br', span(m.r1, m.r2).map((rr) => [rr, m.c2]));
  if (!out.bt) pick('bt', span(m.c1, m.c2).map((cc) => [m.r1, cc]));
  if (!out.bl) pick('bl', span(m.r1, m.r2).map((rr) => [rr, m.c1]));
  return out;
}

export function borderCss(side, kind, color) {
  const [w, s] = BORDER_CSS[kind ?? 'thin'] ?? BORDER_CSS.thin;
  return `border-${side}:${w}px ${s} ${color ?? '#000'}`;
}

/** 도형 글자: 문단 · 글자 조각(run)의 서식 · 맞춤 · 세로 위치 · 안쪽 여백 (엑셀과 같은 모양) */
export function shapeTextHtml(o) {
  const vj = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[o.valign ?? (o.kind === 'textbox' ? 'top' : 'middle')];
  const pad = o.pad ? o.pad.map((v) => `${v}px`).join(' ') : '4.8px 9.6px';
  // 텍스트 채우기 · 윤곽선 · 효과(그림자 · 네온) — 엑셀 WordArt 서식
  const tfx = [];
  if (o.textShadow) tfx.push('2px 2px 3px rgba(0,0,0,.45)');
  if (o.textGlow) tfx.push(`0 0 4px ${esc(o.textGlow)}`, `0 0 8px ${esc(o.textGlow)}`);
  const textFx = `${o.textOutline ? `-webkit-text-stroke:${o.textOutline.w ?? 0.75}px ${esc(o.textOutline.color ?? '#000')};` : ''}${tfx.length ? `text-shadow:${tfx.join(',')};` : ''}${o.font ? `font-family:${fontStack(o.font)};` : ''}${o.italic ? 'font-style:italic;' : ''}${o.underline ? 'text-decoration:underline;' : ''}`;
  const rotation = [90, 270].includes(Number(o.textRot)) ? Number(o.textRot) : 0;
  const rotateCss = rotation ? `inset:auto;left:50%;top:50%;width:${Math.max(1, o.h)}px;height:${Math.max(1, o.w)}px;transform:translate(-50%,-50%) rotate(${rotation}deg);` : '';
  const base = `${rotateCss}justify-content:${vj};padding:${pad};text-align:${o.align ?? (o.kind === 'textbox' ? 'left' : 'center')};color:${esc(o.color ?? '#000')};font-size:${o.size ?? 11}pt;${o.bold ? 'font-weight:700;' : ''}${o.nowrap ? 'white-space:pre;' : ''}${textFx}`;
  if (!o.paras) return `<div class="sh-text" data-text-fit="${o.textFit === 'shrink' ? 'shrink' : 'none'}" style="${base}"><div class="sh-text-content">${esc(o.text)}</div></div>`;
  const runCss = (r) => [r.b ? 'font-weight:700' : '', r.i ? 'font-style:italic' : '', r.u || r.s ? `text-decoration:${r.u ? 'underline ' : ''}${r.s ? 'line-through' : ''}` : '',
    r.sz ? `font-size:${r.sz}pt` : '', r.color ? `color:${esc(r.color)}` : '', r.font ? `font-family:${fontStack(r.font)}` : ''].filter(Boolean).join(';');
  const paras = o.paras.map((p) => {
    const inner = p.runs.length
      ? p.runs.map((r) => (r.t === '\n' ? '<br>' : `<span style="${runCss(r)}">${esc(r.t)}</span>`)).join('')
      : `<span style="font-size:${p.sz ?? o.size ?? 11}pt">&#8203;</span>`;
    return `<div${p.align ? ` style="text-align:${p.align}"` : ''}>${inner}</div>`;
  }).join('');
  return `<div class="sh-text rich" data-text-fit="${o.textFit === 'shrink' ? 'shrink' : 'none'}" style="${base}"><div class="sh-text-content">${paras}</div></div>`;
}

function shapePointHandlesHtml(shape, editing, zoom) {
  const nodes = shapePathHandles(shape), w = Math.max(1, shape.w), h = Math.max(1, shape.h);
  const xy = (x, y) => [(shape.flip ? 1 - x : x) * w, (shape.flipV ? 1 - y : y) * h];
  const size = 8 / zoom, guides = [], points = [];
  for (const point of nodes) {
    const [x, y] = xy(point.x, point.y);
    if (point.origin) { const [x1, y1] = xy(...point.origin); guides.push(`<line x1="${x1}" y1="${y1}" x2="${x}" y2="${y}"/>`); }
    points.push(`<i class="shape-point ${point.anchor ? 'anchor' : 'control'}${editing.selected === point.id ? ' active' : ''}" data-point="${point.id}" role="button" aria-label="${point.anchor ? '편집점' : '곡선 조절점'}" style="left:${x - size / 2}px;top:${y - size / 2}px;width:${size}px;height:${size}px;"></i>`);
  }
  return `<svg class="shape-point-guides" width="${w}" height="${h}" style="overflow:visible;" stroke="#a22d28" stroke-width="${1 / zoom}" fill="none">${guides.join('')}</svg>${points.join('')}`;
}
function markChartSelection(node, part) {
  if (!part || node.dataset.id !== part.id || !node.classList.contains('chart')) return;
  const svg = node.querySelector('svg'); if (!svg) return;
  svg.querySelectorAll('.chart-element-selection').forEach(n => n.remove());
  if (!['title', 'legend', 'plot', 'dataTable'].includes(part.kind)) return;
  const element = svg.querySelector(`[data-el="${part.kind}"]`); if (!element?.getBBox) return;
  const box = element.getBBox(), transform = element.getCTM(), inverse = svg.getCTM()?.inverse(); if (!transform || !inverse) return;
  const corners = [[box.x, box.y], [box.x + box.width, box.y + box.height]].map(([x, y]) => new DOMPoint(x, y).matrixTransform(transform).matrixTransform(inverse));
  const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  for (const [key, value] of Object.entries({ x: corners[0].x - 2, y: corners[0].y - 2, width: Math.max(1, corners[1].x - corners[0].x) + 4, height: Math.max(1, corners[1].y - corners[0].y) + 4, fill: 'none', stroke: '#1765b3', 'stroke-width': 1, 'stroke-dasharray': '4 2', 'vector-effect': 'non-scaling-stroke', 'pointer-events': 'none', class: 'chart-element-selection' })) rect.setAttribute(key, value);
  svg.append(rect);
}

export function fitShapeText(node, appearance) {
  const box = node.querySelector('.sh-text[data-text-fit="shrink"]');
  if (!box || box.dataset.fitDone === appearance) return;
  const content = box.firstElementChild, css = getComputedStyle(box);
  content.style.zoom = ''; // Refit cached DOM when sheet zoom or device scale changes.
  const width = Math.max(1, box.clientWidth - parseFloat(css.paddingLeft) - parseFloat(css.paddingRight));
  const height = Math.max(1, box.clientHeight - parseFloat(css.paddingTop) - parseFloat(css.paddingBottom));
  const fits = () => content.scrollWidth * Number(content.style.zoom || 1) <= width + 0.5 && content.scrollHeight * Number(content.style.zoom || 1) <= height + 0.5;
  let low = 0.05, high = 1;
  if (!fits()) {
    for (let i = 0; i < 12; i++) { const mid = (low + high) / 2; content.style.zoom = String(mid); if (fits()) low = mid; else high = mid; }
    content.style.zoom = String(low);
  }
  box.dataset.fitDone = appearance;
}

/**
 * host.state() → { wb, si, sel, selKind, active, editing, clip, fillPreview, refs, chartSel,
 *                  showGrid, showFormulas, showHeaders }
 * host.onViewScroll() : 스크롤 후 호출 (편집기 위치 갱신 등)
 */
/** 엑셀 무늬 채우기 18종 → CSS 배경 (무늬 색 fg, 배경색 bg) */
export const PATTERNS = [
  ['gray125', '12.5% 회색'], ['gray0625', '6.25% 회색'], ['lightGray', '25% 회색'], ['mediumGray', '50% 회색'], ['darkGray', '75% 회색'],
  ['lightHorizontal', '가는 가로 줄무늬'], ['lightVertical', '가는 세로 줄무늬'], ['lightDown', '가는 역대각선 줄무늬'], ['lightUp', '가는 대각선 줄무늬'], ['lightGrid', '가는 격자'], ['lightTrellis', '가는 대각선 격자'],
  ['darkHorizontal', '가로 줄무늬'], ['darkVertical', '세로 줄무늬'], ['darkDown', '역대각선 줄무늬'], ['darkUp', '대각선 줄무늬'], ['darkGrid', '격자'], ['darkTrellis', '대각선 격자'],
];
// 엑셀 무늬 채우기의 픽셀 모양 (8×8, '1' = 무늬 색). 부드러운 그라데이션은 번져서 실제보다 진해 보임
const PATTERN_BITS = {
  gray0625: ['10000000', '00000000', '00001000', '00000000', '10000000', '00000000', '00001000', '00000000'],
  gray125: ['10001000', '00000000', '00100010', '00000000', '10001000', '00000000', '00100010', '00000000'],
  lightGray: ['10001000', '00100010', '10001000', '00100010', '10001000', '00100010', '10001000', '00100010'],
  mediumGray: ['10101010', '01010101', '10101010', '01010101', '10101010', '01010101', '10101010', '01010101'],
  darkGray: ['11101110', '10111011', '11101110', '10111011', '11101110', '10111011', '11101110', '10111011'],
  lightHorizontal: ['11111111', '00000000', '00000000', '00000000', '11111111', '00000000', '00000000', '00000000'],
  darkHorizontal: ['11111111', '11111111', '00000000', '00000000', '11111111', '11111111', '00000000', '00000000'],
  lightVertical: ['10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000', '10001000'],
  darkVertical: ['11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100', '11001100'],
  lightDown: ['10001000', '01000100', '00100010', '00010001', '10001000', '01000100', '00100010', '00010001'],
  darkDown: ['11001100', '01100110', '00110011', '10011001', '11001100', '01100110', '00110011', '10011001'],
  lightUp: ['00010001', '00100010', '01000100', '10001000', '00010001', '00100010', '01000100', '10001000'],
  darkUp: ['00110011', '01100110', '11001100', '10011001', '00110011', '01100110', '11001100', '10011001'],
  lightGrid: ['11111111', '10001000', '10001000', '10001000', '11111111', '10001000', '10001000', '10001000'],
  darkGrid: ['11111111', '11111111', '11001100', '11001100', '11111111', '11111111', '11001100', '11001100'],
  lightTrellis: ['10011001', '01100110', '01100110', '10011001', '10011001', '01100110', '01100110', '10011001'],
  darkTrellis: ['11111111', '01100110', '11111111', '10011001', '11111111', '01100110', '11111111', '10011001'],
};
const patternMemo = new Map();
/** 셀 그라데이션 채우기 (xlsx gradientFill): { deg, stops: [[pos, color]] } 선형 또는 { path: true, l, r, t, b, stops } 사각 경로 */
export function gradientCss(g) {
  const stops = (g?.stops ?? []).map(([p, c]) => `${c} ${Math.round(p * 1000) / 10}%`).join(', ');
  if (!stops) return '';
  if (g.path) return `radial-gradient(farthest-corner at ${Math.round(((g.l ?? 0) + (g.r ?? 0)) * 50)}% ${Math.round(((g.t ?? 0) + (g.b ?? 0)) * 50)}%, ${stops})`;
  return `linear-gradient(${Math.round((g.deg ?? 0) + 90)}deg, ${stops})`;
}

export function patternCss(p, fg, bg) {
  const b = bg || 'transparent';
  const bits = PATTERN_BITS[p];
  if (!bits) return b;
  const key = `${p}|${fg}|${b}`;
  let css = patternMemo.get(key);
  if (!css) {
    let d = '';
    bits.forEach((row, y) => { for (let x = 0; x < 8; x++) if (row[x] === '1') d += `M${x} ${y}h1v1h-1z`; });
    const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8' shape-rendering='crispEdges'>${bg ? `<rect width='8' height='8' fill='${bg}'/>` : ''}<path d='${d}' fill='${fg}'/></svg>`;
    // 따옴표 없는 url(): 셀 style="" 속성 안에 들어가므로 ' ( ) 도 인코딩
    css = `url(data:image/svg+xml,${encodeURIComponent(svg).replace(/['()]/g, (ch) => `%${ch.charCodeAt(0).toString(16)}`)}) 0 0 / 8px 8px`;
    patternMemo.set(key, css);
  }
  return css;
}

/** 슬라이서 단추 글자가 잘리면 단추 크기에 맞게 글자를 줄임 (최소 60%) */
function fitSlicerText(root) {
  for (const b of root.querySelectorAll('.sl-item:not([data-fit])')) {
    b.dataset.fit = '1';
    if (b.scrollWidth <= b.clientWidth + 1) continue;
    const base = parseFloat(getComputedStyle(b).fontSize) || 14;
    let size = base;
    while (size > base * 0.6 && b.scrollWidth > b.clientWidth + 1) { size -= 0.5; b.style.fontSize = `${size}px`; }
  }
}

// 선택한 차트 옆 단추 (엑셀: + 차트 요소 · 붓 차트 스타일 · 깔때기 차트 필터)
const CHART_SIDE = '<div class="ch-side">'
  + '<button type="button" class="ch-sb" data-a="elements" title="차트 요소"><svg viewBox="0 0 16 16" fill="none" stroke="#217346" stroke-width="1.6"><path d="M8 2.5v11M2.5 8h11"/></svg></button>'
  + '<button type="button" class="ch-sb" data-a="styles" title="차트 스타일"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2"><path d="M14 2 7.5 9.2"/><path d="M7.4 9.3c-.3-1-1.8-1.3-2.7-.4-1 1-.3 2.2-2.2 3.6 2 .8 4.3.3 5-1 .4-.7.3-1.5-.1-2.2z" fill="#4472c4" stroke="#4472c4"/></svg></button>'
  + '<button type="button" class="ch-sb" data-a="filter" title="차트 필터"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.2"><path d="M2 2.5h12L9.3 8.2v5.3l-2.6-1.3V8.2z"/></svg></button></div>';

export class GridView {
  constructor(host) {
    this.host = host;
    this.wrap = document.getElementById('gridWrap');
    this.scroll = document.getElementById('gridScroll');
    this.sizer = document.getElementById('gridSizer');
    this.viewEl = document.getElementById('gridView');
    this.z = 1;
    this.sx = 0;
    this.sy = 0;
    this.extR = 100;
    this.extC = 26;
    this.hw = 34;
    this.hh = HEAD_H;
    this.cond = null;

    const mk = (cls, parent = this.viewEl) => { const d = document.createElement('div'); d.className = cls; parent.append(d); return d; };
    this.paneBox = mk('panes');
    this.panes = ['tl', 'tr', 'bl', 'br'].map((id) => {
      const el = mk(`pane pane-${id}`, this.paneBox);
      const content = mk('pane-content', el);
      return {
        id, el, content,
        grid: mk('gl', content), cells: mk('cells', content), objects: mk('objects', content), overlay: mk('overlay', content),
        scrollX: id === 'tr' || id === 'br', scrollY: id === 'bl' || id === 'br', win: null, ox: 0, oy: 0,
      };
    });
    this.colHead = mk('col-head');
    this.colHeadFrozen = mk('head-clip', this.colHead);
    this.colHeadScroll = mk('head-clip', this.colHead);
    this.rowHead = mk('row-head');
    this.rowHeadFrozen = mk('head-clip', this.rowHead);
    this.rowHeadScroll = mk('head-clip', this.rowHead);
    this.corner = mk('corner');
    this.corner.title = '모두 선택';
    this.freezeV = mk('freeze-line v');
    this.freezeH = mk('freeze-line h');
    // Visual panes duplicate frozen/merged cells. Expose one logical grid instead;
    // floating charts/slicers remain accessible through each pane's objects layer.
    for (const p of this.panes) for (const el of [p.grid, p.cells, p.overlay]) el.setAttribute('aria-hidden', 'true');
    for (const el of [this.colHead, this.rowHead, this.corner, this.freezeV, this.freezeH]) el.setAttribute('aria-hidden', 'true');
    this.a11y = new GridAccessibility(this, document.getElementById('cellEditor'));

    this.scroll.addEventListener('scroll', () => this.onScroll());
    this.viewEl.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.bindTouch();
    const ro = new ResizeObserver(() => { this._cw = null; this._ch = null; this.layout(); });
    ro.observe(this.wrap);
    ro.observe(this.scroll); // 스크롤 막대가 생기거나 없어질 때
  }

  // ───────────── 좌표 ─────────────
  refreshAxes() {
    const { wb, si } = this.host.state();
    const s = wb.sheets[si];
    this.cols = new Axis(s.defColW ?? DEFAULT_COL_WIDTH, s.colWidths, [s.hiddenCols], MAX_COLS);
    this.rows = new Axis(s.defRowH ?? DEFAULT_ROW_HEIGHT, s.rowHeights, [s.hiddenRows, s.filter?.hidden, ...(s.tables ?? []).map((t) => t.filter?.hidden)], MAX_ROWS);
    this.fr = Math.min(s.freeze?.rows || 0, MAX_ROWS - 1);
    this.fc = Math.min(s.freeze?.cols || 0, MAX_COLS - 1);
    this.frozenW = this.cols.pos(this.fc);
    this.frozenH = this.rows.pos(this.fr);
  }

  // 보이는 영역 크기: DOM 을 바꾼 뒤 읽으면 강제 레이아웃이 일어나므로 크기가 바뀔 때만 다시 잼
  get viewW() { return (this._cw ??= this.scroll.clientWidth) / this.z; }
  get viewH() { return (this._ch ??= this.scroll.clientHeight) / this.z; }

  sheetRect(rg) {
    const c2 = Math.min(rg.c2, MAX_COLS - 1);
    const r2 = Math.min(rg.r2, MAX_ROWS - 1);
    const x = this.cols.pos(rg.c1);
    const y = this.rows.pos(rg.r1);
    return { x, y, w: this.cols.pos(c2 + 1) - x, h: this.rows.pos(r2 + 1) - y };
  }

  /** 범위의 화면 좌표 (보기 영역 기준, 확대 전 단위) */
  screenRect(rg) {
    const r = this.sheetRect(rg);
    const fx = rg.c1 < this.fc;
    const fy = rg.r1 < this.fr;
    return { x: this.hw + (fx ? r.x : r.x - this.sx), y: this.hh + (fy ? r.y : r.y - this.sy), w: r.w, h: r.h };
  }

  /** 범위의 브라우저 좌표 */
  clientRect(rg) {
    const s = this.screenRect(rg);
    const v = this.viewEl.getBoundingClientRect();
    return { left: v.left + s.x * this.z, top: v.top + s.y * this.z, right: v.left + (s.x + s.w) * this.z, bottom: v.top + (s.y + s.h) * this.z, width: s.w * this.z, height: s.h * this.z };
  }

  /** 브라우저 좌표 → 영역/셀 */
  hitTest(clientX, clientY, clampToCells = false) {
    const v = this.viewEl.getBoundingClientRect();
    let x = (clientX - v.left) / this.z;
    let y = (clientY - v.top) / this.z;
    const out = { dx: 0, dy: 0 };
    if (clampToCells) {
      const minX = this.hw + 1;
      const minY = this.hh + 1;
      const maxX = this.viewW - 2;
      const maxY = this.viewH - 2;
      if (x < minX) { out.dx = -1; x = minX; }
      if (y < minY) { out.dy = -1; y = minY; }
      if (x > maxX) { out.dx = 1; x = maxX; }
      if (y > maxY) { out.dy = 1; y = maxY; }
    }
    let zone = y < this.hh && x < this.hw ? 'corner' : y < this.hh ? 'colHeader' : x < this.hw ? 'rowHeader' : 'cell';
    if ((zone === 'rowHeader' && x < (this.olw ?? 0)) || (zone === 'colHeader' && y < (this.olh ?? 0))) zone = 'outline';
    const sheetX = x - this.hw < this.frozenW ? Math.max(0, x - this.hw) : x - this.hw + this.sx;
    const sheetY = y - this.hh < this.frozenH ? Math.max(0, y - this.hh) : y - this.hh + this.sy;
    const c = this.cols.indexAt(sheetX);
    const r = this.rows.indexAt(sheetY);
    let edgeCol = null;
    let edgeRow = null;
    if (zone === 'colHeader') {
      if (this.cols.pos(c + 1) - sheetX < 5) edgeCol = c;
      else if (sheetX - this.cols.pos(c) < 4 && c > 0) edgeCol = this.cols.nextVisible(c - 1, -1);
    }
    if (zone === 'rowHeader') {
      if (this.rows.pos(r + 1) - sheetY < 4) edgeRow = r;
      else if (sheetY - this.rows.pos(r) < 3 && r > 0) edgeRow = this.rows.nextVisible(r - 1, -1);
    }
    return { zone, r, c, x, y, sheetX, sheetY, edgeCol, edgeRow, ...out };
  }

  // ───────────── 스크롤 ─────────────
  maxScroll() {
    const totalW = this.cols.pos(this.extC);
    const totalH = this.rows.pos(this.extR);
    return {
      x: Math.max(0, this.hw + totalW - this.viewW),
      y: Math.max(0, this.hh + totalH - this.viewH),
      pxW: Math.min(MAX_PX, (this.hw + totalW) * this.z),
      pxH: Math.min(MAX_PX, (this.hh + totalH) * this.z),
    };
  }

  updateSizer() {
    const m = this.maxScroll();
    this.sizer.style.width = `${Math.ceil(m.pxW)}px`;
    this.sizer.style.height = `${Math.ceil(m.pxH)}px`;
    return m;
  }

  readScroll() {
    const m = this.maxScroll();
    const maxPxX = Math.max(1, this.scroll.scrollWidth - this.scroll.clientWidth);
    const maxPxY = Math.max(1, this.scroll.scrollHeight - this.scroll.clientHeight);
    this.sx = m.pxW >= MAX_PX ? (this.scroll.scrollLeft / maxPxX) * m.x : this.scroll.scrollLeft / this.z;
    this.sy = m.pxH >= MAX_PX ? (this.scroll.scrollTop / maxPxY) * m.y : this.scroll.scrollTop / this.z;
    this.sx = Math.max(0, Math.min(this.sx, m.x));
    this.sy = Math.max(0, Math.min(this.sy, m.y));
  }

  setScroll(sx, sy) {
    this.ensureExtentFor(this.rows.indexAt(sy + this.frozenH + this.viewH), this.cols.indexAt(sx + this.frozenW + this.viewW));
    const m = this.updateSizer();
    sx = Math.max(0, Math.min(sx, m.x));
    sy = Math.max(0, Math.min(sy, m.y));
    const maxPxX = Math.max(1, this.scroll.scrollWidth - this.scroll.clientWidth);
    const maxPxY = Math.max(1, this.scroll.scrollHeight - this.scroll.clientHeight);
    this.scroll.scrollLeft = m.pxW >= MAX_PX ? (sx / Math.max(1, m.x)) * maxPxX : sx * this.z;
    this.scroll.scrollTop = m.pxH >= MAX_PX ? (sy / Math.max(1, m.y)) * maxPxY : sy * this.z;
    this.sx = sx;
    this.sy = sy;
    this.update();
  }

  scrollBy(dx, dy) { this.setScroll(this.sx + dx, this.sy + dy); }

  /** 시트 확장 범위 (스크롤바가 닿는 곳) */
  ensureExtentFor(r, c) {
    const { wb, si } = this.host.state();
    const used = wb.extent(si);
    const needR = Math.min(MAX_ROWS, Math.max(100, used.rows + 50, r + 50));
    const needC = Math.min(MAX_COLS, Math.max(26, used.cols + 10, c + 10));
    let changed = false;
    if (needR > this.extR) { this.extR = needR; changed = true; }
    if (needC > this.extC) { this.extC = needC; changed = true; }
    return changed;
  }

  resetExtent() {
    this.extR = 100;
    this.extC = 26;
  }

  onScroll() {
    this.readScroll();
    // 끝 가까이 가면 영역 확장 (무한 스크롤)
    const m = this.maxScroll();
    let grown = false;
    if (this.sy > m.y - this.viewH && this.extR < MAX_ROWS) { this.extR = Math.min(MAX_ROWS, this.extR + 500); grown = true; }
    if (this.sx > m.x - this.viewW && this.extC < MAX_COLS) { this.extC = Math.min(MAX_COLS, this.extC + 30); grown = true; }
    if (grown) this.updateSizer();
    this.update();
  }

  onWheel(e) {
    // 슬라이서 항목 목록은 그 안에서 스크롤
    const list = e.target.closest?.('.sl-items');
    if (list && !e.ctrlKey && list.scrollHeight > list.clientHeight) {
      const atTop = list.scrollTop <= 0 && e.deltaY < 0;
      const atEnd = list.scrollTop + list.clientHeight >= list.scrollHeight - 1 && e.deltaY > 0;
      if (!atTop && !atEnd) return;
    }
    if (e.ctrlKey) {
      e.preventDefault();
      this.host.onZoomWheel?.(e.deltaY < 0 ? 10 : -10);
      return;
    }
    e.preventDefault();
    const unit = e.deltaMode === 1 ? 20 : e.deltaMode === 2 ? this.viewH : 1;
    let dx = e.deltaX * unit;
    let dy = e.deltaY * unit;
    if (e.shiftKey && !dx) { dx = dy; dy = 0; }
    if (e.deltaMode === 1) dy = Math.sign(dy) * Math.max(Math.abs(dy), DEFAULT_ROW_HEIGHT * 3);
    this.scroll.scrollLeft += dx * this.z;
    this.scroll.scrollTop += dy * this.z;
  }

  bindTouch() {
    this.unbindTouch = bindGridTouch(this.viewEl, {
      context: () => { const state = this.host.state(); return state.wb.sheets[state.si]; },
      hit: (x, y, clamp) => this.hitTest(x, y, clamp),
      begin: () => this.host.onTouchStart?.(),
      tap: hit => this.host.onTouchSelect?.(hit),
      edit: hit => this.host.onTouchEdit?.(hit),
      end: (phase, cancelled) => { if (phase === 'range' && !cancelled) this.host.onTouchRangeEnd?.(); },
      range: (from, to) => {
        this.host.onTouchRange?.(from, to);
        if (to.dx || to.dy) this.scrollBy((to.dx || 0) * 18, (to.dy || 0) * 18);
      },
      scroll: (dx, dy) => this.scrollBy(dx / this.z, dy / this.z),
      zoom: () => this.z * 100,
      pinch: (pct, previous, current) => {
        const before = this.hitTest(previous.x, previous.y);
        this.host.onTouchZoom?.(pct);
        const after = this.hitTest(current.x, current.y);
        this.setScroll(this.sx + (before.sheetX >= this.frozenW ? before.sheetX - after.sheetX : 0),
          this.sy + (before.sheetY >= this.frozenH ? before.sheetY - after.sheetY : 0));
      },
    });
  }

  setZoom(pct) {
    const keep = { r: this.rows.indexAt(this.sy + this.frozenH), c: this.cols.indexAt(this.sx + this.frozenW) };
    this.z = pct / 100;
    this._cw = null;
    this._ch = null;
    this.viewEl.style.zoom = this.z;
    this.layout(false);
    this.setScroll(this.cols.pos(keep.c) - this.frozenW, this.rows.pos(keep.r) - this.frozenH);
  }

  /** (r,c)가 보이도록 스크롤 */
  ensureVisible(r, c) {
    let { sx, sy } = this;
    const paneW = this.viewW - this.hw - this.frozenW;
    const paneH = this.viewH - this.hh - this.frozenH;
    if (r >= this.fr) {
      const top = this.rows.pos(r) - this.frozenH;
      const bottom = this.rows.pos(r + 1) - this.frozenH;
      if (top < sy) sy = top;
      else if (bottom > sy + paneH) sy = Math.min(top, bottom - paneH);
    }
    if (c >= this.fc) {
      const left = this.cols.pos(c) - this.frozenW;
      const right = this.cols.pos(c + 1) - this.frozenW;
      if (left < sx) sx = left;
      else if (right > sx + paneW) sx = Math.min(left, right - paneW);
    }
    if (sx !== this.sx || sy !== this.sy) {
      this.ensureExtentFor(r + 50, c + 10);
      this.setScroll(sx, sy);
    }
  }

  /** 한 화면에 보이는 행 수 */
  pageRows() {
    return Math.max(1, Math.floor((this.viewH - this.hh - this.frozenH) / DEFAULT_ROW_HEIGHT) - 1);
  }

  firstVisibleRow() { return this.rows.indexAt(this.sy + this.frozenH); }

  // ───────────── 배치 ─────────────
  layout(render = true) {
    if (!this.host.state().wb) return;
    this.viewEl.style.width = `${this.viewW}px`;
    this.viewEl.style.height = `${this.viewH}px`;
    this.refreshAxes();
    this.updateSizer();
    this.readScroll();
    if (render) this.renderAll();
  }

  computeHeaderSize() {
    const { showHeaders, wb, si } = this.host.state();
    if (!showHeaders) { this.hw = 0; this.hh = 0; this.olw = 0; this.olh = 0; return; }
    // 개요(그룹) 기호 자리: 수준마다 14px (수준 단추 1 … 최대+1)
    const ol = wb?.sheets[si]?.outline;
    const rl = maxLevel(ol?.rows);
    const cl = maxLevel(ol?.cols);
    this.olw = rl ? (rl + 1) * 14 + 4 : 0;
    this.olh = cl ? (cl + 1) * 14 + 4 : 0;
    this.hh = HEAD_H + this.olh;
    const bottom = this.rows.indexAt(this.sy + this.frozenH + this.viewH);
    this.hw = Math.max(34, String(bottom + 1).length * 8 + 12) + this.olw;
  }

  paneRects() {
    const { hw, hh, frozenW, frozenH } = this;
    const W = this.viewW;
    const H = this.viewH;
    const fw = Math.min(frozenW, Math.max(0, W - hw));
    const fh = Math.min(frozenH, Math.max(0, H - hh));
    return {
      tl: { x: hw, y: hh, w: fw, h: fh },
      tr: { x: hw + fw, y: hh, w: Math.max(0, W - hw - fw), h: fh },
      bl: { x: hw, y: hh + fh, w: fw, h: Math.max(0, H - hh - fh) },
      br: { x: hw + fw, y: hh + fh, w: Math.max(0, W - hw - fw), h: Math.max(0, H - hh - fh) },
    };
  }

  /** 창에 보이는 시트 영역 */
  visibleRange(p, rect) {
    const x0 = p.scrollX ? this.frozenW + this.sx : 0;
    const y0 = p.scrollY ? this.frozenH + this.sy : 0;
    let c1 = this.cols.indexAt(x0);
    let c2 = this.cols.indexAt(x0 + Math.max(0, rect.w - 0.5));
    let r1 = this.rows.indexAt(y0);
    let r2 = this.rows.indexAt(y0 + Math.max(0, rect.h - 0.5));
    if (p.scrollX) c1 = Math.max(c1, this.fc); else { c1 = 0; c2 = Math.max(0, this.fc - 1); }
    if (p.scrollY) r1 = Math.max(r1, this.fr); else { r1 = 0; r2 = Math.max(0, this.fr - 1); }
    return { r1, r2: Math.max(r1, r2), c1, c2: Math.max(c1, c2), x0, y0 };
  }

  /** 스크롤 후 창 이동 · 필요하면 다시 그림 */
  update(force = false) {
    const prevHw = this.hw;
    const prevHh = this.hh;
    this.computeHeaderSize();
    if (this.hw !== prevHw || this.hh !== prevHh) force = true;
    const rects = this.paneRects();
    for (const p of this.panes) {
      const rect = rects[p.id];
      const visible = rect.w > 0 && rect.h > 0;
      p.el.style.display = visible ? 'block' : 'none';
      if (!visible) { p.win = null; continue; }
      Object.assign(p.el.style, { left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.w}px`, height: `${rect.h}px` });
      const need = this.visibleRange(p, rect);
      const w = p.win;
      if (force || !w || need.r1 < w.r1 || need.r2 > w.r2 || need.c1 < w.c1 || need.c2 > w.c2) {
        p.win = {
          r1: p.scrollY ? Math.max(this.fr, need.r1 - OVER_R) : need.r1,
          r2: p.scrollY ? Math.min(MAX_ROWS - 1, need.r2 + OVER_R) : need.r2,
          c1: p.scrollX ? Math.max(this.fc, need.c1 - OVER_C) : need.c1,
          c2: p.scrollX ? Math.min(MAX_COLS - 1, need.c2 + OVER_C) : need.c2,
        };
        this.renderPane(p);
        this.renderPaneOverlay(p);
      }
      p.content.style.transform = `translate(${p.ox - need.x0}px, ${p.oy - need.y0}px)`;
    }
    this.renderHeaders(rects);
    this.freezeV.style.display = this.fc ? 'block' : 'none';
    this.freezeH.style.display = this.fr ? 'block' : 'none';
    this.freezeV.style.left = `${this.hw + this.frozenW - 1}px`;
    this.freezeH.style.top = `${this.hh + this.frozenH - 1}px`;
    this.host.onViewScroll?.();
    this.a11y?.update();
  }

  renderAll() {
    const { wb, si } = this.host.state();
    this.cond = wb.sheets[si].cond.length ? prepareCond(wb, si) : null;
    this.update(true);
  }

  // ───────────── 셀 그리기 ─────────────
  renderPane(p) {
    const st = this.host.state();
    const { wb, si } = st;
    const sheet = wb.sheets[si];
    const { r1, r2, c1, c2 } = p.win;
    const cols = this.cols;
    const rows = this.rows;
    p.ox = cols.pos(c1);
    p.oy = rows.pos(r1);
    const W = cols.pos(c2 + 1) - p.ox;
    const H = rows.pos(r2 + 1) - p.oy;
    const visRows = visibleAxisIndices(rows, r1, r2);
    const visCols = visibleAxisIndices(cols, c1, c2);

    // SVG 선은 CSS border의 최소 1 CSS px 강제 반올림 없이 화면 픽셀에 맞춥니다.
    const scale = this.z * (globalThis.devicePixelRatio || 1);
    const gridWidth = 1 / scale;
    p.borderSegments = [];
    p.gridCovers = [];
    const g = [];
    if (st.showGrid) {
      for (const c of visCols) g.push(`<rect x="${cols.pos(c + 1) - gridWidth - p.ox}" width="${gridWidth}" height="${H}"/>`);
      for (const r of visRows) g.push(`<rect y="${rows.pos(r + 1) - gridWidth - p.oy}" width="${W}" height="${gridWidth}"/>`);
    }


    const merges = sheet.merges.filter((m) => m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1);
    const inMerge = (r, c) => merges.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
    const html = [];
    // 피벗 +/− 단추가 있는 셀: 글자를 단추 오른쪽으로
    this._tog = null;
    for (const pd of [sheet.pivot, ...(sheet.pivotsExtra ?? [])]) {
      for (const b of pd?.buttons ?? []) {
        if (b.kind !== 'toggle' || b.r < r1 || b.r > r2) continue;
        (this._tog ??= new Set()).add(`${b.r},${b.c}`);
      }
    }
    const hasLine = !!(sheet.allStyle || Object.keys(sheet.colStyles).length || Object.keys(sheet.rowStyles).length);
    const tables = (sheet.tables ?? []).filter((t) => t.r1 <= r2 && t.r2 >= r1 && t.c1 <= c2 && t.c2 >= c1 && styleByName(t.style));
    const inTable = (r, c) => tables.some((t) => r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2);
    const emptyRules = (this.cond ?? []).map((pr) => pr.rule).filter((rl) => EMPTY_MATCH_TYPES.has(rl.type) && ruleRanges(rl).some((g) => g.r1 <= r2 && g.r2 >= r1 && g.c1 <= c2 && g.c2 >= c1));
    const emptyCond = (r, c) => emptyRules.some((rl) => inRule(rl, r, c));
    const spills = wb.spillsOf(si).filter((sp) => sp.r <= r2 && sp.r + sp.h - 1 >= r1 && sp.c <= c2 && sp.c + sp.w - 1 >= c1);
    const inSpill = (r, c) => spills.some((sp) => r >= sp.r && r < sp.r + sp.h && c >= sp.c && c < sp.c + sp.w);
    for (const r of visRows) {
      // 창 왼쪽 밖에서 넘쳐 들어오는 텍스트
      if (c1 > 0) {
        for (let k = c1 - 1; k >= Math.max(0, c1 - 40); k--) {
          const cell = wb.getCell(si, r, k);
          if (!cell?.raw) continue;
          if (!inMerge(r, k) && cols.size(k)) html.push(this.cellHtml(r, k, p, sheet, merges));
          break;
        }
      }
      for (const c of visCols) {
        if (inMerge(r, c)) continue;
        if (!hasLine && !sheet.cells.has(`${r},${c}`) && !(sheet.blocks?.length && wb.blockAt(si, r, c)) && !inTable(r, c) && !emptyCond(r, c) && !inSpill(r, c)) continue;
        html.push(this.cellHtml(r, c, p, sheet, merges));
      }
    }
    for (const m of merges) html.push(this.cellHtml(m.r1, m.c1, p, sheet, merges, m));

    // 스파크라인 (셀 안의 작은 차트)
    for (const g of sheet.sparklines ?? []) {
      for (const it of g.items) {
        if (it.r < r1 || it.r > r2 || it.c < c1 || it.c > c2 || !rows.size(it.r) || !cols.size(it.c)) continue;
        const w = cols.size(it.c);
        const h = rows.size(it.r);
        const key = `${wb.version}:${w}:${h}:${g.type}:${g.color}:${g.negColor}:${g.markers}:${g.high}:${g.low}:${g.first}:${g.last}:${g.negative}:${it.ref}`;
        let hit = sparkCache.get(it);
        if (!hit || hit.key !== key) { hit = { key, svg: sparkSvg(sparkValues(wb, si, it.ref), g, w, h) }; sparkCache.set(it, hit); }
        html.push(`<svg class="spark" style="left:${cols.pos(it.c) - p.ox}px;top:${rows.pos(it.r) - p.oy}px;width:${w}px;height:${h}px">${hit.svg}</svg>`);
      }
    }

    // 필터 단추 (시트 필터 + 표마다)
    const targets = [
      ...(sheet.filter ? [['', sheet.filter]] : []),
      ...(sheet.tables ?? []).filter((t) => t.filter && t.header).map((t) => [t.id, tableFilterRange(t)]),
    ];
    for (const [tid, f] of targets) {
      if (!(f.r1 >= r1 && f.r1 <= r2)) continue;
      for (let c = Math.max(f.c1, c1); c <= Math.min(f.c2, c2); c++) {
        if (!cols.size(c) || !rows.size(f.r1)) continue;
        const active = Array.isArray(f.criteria?.[c]);
        const sort = f.sort?.col === c ? (f.sort.asc ? ' asc' : ' desc') : '';
        html.push(`<div class="fbtn${active ? ' on' : ''}${sort}" data-c="${c}" data-t="${esc(tid)}" title="${active ? '필터 적용됨' : '필터'}" style="left:${cols.pos(c + 1) - 18 - p.ox}px;top:${rows.pos(f.r1 + 1) - 18 - p.oy}px"></div>`);
      }
    }
    // 피벗 테이블 필터 단추 (행 레이블 · 열 레이블 · 보고서 필터)
    [sheet.pivot, ...(sheet.pivotsExtra ?? [])].filter(Boolean).forEach((pd, pi) => {
      const area = pd.area;
      if (pd.classic && area && area.r1 <= r2 && area.r2 >= r1 && area.c1 <= c2 && area.c2 >= c1) {
        const x = cols.pos(area.c1) - p.ox, y = rows.pos(area.r1) - p.oy;
        const width = cols.pos(area.c2 + 1) - cols.pos(area.c1), height = rows.pos(area.r2 + 1) - rows.pos(area.r1);
        html.push(`<div class="pv-classic-frame" aria-hidden="true" style="left:${x}px;top:${y}px;width:${width}px;height:${height}px"></div>`);
        if (!st.readonly) {
          const labels = [['pages', '필터'], ['cols', '열'], ['rows', '행'], ['values', '값']];
          // 드래그 시작 버튼을 덮으면 Chromium의 네이티브 dragstart가 중단될 수 있다.
          const visibleTop = (p.scrollY ? this.frozenH + this.sy : 0) - p.oy, paneHeight = this.paneRects()[p.id].h;
          const zoneTop = y >= visibleTop + 28 ? y - 26 : Math.max(visibleTop, Math.min(y + height + 4, visibleTop + paneHeight - 26));
          html.push(`<div class="pv-classic-zones" style="left:${x}px;top:${zoneTop}px;width:${Math.max(width, 240)}px;height:24px">${labels.map(([kind, label]) => `<div class="pv-classic-zone" data-p="${pi}" data-area="${kind}">${label}에 놓기</div>`).join('')}</div>`);
        }
      }
      if (!pd.buttons) return;
      const filtered = (field) => !!(field && (pd.filters?.[field] || pd.fieldFilters?.[field]));
      for (const b of pd.buttons) {
        if (b.kind === 'toggle') {
          if (b.r < r1 || b.r > r2 || b.c < c1 || b.c > c2 || !cols.size(b.c) || !rows.size(b.r)) continue;
          const ind = (wb.styleAt(si, b.r, b.c).indent ?? 0) * 9;
          html.push(`<div class="pxbtn${b.collapsed ? ' coll' : ''}" data-p="${pi}" data-f="${esc(b.field)}" data-i="${esc(b.item)}" title="${b.collapsed ? '확장' : '축소'}" style="left:${cols.pos(b.c) + 3 + ind - p.ox}px;top:${rows.pos(b.r) + Math.max(0, (rows.size(b.r) - 11) / 2) - p.oy}px"></div>`);
          continue;
        }
        if (b.r < r1 || b.r > r2 || b.c < c1 || b.c > c2 || !cols.size(b.c) || !rows.size(b.r)) continue;
        const fields = b.sigma ? [] : b.field ? [b.field] : b.kind === 'rows' ? pd.rows ?? [] : b.kind === 'cols' ? pd.cols ?? [] : [];
        const on = fields.some(filtered) || (b.kind !== 'page' && fields.some((f) => pd.sort?.[f]));
        if (pd.classic && pd.fieldCaptions !== false && (fields.length || b.sigma) && !st.readonly) {
          const dragFields = b.sigma ? ['Σ 값'] : fields;
          const width = Math.max(20, (cols.size(b.c) - 20) / dragFields.length), height = Math.max(12, rows.size(b.r) - 2);
          dragFields.forEach((field, i) => html.push(`<button type="button" class="pv-classic-field" draggable="true" data-sigma="${!!b.sigma}" data-p="${pi}" data-f="${esc(field)}" data-area="${b.kind === 'page' ? 'pages' : b.kind}" title="${esc(field)}: 끌어서 영역 이동 · 클릭하여 이동 메뉴" aria-label="${esc(field)} 피벗 필드 이동" style="left:${cols.pos(b.c) - p.ox + width * i}px;top:${rows.pos(b.r) - p.oy + 1}px;width:${width}px;height:${height}px">${esc(field)}</button>`));
        }
        html.push(`<div class="fbtn pbtn${on ? ' on' : ''}" data-p="${pi}" data-k="${b.kind}" data-f="${esc(b.field ?? '')}" title="${on ? '필터 적용됨' : '필터'}" style="left:${cols.pos(b.c + 1) - 18 - p.ox}px;top:${rows.pos(b.r + 1) - 18 - p.oy}px"></div>`);
      }
    });
    const lines = resolveGridBorders(p.borderSegments).map((e) => {
      const width = gridLineWidth(e.width, this.z, globalThis.devicePixelRatio || 1, e.pattern);
      const path = (offset, strokeWidth) => {
        const at = e.at === 0 ? width + 1 / scale - offset : e.at - offset;
        if (e.pattern === 'solid' || e.pattern === 'double') return `<rect x="${e.vertical ? at - strokeWidth / 2 : e.start}" y="${e.vertical ? e.start : at - strokeWidth / 2}" width="${e.vertical ? strokeWidth : e.end - e.start}" height="${e.vertical ? e.end - e.start : strokeWidth}" fill="${esc(e.color)}"/>`;
        const d = e.vertical ? `M${at} ${e.start}V${e.end}` : `M${e.start} ${at}H${e.end}`;
        const dash = e.pattern === 'dotted' ? ` stroke-dasharray="${strokeWidth} ${strokeWidth}"` : e.pattern === 'dashed' ? ` stroke-dasharray="${strokeWidth * 3} ${strokeWidth * 2}"` : '';
        return `<path d="${d}" stroke="${esc(e.color)}" stroke-width="${strokeWidth}"${dash}/>`;
      };
      if (e.pattern !== 'double') return path(width / 2, width);
      const stroke = Math.max(1, Math.floor(width * scale / 3)) / scale;
      return path(stroke / 2, stroke) + path(width - stroke / 2, stroke);
    });
    html.push(`<svg class="cell-borders" width="${W}" height="${H}" shape-rendering="crispEdges" fill="none">${lines.join('')}</svg>`);
    const maskId = `grid-mask-${p.id}`;
    const mask = p.gridCovers.length ? `<defs><mask id="${maskId}" maskUnits="userSpaceOnUse" x="0" y="0" width="${W}" height="${H}"><rect width="${W}" height="${H}" fill="white"/>${p.gridCovers.join('')}</mask></defs>` : '';
    setSafeHtml(p.grid, g.length ? `<svg width="${W}" height="${H}" shape-rendering="crispEdges" fill="none">${mask}<g fill="var(--grid-line)"${mask ? ` mask="url(#${maskId})"` : ''}>${g.join('')}</g></svg>` : '');
    setSafeHtml(p.cells, html.join(''));
    this.renderObjects(p);
  }

  cellHtml(r, c, p, sheet, merges, merge = null) {
    const st = this.host.state();
    const { wb, si } = st;
    const cell = wb.getCell(si, r, c);
    let style = wb.styleAt(si, r, c);
    const tbl = tableAt(sheet, r, c);
    if (tbl) {
      // 표 서식은 셀에 직접 지정한 서식 아래에 깔림
      const ts = tableCellStyle(tbl, r, c);
      if (ts) {
        const own = {};
        for (const [k, val] of Object.entries(style)) if (val !== undefined) own[k] = val;
        style = { ...ts, ...own };
      }
    }
    if (merge && (merge.r2 > merge.r1 || merge.c2 > merge.c1)) style = mergeEdgeBorders(wb, si, merge, style);
    const v = wb.getValue(si, r, c);
    const x = this.cols.pos(c);
    const y = this.rows.pos(r);
    const w = merge ? this.cols.pos(merge.c2 + 1) - x : this.cols.size(c);
    const h = merge ? this.rows.pos(merge.r2 + 1) - y : this.rows.size(r);
    if (!w || !h) return '';
    let bar = null;
    let icon = null;
    let hideValue = false;
    if (this.cond) {
      const cf = condFormatAt(this.cond, wb, si, r, c, v);
      if (cf.style) style = { ...style, ...cf.style };
      ({ bar, icon, hideValue } = cf);
    }
    let text;
    let align;
    let fmtColor = null;
    let image = null;
    if (st.showFormulas && cell?.formula) { text = cell.raw; align = 'left'; } else ({ text, align, color: fmtColor, image } = formatValue(v, style, wb.date1904));
    // 확인란 칸: 논리값(또는 빈 칸)을 체크 상자로
    let checkbox = null;
    if (style.checkbox && (typeof v === 'boolean' || v === null || v === undefined || v === '') && !(st.showFormulas && cell?.formula)) { checkbox = v === true; text = ''; }
    if (v === 0 && wb.sheets[si].noZeros && !(st.showFormulas && cell?.formula)) text = ''; // 0 값이 있는 셀에 0 표시 안 함 (엑셀 옵션 › 고급)
    // 선택 영역의 가운데로 (centerContinuous): 오른쪽의 빈 같은 맞춤 칸들까지 합친 너비의 가운데
    let across = 0;
    if (style.align === 'centerContinuous' && text && !merge) {
      for (let cc = c + 1; cc < c + 64; cc++) {
        if (wb.getCell(si, r, cc)?.raw || wb.styleAt(si, r, cc)?.align !== 'centerContinuous') break;
        across += this.cols.size(cc);
      }
    }
    const eff = style.align === 'centerContinuous' ? 'center' : style.align && style.align !== 'general' ? style.align : align;
    const css = [`left:${x - p.ox}px`, `top:${y - p.oy}px`, `width:${w}px`, `height:${h}px`];
    if (style.bold) css.push('font-weight:700');
    if (style.italic) css.push('font-style:italic');
    if (style.underline || style.strike) css.push(`text-decoration:${style.underline ? 'underline ' : ''}${style.strike ? 'line-through' : ''}`);
    if (st.valueHighlight && text) css.push(`color:${cell?.formula ? '#008000' : typeof v === 'number' ? '#0000ff' : '#000000'}`); // LibreOffice 값 강조 (Ctrl+F8)
    else if (fmtColor || style.color) css.push(`color:${fmtColor || style.color}`);
    if (style.font) css.push(`font-family:${fontStack(style.font)}`);
    if (style.size) css.push(`font-size:${style.size}pt`);
    if (eff !== 'left') css.push(`justify-content:${eff === 'center' ? 'center' : 'flex-end'};text-align:${eff}`);
    if (style.valign === 'top') css.push('align-items:flex-start');
    else if (style.valign === 'middle') css.push('align-items:center');
    if (this._tog?.has(`${r},${c}`)) css.push(`padding-left:${3 + (style.indent ?? 0) * 9 + 14}px`);
    else if (style.indent) css.push(`padding-${eff === 'right' ? 'right' : 'left'}:${3 + style.indent * 9}px`);
    const bg = style.fill || (merge ? '#fff' : null);
    if (bar) {
      // 그라데이션(엑셀 기본) 또는 단색, 음수 막대는 오른쪽에서 왼쪽으로
      const img = bar.gradient
        ? `linear-gradient(90deg, ${bar.color}, ${bar.color}33)`
        : `linear-gradient(${bar.color}, ${bar.color})`;
      css.push(`background:${img} no-repeat ${bar.neg ? '100%' : '0'} 50% / ${bar.pct}% 72%${bg ? `, ${bg}` : ''};background-clip:padding-box`);
    }
    else if (style.gradient) css.push(`background:${gradientCss(style.gradient)}`);
    else if (style.pattern) css.push(`background:${patternCss(style.pattern, style.patternColor ?? '#000000', bg)}`);
    else if (bg) css.push(`background-color:${bg}`);
    // 채우기는 네 면의 눈금선을 마스크로 지웁니다. 명시적 셀 테두리는 별도 층에 남습니다.
    const covers = !bar && !!(bg || style.pattern || style.gradient);
    if (covers) {
      css.push('background-clip:border-box');
      const inset = 2 / (this.z * (globalThis.devicePixelRatio || 1));
      p.gridCovers.push(`<rect x="${x - p.ox - inset}" y="${y - p.oy - inset}" width="${w + inset}" height="${h + inset}" fill="black"/>`);
    }
    // 셀 배경과 독립된 선 층: 맞닿은 두 셀·병합 셀의 경계는 굵기 우선으로 한 번만 그립니다.
    const edge = (key, vertical, at, start, end) => {
      if (!style[key]) return;
      const [width, pattern] = BORDER_CSS[style[`${key}s`] ?? 'thin'] ?? BORDER_CSS.thin;
      p.borderSegments.push({ vertical, at, start, end, width, pattern, color: style[`${key}c`] ?? '#000' });
    };
    edge('bt', false, y - p.oy, x - p.ox, x + w - p.ox);
    edge('bb', false, y + h - p.oy, x - p.ox, x + w - p.ox);
    edge('bl', true, x - p.ox, y - p.oy, y + h - p.oy);
    edge('br', true, x + w - p.ox, y - p.oy, y + h - p.oy);
    const cls = [];
    // 숫자는 자동 줄 바꿈이어도 한 줄 (엑셀: 들어가지 않으면 ###)
    const wrap = style.wrap && typeof v !== 'number';
    if (wrap) cls.push('wrap');
    else if (text && typeof v !== 'number' && !merge && (eff === 'left' || eff === 'center' || eff === 'right')) {
      // 넘친 글자: 왼쪽 맞춤은 오른쪽 빈 칸으로, 오른쪽 맞춤은 왼쪽 빈 칸으로, 가운데는 양쪽이 모두 비었을 때 양쪽으로 (엑셀과 같음)
      const emptyAt = (cc) => cc < 0 || (!wb.getCell(si, r, cc)?.raw && !merges.some((m) => r >= m.r1 && r <= m.r2 && cc >= m.c1 && cc <= m.c2));
      const ok = eff === 'left' ? emptyAt(c + 1) : eff === 'right' ? c > 0 && emptyAt(c - 1) : c > 0 && emptyAt(c - 1) && emptyAt(c + 1);
      if (ok) cls.push('ovf');
    }
    if (cell?.comment) cls.push('cm');
    // 아이콘 집합: 아이콘은 왼쪽 끝에 고정하고 글자는 남은 너비 안에 (엑셀과 같음)
    const ICON_W = 18;
    if (icon) {
      css.push(`padding-left:${(style.indent && eff !== 'right' ? 3 + style.indent * 9 : 3) + ICON_W}px`);
      const i = cls.indexOf('ovf');
      if (i >= 0) cls.splice(i, 1);
    }
    const room = w - 6 - (icon ? ICON_W : 0);
    if (typeof v === 'number' && text && !st.showFormulas) {
      const tw = measureText(text, style);
      if (tw > room) {
        // 파일의 글꼴(돋움 등)이 없어 대체 글꼴이 더 넓어서 넘치는 경우: ### 대신 조금 작게 (엑셀 화면에서는 들어맞음)
        if (tw <= room * 1.3 && fontMissing(style.font || BASE_FONT.name)) css.push(`font-size:${(((style.size ?? BASE_FONT.size) * room) / tw).toFixed(2)}pt`);
        else text = fitNumber(v, room, style);
      }
    }
    // 셀에 맞춤(축소): 글자가 칸보다 넓으면 글꼴을 줄여서 한 줄에 맞춤
    let spanCss = '';
    if (style.shrink && text && !style.wrap && typeof v !== 'number') {
      const tw = measureText(text, style);
      if (tw > room) css.push(`font-size:${(((style.size ?? BASE_FONT.size) * room) / tw).toFixed(2)}pt`);
    }
    // 텍스트 방향: 각도(시계 반대 방향 +), 255 = 세로 쓰기
    let rotBox = null;
    if (style.rotate === 255) spanCss = ' style="writing-mode:vertical-rl;text-orientation:upright;letter-spacing:-1px"';
    else if (style.rotate && text) {
      // 회전한 글자가 차지하는 사각형을 맞춤 위치에 두고, 그 가운데에서 글자를 돌림 (엑셀과 같은 자리)
      const tw = measureText(text, style);
      const lh = ((style.size ?? BASE_FONT.size) * 4 / 3) * 1.2;
      const a = (Math.abs(style.rotate) * Math.PI) / 180;
      const bw = tw * Math.cos(a) + lh * Math.sin(a);
      const bh = tw * Math.sin(a) + lh * Math.cos(a);
      rotBox = `<span style="display:inline-block;position:relative;flex:none;width:${bw.toFixed(1)}px;height:${bh.toFixed(1)}px"><span style="position:absolute;left:50%;top:50%;width:${Math.ceil(tw)}px;white-space:nowrap;transform:translate(-50%,-50%) rotate(${-style.rotate}deg)">${hideValue ? '' : esc(text)}</span></span>`;
      const i = cls.indexOf('ovf');
      if (i >= 0) cls.splice(i, 1);
    }
    const comment = cell?.comment ? ` data-cm="${esc(cell.comment)}"` : '';
    // 대각선 테두리 (↘ dd, ↗ du)
    const dline = (on, sty, col, x1, y1, x2, y2) => {
      if (!on) return '';
      const [bw, kind] = { medium: [2, ''], thick: [3, ''], double: [3, ''], dashed: [1, '4,3'], dotted: [1, '1,2'], hair: [1, '1,1'], dashDot: [1, '6,2,1,2'], dashDotDot: [1, '6,2,1,2,1,2'], mediumDashed: [2, '6,3'], mediumDashDot: [2, '8,3,2,3'], mediumDashDotDot: [2, '8,3,2,3,2,3'], slantDashDot: [2, '8,2,2,2'] }[sty] ?? [1, ''];
      return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${esc(col ?? '#000')}" stroke-width="${bw}"${kind ? ` stroke-dasharray="${kind}"` : ''}/>`;
    };
    const diagHtml = style.dd || style.du ? `<svg class="cdiag" width="${w}" height="${h}">${dline(style.dd, style.dds, style.ddc, 0, 0, w, h)}${dline(style.du, style.dus, style.duc, 0, h, w, 0)}</svg>` : '';
    const iconHtml = icon ? `<i class="cf-icon">${ICON_SVG[icon] ?? ''}</i>` : '';
    if (icon) cls.push('has-icon');
    if (image) {
      // 셀 안 그림: 0 셀에 맞춤(비율 유지), 1 셀 채우기, 2 원래 크기, 3 높이 · 너비 지정
      const fit = image.sizing === 1 ? 'fill' : image.sizing === 2 ? 'none' : 'contain';
      const size = image.sizing === 3 ? `width:${image.w ? `${image.w}px` : 'auto'};height:${image.h ? `${image.h}px` : 'auto'};` : 'width:100%;height:100%;';
      const img = `<img class="cimg" src="${esc(image.src)}" alt="${esc(image.alt ?? '')}" title="${esc(image.alt ?? '')}" draggable="false" loading="lazy" style="${size}object-fit:${fit}">`;
      const i = cls.indexOf('ovf');
      if (i >= 0) cls.splice(i, 1);
      return `<div class="c cimg-cell${cls.length ? ` ${cls.join(' ')}` : ''}" data-r="${r}" data-c="${c}" style="${css.join(';')}"${comment}>${img}</div>`;
    }
    if (across && !rotBox && !wrap) {
      // 칸 밖으로 이어진 영역 가운데에 글자 (다음 칸들은 비어 있음)
      const i = cls.indexOf('ovf');
      if (i < 0) cls.push('ovf');
      spanCss = ` style="position:absolute;left:0;top:0;bottom:0;width:${w + across}px;display:flex;align-items:inherit;justify-content:center;white-space:nowrap"`;
    }
    if (checkbox !== null) {
      const col = style.color || '#217346';
      const box = `<svg class="cbx" viewBox="0 0 16 16" width="15" height="15"><rect x="1" y="1" width="14" height="14" rx="2.5" fill="${checkbox ? col : '#fff'}" stroke="${checkbox ? col : '#8a8a8a'}" stroke-width="1.3"/>${checkbox ? '<path d="M4.2 8.3l2.5 2.5 5-5.3" fill="none" stroke="#fff" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>' : ''}</svg>`;
      return `<div class="c cbx-cell${cls.length ? ` ${cls.join(' ')}` : ''}" data-r="${r}" data-c="${c}" style="${css.join(';')}"${comment}>${diagHtml}${box}</div>`;
    }
    return `<div class="c${cls.length ? ` ${cls.join(' ')}` : ''}" data-r="${r}" data-c="${c}" style="${css.join(';')}"${comment}>${diagHtml}${iconHtml}${rotBox ?? `<span${spanCss}>${hideValue ? '' : esc(text)}</span>`}</div>`;
  }

  /**
   * 연결된 그림 (엑셀 카메라 · 그림으로 연결하여 붙여넣기): 원본 범위를 값 · 서식 그대로 그려 개체 크기에 맞춤 (원본이 바뀌면 같이 바뀜)
   * o.linked = { sheet, r1, c1, r2, c2 }
   */
  linkedHtml(o) {
    const { wb } = this.host.state();
    const L = o.linked;
    const s = wb.sheetIndexByName(L.sheet);
    if (s < 0) return '<div class="lnk-ref">#REF!</div>';
    const r2 = Math.min(L.r2, L.r1 + 199);
    const c2 = Math.min(L.c2, L.c1 + 49);
    const ws = [];
    for (let c = L.c1; c <= c2; c++) ws.push(wb.colWidth(s, c));
    const hs = [];
    for (let r = L.r1; r <= r2; r++) hs.push(wb.rowHeight(s, r));
    const W = ws.reduce((a, b) => a + b, 0) || 1;
    const H = hs.reduce((a, b) => a + b, 0) || 1;
    const merges = (wb.sheets[s].merges ?? []).filter((m) => m.r1 >= L.r1 && m.c1 >= L.c1 && m.r1 <= r2 && m.c1 <= c2);
    const covered = new Set();
    for (const m of merges) for (let r = m.r1; r <= Math.min(m.r2, r2); r++) for (let c = m.c1; c <= Math.min(m.c2, c2); c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
    const rows = [];
    for (let r = L.r1; r <= r2; r++) {
      const tds = [];
      for (let c = L.c1; c <= c2; c++) {
        if (covered.has(`${r},${c}`)) continue;
        const m = merges.find((x) => x.r1 === r && x.c1 === c);
        const v = wb.getValue(s, r, c);
        const st = wb.styleAt(s, r, c) ?? {};
        let text = '';
        let color = null;
        if (v !== null && v !== undefined && v !== '') {
          if (typeof v === 'object' && v.code) text = v.code;
          else { const f = formatValue(v, st, wb.date1904); text = f.text; color = f.color; }
        }
        const css = [];
        if (st.bold) css.push('font-weight:700');
        if (st.italic) css.push('font-style:italic');
        if (st.underline || st.strike) css.push(`text-decoration:${st.underline ? 'underline ' : ''}${st.strike ? 'line-through' : ''}`);
        if (color || st.color) css.push(`color:${color || st.color}`);
        if (st.font) css.push(`font-family:${fontStack(st.font)}`);
        if (st.size) css.push(`font-size:${st.size}pt`);
        const al = st.align === 'centerContinuous' ? 'center' : st.align && st.align !== 'general' ? st.align : (typeof v === 'number' ? 'right' : typeof v === 'boolean' || (v && v.code) ? 'center' : 'left');
        css.push(`text-align:${al}`);
        css.push(`vertical-align:${st.valign === 'top' ? 'top' : st.valign === 'middle' ? 'middle' : 'bottom'}`);
        if (st.gradient) css.push(`background:${gradientCss(st.gradient)}`);
        else if (st.pattern) css.push(`background:${patternCss(st.pattern, st.patternColor ?? '#000000', st.fill)}`);
        else if (st.fill) css.push(`background-color:${st.fill}`);
        if (st.bt) css.push(borderCss('top', st.bts, st.btc));
        if (st.bb) css.push(borderCss('bottom', st.bbs, st.bbc));
        if (st.bl) css.push(borderCss('left', st.bls, st.blc));
        if (st.br) css.push(borderCss('right', st.brs, st.brc));
        if (st.wrap) css.push('white-space:pre-wrap');
        const span = m ? `${m.c2 > m.c1 ? ` colspan="${Math.min(m.c2, c2) - m.c1 + 1}"` : ''}${m.r2 > m.r1 ? ` rowspan="${Math.min(m.r2, r2) - m.r1 + 1}"` : ''}` : '';
        tds.push(`<td${span} style="${css.join(';')}">${esc(text)}</td>`);
      }
      rows.push(`<tr style="height:${hs[r - L.r1]}px">${tds.join('')}</tr>`);
    }
    const cols = ws.map((w) => `<col style="width:${w}px">`).join('');
    return `<div class="lnk-box"><table class="lnk-tbl" style="width:${W}px;height:${H}px;font:${BASE_FONT.size}pt ${fontStack(BASE_FONT.name)};transform:scale(${o.w / W},${o.h / H})"><colgroup>${cols}</colgroup>${rows.join('')}</table></div>`;
  }

  /** 그림 개체: 차트 · 그림 · 도형 */
  renderObjects(p) {
    const st = this.host.state();
    const { wb, si } = st;
    const sheet = wb.sheets[si];
    const images = sheet.images ?? [];
    const shapes = st.shapePreview ? [...(sheet.shapes ?? []).filter((o) => o.id !== st.shapePreview.id), st.shapePreview] : sheet.shapes ?? [];
    const slicers = sheet.slicers ?? [];
    if (!sheet.charts.length && !images.length && !shapes.length && !slicers.length && !sheet.noteVisibility) { p.objects.replaceChildren(); p.objHtml = null; return; }
    if (!p.win) return;
    // 가시 화면보다 넓은 셀 렌더 창을 사용한다. 그 안의 작은 스크롤에서는
    // renderPane가 재실행되지 않으므로 화면만 기준으로 버리면 개체가 늦게 나타난다.
    const windowRect = {
      x1: Math.max(p.scrollX ? this.frozenW : 0, this.cols.pos(p.win.c1)),
      y1: Math.max(p.scrollY ? this.frozenH : 0, this.rows.pos(p.win.r1)),
      x2: Math.min(p.scrollX ? Infinity : this.frozenW, this.cols.pos(p.win.c2 + 1)),
      y2: Math.min(p.scrollY ? Infinity : this.frozenH, this.rows.pos(p.win.r2 + 1)),
    };
    const appearance = `${THEME.key}|${BASE_FONT.name}|${BASE_FONT.size}|${this.z}|${globalThis.devicePixelRatio || 1}|${!!st.readonly}|${!!st.viewOnly}`;
    const previous = this._objectRenderState;
    if (!this._objectRenderCache || previous?.wb !== wb || previous.si !== si || previous.sheet !== sheet || previous.version !== wb.version || previous.appearance !== appearance) {
      this._objectRenderCache = new Map();
      this._objectRenderState = { wb, si, sheet, version: wb.version, appearance };
    }
    const content = (o, kind, build) => {
      // 개체를 직접 움직이는 미리보기에서도 위치·크기·서식 변경을 놓치지 않는다.
      // 연결 그림은 보존용 원본 base64를 사용하지 않으므로 키에서도 제외한다.
      const { src, emf, ...linkedOptions } = kind === 'linked' ? o : {};
      const key = JSON.stringify(kind === 'linked' ? linkedOptions : o);
      const hit = this._objectRenderCache.get(o);
      if (hit?.kind === kind && hit.key === key) return hit.html;
      const html = build();
      // 긴 시트를 두루 스크롤해도 방문한 모든 SVG를 영구 보관하지 않는다.
      if (!hit && this._objectRenderCache.size >= 128) this._objectRenderCache.delete(this._objectRenderCache.keys().next().value);
      this._objectRenderCache.set(o, { kind, key, html });
      return html;
    };
    const html = [];
    const handles = '<i class="ch-h nw"></i><i class="ch-h ne"></i><i class="ch-h sw"></i><i class="ch-h se"></i>';
    const box = (o, cls, inner, extraCss = '') => {
      const h = Math.max(o.h, cls.includes('line') ? 1 : 0);
      const selected = st.chartSel === o.id || !!st.objMulti?.has(o.id);
      html.push(`<div class="obj ${cls}${selected ? ' sel' : ''}${o.macro ? ' macro' : ''}" data-id="${esc(o.id)}" style="left:${o.x - p.ox}px;top:${o.y - p.oy}px;width:${o.w}px;height:${h}px;${extraCss}">${inner}${selected ? (st.shapeEdit?.id === o.id ? shapePointHandlesHtml(o, st.shapeEdit, this.z) : handles) : ''}</div>`);
    };
    // 엑셀처럼 그림 → 도형 → 차트 순서가 아니라 저장된 순서(z)대로 겹침
    const all = [
      ...sheet.charts.map((o) => ['charts', st.chartPreview?.id === o.id ? st.chartPreview : o]), ...images.map((o) => ['images', o]), ...shapes.map((o) => ['shapes', o]),
      ...slicers.map((o) => ['slicers', o]),
    ].filter(([, o]) => objectIntersectsWindow(o, windowRect)).sort((a, b) => (a[1].z ?? 0) - (b[1].z ?? 0));
    for (const [prop, o] of all) {
      if (prop === 'charts') {
        box(o, 'chart', content(o, 'chart', () => this.chartSvg(o) + this.pivotChartButtons(o)) + (st.chartSel === o.id && !st.objMulti?.size && !st.viewOnly ? CHART_SIDE : ''));
      }
      else if (prop === 'slicers') box(o, o.timeline ? 'slicer timeline' : 'slicer', content(o, 'slicer', () => this.slicerHtml(o)), slicerCssVars(o));
      else if (prop === 'images' && o.linked) box(o, 'pic linked', content(o, 'linked', () => this.linkedHtml(o)), o.rot ? `transform:rotate(${o.rot}deg)` : '');
      else if (prop === 'images') {
        // 그림 스타일: 테두리 · 둥근 모서리 · 그림자 · 회전 · 투명도
        const ic = [o.border ? `border:${o.borderW ?? 2}px solid ${esc(o.border)}` : '', o.radius ? `border-radius:${pictureEffects(o).radius}px` : '', o.shadow ? `box-shadow:${pictureShadowStyle(o)}` : '', o.opacity !== undefined ? `opacity:${pictureEffects(o).opacity}` : ''].filter(Boolean).join(';');
        // 자르기(crop: 위 · 아래 · 왼쪽 · 오른쪽 비율): 원본을 키워 보이는 부분만 틀에 맞춤
        const cr = o.crop;
        const img = cr
          ? (() => {
            const css = pictureCropStyle(cr);
            return `<div style="position:absolute;inset:0;overflow:hidden;${ic}"><img src="${esc(o.src)}" alt="${esc(o.alt ?? o.name ?? '')}" draggable="false" style="position:absolute;max-width:none;left:${css.left};top:${css.top};width:${css.width};height:${css.height}"></div>`;
          })()
          : `<img src="${esc(o.src)}" alt="${esc(o.alt ?? o.name ?? '')}" draggable="false"${ic ? ` style="${ic};box-sizing:border-box"` : ''}>`;
        box(o, 'pic', img, pictureTransform(o) ? `transform:${pictureTransform(o)}` : '');
      }
      else {
        const isLine = isShapeLine(o);
        const inner = content(o, 'shape', () => isSmartArt(o) ? smartArtSvg(o) : shapeSvg(o) + ((o.text || o.paras) && !isLine ? shapeTextHtml(o) : ''));
        box(o, `shape ${isLine ? 'line' : ''}${o.draft ? ' drawing-preview' : ''}`, inner, o.rot ? `transform:rotate(${o.rot}deg)` : '');
      }
    }
    // 표시한 메모는 현재 렌더 창의 셀만 조회한다(전체 셀 저장소를 매번 훑지 않는다).
    if (sheet.noteVisibility) {
      const noteRows = visibleAxisIndices(this.rows, p.win.r1, p.win.r2);
      const noteCols = visibleAxisIndices(this.cols, p.win.c1, p.win.c2);
      for (const c of noteCols) for (const r of noteRows) {
        if (!noteVisible(sheet.noteVisibility, r, c)) continue;
        const text = sheet.cells.getRC(r, c)?.comment; if (!text) continue;
        html.push(`<div class="cell-note-visible" data-note-r="${r}" data-note-c="${c}" style="left:${this.cols.pos(c + 1) + 8 - p.ox}px;top:${this.rows.pos(r) - p.oy}px"><b>${esc(cellName(r, c))}</b><div>${esc(text)}</div></div>`);
      }
    }
    // 바뀐 개체만 다시 만듦 (슬라이서 · 차트가 많아도 클릭마다 전부 다시 그리지 않게)
    p.objHtml ??= new Map();
    const next = new Map();
    const nodes = html.map((h) => {
      // 선택 강조도 캐시에 반영하고, 외부 값으로 style 태그/선택자를 생성하지 않는다.
      const key = h + '\0' + JSON.stringify(st.chartPart ?? null);
      let node = p.objHtml.get(key);
      if (!node) {
        node = sanitizeHtml(h).firstElementChild;
        const part = st.chartPart;
        if (node?.classList.contains('chart') && node.dataset.id === st.chartSel && part?.id === st.chartSel) {
          for (const item of node.querySelectorAll('svg [data-s], svg [data-el]')) {
            const selected = part.kind === 'series' ? item.dataset.s === String(part.s)
              : part.kind === 'point' ? item.dataset.s === String(part.s) && item.dataset.p === String(part.p)
                : item.dataset.el === part.kind && (part.kind !== 'label' || item.dataset.s === String(part.s));
            if (!selected) continue;
            item.style.filter = 'drop-shadow(0 0 1.5px #1f6fd1) drop-shadow(0 0 1px #1f6fd1)';
            if (['rect', 'circle'].includes(item.localName) || item.localName === 'path' && item.getAttribute('fill') !== 'none') {
              item.style.stroke = '#1f6fd1'; item.style.strokeWidth = '1.5px'; item.style.strokeDasharray = '3 2';
            }
            if (part.kind === 'legend' || part.kind === 'title') item.style.outline = '1px dashed #1f6fd1';
          }
        }
      }
      next.set(key, node);
      return node;
    });
    p.objHtml = next;
    const cur = p.objects.childNodes;
    if (cur.length !== nodes.length || nodes.some((n, i) => cur[i] !== n)) {
      // 다시 만든 슬라이서의 항목 목록 스크롤 위치 유지
      const keep = new Map();
      for (const l of p.objects.querySelectorAll('.sl-items')) if (l.scrollTop || l.scrollLeft) keep.set(l.closest('.obj')?.dataset.id, [l.scrollTop, l.scrollLeft]);
      p.objects.replaceChildren(...nodes);
      fitSlicerText(p.objects);
      for (const [id, [t, l]] of keep) {
        const list = [...p.objects.querySelectorAll('.obj')].find((o) => o.dataset.id === id)?.querySelector('.sl-items');
        if (list) { list.scrollTop = t; list.scrollLeft = l; }
      }
    }
    for (const node of nodes) { fitShapeText(node, appearance); markChartSelection(node, st.chartPart); }
  }

  /** 시간 표시 막대 (엑셀 Timeline): 날짜 필드를 연 · 분기 · 월 · 일 칸으로, 끌어서 기간 선택 */
  timelineHtml(sl) {
    const m = this.host.slicerModel?.(sl) ?? { items: [], filtered: false };
    const level = sl.level ?? 'M';
    const periods = timelinePeriods(m.items, level);
    const on = (p) => m.filtered && p.items.some((it) => it.selected);
    const sel = periods.filter(on);
    const LV = { Y: '연도', Q: '분기', M: '월', D: '일' };
    const rangeText = !m.filtered ? '모든 기간' : sel.length ? (sel.length === 1 ? sel[0].long : `${sel[0].long} - ${sel.at(-1).long}`) : '선택 없음';
    // 위 줄: 상위 단위(연도 · 월) 이름, 아래 줄: 칸
    let groups = '';
    let last = null;
    periods.forEach((p, i) => { if (p.group !== last) { groups += `<span class="tl-grp" style="grid-column:${i + 1}">${esc(p.group)}</span>`; last = p.group; } });
    const cells = periods.map((p, i) => `<button type="button" class="tl-cell${on(p) ? ' on' : ''}${p.items.some((it) => it.hasData) ? '' : ' nodata'}" data-i="${i}" style="grid-column:${i + 1}" title="${esc(p.long)}">${esc(p.short)}</button>`).join('');
    const w = Math.max(36, sl.cellW ?? 44);
    return `<div class="sl-head tl-head"><span class="sl-cap">${esc(sl.caption ?? '')}</span>`
      + `<button type="button" class="sl-clear${m.filtered ? '' : ' off'}" title="필터 지우기 (Alt+C)"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M1.5 2h11l-4.2 5v5l-2.6 1.5V7z"/><path d="M10.5 10l4 4M14.5 10l-4 4" stroke="#d13438"/></svg></button></div>`
      + `<div class="tl-sub"><span class="tl-range">${esc(rangeText)}</span><button type="button" class="tl-level" title="시간 수준">${LV[level]} ▾</button></div>`
      + `<div class="tl-bar" data-n="${periods.length}"><div class="tl-grid" style="grid-template-columns:repeat(${Math.max(1, periods.length)}, ${w}px)">${groups}${cells}</div></div>`;
  }

  /** 슬라이서: 머리글(캡션·다중 선택·필터 지우기) + 항목 단추 */
  slicerHtml(sl) {
    if (sl.timeline) return this.timelineHtml(sl);
    const m = this.host.slicerModel?.(sl) ?? { items: [], filtered: false };
    const items = m.broken
      ? `<div class="sl-broken">${esc(m.broken)}</div>`
      : m.items.map((it) => `<button type="button" class="sl-item${it.selected ? ' on' : ''}${it.hasData ? '' : ' nodata'}" data-k="${esc(it.key)}" title="${esc(it.text)}">${esc(it.text)}</button>`).join('');
    const head = sl.showHeader === false ? '' : `<div class="sl-head"><span class="sl-cap">${esc(sl.caption ?? '')}</span>`
      + `<button type="button" class="sl-multi${sl.multi ? ' on' : ''}" title="다중 선택 (Alt+S)"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M1.5 3.5l1.3 1.3 2.2-2.4M1.5 8.5l1.3 1.3 2.2-2.4M1.5 13.2l1.3 1.3 2.2-2.4"/><path d="M7 4h7.5M7 9h7.5M7 14h7.5"/></svg></button>`
      + `<button type="button" class="sl-clear${m.filtered ? '' : ' off'}" title="필터 지우기 (Alt+C)"><svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M1.5 2h11l-4.2 5v5l-2.6 1.5V7z"/><path d="M10.5 10l4 4M14.5 10l-4 4" stroke="#d13438"/></svg></button></div>`;
    const gap = Number.isFinite(sl.gap) ? `;gap:${sl.gap}px` : '';
    const btnW = sl.buttonWidth ? `${sl.buttonWidth}px` : 'minmax(0, 1fr)';
    return `${head}<div class="sl-items" style="grid-template-columns:repeat(${Math.max(1, sl.columns ?? 1)}, ${btnW})${gap}">${items}</div>`;
  }

  /** 피벗 차트 필드 단추 (엑셀처럼 차트 위에서 바로 거르기) */
  pivotChartButtons(ch) {
    if (!ch.pivot || ch.fieldButtons === false) return '';
    const fields = this.host.pivotChartFields?.(ch) ?? null;
    if (!fields) return '';
    const btn = (f, kind, cls) => `<button type="button" class="pc-field ${cls}" data-k="${kind}" data-f="${esc(f)}" title="${esc(f)} 거르기">${esc(f)} ▾</button>`;
    return `<div class="pc-fields">${fields.pages.map((f) => btn(f, 'pages', 'page')).join('')}${fields.values.map((f) => `<span class="pc-field val">${esc(f)}</span>`).join('')}</div>`
      + `<div class="pc-fields axis">${fields.rows.map((f) => btn(f, 'rows', 'row')).join('')}</div>`
      + (fields.cols.length ? `<div class="pc-fields legend">${fields.cols.map((f) => btn(f, 'cols', 'col')).join('')}</div>` : '');
  }

  chartSvg(ch) {
    const { wb, si } = this.host.state();
    return renderChartSvg(ch, chartModelData(wb, si, ch));
  }

  /** 선택 영역이 바뀌었을 때 (셀은 그대로) */
  renderSelection() {
    this.renderOverlays();
    this.renderHeaders(this.paneRects());
    this.a11y?.update();
  }

  renderObjectsAll() { for (const p of this.panes) if (p.win) this.renderObjects(p); }

  // ───────────── 선택 영역 등 겹쳐 그리기 ─────────────
  renderOverlays() { for (const p of this.panes) if (p.win) this.renderPaneOverlay(p); }

  renderPaneOverlay(p) {
    const st = this.host.state();
    const { wb, si, sel, active, selKind } = st;
    const win = p.win;
    const bx1 = this.cols.pos(win.c1) - 3;
    const by1 = this.rows.pos(win.r1) - 3;
    const bx2 = this.cols.pos(win.c2 + 1) + 3;
    const by2 = this.rows.pos(win.r2 + 1) + 3;
    const clip = (r) => {
      const x1 = Math.max(r.x, bx1);
      const y1 = Math.max(r.y, by1);
      const x2 = Math.min(r.x + r.w, bx2);
      const y2 = Math.min(r.y + r.h, by2);
      return x2 <= x1 || y2 <= y1 ? null : { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    };
    const box = (cls, r, extra = '') => {
      const c = clip(r);
      return c ? `<div class="${cls}" style="left:${c.x - p.ox}px;top:${c.y - p.oy}px;width:${c.w}px;height:${c.h}px;${extra}"></div>` : '';
    };
    const html = [];
    const selRect = this.sheetRect(sel);
    const am = wb.mergeAt(si, active.r, active.c) ?? { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
    // 포커스 셀 (엑셀 365): 활성 셀의 행 · 열을 띠로 강조
    if (st.focusCell && !st.editing) {
      const ar = this.sheetRect(am);
      html.push(box('focus-band', { x: bx1, y: ar.y, w: bx2 - bx1, h: ar.h }), box('focus-band', { x: ar.x, y: by1, w: ar.w, h: by2 - by1 }));
    }
    const single = sel.r1 === am.r1 && sel.c1 === am.c1 && sel.r2 === am.r2 && sel.c2 === am.c2;
    if (!single) {
      const a = this.sheetRect(am);
      const t = selRect;
      html.push(box('tint', { x: t.x, y: t.y, w: t.w, h: a.y - t.y }));
      html.push(box('tint', { x: t.x, y: a.y + a.h, w: t.w, h: t.y + t.h - a.y - a.h }));
      html.push(box('tint', { x: t.x, y: a.y, w: a.x - t.x, h: a.h }));
      html.push(box('tint', { x: a.x + a.w, y: a.y, w: t.x + t.w - a.x - a.w, h: a.h }));
    }
    html.push(box('sel-border', { x: selRect.x - 1, y: selRect.y - 1, w: selRect.w + 1, h: selRect.h + 1 }));
    if (!st.editing && selKind === 'cells' && !st.chartSel && st.fillHandle !== false) {
      const hx = selRect.x + selRect.w - 4;
      const hy = selRect.y + selRect.h - 4;
      if (hx >= bx1 && hx <= bx2 && hy >= by1 && hy <= by2) html.push(`<div class="fill-handle" style="left:${hx - p.ox}px;top:${hy - p.oy}px"></div>`);
    }
    // 데이터 유효성 검사: 목록 단추 · 잘못된 데이터 동그라미
    if (!st.editing && !st.chartSel) {
      const rule = validationAt(wb.sheets[si], active.r, active.c);
      const tt = tableAt(wb.sheets[si], active.r, active.c);
      if (tt?.totals && active.r === tt.r2) {
        const a = this.sheetRect(am);
        const x = a.x + a.w + 1;
        const y = a.y + a.h - 18;
        if (x >= bx1 && x <= bx2 && y + 18 >= by1 && y <= by2) html.push(`<div class="dv-btn" data-tt="1" title="요약 함수 선택" style="left:${x - p.ox}px;top:${y - p.oy}px"></div>`);
      } else if (rule?.type === 'list' && rule.showDropdown !== false) {
        const a = this.sheetRect(am);
        const x = a.x + a.w + 1;
        const y = a.y + a.h - 18;
        if (x >= bx1 && x <= bx2 && y + 18 >= by1 && y <= by2) html.push(`<div class="dv-btn" title="목록에서 선택" style="left:${x - p.ox}px;top:${y - p.oy}px"></div>`);
      }
    }
    // 동적 배열 분산 범위 (파란 점선)
    if (!st.editing && !st.chartSel) {
      const own = wb.getCell(si, active.r, active.c);
      const anc = own?.formula ? { r: active.r, c: active.c } : !own?.raw ? wb.spillAnchorOf(si, active.r, active.c) : null;
      const sp = anc ? wb.spillRange(si, anc.r, anc.c) : null;
      if (sp) {
        const r = this.sheetRect(sp);
        html.push(box('spill-border', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
      }
    }
    // 이동 옵션으로 고른 칸 (보이는 창 안만)
    if (st.special) {
      let n = 0;
      for (const [sr, sc] of st.special) {
        if (sr < win.r1 || sr > win.r2 || sc < win.c1 || sc > win.c2) continue;
        const r = this.sheetRect({ r1: sr, c1: sc, r2: sr, c2: sc });
        html.push(box('special-hl', { x: r.x, y: r.y, w: r.w, h: r.h }));
        if (++n > 5000) break;
      }
    }
    // 추적 화살표: 참조 범위(파란 테두리 + 점) → 수식 셀, 다른 시트는 점선 + 시트 아이콘
    if (st.arrows?.length) {
      const svg = [];
      const center = (rg) => { const r = this.sheetRect(rg); return [r.x + Math.min(r.w, this.cols.size(rg.c1)) / 2, r.y + Math.min(r.h, this.rows.size(rg.r1)) / 2]; };
      for (const a of st.arrows) {
        if (a.to.si !== si && a.from.si !== si) continue;
        const color = a.err ? '#d13438' : '#2f5bd3';
        const [x2, y2] = a.to.si === si ? center({ r1: a.to.r, c1: a.to.c, r2: a.to.r, c2: a.to.c }) : [null, null];
        if (a.from.si === si) {
          const multi = a.from.r1 !== a.from.r2 || a.from.c1 !== a.from.c2;
          if (multi) { const r = this.sheetRect(a.from); svg.push(`<rect x="${r.x}" y="${r.y}" width="${r.w}" height="${r.h}" fill="none" stroke="${color}" stroke-width="1.5"/>`); }
          const [x1, y1] = center(a.from);
          svg.push(`<circle cx="${x1}" cy="${y1}" r="3" fill="${color}"/>`);
          if (x2 !== null) svg.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" stroke-width="1.5" marker-end="url(#arrowHead${a.err ? 'E' : ''})"/>`);
        } else if (x2 !== null) {
          // 다른 시트의 참조: 왼쪽 위에서 점선 + 표 아이콘
          const x1 = x2 - 60;
          const y1 = y2 - 40;
          svg.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#333" stroke-dasharray="4 3" stroke-width="1.2" marker-end="url(#arrowHead)"/><rect x="${x1 - 9}" y="${y1 - 8}" width="16" height="13" fill="#fff" stroke="#333"/><path d="M${x1 - 9} ${y1 - 3}h16M${x1 - 9} ${y1 + 1}h16M${x1 - 3} ${y1 - 8}v13" stroke="#333"/>`);
        }
      }
      html.push(`<svg class="trace-svg" style="left:${-p.ox}px;top:${-p.oy}px"><defs><marker id="arrowHead" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#2f5bd3"/></marker><marker id="arrowHeadE" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#d13438"/></marker></defs>${svg.join('')}</svg>`);
    }
    for (const cc of st.circles ?? []) {
      const r = this.sheetRect({ r1: cc.r, c1: cc.c, r2: cc.r, c2: cc.c });
      html.push(box('dv-circle', { x: r.x - 4, y: r.y - 3, w: r.w + 8, h: r.h + 6 }));
    }
    if (st.clip && st.clip.si === si) {
      const r = this.sheetRect(st.clip);
      html.push(box('marquee', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
    }
    if (st.fillPreview) {
      const r = this.sheetRect(st.fillPreview);
      html.push(box('fill-preview', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
    }
    for (const ref of st.refs ?? []) {
      const r = this.sheetRect(ref.rg);
      html.push(box('ref-box', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }, `border-color:${ref.color};background:${ref.color}14`));
    }
    setSafeHtml(p.overlay, html.join(''));
  }

  /** 개요 기호: 그룹 괄호 선 + 요약 행(열)의 +/− 단추 (보이는 범위만) */
  outlineMarks(ax, ol, i1, i2, offset, band) {
    const isRow = ax === 'r';
    const levels = isRow ? ol.rows : ol.cols;
    const coll = (isRow ? ol.rowsColl : ol.colsColl) ?? {};
    const after = isRow ? ol.below !== false : ol.right !== false;
    const axis = isRow ? this.rows : this.cols;
    const out = [`<div class="olband ${ax}" style="${isRow ? `left:0;top:0;width:${band}px;bottom:0` : `left:0;top:0;height:${band}px;right:0`}"></div>`];
    for (const g of groupsOf(levels)) {
      const s = after ? g.b + 1 : g.a - 1;
      const lo = Math.min(g.a, s);
      const hi = Math.max(g.b, s);
      if (hi < i1 || lo > i2) continue;
      const off = (g.level - 1) * 14 + 3;
      const closed = !!coll[s];
      // 괄호 선 (펼친 그룹만)
      if (!closed && axis.size(g.a) + axis.size(g.b) > 0) {
        const p1 = axis.pos(g.a) - offset;
        const p2 = axis.pos(g.b + 1) - offset;
        out.push(isRow
          ? `<i class="oll" style="left:${off + 6}px;top:${p1 + 2}px;height:${Math.max(0, p2 - p1 - 4)}px;width:1px"></i><i class="oll" style="left:${off + 6}px;top:${after ? p2 - 3 : p1 + 2}px;width:5px;height:1px"></i>`
          : `<i class="oll" style="top:${off + 6}px;left:${p1 + 2}px;width:${Math.max(0, p2 - p1 - 4)}px;height:1px"></i><i class="oll" style="top:${off + 6}px;left:${after ? p2 - 3 : p1 + 2}px;height:5px;width:1px"></i>`);
      }
      if (s >= 0 && axis.size(s) > 0) {
        const p = axis.pos(s) - offset + (axis.size(s) - 13) / 2;
        out.push(`<div class="olb${closed ? ' closed' : ''}" data-ax="${ax}" data-a="${g.a}" data-b="${g.b}" data-l="${g.level}" title="${closed ? '세부 정보 표시' : '세부 정보 숨기기'}" style="${isRow ? `left:${off}px;top:${p}px` : `top:${off}px;left:${p}px`}">${closed ? '+' : '−'}</div>`);
      }
    }
    return out.join('');
  }

  // ───────────── 머리글 ─────────────
  renderHeaders(rects) {
    const st = this.host.state();
    const { sel, selKind } = st;
    const show = st.showHeaders;
    this.colHead.style.display = show ? 'block' : 'none';
    this.rowHead.style.display = show ? 'block' : 'none';
    this.corner.style.display = show ? 'block' : 'none';
    if (!show) return;
    const { hw, hh } = this;
    const olw = this.olw ?? 0;
    const olh = this.olh ?? 0;
    const ol = st.wb.sheets[st.si]?.outline;
    this.corner.style.width = `${hw}px`;
    this.corner.style.height = `${hh}px`;
    // 수준 단추: 행은 모서리 아래쪽 가로로, 열은 오른쪽 세로로
    const lv = [];
    if (olw && ol) for (let L = 1; L <= maxLevel(ol.rows) + 1; L++) lv.push(`<div class="olv" data-ax="r" data-l="${L}" title="수준 ${L} 표시" style="left:${(L - 1) * 14 + 2}px;top:${hh - 17}px">${L}</div>`);
    if (olh && ol) for (let L = 1; L <= maxLevel(ol.cols) + 1; L++) lv.push(`<div class="olv" data-ax="c" data-l="${L}" title="수준 ${L} 표시" style="left:${hw - 17}px;top:${(L - 1) * 14 + 2}px">${L}</div>`);
    const lvHtml = lv.join('');
    if (this.corner._lv !== lvHtml) { setSafeHtml(this.corner, lvHtml); this.corner._lv = lvHtml; }
    Object.assign(this.colHead.style, { left: '0px', top: '0px', width: `${this.viewW}px`, height: `${hh}px` });
    Object.assign(this.rowHead.style, { left: '0px', top: '0px', width: `${hw}px`, height: `${this.viewH}px` });
    const colFull = selKind === 'cols' || selKind === 'all';
    const rowFull = selKind === 'rows' || selKind === 'all';
    const colCls = (c) => (c >= sel.c1 && c <= sel.c2 ? (colFull ? ' full' : ' hl') : '');
    const rowCls = (r) => (r >= sel.r1 && r <= sel.r2 ? (rowFull ? ' full' : ' hl') : '');

    const colPart = (clipEl, rect, c1, c2, offset) => {
      Object.assign(clipEl.style, { left: `${rect.x}px`, top: '0px', width: `${rect.w}px`, height: `${hh}px`, display: rect.w > 0 ? 'block' : 'none' });
      const out = [];
      for (const c of visibleAxisIndices(this.cols, c1, c2)) {
        const w = this.cols.size(c);
        out.push(`<div class="hc${colCls(c)}" style="left:${this.cols.pos(c) - offset}px;width:${w}px${olh ? `;top:${olh}px` : ''}">${colToName(c)}</div>`);
      }
      if (olh) out.push(this.outlineMarks('c', ol, c1, c2, offset, olh));
      setSafeHtml(clipEl, out.join(''));
    };
    const rowPart = (clipEl, rect, r1, r2, offset) => {
      Object.assign(clipEl.style, { left: '0px', top: `${rect.y}px`, width: `${hw}px`, height: `${rect.h}px`, display: rect.h > 0 ? 'block' : 'none' });
      const out = [];
      for (const r of visibleAxisIndices(this.rows, r1, r2)) {
        const h = this.rows.size(r);
        out.push(`<div class="hr${rowCls(r)}" style="top:${this.rows.pos(r) - offset}px;height:${h}px;line-height:${h - 1}px${olw ? `;left:${olw}px;width:${hw - olw}px` : ''}">${r + 1}</div>`);
      }
      if (olw) out.push(this.outlineMarks('r', ol, r1, r2, offset, olw));
      setSafeHtml(clipEl, out.join(''));
    };
    const tl = rects.tl;
    const br = rects.br;
    colPart(this.colHeadFrozen, { x: tl.x, w: tl.w }, 0, this.fc - 1, 0);
    const vx = this.visibleRange(this.panes[3], br);
    colPart(this.colHeadScroll, { x: br.x, w: br.w }, vx.c1, vx.c2, vx.x0);
    rowPart(this.rowHeadFrozen, { y: tl.y, h: tl.h }, 0, this.fr - 1, 0);
    rowPart(this.rowHeadScroll, { y: br.y, h: br.h }, vx.r1, vx.r2, vx.y0);
  }
}
