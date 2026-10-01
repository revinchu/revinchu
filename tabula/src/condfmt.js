// 조건부 서식: 규칙 판정 · 설명 · 아이콘 (DOM 없음)
// 규칙: { r1, c1, r2, c2, more: [추가 범위], type, v1, v2, style, color, colors, icons, reverse, iconOnly, percent, period, formula, stopIfTrue,
//         cfvo: [{ type: min|max|num|percent|percentile|formula|autoMin|autoMax, v, gte }], gradient, negColor, equal, stdDev }
// sheet.cond 배열의 순서가 우선순위 (0번이 가장 먼저 적용되고, 겹치는 서식은 0번이 이김)
import { parse, evaluateFormula, shiftFormula, isError, compareValues } from './formula.js';
import { parseInput, formatGeneral } from './format.js';

export const CELL_OPS = [
  { id: 'between', label: '다음 값 사이', two: true },
  { id: 'notBetween', label: '다음 값 사이에 있지 않음', two: true },
  { id: 'eq', label: '=' },
  { id: 'ne', label: '<>' },
  { id: 'gt', label: '>' },
  { id: 'lt', label: '<' },
  { id: 'ge', label: '>=' },
  { id: 'le', label: '<=' },
];
export const TEXT_OPS = [
  { id: 'text', label: '포함' },
  { id: 'notText', label: '포함하지 않음' },
  { id: 'begins', label: '시작 문자' },
  { id: 'ends', label: '끝 문자' },
];
export const DATE_PERIODS = [
  { id: 'yesterday', label: '어제' }, { id: 'today', label: '오늘' }, { id: 'tomorrow', label: '내일' },
  { id: 'last7Days', label: '지난 7일' }, { id: 'lastWeek', label: '지난주' }, { id: 'thisWeek', label: '이번 주' },
  { id: 'nextWeek', label: '다음 주' }, { id: 'lastMonth', label: '지난달' }, { id: 'thisMonth', label: '이번 달' },
  { id: 'nextMonth', label: '다음 달' },
];

/** 아이콘 집합: 엑셀 이름 → 아이콘(낮은 값부터) */
export const ICON_SETS = [
  { id: '3Arrows', label: '방향 (화살표 3개)', icons: ['arrowDownRed', 'arrowSideYellow', 'arrowUpGreen'] },
  { id: '3ArrowsGray', label: '방향 (회색 화살표 3개)', icons: ['arrowDownGray', 'arrowSideGray', 'arrowUpGray'] },
  { id: '3TrafficLights1', label: '신호등 3개', icons: ['circleRed', 'circleYellow', 'circleGreen'] },
  { id: '3Symbols', label: '기호 3개 (원)', icons: ['xRed', 'bangYellow', 'checkGreen'] },
  { id: '3Flags', label: '깃발 3개', icons: ['flagRed', 'flagYellow', 'flagGreen'] },
  { id: '4Arrows', label: '방향 (화살표 4개)', icons: ['arrowDownRed', 'arrowDiagDownYellow', 'arrowDiagUpYellow', 'arrowUpGreen'] },
  { id: '5Arrows', label: '방향 (화살표 5개)', icons: ['arrowDownRed', 'arrowDiagDownYellow', 'arrowSideYellow', 'arrowDiagUpYellow', 'arrowUpGreen'] },
  { id: '5Rating', label: '등급 5개 (막대)', icons: ['bars0', 'bars1', 'bars2', 'bars3', 'bars4'] },
  { id: '3TrafficLights2', label: '신호등 3개 (테두리)', icons: ['lightRed', 'lightYellow', 'lightGreen'] },
  { id: '3Signs', label: '표지판 3개', icons: ['diamondRed', 'triangleYellow', 'circleGreen'] },
  { id: '3Symbols2', label: '기호 3개', icons: ['xPlainRed', 'bangPlainYellow', 'checkPlainGreen'] },
  { id: '3Stars', label: '별 3개', icons: ['star0', 'star1', 'star2'] },
  { id: '3Triangles', label: '삼각형 3개', icons: ['triDownRed', 'dashYellow', 'triUpGreen'] },
  { id: '4ArrowsGray', label: '방향 (회색 화살표 4개)', icons: ['arrowDownGray', 'arrowDiagDownGray', 'arrowDiagUpGray', 'arrowUpGray'] },
  { id: '4RedToBlack', label: '빨강-검정 4개', icons: ['circleBlack', 'circleGray', 'circlePink', 'circleRed'] },
  { id: '4Rating', label: '등급 4개 (막대)', icons: ['bars1', 'bars2', 'bars3', 'bars4'] },
  { id: '4TrafficLights', label: '신호등 4개', icons: ['circleBlack', 'circleRed', 'circleYellow', 'circleGreen'] },
  { id: '5ArrowsGray', label: '방향 (회색 화살표 5개)', icons: ['arrowDownGray', 'arrowDiagDownGray', 'arrowSideGray', 'arrowDiagUpGray', 'arrowUpGray'] },
  { id: '5Quarters', label: '사분원 5개', icons: ['quarter0', 'quarter1', 'quarter2', 'quarter3', 'quarter4'] },
  { id: '5Boxes', label: '상자 5개', icons: ['boxes0', 'boxes1', 'boxes2', 'boxes3', 'boxes4'] },
];

const ARROW = (color, deg) => `<svg viewBox="0 0 16 16" width="14" height="14"><g transform="rotate(${deg} 8 8)"><path d="M8 2l5 6H9.6v6H6.4V8H3z" fill="${color}"/></g></svg>`;
const DOT = (color) => `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="6" fill="${color}" stroke="rgba(0,0,0,.25)"/></svg>`;
const FLAG = (color) => `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M3.5 2v12" stroke="#555" stroke-width="1.2"/><path d="M4 2.5h9l-2.5 3 2.5 3H4z" fill="${color}"/></svg>`;
const BARS = (n) => `<svg viewBox="0 0 16 16" width="14" height="14">${[0, 1, 2, 3].map((i) => `<rect x="${1 + i * 3.7}" y="${12 - i * 3}" width="2.8" height="${3 + i * 3}" fill="${i < n ? '#2f6db5' : '#cfd8e3'}"/>`).join('')}</svg>`;
const RIM = (color) => `<svg viewBox="0 0 16 16" width="14" height="14"><rect x="1" y="1" width="14" height="14" rx="3" fill="#333"/><circle cx="8" cy="8" r="5" fill="${color}"/></svg>`;
const STAR = (fill) => `<svg viewBox="0 0 16 16" width="14" height="14"><defs><linearGradient id="sg${fill * 2}"><stop offset="${fill * 100}%" stop-color="#f0b400"/><stop offset="${fill * 100}%" stop-color="#fff"/></linearGradient></defs><path d="M8 1.5l1.9 4.2 4.6.4-3.5 3 1.1 4.5L8 11.2l-4.1 2.4 1.1-4.5-3.5-3 4.6-.4z" fill="url(#sg${fill * 2})" stroke="#c89600" stroke-width=".8"/></svg>`;
const QUARTER = (q) => {
  const a = (q / 4) * 2 * Math.PI;
  const x = 8 + 6 * Math.sin(a);
  const y = 8 - 6 * Math.cos(a);
  const slice = q === 4 ? '<circle cx="8" cy="8" r="6" fill="#444"/>' : q ? `<path d="M8 8V2A6 6 0 ${q > 2 ? 1 : 0} 1 ${x.toFixed(2)} ${y.toFixed(2)}z" fill="#444"/>` : '';
  return `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="6" fill="#fff" stroke="#444"/>${slice}</svg>`;
};
const BOXES = (n) => `<svg viewBox="0 0 16 16" width="14" height="14">${[0, 1, 2, 3].map((i) => `<rect x="${i % 2 ? 8.5 : 1.5}" y="${i < 2 ? 1.5 : 8.5}" width="6" height="6" fill="${[2, 3, 0, 1][i] < n ? '#2f6db5' : '#cfd8e3'}"/>`).join('')}</svg>`;
export const ICON_SVG = {
  arrowUpGreen: ARROW('#00a650', 0), arrowSideYellow: ARROW('#f0b400', 90), arrowDownRed: ARROW('#e0322c', 180),
  arrowDiagUpYellow: ARROW('#f0b400', 45), arrowDiagDownYellow: ARROW('#f0b400', 135),
  arrowUpGray: ARROW('#808080', 0), arrowSideGray: ARROW('#808080', 90), arrowDownGray: ARROW('#808080', 180),
  circleGreen: DOT('#00a650'), circleYellow: DOT('#f0c000'), circleRed: DOT('#e0322c'),
  checkGreen: `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="7" fill="#00a650"/><path d="M4.5 8.2l2.3 2.3 4.7-4.8" stroke="#fff" stroke-width="1.8" fill="none"/></svg>`,
  bangYellow: `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="7" fill="#f0b400"/><path d="M8 3.8v5.2" stroke="#fff" stroke-width="2"/><circle cx="8" cy="11.8" r="1.1" fill="#fff"/></svg>`,
  xRed: `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="7" fill="#e0322c"/><path d="M5.2 5.2l5.6 5.6m0-5.6l-5.6 5.6" stroke="#fff" stroke-width="1.8"/></svg>`,
  flagGreen: FLAG('#00a650'), flagYellow: FLAG('#f0c000'), flagRed: FLAG('#e0322c'),
  bars0: BARS(0), bars1: BARS(1), bars2: BARS(2), bars3: BARS(3), bars4: BARS(4),
  arrowDiagUpGray: ARROW('#808080', 45), arrowDiagDownGray: ARROW('#808080', 135),
  lightRed: RIM('#e0322c'), lightYellow: RIM('#f0c000'), lightGreen: RIM('#00a650'),
  circleBlack: DOT('#3a3a3a'), circleGray: DOT('#9a9a9a'), circlePink: DOT('#f4a0a0'),
  diamondRed: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M8 1.5l6.5 6.5L8 14.5 1.5 8z" fill="#e0322c"/></svg>`,
  triangleYellow: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M8 2l6.5 12h-13z" fill="#f0c000"/></svg>`,
  xPlainRed: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M4 4l8 8m0-8l-8 8" stroke="#e0322c" stroke-width="2.4"/></svg>`,
  bangPlainYellow: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M8 2.5v7.5" stroke="#f0b400" stroke-width="2.6"/><circle cx="8" cy="13" r="1.4" fill="#f0b400"/></svg>`,
  checkPlainGreen: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M3 8.5l3.2 3.2L13 4.5" stroke="#00a650" stroke-width="2.4" fill="none"/></svg>`,
  star0: STAR(0), star1: STAR(0.5), star2: STAR(1),
  triUpGreen: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M8 4l5 7H3z" fill="#00a650"/></svg>`,
  triDownRed: `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M8 12l5-7H3z" fill="#e0322c"/></svg>`,
  dashYellow: `<svg viewBox="0 0 16 16" width="14" height="14"><rect x="3" y="6.5" width="10" height="3" fill="#f0c000"/></svg>`,
  quarter0: QUARTER(0), quarter1: QUARTER(1), quarter2: QUARTER(2), quarter3: QUARTER(3), quarter4: QUARTER(4),
  boxes0: BOXES(0), boxes1: BOXES(1), boxes2: BOXES(2), boxes3: BOXES(3), boxes4: BOXES(4),
};

export const iconSetById = (id) => ICON_SETS.find((s) => s.id === id) ?? ICON_SETS[0];

const valueOf = (s) => parseInput(String(s ?? '')).value;
const dupKey = (v) => (typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${v}`);
const textOf = (v) => String(typeof v === 'number' ? formatGeneral(v) : v ?? '');

export function scaleColor(colors, t) {
  const hex = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const seg = colors.length - 1;
  const pos = Math.max(0, Math.min(1, t)) * seg;
  const i = Math.min(seg - 1, Math.floor(pos));
  const f = pos - i;
  const a = hex(colors[i]);
  const b = hex(colors[i + 1]);
  return `rgb(${a.map((x, k) => Math.round(x + (b[k] - x) * f)).join(',')})`;
}

/** 오늘 기준 날짜 범위 [시작, 끝) (엑셀 일련 번호) */
export function periodRange(period, now = new Date()) {
  const EPOCH = Date.UTC(1899, 11, 30);
  const today = (Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - EPOCH) / 86400000;
  const dow = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate())).getUTCDay();
  const monthStart = (y, m) => (Date.UTC(y, m, 1) - EPOCH) / 86400000;
  const y = now.getFullYear();
  const m = now.getMonth();
  switch (period) {
    case 'yesterday': return [today - 1, today];
    case 'tomorrow': return [today + 1, today + 2];
    case 'last7Days': return [today - 6, today + 1];
    case 'thisWeek': return [today - dow, today - dow + 7];
    case 'lastWeek': return [today - dow - 7, today - dow];
    case 'nextWeek': return [today - dow + 7, today - dow + 14];
    case 'thisMonth': return [monthStart(y, m), monthStart(y, m + 1)];
    case 'lastMonth': return [monthStart(y, m - 1), monthStart(y, m)];
    case 'nextMonth': return [monthStart(y, m + 1), monthStart(y, m + 2)];
    default: return [today, today + 1];
  }
}

/** 규칙이 적용되는 모든 범위 (sqref 여러 개) */
export const ruleRanges = (rule) => [rule, ...(rule.more ?? [])];
export const inRule = (rule, r, c) => ruleRanges(rule).some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2);

/** 시트의 규칙마다 범위 통계를 미리 계산 */
export function prepareCond(wb, si) {
  const sheet = wb.sheets[si];
  const used = wb.usedRange(si);
  return (sheet.cond ?? []).map((rule) => {
    const nums = [];
    const counts = new Map();
    const needCounts = rule.type === 'dup' || rule.type === 'unique';
    for (const g of ruleRanges(rule)) {
      const r2 = Math.min(g.r2, used.rows - 1);
      const c2 = Math.min(g.c2, used.cols - 1);
      for (let r = g.r1; r <= r2; r++) {
        for (let c = g.c1; c <= c2; c++) {
          const v = wb.getValue(si, r, c);
          if (typeof v === 'number') nums.push(v);
          if (needCounts && v !== null && v !== '' && !isError(v)) counts.set(dupKey(v), (counts.get(dupKey(v)) ?? 0) + 1);
        }
      }
    }
    const desc = [...nums].sort((a, b) => b - a);
    let min = Infinity;
    let max = -Infinity;
    let sum = 0;
    for (const n of nums) { if (n < min) min = n; if (n > max) max = n; sum += n; }
    const rank = Math.max(1, Number(rule.v1) || 10);
    const k = rule.percent ? Math.max(1, Math.floor((desc.length * Math.min(100, rank)) / 100)) : rank;
    const avg = nums.length ? sum / nums.length : 0;
    const sd = nums.length > 1 ? Math.sqrt(nums.reduce((s2, x) => s2 + (x - avg) ** 2, 0) / (nums.length - 1)) : 0;
    const prep = {
      rule, counts, nums: desc,
      min: nums.length ? min : 0,
      max: nums.length ? max : 0,
      avg, sd,
      topCut: desc[Math.min(desc.length, k) - 1],
      bottomCut: desc[Math.max(0, desc.length - k)],
      asts: new Map(),
      period: rule.type === 'date' ? periodRange(rule.period) : null,
    };
    if (rule.cfvo?.length) prep.points = rule.cfvo.map((p) => cfvoValue(p, prep, wb, si));
    return prep;
  });
}

/** 임계값(cfvo) → 숫자 */
function cfvoValue(p, prep, wb, si) {
  const { min, max } = prep;
  const v = Number(p.v);
  switch (p.type) {
    case 'min': return min;
    case 'max': return max;
    case 'autoMin': return Math.min(0, min);
    case 'autoMax': return Math.max(0, max);
    case 'percent': return min + ((max - min) * (Number.isFinite(v) ? v : 0)) / 100;
    case 'percentile': {
      const a = [...prep.nums].reverse();
      if (!a.length) return 0;
      const h = (a.length - 1) * Math.max(0, Math.min(100, v)) / 100;
      const lo = Math.floor(h);
      return a[lo] + (h - lo) * ((a[lo + 1] ?? a[lo]) - a[lo]);
    }
    case 'formula': {
      const text = String(p.v ?? '');
      let r = null;
      try { r = evaluateFormula(parse(text.replace(/^=/, '')), wb.ctxFor(si, prep.rule.r1, prep.rule.c1)); } catch { r = null; }
      return typeof r === 'number' ? r : Number(r) || 0;
    }
    default: return Number.isFinite(v) ? v : Number(valueOf(p.v)) || 0;
  }
}

/** 수식 조건 (범위 왼쪽 위 셀 기준 상대 참조) */
function formulaValue(prep, wb, si, r, c, text) {
  const shifted = shiftFormula(text.startsWith('=') ? text : `=${text}`, r - prep.rule.r1, c - prep.rule.c1);
  let ast = prep.asts.get(shifted);
  if (ast === undefined) {
    try { ast = parse(shifted.slice(1)); } catch { ast = null; }
    if (prep.asts.size < 5000) prep.asts.set(shifted, ast);
  }
  if (!ast) return null;
  try { return evaluateFormula(ast, wb.ctxFor(si, r, c)); } catch { return null; }
}

function operand(prep, wb, si, r, c, s) {
  const t = String(s ?? '');
  return t.startsWith('=') ? formulaValue(prep, wb, si, r, c, t) : valueOf(t);
}

/** 규칙이 이 셀에 맞는지 (막대·색조·아이콘 제외) */
export function condMatch(prep, v, wb = null, si = 0, r = 0, c = 0) {
  const { rule } = prep;
  const t = rule.type;
  if (t === 'blank') return v === null || v === '' || (typeof v === 'string' && !v.trim());
  if (t === 'noBlank') return !(v === null || v === '' || (typeof v === 'string' && !v.trim()));
  if (t === 'errors') return isError(v);
  if (t === 'noErrors') return !isError(v);
  if (t === 'formula') {
    const res = wb ? formulaValue(prep, wb, si, r, c, rule.formula ?? '') : null;
    return res === true || (typeof res === 'number' && res !== 0);
  }
  if (v === null || v === '' || isError(v)) return false;
  const num = () => typeof v === 'number';
  const a = () => operand(prep, wb, si, r, c, rule.v1);
  const b = () => operand(prep, wb, si, r, c, rule.v2);
  const cmp = (x) => (typeof x === typeof v || (typeof x === 'number' && num()) ? compareValues(v, x) : null);
  switch (t) {
    case 'gt': { const x = a(); return num() && typeof x === 'number' && v > x; }
    case 'lt': { const x = a(); return num() && typeof x === 'number' && v < x; }
    case 'ge': { const x = a(); return num() && typeof x === 'number' && v >= x; }
    case 'le': { const x = a(); return num() && typeof x === 'number' && v <= x; }
    case 'between': case 'notBetween': {
      const x = a();
      const y = b();
      if (!num() || typeof x !== 'number' || typeof y !== 'number') return false;
      const inside = v >= Math.min(x, y) && v <= Math.max(x, y);
      return t === 'between' ? inside : !inside;
    }
    case 'eq': return cmp(a()) === 0;
    case 'ne': return cmp(a()) !== 0;
    case 'text': return textOf(v).toLowerCase().includes(String(rule.v1 ?? '').toLowerCase());
    case 'notText': return !textOf(v).toLowerCase().includes(String(rule.v1 ?? '').toLowerCase());
    case 'begins': return textOf(v).toLowerCase().startsWith(String(rule.v1 ?? '').toLowerCase());
    case 'ends': return textOf(v).toLowerCase().endsWith(String(rule.v1 ?? '').toLowerCase());
    case 'date': return num() && Math.floor(v) >= prep.period[0] && Math.floor(v) < prep.period[1];
    case 'dup': return (prep.counts.get(dupKey(v)) ?? 0) > 1;
    case 'unique': return (prep.counts.get(dupKey(v)) ?? 0) === 1;
    case 'top': return num() && prep.topCut !== undefined && v >= prep.topCut;
    case 'bottom': return num() && prep.bottomCut !== undefined && v <= prep.bottomCut;
    case 'aboveAvg': {
      const th = prep.avg + (Number(rule.stdDev) || 0) * prep.sd;
      return num() && (rule.equal ? v >= th : v > th);
    }
    case 'belowAvg': {
      const th = prep.avg - (Number(rule.stdDev) || 0) * prep.sd;
      return num() && (rule.equal ? v <= th : v < th);
    }
    default: return false;
  }
}

export const VISUAL_TYPES = new Set(['bar', 'scale', 'icons']);
/** 빈 셀에도 맞을 수 있는 규칙 (화면에서 빈 셀도 그려야 함) */
export const EMPTY_MATCH_TYPES = new Set(['blank', 'formula', 'noErrors']);

/**
 * 셀 하나의 조건부 서식 결과 → { style: 덮어쓸 서식 | null, bar, icon, hideValue }
 * 우선순위가 높은(앞쪽) 규칙의 서식이 이기고, '(True일 경우) 중지' 규칙이 맞으면 뒤 규칙은 보지 않음
 */
export function condFormatAt(preps, wb, si, r, c, v) {
  const matched = [];
  let bar = null;
  let icon = null;
  let hideValue = false;
  let fill = null;
  for (const prep of preps) {
    const rule = prep.rule;
    if (!inRule(rule, r, c)) continue;
    if (rule.type === 'bar') {
      if (!bar && typeof v === 'number') {
        const lo = prep.points ? prep.points[0] : Math.min(0, prep.min);
        const hi = prep.points ? prep.points[1] : prep.max;
        const neg = v < 0 && rule.negColor && lo < 0;
        // 값이 모두 0 이면 막대 없음 (엑셀: 축 0 에서 길이 0)
        const pct = neg ? Math.min(100, (v / lo) * 100) : hi === lo ? (v === 0 ? 0 : 100) : Math.max(0, Math.min(100, ((v - lo) / (hi - lo)) * 100));
        bar = { pct, color: neg ? rule.negColor : rule.color ?? '#638ec6', gradient: rule.gradient !== false, neg };
        if (rule.iconOnly) hideValue = true;
      }
    } else if (rule.type === 'scale') {
      if (!fill && typeof v === 'number' && rule.colors?.length >= 2) {
        const pts = prep.points?.length === rule.colors.length ? prep.points : null;
        let t;
        if (pts) {
          // 임계값 사이 구간별 보간
          const n = pts.length - 1;
          if (v <= pts[0]) t = 0;
          else if (v >= pts[n]) t = 1;
          else {
            let i = 0;
            while (i < n - 1 && v > pts[i + 1]) i++;
            const span = pts[i + 1] - pts[i];
            t = (i + (span ? (v - pts[i]) / span : 0)) / n;
          }
        } else t = prep.max === prep.min ? 0.5 : (v - prep.min) / (prep.max - prep.min);
        fill = scaleColor(rule.colors, t);
      }
    } else if (rule.type === 'icons') {
      if (!icon && typeof v === 'number') {
        const set = iconSetById(rule.icons);
        const n = set.icons.length;
        let idx;
        if (prep.points?.length === n) {
          idx = 0;
          for (let i = n - 1; i >= 1; i--) {
            const th = prep.points[i];
            if (rule.cfvo[i].gte === false ? v > th : v >= th) { idx = i; break; }
          }
        } else {
          const span = prep.max - prep.min;
          const pct = span ? ((v - prep.min) / span) * 100 : 100;
          idx = Math.min(n - 1, Math.floor((pct * n) / 100));
          if (pct >= 100) idx = n - 1;
        }
        if (rule.reverse) idx = n - 1 - idx;
        icon = set.icons[idx];
        if (rule.iconOnly) hideValue = true;
      }
    } else if (condMatch(prep, v, wb, si, r, c)) {
      matched.push(rule.style ?? {});
      if (rule.stopIfTrue) break;
    }
  }
  let style = null;
  if (matched.length || fill) {
    style = fill ? { fill } : {};
    for (let i = matched.length - 1; i >= 0; i--) Object.assign(style, matched[i]);
  }
  return { style, bar, icon, hideValue };
}

/** 규칙 설명 (규칙 관리자 목록) */
export function describeCond(rule) {
  const op = CELL_OPS.find((o) => o.id === rule.type);
  if (op) return op.two ? `셀 값 ${op.label} ${rule.v1 ?? ''} 및 ${rule.v2 ?? ''}` : `셀 값 ${op.label} ${rule.v1 ?? ''}`;
  const tx = TEXT_OPS.find((o) => o.id === rule.type);
  if (tx) return `특정 텍스트 ${tx.label}: "${rule.v1 ?? ''}"`;
  switch (rule.type) {
    case 'date': return `발생 날짜: ${DATE_PERIODS.find((p) => p.id === rule.period)?.label ?? '오늘'}`;
    case 'blank': return '빈 셀';
    case 'noBlank': return '내용 있는 셀';
    case 'errors': return '오류';
    case 'noErrors': return '오류 없음';
    case 'dup': return '중복 값';
    case 'unique': return '고유 값';
    case 'top': return `상위 ${rule.v1 ?? 10}${rule.percent ? '%' : ''}`;
    case 'bottom': return `하위 ${rule.v1 ?? 10}${rule.percent ? '%' : ''}`;
    case 'aboveAvg': return rule.stdDev ? `평균 초과 (표준 편차 ${rule.stdDev})` : rule.equal ? '평균 이상' : '평균 초과';
    case 'belowAvg': return rule.stdDev ? `평균 미만 (표준 편차 ${rule.stdDev})` : rule.equal ? '평균 이하' : '평균 미만';
    case 'formula': return `수식: ${rule.formula ?? ''}`;
    case 'bar': return '데이터 막대';
    case 'scale': return `${rule.colors?.length === 3 ? 3 : 2}색조`;
    case 'icons': return `아이콘 집합: ${iconSetById(rule.icons).label}`;
    default: return rule.type;
  }
}

// ───────────── 색조 · 데이터 막대 모음 (엑셀 기본 12종 + WIXEL 모던) ─────────────
export const SCALE_PRESETS = {
  modern: [
    ['에메랄드 - 앰버 - 로즈', ['#34d399', '#fde68a', '#fb7185']],
    ['로즈 - 앰버 - 에메랄드', ['#fb7185', '#fde68a', '#34d399']],
    ['인디고 - 흰색 - 로즈', ['#6366f1', '#ffffff', '#f43f5e']],
    ['로즈 - 흰색 - 인디고', ['#f43f5e', '#ffffff', '#6366f1']],
    ['흰색 - 에메랄드', ['#ffffff', '#34d399']],
    ['흰색 - 인디고', ['#ffffff', '#818cf8']],
    ['흰색 - 오션', ['#ffffff', '#38bdf8']],
    ['흰색 - 로즈', ['#ffffff', '#fb7185']],
    ['비리디스', ['#fde725', '#21918c', '#440154']],
    ['선셋', ['#fef3c7', '#fb923c', '#be123c']],
    ['오로라', ['#ecfeff', '#67e8f9', '#0e7490']],
    ['슬레이트', ['#f8fafc', '#94a3b8', '#334155']],
  ],
  excel: [
    ['녹색 - 노랑 - 빨강', ['#63be7b', '#ffeb84', '#f8696b']],
    ['빨강 - 노랑 - 녹색', ['#f8696b', '#ffeb84', '#63be7b']],
    ['녹색 - 흰색 - 빨강', ['#63be7b', '#fcfcff', '#f8696b']],
    ['빨강 - 흰색 - 녹색', ['#f8696b', '#fcfcff', '#63be7b']],
    ['파랑 - 흰색 - 빨강', ['#5a8ac6', '#fcfcff', '#f8696b']],
    ['빨강 - 흰색 - 파랑', ['#f8696b', '#fcfcff', '#5a8ac6']],
    ['흰색 - 빨강', ['#fcfcff', '#f8696b']],
    ['빨강 - 흰색', ['#f8696b', '#fcfcff']],
    ['녹색 - 흰색', ['#63be7b', '#fcfcff']],
    ['흰색 - 녹색', ['#fcfcff', '#63be7b']],
    ['녹색 - 노랑', ['#63be7b', '#ffef9c']],
    ['노랑 - 녹색', ['#ffef9c', '#63be7b']],
  ],
};
// 기본 색조 (새 규칙): 모던 3색
export const DEFAULT_SCALE3 = SCALE_PRESETS.modern[1][1];
export const DEFAULT_SCALE2 = SCALE_PRESETS.modern[4][1];
export const BAR_PRESETS = {
  modern: [['인디고', '#818cf8'], ['오션', '#38bdf8'], ['에메랄드', '#34d399'], ['앰버', '#fbbf24'], ['로즈', '#fb7185'], ['바이올렛', '#a78bfa'], ['슬레이트', '#94a3b8']],
  excel: [['파랑', '#638ec6'], ['녹색', '#63c384'], ['빨강', '#ff555a'], ['주황', '#ffb628'], ['하늘색', '#008aef'], ['보라', '#d6007b']],
};
