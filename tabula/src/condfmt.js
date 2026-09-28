// 조건부 서식: 규칙 판정 · 설명 · 아이콘 (DOM 없음)
// 규칙: { r1, c1, r2, c2, type, v1, v2, style, color, colors, icons, reverse, iconOnly, percent, period, formula, stopIfTrue }
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
];

const ARROW = (color, deg) => `<svg viewBox="0 0 16 16" width="14" height="14"><g transform="rotate(${deg} 8 8)"><path d="M8 2l5 6H9.6v6H6.4V8H3z" fill="${color}"/></g></svg>`;
const DOT = (color) => `<svg viewBox="0 0 16 16" width="14" height="14"><circle cx="8" cy="8" r="6" fill="${color}" stroke="rgba(0,0,0,.25)"/></svg>`;
const FLAG = (color) => `<svg viewBox="0 0 16 16" width="14" height="14"><path d="M3.5 2v12" stroke="#555" stroke-width="1.2"/><path d="M4 2.5h9l-2.5 3 2.5 3H4z" fill="${color}"/></svg>`;
const BARS = (n) => `<svg viewBox="0 0 16 16" width="14" height="14">${[0, 1, 2, 3].map((i) => `<rect x="${1 + i * 3.7}" y="${12 - i * 3}" width="2.8" height="${3 + i * 3}" fill="${i < n ? '#2f6db5' : '#cfd8e3'}"/>`).join('')}</svg>`;
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

/** 시트의 규칙마다 범위 통계를 미리 계산 */
export function prepareCond(wb, si) {
  const sheet = wb.sheets[si];
  const used = wb.usedRange(si);
  return (sheet.cond ?? []).map((rule) => {
    const nums = [];
    const counts = new Map();
    const r2 = Math.min(rule.r2, used.rows - 1);
    const c2 = Math.min(rule.c2, used.cols - 1);
    const needCounts = rule.type === 'dup' || rule.type === 'unique';
    for (let r = rule.r1; r <= r2; r++) {
      for (let c = rule.c1; c <= c2; c++) {
        const v = wb.getValue(si, r, c);
        if (typeof v === 'number') nums.push(v);
        if (needCounts && v !== null && v !== '' && !isError(v)) counts.set(dupKey(v), (counts.get(dupKey(v)) ?? 0) + 1);
      }
    }
    const desc = [...nums].sort((a, b) => b - a);
    let min = Infinity;
    let max = -Infinity;
    let sum = 0;
    for (const n of nums) { if (n < min) min = n; if (n > max) max = n; sum += n; }
    const rank = Math.max(1, Number(rule.v1) || 10);
    const k = rule.percent ? Math.max(1, Math.floor((desc.length * Math.min(100, rank)) / 100)) : rank;
    return {
      rule, counts, nums: desc,
      min: nums.length ? min : 0,
      max: nums.length ? max : 0,
      avg: nums.length ? sum / nums.length : 0,
      topCut: desc[Math.min(desc.length, k) - 1],
      bottomCut: desc[Math.max(0, desc.length - k)],
      asts: new Map(),
      period: rule.type === 'date' ? periodRange(rule.period) : null,
    };
  });
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
    case 'aboveAvg': return num() && v > prep.avg;
    case 'belowAvg': return num() && v < prep.avg;
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
    if (r < rule.r1 || r > rule.r2 || c < rule.c1 || c > rule.c2) continue;
    if (rule.type === 'bar') {
      if (!bar && typeof v === 'number') {
        const lo = Math.min(0, prep.min);
        bar = { pct: prep.max === lo ? 100 : Math.max(0, ((v - lo) / (prep.max - lo)) * 100), color: rule.color ?? '#638ec6' };
        if (rule.iconOnly) hideValue = true;
      }
    } else if (rule.type === 'scale') {
      if (!fill && typeof v === 'number' && rule.colors?.length >= 2) fill = scaleColor(rule.colors, prep.max === prep.min ? 0.5 : (v - prep.min) / (prep.max - prep.min));
    } else if (rule.type === 'icons') {
      if (!icon && typeof v === 'number') {
        const set = iconSetById(rule.icons);
        const n = set.icons.length;
        const span = prep.max - prep.min;
        const pct = span ? ((v - prep.min) / span) * 100 : 100;
        let idx = Math.min(n - 1, Math.floor((pct * n) / 100));
        if (pct >= 100) idx = n - 1;
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
    case 'aboveAvg': return '평균 초과';
    case 'belowAvg': return '평균 미만';
    case 'formula': return `수식: ${rule.formula ?? ''}`;
    case 'bar': return '데이터 막대';
    case 'scale': return `${rule.colors?.length === 3 ? 3 : 2}색조`;
    case 'icons': return `아이콘 집합: ${iconSetById(rule.icons).label}`;
    default: return rule.type;
  }
}
