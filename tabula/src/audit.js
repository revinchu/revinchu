// 수식 분석 · 가상 분석 (DOM 없음): 수식 계산 단계, 목표값 찾기, 데이터 표, 이동 옵션
import { evaluate, isError, Range } from './formula.js';
import { formatGeneral } from './format.js';

// ───────────── 수식 계산 (단계별) ─────────────
const LITERAL = new Set(['num', 'str', 'bool', 'err', 'empty', 'array']);
/** AST 노드의 바로 아래 자식 노드들 (원문 위치 s/e 가 있는 것) */
function childNodes(n) {
  const out = [];
  for (const k of ['a', 'b', 'fn']) if (n[k] && typeof n[k] === 'object' && n[k].type) out.push(n[k]);
  for (const k of ['args', 'items']) if (Array.isArray(n[k])) for (const x of n[k]) if (x && x.type) out.push(x);
  return out.filter((x) => x.s !== undefined && x.e !== undefined).sort((x, y) => x.s - y.s);
}

/** 계산 결과 → 수식 계산 창에 보일 글자 */
export function valueText(v) {
  if (v === null || v === undefined) return '0';
  if (isError(v)) return v.code;
  if (typeof v === 'number') return formatGeneral(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
  if (v instanceof Range) {
    const h = v.height;
    const w = v.width;
    const rows = v.rows.slice(0, 4).map((row) => row.slice(0, 6).map(valueText).join(',') + (w > 6 ? ',…' : ''));
    return `{${rows.join(';')}${h > 4 ? ';…' : ''}}`;
  }
  return String(v);
}

/**
 * 수식 계산 단계: 안쪽 · 왼쪽 부분식부터 하나씩 값으로 바꾼 글자 목록.
 * → [{ text: '=…', mark: [s, e] (다음에 계산할 부분, 없으면 null) }]. 마지막은 결과
 */
export function evalSteps(src, ast, ctx) {
  const order = [];
  const walk = (n) => {
    for (const ch of childNodes(n)) walk(ch);
    if (!LITERAL.has(n.type) && n.type !== 'paren') order.push(n);
  };
  walk(ast);
  const done = new Map(); // 노드 → 결과 글자
  const render = (n, marks) => {
    if (done.has(n)) {
      const t = done.get(n);
      if (marks.target === n) marks.pos = [marks.off, marks.off + t.length];
      return t;
    }
    let out = '';
    let pos = n.s;
    const startOff = marks.off;
    for (const ch of childNodes(n)) {
      const lit = src.slice(pos, ch.s);
      out += lit;
      marks.off += lit.length;
      const t = render(ch, marks);
      out += t;
      marks.off += t.length;
      pos = ch.e;
    }
    out += src.slice(pos, n.e);
    if (marks.target === n) marks.pos = [startOff, startOff + out.length];
    marks.off = startOff;
    return out;
  };
  const snapshot = (target) => {
    const marks = { off: 1, target, pos: null };
    const body = render(ast, marks);
    return { text: `=${src.slice(0, ast.s)}${body}${src.slice(ast.e)}`, mark: target ? marks.pos : null };
  };
  const steps = [];
  for (const n of order) {
    steps.push(snapshot(n));
    let v;
    try { v = evaluate(n, ctx); } catch (e) { v = e; }
    done.set(n, valueText(v));
  }
  steps.push({ text: `=${valueText((() => { try { return evaluate(ast, ctx); } catch (e) { return e; } })())}`.replace(/^=/, ''), mark: null, final: true });
  return steps;
}

// ───────────── 목표값 찾기 ─────────────
/**
 * f(x) = 목표 가 되는 x 찾기 (할선법 + 부호가 바뀌면 이분법). f 는 값(숫자가 아니면 NaN)을 돌려줌
 * → { x, value, ok, iterations }
 */
export function goalSeek(f, x0, target, { maxIter = 100, tol = 0.001 } = {}) {
  const g = (x) => { const v = f(x); return typeof v === 'number' && Number.isFinite(v) ? v - target : NaN; };
  let a = Number.isFinite(x0) ? x0 : 0;
  let fa = g(a);
  if (Math.abs(fa) <= tol) return { x: a, value: fa + target, ok: true, iterations: 0 };
  let b = a === 0 ? 1 : a * 1.01 + 0.01;
  let fb = g(b);
  let lo = null;
  let hi = null; // 부호가 바뀌는 구간
  const bracket = (x, fx) => {
    if (!Number.isFinite(fx)) return;
    if (fx < 0 && (lo === null || Math.abs(fx) < Math.abs(lo[1]))) lo = [x, fx];
    if (fx > 0 && (hi === null || Math.abs(fx) < Math.abs(hi[1]))) hi = [x, fx];
  };
  bracket(a, fa);
  bracket(b, fb);
  for (let i = 1; i <= maxIter; i++) {
    let x;
    if (Number.isFinite(fa) && Number.isFinite(fb) && fb !== fa) x = b - (fb * (b - a)) / (fb - fa);
    else x = b + (b - a || 1) * 2;
    // 할선법이 튀면 알려진 구간의 가운데로
    if (lo && hi && (!Number.isFinite(x) || x < Math.min(lo[0], hi[0]) || x > Math.max(lo[0], hi[0]))) x = (lo[0] + hi[0]) / 2;
    if (!Number.isFinite(x)) x = b * 2 + 1;
    const fx = g(x);
    bracket(x, fx);
    if (Number.isFinite(fx) && Math.abs(fx) <= tol) return { x, value: fx + target, ok: true, iterations: i };
    a = b; fa = fb;
    b = x; fb = fx;
    if (!Number.isFinite(fx) && lo && hi) { b = (lo[0] + hi[0]) / 2; fb = g(b); }
  }
  const best = [lo, hi].filter(Boolean).sort((p, q) => Math.abs(p[1]) - Math.abs(q[1]))[0];
  return best ? { x: best[0], value: best[1] + target, ok: false, iterations: maxIter } : { x: b, value: fb + target, ok: false, iterations: maxIter };
}

// ───────────── 데이터 표 ─────────────
/**
 * 데이터 표(가상 분석): 범위 rg 의 첫 행 · 첫 열이 입력 값, 모서리(또는 첫 행/열)에 수식.
 *  - 행 입력만: 첫 행(모서리 오른쪽)에 입력 값, 첫 열(모서리 아래)에 수식들
 *  - 열 입력만: 첫 열(모서리 아래)에 입력 값, 첫 행(모서리 오른쪽)에 수식들
 *  - 둘 다: 모서리에 수식 하나
 * setInput(셀, 값) 으로 입력 셀을 바꾸고 valueOf(r, c) 로 수식 결과를 읽음 → [[r, c, 값]]
 */
export function dataTable(rg, rowInput, colInput, setInput, valueOf) {
  const out = [];
  if (rowInput && colInput) {
    for (let r = rg.r1 + 1; r <= rg.r2; r++) {
      for (let c = rg.c1 + 1; c <= rg.c2; c++) {
        setInput(rowInput, valueOf(rg.r1, c));
        setInput(colInput, valueOf(r, rg.c1));
        out.push([r, c, valueOf(rg.r1, rg.c1)]);
      }
    }
  } else if (colInput) {
    for (let r = rg.r1 + 1; r <= rg.r2; r++) {
      setInput(colInput, valueOf(r, rg.c1));
      for (let c = rg.c1 + 1; c <= rg.c2; c++) out.push([r, c, valueOf(rg.r1, c)]);
    }
  } else if (rowInput) {
    for (let c = rg.c1 + 1; c <= rg.c2; c++) {
      setInput(rowInput, valueOf(rg.r1, c));
      for (let r = rg.r1 + 1; r <= rg.r2; r++) out.push([r, c, valueOf(r, rg.c1)]);
    }
  }
  return out;
}

// ───────────── 이동 옵션 ─────────────
export const GOTO_KINDS = [
  { id: 'comments', label: '메모' },
  { id: 'constants', label: '상수' },
  { id: 'formulas', label: '수식' },
  { id: 'blanks', label: '빈 셀' },
  { id: 'errors', label: '오류 (수식 · 상수)' },
  { id: 'lastCell', label: '마지막 셀' },
  { id: 'visible', label: '화면에 보이는 셀만' },
  { id: 'rowDiff', label: '동일 행에서 값이 다른 셀' },
  { id: 'colDiff', label: '동일 열에서 값이 다른 셀' },
  { id: 'condfmt', label: '조건부 서식' },
  { id: 'validation', label: '데이터 유효성' },
];

/**
 * 범위 rg 에서 조건에 맞는 칸 [[r, c]] (최대 limit 개).
 * cellAt(r, c) → 셀 객체, valueAt(r, c) → 값, types: { numbers, text, logical, errors } (상수 · 수식일 때)
 */
const sameVal = (a, b) => (a ?? '') === (b ?? '') || (typeof a === 'string' && typeof b === 'string' && a.toLowerCase() === b.toLowerCase()) || (isError(a) && isError(b) && a.code === b.code);

export function specialCells(kind, rg, { cellAt, valueAt, hidden, inCond, inValidation, types = null, limit = 1000000, active = null }) {
  const out = [];
  const t = types ?? { numbers: true, text: true, logical: true, errors: true };
  const typeOk = (v) => (typeof v === 'number' ? t.numbers : typeof v === 'string' ? t.text : typeof v === 'boolean' ? t.logical : isError(v) ? t.errors : false);
  for (let r = rg.r1; r <= rg.r2 && out.length < limit; r++) {
    if (kind === 'visible' && hidden?.(r)) continue;
    for (let c = rg.c1; c <= rg.c2 && out.length < limit; c++) {
      const cell = cellAt(r, c);
      let ok = false;
      switch (kind) {
        case 'comments': ok = !!cell?.comment; break;
        case 'constants': ok = !!cell && !cell.formula && (cell.raw !== '' || !!cell.image) && typeOk(valueAt(r, c)); break;
        case 'formulas': ok = !!cell?.formula && typeOk(valueAt(r, c)); break;
        case 'blanks': ok = !cell || (cell.raw === '' && !cell.image); break;
        case 'errors': ok = isError(valueAt(r, c)); break;
        case 'visible': ok = true; break;
        // 행 내용 차이 (Ctrl+\\): 각 행에서 활성 셀과 같은 열의 값과 다른 칸 · 열 내용 차이 (Ctrl+Shift+|)
        case 'rowDiff': ok = c !== (active?.c ?? rg.c1) && !sameVal(valueAt(r, c), valueAt(r, active?.c ?? rg.c1)); break;
        case 'colDiff': ok = r !== (active?.r ?? rg.r1) && !sameVal(valueAt(r, c), valueAt(active?.r ?? rg.r1, c)); break;
        case 'condfmt': ok = !!inCond?.(r, c); break;
        case 'validation': ok = !!inValidation?.(r, c); break;
        default: break;
      }
      if (ok) out.push([r, c]);
    }
  }
  return out;
}
