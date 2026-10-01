// 셀 단위 의존 그래프 (DOM 없음): 값이 바뀐 칸에서 시작해 그 칸을 (직간접으로) 참조하는 수식만 찾아 다시 계산하게 함.
//
// 아래로 채운 수식은 참조도 규칙적으로 움직인다 (=H2/G2, =H3/G3 … / =SUM($H$2:H2), =SUM($H$2:H3) …).
// 그래서 수식 하나하나가 아니라 "연속된 수식 묶음(run)" 단위로 기억한다:
//   수식 칸: (fr0 + i, fc), i = 0 … n-1
//   참조 범위: 행 r1 + r1s·i ~ r2 + r2s·i (r1s · r2s 는 0 또는 1), 열 c1 ~ c2 (고정)
// 백만 행을 채운 수식 열도 참조 하나당 묶음 하나라서, 그래프를 만들고 찾는 비용이 거의 들지 않는다.
// 바뀐 칸 r 을 참조하는 i 의 범위는 부등식 두 개로 바로 구한다.
//
// 수식이 바뀌거나 지워지면 묶음은 그대로 두고, 찾을 때 그 칸의 수식이 묶음을 만든 수식과 같은지 확인한다(verify).
import { RefValue } from './fxcore.js';
import { FUNCS } from './formula.js';
import { resolveStructRef } from './tables.js';

const R_SPAN = 33554432; // 2^25 행
const C_SPAN = 16384; // 2^14 열
/** 시트 · 행 · 열 → 칸 번호 (2^53 안) */
export const cellNum = (s, r, c) => (s * R_SPAN + r) * C_SPAN + c;

const DYN_FUNCS = new Set(['INDIRECT', 'OFFSET', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'CELL', 'INFO', 'EVALUATE']);
const WIDE = 32; // 이보다 넓은 범위는 시트 목록에 (열마다 넣지 않음)
const LINEAR = 24; // 묶음이 이보다 적으면 정렬 없이 훑음

/** AST 가 참조하는 것 (AST 는 공유되므로 한 번만 훑음) */
const refsMemo = new WeakMap();
export function astRefs(ast) {
  let out = refsMemo.get(ast);
  if (out) return out;
  out = { refs: [], srefs: [], names: [], pivots: [], dyn: false };
  const walk = (n) => {
    if (!n || typeof n !== 'object') return;
    if (Array.isArray(n)) { for (const x of n) walk(x); return; }
    switch (n.type) {
      case 'ref': out.refs.push(n.ref); return;
      case 'sref': out.srefs.push(n); return;
      case 'name': out.names.push(n); return;
      case 'spill': out.dyn = true; break;
      case 'func':
        if (!FUNCS[n.name]) out.names.push({ v: n.name });
        if (DYN_FUNCS.has(n.name)) out.dyn = true;
        // GETPIVOTDATA: 피벗 테이블이 있는 시트 전체에 의존 (피벗 결과는 그 시트에만 쓰임). 참조가 아니면 동적
        else if (n.name === 'GETPIVOTDATA') {
          const a = n.args?.[1];
          if (a?.type === 'ref') out.pivots.push({ node: n, ref: a.ref });
          else out.dyn = true;
        }
        break;
      default: break;
    }
    for (const k in n) {
      if (k === 'ref' || k === 's' || k === 'e') continue;
      const v = n[k];
      if (v && typeof v === 'object') walk(v);
    }
  };
  walk(ast);
  refsMemo.set(ast, out);
  return out;
}

/** 열 하나의 묶음 목록: 정렬된 부분(행 범위 시작 순 + 끝의 누적 최댓값) + 정렬 안 된 꼬리 */
class Bucket {
  constructor() { this.runs = []; this.sorted = 0; this.lo = null; this.maxHi = null; }
  push(run) { this.runs.push(run); }
  /** 행 r 을 덮을 수 있는 묶음마다 fn(run) */
  stab(r, fn) {
    const runs = this.runs;
    if (runs.length - this.sorted > 512 || (this.sorted === 0 && runs.length > LINEAR)) this.sort();
    const S = this.sorted;
    if (S) {
      const lo = this.lo;
      let a = 0;
      let b = S;
      while (a < b) { const m = (a + b) >> 1; if (lo[m] <= r) a = m + 1; else b = m; }
      // a-1 부터 거꾸로: 앞쪽 묶음들의 끝 최댓값이 r 보다 작아지면 멈춤
      for (let i = a - 1; i >= 0 && this.maxHi[i] >= r; i--) { const run = runs[i]; if (run.hi >= r) fn(run); }
    }
    for (let i = S; i < runs.length; i++) { const run = runs[i]; if (run.lo <= r && run.hi >= r) fn(run); }
  }
  sort() {
    const runs = this.runs.filter((x) => !x.dead);
    runs.sort((x, y) => x.lo - y.lo);
    this.runs = runs;
    this.sorted = runs.length;
    this.lo = new Float64Array(runs.length);
    this.maxHi = new Float64Array(runs.length);
    let mx = -1;
    runs.forEach((x, i) => { this.lo[i] = x.lo; mx = Math.max(mx, x.hi); this.maxHi[i] = mx; });
  }
}

export class DepGraph {
  constructor(wb, lazy = false) {
    this.wb = wb;
    this.cols = []; // 대상 시트별 Map(열 → Bucket)
    this.wide = []; // 대상 시트별 넓은 범위 묶음 목록
    this.dyn = new Set(); // 동적 수식 칸 번호
    this.open = new Map(); // 이어 붙일 수 있는 마지막 묶음: 참조 객체 → run
    this.nameMemo = new Map();
    this.added = 0;
    if (!lazy) { const it = this.steps(); while (!it.next().done) { /* 한 번에 */ } }
  }

  /** 그래프 만들기 (나눠서: 수식 약 2만 개마다 멈춤 — 화면이 멈추지 않게 백그라운드로 만들 때) */
  *steps() {
    let n = 0;
    for (let si = 0; si < this.wb.sheets.length; si++) {
      const sheet = this.wb.sheets[si];
      // 열마다 위에서 아래로 (아래로 채운 수식이 한 묶음이 되도록)
      for (const [c, m] of sheet.cells.cols) {
        const rows = [];
        for (const [r, cell] of m) if (cell.formula) rows.push(r);
        rows.sort((a, b) => a - b);
        for (const r of rows) {
          this.add(si, r, c, m.get(r));
          if (++n % 20000 === 0) yield n;
        }
      }
    }
    this.open.clear();
  }

  /** 다시 만드는 편이 나을 만큼 나중에 추가된 것이 많은지 */
  get stale() { return this.added > 200000; }

  /** 참조 하나를 묶음에 추가 (바로 위 칸의 같은 수식 묶음에 이어 붙일 수 있으면 이어 붙임) */
  addRef(tag, ts, R1, C1, R2, C2, fs, fr, fc, cell) {
    const kr = (cell.dr ?? 0) - fr;
    const kc = (cell.dc ?? 0) - fc;
    const last = this.open.get(tag);
    if (last && last.ts === ts && last.s === fs && last.fc === fc && last.fr0 + last.n === fr && last.c1 === C1 && last.c2 === C2
      && last.ast === cell.ast && last.kr === kr && last.kc === kc) {
      if (last.n === 1) {
        const a = R1 - last.r1;
        const b = R2 - last.r2;
        if ((a === 0 || a === 1) && (b === 0 || b === 1)) {
          last.r1s = a;
          last.r2s = b;
          last.n = 2;
          this.grow(last, R2);
          return;
        }
      } else if (R1 === last.r1 + last.r1s * last.n && R2 === last.r2 + last.r2s * last.n) {
        last.n++;
        this.grow(last, R2);
        return;
      }
    }
    const run = { ts, s: fs, fc, fr0: fr, n: 1, r1: R1, r2: R2, r1s: 0, r2s: 0, c1: C1, c2: C2, lo: R1, hi: R2, ast: cell.ast, kr, kc, buckets: null };
    this.open.set(tag, run);
    if (C2 - C1 >= WIDE) { (this.wide[ts] ??= []).push(run); return; }
    let cols = this.cols[ts];
    if (!cols) { cols = new Map(); this.cols[ts] = cols; }
    run.buckets = [];
    for (let c = C1; c <= C2; c++) {
      let b = cols.get(c);
      if (!b) { b = new Bucket(); cols.set(c, b); }
      b.push(run);
      run.buckets.push(b);
    }
  }

  /** 묶음이 아래로 늘어남: 이미 정렬해 둔 열 목록은 다시 정렬하게 함 (끝 최댓값이 바뀜) */
  grow(run, R2) {
    if (R2 <= run.hi) return;
    run.hi = R2;
    if (run.buckets) for (const b of run.buckets) if (b.sorted) b.sorted = 0;
  }

  /** 수식 셀 등록 */
  add(si, r, c, cell) {
    this.refBoxes(si, r, c, cell, (ts, R1, C1, R2, C2, tag) => {
      if (ts === null) this.dyn.add(cellNum(si, r, c));
      else this.addRef(tag, ts, R1, C1, R2, C2, si, r, c, cell);
    });
  }

  /**
   * 수식 셀이 참조하는 범위마다 cb(대상 시트, r1, c1, r2, c2, 참조 노드). 동적이면 cb(null)
   * (참조되는 셀 추적 · 그래프 만들기가 같이 씀)
   */
  refBoxes(si, r, c, cell, cb) {
    // 해석할 수 없는 수식은 참조를 추측하지 않고 모든 입력 변경에 의존시킨다.
    if (!cell.ast) { cb(null); return; }
    const info = astRefs(cell.ast);
    const wb = this.wb;
    if (info.dyn) cb(null);
    const dr = cell.dr ?? 0;
    const dc = cell.dc ?? 0;
    for (const f of info.refs) {
      let s = si;
      if (f.sheet) {
        if (f.sheet.includes(':')) { cb(null); continue; }
        s = wb.sheetIndexByName(f.sheet);
        if (s < 0) continue; // 없는 시트 → #REF!
      }
      const r1 = f.ar1 || f.cols ? f.r1 : f.r1 + dr;
      const c1 = f.ac1 || f.rows ? f.c1 : f.c1 + dc;
      const r2 = f.ar2 || f.cols ? f.r2 : f.r2 + dr;
      const c2 = f.ac2 || f.rows ? f.c2 : f.c2 + dc;
      cb(s, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2), f);
    }
    for (const p of info.pivots) {
      const f = p.ref;
      const s = f.sheet ? (f.sheet.includes(':') ? -2 : wb.sheetIndexByName(f.sheet)) : si;
      if (s === -2) { cb(null); continue; }
      if (s < 0) continue;
      // 참조가 든 피벗 테이블의 열 범위 전체 (피벗 결과는 그 영역에만 쓰임). 피벗이 없으면 시트 전체
      const r0 = f.ar1 ? f.r1 : f.r1 + dr;
      const c0 = f.ac1 ? f.c1 : f.c1 + dc;
      const sh = wb.sheets[s];
      const a = [sh.pivot, ...(sh.pivotsExtra ?? [])].find((d) => d?.area && r0 >= d.area.r1 && r0 <= d.area.r2 && c0 >= d.area.c1 && c0 <= d.area.c2)?.area;
      if (a) cb(s, 0, a.c1, R_SPAN - 1, a.c2, p.node);
      else cb(s, 0, 0, R_SPAN - 1, C_SPAN - 1, p.node);
    }
    if (info.srefs.length) {
      const here = { si, r, c, sheet: wb.sheets[si]?.name };
      for (const n of info.srefs) {
        let rg = null;
        try { rg = resolveStructRef(wb, n.table, n.spec, here); } catch { rg = null; }
        if (rg) cb(rg.si, rg.r1, rg.c1, rg.r2, rg.c2, n);
      }
    }
    for (const n of info.names) {
      const d = this.nameInfo(si, r, c, n);
      if (d.dyn) cb(null);
      else if (d.box) cb(d.box.si, d.box.r1, d.box.c1, d.box.r2, d.box.c2, n);
    }
  }

  /** 이름: 범위 이름은 그 범위, 상수는 없음, 그 밖(수식 · LAMBDA · 상대 참조 이름)은 동적 */
  nameInfo(si, r, c, n) {
    const wb = this.wb;
    const mk = `${si}\u0001${n.sheet ?? ''}\u0001${n.v.toLowerCase()}`;
    let d = this.nameMemo.get(mk);
    if (d === undefined) {
      const def = wb.findName(n.v, si, n.sheet);
      if (!def) {
        // 정의된 이름이 아니면 LET 변수 · LAMBDA 인수 · 함수 이름 · 표 이름
        let rg = null;
        try { rg = resolveStructRef(wb, n.v, '', { si, r, c, sheet: wb.sheets[si]?.name }); } catch { rg = null; }
        d = rg ? { box: rg } : { none: true };
      } else {
        let v;
        try { v = wb.nameValue(n.v, si, n.sheet, null); } catch { v = undefined; }
        if (v instanceof RefValue) {
          const s = v.sheet ? wb.sheetIndexByName(v.sheet) : si;
          const rel = /(^|[^$A-Za-z])[A-Za-z]{1,3}\d/.test(def.ref ?? '');
          d = s < 0 ? { none: true } : rel ? { dyn: true } : { box: { si: s, r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 } };
        } else if (v === null || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') d = { none: true };
        else d = { dyn: true };
      }
      this.nameMemo.set(mk, d);
    }
    return d;
  }

  /** 칸 (si, r, c) 의 수식이 바뀜: 새 수식의 참조만 추가 (옛 묶음은 찾을 때 확인해서 무시) */
  set(si, r, c, cell) {
    this.dyn.delete(cellNum(si, r, c));
    if (cell?.formula) { this.added++; this.add(si, r, c, cell); }
  }

  /** 행 r 에 걸리는 묶음 run 의 수식 i 범위 → 수식마다 fn(시트, 행, 열) */
  static hits(run, r, fn, cover = null) {
    let iLo = 0;
    let iHi = run.n - 1;
    if (run.r2s) iLo = Math.max(iLo, r - run.r2); else if (run.r2 < r) return;
    if (run.r1s) iHi = Math.min(iHi, r - run.r1); else if (run.r1 > r) return;
    if (iLo > iHi) return;
    if (cover) {
      // 한 번의 전파에서 이미 훑은 구간은 건너뜀: 고정 범위를 참조하는 수식 묶음(=COUNTIF($L$5:$L$4000,…) 을 채운 열)은
      // 바뀐 칸이 수만 개여도 묶음 전체를 한 번만 훑음 (O(바뀐 칸 × 묶음 길이) → O(바뀐 칸 + 묶음 길이))
      const cv = cover.get(run);
      if (!cv) cover.set(run, [iLo, iHi]);
      else if (iLo >= cv[0] && iHi <= cv[1]) return;
      else if (iHi >= cv[0] - 1 && iLo <= cv[1] + 1) {
        for (let i = iLo; i < cv[0]; i++) fn(run, run.fr0 + i);
        for (let i = Math.max(iLo, cv[1] + 1); i <= iHi; i++) fn(run, run.fr0 + i);
        cv[0] = Math.min(cv[0], iLo);
        cv[1] = Math.max(cv[1], iHi);
        return;
      }
    }
    for (let i = iLo; i <= iHi; i++) fn(run, run.fr0 + i);
  }

  /** 묶음의 수식이 아직 그대로인지 (지우거나 바꾼 수식이면 무시) */
  verify(run, fr) {
    const cell = this.wb.sheets[run.s]?.cells.getRC(fr, run.fc);
    return !!cell && cell.ast === run.ast && (cell.dr ?? 0) - fr === run.kr && (cell.dc ?? 0) - run.fc === run.kc;
  }

  /** 칸 (s, r, c) 를 참조하는 묶음마다 fn(run, 수식 행) */
  each(s, r, c, fn, cover = null) {
    const b = this.cols[s]?.get(c);
    if (b) b.stab(r, (run) => DepGraph.hits(run, r, fn, cover));
    const w = this.wide[s];
    if (w) for (const run of w) if (c >= run.c1 && c <= run.c2 && run.lo <= r && run.hi >= r) DepGraph.hits(run, r, fn, cover);
  }

  /**
   * 바뀐 칸 목록(points: [s, r, c, …] 평평한 배열)에서 시작해 다시 계산할 수식 칸 목록 [s, r, c, …]
   */
  propagate(points, limit = Infinity) {
    const seen = new Set();
    const out = [];
    const queue = points.slice();
    const over = {};
    const mark = (run, fr) => {
      const n = cellNum(run.s, fr, run.fc);
      if (seen.has(n) || !this.verify(run, fr)) return;
      // 다시 계산할 수식이 너무 많음: 칸마다 표시하기보다 시트 단위가 빠름 (호출한 쪽이 시트 단위로 처리)
      if (seen.size >= limit) throw over;
      seen.add(n);
      out.push(run.s, fr, run.fc);
      queue.push(run.s, fr, run.fc);
    };
    // 동적 수식은 무엇이 바뀌어도 다시 계산
    for (const n of this.dyn) {
      if (seen.has(n)) continue;
      const c = n % C_SPAN;
      const rest = (n - c) / C_SPAN;
      const r = rest % R_SPAN;
      const s = (rest - r) / R_SPAN;
      if (!this.wb.sheets[s]?.cells.getRC(r, c)?.formula) continue;
      seen.add(n);
      out.push(s, r, c);
      queue.push(s, r, c);
    }
    const cover = new Map();
    try {
      for (let q = 0; q < queue.length; q += 3) this.each(queue[q], queue[q + 1], queue[q + 2], mark, cover);
    } catch (e) {
      if (e === over) return null;
      throw e;
    }
    return out;
  }

  /** 칸 (s, r, c) 를 바로 참조하는 수식 [{ si, r, c }] (참조하는 셀 추적용) */
  dependentsOf(s, r, c) {
    const res = new Map();
    this.each(s, r, c, (run, fr) => { if (this.verify(run, fr)) res.set(cellNum(run.s, fr, run.fc), { si: run.s, r: fr, c: run.fc }); });
    return [...res.values()];
  }
}
