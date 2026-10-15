// 빠른 채우기 (Ctrl+E): 예시에서 규칙을 찾아 나머지 행을 채움 (DOM 없음)
// 규칙 = 조각들의 연결. 조각: 원본 열의 전체/토큰(앞·뒤에서 n번째)/토큰 첫 글자 (대소문자 변환 포함) 또는 고정 글자

const SPLIT = /[^\p{L}\p{N}]+/u;

function tokensOf(s) {
  const parts = String(s).split(SPLIT).filter(Boolean);
  // 숫자·글자 경계도 나눔 (abc123 → abc, 123)
  const fine = parts.flatMap((p) => p.match(/\d+|[^\d]+/g) ?? []);
  return { parts, fine };
}

const CASES = ['same', 'upper', 'lower', 'proper'];
function applyCase(s, c) {
  if (c === 'upper') return s.toUpperCase();
  if (c === 'lower') return s.toLowerCase();
  if (c === 'proper') return s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a, b) => a + b.toUpperCase());
  return s;
}

/** 조각 → 글자 (없으면 null) */
function render(piece, sources) {
  if (piece.lit !== undefined) return piece.lit;
  const src = String(sources[piece.col] ?? '');
  let base;
  if (piece.kind === 'whole') base = src;
  else {
    const t = tokensOf(src)[piece.set];
    const i = piece.idx >= 0 ? piece.idx : t.length + piece.idx;
    if (i < 0 || i >= t.length) return null;
    base = t[i];
    if (piece.kind === 'first') base = base[0] ?? '';
  }
  return applyCase(base, piece.case);
}

/** 한 예시에서 가능한 조각 후보 */
function candidates(sources) {
  const out = [];
  sources.forEach((s, col) => {
    const str = String(s ?? '');
    if (!str) return;
    for (const c of CASES) out.push({ col, kind: 'whole', case: c });
    const t = tokensOf(str);
    for (const set of ['parts', 'fine']) {
      const n = t[set].length;
      for (let i = 0; i < n; i++) {
        for (const idx of [i, i - n]) {
          for (const c of CASES) {
            out.push({ col, kind: 'tok', set, idx, case: c });
            out.push({ col, kind: 'first', set, idx, case: c });
          }
        }
      }
    }
  });
  return out;
}

/** 대상 글자를 조각들로 나누기: 구간마다 가능한 조각 목록 */
function segment(sources, target) {
  const cands = candidates(sources).map((p) => ({ p, v: render(p, sources) })).filter((x) => x.v);
  const segs = [];
  let pos = 0;
  while (pos < target.length) {
    let best = 0;
    for (const { v } of cands) if (v.length > best && target.startsWith(v, pos)) best = v.length;
    // 한 글자짜리 조각은 고정 글자보다 약하게: 구분 기호 같은 글자는 고정 글자로
    if (best >= 1 && !(best === 1 && /[^\p{L}\p{N}]/u.test(target[pos]))) {
      const alts = cands.filter(({ v }) => v.length === best && target.startsWith(v, pos)).map(({ p }) => p);
      segs.push({ alts });
      pos += best;
    } else {
      const last = segs[segs.length - 1];
      if (last?.lit !== undefined) last.lit += target[pos];
      else segs.push({ lit: target[pos] });
      pos++;
    }
  }
  return segs;
}

const run = (prog, sources) => {
  let s = '';
  for (const p of prog) {
    const v = render(p, sources);
    if (v === null) return null;
    s += v;
  }
  return s;
};

// 선호 순서: 앞에서 센 토큰, 원래 대소문자, 굵은 토큰(parts) 먼저
const score = (p) => (p.kind === 'whole' ? 0 : 1) + (p.case === 'same' ? 0 : 2) + (p.set === 'fine' ? 1 : 0) + (p.idx < 0 ? 0.5 : 0) + (p.kind === 'first' ? 0.3 : 0);

/**
 * examples: [{ sources: [원본 열 값…], target }], inputs: [[원본 열 값…]]
 * 반환: 채울 값 배열 (규칙을 못 찾으면 null)
 */
export function flashFill(examples, inputs) {
  if (!examples.length) return null;
  for (const ex of examples) {
    const segs = segment(ex.sources, String(ex.target));
    const alts = segs.map((sg) => (sg.lit !== undefined ? [{ lit: sg.lit }] : [...sg.alts].sort((a, b) => score(a) - score(b)).slice(0, 12)));
    // 모든 예시에 맞는 조합 찾기 (앞쪽 선호 순)
    let tries = 0;
    const pick = [];
    const search = (i) => {
      if (tries++ > 20000) return false;
      if (i === alts.length) return examples.every((e) => run(pick, e.sources) === String(e.target));
      for (const a of alts[i]) {
        pick.push(a);
        if (search(i + 1)) return true;
        pick.pop();
      }
      return false;
    };
    if (search(0)) {
      const prog = [...pick];
      return inputs.map((src) => run(prog, src));
    }
  }
  return null;
}
