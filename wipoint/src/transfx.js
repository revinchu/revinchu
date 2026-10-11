// 전환 · 애니메이션용 clip-path 모양 (DOM 없음) — 블라인드, 바둑판, 실선 무늬, 디졸브, 시계, 다이아몬드, 더하기 …
// 모든 값은 요소 크기의 % 좌표. 한 효과의 프레임들은 점 개수가 같아서 Web Animations 가 사이를 이어 줌.

const f = (v) => `${Math.round(v * 100) / 100}%`;

/** 사각형 여러 개 → polygon 하나 (각 사각형 뒤에 기준점 (0,0) 으로 돌아가서 이음선 넓이가 0) */
export function rectsPoly(rects) {
  const pts = ['0% 0%'];
  for (const [x1, y1, x2, y2] of rects) pts.push(`${f(x1)} ${f(y1)}`, `${f(x2)} ${f(y1)}`, `${f(x2)} ${f(y2)}`, `${f(x1)} ${f(y2)}`, `${f(x1)} ${f(y1)}`, '0% 0%');
  return `polygon(${pts.join(', ')})`;
}

/** 블라인드: n 줄이 각자 위(왼쪽)부터 펼쳐짐 — t 0…1 */
export function blindsPoly(t, n = 6, vertical = false) {
  const s = 100 / n;
  const rects = [];
  for (let i = 0; i < n; i++) rects.push(vertical ? [i * s, 0, i * s + s * t, 100] : [0, i * s, 100, i * s + s * t]);
  return rectsPoly(rects);
}

/** 빗질: 줄마다 번갈아 왼쪽/오른쪽에서 들어옴 */
export function combPoly(t, n = 8, vertical = false) {
  const s = 100 / n;
  const rects = [];
  for (let i = 0; i < n; i++) {
    const a = i % 2 ? 100 - 100 * t : 0;
    const b = i % 2 ? 100 : 100 * t;
    rects.push(vertical ? [i * s, a, i * s + s, b] : [a, i * s, b, i * s + s]);
  }
  return rectsPoly(rects);
}

/** 바둑판 무늬: 칸이 왼쪽부터 채워지고 홀수 줄은 반 칸 어긋남 */
export function checkerPoly(t, cols = 8, rows = 6) {
  const cw = 100 / cols;
  const rh = 100 / rows;
  const rects = [];
  for (let r = 0; r < rows; r++) {
    const shift = r % 2 ? -cw / 2 : 0;
    for (let c = 0; c <= cols; c++) {
      const x = c * cw + shift;
      // 두 배 너비 동안 채움 (어긋난 칸도 끝까지 덮이게)
      rects.push([x, r * rh, x + cw * Math.min(1, t * 2), r * rh + rh]);
    }
  }
  return rectsPoly(rects);
}

/** 고정 시드 난수 (같은 효과는 늘 같은 무늬) */
function rng(seed) {
  let s = seed >>> 0 || 1;
  return () => { s ^= s << 13; s >>>= 0; s ^= s >>> 17; s ^= s << 5; s >>>= 0; return s / 4294967296; };
}
function shuffled(n, seed) {
  const r = rng(seed);
  const a = Array.from({ length: n }, (_, i) => i);
  for (let i = n - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
}

/** 실선 무늬: 가는 막대가 무작위 순서로 나타남 — 프레임 목록 */
export function randomBarsFrames(frames = 12, n = 48, vertical = false, seed = 7) {
  const order = shuffled(n, seed);
  const at = new Array(n);
  order.forEach((bar, k) => { at[bar] = k / n; });
  const s = 100 / n;
  const out = [];
  for (let k = 0; k <= frames; k++) {
    const t = k / frames;
    const rects = [];
    for (let i = 0; i < n; i++) {
      const on = t >= 1 || at[i] < t ? 1 : 0;
      rects.push(vertical ? [i * s, 0, i * s + s * on, 100] : [0, i * s, 100, i * s + s * on]);
    }
    out.push(rectsPoly(rects));
  }
  return out;
}

/** 디졸브 · 벌집 · 반짝이: 칸이 무작위 순서로 나타남 */
export function randomCellsFrames(frames = 12, cols = 16, rows = 9, seed = 11) {
  const n = cols * rows;
  const order = shuffled(n, seed);
  const at = new Array(n);
  order.forEach((cell, k) => { at[cell] = k / n; });
  const cw = 100 / cols;
  const rh = 100 / rows;
  const out = [];
  for (let k = 0; k <= frames; k++) {
    const t = k / frames;
    const rects = [];
    for (let i = 0; i < n; i++) {
      const x = (i % cols) * cw;
      const y = Math.floor(i / cols) * rh;
      const on = t >= 1 || at[i] < t;
      // 꺼진 칸은 가운데 점 (넓이 0)
      rects.push(on ? [x, y, x + cw, y + rh] : [x + cw / 2, y + rh / 2, x + cw / 2, y + rh / 2]);
    }
    out.push(rectsPoly(rects));
  }
  return out;
}

/** 시계 · 쐐기 · 바퀴: 가운데에서 휩쓸며 나타남. mode 'clock' (12시부터 시계 방향) | 'wedge' (위에서 양쪽으로) */
export function sweepFrames(frames = 16, mode = 'clock', spokes = 1) {
  const N = 48;
  const out = [];
  const R = 150;
  const pt = (a) => `${f(50 + R * Math.sin(a))} ${f(50 - R * Math.cos(a))}`;
  for (let k = 0; k <= frames; k++) {
    const t = k / frames;
    const pts = [];
    if (mode === 'wedge') {
      pts.push('50% 50%');
      for (let i = 0; i <= N; i++) pts.push(pt(-Math.PI * t + (2 * Math.PI * t * i) / N));
    } else {
      // 바퀴살 spokes 개: 각 살이 2π/spokes 만큼 돔
      const per = Math.max(1, Math.floor(N / spokes));
      for (let sp = 0; sp < spokes; sp++) {
        const a0 = (2 * Math.PI * sp) / spokes;
        pts.push('50% 50%');
        for (let i = 0; i <= per; i++) pts.push(pt(a0 + ((2 * Math.PI) / spokes) * t * (i / per)));
      }
      pts.push('50% 50%');
    }
    out.push(`polygon(${pts.join(', ')})`);
  }
  return out;
}

/** 다이아몬드 (t 0 → 점, 1 → 화면 전체) */
export function diamondPoly(t) {
  const r = 100 * t;
  return `polygon(50% ${f(50 - r)}, ${f(50 + r)} 50%, 50% ${f(50 + r)}, ${f(50 - r)} 50%)`;
}
/** 더하기 모양 (가운데에서 커짐) */
export function plusPoly(t) {
  const a = 50 * t; // 팔 반폭
  const b = 50 + 60 * t; // 팔 길이
  const L = 50 - b; const R = 50 + b; const l = 50 - a; const r = 50 + a;
  return `polygon(${[[l, L], [r, L], [r, l], [R, l], [R, r], [r, r], [r, R], [l, R], [l, r], [L, r], [L, l], [l, l]].map(([x, y]) => `${f(x)} ${f(y)}`).join(', ')})`;
}
/** 계단 (대각선): 왼쪽 위 → 오른쪽 아래 */
export function stripsPoly(t, dir = 'rd') {
  const d = 200 * t;
  const tri = { rd: [[0, 0], [d, 0], [0, d]], ld: [[100, 0], [100 - d, 0], [100, d]], ru: [[0, 100], [d, 100], [0, 100 - d]], lu: [[100, 100], [100 - d, 100], [100, 100 - d]] }[dir] ?? [[0, 0], [d, 0], [0, d]];
  return `polygon(${tri.map(([x, y]) => `${f(x)} ${f(y)}`).join(', ')})`;
}

/** 0…1 사이 프레임 목록 */
export function frames(fn, n = 8) {
  return Array.from({ length: n + 1 }, (_, k) => fn(k / n));
}
