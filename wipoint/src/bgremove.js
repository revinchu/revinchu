// 그림 배경 제거 (DOM 없음): 가장자리에서 이어진 배경색 영역을 찾아 투명하게
//   PowerPoint [배경 제거] 처럼 [제거할 영역 표시] (remove 점) / [보관할 영역 표시] (keep 점) 로 고칠 수 있음

/** 가장자리 픽셀의 대표 색 (많이 나온 색 묶음 최대 6개) */
export function borderColors(data, w, h) {
  const buckets = new Map();
  const add = (x, y) => {
    const i = (y * w + x) * 4;
    if (data[i + 3] < 16) return;
    const k = ((data[i] >> 4) << 8) | ((data[i + 1] >> 4) << 4) | (data[i + 2] >> 4);
    const b = buckets.get(k) ?? { n: 0, r: 0, g: 0, b: 0 };
    b.n++; b.r += data[i]; b.g += data[i + 1]; b.b += data[i + 2];
    buckets.set(k, b);
  };
  for (let x = 0; x < w; x++) { add(x, 0); add(x, h - 1); }
  for (let y = 1; y < h - 1; y++) { add(0, y); add(w - 1, y); }
  const list = [...buckets.values()].sort((a, b) => b.n - a.n);
  const total = list.reduce((s, b) => s + b.n, 0) || 1;
  const out = [];
  let cov = 0;
  for (const b of list) {
    if (out.length >= 6 || (cov / total > 0.85 && out.length)) break;
    // 가장자리의 2% 도 안 되는 색은 배경이 아님 (가장자리에 걸친 피사체)
    if (b.n / total < 0.02 && out.length) break;
    out.push([b.r / b.n, b.g / b.n, b.b / b.n]);
    cov += b.n;
  }
  return out;
}

const dist = (data, i, c) => {
  const dr = data[i] - c[0]; const dg = data[i + 1] - c[1]; const db = data[i + 2] - c[2];
  // 사람 눈에 맞춘 가중 거리 (0 … 약 255)
  return Math.sqrt((dr * dr * 3 + dg * dg * 4 + db * db * 2) / 9);
};

/**
 * 알파 마스크 계산 — 0 = 배경(투명), 255 = 남김
 * @param {Uint8ClampedArray} data RGBA
 * @param {{tolerance?: number, remove?: number[][], keep?: number[][], brush?: number}} opts
 *   tolerance 0–100 (기본 25), remove/keep = [[x, y]] 픽셀 좌표, brush = 보관 표시 반지름(px)
 */
export function backgroundMask(data, w, h, opts = {}) {
  const tol = 6 + (opts.tolerance ?? 25) * 1.6;
  const brush = opts.brush ?? Math.max(3, Math.round(Math.min(w, h) / 60));
  const colors = borderColors(data, w, h);
  const n = w * h;
  const bg = new Uint8Array(n);       // 1 = 배경
  const keep = new Uint8Array(n);     // 1 = 보관 (채우기가 넘어가지 못함)
  for (const [kx, ky] of opts.keep ?? []) {
    for (let y = Math.max(0, ky - brush); y <= Math.min(h - 1, ky + brush); y++) {
      for (let x = Math.max(0, kx - brush); x <= Math.min(w - 1, kx + brush); x++) if ((x - kx) ** 2 + (y - ky) ** 2 <= brush * brush) keep[y * w + x] = 1;
    }
  }
  const near = (p, list) => { let m = Infinity; for (const c of list) { const d = dist(data, p * 4, c); if (d < m) m = d; } return m; };
  const queue = new Int32Array(n);
  let qh = 0; let qt = 0;
  const fill = (seeds, list) => {
    qh = qt = 0;
    for (const p of seeds) if (!bg[p] && !keep[p] && (data[p * 4 + 3] < 16 || near(p, list) <= tol)) { bg[p] = 1; queue[qt++] = p; }
    while (qh < qt) {
      const p = queue[qh++];
      const x = p % w;
      const nb = [x > 0 ? p - 1 : -1, x < w - 1 ? p + 1 : -1, p >= w ? p - w : -1, p < n - w ? p + w : -1];
      for (const q of nb) {
        if (q < 0 || bg[q] || keep[q]) continue;
        if (data[q * 4 + 3] < 16 || near(q, list) <= tol) { bg[q] = 1; queue[qt++] = q; }
      }
    }
  };
  // 1) 가장자리에서 시작
  if (colors.length) {
    const seeds = [];
    for (let x = 0; x < w; x++) seeds.push(x, (h - 1) * w + x);
    for (let y = 1; y < h - 1; y++) seeds.push(y * w, y * w + w - 1);
    fill(seeds, colors);
  }
  // 2) [제거할 영역 표시] 점: 그 점의 색으로 이어진 영역
  for (const [rx, ry] of opts.remove ?? []) {
    const p = Math.min(h - 1, Math.max(0, ry | 0)) * w + Math.min(w - 1, Math.max(0, rx | 0));
    const i = p * 4;
    keep[p] = 0;
    fill([p], [[data[i], data[i + 1], data[i + 2]], ...colors]);
  }
  // 3) 경계는 부드럽게: 배경 옆 남김 픽셀은 배경색과 가까운 만큼 반투명
  const alpha = new Uint8Array(n);
  for (let p = 0; p < n; p++) {
    if (bg[p]) continue;
    const x = p % w;
    const edge = (x > 0 && bg[p - 1]) || (x < w - 1 && bg[p + 1]) || (p >= w && bg[p - w]) || (p < n - w && bg[p + w]);
    if (edge && !keep[p] && colors.length) {
      const d = near(p, colors);
      alpha[p] = d >= tol * 2 ? 255 : Math.round(255 * Math.max(0.25, (d - tol) / tol));
    } else alpha[p] = 255;
  }
  return alpha;
}

/** 마스크를 RGBA 에 적용 (원본 알파와 곱함) */
export function applyMask(data, alpha) {
  const out = new Uint8ClampedArray(data);
  for (let p = 0; p < alpha.length; p++) out[p * 4 + 3] = Math.round((data[p * 4 + 3] * alpha[p]) / 255);
  return out;
}
