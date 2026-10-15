// 그림 꾸밈 효과 (PowerPoint [그림 서식 › 꾸밈 효과] 22종) — DOM 없음, RGBA 픽셀 배열에서 계산
// PowerPoint 와 같이 원본은 따로 두고(o.artOrig) 효과를 입힌 결과 그림을 슬라이드에 씀.

/** [키, 이름, pptx 요소 이름, 기본 강도(0…100)] */
export const ART_EFFECTS = [
  ['marker', '표식', 'artisticMarker', 50], ['pencilGrayscale', '연필 회색조', 'artisticPencilGrayscale', 50], ['pencilSketch', '연필 스케치', 'artisticPencilSketch', 50],
  ['lineDrawing', '선 그리기', 'artisticLineDrawing', 50], ['chalkSketch', '분필 스케치', 'artisticChalkSketch', 50], ['paintStrokes', '페인트 스트로크', 'artisticPaintStrokes', 50],
  ['paintBrush', '페인트 브러시', 'artisticPaintBrush', 50], ['glowDiffused', '광채 확산', 'artisticGlowDiffused', 50], ['blur', '흐리게', 'artisticBlur', 30],
  ['lightScreen', '밝은 화면', 'artisticLightScreen', 50], ['watercolorSponge', '수채화 스펀지', 'artisticWatercolorSponge', 50], ['filmGrain', '필름 입자', 'artisticFilmGrain', 50],
  ['mosaicBubbles', '모자이크 방울', 'artisticMosiaicBubbles', 50], ['glass', '유리', 'artisticGlass', 50], ['cement', '시멘트', 'artisticCement', 50],
  ['texturizer', '질감 표현', 'artisticTexturizer', 50], ['crisscrossEtching', '십자형 에칭', 'artisticCrisscrossEtching', 50], ['pastelsSmooth', '파스텔 부드럽게', 'artisticPastelsSmooth', 50],
  ['plasticWrap', '비닐 랩', 'artisticPlasticWrap', 50], ['cutout', '컷아웃', 'artisticCutout', 50], ['photocopy', '사진 복사', 'artisticPhotocopy', 50], ['glowEdges', '광채 가장자리', 'artisticGlowEdges', 50],
];
export const ART_LABEL = Object.fromEntries(ART_EFFECTS.map(([k, l]) => [k, l]));
export const ART_BY_XML = Object.fromEntries(ART_EFFECTS.map(([k, , x]) => [x, k]));
export const ART_XML = Object.fromEntries(ART_EFFECTS.map(([k, , x]) => [k, x]));

// ── 도움 함수 ──
const clamp = (v) => (v < 0 ? 0 : v > 255 ? 255 : v);
function rng(seed = 1) {
  let s = seed >>> 0 || 1;
  return () => { s ^= s << 13; s >>>= 0; s ^= s >>> 17; s ^= s << 5; s >>>= 0; return s / 4294967296; };
}
/** 밝기 (0…255) */
function lumOf(d, n) {
  const g = new Float32Array(n);
  for (let p = 0, i = 0; p < n; p++, i += 4) g[p] = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
  return g;
}
/** 한 채널 상자 흐림 (가로 · 세로 누적합, r 픽셀) */
function boxBlur1(src, w, h, r) {
  if (r < 1) return Float32Array.from(src);
  const tmp = new Float32Array(w * h);
  const out = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    let acc = 0; const row = y * w;
    for (let x = -r; x <= r; x++) acc += src[row + Math.min(w - 1, Math.max(0, x))];
    for (let x = 0; x < w; x++) {
      tmp[row + x] = acc / (2 * r + 1);
      acc += src[row + Math.min(w - 1, x + r + 1)] - src[row + Math.max(0, x - r)];
    }
  }
  for (let x = 0; x < w; x++) {
    let acc = 0;
    for (let y = -r; y <= r; y++) acc += tmp[Math.min(h - 1, Math.max(0, y)) * w + x];
    for (let y = 0; y < h; y++) {
      out[y * w + x] = acc / (2 * r + 1);
      acc += tmp[Math.min(h - 1, y + r + 1) * w + x] - tmp[Math.max(0, y - r) * w + x];
    }
  }
  return out;
}
/** 가우스 흐림 근사 (상자 흐림 3번) */
const blur1 = (src, w, h, r) => { let a = src; const k = Math.max(1, Math.round(r / 2)); for (let i = 0; i < 3; i++) a = boxBlur1(a, w, h, k); return a; };
function channels(d, n) {
  const c = [new Float32Array(n), new Float32Array(n), new Float32Array(n)];
  for (let p = 0, i = 0; p < n; p++, i += 4) { c[0][p] = d[i]; c[1][p] = d[i + 1]; c[2][p] = d[i + 2]; }
  return c;
}
const blurRGB = (c, w, h, r) => c.map((x) => blur1(x, w, h, r));
/** 소벨 경계 세기 (0…약 1) */
function sobel(g, w, h) {
  const out = new Float32Array(w * h);
  const at = (x, y) => g[Math.min(h - 1, Math.max(0, y)) * w + Math.min(w - 1, Math.max(0, x))];
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const gx = -at(x - 1, y - 1) - 2 * at(x - 1, y) - at(x - 1, y + 1) + at(x + 1, y - 1) + 2 * at(x + 1, y) + at(x + 1, y + 1);
    const gy = -at(x - 1, y - 1) - 2 * at(x, y - 1) - at(x + 1, y - 1) + at(x - 1, y + 1) + 2 * at(x, y + 1) + at(x + 1, y + 1);
    out[y * w + x] = Math.min(1, Math.sqrt(gx * gx + gy * gy) / 360);
  }
  return out;
}
/** 돋을새김 (빛 왼쪽 위) -1…1 */
function emboss(g, w, h, k = 1) {
  const out = new Float32Array(w * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const a = g[Math.max(0, y - k) * w + Math.max(0, x - k)];
    const b = g[Math.min(h - 1, y + k) * w + Math.min(w - 1, x + k)];
    out[y * w + x] = Math.max(-1, Math.min(1, (b - a) / 128));
  }
  return out;
}
/** 부드러운 잡음 무늬 (값 0…1) — 칸 크기 cell, 쌍선형 보간 */
function valueNoise(w, h, cell, seed) {
  const r = rng(seed);
  const gw = Math.ceil(w / cell) + 2; const gh = Math.ceil(h / cell) + 2;
  const grid = new Float32Array(gw * gh);
  for (let i = 0; i < grid.length; i++) grid[i] = r();
  const out = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    const fy = y / cell; const y0 = Math.floor(fy); const ty = fy - y0; const sy = ty * ty * (3 - 2 * ty);
    for (let x = 0; x < w; x++) {
      const fx = x / cell; const x0 = Math.floor(fx); const tx = fx - x0; const sx = tx * tx * (3 - 2 * tx);
      const a = grid[y0 * gw + x0]; const b = grid[y0 * gw + x0 + 1]; const c = grid[(y0 + 1) * gw + x0]; const d = grid[(y0 + 1) * gw + x0 + 1];
      out[y * w + x] = a + (b - a) * sx + (c - a) * sy + (a - b - c + d) * sx * sy;
    }
  }
  return out;
}
const levels = (v, n) => Math.round((v / 255) * (n - 1)) * (255 / (n - 1));

/**
 * 꾸밈 효과 계산
 * @param {Uint8ClampedArray} data RGBA (바뀌지 않음)
 * @param {number} w @param {number} h
 * @param {string} key ART_EFFECTS 키
 * @param {{amount?: number}} opts 강도 0…100 (기본 = 효과 기본값)
 * @returns {Uint8ClampedArray} 새 RGBA
 */
export function applyArtEffect(data, w, h, key, opts = {}) {
  const n = w * h;
  const def = ART_EFFECTS.find((e) => e[0] === key);
  const amt = Math.max(0, Math.min(100, opts.amount ?? def?.[3] ?? 50)) / 100;
  const out = new Uint8ClampedArray(data.length);
  // 그림 크기에 맞춘 반지름 (긴 쪽 1000px 기준)
  const S = Math.max(w, h) / 1000;
  const R = (px) => Math.max(1, Math.round(px * Math.max(0.3, S)));
  const put = (p, r, g, b) => { const i = p * 4; out[i] = clamp(r); out[i + 1] = clamp(g); out[i + 2] = clamp(b); out[i + 3] = data[i + 3]; };
  const g = lumOf(data, n);
  const c = channels(data, n);
  const rnd = rng(7);
  switch (key) {
    case 'blur': {
      const b = blurRGB(c, w, h, R(2 + amt * 30));
      for (let p = 0; p < n; p++) put(p, b[0][p], b[1][p], b[2][p]);
      break;
    }
    case 'pencilGrayscale': case 'pencilSketch': {
      // 색 닷지: 밝기 / (1 - 흐린 반전) → 연필 선
      const inv = g.map((v) => 255 - v);
      const bi = blur1(inv, w, h, R(3 + amt * 10));
      const hatch = key === 'pencilSketch' ? valueNoise(w, h, R(2), 3) : null;
      for (let p = 0; p < n; p++) {
        let v = bi[p] >= 255 ? 255 : Math.min(255, (g[p] * 255) / (255 - bi[p]));
        if (hatch) {
          // 어두운 곳에 비스듬한 빗금
          const x = p % w; const y = (p / w) | 0;
          const line = ((x + y) % R(5)) < 1 ? 1 : 0;
          if (v < 200 && line) v *= 0.75 + hatch[p] * 0.2;
        }
        v = 255 - (255 - v) * (0.7 + amt * 0.6);
        put(p, v, v, v);
      }
      break;
    }
    case 'lineDrawing': {
      const e = sobel(blur1(g, w, h, R(1.5)), w, h);
      const th = 0.25 - amt * 0.18;
      for (let p = 0; p < n; p++) { const v = e[p] > th ? 255 - Math.min(255, (e[p] - th) * 900) : 255; put(p, v, v, v); }
      break;
    }
    case 'chalkSketch': {
      const e = sobel(blur1(g, w, h, R(1)), w, h);
      const nz = valueNoise(w, h, R(1.2), 5);
      for (let p = 0; p < n; p++) {
        const k = Math.min(1, e[p] * (1.5 + amt * 2.5) + (g[p] / 255) * 0.35) * (0.55 + nz[p] * 0.6);
        put(p, 40 + k * 215, 52 + k * 203, 60 + k * 195);
      }
      break;
    }
    case 'marker': case 'cutout': case 'paintBrush': case 'watercolorSponge': case 'pastelsSmooth': {
      const rad = { marker: 2, cutout: 4, paintBrush: 3, watercolorSponge: 4, pastelsSmooth: 5 }[key];
      const lv = { marker: 6, cutout: 3 + Math.round((1 - amt) * 4), paintBrush: 7, watercolorSponge: 6, pastelsSmooth: 8 }[key];
      const b = blurRGB(c, w, h, R(rad + amt * 3));
      const e = key === 'pastelsSmooth' || key === 'cutout' ? null : sobel(blur1(g, w, h, R(1)), w, h);
      const paper = key === 'watercolorSponge' ? valueNoise(w, h, R(3), 9) : null;
      for (let p = 0; p < n; p++) {
        let r = levels(b[0][p], lv); let gg = levels(b[1][p], lv); let bb = levels(b[2][p], lv);
        if (key === 'pastelsSmooth') { r = r * 0.7 + 255 * 0.3; gg = gg * 0.7 + 255 * 0.3; bb = bb * 0.7 + 255 * 0.3; }
        if (e) { const k = 1 - Math.min(0.85, e[p] * (key === 'marker' ? 2.2 : 1.4)); r *= k; gg *= k; bb *= k; }
        if (paper) { const k = 0.88 + paper[p] * 0.24; r *= k; gg *= k; bb *= k; }
        put(p, r, gg, bb);
      }
      break;
    }
    case 'paintStrokes': {
      // 비스듬한 방향 흐림 (붓 자국) + 단계화
      const len = R(3 + amt * 8);
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        let r = 0; let gg = 0; let bb = 0; let k = 0;
        const dir = ((x / R(12)) | 0) % 2 ? 1 : -1;
        for (let t = -len; t <= len; t += 2) {
          const xx = Math.min(w - 1, Math.max(0, x + t)); const yy = Math.min(h - 1, Math.max(0, y + t * dir));
          const q = (yy * w + xx) * 4; r += data[q]; gg += data[q + 1]; bb += data[q + 2]; k++;
        }
        put(p, levels(r / k, 10), levels(gg / k, 10), levels(bb / k, 10));
      }
      break;
    }
    case 'glowDiffused': {
      const b = blurRGB(c, w, h, R(6 + amt * 10));
      for (let p = 0; p < n; p++) {
        const grain = rnd() < 0.08 + amt * 0.1 ? 60 : 0;
        // 스크린 혼합 (밝게) + 흰 입자
        put(p, 255 - ((255 - c[0][p]) * (255 - b[0][p])) / 255 + grain, 255 - ((255 - c[1][p]) * (255 - b[1][p])) / 255 + grain, 255 - ((255 - c[2][p]) * (255 - b[2][p])) / 255 + grain);
      }
      break;
    }
    case 'lightScreen': {
      const step = R(4);
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        const on = (x % step === 0) || (y % step === 0);
        const k = on ? 0.35 + amt * 0.35 : 0.15;
        put(p, c[0][p] + (255 - c[0][p]) * k, c[1][p] + (255 - c[1][p]) * k, c[2][p] + (255 - c[2][p]) * k);
      }
      break;
    }
    case 'filmGrain': {
      for (let p = 0; p < n; p++) {
        const nz = (rnd() + rnd() + rnd() - 1.5) * (30 + amt * 90);
        const ct = (v) => (v - 128) * 1.1 + 128 + nz;
        put(p, ct(c[0][p]), ct(c[1][p]), ct(c[2][p]));
      }
      break;
    }
    case 'mosaicBubbles': {
      const cell = R(6 + amt * 14);
      const avg = new Map();
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        const key2 = ((y / cell) | 0) * 100000 + ((x / cell) | 0);
        let a = avg.get(key2);
        if (!a) { a = [0, 0, 0, 0]; avg.set(key2, a); }
        a[0] += c[0][p]; a[1] += c[1][p]; a[2] += c[2][p]; a[3]++;
      }
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        const a = avg.get(((y / cell) | 0) * 100000 + ((x / cell) | 0));
        // 방울: 칸 가운데서 멀수록 어둡고, 왼쪽 위에 반짝임
        const dx = (x % cell) / cell - 0.5; const dy = (y % cell) / cell - 0.5;
        const d = Math.sqrt(dx * dx + dy * dy);
        const shade = d > 0.5 ? 0.7 : 1 - d * 0.5;
        const hl = Math.max(0, 0.25 - Math.hypot(dx + 0.18, dy + 0.18)) * 400;
        put(p, (a[0] / a[3]) * shade + hl, (a[1] / a[3]) * shade + hl, (a[2] / a[3]) * shade + hl);
      }
      break;
    }
    case 'glass': {
      const nx = valueNoise(w, h, R(6), 11); const ny = valueNoise(w, h, R(6), 13);
      const amp = R(2 + amt * 10);
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        const xx = Math.min(w - 1, Math.max(0, Math.round(x + (nx[p] - 0.5) * 2 * amp)));
        const yy = Math.min(h - 1, Math.max(0, Math.round(y + (ny[p] - 0.5) * 2 * amp)));
        const q = yy * w + xx;
        put(p, c[0][q] * 0.92 + 20, c[1][q] * 0.92 + 20, c[2][q] * 0.92 + 20);
      }
      break;
    }
    case 'cement': case 'texturizer': {
      const tex = valueNoise(w, h, key === 'cement' ? R(1.5) : R(3), 17);
      const em = emboss(tex.map((v) => v * 255), w, h, 1);
      const gem = key === 'cement' ? emboss(g, w, h, R(1)) : null;
      for (let p = 0; p < n; p++) {
        const relief = em[p] * (20 + amt * 50) + (gem ? gem[p] * 40 : 0);
        if (key === 'cement') { const v = g[p] * 0.6 + 70 + relief; put(p, v * 0.98 + c[0][p] * 0.1, v * 0.98 + c[1][p] * 0.1, v + c[2][p] * 0.1 - 10); } else put(p, c[0][p] + relief, c[1][p] + relief, c[2][p] + relief);
      }
      break;
    }
    case 'crisscrossEtching': {
      const step = R(4);
      for (let p = 0; p < n; p++) {
        const x = p % w; const y = (p / w) | 0;
        const v = g[p] / 255;
        const d1 = (x + y) % step === 0; const d2 = (x - y + 100000 * step) % step === 0;
        const dark = (v < 0.75 && d1) || (v < 0.45 + amt * 0.1 && d2) || v < 0.15;
        const ink = dark ? 40 : 245;
        put(p, ink, ink, ink - 8);
      }
      break;
    }
    case 'plasticWrap': {
      const em = emboss(blur1(g, w, h, R(3)), w, h, R(2));
      for (let p = 0; p < n; p++) {
        const spec = Math.max(0, em[p] - (0.25 - amt * 0.15)) * 600;
        put(p, c[0][p] * 0.85 + spec, c[1][p] * 0.85 + spec, c[2][p] * 0.85 + spec);
      }
      break;
    }
    case 'photocopy': {
      const b = blur1(g, w, h, R(4 + amt * 6));
      for (let p = 0; p < n; p++) { const v = g[p] - b[p] < -(4 + (1 - amt) * 10) || g[p] < 40 ? 20 : 250; put(p, v, v, v); }
      break;
    }
    case 'glowEdges': {
      const er = sobel(blur1(c[0], w, h, R(1)), w, h); const eg = sobel(blur1(c[1], w, h, R(1)), w, h); const eb = sobel(blur1(c[2], w, h, R(1)), w, h);
      const k = 400 + amt * 800;
      for (let p = 0; p < n; p++) put(p, er[p] * k, eg[p] * k, eb[p] * k);
      break;
    }
    default:
      out.set(data);
  }
  return out;
}

// ───────────── 그림 압축 (PowerPoint [그림 압축]) ─────────────
/** 그림 압축 해상도 [ppi, 이름] — 0 = 원본 해상도 유지 (잘린 영역 삭제만) */
export const COMPRESS_PPI = [[0, '고화질: 원본 해상도 유지'], [330, 'HD (330ppi): 고화질 디스플레이에 적합'], [220, '인쇄 (220ppi): 대부분 프린터와 화면에서 우수한 품질'], [150, '웹 (150ppi): 웹 페이지와 프로젝터에 적합'], [96, '전자 메일 (96ppi): 문서 크기 최소화']];
/**
 * 그림 하나를 어떻게 줄일지 — 원본 픽셀 크기, 자르기, 슬라이드 위 크기(px, 96dpi) 로 계산
 * @returns {{ sx, sy, sw, sh, tw, th, crop: boolean, resize: boolean }} 원본에서 잘라낼 영역(sx…) 과 결과 크기(tw, th)
 */
export function compressPlan({ natW, natH, crop, w, h }, ppi, deleteCrop = true) {
  const c = deleteCrop && crop ? crop : { l: 0, t: 0, r: 0, b: 0 };
  const sx = Math.max(0, Math.round(natW * (c.l || 0)));
  const sy = Math.max(0, Math.round(natH * (c.t || 0)));
  const sw = Math.max(1, Math.round(natW * (1 - (c.l || 0) - (c.r || 0))));
  const sh = Math.max(1, Math.round(natH * (1 - (c.t || 0) - (c.b || 0))));
  let tw = sw; let th = sh;
  if (ppi > 0) {
    // 보이는 부분이 슬라이드에서 차지하는 크기 (자르기 안 지운 경우에는 원본 전체가 그 비율로 늘어남)
    const visW = deleteCrop ? w : w / Math.max(0.01, 1 - (crop?.l || 0) - (crop?.r || 0));
    const visH = deleteCrop ? h : h / Math.max(0.01, 1 - (crop?.t || 0) - (crop?.b || 0));
    const k = Math.min(1, (visW / 96) * ppi / sw, (visH / 96) * ppi / sh);
    tw = Math.max(1, Math.round(sw * k)); th = Math.max(1, Math.round(sh * k));
  }
  const cropped = sx > 0 || sy > 0 || sw < natW || sh < natH;
  return { sx, sy, sw, sh, tw, th, crop: cropped, resize: tw < sw || th < sh };
}
