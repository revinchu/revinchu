// 쇼 녹화 도움 함수 (DOM 없음): WAV 만들기, 녹화한 잉크를 시간에 맞춰 보여 주기, 자막 줄 나누기

/** 여러 채널 → 모노 (평균) */
export function mixDown(channels) {
  if (channels.length === 1) return channels[0];
  const n = Math.min(...channels.map((c) => c.length));
  const out = new Float32Array(n);
  for (const c of channels) for (let i = 0; i < n; i++) out[i] += c[i] / channels.length;
  return out;
}

/** 표본 수 바꾸기 (선형 보간) — 내레이션은 22.05kHz 면 충분 */
export function resample(samples, from, to) {
  if (from === to) return samples;
  const n = Math.max(1, Math.round((samples.length * to) / from));
  const out = new Float32Array(n);
  const k = from / to;
  for (let i = 0; i < n; i++) {
    const x = i * k; const a = Math.floor(x); const t = x - a;
    out[i] = (samples[a] ?? 0) * (1 - t) + (samples[Math.min(samples.length - 1, a + 1)] ?? 0) * t;
  }
  return out;
}

/** 16비트 PCM 모노 WAV (PowerPoint · 모든 재생기에서 열림) */
export function encodeWav(samples, sampleRate) {
  const n = samples.length;
  const buf = new ArrayBuffer(44 + n * 2);
  const v = new DataView(buf);
  const str = (o, s) => { for (let i = 0; i < s.length; i++) v.setUint8(o + i, s.charCodeAt(i)); };
  str(0, 'RIFF'); v.setUint32(4, 36 + n * 2, true); str(8, 'WAVE');
  str(12, 'fmt '); v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true);
  v.setUint32(24, sampleRate, true); v.setUint32(28, sampleRate * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
  str(36, 'data'); v.setUint32(40, n * 2, true);
  for (let i = 0; i < n; i++) { const s = Math.max(-1, Math.min(1, samples[i])); v.setInt16(44 + i * 2, s < 0 ? s * 0x8000 : s * 0x7fff, true); }
  return new Uint8Array(buf);
}

/**
 * 녹화한 잉크 중 t 초까지 그려진 부분 — strokes: [{ color, width, pts: [[x, y]], ts: [초] }]
 * @returns 같은 모양의 목록 (점이 하나도 없는 획은 빠짐)
 */
export function inkUntil(strokes, t) {
  const out = [];
  for (const s of strokes ?? []) {
    const ts = s.ts ?? [];
    let k = 0;
    while (k < s.pts.length && (ts[k] ?? 0) <= t) k++;
    if (k) out.push({ ...s, pts: s.pts.slice(0, k), ts: ts.slice(0, k) });
  }
  return out;
}

/** 자막: 긴 글은 마지막 max 글자만 두 줄 안에 (말하는 동안 계속 바뀜) */
export function captionText(text, max = 90) {
  const t = String(text ?? '').replace(/\s+/g, ' ').trim();
  if (t.length <= max) return t;
  const cut = t.slice(-max);
  const sp = cut.indexOf(' ');
  return sp > 0 && sp < 20 ? cut.slice(sp + 1) : cut;
}

/** 자막 언어 (Web Speech API 언어 태그) */
export const CAPTION_LANGS = [['ko-KR', '한국어'], ['en-US', '영어 (미국)'], ['en-GB', '영어 (영국)'], ['ja-JP', '일본어'], ['zh-CN', '중국어 (간체)'], ['zh-TW', '중국어 (번체)'], ['es-ES', '스페인어'], ['fr-FR', '프랑스어'], ['de-DE', '독일어'], ['vi-VN', '베트남어']];
/** 자막 위치 (PowerPoint [자막 설정 › 위치]) */
export const CAPTION_POS = [['bottom', '아래쪽 (겹침)'], ['top', '위쪽 (겹침)'], ['below', '슬라이드 아래'], ['above', '슬라이드 위']];
/** 자막 번역 대상 (Chrome 내장 Translator API, 이 기기에서 번역) — '' = 번역 안 함 */
export const CAPTION_TARGETS = [['', '번역 안 함 (말한 언어 그대로)'], ['ko', '한국어'], ['en', '영어'], ['ja', '일본어'], ['zh', '중국어'], ['es', '스페인어'], ['fr', '프랑스어'], ['de', '독일어'], ['vi', '베트남어'], ['th', '태국어'], ['ru', '러시아어']];
