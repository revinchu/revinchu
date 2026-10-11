// 애니메이션 · 전환 소리 (PowerPoint [소리] 목록) — 파일 없이 합성 (DOM 없음). pptx 에는 WAV 로 넣음.
import { encodeWav } from './recording.js';

/** [키, 이름, pptx 파일 이름] */
export const SOUNDS = [
  ['applause', '박수', 'applause.wav'], ['arrow', '화살', 'arrow.wav'], ['bomb', '폭탄', 'bomb.wav'], ['breeze', '바람', 'breeze.wav'],
  ['camera', '카메라', 'camera.wav'], ['cashreg', '금전 등록기', 'cashreg.wav'], ['chime', '차임', 'chime.wav'], ['click', '클릭', 'click.wav'],
  ['coin', '동전', 'coin.wav'], ['drumroll', '드럼롤', 'drumroll.wav'], ['explode', '폭발', 'explode.wav'], ['hammer', '망치', 'hammer.wav'],
  ['laser', '레이저', 'laser.wav'], ['push', '밀기', 'push.wav'], ['suction', '흡입', 'suction.wav'], ['typewriter', '타자기', 'type.wav'],
  ['voltage', '전압', 'voltage.wav'], ['whoosh', '휙', 'whoosh.wav'], ['wind', '바람 소리', 'wind.wav'],
];
export const SOUND_LABEL = Object.fromEntries(SOUNDS.map(([k, l]) => [k, l]));
export const SOUND_BY_FILE = Object.fromEntries(SOUNDS.map(([k, , f]) => [f, k]));

function rng(seed) {
  let s = seed >>> 0 || 1;
  return () => { s ^= s << 13; s >>>= 0; s ^= s >>> 17; s ^= s << 5; s >>>= 0; return s / 4294967296 * 2 - 1; };
}

/** 소리 합성 → Float32Array (-1…1, 모노) */
export function synthSound(key, rate = 22050) {
  const len = { applause: 1.6, drumroll: 1.4, breeze: 1.2, wind: 1.5, typewriter: 0.9, chime: 1.2, cashreg: 0.9, bomb: 1.3, explode: 1.2, voltage: 0.8, laser: 0.5, coin: 0.6, whoosh: 0.5, suction: 0.6, push: 0.4, arrow: 0.45, camera: 0.35, hammer: 0.5, click: 0.08 }[key] ?? 0.5;
  const n = Math.round(len * rate);
  const out = new Float32Array(n);
  const noise = rng(key.length * 7919 + key.charCodeAt(0));
  const tone = (f, t) => Math.sin(2 * Math.PI * f * t);
  // 간단한 1차 저역 통과 (바람 · 폭발)
  let lp = 0;
  const low = (x, a) => (lp += a * (x - lp));
  for (let i = 0; i < n; i++) {
    const t = i / rate;
    const u = t / len;
    let v = 0;
    switch (key) {
      case 'applause': { // 여러 박수 소리 (짧은 잡음 덩어리)
        const clap = ((i * 2654435761) >>> 0) % 997 < 9 ? 1 : 0;
        v = noise() * (0.25 + clap * 0.75) * Math.exp(-((i % 1103) / rate) * 40) * (1 - u * 0.6);
        break;
      }
      case 'drumroll': v = noise() * Math.exp(-((t * 28) % 1) * 6) * (0.5 + u * 0.5); break;
      case 'breeze': case 'wind': v = low(noise(), 0.05 + 0.04 * tone(0.7, t)) * 3 * Math.sin(Math.PI * u); break;
      case 'typewriter': v = (((t * 11) % 1) < 0.04 ? noise() : 0) + (u > 0.85 ? tone(1900, t) * 0.4 * Math.exp(-(u - 0.85) * 40) : 0); break;
      case 'chime': v = (tone(1318, t) * 0.5 + tone(1568, t) * 0.3 + tone(2093, t) * 0.2) * Math.exp(-t * 3); break;
      case 'coin': v = (t < 0.08 ? tone(988, t) : tone(1319, t)) * Math.exp(-t * 6) * 0.7; break;
      case 'cashreg': v = (t < 0.15 ? noise() * 0.6 * Math.exp(-t * 30) : 0) + tone(2637, t) * Math.exp(-Math.max(0, t - 0.15) * 5) * (t > 0.15 ? 0.5 : 0); break;
      case 'bomb': case 'explode': {
        const whistle = key === 'bomb' && t < 0.6 ? tone(1800 - t * 2200, t) * 0.3 : 0;
        const boomT = key === 'bomb' ? t - 0.6 : t;
        v = whistle + (boomT > 0 ? low(noise(), 0.08) * 4 * Math.exp(-boomT * 4) : 0);
        break;
      }
      case 'voltage': v = Math.sign(tone(60, t)) * 0.3 * (0.6 + 0.4 * noise()) * Math.sin(Math.PI * u); break;
      case 'laser': v = tone(2200 * Math.exp(-t * 6), t) * Math.exp(-t * 4) * 0.6; break;
      case 'whoosh': case 'arrow': v = low(noise(), 0.1 + 0.3 * Math.sin(Math.PI * u)) * 2.5 * Math.sin(Math.PI * u) + (key === 'arrow' && u > 0.85 ? noise() * 0.6 * Math.exp(-(u - 0.85) * 60) : 0); break;
      case 'suction': v = low(noise(), 0.02 + 0.3 * u) * 2 * Math.sin(Math.PI * u); break;
      case 'push': v = low(noise(), 0.15) * 2 * Math.exp(-t * 8); break;
      case 'camera': v = (t < 0.03 || (t > 0.18 && t < 0.22) ? noise() * 0.9 : 0); break;
      case 'hammer': v = (noise() * 0.5 + tone(180, t)) * Math.exp(-t * 18) * 0.8; break;
      case 'click': v = noise() * Math.exp(-t * 120); break;
      default: v = tone(880, t) * Math.exp(-t * 6);
    }
    // 시작 · 끝 잘림 소리 막기
    const fade = Math.min(1, i / 64, (n - i) / 256);
    out[i] = Math.max(-1, Math.min(1, v * 0.8 * fade));
  }
  return out;
}

/** 소리 → WAV 바이트 */
export function soundWav(key, rate = 22050) {
  return encodeWav(synthSound(key, rate), rate);
}
