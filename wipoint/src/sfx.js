// 애니메이션 · 전환 소리 재생 (브라우저): 합성한 소리를 AudioBuffer 로 한 번 만들어 둠
import { synthSound } from './sounds.js';

let ac = null;
const cache = new Map();
/** 소리 재생 (key = SOUNDS 키). 오디오를 쓸 수 없으면 조용히 넘어감 */
export function playSound(key, volume = 0.8) {
  if (!key) return;
  try {
    ac ??= new (window.AudioContext || window.webkitAudioContext)();
    if (ac.state === 'suspended') ac.resume?.();
    let buf = cache.get(key);
    if (!buf) {
      const rate = 22050;
      const data = synthSound(key, rate);
      buf = ac.createBuffer(1, data.length, rate);
      buf.getChannelData(0).set(data);
      cache.set(key, buf);
    }
    const src = ac.createBufferSource();
    const g = ac.createGain();
    g.gain.value = volume;
    src.buffer = buf;
    src.connect(g).connect(ac.destination);
    src.start();
  } catch { /* 소리 없이 */ }
}
