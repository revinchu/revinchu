// 문서 암호 풀기에 쓰는 암호 함수 (DOM 없음, 의존성 없음)
//  SHA-1 · SHA-256 · HMAC · PBKDF2 · AES(128/256, ECB · CBC · CFB-1) · DES(ECB)
//  배포용 · 암호 한글 문서(HWP 3 · HWP 5 · HWPX) 를 여는 데만 씀. 표준(FIPS 180-4 · 197 · 46-3, RFC 2898) 대로 구현.

// ───────────── SHA ─────────────
const rotl = (x, n) => (x << n) | (x >>> (32 - n));
const rotr = (x, n) => (x >>> n) | (x << (32 - n));
function pad64(msg, bigLen = 8) {
  const len = msg.length;
  const total = Math.ceil((len + 1 + bigLen) / 64) * 64;
  const b = new Uint8Array(total);
  b.set(msg); b[len] = 0x80;
  const bits = len * 8;
  const dv = new DataView(b.buffer);
  dv.setUint32(total - 4, bits >>> 0); dv.setUint32(total - 8, Math.floor(bits / 2 ** 32));
  return { b, dv };
}
export function sha1(msg) {
  const { b, dv } = pad64(msg);
  let h0 = 0x67452301, h1 = 0xefcdab89, h2 = 0x98badcfe, h3 = 0x10325476, h4 = 0xc3d2e1f0;
  const w = new Int32Array(80);
  for (let o = 0; o < b.length; o += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getInt32(o + i * 4);
    for (let i = 16; i < 80; i++) w[i] = rotl(w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16], 1);
    let a = h0, bb = h1, c = h2, d = h3, e = h4;
    for (let i = 0; i < 80; i++) {
      const f = i < 20 ? ((bb & c) | (~bb & d)) + 0x5a827999 : i < 40 ? (bb ^ c ^ d) + 0x6ed9eba1 : i < 60 ? ((bb & c) | (bb & d) | (c & d)) + 0x8f1bbcdc : (bb ^ c ^ d) + 0xca62c1d6;
      const t = (rotl(a, 5) + f + e + w[i]) | 0;
      e = d; d = c; c = rotl(bb, 30); bb = a; a = t;
    }
    h0 = (h0 + a) | 0; h1 = (h1 + bb) | 0; h2 = (h2 + c) | 0; h3 = (h3 + d) | 0; h4 = (h4 + e) | 0;
  }
  const out = new Uint8Array(20); const ov = new DataView(out.buffer);
  [h0, h1, h2, h3, h4].forEach((h, i) => ov.setInt32(i * 4, h));
  return out;
}
const K256 = new Int32Array([0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2]);
export function sha256(msg) {
  const { b, dv } = pad64(msg);
  const H = new Int32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const w = new Int32Array(64);
  for (let o = 0; o < b.length; o += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getInt32(o + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) | 0;
    }
    let [a, bb, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i++) {
      const t1 = (h + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + K256[i] + w[i]) | 0;
      const t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & bb) ^ (a & c) ^ (bb & c))) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = bb; bb = a; a = (t1 + t2) | 0;
    }
    H[0] += a; H[1] += bb; H[2] += c; H[3] += d; H[4] += e; H[5] += f; H[6] += g; H[7] += h;
  }
  const out = new Uint8Array(32); const ov = new DataView(out.buffer);
  for (let i = 0; i < 8; i++) ov.setInt32(i * 4, H[i]);
  return out;
}
const concat = (...parts) => { const n = parts.reduce((s, p) => s + p.length, 0); const o = new Uint8Array(n); let k = 0; for (const p of parts) { o.set(p, k); k += p.length; } return o; };
export function hmac(hash, key, msg) {
  if (key.length > 64) key = hash(key);
  const k = new Uint8Array(64); k.set(key);
  const ip = k.map((x) => x ^ 0x36); const op = k.map((x) => x ^ 0x5c);
  return hash(concat(op, hash(concat(ip, msg))));
}
/** PBKDF2 (RFC 2898). hash = sha1 | sha256 */
export function pbkdf2(hash, password, salt, iterations, len) {
  const out = new Uint8Array(len);
  // 내부 · 외부 패드를 한 번만 만들어 반복마다 다시 쓰기 (반복 수가 수만 번이라)
  let key = password; if (key.length > 64) key = hash(key);
  const k = new Uint8Array(64); k.set(key);
  const ip = k.map((x) => x ^ 0x36); const op = k.map((x) => x ^ 0x5c);
  const mac = (m) => hash(concat(op, hash(concat(ip, m))));
  for (let block = 1, pos = 0; pos < len; block++) {
    let u = mac(concat(salt, new Uint8Array([block >>> 24, (block >>> 16) & 255, (block >>> 8) & 255, block & 255])));
    const t = u.slice();
    for (let i = 1; i < iterations; i++) { u = mac(u); for (let j = 0; j < t.length; j++) t[j] ^= u[j]; }
    out.set(t.subarray(0, Math.min(t.length, len - pos)), pos);
    pos += t.length;
  }
  return out;
}

// ───────────── AES ─────────────
const SBOX = new Uint8Array(256); const INV = new Uint8Array(256);
const TE = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)];
const TD = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)];
(() => {
  const x2 = (a) => ((a << 1) ^ (a & 0x80 ? 0x11b : 0)) & 0xff;
  const mul = (a, b) => { let r = 0; while (b) { if (b & 1) r ^= a; a = x2(a); b >>= 1; } return r; };
  // 곱셈 역원으로 S-box 만들기
  const inv = new Uint8Array(256);
  for (let a = 1; a < 256; a++) for (let b = 1; b < 256; b++) if (mul(a, b) === 1) { inv[a] = b; break; }
  for (let a = 0; a < 256; a++) {
    let s = inv[a]; let x = s;
    for (let i = 0; i < 4; i++) { s = ((s << 1) | (s >> 7)) & 0xff; x ^= s; }
    x ^= 0x63; SBOX[a] = x; INV[x] = a;
  }
  for (let a = 0; a < 256; a++) {
    const s = SBOX[a];
    const e = (mul(s, 2) << 24) | (s << 16) | (s << 8) | mul(s, 3);
    const v = INV[a];
    const d = (mul(v, 14) << 24) | (mul(v, 9) << 16) | (mul(v, 13) << 8) | mul(v, 11);
    for (let r = 0; r < 4; r++) { TE[r][a] = ((e >>> (8 * r)) | (e << (32 - 8 * r))) >>> 0; TD[r][a] = ((d >>> (8 * r)) | (d << (32 - 8 * r))) >>> 0; }
  }
})();
/** AES 키 확장 → { enc, dec, rounds } */
export function aesKey(key) {
  const nk = key.length / 4; const rounds = nk + 6;
  const w = new Uint32Array(4 * (rounds + 1));
  for (let i = 0; i < nk; i++) w[i] = ((key[4 * i] << 24) | (key[4 * i + 1] << 16) | (key[4 * i + 2] << 8) | key[4 * i + 3]) >>> 0;
  let rcon = 1;
  for (let i = nk; i < w.length; i++) {
    let t = w[i - 1];
    if (i % nk === 0) { t = ((SBOX[(t >>> 16) & 255] << 24) | (SBOX[(t >>> 8) & 255] << 16) | (SBOX[t & 255] << 8) | SBOX[t >>> 24]) ^ (rcon << 24); rcon = (rcon << 1) ^ (rcon & 0x80 ? 0x11b : 0); }
    else if (nk > 6 && i % nk === 4) t = (SBOX[t >>> 24] << 24) | (SBOX[(t >>> 16) & 255] << 16) | (SBOX[(t >>> 8) & 255] << 8) | SBOX[t & 255];
    w[i] = (w[i - nk] ^ t) >>> 0;
  }
  // 복호 키: 순서를 뒤집고 가운데 라운드에 InvMixColumns
  const d = new Uint32Array(w.length);
  for (let r = 0; r <= rounds; r++) for (let c = 0; c < 4; c++) {
    const k = w[(rounds - r) * 4 + c];
    d[r * 4 + c] = r === 0 || r === rounds ? k : (TD[0][SBOX[k >>> 24]] ^ TD[1][SBOX[(k >>> 16) & 255]] ^ TD[2][SBOX[(k >>> 8) & 255]] ^ TD[3][SBOX[k & 255]]) >>> 0;
  }
  return { enc: w, dec: d, rounds };
}
const put = (o, i, v) => { o[i] = v >>> 24; o[i + 1] = (v >>> 16) & 255; o[i + 2] = (v >>> 8) & 255; o[i + 3] = v & 255; };
function crypt(K, inp, io, out, oo, decrypt) {
  const k = decrypt ? K.dec : K.enc; const T = decrypt ? TD : TE; const S = decrypt ? INV : SBOX;
  let s0 = ((inp[io] << 24) | (inp[io + 1] << 16) | (inp[io + 2] << 8) | inp[io + 3]) ^ k[0];
  let s1 = ((inp[io + 4] << 24) | (inp[io + 5] << 16) | (inp[io + 6] << 8) | inp[io + 7]) ^ k[1];
  let s2 = ((inp[io + 8] << 24) | (inp[io + 9] << 16) | (inp[io + 10] << 8) | inp[io + 11]) ^ k[2];
  let s3 = ((inp[io + 12] << 24) | (inp[io + 13] << 16) | (inp[io + 14] << 8) | inp[io + 15]) ^ k[3];
  const [T0, T1, T2, T3] = T;
  let ki = 4;
  for (let r = 1; r < K.rounds; r++) {
    let t0, t1, t2, t3;
    if (decrypt) {
      t0 = T0[s0 >>> 24] ^ T1[(s3 >>> 16) & 255] ^ T2[(s2 >>> 8) & 255] ^ T3[s1 & 255] ^ k[ki];
      t1 = T0[s1 >>> 24] ^ T1[(s0 >>> 16) & 255] ^ T2[(s3 >>> 8) & 255] ^ T3[s2 & 255] ^ k[ki + 1];
      t2 = T0[s2 >>> 24] ^ T1[(s1 >>> 16) & 255] ^ T2[(s0 >>> 8) & 255] ^ T3[s3 & 255] ^ k[ki + 2];
      t3 = T0[s3 >>> 24] ^ T1[(s2 >>> 16) & 255] ^ T2[(s1 >>> 8) & 255] ^ T3[s0 & 255] ^ k[ki + 3];
    } else {
      t0 = T0[s0 >>> 24] ^ T1[(s1 >>> 16) & 255] ^ T2[(s2 >>> 8) & 255] ^ T3[s3 & 255] ^ k[ki];
      t1 = T0[s1 >>> 24] ^ T1[(s2 >>> 16) & 255] ^ T2[(s3 >>> 8) & 255] ^ T3[s0 & 255] ^ k[ki + 1];
      t2 = T0[s2 >>> 24] ^ T1[(s3 >>> 16) & 255] ^ T2[(s0 >>> 8) & 255] ^ T3[s1 & 255] ^ k[ki + 2];
      t3 = T0[s3 >>> 24] ^ T1[(s0 >>> 16) & 255] ^ T2[(s1 >>> 8) & 255] ^ T3[s2 & 255] ^ k[ki + 3];
    }
    s0 = t0; s1 = t1; s2 = t2; s3 = t3; ki += 4;
  }
  // 마지막 라운드 (MixColumns 없음)
  const a0 = decrypt ? s3 : s1, a1 = decrypt ? s0 : s2, a2 = decrypt ? s1 : s3, a3 = decrypt ? s2 : s0;
  const c0 = decrypt ? s1 : s3, c1 = decrypt ? s2 : s0, c2 = decrypt ? s3 : s1, c3 = decrypt ? s0 : s2;
  put(out, oo, ((S[s0 >>> 24] << 24) | (S[(a0 >>> 16) & 255] << 16) | (S[(s2 >>> 8) & 255] << 8) | S[c0 & 255]) ^ k[ki]);
  put(out, oo + 4, ((S[s1 >>> 24] << 24) | (S[(a1 >>> 16) & 255] << 16) | (S[(s3 >>> 8) & 255] << 8) | S[c1 & 255]) ^ k[ki + 1]);
  put(out, oo + 8, ((S[s2 >>> 24] << 24) | (S[(a2 >>> 16) & 255] << 16) | (S[(s0 >>> 8) & 255] << 8) | S[c2 & 255]) ^ k[ki + 2]);
  put(out, oo + 12, ((S[s3 >>> 24] << 24) | (S[(a3 >>> 16) & 255] << 16) | (S[(s1 >>> 8) & 255] << 8) | S[c3 & 255]) ^ k[ki + 3]);
}
export const aesEncryptBlock = (K, inp, io, out, oo) => crypt(K, inp, io, out, oo, false);
export function aesEcbDecrypt(data, key) {
  const K = aesKey(key); const n = data.length - (data.length % 16); const out = new Uint8Array(n);
  for (let o = 0; o < n; o += 16) crypt(K, data, o, out, o, true);
  return out;
}
export function aesCbcDecrypt(data, key, iv) {
  const K = aesKey(key); const n = data.length - (data.length % 16); const out = new Uint8Array(n);
  for (let o = 0; o < n; o += 16) { crypt(K, data, o, out, o, true); for (let j = 0; j < 16; j++) out[o + j] ^= o ? data[o - 16 + j] : iv[j]; }
  return out;
}
/** AES 비트 단위 CFB(CFB-1) 복호, IV 0 — 레지스터는 스트림 전체에서 이어짐 (HWP 5 암호 문서) */
export function aesCfb1Decrypt(data, key) {
  const K = aesKey(key);
  const reg = new Uint8Array(16); const ks = new Uint8Array(16); const out = new Uint8Array(data.length);
  for (let i = 0; i < data.length; i++) {
    const c = data[i]; let p = 0;
    for (let bit = 7; bit >= 0; bit--) {
      crypt(K, reg, 0, ks, 0, false);
      const cb = (c >> bit) & 1;
      p |= (((ks[0] >> 7) & 1) ^ cb) << bit;
      for (let j = 0; j < 15; j++) reg[j] = ((reg[j] << 1) | (reg[j + 1] >> 7)) & 255;
      reg[15] = ((reg[15] << 1) | cb) & 255;
    }
    out[i] = p;
  }
  return out;
}

// ───────────── DES ─────────────
const PC1 = [57, 49, 41, 33, 25, 17, 9, 1, 58, 50, 42, 34, 26, 18, 10, 2, 59, 51, 43, 35, 27, 19, 11, 3, 60, 52, 44, 36, 63, 55, 47, 39, 31, 23, 15, 7, 62, 54, 46, 38, 30, 22, 14, 6, 61, 53, 45, 37, 29, 21, 13, 5, 28, 20, 12, 4];
const PC2 = [14, 17, 11, 24, 1, 5, 3, 28, 15, 6, 21, 10, 23, 19, 12, 4, 26, 8, 16, 7, 27, 20, 13, 2, 41, 52, 31, 37, 47, 55, 30, 40, 51, 45, 33, 48, 44, 49, 39, 56, 34, 53, 46, 42, 50, 36, 29, 32];
const SHIFTS = [1, 1, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 1];
const IP = [58, 50, 42, 34, 26, 18, 10, 2, 60, 52, 44, 36, 28, 20, 12, 4, 62, 54, 46, 38, 30, 22, 14, 6, 64, 56, 48, 40, 32, 24, 16, 8, 57, 49, 41, 33, 25, 17, 9, 1, 59, 51, 43, 35, 27, 19, 11, 3, 61, 53, 45, 37, 29, 21, 13, 5, 63, 55, 47, 39, 31, 23, 15, 7];
const FP = [40, 8, 48, 16, 56, 24, 64, 32, 39, 7, 47, 15, 55, 23, 63, 31, 38, 6, 46, 14, 54, 22, 62, 30, 37, 5, 45, 13, 53, 21, 61, 29, 36, 4, 44, 12, 52, 20, 60, 28, 35, 3, 43, 11, 51, 19, 59, 27, 34, 2, 42, 10, 50, 18, 58, 26, 33, 1, 41, 9, 49, 17, 57, 25];
const E = [32, 1, 2, 3, 4, 5, 4, 5, 6, 7, 8, 9, 8, 9, 10, 11, 12, 13, 12, 13, 14, 15, 16, 17, 16, 17, 18, 19, 20, 21, 20, 21, 22, 23, 24, 25, 24, 25, 26, 27, 28, 29, 28, 29, 30, 31, 32, 1];
const P = [16, 7, 20, 21, 29, 12, 28, 17, 1, 15, 23, 26, 5, 18, 31, 10, 2, 8, 24, 14, 32, 27, 3, 9, 19, 13, 30, 6, 22, 11, 4, 25];
const SB = [
  [14, 4, 13, 1, 2, 15, 11, 8, 3, 10, 6, 12, 5, 9, 0, 7, 0, 15, 7, 4, 14, 2, 13, 1, 10, 6, 12, 11, 9, 5, 3, 8, 4, 1, 14, 8, 13, 6, 2, 11, 15, 12, 9, 7, 3, 10, 5, 0, 15, 12, 8, 2, 4, 9, 1, 7, 5, 11, 3, 14, 10, 0, 6, 13],
  [15, 1, 8, 14, 6, 11, 3, 4, 9, 7, 2, 13, 12, 0, 5, 10, 3, 13, 4, 7, 15, 2, 8, 14, 12, 0, 1, 10, 6, 9, 11, 5, 0, 14, 7, 11, 10, 4, 13, 1, 5, 8, 12, 6, 9, 3, 2, 15, 13, 8, 10, 1, 3, 15, 4, 2, 11, 6, 7, 12, 0, 5, 14, 9],
  [10, 0, 9, 14, 6, 3, 15, 5, 1, 13, 12, 7, 11, 4, 2, 8, 13, 7, 0, 9, 3, 4, 6, 10, 2, 8, 5, 14, 12, 11, 15, 1, 13, 6, 4, 9, 8, 15, 3, 0, 11, 1, 2, 12, 5, 10, 14, 7, 1, 10, 13, 0, 6, 9, 8, 7, 4, 15, 14, 3, 11, 5, 2, 12],
  [7, 13, 14, 3, 0, 6, 9, 10, 1, 2, 8, 5, 11, 12, 4, 15, 13, 8, 11, 5, 6, 15, 0, 3, 4, 7, 2, 12, 1, 10, 14, 9, 10, 6, 9, 0, 12, 11, 7, 13, 15, 1, 3, 14, 5, 2, 8, 4, 3, 15, 0, 6, 10, 1, 13, 8, 9, 4, 5, 11, 12, 7, 2, 14],
  [2, 12, 4, 1, 7, 10, 11, 6, 8, 5, 3, 15, 13, 0, 14, 9, 14, 11, 2, 12, 4, 7, 13, 1, 5, 0, 15, 10, 3, 9, 8, 6, 4, 2, 1, 11, 10, 13, 7, 8, 15, 9, 12, 5, 6, 3, 0, 14, 11, 8, 12, 7, 1, 14, 2, 13, 6, 15, 0, 9, 10, 4, 5, 3],
  [12, 1, 10, 15, 9, 2, 6, 8, 0, 13, 3, 4, 14, 7, 5, 11, 10, 15, 4, 2, 7, 12, 9, 5, 6, 1, 13, 14, 0, 11, 3, 8, 9, 14, 15, 5, 2, 8, 12, 3, 7, 0, 4, 10, 1, 13, 11, 6, 4, 3, 2, 12, 9, 5, 15, 10, 11, 14, 1, 7, 6, 0, 8, 13],
  [4, 11, 2, 14, 15, 0, 8, 13, 3, 12, 9, 7, 5, 10, 6, 1, 13, 0, 11, 7, 4, 9, 1, 10, 14, 3, 5, 12, 2, 15, 8, 6, 1, 4, 11, 13, 12, 3, 7, 14, 10, 15, 6, 8, 0, 5, 9, 2, 6, 11, 13, 8, 1, 4, 10, 7, 9, 5, 0, 15, 14, 2, 3, 12],
  [13, 2, 8, 4, 6, 15, 11, 1, 10, 9, 3, 14, 5, 0, 12, 7, 1, 15, 13, 8, 10, 3, 7, 4, 12, 5, 6, 11, 0, 14, 9, 2, 7, 11, 4, 1, 9, 12, 14, 2, 0, 6, 10, 13, 15, 3, 5, 8, 2, 1, 14, 7, 4, 10, 8, 13, 15, 12, 9, 0, 3, 5, 6, 11],
];
const bitsOf = (bytes) => { const b = []; for (const x of bytes) for (let i = 7; i >= 0; i--) b.push((x >> i) & 1); return b; };
const perm = (bits, table) => table.map((i) => bits[i - 1]);
function desSubkeys(key) {
  const k = perm(bitsOf(key), PC1);
  let c = k.slice(0, 28); let d = k.slice(28);
  return SHIFTS.map((s) => { c = c.slice(s).concat(c.slice(0, s)); d = d.slice(s).concat(d.slice(0, s)); return perm(c.concat(d), PC2); });
}
function desBlock(sub, block, decrypt) {
  const b = perm(bitsOf(block), IP);
  let l = b.slice(0, 32); let r = b.slice(32);
  for (let i = 0; i < 16; i++) {
    const k = sub[decrypt ? 15 - i : i];
    const x = perm(r, E).map((v, j) => v ^ k[j]);
    const f = [];
    for (let s = 0; s < 8; s++) {
      const o = s * 6; const row = (x[o] << 1) | x[o + 5]; const col = (x[o + 1] << 3) | (x[o + 2] << 2) | (x[o + 3] << 1) | x[o + 4];
      const v = SB[s][row * 16 + col]; f.push((v >> 3) & 1, (v >> 2) & 1, (v >> 1) & 1, v & 1);
    }
    const pr = perm(f, P);
    const nr = l.map((v, j) => v ^ pr[j]); l = r; r = nr;
  }
  const out = perm(r.concat(l), FP);
  const o = new Uint8Array(8);
  for (let i = 0; i < 64; i++) o[i >> 3] |= out[i] << (7 - (i & 7));
  return o;
}
export function desEcbDecrypt(data, key) {
  const sub = desSubkeys(key); const n = data.length - (data.length % 8); const out = new Uint8Array(n);
  for (let o = 0; o < n; o += 8) out.set(desBlock(sub, data.subarray(o, o + 8), true), o);
  return out;
}
export function desEcbEncrypt(data, key) {
  const sub = desSubkeys(key); const n = data.length - (data.length % 8); const out = new Uint8Array(n);
  for (let o = 0; o < n; o += 8) out.set(desBlock(sub, data.subarray(o, o + 8), false), o);
  return out;
}

// ───────────── MD5 · RC4 (Word 97-2003 암호 문서) ─────────────
const MD5_S = [7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21];
const MD5_K = Array.from({ length: 64 }, (_, i) => Math.floor(Math.abs(Math.sin(i + 1)) * 2 ** 32) | 0);
export function md5(msg) {
  const len = msg.length; const total = Math.ceil((len + 9) / 64) * 64;
  const b = new Uint8Array(total); b.set(msg); b[len] = 0x80;
  const dv = new DataView(b.buffer);
  dv.setUint32(total - 8, (len * 8) >>> 0, true); dv.setUint32(total - 4, Math.floor(len / 2 ** 29), true);
  let a0 = 0x67452301, b0 = 0xefcdab89 | 0, c0 = 0x98badcfe | 0, d0 = 0x10325476;
  const m = new Int32Array(16);
  for (let o = 0; o < total; o += 64) {
    for (let i = 0; i < 16; i++) m[i] = dv.getInt32(o + i * 4, true);
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F, g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; } else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) % 16; } else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) % 16; } else { F = C ^ (B | ~D); g = (7 * i) % 16; }
      F = (F + A + MD5_K[i] + m[g]) | 0;
      A = D; D = C; C = B; B = (B + rotl(F, MD5_S[i])) | 0;
    }
    a0 = (a0 + A) | 0; b0 = (b0 + B) | 0; c0 = (c0 + C) | 0; d0 = (d0 + D) | 0;
  }
  const out = new Uint8Array(16); const ov = new DataView(out.buffer);
  [a0, b0, c0, d0].forEach((h, i) => ov.setInt32(i * 4, h, true));
  return out;
}
/** RC4 상태 → 바이트를 차례로 XOR 하는 함수 */
export function rc4(key) {
  const S = new Uint8Array(256); for (let i = 0; i < 256; i++) S[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + S[i] + key[i % key.length]) & 255; const t = S[i]; S[i] = S[j]; S[j] = t; }
  let i = 0; let j = 0;
  return (data) => {
    const out = new Uint8Array(data.length);
    for (let k = 0; k < data.length; k++) { i = (i + 1) & 255; j = (j + S[i]) & 255; const t = S[i]; S[i] = S[j]; S[j] = t; out[k] = data[k] ^ S[(S[i] + S[j]) & 255]; }
    return out;
  };
}

// ───────────── SHA-512 (Office 2010+ 암호 문서) — 32비트 두 쪽으로 계산 ─────────────
const K512 = [
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc, 0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118,
  0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2, 0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694,
  0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65, 0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4, 0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725, 0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70,
  0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df, 0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30, 0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8, 0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3,
  0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec, 0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b,
  0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178, 0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c, 0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
];
export function sha512(msg) {
  const len = msg.length; const total = Math.ceil((len + 17) / 128) * 128;
  const b = new Uint8Array(total); b.set(msg); b[len] = 0x80;
  const dv = new DataView(b.buffer);
  dv.setUint32(total - 4, (len * 8) >>> 0); dv.setUint32(total - 8, Math.floor(len / 2 ** 29));
  const H = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];
  const W = new Int32Array(160);
  // 64비트 더하기: (ah, al) + (bh, bl)
  for (let o = 0; o < total; o += 128) {
    for (let i = 0; i < 32; i++) W[i] = dv.getInt32(o + i * 4);
    for (let i = 16; i < 80; i++) {
      const xh = W[(i - 15) * 2], xl = W[(i - 15) * 2 + 1];
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      const yh = W[(i - 2) * 2], yl = W[(i - 2) * 2 + 1];
      const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
      const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
      let l = (W[(i - 16) * 2 + 1] >>> 0) + (s0l >>> 0) + (W[(i - 7) * 2 + 1] >>> 0) + (s1l >>> 0);
      let h = W[(i - 16) * 2] + s0h + W[(i - 7) * 2] + s1h + Math.floor(l / 2 ** 32);
      W[i * 2] = h | 0; W[i * 2 + 1] = l | 0;
    }
    let [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl] = H;
    for (let i = 0; i < 80; i++) {
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
      let t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i * 2 + 1] >>> 0) + (W[i * 2 + 1] >>> 0);
      let t1h = hh + S1h + chh + K512[i * 2] + W[i * 2] + Math.floor(t1l / 2 ** 32);
      t1l |= 0; t1h |= 0;
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      const t2l = (S0l >>> 0) + (mjl >>> 0); const t2h = S0h + mjh + Math.floor(t2l / 2 ** 32);
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      let nl = (dl >>> 0) + (t1l >>> 0); eh = (dh + t1h + Math.floor(nl / 2 ** 32)) | 0; el = nl | 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      nl = (t1l >>> 0) + (t2l >>> 0); ah = (t1h + t2h + Math.floor(nl / 2 ** 32)) | 0; al = nl | 0;
    }
    const add = (k, h, l) => { const s = (H[k + 1] >>> 0) + (l >>> 0); H[k] = (H[k] + h + Math.floor(s / 2 ** 32)) | 0; H[k + 1] = s | 0; };
    add(0, ah, al); add(2, bh, bl); add(4, ch, cl); add(6, dh, dl); add(8, eh, el); add(10, fh, fl); add(12, gh, gl); add(14, hh, hl);
  }
  const out = new Uint8Array(64); const ov = new DataView(out.buffer);
  for (let i = 0; i < 16; i++) ov.setInt32(i * 4, H[i]);
  return out;
}
