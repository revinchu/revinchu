// 암호로 보호된 Office 2007+ 문서 (.docx 등): OLE 복합 문서 안의 EncryptionInfo + EncryptedPackage 풀기 — DOM 없음
// MS-OFFCRYPTO 의 Agile 암호화(Office 2010+, SHA-512 + AES-CBC, 4096바이트 조각) 와 Standard 암호화(Office 2007, SHA-1 + AES-ECB)
import { openCfb } from './cfb.js';
import { sha1, sha512, sha256, aesCbcDecrypt, aesEcbDecrypt } from './crypto.js';

const HASH = { SHA512: sha512, SHA1: sha1, 'SHA-1': sha1, SHA256: sha256 };
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
const le32 = (n) => new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
const cat = (...a) => { const n = a.reduce((x, y) => x + y.length, 0); const o = new Uint8Array(n); let k = 0; for (const x of a) { o.set(x, k); k += x.length; } return o; };
const utf16 = (s) => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { b[2 * i] = s.charCodeAt(i) & 255; b[2 * i + 1] = s.charCodeAt(i) >> 8; } return b; };
const b64 = (s) => Uint8Array.from(atob(s ?? ''), (c) => c.charCodeAt(0));
/** 길이를 맞춤: 길면 자르고, 짧으면 0x36 으로 채움 */
const fit = (b, n) => { if (b.length >= n) return b.slice(0, n); const o = new Uint8Array(n).fill(0x36); o.set(b); return o; };
const wrong = () => { const e = new Error('암호가 맞지 않습니다.'); e.code = 'PASSWORD'; e.wrong = true; return e; };

export const isEncryptedOoxml = (cfb) => cfb.names.includes('EncryptionInfo') && cfb.names.includes('EncryptedPackage');

/** → 풀린 ZIP 바이트 */
export function decryptOoxml(bytes, password) {
  const cfb = openCfb(bytes);
  const info = cfb.stream('EncryptionInfo'); const pkg = cfb.stream('EncryptedPackage');
  if (!info || !pkg) throw new Error('암호 문서의 정보가 없습니다.');
  if (password == null || password === '') { const e = new Error('암호로 보호된 문서입니다. 열기 암호를 입력하세요.'); e.code = 'PASSWORD'; throw e; }
  const major = info[0] | (info[1] << 8); const minor = info[2] | (info[3] << 8);
  const size = u32(pkg, 0) + u32(pkg, 4) * 2 ** 32;
  const data = pkg.subarray(8);
  if (major === 4 && minor === 4) {
    // Agile: 나머지는 XML
    const xml = new TextDecoder().decode(info.subarray(8));
    const tag = (name) => new RegExp(`<(?:\\w+:)?${name}\\b([^>]*)>`).exec(xml)?.[1] ?? '';
    const at = (attrs, n) => new RegExp(`\\b${n}="([^"]*)"`).exec(attrs)?.[1];
    const kd = tag('keyData'); const ek = tag('encryptedKey');
    const kdHash = HASH[at(kd, 'hashAlgorithm')] ?? sha1; const ekHash = HASH[at(ek, 'hashAlgorithm')] ?? sha1;
    const spin = Number(at(ek, 'spinCount') ?? 100000);
    if (!(spin >= 0 && spin <= 10_000_000)) throw new Error('암호 정보가 이상합니다.');
    const ekSalt = b64(at(ek, 'saltValue')); const ekBits = Number(at(ek, 'keyBits') ?? 256);
    let h = ekHash(cat(ekSalt, utf16(password)));
    for (let i = 0; i < spin; i++) h = ekHash(cat(le32(i), h));
    const derive = (block) => fit(ekHash(cat(h, block)), ekBits / 8);
    const BLK = { vin: [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79], vhash: [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e], key: [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6] };
    const dec = (blk, v) => aesCbcDecrypt(b64(v), derive(new Uint8Array(blk)), ekSalt);
    const vin = dec(BLK.vin, at(ek, 'encryptedVerifierHashInput'));
    const vhash = dec(BLK.vhash, at(ek, 'encryptedVerifierHashValue'));
    const calc = ekHash(vin.subarray(0, ekSalt.length));
    if (!calc.every((x, i) => x === vhash[i])) throw wrong();
    const secret = dec(BLK.key, at(ek, 'encryptedKeyValue')).subarray(0, ekBits / 8);
    const kdSalt = b64(at(kd, 'saltValue')); const bs = Number(at(kd, 'blockSize') ?? 16);
    const out = new Uint8Array(Math.ceil(data.length / 16) * 16);
    for (let o = 0, seg = 0; o < data.length; o += 4096, seg++) {
      const iv = fit(kdHash(cat(kdSalt, le32(seg))), bs);
      out.set(aesCbcDecrypt(data.subarray(o, Math.min(data.length, o + 4096)), secret, iv), o);
    }
    return out.subarray(0, size);
  }
  if ((major === 3 || major === 4) && minor === 2) {
    // Standard: EncryptionHeader + EncryptionVerifier
    const hsz = u32(info, 8); const h0 = 12; const keyBits = u32(info, h0 + 16) || 128;
    const vo = h0 + hsz; const saltSize = u32(info, vo); const salt = info.subarray(vo + 4, vo + 4 + saltSize);
    const ev = info.subarray(vo + 4 + saltSize, vo + 20 + saltSize); const evh = info.subarray(vo + 24 + saltSize, vo + 24 + saltSize + 32);
    let h = sha1(cat(salt, utf16(password)));
    for (let i = 0; i < 50000; i++) h = sha1(cat(le32(i), h));
    const hf = sha1(cat(h, le32(0)));
    const x1 = sha1(new Uint8Array(64).fill(0x36).map((v, i) => v ^ (hf[i] ?? 0)));
    const x2 = sha1(new Uint8Array(64).fill(0x5c).map((v, i) => v ^ (hf[i] ?? 0)));
    const key = cat(x1, x2).subarray(0, keyBits / 8);
    const v = aesEcbDecrypt(ev, key); const vh = aesEcbDecrypt(evh, key);
    if (!sha1(v).every((x, i) => x === vh[i])) throw wrong();
    return aesEcbDecrypt(data.subarray(0, data.length - (data.length % 16)), key).subarray(0, size);
  }
  throw new Error(`지원하지 않는 Office 암호 방식입니다 (${major}.${minor}).`);
}
