// Node 서버 전용: 목적지를 검사한 뒤 같은 IP로 연결하여 DNS 재조회 우회를 막습니다.
import { lookup as dnsLookup } from 'node:dns/promises';
import { isIP } from 'node:net';
import { request as httpRequest } from 'node:http';
import { request as httpsRequest } from 'node:https';
import { Transform, Writable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { createGunzip, createInflate, createBrotliDecompress } from 'node:zlib';

export class ProxyError extends Error {
  constructor(message, status = 400, options) { super(message, options); this.status = status; }
}

const MAX_BYTES = 5 * 1024 * 1024;
const REDIRECTS = new Set([301, 302, 303, 307, 308]);

function bareHost(host) { return host.replace(/^\[|\]$/g, '').toLowerCase().replace(/\.$/, ''); }

function ipv6Number(address) {
  let text = address;
  if (text.includes('.')) {
    const at = text.lastIndexOf(':');
    const bytes = text.slice(at + 1).split('.').map(Number);
    text = text.slice(0, at + 1) + ((bytes[0] << 8) | bytes[1]).toString(16) + ':' + ((bytes[2] << 8) | bytes[3]).toString(16);
  }
  const [left, right] = text.split('::');
  const a = left ? left.split(':') : [], b = right ? right.split(':') : [];
  const words = right === undefined ? a : [...a, ...Array(8 - a.length - b.length).fill('0'), ...b];
  return words.reduce((n, word) => (n << 16n) | BigInt('0x' + word), 0n);
}

function inPrefix(value, base, bits) { const shift = BigInt(128 - bits); return (value >> shift) === (base >> shift); }

/** 외부 HTTP 중계에 허용되는 전역 유니캐스트 주소만 true. 특수·전환용 대역은 보수적으로 제외. */
export function isPublicAddress(address) {
  if (typeof address !== 'string' || address.includes('%')) return false;
  const family = isIP(address);
  if (family === 4) {
    // Azure 플랫폼 가상 IP는 공인 주소 형태지만 VM 내부 제어 채널이므로 중계하지 않습니다.
    if (address === '168.63.129.16') return false;
    const [a, b, c] = address.split('.').map(Number);
    return !(a === 0 || a === 10 || a === 127 || a >= 224 ||
      (a === 100 && b >= 64 && b <= 127) || (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && (b === 168 || (b === 0 && (c === 0 || c === 2)) || (b === 88 && c === 99))) ||
      (a === 198 && (b === 18 || b === 19 || (b === 51 && c === 100))) ||
      (a === 203 && b === 0 && c === 113));
  }
  if (family !== 6) return false;
  const n = ipv6Number(address);
  // 2000::/3 밖에는 루프백, mapped IPv4, NAT64, 링크로컬, ULA, 멀티캐스트 등이 포함됩니다.
  return inPrefix(n, 0x20000000000000000000000000000000n, 3) &&
    !inPrefix(n, 0x20010000000000000000000000000000n, 23) &&
    !inPrefix(n, 0x20010db8000000000000000000000000n, 32) &&
    !inPrefix(n, 0x20020000000000000000000000000000n, 16) &&
    !inPrefix(n, 0x3fff0000000000000000000000000000n, 20);
}

export function isLoopbackAddress(address) {
  const host = bareHost(String(address));
  if (isIP(host) === 4) return host.split('.')[0] === '127';
  if (isIP(host) !== 6 || host.includes('%')) return false;
  const n = ipv6Number(host);
  return n === 1n || ((n >> 32n) === 0xffffn && Number((n >> 24n) & 255n) === 127);
}

export function parsePublicUrl(input) {
  let url;
  try { url = new URL(input); } catch { throw new ProxyError('잘못된 주소입니다'); }
  if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password ||
      (url.port && url.port !== '80' && url.port !== '443')) {
    throw new ProxyError('사용자 정보가 없는 공개 HTTP/HTTPS 주소와 포트 80·443만 가져올 수 있습니다');
  }
  const host = bareHost(url.hostname);
  if (host === 'localhost' || /\.(localhost|local|internal|home\.arpa)$/.test(host) ||
      (isIP(host) && !isPublicAddress(host))) {
    throw new ProxyError('내부망·특수 목적 주소는 가져올 수 없습니다');
  }
  return url;
}

export async function resolvePublicTarget(input, lookup = dnsLookup) {
  const url = parsePublicUrl(input), hostname = bareHost(url.hostname), family = isIP(hostname);
  const addresses = family ? [{ address: hostname, family }] : await lookup(hostname, { all: true, verbatim: true });
  if (!Array.isArray(addresses) || !addresses.length || addresses.some((entry) =>
    !entry || isIP(entry.address) !== entry.family || !isPublicAddress(entry.address))) {
    throw new ProxyError('주소가 내부망 또는 허용되지 않는 IP로 연결됩니다');
  }
  // IPv4 경로를 우선하며, 새 연결의 lookup은 이 이미 검사된 주소만 돌려줍니다.
  const selected = addresses.find((entry) => entry.family === 4) ?? addresses[0];
  return { url, hostname, address: selected.address, family: selected.family };
}

/** 검증된 목적지 전용 전송. 테스트는 request 함수만 주입하며 운영 HTTP 입력으로 정책을 바꿀 수 없습니다. */
export function requestPinned(target, { signal, headers = {}, requestHttp = httpRequest, requestHttps = httpsRequest } = {}) {
  return new Promise((resolve, reject) => {
    const { url, hostname, address, family } = target;
    const request = url.protocol === 'https:' ? requestHttps : requestHttp;
    const options = {
      protocol: url.protocol, hostname, port: url.port || undefined,
      path: url.pathname + url.search, method: 'GET', signal,
      agent: false, family, autoSelectFamily: false,
      // IP를 hostname으로 바꾸지 않아 원래 호스트에 대한 SNI/인증서 검증을 유지합니다.
      rejectUnauthorized: true,
      ...(url.protocol === 'https:' && !isIP(hostname) ? { servername: hostname } : {}),
      lookup(_name, opts, callback) {
        if (typeof opts === 'function') { callback = opts; opts = {}; }
        if (opts?.all) callback(null, [{ address, family }]);
        else callback(null, address, family);
      },
      headers: {
        'User-Agent': 'Mozilla/5.0 WIXEL', 'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8',
        ...headers, Host: url.host, 'Accept-Encoding': 'identity',
      },
    };
    const req = request(options, resolve);
    req.once('error', reject);
    req.end();
  });
}

function abortable(promise, signal) {
  if (signal.aborted) return Promise.reject(signal.reason);
  return new Promise((resolve, reject) => {
    const abort = () => reject(signal.reason);
    signal.addEventListener('abort', abort, { once: true });
    promise.then(resolve, reject).finally(() => signal.removeEventListener('abort', abort));
  });
}

async function readLimitedBody(response, maxBytes, signal) {
  const tooLarge = () => new ProxyError('허용된 응답 크기(최대 5MB)를 초과했습니다', 413);
  const length = Number(response.headers['content-length']);
  if (Number.isFinite(length) && length > maxBytes) { response.destroy(); throw tooLarge(); }
  const encoding = String(response.headers['content-encoding'] ?? 'identity').trim().toLowerCase();
  const decode = encoding === 'gzip' ? createGunzip() : encoding === 'deflate' ? createInflate() : encoding === 'br' ? createBrotliDecompress() : null;
  if (!decode && encoding !== 'identity' && encoding !== '') {
    response.destroy(); throw new ProxyError('지원하지 않는 응답 압축 형식입니다', 502);
  }
  let wireSize = 0, size = 0;
  const chunks = [];
  const wireLimit = new Transform({ transform(chunk, _encoding, callback) {
    wireSize += chunk.length;
    callback(wireSize > maxBytes ? tooLarge() : null, chunk);
  } });
  const output = new Writable({ write(chunk, _encoding, callback) {
    size += chunk.length;
    if (size > maxBytes) return callback(tooLarge());
    chunks.push(chunk); callback();
  } });
  const streams = decode ? [response, wireLimit, decode, output] : [response, wireLimit, output];
  await pipeline(...streams, { signal });
  return Buffer.concat(chunks, size);
}

/** DNS·리디렉션·본문 전체에 같은 시간 제한을 적용합니다. 네이버 서명 헤더 요청은 maxRedirects: 0. */
export async function fetchPublicResource(input, {
  lookup = dnsLookup, request = requestPinned, headers = {}, maxRedirects = 5,
  maxBytes = MAX_BYTES, timeoutMs = 15000,
} = {}) {
  const controller = new AbortController(), { signal } = controller;
  const timer = setTimeout(() => controller.abort(new ProxyError('외부 서버 응답 시간(15초)을 초과했습니다', 504)), timeoutMs);
  let current = input;
  try {
    for (let redirects = 0; ; redirects++) {
      const target = await abortable(resolvePublicTarget(current, lookup), signal);
      if (signal.aborted) throw signal.reason;
      const response = await abortable(request(target, { signal, headers }), signal);
      const status = response.statusCode;
      if (REDIRECTS.has(status) && response.headers.location) {
        const location = response.headers.location;
        response.destroy();
        if (redirects >= maxRedirects) throw new ProxyError('허용된 리디렉션 횟수를 초과했습니다', 502);
        try { current = new URL(location, target.url); } catch { throw new ProxyError('이동할 주소가 올바르지 않습니다', 502); }
        continue;
      }
      const body = await readLimitedBody(response, maxBytes, signal);
      return { status, headers: response.headers, body, url: target.url.href };
    }
  } catch (error) {
    if (signal.aborted) throw signal.reason;
    if (error instanceof ProxyError) throw error;
    throw new ProxyError('외부 서버에 연결하지 못했습니다', 502, { cause: error });
  } finally { clearTimeout(timer); }
}
