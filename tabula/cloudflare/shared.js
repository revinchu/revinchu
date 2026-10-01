export const LIMITS = Object.freeze({ maxDocumentBytes: 20 * 1024 * 1024, maxVaultBytes: 100 * 1024 * 1024, maxDocuments: 50, maxPublications: 20, maxProxyBytes: 5 * 1024 * 1024, maxVersions: 20, maxHistoryBytes: 100 * 1024 * 1024, importTimeoutMs: 30 * 60 * 1000 });
export class ApiError extends Error {
  constructor(status, message, code = 'REQUEST_FAILED', extra = {}) { super(message); this.status = status; this.code = code; this.extra = extra; }
}
export function errorResponse(error) {
  if (!(error instanceof ApiError)) return json({ error: '서버 처리 중 오류가 발생했습니다. 잠시 후 다시 시도하세요.', code: 'INTERNAL_ERROR' }, 500);
  return json({ error: error.message, code: error.code, ...error.extra }, error.status, error.status === 429 ? { 'Retry-After': '60' } : {});
}
export function json(value, status = 200, headers = {}) {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ...headers } });
}
export function securityHeaders(response) {
  const out = new Response(response.body, response);
  out.headers.set('X-Content-Type-Options', 'nosniff');
  out.headers.set('Referrer-Policy', 'same-origin');
  out.headers.set('X-Frame-Options', 'SAMEORIGIN');
  out.headers.set('Content-Security-Policy', "script-src 'self'; script-src-attr 'none'; object-src 'none'; base-uri 'self'; frame-ancestors 'self'; form-action 'self'");
  return out;
}
export function validateName(name) {
  if (typeof name !== 'string' || !name.trim() || name.length > 120 || /[\u0000-\u001f\u007f]/.test(name)) throw new ApiError(400, '문서 이름은 제어 문자 없이 1~120자로 입력하세요.', 'INVALID_NAME');
  return name.trim();
}
export function parseRevision(request) {
  const value = request.headers.get('If-Match');
  if (value === null) throw new ApiError(428, '현재 문서 버전을 먼저 확인하세요.', 'REVISION_REQUIRED');
  const match = /^(?:"(\d+)"|(\d+))$/.exec(value);
  const revision = match ? Number(match[1] ?? match[2]) : NaN;
  if (!Number.isSafeInteger(revision) || revision < 0) throw new ApiError(400, '문서 버전 값이 올바르지 않습니다.', 'INVALID_REVISION');
  return revision;
}
export function documentHeaders(meta) {
  return { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', ETag: '"' + meta.revision + '"', 'X-Wixel-Revision': String(meta.revision), 'X-Modified': String(meta.modified) };
}
export function validateOrigin(request) {
  const origin = request.headers.get('Origin'), site = request.headers.get('Sec-Fetch-Site');
  if (site === 'cross-site' || (origin && origin !== new URL(request.url).origin)) throw new ApiError(403, '다른 사이트에서 보낸 요청은 허용하지 않습니다.', 'CROSS_ORIGIN');
}
export function vaultKey(request) {
  const key = request.headers.get('X-Wixel-Vault') ?? '';
  if (!/^[A-Za-z0-9_-]{43}$/.test(key) || !/[AEIMQUYcgkosw048]$/.test(key)) throw new ApiError(401, '개인 보관함을 만들거나 복구키로 연결하세요.', 'VAULT_KEY_REQUIRED');
  return key;
}
export async function sha256(value) {
  const bytes = new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value)));
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}
export function randomId() {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  let binary = ''; for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
export function validPublicId(id) { return /^[A-Za-z0-9_-]{43}$/.test(id); }
export function sameHash(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== 64 || b.length !== 64) return false;
  const x = new TextEncoder().encode(a), y = new TextEncoder().encode(b);
  if (typeof crypto.subtle.timingSafeEqual === 'function') return crypto.subtle.timingSafeEqual(x, y);
  let different = 0; for (let i = 0; i < x.length; i++) different |= x[i] ^ y[i]; return different === 0;
}
