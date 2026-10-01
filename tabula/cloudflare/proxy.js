import { ApiError, LIMITS } from './shared.js';

export function publicUrl(value, selfHost) {
  let url; try { url = new URL(value); } catch { throw new ApiError(400, '가져올 HTTPS 주소가 올바르지 않습니다.', 'INVALID_URL'); }
  const host = url.hostname.toLowerCase().replace(/\.$/, '');
  if (url.protocol !== 'https:' || (url.port && url.port !== '443') || url.username || url.password ||
      host === selfHost?.toLowerCase().replace(/\.$/, '') || !host.includes('.') ||
      host.includes(':') || /^\d+(?:\.\d+){3}$/.test(host) || /(?:^|\.)(?:localhost|local|internal|invalid|test|example|onion)$/.test(host)) {
    throw new ApiError(403, '공개 HTTPS 주소의 기본 포트만 가져올 수 있습니다.', 'UNSAFE_URL');
  }
  url.hash = ''; return url;
}
export async function boundedBytes(response, maxBytes, signal) {
  if (Number(response.headers.get('Content-Length')) > maxBytes) {
    await response.body?.cancel(); throw new ApiError(413, '가져올 응답의 크기가 너무 큽니다.', 'RESPONSE_TOO_LARGE');
  }
  if (!response.body) return new Uint8Array();
  const reader = response.body.getReader(), chunks = []; let length = 0, complete = false;
  const aborted = () => reader.cancel().catch(() => {});
  signal?.addEventListener('abort', aborted, { once: true });
  try {
    for (;;) {
      if (signal?.aborted) throw new ApiError(504, '가져오기 제한 시간을 초과했습니다.', 'UPSTREAM_TIMEOUT');
      const item = await reader.read(); if (item.done) break;
      length += item.value.byteLength;
      if (length > maxBytes) throw new ApiError(413, '가져올 응답의 크기가 너무 큽니다.', 'RESPONSE_TOO_LARGE');
      chunks.push(item.value);
    }
    if (signal?.aborted) throw new ApiError(504, '가져오기 제한 시간을 초과했습니다.', 'UPSTREAM_TIMEOUT');
    complete = true; const bytes = new Uint8Array(length); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; } return bytes;
  } finally {
    signal?.removeEventListener('abort', aborted);
    if (!complete) await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
}
export async function fetchPublicText(value, selfHost, fetcher = fetch) {
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 15000);
  let url = publicUrl(value, selfHost);
  try {
    for (let redirects = 0;; redirects++) {
      const response = await fetcher(url.href, {
        method: 'GET', redirect: 'manual', signal: controller.signal,
        headers: { Accept: 'text/plain,text/csv,application/json,text/html,application/xml;q=0.9,*/*;q=0.1', 'Accept-Encoding': 'identity' },
      });
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        await response.body?.cancel();
        if (redirects >= 5) throw new ApiError(502, '리디렉션 횟수가 5회를 초과했습니다.', 'TOO_MANY_REDIRECTS');
        const location = response.headers.get('Location');
        if (!location) throw new ApiError(502, '리디렉션 주소가 없습니다.', 'INVALID_REDIRECT');
        url = publicUrl(new URL(location, url).href, selfHost); continue;
      }
      const bytes = await boundedBytes(response, LIMITS.maxProxyBytes, controller.signal);
      // 원격 Content-Type/Set-Cookie/인증 헤더를 전달하지 않아 실행 가능한 HTML 중계를 막습니다.
      return new Response(bytes, { status: response.ok ? 200 : (response.status >= 400 ? response.status : 502),
        headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' } });
    }
  } catch (error) {
    if (error instanceof ApiError) throw error;
    if (controller.signal.aborted) throw new ApiError(504, '가져오기 제한 시간을 초과했습니다.', 'UPSTREAM_TIMEOUT');
    throw new ApiError(502, '공개 주소에서 데이터를 가져오지 못했습니다.', 'UPSTREAM_FAILED');
  } finally { clearTimeout(timer); }
}
export async function naverKeywords(request, fetcher = fetch) {
  const customer = request.headers.get('X-Customer'), key = request.headers.get('X-API-Key'), secret = request.headers.get('X-Secret');
  const keywords = new URL(request.url).searchParams.get('hintKeywords');
  if (!customer || !key || !secret || !keywords || [customer, key, secret].some(x => x.length > 1024) || keywords.length > 500)
    throw new ApiError(400, '계정 ID · 액세스 라이선스 · 비밀 키 · 키워드를 확인하세요.', 'NAVER_INPUT');
  const timestamp = String(Date.now()), encoder = new TextEncoder();
  const cryptoKey = await crypto.subtle.importKey('raw', encoder.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const signed = new Uint8Array(await crypto.subtle.sign('HMAC', cryptoKey, encoder.encode(timestamp + '.GET./keywordstool')));
  let binary = ''; for (const b of signed) binary += String.fromCharCode(b);
  const url = new URL('https://api.searchad.naver.com/keywordstool');
  url.searchParams.set('hintKeywords', keywords); url.searchParams.set('showDetail', '1');
  const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetcher(url.href, { redirect: 'manual', signal: controller.signal,
      headers: { 'X-Timestamp': timestamp, 'X-API-KEY': key, 'X-Customer': customer, 'X-Signature': btoa(binary), 'Accept-Encoding': 'identity' } });
    if (response.status >= 300 && response.status < 400) { await response.body?.cancel(); throw new ApiError(502, '네이버 API의 리디렉션은 허용하지 않습니다.', 'NAVER_REDIRECT'); }
    const bytes = await boundedBytes(response, LIMITS.maxProxyBytes, controller.signal);
    return new Response(bytes, { status: response.status, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
  } catch (error) {
    if (error instanceof ApiError) throw error;
    throw new ApiError(controller.signal.aborted ? 504 : 502, '네이버 API 요청에 실패했습니다.', 'NAVER_FAILED');
  } finally { clearTimeout(timer); }
}
