// Node와 Worker가 공유하는 고정 공급자 API. 사용자 지정 URL/인증 헤더는 전달하지 않는다.
import { STOCK_MEDIA_PROVIDERS, normalizeStockMedia, mediaSiteSearch } from '../src/stock-media.js';
const cache = new Map();
function fail(status, message) { return Object.assign(new Error(message), { status }); }
async function jsonRequest(url, headers, fetcher) {
  const controller = new AbortController(); let timer;
  const timeout = new Promise((resolve, reject) => { timer = setTimeout(() => { controller.abort(); reject(fail(504, '미디어 검색 응답 시간이 초과되었습니다.')); }, 10000); });
  try {
    return await Promise.race([timeout, (async () => {
    const response = await fetcher(url.href, { method: 'GET', redirect: 'error', credentials: 'omit', signal: controller.signal, headers: { Accept: 'application/json', ...headers } });
    if (!response.ok) throw fail(response.status === 429 ? 429 : 502, response.status === 429 ? '검색 제공자 요청 한도에 도달했습니다. 잠시 후 다시 시도하세요.' : '검색 제공자에 연결하지 못했습니다. 서버의 API 키와 이용 한도를 확인하세요.');
    if (Number(response.headers.get('content-length')) > 2 * 1024 * 1024) throw fail(502, '검색 응답이 너무 큽니다.');
    const reader = response.body.getReader(), chunks = []; let size = 0;
    const abort = () => { reader.cancel().catch(() => {}); }; controller.signal.addEventListener('abort', abort, { once: true });
    try { for (;;) { if (controller.signal.aborted) throw fail(504, '미디어 검색 응답 시간이 초과되었습니다.'); const { value, done } = await reader.read(); if (done) break; size += value.byteLength; if (size > 2 * 1024 * 1024) throw fail(502, '검색 응답이 너무 큽니다.'); chunks.push(value); } }
    finally { controller.signal.removeEventListener('abort', abort); await reader.cancel().catch(() => {}); reader.releaseLock(); }
    if (controller.signal.aborted) throw fail(504, '미디어 검색 응답 시간이 초과되었습니다.');
    const bytes = new Uint8Array(size); let at = 0; for (const part of chunks) { bytes.set(part, at); at += part.length; }
    return JSON.parse(new TextDecoder().decode(bytes));
    })()]);
  } catch (e) { if (e.status) throw e; throw fail(controller.signal.aborted ? 504 : 502, controller.signal.aborted ? '미디어 검색 응답 시간이 초과되었습니다.' : '미디어 검색 응답을 읽지 못했습니다.'); }
  finally { clearTimeout(timer); }
}
export async function mediaSearch(params, env = {}, fetcher = fetch) {
  const source = params.get('source'), provider = STOCK_MEDIA_PROVIDERS.find(x => x.id === source), kind = params.get('type') || 'image';
  const q = params.get('q')?.trim() || '', page = Number(params.get('page') || 1);
  if (!provider || !['image', 'video', 'gif'].includes(kind) || !q || q.length > 200 || !Number.isInteger(page) || page < 1 || page > 50) throw fail(400, '검색 사이트·종류·검색어(200자 이내)·페이지를 확인하세요.');
  const info = { source, items: [], hasMore: false, searchUrl: mediaSiteSearch(source, q, kind), setupUrl: provider.docs };
  if (!provider.kinds.includes(kind)) return { ...info, status: 'unsupported', message: `${provider.label} API는 이 미디어 종류를 지원하지 않습니다.` };
  const key = typeof env[provider.env] === 'string' ? env[provider.env].trim() : '';
  if (!key) return { ...info, status: 'unconfigured', message: `${provider.label}: 서버에 ${provider.env}를 설정하면 앱 안에서 검색할 수 있습니다.` };
  if (key.length > 2048 || /[\r\n]/.test(key)) throw fail(503, '서버의 미디어 API 키 설정을 확인하세요.');
  if (source === 'pixabay' && q.length > 100) throw fail(400, 'Pixabay 검색어는 100자 이내로 입력하세요.');
  // Pixabay API 권고대로 검색 결과를 24시간 캐시. 반환/로그에 키를 포함하지 않는다.
  const cacheKey = `${source}:${kind}:${q}:${page}`, stored = cache.get(cacheKey);
  if (stored && stored.key === key && stored.expires > Date.now()) return structuredClone(stored.value);
  const url = new URL(source === 'unsplash' ? 'https://api.unsplash.com/search/photos' : source === 'pexels' ? `https://api.pexels.com/v1/${kind === 'video' ? 'videos/search' : 'search'}` : `https://pixabay.com/api/${kind === 'video' ? 'videos/' : ''}`);
  url.searchParams.set(source === 'pixabay' ? 'q' : 'query', q); url.searchParams.set('page', String(page)); url.searchParams.set('per_page', '30');
  const headers = source === 'unsplash' ? { Authorization: `Client-ID ${key}`, 'Accept-Version': 'v1' } : source === 'pexels' ? { Authorization: key } : {};
  if (source === 'pixabay') { url.searchParams.set('key', key); url.searchParams.set('safesearch', 'true'); }
  if (source === 'unsplash') url.searchParams.set('content_filter', 'high');
  const json = await jsonRequest(url, headers, fetcher), items = normalizeStockMedia(source, json, kind);
  const total = Number(source === 'unsplash' ? json.total : source === 'pexels' ? json.total_results : json.totalHits);
  const value = { source, status: 'ok', items, hasMore: page < 50 && (Number.isFinite(total) ? page * 30 < total : !!json.next_page) };
  if (cache.size >= 150) cache.delete(cache.keys().next().value);
  cache.set(cacheKey, { key, value, expires: Date.now() + (source === 'pixabay' ? 86400000 : 300000) });
  return structuredClone(value);
}
/** 사용자 삽입 이벤트만 호출한다. URL 대신 검증된 사진 ID로 공식 download 경로를 만든다. */
export async function trackMediaDownload(params, env = {}, fetcher = fetch) {
  const id = params.get('id');
  if (params.get('source') !== 'unsplash' || !/^[\w-]{1,100}$/.test(id || '')) throw fail(400, '그림 출처와 ID를 확인하세요.');
  const key = env.UNSPLASH_ACCESS_KEY;
  if (typeof key !== 'string' || !key.trim() || key.length > 2048 || /[\r\n]/.test(key)) throw fail(503, 'Unsplash API 키를 서버에 설정하세요.');
  await jsonRequest(new URL(`https://api.unsplash.com/photos/${encodeURIComponent(id)}/download`), { Authorization: `Client-ID ${key.trim()}`, 'Accept-Version': 'v1' }, fetcher);
  return { ok: true };
}
