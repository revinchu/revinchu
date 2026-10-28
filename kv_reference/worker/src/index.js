/* KV 아틀라스 호스팅용 Cloudflare Worker.
 *
 *   GET  /api/health                    사용 가능한 기능
 *   GET  /api/preview?url=<페이지 URL>   그 페이지의 대표 이미지(og:image 등) URL만 돌려준다. 이미지는 브라우저가 원본 사이트에서 직접 불러온다.
 *   GET  /api/images?q=<검색어>&src=a,b  이미지 검색 결과(썸네일 URL + 원본 페이지 링크)를 여러 출처에서 모아 돌려준다.
 *   POST /api/claude                    페이지의 AI 기능이 쓰는 Claude 호출 (ACCESS_TOKEN 헤더 필요)
 *   그 외                               정적 파일(public/)
 *
 * 이미지 파일은 저장하거나 중계하지 않는다. 응답은 URL과 메타데이터뿐이다.
 */
import Anthropic from '@anthropic-ai/sdk';
import { extractPreview, normalizeTarget } from './preview.js';
import { SOURCES, searchImages } from './images.js';

const MODEL = 'claude-opus-5-5';
const EFFORT = { quick: 'low', default: 'medium', complex: 'high' };

const json = (data, status = 200, extra = {}) =>
  new Response(JSON.stringify(data), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'access-control-allow-origin': '*', ...extra } });

async function cached(request, key, ttl, produce) {
  const cache = caches.default;
  const cacheKey = new Request('https://kv-atlas-cache.internal/' + key);
  const hit = await cache.match(cacheKey);
  if (hit) return hit;
  const res = await produce();
  if (res.status === 200) {
    const copy = new Response(res.body, res);
    copy.headers.set('cache-control', 'public, max-age=' + ttl);
    await cache.put(cacheKey, copy.clone());
    return copy;
  }
  return res;
}

async function preview(url) {
  const target = normalizeTarget(url);
  if (!target) return json({ error: 'http(s) 공개 주소만 미리보기할 수 있습니다.' }, 400);
  const res = await fetch(target, {
    headers: { 'user-agent': 'Mozilla/5.0 (compatible; KVAtlasPreview/1.0; +link preview)', accept: 'text/html,application/xhtml+xml' },
    redirect: 'follow', signal: AbortSignal.timeout(8000),
  });
  if (!res.ok) return json({ error: 'status ' + res.status }, 502);
  const type = res.headers.get('content-type') || '';
  if (type.startsWith('image/')) return json({ image: res.url, title: '', site: new URL(res.url).hostname });
  if (!type.includes('html')) return json({ error: 'html 아님' }, 422);
  // 앞부분 600KB만 읽는다 (메타 태그는 head 에 있다)
  const reader = res.body.getReader(); const chunks = []; let size = 0;
  while (size < 600000) { const { done, value } = await reader.read(); if (done) break; chunks.push(value); size += value.length; }
  reader.cancel().catch(() => {});
  const html = new TextDecoder().decode(concat(chunks));
  return json(extractPreview(html, res.url));
}
const concat = (parts) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };

async function claude(request, env) {
  if (!env.ANTHROPIC_API_KEY || !env.ACCESS_TOKEN) return json({ code: 'sampling_disabled', message: '서버에 ANTHROPIC_API_KEY와 ACCESS_TOKEN이 설정되지 않았습니다.' }, 503);
  if (request.headers.get('x-kva-token') !== env.ACCESS_TOKEN) return json({ code: 'token_required', message: '접근 토큰이 맞지 않습니다.' }, 401);
  const body = await request.json().catch(() => null);
  if (!body || !(body.prompt || (body.messages && body.messages.length))) return json({ code: 'invalid_request', message: '빈 요청입니다.' }, 400);
  const messages = body.messages ? body.messages.map((m) => ({ role: m.role, content: String(m.content) })) : [{ role: 'user', content: String(body.prompt) }];
  const images = Array.isArray(body.images) ? body.images.slice(0, 5) : [];
  if (images.length) {
    const last = messages[messages.length - 1];
    last.content = images.map((im) => ({ type: 'image', source: { type: 'base64', media_type: im.media_type, data: im.data } })).concat([{ type: 'text', text: last.content }]);
  }
  const client = new Anthropic({ apiKey: env.ANTHROPIC_API_KEY });
  try {
    const response = await client.beta.messages.create({
      model: MODEL, max_tokens: 16000, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default',
      output_config: { effort: EFFORT[body.tier] || 'medium' }, messages,
    });
    if (response.stop_reason === 'refusal') return json({ code: 'refused', message: 'Claude가 이 요청에 답하지 않았습니다.' }, 422);
    const text = response.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
    if (!text.trim()) return json({ code: 'empty_completion', message: '빈 응답입니다.' }, 502);
    return json({ text, truncated: response.stop_reason === 'max_tokens' });
  } catch (e) {
    if (e instanceof Anthropic.RateLimitError) return json({ code: 'rate_limited', message: '요청이 많습니다.' }, 429);
    if (e instanceof Anthropic.BadRequestError) return json({ code: 'invalid_request', message: e.message }, 400);
    if (e instanceof Anthropic.APIError) return json({ code: 'upstream_error', message: e.message }, 502);
    throw e;
  }
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(request);
    if (request.method === 'OPTIONS') return new Response(null, { headers: { 'access-control-allow-origin': '*', 'access-control-allow-headers': 'content-type, x-kva-token', 'access-control-allow-methods': 'GET, POST' } });
    try {
      if (url.pathname === '/api/health') {
        return json({ ok: true, preview: true, claude: !!(env.ANTHROPIC_API_KEY && env.ACCESS_TOKEN), sources: SOURCES.filter((s) => s.enabled(env)).map((s) => ({ id: s.id, name: s.name })) });
      }
      if (url.pathname === '/api/preview') {
        const target = url.searchParams.get('url') || '';
        return cached(request, 'preview?u=' + encodeURIComponent(target), 7 * 86400, () => preview(target));
      }
      if (url.pathname === '/api/images') {
        const q = (url.searchParams.get('q') || '').trim().slice(0, 200);
        if (!q) return json({ error: '검색어가 없습니다.' }, 400);
        const src = (url.searchParams.get('src') || '').split(',').filter(Boolean);
        const page = Math.max(1, Math.min(20, Number(url.searchParams.get('page')) || 1));
        return cached(request, 'images?q=' + encodeURIComponent(q) + '&src=' + src.sort().join(',') + '&p=' + page, 86400, async () => json(await searchImages(q, src, page, env)));
      }
      if (url.pathname === '/api/claude' && request.method === 'POST') return claude(request, env);
      return json({ error: 'not found' }, 404);
    } catch (e) {
      return json({ error: String(e && e.message || e) }, 500);
    }
  },
};
