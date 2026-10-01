// 공개 이미지 검색. 반환 문자열은 HTML이 아니며, UI는 textContent로 표시한다.
export const ONLINE_IMAGE_SOURCES = [
  { id: 'all', label: '모든 검색 사이트' },
  { id: 'openverse', label: 'Openverse (Flickr·박물관 등)' },
  { id: 'wikimedia', label: 'Wikimedia Commons' },
  { id: 'inaturalist', label: 'iNaturalist' },
  { id: 'nasa', label: 'NASA 이미지' },
];

const CC_CODES = ['by', 'by-sa', 'by-nc', 'by-nd', 'by-nc-sa', 'by-nc-nd'];
const CC_VERSIONS = '(?:1\\.0|2\\.0|2\\.1|2\\.5|3\\.0|4\\.0)';
const CC_PATH = new RegExp(`^/licenses/(${CC_CODES.join('|')})/(${CC_VERSIONS})(?:/[a-z-]+)?/?$`, 'i');

function webUrl(value) {
  if (typeof value !== 'string' || !/^https?:\/\//i.test(value.trim()) || /[\u0000-\u001f\u007f]/.test(value)) return '';
  try {
    const u = new URL(value.trim());
    return /^(https?:)$/.test(u.protocol) && !u.username && !u.password ? u.href : '';
  } catch { return ''; }
}

function text(value) {
  return String(value ?? '').replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '').replace(/<[^>]*>/g, '').replace(/&#(x[0-9a-f]+|\d+);/gi, (m, n) => {
    const code = n[0].toLowerCase() === 'x' ? parseInt(n.slice(1), 16) : Number(n);
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : m;
  }).replace(/&(amp|quot|apos|lt|gt|nbsp);/g, (_, k) => ({ amp: '&', quot: '"', apos: "'", lt: '<', gt: '>', nbsp: ' ' })[k]).replace(/\s+/g, ' ').trim().slice(0, 2000);
}

// URL이 제공된 경우 호스트와 경로를 먼저 검증한다. PDM/PD는 CC 체크에 포함하지 않는다.
function licenseInfo(code, url, label) {
  const licenseUrl = webUrl(url);
  const normalized = text(code).toLowerCase().replace(/^cc[ -](?=by)/, '');
  const match = normalized.match(new RegExp(`^(${CC_CODES.join('|')})(?:[ -]${CC_VERSIONS})?$`, 'i'));
  let cc = normalized === 'cc0' || normalized === 'cc0-1.0' || normalized === 'cc0 1.0' || !!match;
  if (url) {
    cc = false;
    if (licenseUrl) {
      const u = new URL(licenseUrl);
      cc = /^(www\.)?creativecommons\.org$/i.test(u.hostname) && !u.port && (CC_PATH.test(u.pathname) || /^\/publicdomain\/zero\/1\.0\/?$/i.test(u.pathname));
    }
  }
  return { license: text(label || code) || '라이선스 미확인', licenseUrl, cc };
}

function ccUrl(code) {
  const c = String(code ?? '').toLowerCase().replace(/^cc-/, '');
  return c === 'cc0' ? 'https://creativecommons.org/publicdomain/zero/1.0/' : CC_CODES.includes(c) ? `https://creativecommons.org/licenses/${c}/4.0/` : '';
}

function item(source, data) {
  const full = webUrl(data.full), thumb = webUrl(data.thumb) || full;
  if (!full || !thumb) return null;
  return { id: `${source}:${text(data.id) || full}`, thumb, full, title: text(data.title) || '제목 없는 그림', credit: text(data.credit), page: webUrl(data.page), license: data.license.license, licenseUrl: data.license.licenseUrl, cc: data.license.cc, source, provider: text(data.provider) };
}

function queryUrl(base, params) {
  const url = new URL(base);
  for (const [key, value] of Object.entries(params)) if (value !== undefined) url.searchParams.set(key, String(value));
  return url.href;
}

function abortError() { const e = new Error('그림 검색을 취소했습니다.'); e.name = 'AbortError'; return e; }

async function requestJson(url, { signal, fetcher }) {
  if (signal?.aborted) throw abortError();
  const controller = new AbortController();
  let timer, onAbort;
  const stopped = new Promise((resolve, reject) => {
    onAbort = () => { controller.abort(); reject(abortError()); };
    signal?.addEventListener('abort', onAbort, { once: true });
    timer = setTimeout(() => { controller.abort(); reject(new Error('그림 검색 응답 시간이 초과되었습니다.')); }, 10000);
  });
  try {
    return await Promise.race([
      stopped,
      Promise.resolve().then(async () => {
        if (signal?.aborted) throw abortError();
        const response = await fetcher(url, { method: 'GET', mode: 'cors', credentials: 'omit', signal: controller.signal });
        if (!response?.ok) throw new Error('그림 검색 사이트에 연결하지 못했습니다.');
        const json = await response.json();
        if (!json || typeof json !== 'object' || json.error || json.errors) throw new Error('그림 검색 결과를 읽지 못했습니다.');
        return json;
      }),
    ]);
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', onAbort);
  }
}

async function openverse(q, options) {
  const { page, ccOnly } = options;
  // 인증 없는 Openverse 요청은 page_size 20까지만 허용한다(30은 HTTP 401).
  const json = await requestJson(queryUrl('https://api.openverse.org/v1/images/', { q, page, page_size: 20, license: ccOnly ? [...CC_CODES, 'cc0'].join(',') : undefined }), options);
  if (!Array.isArray(json.results)) throw new Error('Openverse 검색 결과 형식이 올바르지 않습니다.');
  const items = json.results.slice(0, 20).map(x => {
    const license = licenseInfo(x.license, x.license_url, x.license === 'cc0' ? 'CC0' : x.license?.startsWith('by') ? `CC ${x.license.toUpperCase()}${x.license_version ? ` ${x.license_version}` : ''}` : x.license);
    return item('openverse', { id: x.id, thumb: x.thumbnail, full: x.url, title: x.title || q, credit: [x.creator, license.license].filter(Boolean).join(' · '), page: x.foreign_landing_url, license, provider: x.source || x.provider });
  });
  const pages = Number(json.page_count);
  return { items, hasMore: Number.isFinite(pages) ? page < pages : !!json.next || json.results.length === 20 };
}

async function wikimedia(q, options) {
  const json = await requestJson(queryUrl('https://commons.wikimedia.org/w/api.php', { action: 'query', format: 'json', origin: '*', generator: 'search', gsrnamespace: 6, gsrlimit: 30, gsroffset: (options.page - 1) * 30, gsrsearch: q, prop: 'imageinfo', iiprop: 'url|extmetadata|mime', iiurlwidth: 480 }), options);
  // MediaWiki는 일치하는 문서가 없으면 query 자체를 생략한다.
  const pages = Object.values(json.query?.pages ?? {}).sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
  const items = pages.slice(0, 30).map(p => {
    const ii = p.imageinfo?.[0]; if (!ii || ii.mime && !/^image\//i.test(ii.mime)) return null;
    const md = ii.extmetadata ?? {}, license = licenseInfo(md.License?.value, md.LicenseUrl?.value, md.LicenseShortName?.value);
    return item('wikimedia', { id: p.pageid, thumb: ii.thumburl, full: ii.url, title: String(p.title ?? q).replace(/^File:/i, ''), credit: [text(md.Artist?.value), license.license].filter(Boolean).join(' · '), page: ii.descriptionurl, license });
  });
  return { items, hasMore: Number.isFinite(Number(json.continue?.gsroffset)) && Number(json.continue.gsroffset) > (options.page - 1) * 30 };
}

function naturalistImage(url, size) {
  const safe = webUrl(url); if (!safe) return '';
  const u = new URL(safe);
  // iNaturalist가 실제로 제공하는 사진 크기 규칙에만 적용한다.
  if (/^(?:static\.inaturalist\.org|inaturalist-open-data\.s3\.amazonaws\.com)$/i.test(u.hostname) && /^\/photos\/\d+\/(?:square|small|medium|large|original)\.(?:jpe?g|png|webp)$/i.test(u.pathname)) u.pathname = u.pathname.replace(/\/(?:square|small|medium|large|original)(?=\.)/i, `/${size}`);
  return u.href;
}

async function inaturalist(q, options) {
  const json = await requestJson(queryUrl('https://api.inaturalist.org/v1/observations', { q, photos: true, per_page: 20, page: options.page, locale: 'ko', photo_license: options.ccOnly ? ['cc0', ...CC_CODES.map(c => `cc-${c}`)].join(',') : undefined }), options);
  if (!Array.isArray(json.results)) throw new Error('iNaturalist 검색 결과 형식이 올바르지 않습니다.');
  const items = [];
  for (const observation of json.results.slice(0, 20)) {
    for (const photo of observation.photos ?? []) {
      if (!photo || photo.hidden || photo.is_placeholder || !/^\d+$/.test(String(photo.id ?? '')) || items.length >= 40) continue;
      const license = licenseInfo(photo.license_code, '', photo.license_code);
      // 관찰 기록의 license_code는 사진 저작권과 별개이므로 사용하지 않는다.
      license.licenseUrl = ccUrl(photo.license_code);
      if (options.ccOnly && !license.cc) continue;
      const value = item('inaturalist', { id: photo.id, thumb: naturalistImage(photo.url, 'medium'), full: webUrl(photo.original_url) || naturalistImage(photo.url, 'original'), title: observation.taxon?.preferred_common_name || observation.taxon?.name || q, credit: photo.attribution, page: /^\d+$/.test(String(photo.id ?? '')) ? `https://www.inaturalist.org/photos/${photo.id}` : '', license });
      if (value) items.push(value);
    }
  }
  const total = Number(json.total_results);
  return { items, hasMore: Number.isFinite(total) ? options.page * 20 < total : json.results.length === 20 };
}

async function nasa(q, options) {
  const json = await requestJson(queryUrl('https://images-api.nasa.gov/search', { q, media_type: 'image', page_size: 30, page: options.page }), options);
  if (!Array.isArray(json.collection?.items)) throw new Error('NASA 검색 결과 형식이 올바르지 않습니다.');
  const items = json.collection.items.slice(0, 30).map(x => {
    const data = x.data?.find(d => d.media_type === 'image') ?? x.data?.[0];
    if (!data || data.media_type && data.media_type !== 'image') return null;
    const links = (x.links ?? []).filter(l => webUrl(l.href) && (l.render === 'image' || /\.(jpe?g|png|webp)(?:[?#]|$)/i.test(l.href)));
    const preview = links.find(l => l.rel === 'preview');
    const alternate = links.filter(l => l.rel === 'alternate').sort((a, b) => (Number(b.width) || 0) * (Number(b.height) || 1) - (Number(a.width) || 0) * (Number(a.height) || 1));
    const full = links.find(l => l.rel === 'canonical') ?? alternate[0] ?? preview ?? links[0];
    return item('nasa', { id: data.nasa_id, thumb: preview?.href || full?.href, full: full?.href, title: data.title || q, credit: [data.photographer || data.secondary_creator, data.center || 'NASA'].filter(Boolean).join(' · '), page: data.nasa_id ? `https://images.nasa.gov/details/${encodeURIComponent(data.nasa_id)}` : '', license: { license: 'NASA 이용 조건 확인', licenseUrl: 'https://www.nasa.gov/nasa-brand-center/images-and-media/', cc: false } });
  });
  const total = Number(json.collection.metadata?.total_hits);
  return { items, hasMore: (json.collection.links ?? []).some(l => l.rel === 'next') || Number.isFinite(total) && options.page * 30 < total };
}

/** 공개 검색 API만 사용한다. 삽입/다운로드·사용자 파일 전송은 호출자의 별도 동작이다. */
export async function searchOnlineImages(q, { page = 1, source = 'all', ccOnly = false, signal, fetcher = fetch } = {}) {
  if (signal?.aborted) throw abortError();
  if (typeof q !== 'string' || q.trim().length > 200) throw new Error('검색어는 200자 이내로 입력하세요.');
  if (!Number.isInteger(page) || page < 1 || page > 50) throw new Error('검색 페이지는 1부터 50까지 지정하세요.');
  if (!ONLINE_IMAGE_SOURCES.some(s => s.id === source)) throw new Error('지원하지 않는 그림 검색 사이트입니다.');
  q = q.trim();
  if (!q || source === 'nasa' && ccOnly) return { items: [], hasMore: false, failures: [] };
  const adapters = { openverse, wikimedia, inaturalist, nasa };
  const sources = ONLINE_IMAGE_SOURCES.filter(s => s.id !== 'all' && (source === 'all' || source === s.id) && !(ccOnly && s.id === 'nasa'));
  const results = await Promise.allSettled(sources.map(s => adapters[s.id](q, { page, ccOnly, signal, fetcher })));
  if (signal?.aborted) throw abortError();
  const failures = [], lists = []; let hasMore = false;
  for (let i = 0; i < results.length; i++) {
    const result = results[i];
    if (result.status === 'rejected') { failures.push(sources[i].label); continue; }
    hasMore ||= result.value.hasMore;
    lists.push(result.value.items.filter(x => x && (!ccOnly || x.cc)));
  }
  if (failures.length === sources.length) throw new Error(`그림 검색 사이트에 연결하지 못했습니다. 잠시 후 다시 시도하세요. (${failures.join(', ')})`);
  const items = [], seen = new Set();
  for (let row = 0; lists.some(list => row < list.length); row++) {
    for (const list of lists) {
      const value = list[row]; if (!value) continue;
      const key = new URL(value.full); key.hash = '';
      if (seen.has(key.href)) continue;
      seen.add(key.href); items.push(value);
    }
  }
  return { items, hasMore: page < 50 && hasMore, failures };
}
