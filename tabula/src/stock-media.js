// 공급자 공식 API의 공개 메타데이터만 정규화한다. API 키는 서버에서만 사용한다.
export const STOCK_MEDIA_PROVIDERS = [
  { id: 'unsplash', label: 'Unsplash', env: 'UNSPLASH_ACCESS_KEY', docs: 'https://unsplash.com/documentation', kinds: ['image'] },
  { id: 'pexels', label: 'Pexels', env: 'PEXELS_API_KEY', docs: 'https://www.pexels.com/api/documentation/', kinds: ['image', 'video'] },
  { id: 'pixabay', label: 'Pixabay', env: 'PIXABAY_API_KEY', docs: 'https://pixabay.com/api/docs/', kinds: ['image', 'video'] },
];
export function safeMediaUrl(value) {
  if (typeof value !== 'string' || !/^https?:\/\//i.test(value.trim()) || /[\u0000-\u001f\u007f]/.test(value)) return '';
  try { const u = new URL(value.trim()); return /^https?:$/.test(u.protocol) && !u.username && !u.password ? u.href : ''; } catch { return ''; }
}
export function mediaSiteSearch(source, q, kind = 'image') {
  const query = encodeURIComponent(String(q || '').trim().slice(0, 200));
  if (source === 'unsplash') return `https://unsplash.com/s/photos/${query}`;
  if (source === 'pexels') return `https://www.pexels.com/search/${kind === 'video' ? 'videos/' : ''}${query}/`;
  if (source === 'pixabay') return `https://pixabay.com/${kind === 'video' ? 'videos/' : ''}search/${query}/`;
  return '';
}
const clean = value => String(value ?? '').replace(/<[^>]*>/g, '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, 2000);
const positive = n => Number.isFinite(Number(n)) && Number(n) > 0 ? Number(n) : undefined;
const utm = value => { const s = safeMediaUrl(value); if (!s) return ''; const u = new URL(s); u.searchParams.set('utm_source', 'wixel'); u.searchParams.set('utm_medium', 'referral'); return u.href; };
function normalized(source, data) {
  const full = safeMediaUrl(data.full), thumb = safeMediaUrl(data.thumb), page = safeMediaUrl(data.page);
  if (!full || !thumb) return null;
  return { id: `${source}:${clean(data.id)}`, kind: data.kind, full, thumb, poster: data.kind === 'video' ? thumb : undefined,
    mime: data.mime || '', title: clean(data.title) || '제목 없는 미디어', credit: clean(data.credit), page, source, provider: source,
    width: positive(data.width), height: positive(data.height), duration: positive(data.duration),
    license: data.license, licenseUrl: data.licenseUrl, cc: false, ...(data.extra || {}) };
}
/** Photo API 검색은 CC 라이선스와 다르다. CC로 추정하지 않는다. */
export function normalizeStockMedia(source, json, kind = 'image') {
  const out = [];
  const provider = STOCK_MEDIA_PROVIDERS.find(x => x.id === source);
  if (!provider?.kinds.includes(kind)) return out;
  const rows = source === 'unsplash' ? json.results : source === 'pexels' ? (kind === 'video' ? json.videos : json.photos) : json.hits;
  if (!Array.isArray(rows)) throw new Error('미디어 검색 결과 형식이 올바르지 않습니다.');
  for (const x of rows.slice(0, 30)) {
    let data;
    if (source === 'unsplash') {
      data = { id: x.id, kind, full: x.urls?.regular || x.urls?.full, thumb: x.urls?.small || x.urls?.thumb,
        title: x.alt_description || x.description, credit: `${clean(x.user?.name)} / Unsplash`, page: utm(x.links?.html), width: x.width, height: x.height,
        license: 'Unsplash 라이선스', licenseUrl: 'https://unsplash.com/license', extra: { linkOnly: true, trackId: clean(x.id), creatorUrl: utm(x.user?.links?.html) } };
    } else if (source === 'pexels') {
      const files = (x.video_files || []).filter(v => v.file_type === 'video/mp4' && safeMediaUrl(v.link));
      files.sort((a, b) => Math.abs((Number(a.width) || 1280) - 1280) - Math.abs((Number(b.width) || 1280) - 1280));
      const video = files[0];
      data = { id: x.id, kind, full: kind === 'video' ? video?.link : x.src?.original, thumb: kind === 'video' ? x.image : x.src?.medium || x.src?.small,
        title: x.alt || (kind === 'video' ? `Pexels 영상 ${x.id}` : ''), credit: `${clean(x.photographer || x.user?.name)} / Pexels`, page: x.url,
        width: video?.width || x.width, height: video?.height || x.height, duration: x.duration, mime: kind === 'video' ? 'video/mp4' : '',
        license: 'Pexels 라이선스', licenseUrl: 'https://www.pexels.com/license/', extra: { creatorUrl: safeMediaUrl(x.photographer_url || x.user?.url) } };
    } else {
      const video = ['medium', 'small', 'large', 'tiny'].map(k => x.videos?.[k]).find(v => safeMediaUrl(v?.url));
      data = { id: x.id, kind, full: kind === 'video' ? video?.url : x.largeImageURL || x.webformatURL,
        thumb: kind === 'video' ? video?.thumbnail : x.previewURL || x.webformatURL, title: x.tags, credit: `${clean(x.user)} / Pixabay`, page: x.pageURL,
        width: video?.width || x.imageWidth, height: video?.height || x.imageHeight, duration: x.duration, mime: kind === 'video' ? 'video/mp4' : '',
        license: 'Pixabay 콘텐츠 라이선스', licenseUrl: 'https://pixabay.com/service/license-summary/', extra: kind === 'image' ? { embedRequired: true } : {} };
    }
    const it = normalized(source, data); if (it) out.push(it);
  }
  return out;
}
