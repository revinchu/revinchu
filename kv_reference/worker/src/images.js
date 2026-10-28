/* 이미지 검색 출처들. 각 출처는 썸네일 URL과 원본 페이지 링크만 돌려준다.
 * 키가 필요 없는 출처: Openverse(크리에이티브 커먼즈 이미지 수억 장), Wikimedia Commons, Art Institute of Chicago, The Met.
 * 키를 넣으면 켜지는 출처: Brave 이미지 검색(웹 전체), Google 프로그래머블 검색(웹 전체), Unsplash, Pexels, Pixabay. */

const UA = 'KVAtlas/1.0 (ad key visual reference tool)';
const get = async (url, headers = {}) => {
  const res = await fetch(url, { headers: { 'user-agent': UA, accept: 'application/json', ...headers }, signal: AbortSignal.timeout(9000) });
  if (!res.ok) throw new Error(url.split('?')[0] + ' ' + res.status);
  return res.json();
};
const item = (source, o) => ({ source, thumb: o.thumb, url: o.url, title: (o.title || '').slice(0, 160), by: (o.by || '').slice(0, 80), w: o.w || 0, h: o.h || 0 });

export const SOURCES = [
  {
    id: 'openverse', name: 'Openverse', enabled: () => true,
    async search(q, page) {
      const d = await get('https://api.openverse.org/v1/images/?q=' + encodeURIComponent(q) + '&page_size=20&page=' + page);
      return (d.results || []).map((r) => item('Openverse', { thumb: r.thumbnail || r.url, url: r.foreign_landing_url || r.url, title: r.title, by: r.creator, w: r.width, h: r.height }));
    },
  },
  {
    id: 'wikimedia', name: 'Wikimedia Commons', enabled: () => true,
    async search(q, page) {
      const d = await get('https://commons.wikimedia.org/w/api.php?action=query&format=json&generator=search&gsrnamespace=6&gsrlimit=20&gsroffset=' + (page - 1) * 20 +
        '&gsrsearch=' + encodeURIComponent(q) + '&prop=imageinfo&iiprop=url|size&iiurlwidth=400&origin=*');
      return Object.values((d.query && d.query.pages) || {}).filter((p) => p.imageinfo && p.imageinfo[0].thumburl).map((p) => {
        const ii = p.imageinfo[0];
        return item('Wikimedia', { thumb: ii.thumburl, url: ii.descriptionurl, title: p.title.replace(/^File:/, '').replace(/\.\w+$/, ''), w: ii.thumbwidth, h: ii.thumbheight });
      });
    },
  },
  {
    id: 'aic', name: 'Art Institute of Chicago', enabled: () => true,
    async search(q, page) {
      const d = await get('https://api.artic.edu/api/v1/artworks/search?q=' + encodeURIComponent(q) + '&limit=20&page=' + page + '&fields=id,title,image_id,artist_display');
      const iiif = (d.config && d.config.iiif_url) || 'https://www.artic.edu/iiif/2';
      return (d.data || []).filter((r) => r.image_id).map((r) => item('Art Institute of Chicago', { thumb: iiif + '/' + r.image_id + '/full/400,/0/default.jpg', url: 'https://www.artic.edu/artworks/' + r.id, title: r.title, by: r.artist_display }));
    },
  },
  {
    id: 'met', name: 'The Met', enabled: () => true,
    async search(q, page) {
      const d = await get('https://collectionapi.metmuseum.org/public/collection/v1/search?hasImages=true&q=' + encodeURIComponent(q));
      const ids = (d.objectIDs || []).slice((page - 1) * 12, page * 12);
      const objs = await Promise.allSettled(ids.map((id) => get('https://collectionapi.metmuseum.org/public/collection/v1/objects/' + id)));
      return objs.filter((o) => o.status === 'fulfilled' && o.value.primaryImageSmall).map((o) => item('The Met', { thumb: o.value.primaryImageSmall, url: o.value.objectURL, title: o.value.title, by: o.value.artistDisplayName }));
    },
  },
  {
    id: 'brave', name: 'Brave 이미지 검색 (웹 전체)', enabled: (env) => !!env.BRAVE_API_KEY,
    async search(q, page, env) {
      const d = await get('https://api.search.brave.com/res/v1/images/search?count=50&q=' + encodeURIComponent(q) + '&offset=' + (page - 1), { 'x-subscription-token': env.BRAVE_API_KEY });
      return (d.results || []).map((r) => item(r.source || 'Web', { thumb: (r.thumbnail && r.thumbnail.src) || (r.properties && r.properties.url), url: r.url, title: r.title, w: r.properties && r.properties.width, h: r.properties && r.properties.height }));
    },
  },
  {
    id: 'google', name: 'Google 프로그래머블 검색 (웹 전체)', enabled: (env) => !!(env.GOOGLE_CSE_KEY && env.GOOGLE_CSE_CX),
    async search(q, page, env) {
      const d = await get('https://www.googleapis.com/customsearch/v1?searchType=image&num=10&start=' + ((page - 1) * 10 + 1) + '&key=' + env.GOOGLE_CSE_KEY + '&cx=' + env.GOOGLE_CSE_CX + '&q=' + encodeURIComponent(q));
      return (d.items || []).map((r) => item(r.displayLink || 'Google', { thumb: (r.image && r.image.thumbnailLink) || r.link, url: (r.image && r.image.contextLink) || r.link, title: r.title, w: r.image && r.image.width, h: r.image && r.image.height }));
    },
  },
  {
    id: 'unsplash', name: 'Unsplash', enabled: (env) => !!env.UNSPLASH_ACCESS_KEY,
    async search(q, page, env) {
      const d = await get('https://api.unsplash.com/search/photos?per_page=20&page=' + page + '&query=' + encodeURIComponent(q), { authorization: 'Client-ID ' + env.UNSPLASH_ACCESS_KEY });
      return (d.results || []).map((r) => item('Unsplash', { thumb: r.urls.small, url: r.links.html, title: r.alt_description, by: r.user && r.user.name, w: r.width, h: r.height }));
    },
  },
  {
    id: 'pexels', name: 'Pexels', enabled: (env) => !!env.PEXELS_API_KEY,
    async search(q, page, env) {
      const d = await get('https://api.pexels.com/v1/search?per_page=20&page=' + page + '&query=' + encodeURIComponent(q), { authorization: env.PEXELS_API_KEY });
      return (d.photos || []).map((r) => item('Pexels', { thumb: r.src.medium, url: r.url, title: r.alt, by: r.photographer, w: r.width, h: r.height }));
    },
  },
  {
    id: 'pixabay', name: 'Pixabay', enabled: (env) => !!env.PIXABAY_API_KEY,
    async search(q, page, env) {
      const d = await get('https://pixabay.com/api/?per_page=20&page=' + page + '&key=' + env.PIXABAY_API_KEY + '&q=' + encodeURIComponent(q));
      return (d.hits || []).map((r) => item('Pixabay', { thumb: r.webformatURL, url: r.pageURL, title: r.tags, by: r.user, w: r.webformatWidth, h: r.webformatHeight }));
    },
  },
];

export async function searchImages(q, wanted, page, env) {
  const active = SOURCES.filter((s) => s.enabled(env) && (!wanted.length || wanted.includes(s.id)));
  const settled = await Promise.allSettled(active.map((s) => s.search(q, page, env)));
  const status = {}; const lists = [];
  settled.forEach((r, i) => {
    const s = active[i];
    if (r.status === 'fulfilled') { status[s.id] = { ok: true, count: r.value.length }; lists.push(r.value.filter((x) => x.thumb && x.url)); }
    else status[s.id] = { ok: false, error: String(r.reason && r.reason.message || r.reason).slice(0, 160) };
  });
  // 출처별로 번갈아 섞는다 (한 출처가 화면을 독차지하지 않게)
  const results = []; const seen = new Set();
  for (let i = 0; lists.some((l) => i < l.length); i++) for (const l of lists) { const x = l[i]; if (x && !seen.has(x.thumb)) { seen.add(x.thumb); results.push(x); } }
  return { q, page, results, sources: status };
}
