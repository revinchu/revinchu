// 온라인 그림 · 스톡 이미지 검색 (키 없이 쓰는 공개 API) — DOM 없음
//   Openverse (CC 라이선스 · 상업적 이용 가능만), Wikimedia Commons. 결과는 같은 모양으로 정리:
//   { id, title, thumb, full: [시도할 원본 URL 순서], w, h, creator, license, licenseUrl, page, source }
export const STOCK_CATEGORIES = [['비즈니스', 'business office'], ['사람', 'people'], ['자연', 'nature landscape'], ['기술', 'technology computer'], ['음식', 'food'], ['여행', 'travel'], ['도시', 'city'], ['동물', 'animal'], ['교육', 'education school'], ['배경', 'abstract background'], ['한국', 'korea seoul'], ['스포츠', 'sport']];

export const openverseUrl = (q, page = 1) => `https://api.openverse.org/v1/images/?q=${encodeURIComponent(q)}&page=${page}&page_size=40&license_type=commercial&mature=false`;
export function parseOpenverse(json) {
  return (json?.results ?? []).map((r) => ({
    id: `ov:${r.id}`, title: r.title || '', source: 'Openverse',
    thumb: r.thumbnail || `https://api.openverse.org/v1/images/${r.id}/thumb/`,
    full: [r.url, `https://api.openverse.org/v1/images/${r.id}/thumb/?full_size=true`, r.thumbnail].filter(Boolean),
    w: r.width ?? 0, h: r.height ?? 0, creator: r.creator || '', license: `${(r.license ?? '').toUpperCase() === 'CC0' || r.license === 'pdm' ? (r.license === 'pdm' ? '퍼블릭 도메인' : 'CC0') : `CC ${String(r.license ?? '').toUpperCase()} ${r.license_version ?? ''}`.trim()}`,
    licenseUrl: r.license_url || '', page: r.foreign_landing_url || '', free: r.license === 'cc0' || r.license === 'pdm',
  }));
}
export const wikimediaUrl = (q, offset = 0) => `https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch=${encodeURIComponent(`${q} filetype:bitmap`)}&gsrnamespace=6&gsrlimit=40&gsroffset=${offset}&prop=imageinfo&iiprop=url|size|extmetadata|mime&iiurlwidth=360&format=json&origin=*`;
const stripHtml = (s) => String(s ?? '').replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#039;/g, "'").trim();
export function parseWikimedia(json) {
  const pages = Object.values(json?.query?.pages ?? {}).sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
  return pages.filter((p) => p.imageinfo?.[0] && /^image\/(jpeg|png|gif|webp)/.test(p.imageinfo[0].mime ?? 'image/jpeg')).map((p) => {
    const ii = p.imageinfo[0];
    const md = ii.extmetadata ?? {};
    const lic = stripHtml(md.LicenseShortName?.value);
    return {
      id: `wm:${p.pageid}`, title: stripHtml(md.ObjectName?.value) || String(p.title ?? '').replace(/^File:|\.[a-z]+$/gi, ''), source: 'Wikimedia Commons',
      thumb: ii.thumburl || ii.url, full: [ii.width > 2560 && ii.thumburl ? ii.thumburl.replace(/\/\d+px-/, '/2048px-') : null, ii.url, ii.thumburl].filter(Boolean),
      w: ii.width ?? 0, h: ii.height ?? 0, creator: stripHtml(md.Artist?.value), license: lic || '저작권 표시',
      licenseUrl: stripHtml(md.LicenseUrl?.value), page: ii.descriptionurl || '', free: /public domain|cc0|퍼블릭/i.test(lic),
    };
  });
}
/** 출처 표시 글 (PowerPoint 처럼 그림 아래 캡션 · 대체 텍스트) */
export const creditText = (r) => [r.title || '그림', r.creator ? `작성자: ${r.creator}` : '', r.license ? `라이선스: ${r.license}` : '', r.page].filter(Boolean).join(' · ');
