// 온라인 그림·GIF·영상 검색. 출처/권한 상태를 구분하고 늦은 응답은 폐기한다.
import { el, openDialog } from './ui.js';
import { ONLINE_IMAGE_SOURCES, searchOnlineImages, resolveOnlineMedia, trackOnlineMediaInsert } from './online-images.js';
import { STOCK_MEDIA_PROVIDERS, safeMediaUrl, mediaSiteSearch } from './stock-media.js';
import { server } from './storage.js';

function mediaFetch(url, options) {
  // 검색 중계에는 Node 로그인 토큰만 사용한다. 보관함 복구키는 전송하지 않는다.
  const token = String(url).startsWith('/api/media/') ? server.headers()['X-Tabula-Token'] : '';
  return fetch(url, { ...options, ...(token ? { headers: { 'X-Tabula-Token': token } } : {}) });
}
const kindName = kind => ({ image: '그림', gif: 'GIF', video: '영상' }[kind] || '그림');
const link = (href, label) => href ? el('a', { href, target: '_blank', rel: 'noopener noreferrer', referrerpolicy: 'no-referrer' }, label) : null;

export function onlinePicturePicker({ onInsert, kind = 'image', allowVideo = true }) {
  const query = el('input', { type: 'search', maxlength: 200, 'aria-label': '그림 검색어', placeholder: '검색어 (예: 커피, 사무실, nature)' });
  const cc = el('input', { type: 'checkbox', 'aria-label': 'Creative Commons만', accessKey: 'c' });
  const source = el('select', { 'aria-label': '검색 사이트', accessKey: 's' }, ONLINE_IMAGE_SOURCES.map(x => el('option', { value: x.id }, x.label)));
  const type = el('select', { 'aria-label': '미디어 종류', accessKey: 'm' }, ['image', 'gif', ...(allowVideo ? ['video'] : [])].map(x => el('option', { value: x }, kindName(x))));
  type.value = ['image', 'gif', ...(allowVideo ? ['video'] : [])].includes(kind) ? kind : 'image';
  const url = el('input', { type: 'url', 'aria-label': '그림 웹 주소', placeholder: 'https://… 그림 파일 주소' });
  const grid = el('div', { class: 'online-grid', 'aria-label': '그림 검색 결과' });
  const status = el('div', { class: 'online-status muted', role: 'status', 'aria-live': 'polite' }, '검색어를 입력하세요. 여러 사이트의 미디어를 함께 검색합니다.');
  const note = el('div', { class: 'online-note muted' }), availability = el('div', { class: 'online-availability', 'aria-live': 'polite' });
  const chosen = new Map(), seen = new Set();
  let request = 0, controller, closed = false, page = 0, last = null, hasMore = false, count = 0, failures = [];
  const search = el('button', { type: 'button', class: 'btn primary', accessKey: 'f', onclick: () => run(false) }, '검색');
  const more = el('button', { type: 'button', class: 'btn', disabled: true, onclick: () => run(true) }, '더 보기');
  const hint = () => {
    note.textContent = cc.checked ? 'CC·CC0 라이선스가 확인된 결과만 표시합니다. Unsplash·Pexels·Pixabay·NASA는 제외됩니다.' : '출처 보기에서 이용 조건을 확인하세요. 추가 검색 사이트의 API 키는 서버에 설정합니다.';
    if (type.value === 'gif') note.textContent += ' GIF 검색은 Wikimedia의 GIF 파일을 사용합니다(정지 GIF가 포함될 수 있음).';
    if (type.value === 'video') note.textContent += ' 영상은 웹 연결로 재생하며 자동 재생하지 않습니다.';
    url.placeholder = type.value === 'video' ? 'https://… .mp4 또는 .webm 직접 영상 주소' : type.value === 'gif' ? 'https://… .gif 원본 파일 주소' : 'https://… 그림 파일 주소';
  };
  const summary = () => `${count}개 미디어${chosen.size ? ` · ${chosen.size}개 선택됨` : ' · 선택한 뒤 삽입하세요.'}${failures.length ? ` · 연결 실패: ${failures.join(', ')} (다른 출처의 결과 표시)` : ''}`;
  const showAvailability = list => {
    availability.replaceChildren(...list.map(it => el('div', { class: 'online-provider-status' }, el('span', {}, it.message || '이 출처를 사용할 수 없습니다.'), link(it.searchUrl, '사이트에서 검색'), link(it.setupUrl, 'API 설정 문서'))));
  };
  const preview = async it => {
    const abort = new AbortController(), host = el('div', { class: 'online-media-preview' }, '미리보기를 불러오는 중…'); let gone = false;
    const dlg = openDialog({ title: `${kindName(it.kind)} 미리보기`, width: 760, body: host, onClose: () => { gone = true; abort.abort(); host.querySelector('video')?.pause(); host.replaceChildren(); }, buttons: [{ label: '닫기' }] });
    try {
      const value = await resolveOnlineMedia(it, { signal: abort.signal, fetcher: mediaFetch }); if (gone || closed) return;
      const media = value.kind === 'video' ? el('video', { src: value.full, poster: value.poster || value.thumb, controls: true, playsinline: true, preload: 'metadata', 'aria-label': value.title }) : el('img', { src: value.full, alt: value.title, referrerpolicy: 'no-referrer' });
      media.addEventListener('error', () => { if (!gone) host.append(el('p', {}, '이 브라우저에서 미리보기를 불러올 수 없습니다. 출처 사이트에서 확인하세요.')); }, { once: true });
      host.replaceChildren(media, el('p', {}, value.title), el('p', { class: 'muted' }, value.credit || value.license), link(value.page, '출처 보기'));
    } catch (e) { if (!gone && e.name !== 'AbortError') host.replaceChildren(el('p', {}, e.message), link(it.page, '출처 보기')); }
    return dlg;
  };
  const addCard = it => {
    const clean = new URL(it.full); clean.hash = ''; const key = clean.href;
    if (seen.has(key)) return; seen.add(key); count++;
    const sourceName = ONLINE_IMAGE_SOURCES.find(x => x.id === it.source)?.label || it.source;
    const sourceLabel = it.provider && it.provider !== it.source ? `${sourceName} · ${it.provider}` : sourceName;
    const card = el('button', { type: 'button', class: 'online-item', 'aria-label': it.title, 'aria-pressed': 'false', title: [it.title, it.credit, sourceLabel].filter(Boolean).join('\n') },
      el('img', { src: it.thumb, alt: '', loading: 'lazy', referrerpolicy: 'no-referrer' }), el('span', { class: 'online-media-kind' }, kindName(it.kind)), el('span', { class: 'online-title' }, it.title));
    card.addEventListener('click', () => {
      if (chosen.has(key)) chosen.delete(key); else if (chosen.size < 20) chosen.set(key, it); else { status.textContent = '한 번에 최대 20개까지 삽입할 수 있습니다.'; return; }
      card.classList.toggle('on', chosen.has(key)); card.setAttribute('aria-pressed', String(chosen.has(key))); status.textContent = summary();
    });
    const details = el('div', { class: 'online-card-actions' }, el('button', { type: 'button', class: 'btn', 'aria-label': `${it.title} 미리보기`, onclick: () => preview(it) }, '미리보기'), link(it.page, '출처 보기'));
    grid.append(el('article', { class: 'online-card' }, card,
      el('span', { class: 'online-source', title: [sourceLabel, it.credit].filter(Boolean).join(' · ') }, sourceLabel),
      el('span', { class: 'online-credit' }, it.credit),
      el('span', { class: 'online-license', title: it.credit || it.license }, it.license || '이용 조건 확인'), details));
  };
  const run = async append => {
    if (closed || body.inert || append && (!hasMore || more.disabled)) return;
    const text = query.value.trim(); if (!append && !text) { query.focus(); status.textContent = '검색어를 입력하세요.'; return; }
    const id = ++request; controller?.abort(); controller = new AbortController();
    if (!append) { last = { q: text, source: source.value, ccOnly: cc.checked, kind: type.value }; page = 0; hasMore = false; failures = []; count = 0; chosen.clear(); seen.clear(); grid.replaceChildren(); availability.replaceChildren(); grid.scrollTop = 0; }
    const next = page + 1; more.disabled = true; grid.setAttribute('aria-busy', 'true'); status.textContent = '검색 중…';
    try {
      const result = await searchOnlineImages(last.q, { ...last, page: next, signal: controller.signal, fetcher: mediaFetch });
      if (closed || request !== id) return;
      page = next; hasMore = result.hasMore && page < 50; failures = result.failures; result.items.forEach(addCard); showAvailability(result.unavailable || []);
      status.textContent = count ? summary() : (last.ccOnly && ['nasa', 'unsplash', 'pexels', 'pixabay'].includes(last.source) ? '선택한 사이트는 CC 전용 검색에 포함되지 않습니다. 다른 사이트를 선택하거나 체크를 해제하세요.' : `결과가 없습니다. 다른 검색어 또는 영어 검색어도 사용해 보세요.${failures.length ? ` 연결 실패: ${failures.join(', ')}.` : ''}`);
    } catch (error) {
      if (closed || request !== id || error.name === 'AbortError') return;
      status.textContent = `검색할 수 없습니다: ${error.message} 다시 검색하거나 웹 주소로 삽입하세요.`;
      const provider = STOCK_MEDIA_PROVIDERS.find(x => x.id === source.value);
      if (provider) showAvailability([{ message: '검색 중계와 API 키 설정을 확인하세요.', searchUrl: mediaSiteSearch(provider.id, text, type.value), setupUrl: provider.docs }]);
    } finally { if (!closed && request === id) { grid.removeAttribute('aria-busy'); more.disabled = !hasMore; } }
  };
  query.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.isComposing) { e.preventDefault(); e.stopPropagation(); run(false); } });
  const changeFilter = () => {
    hint(); if (query.value.trim()) { run(false); return; }
    request++; controller?.abort(); last = null; page = 0; count = 0; hasMore = false; failures = []; chosen.clear(); seen.clear(); grid.replaceChildren(); availability.replaceChildren(); grid.removeAttribute('aria-busy'); more.disabled = true; status.textContent = '검색어를 입력하세요.';
  };
  cc.addEventListener('change', changeFilter); source.addEventListener('change', changeFilter); type.addEventListener('change', changeFilter); hint();
  const setup = () => openDialog({ title: '추가 검색 사이트 설정', width: 650, body: el('div', { class: 'online-provider-help' },
    el('p', {}, 'API 키는 WIXEL 서버 관리자만 설정합니다. 이 창이나 통합 문서에 키를 입력하지 마세요.'),
    ...STOCK_MEDIA_PROVIDERS.map(p => el('p', {}, el('strong', {}, p.label), ` — ${p.env} `, link(p.docs, '공식 API 문서'))),
    el('p', {}, 'Node: 환경 변수를 설정하고 서버를 다시 시작하세요. Cloudflare: Worker의 설정 → 변수 및 비밀에서 해당 이름의 Secret을 추가하세요.'),
    el('p', {}, '키가 없거나 단일 HTML로 실행할 때는 기존 공개 검색 사이트 또는 사이트 검색 링크·웹 주소 삽입을 사용할 수 있습니다.')),
    buttons: [{ label: '닫기' }] });
  const body = el('div', { class: 'online-picture-dialog' },
    el('div', { class: 'online-search-row' }, query, search),
    el('div', { class: 'online-filter-row' }, el('label', {}, '종류 ', type), el('label', {}, cc, 'Creative Commons만'), el('label', {}, '검색 사이트 ', source), el('button', { type: 'button', class: 'btn', onclick: setup }, '검색 사이트 설정')), note,
    availability, grid, el('div', { class: 'online-search-footer' }, status, more),
    el('label', { class: 'online-url-row' }, '웹 주소로 삽입', url), el('div', { class: 'online-note muted' }, 'GIF는 원본을 유지합니다. 영상은 웹 연결로 재생하며 XLSX에는 미리보기/링크가 남을 수 있습니다. 외부 사이트의 삭제·접속 제한에 영향을 받습니다.'));
  const dialog = openDialog({ title: '온라인 그림', width: 1000, body, onClose: () => { closed = true; request++; controller?.abort(); }, buttons: [
    { label: '삽입', primary: true, action: async () => {
      const list = [...chosen.values()], direct = url.value.trim();
      if (direct) {
        const safe = safeMediaUrl(direct); if (!safe) throw new Error('https:// 또는 http:// 미디어 주소만 사용할 수 있습니다.');
        const parsed = new URL(safe), mediaKind = type.value;
        if (mediaKind === 'video' && !/\.(?:mp4|webm)(?:$)/i.test(parsed.pathname)) throw new Error('영상은 MP4 또는 WebM 직접 파일 주소를 입력하세요. 영상 웹페이지 주소는 출처 링크로 사용하세요.');
        let title; try { title = decodeURIComponent(parsed.pathname.split('/').pop() || kindName(mediaKind)); } catch { title = kindName(mediaKind); }
        list.push({ full: safe, title, kind: mediaKind, mime: mediaKind === 'gif' ? 'image/gif' : mediaKind === 'video' ? /\.webm$/i.test(parsed.pathname) ? 'video/webm' : 'video/mp4' : '' });
      }
      if (!list.length) { status.textContent = '미디어를 선택하거나 웹 주소를 입력하세요.'; return false; }
      if (list.length > 20) throw new Error('한 번에 최대 20개까지 삽입할 수 있습니다.');
      request++; controller?.abort(); controller = new AbortController(); body.inert = true; status.textContent = `${list.length}개 미디어를 불러오는 중…`;
      try {
        const ready = [];
        for (const it of list) { const value = await resolveOnlineMedia(it, { signal: controller.signal, fetcher: mediaFetch }); if (closed) return false; await trackOnlineMediaInsert(value, { signal: controller.signal, fetcher: mediaFetch }); ready.push(value); }
        if (closed) return false; return await onInsert(ready);
      } finally { body.inert = false; }
    } }, { label: '취소' },
  ] });
  return dialog;
}
