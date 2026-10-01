// 온라인 그림 검색 창: 외부 출처를 표시하고 비동기 검색 결과를 현재 요청에만 반영한다.
import { el, openDialog } from './ui.js';
import { ONLINE_IMAGE_SOURCES, searchOnlineImages } from './online-images.js';

export function onlinePicturePicker({ onInsert }) {
  const query = el('input', { type: 'search', maxlength: 200, 'aria-label': '그림 검색어', placeholder: '검색어 (예: 커피, 사무실, nature)' });
  const cc = el('input', { type: 'checkbox', 'aria-label': 'Creative Commons만', accessKey: 'c' });
  const source = el('select', { 'aria-label': '검색 사이트', accessKey: 's' }, ONLINE_IMAGE_SOURCES.map(x => el('option', { value: x.id }, x.label)));
  const url = el('input', { type: 'url', 'aria-label': '그림 웹 주소', placeholder: 'https://… 그림 파일 주소' });
  const grid = el('div', { class: 'online-grid', 'aria-label': '그림 검색 결과' });
  const status = el('div', { class: 'online-status muted', role: 'status', 'aria-live': 'polite' }, '검색어를 입력하세요. 여러 사이트의 그림을 함께 검색합니다.');
  const note = el('div', { class: 'online-note muted' });
  const chosen = new Map(), seen = new Set();
  let request = 0, controller, closed = false, page = 0, last = null, hasMore = false, count = 0, failures = [];
  const search = el('button', { type: 'button', class: 'btn primary', accessKey: 'f', onclick: () => run(false) }, '검색');
  const more = el('button', { type: 'button', class: 'btn', disabled: true, onclick: () => run(true) }, '더 보기');
  const hint = () => { note.textContent = cc.checked ? 'CC·CC0 라이선스가 확인된 그림만 표시합니다. NASA 등 다른 이용 조건의 그림은 제외됩니다.' : 'CC 이외의 라이선스도 포함합니다. 출처 보기에서 각 그림의 이용 조건을 확인하세요.'; };
  const summary = () => `${count}개 그림${chosen.size ? ` · ${chosen.size}개 선택됨` : ' · 그림을 선택한 뒤 삽입하세요.'}${failures.length ? ` · 연결 실패: ${failures.join(', ')} (다른 출처의 결과 표시)` : ''}`;
  const addCard = it => {
    const clean = new URL(it.full); clean.hash = ''; const key = clean.href;
    if (seen.has(key)) return;
    seen.add(key); count++;
    const sourceName = ONLINE_IMAGE_SOURCES.find(x => x.id === it.source)?.label || it.source;
    const sourceLabel = it.provider ? `${sourceName} · ${it.provider}` : sourceName;
    const card = el('button', { type: 'button', class: 'online-item', 'aria-label': it.title, 'aria-pressed': 'false', title: [it.title, it.credit, it.source].filter(Boolean).join('\n') },
      el('img', { src: it.thumb, alt: '', loading: 'lazy', referrerpolicy: 'no-referrer' }), el('span', { class: 'online-title' }, it.title));
    card.addEventListener('click', () => {
      if (chosen.has(key)) chosen.delete(key);
      else if (chosen.size < 20) chosen.set(key, it);
      else { status.textContent = '한 번에 최대 20개까지 삽입할 수 있습니다.'; return; }
      card.classList.toggle('on', chosen.has(key)); card.setAttribute('aria-pressed', String(chosen.has(key))); status.textContent = summary();
    });
    grid.append(el('article', { class: 'online-card' }, card,
      el('span', { class: 'online-source', title: [sourceLabel, it.credit].filter(Boolean).join(' · ') }, sourceLabel),
      el('span', { class: 'online-license', title: it.credit || it.license }, it.license || '이용 조건 확인'),
      it.page ? el('a', { href: it.page, target: '_blank', rel: 'noopener noreferrer', referrerpolicy: 'no-referrer', 'aria-label': `${it.title} 출처 보기` }, '출처 보기') : null));
  };
  const run = async append => {
    if (closed || body.inert) return;
    if (append && (!hasMore || more.disabled)) return;
    const text = query.value.trim();
    if (!append && !text) { query.focus(); status.textContent = '검색어를 입력하세요.'; return; }
    const id = ++request; controller?.abort(); controller = new AbortController();
    if (!append) { last = { q: text, source: source.value, ccOnly: cc.checked }; page = 0; hasMore = false; failures = []; count = 0; chosen.clear(); seen.clear(); grid.replaceChildren(); grid.scrollTop = 0; }
    const next = page + 1;
    more.disabled = true; grid.setAttribute('aria-busy', 'true'); status.textContent = '검색 중…';
    try {
      const result = await searchOnlineImages(last.q, { ...last, page: next, signal: controller.signal });
      if (closed || request !== id) return;
      page = next; hasMore = result.hasMore && page < 50; failures = result.failures;
      result.items.forEach(addCard);
      status.textContent = count ? summary() : (last.ccOnly && last.source === 'nasa' ? 'NASA는 CC 라이선스 전용 검색에 포함되지 않습니다. 다른 사이트를 선택하거나 체크를 해제하세요.' : `결과가 없습니다. 다른 검색어 또는 영어 검색어도 사용해 보세요.${failures.length ? ` 연결 실패: ${failures.join(', ')}.` : ''}`);
    } catch (error) {
      if (closed || request !== id || error.name === 'AbortError') return;
      status.textContent = `검색할 수 없습니다: ${error.message} 다시 검색하거나 웹 주소로 삽입하세요.`;
    } finally {
      if (!closed && request === id) { grid.removeAttribute('aria-busy'); more.disabled = !hasMore; }
    }
  };
  query.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.isComposing) { e.preventDefault(); e.stopPropagation(); run(false); } });
  const changeFilter = () => {
    hint();
    if (query.value.trim()) { run(false); return; }
    request++; controller?.abort(); last = null; page = 0; count = 0; hasMore = false; failures = [];
    chosen.clear(); seen.clear(); grid.replaceChildren(); grid.removeAttribute('aria-busy'); more.disabled = true;
    status.textContent = '검색어를 입력하세요.';
  };
  cc.addEventListener('change', changeFilter);
  source.addEventListener('change', changeFilter);
  hint();
  const body = el('div', { class: 'online-picture-dialog' },
    el('div', { class: 'online-search-row' }, query, search),
    el('div', { class: 'online-filter-row' }, el('label', {}, cc, 'Creative Commons만'), el('label', {}, '검색 사이트 ', source)), note,
    grid, el('div', { class: 'online-search-footer' }, status, more),
    el('label', { class: 'online-url-row' }, '웹 주소로 삽입', url), el('div', { class: 'online-note muted' }, 'NASA 등 외부 다운로드가 제한된 그림은 웹 연결로 삽입되며, 파일에 이미지가 포함되지 않을 수 있습니다.'));
  const dialog = openDialog({ title: '온라인 그림', width: 780, body, onClose: () => { closed = true; request++; controller?.abort(); }, buttons: [
    { label: '삽입', primary: true, action: async () => {
      const list = [...chosen.values()], direct = url.value.trim();
      if (direct) {
        let parsed; try { parsed = new URL(direct); } catch { throw new Error('올바른 https:// 또는 http:// 그림 주소를 입력하세요.'); }
        if (!['https:', 'http:'].includes(parsed.protocol) || parsed.username || parsed.password) throw new Error('https:// 또는 http:// 그림 주소만 사용할 수 있습니다.');
        list.push({ full: parsed.href, title: decodeURI(parsed.pathname.split('/').pop() || '그림') });
      }
      if (!list.length) { status.textContent = '그림을 선택하거나 웹 주소를 입력하세요.'; return false; }
      request++; controller?.abort(); body.inert = true; status.textContent = `${list.length}개 그림을 불러오는 중…`;
      try { return await onInsert(list); } finally { body.inert = false; }
    } }, { label: '취소' },
  ] });
  return dialog;
}
