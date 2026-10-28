/* 호스팅 모드: 같은 주소의 Worker(/api/*)가 있을 때 켜진다 (kv_reference/worker).
 * - 카드 썸네일: 이미지 주소가 없는 레퍼런스는 출처 링크의 대표 이미지(og:image)를 찾아 원본 사이트에서 직접 불러온다
 * - 통합 검색: 검색어로 여러 이미지 출처의 썸네일을 한 화면에 보여 준다
 * - Claude: claude.ai 밖에서는 window.claude 가 없으므로 Worker 의 /api/claude 로 같은 모양의 sample 함수를 만든다
 * claude.ai 아티팩트 안에서는 외부 이미지가 차단되므로 이 기능들이 켜지지 않는다. */
(function () {
  'use strict';
  const inArtifact = !!(window.claude && window.claude.use);
  const hosted = { api: null, health: null };
  const ready = (async () => {
    if (inArtifact || location.protocol === 'file:') return null;
    try {
      const res = await fetch('/api/health', { cache: 'no-store' });
      if (!res.ok) return null;
      hosted.health = await res.json();
      hosted.api = '';
      return hosted.health;
    } catch (e) { return null; }
  })();
  window.KVA_HOSTED = hosted;

  // ---------- window.claude 대체 (호스팅 모드) ----------
  const TOKEN_KEY = 'kvatlas.token';
  const getToken = () => { try { return localStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; } };
  hosted.setToken = (t) => { try { localStorage.setItem(TOKEN_KEY, t); } catch (e) { /* 저장 불가 */ } };
  async function toBase64Image(blob) {
    const bmp = await createImageBitmap(blob);
    const sc = Math.min(1, Math.sqrt(1200000 / (bmp.width * bmp.height)));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    const url = c.toDataURL('image/jpeg', 0.85);
    return { media_type: 'image/jpeg', data: url.split(',')[1] };
  }
  async function call(input, opts) {
    opts = opts || {};
    const body = typeof input === 'string' ? { prompt: input } : { messages: input };
    body.tier = opts.modelTier || 'default';
    if (opts.images) body.images = await Promise.all(Array.from(opts.images).slice(0, 5).map(toBase64Image));
    let res;
    try { res = await fetch('/api/claude', { method: 'POST', headers: { 'content-type': 'application/json', 'x-kva-token': getToken() }, body: JSON.stringify(body), signal: opts.signal }); }
    catch (e) { throw { code: e && e.name === 'AbortError' ? 'cancelled' : 'upstream_error', message: String(e) }; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw { code: data.code || 'upstream_error', message: data.message || String(res.status) };
    if (opts.onText) opts.onText({ text: data.text, delta: data.text });
    return data;
  }
  function parseJSON(text) {
    try { return JSON.parse(text); } catch (e) { /* 다음 방법 */ }
    const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/); if (fence) { try { return JSON.parse(fence[1]); } catch (e) { /* 다음 */ } }
    const s = text.search(/[[{]/), e2 = Math.max(text.lastIndexOf('}'), text.lastIndexOf(']'));
    if (s >= 0 && e2 > s) { try { return JSON.parse(text.slice(s, e2 + 1)); } catch (e) { /* 실패 */ } }
    throw { code: 'invalid_json', message: '응답에서 JSON을 찾지 못했습니다.', text: text };
  }
  if (!inArtifact) {
    const sample = async (input, opts) => { const d = await call(input, opts); return { text: d.text, truncated: !!d.truncated, modelTierApplied: (opts && opts.modelTier) || 'default' }; };
    sample.json = async (input, opts) => parseJSON((await call(typeof input === 'string' ? input + '\n\nJSON 값 하나로만 답하라.' : input, opts)).text);
    sample.limits = async () => ({ maxPromptBytes: 262144, images: { maxCount: 5, maxInputBytes: 20000000, mediaTypes: ['image/jpeg', 'image/png', 'image/webp', 'image/gif'] } });
    const downloads = { save: async ({ filename, data }) => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([data])); a.download = filename; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); } };
    window.claude = {
      use: async (name) => {
        const h = await ready;
        if (name === 'downloads') return downloads;
        if (name === 'sample') return h && h.claude ? sample : null;
        return null;
      },
    };
  }

  // ---------- 링크 미리보기 썸네일 ----------
  const PREVIEW_KEY = 'kvatlas.previews.v1';
  let memo = {};
  try { memo = JSON.parse(localStorage.getItem(PREVIEW_KEY) || '{}'); } catch (e) { memo = {}; }
  let saveTimer;
  const persist = () => { clearTimeout(saveTimer); saveTimer = setTimeout(() => { try { const keys = Object.keys(memo); if (keys.length > 3000) keys.slice(0, keys.length - 3000).forEach((k) => delete memo[k]); localStorage.setItem(PREVIEW_KEY, JSON.stringify(memo)); } catch (e) { /* 공간 부족 */ } }, 1000); };
  const queue = []; let active = 0;
  function pump() {
    while (active < 6 && queue.length) {
      const { url, resolve } = queue.shift(); active++;
      fetch('/api/preview?url=' + encodeURIComponent(url)).then((r) => (r.ok ? r.json() : {})).catch(() => ({}))
        .then((d) => { memo[url] = d.image || ''; persist(); resolve(memo[url]); })
        .finally(() => { active--; pump(); });
    }
  }
  const previewOf = (url) => (url in memo ? Promise.resolve(memo[url]) : new Promise((resolve) => { queue.push({ url, resolve }); pump(); }));
  function attach(th, src) {
    if (!src || th.querySelector('img.real')) return;
    const img = new Image(); img.className = 'real'; img.alt = ''; img.loading = 'lazy'; img.referrerPolicy = 'no-referrer';
    img.onerror = () => img.remove(); img.src = src; th.appendChild(img);
  }
  const io = 'IntersectionObserver' in window ? new IntersectionObserver((entries) => {
    entries.forEach((e) => { if (!e.isIntersecting) return; io.unobserve(e.target); const th = e.target; previewOf(th.dataset.preview).then((src) => attach(th, src)); });
  }, { rootMargin: '400px' }) : null;
  function scan(root) { if (!io) return; (root.querySelectorAll ? root : document).querySelectorAll('.th[data-preview]').forEach((th) => { if (!th.dataset.observed) { th.dataset.observed = '1'; io.observe(th); } }); }

  ready.then((h) => {
    if (!h) return;
    document.documentElement.classList.add('hosted');
    scan(document);
    new MutationObserver((muts) => muts.forEach((m) => m.addedNodes.forEach((n) => { if (n.nodeType === 1) scan(n.parentNode || n); }))).observe(document.body, { childList: true, subtree: true });
    if (window.KVA) window.KVA.updateCounts();
  });

  // ---------- 통합 검색: 이미지 결과 ----------
  const imgState = { q: '', page: 1, sources: null, results: [], status: {}, loading: false };
  function renderImages(box, variants) {
    const h = hosted.health;
    if (!h) {
      box.innerHTML = '<div class="img-note"><b>이미지 미리보기</b> — 검색 결과 썸네일은 KV 아틀라스를 직접 호스팅한 주소(Cloudflare Worker)에서 보입니다. claude.ai 화면은 보안 정책상 외부 이미지를 표시할 수 없어 아래 링크로 각 사이트를 엽니다.</div>';
      return;
    }
    if (!imgState.sources) imgState.sources = new Set(h.sources.map((s) => s.id));
    const esc = window.KVA.esc;
    box.innerHTML = '<div class="toolbar"><button type="button" class="btn" id="img-go">이미지로 보기</button>' +
      h.sources.map((s) => '<button type="button" class="chip" data-imgsrc="' + s.id + '" aria-pressed="' + imgState.sources.has(s.id) + '">' + esc(s.name) + '</button>').join('') +
      '<span class="meta-line" id="img-status"></span></div><div class="img-grid" id="img-grid"></div><p><button type="button" class="btn ghost" id="img-more" hidden>더 불러오기</button></p>';
    const grid = box.querySelector('#img-grid'), status = box.querySelector('#img-status'), more = box.querySelector('#img-more');
    const draw = () => {
      grid.innerHTML = imgState.results.map((r) => '<a class="img-card" href="' + esc(r.url) + '" target="_blank" rel="noopener" title="' + esc(r.title) + '"><img alt="' + esc(r.title) + '" loading="lazy" referrerpolicy="no-referrer" src="' + esc(r.thumb) + '" onerror="this.parentNode.remove()"><span>' + esc(r.source) + '</span></a>').join('');
      const failed = Object.entries(imgState.status).filter(([, v]) => !v.ok).map(([k]) => k);
      status.textContent = imgState.loading ? '불러오는 중...' : imgState.results.length ? imgState.results.length + '장' + (failed.length ? ' · 응답 없음: ' + failed.join(', ') : '') : '';
      more.hidden = !imgState.results.length || imgState.loading;
    };
    const load = async (page) => {
      const v = variants(); if (!v.length) { status.textContent = '검색어를 입력하세요'; return; }
      const q = v[1] || v[0];
      if (page === 1) { imgState.results = []; imgState.q = q; }
      imgState.page = page; imgState.loading = true; draw();
      try {
        const res = await fetch('/api/images?q=' + encodeURIComponent(q) + '&page=' + page + '&src=' + Array.from(imgState.sources).join(','));
        const d = await res.json();
        imgState.results = imgState.results.concat(d.results || []); imgState.status = d.sources || {};
      } catch (e) { status.textContent = '검색하지 못했습니다'; }
      imgState.loading = false; draw();
    };
    box.querySelector('#img-go').addEventListener('click', () => load(1));
    more.addEventListener('click', () => load(imgState.page + 1));
    box.querySelectorAll('[data-imgsrc]').forEach((b) => b.addEventListener('click', () => { const id = b.dataset.imgsrc; imgState.sources.has(id) ? imgState.sources.delete(id) : imgState.sources.add(id); b.setAttribute('aria-pressed', imgState.sources.has(id)); }));
    draw();
  }
  function hook() {
    if (!window.KVA) return setTimeout(hook, 50);
    window.KVA.on('search-view', ({ root, variants }) => { const box = root.querySelector('#img-results'); if (box) ready.then(() => renderImages(box, variants)); });
  }
  hook();
})();
