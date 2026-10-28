(function () {
  'use strict';
  const D = window.KV_DATA;
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];
  const shuffle = (arr) => { const a = arr.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };

  const TECH = Object.fromEntries(D.TECHNIQUES.map((t) => [t[0], { id: t[0], name: t[1], desc: t[2] }]));
  const MEDIA = Object.fromEntries(D.MEDIA);
  const CAT = Object.fromEntries(D.CATEGORIES.map((c) => [c[0], { id: c[0], name: c[1], desc: c[2] }]));
  const KVBY = Object.fromEntries(D.KV.map((k) => [k.id, k]));
  const SRCBY = Object.fromEntries(D.SOURCES.map((s) => [s.id, s]));
  const decade = (y) => Math.floor(y / 10) * 10;
  const decadeLabel = (d) => (d < 1950 ? '1950 이전' : d + '년대');
  const decadeKey = (y) => (y < 1950 ? 1940 : decade(y));

  // ---------- 저장소 ----------
  const STORE_KEY = 'kvatlas.board.v1';
  let board = { kv: [], site: [], custom: [] };
  try { const raw = localStorage.getItem(STORE_KEY); if (raw) board = Object.assign(board, JSON.parse(raw)); } catch (e) { /* 저장소 없음 */ }
  function saveBoard() { try { localStorage.setItem(STORE_KEY, JSON.stringify(board)); } catch (e) { toast('이 브라우저에서는 저장할 수 없습니다. 내보내기로 백업하세요.'); } updateCounts(); }
  function toggleSaved(kind, id) {
    const list = board[kind]; const i = list.indexOf(id);
    if (i >= 0) { list.splice(i, 1); toast('보드에서 뺐습니다'); } else { list.unshift(id); toast('보드에 담았습니다'); }
    saveBoard();
    return i < 0;
  }
  const isSaved = (kind, id) => board[kind].indexOf(id) >= 0;

  let toastTimer;
  function toast(msg) {
    let t = $('.toast'); if (!t) { t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
    t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 1800);
  }
  function copy(text, label) {
    const done = () => toast((label || '내용') + '을 복사했습니다');
    try { navigator.clipboard.writeText(text).then(done, () => fallbackCopy(text, done)); } catch (e) { fallbackCopy(text, done); }
  }
  function fallbackCopy(text, done) {
    const ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta); ta.select();
    let ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove(); ok ? done() : toast('복사가 막혀 있습니다. 직접 선택해 복사하세요.');
  }

  // ---------- 검색 URL ----------
  function siteScope(url) { try { const u = new URL(url); return (u.hostname.replace(/^www\./, '') + u.pathname).replace(/\/$/, ''); } catch (e) { return url; } }
  function searchUrl(src, q) {
    if (!src.search) return null;
    if (src.search === 'site') return 'https://www.google.com/search?tbm=isch&q=' + encodeURIComponent('site:' + siteScope(src.url) + ' ' + q);
    return src.search.replace('{q}', encodeURIComponent(q));
  }
  const byName = (n) => D.SOURCES.find((s) => s.name === n);
  const KV_ENGINES = ['Google 이미지', 'Pinterest', 'Behance', 'Yandex 이미지', 'Ads of the World', 'Bing 이미지'].map(byName).filter(Boolean);

  // ---------- 썸네일 ----------
  const THUMB_KINDS = ['type', 'negative', 'geometric', 'grid', 'silhouette', 'crop', 'color', 'data', 'logo', 'illust', 'photo', 'copy', 'handmade'];
  function thumbKind(k) { return k.tech.find((t) => THUMB_KINDS.indexOf(t) >= 0) || 'copy'; }
  function thumb(k) {
    const kind = thumbKind(k);
    let h = 0; for (const ch of k.id) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    const vars = k.col.map((c, i) => '--c' + i + ':' + c).join(';') + ';--rx:' + ((h % 5) - 2) * 9 + 'deg;--dx:' + (h % 4) * 9 + '%;--dy:' + ((h >> 3) % 4) * 6 + '%';
    const objs = { negative: '<i class="obj"></i>', geometric: '<i class="obj"></i><i class="obj2"></i>', grid: '<i class="obj"></i>', silhouette: '<i class="obj"></i>', photo: '<i class="obj"></i>', crop: '<i class="obj"></i>', color: '<i class="obj"></i>', illust: '<i class="obj"></i>', data: '<i class="obj"></i>', logo: '<i class="obj"></i>' }[kind] || '';
    const label = kind === 'negative' || kind === 'silhouette' ? k.brand : k.title;
    return '<div class="th th--' + kind + '" style="' + esc(vars) + '" role="img" aria-label="' + esc(k.title + ' 대표 색과 구성 도식') + '">' +
      objs + '<span class="yr">' + esc(k.year) + '</span><span class="ttl">' + esc(label) + '</span>' +
      '<span class="pal">' + k.col.map((c) => '<i style="background:' + esc(c) + '"></i>').join('') + '</span></div>';
  }
  function kvCard(k) {
    return '<button class="card" type="button" data-kv="' + esc(k.id) + '">' + thumb(k) +
      '<span class="cap"><b>' + esc(k.title) + (isSaved('kv', k.id) ? ' <span class="saved-mark" aria-label="보드에 있음">●</span>' : '') + '</b>' +
      '<span>' + esc(k.brand) + ' · ' + esc(k.by) + '</span></span></button>';
  }
  const dots = (n) => '<span class="dots" title="KV 활용도 ' + n + '/3">' + '■'.repeat(n) + '<span class="off">' + '■'.repeat(3 - n) + '</span></span>';

  // ---------- 상태 ----------
  const kvState = { q: '', ind: '', dec: '', med: new Set(), tech: new Set(), mood: new Set(), sort: 'year-asc', order: null };
  const srcState = { q: '', cat: new Set(), region: '', cost: '', searchable: false, sort: 'kv', group: true };
  const searchState = { q: '', terms: new Set(['en|key visual']), engines: new Set(KV_ENGINES.map((s) => s.id)) };
  const labState = { locks: {}, slots: {}, trio: null };

  function counts(list, fn) { const m = {}; list.forEach((x) => [].concat(fn(x)).forEach((v) => { m[v] = (m[v] || 0) + 1; })); return m; }

  // ---------- KV 보기 ----------
  function kvFiltered() {
    const q = kvState.q.trim().toLowerCase();
    let list = D.KV.filter((k) => {
      if (kvState.ind && k.ind !== kvState.ind) return false;
      if (kvState.dec && String(decadeKey(k.year)) !== kvState.dec) return false;
      for (const m of kvState.med) if (k.med.indexOf(m) < 0) return false;
      for (const t of kvState.tech) if (k.tech.indexOf(t) < 0) return false;
      for (const m of kvState.mood) if (k.mood.indexOf(m) < 0) return false;
      if (q) {
        const hay = [k.title, k.brand, k.by, k.ind, k.note, k.year, k.tech.map((t) => TECH[t].name).join(' '), k.mood.join(' ')].join(' ').toLowerCase();
        if (q.split(/\s+/).some((w) => hay.indexOf(w) < 0)) return false;
      }
      return true;
    });
    if (kvState.sort === 'shuffle') {
      if (!kvState.order) kvState.order = shuffle(D.KV.map((k) => k.id));
      list.sort((a, b) => kvState.order.indexOf(a.id) - kvState.order.indexOf(b.id));
    } else if (kvState.sort === 'year-desc') list.sort((a, b) => b.year - a.year);
    else if (kvState.sort === 'brand') list.sort((a, b) => a.brand.localeCompare(b.brand, 'ko'));
    else list.sort((a, b) => a.year - b.year);
    return list;
  }
  function chipGroup(name, items, set, cnt) {
    return items.map(([v, label]) => '<button type="button" class="chip" data-chip="' + name + '" data-v="' + esc(v) + '" aria-pressed="' + set.has(v) + '">' + esc(label) + ' <span class="n">' + (cnt[v] || 0) + '</span></button>').join('');
  }
  function renderKV(root) {
    const cTech = counts(D.KV, (k) => k.tech), cMed = counts(D.KV, (k) => k.med), cMood = counts(D.KV, (k) => k.mood), cInd = counts(D.KV, (k) => k.ind);
    const decs = Array.from(new Set(D.KV.map((k) => decadeKey(k.year)))).sort((a, b) => a - b);
    root.innerHTML =
      '<div class="view-head"><h2>키비주얼 데이터베이스</h2><p>디자인사와 광고사의 기준점이 된 키비주얼을 업종·매체·기법·무드·시대로 분류했습니다. 카드를 열면 아이디어 포인트와 원본 이미지 검색 링크가 나옵니다.</p></div>' +
      '<div class="split"><aside class="rail" aria-label="필터">' +
        '<div class="field"><label for="kv-q">검색</label><input id="kv-q" type="search" placeholder="브랜드, 제작자, 키워드" value="' + esc(kvState.q) + '"></div>' +
        '<div class="field"><label for="kv-ind">업종</label><select id="kv-ind"><option value="">전체</option>' + D.INDUSTRIES.filter((i) => cInd[i]).map((i) => '<option' + (kvState.ind === i ? ' selected' : '') + ' value="' + esc(i) + '">' + esc(i) + ' (' + cInd[i] + ')</option>').join('') + '</select></div>' +
        '<div class="field"><label for="kv-dec">시대</label><select id="kv-dec"><option value="">전체</option>' + decs.map((d) => '<option' + (kvState.dec === String(d) ? ' selected' : '') + ' value="' + d + '">' + decadeLabel(d) + '</option>').join('') + '</select></div>' +
        '<details open><summary>기법</summary><div class="opts">' + chipGroup('tech', D.TECHNIQUES.map((t) => [t[0], t[1]]), kvState.tech, cTech) + '</div></details>' +
        '<details open><summary>매체</summary><div class="opts">' + chipGroup('med', D.MEDIA, kvState.med, cMed) + '</div></details>' +
        '<details><summary>무드</summary><div class="opts">' + chipGroup('mood', D.MOODS.map((m) => [m, m]), kvState.mood, cMood) + '</div></details>' +
      '</aside><section style="min-width:0">' +
        '<div class="toolbar"><span class="meta-line grow" id="kv-count"></span>' +
          '<select id="kv-sort" aria-label="정렬" style="width:auto"><option value="year-asc">연도 오름차순</option><option value="year-desc">연도 내림차순</option><option value="brand">브랜드 가나다</option><option value="shuffle">무작위 배열</option></select>' +
          '<button class="btn ghost sm" type="button" id="kv-reset">필터 초기화</button>' +
          '<button class="btn sm" type="button" id="kv-rand">필터 안에서 하나 뽑기</button></div>' +
        '<div class="sheet" id="kv-grid"></div></section></div>';
    $('#kv-sort', root).value = kvState.sort;
    const draw = () => {
      const list = kvFiltered();
      $('#kv-count', root).textContent = list.length + ' / ' + D.KV.length + '개';
      $('#kv-grid', root).innerHTML = list.length ? list.map(kvCard).join('') : '<p class="empty">조건에 맞는 키비주얼이 없습니다. 기법이나 매체 필터를 하나 풀어 보세요.</p>';
    };
    draw();
    $('#kv-q', root).addEventListener('input', (e) => { kvState.q = e.target.value; draw(); });
    $('#kv-ind', root).addEventListener('change', (e) => { kvState.ind = e.target.value; draw(); });
    $('#kv-dec', root).addEventListener('change', (e) => { kvState.dec = e.target.value; draw(); });
    $('#kv-sort', root).addEventListener('change', (e) => { kvState.sort = e.target.value; kvState.order = null; draw(); });
    $$('[data-chip]', root).forEach((b) => b.addEventListener('click', () => {
      const set = kvState[b.dataset.chip]; const v = b.dataset.v;
      set.has(v) ? set.delete(v) : set.add(v); b.setAttribute('aria-pressed', set.has(v)); draw();
    }));
    $('#kv-reset', root).addEventListener('click', () => { Object.assign(kvState, { q: '', ind: '', dec: '' }); kvState.med.clear(); kvState.tech.clear(); kvState.mood.clear(); render(); });
    $('#kv-rand', root).addEventListener('click', () => { const l = kvFiltered(); if (l.length) openKV(pick(l).id); });
  }

  // ---------- KV 상세 ----------
  function openKV(id) {
    const k = KVBY[id]; if (!k) return;
    closeModal();
    const similar = D.KV.filter((o) => o.id !== k.id).map((o) => ({ o, s: o.tech.filter((t) => k.tech.indexOf(t) >= 0).length * 2 + o.mood.filter((m) => k.mood.indexOf(m) >= 0).length + (o.ind === k.ind ? 1 : 0) }))
      .filter((x) => x.s > 1).sort((a, b) => b.s - a.s).slice(0, 4).map((x) => x.o);
    const queries = [k.q, k.q + ' key visual'];
    const m = document.createElement('div');
    m.className = 'modal'; m.setAttribute('role', 'dialog'); m.setAttribute('aria-modal', 'true'); m.setAttribute('aria-label', k.title);
    m.innerHTML = '<div class="modal-box"><div>' + thumb(k) + '<p class="th-note">대표 색과 구성을 도식화한 썸네일입니다. 원본은 아래 검색 링크로 출처 사이트에서 확인하세요.</p></div><div>' +
      '<button type="button" class="btn ghost sm close" data-close>닫기</button>' +
      '<h2>' + esc(k.title) + '</h2><p class="sub">' + esc(k.brand) + ' · ' + esc(k.year) + ' · ' + esc(k.by) + '</p>' +
      '<p class="note">' + esc(k.note) + '</p>' +
      '<dl><dt>업종</dt><dd><span class="tag">' + esc(k.ind) + '</span></dd>' +
      '<dt>매체</dt><dd>' + k.med.map((x) => '<span class="tag">' + esc(MEDIA[x]) + '</span>').join('') + '</dd>' +
      '<dt>무드</dt><dd>' + k.mood.map((x) => '<span class="tag">' + esc(x) + '</span>').join('') + '</dd>' +
      '<dt>색 인상</dt><dd class="hexes">' + k.col.map((c) => '<button type="button" class="hex" data-hex="' + esc(c) + '"><i style="background:' + esc(c) + '"></i>' + esc(c) + '</button>').join('') + '</dd></dl>' +
      '<p class="lbl">쓰인 기법</p><ul class="techlist">' + k.tech.map((t) => '<li><b>' + esc(TECH[t].name) + '</b> — ' + esc(TECH[t].desc) + '</li>').join('') + '</ul>' +
      '<p class="lbl">원본 이미지 찾기</p><div class="toolbar">' + KV_ENGINES.map((s) => '<a class="btn ghost sm" target="_blank" rel="noopener" href="' + esc(searchUrl(s, queries[0])) + '">' + esc(s.name) + '</a>').join('') + '</div>' +
      '<div class="toolbar"><button type="button" class="btn sm" data-save>' + (isSaved('kv', k.id) ? '보드에서 빼기' : '보드에 담기') + '</button>' +
      '<button type="button" class="btn ghost sm" data-tosearch>통합 검색으로 보내기</button>' +
      '<button type="button" class="btn ghost sm" data-copyq>검색어 복사</button></div>' +
      (similar.length ? '<p class="lbl">비슷한 결의 레퍼런스</p><div class="mini-row">' + similar.map(kvCard).join('') + '</div>' : '') +
      '</div></div>';
    document.body.appendChild(m);
    const prev = document.activeElement;
    m._prev = prev;
    $('[data-close]', m).focus();
    m.addEventListener('click', (e) => {
      if (e.target === m || e.target.closest('[data-close]')) return closeModal();
      const c = e.target.closest('[data-kv]'); if (c) return openKV(c.dataset.kv);
      const h = e.target.closest('[data-hex]'); if (h) return copy(h.dataset.hex, '색상 코드');
      if (e.target.closest('[data-save]')) { const on = toggleSaved('kv', k.id); e.target.textContent = on ? '보드에서 빼기' : '보드에 담기'; refreshCards(); }
      if (e.target.closest('[data-copyq]')) copy(k.q, '검색어');
      if (e.target.closest('[data-tosearch]')) { searchState.q = k.q; closeModal(); location.hash = 'search'; }
    });
  }
  function closeModal() { const m = $('.modal'); if (m) { const p = m._prev; m.remove(); if (p && p.focus) p.focus(); } }
  function refreshCards() { $$('#app .card[data-kv]').forEach((c) => { c.outerHTML = kvCard(KVBY[c.dataset.kv]); }); }
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });

  // ---------- 사이트 디렉터리 ----------
  function srcFiltered() {
    const q = srcState.q.trim().toLowerCase();
    let list = D.SOURCES.filter((s) => {
      if (srcState.cat.size && !srcState.cat.has(s.cat)) return false;
      if (srcState.region && s.region !== srcState.region) return false;
      if (srcState.cost && s.cost !== srcState.cost) return false;
      if (srcState.searchable && !s.search) return false;
      if (q && (s.name + ' ' + s.desc + ' ' + s.url + ' ' + s.region + ' ' + CAT[s.cat].name).toLowerCase().indexOf(q) < 0) return false;
      return true;
    });
    if (srcState.sort === 'kv') list.sort((a, b) => b.kv - a.kv || a.name.localeCompare(b.name, 'ko'));
    else list.sort((a, b) => a.name.localeCompare(b.name, 'ko'));
    return list;
  }
  function srcRow(s) {
    return '<article class="src"><div style="min-width:0"><h3><a href="' + esc(s.url) + '" target="_blank" rel="noopener">' + esc(s.name) + '</a><span class="host">' + esc(siteScope(s.url)) + '</span></h3>' +
      '<p>' + esc(s.desc) + '</p><div class="tags">' + dots(s.kv) + ' <span class="tag">' + esc(CAT[s.cat].name) + '</span><span class="tag">' + esc(s.region) + '</span><span class="tag">' + esc(s.cost) + '</span>' +
      (s.search ? '<span class="tag">' + (s.search === 'site' ? '구글 사이트 검색' : '직접 검색') + '</span>' : '') + '</div></div>' +
      '<div class="acts"><a class="btn ghost sm" href="' + esc(s.url) + '" target="_blank" rel="noopener">열기</a>' +
      '<button type="button" class="btn ghost sm" data-savesite="' + s.id + '" aria-pressed="' + isSaved('site', s.id) + '">' + (isSaved('site', s.id) ? '담음' : '담기') + '</button></div></article>';
  }
  function renderSources(root) {
    const regions = Array.from(new Set(D.SOURCES.map((s) => s.region))).sort((a, b) => a.localeCompare(b, 'ko'));
    const cCat = counts(D.SOURCES, (s) => s.cat);
    root.innerHTML =
      '<div class="view-head"><h2>이미지·포트폴리오 사이트 디렉터리</h2><p>' + D.SOURCES.length + '개 사이트를 ' + D.CATEGORIES.length + '개 분류로 정리했습니다. 네모 세 칸은 광고 KV 레퍼런스로서의 활용도입니다.</p></div>' +
      '<div class="split"><aside class="rail" aria-label="필터">' +
        '<div class="field"><label for="src-q">검색</label><input id="src-q" type="search" placeholder="사이트 이름, 설명" value="' + esc(srcState.q) + '"></div>' +
        '<details open><summary>분류</summary><div class="opts">' + D.CATEGORIES.map((c) => '<button type="button" class="chip" data-cat="' + c[0] + '" aria-pressed="' + srcState.cat.has(c[0]) + '">' + esc(c[1]) + ' <span class="n">' + cCat[c[0]] + '</span></button>').join('') + '</div></details>' +
        '<div class="field"><label for="src-region">지역</label><select id="src-region"><option value="">전체</option>' + regions.map((r) => '<option' + (srcState.region === r ? ' selected' : '') + '>' + esc(r) + '</option>').join('') + '</select></div>' +
        '<div class="field"><label for="src-cost">비용</label><select id="src-cost"><option value="">전체</option>' + ['무료', '부분유료', '유료'].map((r) => '<option' + (srcState.cost === r ? ' selected' : '') + '>' + r + '</option>').join('') + '</select></div>' +
        '<label style="display:flex;gap:8px;align-items:center;font-size:var(--t-sm)"><input type="checkbox" id="src-searchable"' + (srcState.searchable ? ' checked' : '') + '> 검색어로 바로 검색되는 곳만</label>' +
      '</aside><section style="min-width:0">' +
        '<div class="toolbar"><span class="meta-line grow" id="src-count"></span>' +
          '<select id="src-sort" aria-label="정렬" style="width:auto"><option value="kv">KV 활용도순</option><option value="name">이름순</option></select>' +
          '<button type="button" class="btn ghost sm" id="src-group" aria-pressed="' + srcState.group + '">분류별 묶기</button>' +
          '<button type="button" class="btn sm" id="src-rand">사이트 하나 뽑기</button></div>' +
        '<div id="src-list"></div></section></div>';
    $('#src-sort', root).value = srcState.sort;
    const draw = () => {
      const list = srcFiltered();
      $('#src-count', root).textContent = list.length + ' / ' + D.SOURCES.length + '곳';
      const out = $('#src-list', root);
      if (!list.length) { out.innerHTML = '<p class="empty">조건에 맞는 사이트가 없습니다.</p>'; return; }
      if (srcState.group) {
        out.innerHTML = D.CATEGORIES.map((c) => {
          const items = list.filter((s) => s.cat === c[0]); if (!items.length) return '';
          return '<h3 class="cat-head">' + esc(c[1]) + ' <small>' + items.length + '</small></h3><p class="cat-desc">' + esc(c[2]) + '</p><div class="src-list">' + items.map(srcRow).join('') + '</div>';
        }).join('');
      } else out.innerHTML = '<div class="src-list">' + list.map(srcRow).join('') + '</div>';
    };
    draw();
    $('#src-q', root).addEventListener('input', (e) => { srcState.q = e.target.value; draw(); });
    $('#src-region', root).addEventListener('change', (e) => { srcState.region = e.target.value; draw(); });
    $('#src-cost', root).addEventListener('change', (e) => { srcState.cost = e.target.value; draw(); });
    $('#src-sort', root).addEventListener('change', (e) => { srcState.sort = e.target.value; draw(); });
    $('#src-searchable', root).addEventListener('change', (e) => { srcState.searchable = e.target.checked; draw(); });
    $('#src-group', root).addEventListener('click', (e) => { srcState.group = !srcState.group; e.target.setAttribute('aria-pressed', srcState.group); draw(); });
    $$('[data-cat]', root).forEach((b) => b.addEventListener('click', () => { const v = b.dataset.cat; srcState.cat.has(v) ? srcState.cat.delete(v) : srcState.cat.add(v); b.setAttribute('aria-pressed', srcState.cat.has(v)); draw(); }));
    $('#src-rand', root).addEventListener('click', () => {
      const l = srcFiltered(); if (!l.length) return;
      const weighted = [].concat.apply([], l.map((s) => Array(s.kv * s.kv).fill(s)));
      const s = pick(weighted);
      srcState.group = false; $('#src-group', root).setAttribute('aria-pressed', 'false');
      $('#src-list', root).innerHTML = '<p class="meta-line">무작위로 뽑은 사이트 (활용도가 높을수록 잘 뽑힙니다)</p><div class="src-list">' + srcRow(s) + '</div><p style="margin-top:12px"><button type="button" class="btn ghost sm" id="src-back">전체 목록으로</button></p>';
      $('#src-back', root).addEventListener('click', draw);
    });
    root.addEventListener('click', (e) => {
      const b = e.target.closest('[data-savesite]'); if (!b) return;
      const on = toggleSaved('site', b.dataset.savesite); b.setAttribute('aria-pressed', on); b.textContent = on ? '담음' : '담기';
    });
  }

  // ---------- 통합 검색 ----------
  const PRESETS = [
    ['핵심', () => KV_ENGINES.concat([byName('Savee')]).filter(Boolean)],
    ['광고 아카이브', () => D.SOURCES.filter((s) => (s.cat === 'ad-archive' || s.cat === 'ad-library') && s.search)],
    ['아시아', () => ['네이버 이미지', 'Baidu 이미지', 'Yahoo! JAPAN 이미지', 'ZCOOL 站酷', '花瓣 Huaban', '노트폴리오', 'TVCF', 'AdverTimes 宣伝会議', '数英 Digitaling'].map(byName).filter(Boolean)],
    ['큐레이션', () => D.SOURCES.filter((s) => s.cat === 'curation' && s.search)],
    ['고전·아카이브', () => D.SOURCES.filter((s) => (s.cat === 'archive' || s.cat === 'type') && s.search)],
    ['전체', () => D.SOURCES.filter((s) => s.search)],
  ];
  function queryVariants() {
    const base = searchState.q.trim(); if (!base) return [];
    const v = [base];
    D.TERMS.forEach(([lang, term]) => { if (searchState.terms.has(lang + '|' + term)) v.push(base + ' ' + term); });
    return v;
  }
  function renderSearch(root) {
    const searchable = D.SOURCES.filter((s) => s.search);
    root.innerHTML =
      '<div class="view-head"><h2>통합 검색 런처</h2><p>검색어 하나를 여러 이미지 사이트에 동시에 겁니다. 검색창이 없는 사이트는 구글 이미지의 site: 검색으로 대신합니다. 지역마다 KV를 부르는 말이 달라서, 나라별 키워드를 붙이면 결과가 크게 달라집니다.</p></div>' +
      '<div class="launch-q"><input id="sq" type="search" placeholder="예: whisky poster, 화장품 런칭, minimal car ad" value="' + esc(searchState.q) + '" aria-label="검색어"><button class="btn" type="button" id="sq-copy">링크 전체 복사</button></div>' +
      '<p class="lbl">붙일 키워드</p><div class="toolbar">' + D.TERMS.map(([lang, term]) => { const key = lang + '|' + term; return '<button type="button" class="chip" data-term="' + esc(key) + '" aria-pressed="' + searchState.terms.has(key) + '"><span class="n">' + lang + '</span> ' + esc(term) + '</button>'; }).join('') + '</div>' +
      '<p class="lbl">검색할 곳 <span class="meta-line" id="eng-count"></span></p><div class="toolbar">' + PRESETS.map((p, i) => '<button type="button" class="btn ghost sm" data-preset="' + i + '">' + esc(p[0]) + '</button>').join('') + '<button type="button" class="btn ghost sm" data-preset="none">모두 해제</button></div>' +
      '<div class="engines" id="engines">' + searchable.map((s) => '<div class="eng" data-eng="' + s.id + '"><input type="checkbox" id="eng-' + s.id + '" aria-label="' + esc(s.name) + '"><label class="nm" for="eng-' + s.id + '">' + esc(s.name) + ' <span class="meta-line">' + esc(s.region) + '</span></label><div class="links"></div></div>').join('') + '</div>';
    const draw = () => {
      const vars = queryVariants();
      let n = 0;
      $$('.eng', root).forEach((row) => {
        const s = SRCBY[row.dataset.eng]; const on = searchState.engines.has(s.id); if (on) n++;
        $('input', row).checked = on; row.classList.toggle('off', !on);
        $('.links', row).innerHTML = on ? (vars.length ? vars.map((q) => '<a target="_blank" rel="noopener" href="' + esc(searchUrl(s, q)) + '" title="' + esc(q) + '">' + esc(q) + '</a>').join('') : '<span class="meta-line">검색어를 입력하세요</span>') : '';
      });
      $('#eng-count', root).textContent = '(' + n + '곳 선택)';
    };
    draw();
    $('#sq', root).addEventListener('input', (e) => { searchState.q = e.target.value; draw(); });
    $$('[data-term]', root).forEach((b) => b.addEventListener('click', () => { const k = b.dataset.term; searchState.terms.has(k) ? searchState.terms.delete(k) : searchState.terms.add(k); b.setAttribute('aria-pressed', searchState.terms.has(k)); draw(); }));
    $$('[data-preset]', root).forEach((b) => b.addEventListener('click', () => {
      searchState.engines = new Set(b.dataset.preset === 'none' ? [] : PRESETS[+b.dataset.preset][1]().map((s) => s.id)); draw();
    }));
    $('#engines', root).addEventListener('change', (e) => { const row = e.target.closest('.eng'); if (!row) return; e.target.checked ? searchState.engines.add(row.dataset.eng) : searchState.engines.delete(row.dataset.eng); draw(); });
    $('#sq-copy', root).addEventListener('click', () => {
      const vars = queryVariants(); if (!vars.length) return toast('검색어를 먼저 입력하세요');
      const lines = []; searchState.engines.forEach((id) => vars.forEach((q) => lines.push(SRCBY[id].name + ' | ' + q + ' | ' + searchUrl(SRCBY[id], q))));
      copy(lines.join('\n'), lines.length + '개 링크');
    });
    if (!searchState.q) $('#sq', root).focus();
  }

  // ---------- 무작위·훈련 ----------
  const CONSTRAINTS = ['헤드라인은 한 줄, 다섯 단어 이하', '제품은 화면의 10% 이하로', '색은 두 가지만', '사진 없이 타이포그래피만', '로고는 일부만 보이게', '사람 얼굴 클로즈업을 반드시 넣는다', '흑백에 포인트 색 하나', '3단 그리드를 엄격히 지킨다', '손으로 그린 요소 하나를 넣는다', '같은 틀로 3장 시리즈를 만든다', '여백을 화면의 60% 이상 남긴다', '1:1과 9:16 두 버전을 동시에 설계한다', '카피 없이 이미지만으로 이해되게', '제품을 다른 사물로 바꿔 보여준다'];
  const TARGETS = ['20대 초반 대학생', '30대 맞벌이 부부', '40대 직장인 남성', '첫 차를 사려는 사회초년생', '반려동물과 사는 1인 가구', '은퇴를 앞둔 50대', '10대 게이머', '해외 여행객', '스타트업 창업자', '육아 중인 부모'];
  const TIMES = ['15분 썸네일 스케치 6장', '30분 러프 시안 3장', '1시간 완성 시안 1장', '10분 카피만 20개'];
  function randomPalette() { const k = pick(D.KV); return { colors: k.col, from: k }; }
  function slotGen(key) {
    switch (key) {
      case 'ind': { const v = pick(D.INDUSTRIES.filter((i) => i !== '브랜드 아이덴티티')); return { v: v, d: '업종' }; }
      case 'target': return { v: pick(TARGETS), d: '타깃' };
      case 'mood': return { v: pick(D.MOODS), d: '브랜드 톤' };
      case 'tech': { const t = pick(D.TECHNIQUES); return { v: t[1], d: t[2] }; }
      case 'spec': { const s = pick(D.SPECS); return { v: s[1] + ' ' + s[2] + '×' + s[3] + s[4], d: s[0] + (s[5] ? ' · ' + s[5] : '') }; }
      case 'palette': { const p = randomPalette(); return { v: p.colors.join(' '), d: p.from.title + ' (' + p.from.brand + ')의 색 인상', colors: p.colors }; }
      case 'rule': return { v: pick(CONSTRAINTS), d: '제약 조건' };
      case 'ref': { const k = pick(D.KV); return { v: k.title + ' — ' + k.brand, d: '출발점 레퍼런스', kv: k.id }; }
      case 'time': return { v: pick(TIMES), d: '시간 제한' };
    }
  }
  const SLOT_LABELS = [['ind', '업종'], ['target', '타깃'], ['mood', '톤'], ['tech', '핵심 기법'], ['spec', '매체 규격'], ['palette', '팔레트'], ['rule', '제약'], ['ref', '레퍼런스'], ['time', '시간']];
  function rollBrief(all) { SLOT_LABELS.forEach(([k]) => { if (all || !labState.locks[k] || !labState.slots[k]) labState.slots[k] = slotGen(k); }); }
  function rollTrio() {
    const k = pick(D.KV);
    const pool = [].concat.apply([], D.SOURCES.filter((s) => s.kv >= 2).map((s) => Array(s.kv).fill(s)));
    const s = pick(pool); const t = pick(D.TECHNIQUES);
    labState.trio = { kv: k.id, src: s.id, tech: t[0] };
  }
  function briefText() {
    return 'KV 연습 브리프\n' + SLOT_LABELS.map(([k, l]) => '- ' + l + ': ' + labState.slots[k].v + (k === 'tech' || k === 'ref' || k === 'palette' ? ' (' + labState.slots[k].d + ')' : '')).join('\n');
  }
  function renderLab(root) {
    if (!labState.trio) rollTrio();
    if (!labState.slots.ind) rollBrief(true);
    const k = KVBY[labState.trio.kv], s = SRCBY[labState.trio.src], t = TECH[labState.trio.tech];
    const techKVs = D.KV.filter((x) => x.tech.indexOf(t.id) >= 0);
    root.innerHTML =
      '<div class="view-head"><h2>무작위 추천과 감도 훈련</h2><p>매일 하나의 레퍼런스, 하나의 사이트, 하나의 기법을 뽑아 봅니다. 아래 브리프 생성기는 조건을 무작위로 조합해 연습 과제를 만듭니다. 마음에 드는 칸은 고정하고 나머지만 다시 섞을 수 있습니다.</p></div>' +
      '<div class="toolbar"><button type="button" class="btn" id="trio-roll">세 가지 다시 뽑기</button><span class="meta-line">추천 루틴: 레퍼런스를 3분 관찰 → 왜 작동하는지 한 문장 → 같은 기법으로 썸네일 3장</span></div>' +
      '<div class="lab-grid">' +
        '<section class="panel"><h3>오늘의 키비주얼</h3><div style="max-width:220px">' + kvCard(k) + '</div><p style="font-size:var(--t-sm)">' + esc(k.note) + '</p></section>' +
        '<section class="panel"><h3>오늘 둘러볼 사이트</h3><p class="big"><a href="' + esc(s.url) + '" target="_blank" rel="noopener">' + esc(s.name) + '</a></p><p class="meta-line">' + esc(CAT[s.cat].name) + ' · ' + esc(s.region) + ' · ' + dots(s.kv) + '</p><p style="font-size:var(--t-sm)">' + esc(s.desc) + '</p>' +
          (s.search ? '<p class="toolbar"><a class="btn ghost sm" target="_blank" rel="noopener" href="' + esc(searchUrl(s, k.q)) + '">여기서 「' + esc(k.title) + '」 찾기</a></p>' : '') + '</section>' +
        '<section class="panel"><h3>오늘의 기법</h3><p class="big">' + esc(t.name) + '</p><p style="font-size:var(--t-sm)">' + esc(t.desc) + '</p><p class="meta-line">이 기법을 쓴 레퍼런스 ' + techKVs.length + '개</p><div class="mini-row">' + techKVs.slice(0, 4).map(kvCard).join('') + '</div></section>' +
      '</div>' +
      '<section class="brief" aria-label="브리프 생성기"><div class="brief-head"><h3>연습 브리프 생성기</h3><div class="toolbar" style="margin:0"><button type="button" class="btn sm" id="brief-roll">고정 안 한 칸 섞기</button><button type="button" class="btn ghost sm" id="brief-copy">브리프 복사</button></div></div>' +
      '<div class="slots">' + SLOT_LABELS.map(([key, label]) => {
        const v = labState.slots[key];
        return '<div class="slot"><div class="k"><span class="lbl">' + esc(label) + '</span><button type="button" class="lock" data-lock="' + key + '" aria-pressed="' + !!labState.locks[key] + '">' + (labState.locks[key] ? '고정됨' : '고정') + '</button></div>' +
          (v.colors ? '<div class="swatches">' + v.colors.map((c) => '<i style="background:' + esc(c) + '"></i>').join('') + '</div>' : '') +
          '<span class="v">' + (v.kv ? '<a href="#lab" data-kv="' + esc(v.kv) + '">' + esc(v.v) + '</a>' : esc(v.v)) + '</span><span class="d">' + esc(v.d) + '</span></div>';
      }).join('') + '</div></section>';
    $('#trio-roll', root).addEventListener('click', () => { rollTrio(); renderLab(root); });
    $('#brief-roll', root).addEventListener('click', () => { rollBrief(false); renderLab(root); });
    $('#brief-copy', root).addEventListener('click', () => copy(briefText(), '브리프'));
    $$('[data-lock]', root).forEach((b) => b.addEventListener('click', () => { const k2 = b.dataset.lock; labState.locks[k2] = !labState.locks[k2]; renderLab(root); }));
  }

  // ---------- 내 보드 ----------
  let pendingImg = '';
  function resizeImage(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => {
        const img = new Image();
        img.onload = () => {
          const max = 640, sc = Math.min(1, max / Math.max(img.width, img.height));
          const c = document.createElement('canvas'); c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
          c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
          resolve(c.toDataURL('image/jpeg', 0.8));
        };
        img.onerror = reject; img.src = r.result;
      };
      r.onerror = reject; r.readAsDataURL(file);
    });
  }
  function customCard(c) {
    const img = c.img ? '<img alt="" src="' + esc(c.img) + '" onerror="this.replaceWith(document.createTextNode(\'이미지를 불러올 수 없음\'))">' : esc(c.title);
    return '<div class="card" style="cursor:default"><div class="custom-th">' + img + '</div><span class="cap"><b>' + (c.url ? '<a href="' + esc(c.url) + '" target="_blank" rel="noopener">' + esc(c.title) + '</a>' : esc(c.title)) + '</b>' +
      '<span>' + esc((c.tags || []).join(', ')) + '</span></span>' + (c.note ? '<span class="th-note">' + esc(c.note) + '</span>' : '') +
      '<span><button type="button" class="btn ghost sm" data-delcustom="' + esc(c.id) + '">삭제</button></span></div>';
  }
  function renderBoard(root) {
    const tags = Array.from(new Set([].concat.apply([], board.custom.map((c) => c.tags || [])))).sort();
    const sites = board.site.map((id) => SRCBY[id]).filter(Boolean);
    const kvs = board.kv.map((id) => KVBY[id]).filter(Boolean);
    const total = sites.length + kvs.length + board.custom.length;
    root.innerHTML =
      '<div class="view-head"><h2>내 보드</h2><p>담아 둔 키비주얼과 사이트, 직접 모은 레퍼런스를 한곳에 둡니다. 이미지 파일을 끌어 놓으면 작게 줄여서 이 브라우저에 저장합니다.</p></div>' +
      '<form class="board-form" id="bf" autocomplete="off">' +
        '<div class="field"><label for="bf-title">제목</label><input id="bf-title" type="text" required placeholder="예: 위스키 브랜드 다크 톤 KV"></div>' +
        '<div class="field"><label for="bf-url">출처 링크</label><input id="bf-url" type="url" placeholder="https://www.behance.net/gallery/..."></div>' +
        '<div class="field"><label for="bf-tags">태그 (쉼표로 구분)</label><input id="bf-tags" type="text" placeholder="주류, 다크, 제품 히어로"></div>' +
        '<div class="field"><label for="bf-note">메모</label><input id="bf-note" type="text" placeholder="왜 좋은지 한 문장"></div>' +
        '<div class="field full"><span class="lbl">이미지</span><label class="drop" id="bf-drop" for="bf-file">' + (pendingImg ? '<img alt="" src="' + pendingImg + '">' : '') + '이미지 파일을 끌어 놓거나 눌러서 고르세요 (640px로 줄여 저장)</label><input id="bf-file" type="file" accept="image/*" hidden></div>' +
        '<div class="full toolbar" style="margin:0"><button class="btn" type="submit">보드에 추가</button><span class="meta-line">총 ' + total + '개 저장됨</span></div>' +
      '</form>' +
      (total === 0 ? '<p class="empty">아직 담은 것이 없습니다. 키비주얼 카드를 열어 "보드에 담기"를 누르거나, 위 양식으로 직접 모은 레퍼런스를 추가하세요.</p>' : '') +
      (board.custom.length ? '<h3 class="cat-head">직접 모은 레퍼런스 <small>' + board.custom.length + '</small></h3>' +
        (tags.length ? '<div class="toolbar"><button type="button" class="chip" data-tagf="" aria-pressed="true">전체</button>' + tags.map((t) => '<button type="button" class="chip" data-tagf="' + esc(t) + '" aria-pressed="false">' + esc(t) + '</button>').join('') + '</div>' : '') +
        '<div class="sheet" id="custom-grid">' + board.custom.map(customCard).join('') + '</div>' : '') +
      (kvs.length ? '<h3 class="cat-head">담은 키비주얼 <small>' + kvs.length + '</small></h3><div class="sheet">' + kvs.map(kvCard).join('') + '</div>' : '') +
      (sites.length ? '<h3 class="cat-head">담은 사이트 <small>' + sites.length + '</small></h3><div class="src-list">' + sites.map(srcRow).join('') + '</div>' : '') +
      '<h3 class="cat-head">백업</h3><p class="cat-desc">아래 내용을 복사해 두었다가 다른 기기에서 붙여 넣고 "가져오기"를 누르면 보드가 합쳐집니다.</p>' +
      '<div class="field"><label for="bf-json">보드 데이터 (JSON)</label><textarea id="bf-json" spellcheck="false"></textarea></div>' +
      '<div class="toolbar"><button type="button" class="btn ghost sm" id="bf-export">현재 보드 채우기</button><button type="button" class="btn ghost sm" id="bf-copy">복사</button><button type="button" class="btn sm" id="bf-import">가져오기</button></div>';

    const drop = $('#bf-drop', root), file = $('#bf-file', root);
    const takeFile = (f) => { if (!f || !/^image\//.test(f.type)) return toast('이미지 파일만 넣을 수 있습니다'); resizeImage(f).then((u) => { pendingImg = u; drop.innerHTML = '<img alt="" src="' + u + '">이미지를 바꾸려면 다시 고르세요'; }, () => toast('이미지를 읽지 못했습니다')); };
    file.addEventListener('change', () => takeFile(file.files[0]));
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => takeFile(e.dataTransfer.files[0]));
    $('#bf', root).addEventListener('submit', (e) => {
      e.preventDefault();
      const title = $('#bf-title', root).value.trim(); if (!title) return;
      board.custom.unshift({ id: 'c' + Date.now().toString(36), title: title, url: $('#bf-url', root).value.trim(), tags: $('#bf-tags', root).value.split(',').map((x) => x.trim()).filter(Boolean), note: $('#bf-note', root).value.trim(), img: pendingImg, created: new Date().toISOString().slice(0, 10) });
      pendingImg = ''; saveBoard(); toast('보드에 추가했습니다'); render();
    });
    root.addEventListener('click', (e) => {
      const d = e.target.closest('[data-delcustom]');
      if (d) {
        if (d.dataset.confirm) { board.custom = board.custom.filter((c) => c.id !== d.dataset.delcustom); saveBoard(); render(); }
        else { d.dataset.confirm = '1'; d.textContent = '한 번 더 누르면 삭제'; }
        return;
      }
      const sv = e.target.closest('[data-savesite]');
      if (sv) { toggleSaved('site', sv.dataset.savesite); render(); return; }
      const tf = e.target.closest('[data-tagf]');
      if (tf) {
        $$('[data-tagf]', root).forEach((b) => b.setAttribute('aria-pressed', b === tf));
        const t = tf.dataset.tagf;
        $('#custom-grid', root).innerHTML = board.custom.filter((c) => !t || (c.tags || []).indexOf(t) >= 0).map(customCard).join('');
      }
    });
    $('#bf-export', root).addEventListener('click', () => { $('#bf-json', root).value = JSON.stringify(board, null, 1); });
    $('#bf-copy', root).addEventListener('click', () => { const ta = $('#bf-json', root); if (!ta.value) ta.value = JSON.stringify(board, null, 1); copy(ta.value, '보드 데이터'); });
    $('#bf-import', root).addEventListener('click', () => {
      let data; try { data = JSON.parse($('#bf-json', root).value); } catch (err) { return toast('JSON 형식이 올바르지 않습니다'); }
      ['kv', 'site'].forEach((k) => (data[k] || []).forEach((id) => { if (board[k].indexOf(id) < 0) board[k].push(id); }));
      const have = new Set(board.custom.map((c) => c.id));
      (data.custom || []).forEach((c) => { if (c && c.id && c.title && !have.has(c.id)) board.custom.push(c); });
      saveBoard(); toast('보드를 합쳤습니다'); render();
    });
  }

  // ---------- 규격 ----------
  function renderSpecs(root) {
    let lastGroup = '';
    const rows = D.SPECS.map((s) => {
      const [g, name, w, h, unit, memo] = s;
      const maxSide = 44, sc = maxSide / Math.max(w, h);
      const box = '<span class="ratio-box" style="width:' + Math.max(2, Math.round(w * sc)) + 'px;height:' + Math.max(2, Math.round(h * sc)) + 'px"></span>';
      const ratio = (w / h).toFixed(2);
      const head = g !== lastGroup ? '<tr class="grp"><td colspan="5">' + esc(g) + '</td></tr>' : '';
      lastGroup = g;
      return head + '<tr><td style="width:60px">' + box + '</td><td>' + esc(name) + '</td><td class="num">' + w + ' × ' + h + ' ' + unit + '</td><td class="num">' + ratio + ' : 1</td><td style="color:var(--muted)">' + esc(memo) + '</td></tr>';
    }).join('');
    root.innerHTML = '<div class="view-head"><h2>광고 규격 참고표</h2><p>배너·소셜·인쇄·옥외의 대표 규격입니다. 왼쪽 도형은 같은 기준으로 축소한 실제 비율입니다. 매체사마다 세부 규격과 안전 영역이 다르니 집행 전에는 매체 가이드를 확인하세요.</p></div>' +
      '<div class="scroll-x"><table class="spec-table"><thead><tr><th>비율</th><th>이름</th><th>크기</th><th>가로:세로</th><th>메모</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
  }

  // ---------- 라우팅 ----------
  const VIEWS = { kv: renderKV, sources: renderSources, search: renderSearch, lab: renderLab, board: renderBoard, specs: renderSpecs };
  function current() { const h = location.hash.replace('#', ''); return VIEWS[h] ? h : 'kv'; }
  function updateCounts() {
    const saved = board.kv.length + board.site.length + board.custom.length;
    $('#counts').innerHTML = '사이트 <b>' + D.SOURCES.length + '</b> · 분류 <b>' + D.CATEGORIES.length + '</b><br>키비주얼 <b>' + D.KV.length + '</b> · 기법 <b>' + D.TECHNIQUES.length + '</b><br>내 보드 <b>' + saved + '</b>';
    const a = $('#tabs a[href="#board"]'); if (a) a.innerHTML = '내 보드' + (saved ? '<small>' + saved + '</small>' : '');
  }
  function render() {
    closeModal();
    const v = current();
    $$('#tabs a').forEach((a) => a.setAttribute('aria-current', a.getAttribute('href') === '#' + v ? 'page' : 'false'));
    const root = $('#app'); const fresh = root.cloneNode(false); root.replaceWith(fresh);
    VIEWS[v](fresh);
    if (window.matchMedia('(max-width: 860px)').matches) $$('.rail details', fresh).forEach((d) => { d.open = false; });
  }
  document.addEventListener('click', (e) => {
    const c = e.target.closest('[data-kv]');
    if (c && !e.target.closest('.modal')) { e.preventDefault(); openKV(c.dataset.kv); }
  });
  window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
  updateCounts();
  render();
})();
