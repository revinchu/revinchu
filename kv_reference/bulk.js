/* 대량 데이터 로더: bulk/manifest.json 의 샤드를 필요할 때 불러온다 (ingest/build_bulk.py 가 만든다).
 * 처음에는 AUTO_LIMIT 건까지 자동으로 불러오고, 나머지는 출처·국가 필터를 고르거나 데이터 화면에서 불러온다.
 * 로컬 file:// 로 열면 fetch 가 막혀 대량 데이터 없이 동작한다 (http.server 로 열면 된다). */
(function () {
  'use strict';
  const K = window.KVA;
  const { $, esc, toast, D } = K;
  const AUTO_LIMIT = 80000;
  const bulk = { manifest: null, loaded: new Set(), loading: new Set(), error: '' };
  K.bulk = bulk;

  const cols = () => bulk.manifest.cols;
  function toItem(row) {
    const c = cols(); const r = {};
    c.forEach((k, i) => { r[k] = row[i]; });
    ['medium', 'techniques', 'moods'].forEach((k) => { r[k] = r[k] ? String(r[k]).split('|') : []; });
    if (!r.medium.length) r.medium = ['print'];
    return r;
  }
  function note() {
    if (!bulk.manifest) return;
    const loaded = bulk.manifest.shards.filter((s) => bulk.loaded.has(s.file)).reduce((n, s) => n + s.count, 0);
    K.store.bulkNote = loaded < bulk.manifest.total ? '대량 데이터 ' + bulk.manifest.total.toLocaleString() + '건 중 ' + loaded.toLocaleString() + '건 불러옴' : '';
  }
  async function loadShard(s) {
    if (bulk.loaded.has(s.file) || bulk.loading.has(s.file)) return 0;
    bulk.loading.add(s.file);
    try {
      const res = await fetch('bulk/' + s.file);
      if (!res.ok) throw new Error(res.status);
      const data = await res.json();
      const origin = s.origin;
      const items = data.rows.map((row) => K.normalize(Object.assign(toItem(row), { origin: origin })));
      bulk.loaded.add(s.file);
      return items;
    } finally { bulk.loading.delete(s.file); }
  }
  async function loadMany(shards, label) {
    const todo = shards.filter((s) => !bulk.loaded.has(s.file) && !bulk.loading.has(s.file));
    if (!todo.length) return;
    if (label) toast(label + ' ' + todo.reduce((n, s) => n + s.count, 0).toLocaleString() + '건을 불러오는 중');
    const all = [];
    for (const s of todo) { try { const items = await loadShard(s); if (items) all.push(...items); } catch (e) { bulk.error = s.file + ' 불러오기 실패'; } }
    note();
    if (all.length) K.addBulk(all);
  }
  bulk.loadMany = loadMany;
  bulk.loadWhere = (fn, label) => bulk.manifest && loadMany(bulk.manifest.shards.filter(fn), label);

  const refresh = () => { if (K.currentView() === 'kv' && !document.querySelector('.modal')) K.render(); };
  // 출처나 국가 필터를 고르면 해당 샤드를 불러온다
  document.addEventListener('click', (e) => {
    const chip = e.target.closest('[data-chip="origin"]');
    if (chip && bulk.manifest) setTimeout(() => { const o = K.kvState.origin; if (o.size) Promise.resolve(bulk.loadWhere((s) => o.has(s.origin), '')).then(refresh); }, 0);
  });
  document.addEventListener('change', (e) => {
    if (e.target.id === 'kv-country' && bulk.manifest && e.target.value) Promise.resolve(bulk.loadWhere((s) => s.country === e.target.value, '')).then(refresh);
  });

  // 데이터 화면에 넣을 대량 데이터 패널
  bulk.panel = function (root) {
    const sec = document.createElement('section');
    const m = bulk.manifest;
    if (!m) {
      sec.innerHTML = '<h3 class="cat-head">대량 데이터</h3><p class="cat-desc">' + (bulk.error ? esc(bulk.error) : '불러온 대량 데이터 목록이 없습니다.') + ' ingest 폴더의 수집 스크립트와 build_bulk.py 로 만듭니다.</p>';
      root.appendChild(sec); return;
    }
    const byOrigin = {};
    m.shards.forEach((s) => { const o = (byOrigin[s.origin] = byOrigin[s.origin] || { count: 0, loaded: 0, files: 0 }); o.count += s.count; o.files++; if (bulk.loaded.has(s.file)) o.loaded += s.count; });
    sec.innerHTML = '<h3 class="cat-head">대량 데이터 <small>' + m.total.toLocaleString() + '건 · 샤드 ' + m.shards.length + '개</small></h3>' +
      '<p class="cat-desc">수집 스크립트로 모은 외부 데이터입니다. 브라우저가 무거워지지 않도록 ' + AUTO_LIMIT.toLocaleString() + '건까지만 자동으로 불러오고, 나머지는 필요할 때 불러옵니다. 출처 칩이나 국가를 고르면 해당 데이터가 자동으로 붙습니다.</p>' +
      '<div class="scroll-x"><table class="spec-table"><thead><tr><th>출처</th><th>건수</th><th>불러옴</th><th>샤드</th><th></th></tr></thead><tbody>' +
      Object.keys(byOrigin).map((o) => '<tr><td>' + esc(K.ORIGIN[o] || o) + '</td><td class="num">' + byOrigin[o].count.toLocaleString() + '</td><td class="num">' + byOrigin[o].loaded.toLocaleString() + '</td><td class="num">' + byOrigin[o].files + '</td><td>' +
        (byOrigin[o].loaded < byOrigin[o].count ? '<button type="button" class="btn ghost sm" data-bulkload="' + esc(o) + '">불러오기</button>' : '') + '</td></tr>').join('') +
      '</tbody></table></div>';
    root.appendChild(sec);
    sec.addEventListener('click', async (e) => {
      const b = e.target.closest('[data-bulkload]'); if (!b) return;
      b.disabled = true; b.textContent = '불러오는 중';
      await bulk.loadWhere((s) => s.origin === b.dataset.bulkload, K.ORIGIN[b.dataset.bulkload] || b.dataset.bulkload);
      K.render();
    });
  };

  (async function init() {
    try {
      const res = await fetch('bulk/manifest.json');
      if (!res.ok) throw new Error(res.status);
      bulk.manifest = await res.json();
    } catch (e) { bulk.error = '대량 데이터를 불러오지 못했습니다 (로컬 파일로 열었다면 http 서버로 여세요).'; return; }
    // 큰 샤드부터가 아니라 출처별로 고르게: 건수가 작은 샤드부터 AUTO_LIMIT 까지
    let budget = AUTO_LIMIT; const first = [];
    bulk.manifest.shards.slice().sort((a, b) => a.count - b.count).forEach((s) => { if (s.count <= budget) { first.push(s); budget -= s.count; } });
    await loadMany(first, '');
  })();
})();
