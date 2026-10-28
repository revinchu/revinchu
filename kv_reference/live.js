/* 공유 데이터베이스 레이어 (claude.ai 아티팩트의 db 기능).
 * 컬렉션
 *   ads/{id}       광고 레퍼런스: TVCF 우수 인쇄광고, 주기 수집 트렌드 사례, 역분석 랩에서 등록한 이미지
 *   trends/{id}    주기 트렌드 리포트 (예약 작업이 매주 기록)
 *   overlays/{id}  클래식 레퍼런스에 덧붙인 역분석 결과
 *   meta/status    마지막 갱신 시각과 로그
 * 로컬 파일로 열면 db 가 없으므로 클래식 데이터만으로 동작한다. */
(function () {
  'use strict';
  const K = window.KVA;
  const { $, $$, esc, toast, D } = K;

  const live = { db: null, user: null, canWrite: null, ads: [], trends: [], overlays: {}, status: null, ready: false, downloads: null };
  K.live = live;

  function overlayInto(list) {
    K.store.classic.forEach((k) => { const o = live.overlays[k.id]; if (o) { k.concept = o.concept || k.concept; k.prompts = o.prompts || k.prompts; } });
    return list;
  }
  function publishItems() { overlayInto(); K.setLive(live.ads); }

  async function connect() {
    if (!window.claude || !window.claude.use) { K.store.liveStatus = '공유 DB 미연결 (로컬 파일)'; K.updateCounts(); return; }
    const [db, user, downloads] = await Promise.all([window.claude.use('db'), window.claude.use('user'), window.claude.use('downloads')]);
    live.downloads = downloads;
    live.user = user;
    if (!db) { K.store.liveStatus = '공유 DB 미연결'; K.updateCounts(); return; }
    live.db = db;
    try { live.canWrite = user ? await user.can('data.write') : null; } catch (e) { live.canWrite = null; }
    db.collection('ads').onSnapshot((snap) => {
      live.ads = snap.docs.map((d) => Object.assign({ id: d.id }, d.data()));
      live.ready = true; publishItems();
      K.store.liveStatus = '공유 DB ' + live.ads.length + '건'; K.updateCounts();
    }, () => { K.store.liveStatus = '공유 DB 연결 끊김'; K.updateCounts(); });
    db.collection('trends').onSnapshot((snap) => {
      live.trends = snap.docs.map((d) => Object.assign({ id: d.id }, d.data())).sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
      K.emit('trends');
    }, () => {});
    db.collection('overlays').onSnapshot((snap) => {
      live.overlays = {}; snap.docs.forEach((d) => { live.overlays[d.id] = d.data(); });
      publishItems();
    }, () => {});
    db.doc('meta/status').onSnapshot((snap) => { live.status = snap.exists ? snap.data() : null; K.emit('trends'); }, () => {});
  }

  // ---------- 쓰기 ----------
  const slug = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9가-힣]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'item';
  function cleanAd(r, defaults) {
    const d = defaults || {};
    const pickList = (v, allowed) => (Array.isArray(v) ? v : v ? [v] : []).map(String).filter((x) => !allowed || allowed.indexOf(x) >= 0);
    const techIds = D.TECHNIQUES.map((t) => t[0]), medIds = D.MEDIA.map((m) => m[0]), countries = D.COUNTRIES.map((c) => c[0]);
    const origin = r.origin || d.origin || 'user';
    const ad = {
      title: String(r.title || '').slice(0, 200), brand: String(r.brand || r.advertiser || '').slice(0, 120), agency: String(r.agency || r.by || '').slice(0, 160),
      year: Number(r.year) || (r.date ? Number(String(r.date).slice(0, 4)) : 0) || 0, date: String(r.date || '').slice(0, 10),
      country: countries.indexOf(String(r.country || '').toUpperCase()) >= 0 ? String(r.country).toUpperCase() : d.country || 'XX',
      field: D.FIELDS.indexOf(r.field) >= 0 ? r.field : d.field || '브랜드',
      industry: D.INDUSTRIES.indexOf(r.industry || r.ind) >= 0 ? r.industry || r.ind : '기타',
      medium: pickList(r.medium || r.med, medIds).length ? pickList(r.medium || r.med, medIds) : d.medium || ['print'],
      techniques: pickList(r.techniques || r.tech, techIds), moods: pickList(r.moods || r.mood, D.MOODS),
      col: pickList(r.col || r.palette).filter((c) => /^#[0-9a-f]{6}$/i.test(c)).slice(0, 6),
      note: String(r.note || r.summary || '').slice(0, 600), q: String(r.q || '').slice(0, 160),
      sourceUrl: /^https?:\/\//.test(r.sourceUrl || r.url || '') ? String(r.sourceUrl || r.url) : '',
      imageUrl: /^https?:\/\//.test(r.imageUrl || r.img || '') ? String(r.imageUrl || r.img) : '',
      origin: origin, addedAt: new Date().toISOString(),
    };
    if (r.thumb && /^data:image\//.test(r.thumb) && r.thumb.length < 120000) ad.thumb = r.thumb;
    if (r.features) ad.features = r.features;
    if (r.concept) ad.concept = r.concept;
    if (r.prompts) ad.prompts = r.prompts;
    if (r.analysis) ad.analysis = r.analysis;
    const id = r.id ? slug(r.id) : origin + '-' + slug((ad.brand ? ad.brand + '-' : '') + ad.title) + (ad.date ? '-' + ad.date.replace(/-/g, '') : '');
    return { id: id, data: ad };
  }
  async function saveAds(rows, defaults, onProgress) {
    if (!live.db) throw new Error('공유 DB에 연결되지 않았습니다.');
    let done = 0, failed = 0;
    for (const r of rows.slice(0, 500)) {
      const { id, data } = cleanAd(r, defaults);
      if (!data.title) { failed++; continue; }
      try { await live.db.collection('ads').doc(id).set(data); done++; } catch (e) { failed++; if (e && e.code === 'invalid_argument' && !done) throw new Error('이 계정은 공유 DB에 쓸 수 없습니다 (편집 권한 필요).'); }
      if (onProgress) onProgress(done, failed);
    }
    return { done, failed };
  }
  async function updateAd(id, patch) { if (!live.db) throw new Error('공유 DB 미연결'); await live.db.collection('ads').doc(id).update(patch); }
  async function saveOverlay(id, patch) { if (!live.db) throw new Error('공유 DB 미연결'); await live.db.collection('overlays').doc(id).set(patch); }
  async function deleteAd(id) { if (!live.db) return; await live.db.collection('ads').doc(id).delete(); }
  async function download(filename, text) {
    if (!live.downloads) { K.copy(text, '내보낼 데이터'); return; }
    try { await live.downloads.save({ filename: filename, data: text }); } catch (e) { if (e && e.code !== 'declined') { K.copy(text, '내보낼 데이터'); } }
  }
  Object.assign(live, { saveAds, updateAd, saveOverlay, deleteAd, cleanAd, download });

  // ---------- 트렌드 화면 ----------
  const trState = { region: '', industry: '', q: '' };
  function trendCard(t) {
    const list = (a) => (a || []).map((x) => '<li>' + esc(typeof x === 'string' ? x : x.text || x.title || '') + '</li>').join('');
    const ex = (t.examples || []).map((e) => '<li>' + (e.url ? '<a href="' + esc(e.url) + '" target="_blank" rel="noopener">' + esc(e.title || e.brand) + '</a>' : esc(e.title || e.brand)) + (e.brand && e.title ? ' <span class="meta-line">' + esc(e.brand) + (e.country ? ' · ' + esc(K.countryName(e.country)) : '') + '</span>' : '') + '</li>').join('');
    const src = (t.sources || []).map((s) => '<a href="' + esc(s.url) + '" target="_blank" rel="noopener">' + esc(s.title || s.url) + '</a>').join(' · ');
    return '<article class="trend"><div class="trend-meta"><span class="tag">' + esc(t.period || t.date || '') + '</span>' + (t.region ? '<span class="tag">' + esc(t.region) + '</span>' : '') + (t.industry ? '<span class="tag">' + esc(t.industry) + '</span>' : '') + (t.medium ? '<span class="tag">' + esc(t.medium) + '</span>' : '') + '</div>' +
      '<h3>' + esc(t.title) + '</h3><p>' + esc(t.summary) + '</p>' +
      (t.signals && t.signals.length ? '<p class="lbl">신호</p><ul>' + list(t.signals) + '</ul>' : '') +
      (ex ? '<p class="lbl">사례</p><ul>' + ex + '</ul>' : '') +
      (t.implications && t.implications.length ? '<p class="lbl">크리에이티브 시사점</p><ul>' + list(t.implications) + '</ul>' : '') +
      (src ? '<p class="meta-line">출처: ' + src + '</p>' : '') + '</article>';
  }
  function trendsFiltered() {
    const q = trState.q.trim().toLowerCase();
    return live.trends.filter((t) => (!trState.region || t.region === trState.region) && (!trState.industry || t.industry === trState.industry) &&
      (!q || JSON.stringify([t.title, t.summary, t.signals, t.examples, t.tags]).toLowerCase().indexOf(q) >= 0));
  }
  function renderTrends(root) {
    const regions = Array.from(new Set(live.trends.map((t) => t.region).filter(Boolean)));
    const inds = Array.from(new Set(live.trends.map((t) => t.industry).filter(Boolean)));
    const st = live.status;
    const recent = K.allItems().filter((k) => k.origin === 'trend').sort((a, b) => (b.date || '').localeCompare(a.date || '')).slice(0, 8);
    root.innerHTML = '<div class="view-head"><h2>글로벌 광고 트렌드</h2><p>예약 작업이 매주 전 세계 광고 업계 소식을 조사해 권역·업종·매체별 트렌드 리포트와 사례를 이 데이터베이스에 쌓습니다. 사례는 키비주얼 DB의 "트렌드 수집" 출처로도 들어갑니다.</p></div>' +
      '<div class="toolbar"><span class="meta-line grow">' + (st && st.lastRun ? '마지막 갱신 ' + esc(st.lastRun) + (st.note ? ' · ' + esc(st.note) : '') : live.db ? '아직 수집된 리포트가 없습니다. 첫 수집이 끝나면 여기에 나타납니다.' : '공유 DB는 claude.ai에서 이 페이지를 열 때 연결됩니다.') + '</span>' +
        '<input id="tr-q" type="search" placeholder="트렌드 검색" style="width:200px" value="' + esc(trState.q) + '">' +
        '<select id="tr-region" style="width:auto"><option value="">전체 권역</option>' + regions.map((r) => '<option' + (r === trState.region ? ' selected' : '') + '>' + esc(r) + '</option>').join('') + '</select>' +
        '<select id="tr-ind" style="width:auto"><option value="">전체 업종</option>' + inds.map((r) => '<option' + (r === trState.industry ? ' selected' : '') + '>' + esc(r) + '</option>').join('') + '</select>' +
        '<button class="btn sm" type="button" id="tr-brief">AI 트렌드 브리핑</button></div>' +
      '<div id="tr-brief-out"></div>' +
      (recent.length ? '<h3 class="cat-head">최근 수집된 사례 <small>' + recent.length + '</small></h3><div class="sheet">' + recent.map(K.kvCard).join('') + '</div>' : '') +
      '<h3 class="cat-head">리포트 <small id="tr-count"></small></h3><div class="trend-list" id="tr-list"></div>';
    const draw = () => {
      const l = trendsFiltered();
      $('#tr-count', root).textContent = l.length;
      $('#tr-list', root).innerHTML = l.length ? l.map(trendCard).join('') : '<p class="empty">표시할 트렌드 리포트가 없습니다.</p>';
    };
    draw();
    $('#tr-q', root).addEventListener('input', (e) => { trState.q = e.target.value; draw(); });
    $('#tr-region', root).addEventListener('change', (e) => { trState.region = e.target.value; draw(); });
    $('#tr-ind', root).addEventListener('change', (e) => { trState.industry = e.target.value; draw(); });
    $('#tr-brief', root).addEventListener('click', () => K.ai && K.ai.trendBrief($('#tr-brief-out', root), trendsFiltered()));
  }
  K.VIEWS.trends = renderTrends;
  K.on('trends', () => { if (K.currentView() === 'trends' && !(document.activeElement && /INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName))) K.render(); });

  // ---------- 데이터 관리 화면 ----------
  function renderData(root) {
    const byOrigin = {}; K.allItems().forEach((k) => { byOrigin[k.origin] = (byOrigin[k.origin] || 0) + 1; });
    const st = live.status;
    root.innerHTML = '<div class="view-head"><h2>데이터 관리</h2><p>공유 데이터베이스의 상태를 보고, TVCF 같은 외부 목록을 가져오고, 전체 데이터를 내보냅니다.</p></div>' +
      '<div class="stat-row">' + D.ORIGINS.map(([k, l]) => '<div class="stat"><span class="lbl">' + esc(l) + '</span><b>' + (byOrigin[k] || 0) + '</b></div>').join('') +
        '<div class="stat"><span class="lbl">트렌드 리포트</span><b>' + live.trends.length + '</b></div></div>' +
      '<p class="meta-line">' + (live.db ? '공유 DB 연결됨 · ' + (live.canWrite === false ? '읽기 전용 권한' : '쓰기 가능') : '공유 DB 미연결: claude.ai에서 열어야 등록·수집 데이터가 보입니다.') + (st && st.lastRun ? ' · 마지막 자동 수집 ' + esc(st.lastRun) : '') + '</p>' +
      (st && st.log && st.log.length ? '<details class="log"><summary>수집 로그</summary><ul>' + st.log.slice(-12).reverse().map((l) => '<li>' + esc(l) + '</li>').join('') + '</ul></details>' : '') +

      '<h3 class="cat-head">TVCF 우수 인쇄광고 가져오기</h3>' +
      '<p class="cat-desc">TVCF 목록(예: 우수 인쇄광고 · 한국 · 시작일순)의 내용을 가져오는 방법은 세 가지입니다.</p>' +
      '<ol class="steps">' +
        '<li><b>붙여넣기 + AI 정리</b> — TVCF 목록 페이지에서 전체 선택(Ctrl+A) 후 복사해 아래 칸에 붙여 넣고 "AI로 정리"를 누르면 Claude가 광고별 제목·광고주·대행사·날짜로 나눠 줍니다. 상세 링크까지 담으려면 페이지 HTML(페이지 소스 보기)을 붙여 넣으세요.</li>' +
        '<li><b>수집 스크립트</b> — 저장소의 <code>kv_reference/ingest/tvcf_ingest.py</code>를 TVCF에 접속되는 PC에서 실행하면 JSON 파일이 나옵니다. 그 내용을 아래 JSON 가져오기에 붙여 넣습니다.</li>' +
        '<li><b>자동 수집</b> — 예약 작업이 매주 TVCF 목록을 확인합니다. Claude Code 환경의 네트워크 허용 도메인에 <code>tvcf.co.kr</code>이 있어야 동작합니다.</li>' +
      '</ol>' +
      '<div class="board-form">' +
        '<div class="field full"><label for="im-text">붙여 넣을 내용 (텍스트, HTML, CSV, JSON 모두 가능)</label><textarea id="im-text" spellcheck="false" placeholder="TVCF 목록 페이지 내용이나 광고 목록을 붙여 넣으세요"></textarea></div>' +
        '<div class="field"><label for="im-origin">출처</label><select id="im-origin">' + D.ORIGINS.filter((o) => o[0] !== 'classic').map((o) => '<option value="' + o[0] + '">' + esc(o[1]) + '</option>').join('') + '</select></div>' +
        '<div class="field"><label for="im-country">기본 국가</label><select id="im-country">' + D.COUNTRIES.map((c) => '<option value="' + c[0] + '"' + (c[0] === 'KR' ? ' selected' : '') + '>' + esc(c[1]) + '</option>').join('') + '</select></div>' +
        '<div class="full toolbar" style="margin:0"><button type="button" class="btn" id="im-ai">AI로 정리</button><button type="button" class="btn ghost" id="im-json">JSON 그대로 가져오기</button><span class="meta-line" id="im-status"></span></div>' +
      '</div>' +
      '<div id="im-preview"></div>' +

      '<h3 class="cat-head">내보내기</h3>' +
      '<div class="toolbar"><button type="button" class="btn ghost sm" id="ex-ads">레퍼런스 전체 JSON</button><button type="button" class="btn ghost sm" id="ex-csv">레퍼런스 전체 CSV</button><button type="button" class="btn ghost sm" id="ex-trends">트렌드 리포트 JSON</button></div>' +
      '<h3 class="cat-head">데이터 형식</h3><p class="cat-desc">JSON 가져오기는 아래 필드를 가진 객체 배열을 받습니다. 업종·분야·기법·무드는 이 사이트의 분류 목록 값을 씁니다.</p>' +
      '<pre class="code">' + esc(JSON.stringify([{ title: '광고 제목', brand: '광고주', agency: '대행사', date: '2026-10-01', country: 'KR', field: '브랜드', industry: '식음료', medium: ['print'], techniques: ['metaphor'], moods: ['유머러스'], sourceUrl: 'https://...', imageUrl: 'https://...', note: '한 줄 설명' }], null, 1)) + '</pre>';

    if (K.bulk) K.bulk.panel(root);
    let parsed = [];
    const status = (t) => { $('#im-status', root).textContent = t; };
    const preview = () => {
      const box = $('#im-preview', root);
      if (!parsed.length) { box.innerHTML = ''; return; }
      box.innerHTML = '<div class="scroll-x"><table class="spec-table"><thead><tr><th>제목</th><th>광고주</th><th>대행사</th><th>날짜</th><th>국가</th><th>업종</th><th>링크</th></tr></thead><tbody>' +
        parsed.slice(0, 200).map((r) => '<tr><td>' + esc(r.title) + '</td><td>' + esc(r.brand || r.advertiser) + '</td><td>' + esc(r.agency) + '</td><td class="num">' + esc(r.date || r.year) + '</td><td>' + esc(r.country) + '</td><td>' + esc(r.industry) + '</td><td>' + (r.sourceUrl ? '<a href="' + esc(r.sourceUrl) + '" target="_blank" rel="noopener">열기</a>' : '') + '</td></tr>').join('') +
        '</tbody></table></div><div class="toolbar" style="margin-top:12px"><button type="button" class="btn" id="im-save">' + parsed.length + '건을 공유 DB에 등록</button><button type="button" class="btn ghost" id="im-clear">취소</button></div>';
      $('#im-save', root).addEventListener('click', async (e) => {
        e.target.disabled = true;
        try {
          const res = await saveAds(parsed, { origin: $('#im-origin', root).value, country: $('#im-country', root).value }, (d, f) => status('등록 중 ' + d + '건' + (f ? ', 실패 ' + f : '')));
          toast(res.done + '건을 등록했습니다'); status('등록 완료 ' + res.done + '건' + (res.failed ? ', 실패 ' + res.failed + '건' : '')); parsed = []; preview();
        } catch (err) { status(err.message || '등록하지 못했습니다'); e.target.disabled = false; }
      });
      $('#im-clear', root).addEventListener('click', () => { parsed = []; preview(); status(''); });
    };
    $('#im-json', root).addEventListener('click', () => {
      try { const v = JSON.parse($('#im-text', root).value); parsed = Array.isArray(v) ? v : v.items || v.ads || []; status(parsed.length + '건을 읽었습니다. 확인 후 등록하세요.'); preview(); }
      catch (e) { status('JSON 형식이 아닙니다. 텍스트라면 "AI로 정리"를 쓰세요.'); }
    });
    $('#im-ai', root).addEventListener('click', async (e) => {
      const text = $('#im-text', root).value.trim(); if (!text) return status('붙여 넣은 내용이 없습니다.');
      if (!K.ai || !K.ai.structureAds) return status('AI 기능은 claude.ai에서 열 때 동작합니다.');
      e.target.disabled = true;
      try { parsed = await K.ai.structureAds(text, { origin: $('#im-origin', root).value, country: $('#im-country', root).value }, status); status(parsed.length + '건으로 정리했습니다. 확인 후 등록하세요.'); preview(); }
      catch (err) { status(err.message || '정리하지 못했습니다'); }
      e.target.disabled = false;
    });
    const rowsForExport = () => K.allItems().map((k) => ({ id: k.id, title: k.title, brand: k.brand, agency: k.by, year: k.year, date: k.date, country: k.country, field: k.field, industry: k.ind, medium: k.med, techniques: k.tech, moods: k.mood, colors: k.col, origin: k.origin, sourceUrl: k.url, note: k.note, concept: k.concept && k.concept.concept }));
    $('#ex-ads', root).addEventListener('click', () => download('kv-atlas-references.json', JSON.stringify(rowsForExport(), null, 1)));
    $('#ex-trends', root).addEventListener('click', () => download('kv-atlas-trends.json', JSON.stringify(live.trends, null, 1)));
    $('#ex-csv', root).addEventListener('click', () => {
      const rows = rowsForExport(); const cols = ['id', 'title', 'brand', 'agency', 'year', 'date', 'country', 'field', 'industry', 'medium', 'techniques', 'moods', 'origin', 'sourceUrl', 'concept', 'note'];
      const cell = (v) => '"' + String(Array.isArray(v) ? v.join('|') : v == null ? '' : v).replace(/"/g, '""') + '"';
      download('kv-atlas-references.csv', '﻿' + cols.join(',') + '\n' + rows.map((r) => cols.map((c) => cell(r[c])).join(',')).join('\n'));
    });
  }
  K.VIEWS.data = renderData;

  connect();
})();
