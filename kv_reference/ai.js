/* Claude 연동 기능 (claude.ai 아티팩트의 sample 기능, 보는 사람의 Claude 계정으로 실행).
 *   역분석 랩      이미지 → 분류, 커뮤니케이션 컨셉 역추론, 생각의 고리, 시각 문법, 역프롬프트, 유사 레퍼런스
 *   상세 역분석    이미지가 없는 레퍼런스도 메타데이터로 컨셉·프롬프트 역추론
 *   콘셉트 스튜디오 질문 하나에서 브리프 → 타깃 → KPI → 인사이트·컨셉 → KV 시안 → 영상 시나리오 → 기획서
 *   트렌드 브리핑  수집된 트렌드 리포트 요약
 * 이미지 유사도는 브라우저 안에서 계산한다 (주요 색 k-means 팔레트 + dHash 지각 해시 + 밝기·채도·에지 밀도). */
(function () {
  'use strict';
  const K = window.KVA;
  const { $, $$, esc, toast, D } = K;

  const ai = { sample: null, limits: null, available: false };
  K.ai = ai;

  const ERR = {
    not_granted: 'Claude 사용이 허용되지 않았습니다. 아티팩트 메뉴의 권한에서 허용할 수 있습니다.',
    sampling_disabled: '이 계정에서는 Claude를 쓸 수 없습니다.', rate_limited: '요청이 많습니다. 잠시 뒤 다시 시도하세요.',
    session_expired: '다시 로그인해야 합니다.', image_rejected: '이 이미지는 보낼 수 없습니다. JPG, PNG, WebP 파일로 다시 시도하세요.',
    refused: 'Claude가 이 요청에 답하지 않았습니다. 다른 이미지나 표현으로 시도하세요.', invalid_json: '응답 형식이 맞지 않았습니다. 다시 시도하세요.',
    prompt_too_large: '보낼 내용이 너무 깁니다. 일부만 붙여 넣으세요.', images_unavailable: '이 화면에서는 이미지를 보낼 수 없습니다.', cancelled: '중단했습니다.',
  };
  const errText = (e) => (e && ERR[e.code]) || '응답을 받지 못했습니다. 다시 시도하세요.';
  const unavailableMsg = 'AI 기능은 claude.ai에서 이 페이지를 열 때 보는 사람의 Claude 계정으로 동작합니다.';

  (async function init() {
    if (!window.claude || !window.claude.use) return;
    const s = await window.claude.use('sample');
    if (!s) return;
    ai.sample = s; ai.available = true;
    try { ai.limits = await s.limits(); } catch (e) { ai.limits = null; }
    K.emit('ai-ready');
  })();

  // 진행 표시와 함께 JSON 요청
  async function askJSON(prompt, opts, statusEl) {
    if (!ai.sample) throw { code: 'unavailable', message: unavailableMsg };
    const t0 = Date.now();
    let timer;
    if (statusEl) { statusEl.textContent = 'Claude가 생각하는 중...'; timer = setInterval(() => { statusEl.textContent = 'Claude가 생각하는 중... ' + Math.round((Date.now() - t0) / 1000) + '초'; }, 1000); }
    try {
      return await ai.sample.json(prompt, Object.assign({ onText: ({ text }) => { if (statusEl) { clearInterval(timer); statusEl.textContent = '작성 중... ' + text.length + '자'; } } }, opts || {}));
    } finally { clearInterval(timer); }
  }

  // ---------- 분류 체계 (프롬프트에 넣을 허용값) ----------
  const TAXO = 'field(하나): ' + D.FIELDS.join(', ') + '\nindustry(하나): ' + D.INDUSTRIES.join(', ') +
    '\nmedium(여러 개, 키 값): ' + D.MEDIA.map((m) => m[0] + '=' + m[1]).join(', ') +
    '\ntechniques(1~4개, 키 값): ' + D.TECHNIQUES.map((t) => t[0] + '=' + t[1]).join(', ') +
    '\nmoods(1~3개): ' + D.MOODS.join(', ') + '\ncountry: ISO 3166-1 alpha-2 대문자, 여러 나라나 판단 불가면 XX';
  function catalog(limit) {
    return K.allItems().slice(0, limit || 400).map((k) => k.id + ' | ' + k.title + ' | ' + k.brand + ' | ' + k.ind + ' | ' + k.tech.join(',') + ' | ' + k.mood.join(',')).join('\n');
  }

  // ---------- 이미지 특징 (브라우저 계산) ----------
  function dataUrlToBlob(u) {
    const [head, body] = u.split(','); const bin = atob(body); const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: (head.match(/data:([^;]+)/) || [])[1] || 'image/jpeg' });
  }
  function loadImage(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
  const toHex = (rgb) => '#' + rgb.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
  const hexToRgb = (h) => { const n = parseInt(h.slice(1), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  function rgbToLab(rgb) {
    let [r, g, b] = rgb.map((v) => { v /= 255; return v > 0.04045 ? Math.pow((v + 0.055) / 1.055, 2.4) : v / 12.92; });
    let x = (r * 0.4124 + g * 0.3576 + b * 0.1805) / 0.95047, y = r * 0.2126 + g * 0.7152 + b * 0.0722, z = (r * 0.0193 + g * 0.1192 + b * 0.9505) / 1.08883;
    const f = (t) => (t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116);
    x = f(x); y = f(y); z = f(z);
    return [116 * y - 16, 500 * (x - y), 200 * (y - z)];
  }
  const labDist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]);
  async function imageFeatures(dataUrl) {
    const img = await loadImage(dataUrl);
    const c = document.createElement('canvas'); const ctx = c.getContext('2d', { willReadFrequently: true });
    c.width = 48; c.height = 48; ctx.drawImage(img, 0, 0, 48, 48);
    const px = ctx.getImageData(0, 0, 48, 48).data; const pts = [];
    let lumSum = 0, satSum = 0;
    for (let i = 0; i < px.length; i += 4) {
      const p = [px[i], px[i + 1], px[i + 2]]; pts.push(p);
      const mx = Math.max(...p), mn = Math.min(...p); lumSum += (0.299 * p[0] + 0.587 * p[1] + 0.114 * p[2]) / 255; satSum += mx ? (mx - mn) / mx : 0;
    }
    // k-means 팔레트 (k=5)
    const sorted = pts.slice().sort((a, b) => a[0] + a[1] + a[2] - (b[0] + b[1] + b[2]));
    let cent = [0.1, 0.3, 0.5, 0.7, 0.9].map((q) => sorted[Math.floor(q * (sorted.length - 1))].slice());
    let assign = new Array(pts.length).fill(0);
    for (let it = 0; it < 10; it++) {
      for (let i = 0; i < pts.length; i++) { let best = 0, bd = Infinity; for (let j = 0; j < cent.length; j++) { const p = pts[i], q = cent[j]; const d = (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2 + (p[2] - q[2]) ** 2; if (d < bd) { bd = d; best = j; } } assign[i] = best; }
      const sum = cent.map(() => [0, 0, 0, 0]);
      for (let i = 0; i < pts.length; i++) { const s = sum[assign[i]]; s[0] += pts[i][0]; s[1] += pts[i][1]; s[2] += pts[i][2]; s[3]++; }
      cent = cent.map((q, j) => (sum[j][3] ? [sum[j][0] / sum[j][3], sum[j][1] / sum[j][3], sum[j][2] / sum[j][3]] : q));
    }
    const sizes = cent.map((_, j) => assign.filter((a) => a === j).length);
    const palette = cent.map((q, j) => ({ hex: toHex(q), share: sizes[j] / pts.length })).filter((p) => p.share > 0.01).sort((a, b) => b.share - a.share);
    // dHash 64비트
    c.width = 9; c.height = 8; ctx.drawImage(img, 0, 0, 9, 8);
    const g = ctx.getImageData(0, 0, 9, 8).data; let bits = '';
    for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) { const i = (y * 9 + x) * 4, j = (y * 9 + x + 1) * 4; bits += (g[i] + g[i + 1] + g[i + 2]) > (g[j] + g[j + 1] + g[j + 2]) ? '1' : '0'; }
    let dhash = ''; for (let i = 0; i < 64; i += 4) dhash += parseInt(bits.slice(i, i + 4), 2).toString(16);
    // 에지 밀도 (32px 그레이 라플라시안)
    c.width = 32; c.height = 32; ctx.drawImage(img, 0, 0, 32, 32);
    const e = ctx.getImageData(0, 0, 32, 32).data; let edges = 0;
    const L = (x, y) => { const i = (y * 32 + x) * 4; return (e[i] + e[i + 1] + e[i + 2]) / 3; };
    for (let y = 1; y < 31; y++) for (let x = 1; x < 31; x++) if (Math.abs(4 * L(x, y) - L(x - 1, y) - L(x + 1, y) - L(x, y - 1) - L(x, y + 1)) > 40) edges++;
    return { palette: palette.map((p) => p.hex), shares: palette.map((p) => +p.share.toFixed(3)), dhash, lum: +(lumSum / pts.length).toFixed(3), sat: +(satSum / pts.length).toFixed(3), edge: +(edges / 900).toFixed(3), aspect: +(img.width / img.height).toFixed(3) };
  }
  function paletteSim(a, b) {
    if (!a || !b || !a.length || !b.length) return 0;
    const la = a.map((h) => rgbToLab(hexToRgb(h))), lb = b.map((h) => rgbToLab(hexToRgb(h)));
    const side = (x, y) => x.reduce((s, p) => s + Math.min(...y.map((q) => labDist(p, q))), 0) / x.length;
    return Math.max(0, 1 - (side(la, lb) + side(lb, la)) / 2 / 55);
  }
  function hashSim(a, b) { if (!a || !b) return null; let d = 0; for (let i = 0; i < 16; i++) { let x = parseInt(a[i], 16) ^ parseInt(b[i], 16); while (x) { d += x & 1; x >>= 1; } } return 1 - d / 64; }
  const jacc = (a, b) => { const A = new Set(a || []), B = new Set(b || []); if (!A.size && !B.size) return 0; let n = 0; A.forEach((x) => { if (B.has(x)) n++; }); return n / (A.size + B.size - n); };
  function rankSimilar(feat, tags, excludeId) {
    // 실제 색 정보(클래식·등록 이미지)가 있는 항목만 비교한다. 대량 데이터는 색 정보가 없어 제외.
    return K.allItems().filter((k) => k.id !== excludeId && (k.colReal || k.features)).map((k) => {
      const pal = paletteSim(feat.palette, k.features ? k.features.palette : k.col);
      const hs = k.features ? hashSim(feat.dhash, k.features.dhash) : null;
      let tone = 0; if (k.features) tone = 1 - Math.min(1, (Math.abs(k.features.lum - feat.lum) + Math.abs(k.features.sat - feat.sat) + Math.abs(k.features.edge - feat.edge)) / 1.2);
      const visual = hs == null ? pal : 0.45 * pal + 0.35 * hs + 0.2 * tone;
      let concept = 0;
      if (tags) concept = 0.45 * jacc(tags.techniques, k.tech) + 0.2 * jacc(tags.moods, k.mood) + (tags.industry === k.ind ? 0.15 : 0) + (tags.field === k.field ? 0.08 : 0) + (tags.country === k.country ? 0.07 : 0) + 0.05 * jacc(tags.medium, k.med);
      return { k, visual, concept, score: tags ? 0.5 * visual + 0.5 * Math.min(1, concept * 1.6) : visual };
    }).sort((a, b) => b.score - a.score);
  }
  ai.imageFeatures = imageFeatures; ai.rankSimilar = rankSimilar;

  // ---------- 결과 렌더링 ----------
  const li = (arr) => (arr || []).map((x) => '<li>' + esc(typeof x === 'string' ? x : x.text || x.idea || JSON.stringify(x)) + '</li>').join('');
  function promptBlock(prompts) {
    if (!prompts) return '';
    return '<div class="prompts">' + D.PROMPT_TARGETS.concat([['negative', '네거티브 프롬프트']]).filter(([k]) => prompts[k]).map(([k, l]) =>
      '<div class="prompt"><div class="k"><span class="lbl">' + esc(l) + '</span><button type="button" class="btn ghost sm" data-copytext="' + esc(prompts[k]) + '">복사</button></div><p>' + esc(prompts[k]) + '</p></div>').join('') + '</div>';
  }
  function analysisHTML(a, opts) {
    if (!a) return '';
    const c = a.communication || {}, vg = a.visual_grammar || {};
    const row = (l, v) => (v ? '<dt>' + esc(l) + '</dt><dd>' + esc(v) + '</dd>' : '');
    const tagsLine = [D.FIELDS.indexOf(a.field) >= 0 ? a.field : '', a.industry, K.countryName(a.country)].filter(Boolean).map((t) => '<span class="tag">' + esc(t) + '</span>').join('') +
      (a.techniques || []).filter((t) => K.TECH[t]).map((t) => '<span class="tag">' + esc(K.TECH[t].name) + '</span>').join('') + (a.moods || []).map((m) => '<span class="tag">' + esc(m) + '</span>').join('');
    return '<div class="analysis">' +
      '<section><h3>분류</h3><div class="tags">' + tagsLine + '</div>' + (a.palette ? '<div class="swatches" style="margin-top:8px">' + a.palette.map((h) => '<i title="' + esc(h) + '" style="background:' + esc(h) + '"></i>').join('') + '</div>' : '') +
        (a.copy_read ? '<p class="meta-line">읽힌 카피: ' + esc(a.copy_read) + '</p>' : '') + (a.confidence ? '<p class="meta-line">신뢰도: ' + esc(a.confidence) + '</p>' : '') + '</section>' +
      '<section><h3>커뮤니케이션 컨셉 역추론</h3><dl class="concept">' + row('컨셉', c.concept) + row('빅 아이디어', c.big_idea) + row('타깃', c.target) + row('인사이트', c.insight) + row('풀어야 했던 문제', c.problem) + row('약속·효익', c.proposition) + row('키 메시지', c.key_message) + row('톤앤매너', c.tone) + row('행동·이유', c.action_reason) + '</dl></section>' +
      (a.logic_chain && a.logic_chain.length ? '<section><h3>이렇게 생각했을 것이다 — 생각의 고리</h3><ol class="chain">' + a.logic_chain.map((s) => '<li><b>' + esc(s.step || '') + '</b><span>' + esc(s.thought || s) + '</span>' + (s.decision ? '<em>' + esc(s.decision) + '</em>' : '') + '</li>').join('') + '</ol></section>' : '') +
      '<section><h3>시각 문법</h3><dl class="concept">' + row('레이아웃', vg.layout) + row('시선 시작점', vg.focal_point) + row('정보 위계', vg.hierarchy) + row('타이포', vg.typography) + row('색의 역할', vg.color_role) + row('카피의 역할', vg.copy_role) + row('이미지 방식', vg.image_method) + '</dl>' +
        (a.why_it_works && a.why_it_works.length ? '<p class="lbl">왜 작동하는가</p><ul>' + li(a.why_it_works) + '</ul>' : '') + '</section>' +
      '<section><h3>역프롬프트</h3>' + promptBlock(a.prompts) + '</section>' +
      (a.variations && a.variations.length ? '<section><h3>변주 아이디어</h3><ul>' + a.variations.map((v) => '<li><b>' + esc(v.idea || v) + '</b>' + (v.how ? ' — ' + esc(v.how) : '') + '</li>').join('') + '</ul></section>' : '') +
      (a.keywords ? '<section><h3>레퍼런스 검색어</h3><div class="toolbar">' + ['en', 'ko', 'ja', 'zh'].map((l) => (a.keywords[l] || []).map((q) => '<button type="button" class="chip" data-searchq="' + esc(q) + '"><span class="n">' + l + '</span> ' + esc(q) + '</button>').join('')).join('') + '</div></section>' : '') +
      (opts && opts.similarAI && a.similar && a.similar.length ? '<section><h3>Claude가 고른 비슷한 레퍼런스</h3><div class="sim-list">' + a.similar.filter((s) => K.store.byId[s.id]).map((s) => '<div class="sim">' + K.kvCard(K.store.byId[s.id]) + '<p class="th-note">' + esc(s.reason) + '</p></div>').join('') + '</div></section>' : '') +
      '</div>';
  }
  function bindCommon(box) {
    box.addEventListener('click', (e) => {
      const c = e.target.closest('[data-copytext]'); if (c) K.copy(c.dataset.copytext, '프롬프트');
      const s = e.target.closest('[data-searchq]'); if (s) { K.emit('search-q', s.dataset.searchq); location.hash = 'search'; }
    });
  }
  ai.analysisHTML = analysisHTML;

  // ---------- 역분석 프롬프트 ----------
  function reversePrompt(ctx) {
    return [
      '너는 칸 라이언즈 심사위원급 크리에이티브 디렉터이자 아트 디렉터다. ' + (ctx.image ? '첨부한 이미지는 광고 키비주얼(또는 포스터·배너)이다.' : '아래 광고 레퍼런스의 메타데이터만 보고 판단한다 (이미지는 없다).'),
      '이 광고를 만든 팀이 어떤 과제에서 출발해 어떤 인사이트와 컨셉을 거쳐 이 비주얼에 도달했는지 역으로 추론하라.',
      '확인할 수 없는 사실(브랜드, 대행사, 연도)은 지어내지 말고 "추정" 또는 "미상"으로 쓴다. 모든 설명은 한국어, 프롬프트는 영어로 쓴다.',
      ctx.meta ? '\n[레퍼런스 정보]\n' + ctx.meta : '',
      ctx.note ? '\n[사용자 메모]\n' + ctx.note : '',
      '\n[허용 분류값]\n' + TAXO,
      ctx.catalog ? '\n[데이터베이스 카탈로그: id | 제목 | 브랜드 | 업종 | 기법 | 무드]\n' + ctx.catalog : '',
      '\n다음 JSON 하나로만 답하라:',
      JSON.stringify({
        title_guess: '광고 제목 또는 한 줄 설명', brand_guess: '브랜드(추정이면 추정 표기)', country: 'KR', field: '분야 하나', industry: '업종 하나', medium: ['print'], techniques: ['metaphor'], moods: ['미니멀'],
        palette: ['#RRGGBB x5 주요 색'], copy_read: '이미지에 보이는 카피 원문 (없으면 빈 문자열)',
        communication: { target: '', insight: '타깃의 숨은 진실 한 문장', problem: '광고가 풀어야 했던 문제', proposition: '약속·효익', concept: '커뮤니케이션 컨셉 한 줄', big_idea: '빅 아이디어', key_message: '', tone: '', action_reason: '나는 ~해야 한다. 왜냐하면 ~이기 때문이다.' },
        logic_chain: [{ step: '1. 과제', thought: '어떤 비즈니스·커뮤니케이션 과제였을지', decision: '그래서 내린 결정' }, { step: '2. 인사이트', thought: '', decision: '' }, { step: '3. 컨셉', thought: '', decision: '' }, { step: '4. 시각 번역', thought: '컨셉을 어떤 이미지 장치로 바꿨는지', decision: '' }, { step: '5. 실행 디테일', thought: '구도·색·타이포·카피 선택의 이유', decision: '' }],
        visual_grammar: { layout: '', focal_point: '', hierarchy: '', typography: '', color_role: '', copy_role: '', image_method: '사진/3D/일러스트/합성 등' },
        why_it_works: ['이유 3개'],
        prompts: { midjourney: '... --ar 4:5 --style raw', flux: '', gptimage: '', imagen: '', firefly: '', video: '같은 컨셉을 6초 영상으로 만드는 프롬프트', negative: '' },
        variations: [{ idea: '같은 컨셉의 다른 실행안', how: '' }],
        keywords: { en: ['검색어'], ko: [''], ja: [''], zh: [''] },
        similar: ctx.catalog ? [{ id: '카탈로그 id', reason: '비슷한 이유' }] : [],
        confidence: '높음/중간/낮음과 근거 한 문장',
      }),
      ctx.catalog ? 'similar 에는 카탈로그에 있는 id만 최대 6개 넣는다.' : '',
    ].join('\n');
  }
  ai.reversePrompt = reversePrompt;

  // ---------- 역분석 랩 화면 ----------
  const rv = { file: null, dataUrl: '', thumb: '', feat: null, result: null, note: '', tier: 'default', ctl: null };
  function renderReverse(root) {
    const lim = ai.limits && ai.limits.images;
    root.innerHTML = '<div class="view-head"><h2>역분석 랩</h2><p>어떤 광고 이미지를 올려도 비슷한 레퍼런스를 찾고, 그 이미지가 어떤 과제·인사이트·컨셉·생각의 고리를 거쳐 나왔는지 역으로 추론하고, 같은 결의 이미지를 만드는 프롬프트를 생성합니다.</p></div>' +
      '<div class="rev-grid"><div class="rev-input">' +
        '<label class="drop big" id="rv-drop" for="rv-file">' + (rv.dataUrl ? '<img alt="" src="' + rv.dataUrl + '">' : '<span>광고 이미지를 끌어 놓거나 눌러서 고르세요<br><small>JPG · PNG · WebP</small></span>') + '</label><input id="rv-file" type="file" accept="image/jpeg,image/png,image/webp,image/gif" hidden>' +
        '<div class="field"><label for="rv-note">알고 있는 정보 (선택)</label><input id="rv-note" type="text" placeholder="예: 2025 한국 화장품 브랜드 지하철 광고" value="' + esc(rv.note) + '"></div>' +
        '<div class="field"><label for="rv-tier">분석 깊이</label><select id="rv-tier"><option value="quick">빠르게</option><option value="default">표준</option><option value="complex">깊게 (느림)</option></select></div>' +
        '<div class="toolbar"><button type="button" class="btn" id="rv-run"' + (rv.file ? '' : ' disabled') + '>역분석 시작</button><button type="button" class="btn ghost" id="rv-stop" hidden>중단</button></div>' +
        '<p class="meta-line" id="rv-status">' + (ai.available ? (lim ? '' : '이 화면에서는 이미지를 Claude에 보낼 수 없습니다. 유사 이미지 찾기는 동작합니다.') : unavailableMsg) + '</p>' +
        '<p class="lbl">외부 역이미지 검색 (이미지를 직접 올려 검색)</p><div class="toolbar">' +
          [['Google Lens', 'https://lens.google.com/'], ['Yandex', 'https://yandex.com/images/'], ['TinEye', 'https://tineye.com/'], ['Bing', 'https://www.bing.com/visualsearch'], ['Pinterest Lens', 'https://www.pinterest.com/']].map(([n, u]) => '<a class="btn ghost sm" href="' + u + '" target="_blank" rel="noopener">' + n + '</a>').join('') + '</div>' +
      '</div><div class="rev-out" style="min-width:0">' +
        '<section id="rv-sim"></section><div id="rv-result"></div><div id="rv-save"></div>' +
      '</div></div>';
    $('#rv-tier', root).value = rv.tier;
    const drop = $('#rv-drop', root), file = $('#rv-file', root);
    const take = async (f) => {
      if (!f || !/^image\//.test(f.type)) return toast('이미지 파일을 골라 주세요');
      rv.file = f; rv.result = null;
      rv.dataUrl = await K.resizeImage(f, 900, 0.85); rv.thumb = await K.resizeImage(f, 320, 0.72);
      rv.feat = await imageFeatures(rv.dataUrl);
      drop.innerHTML = '<img alt="" src="' + rv.dataUrl + '">';
      $('#rv-run', root).disabled = false; $('#rv-result', root).innerHTML = ''; $('#rv-save', root).innerHTML = '';
      drawSimilar(root);
    };
    file.addEventListener('change', () => take(file.files[0]));
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => take(e.dataTransfer.files[0]));
    root.addEventListener('paste', (e) => { const f = Array.from(e.clipboardData.files || [])[0]; if (f) take(f); });
    $('#rv-note', root).addEventListener('input', (e) => { rv.note = e.target.value; });
    $('#rv-tier', root).addEventListener('change', (e) => { rv.tier = e.target.value; });
    $('#rv-stop', root).addEventListener('click', () => rv.ctl && rv.ctl.abort());
    $('#rv-run', root).addEventListener('click', () => runReverse(root));
    bindCommon(root);
    if (rv.feat) drawSimilar(root);
    if (rv.result) showResult(root);
  }
  function drawSimilar(root) {
    const box = $('#rv-sim', root); if (!box || !rv.feat) return;
    const ranked = rankSimilar(rv.feat, rv.result, null).slice(0, 12);
    box.innerHTML = '<h3 class="cat-head">비슷한 레퍼런스 <small>' + (rv.result ? '색·구성 + 컨셉·기법 종합' : '색·구성 기준 · 역분석 후 컨셉까지 반영') + '</small></h3>' +
      '<div class="swatches" title="추출한 주요 색">' + rv.feat.palette.map((h, i) => '<i style="background:' + h + ';flex:' + (rv.feat.shares[i] || 0.1) + '"></i>').join('') + '</div>' +
      '<p class="meta-line">밝기 ' + rv.feat.lum + ' · 채도 ' + rv.feat.sat + ' · 디테일 ' + rv.feat.edge + ' · 비율 ' + rv.feat.aspect + ' · dHash ' + rv.feat.dhash + '</p>' +
      '<div class="sim-list">' + ranked.map((r) => '<div class="sim">' + K.kvCard(r.k) + '<p class="meter"><span style="width:' + Math.round(r.score * 100) + '%"></span></p><p class="th-note">유사도 ' + Math.round(r.score * 100) + ' · 색·구성 ' + Math.round(r.visual * 100) + (rv.result ? ' · 컨셉 ' + Math.round(Math.min(1, r.concept * 1.6) * 100) : '') + '</p></div>').join('') + '</div>';
  }
  async function runReverse(root) {
    const status = $('#rv-status', root);
    if (!ai.available) { status.textContent = unavailableMsg; return; }
    if (!ai.limits || !ai.limits.images) { status.textContent = '이 화면에서는 이미지를 Claude에 보낼 수 없습니다.'; return; }
    rv.ctl = new AbortController();
    $('#rv-run', root).disabled = true; $('#rv-stop', root).hidden = false;
    try {
      const a = await askJSON(reversePrompt({ image: true, note: rv.note, catalog: catalog(350) }), { images: [rv.file], modelTier: rv.tier, signal: rv.ctl.signal }, status);
      rv.result = a; status.textContent = '역분석을 마쳤습니다.';
      showResult(root); drawSimilar(root);
    } catch (e) { status.textContent = errText(e); }
    $('#rv-run', root).disabled = false; $('#rv-stop', root).hidden = true;
  }
  function showResult(root) {
    $('#rv-result', root).innerHTML = analysisHTML(rv.result, { similarAI: true });
    const live = K.live;
    const save = $('#rv-save', root);
    if (!live || !live.db) { save.innerHTML = '<p class="meta-line">공유 DB에 연결되면 이 분석을 데이터베이스에 등록할 수 있습니다.</p>'; return; }
    save.innerHTML = '<div class="board-form"><div class="field"><label for="rv-title">제목</label><input id="rv-title" type="text" value="' + esc(rv.result.title_guess || '') + '"></div>' +
      '<div class="field"><label for="rv-brand">브랜드</label><input id="rv-brand" type="text" value="' + esc(rv.result.brand_guess || '') + '"></div>' +
      '<div class="field"><label for="rv-url">출처 링크 (선택)</label><input id="rv-url" type="url" placeholder="https://"></div>' +
      '<div class="field"><label for="rv-year">연도 (선택)</label><input id="rv-year" type="text" inputmode="numeric" placeholder="2026"></div>' +
      '<div class="full toolbar" style="margin:0"><button type="button" class="btn" id="rv-commit">데이터베이스에 등록</button><span class="meta-line">썸네일(320px)과 분석 결과가 팀과 공유됩니다. 권리를 확인한 이미지만 등록하세요.</span></div></div>';
    $('#rv-commit', root).addEventListener('click', async (e) => {
      e.target.disabled = true;
      const a = rv.result;
      try {
        await live.saveAds([{
          title: $('#rv-title', root).value || a.title_guess, brand: $('#rv-brand', root).value, sourceUrl: $('#rv-url', root).value, year: $('#rv-year', root).value,
          country: a.country, field: a.field, industry: a.industry, medium: a.medium, techniques: a.techniques, moods: a.moods, col: (a.palette || []).length ? a.palette : rv.feat.palette,
          note: (a.communication && a.communication.concept) || '', thumb: rv.thumb, features: rv.feat,
          concept: a.communication || null, prompts: a.prompts || null, analysis: { logic_chain: a.logic_chain || [], visual_grammar: a.visual_grammar || {}, why_it_works: a.why_it_works || [], variations: a.variations || [] },
          origin: 'user', date: new Date().toISOString().slice(0, 10),
        }]);
        toast('데이터베이스에 등록했습니다'); e.target.textContent = '등록됨';
      } catch (err) { toast(err.message || '등록하지 못했습니다'); e.target.disabled = false; }
    });
  }
  K.VIEWS.reverse = renderReverse;
  K.on('ai-ready', () => { if (K.currentView() === 'reverse' || K.currentView() === 'studio') K.render(); });

  // ---------- 상세 화면: 메타데이터 기반 역분석 ----------
  K.modalHooks.push((k) => ({
    html: '<div class="ai-box"><p class="lbl">AI 역분석</p>' +
      (k.analysis ? analysisHTML(Object.assign({ communication: k.concept, prompts: k.prompts }, k.analysis)) : k.prompts ? '<section><h3>역프롬프트</h3>' + promptBlock(k.prompts) + '</section>' : '') +
      '<div class="toolbar"><button type="button" class="btn sm" data-ai-meta>' + (k.concept ? '컨셉·프롬프트 다시 추론' : '컨셉·생각의 고리·프롬프트 역추론') + '</button><span class="meta-line" data-ai-status>' + (ai.available ? (k.thumb ? '등록된 썸네일로 분석합니다' : '이미지 없이 메타데이터로 추론합니다') : unavailableMsg) + '</span></div><div data-ai-out></div></div>',
    bind(box) {
      bindCommon(box);
      const btn = $('[data-ai-meta]', box); if (!btn) return;
      btn.addEventListener('click', async () => {
        const st = $('[data-ai-status]', box);
        if (!ai.available) { st.textContent = unavailableMsg; return; }
        btn.disabled = true;
        const meta = ['제목: ' + k.title, '브랜드: ' + k.brand, '연도: ' + (k.year || '미상'), '제작: ' + k.by, '국가: ' + K.countryName(k.country), '분야: ' + k.field, '업종: ' + k.ind, '매체: ' + k.med.join(','), '기법: ' + k.tech.map((t) => K.TECH[t].name).join(','), '무드: ' + k.mood.join(','), '색: ' + k.col.join(' '), '설명: ' + k.note].join('\n');
        try {
          const opts = { modelTier: 'default' };
          let image = false;
          if (k.thumb && ai.limits && ai.limits.images) { opts.images = [dataUrlToBlob(k.thumb)]; image = true; }
          const a = await askJSON(reversePrompt({ image: image, meta: meta }), opts, st);
          $('[data-ai-out]', box).innerHTML = analysisHTML(a);
          st.textContent = image ? '썸네일과 메타데이터로 추론했습니다.' : '메타데이터 기반 추론입니다. 실제 제작 의도와 다를 수 있습니다.';
          const live = K.live;
          if (live && live.db) {
            const patch = { concept: a.communication || null, prompts: a.prompts || null, analysis: { logic_chain: a.logic_chain || [], visual_grammar: a.visual_grammar || {}, why_it_works: a.why_it_works || [], variations: a.variations || [] } };
            const saveBtn = document.createElement('button'); saveBtn.type = 'button'; saveBtn.className = 'btn ghost sm'; saveBtn.textContent = '이 분석을 레퍼런스에 저장';
            saveBtn.addEventListener('click', async () => {
              saveBtn.disabled = true;
              try { if (k.origin === 'classic') await live.saveOverlay(k.id, patch); else await live.updateAd(k.id, patch); toast('저장했습니다'); saveBtn.textContent = '저장됨'; }
              catch (e) { toast('저장하지 못했습니다'); saveBtn.disabled = false; }
            });
            $('[data-ai-out]', box).prepend(saveBtn);
          }
        } catch (e) { st.textContent = errText(e); }
        btn.disabled = false;
      });
    },
  }));

  // ---------- 붙여넣기 → 광고 목록 구조화 (데이터 관리) ----------
  ai.structureAds = async function (text, defaults, status) {
    const chunk = text.slice(0, 60000);
    const prompt = '아래는 광고 아카이브 페이지(예: TVCF 우수 인쇄광고 목록)에서 복사한 내용이다. 광고 한 편마다 객체 하나로 정리해 JSON 배열로만 답하라.\n' +
      '필드: title, brand(광고주), agency(대행사·제작사, 없으면 빈 문자열), date(YYYY-MM-DD, 모르면 빈 문자열), year, country(ISO 대문자, 기본 ' + defaults.country + '), field, industry, medium(키 배열, 인쇄광고면 ["print"]), sourceUrl(상세 페이지 절대 URL, 상대경로면 https://tvcf.co.kr 를 붙인다), imageUrl(이미지 절대 URL, 있으면), note(한 줄 설명).\n' +
      '내용에 없는 정보는 지어내지 말고 빈 문자열로 둔다. 메뉴·광고 배너·페이지 안내 문구는 제외한다.\n[허용 분류값]\n' + TAXO + '\n\n[내용]\n' + chunk;
    const rows = await askJSON(prompt, { modelTier: 'default' }, status ? { set textContent(v) { status(v); } } : null);
    if (!Array.isArray(rows)) throw new Error('목록을 찾지 못했습니다.');
    return rows.map((r) => Object.assign({ origin: defaults.origin }, r));
  };

  // ---------- 트렌드 브리핑 ----------
  ai.trendBrief = async function (out, trends) {
    if (!ai.available) { out.innerHTML = '<p class="meta-line">' + unavailableMsg + '</p>'; return; }
    const recent = K.allItems().filter((k) => k.origin !== 'classic').slice(0, 80).map((k) => k.title + ' / ' + k.brand + ' / ' + K.countryName(k.country) + ' / ' + k.ind + ' / ' + k.tech.join(',')).join('\n');
    const src = trends.slice(0, 30).map((t) => '- [' + (t.period || t.date) + ' ' + (t.region || '') + ' ' + (t.industry || '') + '] ' + t.title + ': ' + t.summary).join('\n');
    if (!src && !recent) { out.innerHTML = '<p class="meta-line">브리핑할 트렌드 데이터가 아직 없습니다.</p>'; return; }
    out.innerHTML = '<div class="ai-box"><p class="meta-line" data-st>Claude가 생각하는 중...</p><div data-o></div></div>';
    const st = $('[data-st]', out);
    try {
      const b = await askJSON('너는 글로벌 광고 전략 디렉터다. 아래 트렌드 리포트와 최근 수집 사례만 근거로, 디자이너와 마케터를 위한 브리핑을 한국어로 작성하라. 근거에 없는 사실은 쓰지 않는다.\n' +
        'JSON: {"headline":"", "top_trends":[{"name":"","what":"","evidence":"","how_to_use":""}], "by_region":[{"region":"","note":""}], "visual_signals":["색·타이포·구도 경향"], "next_brief_ideas":["바로 써볼 아이디어"]}\n\n[트렌드 리포트]\n' + src + '\n\n[최근 사례]\n' + recent, { modelTier: 'default' }, st);
      st.textContent = '';
      $('[data-o]', out).innerHTML = '<h3>' + esc(b.headline) + '</h3>' + (b.top_trends || []).map((t) => '<div class="trend"><h3>' + esc(t.name) + '</h3><p>' + esc(t.what) + '</p><p class="meta-line">근거: ' + esc(t.evidence) + '</p><p><b>활용:</b> ' + esc(t.how_to_use) + '</p></div>').join('') +
        (b.by_region && b.by_region.length ? '<p class="lbl">권역별</p><ul>' + b.by_region.map((r) => '<li><b>' + esc(r.region) + '</b> ' + esc(r.note) + '</li>').join('') + '</ul>' : '') +
        (b.visual_signals ? '<p class="lbl">시각 신호</p><ul>' + li(b.visual_signals) + '</ul>' : '') + (b.next_brief_ideas ? '<p class="lbl">바로 써볼 아이디어</p><ul>' + li(b.next_brief_ideas) + '</ul>' : '');
    } catch (e) { st.textContent = errText(e); }
  };

  // ---------- 콘셉트 스튜디오 ----------
  const STUDIO_KEY = 'kvatlas.studio.v1';
  let studio = { projects: [], current: null };
  try { const raw = localStorage.getItem(STUDIO_KEY); if (raw) studio = Object.assign(studio, JSON.parse(raw)); } catch (e) { /* 없음 */ }
  const saveStudio = () => { try { localStorage.setItem(STUDIO_KEY, JSON.stringify(studio)); } catch (e) { toast('브라우저 저장 공간이 부족합니다'); } };
  const project = () => studio.projects.find((p) => p.id === studio.current);
  function newProject(question) { const p = { id: 'p' + Date.now().toString(36), name: (question || '새 프로젝트').slice(0, 40), question: question || '', inputs: {}, steps: {}, created: new Date().toISOString().slice(0, 10) }; studio.projects.unshift(p); studio.current = p.id; saveStudio(); return p; }
  K.on('studio-brief', (text) => { newProject(text); });

  const ctxOf = (p, keys) => keys.filter((k) => p.steps[k]).map((k) => '[' + STEPS.find((s) => s.key === k).title + ']\n' + JSON.stringify(p.steps[k])).join('\n\n');
  const STEPS = [
    { key: 'brief', title: '브리프 정리', desc: '질문 한 줄이나 메모를 광고 기획 브리프로 구조화합니다.',
      prompt: (p) => '너는 광고대행사의 시니어 AE다. 아래 질문·메모를 광고 캠페인 브리프로 정리하라. 빠진 정보는 합리적으로 가정하되 assumptions 에 명시한다. 한국어.\n[질문]\n' + p.question + '\n[추가 정보]\n' + JSON.stringify(p.inputs) +
        '\nJSON: {"brand":"","product":"","industry":"(허용값 중 하나)","market_country":"ISO","objective":"","business_problem":"","communication_problem":"","budget_level":"","period":"","media":[""],"mandatories":[""],"assumptions":[""]}\n허용 업종: ' + D.INDUSTRIES.join(', '),
      view: (r) => kv([['브랜드', r.brand], ['제품', r.product], ['업종', r.industry], ['시장', K.countryName(r.market_country)], ['목표', r.objective], ['비즈니스 문제', r.business_problem], ['커뮤니케이션 문제', r.communication_problem], ['예산', r.budget_level], ['기간', r.period], ['매체', (r.media || []).join(', ')], ['필수 사항', (r.mandatories || []).join(', ')]]) + list('가정', r.assumptions) },
    { key: 'target', title: '타깃 세그먼트', desc: '후보 세그먼트를 비교하고 주 타깃을 고릅니다.',
      prompt: (p) => '너는 소비자 인사이트 플래너다. 브리프를 바탕으로 타깃 세그먼트 3개를 비교하고 주 타깃을 추천하라. 추정치는 추정이라고 표기한다. 한국어.\n' + ctxOf(p, ['brief']) +
        '\nJSON: {"segments":[{"name":"","profile":"","needs":"","barriers":"","media_habits":"","size_note":"","fit_score":1}],"primary":"세그먼트 이름","rationale":"","persona":{"name":"","day_in_life":"","quote":""}}',
      view: (r) => '<div class="seg-grid">' + (r.segments || []).map((s) => '<div class="seg' + (s.name === r.primary ? ' on' : '') + '"><b>' + esc(s.name) + '</b><span class="meta-line">적합도 ' + esc(s.fit_score) + '/5</span>' + kv([['프로필', s.profile], ['니즈', s.needs], ['장벽', s.barriers], ['매체 습관', s.media_habits], ['규모', s.size_note]]) + '</div>').join('') + '</div>' +
        '<p><b>주 타깃: ' + esc(r.primary) + '</b> — ' + esc(r.rationale) + '</p>' + (r.persona ? '<blockquote>' + esc(r.persona.quote) + '<br><small>' + esc(r.persona.name) + ' · ' + esc(r.persona.day_in_life) + '</small></blockquote>' : '') },
    { key: 'kpi', title: 'KPI 설계', desc: '퍼널 단계별 목표와 측정 방법을 정합니다.',
      prompt: (p) => '너는 퍼포먼스·브랜드 마케팅 전략가다. 퍼널 단계별 KPI를 설계하라. 목표치는 업종 일반 범위의 가정이라고 밝힌다. 한국어.\n' + ctxOf(p, ['brief', 'target']) +
        '\nJSON: {"north_star":"","funnel":[{"stage":"인지/고려/전환/충성","objective":"","kpi":"","target_range":"","measurement":""}],"assumptions":[""]}',
      view: (r) => '<p><b>북극성 지표:</b> ' + esc(r.north_star) + '</p><div class="scroll-x"><table class="spec-table"><thead><tr><th>단계</th><th>목표</th><th>KPI</th><th>목표 범위</th><th>측정</th></tr></thead><tbody>' + (r.funnel || []).map((f) => '<tr><td>' + esc(f.stage) + '</td><td>' + esc(f.objective) + '</td><td>' + esc(f.kpi) + '</td><td>' + esc(f.target_range) + '</td><td>' + esc(f.measurement) + '</td></tr>').join('') + '</tbody></table></div>' + list('가정', r.assumptions) },
    { key: 'concept', title: '인사이트·컨셉', desc: '인사이트에서 컨셉 후보를 뽑고, 생각의 고리와 평가 점수를 붙입니다.',
      prompt: (p) => '너는 칸 라이언즈 수상 경력의 크리에이티브 디렉터다. 인사이트 3개와 서로 결이 다른 컨셉 3개를 제안하라. 각 컨셉은 인사이트에서 컨셉까지의 생각의 고리와 리스크, 4개 기준 점수(1~5)를 가진다. 진부한 표현은 피한다. 한국어.\n' + ctxOf(p, ['brief', 'target', 'kpi']) +
        '\nJSON: {"insights":[{"insight":"","tension":"","evidence_type":"관찰/데이터/문화"}],"concepts":[{"name":"","concept_line":"","key_message":"","logic_chain":["인사이트→","→"],"why_it_works":"","risks":"","scores":{"차별성":4,"적합성":4,"확장성":4,"실행성":4}}],"recommended":"컨셉 이름","recommend_reason":""}',
      view: (r) => list('인사이트', (r.insights || []).map((i) => i.insight + ' (' + i.tension + ')')) + '<div class="seg-grid">' + (r.concepts || []).map((c) => '<div class="seg' + (c.name === r.recommended ? ' on' : '') + '"><b>' + esc(c.name) + '</b><p class="big-line">' + esc(c.concept_line) + '</p>' + kv([['키 메시지', c.key_message], ['작동 이유', c.why_it_works], ['리스크', c.risks]]) +
        '<ol class="chain small">' + (c.logic_chain || []).map((s) => '<li><span>' + esc(s) + '</span></li>').join('') + '</ol><p class="meta-line">' + Object.entries(c.scores || {}).map(([k2, v]) => esc(k2) + ' ' + esc(v)).join(' · ') + '</p></div>').join('') + '</div><p><b>추천: ' + esc(r.recommended) + '</b> — ' + esc(r.recommend_reason) + '</p>' },
    { key: 'creative', title: 'KV 시안·매체 전개', desc: '추천 컨셉을 KV 시안 3종과 매체별 전개, 이미지 생성 프롬프트로 구체화합니다.',
      prompt: (p) => '너는 아트 디렉터다. 추천 컨셉(또는 사용자가 지정한 컨셉)을 KV 시안 3종으로 구체화하라. 각 시안은 레이아웃, 헤드라인, 서브카피, 비주얼 묘사, 기법, 팔레트, 이미지 생성 프롬프트(영어), 참고할 데이터베이스 레퍼런스 id(카탈로그에 있는 것만)를 가진다. 한국어(프롬프트만 영어).\n' + ctxOf(p, ['brief', 'target', 'concept']) +
        (p.inputs.pick ? '\n사용자 지정 컨셉: ' + p.inputs.pick : '') + '\n[허용 기법 id] ' + D.TECHNIQUES.map((t) => t[0]).join(', ') + '\n[카탈로그]\n' + catalog(250) +
        '\nJSON: {"concept_used":"","kv_options":[{"title":"","layout":"","headline":"","subcopy":"","visual":"","techniques":["id"],"palette":["#RRGGBB"],"prompt_midjourney":"","prompt_general":"","reference_ids":["id"]}],"media_rollout":[{"medium":"","spec":"","adaptation":""}],"copy_variants":{"ko":[""],"en":[""]}}',
      view: (r) => '<p class="meta-line">사용한 컨셉: ' + esc(r.concept_used) + '</p><div class="seg-grid">' + (r.kv_options || []).map((o) => '<div class="seg"><b>' + esc(o.title) + '</b>' + (o.palette ? '<div class="swatches">' + o.palette.map((h) => '<i style="background:' + esc(h) + '"></i>').join('') + '</div>' : '') +
        '<p class="big-line">' + esc(o.headline) + '</p><p>' + esc(o.subcopy) + '</p>' + kv([['레이아웃', o.layout], ['비주얼', o.visual], ['기법', (o.techniques || []).map((t) => (K.TECH[t] ? K.TECH[t].name : t)).join(', ')]]) +
        '<div class="prompt"><div class="k"><span class="lbl">Midjourney</span><button type="button" class="btn ghost sm" data-copytext="' + esc(o.prompt_midjourney) + '">복사</button></div><p>' + esc(o.prompt_midjourney) + '</p></div>' +
        '<div class="prompt"><div class="k"><span class="lbl">범용 프롬프트</span><button type="button" class="btn ghost sm" data-copytext="' + esc(o.prompt_general) + '">복사</button></div><p>' + esc(o.prompt_general) + '</p></div>' +
        ((o.reference_ids || []).filter((id) => K.store.byId[id]).length ? '<div class="mini-row">' + o.reference_ids.filter((id) => K.store.byId[id]).slice(0, 4).map((id) => K.kvCard(K.store.byId[id])).join('') + '</div>' : '') + '</div>').join('') + '</div>' +
        '<p class="lbl">매체 전개</p><ul>' + (r.media_rollout || []).map((m) => '<li><b>' + esc(m.medium) + '</b> ' + esc(m.spec) + ' — ' + esc(m.adaptation) + '</li>').join('') + '</ul>' + (r.copy_variants ? list('카피 변주 (한국어)', r.copy_variants.ko) + list('카피 변주 (영어)', r.copy_variants.en) : '') },
    { key: 'film', title: '영상 시나리오', desc: '15초·30초 스토리보드와 컷별 영상 생성 프롬프트를 만듭니다.',
      prompt: (p) => '너는 광고 영상 감독이다. 선택된 컨셉과 KV를 바탕으로 15초와 30초 광고 스토리보드를 만들어라. 컷마다 화면, 카피·내레이션, 사운드, 영상 생성 프롬프트(영어)를 쓴다. 한국어.\n' + ctxOf(p, ['brief', 'concept', 'creative']) +
        '\nJSON: {"films":[{"length":"15s","title":"","logline":"","scenes":[{"no":1,"duration":"0-3s","visual":"","copy":"","sound":"","video_prompt":""}]}]}',
      view: (r) => (r.films || []).map((f) => '<h4>' + esc(f.length) + ' · ' + esc(f.title) + '</h4><p>' + esc(f.logline) + '</p><div class="scroll-x"><table class="spec-table"><thead><tr><th>#</th><th>시간</th><th>화면</th><th>카피</th><th>사운드</th><th>영상 프롬프트</th></tr></thead><tbody>' +
        (f.scenes || []).map((s) => '<tr><td class="num">' + esc(s.no) + '</td><td class="num">' + esc(s.duration) + '</td><td>' + esc(s.visual) + '</td><td>' + esc(s.copy) + '</td><td>' + esc(s.sound) + '</td><td><button type="button" class="btn ghost sm" data-copytext="' + esc(s.video_prompt) + '">복사</button> ' + esc(s.video_prompt) + '</td></tr>').join('') + '</tbody></table></div>').join('') },
    { key: 'review', title: '반론·리스크 점검', desc: '심사위원과 광고주 입장에서 약점을 찾고 고칠 점을 제안합니다.',
      prompt: (p) => '너는 깐깐한 광고제 심사위원이자 광고주 마케팅 임원이다. 지금까지의 기획을 비판적으로 검토하라: 진부함, 브랜드 적합성, 법·윤리 리스크(과장, 비교, 인물·저작권), 실행 난이도, KPI와 크리에이티브의 연결. 각 지적에 고칠 방법을 붙인다. 한국어.\n' + ctxOf(p, ['brief', 'target', 'kpi', 'concept', 'creative']) +
        '\nJSON: {"verdict":"한 문장 총평","score":7,"issues":[{"area":"","problem":"","fix":""}],"stronger_version":"한 단계 더 나간 컨셉 제안"}',
      view: (r) => '<p class="big-line">' + esc(r.verdict) + ' <span class="meta-line">' + esc(r.score) + '/10</span></p><ul>' + (r.issues || []).map((i) => '<li><b>' + esc(i.area) + '</b> ' + esc(i.problem) + '<br><span class="meta-line">고칠 점: ' + esc(i.fix) + '</span></li>').join('') + '</ul><p><b>더 나간 버전:</b> ' + esc(r.stronger_version) + '</p>' },
  ];
  function kv(rows) { return '<dl class="concept">' + rows.filter((r) => r[1]).map((r) => '<dt>' + esc(r[0]) + '</dt><dd>' + esc(r[1]) + '</dd>').join('') + '</dl>'; }
  function list(label, arr) { return arr && arr.length ? '<p class="lbl">' + esc(label) + '</p><ul>' + li(arr) + '</ul>' : ''; }

  function docMarkdown(p) {
    const s = p.steps; const out = ['# ' + p.name, '', '> ' + p.question.replace(/\n/g, ' '), ''];
    if (s.brief) out.push('## 브리프', '- 브랜드: ' + s.brief.brand, '- 제품: ' + s.brief.product, '- 시장: ' + K.countryName(s.brief.market_country), '- 목표: ' + s.brief.objective, '- 비즈니스 문제: ' + s.brief.business_problem, '- 커뮤니케이션 문제: ' + s.brief.communication_problem, '');
    if (s.target) out.push('## 타깃', '주 타깃: **' + s.target.primary + '** — ' + s.target.rationale, '', ...(s.target.segments || []).map((g) => '- ' + g.name + ': ' + g.profile + ' / 니즈 ' + g.needs + ' / 장벽 ' + g.barriers), '');
    if (s.kpi) out.push('## KPI', '북극성 지표: ' + s.kpi.north_star, '', '| 단계 | KPI | 목표 범위 | 측정 |', '|---|---|---|---|', ...(s.kpi.funnel || []).map((f) => '| ' + [f.stage, f.kpi, f.target_range, f.measurement].join(' | ') + ' |'), '');
    if (s.concept) out.push('## 인사이트와 컨셉', ...(s.concept.insights || []).map((i) => '- 인사이트: ' + i.insight), '', ...(s.concept.concepts || []).map((c) => '### ' + c.name + (c.name === s.concept.recommended ? ' (추천)' : '') + '\n' + c.concept_line + '\n\n- 키 메시지: ' + c.key_message + '\n- 생각의 고리: ' + (c.logic_chain || []).join(' ') + '\n- 리스크: ' + c.risks), '');
    if (s.creative) out.push('## KV 시안', ...(s.creative.kv_options || []).map((o) => '### ' + o.title + '\n- 헤드라인: ' + o.headline + '\n- 서브카피: ' + o.subcopy + '\n- 레이아웃: ' + o.layout + '\n- 비주얼: ' + o.visual + '\n- 프롬프트: `' + o.prompt_midjourney + '`'), '', '### 매체 전개', ...(s.creative.media_rollout || []).map((m) => '- ' + m.medium + ' (' + m.spec + '): ' + m.adaptation), '');
    if (s.film) out.push('## 영상 시나리오', ...(s.film.films || []).map((f) => '### ' + f.length + ' ' + f.title + '\n' + f.logline + '\n\n| # | 시간 | 화면 | 카피 | 사운드 |\n|---|---|---|---|---|\n' + (f.scenes || []).map((c) => '| ' + [c.no, c.duration, c.visual, c.copy, c.sound].join(' | ') + ' |').join('\n')), '');
    if (s.review) out.push('## 리스크 점검', s.review.verdict, '', ...(s.review.issues || []).map((i) => '- ' + i.area + ': ' + i.problem + ' → ' + i.fix), '', '더 나간 버전: ' + s.review.stronger_version, '');
    return out.join('\n');
  }

  function renderStudio(root) {
    const p = project();
    root.innerHTML = '<div class="view-head"><h2>콘셉트 스튜디오</h2><p>질문 하나에서 시작해 브리프, 타깃 세그먼트, KPI, 인사이트와 컨셉, KV 시안과 매체 전개, 영상 시나리오, 리스크 점검, 기획서까지 한 흐름으로 만듭니다. 단계마다 고칠 점을 지시해 다시 만들 수 있습니다.</p></div>' +
      '<div class="studio-bar"><select id="st-proj" aria-label="프로젝트"><option value="">새 프로젝트</option>' + studio.projects.map((x) => '<option value="' + x.id + '"' + (p && x.id === p.id ? ' selected' : '') + '>' + esc(x.name) + ' · ' + esc(x.created) + '</option>').join('') + '</select>' +
        (p ? '<button type="button" class="btn ghost sm" id="st-del">프로젝트 삭제</button>' : '') + '</div>' +
      (!p ? '<div class="board-form"><div class="field full"><label for="st-q">무엇을 광고하나요? 한 줄이면 충분합니다</label><textarea id="st-q" placeholder="예: 30대 직장인을 위한 저당 단백질 음료를 한국에서 출시한다. 3개월 안에 인지도를 높이고 싶다."></textarea></div>' +
        '<div class="field"><label for="st-brand">브랜드 (선택)</label><input id="st-brand" type="text"></div><div class="field"><label for="st-country">시장</label><select id="st-country">' + D.COUNTRIES.map((c) => '<option value="' + c[0] + '"' + (c[0] === 'KR' ? ' selected' : '') + '>' + esc(c[1]) + '</option>').join('') + '</select></div>' +
        '<div class="full toolbar" style="margin:0"><button type="button" class="btn" id="st-new">프로젝트 시작</button><span class="meta-line">' + (ai.available ? '' : unavailableMsg) + '</span></div></div>' :
      '<p class="question">' + esc(p.question) + '</p>' +
      '<div class="toolbar"><button type="button" class="btn" id="st-all">남은 단계 모두 실행</button><button type="button" class="btn ghost" id="st-stop" hidden>중단</button><button type="button" class="btn ghost" id="st-doc">기획서 복사</button><button type="button" class="btn ghost" id="st-dl">기획서 내려받기 (.md)</button><span class="meta-line" id="st-status"></span></div>' +
      '<ol class="steps-list">' + STEPS.map((s, i) => {
        const r = p.steps[s.key];
        return '<li class="step' + (r ? ' done' : '') + '" data-step="' + s.key + '"><div class="step-head"><span class="num">' + (i + 1) + '</span><div><h3>' + esc(s.title) + '</h3><p class="meta-line">' + esc(s.desc) + '</p></div>' +
          '<div class="toolbar" style="margin:0"><button type="button" class="btn sm" data-run="' + s.key + '">' + (r ? '다시 만들기' : '실행') + '</button></div></div>' +
          (r ? '<div class="step-body">' + s.view(r) + '</div>' : '') +
          '<div class="field step-fb"><label for="fb-' + s.key + '">고칠 점 지시 (다시 만들 때 반영)</label><input id="fb-' + s.key + '" type="text" placeholder="' + (s.key === 'creative' ? '예: 두 번째 컨셉으로, 더 미니멀하게' : '예: 더 대담하게, 20대 여성 중심으로') + '" value="' + esc((p.inputs.fb || {})[s.key] || '') + '"></div></li>';
      }).join('') + '</ol>');
    bindCommon(root);
    $('#st-proj', root).addEventListener('change', (e) => { studio.current = e.target.value || null; saveStudio(); K.render(); });
    if (!p) {
      $('#st-new', root).addEventListener('click', () => {
        const q = $('#st-q', root).value.trim(); if (!q) return toast('무엇을 광고하는지 한 줄 적어 주세요');
        const np = newProject(q); np.inputs = { brand: $('#st-brand', root).value.trim(), market: $('#st-country', root).value }; saveStudio(); K.render();
      });
      return;
    }
    $('#st-del', root).addEventListener('click', (e) => {
      if (!e.target.dataset.confirm) { e.target.dataset.confirm = '1'; e.target.textContent = '한 번 더 누르면 삭제'; return; }
      studio.projects = studio.projects.filter((x) => x.id !== p.id); studio.current = null; saveStudio(); K.render();
    });
    $$('.step-fb input', root).forEach((inp) => inp.addEventListener('change', () => { p.inputs.fb = p.inputs.fb || {}; p.inputs.fb[inp.id.slice(3)] = inp.value; saveStudio(); }));
    let ctl = null;
    const status = $('#st-status', root);
    async function run(key) {
      const step = STEPS.find((s) => s.key === key);
      if (!ai.available) { status.textContent = unavailableMsg; return false; }
      const fb = (p.inputs.fb || {})[key];
      if (key === 'creative' && fb) p.inputs.pick = fb;
      const prompt = step.prompt(p) + (fb ? '\n\n[사용자 수정 지시] ' + fb : '') + (p.steps[key] ? '\n\n[이전 결과, 이번에는 개선해서 다시]\n' + JSON.stringify(p.steps[key]).slice(0, 6000) : '');
      ctl = new AbortController(); $('#st-stop', root).hidden = false;
      const li2 = $('[data-step="' + key + '"]', root); li2.classList.add('running');
      try {
        const r = await askJSON(prompt, { modelTier: key === 'concept' || key === 'review' ? 'complex' : 'default', signal: ctl.signal, cache: false }, status);
        p.steps[key] = r; saveStudio(); status.textContent = step.title + ' 완료'; return true;
      } catch (e) { status.textContent = step.title + ': ' + errText(e); return false; }
      finally { $('#st-stop', root).hidden = true; li2.classList.remove('running'); }
    }
    $$('[data-run]', root).forEach((b) => b.addEventListener('click', async () => { b.disabled = true; if (await run(b.dataset.run)) K.render(); else b.disabled = false; }));
    $('#st-stop', root).addEventListener('click', () => ctl && ctl.abort());
    $('#st-all', root).addEventListener('click', async (e) => {
      e.target.disabled = true;
      for (const s of STEPS) { if (p.steps[s.key]) continue; if (!(await run(s.key))) break; }
      K.render();
    });
    $('#st-doc', root).addEventListener('click', () => K.copy(docMarkdown(p), '기획서'));
    $('#st-dl', root).addEventListener('click', () => K.live && K.live.download ? K.live.download((p.name || 'concept').replace(/[\\/:*?"<>|]/g, '') + '.md', docMarkdown(p)) : K.copy(docMarkdown(p), '기획서'));
  }
  K.VIEWS.studio = renderStudio;

})();
