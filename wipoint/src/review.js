// 사진 앨범 · 프레젠테이션 비교 · 접근성 검사 (DOM 없음)
import { newSlide, newImage, newTextBox, setPlainText, slideTitle, plainText, isEmptyText, defaultColor, defaultSize, uid } from './model.js';
import { resolveColor } from './themes.js';
import { slideBackground } from './render.js';

const clone = (v) => JSON.parse(JSON.stringify(v));

// ───────────── 사진 앨범 (PowerPoint [삽입 › 사진 앨범]) ─────────────
/** 그림 레이아웃: [키, 이름, 슬라이드당 그림 수, 제목 있음] */
export const ALBUM_LAYOUTS = [['fit', '슬라이드에 맞춤', 1, false], ['1', '그림 1개', 1, false], ['2', '그림 2개', 2, false], ['4', '그림 4개', 4, false], ['1t', '제목을 가진 그림 1개', 1, true], ['2t', '제목을 가진 그림 2개', 2, true], ['4t', '제목을 가진 그림 4개', 4, true]];
/** 프레임 모양 */
export const ALBUM_FRAMES = [['rect', '사각형'], ['roundRect', '둥근 사각형'], ['whiteFrame', '단순형 프레임, 흰색'], ['blackFrame', '단순형 프레임, 검은색'], ['doubleBlack', '복합형 프레임, 검은색'], ['shadow', '가운데 그림자 사각형'], ['soft', '부드러운 가장자리 사각형'], ['oval', '타원']];

function frameProps(frame) {
  switch (frame) {
    case 'roundRect': return { shape: 'roundRect', adj: [8000] };
    case 'whiteFrame': return { line: { color: '#FFFFFF', width: 8, dash: 'solid' }, fx: { shadow: { dist: 2, dir: 45, blur: 4, color: '#000000', alpha: 0.35 } } };
    case 'blackFrame': return { line: { color: '#000000', width: 6, dash: 'solid' } };
    case 'doubleBlack': return { line: { color: '#000000', width: 6, dash: 'solid', cmpd: 'dbl' } };
    case 'shadow': return { fx: { shadow: { dist: 0, dir: 0, blur: 12, color: '#000000', alpha: 0.5 } } };
    case 'soft': return { fx: { soft: 10 } };
    case 'oval': return { shape: 'ellipse' };
    default: return {};
  }
}

/**
 * 앨범 슬라이드 만들기
 * @param pres 프레젠테이션 (크기 · 레이아웃용)
 * @param pictures [{ media, w, h, name }] (w/h = 원본 픽셀 크기, 비율 계산용)
 * @param opts { layout, frame, captions, bw, title, subtitle }
 * @returns 슬라이드 목록 (첫 장 = 앨범 제목)
 */
export function albumSlides(pres, pictures, opts = {}) {
  const { w: W, h: H } = pres.size;
  const lay = ALBUM_LAYOUTS.find((l) => l[0] === opts.layout) ?? ALBUM_LAYOUTS[0];
  const [key, , per, titled] = lay;
  const out = [];
  const t = newSlide(pres, 'title');
  setPlainText(t.objects.find((o) => o.ph === 'ctrTitle').text, opts.title ?? '사진 앨범');
  setPlainText(t.objects.find((o) => o.ph === 'subTitle').text, opts.subtitle ?? '');
  t.album = true;
  out.push(t);
  for (let i = 0; i < pictures.length; i += per) {
    const group = pictures.slice(i, i + per);
    const s = newSlide(pres, titled ? 'titleOnly' : 'blank');
    s.album = true;
    if (titled) setPlainText(s.objects.find((o) => o.ph === 'title').text, group.map((p) => p.name).join(', '));
    // 그림이 들어갈 영역
    const top = titled ? H * 0.25 : key === 'fit' ? 0 : H * 0.06;
    const area = key === 'fit' ? { x: 0, y: 0, w: W, h: H } : { x: W * 0.06, y: top, w: W * 0.88, h: H - top - H * 0.05 };
    const cols = per === 1 ? 1 : 2;
    const rows = per === 4 ? 2 : 1;
    const capH = opts.captions ? Math.min(40, area.h * 0.08) : 0;
    const gap = per === 1 ? 0 : W * 0.02;
    const cw = (area.w - gap * (cols - 1)) / cols;
    const chh = (area.h - gap * (rows - 1)) / rows;
    group.forEach((p, k) => {
      const cx = area.x + (k % cols) * (cw + gap);
      const cy = area.y + Math.floor(k / cols) * (chh + gap);
      const boxH = chh - capH;
      const sc = Math.min(cw / (p.w || 1), boxH / (p.h || 1));
      const iw = (p.w || 1) * sc; const ih = (p.h || 1) * sc;
      const ix = cx + (cw - iw) / 2; const iy = cy + (boxH - ih) / 2;
      const img = newImage(p.media, { x: ix, y: iy, w: iw, h: ih }, { name: p.name, alt: p.name, ...frameProps(key === 'fit' ? 'rect' : opts.frame) });
      if (opts.bw) img.gray = true;
      s.objects.push(img);
      if (opts.captions) {
        const tb = newTextBox({ x: cx, y: iy + ih + 4, w: cw, h: capH }, p.name ?? '');
        tb.text.paras[0].align = 'ctr';
        tb.text.autofit = 'none';
        s.objects.push(tb);
      }
    });
    out.push(s);
  }
  return out;
}

// ───────────── 프레젠테이션 비교 (PowerPoint [검토 › 비교]) ─────────────
const objText = (o) => (o.text ? plainText(o.text) : o.type === 'table' ? o.rows.map((r) => r.cells.map((c) => plainText(c.text)).join('\t')).join('\n') : '');
const near = (a, b, tol = 2) => Math.abs(a - b) <= tol;
const sameBox = (a, b) => near(a.x, b.x) && near(a.y, b.y) && near(a.w, b.w) && near(a.h, b.h) && Math.round(a.rot ?? 0) === Math.round(b.rot ?? 0);
const AUTO_NAME = /^(Picture|TextBox|Rectangle|Shape|Placeholder|Title|Chart|Table|Video|Audio|Group|그림|텍스트 상자|직사각형|도형) ?\d*$/i;
const r1 = (v) => (typeof v === 'number' ? Math.round(v * 10) / 10 : v);
/** 서식 비교용 요약 (파일로 저장했다 읽어도 같은 값만) */
export const fmtOf = (o, theme) => JSON.stringify([
  o.fill?.type === 'solid' ? o.fill.color : o.fill?.type ?? null, o.line ? [o.line.color, r1(o.line.width)] : null, o.shape ?? 'rect', !!o.gray, o.fx ? Object.keys(o.fx).filter((k) => o.fx[k]).sort() : [],
  o.crop ? ['l', 't', 'r', 'b'].map((k) => Math.round((o.crop[k] ?? 0) * 100)) : null,
  // 글자: 실제로 보이는 값 (기본 크기 · 색을 파일에 적어 넣어도 같게)
  o.text?.paras.map((p) => [p.align ?? 'l', ...p.runs.filter((r) => r.t).map((r) => [!!r.b, !!r.i, !!r.u, r1(r.size ?? defaultSize(o, p.lvl ?? 0)), theme ? resolveColor(theme, r.color ?? defaultColor(o)) : r.color ?? null])]) ?? null,
]);
/** 다른 문서의 개체와 짝 점수: id → 개체 틀 → 같은 글 → 이름 → 같은 자리 */
function matchScore(a, b) {
  if (a.type !== b.type) return 0;
  let sc = 0;
  if (a.id === b.id) sc += 100;
  if (a.ph && a.ph === b.ph) sc += 50;
  const ta = objText(a);
  if (ta && ta === objText(b)) sc += 40;
  if (a.name && a.name === b.name && !AUTO_NAME.test(a.name)) sc += 30;
  if (sameBox(a, b)) sc += 25;
  else if (near(a.x + a.w / 2, b.x + b.w / 2, 40) && near(a.y + a.h / 2, b.y + b.h / 2, 40)) sc += 10;
  return sc;
}
/** 슬라이드 짝 찾기: id → 제목 → 순서 */
function matchSlides(a, b) {
  const pairs = new Map(); // b index → a index
  const usedA = new Set();
  b.slides.forEach((s, j) => { const i = a.slides.findIndex((x, k) => !usedA.has(k) && x.id === s.id); if (i >= 0) { pairs.set(j, i); usedA.add(i); } });
  b.slides.forEach((s, j) => {
    if (pairs.has(j)) return;
    const t = slideTitle(s).trim();
    if (!t) return;
    const i = a.slides.findIndex((x, k) => !usedA.has(k) && slideTitle(x).trim() === t);
    if (i >= 0) { pairs.set(j, i); usedA.add(i); }
  });
  b.slides.forEach((s, j) => { if (pairs.has(j)) return; if (j < a.slides.length && !usedA.has(j)) { pairs.set(j, j); usedA.add(j); } });
  return { pairs, usedA };
}

/**
 * 현재 문서(a)와 다른 문서(b)의 차이 — b 의 내용을 a 에 적용할 수 있는 변경 목록
 * 변경: { id, kind, slide (a 의 번호 또는 b 의 번호), label, detail, data }
 */
export function compareDecks(a, b) {
  const changes = [];
  const add = (kind, slide, label, detail, data) => changes.push({ id: `c${changes.length + 1}`, kind, slide, label, detail, data });
  const { pairs, usedA } = matchSlides(a, b);
  // 순서가 실제로 바뀐 슬라이드만 '이동' (다른 슬라이드 삽입/삭제로 번호만 밀린 것은 제외): 최장 증가 부분 수열 밖의 것
  const seq = [...pairs].sort((x, y) => x[0] - y[0]).map(([, i]) => i);
  const keep = new Set();
  {
    const tails = []; const prevIdx = new Array(seq.length); const tailIdx = [];
    seq.forEach((v, k) => {
      let lo = 0; let hi = tails.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (tails[mid] < v) lo = mid + 1; else hi = mid; }
      tails[lo] = v; tailIdx[lo] = k; prevIdx[k] = lo ? tailIdx[lo - 1] : -1;
    });
    for (let k = tailIdx[tails.length - 1] ?? -1; k >= 0; k = prevIdx[k]) keep.add(seq[k]);
  }
  b.slides.forEach((sb, j) => {
    if (!pairs.has(j)) { add('slideAdded', j, `슬라이드 ${j + 1} 삽입됨`, slideTitle(sb) || '(제목 없음)', { slide: sb, at: j }); return; }
    const i = pairs.get(j);
    const sa = a.slides[i];
    if (!keep.has(i)) add('slideMoved', i, `슬라이드 ${i + 1} 이동 (${i + 1} → ${j + 1})`, slideTitle(sa) || '', { from: i, to: j, sid: sa.id });
    const usedO = new Set();
    for (const ob of sb.objects) {
      if (ob.decor) continue;
      let oa = null; let best = 19;
      for (const x of sa.objects) { if (usedO.has(x.id) || x.decor) continue; const sc = matchScore(x, ob); if (sc > best) { best = sc; oa = x; } }
      if (!oa) { add('objAdded', i, `슬라이드 ${i + 1}: 개체 삽입됨`, objText(ob).slice(0, 60) || ob.name || ob.type, { sid: sa.id, obj: ob }); continue; }
      usedO.add(oa.id);
      if (objText(oa) !== objText(ob)) add('text', i, `슬라이드 ${i + 1}: 텍스트 변경`, `"${objText(oa).slice(0, 40)}" → "${objText(ob).slice(0, 40)}"`, { sid: sa.id, oid: oa.id, obj: ob });
      else if (!sameBox(oa, ob)) add('objMoved', i, `슬라이드 ${i + 1}: 위치/크기 변경`, oa.name || objText(oa).slice(0, 40) || oa.type, { sid: sa.id, oid: oa.id, obj: ob });
      else if (fmtOf(oa, a.theme) !== fmtOf(ob, a.theme)) add('format', i, `슬라이드 ${i + 1}: 서식 변경`, oa.name || objText(oa).slice(0, 40) || oa.type, { sid: sa.id, oid: oa.id, obj: ob });
    }
    for (const oa of sa.objects) if (!usedO.has(oa.id) && !oa.decor) add('objRemoved', i, `슬라이드 ${i + 1}: 개체 삭제됨`, objText(oa).slice(0, 60) || oa.name || oa.type, { sid: sa.id, oid: oa.id });
    if ((sa.notes ?? '') !== (sb.notes ?? '')) add('notes', i, `슬라이드 ${i + 1}: 메모 변경`, (sb.notes ?? '').slice(0, 60), { sid: sa.id, notes: sb.notes ?? '' });
    if (JSON.stringify(sa.transition ?? null) !== JSON.stringify(sb.transition ?? null)) add('transition', i, `슬라이드 ${i + 1}: 전환 변경`, sb.transition?.type ?? '없음', { sid: sa.id, transition: sb.transition ?? null });
  });
  a.slides.forEach((sa, i) => { if (!usedA.has(i)) add('slideRemoved', i, `슬라이드 ${i + 1} 삭제됨`, slideTitle(sa) || '(제목 없음)', { sid: sa.id }); });
  return changes;
}

/** 변경 하나를 문서 a 에 적용 (b 의 미디어는 mediaB 에서 복사) */
export function applyChange(a, ch, mediaB = {}) {
  const d = ch.data;
  const copyMedia = (o) => { const walk = (v) => { if (!v || typeof v !== 'object') return; for (const [k, x] of Object.entries(v)) { if (typeof x === 'string' && ['media', 'poster'].includes(k) && mediaB[x] && !a.media[x]) a.media[x] = mediaB[x]; else walk(x); } }; walk(o); };
  const slide = () => a.slides.find((s) => s.id === d.sid);
  switch (ch.kind) {
    case 'slideAdded': { const s = clone(d.slide); copyMedia(s); if (a.slides.some((x) => x.id === s.id)) s.id = uid('s'); a.slides.splice(Math.min(d.at, a.slides.length), 0, s); break; }
    case 'slideRemoved': a.slides = a.slides.filter((s) => s.id !== d.sid); break;
    case 'slideMoved': { const k = a.slides.findIndex((s) => s.id === d.sid); if (k >= 0) { const [s] = a.slides.splice(k, 1); a.slides.splice(Math.min(d.to, a.slides.length), 0, s); } break; }
    case 'objAdded': { const s = slide(); if (s) { const o = clone(d.obj); copyMedia(o); if (s.objects.some((x) => x.id === o.id)) o.id = uid(); s.objects.push(o); } break; }
    case 'objRemoved': { const s = slide(); if (s) { s.objects = s.objects.filter((o) => o.id !== d.oid); s.anims = (s.anims ?? []).filter((x) => x.obj !== d.oid); } break; }
    case 'text': case 'objMoved': case 'format': {
      const s = slide();
      const k = s?.objects.findIndex((o) => o.id === d.oid) ?? -1;
      if (k >= 0) { const o = clone(d.obj); copyMedia(o); o.id = d.oid; s.objects[k] = o; }
      break;
    }
    case 'notes': { const s = slide(); if (s) s.notes = d.notes; break; }
    case 'transition': { const s = slide(); if (s) s.transition = clone(d.transition); break; }
    default: break;
  }
}

// ───────────── 접근성 검사 (PowerPoint [검토 › 접근성 검사]) ─────────────
function lum(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!m) return null;
  const n = parseInt(m[1], 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => { const c = v / 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
}
/** WCAG 대비율 (1…21) */
export function contrastRatio(c1, c2) {
  const a = lum(c1); const b = lum(c2);
  if (a == null || b == null) return null;
  return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
}
const VAGUE_LINK = /^(여기|여기를 클릭|클릭|링크|자세히|더 보기|here|click here|link|more|read more)$/i;
/** 접근성 규칙 → 문제 목록 { level: 'error'|'warning'|'tip', rule, slide, obj, msg, fix } */
export function accessibilityIssues(pres) {
  const theme = pres.theme;
  const issues = [];
  const add = (level, rule, slide, obj, msg, fix = null) => issues.push({ level, rule, slide, obj, msg, fix });
  const titles = new Map();
  pres.slides.forEach((s, i) => {
    const title = slideTitle(s).trim();
    if (!title) add('error', '슬라이드 제목 없음', i, null, '화면 읽기 프로그램이 슬라이드를 구분할 수 있게 제목을 넣으세요. (제목을 화면에 안 보이게 하려면 슬라이드 밖에 두세요)', 'title');
    else titles.set(title, [...(titles.get(title) ?? []), i]);
    const bg = slideBackground(pres, s);
    const bgHex = bg?.type === 'solid' ? resolveColor(theme, bg.color) : null;
    // 읽기 순서: 제목이 맨 앞이 아니면
    const ti = s.objects.findIndex((o) => o.ph === 'title' || o.ph === 'ctrTitle');
    if (ti > 0 && s.objects.slice(0, ti).some((o) => !o.decor && (o.text && !isEmptyText(o.text) || o.type === 'image'))) add('warning', '읽기 순서 확인', i, s.objects[ti].id, '제목이 읽기 순서에서 첫 번째가 아닙니다. [선택 창]에서 순서를 확인하세요.', 'order');
    for (const o of s.objects) {
      if (o.decor || o.hidden) continue;
      const visual = o.type === 'image' && o.media || o.type === 'chart' || o.type === 'zoom' || (o.type === 'shape' && o.smartItems) || o.type === 'equation';
      if (visual && !o.alt && !o.decorative) add('error', '대체 텍스트 없음', i, o.id, `${o.type === 'chart' ? '차트' : o.type === 'image' ? '그림' : o.smartItems ? 'SmartArt' : '개체'}${o.name ? ` (${o.name})` : ''}에 설명을 넣거나 [장식용으로 표시]하세요.`, 'alt');
      if (o.type === 'media' && o.kind === 'video' && !o.alt) add('warning', '동영상 설명 · 자막 없음', i, o.id, '동영상에 대체 텍스트를 넣고, 말소리가 있으면 자막을 함께 제공하세요.', 'alt');
      if (o.type === 'table') {
        if (!o.style?.firstRow) add('error', '표 머리글 없음', i, o.id, '표의 첫 행을 머리글 행으로 지정하세요.', 'header');
        if (o.rows.some((r) => r.cells.some((c) => c.span > 1 || c.rowSpan > 1))) add('warning', '병합된 셀', i, o.id, '병합된 셀은 화면 읽기 프로그램이 읽기 어렵습니다. 가능하면 병합을 푸세요.');
      }
      // 글: 대비 · 크기 · 링크
      if (o.text && !isEmptyText(o.text)) {
        const fill = o.fill?.type === 'solid' ? resolveColor(theme, o.fill.color) : null;
        const back = fill ?? bgHex;
        let lowShown = false; let smallShown = false; let linkShown = false;
        for (const p of o.text.paras) for (const r of p.runs) {
          if (!r.t.trim()) continue;
          const size = (r.size ?? defaultSize(o, p.lvl ?? 0)) * (o.text.fontScale ?? 1);
          const fg = resolveColor(theme, r.color ?? defaultColor(o));
          const cr = back && fg ? contrastRatio(fg, back) : null;
          const need = size >= 18 || (r.b && size >= 14) ? 3 : 4.5;
          if (cr != null && cr < need && !lowShown) { add('warning', '읽기 어려운 텍스트 대비', i, o.id, `글자와 배경의 대비가 ${cr.toFixed(1)}:1 입니다 (권장 ${need}:1 이상). "${r.t.trim().slice(0, 20)}"`, 'contrast'); lowShown = true; }
          if (size < 12 && !smallShown) { add('tip', '작은 글꼴', i, o.id, `${Math.round(size)}pt 글자가 있습니다. 발표용은 18pt 이상을 권장합니다.`); smallShown = true; }
          if (r.link && VAGUE_LINK.test(r.t.trim()) && !linkShown) { add('warning', '의미 없는 하이퍼링크 텍스트', i, o.id, `"${r.t.trim()}" 대신 연결 대상을 알 수 있는 글을 쓰세요.`); linkShown = true; }
        }
      }
    }
  });
  for (const [t, list] of titles) if (list.length > 1) for (const i of list) add('warning', '중복된 슬라이드 제목', i, null, `"${t}" 제목이 슬라이드 ${list.map((k) => k + 1).join(', ')}에 같습니다. 구분되게 바꾸세요.`, 'title');
  const rank = { error: 0, warning: 1, tip: 2 };
  return issues.sort((a, b) => rank[a.level] - rank[b.level] || a.slide - b.slide);
}
