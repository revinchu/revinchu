// 프레젠테이션 문서 모델 — DOM 없음 (Node 단위 테스트 대상)
//
// pres = { version, size: {w, h}(px, 96dpi), theme, slides, media: {id: dataURL}, props, footer }
// slide = { id, layout, bg, hidden, transition, notes, objects, anims, bgObjects?, section? }
// 개체 = { id, type: 'shape'|'image'|'table'|'chart', x, y, w, h, rot, flipH, flipV, name, ph?, grp?, ... }
// 글 = { paras: [{ runs: [{ t, b, i, u, s, size(pt), color, font, hl }], align, lvl, bullet, lineSpacing, spcBef, spcAft }], anchor, insets, wrap, autofit }
import { THEMES, cloneTheme, themeByName } from './themes.js';

export const PX_PER_PT = 96 / 72;
export const EMU_PER_PX = 9525;
export const SIZES = {
  wide: { w: 1280, h: 720, label: '와이드스크린 (16:9)' },
  standard: { w: 960, h: 720, label: '표준 (4:3)' },
  a4: { w: 1040, h: 720, label: 'A4 용지 (210x297mm)' },
};

let seq = 0;
export const uid = (p = 'o') => `${p}${Date.now().toString(36)}${(seq++).toString(36)}${Math.random().toString(36).slice(2, 5)}`;
const clone = (v) => JSON.parse(JSON.stringify(v));

// ───────────── 글 ─────────────
export const run = (t, props = {}) => ({ t, ...props });
export const para = (text = '', props = {}, runProps = {}) => ({ runs: text === '' ? [] : [run(text, runProps)], align: 'l', lvl: 0, ...props });
export function textBody(paras = [para()], props = {}) {
  return { paras, anchor: 't', insets: [9.6, 4.8, 9.6, 4.8], wrap: true, autofit: 'none', ...props };
}
/** 글 상자 내용 → 일반 텍스트 (단락은 줄 바꿈) */
export function plainText(body) {
  if (!body) return '';
  return body.paras.map((p) => p.runs.map((r) => r.t).join('')).join('\n');
}
export const isEmptyText = (body) => !body || body.paras.every((p) => p.runs.every((r) => r.t === ''));

/** 일반 텍스트 → 단락 (첫 단락의 서식을 이어받음) */
export function setPlainText(body, text) {
  const tpl = body.paras[0] ?? para();
  const r0 = tpl.runs[0] ?? {};
  const { t: _t, ...rp } = r0;
  body.paras = String(text).split('\n').map((line) => ({ ...clone({ ...tpl, runs: [] }), runs: line ? [run(line, clone(rp))] : [] }));
  return body;
}

/** 글 전체 또는 모든 단락에 글자 서식 적용 (값이 undefined 면 지움) */
export function applyRunProps(body, props) {
  for (const p of body.paras) {
    for (const r of p.runs) for (const [k, v] of Object.entries(props)) { if (v === undefined || v === null) delete r[k]; else r[k] = v; }
    // 빈 단락의 글자 크기 등도 기억 (나중에 입력할 글자에 적용)
    p.end = { ...(p.end ?? {}) };
    for (const [k, v] of Object.entries(props)) { if (v === undefined || v === null) delete p.end[k]; else p.end[k] = v; }
  }
  return body;
}
export function applyParaProps(body, props) {
  for (const p of body.paras) for (const [k, v] of Object.entries(props)) { if (v === undefined || v === null) delete p[k]; else p[k] = v; }
  return body;
}
/** 글자 서식 공통값 (모든 글자가 같으면 그 값, 아니면 null) */
export function commonRunProp(body, key, dflt) {
  let val;
  let first = true;
  for (const p of body?.paras ?? []) {
    for (const r of p.runs) {
      if (!r.t) continue;
      const v = r[key] ?? dflt?.(p) ?? null;
      if (first) { val = v; first = false; } else if (v !== val) return null;
    }
  }
  return first ? (body?.paras[0]?.end?.[key] ?? dflt?.(body?.paras[0] ?? { lvl: 0 }) ?? null) : val;
}
/** 인접한 같은 서식의 글자 묶음 합치기 */
export function normalizeRuns(body) {
  for (const p of body.paras) {
    const out = [];
    for (const r of p.runs) {
      if (r.t === '') continue;
      const prev = out[out.length - 1];
      if (prev && sameRunProps(prev, r)) prev.t += r.t; else out.push({ ...r });
    }
    p.runs = out;
  }
  return body;
}
function sameRunProps(a, b) {
  const ka = Object.keys(a).filter((k) => k !== 't');
  const kb = Object.keys(b).filter((k) => k !== 't');
  return ka.length === kb.length && ka.every((k) => a[k] === b[k]);
}

// ───────────── 개체 기본값 ─────────────
const BASE = (type, rect, extra = {}) => ({ id: uid(), type, x: Math.round(rect.x), y: Math.round(rect.y), w: Math.round(rect.w), h: Math.round(rect.h), rot: 0, ...extra });

/** 도형 (PowerPoint 기본 도형 스타일: 강조 1 채우기, 진한 윤곽선, 흰 글자 가운데) */
export function newShape(kind, rect) {
  const line = kind === 'line' || kind === 'straightConnector1' || kind === 'arc';
  return BASE('shape', rect, {
    shape: kind,
    fill: line ? null : { type: 'solid', color: '@accent1' },
    line: { color: line ? '@accent1' : '@accent1:s75', width: line ? 1.5 : 1.33, dash: 'solid', ...(kind === 'straightConnector1' ? { tail: 'triangle' } : {}) },
    text: line ? null : textBody([para('', { align: 'ctr' })], { anchor: 'ctr', defColor: '@lt1' }),
  });
}
export function newTextBox(rect, text = '') {
  return BASE('shape', rect, {
    shape: 'rect', txBox: true, fill: null, line: null,
    text: textBody([para(text, {}, {})], { autofit: 'resize', wrap: true }),
  });
}
export function newImage(media, rect, extra = {}) {
  return BASE('image', rect, { media, ...extra });
}
export function newTable(rows, cols, rect) {
  const cw = rect.w / cols;
  const rh = Math.max(30, Math.min(48, rect.h / rows));
  return BASE('table', { ...rect, h: rh * rows }, {
    cols: Array.from({ length: cols }, () => cw),
    rows: Array.from({ length: rows }, () => ({ h: rh, cells: Array.from({ length: cols }, () => ({ text: textBody([para()], { anchor: 'ctr', insets: [7, 3.5, 7, 3.5] }) })) })),
    style: { firstRow: true, banded: true, accent: 'accent1', lastRow: false, firstCol: false },
  });
}
export function newChart(rect, chart) {
  return BASE('chart', rect, {
    chart: chart ?? {
      kind: 'col', title: '차트 제목', legend: 'b', labels: false,
      cats: ['항목 1', '항목 2', '항목 3', '항목 4'],
      series: [{ name: '계열 1', vals: [4.3, 2.5, 3.5, 4.5] }, { name: '계열 2', vals: [2.4, 4.4, 1.8, 2.8] }, { name: '계열 3', vals: [2, 2, 3, 5] }],
    },
  });
}

// ───────────── 레이아웃 (와이드 1280×720 기준, 다른 크기는 가로 비율로) ─────────────
export const LAYOUTS = [
  ['title', '제목 슬라이드'], ['titleContent', '제목 및 내용'], ['section', '구역 머리글'], ['twoContent', '콘텐츠 2개'],
  ['comparison', '비교'], ['titleOnly', '제목만'], ['blank', '빈 화면'], ['contentCaption', '캡션 있는 콘텐츠'], ['pictureCaption', '캡션 있는 그림'],
];
export const LAYOUT_LABEL = Object.fromEntries(LAYOUTS);

const PH_TEXT = {
  ctrTitle: '제목을 입력하십시오', title: '제목을 입력하십시오', subTitle: '부제목을 입력하십시오',
  body: '텍스트를 입력하십시오', pic: '그림을 추가하려면 아이콘을 클릭하십시오',
};
export const placeholderPrompt = (o) => o.prompt ?? PH_TEXT[o.ph] ?? PH_TEXT.body;

/** 레이아웃의 개체 틀 정의: [ph, x, y, w, h, 글 기본값] */
function layoutBoxes(layout) {
  const T = ['title', 88, 38, 1104, 139, { anchor: 'ctr' }];
  switch (layout) {
    case 'title': return [['ctrTitle', 160, 118, 960, 250, { anchor: 'b', align: 'ctr' }], ['subTitle', 160, 378, 960, 174, { anchor: 't', align: 'ctr' }]];
    case 'titleContent': return [T, ['body', 88, 192, 1104, 456, {}]];
    case 'section': return [['title', 87, 180, 1104, 300, { anchor: 'b' }], ['body', 87, 482, 1104, 158, { noBullet: true, sub: true }]];
    case 'twoContent': return [T, ['body', 88, 192, 544, 456, {}], ['body', 648, 192, 544, 456, {}]];
    case 'comparison': return [T, ['body', 88, 176, 541, 89, { noBullet: true, head: true, anchor: 'b' }], ['body', 88, 265, 541, 383, {}], ['body', 648, 176, 544, 89, { noBullet: true, head: true, anchor: 'b' }], ['body', 648, 265, 544, 383, {}]];
    case 'titleOnly': return [T];
    case 'contentCaption': return [['title', 88, 48, 413, 168, { anchor: 'b', small: true }], ['body', 544, 104, 648, 516, {}], ['body', 88, 216, 413, 404, { noBullet: true, caption: true }]];
    case 'pictureCaption': return [['title', 88, 48, 413, 168, { anchor: 'b', small: true }], ['pic', 544, 104, 648, 516, {}], ['body', 88, 216, 413, 404, { noBullet: true, caption: true }]];
    default: return [];
  }
}

/** 개체 틀 기본 글꼴 크기 (pt) */
export function defaultSize(o, lvl = 0) {
  const k = o.phKind ?? {};
  if (o.ph === 'ctrTitle') return 54;
  if (o.ph === 'title') return k.small ? 32 : (o.phLayout === 'section' ? 54 : 40);
  if (o.ph === 'subTitle') return 24;
  if (o.ph === 'body' || o.ph === 'obj') {
    if (k.sub) return 24;
    if (k.head) return 24;
    if (k.caption) return 16;
    return [28, 24, 20, 18, 18][Math.min(lvl, 4)];
  }
  if (o.type === 'table') return 18;
  return 18;
}
/** 기본 글자 색 */
export function defaultColor(o) {
  if (o.text?.defColor) return o.text.defColor;
  if (o.ph === 'subTitle' || (o.ph === 'body' && o.phKind?.sub)) return '@tx1:lm65:lo35';
  return '@tx1';
}
export const defaultFont = (o) => (o.ph === 'title' || o.ph === 'ctrTitle' ? '+mj' : '+mn');

export function scaleRect(size, x, y, w, h) {
  const sx = size.w / 1280;
  const sy = size.h / 720;
  return { x: x * sx, y: y * sy, w: w * sx, h: h * sy };
}

/** 레이아웃의 빈 개체 틀 만들기 */
export function layoutPlaceholders(layout, size) {
  return layoutBoxes(layout).map(([ph, x, y, w, h, k]) => {
    const rect = scaleRect(size, x, y, w, h);
    if (ph === 'pic') return BASE('image', rect, { ph: 'pic', media: null, phLayout: layout });
    const bullet = ph === 'body' && !k.noBullet ? { type: 'char', char: '•' } : null;
    const o = BASE('shape', rect, {
      shape: 'rect', fill: null, line: null, ph, phLayout: layout, phKind: { ...k },
      text: textBody([para('', { align: k.align ?? 'l', ...(bullet ? { bullet } : {}), ...(ph === 'title' || ph === 'ctrTitle' ? { lineSpacing: 0.9 } : {}) }, k.head ? { b: true } : {})], { anchor: k.anchor ?? 't', autofit: 'shrink' }),
    });
    if (k.align) o.text.paras[0].align = k.align;
    return o;
  });
}

// ───────────── 슬라이드 · 문서 ─────────────
export function newSlide(pres, layout = 'titleContent') {
  return { id: uid('s'), layout, bg: null, hidden: false, transition: null, notes: '', objects: layoutPlaceholders(layout, pres.size), anims: [] };
}

export function newPresentation({ theme = 'office', size = 'wide', firstLayout = 'title' } = {}) {
  const t = cloneTheme(themeByName(theme) ?? THEMES[0]);
  const pres = {
    version: 1,
    size: { w: SIZES[size]?.w ?? 1280, h: SIZES[size]?.h ?? 720 },
    theme: t,
    slides: [],
    media: {},
    props: { title: '', author: '', created: new Date().toISOString() },
    footer: { slideNum: false, date: false, text: '', hideOnTitle: true },
  };
  if (firstLayout) pres.slides.push(newSlide(pres, firstLayout));
  return pres;
}

/** 새 슬라이드 위치에 맞는 다음 레이아웃 (제목 슬라이드 다음은 제목 및 내용) */
export const nextLayout = (layout) => (layout === 'title' || layout === 'section' ? 'titleContent' : layout ?? 'titleContent');

/** 슬라이드 복제 (개체 · 애니메이션 id 새로) */
export function duplicateSlide(slide) {
  const s = clone(slide);
  s.id = uid('s');
  const map = new Map();
  for (const o of s.objects) { const nid = uid(); map.set(o.id, nid); o.id = nid; }
  const gmap = new Map();
  for (const o of s.objects) if (o.grp) { if (!gmap.has(o.grp)) gmap.set(o.grp, uid('g')); o.grp = gmap.get(o.grp); }
  s.anims = (s.anims ?? []).map((a) => ({ ...a, id: uid('a'), obj: map.get(a.obj) ?? a.obj }));
  delete s.section;
  return s;
}

/** 개체 복제 (붙여넣기 · Ctrl+D) — 같은 그룹은 새 그룹으로 */
export function cloneObjects(objs, dx = 0, dy = 0) {
  const gmap = new Map();
  return objs.map((o) => {
    const c = clone(o);
    c.id = uid();
    c.x += dx; c.y += dy;
    if (c.grp) { if (!gmap.has(c.grp)) gmap.set(c.grp, uid('g')); c.grp = gmap.get(c.grp); }
    delete c.ph; delete c.phKind; delete c.phLayout;
    return c;
  });
}

/** 레이아웃 바꾸기: 같은 종류의 개체 틀은 내용을 옮기고 위치만 새 레이아웃으로 */
export function changeLayout(pres, slide, layout) {
  const fresh = layoutPlaceholders(layout, pres.size);
  const old = slide.objects.filter((o) => o.ph);
  const rest = slide.objects.filter((o) => !o.ph);
  const used = new Set();
  const take = (kinds) => old.find((o) => !used.has(o) && kinds.includes(o.ph));
  const out = fresh.map((f) => {
    const kinds = f.ph === 'ctrTitle' || f.ph === 'title' ? ['title', 'ctrTitle'] : f.ph === 'subTitle' ? ['subTitle', 'body'] : f.ph === 'pic' ? ['pic'] : ['body', 'subTitle', 'obj'];
    const src = take(kinds);
    if (!src) return f;
    used.add(src);
    if (f.ph === 'pic') return { ...f, media: src.media };
    const text = clone(src.text);
    text.anchor = f.text.anchor;
    if (f.ph === 'body' && src.ph === 'body') return { ...f, text, id: src.id };
    for (const p of text.paras) { p.bullet = f.text.paras[0].bullet; if (f.text.paras[0].align) p.align = f.text.paras[0].align; }
    return { ...f, text, id: src.id };
  });
  // 새 레이아웃에 자리가 없는 내용 있는 개체 틀은 일반 개체로 남김
  for (const o of old) if (!used.has(o) && (o.type === 'image' ? o.media : !isEmptyText(o.text))) { const c = { ...o }; c.phOrphan = true; out.push(c); }
  slide.objects = [...out, ...rest];
  slide.layout = layout;
  const ids = new Set(slide.objects.map((o) => o.id));
  slide.anims = (slide.anims ?? []).filter((a) => ids.has(a.obj));
}

/** 레이아웃 다시 설정: 개체 틀 위치 · 크기를 레이아웃 기본으로 */
export function resetSlideLayout(pres, slide) {
  const fresh = layoutPlaceholders(slide.layout, pres.size);
  const byPh = new Map();
  for (const f of fresh) { const k = f.ph === 'ctrTitle' ? 'title' : f.ph; if (!byPh.has(k)) byPh.set(k, []); byPh.get(k).push(f); }
  for (const o of slide.objects) {
    if (!o.ph) continue;
    const k = o.ph === 'ctrTitle' ? 'title' : o.ph;
    const f = byPh.get(k)?.shift();
    if (!f) continue;
    Object.assign(o, { x: f.x, y: f.y, w: f.w, h: f.h, rot: 0 });
  }
}

/** 슬라이드 크기 바꾸기 (개체 위치 · 크기를 비율대로) */
export function resizePresentation(pres, w, h, mode = 'fit') {
  const sx = w / pres.size.w;
  const sy = h / pres.size.h;
  const s = mode === 'fit' ? Math.min(sx, sy) : Math.max(sx, sy);
  const ox = (w - pres.size.w * s) / 2;
  const oy = (h - pres.size.h * s) / 2;
  const fix = (o) => {
    o.x = o.x * s + ox; o.y = o.y * s + oy; o.w *= s; o.h *= s;
    if (o.type === 'table') { o.cols = o.cols.map((c) => c * s); for (const r of o.rows) r.h *= s; }
    const scaleText = (t) => { for (const p of t?.paras ?? []) for (const r of p.runs) if (r.size) r.size = Math.round(r.size * s * 10) / 10; };
    if (s < 0.999) { scaleText(o.text); if (o.type === 'table') for (const r of o.rows) for (const c of r.cells) scaleText(c.text); }
  };
  for (const sl of pres.slides) { for (const o of sl.objects) fix(o); for (const o of sl.bgObjects ?? []) fix(o); }
  pres.size = { w: Math.round(w), h: Math.round(h) };
}

// ───────────── 그룹 ─────────────
export function groupMembers(slide, o) {
  return o?.grp ? slide.objects.filter((x) => x.grp === o.grp) : o ? [o] : [];
}
export function bbox(objs) {
  if (!objs.length) return null;
  let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
  for (const o of objs) {
    const [a, b, c, d] = rotatedBox(o);
    x1 = Math.min(x1, a); y1 = Math.min(y1, b); x2 = Math.max(x2, c); y2 = Math.max(y2, d);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}
/** 회전한 개체를 감싸는 상자 [x1, y1, x2, y2] */
export function rotatedBox(o) {
  if (!o.rot) return [o.x, o.y, o.x + o.w, o.y + o.h];
  const a = (o.rot * Math.PI) / 180;
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const hw = (Math.abs(Math.cos(a)) * o.w + Math.abs(Math.sin(a)) * o.h) / 2;
  const hh = (Math.abs(Math.sin(a)) * o.w + Math.abs(Math.cos(a)) * o.h) / 2;
  return [cx - hw, cy - hh, cx + hw, cy + hh];
}

// ───────────── 맞춤 · 배분 ─────────────
/** how: l|c|r|t|m|b|dh|dv, toSlide: 슬라이드 기준 */
export function alignObjects(objs, how, size, toSlide = false) {
  if (!objs.length) return;
  const box = toSlide || objs.length === 1 ? { x: 0, y: 0, w: size.w, h: size.h } : bbox(objs);
  if (how === 'dh' || how === 'dv') {
    const horiz = how === 'dh';
    const list = [...objs].sort((a, b) => (horiz ? a.x - b.x : a.y - b.y));
    if (list.length < 2) return;
    const total = list.reduce((s, o) => s + (horiz ? o.w : o.h), 0);
    const span = toSlide || objs.length < 3 ? (horiz ? box.w : box.h) : (horiz ? bbox(list).w : bbox(list).h);
    const start = toSlide || objs.length < 3 ? (horiz ? box.x : box.y) : (horiz ? bbox(list).x : bbox(list).y);
    const gap = (span - total) / (list.length - 1);
    let p = start;
    for (const o of list) { if (horiz) o.x = p; else o.y = p; p += (horiz ? o.w : o.h) + gap; }
    return;
  }
  for (const o of objs) {
    if (how === 'l') o.x = box.x;
    if (how === 'c') o.x = box.x + (box.w - o.w) / 2;
    if (how === 'r') o.x = box.x + box.w - o.w;
    if (how === 't') o.y = box.y;
    if (how === 'm') o.y = box.y + (box.h - o.h) / 2;
    if (how === 'b') o.y = box.y + box.h - o.h;
  }
}

/** 순서: front | back | forward | backward (그룹은 함께) */
export function reorder(slide, ids, how) {
  const set = new Set(ids);
  const objs = slide.objects;
  const sel = objs.filter((o) => set.has(o.id));
  const rest = objs.filter((o) => !set.has(o.id));
  if (how === 'front') slide.objects = [...rest, ...sel];
  else if (how === 'back') slide.objects = [...sel, ...rest];
  else if (how === 'forward' || how === 'backward') {
    const arr = [...objs];
    const idx = arr.map((o, i) => (set.has(o.id) ? i : -1)).filter((i) => i >= 0);
    if (how === 'forward') {
      for (let k = idx.length - 1; k >= 0; k--) { const i = idx[k]; if (i < arr.length - 1 && !set.has(arr[i + 1].id)) [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]]; }
    } else {
      for (const i of idx) if (i > 0 && !set.has(arr[i - 1].id)) [arr[i], arr[i - 1]] = [arr[i - 1], arr[i]];
    }
    slide.objects = arr;
  }
}

// ───────────── 찾기 · 바꾸기 ─────────────
function textBodies(slide) {
  const out = [];
  for (const o of slide.objects) {
    if (o.text) out.push({ o, body: o.text });
    if (o.type === 'table') for (const [ri, r] of o.rows.entries()) for (const [ci, c] of r.cells.entries()) if (c.text) out.push({ o, body: c.text, cell: [ri, ci] });
  }
  return out;
}
/** 일치 목록: [{ slide(index), obj, cell, para, start, len }] */
export function findAll(pres, query, { matchCase = false, wholeWord = false } = {}) {
  if (!query) return [];
  const esc = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(wholeWord ? `(?<![\\p{L}\\p{N}_])${esc}(?![\\p{L}\\p{N}_])` : esc, `gu${matchCase ? '' : 'i'}`);
  const out = [];
  pres.slides.forEach((s, si) => {
    for (const { o, body, cell } of textBodies(s)) {
      body.paras.forEach((p, pi) => {
        const txt = p.runs.map((r) => r.t).join('');
        for (const m of txt.matchAll(re)) out.push({ slide: si, obj: o.id, cell, para: pi, start: m.index, len: m[0].length });
      });
    }
    // 메모
    if (s.notes) for (const m of s.notes.matchAll(re)) out.push({ slide: si, notes: true, start: m.index, len: m[0].length });
  });
  return out;
}
/** 단락의 [start, start+len) 를 글자로 바꿈 (첫 글자의 서식 유지) */
export function replaceInPara(p, start, len, text) {
  let pos = 0;
  let done = false;
  const out = [];
  for (const r of p.runs) {
    const a = pos;
    const b = pos + r.t.length;
    pos = b;
    if (b <= start || a >= start + len) { out.push(r); continue; }
    const before = r.t.slice(0, Math.max(0, start - a));
    const after = r.t.slice(Math.min(r.t.length, start + len - a));
    if (before) out.push({ ...r, t: before });
    if (!done) { out.push({ ...r, t: text }); done = true; }
    if (after) out.push({ ...r, t: after });
  }
  p.runs = out.filter((r) => r.t !== '');
}
export function replaceAll(pres, query, repl, opts = {}) {
  const hits = findAll(pres, query, opts);
  // 뒤에서부터 바꿔야 앞 위치가 유지됨
  for (const h of [...hits].reverse()) {
    const s = pres.slides[h.slide];
    if (h.notes) { s.notes = s.notes.slice(0, h.start) + repl + s.notes.slice(h.start + h.len); continue; }
    const o = s.objects.find((x) => x.id === h.obj);
    const body = h.cell ? o.rows[h.cell[0]].cells[h.cell[1]].text : o.text;
    replaceInPara(body.paras[h.para], h.start, h.len, repl);
  }
  return hits.length;
}

// ───────────── 개요 (보기 › 개요) ─────────────
export function slideTitle(slide) {
  const t = slide.objects.find((o) => o.ph === 'title' || o.ph === 'ctrTitle');
  return t ? plainText(t.text).split('\n')[0] : '';
}
export function outline(pres) {
  return pres.slides.map((s, i) => ({
    index: i, title: slideTitle(s),
    body: s.objects.filter((o) => o.ph && o.ph !== 'title' && o.ph !== 'ctrTitle' && o.text).flatMap((o) => o.text.paras.map((p) => ({ lvl: p.lvl ?? 0, text: p.runs.map((r) => r.t).join('') }))).filter((x) => x.text),
  }));
}

// ───────────── 애니메이션 ─────────────
export const ANIM_EFFECTS = {
  entr: [['appear', '나타내기'], ['fade', '밝기 변화'], ['fly', '날아오기'], ['float', '떠오르기'], ['split', '내밀기'], ['wipe', '닦아내기'], ['zoom', '확대/축소'], ['grow', '확대/축소 회전'], ['wheel', '시계 방향 회전'], ['bounce', '바운드']],
  emph: [['pulse', '흔들기'], ['spin', '회전'], ['growShrink', '크게/작게'], ['teeter', '시소'], ['flash', '깜박이기']],
  exit: [['disappear', '사라지기'], ['fade', '밝기 변화'], ['fly', '날아가기'], ['zoom', '확대/축소'], ['wipe', '닦아내기']],
};
export const ANIM_CLASS_LABEL = { entr: '나타내기', emph: '강조', exit: '끝내기' };

export function addAnim(slide, objId, cls, effect, opts = {}) {
  const a = { id: uid('a'), obj: objId, cls, effect, start: 'click', dur: effect === 'appear' || effect === 'disappear' ? 0 : 0.5, delay: 0, dir: 'b', ...opts };
  slide.anims = [...(slide.anims ?? []), a];
  return a;
}
/** 슬라이드 쇼 단계: [[anim...]] — 클릭 한 번에 실행할 묶음 (with/after 는 앞 묶음에 붙음) */
export function animSteps(slide) {
  const steps = [];
  for (const a of slide.anims ?? []) {
    if (!slide.objects.some((o) => o.id === a.obj)) continue;
    if (a.start === 'click' || !steps.length) steps.push([a]);
    else steps[steps.length - 1].push(a);
  }
  return steps;
}
/** 각 애니메이션의 시작 시각(초) — 한 묶음 안에서 with 는 앞과 동시, after 는 앞이 끝난 뒤 */
export function stepTimeline(step) {
  let t = 0;
  let prevStart = 0;
  let prevEnd = 0;
  return step.map((a, i) => {
    const start = i === 0 ? 0 : a.start === 'with' ? prevStart : prevEnd;
    const s = start + (a.delay ?? 0);
    const e = s + (a.dur ?? 0.5);
    prevStart = s; prevEnd = Math.max(prevEnd, e); t = Math.max(t, e);
    return { anim: a, at: s, end: e };
  });
}

export const TRANSITIONS = [
  ['none', '없음'], ['cut', '컷'], ['fade', '페이드'], ['push', '밀어내기'], ['wipe', '닦아내기'], ['split', '나누기'],
  ['cover', '덮기'], ['uncover', '나타내기'], ['zoom', '확대/축소'], ['circle', '원형'], ['dissolve', '디졸브'], ['flip', '뒤집기'],
];
export const TRANSITION_LABEL = Object.fromEntries(TRANSITIONS);

// ───────────── 실행 취소 ─────────────
/** 문서 상태(미디어 제외) 직렬화 — 미디어는 추가만 되므로 따로 둠 */
/** 실행 취소 · 자동 저장용 문서 글 (그림 · 포함된 글꼴은 커서 따로 보관) */
export const snapshot = (pres) => JSON.stringify({ ...pres, media: undefined, fonts: undefined });
export function restore(pres, snap) {
  const { media, fonts } = pres;
  const data = JSON.parse(snap);
  for (const k of Object.keys(pres)) delete pres[k];
  Object.assign(pres, data, { media }, fonts ? { fonts } : {});
}

export class History {
  constructor(limit = 100) { this.undo = []; this.redo = []; this.limit = limit; this.lastKey = null; this.lastAt = 0; }
  /** 바꾸기 전에 부름. key 가 같고 1.5초 안이면 앞 기록과 합침 (글자 입력 · 끌기 연속) */
  record(pres, key = null) {
    const now = Date.now();
    if (key && key === this.lastKey && now - this.lastAt < 1500) { this.lastAt = now; return; }
    this.undo.push(snapshot(pres));
    if (this.undo.length > this.limit) this.undo.shift();
    this.redo = [];
    this.lastKey = key; this.lastAt = now;
  }
  canUndo() { return this.undo.length > 0; }
  canRedo() { return this.redo.length > 0; }
  doUndo(pres) {
    if (!this.undo.length) return false;
    this.redo.push(snapshot(pres));
    restore(pres, this.undo.pop());
    this.lastKey = null;
    return true;
  }
  doRedo(pres) {
    if (!this.redo.length) return false;
    this.undo.push(snapshot(pres));
    restore(pres, this.redo.pop());
    this.lastKey = null;
    return true;
  }
  clear() { this.undo = []; this.redo = []; this.lastKey = null; }
}

/** 쓰지 않는 미디어 정리 (저장 전) */
export function usedMedia(pres) {
  const used = new Set();
  const walk = (o) => { if (o.media) used.add(o.media); };
  for (const s of pres.slides) {
    for (const o of s.objects) walk(o);
    for (const o of s.bgObjects ?? []) walk(o);
    if (s.bg?.media) used.add(s.bg.media);
    if (s.bg?.type === 'image') used.add(s.bg.media);
  }
  return used;
}
export function pruneMedia(pres) {
  const used = usedMedia(pres);
  for (const k of Object.keys(pres.media)) if (!used.has(k)) delete pres.media[k];
}

/** 문서 검사 (불러온 JSON 이 모양을 갖췄는지) */
export function validatePresentation(p) {
  if (!p || typeof p !== 'object' || !Array.isArray(p.slides) || !p.size || !p.theme) throw new Error('WIPOINT 문서 형식이 아닙니다');
  p.media ??= {};
  p.props ??= {};
  p.footer ??= { slideNum: false, date: false, text: '', hideOnTitle: true };
  for (const s of p.slides) { s.objects ??= []; s.anims ??= []; s.notes ??= ''; s.id ??= uid('s'); }
  return p;
}

// ───────────── 선택 영역 서식 (편집 중 일부 글자) ─────────────
/** 단락의 offset 위치에서 글자 묶음을 나눔 → 그 위치의 run 인덱스 */
function splitAt(p, off) {
  let pos = 0;
  for (let i = 0; i < p.runs.length; i++) {
    const r = p.runs[i];
    const end = pos + r.t.length;
    if (off === pos) return i;
    if (off < end) {
      p.runs.splice(i, 1, { ...r, t: r.t.slice(0, off - pos) }, { ...r, t: r.t.slice(off - pos) });
      return i + 1;
    }
    pos = end;
  }
  return p.runs.length;
}
/** a, b = { p: 단락, o: 글자 위치 }. 범위의 글자에 서식 적용 (빈 범위면 false) */
export function formatRange(body, a, b, props) {
  if (a.p > b.p || (a.p === b.p && a.o > b.o)) [a, b] = [b, a];
  let touched = false;
  for (let pi = a.p; pi <= b.p; pi++) {
    const p = body.paras[pi];
    if (!p) continue;
    const len = p.runs.reduce((s, r) => s + r.t.length, 0);
    const from = pi === a.p ? a.o : 0;
    const to = pi === b.p ? b.o : len;
    if (to <= from) {
      if (!len) { p.end = { ...(p.end ?? {}) }; for (const [k, v] of Object.entries(props)) { if (v == null) delete p.end[k]; else p.end[k] = v; } }
      continue;
    }
    const i = splitAt(p, from);
    const j = splitAt(p, to);
    for (let k = i; k < j; k++) {
      const r = p.runs[k];
      for (const [key, v] of Object.entries(props)) { if (v == null) delete r[key]; else r[key] = v; }
      touched = true;
    }
  }
  normalizeRuns(body);
  return touched;
}
/** 범위 안 글자의 서식 값 (모두 같으면 그 값, 아니면 null) */
export function rangeRunProp(body, a, b, key, dflt) {
  if (a.p > b.p || (a.p === b.p && a.o > b.o)) [a, b] = [b, a];
  let val;
  let first = true;
  for (let pi = a.p; pi <= b.p; pi++) {
    const p = body.paras[pi];
    if (!p) continue;
    let pos = 0;
    const from = pi === a.p ? a.o : 0;
    const to = pi === b.p ? b.o : Infinity;
    for (const r of p.runs) {
      const s = pos;
      const e = pos + r.t.length;
      pos = e;
      // 빈 범위(커서)면 커서 앞 글자의 서식
      const hit = from === to ? (s < from && from <= e) || (from === 0 && s === 0) : e > from && s < to;
      if (!hit) continue;
      const v = r[key] ?? dflt?.(p) ?? null;
      if (first) { val = v; first = false; } else if (v !== val) return null;
    }
  }
  if (first) { const p = body.paras[a.p]; return p?.end?.[key] ?? dflt?.(p ?? { lvl: 0 }) ?? null; }
  return val;
}
