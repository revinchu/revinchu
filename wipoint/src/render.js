// 슬라이드 → HTML 문자열 (DOM 없음): 편집 화면 · 미리 보기 · 슬라이드 쇼 · 인쇄가 함께 씀
import { resolveColor, isDark } from './themes.js';
import { shapePath, customPath, textRect, OPEN_SHAPES } from './shapes.js';
import { PX_PER_PT, defaultSize, defaultColor, defaultFont, placeholderPrompt, isEmptyText } from './model.js';
import { chartSvg } from './chart.js';

export const escHtml = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const f = (n) => Math.round(n * 100) / 100;
const KO_FALLBACK = "'맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif";

export function fontStack(theme, font) {
  const name = font === '+mj' || font === '+mj-lt' ? theme.fonts.major : font === '+mn' || font === '+mn-lt' || !font ? theme.fonts.minor : font;
  return `'${String(name).replace(/'/g, '')}', ${KO_FALLBACK}`;
}

/** 채우기 → CSS 배경 */
export function fillCss(theme, fill, media = {}) {
  if (!fill) return 'transparent';
  if (fill.type === 'solid') return resolveColor(theme, fill.color, 'transparent');
  if (fill.type === 'gradient') {
    const stops = (fill.stops ?? []).map(([p, c]) => `${resolveColor(theme, c, '#fff')} ${f(p * 100)}%`).join(', ');
    return fill.path ? `radial-gradient(circle, ${stops})` : `linear-gradient(${f((fill.angle ?? 90) + 90)}deg, ${stops})`;
  }
  if (fill.type === 'image') {
    const src = media[fill.media] ?? fill.src;
    return src ? `url("${src}") center / ${fill.tile ? 'auto' : '100% 100%'} ${fill.tile ? 'repeat' : 'no-repeat'}` : 'transparent';
  }
  return 'transparent';
}

const DASH = { solid: '', dash: '4 3', sysDash: '3 1', dot: '1 2', sysDot: '1 1', dashDot: '4 3 1 3', lgDash: '8 3', lgDashDot: '8 3 1 3', lgDashDotDot: '8 3 1 3 1 3' };
export const DASH_LABEL = [['solid', '실선'], ['sysDot', '둥근 점선'], ['sysDash', '사각 점선'], ['dash', '파선'], ['dashDot', '파선-점선'], ['lgDash', '긴 파선'], ['lgDashDot', '긴 파선-점선']];

// ───────────── 글 ─────────────
const NUM_FMT = {
  arabicPeriod: (n) => `${n}.`, arabicParenR: (n) => `${n})`, arabicParenBoth: (n) => `(${n})`, arabicPlain: (n) => `${n}`,
  alphaLcPeriod: (n) => `${alpha(n)}.`, alphaUcPeriod: (n) => `${alpha(n).toUpperCase()}.`, alphaLcParenR: (n) => `${alpha(n)})`, alphaUcParenR: (n) => `${alpha(n).toUpperCase()})`,
  romanUcPeriod: (n) => `${roman(n)}.`, romanLcPeriod: (n) => `${roman(n).toLowerCase()}.`,
  circleNumDbPlain: (n) => (n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : `${n}.`),
  ganada: (n) => `${'가나다라마바사아자차카타파하'[(n - 1) % 14]}.`,
};
export const NUM_SCHEMES = [['arabicPeriod', '1. 2. 3.'], ['arabicParenR', '1) 2) 3)'], ['romanUcPeriod', 'I. II. III.'], ['alphaUcPeriod', 'A. B. C.'], ['alphaLcParenR', 'a) b) c)'], ['alphaLcPeriod', 'a. b. c.'], ['romanLcPeriod', 'i. ii. iii.'], ['circleNumDbPlain', '① ② ③'], ['ganada', '가. 나. 다.']];
export const BULLET_CHARS = [['•', '채워진 둥근 글머리 기호'], ['○', '속이 빈 둥근 글머리 기호'], ['■', '채워진 사각형 글머리 기호'], ['□', '속이 빈 사각형 글머리 기호'], ['◆', '별표 글머리 기호'], ['➢', '화살표 글머리 기호'], ['✓', '체크 표시 글머리 기호'], ['–', '대시']];
function alpha(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; }
function roman(n) {
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return s;
}
export const numberLabel = (scheme, n) => (NUM_FMT[scheme] ?? NUM_FMT.arabicPeriod)(n);

/** 단락의 실제 서식 (기본값 포함) */
export function paraDefaults(o, p) {
  const isBody = o.ph === 'body' || o.ph === 'obj';
  const isTitle = o.ph === 'title' || o.ph === 'ctrTitle';
  const lvl = p.lvl ?? 0;
  const bullet = p.bullet && p.bullet.type !== 'none' ? p.bullet : null;
  const step = 48;
  const hang = bullet?.type === 'num' ? 36 : 24;
  return {
    marL: p.marL ?? (bullet ? hang + lvl * step : lvl * step),
    indent: p.indent ?? (bullet ? -hang : 0),
    lineSpacing: p.lineSpacing ?? (isTitle || isBody ? 0.9 : 1),
    lineSpacingPt: p.lineSpacingPt ?? null,
    spcBef: p.spcBef ?? (isBody ? (lvl ? 5 : 10) : 0),
    spcAft: p.spcAft ?? 0,
    bullet,
  };
}

/** 글 상자 → HTML (편집기용이면 data-p/data-r 에 서식을 담음) */
export function textHtml(theme, o, body, { editable = false, prompt = false, scale = 1 } = {}) {
  const fsc = body.fontScale ?? 1;
  const lsr = body.lnSpcReduction ?? 0;
  const counters = [];
  const parts = [];
  body.paras.forEach((p, pi) => {
    const d = paraDefaults(o, p);
    const lvl = p.lvl ?? 0;
    // 번호 매기기: 같은 수준의 연속 번호 단락
    let label = '';
    if (d.bullet?.type === 'num') {
      counters.length = lvl + 1;
      counters[lvl] = (counters[lvl] ?? (d.bullet.start ?? 1) - 1) + 1;
      label = numberLabel(d.bullet.scheme, counters[lvl]);
    } else {
      counters.length = Math.min(counters.length, lvl);
      if (d.bullet?.type === 'char') label = d.bullet.char ?? '•';
      if (d.bullet?.type === 'pic') label = '■';
    }
    const end = p.end ?? {};
    const baseSize = (end.size ?? p.runs[0]?.size ?? defaultSize(o, lvl)) * fsc;
    const lh = d.lineSpacingPt ? `${f(d.lineSpacingPt * PX_PER_PT)}px` : f(1.2 * d.lineSpacing * (1 - lsr));
    const pStyle = [
      `text-align:${{ l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[p.align ?? 'l'] ?? 'left'}`,
      `padding-left:${f(Math.max(0, d.marL))}px`,
      `text-indent:${f(d.indent)}px`,
      `line-height:${lh}`,
      pi > 0 || d.spcBef ? `margin-top:${f(d.spcBef * PX_PER_PT * (pi === 0 ? 0 : 1))}px` : '',
      d.spcAft ? `margin-bottom:${f(d.spcAft * PX_PER_PT)}px` : '',
      `font-size:${f(baseSize * PX_PER_PT)}px`,
    ].filter(Boolean).join(';');
    const pAttr = editable ? ` data-p="${escHtml(JSON.stringify({ ...p, runs: undefined }))}"` : '';
    let inner = '';
    if (label && (p.runs.length || editable)) {
      const bc = d.bullet.color ? resolveColor(theme, d.bullet.color) : runColor(theme, o, p.runs[0] ?? end);
      const bsz = (d.bullet.size ?? 1) * baseSize;
      const bfont = d.bullet.font ? `font-family:${fontStack(theme, d.bullet.font)};` : '';
      inner += `<span class="bu" contenteditable="false" style="display:inline-block;min-width:${f(Math.max(0, -d.indent))}px;text-indent:0;color:${bc};font-size:${f(bsz * PX_PER_PT)}px;${bfont}">${escHtml(label)}</span>`;
    }
    if (!p.runs.length) {
      inner += prompt && pi === 0 ? `<span class="ph-prompt" data-prompt="1">${escHtml(placeholderPrompt(o))}</span>` : '<br>';
    } else {
      for (const r of p.runs) inner += runHtml(theme, o, p, r, fsc, editable);
    }
    parts.push(`<div class="p" style="${pStyle}"${pAttr}>${inner}</div>`);
  });
  return parts.join('');
}

function runColor(theme, o, r) {
  return resolveColor(theme, r.color ?? defaultColor(o), '#000');
}

function runHtml(theme, o, p, r, fsc, editable) {
  const size = (r.size ?? defaultSize(o, p.lvl ?? 0)) * fsc;
  const deco = [r.u ? 'underline' : '', r.s ? 'line-through' : ''].filter(Boolean).join(' ');
  const font = r.font ?? defaultFont(o);
  const st = [
    `font-size:${f(size * PX_PER_PT)}px`,
    `color:${runColor(theme, o, r)}`,
    `font-family:${fontStack(theme, font)}`,
    (r.b ?? (o.phKind?.head && r.b !== false)) ? 'font-weight:700' : '',
    r.i ? 'font-style:italic' : '',
    deco ? `text-decoration:${deco}` : '',
    r.hl ? `background:${resolveColor(theme, r.hl)}` : '',
    r.spc ? `letter-spacing:${f(r.spc * PX_PER_PT)}px` : '',
    r.base ? `vertical-align:${r.base > 0 ? 'super' : 'sub'};font-size:${f(size * PX_PER_PT * 0.65)}px` : '',
    r.cap === 'all' ? 'text-transform:uppercase' : r.cap === 'small' ? 'font-variant:small-caps' : '',
    r.shadow ? 'text-shadow:1px 1.5px 2px rgba(0,0,0,.45)' : '',
    r.outline ? `-webkit-text-stroke:${f(r.outline.width ?? 1)}px ${resolveColor(theme, r.outline.color ?? '@tx1')}` : '',
  ].filter(Boolean).join(';');
  const attr = editable ? ` data-r="${escHtml(JSON.stringify({ ...r, t: undefined }))}"` : '';
  const txt = escHtml(r.t).replace(/\v|\u000b/g, '<br>');
  if (r.link && !editable) return `<a class="lnk" href="${escHtml(r.link)}" target="_blank" rel="noopener" style="${st}">${txt}</a>`;
  return `<span style="${st}"${attr}>${txt}</span>`;
}

// ───────────── 개체 ─────────────
const nextId = (() => { let n = 0; return () => `g${(n++).toString(36)}`; })();

function svgFill(theme, fill, w, h, media, defs) {
  if (!fill) return 'none';
  if (fill.type === 'solid') return resolveColor(theme, fill.color, 'none');
  const id = nextId();
  if (fill.type === 'gradient') {
    const stops = (fill.stops ?? []).map(([p, c]) => { const col = resolveColor(theme, c, '#fff'); return `<stop offset="${f(p * 100)}%" stop-color="${col}"/>`; }).join('');
    if (fill.path) defs.push(`<radialGradient id="${id}">${stops}</radialGradient>`);
    else {
      const a = ((fill.angle ?? 90) * Math.PI) / 180;
      const x = Math.cos(a) / 2;
      const y = Math.sin(a) / 2;
      defs.push(`<linearGradient id="${id}" x1="${f(0.5 - x)}" y1="${f(0.5 - y)}" x2="${f(0.5 + x)}" y2="${f(0.5 + y)}">${stops}</linearGradient>`);
    }
    return `url(#${id})`;
  }
  if (fill.type === 'image') {
    const src = media[fill.media] ?? fill.src;
    if (!src) return 'none';
    defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${f(w)}" height="${f(h)}"><image href="${escHtml(src)}" width="${f(w)}" height="${f(h)}" preserveAspectRatio="none"/></pattern>`);
    return `url(#${id})`;
  }
  return 'none';
}

function markers(theme, line, defs) {
  const color = resolveColor(theme, line.color, '#000');
  const out = [];
  for (const end of ['head', 'tail']) {
    const kind = line[end];
    if (!kind || kind === 'none') continue;
    const id = nextId();
    const shape = kind === 'oval' ? '<circle cx="5" cy="5" r="4"/>' : kind === 'diamond' ? '<path d="M5 0 L10 5 L5 10 L0 5 Z"/>' : kind === 'arrow' ? '<path d="M0 0 L10 5 L0 10" fill="none" stroke-width="1.5" stroke="currentColor"/>' : '<path d="M0 0 L10 5 L0 10 Z"/>';
    defs.push(`<marker id="${id}" viewBox="0 0 10 10" refX="${kind === 'triangle' ? 8 : 5}" refY="5" markerWidth="4" markerHeight="4" orient="${end === 'head' ? 'auto-start-reverse' : 'auto'}" style="color:${color}" fill="${color}">${shape}</marker>`);
    out.push(`marker-${end === 'head' ? 'start' : 'end'}="url(#${id})"`);
  }
  return out.join(' ');
}

const EFFECT_SHADOW = 'drop-shadow(2px 3px 4px rgba(0,0,0,.35))';

export function shapeHtml(theme, o, media, opts = {}) {
  const defs = [];
  const w = Math.max(o.w, 0.01);
  const h = Math.max(o.h, 0.01);
  const d = o.path ? customPath(o.path, w, h) : shapePath(o.shape ?? 'rect', w, h, o.adj);
  const open = OPEN_SHAPES.has(o.shape) || o.openPath;
  const fill = open ? 'none' : svgFill(theme, o.fill, w, h, media, defs);
  const ln = o.line;
  const stroke = ln ? resolveColor(theme, ln.color, 'none') : 'none';
  const sw = ln ? Math.max(0.5, ln.width ?? 1) : 0;
  const dash = ln && DASH[ln.dash] ? `stroke-dasharray="${DASH[ln.dash].split(' ').map((x) => f(Number(x) * sw)).join(' ')}"` : '';
  const mk = ln ? markers(theme, ln, defs) : '';
  let html = '';
  const needSvg = fill !== 'none' || stroke !== 'none';
  if (needSvg) {
    html += `<svg class="geo" width="${f(w)}" height="${f(h)}" viewBox="0 0 ${f(w)} ${f(h)}" overflow="visible" style="position:absolute;left:0;top:0;${o.shadow ? `filter:${EFFECT_SHADOW}` : ''}">${defs.length ? `<defs>${defs.join('')}</defs>` : ''}<path d="${d}" fill="${fill}" fill-rule="evenodd" ${fill !== 'none' && o.fill?.alpha != null ? `fill-opacity="${o.fill.alpha}"` : ''} stroke="${stroke}" stroke-width="${f(sw)}" ${dash} stroke-linejoin="round" ${mk} vector-effect="non-scaling-stroke"/></svg>`;
  }
  if (o.text && (!isEmptyText(o.text) || opts.editable || (opts.prompt && o.ph))) html += textBoxHtml(theme, o, o.text, w, h, opts);
  return html;
}

/** 글 영역 (도형 모양 안쪽 + 여백) */
export function textBoxHtml(theme, o, body, w, h, opts = {}) {
  const [l, t, r, b] = o.txBox || o.ph ? [0, 0, w, h] : textRect(o.shape ?? 'rect', w, h);
  const ins = body.insets ?? [9.6, 4.8, 9.6, 4.8];
  const jc = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[body.anchor ?? 't'] ?? 'flex-start';
  const vert = body.vert === 'vert' || body.vert === 'eaVert' ? 'writing-mode:vertical-rl;' : body.vert === 'vert270' ? 'writing-mode:vertical-rl;transform:rotate(180deg);' : '';
  const prompt = opts.prompt && o.ph && isEmptyText(body);
  const cls = `tx${prompt ? ' empty-ph' : ''}`;
  const st = `position:absolute;left:${f(l + ins[0])}px;top:${f(t + ins[1])}px;width:${f(Math.max(1, r - l - ins[0] - ins[2]))}px;height:${f(Math.max(1, b - t - ins[1] - ins[3]))}px;display:flex;flex-direction:column;justify-content:${jc};${body.wrap === false ? 'white-space:pre;' : 'white-space:pre-wrap;overflow-wrap:break-word;word-break:keep-all;'}${vert}`;
  return `<div class="${cls}" data-tx="${escHtml(o.id)}" style="${st}"><div class="txi">${textHtml(theme, o, body, { editable: opts.editable, prompt })}</div></div>`;
}

function imageHtml(theme, o, media, opts) {
  const src = media[o.media] ?? o.src;
  if (!src) {
    if (!opts.prompt) return '';
    return `<div class="pic-ph" style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:8px;color:#7f7f7f;font-size:20px;text-align:center"><span class="pic-ph-ico">🖼</span><span>${escHtml(placeholderPrompt(o))}</span></div>`;
  }
  const c = o.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const iw = o.w / Math.max(0.01, 1 - c.l - c.r);
  const ih = o.h / Math.max(0.01, 1 - c.t - c.b);
  const clip = o.shape && o.shape !== 'rect' ? `clip-path:path('${shapePath(o.shape, o.w, o.h, o.adj)}');` : '';
  const ln = o.line;
  const border = ln ? `<svg width="${f(o.w)}" height="${f(o.h)}" style="position:absolute;left:0;top:0;overflow:visible"><path d="${shapePath(o.shape ?? 'rect', o.w, o.h, o.adj)}" fill="none" stroke="${resolveColor(theme, ln.color, '#000')}" stroke-width="${f(ln.width ?? 1)}"/></svg>` : '';
  const filt = [o.shadow ? EFFECT_SHADOW : '', o.gray ? 'grayscale(1)' : '', o.bright ? `brightness(${1 + o.bright})` : '', o.contrast ? `contrast(${1 + o.contrast})` : ''].filter(Boolean).join(' ');
  return `<div style="position:absolute;inset:0;overflow:hidden;${clip}${filt ? `filter:${filt};` : ''}"><img src="${escHtml(src)}" alt="${escHtml(o.alt ?? '')}" draggable="false" style="position:absolute;left:${f(-c.l * iw)}px;top:${f(-c.t * ih)}px;width:${f(iw)}px;height:${f(ih)}px;max-width:none;${o.alpha != null ? `opacity:${o.alpha};` : ''}"></div>${border}`;
}

/** 표 스타일 색 */
export function tableCellStyle(theme, o, ri, ci) {
  const st = o.style ?? {};
  const acc = st.accent ?? 'accent1';
  const nRows = o.rows.length;
  const nCols = o.cols.length;
  const header = st.firstRow && ri === 0;
  const total = st.lastRow && ri === nRows - 1;
  const firstCol = st.firstCol && ci === 0;
  const lastCol = st.lastCol && ci === nCols - 1;
  if (st.none) return { fill: null, color: '@tx1', bold: false, border: '@tx1' };
  if (header) return { fill: `@${acc}`, color: '@lt1', bold: true, border: '@lt1' };
  if (total) return { fill: `@${acc}`, color: '@lt1', bold: true, border: '@lt1' };
  const bodyIdx = ri - (st.firstRow ? 1 : 0);
  const band = st.banded && bodyIdx % 2 === 1;
  const light = st.light;
  return {
    fill: light ? (band ? `@${acc}:lm20:lo80` : null) : band ? `@${acc}:lm40:lo60` : `@${acc}:lm20:lo80`,
    color: '@tx1', bold: firstCol || lastCol, border: light ? `@${acc}` : '@lt1',
  };
}

function tableHtml(theme, o, opts) {
  const rows = o.rows.map((r, ri) => {
    let out = '';
    r.cells.forEach((c, ci) => {
      if (c.hMerge || c.vMerge) return;
      const ts = tableCellStyle(theme, o, ri, ci);
      const fill = c.fill !== undefined ? fillCss(theme, c.fill) : ts.fill ? resolveColor(theme, ts.fill) : 'transparent';
      const bcol = resolveColor(theme, ts.border, '#fff');
      const pseudo = { ...o, type: 'table', ph: null, text: c.text, phKind: null };
      const body = c.text;
      const defColor = ts.color;
      const b2 = { ...body, paras: body.paras.map((p) => ({ ...p, runs: p.runs.map((rr) => ({ ...rr, color: rr.color ?? defColor, b: rr.b ?? (ts.bold || undefined) })), end: { ...(p.end ?? {}), color: p.end?.color ?? defColor } })) };
      const ins = body.insets ?? [7, 3.5, 7, 3.5];
      const va = { t: 'top', ctr: 'middle', b: 'bottom' }[body.anchor ?? 't'];
      const borders = c.borders ? Object.entries(c.borders).map(([k, v]) => `border-${{ l: 'left', r: 'right', t: 'top', b: 'bottom' }[k]}:${f(v.width ?? 1)}px solid ${resolveColor(theme, v.color, '#000')}`).join(';') : `border:1px solid ${bcol}`;
      out += `<td${c.span ? ` colspan="${c.span}"` : ''}${c.rowSpan ? ` rowspan="${c.rowSpan}"` : ''} data-cell="${ri},${ci}" style="background:${fill};${borders};padding:${f(ins[1])}px ${f(ins[2])}px ${f(ins[3])}px ${f(ins[0])}px;vertical-align:${va};overflow:hidden"><div class="tx tcell" data-tx="${escHtml(o.id)}" data-cell="${ri},${ci}" style="white-space:pre-wrap;overflow-wrap:break-word;word-break:keep-all"><div class="txi">${textHtml(theme, pseudo, b2, { editable: opts.editable })}</div></div></td>`;
    });
    return `<tr style="height:${f(r.h)}px">${out}</tr>`;
  }).join('');
  const cols = o.cols.map((w) => `<col style="width:${f(w)}px">`).join('');
  return `<table class="tbl" style="position:absolute;left:0;top:0;width:${f(o.cols.reduce((a, b) => a + b, 0))}px;border-collapse:collapse;table-layout:fixed"><colgroup>${cols}</colgroup>${rows}</table>`;
}

/** 개체 하나 → 위치 · 회전이 들어간 div */
export function objectHtml(theme, o, media, opts = {}) {
  let inner;
  if (o.type === 'image') inner = imageHtml(theme, o, media, opts);
  else if (o.type === 'table') inner = tableHtml(theme, o, opts);
  else if (o.type === 'chart') inner = chartSvg(o.chart, o.w, o.h, theme, { font: theme.fonts.minor });
  else if (o.type === 'video') inner = `<div style="position:absolute;inset:0;background:#000;display:flex;align-items:center;justify-content:center;color:#fff;font-size:40px">▶</div>`;
  else inner = shapeHtml(theme, o, media, opts);
  const tr = [];
  if (o.rot) tr.push(`rotate(${f(o.rot)}deg)`);
  if (o.flipH || o.flipV) tr.push(`scale(${o.flipH ? -1 : 1},${o.flipV ? -1 : 1})`);
  const hidden = opts.hide?.has(o.id) ? 'visibility:hidden;' : '';
  const ph = o.ph && opts.prompt ? ' ph' : '';
  const link = o.link && opts.links ? ` data-link="${escHtml(o.link)}"` : '';
  return `<div class="ob${ph}" data-id="${escHtml(o.id)}"${link} style="left:${f(o.x)}px;top:${f(o.y)}px;width:${f(o.w)}px;height:${f(o.h)}px;${tr.length ? `transform:${tr.join(' ')};` : ''}${hidden}">${inner}</div>`;
}

// ───────────── 테마 장식 · 바닥글 ─────────────
/** 테마별 배경 장식 (마스터 도형에 해당) */
export function themeDecor(theme, layout, size) {
  const W = size.w;
  const H = size.h;
  const sx = W / 1280;
  const out = [];
  const R = (id, x, y, w, h, color, shape = 'rect') => out.push({ id: `decor-${id}`, type: 'shape', shape, x, y, w, h, rot: 0, fill: { type: 'solid', color }, line: null, text: null, decor: true });
  const titleSlide = layout === 'title' || layout === 'section';
  if (layout === 'blank') return out;
  if (theme.accentBand) {
    if (titleSlide) { R('band', 0, H - 28, W, 28, '@accent1'); R('band2', 0, H - 34, W, 6, '@accent2'); } else R('bar', 40 * sx, 62, 10 * sx, 92, '@accent1');
  }
  if (theme.titleBar) {
    if (titleSlide) { R('top', 0, 0, W, 14, '@accent1'); R('line', 160 * sx, 372, 960 * sx, 3, '@accent1'); } else R('rule', 88 * sx, 180, 1104 * sx, 3, '@accent1');
  }
  if (theme.dark && titleSlide) R('glow', W * 0.62, -H * 0.3, W * 0.6, H * 1.1, '@accent1:a18', 'ellipse');
  return out;
}

function footerHtml(pres, slide, index, opts) {
  const ft = pres.footer ?? {};
  if (!ft.slideNum && !ft.date && !ft.text) return '';
  if (ft.hideOnTitle && slide.layout === 'title') return '';
  const theme = pres.theme;
  const color = resolveColor(theme, '@tx1:lm50:lo50');
  const W = pres.size.w;
  const H = pres.size.h;
  const y = H - 54;
  const st = `position:absolute;top:${y}px;height:32px;line-height:32px;font-size:16px;color:${color};font-family:${fontStack(theme, '+mn')}`;
  let html = '';
  if (ft.date) html += `<div class="ftr" style="${st};left:${f(W * 0.069)}px;width:${f(W * 0.225)}px">${escHtml(ft.dateText || opts.today || new Date().toLocaleDateString('ko-KR'))}</div>`;
  if (ft.text) html += `<div class="ftr" style="${st};left:${f(W * 0.33)}px;width:${f(W * 0.34)}px;text-align:center">${escHtml(ft.text)}</div>`;
  if (ft.slideNum) html += `<div class="ftr" style="${st};left:${f(W * 0.706)}px;width:${f(W * 0.225)}px;text-align:right">${index + 1}</div>`;
  return html;
}

/** 슬라이드 배경 채우기 */
export function slideBackground(pres, slide) {
  return slide.bg ?? pres.theme.bg ?? { type: 'solid', color: '@bg1' };
}

/**
 * 슬라이드 하나 → HTML
 * opts: { prompt (빈 개체 틀 안내 글), editable (data-p/r), hide: Set(숨길 개체), index, links, skipIds: Set }
 */
export function slideHtml(pres, slide, opts = {}) {
  const theme = pres.theme;
  const media = pres.media ?? {};
  const bg = fillCss(theme, slideBackground(pres, slide), media);
  const decor = slide.bgObjects ?? (slide.hideDecor ? [] : themeDecor(theme, slide.layout, pres.size));
  let html = '';
  for (const o of decor) html += objectHtml(theme, o, media, {}).replace('class="ob"', 'class="ob decor"');
  for (const o of slide.objects) {
    if (opts.skipIds?.has(o.id) || o.hidden) continue;
    if (!opts.prompt && o.ph && (o.type === 'image' ? !o.media : isEmptyText(o.text))) continue;
    html += objectHtml(theme, o, media, o.id === opts.editId ? { ...opts, editable: true, prompt: false } : opts);
  }
  html += footerHtml(pres, slide, opts.index ?? 0, opts);
  const dark = isDark(bg.startsWith('#') || bg.startsWith('rgb') ? bg : '#ffffff');
  return `<div class="sl${dark ? ' dark' : ''}" style="position:relative;width:${pres.size.w}px;height:${pres.size.h}px;background:${bg};overflow:hidden">${html}</div>`;
}

/** 슬라이드 HTML 을 그리는 데 필요한 CSS (편집 화면 · 쇼 · 청중 창 · 인쇄 · PNG 내보내기가 같이 씀) */
export const SLIDE_CSS = `.sl{line-height:normal;text-align:left;color:#000;-webkit-font-smoothing:antialiased;font-kerning:normal;box-sizing:border-box}
.sl *{box-sizing:border-box}
.sl .ob{position:absolute;transform-origin:50% 50%}
.sl .ob.decor{pointer-events:none}
.sl .tx{overflow:visible}
.sl .txi{width:100%;outline:none;min-height:1em}
.sl .p{margin:0}
.sl .bu{user-select:none}
.sl svg{overflow:visible}
.sl table{border-spacing:0}
.sl td{position:relative}
.sl .lnk{text-decoration:underline;cursor:pointer}
.tw{position:relative;overflow:hidden}
.tw-in{transform-origin:0 0;position:absolute;left:0;top:0}`;
