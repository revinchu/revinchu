// 문서 → HTML 글 (DOM 없음). 편집 화면 · 인쇄 · 내보내기 · 측정이 같은 결과를 씀.
// 문단 = <div class="p" data-pid> · 런 = <span class="r" data-o="시작 위치"> (글자 위치 ↔ DOM 변환 기준)
import { effPara, effRun, fontName, paraLen, runLen, HIGHLIGHTS, styleChain, formatNumber } from './model.js';

export const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const f2 = (v) => Math.round(v * 100) / 100;

// 글꼴 이름 → CSS 글꼴 목록 (한국어 글꼴은 운영체제마다 이름이 달라 대체 이름을 덧붙임)
const FONT_ALIASES = {
  '맑은 고딕': ['Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', 'Noto Sans CJK KR', 'NanumGothic'],
  'Malgun Gothic': ['맑은 고딕', 'Apple SD Gothic Neo', 'Noto Sans KR'],
  '굴림': ['Gulim', 'Noto Sans KR', 'Apple SD Gothic Neo'], '굴림체': ['GulimChe', 'D2Coding', 'monospace'],
  '돋움': ['Dotum', 'Noto Sans KR', 'Apple SD Gothic Neo'], '돋움체': ['DotumChe', 'D2Coding', 'monospace'],
  '바탕': ['Batang', 'Noto Serif KR', 'AppleMyungjo', 'Nanum Myeongjo', 'serif'], '바탕체': ['BatangChe', 'Noto Serif KR', 'serif'],
  '궁서': ['Gungsuh', 'Noto Serif KR', 'serif'], '궁서체': ['GungsuhChe', 'Noto Serif KR', 'serif'],
  '함초롬바탕': ['HCR Batang', 'HCRBatang', 'Noto Serif KR', 'Nanum Myeongjo', 'serif'], '함초롬돋움': ['HCR Dotum', 'HCRDotum', 'Noto Sans KR', 'NanumGothic'],
  '한컴바탕': ['함초롬바탕', 'HCR Batang', 'Noto Serif KR', 'serif'], '한컴돋움': ['함초롬돋움', 'HCR Dotum', 'Noto Sans KR'],
  '휴먼명조': ['Noto Serif KR', 'serif'], 'HY헤드라인M': ['Noto Sans KR', 'sans-serif'], '신명조': ['Noto Serif KR', 'serif'], '중고딕': ['Noto Sans KR'], '견고딕': ['Noto Sans KR'],
  '나눔고딕': ['NanumGothic', 'Nanum Gothic', 'Noto Sans KR'], '나눔명조': ['NanumMyeongjo', 'Nanum Myeongjo', 'Noto Serif KR', 'serif'],
  Calibri: ['Carlito', 'Segoe UI', 'Arial'], Cambria: ['Caladea', 'Georgia', 'serif'], 'Times New Roman': ['Liberation Serif', 'Times', 'serif'], Arial: ['Liberation Sans', 'Helvetica', 'sans-serif'],
  'Courier New': ['Liberation Mono', 'Courier', 'monospace'], Consolas: ['Menlo', 'D2Coding', 'monospace'],
};
const GENERIC = ['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'];
const qf = (n) => (GENERIC.includes(n) ? n : `'${n.replace(/['"]/g, '')}'`);
export function fontStack(latin, ea) {
  const out = [];
  const add = (n) => { if (n && !out.includes(n)) out.push(n); };
  add(latin);
  // 라틴 글꼴에 없는 한글은 한글 글꼴로 (Word 의 동아시아 글꼴과 같은 효과)
  if (ea && ea !== latin) add(ea);
  for (const n of FONT_ALIASES[latin] ?? []) add(n);
  for (const n of FONT_ALIASES[ea] ?? []) add(n);
  add('Malgun Gothic'); add('Apple SD Gothic Neo'); add('Noto Sans KR'); add('sans-serif');
  return out.map(qf).join(',');
}
/** 글꼴별 기본 줄 높이 (글자 크기 배수) — Word '1줄' 의 높이에 가깝게 */
const LINE_FACTOR = { '맑은 고딕': 1.33, 'Malgun Gothic': 1.33, Calibri: 1.22, Cambria: 1.17, 'Times New Roman': 1.15, Arial: 1.15, '바탕': 1.2, '굴림': 1.2, '돋움': 1.2, '궁서': 1.2, '함초롬바탕': 1.2, '함초롬돋움': 1.2, Consolas: 1.17, 'Courier New': 1.13 };
export const lineFactor = (font) => LINE_FACTOR[font] ?? 1.25;

/** 색 이름 · 자동 → CSS */
const color = (c, auto = null) => (!c || c === 'auto' ? auto : c);

// ───────────── 표 스타일 (Word 기본 표 스타일 일부 · 색은 테마 강조색) ─────────────
export const ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const lighten = (hex, f) => { const n = parseInt(hex.slice(1), 16); const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * f)); return `#${ch.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`; };
/** 표 스타일 id → 칸별 모양 함수 (Word: 'GridTable4-Accent1' 등) */
export function tableLook(doc, t) {
  const id = t.style ?? 'TableGrid';
  const m = /^(Grid|List|Plain)Table(\d)(Light)?(?:-Accent(\d))?$/.exec(id);
  const look = { firstRow: true, firstCol: true, bandRows: true, lastRow: false, lastCol: false, bandCols: false, ...(t.tblPr?.look ?? {}) };
  const base = styleChain(doc.styles, id);
  const borders0 = base.tblPr?.borders ?? null;
  if (!m) return { borders: borders0, cell: () => null, look };
  const [, kind, n] = m;
  const acc = m[4] ? ACCENTS[Number(m[4]) - 1] : '#000000';
  const accL = m[4] ? lighten(acc, 0.8) : '#CCCCCC';
  const accM = m[4] ? lighten(acc, 0.6) : '#999999';
  const line = (c, w = 0.5) => ({ style: 'single', w, color: c });
  let borders = null;
  let cell = () => null;
  const rows = (t.rows ?? []).length;
  const band = (r, c) => (look.bandRows && ((r - (look.firstRow ? 1 : 0)) % 2 === 0)) || (look.bandCols && c % 2 === 0);
  if (kind === 'Grid') {
    borders = { top: line(accM), bottom: line(accM), left: line(accM), right: line(accM), insideH: line(accM), insideV: line(accM) };
    cell = (r, c) => {
      if (look.firstRow && r === 0) return n === '4' || n === '5' ? { shd: acc, color: '#FFFFFF', b: true } : { b: true, borderB: line(acc, 1.5) };
      if (look.lastRow && r === rows - 1) return { b: true, borderT: line(acc, 1.5) };
      const o = {};
      if (look.firstCol && c === 0) o.b = true;
      if (band(r, c)) o.shd = n === '5' ? accM : accL;
      return o;
    };
  } else if (kind === 'List') {
    borders = { top: line(acc), bottom: line(acc), insideH: n === '1' ? null : line(accM) };
    cell = (r, c) => {
      if (look.firstRow && r === 0) return n === '4' ? { shd: acc, color: '#FFFFFF', b: true } : { b: true, borderB: line(acc) };
      const o = {};
      if (look.firstCol && c === 0) o.b = true;
      if (band(r, c)) o.shd = accL;
      return o;
    };
  } else {
    borders = n === '1' ? { top: line('#BFBFBF'), bottom: line('#BFBFBF'), left: line('#BFBFBF'), right: line('#BFBFBF'), insideH: line('#BFBFBF'), insideV: line('#BFBFBF') } : n === '2' ? { top: line('#7F7F7F'), bottom: line('#7F7F7F'), insideH: line('#7F7F7F') } : null;
    cell = (r, c) => {
      if (look.firstRow && r === 0) return { b: true, ...(n === '3' ? { caps: true, borderB: line('#7F7F7F') } : {}) };
      const o = {};
      if (look.firstCol && c === 0) o.b = true;
      if (band(r, c) && n !== '2') o.shd = '#F2F2F2';
      return o;
    };
  }
  return { borders, cell, look };
}
/** 표 스타일 갤러리 목록 */
export function tableStyleList() {
  const out = [['TableGrid', '표 눈금'], ['PlainTable1', '일반 표 1'], ['PlainTable2', '일반 표 2'], ['PlainTable3', '일반 표 3']];
  for (const n of ['1Light', '2', '4', '5']) for (let a = 0; a <= 6; a++) out.push([`GridTable${n}${a ? `-Accent${a}` : ''}`, `눈금 표 ${n.replace('Light', ' 밝게')}${a ? ` - 강조 ${a}` : ''}`]);
  for (const n of ['1Light', '2', '4']) for (let a = 0; a <= 6; a++) out.push([`ListTable${n}${a ? `-Accent${a}` : ''}`, `목록 표 ${n.replace('Light', ' 밝게')}${a ? ` - 강조 ${a}` : ''}`]);
  return out;
}

// ───────────── 테두리 ─────────────
const BORDER_CSS = { single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dotDash: 'dashed', dotDotDash: 'dotted', triple: 'double', thinThickSmallGap: 'double', thickThinSmallGap: 'double', wave: 'solid', none: 'none', nil: 'none' };
export function borderCss(b) {
  if (!b || b.style === 'none' || b.style === 'nil') return 'none';
  const w = Math.max(b.style === 'double' ? 3 : 1, f2(((b.w ?? 0.5) * 4) / 3));
  return `${w}px ${BORDER_CSS[b.style] ?? 'solid'} ${color(b.color, '#000000') ?? '#000'}`;
}

// ───────────── 글자 서식 → CSS ─────────────
export function runCss(doc, rp, opts = {}) {
  const css = [];
  const latin = fontName(doc, rp.font ?? '+mn');
  const ea = fontName(doc, rp.fontEa ?? rp.font ?? '+mn', true);
  css.push(`font-family:${fontStack(latin, ea)}`);
  let sz = rp.sz ?? 10;
  if (rp.va === 'super' || rp.va === 'sub') { css.push(`vertical-align:${rp.va}`); sz *= 0.65; }
  css.push(`font-size:${f2(sz)}pt`);
  if (rp.b) css.push('font-weight:bold');
  if (rp.i) css.push('font-style:italic');
  const deco = [];
  if (rp.u && rp.u !== 'none') deco.push('underline');
  if (rp.s || rp.dstrike) deco.push('line-through');
  if (opts.del) deco.push('line-through');
  if (deco.length) {
    css.push(`text-decoration-line:${deco.join(' ')}`);
    const u = rp.u;
    if (rp.dstrike || u === 'double') css.push('text-decoration-style:double');
    else if (u === 'dotted' || u === 'dottedHeavy') css.push('text-decoration-style:dotted');
    else if (u === 'dash' || u === 'dashLong' || u === 'dashedHeavy') css.push('text-decoration-style:dashed');
    else if (u === 'wave' || u === 'wavyHeavy' || u === 'wavyDouble') css.push('text-decoration-style:wavy');
    if (u === 'thick' || /Heavy/.test(u ?? '')) css.push('text-decoration-thickness:2px');
    if (rp.uColor && rp.uColor !== 'auto') css.push(`text-decoration-color:${rp.uColor}`);
    css.push('text-decoration-skip-ink:none');
  }
  const col = color(rp.color);
  if (col) css.push(`color:${col}`);
  if (rp.hl && rp.hl !== 'none') css.push(`background-color:${HIGHLIGHTS[rp.hl] ?? rp.hl}`);
  else if (rp.shd) css.push(`background-color:${rp.shd}`);
  if (rp.caps) css.push('text-transform:uppercase');
  if (rp.smallCaps) css.push('font-variant:small-caps');
  if (rp.spc) css.push(`letter-spacing:${f2(rp.spc)}pt`);
  if (rp.scale && rp.scale !== 100) css.push(`display:inline-block;transform:scaleX(${rp.scale / 100});transform-origin:left`);
  if (rp.pos) css.push(`position:relative;top:${f2(-rp.pos)}pt`);
  if (rp.hidden && !opts.marks) css.push('display:none');
  else if (rp.hidden) css.push('text-decoration:underline dotted');
  if (rp.outline) css.push(`-webkit-text-stroke:0.6px ${col ?? '#000'};color:transparent`);
  if (rp.shadow) css.push('text-shadow:1px 1px 1px rgba(0,0,0,.45)');
  if (rp.emboss) css.push('text-shadow:-1px -1px 0 rgba(255,255,255,.7),1px 1px 0 rgba(0,0,0,.35)');
  if (rp.imprint) css.push('text-shadow:1px 1px 0 rgba(255,255,255,.7),-1px -1px 0 rgba(0,0,0,.35)');
  if (rp.emph) css.push(`text-emphasis:${rp.emph === 'comma' ? 'sesame' : rp.emph === 'circle' ? 'open circle' : 'filled dot'};text-emphasis-position:over right`);
  return css.join(';');
}

// ───────────── 문단 서식 → CSS ─────────────
export function paraCss(doc, ep, mainRpr, ctx = {}) {
  const css = [];
  const ind = ep.ind ?? {};
  const spc = ep.spc ?? {};
  css.push(`text-align:${ep.align === 'both' ? 'justify' : ep.align === 'distribute' ? 'justify' : ep.align ?? 'left'}`);
  if (ep.align === 'distribute') css.push('text-align-last:justify');
  if (ind.l) css.push(`padding-left:${f2(ind.l)}px`);
  if (ind.r) css.push(`padding-right:${f2(ind.r)}px`);
  const first = ind.first ?? (ind.hanging ? -ind.hanging : 0);
  if (first) css.push(`text-indent:${f2(first)}px`);
  const before = ctx.noBefore ? 0 : spc.before ?? 0;
  const after = ctx.noAfter ? 0 : spc.after ?? 0;
  const bd = ep.border;
  if (bd) {
    for (const side of ['top', 'bottom', 'left', 'right']) if (bd[side] && !(side === 'top' && ctx.joinTop) && !(side === 'bottom' && ctx.joinBottom)) css.push(`border-${side}:${borderCss(bd[side])}`);
    const sp = (s) => f2(((bd[s]?.space ?? 1) * 4) / 3);
    if (bd.top && !ctx.joinTop) css.push(`padding-top:${f2(before)}px;--bpt:${sp('top')}px`);
    if (bd.left) css.push(`padding-left:${f2((ind.l ?? 0) + Number(sp('left')))}px`);
  }
  // 간격: padding (여백 상쇄 없음 → 측정한 높이를 그대로 더할 수 있음)
  if (before) css.push(`padding-top:${f2(before + (bd?.top && !ctx.joinTop ? ((bd.top.space ?? 1) * 4) / 3 : 0))}px`);
  else if (bd?.top && !ctx.joinTop) css.push(`padding-top:${f2(((bd.top.space ?? 1) * 4) / 3)}px`);
  if (after) css.push(`padding-bottom:${f2(after + (bd?.bottom && !ctx.joinBottom ? ((bd.bottom.space ?? 1) * 4) / 3 : 0))}px`);
  else if (bd?.bottom && !ctx.joinBottom) css.push(`padding-bottom:${f2(((bd.bottom.space ?? 1) * 4) / 3)}px`);
  if (ep.shd) css.push(`background:${ep.shd}`);
  // 줄 간격: auto = 글꼴 기본 줄 높이 × 배수 (단위 없는 값 → 줄마다 가장 큰 글자 기준, Word 와 같음)
  const fam = fontName(doc, mainRpr.fontEa ?? mainRpr.font ?? '+mn', true);
  const fac = lineFactor(fam);
  const rule = spc.rule ?? 'auto';
  if (rule === 'exact') css.push(`line-height:${f2(spc.line)}px`);
  else if (rule === 'atLeast') css.push(`line-height:${f2(Math.max(spc.line ?? 0, ((mainRpr.sz ?? 10) * 4) / 3 * fac))}px`);
  else if (rule === 'pct') css.push(`line-height:${f2((spc.line ?? 1.6) * 1.0)}`); // 한글: 글자 크기의 %
  else css.push(`line-height:${f2((spc.line ?? 1) * fac)}`);
  if (ep.keepLines) css.push('break-inside:avoid');
  if (ep.bidi) css.push('direction:rtl');
  return css.join(';');
}

// ───────────── 번호 붙임표 ─────────────
function labelHtml(doc, p, ep, num, mainRpr) {
  if (!num) return '';
  const ind = ep.ind ?? {};
  const hanging = ind.hanging ?? (ind.first != null && ind.first < 0 ? -ind.first : 0);
  const rp = { ...mainRpr, ...(num.rPr ?? {}), ...(num.font ? { font: num.font, fontEa: num.font } : {}) };
  delete rp.u; delete rp.hl; delete rp.shd;
  const w = num.suff === 'nothing' ? 0 : num.suff === 'space' ? null : hanging;
  const style = `${runCss(doc, rp)};display:inline-block;text-indent:0;${w ? `min-width:${f2(w)}px;` : ''}white-space:pre`;
  return `<span class="lbl" contenteditable="false" style="${style}">${esc(num.text)}${num.suff === 'space' ? ' ' : ''}</span>`;
}

// ───────────── 런 HTML ─────────────
function imgHtml(doc, r, opts) {
  const src = opts.media?.[r.media] ?? doc.media?.[r.media] ?? r.src ?? '';
  const w = f2(r.w ?? 100); const h = f2(r.h ?? 100);
  const fx = [];
  if (r.rot) fx.push(`transform:rotate(${r.rot}deg)`);
  const filt = [];
  if (r.bright) filt.push(`brightness(${1 + r.bright / 100})`);
  if (r.contrast) filt.push(`contrast(${1 + r.contrast / 100})`);
  if (r.gray) filt.push('grayscale(1)');
  if (filt.length) fx.push(`filter:${filt.join(' ')}`);
  if (r.border) fx.push(`outline:${borderCss(r.border)}`);
  let inner;
  if (r.crop) {
    const c = r.crop; const fw = w / (1 - c.l - c.r); const fh = h / (1 - c.t - c.b);
    inner = `<span style="display:inline-block;width:${w}px;height:${h}px;overflow:hidden;position:relative;vertical-align:bottom"><img src="${esc(src)}" alt="${esc(r.alt ?? '')}" draggable="false" style="position:absolute;left:${f2(-c.l * fw)}px;top:${f2(-c.t * fh)}px;width:${f2(fw)}px;height:${f2(fh)}px"></span>`;
  } else inner = r.svg ? `<span class="svgobj" style="display:inline-block;width:${w}px;height:${h}px">${r.svg}</span>` : `<img src="${esc(src)}" alt="${esc(r.alt ?? '')}" draggable="false" style="width:${w}px;height:${h}px;vertical-align:bottom">`;
  return { inner, fx: fx.join(';') };
}
function runHtml(doc, p, r, off, opts, ctxRun) {
  const rp = effRun(doc, p, r, opts.tblStyle);
  if (ctxRun) Object.assign(rp, ctxRun);
  const tc = r.ins ? ' tc-ins' : r.del ? ' tc-del' : '';
  const cm = opts.cmtOpen?.size ? ` cmt" data-cmt="${[...opts.cmtOpen].join(' ')}` : '';
  if (r.del && opts.markup === 'none') return '';
  const css = runCss(doc, rp, { marks: opts.marks, del: r.del && opts.markup !== 'none' });
  const tcCss = r.ins && opts.markup !== 'none' ? `;color:${opts.authorColor?.(r.ins.author) ?? '#C00000'};text-decoration-line:underline` : r.del && opts.markup !== 'none' ? `;color:${opts.authorColor?.(r.del.author) ?? '#C00000'}` : '';
  const lnk = r.link ? ` data-link="${esc(r.link.url ?? `#${r.link.anchor}`)}"` : '';
  if (!r.obj) {
    let t = esc(r.text);
    if (opts.marks) t = t.replace(/ /g, '<span class="sp"> </span>').replace(/　/g, '<span class="sp ide">　</span>');
    return `<span class="r${tc}${cm}" data-o="${off}"${lnk} style="${css}${tcCss}">${t}</span>`;
  }
  switch (r.obj) {
    case 'tab': return `<span class="r tab${tc}" data-o="${off}" style="${css}${tcCss}">${opts.marks ? '<i class="tabm">→</i>' : ''}​</span>`;
    case 'br':
      if (r.kind === 'page') return `<span class="r pbr" data-o="${off}" data-br="page" style="${css}">${opts.marks ? '<i class="brm">··········· 페이지 나누기 ···········</i>' : '​'}</span>`;
      if (r.kind === 'column') return `<span class="r cbr" data-o="${off}" data-br="column" style="${css}">${opts.marks ? '<i class="brm">··· 단 나누기 ···</i>' : '​'}</span>`;
      return `<span class="r lbr" data-o="${off}" style="${css}">${opts.marks ? '<i class="lbrm">↵</i>' : '​'}</span><br>`;
    case 'img': {
      const { inner, fx } = imgHtml(doc, r, opts);
      if (r.float) {
        const fl = r.float;
        if (fl.wrap === 'behind' || fl.wrap === 'front' || fl.wrap === 'none') return `<span class="r anc" data-o="${off}" data-anchor="${esc(r.id ?? '')}">​</span>`;
        const side = fl.side ?? (fl.x != null && fl.x > 200 ? 'right' : 'left');
        const m = fl.dist ?? { t: 0, b: 0, l: 12, r: 12 };
        const top = fl.wrap === 'topBottom' ? `clear:both;float:none;display:block;margin:${f2(m.t ?? 0)}px auto ${f2(m.b ?? 0)}px ${f2(fl.x ?? 0)}px` : `float:${side};margin:${f2(m.t ?? 0)}px ${f2(m.r ?? 12)}px ${f2(m.b ?? 0)}px ${f2(m.l ?? 12)}px`;
        return `<span class="r obj fl" data-o="${off}" data-img="${esc(r.id ?? '')}" style="${top};${fx};line-height:0">${inner}</span>`;
      }
      return `<span class="r obj" data-o="${off}" data-img="${esc(r.id ?? '')}" style="display:inline-block;line-height:0;${fx}">${inner}</span>`;
    }
    case 'field': {
      const code = (r.code ?? '').trim().split(/\s+/)[0]?.toUpperCase();
      const live = code === 'PAGE' || code === 'NUMPAGES' || code === 'SECTIONPAGES';
      let txt = live ? (opts.fields?.[code] ?? r.result ?? '1') : r.result ?? '';
      if (code === 'MERGEFIELD') { const name = (r.code ?? '').trim().split(/\s+/)[1]?.replace(/"/g, '') ?? ''; txt = opts.merge ? String(opts.merge[name] ?? '') : `«${name}»`; return `<span class="r fld mf${opts.mergeShade ? ' fshade' : ''}" data-o="${off}" data-field="MERGEFIELD" style="${css}${tcCss}">${esc(txt) || '\u200b'}</span>`; }
      return `<span class="r fld${opts.fieldShade ? ' fshade' : ''}${tc}" data-o="${off}" data-field="${esc(code ?? '')}"${lnk} style="${css}${tcCss}">${esc(txt) || '​'}</span>`;
    }
    case 'note': {
      const n = opts.noteNums?.get(r.id) ?? '*';
      return `<span class="r note${tc}" data-o="${off}" data-note="${esc(r.id)}" data-kind="${r.kind ?? 'foot'}" style="${runCss(doc, { ...rp, va: 'super' })}">${esc(r.custom ?? n)}</span>`;
    }
    case 'sym': return `<span class="r${tc}" data-o="${off}" style="${css};${r.font ? `font-family:${qf(r.font)}` : ''}">${esc(r.char ?? '')}</span>`;
    case 'eq': return `<span class="r obj eq" data-o="${off}" style="${css};display:inline-block">${r.html ?? esc(r.latex ?? '')}</span>`;
    case 'mark': return r.kind === 'bm' && opts.marks ? `<span class="r bm" data-o="${off}">${r.edge === 'start' ? '[' : ']'}</span>` : '';
    case 'shape': {
      return `<span class="r anc" data-o="${off}" data-anchor="${esc(r.id ?? '')}">​</span>`;
    }
    default: return `<span class="r" data-o="${off}">​</span>`;
  }
}

/** 문단 HTML. opts: { num (computeNumbering 결과), tblStyle, media, marks, markup, fields, noteNums, noBefore/noAfter, cellRun, authorColor } */
export function paraHtml(doc, p, opts = {}) {
  const ep = effPara(doc, p, opts.tblStyle);
  const firstRun = p.runs.find((r) => !r.obj || r.obj === 'tab') ?? { ...(p.mark ?? {}), text: '' };
  const mainRpr = effRun(doc, p, firstRun, opts.tblStyle);
  if (opts.cellRun) Object.assign(mainRpr, opts.cellRun);
  const num = opts.num?.get(p.id);
  let ep2 = ep;
  if (num?.ind) {
    // 목록 들여쓰기: 문단에 직접 들여쓰기가 없으면 목록 수준의 들여쓰기
    const ind = { ...ep.ind };
    if (p.pPr?.ind?.l == null) ind.l = num.ind.l;
    if (p.pPr?.ind?.hanging == null && p.pPr?.ind?.first == null) { ind.hanging = num.ind.hanging ?? 0; delete ind.first; }
    ep2 = { ...ep, ind };
  }
  const css = paraCss(doc, ep2, mainRpr, opts);
  const cmtOpen = new Set(opts.cmtOpenIn ?? []);
  let html = labelHtml(doc, p, ep2, num, mainRpr);
  let off = 0;
  const o2 = { ...opts, cmtOpen };
  let runs = p.runs;
  // 드롭 캡: 첫 글자를 크게 (여러 줄 높이) · 왼쪽에 띄움
  if (ep.dropCap && runs[0] && !runs[0].obj && runs[0].text) {
    const r0 = runs[0];
    const rp0 = effRun(doc, p, r0, opts.tblStyle);
    const lines = ep.dropCap.lines ?? 3;
    const fam = fontName(doc, rp0.fontEa ?? rp0.font ?? '+mn', true);
    const lh = (rp0.sz ?? 10) * lineFactor(fam) * (ep.spc?.line && (ep.spc.rule ?? 'auto') === 'auto' ? ep.spc.line : 1);
    const dcCss = `${runCss(doc, { ...rp0, sz: lh * lines * 0.98 })};float:left;line-height:${f2(((lh * lines) * 4) / 3)}px;margin:${f2(-lh * 0.12)}px 4px 0 0${ep.dropCap.kind === 'margin' ? ';position:relative;margin-left:-1.1em' : ''};text-indent:0`;
    html += `<span class="r dc" data-o="0" style="${dcCss}">${esc(r0.text[0])}</span>`;
    off = 1;
    runs = [{ ...r0, text: r0.text.slice(1) }, ...runs.slice(1)];
    if (!runs[0].text) runs = runs.slice(1);
  }
  for (const r of runs) {
    if (r.obj === 'mark' && r.kind === 'cmt') { if (r.edge === 'start') cmtOpen.add(r.id); else cmtOpen.delete(r.id); }
    html += runHtml(doc, p, r, off, o2, opts.cellRun);
    off += runLen(r);
  }
  // 빈 문단 · 끝 표시: 문단 표시 서식의 빈 글자 (높이 · 커서 자리)
  const markRp = effRun(doc, p, { ...(p.mark ?? (p.runs.length ? {} : {})), text: '' }, opts.tblStyle);
  if (opts.cellRun) Object.assign(markRp, opts.cellRun);
  const markCss = runCss(doc, p.runs.length ? effRun(doc, p, lastTextRun(p) ?? {}, opts.tblStyle) : markRp);
  html += `<span class="r end" data-o="${off}" style="${markCss}">​${opts.marks ? `<i class="pm">${p.secBreak ? `::::: 구역 나누기 (${SEC_KIND[p.secBreak.kind] ?? '다음 페이지'}) :::::` : '¶'}</i>` : ''}</span>`;
  const cls = `p${num ? ' li' : ''}${opts.extraClass ? ` ${opts.extraClass}` : ''}`;
  const lvl = ep.outline != null && ep.outline < 9 ? ` data-lvl="${ep.outline + 1}"` : '';
  return `<div class="${cls}" data-pid="${p.id}"${lvl} style="${css}">${html}</div>`;
}
const SEC_KIND = { nextPage: '다음 페이지', continuous: '이어서', evenPage: '짝수 페이지', oddPage: '홀수 페이지' };
const lastTextRun = (p) => { for (let i = p.runs.length - 1; i >= 0; i--) if (!p.runs[i].obj) return p.runs[i]; return null; };

// ───────────── 표 ─────────────
/** 칸 병합 정보: rows[r].cells[c] → { gc (격자 시작 열), span, rowSpan, hidden(세로 병합 이어짐) } */
export function tableGrid(t) {
  const info = t.rows.map((row) => {
    let gc = 0;
    return row.cells.map((cell) => { const o = { gc, span: cell.span ?? 1, rowSpan: 1, hidden: cell.vMerge === 'cont' }; gc += o.span; return o; });
  });
  // 세로 병합: restart 칸 아래로 cont 칸 수
  for (let r = 0; r < t.rows.length; r++) {
    t.rows[r].cells.forEach((cell, c) => {
      if (cell.vMerge !== 'restart') return;
      const gc = info[r][c].gc;
      let n = 1;
      for (let k = r + 1; k < t.rows.length; k++) {
        const j = info[k].findIndex((x) => x.gc === gc);
        if (j < 0 || t.rows[k].cells[j].vMerge !== 'cont') break;
        n++;
      }
      info[r][c].rowSpan = n;
    });
  }
  return info;
}
/** 표 HTML (rowsFrom~rowsTo 만 · 머리글 행 반복은 layout 이 rows 를 지정) */
export function tableHtml(doc, t, opts = {}) {
  const look = tableLook(doc, t);
  const tp = { ...styleChain(doc.styles, t.style ?? 'TableGrid').tblPr, ...t.tblPr };
  const borders = { ...(look.borders ?? {}), ...(tp.borders ?? {}) };
  const info = tableGrid(t);
  const grid = t.grid;
  const width = grid.reduce((a, b) => a + b, 0);
  const mar = { l: 7.2, r: 7.2, t: 0, b: 0, ...(tp.cellMar ?? {}) };
  const ml = tp.ind ?? 0;
  const align = tp.align === 'center' ? 'margin-left:auto;margin-right:auto' : tp.align === 'right' ? 'margin-left:auto' : `margin-left:${f2(ml - mar.l)}px`;
  let html = `<table class="tbl" data-tid="${t.id}" style="width:${f2(width)}px;${align}${tp.shd ? `;background:${tp.shd}` : ''}"><colgroup>${grid.map((w) => `<col style="width:${f2(w)}px">`).join('')}</colgroup><tbody>`;
  const rows = opts.rows ?? t.rows.map((_, i) => i);
  const nrows = t.rows.length;
  for (const r of rows) {
    const row = t.rows[r];
    const hcss = row.h ? (row.hRule === 'exact' ? `height:${f2(row.h)}px` : `height:${f2(row.h)}px`) : '';
    html += `<tr data-r="${r}"${hcss ? ` style="${hcss}"` : ''}>`;
    row.cells.forEach((cell, c) => {
      let ci = info[r][c];
      if (opts.rows) {
        // 잘린 표 (쪽 나눔): 세로 병합을 보이는 행 안으로 줄임
        const at = rows.indexOf(r);
        const restartShown = (() => { if (!ci.hidden) return true; for (let k = r - 1; k >= 0; k--) { const j = info[k].findIndex((x) => x.gc === ci.gc); if (j < 0) return false; if (!info[k][j].hidden) return rows.includes(k) && rows.indexOf(k) < at; } return false; })();
        if (ci.hidden && restartShown) return;
        let span = 1;
        if (ci.rowSpan > 1 || ci.hidden) {
          for (let k = r + 1; rows[at + span] === k && k < t.rows.length; k++) { const j = info[k].findIndex((x) => x.gc === ci.gc); if (j < 0 || t.rows[k].cells[j].vMerge !== 'cont') break; span++; }
        }
        ci = { ...ci, hidden: false, rowSpan: span };
      } else if (ci.hidden) return;
      const st = look.cell(r, ci.gc) ?? {};
      const cb = cell.borders ?? {};
      const last = ci.gc + ci.span >= grid.length;
      const bottomRow = r + ci.rowSpan >= nrows;
      const side = (k, outer, inner) => cb[k] ?? (outer ? borders[k] : borders[inner]);
      const bt = st.borderT ?? side('top', r === 0, 'insideH');
      const bb = st.borderB ?? side('bottom', bottomRow, 'insideH');
      const bl = side('left', ci.gc === 0, 'insideV');
      const br = side('right', last, 'insideV');
      const cm = { ...mar, ...(cell.mar ?? {}) };
      const w = grid.slice(ci.gc, ci.gc + ci.span).reduce((a, b) => a + b, 0);
      const css = [`width:${f2(w)}px`, `border-top:${borderCss(bt)}`, `border-bottom:${borderCss(bb)}`, `border-left:${borderCss(bl)}`, `border-right:${borderCss(br)}`, `padding:${f2(cm.t)}px ${f2(cm.r)}px ${f2(cm.b)}px ${f2(cm.l)}px`, `vertical-align:${cell.vAlign === 'center' ? 'middle' : cell.vAlign === 'bottom' ? 'bottom' : 'top'}`];
      const shd = cell.shd ?? st.shd;
      if (shd) css.push(`background:${shd}`);
      if (cell.dir === 'vert') css.push('writing-mode:vertical-rl');
      const cellRun = {};
      if (st.b) cellRun.b = true;
      if (st.color) cellRun.color = st.color;
      if (st.caps) cellRun.caps = true;
      const inner = cell.blocks.map((b) => blockHtml(doc, b, { ...opts, tblStyle: t.style, cellRun: Object.keys(cellRun).length ? cellRun : null, inCell: true, rows: undefined })).join('');
      html += `<td data-c="${c}"${ci.span > 1 ? ` colspan="${ci.span}"` : ''}${ci.rowSpan > 1 ? ` rowspan="${ci.rowSpan}"` : ''} style="${css.join(';')}">${inner}</td>`;
    });
    html += '</tr>';
  }
  return `${html}</tbody></table>`;
}

export function blockHtml(doc, b, opts = {}) {
  if (b.type === 'tbl') return tableHtml(doc, b, opts);
  return paraHtml(doc, b, opts);
}

// ───────────── 문서 CSS ─────────────
export const DOC_CSS = `
.wd{color:#000;font-kerning:normal;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
.wd.hyph .p,.wd .hyph .p{hyphens:auto}
.wd .p{white-space:pre-wrap;overflow-wrap:break-word;word-break:keep-all;margin:0;position:relative;tab-size:48px;min-height:1px}
.wd .p.li{}
.wd .r{white-space:pre-wrap}
.wd .tab{display:inline-block;min-width:1px;white-space:pre;text-indent:0}
.wd .tab.lead-dot{background:radial-gradient(circle,currentColor .7px,transparent 1px) bottom/5px 4px repeat-x}
.wd .tab.lead-hyphen{background:linear-gradient(currentColor,currentColor) bottom 3px left 0/100% 1px no-repeat}
.wd .tab.lead-underscore{border-bottom:1px solid currentColor}
.wd .tab.lead-middleDot{background:radial-gradient(circle,currentColor .7px,transparent 1px) center/5px 4px repeat-x}
.wd .lbl{user-select:none}
.wd .tbl{border-collapse:collapse;table-layout:fixed;border-spacing:0}
.wd .tbl td{overflow-wrap:break-word;box-sizing:border-box}
.wd .tbl td>.p:last-child{}
.wd a{color:inherit}
.wd .note{cursor:default}
.wd .fshade{background:#D9D9D9}
.wd .tc-ins{text-decoration-line:underline}
.wd .tc-del{text-decoration-line:line-through}
.wd .cmt{background:rgba(255,209,102,.35)}
.wd.marks .pm,.wd.marks .tabm,.wd.marks .brm,.wd.marks .lbrm{color:#4472C4;font-style:normal;font-family:Arial,sans-serif;font-weight:normal;text-decoration:none;user-select:none;pointer-events:none}
.wd .pm{position:absolute;white-space:nowrap}
.wd .pm{position:static}
.wd .tabm{position:absolute;font-size:9px}
.wd .brm{display:block;text-align:center;font-size:9pt;white-space:nowrap}
.wd .sp{position:relative}
.wd.marks .sp::after{content:'·';position:absolute;left:50%;top:0;transform:translateX(-50%);color:#4472C4;pointer-events:none}
.wd .bm{color:#808080;font-weight:bold}
.wd img{user-select:none}
.wd .obj.sel img,.wd .obj.sel .svgobj{outline:2px solid #2B579A;outline-offset:1px}
`;
