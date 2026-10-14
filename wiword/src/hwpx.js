// 한글 문서 (.hwpx · OWPML, KS X 6101) 읽기 · 쓰기 — DOM 없음
// 구조 참고: kordoc (MIT, chrisryugj/kordoc) 의 hwpx 생성기 · 파서, 한컴오피스 실제 저장본
// 단위: HWPUNIT = 1/7200 인치 = px × 75, 글자 크기 height = pt × 100
import { unzip } from './zip.js';
import { parseXml, child, kids, esc } from './xml.js';
import { newDoc, newSection, para, uid, normalizeRuns, effRun, effPara, walkParas, runLen, computeNumbering, docText, paraPlain } from './model.js';
import { mediaDataUrl, dataUrlBytes } from './docx.js';
import { lineFactor } from './render.js';

const dec = new TextDecoder();
const HU = 75;
const hu = (v) => (v == null || v === '' ? 0 : Number(v) / HU);
const toHU = (px) => Math.round((px ?? 0) * HU);
const deep = (el, name, out = []) => { for (const c of el?.children ?? []) { if (c.name === name) out.push(c); deep(c, name, out); } return out; };
const first = (el, name) => deep(el, name)[0] ?? null;
const A = (el, a) => el?.attrs?.[a];
const num = (v, d = 0) => { const n = Number(v); return Number.isFinite(n) ? n : d; };
const colorOf = (v) => (!v || v === 'none' || v === '-1' ? null : /^#[0-9a-f]{6}$/i.test(v) ? v.toUpperCase() : /^#[0-9a-f]{8}$/i.test(v) ? `#${v.slice(3).toUpperCase()}` : null);
const mmOf = (w) => { const m = /([\d.]+)\s*mm/.exec(w ?? ''); return m ? Number(m[1]) : 0.12; };

// 번호 모양
const NUMFMT_IN = { DIGIT: 'decimal', CIRCLED_DIGIT: 'decimalCircle', HANGUL_SYLLABLE: 'ganada', CIRCLED_HANGUL_SYLLABLE: 'ganada', HANGUL_JAMO: 'chosung', CIRCLED_HANGUL_JAMO: 'chosung', LATIN_CAPITAL: 'upperLetter', LATIN_SMALL: 'lowerLetter', CIRCLED_LATIN_CAPITAL: 'upperLetter', CIRCLED_LATIN_SMALL: 'lowerLetter', ROMAN_CAPITAL: 'upperRoman', ROMAN_SMALL: 'lowerRoman', HANGUL_PHONETIC: 'koreanDigital', IDEOGRAPH: 'ideographDigital', CIRCLED_IDEOGRAPH: 'ideographDigital' };
const NUMFMT_OUT = { decimal: 'DIGIT', decimalCircle: 'CIRCLED_DIGIT', ganada: 'HANGUL_SYLLABLE', chosung: 'HANGUL_JAMO', upperLetter: 'LATIN_CAPITAL', lowerLetter: 'LATIN_SMALL', upperRoman: 'ROMAN_CAPITAL', lowerRoman: 'ROMAN_SMALL', koreanDigital: 'HANGUL_PHONETIC', koreanCounting: 'HANGUL_PHONETIC', ideographDigital: 'IDEOGRAPH', decimalZero: 'DIGIT' };
const ALIGN_IN = { JUSTIFY: 'justify', LEFT: 'left', RIGHT: 'right', CENTER: 'center', DISTRIBUTE: 'distribute', DISTRIBUTE_SPACE: 'distribute' };
const ALIGN_OUT = { justify: 'JUSTIFY', left: 'LEFT', right: 'RIGHT', center: 'CENTER', distribute: 'DISTRIBUTE' };
const LINE_IN = { SOLID: 'single', DASH: 'dashed', DOT: 'dotted', DASH_DOT: 'dotDash', DASH_DOT_DOT: 'dotDotDash', LONG_DASH: 'dashed', CIRCLE: 'dotted', DOUBLE_SLIM: 'double', SLIM_THICK: 'thinThickSmallGap', THICK_SLIM: 'thickThinSmallGap', SLIM_THICK_SLIM: 'triple', WAVE: 'wave', DOUBLE_WAVE: 'wave' };
const LINE_OUT = { single: 'SOLID', thick: 'SOLID', dashed: 'DASH', dotted: 'DOT', dotDash: 'DASH_DOT', dotDotDash: 'DASH_DOT_DOT', double: 'DOUBLE_SLIM', thinThickSmallGap: 'SLIM_THICK', thickThinSmallGap: 'THICK_SLIM', triple: 'SLIM_THICK_SLIM', wave: 'WAVE' };
const borderIn = (el) => (!el || A(el, 'type') === 'NONE' ? { style: 'none' } : { style: LINE_IN[A(el, 'type')] ?? 'single', w: Math.max(0.25, (mmOf(A(el, 'width')) * 72) / 25.4), color: colorOf(A(el, 'color')) ?? '#000000' });

/** <hh:margin> 은 hp:switch 의 hp:case (HwpUnitChar) 값이 진짜 — hp:default 는 2배인 경우가 있음 */
function marginOf(pp) {
  const sw = child(pp, 'switch');
  const cs = sw ? child(sw, 'case') : null;
  return { margin: child(cs, 'margin') ?? child(pp, 'margin') ?? child(child(sw, 'default'), 'margin'), ls: child(cs, 'lineSpacing') ?? child(pp, 'lineSpacing') ?? child(child(sw, 'default'), 'lineSpacing') };
}

// ───────────── 머리 (header.xml) ─────────────
function readHead(xml) {
  const H = { fonts: { HANGUL: [], LATIN: [] }, bf: {}, char: {}, para: {}, styles: [], numbering: {}, bullets: {}, tabs: {} };
  if (!xml) return H;
  const x = parseXml(xml);
  for (const ff of deep(x, 'fontface')) { const lang = A(ff, 'lang'); H.fonts[lang] = kids(ff, 'font').map((f) => A(f, 'face')); }
  for (const b of deep(x, 'borderFill')) {
    const brush = first(b, 'winBrush');
    const grad = first(b, 'gradation');
    H.bf[A(b, 'id')] = {
      l: borderIn(child(b, 'leftBorder')), r: borderIn(child(b, 'rightBorder')), t: borderIn(child(b, 'topBorder')), b: borderIn(child(b, 'bottomBorder')),
      fill: brush ? colorOf(A(brush, 'faceColor')) : grad ? colorOf(A(kids(grad, 'color')[0], 'value')) : null,
    };
  }
  for (const c of deep(x, 'charPr')) {
    const fr = child(c, 'fontRef');
    const o = { sz: num(A(c, 'height'), 1000) / 100 };
    const col = colorOf(A(c, 'textColor'));
    if (col && col !== '#000000') o.color = col;
    const sh = colorOf(A(c, 'shadeColor'));
    if (sh && sh !== '#FFFFFF') o.shd = sh;
    const ea = H.fonts.HANGUL?.[num(A(fr, 'hangul'))];
    const lat = H.fonts.LATIN?.[num(A(fr, 'latin'))];
    if (ea) o.fontEa = ea;
    if (lat) o.font = lat;
    if (child(c, 'bold') || A(c, 'bold') === '1') o.b = true;
    if (child(c, 'italic') || A(c, 'italic') === '1') o.i = true;
    const u = child(c, 'underline');
    if (u && A(u, 'type') && A(u, 'type') !== 'NONE') { o.u = A(u, 'shape') === 'DOUBLE_SLIM' ? 'double' : A(u, 'shape') === 'DOT' ? 'dotted' : A(u, 'shape') === 'DASH' ? 'dash' : A(u, 'shape') === 'WAVE' ? 'wave' : true; const uc = colorOf(A(u, 'color')); if (uc && uc !== '#000000') o.uColor = uc; }
    const st = child(c, 'strikeout');
    if (st && A(st, 'shape') && A(st, 'shape') !== 'NONE' && A(st, 'shape') !== '3D') o.s = true;
    if (child(c, 'supscript')) o.va = 'super';
    if (child(c, 'subscript')) o.va = 'sub';
    const ol = child(c, 'outline');
    if (ol && A(ol, 'type') && A(ol, 'type') !== 'NONE') o.outline = true;
    const sd = child(c, 'shadow');
    if (sd && A(sd, 'type') && A(sd, 'type') !== 'NONE') o.shadow = true;
    if (child(c, 'emboss')) o.emboss = true;
    if (child(c, 'engrave')) o.imprint = true;
    const ratio = num(A(child(c, 'ratio'), 'hangul'), 100);
    if (ratio !== 100) o.scale = ratio;
    const spc = num(A(child(c, 'spacing'), 'hangul'), 0);
    if (spc) o.spc = Math.round(((spc / 100) * o.sz) * 100) / 100;
    const off = num(A(child(c, 'offset'), 'hangul'), 0);
    if (off) o.pos = Math.round((-off / 100) * o.sz * 100) / 100;
    if (A(c, 'symMark') && A(c, 'symMark') !== 'NONE') o.emph = A(c, 'symMark') === 'DOT_ABOVE' ? 'dot' : A(c, 'symMark') === 'RING_ABOVE' ? 'circle' : 'dot';
    H.char[A(c, 'id')] = o;
  }
  for (const t of deep(x, 'tabPr')) H.tabs[A(t, 'id')] = kids(t, 'tabItem').map((it) => ({ pos: hu(A(it, 'pos')), align: { LEFT: 'left', RIGHT: 'right', CENTER: 'center', DECIMAL: 'decimal' }[A(it, 'type')] ?? 'left', ...(A(it, 'leader') && A(it, 'leader') !== 'NONE' ? { leader: A(it, 'leader') === 'DOT' ? 'dot' : 'hyphen' } : {}) }));
  for (const p of deep(x, 'paraPr')) {
    const o = {};
    const al = A(child(p, 'align'), 'horizontal');
    o.align = ALIGN_IN[al] ?? 'justify';
    const { margin, ls } = marginOf(p);
    if (margin) {
      const v = (n) => num(A(child(margin, n), 'value'));
      const unit = A(child(margin, 'left'), 'unit') ?? 'HWPUNIT';
      const k = unit === 'CHAR' ? 13.33 : 1 / HU;
      const left = v('left') * k; const right = v('right') * k; const intent = v('intent') * (A(child(margin, 'intent'), 'unit') === 'CHAR' ? 13.33 : 1 / HU);
      const ind = {};
      if (left) ind.l = left;
      if (right) ind.r = right;
      // 한글 '내어쓰기'(intent < 0): 둘째 줄부터 들어감 — 왼쪽 여백이 둘째 줄 기준
      if (intent > 0) ind.first = intent;
      else if (intent < 0) { ind.l = (ind.l ?? 0) - intent; ind.hanging = -intent; }
      if (Object.keys(ind).length) o.ind = ind;
      const prev = v('prev') / HU; const next = v('next') / HU;
      o.spc = { before: prev, after: next };
    }
    if (ls) {
      const t = A(ls, 'type'); const val = num(A(ls, 'value'), 160);
      o.spc = { ...(o.spc ?? {}), ...(t === 'PERCENT' ? { line: val / 100, rule: 'pct' } : t === 'FIXED' ? { line: val / HU, rule: 'exact' } : t === 'AT_LEAST' ? { line: val / HU, rule: 'atLeast' } : { line: 1.6, rule: 'pct' }) };
    }
    const bs = child(p, 'breakSetting');
    if (A(bs, 'keepWithNext') === '1') o.keepNext = true;
    if (A(bs, 'keepLines') === '1') o.keepLines = true;
    if (A(bs, 'pageBreakBefore') === '1') o.pageBreakBefore = true;
    if (A(bs, 'widowOrphan') === '0') o.widow = false;
    const hd = child(p, 'heading');
    if (hd && A(hd, 'type') !== 'NONE') o._head = { type: A(hd, 'type'), id: A(hd, 'idRef'), level: num(A(hd, 'level')) };
    const bd = child(p, 'border');
    if (bd && A(bd, 'borderFillIDRef') && A(bd, 'borderFillIDRef') !== '1') o._bf = A(bd, 'borderFillIDRef');
    if (A(p, 'tabPrIDRef') && H.tabs[A(p, 'tabPrIDRef')]?.length) o.tabs = H.tabs[A(p, 'tabPrIDRef')];
    H.para[A(p, 'id')] = o;
  }
  for (const s of deep(x, 'style')) H.styles[A(s, 'id')] = { name: A(s, 'name'), eng: A(s, 'engName'), type: A(s, 'type'), para: A(s, 'paraPrIDRef'), char: A(s, 'charPrIDRef'), next: A(s, 'nextStyleIDRef') };
  for (const n of deep(x, 'numbering')) {
    const heads = {};
    for (const h of kids(n, 'paraHead')) heads[num(A(h, 'level'))] = { fmt: NUMFMT_IN[A(h, 'numFormat')] ?? 'decimal', text: (h.text ?? '').replace(/\^(\d+)/g, '%$1'), start: num(A(h, 'start'), 1) };
    H.numbering[A(n, 'id')] = { heads, start: num(A(n, 'start'), 1) };
  }
  for (const b of deep(x, 'bullet')) H.bullets[A(b, 'id')] = A(b, 'char') || '●';
  return H;
}

// ───────────── 본문 (sectionN.xml) ─────────────
function readSecPr(sp, ctx, firstPara) {
  const sec = newSection();
  sec.headers = {}; sec.footers = {};
  const pg = child(sp, 'pagePr');
  if (pg) {
    const w = hu(A(pg, 'width')); const h = hu(A(pg, 'height'));
    const land = A(pg, 'landscape') === 'NARROWLY';
    sec.page = { w: land ? Math.max(w, h) : w, h: land ? Math.min(w, h) : h, orient: land ? 'landscape' : 'portrait' };
    const m = child(pg, 'margin');
    // 한글: 위 여백 = 종이 위 ~ 머리말 시작, 머리말 영역이 따로 있음 → Word 처럼 본문 시작 = top + header
    const top = hu(A(m, 'top')); const bottom = hu(A(m, 'bottom')); const header = hu(A(m, 'header')); const footer = hu(A(m, 'footer'));
    sec.margin = { t: top + header, b: bottom + footer, l: hu(A(m, 'left')), r: hu(A(m, 'right')), header: top, footer: bottom, gutter: hu(A(m, 'gutter')) };
  }
  const sn = child(sp, 'startNum');
  if (sn && num(A(sn, 'page')) > 0) sec.pgNum = { fmt: 'decimal', start: num(A(sn, 'page')) };
  const col = first(firstPara, 'colPr');
  if (col && num(A(col, 'colCount'), 1) > 1) sec.cols = { n: num(A(col, 'colCount'), 1), space: hu(A(col, 'sameGap')) || 48, sep: !!child(col, 'colLine') && A(child(col, 'colLine'), 'type') !== 'NONE' };
  const pbf = child(sp, 'pageBorderFill');
  if (pbf) { const bf = ctx.H.bf[A(pbf, 'borderFillIDRef')]; if (bf && bf.t.style !== 'none') sec.pageBorder = { ...bf.t, space: 12 }; }
  sec.outlineId = A(sp, 'outlineShapeIDRef');
  return sec;
}

function headList(ctx, key, build) { if (!ctx.doc.lists[key]) ctx.doc.lists[key] = build(); return key; }
function paraProps(id, ctx) {
  const src = ctx.H.para[id] ?? {};
  const { _head, _bf, ...pPr } = src;
  if (_bf) {
    const bf = ctx.H.bf[_bf];
    if (bf) {
      const b = {};
      for (const [k, s] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']]) if (bf[s] && bf[s].style !== 'none') b[k] = { ...bf[s], space: 2 };
      if (Object.keys(b).length) pPr.border = b;
      if (bf.fill && bf.fill !== '#FFFFFF') pPr.shd = bf.fill;
    }
  }
  if (_head) {
    const lvl = Math.max(0, Math.min(8, _head.level));
    if (_head.type === 'OUTLINE') {
      pPr.outline = lvl;
      const nid = ctx.outlineId && ctx.H.numbering[ctx.outlineId] ? ctx.outlineId : null;
      const heads = nid ? ctx.H.numbering[nid].heads : null;
      if (heads && Object.values(heads).some((h) => h.text)) pPr.num = { id: headList(ctx, `hwpOutline${nid}`, () => listFromHeads(heads)), lvl };
    } else if (_head.type === 'NUMBER' && ctx.H.numbering[_head.id]) pPr.num = { id: headList(ctx, `hwpNum${_head.id}`, () => listFromHeads(ctx.H.numbering[_head.id].heads)), lvl };
    else if (_head.type === 'BULLET') pPr.num = { id: headList(ctx, `hwpBullet${_head.id}`, () => ({ levels: Array.from({ length: 9 }, (_, i) => ({ fmt: 'bullet', text: ctx.H.bullets[_head.id] ?? '●', start: 1, ind: { l: 0, hanging: 0 }, suff: 'space' })) })), lvl };
  }
  return pPr;
}
function listFromHeads(heads) {
  return { levels: Array.from({ length: 9 }, (_, i) => { const h = heads[i + 1] ?? { fmt: 'decimal', text: `%${i + 1}.`, start: 1 }; return { fmt: h.fmt, text: h.text || '', start: h.start ?? 1, ind: { l: 0, hanging: 0 }, suff: 'space' }; }) };
}

function readPara(p, ctx) {
  const out = [];
  const style = ctx.styleMap[A(p, 'styleIDRef')] ?? 'Normal';
  const pPr = diffProps(paraProps(A(p, 'paraPrIDRef'), ctx), ctx.doc.styles[style]?.pPr ?? {});
  const cur = () => ({ ...para(''), style, pPr: { ...pPr }, runs: [] });
  let pg = cur();
  if (A(p, 'pageBreak') === '1' && !ctx.firstPara) pg.pPr.pageBreakBefore = true;
  if (A(p, 'columnBreak') === '1') pg.runs.push({ obj: 'br', kind: 'column' });
  ctx.firstPara = false;
  const flush = (force = false) => { if (force || pg.runs.length) { pg.runs = normalizeRuns(pg.runs); out.push(pg); } pg = cur(); };
  for (const r of kids(p, 'run')) {
    const rp = diffProps(ctx.H.char[A(r, 'charPrIDRef')] ?? {}, ctx.doc.styles[style]?.rPr ?? {});
    for (const c of r.children) {
      switch (c.name) {
        case 't': readT(c, rp, pg.runs, ctx); break;
        case 'tbl': { flush(); const t = readTbl(c, ctx); if (t) out.push(t); break; }
        case 'pic': { const im = readPic(c, ctx); if (im) pg.runs.push(im); break; }
        case 'rect': case 'ellipse': case 'container': case 'polygon': { const s = readShape(c, ctx); if (s) pg.runs.push(s); break; }
        case 'equation': { const script = first(c, 'script')?.text ?? ''; pg.runs.push({ obj: 'eq', latex: hwpEqToLatex(script) }); break; }
        case 'ctrl': readCtrl(c, rp, pg.runs, ctx); break;
        case 'secPr': break;
        case 'switch': { const cs = child(c, 'case') ?? child(c, 'default'); for (const cc of cs?.children ?? []) if (cc.name === 'equation') pg.runs.push({ obj: 'eq', latex: hwpEqToLatex(first(cc, 'script')?.text ?? '') }); break; }
        default: break;
      }
    }
  }
  if (pg.runs.length || !out.length) flush(true);
  return out;
}
function readT(t, rp, runs, ctx) {
  // <hp:t> 안: 글자 + <hp:tab/> · <hp:lineBreak/> · <hp:fwSpace/> · <hp:nbSpace/> · <hp:hyphen/> · <hp:markpenBegin/> …
  if (!t.children.length) { if (t.text) runs.push({ ...rp, ...(ctx.link ? { link: ctx.link } : {}), text: t.text }); return; }
  // xml.js 는 섞인 글자를 요소의 text 로 모으므로, 자식 위치 정보 없이: 글자 먼저, 뒤에 특수 문자
  // 정확한 순서를 위해 원문을 다시 나눔
  const raw = t._raw;
  if (raw) {
    const re = /<hp:(tab|lineBreak|fwSpace|nbSpace|hyphen|markpenBegin|markpenEnd|titleMark)\b[^>]*?(?:\/>|>[\s\S]*?<\/hp:\1>)|([^<]+)/g;
    let m;
    let hl = null;
    while ((m = re.exec(raw))) {
      if (m[2]) { const txt = m[2].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'); if (txt) runs.push({ ...rp, ...(hl ? { hl } : {}), ...(ctx.link ? { link: ctx.link } : {}), text: txt }); continue; }
      const k = m[1];
      if (k === 'tab') runs.push({ ...rp, obj: 'tab' });
      else if (k === 'lineBreak') runs.push({ obj: 'br' });
      else if (k === 'fwSpace') runs.push({ ...rp, text: '\u3000' });
      else if (k === 'nbSpace') runs.push({ ...rp, text: '\u00A0' });
      else if (k === 'hyphen') runs.push({ ...rp, text: '-' });
      else if (k === 'markpenBegin') { const c = /color="(#[0-9A-Fa-f]{6})"/.exec(m[0]); hl = c ? c[1].toUpperCase() : '#FFFF00'; }
      else if (k === 'markpenEnd') hl = null;
    }
    return;
  }
  for (const c of t.children) if (c.name === 'tab') runs.push({ ...rp, obj: 'tab' }); else if (c.name === 'lineBreak') runs.push({ obj: 'br' });
  if (t.text) runs.push({ ...rp, text: t.text });
}
function readCtrl(c, rp, runs, ctx) {
  for (const k of c.children) {
    switch (k.name) {
      case 'header': case 'footer': {
        const blocks = readSubList(child(k, 'subList'), ctx);
        const which = A(k, 'applyPageType') === 'EVEN' ? 'even' : 'default';
        const prop = k.name === 'header' ? 'headers' : 'footers';
        ctx.hf[prop][which] = blocks.length ? blocks : [para('')];
        if (A(k, 'applyPageType') === 'ODD' && !ctx.hf[prop].default) ctx.hf[prop].default = ctx.hf[prop][which];
        if (A(k, 'applyPageType') === 'EVEN') ctx.doc.settings.evenOdd = true;
        break;
      }
      case 'footNote': case 'endNote': {
        const kind = k.name === 'footNote' ? 'foot' : 'end';
        const id = uid('n');
        const bl = readSubList(child(k, 'subList'), ctx);
        ctx.doc.notes[kind][id] = bl.length ? bl : [para('')];
        runs.push({ obj: 'note', id, kind });
        break;
      }
      case 'pageNum': {
        // 쪽 번호 위치 → 머리말/꼬리말에 PAGE 필드 (그 영역이 없을 때)
        const pos = A(k, 'pos') ?? 'BOTTOM_CENTER';
        const side = A(k, 'sideChar') ?? '';
        const prop = pos.startsWith('TOP') ? 'headers' : 'footers';
        const align = pos.endsWith('LEFT') ? 'left' : pos.endsWith('RIGHT') ? 'right' : 'center';
        if (pos !== 'NONE' && !ctx.hf[prop].default) ctx.hf[prop].default = [{ ...para(''), style: prop === 'headers' ? 'Header' : 'Footer', pPr: { align }, runs: [...(side ? [{ text: `${side} ` }] : []), { obj: 'field', code: 'PAGE', result: '1' }, ...(side ? [{ text: ` ${side}` }] : [])] }];
        ctx.pageNumFmt = A(k, 'formatType');
        break;
      }
      case 'fieldBegin': {
        const type = A(k, 'type');
        if (type === 'HYPERLINK') {
          const cmd = deep(k, 'stringParam').find((s) => A(s, 'name') === 'Command')?.text ?? A(k, 'command') ?? '';
          const url = cmd.split(';')[0].replace(/\\(.)/g, '$1');
          ctx.link = url.startsWith('#') || /^[^:]+$/.test(url) ? { anchor: url.replace(/^#/, '') } : { url };
        }
        break;
      }
      case 'fieldEnd': ctx.link = null; break;
      case 'bookmark': runs.push({ obj: 'mark', kind: 'bm', name: A(k, 'name') ?? '', edge: 'start' }, { obj: 'mark', kind: 'bm', name: A(k, 'name') ?? '', edge: 'end' }); break;
      case 'colPr': break;
      case 'autoNum': case 'newNum': break;
      default: break;
    }
  }
}
function readSubList(sl, ctx) {
  const out = [];
  const saved = ctx.firstPara;
  ctx.firstPara = true;
  for (const p of kids(sl, 'p')) for (const b of readPara(p, ctx)) out.push(b);
  ctx.firstPara = saved;
  return out;
}
function readTbl(t, ctx) {
  const rows = num(A(t, 'rowCnt'));
  const cols = num(A(t, 'colCnt'));
  const cells = [];
  for (const tr of kids(t, 'tr')) for (const tc of kids(tr, 'tc')) {
    const addr = child(tc, 'cellAddr'); const sp = child(tc, 'cellSpan'); const sz = child(tc, 'cellSz'); const mar = child(tc, 'cellMargin');
    const sl = child(tc, 'subList');
    const bf = ctx.H.bf[A(tc, 'borderFillIDRef')];
    const cell = { r: num(A(addr, 'rowAddr')), c: num(A(addr, 'colAddr')), cs: num(A(sp, 'colSpan'), 1), rs: num(A(sp, 'rowSpan'), 1), w: hu(A(sz, 'width')), h: hu(A(sz, 'height')), header: A(tc, 'header') === '1', blocks: readSubList(sl, ctx), va: A(sl, 'vertAlign'), bf, mar: mar && A(tc, 'hasMargin') === '1' ? { l: hu(A(mar, 'left')), r: hu(A(mar, 'right')), t: hu(A(mar, 'top')), b: hu(A(mar, 'bottom')) } : null };
    cells.push(cell);
  }
  const R = Math.max(rows, ...cells.map((c) => c.r + c.rs));
  const C = Math.max(cols, ...cells.map((c) => c.c + c.cs));
  if (!R || !C) return null;
  // 열 너비: 한 칸짜리 칸에서, 없으면 나눠서
  const grid = Array(C).fill(0);
  for (const c of cells) if (c.cs === 1 && !grid[c.c]) grid[c.c] = c.w;
  for (const c of cells) if (c.cs > 1) { const known = grid.slice(c.c, c.c + c.cs).reduce((a, b) => a + b, 0); const missing = grid.slice(c.c, c.c + c.cs).filter((g) => !g).length; if (missing) for (let k = c.c; k < c.c + c.cs; k++) if (!grid[k]) grid[k] = Math.max(8, (c.w - known) / missing); }
  for (let k = 0; k < C; k++) if (!grid[k]) grid[k] = 48;
  const at = new Map(cells.map((c) => [`${c.r},${c.c}`, c]));
  const cover = new Map(); // 'r,c' → 시작 칸 (세로 병합으로 덮인 자리)
  for (const c of cells) for (let r = c.r + 1; r < c.r + c.rs; r++) cover.set(`${r},${c.c}`, c);
  const inM = child(t, 'inMargin');
  const tbl = { type: 'tbl', id: uid('t'), style: 'TableNormal', tblPr: { look: { firstRow: false, firstCol: false, bandRows: false }, ...(inM ? { cellMar: { l: hu(A(inM, 'left')), r: hu(A(inM, 'right')), t: hu(A(inM, 'top')), b: hu(A(inM, 'bottom')) } } : {}) }, grid, rows: [] };
  const pos = child(t, 'pos');
  if (A(pos, 'horzAlign') === 'CENTER' && A(pos, 'treatAsChar') !== '1') tbl.tblPr.align = 'center';
  const repeat = A(t, 'repeatHeader') === '1';
  for (let r = 0; r < R; r++) {
    const row = { cells: [] };
    let hgt = 0;
    for (let c = 0; c < C;) {
      const cell = at.get(`${r},${c}`);
      const cov = cover.get(`${r},${c}`);
      if (cell) {
        const o = { blocks: cell.blocks.length ? cell.blocks : [para('')] };
        if (cell.cs > 1) o.span = cell.cs;
        if (cell.rs > 1) o.vMerge = 'restart';
        if (cell.bf) { o.borders = { top: cell.bf.t, bottom: cell.bf.b, left: cell.bf.l, right: cell.bf.r }; if (cell.bf.fill && cell.bf.fill !== '#FFFFFF') o.shd = cell.bf.fill; }
        if (cell.va === 'CENTER') o.vAlign = 'center'; else if (cell.va === 'BOTTOM') o.vAlign = 'bottom';
        if (cell.mar) o.mar = cell.mar;
        if (cell.header && repeat) row.header = true;
        if (cell.rs === 1) hgt = Math.max(hgt, cell.h);
        row.cells.push(o);
        c += cell.cs;
      } else if (cov) {
        const o = { blocks: [para('')], vMerge: 'cont' };
        if (cov.cs > 1) o.span = cov.cs;
        if (cov.bf) { o.borders = { top: cov.bf.t, bottom: cov.bf.b, left: cov.bf.l, right: cov.bf.r }; if (cov.bf.fill && cov.bf.fill !== '#FFFFFF') o.shd = cov.bf.fill; }
        row.cells.push(o);
        c += cov.cs;
      } else { row.cells.push({ blocks: [para('')] }); c++; }
    }
    if (hgt) { row.h = hgt; row.hRule = 'atLeast'; }
    tbl.rows.push(row);
  }
  return tbl;
}
function readPic(p, ctx) {
  const img = first(p, 'img');
  const ref = A(img, 'binaryItemIDRef');
  const media = ctx.binMedia(ref);
  const sz = child(p, 'sz') ?? child(p, 'curSz');
  const w = hu(A(sz, 'width')) || hu(A(child(p, 'curSz'), 'width')) || 100;
  const h = hu(A(sz, 'height')) || hu(A(child(p, 'curSz'), 'height')) || 100;
  const run = { obj: 'img', id: uid('i'), w, h, alt: first(p, 'shapeComment')?.text ?? '', ...(media ? { media } : { src: '' }) };
  const pos = child(p, 'pos');
  if (pos && A(pos, 'treatAsChar') !== '1') run.float = floatOf(p, pos);
  const clip = child(p, 'imgClip'); const dim = child(p, 'imgDim');
  if (clip && dim) { const dw = num(A(dim, 'dimwidth')); const dh = num(A(dim, 'dimheight')); if (dw && dh) { const c = { l: num(A(clip, 'left')) / dw, t: num(A(clip, 'top')) / dh, r: 1 - num(A(clip, 'right')) / dw, b: 1 - num(A(clip, 'bottom')) / dh }; if (c.l > 0.001 || c.t > 0.001 || c.r > 0.001 || c.b > 0.001) run.crop = c; } }
  return run;
}
function floatOf(el, pos) {
  const wrap = { SQUARE: 'square', TIGHT: 'tight', THROUGH: 'tight', TOP_AND_BOTTOM: 'topBottom', BEHIND_TEXT: 'behind', IN_FRONT_OF_TEXT: 'front' }[A(el, 'textWrap')] ?? 'square';
  const relH = { PAPER: 'page', PAGE: 'margin', COLUMN: 'column', PARA: 'column' }[A(pos, 'horzRelTo')] ?? 'column';
  const relV = { PAPER: 'page', PAGE: 'margin', PARA: 'paragraph' }[A(pos, 'vertRelTo')] ?? 'paragraph';
  const f = { wrap, relH, relV, x: hu(A(pos, 'horzOffset')), y: hu(A(pos, 'vertOffset')) };
  const ha = A(pos, 'horzAlign'); const va = A(pos, 'vertAlign');
  if (ha && ha !== 'LEFT') f.alignH = ha === 'CENTER' ? 'center' : ha === 'RIGHT' ? 'right' : undefined;
  if (va && va !== 'TOP' && relV !== 'paragraph') f.alignV = va === 'CENTER' ? 'center' : va === 'BOTTOM' ? 'bottom' : undefined;
  if (wrap === 'square' || wrap === 'tight') f.side = ha === 'RIGHT' ? 'right' : 'left';
  return f;
}
function readShape(s, ctx) {
  const dt = child(s, 'drawText');
  const sz = child(s, 'sz');
  const w = hu(A(sz, 'width')) || 100; const h = hu(A(sz, 'height')) || 60;
  const ls = child(s, 'lineShape');
  const fill = first(s, 'winBrush');
  const run = { obj: 'shape', id: uid('s'), w, h, prst: s.name === 'ellipse' ? 'ellipse' : 'rect', blocks: dt ? readSubList(child(dt, 'subList'), ctx) : [] };
  const fc = colorOf(A(fill, 'faceColor'));
  if (fc && A(fill, 'alpha') !== '255') run.fill = fc;
  if (ls && A(ls, 'style') !== 'NONE') run.line = { w: Math.max(0.25, hu(A(ls, 'width')) * 0.75), color: colorOf(A(ls, 'color')) ?? '#000000' };
  const pos = child(s, 'pos');
  run.float = pos && A(pos, 'treatAsChar') !== '1' ? floatOf(s, pos) : { wrap: 'topBottom', relH: 'column', relV: 'paragraph', x: 0, y: 0 };
  if (A(child(dt, 'subList'), 'vertAlign') === 'CENTER') run.anchor = 'ctr';
  if (!run.blocks.length && !run.fill && !run.line) return null;
  return run;
}
/** 한글 수식 스크립트 → LaTeX (자주 쓰는 것만) */
export function hwpEqToLatex(s) {
  let t = ` ${s} `;
  t = t.replace(/\{([^{}]*)\}\s*over\s*\{([^{}]*)\}/gi, '\\frac{$1}{$2}');
  t = t.replace(/(\S+)\s+over\s+(\S+)/gi, '\\frac{$1}{$2}');
  t = t.replace(/\bsqrt\s*\{/gi, '\\sqrt{').replace(/\bsqrt\s+(\S+)/gi, '\\sqrt{$1}');
  t = t.replace(/\b(sum|int|prod|oint|lim)\b/gi, (m) => `\\${m.toLowerCase()}`);
  t = t.replace(/\b(alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|phi|omega|infinity|inf)\b/gi, (m) => (m.toLowerCase().startsWith('inf') ? '\\infty' : `\\${m.toLowerCase()}`));
  t = t.replace(/\btimes\b/gi, '\\times').replace(/\bdiv\b/gi, '\\div').replace(/\bleq\b|\ble\b/gi, '\\le').replace(/\bgeq\b|\bge\b/gi, '\\ge').replace(/\bpm\b/gi, '\\pm').replace(/\bcdot\b/gi, '\\cdot');
  t = t.replace(/\b(sin|cos|tan|log|ln|exp)\b/g, '\\$1');
  t = t.replace(/\bLEFT\b|\bRIGHT\b/g, '').replace(/~/g, ' ').replace(/`/g, ' ');
  return t.replace(/\s+/g, ' ').trim();
}
function diffProps(full, base) {
  const out = {};
  for (const [k, v] of Object.entries(full)) if (JSON.stringify(v) !== JSON.stringify(base?.[k])) out[k] = v;
  return out;
}

export function readHwpx(input) {
  const files = input instanceof Uint8Array ? unzip(input) : input;
  const warnings = [];
  const txt = (p) => (files[p] ? dec.decode(files[p]) : null);
  if (files['META-INF/manifest.xml'] && /encryption-data/.test(txt('META-INF/manifest.xml') ?? '')) throw new Error('암호로 보호된 한글 문서입니다. 한글에서 암호를 풀고 저장한 뒤 열어 주세요.');
  // 매니페스트 (content.hpf): 그림 · 구역 순서
  const hpfPath = Object.keys(files).find((k) => /content\.hpf$/i.test(k)) ?? 'Contents/content.hpf';
  const hpf = txt(hpfPath);
  const items = {};
  let spine = [];
  if (hpf) {
    const x = parseXml(hpf);
    for (const it of deep(x, 'item')) items[A(it, 'id')] = A(it, 'href');
    spine = deep(x, 'itemref').map((r) => items[A(r, 'idref')]).filter((h) => h && /section\d+\.xml$/i.test(h));
  }
  if (!spine.length) spine = Object.keys(files).filter((k) => /Contents\/section\d+\.xml$/i.test(k)).sort((a, b) => Number(/(\d+)\.xml$/.exec(a)[1]) - Number(/(\d+)\.xml$/.exec(b)[1]));
  if (!spine.length) throw new Error('한글 문서(.hwpx)의 본문(Contents/section0.xml)이 없습니다.');
  const H = readHead(txt(Object.keys(files).find((k) => /Contents\/header\.xml$/i.test(k)) ?? 'Contents/header.xml'));
  const doc = newDoc();
  doc.media = {};
  doc.lists = {};
  // 스타일: 바탕글 = 표준, 개요 N = 제목 N (탐색 창 · 목차), 나머지 이름 그대로
  const styleMap = {};
  const st = doc.styles;
  H.styles.forEach((s, id) => {
    if (!s || s.type === 'CHAR') return;
    const outline = /^개요\s*(\d)$/.exec(s.name ?? '') ?? /^Outline\s*(\d)$/i.exec(s.eng ?? '');
    const key = id === 0 || s.eng === 'Normal' ? 'Normal' : outline ? `Heading${outline[1]}` : `hwp${id}`;
    styleMap[String(id)] = key;
    const pPr = paraProps(s.para, { H, doc, outlineId: null });
    delete pPr.num;
    const rPr = { ...(H.char[s.char] ?? {}) };
    st[key] = { ...(st[key] ?? {}), id: key, type: 'para', name: s.name, basedOn: key === 'Normal' ? undefined : 'Normal', quick: key === 'Normal' || !!outline || /^본문|^제목/.test(s.name ?? ''), order: key === 'Normal' ? 0 : outline ? 9 + Number(outline[1]) : 100 + id, pPr, rPr };
    if (key === 'Normal') { doc.defaults = { rPr: { ...doc.defaults.rPr, ...rPr, font: rPr.font ?? doc.defaults.rPr.font, fontEa: rPr.fontEa ?? doc.defaults.rPr.fontEa }, pPr: { ...pPr } }; st.Normal.pPr = {}; st.Normal.rPr = {}; }
  });
  // 스타일 서식은 바탕글과의 차이만
  for (const [k, s] of Object.entries(st)) if (k !== 'Normal' && (k.startsWith('hwp') || /^Heading\d$/.test(k)) && s.pPr && H.styles.length) { s.pPr = diffProps(s.pPr, doc.defaults.pPr); s.rPr = diffProps(s.rPr ?? {}, doc.defaults.rPr); }
  const mediaIds = new Map();
  const binMedia = (ref) => {
    if (!ref) return null;
    if (mediaIds.has(ref)) return mediaIds.get(ref);
    let href = items[ref];
    if (!href) href = Object.keys(files).find((k) => k.startsWith('BinData/') && k.replace(/^BinData\//, '').replace(/\.[^.]+$/, '') === ref);
    const bytes = href && files[href];
    if (!bytes) return null;
    const id = uid('m');
    doc.media[id] = mediaDataUrl(bytes, href);
    mediaIds.set(ref, id);
    return id;
  };
  const body = [];
  let lastSec = null;
  spine.forEach((path, si) => {
    const ctx = { H, doc, styleMap, binMedia, hf: { headers: {}, footers: {} }, firstPara: true, outlineId: null, link: null };
    // 개요 번호: secPr 의 outlineShapeIDRef
    const raw = txt(path);
    const om = /outlineShapeIDRef="(\d+)"/.exec(raw);
    ctx.outlineId = om?.[1] ?? null;
    const x = parseXml(raw, true);
    // <hp:t> 원문 (안에 섞인 탭 · 줄바꿈 순서를 지키려고)
    for (const t of deep(x, 't')) if (t.children.length && t.s != null && t.e != null) t._raw = raw.slice(raw.indexOf('>', t.s) + 1, raw.lastIndexOf('</', t.e - 1));
    const blocks = [];
    let sec = null;
    for (const p of kids(x, 'p')) {
      for (const b of readPara(p, ctx)) blocks.push(b);
      if (!sec) { const sp = first(p, 'secPr'); if (sp) sec = readSecPr(sp, ctx, p); }
    }
    sec = sec ?? newSection();
    sec.headers = ctx.hf.headers; sec.footers = ctx.hf.footers;
    if (ctx.pageNumFmt && ctx.pageNumFmt !== 'DIGIT') sec.pgNum = { ...(sec.pgNum ?? {}), fmt: { ROMAN_CAPITAL: 'upperRoman', ROMAN_SMALL: 'lowerRoman', LATIN_CAPITAL: 'upperLetter', LATIN_SMALL: 'lowerLetter', HANGUL_SYLLABLE: 'ganada' }[ctx.pageNumFmt] ?? 'decimal' };
    delete sec.outlineId;
    if (!blocks.length) blocks.push(para(''));
    if (lastSec) { const prev = body[body.length - 1]; if (prev?.type === 'p') prev.secBreak = { kind: 'nextPage', sec: lastSec }; else body.push({ ...para(''), secBreak: { kind: 'nextPage', sec: lastSec } }); }
    body.push(...blocks);
    lastSec = sec;
  });
  doc.body = body;
  doc.section = lastSec ?? newSection();
  // 문서 정보
  if (hpf) {
    const x = parseXml(hpf);
    const t = first(x, 'title')?.text;
    if (t) doc.props.title = t;
    for (const m of deep(x, 'meta')) { const n = A(m, 'name'); if (n === 'creator' && m.text && m.text !== 'text') doc.props.author = m.text; if (n === 'subject' && m.text) doc.props.subject = m.text; }
  }
  return { doc, warnings };
}

// ════════════════════════════════════ 쓰기 ════════════════════════════════════
const NS = 'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>';
const mm = (pt) => `${Math.max(0.1, Math.round(((pt ?? 0.5) * 25.4 / 72) * 100) / 100)} mm`;
const hexc = (c) => (c ? (c.startsWith('#') ? c.toUpperCase().slice(0, 7) : `#${c}`) : '#000000');

class Pool {
  constructor() { this.list = []; this.map = new Map(); }
  id(obj) { const k = JSON.stringify(obj); if (!this.map.has(k)) { this.map.set(k, this.list.length); this.list.push(obj); } return this.map.get(k); }
}

/** 문서 → { 경로: 문자열|바이트 } (mimetype 이 첫 항목) */
export function writeHwpx(doc, opts = {}) {
  const fonts = new Pool(); fonts.id('함초롬바탕');
  const bfs = new Pool(); bfs.id({}); // id 1 = 테두리 없음
  const bfThin = bfs.id({ l: 1, r: 1, t: 1, b: 1 });
  const chars = new Pool();
  const paras = new Pool();
  const nums = new Pool();
  const bullets = new Pool();
  const bins = [];
  const binOf = new Map();
  const numbering = computeNumbering(doc);
  const fontName = (f, ea) => { const t = doc.themeFonts ?? {}; if (f === '+mj') return (ea ? t.major?.ea : t.major?.latin) ?? '맑은 고딕'; if (f === '+mn' || !f) return (ea ? t.minor?.ea : t.minor?.latin) ?? '맑은 고딕'; return f; };
  const charId = (rp) => {
    const o = { sz: rp.sz ?? 10, ea: fonts.id(fontName(rp.fontEa ?? rp.font, true)), lat: fonts.id(fontName(rp.font ?? rp.fontEa, false)), b: !!rp.b, i: !!rp.i, u: rp.u ? (rp.u === true ? 'SOLID' : rp.u === 'double' ? 'DOUBLE_SLIM' : rp.u === 'dotted' ? 'DOT' : rp.u === 'wave' ? 'WAVE' : 'SOLID') : null, s: !!(rp.s || rp.dstrike), color: rp.color ?? '#000000', shd: rp.shd ?? null, va: rp.va ?? null, scale: rp.scale ?? 100, spc: rp.spc ? Math.round((rp.spc / (rp.sz ?? 10)) * 100) : 0, outline: !!rp.outline, shadow: !!rp.shadow, emboss: !!rp.emboss, imprint: !!rp.imprint, pos: rp.pos ? Math.round((-rp.pos / (rp.sz ?? 10)) * 100) : 0 };
    return chars.id(o);
  };
  const paraId = (ep, numInfo, mainSz) => {
    const ind = ep.ind ?? {};
    const spc = ep.spc ?? {};
    const hanging = ind.hanging ?? (ind.first != null && ind.first < 0 ? -ind.first : 0);
    const left = (ind.l ?? 0) - hanging;
    const intent = hanging ? -hanging : ind.first ?? 0;
    let ls;
    if (spc.rule === 'exact') ls = { type: 'FIXED', v: toHU(spc.line) };
    else if (spc.rule === 'atLeast') ls = { type: 'AT_LEAST', v: toHU(spc.line) };
    else if (spc.rule === 'pct') ls = { type: 'PERCENT', v: Math.round((spc.line ?? 1.6) * 100) };
    else ls = { type: 'PERCENT', v: Math.round((spc.line ?? 1) * 100 * lineFactor(fontName(ep._font, true)) * 0.95) };
    const o = { align: ALIGN_OUT[ep.align] ?? 'JUSTIFY', left: toHU(left), right: toHU(ind.r ?? 0), intent: toHU(intent), prev: toHU(spc.before ?? 0), next: toHU(spc.after ?? 0), ls, keepNext: !!ep.keepNext, keepLines: !!ep.keepLines, pbb: !!ep.pageBreakBefore, widow: ep.widow !== false, head: numInfo, bf: 1, tabs: null };
    if (ep.border || ep.shd) {
      const b = ep.border ?? {};
      o.bf = bfs.id({ l: b.left, r: b.right, t: b.top, b: b.bottom, fill: ep.shd ?? null, para: true }) + 1;
    }
    if (ep.outline != null && ep.outline < 9 && !numInfo) o.head = { type: 'OUTLINE', id: 1, level: ep.outline };
    return paras.id(o);
  };
  const numId = (numId0, lvl) => {
    const def = doc.lists?.[numId0];
    if (!def) return null;
    const L = def.levels?.[lvl] ?? def.levels?.[0];
    if (L?.fmt === 'bullet') return { type: 'BULLET', id: bullets.id(L.text || '●') + 1, level: lvl };
    const heads = (def.levels ?? []).slice(0, 7).map((x, i) => ({ fmt: NUMFMT_OUT[x.fmt] ?? 'DIGIT', text: (x.text ?? `%${i + 1}.`).replace(/%(\d)/g, '^$1'), start: def.override?.[i]?.start ?? x.start ?? 1 }));
    return { type: 'NUMBER', id: nums.id(heads) + 2, level: lvl }; // numbering 1 = 개요용 (빈 서식)
  };
  const binFor = (r) => {
    const src = r.media ? doc.media?.[r.media] : r.src;
    if (!src || !src.startsWith('data:')) return null;
    if (binOf.has(src)) return binOf.get(src);
    const d = dataUrlBytes(src);
    if (!d) return null;
    const id = `image${bins.length + 1}`;
    bins.push({ id, name: `BinData/${id}.${d.ext}`, bytes: d.bytes, mime: d.mime });
    binOf.set(src, id);
    return id;
  };
  let objId = 1000;
  const ctxBreak = { page: false, column: false };
  let tblN = 0;
  // ── 문단 · 표
  const paraXml = (p, { secXml = '', tblStyle = null, cellRun = null } = {}) => {
    const ep = effPara(doc, p, tblStyle);
    const firstRun = p.runs.find((r) => !r.obj) ?? { ...(p.mark ?? {}), text: '' };
    const mainRp = { ...effRun(doc, p, firstRun, tblStyle), ...(cellRun ?? {}) };
    const nm = numbering.get(p.id);
    const head = nm ? numId(nm.numId, nm.lvl) : null;
    // 목록 들여쓰기 (문단에 직접 값이 없으면 목록 수준 값)
    const ep2 = { ...ep, _font: mainRp.fontEa ?? mainRp.font };
    if (nm?.ind && p.pPr?.ind?.l == null) ep2.ind = { ...ep.ind, l: nm.ind.l, hanging: nm.ind.hanging ?? 0 };
    const pid = paraId(ep2, head, mainRp.sz);
    const pb = ep.pageBreakBefore || ctxBreak.page;
    const cb = ctxBreak.column;
    ctxBreak.page = false; ctxBreak.column = false;
    let s = `<hp:p id="0" paraPrIDRef="${pid}" styleIDRef="${styleIdx(p.style)}" pageBreak="${pb ? 1 : 0}" columnBreak="${cb ? 1 : 0}" merged="0">`;
    let open = null;
    const closeRun = () => { if (open != null) { s += '</hp:run>'; open = null; } };
    const runOpen = (cid) => { if (open !== cid) { closeRun(); s += `<hp:run charPrIDRef="${cid}">`; open = cid; } };
    if (secXml) { runOpen(charId(mainRp)); s += secXml; }
    let text = '';
    let tcid = null;
    const flushT = () => { if (text) { s += `<hp:t>${text}</hp:t>`; text = ''; } };
    for (const r of p.runs) {
      if (r.del) continue;
      const rp = { ...effRun(doc, p, r, tblStyle), ...(cellRun ?? {}) };
      const cid = charId(rp);
      if (cid !== tcid) { flushT(); runOpen(cid); tcid = cid; }
      if (!r.obj) {
        let t = esc(r.text ?? '').replace(/\u00A0/g, '<hp:nbSpace/>');
        if (r.hl || (rp.hl && !rp.shd)) { const hc = doc.styles && rp.hl ? ({ yellow: '#FFFF00', brightGreen: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', red: '#FF0000', blue: '#0000FF' }[rp.hl] ?? (rp.hl.startsWith('#') ? rp.hl : '#FFFF00')) : '#FFFF00'; t = `<hp:markpenBegin color="${hc}"/>${t}<hp:markpenEnd/>`; }
        if (r.link?.url) { flushT(); s += `<hp:ctrl><hp:fieldBegin id="${++objId}" type="HYPERLINK" name="" editable="0" dirty="1" zorder="-1" fieldid="${objId}"><hp:parameters cnt="1" name=""><hp:stringParam name="Command">${esc(r.link.url.replace(/[:;\\]/g, '\\$&'))};1;0;0;</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl>`; text += t; flushT(); s += `<hp:ctrl><hp:fieldEnd beginIDRef="${objId}" fieldid="${objId}"/></hp:ctrl>`; continue; }
        text += t;
        continue;
      }
      switch (r.obj) {
        case 'tab': text += '<hp:tab width="4000" leader="0" type="1"/>'; break;
        case 'br': if (r.kind === 'page') ctxBreak.page = true; else if (r.kind === 'column') ctxBreak.column = true; else text += '<hp:lineBreak/>'; break;
        case 'sym': text += esc(r.char ?? ''); break;
        case 'field': {
          const code = (r.code ?? '').trim().split(/\s+/)[0].toUpperCase();
          if (code === 'PAGE') { flushT(); s += `<hp:ctrl><hp:autoNum num="1" numType="PAGE"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="0"/></hp:autoNum></hp:ctrl>`; }
          else if (code === 'NUMPAGES') { flushT(); s += `<hp:ctrl><hp:autoNum num="1" numType="TOTAL_PAGE"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="0"/></hp:autoNum></hp:ctrl>`; }
          else text += esc(r.result ?? '');
          break;
        }
        case 'note': {
          flushT();
          const kind = r.kind === 'end' ? 'endNote' : 'footNote';
          const bl = (r.kind === 'end' ? doc.notes?.end : doc.notes?.foot)?.[r.id] ?? [para('')];
          s += `<hp:ctrl><hp:${kind} number="0" suffixChar="41" instId="${++objId}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${blocksXml(bl)}</hp:subList></hp:${kind}></hp:ctrl>`;
          break;
        }
        case 'img': { flushT(); const x = picXml(r); if (x) s += x; break; }
        case 'shape': { flushT(); s += shapeXml(r); break; }
        case 'eq': { flushT(); s += eqXml(r, rp); break; }
        case 'mark': if (r.kind === 'bm' && r.edge === 'start') { flushT(); s += `<hp:ctrl><hp:bookmark name="${esc(r.name ?? '')}"/></hp:ctrl>`; } break;
        default: break;
      }
    }
    flushT();
    if (open == null) runOpen(charId(mainRp));
    if (open != null && !/<hp:t>|<hp:ctrl>|<hp:pic|<hp:rect|<hp:equation/.test(s.slice(s.lastIndexOf('<hp:run')))) s += '<hp:t/>';
    closeRun();
    return `${s}</hp:p>`;
  };
  const posXml = (r, inline) => {
    if (inline || !r.float) return `<hp:sz width="${toHU(r.w)}" widthRelTo="ABSOLUTE" height="${toHU(r.h)}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/>`;
    const f = r.float;
    const hr = { page: 'PAPER', margin: 'PAGE', column: 'COLUMN' }[f.relH] ?? 'COLUMN';
    const vr = { page: 'PAPER', margin: 'PAGE', paragraph: 'PARA' }[f.relV] ?? 'PARA';
    return `<hp:sz width="${toHU(r.w)}" widthRelTo="ABSOLUTE" height="${toHU(r.h)}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="${vr}" horzRelTo="${hr}" vertAlign="${f.alignV === 'center' ? 'CENTER' : f.alignV === 'bottom' ? 'BOTTOM' : 'TOP'}" horzAlign="${f.alignH === 'center' ? 'CENTER' : f.alignH === 'right' ? 'RIGHT' : 'LEFT'}" vertOffset="${toHU(f.y)}" horzOffset="${toHU(f.x)}"/><hp:outMargin left="0" right="0" top="0" bottom="0"/>`;
  };
  const wrapAttr = (r) => { const w = r.float?.wrap; return { square: 'SQUARE', tight: 'TIGHT', topBottom: 'TOP_AND_BOTTOM', behind: 'BEHIND_TEXT', front: 'IN_FRONT_OF_TEXT' }[w] ?? 'TOP_AND_BOTTOM'; };
  const picXml = (r) => {
    const bin = r.svg ? null : binFor(r);
    if (!bin) return '';
    const id = ++objId;
    const w = toHU(r.w); const h = toHU(r.h);
    const c = r.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const dw = Math.round(w / Math.max(0.01, 1 - c.l - c.r)); const dh = Math.round(h / Math.max(0.01, 1 - c.t - c.b));
    return `<hp:pic id="${id}" zOrder="0" numberingType="PICTURE" textWrap="${wrapAttr(r)}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${id}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${dw}" height="${dh}"/><hp:curSz width="${w}" height="${h}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="${Math.round(r.rot ?? 0)}" centerX="${Math.floor(w / 2)}" centerY="${Math.floor(h / 2)}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${w}" y="0"/><hc:pt2 x="${w}" y="${h}"/><hc:pt3 x="0" y="${h}"/></hp:imgRect><hp:imgClip left="${Math.round(c.l * dw)}" right="${Math.round((1 - c.r) * dw)}" top="${Math.round(c.t * dh)}" bottom="${Math.round((1 - c.b) * dh)}"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="${dw}" dimheight="${dh}"/><hc:img binaryItemIDRef="${bin}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:effects/>${posXml(r, !r.float)}${r.alt ? `<hp:shapeComment>${esc(r.alt)}</hp:shapeComment>` : ''}</hp:pic>`;
  };
  const shapeXml = (r) => {
    const id = ++objId;
    const w = toHU(r.w); const h = toHU(r.h);
    const fill = r.fill ? `<hc:fillBrush><hc:winBrush faceColor="${hexc(r.fill)}" hatchColor="#000000" alpha="0"/></hc:fillBrush>` : '';
    const ln = r.line ? `<hp:lineShape color="${hexc(r.line.color)}" width="${toHU(((r.line.w ?? 0.75) * 4) / 3)}" style="SOLID" endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL" headfill="1" tailfill="1" headSz="MEDIUM_MEDIUM" tailSz="MEDIUM_MEDIUM" outlineStyle="NORMAL" alpha="0"/>` : '<hp:lineShape color="#000000" width="0" style="NONE" endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL" headfill="1" tailfill="1" headSz="MEDIUM_MEDIUM" tailSz="MEDIUM_MEDIUM" outlineStyle="NORMAL" alpha="0"/>';
    const tag = r.prst === 'ellipse' ? 'ellipse' : 'rect';
    const ins = r.inset ?? { l: 9.6, r: 9.6, t: 4.8, b: 4.8 };
    const geo = tag === 'rect' ? `<hc:pt0 x="0" y="0"/><hc:pt1 x="${w}" y="0"/><hc:pt2 x="${w}" y="${h}"/><hc:pt3 x="0" y="${h}"/>` : `<hc:center x="${w / 2}" y="${h / 2}"/><hc:ax1 x="${w}" y="${h / 2}"/><hc:ax2 x="${w / 2}" y="0"/><hc:start1 x="0" y="0"/><hc:end1 x="0" y="0"/><hc:start2 x="0" y="0"/><hc:end2 x="0" y="0"/>`;
    return `<hp:${tag} id="${id}" zOrder="1" numberingType="PICTURE" textWrap="${wrapAttr({ float: r.float ?? { wrap: 'front' } })}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${id}"${tag === 'rect' ? ' ratio="0"' : ' intervalDirty="0" hasArcPr="0" arcType="NORMAL"'}><hp:offset x="0" y="0"/><hp:orgSz width="${w}" height="${h}"/><hp:curSz width="${w}" height="${h}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="${Math.floor(w / 2)}" centerY="${Math.floor(h / 2)}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>${ln}${fill}<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0"/><hp:drawText lastWidth="${w}" name="" editable="0"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${r.anchor === 'ctr' ? 'CENTER' : 'TOP'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${blocksXml(r.blocks?.length ? r.blocks : [para('')])}</hp:subList><hp:textMargin left="${toHU(ins.l)}" right="${toHU(ins.r)}" top="${toHU(ins.t)}" bottom="${toHU(ins.b)}"/></hp:drawText>${geo}${posXml(r, false)}</hp:${tag}>`;
  };
  const eqXml = (r, rp) => {
    const id = ++objId;
    const script = latexToHwpEq(r.latex ?? '');
    const w = Math.max(1000, script.length * 600); const h = 1500;
    return `<hp:equation id="${id}" zOrder="0" numberingType="EQUATION" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" version="Equation Version 60" baseLine="85" textColor="#000000" baseUnit="${Math.round((rp.sz ?? 10) * 100)}" lineMode="CHAR" font="HancomEQN"><hp:sz width="${w}" widthRelTo="ABSOLUTE" height="${h}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="56" right="56" top="0" bottom="0"/><hp:shapeComment>수식입니다.</hp:shapeComment><hp:script>${esc(script)}</hp:script></hp:equation>`;
  };
  const tableXml = (t, secXml = '') => {
    const id = ++objId; tblN++;
    const look = (doc.styles?.[t.style]?.tblPr) ?? {};
    const tb = { ...(look.borders ?? {}), ...(t.tblPr?.borders ?? {}) };
    const R = t.rows.length; const C = t.grid.length;
    // 칸 위치 계산 (세로 병합 묶음)
    const pos = t.rows.map((row) => { let gc = 0; return row.cells.map((c) => { const o = { gc, span: c.span ?? 1 }; gc += o.span; return o; }); });
    const rowSpan = (r, ci) => { const gc = pos[r][ci].gc; let n = 1; for (let k = r + 1; k < R; k++) { const j = pos[k].findIndex((x) => x.gc === gc); if (j < 0 || t.rows[k].cells[j].vMerge !== 'cont') break; n++; } return n; };
    const mar = { l: 7.2, r: 7.2, t: 0, b: 0, ...(t.tblPr?.cellMar ?? {}) };
    let trs = '';
    let totalH = 0;
    for (let r = 0; r < R; r++) {
      const row = t.rows[r];
      const h = toHU(row.h ?? 24);
      totalH += h;
      let tcs = '';
      row.cells.forEach((c, ci) => {
        if (c.vMerge === 'cont') return;
        const { gc, span } = pos[r][ci];
        const rs = c.vMerge === 'restart' ? rowSpan(r, ci) : 1;
        const w = t.grid.slice(gc, gc + span).reduce((a, b) => a + b, 0);
        const side = (k, outer, inner) => c.borders?.[k] ?? (outer ? tb[k] : tb[inner]);
        const bdr = { t: side('top', r === 0, 'insideH'), b: side('bottom', r + rs >= R, 'insideH'), l: side('left', gc === 0, 'insideV'), r: side('right', gc + span >= C, 'insideV') };
        const bfid = bfs.id({ l: bdr.l, r: bdr.r, t: bdr.t, b: bdr.b, fill: c.shd ?? null }) + 1;
        const cm = { ...mar, ...(c.mar ?? {}) };
        tcs += `<hp:tc name="" header="${row.header ? 1 : 0}" hasMargin="${c.mar ? 1 : 0}" protect="0" editable="0" dirty="0" borderFillIDRef="${bfid}"><hp:subList id="" textDirection="${c.dir === 'vert' ? 'VERTICAL' : 'HORIZONTAL'}" lineWrap="BREAK" vertAlign="${c.vAlign === 'center' ? 'CENTER' : c.vAlign === 'bottom' ? 'BOTTOM' : 'TOP'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${blocksXml(c.blocks, { tblStyle: t.style })}</hp:subList><hp:cellAddr colAddr="${gc}" rowAddr="${r}"/><hp:cellSpan colSpan="${span}" rowSpan="${rs}"/><hp:cellSz width="${toHU(w)}" height="${h}"/><hp:cellMargin left="${toHU(cm.l)}" right="${toHU(cm.r)}" top="${toHU(cm.t)}" bottom="${toHU(cm.b)}"/></hp:tc>`;
      });
      trs += `<hp:tr>${tcs}</hp:tr>`;
    }
    const W = toHU(t.grid.reduce((a, b) => a + b, 0));
    const center = t.tblPr?.align === 'center';
    const tbl = `<hp:tbl id="${id}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="${t.rows[0]?.header ? 1 : 0}" rowCnt="${R}" colCnt="${C}" cellSpacing="0" borderFillIDRef="${bfThin + 1}" noShading="0"><hp:sz width="${W}" widthRelTo="ABSOLUTE" height="${totalH}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="${center ? 'CENTER' : 'LEFT'}" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="${toHU(6)}"/><hp:inMargin left="${toHU(mar.l)}" right="${toHU(mar.r)}" top="${toHU(mar.t)}" bottom="${toHU(mar.b)}"/>${trs}</hp:tbl>`;
    const pid = paraId({ align: center ? 'center' : 'left', spc: { before: 0, after: 0, line: 1, rule: 'pct' } }, null, 10);
    return `<hp:p id="0" paraPrIDRef="${pid}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="${charId(doc.defaults?.rPr ?? {})}">${secXml}${tbl}<hp:t/></hp:run></hp:p>`;
  };
  const blocksXml = (bl, o = {}) => {
    let s = '';
    for (const b of bl ?? []) s += b.type === 'tbl' ? tableXml(b) : paraXml(b, o);
    return s;
  };
  const styleIdx = (id) => (id === 'Normal' ? 0 : styleOrder.indexOf(id) >= 0 ? styleOrder.indexOf(id) : 0);
  // 스타일: 바탕글 + 문서에서 쓰는 단락 스타일
  const styleOrder = ['Normal'];
  walkParas(doc.body, (p) => { if (p.style && !styleOrder.includes(p.style) && doc.styles[p.style]?.type === 'para') styleOrder.push(p.style); });
  // ── 구역 나누기 기준으로 sectionN.xml
  const sects = [];
  let cur = [];
  for (const b of doc.body) { cur.push(b); if (b.type === 'p' && b.secBreak) { sects.push({ blocks: cur, sec: b.secBreak.sec }); cur = []; } }
  sects.push({ blocks: cur.length ? cur : [para('')], sec: doc.section });
  const secFiles = sects.map(({ blocks, sec }) => {
    const m = sec.margin;
    const land = sec.page.orient === 'landscape';
    // 한글 여백: top = 머리말까지, header = 머리말 영역 → Word 의 본문 위치와 같게
    const hdrH = Math.max(0, (m.t ?? 96) - (m.header ?? 48));
    const ftrH = Math.max(0, (m.b ?? 96) - (m.footer ?? 48));
    let hf = '';
    for (const [prop, tag] of [['headers', 'header'], ['footers', 'footer']]) for (const [which, apply] of [['default', 'BOTH'], ['even', 'EVEN']]) { const bl = sec[prop]?.[which]; if (bl) hf += `<hp:ctrl><hp:${tag} id="${++objId}" applyPageType="${apply}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${tag === 'header' ? 'TOP' : 'BOTTOM'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="${toHU(sec.page.w - m.l - m.r)}" textHeight="${toHU(tag === 'header' ? hdrH : ftrH)}" hasTextRef="0" hasNumRef="0">${blocksXml(bl)}</hp:subList></hp:${tag}></hp:ctrl>`; }
    const colN = sec.cols?.n ?? 1;
    const secXml = `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="${toHU(sec.cols?.space ?? 0)}" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="${sec.pgNum?.start ?? 0}" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="${sec.titlePg ? 1 : 0}" hideFirstFooter="${sec.titlePg ? 1 : 0}" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${land ? 'NARROWLY' : 'WIDELY'}" width="${toHU(land ? Math.min(sec.page.w, sec.page.h) : sec.page.w)}" height="${toHU(land ? Math.max(sec.page.w, sec.page.h) : sec.page.h)}" gutterType="LEFT_ONLY"><hp:margin header="${toHU(hdrH)}" footer="${toHU(ftrH)}" gutter="${toHU(m.gutter ?? 0)}" left="${toHU(m.l)}" right="${toHU(m.r)}" top="${toHU(m.header ?? 48)}" bottom="${toHU(m.footer ?? 48)}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="${sec.pageBorder ? bfs.id({ l: sec.pageBorder, r: sec.pageBorder, t: sec.pageBorder, b: sec.pageBorder }) + 1 : 1}" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="${colN}" sameSz="1" sameGap="${colN > 1 ? toHU(sec.cols.space) : 0}"/></hp:ctrl>${hf}`;
    let body = '';
    blocks.forEach((b, i) => {
      if (i === 0) body += b.type === 'tbl' ? tableXml(b, secXml) : paraXml(b, { secXml });
      else body += b.type === 'tbl' ? tableXml(b) : paraXml(b);
    });
    return `${HEAD}<hs:sec ${NS}>${body}</hs:sec>`;
  });
  // ── 머리 (header.xml)
  const fontList = fonts.list.map((f, i) => `<hh:font id="${i}" face="${esc(f)}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font>`).join('');
  const langs = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
  const side = (n, b) => (b && b !== 1 && b.style !== 'none' ? `<hh:${n} type="${LINE_OUT[b.style] ?? 'SOLID'}" width="${mm(b.w)}" color="${hexc(b.color)}"/>` : b === 1 ? `<hh:${n} type="SOLID" width="0.12 mm" color="#000000"/>` : `<hh:${n} type="NONE" width="0.1 mm" color="#000000"/>`);
  const bfXml = bfs.list.map((b, i) => `<hh:borderFill id="${i + 1}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>${side('leftBorder', b.l)}${side('rightBorder', b.r)}${side('topBorder', b.t)}${side('bottomBorder', b.b)}<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>${b.fill ? `<hc:fillBrush><hc:winBrush faceColor="${hexc(b.fill)}" hatchColor="#999999" alpha="0"/></hc:fillBrush>` : ''}</hh:borderFill>`).join('');
  const charXml = chars.list.map((c, i) => `<hh:charPr id="${i}" height="${Math.round(c.sz * 100)}" textColor="${hexc(c.color)}" shadeColor="${c.shd ? hexc(c.shd) : 'none'}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1"><hh:fontRef hangul="${c.ea}" latin="${c.lat}" hanja="${c.ea}" japanese="${c.ea}" other="${c.ea}" symbol="${c.ea}" user="${c.ea}"/><hh:ratio hangul="${c.scale}" latin="${c.scale}" hanja="${c.scale}" japanese="${c.scale}" other="${c.scale}" symbol="${c.scale}" user="${c.scale}"/><hh:spacing hangul="${c.spc}" latin="${c.spc}" hanja="${c.spc}" japanese="${c.spc}" other="${c.spc}" symbol="${c.spc}" user="${c.spc}"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="${c.pos}" latin="${c.pos}" hanja="${c.pos}" japanese="${c.pos}" other="${c.pos}" symbol="${c.pos}" user="${c.pos}"/>${c.b ? '<hh:bold/>' : ''}${c.i ? '<hh:italic/>' : ''}<hh:underline type="${c.u ? 'BOTTOM' : 'NONE'}" shape="${c.u ?? 'SOLID'}" color="${hexc(c.color)}"/><hh:strikeout shape="${c.s ? 'SOLID' : 'NONE'}" color="${hexc(c.color)}"/><hh:outline type="${c.outline ? 'SOLID' : 'NONE'}"/><hh:shadow type="${c.shadow ? 'DROP' : 'NONE'}" color="#B2B2B2" offsetX="10" offsetY="10"/>${c.emboss ? '<hh:emboss/>' : ''}${c.imprint ? '<hh:engrave/>' : ''}${c.va === 'super' ? '<hh:supscript/>' : c.va === 'sub' ? '<hh:subscript/>' : ''}</hh:charPr>`).join('');
  const marginXml = (p) => `<hh:margin><hc:intent value="${p.intent}" unit="HWPUNIT"/><hc:left value="${p.left}" unit="HWPUNIT"/><hc:right value="${p.right}" unit="HWPUNIT"/><hc:prev value="${p.prev}" unit="HWPUNIT"/><hc:next value="${p.next}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="${p.ls.type}" value="${p.ls.v}" unit="HWPUNIT"/>`;
  const paraXmlH = paras.list.map((p, i) => `<hh:paraPr id="${i}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0"><hh:align horizontal="${p.align}" vertical="BASELINE"/><hh:heading type="${p.head?.type ?? 'NONE'}" idRef="${p.head?.id ?? 0}" level="${p.head?.level ?? 0}"/><hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="${p.widow ? 1 : 0}" keepWithNext="${p.keepNext ? 1 : 0}" keepLines="${p.keepLines ? 1 : 0}" pageBreakBefore="${p.pbb ? 1 : 0}" lineWrap="BREAK"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/><hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">${marginXml(p)}</hp:case><hp:default>${marginXml(p)}</hp:default></hp:switch><hh:border borderFillIDRef="${p.bf}" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`).join('');
  const heads = (list) => list.map((h, i) => `<hh:paraHead start="${h.start}" level="${i + 1}" align="LEFT" useInstWidth="1" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="${h.fmt}" charPrIDRef="4294967295" checkable="0">${esc(h.text)}</hh:paraHead>`).join('');
  const outlineHeads = Array.from({ length: 7 }, (_, i) => ({ start: 1, fmt: 'DIGIT', text: '' }));
  const numXml = `<hh:numberings itemCnt="${nums.list.length + 1}"><hh:numbering id="1" start="0">${heads(outlineHeads)}</hh:numbering>${nums.list.map((h, i) => `<hh:numbering id="${i + 2}" start="0">${heads(h)}</hh:numbering>`).join('')}</hh:numberings>`;
  const bulXml = `<hh:bullets itemCnt="${bullets.list.length}">${bullets.list.map((c, i) => `<hh:bullet id="${i + 1}" char="${esc(c)}" useImage="0"><hh:paraHead level="0" align="LEFT" useInstWidth="0" autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295" checkable="0"/></hh:bullet>`).join('')}</hh:bullets>`;
  // 스타일 목록 (단락 스타일만; 각 스타일의 서식은 그 스타일 기본 문단에서)
  const styleXml = styleOrder.map((id, i) => {
    const st = doc.styles[id] ?? { name: id };
    const fake = { type: 'p', id: `s${i}`, style: id, pPr: {}, runs: [{ text: 'x' }] };
    const pid = paraId(effPara(doc, fake), null, 10);
    const cid = charId(effRun(doc, fake, fake.runs[0]));
    const name = id === 'Normal' ? '바탕글' : /^Heading(\d)$/.test(id) ? `개요 ${/^Heading(\d)$/.exec(id)[1]}` : st.name ?? id;
    const eng = id === 'Normal' ? 'Normal' : /^Heading(\d)$/.test(id) ? `Outline ${/^Heading(\d)$/.exec(id)[1]}` : id;
    return `<hh:style id="${i}" type="PARA" name="${esc(name)}" engName="${esc(eng)}" paraPrIDRef="${pid}" charPrIDRef="${cid}" nextStyleIDRef="0" langID="1042" lockForm="0"/>`;
  }).join('');
  const header = `${HEAD}<hh:head ${NS} version="1.4" secCnt="${secFiles.length}"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList><hh:fontfaces itemCnt="7">${langs.map((l) => `<hh:fontface lang="${l}" fontCnt="${fonts.list.length}">${fontList}</hh:fontface>`).join('')}</hh:fontfaces><hh:borderFills itemCnt="${bfs.list.length}">${bfXml}</hh:borderFills><hh:charProperties itemCnt="${chars.list.length}">${charXml}</hh:charProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>${numXml}${bulXml}<hh:paraProperties itemCnt="${paras.list.length}">${paraXmlH}</hh:paraProperties><hh:styles itemCnt="${styleOrder.length}">${styleXml}</hh:styles></hh:refList><hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument><hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption><hh:trackchageConfig flags="56"/></hh:head>`;
  const p = doc.props ?? {};
  const manifest = `${HEAD}<opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title>${esc(p.title ?? '')}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">${esc(p.author ?? 'WIWORD')}</opf:meta><opf:meta name="subject" content="text">${esc(p.subject ?? '')}</opf:meta><opf:meta name="description" content="text">${esc(p.comments ?? '')}</opf:meta><opf:meta name="CreatedDate" content="text">${esc(p.created ?? '')}</opf:meta><opf:meta name="ModifiedDate" content="text">${new Date().toISOString()}</opf:meta><opf:meta name="keyword" content="text">${esc(p.keywords ?? '')}</opf:meta></opf:metadata><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>${secFiles.map((_, i) => `<opf:item id="section${i}" href="Contents/section${i}.xml" media-type="application/xml"/>`).join('')}${bins.map((b) => `<opf:item id="${b.id}" href="${b.name}" media-type="${b.mime}" isEmbeded="1"/>`).join('')}<opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/>${secFiles.map((_, i) => `<opf:itemref idref="section${i}" linear="yes"/>`).join('')}</opf:spine></opf:package>`;
  const out = {
    mimetype: 'application/hwp+zip',
    'version.xml': `${HEAD}<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="WIWORD" appVersion="1.0"/>`,
    'Contents/header.xml': header,
  };
  secFiles.forEach((s, i) => { out[`Contents/section${i}.xml`] = s; });
  for (const b of bins) out[b.name] = b.bytes;
  out['Contents/content.hpf'] = manifest;
  out['settings.xml'] = `${HEAD}<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`;
  out['META-INF/container.xml'] = `${HEAD}<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/></ocf:rootfiles></ocf:container>`;
  out['META-INF/manifest.xml'] = `${HEAD}<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`;
  out['Preview/PrvText.txt'] = docText(doc).slice(0, 1024);
  return out;
}

/** LaTeX → 한글 수식 스크립트 (자주 쓰는 것만) */
export function latexToHwpEq(s) {
  let t = s;
  for (let k = 0; k < 4; k++) t = t.replace(/\\frac\{([^{}]*)\}\{([^{}]*)\}/g, '{$1} over {$2}');
  t = t.replace(/\\sqrt\[([^\]]*)\]\{([^{}]*)\}/g, 'root {$1} of {$2}').replace(/\\sqrt\{([^{}]*)\}/g, 'sqrt {$1}');
  t = t.replace(/\\left|\\right/g, '').replace(/\\cdot/g, 'cdot').replace(/\\times/g, 'times').replace(/\\div/g, 'div').replace(/\\le\b|\\leq/g, 'leq').replace(/\\ge\b|\\geq/g, 'geq').replace(/\\pm/g, '+-').replace(/\\infty/g, 'inf');
  t = t.replace(/\\(sum|int|prod|oint|lim|sin|cos|tan|log|ln|exp|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|phi|omega|Delta|Sigma|Omega|Pi)\b/g, '$1');
  t = t.replace(/\\text\{([^{}]*)\}/g, '"$1"').replace(/\\,|\\;|\\!/g, ' ');
  return t;
}

