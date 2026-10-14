// Word 문서 (.docx · .dotx · .docm) 읽기 · 쓰기 — DOM 없음 (MCP 서버 · 테스트에서도 씀)
// 읽기: 스타일(이름으로 한국어판 ID 대응) · 테마 글꼴/색 · 번호 매기기 · 구역(용지 · 여백 · 단 · 쪽 번호 · 머리글/바닥글 첫 쪽/짝수) · 표(병합 · 테두리 · 음영 · 머리글 행)
//       그림(인라인 · 떠 있는 그림 · 자르기 · EMF/WMF → SVG) · 글상자 · 하이퍼링크 · 책갈피 · 필드(PAGE · TOC · MERGEFIELD …) · 각주/미주 · 메모 · 변경 내용 · 수식(OMML) · 워터마크 · 기호
// 쓰기: 같은 내용을 Word 형식으로. WIWORD 만의 정보는 word/wiword.json 에 함께 넣고, 다시 열 때 본문이 그대로면 그것을 씀
import { unzip } from './zip.js';
import { parseXml, child, kids, esc } from './xml.js';
import { newDoc, newSection, para, uid, clone, builtinStyles, normalizeRuns, docText, walkParas, runLen, PAPERS, cm, styleChain, computeNumbering } from './model.js';
import { ommlToLatex, parseLatex, toOMML, toMathML } from './math.js';
import { emfToSvg } from './emf.js';
import { wmfToSvg } from './wmf.js';
import { symbolText } from './symfont.js';
import { tableLook } from './render.js';

const dec = new TextDecoder();
const tw = (v) => (v == null || v === '' ? null : Number(v) / 15); // twip → px
const emu = (v) => (v == null ? null : Number(v) / 9525);
const onOff = (el) => (el ? !['0', 'false', 'off', 'none'].includes(String(el.attrs['w:val'] ?? 'true')) : undefined);
const av = (el, a = 'w:val') => el?.attrs[a];
const kidsDeep = (el, name, out = []) => { for (const c of el?.children ?? []) { if (c.name === name) out.push(c); else kidsDeep(c, name, out); } return out; };
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf' };
function b64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
function unb64(s) {
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(s, 'base64'));
  const bin = atob(s); const out = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out;
}
export function mediaDataUrl(bytes, name) {
  const ext = (/\.([a-z0-9]+)$/i.exec(name)?.[1] ?? 'png').toLowerCase();
  if (ext === 'emf') { const svg = emfToSvg(bytes); if (svg) return `data:image/svg+xml;base64,${b64(new TextEncoder().encode(svg))}`; }
  if (ext === 'wmf') { const svg = wmfToSvg(bytes); if (svg) return `data:image/svg+xml;base64,${b64(new TextEncoder().encode(svg))}`; }
  return `data:${MIME[ext] ?? 'application/octet-stream'};base64,${b64(bytes)}`;
}
export function dataUrlBytes(url) {
  const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url ?? '');
  if (!m) return null;
  const mime = m[1] ?? 'application/octet-stream';
  const bytes = m[2] ? unb64(m[3]) : new TextEncoder().encode(decodeURIComponent(m[3]));
  const ext = Object.entries(MIME).find(([, v]) => v === mime)?.[0] ?? (mime.includes('svg') ? 'svg' : 'png');
  return { bytes, mime, ext: ext === 'jpeg' ? 'jpg' : ext };
}

// 이름 → WIWORD 스타일 ID (한국어판 Word 는 ID 가 '1', 'a' 같은 글자라 이름으로 맞춤)
const NAME_ID = { normal: 'Normal', title: 'Title', subtitle: 'Subtitle', 'list paragraph': 'ListParagraph', caption: 'Caption', header: 'Header', footer: 'Footer', 'footnote text': 'FootnoteText', 'endnote text': 'EndnoteText', 'toc heading': 'TOCHeading', bibliography: 'Bibliography', quote: 'Quote', 'intense quote': 'IntenseQuote', 'no spacing': 'NoSpacing', 'table grid': 'TableGrid', 'normal table': 'TableNormal', 'default paragraph font': 'DefaultParagraphFont', hyperlink: 'Hyperlink', followedhyperlink: 'FollowedHyperlink', strong: 'Strong', emphasis: 'Emphasis', 'subtle emphasis': 'SubtleEmphasis', 'intense emphasis': 'IntenseEmphasis', 'subtle reference': 'SubtleReference', 'intense reference': 'IntenseReference', 'book title': 'BookTitle', 'footnote reference': 'FootnoteReference', 'endnote reference': 'EndnoteReference', 'placeholder text': 'PlaceholderText' };
const ID_NAME = { ...Object.fromEntries(Object.entries(NAME_ID).map(([k, v]) => [v, k])), Normal: 'Normal', Title: 'Title', Subtitle: 'Subtitle', Hyperlink: 'Hyperlink', Strong: 'Strong', Emphasis: 'Emphasis', Quote: 'Quote' };
for (let i = 1; i <= 9; i++) { NAME_ID[`heading ${i}`] = `Heading${i}`; NAME_ID[`toc ${i}`] = `TOC${i}`; ID_NAME[`Heading${i}`] = `heading ${i}`; ID_NAME[`TOC${i}`] = `toc ${i}`; }
function styleIdFromName(name, fileId) {
  const n = (name ?? '').toLowerCase().trim();
  if (NAME_ID[n]) return NAME_ID[n];
  const t = /^(grid|list|plain) table (\d)( light)?(?:\s*-?\s*accent (\d))?$/i.exec(name ?? '');
  if (t) return `${t[1][0].toUpperCase()}${t[1].slice(1).toLowerCase()}Table${t[2]}${t[3] ? 'Light' : ''}${t[4] ? `-Accent${t[4]}` : ''}`;
  return fileId;
}

// ───────────── 테마 ─────────────
function readTheme(xml) {
  const out = { colors: {}, fonts: null };
  if (!xml) return out;
  const t = parseXml(xml);
  const cs = kidsDeep(t, 'clrScheme')[0];
  for (const c of cs?.children ?? []) { const v = c.children[0]; out.colors[c.name] = v?.name === 'sysClr' ? `#${v.attrs.lastClr ?? '000000'}` : `#${v?.attrs.val ?? '000000'}`; }
  const fs = kidsDeep(t, 'fontScheme')[0];
  const pick = (el) => { if (!el) return null; const latin = child(el, 'latin')?.attrs.typeface; const ea = child(el, 'ea')?.attrs.typeface || kids(el, 'font').find((f) => f.attrs.script === 'Hang')?.attrs.typeface || latin; return { latin: latin || '맑은 고딕', ea: ea || '맑은 고딕' }; };
  if (fs) out.fonts = { major: pick(child(fs, 'majorFont')), minor: pick(child(fs, 'minorFont')) };
  return out;
}
const THEME_KEY = { dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hyperlink: 'hlink', followedHyperlink: 'folHlink' };
function tintHex(hex, tint, shade) {
  let [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  if (tint != null) { const t = parseInt(tint, 16) / 255; [r, g, b] = [r, g, b].map((v) => v + (255 - v) * (1 - t)); }
  if (shade != null) { const s = parseInt(shade, 16) / 255; [r, g, b] = [r, g, b].map((v) => v * s); }
  return `#${[r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

// ───────────── 글자 · 문단 서식 ─────────────
const BORDER_STYLE = { single: 'single', thick: 'thick', double: 'double', dotted: 'dotted', dashed: 'dashed', dotDash: 'dotDash', dotDotDash: 'dotDotDash', triple: 'triple', wave: 'wave', thinThickSmallGap: 'thinThickSmallGap', thickThinSmallGap: 'thickThinSmallGap', none: 'none', nil: 'none', dashSmallGap: 'dashed', threeDEmboss: 'single', threeDEngrave: 'single', outset: 'single', inset: 'single' };
function border(el, ctx) {
  if (!el) return undefined;
  const v = av(el);
  if (!v || v === 'nil' || v === 'none') return { style: 'none' };
  return { style: BORDER_STYLE[v] ?? 'single', w: Math.max(0.25, Number(av(el, 'w:sz') ?? 4) / 8), color: color(el, ctx) ?? '#000000', space: Number(av(el, 'w:space') ?? 0) };
}
function color(el, ctx, attr = 'w:color') {
  const th = el?.attrs['w:themeColor'];
  if (th && ctx?.theme?.colors[THEME_KEY[th]]) return tintHex(ctx.theme.colors[THEME_KEY[th]], el.attrs['w:themeTint'], el.attrs['w:themeShade']);
  const v = el?.attrs[attr];
  if (!v || v === 'auto') return null;
  return /^[0-9A-Fa-f]{6}$/.test(v) ? `#${v.toUpperCase()}` : null;
}
const HL = { yellow: 'yellow', green: 'brightGreen', cyan: 'cyan', magenta: 'magenta', blue: 'blue', red: 'red', darkBlue: 'darkBlue', darkCyan: 'darkCyan', darkGreen: 'darkGreen', darkMagenta: 'darkMagenta', darkRed: 'darkRed', darkYellow: 'darkYellow', darkGray: 'darkGray', lightGray: 'lightGray', black: 'black', white: 'white' };
const HL_OUT = Object.fromEntries(Object.entries(HL).map(([k, v]) => [v, k]));
function readRpr(el, ctx) {
  const o = {};
  if (!el) return o;
  for (const c of el.children) {
    switch (c.name) {
      case 'rFonts': {
        const A = c.attrs;
        const lat = A['w:asciiTheme'] ?? A['w:hAnsiTheme'];
        if (lat) o.font = lat.startsWith('major') ? '+mj' : '+mn'; else if (A['w:ascii'] || A['w:hAnsi']) o.font = A['w:ascii'] ?? A['w:hAnsi'];
        if (A['w:eastAsiaTheme']) o.fontEa = A['w:eastAsiaTheme'].startsWith('major') ? '+mj' : '+mn'; else if (A['w:eastAsia']) o.fontEa = A['w:eastAsia'];
        break;
      }
      case 'b': o.b = onOff(c); break;
      case 'i': o.i = onOff(c); break;
      case 'u': { const v = av(c) ?? 'single'; o.u = v === 'none' ? false : v === 'single' ? true : v === 'words' ? true : v; const uc = color(c, ctx); if (uc) o.uColor = uc; break; }
      case 'strike': o.s = onOff(c); break;
      case 'dstrike': o.dstrike = onOff(c); break;
      case 'color': { const v = color(c, ctx, 'w:val'); if (v) o.color = v; break; }
      case 'sz': o.sz = Number(av(c)) / 2; break;
      case 'highlight': { const v = av(c); if (v && v !== 'none') o.hl = HL[v] ?? v; break; }
      case 'shd': { const f = color(c, ctx, 'w:fill'); if (f) o.shd = f; break; }
      case 'vertAlign': { const v = av(c); if (v === 'superscript') o.va = 'super'; else if (v === 'subscript') o.va = 'sub'; break; }
      case 'caps': o.caps = onOff(c); break;
      case 'smallCaps': o.smallCaps = onOff(c); break;
      case 'spacing': o.spc = Number(av(c)) / 20; break;
      case 'w': o.scale = Number(av(c)); break;
      case 'position': o.pos = Number(av(c)) / 2; break;
      case 'vanish': o.hidden = onOff(c); break;
      case 'rStyle': o.cstyle = ctx.styleMap?.[av(c)] ?? av(c); break;
      case 'lang': if (av(c, 'w:eastAsia') || av(c)) o.lang = av(c, 'w:eastAsia') ?? av(c); break;
      case 'outline': o.outline = onOff(c); break;
      case 'shadow': o.shadow = onOff(c); break;
      case 'emboss': o.emboss = onOff(c); break;
      case 'imprint': o.imprint = onOff(c); break;
      case 'em': { const v = av(c); if (v && v !== 'none') o.emph = v === 'comma' ? 'comma' : v === 'circle' ? 'circle' : 'dot'; break; }
      case 'kern': o.kern = Number(av(c)) / 2; break;
      case 'rtl': o.rtl = onOff(c); break;
      default: break;
    }
  }
  for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k];
  return o;
}
const JC = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'distribute', thaiDistribute: 'distribute', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify' };
function readPpr(el, ctx) {
  const o = {};
  if (!el) return o;
  for (const c of el.children) {
    switch (c.name) {
      case 'jc': o.align = JC[av(c)] ?? 'left'; break;
      case 'ind': {
        const A = c.attrs; const ind = {};
        const l = A['w:left'] ?? A['w:start']; const r = A['w:right'] ?? A['w:end'];
        if (l != null) ind.l = tw(l);
        if (r != null) ind.r = tw(r);
        if (A['w:hanging'] != null) ind.hanging = tw(A['w:hanging']);
        else if (A['w:firstLine'] != null) ind.first = tw(A['w:firstLine']);
        if (A['w:leftChars'] != null && l == null) ind.l = (Number(A['w:leftChars']) / 100) * 13.33;
        if (A['w:firstLineChars'] != null && A['w:firstLine'] == null) ind.first = (Number(A['w:firstLineChars']) / 100) * 13.33;
        o.ind = ind;
        break;
      }
      case 'spacing': {
        const A = c.attrs; const s = {};
        if (A['w:before'] != null) s.before = tw(A['w:before']);
        if (A['w:after'] != null) s.after = tw(A['w:after']);
        if (A['w:beforeAutospacing'] === '1') s.before = 18.67;
        if (A['w:afterAutospacing'] === '1') s.after = 18.67;
        if (A['w:line'] != null) {
          const rule = A['w:lineRule'] ?? 'auto';
          if (rule === 'auto') { s.line = Number(A['w:line']) / 240; s.rule = 'auto'; } else { s.line = tw(A['w:line']); s.rule = rule === 'exact' ? 'exact' : 'atLeast'; }
        }
        o.spc = s;
        break;
      }
      case 'keepNext': o.keepNext = onOff(c); break;
      case 'keepLines': o.keepLines = onOff(c); break;
      case 'pageBreakBefore': o.pageBreakBefore = onOff(c); break;
      case 'widowControl': o.widow = onOff(c); break;
      case 'numPr': { const id = av(child(c, 'numId')); if (id != null) o.num = { id: id === '0' ? null : String(id), lvl: Number(av(child(c, 'ilvl')) ?? 0) }; break; }
      case 'tabs': o.tabs = kids(c, 'tab').map((t) => ({ pos: tw(av(t, 'w:pos')), align: { start: 'left', end: 'right', num: 'decimal' }[av(t)] ?? av(t) ?? 'left', ...(av(t, 'w:leader') && av(t, 'w:leader') !== 'none' ? { leader: av(t, 'w:leader') } : {}) })); break;
      case 'pBdr': { const b = {}; for (const k of ['top', 'bottom', 'left', 'right', 'between']) { const x = child(c, k) ?? (k === 'left' ? child(c, 'start') : k === 'right' ? child(c, 'end') : null); if (x) b[k] = border(x, ctx); } o.border = b; break; }
      case 'shd': { const f = color(c, ctx, 'w:fill'); if (f) o.shd = f; break; }
      case 'outlineLvl': o.outline = Number(av(c)); break;
      case 'contextualSpacing': o.contextual = onOff(c); break;
      case 'suppressAutoHyphens': o.noHyphen = onOff(c); break;
      case 'bidi': o.bidi = onOff(c); break;
      case 'snapToGrid': o.snapGrid = onOff(c); break;
      case 'framePr': if (av(c, 'w:dropCap') && av(c, 'w:dropCap') !== 'none') o.dropCap = { lines: Number(av(c, 'w:lines') ?? 3), kind: av(c, 'w:dropCap') === 'margin' ? 'margin' : 'drop' }; break;
      default: break;
    }
  }
  for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k];
  return o;
}

// ───────────── 구역 ─────────────
function readSect(el, ctx) {
  const sec = newSection();
  sec.headers = {}; sec.footers = {};
  let kind = 'nextPage';
  if (!el) return { sec, kind };
  const sz = child(el, 'pgSz');
  if (sz) { sec.page = { w: tw(av(sz, 'w:w')) ?? sec.page.w, h: tw(av(sz, 'w:h')) ?? sec.page.h, orient: av(sz, 'w:orient') === 'landscape' ? 'landscape' : 'portrait' }; }
  const mg = child(el, 'pgMar');
  if (mg) { const A = mg.attrs; sec.margin = { t: Math.abs(tw(A['w:top']) ?? 96), b: Math.abs(tw(A['w:bottom']) ?? 96), l: tw(A['w:left'] ?? A['w:start']) ?? 120, r: tw(A['w:right'] ?? A['w:end']) ?? 120, header: tw(A['w:header']) ?? 48, footer: tw(A['w:footer']) ?? 48, gutter: tw(A['w:gutter']) ?? 0 }; }
  const cols = child(el, 'cols');
  if (cols) {
    const n = Number(av(cols, 'w:num') ?? 1);
    const cl = kids(cols, 'col');
    sec.cols = { n: Math.max(1, cl.length || n), space: tw(av(cols, 'w:space')) ?? cm(1.25), sep: av(cols, 'w:sep') === '1' || av(cols, 'w:sep') === 'true' };
    if (cl.length && av(cols, 'w:equalWidth') === '0') sec.cols.widths = cl.map((c) => tw(av(c, 'w:w')));
  }
  const pn = child(el, 'pgNumType');
  if (pn && (av(pn, 'w:fmt') || av(pn, 'w:start'))) sec.pgNum = { fmt: av(pn, 'w:fmt') ?? 'decimal', ...(av(pn, 'w:start') != null ? { start: Number(av(pn, 'w:start')) } : {}) };
  if (child(el, 'titlePg') && onOff(child(el, 'titlePg'))) sec.titlePg = true;
  const va = av(child(el, 'vAlign'));
  if (va) sec.vAlign = va === 'both' ? 'justify' : va;
  const t = av(child(el, 'type'));
  if (t) kind = t;
  const pb = child(el, 'pgBorders');
  if (pb) { const b = border(child(pb, 'top') ?? child(pb, 'left'), ctx); if (b && b.style !== 'none') sec.pageBorder = { ...b, space: b.space || 24 }; }
  const ln = child(el, 'lnNumType');
  if (ln) sec.lineNumbers = { countBy: Number(av(ln, 'w:countBy') ?? 1), restart: av(ln, 'w:restart') ?? 'newPage' };
  for (const ref of kids(el, 'headerReference')) { const bl = ctx.part(ref.attrs['r:id'], 'hdr'); if (bl) sec.headers[ref.attrs['w:type'] ?? 'default'] = bl; }
  for (const ref of kids(el, 'footerReference')) { const bl = ctx.part(ref.attrs['r:id'], 'ftr'); if (bl) sec.footers[ref.attrs['w:type'] ?? 'default'] = bl; }
  return { sec, kind, inherit: { h: kids(el, 'headerReference').map((r) => r.attrs['w:type'] ?? 'default'), f: kids(el, 'footerReference').map((r) => r.attrs['w:type'] ?? 'default') } };
}

// ───────────── 본문 ─────────────
const SIMPLE_FIELDS = /^(PAGE|NUMPAGES|SECTIONPAGES|DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE|MERGEFIELD|REF|PAGEREF|NOTEREF|SEQ|AUTHOR|TITLE|SUBJECT|FILENAME|DOCPROPERTY|STYLEREF|NUMWORDS|NUMCHARS|USERNAME|SYMBOL|FORMTEXT|QUOTE|LASTSAVEDBY|KEYWORDS|COMMENTS|INFO|SECTION|LISTNUM|AUTONUM|AUTONUMLGL|BIBLIOGRAPHY_CITATION|CITATION)$/i;
function readBlocks(container, ctx) {
  const out = [];
  for (const c of container?.children ?? []) {
    if (c.name === 'p') out.push(...readPara(c, ctx));
    else if (c.name === 'tbl') { const t = readTable(c, ctx); if (t) out.push(t); }
    else if (c.name === 'sdt') out.push(...readBlocks(child(c, 'sdtContent'), ctx));
    else if (c.name === 'customXml' || c.name === 'smartTag') out.push(...readBlocks(c, ctx));
    else if (c.name === 'AlternateContent') out.push(...readBlocks(child(c, 'Choice') ?? child(c, 'Fallback'), ctx));
    else if (c.name === 'ins' || c.name === 'moveTo') out.push(...readBlocks(c, ctx));
    else if (c.name === 'bookmarkStart' || c.name === 'bookmarkEnd') { /* 블록 사이 책갈피: 무시 */ }
  }
  return out;
}

function readPara(el, ctx) {
  const pPrEl = child(el, 'pPr');
  const p = para('');
  const styleFile = av(child(pPrEl, 'pStyle'));
  p.style = styleFile ? (ctx.styleMap[styleFile] ?? styleFile) : 'Normal';
  if (!ctx.doc.styles[p.style]) p.style = 'Normal';
  p.pPr = readPpr(pPrEl, ctx);
  const markRpr = readRpr(child(pPrEl, 'rPr'), ctx);
  delete markRpr.cstyle;
  if (Object.keys(markRpr).length) p.mark = markRpr;
  if (child(child(pPrEl, 'rPr'), 'ins')) p.paraIns = true;
  const runs = [];
  const extra = []; // 문단을 쪼개야 하는 경우 (여러 문단짜리 필드 결과는 그대로 이어 붙임)
  const st = { fld: ctx.fldStack, runs, link: null, rev: null };
  readInline(el, ctx, st, {});
  p.runs = normalizeRuns(runs);
  // 구역
  const sp = child(pPrEl, 'sectPr');
  if (sp) { const s = readSect(sp, ctx); ctx.sections.push(s); p.secBreak = { kind: 'nextPage', sec: s.sec, _kind: s.kind }; }
  return [p, ...extra];
}

/** 문단 안 내용 (런 · 하이퍼링크 · 변경 내용 · 필드 · 책갈피 · 메모) */
function readInline(el, ctx, st, rev) {
  for (const c of el.children) {
    switch (c.name) {
      case 'pPr': break;
      case 'r': readRun(c, ctx, st, rev); break;
      case 'hyperlink': {
        const rid = c.attrs['r:id'];
        const url = rid ? ctx.rels[rid]?.target : null;
        const anchor = c.attrs['w:anchor'];
        const prev = st.link;
        st.link = url ? { url } : anchor ? { anchor } : null;
        readInline(c, ctx, st, rev);
        st.link = prev;
        break;
      }
      case 'ins': case 'moveTo': readInline(c, ctx, st, { ...rev, ins: { author: c.attrs['w:author'] ?? '', date: c.attrs['w:date'] ?? '' } }); break;
      case 'del': case 'moveFrom': readInline(c, ctx, st, { ...rev, del: { author: c.attrs['w:author'] ?? '', date: c.attrs['w:date'] ?? '' } }); break;
      case 'smartTag': case 'customXml': case 'dir': case 'bdo': readInline(c, ctx, st, rev); break;
      case 'sdt': readInline(child(c, 'sdtContent') ?? c, ctx, st, rev); break;
      case 'fldSimple': {
        const code = (c.attrs['w:instr'] ?? '').trim();
        const inner = { ...st, runs: [] };
        readInline(c, ctx, inner, rev);
        const result = inner.runs.map((r) => r.text ?? '').join('');
        const head = code.split(/\s+/)[0];
        if (/^HYPERLINK$/i.test(head)) { const m = /HYPERLINK\s+"([^"]*)"(?:\s+\\l\s+"([^"]*)")?/i.exec(code); for (const r of inner.runs) st.runs.push({ ...r, link: m?.[2] ? { anchor: m[2] } : { url: m?.[1] ?? '' } }); } else st.runs.push({ obj: 'field', code, result, ...firstProps(inner.runs) });
        break;
      }
      case 'bookmarkStart': { const name = c.attrs['w:name']; if (name && name !== '_GoBack') { ctx.bm[c.attrs['w:id']] = name; st.runs.push({ obj: 'mark', kind: 'bm', name, edge: 'start' }); } break; }
      case 'bookmarkEnd': { const name = ctx.bm[c.attrs['w:id']]; if (name) st.runs.push({ obj: 'mark', kind: 'bm', name, edge: 'end' }); break; }
      case 'commentRangeStart': st.runs.push({ obj: 'mark', kind: 'cmt', id: `c${c.attrs['w:id']}`, edge: 'start' }); ctx.cmtUsed.add(`c${c.attrs['w:id']}`); break;
      case 'commentRangeEnd': st.runs.push({ obj: 'mark', kind: 'cmt', id: `c${c.attrs['w:id']}`, edge: 'end' }); break;
      case 'oMath': case 'oMathPara': { const latex = ommlToLatex(c); if (latex) st.runs.push({ obj: 'eq', latex, html: safeMathML(latex), ...rev }); break; }
      case 'AlternateContent': readInline(child(c, 'Choice') ?? child(c, 'Fallback') ?? c, ctx, st, rev); break;
      default: break;
    }
  }
}
function safeMathML(latex) { try { return toMathML(parseLatex(latex), { display: false }); } catch { return null; } }
const firstProps = (runs) => { const r = runs.find((x) => !x.obj) ?? {}; const { text: _t, ...rest } = r; return rest; };

function readRun(el, ctx, st, rev) {
  const rp = { ...readRpr(child(el, 'rPr'), ctx), ...rev };
  if (st.link) { rp.link = st.link; }
  const f = st.fld;
  const push = (r) => {
    // 필드 코드 · 결과 처리
    const top = f[f.length - 1];
    if (top && top.phase === 'code') return;
    if (top && top.phase === 'result' && top.collect) { top.result.push(r); return; }
    if (top && top.phase === 'result' && top.link) r = { ...r, link: top.link };
    st.runs.push(r);
  };
  for (const c of el.children) {
    switch (c.name) {
      case 't': if (c.text) push({ ...rp, text: c.text }); break;
      case 'delText': if (c.text) push({ ...rp, text: c.text }); break;
      case 'instrText': case 'delInstrText': { const top = f[f.length - 1]; if (top && top.phase === 'code') top.code += c.text; break; }
      case 'tab': push({ ...rp, obj: 'tab' }); break;
      case 'ptab': push({ ...rp, obj: 'tab' }); break;
      case 'br': { const t = av(c, 'w:type'); push({ obj: 'br', ...(t === 'page' ? { kind: 'page' } : t === 'column' ? { kind: 'column' } : {}) }); break; }
      case 'cr': push({ obj: 'br' }); break;
      case 'noBreakHyphen': push({ ...rp, text: '‑' }); break;
      case 'softHyphen': push({ ...rp, text: '­' }); break;
      case 'sym': { const font = av(c, 'w:font'); const code = parseInt(av(c, 'w:char') ?? '0', 16); const ch = symbolText(font, String.fromCharCode(code >= 0xf000 ? code : code + 0xf000)) || String.fromCharCode(code); push({ ...rp, obj: 'sym', char: ch, font }); break; }
      case 'footnoteReference': case 'endnoteReference': { const kind = c.name === 'footnoteReference' ? 'foot' : 'end'; const id = `${kind[0]}${av(c, 'w:id')}`; ctx.noteRefs.push({ kind, fileId: av(c, 'w:id'), id }); push({ obj: 'note', id, kind, ...(c.attrs['w:customMarkFollows'] === '1' ? { custom: '*' } : {}) }); break; }
      case 'footnoteRef': case 'endnoteRef': break;
      case 'commentReference': break;
      case 'separator': case 'continuationSeparator': break;
      case 'fldChar': {
        const t = av(c, 'w:fldCharType');
        if (t === 'begin') f.push({ phase: 'code', code: '', result: [], rp });
        else if (t === 'separate') {
          const top = f[f.length - 1];
          if (!top) break;
          // 바깥 필드가 결과를 모으는 중이면 안쪽 필드는 결과 글만 (중첩)
          const outer = f[f.length - 2];
          const code = top.code.trim();
          const head = code.split(/\s+/)[0]?.toUpperCase() ?? '';
          top.phase = 'result';
          if (outer && (outer.phase === 'code' || outer.collect)) { top.collect = outer.collect; top.result = outer.collect ? outer.result : []; top.nested = true; break; }
          if (head === 'HYPERLINK') { const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]*)"(?:\s+\\l\s+"([^"]*)")?/i.exec(code); const isAnchor = /\\l/.test(code) && !m?.[2]; top.link = m?.[2] ? { anchor: m[2] } : isAnchor ? { anchor: m?.[1] ?? '' } : { url: m?.[1] ?? '' }; }
          else if (SIMPLE_FIELDS.test(head)) top.collect = true;
          else top.pass = true; // TOC · INDEX 등: 결과를 그대로 본문으로
        } else if (t === 'end') {
          const top = f.pop();
          if (!top || top.nested) break;
          const code = top.code.trim();
          const head = code.split(/\s+/)[0]?.toUpperCase() ?? '';
          if (top.phase === 'code') { if (SIMPLE_FIELDS.test(head) || /^(PAGE|NUMPAGES)$/.test(head)) st.runs.push({ ...top.rp, obj: 'field', code, result: '' }); break; }
          if (top.collect) {
            const result = top.result.map((r) => (r.obj === 'tab' ? '\t' : r.obj === 'br' ? '\n' : r.text ?? '')).join('');
            st.runs.push({ ...firstProps(top.result.length ? top.result : [top.rp]), obj: 'field', code, result });
          }
        }
        break;
      }
      case 'drawing': { const o = readDrawing(c, ctx); if (o) push({ ...o, ...(rev.ins ? { ins: rev.ins } : {}), ...(st.link ? { link: st.link } : {}) }); break; }
      case 'pict': case 'object': { const o = readVml(c, ctx); if (o) push(o); break; }
      case 'AlternateContent': { const ch = child(c, 'Choice'); const fb = child(c, 'Fallback'); const tmp = { name: 'r', attrs: {}, children: [...(child(el, 'rPr') ? [child(el, 'rPr')] : []), ...((ch ?? fb)?.children ?? [])] }; readRun(tmp, ctx, st, rev); break; }
      case 'ruby': { const base = child(c, 'rubyBase'); for (const r of kids(base, 'r')) readRun(r, ctx, st, rev); break; }
      default: break;
    }
  }
}

// 그림 (DrawingML)
function readDrawing(el, ctx) {
  const box = child(el, 'inline') ?? child(el, 'anchor');
  if (!box) return null;
  const ext = child(box, 'extent');
  const w = emu(av(ext, 'cx')) ?? 100; const h = emu(av(ext, 'cy')) ?? 100;
  const docPr = child(box, 'docPr');
  const alt = docPr?.attrs.descr ?? docPr?.attrs.title ?? '';
  let float = null;
  if (box.name === 'anchor') {
    const ph = child(box, 'positionH'); const pv = child(box, 'positionV');
    const rel = (x, def) => ({ page: 'page', margin: 'margin', column: 'column', character: 'column', paragraph: 'paragraph', line: 'paragraph', leftMargin: 'page', rightMargin: 'page', insideMargin: 'margin', outsideMargin: 'margin', topMargin: 'page', bottomMargin: 'page' }[av(x, 'relativeFrom')] ?? def);
    const wrapEl = box.children.find((x) => /^wrap/.test(x.name));
    const wrap = box.attrs.behindDoc === '1' && (!wrapEl || wrapEl.name === 'wrapNone') ? 'behind' : { wrapNone: 'front', wrapSquare: 'square', wrapTight: 'tight', wrapThrough: 'tight', wrapTopAndBottom: 'topBottom' }[wrapEl?.name] ?? 'front';
    float = { wrap, relH: rel(ph, 'column'), relV: rel(pv, 'paragraph'), x: emu(child(ph, 'posOffset')?.text) ?? 0, y: emu(child(pv, 'posOffset')?.text) ?? 0 };
    const ah = child(ph, 'align')?.text; const avv = child(pv, 'align')?.text;
    if (ah) float.alignH = ah === 'inside' ? 'left' : ah === 'outside' ? 'right' : ah;
    if (avv) float.alignV = avv === 'inside' ? 'top' : avv === 'outside' ? 'bottom' : avv;
    if (wrap === 'square' || wrap === 'tight') float.side = ah === 'right' ? 'right' : float.x > 200 ? 'right' : 'left';
    float.dist = { t: emu(box.attrs.distT) ?? 0, b: emu(box.attrs.distB) ?? 0, l: emu(box.attrs.distL) ?? 12, r: emu(box.attrs.distR) ?? 12 };
  }
  const gd = kidsDeep(box, 'graphicData')[0];
  const pic = kidsDeep(gd, 'pic')[0];
  if (pic) {
    const blip = kidsDeep(pic, 'blip')[0];
    const svgBlip = kidsDeep(blip, 'svgBlip')[0];
    const rid = svgBlip?.attrs['r:embed'] ?? blip?.attrs['r:embed'] ?? blip?.attrs['r:link'];
    const media = ctx.media(rid);
    const run = { obj: 'img', id: uid('i'), w, h, alt, ...(media ? { media } : { src: '' }) };
    if (!media && blip?.attrs['r:link']) run.src = ctx.rels[blip.attrs['r:link']]?.target ?? '';
    const sr = kidsDeep(pic, 'srcRect')[0];
    if (sr) { const k = (a) => Number(sr.attrs[a] ?? 0) / 100000; const c = { l: k('l'), t: k('t'), r: k('r'), b: k('b') }; if (c.l || c.t || c.r || c.b) run.crop = c; }
    const xf = kidsDeep(pic, 'xfrm')[0];
    if (xf?.attrs.rot) run.rot = Number(xf.attrs.rot) / 60000;
    if (float) run.float = float;
    return run;
  }
  // 글상자 · 도형 (wps)
  const wsp = kidsDeep(gd, 'wsp')[0];
  if (wsp) {
    const txbx = kidsDeep(wsp, 'txbxContent')[0];
    const sp = child(wsp, 'spPr');
    const fill = kidsDeep(child(sp, 'solidFill'), 'srgbClr')[0]?.attrs.val;
    const noFill = !!child(sp, 'noFill');
    const ln = child(sp, 'ln');
    const lnc = kidsDeep(child(ln, 'solidFill'), 'srgbClr')[0]?.attrs.val;
    const prst = child(sp, 'prstGeom')?.attrs.prst ?? 'rect';
    const body = child(wsp, 'bodyPr');
    const run = { obj: 'shape', id: uid('s'), w, h, prst, blocks: txbx ? readBlocks(txbx, ctx) : [], float: float ?? { wrap: 'front', relH: 'column', relV: 'paragraph', x: 0, y: 0 } };
    if (fill && !noFill) run.fill = `#${fill}`;
    else if (!noFill && !txbx) run.fill = '#4472C4';
    if (ln && !child(ln, 'noFill')) run.line = { w: Number(ln.attrs.w ?? 9525) / 12700, color: lnc ? `#${lnc}` : '#2F528F' };
    if (body?.attrs.anchor === 'ctr') run.anchor = 'ctr';
    const ins = (a, d) => (body?.attrs[a] != null ? emu(body.attrs[a]) : d);
    run.inset = { l: ins('lIns', 9.6), r: ins('rIns', 9.6), t: ins('tIns', 4.8), b: ins('bIns', 4.8) };
    if (!run.blocks.length && !float) return null;
    return run;
  }
  // 차트 · SmartArt 등: 대체 그림이 없으면 회색 상자
  if (kidsDeep(gd, 'chart').length || kidsDeep(gd, 'relIds').length) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${Math.round(w)} ${Math.round(h)}"><rect width="100%" height="100%" fill="#F2F2F2" stroke="#BFBFBF"/><text x="50%" y="50%" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="13" fill="#7F7F7F">${kidsDeep(gd, 'chart').length ? '차트' : 'SmartArt'}</text></svg>`;
    return { obj: 'img', id: uid('i'), w, h, alt: alt || '차트', svg, ...(float ? { float } : {}) };
  }
  return null;
}
// VML (예전 그림 · 글상자 · 워터마크 · OLE 미리 보기)
function readVml(el, ctx) {
  const shape = kidsDeep(el, 'shape')[0] ?? kidsDeep(el, 'rect')[0];
  if (!shape) return null;
  const css = Object.fromEntries((shape.attrs.style ?? '').split(';').map((x) => x.split(':').map((y) => y.trim())).filter((x) => x[0]));
  const len = (v) => { if (!v) return null; const n = parseFloat(v); if (/pt$/.test(v)) return (n * 4) / 3; if (/in$/.test(v)) return n * 96; if (/cm$/.test(v)) return (n * 96) / 2.54; if (/mm$/.test(v)) return (n * 96) / 25.4; if (/px$/.test(v)) return n; return n / 15; };
  const w = len(css.width) ?? 100; const h = len(css.height) ?? 100;
  const tp = kidsDeep(shape, 'textpath')[0];
  if (tp) { ctx.watermark = { text: tp.attrs.string ?? '', color: shape.attrs.fillcolor ?? '#C0C0C0', diagonal: /rotation:\s*315/.test(shape.attrs.style ?? '') }; return null; }
  const im = kidsDeep(shape, 'imagedata')[0];
  if (im) { const media = ctx.media(im.attrs['r:id'] ?? im.attrs['r:href']); if (media) return { obj: 'img', id: uid('i'), w, h, media, alt: shape.attrs.alt ?? '' }; }
  const tx = kidsDeep(shape, 'txbxContent')[0];
  if (tx) return { obj: 'shape', id: uid('s'), w, h, prst: 'rect', blocks: readBlocks(tx, ctx), line: { w: 0.75, color: '#000000' }, float: { wrap: 'front', relH: 'column', relV: 'paragraph', x: len(css['margin-left']) ?? 0, y: len(css['margin-top']) ?? 0 } };
  return null;
}

// 표
function readTable(el, ctx) {
  // 너무 깊이 겹친 표 (손상 · 공격용 파일): 글만
  if ((ctx.depth ?? 0) > 24) return null;
  ctx = { ...ctx, depth: (ctx.depth ?? 0) + 1 };
  const tp = child(el, 'tblPr');
  const t = { type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: {}, grid: [], rows: [] };
  const sid = av(child(tp, 'tblStyle'));
  if (sid) t.style = ctx.styleMap[sid] ?? sid;
  else t.style = 'TableNormal';
  if (!ctx.doc.styles[t.style] && !/^(Grid|List|Plain)Table/.test(t.style)) t.style = 'TableGrid';
  const jc = av(child(tp, 'jc'));
  if (jc === 'center' || jc === 'right' || jc === 'end') t.tblPr.align = jc === 'end' ? 'right' : jc;
  const ind = child(tp, 'tblInd');
  if (ind) t.tblPr.ind = tw(av(ind, 'w:w')) ?? 0;
  const bd = child(tp, 'tblBorders');
  if (bd) { const b = {}; for (const k of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) { const x = child(bd, k) ?? (k === 'left' ? child(bd, 'start') : k === 'right' ? child(bd, 'end') : null); if (x) b[k] = border(x, ctx); } t.tblPr.borders = b; }
  const mar = child(tp, 'tblCellMar');
  if (mar) { const m = {}; for (const [k, n] of [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom']]) { const x = child(mar, n) ?? (n === 'left' ? child(mar, 'start') : n === 'right' ? child(mar, 'end') : null); if (x) m[k] = tw(av(x, 'w:w')); } t.tblPr.cellMar = m; }
  const look = child(tp, 'tblLook');
  if (look) {
    const A = look.attrs;
    const v = A['w:val'] ? parseInt(A['w:val'], 16) : null;
    const f = (a, bit, inv = false) => (A[a] != null ? A[a] === '1' : v != null ? (inv ? !(v & bit) : !!(v & bit)) : undefined);
    t.tblPr.look = { firstRow: f('w:firstRow', 0x20), lastRow: f('w:lastRow', 0x40), firstCol: f('w:firstColumn', 0x80), lastCol: f('w:lastColumn', 0x100), bandRows: A['w:noHBand'] != null ? A['w:noHBand'] !== '1' : v != null ? !(v & 0x200) : true, bandCols: A['w:noVBand'] != null ? A['w:noVBand'] !== '1' : v != null ? !(v & 0x400) : false };
  }
  const shd = color(child(tp, 'shd'), ctx, 'w:fill');
  if (shd) t.tblPr.shd = shd;
  const tw0 = child(tp, 'tblW');
  t.grid = kids(child(el, 'tblGrid'), 'gridCol').map((g) => tw(av(g, 'w:w')) ?? 0);
  for (const tr of kids(el, 'tr').concat(kids(el, 'sdt').flatMap((s) => kids(child(s, 'sdtContent'), 'tr')))) {
    const trPr = child(tr, 'trPr');
    const row = { cells: [] };
    const th = child(trPr, 'trHeight');
    if (th) { row.h = tw(av(th, 'w:val')); row.hRule = av(th, 'w:hRule') === 'exact' ? 'exact' : 'atLeast'; }
    if (child(trPr, 'tblHeader') && onOff(child(trPr, 'tblHeader'))) row.header = true;
    if (child(trPr, 'cantSplit') && onOff(child(trPr, 'cantSplit'))) row.cantSplit = true;
    const gb = Number(av(child(trPr, 'gridBefore')) ?? 0);
    if (gb) row.cells.push({ span: gb, blocks: [para('')], borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } } });
    const tcs = tr.children.flatMap((c) => (c.name === 'tc' ? [c] : c.name === 'sdt' ? kids(child(c, 'sdtContent'), 'tc') : c.name === 'customXml' ? kids(c, 'tc') : []));
    for (const tc of tcs) {
      const pr = child(tc, 'tcPr');
      const cell = { blocks: readBlocks(tc, ctx) };
      if (!cell.blocks.length) cell.blocks = [para('')];
      const span = Number(av(child(pr, 'gridSpan')) ?? 1);
      if (span > 1) cell.span = span;
      const vm = child(pr, 'vMerge');
      if (vm) cell.vMerge = av(vm) === 'restart' ? 'restart' : 'cont';
      if (child(pr, 'hMerge')) { const hv = av(child(pr, 'hMerge')); if (hv !== 'restart' && row.cells.length) { row.cells[row.cells.length - 1].span = (row.cells[row.cells.length - 1].span ?? 1) + 1; continue; } }
      const s = color(child(pr, 'shd'), ctx, 'w:fill');
      if (s) cell.shd = s;
      const cb = child(pr, 'tcBorders');
      if (cb) { const b = {}; for (const k of ['top', 'bottom', 'left', 'right']) { const x = child(cb, k) ?? (k === 'left' ? child(cb, 'start') : k === 'right' ? child(cb, 'end') : null); if (x) b[k] = border(x, ctx); } cell.borders = b; }
      const va = av(child(pr, 'vAlign'));
      if (va && va !== 'top') cell.vAlign = va === 'both' ? 'center' : va;
      const td = av(child(pr, 'textDirection'));
      if (td === 'tbRl' || td === 'btLr' || td === 'tbRlV' || td === 'tbLrV') cell.dir = 'vert';
      const cm2 = child(pr, 'tcMar');
      if (cm2) { const m = {}; for (const [k, n] of [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom']]) { const x = child(cm2, n) ?? (n === 'left' ? child(cm2, 'start') : n === 'right' ? child(cm2, 'end') : null); if (x) m[k] = tw(av(x, 'w:w')); } cell.mar = m; }
      const cw = child(pr, 'tcW');
      if (cw && av(cw, 'w:type') === 'dxa') cell.w = tw(av(cw, 'w:w'));
      row.cells.push(cell);
    }
    if (!row.cells.length) continue;
    t.rows.push(row);
  }
  if (!t.rows.length) return null;
  // 격자가 없거나 모자라면 칸 너비 · 표 너비로 채움
  const ncols = Math.max(...t.rows.map((r) => r.cells.reduce((s, c) => s + (c.span ?? 1), 0)));
  if (t.grid.length < ncols || t.grid.every((g) => !g)) {
    const tblW = tw0 && av(tw0, 'w:type') === 'dxa' ? tw(av(tw0, 'w:w')) : ctx.bodyW;
    const fromCells = t.rows[0].cells.length === ncols ? t.rows[0].cells.map((c) => c.w) : null;
    t.grid = Array.from({ length: ncols }, (_, i) => fromCells?.[i] || t.grid[i] || tblW / ncols);
  }
  // 표 너비가 pct 이면 본문 너비에 맞춤
  if (tw0 && av(tw0, 'w:type') === 'pct') { const pct = Number(av(tw0, 'w:w')) / 5000; const total = t.grid.reduce((a, b) => a + b, 0); if (total > 0) t.grid = t.grid.map((g) => (g / total) * ctx.bodyW * pct); }
  for (const r of t.rows) for (const c of r.cells) delete c.w;
  return t;
}

// ───────────── 읽기 ─────────────
function relsOf(files, path) {
  const dir = path.includes('/') ? path.slice(0, path.lastIndexOf('/')) : '';
  const rp = `${dir ? `${dir}/` : ''}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
  const out = {};
  const x = files[rp];
  if (!x) return out;
  for (const r of kids(parseXml(dec.decode(x)), 'Relationship')) {
    const ext = r.attrs.TargetMode === 'External';
    let target = r.attrs.Target ?? '';
    if (!ext) target = target.startsWith('/') ? target.slice(1) : normPath(`${dir ? `${dir}/` : ''}${target}`);
    out[r.attrs.Id] = { type: (r.attrs.Type ?? '').split('/').pop(), target, ext };
  }
  return out;
}
function normPath(p) { const out = []; for (const s of p.split('/')) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); } return out.join('/'); }

export function readDocx(input) {
  const files = input instanceof Uint8Array ? unzip(input) : input;
  const warnings = [];
  const xmlOf = (p) => (files[p] ? dec.decode(files[p]) : null);
  // 본문 파트 찾기
  let main = 'word/document.xml';
  const root = xmlOf('_rels/.rels');
  if (root) { const r = kids(parseXml(root), 'Relationship').find((x) => /officeDocument$/.test(x.attrs.Type)); if (r) main = normPath(r.attrs.Target.replace(/^\//, '')); }
  if (!files[main]) throw new Error('Word 문서(.docx)의 본문(word/document.xml)이 없습니다.');
  const rels = relsOf(files, main);
  const byType = (t) => Object.values(rels).find((r) => r.type === t && !r.ext)?.target;
  const doc = newDoc();
  doc.body = [];
  doc.media = {};
  doc.lists = {};
  const ctx = { doc, rels, styleMap: {}, theme: readTheme(xmlOf(byType('theme'))), bm: {}, cmtUsed: new Set(), noteRefs: [], fldStack: [], sections: [], watermark: null, bodyW: 600 };
  if (ctx.theme.fonts?.major && ctx.theme.fonts?.minor) doc.themeFonts = ctx.theme.fonts;
  // 그림
  const mediaIds = new Map();
  const mediaFor = (relMap) => (rid) => {
    const r = relMap[rid];
    if (!r || r.ext) return null;
    if (mediaIds.has(r.target)) return mediaIds.get(r.target);
    const bytes = files[r.target];
    if (!bytes) return null;
    const id = uid('m');
    doc.media[id] = mediaDataUrl(bytes, r.target);
    mediaIds.set(r.target, id);
    return id;
  };
  ctx.media = mediaFor(rels);
  // 머리글 · 바닥글 파트 (자기 rels 로 그림)
  ctx.part = (rid, kind) => {
    const r = rels[rid];
    if (!r || !files[r.target]) return null;
    const sub = { ...ctx, rels: relsOf(files, r.target), fldStack: [] };
    sub.media = mediaFor(sub.rels);
    const x = parseXml(dec.decode(files[r.target]));
    const bl = readBlocks(x, sub);
    if (sub.watermark) ctx.watermark = sub.watermark;
    for (const b of bl) if (b.type === 'p' && b.style === 'Normal') b.style = kind === 'hdr' ? 'Header' : 'Footer';
    return bl.length ? bl : [para('', { style: kind === 'hdr' ? 'Header' : 'Footer' })];
  };
  readStyles(xmlOf(byType('styles')), ctx);
  readNumbering(xmlOf(byType('numbering')), ctx);
  readSettings(xmlOf(byType('settings')), ctx);
  // 본문
  const dx = parseXml(xmlOf(main));
  const bg = child(dx, 'background');
  if (bg && color(bg, ctx)) doc.settings.pageColor = color(bg, ctx);
  const body = child(dx, 'body');
  const lastSect = child(body, 'sectPr');
  const finalSec = readSect(lastSect, ctx);
  ctx.bodyW = finalSec.sec.page.w - finalSec.sec.margin.l - finalSec.sec.margin.r;
  doc.body = readBlocks(body, ctx);
  if (!doc.body.length) doc.body = [para('')];
  // 구역: Word 는 구역마다 머리글을 따로 안 적으면 앞 구역 것을 이어 씀 → 펼쳐 줌. 구역 시작 종류는 그 구역의 sectPr type
  const secParas = doc.body.filter((b) => b.type === 'p' && b.secBreak);
  const all = [...secParas.map((p) => ({ sec: p.secBreak.sec, kind: p.secBreak._kind, p })), { sec: finalSec.sec, kind: finalSec.kind }];
  for (let i = 1; i < all.length; i++) for (const prop of ['headers', 'footers']) for (const w of ['default', 'first', 'even']) if (!all[i].sec[prop][w] && all[i - 1].sec[prop][w]) all[i].sec[prop][w] = all[i - 1].sec[prop][w];
  for (const s of all) if (s.p) { s.p.secBreak = { kind: s.kind ?? 'nextPage', sec: s.sec }; }
  doc.section = finalSec.sec;
  if (finalSec.kind && finalSec.kind !== 'nextPage') doc.sectionKind = finalSec.kind;
  if (ctx.watermark?.text) doc.settings.watermark = ctx.watermark;
  // 각주 · 미주
  for (const [kind, type, tag] of [['foot', 'footnotes', 'footnote'], ['end', 'endnotes', 'endnote']]) {
    const p = byType(type);
    if (!p || !files[p]) continue;
    const sub = { ...ctx, rels: relsOf(files, p), fldStack: [] };
    sub.media = mediaFor(sub.rels);
    const x = parseXml(dec.decode(files[p]));
    for (const n of kids(x, tag)) {
      if (n.attrs['w:type'] && n.attrs['w:type'] !== 'normal') continue;
      const id = `${kind[0]}${n.attrs['w:id']}`;
      const bl = readBlocks(n, sub);
      // 맨 앞 번호 뒤 공백 정리
      const f = bl[0];
      if (f?.type === 'p' && f.runs[0] && !f.runs[0].obj && /^\s/.test(f.runs[0].text)) f.runs[0] = { ...f.runs[0], text: f.runs[0].text.replace(/^\s+/, '') };
      doc.notes[kind][id] = bl.length ? bl : [para('', { style: kind === 'foot' ? 'FootnoteText' : 'EndnoteText' })];
    }
  }
  // 메모
  const cp = byType('comments');
  if (cp && files[cp]) {
    const x = parseXml(dec.decode(files[cp]));
    doc.comments = kids(x, 'comment').map((c) => ({ id: `c${c.attrs['w:id']}`, author: c.attrs['w:author'] ?? '', initials: c.attrs['w:initials'] ?? '', date: c.attrs['w:date'] ?? '', text: kidsDeep(c, 'p').map((p) => kidsDeep(p, 't').map((t) => t.text).join('')).join('\n'), replies: [], done: false }));
    // 회신 (commentsExtended): paraIdParent 가 있으면 앞 메모의 회신으로
    const cx = byType('commentsExtended');
    if (cx && files[cx]) {
      const ext = kids(parseXml(dec.decode(files[cx])), 'commentEx');
      const paraOf = new Map(kids(x, 'comment').map((c, i) => [kidsDeep(c, 'p').slice(-1)[0]?.attrs['w14:paraId'], i]));
      const drop = new Set();
      for (const e of ext) {
        const i = paraOf.get(e.attrs['w15:paraId']);
        if (i == null) continue;
        if (e.attrs['w15:done'] === '1') doc.comments[i].done = true;
        const pi = paraOf.get(e.attrs['w15:paraIdParent']);
        if (pi != null && pi !== i) { doc.comments[pi].replies.push({ author: doc.comments[i].author, date: doc.comments[i].date, text: doc.comments[i].text }); drop.add(i); }
      }
      doc.comments = doc.comments.filter((_, i) => !drop.has(i));
    }
  }
  // 문서 속성
  const core = xmlOf('docProps/core.xml');
  if (core) {
    const x = parseXml(core);
    const g = (n) => child(x, n)?.text ?? '';
    doc.props = { ...doc.props, title: g('title'), subject: g('subject'), author: g('creator'), keywords: g('keywords'), comments: g('description'), created: g('created') || doc.props.created, modified: g('modified') || doc.props.modified, lastModifiedBy: g('lastModifiedBy') };
  }
  // WIWORD 가 저장한 원본 (본문 글이 그대로면 그것을 씀 — Word 에서 고쳤으면 무시)
  const own = files['word/wiword.json'];
  if (own) {
    try {
      const j = JSON.parse(dec.decode(own));
      if (j.text === hashText(docText(doc)) && j.doc) {
        const d2 = j.doc;
        d2.media = { ...doc.media, ...Object.fromEntries(Object.entries(j.mediaMap ?? {}).map(([id, target]) => [id, files[target] ? mediaDataUrl(files[target], target) : doc.media[id]])) };
        return { doc: d2, warnings };
      }
    } catch { /* Word 형식 그대로 */ }
  }
  if (Object.keys(files).some((k) => /vbaProject\.bin$/i.test(k))) warnings.push('매크로(VBA)가 들어 있습니다. 매크로는 실행하지 않으며, 저장하면 빠집니다.');
  if (Object.keys(files).some((k) => /^word\/charts\//.test(k))) warnings.push('차트는 회색 상자로 표시합니다 (원래 차트는 Word 에서 보세요).');
  return { doc, warnings };
}
function hashText(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return `${s.length}:${(h >>> 0).toString(36)}`; }

function readStyles(xml, ctx) {
  const doc = ctx.doc;
  if (!xml) return;
  const x = parseXml(xml);
  const dd = child(x, 'docDefaults');
  const rd = readRpr(child(child(dd, 'rPrDefault'), 'rPr'), ctx);
  const pd = readPpr(child(child(dd, 'pPrDefault'), 'pPr'), ctx);
  // Word 의 문서 기본값이 없으면 (오래된 파일) 10pt · 줄 간격 1 · 뒤 0
  doc.defaults = { rPr: { font: '+mn', fontEa: '+mn', sz: 10, lang: 'ko-KR', ...rd }, pPr: { align: 'left', spc: { before: 0, after: 0, line: 1, rule: 'auto' }, widow: true, ...pd, spc: { before: 0, after: 0, line: 1, rule: 'auto', ...(pd.spc ?? {}) } } };
  if (!rd.sz) doc.defaults.rPr.sz = 10;
  const styles = builtinStyles();
  // 기본 제공 스타일의 서식은 파일 것으로 덮어씀 (Word 버전마다 다름)
  const els = kids(x, 'style');
  for (const s of els) { const fid = s.attrs['w:styleId']; ctx.styleMap[fid] = styleIdFromName(av(child(s, 'name')), fid); }
  for (const s of els) {
    const fid = s.attrs['w:styleId'];
    const id = ctx.styleMap[fid];
    const type = { paragraph: 'para', character: 'char', table: 'table', numbering: 'numbering' }[s.attrs['w:type']] ?? 'para';
    if (type === 'numbering') continue;
    const name = av(child(s, 'name')) ?? id;
    const base = styles[id];
    const st = {
      id, type,
      name: base?.name ?? name,
      ...(child(s, 'basedOn') ? { basedOn: ctx.styleMap[av(child(s, 'basedOn'))] ?? av(child(s, 'basedOn')) } : {}),
      ...(child(s, 'next') ? { next: ctx.styleMap[av(child(s, 'next'))] ?? av(child(s, 'next')) } : {}),
      ...(child(s, 'link') ? { link: ctx.styleMap[av(child(s, 'link'))] ?? av(child(s, 'link')) } : {}),
      quick: !!child(s, 'qFormat'),
      hidden: !!child(s, 'semiHidden') || !!child(s, 'hidden'),
      order: Number(av(child(s, 'uiPriority')) ?? base?.order ?? 99),
      pPr: readPpr(child(s, 'pPr'), ctx),
      rPr: readRpr(child(s, 'rPr'), ctx),
    };
    if (type === 'table') {
      const tp = child(s, 'tblPr');
      st.tblPr = {};
      const bd = child(tp, 'tblBorders');
      if (bd) { const b = {}; for (const k of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']) { const xx = child(bd, k) ?? (k === 'left' ? child(bd, 'start') : k === 'right' ? child(bd, 'end') : null); if (xx) b[k] = border(xx, ctx); } st.tblPr.borders = b; }
      const mar = child(tp, 'tblCellMar');
      if (mar) { const m = {}; for (const [k, n] of [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom']]) { const xx = child(mar, n) ?? (n === 'left' ? child(mar, 'start') : n === 'right' ? child(mar, 'end') : null); if (xx) m[k] = tw(av(xx, 'w:w')); } st.tblPr.cellMar = m; }
      // 조건부 서식 (머리글 행 · 줄무늬): 사용자 표 스타일
      const cond = {};
      for (const c of kids(s, 'tblStylePr')) {
        const ty = c.attrs['w:type'];
        const shd = color(child(child(c, 'tcPr'), 'shd'), ctx, 'w:fill');
        const r = readRpr(child(c, 'rPr'), ctx);
        cond[ty] = { ...(shd ? { shd } : {}), ...(r.b ? { b: true } : {}), ...(r.color ? { color: r.color } : {}) };
      }
      if (Object.keys(cond).length) st.cond = cond;
    }
    if (st.basedOn === id) delete st.basedOn;
    if (s.attrs['w:default'] === '1' && type === 'para' && id !== 'Normal') { ctx.styleMap[fid] = 'Normal'; styles.Normal = { ...styles.Normal, pPr: st.pPr, rPr: st.rPr }; continue; }
    styles[id] = st;
  }
  doc.styles = styles;
}

function readNumbering(xml, ctx) {
  if (!xml) return;
  const x = parseXml(xml);
  const abs = {};
  for (const a of kids(x, 'abstractNum')) {
    const levels = [];
    for (const l of kids(a, 'lvl')) {
      const i = Number(l.attrs['w:ilvl'] ?? levels.length);
      const fmt = av(child(l, 'numFmt')) ?? 'decimal';
      const rp = readRpr(child(l, 'rPr'), ctx);
      let text = av(child(l, 'lvlText')) ?? '';
      const font = rp.font;
      if (fmt === 'bullet') text = symbolText(font, text) || text;
      if (fmt === 'bullet') text = ({ '': '•', '': '▪', '': '➢', '': '❖', '': '✓', '': '■', '': '◻', o: '○', '': '•' })[text] ?? text;
      const ind = readPpr(child(l, 'pPr'), ctx).ind ?? {};
      const FMT = { decimal: 'decimal', upperRoman: 'upperRoman', lowerRoman: 'lowerRoman', upperLetter: 'upperLetter', lowerLetter: 'lowerLetter', bullet: 'bullet', ganada: 'ganada', chosung: 'chosung', decimalEnclosedCircle: 'decimalCircle', decimalEnclosedCircleChinese: 'decimalCircle', koreanDigital: 'koreanDigital', koreanCounting: 'koreanCounting', koreanDigital2: 'koreanDigital', koreanLegal: 'koreanCounting', ideographDigital: 'ideographDigital', decimalZero: 'decimalZero', none: 'none', ordinal: 'decimal', cardinalText: 'decimal', ordinalText: 'decimal', decimalFullWidth: 'decimal', japaneseCounting: 'ideographDigital', chineseCounting: 'ideographDigital', taiwaneseCounting: 'ideographDigital', decimalEnclosedParen: 'decimal' };
      levels[i] = { fmt: FMT[fmt] ?? 'decimal', text, start: Number(av(child(l, 'start')) ?? 1), ind: { l: ind.l ?? 0, hanging: ind.hanging ?? (ind.first != null ? -ind.first : 0) }, ...(fmt === 'bullet' ? {} : {}), ...(child(l, 'isLgl') ? { legal: true } : {}), suff: av(child(l, 'suff')) ?? 'tab' };
      const lr = { ...rp };
      delete lr.font; delete lr.fontEa;
      if (Object.keys(lr).length) levels[i].rPr = lr;
      if (fmt === 'bullet' && font && !/symbol|wingdings/i.test(font)) levels[i].font = font;
    }
    for (let i = 0; i < 9; i++) if (!levels[i]) levels[i] = { fmt: 'decimal', text: `%${i + 1}.`, start: 1, ind: { l: 24 * (i + 1), hanging: 24 } };
    abs[a.attrs['w:abstractNumId']] = levels;
  }
  for (const n of kids(x, 'num')) {
    const id = n.attrs['w:numId'];
    const aid = av(child(n, 'abstractNumId'));
    if (!abs[aid]) continue;
    const def = { levels: abs[aid], abs: `a${aid}` };
    const ov = {};
    for (const o of kids(n, 'lvlOverride')) { const s = av(child(o, 'startOverride')); if (s != null) ov[o.attrs['w:ilvl']] = { start: Number(s) }; }
    if (Object.keys(ov).length) { def.override = ov; def.abs = `n${id}`; }
    ctx.doc.lists[String(id)] = def;
  }
}
function readSettings(xml, ctx) {
  if (!xml) return;
  const x = parseXml(xml);
  const s = ctx.doc.settings;
  if (child(x, 'evenAndOddHeaders') && onOff(child(x, 'evenAndOddHeaders'))) s.evenOdd = true;
  if (child(x, 'trackRevisions') && onOff(child(x, 'trackRevisions'))) s.track = true;
  const dt = av(child(x, 'defaultTabStop'));
  if (dt) s.defaultTab = tw(dt);
  if (child(x, 'autoHyphenation') && onOff(child(x, 'autoHyphenation'))) s.hyphenate = true;
  const dp = child(x, 'documentProtection');
  if (dp && dp.attrs['w:enforcement'] === '1') s.protect = { mode: dp.attrs['w:edit'] ?? 'readOnly' };
}

// ════════════════════════════════════ 쓰기 ════════════════════════════════════
const NS = 'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" mc:Ignorable="w14 w15 wp14"';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const T = (px) => Math.round((px ?? 0) * 15); // px → twip
const E = (px) => Math.round((px ?? 0) * 9525);
const hex = (c) => (c ? c.replace('#', '').toUpperCase().slice(0, 6) : 'auto');
const BORDER_OUT = { single: 'single', thick: 'thick', double: 'double', dotted: 'dotted', dashed: 'dashed', dotDash: 'dotDash', dotDotDash: 'dotDotDash', triple: 'triple', wave: 'wave', thinThickSmallGap: 'thinThickSmallGap', thickThinSmallGap: 'thickThinSmallGap', none: 'nil', nil: 'nil' };
function borderXml(tag, b) {
  if (!b) return '';
  if (b.style === 'none' || b.style === 'nil') return `<w:${tag} w:val="nil"/>`;
  return `<w:${tag} w:val="${BORDER_OUT[b.style] ?? 'single'}" w:sz="${Math.max(2, Math.round((b.w ?? 0.5) * 8))}" w:space="${Math.round(b.space ?? 0)}" w:color="${hex(b.color ?? '#000000')}"/>`;
}
function fontsXml(rp) {
  if (!rp.font && !rp.fontEa) return '';
  const a = [];
  if (rp.font === '+mj' || rp.font === '+mn') a.push(`w:asciiTheme="${rp.font === '+mj' ? 'major' : 'minor'}HAnsi" w:hAnsiTheme="${rp.font === '+mj' ? 'major' : 'minor'}HAnsi"`);
  else if (rp.font) a.push(`w:ascii="${esc(rp.font)}" w:hAnsi="${esc(rp.font)}" w:cs="${esc(rp.font)}"`);
  if (rp.fontEa === '+mj' || rp.fontEa === '+mn') a.push(`w:eastAsiaTheme="${rp.fontEa === '+mj' ? 'major' : 'minor'}EastAsia"`);
  else if (rp.fontEa) a.push(`w:eastAsia="${esc(rp.fontEa)}"`);
  return `<w:rFonts ${a.join(' ')}/>`;
}
function rprXml(rp, { styleIds } = {}) {
  let s = '';
  if (rp.cstyle) s += `<w:rStyle w:val="${esc(styleIds?.[rp.cstyle] ?? rp.cstyle)}"/>`;
  s += fontsXml(rp);
  const tog = (k, tag) => { if (rp[k] === true) s += `<w:${tag}/>`; else if (rp[k] === false) s += `<w:${tag} w:val="0"/>`; };
  tog('b', 'b'); tog('b', 'bCs'); tog('i', 'i'); tog('i', 'iCs'); tog('caps', 'caps'); tog('smallCaps', 'smallCaps'); tog('s', 'strike'); tog('dstrike', 'dstrike'); tog('outline', 'outline'); tog('shadow', 'shadow'); tog('emboss', 'emboss'); tog('imprint', 'imprint'); tog('hidden', 'vanish');
  if (rp.color) s += `<w:color w:val="${hex(rp.color)}"/>`;
  if (rp.spc) s += `<w:spacing w:val="${Math.round(rp.spc * 20)}"/>`;
  if (rp.scale && rp.scale !== 100) s += `<w:w w:val="${Math.round(rp.scale)}"/>`;
  if (rp.kern) s += `<w:kern w:val="${Math.round(rp.kern * 2)}"/>`;
  if (rp.pos) s += `<w:position w:val="${Math.round(rp.pos * 2)}"/>`;
  if (rp.sz) s += `<w:sz w:val="${Math.round(rp.sz * 2)}"/><w:szCs w:val="${Math.round(rp.sz * 2)}"/>`;
  if (rp.hl) { const n = HL_OUT[rp.hl]; if (n) s += `<w:highlight w:val="${n}"/>`; else s += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(rp.hl)}"/>`; }
  if (rp.u) s += `<w:u w:val="${rp.u === true ? 'single' : rp.u}"${rp.uColor ? ` w:color="${hex(rp.uColor)}"` : ''}/>`;
  else if (rp.u === false) s += '<w:u w:val="none"/>';
  if (rp.shd && !rp.hl) s += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(rp.shd)}"/>`;
  if (rp.va) s += `<w:vertAlign w:val="${rp.va === 'super' ? 'superscript' : 'subscript'}"/>`;
  if (rp.rtl) s += '<w:rtl/>';
  if (rp.emph) s += `<w:em w:val="${rp.emph === 'comma' ? 'comma' : rp.emph === 'circle' ? 'circle' : 'dot'}"/>`;
  if (rp.lang) s += `<w:lang w:eastAsia="${esc(rp.lang)}"/>`;
  return s ? `<w:rPr>${s}</w:rPr>` : '';
}
function pprBody(pp, { styleIds, numIds } = {}) {
  let s = '';
  if (pp.keepNext) s += '<w:keepNext/>';
  if (pp.keepLines) s += '<w:keepLines/>';
  if (pp.pageBreakBefore) s += '<w:pageBreakBefore/>';
  if (pp.widow === false) s += '<w:widowControl w:val="0"/>';
  if (pp.num) s += pp.num.id == null ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>' : `<w:numPr><w:ilvl w:val="${pp.num.lvl ?? 0}"/><w:numId w:val="${numIds?.[pp.num.id] ?? pp.num.id}"/></w:numPr>`;
  if (pp.border) s += `<w:pBdr>${['top', 'left', 'bottom', 'right', 'between'].map((k) => borderXml(k, pp.border[k])).join('')}</w:pBdr>`;
  if (pp.shd) s += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(pp.shd)}"/>`;
  if (pp.tabs?.length) s += `<w:tabs>${pp.tabs.map((t) => `<w:tab w:val="${t.align === 'decimal' ? 'decimal' : t.align ?? 'left'}"${t.leader ? ` w:leader="${t.leader}"` : ''} w:pos="${T(t.pos)}"/>`).join('')}</w:tabs>`;
  if (pp.noHyphen) s += '<w:suppressAutoHyphens/>';
  if (pp.bidi) s += '<w:bidi/>';
  if (pp.snapGrid === false) s += '<w:snapToGrid w:val="0"/>';
  if (pp.spc) {
    const sp = pp.spc;
    const a = [];
    if (sp.before != null) a.push(`w:before="${T(sp.before)}"`);
    if (sp.after != null) a.push(`w:after="${T(sp.after)}"`);
    if (sp.line != null) {
      if (sp.rule === 'exact' || sp.rule === 'atLeast') a.push(`w:line="${T(sp.line)}" w:lineRule="${sp.rule}"`);
      else a.push(`w:line="${Math.round(sp.line * 240)}" w:lineRule="auto"`);
    }
    if (a.length) s += `<w:spacing ${a.join(' ')}/>`;
  }
  if (pp.ind) {
    const i = pp.ind; const a = [];
    if (i.l != null) a.push(`w:left="${T(i.l)}"`);
    if (i.r != null) a.push(`w:right="${T(i.r)}"`);
    if (i.hanging) a.push(`w:hanging="${T(i.hanging)}"`); else if (i.first != null) a.push(`w:firstLine="${T(i.first)}"`);
    if (a.length) s += `<w:ind ${a.join(' ')}/>`;
  }
  if (pp.contextual) s += '<w:contextualSpacing/>';
  if (pp.align) s += `<w:jc w:val="${pp.align === 'justify' ? 'both' : pp.align}"/>`;
  if (pp.outline != null) s += `<w:outlineLvl w:val="${pp.outline}"/>`;
  return s;
}
function sectXml(sec, kind, W) {
  let s = '<w:sectPr>';
  for (const [prop, tag] of [['headers', 'header'], ['footers', 'footer']]) for (const which of ['default', 'first', 'even']) { const bl = sec[prop]?.[which]; if (bl) s += `<w:${tag}Reference w:type="${which}" r:id="${W.hf(bl, tag)}"/>`; }
  if (kind && kind !== 'nextPage') s += `<w:type w:val="${kind}"/>`;
  s += `<w:pgSz w:w="${T(sec.page.w)}" w:h="${T(sec.page.h)}"${sec.page.orient === 'landscape' ? ' w:orient="landscape"' : ''}/>`;
  const m = sec.margin;
  s += `<w:pgMar w:top="${T(m.t)}" w:right="${T(m.r)}" w:bottom="${T(m.b)}" w:left="${T(m.l)}" w:header="${T(m.header ?? 48)}" w:footer="${T(m.footer ?? 48)}" w:gutter="${T(m.gutter ?? 0)}"/>`;
  if (sec.pageBorder) { const b = sec.pageBorder; s += `<w:pgBorders w:offsetFrom="page">${['top', 'left', 'bottom', 'right'].map((k) => borderXml(k, { ...b, space: b.space ?? 24 })).join('')}</w:pgBorders>`; }
  if (sec.lineNumbers) s += `<w:lnNumType w:countBy="${sec.lineNumbers.countBy ?? 1}" w:restart="${sec.lineNumbers.restart ?? 'newPage'}"/>`;
  if (sec.pgNum) s += `<w:pgNumType${sec.pgNum.fmt && sec.pgNum.fmt !== 'decimal' ? ` w:fmt="${sec.pgNum.fmt}"` : ''}${sec.pgNum.start != null ? ` w:start="${sec.pgNum.start}"` : ''}/>`;
  const c = sec.cols ?? { n: 1 };
  if (c.widths?.length) s += `<w:cols w:num="${c.n}" w:space="${T(c.space)}" w:equalWidth="0"${c.sep ? ' w:sep="1"' : ''}>${c.widths.map((w) => `<w:col w:w="${T(w)}" w:space="${T(c.space)}"/>`).join('')}</w:cols>`;
  else s += `<w:cols w:space="${T(c.space ?? cm(1.25))}"${c.n > 1 ? ` w:num="${c.n}"` : ''}${c.sep ? ' w:sep="1"' : ''}/>`;
  if (sec.vAlign && sec.vAlign !== 'top') s += `<w:vAlign w:val="${sec.vAlign === 'justify' ? 'both' : sec.vAlign}"/>`;
  if (sec.titlePg) s += '<w:titlePg/>';
  s += '<w:docGrid w:linePitch="360"/></w:sectPr>';
  return s;
}

/** 문서 → { 경로: 문자열|바이트 } */
export function writeDocx(doc, { template = false } = {}) {
  const files = {};
  const rels = [];
  let rid = 0;
  const addRel = (type, target, ext = false) => { const id = `rId${++rid}`; rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`); return id; };
  const ct = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
  const over = [];
  const mediaOut = new Map(); // media id → { name, target }
  const mediaMap = {};
  let imgN = 0;
  const hfParts = [];
  // 스타일 · 목록 ID
  const styleIds = {};
  for (const id of Object.keys(doc.styles)) styleIds[id] = id.replace(/[^A-Za-z0-9_-]/g, '') || `s${Object.keys(styleIds).length}`;
  const numIds = {};
  let nn = 0;
  for (const k of Object.keys(doc.lists ?? {})) numIds[k] = ++nn;
  // 쓰는 중 상태 (본문 · 머리글 파트마다 rels 가 다름)
  const W = {
    rels: null,
    media(r) {
      const src = r.media ? doc.media?.[r.media] : r.src;
      if (!src) return null;
      if (/^https?:/.test(src)) return { id: W.rels.add('image', src, true), link: true };
      const key = r.media ?? src;
      let m = mediaOut.get(key);
      if (!m) {
        const d = dataUrlBytes(src);
        if (!d) return null;
        const name = `image${++imgN}.${d.ext}`;
        files[`word/media/${name}`] = d.bytes;
        ct.set(d.ext, MIME[d.ext] ?? d.mime);
        m = { name };
        mediaOut.set(key, m);
        if (r.media) mediaMap[r.media] = `word/media/${name}`;
      }
      return { id: W.rels.add('image', `media/${m.name}`) };
    },
    hf(blocks, tag) {
      let part = hfParts.find((p) => p.blocks === blocks);
      if (!part) {
        const n = hfParts.filter((p) => p.tag === tag).length + 1;
        part = { blocks, tag, name: `${tag}${n}.xml`, rels: relSet(`word/_rels/${tag}${n}.xml.rels`) };
        hfParts.push(part);
        const prev = W.rels; W.rels = part.rels;
        const body = blocksXml(blocks, W);
        W.rels = prev;
        files[`word/${part.name}`] = `${XML_HEAD}<w:${tag === 'header' ? 'hdr' : 'ftr'} ${NS}>${body || '<w:p/>'}</w:${tag === 'header' ? 'hdr' : 'ftr'}>`;
        over.push([`/word/${part.name}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${tag}+xml`]);
        part.id = addRel(tag, part.name);
      }
      return part.id;
    },
    styleIds, numIds, doc,
    noteIds: new Map(), notes: { foot: [], end: [] }, cmtIds: new Map(), bmIds: new Map(), drawId: 0,
  };
  const relSet = (path) => {
    const list = [];
    let n = 0;
    const set = { path, list, add(type, target, ext = false) { const id = `rId${++n}`; list.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`); return id; } };
    return set;
  };
  // 본문 rels 는 문서 rels 와 같음
  W.rels = { add: (type, target, ext) => addRel(type, target, ext) };
  // 구역: 구역 나누기 문단 + 마지막 구역
  let body = blocksXml(doc.body, W);
  body += sectXml(doc.section, doc.sectionKind, W);
  const bg = doc.settings?.pageColor ? `<w:background w:color="${hex(doc.settings.pageColor)}"/>` : '';
  files['word/document.xml'] = `${XML_HEAD}<w:document ${NS}>${bg}<w:body>${body}</w:body></w:document>`;
  // 머리글 파트 rels
  for (const p of hfParts) if (p.rels.list.length) files[`word/_rels/${p.name}.rels`] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${p.rels.list.join('')}</Relationships>`;
  // 워터마크: 첫 구역 기본 머리글에 VML 글자 (없으면 머리글 만들어 넣음) — 간단히 settings 로만 보관하고 VML 은 머리글 안에
  // 스타일 · 번호 · 설정 · 테마
  files['word/styles.xml'] = stylesXml(doc, W);
  addRel('styles', 'styles.xml');
  over.push(['/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml']);
  if (nn) {
    files['word/numbering.xml'] = numberingXml(doc, numIds);
    addRel('numbering', 'numbering.xml');
    over.push(['/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml']);
  }
  files['word/settings.xml'] = settingsXml(doc);
  addRel('settings', 'settings.xml');
  over.push(['/word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml']);
  files['word/theme/theme1.xml'] = themeXml(doc);
  addRel('theme', 'theme/theme1.xml');
  over.push(['/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml']);
  files['word/fontTable.xml'] = `${XML_HEAD}<w:fonts ${NS}><w:font w:name="맑은 고딕"><w:panose1 w:val="020B0503020000020004"/><w:charset w:val="81"/><w:family w:val="modern"/><w:pitch w:val="variable"/></w:font></w:fonts>`;
  addRel('fontTable', 'fontTable.xml');
  over.push(['/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml']);
  // 각주 · 미주 (구분선 -1 · 0 은 Word 필수)
  for (const [kind, tag] of [['foot', 'footnote'], ['end', 'endnote']]) {
    const list = W.notes[kind];
    const nrels = relSet(`word/_rels/${tag}s.xml.rels`);
    const prev = W.rels; W.rels = nrels;
    const parts = list.map(({ id, num }) => {
      const bl = doc.notes?.[kind]?.[id] ?? [para('')];
      let x = blocksXml(bl, W);
      // 첫 문단 맨 앞에 번호 표시
      x = x.replace(/(<w:p\b[^>]*>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?)/, `$1<w:r><w:rPr><w:rStyle w:val="${kind === 'foot' ? 'FootnoteReference' : 'EndnoteReference'}"/><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`);
      return `<w:${tag} w:id="${num}">${x}</w:${tag}>`;
    });
    W.rels = prev;
    if (!list.length && kind === 'end' && !W.notes.foot.length) continue;
    files[`word/${tag}s.xml`] = `${XML_HEAD}<w:${tag}s ${NS}><w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>${parts.join('')}</w:${tag}s>`;
    if (nrels.list.length) files[`word/_rels/${tag}s.xml.rels`] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${nrels.list.join('')}</Relationships>`;
    addRel(`${tag}s`, `${tag}s.xml`);
    over.push([`/word/${tag}s.xml`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${tag}s+xml`]);
  }
  // 메모
  if ((doc.comments ?? []).length) {
    const xs = doc.comments.map((c) => {
      const num = W.cmtIds.get(c.id) ?? W.cmtIds.size;
      if (!W.cmtIds.has(c.id)) W.cmtIds.set(c.id, num);
      const ps = String(c.text ?? '').split('\n').map((l, i) => `<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>${i === 0 ? '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' : ''}<w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`).join('');
      const reps = (c.replies ?? []).map((r) => `<w:p><w:r><w:t xml:space="preserve">${esc(`↳ ${r.author ?? ''}: ${r.text ?? ''}`)}</w:t></w:r></w:p>`).join('');
      return `<w:comment w:id="${num}" w:author="${esc(c.author ?? '')}" w:date="${esc(c.date || new Date().toISOString())}" w:initials="${esc((c.author ?? '').slice(0, 2))}">${ps}${reps}</w:comment>`;
    });
    files['word/comments.xml'] = `${XML_HEAD}<w:comments ${NS}>${xs.join('')}</w:comments>`;
    addRel('comments', 'comments.xml');
    over.push(['/word/comments.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml']);
  }
  // WIWORD 원본 (다시 열 때 정확히 복원)
  const lite = { ...doc, media: Object.fromEntries(Object.keys(doc.media ?? {}).filter((k) => !mediaMap[k]).map((k) => [k, doc.media[k]])) };
  files['word/wiword.json'] = JSON.stringify({ wiword: 1, text: hashText(docText(doc)), mediaMap, doc: lite });
  rels.push(`<Relationship Id="rId${++rid}" Type="http://wiword.app/relationships/source" Target="wiword.json"/>`);
  ct.set('json', 'application/json');
  // 문서 속성
  const p = doc.props ?? {};
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  files['docProps/core.xml'] = `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(p.title ?? '')}</dc:title><dc:subject>${esc(p.subject ?? '')}</dc:subject><dc:creator>${esc(p.author ?? '')}</dc:creator><cp:keywords>${esc(p.keywords ?? '')}</cp:keywords><dc:description>${esc(p.comments ?? '')}</dc:description><cp:lastModifiedBy>${esc(p.lastModifiedBy ?? p.author ?? '')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${esc((p.created ?? now).replace(/\.\d+Z$/, 'Z'))}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  files['docProps/app.xml'] = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>WIWORD</Application><DocSecurity>0</DocSecurity></Properties>`;
  over.push(['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'], ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']);
  const mainType = template ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  files['[Content_Types].xml'] = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[...ct].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('')}<Override PartName="/word/document.xml" ContentType="${mainType}"/>${over.map(([pn, t]) => `<Override PartName="${pn}" ContentType="${t}"/>`).join('')}</Types>`;
  files['_rels/.rels'] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  files['word/_rels/document.xml.rels'] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`;
  // 순서: [Content_Types] 를 맨 앞에
  const ordered = { '[Content_Types].xml': files['[Content_Types].xml'], '_rels/.rels': files['_rels/.rels'] };
  for (const [k, v] of Object.entries(files)) if (!(k in ordered)) ordered[k] = v;
  return ordered;
}

function blocksXml(blocks, W) {
  let s = '';
  for (const b of blocks ?? []) s += b.type === 'tbl' ? tableXml(b, W) : paraXml(b, W);
  return s;
}
function paraXml(p, W) {
  let ppr = '';
  if (p.style && p.style !== 'Normal') ppr += `<w:pStyle w:val="${esc(W.styleIds[p.style] ?? p.style)}"/>`;
  ppr += pprBody(p.pPr ?? {}, W);
  const markRpr = p.mark ? rprXml(p.mark, W) : '';
  const insMark = p.paraIns ? '<w:ins w:id="999" w:author="WIWORD" w:date="2024-01-01T00:00:00Z"/>' : '';
  if (markRpr || insMark) ppr += markRpr ? markRpr.replace('<w:rPr>', `<w:rPr>${insMark}`) : `<w:rPr>${insMark}</w:rPr>`;
  if (p.secBreak) ppr += sectXml(p.secBreak.sec, p.secBreak.kind, W);
  let s = `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}`;
  let link = null; // 현재 열린 하이퍼링크
  let rev = null;  // 현재 열린 w:ins/w:del
  let revId = 0;
  const closeLink = () => { if (link) { s += '</w:hyperlink>'; link = null; } };
  const closeRev = () => { if (rev) { s += `</w:${rev.tag}>`; rev = null; } };
  const sameRev = (a, b) => a && b && a.author === b.author && a.date === b.date;
  for (const r of p.runs) {
    // 변경 내용 묶음
    const tag = r.del ? 'del' : r.ins ? 'ins' : null;
    const info = r.del ?? r.ins;
    if (!tag || !rev || rev.tag !== tag || !sameRev(rev.info, info)) {
      closeLink();
      closeRev();
      if (tag) { s += `<w:${tag} w:id="${++revId + 1000}" w:author="${esc(info.author ?? '')}" w:date="${esc(info.date || new Date().toISOString())}">`; rev = { tag, info }; }
    }
    // 하이퍼링크 묶음
    const lk = r.link ? JSON.stringify(r.link) : null;
    if (lk !== (link ? JSON.stringify(link) : null)) {
      closeLink();
      if (r.link) {
        if (r.link.url) s += `<w:hyperlink r:id="${W.rels.add('hyperlink', r.link.url, true)}" w:history="1">`;
        else s += `<w:hyperlink w:anchor="${esc(r.link.anchor ?? '')}" w:history="1">`;
        link = r.link;
      }
    }
    s += runXml(r, W, !!r.del);
  }
  closeLink();
  closeRev();
  return `${s}</w:p>`;
}
function runXml(r, W, del) {
  const { text: _t, obj: _o, link: _l, ins: _i, del: _d, ...props } = r;
  const rp = (extra = {}) => rprXml({ ...props, ...extra, ...(r.link && !props.cstyle ? { cstyle: 'Hyperlink' } : {}) }, W);
  if (!r.obj) {
    const parts = String(r.text ?? '').split(/(\t|\n)/);
    let inner = '';
    for (const t of parts) {
      if (t === '\t') inner += '<w:tab/>';
      else if (t === '\n') inner += '<w:br/>';
      else if (t) inner += del ? `<w:delText xml:space="preserve">${esc(t)}</w:delText>` : `<w:t xml:space="preserve">${esc(t)}</w:t>`;
    }
    return inner ? `<w:r>${rp()}${inner}</w:r>` : '';
  }
  switch (r.obj) {
    case 'tab': return `<w:r>${rp()}<w:tab/></w:r>`;
    case 'br': return `<w:r><w:br${r.kind === 'page' ? ' w:type="page"' : r.kind === 'column' ? ' w:type="column"' : ''}/></w:r>`;
    case 'sym': return `<w:r>${rp()}<w:t xml:space="preserve">${esc(r.char ?? '')}</w:t></w:r>`;
    case 'note': {
      const kind = r.kind === 'end' ? 'end' : 'foot';
      const tag = kind === 'foot' ? 'footnote' : 'endnote';
      let num = W.noteIds.get(r.id);
      if (num == null) { num = W.notes[kind].length + 1; W.noteIds.set(r.id, num); W.notes[kind].push({ id: r.id, num }); }
      return `<w:r><w:rPr><w:rStyle w:val="${kind === 'foot' ? 'FootnoteReference' : 'EndnoteReference'}"/><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Reference w:id="${num}"/></w:r>`;
    }
    case 'field': {
      const code = r.code ?? '';
      const res = String(r.result ?? '');
      return `<w:r>${rp()}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rp()}<w:instrText xml:space="preserve"> ${esc(code)} </w:instrText></w:r><w:r>${rp()}<w:fldChar w:fldCharType="separate"/></w:r>${res ? `<w:r>${rp()}<w:t xml:space="preserve">${esc(res)}</w:t></w:r>` : ''}<w:r>${rp()}<w:fldChar w:fldCharType="end"/></w:r>`;
    }
    case 'mark': {
      if (r.kind === 'bm') {
        let id = W.bmIds.get(r.name);
        if (id == null) { id = W.bmIds.size + 1; W.bmIds.set(r.name, id); }
        return r.edge === 'start' ? `<w:bookmarkStart w:id="${id}" w:name="${esc(r.name)}"/>` : `<w:bookmarkEnd w:id="${id}"/>`;
      }
      if (r.kind === 'cmt') {
        let id = W.cmtIds.get(r.id);
        if (id == null) { id = W.cmtIds.size; W.cmtIds.set(r.id, id); }
        return r.edge === 'start' ? `<w:commentRangeStart w:id="${id}"/>` : `<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
      }
      return '';
    }
    case 'eq': {
      let om = '';
      try { om = toOMML(parseLatex(r.latex ?? '')); } catch { om = ''; }
      om = om.replace(/<a:rPr\b[\s\S]*?<\/a:rPr>/g, '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>').replace(/ xmlns:m="[^"]*"/, '').replace(/^<m:oMathPara>/, '').replace(/<\/m:oMathPara>$/, '');
      return om;
    }
    case 'img': return drawingXml(r, W);
    case 'shape': return shapeXml(r, W);
    default: return '';
  }
}
function drawingXml(r, W) {
  const m = r.svg ? W.media({ src: `data:image/svg+xml;base64,${b64(new TextEncoder().encode(r.svg))}` }) : W.media(r);
  if (!m) return '';
  const id = ++W.drawId;
  const cx = E(r.w); const cy = E(r.h);
  const crop = r.crop ? `<a:srcRect l="${Math.round(r.crop.l * 100000)}" t="${Math.round(r.crop.t * 100000)}" r="${Math.round(r.crop.r * 100000)}" b="${Math.round(r.crop.b * 100000)}"/>` : '';
  const blip = m.link ? `<a:blip r:link="${m.id}"/>` : `<a:blip r:embed="${m.id}"/>`;
  const pic = `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="그림 ${id}" descr="${esc(r.alt ?? '')}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>${blip}${crop}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm${r.rot ? ` rot="${Math.round(r.rot * 60000)}"` : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  return `<w:r><w:drawing>${anchorOpen(r, id, cx, cy, `그림 ${id}`, r.alt)}${pic}${r.float ? '</wp:anchor>' : '</wp:inline>'}</w:drawing></w:r>`;
}
function anchorOpen(r, id, cx, cy, name, alt) {
  const doc = `<wp:docPr id="${id}" name="${esc(name)}" descr="${esc(alt ?? '')}"/><wp:cNvGraphicFramePr/>`;
  if (!r.float) return `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${doc}`;
  const f = r.float;
  const relH = { page: 'page', margin: 'margin', column: 'column' }[f.relH] ?? 'column';
  const relV = { page: 'page', margin: 'margin', paragraph: 'paragraph' }[f.relV] ?? 'paragraph';
  const ph = f.alignH ? `<wp:align>${f.alignH}</wp:align>` : `<wp:posOffset>${E(f.x)}</wp:posOffset>`;
  const pv = f.alignV ? `<wp:align>${f.alignV}</wp:align>` : `<wp:posOffset>${E(f.y)}</wp:posOffset>`;
  const wrap = f.wrap === 'square' ? '<wp:wrapSquare wrapText="bothSides"/>' : f.wrap === 'tight' ? '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>' : f.wrap === 'topBottom' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapNone/>';
  const d = f.dist ?? {};
  return `<wp:anchor distT="${E(d.t ?? 0)}" distB="${E(d.b ?? 0)}" distL="${E(d.l ?? 12)}" distR="${E(d.r ?? 12)}" simplePos="0" relativeHeight="${251658240 + id}" behindDoc="${f.wrap === 'behind' ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${relH}">${ph}</wp:positionH><wp:positionV relativeFrom="${relV}">${pv}</wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}${doc}`;
}
function shapeXml(r, W) {
  const id = ++W.drawId;
  const cx = E(r.w); const cy = E(r.h);
  const fl = r.float ?? { wrap: 'front', relH: 'column', relV: 'paragraph', x: 0, y: 0 };
  const fill = r.fill ? `<a:solidFill><a:srgbClr val="${hex(r.fill)}"/></a:solidFill>` : '<a:noFill/>';
  const ln = r.line ? `<a:ln w="${Math.round((r.line.w ?? 0.75) * 12700)}"><a:solidFill><a:srgbClr val="${hex(r.line.color ?? '#000000')}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const ins = r.inset ?? { l: 9.6, r: 9.6, t: 4.8, b: 4.8 };
  const txt = blocksXml(r.blocks?.length ? r.blocks : [para('')], W);
  return `<w:r><w:drawing>${anchorOpen({ ...r, float: fl }, id, cx, cy, `글상자 ${id}`, '')}<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${r.prst ?? 'rect'}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr><wps:txbx><w:txbxContent>${txt}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${E(ins.l)}" tIns="${E(ins.t)}" rIns="${E(ins.r)}" bIns="${E(ins.b)}" anchor="${r.anchor === 'ctr' ? 'ctr' : 't'}" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}
function tableXml(t, W) {
  const tp = t.tblPr ?? {};
  let pr = `<w:tblStyle w:val="${esc(W.styleIds[t.style] ?? t.style ?? 'TableGrid')}"/>`;
  const total = t.grid.reduce((a, b) => a + b, 0);
  pr += `<w:tblW w:w="${T(total)}" w:type="dxa"/>`;
  if (tp.align) pr += `<w:jc w:val="${tp.align}"/>`;
  if (tp.ind) pr += `<w:tblInd w:w="${T(tp.ind)}" w:type="dxa"/>`;
  if (tp.borders) pr += `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => borderXml(k, tp.borders[k])).join('')}</w:tblBorders>`;
  if (tp.shd) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(tp.shd)}"/>`;
  pr += '<w:tblLayout w:type="fixed"/>';
  if (tp.cellMar) pr += `<w:tblCellMar>${[['t', 'top'], ['l', 'left'], ['b', 'bottom'], ['r', 'right']].filter(([k]) => tp.cellMar[k] != null).map(([k, n]) => `<w:${n} w:w="${T(tp.cellMar[k])}" w:type="dxa"/>`).join('')}</w:tblCellMar>`;
  const lk = { firstRow: true, firstCol: true, bandRows: true, ...(tp.look ?? {}) };
  pr += `<w:tblLook w:val="04A0" w:firstRow="${lk.firstRow ? 1 : 0}" w:lastRow="${lk.lastRow ? 1 : 0}" w:firstColumn="${lk.firstCol ? 1 : 0}" w:lastColumn="${lk.lastCol ? 1 : 0}" w:noHBand="${lk.bandRows ? 0 : 1}" w:noVBand="${lk.bandCols ? 0 : 1}"/>`;
  let s = `<w:tbl><w:tblPr>${pr}</w:tblPr><w:tblGrid>${t.grid.map((g) => `<w:gridCol w:w="${T(g)}"/>`).join('')}</w:tblGrid>`;
  for (const row of t.rows) {
    let trPr = '';
    if (row.cantSplit) trPr += '<w:cantSplit/>';
    if (row.h) trPr += `<w:trHeight w:val="${T(row.h)}"${row.hRule === 'exact' ? ' w:hRule="exact"' : ''}/>`;
    if (row.header) trPr += '<w:tblHeader/>';
    s += `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}`;
    let gc = 0;
    for (const c of row.cells) {
      const span = c.span ?? 1;
      const w = t.grid.slice(gc, gc + span).reduce((a, b) => a + b, 0);
      gc += span;
      let cp = `<w:tcW w:w="${T(w)}" w:type="dxa"/>`;
      if (span > 1) cp += `<w:gridSpan w:val="${span}"/>`;
      if (c.vMerge) cp += c.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>';
      if (c.borders) cp += `<w:tcBorders>${['top', 'left', 'bottom', 'right'].map((k) => borderXml(k, c.borders[k])).join('')}</w:tcBorders>`;
      if (c.shd) cp += `<w:shd w:val="clear" w:color="auto" w:fill="${hex(c.shd)}"/>`;
      if (c.mar) cp += `<w:tcMar>${[['t', 'top'], ['l', 'left'], ['b', 'bottom'], ['r', 'right']].filter(([k]) => c.mar[k] != null).map(([k, n]) => `<w:${n} w:w="${T(c.mar[k])}" w:type="dxa"/>`).join('')}</w:tcMar>`;
      if (c.dir === 'vert') cp += '<w:textDirection w:val="tbRl"/>';
      if (c.vAlign && c.vAlign !== 'top') cp += `<w:vAlign w:val="${c.vAlign}"/>`;
      const inner = blocksXml(c.blocks, W);
      s += `<w:tc><w:tcPr>${cp}</w:tcPr>${/<w:p[ >]|<w:p\/>/.test(inner) && /<\/w:p>$|<w:p\/>$/.test(inner) ? inner : `${inner}<w:p/>`}</w:tc>`;
    }
    s += '</w:tr>';
  }
  return `${s}</w:tbl>`;
}

function stylesXml(doc, W) {
  const d = doc.defaults ?? {};
  let s = `${XML_HEAD}<w:styles ${NS}><w:docDefaults><w:rPrDefault>${rprXml(d.rPr ?? {}, W) || '<w:rPr/>'}</w:rPrDefault><w:pPrDefault><w:pPr>${pprBody(d.pPr ?? {}, W)}</w:pPr></w:pPrDefault></w:docDefaults>`;
  s += '<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376"/>';
  const ensure = { CommentText: { id: 'CommentText', name: 'annotation text', type: 'para', basedOn: 'Normal', pPr: {}, rPr: { sz: 10 } }, CommentReference: { id: 'CommentReference', name: 'annotation reference', type: 'char', rPr: { sz: 9 } } };
  const all = { ...ensure, ...doc.styles };
  const usedTables = new Set();
  walkParas(doc.body, () => {});
  const visitT = (bs) => { for (const b of bs) if (b.type === 'tbl') { usedTables.add(b.style); for (const r of b.rows) for (const c of r.cells) visitT(c.blocks); } };
  visitT(doc.body);
  for (const st of Object.values(all)) {
    const id = st.id ?? Object.keys(all).find((k) => all[k] === st);
    const type = { para: 'paragraph', char: 'character', table: 'table' }[st.type] ?? 'paragraph';
    const fileName = ID_NAME[id] ?? st.name ?? id;
    const def = (id === 'Normal' && type === 'paragraph') || (id === 'DefaultParagraphFont') || (id === 'TableNormal');
    s += `<w:style w:type="${type}"${def ? ' w:default="1"' : ''} w:styleId="${esc(W.styleIds[id] ?? id)}"><w:name w:val="${esc(fileName)}"/>`;
    if (st.basedOn && all[st.basedOn]) s += `<w:basedOn w:val="${esc(W.styleIds[st.basedOn] ?? st.basedOn)}"/>`;
    if (st.next && all[st.next]) s += `<w:next w:val="${esc(W.styleIds[st.next] ?? st.next)}"/>`;
    if (st.link && all[st.link]) s += `<w:link w:val="${esc(W.styleIds[st.link] ?? st.link)}"/>`;
    if (st.order != null) s += `<w:uiPriority w:val="${Math.round(st.order)}"/>`;
    if (st.hidden) s += '<w:semiHidden/><w:unhideWhenUsed/>';
    if (st.quick) s += '<w:qFormat/>';
    const pp = pprBody(st.pPr ?? {}, W);
    if (pp && type !== 'character') s += `<w:pPr>${pp}</w:pPr>`;
    const rp = rprXml({ ...(st.rPr ?? {}), cstyle: undefined }, W);
    if (rp) s += rp;
    if (type === 'table') {
      const tp = st.tblPr ?? {};
      s += `<w:tblPr>${tp.borders ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => borderXml(k, tp.borders[k])).join('')}</w:tblBorders>` : ''}<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${T(tp.cellMar?.l ?? 7.2)}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${T(tp.cellMar?.r ?? 7.2)}" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
    }
    s += '</w:style>';
  }
  // 기본 제공 표 스타일 (눈금 표 4 - 강조 1 …): 머리글 행 · 줄무늬 · 테두리를 Word 조건부 서식으로
  for (const id of usedTables) {
    if (!id || all[id] || !/^(Grid|List|Plain)Table/.test(id)) continue;
    const fake = { style: id, rows: Array.from({ length: 6 }, () => ({ cells: [{}, {}, {}] })), tblPr: { look: { firstRow: true, firstCol: true, bandRows: true, lastRow: true } } };
    const look = tableLook(doc, fake);
    const m = /^(Grid|List|Plain)Table(\d)(Light)?(?:-Accent(\d))?$/.exec(id);
    const name = `${m[1]} Table ${m[2]}${m[3] ? ' Light' : ''}${m[4] ? ` Accent ${m[4]}` : ''}`;
    const b = look.borders ?? {};
    const cond = (type, c) => {
      if (!c || !Object.keys(c).length) return '';
      const rp = rprXml({ ...(c.b ? { b: true } : {}), ...(c.color ? { color: c.color } : {}) });
      const bd = c.borderB || c.borderT ? `<w:tcBorders>${c.borderT ? borderXml('top', c.borderT) : ''}${c.borderB ? borderXml('bottom', c.borderB) : ''}</w:tcBorders>` : '';
      return `<w:tblStylePr w:type="${type}">${rp}${c.shd || bd ? `<w:tcPr>${bd}${c.shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${hex(c.shd)}"/>` : ''}</w:tcPr>` : ''}</w:tblStylePr>`;
    };
    s += `<w:style w:type="table" w:styleId="${esc(id)}"><w:name w:val="${name}"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="49"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => borderXml(k, b[k])).join('')}</w:tblBorders></w:tblPr>${cond('firstRow', look.cell(0, 1))}${cond('lastRow', look.cell(5, 1))}${cond('firstCol', look.cell(2, 0))}${cond('band1Horz', look.cell(1, 1))}</w:style>`;
  }
  return `${s}</w:styles>`;
}
const FMT_OUT = { decimal: 'decimal', upperRoman: 'upperRoman', lowerRoman: 'lowerRoman', upperLetter: 'upperLetter', lowerLetter: 'lowerLetter', bullet: 'bullet', ganada: 'ganada', chosung: 'chosung', decimalCircle: 'decimalEnclosedCircle', koreanDigital: 'koreanDigital', koreanCounting: 'koreanCounting', ideographDigital: 'ideographDigital', decimalZero: 'decimalZero', none: 'none' };
function numberingXml(doc, numIds) {
  let a = ''; let n = '';
  for (const [k, def] of Object.entries(doc.lists ?? {})) {
    const id = numIds[k];
    a += `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>`;
    (def.levels ?? []).slice(0, 9).forEach((L, i) => {
      const bullet = L.fmt === 'bullet';
      a += `<w:lvl w:ilvl="${i}"><w:start w:val="${def.override?.[i]?.start ?? L.start ?? 1}"/><w:numFmt w:val="${FMT_OUT[L.fmt] ?? 'decimal'}"/>${L.legal ? '<w:isLgl/>' : ''}${L.suff && L.suff !== 'tab' ? `<w:suff w:val="${L.suff}"/>` : ''}<w:lvlText w:val="${esc(bullet ? L.text || '•' : L.text ?? `%${i + 1}.`)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${T(L.ind?.l ?? 24 * (i + 1))}" w:hanging="${T(L.ind?.hanging ?? 24)}"/></w:pPr>${bullet ? `<w:rPr><w:rFonts w:ascii="${esc(L.font ?? '맑은 고딕')}" w:hAnsi="${esc(L.font ?? '맑은 고딕')}" w:eastAsia="${esc(L.font ?? '맑은 고딕')}"/></w:rPr>` : L.rPr ? rprXml(L.rPr) : ''}</w:lvl>`;
    });
    a += '</w:abstractNum>';
    n += `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`;
  }
  return `${XML_HEAD}<w:numbering ${NS}>${a}${n}</w:numbering>`;
}
function settingsXml(doc) {
  const s = doc.settings ?? {};
  return `${XML_HEAD}<w:settings ${NS}><w:zoom w:percent="100"/>${s.protect ? `<w:documentProtection w:edit="${s.protect.mode ?? 'readOnly'}" w:enforcement="1"/>` : ''}${s.track ? '<w:trackRevisions/>' : ''}<w:defaultTabStop w:val="${T(s.defaultTab ?? 48)}"/>${s.hyphenate ? '<w:autoHyphenation/>' : ''}${s.evenOdd ? '<w:evenAndOddHeaders/>' : ''}<w:characterSpacingControl w:val="doNotCompress"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="en-US" w:eastAsia="ko-KR"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/></w:settings>`;
}
function themeXml(doc) {
  const f = doc.themeFonts ?? { major: { latin: '맑은 고딕', ea: '맑은 고딕' }, minor: { latin: '맑은 고딕', ea: '맑은 고딕' } };
  const font = (x) => `<a:latin typeface="${esc(x.latin)}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="${esc(x.ea)}"/>`;
  const acc = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  return `${XML_HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 테마"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>${acc.map((c, i) => `<a:accent${i + 1}><a:srgbClr val="${c}"/></a:accent${i + 1}>`).join('')}<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont>${font(f.major)}</a:majorFont><a:minorFont>${font(f.minor)}</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

