// .pptx 쓰기 · 읽기 (PresentationML) — DOM 없음
// 쓰기: pptxEntries(pres) → { 경로: 문자열|바이트 } → zip / zipAsync
// 읽기: readPptx(bytes) → pres (마스터 · 레이아웃 상속, 테마 색, 표, 차트, SmartArt 그림, 그룹, 전환, 애니메이션, 메모)
import { zip, unzip, textOf } from './zip.js';
import { parseXml, child, kids, descendants, esc } from './xml.js';
import { SLOT_NAMES, resolveColor, applyMods, cloneTheme, THEMES } from './themes.js';
import { EMU_PER_PX, uid, layoutPlaceholders, defaultSize, defaultColor, defaultFont, isEmptyText, animSteps, stepTimeline, textBody, para, newSlide } from './model.js';
import { paraDefaults, tableCellStyle, slideBackground, themeDecor } from './render.js';
import { KNOWN_SHAPES } from './shapes.js';
import { decodeEmbeddedFont } from './fonts.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL = (t) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;
const CT = (t) => `application/vnd.openxmlformats-officedocument.presentationml.${t}+xml`;
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
/** WIPOINT 확장(엑셀의 WIXEL 확장과 같은 방식): PowerPoint 는 모르는 ext 를 무시함 */
export const WP_EXT_URI = '{6F1B2D57-7A1C-4E7B-9C1D-57A1F0E1C0DE}';
const WP_NS = 'https://wipoint.app/x';

const E = (px) => Math.round((Number(px) || 0) * EMU_PER_PX);
const PX = (emu) => (Number(emu) || 0) / EMU_PER_PX;
const pct = (v) => Math.round(v * 100000);

// ───────────── 쓰기 도우미 ─────────────
const SLOT_TO_SCHEME = { dk1: 'tx1', lt1: 'bg1', dk2: 'tx2', lt2: 'bg2' };
const MOD_TAG = { lumMod: 'lumMod', lumOff: 'lumOff', tint: 'tint', shade: 'shade', alpha: 'alpha' };
const MOD_CODES = { lm: 'lumMod', lo: 'lumOff', t: 'tint', s: 'shade', a: 'alpha' };

function colorXml(c) {
  if (!c) return '';
  const [base, ...rest] = String(c).split(':');
  const mods = rest.map((m) => { const x = /^([a-z]+)(-?[\d.]+)$/.exec(m); return x && MOD_CODES[x[1]] ? `<a:${MOD_TAG[MOD_CODES[x[1]]]} val="${Math.round(Number(x[2]) * 1000)}"/>` : ''; }).join('');
  if (base.startsWith('@')) {
    const slot = base.slice(1);
    const val = SLOT_TO_SCHEME[slot] ?? ({ tx1: 'tx1', bg1: 'bg1', tx2: 'tx2', bg2: 'bg2' }[slot] ?? slot);
    return mods ? `<a:schemeClr val="${val}">${mods}</a:schemeClr>` : `<a:schemeClr val="${val}"/>`;
  }
  const rgba = /^rgba?\((\d+),(\d+),(\d+)(?:,([\d.]+))?\)$/.exec(base.replace(/\s/g, ''));
  if (rgba) {
    const hex = [rgba[1], rgba[2], rgba[3]].map((v) => Number(v).toString(16).padStart(2, '0')).join('').toUpperCase();
    const al = rgba[4] != null ? `<a:alpha val="${Math.round(Number(rgba[4]) * 100000)}"/>` : '';
    return al || mods ? `<a:srgbClr val="${hex}">${al}${mods}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
  }
  const hex = base.replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6);
  return mods ? `<a:srgbClr val="${hex}">${mods}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}
const solidFill = (c) => (c ? `<a:solidFill>${colorXml(c)}</a:solidFill>` : '');

function fillXml(fill, ctx) {
  if (!fill) return '<a:noFill/>';
  if (fill.type === 'solid') return solidFill(fill.color);
  if (fill.type === 'gradient') {
    const gs = (fill.stops ?? []).map(([p, c]) => `<a:gs pos="${pct(p)}">${colorXml(c)}</a:gs>`).join('');
    const shade = fill.path ? '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' : `<a:lin ang="${Math.round((((fill.angle ?? 90) % 360) + 360) % 360 * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
  }
  if (fill.type === 'image') {
    const rid = ctx.image(fill.media);
    return rid ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"/>${fill.tile ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>` : '<a:noFill/>';
  }
  return '<a:noFill/>';
}

function lineXml(line) {
  if (!line) return '<a:ln><a:noFill/></a:ln>';
  const w = Math.round((line.width ?? 1) * EMU_PER_PX);
  const dash = line.dash && line.dash !== 'solid' ? `<a:prstDash val="${line.dash}"/>` : '';
  const end = (tag, k) => (line[k] && line[k] !== 'none' ? `<a:${tag} type="${line[k]}"/>` : '');
  return `<a:ln w="${w}">${solidFill(line.color)}${dash}<a:round/>${end('headEnd', 'head')}${end('tailEnd', 'tail')}</a:ln>`;
}

function xfrmXml(o, tag = 'a:xfrm') {
  const rot = o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : '';
  const fl = `${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}`;
  return `<${tag}${rot}${fl}><a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${E(o.w)}" cy="${E(o.h)}"/></${tag}>`;
}

function geomXml(o) {
  if (o.path) {
    const paths = o.path.map((p) => `<a:path w="${Math.round(p.w)}" h="${Math.round(p.h)}"${p.fill === 'none' ? ' fill="none"' : ''}>${p.cmds.map(([op, ...v]) => {
      const pt = (x, y) => `<a:pt x="${Math.round(x)}" y="${Math.round(y)}"/>`;
      if (op === 'M') return `<a:moveTo>${pt(v[0], v[1])}</a:moveTo>`;
      if (op === 'L') return `<a:lnTo>${pt(v[0], v[1])}</a:lnTo>`;
      if (op === 'C') return `<a:cubicBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}${pt(v[4], v[5])}</a:cubicBezTo>`;
      if (op === 'Q') return `<a:quadBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}</a:quadBezTo>`;
      if (op === 'A') return `<a:arcTo wR="${Math.round(v[0])}" hR="${Math.round(v[1])}" stAng="${Math.round(v[2])}" swAng="${Math.round(v[3])}"/>`;
      return '<a:close/>';
    }).join('')}</a:path>`).join('');
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
  }
  const av = o.adj ? Object.entries(o.adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join('') : '';
  return `<a:prstGeom prst="${o.shape ?? 'rect'}"><a:avLst>${av}</a:avLst></a:prstGeom>`;
}

function fontXml(font) {
  if (!font) return '';
  if (font === '+mj') return '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/>';
  if (font === '+mn') return '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>';
  return `<a:latin typeface="${esc(font)}"/><a:ea typeface="${esc(font)}"/>`;
}

function rPrXml(r, o, p, ctx, tag = 'a:rPr', fallbackSize) {
  const size = r.size ?? fallbackSize ?? defaultSize(o, p.lvl ?? 0);
  const bold = r.b ?? (o.phKind?.head ? true : undefined);
  const attrs = [
    'lang="ko-KR"', 'altLang="en-US"', `sz="${Math.round(size * 100)}"`,
    bold != null ? `b="${bold ? 1 : 0}"` : '', r.i ? 'i="1"' : '', r.u ? 'u="sng"' : '', r.s ? 'strike="sngStrike"' : '',
    r.base ? `baseline="${r.base > 0 ? 30000 : -25000}"` : '', r.spc ? `spc="${Math.round(r.spc * 100)}"` : '',
    r.cap === 'all' ? 'cap="all"' : r.cap === 'small' ? 'cap="small"' : '', 'dirty="0"',
  ].filter(Boolean).join(' ');
  const ln = r.outline ? `<a:ln w="${Math.round((r.outline.width ?? 1) * EMU_PER_PX)}">${solidFill(r.outline.color ?? '@tx1')}</a:ln>` : '';
  const fill = solidFill(r.color ?? defaultColor(o));
  const eff = r.shadow ? '<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:prstClr val="black"><a:alpha val="43137"/></a:prstClr></a:outerShdw></a:effectLst>' : '';
  const hl = r.hl ? `<a:highlight>${colorXml(r.hl)}</a:highlight>` : '';
  const link = r.link && ctx ? `<a:hlinkClick r:id="${ctx.link(r.link)}"/>` : '';
  return `<${tag} ${attrs}>${ln}${fill}${eff}${hl}${fontXml(r.font ?? defaultFont(o))}${link}</${tag}>`;
}

function pPrXml(o, p) {
  const d = paraDefaults(o, p);
  const algn = { l: 'l', ctr: 'ctr', r: 'r', just: 'just', dist: 'dist' }[p.align ?? 'l'] ?? 'l';
  const attrs = `marL="${E(d.marL)}" indent="${E(d.indent)}"${p.lvl ? ` lvl="${p.lvl}"` : ''} algn="${algn}"`;
  const lnSpc = d.lineSpacingPt ? `<a:lnSpc><a:spcPts val="${Math.round(d.lineSpacingPt * 100)}"/></a:lnSpc>` : `<a:lnSpc><a:spcPct val="${pct(d.lineSpacing)}"/></a:lnSpc>`;
  const bef = `<a:spcBef><a:spcPts val="${Math.round(d.spcBef * 100)}"/></a:spcBef>`;
  const aft = `<a:spcAft><a:spcPts val="${Math.round(d.spcAft * 100)}"/></a:spcAft>`;
  let bu = '<a:buNone/>';
  const b = d.bullet;
  if (b?.type === 'char') bu = `${b.color ? `<a:buClr>${colorXml(b.color)}</a:buClr>` : ''}${b.size ? `<a:buSzPct val="${pct(b.size)}"/>` : ''}<a:buFont typeface="${esc(b.font ?? 'Arial')}" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="${esc(b.char ?? '•')}"/>`;
  else if (b?.type === 'num') bu = `${b.color ? `<a:buClr>${colorXml(b.color)}</a:buClr>` : ''}<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${b.scheme === 'ganada' ? 'ea1JpnKorPeriod' : (b.scheme ?? 'arabicPeriod')}"${b.start && b.start !== 1 ? ` startAt="${b.start}"` : ''}/>`;
  return `<a:pPr ${attrs}>${lnSpc}${bef}${aft}${bu}</a:pPr>`;
}

function txBodyXml(o, body, ctx, tag = 'p:txBody') {
  const ins = body.insets ?? [9.6, 4.8, 9.6, 4.8];
  const anchor = { t: 't', ctr: 'ctr', b: 'b' }[body.anchor ?? 't'];
  const fit = body.autofit === 'resize' ? '<a:spAutoFit/>' : body.autofit === 'shrink' ? `<a:normAutofit${body.fontScale && body.fontScale < 1 ? ` fontScale="${pct(body.fontScale)}"` : ''}${body.lnSpcReduction ? ` lnSpcReduction="${pct(body.lnSpcReduction)}"` : ''}/>` : '<a:noAutofit/>';
  const bodyPr = `<a:bodyPr vert="${body.vert ?? 'horz'}" wrap="${body.wrap === false ? 'none' : 'square'}" lIns="${E(ins[0])}" tIns="${E(ins[1])}" rIns="${E(ins[2])}" bIns="${E(ins[3])}" rtlCol="0" anchor="${anchor}" anchorCtr="0">${fit}</a:bodyPr>`;
  const paras = body.paras.map((p) => {
    const runs = p.runs.map((r) => {
      const parts = String(r.t).split(/\v|\u000b/);
      return parts.map((t, i) => `${i ? `<a:br>${rPrXml(r, o, p, null)}</a:br>` : ''}${t ? `<a:r>${rPrXml(r, o, p, ctx)}<a:t>${esc(t)}</a:t></a:r>` : ''}`).join('');
    }).join('');
    const end = { ...(p.runs[p.runs.length - 1] ?? {}), ...(p.end ?? {}) };
    delete end.t; delete end.link;
    return `<a:p>${pPrXml(o, p)}${runs}${rPrXml(end, o, p, null, 'a:endParaRPr')}</a:p>`;
  }).join('');
  return `<${tag}>${bodyPr}<a:lstStyle/>${paras || '<a:p><a:endParaRPr lang="ko-KR"/></a:p>'}</${tag}>`;
}

const PH_TYPE = { ctrTitle: 'ctrTitle', title: 'title', subTitle: 'subTitle', body: 'body', obj: 'obj', pic: 'pic', dt: 'dt', ftr: 'ftr', sldNum: 'sldNum', tbl: 'tbl', chart: 'chart' };

function phXml(o, phIdx) {
  if (!o.ph || !phIdx) return '<p:nvPr/>';
  const info = phIdx.get(o.id);
  if (!info) return '<p:nvPr/>';
  const plainBody = info.type === 'body' && !o.phKind?.sub && !o.phKind?.head && !o.phKind?.caption;
  const typeAttr = plainBody ? '' : ` type="${info.type}"`;
  return `<p:nvPr><p:ph${typeAttr}${info.idx ? ` idx="${info.idx}"` : ''}/></p:nvPr>`;
}

function shapeXml(o, id, ctx, phIdx) {
  const name = esc(o.name ?? (o.txBox ? `TextBox ${id}` : o.ph ? `Placeholder ${id}` : `Shape ${id}`));
  const link = o.link ? `<a:hlinkClick r:id="${ctx.link(o.link)}"/>` : '';
  const isLine = o.shape === 'line' || o.shape === 'straightConnector1';
  if (isLine && !o.text) {
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}">${link}</p:cNvPr><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrmXml(o)}${geomXml(o)}${lineXml(o.line)}</p:spPr></p:cxnSp>`;
  }
  const effect = o.shadow ? '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>' : '';
  const spPr = `<p:spPr>${xfrmXml(o)}${geomXml(o)}${o.ph ? (o.fill ? fillXml(o.fill, ctx) : '') : fillXml(o.fill, ctx)}${o.ph && !o.line ? '' : lineXml(o.line)}${effect}</p:spPr>`;
  const text = o.text ? txBodyXml(o, o.text, ctx) : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"${o.alt ? ` descr="${esc(o.alt)}"` : ''}>${link}</p:cNvPr><p:cNvSpPr${o.txBox ? ' txBox="1"' : ''}>${o.ph ? '<a:spLocks noGrp="1"/>' : ''}</p:cNvSpPr>${phXml(o, phIdx)}</p:nvSpPr>${spPr}${text}</p:sp>`;
}

function picXml(o, id, ctx, phIdx) {
  if (!o.media) return o.ph ? `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>${phXml(o, phIdx)}</p:nvSpPr><p:spPr>${xfrmXml(o)}</p:spPr></p:sp>` : '';
  const rid = ctx.image(o.media);
  if (!rid) return '';
  const c = o.crop;
  const src = c && (c.l || c.t || c.r || c.b) ? `<a:srcRect l="${pct(c.l)}" t="${pct(c.t)}" r="${pct(c.r)}" b="${pct(c.b)}"/>` : '';
  const alpha = o.alpha != null ? `<a:alphaModFix amt="${pct(o.alpha)}"/>` : '';
  const gray = o.gray ? '<a:grayscl/>' : '';
  const lum = o.bright || o.contrast ? `<a:lum bright="${pct(o.bright ?? 0)}" contrast="${pct(o.contrast ?? 0)}"/>` : '';
  const link = o.link ? `<a:hlinkClick r:id="${ctx.link(o.link)}"/>` : '';
  const effect = o.shadow ? '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>' : '';
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(o.name ?? `Picture ${id}`)}" descr="${esc(o.alt ?? '')}">${link}</p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>${o.ph ? phXml(o, phIdx) : '<p:nvPr/>'}</p:nvPicPr><p:blipFill><a:blip r:embed="${rid}">${alpha}${gray}${lum}</a:blip>${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(o)}${geomXml({ shape: o.shape ?? 'rect', adj: o.adj })}${o.line ? lineXml(o.line) : ''}${effect}</p:spPr></p:pic>`;
}

function tableXml(o, id, ctx, theme) {
  const own = [];
  const rows = o.rows.map((r, ri) => `<a:tr h="${E(r.h)}">${r.cells.map((c, ci) => {
    const ts = tableCellStyle(theme, o, ri, ci);
    if (c.fill !== undefined) own.push([ri, ci]);
    const fill = c.fill !== undefined ? fillXml(c.fill, ctx) : ts.fill ? solidFill(ts.fill) : '<a:noFill/>';
    const pseudo = { ...o, ph: null, phKind: null };
    const body = { ...c.text, paras: c.text.paras.map((p) => ({ ...p, runs: p.runs.map((x) => ({ ...x, color: x.color ?? ts.color, b: x.b ?? (ts.bold || undefined) })), end: { ...(p.end ?? {}), color: p.end?.color ?? ts.color } })) };
    const ins = c.text.insets ?? [7, 3.5, 7, 3.5];
    const anchor = { t: 't', ctr: 'ctr', b: 'b' }[c.text.anchor ?? 't'];
    const bl = (tag, side) => {
      const b = c.borders?.[side];
      if (b) return `<a:${tag} w="${Math.round((b.width ?? 1) * EMU_PER_PX)}">${solidFill(b.color)}</a:${tag}>`;
      return `<a:${tag} w="12700">${solidFill(ts.border)}</a:${tag}>`;
    };
    const merge = `${c.span ? ` gridSpan="${c.span}"` : ''}${c.rowSpan ? ` rowSpan="${c.rowSpan}"` : ''}${c.hMerge ? ' hMerge="1"' : ''}${c.vMerge ? ' vMerge="1"' : ''}`;
    return `<a:tc${merge}>${txBodyXml(pseudo, body, ctx, 'a:txBody').replace(/<a:bodyPr[^>]*>(<a:[a-zA-Z]+[^>]*\/>)?<\/a:bodyPr>/, '<a:bodyPr/>')}<a:tcPr marL="${E(ins[0])}" marR="${E(ins[2])}" marT="${E(ins[1])}" marB="${E(ins[3])}" anchor="${anchor}">${bl('lnL', 'l')}${bl('lnR', 'r')}${bl('lnT', 't')}${bl('lnB', 'b')}${fill}</a:tcPr></a:tc>`;
  }).join('')}</a:tr>`).join('');
  const st = o.style ?? {};
  const ext = `<a:extLst><a:ext uri="${WP_EXT_URI}"><wp:table xmlns:wp="${WP_NS}" json="${esc(JSON.stringify({ style: st, own }))}"/></a:ext></a:extLst>`;
  const tblPr = `<a:tblPr${st.firstRow ? ' firstRow="1"' : ''}${st.banded ? ' bandRow="1"' : ''}${st.lastRow ? ' lastRow="1"' : ''}${st.firstCol ? ' firstCol="1"' : ''}>${ext}</a:tblPr>`;
  const grid = o.cols.map((w) => `<a:gridCol w="${E(w)}"/>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name ?? `Table ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(o, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}<a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function chartFrameXml(o, id, ctx) {
  const rid = ctx.chart(o.chart);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name ?? `Chart ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(o, 'p:xfrm')}<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

/** 차트 파트 (값은 리터럴로 — 내장 통합 문서 없이도 PowerPoint 가 그림) */
export function chartXml(ch) {
  const kind = ch.kind ?? 'col';
  const lit = (vals, num) => `<c:${num ? 'numLit' : 'strLit'}>${num ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(num ? (Number(v) || 0) : v)}</c:v></c:pt>`).join('')}</c:${num ? 'numLit' : 'strLit'}>`;
  const pie = kind === 'pie' || kind === 'doughnut';
  const sers = ch.series.map((s, i) => {
    const color = s.color ? `<c:spPr>${kind === 'line' || kind === 'scatter' ? `<a:ln w="28575" cap="rnd">${solidFill(s.color)}<a:round/></a:ln>` : solidFill(s.color)}</c:spPr>` : '';
    const tx = `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
    if (kind === 'scatter') {
      const xs = ch.cats.map((c, k) => (Number.isFinite(Number(c)) ? Number(c) : k + 1));
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${color}<c:xVal>${lit(xs, true)}</c:xVal><c:yVal>${lit(s.vals, true)}</c:yVal><c:smooth val="0"/></c:ser>`;
    }
    const extra = kind === 'col' || kind === 'bar' ? '<c:invertIfNegative val="0"/>' : kind === 'line' ? `<c:marker><c:symbol val="${ch.markers ? 'circle' : 'none'}"/></c:marker>` : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${color}${extra}<c:cat>${lit(ch.cats, false)}</c:cat><c:val>${lit(s.vals, true)}</c:val>${kind === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  }).join('');
  const dl = ch.labels ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${pie ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pie ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>` : '';
  const axIds = '<c:axId val="111111111"/><c:axId val="222222222"/>';
  let plot;
  if (kind === 'col' || kind === 'bar') plot = `<c:barChart><c:barDir val="${kind}"/><c:grouping val="${ch.stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${sers}${dl}<c:gapWidth val="${ch.stacked ? 150 : 219}"/><c:overlap val="${ch.stacked ? 100 : -27}"/>${axIds}</c:barChart>`;
  else if (kind === 'line') plot = `<c:lineChart><c:grouping val="${ch.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${sers}${dl}<c:marker val="1"/>${axIds}</c:lineChart>`;
  else if (kind === 'area') plot = `<c:areaChart><c:grouping val="${ch.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${sers}${dl}${axIds}</c:areaChart>`;
  else if (kind === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${sers}${dl}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (kind === 'doughnut') plot = `<c:doughnutChart><c:varyColors val="1"/>${sers}${dl}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
  else plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${dl}${axIds}</c:scatterChart>`;
  const ax = pie ? '' : `<c:${kind === 'scatter' ? 'valAx' : 'catAx'}><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${kind === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/>${kind === 'scatter' ? '<c:crossBetween val="midCat"/>' : '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>'}</c:${kind === 'scatter' ? 'valAx' : 'catAx'}><c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${kind === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525">${solidFill('@tx1:lm15:lo85')}</a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1862" b="0"/></a:pPr><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>${esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.legend ? `<c:legend><c:legendPos val="${ch.legend}"/><c:overlay val="0"/></c:legend>` : '';
  const wx = `<c:extLst><c:ext uri="${WP_EXT_URI}"><wp:chart xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(ch))}"/></c:ext></c:extLst>`;
  return `${XML_HEAD}<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${ax}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1197">${solidFill('@tx1:lm65:lo35')}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>${wx}</c:chartSpace>`;
}

// ───────────── 애니메이션 → p:timing ─────────────
const ENTR_PRESET = { appear: 1, fly: 2, split: 16, wheel: 21, wipe: 22, zoom: 53, fade: 10, float: 42, grow: 31, bounce: 26 };
const EMPH_PRESET = { pulse: 26, spin: 8, growShrink: 6, teeter: 32, flash: 35 };
const EXIT_PRESET = { disappear: 1, fly: 2, wipe: 22, zoom: 53, fade: 10 };
const DIR_SUB = { b: 4, t: 1, l: 8, r: 2 };

function timingXml(slide, spidOf) {
  const steps = animSteps(slide);
  if (!steps.length) return '';
  let n = 3;
  const nid = () => n++;
  const tgt = (spid) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
  const set = (spid, val, delay = 0) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
  const effect = (spid, dur, filter, dir = 'in') => `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
  const anim = (spid, dur, attr, from, to) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const flyFrom = { b: ['ppt_y', '1+#ppt_h/2'], t: ['ppt_y', '0-#ppt_h/2'], l: ['ppt_x', '0-#ppt_w/2'], r: ['ppt_x', '1+#ppt_w/2'] };
  const WIPE = { b: 'up', t: 'down', l: 'right', r: 'left' };
  const behaviors = (a, spid) => {
    const dur = Math.max(1, Math.round((a.dur ?? 0.5) * 1000));
    const dir = a.dir ?? 'b';
    if (a.cls === 'entr') {
      switch (a.effect) {
        case 'appear': return set(spid, 'visible');
        case 'fly': case 'bounce': { const [attr, from] = flyFrom[dir] ?? flyFrom.b; const other = attr === 'ppt_y' ? 'ppt_x' : 'ppt_y'; return set(spid, 'visible') + anim(spid, dur, attr, from, `#${attr}`) + anim(spid, dur, other, `#${other}`, `#${other}`); }
        case 'wipe': return set(spid, 'visible') + effect(spid, dur, `wipe(${WIPE[dir] ?? 'up'})`);
        case 'split': return set(spid, 'visible') + effect(spid, dur, 'barn(inVertical)');
        case 'wheel': return set(spid, 'visible') + effect(spid, dur, 'wheel(1)');
        case 'zoom': return set(spid, 'visible') + anim(spid, dur, 'ppt_w', '0', '#ppt_w') + anim(spid, dur, 'ppt_h', '0', '#ppt_h') + effect(spid, dur, 'fade');
        case 'grow': return set(spid, 'visible') + anim(spid, dur, 'ppt_w', '0', '#ppt_w') + anim(spid, dur, 'ppt_h', '0', '#ppt_h') + anim(spid, dur, 'style.rotation', '90', '0') + effect(spid, dur, 'fade');
        case 'float': return set(spid, 'visible') + effect(spid, dur, 'fade') + anim(spid, dur, 'ppt_y', '#ppt_y+.1', '#ppt_y');
        default: return set(spid, 'visible') + effect(spid, dur, 'fade');
      }
    }
    if (a.cls === 'exit') {
      const hide = set(spid, 'hidden', Math.max(0, dur - 1));
      switch (a.effect) {
        case 'disappear': return set(spid, 'hidden');
        case 'fly': { const [attr, to] = flyFrom[dir] ?? flyFrom.b; return anim(spid, dur, attr, `#${attr}`, to) + hide; }
        case 'wipe': return effect(spid, dur, `wipe(${WIPE[dir] ?? 'down'})`, 'out') + hide;
        case 'zoom': return anim(spid, dur, 'ppt_w', '#ppt_w', '0') + anim(spid, dur, 'ppt_h', '#ppt_h', '0') + effect(spid, dur, 'fade', 'out') + hide;
        default: return effect(spid, dur, 'fade', 'out') + hide;
      }
    }
    // 강조
    if (a.effect === 'spin' || a.effect === 'teeter') return `<p:animRot by="${a.effect === 'spin' ? 21600000 : 300000}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${a.effect === 'teeter' ? ' autoRev="1" repeatCount="2000"' : ''}/>${tgt(spid)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    if (a.effect === 'flash') return `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" autoRev="1"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
    const by = a.effect === 'growShrink' ? 150000 : 110000;
    return `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${a.effect === 'pulse' ? ' autoRev="1"' : ''}/>${tgt(spid)}</p:cBhvr><p:by x="${by}" y="${by}"/></p:animScale>`;
  };
  const clickPars = steps.map((step) => {
    const outerId = nid();
    const tl = stepTimeline(step);
    // 같은 시작 시각끼리 묶음 (PowerPoint 구조: 클릭 → 시간 그룹 → 효과)
    const groups = [];
    for (const t of tl) {
      const at = Math.round((t.at - (t.anim.delay ?? 0)) * 1000);
      let g = groups.find((x) => x.at === at);
      if (!g) { g = { at, items: [] }; groups.push(g); }
      g.items.push(t);
    }
    const inner = groups.map((g) => {
      const gid = nid();
      const effs = g.items.map((t, k) => {
        const a = t.anim;
        const spid = spidOf(a.obj);
        if (!spid) return '';
        const table = a.cls === 'entr' ? ENTR_PRESET : a.cls === 'exit' ? EXIT_PRESET : EMPH_PRESET;
        const node = k === 0 && g === groups[0] && step[0] === a ? (a.start === 'after' ? 'afterEffect' : a.start === 'with' ? 'withEffect' : 'clickEffect') : a.start === 'after' ? 'afterEffect' : 'withEffect';
        const sub = a.cls === 'emph' ? 0 : a.effect === 'fly' || a.effect === 'wipe' ? DIR_SUB[a.dir ?? 'b'] ?? 4 : a.effect === 'zoom' ? 16 : a.effect === 'split' ? 37 : 0;
        const eid = nid();
        return `<p:par><p:cTn id="${eid}" presetID="${table[a.effect] ?? 10}" presetClass="${a.cls}" presetSubtype="${sub}" fill="hold" grpId="0" nodeType="${node}"><p:stCondLst><p:cond delay="${Math.round((a.delay ?? 0) * 1000)}"/></p:stCondLst><p:childTnLst>${behaviors(a, spid)}</p:childTnLst></p:cTn></p:par>`;
      }).join('');
      return `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${g.at}"/></p:stCondLst><p:childTnLst>${effs}</p:childTnLst></p:cTn></p:par>`;
    }).join('');
    const first = step[0];
    const cond = first.start === 'click' ? '<p:cond delay="indefinite"/>' : '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>';
    return `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
  const blds = [...new Set(steps.flat().map((a) => spidOf(a.obj)).filter(Boolean))].map((spid) => `<p:bldP spid="${spid}" grpId="0"/>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clickPars}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${blds ? `<p:bldLst>${blds}</p:bldLst>` : ''}</p:timing>`;
}

const TRANS_XML = {
  cut: () => '<p:cut/>', fade: () => '<p:fade/>', dissolve: () => '<p:dissolve/>', zoom: () => '<p:zoom/>', circle: () => '<p:circle/>',
  push: (d) => `<p:push dir="${{ b: 'u', t: 'd', l: 'r', r: 'l' }[d] ?? 'u'}"/>`, wipe: (d) => `<p:wipe dir="${{ b: 'u', t: 'd', l: 'r', r: 'l' }[d] ?? 'r'}"/>`,
  cover: (d) => `<p:cover dir="${{ b: 'u', t: 'd', l: 'r', r: 'l' }[d] ?? 'l'}"/>`, uncover: (d) => `<p:pull dir="${{ b: 'u', t: 'd', l: 'r', r: 'l' }[d] ?? 'l'}"/>`,
  split: () => '<p:split orient="vert" dir="out"/>', flip: () => '<p:fade/>',
};
function transitionXml(t) {
  if (!t || t.type === 'none') return t?.advAfter != null ? `<p:transition advTm="${Math.round(t.advAfter * 1000)}"${t.advClick === false ? ' advClick="0"' : ''}/>` : '';
  const dur = t.dur ?? 1;
  const spd = dur <= 0.5 ? 'fast' : dur <= 0.75 ? 'med' : 'slow';
  const adv = `${t.advClick === false ? ' advClick="0"' : ''}${t.advAfter != null ? ` advTm="${Math.round(t.advAfter * 1000)}"` : ''}`;
  return `<p:transition spd="${spd}"${adv}>${(TRANS_XML[t.type] ?? TRANS_XML.fade)(t.dir)}</p:transition>`;
}

// ───────────── 테마 · 마스터 · 레이아웃 ─────────────
function themeXml(theme, name = 'WIPOINT') {
  const clr = SLOT_NAMES.map((k) => {
    const v = (theme.colors[k] ?? '#000000').replace('#', '').toUpperCase();
    if (k === 'dk1' && v === '000000') return '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
    if (k === 'lt1' && v === 'FFFFFF') return '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
    return `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`;
  }).join('');
  const font = (f) => `<a:latin typeface="${esc(f)}" panose="020F0302020204030204"/><a:ea typeface="${esc(f)}"/><a:cs typeface=""/><a:font script="Hang" typeface="${esc(f)}"/>`;
  const ph = (inner) => `<a:solidFill><a:schemeClr val="phClr">${inner}</a:schemeClr></a:solidFill>`;
  return `${XML_HEAD}<a:theme xmlns:a="${NS_A}" name="${esc(theme.label ?? name)}"><a:themeElements><a:clrScheme name="${esc(theme.label ?? name)}">${clr}</a:clrScheme><a:fontScheme name="${esc(theme.label ?? name)}"><a:majorFont>${font(theme.fonts.major)}</a:majorFont><a:minorFont>${font(theme.fonts.minor)}</a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst>${ph('')}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">${ph('')}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">${ph('')}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">${ph('')}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${ph('')}${ph('<a:tint val="95000"/><a:satMod val="170000"/>')}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function lvlStyles(sizes, bullets, extra = '') {
  return sizes.map((sz, i) => {
    const marL = bullets ? 228600 + i * 457200 : i * 457200;
    const bu = bullets ? `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/>` : '<a:buNone/>';
    return `<a:lvl${i + 1}pPr marL="${marL}" indent="${bullets ? -228600 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${i ? 500 : 1000}"/></a:spcBef>${bu}<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`;
  }).join('') + extra;
}

function masterXml(pres, layoutRids) {
  const size = pres.size;
  const sx = size.w / 1280;
  const sy = size.h / 720;
  const ph = (id, name, type, x, y, w, h, idx) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x * sx)}" y="${E(y * sy)}"/><a:ext cx="${E(w * sx)}" cy="${E(h * sy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="${type === 'title' ? 'ctr' : 't'}"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>${type === 'title' ? '마스터 제목 스타일 편집' : '마스터 텍스트 스타일을 편집합니다'}</a:t></a:r><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>`;
  const sizes = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800];
  const other = [1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800].map((sz, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('');
  const ids = layoutRids.map((rid, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="${rid}"/>`).join('');
  return `${XML_HEAD}<p:sldMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${ph(2, 'Title Placeholder 1', 'title', 88, 38, 1104, 139)}${ph(3, 'Text Placeholder 2', 'body', 88, 192, 1104, 456, 1)}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>${lvlStyles(sizes, true)}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr>${other}</p:otherStyle></p:txStyles></p:sldMaster>`;
}

const LAYOUT_TYPE = { title: 'title', titleContent: 'obj', section: 'secHead', twoContent: 'twoObj', comparison: 'twoTxTwoObj', titleOnly: 'titleOnly', blank: 'blank', contentCaption: 'objTx', pictureCaption: 'picTx' };
const LAYOUT_NAME = { title: '제목 슬라이드', titleContent: '제목 및 내용', section: '구역 머리글', twoContent: '콘텐츠 2개', comparison: '비교', titleOnly: '제목만', blank: '빈 화면', contentCaption: '캡션 있는 콘텐츠', pictureCaption: '캡션 있는 그림' };

/** 슬라이드 개체 틀 → (종류, idx) 배정. 레이아웃과 슬라이드가 같은 규칙을 씀 */
function phAssign(objs) {
  const map = new Map();
  const count = {};
  for (const o of objs) {
    if (!o.ph) continue;
    const type = o.ph === 'obj' ? 'body' : o.ph;
    count[type] = (count[type] ?? 0) + 1;
    const key = `${type}#${count[type]}`;
    const idx = type === 'title' || type === 'ctrTitle' ? 0 : ({ body: 0, subTitle: 100, pic: 20, tbl: 30, chart: 40 }[type] ?? 50) + count[type];
    map.set(o.id, { type, idx, key });
  }
  return map;
}

function layoutXml(pres, def, ctx) {
  const objs = [];
  let id = 2;
  const parts = [];
  for (const o of def.decor) {
    const c = { ...o };
    parts.push(objXml(c, id++, ctx, null, pres.theme));
  }
  const phIdx = phAssign(def.phs);
  for (const o of def.phs) {
    const c = { ...o, text: o.text ? { ...o.text, paras: [{ ...(o.text.paras[0] ?? para()), runs: [] }] } : null };
    if (c.type === 'image') parts.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>${phXml(c, phIdx)}</p:nvSpPr><p:spPr>${xfrmXml(c)}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>`);
    else parts.push(shapeXml(c, id, ctx, phIdx));
    id++;
    objs.push(c);
  }
  const W = pres.size.w;
  const H = pres.size.h;
  for (const [type, idx, x, w] of [['dt', 10, W * 0.069, W * 0.225], ['ftr', 11, W * 0.33, W * 0.34], ['sldNum', 12, W * 0.706, W * 0.225]]) {
    parts.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${type} ${idx}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}" sz="quarter" idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x)}" y="${E(H - 54)}"/><a:ext cx="${E(w)}" cy="${E(32)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>`);
  }
  const bg = def.bg ? `<p:bg><p:bgPr>${fillXml(def.bg, ctx)}<a:effectLst/></p:bgPr></p:bg>` : '';
  return `${XML_HEAD}<p:sldLayout xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" type="${LAYOUT_TYPE[def.kind] ?? 'cust'}" preserve="1"${def.showMaster === false ? ' showMasterSp="0"' : ''}><p:cSld name="${esc(def.name)}">${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${parts.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function objXml(o, id, ctx, phIdx, theme) {
  if (o.type === 'image') return picXml(o, id, ctx, phIdx);
  if (o.type === 'table') return tableXml(o, id, ctx, theme);
  if (o.type === 'chart') return chartFrameXml(o, id, ctx);
  return shapeXml(o, id, ctx, phIdx);
}

function notesMasterXml() {
  return `${XML_HEAD}<p:notesMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>마스터 텍스트 스타일을 편집합니다</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

function notesSlideXml(text) {
  const paras = String(text).split('\n').map((l) => `<a:p>${l ? `<a:r><a:rPr lang="ko-KR" altLang="en-US" dirty="0"/><a:t>${esc(l)}</a:t></a:r>` : ''}<a:endParaRPr lang="ko-KR" altLang="en-US"/></a:p>`).join('');
  return `${XML_HEAD}<p:notes xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

function dataUrlBytes(url) {
  const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(url ?? '');
  if (!m) return null;
  const mime = m[1];
  let bytes;
  if (m[2]) {
    const bin = atob(m[3]);
    bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  } else bytes = new TextEncoder().encode(decodeURIComponent(m[3]));
  const ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'image/tiff': 'tiff' }[mime] ?? 'png';
  return { bytes, ext, mime };
}
export const MIME_BY_EXT = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' };

/**
 * 문서 → pptx 파트 목록
 * opts.png: { mediaId: dataURL } — SVG 그림의 PNG 대체본 (앱이 미리 그려 넘김, PowerPoint 2016 이전 호환)
 */
export function pptxEntries(pres, opts = {}) {
  const files = {};
  const ct = new Map();
  const defaults = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
  const override = (path, type) => ct.set(`/${path}`, type);
  const mediaPath = new Map();
  let mediaN = 0;
  const putMedia = (id) => {
    if (mediaPath.has(id)) return mediaPath.get(id);
    let url = pres.media?.[id];
    if (!url) return null;
    let d = dataUrlBytes(url);
    if (!d) return null;
    if (d.ext === 'svg' && opts.png?.[id]) d = dataUrlBytes(opts.png[id]) ?? d;
    const path = `ppt/media/image${++mediaN}.${d.ext}`;
    files[path] = d.bytes;
    defaults.set(d.ext, d.mime === 'image/jpg' ? 'image/jpeg' : d.mime);
    mediaPath.set(id, path);
    return path;
  };
  let chartN = 0;
  /** 파트별 관계 목록 */
  const relsFor = () => {
    const list = [];
    const byKey = new Map();
    const add = (type, target, ext = false, key = `${type}|${target}`) => {
      if (byKey.has(key)) return byKey.get(key);
      const rid = `rId${list.length + 1}`;
      list.push(`<Relationship Id="${rid}" Type="${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`);
      byKey.set(key, rid);
      return rid;
    };
    return {
      add,
      image: (id) => { const p = putMedia(id); return p ? add(REL('image'), `../media/${p.split('/').pop()}`) : null; },
      link: (url) => add(REL('hyperlink'), url, true),
      chart: (ch) => {
        const path = `ppt/charts/chart${++chartN}.xml`;
        files[path] = chartXml(ch);
        override(path, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
        return add(REL('chart'), `../charts/chart${chartN}.xml`, false, `chart${chartN}`);
      },
      xml: () => `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.join('')}</Relationships>`,
    };
  };

  // 레이아웃: (종류, 배경 장식) 조합마다 하나 — 슬라이드가 쓰는 것만 + 기본 9종
  const layoutDefs = [];
  const layoutKey = new Map();
  const layoutOf = (slide) => {
    const kind = LAYOUT_TYPE[slide.layout] ? slide.layout : 'blank';
    const decor = slide.bgObjects ?? (slide.hideDecor ? [] : themeDecor(pres.theme, kind, pres.size));
    const key = `${kind}|${JSON.stringify(decor)}`;
    if (layoutKey.has(key)) {
      const def = layoutKey.get(key);
      mergePhs(def, slide);
      return def;
    }
    const def = { kind, decor, name: slide.bgObjects ? `${LAYOUT_NAME[kind] ?? '사용자 지정'} ${layoutDefs.length + 1}` : LAYOUT_NAME[kind] ?? '사용자 지정', phs: slide.bgObjects ? [] : layoutPlaceholders(kind, pres.size), index: layoutDefs.length + 1, showMaster: true };
    mergePhs(def, slide);
    layoutDefs.push(def);
    layoutKey.set(key, def);
    return def;
  };
  // 슬라이드에만 있는 개체 틀 자리도 레이아웃에 만들어 둠 (PowerPoint 가 연결을 찾도록)
  function mergePhs(def, slide) {
    const have = phAssign(def.phs);
    const keys = new Set([...have.values()].map((v) => v.key));
    for (const [oid, info] of phAssign(slide.objects)) {
      if (keys.has(info.key)) continue;
      const o = slide.objects.find((x) => x.id === oid);
      def.phs.push({ ...o, id: uid(), text: o.text ? { ...o.text, paras: [{ ...(o.text.paras[0] ?? para()), runs: [] }] } : null, media: null });
      keys.add(info.key);
    }
  }
  for (const k of Object.keys(LAYOUT_TYPE)) layoutOf({ layout: k, objects: [] });
  const slideLayouts = pres.slides.map((s) => layoutOf(s));

  // 프레젠테이션
  const presRels = relsFor();
  const masterRid = presRels.add(REL('slideMaster'), 'slideMasters/slideMaster1.xml');
  const slideRids = pres.slides.map((_, i) => presRels.add(REL('slide'), `slides/slide${i + 1}.xml`));
  const hasNotes = pres.slides.some((s) => s.notes);
  const notesMasterRid = hasNotes ? presRels.add(REL('notesMaster'), 'notesMasters/notesMaster1.xml') : null;
  presRels.add(REL('presProps'), 'presProps.xml');
  presRels.add(REL('viewProps'), 'viewProps.xml');
  presRels.add(REL('theme'), 'theme/theme1.xml');
  presRels.add(REL('tableStyles'), 'tableStyles.xml');
  // 포함된 글꼴 (읽은 .fntdata 그대로)
  const fontEls = new Map();
  (pres.fonts ?? []).forEach((f, i) => {
    const d = dataUrlBytes(f.data);
    if (!d) return;
    const path = `fonts/font${i + 1}.fntdata`;
    files[`ppt/${path}`] = d.bytes;
    defaults.set('fntdata', 'application/x-fontdata');
    const rid2 = presRels.add(REL('font'), path);
    if (!fontEls.has(f.typeface)) fontEls.set(f.typeface, { f, parts: [] });
    fontEls.get(f.typeface).parts.push(`<p:${f.style} r:id="${rid2}"/>`);
  });
  const order = ['regular', 'bold', 'italic', 'boldItalic'];
  const fontXml = fontEls.size ? `<p:embeddedFontLst>${[...fontEls.values()].map(({ f, parts }) => `<p:embeddedFont><p:font typeface="${esc(f.typeface)}"${f.panose ? ` panose="${esc(f.panose)}"` : ''}${f.pitchFamily ? ` pitchFamily="${esc(f.pitchFamily)}"` : ''}${f.charset ? ` charset="${esc(f.charset)}"` : ''}/>${parts.sort((a, b) => order.indexOf(a.slice(3, a.indexOf(' '))) - order.indexOf(b.slice(3, b.indexOf(' ')))).join('')}</p:embeddedFont>`).join('')}</p:embeddedFontLst>` : '';

  // 마스터 · 레이아웃
  const masterRels = relsFor();
  const layoutRids = layoutDefs.map((d) => masterRels.add(REL('slideLayout'), `../slideLayouts/slideLayout${d.index}.xml`));
  masterRels.add(REL('theme'), '../theme/theme1.xml');
  files['ppt/slideMasters/slideMaster1.xml'] = masterXml(pres, layoutRids);
  files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = masterRels.xml();
  override('ppt/slideMasters/slideMaster1.xml', CT('slideMaster'));
  for (const d of layoutDefs) {
    const rels = relsFor();
    rels.add(REL('slideMaster'), '../slideMasters/slideMaster1.xml');
    files[`ppt/slideLayouts/slideLayout${d.index}.xml`] = layoutXml(pres, d, rels);
    files[`ppt/slideLayouts/_rels/slideLayout${d.index}.xml.rels`] = rels.xml();
    override(`ppt/slideLayouts/slideLayout${d.index}.xml`, CT('slideLayout'));
  }
  files['ppt/theme/theme1.xml'] = themeXml(pres.theme);
  override('ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
  if (hasNotes) {
    files['ppt/notesMasters/notesMaster1.xml'] = notesMasterXml();
    files['ppt/notesMasters/_rels/notesMaster1.xml.rels'] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL('theme')}" Target="../theme/theme2.xml"/></Relationships>`;
    files['ppt/theme/theme2.xml'] = themeXml(cloneTheme(THEMES[0]), 'Notes');
    override('ppt/notesMasters/notesMaster1.xml', CT('notesMaster'));
    override('ppt/theme/theme2.xml', 'application/vnd.openxmlformats-officedocument.theme+xml');
  }

  // 슬라이드
  pres.slides.forEach((slide, si) => {
    const n = si + 1;
    const rels = relsFor();
    rels.add(REL('slideLayout'), `../slideLayouts/slideLayout${slideLayouts[si].index}.xml`);
    const phIdx = phAssign(slide.objects);
    const spid = new Map();
    let id = 2;
    const parts = [];
    // 그룹: 같은 grp 의 개체를 p:grpSp 로 묶음 (처음 나온 위치에)
    const done = new Set();
    for (const o of slide.objects) {
      if (done.has(o.id)) continue;
      if (o.grp) {
        const members = slide.objects.filter((x) => x.grp === o.grp);
        const gid = id++;
        let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
        for (const m of members) { x1 = Math.min(x1, m.x); y1 = Math.min(y1, m.y); x2 = Math.max(x2, m.x + m.w); y2 = Math.max(y2, m.y + m.h); }
        const inner = members.map((m) => { done.add(m.id); const mid = id++; spid.set(m.id, mid); return objXml(m, mid, rels, phIdx, pres.theme); }).join('');
        const ext = `<a:off x="${E(x1)}" y="${E(y1)}"/><a:ext cx="${E(x2 - x1)}" cy="${E(y2 - y1)}"/>`;
        parts.push(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gid}" name="Group ${gid}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${ext}<a:chOff x="${E(x1)}" y="${E(y1)}"/><a:chExt cx="${E(x2 - x1)}" cy="${E(y2 - y1)}"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`);
        for (const m of members) spid.set(`grp:${o.grp}`, gid);
        continue;
      }
      done.add(o.id);
      const oid = id++;
      spid.set(o.id, oid);
      parts.push(objXml(o, oid, rels, phIdx, pres.theme));
    }
    // 바닥글 (슬라이드 번호 · 날짜 · 글)
    const ft = pres.footer ?? {};
    if ((ft.slideNum || ft.date || ft.text) && !(ft.hideOnTitle && slide.layout === 'title')) {
      const W = pres.size.w;
      const H = pres.size.h;
      const mk = (type, x, w, text, algn, fld) => `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${type}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}" sz="quarter" idx="${{ dt: 10, ftr: 11, sldNum: 12 }[type]}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x)}" y="${E(H - 54)}"/><a:ext cx="${E(w)}" cy="${E(32)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${algn}"/>${fld ? `<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="${fld}"><a:rPr lang="ko-KR" altLang="en-US" sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>${esc(text)}</a:t></a:fld>` : `<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>${esc(text)}</a:t></a:r>`}<a:endParaRPr lang="ko-KR" sz="1200"/></a:p></p:txBody></p:sp>`;
      if (ft.date) parts.push(mk('dt', W * 0.069, W * 0.225, ft.dateText || new Date().toLocaleDateString('ko-KR'), 'l', ft.dateText ? null : 'datetime1'));
      if (ft.text) parts.push(mk('ftr', W * 0.33, W * 0.34, ft.text, 'ctr'));
      if (ft.slideNum) parts.push(mk('sldNum', W * 0.706, W * 0.225, String(n), 'r', 'slidenum'));
    }
    const bgFill = slide.bg ? slide.bg : null;
    const bg = bgFill ? `<p:bg><p:bgPr>${fillXml(bgFill, rels)}<a:effectLst/></p:bgPr></p:bg>` : '';
    const meta = { layout: slide.layout, transition: slide.transition ?? null, anims: slide.anims ?? [], hideDecor: slide.hideDecor ?? false, section: slide.section ?? null, bgObjects: slide.bgObjects ? true : undefined, ids: Object.fromEntries([...spid].filter(([k]) => !String(k).startsWith('grp:')).map(([k, v]) => [v, k])) };
    const ext = `<p:extLst><p:ext uri="${WP_EXT_URI}"><wp:slide xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(meta))}"/></p:ext></p:extLst>`;
    const timing = timingXml(slide, (oid) => spid.get(oid));
    files[`ppt/slides/slide${n}.xml`] = `${XML_HEAD}<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"${slide.hidden ? ' show="0"' : ''}><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${parts.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml(slide.transition)}${timing}${ext}</p:sld>`;
    if (slide.notes) {
      const nrels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL('notesMaster')}" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL('slide')}" Target="../slides/slide${n}.xml"/></Relationships>`;
      files[`ppt/notesSlides/notesSlide${n}.xml`] = notesSlideXml(slide.notes);
      files[`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`] = nrels;
      override(`ppt/notesSlides/notesSlide${n}.xml`, CT('notesSlide'));
      rels.add(REL('notesSlide'), `../notesSlides/notesSlide${n}.xml`);
    }
    files[`ppt/slides/_rels/slide${n}.xml.rels`] = rels.xml();
    override(`ppt/slides/slide${n}.xml`, CT('slide'));
  });

  // 구역
  const sections = [];
  pres.slides.forEach((s, i) => { if (s.section || (i === 0 && pres.slides.some((x) => x.section))) sections.push({ name: s.section ?? '기본 구역', ids: [] }); sections[sections.length - 1]?.ids.push(256 + i); });
  const secExt = sections.length ? `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${sections.map((s) => `<p14:section name="${esc(s.name)}" id="{${uuidish()}}"><p14:sldIdLst>${s.ids.map((x) => `<p14:sldId id="${x}"/>`).join('')}</p14:sldIdLst></p14:section>`).join('')}</p14:sectionLst></p:ext>` : '';
  const th = pres.theme;
  const presMeta = { theme: { name: th.name, label: th.label, accentBand: !!th.accentBand, titleBar: !!th.titleBar, dark: !!th.dark, bg: th.bg ?? null }, footer: pres.footer ?? null };
  const secXml = `<p:extLst>${secExt}<p:ext uri="${WP_EXT_URI}"><wp:pres xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(presMeta))}"/></p:ext></p:extLst>`;
  const sz = pres.size;
  const sldIds = slideRids.map((rid, i) => `<p:sldId id="${256 + i}" r:id="${rid}"/>`).join('');
  const defText = `<p:defaultTextStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((l) => `<a:lvl${l}pPr marL="${(l - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${l}pPr>`).join('')}</p:defaultTextStyle>`;
  files['ppt/presentation.xml'] = `${XML_HEAD}<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" saveSubsetFonts="1"${fontXml ? ' embedTrueTypeFonts="1"' : ''}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>${notesMasterRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>` : ''}${sldIds ? `<p:sldIdLst>${sldIds}</p:sldIdLst>` : ''}<p:sldSz cx="${E(sz.w)}" cy="${E(sz.h)}"/><p:notesSz cx="6858000" cy="9144000"/>${fontXml}${defText}${secXml}</p:presentation>`;
  files['ppt/_rels/presentation.xml.rels'] = presRels.xml();
  override('ppt/presentation.xml', CT('presentation.main'));
  files['ppt/presProps.xml'] = `${XML_HEAD}<p:presentationPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"/>`;
  files['ppt/viewProps.xml'] = `${XML_HEAD}<p:viewPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
  files['ppt/tableStyles.xml'] = `${XML_HEAD}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
  override('ppt/presProps.xml', CT('presProps'));
  override('ppt/viewProps.xml', CT('viewProps'));
  override('ppt/tableStyles.xml', CT('tableStyles'));
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const p = pres.props ?? {};
  files['docProps/core.xml'] = `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(p.title ?? '')}</dc:title><dc:creator>${esc(p.author ?? '')}</dc:creator><cp:lastModifiedBy>${esc(p.author ?? '')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${esc((p.created ?? now).replace(/\.\d+Z$/, 'Z'))}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  files['docProps/app.xml'] = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>WIPOINT</Application><Slides>${pres.slides.length}</Slides><Notes>${pres.slides.filter((s) => s.notes).length}</Notes><HiddenSlides>${pres.slides.filter((s) => s.hidden).length}</HiddenSlides><PresentationFormat>${sz.w / sz.h > 1.5 ? '와이드스크린' : '화면 슬라이드 쇼(4:3)'}</PresentationFormat></Properties>`;
  override('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml');
  override('docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
  files['_rels/.rels'] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="${REL('extended-properties')}" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${REL('officeDocument')}" Target="ppt/presentation.xml"/></Relationships>`;
  const ctXml = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${[...defaults].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('')}${[...ct].map(([p2, t]) => `<Override PartName="${esc(p2)}" ContentType="${t}"/>`).join('')}</Types>`;
  // [Content_Types].xml 이 맨 앞에 오도록
  return { '[Content_Types].xml': ctXml, ...files };
}

function uuidish() {
  const h = () => Math.floor(Math.random() * 0x10000).toString(16).toUpperCase().padStart(4, '0');
  return `${h()}${h()}-${h()}-${h()}-${h()}-${h()}${h()}${h()}`;
}

/** 동기 쓰기 (테스트 · 작은 문서) — 브라우저는 zipAsync(pptxEntries(pres)) 로 압축 */
export function writePptx(pres, opts) {
  return zip(pptxEntries(pres, opts));
}

// ═════════════ 읽기 ═════════════
const sanitizeName = (p) => p.replace(/^\/+/, '');
function resolvePath(base, target) {
  if (target.startsWith('/')) return sanitizeName(target);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}
const relsPath = (part) => { const i = part.lastIndexOf('/'); return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`; };
const attr = (el, name) => el?.attrs?.[name] ?? el?.attrs?.[Object.keys(el?.attrs ?? {}).find((k) => k.endsWith(`:${name}`)) ?? ''];
/** r:id (같은 요소에 id 와 r:id 가 함께 있을 때) */
const rid = (el) => el?.attrs?.['r:id'] ?? el?.attrs?.[Object.keys(el?.attrs ?? {}).find((k) => k.endsWith(':id')) ?? ''];
const num = (v, d = 0) => (v === undefined || v === null || v === '' ? d : Number(v));
const bool = (v) => v === '1' || v === 'true' || v === 'on';

/** 요소 안의 첫 번째 색 → WIPOINT 색 문자열 */
function readColor(el, ctx, phClr = null) {
  if (!el) return null;
  const c = el.children.find((x) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'].includes(x.name));
  if (!c) return null;
  const mods = [];
  for (const m of c.children) {
    if (['lumMod', 'lumOff', 'tint', 'shade', 'alpha'].includes(m.name)) mods.push([m.name, num(attr(m, 'val')) / 100000]);
  }
  const modStr = (list) => list.map(([k, v]) => `:${{ lumMod: 'lm', lumOff: 'lo', tint: 't', shade: 's', alpha: 'a' }[k]}${Math.round(v * 1000) / 10}`).join('');
  if (c.name === 'schemeClr') {
    let val = attr(c, 'val');
    if (val === 'phClr') {
      if (!phClr) return null;
      return mods.length ? `${phClr}${modStr(mods)}` : phClr;
    }
    const mapped = ctx?.clrMap?.[val] ?? { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }[val] ?? val;
    return `@${mapped}${modStr(mods)}`;
  }
  let hex = '#000000';
  if (c.name === 'srgbClr') hex = `#${(attr(c, 'val') ?? '000000').toUpperCase()}`;
  else if (c.name === 'sysClr') hex = `#${(attr(c, 'lastClr') ?? (attr(c, 'val') === 'window' ? 'FFFFFF' : '000000')).toUpperCase()}`;
  else if (c.name === 'prstClr') hex = PRESET_COLORS[attr(c, 'val')] ?? '#000000';
  else if (c.name === 'scrgbClr') hex = `#${['r', 'g', 'b'].map((k) => Math.round(Math.min(1, num(attr(c, k)) / 100000) ** (1 / 2.2) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
  const nonAlpha = mods.filter(([k]) => k !== 'alpha');
  const alpha = mods.find(([k]) => k === 'alpha');
  if (nonAlpha.length) hex = applyMods(hex, nonAlpha).hex;
  return alpha && alpha[1] < 1 ? `${hex}:a${Math.round(alpha[1] * 1000) / 10}` : hex;
}
const PRESET_COLORS = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', orange: '#FFA500', purple: '#800080', navy: '#000080', silver: '#C0C0C0' };

function readFill(spPr, ctx, phClr) {
  if (!spPr) return undefined;
  for (const c of spPr.children) {
    if (c.name === 'noFill') return null;
    if (c.name === 'solidFill') return { type: 'solid', color: readColor(c, ctx, phClr) ?? '#000000' };
    if (c.name === 'gradFill') {
      const stops = descendants(child(c, 'gsLst'), 'gs').map((g) => [num(attr(g, 'pos')) / 100000, readColor(g, ctx, phClr) ?? '#FFFFFF']).sort((a, b) => a[0] - b[0]);
      const lin = child(c, 'lin');
      const path = child(c, 'path');
      return { type: 'gradient', stops, angle: lin ? num(attr(lin, 'ang')) / 60000 : 90, ...(path ? { path: true } : {}) };
    }
    if (c.name === 'blipFill') {
      const blip = child(c, 'blip');
      const media = ctx.media(attr(blip, 'embed'));
      return media ? { type: 'image', media, ...(child(c, 'tile') ? { tile: true } : {}) } : null;
    }
    if (c.name === 'pattFill') return { type: 'solid', color: readColor(child(c, 'fgClr'), ctx, phClr) ?? '#000000' };
    if (c.name === 'grpFill') return ctx.groupFill ?? null;
  }
  return undefined;
}

function readLine(ln, ctx, phClr) {
  if (!ln) return undefined;
  if (child(ln, 'noFill')) return null;
  const fillEl = child(ln, 'solidFill') ?? child(ln, 'gradFill');
  const color = fillEl ? readColor(fillEl.name === 'gradFill' ? descendants(fillEl, 'gs')[0] : fillEl, ctx, phClr) : phClr;
  const out = { color: color ?? '#000000', width: attr(ln, 'w') ? PX(attr(ln, 'w')) : 1 };
  const dash = attr(child(ln, 'prstDash'), 'val');
  if (dash && dash !== 'solid') out.dash = dash;
  const he = attr(child(ln, 'headEnd'), 'type');
  const te = attr(child(ln, 'tailEnd'), 'type');
  if (he && he !== 'none') out.head = he === 'stealth' ? 'triangle' : he;
  if (te && te !== 'none') out.tail = te === 'stealth' ? 'triangle' : te;
  if (!fillEl && !phClr) return undefined;
  return out;
}

function readXfrm(x) {
  if (!x) return null;
  const off = child(x, 'off');
  const ext = child(x, 'ext');
  const out = { x: PX(attr(off, 'x')), y: PX(attr(off, 'y')), w: PX(attr(ext, 'cx')), h: PX(attr(ext, 'cy')), rot: num(attr(x, 'rot')) / 60000 };
  if (bool(attr(x, 'flipH'))) out.flipH = true;
  if (bool(attr(x, 'flipV'))) out.flipV = true;
  const chOff = child(x, 'chOff');
  const chExt = child(x, 'chExt');
  if (chOff) { out.chx = PX(attr(chOff, 'x')); out.chy = PX(attr(chOff, 'y')); out.chw = PX(attr(chExt, 'cx')); out.chh = PX(attr(chExt, 'cy')); }
  return out;
}

/** 단락 수준 서식 묶음 (lvlNpPr) 읽기 */
function readLvl(el, ctx) {
  if (!el) return null;
  const out = {};
  const a = el.attrs;
  if (a.algn) out.align = a.algn;
  if (a.marL !== undefined) out.marL = PX(a.marL);
  if (a.indent !== undefined) out.indent = PX(a.indent);
  const lnSpc = child(el, 'lnSpc');
  if (lnSpc) {
    const p = child(lnSpc, 'spcPct');
    const pts = child(lnSpc, 'spcPts');
    if (p) out.lineSpacing = num(attr(p, 'val')) / 100000;
    if (pts) out.lineSpacingPt = num(attr(pts, 'val')) / 100;
  }
  for (const [tag, key] of [['spcBef', 'spcBef'], ['spcAft', 'spcAft']]) {
    const s = child(el, tag);
    if (!s) continue;
    const pts = child(s, 'spcPts');
    const p = child(s, 'spcPct');
    if (pts) out[key] = num(attr(pts, 'val')) / 100;
    else if (p) out[key] = (num(attr(p, 'val')) / 100000) * 18 * 1.2;
  }
  if (child(el, 'buNone')) out.bullet = { type: 'none' };
  const buChar = child(el, 'buChar');
  const buAuto = child(el, 'buAutoNum');
  if (buChar) out.bullet = { type: 'char', char: attr(buChar, 'char') };
  if (buAuto) out.bullet = { type: 'num', scheme: attr(buAuto, 'type') === 'ea1JpnKorPeriod' ? 'ganada' : attr(buAuto, 'type'), start: num(attr(buAuto, 'startAt'), 1) };
  if (child(el, 'buBlip')) out.bullet = { type: 'char', char: '■' };
  const buClr = child(el, 'buClr');
  if (buClr) out.buColor = readColor(buClr, ctx);
  const buSz = child(el, 'buSzPct');
  if (buSz) out.buSize = num(attr(buSz, 'val')) / 100000;
  const buFont = child(el, 'buFont');
  if (buFont) out.buFont = attr(buFont, 'typeface');
  const def = child(el, 'defRPr');
  if (def) out.rpr = readRPr(def, ctx);
  return out;
}

function readRPr(el, ctx, phClr) {
  const out = {};
  if (!el) return out;
  const a = el.attrs;
  if (a.sz) out.size = num(a.sz) / 100;
  if (a.b !== undefined) out.b = bool(a.b);
  if (a.i !== undefined) out.i = bool(a.i);
  if (a.u && a.u !== 'none') out.u = true;
  if (a.strike && a.strike !== 'noStrike') out.s = true;
  if (a.baseline) out.base = num(a.baseline) > 0 ? 1 : -1;
  if (a.spc) out.spc = num(a.spc) / 100;
  if (a.cap && a.cap !== 'none') out.cap = a.cap;
  const fillEl = child(el, 'solidFill') ?? child(el, 'gradFill');
  if (fillEl) out.color = readColor(fillEl.name === 'gradFill' ? descendants(fillEl, 'gs')[0] : fillEl, ctx, phClr);
  const hl = child(el, 'highlight');
  if (hl) out.hl = readColor(hl, ctx);
  const ea = attr(child(el, 'ea'), 'typeface');
  const latin = attr(child(el, 'latin'), 'typeface');
  const font = [ea, latin].find((f) => f && !/^\+(mj|mn)-(cs)$/.test(f));
  if (font) out.font = font.startsWith('+mj') ? '+mj' : font.startsWith('+mn') ? '+mn' : font;
  if (child(child(el, 'effectLst'), 'outerShdw')) out.shadow = true;
  const ln = child(el, 'ln');
  if (ln && !child(ln, 'noFill') && child(ln, 'solidFill')) out.outline = { color: readColor(child(ln, 'solidFill'), ctx), width: PX(attr(ln, 'w') ?? 9525) };
  const link = child(el, 'hlinkClick');
  if (link && ctx.links) { const url = ctx.links(attr(link, 'id')); if (url) out.link = url; }
  return out;
}

/** 목록 스타일 (lstStyle / txStyles 의 각 수준) */
function readListStyle(el, ctx) {
  if (!el) return null;
  const out = {};
  for (let l = 1; l <= 9; l++) {
    const v = readLvl(child(el, `lvl${l}pPr`), ctx);
    if (v) out[l - 1] = v;
  }
  const def = readLvl(child(el, 'defPPr'), ctx);
  if (def) out.def = def;
  return out;
}

function mergeLvl(base, over) {
  if (!over) return base;
  const out = { ...base, ...over };
  out.rpr = { ...(base?.rpr ?? {}), ...(over.rpr ?? {}) };
  return out;
}

function readBodyPr(bp) {
  if (!bp) return {};
  const a = bp.attrs;
  const out = {};
  if (a.anchor) out.anchor = a.anchor === 'ctr' ? 'ctr' : a.anchor === 'b' ? 'b' : 't';
  const ins = ['lIns', 'tIns', 'rIns', 'bIns'];
  if (ins.some((k) => a[k] !== undefined)) out.insets = ins.map((k, i) => (a[k] !== undefined ? PX(a[k]) : [9.6, 4.8, 9.6, 4.8][i]));
  if (a.wrap) out.wrap = a.wrap !== 'none';
  if (a.vert && a.vert !== 'horz') out.vert = a.vert;
  if (child(bp, 'spAutoFit')) out.autofit = 'resize';
  const nf = child(bp, 'normAutofit');
  if (nf) { out.autofit = 'shrink'; if (attr(nf, 'fontScale')) out.fontScale = num(attr(nf, 'fontScale')) / 100000; if (attr(nf, 'lnSpcReduction')) out.lnSpcReduction = num(attr(nf, 'lnSpcReduction')) / 100000; }
  if (child(bp, 'noAutofit')) out.autofit = 'none';
  return out;
}

/** txBody → 글 (상속 단계: styles = [수준별 기본값 …]) */
function readTxBody(tx, ctx, chain, bodyBase, fontRefColor) {
  if (!tx) return null;
  const bp = { ...bodyBase, ...readBodyPr(child(tx, 'bodyPr')) };
  const own = readListStyle(child(tx, 'lstStyle'), ctx);
  const levels = [...chain, own].filter(Boolean);
  const lvlStyle = (lvl) => {
    let s = {};
    for (const L of levels) { s = mergeLvl(s, L.def); s = mergeLvl(s, L[lvl]); }
    return s;
  };
  const paras = kids(tx, 'p').map((p) => {
    const pPr = child(p, 'pPr');
    const lvl = num(attr(pPr, 'lvl'));
    const base = mergeLvl(lvlStyle(lvl), readLvl(pPr, ctx));
    const rp = { ...(fontRefColor ? { color: fontRefColor } : {}), ...(base.rpr ?? {}) };
    const runs = [];
    for (const c of p.children) {
      if (c.name === 'r' || c.name === 'fld') {
        const r = { ...rp, ...readRPr(child(c, 'rPr'), ctx) };
        const t = child(c, 't')?.text ?? '';
        if (t) runs.push(cleanRun({ t, ...r }));
      } else if (c.name === 'br') {
        const last = runs[runs.length - 1];
        if (last) last.t += '\v'; else runs.push(cleanRun({ t: '\v', ...rp }));
      }
    }
    const endR = { ...rp, ...readRPr(child(p, 'endParaRPr'), ctx) };
    const bullet = base.bullet && base.bullet.type !== 'none' ? { ...base.bullet, ...(base.buColor ? { color: base.buColor } : {}), ...(base.buSize ? { size: base.buSize } : {}), ...(base.buFont && base.bullet.type === 'char' ? { font: base.buFont } : {}) } : null;
    const out = {
      runs, align: { l: 'l', ctr: 'ctr', r: 'r', just: 'just', dist: 'dist' }[base.align] ?? 'l', lvl,
      bullet, marL: base.marL ?? 0, indent: base.indent ?? 0,
      lineSpacing: base.lineSpacing ?? 1, spcBef: base.spcBef ?? 0, spcAft: base.spcAft ?? 0,
      end: cleanRun({ ...endR }),
    };
    if (base.lineSpacingPt) out.lineSpacingPt = base.lineSpacingPt;
    delete out.end.t;
    return out;
  });
  return { paras: paras.length ? paras : [para()], anchor: bp.anchor ?? 't', insets: bp.insets ?? [9.6, 4.8, 9.6, 4.8], wrap: bp.wrap ?? true, autofit: bp.autofit ?? 'none', ...(bp.fontScale ? { fontScale: bp.fontScale } : {}), ...(bp.lnSpcReduction ? { lnSpcReduction: bp.lnSpcReduction } : {}), ...(bp.vert ? { vert: bp.vert } : {}) };
}
function cleanRun(r) {
  for (const k of Object.keys(r)) if (r[k] === undefined || r[k] === null || r[k] === false) delete r[k];
  if (r.size == null) r.size = 18;
  return r;
}

function readGeom(spPr) {
  const pg = child(spPr, 'prstGeom');
  if (pg) {
    const prst = attr(pg, 'prst');
    const adj = {};
    for (const gd of descendants(child(pg, 'avLst'), 'gd')) { const m = /val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? ''); if (m) adj[attr(gd, 'name')] = Number(m[1]); }
    return { shape: prst, ...(Object.keys(adj).length ? { adj } : {}) };
  }
  const cg = child(spPr, 'custGeom');
  if (cg) {
    const paths = kids(child(cg, 'pathLst'), 'path').map((p) => {
      const cmds = [];
      for (const c of p.children) {
        const pts = kids(c, 'pt').flatMap((pt) => [num(attr(pt, 'x')), num(attr(pt, 'y'))]);
        if (c.name === 'moveTo') cmds.push(['M', ...pts]);
        else if (c.name === 'lnTo') cmds.push(['L', ...pts]);
        else if (c.name === 'cubicBezTo') cmds.push(['C', ...pts]);
        else if (c.name === 'quadBezTo') cmds.push(['Q', ...pts]);
        else if (c.name === 'arcTo') cmds.push(['A', num(attr(c, 'wR')), num(attr(c, 'hR')), num(attr(c, 'stAng')), num(attr(c, 'swAng'))]);
        else if (c.name === 'close') cmds.push(['Z']);
      }
      return { w: num(attr(p, 'w')), h: num(attr(p, 'h')), cmds, ...(attr(p, 'fill') === 'none' ? { fill: 'none' } : {}) };
    });
    return { shape: 'rect', path: paths };
  }
  return { shape: 'rect' };
}

function themeFromXml(xml) {
  const root = parseXml(xml);
  const cs = descendants(root, 'clrScheme')[0];
  const colors = {};
  for (const k of SLOT_NAMES) {
    const el = child(cs, k);
    const c = el?.children[0];
    colors[k] = c?.name === 'sysClr' ? `#${(attr(c, 'lastClr') ?? '000000').toUpperCase()}` : c ? `#${(attr(c, 'val') ?? '000000').toUpperCase()}` : '#000000';
  }
  const fontOf = (el) => {
    const hang = kids(el, 'font').find((f) => attr(f, 'script') === 'Hang');
    const ea = attr(child(el, 'ea'), 'typeface');
    return attr(hang, 'typeface') || ea || attr(child(el, 'latin'), 'typeface') || '맑은 고딕';
  };
  const fs = descendants(root, 'fontScheme')[0];
  const fmt = descendants(root, 'fmtScheme')[0];
  return {
    theme: { name: 'file', label: attr(root, 'name') ?? '파일 테마', colors, fonts: { major: fontOf(child(fs, 'majorFont')), minor: fontOf(child(fs, 'minorFont')) } },
    fillStyles: child(fmt, 'fillStyleLst')?.children ?? [],
    lineStyles: child(fmt, 'lnStyleLst')?.children ?? [],
    bgStyles: child(fmt, 'bgFillStyleLst')?.children ?? [],
  };
}

const TABLE_STYLE_ACCENT = {
  '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': 'accent1', '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}': 'accent2', '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}': 'accent3',
  '{00A15C55-8517-42AA-B614-E9B94910E393}': 'accent4', '{7DF18680-E054-41AD-8BC1-D1AEF772440D}': 'accent5', '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}': 'accent6',
};
const PRESET_ANIM = {
  entr: { 1: 'appear', 2: 'fly', 10: 'fade', 16: 'split', 21: 'wheel', 22: 'wipe', 23: 'zoom', 53: 'zoom', 42: 'float', 47: 'float', 31: 'grow', 26: 'bounce', 9: 'fade', 37: 'float', 55: 'zoom' },
  emph: { 26: 'pulse', 8: 'spin', 6: 'growShrink', 32: 'teeter', 35: 'flash' },
  exit: { 1: 'disappear', 2: 'fly', 10: 'fade', 22: 'wipe', 53: 'zoom', 23: 'zoom', 9: 'fade' },
};
const SUB_DIR = { 4: 'b', 1: 't', 8: 'l', 2: 'r' };
const LAYOUT_KIND = Object.fromEntries(Object.entries(LAYOUT_TYPE).map(([k, v]) => [v, k]));
const TRANS_NAME = { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'uncover', zoom: 'zoom', circle: 'circle', dissolve: 'dissolve', cut: 'cut', flip: 'flip', random: 'fade', blinds: 'wipe', checker: 'dissolve', comb: 'wipe', randomBar: 'wipe', strips: 'cover', wedge: 'circle', wheel: 'circle', newsflash: 'zoom', plus: 'zoom', diamond: 'zoom', vortex: 'zoom', ripple: 'zoom', honeycomb: 'dissolve', glitter: 'dissolve', shred: 'dissolve', switch: 'flip', gallery: 'push', cube: 'push', doors: 'split', box: 'push', conveyor: 'push', pan: 'push', ferris: 'push', flythrough: 'zoom', warp: 'zoom', prism: 'push', reveal: 'fade', flash: 'fade', vortex2: 'zoom', morph: 'fade', prstTrans: 'fade' };
const DIR_FROM = { u: 'b', d: 't', l: 'r', r: 'l' };

/**
 * pptx 바이트 → 문서
 * 반환 { pres, warnings }
 */
export function readPptx(bytes) {
  const files = unzip(bytes);
  const has = (p) => Object.prototype.hasOwnProperty.call(files, p);
  const xmlOf = (p) => (has(p) ? parseXml(textOf(files[p])) : null);
  const relCache = new Map();
  const relsOf = (part) => {
    if (relCache.has(part)) return relCache.get(part);
    const map = new Map();
    const x = xmlOf(relsPath(part));
    for (const r of kids(x, 'Relationship')) map.set(attr(r, 'Id'), { type: attr(r, 'Type') ?? '', target: attr(r, 'TargetMode') === 'External' ? attr(r, 'Target') : resolvePath(part, attr(r, 'Target') ?? ''), external: attr(r, 'TargetMode') === 'External' });
    relCache.set(part, map);
    return map;
  };
  const warnings = new Set();
  if (has('EncryptedPackage') || has('\u0006DataSpaces/DataSpaceMap')) throw new Error('암호로 보호된 프레젠테이션은 열 수 없습니다. PowerPoint 에서 암호를 해제한 뒤 열어 주세요.');
  const root = xmlOf('_rels/.rels');
  const office = kids(root, 'Relationship').find((r) => (attr(r, 'Type') ?? '').endsWith('/officeDocument'));
  const presPath = office ? sanitizeName(attr(office, 'Target')) : 'ppt/presentation.xml';
  const presXml = xmlOf(presPath);
  if (!presXml) throw new Error('PowerPoint 프레젠테이션(.pptx)이 아닙니다');
  const presRels = relsOf(presPath);
  const sz = child(presXml, 'sldSz');
  const size = { w: Math.round(PX(attr(sz, 'cx') ?? 12192000)), h: Math.round(PX(attr(sz, 'cy') ?? 6858000)) };

  const media = {};
  const mediaByPath = new Map();
  const loadMedia = (path) => {
    if (!path || !has(path)) return null;
    if (mediaByPath.has(path)) return mediaByPath.get(path);
    const ext = path.split('.').pop().toLowerCase();
    const mime = MIME_BY_EXT[ext];
    if (!mime) { warnings.add(`지원하지 않는 그림 형식(.${ext})은 빈 칸으로 표시합니다`); mediaByPath.set(path, null); return null; }
    if (ext === 'emf' || ext === 'wmf') warnings.add('EMF/WMF 그림은 브라우저에서 보이지 않을 수 있습니다');
    const b = files[path];
    let bin = '';
    for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    const id = `m${mediaByPath.size + 1}${uid('').slice(-4)}`;
    media[id] = `data:${mime};base64,${btoa(bin)}`;
    mediaByPath.set(path, id);
    return id;
  };

  // 테마 (첫 마스터의 테마)
  const masterIds = descendants(child(presXml, 'sldMasterIdLst'), 'sldMasterId');
  const firstMaster = masterIds.length ? presRels.get(rid(masterIds[0]))?.target : null;
  let themeInfo = null;
  if (firstMaster) {
    const tRel = [...relsOf(firstMaster).values()].find((r) => r.type.endsWith('/theme'));
    if (tRel && has(tRel.target)) themeInfo = themeFromXml(textOf(files[tRel.target]));
  }
  const theme = themeInfo?.theme ?? cloneTheme(THEMES[0]);

  // 마스터 · 레이아웃 캐시
  const masterCache = new Map();
  const layoutCache = new Map();
  const partCtx = (part, clrMap) => ({
    clrMap,
    media: (rid) => { const r = relsOf(part).get(rid); return r ? loadMedia(r.target) : null; },
    links: (rid) => { const r = relsOf(part).get(rid); return r?.external ? r.target : null; },
  });
  const readClrMap = (el) => (el ? Object.fromEntries(Object.entries(el.attrs)) : null);
  const phKey = (sp) => { const ph = descendants(child(sp, 'nvSpPr') ?? child(sp, 'nvPicPr') ?? child(sp, 'nvGraphicFramePr'), 'ph')[0]; return ph ? { type: attr(ph, 'type') ?? 'body', idx: attr(ph, 'idx') ?? null } : null; };
  const phList = (spTree) => (spTree?.children ?? []).filter((c) => ['sp', 'pic'].includes(c.name)).map((sp) => ({ key: phKey(sp), sp })).filter((x) => x.key);
  const loadMaster = (path) => {
    if (masterCache.has(path)) return masterCache.get(path);
    const x = xmlOf(path);
    const clrMap = readClrMap(child(x, 'clrMap'));
    const ctx = partCtx(path, clrMap);
    const tx = child(x, 'txStyles');
    const m = { path, x, clrMap, ctx, phs: phList(descendants(x, 'spTree')[0]), spTree: descendants(x, 'spTree')[0], bg: child(child(x, 'cSld'), 'bg'), styles: { title: readListStyle(child(tx, 'titleStyle'), ctx), body: readListStyle(child(tx, 'bodyStyle'), ctx), other: readListStyle(child(tx, 'otherStyle'), ctx) } };
    masterCache.set(path, m);
    return m;
  };
  const loadLayout = (path) => {
    if (layoutCache.has(path)) return layoutCache.get(path);
    const x = xmlOf(path);
    const mRel = [...relsOf(path).values()].find((r) => r.type.endsWith('/slideMaster'));
    const master = loadMaster(mRel?.target ?? firstMaster);
    const clrOvr = descendants(child(x, 'clrMapOvr'), 'overrideClrMapping')[0];
    const clrMap = clrOvr ? readClrMap(clrOvr) : master.clrMap;
    const l = { path, x, master, clrMap, ctx: partCtx(path, clrMap), phs: phList(descendants(x, 'spTree')[0]), spTree: descendants(x, 'spTree')[0], bg: child(child(x, 'cSld'), 'bg'), type: attr(x, 'type'), showMasterSp: attr(x, 'showMasterSp') !== '0' };
    layoutCache.set(path, l);
    return l;
  };
  const findPh = (list, key) => {
    if (!key) return null;
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'obj' ? 'body' : t);
    if (key.idx != null) { const byIdx = list.find((p) => p.key.idx === key.idx && (norm(p.key.type) === norm(key.type) || key.type === 'body' || key.type === 'obj')); if (byIdx) return byIdx.sp; }
    return list.find((p) => p.key.type === key.type)?.sp ?? list.find((p) => norm(p.key.type) === norm(key.type))?.sp ?? (key.idx != null ? list.find((p) => p.key.idx === key.idx)?.sp : null) ?? null;
  };
  const styleFor = (master, key) => {
    if (!key) return master.styles.other;
    if (key.type === 'title' || key.type === 'ctrTitle') return master.styles.title;
    if (['dt', 'ftr', 'sldNum'].includes(key.type)) return master.styles.other;
    return master.styles.body;
  };

  /** 그림 · 도형 · 표 등 개체 하나 읽기 (layout: 상속용) */
  const readShapeTree = (spTree, ctx, layout, xf, out, opts = {}) => {
    for (const c of spTree?.children ?? []) {
      try {
        readElement(c, ctx, layout, xf, out, opts);
      } catch (e) {
        warnings.add(`일부 개체를 읽지 못했습니다 (${c.name}): ${e.message}`);
      }
    }
  };
  const applyXf = (xf, r) => {
    if (!xf || !r) return r;
    const sx = xf.chw ? xf.w / xf.chw : 1;
    const sy = xf.chh ? xf.h / xf.chh : 1;
    return { ...r, x: xf.x + (r.x - xf.chx) * sx, y: xf.y + (r.y - xf.chy) * sy, w: r.w * sx, h: r.h * sy };
  };
  const styleRefs = (sp, ctx) => {
    const st = child(sp, 'style');
    if (!st) return {};
    const out = {};
    const fr = child(st, 'fillRef');
    const lr = child(st, 'lnRef');
    const fo = child(st, 'fontRef');
    if (fr && num(attr(fr, 'idx')) > 0) {
      const col = readColor(fr, ctx);
      const tpl = themeInfo?.fillStyles?.[num(attr(fr, 'idx')) - 1];
      if (tpl?.name === 'gradFill') { const g = readFill({ children: [tpl] }, ctx, col); out.fill = g ?? { type: 'solid', color: col }; } else out.fill = col ? { type: 'solid', color: col } : undefined;
    } else if (fr) out.fill = null;
    if (lr && num(attr(lr, 'idx')) > 0) {
      const col = readColor(lr, ctx);
      const tpl = themeInfo?.lineStyles?.[num(attr(lr, 'idx')) - 1];
      out.line = { color: col ?? '#000000', width: tpl && attr(tpl, 'w') ? PX(attr(tpl, 'w')) : 1 };
    } else if (lr) out.line = null;
    if (fo) out.fontColor = readColor(fo, ctx);
    return out;
  };

  function readElement(c, ctx, layout, xf, out, opts) {
    if (c.name === 'AlternateContent') {
      // 대체본(Fallback)이 있으면 그것을 — 새 기능(Choice)은 확장 네임스페이스가 필요할 수 있음
      const target = child(c, 'Fallback') ?? child(c, 'Choice');
      for (const x of target?.children ?? []) readElement(x, ctx, layout, xf, out, opts);
      return;
    }
    if (c.name === 'grpSp') {
      const gx = readXfrm(child(child(c, 'grpSpPr'), 'xfrm'));
      const composed = gx ? (xf ? { ...applyXf(xf, gx), chx: gx.chx, chy: gx.chy, chw: gx.chw, chh: gx.chh } : gx) : xf;
      const gFill = readFill(child(c, 'grpSpPr'), ctx);
      const sub = [];
      readShapeTree(c, { ...ctx, groupFill: gFill ?? ctx.groupFill }, layout, composed, sub, opts);
      const gid = opts.noGroup ? null : uid('g');
      for (const o of sub) { if (gid) o.grp = gid; out.push(o); }
      return;
    }
    if (c.name === 'sp' || c.name === 'cxnSp') {
      const nv = child(c, 'nvSpPr') ?? child(c, 'nvCxnSpPr');
      const cNvPr = child(nv, 'cNvPr');
      if (bool(attr(cNvPr, 'hidden'))) return;
      const key = phKey(c);
      if (key && opts.skipPh) return;
      if (key && ['dt', 'ftr', 'sldNum'].includes(key.type) && (opts.isMaster || opts.skipFooters)) return;
      const spPr = child(c, 'spPr');
      let lph = null;
      let mph = null;
      if (key && layout) { lph = findPh(layout.phs, key); mph = findPh(layout.master.phs, key); }
      const xr = readXfrm(child(spPr, 'xfrm')) ?? readXfrm(child(child(lph, 'spPr'), 'xfrm')) ?? readXfrm(child(child(mph, 'spPr'), 'xfrm'));
      if (!xr) return;
      const rect = applyXf(xf, xr);
      const refs = styleRefs(c, ctx);
      const geom = readGeom(spPr);
      if (!child(spPr, 'prstGeom') && !child(spPr, 'custGeom')) { const g2 = readGeom(child(lph, 'spPr') ?? child(mph, 'spPr')); Object.assign(geom, g2); }
      if (!KNOWN_SHAPES.has(geom.shape) && !geom.path && !['rect', 'line', 'bentConnector3', 'curvedConnector3', 'roundRect', 'ellipse'].includes(geom.shape)) geom.shape = ['bentConnector2', 'bentConnector4', 'curvedConnector2', 'curvedConnector4'].includes(geom.shape) ? 'line' : geom.shape;
      let fill = readFill(spPr, ctx);
      if (fill === undefined && lph) fill = readFill(child(lph, 'spPr'), layout.ctx);
      if (fill === undefined) fill = refs.fill ?? null;
      let line = readLine(child(spPr, 'ln'), ctx, refs.line?.color);
      if (line === undefined) line = refs.line ?? null;
      const isTx = bool(attr(child(nv, 'cNvSpPr'), 'txBox'));
      const chain = [];
      let bodyBase = {};
      if (layout) {
        chain.push(styleFor(layout.master, key));
        if (mph) { chain.push(readListStyle(child(child(mph, 'txBody'), 'lstStyle'), layout.master.ctx)); bodyBase = { ...bodyBase, ...readBodyPr(child(child(mph, 'txBody'), 'bodyPr')) }; }
        if (lph) { chain.push(readListStyle(child(child(lph, 'txBody'), 'lstStyle'), layout.ctx)); bodyBase = { ...bodyBase, ...readBodyPr(child(child(lph, 'txBody'), 'bodyPr')) }; }
      }
      const text = readTxBody(child(c, 'txBody'), ctx, chain.filter(Boolean), bodyBase, refs.fontColor);
      const o = {
        id: uid(), type: 'shape', ...rect, ...geom, fill, line,
        text: text ?? (key && key.type !== 'pic' ? readTxBody({ children: [{ name: 'bodyPr', attrs: {}, children: [] }, { name: 'p', attrs: {}, children: [] }] }, ctx, chain.filter(Boolean), bodyBase, refs.fontColor) : null),
        name: attr(cNvPr, 'name'),
      };
      if (attr(cNvPr, 'descr')) o.alt = attr(cNvPr, 'descr');
      const link = child(cNvPr, 'hlinkClick');
      if (link) { const url = ctx.links?.(attr(link, 'id')); if (url) o.link = url; }
      if (isTx) o.txBox = true;
      if (child(child(spPr, 'effectLst'), 'outerShdw')) o.shadow = true;
      if (key) {
        o.ph = { ctrTitle: 'ctrTitle', title: 'title', subTitle: 'subTitle', body: 'body', obj: 'body', pic: 'pic', dt: 'dt', ftr: 'ftr', sldNum: 'sldNum', tbl: 'body', chart: 'body', media: 'body', clipArt: 'body', dgm: 'body' }[key.type] ?? 'body';
        if (o.ph === 'pic' && (!text || isEmptyText(text))) { out.push({ id: o.id, type: 'image', ...rect, ph: 'pic', media: null }); return; }
        if (['dt', 'ftr', 'sldNum'].includes(o.ph)) delete o.ph;
      }
      if (o.text && isEmptyText(o.text) && !o.ph && !o.fill && !o.line && !geom.path) return;
      out.push(o);
      return;
    }
    if (c.name === 'pic') {
      const nv = child(c, 'nvPicPr');
      const cNvPr = child(nv, 'cNvPr');
      if (bool(attr(cNvPr, 'hidden'))) return;
      const spPr = child(c, 'spPr');
      const key = phKey(c);
      let lph = null;
      if (key && layout) lph = findPh(layout.phs, key) ?? findPh(layout.master.phs, key);
      const xr = readXfrm(child(spPr, 'xfrm')) ?? readXfrm(child(child(lph, 'spPr'), 'xfrm'));
      if (!xr) return;
      const rect = applyXf(xf, xr);
      const bf = child(c, 'blipFill');
      const blip = child(bf, 'blip');
      // SVG 원본이 있으면 SVG 사용 (asvg:svgBlip)
      const svg = descendants(blip, 'svgBlip')[0];
      const mid = (svg && ctx.media(attr(svg, 'embed'))) || ctx.media(attr(blip, 'embed'));
      const sr = child(bf, 'srcRect');
      const o = { id: uid(), type: 'image', ...rect, media: mid, name: attr(cNvPr, 'name') };
      if (sr) o.crop = { l: num(attr(sr, 'l')) / 100000, t: num(attr(sr, 't')) / 100000, r: num(attr(sr, 'r')) / 100000, b: num(attr(sr, 'b')) / 100000 };
      if (attr(cNvPr, 'descr')) o.alt = attr(cNvPr, 'descr');
      const g = readGeom(spPr);
      if (g.shape && g.shape !== 'rect') { o.shape = g.shape; if (g.adj) o.adj = g.adj; }
      const ln = readLine(child(spPr, 'ln'), ctx);
      if (ln) o.line = ln;
      const am = child(blip, 'alphaModFix');
      if (am) o.alpha = num(attr(am, 'amt'), 100000) / 100000;
      if (child(blip, 'grayscl')) o.gray = true;
      if (child(child(spPr, 'effectLst'), 'outerShdw')) o.shadow = true;
      const link = child(cNvPr, 'hlinkClick');
      if (link) { const url = ctx.links?.(attr(link, 'id')); if (url) o.link = url; }
      if (key) o.ph = 'pic';
      if (!mid && !key) return;
      out.push(o);
      return;
    }
    if (c.name === 'graphicFrame') {
      const xr = readXfrm(child(c, 'xfrm'));
      if (!xr) return;
      const rect = applyXf(xf, xr);
      const gd = descendants(c, 'graphicData')[0];
      const uri = attr(gd, 'uri') ?? '';
      if (uri.endsWith('/table')) { out.push(readTable(child(gd, 'tbl'), ctx, rect)); return; }
      if (uri.endsWith('/chart')) {
        const ch = child(gd, 'chart');
        const rel = relsOf(ctx.part).get(attr(ch, 'id'));
        const chart = rel ? readChart(rel.target) : null;
        if (chart) out.push({ id: uid(), type: 'chart', ...rect, chart });
        return;
      }
      if (uri.endsWith('/diagram')) {
        const rel = [...relsOf(ctx.part).values()].find((r) => r.type.endsWith('/diagramDrawing'));
        const relIds = descendants(gd, 'relIds')[0];
        const dmRel = relIds ? relsOf(ctx.part).get(attr(relIds, 'dm')) : null;
        let drawingPath = rel?.target;
        if (dmRel) {
          // 데이터 파트에 연결된 그림 파트 (dataModelExt relId)
          const dm = xmlOf(dmRel.target);
          const ext = descendants(dm, 'dataModelExt')[0];
          if (ext) { const r2 = relsOf(ctx.part).get(attr(ext, 'relId')); if (r2) drawingPath = r2.target; }
        }
        if (drawingPath && has(drawingPath)) {
          const dx = xmlOf(drawingPath);
          const tree = descendants(dx, 'spTree')[0];
          const dctx = { ...partCtx(drawingPath, ctx.clrMap), part: drawingPath };
          const sub = [];
          readShapeTree(tree, dctx, null, { x: rect.x, y: rect.y, w: rect.w, h: rect.h, chx: 0, chy: 0, chw: rect.w, chh: rect.h }, sub, { noGroup: true });
          const gid = uid('g');
          for (const o of sub) { o.grp = gid; out.push(o); }
        } else warnings.add('SmartArt 일부는 그림으로 저장되지 않아 표시하지 못했습니다');
        return;
      }
      // OLE 개체 등: 대체 그림
      const pic = descendants(c, 'pic')[0];
      if (pic) { readElement(pic, ctx, layout, null, out, opts); const last = out[out.length - 1]; if (last) Object.assign(last, rect); return; }
      warnings.add('표시할 수 없는 개체(OLE 등)를 건너뛰었습니다');
    }
  }

  function readTable(tbl, ctx, rect) {
    const tblPr = child(tbl, 'tblPr');
    const cols = kids(child(tbl, 'tblGrid'), 'gridCol').map((g) => PX(attr(g, 'w')));
    const wpExt = descendants(tblPr, 'table')[0];
    let wp = null;
    try { wp = wpExt ? JSON.parse(attr(wpExt, 'json')) : null; } catch { wp = null; }
    const styleId = child(tblPr, 'tableStyleId')?.text?.trim();
    const style = wp?.style ?? (styleId ? { firstRow: bool(attr(tblPr, 'firstRow')), banded: bool(attr(tblPr, 'bandRow')), lastRow: bool(attr(tblPr, 'lastRow')), firstCol: bool(attr(tblPr, 'firstCol')), accent: TABLE_STYLE_ACCENT[styleId] ?? 'accent1' } : { none: true });
    const own = wp ? new Set((wp.own ?? []).map(([r, c]) => `${r},${c}`)) : null;
    const rows = kids(tbl, 'tr').map((tr, ri) => ({
      h: PX(attr(tr, 'h')),
      cells: kids(tr, 'tc').map((tc, ci) => {
        const tcPr = child(tc, 'tcPr');
        const text = readTxBody(child(tc, 'txBody'), ctx, [], {}, null) ?? textBody();
        const ins = [attr(tcPr, 'marL') != null ? PX(attr(tcPr, 'marL')) : 9.6, attr(tcPr, 'marT') != null ? PX(attr(tcPr, 'marT')) : 4.8, attr(tcPr, 'marR') != null ? PX(attr(tcPr, 'marR')) : 9.6, attr(tcPr, 'marB') != null ? PX(attr(tcPr, 'marB')) : 4.8];
        text.insets = ins;
        text.anchor = { ctr: 'ctr', b: 'b' }[attr(tcPr, 'anchor')] ?? 't';
        const cell = { text };
        const fill = readFill(tcPr, ctx);
        if (fill !== undefined && (!own || own.has(`${ri},${ci}`))) cell.fill = fill;
        if (style.none && fill === undefined) cell.fill = null;
        if (attr(tc, 'gridSpan')) cell.span = num(attr(tc, 'gridSpan'));
        if (attr(tc, 'rowSpan')) cell.rowSpan = num(attr(tc, 'rowSpan'));
        if (bool(attr(tc, 'hMerge'))) cell.hMerge = true;
        if (bool(attr(tc, 'vMerge'))) cell.vMerge = true;
        if (!wp) {
          const borders = {};
          for (const [tag, side] of [['lnL', 'l'], ['lnR', 'r'], ['lnT', 't'], ['lnB', 'b']]) {
            const ln = child(tcPr, tag);
            if (!ln) continue;
            const l = readLine(ln, ctx);
            if (l) borders[side] = { color: l.color, width: l.width };
            else if (l === null) borders[side] = { color: 'rgba(0,0,0,0)', width: 0 };
          }
          if (Object.keys(borders).length) cell.borders = borders;
          // 표 스타일에 맡긴 글자 색은 스타일 색을 쓰도록 비움
          if (styleId) for (const p of text.paras) for (const r of p.runs) if (r.color === undefined) delete r.color;
        } else {
          for (const p of text.paras) {
            for (const r of p.runs) { delete r.color; delete r.b; }
            if (p.end) { delete p.end.color; delete p.end.b; }
          }
        }
        return cell;
      }),
    }));
    return { id: uid(), type: 'table', ...rect, cols, rows, style, h: rows.reduce((s, r) => s + r.h, 0) || rect.h };
  }

  function readChart(path) {
    const x = xmlOf(path);
    if (!x) return null;
    const ext = descendants(x, 'chart').find((e) => attr(e, 'json'));
    if (ext) { try { return JSON.parse(attr(ext, 'json')); } catch { /* 아래로 */ } }
    const plot = descendants(x, 'plotArea')[0];
    const typeEl = plot?.children.find((c) => /Chart$/.test(c.name));
    if (!typeEl) return null;
    const kindMap = { barChart: attr(child(typeEl, 'barDir'), 'val') === 'bar' ? 'bar' : 'col', bar3DChart: attr(child(typeEl, 'barDir'), 'val') === 'bar' ? 'bar' : 'col', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', radarChart: 'line', bubbleChart: 'scatter' };
    const kind = kindMap[typeEl.name] ?? 'col';
    const vals = (el) => {
      const cache = descendants(el, 'numCache')[0] ?? descendants(el, 'numLit')[0] ?? descendants(el, 'strCache')[0] ?? descendants(el, 'strLit')[0];
      if (!cache) return [];
      const n = num(attr(child(cache, 'ptCount'), 'val'), 0);
      const arr = Array.from({ length: n }, () => null);
      for (const pt of kids(cache, 'pt')) arr[num(attr(pt, 'idx'))] = child(pt, 'v')?.text ?? '';
      return arr;
    };
    const series = kids(typeEl, 'ser').map((s, i) => {
      const tx = child(s, 'tx');
      const name = tx ? (child(tx, 'v')?.text ?? vals(tx)[0] ?? `계열 ${i + 1}`) : `계열 ${i + 1}`;
      const color = readColor(child(child(s, 'spPr'), 'solidFill'), { clrMap: null });
      return { name, vals: vals(child(s, 'val') ?? child(s, 'yVal')).map((v) => Number(v) || 0), ...(color && color.startsWith('#') ? { color } : {}), _cats: vals(child(s, 'cat') ?? child(s, 'xVal')) };
    });
    const cats = (series.find((s) => s._cats.length)?._cats ?? []).map((v) => v ?? '');
    for (const s of series) delete s._cats;
    const chartEl = descendants(x, 'chart')[0];
    const titleEl = child(chartEl, 'title');
    const autoDeleted = bool(attr(child(chartEl, 'autoTitleDeleted'), 'val'));
    // 제목 글이 없으면: 계열이 하나일 때 계열 이름이 자동 제목 (PowerPoint 와 같음)
    const title = descendants(titleEl, 't').map((t) => t.text).join('') || ((titleEl || !autoDeleted) && series.length === 1 && !['pie', 'doughnut'].includes(kind) ? series[0].name : (titleEl && series.length === 1 ? series[0].name : ''));
    const legend = attr(child(descendants(x, 'legend')[0], 'legendPos'), 'val');
    const dl = descendants(typeEl, 'dLbls')[0];
    return { kind, title, cats: cats.length ? cats : series[0]?.vals.map((_, i) => String(i + 1)) ?? [], series, legend: legend ? (legend === 'r' ? 'r' : 'b') : null, labels: !!dl && (bool(attr(child(dl, 'showVal'), 'val')) || bool(attr(child(dl, 'showPercent'), 'val'))), stacked: /stacked/i.test(attr(child(typeEl, 'grouping'), 'val') ?? '') };
  }

  /** 배경 채우기 (bgPr 또는 테마 bgRef) */
  const readBg = (bg, ctx) => {
    if (!bg) return undefined;
    const bgPr = child(bg, 'bgPr');
    if (bgPr) return readFill(bgPr, ctx) ?? null;
    const ref = child(bg, 'bgRef');
    if (ref) {
      const idx = num(attr(ref, 'idx'));
      const col = readColor(ref, ctx);
      const tpl = idx >= 1001 ? themeInfo?.bgStyles?.[idx - 1001] : themeInfo?.fillStyles?.[idx - 1];
      if (tpl && tpl.name !== 'solidFill') { const f = readFill({ children: [tpl] }, ctx, col); if (f) return f; }
      return col ? { type: 'solid', color: col } : null;
    }
    return undefined;
  };

  // 문서 수준 WIPOINT 확장 (테마 장식 · 바닥글)
  let presMeta = null;
  const pm = descendants(child(presXml, 'extLst'), 'pres').find((e) => attr(e, 'json'));
  if (pm) { try { presMeta = JSON.parse(attr(pm, 'json')); } catch { presMeta = null; } }
  if (presMeta?.theme) { Object.assign(theme, { name: presMeta.theme.name ?? theme.name, label: presMeta.theme.label ?? theme.label, accentBand: presMeta.theme.accentBand, titleBar: presMeta.theme.titleBar, dark: presMeta.theme.dark }); if (presMeta.theme.bg) theme.bg = presMeta.theme.bg; }

  // 구역
  const sectionOf = new Map();
  for (const sec of descendants(presXml, 'section')) for (const sid of descendants(sec, 'sldId')) if (!sectionOf.has(attr(sid, 'id'))) sectionOf.set(attr(sid, 'id'), attr(sec, 'name'));
  const firstOfSection = new Set();

  const slides = [];
  for (const sid of descendants(child(presXml, 'sldIdLst'), 'sldId')) {
    const rel = presRels.get(rid(sid));
    if (!rel || !has(rel.target)) continue;
    const path = rel.target;
    const x = xmlOf(path);
    const rels = relsOf(path);
    const lRel = [...rels.values()].find((r) => r.type.endsWith('/slideLayout'));
    const layout = lRel ? loadLayout(lRel.target) : null;
    const clrOvr = descendants(child(x, 'clrMapOvr'), 'overrideClrMapping')[0];
    const clrMap = clrOvr ? readClrMap(clrOvr) : layout?.clrMap ?? layout?.master.clrMap;
    const ctx = { ...partCtx(path, clrMap), part: path };
    // WIPOINT 확장 (우리 앱이 쓴 파일이면 전환 · 애니메이션 · 레이아웃을 그대로)
    let meta = null;
    const wpEl = descendants(child(x, 'extLst'), 'slide').find((e) => attr(e, 'json'));
    if (wpEl) { try { meta = JSON.parse(attr(wpEl, 'json')); } catch { meta = null; } }
    const objects = [];
    readShapeTree(descendants(child(x, 'cSld'), 'spTree')[0], ctx, layout, null, objects, { skipFooters: !!(meta && presMeta?.footer) });
    // 배경 장식: 마스터(레이아웃이 숨기지 않으면) + 레이아웃의 개체 틀 아닌 도형
    const bgObjects = [];
    if (layout) {
      const showMaster = layout.showMasterSp && attr(x, 'showMasterSp') !== '0';
      if (showMaster) readShapeTree(layout.master.spTree, { ...layout.master.ctx, part: layout.master.path }, null, null, bgObjects, { skipPh: true, isMaster: true });
      readShapeTree(layout.spTree, { ...layout.ctx, part: layout.path }, null, null, bgObjects, { skipPh: true, isMaster: true });
      for (const o of bgObjects) { o.decor = true; o.id = `bg-${o.id}`; delete o.grp; }
    }
    let bg = readBg(child(child(x, 'cSld'), 'bg'), ctx);
    if (bg === undefined && layout) bg = readBg(layout.bg, layout.ctx);
    if (bg === undefined && layout) bg = readBg(layout.master.bg, layout.master.ctx);
    // 메모
    let notes = '';
    const nRel = [...rels.values()].find((r) => r.type.endsWith('/notesSlide'));
    if (nRel && has(nRel.target)) {
      const nx = xmlOf(nRel.target);
      const body = (descendants(nx, 'sp')).find((sp) => { const k = phKey(sp); return k && k.type === 'body'; });
      if (body) notes = kids(child(body, 'txBody'), 'p').map((p) => descendants(p, 't').map((t) => t.text).join('')).join('\n').replace(/\n+$/, '');
    }
    const slide = {
      id: uid('s'), layout: meta?.layout ?? LAYOUT_KIND[layout?.type] ?? 'blank', bg: bg ?? null, hidden: attr(x, 'show') === '0',
      transition: null, notes, objects, anims: [],
    };
    // 레이아웃 이름이 사용자 지정이면 배경 장식은 원본대로
    // 우리가 쓴 파일의 테마 장식은 테마에서 다시 만들어지므로 원본 장식만 담음
    if (meta?.bgObjects || !meta) slide.bgObjects = bgObjects;
    if (meta?.hideDecor) slide.hideDecor = true;
    // 개체 id 연결 (spid → 개체): 원래 순서대로 cNvPr id 를 모음
    const spids = [];
    const collectIds = (tree) => {
      for (const c of tree?.children ?? []) {
        if (c.name === 'grpSp') { spids.push({ id: attr(child(child(c, 'nvGrpSpPr'), 'cNvPr'), 'id'), group: true }); collectIds(c); continue; }
        if (c.name === 'AlternateContent') { collectIds(child(c, 'Fallback') ?? child(c, 'Choice')); continue; }
        const nv = c.children.find((k) => k.name.startsWith('nv'));
        const cNvPr = child(nv, 'cNvPr');
        if (cNvPr) spids.push({ id: attr(cNvPr, 'id'), el: c });
      }
    };
    collectIds(descendants(child(x, 'cSld'), 'spTree')[0]);
    // 읽은 개체 순서와 spid 순서는 같음 (숨김 · 빈 도형을 건너뛴 것만 다름) → 이름 · 위치로 맞춤
    const spidToObj = new Map();
    {
      let k = 0;
      const shapes = spids.filter((s) => !s.group);
      for (const s of shapes) {
        const nm = attr(child(s.el.children.find((q) => q.name.startsWith('nv')), 'cNvPr'), 'name');
        for (let j = k; j < objects.length; j++) if (objects[j].name === nm || (objects[j].type === 'table' && s.el.name === 'graphicFrame') || (objects[j].type === 'chart' && s.el.name === 'graphicFrame')) { spidToObj.set(s.id, objects[j].id); k = j + 1; break; }
      }
    }
    if (meta) {
      slide.transition = meta.transition ?? null;
      const map = new Map(Object.entries(meta.ids ?? {}).map(([sp, oldId]) => [oldId, spidToObj.get(sp)]));
      slide.anims = (meta.anims ?? []).map((a) => ({ ...a, obj: map.get(a.obj) })).filter((a) => a.obj);
      if (meta.section) slide.section = meta.section;
    } else {
      slide.transition = readTransition(x);
      slide.anims = readTiming(x, spidToObj);
    }
    if (!slide.section && sectionOf.has(attr(sid, 'id'))) {
      const name = sectionOf.get(attr(sid, 'id'));
      if (!firstOfSection.has(name)) { firstOfSection.add(name); slide.section = name; }
    }
    for (const o of objects) delete o.name;
    slides.push(slide);
  }

  function readTransition(x) {
    const t = descendants(x, 'transition')[0];
    if (!t) return null;
    const el = t.children.find((c) => c.name !== 'sndAc' && c.name !== 'extLst');
    const name = el ? TRANS_NAME[el.name] ?? 'fade' : 'none';
    const spd = attr(t, 'spd');
    const dur = attr(t, 'dur') ? num(attr(t, 'dur')) / 1000 : spd === 'fast' ? 0.5 : spd === 'slow' ? 1 : 0.75;
    const dirAttr = attr(el, 'dir');
    const out = { type: name, dur };
    if (dirAttr && DIR_FROM[dirAttr]) out.dir = DIR_FROM[dirAttr];
    if (attr(t, 'advClick') === '0') out.advClick = false;
    if (attr(t, 'advTm')) out.advAfter = num(attr(t, 'advTm')) / 1000;
    return out;
  }

  function readTiming(x, spidToObj) {
    const timing = child(x, 'timing');
    if (!timing) return [];
    const main = descendants(timing, 'cTn').find((c) => attr(c, 'nodeType') === 'mainSeq');
    if (!main) return [];
    const anims = [];
    for (const eff of descendants(main, 'cTn').filter((c) => attr(c, 'presetClass'))) {
      const cls = attr(eff, 'presetClass');
      if (!['entr', 'exit', 'emph'].includes(cls)) continue;
      const spTgt = descendants(eff, 'spTgt')[0];
      const obj = spidToObj.get(attr(spTgt, 'spid'));
      if (!obj) continue;
      const pid = num(attr(eff, 'presetID'));
      const effect = PRESET_ANIM[cls]?.[pid] ?? (cls === 'entr' ? 'fade' : cls === 'exit' ? 'fade' : 'pulse');
      const node = attr(eff, 'nodeType');
      const durs = descendants(eff, 'cTn').map((c) => num(attr(c, 'dur'), 0)).filter((d) => d > 1);
      const delay = num(attr(child(child(eff, 'stCondLst'), 'cond'), 'delay'), 0);
      const a = { id: uid('a'), obj, cls, effect, start: node === 'withEffect' ? 'with' : node === 'afterEffect' ? 'after' : 'click', dur: durs.length ? Math.max(...durs) / 1000 : effect === 'appear' || effect === 'disappear' ? 0 : 0.5, delay: delay / 1000, dir: SUB_DIR[num(attr(eff, 'presetSubtype'))] ?? 'b' };
      // 같은 문단별 효과(bldP build="p")가 여러 번 나오면 하나로
      if (anims.some((b) => b.obj === obj && b.cls === cls && b.effect === effect && b.start !== 'click' && a.start !== 'click')) continue;
      anims.push(a);
    }
    return anims;
  }

  // 문서 속성
  const core = xmlOf('docProps/core.xml');
  const props = { title: descendants(core, 'title')[0]?.text ?? '', author: descendants(core, 'creator')[0]?.text ?? '', created: descendants(core, 'created')[0]?.text ?? new Date().toISOString() };
  const pres = { version: 1, size, theme, slides, media, props, footer: presMeta?.footer ?? { slideNum: false, date: false, text: '', hideOnTitle: true } };
  // 포함된 글꼴 (ppt/fonts/*.fntdata = EOT, MTX 압축일 수 있음) — 원본 그대로 보관 (다시 저장할 때 그대로 넣음), 화면에는 fileio 가 풀어서 등록
  const fonts = [];
  for (const ef of kids(child(presXml, 'embeddedFontLst'), 'embeddedFont')) {
    const f = child(ef, 'font');
    const typeface = attr(f, 'typeface');
    if (!typeface) continue;
    for (const style of ['regular', 'bold', 'italic', 'boldItalic']) {
      const el = child(ef, style);
      const target = el ? presRels.get(rid(el))?.target : null;
      if (!target || !has(target)) continue;
      const bytes = files[target];
      try { decodeEmbeddedFont(bytes); } catch (e) { warnings.add(`포함된 글꼴 '${typeface}' 을(를) 읽지 못했습니다 (${e.message}) — 컴퓨터에 있는 글꼴로 표시합니다`); continue; }
      let bin = '';
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      fonts.push({ typeface, style, panose: attr(f, 'panose') ?? null, pitchFamily: attr(f, 'pitchFamily') ?? null, charset: attr(f, 'charset') ?? null, data: `data:application/x-fontdata;base64,${btoa(bin)}` });
    }
  }
  if (fonts.length) pres.fonts = fonts;
  if (!slides.length) pres.slides.push(newSlide(pres, 'title'));
  return { pres, warnings: [...warnings] };
}
