// .pptx 쓰기 · 읽기 (PresentationML) — DOM 없음
// 쓰기: pptxEntries(pres) → { 경로: 문자열|바이트 } → zip / zipAsync
// 읽기: readPptx(bytes) → pres (마스터 · 레이아웃 상속, 테마 색, 표, 차트, SmartArt 그림, 그룹, 전환, 애니메이션, 메모)
import { placeInGroups, guideVars } from './geom.js';
import { zip, unzip, textOf } from './zip.js';
import { parseXml, child, kids, descendants, esc } from './xml.js';
import { SLOT_NAMES, resolveColor, applyMods, cloneTheme, THEMES } from './themes.js';
import { EMU_PER_PX, uid, layoutPlaceholders, defaultSize, defaultColor, defaultFont, isEmptyText, animSteps, triggerSteps, stepTimeline, MOTION_PATHS, isCustomLayout, layoutName, presForSlide, textBody, para, newSlide, slideTitle, setMasterText } from './model.js';
import { parseLatex, toOMML, ommlToLatex } from './math.js';
import { paraDefaults, tableCellStyle, slideBackground, themeDecor, slideDecor } from './render.js';
import { decodeEmbeddedFont } from './fonts.js';
import { fxOf, fxXml, fxFromXml } from './effects.js';
import { fallbackChart, is3d, VIEW3D_DEFAULT } from './chart.js';
import { ART_BY_XML } from './arteffects.js';
import { SOUNDS, SOUND_BY_FILE, soundWav } from './sounds.js';

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
const STD_CLRMAP = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink' };
const clrMapAttrs = (m) => Object.entries({ ...STD_CLRMAP, ...(m ?? {}) }).filter(([k]) => STD_CLRMAP[k]).map(([k, v]) => `${k}="${v}"`).join(' ');
const isStdMap = (m) => !m || Object.entries(STD_CLRMAP).every(([k, v]) => (m[k] ?? v) === v);
/** 마스터 색 매핑이 표준이 아니면 (예전 PowerPoint 파일: tx1 = lt1) 테마 칸 이름(dk1 · lt1)을 그대로 씀 */
let directSlots = false;
const MOD_CODES = { lm: 'lumMod', lo: 'lumOff', t: 'tint', s: 'shade', a: 'alpha', sm: 'satMod', so: 'satOff', hm: 'hueMod', ho: 'hueOff', am: 'alphaMod', ao: 'alphaOff', cp: 'comp', iv: 'inv', gr: 'gray', rd: 'red', gn: 'green', bl: 'blue', rm: 'redMod', gm: 'greenMod', bm: 'blueMod', ro: 'redOff', go: 'greenOff', bo: 'blueOff' };

function colorXml(c) {
  if (!c) return '';
  const [base, ...rest] = String(c).split(':');
  const mods = rest.map((m) => {
    const x = /^([a-z]+)(-?[\d.]+)$/.exec(m);
    const tag = x && MOD_CODES[x[1]];
    if (!tag) return '';
    if (tag === 'comp' || tag === 'inv' || tag === 'gray') return `<a:${tag}/>`;
    return `<a:${tag} val="${Math.round(Number(x[2]) * (tag === 'hueOff' ? 60000 : 1000))}"/>`;
  }).join('');
  if (base.startsWith('@')) {
    const slot = base.slice(1);
    const val = directSlots ? ({ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[slot] ?? slot) : SLOT_TO_SCHEME[slot] ?? ({ tx1: 'tx1', bg1: 'bg1', tx2: 'tx2', bg2: 'bg2' }[slot] ?? slot);
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
    const fo = fill.focus ?? [0.5, 0.5, 0.5, 0.5];
    const shade = fill.path ? `<a:path path="${fill.path === 'rect' || fill.path === 'shape' ? fill.path : 'circle'}"><a:fillToRect l="${pct(fo[0])}" t="${pct(fo[1])}" r="${pct(fo[2])}" b="${pct(fo[3])}"/></a:path>` : `<a:lin ang="${Math.round((((fill.angle ?? 90) % 360) + 360) % 360 * 60000)}" scaled="0"/>`;
    return `<a:gradFill rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${shade}</a:gradFill>`;
  }
  if (fill.type === 'pattern') return `<a:pattFill prst="${fill.prst ?? 'pct5'}"><a:fgClr>${colorXml(fill.fg ?? '#000000')}</a:fgClr><a:bgClr>${colorXml(fill.bg ?? '#FFFFFF')}</a:bgClr></a:pattFill>`;
  if (fill.type === 'image') {
    const rid = ctx.image(fill.media);
    const t = fill.tile && typeof fill.tile === 'object' ? fill.tile : {};
    const blipKids = `${fill.alpha != null && fill.alpha < 1 ? `<a:alphaModFix amt="${pct(fill.alpha)}"/>` : ''}${fill.duotone?.length === 2 ? `<a:duotone>${colorXml(fill.duotone[0])}${colorXml(fill.duotone[1])}</a:duotone>` : ''}`;
    return rid ? `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"${blipKids ? `>${blipKids}</a:blip>` : '/>'}${fill.tile ? `<a:tile tx="${E(t.tx ?? 0)}" ty="${E(t.ty ?? 0)}" sx="${pct(t.sx ?? 1)}" sy="${pct(t.sy ?? 1)}" flip="${t.flip ?? 'none'}" algn="${t.algn ?? 'tl'}"/>` : '<a:stretch><a:fillRect/></a:stretch>'}</a:blipFill>` : '<a:noFill/>';
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
    const K = 100; // 경로 좌표를 0.01 단위로 (w · h 도 같이 늘려 비율은 같음)
    const paths = o.path.map((p) => `<a:path w="${Math.round(p.w * K)}" h="${Math.round(p.h * K)}"${p.fill === 'none' ? ' fill="none"' : ''}${p.stroke === false ? ' stroke="0"' : ''}>${p.cmds.map(([op, ...v]) => {
      const pt = (x, y) => `<a:pt x="${Math.round(x * K)}" y="${Math.round(y * K)}"/>`;
      if (op === 'M') return `<a:moveTo>${pt(v[0], v[1])}</a:moveTo>`;
      if (op === 'L') return `<a:lnTo>${pt(v[0], v[1])}</a:lnTo>`;
      if (op === 'C') return `<a:cubicBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}${pt(v[4], v[5])}</a:cubicBezTo>`;
      if (op === 'Q') return `<a:quadBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}</a:quadBezTo>`;
      if (op === 'A') return `<a:arcTo wR="${Math.round(v[0] * K)}" hR="${Math.round(v[1] * K)}" stAng="${Math.round(v[2])}" swAng="${Math.round(v[3])}"/>`;
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
    bold != null ? `b="${bold ? 1 : 0}"` : '', r.i ? 'i="1"' : '', r.u ? `u="${r.u === true ? 'sng' : esc(r.u)}"` : '', r.s ? `strike="${r.s === 'dbl' ? 'dblStrike' : 'sngStrike'}"` : '',
    r.kern != null ? `kern="${Math.round(r.kern * 100)}"` : '',
    r.base ? `baseline="${r.base > 0 ? 30000 : -25000}"` : '', r.spc ? `spc="${Math.round(r.spc * 100)}"` : '',
    r.cap === 'all' ? 'cap="all"' : r.cap === 'small' ? 'cap="small"' : '', 'dirty="0"',
  ].filter(Boolean).join(' ');
  const ln = r.outline ? `<a:ln w="${Math.round((r.outline.width ?? 1) * EMU_PER_PX)}">${solidFill(r.outline.color ?? '@tx1')}</a:ln>` : '';
  const fill = r.fill?.type === 'none' ? '<a:noFill/>' : r.fill?.type === 'gradient' ? fillXml(r.fill) : solidFill(r.color ?? defaultColor(o));
  const sh = r.shadow === true ? '<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:prstClr val="black"><a:alpha val="43137"/></a:prstClr></a:outerShdw>'
    : r.shadow ? `<a:outerShdw blurRad="${Math.round((r.shadow.blur ?? 3) * 12700)}" dist="${Math.round((r.shadow.dist ?? 3) * 12700)}" dir="${Math.round((r.shadow.dir ?? 45) * 60000)}" algn="tl" rotWithShape="0">${colorXml(r.shadow.color ?? '#000000:a43')}</a:outerShdw>` : '';
  const gl = r.glow ? `<a:glow rad="${Math.round((r.glow.rad ?? 5) * 12700)}">${colorXml(r.glow.color ?? '@accent1:a40')}</a:glow>` : '';
  const eff = gl || sh ? `<a:effectLst>${gl}${sh}</a:effectLst>` : '';
  const hl = r.hl ? `<a:highlight>${colorXml(r.hl)}</a:highlight>` : '';
  const uf = r.u && r.uColor ? `<a:uFill>${solidFill(r.uColor)}</a:uFill>` : '';
  const link = r.link && ctx ? ctx.hlink(r.link) : '';
  const sym = r.sym ? `<a:sym typeface="${esc(r.sym)}" pitchFamily="2" charset="2"/>` : '';
  return `<${tag} ${attrs}>${ln}${fill}${eff}${hl}${uf}${fontXml(r.font ?? defaultFont(o))}${sym}${link}</${tag}>`;
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
  const tabs = p.tabs?.length ? `<a:tabLst>${p.tabs.map((t) => `<a:tab pos="${E(t.pos)}" algn="${{ l: 'l', ctr: 'ctr', r: 'r', dec: 'dec' }[t.algn] ?? 'l'}"/>`).join('')}</a:tabLst>` : '';
  return `<a:pPr ${attrs}>${lnSpc}${bef}${aft}${bu}${tabs}</a:pPr>`;
}

/** 필드 id (GUID, 종류마다 고정) */
function fldId(type) {
  let h = 0;
  for (const ch of String(type)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  const x = h.toString(16).toUpperCase().padStart(8, '0');
  return `{${x}-6B2D-4C1E-9A3F-${x}5D7E}`;
}

function txBodyXml(o, body, ctx, tag = 'p:txBody') {
  const ins = body.insets ?? [9.6, 4.8, 9.6, 4.8];
  const anchor = { t: 't', ctr: 'ctr', b: 'b' }[body.anchor ?? 't'];
  const fit = body.autofit === 'resize' ? '<a:spAutoFit/>' : body.autofit === 'shrink' ? `<a:normAutofit${body.fontScale && body.fontScale < 1 ? ` fontScale="${pct(body.fontScale)}"` : ''}${body.lnSpcReduction ? ` lnSpcReduction="${pct(body.lnSpcReduction)}"` : ''}/>` : '<a:noAutofit/>';
  const warp = body.warp ? `<a:prstTxWarp prst="${esc(body.warp.prst)}"><a:avLst>${Object.entries(body.warp.adj ?? {}).map(([k, v]) => `<a:gd name="${esc(k)}" fmla="val ${Math.round(v)}"/>`).join('')}</a:avLst></a:prstTxWarp>` : '';
  const extra = `${body.rot ? ` rot="${Math.round(body.rot * 60000)}"` : ''}${body.cols > 1 ? ` numCol="${Math.min(16, Math.round(body.cols))}"${body.colSpc ? ` spcCol="${E(body.colSpc)}"` : ''}` : ''}${body.upright ? ' upright="1"' : ''}`;
  const bodyPr = `<a:bodyPr${extra} vert="${body.vert ?? 'horz'}" wrap="${body.wrap === false ? 'none' : 'square'}" lIns="${E(ins[0])}" tIns="${E(ins[1])}" rIns="${E(ins[2])}" bIns="${E(ins[3])}" rtlCol="0" anchor="${anchor}" anchorCtr="${body.anchorCtr ? 1 : 0}">${warp}${fit}</a:bodyPr>`;
  const paras = body.paras.map((p) => {
    const runs = p.runs.map((r) => {
      // 필드 (슬라이드 번호 · 날짜 …)
      if (r.fld) return `<a:fld id="${esc(r.fld.id ?? fldId(r.fld.type))}" type="${esc(r.fld.type)}">${rPrXml(r, o, p, ctx)}<a:t>${esc(r.t || (r.fld.type === 'slidenum' ? '‹#›' : ''))}</a:t></a:fld>`;
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

/** 장식용 개체 (PowerPoint [장식용으로 표시]: 화면 읽기 프로그램이 건너뜀) */
const decoXml = (o) => (o.decorative ? '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>' : '');
/** PowerPoint SmartArt 레이아웃 (layoutDef uniqueId 끝) → WIPOINT SMART_KINDS */
const DGM_LAYOUT = {
  default: 'blocks', list1: 'list', vList2: 'vBullet', vList5: 'boxList', vList6: 'boxList', hList1: 'hBullet', hList6: 'hBullet', lProcess2: 'list', trapezoid2: 'trapezoidList', hexagonList: 'hexList', StackedList: 'stackedList',
  process1: 'process', process2: 'vProcess', process4: 'arrowProcess', chevron1: 'chevron', hChevron3: 'chevron', vProcess5: 'vProcess', bProcess3: 'bentProcess', bProcess4: 'bentProcess', StepUpProcess: 'stepUp', hProcess9: 'arrowProcess', funnel1: 'funnel', chevron2: 'vChevron',
  cycle1: 'cycle', cycle2: 'cycle', cycle3: 'textCycle', cycle5: 'continuousCycle', cycle6: 'blockCycle', cycle8: 'segmentedCycle', radial1: 'radialCycle', radial3: 'radialVenn', radial5: 'converging',
  hierarchy1: 'hierarchy', orgChart1: 'hierarchy', hierarchy2: 'hHierarchy', hierarchy4: 'tableHierarchy', NameandTitleOrganizationalChart: 'nameHierarchy',
  venn1: 'venn', venn2: 'stackedVenn', venn3: 'venn', target1: 'targetVenn', target3: 'targetVenn', arrow4: 'counterArrows', arrow1: 'counterArrows', equation1: 'equation', balance1: 'balance',
  matrix1: 'matrix', matrix2: 'titledMatrix', matrix3: 'matrix', cycle4: 'cycleMatrix', pyramid1: 'pyramid', pyramid3: 'invPyramid', pyramid2: 'pyramidList', pyramid4: 'segPyramid',
  pList1: 'picAccent', pList2: 'picGrid', CircularPictureCallout: 'picCircles', TitledPictureBlocks: 'picTitles',
};
/** SmartArt 데이터 모델 (dgm:dataModel) → 글 줄 목록 (수준 = 앞의 탭 수) */
function dgmItems(dm) {
  const pts = new Map();
  for (const pt of descendants(child(dm, 'ptLst'), 'pt')) {
    const type = attr(pt, 'type') ?? 'node';
    pts.set(attr(pt, 'modelId'), { type, t: descendants(child(pt, 't'), 'p').map((p) => descendants(p, 't').map((x) => x.text ?? '').join('')).join(' ').trim(), kids: [] });
  }
  for (const c of descendants(child(dm, 'cxnLst'), 'cxn')) {
    const type = attr(c, 'type') ?? 'parOf';
    if (type !== 'parOf') continue;
    const src = pts.get(attr(c, 'srcId')); const dst = pts.get(attr(c, 'destId'));
    if (src && dst && dst.type === 'node') src.kids.push([Number(attr(c, 'srcOrd') ?? 0), dst]);
  }
  const doc = [...pts.values()].find((p) => p.type === 'doc');
  const out = [];
  const walk = (n, depth) => { for (const [, k] of n.kids.sort((a, b) => a[0] - b[0])) { out.push(`${'\t'.repeat(depth)}${k.t}`); walk(k, depth + 1); } };
  if (doc) walk(doc, 0);
  return out.filter((l) => l.trim());
}

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
  const link = o.link ? ctx.hlink(o.link) : '';
  const isLine = o.shape === 'line' || o.shape === 'straightConnector1';
  if (isLine && !o.text) {
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name}">${link}</p:cNvPr><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrmXml(o)}${geomXml(o)}${lineXml(o.line)}</p:spPr></p:cxnSp>`;
  }
  const effect = fxXml(fxOf(o));
  const spPr = `<p:spPr>${xfrmXml(o)}${geomXml(o)}${o.ph ? (o.fill ? fillXml(o.fill, ctx) : '') : fillXml(o.fill, ctx)}${o.ph && !o.line ? '' : lineXml(o.line)}${effect}</p:spPr>`;
  const text = o.text ? txBodyXml(o, o.text, ctx) : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"${o.alt ? ` descr="${esc(o.alt)}"` : ''}>${link}${decoXml(o)}</p:cNvPr><p:cNvSpPr${o.txBox ? ' txBox="1"' : ''}>${o.ph ? '<a:spLocks noGrp="1"/>' : ''}</p:cNvSpPr>${phXml(o, phIdx)}</p:nvSpPr>${spPr}${text}</p:sp>`;
}

/** OLE 개체 (python-pptx · PowerPoint Fallback 과 같은 꼴: oleObj 안에 미리 보기 p:pic) */
function oleXml(o, id, ctx, phIdx) {
  const ole = o.ole;
  const r = ole.link ? ctx.add(REL('oleObject'), ole.link, true) : ctx.ole(ole.bin, ole.pkg, ole.ext);
  if (!r) return null;
  // 안쪽 그림의 id 는 0 (PowerPoint 가 쓰는 꼴 — 슬라이드의 다른 개체 id 와 겹치지 않음)
  const pic = picXml({ ...o, ole: undefined, ph: null, name: '', link: null, alt: null }, 0, ctx, phIdx);
  const size = ole.imgW ? ` imgW="${Math.round(ole.imgW)}" imgH="${Math.round(ole.imgH)}"` : ` imgW="${E(o.w)}" imgH="${E(o.h)}"`;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name ?? `Object ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(o, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="${esc(ole.name ?? '')}"${ole.icon ? ' showAsIcon="1"' : ''} r:id="${r}"${size} progId="${esc(ole.progId ?? '')}">${ole.link ? '<p:link updateAutomatic="1"/>' : '<p:embed/>'}${pic}</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
}

function picXml(o, id, ctx, phIdx) {
  if (o.ole && (o.ole.bin || o.ole.link) && o.media) { const x = oleXml(o, id, ctx, phIdx); if (x) return x; }
  // SmartArt 그림 칸 (그림 없음): 회색 사각형/타원으로
  if (!o.media && o.picSlot) return shapeXml({ ...o, type: 'shape', shape: o.shape ?? 'rect', fill: { type: 'solid', color: '#E7E6E6' }, line: null, text: null, picSlot: undefined }, id, ctx, phIdx);
  if (!o.media) return o.ph ? `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>${phXml(o, phIdx)}</p:nvSpPr><p:spPr>${xfrmXml(o)}</p:spPr></p:sp>` : '';
  const rid = ctx.image(o.media);
  if (!rid) return '';
  const c = o.crop;
  const src = c && (c.l || c.t || c.r || c.b) ? `<a:srcRect l="${pct(c.l)}" t="${pct(c.t)}" r="${pct(c.r)}" b="${pct(c.b)}"/>` : '';
  const alpha = o.alpha != null ? `<a:alphaModFix amt="${pct(o.alpha)}"/>` : '';
  const gray = `${o.gray ? '<a:grayscl/>' : ''}${o.duotone?.length === 2 ? `<a:duotone>${colorXml(o.duotone[0])}${colorXml(o.duotone[1])}</a:duotone>` : ''}`;
  const lum = o.bright || o.contrast ? `<a:lum bright="${pct(o.bright ?? 0)}" contrast="${pct(o.contrast ?? 0)}"/>` : '';
  const link = o.link ? ctx.hlink(o.link) : '';
  const effect = fxXml(fxOf(o));
  // 꾸밈 효과: 결과 그림이 본 그림, 원본과 효과는 WIPOINT 확장에 (PowerPoint 는 결과 그림을 보여 줌)
  const origRid = o.art && o.artOrig ? ctx.image(o.artOrig) : null;
  const artExt = o.art ? `<a:extLst><a:ext uri="${WP_EXT_URI}"><wp:art xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(o.art))}"${origRid ? ` r:embed="${origRid}"` : ''}/></a:ext></a:extLst>` : '';
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(o.name ?? `Picture ${id}`)}" descr="${esc(o.alt ?? '')}">${link}${decoXml(o)}</p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>${o.ph ? phXml(o, phIdx) : '<p:nvPr/>'}</p:nvPicPr><p:blipFill><a:blip r:embed="${rid}">${alpha}${gray}${lum}${artExt}</a:blip>${src}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(o)}${geomXml({ shape: o.shape ?? 'rect', adj: o.adj })}${o.line ? lineXml(o.line) : ''}${effect}</p:spPr></p:pic>`;
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
      if ((side === 'dd' || side === 'du') && !b) return '';
      if (b) return `<a:${tag} w="${Math.round((b.width ?? 1) * EMU_PER_PX)}">${solidFill(b.color)}</a:${tag}>`;
      return `<a:${tag} w="12700">${solidFill(ts.border)}</a:${tag}>`;
    };
    const merge = `${c.span ? ` gridSpan="${c.span}"` : ''}${c.rowSpan ? ` rowSpan="${c.rowSpan}"` : ''}${c.hMerge ? ' hMerge="1"' : ''}${c.vMerge ? ' vMerge="1"' : ''}`;
    return `<a:tc${merge}>${txBodyXml(pseudo, body, ctx, 'a:txBody').replace(/<a:bodyPr[^>]*>(<a:[a-zA-Z]+[^>]*\/>)?<\/a:bodyPr>/, '<a:bodyPr/>')}<a:tcPr marL="${E(ins[0])}" marR="${E(ins[2])}" marT="${E(ins[1])}" marB="${E(ins[3])}" anchor="${anchor}">${bl('lnL', 'l')}${bl('lnR', 'r')}${bl('lnT', 't')}${bl('lnB', 'b')}${bl('lnTlToBr', 'dd')}${bl('lnBlToTr', 'du')}${fill}</a:tcPr></a:tc>`;
  }).join('')}</a:tr>`).join('');
  const st = o.style ?? {};
  const ext = `<a:extLst><a:ext uri="${WP_EXT_URI}"><wp:table xmlns:wp="${WP_NS}" json="${esc(JSON.stringify({ style: st, own }))}"/></a:ext></a:extLst>`;
  const tblPr = `<a:tblPr${st.firstRow ? ' firstRow="1"' : ''}${st.banded ? ' bandRow="1"' : ''}${st.lastRow ? ' lastRow="1"' : ''}${st.firstCol ? ' firstCol="1"' : ''}>${ext}</a:tblPr>`;
  const grid = o.cols.map((w) => `<a:gridCol w="${E(w)}"/>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name ?? `Table ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(o, 'p:xfrm')}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}<a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function chartFrameXml(o, id, ctx) {
  const rid = ctx.chart(o.chart, o.chartRaw);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name ?? `Chart ${id}`)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(o, 'p:xfrm')}<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

/** 차트 파트 (값은 리터럴로 — 내장 통합 문서 없이도 PowerPoint 가 그림) */
/** 열 번호 → 문자 (0 → A) */
const colName = (i) => { let s = ''; for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s; return s; };
/** 차트 시트의 열: 계열 i 의 값 열 (거품형은 계열마다 값 · 크기 두 열) */
const chartCol = (ch, i, size = false) => (ch.kind === 'bubble' ? 1 + i * 2 + (size ? 1 : 0) : 1 + i);
/**
 * 차트 데이터 통합 문서 (PowerPoint [데이터 편집] 용): A열 = 항목, 1행 = 계열 이름
 * 최소 xlsx (inlineStr 글자, 스타일 1개) 를 zip 으로 만듦
 */
export function chartWorkbook(ch0) {
  const ch = fallbackChart(ch0);
  const cell = (ref, v) => (typeof v === 'number' && Number.isFinite(v) ? `<c r="${ref}"><v>${v}</v></c>` : v == null || v === '' ? '' : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(String(v))}</t></is></c>`);
  const bub = ch.kind === 'bubble';
  const xyNum = ch.kind === 'scatter' || bub;
  const rows = [`<row r="1">${ch.series.map((s, i) => cell(`${colName(chartCol(ch, i))}1`, s.name) + (bub ? cell(`${colName(chartCol(ch, i, true))}1`, `${s.name} 크기`) : '')).join('')}</row>`];
  ch.cats.forEach((c, r) => {
    const num = xyNum && Number.isFinite(Number(c)) ? Number(c) : c;
    rows.push(`<row r="${r + 2}">${cell(`A${r + 2}`, num)}${ch.series.map((s, i) => cell(`${colName(chartCol(ch, i))}${r + 2}`, Number(s.vals[r]) || 0) + (bub ? cell(`${colName(chartCol(ch, i, true))}${r + 2}`, Number(s.sizes?.[r]) || 0) : '')).join('')}</row>`);
  });
  const last = `${colName(chartCol(ch, ch.series.length - 1, bub))}${ch.cats.length + 1}`;
  const enc = new TextEncoder();
  const files = {
    '[Content_Types].xml': `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL('officeDocument')}" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `${XML_HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL('worksheet')}" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL('styles')}" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `${XML_HEAD}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="맑은 고딕"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    'xl/worksheets/sheet1.xml': `${XML_HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${last}"/><sheetData>${rows.join('')}</sheetData></worksheet>`,
  };
  return zip(Object.fromEntries(Object.entries(files).map(([name, text]) => [name, enc.encode(text)])));
}

// 데이터 레이블 위치: 차트 종류마다 PowerPoint 가 허용하는 값만 (다른 값이면 파일이 손상되었다고 함)
const DLBL_POS = { barClustered: ['ctr', 'inEnd', 'inBase', 'outEnd'], barStacked: ['ctr', 'inEnd', 'inBase'], line: ['t', 'b', 'l', 'r', 'ctr'], scatter: ['t', 'b', 'l', 'r', 'ctr'], bubble: ['t', 'b', 'l', 'r', 'ctr'], pie: ['bestFit', 'ctr', 'inEnd', 'outEnd'] };
const TREND_XML = { linear: 'linear', exp: 'exp', log: 'log', poly: 'poly', power: 'power', movingAvg: 'movingAvg' };
const ERR_XML = { fixed: 'fixedVal', percent: 'percentage', stdDev: 'stdDev', stdErr: 'stdErr' };
const MARKER_XML = { circle: 'circle', square: 'square', diamond: 'diamond', triangle: 'triangle', x: 'x', none: 'none' };

/** opts.embed: 데이터 통합 문서 관계 id (있으면 Sheet1 셀 참조 + 캐시로 씀) */
/** c:view3D (+ 벽 · 바닥) — X/Y 회전 · 깊이 · 원근 */
function view3dXml(ch) {
  const v = { ...VIEW3D_DEFAULT, ...(ch.kind === 'pie' ? { rotX: 30, rotY: 0 } : {}), ...(ch.view3d ?? {}) };
  return `<c:view3D><c:rotX val="${Math.round(v.rotX)}"/><c:rotY val="${Math.round(((v.rotY % 360) + 360) % 360)}"/><c:depthPercent val="${Math.max(20, Math.min(2000, Math.round(v.depth)))}"/><c:rAngAx val="${ch.kind === 'pie' ? 0 : 1}"/>${ch.kind === 'pie' ? `<c:perspective val="${Math.round(v.persp ?? 30)}"/>` : ''}</c:view3D>`;
}
export function chartXml(orig, opts = {}) {
  // 폭포 · 깔때기 · 트리맵 등: 다른 프로그램용 일반 차트로 쓰고, 원본은 WIPOINT 확장 JSON 에
  const ch = fallbackChart(orig);
  const kind = ch.kind ?? 'col';
  const n = ch.cats.length;
  const ref = (col, r1, r2) => `Sheet1!$${col}$${r1}${r2 != null ? `:$${col}$${r2}` : ''}`;
  const cache = (vals, num, fm) => `<c:${num ? 'numRef' : 'strRef'}><c:f>${fm}</c:f><c:${num ? 'numCache' : 'strCache'}>${num ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(num ? (Number(v) || 0) : v)}</c:v></c:pt>`).join('')}</c:${num ? 'numCache' : 'strCache'}></c:${num ? 'numRef' : 'strRef'}>`;
  const lit0 = (vals, num) => `<c:${num ? 'numLit' : 'strLit'}>${num ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(num ? (Number(v) || 0) : v)}</c:v></c:pt>`).join('')}</c:${num ? 'numLit' : 'strLit'}>`;
  const pie = kind === 'pie' || kind === 'doughnut';
  const xy = kind === 'scatter' || kind === 'bubble';
  const grp = ch.grouping ?? (ch.stacked ? 'stacked' : null);
  const typeOf = (s, i) => (kind === 'combo' ? s.type ?? (i === ch.series.length - 1 && ch.series.length > 1 ? 'line' : 'col') : kind);
  const grouping = (t) => {
    if (t === 'col' || t === 'bar') return grp === 'stacked' ? 'stacked' : grp === 'percent' ? 'percentStacked' : 'clustered';
    if (kind === 'combo' && t === 'line') return 'standard';
    return grp === 'stacked' ? 'stacked' : grp === 'percent' ? 'percentStacked' : 'standard';
  };
  const pt = (v) => Math.round(v * 12700);
  const lnXml = (s, t) => {
    const w = s.lineWidth ? pt(s.lineWidth) : 28575;
    return `<a:ln w="${w}" cap="rnd">${s.color ? solidFill(s.color) : ''}${s.dash ? `<a:prstDash val="${{ dash: 'dash', dot: 'sysDot', dashDot: 'dashDot', longDash: 'lgDash' }[s.dash] ?? 'solid'}"/>` : ''}<a:round/></a:ln>`;
  };
  const dlXml = (lo, t, stacked, bar) => {
    if (!lo) return '';
    const o = lo === true ? {} : lo;
    const any = o.val || o.cat || o.ser || o.pct;
    const val = any ? !!o.val : !pie;
    const pct = any ? !!o.pct : pie;
    const allowed = pie ? DLBL_POS.pie : bar ? (stacked ? DLBL_POS.barStacked : DLBL_POS.barClustered) : DLBL_POS[t] ?? [];
    const pos = o.pos && allowed.includes(o.pos) ? `<c:dLblPos val="${o.pos}"/>` : '';
    return `<c:dLbls>${pos}<c:showLegendKey val="0"/><c:showVal val="${val ? 1 : 0}"/><c:showCatName val="${o.cat ? 1 : 0}"/><c:showSerName val="${o.ser ? 1 : 0}"/><c:showPercent val="${pct ? 1 : 0}"/><c:showBubbleSize val="0"/>${pie ? '<c:showLeaderLines val="1"/>' : ''}</c:dLbls>`;
  };
  const trendXml = (s) => {
    const tr = s.trend;
    if (!tr?.type || !TREND_XML[tr.type]) return '';
    return `<c:trendline>${tr.name ? `<c:name>${esc(tr.name)}</c:name>` : ''}<c:spPr><a:ln w="19050" cap="rnd">${s.color ? solidFill(s.color) : ''}<a:prstDash val="sysDot"/></a:ln></c:spPr><c:trendlineType val="${TREND_XML[tr.type]}"/>${tr.type === 'poly' ? `<c:order val="${Math.max(2, Math.min(6, tr.order ?? 2))}"/>` : ''}${tr.type === 'movingAvg' ? `<c:period val="${Math.max(2, tr.period ?? 2)}"/>` : ''}${tr.forward ? `<c:forward val="${Number(tr.forward)}"/>` : ''}${tr.backward ? `<c:backward val="${Number(tr.backward)}"/>` : ''}<c:dispRSqr val="${tr.r2 ? 1 : 0}"/><c:dispEq val="${tr.eq ? 1 : 0}"/></c:trendline>`;
  };
  const errXml = (s, t) => (s.err && ERR_XML[s.err.type] ? `<c:errBars>${xy ? '<c:errDir val="y"/>' : ''}<c:errBarType val="both"/><c:errValType val="${ERR_XML[s.err.type]}"/><c:noEndCap val="0"/>${s.err.type !== 'stdErr' ? `<c:val val="${Number(s.err.val ?? (s.err.type === 'percent' ? 5 : 1))}"/>` : ''}</c:errBars>` : '');
  const dPtXml = (s, t) => {
    const out = [];
    const idxs = new Set([...Object.keys(s.pointColors ?? {}), ...Object.keys(s.explodeAt ?? {})].map(Number));
    for (const i of [...idxs].sort((a, b) => a - b)) {
      const col = s.pointColors?.[i];
      const ex = s.explodeAt?.[i];
      out.push(`<c:dPt><c:idx val="${i}"/>${t === 'col' || t === 'bar' || t === 'bubble' ? '<c:invertIfNegative val="0"/>' : ''}${t === 'line' || t === 'scatter' ? '<c:marker><c:symbol val="circle"/></c:marker>' : ''}${t === 'pie' && ex ? `<c:explosion val="${Math.round(ex)}"/>` : ''}<c:bubble3D val="0"/>${col ? `<c:spPr>${solidFill(col)}</c:spPr>` : ''}</c:dPt>`);
    }
    return out.join('');
  };
  const serXml = (s, i, t, stacked) => {
    const col = colName(chartCol(ch, i));
    const tx = opts.embed ? `<c:tx>${cache([s.name], false, ref(col, 1))}</c:tx>` : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
    const lit = (vals, num, which) => (opts.embed ? cache(vals, num, which === 'cat' ? ref('A', 2, n + 1) : ref(which === 'size' ? colName(chartCol(ch, i, true)) : col, 2, n + 1)) : lit0(vals, num));
    const isLine = t === 'line' || t === 'scatter' || t === 'radar';
    const mk = s.marker ?? (t === 'scatter' || (t === 'line' && (ch.variant === 'markers' || ch.markers)) || (t === 'radar' && ch.variant === 'marker') ? 'circle' : 'none');
    const markerXml = isLine ? `<c:marker><c:symbol val="${MARKER_XML[mk] ?? 'circle'}"/>${mk !== 'none' ? `<c:size val="${Math.max(2, Math.min(72, s.markerSize ?? 5))}"/>${s.color ? `<c:spPr>${solidFill(s.color)}<a:ln>${solidFill(s.color)}</a:ln></c:spPr>` : ''}` : ''}</c:marker>` : '';
    const scatterLine = t === 'scatter' && !(ch.variant === 'smooth' || ch.variant === 'line' || s.lineWidth);
    const spPr = s.noFill ? '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' : isLine && !(t === 'radar' && ch.variant === 'filled') ? `<c:spPr>${scatterLine ? '<a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln>' : lnXml(s, t)}</c:spPr>` : s.color ? `<c:spPr>${solidFill(s.color)}${ch.barOutline ? `<a:ln w="9525">${solidFill(ch.barOutline)}</a:ln>` : ''}</c:spPr>` : '';
    const dl = s.noFill ? (ch.labels ? '<c:dLbls><c:delete val="1"/></c:dLbls>' : '') : s.labels != null ? dlXml(s.labels, t, stacked, t === 'col' || t === 'bar') : '';
    const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}`;
    const smooth = `<c:smooth val="${(s.smooth ?? (ch.variant === 'smooth' || ch.smooth)) ? 1 : 0}"/>`;
    if (t === 'scatter') return `<c:ser>${head}${markerXml}${dPtXml(s, t)}${dl}${trendXml(s)}${errXml(s, t)}<c:xVal>${lit(ch.cats.map((c, k) => (Number.isFinite(Number(c)) ? Number(c) : k + 1)), true, 'cat')}</c:xVal><c:yVal>${lit(s.vals, true)}</c:yVal>${smooth}</c:ser>`;
    if (t === 'bubble') return `<c:ser>${head}<c:invertIfNegative val="0"/>${dPtXml(s, t)}${dl}${trendXml(s)}${errXml(s, t)}<c:xVal>${lit(ch.cats.map((c, k) => (Number.isFinite(Number(c)) ? Number(c) : k + 1)), true, 'cat')}</c:xVal><c:yVal>${lit(s.vals, true)}</c:yVal><c:bubbleSize>${lit((s.sizes ?? s.vals.map(() => 1)).map((v) => Number(v) || 0), true, 'size')}</c:bubbleSize><c:bubble3D val="0"/></c:ser>`;
    if (t === 'pie') return `<c:ser>${head}${ch.variant === 'explode' || ch.explode ? `<c:explosion val="${ch.variant === 'explode' ? 18 : Math.round(ch.explode)}"/>` : ''}${dPtXml(s, t)}${dl}<c:cat>${lit(ch.cats, false, 'cat')}</c:cat><c:val>${lit(s.vals, true)}</c:val></c:ser>`;
    if (t === 'radar') return `<c:ser>${head}${markerXml}${dPtXml(s, t)}${dl}<c:cat>${lit(ch.cats, false, 'cat')}</c:cat><c:val>${lit(s.vals, true)}</c:val></c:ser>`;
    const extra = t === 'col' || t === 'bar' ? '<c:invertIfNegative val="0"/>' : t === 'line' ? markerXml : '';
    return `<c:ser>${head}${extra}${dPtXml(s, t)}${dl}${t !== 'area' ? trendXml(s) : ''}${errXml(s, t)}<c:cat>${lit(ch.cats, false, 'cat')}</c:cat><c:val>${lit(s.vals, true)}</c:val>${t === 'line' ? smooth : ''}</c:ser>`;
  };
  // 계열을 (종류, 축) 묶음으로 — 콤보는 여러 묶음
  const groups = [];
  ch.series.forEach((s, i) => {
    let t = typeOf(s, i);
    if (kind === 'doughnut') t = 'pie';
    const axis = !pie && !xy && kind !== 'radar' && s.axis === 'secondary' ? 2 : 1;
    const key = `${t}|${axis}`;
    let g = groups.find((x) => x.key === key);
    if (!g) groups.push(g = { key, t, axis, items: [] });
    g.items.push(i);
  });
  const has2 = groups.some((g) => g.axis === 2);
  const d3 = is3d(ch) && groups.length === 1;
  const AX = { 1: '<c:axId val="111111111"/><c:axId val="222222222"/>', 2: '<c:axId val="333333333"/><c:axId val="444444444"/>' };
  const chartDl = (t, stacked) => (ch.labels ? dlXml(ch.labels, t, stacked, t === 'col' || t === 'bar') : '');
  const plots = groups.map((g) => {
    const t = g.t;
    const gp = grouping(t);
    const stacked = gp !== 'clustered' && gp !== 'standard';
    const sers = g.items.map((i) => serXml(ch.series[i], i, t === 'pie' ? 'pie' : t, stacked)).join('');
    const dl = chartDl(t === 'pie' ? 'pie' : t, stacked);
    const vary = `<c:varyColors val="${pie || ch.vary ? 1 : 0}"/>`;
    // 3차원: bar3DChart · line3DChart · area3DChart · pie3DChart (계열을 깊이로 두면 grouping standard + 계열 축)
    if (d3) {
      const full3 = ch.variant3d === 'full' || ((t === 'line' || t === 'area') && gp === 'standard');
      const ax3 = full3 ? `${AX[1]}<c:axId val="555555555"/>` : AX[1];
      if (t === 'col' || t === 'bar') return `<c:bar3DChart><c:barDir val="${kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${full3 ? 'standard' : gp}"/>${vary}${sers}${dl}<c:gapWidth val="${Math.round(ch.gap ?? 150)}"/><c:shape val="box"/>${ax3}</c:bar3DChart>`;
      if (t === 'line') return `<c:line3DChart><c:grouping val="${gp}"/>${vary}${sers}${dl}${AX[1]}<c:axId val="555555555"/></c:line3DChart>`;
      if (t === 'area') return `<c:area3DChart><c:grouping val="${gp}"/>${vary}${sers}${dl}${ax3}</c:area3DChart>`;
      if (t === 'pie' && kind === 'pie') return `<c:pie3DChart>${vary}${sers}${dl}</c:pie3DChart>`;
    }
    if (t === 'col' || t === 'bar') return `<c:barChart><c:barDir val="${kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${gp}"/>${vary}${sers}${dl}<c:gapWidth val="${Math.round(ch.gap ?? (stacked ? 150 : 219))}"/><c:overlap val="${Math.round(ch.overlap ?? (stacked ? 100 : -27))}"/>${AX[g.axis]}</c:barChart>`;
    if (t === 'line') return `<c:lineChart><c:grouping val="${gp}"/>${vary}${sers}${dl}<c:marker val="1"/>${AX[g.axis]}</c:lineChart>`;
    if (t === 'area') return `<c:areaChart><c:grouping val="${gp}"/>${vary}${sers}${dl}${AX[g.axis]}</c:areaChart>`;
    if (t === 'pie' && kind === 'pie') return `<c:pieChart>${vary}${sers}${dl}<c:firstSliceAng val="${Math.round(((ch.firstAngle ?? 0) % 360 + 360) % 360)}"/></c:pieChart>`;
    if (t === 'pie') return `<c:doughnutChart>${vary}${sers}${dl}<c:firstSliceAng val="${Math.round(((ch.firstAngle ?? 0) % 360 + 360) % 360)}"/><c:holeSize val="${Math.max(10, Math.min(90, Math.round(ch.hole ?? 50)))}"/></c:doughnutChart>`;
    if (t === 'scatter') return `<c:scatterChart><c:scatterStyle val="${ch.variant === 'smooth' ? 'smoothMarker' : 'lineMarker'}"/>${vary}${sers}${dl}${AX[g.axis]}</c:scatterChart>`;
    if (t === 'radar') return `<c:radarChart><c:radarStyle val="${ch.variant === 'filled' ? 'filled' : 'marker'}"/>${vary}${sers}${dl}${AX[g.axis]}</c:radarChart>`;
    return `<c:bubbleChart>${vary}${sers}${dl}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/>${AX[g.axis]}</c:bubbleChart>`;
  }).join('');
  const axes = ch.axes ?? {};
  const grid = ch.grid ?? { major: true };
  const gridLn = (minor) => `<c:${minor ? 'minor' : 'major'}Gridlines><c:spPr><a:ln w="${minor ? 6350 : 9525}">${solidFill(ch.gridColor ?? (minor ? '@tx1:lm5:lo95' : '@tx1:lm15:lo85'))}${ch.dashGrid ? '<a:prstDash val="dash"/>' : ''}</a:ln></c:spPr></c:${minor ? 'minor' : 'major'}Gridlines>`;
  const axTitle = (t, rot) => (t ? `<c:title><c:tx><c:rich><a:bodyPr${rot ? ` rot="${rot}" vert="horz"` : ''}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1330" b="0"/></a:pPr><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '');
  const scaling = (a) => `<c:scaling><c:orientation val="${a?.reverse ? 'maxMin' : 'minMax'}"/>${a?.max != null && a.max !== '' ? `<c:max val="${Number(a.max)}"/>` : ''}${a?.min != null && a.min !== '' ? `<c:min val="${Number(a.min)}"/>` : ''}</c:scaling>`;
  const numFmt = (a) => `<c:numFmt formatCode="${esc(a?.fmt ?? (grp === 'percent' ? '0%' : 'General'))}" sourceLinked="${a?.fmt ? 0 : 1}"/>`;
  const txt = ch.fontColor ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr>${solidFill(ch.fontColor)}</a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>` : '';
  const horiz = kind === 'bar';
  let ax = '';
  if (!pie) {
    const catTag = xy ? 'valAx' : 'catAx';
    const catTail = xy ? `<c:crossBetween val="midCat"/>${axes.x?.major ? `<c:majorUnit val="${Number(axes.x.major)}"/>` : ''}` : '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>';
    const catAx = `<c:${catTag}><c:axId val="111111111"/>${scaling(axes.x)}<c:delete val="${axes.x?.hide || kind === 'radar' && false ? 1 : 0}"/><c:axPos val="${horiz ? 'l' : 'b'}"/>${grid.cat ? gridLn(false) : ''}${axTitle(axes.x?.title, horiz ? '-5400000' : '')}${xy ? numFmt(axes.x) : '<c:numFmt formatCode="General" sourceLinked="1"/>'}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525">${solidFill('@tx1:lm25:lo75')}</a:ln></c:spPr>${txt}<c:crossAx val="222222222"/><c:crosses val="autoZero"/>${catTail}</c:${catTag}>`;
    const valAx = `<c:valAx><c:axId val="222222222"/>${scaling(axes.y)}<c:delete val="${axes.y?.hide ? 1 : 0}"/><c:axPos val="${horiz ? 'b' : 'l'}"/>${grid.major !== false ? gridLn(false) : ''}${grid.minor ? gridLn(true) : ''}${axTitle(axes.y?.title, horiz ? '' : '-5400000')}${numFmt(axes.y)}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txt}<c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="${xy ? 'midCat' : 'between'}"/>${axes.y?.major ? `<c:majorUnit val="${Number(axes.y.major)}"/>` : ''}</c:valAx>`;
    ax = catAx + valAx;
    // 3차원에서 계열 축 (깊이 방향)
    if (d3 && (ch.variant3d === 'full' || kind === 'line' || (kind === 'area' && !grp))) ax += '<c:serAx><c:axId val="555555555"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/></c:serAx>';
    if (has2) ax += `<c:catAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="444444444"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="444444444"/>${scaling(axes.y2)}<c:delete val="${axes.y2?.hide ? 1 : 0}"/><c:axPos val="r"/>${axTitle(axes.y2?.title, '5400000')}${numFmt(axes.y2)}<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txt}<c:crossAx val="333333333"/><c:crosses val="max"/><c:crossBetween val="between"/>${axes.y2?.major ? `<c:majorUnit val="${Number(axes.y2.major)}"/>` : ''}</c:valAx>`;
  }
  const dTable = ch.dataTable && !pie && !xy && kind !== 'radar' ? `<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="${ch.dataTable === 'nokeys' ? 0 : 1}"/></c:dTable>` : '';
  const plotSp = ch.plotBg ? `<c:spPr>${solidFill(ch.plotBg)}</c:spPr>` : '';
  const tsz = ch.titleSize ? Math.round(ch.titleSize * 100) : 1862;
  const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${tsz}" b="${ch.titleBold ? 1 : 0}"/></a:pPr><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>${esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="${ch.titleOverlay ? 1 : 0}"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.legend ? `<c:legend><c:legendPos val="${['t', 'b', 'l', 'r', 'tr'].includes(ch.legend) ? ch.legend : 'b'}"/><c:overlay val="0"/></c:legend>` : '';
  const chartSp = opts.bgRid ? `<c:spPr><a:blipFill><a:blip r:embed="${opts.bgRid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>${ch.border ? `<a:ln w="9525">${solidFill(ch.border)}</a:ln>` : '<a:ln><a:noFill/></a:ln>'}</c:spPr>` : ch.bg || ch.border ? `<c:spPr>${ch.bg ? solidFill(ch.bg) : '<a:noFill/>'}${ch.border ? `<a:ln w="9525">${solidFill(ch.border)}</a:ln>` : '<a:ln><a:noFill/></a:ln>'}</c:spPr>` : '';
  const fsz = ch.font?.size ? Math.round(ch.font.size * 100) : 1197;
  const wx = `<c:extLst><c:ext uri="${WP_EXT_URI}"><wp:chart xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(orig))}"/></c:ext></c:extLst>`;
  return `${XML_HEAD}<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>${title}${d3 ? view3dXml(ch) : ''}<c:plotArea><c:layout/>${plots}${ax}${dTable}${plotSp}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${chartSp}<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${fsz}">${solidFill(ch.fontColor ?? ch.font?.color ?? '@tx1:lm65:lo35')}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>${opts.embed ? `<c:externalData r:id="${opts.embed}"><c:autoUpdate val="0"/></c:externalData>` : ''}${wx}</c:chartSpace>`;
}

// ───────────── 애니메이션 → p:timing ─────────────
const ENTR_PRESET = { appear: 1, fly: 2, blinds: 3, box: 4, checker: 5, shape: 6, circleIn: 6, crawl: 7, diamond: 8, dissolveIn: 9, fade: 10, peek: 12, plus: 13, randomBars: 14, spiralIn: 15, split: 16, stretch: 17, strips: 18, swivel: 19, wedge: 20, wheel: 21, wipe: 22, basicZoom: 23, boomerang: 25, bounce: 26, dropIn: 26, credits: 28, grow: 31, pinwheel: 35, riseUp: 37, float: 42, centerRevolve: 43, sinkDown: 47, compress: 50, curveUp: 52, zoom: 53, expand: 55, flip: 56 };
const EMPH_PRESET = { fillColor: 1, growShrink: 6, lineColor: 7, spin: 8, transparency: 9, boldFlash: 10, bold: 15, brushColor: 16, underline: 18, objectColor: 19, complementary: 21, darken: 24, desaturate: 25, pulse: 26, colorPulse: 27, lighten: 30, teeter: 32, wave: 34, flash: 35, shimmer: 36 };
const EXIT_PRESET = { ...ENTR_PRESET, disappear: 1 };
const DIR_SUB = { b: 4, t: 1, l: 8, r: 2 };
const PATH_PRESET = { pathRight: 64, pathLeft: 64, pathUp: 64, pathDown: 64, pathArc: 37, pathCircle: 1, pathWave: 61 };

function timingXml(slide, spidOf, sound = () => null) {
  const steps = animSteps(slide);
  const trigs = [...triggerSteps(slide)].filter(([t]) => spidOf(t));
  if (!steps.length && !trigs.length) return '';
  let n = 3;
  const nid = () => n++;
  const tgt = (spid) => { const [id, pi] = String(spid).split('#'); return pi != null ? `<p:tgtEl><p:spTgt spid="${id}"><p:txEl><p:pRg st="${pi}" end="${pi}"/></p:txEl></p:spTgt></p:tgtEl>` : `<p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>`; };
  const set = (spid, val, delay = 0) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
  const effect = (spid, dur, filter, dir = 'in') => `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
  const anim = (spid, dur, attr, from, to) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  const flyFrom = { b: ['ppt_y', '1+#ppt_h/2'], t: ['ppt_y', '0-#ppt_h/2'], l: ['ppt_x', '0-#ppt_w/2'], r: ['ppt_x', '1+#ppt_w/2'] };
  const WIPE = { b: 'up', t: 'down', l: 'right', r: 'left' };
  // 모양 필터로 표현되는 나타내기/끝내기 (PowerPoint animEffect filter)
  const FILTER_OF = (e, dir, io) => {
    const hv = dir === 'l' || dir === 'r' ? 'vertical' : 'horizontal';
    switch (e) {
      case 'blinds': return `blinds(${hv})`;
      case 'box': return `box(${io})`;
      case 'checker': return 'checkerboard(across)';
      case 'shape': case 'circleIn': return `circle(${io})`;
      case 'diamond': return `diamond(${io})`;
      case 'plus': return `plus(${io})`;
      case 'dissolveIn': return 'dissolve';
      case 'randomBars': return `randombar(${hv})`;
      case 'strips': return 'strips(downRight)';
      case 'wedge': return 'wedge';
      case 'peek': return `slide(${{ b: 'fromBottom', t: 'fromTop', l: 'fromLeft', r: 'fromRight' }[dir] ?? 'fromBottom'})`;
      default: return null;
    }
  };
  // 크기 · 위치 · 회전으로 표현되는 효과 (out: 끝내기는 거꾸로)
  const MOVE_OF = {
    expand: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0.70*#ppt_w', out ? '0.70*#ppt_w' : '#ppt_w'),
    compress: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '1.5*#ppt_w', out ? '1.5*#ppt_w' : '#ppt_w'),
    stretch: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0', out ? '0' : '#ppt_w'),
    basicZoom: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0', out ? '0' : '#ppt_w') + anim(sp, d, 'ppt_h', out ? '#ppt_h' : '0', out ? '0' : '#ppt_h'),
    pinwheel: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0', out ? '0' : '#ppt_w') + anim(sp, d, 'ppt_h', out ? '#ppt_h' : '0', out ? '0' : '#ppt_h') + anim(sp, d, 'style.rotation', out ? '0' : '720', out ? '720' : '0'),
    swivel: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0', out ? '0' : '#ppt_w'),
    flip: (sp, d, out) => anim(sp, d, 'ppt_h', out ? '#ppt_h' : '0', out ? '0' : '#ppt_h'),
    riseUp: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '#ppt_y+.1', out ? '#ppt_y+.1' : '#ppt_y'),
    sinkDown: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '#ppt_y-.1', out ? '#ppt_y-.1' : '#ppt_y'),
    crawl: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '1+#ppt_h/2', out ? '1+#ppt_h/2' : '#ppt_y'),
    credits: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '1+#ppt_h/2', out ? '0-#ppt_h/2' : '#ppt_y'),
    dropIn: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '0-#ppt_h/2', out ? '1+#ppt_h/2' : '#ppt_y'),
    curveUp: (sp, d, out) => anim(sp, d, 'ppt_x', out ? '#ppt_x' : '#ppt_x-.15', out ? '#ppt_x+.15' : '#ppt_x') + anim(sp, d, 'ppt_y', out ? '#ppt_y' : '#ppt_y+.2', out ? '#ppt_y+.2' : '#ppt_y'),
    boomerang: (sp, d, out) => anim(sp, d, 'ppt_x', out ? '#ppt_x' : '#ppt_x+.25', out ? '#ppt_x+.25' : '#ppt_x') + anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0.3*#ppt_w', out ? '0.3*#ppt_w' : '#ppt_w'),
    spiralIn: (sp, d, out) => anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0', out ? '0' : '#ppt_w') + anim(sp, d, 'ppt_h', out ? '#ppt_h' : '0', out ? '0' : '#ppt_h') + anim(sp, d, 'ppt_x', out ? '#ppt_x' : '0', out ? '1' : '#ppt_x'),
    centerRevolve: (sp, d, out) => anim(sp, d, 'ppt_y', out ? '#ppt_y' : '#ppt_y+.15', out ? '#ppt_y+.15' : '#ppt_y') + anim(sp, d, 'ppt_w', out ? '#ppt_w' : '0.3*#ppt_w', out ? '0.3*#ppt_w' : '#ppt_w'),
  };
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
        default: {
          const f = FILTER_OF(a.effect, dir, 'in');
          if (f) return set(spid, 'visible') + effect(spid, dur, f);
          const m = MOVE_OF[a.effect];
          return set(spid, 'visible') + effect(spid, dur, 'fade') + (m ? m(spid, dur, false) : '');
        }
      }
    }
    if (a.cls === 'exit') {
      const hide = set(spid, 'hidden', Math.max(0, dur - 1));
      switch (a.effect) {
        case 'disappear': return set(spid, 'hidden');
        case 'fly': { const [attr, to] = flyFrom[dir] ?? flyFrom.b; return anim(spid, dur, attr, `#${attr}`, to) + hide; }
        case 'wipe': return effect(spid, dur, `wipe(${WIPE[dir] ?? 'down'})`, 'out') + hide;
        case 'zoom': return anim(spid, dur, 'ppt_w', '#ppt_w', '0') + anim(spid, dur, 'ppt_h', '#ppt_h', '0') + effect(spid, dur, 'fade', 'out') + hide;
        default: {
          const f = FILTER_OF(a.effect, dir, 'out');
          if (f) return effect(spid, dur, f, 'out') + hide;
          const m = MOVE_OF[a.effect];
          return effect(spid, dur, 'fade', 'out') + (m ? m(spid, dur, true) : '') + hide;
        }
      }
    }
    if (a.cls === 'path') {
      const path = `${String(a.motion ?? 'M 0 0 L 0.25 0').trim()} E`;
      return `<p:animMotion origin="layout" path="${esc(path)}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="0" y="0"/></p:animMotion>`;
    }
    // 강조
    const clr = (attrName, inner) => `<p:animClr clrSpc="${inner.startsWith('<p:by><p:hsl') ? 'hsl' : 'rgb'}" dir="cw"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst></p:cBhvr>${inner}</p:animClr>`;
    const setAttr = (attrName, val) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`;
    switch (a.effect) {
      case 'transparency': return setAttr('style.opacity', '0.5');
      case 'bold': return setAttr('style.fontWeight', 'bold');
      case 'underline': return setAttr('style.textDecorationUnderline', 'true');
      case 'desaturate': return clr('fillcolor', '<p:by><p:hsl h="0" s="-100000" l="0"/></p:by>');
      case 'darken': return clr('fillcolor', '<p:by><p:hsl h="0" s="0" l="-30000"/></p:by>');
      case 'lighten': return clr('fillcolor', '<p:by><p:hsl h="0" s="0" l="30000"/></p:by>');
      case 'complementary': return clr('fillcolor', '<p:by><p:hsl h="10800000" s="0" l="0"/></p:by>');
      case 'objectColor': case 'fillColor': case 'brushColor': return clr('fillcolor', '<p:to><a:srgbClr val="FF0000"/></p:to>');
      case 'lineColor': return clr('stroke.color', '<p:to><a:srgbClr val="FF0000"/></p:to>');
      default: break;
    }
    if (a.effect === 'spin' || a.effect === 'teeter') return `<p:animRot by="${a.effect === 'spin' ? 21600000 : 300000}"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${a.effect === 'teeter' ? ' autoRev="1" repeatCount="2000"' : ''}/>${tgt(spid)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    if (a.effect === 'flash') return `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" autoRev="1"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
    const by = a.effect === 'growShrink' ? 150000 : 110000;
    return `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"${a.effect === 'pulse' ? ' autoRev="1"' : ''}/>${tgt(spid)}</p:cBhvr><p:by x="${by}" y="${by}"/></p:animScale>`;
  };
  const seqPars = (steps, interactive) => steps.map((step, si) => {
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
        const spid0 = spidOf(a.obj);
        if (!spid0) return '';
        const spid = a.para != null ? `${spid0}#${a.para}` : spid0;
        const table = a.cls === 'entr' ? ENTR_PRESET : a.cls === 'exit' ? EXIT_PRESET : a.cls === 'path' ? PATH_PRESET : EMPH_PRESET;
        const node = k === 0 && g === groups[0] && step[0] === a ? (a.start === 'after' ? 'afterEffect' : a.start === 'with' ? 'withEffect' : 'clickEffect') : a.start === 'after' ? 'afterEffect' : 'withEffect';
        const sub = a.cls === 'emph' ? 0 : ['fly', 'wipe', 'peek', 'crawl'].includes(a.effect) ? DIR_SUB[a.dir ?? 'b'] ?? 4 : a.effect === 'zoom' || a.effect === 'box' || a.effect === 'shape' || a.effect === 'circleIn' || a.effect === 'diamond' || a.effect === 'plus' ? 16 : a.effect === 'split' ? 37 : a.effect === 'blinds' || a.effect === 'randomBars' ? (a.dir === 'l' || a.dir === 'r' ? 5 : 10) : a.effect === 'wheel' ? 1 : a.effect === 'strips' ? 12 : 0;
        const eid = nid();
        // 텍스트 애니메이션: 단어별 · 문자별 (p:iterate, 지연 = 효과 시간의 %)
        const iter = a.textUnit && a.para == null ? `<p:iterate type="${a.textUnit === 'letter' ? 'lt' : 'wd'}"><p:tmPct val="${Math.round((a.unitDelay ?? 10) * 1000)}"/></p:iterate>` : '';
        const kids = behaviors(a, spid);
        const snd = a.sound ? sound(a.sound) : null;
        const sndXml = snd ? `<p:subTnLst><p:audio><p:cMediaNode vol="80000" showWhenStopped="0"><p:cTn id="${nid()}" display="0" masterRel="sameClick"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:sndTgt r:embed="${snd.rid}" name="${snd.name}"/></p:tgtEl></p:cMediaNode></p:audio></p:subTnLst>` : '';
        return `<p:par><p:cTn id="${eid}" presetID="${table[a.effect] ?? (a.cls === 'path' ? 0 : 10)}" presetClass="${a.cls}" presetSubtype="${sub}" fill="hold" grpId="0" nodeType="${node}"><p:stCondLst><p:cond delay="${Math.round((a.delay ?? 0) * 1000)}"/></p:stCondLst>${iter}<p:childTnLst>${kids}</p:childTnLst>${sndXml}</p:cTn></p:par>`;
      }).join('');
      return `<p:par><p:cTn id="${gid}" fill="hold"><p:stCondLst><p:cond delay="${g.at}"/></p:stCondLst><p:childTnLst>${effs}</p:childTnLst></p:cTn></p:par>`;
    }).join('');
    const first = step[0];
    const cond = interactive ? (si === 0 ? '<p:cond delay="0"/>' : '<p:cond delay="indefinite"/>') : first.start === 'click' ? '<p:cond delay="indefinite"/>' : '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>';
    return `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
  const clickPars = seqPars(steps, false);
  // 트리거: 개체를 누르면 시작하는 interactiveSeq (PowerPoint [트리거 › 클릭할 때])
  const trigSeqs = trigs.map(([t, tsteps]) => {
    const on = `<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${spidOf(t)}"/></p:tgtEl></p:cond>`;
    const id = nid();
    return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst>${on}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>${seqPars(tsteps, true)}</p:childTnLst></p:cTn><p:nextCondLst>${on}</p:nextCondLst></p:seq>`;
  }).join('');
  const allSteps = [...steps, ...trigs.flatMap(([, t]) => t)];
  const byPara = new Set(allSteps.flat().filter((a) => a.para != null).map((a) => spidOf(a.obj)));
  const blds = [...new Set(allSteps.flat().filter((a) => a.cls !== 'path' || byPara.has(spidOf(a.obj))).map((a) => spidOf(a.obj)).filter(Boolean))].map((spid) => `<p:bldP spid="${spid}" grpId="0"${byPara.has(spid) ? ' build="p"' : ''}/>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${clickPars ? `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clickPars}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` : ''}${trigSeqs}</p:childTnLst></p:cTn></p:par></p:tnLst>${blds ? `<p:bldLst>${blds}</p:bldLst>` : ''}</p:timing>`;
}

const PUSH_DIR = { b: 'u', t: 'd', l: 'r', r: 'l' };
const ORIENT = (d) => (d === 'l' || d === 'r' ? 'vert' : 'horz');
const STRIP_DIR = { b: 'ru', t: 'rd', l: 'ld', r: 'lu' };
const LR = (d) => (d === 'l' ? 'r' : 'l');
/** 기본 (ECMA) 전환 요소 */
const TRANS_XML = {
  cut: () => '<p:cut/>', fade: () => '<p:fade/>', dissolve: () => '<p:dissolve/>', zoom: () => '<p:zoom/>', circle: () => '<p:circle/>',
  push: (d) => `<p:push dir="${PUSH_DIR[d] ?? 'u'}"/>`, wipe: (d) => `<p:wipe dir="${PUSH_DIR[d] ?? 'r'}"/>`,
  cover: (d) => `<p:cover dir="${PUSH_DIR[d] ?? 'l'}"/>`, uncover: (d) => `<p:pull dir="${PUSH_DIR[d] ?? 'l'}"/>`,
  split: (d) => `<p:split orient="${d === 'l' ? 'horz' : 'vert'}" dir="out"/>`,
  blinds: (d) => `<p:blinds dir="${ORIENT(d)}"/>`, checker: (d) => `<p:checker dir="${ORIENT(d)}"/>`, comb: (d) => `<p:comb dir="${ORIENT(d)}"/>`, randomBars: (d) => `<p:randomBar dir="${ORIENT(d)}"/>`,
  strips: (d) => `<p:strips dir="${STRIP_DIR[d] ?? 'ru'}"/>`, diamond: () => '<p:diamond/>', plus: () => '<p:plus/>', clock: () => '<p:wheel spokes="1"/>', random: () => '<p:random/>',
};
/** PowerPoint 2010 (p14) 전환 */
const TRANS_P14 = {
  vortex: (d) => `<p14:vortex dir="${LR(d)}"/>`, switch: (d) => `<p14:switch dir="${LR(d)}"/>`, flip: (d) => `<p14:flip dir="${LR(d)}"/>`, ripple: () => '<p14:ripple/>', honeycomb: () => '<p14:honeycomb/>',
  cube: (d) => `<p14:prism dir="${LR(d)}"/>`, box: (d) => `<p14:prism dir="${LR(d)}" isInverted="1"/>`, rotate: (d) => `<p14:prism dir="${LR(d)}" isContent="1"/>`, orbit: (d) => `<p14:prism dir="${LR(d)}" isContent="1" isInverted="1"/>`,
  doors: () => '<p14:doors dir="vert"/>', window: () => '<p14:window dir="vert"/>', ferris: (d) => `<p14:ferris dir="${LR(d)}"/>`, gallery: (d) => `<p14:gallery dir="${LR(d)}"/>`, conveyor: (d) => `<p14:conveyor dir="${LR(d)}"/>`,
  pan: (d) => `<p14:pan dir="${PUSH_DIR[d] ?? 'u'}"/>`, glitter: () => '<p14:glitter dir="l" pattern="hexagon"/>', flythrough: () => '<p14:flythrough/>', flash: () => '<p14:flash/>', shred: () => '<p14:shred/>', reveal: () => '<p14:reveal/>',
};
/** PowerPoint 2013 (p15:prstTrans) 전환 */
const TRANS_P15 = { fall: 'fallOver', drape: 'drape', curtains: 'curtains', wind: 'wind', prestige: 'prestige', fracture: 'fracture', crush: 'crush', peelOff: 'peelOff', pageCurl: 'pageCurlDouble', airplane: 'airplane', origami: 'origami' };
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
function transitionXml(t, sound = () => null) {
  if (!t || t.type === 'none') return t?.advAfter != null ? `<p:transition advTm="${Math.round(t.advAfter * 1000)}"${t.advClick === false ? ' advClick="0"' : ''}/>` : '';
  const dur = t.dur ?? 1;
  const spd = dur <= 0.5 ? 'fast' : dur <= 0.75 ? 'med' : 'slow';
  const adv = `${t.advClick === false ? ' advClick="0"' : ''}${t.advAfter != null ? ` advTm="${Math.round(t.advAfter * 1000)}"` : ''}`;
  const ms = Math.round(dur * 1000);
  const snd = t.sound ? sound(t.sound) : null;
  const sndAc = snd ? `<p:sndAc><p:stSnd><p:snd r:embed="${snd.rid}" name="${snd.name}"/></p:stSnd></p:sndAc>` : '';
  // 새 버전 전환은 mc:AlternateContent — 이전 버전 PowerPoint · LibreOffice 는 페이드
  const alt = (req, ns, inner, fallback = '<p:fade/>') => `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:${req}="${ns}" Requires="${req}"><p:transition spd="${spd}"${adv} xmlns:p14="${NS_P14}" p14:dur="${ms}">${inner}${sndAc}</p:transition></mc:Choice><mc:Fallback><p:transition spd="${spd}"${adv}>${fallback}${sndAc}</p:transition></mc:Fallback></mc:AlternateContent>`;
  if (t.type === 'morph') return alt('p159', 'http://schemas.microsoft.com/office/powerpoint/2015/09/main', '<p159:morph option="byObject"/>');
  if (TRANS_P14[t.type]) return alt('p14', NS_P14, TRANS_P14[t.type](t.dir));
  if (TRANS_P15[t.type]) return alt('p15', NS_P15, `<p15:prstTrans prst="${TRANS_P15[t.type]}"${t.dir === 'l' ? ' invX="1"' : ''}/>`);
  const inner = (TRANS_XML[t.type] ?? TRANS_XML.fade)(t.dir);
  // 빠르게/보통/느리게 와 다른 시간은 p14:dur 로 (PowerPoint 2010+ 과 같음)
  if (![0.5, 0.75, 1].includes(dur)) return alt('p14', NS_P14, inner, inner);
  return `<p:transition spd="${spd}"${adv}>${inner}${sndAc}</p:transition>`;
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

/** 원래 테마 XML 의 색 구성 · 글꼴만 바꿈 */
function patchTheme(xml, theme) {
  const ours = themeXml(theme);
  const cs = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(ours)?.[0];
  let out = cs ? xml.replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, cs) : xml;
  const setFont = (tag, face) => {
    out = out.replace(new RegExp(`(<a:${tag}>[\\s\\S]*?<\\/a:${tag}>)`), (blk) => blk
      .replace(/(<a:latin\b[^>]*?typeface=")[^"]*"/, `$1${esc(face)}"`)
      .replace(/(<a:ea\b[^>]*?typeface=")[^"]*"/, `$1${esc(face)}"`)
      .replace(/(<a:font\s+script="Hang"\s+typeface=")[^"]*"/, `$1${esc(face)}"`));
  };
  if (theme.fonts?.major) setFont('majorFont', theme.fonts.major);
  if (theme.fonts?.minor) setFont('minorFont', theme.fonts.minor);
  return out.startsWith('<?xml') ? out : XML_HEAD + out;
}

function lvlStyles(sizes, bullets, extra = '') {
  return sizes.map((sz, i) => {
    const marL = bullets ? 228600 + i * 457200 : i * 457200;
    const bu = bullets ? `<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/>` : '<a:buNone/>';
    return `<a:lvl${i + 1}pPr marL="${marL}" indent="${bullets ? -228600 : 0}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${i ? 500 : 1000}"/></a:spcBef>${bu}<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`;
  }).join('') + extra;
}

function masterXml(pres, layoutRids, idBase = 2147483649) {
  const size = pres.size;
  const sx = size.w / 1280;
  const sy = size.h / 720;
  const ph = (id, name, type, x, y, w, h, idx) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}"${idx ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x * sx)}" y="${E(y * sy)}"/><a:ext cx="${E(w * sx)}" cy="${E(h * sy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="${type === 'title' ? 'ctr' : 't'}"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"/><a:t>${type === 'title' ? '마스터 제목 스타일 편집' : '마스터 텍스트 스타일을 편집합니다'}</a:t></a:r><a:endParaRPr lang="ko-KR"/></a:p></p:txBody></p:sp>`;
  const sizes = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800];
  const other = [1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800].map((sz, i) => `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`).join('');
  const ids = layoutRids.map((rid, i) => `<p:sldLayoutId id="${idBase + i}" r:id="${rid}"/>`).join('');
  return `${XML_HEAD}<p:sldMaster xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${ph(2, 'Title Placeholder 1', 'title', 88, 38, 1104, 139)}${ph(3, 'Text Placeholder 2', 'body', 88, 192, 1104, 456, 1)}</p:spTree></p:cSld><p:clrMap ${clrMapAttrs(pres.clrMap)}/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>${lvlStyles(sizes, true)}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr>${other}</p:otherStyle></p:txStyles></p:sldMaster>`;
}

const LAYOUT_TYPE = { title: 'title', titleContent: 'obj', section: 'secHead', twoContent: 'twoObj', comparison: 'twoTxTwoObj', titleOnly: 'titleOnly', blank: 'blank', contentCaption: 'objTx', pictureCaption: 'picTx' };
const LAYOUT_NAME = { title: '제목 슬라이드', titleContent: '제목 및 내용', section: '구역 머리글', twoContent: '콘텐츠 2개', comparison: '비교', titleOnly: '제목만', blank: '빈 화면', contentCaption: '캡션 있는 콘텐츠', pictureCaption: '캡션 있는 그림' };

/** 슬라이드 개체 틀 → (종류, idx) 배정. 레이아웃과 슬라이드가 같은 규칙을 씀 */
function phAssign(objs) {
  const map = new Map();
  const count = {};
  for (const o of objs) {
    if (!o.ph) continue;
    const type = o.phFor && o.ph !== 'pic' ? o.phFor : o.ph === 'obj' ? 'body' : o.ph;
    count[type] = (count[type] ?? 0) + 1;
    const key = `${type}#${count[type]}`;
    const idx = type === 'title' || type === 'ctrTitle' ? 0 : ({ body: 0, subTitle: 100, pic: 20, tbl: 30, chart: 40, dgm: 60, media: 70 }[type] ?? 50) + count[type];
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

/** 비디오 · 오디오 (PowerPoint 2010+ 형식: p:pic + a:videoFile/audioFile + p14:media) */
function mediaXml(o, id, ctx) {
  const kind = o.kind === 'audio' ? 'audio' : 'video';
  const r = ctx.av(o.media, kind);
  if (!r) return '';
  const poster = o.poster ? ctx.image(o.poster) : null;
  const blip = poster ? `<a:blip r:embed="${poster}"/>` : '<a:blip/>';
  const trim = o.trim && (o.trim.st || o.trim.end) ? `<p14:trim${o.trim.st ? ` st="${Math.round(o.trim.st * 1000)}"` : ''}${o.trim.end ? ` end="${Math.round(o.trim.end * 1000)}"` : ''}/>` : '';
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(o.name ?? `${kind === 'audio' ? 'Audio' : 'Video'} ${id}`)}" descr="${esc(o.alt ?? '')}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${kind}File r:link="${r.link}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${r.embed}">${trim}</p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

/** 수식: PowerPoint 2010+ 는 a14:m (OMML), 이전 버전 · 다른 프로그램은 LaTeX 글 상자 */
function equationXml(o, id, ctx) {
  const tb = { ...o, type: 'shape', shape: 'rect', fill: null, line: null, txBox: true, name: o.name ?? `수식 ${id}`, text: textBody([{ ...para(), align: 'ctr', runs: [{ t: o.latex ?? '', size: o.size ?? 28, ...(o.color ? { color: o.color } : {}) }] }], { anchor: 'ctr', wrap: false }) };
  const fallback = shapeXml(tb, id, ctx);
  const color = o.color && /^#/.test(o.color) ? o.color : null;
  const math = `<a:p><a:pPr algn="ctr"/><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">${toOMML(parseLatex(o.latex ?? ''), { size: o.size ?? 28, color })}</a14:m><a:endParaRPr lang="en-US" sz="${Math.round((o.size ?? 28) * 100)}"/></a:p>`;
  const choice = fallback.replace(/<a:p>[\s\S]*<\/a:p>/, math);
  return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}

function objXml(o, id, ctx, phIdx, theme) {
  if (o.type === 'equation') return equationXml(o, id, ctx);
  if (o.type === 'image') return picXml(o, id, ctx, phIdx);
  if (o.type === 'media') return mediaXml(o, id, ctx);
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
  return { bytes, ext: extOfMime(mime), mime };
}
const extOfMime = (mime) => OLE_EXT[mime] ?? ({ 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/svg+xml': 'svg', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf', 'image/tiff': 'tiff', ...AV_EXT }[mime] ?? 'png');
/** 비디오 · 오디오 (mime → 확장자) */
const AV_EXT = { 'video/mp4': 'mp4', 'video/webm': 'webm', 'video/quicktime': 'mov', 'video/x-m4v': 'm4v', 'video/ogg': 'ogv', 'audio/mpeg': 'mp3', 'audio/mp3': 'mp3', 'audio/mp4': 'm4a', 'audio/x-m4a': 'm4a', 'audio/aac': 'aac', 'audio/wav': 'wav', 'audio/x-wav': 'wav', 'audio/wave': 'wav', 'audio/ogg': 'ogg', 'audio/webm': 'weba' };
export const AV_MIME = { mp4: 'video/mp4', m4v: 'video/x-m4v', webm: 'video/webm', mov: 'video/quicktime', ogv: 'video/ogg', mp3: 'audio/mpeg', m4a: 'audio/mp4', aac: 'audio/aac', wav: 'audio/wav', ogg: 'audio/ogg', weba: 'audio/webm', wma: 'audio/x-ms-wma', wmv: 'video/x-ms-wmv', avi: 'video/x-msvideo' };
/** OLE 포함 파일 (확장자 → mime) */
const OLE_MIME = { bin: 'application/vnd.openxmlformats-officedocument.oleObject', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', xls: 'application/vnd.ms-excel', doc: 'application/msword', ppt: 'application/vnd.ms-powerpoint', sldx: 'application/vnd.openxmlformats-officedocument.presentationml.slide' };
const OLE_EXT = Object.fromEntries(Object.entries(OLE_MIME).map(([k, v]) => [v, k]));
export const MIME_BY_EXT = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf', tif: 'image/tiff', tiff: 'image/tiff' };

/** 원래 XML 의 관계 id 를 새 파일의 관계로 바꿈 (못 바꾸면 null) */
function remapRels(xml, relMap, rels, pres) {
  const map = {};
  for (const [old, r] of Object.entries(relMap ?? {})) {
    let nid = null;
    if (r.kind === 'image') nid = rels.image(r.media);
    else if (r.kind === 'audio') nid = rels.audio(r.media);
    else if (r.kind === 'link') nid = r.external === false ? rels.add(REL('hyperlink'), r.url) : rels.link(r.url);
    else if (r.kind === 'slide') { const k = pres.slides.findIndex((x) => x.id === r.slide); if (k >= 0) nid = rels.add(REL('slide'), `slide${k + 1}.xml`); }
    if (!nid) return null;
    map[old] = nid;
  }
  return xml.replace(/\br:(\w+)="([^"]*)"/g, (m, k, v) => (v && map[v] ? `r:${k}="${map[v]}"` : m));
}
function rawPartXml(raw, rels) { return remapRels(raw.xml, raw.rels, rels, null); }
/** 고치지 않은 개체 → 원래 XML (위치 · 크기만 지금 값으로) */
function rawXml(o, members, rels, pres) {
  let xml = remapRels(o.raw.xml, o.raw.rels, rels, pres);
  if (!xml) return null;
  const m = /<(a|p):xfrm\b[^>]*?(?:\/>|>[\s\S]*?<\/\1:xfrm>)/.exec(xml);
  if (!m) return null;
  if (o.raw.group) {
    // 묶음 전체를 옮기거나 늘림: 묶음 xfrm 의 off · ext 만 비율대로, chOff · chExt 는 그대로
    const b0 = o.raw.bbox; const b1 = bboxOf(members); const g = o.raw.g;
    const sx = b1.w / b0.w; const sy = b1.h / b0.h;
    const nx = b1.x + (g.x - b0.x) * sx; const ny = b1.y + (g.y - b0.y) * sy;
    const patched = m[0].replace(/<a:off\b[^>]*\/>/, `<a:off x="${E(nx)}" y="${E(ny)}"/>`).replace(/<a:ext\b[^>]*\/>/, `<a:ext cx="${E(g.w * sx)}" cy="${E(g.h * sy)}"/>`);
    return xml.slice(0, m.index) + patched + xml.slice(m.index + m[0].length);
  }
  return xml.slice(0, m.index) + xfrmXml(o, `${m[1]}:xfrm`) + xml.slice(m.index + m[0].length);
}

/**
 * 문서 → pptx 파트 목록
 * opts.png: { mediaId: dataURL } — SVG 그림의 PNG 대체본 (앱이 미리 그려 넘김, PowerPoint 2016 이전 호환)
 * opts.template: true → 서식 파일(.potx) 내용 형식
 * opts.blob(url) → Blob — blob: URL 미디어 (큰 동영상 등; 파트 값이 Blob 이 되므로 zipAsync 로 묶어야 함)
 */
export function pptxEntries(pres, opts = {}) {
  setMasterText(pres); // 마스터 글자 서식을 기본값으로 (글자마다 명시적으로 씀)
  directSlots = !isStdMap(pres.clrMap) || (pres.designs ?? []).some((d) => !isStdMap(d.clrMap)) || pres.slides.some((s) => s.clrMap && !isStdMap(s.clrMap));
  const files = {};
  const ct = new Map();
  const defaults = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
  const override = (path, type) => ct.set(`/${path}`, type);
  const mediaPath = new Map();
  let mediaN = 0;
  const putMedia = (id, dir = 'ppt/media', stem = null) => {
    if (mediaPath.has(id)) return mediaPath.get(id);
    let url = pres.media?.[id];
    if (!url) return null;
    // 큰 미디어: blob URL → Blob 그대로 (zipAsync 가 복사 없이 담음)
    const blob = url.startsWith('blob:') ? opts.blob?.(url) : null;
    let d = blob ? { bytes: blob, mime: blob.type || 'application/octet-stream', ext: extOfMime(blob.type) } : dataUrlBytes(url);
    if (!d) return null;
    if (d.ext === 'svg' && opts.png?.[id]) d = dataUrlBytes(opts.png[id]) ?? d;
    const path = `${dir}/${stem ?? (/^(video|audio)\//.test(d.mime) ? 'media' : 'image')}${++mediaN}.${d.ext}`;
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
      /** OLE 포함 파일 (oleObject .bin 또는 package .xlsx 등) */
      ole: (id, pkg, ext) => { const p = putMedia(id, 'ppt/embeddings', pkg ? 'Embedded' : 'oleObject'); return p ? add(pkg ? REL('package') : REL('oleObject'), `../embeddings/${p.split('/').pop()}`) : null; },
      /** a:hlinkClick — '#slideN' 은 슬라이드 관계 + hlinksldjump (PowerPoint 의 '이 문서의 슬라이드') */
      hlink: (url) => {
        const m = /^#slide(\d+)$/.exec(url);
        if (m && Number(m[1]) >= 1 && Number(m[1]) <= pres.slides.length) return `<a:hlinkClick r:id="${add(REL('slide'), `slide${m[1]}.xml`)}" action="ppaction://hlinksldjump"/>`;
        return `<a:hlinkClick r:id="${add(REL('hyperlink'), url, true)}"/>`;
      },
      /** 애니메이션 · 전환 소리 (합성한 WAV, 파일당 하나) → { rid, name } */
      sound: (key) => {
        const def = SOUNDS.find((x) => x[0] === key);
        if (!def) return null;
        const path = `ppt/media/${def[2]}`;
        if (!files[path]) { files[path] = soundWav(key); defaults.set('wav', 'audio/wav'); }
        return { rid: add(REL('audio'), `../media/${def[2]}`), name: def[2] };
      },
      /** 원래 XML 의 소리 (전환 · 애니메이션) */
      audio: (id) => { const p = putMedia(id); return p ? add(REL('audio'), `../media/${p.split('/').pop()}`) : null; },
      /** 비디오/오디오 파트: a:videoFile r:link + p14:media r:embed 두 관계 */
      av: (id, kind) => { const p = putMedia(id); if (!p) return null; const t = `../media/${p.split('/').pop()}`; return { link: add(REL(kind), t), embed: add('http://schemas.microsoft.com/office/2007/relationships/media', t) }; },
      chart: (ch, raw) => {
        const path = `ppt/charts/chart${++chartN}.xml`;
        // 손대지 않은 차트: 원래 차트 XML · 스타일 · 색 · 데이터 통합 문서를 그대로
        if (raw?.sig && raw.sig === stableJson(ch) && raw.rels.every((r) => r.external || r.xml != null || pres.media?.[r.media])) {
          const rx = raw.rels.map((r, k) => {
            if (r.external) return `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}" TargetMode="External"/>`;
            let target;
            if (r.xml != null) {
              const name = `${(r.name ?? 'part').replace(/\d*\.xml$/i, '')}${chartN}_${k + 1}.xml`;
              files[`ppt/charts/${name}`] = r.xml;
              override(`ppt/charts/${name}`, r.ct);
              target = name;
            } else {
              const img = r.type.endsWith('/image');
              const p2 = putMedia(r.media, img ? 'ppt/media' : 'ppt/embeddings', img ? null : 'Microsoft_Excel_Worksheet');
              target = `../${p2.slice(4)}`;
            }
            return `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(target)}"/>`;
          }).join('');
          files[path] = raw.xml;
          if (rx) files[`ppt/charts/_rels/chart${chartN}.xml.rels`] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rx}</Relationships>`;
          override(path, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
          return add(REL('chart'), `../charts/chart${chartN}.xml`, false, `chart${chartN}`);
        }
        // 데이터 통합 문서 (PowerPoint 의 [데이터 편집] 이 되도록)
        const wb = `ppt/embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx`;
        files[wb] = chartWorkbook(ch);
        defaults.set('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        const bgPath = ch.bgMedia ? putMedia(ch.bgMedia) : null;
        files[`ppt/charts/_rels/chart${chartN}.xml.rels`] = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL('package')}" Target="../embeddings/Microsoft_Excel_Worksheet${chartN}.xlsx"/>${bgPath ? `<Relationship Id="rId2" Type="${REL('image')}" Target="../media/${bgPath.split('/').pop()}"/>` : ''}</Relationships>`;
        files[path] = chartXml(ch, { embed: 'rId1', bgRid: bgPath ? 'rId2' : null });
        override(path, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
        return add(REL('chart'), `../charts/chart${chartN}.xml`, false, `chart${chartN}`);
      },
      xml: () => `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.join('')}</Relationships>`,
    };
  };

  // 레이아웃: (종류, 배경 장식) 조합마다 하나 — 슬라이드가 쓰는 것만 + 기본 9종
  const layoutDefs = [];
  const layoutKey = new Map();
  const designOf = (slide) => (slide.design && pres.designs?.some((d) => d.id === slide.design) ? slide.design : null);
  const layoutOf = (slide) => {
    const design = designOf(slide);
    const custom = !design && isCustomLayout(slide.layout) && pres.masters?.[slide.layout];
    const kind = LAYOUT_TYPE[slide.layout] || custom ? slide.layout : 'blank';
    const decor = slideDecor(pres, { ...slide, layout: kind });
    const key = `${design ?? ''}|${kind}|${JSON.stringify(decor)}`;
    if (layoutKey.has(key)) {
      const def = layoutKey.get(key);
      mergePhs(def, slide);
      return def;
    }
    // 사용자 지정 레이아웃: 이름 · 개체 틀은 그 레이아웃 마스터에서 (PowerPoint type="cust")
    const def = custom
      ? { kind, decor, name: layoutName(pres, kind), phs: custom.objects.filter((o) => o.ph).map((o) => ({ ...o })), index: layoutDefs.length + 1, showMaster: !custom.hideMaster }
      : { kind, decor, name: slide.bgObjects ? `${LAYOUT_NAME[kind] ?? '사용자 지정'} ${layoutDefs.length + 1}` : LAYOUT_NAME[kind] ?? '사용자 지정', phs: slide.bgObjects ? [] : layoutPlaceholders(kind, pres.size), index: layoutDefs.length + 1, showMaster: true };
    def.design = design;
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
  for (const c of pres.customLayouts ?? []) if (pres.masters?.[c.key]) layoutOf({ layout: c.key, objects: [], bgObjects: c.bgObjects });
  // 추가 슬라이드 마스터(디자인)마다 기본 레이아웃 9종
  for (const d of pres.designs ?? []) for (const k of Object.keys(LAYOUT_TYPE)) layoutOf({ layout: k, objects: [], design: d.id });
  const slideLayouts = pres.slides.map((s) => layoutOf(s));

  // 프레젠테이션
  const presRels = relsFor();
  // 마스터 묶음: [첫 마스터, …디자인] — 레이아웃이 있는 것만
  const masterGroups = [null, ...(pres.designs ?? []).map((d) => d.id)].filter((id, k) => k === 0 || layoutDefs.some((d) => d.design === id));
  const masterRids = masterGroups.map((_, k) => presRels.add(REL('slideMaster'), `slideMasters/slideMaster${k + 1}.xml`));
  const slideRids = pres.slides.map((_, i) => presRels.add(REL('slide'), `slides/slide${i + 1}.xml`));
  const authors = [];
  const authorId = (name) => { let i = authors.findIndex((a) => a.name === name); if (i < 0) { authors.push({ name, last: 0 }); i = authors.length - 1; } return i; };
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

  // 마스터 · 레이아웃 · 테마 (마스터마다 하나씩)
  let idNext = 2147483648;
  const masterIdOf = [];
  masterGroups.forEach((did, k) => {
    const n = k + 1;
    const pd = did ? presForSlide(pres, { design: did }) : pres;
    const defs = layoutDefs.filter((d) => (d.design ?? null) === did);
    const masterRels = relsFor();
    const layoutRids = defs.map((d) => masterRels.add(REL('slideLayout'), `../slideLayouts/slideLayout${d.index}.xml`));
    const themeN = n === 1 ? 1 : n + 1; // theme2 는 메모 마스터용
    masterRels.add(REL('theme'), `../theme/theme${themeN}.xml`);
    masterIdOf.push(idNext);
    files[`ppt/slideMasters/slideMaster${n}.xml`] = masterXml(pd, layoutRids, idNext + 1);
    idNext += defs.length + 1;
    files[`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`] = masterRels.xml();
    override(`ppt/slideMasters/slideMaster${n}.xml`, CT('slideMaster'));
    for (const d of defs) {
      const rels = relsFor();
      rels.add(REL('slideMaster'), `../slideMasters/slideMaster${n}.xml`);
      files[`ppt/slideLayouts/slideLayout${d.index}.xml`] = layoutXml(pd, d, rels);
      files[`ppt/slideLayouts/_rels/slideLayout${d.index}.xml.rels`] = rels.xml();
      override(`ppt/slideLayouts/slideLayout${d.index}.xml`, CT('slideLayout'));
    }
    // 원래 테마 (채우기 · 선 · 효과 · 배경 스타일, 글꼴 스크립트) 를 살리고 색 · 글꼴만 지금 값으로
    const traw = !did ? pres.themeRaw : (pres.designs.find((x) => x.id === did)?.themeRaw ?? null);
    let tx = null;
    if (traw?.xml) {
      const trels = relsFor();
      const body = remapRels(traw.xml, traw.rels, trels, pres);
      if (body) {
        tx = patchTheme(body, pd.theme);
        if (Object.keys(traw.rels ?? {}).length) files[`ppt/theme/_rels/theme${themeN}.xml.rels`] = trels.xml();
      }
    }
    files[`ppt/theme/theme${themeN}.xml`] = tx ?? themeXml(pd.theme, did ? (pres.designs.find((x) => x.id === did)?.name ?? 'WIPOINT') : 'WIPOINT');
    override(`ppt/theme/theme${themeN}.xml`, 'application/vnd.openxmlformats-officedocument.theme+xml');
  });
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
    const zooms = {};
    // 개체 id: 읽은 파일의 id 를 그대로 (애니메이션 · 연결선이 가리키는 id 유지), 새 개체는 비어 있는 번호
    const used = new Set([1]);
    const rawUse = new Map(); // 첫 개체 id → 원래 XML 로 쓸 개체들
    for (let i = 0; i < slide.objects.length; i++) {
      const o = slide.objects[i];
      if (!o.raw?.sig) continue;
      const members = o.raw.group ? slide.objects.slice(i, i + o.raw.n) : [o];
      if (o.raw.group && (members.length !== o.raw.n || members.some((m) => m.grp !== o.grp) || slide.objects.filter((m) => m.grp === o.grp).length !== o.raw.n)) continue;
      if (rawSig(members, o.raw.group ? bboxOf(members) : null) !== o.raw.sig) continue;
      if (o.raw.ids.some((k) => used.has(k))) continue;
      for (const k of o.raw.ids) used.add(k);
      rawUse.set(o.id, members);
    }
    let idNext = 2;
    const freeId = () => { while (used.has(idNext)) idNext++; used.add(idNext); return idNext; };
    const idFor = (o) => (o.spid && !used.has(o.spid) ? (used.add(o.spid), o.spid) : freeId());
    const parts = [];
    let rawCount = 0;
    // 그룹: 같은 grp 의 개체를 p:grpSp 로 묶음 (처음 나온 위치에)
    const done = new Set();
    for (const o of slide.objects) {
      if (done.has(o.id)) continue;
      if (rawUse.has(o.id)) {
        const members = rawUse.get(o.id);
        const xml = rawXml(o, members, rels, pres);
        if (xml) {
          parts.push(xml);
          rawCount++;
          for (const m of members) { done.add(m.id); if (m.spid) spid.set(m.id, m.spid); }
          if (o.raw.group) spid.set(`grp:${o.grp}`, o.raw.ids[0]);
          continue;
        }
        for (const k of o.raw.ids) used.delete(k);
      }
      if (o.grp) {
        const members = slide.objects.filter((x) => x.grp === o.grp);
        const gid = freeId();
        let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
        for (const m of members) { x1 = Math.min(x1, m.x); y1 = Math.min(y1, m.y); x2 = Math.max(x2, m.x + m.w); y2 = Math.max(y2, m.y + m.h); }
        const inner = members.map((m) => { done.add(m.id); const mid = idFor(m); spid.set(m.id, mid); return objXml(m, mid, rels, phIdx, pres.theme); }).join('');
        const ext = `<a:off x="${E(x1)}" y="${E(y1)}"/><a:ext cx="${E(x2 - x1)}" cy="${E(y2 - y1)}"/>`;
        parts.push(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gid}" name="Group ${gid}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>${ext}<a:chOff x="${E(x1)}" y="${E(y1)}"/><a:chExt cx="${E(x2 - x1)}" cy="${E(y2 - y1)}"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`);
        for (const m of members) spid.set(`grp:${o.grp}`, gid);
        continue;
      }
      done.add(o.id);
      const oid = idFor(o);
      spid.set(o.id, oid);
      if (o.type === 'zoom') {
        // 확대/축소: 대상 슬라이드로 가는 링크가 걸린 사각형 (PowerPoint 이전 버전의 대체 그림과 같은 동작) + 확장에 설정
        const ti = pres.slides.findIndex((x) => x.id === o.zoom?.slide);
        zooms[oid] = { ...o.zoom, slide: ti };
        const title = ti >= 0 ? slideTitle(pres.slides[ti]) || `슬라이드 ${ti + 1}` : '';
        const t = textBody([{ ...para(), align: 'ctr', runs: [{ t: title, size: 14, color: '#404040' }] }], { anchor: 'ctr' });
        parts.push(objXml({ ...o, type: 'shape', shape: 'rect', fill: { type: 'solid', color: '#F2F2F2' }, line: { color: '#A6A6A6', width: 1, dash: 'solid' }, text: t, link: ti >= 0 ? `#slide${ti + 1}` : undefined, name: o.name ?? `확대/축소 ${oid}` }, oid, rels, phIdx, pres.theme));
        continue;
      }
      parts.push(objXml(o, oid, rels, phIdx, pres.theme));
    }
    // 바닥글 (슬라이드 번호 · 날짜 · 글)
    const ft = pres.footer ?? {};
    if ((ft.slideNum || ft.date || ft.text) && !(ft.hideOnTitle && slide.layout === 'title')) {
      const W = pres.size.w;
      const H = pres.size.h;
      const mk = (type, x, w, text, algn, fld) => `<p:sp><p:nvSpPr><p:cNvPr id="${freeId()}" name="${type}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${type}" sz="quarter" idx="${{ dt: 10, ftr: 11, sldNum: 12 }[type]}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E(x)}" y="${E(H - 54)}"/><a:ext cx="${E(w)}" cy="${E(32)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="${algn}"/>${fld ? `<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="${fld}"><a:rPr lang="ko-KR" altLang="en-US" sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>${esc(text)}</a:t></a:fld>` : `<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>${esc(text)}</a:t></a:r>`}<a:endParaRPr lang="ko-KR" sz="1200"/></a:p></p:txBody></p:sp>`;
      if (ft.date) parts.push(mk('dt', W * 0.069, W * 0.225, ft.dateText || new Date().toLocaleDateString('ko-KR'), 'l', ft.dateText ? null : 'datetime1'));
      if (ft.text) parts.push(mk('ftr', W * 0.33, W * 0.34, ft.text, 'ctr'));
      if (ft.slideNum) parts.push(mk('sldNum', W * 0.706, W * 0.225, String(n), 'r', 'slidenum'));
    }
    const bgFill = slide.bg ? slide.bg : null;
    const bg = bgFill ? `<p:bg><p:bgPr>${fillXml(bgFill, rels)}<a:effectLst/></p:bgPr></p:bg>` : '';
    // 메모 (PowerPoint 이전 형식 p:cmLst — 모든 버전에서 열림; 답글은 ↳ 로 이어 붙임)
    if (slide.comments?.length) {
      let idx = 0;
      const cms = slide.comments.flatMap((c) => [c, ...(c.replies ?? []).map((r) => ({ ...r, x: c.x, y: c.y, text: `↳ ${r.text}` }))]).map((c) => {
        const aid = authorId(c.author ?? '사용자');
        const ai = authors[aid];
        ai.last = ++idx;
        return `<p:cm authorId="${aid}" dt="${esc((c.at ?? new Date().toISOString()).replace(/Z$/, ''))}" idx="${idx}"><p:pos x="${Math.round((c.x ?? 0) * 6)}" y="${Math.round((c.y ?? 0) * 6)}"/><p:text>${esc(c.text ?? '')}</p:text></p:cm>`;
      });
      const cpath = `ppt/comments/comment${n}.xml`;
      files[cpath] = `${XML_HEAD}<p:cmLst xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">${cms.join('')}</p:cmLst>`;
      override(cpath, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
      rels.add(REL('comments'), `../comments/comment${n}.xml`);
    }
    // 미디어 재생 옵션 · 내레이션 (PowerPoint 타이밍 대신 WIPOINT 확장에 그대로)
    const props = {};
    for (const o of slide.objects) {
      const pr = Object.fromEntries(['autoplay', 'loop', 'narration', 'hideInShow', 'volume'].filter((k) => o[k] != null).map((k) => [k, o[k]]));
      if (Object.keys(pr).length && spid.has(o.id)) props[spid.get(o.id)] = pr;
    }
    const meta = { design: slide.design ?? undefined, props: Object.keys(props).length ? props : undefined, inkRec: slide.inkRec?.length ? slide.inkRec : undefined, comments: slide.comments?.length ? slide.comments : undefined, layout: slide.layout, transition: slide.transition ?? null, anims: slide.anims ?? [], hideDecor: slide.hideDecor ?? false, section: slide.section ?? null, zooms: Object.keys(zooms).length ? zooms : undefined, bgObjects: slide.bgObjects ? true : undefined, ids: Object.fromEntries([...spid].filter(([k]) => !String(k).startsWith('grp:')).map(([k, v]) => [v, k])) };
    const ext = `<p:extLst><p:ext uri="${WP_EXT_URI}"><wp:slide xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(meta))}"/></p:ext></p:extLst>`;
    // 원래 애니메이션 XML: 애니메이션이 그대로이고 가리키는 개체 id 가 모두 같은 개체이면
    const byId = new Map([...spid].filter(([k]) => !String(k).startsWith('grp:')).map(([k, v]) => [v, k]));
    const rawTim = slide.rawTiming && slide.rawTiming.sig === stableJson(slide.anims) && slide.rawTiming.spids.every((sp) => { const oid = byId.get(sp); const o = oid && slide.objects.find((q) => q.id === oid); return o && o.spid === sp; }) ? rawPartXml(slide.rawTiming, rels) : null;
    const rawTr = slide.rawTransition && slide.rawTransition.sig === stableJson(slide.transition) ? rawPartXml(slide.rawTransition, rels) : null;
    const timing = rawTim ?? timingXml(slide, (oid) => spid.get(oid), (k) => rels.sound(k));
    // 원래 XML 을 쓰면 그 파일의 네임스페이스 선언도 (a14 · p14 · mc 등)
    const extraNs = (rawCount || rawTim || rawTr) && slide.rawNs ? Object.entries(slide.rawNs).filter(([k]) => !['xmlns:a', 'xmlns:r', 'xmlns:p'].includes(k)).map(([k, v]) => ` ${k}="${esc(v)}"`).join('') : '';
    files[`ppt/slides/slide${n}.xml`] = `${XML_HEAD}<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"${extraNs}${slide.hidden ? ' show="0"' : ''}><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${parts.join('')}</p:spTree></p:cSld><p:clrMapOvr>${slide.clrMap && stableJson({ ...STD_CLRMAP, ...slide.clrMap }) !== stableJson({ ...STD_CLRMAP, ...(presForSlide(pres, slide).clrMap ?? {}) }) ? `<a:overrideClrMapping ${clrMapAttrs(slide.clrMap)}/>` : '<a:masterClrMapping/>'}</p:clrMapOvr>${rawTr ?? transitionXml(slide.transition, (k) => rels.sound(k))}${timing}${ext}</p:sld>`;
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
  // 마스터 개체의 그림: 파일 안 경로를 기록 (다시 읽을 때 새 media id 로 연결)
  const masterMedia = {};
  for (const o of [...Object.values(pres.masters ?? {}).flatMap((m) => m.objects), ...(pres.designs ?? []).flatMap((d) => Object.values(d.masters ?? {}).flatMap((m) => m.objects)), ...(pres.customLayouts ?? []).flatMap((c) => c.bgObjects ?? [])]) for (const id of [o.media, o.fill?.media]) if (id && mediaPath.has(id)) masterMedia[id] = mediaPath.get(id);
  const presMeta = { designs: pres.designs?.length ? pres.designs.map(({ themeRaw, ...d }) => d) : undefined, customLayouts: pres.customLayouts?.length ? pres.customLayouts : undefined, masters: pres.masters ?? undefined, masterMedia: Object.keys(masterMedia).length ? masterMedia : undefined, theme: { name: th.name, label: th.label, accentBand: !!th.accentBand, titleBar: !!th.titleBar, dark: !!th.dark, bg: th.bg ?? null }, footer: pres.footer ?? null };
  // 안내선 (PowerPoint 2013+: p15:sldGuideLst, pos = 1/8 pt = px × 6)
  const guideExt = pres.guides?.length ? `<p:ext uri="{EFAFB233-063F-42B5-8137-9DF3F51BA10A}"><p15:sldGuideLst xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">${pres.guides.map((g, i) => `<p15:guide id="${i + 1}"${g.dir === 'h' ? ' orient="horz"' : ''} pos="${Math.round(g.pos * 6)}"/>`).join('')}</p15:sldGuideLst></p:ext>` : '';
  const secXml = `<p:extLst>${secExt}${guideExt}<p:ext uri="${WP_EXT_URI}"><wp:pres xmlns:wp="${WP_NS}" json="${esc(JSON.stringify(presMeta))}"/></p:ext></p:extLst>`;
  const sz = pres.size;
  const sldIds = slideRids.map((rid, i) => `<p:sldId id="${256 + i}" r:id="${rid}"/>`).join('');
  const defText = pres.defaultTextStyleRaw ?? `<p:defaultTextStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((l) => `<a:lvl${l}pPr marL="${(l - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${l}pPr>`).join('')}</p:defaultTextStyle>`;
  if (authors.length) {
    files['ppt/commentAuthors.xml'] = `${XML_HEAD}<p:cmAuthorLst xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}">${authors.map((a, i) => `<p:cmAuthor id="${i}" name="${esc(a.name)}" initials="${esc([...a.name][0] ?? '')}" lastIdx="${a.last}" clrIdx="${i % 8}"/>`).join('')}</p:cmAuthorLst>`;
    presRels.add(REL('commentAuthors'), 'commentAuthors.xml');
    override('ppt/commentAuthors.xml', 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
  }
  // 사용자 지정 쇼
  const custXml = (pres.customShows ?? []).length ? `<p:custShowLst>${pres.customShows.map((c, k) => `<p:custShow name="${esc(c.name)}" id="${k}"><p:sldLst>${c.slides.map((id) => pres.slides.findIndex((x) => x.id === id)).filter((i) => i >= 0).map((i) => `<p:sld r:id="${slideRids[i]}"/>`).join('')}</p:sldLst></p:custShow>`).join('')}</p:custShowLst>` : '';
  files['ppt/presentation.xml'] = `${XML_HEAD}<p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}" saveSubsetFonts="1"${fontXml ? ' embedTrueTypeFonts="1"' : ''}><p:sldMasterIdLst>${masterRids.map((r, k) => `<p:sldMasterId id="${masterIdOf[k]}" r:id="${r}"/>`).join('')}</p:sldMasterIdLst>${notesMasterRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>` : ''}${sldIds ? `<p:sldIdLst>${sldIds}</p:sldIdLst>` : ''}<p:sldSz cx="${E(sz.w)}" cy="${E(sz.h)}"/><p:notesSz cx="6858000" cy="9144000"/>${fontXml}${custXml}${defText}${secXml}</p:presentation>`;
  files['ppt/_rels/presentation.xml.rels'] = presRels.xml();
  override('ppt/presentation.xml', CT(opts.template ? 'template.main' : 'presentation.main'));
  files['ppt/presProps.xml'] = `${XML_HEAD}<p:presentationPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"/>`;
  files['ppt/viewProps.xml'] = `${XML_HEAD}<p:viewPr xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
  files['ppt/tableStyles.xml'] = pres.tableStylesRaw ?? `${XML_HEAD}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
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
    if (MOD_SHORT_R[m.name] == null) continue;
    if (['comp', 'inv', 'gray'].includes(m.name)) mods.push([m.name, 1]);
    else if (m.name === 'hueOff') mods.push([m.name, num(attr(m, 'val')) / 60000]);
    else mods.push([m.name, num(attr(m, 'val')) / 100000]);
  }
  const modStr = (list) => list.map(([k, v]) => `:${MOD_SHORT_R[k]}${k === 'hueOff' ? Math.round(v * 100) / 100 : Math.round(v * 1000) / 10}`).join('');
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
/** 색 변환 요소 → WIPOINT 색 글의 짧은 이름 (themes.js MOD_CODES 와 같음) */
const MOD_SHORT_R = { lumMod: 'lm', lumOff: 'lo', tint: 't', shade: 's', alpha: 'a', satMod: 'sm', satOff: 'so', hueMod: 'hm', hueOff: 'ho', alphaMod: 'am', alphaOff: 'ao', comp: 'cp', inv: 'iv', gray: 'gr', red: 'rd', green: 'gn', blue: 'bl', redMod: 'rm', greenMod: 'gm', blueMod: 'bm', redOff: 'ro', greenOff: 'go', blueOff: 'bo' };
const PRESET_COLORS = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', orange: '#FFA500', purple: '#800080', navy: '#000080', silver: '#C0C0C0' };

function readFill(spPr, ctx, phClr) {
  if (!spPr) return undefined;
  for (const c of spPr.children) {
    if (c.name === 'noFill') return null;
    if (c.name === 'solidFill') return { type: 'solid', color: readColor(c, ctx, phClr) ?? '#000000' };
    if (c.name === 'gradFill') {
      // 색 지점이 없는 그라데이션은 스타일 채우기를 따름 (PowerPoint 와 같음)
      if (!descendants(child(c, 'gsLst'), 'gs').length) return undefined;
      const stops = descendants(child(c, 'gsLst'), 'gs').map((g) => [num(attr(g, 'pos')) / 100000, readColor(g, ctx, phClr) ?? '#FFFFFF']).sort((a, b) => a[0] - b[0]);
      const lin = child(c, 'lin');
      const path = child(c, 'path');
      const out = { type: 'gradient', stops, angle: lin ? num(attr(lin, 'ang')) / 60000 : 90 };
      if (path) {
        // 방사형 (circle) · 사각형 (rect) · 도형 모양 (shape), 중심 = fillToRect
        const kind = attr(path, 'path') ?? 'circle';
        out.path = kind === 'circle' ? true : kind;
        const fr = child(path, 'fillToRect');
        const fv = ['l', 't', 'r', 'b'].map((k) => num(attr(fr, k), 0) / 100000);
        if (fr && fv.some((v) => Math.abs(v - 0.5) > 1e-6)) out.focus = fv;
      }
      return out;
    }
    if (c.name === 'blipFill') {
      const blip = child(c, 'blip');
      const media = ctx.media(attr(blip, 'embed'));
      if (!media) return null;
      const out = { type: 'image', media };
      const tl = child(c, 'tile');
      if (tl) {
        // 바둑판식: 그림 원래 크기 × 배율, 시작 위치 · 대칭 (PowerPoint [질감 · 그림 바둑판식])
        const d = (ctx.dims ?? (() => null))(media) ?? (typeof ctx.mediaDims === 'function' ? ctx.mediaDims(media) : null);
        out.tile = { sx: num(attr(tl, 'sx'), 100000) / 100000, sy: num(attr(tl, 'sy'), 100000) / 100000, tx: PX(attr(tl, 'tx') ?? 0), ty: PX(attr(tl, 'ty') ?? 0), ...(attr(tl, 'algn') && attr(tl, 'algn') !== 'tl' ? { algn: attr(tl, 'algn') } : {}), ...(attr(tl, 'flip') && attr(tl, 'flip') !== 'none' ? { flip: attr(tl, 'flip') } : {}), ...(d ? { w: d.w, h: d.h } : {}) };
      }
      const duo = child(blip, 'duotone');
      if (duo) { const cs = duo.children.map((k) => readColor({ children: [k] }, ctx, phClr)).filter(Boolean); if (cs.length === 2) out.duotone = cs; }
      const am = child(blip, 'alphaModFix');
      if (am) out.alpha = num(attr(am, 'amt'), 100000) / 100000;
      return out;
    }
    if (c.name === 'pattFill') return { type: 'pattern', prst: attr(c, 'prst') ?? 'pct5', fg: readColor(child(c, 'fgClr'), ctx, phClr) ?? '#000000', bg: readColor(child(c, 'bgClr'), ctx, phClr) ?? '#FFFFFF' };
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

const SHOW_KEYS = ['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize'];
/** c:dLbls 의 생략된 show* 요소를 val="0" 으로 채움 (개별 c:dLbl · delete 가 있는 묶음은 그대로) */
export function explicitDLbls(xml) {
  return xml.replace(/<c:dLbls>((?:(?!<\/c:dLbls>)[\s\S])*)<\/c:dLbls>/g, (all, inner) => {
    if (/<c:dLbl>|<c:delete\b/.test(inner)) return all;
    const at = inner.search(/<c:(showLegendKey|showVal|showCatName|showSerName|showPercent|showBubbleSize|separator|showLeaderLines|leaderLines|extLst)\b/);
    const head = at < 0 ? inner : inner.slice(0, at);
    const rest = at < 0 ? '' : inner.slice(at);
    const own = {};
    const tail = rest.replace(/<c:(showLegendKey|showVal|showCatName|showSerName|showPercent|showBubbleSize)\b[^>]*\/>/g, (m, k) => { own[k] = m; return ''; });
    return `<c:dLbls>${head}${SHOW_KEYS.map((k) => own[k] ?? `<c:${k} val="0"/>`).join('')}${tail}</c:dLbls>`;
  });
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
  const tabs = kids(child(el, 'tabLst'), 'tab').map((t) => ({ pos: PX(attr(t, 'pos') ?? 0), algn: attr(t, 'algn') ?? 'l' }));
  if (tabs.length) out.tabs = tabs;
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
  // 밑줄 종류 (sng = true, 그 밖은 이름: dbl · heavy · dotted · dash · wavy …), 취소선 이중
  if (a.u && a.u !== 'none') out.u = a.u === 'sng' ? true : a.u;
  if (a.strike && a.strike !== 'noStrike') out.s = a.strike === 'dblStrike' ? 'dbl' : true;
  const uFill = child(child(el, 'uFill'), 'solidFill');
  if (uFill) out.uColor = readColor(uFill, ctx, phClr);
  if (a.kern !== undefined) out.kern = num(a.kern) / 100;
  if (a.baseline && num(a.baseline) !== 0) out.base = num(a.baseline) > 0 ? 1 : -1;
  if (a.spc) out.spc = num(a.spc) / 100;
  if (a.cap && a.cap !== 'none') out.cap = a.cap;
  const fillEl = child(el, 'solidFill') ?? child(el, 'gradFill');
  if (fillEl) out.color = readColor(fillEl.name === 'gradFill' ? descendants(fillEl, 'gs')[0] : fillEl, ctx, phClr);
  // 글자 그라데이션 · 채우기 없음 (WordArt): 색은 첫 지점 색도 같이 둠 (편집 · 대체용)
  if (fillEl?.name === 'gradFill') { const g = readFill({ children: [fillEl] }, ctx, phClr); if (g?.type === 'gradient') out.fill = g; }
  if (child(el, 'noFill')) out.fill = { type: 'none' };
  const hl = child(el, 'highlight');
  if (hl) out.hl = readColor(hl, ctx);
  const ea = attr(child(el, 'ea'), 'typeface');
  const latin = attr(child(el, 'latin'), 'typeface');
  const font = [ea, latin].find((f) => f && !/^\+(mj|mn)-(cs)$/.test(f));
  if (font) out.font = font.startsWith('+mj') ? '+mj' : font.startsWith('+mn') ? '+mn' : font;
  const sym = attr(child(el, 'sym'), 'typeface');
  if (sym && !/^\+/.test(sym)) out.sym = sym;
  const eff = child(el, 'effectLst');
  const sh = child(eff, 'outerShdw');
  if (sh) {
    // 기본 그림자 (blur 3pt · dist 3pt · 45° · 검정 43%) 는 true, 다르면 값 그대로
    const col = readColor(sh, ctx, phClr) ?? '#000000';
    const v = { blur: num(attr(sh, 'blurRad')) / 12700, dist: num(attr(sh, 'dist')) / 12700, dir: num(attr(sh, 'dir')) / 60000, color: col };
    const isDef = Math.abs(v.blur - 3) < 0.05 && Math.abs(v.dist - 3) < 0.05 && Math.abs(v.dir - 45) < 0.5 && /^#000000:a4[23]/.test(col);
    out.shadow = isDef ? true : v;
  }
  const gl = child(eff, 'glow');
  if (gl) out.glow = { rad: num(attr(gl, 'rad')) / 12700, color: readColor(gl, ctx, phClr) ?? '#FFC000' };
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
  // 단 (numCol · spcCol), 글 회전, 똑바로 세우기, WordArt 변환 (prstTxWarp)
  if (num(a.numCol) > 1) { out.cols = num(a.numCol); if (a.spcCol) out.colSpc = PX(a.spcCol); }
  if (a.rot && num(a.rot)) out.rot = num(a.rot) / 60000;
  if (bool(a.upright)) out.upright = true;
  if (bool(a.anchorCtr)) out.anchorCtr = true;
  const warp = child(bp, 'prstTxWarp');
  if (warp && attr(warp, 'prst') && attr(warp, 'prst') !== 'textNoShape') {
    const adj = {};
    for (const gd of descendants(child(warp, 'avLst'), 'gd')) { const m = /val\s+(-?\d+)/.exec(attr(gd, 'fmla') ?? ''); if (m) adj[attr(gd, 'name')] = Number(m[1]); }
    out.warp = { prst: attr(warp, 'prst'), ...(Object.keys(adj).length ? { adj } : {}) };
  }
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
        // 필드 (슬라이드 번호 · 날짜): 종류를 두고 화면에서 현재 값으로
        if (c.name === 'fld' && attr(c, 'type')) r.fld = { type: attr(c, 'type'), id: attr(c, 'id') };
        if (t || r.fld) runs.push(cleanRun({ t, ...r }));
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
    if (base.tabs) out.tabs = base.tabs;
    delete out.end.t;
    return out;
  });
  return { paras: paras.length ? paras : [para()], anchor: bp.anchor ?? 't', insets: bp.insets ?? [9.6, 4.8, 9.6, 4.8], wrap: bp.wrap ?? true, autofit: bp.autofit ?? 'none', ...(bp.fontScale ? { fontScale: bp.fontScale } : {}), ...(bp.lnSpcReduction ? { lnSpcReduction: bp.lnSpcReduction } : {}), ...(bp.vert ? { vert: bp.vert } : {}), ...Object.fromEntries(['cols', 'colSpc', 'rot', 'upright', 'anchorCtr', 'warp'].filter((k) => bp[k] != null).map((k) => [k, bp[k]])) };
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
    // 안내선 (avLst · gdLst) 은 도형 크기(EMU)로 계산해 숫자로 바꿈. 경로에 w · h 가 없으면 좌표 단위 = 도형 EMU
    const ext = child(child(spPr, 'xfrm'), 'ext');
    const W = num(attr(ext, 'cx')) || 914400; const H = num(attr(ext, 'cy')) || 914400;
    const gds = (el) => kids(el, 'gd').map((g) => [attr(g, 'name'), attr(g, 'fmla') ?? 'val 0']);
    const v = guideVars(W, H, gds(child(cg, 'avLst')), gds(child(cg, 'gdLst')));
    const val = (x) => { const n = Number(x); return Number.isFinite(n) ? n : (v[x] ?? 0); };
    const paths = kids(child(cg, 'pathLst'), 'path').map((p) => {
      const cmds = [];
      for (const c of p.children) {
        const pts = kids(c, 'pt').flatMap((pt) => [val(attr(pt, 'x')), val(attr(pt, 'y'))]);
        if (c.name === 'moveTo') cmds.push(['M', ...pts]);
        else if (c.name === 'lnTo') cmds.push(['L', ...pts]);
        else if (c.name === 'cubicBezTo') cmds.push(['C', ...pts]);
        else if (c.name === 'quadBezTo') cmds.push(['Q', ...pts]);
        else if (c.name === 'arcTo') cmds.push(['A', val(attr(c, 'wR')), val(attr(c, 'hR')), val(attr(c, 'stAng')), val(attr(c, 'swAng'))]);
        else if (c.name === 'close') cmds.push(['Z']);
      }
      const pw = num(attr(p, 'w')); const ph = num(attr(p, 'h'));
      return { w: pw || W, h: ph || H, cmds, ...(attr(p, 'fill') === 'none' ? { fill: 'none' } : {}), ...(attr(p, 'stroke') === '0' || attr(p, 'stroke') === 'false' ? { stroke: false } : {}) };
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
    effectStyles: child(fmt, 'effectStyleLst')?.children ?? [],
  };
}

const TABLE_STYLE_ACCENT = {
  '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': 'accent1', '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}': 'accent2', '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}': 'accent3',
  '{00A15C55-8517-42AA-B614-E9B94910E393}': 'accent4', '{7DF18680-E054-41AD-8BC1-D1AEF772440D}': 'accent5', '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}': 'accent6',
};
const PRESET_ANIM = {
  entr: { 1: 'appear', 2: 'fly', 3: 'blinds', 4: 'box', 5: 'checker', 6: 'shape', 7: 'crawl', 8: 'diamond', 9: 'dissolveIn', 10: 'fade', 12: 'peek', 13: 'plus', 14: 'randomBars', 15: 'spiralIn', 16: 'split', 17: 'stretch', 18: 'strips', 19: 'swivel', 20: 'wedge', 21: 'wheel', 22: 'wipe', 23: 'basicZoom', 25: 'boomerang', 26: 'bounce', 28: 'credits', 29: 'float', 30: 'float', 31: 'grow', 35: 'pinwheel', 37: 'riseUp', 42: 'float', 43: 'centerRevolve', 45: 'swivel', 47: 'sinkDown', 50: 'compress', 52: 'curveUp', 53: 'zoom', 55: 'expand', 56: 'flip' },
  emph: { 1: 'fillColor', 3: 'objectColor', 6: 'growShrink', 7: 'lineColor', 8: 'spin', 9: 'transparency', 10: 'boldFlash', 15: 'bold', 16: 'brushColor', 18: 'underline', 19: 'objectColor', 21: 'complementary', 22: 'complementary', 24: 'darken', 25: 'desaturate', 26: 'pulse', 27: 'colorPulse', 30: 'lighten', 32: 'teeter', 34: 'wave', 35: 'flash', 36: 'shimmer' },
};
PRESET_ANIM.exit = { ...PRESET_ANIM.entr, 1: 'disappear' };
const SUB_DIR = { 4: 'b', 1: 't', 8: 'l', 2: 'r' };
const LAYOUT_KIND = Object.fromEntries(Object.entries(LAYOUT_TYPE).map(([k, v]) => [v, k]));
const TRANS_NAME = { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'uncover', zoom: 'zoom', circle: 'circle', dissolve: 'dissolve', cut: 'cut', flip: 'flip', random: 'random', blinds: 'blinds', checker: 'checker', comb: 'comb', randomBar: 'randomBars', strips: 'strips', wedge: 'clock', wheel: 'clock', wheelReverse: 'clock', newsflash: 'zoom', plus: 'plus', diamond: 'diamond', vortex: 'vortex', ripple: 'ripple', honeycomb: 'honeycomb', glitter: 'glitter', shred: 'shred', switch: 'switch', gallery: 'gallery', cube: 'cube', doors: 'doors', window: 'window', box: 'box', conveyor: 'conveyor', pan: 'pan', ferris: 'ferris', flythrough: 'flythrough', warp: 'zoom', prism: 'cube', reveal: 'reveal', flash: 'flash', morph: 'morph', prstTrans: 'fade' };
const PRST_TRANS = { fallOver: 'fall', drape: 'drape', curtains: 'curtains', wind: 'wind', prestige: 'prestige', fracture: 'fracture', crush: 'crush', peelOff: 'peelOff', pageCurlDouble: 'pageCurl', pageCurlSingle: 'pageCurl', airplane: 'airplane', origami: 'origami' };
const DIR_FROM = { u: 'b', d: 't', l: 'r', r: 'l' };

/** 그림 바이트 → 화면 크기 px (96dpi 기준, 파일의 dpi 반영) — PNG · JPEG · GIF · BMP, 모르면 null */
export function imageSize(b) {
  if (!b || b.length < 24) return null;
  const be16 = (o) => (b[o] << 8) | b[o + 1];
  const be32 = (o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
  if (b[0] === 0x89 && b[1] === 0x50) {
    const w = be32(16); const h = be32(20);
    let dpi = 96;
    for (let o = 8; o + 12 < b.length && o < 4096;) {
      const len = be32(o); const t = String.fromCharCode(b[o + 4], b[o + 5], b[o + 6], b[o + 7]);
      if (t === 'pHYs' && b[o + 16] === 1) { dpi = be32(o + 8) * 0.0254; break; }
      if (t === 'IDAT') break;
      o += 12 + len;
    }
    return { w: (w * 96) / (dpi || 96), h: (h * 96) / (dpi || 96) };
  }
  if (b[0] === 0xff && b[1] === 0xd8) {
    let dpi = 96;
    for (let o = 2; o + 9 < b.length;) {
      if (b[o] !== 0xff) { o++; continue; }
      const m = b[o + 1]; const len = be16(o + 2);
      if (m === 0xe0 && b[o + 4] === 0x4a && b[o + 11] === 1 && be16(o + 12)) dpi = be16(o + 12);
      if ((m >= 0xc0 && m <= 0xc3) || (m >= 0xc5 && m <= 0xc7) || (m >= 0xc9 && m <= 0xcb) || (m >= 0xcd && m <= 0xcf)) return { w: (be16(o + 7) * 96) / dpi, h: (be16(o + 5) * 96) / dpi };
      o += 2 + len;
    }
    return null;
  }
  if (b[0] === 0x47 && b[1] === 0x49) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
  if (b[0] === 0x42 && b[1] === 0x4d) { const w = b[18] | (b[19] << 8) | (b[20] << 16); const h = Math.abs((b[22] | (b[23] << 8) | (b[24] << 16) | (b[25] << 24)) >> 0); return { w, h }; }
  return null;
}

// ───────────── 원래 XML 그대로 쓰기 (고치지 않은 개체) ─────────────
/** 키 순서와 무관한 JSON */
export function stableJson(v) {
  if (Array.isArray(v)) return `[${v.map(stableJson).join(',')}]`;
  if (v && typeof v === 'object') return `{${Object.keys(v).filter((k) => v[k] !== undefined).sort().map((k) => `${JSON.stringify(k)}:${stableJson(v[k])}`).join(',')}}`;
  return JSON.stringify(v ?? null);
}
const GEOM_KEYS = new Set(['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV']);
const SKIP_KEYS = new Set(['id', 'raw', 'grp', 'spid', 'name', 'decor']);
function bboxOf(objs) {
  let x1 = Infinity; let y1 = Infinity; let x2 = -Infinity; let y2 = -Infinity;
  for (const m of objs) { x1 = Math.min(x1, m.x); y1 = Math.min(y1, m.y); x2 = Math.max(x2, m.x + m.w); y2 = Math.max(y2, m.y + m.h); }
  return { x: x1, y: y1, w: Math.max(1e-6, x2 - x1), h: Math.max(1e-6, y2 - y1) };
}
/** 서명: 위치 · 크기를 뺀 모든 속성 (그룹은 묶음 안 상대 위치 포함 — 묶음 전체 이동 · 크기 조절은 그대로 씀) */
function rawSig(objs, bb) {
  const r = (v) => Math.round(v * 100) / 100;
  return stableJson(objs.map((o) => {
    const c = {};
    for (const [k, v] of Object.entries(o)) if (!SKIP_KEYS.has(k) && !(GEOM_KEYS.has(k) && !bb)) c[k] = v;
    if (bb) { c.x = r((o.x - bb.x) / bb.w * 1000); c.y = r((o.y - bb.y) / bb.h * 1000); c.w = r(o.w / bb.w * 1000); c.h = r(o.h / bb.h * 1000); }
    return c;
  }));
}

/**
 * pptx 바이트 → 문서
 * 반환 { pres, warnings }
 */
export function readPptx(input, opts = {}) {
  // input: pptx 바이트, 또는 openZipFile 결과 ({ 경로: Uint8Array|Blob } — 큰 미디어는 Blob, opts.blobUrl 로 연결)
  const files = input instanceof Uint8Array ? unzip(input) : input;
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
  // 메모 작성자
  const legacyAuthors = new Map();
  const modernAuthors = new Map();
  for (const r of presRels.values()) {
    if (r.type.endsWith('/commentAuthors') && has(r.target)) for (const a of kids(xmlOf(r.target), 'cmAuthor')) legacyAuthors.set(attr(a, 'id'), attr(a, 'name') ?? '');
    if (r.type.endsWith('/authors') && has(r.target)) for (const a of kids(xmlOf(r.target), 'author')) modernAuthors.set(attr(a, 'id'), attr(a, 'name') ?? '');
  }
  const sz = child(presXml, 'sldSz');
  const size = { w: Math.round(PX(attr(sz, 'cx') ?? 12192000)), h: Math.round(PX(attr(sz, 'cy') ?? 6858000)) };

  const media = {};
  const mediaByPath = new Map();
  const mediaDims = new Map();
  const loadMedia = (path, av = false) => {
    if (!path || !has(path)) return null;
    if (mediaByPath.has(path)) return mediaByPath.get(path);
    const ext = path.split('.').pop().toLowerCase();
    const mime = av === 'bin' ? OLE_MIME[ext] ?? 'application/vnd.openxmlformats-officedocument.oleObject' : av ? AV_MIME[ext] ?? 'video/mp4' : MIME_BY_EXT[ext];
    if (!mime) { warnings.add(`지원하지 않는 그림 형식(.${ext})은 빈 칸으로 표시합니다`); mediaByPath.set(path, null); return null; }
    const b = files[path];
    if (typeof Blob === 'function' && b instanceof Blob) {
      const id = `m${mediaByPath.size + 1}${uid('').slice(-4)}`;
      media[id] = opts.blobUrl ? opts.blobUrl(b, mime) : null;
      mediaByPath.set(path, media[id] ? id : null);
      return media[id] ? id : null;
    }
    let bin = '';
    for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    const id = `m${mediaByPath.size + 1}${uid('').slice(-4)}`;
    media[id] = `data:${mime};base64,${btoa(bin)}`;
    mediaByPath.set(path, id);
    if (!av) { const d = imageSize(b); if (d) mediaDims.set(id, d); }
    return id;
  };

  // 테마 (첫 마스터의 테마)
  const masterIds = descendants(child(presXml, 'sldMasterIdLst'), 'sldMasterId');
  const firstMaster = masterIds.length ? presRels.get(rid(masterIds[0]))?.target : null;
  let themeInfo = null;
  if (firstMaster) {
    const tRel = [...relsOf(firstMaster).values()].find((r) => r.type.endsWith('/theme'));
    if (tRel && has(tRel.target)) themeInfo = { ...themeFromXml(textOf(files[tRel.target])), path: tRel.target };
  }
  const theme = themeInfo?.theme ?? cloneTheme(THEMES[0]);

  // 마스터마다 자기 테마 (채우기 · 선 · 효과 · 배경 스타일) — 테마 그림은 테마 파트의 관계로
  const themeInfoCache = new Map();
  const themeInfoFor = (mpath) => {
    if (themeInfoCache.has(mpath)) return themeInfoCache.get(mpath);
    const tRel = [...relsOf(mpath).values()].find((r) => r.type.endsWith('/theme'));
    let ti = themeInfo;
    if (tRel && has(tRel.target) && tRel.target !== themeInfo?.path) { try { ti = { ...themeFromXml(textOf(files[tRel.target])), path: tRel.target }; } catch { ti = themeInfo; } }
    themeInfoCache.set(mpath, ti);
    return ti;
  };
  const TI = (ctx) => ctx?.ti ?? themeInfo;
  const tplCtx = (ctx) => ({ ...ctx, media: (r) => { const t = TI(ctx)?.path; const rel = t ? relsOf(t).get(r) : null; return rel && !rel.external ? loadMedia(rel.target) : null; } });

  // 마스터 · 레이아웃 캐시
  const masterCache = new Map();
  const layoutCache = new Map();
  const partCtx = (part, clrMap) => ({
    clrMap,
    media: (rid) => { const r = relsOf(part).get(rid); return r ? loadMedia(r.target) : null; },
    dims: (id) => mediaDims.get(id) ?? null,
    rel: (rid) => (rid ? relsOf(part).get(rid) ?? null : null),
    avMedia: (path) => loadMedia(path, true),
    binMedia: (path) => loadMedia(path, 'bin'),
    links: (rid) => {
      const r = relsOf(part).get(rid);
      if (r?.external) return r.target;
      // 슬라이드로 가는 링크 (ppaction://hlinksldjump) → '#slideN'
      if (r?.type?.endsWith('/slide')) { const n = slideOrder.indexOf(r.target); return n >= 0 ? `#slide${n + 1}` : null; }
      return null;
    },
  });
  /** 예전 형식 OLE 미리 보기: 슬라이드의 VML 그림 (v:shape id = spid → v:imagedata o:relid) */
  const vmlCache = new Map();
  const vmlImage = (part, spid) => {
    if (!spid) return null;
    for (const r of relsOf(part).values()) {
      if (!r.type.endsWith('/vmlDrawing') || !has(r.target)) continue;
      if (!vmlCache.has(r.target)) {
        const m = new Map();
        try {
          const vx = parseXml(textOf(files[r.target]).replace(/<br\s*>/gi, '<br/>'));
          // VML 도형은 id 또는 o:spid 로 OLE 개체(spid)와 이어짐
          for (const sh of descendants(vx, 'shape')) { const im = descendants(sh, 'imagedata')[0]; const rr = attr(im, 'relid') ?? attr(im, 'id'); if (!rr) continue; const t = relsOf(r.target).get(rr)?.target; for (const k of [attr(sh, 'id'), attr(sh, 'spid'), sh.attrs?.['o:spid']]) if (k) m.set(k, t); }
        } catch { /* 무시 */ }
        vmlCache.set(r.target, m);
      }
      const t = vmlCache.get(r.target).get(spid);
      if (t) return loadMedia(t);
    }
    return null;
  };
  const slideOrder = descendants(child(presXml, 'sldIdLst'), 'sldId').map((sid) => presRels.get(rid(sid))?.target).filter((t) => t && has(t));
  const readClrMap = (el) => (el ? Object.fromEntries(Object.entries(el.attrs)) : null);
  const phKey = (sp) => { const ph = descendants(child(sp, 'nvSpPr') ?? child(sp, 'nvPicPr') ?? child(sp, 'nvGraphicFramePr'), 'ph')[0]; return ph ? { type: attr(ph, 'type') ?? 'body', idx: attr(ph, 'idx') ?? null } : null; };
  const phList = (spTree) => (spTree?.children ?? []).filter((c) => ['sp', 'pic'].includes(c.name)).map((sp) => ({ key: phKey(sp), sp })).filter((x) => x.key);
  const loadMaster = (path) => {
    if (masterCache.has(path)) return masterCache.get(path);
    const x = xmlOf(path);
    const clrMap = readClrMap(child(x, 'clrMap'));
    const ctx = partCtx(path, clrMap);
    const tx = child(x, 'txStyles');
    ctx.ti = themeInfoFor(path);
    const m = { path, x, clrMap, ctx, ti: ctx.ti, phs: phList(descendants(x, 'spTree')[0]), spTree: descendants(x, 'spTree')[0], bg: child(child(x, 'cSld'), 'bg'), styles: { title: readListStyle(child(tx, 'titleStyle'), ctx), body: readListStyle(child(tx, 'bodyStyle'), ctx), other: readListStyle(child(tx, 'otherStyle'), ctx) } };
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
    const l = { path, x, master, clrMap, ctx: { ...partCtx(path, clrMap), ti: master.ti }, phs: phList(descendants(x, 'spTree')[0]), spTree: descendants(x, 'spTree')[0], bg: child(child(x, 'cSld'), 'bg'), type: attr(x, 'type'), showMasterSp: attr(x, 'showMasterSp') !== '0' };
    layoutCache.set(path, l);
    return l;
  };
  const findPh = (list, key) => {
    if (!key) return null;
    const norm = (t) => (t === 'ctrTitle' ? 'title' : t === 'obj' ? 'body' : t);
    if (key.idx != null) { const byIdx = list.find((p) => p.key.idx === key.idx && (norm(p.key.type) === norm(key.type) || key.type === 'body' || key.type === 'obj')); if (byIdx) return byIdx.sp; }
    return list.find((p) => p.key.type === key.type)?.sp ?? list.find((p) => norm(p.key.type) === norm(key.type))?.sp ?? (key.idx != null ? list.find((p) => p.key.idx === key.idx)?.sp : null) ?? null;
  };
  // 마스터 글 스타일은 쓰는 쪽(슬라이드 · 레이아웃)의 색 매핑(clrMap)으로 읽어야 함 (레이아웃의 clrMapOvr: 어두운 배경 + 밝은 글자)
  const styleCache = new Map();
  const styleFor = (master, key, clrMap) => {
    const ck = `${master.path}|${JSON.stringify(clrMap ?? master.clrMap)}`;
    if (!styleCache.has(ck)) {
      const c2 = { ...master.ctx, clrMap: clrMap ?? master.clrMap };
      const tx = child(master.x, 'txStyles');
      styleCache.set(ck, { title: readListStyle(child(tx, 'titleStyle'), c2), body: readListStyle(child(tx, 'bodyStyle'), c2), other: readListStyle(child(tx, 'otherStyle'), c2) });
    }
    const st = styleCache.get(ck);
    if (!key) return st.other;
    if (key.type === 'title' || key.type === 'ctrTitle') return st.title;
    if (['dt', 'ftr', 'sldNum'].includes(key.type)) return st.other;
    return st.body;
  };
  const withMap = (c, clrMap) => (clrMap ? { ...c, clrMap } : c);

  /** 다른 프로그램 파일의 두 번째 이후 마스터: 테마가 첫 마스터와 다르면 디자인으로 */
  const foreignDesigns = new Map();
  const foreignDesign = (mpath) => {
    if (foreignDesigns.has(mpath)) return foreignDesigns.get(mpath);
    const tRel = [...relsOf(mpath).values()].find((r) => r.type.endsWith('/theme'));
    const firstRel = firstMaster ? [...relsOf(firstMaster).values()].find((r) => r.type.endsWith('/theme')) : null;
    let d = null;
    if (tRel && has(tRel.target) && tRel.target !== firstRel?.target) {
      const th = themeFromXml(textOf(files[tRel.target]))?.theme;
      if (th && JSON.stringify(th.colors) + JSON.stringify(th.fonts) !== JSON.stringify(theme.colors) + JSON.stringify(theme.fonts)) d = { id: uid('d'), name: th.label ?? th.name ?? `디자인 ${foreignDesigns.size + 2}`, theme: th, masters: {} };
    }
    foreignDesigns.set(mpath, d);
    return d;
  };

  /** 다른 프로그램의 사용자 지정 레이아웃 → WIPOINT 사용자 지정 레이아웃 (개체 틀 + 배경 장식) */
  const foreignLayouts = [];
  const foreignByPath = new Map();
  const foreignLayout = (layout, decor) => {
    if (foreignByPath.has(layout.path)) return foreignByPath.get(layout.path);
    const objs = [];
    try { readShapeTree(layout.spTree, { ...layout.ctx, part: layout.path }, layout, null, objs, {}); } catch { /* 개체 틀 없이 */ }
    const key = `cust:${uid('l')}`;
    const phs = objs.filter((o) => o.ph && !['dt', 'ftr', 'sldNum'].includes(o.ph)).map((o) => ({ ...o, phLayout: key }));
    const name = attr(child(layout.x, 'cSld'), 'name') || '사용자 지정 레이아웃';
    foreignLayouts.push({ key, name, objects: phs, bgObjects: decor.map((o) => ({ ...o })) });
    foreignByPath.set(layout.path, key);
    return key;
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
  // 그룹 좌표 (xf = 그룹 xfrm + up = 바깥 그룹) → 슬라이드 좌표. 그룹 회전 · 대칭은 자식 위치 · 회전 · 대칭에 반영
  const applyXf = (xf, r) => placeInGroups(xf, r);
  const styleRefs = (sp, ctx) => {
    const st = child(sp, 'style');
    if (!st) return {};
    const out = {};
    const fr = child(st, 'fillRef');
    const lr = child(st, 'lnRef');
    const fo = child(st, 'fontRef');
    if (fr && num(attr(fr, 'idx')) > 0) {
      const col = readColor(fr, ctx);
      const tpl = TI(ctx)?.fillStyles?.[num(attr(fr, 'idx')) - 1];
      if (tpl && tpl.name !== 'solidFill' && tpl.name !== 'noFill') { const g = readFill({ children: [tpl] }, tplCtx(ctx), col); out.fill = g ?? { type: 'solid', color: col }; } else if (tpl?.name === 'noFill') out.fill = null; else out.fill = col ? { type: 'solid', color: col } : undefined;
    } else if (fr) out.fill = null;
    if (lr && num(attr(lr, 'idx')) > 0) {
      const col = readColor(lr, ctx);
      const tpl = TI(ctx)?.lineStyles?.[num(attr(lr, 'idx')) - 1];
      out.line = (tpl && readLine(tpl, ctx, col ?? '#000000')) ?? { color: col ?? '#000000', width: tpl && attr(tpl, 'w') ? PX(attr(tpl, 'w')) : 1 };
    } else if (lr) out.line = null;
    if (fo) out.fontColor = readColor(fo, ctx);
    // 테마 효과 스타일 (effectRef: 그림자 · 입체 등)
    const er = child(st, 'effectRef');
    if (er && num(attr(er, 'idx')) > 0) {
      const col = readColor(er, ctx);
      const tpl = TI(ctx)?.effectStyles?.[num(attr(er, 'idx')) - 1];
      if (tpl) { const fx = fxFromXml(tpl, { child, attr, readColor: (e) => readColor(e, ctx, col) }); if (fx) out.fx = fx; }
    }
    return out;
  };

  function readElement(c, ctx, layout, xf, out, opts) {
    if (c.name === 'AlternateContent') {
      // 수식 (a14:m): 대체 글 상자의 자리 · 크기에 OMML 을 LaTeX 로 읽은 수식 개체
      const choice = child(c, 'Choice');
      const math = choice ? descendants(choice, 'oMathPara')[0] ?? descendants(choice, 'oMath')[0] : null;
      if (math) {
        const tmp = [];
        for (const x of (child(c, 'Fallback') ?? choice).children) readElement(x, ctx, layout, xf, tmp, opts);
        const base = tmp[0];
        if (base) {
          const rPr = descendants(math, 'rPr').find((r) => attr(r, 'sz'));
          const col = descendants(math, 'srgbClr')[0];
          out.push({ id: base.id, name: base.name, type: 'equation', x: base.x, y: base.y, w: base.w, h: base.h, rot: base.rot ?? 0, latex: ommlToLatex(math), size: rPr ? num(attr(rPr, 'sz')) / 100 : 28, ...(col ? { color: `#${attr(col, 'val')}` } : {}) });
          return;
        }
      }
      // 대체본(Fallback)이 있으면 그것을 — 새 기능(Choice)은 확장 네임스페이스가 필요할 수 있음
      const target = child(c, 'Fallback') ?? child(c, 'Choice');
      for (const x of target?.children ?? []) readElement(x, ctx, layout, xf, out, opts);
      return;
    }
    if (c.name === 'grpSp') {
      const gx = readXfrm(child(child(c, 'grpSpPr'), 'xfrm'));
      const composed = gx ? { ...gx, up: xf } : xf;
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
      let fill = readFill(spPr, ctx);
      if (fill === undefined && lph) fill = readFill(child(lph, 'spPr'), withMap(layout.ctx, ctx.clrMap));
      if (fill === undefined) fill = refs.fill ?? null;
      let line = readLine(child(spPr, 'ln'), ctx, refs.line?.color);
      if (line === undefined) line = refs.line ?? null;
      const isTx = bool(attr(child(nv, 'cNvSpPr'), 'txBox'));
      const chain = [];
      let bodyBase = {};
      if (layout) {
        chain.push(styleFor(layout.master, key, ctx.clrMap));
        if (mph) { chain.push(readListStyle(child(child(mph, 'txBody'), 'lstStyle'), withMap(layout.master.ctx, ctx.clrMap))); bodyBase = { ...bodyBase, ...readBodyPr(child(child(mph, 'txBody'), 'bodyPr')) }; }
        if (lph) { chain.push(readListStyle(child(child(lph, 'txBody'), 'lstStyle'), withMap(layout.ctx, ctx.clrMap))); bodyBase = { ...bodyBase, ...readBodyPr(child(child(lph, 'txBody'), 'bodyPr')) }; }
      }
      const text = readTxBody(child(c, 'txBody'), ctx, chain.filter(Boolean), bodyBase, refs.fontColor);
      const o = {
        id: uid(), type: 'shape', ...rect, ...geom, fill, line,
        text: text ?? (key && key.type !== 'pic' ? readTxBody({ children: [{ name: 'bodyPr', attrs: {}, children: [] }, { name: 'p', attrs: {}, children: [] }] }, ctx, chain.filter(Boolean), bodyBase, refs.fontColor) : null),
        name: attr(cNvPr, 'name'),
      };
      if (attr(cNvPr, 'id')) o.spid = num(attr(cNvPr, 'id'));
      if (attr(cNvPr, 'descr')) o.alt = attr(cNvPr, 'descr');
      const link = child(cNvPr, 'hlinkClick');
      if (link) { const url = ctx.links?.(attr(link, 'id')); if (url) o.link = url; }
      if (isTx) o.txBox = true;
      { const fx = fxFromXml(spPr, { child, attr, readColor: (e) => readColor(e, ctx) }); if (fx) o.fx = fx; else if (refs.fx && !child(spPr, 'effectLst') && !child(spPr, 'effectDag')) o.fx = refs.fx; }
      if (key) {
        o.ph = { ctrTitle: 'ctrTitle', title: 'title', subTitle: 'subTitle', body: 'body', obj: 'body', pic: 'pic', dt: 'dt', ftr: 'ftr', sldNum: 'sldNum', tbl: 'body', chart: 'body', media: 'body', clipArt: 'body', dgm: 'body' }[key.type] ?? 'body';
        if (['chart', 'tbl', 'dgm', 'media'].includes(key.type)) o.phFor = key.type;
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
      // 비디오 · 오디오
      const nvPr = child(nv, 'nvPr');
      const avEl = child(nvPr, 'videoFile') ?? child(nvPr, 'audioFile');
      if (avEl) {
        const p14 = descendants(nvPr, 'media')[0];
        const rel = ctx.rel?.(attr(p14, 'embed')) ?? ctx.rel?.(attr(avEl, 'link'));
        const avId = rel && !rel.external ? ctx.avMedia(rel.target) : null;
        if (avId) {
          const mo = { id: uid(), type: 'media', kind: avEl.name === 'audioFile' ? 'audio' : 'video', ...rect, media: avId, name: attr(cNvPr, 'name') };
          if (mid) mo.poster = mid;
          const tr = descendants(p14, 'trim')[0];
          if (tr) mo.trim = { st: num(attr(tr, 'st'), 0) / 1000, end: num(attr(tr, 'end'), 0) / 1000 };
          out.push(mo);
          return;
        }
        if (rel?.external) warnings.add('외부 파일에 연결된 비디오/오디오는 함께 열 수 없습니다');
      }
      const sr = child(bf, 'srcRect');
      const o = { id: uid(), type: 'image', ...rect, media: mid, name: attr(cNvPr, 'name') };
      if (attr(cNvPr, 'id')) o.spid = num(attr(cNvPr, 'id'));
      if (sr) o.crop = { l: num(attr(sr, 'l')) / 100000, t: num(attr(sr, 't')) / 100000, r: num(attr(sr, 'r')) / 100000, b: num(attr(sr, 'b')) / 100000 };
      if (attr(cNvPr, 'descr')) o.alt = attr(cNvPr, 'descr');
      if (attr(descendants(cNvPr, 'decorative')[0], 'val') === '1') o.decorative = true;
      const g = readGeom(spPr);
      if (g.shape && g.shape !== 'rect') { o.shape = g.shape; if (g.adj) o.adj = g.adj; }
      const ln = readLine(child(spPr, 'ln'), ctx);
      if (ln) o.line = ln;
      const am = child(blip, 'alphaModFix');
      if (am) o.alpha = num(attr(am, 'amt'), 100000) / 100000;
      if (child(blip, 'grayscl')) o.gray = true;
      const duo = child(blip, 'duotone');
      if (duo) { const cs = duo.children.map((k) => readColor({ children: [k] }, ctx)).filter(Boolean); if (cs.length === 2) o.duotone = cs; }
      const lum = child(blip, 'lum');
      if (lum) { if (attr(lum, 'bright')) o.bright = num(attr(lum, 'bright')) / 100000; if (attr(lum, 'contrast')) o.contrast = num(attr(lum, 'contrast')) / 100000; }
      // 꾸밈 효과: WIPOINT 확장 (원본 그림 포함) 또는 PowerPoint a14:imgEffect (원본은 HD Photo 라 효과 이름만)
      const wpArt = descendants(blip, 'art')[0];
      if (wpArt) {
        try { o.art = JSON.parse(attr(wpArt, 'json')); } catch { /* 무시 */ }
        const om = attr(wpArt, 'embed') ? ctx.media(attr(wpArt, 'embed')) : null;
        if (om) o.artOrig = om;
      } else {
        const ie = descendants(blip, 'imgEffect')[0]?.children?.find((x) => ART_BY_XML[x.name]);
        if (ie) o.art = { effect: ART_BY_XML[ie.name] };
      }
      { const fx = fxFromXml(spPr, { child, attr, readColor: (e) => readColor(e, ctx) }); if (fx) o.fx = fx; }
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
      const fid = num(attr(child(child(c, 'nvGraphicFramePr'), 'cNvPr'), 'id'), 0);
      if (uri.endsWith('/table')) { const t = readTable(child(gd, 'tbl'), ctx, rect); if (fid) t.spid = fid; out.push(t); return; }
      if (uri.endsWith('/chart')) {
        const ch = child(gd, 'chart');
        const rel = relsOf(ctx.part).get(attr(ch, 'id'));
        const chart = rel ? readChart(rel.target) : null;
        if (chart) { const chartRaw = chartRawOf(rel.target); out.push({ id: uid(), type: 'chart', ...rect, chart, ...(fid ? { spid: fid } : {}), ...(chartRaw ? { chartRaw } : {}) }); }
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
          // 데이터 모델의 글 계층 + 레이아웃 → WIPOINT SmartArt (텍스트 창으로 고칠 수 있게)
          if (dmRel && sub.length) {
            try {
              const loRel = relIds ? relsOf(ctx.part).get(attr(relIds, 'lo')) : null;
              const uniq = loRel && has(loRel.target) ? attr(xmlOf(loRel.target), 'uniqueId') ?? '' : '';
              const kind = DGM_LAYOUT[uniq.split('/').pop()] ?? null;
              const items = dgmItems(xmlOf(dmRel.target));
              if (kind && items.length) {
                for (const o of sub) o.smart = kind;
                sub[0].smartItems = items;
                sub[0].smartMulti = true;
              }
            } catch { /* 그림 그대로 */ }
          }
        } else warnings.add('SmartArt 일부는 그림으로 저장되지 않아 표시하지 못했습니다');
        return;
      }
      // OLE 개체 (엑셀 표 · 문서 · 수식 편집기 …): 미리 보기 그림 + 포함된 원본 (다시 저장하면 PowerPoint 에서 두 번 눌러 편집)
      const oles = descendants(gd, 'oleObj');
      const ole = oles.find((e) => child(e, 'pic')) ?? oles[0];
      if (ole) {
        let mid = null;
        const pic0 = child(ole, 'pic');
        if (pic0) { const tmp = []; readElement(pic0, ctx, layout, null, tmp, opts); mid = tmp[0]?.media ?? null; }
        if (!mid) mid = vmlImage(ctx.part, attr(ole, 'spid'));
        const binRel = ctx.rel?.(rid(ole));
        const o = { id: uid(), type: 'image', ...rect, media: mid, name: attr(child(child(c, 'nvGraphicFramePr'), 'cNvPr'), 'name') };
        o.ole = { progId: attr(ole, 'progId') ?? '', name: attr(ole, 'name') ?? '', ...(attr(ole, 'showAsIcon') === '1' ? { icon: true } : {}), ...(attr(ole, 'imgW') ? { imgW: num(attr(ole, 'imgW')), imgH: num(attr(ole, 'imgH')) } : {}) };
        if (binRel?.external) o.ole.link = binRel.target;
        else if (binRel) { const b = ctx.binMedia?.(binRel.target); if (b) { o.ole.bin = b; o.ole.ext = binRel.target.split('.').pop().toLowerCase(); o.ole.pkg = /\/package$/.test(binRel.type); } }
        if (!mid) warnings.add('미리 보기 그림이 없는 OLE 개체는 빈 칸으로 표시합니다');
        out.push(o);
        return;
      }
      // 그 밖의 개체: 대체 그림
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
          for (const [tag, side] of [['lnL', 'l'], ['lnR', 'r'], ['lnT', 't'], ['lnB', 'b'], ['lnTlToBr', 'dd'], ['lnBlToTr', 'du']]) {
            const ln = child(tcPr, tag);
            if (!ln) continue;
            const l = readLine(ln, ctx);
            if (l) borders[side] = { color: l.color, width: l.width };
            else if (l === null && side !== 'dd' && side !== 'du') borders[side] = { color: 'rgba(0,0,0,0)', width: 0 };
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
    // 차트 영역 그림 채우기 (미디어 id 는 파일마다 새로 정해짐)
    const bgBlip = child(child(child(descendants(x, 'chartSpace')[0] ?? x, 'spPr'), 'blipFill'), 'blip');
    const bgRel = bgBlip ? relsOf(path).get(attr(bgBlip, 'embed')) : null;
    const bgMedia = bgRel && !bgRel.external ? loadMedia(bgRel.target) : null;
    if (ext) { try { const j = JSON.parse(attr(ext, 'json')); if (bgMedia) j.bgMedia = bgMedia; else delete j.bgMedia; return j; } catch { /* 아래로 */ } }
    const plot = descendants(x, 'plotArea')[0];
    const typeEls = (plot?.children ?? []).filter((c) => /Chart$/.test(c.name));
    if (!typeEls.length) return null;
    const barDir = (el) => (attr(child(el, 'barDir'), 'val') === 'bar' ? 'bar' : 'col');
    const KIND = { barChart: barDir, bar3DChart: barDir, lineChart: () => 'line', line3DChart: () => 'line', stockChart: () => 'line', areaChart: () => 'area', area3DChart: () => 'area', pieChart: () => 'pie', pie3DChart: () => 'pie', ofPieChart: () => 'pie', doughnutChart: () => 'doughnut', scatterChart: () => 'scatter', radarChart: () => 'radar', bubbleChart: () => 'bubble', surfaceChart: () => 'col', surface3DChart: () => 'col' };
    const kindOf = (el) => (KIND[el.name] ?? (() => 'col'))(el);
    const vals = (el) => {
      const cache = descendants(el, 'numCache')[0] ?? descendants(el, 'numLit')[0] ?? descendants(el, 'strCache')[0] ?? descendants(el, 'strLit')[0];
      if (!cache) return [];
      const n = num(attr(child(cache, 'ptCount'), 'val'), 0);
      const arr = Array.from({ length: n }, () => null);
      for (const pt of kids(cache, 'pt')) arr[num(attr(pt, 'idx'))] = child(pt, 'v')?.text ?? '';
      return arr;
    };
    // 테마 색(@accent1:lm75 …)도 그대로 (차트 계열 색은 보통 테마 색)
    const fillOf = (sp) => readColor(child(sp, 'solidFill'), { clrMap: null }) ?? null;
    const lnOf = (sp) => { const ln = child(sp, 'ln'); return ln ? { color: readColor(child(ln, 'solidFill'), { clrMap: null }), w: attr(ln, 'w'), dash: attr(child(ln, 'prstDash'), 'val'), none: !!child(ln, 'noFill') } : null; };
    const dlOf = (dl) => {
      if (!dl || bool(attr(child(dl, 'delete'), 'val'))) return null;
      const on = (k) => bool(attr(child(dl, k), 'val'));
      if (!(on('showVal') || on('showPercent') || on('showCatName') || on('showSerName'))) return null;
      const pos = attr(child(dl, 'dLblPos'), 'val');
      const nf = child(dl, 'numFmt');
      const fmt = nf && attr(nf, 'sourceLinked') !== '1' && attr(nf, 'formatCode') && attr(nf, 'formatCode') !== 'General' ? attr(nf, 'formatCode') : null;
      return { ...(pos ? { pos } : {}), val: on('showVal'), cat: on('showCatName'), ser: on('showSerName'), pct: on('showPercent'), ...(fmt ? { fmt } : {}) };
    };
    const axById = new Map();
    for (const a of (plot?.children ?? []).filter((c) => /Ax$/.test(c.name))) axById.set(attr(child(a, 'axId'), 'val'), a);
    const typeIds = (el) => kids(el, 'axId').map((a) => attr(a, 'val'));
    const valAxOf = (el) => typeIds(el).map((id) => axById.get(id)).find((a) => a?.name === 'valAx' && ['l', 'r'].includes(attr(child(a, 'axPos'), 'val')) || a?.name === 'valAx' && kindOf(el) === 'bar') ?? typeIds(el).map((id) => axById.get(id)).find((a) => a?.name === 'valAx');
    const firstKind = kindOf(typeEls[0]);
    const kinds = new Set(typeEls.map(kindOf));
    const combo = typeEls.length > 1 && kinds.size > 1;
    const primaryVal = valAxOf(typeEls[0]);
    const series = [];
    let grouping = null;
    let variant;
    let unnamed = false;
    let srcFmt = null;
    for (const typeEl of typeEls) {
      const k = kindOf(typeEl);
      const g = attr(child(typeEl, 'grouping'), 'val');
      const va = valAxOf(typeEl);
      const secondary = !!(va && primaryVal && va !== primaryVal && attr(child(va, 'axPos'), 'val') === 'r');
      if (g && (k === 'col' || k === 'bar' || !grouping)) grouping = g === 'percentStacked' ? 'percent' : g === 'stacked' ? 'stacked' : grouping ?? null;
      if (k === 'scatter' && attr(child(typeEl, 'scatterStyle'), 'val')?.startsWith('smooth')) variant = 'smooth';
      if (k === 'radar') { const st = attr(child(typeEl, 'radarStyle'), 'val'); if (st === 'filled') variant = 'filled'; }
      for (const sEl of kids(typeEl, 'ser')) {
        const i = series.length;
        const tx = child(sEl, 'tx');
        const name = tx ? (child(tx, 'v')?.text ?? vals(tx)[0] ?? `계열 ${i + 1}`) : `계열 ${i + 1}`;
        if (!tx) unnamed = true;
        const fc = descendants(child(sEl, 'val') ?? child(sEl, 'yVal'), 'formatCode')[0]?.text;
        if (fc && fc !== 'General' && !srcFmt) srcFmt = fc;
        const sp = child(sEl, 'spPr');
        const ln = lnOf(sp);
        const lineLike = k === 'line' || k === 'scatter' || (k === 'radar' && variant !== 'filled');
        const color = lineLike ? (ln?.color ?? null) : fillOf(sp);
        const o = { name, vals: vals(child(sEl, 'val') ?? child(sEl, 'yVal')).map((v) => Number(v) || 0), ...(color ? { color } : {}), _cats: vals(child(sEl, 'cat') ?? child(sEl, 'xVal')) };
        if (combo) o.type = k === 'bar' ? 'col' : k;
        if (secondary) o.axis = 'secondary';
        if (k === 'bubble') o.sizes = vals(child(sEl, 'bubbleSize')).map((v) => Number(v) || 0);
        if (lineLike && ln?.w && Number(ln.w) !== 28575) o.lineWidth = Math.round((Number(ln.w) / 12700) * 4) / 4;
        if (ln?.dash && ln.dash !== 'solid') o.dash = { dash: 'dash', sysDash: 'dash', sysDot: 'dot', dashDot: 'dashDot', lgDash: 'longDash' }[ln.dash] ?? 'dash';
        if (k === 'scatter' && ln && !ln.none && !variant) variant = 'line';
        const mk = attr(child(child(sEl, 'marker'), 'symbol'), 'val');
        if (k === 'line' && mk && mk !== 'none') { o.marker = MARKER_XML[mk] ?? 'circle'; }
        if (k === 'radar' && mk && mk !== 'none' && !variant) variant = 'marker';
        if (bool(attr(child(sEl, 'smooth'), 'val')) && k === 'line') o.smooth = true;
        const pc = {};
        const ex = {};
        for (const d of kids(sEl, 'dPt')) { const di = num(attr(child(d, 'idx'), 'val')); const c2 = fillOf(child(d, 'spPr')); if (c2) pc[di] = c2; const e = num(attr(child(d, 'explosion'), 'val'), 0); if (e) ex[di] = e; }
        if (Object.keys(pc).length) o.pointColors = pc;
        if (Object.keys(ex).length) o.explodeAt = ex;
        const sdl = dlOf(child(sEl, 'dLbls'));
        if (sdl) o.labels = sdl;
        const tr = child(sEl, 'trendline');
        if (tr) {
          const tt = attr(child(tr, 'trendlineType'), 'val');
          const t2 = Object.entries(TREND_XML).find(([, v]) => v === tt)?.[0];
          if (t2) o.trend = { type: t2, ...(t2 === 'poly' ? { order: num(attr(child(tr, 'order'), 'val'), 2) } : {}), ...(t2 === 'movingAvg' ? { period: num(attr(child(tr, 'period'), 'val'), 2) } : {}), ...(child(tr, 'forward') ? { forward: num(attr(child(tr, 'forward'), 'val'), 0) } : {}), eq: bool(attr(child(tr, 'dispEq'), 'val')) || undefined, r2: bool(attr(child(tr, 'dispRSqr'), 'val')) || undefined, ...(child(tr, 'name') ? { name: child(tr, 'name').text } : {}) };
        }
        const eb = child(sEl, 'errBars');
        if (eb) {
          const et = attr(child(eb, 'errValType'), 'val');
          const e2 = Object.entries(ERR_XML).find(([, v]) => v === et)?.[0];
          if (e2) o.err = { type: e2, ...(child(eb, 'val') ? { val: Number(attr(child(eb, 'val'), 'val')) } : {}) };
        }
        if ((k === 'pie' || k === 'doughnut') && num(attr(child(sEl, 'explosion'), 'val'), 0)) variant = variant ?? 'explode';
        series.push(o);
      }
    }
    const cats = (series.find((s) => s._cats.length)?._cats ?? []).map((v) => v ?? '');
    for (const s of series) delete s._cats;
    let kind = combo ? 'combo' : firstKind;
    if (combo && [...kinds].every((k) => k === 'col' || k === 'bar')) kind = firstKind;
    if (firstKind === 'line' && !combo && series.some((s) => s.marker) && series.every((s) => s.marker === 'circle' || !s.marker)) { variant = 'markers'; for (const s of series) delete s.marker; }
    const chartEl = descendants(x, 'chart')[0];
    const titleEl = child(chartEl, 'title');
    const autoDeleted = bool(attr(child(chartEl, 'autoTitleDeleted'), 'val'));
    // 제목 글이 없으면: 계열이 하나일 때 계열 이름이 자동 제목 (PowerPoint 와 같음)
    const title = descendants(titleEl, 't').map((t) => t.text).join('') || (unnamed ? '' : (titleEl || !autoDeleted) && series.length === 1 && !['pie', 'doughnut'].includes(kind) ? series[0].name : (titleEl && series.length === 1 ? series[0].name : ''));
    const tsz = attr(descendants(titleEl, 'defRPr')[0], 'sz') ?? attr(descendants(titleEl, 'rPr')[0], 'sz');
    const legend = attr(child(descendants(x, 'legend')[0], 'legendPos'), 'val');
    const dl = dlOf(kids(typeEls[0], 'dLbls')[0]);
    const out = { kind, title, cats: cats.length ? cats : series[0]?.vals.map((_, i) => String(i + 1)) ?? [], series, legend: legend && ['t', 'b', 'l', 'r', 'tr'].includes(legend) ? legend : null, labels: dl ?? false };
    // 3차원 (bar3DChart 등 + c:view3D)
    if (typeEls.length === 1 && /3DChart$/.test(typeEls[0].name) && ['col', 'bar', 'line', 'area', 'pie'].includes(kind)) {
      out.d3 = true;
      const v = descendants(chartEl, 'view3D')[0];
      const n = (k) => (child(v, k) ? num(attr(child(v, k), 'val')) : null);
      out.view3d = Object.fromEntries([['rotX', n('rotX')], ['rotY', n('rotY')], ['depth', n('depthPercent')], ['persp', n('perspective')]].filter(([, x]) => x != null));
      if ((kind === 'col' || kind === 'bar') && attr(child(typeEls[0], 'grouping'), 'val') === 'standard') out.variant3d = 'full';
    }
    if (grouping) out.grouping = grouping;
    if (variant) out.variant = variant;
    if (tsz && Number(tsz) !== 1862) out.titleSize = Number(tsz) / 100;
    // 축
    const axisOpts = (a) => {
      if (!a) return undefined;
      const o = {};
      const sc = child(a, 'scaling');
      if (attr(child(sc, 'orientation'), 'val') === 'maxMin') o.reverse = true;
      if (child(sc, 'min')) o.min = Number(attr(child(sc, 'min'), 'val'));
      if (child(sc, 'max')) o.max = Number(attr(child(sc, 'max'), 'val'));
      if (child(a, 'majorUnit')) o.major = Number(attr(child(a, 'majorUnit'), 'val'));
      if (bool(attr(child(a, 'delete'), 'val'))) o.hide = true;
      const nf = child(a, 'numFmt');
      if (nf && attr(nf, 'sourceLinked') === '0' && attr(nf, 'formatCode') && attr(nf, 'formatCode') !== 'General') o.fmt = attr(nf, 'formatCode');
      const t = descendants(child(a, 'title'), 't').map((e) => e.text).join('');
      if (t) o.title = t;
      const asz = attr(descendants(child(a, 'txPr'), 'defRPr')[0], 'sz');
      if (asz) o.size = Number(asz) / 100;
      return Object.keys(o).length ? o : undefined;
    };
    if (!['pie', 'doughnut'].includes(kind)) {
      const ids = typeIds(typeEls[0]).map((id) => axById.get(id)).filter(Boolean);
      const catAx = ids.find((a) => a.name === 'catAx' || a.name === 'dateAx') ?? ids.find((a) => a !== primaryVal);
      const y2 = [...axById.values()].find((a) => a.name === 'valAx' && a !== primaryVal && attr(child(a, 'axPos'), 'val') === 'r');
      const axes = { x: axisOpts(catAx), y: axisOpts(primaryVal), ...(y2 && series.some((s2) => s2.axis) ? { y2: axisOpts(y2) } : {}) };
      // 축 서식이 '원본에 연결' 이면 값의 서식 코드 (#,##0 등) 를 축 · 레이블에 씀
      const vnf = child(primaryVal, 'numFmt');
      if (srcFmt && !axes.y?.fmt && (!vnf || attr(vnf, 'sourceLinked') !== '0')) axes.y = { ...(axes.y ?? {}), fmt: srcFmt };
      for (const k of Object.keys(axes)) if (!axes[k]) delete axes[k];
      if (Object.keys(axes).length) out.axes = axes;
      const grid = { major: !!child(primaryVal, 'majorGridlines'), ...(child(primaryVal, 'minorGridlines') ? { minor: true } : {}), ...(catAx && child(catAx, 'majorGridlines') ? { cat: true } : {}) };
      if (!grid.major || grid.minor || grid.cat) out.grid = grid;
    }
    const bar = typeEls.find((e) => e.name === 'barChart' || e.name === 'bar3DChart');
    if (bar) {
      const gw = attr(child(bar, 'gapWidth'), 'val');
      const ov = attr(child(bar, 'overlap'), 'val');
      if (gw != null) out.gap = Number(gw);
      if (ov != null) out.overlap = Number(ov);
    }
    const dn = typeEls.find((e) => e.name === 'doughnutChart');
    if (dn && child(dn, 'holeSize')) out.hole = Number(attr(child(dn, 'holeSize'), 'val'));
    const fsa = typeEls.map((e) => child(e, 'firstSliceAng')).find(Boolean);
    if (fsa && Number(attr(fsa, 'val'))) out.firstAngle = Number(attr(fsa, 'val'));
    if (bool(attr(child(typeEls[0], 'varyColors'), 'val')) && !['pie', 'doughnut'].includes(kind)) out.vary = true;
    const dt = child(plot, 'dTable');
    if (dt) out.dataTable = bool(attr(child(dt, 'showKeys'), 'val')) ? true : 'nokeys';
    const pbg = fillOf(child(plot, 'spPr'));
    if (pbg) out.plotBg = pbg;
    const cs = descendants(x, 'chartSpace')[0] ?? x;
    const cbg = fillOf(child(cs, 'spPr'));
    if (cbg) out.bg = cbg;
    if (bgMedia) out.bgMedia = bgMedia;
    const fsz = attr(descendants(child(cs, 'txPr'), 'defRPr')[0], 'sz');
    // 차트 글자 크기: 파일 값 그대로 (없으면 PowerPoint 기본 10pt)
    out.font = { ...(out.font ?? {}), size: fsz ? Number(fsz) / 100 : 10, exact: true };
    return out;
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
      const tpl = idx >= 1001 ? TI(ctx)?.bgStyles?.[idx - 1001] : TI(ctx)?.fillStyles?.[idx - 1];
      if (tpl && tpl.name !== 'solidFill') { const f = readFill({ children: [tpl] }, tplCtx(ctx), col); if (f) return f; }
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

  /** 원래 XML 로 다시 쓸 수 있는 개체인지 + 관계(그림 · 링크 · 슬라이드) 목록 */
  const slide0Ns = new Map();
  const RAW_TAGS = new Set(['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame']);
  /** 차트 파트 원본 (손대지 않으면 그대로 다시 씀): XML + 관계 (스타일 · 색 XML, 데이터 통합 문서 · 그림은 미디어로) */
  // PowerPoint 2007 이 만든 차트는 데이터 레이블 표시 요소를 생략함 (없음 = 끔). 다른 앱이 같은 뜻으로 읽도록 명시해서 저장
  const appXml = has('docProps/app.xml') ? textOf(files['docProps/app.xml']) : '';
  const mso2007 = /<Application>Microsoft[^<]*<\/Application>/.test(appXml) && /<AppVersion>12\./.test(appXml);
  const ctMap = new Map(kids(xmlOf('[Content_Types].xml'), 'Override').map((o) => [String(attr(o, 'PartName') ?? '').replace(/^\//, ''), attr(o, 'ContentType')]));
  function chartRawOf(path) {
    if (!has(path)) return null;
    const rels = [];
    for (const [id, r] of relsOf(path)) {
      if (r.external) { rels.push({ id, type: r.type, target: r.target, external: true }); continue; }
      if (!has(r.target)) return null;
      const name = r.target.split('/').pop();
      if (/\.xml$/i.test(name)) {
        const ct = ctMap.get(r.target);
        if (!ct || has(relsPath(r.target))) return null;
        rels.push({ id, type: r.type, xml: textOf(files[r.target]), ct, name });
      } else {
        const mid = r.type.endsWith('/image') ? loadMedia(r.target) : loadMedia(r.target, 'bin');
        if (!mid) return null;
        rels.push({ id, type: r.type, media: mid, name });
      }
    }
    let xml = textOf(files[path]);
    if (mso2007) xml = explicitDLbls(xml);
    return { xml, rels };
  }

  function rawObject(el, xml, made, part) {
    if (!RAW_TAGS.has(el.name)) return null;
    if (made.some((o) => o.ph || ['media', 'chart', 'equation', 'zoom'].includes(o.type) || o.ole || o.smart)) return null;
    if (el.name === 'graphicFrame' && !/drawingml\/2006\/table"/.test(xml)) return null;
    if (el.name !== 'grpSp' && made.length !== 1) return null;
    if (/<(?:\w+:)?(?:contentPart|AlternateContent)\b/.test(xml)) return null;
    const relMap = {};
    for (const m of xml.matchAll(/\br:(\w+)="([^"]*)"/g)) {
      const id = m[2];
      if (!id || relMap[id]) continue;
      const r = relsOf(part).get(id);
      if (!r) return null;
      if (r.type.endsWith('/image') && !r.external) { const mid = loadMedia(r.target); if (!mid) return null; relMap[id] = { kind: 'image', media: mid }; } else if (r.type.endsWith('/hyperlink')) relMap[id] = { kind: 'link', url: r.target, external: r.external };
      else if (r.type.endsWith('/slide')) relMap[id] = { kind: 'slide', path: r.target };
      else return null;
    }
    const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)].map((m) => Number(m[1]));
    const out = { xml, rels: relMap, ids, n: made.length };
    if (el.name === 'grpSp') {
      out.group = true;
      const gx = readXfrm(child(child(el, 'grpSpPr'), 'xfrm'));
      if (!gx) return null;
      out.g = { x: gx.x, y: gx.y, w: gx.w, h: gx.h };
    }
    return out;
  }

  const slides = [];
  const slideByRid = new Map();
  const pendingZooms = [];
  for (const sid of descendants(child(presXml, 'sldIdLst'), 'sldId')) {
    const rel = presRels.get(rid(sid));
    if (!rel || !has(rel.target)) continue;
    const path = rel.target;
    const src = textOf(files[path]);
    const x = parseXml(src, true);
    const rels = relsOf(path);
    const lRel = [...rels.values()].find((r) => r.type.endsWith('/slideLayout'));
    const layout = lRel ? loadLayout(lRel.target) : null;
    const clrOvr = descendants(child(x, 'clrMapOvr'), 'overrideClrMapping')[0];
    const clrMap = clrOvr ? readClrMap(clrOvr) : layout?.clrMap ?? layout?.master.clrMap;
    const ctx = { ...partCtx(path, clrMap), part: path, ti: layout?.master.ti };
    // WIPOINT 확장 (우리 앱이 쓴 파일이면 전환 · 애니메이션 · 레이아웃을 그대로)
    let meta = null;
    const wpEl = descendants(child(x, 'extLst'), 'slide').find((e) => attr(e, 'json'));
    if (wpEl) { try { meta = JSON.parse(attr(wpEl, 'json')); } catch { meta = null; } }
    const objects = [];
    {
      // 개체마다 원래 XML 을 함께 보관 (고치지 않은 개체는 저장할 때 그대로 → PowerPoint 고유 서식 · 효과가 사라지지 않음)
      const tree = descendants(child(x, 'cSld'), 'spTree')[0];
      const rootNs = x.attrs ?? {};
      const nsOk = (!rootNs['xmlns:a'] || rootNs['xmlns:a'] === NS_A) && (!rootNs['xmlns:p'] || rootNs['xmlns:p'] === NS_P) && (!rootNs['xmlns:r'] || rootNs['xmlns:r'] === NS_R);
      for (const c of tree?.children ?? []) {
        const before = objects.length;
        try { readElement(c, ctx, layout, null, objects, { skipFooters: !!(meta && presMeta?.footer) }); } catch (e) { warnings.add(`일부 개체를 읽지 못했습니다 (${c.name}): ${e.message}`); continue; }
        const made = objects.slice(before);
        if (!nsOk || !made.length || c.s == null || c.e == null) continue;
        const raw = rawObject(c, src.slice(c.s, c.e), made, path);
        if (raw) made[0].raw = raw;
      }
      slide0Ns.set(path, Object.fromEntries(Object.entries(rootNs).filter(([k]) => k.startsWith('xmlns:') || k === 'mc:Ignorable')));
    }
    // 배경 장식: 마스터(레이아웃이 숨기지 않으면) + 레이아웃의 개체 틀 아닌 도형
    const bgObjects = [];
    if (layout) {
      const showMaster = layout.showMasterSp && attr(x, 'showMasterSp') !== '0';
      if (showMaster) readShapeTree(layout.master.spTree, { ...layout.master.ctx, clrMap, part: layout.master.path }, null, null, bgObjects, { skipPh: true, isMaster: true });
      readShapeTree(layout.spTree, { ...layout.ctx, clrMap, part: layout.path }, null, null, bgObjects, { skipPh: true, isMaster: true });
      for (const o of bgObjects) { o.decor = true; o.id = `bg-${o.id}`; delete o.grp; }
    }
    let bg = readBg(child(child(x, 'cSld'), 'bg'), ctx);
    if (bg === undefined && layout) bg = readBg(layout.bg, withMap(layout.ctx, clrMap));
    if (bg === undefined && layout) bg = readBg(layout.master.bg, withMap(layout.master.ctx, clrMap));
    // 메모
    let notes = '';
    const nRel = [...rels.values()].find((r) => r.type.endsWith('/notesSlide'));
    if (nRel && has(nRel.target)) {
      const nx = xmlOf(nRel.target);
      const body = (descendants(nx, 'sp')).find((sp) => { const k = phKey(sp); return k && k.type === 'body'; });
      if (body) notes = kids(child(body, 'txBody'), 'p').map((p) => descendants(p, 't').map((t) => t.text).join('')).join('\n').replace(/\n+$/, '');
    }
    const slide = {
      id: uid('s'), layout: meta?.layout ?? LAYOUT_KIND[layout?.type] ?? (layout && !meta ? foreignLayout(layout, bgObjects) : null) ?? 'blank', bg: bg ?? null, hidden: attr(x, 'show') === '0',
      transition: null, notes, objects, anims: [],
    };
    // 레이아웃 이름이 사용자 지정이면 배경 장식은 원본대로
    // 우리가 쓴 파일의 테마 장식은 테마에서 다시 만들어지므로 원본 장식만 담음
    if (meta?.bgObjects || !meta) slide.bgObjects = bgObjects;
    if (meta?.hideDecor) slide.hideDecor = true;
    if (!meta && layout && layout.master.path !== firstMaster) {
      const d = foreignDesign(layout.master.path);
      if (d) slide.design = d.id;
    }
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
    // 메모: WIPOINT 확장 → 이전 형식(p:cmLst) → 새 형식(p188 스레드 메모)
    if (meta?.comments) slide.comments = meta.comments;
    else {
      const cmts = readComments(path);
      if (cmts.length) slide.comments = cmts;
    }
    if (meta) {
      slide.transition = meta.transition ?? null;
      const map = new Map(Object.entries(meta.ids ?? {}).map(([sp, oldId]) => [oldId, spidToObj.get(sp)]));
      slide.anims = (meta.anims ?? []).map((a) => ({ ...a, obj: map.get(a.obj), trigger: a.trigger ? map.get(a.trigger) : undefined })).filter((a) => a.obj);
      for (const a of slide.anims) if (a.trigger === undefined) delete a.trigger;
      if (meta.section) slide.section = meta.section;
      if (meta.design) slide.design = meta.design;
      for (const [sp, pr] of Object.entries(meta.props ?? {})) { const o = objects.find((q) => q.id === spidToObj.get(sp)); if (o) Object.assign(o, pr); }
      if (meta.inkRec) slide.inkRec = meta.inkRec;
      // 확대/축소 개체 (대상 슬라이드는 모두 읽은 뒤 id 로 바꿈)
      for (const [sp, z] of Object.entries(meta.zooms ?? {})) {
        const o = objects.find((q) => q.id === spidToObj.get(sp));
        if (!o) continue;
        for (const k of ['shape', 'fill', 'line', 'text', 'link', 'adj']) delete o[k];
        Object.assign(o, { type: 'zoom', zoom: { ...z } });
        pendingZooms.push(o);
      }
    } else {
      slide.transition = readTransition(x);
      slide.anims = readTiming(x, spidToObj);
    }
    {
      // PowerPoint 의 애니메이션 · 전환 원본 (고치지 않으면 그대로 저장 → 반복 · 되감기 · 부드럽게 시작 등 세부 설정 유지)
      const rawPart = (el) => {
        if (!el || el.s == null || el.e == null) return null;
        const xml = src.slice(el.s, el.e);
        const relMap = {};
        for (const m of xml.matchAll(/\br:(\w+)="([^"]*)"/g)) {
          const r = m[2] ? rels.get(m[2]) : null;
          if (!m[2] || relMap[m[2]]) continue;
          if (!r || r.external) return null;
          const mid = /\/(audio|media)$/.test(r.type) ? loadMedia(r.target, true) : null;
          if (!mid) return null;
          relMap[m[2]] = { kind: 'audio', media: mid, name: r.target.split('/').pop() };
        }
        return { xml, rels: relMap };
      };
      const tEl = x.children.find((c) => c.name === 'transition' || (c.name === 'AlternateContent' && descendants(c, 'transition').length));
      const tim = child(x, 'timing');
      const rt = rawPart(tEl);
      if (rt) slide.rawTransition = rt;
      const rtim = rawPart(tim);
      if (rtim) {
        rtim.spids = [...new Set([...rtim.xml.matchAll(/\bspid="(\d+)"/g)].map((m) => Number(m[1])))];
        slide.rawTiming = rtim;
      }
    }
    if (!slide.section && sectionOf.has(attr(sid, 'id'))) {
      const name = sectionOf.get(attr(sid, 'id'));
      if (!firstOfSection.has(name)) { firstOfSection.add(name); slide.section = name; }
    }
    for (const o of objects) delete o.name;
    if (slide0Ns.has(path)) slide.rawNs = slide0Ns.get(path);
    // 이 슬라이드의 실제 색 매핑 (원래 XML 의 tx1 · bg1 이 같은 색이 되도록 저장할 때 씀)
    if (clrMap && !isStdMap(clrMap)) slide.clrMap = Object.fromEntries(Object.entries(clrMap).filter(([k]) => STD_CLRMAP[k]));
    slides.push(slide);
    slideByRid.set(rid(sid), slide.id);
  }

  for (const o of pendingZooms) o.zoom.slide = slides[o.zoom.slide]?.id ?? null;

  /** 메모 읽기 (이전 형식 cmLst + commentAuthors / 새 형식 modernComment + authors) */
  function readComments(slidePath) {
    const out = [];
    for (const r of relsOf(slidePath).values()) {
      if (!/\/comments$/.test(r.type) || !has(r.target)) continue;
      const cx = xmlOf(r.target);
      const modern = /2018\/10/.test(r.type);
      if (!modern) {
        for (const cm of kids(cx, 'cm')) {
          const pos = child(cm, 'pos');
          const text = child(cm, 'text')?.text ?? '';
          const c = { id: uid('c'), author: legacyAuthors.get(attr(cm, 'authorId')) ?? '사용자', text, at: attr(cm, 'dt') ?? '', x: Math.round(num(attr(pos, 'x'), 0) / 6), y: Math.round(num(attr(pos, 'y'), 0) / 6), replies: [] };
          if (text.startsWith('↳ ') && out.length) out[out.length - 1].replies.push({ author: c.author, text: text.slice(2), at: c.at });
          else out.push(c);
        }
      } else {
        const tx = (el) => descendants(child(el, 'txBody'), 'p').map((p) => descendants(p, 't').map((t) => t.text ?? '').join('')).join('\n');
        for (const cm of kids(cx, 'cm')) {
          const pos = child(cm, 'pos');
          out.push({ id: uid('c'), author: modernAuthors.get(attr(cm, 'authorId')) ?? '사용자', text: tx(cm), at: attr(cm, 'created') ?? '', x: Math.round(PX(num(attr(pos, 'x'), 0))), y: Math.round(PX(num(attr(pos, 'y'), 0))), replies: kids(child(cm, 'replyLst'), 'reply').map((rp) => ({ author: modernAuthors.get(attr(rp, 'authorId')) ?? '', text: tx(rp), at: attr(rp, 'created') ?? '' })), ...(attr(cm, 'status') === 'resolved' ? { done: true } : {}) });
        }
      }
    }
    return out;
  }

  function readTransition(x) {
    const t = descendants(x, 'transition')[0];
    if (!t) return null;
    const el = t.children.find((c) => c.name !== 'sndAc' && c.name !== 'extLst');
    let name = el ? TRANS_NAME[el.name] ?? 'fade' : 'none';
    if (el?.name === 'prstTrans') name = PRST_TRANS[attr(el, 'prst')] ?? 'fade';
    if (el?.name === 'prism') name = attr(el, 'isContent') === '1' ? (attr(el, 'isInverted') === '1' ? 'orbit' : 'rotate') : attr(el, 'isInverted') === '1' ? 'box' : 'cube';
    const spd = attr(t, 'spd');
    const dur = attr(t, 'dur') ? num(attr(t, 'dur')) / 1000 : spd === 'fast' ? 0.5 : spd === 'slow' ? 1 : 0.75;
    const dirAttr = attr(el, 'dir');
    const out = { type: name, dur };
    if (dirAttr && DIR_FROM[dirAttr]) out.dir = DIR_FROM[dirAttr];
    // 가로/세로 효과 · 계단 방향 · 나누기 방향
    if (dirAttr === 'vert' && ['blinds', 'checker', 'comb', 'randomBars'].includes(name)) out.dir = 'l';
    if (name === 'strips' && dirAttr) out.dir = { ru: 'b', rd: 't', ld: 'l', lu: 'r' }[dirAttr] ?? 'b';
    if (name === 'split' && attr(el, 'orient') === 'horz') out.dir = 'l';
    if (el?.name === 'prstTrans' && attr(el, 'invX') === '1') out.dir = 'l';
    const snd = descendants(t, 'snd')[0];
    if (snd && SOUND_BY_FILE[attr(snd, 'name')]) out.sound = SOUND_BY_FILE[attr(snd, 'name')];
    if (attr(t, 'advClick') === '0') out.advClick = false;
    if (attr(t, 'advTm')) out.advAfter = num(attr(t, 'advTm')) / 1000;
    return out;
  }

  function readTiming(x, spidToObj) {
    const timing = child(x, 'timing');
    if (!timing) return [];
    // 기본 순서 (mainSeq) + 트리거 순서 (interactiveSeq: 누를 개체 = stCondLst 의 onClick spTgt)
    const seqs = descendants(timing, 'cTn').filter((c) => ['mainSeq', 'interactiveSeq'].includes(attr(c, 'nodeType')));
    const anims = [];
    for (const seq of seqs) {
    let trigger = null;
    if (attr(seq, 'nodeType') === 'interactiveSeq') {
      const tg = descendants(child(seq, 'stCondLst'), 'spTgt')[0];
      trigger = spidToObj.get(attr(tg, 'spid'));
      if (!trigger) continue;
    }
    for (const eff of descendants(seq, 'cTn').filter((c) => attr(c, 'presetClass'))) {
      const cls = attr(eff, 'presetClass');
      if (!['entr', 'exit', 'emph', 'path'].includes(cls)) continue;
      const spTgt = descendants(eff, 'spTgt')[0];
      const obj = spidToObj.get(attr(spTgt, 'spid'));
      if (!obj) continue;
      const pid = num(attr(eff, 'presetID'));
      const effect = cls === 'path' ? 'pathRight' : PRESET_ANIM[cls]?.[pid] ?? (cls === 'entr' ? 'fade' : cls === 'exit' ? 'fade' : 'pulse');
      const node = attr(eff, 'nodeType');
      const durs = descendants(eff, 'cTn').map((c) => num(attr(c, 'dur'), 0)).filter((d) => d > 1);
      const delay = num(attr(child(child(eff, 'stCondLst'), 'cond'), 'delay'), 0);
      const a = { id: uid('a'), obj, cls, effect, start: node === 'withEffect' ? 'with' : node === 'afterEffect' ? 'after' : 'click', dur: durs.length ? Math.max(...durs) / 1000 : effect === 'appear' || effect === 'disappear' ? 0 : 0.5, delay: delay / 1000, dir: SUB_DIR[num(attr(eff, 'presetSubtype'))] ?? 'b' };
      if (cls === 'path') {
        const mo = descendants(eff, 'animMotion')[0];
        a.motion = (attr(mo, 'path') ?? 'M 0 0 L 0.25 0').replace(/\s*[Ee]\s*$/, '').trim();
        a.effect = Object.keys(MOTION_PATHS).find((k) => MOTION_PATHS[k] === a.motion) ?? 'pathRight';
        if (!durs.length) a.dur = 2;
      }
      const it = child(eff, 'iterate');
      if (it) { a.textUnit = attr(it, 'type') === 'lt' ? 'letter' : 'word'; const pct = child(it, 'tmPct'); if (pct) a.unitDelay = num(attr(pct, 'val'), 10000) / 1000; }
      const st = descendants(child(eff, 'subTnLst'), 'sndTgt')[0];
      if (st && SOUND_BY_FILE[attr(st, 'name')]) a.sound = SOUND_BY_FILE[attr(st, 'name')];
      if (trigger) a.trigger = trigger;
      // 단락별 효과 (txEl pRg): 같은 개체 · 같은 효과가 이어지면 하나로 묶고 byPara
      const isPara = !!descendants(spTgt, 'pRg')[0];
      if (isPara) {
        const prev = anims.find((b) => b.obj === obj && b.cls === cls && b.effect === effect && b.byPara);
        if (prev) continue;
        a.byPara = true;
      }
      if (anims.some((b) => b.obj === obj && b.cls === cls && b.effect === effect && b.start !== 'click' && a.start !== 'click' && (b.trigger ?? null) === (a.trigger ?? null))) continue;
      anims.push(a);
    }
    }
    return anims;
  }

  // 문서 속성
  const core = xmlOf('docProps/core.xml');
  const props = { title: descendants(core, 'title')[0]?.text ?? '', author: descendants(core, 'creator')[0]?.text ?? '', created: descendants(core, 'created')[0]?.text ?? new Date().toISOString() };
  const pres = { version: 1, size, theme, slides, media, props, footer: presMeta?.footer ?? { slideNum: false, date: false, text: '', hideOnTitle: true } };
  if (firstMaster) { const cm = loadMaster(firstMaster).clrMap; if (cm && !isStdMap(cm)) pres.clrMap = Object.fromEntries(Object.entries(cm).filter(([k]) => STD_CLRMAP[k])); }
  for (const [mp, d] of foreignDesigns) if (d) { const cm = loadMaster(mp).clrMap; if (cm && !isStdMap(cm)) d.clrMap = Object.fromEntries(Object.entries(cm).filter(([k]) => STD_CLRMAP[k])); }
  // 원래 테마 · 기본 글 스타일 · 표 스타일을 보관 (다시 저장할 때 그대로)
  {
    const rawTheme = (ti) => {
      if (!ti?.path || !has(ti.path)) return null;
      const xml = textOf(files[ti.path]);
      const relMap = {};
      for (const m of xml.matchAll(/\br:(\w+)="([^"]*)"/g)) {
        const r = relsOf(ti.path).get(m[2]);
        if (!r || r.external || !r.type.endsWith('/image')) return null;
        const mid = loadMedia(r.target);
        if (!mid) return null;
        relMap[m[2]] = { kind: 'image', media: mid };
      }
      // 네임스페이스 접두사가 표준(a:)일 때만
      return /<a:theme\b/.test(xml) ? { xml: xml.replace(/^<\?xml[^>]*\?>\s*/, ''), rels: relMap } : null;
    };
    const tr = rawTheme(themeInfo);
    if (tr) pres.themeRaw = tr;
    for (const [mp, d] of foreignDesigns) if (d) { const t2 = rawTheme(themeInfoFor(mp)); if (t2) d.themeRaw = t2; }
    const pSrc = textOf(files[presPath]);
    const dts = /<p:defaultTextStyle\b[\s\S]*?<\/p:defaultTextStyle>/.exec(pSrc);
    if (dts && /xmlns:a="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/main"/.test(pSrc)) pres.defaultTextStyleRaw = dts[0];
    const tsRel = [...presRels.values()].find((r) => r.type.endsWith('/tableStyles'));
    if (tsRel && has(tsRel.target)) { const t = textOf(files[tsRel.target]); if (/<a:tblStyleLst\b/.test(t)) pres.tableStylesRaw = t.startsWith('<?xml') ? t : XML_HEAD + t; }
  }
  if (presMeta?.masters) {
    const remap = new Map(Object.entries(presMeta.masterMedia ?? {}).map(([id, path]) => [id, has(path) ? loadMedia(path) : null]));
    for (const m of Object.values(presMeta.masters)) for (const o of m.objects ?? []) { if (o.media) o.media = remap.get(o.media) ?? null; if (o.fill?.media) o.fill.media = remap.get(o.fill.media) ?? null; }
    pres.masters = presMeta.masters;
  }
  if (presMeta?.designs) {
    const remap = new Map(Object.entries(presMeta.masterMedia ?? {}).map(([id, path]) => [id, has(path) ? loadMedia(path) : null]));
    for (const d of presMeta.designs) for (const m of Object.values(d.masters ?? {})) for (const o of m.objects ?? []) { if (o.media) o.media = remap.get(o.media) ?? null; if (o.fill?.media) o.fill.media = remap.get(o.fill.media) ?? null; }
    pres.designs = presMeta.designs;
  } else if ([...foreignDesigns.values()].some(Boolean)) {
    // 다른 프로그램 파일: 첫 마스터와 테마가 다른 마스터 → 추가 디자인
    pres.designs = [...foreignDesigns.values()].filter(Boolean);
  }
  if (presMeta?.customLayouts) {
    const remap = new Map(Object.entries(presMeta.masterMedia ?? {}).map(([id, path]) => [id, has(path) ? loadMedia(path) : null]));
    for (const c of presMeta.customLayouts) for (const o of c.bgObjects ?? []) { if (o.media) o.media = remap.get(o.media) ?? null; if (o.fill?.media) o.fill.media = remap.get(o.fill.media) ?? null; }
    pres.customLayouts = presMeta.customLayouts;
  } else if (foreignLayouts.length) {
    // 다른 프로그램 파일의 사용자 지정 레이아웃 (type="cust")
    pres.customLayouts = foreignLayouts.map(({ key, name, bgObjects }) => ({ key, name, bgObjects }));
    pres.masters = { ...(pres.masters ?? {}) };
    for (const f of foreignLayouts) pres.masters[f.key] = { id: `master-${f.key}`, layout: f.key, master: f.key, name: f.name, bg: null, objects: f.objects, anims: [] };
  }
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
  // 사용자 지정 쇼 (슬라이드 관계 id → 슬라이드 id)
  const shows = kids(child(presXml, 'custShowLst'), 'custShow').map((c) => ({ name: attr(c, 'name') ?? '쇼', slides: kids(child(c, 'sldLst'), 'sld').map((x) => slideByRid.get(rid(x))).filter(Boolean) })).filter((c) => c.slides.length);
  if (shows.length) pres.customShows = shows;
  const gl = descendants(child(presXml, 'extLst'), 'sldGuideLst')[0];
  if (gl) pres.guides = kids(gl, 'guide').map((g) => ({ dir: attr(g, 'orient') === 'horz' ? 'h' : 'v', pos: num(attr(g, 'pos'), 0) / 6 }));
  if (!slides.length) pres.slides.push(newSlide(pres, 'title'));
  // 원래 XML 의 서명 (읽은 그대로인지 저장할 때 비교)
  for (const s of slides) {
    for (let i = 0; i < s.objects.length; i++) {
      const o = s.objects[i];
      if (o.chartRaw) o.chartRaw.sig = stableJson(o.chart);
      if (!o.raw) continue;
      if (o.raw.group) {
        const members = s.objects.slice(i, i + o.raw.n);
        if (members.length !== o.raw.n || members.some((m) => m.grp !== o.grp)) { delete o.raw; continue; }
        o.raw.bbox = bboxOf(members);
        o.raw.sig = rawSig(members, o.raw.bbox);
      } else o.raw.sig = rawSig([o], null);
      for (const r of Object.values(o.raw.rels)) if (r.kind === 'slide') { const k = slideOrder.indexOf(r.path); r.slide = slides[k]?.id ?? null; delete r.path; if (!r.slide) { delete o.raw; break; } }
    }
    if (s.rawTiming) s.rawTiming.sig = stableJson(s.anims);
    if (s.rawTransition) s.rawTransition.sig = stableJson(s.transition);
  }
  return { pres, warnings: [...warnings] };
}
