// 슬라이서 스타일: 엑셀 기본 제공 14개 (밝게 1~6, 기타 1~2, 어둡게 1~6) + 사용자 지정 색 (DOM 없음)
import { mergeObjectStyle } from './object-styles.js';
import { objectStyleFillCss } from './object-style-fill.js';
import { ACCENTS, tint, shade } from './tables.js';
import { MODERN_PALETTES, presetColor } from './stylepresets.js';
import { SHEET_SLICER_STYLES } from './stylesheetdata.js';

// 옛 색 이름 → 스타일
const LEGACY = { blue: 'SlicerStyleLight1', orange: 'SlicerStyleLight2', gray: 'SlicerStyleLight3', gold: 'SlicerStyleLight4', sky: 'SlicerStyleLight5', green: 'SlicerStyleLight6' };

// 스타일시트 v1.0 슬라이서 스타일 (슬라이서_001…): 서식 문자열 → 색
const SHEET_SLICERS = new Map(SHEET_SLICER_STYLES);
function dxfColors(str) {
  const o = {};
  for (const part of str ? str.split(';') : []) {
    if (part.startsWith('c=')) o.color = presetColor(part.slice(2));
    else if (part.startsWith('f=')) o.fill = presetColor(part.slice(2));
    else if (/^[LRTB]=/.test(part)) { const col = part.split('/')[1]; o[`side${part[0]}`] = presetColor(col); o.edge ??= o[`side${part[0]}`]; }
  }
  return o;
}
// 엑셀 gradientFill (degree 0 = 왼쪽→오른쪽, 90 = 위→아래) → CSS
const gradCss = (g) => (g?.stops?.length > 1 ? `linear-gradient(${(g.deg + 90) % 360}deg, ${g.stops.map((c) => presetColor(c)).join(', ')})` : undefined);
function buildSheet(el) {
  const wt = dxfColors(el.wt);
  const hr = dxfColors(el.hr);
  const sel = dxfColors(el.selectedItemWithData);
  const item = dxfColors(el.unselectedItemWithData);
  const nd = dxfColors(el.unselectedItemWithNoData);
  const snd = dxfColors(el.selectedItemWithNoData);
  const hov = dxfColors(el.hoveredUnselectedItemWithData);
  const hovSel = dxfColors(el.hoveredSelectedItemWithData);
  const out = {
    frame: wt.fill ?? '#ffffff', border: wt.edge ?? '#d9d9d9', head: hr.color ?? wt.color ?? '#000000', headLine: hr.sideB,
    selFill: sel.fill ?? '#bdd7ee', selText: sel.color ?? '#000000', selBorder: sel.edge ?? '#a6a6a6',
    item: item.fill ?? '#ffffff', itemText: item.color ?? '#000000', itemBorder: item.edge ?? '#d9d9d9',
    noData: nd.color ?? '#a6a6a6', noDataFill: nd.fill, selNoDataFill: snd.fill, selNoData: snd.color,
    hover: gradCss(el.hoveredUnselectedItemWithDataG) ?? hov.fill, hoverText: hov.color,
    hoverSel: gradCss(el.hoveredSelectedItemWithDataG) ?? hovSel.fill, hoverSelText: hovSel.color,
    selGrad: gradCss(el.selectedItemWithDataG), itemGrad: gradCss(el.unselectedItemWithDataG),
  };
  return Object.fromEntries(Object.entries(out).filter(([, v]) => v));
}

function build(name) {
  if (SHEET_SLICERS.has(name)) return buildSheet(SHEET_SLICERS.get(name));
  // WIXEL 모던: 채움(선택 항목이 진한 색) · 소프트(옅은 색) · 테두리(선만)
  const t = /^(?:Wixel|WIXEL)Slicer(Solid|Soft|Line)(\d)$/.exec(name ?? '');
  if (t) {
    const [, dk, ac, soft, line] = MODERN_PALETTES[Math.min(MODERN_PALETTES.length, Math.max(1, Number(t[2]))) - 1];
    const common = { frame: '#ffffff', border: line, head: dk, noData: '#b8bec8' };
    if (t[1] === 'Solid') return { ...common, selFill: ac, selText: '#ffffff', selBorder: ac, item: '#ffffff', itemText: '#334155', itemBorder: line };
    if (t[1] === 'Soft') return { ...common, selFill: soft, selText: dk, selBorder: ac, item: '#ffffff', itemText: '#475569', itemBorder: line };
    return { ...common, selFill: '#ffffff', selText: dk, selBorder: dk, item: '#ffffff', itemText: '#94a3b8', itemBorder: line };
  }
  const m = /^SlicerStyle(Light|Other|Dark)(\d)$/i.exec(name ?? '');
  const kind = m ? m[1].toLowerCase() : 'light';
  const n = m ? Number(m[2]) : 1;
  const a = kind === 'other' ? ACCENTS[0] : ACCENTS[Math.min(6, Math.max(1, n))];
  const base = a.hex;
  if (kind === 'light') {
    return { frame: '#ffffff', border: '#bfbfbf', head: '#000000', selFill: tint(base, 0.6), selText: '#000000', selBorder: tint(base, 0.2), item: '#ffffff', itemText: '#000000', itemBorder: '#d9d9d9', noData: '#a6a6a6' };
  }
  if (kind === 'other') {
    return n === 1
      ? { frame: '#ffffff', border: '#bfbfbf', head: '#000000', selFill: '#d9d9d9', selText: '#000000', selBorder: '#808080', item: '#ffffff', itemText: '#000000', itemBorder: '#d9d9d9', noData: '#a6a6a6' }
      : { frame: '#f2f2f2', border: '#808080', head: '#000000', selFill: '#595959', selText: '#ffffff', selBorder: '#404040', item: '#ffffff', itemText: '#000000', itemBorder: '#bfbfbf', noData: '#a6a6a6' };
  }
  return { frame: '#ffffff', border: shade(base, 0.25), head: '#000000', selFill: base, selText: '#ffffff', selBorder: shade(base, 0.25), item: tint(base, 0.8), itemText: '#000000', itemBorder: tint(base, 0.6), noData: '#8c8c8c' };
}

export const SLICER_STYLES = [
  ...SHEET_SLICER_STYLES.map(([name]) => ({ name, group: '스타일시트', label: name })),
  ...[['Solid', '모던 채움'], ['Soft', '모던 소프트'], ['Line', '모던 선']].flatMap(([k, g]) => MODERN_PALETTES.map((p, i) => ({ name: `WixelSlicer${k}${i + 1}`, group: `WIXEL ${g}`, label: `${g} · ${p[0]}` }))),
  ...[1, 2, 3, 4, 5, 6].map((n) => ({ name: `SlicerStyleLight${n}`, group: '밝게', label: `슬라이서 스타일 밝게 ${n}` })),
  ...[1, 2].map((n) => ({ name: `SlicerStyleOther${n}`, group: '기타', label: `슬라이서 스타일 기타 ${n}` })),
  ...[1, 2, 3, 4, 5, 6].map((n) => ({ name: `SlicerStyleDark${n}`, group: '어둡게', label: `슬라이서 스타일 어둡게 ${n}` })),
].map((s) => ({ ...s, colors: build(s.name) }));

export const SLICER_STYLE_GROUPS = [...new Set(SLICER_STYLES.map((s) => s.group))];
/** 파일에 사용자 지정 슬라이서 스타일로 써야 하는 스타일 (WIXEL 모던 · 스타일시트) */
export const isModernSlicer = (name) => /^(Wixel|WIXEL)Slicer/.test(name ?? '') || SHEET_SLICERS.has(name);
export const slicerStyleColors = build;

/** 슬라이서의 스타일 이름 (옛 color 속성 포함) */
export const slicerStyleName = (sl) => sl.style ?? LEGACY[sl.color] ?? 'SlicerStyleLight1';

/** 실제 색: 스타일 + 사용자 지정(sl.custom) */
export function slicerColors(sl) {
  if (sl.style === 'None') return { ...PLAIN_SLICER };
  if (Array.isArray(sl.styleElements)) return { ...styleElementColors(sl.styleElements), ...(sl.custom ?? {}) };
  return { ...build(slicerStyleName(sl)), ...(sl.custom ?? {}) };
}

/** CSS 변수 문자열 */
export function slicerCssVars(sl) {
  const c = slicerColors(sl);
  return `--sl-frame:${c.frame};--sl-border:${c.border};--sl-head:${c.head};--sl-sel:${c.selFill};--sl-seltext:${c.selText};--sl-selborder:${c.selBorder};`
    + `--sl-item:${c.item};--sl-itemtext:${c.itemText};--sl-itemborder:${c.itemBorder};--sl-nodata:${c.noData};--sl-h:${sl.buttonHeight ?? 24}px`
    + [['hline', c.headLine], ['selg', c.selGrad], ['itemg', c.itemGrad], ['hov', c.hover], ['hovtext', c.hoverText], ['hovsel', c.hoverSel], ['hovseltext', c.hoverSelText],
      ['ndfill', c.noDataFill], ['selndfill', c.selNoDataFill], ['selnd', c.selNoData]].map(([k, v]) => (v ? `;--sl-${k}:${v}` : '')).join('')
    + slicerElementVars(sl)
    + `${sl.fontSize ? `;--sl-fs:${sl.fontSize}pt` : ''}${sl.headSize ? `;--sl-hfs:${sl.headSize}pt` : ''}${sl.bold ? ';--sl-fw:700' : ''}${sl.font ? `;--sl-ff:"${String(sl.font).replace(/"/g, '')}"` : ''}`;
}

export const CUSTOM_KEYS = [
  ['frame', '슬라이서 배경'], ['border', '테두리'], ['head', '머리글 글자'], ['selFill', '선택한 항목 채우기'], ['selText', '선택한 항목 글자'],
  ['selBorder', '선택한 항목 테두리'], ['item', '선택하지 않은 항목 채우기'], ['itemText', '선택하지 않은 항목 글자'], ['itemBorder', '항목 테두리'], ['noData', '데이터 없는 항목 글자'],
];

const PLAIN_SLICER={frame:'#ffffff',border:'transparent',head:'#000000',selFill:'#ffffff',selText:'#000000',selBorder:'transparent',item:'#ffffff',itemText:'#000000',itemBorder:'transparent',noData:'#808080'};
const STATE_CODES={wholeTable:'whole',headerRow:'header',selectedItemWithData:'selected',unselectedItemWithData:'item',selectedItemWithNoData:'selected-empty',unselectedItemWithNoData:'empty',hoveredSelectedItemWithData:'hover-selected',hoveredUnselectedItemWithData:'hover-item',hoveredSelectedItemWithNoData:'hover-selected-empty',hoveredUnselectedItemWithNoData:'hover-empty'};
const borderColor=s=>s?.bbc??s?.btc??s?.blc??s?.brc;
const gradientCss=g=>g?.stops?.length>1&&g.stops.every(s=>Array.isArray(s)&&Number.isFinite(s[0])&&/^#[0-9a-f]{6}$/i.test(s[1]))?`linear-gradient(${((Number(g.deg)||0)+90)%360}deg, ${g.stops.map(([p,c])=>`${c} ${p*100}%`).join(', ')})`:undefined;
function cssGradient(value) {
  const m=/^linear-gradient\((-?[\d.]+)deg,\s*(#[0-9a-f]{6}(?:,\s*#[0-9a-f]{6})+)\)$/i.exec(value??'');
  if(!m)return {};
  const colors=m[2].split(/,\s*/);return {gradient:{deg:(Number(m[1])-90+360)%360,stops:colors.map((c,i)=>[i/(colors.length-1),c])}};
}
function elementState(map,type) {
  const whole=map.wholeTable??{};if(type==='wholeTable')return {...whole};
  const parent=type.startsWith('hovered')?type.replace(/^hovered([A-Z])/,(_,c)=>c.toLowerCase()):type.includes('NoData')?type.replace('NoData','Data'):null;
  return mergeObjectStyle(parent?elementState(map,parent):whole,map[type]);
}
function styleElementColors(elements) {
  const m=Object.fromEntries(elements.map(e=>[e.type,e.style??{}])),get=t=>elementState(m,t),w=get('wholeTable'),h=get('headerRow'),item=get('unselectedItemWithData'),sel=get('selectedItemWithData'),nd=get('unselectedItemWithNoData'),snd=get('selectedItemWithNoData'),hov=get('hoveredUnselectedItemWithData'),hs=get('hoveredSelectedItemWithData');
  const out={...PLAIN_SLICER};
  const put=(key,v)=>{if(v!==undefined)out[key]=v};
  for(const [key,v] of Object.entries({frame:w.fill,border:borderColor(w),head:h.color,headFill:h.fill,headLine:h.bb===false?'transparent':h.bbc,selFill:sel.fill,selText:sel.color,selBorder:borderColor(sel),item:item.fill,itemText:item.color,itemBorder:borderColor(item),noData:nd.color,noDataFill:nd.fill,selNoDataFill:snd.fill,selNoData:snd.color,hover:gradientCss(hov.gradient)??hov.fill,hoverText:hov.color,hoverSel:gradientCss(hs.gradient)??hs.fill,hoverSelText:hs.color,selGrad:gradientCss(sel.gradient),itemGrad:gradientCss(item.gradient)}))put(key,v);
  return out;
}
/** UI 복제 및 표준 XLSX에 쓸 스타일 원소. 기존 객체의 색 재정의도 여기서 확정한다. */
export function slicerStyleElements(sl) {
  if(sl.style==='None')return [];
  if(Array.isArray(sl.styleElements)&&!sl.custom)return structuredClone(sl.styleElements);
  const c=slicerColors(sl),bd=col=>col&&col!=='transparent'?{bt:true,bb:true,bl:true,br:true,btc:col,bbc:col,blc:col,brc:col}:{};
  const rows=[['wholeTable',{fill:c.frame,...bd(c.border)}],['headerRow',{color:c.head,...(c.headFill?{fill:c.headFill}:{}),...(c.headLine?{bb:true,bbc:c.headLine}:{}),bold:true}],['selectedItemWithData',{fill:c.selFill,color:c.selText,...bd(c.selBorder)}],['unselectedItemWithData',{fill:c.item,color:c.itemText,...bd(c.itemBorder)}],['selectedItemWithNoData',{fill:c.selNoDataFill??c.selFill,color:c.selNoData??c.noData,...bd(c.selBorder)}],['unselectedItemWithNoData',{fill:c.noDataFill??c.item,color:c.noData,...bd(c.itemBorder)}],['hoveredSelectedItemWithData',{fill:/^#/.test(c.hoverSel??'')?c.hoverSel:c.selFill,color:c.hoverSelText??c.selText,...bd(c.selBorder)}],['hoveredUnselectedItemWithData',{fill:/^#/.test(c.hover??'')?c.hover:c.item,color:c.hoverText??c.itemText,...bd(c.itemBorder)}]];
  const gradients={selectedItemWithData:c.selGrad,unselectedItemWithData:c.itemGrad,hoveredSelectedItemWithData:c.hoverSel,hoveredUnselectedItemWithData:c.hover};
  return rows.map(([type,style])=>({type,style:{...style,...cssGradient(gradients[type])}}));
}
function slicerElementVars(sl) {
  const elements=sl.style==='None'?[]:sl.styleElements;if(!Array.isArray(elements))return '';
  const map=Object.fromEntries(elements.map(e=>[e.type,e.style??{}])),whole=map.wholeTable??{};let css='';
  const safeColor=v=>/^#[0-9a-f]{6}$/i.test(v??'')?v:undefined;
  for(const [type,code] of Object.entries(STATE_CODES)) {
    const st=elementState(map,type),vars={};
    if(st.font)vars.font='"'+String(st.font).replace(/[";{}<>\\]/g,'')+'"';
    if(Number.isFinite(st.size)&&st.size>0)vars.size=st.size+'pt';
    if(st.bold!==undefined)vars.weight=st.bold?700:400;if(st.italic!==undefined)vars.italic=st.italic?'italic':'normal';
    if(st.underline!==undefined||st.strike!==undefined)vars.decoration=[st.underline?'underline':'',st.strike?'line-through':''].filter(Boolean).join(' ')||'none';
    const fill=objectStyleFillCss(st);if(fill!==undefined)vars.fill=fill;if(safeColor(st.color))vars.color=st.color;
    for(const [side,key]of Object.entries({top:'bt',right:'br',bottom:'bb',left:'bl'}))if(st[key]!==undefined){const kind=st[key+'s']??'thin',width=kind==='double'||kind==='thick'?3:kind==='medium'?2:1;vars[side]=st[key]?`${width}px ${kind==='double'?'double':/dash/i.test(kind)?'dashed':/dot/i.test(kind)?'dotted':'solid'} ${safeColor(st[key+'c'])??'#000000'}`:'0 solid transparent';}
    for(const [key,value]of Object.entries(vars))css+=`;--sl-${code}-${key}:${value}`;
  }
  return css;
}
