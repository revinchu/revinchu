// 선택 범위의 셀만 독립 SVG로 출력한다. 화면 배율/스크롤/DOM 및 Workbook 상태를 바꾸지 않는다.
import { Axis } from './axis.js';
import { CellMap } from './cellmap.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { MAX_ROWS, MAX_COLS } from './formula.js';
import { formatValue, formatGeneral } from './format.js';
import { prepareCond, condFormatAt, ICON_SVG, ruleRanges } from './condfmt.js';
import { tableAt, tableCellStyle } from './tables.js';
import { resolveGridBorders, gridBorderPaintOrder } from './grid-lines.js';
import { fontAlias } from './fonts.js';
import { safeUrl } from './safe-html.js';
import { esc } from './xml.js';

export const RANGE_IMAGE_LIMITS = { cells: 100000, rows: 20000, cols: 2000, dimension: 65536, conditionalCells: 1000000 };
const BORDER = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const DASH = { hair:'1 1', dotted:'1 2', dashed:'3 2', dashDot:'6 2 1 2', dashDotDot:'6 2 1 2 1 2', mediumDashed:'6 4', mediumDashDot:'8 4 2 4', mediumDashDotDot:'8 4 2 4 2 4', slantDashDot:'8 2 2 2' };
const color = (value, fallback = '#000000') => typeof value === 'string' && /^(?:#[\da-f]{3,8}|[a-z]+|rgba?\([\d\s.,%/+-]+\))$/i.test(value.trim()) ? value.trim() : fallback;
const number = (value, fallback = 0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
const n = value => String(Math.round(value * 1000) / 1000);
const family = value => {
  const name = String(value || '맑은 고딕').replace(/['";\\]/g, '');
  return `'${name}'${fontAlias(name) ? `, '${fontAlias(name)}'` : ''}, 'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif`;
};

// 수식 결과·이름·분산 및 조건부서식의 계산 캐시는 출력용 인스턴스에만 기록한다.
// 셀과 대형 열 블록은 읽기 전용으로 공유하므로 전체 통합 문서를 복사하지 않는다.
function exportBook(wb) {
  if (wb.pending?.length) throw new Error('셀 입력을 확정한 다음 그림으로 저장하세요.');
  const book = Object.create(wb);
  for (const [key, value] of Object.entries(wb)) {
    if (value instanceof Map) book[key] = new Map(value);
    else if (value instanceof Set) book[key] = new Set(value);
    else if (Array.isArray(value)) book[key] = value.slice();
  }
  book.sheets = wb.sheets.map(sheet => ({ ...sheet }));
  book.names = wb.names.map(entry => ({ ...entry }));
  book.caches = wb.caches.map(cache => cache ? new CellMap(cache) : cache);
  book.ctxs = []; book.colIdx = []; book.pending = []; book.depth = 0;
  book.evaluating = new Set(); book.nameStack = new Set(); book.rangeMemo = new Map();
  book.pivotMemo = null; book.sheetNameMemo = new Map(); book.circHits = new Set();
  return book;
}

function visible(axis, first, last, limit) {
  const out = [];
  for (let i = axis.nextVisible(first, 1); i <= last; i = axis.nextVisible(i + 1, 1)) {
    if (out.length >= limit) throw new Error('선택 범위가 너무 큽니다. 20,000행·2,000열·100,000셀 이하로 나누어 저장하세요.');
    out.push(i);
  }
  return out;
}

function measureFor(options) {
  if (typeof options.measureText === 'function') return options.measureText;
  const ctx = typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d');
  return (text, style) => {
    if (ctx) {
      ctx.font = `${style.italic ? 'italic ' : ''}${style.bold ? '700 ' : ''}${style.size}pt ${family(style.font)}`;
      return ctx.measureText(String(text)).width;
    }
    let units = 0; for (const ch of String(text)) units += /[\x00-\x7f]/.test(ch) ? .55 : 1;
    return units * style.size * 4 / 3;
  };
}

function wrapLines(text, width, style, measure) {
  const lines = [];
  for (const paragraph of String(text).split(/\r\n?|\n/)) {
    let line = '';
    for (const ch of paragraph) {
      if (line && measure(line + ch, style) > width) { lines.push(line); line = ch; } else line += ch;
    }
    lines.push(line);
  }
  return lines;
}

function fitNumber(value, width, style, measure) {
  if ((!style.numFmt || style.numFmt === 'general') && style.decimals === undefined) {
    if (Math.abs(value) >= 1e-4 && Math.abs(value) < 1e11 && !Number.isInteger(value)) for (let d = 9; d >= 0; d--) {
      const text = formatGeneral(Number(value.toFixed(d)));
      if (text !== '0' && text !== '-0' && measure(text, style) <= width) return text;
    }
    for (let d = 4; d >= 0; d--) {
      const [mantissa, exponent] = value.toExponential(d).split('e');
      const text = `${mantissa.includes('.') ? mantissa.replace(/\.?0+$/, '') : mantissa}E${exponent[0] === '-' ? '-' : '+'}${exponent.replace(/^[+-]/, '').padStart(2, '0')}`;
      if (measure(text, style) <= width) return text;
    }
  }
  return '#'.repeat(Math.max(1, Math.min(10000, Math.floor(width / Math.max(1, measure('#', style))))));
}

function patternBits(kind, x, y) {
  switch (kind) {
    case 'gray0625': return y % 4 === 0 && x === (y ? 4 : 0);
    case 'gray125': return y % 2 === 0 && x % 4 === (y % 4 ? 2 : 0);
    case 'lightGray': return x % 4 === (y % 2 ? 2 : 0);
    case 'mediumGray': return (x + y) % 2 === 0;
    case 'darkGray': return x % 4 !== (y % 2 ? 1 : 3);
    case 'lightHorizontal': return y % 4 === 0;
    case 'darkHorizontal': return y % 4 < 2;
    case 'lightVertical': return x % 4 === 0;
    case 'darkVertical': return x % 4 < 2;
    case 'lightDown': return (x - y + 8) % 4 === 0;
    case 'darkDown': return (x - y + 8) % 4 < 2;
    case 'lightUp': return (x + y) % 4 === 3;
    case 'darkUp': return (x + y) % 4 >= 2;
    case 'lightGrid': return x % 4 === 0 || y % 4 === 0;
    case 'darkGrid': return x % 4 < 2 || y % 4 < 2;
    case 'lightTrellis': return (x + y) % 4 === 3 || (x - y + 8) % 4 === 0;
    case 'darkTrellis': return y % 2 === 0 || (x + y) % 4 < 2;
    default: return false;
  }
}

/** 셀 범위 → SVG. 셀 그림은 prepareImageSvg의 명시 resolver로 포함해야 한다. 부유 개체는 제외한다. */
export function rangeImageSvg(wb, si, range, options = {}) {
  if (!wb?.sheets?.[si]) throw new Error('그림으로 저장할 시트를 찾을 수 없습니다.');
  const rg = Object.fromEntries(['r1','c1','r2','c2'].map(key => [key, range?.[key]]));
  if (!Object.values(rg).every(Number.isInteger) || rg.r1 < 0 || rg.c1 < 0 || rg.r2 < rg.r1 || rg.c2 < rg.c1 || rg.r2 >= MAX_ROWS || rg.c2 >= MAX_COLS) throw new Error('그림으로 저장할 셀 범위를 확인하세요.');
  const book = exportBook(wb), sheet = book.sheets[si];
  const rowsAxis = new Axis(sheet.defRowH ?? DEFAULT_ROW_HEIGHT, sheet.rowHeights, [sheet.hiddenRows, sheet.filter?.hidden, ...(sheet.tables ?? []).map(table => table.filter?.hidden)], MAX_ROWS);
  const colsAxis = new Axis(sheet.defColW ?? DEFAULT_COL_WIDTH, sheet.colWidths, [sheet.hiddenCols], MAX_COLS);
  const rows = visible(rowsAxis, rg.r1, rg.r2, RANGE_IMAGE_LIMITS.rows), cols = visible(colsAxis, rg.c1, rg.c2, RANGE_IMAGE_LIMITS.cols);
  if (!rows.length || !cols.length) throw new Error('선택 범위의 모든 행 또는 열이 숨겨져 있습니다.');
  if (rows.length * cols.length > RANGE_IMAGE_LIMITS.cells) throw new Error('선택 범위가 100,000셀을 넘습니다. 범위를 나누어 저장하세요.');
  const x0 = colsAxis.pos(rg.c1), y0 = rowsAxis.pos(rg.r1), w = colsAxis.pos(rg.c2 + 1) - x0, h = rowsAxis.pos(rg.r2 + 1) - y0;
  if (!(w > 0 && h > 0) || w + 4 > RANGE_IMAGE_LIMITS.dimension || h + 4 > RANGE_IMAGE_LIMITS.dimension) throw new Error('그림의 한 변은 여백을 포함해 65,536픽셀 이하여야 합니다. 범위를 나누어 저장하세요.');
  const warnings = new Set(), measure = measureFor(options), defs = [], backgrounds = [], content = [], edges = [], grid = [], memoFill = new Map();
  const gridlines = options.gridlines === undefined ? !sheet.noGrid : !!options.gridlines;
  const defaultStyle = { font: book.defaultFont?.name || '맑은 고딕', size: book.defaultFont?.size || 11 };
  const used = book.usedRange(si); let condCells = 0;
  for (const rule of sheet.cond ?? []) if (['dup','unique','top','bottom','aboveAvg','belowAvg','bar','scale','icons'].includes(rule.type)) for (const area of ruleRanges(rule)) {
    condCells += Math.max(0, Math.min(area.r2, used.rows - 1) - area.r1 + 1) * Math.max(0, Math.min(area.c2, used.cols - 1) - area.c1 + 1);
  }
  if (condCells > RANGE_IMAGE_LIMITS.conditionalCells) throw new Error('조건부 서식 계산 범위가 너무 큽니다. 1,000,000셀 이하로 나누어 저장하세요.');
  const cond = prepareCond(book, si), styleCache = new Map(), values = new Map();
  const valueAt = (r, c) => { const key = `${r},${c}`; if (!values.has(key)) values.set(key, book.getValue(si, r, c)); return values.get(key); };
  const styleAt = (r, c) => {
    const key = `${r},${c}`;
    if (!styleCache.has(key)) {
      const table = tableAt(sheet, r, c), own = Object.fromEntries(Object.entries(book.styleAt(si, r, c)).filter(([,v]) => v !== undefined));
      const cf = condFormatAt(cond, book, si, r, c, valueAt(r, c));
      styleCache.set(key, { style: { ...defaultStyle, ...(table ? tableCellStyle(table, r, c) : null), ...own, ...cf.style }, cf });
    }
    return styleCache.get(key);
  };
  const fill = style => {
    const key = JSON.stringify([style.gradient,style.pattern,style.patternColor,style.fill]);
    if (memoFill.has(key)) return memoFill.get(key);
    let paint = color(style.fill, 'none');
    if (style.gradient?.stops?.length) {
      const id = `rif${memoFill.size}`, g = style.gradient, stops = g.stops.map(([p,c]) => `<stop offset="${n(Math.max(0,Math.min(1,number(p))))}" stop-color="${esc(color(c))}"/>`).join('');
      if (g.path) defs.push(`<radialGradient id="${id}" cx="${n((number(g.l)+number(g.r))/2)}" cy="${n((number(g.t)+number(g.b))/2)}" r="1">${stops}</radialGradient>`);
      else { const a = number(g.deg)*Math.PI/180, dx = Math.cos(a)/2, dy = Math.sin(a)/2; defs.push(`<linearGradient id="${id}" x1="${n(.5-dx)}" y1="${n(.5-dy)}" x2="${n(.5+dx)}" y2="${n(.5+dy)}">${stops}</linearGradient>`); }
      paint = `url(#${id})`;
    } else if (style.pattern) {
      const id = `rif${memoFill.size}`; let path = '';
      for (let y=0;y<8;y++) for (let x=0;x<8;x++) if (patternBits(style.pattern,x,y)) path += `M${x} ${y}h1v1h-1z`;
      defs.push(`<pattern id="${id}" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="${esc(paint)}"/><path d="${path}" fill="${esc(color(style.patternColor))}"/></pattern>`); paint=`url(#${id})`;
    }
    memoFill.set(key,paint);return paint;
  };
  const edge = (style,key,vertical,at,start,end) => {
    if (!style[key]) return;
    const pattern = style[key+'s'] || 'thin'; edges.push({vertical,at,start,end,width:BORDER[pattern]||1,pattern,color:color(style[key+'c'])});
  };
  const mergeAt = new Map();
  for (const merge of sheet.merges ?? []) if (merge.r1<=rg.r2 && merge.r2>=rg.r1 && merge.c1<=rg.c2 && merge.c2>=rg.c1) {
    for (const r of rows) if(r>=merge.r1&&r<=merge.r2) for(const c of cols) if(c>=merge.c1&&c<=merge.c2) mergeAt.set(`${r},${c}`,merge);
  }
  const colIndex = new Map(cols.map((c,i)=>[c,i])), rendered = new Set();
  const emptyAt = (r,c) => !mergeAt.has(`${r},${c}`) && !book.getCell(si,r,c)?.raw && (valueAt(r,c)==null || valueAt(r,c)==='');
  for (const r of rows) for (const c of cols) {
    const merge=mergeAt.get(`${r},${c}`); if(merge && rendered.has(merge))continue;if(merge)rendered.add(merge);
    const rr=merge?.r1??r,cc=merge?.c1??c,r2=merge?.r2??r,c2=merge?.c2??c;
    const x=colsAxis.pos(cc)-x0,y=rowsAxis.pos(rr)-y0,cw=colsAxis.pos(c2+1)-colsAxis.pos(cc),ch=rowsAxis.pos(r2+1)-rowsAxis.pos(rr);
    if(!(cw>0&&ch>0))continue;
    const {style,cf}=styleAt(rr,cc),value=valueAt(rr,cc),cell=book.getCell(si,rr,cc),formatted=formatValue(value,style,book.date1904);
    if(gridlines)grid.push(`M${n(x)} ${n(y)}H${n(x+cw)}V${n(y+ch)}H${n(x)}Z`);
    const bg=fill(style);
    if(bg!=='none')backgrounds.push(`<rect x="${n(x-.5)}" y="${n(y-.5)}" width="${n(cw+1)}" height="${n(ch+1)}" fill="${esc(bg)}"/>`);
    if(merge) {
      for(const col of cols) if(col>=cc&&col<=c2){const a=colsAxis.pos(col)-x0,b=a+colsAxis.size(col);edge(styleAt(rr,col).style,'bt',false,y,a,b);edge(styleAt(r2,col).style,'bb',false,y+ch,a,b);}
      for(const row of rows) if(row>=rr&&row<=r2){const a=rowsAxis.pos(row)-y0,b=a+rowsAxis.size(row);edge(styleAt(row,cc).style,'bl',true,x,a,b);edge(styleAt(row,c2).style,'br',true,x+cw,a,b);}
      // 왼쪽 위 셀에만 저장된 병합 외곽 테두리도 보존한다.
      edge(style,'bt',false,y,x,x+cw);edge(style,'bb',false,y+ch,x,x+cw);edge(style,'bl',true,x,y,y+ch);edge(style,'br',true,x+cw,y,y+ch);
    }else{edge(style,'bt',false,y,x,x+cw);edge(style,'bb',false,y+ch,x,x+cw);edge(style,'bl',true,x,y,y+ch);edge(style,'br',true,x+cw,y,y+ch);}
    const align=style.align==='centerContinuous'?'center':style.align&&style.align!=='general'?style.align:formatted.align;
    let clipX=x,clipW=cw,textW=cw;
    const ci=colIndex.get(c), emptyLeft=ci>0&&emptyAt(r,cols[ci-1]),emptyRight=ci+1<cols.length&&emptyAt(r,cols[ci+1]);
    if(!merge && formatted.text && !style.wrap && !style.shrink && !style.rotate && !cf.icon && typeof value==='string') {
      if(align==='right'||align==='center'&&emptyLeft&&emptyRight) for(let i=ci-1;i>=0&&emptyAt(r,cols[i]);i--){const size=colsAxis.size(cols[i]);clipX-=size;clipW+=size;}
      if(align==='left'||align==='center'&&emptyLeft&&emptyRight) for(let i=ci+1;i<cols.length&&emptyAt(r,cols[i]);i++)clipW+=colsAxis.size(cols[i]);
    }
    if(!merge&&style.align==='centerContinuous'&&formatted.text)for(let i=ci+1;i<cols.length&&emptyAt(r,cols[i])&&styleAt(r,cols[i]).style.align==='centerContinuous';i++)textW+=colsAxis.size(cols[i]);
    if(textW>cw){clipX=x;clipW=textW;}
    const clip=`ric${content.length}`,parts=[];defs.push(`<clipPath id="${clip}"><rect x="${n(clipX)}" y="${n(y)}" width="${n(clipW)}" height="${n(ch)}"/></clipPath>`);
    if(cf.bar){const bw=cw*Math.max(0,Math.min(100,cf.bar.pct))/100;let paint=color(cf.bar.color);if(cf.bar.gradient){const id=`rib${content.length}`;defs.push(`<linearGradient id="${id}"><stop stop-color="${esc(paint)}"/><stop offset="1" stop-color="${esc(paint)}" stop-opacity=".2"/></linearGradient>`);paint=`url(#${id})`;}parts.push(`<rect x="${n(x+(cf.bar.neg?cw-bw:0))}" y="${n(y+ch*.14)}" width="${n(bw)}" height="${n(ch*.72)}" fill="${esc(paint)}"/>`);}
    for(const [key,x1,y1,x2,y2] of [['dd',x,y,x+cw,y+ch],['du',x,y+ch,x+cw,y]])if(style[key])parts.push(`<path d="M${n(x1)} ${n(y1)}L${n(x2)} ${n(y2)}" stroke="${esc(color(style[key+'c']))}" stroke-width="${BORDER[style[key+'s']]||1}"${DASH[style[key+'s']]?` stroke-dasharray="${DASH[style[key+'s']]}"`:''}/>`);
    if(cf.icon&&ICON_SVG[cf.icon]){const icon=ICON_SVG[cf.icon].replace('<svg ',`<svg x="${n(x+2)}" y="${n(y+(ch-14)/2)}" `).replace(/id="([^"]+)"/g,(_,id)=>`id="${clip}${id}"`).replace(/url\(#([^)]+)\)/g,(_,id)=>`url(#${clip}${id})`);parts.push(icon);}
    let text=cf.hideValue||value===0&&sheet.noZeros?'':formatted.text;
    const checkbox=style.checkbox&&(typeof value==='boolean'||value==null||value==='');
    if(checkbox){const bx=x+(cw-14)/2,by=y+(ch-14)/2;parts.push(`<rect x="${n(bx)}" y="${n(by)}" width="14" height="14" rx="2" fill="${value===true?esc(color(style.color,'#217346')):'#fff'}" stroke="#777"/>`);if(value===true)parts.push(`<path d="M${n(bx+3)} ${n(by+7)}l3 3 5-6" fill="none" stroke="#fff" stroke-width="2"/>`);text='';}
    if(formatted.image){const image=formatted.image,src=safeUrl(image.src,'image');if(!src||! /^(?:https?:|blob:|data:image\/)/i.test(src))throw new Error('셀 그림 주소를 안전하게 저장할 수 없습니다.');parts.push(`<image x="${n(x)}" y="${n(y)}" width="${n(cw)}" height="${n(ch)}" preserveAspectRatio="${image.sizing===1?'none':'xMidYMid meet'}" href="${esc(src)}"/>`);text='';}
    if(cell?.phonetic?.visible===true&&cell?.phonetic?.runs?.length)warnings.add('윗주는 셀의 원문만 내보냅니다.');
    if(text){
      const st={...style,size:Math.max(1,Math.min(4096,number(style.size,11)))};
      const iconSpace=cf.icon?18:0,indent=Math.max(0,number(st.indent))*9,padLeft=2+iconSpace+(align==='left'?indent:0),padRight=3+(align==='right'?indent:0),room=Math.max(1,cw-padLeft-padRight);
      const wrap=st.wrap&&typeof value!=='number';
      if(!wrap&&measure(text,st)>room){if(st.shrink)st.size=Math.max(.1,st.size*room/measure(text,st));else if(typeof value==='number')text=fitNumber(value,room,st,measure);}
      const lines=wrap?wrapLines(text,room,st,measure):[String(text).replace(/\r\n?|\n/g,' ')],fontPx=st.size*4/3,lh=fontPx*1.2,total=lines.length*lh;
      const baseline=(st.valign==='top'?y+1:st.valign==='middle'?y+(ch-total)/2:y+ch-total-1)+fontPx*.96;
      const tx=align==='center'?x+(textW+padLeft-padRight)/2:align==='right'?x+cw-padRight:x+padLeft,anchor=align==='center'?'middle':align==='right'?'end':'start';
      const attrs=`font-family="${esc(family(st.font))}" font-size="${n(fontPx)}"${st.bold?' font-weight="700"':''}${st.italic?' font-style="italic"':''} fill="${esc(color(formatted.color||st.color))}"${st.underline||st.strike?` text-decoration="${st.underline?'underline ':''}${st.strike?'line-through':''}"`:''}`;
      if(st.rotate===255){let yy=y+fontPx;const xx=align==='left'?x+padLeft+fontPx/2:align==='right'?x+cw-padRight-fontPx/2:x+cw/2;parts.push(`<text ${attrs} text-anchor="middle">${[...String(text)].map(ch=>{const t=`<tspan x="${n(xx)}" y="${n(yy)}">${esc(ch)}</tspan>`;yy+=fontPx;return t;}).join('')}</text>`);}
      else if(number(st.rotate)){const a=Math.max(-90,Math.min(90,number(st.rotate))),tw=measure(text,st),rad=Math.abs(a)*Math.PI/180,bw=tw*Math.cos(rad)+lh*Math.sin(rad),bh=tw*Math.sin(rad)+lh*Math.cos(rad),cx=align==='left'?x+padLeft+bw/2:align==='right'?x+cw-padRight-bw/2:x+cw/2,cy=st.valign==='top'?y+bh/2:st.valign==='middle'?y+ch/2:y+ch-bh/2;parts.push(`<text ${attrs} x="0" y="${n(fontPx*.34)}" text-anchor="middle" transform="translate(${n(cx)} ${n(cy)}) rotate(${-a})">${esc(text)}</text>`);}
      else parts.push(`<text ${attrs} text-anchor="${anchor}">${lines.map((line,i)=>`<tspan x="${n(tx)}" y="${n(baseline+i*lh)}">${esc(line)}</tspan>`).join('')}</text>`);
    }
    content.push(`<g data-cell="${rr},${cc}" clip-path="url(#${clip})">${parts.join('')}</g>`);
  }
  const borders=gridBorderPaintOrder(resolveGridBorders(edges)).map(e=>{const line=(at,width)=>`<path d="${e.vertical?`M${n(at)} ${n(e.start)}V${n(e.end)}`:`M${n(e.start)} ${n(at)}H${n(e.end)}`}" fill="none" stroke="${esc(e.color)}" stroke-width="${width}"${DASH[e.pattern]?` stroke-dasharray="${DASH[e.pattern]}"`:' shape-rendering="crispEdges"'}/>`;return e.pattern==='double'?line(e.at-1,1)+line(e.at+1,1):line(e.at,e.width);}).join('');
  const padding=2,width=w+padding*2,height=h+padding*2;
  defs.push(`<clipPath id="riRange"><rect x="-2" y="-2" width="${n(width)}" height="${n(height)}"/></clipPath>`);
  const background=options.background==='transparent'?'none':color(options.background,'#ffffff');
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${n(width)}" height="${n(height)}" viewBox="0 0 ${n(width)} ${n(height)}"><title>선택한 셀 범위</title><defs>${defs.join('')}</defs><rect width="${n(width)}" height="${n(height)}" fill="${esc(background)}"/><g transform="translate(2 2)" clip-path="url(#riRange)">${grid.length?`<path d="${grid.join('')}" fill="none" stroke="#d9d9d9" stroke-width="1" shape-rendering="crispEdges"/>`:''}${backgrounds.join('')}${content.join('')}${borders}</g></svg>`;
  return {svg,width,height,range:rg,rows:rows.length,cols:cols.length,warnings:[...warnings]};
}
