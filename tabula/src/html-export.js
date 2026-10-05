// 현재 시트를 인쇄 페이지/DOM 없이 자기 완결 HTML로 저장합니다.
import { Axis } from './axis.js';
import { formatValue } from './format.js';
import { tableCellDisplayStyle } from './table-format.js';
import { prepareCond, condFormatAt } from './condfmt.js';
import { safeUrl } from './safe-html.js';
import { slicerColors } from './slicerstyle.js';
import { normalizePrintAreas } from './print-layout.js';
import { fontFamilyCandidates } from './fonts.js';
import { createFontUsage, addFontUsage, collectMarkupFontUsage, embedFontCss } from './font-export.js';

const esc = value => String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const finite = (value, fallback = 0) => Number.isFinite(Number(value)) ? Number(value) : fallback;
const color = value => /^(#[0-9a-f]{3,8}|[a-z]+|rgba?\([\d.,%\s]+\))$/i.test(String(value ?? '')) ? String(value) : null;
const json = value => JSON.stringify(value).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
const nextFrame = () => new Promise(resolve => setTimeout(resolve, 0));

/** 서식만 있는 거대한 꼬리/인쇄 영역은 기본 HTML 범위에 포함하지 않습니다. 실제 값과 병합 내용은 자르지 않습니다. */
export function htmlSheetRange(wb, si) {
  const sheet = wb.sheets[si];
  if (!sheet) throw new Error('저장할 시트가 없습니다.');
  const used = wb.usedRange(si);
  let rows = Math.max(1, used.rows), cols = Math.max(1, used.cols);
  for (const m of sheet.merges ?? []) if (m.r1 < used.rows && m.c1 < used.cols) { rows = Math.max(rows, m.r2 + 1); cols = Math.max(cols, m.c2 + 1); }
  for (const sp of wb.spillsOf?.(si) ?? []) { rows = Math.max(rows, sp.r + sp.h); cols = Math.max(cols, sp.c + sp.w); }
  return { r1: 0, c1: 0, r2: rows - 1, c2: cols - 1 };
}

function axes(wb, si, options) {
  const s = wb.sheets[si];
  return {
    rows: options.rows ?? new Axis(s.defRowH ?? 20, s.rowHeights, [s.hiddenRows, s.filter?.hidden, ...(s.tables ?? []).map(t => t.filter?.hidden)]),
    cols: options.cols ?? new Axis(s.defColW ?? 64, s.colWidths, [s.hiddenCols], 16384),
  };
}
function* visible(axis, first, last) {
  for (let i = first; i <= last;) {
    const n = axis.nextVisible(i, 1);
    if (n < i || n > last || axis.isHidden(n)) return;
    yield n; i = n + 1;
  }
}
function cssFor(st, formatted) {
  const align = st.align && st.align !== 'general' ? st.align === 'centerContinuous' ? 'center' : st.align : formatted.align;
  const css = [`text-align:${['left','center','right','justify'].includes(align) ? align : 'left'}`, `font-size:${Math.max(1, Math.min(512, finite(st.size, 11)))}pt`];
  if (st.font && !/[;{}<>\\\r\n]/.test(st.font)) css.push(`font-family:${fontFamilyCandidates(st.font).map(f => `'${String(f).replace(/[\\'"\r\n\f<>]/g, '')}'`).join(',')},sans-serif`);
  if (st.bold) css.push('font-weight:700');
  if (st.italic) css.push('font-style:italic');
  if (color(formatted.color ?? st.color)) css.push(`color:${color(formatted.color ?? st.color)}`);
  if (color(st.fill)) css.push(`background:${color(st.fill)}`);
  if (st.wrap) css.push('white-space:pre-wrap');
  if (['top','middle','bottom'].includes(st.valign)) css.push(`vertical-align:${st.valign}`);
  if (st.underline || st.strike) css.push(`text-decoration:${[st.underline && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ')}`);
  for (const [key, side] of [['bt','top'],['bb','bottom'],['bl','left'],['br','right']]) if (st[key]) {
    const kind = st[`${key}s`], width = kind === 'thick' || kind === 'double' ? 3 : /^(medium|slant)/.test(kind ?? '') ? 2 : 1;
    css.push(`border-${side}:${width}px ${kind === 'double' ? 'double' : /dash/i.test(kind ?? '') ? 'dashed' : kind === 'dotted' || kind === 'hair' ? 'dotted' : 'solid'} ${color(st[`${key}c`]) ?? '#000'}`);
  }
  return css.join(';');
}
function cellHtml(wb, si, r, c, merge, cond, options) {
  const rr = merge?.r ?? r, cc = merge?.c ?? c, value = wb.getValue(si, rr, cc);
  const st = { ...tableCellDisplayStyle(wb, si, rr, cc), ...condFormatAt(cond, wb, si, rr, cc, value).style };
  const formatted = formatValue(value, st, wb.date1904), cell = wb.getCell(si, rr, cc);
  st.font ||= wb.defaultFont?.name || '맑은 고딕';
  addFontUsage(options.fontUsage, st.font, formatted.text, st);
  let content = esc(formatted.text);
  const image = formatted.image ?? cell?.image, src = image && safeUrl(image.src, 'image');
  if (src) content = `<img src="${esc(src)}" alt="${esc(image.alt ?? '')}" style="max-width:100%;max-height:100%;object-fit:contain">`;
  const href = cell?.link && safeUrl(cell.link, 'link');
  if (href) content = `<a href="${esc(href)}" rel="noopener noreferrer">${content}</a>`;
  return `<td data-row="${r}" data-col="${c}"${merge ? ` rowspan="${merge.rows}" colspan="${merge.cols}"` : ''} style="${esc(cssFor(st, formatted))}">${content}</td>`;
}
// 셀마다 거대한 Map을 만들지 않고 현재 표시 구간의 병합만 계산합니다.
function mergeBands(merges, rows, cols) {
  const bands = new Map();
  for (const m of merges) {
    const rs = rows.filter(r => r >= m.r1 && r <= m.r2), cs = cols.filter(c => c >= m.c1 && c <= m.c2);
    if (!rs.length || !cs.length) continue;
    const anchor = { r:m.r1, c:m.c1, rows:rs.length, cols:cs.length, first:cs[0], last:cs[cs.length-1] };
    for (const r of rs) { if (!bands.has(r)) bands.set(r, []); bands.get(r).push({ ...anchor, top:r === rs[0] }); }
  }
  for (const list of bands.values()) list.sort((a,b)=>a.first-b.first);
  return bands;
}
const CSS = `*{box-sizing:border-box}body{margin:16px;color:#17221c;background:#fff;font:11pt "맑은 고딕",sans-serif}h1{font-size:16pt;margin:0 0 12px}table{border-collapse:collapse;table-layout:fixed;background:white}td{padding:2px 3px;overflow:hidden;white-space:pre;vertical-align:bottom}table.grid td{border:1px solid #ddd}td a{color:inherit}.wx-sheet{overflow:auto;margin-bottom:20px}.wx-stage{position:relative;min-width:max-content}.wx-object{position:absolute;overflow:hidden}.wx-object svg{max-width:100%;max-height:100%}.wx-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;position:sticky;top:0;padding:10px 0;background:white;z-index:2}.wx-toolbar input{width:90px}.wx-toolbar button,.wx-toolbar input{font:inherit;padding:4px 8px}.wx-note{font-size:10pt;color:#536359}.wx-canvas-view{max-width:100%;max-height:75vh;overflow:auto;border:1px solid #ddd}.wx-canvas{position:relative}noscript{display:block;color:#8a2424}@media print{.wx-toolbar,.wx-note{display:none}body{margin:0}.wx-sheet{overflow:visible}}`;
// 고정 프로그램만 실행합니다. 데이터는 HTML 종료 문자를 이스케이프한 비실행 JSON입니다.
const PAGER = `(()=>{const data=[...document.querySelectorAll('script[data-wx-page]')],host=document.getElementById('wx-body'),status=document.getElementById('wx-status'),input=document.getElementById('wx-jump');let current=0;function show(n){current=Math.max(0,Math.min(data.length-1,n));const p=JSON.parse(data[current].textContent);host.innerHTML=p.html.join('');status.textContent=(current+1)+' / '+data.length+' · '+p.r1+'–'+p.r2+'행 · '+p.c1+'–'+p.c2+'열';input.value=p.r1;document.getElementById('wx-prev').disabled=current===0;document.getElementById('wx-next').disabled=current===data.length-1;}document.getElementById('wx-prev').onclick=()=>show(current-1);document.getElementById('wx-next').onclick=()=>show(current+1);document.getElementById('wx-go').onclick=()=>{const r=Number(input.value);let i=data.findIndex(e=>r>=Number(e.dataset.r1)&&r<=Number(e.dataset.r2));if(i>=0)show(i);};if(data.length)show(0);})();`;

/** 슬라이서의 현재 크기·선택 상태를 읽기 전용으로 보존합니다. 가상화된 빈 DOM을 복제하지 않습니다. */
export function exportSlicerHtml(sl, model, { start = 0 } = {}) {
  const colors = slicerColors(sl), items = model?.items ?? [], columns = Math.max(1, Math.min(100, Math.floor(finite(sl.columns, 1))));
  const buttonHeight = Math.max(1, finite(sl.buttonHeight, 24)), gap = Math.max(0, finite(sl.gap, 3)), head = sl.showHeader === false ? 0 : 26;
  const viewport = Math.max(buttonHeight, finite(sl.h, 200)-head-26);
  const count = Math.max(columns, Math.min(2000, Math.ceil(viewport/(buttonHeight+gap))*columns));
  const first = Math.max(0, Math.min(items.length, Math.floor(finite(start)))), last = Math.min(items.length,first+count);
  let selected=0;for(const item of items)if(item.selected)selected++;
  const title=sl.caption ?? model?.caption ?? '', state=`전체 ${items.length.toLocaleString('ko-KR')}개 · 선택 ${selected.toLocaleString('ko-KR')}개`;
  const font=cssFor({size:sl.fontSize??11,font:sl.font,bold:sl.bold},{align:'left'});
  const parts=[`<div role="img" aria-label="${esc(title)} · 읽기 전용 필터 스냅샷 · ${esc(state)}" style="width:100%;height:100%;overflow:hidden;border:1px solid ${color(colors.border)??'#bfbfbf'};background:${color(colors.frame)??'#fff'};${esc(font)}">`];
  if(head)parts.push(`<div style="height:26px;padding:3px 5px;font-weight:700;color:${color(colors.head)??'#000'};overflow:hidden;white-space:nowrap">${esc(title)}</div>`);
  if(model?.broken)parts.push(`<div style="padding:5px">${esc(model.broken)}</div>`);
  else {
    parts.push(`<div style="display:grid;grid-template-columns:repeat(${columns},minmax(0,1fr));gap:${gap}px;padding:3px 5px;max-height:${viewport}px;overflow:hidden">`);
    for(let i=first;i<last;i++){
      const it=items[i],on=!!it.selected,has=it.hasData!==false;
      const fill=on?(has?colors.selFill:colors.selNoDataFill??colors.selFill):(has?colors.item:colors.noDataFill??colors.item);
      const fg=on?(has?colors.selText:colors.selNoData??colors.noData):(has?colors.itemText:colors.noData);
      const line=on?colors.selBorder:colors.itemBorder;
      parts.push(`<span title="${esc(it.text)}" data-slicer-selected="${on?'1':'0'}" style="height:${buttonHeight}px;line-height:${Math.max(1,buttonHeight-2)}px;padding:0 4px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;border:1px solid ${color(line)??'#ddd'};background:${color(fill)??'#fff'};color:${color(fg)??'#000'}">${esc(it.text)}</span>`);
    }
    parts.push('</div>');
  }
  const interval=first>0||last<items.length?` · ${first+1}–${last}번째 항목 표시`:'';
  parts.push(`<div style="font-size:9pt;padding:2px 5px;white-space:nowrap;overflow:hidden" title="${esc(state+interval)}">${esc(state+interval)}</div></div>`);
  return parts.join('');
}

/** renderObject(kind, object)는 개체 하나씩 정화한 HTML을 반환해야 합니다. 전체 시트 DOM은 만들지 않습니다. */
export async function createSheetHtmlBlob(wb, si, options = {}) {
  return writeSheetHtml(wb, si, options);
}

/** 디스크 저장: 완성 HTML 전체를 보관하지 않고 한 구간씩 기록합니다. */
export async function writeSheetHtmlToSink(wb, si, sink, options = {}) {
  if (!sink || typeof sink.write !== 'function') throw new TypeError('웹페이지 저장 스트림이 필요합니다.');
  return writeSheetHtml(wb, si, options, sink);
}

async function writeSheetHtml(wb, si, options, sink = null) {
  const fontUsage = createFontUsage(); options = { ...options, fontUsage };
  const sheet = wb.sheets[si];
  if (!sheet) throw new Error('저장할 시트가 없습니다.');
  const { rows:rowAxis, cols:colAxis } = axes(wb, si, options), fallback = options.range ?? htmlSheetRange(wb, si);
  const areas = options.respectPrintArea ? normalizePrintAreas(sheet.page, fallback) : [fallback];
  const estimate = areas.reduce((sum,a)=>sum+(a.r2-a.r1+1)*(a.c2-a.c1+1),0);
  const paged = estimate > (options.staticCellLimit ?? 50000), pageRows = Math.max(1, Math.min(250, options.pageRows ?? 250));
  const title = `${options.name ?? '통합 문서'} — ${sheet.name}`, parts = [], pending = []; let size = 0, work = 0, tick = Date.now(), pageCount = 0;
  const check = () => { options.assertCurrent?.(); if (options.signal?.aborted) throw new Error('웹페이지 저장을 취소했습니다.'); };
  const flush = () => { if (pending.length) { parts.push(new Blob(pending)); pending.length = 0; size = 0; } };
  const push = text => { pending.push(text); size += text.length; if (size >= 262144) flush(); };
  let bytesWritten = 0;
  const drain = async () => {
    if (!sink) return;
    for (const part of parts) for (let at=0;at<part.size;at+=1<<20) {
      check(); const bytes = new Uint8Array(await part.slice(at,at+(1<<20)).arrayBuffer());
      await sink.write(bytes); bytesWritten += bytes.length;
    }
    parts.length = 0;
  };
  const yieldWork = async () => { check(); await drain(); if (Date.now()-tick>=12) { options.onProgress?.(Math.min(.98,work/Math.max(1,estimate)), '웹페이지 만드는 중'); await nextFrame(); check(); tick=Date.now(); } };
  check();
  push(`<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data: https: http:; style-src 'unsafe-inline'; font-src data:; ${paged ? "script-src 'nonce-wixel-html-export'; " : ''}base-uri 'none'; form-action 'none'"><title>${esc(title)}</title><style>${CSS}</style></head><body><h1>${esc(title)}</h1>`);
  if (paged) push('<p class="wx-note">현재 시트의 모든 데이터가 이 파일에 포함되어 있습니다. 표시할 행·열 구간을 이동하세요.</p><nav class="wx-toolbar" aria-label="표 구간 이동"><button id="wx-prev" type="button">이전</button><button id="wx-next" type="button">다음</button><span id="wx-status" role="status"></span><label>행 <input id="wx-jump" type="number" min="1" value="1"></label><button id="wx-go" type="button">이동</button></nav><div id="wx-body" class="wx-sheet"></div><noscript>이 큰 표의 구간을 보려면 JavaScript를 켜세요. 모든 데이터는 파일 안에 저장되어 있습니다.</noscript>');
  const objects=[...(sheet.charts??[]).map(o=>['chart',o]),...(sheet.images??[]).map(o=>['image',o]),...(sheet.shapes??[]).map(o=>['shape',o]),...(sheet.slicers??[]).map(o=>['slicer',o])].filter(([,o])=>!o.hidden).sort((a,b)=>(a[1].z??0)-(b[1].z??0));
  let canvasW=0,canvasH=0;
  for(const a of areas){canvasW=Math.max(canvasW,colAxis.pos(a.c2+1));canvasH=Math.max(canvasH,rowAxis.pos(a.r2+1));}
  let objectW=0,objectH=0;for(const [,o]of objects){objectW=Math.max(objectW,finite(o.x)+Math.max(1,finite(o.w,100)));objectH=Math.max(objectH,finite(o.y)+Math.max(1,finite(o.h,100)));}
  if(!paged)push(`<div class="wx-sheet"><div class="wx-canvas" style="width:${Math.max(canvasW,objectW)}px;height:${Math.max(canvasH,objectH)}px">`);
  const cond = prepareCond(wb, si), allMerges=sheet.merges ?? [];
  for (const area of areas) {
    const allCols = [...visible(colAxis, area.c1, area.c2)];
    const colStep = paged ? 64 : Math.max(1,allCols.length);
    if (!allCols.length) continue;
    for (let at=0;at<allCols.length;at+=colStep) {
      const cols=allCols.slice(at,at+colStep), width=cols.reduce((sum,c)=>sum+colAxis.size(c),0);
      const colgroup=`<colgroup>${cols.map(c=>`<col style="width:${colAxis.size(c)}px">`).join('')}</colgroup>`;
      const grid = !sheet.noGrid || sheet.page?.gridlines;
      let batch=[];
      const emit = async rs => {
        const bands=mergeBands(allMerges,rs,cols);let firstPart=true;
        if(paged)push(`<script type="application/json" data-wx-page data-r1="${rs[0]+1}" data-r2="${rs[rs.length-1]+1}">{"r1":${rs[0]+1},"r2":${rs[rs.length-1]+1},"c1":${cols[0]+1},"c2":${cols[cols.length-1]+1},"html":[`);
        const add = paged ? text=>{if(!firstPart)push(',');firstPart=false;push(json(text));} : push;
        add(`<section class="wx-stage"${paged?'':` style="position:absolute;left:${colAxis.pos(cols[0])}px;top:${rowAxis.pos(rs[0])}px"`}><table class="${grid?'grid':''}" aria-label="${esc(sheet.name)}" style="width:${width}px">${colgroup}<tbody>`);
        for(const r of rs){
          add(`<tr style="height:${rowAxis.size(r)}px">`);const merges=bands.get(r)??[];let mi=0;
          for(const c of cols){while(mi<merges.length&&merges[mi].last<c)mi++;const m=merges[mi];if(m&&c>=m.first&&c<=m.last){if(m.top&&c===m.first)add(cellHtml(wb,si,r,c,m,cond,options));}else add(cellHtml(wb,si,r,c,null,cond,options));work++;}
          add('</tr>'); if (sink && parts.length) await drain();
        }
        add('</tbody></table></section>');
        if(paged)push(']}</script>');
        pageCount++;
      };
      // 작은 표는 하나의 병합 계산으로 유지하고, 큰 표만 표시 구간별로 나눕니다.
      for(const r of visible(rowAxis,area.r1,area.r2)){batch.push(r);if(paged&&batch.length>=pageRows){await emit(batch);batch=[];await yieldWork();}}
      if(batch.length){await emit(batch);await yieldWork();}
    }
  }
  if(objects.length){
    if(!options.renderObject)throw new Error('차트·그림·도형 저장 준비가 필요합니다.');
    // HTML은 인쇄물이 아니므로 noPrint 개체도 표시합니다. 숨겨진 개체만 제외합니다.
    if(paged)push(`<section aria-label="차트 · 그림 · 도형 · 슬라이서"><p class="wx-note">개체는 표 구간과 별도로 시트의 원래 좌표에 배치했습니다.</p><div class="wx-canvas-view"><div class="wx-canvas" style="width:${objectW}px;height:${objectH}px">`);
    for(const [kind,o]of objects){check();const html=await options.renderObject(kind,o);if(typeof html!=='string')throw new Error('개체를 웹페이지로 변환하지 못했습니다.');collectMarkupFontUsage(html,fontUsage);push(`<div class="wx-object" data-object="${esc(o.id)}" style="left:${finite(o.x)}px;top:${finite(o.y)}px;width:${Math.max(1,finite(o.w,100))}px;height:${Math.max(1,finite(o.h,100))}px;${o.rot?`transform:rotate(${finite(o.rot)}deg);`:''}">${html}</div>`);await yieldWork();}
    if(paged)push('</div></div></section>');
  }
  if(!paged)push('</div></div>');
  if(!pageCount)push('<p>표시할 셀이 없습니다.</p>');
  if(paged&&pageCount)push(`<script nonce="wixel-html-export">${PAGER}</script>`);
  options.onProgress?.(.99,'출력 글꼴 준비 중');
  const fonts = await embedFontCss(fontUsage, { signal: options.signal, onWarning: options.onWarning, fetchImpl: options.fetchFont });
  check();
  if (fonts.css) push(`<style data-wixel-embedded-fonts="true">${fonts.css}</style>`);
  for (const warning of fonts.warnings) push(`<p class="wx-note" role="alert">${esc(warning)}</p>`);
  push('</body></html>');check();flush();await drain();check();options.onProgress?.(1,'웹페이지 준비 완료');return sink ? { bytesWritten } : new Blob(parts,{type:'text/html;charset=utf-8'});
}
