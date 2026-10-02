import { el, openDialog } from './ui.js';
import { normPage, paperOf, printScale, headerParts } from './page.js';
import { splitPrintIndexes, imagePagesPdf } from './print-document.js';

const PRINT_CSS = `
.wixel-print-page{position:relative;box-sizing:border-box;background:#fff;color:#000;font-family:"Malgun Gothic",sans-serif;font-size:10pt;overflow:hidden;break-after:page;print-color-adjust:exact;-webkit-print-color-adjust:exact}
.wixel-print-page:last-child{break-after:auto}
.wixel-print-page *{box-sizing:border-box}
.wixel-print-page table{border-collapse:collapse;table-layout:fixed;font-family:inherit;font-size:10pt;margin:0;zoom:1!important}
.wixel-print-page td{padding:1px 3px;white-space:nowrap;overflow:hidden;vertical-align:bottom}
.wixel-print-page table.grid-lines td{border:1px solid #d9d9d9}
.wixel-print-page td.ph{background:#f2f2f2;color:#555;text-align:center;border:1px solid #d9d9d9}
.wixel-print-hf{display:flex;align-items:center;justify-content:space-between;min-height:24px;font-size:9pt;color:#444;gap:8px}
.wixel-print-hf span{flex:1;min-width:0;overflow:hidden;white-space:nowrap}
.wixel-print-hf span:nth-child(2){text-align:center}.wixel-print-hf span:nth-child(3){text-align:right}
.wixel-print-body{position:relative;overflow:hidden}
.wixel-print-body .chart-print{position:absolute;margin:0!important}
.wixel-print-body svg{display:block;max-width:none}
.wixel-print-page .sh-text{position:absolute;inset:0;display:flex;flex-direction:column;box-sizing:border-box;line-height:1.2;overflow:hidden;white-space:pre-wrap;word-break:normal;overflow-wrap:anywhere}
.wixel-print-page .sh-text-content{flex-shrink:0;min-width:0;overflow-wrap:anywhere}.wixel-print-page .sh-text.rich .sh-text-content>div{min-height:1.2em}
@media print{#printArea:has(.wixel-print-page){margin:0!important;padding:0!important}#printArea .wixel-print-page table{page-break-after:auto!important}#printArea .wixel-print-page .chart-print{page-break-inside:auto!important}}
`;

function ensurePrintStyle() {
  if (document.getElementById('wixelPrintStyle')) return;
  const style = el('style', { id: 'wixelPrintStyle' }); style.textContent = PRINT_CSS; document.head.append(style);
}

function hfNode(value) {
  return el('div', { class: 'wixel-print-hf' }, ['left', 'center', 'right'].map(k => el('span', {}, value[k])));
}

// DOM child index is not a worksheet column once a merged cell is present.
function tableSection(section) {
  const rows = [...(section?.children ?? [])], grid = [], cells = [];
  for (let r = 0; r < rows.length; r++) {
    grid[r] ||= []; let c = 0;
    for (const node of rows[r].children) {
      while (grid[r][c]) c++;
      const rs = Math.min(rows.length - r, Math.max(1, node.rowSpan)), cs = Math.max(1, node.colSpan);
      const cell = { node, r, c, rs, cs }; cells.push(cell);
      for (let y = r; y < r + rs; y++) { grid[y] ||= []; for (let x = c; x < c + cs; x++) grid[y][x] = cell; }
      c += cs;
    }
  }
  return { rows, grid, cells };
}

function copyTable(table, model, columns, rows) {
  const result = table.cloneNode(false); result.style.zoom = ''; result.style.margin = '0';
  const originalCols = [...(table.querySelector('colgroup')?.children ?? [])], colgroup = el('colgroup');
  for (const c of columns) if (originalCols[c]) colgroup.append(originalCols[c].cloneNode(true));
  result.append(colgroup);
  const section = (source, selected, tag) => {
    const target = el(tag), seen = new Set();
    for (const r of selected) {
      const row = source.rows[r].cloneNode(false);
      for (const c of columns) {
        const cell = source.grid[r]?.[c]; if (!cell || seen.has(cell)) continue;
        seen.add(cell); const node = cell.node.cloneNode(true);
        node.rowSpan = selected.filter(v => v >= cell.r && v < cell.r + cell.rs).length;
        node.colSpan = columns.filter(v => v >= cell.c && v < cell.c + cell.cs).length;
        row.append(node);
      }
      target.append(row);
    }
    return target;
  };
  if (model.head.rows.length) result.append(section(model.head, model.head.rows.map((_, i) => i), 'thead'));
  result.append(section(model.body, rows, 'tbody')); return result;
}

// Consecutive source coordinates become a page tile. Repeated titles intentionally
// create another tile so objects placed in a repeated title also repeat.
function printSegments(items) {
  const segments = []; let target = 0;
  for (const { start, size } of items) {
    if (Number.isFinite(start) && size > 0) {
      const last = segments[segments.length - 1];
      if (last && Math.abs(last.end - start) < 0.1 && Math.abs(last.target + last.end - last.start - target) < 0.1) last.end = start + size;
      else segments.push({ start, end: start + size, target });
    }
    target += size;
  }
  return segments;
}

function appendPrintObjects(layer, objects, xs, ys) {
  for (const object of objects) for (const x of xs) for (const y of ys) {
    const { node, ox, oy, w, h, bounds } = object;
    const x1 = Math.max(x.start, bounds.x1), x2 = Math.min(x.end, bounds.x2), y1 = Math.max(y.start, bounds.y1), y2 = Math.min(y.end, bounds.y2);
    if (x2 <= x1 || y2 <= y1) continue;
    const clip = el('div', { class: 'wixel-print-object-clip', style: { position: 'absolute', overflow: 'hidden', left: `${x.target}px`, top: `${y.target}px`, width: `${x.end - x.start}px`, height: `${y.end - y.start}px` } });
    const copy = node.cloneNode(true);
    Object.assign(copy.style, { position: 'absolute', left: `${ox - x.start}px`, top: `${oy - y.start}px`, width: `${w}px`, height: `${h}px` });
    clip.append(copy); layer.append(clip);
  }
}

/** 이미 생성·정화한 인쇄 DOM을 값 스냅숏으로 고정한다. 원본 통합 문서는 변경하지 않는다. */
function preparePages(source, page, name, sheet) {
  ensurePrintStyle();
  const pg = normPage(page), paper = paperOf(pg.paper), landscape = pg.orientation === 'landscape';
  const width = (landscape ? paper.h : paper.w) * 96, height = (landscape ? paper.w : paper.h) * 96;
  const margins = pg.margins, contentW = width - (margins.left + margins.right) * 96;
  const contentH = height - (margins.top + margins.bottom) * 96 - 48;
  if (contentW < 20 || contentH < 20) throw new Error('용지에 비해 여백이 너무 큽니다. 페이지 설정에서 여백을 줄이세요.');
  const snapshot = source.cloneNode(true); snapshot.removeAttribute('id'); snapshot.className = 'wixel-print-page';
  snapshot.style.cssText = 'display:block;position:fixed;left:-100000px;top:0;visibility:hidden;width:max-content;overflow:visible;';
  const table = snapshot.querySelector('table');
  if (table) { table.style.zoom = ''; table.style.margin = '0'; table.style.width = `${[...(table.querySelector('colgroup')?.children ?? [])].reduce((sum, col) => sum + (parseFloat(col.style.width) || 80), 0)}px`; }
  document.body.append(snapshot);
  const descriptors = [], objects = []; let scale = 1, oversized = false, model;
  try {
    if (table) {
      model = { head: tableSection(table.querySelector('thead')), body: tableSection(table.querySelector('tbody')) };
      const columns = [...table.querySelector('colgroup').children], sizes = columns.map(col => parseFloat(col.style.width) || 80);
      const body = [...(table.querySelector('tbody')?.children ?? [])], rows = body.map(row => Math.max(1, row.getBoundingClientRect().height));
      const headH = table.querySelector('thead')?.getBoundingClientRect().height || 0;
      const tableW = sizes.reduce((a, b) => a + b, 0), tableH = rows.reduce((a, b) => a + b, headH);
      // 머리글/바닥글이 차지하는 공간도 맞춤 계산에 포함한다.
      scale = printScale({ ...pg, margins: { ...margins, top: margins.top + 24 / 96, bottom: margins.bottom + 24 / 96 } }, tableW, tableH);
      const repeatCols = columns.map((col, i) => ({ c: Number(col.dataset.printCol), i })).filter(x => pg.headings && x.i === 0 || pg.titleCols && Number.isFinite(x.c) && x.c >= pg.titleCols[0] && x.c <= pg.titleCols[1]).map(x => x.i);
      const colPages = splitPrintIndexes(sizes, contentW / scale, repeatCols);
      const rowPages = splitPrintIndexes(rows, contentH / scale - headH);
      oversized = rows.some(h => (h + headH) * scale > contentH + 1) || sizes.some(w => w * scale > contentW + 1);
      for (const cols of colPages) for (const rs of rowPages) descriptors.push({ cols, rows: rs,
        xs: printSegments(cols.map(c => ({ start: columns[c].hasAttribute('data-print-x') ? Number(columns[c].dataset.printX) : NaN, size: sizes[c] }))),
        ys: printSegments([...model.head.rows, ...rs.map(r => body[r])].map(row => ({ start: row.hasAttribute('data-print-y') ? Number(row.dataset.printY) : NaN, size: row.getBoundingClientRect().height }))),
        tableW: cols.reduce((sum, c) => sum + sizes[c], 0), tableH: rs.reduce((sum, r) => sum + rows[r], headH) });
    }
    for (const node of snapshot.querySelectorAll(':scope > .chart-print')) {
      const rect = node.getBoundingClientRect();
      const image = node.querySelector('img'), svg = node.querySelector('svg');
      const w = parseFloat(node.style.width) || parseFloat(image?.style.width) || parseFloat(svg?.getAttribute('width')) || rect.width || contentW;
      const h = parseFloat(node.style.height) || parseFloat(image?.style.height) || parseFloat(svg?.getAttribute('height')) || rect.height || contentH;
      const ox = Number(node.dataset.printX), oy = Number(node.dataset.printY), angle = Number(/rotate\(([-.\d]+)deg\)/.exec(node.style.transform)?.[1] || 0) * Math.PI / 180;
      const rw = Math.abs(w * Math.cos(angle)) + Math.abs(h * Math.sin(angle)), rh = Math.abs(w * Math.sin(angle)) + Math.abs(h * Math.cos(angle));
      if (table && node.hasAttribute('data-print-x') && node.hasAttribute('data-print-y')) objects.push({ node, ox, oy, w, h, bounds: { x1: ox + (w - rw) / 2, x2: ox + (w + rw) / 2, y1: oy + (h - rh) / 2, y2: oy + (h + rh) / 2 } });
      else descriptors.push({ object: node, w, h });
    }
  } finally { snapshot.remove(); }
  if (!descriptors.length) descriptors.push({ empty: true });
  const build = index => {
    const descriptor = descriptors[index];
    const pageNode = el('section', { class: 'wixel-print-page', 'aria-label': `${index + 1}쪽`, style: { width: `${width}px`, height: `${height}px`, padding: `${margins.top * 96}px ${margins.right * 96}px ${margins.bottom * 96}px ${margins.left * 96}px` } });
    const context = { page: index + 1, pages: descriptors.length, file: name, sheet };
    pageNode.append(hfNode(pg.header ? headerParts(pg.header, context) : { left: '', center: `${name} — ${sheet}`, right: '' }));
    const content = el('div', { class: 'wixel-print-body', style: { width: `${contentW}px`, height: `${contentH}px` } });
    if (descriptor.cols) {
      const part = copyTable(table, model, descriptor.cols, descriptor.rows);
      Object.assign(part.style, { width: `${descriptor.tableW}px`, position: 'absolute', left: `${pg.hCenter ? Math.max(0, (contentW - descriptor.tableW * scale) / 2) : 0}px`, top: `${pg.vCenter ? Math.max(0, (contentH - descriptor.tableH * scale) / 2) : 0}px`, transform: `scale(${scale})`, transformOrigin: 'top left' });
      content.append(part);
      const drawings = el('div', { class: 'wixel-print-drawings', style: { position: 'absolute', left: part.style.left, top: part.style.top, width: `${descriptor.tableW}px`, height: `${descriptor.tableH}px`, transform: part.style.transform, transformOrigin: 'top left' } });
      appendPrintObjects(drawings, objects, descriptor.xs, descriptor.ys); content.append(drawings);
    } else if (descriptor.object) {
      const object = descriptor.object.cloneNode(true), fit = Math.min(1, contentW / descriptor.w, contentH / descriptor.h);
      Object.assign(object.style, { left: '0', top: '0', width: `${descriptor.w}px`, height: `${descriptor.h}px`, transform: `scale(${fit}) ${object.style.transform || ''}`, transformOrigin: 'top left' }); content.append(object);
    }
    pageNode.append(content, hfNode(pg.footer ? headerParts(pg.footer, context) : { left: '', center: '', right: '' }));
    return pageNode;
  };
  return { build, width, height, count: descriptors.length, scale, oversized };
}

async function inlineImages(node, signal) {
  for (const image of node.querySelectorAll('img,svg image')) {
    const attr = image.localName === 'img' ? 'src' : 'href', src = image.getAttribute(attr) || image.getAttribute('xlink:href');
    if (!src || src.startsWith('data:')) continue;
    const controller = new AbortController(), cancel = () => controller.abort(); signal?.addEventListener('abort', cancel, { once: true });
    const timer = setTimeout(cancel, 10000);
    try {
      if (signal?.aborted) throw new Error('PDF 저장을 취소했습니다.');
      const result = await fetch(src, { credentials: 'omit', signal: controller.signal });
      if (!result.ok) throw new Error();
      const blob = await result.blob(); if (!blob.type.startsWith('image/') || blob.size > 20 * 1024 * 1024) throw new Error();
      const data = await new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = reject; reader.readAsDataURL(blob); });
      image.setAttribute(attr, data); image.removeAttribute('crossorigin'); image.removeAttribute('srcset');
    } catch { throw new Error('외부 그림을 PDF에 담지 못했습니다. 그림을 파일로 삽입하거나 브라우저 인쇄를 이용하세요.'); }
    finally { clearTimeout(timer); signal?.removeEventListener('abort', cancel); }
  }
}

async function rasterPage(node, width, height, signal) {
  await inlineImages(node, signal); if (signal.aborted) throw new Error('PDF 저장을 취소했습니다.');
  const content = new XMLSerializer().serializeToString(node);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml"><style>${PRINT_CSS}</style>${content}</div></foreignObject></svg>`;
  const image = new Image(); image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`; await image.decode();
  const canvas = document.createElement('canvas'); canvas.width = Math.ceil(width * 2); canvas.height = Math.ceil(height * 2);
  const ctx = canvas.getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height); ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
  const data = canvas.toDataURL('image/jpeg', 0.94).split(',')[1], raw = atob(data), jpeg = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) jpeg[i] = raw.charCodeAt(i);
  return { jpeg, width: canvas.width, height: canvas.height, paperWidth: width * 0.75, paperHeight: height * 0.75 };
}

export function openPrintPreview({ source, page, name, sheet, saveFile = null, pdfPreferred = false }) {
  const prepared = preparePages(source, page, name, sheet), controller = new AbortController();
  let index = 0, busy = false, closed = false;
  const status = el('div', { role: 'status', style: { minHeight: '20px', fontSize: '12px' } });
  const counter = el('span', { style: { minWidth: '100px', textAlign: 'center' } });
  const stage = el('div', { class: 'print-preview-stage', style: { overflow: 'auto', maxHeight: 'min(65vh,750px)', background: '#dfe3e8', padding: '12px', display: 'grid', justifyItems: 'center' } });
  const show = next => {
    index = Math.max(0, Math.min(prepared.count - 1, next));
    const node = prepared.build(index), zoom = Math.min(0.72, Math.max(0.15, (Math.min(innerWidth - 90, 900)) / prepared.width));
    node.style.zoom = String(zoom); stage.replaceChildren(node); counter.textContent = `${index + 1} / ${prepared.count}쪽`;
    prev.disabled = index === 0; nextButton.disabled = index === prepared.count - 1;
  };
  const prev = el('button', { type: 'button', class: 'btn', 'aria-label': '이전 인쇄 페이지', onclick: () => show(index - 1) }, '이전');
  const nextButton = el('button', { type: 'button', class: 'btn', 'aria-label': '다음 인쇄 페이지', onclick: () => show(index + 1) }, '다음');
  const nativePrint = () => {
    if (busy) return;
    const printArea = document.getElementById('printArea');
    printArea.replaceChildren(); for (let i = 0; i < prepared.count; i++) printArea.append(prepared.build(i));
    const pageStyle = document.getElementById('pageStyle') || el('style', { id: 'pageStyle' });
    if (!pageStyle.isConnected) document.head.append(pageStyle);
    pageStyle.textContent = `@page{size:${prepared.width / 96}in ${prepared.height / 96}in;margin:0;}`;
    status.textContent = '인쇄 창이 열리지 않으면 PDF 파일 저장을 이용하세요. 브라우저 인쇄 창에서도 PDF로 저장할 수 있습니다.';
    try { window.print(); } catch { status.textContent = '이 환경에서는 인쇄 창을 열 수 없습니다. PDF 파일 저장을 이용하세요.'; }
  };
  const savePdf = async () => {
    if (busy) return; busy = true; pdf.disabled = true;
    try {
      if (prepared.count > 150) throw new Error('직접 PDF 저장은 150쪽까지 지원합니다. 인쇄 영역을 줄이거나 브라우저 인쇄를 이용하세요.');
      const make = async () => {
        await document.fonts?.ready;
        const pages = [];
        for (let i = 0; i < prepared.count; i++) {
          if (controller.signal.aborted) return null;
          status.textContent = `PDF 만드는 중… ${i + 1} / ${prepared.count}쪽`;
          pages.push(await rasterPage(prepared.build(i), prepared.width, prepared.height, controller.signal));
          await new Promise(resolve => setTimeout(resolve, 0));
        }
        return closed ? null : new Blob([imagePagesPdf(pages, { title: `${name} — ${sheet}` })], { type: 'application/pdf' });
      };
      const fileName = `${String(name || '통합 문서').replace(/[\\/:*?"<>|]/g, '_')}.pdf`;
      if (saveFile) {
        const done = await saveFile(fileName, make);
        if (!closed) status.textContent = done ? `${prepared.count}쪽 PDF 파일을 저장했습니다.` : 'PDF 저장을 취소했습니다.';
      } else {
        const blob = await make(); if (!blob) return;
        const url = URL.createObjectURL(blob), a = el('a', { href: url, download: fileName });
        document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 60000);
        status.textContent = `${prepared.count}쪽 PDF 파일의 다운로드를 시작했습니다.`;
      }
    } catch (error) { if (!closed) status.textContent = error.message || 'PDF를 저장하지 못했습니다. 브라우저 인쇄를 이용하세요.'; }
    finally { busy = false; pdf.disabled = false; }
  };
  const print = el('button', { type: 'button', class: pdfPreferred ? 'btn' : 'btn primary', onclick: nativePrint }, '인쇄');
  const pdf = el('button', { type: 'button', class: pdfPreferred ? 'btn primary' : 'btn', onclick: savePdf }, 'PDF 파일 저장');
  const body = el('div', { class: 'print-preview', style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    el('div', { style: { display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' } }, print, pdf, prev, counter, nextButton),
    el('div', { class: 'muted', style: { fontSize: '12px' } }, `현재 시트 · ${prepared.count}쪽 · 배율 ${Math.round(prepared.scale * 100)}%. 표와 개체를 시트의 위치에 맞춰 함께 출력합니다. 직접 저장하는 PDF는 화면 모양을 보존하는 이미지 PDF이며 본문 텍스트 검색은 지원하지 않습니다.`),
    ...(prepared.oversized ? [el('div', { role: 'alert', style: { color: '#a65b00' } }, '한 행 또는 열이 인쇄 가능 영역보다 큽니다. 페이지 설정에서 배율을 줄여 잘림을 방지하세요.')] : []), status, stage);
  const dialog = openDialog({ title: pdfPreferred ? 'PDF 다운로드' : '인쇄 미리보기', initialFocus: () => pdfPreferred ? pdf : print, width: Math.min(1000, innerWidth - 24), body, buttons: [{ label: '닫기' }], onClose: () => { closed = true; controller.abort(); } });
  dialog.root.style.minWidth = '0'; dialog.root.style.maxWidth = 'calc(100vw - 24px)';
  show(0); return dialog;
}

/** Static, script-free web page with the same page geometry as print/PDF. */
export function htmlPrintDocument({ source, page, name, sheet }) {
  const prepared=preparePages(source,page,name,sheet);
  if(prepared.count>300)throw new Error('웹페이지 저장은 300쪽 이내로 인쇄 영역을 나누어 주세요.');
  const html=document.implementation.createHTMLDocument(`${name} — ${sheet}`);
  html.documentElement.lang='ko';
  const charset=html.createElement('meta');charset.setAttribute('charset','utf-8');html.head.prepend(charset);
  const viewport=html.createElement('meta');viewport.name='viewport';viewport.content='width=device-width, initial-scale=1';html.head.append(viewport);
  const policy=html.createElement('meta');policy.httpEquiv='Content-Security-Policy';policy.content="default-src 'none'; img-src data: https: http:; style-src 'unsafe-inline'; font-src data:; base-uri 'none'; form-action 'none'";html.head.append(policy);
  const style=html.createElement('style');style.textContent=PRINT_CSS+`body{margin:0;padding:24px;background:#e9edf1;overflow:auto}.wixel-print-page{margin:0 auto 24px;box-shadow:0 2px 16px #0001}@media print{body{padding:0;background:#fff}.wixel-print-page{margin:0;box-shadow:none}@page{size:${prepared.width/96}in ${prepared.height/96}in;margin:0}}`;html.head.append(style);
  for(let i=0;i<prepared.count;i++)html.body.append(html.importNode(prepared.build(i),true));
  return '<!doctype html>\n'+html.documentElement.outerHTML;
}
