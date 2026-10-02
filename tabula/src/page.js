// 페이지 설정 (DOM 없음): 용지 방향 · 크기 · 여백 · 배율 · 인쇄 영역 · 인쇄 제목 · 머리글/바닥글
// 시트 속성 page = { orientation: 'portrait'|'landscape', paper: 9, margins: { left, right, top, bottom, header, footer } (인치),
//   scale: 100, fitW: 0, fitH: 0 (0 = 맞추지 않음), hCenter, vCenter, gridlines, headings, header, footer,
//   area: { r1, c1, r2, c2 } | null, areas?: [range,...] (별도 페이지, area는 첫 영역),
//   rowBreaks/colBreaks: [다음 페이지 첫 zero-based 인덱스], order: 'downThenOver'|'overThenDown',
//   titleRows: [r1, r2] | null, titleCols: [c1, c2] | null }
export const PAPERS = [
  { id: 9, label: 'A4 (210 × 297mm)', w: 8.27, h: 11.69, css: 'A4' },
  { id: 8, label: 'A3 (297 × 420mm)', w: 11.69, h: 16.54, css: 'A3' },
  { id: 11, label: 'A5 (148 × 210mm)', w: 5.83, h: 8.27, css: 'A5' },
  { id: 13, label: 'B5 (182 × 257mm)', w: 7.17, h: 10.12, css: 'B5' },
  { id: 1, label: 'Letter (8.5 × 11in)', w: 8.5, h: 11, css: 'letter' },
  { id: 5, label: 'Legal (8.5 × 14in)', w: 8.5, h: 14, css: 'legal' },
];

export const MARGINS = [
  { id: 'normal', label: '보통', m: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } },
  { id: 'wide', label: '넓게', m: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 } },
  { id: 'narrow', label: '좁게', m: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } },
];

export function normPage(p) {
  const result = {
    orientation: 'portrait', paper: 9, margins: { ...MARGINS[0].m }, scale: 100, fitW: 0, fitH: 0, hCenter: false, vCenter: false,
    gridlines: false, headings: false, header: '', footer: '', area: null, titleRows: null, titleCols: null, order: 'downThenOver', ...(p ?? {}),
  };
  result.margins = { ...MARGINS[0].m, ...(p?.margins ?? {}) };
  for (const key of Object.keys(MARGINS[0].m)) if (!Number.isFinite(Number(result.margins[key])) || Number(result.margins[key]) < 0) result.margins[key] = MARGINS[0].m[key]; else result.margins[key] = Number(result.margins[key]);
  result.scale = Math.max(10, Math.min(400, Number(result.scale) || 100));
  for (const key of ['fitW', 'fitH']) result[key] = Math.max(0, Math.min(32767, Math.floor(Number(result[key]) || 0)));
  result.orientation = result.orientation === 'landscape' ? 'landscape' : 'portrait';
  result.order = result.order === 'overThenDown' ? 'overThenDown' : 'downThenOver';
  return result;
}

/** 리본의 자동 맞춤 입력. 수동 배율을 입력하면 너비/높이 맞춤을 해제한다. */
export function pageScalePatch(key, value) {
  const n = Number(value);
  if (!Number.isFinite(n)) throw new Error('인쇄 배율과 쪽 수는 숫자로 입력하세요.');
  if (key === 'scale') return { scale: Math.max(10, Math.min(400, Math.round(n))), fitW: 0, fitH: 0 };
  if (key === 'fitW' || key === 'fitH') return { [key]: Math.max(0, Math.min(32767, Math.floor(n))) };
  throw new Error('지원하지 않는 인쇄 배율 설정입니다.');
}

export const paperOf = (id) => PAPERS.find((p) => p.id === Number(id)) ?? PAPERS[0];

/** 인쇄 가능한 너비 · 높이 (px, 96dpi) */
export function printableSize(page) {
  const p = normPage(page);
  const paper = paperOf(p.paper);
  const [w, h] = p.orientation === 'landscape' ? [paper.h, paper.w] : [paper.w, paper.h];
  return { w: (w - p.margins.left - p.margins.right) * 96, h: (h - p.margins.top - p.margins.bottom) * 96 };
}

/** 배율: 자동 맞춤이면 표 너비 · 높이로 계산 (1 = 100%) */
export function printScale(page, tableW, tableH) {
  const p = normPage(page);
  if (!p.fitW && !p.fitH) return Math.max(0.1, Math.min(4, (p.scale || 100) / 100));
  const { w, h } = printableSize(p);
  let s = 1;
  if (p.fitW) s = Math.min(s, (w * p.fitW) / Math.max(1, tableW));
  if (p.fitH) s = Math.min(s, (h * p.fitH) / Math.max(1, tableH));
  return Math.max(0.1, s);
}

/** 머리글/바닥글 코드 (&P 쪽 번호, &N 전체 쪽 수, &D 날짜, &T 시간, &F 파일 이름, &A 시트 이름) → 글자. &L &C &R 로 왼쪽 · 가운데 · 오른쪽 */
export function headerParts(code, { page = 1, pages = 1, file = '', sheet = '', now = new Date() } = {}) {
  const parts = { left: '', center: '', right: '' };
  let cur = 'center';
  const pad = (n) => String(n).padStart(2, '0');
  let i = 0;
  const s = String(code ?? '');
  while (i < s.length) {
    const ch = s[i];
    if (ch === '&' && i + 1 < s.length) {
      const k = s[i + 1].toUpperCase();
      i += 2;
      if (k === 'L') cur = 'left';
      else if (k === 'C') cur = 'center';
      else if (k === 'R') cur = 'right';
      else if (k === 'P') parts[cur] += String(page);
      else if (k === 'N') parts[cur] += String(pages);
      else if (k === 'D') parts[cur] += `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
      else if (k === 'T') parts[cur] += `${pad(now.getHours())}:${pad(now.getMinutes())}`;
      else if (k === 'F') parts[cur] += file;
      else if (k === 'A') parts[cur] += sheet;
      else if (k === '&') parts[cur] += '&';
      else if (k === '"' || k === 'B' || k === 'I' || k === 'U' || /\d/.test(k)) {
        // 글꼴 · 크기 지정은 건너뜀 (&"굴림,굵게" / &12)
        if (k === '"') { const e = s.indexOf('"', i); i = e < 0 ? s.length : e + 1; } else while (i < s.length && /\d/.test(s[i])) i++;
      }
      continue;
    }
    parts[cur] += ch;
    i++;
  }
  return parts;
}

const inch = (v) => Number(Number(v).toFixed(4));
/** xlsx: <printOptions/> <pageMargins/> <pageSetup/> <headerFooter/> (+ sheetPr 의 pageSetUpPr fitToPage) */
export function pageXml(page, esc) {
  if (!page) return { printOptions: '', margins: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', setup: '', headerFooter: '', fitToPage: false };
  const p = normPage(page);
  const m = p.margins;
  const po = [p.gridlines ? 'gridLines="1"' : '', p.headings ? 'headings="1"' : '', p.hCenter ? 'horizontalCentered="1"' : '', p.vCenter ? 'verticalCentered="1"' : ''].filter(Boolean);
  const fit = !!(p.fitW || p.fitH);
  const setup = [`paperSize="${p.paper}"`, fit ? '' : `scale="${Math.round(p.scale || 100)}"`, fit ? `fitToWidth="${p.fitW || 0}" fitToHeight="${p.fitH || 0}"` : '', `orientation="${p.orientation}"`, `pageOrder="${p.order}"`].filter(Boolean);
  const hf = p.header || p.footer ? `<headerFooter>${p.header ? `<oddHeader>${esc(p.header)}</oddHeader>` : ''}${p.footer ? `<oddFooter>${esc(p.footer)}</oddFooter>` : ''}</headerFooter>` : '';
  return {
    printOptions: po.length ? `<printOptions ${po.join(' ')}/>` : '',
    margins: `<pageMargins left="${inch(m.left)}" right="${inch(m.right)}" top="${inch(m.top)}" bottom="${inch(m.bottom)}" header="${inch(m.header)}" footer="${inch(m.footer)}"/>`,
    setup: `<pageSetup ${setup.join(' ')}/>`,
    headerFooter: hf,
    fitToPage: fit,
  };
}

/** xlsx 요소들 → page (기본값과 같으면 null) */
export function pageFromXml({ printOptions, pageMargins, pageSetup, headerFooter, fitToPage }) {
  const p = normPage();
  let any = false;
  const on = (v) => v === '1' || v === 'true';
  if (printOptions) {
    const a = printOptions.attrs;
    p.gridlines = on(a.gridLines); p.headings = on(a.headings); p.hCenter = on(a.horizontalCentered); p.vCenter = on(a.verticalCentered);
    any ||= p.gridlines || p.headings || p.hCenter || p.vCenter;
  }
  if (pageMargins) {
    const a = pageMargins.attrs;
    const m = { left: +a.left, right: +a.right, top: +a.top, bottom: +a.bottom, header: +a.header, footer: +a.footer };
    if (Object.values(m).every(Number.isFinite)) {
      p.margins = m;
      any ||= JSON.stringify(m) !== JSON.stringify(MARGINS[0].m);
    }
  }
  if (pageSetup) {
    const a = pageSetup.attrs;
    if (a.pageOrder === 'overThenDown') { p.order = 'overThenDown'; any = true; }
    if (a.orientation === 'landscape') { p.orientation = 'landscape'; any = true; }
    if (a.paperSize && paperOf(a.paperSize).id === Number(a.paperSize) && Number(a.paperSize) !== 9) { p.paper = Number(a.paperSize); any = true; }
    if (fitToPage) { p.fitW = a.fitToWidth === undefined ? 1 : Number(a.fitToWidth); p.fitH = a.fitToHeight === undefined ? 1 : Number(a.fitToHeight); any = true; }
    else if (a.scale && Number(a.scale) !== 100) { p.scale = Number(a.scale); any = true; }
  }
  if (headerFooter) {
    const h = headerFooter.children.find((c) => c.name === 'oddHeader')?.text ?? '';
    const f = headerFooter.children.find((c) => c.name === 'oddFooter')?.text ?? '';
    if (h || f) { p.header = h; p.footer = f; any = true; }
  }
  return any ? p : null;
}
