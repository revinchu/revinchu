// PDF → 편집할 수 있는 문서 (DOM 없음)
// pdf.js (WIPOINT 와 같은 해석기) 가 쪽마다 글 상자 · 그림을 만들면,
//  흐름(flow, 기본): 줄을 문단으로 다시 잇고 (줄 끝 · 들여쓰기 · 간격으로 판단), 큰 글자는 제목, 그림은 본문 안에 — Word 의 [PDF 다시 배치] 와 같은 방식
//  모양 유지(layout): 쪽마다 글 상자를 원래 자리에 (모양은 같지만 고치기는 불편)
import { readPdf } from './pdf.js';
import { newDoc, newSection, para, uid, normalizeRuns } from './model.js';

const PT = 4 / 3;

function runsOf(p) {
  return p.runs.map((r) => {
    const o = { text: r.t };
    if (r.size) o.sz = r.size;
    if (r.font) { o.font = cleanFont(r.font); o.fontEa = o.font; }
    if (r.b) o.b = true;
    if (r.i) o.i = true;
    if (r.color) o.color = r.color;
    return o;
  });
}
const cleanFont = (f) => f.replace(/^[A-Z]{6}\+/, '').replace(/[-,](Bold|Italic|BoldItalic|Regular|Roman|Medium|Light)$/i, '').replace(/MT$/, '');
const plain = (p) => p.runs.map((r) => r.t).join('');

export function readPdfDoc(bytes, opts = {}) {
  const { pres, warnings, jobs } = readPdf(bytes, opts);
  const doc = newDoc();
  doc.media = pres.media;
  const W = pres.size.w; const H = pres.size.h;
  // 본문 글자 크기 (가장 흔한 크기) → 제목 판단
  const sizes = new Map();
  for (const s of pres.slides) for (const o of s.objects) if (o.text) for (const p of o.text.paras) for (const r of p.runs) sizes.set(r.size, (sizes.get(r.size) ?? 0) + r.t.length);
  const bodySz = [...sizes].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 10;
  // 여백: 모든 쪽 글 상자의 바깥 경계
  let minX = Infinity; let maxX = 0; let minY = Infinity; let maxY = 0;
  for (const s of pres.slides) for (const o of s.objects) if (o.text || o.type === 'image') { minX = Math.min(minX, o.x); maxX = Math.max(maxX, o.x + o.w); minY = Math.min(minY, o.y); maxY = Math.max(maxY, o.y + o.h); }
  const sec = newSection();
  sec.page = { w: W, h: H, orient: W > H ? 'landscape' : 'portrait' };
  const clampM = (v, d) => (Number.isFinite(v) ? Math.max(24, Math.min(v, d)) : d);
  sec.margin = { ...sec.margin, l: clampM(minX - 2, 96), r: clampM(W - maxX - 2, 96), t: clampM(minY - 2, 96), b: clampM(H - maxY - 2, 96), header: 36, footer: 36 };
  doc.section = sec;
  const bodyW = W - sec.margin.l - sec.margin.r;
  const body = [];
  if (opts.mode === 'layout') {
    pres.slides.forEach((s, i) => {
      const anchor = { ...para(''), runs: [], pPr: { spc: { before: 0, after: 0, line: 1, rule: 'auto' }, ...(i ? { pageBreakBefore: true } : {}) } };
      for (const o of s.objects) {
        if (o.type === 'image') anchor.runs.push({ obj: 'img', id: uid('i'), media: o.media, w: o.w, h: o.h, float: { wrap: 'behind', relH: 'page', relV: 'page', x: o.x, y: o.y } });
        else if (o.text) anchor.runs.push({ obj: 'shape', id: uid('s'), w: o.w + 4, h: o.h + 2, prst: 'rect', inset: { l: 0, r: 0, t: 0, b: 0 }, float: { wrap: 'front', relH: 'page', relV: 'page', x: o.x, y: o.y }, blocks: o.text.paras.map((p) => ({ ...para(''), pPr: { align: p.align === 'ctr' ? 'center' : p.align === 'r' ? 'right' : 'left', spc: { before: 0, after: 0, line: (p.lineSpacingPt ?? 12) * PT, rule: 'exact' } }, runs: normalizeRuns(runsOf(p)) })) });
      }
      body.push(anchor);
    });
    doc.body = body.length ? body : [para('')];
    return { doc, warnings: ['PDF 를 원래 모양대로 열었습니다 (쪽마다 글 상자). 글을 이어 고치려면 [흐름으로 열기]를 쓰세요.', ...warnings], jobs };
  }
  pres.slides.forEach((s, si) => {
    const objs = s.objects.filter((o) => o.text || o.type === 'image').sort((a, b) => (Math.abs(a.y - b.y) < 4 ? a.x - b.x : a.y - b.y));
    let prevBottom = null;
    let first = true;
    for (const o of objs) {
      const gap = prevBottom == null ? 0 : Math.max(0, o.y - prevBottom);
      if (o.type === 'image') {
        const k = Math.min(1, bodyW / o.w);
        const p = { ...para(''), pPr: { align: Math.abs(o.x + o.w / 2 - W / 2) < 30 ? 'center' : 'left', spc: { before: Math.min(24, gap), after: 0 } }, runs: [{ obj: 'img', id: uid('i'), media: o.media, w: o.w * k, h: o.h * k }] };
        if (first && si > 0) p.pPr.pageBreakBefore = true;
        body.push(p); first = false; prevBottom = o.y + o.h;
        continue;
      }
      const paras = o.text.paras;
      const blockRight = o.x + o.w;
      let cur = null;
      paras.forEach((p, li) => {
        const txt = plain(p);
        const size = Math.max(...p.runs.map((r) => r.size ?? bodySz));
        const align = p.align === 'ctr' ? 'center' : p.align === 'r' ? 'right' : null;
        const startNew = !cur || align || size > bodySz * 1.25 || /^[\s]*([•▪■●◆❖➢✓·\-–]|\d+[.)]|[가-하][.)]|\(\d+\)|[①-⑳])\s/.test(txt);
        if (startNew) {
          cur = { ...para(''), pPr: {}, runs: [] };
          const ind = o.x - sec.margin.l;
          if (ind > 8 && !align) cur.pPr.ind = { l: Math.round(ind) };
          if (align) cur.pPr.align = align;
          const lineGap = (p.lineSpacingPt ?? size * 1.2) * PT;
          cur.pPr.spc = { before: li === 0 ? Math.min(32, Math.round(gap)) : 0, after: 0, line: Math.max(1, Math.min(2.5, lineGap / (size * PT * 1.25))), rule: 'auto' };
          if (size >= bodySz * 1.6) cur.style = 'Heading1';
          else if (size >= bodySz * 1.25) cur.style = 'Heading2';
          if (first && si > 0) cur.pPr.pageBreakBefore = true;
          first = false;
          body.push(cur);
        } else if (cur.runs.length) {
          // 줄 이음: 한글은 줄 끝 공백 없이 이어지는 경우가 많음 → 영문 단어끼리만 공백
          const last = cur.runs[cur.runs.length - 1];
          const prevText = last.text ?? '';
          if (/[A-Za-z0-9,.;:!?)]$/.test(prevText) && /^[A-Za-z0-9(]/.test(txt)) cur.runs.push({ ...last, text: ' ' });
          else if (/[가-힣]$/.test(prevText) && /^[가-힣]/.test(txt) && !prevText.endsWith('-')) cur.runs.push({ ...last, text: ' ' });
        }
        cur.runs.push(...runsOf(p));
        // 줄이 블록 오른쪽 끝보다 많이 짧고 문장이 끝나면 문단 끝
        const lineW = p.runs.reduce((n, r) => n + r.t.length * (r.size ?? bodySz) * PT * 0.55, 0);
        if (/[.!?。:다요]\s*$/.test(txt) && o.x + lineW < blockRight - (o.w * 0.15)) cur = null;
      });
      for (const b of body.slice(-paras.length)) if (b.runs) b.runs = normalizeRuns(b.runs);
      prevBottom = o.y + o.h;
    }
  });
  doc.body = body.length ? body.map((b) => ({ ...b, runs: normalizeRuns(b.runs) })) : [para('')];
  if (pres.props?.title) doc.props.title = pres.props.title;
  if (pres.props?.author) doc.props.author = pres.props.author;
  return { doc, warnings: ['PDF 를 편집할 수 있는 문서로 바꾸어 열었습니다 (줄을 문단으로 다시 이음). 표 · 다단 · 도형은 모양이 달라질 수 있습니다.', ...warnings], jobs };
}
