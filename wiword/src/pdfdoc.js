// PDF → 편집할 수 있는 문서 (DOM 없음)
// pdf.js (WIPOINT 와 같은 해석기) 가 쪽마다 글 상자 · 그림을 만들면,
//  흐름(flow, 기본): 줄을 문단으로 다시 잇고 (줄 끝 · 들여쓰기 · 간격으로 판단), 큰 글자는 제목, 그림은 본문 안에 — Word 의 [PDF 다시 배치] 와 같은 방식
//  모양 유지(layout): 쪽마다 글 상자를 원래 자리에 (모양은 같지만 고치기는 불편)
import { readPdf } from './pdf.js';
import { newDoc, newSection, para, uid, normalizeRuns } from './model.js';
import { gridTable } from './hwp.js';

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
/** 줄이 목록 항목으로 시작하는가 (가. 나. 는 앞 줄이 문장으로 끝났을 때만 — "…했습니다. 두 줄" 같은 줄 이음을 목록으로 오인하지 않게) */
const listStart = (txt, prev = '') => /^\s*([•▪■●◆❖➢✓·\-–]|\d+[.)]|\(\d+\)|[①-⑳])\s/.test(txt) || (/^\s*[가-하][.)]\s/.test(txt) && (!prev || /[.!?:)]\s*$/.test(prev)));


// ───────────── 표 찾기 (괘선) ─────────────
/** 도형 → 가로 · 세로 괘선 조각 { o:'h'|'v', at, a, b } 과 칠한 사각형 */
function rulings(objs) {
  const segs = []; const fills = [];
  for (const o of objs) {
    if (o.type !== 'shape' || o.text) continue;
    const x = o.x; const y = o.y; const w = o.w; const h = o.h;
    if (o.shape === 'line' || (!o.path && (h <= 3 || w <= 3))) {
      if (h <= 3 && w > 3) segs.push({ o: 'h', at: y + h / 2, a: x, b: x + w });
      else if (w <= 3 && h > 3) segs.push({ o: 'v', at: x + w / 2, a: y, b: y + h });
      continue;
    }
    if (o.path) {
      for (const p of o.path) {
        let cx = 0; let cy = 0; let sx = 0; let sy = 0;
        for (const [op, ...v] of p.cmds) {
          if (op === 'M') { cx = v[0]; cy = v[1]; sx = cx; sy = cy; } else if (op === 'L' || op === 'Z') {
            const nx = op === 'Z' ? sx : v[0]; const ny = op === 'Z' ? sy : v[1];
            if (Math.abs(ny - cy) < 0.6 && Math.abs(nx - cx) > 3) segs.push({ o: 'h', at: y + cy, a: x + Math.min(cx, nx), b: x + Math.max(cx, nx) });
            else if (Math.abs(nx - cx) < 0.6 && Math.abs(ny - cy) > 3) segs.push({ o: 'v', at: x + cx, a: y + Math.min(cy, ny), b: y + Math.max(cy, ny) });
            cx = nx; cy = ny;
          } else if (op === 'C') { cx = v[4]; cy = v[5]; }
        }
      }
      continue;
    }
    if (o.line) { segs.push({ o: 'h', at: y, a: x, b: x + w }, { o: 'h', at: y + h, a: x, b: x + w }, { o: 'v', at: x, a: y, b: y + h }, { o: 'v', at: x + w, a: y, b: y + h }); }
    const fc = o.fill?.color;
    if (fc && !/^#(FFFFFF|FEFEFE)$/i.test(fc)) fills.push({ x0: x, y0: y, x1: x + w, y1: y + h, color: fc });
  }
  // 같은 줄의 조각 잇기
  const merged = [];
  for (const kind of ['h', 'v']) {
    const list = segs.filter((g) => g.o === kind).sort((a, b) => a.at - b.at || a.a - b.a);
    for (const g of list) {
      const m = merged.find((q) => q.o === kind && Math.abs(q.at - g.at) <= 1.5 && g.a <= q.b + 2 && g.b >= q.a - 2);
      if (m) { m.a = Math.min(m.a, g.a); m.b = Math.max(m.b, g.b); } else merged.push({ ...g });
    }
  }
  return { segs: merged, fills };
}
const uniq = (vals, tol = 2) => { const out = []; for (const v of [...vals].sort((a, b) => a - b)) if (!out.length || v - out[out.length - 1] > tol) out.push(v); return out; };
function findTables(segs, W, H) {
  const hs = segs.filter((g) => g.o === 'h'); const vs = segs.filter((g) => g.o === 'v');
  const parent = segs.map((_, i) => i);
  const root = (i) => { while (parent[i] !== i) i = parent[i] = parent[parent[i]]; return i; };
  const idx = new Map(segs.map((g, i) => [g, i]));
  for (const h of hs) for (const v of vs) if (v.at >= h.a - 2 && v.at <= h.b + 2 && h.at >= v.a - 2 && h.at <= v.b + 2) parent[root(idx.get(h))] = root(idx.get(v));
  const groups = new Map();
  segs.forEach((g, i) => { const r = root(i); if (!groups.has(r)) groups.set(r, []); groups.get(r).push(g); });
  const tables = [];
  for (const g of groups.values()) {
    const gh = g.filter((x) => x.o === 'h'); const gv = g.filter((x) => x.o === 'v');
    if (gh.length < 2 || gv.length < 2) continue;
    const xs = uniq(gv.map((v) => v.at)); const ys = uniq(gh.map((h) => h.at));
    if (xs.length < 2 || ys.length < 2) continue;
    const bw = xs[xs.length - 1] - xs[0]; const bh = ys[ys.length - 1] - ys[0];
    if (bw < 20 || bh < 8) continue;
    if (xs.length === 2 && ys.length === 2 && bw * bh > W * H * 0.4) continue; // 쪽 테두리
    const vOn = (k, r) => { const mid = (ys[r] + ys[r + 1]) / 2; return gv.some((v) => Math.abs(v.at - xs[k]) <= 2 && v.a <= mid && v.b >= mid); };
    const hOn = (r, c) => { const mid = (xs[c] + xs[c + 1]) / 2; return gh.some((h) => Math.abs(h.at - ys[r]) <= 2 && h.a <= mid && h.b >= mid); };
    const R = ys.length - 1; const Cn = xs.length - 1;
    const seen = Array.from({ length: R }, () => new Array(Cn).fill(false));
    const cells = [];
    for (let r = 0; r < R; r++) for (let c = 0; c < Cn; c++) {
      if (seen[r][c]) continue;
      let cs = 1; while (c + cs < Cn && !vOn(c + cs, r)) cs++;
      let rs = 1; while (r + rs < R && Array.from({ length: cs }, (_, k) => c + k).every((cc) => !hOn(r + rs, cc))) rs++;
      for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) seen[r + a][c + b] = true;
      cells.push({ r, c, rs, cs, x0: xs[c], x1: xs[c + cs], y0: ys[r], y1: ys[r + rs], lines: [] });
    }
    tables.push({ xs, ys, x0: xs[0], x1: xs[xs.length - 1], y0: ys[0], y1: ys[ys.length - 1], cells });
  }
  return tables;
}

// ───────────── 읽는 순서 (단 나누기) ─────────────
function readingOrder(items, W) {
  const narrow = items.filter((it) => it.x1 - it.x0 < W * 0.55);
  const total = narrow.reduce((n, it) => n + (it.y1 - it.y0), 0);
  let best = null;
  if (narrow.length >= 2 && total > 0) {
    let run = null;
    for (let x = Math.round(W * 0.3); x <= W * 0.7; x += 2) {
      const cross = narrow.some((it) => it.x0 < x - 2 && it.x1 > x + 2);
      if (!cross) {
        const left = narrow.reduce((n, it) => n + (it.x1 <= x ? it.y1 - it.y0 : 0), 0);
        const right = narrow.reduce((n, it) => n + (it.x0 >= x ? it.y1 - it.y0 : 0), 0);
        if (left > total * 0.2 && right > total * 0.2) { if (!run) run = { a: x, b: x }; else run.b = x; if (!best || run.b - run.a > best.b - best.a) best = { ...run }; continue; }
      }
      run = null;
    }
  }
  const byY = [...items].sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  if (!best) return { order: byY, cols: false };
  const g = (best.a + best.b) / 2;
  const out = []; let left = []; let right = [];
  const flush = () => { out.push(...left.sort((a, b) => a.y0 - b.y0), ...right.sort((a, b) => a.y0 - b.y0)); left = []; right = []; };
  for (const it of byY) {
    if (it.x0 < g - 2 && it.x1 > g + 2) { flush(); out.push(it); } else if (it.x1 <= g + 2) left.push(it); else right.push(it);
  }
  flush();
  return { order: out, cols: true };
}

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
  let colPages = 0; let nTables = 0;
  pres.slides.forEach((s, si) => {
    // 괘선 → 표, 표 안 줄은 칸으로
    const { segs, fills } = rulings(s.objects);
    const tables = findTables(segs, W, H);
    const texts = [];
    for (const o of s.objects.filter((x) => x.text)) {
      const n = o.text.paras.length; const lh = o.h / Math.max(1, n);
      let keep = [];
      const pushKeep = () => { if (keep.length) { const y0 = o.y + keep[0].i * lh; texts.push({ ...o, y: y0, h: keep.length * lh, text: { ...o.text, paras: keep.map((k) => k.p) } }); keep = []; } };
      o.text.paras.forEach((p, i) => {
        const cy = p.cy ?? o.y + (i + 0.5) * lh;
        const pcs = p.pieces ?? [{ t: plain(p), x0: o.x, x1: o.x + o.w }];
        const t = tables.find((tb) => cy > tb.y0 && cy < tb.y1 && pcs.some((q) => (q.x0 + q.x1) / 2 > tb.x0 && (q.x0 + q.x1) / 2 < tb.x1));
        if (!t) { keep.push({ p, i }); return; }
        pushKeep();
        // 조각을 칸별로 나눔 (한 줄이 여러 칸에 걸친 경우)
        const byCell = new Map();
        for (const q of pcs) {
          const cx = (q.x0 + q.x1) / 2;
          const cell = t.cells.find((c) => cx > c.x0 && cx < c.x1 && cy > c.y0 && cy < c.y1);
          if (!cell) continue;
          if (!byCell.has(cell)) byCell.set(cell, []);
          byCell.get(cell).push(q);
        }
        for (const [cell, qs] of byCell) {
          while (qs.length && !qs[0].t.trim()) qs.shift();
          while (qs.length && !qs[qs.length - 1].t.trim()) qs.pop();
          if (!qs.length) continue;
          const runs = qs.map(({ t: tt, x0, x1, ...rest }) => ({ t: tt, ...rest }));
          cell.lines.push({ p: { runs }, y: cy, x: qs[0].x0, w: qs[qs.length - 1].x1 - qs[0].x0, n: 1 });
        }
      });
      pushKeep();
    }
    const items = [
      ...texts.map((o) => ({ kind: 'text', o, x0: o.x, x1: o.x + o.w, y0: o.y, y1: o.y + o.h })),
      ...s.objects.filter((o) => o.type === 'image').map((o) => ({ kind: 'image', o, x0: o.x, x1: o.x + o.w, y0: o.y, y1: o.y + o.h })),
      ...tables.map((t) => ({ kind: 'table', t, x0: t.x0, x1: t.x1, y0: t.y0, y1: t.y1 })),
    ];
    const ro = readingOrder(items, W);
    if (ro.cols) colPages++;
    let prevBottom = null;
    let first = true;
    for (const it of ro.order) {
      if (it.kind === 'table') {
        const t = it.t;
        const cells = t.cells.map((c) => {
          const lines = c.lines.sort((a, b) => a.y - b.y);
          const blocks = [];
          let cur = null;
          for (const l of lines) {
            const txt = plain(l.p);
            const gl = l.x - c.x0; const gr = c.x1 - (l.x + l.w);
            const align = gl > 12 && gr > 12 && Math.abs(gl - gr) < 4 ? 'center' : gr < 12 && gl > gr + 10 ? 'right' : null;
            if (!cur || listStart(txt, cur.runs.map((r) => r.text ?? '').join(''))) { cur = { ...para(''), pPr: { spc: { before: 0, after: 0 }, ...(align ? { align } : {}) }, runs: [] }; blocks.push(cur); }
            else if (cur.runs.length) { const pt = cur.runs[cur.runs.length - 1].text ?? ''; if ((/[A-Za-z0-9,.;:!?)]$/.test(pt) && /^[A-Za-z0-9(]/.test(txt)) || (/[가-힣]$/.test(pt) && /^[가-힣]/.test(txt))) cur.runs.push({ ...cur.runs[cur.runs.length - 1], text: ' ' }); }
            cur.runs.push(...runsOf(l.p));
          }
          for (const b of blocks) b.runs = normalizeRuns(b.runs);
          const fill = fills.find((f) => f.x0 <= c.x0 + 2 && f.x1 >= c.x1 - 2 && f.y0 <= c.y0 + 2 && f.y1 >= c.y1 - 2);
          const line = { style: 'single', w: 0.75, color: '#000000' };
          return { r: c.r, c: c.c, rs: c.rs, cs: c.cs, h: c.rs === 1 ? Math.round(c.y1 - c.y0) : 0, blocks, borders: { top: line, bottom: line, left: line, right: line }, ...(fill ? { shd: fill.color } : {}) };
        });
        const grid = t.xs.slice(1).map((x, k) => Math.max(8, Math.round(x - t.xs[k])));
        const tb = gridTable(cells, t.ys.length - 1, t.xs.length - 1, grid, { look: { firstRow: false, firstCol: false, bandRows: false }, cellMar: { l: 5, r: 5, t: 2, b: 2 }, ...(Math.abs((t.x0 + t.x1) / 2 - W / 2) < 20 && t.x1 - t.x0 < bodyW * 0.9 ? { align: 'center' } : {}), ...(t.x0 - sec.margin.l > 8 ? { ind: Math.round(t.x0 - sec.margin.l) } : {}) });
        if (first && si > 0) body.push({ ...para(''), pPr: { pageBreakBefore: true, spc: { before: 0, after: 0, line: 1, rule: 'auto' } } });
        body.push(tb);
        nTables++;
        first = false; prevBottom = t.y1;
        continue;
      }
      const o = it.o;
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
        const startNew = !cur || align || size > bodySz * 1.25 || listStart(txt, cur ? cur.runs.map((r) => r.text ?? '').join('') : '');
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
  doc.body = body.length ? body.map((b) => (b.runs ? { ...b, runs: normalizeRuns(b.runs) } : b)) : [para('')];
  if (pres.props?.title) doc.props.title = pres.props.title;
  if (pres.props?.author) doc.props.author = pres.props.author;
  const notes = [`PDF 를 편집할 수 있는 문서로 바꾸어 열었습니다 (줄을 문단으로 다시 이음${nTables ? `, 표 ${nTables}개` : ''}${colPages ? `, 두 단 ${colPages}쪽은 읽는 순서대로` : ''}). 도형과 글자 위치는 원본과 다를 수 있습니다 — 원래 모양이 필요하면 [모양 유지]로 여세요.`];
  return { doc, warnings: [...notes, ...warnings], jobs };
}
