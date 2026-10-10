// SmartArt 비슷한 도식 만들기 (도형 묶음으로 생성 → 일반 도형처럼 편집 · pptx 그룹으로 저장) — DOM 없음
import { newShape, uid, para } from './model.js';

export const SMART_KINDS = [
  ['process', '기본 프로세스'], ['chevron', '갈매기형 프로세스'], ['list', '세로 목록'], ['blocks', '기본 블록 목록'],
  ['cycle', '기본 주기형'], ['hierarchy', '조직도'], ['pyramid', '기본 피라미드'], ['timeline', '기본 타임라인'], ['venn', '기본 벤형'],
];

const ACC = ['@accent1', '@accent2', '@accent3', '@accent4', '@accent5', '@accent6'];

function box(kind, x, y, w, h, text, color, opts = {}) {
  const o = newShape(kind, { x, y, w, h });
  o.fill = { type: 'solid', color };
  o.line = opts.line === undefined ? { color: '@lt1', width: 1.33 } : opts.line;
  o.text.paras = [para(text, { align: 'ctr' }, { size: opts.size ?? 18, ...(opts.b ? { b: true } : {}) })];
  o.text.anchor = 'ctr';
  o.text.defColor = opts.textColor ?? '@lt1';
  return o;
}

/**
 * kind, items: 글 목록, rect: {x, y, w, h}, opts: { multicolor }
 * 반환: 개체 배열 (같은 grp)
 */
export function smartArt(kind, items, rect, { multicolor = true } = {}) {
  const list = items.map((s) => String(s).trim()).filter(Boolean);
  if (!list.length) list.push('텍스트');
  const n = list.length;
  const col = (i) => (multicolor ? ACC[i % ACC.length] : '@accent1');
  const { x, y, w, h } = rect;
  const fs = Math.max(11, Math.min(24, Math.round(Math.min(w / n, h) / 7)));
  const out = [];
  switch (kind) {
    case 'chevron': {
      const gap = w * 0.01;
      const bw = (w - gap * (n - 1)) / n;
      const bh = Math.min(h, bw * 0.55);
      list.forEach((t, i) => out.push(box(i === 0 ? 'homePlate' : 'chevron', x + i * (bw + gap), y + (h - bh) / 2, bw, bh, t, col(i), { size: fs, line: null })));
      break;
    }
    case 'list': {
      const gap = h * 0.03;
      const bh = (h - gap * (n - 1)) / n;
      list.forEach((t, i) => {
        out.push(box('roundRect', x, y + i * (bh + gap), w * 0.18, bh, String(i + 1), col(i), { size: fs + 4, b: true }));
        const o = box('rect', x + w * 0.19, y + i * (bh + gap), w * 0.81, bh, t, `${col(i)}:lm20:lo80`, { size: fs, textColor: '@tx1', line: null });
        o.text.paras[0].align = 'l';
        out.push(o);
      });
      break;
    }
    case 'blocks': {
      const cols = n <= 3 ? n : n === 4 ? 2 : 3;
      const rows = Math.ceil(n / cols);
      const gx = w * 0.03;
      const gy = h * 0.05;
      const bw = (w - gx * (cols - 1)) / cols;
      const bh = Math.min((h - gy * (rows - 1)) / rows, bw * 0.75);
      list.forEach((t, i) => out.push(box('rect', x + (i % cols) * (bw + gx), y + Math.floor(i / cols) * (bh + gy), bw, bh, t, col(i), { size: fs })));
      break;
    }
    case 'cycle': {
      const r = Math.min(w, h) / 2;
      const cx = x + w / 2;
      const cy = y + h / 2;
      const d = Math.min(r * 0.8, (Math.PI * 2 * r * 0.75) / n / 1.15);
      const ring = r - d / 2;
      list.forEach((t, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        out.push(box('ellipse', cx + Math.cos(a) * ring - d / 2, cy + Math.sin(a) * ring - d / 2, d, d, t, col(i), { size: Math.max(10, fs - 2) }));
        const am = a + Math.PI / n;
        const ar = newShape('rightArrow', { x: cx + Math.cos(am) * ring - d * 0.18, y: cy + Math.sin(am) * ring - d * 0.12, w: d * 0.36, h: d * 0.24 });
        ar.rot = Math.round(((am * 180) / Math.PI + 90 + 360) % 360);
        ar.fill = { type: 'solid', color: '@tx1:lm25:lo75' };
        ar.line = null;
        ar.text = null;
        out.push(ar);
      });
      break;
    }
    case 'hierarchy': {
      const [top, ...kids] = list;
      const bh = Math.min(h * 0.28, 90);
      const tw = Math.min(w * 0.4, 280);
      out.push(box('roundRect', x + (w - tw) / 2, y, tw, bh, top, '@accent1', { size: fs, b: true }));
      const m = Math.max(1, kids.length);
      const gap = w * 0.03;
      const kw = Math.min((w - gap * (m - 1)) / m, 260);
      const total = kw * m + gap * (m - 1);
      const ky = y + h - bh;
      const busY = y + bh + (ky - y - bh) / 2;
      const line = (x1, y1, x2, y2) => { const l = newShape('line', { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) }); l.line = { color: '@tx1:lm50:lo50', width: 1.5 }; out.push(l); };
      if (kids.length) {
        line(x + w / 2, y + bh, x + w / 2, busY);
        const first = x + (w - total) / 2 + kw / 2;
        const last = first + (m - 1) * (kw + gap);
        if (m > 1) line(first, busY, last, busY);
      }
      kids.forEach((t, i) => {
        const kx = x + (w - total) / 2 + i * (kw + gap);
        line(kx + kw / 2, busY, kx + kw / 2, ky);
        out.push(box('roundRect', kx, ky, kw, bh, t, multicolor ? ACC[(i + 1) % ACC.length] : '@accent1:lm75', { size: Math.max(10, fs - 2) }));
      });
      break;
    }
    case 'pyramid': {
      const bh = h / n;
      list.forEach((t, i) => {
        const tw = w * ((i + 1) / n);
        const kind2 = i === 0 ? 'triangle' : 'trapezoid';
        const o = box(kind2, x + (w - tw) / 2, y + i * bh, tw, bh, t, col(i), { size: fs });
        if (i > 0) o.adj = { adj: Math.round(((w / n / 2) / Math.min(tw, bh)) * 100000) };
        out.push(o);
      });
      break;
    }
    case 'timeline': {
      const ly = y + h / 2;
      const l = newShape('straightConnector1', { x, y: ly, w, h: 0 });
      l.line = { color: '@tx1:lm50:lo50', width: 3, tail: 'triangle' };
      out.push(l);
      const step = w / (n + 0.5);
      list.forEach((t, i) => {
        const cx = x + step * (i + 0.6);
        const dot = box('ellipse', cx - 12, ly - 12, 24, 24, '', col(i), { line: { color: '@lt1', width: 3 } });
        dot.text = null;
        out.push(dot);
        const up = i % 2 === 0;
        const tb = box('rect', cx - step * 0.48, up ? ly - h * 0.42 : ly + h * 0.1, step * 0.96, h * 0.32, t, '@lt1', { size: fs, textColor: '@tx1', line: null });
        tb.fill = null;
        tb.text.anchor = up ? 'b' : 't';
        out.push(tb);
      });
      break;
    }
    case 'venn': {
      const d = Math.min(h, (w / (n * 0.7 + 0.3)));
      const total = d * (0.7 * (n - 1) + 1);
      list.forEach((t, i) => {
        const o = box('ellipse', x + (w - total) / 2 + i * d * 0.7, y + (h - d) / 2, d, d, t, `${col(i)}:a55`, { size: fs, textColor: '@tx1', line: null });
        out.push(o);
      });
      break;
    }
    default: {
      // 기본 프로세스: 상자 → 화살표 → 상자
      const aw = w * 0.06;
      const bw = (w - aw * (n - 1) * 1.6) / n;
      const bh = Math.min(h, bw * 0.65);
      list.forEach((t, i) => {
        const bx = x + i * (bw + aw * 1.6);
        out.push(box('roundRect', bx, y + (h - bh) / 2, bw, bh, t, col(i), { size: fs }));
        if (i < n - 1) {
          const a = newShape('rightArrow', { x: bx + bw + aw * 0.3, y: y + h / 2 - aw * 0.6, w: aw, h: aw * 1.2 });
          a.fill = { type: 'solid', color: '@tx1:lm25:lo75' };
          a.line = null;
          a.text = null;
          out.push(a);
        }
      });
    }
  }
  const g = uid('g');
  for (const o of out) { o.grp = g; o.smart = kind; }
  if (out[0]) { out[0].smartItems = list; out[0].smartMulti = multicolor; }
  return out;
}

