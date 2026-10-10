// 슬라이드 → 진짜 SVG (DOM 없음): 도형은 path, 글은 <text>/<tspan> (줄 바꿈 계산), 그림은 <image>, 표는 사각형 + 글, 차트는 차트 SVG
//   foreignObject 를 쓰지 않으므로 일러스트레이터 · 잉크스케이프 · 피그마에서도 열림.
//   measure(fontCss, text) → 너비(px): 브라우저는 canvas measureText, 없으면 근사값
import { resolveColor } from './themes.js';
import { shapePath, customPath, textRect, OPEN_SHAPES } from './shapes.js';
import { PX_PER_PT, defaultSize, defaultColor, defaultFont, isEmptyText, setMasterText } from './model.js';
import { svgFill, markers, DASH, fontStack, paraDefaults, numberLabel, slideBackground, slideDecor, tableCellStyle } from './render.js';
import { chartSvg } from './chart.js';
import { esc } from './xml.js';

const f = (n) => Math.round(n * 100) / 100;
const CJK = /[ᄀ-ᇿ　-鿿가-힯＀-￯]/;
export const approxMeasure = (font, text) => {
  const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? 16);
  let w = 0;
  for (const ch of text) w += CJK.test(ch) ? size : ch === ' ' ? size * 0.28 : /[A-Z0-9]/.test(ch) ? size * 0.62 : size * 0.52;
  return w;
};

/** 글 상자 → <text> 들 (x0,y0 = 글 영역 왼쪽 위, w,h = 글 영역 크기) */
function textSvg(theme, o, body, x0, y0, w, h, measure) {
  const fsc = body.fontScale ?? 1;
  const lines = [];
  const counters = [];
  const wrap = body.wrap !== false;
  body.paras.forEach((p, pi) => {
    const d = paraDefaults(o, p);
    const lvl = p.lvl ?? 0;
    let label = '';
    if (d.bullet?.type === 'num') { counters.length = lvl + 1; counters[lvl] = (counters[lvl] ?? (d.bullet.start ?? 1) - 1) + 1; label = numberLabel(d.bullet.scheme, counters[lvl]); }
    else { counters.length = Math.min(counters.length, lvl); if (d.bullet?.type === 'char') label = d.bullet.char ?? '•'; }
    const baseSize = (p.end?.size ?? p.runs[0]?.size ?? defaultSize(o, lvl)) * fsc;
    const runs = p.runs.map((r) => {
      const size = (r.size ?? defaultSize(o, lvl)) * fsc * (r.base ? 0.7 : 1);
      const font = `${r.i ? 'italic ' : ''}${r.b ? 700 : 400} ${f(size * PX_PER_PT)}px ${fontStack(theme, r.font ?? defaultFont(o))}`;
      return { r, size, font, color: resolveColor(theme, r.color ?? defaultColor(o), '#000') };
    });
    const lh = (d.lineSpacingPt ? d.lineSpacingPt : 1.2 * d.lineSpacing * Math.max(baseSize, ...runs.map((x) => x.size))) * PX_PER_PT;
    const left = Math.max(0, d.marL);
    const avail = Math.max(10, w - left);
    // 글자 조각 단위로 줄 나누기 (띄어쓰기에서, 너무 긴 낱말은 글자에서)
    let line = { segs: [], width: 0, first: true };
    const flush = () => { lines.push({ ...line, p, d, lh, label: line.first ? label : '', spcBef: line.first && pi ? d.spcBef * PX_PER_PT : 0, align: p.align ?? 'l', left }); line = { segs: [], width: 0, first: false }; };
    const indentOf = (ln) => (ln.first ? d.indent : 0);
    for (const run of runs) {
      const parts = String(run.r.t ?? '').split(/(\s+|\v)/);
      for (const part of parts) {
        if (!part) continue;
        if (part === '\v') { flush(); continue; }
        let pw = measure(run.font, part);
        if (wrap && line.segs.length && line.width + pw > avail - indentOf(line) && !/^\s+$/.test(part)) flush();
        if (wrap && pw > avail - indentOf(line)) {
          // 긴 낱말 · 줄 바꿈 없는 한글 문장: 글자 단위
          let buf = '';
          for (const ch of part) {
            const cw = measure(run.font, buf + ch);
            if (buf && line.width + cw > avail - indentOf(line)) { line.segs.push({ run, t: buf }); line.width += measure(run.font, buf); flush(); buf = ch; } else buf += ch;
          }
          if (buf) { line.segs.push({ run, t: buf }); line.width += measure(run.font, buf); }
          continue;
        }
        if (/^\s+$/.test(part) && !line.segs.length && !line.first) continue;
        line.segs.push({ run, t: part });
        line.width += pw;
      }
    }
    flush();
  });
  const total = lines.reduce((a, l) => a + l.lh + l.spcBef, 0);
  let y = y0 + ({ t: 0, ctr: (h - total) / 2, b: h - total }[body.anchor ?? 't'] ?? 0);
  let out = '';
  for (const l of lines) {
    y += l.spcBef;
    const base = y + l.lh * 0.8;
    const ind = l.first ? l.d.indent : 0;
    const startX = x0 + l.left + ind;
    const anchor = { ctr: 'middle', r: 'end' }[l.align] ?? 'start';
    const ax = anchor === 'middle' ? x0 + l.left + (w - l.left) / 2 : anchor === 'end' ? x0 + w : startX;
    if (l.label) {
      const r0 = l.segs[0]?.run ?? { font: `${f(l.lh / 1.2)}px sans-serif`, color: '#000' };
      const lx = l.d.bullet?.type ? x0 + l.left + l.d.indent : startX;
      out += `<text x="${f(lx)}" y="${f(base)}" style="font:${esc(r0.font)}" fill="${resolveColor(theme, l.d.bullet?.color ?? null, r0.color)}">${esc(l.label)}</text>`;
    }
    const spans = l.segs.map((s) => {
      const r = s.run.r;
      const deco = [r.u ? 'underline' : '', r.s ? 'line-through' : ''].filter(Boolean).join(' ');
      return `<tspan style="font:${esc(s.run.font)}${deco ? `;text-decoration:${deco}` : ''}" fill="${s.run.color}"${r.base ? ` baseline-shift="${r.base > 0 ? 'super' : 'sub'}"` : ''}>${esc(s.t)}</tspan>`;
    }).join('');
    if (spans) out += `<text x="${f(ax)}" y="${f(base)}" text-anchor="${anchor}" xml:space="preserve">${spans}</text>`;
    y += l.lh;
  }
  return out;
}

function objectSvg(pres, o, ctx) {
  const { theme } = pres;
  const media = pres.media ?? {};
  const defs = ctx.defs;
  const w = Math.max(o.w, 0.01);
  const h = Math.max(o.h, 0.01);
  const tf = [`translate(${f(o.x)} ${f(o.y)})`];
  if (o.rot) tf.push(`rotate(${f(o.rot)} ${f(w / 2)} ${f(h / 2)})`);
  if (o.flipH || o.flipV) tf.push(`translate(${o.flipH ? f(w) : 0} ${o.flipV ? f(h) : 0}) scale(${o.flipH ? -1 : 1} ${o.flipV ? -1 : 1})`);
  let inner = '';
  if (o.type === 'image' || o.type === 'media') {
    const src = media[o.media] ?? media[o.poster] ?? o.src;
    if (src) {
      const c = o.crop ?? { l: 0, t: 0, r: 0, b: 0 };
      const iw = w / Math.max(0.01, 1 - c.l - c.r);
      const ih = h / Math.max(0.01, 1 - c.t - c.b);
      const clipId = `c${ctx.n++}`;
      defs.push(`<clipPath id="${clipId}"><path d="${shapePath(o.shape ?? 'rect', w, h, o.adj)}"/></clipPath>`);
      inner += `<g clip-path="url(#${clipId})"${o.alpha != null ? ` opacity="${o.alpha}"` : ''}><image href="${esc(src)}" x="${f(-c.l * iw)}" y="${f(-c.t * ih)}" width="${f(iw)}" height="${f(ih)}" preserveAspectRatio="none"/></g>`;
    }
    if (o.line) inner += `<path d="${shapePath(o.shape ?? 'rect', w, h, o.adj)}" fill="none" stroke="${resolveColor(theme, o.line.color, '#000')}" stroke-width="${f(o.line.width ?? 1)}"/>`;
  } else if (o.type === 'chart') {
    inner += chartSvg(o.chart, w, h, theme, { font: theme.fonts.minor });
  } else if (o.type === 'table') {
    let y = 0;
    o.rows.forEach((row, ri) => {
      let x = 0;
      row.cells.forEach((c, ci) => {
        const cw = o.cols.slice(ci, ci + (c.span ?? 1)).reduce((a, b) => a + b, 0);
        const ch = o.rows.slice(ri, ri + (c.rowSpan ?? 1)).reduce((a, r) => a + r.h, 0);
        if (!c.hMerge && !c.vMerge) {
          const st = tableCellStyle(theme, o, ri, ci);
          const fill = c.fill !== undefined ? c.fill : st.fill;
          inner += `<rect x="${f(x)}" y="${f(y)}" width="${f(cw)}" height="${f(ch)}" fill="${fill ? resolveColor(theme, fill, 'none') : 'none'}" stroke="${resolveColor(theme, st.border, '#fff')}" stroke-width="1"/>`;
          if (c.text && !isEmptyText(c.text)) {
            const pseudo = { ...o, ph: null, text: c.text };
            const body = { ...c.text, paras: c.text.paras.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r, color: r.color ?? st.color, b: r.b ?? (st.bold || undefined) })) })) };
            inner += textSvg(theme, pseudo, body, x + 9.6, y + 4.8, cw - 19.2, ch - 9.6, ctx.measure);
          }
        }
        x += o.cols[ci] ?? 0;
      });
      y += row.h;
    });
  } else if (o.type === 'zoom') {
    const target = pres.slides.find((s) => s.id === o.zoom?.slide);
    if (target && !ctx.nested) inner += `<svg width="${f(w)}" height="${f(h)}" viewBox="0 0 ${pres.size.w} ${pres.size.h}" preserveAspectRatio="none">${slideBody(pres, target, { ...ctx, nested: true })}</svg>`;
    inner += `<rect width="${f(w)}" height="${f(h)}" fill="none" stroke="#bfbfbf"/>`;
  } else if (o.type === 'equation') {
    const size = (o.size ?? 28) * PX_PER_PT;
    inner += `<text x="${f(w / 2)}" y="${f(h / 2 + size * 0.35)}" text-anchor="middle" style="font:italic ${f(size)}px 'Cambria Math', serif" fill="${resolveColor(theme, o.color ?? '@tx1', '#000')}">${esc(o.latex ?? '')}</text>`;
  } else {
    const d = o.path ? customPath(o.path, w, h) : shapePath(o.shape ?? 'rect', w, h, o.adj);
    const open = OPEN_SHAPES.has(o.shape) || o.openPath;
    const local = [];
    const fill = open ? 'none' : svgFill(theme, o.fill, w, h, media, local);
    const ln = o.line;
    const stroke = ln ? resolveColor(theme, ln.color, 'none') : 'none';
    const sw = ln ? Math.max(0.5, ln.width ?? 1) : 0;
    const dash = ln && DASH[ln.dash] ? ` stroke-dasharray="${DASH[ln.dash].split(' ').map((x) => f(Number(x) * sw)).join(' ')}"` : '';
    const mk = ln ? markers(theme, ln, local) : '';
    defs.push(...local);
    if (fill !== 'none' || stroke !== 'none') inner += `<path d="${d}" fill="${fill}" fill-rule="evenodd"${fill !== 'none' && o.fill?.alpha != null ? ` fill-opacity="${o.fill.alpha}"` : ''} stroke="${stroke}" stroke-width="${f(sw)}"${dash} stroke-linejoin="round" ${mk}/>`;
  }
  if (o.text && !isEmptyText(o.text) && o.type !== 'table') {
    const [l, t, r, b] = o.txBox || o.ph ? [0, 0, w, h] : textRect(o.shape ?? 'rect', w, h);
    const ins = o.text.insets ?? [9.6, 4.8, 9.6, 4.8];
    inner += textSvg(theme, o, o.text, l + ins[0], t + ins[1], Math.max(1, r - l - ins[0] - ins[2]), Math.max(1, b - t - ins[1] - ins[3]), ctx.measure);
  }
  return `<g transform="${tf.join(' ')}"${o.name ? ` id="${esc(String(o.name).replace(/\s+/g, '_'))}"` : ''}>${inner}</g>`;
}

function slideBody(pres, slide, ctx) {
  const { w, h } = pres.size;
  let out = '';
  if (!ctx.only) {
    const bg = slideBackground(pres, slide);
    const local = [];
    const fill = bg ? svgFill(pres.theme, bg, w, h, pres.media ?? {}, local) : '#ffffff';
    ctx.defs.push(...local);
    out += `<rect width="${w}" height="${h}" fill="${fill === 'none' ? '#ffffff' : fill}"/>`;
    for (const o of slideDecor(pres, slide)) out += objectSvg(pres, o, ctx);
  }
  for (const o of slide.objects) {
    if (o.hidden || (ctx.only && !ctx.only.has(o.id))) continue;
    if (o.ph && (o.type === 'image' ? !o.media : !o.text || isEmptyText(o.text))) continue;
    out += objectSvg(pres, o, ctx);
  }
  return out;
}

/**
 * 슬라이드 (또는 일부 개체) → SVG 문서 문자열
 * opts: { only: Set(개체 id) — 그 개체만, 그 범위로 잘라냄 · measure · scale }
 */
export function slideSvg(pres, slide, opts = {}) {
  setMasterText(pres);
  const ctx = { defs: [], n: 0, measure: opts.measure ?? approxMeasure, only: opts.only ?? null };
  const body = slideBody(pres, slide, ctx);
  let vb = [0, 0, pres.size.w, pres.size.h];
  if (ctx.only) {
    const objs = slide.objects.filter((o) => ctx.only.has(o.id));
    const pad = 4;
    const x1 = Math.min(...objs.map((o) => o.x)) - pad; const y1 = Math.min(...objs.map((o) => o.y)) - pad;
    const x2 = Math.max(...objs.map((o) => o.x + o.w)) + pad; const y2 = Math.max(...objs.map((o) => o.y + o.h)) + pad;
    vb = [x1, y1, x2 - x1, y2 - y1];
  }
  const sc = opts.scale ?? 1;
  return `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${f(vb[2] * sc)}" height="${f(vb[3] * sc)}" viewBox="${vb.map(f).join(' ')}">${ctx.defs.length ? `<defs>${ctx.defs.join('')}</defs>` : ''}${body}</svg>`;
}
