// 문서 → PDF 파일 (브라우저): 쪽마다 화면과 같은 HTML 을 그린 뒤 글자 · 배경 · 테두리 · 그림 · 연결의 위치를 읽어 pdfwrite 로 옮김
// 인쇄 대화상자를 거치지 않고 바로 .pdf 를 만듦 (글자는 고르고 찾을 수 있는 글자로, 제목은 PDF 책갈피로)
import { S } from './state.js';
import { PdfWriter } from './pdfwrite.js';
import { pageNode, Layout } from './layout.js';
import { relayout } from './view.js';
import { walkParas, headingLevel, paraPlain } from './model.js';
import { fontFileFor } from './fontlib.js';
import { openTtf } from './ttfsubset.js';

const SERIF = /바탕|명조|궁서|batang|myeongjo|gungsuh|serif|times|georgia|cambria|garamond|함초롬바탕|book antiqua|palatino/i;
const parseColor = (s) => {
  const m = /rgba?\(([^)]+)\)/.exec(s ?? '');
  if (!m) return null;
  const [r, g, b, a = '1'] = m[1].split(/[\s,/]+/).filter(Boolean);
  const alpha = Number(a);
  if (alpha === 0) return null;
  return { hex: `#${[r, g, b].map((v) => Math.round(Number(v)).toString(16).padStart(2, '0')).join('')}`, alpha };
};
const inter = (a, b) => { if (!b) return a; const x = Math.max(a.x, b.x); const y = Math.max(a.y, b.y); const x2 = Math.min(a.x + a.w, b.x + b.w); const y2 = Math.min(a.y + a.h, b.y + b.h); return { x, y, w: Math.max(0, x2 - x), h: Math.max(0, y2 - y) }; };
const b64bytes = (s) => { const bin = atob(s); const o = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) o[i] = bin.charCodeAt(i); return o; };
/** JPEG 의 크기 · 색 성분 (SOF) */
function jpegInfo(b) {
  for (let o = 2; o + 9 < b.length;) {
    if (b[o] !== 0xff) { o++; continue; }
    const m = b[o + 1]; const len = (b[o + 2] << 8) | b[o + 3];
    if ((m >= 0xc0 && m <= 0xc3) || (m >= 0xc5 && m <= 0xc7) || (m >= 0xc9 && m <= 0xcb) || (m >= 0xcd && m <= 0xcf)) return { h: (b[o + 5] << 8) | b[o + 6], w: (b[o + 7] << 8) | b[o + 8], comps: b[o + 9] };
    o += 2 + len;
  }
  return null;
}

/** 그림 요소 → writer 그림 번호 (같은 그림은 한 번만) */
async function imageId(w, cache, src, el) {
  if (cache.has(src)) return cache.get(src);
  let id = null;
  const jm = /^data:image\/jpe?g;base64,(.*)$/s.exec(src);
  if (jm) {
    const bytes = b64bytes(jm[1]); const info = jpegInfo(bytes);
    if (info && info.comps !== 4) id = w.addImage({ jpeg: bytes, w: info.w, h: info.h, gray: info.comps === 1 });
  }
  if (id == null) {
    try {
      const img = el instanceof HTMLImageElement && el.complete && el.naturalWidth ? el : await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
      let nw = img.naturalWidth || 300; let nh = img.naturalHeight || 150;
      if (/^data:image\/svg/.test(src) || !nw) { const r = el.getBoundingClientRect(); nw = Math.max(1, Math.round(r.width * 3)); nh = Math.max(1, Math.round(r.height * 3)); }
      const k = Math.min(1, 2400 / Math.max(nw, nh)); const cw = Math.max(1, Math.round(nw * k)); const ch = Math.max(1, Math.round(nh * k));
      const c = document.createElement('canvas'); c.width = cw; c.height = ch;
      const g = c.getContext('2d'); g.drawImage(img, 0, 0, cw, ch);
      const d = g.getImageData(0, 0, cw, ch).data;
      const rgb = new Uint8Array(cw * ch * 3); const alpha = new Uint8Array(cw * ch); let hasA = false;
      for (let i = 0; i < cw * ch; i++) { rgb[i * 3] = d[i * 4]; rgb[i * 3 + 1] = d[i * 4 + 1]; rgb[i * 3 + 2] = d[i * 4 + 2]; alpha[i] = d[i * 4 + 3]; if (d[i * 4 + 3] < 255) hasA = true; }
      id = w.addImage({ rgb, alpha: hasA ? alpha : null, w: cw, h: ch });
    } catch { id = null; }
  }
  cache.set(src, id);
  return id;
}

/** 요소 하나를 그림으로 (수식 · SVG 처럼 글자와 선을 그대로 옮기기 어려운 것) */
async function rasterize(el) {
  const r = el.getBoundingClientRect();
  if (!r.width || !r.height) return null;
  let svg;
  if (el instanceof SVGSVGElement) {
    const c = el.cloneNode(true); c.setAttribute('xmlns', 'http://www.w3.org/2000/svg'); c.setAttribute('width', r.width); c.setAttribute('height', r.height);
    svg = new XMLSerializer().serializeToString(c);
  } else {
    // foreignObject 로 감싸기 (계산된 글자 모양을 그대로)
    const clone = el.cloneNode(true);
    const cs = getComputedStyle(el);
    clone.setAttribute('style', `${clone.getAttribute('style') ?? ''};font:${cs.font};color:${cs.color};margin:0`);
    svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${r.width}" height="${r.height}"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml">${new XMLSerializer().serializeToString(clone)}</div></foreignObject></svg>`;
  }
  return `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`;
}

// 글자가 첫 글꼴에 없을 때 차례로 찾아볼 글꼴 (한글 · 라틴)
const FALLBACK = ['Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', 'AppleGothic', 'Noto Sans KR', 'Noto Sans CJK KR', 'Source Han Sans KR', 'NanumGothic', 'Nanum Gothic', 'UnDotum', 'Baekmuk Gulim', 'WenQuanYi Zen Hei', 'Arial Unicode MS', 'Arial', 'Liberation Sans', 'DejaVu Sans', 'Helvetica'];
/** 글꼴 이름 → 넣을 글꼴 참조 (없으면 null) — 문서 하나에서 같은 글꼴은 한 번만 */
async function fontRef(w, refs, family, bold, italic, ask) {
  const key = `${family}|${bold}|${italic}`.toLowerCase();
  if (refs.has(key)) return refs.get(key);
  let ref = null;
  const file = await fontFileFor(family, { bold, italic, ask });
  if (file) {
    const f = openTtf(file.bytes, file.ps);
    if (f && !(f.fsType & 0x0002 && !(f.fsType & 0x000c))) {
      // 같은 글꼴 파일이면 참조 하나 (굵게를 흉내 내는 경우 등)
      const same = [...refs.values()].find((r) => r && r.font.psName === f.psName && r.font.bytes.length === f.bytes.length);
      ref = same ? { ...same, fake: { bold: file.fakeBold, italic: file.fakeItalic } } : Object.assign(w.embedFont(f), { fake: { bold: file.fakeBold, italic: file.fakeItalic } });
      if (same) ref.used = same.used; // 같은 사용 목록을 나눠 씀
    }
  }
  refs.set(key, ref);
  return ref;
}

/** 쪽 DOM → PDF 쪽 */
async function capturePage(w, pg, node, cache, outline, pageIndex, levels, refs, opts) {
  const O = node.getBoundingClientRect();
  const R = (r) => ({ x: r.left - O.left, y: r.top - O.top, w: r.width, h: r.height });
  const pageClip = { x: 0, y: 0, w: O.width, h: O.height };
  const range = document.createRange();
  const bgOf = (el) => parseColor(getComputedStyle(el).backgroundColor);
  const pbg = bgOf(node);
  if (pbg && pbg.hex !== '#ffffff') pg.rect(0, 0, O.width, O.height, pbg.hex, pbg.alpha);
  const seenHead = new Set();
  const walk = async (el, clip) => {
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden' || Number(cs.opacity) === 0) return;
    const box = R(el.getBoundingClientRect());
    // 배경
    if (el !== node) {
      const bg = parseColor(cs.backgroundColor);
      if (bg) {
        const rects = cs.display === 'inline' ? [...el.getClientRects()].map(R) : [box];
        for (const r of rects) { const c = inter(r, clip); if (c.w > 0 && c.h > 0) pg.rect(c.x, c.y, c.w, c.h, bg.hex, bg.alpha); }
      }
      // 테두리
      if (cs.display !== 'inline') {
        for (const [side, x1, y1, x2, y2] of [['Top', 0, 0, 1, 0], ['Bottom', 0, 1, 1, 1], ['Left', 0, 0, 0, 1], ['Right', 1, 0, 1, 1]]) {
          const bw = parseFloat(cs[`border${side}Width`]); const st = cs[`border${side}Style`]; const bc = parseColor(cs[`border${side}Color`]);
          if (!bw || st === 'none' || st === 'hidden' || !bc) continue;
          const half = bw / 2;
          const ax = box.x + x1 * box.w + (side === 'Left' ? half : side === 'Right' ? -half : 0);
          const ay = box.y + y1 * box.h + (side === 'Top' ? half : side === 'Bottom' ? -half : 0);
          const bx = box.x + x2 * box.w + (side === 'Left' ? half : side === 'Right' ? -half : 0);
          const by = box.y + y2 * box.h + (side === 'Top' ? half : side === 'Bottom' ? -half : 0);
          const c = inter({ x: Math.min(ax, bx) - half, y: Math.min(ay, by) - half, w: Math.abs(bx - ax) + bw, h: Math.abs(by - ay) + bw }, clip);
          if (c.w <= 0 || c.h <= 0) continue;
          if (st === 'double' && bw >= 2.5) {
            const t = bw / 3;
            if (side === 'Top' || side === 'Bottom') { pg.rect(c.x, c.y, c.w, t, bc.hex, bc.alpha); pg.rect(c.x, c.y + c.h - t, c.w, t, bc.hex, bc.alpha); } else { pg.rect(c.x, c.y, t, c.h, bc.hex, bc.alpha); pg.rect(c.x + c.w - t, c.y, t, c.h, bc.hex, bc.alpha); }
          } else if (st === 'dashed' || st === 'dotted') {
            const hz = side === 'Top' || side === 'Bottom';
            pg.line(hz ? c.x : c.x + c.w / 2, hz ? c.y + c.h / 2 : c.y, hz ? c.x + c.w : c.x + c.w / 2, hz ? c.y + c.h / 2 : c.y + c.h, bw, bc.hex, st === 'dotted' ? [bw, bw * 1.5] : [bw * 3, bw * 2]);
          } else pg.rect(c.x, c.y, c.w, c.h, bc.hex, bc.alpha);
        }
      }
    }
    // 잘라내기 (넘치는 부분 숨김)
    let inner = clip;
    if (cs.overflow !== 'visible' || cs.overflowX !== 'visible' || cs.overflowY !== 'visible') inner = inter(box, clip);
    if (inner.w <= 0 || inner.h <= 0) return;
    // 제목 → 책갈피
    const pid = el.dataset?.pid;
    if (pid && levels.has(pid) && !seenHead.has(pid) && !outline.done.has(pid)) {
      seenHead.add(pid); outline.done.add(pid);
      const t = levels.get(pid);
      if (t.text.trim()) w.addOutline(t.text.trim().slice(0, 120), pageIndex, box.y, t.level);
    }
    // 그림 · SVG · 수식
    if (el instanceof HTMLImageElement) {
      if (el.src) { const id = await imageId(w, cache, el.src, el); if (id != null) pg.image(id, box.x, box.y, box.w, box.h, inner.w < box.w - 0.5 || inner.h < box.h - 0.5 ? inner : null); }
      return;
    }
    if (el instanceof SVGSVGElement || el.localName === 'math' || el.classList?.contains('eq')) {
      const url = await rasterize(el);
      if (url) { const id = await imageId(w, cache, url, null); if (id != null) pg.image(id, box.x, box.y, box.w, box.h, inner); }
      return;
    }
    if (el instanceof HTMLCanvasElement) {
      try { const id = await imageId(w, cache, el.toDataURL('image/png'), null); if (id != null) pg.image(id, box.x, box.y, box.w, box.h, inner); } catch { /* 무시 */ }
      return;
    }
    // 연결
    if (el instanceof HTMLAnchorElement && /^(https?|mailto):/i.test(el.getAttribute('href') ?? '')) for (const r of el.getClientRects()) { const q = inter(R(r), inner); if (q.w > 0) pg.link(q.x, q.y, q.w, q.h, el.href); }
    for (const ch of el.childNodes) {
      if (ch.nodeType === 1) await walk(ch, inner);
      else if (ch.nodeType === 3 && ch.data.trim()) await textNode(ch, cs, inner);
    }
  };
  const textNode = async (tn, cs, clip) => {
    const size = parseFloat(cs.fontSize) || 13;
    const color = parseColor(cs.color) ?? { hex: '#000000', alpha: 1 };
    const style = { color: color.hex, alpha: color.alpha, serif: SERIF.test(cs.fontFamily.split(',')[0]), bold: Number(cs.fontWeight) >= 600 || cs.fontWeight === 'bold', italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique' };
    // 글꼴: 문단의 글꼴 목록 → 대체 글꼴 (넣을 수 있는 TrueType 만), 마지막은 미리 정의된 한글 글꼴
    const chain = [];
    if (opts.embed !== false) {
      for (const fam of [...cs.fontFamily.split(',').map((x) => x.trim()), ...FALLBACK]) {
        if (/^(serif|sans-serif|monospace|system-ui|cursive|fantasy)$/i.test(fam)) continue;
        const r = await fontRef(w, refs, fam, style.bold, style.italic, opts.ask);
        if (r && !chain.includes(r)) chain.push(r);
      }
    }
    // 라틴 글자: 화면에서 잰 글자 간격과 너비가 가장 잘 맞는 글꼴 (브라우저가 실제로 쓴 글꼴과 같게)
    let latin = null;
    const pick = (cp) => { if (cp < 0x2000 && latin) return [latin, latin.font.gidOf(cp)]; for (const r of chain) { const g = r.font.gidOf(cp); if (g) return [r, g]; } return [null, 0]; };
    const chooseLatin = (chars, xs) => {
      latin = null;
      const idx = chars.map((c, i) => i).filter((i) => i < chars.length - 1 && chars[i].codePointAt(0) < 0x2000 && chars[i] !== ' ');
      if (!idx.length) return;
      let best = Infinity;
      for (const r of new Set(chain)) {
        if (!idx.every((i) => r.font.gidOf(chars[i].codePointAt(0)))) continue;
        let err = 0;
        for (const i of idx) err += Math.abs(r.font.width(r.font.gidOf(chars[i].codePointAt(0))) / 1000 * size - (xs[i + 1] - xs[i]));
        if (err < best) { best = err; latin = r; }
      }
    };
    // 밑줄 · 취소선: 이 글자를 감싼 요소들의 장식
    let deco = ''; for (let p = tn.parentElement; p && !p.classList.contains('p'); p = p.parentElement) deco += ` ${getComputedStyle(p).textDecorationLine}`;
    const s = tn.data;
    let run = null;
    const flush = () => {
      if (!run) return;
      while (run.chars.length && /\s/.test(run.chars[run.chars.length - 1])) { run.chars.pop(); run.xs.pop(); }
      if (run.chars.length) {
        // 글꼴이 바뀌는 곳에서 나눠 그림
        chooseLatin(run.chars, run.xs);
        let a = 0;
        while (a < run.chars.length) {
          const [ref0] = pick(run.chars[a].codePointAt(0));
          let z = a + 1;
          while (z < run.chars.length && (run.chars[z] === ' ' || pick(run.chars[z].codePointAt(0))[0] === ref0)) z++;
          const ch = run.chars.slice(a, z); const xs = run.xs.slice(a, z);
          if (ref0) pg.glyphs(ref0, ch.map((c) => ref0.font.gidOf(c.codePointAt(0)) || ref0.font.gidOf(32)), ch, xs, run.base, size, { ...style, bold: style.bold && ref0.fake.bold, italic: style.italic && ref0.fake.italic });
          else pg.text(ch, xs, run.base, size, style);
          a = z;
        }
        const x1 = run.right;
        if (/underline/.test(deco)) pg.rect(run.xs[0], run.base + size * 0.12, x1 - run.xs[0], Math.max(0.6, size * 0.06), color.hex);
        if (/line-through/.test(deco)) pg.rect(run.xs[0], run.base - size * 0.3, x1 - run.xs[0], Math.max(0.6, size * 0.06), color.hex);
      }
      run = null;
    };
    for (let i = 0; i < s.length;) {
      const cp = s.codePointAt(i); const len = cp > 0xffff ? 2 : 1;
      range.setStart(tn, i); range.setEnd(tn, i + len);
      const rects = range.getClientRects();
      const r = rects[rects.length - 1];
      i += len;
      if (!r || (!r.width && cp !== 32)) continue;
      const q = R(r);
      const cx = q.x + q.w / 2; const cy = q.y + q.h / 2;
      if (cx < clip.x || cx > clip.x + clip.w || cy < clip.y || cy > clip.y + clip.h) { flush(); continue; }
      const ch = String.fromCodePoint(cp);
      if (cp === 0x200b || cp === 0xad) continue;
      const base = q.y + q.h - q.h * 0.2;
      if (run && (Math.abs(base - run.base) > size * 0.3 || q.x < run.right - size * 0.6)) flush();
      if (!run) { if (/\s/.test(ch)) continue; run = { chars: [], xs: [], base, right: q.x }; }
      run.chars.push(/\s/.test(ch) ? ' ' : ch); run.xs.push(q.x); run.right = q.x + q.w;
    }
    flush();
  };
  await walk(node, pageClip);
}

/**
 * 현재 문서 → PDF 바이트
 * opts: { from, to, onProgress(i, n) }
 */
export async function exportPdfBytes(opts = {}) {
  // opts.doc: 화면에 열린 문서가 아닌 문서 (AI 연결에서 파일을 PDF 로) — 따로 쪽을 나눔
  const doc = opts.doc ?? S.doc;
  let measure = null;
  let res;
  if (opts.doc) {
    measure = document.createElement('div');
    measure.style.cssText = 'position:fixed;left:-60000px;top:0;visibility:hidden;';
    document.body.append(measure);
    const L = new Layout(measure);
    res = L.run(doc, {});
    res.layout = L;
  } else res = S.layout ?? relayout();
  const pages = res.pages.slice(Math.max(0, (opts.from ?? 1) - 1), opts.to ?? res.pages.length);
  const host = document.createElement('div');
  host.style.cssText = 'position:fixed;left:-30000px;top:0;pointer-events:none;';
  document.body.append(host);
  const w = new PdfWriter();
  w.info = { title: doc.props?.title || opts.name || S.docName, author: doc.props?.author, subject: doc.props?.subject };
  // 제목 문단 → 수준
  const levels = new Map();
  walkParas(doc.body, (p) => { const lv = headingLevel(doc, p); if (lv) levels.set(p.id, { level: lv - 1, text: paraPlain(p) }); });
  const outline = { done: new Set() };
  const cache = new Map();
  const refs = new Map();
  try {
    if (document.fonts?.ready) await document.fonts.ready;
    for (let i = 0; i < pages.length; i++) {
      const p = pages[i];
      const node = pageNode(doc, p, res, { marks: false, print: true });
      node.style.transform = ''; node.style.margin = '0'; node.style.boxShadow = 'none';
      host.replaceChildren(node);
      const imgs = [...node.querySelectorAll('img')];
      await Promise.all(imgs.map((im) => (im.complete ? null : new Promise((r) => { im.onload = r; im.onerror = r; }))));
      const pg = w.addPage(p.w, p.h);
      await capturePage(w, pg, node, cache, outline, i, levels, refs, opts);
      opts.onProgress?.(i + 1, pages.length);
      if (i % 3 === 2) await new Promise((r) => setTimeout(r, 0));
    }
  } finally { host.remove(); measure?.remove(); }
  return w.build();
}
