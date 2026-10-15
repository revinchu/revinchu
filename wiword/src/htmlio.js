// HTML ↔ 블록 (DOM 없음: 관대한 자체 HTML 해석기). 클립보드 · .html 열기/저장 · 웹 붙여넣기 · Word 에서 복사한 HTML
import { para, uid, effPara, effRun, paraPlain, computeNumbering, fontName, HIGHLIGHTS, normalizeRuns, PX_PER_PT } from './model.js';
import { esc } from './render.js';

// ───────────── 관대한 HTML 해석기 ─────────────
const VOID = new Set(['br', 'img', 'hr', 'meta', 'link', 'input', 'col', 'area', 'base', 'wbr', 'source', 'o:p']);
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', hellip: '…', mdash: '—', ndash: '–', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', middot: '·', times: '×', divide: '÷', laquo: '«', raquo: '»', trade: '™', deg: '°' };
export const decodeEnt = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);?/gi, (m, e) => {
  if (e[0] === '#') { const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10); return Number.isFinite(n) && n > 0 && n < 0x110000 ? String.fromCodePoint(n) : m; }
  return ENT[e.toLowerCase()] ?? m;
});
export function parseHtml(src) {
  const root = { tag: '#root', attrs: {}, children: [] };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<!(?:doctype|DOCTYPE)[^>]*>|<\?[\s\S]*?\?>|<\/?([a-zA-Z][\w:-]*)((?:\s+[^\s=>/]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>|[^<]+|</g;
  let m;
  const top = () => stack[stack.length - 1];
  while ((m = re.exec(src))) {
    const s = m[0];
    if (s.startsWith('<!--') || s.startsWith('<!') || s.startsWith('<?')) continue;
    if (m[1]) {
      const tag = m[1].toLowerCase();
      if (s[1] === '/') {
        for (let i = stack.length - 1; i > 0; i--) if (stack[i].tag === tag) { stack.length = i; break; }
        continue;
      }
      const attrs = {};
      const ar = /([^\s=>/]+)(?:\s*=\s*("[^"]*"|'[^']*'|[^\s>]+))?/g;
      let a;
      while ((a = ar.exec(m[2] ?? ''))) { let v = a[2] ?? ''; if (v[0] === '"' || v[0] === "'") v = v.slice(1, -1); attrs[a[1].toLowerCase()] = decodeEnt(v); }
      // 암묵 닫기 (p · li · td · tr)
      const autoClose = { p: ['p'], li: ['li'], td: ['td', 'th'], th: ['td', 'th'], tr: ['tr', 'td', 'th'], dt: ['dt', 'dd'], dd: ['dt', 'dd'], option: ['option'] }[tag];
      if (autoClose) {
        for (let i = stack.length - 1; i > 0; i--) { if (autoClose.includes(stack[i].tag)) { stack.length = i; break; } if (['table', 'ul', 'ol', 'div', 'body', 'tbody', 'thead'].includes(stack[i].tag)) break; }
      }
      if (['div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'blockquote', 'pre'].includes(tag) && top().tag === 'p') stack.pop();
      const node = { tag, attrs, children: [] };
      top().children.push(node);
      if (!VOID.has(tag) && !m[3] && tag !== 'script' && tag !== 'style') stack.push(node);
      else if (tag === 'script' || tag === 'style') {
        // 내용 건너뛰기
        const end = src.toLowerCase().indexOf(`</${tag}`, re.lastIndex);
        node.text = end < 0 ? '' : src.slice(re.lastIndex, end);
        re.lastIndex = end < 0 ? src.length : src.indexOf('>', end) + 1;
      }
    } else top().children.push({ tag: '#text', text: decodeEnt(s) });
  }
  return root;
}
export const textOf = (n) => (n.tag === '#text' ? n.text : (n.children ?? []).map(textOf).join(''));
function findTag(n, tag) { if (n.tag === tag) return n; for (const c of n.children ?? []) { const r = findTag(c, tag); if (r) return r; } return null; }

// ───────────── CSS 한 줄 → 객체 ─────────────
export function parseStyle(s) {
  const out = {};
  for (const part of String(s ?? '').split(';')) { const i = part.indexOf(':'); if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim(); }
  return out;
}
const NAMED = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', silver: '#C0C0C0', maroon: '#800000', navy: '#000080', purple: '#800080', teal: '#008080', olive: '#808000', orange: '#FFA500', lime: '#00FF00', aqua: '#00FFFF', fuchsia: '#FF00FF', windowtext: '#000000' };
export function cssColor(v) {
  if (!v) return null;
  v = v.trim().toLowerCase();
  if (v === 'transparent' || v === 'inherit' || v === 'initial' || v === 'auto' || v === 'none') return null;
  if (NAMED[v]) return NAMED[v];
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return `#${m[1].split('').map((c) => c + c).join('')}`.toUpperCase();
  m = /^#([0-9a-f]{6})/.exec(v);
  if (m) return `#${m[1]}`.toUpperCase();
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?/.exec(v);
  if (m) { if (m[4] && parseFloat(m[4]) === 0) return null; return `#${[m[1], m[2], m[3]].map((x) => Number(x).toString(16).padStart(2, '0')).join('')}`.toUpperCase(); }
  return null;
}
export function cssLen(v, fontPx = 13.33) {
  if (v == null) return null;
  const m = /^(-?[\d.]+)(px|pt|cm|mm|in|em|rem|%)?$/.exec(String(v).trim());
  if (!m) return null;
  const n = parseFloat(m[1]);
  switch (m[2]) { case 'pt': return n * PX_PER_PT; case 'cm': return (n * 96) / 2.54; case 'mm': return (n * 96) / 25.4; case 'in': return n * 96; case 'em': case 'rem': return n * fontPx; case '%': return null; default: return n; }
}

// ───────────── HTML → 블록 ─────────────
const BLOCK = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'blockquote', 'pre', 'table', 'ul', 'ol', 'hr', 'section', 'article', 'header', 'footer', 'main', 'aside', 'nav', 'figure', 'figcaption', 'dl', 'dt', 'dd', 'address', 'center', 'body', 'html', 'tbody', 'thead', 'tfoot', 'tr', 'td', 'th', 'form', 'details', 'summary', 'caption']);
/** HTML 글 → { blocks, media } */
export function htmlToBlocks(html, doc = null, opts = {}) {
  const root = parseHtml(String(html).replace(/<!--StartFragment-->|<!--EndFragment-->/g, ''));
  const body = findTag(root, 'body') ?? root;
  const media = {};
  const lists = opts.lists ?? null; // 목록 정의를 만들 때 (열기)
  const out = [];
  const ctx = { media, lists, listStack: [], doc };
  walkBlocks(body.children, out, { style: 'Normal', pPr: {}, rPr: {} }, ctx);
  // 앞뒤 빈 문단 정리
  while (out.length > 1 && out[out.length - 1].type === 'p' && !out[out.length - 1].runs.length) out.pop();
  return { blocks: out.length ? out : [], media, lists: ctx.listsMade };
}
function inlineRpr(node, base) {
  const r = { ...base };
  const st = parseStyle(node.attrs?.style);
  const t = node.tag;
  if (t === 'b' || t === 'strong') r.b = true;
  if (t === 'i' || t === 'em' || t === 'cite' || t === 'var' || t === 'dfn') r.i = true;
  if (t === 'u' || t === 'ins') r.u = true;
  if (t === 's' || t === 'strike' || t === 'del') r.s = true;
  if (t === 'sup') r.va = 'super';
  if (t === 'sub') r.va = 'sub';
  if (t === 'code' || t === 'kbd' || t === 'samp' || t === 'tt') { r.font = 'Consolas'; r.fontEa = 'D2Coding'; r.shd = r.shd ?? '#F3F3F3'; }
  if (t === 'mark') r.hl = 'yellow';
  if (t === 'small') r.sz = (base.sz ?? 10) * 0.85;
  if (t === 'big') r.sz = (base.sz ?? 10) * 1.2;
  if (t === 'font') { if (node.attrs.color) r.color = cssColor(node.attrs.color); if (node.attrs.face) r.font = r.fontEa = node.attrs.face.split(',')[0].replace(/['"]/g, '').trim(); if (node.attrs.size) r.sz = [8, 10, 12, 14, 18, 24, 36][Math.max(0, Math.min(6, Number(node.attrs.size) - 1))] ?? 12; }
  if (st['font-weight'] && (st['font-weight'] === 'bold' || Number(st['font-weight']) >= 600)) r.b = true;
  if (st['font-weight'] === 'normal' || Number(st['font-weight']) < 600) { if (st['font-weight']) delete r.b; }
  if (st['font-style'] === 'italic') r.i = true; else if (st['font-style'] === 'normal') delete r.i;
  const deco = st['text-decoration'] ?? st['text-decoration-line'] ?? '';
  if (/underline/.test(deco)) r.u = true;
  if (/line-through/.test(deco)) r.s = true;
  if (/none/.test(deco) && t !== 'a') { delete r.u; delete r.s; }
  const col = cssColor(st.color);
  if (col) r.color = col;
  const bg = cssColor(st['background-color'] ?? st.background ?? st['mso-highlight']);
  if (bg && bg !== '#FFFFFF') r.shd = bg;
  if (st['font-size']) { const px = cssLen(st['font-size']); if (px) r.sz = Math.round((px / PX_PER_PT) * 2) / 2; else if (/smaller|small/.test(st['font-size'])) r.sz = (base.sz ?? 10) * 0.85; else if (/larger|large/.test(st['font-size'])) r.sz = (base.sz ?? 10) * 1.2; }
  if (st['font-family']) { const f = st['font-family'].split(',').map((x) => x.replace(/['"]/g, '').trim()).filter(Boolean); if (f[0] && !/^(sans-serif|serif|monospace|system-ui|-apple-system|inherit)$/i.test(f[0])) { r.font = f[0]; r.fontEa = f.find((x) => /[가-힣]|Gothic|Myeongjo|Batang|Dotum|Gulim|Malgun|Nanum|Noto Sans KR/i.test(x)) ?? f[0]; } }
  if (st['vertical-align'] === 'super') r.va = 'super';
  if (st['vertical-align'] === 'sub') r.va = 'sub';
  if (st['text-transform'] === 'uppercase') r.caps = true;
  if (st['font-variant'] === 'small-caps') r.smallCaps = true;
  if (st['letter-spacing']) { const px = cssLen(st['letter-spacing']); if (px) r.spc = px / PX_PER_PT; }
  if (t === 'a' && node.attrs.href) { const h = node.attrs.href; r.link = h.startsWith('#') ? { anchor: h.slice(1) } : { url: h }; }
  return r;
}
function paraProps(node, base) {
  const p = { ...base };
  const st = parseStyle(node.attrs?.style);
  const al = st['text-align'] ?? node.attrs?.align;
  if (al) p.align = { left: 'left', center: 'center', right: 'right', justify: 'justify', start: 'left', end: 'right' }[al.toLowerCase()] ?? p.align;
  const ind = { ...(p.ind ?? {}) };
  const ml = cssLen(st['margin-left'] ?? st['padding-left']);
  if (ml) ind.l = (ind.l ?? 0) + ml;
  const ti = cssLen(st['text-indent']);
  if (ti) { if (ti < 0) ind.hanging = -ti; else ind.first = ti; }
  if (Object.keys(ind).length) p.ind = ind;
  const spc = {};
  const mt = cssLen(st['margin-top']); const mb = cssLen(st['margin-bottom']);
  if (mt != null) spc.before = mt;
  if (mb != null) spc.after = mb;
  const lh = st['line-height'];
  if (lh) { const n = parseFloat(lh); if (/%$/.test(lh)) { spc.line = n / 100 / 1.15; spc.rule = 'auto'; } else if (/^[\d.]+$/.test(lh)) { spc.line = n / 1.15; spc.rule = 'auto'; } else { const px = cssLen(lh); if (px) { spc.line = px; spc.rule = 'atLeast'; } } }
  if (Object.keys(spc).length) p.spc = { ...(p.spc ?? {}), ...spc };
  const bg = cssColor(st['background-color'] ?? st.background);
  if (bg && bg !== '#FFFFFF') p.shd = bg;
  if (st['page-break-before'] === 'always' || st['break-before'] === 'page') p.pageBreakBefore = true;
  return p;
}
function walkBlocks(nodes, out, ctx0, ctx) {
  let cur = null; // 이어 쓰는 문단 (블록 밖 글)
  const flush = () => { if (cur) { cur.runs = normalizeRuns(cur.runs); if (cur.runs.length || cur.keep) out.push(cur); cur = null; } };
  const ensure = () => { if (!cur) cur = para('', { style: ctx0.style, pPr: { ...ctx0.pPr } }); return cur; };
  for (const n of nodes) {
    if (n.tag === '#text') {
      const t = ctx0.pre ? n.text : n.text.replace(/[\s ]+/g, (w) => (w.includes(' ') ? w.replace(/[^ ]/g, '').replace(/ /g, ' ') || ' ' : ' '));
      if (!t || (!cur && !t.trim())) continue;
      if (ctx0.pre) { const lines = t.split('\n'); lines.forEach((ln, i) => { if (i) { flush(); } if (ln) ensure().runs.push({ ...ctx0.rPr, text: ln }); else if (i) ensure().keep = true; }); continue; }
      ensure().runs.push({ ...ctx0.rPr, text: t });
      continue;
    }
    const t = n.tag;
    const st = parseStyle(n.attrs?.style);
    if (st.display === 'none' || st['mso-hide'] === 'all' || t === 'head' || t === 'title' || t === 'script' || t === 'style' || t === 'meta' || t === 'o:p' && !textOf(n).trim()) continue;
    if (/mso-list:\s*ignore/i.test(n.attrs?.style ?? '')) continue; // Word 목록 기호 (번호는 다시 매김)
    if (t === 'br') { if (ctx0.inline) { ensure().runs.push({ obj: 'br', ...ctx0.rPr }); } else { ensure().runs.push({ obj: 'br', ...ctx0.rPr }); } continue; }
    if (t === 'img') {
      const src = n.attrs.src ?? '';
      const w = cssLen(n.attrs.width) ?? cssLen(st.width) ?? 200;
      const h = cssLen(n.attrs.height) ?? cssLen(st.height) ?? 150;
      let mid = null;
      if (src.startsWith('data:')) { mid = uid('m'); ctx.media[mid] = src; }
      const run = { obj: 'img', id: uid('i'), media: mid, src: mid ? undefined : src, w, h, alt: n.attrs.alt ?? '' };
      ensure().runs.push(run);
      continue;
    }
    if (t === 'hr') { flush(); out.push(para('', { pPr: { border: { bottom: { style: 'single', w: 0.75, color: '#A0A0A0', space: 1 } } } })); continue; }
    if (t === 'table') { flush(); const tb = tableFrom(n, ctx0, ctx); if (tb) out.push(tb); continue; }
    if (t === 'ul' || t === 'ol') {
      flush();
      const ordered = t === 'ol';
      const lvl = ctx.listStack.length;
      const numId = listFor(ctx, ordered, n.attrs.type, lvl);
      ctx.listStack.push(numId);
      for (const li of n.children) {
        if (li.tag === '#text') continue;
        if (li.tag === 'li') {
          const sub = [];
          const lp = { ...paraProps(li, ctx0.pPr), num: { id: numId, lvl } };
          walkBlocks(li.children, sub, { ...ctx0, style: 'ListParagraph', pPr: lp, rPr: inlineRpr(li, ctx0.rPr) }, ctx);
          // li 안 첫 문단만 번호, 나머지는 같은 들여쓰기
          let first = true;
          for (const b of sub) { if (b.type === 'p' && b.pPr.num?.id === numId && b.pPr.num.lvl === lvl) { if (!first) { b.pPr = { ...b.pPr }; delete b.pPr.num; } first = false; } out.push(b); }
          if (!sub.length) out.push(para('', { style: 'ListParagraph', pPr: lp }));
          // 체크 상자 (- [ ] 작업 목록)
        } else if (li.tag === 'ul' || li.tag === 'ol') walkBlocks([li], out, ctx0, ctx);
      }
      ctx.listStack.pop();
      continue;
    }
    if (BLOCK.has(t)) {
      flush();
      let style = ctx0.style;
      let pPr = paraProps(n, ctx0.pPr);
      let rPr = inlineRpr(n, ctx0.rPr);
      const hm = /^h([1-6])$/.exec(t);
      if (hm) { style = `Heading${hm[1]}`; rPr = { ...rPr }; delete rPr.b; delete rPr.sz; delete rPr.color; }
      if (t === 'blockquote') style = 'Quote';
      if (t === 'pre') { rPr = { ...rPr, font: 'Consolas', fontEa: 'D2Coding' }; pPr = { ...pPr, spc: { before: 0, after: 0, line: 1, rule: 'auto' }, shd: '#F6F8FA' }; }
      if (t === 'figcaption' || t === 'caption') style = 'Caption';
      if (t === 'center') pPr.align = 'center';
      const cls = n.attrs?.class ?? '';
      const mh = /MsoHeading(\d)|Heading(\d)/i.exec(cls);
      if (mh) style = `Heading${mh[1] ?? mh[2]}`;
      if (/MsoTitle/i.test(cls)) style = 'Title';
      if (/MsoSubtitle/i.test(cls)) style = 'Subtitle';
      if (/MsoQuote/i.test(cls)) style = 'Quote';
      // Word 목록 문단 (mso-list:l0 level2)
      const ml = /mso-list:\s*l(\d+)\s+level(\d+)/i.exec(n.attrs?.style ?? '');
      if (ml) {
        const lvl = Number(ml[2]) - 1;
        const marker = textOf(n.children.find((c) => /mso-list:\s*ignore/i.test(c.attrs?.style ?? '')) ?? { tag: '#text', text: '' }).trim();
        const ordered = /^[\d가-힣a-zA-Z]+[.)]/.test(marker);
        const numId = listFor(ctx, ordered, null, lvl, `mso${ml[1]}`);
        pPr = { ...pPr, num: { id: numId, lvl } };
        delete pPr.ind;
        style = 'ListParagraph';
      }
      const blockStart = out.length;
      walkBlocks(n.children, out, { ...ctx0, style, pPr, rPr, pre: t === 'pre' || ctx0.pre }, ctx);
      if (out.length === blockStart && ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(t)) out.push(para('', { style, pPr }));
      continue;
    }
    // 글자 서식 요소
    const rPr = inlineRpr(n, ctx0.rPr);
    const sub = [];
    walkBlocks(n.children, sub, { ...ctx0, rPr, inline: true }, ctx);
    // 안쪽 결과: 첫 문단 런은 현재 문단에 이어 붙이고, 나머지는 블록으로
    sub.forEach((b, i) => {
      if (b.type === 'p' && i === 0 && b.style === ctx0.style) { ensure().runs.push(...b.runs); }
      else { flush(); out.push(b); }
    });
  }
  flush();
}
function listFor(ctx, ordered, type, lvl, key = null) {
  ctx.listsMade ??= {};
  ctx.listKeys ??= new Map();
  const k = key ?? (ordered ? `ol${type ?? ''}` : 'ul');
  if (lvl > 0 && ctx.listStack.length) return ctx.listStack[ctx.listStack.length - 1];
  if (key && ctx.listKeys.has(k)) return ctx.listKeys.get(k);
  const id = String(Object.keys(ctx.listsMade).length + 1 + (ctx.doc ? Object.keys(ctx.doc.lists ?? {}).length : 0) + 100);
  const fmt = type === 'a' ? 'lowerLetter' : type === 'A' ? 'upperLetter' : type === 'i' ? 'lowerRoman' : type === 'I' ? 'upperRoman' : 'decimal';
  ctx.listsMade[id] = ordered
    ? { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: i === 0 ? fmt : ['decimal', 'lowerLetter', 'lowerRoman'][i % 3], text: `%${i + 1}.`, start: 1, ind: { l: 26.67 + i * 26.67, hanging: 26.67 } })) }
    : { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: 'bullet', text: ['●', '○', '■'][i % 3], start: 1, ind: { l: 26.67 + i * 26.67, hanging: 26.67 } })) };
  ctx.listKeys.set(k, id);
  return id;
}
function tableFrom(n, ctx0, ctx) {
  const trs = [];
  const collect = (x) => { for (const c of x.children ?? []) { if (c.tag === 'tr') trs.push(c); else if (['thead', 'tbody', 'tfoot'].includes(c.tag)) collect(c); } };
  collect(n);
  if (!trs.length) return null;
  // 격자: colspan · rowspan
  const occupied = [];
  const rows = [];
  let maxCols = 0;
  trs.forEach((tr, r) => {
    occupied[r] ??= [];
    const cells = [];
    let c = 0;
    for (const td of tr.children.filter((x) => x.tag === 'td' || x.tag === 'th')) {
      while (occupied[r][c]) {
        // 세로 병합 이어짐 칸
        cells.push({ vMerge: 'cont', span: occupied[r][c].span, blocks: [para('')] });
        c += occupied[r][c].span;
      }
      const span = Math.max(1, Number(td.attrs.colspan) || 1);
      const rs = Math.max(1, Number(td.attrs.rowspan) || 1);
      const blocks = [];
      walkBlocks(td.children, blocks, { style: 'Normal', pPr: { spc: { before: 0, after: 0, line: 1, rule: 'auto' } }, rPr: td.tag === 'th' ? { b: true } : {} }, ctx);
      const st = parseStyle(td.attrs.style);
      const cell = { span, blocks: blocks.length ? blocks : [para('', { pPr: { spc: { before: 0, after: 0 } } })] };
      const bg = cssColor(st['background-color'] ?? st.background ?? td.attrs.bgcolor);
      if (bg) cell.shd = bg;
      if (st['vertical-align'] === 'middle' || td.attrs.valign === 'middle') cell.vAlign = 'center';
      if (rs > 1) { cell.vMerge = 'restart'; for (let k = 1; k < rs; k++) { occupied[r + k] ??= []; occupied[r + k][c] = { span }; } }
      cells.push(cell);
      c += span;
    }
    while (occupied[r][c]) { cells.push({ vMerge: 'cont', span: occupied[r][c].span, blocks: [para('')] }); c += occupied[r][c].span; }
    maxCols = Math.max(maxCols, c);
    rows.push({ cells, header: tr.children.every((x) => x.tag !== 'td') && r === 0 });
  });
  const W = 600;
  return { type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: { look: { firstRow: true, bandRows: true } }, grid: Array.from({ length: maxCols }, () => W / maxCols), rows };
}

// ───────────── 블록 → HTML (깨끗한 HTML: 클립보드 · .html 저장) ─────────────
export function blocksToHtml(doc, blocks, { full = false, title = '' } = {}) {
  const num = computeNumbering({ ...doc, body: blocks });
  const out = [];
  let listOpen = null; // { tag, lvl }
  const closeList = () => { if (listOpen) { out.push(`</${listOpen.tag}>`); listOpen = null; } };
  for (const b of blocks) {
    if (b.type === 'tbl') { closeList(); out.push(tableToHtml(doc, b)); continue; }
    const n = num.get(b.id);
    if (n && n.lvl === 0) {
      const tag = n.bullet ? 'ul' : 'ol';
      if (!listOpen || listOpen.tag !== tag) { closeList(); out.push(`<${tag}>`); listOpen = { tag }; }
      out.push(`<li>${runsHtml(doc, b)}</li>`);
      continue;
    }
    closeList();
    out.push(paraToHtml(doc, b, n));
  }
  closeList();
  const body = out.join('\n');
  if (!full) return body;
  return `<!doctype html>\n<html lang="ko">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<meta name="generator" content="WIWORD">\n<title>${esc(title || doc.props?.title || '문서')}</title>\n<style>body{font-family:'맑은 고딕','Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif;max-width:${Math.round(doc.section.page.w - doc.section.margin.l - doc.section.margin.r)}px;margin:40px auto;line-height:1.6;color:#222;padding:0 16px}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px 8px;vertical-align:top}img{max-width:100%}blockquote{border-left:4px solid #ccc;margin-left:0;padding-left:16px;color:#555}pre{background:#f6f8fa;padding:12px;overflow:auto}</style>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;
}
function paraToHtml(doc, p, n) {
  const ep = effPara(doc, p);
  const lv = ep.outline != null && ep.outline < 6 && /^Heading/.test(p.style ?? '') ? ep.outline + 1 : 0;
  const tag = lv ? `h${lv}` : p.style === 'Title' ? 'h1' : p.style === 'Quote' || p.style === 'IntenseQuote' ? 'blockquote' : 'p';
  const css = [];
  if (ep.align && ep.align !== 'left' && ep.align !== 'justify') css.push(`text-align:${ep.align}`);
  if (ep.ind?.l) css.push(`margin-left:${Math.round(ep.ind.l)}px`);
  if (ep.ind?.first) css.push(`text-indent:${Math.round(ep.ind.first)}px`);
  if (ep.shd) css.push(`background:${ep.shd}`);
  const lbl = n ? `${esc(n.text)} ` : '';
  const inner = runsHtml(doc, p) || '<br>';
  return `<${tag}${css.length ? ` style="${css.join(';')}"` : ''}>${lbl}${inner}</${tag}>`;
}
export function runsHtml(doc, p) {
  let s = '';
  for (const r of p.runs) {
    if (r.del) continue;
    if (r.obj === 'tab') { s += '&emsp;'; continue; }
    if (r.obj === 'br') { s += '<br>'; continue; }
    if (r.obj === 'img') { const src = doc.media?.[r.media] ?? r.src ?? ''; s += `<img src="${esc(src)}" width="${Math.round(r.w)}" height="${Math.round(r.h)}" alt="${esc(r.alt ?? '')}">`; continue; }
    if (r.obj === 'field') { s += esc(r.result ?? ''); continue; }
    if (r.obj === 'sym') { s += esc(r.char ?? ''); continue; }
    if (r.obj) continue;
    let t = esc(r.text).replace(/ {2}/g, ' &nbsp;');
    const css = [];
    const rp = effRun(doc, p, r);
    if (r.font) css.push(`font-family:'${fontName(doc, r.font)}'`);
    if (r.sz) css.push(`font-size:${r.sz}pt`);
    if (r.color && r.color !== 'auto') css.push(`color:${r.color}`);
    if (r.hl) css.push(`background:${HIGHLIGHTS[r.hl] ?? r.hl}`);
    else if (r.shd) css.push(`background:${r.shd}`);
    if (r.spc) css.push(`letter-spacing:${r.spc}pt`);
    if (css.length) t = `<span style="${css.join(';')}">${t}</span>`;
    if (rp.b && !/^Heading|Title/.test(p.style ?? '')) t = `<strong>${t}</strong>`;
    if (rp.i) t = `<em>${t}</em>`;
    if (rp.u && !r.link) t = `<u>${t}</u>`;
    if (rp.s) t = `<s>${t}</s>`;
    if (rp.va === 'super') t = `<sup>${t}</sup>`;
    if (rp.va === 'sub') t = `<sub>${t}</sub>`;
    if (r.link) t = `<a href="${esc(r.link.url ?? `#${r.link.anchor}`)}">${t}</a>`;
    s += t;
  }
  return s;
}
function tableToHtml(doc, t) {
  let h = '<table>';
  t.rows.forEach((row, ri) => {
    h += '<tr>';
    for (const c of row.cells) {
      if (c.vMerge === 'cont') continue;
      let rs = 1;
      if (c.vMerge === 'restart') { const ci = row.cells.indexOf(c); for (let k = ri + 1; k < t.rows.length && t.rows[k].cells[ci]?.vMerge === 'cont'; k++) rs++; }
      const tag = row.header ? 'th' : 'td';
      h += `<${tag}${c.span > 1 ? ` colspan="${c.span}"` : ''}${rs > 1 ? ` rowspan="${rs}"` : ''}${c.shd ? ` style="background:${c.shd}"` : ''}>${c.blocks.map((b) => (b.type === 'p' ? runsHtml(doc, b) : tableToHtml(doc, b))).join('<br>')}</${tag}>`;
    }
    h += '</tr>';
  });
  return `${h}</table>`;
}
