// 마크다운 ⇄ 문서 (DOM 없음 · MCP 서버에서도 씀)
// 읽기: CommonMark + GFM 의 흔한 부분 (제목 · 문단 · 목록(중첩 · 번호 · 할 일) · 인용 · 코드 블록 · 가로줄 · 표 · 그림 · 링크 · 굵게/기울임/취소선/코드/밑줄 · 각주 [^1])
// 쓰기: 제목 · 목록 · 표 · 인용 · 코드 · 그림(images: 'ref' → media:<id>, 'inline' → data URL, 'none') · 각주
import { newDoc, para, table, uid, paraPlain, headingLevel, computeNumbering, bulletList, numberList, normalizeRuns, effPara, cm } from './model.js';

const CODE_RPR = { font: 'Consolas', fontEa: 'D2Coding', sz: 9.5 };
function ensureMdStyles(doc) {
  doc.styles = { ...doc.styles };
  if (!doc.styles.CodeBlock) doc.styles.CodeBlock = { id: 'CodeBlock', name: '코드 블록', type: 'para', basedOn: 'Normal', next: 'Normal', order: 46, pPr: { spc: { before: 0, after: 0, line: 1, rule: 'auto' }, shd: '#F6F8FA', align: 'left', contextual: true }, rPr: { ...CODE_RPR } };
  if (!doc.styles.HTMLCode) doc.styles.HTMLCode = { id: 'HTMLCode', name: 'HTML 코드', type: 'char', order: 56, rPr: { ...CODE_RPR, shd: '#F3F3F3' } };
}

// ───────────── 인라인 ─────────────
const ESC = /\\([\\`*_{}[\]()#+\-.!|~<>])/g;
/** 인라인 마크다운 → 런 목록 */
export function inlineRuns(src, base = {}, ctx = {}) {
  const out = [];
  let i = 0;
  let buf = '';
  const flush = () => { if (buf) { out.push({ ...base, text: buf.replace(ESC, '$1') }); buf = ''; } };
  const find = (tok, from) => { let k = from; while ((k = src.indexOf(tok, k)) >= 0) { if (src[k - 1] !== '\\') return k; k += tok.length; } return -1; };
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    if (ch === '\\' && i + 1 < src.length && /[\\`*_{}[\]()#+\-.!|~<>]/.test(src[i + 1])) { buf += src[i + 1]; i += 2; continue; }
    if (ch === '`') {
      const m = /^(`+)([\s\S]*?[^`])\1(?!`)/.exec(rest);
      if (m) { flush(); out.push({ ...base, ...CODE_RPR, cstyle: 'HTMLCode', text: m[2].replace(/^ (.*) $/, '$1') }); i += m[0].length; continue; }
    }
    if (ch === '!' && src[i + 1] === '[') {
      const m = /^!\[([^\]]*)\]\(\s*<?([^)\s>]+)>?(?:\s+"([^"]*)")?\s*\)/.exec(rest);
      if (m) {
        flush();
        const r = { obj: 'img', id: uid('i'), alt: m[1], w: 320, h: 200 };
        const mm = /^media:(.+)$/.exec(m[2]);
        if (mm && ctx.media?.[mm[1]]) r.media = mm[1];
        else if (m[2].startsWith('data:') && ctx.media) { const id = uid('m'); ctx.media[id] = m[2]; r.media = id; } else r.src = m[2];
        const sz = ctx.sizes?.[r.media ?? r.src];
        if (sz) { r.w = sz.w; r.h = sz.h; }
        if (m[3]) r.title = m[3];
        out.push(r); i += m[0].length; continue;
      }
    }
    if (ch === '[' && src[i + 1] === '^') {
      const m = /^\[\^([^\]]+)\]/.exec(rest);
      if (m && ctx.notes) { flush(); const id = ctx.noteIds[m[1]] ?? (ctx.noteIds[m[1]] = uid('n')); out.push({ obj: 'note', id, kind: 'foot' }); i += m[0].length; continue; }
    }
    if (ch === '[') {
      // [글](주소 "제목")  — 괄호 짝 맞춤
      let depth = 0; let k = i;
      for (; k < src.length; k++) { if (src[k] === '\\') { k++; continue; } if (src[k] === '[') depth++; else if (src[k] === ']' && --depth === 0) break; }
      if (k < src.length && src[k + 1] === '(') {
        const m = /^\(\s*<?([^)\s>]*)>?(?:\s+"([^"]*)")?\s*\)/.exec(src.slice(k + 1));
        if (m) {
          flush();
          const url = m[1];
          const link = url.startsWith('#') ? { anchor: url.slice(1) } : { url };
          for (const r of inlineRuns(src.slice(i + 1, k), { ...base, link, cstyle: 'Hyperlink', color: '#0563C1', u: true }, ctx)) out.push(r);
          i = k + 1 + m[0].length; continue;
        }
      }
    }
    if (ch === '<') {
      const auto = /^<((?:https?|mailto):[^>\s]+)>/.exec(rest);
      if (auto) { flush(); out.push({ ...base, text: auto[1], link: { url: auto[1] }, cstyle: 'Hyperlink', color: '#0563C1', u: true }); i += auto[0].length; continue; }
      const tag = /^<(u|ins|sub|sup|mark|kbd|br)\s*\/?>/i.exec(rest);
      if (tag) {
        const t = tag[1].toLowerCase();
        if (t === 'br') { flush(); out.push({ obj: 'br' }); i += tag[0].length; continue; }
        const close = rest.toLowerCase().indexOf(`</${t}>`);
        if (close > 0) {
          flush();
          const add = t === 'u' || t === 'ins' ? { u: true } : t === 'sub' ? { va: 'sub' } : t === 'sup' ? { va: 'super' } : t === 'mark' ? { hl: 'yellow' } : { ...CODE_RPR };
          for (const r of inlineRuns(rest.slice(tag[0].length, close), { ...base, ...add }, ctx)) out.push(r);
          i += close + t.length + 3; continue;
        }
      }
    }
    if (ch === '~' && src[i + 1] === '~') {
      const e = find('~~', i + 2);
      if (e > i + 2) { flush(); for (const r of inlineRuns(src.slice(i + 2, e), { ...base, s: true }, ctx)) out.push(r); i = e + 2; continue; }
    }
    if (ch === '=' && src[i + 1] === '=') {
      const e = find('==', i + 2);
      if (e > i + 2) { flush(); for (const r of inlineRuns(src.slice(i + 2, e), { ...base, hl: 'yellow' }, ctx)) out.push(r); i = e + 2; continue; }
    }
    if (ch === '*' || ch === '_') {
      const n = src[i + 1] === ch ? (src[i + 2] === ch ? 3 : 2) : 1;
      const tok = ch.repeat(n);
      // _ 는 낱말 안에서는 강조가 아님 (snake_case)
      const leftOk = ch === '*' || !/[\p{L}\p{N}]/u.test(src[i - 1] ?? '');
      const nextCh = src[i + n];
      if (leftOk && nextCh && !/\s/.test(nextCh)) {
        let e = find(tok, i + n);
        while (e > 0 && (/\s/.test(src[e - 1]) || (ch === '_' && /[\p{L}\p{N}]/u.test(src[e + n] ?? '')))) e = find(tok, e + 1);
        if (e > i + n - 1 + 0 && e > i) {
          flush();
          const add = n === 3 ? { b: true, i: true } : n === 2 ? { b: true } : { i: true };
          for (const r of inlineRuns(src.slice(i + n, e), { ...base, ...add }, ctx)) out.push(r);
          i = e + n; continue;
        }
      }
    }
    if (ch === '\n') { flush(); out.push({ obj: 'br' }); i++; continue; }
    buf += ch; i++;
  }
  flush();
  return out;
}

// ───────────── 블록 ─────────────
const reHead = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$/;
const reHr = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/;
const reFence = /^ {0,3}(`{3,}|~{3,})\s*([^`\s]*)?.*$/;
const reList = /^( *)([-*+]|\d{1,9}[.)])[ \t]+(.*)$/;
const reQuote = /^ {0,3}> ?(.*)$/;
const reTableSep = /^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?\s*$/;
const splitRow = (l) => { let s = l.trim(); if (s.startsWith('|')) s = s.slice(1); if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1); const out = []; let cur = ''; for (let i = 0; i < s.length; i++) { if (s[i] === '\\' && s[i + 1] === '|') { cur += '|'; i++; } else if (s[i] === '|') { out.push(cur.trim()); cur = ''; } else cur += s[i]; } out.push(cur.trim()); return out; };

/** 마크다운 → 문서 */
export function mdToDoc(text, { doc = null, bodyWidth = null, sizes = null } = {}) {
  const d = doc ?? newDoc();
  ensureMdStyles(d);
  d.media = { ...(d.media ?? {}) };
  d.lists = { ...(d.lists ?? {}) };
  const ctx = { media: d.media, notes: {}, noteIds: {}, sizes };
  const lines = text.replace(/\r\n?/g, '\n').replace(/\t/g, '    ').split('\n');
  // 각주 정의 먼저 모음
  const noteDefs = {};
  for (let i = 0; i < lines.length; i++) {
    const m = /^\[\^([^\]]+)\]:\s?(.*)$/.exec(lines[i]);
    if (m) { let t = m[2]; while (i + 1 < lines.length && /^ {2,}\S/.test(lines[i + 1])) t += `\n${lines[++i].trim()}`; noteDefs[m[1]] = t; lines[i] = '\u0000'; if (m.index === 0) lines[i - (t.split('\n').length - 1)] = '\u0000'; }
  }
  const bullet = 'mdBullet';
  const ordered = 'mdNumber';
  const W = bodyWidth ?? (d.section.page.w - d.section.margin.l - d.section.margin.r);
  const blocks = parseBlocks(lines.filter((l) => l !== '\u0000'), 0, ctx, d, { bullet, ordered, W });
  d.body = blocks.length ? blocks : [para('')];
  if (d.body.some((b) => b.pPr?.num?.id === bullet)) d.lists[bullet] = bulletList('●');
  if (d.body.some((b) => b.pPr?.num?.id === ordered) || Object.keys(d.lists).some((k) => k.startsWith(`${ordered}_`))) d.lists[ordered] = numberList();
  for (const k of Object.keys(ctx.listRestarts ?? {})) d.lists[k] = { ...numberList(), override: { 0: { start: ctx.listRestarts[k] } } };
  // 각주
  const foot = { ...(d.notes?.foot ?? {}) };
  for (const [k, id] of Object.entries(ctx.noteIds)) foot[id] = [{ ...para(''), style: 'FootnoteText', runs: inlineRuns(noteDefs[k] ?? '', {}, ctx) }];
  d.notes = { ...(d.notes ?? { end: {} }), foot, end: d.notes?.end ?? {} };
  // 첫 줄 제목 → 문서 제목
  const h1 = d.body.find((b) => b.type === 'p' && b.style === 'Heading1');
  if (h1 && !d.props.title) d.props = { ...d.props, title: paraPlain(h1) };
  return d;
}

function parseBlocks(lines, depth, ctx, d, L) {
  const out = [];
  let i = 0;
  let paraBuf = [];
  let listSeq = 0;
  const flushPara = () => {
    if (!paraBuf.length) return;
    const txt = paraBuf.map((l, k) => (k < paraBuf.length - 1 && / {2,}$/.test(l) ? `${l.trimEnd()}\n` : k < paraBuf.length - 1 && l.endsWith('\\') ? `${l.slice(0, -1)}\n` : `${l.trim()}${k < paraBuf.length - 1 ? ' ' : ''}`)).join('');
    out.push({ ...para(''), runs: normalizeRuns(inlineRuns(txt.trim(), {}, ctx)) });
    paraBuf = [];
  };
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { flushPara(); i++; continue; }
    let m;
    if ((m = reFence.exec(line))) {
      flushPara();
      const fence = m[1];
      const lang = m[2] ?? '';
      const code = [];
      i++;
      while (i < lines.length && !lines[i].trimStart().startsWith(fence)) code.push(lines[i++]);
      i++;
      if (lang === 'math' || lang === 'latex') { out.push({ ...para(''), pPr: { align: 'center' }, runs: [{ obj: 'eq', latex: code.join('\n') }] }); continue; }
      for (const c of code.length ? code : ['']) out.push({ ...para(''), style: 'CodeBlock', runs: c ? [{ text: c }] : [], ...(lang ? { lang } : {}) });
      continue;
    }
    if (/^\s*\$\$/.test(line)) {
      flushPara();
      const one = /^\s*\$\$(.+)\$\$\s*$/.exec(line);
      let tex = '';
      if (one) { tex = one[1]; i++; } else { i++; const buf = []; while (i < lines.length && !/^\s*\$\$/.test(lines[i])) buf.push(lines[i++]); i++; tex = buf.join('\n'); }
      out.push({ ...para(''), pPr: { align: 'center' }, runs: [{ obj: 'eq', latex: tex.trim() }] });
      continue;
    }
    if ((m = reHead.exec(line)) && !/^ {4}/.test(line)) {
      flushPara();
      const lv = m[1].length;
      out.push({ ...para(''), style: `Heading${lv}`, runs: normalizeRuns(inlineRuns((m[2] ?? '').trim(), {}, ctx)) });
      i++; continue;
    }
    // Setext 제목 (=== / ---)
    if (paraBuf.length && /^ {0,3}(=+|-+)\s*$/.test(line)) {
      const lv = line.trim()[0] === '=' ? 1 : 2;
      const txt = paraBuf.join(' ').trim();
      paraBuf = [];
      out.push({ ...para(''), style: `Heading${lv}`, runs: normalizeRuns(inlineRuns(txt, {}, ctx)) });
      i++; continue;
    }
    if (reHr.test(line)) {
      flushPara();
      out.push({ ...para(''), pPr: { border: { bottom: { style: 'single', w: 0.75, color: '#A0A0A0', space: 1 } }, spc: { after: 8 } } });
      i++; continue;
    }
    if (reQuote.test(line)) {
      flushPara();
      const q = [];
      while (i < lines.length && (reQuote.test(lines[i]) || (lines[i].trim() && !reList.test(lines[i]) && q.length && !reHead.test(lines[i])))) { const qm = reQuote.exec(lines[i]); q.push(qm ? qm[1] : lines[i]); i++; }
      // > [!NOTE] 경고 상자 (GitHub)
      const alert = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$/i.exec(q[0] ?? '');
      const inner = parseBlocks(alert ? q.slice(1) : q, depth + 1, ctx, d, L);
      const ALERT = { NOTE: ['참고', '#0969DA'], TIP: ['팁', '#1A7F37'], IMPORTANT: ['중요', '#8250DF'], WARNING: ['경고', '#9A6700'], CAUTION: ['주의', '#CF222E'] };
      if (alert) {
        const [label, color] = ALERT[alert[1].toUpperCase()];
        inner.unshift({ ...para(''), runs: [{ text: `ⓘ ${label}`, b: true, color }] });
        for (const b of inner) if (b.type === 'p') b.pPr = { ...b.pPr, border: { left: { style: 'single', w: 3, color, space: 8 } }, ind: { l: 16 } };
      } else for (const b of inner) if (b.type === 'p') { if (b.style === 'Normal') b.style = 'Quote'; b.pPr = { ...b.pPr, align: 'left', ind: { l: 28.8, r: 28.8 } }; }
      out.push(...inner);
      continue;
    }
    // 표 (GFM)
    if (line.includes('|') && i + 1 < lines.length && reTableSep.test(lines[i + 1]) && lines[i + 1].includes('-')) {
      flushPara();
      const head = splitRow(line);
      const aligns = splitRow(lines[i + 1]).map((s) => (s.startsWith(':') && s.endsWith(':') ? 'center' : s.endsWith(':') ? 'right' : s.startsWith(':') ? 'left' : null));
      i += 2;
      const rows = [head];
      while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(splitRow(lines[i++]));
      const cols = head.length;
      const t = table(rows.length, cols, { width: L.W, style: 'TableGrid' });
      t.rows = rows.map((r, ri) => ({ ...(ri === 0 ? { header: true } : {}), cells: Array.from({ length: cols }, (_, ci) => ({ ...(ri === 0 ? { shd: '#F2F2F2' } : {}), blocks: [{ ...para(''), pPr: { spc: { after: 0, line: 1, rule: 'auto' }, ...(aligns[ci] ? { align: aligns[ci] } : {}) }, runs: normalizeRuns(inlineRuns(r[ci] ?? '', ri === 0 ? { b: true } : {}, ctx)) }] })) }));
      out.push(t, { ...para(''), pPr: { spc: { after: 0 } } });
      // 표 다음 빈 문단은 Word 처럼 하나만 (표가 문서 끝일 때만 필요) — 다음 블록이 있으면 지움
      continue;
    }
    if ((m = reList.exec(line))) {
      flushPara();
      // 목록 묶음: 같은 들여쓰기 이상 또는 이어지는 줄
      const baseIndent = m[1].length;
      const items = [];
      while (i < lines.length) {
        const lm = reList.exec(lines[i]);
        if (lm && lm[1].length <= baseIndent + 1) {
          if (items.length && /\d/.test(lm[2]) !== /\d/.test(items[0].marker)) break;
          items.push({ marker: lm[2], lines: [lm[3]], indent: lm[1].length, contentIndent: lm[1].length + lm[2].length + 1 });
          i++; continue;
        }
        if (!items.length) break;
        const l = lines[i];
        if (!l.trim()) { if (i + 1 < lines.length && (/^ {2,}/.test(lines[i + 1]) || (reList.exec(lines[i + 1])?.[1].length ?? -1) > baseIndent + 1)) { items[items.length - 1].lines.push(''); i++; continue; } break; }
        if (/^ {2,}/.test(l) || (!reHead.test(l) && !reHr.test(l) && !reQuote.test(l) && !reFence.test(l) && items[items.length - 1].lines[items[items.length - 1].lines.length - 1] !== '')) { items[items.length - 1].lines.push(l.replace(new RegExp(`^ {0,${items[items.length - 1].contentIndent}}`), '')); i++; continue; }
        break;
      }
      const isOrd = /\d/.test(items[0].marker);
      let numId = isOrd ? L.ordered : L.bullet;
      const start = isOrd ? parseInt(items[0].marker, 10) : 1;
      if (isOrd && (listSeq++ > 0 || start !== 1)) { numId = `${L.ordered}_${uid('l')}`; ctx.listRestarts = { ...(ctx.listRestarts ?? {}), [numId]: start }; }
      for (const it of items) {
        const first = [];
        let k = 0;
        while (k < it.lines.length && it.lines[k] !== '' && !reList.test(it.lines[k])) first.push(it.lines[k++]);
        let txt = first.join(' ');
        const task = /^\[( |x|X)\]\s+/.exec(txt);
        if (task) txt = `${task[1] === ' ' ? '☐' : '☑'} ${txt.slice(task[0].length)}`;
        const lvl = depth;
        out.push({ ...para(''), style: 'ListParagraph', pPr: { num: { id: task ? null : numId, lvl: Math.min(8, lvl) }, ...(task ? { ind: { l: 26.67 + lvl * 26.67 } } : {}) }, runs: normalizeRuns(inlineRuns(txt, {}, ctx)) });
        if (task) delete out[out.length - 1].pPr.num;
        const sub = it.lines.slice(k);
        if (sub.some((s) => s.trim())) {
          const inner = parseBlocks(sub, depth + 1, ctx, d, L);
          for (const b of inner) { if (b.type === 'p' && !b.pPr?.num && b.style === 'Normal') b.pPr = { ...b.pPr, ind: { l: 26.67 * (depth + 1) } }; out.push(b); }
        }
      }
      continue;
    }
    // HTML 블록 주석 · 이미지 한 줄 → 문단에 그대로
    paraBuf.push(line);
    i++;
  }
  flushPara();
  // 표 뒤에 붙인 빈 문단: 다음 블록이 있으면 없앰
  for (let k = out.length - 2; k >= 0; k--) if (out[k].type === 'tbl' && out[k + 1]?.type === 'p' && !out[k + 1].runs.length && k + 2 < out.length) out.splice(k + 1, 1);
  return out;
}

// ───────────── 쓰기 ─────────────
const mdEsc = (s) => s.replace(/([\\`*[\]<>|])/g, '\\$1').replace(/(^|[^\p{L}\p{N}])_|_(?=[^\p{L}\p{N}]|$)/gu, (m) => m.replace('_', '\\_')).replace(/^(\s*)([#>+-]|\d+\.)(\s)/, '$1\\$2$3');
function wrap(s, tok) { const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(s); return m[2] ? `${m[1]}${tok}${m[2]}${tok}${m[3]}` : s; }

export function runsToMd(doc, runs, opts = {}) {
  let s = '';
  for (const r of runs) {
    if (r.del && !opts.keepDeleted) continue;
    if (r.obj === 'tab') { s += '\t'; continue; }
    if (r.obj === 'br') { s += r.kind === 'page' ? '' : '<br>'; continue; }
    if (r.obj === 'img') {
      const src = opts.images === 'none' ? '' : r.media ? (opts.images === 'inline' ? doc.media?.[r.media] ?? '' : `media:${r.media}`) : r.src ?? '';
      if (src || opts.images !== 'none') s += `![${(r.alt ?? '').replace(/[[\]]/g, '')}](${src})`;
      continue;
    }
    if (r.obj === 'note') { s += `[^${opts.noteNo?.(r) ?? r.id}]`; continue; }
    if (r.obj === 'field') { s += mdEsc(r.result ?? ''); continue; }
    if (r.obj === 'sym') { s += r.char ?? ''; continue; }
    if (r.obj === 'eq') { s += `$${r.latex}$`; continue; }
    if (r.obj) continue;
    let t = r.text ?? '';
    if (!t) continue;
    const code = r.cstyle === 'HTMLCode' || /Consolas|Courier|D2Coding|monospace/i.test(r.font ?? '');
    if (code) t = t.includes('`') ? `\`\` ${t} \`\`` : `\`${t}\``;
    else {
      t = mdEsc(t);
      if (r.s || r.dstrike) t = wrap(t, '~~');
      if (r.b && r.i) t = wrap(t, '***'); else if (r.b) t = wrap(t, '**'); else if (r.i) t = wrap(t, '*');
      if (r.u && !r.link) t = `<u>${t}</u>`;
      if (r.va === 'super') t = `<sup>${t}</sup>`; else if (r.va === 'sub') t = `<sub>${t}</sub>`;
      if (r.hl) t = `==${t}==`;
    }
    if (r.link) t = `[${t}](${r.link.url ?? `#${r.link.anchor}`})`;
    s += t;
  }
  // 이어진 같은 링크 합치기: [a](u)[b](u) → [ab](u)
  return s.replace(/\[([^\]]*)\]\(([^)]+)\)\[([^\]]*)\]\(\2\)/g, '[$1$3]($2)');
}

/** 문서 → 마크다운 */
export function docToMd(doc, opts = {}) {
  const nums = computeNumbering(doc);
  const out = [];
  const noteOrder = [];
  const noteNo = (r) => { let k = noteOrder.indexOf(r.id); if (k < 0) { noteOrder.push(r.id); k = noteOrder.length - 1; } return k + 1; };
  const o = { ...opts, noteNo };
  let prevList = false;
  let inCode = false;
  const emitBlocks = (blocks, quoteDepth = 0) => {
    for (const b of blocks) {
      if (b.type === 'tbl') {
        if (inCode) { out.push('```'); inCode = false; }
        prevList = false;
        out.push('');
        const cells = b.rows.map((r, ri) => r.cells.map((c) => (c.vMerge === 'cont' ? '' : c.blocks.filter((x) => x.type === 'p').map((x) => runsToMd(doc, ri === 0 ? x.runs.map((y) => (y.b ? { ...y, b: false } : y)) : x.runs, o)).join('<br>').replace(/\n/g, '<br>'))));
        const n = Math.max(...cells.map((r) => r.reduce((s, _c, k) => s + (b.rows[cells.indexOf(r)].cells[k].span ?? 1), 0)), 1);
        const expand = (r, ri) => { const row = []; r.forEach((c, k) => { row.push(c); for (let x = 1; x < (b.rows[ri].cells[k].span ?? 1); x++) row.push(''); }); while (row.length < n) row.push(''); return row; };
        const rows = cells.map(expand);
        const al = (b.rows[1] ?? b.rows[0])?.cells.map((c) => effPara(doc, c.blocks[0] ?? { pPr: {} }).align) ?? [];
        out.push(`| ${rows[0].join(' | ')} |`);
        out.push(`|${rows[0].map((_, k) => (al[k] === 'center' ? ':---:' : al[k] === 'right' ? '---:' : '---')).join('|')}|`);
        for (const r of rows.slice(1)) out.push(`| ${r.join(' | ')} |`);
        out.push('');
        continue;
      }
      if (b.type !== 'p') continue;
      if (b.style === 'CodeBlock') {
        if (!inCode) { if (prevList) out.push(''); out.push(`\`\`\`${b.lang ?? ''}`); inCode = true; }
        out.push(b.runs.filter((r) => !r.del).map((r) => (r.obj === 'tab' ? '\t' : r.obj ? '' : r.text)).join(''));
        continue;
      }
      if (inCode) { out.push('```', ''); inCode = false; }
      const q = '> '.repeat(quoteDepth);
      const txt = runsToMd(doc, b.runs, o);
      const lv = headingLevel(doc, b);
      if (b.runs.length === 1 && b.runs[0].obj === 'eq') { out.push('', `$$${b.runs[0].latex}$$`, ''); prevList = false; continue; }
      if (lv && txt.trim()) { if (out.length && out[out.length - 1] !== '') out.push(''); out.push(`${'#'.repeat(Math.min(6, lv))} ${txt.trim()}`, ''); prevList = false; continue; }
      const nm = nums.get(b.id);
      if (nm) {
        const ind = '   '.repeat(nm.lvl);
        out.push(`${q}${ind}${nm.bullet ? '-' : `${String(nm.text).replace(/[^\d]/g, '') || '1'}.`} ${txt}`);
        prevList = true;
        continue;
      }
      const task = /^([☐☑])\s/.exec(txt);
      if (task && b.style === 'ListParagraph') { out.push(`${q}- [${task[1] === '☑' ? 'x' : ' '}] ${txt.slice(2)}`); prevList = true; continue; }
      if (prevList) { out.push(''); prevList = false; }
      if (b.pPr?.border?.bottom && !b.runs.length) { out.push('---', ''); continue; }
      if (b.style === 'Quote' || b.style === 'IntenseQuote') { out.push(`> ${txt}`, ''); continue; }
      if (!txt.trim()) { if (out.length && out[out.length - 1] !== '') out.push(''); continue; }
      out.push(`${q}${txt.replace(/\n/g, '  \n')}`, '');
    }
  };
  emitBlocks(doc.body);
  if (inCode) out.push('```');
  // 각주
  if (noteOrder.length) {
    out.push('');
    noteOrder.forEach((id, k) => { const bl = doc.notes?.foot?.[id] ?? doc.notes?.end?.[id] ?? []; out.push(`[^${k + 1}]: ${bl.filter((x) => x.type === 'p').map((x) => runsToMd(doc, x.runs, o)).join(' ')}`); });
  }
  return `${out.join('\n').replace(/\n{3,}/g, '\n\n').trim()}\n`;
}

export const MD_CODE_RPR = CODE_RPR;
export const mdCm = cm;
