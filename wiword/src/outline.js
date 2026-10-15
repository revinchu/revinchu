// 마크다운 개요 → 새 문서 (DOM 없음) — 이음오피스의 "문서로 새 문서 만들기" 를 참고해 AI 없이 만드는 방식으로 직접 구현
// '# 제목' · '## 장' · '- 항목' 개요를 표지 · 차례 · 번호 매긴 제목 · 빈 절 안내 문구가 있는 문서로. 다른 문서의 제목들로 개요를 뽑을 수도 있음.
import { newDoc, para, uid, headings, buildToc, paraPlain, docText } from './model.js';
import { mdToDoc } from './md.js';

/** 문서 종류 (id, 이름, 기본값) */
export const OUTLINE_KINDS = [
  { id: 'report', name: '보고서', number: 'decimal', cover: true, toc: true, chapterBreak: false, desc: '표지 · 차례 · 1. / 1.1 번호' },
  { id: 'plan', name: '기획서', number: 'plan', cover: true, toc: true, chapterBreak: true, desc: 'Ⅰ. / 1. / 가. 번호, 장마다 새 쪽' },
  { id: 'gongmun', name: '공문서 · 보고 문서', number: 'gongmun', cover: false, toc: false, chapterBreak: false, desc: '1. / 가. / 1) / 가) / (1) / (가) 항목 번호' },
  { id: 'paper', name: '논문 · 과제', number: 'paper', cover: true, toc: true, chapterBreak: false, desc: '1 / 1.1 / 1.1.1 번호, 차례' },
  { id: 'minutes', name: '회의록', number: 'none', cover: false, toc: false, chapterBreak: false, minutes: true, desc: '일시 · 장소 · 참석자 표 + 안건' },
  { id: 'plain', name: '기본 문서', number: 'none', cover: false, toc: false, chapterBreak: false, desc: '개요 그대로' },
];
const HANGUL = '가나다라마바사아자차카타파하';
const ROMAN = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ', 'Ⅺ', 'Ⅻ'];
const han = (n) => HANGUL[(n - 1) % HANGUL.length] ?? String(n);
/** 번호 글: counters = 단계마다 번호 (1부터) */
export function numberLabel(scheme, counters) {
  const lv = counters.length; const n = counters[lv - 1];
  switch (scheme) {
    case 'decimal': return lv === 1 ? `${n}.` : counters.join('.');
    case 'paper': return counters.join('.');
    case 'plan': return lv === 1 ? `${ROMAN[n - 1] ?? n}.` : lv === 2 ? `${n}.` : lv === 3 ? `${han(n)}.` : lv === 4 ? `${n})` : `${han(n)})`;
    case 'gongmun': return [`${n}.`, `${han(n)}.`, `${n})`, `${han(n)})`, `(${n})`, `(${han(n)})`][lv - 1] ?? `${n}.`;
    default: return '';
  }
}
/** 제목 앞에 이미 붙은 번호 지우기 ("1. 개요" → "개요") */
const NUM_PREFIX = /^\s*(?:\d+(?:\.\d+)*\.?|[가-하]\.|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+\.|\d+\)|[가-하]\)|\(\d+\)|\([가-하]\)|제\s*\d+\s*[장절조])\s+/;
export const stripNumber = (t) => String(t).replace(NUM_PREFIX, '');

/** 개요 글 → [{ kind: 'h', level, text } | { kind: 'md', text }] — '#' 이 없으면 "1. / 1.1 / 가." 로 시작하는 짧은 줄을 제목으로 */
export function parseOutline(md) {
  const lines = String(md ?? '').replace(/\r\n?/g, '\n').split('\n');
  const hasHash = lines.some((l) => /^#{1,6}\s+\S/.test(l));
  const items = [];
  let buf = [];
  const flush = () => { const t = buf.join('\n').replace(/^\n+|\n+$/g, ''); if (t.trim()) items.push({ kind: 'md', text: t }); buf = []; };
  let inFence = false;
  for (const line of lines) {
    if (/^\s*(```|~~~)/.test(line)) inFence = !inFence;
    if (inFence) { buf.push(line); continue; }
    let m = /^(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line);
    let level = m ? m[1].length : 0; let text = m ? m[2] : '';
    if (!m && !hasHash && line.trim().length <= 60 && !/[.。!?]$/.test(line.trim().replace(/^\s*\d+\.\s*/, ''))) {
      if ((m = /^(\d+)\.\s+(\S.*)$/.exec(line))) { level = 1; text = m[2]; }
      else if ((m = /^\d+\.\d+\.?\s+(\S.*)$/.exec(line))) { level = 2; text = m[1]; }
      else if ((m = /^\d+\.\d+\.\d+\.?\s+(\S.*)$/.exec(line))) { level = 3; text = m[1]; }
      else if ((m = /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.\s+(\S.*)$/.exec(line))) { level = 1; text = m[1]; }
    }
    if (level) { flush(); items.push({ kind: 'h', level, text: text.trim() }); } else buf.push(line);
  }
  flush();
  return items;
}

/**
 * 개요 → 문서. opts: { kind, title, subtitle, author, date, cover, toc, number, chapterBreak, placeholders, tocLevels }
 */
export function buildFromOutline(md, opts = {}) {
  const kind = OUTLINE_KINDS.find((k) => k.id === (opts.kind ?? 'report')) ?? OUTLINE_KINDS[0];
  const o = { ...kind, placeholders: true, tocLevels: 3, ...Object.fromEntries(Object.entries(opts).filter(([, v]) => v !== undefined && v !== null && v !== '')) };
  let items = parseOutline(md);
  // 맨 앞의 '# 제목' 하나뿐이면 문서 제목으로 (나머지 제목은 한 단계씩 올림)
  let title = o.title ?? '';
  const h1 = items.filter((x) => x.kind === 'h' && x.level === 1);
  if (!title && items[0]?.kind === 'h' && items[0].level === 1 && h1.length === 1 && items.some((x) => x.kind === 'h' && x.level > 1)) {
    title = items[0].text;
    items = items.slice(1).map((x) => (x.kind === 'h' ? { ...x, level: Math.max(1, x.level - 1) } : x));
  }
  if (!title) title = o.kind === 'minutes' ? '회의록' : '제목 없는 문서';
  // 본문 마크다운 (번호는 글로 붙임 — 한글 · Word 어디서 열어도 같게)
  const counters = [];
  const PLACE = '(여기에 내용을 쓰세요)';
  const out = [];
  items.forEach((it, i) => {
    if (it.kind === 'md') { out.push(it.text, ''); return; }
    const lv = Math.min(6, it.level);
    counters.length = lv;
    for (let k = 0; k < lv; k++) counters[k] ??= 0;
    counters[lv - 1]++;
    for (let k = 0; k < lv - 1; k++) if (!counters[k]) counters[k] = 1;
    const label = numberLabel(o.number, counters.slice(0, lv));
    const text = o.number === 'none' ? it.text : stripNumber(it.text);
    out.push(`${'#'.repeat(lv)} ${label ? `${label} ` : ''}${text}`, '');
    const next = items[i + 1];
    if (o.placeholders && (!next || (next.kind === 'h' && next.level <= lv))) out.push(PLACE, '');
  });
  const body = mdToDoc(out.join('\n')).body;
  const doc = newDoc();
  doc.props.title = title;
  if (o.author) doc.props.author = o.author;
  // 빈 절 안내 문구는 회색 기울임
  for (const b of body) if (b.type === 'p' && paraPlain(b) === PLACE) b.runs = [{ text: PLACE, color: '#8C8C8C', i: true }];
  // 장마다 새 쪽 (첫 장 빼고)
  if (o.chapterBreak) {
    let seen = 0;
    for (const b of body) if (b.type === 'p' && b.style === 'Heading1' && seen++ > 0) b.pPr = { ...(b.pPr ?? {}), pageBreakBefore: true };
  }
  const date = o.date ?? '';
  const front = [];
  const pageBreak = () => ({ ...para(''), runs: [{ obj: 'br', kind: 'page' }] });
  if (o.cover) {
    front.push({ ...para(''), pPr: { spc: { before: 216 } } });
    front.push({ ...para(title), style: 'Title', pPr: { align: 'center' } });
    if (o.subtitle) front.push({ ...para(o.subtitle), style: 'Subtitle', pPr: { align: 'center' } });
    const by = [o.author, date].filter(Boolean).join(' · ');
    if (by) front.push({ ...para(by), pPr: { align: 'center', spc: { before: 160 } } });
    front.push(pageBreak());
  } else {
    front.push({ ...para(title), style: 'Title', pPr: { align: o.kind === 'gongmun' || o.kind === 'minutes' ? 'center' : undefined } });
    if (o.subtitle) front.push({ ...para(o.subtitle), style: 'Subtitle' });
    const by = [o.author, date].filter(Boolean).join(' · ');
    if (by && !o.minutes) front.push({ ...para(by), pPr: { align: 'right' } });
  }
  if (o.minutes) {
    const row = (k, v) => ({ cells: [{ blocks: [{ ...para(k), runs: [{ text: k, b: true }] }], shd: '#F2F2F2' }, { blocks: [para(v)] }] });
    front.push({ type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: {}, grid: [120, 480], rows: [row('일시', date), row('장소', ''), row('참석자', ''), row('작성자', o.author ?? '')] }, para(''));
  }
  doc.body = [...front, ...body];
  if (!doc.body.length) doc.body = [para('')];
  if (o.toc && headings(doc, o.tocLevels).length) {
    const toc = buildToc(doc, { levels: o.tocLevels });
    const at = o.cover ? front.length : front.length;
    doc.body.splice(at, 0, ...toc, pageBreak());
  }
  return doc;
}

/** 문서의 제목들 → 개요 (제목마다 그 아래 첫 문장을 '- ' 항목으로) */
export function outlineFromDoc(doc, { bullets = true, maxLevel = 4 } = {}) {
  const hs = headings(doc, maxLevel);
  if (!hs.length) {
    // 제목이 없으면 문단 첫머리로 개요
    const paras = docText(doc).split('\n').map((t) => t.trim()).filter((t) => t.length > 1).slice(0, 30);
    return paras.map((t) => `- ${firstSentence(t)}`).join('\n');
  }
  const lines = [];
  const minLv = Math.min(...hs.map((h) => h.level));
  hs.forEach((h, i) => {
    lines.push(`${'#'.repeat(h.level - minLv + 1)} ${h.text}`);
    if (!bullets) return;
    const end = hs[i + 1]?.path[0] ?? doc.body.length;
    const sent = [];
    for (let k = h.path[0] + 1; k < end && sent.length < 2; k++) {
      const b = doc.body[k];
      if (b.type !== 'p') continue;
      const t = paraPlain(b).trim();
      if (t.length > 1) sent.push(firstSentence(t));
    }
    for (const s of sent) lines.push(`- ${s}`);
    lines.push('');
  });
  return lines.join('\n').trim();
}
function firstSentence(t) {
  const m = /^(.{4,120}?[.!?。])(\s|$)/.exec(t);
  const s = m ? m[1] : t;
  return s.length > 120 ? `${s.slice(0, 117)}…` : s;
}
