// 명령 (리본 · 단축키 · MCP 가 모두 이 이름으로 실행)
import { S, change, run, register, emit, curSel, setDoc } from './state.js';
import {
  blockAt, paraLen, posEq, orderPos, formatRange, rangeParas, formatParas, effPara, effRun, propsAt, insertRuns, splitPara, deleteRange, firstPos, lastPos, walkParas,
  uid, para, table, clone, containerOf, normalizeRuns, splitRuns, runLen, paraPlain, paraText, findAll, insertText, buildToc, headings, cm, PAPERS, MARGINS, newSection,
  bulletList, numberList, multiList, docStats, sections, styleChain, insertBlocks, sliceRange, newDoc, RUN_KEYS,
} from './model.js';
import { E, setSel, typeText, deleteSelection, newParagraph, insertObject, selectedBlocks, pasteBlocks } from './editor.js';
import { V, cellSelection, setZoom, applyZoomMode, relayout, drawSelection, scrollToCaret } from './view.js';

const sel = () => curSel();
const collapsed = () => posEq(sel().a, sel().f);
const author = () => S.opts.userName || '사용자';
const now = () => new Date().toISOString();

/** 선택 범위 (빈 선택이면 그 단어 · Word: 단어 안의 캐럿에 서식 → 단어 전체) */
function fmtRange({ word = true } = {}) {
  const { a, f } = sel();
  if (!posEq(a, f)) return orderPos(a, f);
  if (!word) return null;
  const p = blockAt(S.doc, f.p);
  const t = paraText(p);
  const isW = (c) => /[\p{L}\p{N}_]/u.test(c ?? '');
  if (isW(t[f.o - 1]) && isW(t[f.o])) {
    let s = f.o; let e = f.o;
    while (s > 0 && isW(t[s - 1])) s--;
    while (e < t.length && isW(t[e])) e++;
    return [{ p: f.p, o: s }, { p: f.p, o: e }];
  }
  return null;
}
/** 글자 서식 적용 (빈 선택: 다음에 칠 글자 서식) */
export function applyRun(props, { toggleKey = null, label = null } = {}) {
  void label;
  const cs = cellSelection();
  if (cs) {
    change((d) => {
      for (let r = cs.r1; r <= cs.r2; r++) for (let c = cs.c1; c <= cs.c2; c++) {
        const cell = S.doc.body[cs.t].rows[r]?.cells[c];
        if (!cell) continue;
        cell.blocks.forEach((b, i) => { if (b.type === 'p') formatRange(d, { p: [cs.t, r, c, i], o: 0 }, { p: [cs.t, r, c, i], o: paraLen(b) }, props); });
      }
      return S.sel;
    });
    return;
  }
  const rg = fmtRange();
  if (!rg) {
    S.typing = { ...(S.typing ?? propsAt(blockAt(S.doc, sel().f.p), sel().f.o)), ...(typeof props === 'function' ? props(S.typing ?? propsAt(blockAt(S.doc, sel().f.p), sel().f.o)) : props) };
    for (const [k, v] of Object.entries(S.typing)) if (v === undefined || v === null) delete S.typing[k];
    // 빈 문단: 문단 표시 서식도 바꿈 (보이는 높이)
    const p = blockAt(S.doc, sel().f.p);
    if (p && !paraLen(p)) { const t = S.typing; change((d) => { const b = d.block(sel().f.p); b.mark = { ...(b.mark ?? {}), ...t }; return S.sel; }, { key: 'typing' }); S.typing = t; }
    emit('selection');
    return;
  }
  const keepSel = S.sel;
  change((d) => { formatRange(d, rg[0], rg[1], props); return keepSel; });
  void toggleKey;
}
/** 선택한 글이 모두 그 서식인가 (단추 켜짐 · 토글) */
export function selProp(key) {
  const { a, f } = sel();
  const rp = (r, p) => effRun(S.doc, p, r);
  if (posEq(a, f)) {
    if (S.typing && key in S.typing) return S.typing[key];
    const p = blockAt(S.doc, f.p);
    return p ? rp(propsAt(p, f.o), p)[key] : undefined;
  }
  const [s, e] = orderPos(a, f);
  let val; let first = true;
  for (const path of rangeParas(S.doc, s, e)) {
    const p = blockAt(S.doc, path);
    if (p?.type !== 'p') continue;
    const st = path === s.p || pathEq(path, s.p) ? s.o : 0;
    const en = pathEq(path, e.p) ? e.o : paraLen(p);
    let pos = 0;
    for (const r of p.runs) {
      const n = runLen(r);
      if (!r.obj && pos + n > st && pos < en) { const v = rp(r, p)[key]; if (first) { val = v; first = false; } else if (JSON.stringify(v) !== JSON.stringify(val)) return undefined; }
      pos += n;
    }
  }
  return val;
}
const pathEq = (x, y) => x.length === y.length && x.every((v, i) => v === y[i]);
function toggle(key, on = true, off = undefined) {
  const cur = selProp(key);
  applyRun({ [key]: cur && cur !== 'none' ? off : on });
}
/** 선택한 문단 경로들 (표 칸 선택 포함) */
export function selParas() {
  const cs = cellSelection();
  if (cs) {
    const out = [];
    for (let r = cs.r1; r <= cs.r2; r++) for (let c = cs.c1; c <= cs.c2; c++) S.doc.body[cs.t].rows[r]?.cells[c]?.blocks.forEach((b, i) => { if (b.type === 'p') out.push([cs.t, r, c, i]); });
    return out;
  }
  const { a, f } = sel();
  return rangeParas(S.doc, a, f).filter((p) => blockAt(S.doc, p)?.type === 'p');
}
export function applyPara(fn, key = null) {
  const keepSel = S.sel;
  change((d) => { formatParas(d, selParas(), fn); return keepSel; }, { key });
}
const setPPr = (patch) => (p) => { p.pPr = { ...p.pPr, ...patch }; for (const [k, v] of Object.entries(patch)) if (v === undefined) delete p.pPr[k]; };

// ───────────── 글꼴 ─────────────
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
register({
  bold: () => toggle('b'),
  italic: () => toggle('i'),
  underline: (style) => { if (style) applyRun({ u: style === 'none' ? undefined : style }); else toggle('u', true); },
  strike: () => toggle('s'),
  dstrike: () => toggle('dstrike'),
  superscript: () => applyRun({ va: selProp('va') === 'super' ? undefined : 'super' }),
  subscript: () => applyRun({ va: selProp('va') === 'sub' ? undefined : 'sub' }),
  fontFamily: (name) => applyRun({ font: name, fontEa: name }),
  fontEa: (name) => applyRun({ fontEa: name }),
  fontSize: (pt) => { const v = Number(pt); if (v > 0 && v <= 1638) applyRun({ sz: v }); },
  growFont: () => applyRun((r) => ({ sz: SIZES.find((s) => s > (r.sz ?? curSize())) ?? Math.min(1638, (r.sz ?? curSize()) + 10) })),
  shrinkFont: () => applyRun((r) => ({ sz: [...SIZES].reverse().find((s) => s < (r.sz ?? curSize())) ?? Math.max(1, (r.sz ?? curSize()) - 1) })),
  growFont1: () => applyRun((r) => ({ sz: Math.min(1638, (r.sz ?? curSize()) + 1) })),
  shrinkFont1: () => applyRun((r) => ({ sz: Math.max(1, (r.sz ?? curSize()) - 1) })),
  fontColor: (c) => applyRun({ color: c === 'auto' ? undefined : c ?? S.lastFontColor ?? '#C00000' }),
  highlight: (c) => applyRun({ hl: c === 'none' ? undefined : c ?? S.lastHighlight ?? 'yellow' }),
  charShade: () => applyRun({ shd: selProp('shd') ? undefined : '#D9D9D9' }),
  smallCaps: () => toggle('smallCaps'),
  allCaps: () => toggle('caps'),
  hidden: () => toggle('hidden'),
  charSpacing: (pt) => applyRun({ spc: Number(pt) || undefined }),
  charScale: (pct) => applyRun({ scale: Number(pct) === 100 ? undefined : Number(pct) }),
  emphasisMark: (k) => applyRun({ emph: k === 'none' ? undefined : k ?? 'dot' }),
  textEffect: (k) => applyRun({ outline: k === 'outline' ? true : undefined, shadow: k === 'shadow' ? true : undefined, emboss: k === 'emboss' ? true : undefined, imprint: k === 'imprint' ? true : undefined }),
  clearFormat: () => {
    const keep = (r) => { const o = {}; for (const k of RUN_KEYS) if (!['ins', 'del', 'link'].includes(k)) o[k] = undefined; return o; };
    applyRun(keep);
    applyPara((p) => { if (p.style !== 'Normal' && !/^Heading/.test(p.style ?? '')) p.style = 'Normal'; p.pPr = p.pPr.num ? { num: p.pPr.num } : {}; });
  },
  clearCharFormat: () => applyRun(() => { const o = {}; for (const k of RUN_KEYS) if (!['ins', 'del', 'link'].includes(k)) o[k] = undefined; return o; }),
  changeCase: (mode) => {
    const rg = fmtRange();
    if (!rg) return;
    const keepSel = S.sel;
    change((d) => {
      let firstWord = true;
      for (const path of rangeParas(d.doc, rg[0], rg[1])) {
        const p = d.block(path);
        const s = pathEq(path, rg[0].p) ? rg[0].o : 0;
        const e = pathEq(path, rg[1].p) ? rg[1].o : paraLen(p);
        const [x, rest] = splitRuns(p.runs, s);
        const [mid, y] = splitRuns(rest, e - s);
        const conv = (t) => {
          switch (mode) {
            case 'upper': return t.toUpperCase();
            case 'lower': return t.toLowerCase();
            case 'title': return t.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
            case 'toggle': return [...t].map((c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join('');
            case 'sentence': { const o = t.toLowerCase().replace(/(^|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toUpperCase()); if (firstWord) { firstWord = false; return o.replace(/^\s*\p{L}/u, (c) => c.toUpperCase()); } return o; }
            case 'full': return t.replace(/[!-~]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0)).replace(/ /g, '　');
            case 'half': return t.replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/　/g, ' ');
            default: return t;
          }
        };
        p.runs = normalizeRuns([...x, ...mid.map((r) => (r.obj ? r : { ...r, text: conv(r.text) })), ...y]);
      }
      return keepSel;
    });
  },
  cycleCase: () => { const t = selectedText(); run('changeCase', t === t.toUpperCase() ? 'lower' : t === t.toLowerCase() ? 'title' : 'upper'); },
}, {
  bold: { label: '굵게' }, italic: { label: '기울임꼴' }, underline: { label: '밑줄', args: ['style?'] }, strike: { label: '취소선' },
  superscript: { label: '위 첨자' }, subscript: { label: '아래 첨자' }, fontFamily: { label: '글꼴', args: ['name'] }, fontSize: { label: '글꼴 크기', args: ['pt'] },
  growFont: { label: '글꼴 크기 크게' }, shrinkFont: { label: '글꼴 크기 작게' }, fontColor: { label: '글꼴 색', args: ['#RRGGBB|auto'] }, highlight: { label: '형광펜', args: ['yellow|green|…|none'] },
  clearFormat: { label: '모든 서식 지우기' }, changeCase: { label: '대/소문자 바꾸기', args: ['upper|lower|title|sentence|toggle|full|half'] },
});
function curSize() { const s = selProp('sz'); return s ?? 10; }
export function selectedText() { return selectedBlocks().map((b) => (b.type === 'p' ? paraPlain(b) : '')).join('\n'); }

// ───────────── 단락 ─────────────
register({
  alignLeft: () => applyPara(setPPr({ align: 'left' })),
  alignCenter: () => applyPara((p) => { p.pPr = { ...p.pPr, align: effPara(S.doc, p).align === 'center' ? 'left' : 'center' }; }),
  alignRight: () => applyPara((p) => { p.pPr = { ...p.pPr, align: effPara(S.doc, p).align === 'right' ? 'left' : 'right' }; }),
  alignJustify: () => applyPara((p) => { p.pPr = { ...p.pPr, align: effPara(S.doc, p).align === 'justify' ? 'left' : 'justify' }; }),
  alignDistribute: () => applyPara(setPPr({ align: 'distribute' })),
  lineSpacing: (v) => { const n = Number(v); applyPara((p) => { const ep = effPara(S.doc, p); p.pPr = { ...p.pPr, spc: { ...(p.pPr.spc ?? {}), line: n, rule: ep.spc?.rule === 'pct' ? 'pct' : 'auto' } }; }); },
  spaceBefore: (v) => applyPara((p) => { p.pPr = { ...p.pPr, spc: { ...(p.pPr.spc ?? {}), before: Number(v) } }; }),
  spaceAfter: (v) => applyPara((p) => { p.pPr = { ...p.pPr, spc: { ...(p.pPr.spc ?? {}), after: Number(v) } }; }),
  toggleSpaceBefore: () => applyPara((p) => { const b = effPara(S.doc, p).spc?.before ?? 0; p.pPr = { ...p.pPr, spc: { ...(p.pPr.spc ?? {}), before: b ? 0 : 16 } }; }),
  toggleSpaceAfter: () => applyPara((p) => { const b = effPara(S.doc, p).spc?.after ?? 0; p.pPr = { ...p.pPr, spc: { ...(p.pPr.spc ?? {}), after: b ? 0 : 10.67 } }; }),
  indentMore: () => {
    const paths = selParas();
    if (paths.some((pt) => blockAt(S.doc, pt)?.pPr?.num)) { applyPara((p) => { if (p.pPr.num) p.pPr = { ...p.pPr, num: { ...p.pPr.num, lvl: Math.min(8, (p.pPr.num.lvl ?? 0) + 1) } }; }); return; }
    applyPara((p) => { const ep = effPara(S.doc, p); const l = ep.ind?.l ?? 0; const step = 48; p.pPr = { ...p.pPr, ind: { ...(p.pPr.ind ?? {}), l: (Math.floor(l / step) + 1) * step } }; });
  },
  indentLess: () => {
    const paths = selParas();
    if (paths.some((pt) => blockAt(S.doc, pt)?.pPr?.num)) { applyPara((p) => { if (p.pPr.num) p.pPr = { ...p.pPr, num: { ...p.pPr.num, lvl: Math.max(0, (p.pPr.num.lvl ?? 0) - 1) } }; }); return; }
    applyPara((p) => { const ep = effPara(S.doc, p); const l = ep.ind?.l ?? 0; const step = 48; p.pPr = { ...p.pPr, ind: { ...(p.pPr.ind ?? {}), l: Math.max(0, (Math.ceil(l / step) - 1) * step) } }; });
  },
  hangingIndent: () => applyPara((p) => { const ep = effPara(S.doc, p); const l = (ep.ind?.l ?? 0) + 48; p.pPr = { ...p.pPr, ind: { ...(p.pPr.ind ?? {}), l, hanging: 48 } }; delete p.pPr.ind.first; }),
  firstLineIndent: (v = 20) => applyPara((p) => { p.pPr = { ...p.pPr, ind: { ...(p.pPr.ind ?? {}), first: Number(v) } }; delete p.pPr.ind.hanging; }),
  paraShading: (c) => applyPara(setPPr({ shd: c === 'none' ? undefined : c })),
  paraBorder: (kind, b = { style: 'single', w: 0.5, color: '#000000', space: 1 }) => applyPara((p) => {
    const cur = { ...(p.pPr.border ?? {}) };
    if (kind === 'none') { p.pPr = { ...p.pPr }; delete p.pPr.border; return; }
    const sides = kind === 'all' || kind === 'box' ? ['top', 'bottom', 'left', 'right'] : [kind];
    for (const s of sides) cur[s] = cur[s] && kind !== 'all' && kind !== 'box' ? undefined : { ...b };
    for (const [k, v] of Object.entries(cur)) if (!v) delete cur[k];
    p.pPr = { ...p.pPr, border: Object.keys(cur).length ? cur : undefined };
    if (!p.pPr.border) delete p.pPr.border;
  }),
  keepWithNext: () => applyPara((p) => { p.pPr = { ...p.pPr, keepNext: !effPara(S.doc, p).keepNext }; }),
  keepLinesTogether: () => applyPara((p) => { p.pPr = { ...p.pPr, keepLines: !effPara(S.doc, p).keepLines }; }),
  pageBreakBefore: () => applyPara((p) => { p.pPr = { ...p.pPr, pageBreakBefore: !effPara(S.doc, p).pageBreakBefore }; }),
  setTabs: (tabs) => applyPara(setPPr({ tabs })),
  setIndent: (ind) => applyPara((p) => { p.pPr = { ...p.pPr, ind: { ...(p.pPr.ind ?? {}), ...ind } }; for (const [k, v] of Object.entries(ind)) if (v == null) delete p.pPr.ind[k]; }, 'indent'),
  paragraphProps: (pPr) => applyPara((p) => { p.pPr = { ...p.pPr, ...pPr }; }),
  dropCap: (kind) => applyPara((p) => { p.pPr = { ...p.pPr, dropCap: kind === 'none' ? undefined : { kind: kind ?? 'drop', lines: 3 } }; if (!p.pPr.dropCap) delete p.pPr.dropCap; }),
}, {
  alignLeft: { label: '왼쪽 맞춤' }, alignCenter: { label: '가운데 맞춤' }, alignRight: { label: '오른쪽 맞춤' }, alignJustify: { label: '양쪽 맞춤' }, alignDistribute: { label: '균등 분할' },
  lineSpacing: { label: '줄 간격', args: ['배수 (1, 1.15, 1.5, 2 …)'] }, spaceBefore: { label: '단락 앞 간격', args: ['px'] }, spaceAfter: { label: '단락 뒤 간격', args: ['px'] },
  indentMore: { label: '들여쓰기' }, indentLess: { label: '내어쓰기' }, paraShading: { label: '음영', args: ['#RRGGBB|none'] }, paraBorder: { label: '테두리', args: ['top|bottom|left|right|all|none'] },
  keepWithNext: { label: '다음 단락과 함께' }, keepLinesTogether: { label: '현재 단락을 나누지 않음' }, pageBreakBefore: { label: '단락 앞에서 페이지 나누기' },
});

// ───────────── 스타일 ─────────────
register({
  applyStyle: (id) => {
    const st = S.doc.styles[id];
    if (!st) { emit('toast', `스타일 '${id}' 이(가) 없습니다`); return; }
    if (st.type === 'char') { const rg = fmtRange(); if (rg) { const keep = S.sel; change((d) => { formatRange(d, rg[0], rg[1], { cstyle: id === 'DefaultParagraphFont' ? undefined : id }); return keep; }); } return; }
    applyPara((p) => {
      p.style = id;
      // 스타일을 바꾸면 직접 서식 중 스타일과 겹치는 것은 지움 (Word: 단락 서식 초기화)
      const keepNum = p.pPr.num && !/^Heading|Title/.test(id) ? { num: p.pPr.num } : {};
      p.pPr = { ...keepNum, ...(p.pPr.align && id === 'Normal' ? {} : {}) };
      if (/^Heading|Title|Subtitle/.test(id)) p.runs = p.runs.map((r) => { if (r.obj) return r; const c = { ...r }; delete c.sz; delete c.font; delete c.fontEa; delete c.color; delete c.b; return c; });
    });
  },
  normalStyle: () => run('applyStyle', 'Normal'),
  heading1: () => run('applyStyle', 'Heading1'), heading2: () => run('applyStyle', 'Heading2'), heading3: () => run('applyStyle', 'Heading3'),
  newStyleFromSel: (name, type = 'para') => {
    const p = blockAt(S.doc, sel().f.p);
    const rPr = propsAt(p, sel().f.o);
    delete rPr.link; delete rPr.ins; delete rPr.del;
    const id = uniqueStyleId(name);
    change((d) => {
      const st = d.prop('styles');
      st[id] = { id, name, type, basedOn: type === 'para' ? p.style ?? 'Normal' : undefined, next: id, quick: true, order: 200 + Object.keys(st).length, pPr: type === 'para' ? { ...p.pPr } : undefined, rPr };
      if (type === 'para') { const b = d.block(sel().f.p); b.style = id; }
      return S.sel;
    });
    return id;
  },
  updateStyleFromSel: (id) => {
    const p = blockAt(S.doc, sel().f.p);
    const rPr = propsAt(p, sel().f.o);
    delete rPr.link; delete rPr.ins; delete rPr.del;
    change((d) => { const st = d.prop('styles'); st[id] = { ...st[id], pPr: { ...(st[id].pPr ?? {}), ...p.pPr }, rPr: { ...(st[id].rPr ?? {}), ...rPr } }; return S.sel; });
  },
  modifyStyle: (id, patch) => change((d) => { const st = d.prop('styles'); st[id] = { ...st[id], ...patch }; return S.sel; }),
  deleteStyle: (id) => change((d) => {
    const st = d.prop('styles');
    if (!st[id] || ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Title', 'DefaultParagraphFont', 'TableNormal'].includes(id)) return false;
    delete st[id];
    // 쓰던 문단은 기준 스타일로
    walkParas(d.doc.body, (p, path) => { if (p.style === id) d.block(path).style = 'Normal'; });
    return S.sel;
  }),
  styleSet: (key) => change((d) => { const st = d.prop('styles'); for (const [k, v] of Object.entries(STYLE_SETS[key] ?? {})) st[k] = { ...st[k], rPr: { ...(st[k]?.rPr ?? {}), ...(v.rPr ?? {}) }, pPr: { ...(st[k]?.pPr ?? {}), ...(v.pPr ?? {}) } }; return S.sel; }),
  themeFonts: (major, minor) => change((d) => { d.doc.themeFonts = { major: { latin: major, ea: major }, minor: { latin: minor ?? major, ea: minor ?? major } }; return S.sel; }),
  paragraphSpacingPreset: (k) => change((d) => {
    const P = { none: { before: 0, after: 0, line: 1 }, compact: { before: 0, after: 5.33, line: 1 }, tight: { before: 0, after: 8, line: 1.08 }, open: { before: 0, after: 13.33, line: 1.15 }, relaxed: { before: 0, after: 8, line: 1.5 }, double: { before: 0, after: 10.67, line: 2 }, default: { before: 0, after: 10.67, line: 1.08 } }[k];
    if (!P) return false;
    d.doc.defaults = { ...d.doc.defaults, pPr: { ...d.doc.defaults.pPr, spc: { ...P, rule: 'auto' } } };
    return S.sel;
  }),
}, { applyStyle: { label: '스타일 적용', args: ['Normal|Heading1|Heading2|…|Title|Quote|스타일 id'] }, newStyleFromSel: { label: '선택 영역으로 새 스타일', args: ['이름'] } });
function uniqueStyleId(name) { let id = name.replace(/\s+/g, ''); let k = 1; while (S.doc.styles[id]) id = `${name.replace(/\s+/g, '')}${++k}`; return id; }
const STYLE_SETS = {
  basic: { Heading1: { rPr: { color: '#2F5496', font: '+mj', fontEa: '+mj', sz: 16, b: undefined } }, Title: { rPr: { sz: 28 } } },
  modern: { Heading1: { rPr: { color: '#C55A11', sz: 18, b: true }, pPr: { border: { bottom: { style: 'single', w: 1, color: '#C55A11', space: 2 } } } }, Heading2: { rPr: { color: '#C55A11', sz: 14 } }, Title: { rPr: { color: '#C55A11', sz: 30, b: true } } },
  classic: { Heading1: { rPr: { color: '#000000', font: 'Times New Roman', fontEa: '바탕', sz: 16, b: true } }, Heading2: { rPr: { color: '#000000', font: 'Times New Roman', fontEa: '바탕', sz: 13, b: true } }, Title: { rPr: { font: 'Times New Roman', fontEa: '바탕', sz: 26 } } },
  elegant: { Heading1: { rPr: { color: '#7030A0', smallCaps: true, sz: 16 } }, Heading2: { rPr: { color: '#7030A0', sz: 13, i: true } }, Title: { rPr: { color: '#7030A0', sz: 30, smallCaps: true } } },
  shaded: { Heading1: { rPr: { color: '#FFFFFF', sz: 15, b: true }, pPr: { shd: '#2F5496' } }, Heading2: { rPr: { color: '#2F5496', sz: 13, b: true }, pPr: { shd: '#D9E2F3' } } },
  korean: { Heading1: { rPr: { font: '맑은 고딕', fontEa: '맑은 고딕', color: '#1F3864', sz: 16, b: true } }, Heading2: { rPr: { font: '맑은 고딕', fontEa: '맑은 고딕', color: '#1F3864', sz: 13, b: true } }, Title: { rPr: { font: '맑은 고딕', fontEa: '맑은 고딕', sz: 24, b: true } } },
};
export const STYLE_SET_LIST = [['basic', '기본'], ['modern', '현대적'], ['classic', '고전'], ['elegant', '우아함'], ['shaded', '음영'], ['korean', '한국어 보고서']];

// ───────────── 목록 ─────────────
function ensureList(def, key) {
  // 같은 정의가 이미 있으면 재사용
  for (const [id, l] of Object.entries(S.doc.lists)) if (l.key === key) return id;
  const id = String(Math.max(0, ...Object.keys(S.doc.lists).map(Number).filter(Number.isFinite)) + 1);
  return { id, def: { ...def, key } };
}
function applyList(def, key) {
  const paths = selParas();
  const allOn = paths.every((pt) => { const p = blockAt(S.doc, pt); return p?.pPr?.num && S.doc.lists[p.pPr.num.id]?.key === key; });
  const r = ensureList(def, key);
  const id = typeof r === 'string' ? r : r.id;
  const keepSel = S.sel;
  change((d) => {
    if (typeof r !== 'string') { const ls = d.prop('lists'); ls[id] = r.def; }
    formatParas(d, paths, (p) => {
      if (allOn) { p.pPr = { ...p.pPr }; delete p.pPr.num; if (p.style === 'ListParagraph') p.style = 'Normal'; return; }
      const lvl = p.pPr.num?.lvl ?? 0;
      p.pPr = { ...p.pPr, num: { id, lvl } };
      if (p.style === 'Normal') p.style = 'ListParagraph';
    });
    return keepSel;
  });
}
register({
  bullets: (char = '●') => applyList(bulletList(char), `bullet:${char}`),
  numbering: (fmt = 'decimal', suffix = '.') => {
    const pre = suffix === '()' ? '(' : '';
    const suf = suffix === '()' ? ')' : suffix;
    const def = { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: i === 0 ? fmt : ['decimal', 'lowerLetter', 'lowerRoman'][i % 3], text: i === 0 ? `${pre}%1${suf}` : `%${i + 1}.`, start: 1, ind: { l: 26.67 + i * 26.67, hanging: 26.67 } })) };
    applyList(def, `num:${fmt}:${suffix}`);
  },
  multilevel: (kind = 'outline') => applyList(multiList(kind), `multi:${kind}`),
  restartNumbering: () => applyPara((p) => { if (p.pPr.num) p.pPr = { ...p.pPr, num: { ...p.pPr.num, restart: true } }; }),
  continueNumbering: () => applyPara((p) => { if (p.pPr.num) { p.pPr = { ...p.pPr, num: { ...p.pPr.num } }; delete p.pPr.num.restart; } }),
  setListLevel: (lvl) => applyPara((p) => { if (p.pPr.num) p.pPr = { ...p.pPr, num: { ...p.pPr.num, lvl: Number(lvl) } }; }),
  removeList: () => applyPara((p) => { p.pPr = { ...p.pPr }; delete p.pPr.num; if (p.style === 'ListParagraph') p.style = 'Normal'; }),
}, { bullets: { label: '글머리 기호', args: ['기호?'] }, numbering: { label: '번호 매기기', args: ['decimal|ganada|upperRoman|lowerLetter|decimalCircle…', '접미사 . ) ()'] }, multilevel: { label: '다단계 목록', args: ['outline|legal|korean'] } });

// ───────────── 선택 · 편집 ─────────────
register({
  undo: () => { const s = S.history.undo(S.doc, S.sel); if (!s) return; S.doc = s.doc; S.sel = s.sel ?? S.sel; S.dirty = true; emit('change', { scope: 'doc', undo: true }); },
  redo: () => { const s = S.history.redo(S.doc, S.sel); if (!s) return; S.doc = s.doc; S.sel = s.sel ?? S.sel; S.dirty = true; emit('change', { scope: 'doc', undo: true }); },
  repeat: () => { const r = S.lastRepeat; if (r) run(r.name, ...r.args); },
  selectAll: () => setSel(firstPos(S.doc.body), lastPos(S.doc.body), { scroll: false }),
  delete: () => deleteSelection(),
  typeText: (text) => typeText(String(text ?? '')),
  newParagraph: () => newParagraph(),
  pageBreak: () => insertObject({ obj: 'br', kind: 'page' }),
  columnBreak: () => insertObject({ obj: 'br', kind: 'column' }),
  lineBreak: () => insertObject({ obj: 'br' }),
  sectionBreak: (kind = 'nextPage') => {
    const pos = collapsed() ? sel().f : orderPos(sel().a, sel().f)[1];
    if (pos.p.length > 1) { emit('toast', '표 안에는 구역 나누기를 넣을 수 없습니다'); return; }
    change((d) => {
      const at = splitPara(d, pos);
      const prevPath = [at.p[0] - 1];
      const prev = d.block(prevPath);
      // 앞 구역 = 지금 구역 설정을 복사 (구역 나누기 문단에 저장)
      const cur = sectionAtTop(S.doc, pos.p[0]);
      prev.secBreak = { kind, sec: clone(cur) };
      return at;
    });
  },
  insertSymbol: (ch, font) => (font ? insertObject({ obj: 'sym', char: ch, font, ...propsAt(blockAt(S.doc, sel().f.p), sel().f.o) }) : typeText(ch)),
  insertText: (text) => typeText(String(text ?? '')),
  insertDate: (fmt = 'long') => {
    const d = new Date();
    const t = fmt === 'iso' ? d.toISOString().slice(0, 10) : fmt === 'short' ? `${d.getFullYear()}. ${d.getMonth() + 1}. ${d.getDate()}.` : d.toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' });
    typeText(t);
  },
  insertField: (code) => insertObject({ obj: 'field', code, result: code === 'PAGE' ? '1' : code === 'DATE' ? new Date().toLocaleDateString('ko-KR') : code === 'AUTHOR' ? S.doc.props.author || author() : code === 'TITLE' ? S.doc.props.title || S.docName : code === 'NUMWORDS' ? String(docStats(S.doc).words) : '' }),
  insertLink: (url, text) => {
    if (!url) return;
    const link = url.startsWith('#') ? { anchor: url.slice(1) } : { url: /^[a-z]+:/i.test(url) ? url : `https://${url}` };
    if (!collapsed()) { const [a, b] = orderPos(sel().a, sel().f); const keep = S.sel; change((d) => { formatRange(d, a, b, { link }); return keep; }); return; }
    change((d) => insertRuns(d, sel().f, [{ ...propsAt(blockAt(d.doc, sel().f.p), sel().f.o), text: text || url, link }]));
  },
  removeLink: () => { const rg = fmtRange() ?? linkRangeAt(); if (rg) { const keep = S.sel; change((d) => { formatRange(d, rg[0], rg[1], { link: undefined }); return keep; }); } },
  insertBookmark: (name) => {
    if (!name) return;
    const [a, b] = orderPos(sel().a, sel().f);
    change((d) => {
      insertRuns(d, b, [{ obj: 'mark', kind: 'bm', name, edge: 'end' }]);
      insertRuns(d, a, [{ obj: 'mark', kind: 'bm', name, edge: 'start' }]);
      return S.sel;
    });
  },
  insertFootnote: (kind = 'foot') => {
    const id = uid('n');
    const pos = collapsed() ? sel().f : orderPos(sel().a, sel().f)[1];
    change((d) => {
      const notes = d.prop('notes');
      notes[kind] = { ...notes[kind], [id]: [para('', { style: kind === 'end' ? 'EndnoteText' : 'FootnoteText' })] };
      return insertRuns(d, pos, [{ obj: 'note', id, kind }]);
    });
    emit('editNote', { id, kind });
  },
  insertTable: (rows = 3, cols = 3, style = 'TableGrid') => {
    rows = Math.max(1, Math.min(500, Number(rows))); cols = Math.max(1, Math.min(63, Number(cols)));
    const pos = collapsed() ? sel().f : orderPos(sel().a, sel().f)[0];
    const w = bodyWidthAt(pos);
    const t = table(rows, cols, { style, width: w });
    change((d) => {
      const lst = d.list(pos.p);
      const i = pos.p[pos.p.length - 1];
      const p = blockAt(d.doc, pos.p);
      let at = i + 1;
      if (paraLen(p) && pos.o > 0) { const n = splitPara(d, pos); at = n.p[n.p.length - 1]; }
      else if (!paraLen(p)) at = i; // 빈 문단 자리에
      d.owned.add(t);
      lst.splice(at, 0, t);
      // 표 뒤에 문단이 없으면 하나
      if (!lst[at + 1]) lst.push(para(''));
      return { p: [...pos.p.slice(0, -1), at, 0, 0, 0], o: 0 };
    });
  },
}, {
  undo: { label: '실행 취소' }, redo: { label: '다시 실행' }, selectAll: { label: '모두 선택' }, typeText: { label: '글 입력', args: ['text'] }, newParagraph: { label: '새 단락 (Enter)' },
  pageBreak: { label: '페이지 나누기' }, sectionBreak: { label: '구역 나누기', args: ['nextPage|continuous|evenPage|oddPage'] }, insertLink: { label: '링크', args: ['url', 'text?'] },
  insertFootnote: { label: '각주 삽입', args: ['foot|end'] }, insertTable: { label: '표 삽입', args: ['rows', 'cols', 'style?'] }, insertDate: { label: '날짜 삽입', args: ['long|short|iso'] },
});
function linkRangeAt() {
  const pos = sel().f;
  const p = blockAt(S.doc, pos.p);
  let o = 0;
  for (const r of p.runs) { const n = runLen(r); if (r.link && pos.o >= o && pos.o <= o + n) { let s = o; let e = o + n; return [{ p: pos.p, o: s }, { p: pos.p, o: e }]; } o += n; }
  return null;
}
export function sectionAtTop(doc, i) {
  for (const s of sections(doc)) if (i >= s.start && i < s.end) return s.sec;
  return doc.section;
}
export function bodyWidthAt(pos) {
  const sec = sectionAtTop(S.doc, pos.p[0]);
  const W = sec.page.w - sec.margin.l - sec.margin.r - (sec.margin.gutter ?? 0);
  const n = Math.max(1, sec.cols?.n ?? 1);
  return (W - (n - 1) * (sec.cols?.space ?? 48)) / n;
}

// ───────────── 표 ─────────────
function tableCtx() {
  const pos = sel().f;
  if (pos.p.length < 4) return null;
  // 가장 안쪽 표
  const k = pos.p.length - 4;
  const tpath = pos.p.slice(0, k + 1);
  return { tpath, r: pos.p[k + 1], c: pos.p[k + 2], t: blockAt(S.doc, tpath), cs: cellSelection() };
}
function cellPos(tpath, r, c, end = false) { const t = blockAt(S.doc, tpath); const cell = t.rows[r].cells[c]; return end ? lastPos(cell.blocks, [...tpath, r, c]) : firstPos(cell.blocks, [...tpath, r, c]); }
const emptyCell = (like) => ({ blocks: [para('', { style: like?.blocks?.[0]?.style ?? 'Normal', pPr: { ...(like?.blocks?.[0]?.pPr ?? {}) } })], ...(like?.shd ? {} : {}) });
register({
  nextCell: () => {
    const tc = tableCtx(); if (!tc) return;
    const { t, r, c, tpath } = tc;
    let nr = r; let nc = c + 1;
    while (nr < t.rows.length) { while (nc < t.rows[nr].cells.length && t.rows[nr].cells[nc].vMerge === 'cont') nc++; if (nc < t.rows[nr].cells.length) break; nr++; nc = 0; }
    if (nr >= t.rows.length) { run('insertRowBelow'); const t2 = blockAt(S.doc, tpath); setSel(cellPos(tpath, t2.rows.length - 1, 0)); return; }
    const a = cellPos(tpath, nr, nc); const b = cellPos(tpath, nr, nc, true);
    setSel(a, b);
  },
  prevCell: () => {
    const tc = tableCtx(); if (!tc) return;
    const { t, r, c, tpath } = tc;
    let nr = r; let nc = c - 1;
    while (nr >= 0) { while (nc >= 0 && t.rows[nr].cells[nc].vMerge === 'cont') nc--; if (nc >= 0) break; nr--; nc = nr >= 0 ? t.rows[nr].cells.length - 1 : 0; }
    if (nr < 0) return;
    setSel(cellPos(tpath, nr, nc), cellPos(tpath, nr, nc, true));
  },
  insertRowBelow: () => tableEdit((tb, r) => { const row = tb.rows[r]; tb.rows.splice(r + 1, 0, { ...row, header: false, cells: row.cells.map((c) => ({ ...emptyCell(c), span: c.span, vMerge: c.vMerge === 'restart' || c.vMerge === 'cont' ? 'cont' : undefined, shd: c.shd })) }); return [r + 1, 0]; }),
  insertRowAbove: () => tableEdit((tb, r) => { const row = tb.rows[r]; tb.rows.splice(r, 0, { ...row, cells: row.cells.map((c) => ({ ...emptyCell(c), span: c.span, shd: c.shd })) }); return [r, 0]; }),
  insertColRight: () => tableEdit((tb, r, c) => { const g = gridCol(tb, r, c, true); insertGridCol(tb, g); return [r, c + 1]; }),
  insertColLeft: () => tableEdit((tb, r, c) => { const g = gridCol(tb, r, c, false); insertGridCol(tb, g); return [r, c]; }),
  deleteRow: () => tableEdit((tb, r, c, cs) => { const r1 = cs ? cs.r1 : r; const r2 = cs ? cs.r2 : r; tb.rows.splice(r1, r2 - r1 + 1); if (!tb.rows.length) return 'remove'; return [Math.min(r1, tb.rows.length - 1), 0]; }),
  deleteCol: () => tableEdit((tb, r, c) => { const g = gridCol(tb, r, c, false); deleteGridCol(tb, g); if (!tb.grid.length) return 'remove'; return [r, Math.max(0, c - 1)]; }),
  deleteTable: () => tableEdit(() => 'remove'),
  mergeCells: () => tableEdit((tb, r, c, cs) => {
    if (!cs) return false;
    const { r1, r2, c1, c2 } = cs;
    // 가로 병합 (같은 행) + 세로 병합 (여러 행)
    const span = tb.rows[r1].cells.slice(c1, c2 + 1).reduce((a, x) => a + (x.span ?? 1), 0);
    const blocks = [];
    for (let rr = r1; rr <= r2; rr++) for (let cc = c1; cc <= Math.min(c2, tb.rows[rr].cells.length - 1); cc++) for (const b of tb.rows[rr].cells[cc].blocks) if (b.type !== 'p' || paraLen(b)) blocks.push(b);
    for (let rr = r1; rr <= r2; rr++) {
      const row = { ...tb.rows[rr], cells: tb.rows[rr].cells.slice() };
      const first = { ...row.cells[c1], span, blocks: rr === r1 ? (blocks.length ? blocks : [para('')]) : [para('')] };
      if (r2 > r1) first.vMerge = rr === r1 ? 'restart' : 'cont';
      row.cells.splice(c1, c2 - c1 + 1, first);
      tb.rows[rr] = row;
    }
    return [r1, c1];
  }),
  splitCell: (cols = 2, rows = 1) => tableEdit((tb, r, c) => {
    cols = Math.max(1, Number(cols));
    const cell = tb.rows[r].cells[c];
    const span = cell.span ?? 1;
    if (span >= cols) {
      const parts = Array.from({ length: cols }, (_, i) => ({ ...emptyCell(cell), span: Math.floor(span / cols) + (i < span % cols ? 1 : 0), blocks: i === 0 ? cell.blocks : [para('')] }));
      tb.rows[r] = { ...tb.rows[r], cells: [...tb.rows[r].cells.slice(0, c), ...parts, ...tb.rows[r].cells.slice(c + 1)] };
    } else {
      // 격자 열 나누기: 다른 행은 그 열을 span 으로
      const g = gridCol(tb, r, c, false);
      const w = tb.grid.slice(g, g + span).reduce((a, b) => a + b, 0);
      const extra = cols - span;
      const nw = w / cols;
      tb.grid = [...tb.grid.slice(0, g), ...Array(cols).fill(nw), ...tb.grid.slice(g + span)];
      tb.rows = tb.rows.map((row, ri) => {
        let gc = 0;
        const cells = row.cells.map((cl) => { const s = cl.span ?? 1; const at = gc; gc += s; if (ri === r && at === g) return cl; if (at <= g && g < at + s) return { ...cl, span: s + extra }; return cl; });
        if (ri !== r) return { ...row, cells };
        const k = cells.indexOf(cell);
        const parts = Array.from({ length: cols }, (_, i) => ({ ...emptyCell(cell), shd: cell.shd, blocks: i === 0 ? cell.blocks : [para('')] }));
        return { ...row, cells: [...cells.slice(0, k), ...parts, ...cells.slice(k + 1)] };
      });
    }
    if (rows > 1) for (let k = 1; k < rows; k++) { const row = tb.rows[r]; tb.rows.splice(r + k, 0, { ...row, cells: row.cells.map((cl) => ({ ...emptyCell(cl), span: cl.span })) }); }
    return [r, c];
  }),
  distributeCols: () => tableEdit((tb, r) => { const w = tb.grid.reduce((a, b) => a + b, 0); tb.grid = tb.grid.map(() => w / tb.grid.length); return [r, 0]; }),
  distributeRows: () => tableEdit((tb, r) => { tb.rows = tb.rows.map((row) => { const c = { ...row }; delete c.h; return c; }); return [r, 0]; }),
  autoFitWindow: () => tableEdit((tb, r) => { const w = bodyWidthAt(sel().f); const cur = tb.grid.reduce((a, b) => a + b, 0); tb.grid = tb.grid.map((x) => (x * w) / cur); return [r, 0]; }),
  autoFitContents: () => tableEdit((tb, r) => {
    const lens = tb.grid.map(() => 1);
    for (const row of tb.rows) { let gc = 0; for (const cl of row.cells) { const s = cl.span ?? 1; if (s === 1) lens[gc] = Math.max(lens[gc], Math.min(60, Math.max(...cl.blocks.map((b) => (b.type === 'p' ? paraPlain(b).length : 10))))); gc += s; } }
    const total = tb.grid.reduce((a, b) => a + b, 0);
    const sum = lens.reduce((a, b) => a + b, 0);
    tb.grid = lens.map((l) => Math.max(36, (l / sum) * total));
    return [r, 0];
  }),
  setColWidth: (g, w) => tableEdit((tb, r, c) => { tb.grid = tb.grid.slice(); tb.grid[g] = Math.max(12, Number(w)); return [r, c]; }, 'colw'),
  setRowHeight: (h, rule = 'atLeast') => tableEdit((tb, r, c, cs) => { const r1 = cs ? cs.r1 : r; const r2 = cs ? cs.r2 : r; for (let k = r1; k <= r2; k++) tb.rows[k] = { ...tb.rows[k], h: Number(h) || undefined, hRule: rule }; return [r, c]; }),
  tableStyle: (id) => tableEdit((tb, r, c) => { tb.style = id; return [r, c]; }),
  tableLook: (key, v) => tableEdit((tb, r, c) => { tb.tblPr = { ...tb.tblPr, look: { ...(tb.tblPr.look ?? {}), [key]: v ?? !(tb.tblPr.look ?? {})[key] } }; return [r, c]; }),
  repeatHeader: () => tableEdit((tb, r, c) => { tb.rows[r] = { ...tb.rows[r], header: !tb.rows[r].header }; return [r, c]; }),
  cellShading: (color) => cellsEdit((cl) => ({ ...cl, shd: color === 'none' ? undefined : color })),
  cellAlign: (v, h) => {
    cellsEdit((cl) => ({ ...cl, vAlign: v ?? cl.vAlign }));
    if (h) run(h === 'center' ? 'alignCenter' : h === 'right' ? 'alignRight' : 'alignLeft');
  },
  cellBorders: (kind, b = { style: 'single', w: 0.5, color: '#000000' }) => cellsEdit((cl, r, c, cs) => {
    const bd = { ...(cl.borders ?? {}) };
    const set = (k, v) => { if (v) bd[k] = { ...v }; else bd[k] = { style: 'none' }; };
    const top = r === cs.r1; const bottom = r === cs.r2; const left = c === cs.c1; const right = c === cs.c2;
    switch (kind) {
      case 'all': set('top', b); set('bottom', b); set('left', b); set('right', b); break;
      case 'none': set('top', null); set('bottom', null); set('left', null); set('right', null); break;
      case 'outside': if (top) set('top', b); if (bottom) set('bottom', b); if (left) set('left', b); if (right) set('right', b); break;
      case 'inside': if (!top) set('top', b); if (!bottom) set('bottom', b); if (!left) set('left', b); if (!right) set('right', b); break;
      case 'top': if (top) set('top', b); break;
      case 'bottom': if (bottom) set('bottom', b); break;
      case 'left': if (left) set('left', b); break;
      case 'right': if (right) set('right', b); break;
      case 'insideH': if (!top) set('top', b); if (!bottom) set('bottom', b); break;
      case 'insideV': if (!left) set('left', b); if (!right) set('right', b); break;
      default: break;
    }
    return { ...cl, borders: bd };
  }),
  tableAlign: (al) => tableEdit((tb, r, c) => { tb.tblPr = { ...tb.tblPr, align: al }; return [r, c]; }),
  selectTable: () => { const tc = tableCtx(); if (!tc) return; const { t, tpath } = tc; setSel(cellPos(tpath, 0, 0), cellPos(tpath, t.rows.length - 1, t.rows[t.rows.length - 1].cells.length - 1, true)); },
  selectRow: () => { const tc = tableCtx(); if (!tc) return; const { t, tpath, r } = tc; setSel(cellPos(tpath, r, 0), cellPos(tpath, r, t.rows[r].cells.length - 1, true)); },
  selectCol: () => { const tc = tableCtx(); if (!tc) return; const { t, tpath, c } = tc; setSel(cellPos(tpath, 0, Math.min(c, t.rows[0].cells.length - 1)), cellPos(tpath, t.rows.length - 1, Math.min(c, t.rows[t.rows.length - 1].cells.length - 1), true)); },
  selectCell: () => { const tc = tableCtx(); if (!tc) return; setSel(cellPos(tc.tpath, tc.r, tc.c), cellPos(tc.tpath, tc.r, tc.c, true)); },
  textToTable: (sepChar = '\t') => {
    const paths = selParas().filter((pt) => pt.length === 1);
    if (!paths.length) return;
    const rows = paths.map((pt) => paraPlain(blockAt(S.doc, pt)).split(sepChar === 'tab' ? '\t' : sepChar));
    const cols = Math.max(...rows.map((r) => r.length));
    const w = bodyWidthAt({ p: paths[0], o: 0 });
    const t = table(rows.length, cols, { width: w });
    t.rows.forEach((row, r) => row.cells.forEach((c, k) => { c.blocks = [para(rows[r][k] ?? '')]; }));
    change((d) => { const lst = d.list(paths[0]); lst.splice(paths[0][0], paths.length, t); d.owned.add(t); return { p: [paths[0][0], 0, 0, 0], o: 0 }; });
  },
  tableToText: () => { const tc = tableCtx(); if (!tc) return; const { t, tpath } = tc; change((d) => { const lst = d.list(tpath); const ps = tableToParasLocal(t); lst.splice(tpath[tpath.length - 1], 1, ...ps); return { p: tpath, o: 0 }; }); },
  sortTable: (col = 0, desc = false, header = true) => tableEdit((tb, r, c) => {
    const start = header && tb.rows[0]?.header !== false ? 1 : 0;
    const key = (row) => row.cells[col] ? row.cells[col].blocks.map((b) => (b.type === 'p' ? paraPlain(b) : '')).join(' ') : '';
    const head = tb.rows.slice(0, start);
    const body = tb.rows.slice(start).sort((x, y) => { const a = key(x); const b = key(y); const na = parseFloat(a.replace(/,/g, '')); const nb = parseFloat(b.replace(/,/g, '')); const v = Number.isFinite(na) && Number.isFinite(nb) ? na - nb : a.localeCompare(b, 'ko'); return desc ? -v : v; });
    tb.rows = [...head, ...body];
    return [r, c];
  }),
  tableFormula: (formula) => {
    // =SUM(ABOVE) / =SUM(LEFT) / =AVERAGE(ABOVE) …
    const tc = tableCtx(); if (!tc) return;
    const { t, r, c } = tc;
    const m = /^=?\s*(SUM|AVERAGE|MAX|MIN|COUNT|PRODUCT)\((ABOVE|LEFT|BELOW|RIGHT)\)/i.exec(formula ?? '=SUM(ABOVE)');
    if (!m) { emit('toast', '지원하는 수식: SUM · AVERAGE · MAX · MIN · COUNT · PRODUCT (ABOVE · LEFT · BELOW · RIGHT)'); return; }
    const vals = [];
    const num = (cl) => parseFloat((cl?.blocks ?? []).map((b) => (b.type === 'p' ? paraPlain(b) : '')).join('').replace(/[,\s원₩$%]/g, ''));
    const dir = m[2].toUpperCase();
    if (dir === 'ABOVE') for (let k = r - 1; k >= 0; k--) { const v = num(t.rows[k].cells[c]); if (Number.isFinite(v)) vals.push(v); else if (vals.length) break; }
    if (dir === 'BELOW') for (let k = r + 1; k < t.rows.length; k++) { const v = num(t.rows[k].cells[c]); if (Number.isFinite(v)) vals.push(v); }
    if (dir === 'LEFT') for (let k = c - 1; k >= 0; k--) { const v = num(t.rows[r].cells[k]); if (Number.isFinite(v)) vals.push(v); else if (vals.length) break; }
    if (dir === 'RIGHT') for (let k = c + 1; k < t.rows[r].cells.length; k++) { const v = num(t.rows[r].cells[k]); if (Number.isFinite(v)) vals.push(v); }
    const f = m[1].toUpperCase();
    const res = f === 'SUM' ? vals.reduce((a, b) => a + b, 0) : f === 'AVERAGE' ? (vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0) : f === 'MAX' ? Math.max(...vals) : f === 'MIN' ? Math.min(...vals) : f === 'COUNT' ? vals.length : vals.reduce((a, b) => a * b, 1);
    insertObject({ obj: 'field', code: `=${f}(${dir})`, result: Number.isInteger(res) ? res.toLocaleString('ko-KR') : res.toLocaleString('ko-KR', { maximumFractionDigits: 2 }) });
  },
}, {
  insertRowBelow: { label: '아래에 행 삽입' }, insertRowAbove: { label: '위에 행 삽입' }, insertColRight: { label: '오른쪽에 열 삽입' }, insertColLeft: { label: '왼쪽에 열 삽입' },
  deleteRow: { label: '행 삭제' }, deleteCol: { label: '열 삭제' }, deleteTable: { label: '표 삭제' }, mergeCells: { label: '셀 병합' }, splitCell: { label: '셀 분할', args: ['cols', 'rows'] },
  tableStyle: { label: '표 스타일', args: ['TableGrid|GridTable4-Accent1|ListTable3-Accent2|PlainTable1 …'] }, cellShading: { label: '셀 음영', args: ['#RRGGBB|none'] }, sortTable: { label: '표 정렬', args: ['열 번호', 'desc?'] }, tableFormula: { label: '표 수식', args: ['=SUM(ABOVE)'] },
});
function tableToParasLocal(t) {
  return t.rows.map((r) => { const runs = []; r.cells.forEach((c, i) => { if (i) runs.push({ obj: 'tab' }); runs.push({ text: c.blocks.filter((b) => b.type === 'p').map(paraPlain).join(' ') }); }); return { ...para(''), runs: normalizeRuns(runs) }; });
}
function gridCol(tb, r, c, after) { let gc = 0; for (let k = 0; k < c; k++) gc += tb.rows[r].cells[k].span ?? 1; return after ? gc + (tb.rows[r].cells[c].span ?? 1) : gc; }
function insertGridCol(tb, g) {
  const w = tb.grid[Math.min(g, tb.grid.length - 1)] ?? 60;
  // 전체 너비 유지: 다른 열을 비례로 줄임
  const total = tb.grid.reduce((a, b) => a + b, 0);
  tb.grid = [...tb.grid.slice(0, g), w, ...tb.grid.slice(g)].map((x) => (x * total) / (total + w));
  tb.rows = tb.rows.map((row) => {
    let gc = 0; let done = false;
    const cells = [];
    for (const cl of row.cells) {
      const s = cl.span ?? 1;
      if (!done && gc === g) { cells.push(emptyCell(cl)); done = true; }
      if (!done && gc < g && g < gc + s) { cells.push({ ...cl, span: s + 1 }); done = true; gc += s; continue; }
      cells.push(cl);
      gc += s;
    }
    if (!done) cells.push(emptyCell(row.cells[row.cells.length - 1]));
    return { ...row, cells };
  });
}
function deleteGridCol(tb, g) {
  const total = tb.grid.reduce((a, b) => a + b, 0);
  tb.grid = tb.grid.filter((_, i) => i !== g);
  const rest = tb.grid.reduce((a, b) => a + b, 0) || 1;
  tb.grid = tb.grid.map((x) => (x * total) / rest);
  tb.rows = tb.rows.map((row) => {
    let gc = 0;
    const cells = [];
    for (const cl of row.cells) { const s = cl.span ?? 1; if (gc <= g && g < gc + s) { if (s > 1) cells.push({ ...cl, span: s - 1 }); } else cells.push(cl); gc += s; }
    return { ...row, cells };
  });
}
function tableEdit(fn, key = null) {
  const tc = tableCtx();
  if (!tc) { emit('toast', '표 안을 클릭한 뒤 실행하세요'); return; }
  const { tpath, r, c, cs } = tc;
  change((d) => {
    const tb = d.block(tpath);
    tb.rows = tb.rows.slice();
    const res = fn(tb, r, c, cs);
    if (res === false) return false;
    if (res === 'remove') { const lst = d.list(tpath); lst.splice(tpath[tpath.length - 1], 1); if (!lst.length) lst.push(para('')); const i = Math.min(tpath[tpath.length - 1], lst.length - 1); return lst[i].type === 'p' ? { p: [...tpath.slice(0, -1), i], o: 0 } : firstPos(lst); }
    const [nr, nc] = res;
    const rr = Math.min(nr, tb.rows.length - 1);
    const cc = Math.min(nc, tb.rows[rr].cells.length - 1);
    return firstPos(tb.rows[rr].cells[cc].blocks, [...tpath, rr, cc]);
  }, { key });
}
function cellsEdit(fn) {
  const tc = tableCtx();
  if (!tc) return;
  const cs = tc.cs ?? { r1: tc.r, r2: tc.r, c1: tc.c, c2: tc.c };
  const keep = S.sel;
  change((d) => {
    const tb = d.block(tc.tpath);
    tb.rows = tb.rows.slice();
    for (let r = cs.r1; r <= cs.r2; r++) {
      const row = { ...tb.rows[r], cells: tb.rows[r].cells.slice() };
      for (let c = cs.c1; c <= Math.min(cs.c2, row.cells.length - 1); c++) row.cells[c] = fn(row.cells[c], r, c, cs);
      tb.rows[r] = row;
    }
    return keep;
  });
}

// ───────────── 레이아웃 (쪽 설정) ─────────────
function editSection(fn) {
  const i = sel().f.p[0];
  const keep = S.sel;
  change((d) => {
    // 이 블록이 속한 구역: 구역 나누기 문단의 sec 또는 마지막 구역
    for (const s of sections(d.doc)) {
      if (i >= s.start && i < s.end) {
        if (s.end - 1 < d.doc.body.length && d.doc.body[s.end - 1]?.secBreak && s.end - 1 !== d.doc.body.length - 1 + 1) {
          const last = s.end - 1;
          if (d.doc.body[last]?.secBreak) { const b = d.block([last]); b.secBreak = { ...b.secBreak, sec: fn(clone(b.secBreak.sec)) }; return keep; }
        }
        d.doc.section = fn(clone(d.doc.section));
        return keep;
      }
    }
    d.doc.section = fn(clone(d.doc.section));
    return keep;
  });
}
/** 모든 구역에 */
function editAllSections(fn) {
  const keep = S.sel;
  change((d) => {
    d.doc.body.forEach((b, i) => { if (b.type === 'p' && b.secBreak) { const x = d.block([i]); x.secBreak = { ...x.secBreak, sec: fn(clone(x.secBreak.sec)) }; } });
    d.doc.section = fn(clone(d.doc.section));
    return keep;
  });
}
register({
  margins: (key) => { const m = typeof key === 'object' ? key : MARGINS[key]; if (!m) return; editSection((s) => ({ ...s, margin: { ...s.margin, t: m.t, b: m.b, l: m.l, r: m.r, ...(m.header ? { header: m.header } : {}), ...(m.footer ? { footer: m.footer } : {}) }, mirror: !!m.mirror })); },
  orientation: (o) => editSection((s) => { const land = o === 'landscape'; const w = Math.min(s.page.w, s.page.h); const h = Math.max(s.page.w, s.page.h); return { ...s, page: { ...s.page, orient: o, w: land ? h : w, h: land ? w : h } }; }),
  paperSize: (key, w, h) => editSection((s) => { const p = PAPERS[key] ?? { w: Number(w), h: Number(h) }; const land = s.page.orient === 'landscape'; return { ...s, page: { ...s.page, w: land ? Math.max(p.w, p.h) : Math.min(p.w, p.h), h: land ? Math.min(p.w, p.h) : Math.max(p.w, p.h), paper: key } }; }),
  columns: (n, space = 48, sep = false) => editSection((s) => ({ ...s, cols: { n: Math.max(1, Math.min(13, Number(n))), space: Number(space), sep: !!sep } })),
  columnsPreset: (k) => { const P = { one: [1], two: [2], three: [3], left: [2, [0.33, 0.67]], right: [2, [0.67, 0.33]] }[k]; if (!P) return; editSection((s) => { const W = s.page.w - s.margin.l - s.margin.r - 48 * (P[0] - 1); return { ...s, cols: { n: P[0], space: 48, sep: false, ...(P[1] ? { widths: P[1].map((f) => f * W) } : {}) } }; }); },
  pageNumberFormat: (fmt, start) => editSection((s) => ({ ...s, pgNum: { fmt: fmt ?? 'decimal', start: start == null || start === '' ? undefined : Number(start) } })),
  differentFirstPage: () => editSection((s) => ({ ...s, titlePg: !s.titlePg })),
  differentOddEven: () => change((d) => { d.doc.settings = { ...d.doc.settings, evenOdd: !d.doc.settings.evenOdd }; return S.sel; }),
  pageColor: (c) => change((d) => { d.doc.settings = { ...d.doc.settings, pageColor: c === 'none' ? undefined : c }; return S.sel; }),
  watermark: (text, o = {}) => change((d) => { d.doc.settings = { ...d.doc.settings, watermark: text ? { text, color: o.color ?? '#D9D9D9', size: o.size ?? 0, diagonal: o.diagonal !== false } : undefined }; return S.sel; }),
  pageBorders: (b) => editAllSections((s) => ({ ...s, pageBorder: b === 'none' ? undefined : b ?? { style: 'single', w: 1, color: '#000000', space: 24 } })),
  vAlign: (v) => editSection((s) => ({ ...s, vAlign: v })),
  lineNumbers: (mode) => editSection((s) => ({ ...s, lineNum: mode === 'none' ? undefined : { restart: mode ?? 'page', countBy: 1 } })),
  pageSetup: (patch) => editSection((s) => ({ ...s, ...patch, page: { ...s.page, ...(patch.page ?? {}) }, margin: { ...s.margin, ...(patch.margin ?? {}) } })),
  setHeader: (blocks, which = 'default', kind = 'headers') => editSection((s) => ({ ...s, [kind]: { ...s[kind], [which]: blocks } })),
  removeHeader: (kind = 'headers') => editSection((s) => ({ ...s, [kind]: {} })),
}, {
  margins: { label: '여백', args: ['normal|narrow|moderate|wide|mirrored|hwp'] }, orientation: { label: '용지 방향', args: ['portrait|landscape'] }, paperSize: { label: '크기', args: ['A4|A3|A5|B4|B5|Letter|Legal'] },
  columns: { label: '단', args: ['n', 'space px?', 'line?'] }, pageNumberFormat: { label: '쪽 번호 서식', args: ['decimal|upperRoman|lowerRoman|dash', 'start?'] }, watermark: { label: '워터마크', args: ['글 (빈 글이면 제거)'] }, pageColor: { label: '페이지 색', args: ['#RRGGBB|none'] },
});

// ───────────── 참조 ─────────────
register({
  insertToc: (levels = 3) => {
    const pos = collapsed() ? sel().f : orderPos(sel().a, sel().f)[0];
    if (pos.p.length > 1) { emit('toast', '표 밖에서 목차를 넣으세요'); return; }
    const pageOf = (id) => pageOfBlock(id);
    const blocks = buildToc(S.doc, { levels: Number(levels), pageOf });
    change((d) => insertBlocks(d, pos, [...blocks, para('')]));
    setTimeout(() => run('updateToc'), 50);
  },
  updateToc: () => {
    const ids = new Set();
    for (const b of S.doc.body) if (b.toc) ids.add(b.toc);
    if (!ids.size) { emit('toast', '문서에 목차가 없습니다 (참조 › 목차)'); return; }
    // 두 번: 목차 길이가 바뀌면 쪽 번호가 바뀜
    for (let pass = 0; pass < 2; pass++) {
      change((d) => {
        for (const tocId of ids) {
          const idx = d.doc.body.findIndex((b) => b.toc === tocId);
          if (idx < 0) continue;
          let n = 0; while (d.doc.body[idx + n]?.toc === tocId) n++;
          const levels = Math.max(1, ...d.doc.body.slice(idx, idx + n).map((b) => Number(/TOC(\d)/.exec(b.style ?? '')?.[1] ?? 0)).filter(Boolean), 3);
          const hasTitle = d.doc.body[idx].style === 'TOCHeading';
          const blocks = buildToc(d.doc, { levels, pageOf: (id) => pageOfBlock(id), tocId, title: hasTitle ? paraPlain(d.doc.body[idx]) : '' });
          d.list([idx]).splice(idx, n, ...blocks);
        }
        return S.sel;
      }, { key: pass ? 'toc' : null });
      relayout();
    }
  },
  insertCaption: (label = '표', text = '') => {
    const pos = sel().f;
    let n = 1;
    walkParas(S.doc.body, (p) => { if (p.style === 'Caption' && paraPlain(p).startsWith(`${label} `)) n++; });
    change((d) => {
      const at = splitPara(d, { p: pos.p, o: paraLen(blockAt(d.doc, pos.p)) });
      const b = d.block(at.p);
      b.style = 'Caption';
      b.pPr = {};
      b.runs = [{ text: `${label} ` }, { obj: 'field', code: `SEQ ${label} \\* ARABIC`, result: String(n) }, { text: text ? `. ${text}` : '' }].filter((r) => r.obj || r.text);
      return { p: at.p, o: paraLen(b) };
    });
  },
  updateFields: () => change((d) => {
    const seq = {};
    walkParas(d.doc.body, (p, path) => {
      if (!p.runs.some((r) => r.obj === 'field')) return;
      const b = d.block(path);
      b.runs = b.runs.map((r) => {
        if (r.obj !== 'field') return r;
        const code = (r.code ?? '').trim();
        const head = code.split(/\s+/)[0].toUpperCase();
        if (head === 'SEQ') { const k = code.split(/\s+/)[1]; seq[k] = (seq[k] ?? 0) + 1; return { ...r, result: String(seq[k]) }; }
        if (head === 'DATE' || head === 'TIME') return { ...r, result: new Date().toLocaleDateString('ko-KR') };
        if (head === 'NUMWORDS') return { ...r, result: String(docStats(S.doc).words) };
        if (head === 'PAGEREF') { const id = code.split(/\s+/)[1]; const pg = pageOfBlock(id) ?? pageOfBookmark(id); return { ...r, result: pg != null ? String(pg) : r.result }; }
        if (head === 'TITLE') return { ...r, result: d.doc.props.title || S.docName };
        if (head === 'AUTHOR') return { ...r, result: d.doc.props.author || author() };
        return r;
      });
    });
    return S.sel;
  }),
}, { insertToc: { label: '목차 삽입', args: ['levels (1-9)'] }, updateToc: { label: '목차 업데이트' }, insertCaption: { label: '캡션 삽입', args: ['표|그림|수식', 'text?'] }, updateFields: { label: '필드 업데이트' } });
export function pageOfBlock(id) {
  const res = S.layout;
  if (!res) return null;
  const its = res.byId.get(id);
  if (!its?.length) return null;
  return res.pages[its[0].page]?.numText ?? its[0].page + 1;
}
function pageOfBookmark(name) {
  let found = null;
  S.doc.body.forEach((b) => { if (found == null && b.type === 'p' && b.runs.some((r) => r.obj === 'mark' && r.name === name)) found = b.id; });
  return found ? pageOfBlock(found) : null;
}

// ───────────── 검토 ─────────────
register({
  trackChanges: () => change((d) => { d.doc.settings = { ...d.doc.settings, track: !d.doc.settings.track }; return S.sel; }),
  showMarkup: (mode) => { S.markup = mode; emit('change', { scope: 'view' }); },
  acceptAll: () => acceptReject(true, null),
  rejectAll: () => acceptReject(false, null),
  acceptSel: () => acceptReject(true, orderPos(sel().a, sel().f)),
  rejectSel: () => acceptReject(false, orderPos(sel().a, sel().f)),
  nextChange: (dir = 1) => {
    const out = [];
    walkParas(S.doc.body, (p, path) => { let o = 0; for (const r of p.runs) { const n = runLen(r); if (r.ins || r.del) out.push({ p: path, s: o, e: o + n }); o += n; } });
    if (!out.length) { emit('toast', '변경 내용이 없습니다'); return; }
    const cur = sel().f;
    const cmp = (x) => { for (let i = 0; i < Math.min(x.p.length, cur.p.length); i++) if (x.p[i] !== cur.p[i]) return x.p[i] - cur.p[i]; return x.s - cur.o; };
    const hit = dir > 0 ? out.find((x) => cmp(x) > 0) ?? out[0] : [...out].reverse().find((x) => cmp(x) < 0) ?? out[out.length - 1];
    setSel({ p: hit.p, o: hit.s }, { p: hit.p, o: hit.e });
  },
  newComment: (text = '') => {
    const id = uid('c');
    const [a, b] = collapsed() ? (fmtRange() ?? [sel().f, sel().f]) : orderPos(sel().a, sel().f);
    change((d) => {
      insertRuns(d, b, [{ obj: 'mark', kind: 'cmt', id, edge: 'end' }]);
      insertRuns(d, a, [{ obj: 'mark', kind: 'cmt', id, edge: 'start' }]);
      d.doc.comments = [...(d.doc.comments ?? []), { id, author: author(), date: now(), text, replies: [], done: false }];
      return S.sel;
    });
    S.pane = 'comments';
    emit('pane');
    emit('editComment', id);
    return id;
  },
  editComment: (id, text) => change((d) => { d.doc.comments = d.doc.comments.map((c) => (c.id === id ? { ...c, text } : c)); return S.sel; }, { key: `cmt${id}` }),
  replyComment: (id, text) => change((d) => { d.doc.comments = d.doc.comments.map((c) => (c.id === id ? { ...c, replies: [...(c.replies ?? []), { author: author(), date: now(), text }] } : c)); return S.sel; }),
  resolveComment: (id) => change((d) => { d.doc.comments = d.doc.comments.map((c) => (c.id === id ? { ...c, done: !c.done } : c)); return S.sel; }),
  deleteComment: (id) => change((d) => {
    d.doc.comments = d.doc.comments.filter((c) => c.id !== id);
    walkParas(d.doc.body, (p, path) => { if (p.runs.some((r) => r.obj === 'mark' && r.kind === 'cmt' && r.id === id)) { const b = d.block(path); b.runs = b.runs.filter((r) => !(r.obj === 'mark' && r.kind === 'cmt' && r.id === id)); } });
    return S.sel;
  }),
  deleteAllComments: () => change((d) => { d.doc.comments = []; walkParas(d.doc.body, (p, path) => { if (p.runs.some((r) => r.obj === 'mark' && r.kind === 'cmt')) { const b = d.block(path); b.runs = b.runs.filter((r) => !(r.obj === 'mark' && r.kind === 'cmt')); } }); return S.sel; }),
  wordCount: () => emit('wordCount', docStats(S.doc)),
}, { trackChanges: { label: '변경 내용 추적 켜기/끄기' }, acceptAll: { label: '모든 변경 내용 적용' }, rejectAll: { label: '모든 변경 내용 취소' }, newComment: { label: '새 메모', args: ['text'] }, wordCount: { label: '단어 개수' } });
function acceptReject(accept, range) {
  const keep = S.sel;
  change((d) => {
    const inRange = (path, o) => { if (!range) return true; const [a, b] = range; const cmp = (x, y) => { for (let i = 0; i < Math.min(x.length, y.length); i++) if (x[i] !== y[i]) return x[i] - y[i]; return 0; }; const ca = cmp(path, a.p); const cb = cmp(path, b.p); if (ca < 0 || cb > 0) return false; if (ca === 0 && o < a.o) return false; if (cb === 0 && o >= b.o && !posEq(a, b)) return false; return true; };
    walkParas(d.doc.body, (p, path) => {
      if (!p.runs.some((r) => r.ins || r.del) && !p.paraIns) return;
      const b = d.block(path);
      let o = 0;
      const out = [];
      for (const r of b.runs) {
        const n = runLen(r);
        const hit = inRange(path, o);
        o += n;
        if (!hit || (!r.ins && !r.del)) { out.push(r); continue; }
        if (r.ins) { if (accept) { const c = { ...r }; delete c.ins; out.push(c); } }
        else if (r.del) { if (!accept) { const c = { ...r }; delete c.del; out.push(c); } }
      }
      b.runs = normalizeRuns(out);
      delete b.paraIns;
    });
    return keep;
  });
}

// ───────────── 찾기 · 바꾸기 ─────────────
register({
  find: (q, opts = {}) => {
    const hits = findAll(S.doc, q, opts);
    S.findHits = hits;
    S.findQuery = q;
    emit('findHits', hits);
    return hits.length;
  },
  findNext: (q, opts = {}, dir = 1) => {
    const hits = findAll(S.doc, q ?? S.findQuery, opts);
    if (!hits.length) { emit('toast', '검색 결과가 없습니다'); return false; }
    const cur = dir > 0 ? orderPos(sel().a, sel().f)[1] : orderPos(sel().a, sel().f)[0];
    const key = (h) => [...h.p, h.s];
    const cmpK = (x, y) => { for (let i = 0; i < Math.max(x.length, y.length); i++) { const a = x[i] ?? -1; const b = y[i] ?? -1; if (a !== b) return a - b; } return 0; };
    const ck = [...cur.p, cur.o];
    const h = dir > 0 ? hits.find((x) => cmpK(key(x), ck) >= 0 && !(pathEq(x.p, cur.p) && x.s === cur.o && x.e === cur.o)) ?? hits[0] : [...hits].reverse().find((x) => cmpK([...x.p, x.e], ck) < 0) ?? hits[hits.length - 1];
    setSel({ p: h.p, o: h.s }, { p: h.p, o: h.e });
    return true;
  },
  replace: (q, rep, opts = {}) => {
    const [a, b] = orderPos(sel().a, sel().f);
    const hits = findAll(S.doc, q, opts);
    const on = hits.find((h) => pathEq(h.p, a.p) && h.s === a.o && h.e === b.o && pathEq(h.p, b.p));
    if (on) {
      const text = opts.regex ? on.text.replace(new RegExp(q, opts.matchCase ? 'u' : 'iu'), rep) : rep;
      change((d) => { const props = propsAt(blockAt(d.doc, on.p), on.s + 1); deleteRange(d, { p: on.p, o: on.s }, { p: on.p, o: on.e }); return insertText(d, { p: on.p, o: on.s }, text, props); });
    }
    return run('findNext', q, opts, 1);
  },
  replaceAll: (q, rep, opts = {}) => {
    const hits = findAll(S.doc, q, opts);
    if (!hits.length) { emit('toast', '바꿀 항목이 없습니다'); return 0; }
    change((d) => {
      // 뒤에서부터 (위치가 밀리지 않게)
      for (const h of [...hits].reverse()) {
        const text = opts.regex ? h.text.replace(new RegExp(q, opts.matchCase ? 'u' : 'iu'), rep) : rep;
        const props = propsAt(blockAt(d.doc, h.p), h.s + 1);
        if (opts.format) Object.assign(props, opts.format);
        deleteRange(d, { p: h.p, o: h.s }, { p: h.p, o: h.e });
        insertText(d, { p: h.p, o: h.s }, text, props);
      }
      return S.sel;
    });
    emit('toast', `${hits.length}개 항목을 바꿨습니다`);
    return hits.length;
  },
  goToPage: (n) => { const i = Number(n) - 1; const pg = S.layout?.pages[i]; if (!pg) return; const it = pg.bands[0]?.cols[0]?.items[0]; if (it) { const path = V.pathById.get(it.block.id) ?? (it.block.type === 'tbl' ? null : null); if (path) setSel({ p: path, o: 0 }); else { const bi = S.doc.body.indexOf(it.block); if (bi >= 0) setSel(firstPos(S.doc.body.slice(bi), [bi]).p.length ? firstPos([it.block], [bi]) : { p: [bi], o: 0 }); } } V.scroller.scrollTop = V.holders[i]?.offsetTop - 10; },
  goToHeading: (id) => { const path = V.pathById.get(id); if (path) { setSel({ p: path, o: 0 }); const h = V.holders[S.layout.byId.get(S.doc.body[path[0]]?.id)?.[0]?.page ?? 0]; if (h) V.scroller.scrollTop = h.offsetTop + 0; scrollToCaret(); } },
}, { find: { label: '찾기', args: ['query', '{matchCase, wholeWord, regex}'] }, replace: { label: '바꾸기', args: ['query', 'replacement'] }, replaceAll: { label: '모두 바꾸기', args: ['query', 'replacement', '{matchCase, wholeWord, regex}'] }, goToPage: { label: '페이지로 이동', args: ['n'] } });

// ───────────── 보기 ─────────────
register({
  toggleMarks: () => { S.marks = !S.marks; emit('change', { scope: 'view' }); },
  toggleRuler: () => { S.showRuler = !S.showRuler; emit('ui'); },
  toggleNav: () => { S.showNav = !S.showNav; emit('ui'); },
  zoomIn: () => setZoom(Math.round((S.zoom + 0.1) * 10) / 10),
  zoomOut: () => setZoom(Math.round((S.zoom - 0.1) * 10) / 10),
  zoom100: () => setZoom(1),
  zoomTo: (z) => setZoom(Number(z) / 100),
  zoomPageWidth: () => { S.zoomMode = 'pageWidth'; applyZoomMode(); },
  zoomOnePage: () => { S.zoomMode = 'onePage'; applyZoomMode(); },
  zoomMulti: () => { S.zoomMode = 'multi'; applyZoomMode(); },
  viewPrint: () => { S.view = 'print'; emit('ui'); },
  viewWeb: () => { S.view = 'web'; emit('ui'); },
  viewRead: () => { S.view = 'read'; emit('ui'); },
  viewOutline: () => { S.view = 'outline'; emit('ui'); },
  viewDraft: () => { S.view = 'draft'; emit('ui'); },
  focusMode: () => { document.body.classList.toggle('focus-mode'); },
}, { zoomTo: { label: '확대/축소', args: ['percent'] }, toggleMarks: { label: '편집 기호 표시' }, toggleNav: { label: '탐색 창' } });

// ───────────── 서식 복사 ─────────────
register({
  formatPainter: (keep = false) => {
    if (S.painter) { S.painter = null; document.body.classList.remove('painting'); emit('ui'); return; }
    const pos = sel().f;
    const p = blockAt(S.doc, pos.p);
    S.painter = { run: propsAt(p, collapsed() ? pos.o : orderPos(sel().a, sel().f)[0].o + 1), style: p.style, pPr: clone(p.pPr), keep };
    document.body.classList.add('painting');
    emit('ui');
  },
  applyPainter: () => {
    const pt = S.painter;
    if (!pt) return;
    const [a, b] = orderPos(sel().a, sel().f);
    const props = {};
    for (const k of RUN_KEYS) props[k] = ['ins', 'del', 'link'].includes(k) ? undefined : pt.run[k];
    delete props.ins; delete props.del; delete props.link;
    const keep = S.sel;
    const wholePara = a.o === 0 && b.o >= paraLen(blockAt(S.doc, b.p));
    change((d) => { formatRange(d, a, b, props); if (wholePara) formatParas(d, rangeParas(d.doc, a, b), (p) => { p.style = pt.style; p.pPr = clone(pt.pPr); }); return keep; });
    if (!pt.keep) { S.painter = null; document.body.classList.remove('painting'); emit('ui'); }
  },
});

// ───────────── 문서 ─────────────
register({
  newDocument: (tplDoc) => { setDoc(tplDoc ?? newDoc(), '문서1'); },
  docProps: (patch) => change((d) => { d.doc.props = { ...d.doc.props, ...patch }; return S.sel; }),
  insertBlocksAt: (blocks) => { const pos = collapsed() ? sel().f : orderPos(sel().a, sel().f)[0]; change((d) => insertBlocks(d, pos, blocks)); },
  pasteBlocks: (blocks, media) => pasteBlocks(blocks, media),
  insertImage: (src, w = 300, h = 200, alt = '') => {
    let media = null;
    if (src.startsWith('data:')) media = uid('m');
    const run0 = { obj: 'img', id: uid('i'), media, src: media ? undefined : src, w, h, alt };
    change((d) => {
      if (media) { const m = d.prop('media'); m[media] = src; }
      const at = collapsed() ? sel().f : orderPos(sel().a, sel().f)[0];
      return insertRuns(d, at, [run0]);
    });
  },
  imageProps: (patch) => {
    const o = S.objSel;
    if (!o) { emit('toast', '그림을 먼저 선택하세요'); return; }
    change((d) => {
      const p = d.block(o.p);
      let pos = 0;
      p.runs = p.runs.map((r) => { const at = pos; pos += runLen(r); if (at !== o.o) return r; const c = { ...r, ...patch }; for (const [k, v] of Object.entries(patch)) if (v === undefined) delete c[k]; return c; });
      return { a: { p: o.p, o: o.o }, f: { p: o.p, o: o.o + 1 } };
    });
    S.objSel = o;
    drawSelection();
  },
}, { insertImage: { label: '그림 삽입', args: ['url|dataURL', 'w', 'h', 'alt'] } });

// 내부 도우미 (대화 상자 · 메뉴에서)
register({
  __applyRun: (props) => applyRun(props),
  selectedTextImpl: () => selectedText(),
  __objRun: () => { const o = S.objSel; if (!o) return null; const p = blockAt(S.doc, o.p); let pos = 0; for (const r of p.runs) { if (pos === o.o) return r; pos += runLen(r); } return null; },
  __cellsEditDir: () => cellsEdit((cl) => ({ ...cl, dir: cl.dir === 'vert' ? undefined : 'vert' })),
  __moveLeft: () => { const f = sel().f; if (f.o > 0) setSel({ p: f.p, o: f.o - 1 }); },
});
