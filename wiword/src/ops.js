// AI 연결 (MCP) 도구: 문서에 하는 일 (DOM 없음) — mcp.js (파일 · Node) 와 bridge.js (열린 WIWORD 창) 가 같이 씀
// 도구 = { name, description, input (JSON 스키마 properties), required, run(doc, args) → { doc?, result } }
import { draft, docText, paraPlain, walkParas, findAll, deleteRange, insertText, propsAt, blockAt, insertBlocks, lastPos, firstPos, headings, docStats, para, table, uid, PAPERS, MARGINS, normalizeRuns, runLen, sections, clone } from './model.js';
import { mdToDoc, docToMd } from './md.js';

const STYLE_ALIASES = { 제목1: 'Heading1', 제목2: 'Heading2', 제목3: 'Heading3', 제목: 'Title', 부제: 'Subtitle', 표준: 'Normal', 본문: 'Normal', 인용: 'Quote', 'heading 1': 'Heading1', 'heading 2': 'Heading2', 'heading 3': 'Heading3', h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', title: 'Title', normal: 'Normal', quote: 'Quote', list: 'ListParagraph' };
const styleId = (doc, s) => { if (!s) return null; if (doc.styles[s]) return s; const a = STYLE_ALIASES[String(s).toLowerCase().replace(/\s+/g, '')] ?? STYLE_ALIASES[String(s).toLowerCase()]; if (a) return a; const byName = Object.values(doc.styles).find((x) => x.name === s); return byName?.id ?? null; };
/** 최상위 문단 목록 (표 안 문단 제외) — 번호는 doc.body 인덱스 */
function paras(doc) { const out = []; doc.body.forEach((b, i) => { if (b.type === 'p') out.push({ index: i, style: b.style, text: paraPlain(b) }); else out.push({ index: i, type: 'table', rows: b.rows.length, cols: b.grid.length, text: b.rows.slice(0, 3).map((r) => r.cells.map((c) => c.blocks.map(paraPlain).join(' ')).join(' | ')).join(' / ') }); }); return out; }
const mdBlocks = (doc, md) => { const d = mdToDoc(md, { doc: { ...doc, body: [], lists: { ...doc.lists }, notes: { foot: { ...doc.notes?.foot }, end: { ...doc.notes?.end } }, media: { ...doc.media }, props: { ...doc.props } } }); return d; };
const now = () => new Date().toISOString();

export const TOOLS = [
  {
    name: 'get_text', description: '문서 전체 글을 일반 텍스트로 돌려줍니다 (문단마다 줄바꿈).',
    input: { max_chars: { type: 'number', description: '최대 글자 수 (기본 200000)' } },
    run: (doc, a) => ({ result: docText(doc).slice(0, a.max_chars ?? 200000) }),
  },
  {
    name: 'get_markdown', description: '문서를 마크다운으로 돌려줍니다 (제목 · 목록 · 표 · 굵게/기울임 · 링크 · 각주 유지). 문서 구조를 이해하거나 고칠 때 가장 좋습니다.',
    input: {},
    run: (doc) => ({ result: docToMd(doc, { images: 'none' }) }),
  },
  {
    name: 'set_markdown', description: '문서 본문 전체를 마크다운으로 바꿉니다 (용지 · 스타일 · 머리글은 유지). # 제목, - 목록, 1. 번호, | 표 |, **굵게**, > 인용, ``` 코드, [^1] 각주, $$수식$$ 을 지원합니다.',
    input: { markdown: { type: 'string', description: '새 본문 (마크다운)' } }, required: ['markdown'],
    run: (doc, a) => { const d = mdBlocks(doc, a.markdown); return { doc: { ...doc, body: d.body, lists: d.lists, styles: d.styles, notes: d.notes, media: d.media }, result: `본문을 ${d.body.length}개 블록으로 바꿨습니다` }; },
  },
  {
    name: 'insert_markdown', description: '마크다운 내용을 문서의 정해진 위치에 넣습니다. position: "end"(기본) | "start" | 블록 번호(list_paragraphs 의 index, 그 블록 뒤에 넣음).',
    input: { markdown: { type: 'string' }, position: { type: ['string', 'number'], description: '"end" | "start" | 블록 번호' } }, required: ['markdown'],
    run: (doc, a) => {
      const d = mdBlocks(doc, a.markdown);
      const body = doc.body.slice();
      const at = a.position === 'start' ? 0 : a.position == null || a.position === 'end' ? body.length : Math.max(0, Math.min(body.length, Number(a.position) + 1));
      // 문서가 빈 문단 하나뿐이면 바꿈
      if (body.length === 1 && body[0].type === 'p' && !body[0].runs.length) body.splice(0, 1, ...d.body);
      else body.splice(at, 0, ...d.body);
      return { doc: { ...doc, body, lists: d.lists, styles: d.styles, notes: d.notes, media: d.media }, result: `${d.body.length}개 블록을 넣었습니다` };
    },
  },
  {
    name: 'inspect', description: '문서 개요: 쪽 설정, 문단 · 표 · 그림 수, 제목 목록, 쓰는 스타일, 메모 · 변경 내용, 단어 수.',
    input: {},
    run: (doc) => {
      const st = docStats(doc);
      const used = {};
      let images = 0; let tables = 0; let tracked = 0;
      walkParas(doc.body, (p) => { used[p.style] = (used[p.style] ?? 0) + 1; for (const r of p.runs) { if (r.obj === 'img') images++; if (r.ins || r.del) tracked++; } });
      doc.body.forEach((b) => { if (b.type === 'tbl') tables++; });
      const sec = doc.section;
      const paper = Object.entries(PAPERS).find(([, p]) => Math.abs(p.w - Math.min(sec.page.w, sec.page.h)) < 2)?.[0] ?? 'custom';
      return { result: { title: doc.props?.title, author: doc.props?.author, paper, orientation: sec.page.orient, margins_cm: Object.fromEntries(['t', 'b', 'l', 'r'].map((k) => [k, Math.round((sec.margin[k] * 2.54 / 96) * 100) / 100])), sections: sections(doc).length, blocks: doc.body.length, tables, images, words: st.words, chars: st.chars, headings: headings(doc).map((h) => ({ level: h.level, text: h.text, index: h.path[0] })), styles_used: used, comments: (doc.comments ?? []).map((c) => ({ id: c.id, author: c.author, text: c.text, done: c.done })), tracked_changes: tracked, track_changes_on: !!doc.settings?.track, styles_available: Object.values(doc.styles).filter((s) => s.type === 'para' && !s.hidden).map((s) => `${s.id} (${s.name})`) } };
    },
  },
  {
    name: 'list_paragraphs', description: '블록(문단 · 표) 목록을 번호 · 스타일 · 글과 함께 돌려줍니다. 고칠 위치를 찾을 때 씁니다.',
    input: { from: { type: 'number' }, to: { type: 'number' } },
    run: (doc, a) => { const list = paras(doc); return { result: list.slice(a.from ?? 0, (a.to ?? list.length - 1) + 1).map((x) => ({ ...x, text: x.text.length > 400 ? `${x.text.slice(0, 400)}…` : x.text })) }; },
  },
  {
    name: 'edit_paragraph', description: '블록 번호(index)의 문단 글 · 스타일을 바꿉니다. text 를 주면 글을 바꾸고(첫 글자 서식 유지), markdown 을 주면 서식 있는 글로, style 을 주면 스타일만.',
    input: { index: { type: 'number' }, text: { type: 'string' }, markdown: { type: 'string' }, style: { type: 'string', description: 'Heading1 · Heading2 · Title · Normal · Quote · ListParagraph 또는 스타일 이름 (제목 1 …)' }, align: { type: 'string', enum: ['left', 'center', 'right', 'justify'] } }, required: ['index'],
    run: (doc, a) => {
      const b = doc.body[a.index];
      if (!b || b.type !== 'p') throw new Error(`${a.index}번 블록은 문단이 아닙니다`);
      const nb = { ...b, pPr: { ...b.pPr } };
      if (a.markdown != null) { const d = mdBlocks(doc, a.markdown); const p0 = d.body.find((x) => x.type === 'p'); if (p0) { nb.runs = p0.runs; if (p0.style !== 'Normal') nb.style = p0.style; } }
      else if (a.text != null) { const r0 = b.runs.find((r) => !r.obj) ?? {}; const { text: _t, ...props } = r0; nb.runs = a.text ? [{ ...props, text: a.text }] : []; }
      if (a.style) { const s = styleId(doc, a.style); if (!s) throw new Error(`스타일 '${a.style}' 이(가) 없습니다`); nb.style = s; }
      if (a.align) nb.pPr.align = a.align;
      const body = doc.body.slice(); body[a.index] = nb;
      return { doc: { ...doc, body }, result: 'OK' };
    },
  },
  {
    name: 'delete_blocks', description: '블록 번호 from ~ to (포함) 를 지웁니다.',
    input: { from: { type: 'number' }, to: { type: 'number' } }, required: ['from'],
    run: (doc, a) => { const body = doc.body.slice(); const n = (a.to ?? a.from) - a.from + 1; body.splice(a.from, n); if (!body.length) body.push(para('')); return { doc: { ...doc, body }, result: `${n}개 블록을 지웠습니다` }; },
  },
  {
    name: 'find', description: '글 찾기. 결과마다 블록 경로 · 앞뒤 글. regex=true 면 정규식.',
    input: { query: { type: 'string' }, regex: { type: 'boolean' }, match_case: { type: 'boolean' }, whole_word: { type: 'boolean' }, limit: { type: 'number' } }, required: ['query'],
    run: (doc, a) => {
      const hits = findAll(doc, a.query, { regex: !!a.regex, matchCase: !!a.match_case, wholeWord: !!a.whole_word });
      return { result: { count: hits.length, hits: hits.slice(0, a.limit ?? 50).map((h) => { const t = paraPlain(blockAt(doc, h.p)); return { path: h.p, block: h.p[0], match: h.text, context: `${t.slice(Math.max(0, h.s - 30), h.s)}【${h.text}】${t.slice(h.e, h.e + 30)}` }; }) } };
    },
  },
  {
    name: 'replace', description: '찾아 바꾸기 (글자 서식 유지). regex=true 면 $1 같은 그룹 참조 가능. track=true 면 변경 내용으로 (원문을 지운 표시 + 새 글 표시).',
    input: { find: { type: 'string' }, replace: { type: 'string' }, regex: { type: 'boolean' }, match_case: { type: 'boolean' }, whole_word: { type: 'boolean' }, track: { type: 'boolean' }, author: { type: 'string' } }, required: ['find', 'replace'],
    run: (doc, a) => {
      const opts = { regex: !!a.regex, matchCase: !!a.match_case, wholeWord: !!a.whole_word };
      const hits = findAll(doc, a.find, opts);
      if (!hits.length) return { result: '바꿀 항목이 없습니다' };
      const d = draft(doc);
      const track = a.track ?? !!doc.settings?.track;
      const info = { author: a.author ?? 'AI', date: now() };
      for (const h of [...hits].reverse()) {
        const text = opts.regex ? h.text.replace(new RegExp(a.find, opts.matchCase ? 'u' : 'iu'), a.replace) : a.replace;
        const props = propsAt(blockAt(d.doc, h.p), h.s + 1);
        if (track) {
          const b = d.block(h.p);
          // 지운 글 표시 + 새 글 삽입
          let o = 0; const out = [];
          for (const r of b.runs) { const n = runLen(r); if (r.obj || o + n <= h.s || o >= h.e) { out.push(r); o += n; continue; } const a0 = Math.max(0, h.s - o); const a1 = Math.min(n, h.e - o); if (a0) out.push({ ...r, text: r.text.slice(0, a0) }); out.push({ ...r, text: r.text.slice(a0, a1), del: info }); if (a1 < n) out.push({ ...r, text: r.text.slice(a1) }); o += n; }
          b.runs = normalizeRuns(out);
          insertText(d, { p: h.p, o: h.e }, text, { ...props, ins: info });
        } else {
          deleteRange(d, { p: h.p, o: h.s }, { p: h.p, o: h.e });
          insertText(d, { p: h.p, o: h.s }, text, props);
        }
      }
      return { doc: d.doc, result: `${hits.length}개 바꿈${track ? ' (변경 내용 추적)' : ''}` };
    },
  },
  {
    name: 'apply_style', description: '문단 스타일 적용: index(블록 번호) 또는 match(그 글이 들어 있는 모든 문단).',
    input: { style: { type: 'string' }, index: { type: 'number' }, match: { type: 'string' } }, required: ['style'],
    run: (doc, a) => {
      const s = styleId(doc, a.style);
      if (!s) throw new Error(`스타일 '${a.style}' 이(가) 없습니다. inspect 의 styles_available 을 보세요.`);
      let n = 0;
      const body = doc.body.map((b, i) => { if (b.type !== 'p') return b; if ((a.index != null && i === a.index) || (a.match && paraPlain(b).includes(a.match))) { n++; return { ...b, style: s }; } return b; });
      return { doc: { ...doc, body }, result: `${n}개 문단에 ${s}` };
    },
  },
  {
    name: 'format_text', description: '글자 서식: match 글의 모든 위치에 굵게 · 기울임 · 밑줄 · 색 · 형광펜 · 크기 · 글꼴.',
    input: { match: { type: 'string' }, bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, color: { type: 'string', description: '#RRGGBB' }, highlight: { type: 'string', description: 'yellow · brightGreen · cyan …' }, size: { type: 'number', description: 'pt' }, font: { type: 'string' } }, required: ['match'],
    run: (doc, a) => {
      const hits = findAll(doc, a.match, { matchCase: true });
      const props = {};
      if (a.bold != null) props.b = a.bold; if (a.italic != null) props.i = a.italic; if (a.underline != null) props.u = a.underline;
      if (a.color) props.color = a.color; if (a.highlight) props.hl = a.highlight; if (a.size) props.sz = a.size; if (a.font) { props.font = a.font; props.fontEa = a.font; }
      const d = draft(doc);
      for (const h of [...hits].reverse()) {
        const b = d.block(h.p);
        let o = 0; const out = [];
        for (const r of b.runs) { const n = runLen(r); if (r.obj || o + n <= h.s || o >= h.e) { out.push(r); o += n; continue; } const a0 = Math.max(0, h.s - o); const a1 = Math.min(n, h.e - o); if (a0) out.push({ ...r, text: r.text.slice(0, a0) }); out.push({ ...r, ...props, text: r.text.slice(a0, a1) }); if (a1 < n) out.push({ ...r, text: r.text.slice(a1) }); o += n; }
        b.runs = normalizeRuns(out);
      }
      return { doc: d.doc, result: `${hits.length}곳 서식 적용` };
    },
  },
  {
    name: 'insert_table', description: '표 넣기. rows = [[머리글…], [값…] …]. position 은 insert_markdown 과 같음.',
    input: { rows: { type: 'array', items: { type: 'array', items: { type: 'string' } } }, position: { type: ['string', 'number'] }, header: { type: 'boolean', description: '첫 행을 머리글로 (기본 true)' }, style: { type: 'string', description: 'TableGrid · GridTable4-Accent1 · ListTable3-Accent2 …' } }, required: ['rows'],
    run: (doc, a) => {
      const R = a.rows.length; const C = Math.max(...a.rows.map((r) => r.length));
      const W = doc.section.page.w - doc.section.margin.l - doc.section.margin.r;
      const t = table(R, C, { width: W, style: a.style ?? (a.header === false ? 'TableGrid' : 'GridTable4-Accent1') });
      t.rows = a.rows.map((r, ri) => ({ ...(ri === 0 && a.header !== false ? { header: true } : {}), cells: Array.from({ length: C }, (_, ci) => ({ blocks: [{ ...para(String(r[ci] ?? '')), pPr: { spc: { after: 0, line: 1, rule: 'auto' } } }] })) }));
      const body = doc.body.slice();
      const at = a.position === 'start' ? 0 : a.position == null || a.position === 'end' ? body.length : Math.max(0, Math.min(body.length, Number(a.position) + 1));
      body.splice(at, 0, t, para(''));
      return { doc: { ...doc, body }, result: `${R}×${C} 표를 넣었습니다` };
    },
  },
  {
    name: 'add_comment', description: 'match 글(첫 위치)에 메모를 답니다.',
    input: { match: { type: 'string' }, text: { type: 'string' }, author: { type: 'string' } }, required: ['match', 'text'],
    run: (doc, a) => {
      const h = findAll(doc, a.match, { matchCase: true })[0];
      if (!h) throw new Error(`'${a.match}' 을(를) 찾지 못했습니다`);
      const id = uid('c');
      const d = draft(doc);
      const b = d.block(h.p);
      let o = 0; const out = [];
      for (const r of b.runs) { const n = runLen(r); if (r.obj || o + n <= h.s || o >= h.e) { out.push(r); o += n; continue; } const a0 = Math.max(0, h.s - o); const a1 = Math.min(n, h.e - o); if (a0) out.push({ ...r, text: r.text.slice(0, a0) }); if (o + a0 === h.s) out.push({ obj: 'mark', kind: 'cmt', id, edge: 'start' }); out.push({ ...r, text: r.text.slice(a0, a1) }); if (o + a1 === h.e) out.push({ obj: 'mark', kind: 'cmt', id, edge: 'end' }); if (a1 < n) out.push({ ...r, text: r.text.slice(a1) }); o += n; }
      b.runs = normalizeRuns(out);
      d.doc.comments = [...(d.doc.comments ?? []), { id, author: a.author ?? 'AI', date: now(), text: a.text, replies: [], done: false }];
      return { doc: d.doc, result: `메모 ${id}` };
    },
  },
  {
    name: 'comments', description: '메모 목록 · 회신 · 해결 · 삭제. action: list | reply | resolve | delete.',
    input: { action: { type: 'string', enum: ['list', 'reply', 'resolve', 'delete'] }, id: { type: 'string' }, text: { type: 'string' }, author: { type: 'string' } },
    run: (doc, a) => {
      const list = doc.comments ?? [];
      if (!a.action || a.action === 'list') return { result: list };
      const c = list.find((x) => x.id === a.id);
      if (!c) throw new Error(`메모 ${a.id} 가 없습니다`);
      if (a.action === 'reply') return { doc: { ...doc, comments: list.map((x) => (x === c ? { ...x, replies: [...(x.replies ?? []), { author: a.author ?? 'AI', date: now(), text: a.text ?? '' }] } : x)) }, result: 'OK' };
      if (a.action === 'resolve') return { doc: { ...doc, comments: list.map((x) => (x === c ? { ...x, done: true } : x)) }, result: 'OK' };
      const d = draft(doc);
      walkParas(d.doc.body, (p, path) => { if (p.runs.some((r) => r.obj === 'mark' && r.id === a.id)) { const b = d.block(path); b.runs = b.runs.filter((r) => !(r.obj === 'mark' && r.id === a.id)); } });
      d.doc.comments = list.filter((x) => x !== c);
      return { doc: d.doc, result: 'OK' };
    },
  },
  {
    name: 'track_changes', description: '변경 내용: on | off | accept_all | reject_all | list.',
    input: { action: { type: 'string', enum: ['on', 'off', 'accept_all', 'reject_all', 'list'] } }, required: ['action'],
    run: (doc, a) => {
      if (a.action === 'on' || a.action === 'off') return { doc: { ...doc, settings: { ...doc.settings, track: a.action === 'on' } }, result: `변경 내용 추적 ${a.action === 'on' ? '켬' : '끔'}` };
      if (a.action === 'list') { const out = []; walkParas(doc.body, (p, path) => { for (const r of p.runs) if (r.ins || r.del) out.push({ block: path[0], kind: r.ins ? 'insert' : 'delete', text: r.text ?? '', author: (r.ins ?? r.del).author }); }); return { result: out }; }
      const accept = a.action === 'accept_all';
      const d = draft(doc);
      let n = 0;
      walkParas(d.doc.body, (p, path) => {
        if (!p.runs.some((r) => r.ins || r.del)) return;
        const b = d.block(path);
        b.runs = normalizeRuns(b.runs.flatMap((r) => { if (r.ins) { n++; if (!accept) return []; const { ins: _i, ...x } = r; return [x]; } if (r.del) { n++; if (accept) return []; const { del: _d, ...x } = r; return [x]; } return [r]; }));
      });
      return { doc: d.doc, result: `${n}개 ${accept ? '적용' : '취소'}` };
    },
  },
  {
    name: 'page_setup', description: '쪽 설정: paper(A4 · Letter · B5 · A3 …), orientation(portrait | landscape), margins(normal · narrow · wide · hwp 또는 {t,b,l,r} cm), columns(단 수).',
    input: { paper: { type: 'string' }, orientation: { type: 'string', enum: ['portrait', 'landscape'] }, margins: { type: ['string', 'object'] }, columns: { type: 'number' } },
    run: (doc, a) => {
      const fix = (sec) => {
        const s = clone(sec);
        if (a.paper && PAPERS[a.paper]) { const p = PAPERS[a.paper]; const land = s.page.orient === 'landscape'; s.page = { ...s.page, w: land ? p.h : p.w, h: land ? p.w : p.h }; }
        if (a.orientation && a.orientation !== s.page.orient) s.page = { w: s.page.h, h: s.page.w, orient: a.orientation };
        if (a.margins) { const m = typeof a.margins === 'string' ? MARGINS[a.margins] : Object.fromEntries(Object.entries(a.margins).map(([k, v]) => [k, (v * 96) / 2.54])); if (m) s.margin = { ...s.margin, ...(m.t != null ? { t: m.t } : {}), ...(m.b != null ? { b: m.b } : {}), ...(m.l != null ? { l: m.l } : {}), ...(m.r != null ? { r: m.r } : {}) }; }
        if (a.columns) s.cols = { ...(s.cols ?? {}), n: Math.max(1, Math.min(6, a.columns)) };
        return s;
      };
      const body = doc.body.map((b) => (b.type === 'p' && b.secBreak ? { ...b, secBreak: { ...b.secBreak, sec: fix(b.secBreak.sec) } } : b));
      return { doc: { ...doc, body, section: fix(doc.section) }, result: 'OK' };
    },
  },
  {
    name: 'set_properties', description: '문서 정보: title · subject · author · keywords · comments.',
    input: { title: { type: 'string' }, subject: { type: 'string' }, author: { type: 'string' }, keywords: { type: 'string' }, comments: { type: 'string' } },
    run: (doc, a) => ({ doc: { ...doc, props: { ...doc.props, ...a } }, result: 'OK' }),
  },
  {
    name: 'header_footer', description: '머리글/바닥글 글 설정 (모든 쪽). kind: header | footer. text 안의 {PAGE} 는 쪽 번호, {NUMPAGES} 는 전체 쪽 수. align: left | center | right.',
    input: { kind: { type: 'string', enum: ['header', 'footer'] }, text: { type: 'string' }, align: { type: 'string', enum: ['left', 'center', 'right'] } }, required: ['kind', 'text'],
    run: (doc, a) => {
      const runs = a.text.split(/(\{PAGE\}|\{NUMPAGES\})/).filter(Boolean).map((t) => (t === '{PAGE}' ? { obj: 'field', code: 'PAGE', result: '1' } : t === '{NUMPAGES}' ? { obj: 'field', code: 'NUMPAGES', result: '1' } : { text: t }));
      const bl = [{ ...para(''), style: a.kind === 'header' ? 'Header' : 'Footer', pPr: { align: a.align ?? 'center' }, runs }];
      const prop = a.kind === 'header' ? 'headers' : 'footers';
      const fix = (sec) => ({ ...sec, [prop]: { ...(sec[prop] ?? {}), default: bl } });
      const body = doc.body.map((b) => (b.type === 'p' && b.secBreak ? { ...b, secBreak: { ...b.secBreak, sec: fix(b.secBreak.sec) } } : b));
      return { doc: { ...doc, body, section: fix(doc.section) }, result: 'OK' };
    },
  },
  {
    name: 'insert_toc', description: '맨 앞(또는 position)에 목차를 넣습니다 (제목 1~levels).',
    input: { levels: { type: 'number' }, position: { type: ['string', 'number'] } },
    run: (doc, a) => {
      const hs = headings(doc, a.levels ?? 3);
      const blocks = [{ ...para('목차'), style: 'TOCHeading' }, ...hs.map((h) => ({ ...para(''), style: `TOC${h.level}`, pPr: { tabs: [{ pos: 9999, align: 'right', leader: 'dot' }] }, runs: [{ text: h.text, link: { anchor: h.p.id } }, { obj: 'tab' }, { obj: 'field', code: `PAGEREF ${h.p.id}`, result: '' }] })), { ...para(''), runs: [{ obj: 'br', kind: 'page' }] }];
      const body = doc.body.slice();
      const at = a.position == null || a.position === 'start' ? 0 : a.position === 'end' ? body.length : Number(a.position) + 1;
      body.splice(at, 0, ...blocks);
      return { doc: { ...doc, body }, result: `목차 항목 ${hs.length}개` };
    },
  },
];

export const TOOL_MAP = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
export function toolSchema(t) { return { name: t.name, description: t.description, inputSchema: { type: 'object', properties: t.input ?? {}, ...(t.required ? { required: t.required } : {}) } }; }
export function runTool(doc, name, args = {}) {
  const t = TOOL_MAP[name];
  if (!t) throw new Error(`알 수 없는 도구: ${name}`);
  return t.run(doc, args ?? {});
}
