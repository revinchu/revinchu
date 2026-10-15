// 기능: 머리글/바닥글 편집 · 편지 병합 · 인용/참고 문헌 · 색인 · 그림 목차 · 맞춤법 · 소리 내어 읽기 · 받아쓰기 · 번역 · 접근성 · 비교 · 편집 제한 · 요약 · 메모 탐색
import { S, run, register, emit, change, curSel, setDoc } from './state.js';
import { el, openDialog, formDialog, toast } from './ui.js';
import {
  blockAt, paraLen, paraPlain, walkParas, para, uid, runLen, propsAt, insertRuns, orderPos, posEq, headings, effPara, effRun, firstPos, lastPos, storyBlocks, sectionObj, newDoc, clone, docText, splitRuns, normalizeRuns,
} from './model.js';
import { setSel } from './editor.js';
import { V, relayout, scrollToCaret } from './view.js';
import { pageOfBlock, sectionAtTop } from './commands.js';
import { parseCsv } from './csv.js';

const sel = () => curSel();

// ───────────── 머리글 · 바닥글 편집 ─────────────
register({
  editHeaderFooterAt: (kind = 'header', pageIdx = null) => {
    const res = S.layout;
    const pi = pageIdx ?? (() => { for (const [i, r] of V.rendered) if (r) return i; return 0; })();
    const page = res?.pages[pi] ?? res?.pages[0];
    const si = page?.si ?? 0;
    const which = page?.hfKind ?? 'default';
    const key = `${kind === 'header' ? 'h' : 'f'}:${si}:${which}`;
    if (!storyBlocks(S.doc, key)) {
      // 없으면 빈 머리글 (쪽마다 그리려면 문서에 넣어야 함)
      change((d) => { d.list([key, 0]); return S.sel; });
    }
    S.hf = { key, kind, page: pi };
    document.body.classList.add('hf-editing');
    const bs = storyBlocks(S.doc, key) ?? [];
    const end = lastPos(bs, [key]);
    S.sel = { a: end, f: end };
    emit('change', { scope: 'view' });
    emit('ui');
  },
  closeHeaderFooter: () => {
    if (!S.hf) return;
    S.hf = null;
    document.body.classList.remove('hf-editing');
    S.sel = { a: firstPos(S.doc.body), f: firstPos(S.doc.body) };
    emit('change', { scope: 'view' });
    emit('ui');
  },
  hfGotoHeader: () => run('editHeaderFooterAt', 'header', S.hf?.page),
  hfGotoFooter: () => run('editHeaderFooterAt', 'footer', S.hf?.page),
  toggleShowDocText: () => { S.opts.hideDocText = !S.opts.hideDocText; document.body.classList.toggle('hide-doc-text', S.opts.hideDocText); emit('ui'); },
});

// ───────────── 메모 탐색 ─────────────
function commentAnchors() {
  const out = [];
  walkParas(S.doc.body, (p, path) => { let o = 0; for (const r of p.runs) { if (r.obj === 'mark' && r.kind === 'cmt' && r.edge === 'start') out.push({ id: r.id, p: path, o }); o += runLen(r); } });
  return out;
}
register({
  nextComment: (dir = 1) => {
    const list = commentAnchors();
    if (!list.length) { toast('메모가 없습니다'); return; }
    const cur = sel().f;
    const k = (x) => [...x.p, x.o];
    const cmp = (a, b) => { for (let i = 0; i < Math.max(a.length, b.length); i++) { const d = (a[i] ?? -1) - (b[i] ?? -1); if (d) return d; } return 0; };
    const ck = [...cur.p, cur.o];
    const hit = dir > 0 ? list.find((x) => cmp(k(x), ck) > 0) ?? list[0] : [...list].reverse().find((x) => cmp(k(x), ck) < 0) ?? list[list.length - 1];
    setSel({ p: hit.p, o: hit.o });
    S.activeComment = hit.id;
    if (S.pane !== 'comments') { S.pane = 'comments'; emit('pane'); } else emit('pane');
  },
  prevComment: () => run('nextComment', -1),
  deleteCurrentComment: () => { const id = S.activeComment ?? commentAt(sel().f); if (id) run('deleteComment', id); else toast('메모 안에 커서를 두세요'); },
  commentsPane: () => run('togglePane', 'comments'),
  revisionsPane: () => run('togglePane', 'revisions'),
  showNotes: () => { const ids = Object.keys(S.doc.notes?.foot ?? {}); if (!ids.length) { toast('각주가 없습니다'); return; } run('gotoNote', 1); },
  gotoNote: (dir = 1) => {
    const list = [];
    walkParas(S.doc.body, (p, path) => { let o = 0; for (const r of p.runs) { if (r.obj === 'note') list.push({ p: path, o }); o += runLen(r); } });
    if (!list.length) { toast('각주가 없습니다'); return; }
    const cur = sel().f;
    const ck = [...cur.p, cur.o];
    const cmp = (a, b) => { for (let i = 0; i < Math.max(a.length, b.length); i++) { const d = (a[i] ?? -1) - (b[i] ?? -1); if (d) return d; } return 0; };
    const hit = dir > 0 ? list.find((x) => cmp([...x.p, x.o], ck) > 0) ?? list[0] : [...list].reverse().find((x) => cmp([...x.p, x.o], ck) < 0) ?? list[list.length - 1];
    setSel({ p: hit.p, o: hit.o }, { p: hit.p, o: hit.o + 1 });
  },
});
function commentAt(pos) {
  const p = blockAt(S.doc, pos.p);
  const open = new Set();
  let o = 0;
  for (const r of p.runs) { if (o > pos.o) break; if (r.obj === 'mark' && r.kind === 'cmt') { if (r.edge === 'start') open.add(r.id); else open.delete(r.id); } o += runLen(r); }
  return [...open].pop() ?? null;
}

// ───────────── 편지 병합 ─────────────
register({
  loadRecipients: () => {
    const inp = el('input', { type: 'file', accept: '.csv,.tsv,.txt' });
    inp.onchange = async () => {
      const f = inp.files?.[0];
      if (!f) return;
      const text = await readText(f);
      const rows = parseCsv(text, f.name.endsWith('.tsv') ? '\t' : null);
      if (rows.length < 2) { toast('목록이 비어 있습니다 (첫 행 = 필드 이름)'); return; }
      S.merge = { ...(S.merge ?? {}), cols: rows[0].map((x) => x.trim()), rows: rows.slice(1).filter((r) => r.some((x) => x.trim())), index: 0 };
      toast(`받는 사람 ${S.merge.rows.length}명 · 필드 ${S.merge.cols.length}개`);
      emit('ui');
    };
    inp.click();
  },
  editRecipients: (fresh = false) => {
    const m = fresh || !S.merge?.cols ? { cols: ['이름', '직함', '회사', '주소', '전자 메일'], rows: [['', '', '', '', '']] } : S.merge;
    const ta = el('textarea', { rows: 12, style: { width: '100%', fontFamily: 'Consolas, monospace' } }, [m.cols, ...m.rows].map((r) => r.join('\t')).join('\n'));
    openDialog({ title: '받는 사람 목록 (탭으로 구분 · 첫 줄 = 필드 이름)', width: 640, body: el('div', {}, el('div', { class: 'muted' }, 'Excel 에서 복사해 붙여넣을 수 있습니다.'), ta), buttons: [{ label: '확인', primary: true, action: () => { const rows = ta.value.split(/\r?\n/).filter((l) => l.trim()).map((l) => l.split('\t')); if (rows.length < 1) return; S.merge = { ...(S.merge ?? {}), cols: rows[0], rows: rows.slice(1), index: 0 }; emit('ui'); emit('change', { scope: 'view' }); } }, { label: '취소' }] });
  },
  insertMergeField: (name) => run('__insertRun', { obj: 'field', code: `MERGEFIELD ${name}`, result: `«${name}»` }),
  addressBlock: () => { const c = S.merge?.cols ?? []; const pick = (re) => c.find((x) => re.test(x)); const fields = [pick(/이름|name/i), pick(/회사|company/i), pick(/주소|address/i)].filter(Boolean); if (!fields.length) { toast('먼저 받는 사람 목록을 불러오세요'); return; } fields.forEach((f, i) => { if (i) run('lineBreak'); run('insertMergeField', f); }); },
  greetingLine: () => { const c = S.merge?.cols ?? []; const n = c.find((x) => /이름|name/i.test(x)); if (!n) { toast('이름 필드가 없습니다'); return; } run('insertMergeField', n); run('insertText', ' 님께,'); },
  highlightMergeFields: () => { S.opts.mergeShade = !S.opts.mergeShade; emit('change', { scope: 'view' }); },
  previewResults: () => { if (!S.merge?.rows?.length) { toast('받는 사람 목록이 없습니다'); return; } S.merge.preview = !S.merge.preview; emit('change', { scope: 'view' }); emit('ui'); },
  mergeGoto: (i) => { if (!S.merge?.rows?.length) return; S.merge.index = Math.max(0, Math.min(S.merge.rows.length - 1, i)); S.merge.preview = true; emit('change', { scope: 'view' }); emit('ui'); },
  mergeFirst: () => run('mergeGoto', 0),
  mergePrev: () => run('mergeGoto', (S.merge?.index ?? 0) - 1),
  mergeNext: () => run('mergeGoto', (S.merge?.index ?? 0) + 1),
  mergeLast: () => run('mergeGoto', (S.merge?.rows?.length ?? 1) - 1),
  finishMerge: (mode = 'doc') => {
    const m = S.merge;
    if (!m?.rows?.length) { toast('받는 사람 목록이 없습니다'); return; }
    const out = newDoc({ styles: S.doc.styles, lists: S.doc.lists, defaults: S.doc.defaults, themeFonts: S.doc.themeFonts, section: clone(S.doc.section), media: S.doc.media });
    out.body = [];
    m.rows.forEach((row, ri) => {
      const rec = Object.fromEntries(m.cols.map((c, i) => [c, row[i] ?? '']));
      const sub = (blocks) => blocks.map((b) => {
        if (b.type === 'tbl') return { ...clone(b), id: uid('t'), rows: b.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, blocks: sub(c.blocks) })) })) };
        return { ...clone(b), id: uid('p'), runs: b.runs.map((r) => (r.obj === 'field' && /^MERGEFIELD/i.test(r.code ?? '') ? { ...propsOnly(r), text: String(rec[(r.code ?? '').split(/\s+/)[1]?.replace(/"/g, '')] ?? '') } : clone(r))).filter((r) => r.obj || r.text) };
      });
      const blocks = sub(S.doc.body);
      if (ri) out.body.push({ ...para(''), runs: [{ obj: 'br', kind: 'page' }] });
      out.body.push(...blocks);
    });
    if (mode === 'print') { const keep = S.doc; S.doc = out; relayout(); run('print'); S.doc = keep; relayout(); return; }
    setDoc(out, `${S.docName} (병합)`);
    toast(`${m.rows.length}개 문서를 만들었습니다`);
  },
  envelopesDialog: () => formDialog('봉투', [{ name: 'to', label: '받는 사람 주소', type: 'textarea', value: '' }, { name: 'from', label: '보내는 사람 주소', type: 'textarea', value: '' }], (v) => {
    const d = newDoc();
    d.section = { ...d.section, page: { w: 9.5 * 96, h: 4.125 * 96, orient: 'landscape' }, margin: { t: 30, b: 30, l: 30, r: 30, header: 20, footer: 20, gutter: 0 } };
    d.body = [...v.from.split('\n').map((l) => para(l, { pPr: { spc: { after: 0 } } })), para('', { pPr: { spc: { before: 72 } } }), ...v.to.split('\n').map((l) => para(l, { pPr: { ind: { l: 4.2 * 96 }, spc: { after: 0 } } }))];
    setDoc(d, '봉투');
  }),
  labelsDialog: () => formDialog('레이블', [{ name: 'text', label: '주소 (모든 레이블에 같은 내용 · 병합 필드는 편지 병합으로)', type: 'textarea', value: '' }, { name: 'cols', label: '열', type: 'number', value: 3 }, { name: 'rows', label: '행', type: 'number', value: 8 }], (v) => {
    const cols = Number(v.cols) || 3; const rows = Number(v.rows) || 8;
    const d = newDoc();
    d.section = { ...d.section, margin: { ...d.section.margin, t: 40, b: 40, l: 30, r: 30 } };
    const W = d.section.page.w - 60;
    const t = { type: 'tbl', id: uid('t'), style: 'TableNormal', tblPr: {}, grid: Array(cols).fill(W / cols), rows: Array.from({ length: rows }, () => ({ h: (d.section.page.h - 80) / rows, hRule: 'exact', cells: Array.from({ length: cols }, () => ({ vAlign: 'center', blocks: (v.text || ' ').split('\n').map((l) => para(l, { pPr: { align: 'center', spc: { after: 0 } } })) })) })) };
    d.body = [t, para('')];
    setDoc(d, '레이블');
  }),
});
const propsOnly = (r) => { const o = { ...r }; delete o.obj; delete o.code; delete o.result; return o; };
async function readText(f) {
  const buf = new Uint8Array(await f.arrayBuffer());
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return new TextDecoder().decode(buf.subarray(3));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch { return new TextDecoder('euc-kr').decode(buf); }
}

// ───────────── 인용 · 참고 문헌 · 색인 · 그림 목차 ─────────────
function citeText(s, style) {
  const a = s.author || '저자 미상';
  switch (style) {
    case 'MLA': return `(${a.split(/[ ,]/)[0]})`;
    case 'IEEE': return `[${(S.doc.sources ?? []).indexOf(s) + 1}]`;
    case 'Chicago': return `(${a} ${s.year})`;
    default: return `(${a}, ${s.year || 'n.d.'})`;
  }
}
function bibEntry(s, style, i) {
  const a = s.author || '저자 미상';
  switch (style) {
    case 'MLA': return `${a}. ${s.title}. ${s.publisher ? `${s.publisher}, ` : ''}${s.year}.`;
    case 'IEEE': return `[${i + 1}] ${a}, "${s.title}," ${s.publisher ? `${s.publisher}, ` : ''}${s.year}.`;
    case 'Chicago': return `${a}. ${s.year}. ${s.title}. ${s.publisher ?? ''}.`;
    case 'KS (한국 논문)': return `${a}, 「${s.title}」, ${s.publisher ? `${s.publisher}, ` : ''}${s.year}.`;
    default: return `${a} (${s.year || 'n.d.'}). ${s.title}. ${s.publisher ?? ''}${s.url ? ` ${s.url}` : ''}`;
  }
}
register({
  sourceDialog: () => formDialog('원본 만들기', [{ name: 'type', label: '원본 종류', type: 'select', value: 'book', options: [{ value: 'book', label: '책' }, { value: 'article', label: '학술지 논문' }, { value: 'web', label: '웹 사이트' }, { value: 'report', label: '보고서' }] }, { name: 'author', label: '저자', value: '' }, { name: 'title', label: '제목', value: '' }, { name: 'year', label: '연도', value: String(new Date().getFullYear()) }, { name: 'publisher', label: '출판사 / 학술지', value: '' }, { name: 'url', label: 'URL', value: '' }], (v) => {
    const id = uid('src');
    change((d) => { d.doc.sources = [...(d.doc.sources ?? []), { id, ...v }]; return S.sel; });
    run('insertCitation', id);
  }),
  manageSources: () => {
    const list = S.doc.sources ?? [];
    openDialog({ title: '원본 관리', width: 520, body: el('div', { class: 'dlg-list' }, list.length ? list.map((s) => el('div', { class: 'dlg-item' }, `${s.author} (${s.year}) ${s.title}`, el('button', { type: 'button', onclick: () => run('insertCitation', s.id) }, '인용'), el('button', { type: 'button', onclick: () => change((d) => { d.doc.sources = d.doc.sources.filter((x) => x.id !== s.id); return S.sel; }) }, '삭제'))) : el('div', { class: 'muted' }, '원본이 없습니다. [인용 삽입 › 새 원본 추가]')), buttons: [{ label: '새로 만들기...', action: () => { setTimeout(() => run('sourceDialog'), 0); } }, { label: '닫기', primary: true }] });
  },
  insertCitation: (id) => { const s = (S.doc.sources ?? []).find((x) => x.id === id); if (!s) return; run('__insertRun', { obj: 'field', code: `CITATION ${id}`, result: citeText(s, S.doc.settings.citeStyle ?? 'APA') }); },
  insertBibliography: (title = '참고 문헌') => {
    const style = S.doc.settings.citeStyle ?? 'APA';
    const src = [...(S.doc.sources ?? [])].sort((a, b) => (style === 'IEEE' ? 0 : (a.author ?? '').localeCompare(b.author ?? '', 'ko')));
    if (!src.length) { toast('원본이 없습니다'); return; }
    run('insertBlocksAt', [para(title, { style: 'Heading1' }), ...src.map((s, i) => para(bibEntry(s, style, i), { style: 'Bibliography', pPr: { ind: { l: 32, hanging: 32 } } })), para('')]);
  },
  markEntryDialog: () => {
    const t = posEq(sel().a, sel().f) ? '' : run('selectedTextImpl');
    formDialog('색인 항목 표시', [{ name: 'main', label: '주 항목', value: t }, { name: 'all', label: '모두 표시 (같은 낱말 전부)', type: 'checkbox', value: false }], (v) => {
      if (!v.main) return;
      change((d) => { d.doc.indexEntries = [...new Set([...(d.doc.indexEntries ?? []), v.main])]; return S.sel; });
      toast(`색인 항목 '${v.main}' 을(를) 표시했습니다. [색인 삽입]으로 목록을 만드세요`);
    });
  },
  insertIndex: () => {
    const entries = S.doc.indexEntries ?? [];
    if (!entries.length) { toast('색인 항목이 없습니다 ([항목 표시])'); return; }
    const rows = [];
    for (const e of [...entries].sort((a, b) => a.localeCompare(b, 'ko'))) {
      const pages = new Set();
      walkParas(S.doc.body, (p, path) => { if (paraPlain(p).includes(e)) { const pg = pageOfBlock(S.doc.body[path[0]].id); if (pg != null) pages.add(pg); } });
      rows.push(para([{ text: e }, { obj: 'tab' }, { text: [...pages].join(', ') }].length ? '' : '', { pPr: { tabs: [{ pos: 9999, align: 'right', leader: 'dot' }] } }));
      rows[rows.length - 1].runs = [{ text: e }, { obj: 'tab' }, { text: [...pages].join(', ') }];
    }
    run('insertBlocksAt', [para('색인', { style: 'Heading1' }), ...rows, para('')]);
  },
  tableOfFigures: () => {
    const caps = [];
    walkParas(S.doc.body, (p, path) => { if (p.style === 'Caption') caps.push({ text: paraPlain(p), id: p.id, top: S.doc.body[path[0]].id }); });
    if (!caps.length) { toast('캡션이 없습니다 ([캡션 삽입])'); return; }
    run('insertBlocksAt', [...caps.map((c) => ({ ...para(''), pPr: { tabs: [{ pos: 9999, align: 'right', leader: 'dot' }] }, runs: [{ text: c.text, link: { anchor: c.id } }, { obj: 'tab' }, { obj: 'field', code: `PAGEREF ${c.id}`, result: String(pageOfBlock(c.top) ?? '') }] })), para('')]);
  },
});

// ───────────── 맞춤법 (규칙 기반 한국어 · 영어 반복어) ─────────────
const KO_RULES = [
  [/몇일/g, '며칠'], [/어떻해/g, '어떡해'], [/할께/g, '할게'], [/할껄/g, '할걸'], [/금새/g, '금세'], [/웬지/g, '왠지'], [/왠만/g, '웬만'], [/되요/g, '돼요'], [/되서/g, '돼서'], [/안되/g, '안 되'], [/않되/g, '안 되'], [/됬/g, '됐'], [/역활/g, '역할'], [/설겆이/g, '설거지'], [/희안/g, '희한'], [/어의없/g, '어이없'], [/일일히/g, '일일이'], [/깨끗히/g, '깨끗이'], [/곰곰히/g, '곰곰이'], [/틈틈히/g, '틈틈이'], [/오랫만/g, '오랜만'], [/어떻게든지/g, '어떻게든'], [/내노라/g, '내로라'], [/바램/g, '바람'], [/뵈요/g, '봬요'], [/돋보기/g, '돋보기'], [/구지/g, '굳이'], [/궂이/g, '굳이'], [/문안하/g, '무난하'], [/요번/g, '이번'], [/명예회손/g, '명예훼손'], [/폐쇠/g, '폐쇄'], [/제작년/g, '재작년'], [/결제를 올/g, '결재를 올'], [/짜집기/g, '짜깁기'], [/눈꼽/g, '눈곱'], [/할수있/g, '할 수 있'], [/할수없/g, '할 수 없'], [/것같/g, '것 같'], [/수있/g, '수 있'], [/수없/g, '수 없'], [/때문에서/g, '때문에'],
];
register({
  spellCheck: () => {
    const issues = [];
    walkParas(S.doc.body, (p, path) => {
      const t = paraPlain(p);
      for (const [re, rep] of KO_RULES) { re.lastIndex = 0; let m; while ((m = re.exec(t))) issues.push({ p: path, s: m.index, e: m.index + m[0].length, word: m[0], sug: m[0].replace(new RegExp(re.source), rep), why: '맞춤법' }); }
      const rr = /(?<![\p{L}\p{N}])([\p{L}\p{N}]+)\s+\1(?![\p{L}\p{N}])/gu;
      let m; while ((m = rr.exec(t))) issues.push({ p: path, s: m.index, e: m.index + m[0].length, word: m[0], sug: m[1], why: '반복된 낱말' });
      const sp = / {2,}/g; while ((m = sp.exec(t))) issues.push({ p: path, s: m.index, e: m.index + m[0].length, word: '(공백 여러 개)', sug: ' ', why: '띄어쓰기' });
    });
    S.spellIssues = issues;
    S.pane = 'spelling';
    emit('pane');
    if (!issues.length) toast('맞춤법 검사를 마쳤습니다. 문제가 없습니다.');
  },
  applySpell: (i) => {
    const it = S.spellIssues?.[i];
    if (!it) return;
    change((d) => { const b = d.block(it.p); const [x, rest] = splitRuns(b.runs, it.s); const [mid, y] = splitRuns(rest, it.e - it.s); const props = mid.find((r) => !r.obj) ?? {}; b.runs = normalizeRuns([...x, { ...props, text: it.sug }, ...y]); return { p: it.p, o: it.s + it.sug.length }; });
    const delta = it.sug.length - (it.e - it.s);
    S.spellIssues = S.spellIssues.filter((_, k) => k !== i).map((x) => (x.p.join() === it.p.join() && x.s > it.s ? { ...x, s: x.s + delta, e: x.e + delta } : x));
    emit('pane');
  },
});

// ───────────── 소리 내어 읽기 · 받아쓰기 · 번역 ─────────────
register({
  readAloud: () => {
    const ss = globalThis.speechSynthesis;
    if (!ss) { toast('이 브라우저는 소리 내어 읽기를 지원하지 않습니다'); return; }
    if (S.reading) { ss.cancel(); S.reading = false; emit('ui'); return; }
    const paths = [];
    const start = sel().f;
    let on = false;
    walkParas(S.doc.body, (p, path) => { if (path.join() === start.p.join()) on = true; if (on && paraPlain(p).trim()) paths.push(path); });
    if (!paths.length) return;
    S.reading = true; emit('ui');
    let k = 0;
    const next = () => {
      if (!S.reading || k >= paths.length) { S.reading = false; emit('ui'); return; }
      const path = paths[k++];
      const p = blockAt(S.doc, path);
      setSel({ p: path, o: 0 }, { p: path, o: paraLen(p) });
      const u = new SpeechSynthesisUtterance(paraPlain(p));
      u.lang = /[가-힣]/.test(paraPlain(p)) ? 'ko-KR' : 'en-US';
      u.rate = S.opts.readRate ?? 1;
      u.onend = next; u.onerror = () => { S.reading = false; emit('ui'); };
      ss.speak(u);
    };
    next();
  },
  dictate: () => {
    const R = globalThis.SpeechRecognition || globalThis.webkitSpeechRecognition;
    if (!R) { toast('이 브라우저는 받아쓰기(음성 인식)를 지원하지 않습니다 (Chrome · Edge 에서 됩니다)'); return; }
    if (S.dictating) { S.dictating.stop(); S.dictating = null; emit('ui'); return; }
    const rec = new R();
    rec.lang = S.opts.dictateLang ?? 'ko-KR';
    rec.continuous = true;
    rec.interimResults = false;
    rec.onresult = (e) => { for (let i = e.resultIndex; i < e.results.length; i++) if (e.results[i].isFinal) { let t = e.results[i][0].transcript.trim(); t = t.replace(/\s*(마침표|온점)$/, '.').replace(/\s*쉼표$/, ',').replace(/\s*물음표$/, '?'); if (/새 ?줄$|줄 ?바꿈$/.test(t)) { run('newParagraph'); continue; } run('typeText', `${t} `); } };
    rec.onend = () => { if (S.dictating === rec) { try { rec.start(); } catch { S.dictating = null; emit('ui'); } } };
    rec.onerror = (e) => { toast(`받아쓰기 오류: ${e.error}`); S.dictating = null; emit('ui'); };
    rec.start();
    S.dictating = rec;
    toast('받아쓰기 중… 말하면 글로 입력됩니다 (다시 누르면 끝)');
    emit('ui');
  },
  translateDialog: async (scope = 'sel') => {
    const T = globalThis.Translator;
    const text = scope === 'sel' && !posEq(sel().a, sel().f) ? run('selectedTextImpl') : docText(S.doc);
    if (!text.trim()) { toast('번역할 글이 없습니다'); return; }
    const target = el('select', {}, [['en', '영어'], ['ko', '한국어'], ['ja', '일본어'], ['zh', '중국어'], ['es', '스페인어'], ['fr', '프랑스어'], ['de', '독일어'], ['vi', '베트남어']].map(([v, l]) => el('option', { value: v }, l)));
    const out = el('textarea', { rows: 10, style: { width: '100%' } });
    const go = async () => {
      if (!T) { out.value = '이 브라우저에는 내장 번역기(Translator API)가 없습니다. Chrome 최신 버전에서 쓸 수 있습니다. 또는 [AI 연결]로 Claude 에게 번역을 맡기세요.'; return; }
      try {
        const src = /[가-힣]/.test(text) ? 'ko' : 'en';
        const tr = await T.create({ sourceLanguage: src, targetLanguage: target.value });
        out.value = await tr.translate(text.slice(0, 20000));
      } catch (e) { out.value = `번역할 수 없습니다: ${e.message}`; }
    };
    openDialog({ title: '번역기', width: 560, body: el('div', {}, el('label', { class: 'dlg-row' }, el('span', {}, '대상 언어:'), target, el('button', { type: 'button', class: 'btn', onclick: go }, '번역')), out), buttons: [{ label: '삽입', primary: true, action: () => { if (out.value) run('insertText', out.value); } }, { label: '닫기' }] });
    go();
  },
});

// ───────────── 접근성 · 비교 · 편집 제한 · 요약 ─────────────
register({
  accessibilityCheck: () => {
    const issues = [];
    let lastLevel = 0;
    walkParas(S.doc.body, (p, path) => {
      for (const r of p.runs) if (r.obj === 'img' && !r.alt && !r.decorative) issues.push({ kind: '오류', msg: '대체 텍스트가 없는 그림', path });
      const lv = effPara(S.doc, p).outline;
      if (lv != null && lv < 9) { if (!paraPlain(p).trim()) issues.push({ kind: '경고', msg: '빈 제목', path }); if (lv + 1 > lastLevel + 1 && lastLevel) issues.push({ kind: '경고', msg: `제목 수준을 건너뜀 (제목 ${lastLevel} → 제목 ${lv + 1})`, path }); lastLevel = lv + 1; }
      for (const r of p.runs) if (!r.obj && r.color && /^#F{2}[EF].?/i.test(r.color) && !p.pPr.shd) { issues.push({ kind: '경고', msg: '읽기 어려운 글자 색 (흰색에 가까움)', path }); break; }
    });
    S.doc.body.forEach((b, i) => { if (b.type === 'tbl' && !b.rows[0]?.header) issues.push({ kind: '팁', msg: '표에 머리글 행이 지정되지 않음', path: [i] }); });
    if (!S.doc.props.title) issues.push({ kind: '팁', msg: '문서 제목(속성)이 없음 — 파일 › 정보에서 입력', path: null });
    S.a11yIssues = issues;
    S.pane = 'a11y';
    emit('pane');
  },
  compareDialog: () => {
    const inp = el('input', { type: 'file', accept: '.docx,.hwp,.hwpx,.md,.txt,.html,.wiword' });
    inp.onchange = async () => {
      const f = inp.files?.[0];
      if (!f) return;
      const other = await run('readAnyFile', f);
      if (!other) return;
      const out = compareDocs(other, S.doc);
      setDoc(out, `비교 결과 - ${S.docName}`);
      toast('원본(선택한 파일) → 현재 문서 차이를 변경 내용으로 표시했습니다');
    };
    toast('비교할 원본 문서를 고르세요 (현재 문서 = 수정본)');
    inp.click();
  },
  restrictEditing: () => formDialog('편집 제한', [{ name: 'mode', label: '편집 제한', type: 'select', value: S.doc.settings.protect ?? 'none', options: [{ value: 'none', label: '제한 없음' }, { value: 'readOnly', label: '변경할 수 없음 (읽기 전용)' }, { value: 'tracked', label: '변경 내용 추적만 허용' }, { value: 'comments', label: '메모만 허용' }] }], (v) => {
    change((d) => { d.doc.settings = { ...d.doc.settings, protect: v.mode === 'none' ? undefined : v.mode, track: v.mode === 'tracked' ? true : d.doc.settings.track }; return S.sel; });
    toast(v.mode === 'none' ? '편집 제한을 해제했습니다' : '편집 제한을 적용했습니다');
  }),
  aiSummary: () => {
    // 내장 요약 (추출식): 낱말 빈도로 중요한 문장을 고름. 생성형 요약은 [AI 연결] (MCP) 로
    const text = docText(S.doc);
    const sents = text.split(/(?<=[.!?。]|다\.)\s+|\n+/).map((s) => s.trim()).filter((s) => s.length > 12);
    if (!sents.length) { toast('요약할 글이 없습니다'); return; }
    const freq = new Map();
    for (const w of text.toLowerCase().match(/[\p{L}\p{N}]{2,}/gu) ?? []) freq.set(w, (freq.get(w) ?? 0) + 1);
    const score = (s) => (s.toLowerCase().match(/[\p{L}\p{N}]{2,}/gu) ?? []).reduce((a, w) => a + (freq.get(w) ?? 0), 0) / Math.sqrt(s.length);
    const n = Math.max(3, Math.min(7, Math.round(sents.length / 8)));
    const top = sents.map((s, i) => ({ s, i, sc: score(s) })).sort((a, b) => b.sc - a.sc).slice(0, n).sort((a, b) => a.i - b.i);
    const hs = headings(S.doc, 2).map((h) => h.text).filter(Boolean);
    const box = el('div', {}, hs.length ? el('div', {}, el('b', {}, '구성: '), hs.join(' · ')) : null, el('ul', {}, top.map((t) => el('li', {}, t.s))), el('div', { class: 'muted' }, '문서 안의 문장을 골라 만든 요약입니다. 생성형 AI 요약은 [AI 연결]로 Claude 에게 맡길 수 있습니다.'));
    openDialog({ title: '요약', width: 560, body: box, buttons: [{ label: '문서 앞에 넣기', action: () => run('insertSummary', top.map((t) => t.s)) }, { label: '닫기', primary: true }] });
  },
  insertSummary: (list) => change((d) => { d.doc.body = [para('요약', { style: 'Heading1' }), ...list.map((s) => para(s, { style: 'ListParagraph' })), ...d.doc.body]; return { p: [0], o: 0 }; }),
  immersiveReader: () => { run('viewRead'); run('readAloud'); },
  newWindow: () => { try { sessionStorage.setItem('wiword:copy', JSON.stringify(S.doc)); window.open(`${location.pathname}?copy=1`, '_blank'); } catch { toast('새 창을 열 수 없습니다'); } },
  splitView: () => toast('창 나누기는 [새 창]으로 같은 문서를 나란히 띄워 쓰세요'),
});

/** 두 문서 비교 → 변경 내용(삽입 · 삭제) 표시 문서 (단락 · 낱말 단위 LCS) */
export function compareDocs(a, b) {
  const pa = []; const pb = [];
  walkParas(a.body, (p) => pa.push(p));
  walkParas(b.body, (p) => pb.push(p));
  const ka = pa.map(paraPlain); const kb = pb.map(paraPlain);
  const pairs = lcs(ka, kb, (x, y) => x === y || similar(x, y));
  const out = [];
  const auth = { author: '비교', date: new Date().toISOString() };
  let i = 0; let j = 0;
  for (const [x, y] of [...pairs, [pa.length, pb.length]]) {
    while (i < x) { out.push({ ...clone(pa[i]), id: uid('p'), runs: pa[i].runs.map((r) => (r.obj ? r : { ...r, del: auth })) }); i++; }
    while (j < y) { out.push({ ...clone(pb[j]), id: uid('p'), runs: pb[j].runs.map((r) => (r.obj ? r : { ...r, ins: auth })) }); j++; }
    if (x < pa.length && y < pb.length) {
      if (ka[x] === kb[y]) out.push({ ...clone(pb[y]), id: uid('p') });
      else out.push({ ...clone(pb[y]), id: uid('p'), runs: wordDiff(ka[x], kb[y], propsOnlyRun(pb[y]), auth) });
      i = x + 1; j = y + 1;
    }
  }
  return { ...clone(b), body: out.length ? out : [para('')] };
}
const propsOnlyRun = (p) => { const r = p.runs.find((x) => !x.obj) ?? {}; const o = { ...r }; delete o.text; return o; };
function similar(x, y) { if (!x || !y) return false; const a = new Set(x.split(/\s+/)); const b = y.split(/\s+/); const same = b.filter((w) => a.has(w)).length; return same / Math.max(a.size, b.length) > 0.5; }
function lcs(A, B, eq) {
  const n = A.length; const m = B.length;
  if (n * m > 4e6) return [];
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = eq(A[i], B[j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const out = [];
  let i = 0; let j = 0;
  while (i < n && j < m) { if (eq(A[i], B[j])) { out.push([i, j]); i++; j++; } else if (dp[i + 1][j] >= dp[i][j + 1]) i++; else j++; }
  return out;
}
function wordDiff(a, b, props, auth) {
  const wa = a.split(/(\s+)/); const wb = b.split(/(\s+)/);
  const pairs = lcs(wa, wb, (x, y) => x === y);
  const runs = [];
  let i = 0; let j = 0;
  for (const [x, y] of [...pairs, [wa.length, wb.length]]) {
    if (i < x) runs.push({ ...props, text: wa.slice(i, x).join(''), del: auth });
    if (j < y) runs.push({ ...props, text: wb.slice(j, y).join(''), ins: auth });
    if (x < wa.length) runs.push({ ...props, text: wb[y] });
    i = x + 1; j = y + 1;
  }
  return normalizeRuns(runs.filter((r) => r.text));
}
void insertRuns; void propsAt; void orderPos; void effRun; void sectionObj; void sectionAtTop; void scrollToCaret; void firstPos;
