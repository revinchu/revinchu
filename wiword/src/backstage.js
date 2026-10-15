// 파일 탭 (Backstage): 홈 · 새로 만들기 · 열기 · 정보 · 저장 · 다른 이름으로 저장 · 인쇄 · 공유 · 내보내기 · 닫기
import { S, run, setDoc } from './state.js';
import { TEMPLATES } from './sample.js';
import { blockHtml } from './render.js';
import { docStats, pxToCm, newDoc } from './model.js';
import { recentDocs, openRecent, removeRecent, saveAs, printDoc } from './fileio.js';
import { el, hydrateIcons, toast } from './ui.js';

const thumbCache = new Map();
function docThumb(key, build) {
  if (thumbCache.has(key)) return thumbCache.get(key);
  const d = build();
  const sec = d.section;
  const w = 150;
  const sc = w / sec.page.w;
  let inner = '';
  try { inner = d.body.slice(0, 14).map((b) => blockHtml(d, b, { num: new Map() })).join(''); } catch { inner = ''; }
  const html = `<div class="tw" style="width:${w}px;height:${Math.round(sec.page.h * sc)}px"><div class="tw-in wd" style="transform:scale(${sc});width:${sec.page.w}px;padding:${sec.margin.t}px ${sec.margin.r}px 0 ${sec.margin.l}px">${inner}</div></div>`;
  thumbCache.set(key, html);
  return html;
}

let root = null;
export function closeBackstage() { root?.remove(); root = null; }
export const backstageOpen = () => !!root;

export function openBackstage(page = 'home') {
  closeBackstage();
  const main = el('div', { class: 'backstage-main' });
  const nav = el('nav', { class: 'backstage-nav' },
    el('button', { class: 'back', title: '돌아가기 (Esc)', onclick: closeBackstage }, '←'),
    ...[['home', '홈', 'H'], ['new', '새로 만들기', 'N'], ['open', '열기', 'O'], ['info', '정보', 'I'], ['save', '저장', 'S'], ['saveAs', '다른 이름으로 저장', 'A'], ['print', '인쇄', 'P'], ['share', '공유', 'Z'], ['export', '내보내기', 'E'], ['pdfTools', 'PDF 도구', 'T'], ['close', '닫기', 'C']].map(([k, l, kt]) => el('button', { class: k === page ? 'on' : '', 'data-kt': kt, onclick: () => openBackstage(k) }, l)),
    el('span', { class: 'spacer' }),
    el('button', { 'data-kt': 'T', onclick: () => { closeBackstage(); run('optionsDialog'); } }, '옵션'),
    el('button', { onclick: () => { closeBackstage(); run('about'); } }, '정보 (WIWORD)'));
  root = el('div', { class: 'backstage', tabIndex: -1 }, nav, main);
  root.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.stopPropagation(); closeBackstage(); } });
  document.body.append(root);
  (PAGES[page] ?? PAGES.home)(main);
  if (!root) return;
  hydrateIcons(root);
  root.focus();
}

const h2 = (t) => el('h2', {}, t);
function newCards() {
  return el('div', { class: 'tpl-cards' }, TEMPLATES.map((t) => el('button', { class: 'tpl-card', title: t.desc, onclick: () => { setDoc(t.build(), t.id === 'blank' ? '문서1' : t.name); S.fileHandle = null; S.fileKind = null; S.docId = null; closeBackstage(); } },
    el('div', { class: 'tpl-thumb', html: docThumb(t.id, t.build) }), el('b', {}, t.name), el('small', {}, t.desc))));
}
async function recentList(box) {
  const list = await recentDocs();
  box.replaceChildren();
  if (!list.length) { box.append(el('p', { class: 'muted' }, '최근에 편집한 문서가 없습니다. (이 브라우저에 자동으로 보관됩니다)')); return; }
  box.append(el('table', { class: 'recent' }, list.map((d) => {
    const tr = el('tr', { class: 'clickable', onclick: async () => { await openRecent(d.id); closeBackstage(); } },
      el('td', {}, el('span', { 'data-icon': 'newDoc' })), el('td', {}, el('b', {}, d.name), el('div', { class: 'muted' }, (d.preview ?? '').slice(0, 80))), el('td', { class: 'muted' }, new Date(d.at).toLocaleString('ko-KR')),
      el('td', {}, el('button', { class: 'sb-btn', title: '목록에서 제거', onclick: async (e) => { e.stopPropagation(); await removeRecent(d.id); recentList(box); } }, '✕')));
    return tr;
  })));
  hydrateIcons(box);
}
function big(icon, title, desc, action) {
  return el('button', { class: 'bs-big', onclick: action }, el('span', { 'data-icon': icon }), el('b', {}, title), desc ? el('small', {}, desc) : null);
}

const PAGES = {
  home: (m) => { const rec = el('div', {}); m.append(h2('좋은 하루입니다'), el('h3', {}, '새로 만들기'), newCards(), el('h3', {}, '최근 항목'), rec); recentList(rec); },
  new: (m) => m.append(h2('새로 만들기'), newCards()),
  open: (m) => {
    const rec = el('div', {});
    m.append(h2('열기'), el('div', { class: 'bs-actions' },
      big('open', '찾아보기', 'Word(.docx · .doc) · 한글(.hwp · .hwpx) · PDF · 마크다운 · HTML · 텍스트 · RTF · ODT', () => { closeBackstage(); run('open'); })),
    el('h3', {}, '최근 항목'), rec);
    recentList(rec);
  },
  info: (m) => {
    const p = S.doc.props ?? {};
    const st = docStats(S.doc);
    const field = (key, label) => { const i = el('input', { type: 'text', value: p[key] ?? '' }); i.addEventListener('change', () => run('docProps', { [key]: i.value })); return el('label', {}, el('span', {}, label), i); };
    const sec = S.doc.section;
    m.append(h2('정보'), el('div', { class: 'info-grid' },
      el('div', { class: 'info-card' }, el('h3', {}, '속성'), field('title', '제목'), field('subject', '주제'), field('author', '만든 이'), field('keywords', '태그'), field('comments', '설명'),
        el('p', {}, `쪽 ${S.layout?.pages.length ?? 1} · 단어 ${st.words.toLocaleString()} · 글자 ${st.chars.toLocaleString()} (공백 제외 ${st.charsNoSpace.toLocaleString()}) · 단락 ${st.paras.toLocaleString()}`),
        el('p', {}, `용지 ${pxToCm(sec.page.w).toFixed(1)} × ${pxToCm(sec.page.h).toFixed(1)} cm`),
        el('p', { class: 'muted' }, `만든 날짜 ${p.created ? new Date(p.created).toLocaleString('ko-KR') : '-'} · 수정한 날짜 ${p.modified ? new Date(p.modified).toLocaleString('ko-KR') : '-'}`)),
      el('div', { class: 'info-card' }, el('h3', {}, '문서 보호'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('restrictEditing'); } }, '편집 제한…'),
        el('h3', {}, '문서 검사'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('accessibilityCheck'); } }, '접근성 검사'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('spellCheck'); } }, '맞춤법 검사'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('wordCountDialog'); } }, '단어 개수'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('compareDialog'); } }, '문서 비교…'))));
  },
  save: () => { closeBackstage(); run('save'); },
  saveAs: (m) => m.append(h2('다른 이름으로 저장'), el('div', { class: 'bs-actions' },
    big('save', 'Word 문서 (.docx)', 'Microsoft Word · 한컴오피스 · Google 문서 · Pages 에서 열 수 있습니다', () => saveAs('docx')),
    big('hwp', '한글 문서 (.hwpx)', '한컴오피스 한글 2014 이상 (개방형 한글 문서)', () => saveAs('hwpx')),
    big('pdf', 'PDF', '바로 .pdf 파일로 (글자 · 글꼴 · 그림 · 책갈피 포함)', () => { closeBackstage(); run('exportPdf'); }),
    big('markdown', '마크다운 (.md)', 'GitHub · Notion · Obsidian · AI 에 넣기 좋은 글', () => saveAs('md')),
    big('book', 'EPUB 전자책 (.epub)', '제목 1 마다 장 하나, 차례 포함 — 전자책 앱 · 리더기에서 읽기', () => saveAs('epub')),
    big('webLayout', '웹 페이지 (.html)', '브라우저에서 바로 보기', () => saveAs('html')),
    big('newDoc', '일반 텍스트 (.txt)', '서식 없이 글만', () => saveAs('txt')),
    big('layout', 'Word 서식 파일 (.dotx)', '이 모양으로 새 문서를 만들 때', () => saveAs('dotx')),
    big('save', 'WIWORD 문서 (.wiword)', '모든 정보를 그대로 보관 (WIWORD 전용)', () => saveAs('wiword')))),
  print: (m) => {
    const total = S.layout?.pages.length ?? 1;
    const from = el('input', { type: 'number', min: '1', max: String(total), value: '1' });
    const to = el('input', { type: 'number', min: '1', max: String(total), value: String(total) });
    m.append(h2('인쇄'), el('div', { class: 'print-form' },
      el('label', {}, el('span', {}, '페이지'), from, el('span', {}, '~'), to, el('span', { class: 'muted' }, ` (전체 ${total}쪽)`)),
      el('button', { class: 'btn primary big', onclick: () => { closeBackstage(); printDoc({ from: Number(from.value), to: Number(to.value) }); } }, '인쇄'),
      el('p', { class: 'muted' }, 'PDF 파일은 [파일 › 내보내기 › PDF/XPS 문서 만들기]로 바로 만들 수 있습니다. 인쇄 대화상자의 [PDF로 저장]도 됩니다 (용지 크기는 첫 쪽 크기).'),
      el('button', { class: 'btn', onclick: () => { closeBackstage(); run('pageSetupDialog'); } }, '페이지 설정…')));
  },
  share: (m) => m.append(h2('공유'), el('div', { class: 'bs-actions' },
    big('share', '파일로 보내기 (.docx)', '메일 · 메신저에 첨부', () => saveAs('docx')),
    big('hwp', '한글 파일로 보내기 (.hwpx)', '공공기관 제출용', () => saveAs('hwpx')),
    big('pdf', 'PDF 로 보내기', '받는 사람이 고칠 수 없게', () => { closeBackstage(); run('exportPdf'); }),
    big('mcp', 'AI 에게 맡기기 (MCP)', 'Claude 같은 AI 가 이 문서를 읽고 고치도록 연결', () => { closeBackstage(); run('togglePane', 'mcp'); }))),
  export: (m) => m.append(h2('내보내기'), el('div', { class: 'bs-actions' },
    big('pdf', 'PDF/XPS 문서 만들기', '글꼴을 넣은 PDF 파일 (제목은 책갈피로)', () => { closeBackstage(); run('exportPdf'); }),
    big('hwp', '한글 문서 (.hwpx)', '', () => saveAs('hwpx')),
    big('markdown', '마크다운 (.md)', '그림은 문서 안에 포함 (data URL)', () => saveAs('md')),
    big('book', 'EPUB 전자책', '장 · 차례 · 그림 포함', () => saveAs('epub')),
    big('webLayout', 'HTML', '', () => saveAs('html')),
    big('pdf', 'PDF 도구', '병합 · 분할 · 추출 · 회전 · 모아찍기 · 배경 · 쪽 번호 · 주석 지우기 · 이미지로 PDF', () => { closeBackstage(); run('pdfTools'); }))),
  pdfTools: () => { closeBackstage(); run('pdfTools'); },
  close: () => {
    closeBackstage();
    if (S.dirty && !confirm('저장하지 않은 변경 내용이 있습니다. 닫을까요? (자동 저장된 사본은 [열기 › 최근 항목]에 남습니다)')) return;
    setDoc(newDoc(), '문서1');
    S.fileHandle = null; S.fileKind = null; S.docId = null;
    toast('문서를 닫았습니다');
  },
};

