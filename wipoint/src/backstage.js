// 파일 탭 (Backstage): 홈 · 새로 만들기 · 열기 · 정보 · 저장 · 인쇄 · 공유 · 내보내기
import { S, run, change } from './state.js';
import { slideHtml } from './render.js';
import { TEMPLATES, buildTemplate } from './templates.js';
import { THEMES } from './themes.js';
import { newPresentation } from './model.js';
import { setDocument, recentDocs, openRecent } from './fileio.js';
import { el, hydrateIcons } from './ui.js';

const thumbCache = new Map();
function deckThumb(key, build, w = 196) {
  if (thumbCache.has(key)) return thumbCache.get(key);
  const p = build();
  const sc = w / p.size.w;
  const html = `<div class="tw" style="width:${w}px;height:${Math.round(p.size.h * sc)}px"><div class="tw-in" style="transform:scale(${sc})">${slideHtml(p, p.slides[0], { index: 0 })}</div></div>`;
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
    ...[['home', '홈', 'H'], ['new', '새로 만들기', 'N'], ['open', '열기', 'O'], ['info', '정보', 'I'], ['save', '저장', 'S'], ['saveAs', '다른 이름으로 저장', 'A'], ['print', '인쇄', 'P'], ['share', '공유', 'Z'], ['export', '내보내기', 'E']].map(([k, l, kt]) => el('button', { class: k === page ? 'on' : '', 'data-kt': kt, onclick: () => openBackstage(k) }, l)),
    el('span', { class: 'spacer' }),
    el('button', { onclick: () => { closeBackstage(); run('about'); } }, '정보 (WIPOINT)'));
  root = el('div', { class: 'backstage', tabIndex: -1 }, nav, main);
  root.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.stopPropagation(); closeBackstage(); } });
  document.body.append(root);
  (PAGES[page] ?? PAGES.home)(main);
  if (!root) return; // 페이지가 바로 닫은 경우 (저장 등)
  hydrateIcons(root);
  root.focus();
}

const h2 = (t) => el('h2', {}, t);

function newCards() {
  return el('div', { class: 'tpl-cards' }, TEMPLATES.map((t) => el('button', { class: 'tpl-card', title: t.desc, onclick: () => { setDocument(buildTemplate(t.id), t.id === 'blank' ? '프레젠테이션1' : t.name); closeBackstage(); } },
    el('div', { class: 'tpl-thumb', html: deckThumb(`tpl:${t.id}`, () => buildTemplate(t.id)) }), el('b', {}, t.name), el('small', {}, t.desc))));
}
function themeCards() {
  return el('div', { class: 'tpl-cards' }, THEMES.map((t) => el('button', { class: 'tpl-card', onclick: () => { setDocument(newPresentation({ theme: t.name }), '프레젠테이션1'); closeBackstage(); } },
    el('div', { class: 'tpl-thumb', html: deckThumb(`theme:${t.name}`, () => { const p = newPresentation({ theme: t.name }); p.slides[0].objects[0].text.paras[0].runs = [{ t: t.label }]; return p; }) }), el('b', {}, t.label))));
}
async function recentList(box) {
  const list = await recentDocs().catch(() => []);
  box.replaceChildren();
  if (!list.length) { box.append(el('p', { class: 'muted' }, '최근에 연 프레젠테이션이 없습니다. (이 브라우저에 자동으로 보관됩니다)')); return; }
  box.append(el('table', { class: 'recent' }, list.map((d) => el('tr', { class: 'clickable', onclick: async () => { await openRecent(d.id); closeBackstage(); } },
    el('td', {}, el('span', { 'data-icon': 'slideshow' })), el('td', {}, el('b', {}, d.name)), el('td', { class: 'muted' }, `슬라이드 ${d.slides}장`), el('td', { class: 'muted' }, new Date(d.at).toLocaleString('ko-KR'))))));
  hydrateIcons(box);
}

const PAGES = {
  home: (m) => {
    const rec = el('div', {});
    m.append(h2('좋은 하루입니다'), el('h3', {}, '새로 만들기'), newCards(), el('h3', {}, '최근 항목'), rec);
    recentList(rec);
  },
  new: (m) => m.append(h2('새로 만들기'), el('h3', {}, '서식 파일'), newCards(), el('h3', {}, '테마'), themeCards()),
  open: (m) => {
    const rec = el('div', {});
    m.append(h2('열기'), el('div', { class: 'bs-actions' },
      el('button', { class: 'bs-big', onclick: () => { closeBackstage(); run('open'); } }, el('span', { 'data-icon': 'open' }), el('b', {}, '찾아보기'), el('small', {}, '.pptx · .ppsx · .potx · PDF · WIPOINT(.json)'))),
    el('h3', {}, '최근 항목'), rec);
    recentList(rec);
  },
  info: (m) => {
    const p = S.pres.props ?? {};
    const title = el('input', { type: 'text', value: p.title ?? '' });
    const author = el('input', { type: 'text', value: p.author ?? '' });
    const save = () => change(() => { S.pres.props = { ...p, title: title.value, author: author.value }; }, { scope: 'none' });
    title.addEventListener('change', save);
    author.addEventListener('change', save);
    const media = Object.values(S.pres.media ?? {}).reduce((n, u) => n + u.length * 0.75, 0);
    m.append(h2('정보'), el('div', { class: 'info-grid' },
      el('div', { class: 'info-card' }, el('h3', {}, '속성'),
        el('label', {}, el('span', {}, '제목'), title), el('label', {}, el('span', {}, '만든 이'), author),
        el('p', {}, `슬라이드 ${S.pres.slides.length}장 · 숨긴 슬라이드 ${S.pres.slides.filter((s) => s.hidden).length}장`),
        el('p', {}, `크기 ${Math.round(S.pres.size.w * 2.54 / 96 * 100) / 100} × ${Math.round(S.pres.size.h * 2.54 / 96 * 100) / 100} cm · 테마 ${S.pres.theme.label ?? S.pres.theme.name}`),
        el('p', {}, `그림 ${Object.keys(S.pres.media ?? {}).length}개 (약 ${(media / 1048576).toFixed(1)}MB)`)),
      el('div', { class: 'info-card' }, el('h3', {}, '문제 확인'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('accessibility'); } }, '접근성 검사'),
        el('button', { class: 'btn', onclick: () => { closeBackstage(); run('wordCount'); } }, '통계'))));
  },
  save: () => { closeBackstage(); run('save'); },
  saveAs: (m) => m.append(h2('다른 이름으로 저장'), el('div', { class: 'bs-actions' },
    big('save', 'PowerPoint 프레젠테이션 (.pptx)', 'PowerPoint · Keynote · Google 프레젠테이션에서 열 수 있습니다', () => run('saveAs', 'pptx')),
    big('newDoc', 'WIPOINT 문서 (.json)', '모든 정보를 그대로 보관 (WIPOINT 전용)', () => run('saveAs', 'json')),
    big('print', 'PDF', '인쇄 대화상자에서 [PDF로 저장]을 고르세요', () => run('printNow', { mode: 'slides' })),
    big('picture', 'PNG 그림 (현재 슬라이드)', '1920픽셀 너비', () => run('exportPng', false)))),
  print: (m) => {
    const mode = el('select', {}, [['slides', '전체 페이지 슬라이드'], ['notes', '슬라이드 노트'], ['handout2', '유인물 (한 페이지에 2개)'], ['handout3', '유인물 (한 페이지에 3개, 메모 줄)'], ['handout6', '유인물 (한 페이지에 6개)']].map(([v, l]) => el('option', { value: v }, l)));
    const from = el('input', { type: 'number', min: '1', value: '1' });
    const to = el('input', { type: 'number', min: '1', value: String(S.pres.slides.length) });
    const hidden = el('input', { type: 'checkbox' });
    m.append(h2('인쇄'), el('div', { class: 'print-form' },
      el('label', {}, el('span', {}, '인쇄 모양'), mode),
      el('label', {}, el('span', {}, '슬라이드'), from, el('span', {}, '~'), to),
      el('label', {}, hidden, el('span', {}, '숨겨진 슬라이드 인쇄')),
      el('button', { class: 'btn primary big', onclick: () => { closeBackstage(); run('printNow', { mode: mode.value, range: [Number(from.value), Number(to.value)], hidden: hidden.checked }); } }, '인쇄'),
      el('p', { class: 'muted' }, 'PDF 로 만들려면 인쇄 대화상자의 대상에서 [PDF로 저장]을 고르세요.')));
  },
  share: (m) => m.append(h2('공유'), el('div', { class: 'bs-actions' },
    big('share', '읽기 전용 링크', '링크를 받은 사람은 슬라이드 쇼로 볼 수 있습니다', () => { closeBackstage(); run('share'); }),
    big('save', '파일로 보내기 (.pptx)', '메일 · 메신저에 첨부', () => run('saveAs', 'pptx')))),
  export: (m) => m.append(h2('내보내기'), el('div', { class: 'bs-actions' },
    big('print', 'PDF/XPS 문서 만들기', '인쇄 → PDF로 저장', () => { closeBackstage(); run('printNow', { mode: 'slides' }); }),
    big('picture', '현재 슬라이드를 PNG 로', '', () => run('exportPng', false)),
    big('sorter', '모든 슬라이드를 PNG 로', '슬라이드마다 파일 하나', () => run('exportPng', true)),
    big('video', '동영상 만들기', 'MP4/WebM — 전환 · 설정된 시간 포함', () => { closeBackstage(); run('exportVideo'); }),
    big('notes', '유인물 만들기', '한 페이지에 슬라이드 여러 장 + 메모 줄', () => openBackstage('print')))),
};
function big(icon, title, desc, action) {
  return el('button', { class: 'bs-big', onclick: action }, el('span', { 'data-icon': icon }), el('b', {}, title), desc ? el('small', {}, desc) : null);
}

