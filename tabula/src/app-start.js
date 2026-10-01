// WIXEL 3 파일 화면의 공통 구성. 통합 문서 상태는 app.js에서 전달합니다.
import { el } from './ui.js';
import { ICONS } from './icons.js';

export function hubIcon(name) { return el('span', { class: 'hub-icon', html: ICONS[name] ?? ICONS.table ?? '' }); }
export function hubHeading(eyebrow, title, desc) {
  return el('header', { class: 'hub-heading' }, el('span', { class: 'hub-eyebrow' }, eyebrow), el('h2', {}, title), desc ? el('p', {}, desc) : null);
}
export function hubCard(icon, title, text, action, { tag = '', disabled = false } = {}) {
  return el('button', { class: 'hub-card', type: 'button', onclick: action, disabled }, hubIcon(icon),
    el('span', { class: 'hub-card-copy' }, el('b', {}, title), el('small', {}, text)), tag ? el('span', { class: 'hub-tag' }, tag) : null,
    el('span', { class: 'hub-arrow', 'aria-hidden': true }, '↗'));
}
export function hubPreview() {
  const rows = [['채널', '전환', '성과'], ['검색 광고', '128', '82%'], ['브랜드 캠페인', '96', '68%'], ['소셜 미디어', '74', '56%']];
  return el('div', { class: 'hub-preview', 'aria-hidden': true },
    el('div', { class: 'hub-preview-top' }, el('span', { class: 'hub-dot' }), '월간 캠페인 리포트', el('span', { class: 'hub-preview-check' }, '✓ 저장됨')),
    el('div', { class: 'hub-preview-formula' }, el('i', {}, 'fx'), '=SUM(B3:B5)'),
    el('div', { class: 'hub-preview-grid' }, rows.map((row, r) => el('div', { class: r ? '' : 'head' },
      el('span', { class: 'row' }, r + 2), row.map((v, c) => el('span', { class: c === 2 && r ? 'bar-cell' : '' }, c === 2 && r ? el('i', { style: { width: v } }) : null, el('em', {}, v)))))),
    el('div', { class: 'hub-preview-bottom' }, el('b', {}, '성과 요약'), el('span', {}, '+'), el('span', {}, '합계: 298')));
}
export function hubDropzone(onBrowse, onFile) {
  const node = el('button', { type: 'button', class: 'hub-dropzone', onclick: onBrowse }, hubIcon('open'),
    el('b', {}, '파일을 이곳에 놓거나 클릭해서 열기'), el('span', {}, 'Excel · CSV · ODS · WIXEL 파일'));
  node.addEventListener('dragover', (e) => { e.preventDefault(); node.classList.add('drag-over'); });
  node.addEventListener('dragleave', () => node.classList.remove('drag-over'));
  node.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); node.classList.remove('drag-over'); if (e.dataTransfer?.files?.[0]) onFile(e.dataTransfer.files[0]); });
  return node;
}
export function hubEmpty(title, desc) { return el('div', { class: 'hub-empty' }, hubIcon('folder'), el('b', {}, title), el('p', {}, desc)); }
