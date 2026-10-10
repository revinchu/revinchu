// 오른쪽 클릭 메뉴 — PowerPoint (Microsoft 365, 한국어) 와 같은 구성 · 순서 · 괄호 글자, 위에 미니 도구 모음
import { S, curSlide, selObjects, selOne, run, emit, change } from './state.js';
import { el, openMenu, closeMenus, icon } from './ui.js';
import { startEdit } from './editor.js';

const at = (e) => ({ x: e.clientX, y: e.clientY });
const pasteOptions = () => ({ label: '붙여넣기 옵션:', icon: 'paste', submenu: [
  { label: '대상 테마 사용(H)', icon: 'paste', action: () => run('paste') },
  { label: '원본 서식 유지(K)', icon: 'painter', action: () => run('pasteInternal') },
  { label: '그림(U)', icon: 'picture', action: () => run('pasteAsPicture') },
  { label: '텍스트만 유지(T)', icon: 'textbox', action: async () => { const t = await navigator.clipboard?.readText?.().catch(() => ''); if (t) run('pasteText', t); } },
] });
const arrange = () => [
  { label: '그룹화(G)', icon: 'group', submenu: [{ label: '그룹(G)', key: 'Ctrl+G', action: () => run('group') }, { label: '그룹 해제(U)', key: 'Ctrl+Shift+G', action: () => run('ungroup') }, { label: '재그룹(E)', action: () => run('regroup') }] },
  { label: '맨 앞으로 가져오기(R)', icon: 'bringForward', submenu: [{ label: '맨 앞으로 가져오기(R)', key: 'Ctrl+Shift+]', action: () => run('bringToFront') }, { label: '앞으로 가져오기(F)', key: 'Ctrl+]', action: () => run('bringForward') }] },
  { label: '맨 뒤로 보내기(K)', icon: 'sendBackward', submenu: [{ label: '맨 뒤로 보내기(K)', key: 'Ctrl+Shift+[', action: () => run('sendToBack') }, { label: '뒤로 보내기(B)', key: 'Ctrl+[', action: () => run('sendBackward') }] },
];
const common = () => [
  { label: '링크(I)', icon: 'link', key: 'Ctrl+K', action: () => run('hyperlink') },
  { label: '스마트 조회(L)', icon: 'search', action: () => run('smartLookup') },
];
const tail = (formatLabel, pane = 'shape') => [
  { label: '대체 텍스트 편집(A)...', icon: 'info', action: () => run('altText') },
  { label: '크기 및 위치(Z)...', icon: 'select', action: () => run('openFormatPane', pane) },
  { label: formatLabel, icon: 'effects', action: () => run('openFormatPane', pane) },
  { sep: true },
  { label: '새 메모(M)', icon: 'newComment', action: () => run('newComment') },
];
const cutCopy = (enabled = true) => [
  { label: '잘라내기(T)', icon: 'cut', key: 'Ctrl+X', disabled: !enabled, action: () => run('cut') },
  { label: '복사(C)', icon: 'copy', key: 'Ctrl+C', disabled: !enabled, action: () => run('copy') },
  pasteOptions(),
];

/** 미니 도구 모음 (메뉴 위) */
function miniToolbar(pos, kind) {
  const B = (ic, title, fn) => { const b = el('button', { class: 'mt-btn', title, onmousedown: (e) => e.preventDefault(), onclick: (e) => { e.stopPropagation(); const r = b.getBoundingClientRect(); closeMenus(); fn({ x: r.left, y: r.bottom + 2 }); } }); b.append(icon(ic)); return b; };
  const T = (ch, title, css, fn) => { const b = el('button', { class: 'mt-btn', title, style: css, onmousedown: (e) => e.preventDefault(), onclick: (e) => { e.stopPropagation(); fn(); } }, ch); return b; };
  const L = (ic, title, fn) => { const b = B(ic, title, fn); b.append(el('span', {}, title)); b.classList.add('wide'); return b; };
  const items = kind === 'text' ? [
    T('가', '굵게', 'font-weight:700', () => run('bold')), T('가', '기울임꼴', 'font-style:italic', () => run('italic')), T('가', '밑줄', 'text-decoration:underline', () => run('underline')),
    B('growFont', '글꼴 크기 크게', () => run('growFont')), B('shrinkFont', '글꼴 크기 작게', () => run('shrinkFont')),
    B('alignLeft', '왼쪽 맞춤', () => run('alignLeft')), B('alignCenter', '가운데 맞춤', () => run('alignCenter')), B('alignRight', '오른쪽 맞춤', () => run('alignRight')),
    B('fontColor', '글꼴 색', (a) => run('fontColorMenu', a)), B('highlight', '텍스트 강조 색', (a) => run('highlightMenu', a)), B('bullets', '글머리 기호', (a) => run('bulletsMenu', a)), B('painter', '서식 복사', () => run('formatPainter')),
  ] : kind === 'picture' ? [
    L('quickStyles', '스타일', (a) => run('pictureStyleMenu', a)), L('crop', '자르기', () => run('cropPicture')),
  ] : kind === 'table' ? [
    L('table', '삽입', (a) => run('tblInsertMenu', a)), L('delete', '삭제', (a) => run('tblDeleteMenu', a)), L('shapeFill', '음영', (a) => run('cellFillMenu', a)), L('border', '테두리', (a) => run('cellBorderMenu', a)),
  ] : kind === 'chart' ? [
    L('chartColumn', '종류', (a) => run('chartKindMenu', a)), L('plus', '요소', (a) => run('chartElementsMenu', a)), L('theme', '색', (a) => run('chartColorsMenu', a)),
  ] : [
    L('quickStyles', '스타일', (a) => run('quickStylesMenu', a)), L('shapeFill', '채우기', (a) => run('shapeFillMenu', a)), L('shapeOutline', '윤곽선', (a) => run('shapeOutlineMenu', a)),
  ];
  const bar = el('div', { class: 'menu mini-toolbar' }, items);
  document.getElementById('menuLayer').prepend(bar); // 메뉴가 맨 뒤 (키보드 · 괄호 글자는 메뉴에)
  const r = bar.getBoundingClientRect();
  bar.style.left = `${Math.max(4, Math.min(innerWidth - r.width - 4, pos.x))}px`;
  bar.style.top = `${Math.max(4, pos.y - r.height - 6)}px`;
  emit('hydrate', bar);
}

/** 편집 화면 (슬라이드) 오른쪽 클릭 */
export function canvasMenu(e, hit) {
  if (!hit && !S.editing && S.sel.size) { S.sel = new Set(); emit('selection'); }
  if (hit && !S.sel.has(hit.id)) { S.sel = new Set(hit.grp ? curSlide().objects.filter((o) => o.grp === hit.grp).map((o) => o.id) : [hit.id]); emit('selection'); }
  const objs = selObjects();
  const o = selOne() ?? (objs.length && new Set(objs.map((x) => x.grp)).size === 1 && objs[0].grp ? objs[0] : null);
  let items;
  let kind = 'shape';
  if (S.editing) {
    kind = 'text';
    items = [
      ...cutCopy(),
      { sep: true },
      { label: '글꼴(F)...', icon: 'fontColor', action: () => run('fontDialog') },
      { label: '단락(P)...', icon: 'lineSpacing', action: () => run('paragraphDialog') },
      { label: '글머리 기호(B)', icon: 'bullets', submenu: [{ label: '없음', action: () => run('bulletsNone') }, { label: '•  채우기 원형', action: () => run('bullets', '•') }, { label: '■  채우기 사각형', action: () => run('bullets', '■') }, { label: '➢  화살표', action: () => run('bullets', '➢') }, { label: '✓  확인 표시', action: () => run('bullets', '✓') }, { sep: true }, { label: '글머리 기호 및 번호 매기기(N)...', action: () => run('bulletsMenu', at(e)) }] },
      { label: '번호 매기기(N)', icon: 'numbering', submenu: [{ label: '1. 2. 3.', action: () => run('numbering', 'arabicPeriod') }, { label: '1) 2) 3)', action: () => run('numbering', 'arabicParenR') }, { label: 'Ⅰ. Ⅱ. Ⅲ.', action: () => run('numbering', 'romanUcPeriod') }, { label: '가. 나. 다.', action: () => run('numbering', 'ganada') }, { label: 'a. b. c.', action: () => run('numbering', 'alphaLcPeriod') }] },
      { label: 'SmartArt로 변환(M)', icon: 'convertSmart', action: () => run('convertToSmartArt') },
      { sep: true },
      ...common(),
      { label: '동의어(Y)', icon: 'search', disabled: true },
      { label: '번역(S)', icon: 'search', action: () => run('translateSel') },
      { sep: true },
      { label: '텍스트 효과 서식(S)...', icon: 'effects', action: () => run('openFormatPane', 'shape') },
      { label: '도형 서식(O)...', icon: 'effects', action: () => run('openFormatPane', 'shape') },
      { sep: true },
      { label: '새 메모(M)', icon: 'newComment', action: () => run('newComment') },
    ];
  } else if (!objs.length) {
    items = [
      ...cutCopy(false),
      { sep: true },
      { label: '레이아웃(L)', icon: 'layout', submenu: layoutItems() },
      { label: '슬라이드 다시 설정(R)', icon: 'resetSlide', action: () => run('resetSlide') },
      { sep: true },
      { label: '눈금 및 안내선(I)', icon: 'gridlines', submenu: [
        { label: '안내선(I)', checked: S.showGuides, action: () => run('toggleGuides') },
        { label: '눈금선(G)', checked: S.showGrid, action: () => run('toggleGrid') },
        { label: '스마트 가이드(S)', checked: S.smartGuides !== false, action: () => { S.smartGuides = S.smartGuides === false; } },
        { sep: true },
        { label: '세로 안내선 추가(V)', action: () => run('addGuide', 'v') },
        { label: '가로 안내선 추가(H)', action: () => run('addGuide', 'h') },
      ] },
      { label: '눈금자(R)', checked: !!S.showRuler, action: () => run('toggleRuler') },
      { label: '배경 서식(B)...', icon: 'formatBg', action: () => run('formatBg') },
      { sep: true },
      { label: '새 메모(M)', icon: 'newComment', action: () => run('newComment') },
    ];
  } else if (o?.type === 'image' || o?.type === 'media') {
    kind = 'picture';
    const svg = S.pres.media[o.media]?.startsWith('data:image/svg');
    items = [
      ...cutCopy(),
      { sep: true },
      ...(svg ? [{ label: '도형으로 변환(O)', icon: 'shapes', action: () => run('convertToShapes') }] : []),
      ...(o.type === 'image' ? [{ label: '그림 바꾸기(A)', icon: 'picture', submenu: [{ label: '파일에서(F)...', action: () => run('changePicture') }, { label: '아이콘에서(I)...', action: () => run('insertIcons') }, { label: '클립보드에서(C)', action: () => run('changePictureFromClipboard') }] }] : []),
      ...arrange(),
      { sep: true },
      ...common(),
      { label: '그림으로 저장(S)...', icon: 'save', action: () => run('saveAsPicture') },
      ...(o.type === 'image' ? [{ label: '자르기(C)', icon: 'crop', action: () => run('cropPicture') }, { label: '배경 제거(B)', icon: 'picture', action: () => run('removeBackground') }] : []),
      { sep: true },
      ...tail('그림 서식(O)...'),
    ];
  } else if (o?.type === 'table') {
    kind = 'table';
    items = [
      ...cutCopy(),
      { sep: true },
      { label: '삽입(I)', icon: 'insert', submenu: [{ label: '왼쪽에 열 삽입(L)', action: () => run('tblColLeft') }, { label: '오른쪽에 열 삽입(R)', action: () => run('tblColRight') }, { label: '위에 행 삽입(A)', action: () => run('tblRowAbove') }, { label: '아래에 행 삽입(B)', action: () => run('tblRowBelow') }] },
      { label: '삭제(D)', icon: 'delete', submenu: [{ label: '열 삭제(C)', action: () => run('tblDeleteCol') }, { label: '행 삭제(R)', action: () => run('tblDeleteRow') }, { label: '표 삭제(T)', action: () => run('deleteSelection') }] },
      { label: '셀 병합(M)', icon: 'merge', action: () => run('tblMerge') },
      { label: '셀 분할(E)...', icon: 'merge', action: () => run('tblSplit') },
      { label: '열 너비를 같게(W)', action: () => run('tblEqualCols') },
      { label: '행 높이를 같게(H)', action: () => run('tblEqualRows') },
      { sep: true },
      ...arrange(),
      ...common(),
      { label: '그림으로 저장(S)...', icon: 'save', action: () => run('saveAsPicture') },
      { sep: true },
      ...tail('도형 서식(O)...'),
    ];
  } else if (o?.type === 'chart') {
    kind = 'chart';
    items = [
      ...cutCopy(),
      { sep: true },
      { label: '다시 설정하여 스타일에 맞추기(A)', icon: 'resetSlide', action: () => run('chartResetStyle') },
      { label: '글꼴(F)...', icon: 'fontColor', action: () => run('chartFontDialog') },
      { label: '차트 종류 변경(Y)...', icon: 'chartColumn', action: () => run('chartTypeDialog') },
      { label: '데이터 선택(E)...', icon: 'table', action: () => run('chartData') },
      { label: '데이터 편집(D)', icon: 'table', submenu: [{ label: '데이터 편집(E)', action: () => run('chartData') }, { label: 'CSV 로 데이터 내보내기(X)', action: () => run('chartExportCsv') }] },
      { sep: true },
      ...arrange(),
      { label: '그림으로 저장(S)...', icon: 'save', action: () => run('saveAsPicture') },
      { label: '대체 텍스트 편집(A)...', icon: 'info', action: () => run('altText') },
      { label: '차트 영역 서식(F)...', icon: 'effects', action: () => run('openFormatPane', 'shape') },
      { sep: true },
      { label: '새 메모(M)', icon: 'newComment', action: () => run('newComment') },
    ];
  } else {
    // 도형 · 글 상자 · 여러 개
    items = [
      ...cutCopy(),
      { sep: true },
      ...(o?.text || o?.type === 'shape' ? [{ label: '텍스트 편집(X)', icon: 'textbox', action: () => { const t = selOne(); if (t) startEdit(t, { end: true }); } }] : []),
      ...(o?.type === 'shape' && !o.grp ? [{ label: '점 편집(E)', icon: 'shapes', action: () => run('editPoints') }] : []),
      ...(o?.type === 'equation' ? [{ label: '수식 편집(E)...', icon: 'equation', action: () => run('editEquation', o.id) }] : []),
      ...arrange(),
      ...(objs.length > 1 ? [{ label: '맞춤(A)', icon: 'align', submenu: [['l', '왼쪽 맞춤(L)'], ['c', '가운데 맞춤(C)'], ['r', '오른쪽 맞춤(R)'], ['t', '위쪽 맞춤(T)'], ['m', '중간 맞춤(M)'], ['b', '아래쪽 맞춤(B)'], ['h', '가로 간격을 동일하게(H)'], ['v', '세로 간격을 동일하게(V)']].map(([k, l]) => ({ label: l, action: () => run('align', k) })) }] : []),
      { sep: true },
      ...common(),
      { label: '그림으로 저장(S)...', icon: 'save', action: () => run('saveAsPicture') },
      ...(o?.type === 'shape' ? [{ label: '기본 도형으로 설정(D)', action: () => run('setDefaultShape') }] : []),
      { sep: true },
      ...tail(o?.txBox ? '도형 서식(O)...' : objs.length > 1 ? '개체 서식(O)...' : '도형 서식(O)...'),
    ];
  }
  openMenu(at(e), items.filter(Boolean));
  miniToolbar(at(e), kind);
}

function layoutItems() {
  return [['title', '제목 슬라이드'], ['titleContent', '제목 및 내용'], ['section', '구역 머리글'], ['twoContent', '콘텐츠 2개'], ['comparison', '비교'], ['titleOnly', '제목만'], ['blank', '빈 화면'], ['contentCaption', '캡션 있는 콘텐츠'], ['pictureCaption', '캡션 있는 그림']].map(([k, l]) => ({ label: l, checked: curSlide()?.layout === k, action: () => run('changeLayout', k) }));
}

/** 슬라이드 목록 (축소판) 오른쪽 클릭 — index null 이면 빈 곳 */
export function thumbMenu(e, index) {
  const s = curSlide();
  const items = index == null ? [
    pasteOptions(),
    { sep: true },
    { label: '새 슬라이드(N)', icon: 'newSlide', key: 'Ctrl+M', action: () => run('newSlide') },
    { label: '구역 추가(A)', icon: 'section', action: () => run('addSection') },
  ] : [
    { label: '잘라내기(T)', icon: 'cut', action: () => { S.focusThumbs = true; run('cut'); } },
    { label: '복사(C)', icon: 'copy', action: () => { S.focusThumbs = true; run('copy'); S.focusThumbs = false; } },
    pasteOptions(),
    { sep: true },
    { label: '새 슬라이드(N)', icon: 'newSlide', key: 'Ctrl+M', action: () => run('newSlide') },
    { label: '슬라이드 복제(A)', icon: 'duplicate', key: 'Ctrl+D', action: () => run('duplicateSlide') },
    { label: '슬라이드 삭제(D)', icon: 'delete', key: 'Delete', action: () => run('deleteSlide') },
    { sep: true },
    { label: '구역 추가(A)', icon: 'section', action: () => run('addSection') },
    { sep: true },
    { label: '레이아웃(L)', icon: 'layout', submenu: layoutItems() },
    { label: '슬라이드 다시 설정(R)', icon: 'resetSlide', action: () => run('resetSlide') },
    { label: '배경 서식(B)...', icon: 'formatBg', action: () => run('formatBg') },
    { label: '사진 앨범(P)...', icon: 'picture', action: () => run('photoAlbum') },
    { label: '슬라이드 숨기기(H)', icon: 'hideSlide', checked: !!s?.hidden, action: () => run('hideSlide') },
    { sep: true },
    { label: '새 메모(M)', icon: 'newComment', action: () => run('newComment') },
  ];
  openMenu(at(e), items);
}
