// 리본 메뉴 (PowerPoint 한국어판 배치) — 명령은 state.run 으로 실행
import { S, run, emit } from './state.js';
import { el, hydrateIcons } from './ui.js';

// ───────────── 항목 만들기 ─────────────
const L = (cmd, icon, label, o = {}) => ({ t: 'large', cmd, icon, label, ...o });
const M = (cmd, icon, label, o = {}) => ({ t: 'medium', cmd, icon, label, ...o });
const B = (cmd, icon, title, o = {}) => ({ t: 'small', cmd, icon, title, ...o });
const G = (name, o = {}) => ({ t: 'gallery', name, ...o });
const ROWS = (...rows) => ({ t: 'rows', rows });
const COL = (...items) => ({ t: 'col', items });

export const FONTS = ['+mj', '+mn', '맑은 고딕', '굴림', '돋움', '바탕', '궁서', '나눔고딕', '나눔명조', '나눔스퀘어', '나눔바른고딕', 'Noto Sans KR', 'Noto Serif KR', 'Pretendard', '본고딕', 'Arial', 'Arial Black', 'Calibri', 'Cambria', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Consolas', 'Courier New', 'Impact'];
export const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

const FONT_ROW = [
  { t: 'font' }, { t: 'size' },
  B('growFont', 'growFont', '글꼴 크기 크게 (Ctrl+Shift+>)'), B('shrinkFont', 'shrinkFont', '글꼴 크기 작게 (Ctrl+Shift+<)'),
  B('clearFormat', 'clearFormat', '모든 서식 지우기 (Ctrl+Space)'),
];
const FONT_ROW2 = [
  B('bold', null, '굵게 (Ctrl+B)', { glyph: '<span class="kr b">가</span>', on: 'b' }),
  B('italic', null, '기울임꼴 (Ctrl+I)', { glyph: '<span class="kr i">가</span>', on: 'i' }),
  B('underline', null, '밑줄 (Ctrl+U)', { glyph: '<span class="kr u">가</span>', on: 'u' }),
  B('strike', null, '취소선', { glyph: '<span class="kr s">가</span>', on: 's' }),
  B('textShadow', 'shadowText', '텍스트 그림자', { on: 'shadow' }),
  B(null, 'charSpacing', '문자 간격', { menu: 'charSpacingMenu' }),
  B(null, 'changeCase', '대/소문자 바꾸기', { menu: 'changeCaseMenu' }),
  { t: 'color', cmd: 'highlight', icon: 'highlight', title: '텍스트 강조 색', menu: 'highlightMenu', bar: 'hl' },
  { t: 'color', cmd: 'fontColor', icon: 'fontColor', title: '글꼴 색', menu: 'fontColorMenu', bar: 'font' },
];
const PARA_ROW1 = [
  B('bullets', 'bullets', '글머리 기호', { menu: 'bulletsMenu', on: 'bullets' }),
  B('numbering', 'numbering', '번호 매기기', { menu: 'numberingMenu', on: 'numbering' }),
  B('indentLess', 'indentLess', '목록 수준 줄임 (Shift+Tab)'), B('indentMore', 'indentMore', '목록 수준 늘림 (Tab)'),
  B(null, 'lineSpacing', '줄 간격', { menu: 'lineSpacingMenu' }),
];
const PARA_ROW2 = [
  B('alignLeft', 'alignLeft', '왼쪽 맞춤 (Ctrl+L)', { on: 'al' }), B('alignCenter', 'alignCenter', '가운데 맞춤 (Ctrl+E)', { on: 'ac' }),
  B('alignRight', 'alignRight', '오른쪽 맞춤 (Ctrl+R)', { on: 'ar' }), B('alignJustify', 'alignJustify', '양쪽 맞춤 (Ctrl+J)', { on: 'aj' }),
  B(null, 'textDir', '텍스트 방향', { menu: 'textDirMenu' }), B(null, 'textAlignV', '텍스트 맞춤', { menu: 'textAnchorMenu' }),
];

const HOME = { id: 'home', label: '홈', groups: [
  { label: '클립보드', items: [L('paste', 'paste', '붙여넣기', { menu: 'pasteMenu' }), COL(M('cut', 'cut', '잘라내기'), M('copy', 'copy', '복사'), M('formatPainter', 'painter', '서식 복사', { on: 'painter' }))] },
  { label: '슬라이드', items: [L('newSlide', 'newSlide', '새 슬라이드', { menu: 'newSlideMenu' }), COL(M(null, 'layout', '레이아웃', { menu: 'layoutMenu' }), M('resetSlide', 'resetSlide', '다시 설정'), M(null, 'section', '구역', { menu: 'sectionMenu' }))] },
  { label: '글꼴', items: [ROWS(FONT_ROW, FONT_ROW2)], launcher: 'fontDialog' },
  { label: '단락', items: [ROWS(PARA_ROW1, PARA_ROW2)], launcher: 'paragraphDialog' },
  { label: '그리기', items: [G('shapesMini', { rows: 3 }), L(null, 'arrange', '정렬', { menu: 'arrangeMenu' }), L(null, 'quickStyles', '빠른 스타일', { menu: 'quickStylesMenu' }), COL({ t: 'color', cmd: 'shapeFill', icon: 'shapeFill', title: '도형 채우기', label: '도형 채우기', menu: 'shapeFillMenu', bar: 'fill' }, { t: 'color', cmd: 'shapeOutline', icon: 'shapeOutline', title: '도형 윤곽선', label: '도형 윤곽선', menu: 'shapeOutlineMenu', bar: 'line' }, M(null, 'effects', '도형 효과', { menu: 'shapeEffectsMenu' }))], launcher: 'formatPaneShape' },
  { label: '편집', items: [COL(M('find', 'find', '찾기'), M('replace', 'replace', '바꾸기'), M(null, 'select', '선택', { menu: 'selectMenu' }))] },
] };

const INSERT = { id: 'insert', label: '삽입', groups: [
  { label: '슬라이드', items: [L('newSlide', 'newSlide', '새 슬라이드', { menu: 'newSlideMenu' })] },
  { label: '표', items: [L(null, 'table', '표', { menu: 'tableMenu' })] },
  { label: '이미지', items: [L('insertPicture', 'picture', '그림'), L('screenshot', 'crop', '스크린샷')] },
  { label: '일러스트레이션', items: [L(null, 'shapes', '도형', { menu: 'shapesMenu' }), L('insertIcons', 'icons', '아이콘'), L('insertSmartArt', 'convertSmart', 'SmartArt'), L(null, 'chartColumn', '차트', { menu: 'chartMenu' })] },
  { label: '링크', items: [L('hyperlink', 'link', '링크')] },
  { label: '텍스트', items: [L('drawTextbox', 'textbox', '텍스트 상자'), L('headerFooter', 'headerFooter', '머리글/바닥글'), L(null, 'wordart', 'WordArt', { menu: 'wordArtMenu' }), COL(M('insertDate', 'dateTime', '날짜 및 시간'), M('insertSlideNumber', 'slideNumber', '슬라이드 번호'))] },
  { label: '기호', items: [L('symbol', 'symbol', '기호')] },
] };

const DESIGN = { id: 'design', label: '디자인', groups: [
  { label: '테마', items: [G('themes', { wide: true })] },
  { label: '적용', items: [COL(M(null, 'theme', '색', { menu: 'themeColorsMenu' }), M(null, 'fontColor', '글꼴', { menu: 'themeFontsMenu' }), M(null, 'formatBg', '배경 스타일', { menu: 'bgStylesMenu' }))] },
  { label: '사용자 지정', items: [L(null, 'slideSize', '슬라이드 크기', { menu: 'slideSizeMenu' }), L('formatBg', 'formatBg', '배경 서식')] },
  { label: '디자이너', items: [L('designIdeas', 'ai', '디자인 아이디어')] },
] };

const TRANS = { id: 'transitions', label: '전환', groups: [
  { label: '미리 보기', items: [L('previewTransition', 'preview', '미리 보기')] },
  { label: '슬라이드 화면 전환', items: [G('transitions', { wide: true }), L(null, 'transition', '효과 옵션', { menu: 'transitionOptionsMenu' })] },
  { label: '타이밍', items: [{ t: 'transTiming' }, L('applyTransitionAll', 'check', '모두 적용')] },
] };

const ANIM = { id: 'animations', label: '애니메이션', groups: [
  { label: '미리 보기', items: [L('previewAnim', 'preview', '미리 보기')] },
  { label: '애니메이션', items: [G('animations', { wide: true }), L(null, 'animation', '효과 옵션', { menu: 'animOptionsMenu' })] },
  { label: '고급 애니메이션', items: [L(null, 'newSlide', '애니메이션 추가', { menu: 'addAnimMenu' }), COL(M('animPane', 'animPane', '애니메이션 창', { on: 'animPane' }), M('removeAnim', 'cross', '애니메이션 제거'), M('animPainter', 'painter', '애니메이션 복사'))] },
  { label: '타이밍', items: [{ t: 'animTiming' }] },
] };

const SHOW = { id: 'slideshow', label: '슬라이드 쇼', groups: [
  { label: '슬라이드 쇼 시작', items: [L('showFromStart', 'slideshow', '처음부터'), L('showFromCurrent', 'fromCurrent', '현재 슬라이드부터')] },
  { label: '설정', items: [L('setupShow', 'slideshow', '슬라이드 쇼 설정'), L('hideSlide', 'hideSlide', '슬라이드 숨기기', { on: 'hidden' }), COL(M('rehearse', 'timer', '예행 연습'))] },
  { label: '모니터', items: [L('presenterView', 'presenter', '발표자 보기')] },
] };

const REVIEW = { id: 'review', label: '검토', groups: [
  { label: '언어 교정', items: [L('spellCheck', 'check', '맞춤법 검사')] },
  { label: '접근성', items: [L('accessibility', 'info', '접근성 검사')] },
  { label: '정보', items: [L('wordCount', 'stats', '통계')] },
] };

const VIEW = { id: 'view', label: '보기', groups: [
  { label: '프레젠테이션 보기', items: [L('viewNormal', 'normalView', '기본', { on: 'vNormal' }), L('viewOutline', 'outlineView', '개요 보기'), L('viewSorter', 'sorter', '여러 슬라이드', { on: 'vSorter' }), L('viewNotesPage', 'notes', '슬라이드 노트'), L('viewReading', 'readingView', '읽기용 보기')] },
  { label: '표시', items: [COL({ t: 'check', cmd: 'toggleGrid', label: '눈금선', on: 'grid' }, { t: 'check', cmd: 'toggleGuides', label: '안내선', on: 'guides' }), L('toggleNotes', 'notes', '메모', { on: 'notes' })] },
  { label: '확대/축소', items: [L('zoomDialog', 'zoomIn', '확대/축소'), L('fitZoom', 'fit', '창에 맞춤')] },
  { label: '창', items: [L('selectionPane', 'selectionPane', '선택 창', { on: 'selPane' })] },
] };

const HELP = { id: 'help', label: '도움말', groups: [
  { label: '도움말', items: [L('shortcuts', 'keyboard', '바로 가기 키'), L('whatsNew', 'ai', '새로운 기능'), L('about', 'about', 'WIPOINT 정보')] },
] };

// 상황별 탭
const SHAPE_FMT = { id: 'shapeFormat', label: '도형 서식', context: 'shape', groups: [
  { label: '도형 삽입', items: [G('shapesMini', { rows: 3 }), COL(M(null, 'shapes', '도형 편집', { menu: 'changeShapeMenu' }), M('drawTextbox', 'textbox', '텍스트 상자'))] },
  { label: '도형 스타일', items: [G('shapeStyles', { wide: true }), COL({ t: 'color', cmd: 'shapeFill', icon: 'shapeFill', label: '도형 채우기', title: '도형 채우기', menu: 'shapeFillMenu', bar: 'fill' }, { t: 'color', cmd: 'shapeOutline', icon: 'shapeOutline', label: '도형 윤곽선', title: '도형 윤곽선', menu: 'shapeOutlineMenu', bar: 'line' }, M(null, 'effects', '도형 효과', { menu: 'shapeEffectsMenu' }))], launcher: 'formatPaneShape' },
  { label: 'WordArt 스타일', items: [G('wordart', { wide: true })] },
  { label: '정렬', items: [COL(M('bringForward', 'bringForward', '앞으로 가져오기', { menu: 'bringMenu' }), M('sendBackward', 'sendBackward', '뒤로 보내기', { menu: 'sendMenu' }), M('selectionPane', 'selectionPane', '선택 창')), COL(M(null, 'align', '맞춤', { menu: 'alignMenu' }), M(null, 'group', '그룹', { menu: 'groupMenu' }), M(null, 'rotate', '회전', { menu: 'rotateMenu' }))] },
  { label: '크기', items: [{ t: 'objSize' }], launcher: 'formatPaneShape' },
] };
const PIC_FMT = { id: 'pictureFormat', label: '그림 서식', context: 'picture', groups: [
  { label: '조정', items: [L(null, 'effects', '수정', { menu: 'pictureCorrectionsMenu' }), L(null, 'theme', '색', { menu: 'pictureColorMenu' }), COL(M('changePicture', 'picture', '그림 바꾸기'), M('resetPicture', 'resetSlide', '그림 원래대로'))] },
  { label: '그림 스타일', items: [G('pictureStyles', { wide: true }), COL({ t: 'color', cmd: 'shapeOutline', icon: 'shapeOutline', label: '그림 테두리', title: '그림 테두리', menu: 'shapeOutlineMenu', bar: 'line' }, M(null, 'effects', '그림 효과', { menu: 'shapeEffectsMenu' }))], launcher: 'formatPaneShape' },
  { label: '접근성', items: [L('altText', 'info', '대체 텍스트')] },
  { label: '정렬', items: [COL(M('bringForward', 'bringForward', '앞으로 가져오기', { menu: 'bringMenu' }), M('sendBackward', 'sendBackward', '뒤로 보내기', { menu: 'sendMenu' }), M('selectionPane', 'selectionPane', '선택 창')), COL(M(null, 'align', '맞춤', { menu: 'alignMenu' }), M(null, 'group', '그룹', { menu: 'groupMenu' }), M(null, 'rotate', '회전', { menu: 'rotateMenu' }))] },
  { label: '크기', items: [L('cropPicture', 'crop', '자르기', { menu: 'cropMenu' }), { t: 'objSize' }], launcher: 'formatPaneShape' },
] };
const TABLE_DESIGN = { id: 'tableDesign', label: '테이블 디자인', context: 'table', groups: [
  { label: '표 스타일 옵션', items: [COL({ t: 'check', cmd: 'tblOpt', arg: 'firstRow', label: '머리글 행', on: 'tFirstRow' }, { t: 'check', cmd: 'tblOpt', arg: 'lastRow', label: '요약 행', on: 'tLastRow' }, { t: 'check', cmd: 'tblOpt', arg: 'banded', label: '줄무늬 행', on: 'tBanded' }), COL({ t: 'check', cmd: 'tblOpt', arg: 'firstCol', label: '첫째 열', on: 'tFirstCol' }, { t: 'check', cmd: 'tblOpt', arg: 'lastCol', label: '마지막 열', on: 'tLastCol' })] },
  { label: '표 스타일', items: [G('tableStyles', { wide: true }), COL({ t: 'color', cmd: 'cellFill', icon: 'fill', label: '음영', title: '음영', menu: 'cellFillMenu', bar: 'cell' }, M(null, 'borderAll', '테두리', { menu: 'cellBorderMenu' }))] },
] };
const TABLE_LAYOUT = { id: 'tableLayout', label: '레이아웃', context: 'table', groups: [
  { label: '표', items: [L(null, 'select', '선택', { menu: 'tblSelectMenu' })] },
  { label: '행 및 열', items: [L(null, 'delete', '삭제', { menu: 'tblDeleteMenu' }), L('tblRowAbove', 'rowInsert', '위에 삽입'), COL(M('tblRowBelow', 'rowInsert', '아래에 삽입'), M('tblColLeft', 'colInsert', '왼쪽에 삽입'), M('tblColRight', 'colInsert', '오른쪽에 삽입'))] },
  { label: '병합', items: [COL(M('tblMerge', 'merge', '셀 병합'), M('tblSplit', 'textColumns', '셀 분할'))] },
  { label: '셀 크기', items: [COL(M('tblEqualRows', 'tableResize', '행 높이를 같게'), M('tblEqualCols', 'tableResize', '열 너비를 같게'))] },
  { label: '맞춤', items: [ROWS([B('alignLeft', 'alignLeft', '왼쪽 맞춤'), B('alignCenter', 'alignCenter', '가운데 맞춤'), B('alignRight', 'alignRight', '오른쪽 맞춤')], [B('anchorTop', 'alignTop', '위쪽 맞춤'), B('anchorMiddle', 'alignMiddle', '세로 가운데 맞춤'), B('anchorBottom', 'alignBottom', '아래쪽 맞춤')])] },
] };
const CHART_DESIGN = { id: 'chartDesign', label: '차트 디자인', context: 'chart', groups: [
  { label: '차트 레이아웃', items: [L(null, 'chartColumn', '차트 요소 추가', { menu: 'chartElementsMenu' })] },
  { label: '차트 스타일', items: [L(null, 'theme', '색 변경', { menu: 'chartColorsMenu' })] },
  { label: '데이터', items: [L('chartData', 'table', '데이터 편집')] },
  { label: '종류', items: [L(null, 'chartBar', '차트 종류 변경', { menu: 'chartKindMenu' })] },
] };

export const TABS = [HOME, INSERT, DESIGN, TRANS, ANIM, SHOW, REVIEW, VIEW, HELP, SHAPE_FMT, PIC_FMT, TABLE_DESIGN, TABLE_LAYOUT, CHART_DESIGN];

let tabsEl;
let bodyEl;
let active = 'home';
let galleries = {};
let state = {};

export function initRibbon(tabs, body, galleryRenderers) {
  tabsEl = tabs;
  bodyEl = body;
  galleries = galleryRenderers;
  // 단추를 눌러도 편집 중인 글의 선택이 사라지지 않도록
  body.addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
}
export function setRibbonTab(id) { active = id; renderRibbon(); }
export const ribbonTab = () => active;

export function renderRibbon() {
  if (!tabsEl) return;
  // 리본 입력 칸(글꼴 · 크기 등)을 쓰는 중이면 나중에 (입력 값이 사라지지 않도록)
  const ae = document.activeElement;
  if (ae && bodyEl.contains(ae) && (ae.tagName === 'INPUT' || ae.tagName === 'SELECT')) {
    if (!ae.dataset.waitRender) { ae.dataset.waitRender = '1'; ae.addEventListener('blur', () => setTimeout(renderRibbon, 0), { once: true }); }
    return;
  }
  state = run('ribbonState') ?? {};
  const ctx = state.context ?? [];
  const visible = TABS.filter((t) => !t.context || ctx.includes(t.context));
  if (!visible.some((t) => t.id === active)) active = 'home';
  tabsEl.innerHTML = '';
  tabsEl.append(el('button', { class: 'ribbon-tab file', onclick: () => run('backstage') }, '파일'));
  for (const t of visible) {
    tabsEl.append(el('button', { class: `ribbon-tab${t.id === active ? ' active' : ''}${t.context ? ' ctx' : ''}`, onclick: () => { active = t.id; renderRibbon(); } }, t.label));
  }
  tabsEl.append(el('span', { class: 'spacer' }),
    el('button', { class: 'ribbon-action', onclick: () => run('showFromCurrent'), title: '현재 슬라이드부터 (Shift+F5)' }, el('span', { 'data-icon': 'slideshow' }), '슬라이드 쇼'),
    el('button', { class: 'ribbon-action primary', onclick: () => run('share') }, el('span', { 'data-icon': 'share' }), '공유'));
  const tab = TABS.find((t) => t.id === active);
  bodyEl.innerHTML = '';
  for (const g of tab.groups) {
    const gEl = el('div', { class: 'rgroup' }, el('div', { class: 'rgroup-body' }, g.items.map(renderItem)), el('div', { class: 'rgroup-label' }, g.label));
    if (g.launcher) gEl.append(el('button', { class: 'rgroup-launcher', title: `${g.label} 설정`, onclick: () => run(g.launcher) }, '⬊'));
    bodyEl.append(gEl);
  }
  hydrateIcons(tabsEl);
  hydrateIcons(bodyEl);
}

function iconHtml(name) { return name ? `<span data-icon="${name}"></span>` : ''; }

function renderItem(it) {
  if (!it) return null;
  if (it.t === 'rows') return el('div', { class: 'rcol' }, it.rows.map((r) => el('div', { class: 'rrow' }, r.map(renderItem))));
  if (it.t === 'col') return el('div', { class: 'rcol' }, it.items.map(renderItem));
  if (it.t === 'gallery') return galleries[it.name]?.(it) ?? null;
  if (it.t === 'font' || it.t === 'size') return fontBox(it.t);
  if (it.t === 'transTiming' || it.t === 'animTiming' || it.t === 'objSize') return galleries[it.t]?.(it) ?? null;
  if (it.t === 'check') {
    const c = el('input', { type: 'checkbox', checked: !!state[it.on] });
    c.addEventListener('change', () => run(it.cmd, it.arg));
    return el('label', { class: 'rcheck' }, c, el('span', {}, it.label));
  }
  if (it.t === 'color') {
    const bar = state.bars?.[it.bar] ?? '#000';
    const main = el('button', { class: `rbtn ${it.label ? 'medium' : ''} color-btn`, title: it.title, onclick: () => run(it.cmd) },
      el('span', { class: 'cb-ic', html: iconHtml(it.icon) }, el('i', { class: 'color-bar', style: { background: bar } })), it.label ? el('span', { class: 'lbl' }, it.label) : null);
    const caret = el('button', { class: 'rbtn caret-only', title: it.title, 'data-icon': 'chevronDown' });
    caret.addEventListener('click', () => run(it.menu, caret));
    return el('span', { class: 'rsplit' }, main, caret);
  }
  const on = it.on && state[it.on];
  const cls = `rbtn ${it.t}${on ? ' on' : ''}${it.menu && it.cmd ? ' split' : ''}`;
  const btn = el('button', { class: cls, title: it.title ?? it.label ?? '', disabled: state.disabled?.[it.cmd] || undefined });
  btn.innerHTML = `${it.glyph ?? iconHtml(it.icon)}${it.label ? `<span class="lbl">${it.label}</span>` : ''}${it.menu ? '<span class="caret" data-icon="chevronDown"></span>' : ''}`;
  btn.addEventListener('click', (e) => {
    const caret = e.target.closest('.caret');
    if (it.menu && (!it.cmd || caret)) { run(it.menu, btn); return; }
    run(it.cmd, it.arg);
  });
  return btn;
}

function fontBox(kind) {
  const val = kind === 'font' ? state.font : state.size;
  const list = kind === 'font' ? FONTS : SIZES;
  const id = `dl-${kind}`;
  const shown = kind === 'font' ? (val === '+mj' ? `${S.pres.theme.fonts.major} (제목)` : val === '+mn' ? `${S.pres.theme.fonts.minor} (본문)` : val ?? '') : val == null ? '' : String(val);
  const inp = el('input', { class: `rselect ${kind === 'font' ? 'font-family' : 'font-size'}`, value: shown, list: id, title: kind === 'font' ? '글꼴' : '글꼴 크기', spellcheck: 'false' });
  const dl = el('datalist', { id }, list.map((f) => el('option', { value: kind === 'font' ? (f === '+mj' ? `${S.pres.theme.fonts.major} (제목)` : f === '+mn' ? `${S.pres.theme.fonts.minor} (본문)` : f) : String(f) })));
  const apply = () => {
    let v = inp.value.trim();
    if (!v) return;
    if (kind === 'font') {
      if (v.endsWith('(제목)')) v = '+mj';
      else if (v.endsWith('(본문)')) v = '+mn';
      run('fontFamily', v);
    } else {
      const n = Number(v);
      if (Number.isFinite(n) && n > 0 && n <= 4000) run('fontSize', n);
    }
  };
  inp.addEventListener('change', apply);
  inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); apply(); emit('focusCanvas'); } });
  inp.addEventListener('mousedown', () => { inp.dataset.prev = inp.value; });
  return el('span', { class: 'rfont' }, inp, dl);
}
