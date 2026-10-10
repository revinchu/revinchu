// 리본 메뉴 (PowerPoint 한국어판 배치) — 명령은 state.run 으로 실행
import { S, run, emit } from './state.js';
import { el, hydrateIcons, openMenu } from './ui.js';
import { redrawKeytips } from './keytips.js';

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
  { label: '링크', items: [L(null, 'zoomIn', '확대/축소', { menu: 'zoomMenu' }), L('hyperlink', 'link', '링크')] },
  { label: '미디어', items: [L('insertVideo', 'video', '비디오'), L('insertAudio', 'audio', '오디오')] },
  { label: '텍스트', items: [L('drawTextbox', 'textbox', '텍스트 상자'), L('headerFooter', 'headerFooter', '머리글/바닥글'), L(null, 'wordart', 'WordArt', { menu: 'wordArtMenu' }), COL(M('insertDate', 'dateTime', '날짜 및 시간'), M('insertSlideNumber', 'slideNumber', '슬라이드 번호'))] },
  { label: '기호', items: [L('insertEquation', 'equation', '수식'), L('symbol', 'symbol', '기호')] },
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
  { label: '고급 애니메이션', items: [L(null, 'newSlide', '애니메이션 추가', { menu: 'addAnimMenu' }), COL(M('animPane', 'animPane', '애니메이션 창', { on: 'animPane' }), M('removeAnim', 'cross', '애니메이션 제거'), M('animPainter', 'painter', '애니메이션 복사')), COL(M(null, 'select', '트리거', { menu: 'triggerMenu' }))] },
  { label: '타이밍', items: [{ t: 'animTiming' }] },
] };

const SHOW = { id: 'slideshow', label: '슬라이드 쇼', groups: [
  { label: '슬라이드 쇼 시작', items: [L('showFromStart', 'slideshow', '처음부터'), L('showFromCurrent', 'fromCurrent', '현재 슬라이드부터')] },
  { label: '사용자 지정', items: [L(null, 'sorter', '사용자 지정 슬라이드 쇼', { menu: 'customShowMenu' })] },
  { label: '설정', items: [L('setupShow', 'slideshow', '슬라이드 쇼 설정'), L('hideSlide', 'hideSlide', '슬라이드 숨기기', { on: 'hidden' }), COL(M('rehearse', 'timer', '예행 연습'))] },
  { label: '모니터', items: [L('presenterView', 'presenter', '발표자 보기')] },
] };

const REVIEW = { id: 'review', label: '검토', groups: [
  { label: '언어 교정', items: [L('spellCheck', 'check', '맞춤법 검사')] },
  { label: '접근성', items: [L('accessibility', 'info', '접근성 검사')] },
  { label: '정보', items: [L('wordCount', 'stats', '통계')] },
  { label: '메모', items: [L('newComment', 'newComment', '새 메모'), COL(M('deleteComments', 'deleteComment', '삭제'), M('prevComment', 'prev', '이전'), M('nextComment', 'next', '다음')), L('commentsPane', 'comment', '메모 표시', { on: 'cmPane' })] },
] };

const VIEW = { id: 'view', label: '보기', groups: [
  { label: '프레젠테이션 보기', items: [L('viewNormal', 'normalView', '기본', { on: 'vNormal' }), L('viewOutline', 'outlineView', '개요 보기', { on: 'vOutline' }), L('viewSorter', 'sorter', '여러 슬라이드', { on: 'vSorter' }), L('viewNotesPage', 'notes', '슬라이드 노트'), L('viewReading', 'readingView', '읽기용 보기')] },
  { label: '마스터 보기', items: [L('viewMaster', 'layout', '슬라이드 마스터', { on: 'vMaster' })] },
  { label: '표시', items: [COL({ t: 'check', cmd: 'toggleRuler', label: '눈금자', on: 'ruler' }, { t: 'check', cmd: 'toggleGrid', label: '눈금선', on: 'grid' }, { t: 'check', cmd: 'toggleGuides', label: '안내선', on: 'guides' }), L('toggleNotes', 'notes', '메모', { on: 'notes' })] },
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
  { label: '조정', items: [L('removeBackground', 'picture', '배경 제거'), L(null, 'effects', '수정', { menu: 'pictureCorrectionsMenu' }), L(null, 'theme', '색', { menu: 'pictureColorMenu' }), COL(M('changePicture', 'picture', '그림 바꾸기'), M('resetPicture', 'resetSlide', '그림 원래대로'))] },
  { label: '그림 스타일', items: [G('pictureStyles', { wide: true }), COL({ t: 'color', cmd: 'shapeOutline', icon: 'shapeOutline', label: '그림 테두리', title: '그림 테두리', menu: 'shapeOutlineMenu', bar: 'line' }, M(null, 'effects', '그림 효과', { menu: 'shapeEffectsMenu' }))], launcher: 'formatPaneShape' },
  { label: '접근성', items: [L('altText', 'info', '대체 텍스트')] },
  { label: '정렬', items: [COL(M('bringForward', 'bringForward', '앞으로 가져오기', { menu: 'bringMenu' }), M('sendBackward', 'sendBackward', '뒤로 보내기', { menu: 'sendMenu' }), M('selectionPane', 'selectionPane', '선택 창')), COL(M(null, 'align', '맞춤', { menu: 'alignMenu' }), M(null, 'group', '그룹', { menu: 'groupMenu' }), M(null, 'rotate', '회전', { menu: 'rotateMenu' }))] },
  { label: '크기', items: [L('cropPicture', 'crop', '자르기', { menu: 'cropMenu' }), { t: 'objSize' }], launcher: 'formatPaneShape' },
] };
const MEDIA_TAB = { id: 'playback', label: '재생', context: 'media', groups: [
  { label: '미리 보기', items: [L('previewMedia', 'preview', '재생')] },
  { label: '편집', items: [L('mediaOptions', 'timer', '재생 옵션')] },
  { label: '정렬', items: [COL(M('bringForward', 'bringForward', '앞으로 가져오기', { menu: 'bringMenu' }), M('sendBackward', 'sendBackward', '뒤로 보내기', { menu: 'sendMenu' })), COL(M(null, 'align', '맞춤', { menu: 'alignMenu' }), M('altText', 'info', '대체 텍스트'))] },
  { label: '크기', items: [{ t: 'objSize' }] },
] };
const MASTER_TAB = { id: 'slideMaster', label: '슬라이드 마스터', context: 'master', groups: [
  { label: '마스터 편집', items: [L('masterRename', 'section', '이름 바꾸기'), L('masterReset', 'resetSlide', '원래대로')] },
  { label: '배경', items: [COL({ t: 'check', cmd: 'masterHideGraphics', label: '배경 그래픽 숨기기', on: 'masterHide' }), L('formatBg', 'formatBg', '배경 서식')] },
  { label: '닫기', items: [L('closeMaster', 'cross', '마스터 보기 닫기')] },
] };
const SMART_TAB = { id: 'smartDesign', label: 'SmartArt 디자인', context: 'smart', groups: [
  { label: '그래픽 만들기', items: [L('smartAddShape', 'plus', '도형 추가'), L('smartTextPane', 'textbox', '텍스트 창')] },
  { label: '레이아웃', items: [L(null, 'convertSmart', '레이아웃', { menu: 'smartLayoutMenu' })] },
  { label: 'SmartArt 스타일', items: [L('smartColors', 'theme', '색 변경')] },
  { label: '원래대로', items: [L('smartToShapes', 'shapes', '도형으로 변환')] },
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

export const TABS = [HOME, INSERT, DESIGN, TRANS, ANIM, SHOW, REVIEW, VIEW, HELP, SHAPE_FMT, PIC_FMT, TABLE_DESIGN, TABLE_LAYOUT, CHART_DESIGN, MEDIA_TAB, SMART_TAB, MASTER_TAB];

// ───────────── 키 팁 (Alt 누르기 → 글자) — PowerPoint 한국어판과 같은 글자 ─────────────
const TAB_KT = { home: 'H', insert: 'N', design: 'G', transitions: 'K', animations: 'A', slideshow: 'S', review: 'R', view: 'W', help: 'Y2', playback: 'JN', smartDesign: 'JS', slideMaster: 'M', shapeFormat: 'JD', pictureFormat: 'JP', tableDesign: 'JT', tableLayout: 'JL', chartDesign: 'JC' };
const ARRANGE_KT = { bringForward: 'AF', sendBackward: 'AE', selectionPane: 'AP', alignMenu: 'AA', groupMenu: 'AG', rotateMenu: 'AY' };
const KT = {
  home: { paste: 'V', cut: 'X', copy: 'C', formatPainter: 'FP', newSlide: 'I', layoutMenu: 'L', resetSlide: 'RE', sectionMenu: 'T', font: 'FF', size: 'FS', growFont: 'FG', shrinkFont: 'FK', clearFormat: 'E', bold: '1', italic: '2', underline: '3', strike: '4', textShadow: '5', charSpacingMenu: '6', changeCaseMenu: '7', highlight: 'TC', fontColor: 'FC', bullets: 'U', numbering: 'N', indentLess: 'AO', indentMore: 'AI', lineSpacingMenu: 'K', alignLeft: 'AL', alignCenter: 'AC', alignRight: 'AR', alignJustify: 'AJ', textDirMenu: 'AD', textAnchorMenu: 'AT', shapesMini: 'SH', arrangeMenu: 'G', quickStylesMenu: 'Q', shapeFill: 'SF', shapeOutline: 'SO', shapeEffectsMenu: 'SE', find: 'FD', replace: 'RP', selectMenu: 'SL', 'launcher:fontDialog': 'FN', 'launcher:paragraphDialog': 'PG', 'launcher:formatPaneShape': 'DS' },
  insert: { newSlide: 'I', tableMenu: 'T', insertPicture: 'P', screenshot: 'SC', shapesMenu: 'SH', insertIcons: 'Y1', insertSmartArt: 'M', chartMenu: 'C', hyperlink: 'IL', zoomMenu: 'ZM', drawTextbox: 'X', headerFooter: 'H', wordArtMenu: 'W', insertDate: 'D', insertSlideNumber: 'SN', symbol: 'U', insertEquation: 'E', insertVideo: 'V', insertAudio: 'O' },
  design: { themes: 'TH', themeColorsMenu: 'TC', themeFontsMenu: 'TF', bgStylesMenu: 'TB', slideSizeMenu: 'S', formatBg: 'G', designIdeas: 'D' },
  transitions: { previewTransition: 'P', transitions: 'T', transitionOptionsMenu: 'E', transTiming: 'D', applyTransitionAll: 'L' },
  animations: { previewAnim: 'P', animations: 'S', animOptionsMenu: 'M', addAnimMenu: 'AA', animPane: 'C', removeAnim: 'X', animPainter: 'K', triggerMenu: 'G', animTiming: 'T' },
  slideshow: { showFromStart: 'B', showFromCurrent: 'C', customShowMenu: 'W', setupShow: 'S', hideSlide: 'H', rehearse: 'T', presenterView: 'V' },
  review: { spellCheck: 'S', accessibility: 'A', wordCount: 'W', newComment: 'C', deleteComments: 'D', prevComment: 'V', nextComment: 'N', commentsPane: 'P' },
  view: { viewMaster: 'M', toggleRuler: 'R', viewNormal: 'L', viewOutline: 'O', viewSorter: 'I', viewNotesPage: 'T', viewReading: 'D', toggleGrid: 'G', toggleGuides: 'U', toggleNotes: 'N', zoomDialog: 'Q', fitZoom: 'W', selectionPane: 'P' },
  help: { shortcuts: 'K', whatsNew: 'N', about: 'A' },
  shapeFormat: { shapesMini: 'SH', changeShapeMenu: 'E', drawTextbox: 'X', shapeStyles: 'K', shapeFill: 'SF', shapeOutline: 'SO', shapeEffectsMenu: 'SE', wordart: 'Q', objSize: 'H', 'launcher:formatPaneShape': 'DS', ...ARRANGE_KT },
  pictureFormat: { removeBackground: 'E', pictureCorrectionsMenu: 'R', pictureColorMenu: 'I', changePicture: 'CP', resetPicture: 'Q', pictureStyles: 'K', shapeOutline: 'SO', shapeEffectsMenu: 'SE', altText: 'T', cropPicture: 'V', objSize: 'H', 'launcher:formatPaneShape': 'DS', ...ARRANGE_KT },
  tableDesign: { 'tblOpt:firstRow': 'A', 'tblOpt:lastRow': 'T', 'tblOpt:banded': 'R', 'tblOpt:firstCol': 'C', 'tblOpt:lastCol': 'L', tableStyles: 'S', cellFill: 'H', cellBorderMenu: 'B' },
  tableLayout: { tblSelectMenu: 'K', tblDeleteMenu: 'D', tblRowAbove: 'A', tblRowBelow: 'BE', tblColLeft: 'L', tblColRight: 'R', tblMerge: 'M', tblSplit: 'P', tblEqualRows: 'HE', tblEqualCols: 'WE', alignLeft: 'AL', alignCenter: 'AC', alignRight: 'AR', anchorTop: 'AT', anchorMiddle: 'AV', anchorBottom: 'AB' },
  chartDesign: { chartElementsMenu: 'A', chartColorsMenu: 'C', chartData: 'E', chartKindMenu: 'T' },
  slideMaster: { masterRename: 'R', masterReset: 'E', masterHideGraphics: 'H', formatBg: 'B', closeMaster: 'C' },
  smartDesign: { smartAddShape: 'A', smartTextPane: 'X', smartLayoutMenu: 'L', smartColors: 'C', smartToShapes: 'V' },
  playback: { previewMedia: 'P', mediaOptions: 'O', bringForward: 'AF', sendBackward: 'AE', alignMenu: 'AA', altText: 'T', objSize: 'H' },
};
export const itemKey = (it) => (it.cmd ? (it.arg != null ? `${it.cmd}:${it.arg}` : it.cmd) : it.menu ?? it.name ?? it.t);
export function leafItems(tab) {
  const out = [];
  const walk = (it) => {
    if (!it) return;
    if (it.t === 'rows') { for (const r of it.rows) for (const x of r) walk(x); return; }
    if (it.t === 'col') { for (const x of it.items) walk(x); return; }
    out.push(it);
  };
  for (const g of tab.groups) { for (const it of g.items) walk(it); if (g.launcher) { g.launchItem ??= { t: 'launcher', cmd: g.launcher, label: `${g.label} 설정` }; out.push(g.launchItem); } }
  return out;
}
/** 키 팁 정하기: 표에 있는 글자, 없으면 이름 첫 글자 영문 대응 → 남는 글자. 한 탭 안에서 앞부분이 겹치지 않게 (prefix-free) */
function assignKeytips() {
  for (const t of TABS) {
    t.kt = TAB_KT[t.id];
    const table = KT[t.id] ?? {};
    const used = [];
    const ok = (k) => !used.some((u) => u.startsWith(k) || k.startsWith(u));
    const leaves = leafItems(t);
    for (const it of leaves) {
      const k = table[it.t === 'launcher' ? `launcher:${it.cmd}` : itemKey(it)];
      if (k && ok(k)) { it.kt = k; used.push(k); } else it.kt = null;
    }
    const pool = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    for (const it of leaves) {
      if (it.kt) continue;
      let k = null;
      for (const a of pool) { for (const b of pool) { if (ok(a + b)) { k = a + b; break; } } if (k) break; }
      it.kt = k;
      used.push(k);
    }
  }
}
assignKeytips();
export const FILE_KT = 'F';

// ───────────── 빠른 실행 도구 모음 (리본 아래, Alt+1 … Alt+9) ─────────────
export const QAT_CHOICES = [
  { cmd: 'objAlign', arg: 'c', icon: 'alignCenter', label: '가로 가운데 맞춤' },
  { cmd: 'objAlign', arg: 'm', icon: 'alignMiddle', label: '세로 가운데 맞춤' },
  { cmd: 'drawTextbox', icon: 'textbox', label: '텍스트 상자' },
  { cmd: 'eyedropFill', icon: 'eyedropper', label: '스포이트 (도형 채우기 색)' },
  { cmd: 'save', icon: 'save', label: '저장' },
  { cmd: 'undo', icon: 'undo', label: '실행 취소' },
  { cmd: 'redo', icon: 'redo', label: '다시 실행' },
  { cmd: 'showFromStart', icon: 'slideshow', label: '처음부터 시작' },
  { cmd: 'objAlign', arg: 'l', icon: 'alignLeft', label: '왼쪽 맞춤 (개체)' },
  { cmd: 'objAlign', arg: 'r', icon: 'alignRight', label: '오른쪽 맞춤 (개체)' },
  { cmd: 'objAlign', arg: 't', icon: 'alignTop', label: '위쪽 맞춤 (개체)' },
  { cmd: 'objAlign', arg: 'b', icon: 'alignBottom', label: '아래쪽 맞춤 (개체)' },
  { cmd: 'objAlign', arg: 'dh', icon: 'distributeH', label: '가로 간격을 동일하게' },
  { cmd: 'objAlign', arg: 'dv', icon: 'distributeV', label: '세로 간격을 동일하게' },
  { cmd: 'group', icon: 'group', label: '그룹' },
  { cmd: 'ungroup', icon: 'ungroup', label: '그룹 해제' },
  { cmd: 'bringToFront', icon: 'bringForward', label: '맨 앞으로 가져오기' },
  { cmd: 'sendToBack', icon: 'sendBackward', label: '맨 뒤로 보내기' },
  { cmd: 'eyedropFont', icon: 'fontColor', label: '스포이트 (글꼴 색)' },
  { cmd: 'eyedropLine', icon: 'shapeOutline', label: '스포이트 (윤곽선 색)' },
  { cmd: 'formatPainter', icon: 'painter', label: '서식 복사' },
  { cmd: 'newSlide', icon: 'newSlide', label: '새 슬라이드' },
  { cmd: 'print', icon: 'print', label: '인쇄' },
  { cmd: 'open', icon: 'open', label: '열기' },
];
const QAT_DEFAULT = { below: true, items: QAT_CHOICES.slice(0, 4).map(({ cmd, arg }) => ({ cmd, arg })) };
let qatEl = null;
let qatTopEl = null;
export function qatConfig() {
  try {
    const v = JSON.parse(localStorage.getItem('wipoint:qat') ?? 'null');
    if (v && Array.isArray(v.items)) return v;
  } catch { /* 저장소 없음 */ }
  return structuredClone(QAT_DEFAULT);
}
function saveQat(cfg) { try { localStorage.setItem('wipoint:qat', JSON.stringify(cfg)); } catch { /* 저장소 없음 */ } renderQat(); }
const sameQ = (a, b) => a.cmd === b.cmd && (a.arg ?? null) === (b.arg ?? null);
export function qatItemInfo(q) {
  return QAT_CHOICES.find((c) => sameQ(c, q)) ?? (() => { for (const t of TABS) for (const it of leafItems(t)) if (it.cmd === q.cmd && (it.arg ?? null) === (q.arg ?? null)) return { ...q, icon: it.icon, label: it.label ?? it.title, glyph: it.glyph }; return { ...q, icon: 'check', label: q.label ?? q.cmd }; })();
}
export function qatAdd(q) { const cfg = qatConfig(); if (!cfg.items.some((x) => sameQ(x, q))) { cfg.items.push({ cmd: q.cmd, arg: q.arg, label: q.label }); saveQat(cfg); } }
export function qatRemove(i) { const cfg = qatConfig(); cfg.items.splice(i, 1); saveQat(cfg); }
export function qatMove(i, d) { const cfg = qatConfig(); const j = i + d; if (j < 0 || j >= cfg.items.length) return; [cfg.items[i], cfg.items[j]] = [cfg.items[j], cfg.items[i]]; saveQat(cfg); }
export function qatSetBelow(below) { const cfg = qatConfig(); cfg.below = below; saveQat(cfg); }
export function qatReset() { saveQat(structuredClone(QAT_DEFAULT)); }
/** n번째 (0부터) 빠른 실행 명령 실행 — Alt+1 은 0 */
export function qatRun(i) { const q = qatConfig().items[i]; if (!q) return false; run(q.cmd, q.arg); return true; }
/** 키 팁 숫자: 1–9, 그다음 09, 08 … 01 (Office 와 같음) */
export const qatKeytip = (i) => (i < 9 ? String(i + 1) : `0${Math.max(1, 18 - i)}`);

export function initQat(below, top) { qatEl = below; qatTopEl = top; renderQat(); }
export function renderQat() {
  if (!qatEl) return;
  const cfg = qatConfig();
  const host = cfg.below ? qatEl : qatTopEl;
  qatEl.innerHTML = '';
  qatTopEl.innerHTML = '';
  qatEl.hidden = !cfg.below;
  document.getElementById('app')?.classList.toggle('qat-below', !!cfg.below);
  cfg.items.forEach((q, i) => {
    const info = qatItemInfo(q);
    const b = el('button', { class: 'qat-btn', title: `${info.label} (Alt+${i < 9 ? i + 1 : qatKeytip(i)})`, 'data-kt': qatKeytip(i) });
    b.innerHTML = info.glyph ?? iconHtml(info.icon ?? 'check');
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => run(q.cmd, q.arg));
    b.addEventListener('contextmenu', (e) => { e.preventDefault(); openMenu({ x: e.clientX, y: e.clientY }, [
      { label: '빠른 실행 도구 모음에서 제거', icon: 'delete', action: () => qatRemove(i) },
      { label: '왼쪽으로 이동', disabled: i === 0, action: () => qatMove(i, -1) },
      { label: '오른쪽으로 이동', disabled: i === cfg.items.length - 1, action: () => qatMove(i, 1) },
      { sep: true }, ...qatMenuTail(cfg),
    ]); });
    host.append(b);
  });
  const more = el('button', { class: 'qat-btn qat-more', title: '빠른 실행 도구 모음 사용자 지정', 'data-icon': 'chevronDown' });
  more.addEventListener('mousedown', (e) => e.preventDefault());
  more.addEventListener('click', () => openMenu(more, [
    { title: '빠른 실행 도구 모음 사용자 지정' },
    ...QAT_CHOICES.map((c) => ({ label: c.label, checked: cfg.items.some((x) => sameQ(x, c)), action: () => { const idx = cfg.items.findIndex((x) => sameQ(x, c)); if (idx >= 0) qatRemove(idx); else qatAdd(c); } })),
    { sep: true }, ...qatMenuTail(cfg),
  ], { scroll: true }));
  host.append(more);
  if (cfg.below) qatEl.append(el('span', { class: 'qat-hint' }, 'Alt+숫자로 실행 · 리본 단추를 오른쪽 클릭하면 추가'));
  hydrateIcons(host);
}
function qatMenuTail(cfg) {
  return [
    { label: cfg.below ? '리본 위에 표시' : '리본 아래에 표시', action: () => qatSetBelow(!cfg.below) },
    { label: '기본값으로 되돌리기', action: () => qatReset() },
  ];
}

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
  tabsEl.append(el('button', { class: 'ribbon-tab file', 'data-kt': FILE_KT, 'data-kt-level': 'tab', onclick: () => run('backstage') }, '파일'));
  for (const t of visible) {
    tabsEl.append(el('button', { class: `ribbon-tab${t.id === active ? ' active' : ''}${t.context ? ' ctx' : ''}`, 'data-kt': t.kt, 'data-kt-level': 'tab', 'data-tab': t.id, onclick: () => { active = t.id; renderRibbon(); } }, t.label));
  }
  tabsEl.append(el('span', { class: 'spacer' }),
    el('button', { class: 'ribbon-action', 'data-kt': 'ZR', 'data-kt-level': 'tab', onclick: () => run('showFromCurrent'), title: '현재 슬라이드부터 (Shift+F5)' }, el('span', { 'data-icon': 'slideshow' }), el('span', { class: 'ra-l' }, '슬라이드 쇼')),
    el('button', { class: 'ribbon-action primary', 'data-kt': 'ZS', 'data-kt-level': 'tab', onclick: () => run('share') }, el('span', { 'data-icon': 'share' }), el('span', { class: 'ra-l' }, '공유')));
  const tab = TABS.find((t) => t.id === active);
  bodyEl.innerHTML = '';
  for (const g of tab.groups) {
    const gEl = el('div', { class: 'rgroup' }, el('div', { class: 'rgroup-body' }, g.items.map(renderItem)), el('div', { class: 'rgroup-label' }, g.label));
    if (g.launcher) gEl.append(el('button', { class: 'rgroup-launcher', title: `${g.label} 설정`, 'data-kt': g.launchItem?.kt, onclick: () => run(g.launcher) }, '⬊'));
    bodyEl.append(gEl);
  }
  hydrateIcons(tabsEl);
  hydrateIcons(bodyEl);
  redrawKeytips();
}

function iconHtml(name) { return name ? `<span data-icon="${name}"></span>` : ''; }

/** 리본 항목 그리기 + 키 팁(data-kt) · 키 팁 동작(ktAction) · 오른쪽 클릭 [빠른 실행 도구 모음에 추가] */
function renderItem(it) {
  const node = renderItemRaw(it);
  if (!node || !it || it.t === 'rows' || it.t === 'col') return node;
  if (it.kt) node.dataset.kt = it.kt;
  node.ktAction = () => {
    if (it.t === 'font' || it.t === 'size' || it.t === 'transTiming' || it.t === 'animTiming' || it.t === 'objSize') { const f = node.querySelector('input,select'); f?.focus(); f?.select?.(); return; }
    if (it.t === 'gallery') { const m = node.querySelector('.gal-more') ?? node.querySelector('button:not(:disabled)'); if (m) { m.focus(); if (m.classList.contains('gal-more')) m.click(); } return; }
    if (it.t === 'check') { run(it.cmd, it.arg); return; }
    if (it.menu) { run(it.menu, node.querySelector('.caret-only') ?? node); return; }
    run(it.cmd, it.arg);
  };
  if (it.cmd && it.t !== 'gallery') {
    node.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const label = it.label ?? it.title ?? it.cmd;
      openMenu({ x: e.clientX, y: e.clientY }, [{ label: '빠른 실행 도구 모음에 추가', icon: 'plus', action: () => qatAdd({ cmd: it.cmd, arg: it.arg, label: label.replace(/\s*\(.*\)$/, '') }) }]);
    });
  }
  return node;
}

function renderItemRaw(it) {
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
  const list = kind === 'font' ? [...new Set([...(S.pres.fonts ?? []).map((f) => f.typeface), ...FONTS])] : SIZES;
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
