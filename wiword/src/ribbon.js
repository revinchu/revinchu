// 리본 메뉴 (Word 한국어판 배치) — 명령은 state.run 으로 실행
import { S, run, emit } from './state.js';
import { el, hydrateIcons, openMenu } from './ui.js';
import { redrawKeytips } from './keytips.js';
import { fontSections, ensureFont, touchRecent, loadAllLocalFonts, canQueryLocalFonts } from './fontlib.js';

// ───────────── 항목 만들기 ─────────────
const L = (cmd, icon, label, o = {}) => ({ t: 'large', cmd, icon, label, ...o });
const M = (cmd, icon, label, o = {}) => ({ t: 'medium', cmd, icon, label, ...o });
const B = (cmd, icon, title, o = {}) => ({ t: 'small', cmd, icon, title, ...o });
const G = (name, o = {}) => ({ t: 'gallery', name, ...o });
const C = (cmd, label, on, arg) => ({ t: 'check', cmd, label, on, arg });
const ROWS = (...rows) => ({ t: 'rows', rows });
const COL = (...items) => ({ t: 'col', items });

export const FONTS = ['+mj', '+mn', '맑은 고딕', '굴림', '돋움', '바탕', '궁서', '함초롬바탕', '함초롬돋움', '나눔고딕', '나눔명조', '나눔스퀘어', '나눔바른고딕', 'Noto Sans KR', 'Noto Serif KR', 'Pretendard', 'Arial', 'Arial Black', 'Calibri', 'Cambria', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Consolas', 'Courier New', 'Impact'];
export const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

const HOME = { id: 'home', label: '홈', groups: [
  { label: '클립보드', items: [L('paste', 'paste', '붙여넣기', { menu: 'pasteMenu' }), COL(M('cut', 'cut', '잘라내기'), M('copy', 'copy', '복사'), M('formatPainter', 'painter', '서식 복사', { on: 'painter' }))] },
  { label: '글꼴', launcher: 'fontDialog', items: [ROWS(
    [{ t: 'font' }, { t: 'size' }, B('growFont', 'growFont', '글꼴 크기 크게 (Ctrl+Shift+>)'), B('shrinkFont', 'shrinkFont', '글꼴 크기 작게 (Ctrl+Shift+<)'), B(null, 'changeCase', '대/소문자 바꾸기', { menu: 'changeCaseMenu' }), B('clearFormat', 'clearFormat', '모든 서식 지우기')],
    [B('bold', null, '굵게 (Ctrl+B)', { glyph: '<span class="kr b">가</span>', on: 'b' }), B('italic', null, '기울임꼴 (Ctrl+I)', { glyph: '<span class="kr i">가</span>', on: 'i' }), B('underline', null, '밑줄 (Ctrl+U)', { glyph: '<span class="kr u">가</span>', on: 'u', menu: 'underlineMenu' }), B('strike', null, '취소선', { glyph: '<span class="kr s">가</span>', on: 's' }), B('subscript', null, '아래 첨자 (Ctrl+=)', { glyph: '<span class="kr">x<sub>2</sub></span>', on: 'sub' }), B('superscript', null, '위 첨자 (Ctrl+Shift++)', { glyph: '<span class="kr">x<sup>2</sup></span>', on: 'sup' }), B(null, 'wordart', '텍스트 효과', { menu: 'textEffectMenu' }), { t: 'color', cmd: 'highlight', icon: 'highlight', title: '텍스트 강조 색', menu: 'highlightMenu', bar: 'hl' }, { t: 'color', cmd: 'fontColor', icon: 'fontColor', title: '글꼴 색', menu: 'fontColorMenu', bar: 'font' }],
  )] },
  { label: '단락', launcher: 'paragraphDialog', items: [ROWS(
    [B('bullets', 'bullets', '글머리 기호', { menu: 'bulletsMenu', on: 'bullets' }), B('numbering', 'numbering', '번호 매기기', { menu: 'numberingMenu', on: 'numbering' }), B(null, 'multilevel', '다단계 목록', { menu: 'multilevelMenu' }), B('indentLess', 'indentLess', '내어쓰기'), B('indentMore', 'indentMore', '들여쓰기'), B('sortDialog', 'sort', '정렬'), B('toggleMarks', 'marks', '편집 기호 표시/숨기기 (Ctrl+*)', { on: 'marks' })],
    [B('alignLeft', 'alignLeft', '왼쪽 맞춤 (Ctrl+L)', { on: 'al' }), B('alignCenter', 'alignCenter', '가운데 맞춤 (Ctrl+E)', { on: 'ac' }), B('alignRight', 'alignRight', '오른쪽 맞춤 (Ctrl+R)', { on: 'ar' }), B('alignJustify', 'alignJustify', '양쪽 맞춤 (Ctrl+J)', { on: 'aj' }), B('alignDistribute', 'distributeH', '균등 분할', { on: 'ad' }), B(null, 'lineSpacing', '줄 및 단락 간격', { menu: 'lineSpacingMenu' }), { t: 'color', cmd: 'paraShading', icon: 'shading', title: '음영', menu: 'shadingMenu', bar: 'shd' }, B(null, 'borderPara', '테두리', { menu: 'bordersMenu' })],
  )] },
  { label: '스타일', launcher: 'stylesPane', items: [G('styles')] },
  { label: '편집', items: [COL(M('findPane', 'find', '찾기', { menu: 'findMenu' }), M('replaceDialog', 'replace', '바꾸기'), M(null, 'select', '선택', { menu: 'selectMenu' }))] },
  { label: '음성', items: [L('dictate', 'dictate', '받아쓰기', { on: 'dictating' })] },
  { label: 'AI', items: [L('mcpPane', 'mcp', 'AI 연결'), L('aiSummary', 'summary', '요약')] },
] };

const INSERT = { id: 'insert', label: '삽입', groups: [
  { label: '페이지', items: [COL(M(null, 'coverPage', '표지', { menu: 'coverPageMenu' }), M('blankPage', 'blankPage', '빈 페이지'), M('pageBreak', 'pageBreak', '페이지 나누기'))] },
  { label: '표', items: [L(null, 'table', '표', { menu: 'tableMenu' })] },
  { label: '일러스트레이션', items: [L('insertPicture', 'picture', '그림', { menu: 'pictureMenu' }), L('insertIcons', 'icons', '아이콘'), L(null, 'shapes', '도형', { menu: 'shapesMenu' }), L(null, 'chartColumn', '차트', { menu: 'chartMenu' })] },
  { label: '링크', items: [COL(M('insertLinkDialog', 'link', '링크'), M('bookmarkDialog', 'bookmark', '책갈피'), M('crossRefDialog', 'crossRef', '상호 참조'))] },
  { label: '메모', items: [L('newComment', 'newComment', '메모')] },
  { label: '머리글/바닥글', items: [COL(M(null, 'header', '머리글', { menu: 'headerMenu' }), M(null, 'footer', '바닥글', { menu: 'footerMenu' }), M(null, 'pageNumber', '페이지 번호', { menu: 'pageNumberMenu' }))] },
  { label: '텍스트', items: [L(null, 'textbox', '텍스트 상자', { menu: 'textboxMenu' }), COL(M(null, 'quickParts', '빠른 문서 요소', { menu: 'quickPartsMenu' }), M(null, 'dropCap', '드롭 캡', { menu: 'dropCapMenu' }), M('dateTimeDialog', 'dateTime', '날짜 및 시간')), COL(M(null, 'object', '개체', { menu: 'objectMenu' }), M('signatureLine', 'pen', '서명란'))] },
  { label: '기호', items: [COL(M('equationDialog', 'equation', '수식', { menu: 'equationMenu' }), M(null, 'symbol', '기호', { menu: 'symbolMenu' }))] },
] };

const DESIGN = { id: 'design', label: '디자인', groups: [
  { label: '문서 서식', items: [L(null, 'theme', '테마', { menu: 'themeMenu' }), G('styleSets'), COL(M(null, 'fill', '색', { menu: 'themeColorsMenu' }), M(null, 'fontColor', '글꼴', { menu: 'themeFontsMenu' }), M(null, 'lineSpacing', '단락 간격', { menu: 'paraSpacingMenu' })), M('setDefaultFormat', 'check', '기본값으로 설정')] },
  { label: '페이지 배경', items: [L(null, 'watermark', '워터마크', { menu: 'watermarkMenu' }), L(null, 'pageColor', '페이지 색', { menu: 'pageColorMenu' }), L('pageBordersDialog', 'pageBorders', '페이지 테두리')] },
] };

const LAYOUT = { id: 'layout', label: '레이아웃', groups: [
  { label: '페이지 설정', launcher: 'pageSetupDialog', items: [L(null, 'margins', '여백', { menu: 'marginsMenu' }), L(null, 'orientation', '용지 방향', { menu: 'orientationMenu' }), L(null, 'pageSize', '크기', { menu: 'sizeMenu' }), L(null, 'columns', '단', { menu: 'columnsMenu' }), COL(M(null, 'breaks', '나누기', { menu: 'breaksMenu' }), M(null, 'lineNumbers', '줄 번호', { menu: 'lineNumbersMenu' }), M(null, 'hyphen', '하이픈 넣기', { menu: 'hyphenMenu' }))] },
  { label: '원고지', items: [L('manuscriptDialog', 'manuscript', '원고지 설정')] },
  { label: '단락', launcher: 'paragraphDialog', items: [G('paraSpin')] },
  { label: '정렬', items: [L(null, 'position', '위치', { menu: 'positionMenu' }), L(null, 'textWrapping', '텍스트 줄 바꿈', { menu: 'wrapMenu' }), COL(M('bringForward', 'bringForward', '앞으로 가져오기'), M('sendBackward', 'sendBackward', '뒤로 보내기'), M(null, 'align', '맞춤', { menu: 'objAlignMenu' }))] },
] };

const REFS = { id: 'references', label: '참조', groups: [
  { label: '목차', items: [L(null, 'toc', '목차', { menu: 'tocMenu' }), COL(M(null, 'plus', '텍스트 추가', { menu: 'addTextMenu' }), M('updateToc', 'updateToc', '목차 업데이트'))] },
  { label: '각주', launcher: 'notesDialog', items: [L('insertFootnote', 'footnote', '각주 삽입'), COL(M('insertEndnote', 'endnote', '미주 삽입'), M(null, 'next', '다음 각주', { menu: 'nextNoteMenu' }), M('showNotes', 'notes', '메모 표시'))] },
  { label: '인용 및 참고 문헌', items: [L(null, 'citation', '인용 삽입', { menu: 'citationMenu' }), COL(M('manageSources', 'fnBook', '원본 관리'), M(null, 'styles', '스타일', { menu: 'citeStyleMenu' }), M(null, 'stats', '참고 문헌', { menu: 'bibliographyMenu' }))] },
  { label: '캡션', items: [L('captionDialog', 'caption', '캡션 삽입'), COL(M('tableOfFigures', 'toc', '그림 목차 삽입'), M('updateFields', 'refresh', '필드 업데이트'), M('crossRefDialog', 'crossRef', '상호 참조'))] },
  { label: '색인', items: [L('markEntryDialog', 'bookmark', '항목 표시'), COL(M('insertIndex', 'toc', '색인 삽입'))] },
] };

const MAIL = { id: 'mailings', label: '편지', groups: [
  { label: '만들기', items: [L('envelopesDialog', 'envelope', '봉투'), L('labelsDialog', 'labels', '레이블')] },
  { label: '편지 병합 시작', items: [L(null, 'mailMerge', '편지 병합 시작', { menu: 'mergeStartMenu' }), L(null, 'names', '받는 사람 선택', { menu: 'recipientsMenu' }), L('editRecipients', 'table', '받는 사람 목록 편집')] },
  { label: '필드 쓰기 및 삽입', items: [COL(M('highlightMergeFields', 'highlight', '병합 필드 강조 표시'), M('addressBlock', 'envelope', '주소 블록'), M('greetingLine', 'comment', '인사말')), L(null, 'mergeField', '병합 필드 삽입', { menu: 'mergeFieldMenu' })] },
  { label: '결과 미리 보기', items: [L('previewResults', 'preview', '결과 미리 보기', { on: 'mergePreview' }), COL(ROWS([B('mergeFirst', 'prev', '첫 레코드'), B('mergePrev', 'prev', '이전 레코드'), { t: 'gallery', name: 'mergeRecord' }, B('mergeNext', 'next', '다음 레코드'), B('mergeLast', 'next', '마지막 레코드')]))] },
  { label: '마침', items: [L(null, 'check', '완료 및 병합', { menu: 'mergeFinishMenu' })] },
] };

const REVIEW = { id: 'review', label: '검토', groups: [
  { label: '언어 교정', items: [L('spellCheck', 'spell', '맞춤법 및 문법 검사'), COL(M('wordCountDialog', 'wordCount', '단어 개수'))] },
  { label: '음성', items: [L('readAloud', 'audio', '소리 내어 읽기', { on: 'reading' })] },
  { label: '접근성', items: [L('accessibilityCheck', 'accessibility', '접근성 검사')] },
  { label: '언어', items: [L(null, 'translate', '번역', { menu: 'translateMenu' })] },
  { label: '메모', items: [L('newComment', 'newComment', '새 메모'), COL(M(null, 'deleteComment', '삭제', { menu: 'deleteCommentMenu' }), M('prevComment', 'prev', '이전'), M('nextComment', 'next', '다음')), L('commentsPane', 'comment', '메모 표시')] },
  { label: '추적', launcher: 'trackOptionsDialog', items: [L('trackChanges', 'track', '변경 내용 추적', { on: 'track' }), COL(M(null, 'preview', '표시', { menu: 'markupMenu' }), M('revisionsPane', 'selectionPane', '변경 내용 창'))] },
  { label: '변경 내용', items: [L('acceptSel', 'accept', '적용', { menu: 'acceptMenu' }), L('rejectSel', 'reject', '거부', { menu: 'rejectMenu' }), COL(M('prevChange', 'prev', '이전'), M('nextChange', 'next', '다음'))] },
  { label: '비교', items: [L('compareDialog', 'compare', '비교')] },
  { label: '보호', items: [L('restrictEditing', 'protect', '편집 제한')] },
  { label: '도구', items: [L('pdfTools', 'pdf', 'PDF 도구'), L('pdfEdit', 'pdf', 'PDF 편집')] },
] };

const VIEW = { id: 'view', label: '보기', groups: [
  { label: '보기', items: [L('viewRead', 'readMode', '읽기 모드', { on: 'vRead' }), L('viewPrint', 'printLayout', '인쇄 모양', { on: 'vPrint' }), L('viewWeb', 'webLayout', '웹 모양', { on: 'vWeb' }), COL(M('viewOutline', 'outlineView', '개요', { on: 'vOutline' }), M('viewDraft', 'draftView', '초안', { on: 'vDraft' }))] },
  { label: '몰입형', items: [L('focusMode', 'fit', '포커스'), L('immersiveReader', 'readMode', '몰입형 리더')] },
  { label: '표시', items: [COL(C('toggleRuler', '눈금자', 'ruler'), C('toggleGridlines', '눈금선', 'gridlines'), C('toggleNav', '탐색 창', 'nav'))] },
  { label: '확대/축소', items: [L('zoomDialog', 'zoom100', '확대/축소'), L('zoom100', 'zoom100', '100%'), COL(M('zoomOnePage', 'onePage', '한 페이지'), M('zoomMulti', 'multiPage', '여러 페이지'), M('zoomPageWidth', 'pageWidth', '페이지 너비'))] },
  { label: '창', items: [L('newWindow', 'duplicate', '새 창'), L('markdownPane', 'markdown', '마크다운'), L('splitView', 'freeze', '나누기')] },
] };

const HELP = { id: 'help', label: '도움말', groups: [
  { label: '도움말', items: [L('helpPane', 'about', '도움말'), L('shortcuts', 'keyboard', '바로 가기 키'), L('whatsNew', 'info', '새로운 기능'), L('about', 'about', 'WIWORD 정보')] },
] };

// ── 상황별 탭 ──
const TABLE_DESIGN = { id: 'tableDesign', label: '테이블 디자인', context: 'table', groups: [
  { label: '표 스타일 옵션', items: [COL(C('tableLook', '머리글 행', 'lkFirstRow', 'firstRow'), C('tableLook', '요약 행', 'lkLastRow', 'lastRow'), C('tableLook', '줄무늬 행', 'lkBandRows', 'bandRows')), COL(C('tableLook', '첫째 열', 'lkFirstCol', 'firstCol'), C('tableLook', '마지막 열', 'lkLastCol', 'lastCol'), C('tableLook', '줄무늬 열', 'lkBandCols', 'bandCols'))] },
  { label: '표 스타일', items: [G('tableStyles'), { t: 'color', cmd: 'cellShading', icon: 'fill', title: '음영', label: '음영', menu: 'cellShadingMenu', bar: 'cellShd' }] },
  { label: '테두리', items: [L(null, 'borderAll', '테두리', { menu: 'cellBordersMenu' })] },
] };
const TABLE_LAYOUT = { id: 'tableLayout', label: '표 레이아웃', context: 'table', groups: [
  { label: '표', items: [COL(M(null, 'select', '선택', { menu: 'tableSelectMenu' }), M('toggleGridlines', 'gridlines', '눈금선 보기', { on: 'gridlines' }), M('tablePropsDialog', 'format', '속성'))] },
  { label: '행 및 열', items: [L(null, 'delete', '삭제', { menu: 'tableDeleteMenu' }), L('insertRowAbove', 'rowInsert', '위에 삽입'), COL(M('insertRowBelow', 'rowInsert', '아래에 삽입'), M('insertColLeft', 'colInsert', '왼쪽에 삽입'), M('insertColRight', 'colInsert', '오른쪽에 삽입'))] },
  { label: '병합', items: [COL(M('mergeCells', 'merge', '셀 병합'), M('splitCellDialog', 'splitCells', '셀 분할'), M('splitTable', 'breaks', '표 분할'))] },
  { label: '셀 크기', items: [L(null, 'autofit', '자동 맞춤', { menu: 'autofitMenu' }), G('cellSize'), COL(M('distributeRows', 'rowHeight', '행 높이를 같게'), M('distributeCols', 'colWidth', '열 너비를 같게'))] },
  { label: '맞춤', items: [ROWS(
    [B('cellAlignTL', 'alignTop', '위쪽 왼쪽'), B('cellAlignTC', 'alignTop', '위쪽 가운데'), B('cellAlignTR', 'alignTop', '위쪽 오른쪽')],
    [B('cellAlignML', 'alignMiddle', '가운데 왼쪽'), B('cellAlignMC', 'alignMiddle', '가운데'), B('cellAlignMR', 'alignMiddle', '가운데 오른쪽')],
    [B('cellAlignBL', 'alignBottom', '아래쪽 왼쪽'), B('cellAlignBC', 'alignBottom', '아래쪽 가운데'), B('cellAlignBR', 'alignBottom', '아래쪽 오른쪽')],
  ), L('cellTextDir', 'textDirection', '텍스트 방향'), L('cellMarginsDialog', 'cellMargins', '셀 여백')] },
  { label: '데이터', items: [L('sortTableDialog', 'sort', '정렬'), COL(M('repeatHeader', 'repeatHeader', '머리글 행 반복', { on: 'repeatHeader' }), M('tableToText', 'tableConvert', '텍스트로 변환'), M('tableFormulaDialog', 'fx', '수식'))] },
] };
const PICTURE_FMT = { id: 'pictureFormat', label: '그림 서식', context: 'picture', groups: [
  { label: '조정', items: [L(null, 'bright', '수정', { menu: 'picCorrectionsMenu' }), L(null, 'grayscale', '색', { menu: 'picColorMenu' }), COL(M('changePicture', 'swap', '그림 바꾸기'), M('resetPicture', 'resetPic', '그림 원래대로'))] },
  { label: '그림 스타일', items: [G('pictureStyles'), COL(M(null, 'picBorder', '그림 테두리', { menu: 'picBorderMenu' }), M(null, 'effects', '그림 효과', { menu: 'picEffectsMenu' }))] },
  { label: '접근성', items: [L('altTextDialog', 'accessibility', '대체 텍스트')] },
  { label: '정렬', items: [L(null, 'position', '위치', { menu: 'positionMenu' }), L(null, 'textWrapping', '텍스트 줄 바꿈', { menu: 'wrapMenu' }), COL(M('bringForward', 'bringForward', '앞으로 가져오기'), M('sendBackward', 'sendBackward', '뒤로 보내기'), M(null, 'rotate', '회전', { menu: 'rotateMenu' }))] },
  { label: '크기', items: [L('cropPicture', 'crop', '자르기'), G('picSize')] },
] };
const HF_TAB = { id: 'headerFooter', label: '머리글/바닥글', context: 'hf', groups: [
  { label: '머리글/바닥글', items: [COL(M(null, 'header', '머리글', { menu: 'headerMenu' }), M(null, 'footer', '바닥글', { menu: 'footerMenu' }), M(null, 'pageNumber', '페이지 번호', { menu: 'pageNumberMenu' }))] },
  { label: '삽입', items: [L('dateTimeDialog', 'dateTime', '날짜 및 시간'), L('docInfoMenu', 'info', '문서 정보', { menu: 'docInfoMenu' }), L('insertPicture', 'picture', '그림')] },
  { label: '탐색', items: [L('hfGotoHeader', 'header', '머리글로 이동'), L('hfGotoFooter', 'footer', '바닥글로 이동')] },
  { label: '옵션', items: [COL(C('differentFirstPage', '첫 페이지를 다르게 지정', 'titlePg'), C('differentOddEven', '홀수/짝수 페이지를 다르게 지정', 'evenOdd'), C('toggleShowDocText', '문서 텍스트 표시', 'showDocText'))] },
  { label: '위치', items: [G('hfPos')] },
  { label: '닫기', items: [L('closeHeaderFooter', 'close', '머리글/바닥글 닫기')] },
] };

export const TABS = [HOME, INSERT, DESIGN, LAYOUT, REFS, MAIL, REVIEW, VIEW, HELP, TABLE_DESIGN, TABLE_LAYOUT, PICTURE_FMT, HF_TAB];

// ───────────── 키 팁 (Alt) — Word 한국어판과 같은 글자 ─────────────
const TAB_KT = { home: 'H', insert: 'N', design: 'G', layout: 'P', references: 'S', mailings: 'M', review: 'R', view: 'W', help: 'Y2', tableDesign: 'JT', tableLayout: 'JL', pictureFormat: 'JP', headerFooter: 'JH' };
const KT = {
  home: { paste: 'V', cut: 'X', copy: 'C', formatPainter: 'FP', font: 'FF', size: 'FS', growFont: 'FG', shrinkFont: 'FK', changeCaseMenu: '7', clearFormat: 'E', bold: '1', italic: '2', underline: '3', strike: '4', subscript: '5', superscript: '6', textEffectMenu: 'FT', highlight: 'I', fontColor: 'FC', bullets: 'U', numbering: 'N', multilevelMenu: 'M', indentLess: 'AO', indentMore: 'AI', sortDialog: 'SO', toggleMarks: '8', alignLeft: 'AL', alignCenter: 'AC', alignRight: 'AR', alignJustify: 'AJ', alignDistribute: 'AD', lineSpacingMenu: 'K', paraShading: 'H', bordersMenu: 'B', styles: 'L', findPane: 'FD', replaceDialog: 'R', selectMenu: 'SL', dictate: 'D2', mcpPane: 'Y1', aiSummary: 'Y3', 'launcher:fontDialog': 'FN', 'launcher:paragraphDialog': 'PG', 'launcher:stylesPane': 'FY' },
  insert: { coverPageMenu: 'V', blankPage: 'NP', pageBreak: 'B', tableMenu: 'T', insertPicture: 'P', insertIcons: 'Y1', shapesMenu: 'SH', chartMenu: 'C', insertLinkDialog: 'I', bookmarkDialog: 'K', crossRefDialog: 'RF', newComment: 'L', headerMenu: 'H', footerMenu: 'O', pageNumberMenu: 'NU', textboxMenu: 'X', quickPartsMenu: 'Q', dropCapMenu: 'RC', dateTimeDialog: 'D', objectMenu: 'J', signatureLine: 'G', equationDialog: 'E', symbolMenu: 'U' },
  design: { themeMenu: 'TH', styleSets: 'S', themeColorsMenu: 'TC', themeFontsMenu: 'TF', paraSpacingMenu: 'PS', setDefaultFormat: 'D', watermarkMenu: 'W', pageColorMenu: 'PC', pageBordersDialog: 'PB' },
  layout: { marginsMenu: 'M', orientationMenu: 'O', sizeMenu: 'SZ', columnsMenu: 'J', breaksMenu: 'B', lineNumbersMenu: 'LN', hyphenMenu: 'H', manuscriptDialog: 'G', paraSpin: 'I', positionMenu: 'PO', wrapMenu: 'TW', bringForward: 'AF', sendBackward: 'AE', objAlignMenu: 'AA', 'launcher:pageSetupDialog': 'SP', 'launcher:paragraphDialog': 'PG' },
  references: { tocMenu: 'T', addTextMenu: 'A', updateToc: 'U', insertFootnote: 'F', insertEndnote: 'E', nextNoteMenu: 'O', showNotes: 'H', citationMenu: 'C', manageSources: 'M', citeStyleMenu: 'L', bibliographyMenu: 'B', captionDialog: 'P', tableOfFigures: 'G', updateFields: 'V', crossRefDialog: 'RF', markEntryDialog: 'N', insertIndex: 'X' },
  mailings: { envelopesDialog: 'E', labelsDialog: 'L', mergeStartMenu: 'S', recipientsMenu: 'R', editRecipients: 'D', highlightMergeFields: 'H', addressBlock: 'A', greetingLine: 'G', mergeFieldMenu: 'I', previewResults: 'P', mergeFinishMenu: 'F' },
  review: { spellCheck: 'S', wordCountDialog: 'W', readAloud: 'RA', accessibilityCheck: 'A1', translateMenu: 'L', newComment: 'C', deleteCommentMenu: 'D', prevComment: 'V', nextComment: 'N', commentsPane: 'P', trackChanges: 'G', markupMenu: 'TM', revisionsPane: 'TP', acceptSel: 'A2', rejectSel: 'J', prevChange: 'F', nextChange: 'H', compareDialog: 'M', restrictEditing: 'PE' },
  view: { viewRead: 'F', viewPrint: 'P', viewWeb: 'L', viewOutline: 'U', viewDraft: 'E', focusMode: 'O', immersiveReader: 'L2', toggleRuler: 'R', toggleGridlines: 'G', toggleNav: 'K', zoomDialog: 'Q', zoom100: 'J', zoomOnePage: '1', zoomMulti: 'M', zoomPageWidth: 'I', newWindow: 'N', markdownPane: 'MD', splitView: 'S' },
  help: { helpPane: 'H', shortcuts: 'K', whatsNew: 'N', about: 'A' },
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
  { cmd: 'save', icon: 'save', label: '저장' },
  { cmd: 'undo', icon: 'undo', label: '실행 취소' },
  { cmd: 'redo', icon: 'redo', label: '다시 실행' },
  { cmd: 'print', icon: 'print', label: '인쇄 미리 보기 및 인쇄' },
  { cmd: 'open', icon: 'open', label: '열기' },
  { cmd: 'newBlank', icon: 'newDoc', label: '새 문서' },
  { cmd: 'formatPainter', icon: 'painter', label: '서식 복사' },
  { cmd: 'toggleMarks', icon: 'marks', label: '편집 기호 표시' },
  { cmd: 'insertTableDialog', icon: 'table', label: '표 삽입' },
  { cmd: 'insertPicture', icon: 'picture', label: '그림 삽입' },
  { cmd: 'insertLinkDialog', icon: 'link', label: '링크' },
  { cmd: 'newComment', icon: 'newComment', label: '새 메모' },
  { cmd: 'trackChanges', icon: 'track', label: '변경 내용 추적' },
  { cmd: 'readAloud', icon: 'audio', label: '소리 내어 읽기' },
  { cmd: 'spellCheck', icon: 'spell', label: '맞춤법 검사' },
  { cmd: 'exportPdf', icon: 'pdf', label: 'PDF 로 내보내기' },
  { cmd: 'saveHwpx', icon: 'hwp', label: '한글(HWPX)로 저장' },
  { cmd: 'markdownPane', icon: 'markdown', label: '마크다운 창' },
  { cmd: 'mcpPane', icon: 'mcp', label: 'AI 연결 (MCP)' },
];
const QAT_DEFAULT = { below: true, items: QAT_CHOICES.slice(0, 4).map(({ cmd, arg }) => ({ cmd, arg })) };
let qatEl = null;
let qatTopEl = null;
export function qatConfig() {
  try {
    const v = JSON.parse(localStorage.getItem('wiword:qat') ?? 'null');
    if (v && Array.isArray(v.items)) return v;
  } catch { /* 저장소 없음 */ }
  return structuredClone(QAT_DEFAULT);
}
function saveQat(cfg) { try { localStorage.setItem('wiword:qat', JSON.stringify(cfg)); } catch { /* 저장소 없음 */ } renderQat(); }
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
    el('button', { class: 'ribbon-action', 'data-kt': 'ZC', 'data-kt-level': 'tab', onclick: () => run('togglePane', 'comments'), title: '메모' }, el('span', { 'data-icon': 'comment' }), el('span', { class: 'ra-l' }, '메모')),
    el('button', { class: 'ribbon-action', 'data-kt': 'ZA', 'data-kt-level': 'tab', onclick: () => run('togglePane', 'mcp'), title: 'AI 연결 (MCP)' }, el('span', { 'data-icon': 'mcp' }), el('span', { class: 'ra-l' }, 'AI 연결')),
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

// ───── 글꼴 목록 (PowerPoint: 테마 글꼴 · 최근 · 모든 글꼴, 글꼴 이름을 그 글꼴로) ─────
let fontPop = null;
function closeFontPop() { fontPop?.remove(); fontPop = null; removeEventListener('pointerdown', outsideFontPop, true); }
function outsideFontPop(e) { if (fontPop && !fontPop.contains(e.target) && !e.target.closest('.rfont')) closeFontPop(); }
function fontDropButton(inp) {
  const btn = el('button', { class: 'rbtn caret-only font-drop', title: '글꼴 목록', 'data-icon': 'chevronDown', type: 'button' });
  const open = async (filter = '') => {
    closeFontPop();
    const r = inp.getBoundingClientRect();
    fontPop = el('div', { class: 'font-pop', role: 'listbox' });
    Object.assign(fontPop.style, { left: `${r.left}px`, top: `${r.bottom + 2}px` });
    const draw = () => {
      fontPop.replaceChildren();
      if (canQueryLocalFonts()) fontPop.append(el('button', { class: 'font-all', type: 'button', onclick: async () => { const l = await loadAllLocalFonts({ ask: true }); if (!l) { fontPop?.append(el('div', { class: 'font-note' }, '브라우저가 글꼴 목록을 허용하지 않았습니다')); return; } draw(); } }, '이 컴퓨터의 모든 글꼴 불러오기'));
      // 글꼴 파일 (.ttf · .otf · .woff · .woff2) 불러오기 · 파일에 글꼴 포함
      fontPop.append(el('div', { class: 'font-tools' },
        el('button', { class: 'font-all', type: 'button', onclick: () => run('importFonts', () => draw()) }, '글꼴 파일 불러오기… (.ttf · .otf · .woff)'),
        el('label', { class: 'font-embed', title: '저장할 때 문서에 쓴 TrueType 글꼴을 pptx 에 넣습니다 (PowerPoint [파일의 글꼴 포함])' }, el('input', { type: 'checkbox', checked: !!S.doc.props?.embedFonts, onchange: () => run('embedFonts') }), ' 저장할 때 파일에 글꼴 포함')));
      for (const [title, items] of fontSections(S.doc, filter)) {
        fontPop.append(el('div', { class: 'font-sec' }, title));
        for (const it of items.slice(0, 400)) {
          const row = el('button', { class: 'font-item', type: 'button', title: it.web ? `${it.label} — 고르면 인터넷에서 내려받아 씁니다` : it.label, style: { fontFamily: `'${it.face}', '맑은 고딕', sans-serif` } }, it.label, it.note ? el('small', {}, it.note) : null);
          row.addEventListener('mousedown', (e) => e.preventDefault());
          row.addEventListener('click', () => { closeFontPop(); touchRecent(it.name); run('fontFamily', it.name); });
          if (it.web) row.dataset.web = it.name;
          fontPop.append(row);
        }
      }
      // 보이는 웹 글꼴은 미리 보기용으로 내려받음
      const io = new IntersectionObserver((ents) => { for (const en of ents) if (en.isIntersecting) { ensureFont(en.target.dataset.web); io.unobserve(en.target); } }, { root: fontPop });
      for (const n of fontPop.querySelectorAll('[data-web]')) io.observe(n);
    };
    draw();
    document.body.append(fontPop);
    addEventListener('pointerdown', outsideFontPop, true);
    // 이 컴퓨터의 모든 글꼴: 처음 열 때 한 번 물어봄 (사용자 동작 안에서만 가능)
    if (canQueryLocalFonts() && !open.asked) { open.asked = true; const l = await loadAllLocalFonts({ ask: true }); if (l && fontPop) draw(); }
  };
  btn.addEventListener('click', () => (fontPop ? closeFontPop() : open('')));
  inp.addEventListener('input', () => open(inp.value.replace(/\s*\((제목|본문)\)$/, '')));
  inp.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeFontPop(); if (e.key === 'ArrowDown' && fontPop) { e.preventDefault(); fontPop.querySelector('.font-item')?.focus(); } });
  return btn;
}

function fontBox(kind) {
  const val = kind === 'font' ? state.font : state.size;
  const list = kind === 'font' ? [...new Set([...(S.doc.fonts ?? []).map((f) => f.typeface), ...FONTS])] : SIZES;
  const id = `dl-${kind}`;
  const shown = kind === 'font' ? (val === '+mj' ? `${S.doc.themeFonts.major.ea} (제목)` : val === '+mn' ? `${S.doc.themeFonts.minor.ea} (본문)` : val ?? '') : val == null ? '' : String(val);
  const inp = el('input', { class: `rselect ${kind === 'font' ? 'font-family' : 'font-size'}`, value: shown, list: kind === 'font' ? undefined : id, title: kind === 'font' ? '글꼴' : '글꼴 크기', spellcheck: 'false' });
  const dl = kind === 'font' ? fontDropButton(inp) : el('datalist', { id }, list.map((f) => el('option', { value: String(f) })));
  const apply = () => {
    let v = inp.value.trim();
    if (!v) return;
    if (kind === 'font') {
      if (v.endsWith('(제목)')) v = '+mj';
      else if (v.endsWith('(본문)')) v = '+mn';
      closeFontPop();
      touchRecent(v);
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
