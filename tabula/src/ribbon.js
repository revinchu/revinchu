// 리본 메뉴 정의 및 렌더링
import { ICONS } from './icons.js';
import { el } from './ui.js';
import { NUMBER_FORMATS } from './format.js';

export const FONTS = ['맑은 고딕', '굴림', '돋움', '바탕', '궁서', 'Arial', 'Calibri', 'Consolas', 'Times New Roman', 'Verdana'];
export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

// ── 항목 생성기 ──
const large = (cmd, icon, label, opts = {}) => ({ type: 'large', cmd, icon, label, ...opts });
const medium = (cmd, icon, label, opts = {}) => ({ type: 'medium', cmd, icon, label, ...opts });
const btn = (cmd, icon, title, opts = {}) => ({ type: 'btn', cmd, icon, title, ...opts });
const kr = (cls, text, sup) => `<span class="kr ${cls}">${text}${sup ?? ''}</span>`;
const row = (...items) => ({ type: 'row', items });
const col = (...items) => ({ type: 'col', items });
const sep = { type: 'sep' };
const group = (label, items, launcher) => ({ label, items, launcher });
const check = (cmd, label, stateKey) => ({ type: 'check', cmd, label, stateKey });

export const TABS = [
  { id: 'file', label: '파일', file: true },
  {
    id: 'home', label: '홈', groups: [
      group('클립보드', [
        large('paste', 'paste', '붙여넣기', { menu: 'paste' }),
        col(
          medium('cut', 'cut', '잘라내기', { title: '잘라내기 (Ctrl+X)' }),
          medium('copy', 'copy', '복사', { title: '복사 (Ctrl+C)' }),
          medium('painter', 'painter', '서식 복사', { title: '서식 복사 (두 번 클릭하면 계속 적용)', toggle: 'painter' }),
        ),
      ]),
      group('글꼴', [
        col(
          row(
            { type: 'font', cmd: 'fontFamily', cls: 'font-family', stateKey: 'font', title: '글꼴 (이 PC에 설치된 글꼴 이름을 입력하거나 목록에서 선택)', menu: 'fontList' },
            { type: 'select', cmd: 'fontSize', cls: 'font-size', options: FONT_SIZES.map((s) => ({ value: String(s), label: String(s) })), stateKey: 'size', editable: true, title: '글꼴 크기' },
            btn('growFont', kr('', '가', '<sup>▲</sup>'), '글꼴 크기 크게'),
            btn('shrinkFont', kr('', '가', '<sub>▼</sub>'), '글꼴 크기 작게'),
          ),
          row(
            btn('bold', kr('b', '가'), '굵게 (Ctrl+B)', { toggle: 'bold' }),
            btn('italic', kr('i', '가'), '기울임꼴 (Ctrl+I)', { toggle: 'italic' }),
            btn('underline', kr('u', '가'), '밑줄 (Ctrl+U)', { toggle: 'underline' }),
            btn('strike', kr('s', '가'), '취소선 (Ctrl+5)', { toggle: 'strike' }),
            sep,
            btn('borderLast', 'border', '테두리', { menu: 'borders' }),
            sep,
            { type: 'color', cmd: 'fillColor', icon: 'fill', title: '채우기 색', menu: 'fillColor', stateKey: 'lastFill' },
            { type: 'color', cmd: 'fontColor', icon: 'fontColor', title: '글꼴 색', menu: 'fontColor', stateKey: 'lastFont' },
          ),
        ),
      ], 'fontDialog'),
      group('맞춤', [
        col(
          row(
            btn('valignTop', 'alignTop', '위쪽 맞춤', { toggle: 'valignTop' }),
            btn('valignMiddle', 'alignMiddle', '가운데 맞춤', { toggle: 'valignMiddle' }),
            btn('valignBottom', 'alignBottom', '아래쪽 맞춤', { toggle: 'valignBottom' }),
            sep,
            medium('wrap', 'wrap', '자동 줄 바꿈', { toggle: 'wrap' }),
          ),
          row(
            btn('alignLeft', 'alignLeft', '왼쪽 맞춤', { toggle: 'alignLeft' }),
            btn('alignCenter', 'alignCenter', '가운데 맞춤', { toggle: 'alignCenter' }),
            btn('alignRight', 'alignRight', '오른쪽 맞춤', { toggle: 'alignRight' }),
            sep,
            btn('indentDec', 'indentDec', '내어쓰기'),
            btn('indentInc', 'indentInc', '들여쓰기'),
            sep,
            medium('mergeCenter', 'merge', '병합하고 가운데 맞춤', { menu: 'merge', toggle: 'merged' }),
          ),
        ),
      ]),
      group('표시 형식', [
        col(
          row({ type: 'combo', cls: 'numfmt', options: NUMBER_FORMATS, stateKey: 'numFmt', menu: 'numFormats', title: '표시 형식' }),
          row(
            btn('fmtCurrency', 'currency', '회계 표시 형식'),
            btn('fmtPercent', 'percent', '백분율 스타일 (Ctrl+Shift+%)'),
            btn('fmtComma', 'comma', '쉼표 스타일'),
            sep,
            btn('incDecimal', 'incDecimal', '자릿수 늘림'),
            btn('decDecimal', 'decDecimal', '자릿수 줄임'),
          ),
        ),
      ]),
      group('스타일', [
        large('condFormat', 'condFormat', '조건부 서식', { menu: 'condFormat' }),
        large('tableStyle', 'table', '표 서식', { menu: 'tableStyles' }),
        large('cellStyle', 'cellStyles', '셀 스타일', { menu: 'cellStyles' }),
      ]),
      group('빠른 서식', [
        col(
          medium('cfUpDown', '<span class="qf-ic"><b style="color:#e00">▲</b><b style="color:#06c">▼</b></span>', '증감 ▲▼', { title: '증감 표시 (조건부 서식): 0보다 크면 빨강 ▲, 작으면 파랑 ▼\n[빨강][>0]"▲"#,##0;[파랑]"▼"#,##0;\n셀 하나만 선택하면 데이터 덩어리 전체에 적용' }),
          medium('cfUpDownPct', '<span class="qf-ic"><b style="color:#e00">▲</b><b style="color:#06c">%</b></span>', '증감률 ▲▼%', { title: '증감률 표시 (조건부 서식)\n[빨강]"▲"#,##0.00%;[파랑]"▼"#,##0.00%' }),
          medium('cfWeekend', '<span class="qf-ic"><b style="color:#06c">토</b><b style="color:#e00">일</b></span>', '주말 색', { title: '주말 행 색 (조건부 서식): 토요일 날짜가 든 행은 파랑, 일요일은 빨강' }),
        ),
      ]),
      group('셀', [
        large('insertMenu', 'insert', '삽입', { menu: 'insert' }),
        large('deleteMenu', 'delete', '삭제', { menu: 'delete' }),
        large('formatMenu', 'format', '서식', { menu: 'format' }),
      ]),
      group('편집', [
        col(
          medium('autosum', 'autosum', '자동 합계', { menu: 'autosum', split: true }),
          medium('fillMenu', 'fillDown', '채우기', { menu: 'fill' }),
          medium('clearMenu', 'clear', '지우기', { menu: 'clear' }),
        ),
        large('sortMenu', 'sort', '정렬 및 필터', { menu: 'sort' }),
        large('findMenu', 'find', '찾기 및 선택', { menu: 'find' }),
      ]),
    ],
  },
  {
    id: 'insert', label: '삽입', groups: [
      group('셀', [
        large('insertRows', 'rowInsert', '시트 행 삽입'),
        large('insertCols', 'colInsert', '시트 열 삽입'),
        large('addSheet', 'sheetInsert', '시트 삽입'),
      ]),
      group('표', [
        large('insertPivot', 'pivot', '피벗 테이블'),
        large('recommendPivot', 'recommendPivot', '추천 피벗 테이블', { title: '데이터를 요약하는 피벗 테이블 후보를 미리 보고 고르기' }),
        large('createTable', 'table', '표', { title: '표 만들기 (Ctrl+T, Ctrl+L)' }),
      ]),
      group('일러스트레이션', [
        large('insertPicture', 'picture', '그림', { menu: 'picture', title: '이 기기의 그림 삽입 — 셀에 배치 또는 셀 위에 배치 (붙여넣기·끌어 놓기도 가능)' }),
        large('shapesMenu', 'shapes', '도형', { menu: 'shapes' }),
        large('insertIcons', 'iconsLib', '아이콘', { title: '아이콘 삽입 — 34개 범주 3,600여 개 (엑셀 아이콘과 같은 그림, 색 변경 가능)' }),
        large('insertTextbox', 'textbox', '텍스트 상자'),
      ]),
      group('기호', [large('insertEquation', 'equation', '수식', { menu: 'equations' }), large('insertSymbol', 'symbol', '기호')]),
      group('셀 컨트롤', [large('insertCheckbox', '<svg viewBox="0 0 24 24" width="24" height="24"><rect x="3" y="3" width="18" height="18" rx="3" fill="#fff" stroke="#217346" stroke-width="2"/><path d="M7 12l3.5 3.5L17 9" fill="none" stroke="#217346" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>', '확인란', { title: '확인란 삽입 — 선택한 칸에 체크 상자 (값 TRUE / FALSE, 클릭 · Space 로 켜고 끔)' })]),
      group('링크', [large('hyperlink', 'link', '링크', { title: '하이퍼링크 삽입 (Ctrl+K)' })]),
      group('필터', [large('insertSlicer', 'slicer', '슬라이서', { title: '표나 피벗 테이블에 슬라이서 삽입' }), large('insertTimeline', 'calendar', '시간 표시 막대', { title: '날짜 필드를 기간(연 · 분기 · 월 · 일)으로 거르는 시간 표시 막대 삽입' })]),
      group('차트', [
        large('insertChartAll', 'chartColumn', '추천 차트', { title: '차트 삽입 — 추천 차트 · 모든 차트' }),
        col(
          row(btn('chartsColBar', 'chartColumn', '세로 또는 가로 막대형 차트 삽입', { menu: 'chartsColBar' }), btn('chartsHier', 'table', '계층 구조 차트 삽입', { menu: 'chartsHier' }), btn('chartsWaterfall', 'chartBar', '폭포, 깔때기형, 주식형 차트 삽입', { menu: 'chartsWaterfall' })),
          row(btn('chartsLineArea', 'chartLine', '꺾은선형 또는 영역형 차트 삽입', { menu: 'chartsLineArea' }), btn('chartsStat', 'stats', '통계 차트 삽입', { menu: 'chartsStat' }), btn('chartsCombo', 'chartArea', '콤보 차트 삽입', { menu: 'chartsCombo' })),
          row(btn('chartPie', 'chartPie', '원형 또는 도넛형 차트 삽입', { menu: 'pieCharts' }), btn('chartsScatter', 'chartScatter', '분산형(X, Y) 또는 거품형 차트 삽입', { menu: 'chartsScatter' })),
        ),
        large('insertPivotChart', 'pivot', '피벗 차트', { title: '피벗 차트 삽입 (피벗 테이블 + 차트)' }),
      ]),
      group('스파크라인', [
        medium('sparkLine', 'chartLine', '꺾은선형', { title: '꺾은선형 스파크라인 (셀 안 차트)' }),
        medium('sparkColumn', 'chartColumn', '열', { title: '열 스파크라인' }),
        medium('sparkWinLoss', 'chartBar', '승패', { title: '승패 스파크라인' }),
      ]),
      group('함수', [large('insertFunction', 'function', '함수')]),
      group('날짜 및 시간', [
        large('insertDate', 'calendar', '오늘 날짜'),
        large('insertTime', 'clock', '현재 시간'),
      ]),
      group('메모', [large('editComment', 'newComment', '메모')]),
    ],
  },
  {
    id: 'layout', label: '페이지 레이아웃', groups: [
      group('테마', [large('themesBtn', 'effects', '테마', { menu: 'themes' }), col(medium('themeColorsBtn', 'fill', '색', { menu: 'themeColors' }))]),
      group('페이지 설정', [
        large('marginsBtn', 'borderOutside', '여백', { menu: 'marginsMenu' }),
        large('orientBtn', 'print', '용지 방향', { menu: 'orientMenu' }),
        large('paperBtn', 'table', '크기', { menu: 'paperMenu' }),
        large('printAreaBtn', 'freeze', '인쇄 영역', { menu: 'printAreaMenu' }),
        large('printTitles', 'rowInsert', '인쇄 제목', { title: '반복할 행 · 머리글/바닥글 · 페이지 설정' }),
        large('pageSetup', 'format', '페이지 설정'),
      ]),
      group('크기 조정', [large('fitBtn', 'zoomIn', '배율 조정', { menu: 'fitMenu' })]),
      group('인쇄', [large('print', 'print', '인쇄')]),
      group('시트', [
        large('hideRowsCols', 'hide', '숨기기 및 숨기기 취소', { menu: 'hideMenu' }),
      ]),
      group('시트 옵션', [
        col(
          check('toggleGrid', '눈금선 보기', 'showGrid'),
          check('togglePrintGrid', '눈금선 인쇄', 'printGrid'),
        ),
      ]),
    ],
  },
  {
    id: 'formulas', label: '수식', groups: [
      group('함수 라이브러리', [
        large('insertFunction', 'function', '함수 삽입'),
        large('autosum', 'autosum', '자동 합계', { menu: 'autosum', split: true }),
        large('fnCat', 'fnBook', '재무', { menu: 'fn:재무' }),
        large('fnCat', 'fnBook', '논리', { menu: 'fn:논리' }),
        large('fnCat', 'fnBook', '텍스트', { menu: 'fn:텍스트' }),
        large('fnCat', 'fnBook', '날짜 및 시간', { menu: 'fn:날짜/시간' }),
        large('fnCat', 'fnBook', '찾기/참조 영역', { menu: 'fn:찾기/참조' }),
        large('fnCat', 'fnBook', '수학/삼각', { menu: 'fn:수학/삼각' }),
        large('fnCat', 'fnBook', '함수 더 보기', { menu: 'fn:more' }),
      ]),
      group('정의된 이름', [
        large('nameManager', 'names', '이름 관리자', { title: '이름 관리자 (Ctrl+F3)' }),
        medium('defineName', 'nameDefine', '이름 정의'),
        medium('useInFormula', 'fx', '수식에서 사용', { menu: 'useInFormula' }),
        medium('createNamesFromSel', 'table', '선택 영역에서 만들기', { title: '선택 영역에서 이름 만들기 (Ctrl+Shift+F3)' }),
      ]),
      group('수식 분석', [
        col(
          medium('tracePrecedents', 'prev', '참조되는 셀 추적'),
          medium('traceDependents', 'next', '참조하는 셀 추적'),
          medium('removeArrows', 'clear', '연결선 제거', { menu: 'arrowsMenu', split: true }),
        ),
        col(
          medium('toggleFormulas', 'showFormulas', '수식 표시', { toggle: 'showFormulas' }),
          medium('errorCheck', 'validation', '오류 검사', { menu: 'errorMenu', split: true }),
          medium('evaluateFormula', 'fx', '수식 계산'),
        ),
        large('watchWindow', 'search', '조사식 창'),
      ]),
      group('계산', [large('calcOptionsBtn', 'calc', '계산 옵션', { menu: 'calcOptions' }), col(medium('recalc', 'calc', '지금 계산 (F9)'), medium('calcNowSheet', 'calc', '시트 계산 (Shift+F9)'))]),
      group('옵션', [large('options', 'format', 'WIXEL 옵션')]),
    ],
  },
  {
    id: 'data', label: '데이터', groups: [
      group('데이터 가져오기 및 변환', [
        large('importCsv', 'csvIn', '파일에서 가져오기', { title: 'CSV·TSV·Excel(.xlsx) 파일 가져오기' }),
        large('webData', 'webData', '웹에서', { title: '웹 페이지의 표 · 목록 · CSV 가져오기 (IMPORTHTML)' }),
        large('exportCsv', 'csvOut', 'CSV로 내보내기'),
      ]),
      group('쿼리 및 연결', [large('refreshAll', 'refresh', '모두 새로 고침', { title: '피벗 테이블 새로 고침' })]),
      group('정렬 및 필터', [
        col(
          medium('sortAsc', 'sortAsc', '오름차순'),
          medium('sortDesc', 'sortDesc', '내림차순'),
        ),
        large('sortDialog', 'sort', '정렬'),
        large('toggleFilter', 'filter', '필터', { toggle: 'filterOn', title: '필터 (Ctrl+Shift+L)' }),
        col(
          medium('clearFilter', 'filterClear', '지우기'),
          medium('reapplyFilter', 'refresh', '다시 적용'),
          medium('advancedFilter', 'filter', '고급'),
        ),
      ]),
      group('데이터 도구', [
        large('textToColumns', 'textColumns', '텍스트 나누기', { title: '텍스트 나누기 (Alt+A+E)' }),
        large('dedupe', 'dedupe', '중복된 항목 제거'),
        large('consolidate', 'consolidate', '통합', { title: '여러 범위의 값을 첫 행 · 왼쪽 열 이름으로 모아 합계 · 평균 등을 구함' }),
        large('dataValidation', 'validation', '데이터 유효성 검사', { menu: 'validation', split: true }),
      ]),
      group('예측', [large('whatIfMenu', 'chartLine', '가상 분석', { menu: 'whatIf' }), large('forecastSheet', 'chartLine', '예측 시트', { title: '시간 표시줄의 값으로 미래 값을 예측하는 새 워크시트 (FORECAST.ETS)' })]),
      group('개요', [
        large('outlineGroup', 'rowInsert', '그룹', { title: '그룹 (Shift+Alt+→)', menu: 'outlineGroupMenu', split: true }),
        large('outlineUngroup', 'delete', '그룹 해제', { title: '그룹 해제 (Shift+Alt+←)', menu: 'outlineUngroupMenu', split: true }),
        large('subtotal', 'autosum', '부분합', { title: '그룹마다 요약 행 삽입 (SUBTOTAL)' }),
        col(medium('outlineShow', 'plus', '세부 정보 표시'), medium('outlineHide', 'collapse', '세부 정보 숨기기')),
      ]),
      group('분석', [
        large('dataAnalysis', 'stats', '데이터 분석', { title: '통계 데이터 분석 도구 (상관 · 회귀 · 기술 통계 · t-검정 · 히스토그램 …)' }),
        large('anomalies', 'anomaly', '이상치 찾기', { title: '열마다 급등 · 급락 값을 찾아 조건부 서식으로 강조 (MAD · Z · IQR)' }),
        large('marketingMetrics', 'adMetrics', '광고 지표', { title: '노출 · 클릭 · 비용 · 전환 · 매출 열에서 CTR · CPC · CPM · CVR · CPA · ROAS 열을 수식으로 추가' }),
        large('solver', 'calc', '해 찾기', { title: '제한 조건을 만족하면서 목표 셀을 최대 · 최소 · 지정값으로 만드는 변수 값 찾기' }),
      ]),
    ],
  },
  {
    id: 'review', label: '검토', groups: [
      group('교정', [large('workbookStats', 'stats', '통합 문서 통계')]),
      group('메모', [
        large('editComment', 'newComment', '새 메모'),
        large('deleteComment', 'deleteComment', '삭제'),
        col(
          medium('prevComment', 'prev', '이전 메모'),
          medium('nextComment', 'next', '다음 메모'),
        ),
      ]),
      group('보호', [
        large('protectSheet', 'validation', '시트 보호', { title: '시트 보호 / 시트 보호 해제', toggle: 'sheetProtected' }),
        large('protectWorkbook', 'lock', '통합 문서 보호', { title: '통합 문서 구조 보호 (시트 추가 · 삭제 · 이동 · 이름 변경 막기)', toggle: 'bookProtected' }),
        large('cellProtection', 'format', '셀 잠금', { title: '셀 잠금 · 수식 숨기기 (시트를 보호하면 적용)' }),
      ]),
    ],
  },
  {
    id: 'view', label: '보기', groups: [
      group('표시', [
        col(
          check('toggleGrid', '눈금선', 'showGrid'),
          check('toggleFormulaBar', '수식 입력줄', 'showFormulaBar'),
          check('toggleHeaders', '머리글', 'showHeaders'),
        ),
        col(
          check('focusCellToggle', '포커스 셀', 'focusCellOn'),
          check('valueHighlight', '값 강조', 'valueHighlight'),
        ),
      ]),
      group('확대/축소', [
        large('zoomIn', 'zoomIn', '확대'),
        large('zoomOut', 'zoomOut', '축소'),
        large('zoom100', 'zoom100', '100%'),
        large('zoomSel', 'zoomSel', '선택 영역 확대/축소', { title: '선택한 범위가 창에 꽉 차도록 확대/축소' }),
      ]),
      group('창', [large('freezeMenu', 'freeze', '틀 고정', { menu: 'freeze', toggle: 'frozen' }), large('navigator', 'navigator', '탐색', { title: '시트 · 표 · 피벗 · 이름 · 개체 · 메모 목록에서 찾아 이동' })]),
      group('매크로', [large('macros', 'macro', '매크로', { title: '매크로(VBA) 코드 보기' })]),
    ],
  },
  {
    id: 'tableDesign', label: '테이블 디자인', context: 'table', groups: [
      group('속성', [
        col(
          { type: 'text', cmd: 'tblName', stateKey: 'tblName', label: '표 이름:', title: '표 이름', width: 96 },
          medium('resizeTable', 'tableResize', '표 크기 조정'),
        ),
      ]),
      group('도구', [
        col(
          medium('pivotFromTable', 'pivot', '피벗 테이블로 요약'),
          medium('dedupe', 'dedupe', '중복된 항목 제거'),
          medium('convertToRange', 'tableConvert', '범위로 변환'),
        ),
        large('insertSlicer', 'slicer', '슬라이서 삽입'),
      ]),
      group('표 스타일 옵션', [
        col(check('tblHeader', '머리글 행', 'tblHeader'), check('tblTotals', '요약 행', 'tblTotals'), check('tblBanded', '줄무늬 행', 'tblBanded')),
        col(check('tblFirstCol', '첫째 열', 'tblFirstCol'), check('tblLastCol', '마지막 열', 'tblLastCol'), check('tblBandedCols', '줄무늬 열', 'tblBandedCols')),
        col(check('tblFilter', '필터 단추', 'tblFilter')),
      ]),
      group('표 스타일', [large('tableStyleGallery', 'table', '빠른 스타일', { menu: 'tableStylesDesign' })]),
    ],
  },
  {
    id: 'pivotAnalyze', label: '피벗 테이블 분석', context: 'pivot', groups: [
      group('피벗 테이블', [
        col(
          { type: 'text', cmd: 'pivotName', stateKey: 'pivotName', label: '이름:', title: '피벗 테이블 이름', width: 110 },
          medium('pivotOptions', 'pivot', '옵션'),
        ),
      ]),
      group('활성 필드', [
        col(
          medium('pivotExpandField', 'plus', '전체 필드 확장', { title: '활성 필드의 모든 항목 확장' }),
          medium('pivotCollapseField', 'collapse', '전체 필드 축소', { title: '활성 필드의 모든 항목 축소' }),
        ),
      ]),
      group('그룹', [
        col(
          medium('pivotGroupSelection', 'pivot', '선택 항목 그룹화', { title: '선택한 항목 셀들을 그룹1 · 그룹2 … 로 묶기' }),
          medium('pivotGroupField', 'pivot', '필드 그룹', { title: '날짜를 연 · 분기 · 월로, 숫자를 구간으로 묶기' }),
          medium('pivotUngroup', 'clear', '그룹 해제'),
          medium('pivotDetail', 'table', '세부 정보 표시', { title: '선택한 값 셀의 원본 행을 새 시트에 (값 셀 두 번 클릭)' }),
        ),
      ]),
      group('필터', [large('insertSlicer', 'slicer', '슬라이서 삽입'), large('slicerConnections', 'slicer', '필터 연결', { title: '이 피벗 테이블에 연결할 슬라이서' })]),
      group('데이터', [large('pivotRefresh', 'refresh', '새로 고침', { title: '새로 고침 (Alt+F5)' }), large('pivotChangeSource', 'table', '데이터 원본 변경')]),
      group('동작', [large('pivotClear', 'clear', '지우기', { title: '필드를 모두 지우기' })]),
      group('계산', [large('calcField', 'fx', '필드, 항목 및 집합', { title: '계산 필드 (CPC · CTR · ROAS 등) · 수식 나열', menu: 'calcFields' })]),
      group('표시', [
        large('pivotFieldList', 'pivot', '필드 목록', { title: '피벗 테이블 필드 창 표시/숨기기' }),
        large('pivotShowExpand', 'plus', '+/- 단추', { title: '항목 확장 · 축소 단추 표시', toggle: 'pvShowExpand' }),
      ]),
    ],
  },
  {
    id: 'pivotDesign', label: '디자인', context: 'pivot', groups: [
      group('레이아웃', [
        large('pivotSubtotalsMenu', 'table', '부분합', { menu: 'pivotSubtotals' }),
        large('pivotGrandMenu', 'table', '총합계', { menu: 'pivotGrand' }),
        large('pivotLayoutMenu', 'table', '보고서 레이아웃', { menu: 'pivotLayout' }),
        large('pivotBlankMenu', 'rowInsert', '빈 행', { menu: 'pivotBlank' }),
      ]),
      group('피벗 테이블 스타일 옵션', [
        col(check('pvRowHeaders', '행 머리글', 'pvRowHeaders'), check('pvColHeaders', '열 머리글', 'pvColHeaders')),
        col(check('pvBandRows', '줄무늬 행', 'pvBandRows'), check('pvBandCols', '줄무늬 열', 'pvBandCols')),
      ]),
      group('피벗 테이블 스타일', [large('pivotStyleGalleryBtn', 'table', '빠른 스타일', { menu: 'pivotStylesDesign' })]),
    ],
  },
  {
    id: 'sparkTab', label: '스파크라인', context: 'spark', groups: [
      group('스파크라인', [large('sparkEdit', 'chartLine', '데이터 편집', { title: '스파크라인 그룹의 데이터 · 위치 범위 · 색' })]),
      group('종류', [
        large('sparkTypeLine', 'chartLine', '꺾은선형', { toggle: 'sparkIsLine' }),
        large('sparkTypeColumn', 'chartColumn', '열', { toggle: 'sparkIsColumn' }),
        large('sparkTypeWinLoss', 'chartBar', '승패', { toggle: 'sparkIsWinLoss' }),
      ]),
      group('표시', [
        col(check('sparkHigh', '높은 점', 'sparkHigh'), check('sparkLow', '낮은 점', 'sparkLow'), check('sparkNegative', '음수 점', 'sparkNegative')),
        col(check('sparkFirst', '첫 점', 'sparkFirst'), check('sparkLast', '마지막 점', 'sparkLast'), check('sparkMarkers', '표식', 'sparkMarkers')),
      ]),
      group('그룹', [large('sparkClear', 'clear', '지우기', { title: '선택한 스파크라인 그룹 지우기' })]),
    ],
  },
  {
    id: 'slicerTab', label: '슬라이서', context: 'slicer', groups: [
      group('슬라이서', [
        col(
          { type: 'text', cmd: 'slicerCaption', stateKey: 'slicerCaption', label: '캡션:', title: '슬라이서 캡션', width: 110 },
          medium('slicerSettings', 'slicer', '슬라이서 설정'),
        ),
      ]),
      group('필터', [
        large('slicerClear', 'filterClear', '필터 지우기', { title: '필터 지우기 (Alt+C)' }),
        large('slicerMulti', 'filter', '다중 선택', { title: '다중 선택 (Alt+S)', toggle: 'slicerMultiOn' }),
      ]),
      group('슬라이서 스타일', [large('slicerStyleGalleryBtn', 'slicer', '빠른 스타일', { menu: 'slicerStyles' })]),
      group('단추', [
        col(
          { type: 'spin', cmd: 'slicerCols', stateKey: 'slicerCols', label: '열:', title: '열 수', min: 1, max: 20, step: 1 },
          { type: 'spin', cmd: 'slicerBtnH', stateKey: 'slicerBtnH', label: '높이:', title: '단추 높이 (px)', min: 10, max: 120, step: 1 },
          { type: 'spin', cmd: 'slicerBtnW', stateKey: 'slicerBtnW', label: '너비:', title: '단추 너비 (px, 0 = 자동)', min: 0, max: 600, step: 1 },
        ),
        col(
          { type: 'spin', cmd: 'slicerGap', stateKey: 'slicerGap', label: '간격:', title: '단추 사이 간격 (px)', min: 0, max: 30, step: 1 },
        ),
      ]),
      group('표시', [col(check('slicerHeader', '머리글 표시', 'slicerHeaderOn')), large('slicerConnections', 'slicer', '보고서 연결')]),
      group('글꼴', [
        col(
          { type: 'spin', cmd: 'slicerFontSize', stateKey: 'slicerFontSize', label: '항목:', title: '항목 글꼴 크기 (pt, 비우면 기본)', min: 5, max: 72, step: 0.5 },
          { type: 'spin', cmd: 'slicerHeadSize', stateKey: 'slicerHeadSize', label: '머리글:', title: '머리글 글꼴 크기 (pt, 비우면 기본)', min: 5, max: 72, step: 0.5 },
          check('slicerBold', '굵게', 'slicerBoldOn'),
        ),
      ]),
      group('정렬', [
        col(medium('objForwardBtn', 'bringForward', '앞으로 가져오기', { menu: 'objForward' }), medium('objBackwardBtn', 'sendBackward', '뒤로 보내기', { menu: 'objBackward' }), medium('selectionPane', 'selectionPane', '선택 창')),
        col(medium('objAlignBtn', 'align', '맞춤', { menu: 'objAlign' }), medium('objPlacementBtn', 'placement', '위치 속성', { menu: 'objPlacement' })),
      ]),
      group('크기', [
        col(
          { type: 'spin', cmd: 'objH', stateKey: 'objH', label: '높이:', title: '슬라이서 높이 (px)', min: 20, max: 4000, step: 1 },
          { type: 'spin', cmd: 'objW', stateKey: 'objW', label: '너비:', title: '슬라이서 너비 (px)', min: 30, max: 4000, step: 1 },
        ),
      ]),
    ],
  },
  {
    id: 'objFormat', label: '셰이프 형식', context: 'object', groups: [
      group('도형 삽입', [
        large('shapesMenu', 'shapes', '도형', { menu: 'shapes' }),
        col(medium('shapeChangeBtn', 'shapes', '도형 모양 변경', { menu: 'shapeChange' }), medium('insertTextbox', 'textbox', '텍스트 상자')),
      ]),
      group('도형 스타일', [
        large('shapeStylesBtn', 'effects', '빠른 스타일', { menu: 'shapeStyles' }),
        col(medium('shapeFillBtn', 'fill', '도형 채우기', { menu: 'shapeFill' }), medium('shapeOutlineBtn', 'border', '도형 윤곽선', { menu: 'shapeOutline' }), medium('shapeEffectsBtn', 'effects', '도형 효과', { menu: 'shapeEffects' })),
      ]),
      group('WordArt 스타일', [
        large('wordArtBtn', 'wordart', '빠른 스타일', { menu: 'wordArt' }),
        col(medium('textFillBtn', 'fontColor', '텍스트 채우기', { menu: 'textFill' }), medium('textOutlineBtn', 'wordart', '텍스트 윤곽선', { menu: 'textOutline' }), medium('textEffectsBtn', 'effects', '텍스트 효과', { menu: 'textEffects' })),
      ]),
      group('정렬', [
        col(medium('objForwardBtn', 'bringForward', '앞으로 가져오기', { menu: 'objForward' }), medium('objBackwardBtn', 'sendBackward', '뒤로 보내기', { menu: 'objBackward' }), medium('selectionPane', 'selectionPane', '선택 창')),
        col(medium('objAlignBtn', 'align', '맞춤', { menu: 'objAlign' }), medium('objRotateBtn', 'rotate', '회전', { menu: 'objRotate' }), medium('objPlacementBtn', 'placement', '위치 속성', { menu: 'objPlacement' })),
      ]),
      group('크기', [
        col(
          { type: 'text', cmd: 'objH', stateKey: 'objH', label: '높이:', title: '도형 높이 (px)', width: 58 },
          { type: 'text', cmd: 'objW', stateKey: 'objW', label: '너비:', title: '도형 너비 (px)', width: 58 },
          { type: 'text', cmd: 'objRot', stateKey: 'objRot', label: '회전:', title: '회전 각도 (°)', width: 58 },
        ),
      ]),
    ],
  },
  {
    id: 'chartDesign', label: '차트 디자인', context: 'chart', groups: [
      group('차트 레이아웃', [large('chartElementsBtn', 'chartColumn', '차트 요소 추가', { menu: 'chartElements' }), large('chartLayoutBtn', 'table', '빠른 레이아웃', { menu: 'chartLayouts' })]),
      group('차트 스타일', [large('chartColorsBtn', 'fill', '색 변경', { menu: 'chartColors' }), { type: 'gallery', gallery: 'chartStyles', stateKey: 'chartGalleryKey', menu: 'chartStyles', title: '차트 스타일' }]),
      group('데이터', [large('chartSwitch', 'refresh', '행/열 전환'), large('chartSelectData', 'table', '데이터 선택')]),
      group('종류', [large('chartChangeType', 'chartColumn', '차트 종류 변경')]),
      group('위치', [large('chartMove', 'placement', '차트 이동')]),
      group('서식', [large('chartFormat', 'format', '차트 서식 창'), large('chartPivotFields', 'pivot', '필드 단추', { toggle: 'chartFieldButtons' })]),
    ],
  },
  {
    id: 'help', label: '도움말', groups: [
      group('도움말', [
        large('shortcuts', 'keyboard', '바로 가기 키'),
        large('whatsNew', 'effects', '새로운 기능'),
        large('about', 'about', '정보'),
      ]),
    ],
  },
];

/**
 * app: { run(cmd, arg), openMenu(name, anchorEl, arg), focusGrid() }
 * 반환: { update(state) }
 */
/** 리본의 모든 명령 (빠른 실행 도구 모음 사용자 지정용): [{ cmd, icon, label, tab }] */
export function ribbonCommands() {
  const out = [];
  const seen = new Set();
  const walk = (it, tab) => {
    if (!it) return;
    if (it.items) { for (const x of it.items) walk(x, tab); return; }
    if (!it.cmd || seen.has(it.cmd) || !['large', 'medium', 'btn', 'check'].includes(it.type)) return;
    seen.add(it.cmd);
    const label = String(it.label ?? it.title ?? it.cmd).replace(/\s*\(.*\)$/, '').replace(/\n/g, ' ');
    out.push({ cmd: it.cmd, icon: it.icon, label, tab: tab.label });
  };
  for (const t of TABS) for (const g of t.groups ?? []) for (const it of g.items) walk(it, t);
  return out;
}

export function buildRibbon(app) {
  const tabsEl = document.getElementById('ribbonTabs');
  const ribbonEl = document.getElementById('ribbon');
  const bindings = [];
  let current = 'home';

  const keepFocus = (e) => e.preventDefault();

  let context = new Set(); // 상황별 탭 (예: 'table' → 테이블 디자인)

  function selectTab(id) {
    const t = TABS.find((x) => x.id === id);
    if (!t || t.file || (t.context && !context.has(t.context))) return;
    ribbonEl.classList.remove('collapsed');
    if (current === id) return;
    current = id;
    renderTabs();
    renderRibbon();
    app.refreshRibbon();
  }

  function renderTabs() {
    tabsEl.replaceChildren();
    const hidden = app.hiddenTabs?.() ?? [];
    for (const t of TABS) {
      if (t.context && !context.has(t.context)) continue;
      if (!t.file && hidden.includes(t.id)) continue;
      tabsEl.append(el('button', {
        class: `ribbon-tab${t.file ? ' file' : ''}${t.context ? ' contextual' : ''}${t.id === current ? ' active' : ''}`,
        onmousedown: keepFocus,
        onclick: () => {
          if (t.file) { app.run('backstage'); return; }
          if (t.id === current && !ribbonEl.classList.contains('collapsed')) return;
          ribbonEl.classList.remove('collapsed');
          current = t.id;
          renderTabs();
          renderRibbon();
          app.refreshRibbon();
        },
        ondblclick: () => { if (!t.file) ribbonEl.classList.toggle('collapsed'); },
      }, t.label));
    }
    tabsEl.append(
      el('span', { class: 'spacer' }),
      el('button', { class: 'ribbon-action', title: '버전 기록 (이 브라우저에 보관된 이전 버전 보기 · 복원)', onmousedown: keepFocus, onclick: () => app.run('versionHistory') },
        el('span', { html: ICONS.history }), el('span', { class: 'lbl' }, '버전')),
      el('button', { class: 'ribbon-action', title: '공유 · 웹에 게시 (읽기 전용 링크 · 대시보드 보기)', onmousedown: keepFocus, onclick: () => app.run('publish') },
        el('span', { html: ICONS.share }), el('span', { class: 'lbl' }, '공유')),
      el('button', { class: 'ribbon-action', title: '메모 (Shift+F2)', onmousedown: keepFocus, onclick: () => app.run('editComment') },
        el('span', { html: ICONS.comment }), el('span', { class: 'lbl' }, '메모')),
      el('button', { class: 'ribbon-action primary', title: '저장 (Ctrl+S)', onmousedown: keepFocus, onclick: () => app.run('save') },
        el('span', { html: ICONS.save }), el('span', { class: 'lbl' }, '저장')),
    );
  }

  function makeItem(it) {
    switch (it.type) {
      case 'row': return el('div', { class: 'rrow' }, it.items.map(makeItem));
      case 'col': return el('div', { class: 'rcol' }, it.items.map(makeItem));
      case 'sep': return el('span', { class: 'rsep' });
      case 'select': return makeSelect(it);
      case 'font': return makeFont(it);
      case 'text': return makeText(it);
      case 'spin': return makeSpin(it);
      case 'check': return makeCheck(it);
      case 'color': return makeColor(it);
      case 'gallery': return makeGallery(it);
      case 'combo': return makeCombo(it);
      default: return makeButton(it);
    }
  }

  function makeButton(it) {
    const cls = ['rbtn', it.type === 'large' ? 'large' : it.type === 'medium' ? 'medium' : ''];
    const iconHtml = ICONS[it.icon] ?? it.icon ?? '';
    const onlyMenu = it.menu && !it.split && !['borderLast', 'mergeCenter'].includes(it.cmd) && it.cmd !== 'paste';
    const b = el('button', {
      class: cls.join(' '), title: it.title ?? it.label, onmousedown: keepFocus,
      ondblclick: it.cmd === 'painter' ? () => app.run('painterSticky') : undefined,
    });
    b.innerHTML = iconHtml;
    if (it.label) b.append(el('span', { class: 'lbl' }, it.label));
    if (it.menu) {
      const caret = el('span', { class: 'caret', html: ICONS.chevronDown });
      b.append(caret);
      if (!onlyMenu) b.classList.add('split');
      b.addEventListener('click', (e) => {
        const onCaret = caret.contains(e.target) || (it.type === 'large' && e.clientY - b.getBoundingClientRect().top > 44);
        if (onlyMenu || onCaret) app.openMenu(it.menu, b);
        else app.run(it.cmd);
      });
    } else {
      b.addEventListener('click', () => app.run(it.cmd));
    }
    if (it.toggle) bindings.push((s) => b.classList.toggle('on', !!s[it.toggle]));
    return b;
  }

  function makeSelect(it) {
    if (it.editable) {
      const input = el('input', { class: `rselect ${it.cls}`, title: it.title, list: `dl-${it.cmd}` });
      const dl = el('datalist', { id: `dl-${it.cmd}` }, it.options.map((o) => el('option', { value: o.value })));
      input.addEventListener('change', () => { app.run(it.cmd, input.value); app.focusGrid(); });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { input.blur(); app.focusGrid(); }
        if (e.key === 'Escape') app.focusGrid();
      });
      bindings.push((s) => { if (document.activeElement !== input) input.value = s[it.stateKey] ?? ''; });
      return el('span', {}, input, dl);
    }
    const sel = el('select', { class: `rselect ${it.cls}`, title: it.title },
      it.options.map((o) => el('option', { value: o.value }, o.label)));
    sel.addEventListener('change', () => { app.run(it.cmd, sel.value); app.focusGrid(); });
    bindings.push((s) => { sel.value = s[it.stateKey] ?? it.options[0].value; });
    return sel;
  }

  /** 글꼴 상자: 이름을 직접 입력하거나 ▾ 로 글꼴 목록 (각 글꼴 모양으로 표시) */
  function makeFont(it) {
    const input = el('input', { class: `rselect ${it.cls}`, title: it.title, spellcheck: false });
    const commit = () => { const v = input.value.trim(); if (v) app.run(it.cmd, v); };
    input.addEventListener('change', commit);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); app.focusGrid(); }
      if (e.key === 'Escape') app.focusGrid();
      if (e.key === 'ArrowDown' && e.altKey) { e.preventDefault(); app.openMenu(it.menu, wrap); }
    });
    const caret = el('button', { class: 'rbtn font-caret', title: '글꼴 목록', onmousedown: keepFocus, html: ICONS.chevronDown });
    caret.addEventListener('click', () => app.openMenu(it.menu, wrap));
    const wrap = el('span', { class: 'font-box' }, input, caret);
    bindings.push((s) => {
      if (document.activeElement === input) return;
      input.value = s[it.stateKey] ?? '';
      input.style.fontFamily = `'${String(input.value).replace(/'/g, '')}', 'Malgun Gothic', sans-serif`;
    });
    return wrap;
  }

  function makeText(it) {
    const input = el('input', { class: 'rtext', title: it.title, style: { width: `${it.width ?? 110}px` } });
    const commit = () => { app.run(it.cmd, input.value); };
    input.addEventListener('change', commit);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); app.focusGrid(); }
      if (e.key === 'Escape') { input.blur(); app.focusGrid(); }
    });
    bindings.push((s) => { if (document.activeElement !== input) input.value = s[it.stateKey] ?? ''; });
    return el('label', { class: 'rbtn medium rtext-wrap' }, it.label ? el('span', {}, it.label) : null, input);
  }

  /** 숫자 칸 (▲▼ · 방향키 · 휠로 한 단계씩, 입력하면 바로 적용) */
  function makeSpin(it) {
    const input = el('input', { type: 'number', class: 'rtext rspin', title: it.title, min: it.min, max: it.max, step: it.step ?? 1, style: { width: `${it.width ?? 58}px` } });
    let t = null;
    const commit = () => { clearTimeout(t); t = null; app.run(it.cmd, input.value); };
    input.addEventListener('input', () => { clearTimeout(t); t = setTimeout(commit, 250); });
    input.addEventListener('change', commit);
    input.addEventListener('wheel', (e) => {
      if (document.activeElement !== input) return;
      e.preventDefault();
      const v = (Number(input.value) || 0) + (e.deltaY < 0 ? 1 : -1) * Number(it.step ?? 1);
      input.value = String(Math.min(it.max ?? Infinity, Math.max(it.min ?? -Infinity, Math.round(v * 100) / 100)));
      commit();
    }, { passive: false });
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commit(); input.blur(); app.focusGrid(); }
      if (e.key === 'Escape') { input.blur(); app.focusGrid(); }
    });
    bindings.push((s) => { if (document.activeElement !== input) input.value = s[it.stateKey] ?? ''; });
    return el('label', { class: 'rbtn medium rtext-wrap' }, it.label ? el('span', {}, it.label) : null, input);
  }

  /** 목록 단추 (엑셀 표시 형식 상자): 현재 값 이름 + ▾, 누르면 그림 · 보기가 있는 메뉴 */
  function makeCombo(it) {
    const txt = el('span', { class: 'rcombo-txt' });
    const b = el('button', { class: `rbtn rcombo ${it.cls ?? ''}`, title: it.title, onmousedown: keepFocus }, txt, el('span', { class: 'caret', html: ICONS.chevronDown }));
    b.addEventListener('click', () => app.openMenu(it.menu, b));
    bindings.push((s) => { const v = s[it.stateKey]; txt.textContent = it.options.find((o) => o.id === v)?.label ?? (v ? '사용자 지정' : '일반'); });
    return b;
  }

  /** 리본 안 갤러리 (엑셀 차트 스타일처럼 견본이 바로 보이고 ▾ 로 전체 목록) */
  function makeGallery(it) {
    const strip = el('div', { class: 'rg-strip' });
    const more = el('button', { class: 'rbtn rg-more', title: `${it.title} 더 보기`, onmousedown: keepFocus, html: ICONS.chevronDown });
    more.addEventListener('click', () => app.openMenu(it.menu, more));
    let last = null;
    bindings.push((s) => {
      const k = s[it.stateKey] ?? '';
      if (k === last) return;
      last = k;
      strip.replaceChildren(...(k ? app.gallery?.(it.gallery) ?? [] : []));
    });
    return el('div', { class: 'rgallery', title: it.title }, strip, more);
  }

  function makeCheck(it) {
    const input = el('input', { type: 'checkbox' });
    input.addEventListener('change', () => { app.run(it.cmd, input.checked); app.focusGrid(); });
    bindings.push((s) => { input.checked = !!s[it.stateKey]; });
    return el('label', { class: 'rbtn medium', style: { gap: '6px' } }, input, it.label);
  }

  function makeColor(it) {
    const bar = el('span', { class: 'color-bar' });
    const main = el('button', { class: 'rbtn color-btn', title: it.title, onmousedown: keepFocus, html: ICONS[it.icon] });
    main.append(bar);
    main.addEventListener('click', () => app.run(it.cmd));
    const caret = el('button', {
      class: 'rbtn', style: { minWidth: '12px', padding: '0 1px' }, title: `${it.title} 선택`, onmousedown: keepFocus,
      html: ICONS.chevronDown,
    });
    caret.addEventListener('click', () => app.openMenu(it.menu, main));
    bindings.push((s) => { bar.style.background = s[it.stateKey]; });
    return el('span', { style: { display: 'inline-flex' } }, main, caret);
  }

  function renderRibbon() {
    bindings.length = 0;
    ribbonEl.replaceChildren();
    const tab = TABS.find((t) => t.id === current);
    for (const g of tab.groups) {
      ribbonEl.append(el('div', { class: 'rgroup' },
        el('div', { class: 'rgroup-body' }, g.items.map(makeItem)),
        el('div', { class: 'rgroup-label' }, g.label),
        g.launcher ? el('button', { class: 'rgroup-launcher', title: '자세히', onmousedown: keepFocus, onclick: () => app.run(g.launcher) }, '⇲') : null));
    }
    ribbonEl.append(el('span', { style: { flex: '1' } }),
      el('button', {
        class: 'rbtn', style: { alignSelf: 'flex-end', marginBottom: '2px' }, title: '리본 축소 (Ctrl+F1)',
        onmousedown: keepFocus, html: ICONS.collapse, onclick: () => ribbonEl.classList.add('collapsed'),
      }));
  }

  renderTabs();
  renderRibbon();

  return {
    update(state) {
      const next = new Set(state.context ?? []);
      if ([...next].join() !== [...context].join()) {
        const leaving = TABS.find((t) => t.id === current)?.context;
        context = next;
        if (leaving && !context.has(leaving)) { current = 'home'; renderRibbon(); }
        renderTabs();
      }
      for (const fn of bindings) fn(state);
    },
    selectTab,
    /** 리본 사용자 지정 뒤 다시 그리기 (숨긴 탭이 선택돼 있으면 홈으로) */
    reset() {
      if ((app.hiddenTabs?.() ?? []).includes(current)) current = 'home';
      renderTabs();
      renderRibbon();
      app.refreshRibbon();
    },
    get current() { return current; },
    toggleCollapse() { ribbonEl.classList.toggle('collapsed'); },
  };
}
