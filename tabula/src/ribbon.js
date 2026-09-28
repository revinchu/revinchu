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
            { type: 'select', cmd: 'fontFamily', cls: 'font-family', options: FONTS.map((f) => ({ value: f, label: f })), stateKey: 'font', title: '글꼴' },
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
          row({ type: 'select', cmd: 'numFmt', cls: 'numfmt', options: NUMBER_FORMATS.map((f) => ({ value: f.id, label: f.label })), stateKey: 'numFmt', title: '표시 형식' }),
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
        large('createTable', 'table', '표', { title: '표 만들기 (Ctrl+T, Ctrl+L)' }),
      ]),
      group('일러스트레이션', [
        large('insertPicture', 'picture', '그림', { title: '이 기기의 그림 삽입 (붙여넣기·끌어 놓기도 가능)' }),
        large('shapesMenu', 'shapes', '도형', { menu: 'shapes' }),
        large('insertTextbox', 'textbox', '텍스트 상자'),
      ]),
      group('필터', [large('insertSlicer', 'slicer', '슬라이서', { title: '표나 피벗 테이블에 슬라이서 삽입' })]),
      group('차트', [
        large('chartColumn', 'chartColumn', '세로 막대형', { title: '세로 막대형 차트 삽입 (Alt+F1)' }),
        large('chartBar', 'chartBar', '가로 막대형'),
        large('chartLine', 'chartLine', '꺾은선형'),
        large('chartPie', 'chartPie', '원형', { menu: 'pieCharts', split: true }),
        large('chartArea', 'chartArea', '영역형'),
        large('chartScatter', 'chartScatter', '분산형'),
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
      group('수식 분석', [large('toggleFormulas', 'showFormulas', '수식 표시', { toggle: 'showFormulas' })]),
      group('계산', [large('recalc', 'calc', '지금 계산')]),
    ],
  },
  {
    id: 'data', label: '데이터', groups: [
      group('데이터 가져오기 및 변환', [
        large('importCsv', 'csvIn', '파일에서 가져오기', { title: 'CSV·TSV·Excel(.xlsx) 파일 가져오기' }),
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
        ),
      ]),
      group('데이터 도구', [
        large('textToColumns', 'textColumns', '텍스트 나누기', { title: '텍스트 나누기 (Alt+A+E)' }),
        large('dedupe', 'dedupe', '중복된 항목 제거'),
        large('dataValidation', 'validation', '데이터 유효성 검사', { menu: 'validation', split: true }),
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
      ]),
      group('확대/축소', [
        large('zoomIn', 'zoomIn', '확대'),
        large('zoomOut', 'zoomOut', '축소'),
        large('zoom100', 'zoom100', '100%'),
      ]),
      group('창', [large('freezeMenu', 'freeze', '틀 고정', { menu: 'freeze', toggle: 'frozen' })]),
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
      group('단추', [
        { type: 'select', cmd: 'slicerCols', stateKey: 'slicerCols', cls: 'w60', title: '열 수', options: [1, 2, 3, 4, 5, 6].map((n) => ({ value: String(n), label: `열 ${n}` })) },
      ]),
    ],
  },
  {
    id: 'help', label: '도움말', groups: [
      group('도움말', [
        large('shortcuts', 'keyboard', '바로 가기 키'),
        large('about', 'about', '정보'),
      ]),
    ],
  },
];

/**
 * app: { run(cmd, arg), openMenu(name, anchorEl, arg), focusGrid() }
 * 반환: { update(state) }
 */
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
    for (const t of TABS) {
      if (t.context && !context.has(t.context)) continue;
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
      case 'text': return makeText(it);
      case 'check': return makeCheck(it);
      case 'color': return makeColor(it);
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
    get current() { return current; },
    toggleCollapse() { ribbonEl.classList.toggle('collapsed'); },
  };
}
