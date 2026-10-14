// 리본 메뉴 · 갤러리 · 리본 상태 (Word 한국어판)
import { S, run, register, emit, change, curSel } from './state.js';
import { el, openMenu, openSubmenu, closeMenus, toast, hydrateIcons } from './ui.js';
import { ICONS } from './icons.js';
import {
  blockAt, effPara, effRun, propsAt, paraLen, posEq, orderPos, HIGHLIGHTS, HIGHLIGHT_LABEL, MARGINS, PAPERS, pxToCm, cm, para, uid, styleChain, sections,
  bulletList, numberList, multiList, formatNumber, newDoc, clone, walkParas,
} from './model.js';
import { selProp, selParas, STYLE_SET_LIST, sectionAtTop } from './commands.js';
import { tableStyleList, tableHtml, runCss, ACCENTS } from './render.js';
import { cellSelection } from './view.js';
import { TEMPLATES } from './sample.js';

const sel = () => curSel();
const f1 = (v) => Math.round(v * 10) / 10;

// ───────────── 색 고르기 (Word: 테마 색 · 표준 색 · 다른 색) ─────────────
export const THEME_COLS = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const STD_COLS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
const mix = (hex, f) => { const n = parseInt(hex.slice(1), 16); const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(f >= 0 ? v + (255 - v) * f : v * (1 + f))); return `#${c.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`; };
const SHADES = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
function colorGrid(onPick, { auto = '자동', none = null, recentKey = 'fontColor' } = {}) {
  const wrap = el('div', { class: 'color-menu' });
  if (auto) wrap.append(el('button', { class: 'menu-item', onclick: () => { closeMenus(); onPick('auto'); } }, el('span', { class: 'mi-swatch', style: { background: '#000', width: '14px', height: '14px', display: 'inline-block', marginRight: '8px' } }), auto));
  if (none) wrap.append(el('button', { class: 'menu-item', onclick: () => { closeMenus(); onPick('none'); } }, el('span', { class: 'mi-swatch', style: { border: '1px solid #999', width: '14px', height: '14px', display: 'inline-block', marginRight: '8px' } }), none));
  wrap.append(el('div', { class: 'menu-title' }, '테마 색'));
  const grid = el('div', { class: 'cgrid' });
  for (const f of SHADES) for (const c of THEME_COLS) {
    const v = f === 0 ? c : mix(c, c === '#FFFFFF' ? -Math.abs(f) * 0.6 : c === '#000000' ? Math.abs(f) * 0.9 : f);
    grid.append(el('button', { class: 'csw', title: v, style: { background: v }, onclick: () => { closeMenus(); onPick(v); } }));
  }
  wrap.append(grid, el('div', { class: 'menu-title' }, '표준 색'));
  const std = el('div', { class: 'cgrid' });
  for (const c of STD_COLS) std.append(el('button', { class: 'csw', title: c, style: { background: c }, onclick: () => { closeMenus(); onPick(c); } }));
  wrap.append(std);
  const recent = recentColors(recentKey);
  if (recent.length) { const rc = el('div', { class: 'cgrid' }); for (const c of recent) rc.append(el('button', { class: 'csw', title: c, style: { background: c }, onclick: () => { closeMenus(); onPick(c); } })); wrap.append(el('div', { class: 'menu-title' }, '최근에 사용한 색'), rc); }
  const inp = el('input', { type: 'color', value: '#2B579A', style: { width: '0', height: '0', opacity: '0', position: 'absolute' } });
  inp.addEventListener('change', () => { addRecent(recentKey, inp.value.toUpperCase()); onPick(inp.value.toUpperCase()); closeMenus(); });
  wrap.append(el('button', { class: 'menu-item', onclick: () => inp.click() }, el('span', { class: 'mi-icon', html: ICONS.fill }), '다른 색...'), inp);
  return wrap;
}
function recentColors(k) { try { return JSON.parse(localStorage.getItem(`wiword:recent:${k}`) ?? '[]'); } catch { return []; } }
function addRecent(k, c) { try { localStorage.setItem(`wiword:recent:${k}`, JSON.stringify([c, ...recentColors(k).filter((x) => x !== c)].slice(0, 10))); } catch { /* 없음 */ } }

// ───────────── 리본 상태 (단추 켜짐 · 색 막대 · 상황별 탭) ─────────────
register({
  ribbonState: () => {
    const st = { bars: { font: S.lastFontColor ?? '#C00000', hl: HIGHLIGHTS[S.lastHighlight ?? 'yellow'], shd: S.lastShade ?? '#D9D9D9', cellShd: S.lastCellShade ?? '#D9E2F3' }, context: [] };
    try {
      const p = blockAt(S.doc, sel().f.p);
      if (p?.type === 'p') {
        const ep = effPara(S.doc, p);
        st.b = !!selProp('b'); st.i = !!selProp('i'); st.u = !!selProp('u'); st.s = !!selProp('s'); st.sub = selProp('va') === 'sub'; st.sup = selProp('va') === 'super';
        const al = ep.align ?? 'left';
        st.al = al === 'left'; st.ac = al === 'center'; st.ar = al === 'right'; st.aj = al === 'justify'; st.ad = al === 'distribute';
        const num = ep.num?.id ? S.doc.lists[ep.num.id] : null;
        st.bullets = !!num && num.levels?.[ep.num.lvl ?? 0]?.fmt === 'bullet';
        st.numbering = !!num && !st.bullets;
        const rp = S.typing ? { ...effRun(S.doc, p, propsAt(p, sel().f.o)), ...S.typing } : null;
        st.font = rp?.fontEa ?? selProp('fontEa') ?? selProp('font');
        st.size = rp?.sz ?? selProp('sz');
        st.style = p.style ?? 'Normal';
      }
      if (sel().f.p.length >= 4 || cellSelection()) {
        st.context.push('table');
        const t = blockAt(S.doc, sel().f.p.slice(0, sel().f.p.length - 3));
        if (t?.type === 'tbl') { const lk = { firstRow: true, firstCol: true, bandRows: true, ...(t.tblPr?.look ?? {}) }; st.lkFirstRow = lk.firstRow; st.lkLastRow = lk.lastRow; st.lkBandRows = lk.bandRows; st.lkFirstCol = lk.firstCol; st.lkLastCol = lk.lastCol; st.lkBandCols = lk.bandCols; st.repeatHeader = !!t.rows[sel().f.p[sel().f.p.length - 3]]?.header; }
      }
    } catch { /* 선택 없음 */ }
    if (S.objSel) st.context.push('picture');
    if (S.hf) { st.context.push('hf'); const sec = sectionAtTop(S.doc, 0); st.titlePg = !!sec.titlePg; st.evenOdd = !!S.doc.settings.evenOdd; st.showDocText = S.opts.hideDocText !== true; }
    st.painter = !!S.painter; st.marks = S.marks; st.track = !!S.doc.settings.track; st.ruler = S.showRuler; st.nav = S.showNav; st.gridlines = !!S.opts.gridlines;
    st.vRead = S.view === 'read'; st.vPrint = S.view === 'print'; st.vWeb = S.view === 'web'; st.vOutline = S.view === 'outline'; st.vDraft = S.view === 'draft';
    st.dictating = !!S.dictating; st.reading = !!S.reading; st.mergePreview = !!S.merge?.preview;
    return st;
  },
});

// ───────────── 붙여넣기 · 글꼴 메뉴 ─────────────
register({
  pasteMenu: (a) => openMenu(a, [
    { label: '원본 서식 유지 (K)', icon: 'paste', action: () => run('paste') },
    { label: '텍스트만 유지 (T)', icon: 'paste', action: () => run('paste', 'plain') },
    { label: '마크다운으로 붙여넣기', icon: 'markdown', action: () => run('paste', 'markdown') },
    { sep: true },
    { label: '선택하여 붙여넣기...', action: () => run('paste', 'special') },
  ]),
  changeCaseMenu: (a) => openMenu(a, [
    { label: '문장의 첫 글자를 대문자로.', action: () => run('changeCase', 'sentence') },
    { label: '모두 소문자로', action: () => run('changeCase', 'lower') },
    { label: '모두 대문자로', action: () => run('changeCase', 'upper') },
    { label: '각 단어를 대문자로', action: () => run('changeCase', 'title') },
    { label: '대/소문자 전환', action: () => run('changeCase', 'toggle') },
    { sep: true },
    { label: '반자', action: () => run('changeCase', 'half') },
    { label: '전자', action: () => run('changeCase', 'full') },
  ]),
  underlineMenu: (a) => openMenu(a, [
    ...[['single', '밑줄'], ['double', '이중 밑줄'], ['thick', '굵은 밑줄'], ['dotted', '점선 밑줄'], ['dash', '파선 밑줄'], ['dotDash', '일점 쇄선'], ['wave', '물결선']].map(([k, l]) => ({ label: l, node: el('button', { class: 'menu-item', onclick: () => { closeMenus(); run('underline', k); } }, el('span', { class: 'mi-icon' }), el('span', { style: { textDecoration: 'underline', textDecorationStyle: k === 'double' ? 'double' : k === 'dotted' ? 'dotted' : k === 'dash' || k === 'dotDash' ? 'dashed' : k === 'wave' ? 'wavy' : 'solid', textDecorationThickness: k === 'thick' ? '2px' : 'auto', minWidth: '120px', display: 'inline-block' } }, l)) })),
    { sep: true },
    { label: '밑줄 없음', action: () => run('underline', 'none') },
    { label: '밑줄 색', submenu: [{ node: colorGrid((c) => run('applyRunProps', { uColor: c === 'auto' ? undefined : c })) }] },
  ]),
  textEffectMenu: (a) => openMenu(a, [
    { label: '윤곽선', action: () => run('textEffect', 'outline') },
    { label: '그림자', action: () => run('textEffect', 'shadow') },
    { label: '양각', action: () => run('textEffect', 'emboss') },
    { label: '음각', action: () => run('textEffect', 'imprint') },
    { label: '효과 없음', action: () => run('textEffect', 'none') },
    { sep: true },
    { label: '강조점 (위에 점)', action: () => run('emphasisMark', 'dot') },
    { label: '강조점 없음', action: () => run('emphasisMark', 'none') },
    { label: '작은 대문자', action: () => run('smallCaps') },
    { label: '모두 대문자', action: () => run('allCaps') },
    { label: '숨김', action: () => run('hidden') },
  ]),
  highlightMenu: (a) => {
    const grid = el('div', { class: 'cgrid hlgrid' });
    for (const [k, c] of Object.entries(HIGHLIGHTS)) if (k !== 'white') grid.append(el('button', { class: 'csw', title: HIGHLIGHT_LABEL[k] ?? k, style: { background: c }, onclick: () => { closeMenus(); S.lastHighlight = k; run('highlight', k); } }));
    openMenu(a, [{ node: grid }, { label: '색 없음', action: () => run('highlight', 'none') }]);
  },
  fontColorMenu: (a) => openMenu(a, [{ node: colorGrid((c) => { if (c !== 'auto') S.lastFontColor = c; run('fontColor', c); }) }]),
  shadingMenu: (a) => openMenu(a, [{ node: colorGrid((c) => { if (c !== 'none') S.lastShade = c; run('paraShading', c); }, { auto: null, none: '색 없음', recentKey: 'shade' }) }]),
  applyRunProps: (props) => run('__applyRun', props),
  bordersMenu: (a) => openMenu(a, [
    { label: '아래쪽 테두리', icon: 'borderBottom', action: () => run('paraBorder', 'bottom') },
    { label: '위쪽 테두리', icon: 'borderTop', action: () => run('paraBorder', 'top') },
    { label: '왼쪽 테두리', icon: 'borderLeft', action: () => run('paraBorder', 'left') },
    { label: '오른쪽 테두리', icon: 'borderRight', action: () => run('paraBorder', 'right') },
    { label: '테두리 없음', icon: 'borderNone', action: () => run('paraBorder', 'none') },
    { label: '모든 테두리', icon: 'borderAll', action: () => run('paraBorder', 'all') },
    { label: '바깥쪽 테두리', icon: 'borderOutside', action: () => run('paraBorder', 'box') },
    { sep: true },
    { label: '가로줄', action: () => run('insertBlocksAt', [para('', { pPr: { border: { bottom: { style: 'single', w: 1, color: '#A5A5A5', space: 1 } } } }), para('')]) },
    ...(sel().f.p.length >= 4 ? [{ sep: true }, { title: '표 테두리' }, ...cellBorderItems()] : []),
    { sep: true },
    { label: '테두리 및 음영...', action: () => run('bordersDialog') },
  ]),
});
function cellBorderItems() {
  return [['all', '모든 테두리', 'borderAll'], ['outside', '바깥쪽 테두리', 'borderOutside'], ['inside', '안쪽 테두리', 'borderAll'], ['none', '테두리 없음', 'borderNone'], ['top', '위쪽 테두리', 'borderTop'], ['bottom', '아래쪽 테두리', 'borderBottom'], ['left', '왼쪽 테두리', 'borderLeft'], ['right', '오른쪽 테두리', 'borderRight'], ['insideH', '안쪽 가로 테두리', 'borderAll'], ['insideV', '안쪽 세로 테두리', 'borderAll']]
    .map(([k, l, ic]) => ({ label: l, icon: ic, action: () => run('cellBorders', k, S.borderPen) }));
}

// ───────────── 목록 갤러리 ─────────────
const BULLETS = ['●', '○', '■', '◆', '➢', '✓', '–', '★', '※', '▶', '◇', '□'];
const NUM_STYLES = [['decimal', '.', '1.\n2.\n3.'], ['decimal', ')', '1)\n2)\n3)'], ['upperRoman', '.', 'I.\nII.\nIII.'], ['upperLetter', '.', 'A.\nB.\nC.'], ['lowerLetter', ')', 'a)\nb)\nc)'], ['lowerLetter', '.', 'a.\nb.\nc.'], ['lowerRoman', '.', 'i.\nii.\niii.'], ['ganada', '.', '가.\n나.\n다.'], ['ganada', ')', '가)\n나)\n다)'], ['decimalCircle', '', '①\n②\n③'], ['decimal', '()', '(1)\n(2)\n(3)'], ['chosung', '.', 'ㄱ.\nㄴ.\nㄷ.']];
register({
  bulletsMenu: (a) => {
    const g = el('div', { class: 'list-gal' });
    g.append(el('button', { onclick: () => { closeMenus(); run('removeList'); } }, '없음'));
    for (const b of BULLETS) g.append(el('button', { style: { fontSize: '20px', textAlign: 'center' }, onclick: () => { closeMenus(); run('bullets', b); } }, b));
    openMenu(a, [{ title: '글머리 기호 라이브러리' }, { node: g }, { sep: true }, { label: '목록 수준 변경', submenu: Array.from({ length: 9 }, (_, i) => ({ label: `수준 ${i + 1}`, action: () => run('setListLevel', i) })) }, { label: '새 글머리 기호 정의...', action: () => run('defineBulletDialog') }]);
  },
  numberingMenu: (a) => {
    const g = el('div', { class: 'list-gal' });
    g.append(el('button', { onclick: () => { closeMenus(); run('removeList'); } }, '없음'));
    for (const [fmt, suf, txt] of NUM_STYLES) g.append(el('button', { onclick: () => { closeMenus(); run('numbering', fmt, suf); } }, txt));
    openMenu(a, [{ title: '번호 매기기 라이브러리' }, { node: g }, { sep: true }, { label: '목록 수준 변경', submenu: Array.from({ length: 9 }, (_, i) => ({ label: `수준 ${i + 1}`, action: () => run('setListLevel', i) })) }, { label: '번호 다시 매기기', action: () => run('restartNumbering') }, { label: '이전 목록에 이어서', action: () => run('continueNumbering') }, { label: '번호 매기기 값 설정...', action: () => run('setNumberValueDialog') }]);
  },
  multilevelMenu: (a) => {
    const g = el('div', { class: 'list-gal' });
    g.append(el('button', { onclick: () => { closeMenus(); run('removeList'); } }, '없음'));
    g.append(el('button', { onclick: () => { closeMenus(); run('multilevel', 'outline'); } }, '1.\n  a)\n    i.'));
    g.append(el('button', { onclick: () => { closeMenus(); run('multilevel', 'legal'); } }, '1.\n  1.1.\n    1.1.1.'));
    g.append(el('button', { onclick: () => { closeMenus(); run('multilevel', 'korean'); } }, '1.\n 가.\n  1)\n   가)'));
    g.append(el('button', { onclick: () => { closeMenus(); run('headingNumbering'); } }, '1 제목 1\n1.1 제목 2\n1.1.1 제목 3'));
    openMenu(a, [{ title: '목록 라이브러리' }, { node: g }, { sep: true }, { label: '목록 수준 변경', submenu: Array.from({ length: 9 }, (_, i) => ({ label: `수준 ${i + 1}`, action: () => run('setListLevel', i) })) }]);
  },
  headingNumbering: () => {
    // 제목 1~3 스타일에 번호 (1 / 1.1 / 1.1.1)
    const def = multiList('legal');
    def.levels = def.levels.map((l, i) => ({ ...l, text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.'), ind: { l: 0, hanging: 0 }, suff: 'space' }));
    change((d) => {
      const ls = d.prop('lists');
      const id = String(Math.max(0, ...Object.keys(ls).map(Number).filter(Number.isFinite)) + 1);
      ls[id] = { ...def, key: 'headingNum' };
      const st = d.prop('styles');
      for (let i = 1; i <= 9; i++) st[`Heading${i}`] = { ...st[`Heading${i}`], pPr: { ...st[`Heading${i}`].pPr, num: { id, lvl: i - 1 } } };
      return S.sel;
    });
  },
  lineSpacingMenu: (a) => {
    const cur = (() => { const p = blockAt(S.doc, sel().f.p); return p ? effPara(S.doc, p).spc?.line : null; })();
    openMenu(a, [
      ...[1, 1.08, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: String(v.toFixed(v % 1 ? 2 : 1).replace(/0$/, '')), checked: Math.abs((cur ?? 0) - v) < 0.01, action: () => run('lineSpacing', v) })),
      { sep: true },
      { label: '줄 간격 옵션...', action: () => run('paragraphDialog') },
      { sep: true },
      { label: '단락 앞에 공백 추가/제거', action: () => run('toggleSpaceBefore') },
      { label: '단락 뒤에 공백 추가/제거', action: () => run('toggleSpaceAfter') },
      { sep: true },
      { title: '한글 문서 (글자 크기의 %)' },
      ...[1.3, 1.6, 2].map((v) => ({ label: `${Math.round(v * 100)}%`, action: () => run('paragraphProps', { spc: { ...(blockAt(S.doc, sel().f.p)?.pPr?.spc ?? {}), line: v, rule: 'pct' } }) })),
    ]);
  },
  findMenu: (a) => openMenu(a, [
    { label: '찾기 (Ctrl+F)', icon: 'find', action: () => run('findPane') },
    { label: '고급 찾기...', action: () => run('replaceDialog', 'find') },
    { label: '이동... (Ctrl+G)', action: () => run('replaceDialog', 'goto') },
  ]),
  selectMenu: (a) => openMenu(a, [
    { label: '모두 선택 (Ctrl+A)', action: () => run('selectAll') },
    { label: '비슷한 서식의 텍스트 선택', action: () => run('selectSimilar') },
    { label: '현재 단락 선택', action: () => { const p = blockAt(S.doc, sel().f.p); run('__setSel', { p: sel().f.p, o: 0 }, { p: sel().f.p, o: paraLen(p) }); } },
  ]),
});

// ───────────── 삽입 메뉴 ─────────────
function gridPicker(onPick, maxR = 8, maxC = 10) {
  const title = el('div', { class: 'gp-title' }, '표 삽입');
  const tb = el('table');
  for (let r = 0; r < maxR; r++) { const tr = el('tr'); for (let c = 0; c < maxC; c++) tr.append(el('td', { 'data-r': r, 'data-c': c })); tb.append(tr); }
  tb.addEventListener('mouseover', (e) => {
    const td = e.target.closest('td'); if (!td) return;
    const R = Number(td.dataset.r); const Cc = Number(td.dataset.c);
    for (const x of tb.querySelectorAll('td')) x.classList.toggle('on', Number(x.dataset.r) <= R && Number(x.dataset.c) <= Cc);
    title.textContent = `표 ${Cc + 1}×${R + 1}`;
  });
  tb.addEventListener('click', (e) => { const td = e.target.closest('td'); if (!td) return; closeMenus(); onPick(Number(td.dataset.r) + 1, Number(td.dataset.c) + 1); });
  return el('div', { class: 'grid-pick' }, title, tb);
}
const PAGE_NUM_POS = [['top', 'left', '페이지 위쪽 · 왼쪽'], ['top', 'center', '페이지 위쪽 · 가운데'], ['top', 'right', '페이지 위쪽 · 오른쪽'], ['bottom', 'left', '페이지 아래쪽 · 왼쪽'], ['bottom', 'center', '페이지 아래쪽 · 가운데'], ['bottom', 'right', '페이지 아래쪽 · 오른쪽']];
const HEADER_PRESETS = [
  ['비어 있음', () => [para('', { style: 'Header' })]],
  ['비어 있음 (3단)', () => [{ ...para('', { style: 'Header' }), runs: [{ text: '[여기에 입력]' }, { obj: 'tab' }, { text: '[여기에 입력]' }, { obj: 'tab' }, { text: '[여기에 입력]' }] }]],
  ['제목', () => [{ ...para('', { style: 'Header', pPr: { align: 'center' } }), runs: [{ obj: 'field', code: 'TITLE', result: S.doc.props.title || S.docName }] }]],
  ['제목 + 날짜', () => [{ ...para('', { style: 'Header' }), runs: [{ obj: 'field', code: 'TITLE', result: S.doc.props.title || S.docName }, { obj: 'tab' }, { obj: 'tab' }, { obj: 'field', code: 'DATE', result: new Date().toLocaleDateString('ko-KR') }] }]],
  ['밑줄 있는 제목', () => [{ ...para('', { style: 'Header', pPr: { border: { bottom: { style: 'single', w: 0.75, color: '#2B579A', space: 2 } } } }), runs: [{ obj: 'field', code: 'TITLE', result: S.doc.props.title || S.docName, color: '#2B579A' }] }]],
  ['기관명 (공문서)', () => [{ ...para('', { style: 'Header', pPr: { align: 'center' } }), runs: [{ text: '[기관명]', b: true }] }]],
];
const FOOTER_PRESETS = [
  ['비어 있음', () => [para('', { style: 'Footer' })]],
  ['페이지 번호 (가운데)', () => [{ ...para('', { style: 'Footer', pPr: { align: 'center' } }), runs: [{ obj: 'field', code: 'PAGE', result: '1' }] }]],
  ['- 페이지 번호 -', () => [{ ...para('', { style: 'Footer', pPr: { align: 'center' } }), runs: [{ text: '- ' }, { obj: 'field', code: 'PAGE', result: '1' }, { text: ' -' }] }]],
  ['페이지 X / Y', () => [{ ...para('', { style: 'Footer', pPr: { align: 'center' } }), runs: [{ obj: 'field', code: 'PAGE', result: '1' }, { text: ' / ' }, { obj: 'field', code: 'NUMPAGES', result: '1' }] }]],
  ['작성자 + 페이지', () => [{ ...para('', { style: 'Footer' }), runs: [{ obj: 'field', code: 'AUTHOR', result: S.doc.props.author || '작성자' }, { obj: 'tab' }, { obj: 'tab' }, { obj: 'field', code: 'PAGE', result: '1' }] }]],
];
const COVERS = [
  ['기본', (t) => [para('', { pPr: { spc: { before: cm(7) } } }), para(t || '[문서 제목]', { style: 'Title', pPr: { align: 'center' } }), para('[문서 부제목]', { style: 'Subtitle', pPr: { align: 'center' } }), { ...para('', { pPr: { align: 'center', spc: { before: cm(6) } } }), runs: [{ text: `${S.doc.props.author || '[작성자]'}  ·  ${new Date().toLocaleDateString('ko-KR')}` }] }]],
  ['띠 (파랑)', (t) => [{ ...para('', { pPr: { shd: '#2B579A', spc: { before: cm(5), after: 0 }, ind: { l: 12 } } }), runs: [{ text: t || '[문서 제목]', sz: 30, color: '#FFFFFF', b: true }] }, { ...para('', { pPr: { shd: '#2B579A', ind: { l: 12 } } }), runs: [{ text: '[부제목]', sz: 14, color: '#DDEBF7' }] }]],
  ['보고서 (공공기관)', (t) => [para('', { pPr: { spc: { before: cm(6) } } }), { ...para('', { pPr: { align: 'center', border: { top: { style: 'double', w: 1.5, color: '#000000', space: 6 }, bottom: { style: 'double', w: 1.5, color: '#000000', space: 6 } } } }), runs: [{ text: t || '[보고서 제목]', sz: 24, b: true }] }, para('', { pPr: { spc: { before: cm(8) } } }), { ...para('', { pPr: { align: 'center' } }), runs: [{ text: `${new Date().getFullYear()}. ${new Date().getMonth() + 1}.`, sz: 14 }] }, { ...para('', { pPr: { align: 'center' } }), runs: [{ text: '[기관명]', sz: 18, b: true }] }]],
];
register({
  coverPageMenu: (a) => openMenu(a, [
    { title: '기본 제공' },
    ...COVERS.map(([l, f]) => ({ label: l, icon: 'coverPage', action: () => run('insertCover', l) })),
    { sep: true },
    { label: '현재 표지 제거', action: () => run('removeCover') },
  ]),
  insertCover: (name) => {
    const f = COVERS.find((c) => c[0] === name)?.[1];
    if (!f) return;
    const title = S.doc.props.title;
    const blocks = f(title).map((b) => ({ ...b, cover: true }));
    change((d) => { d.doc.body = [...blocks, { ...para(''), cover: true, runs: [{ obj: 'br', kind: 'page' }] }, ...d.doc.body]; for (const b of blocks) d.owned.add(b); return { p: [0], o: 0 }; });
  },
  removeCover: () => change((d) => { d.doc.body = d.doc.body.filter((b) => !b.cover); if (!d.doc.body.length) d.doc.body = [para('')]; return { p: [0], o: 0 }; }),
  blankPage: () => { run('pageBreak'); run('pageBreak'); run('__moveLeft'); },
  tableMenu: (a) => openMenu(a, [
    { node: gridPicker((r, c) => run('insertTable', r, c)) },
    { sep: true },
    { label: '표 삽입...', icon: 'table', action: () => run('insertTableDialog') },
    { label: '텍스트를 표로 변환...', icon: 'tableConvert', action: () => run('textToTable', '\t') },
    { label: 'Excel 스프레드시트 붙여넣기 (CSV)...', icon: 'csvIn', action: () => run('insertCsvTable') },
    { label: '빠른 표', submenu: QUICK_TABLES.map(([l]) => ({ label: l, action: () => run('quickTable', l) })) },
  ]),
  pictureMenu: (a) => openMenu(a, [
    { label: '이 디바이스...', icon: 'picture', action: () => run('insertPicture') },
    { label: '온라인 그림 (주소)...', icon: 'webData', action: () => run('insertPictureUrl') },
  ]),
  shapesMenu: (a) => {
    const g = el('div', { class: 'list-gal', style: { gridTemplateColumns: 'repeat(6, 40px)' } });
    for (const [k, svg] of Object.entries(SHAPES)) g.append(el('button', { title: k, style: { height: '40px', padding: '4px' }, html: `<svg viewBox="0 0 100 100" width="28" height="28">${svg('#4472C4', '#2F528F')}</svg>`, onclick: () => { closeMenus(); run('insertShape', k); } }));
    openMenu(a, [{ title: '도형 (그림으로 넣습니다)' }, { node: g }]);
  },
  insertShape: (k) => {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" width="100%" height="100%">${SHAPES[k]('#4472C4', '#2F528F')}</svg>`;
    run('insertImage', `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`, 120, k.includes('선') || k.includes('화살표') ? 60 : 120, k);
  },
  chartMenu: (a) => openMenu(a, [
    { label: '세로 막대형', icon: 'chartColumn', action: () => run('insertChartDialog', 'col') },
    { label: '가로 막대형', icon: 'chartBar', action: () => run('insertChartDialog', 'bar') },
    { label: '꺾은선형', icon: 'chartLine', action: () => run('insertChartDialog', 'line') },
    { label: '원형', icon: 'chartPie', action: () => run('insertChartDialog', 'pie') },
  ]),
  headerMenu: (a) => openMenu(a, [
    { title: '기본 제공' },
    ...HEADER_PRESETS.map(([l, f]) => ({ label: l, icon: 'header', action: () => { run('setHeader', f(), 'default', 'headers'); run('editHeaderFooterAt', 'header'); } })),
    { sep: true },
    { label: '머리글 편집', icon: 'editor', action: () => run('editHeaderFooterAt', 'header') },
    { label: '머리글 제거', icon: 'delete', action: () => run('removeHeader', 'headers') },
  ]),
  footerMenu: (a) => openMenu(a, [
    { title: '기본 제공' },
    ...FOOTER_PRESETS.map(([l, f]) => ({ label: l, icon: 'footer', action: () => { run('setHeader', f(), 'default', 'footers'); run('editHeaderFooterAt', 'footer'); } })),
    { sep: true },
    { label: '바닥글 편집', icon: 'editor', action: () => run('editHeaderFooterAt', 'footer') },
    { label: '바닥글 제거', icon: 'delete', action: () => run('removeHeader', 'footers') },
  ]),
  pageNumberMenu: (a) => openMenu(a, [
    ...PAGE_NUM_POS.map(([v, h, l]) => ({ label: l, icon: 'pageNumber', action: () => run('insertPageNumber', v, h) })),
    { sep: true },
    { label: '현재 위치', action: () => run('insertField', 'PAGE') },
    { label: '페이지 번호 서식...', action: () => run('pageNumberFormatDialog') },
    { label: '페이지 번호 제거', action: () => run('removePageNumbers') },
  ]),
  insertPageNumber: (v, h) => {
    const kind = v === 'top' ? 'headers' : 'footers';
    run('setHeader', [{ ...para('', { style: v === 'top' ? 'Header' : 'Footer', pPr: { align: h } }), runs: [{ obj: 'field', code: 'PAGE', result: '1' }] }], 'default', kind);
  },
  removePageNumbers: () => {
    const strip = (bs) => bs?.map((b) => (b.type === 'p' ? { ...b, runs: b.runs.filter((r) => !(r.obj === 'field' && /^(PAGE|NUMPAGES)/i.test(r.code ?? ''))) } : b));
    change((d) => {
      const fix = (sec) => ({ ...sec, headers: Object.fromEntries(Object.entries(sec.headers ?? {}).map(([k, v]) => [k, strip(v)])), footers: Object.fromEntries(Object.entries(sec.footers ?? {}).map(([k, v]) => [k, strip(v)])) });
      d.doc.section = fix(d.doc.section);
      d.doc.body.forEach((b, i) => { if (b.secBreak) { const x = d.block([i]); x.secBreak = { ...x.secBreak, sec: fix(x.secBreak.sec) }; } });
      return S.sel;
    });
  },
  textboxMenu: (a) => openMenu(a, [
    { label: '간단한 글 상자', action: () => run('insertTextBox', 'simple') },
    { label: '강조 상자 (파랑)', action: () => run('insertTextBox', 'blue') },
    { label: '인용 상자', action: () => run('insertTextBox', 'quote') },
    { label: '알림 상자 (노랑)', action: () => run('insertTextBox', 'note') },
  ]),
  insertTextBox: (kind) => {
    const sty = { simple: { border: { top: { style: 'single', w: 0.75, color: '#000000', space: 4 }, bottom: { style: 'single', w: 0.75, color: '#000000', space: 4 }, left: { style: 'single', w: 0.75, color: '#000000', space: 4 }, right: { style: 'single', w: 0.75, color: '#000000', space: 4 } } }, blue: { shd: '#DEEAF6', border: { left: { style: 'single', w: 4, color: '#2B579A', space: 8 } } }, quote: { shd: '#F2F2F2', ind: { l: 24, r: 24 }, border: { left: { style: 'single', w: 3, color: '#A5A5A5', space: 8 } } }, note: { shd: '#FFF2CC', border: { top: { style: 'single', w: 0.75, color: '#FFC000', space: 4 }, bottom: { style: 'single', w: 0.75, color: '#FFC000', space: 4 }, left: { style: 'single', w: 0.75, color: '#FFC000', space: 4 }, right: { style: 'single', w: 0.75, color: '#FFC000', space: 4 } } } }[kind];
    run('insertBlocksAt', [para('[글 상자에 내용을 입력하세요]', { pPr: { ...sty, spc: { before: 8, after: 8 } } }), para('')]);
  },
  quickPartsMenu: (a) => openMenu(a, [
    { title: '문서 속성' },
    { label: '제목', action: () => run('insertField', 'TITLE') },
    { label: '작성자', action: () => run('insertField', 'AUTHOR') },
    { label: '날짜', action: () => run('insertField', 'DATE') },
    { label: '단어 수', action: () => run('insertField', 'NUMWORDS') },
    { sep: true },
    { label: '페이지 번호', action: () => run('insertField', 'PAGE') },
    { label: '전체 페이지 수', action: () => run('insertField', 'NUMPAGES') },
    { sep: true },
    { label: '필드...', action: () => run('fieldDialog') },
  ]),
  dropCapMenu: (a) => openMenu(a, [
    { label: '없음', action: () => run('dropCap', 'none') },
    { label: '본문', action: () => run('dropCap', 'drop') },
    { label: '여백', action: () => run('dropCap', 'margin') },
  ]),
  objectMenu: (a) => openMenu(a, [
    { label: '개체... (다른 문서 넣기)', action: () => run('insertFileContents') },
    { label: '파일의 텍스트...', action: () => run('insertFileContents') },
  ]),
  equationMenu: (a) => openMenu(a, [
    { title: '기본 제공' },
    ...[['이차 방정식의 근', 'x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}'], ['피타고라스 정리', 'a^2+b^2=c^2'], ['원의 넓이', 'A=\\pi r^2'], ['합', '\\sum_{i=1}^{n} i=\\frac{n(n+1)}{2}'], ['적분', '\\int_0^1 x^2\\,dx=\\frac{1}{3}'], ['이항 정리', '(x+a)^n=\\sum_{k=0}^{n}\\binom{n}{k}x^k a^{n-k}']].map(([l, tex]) => ({ label: l, action: () => run('insertEquation', tex) })),
    { sep: true },
    { label: '새 수식 삽입 (Alt+=)', action: () => run('equationDialog') },
  ]),
  symbolMenu: (a) => {
    const g = el('div', { class: 'list-gal', style: { gridTemplateColumns: 'repeat(8, 30px)' } });
    for (const ch of '€£¥©®™±≠≤≥÷×∞µαβπΩ∑√°※★☆○●◎◇◆□■△▲▽▼→←↑↓↔「」『』【】①②③④⑤ㄱㄴ㉠㉡㉮㉯'.split('').concat(['…', '·', '―'])) g.append(el('button', { style: { height: '30px', fontSize: '16px', textAlign: 'center', padding: '0' }, onclick: () => { closeMenus(); run('insertText', ch); } }, ch));
    openMenu(a, [{ node: g }, { sep: true }, { label: '다른 기호...', action: () => run('symbolDialog') }]);
  },
});
const QUICK_TABLES = [
  ['달력 (이번 달)', () => { const d = new Date(); const y = d.getFullYear(); const m = d.getMonth(); const first = new Date(y, m, 1).getDay(); const days = new Date(y, m + 1, 0).getDate(); const rows = [['일', '월', '화', '수', '목', '금', '토']]; let row = Array(7).fill(''); for (let i = 0; i < first; i++) row[i] = ''; for (let dd = 1; dd <= days; dd++) { const k = (first + dd - 1) % 7; row[k] = String(dd); if (k === 6 || dd === days) { rows.push(row); row = Array(7).fill(''); } } return { title: `${y}년 ${m + 1}월`, rows, style: 'GridTable1Light-Accent1' }; }],
  ['일정표', () => ({ rows: [['시간', '내용', '담당', '비고'], ['09:00', '', '', ''], ['10:00', '', '', ''], ['11:00', '', '', ''], ['13:00', '', '', ''], ['14:00', '', '', '']], style: 'GridTable4-Accent1' })],
  ['예산표', () => ({ rows: [['항목', '예산', '집행', '잔액'], ['인건비', '', '', ''], ['운영비', '', '', ''], ['합계', '', '', '']], style: 'GridTable4-Accent6' })],
  ['결재란', () => ({ rows: [['담당', '팀장', '본부장', '대표'], ['', '', '', '']], style: 'TableGrid', align: 'right', h: [24, 48] })],
];
register({
  quickTable: (name) => {
    const q = QUICK_TABLES.find((x) => x[0] === name)?.[1]?.();
    if (!q) return;
    run('insertTable', q.rows.length, q.rows[0].length, q.style);
    change((d) => {
      const pos = S.sel.f;
      const tpath = pos.p.slice(0, -3);
      const t = d.block(tpath);
      t.rows = t.rows.map((row, r) => ({ ...row, header: r === 0, ...(q.h ? { h: q.h[r] ?? q.h[q.h.length - 1] } : {}), cells: row.cells.map((c, k) => ({ ...c, vAlign: 'center', blocks: [{ ...para(q.rows[r][k] ?? ''), pPr: { align: 'center', spc: { after: 0 } } }] })) }));
      if (q.align) { t.tblPr = { ...t.tblPr, align: q.align }; const w = 70 * t.grid.length; t.grid = t.grid.map(() => 70); void w; }
      return S.sel;
    });
    if (q.title) run('insertBlocksAt', [para(q.title, { style: 'Heading3' })]);
  },
});
const SHAPES = {
  '사각형': (f, s) => `<rect x="2" y="2" width="96" height="96" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '둥근 사각형': (f, s) => `<rect x="2" y="2" width="96" height="96" rx="16" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '타원': (f, s) => `<ellipse cx="50" cy="50" rx="48" ry="48" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '삼각형': (f, s) => `<polygon points="50,2 98,98 2,98" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '다이아몬드': (f, s) => `<polygon points="50,2 98,50 50,98 2,50" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '오각형': (f, s) => `<polygon points="50,2 98,38 80,98 20,98 2,38" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '별': (f, s) => `<polygon points="50,2 61,37 98,37 68,59 79,96 50,74 21,96 32,59 2,37 39,37" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '오른쪽 화살표': (f, s) => `<polygon points="2,30 62,30 62,8 98,50 62,92 62,70 2,70" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '왼쪽 화살표': (f, s) => `<polygon points="98,30 38,30 38,8 2,50 38,92 38,70 98,70" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '하트': (f, s) => `<path d="M50 92 C10 62 2 40 18 20 C32 4 48 12 50 26 C52 12 68 4 82 20 C98 40 90 62 50 92Z" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '말풍선': (f, s) => `<path d="M4 6 H96 V70 H40 L20 94 L24 70 H4 Z" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '선': (f, s) => `<line x1="2" y1="50" x2="98" y2="50" stroke="${s}" stroke-width="4"/>`,
  '화살표 선': (f, s) => `<line x1="2" y1="50" x2="86" y2="50" stroke="${s}" stroke-width="4"/><polygon points="98,50 82,40 82,60" fill="${s}"/>`,
  '육각형': (f, s) => `<polygon points="25,4 75,4 98,50 75,96 25,96 2,50" fill="${f}" stroke="${s}" stroke-width="2"/>`,
  '원통': (f, s) => `<path d="M2 16 C2 4 98 4 98 16 V84 C98 96 2 96 2 84Z" fill="${f}" stroke="${s}" stroke-width="2"/><path d="M2 16 C2 28 98 28 98 16" fill="none" stroke="${s}" stroke-width="2"/>`,
  '체크': (f, s) => `<polyline points="6,52 38,84 94,16" fill="none" stroke="${s}" stroke-width="10" stroke-linecap="round" stroke-linejoin="round"/>`,
  '엑스': (f, s) => `<path d="M10 10 90 90M90 10 10 90" stroke="#C00000" stroke-width="10" stroke-linecap="round"/>`,
  '구름': (f, s) => `<path d="M25 80 C8 80 4 60 18 54 C12 36 34 28 42 36 C46 18 74 18 76 38 C94 34 100 58 86 66 C92 80 76 86 66 80 Z" fill="${f}" stroke="${s}" stroke-width="2"/>`,
};

// ───────────── 디자인 메뉴 ─────────────
const THEME_FONT_PAIRS = [['맑은 고딕', '맑은 고딕', 'Office (한국어)'], ['맑은 고딕 Semilight', '맑은 고딕', 'Office 2013-2022'], ['함초롬돋움', '함초롬바탕', '한글 기본'], ['나눔고딕', '나눔명조', '나눔'], ['Pretendard', 'Pretendard', 'Pretendard'], ['Noto Sans KR', 'Noto Serif KR', 'Noto'], ['Calibri Light', 'Calibri', 'Office (영문)'], ['Cambria', 'Calibri', 'Cambria'], ['Georgia', 'Georgia', 'Georgia'], ['Arial', 'Times New Roman', '고전']];
const THEME_COLOR_SETS = [['Office', ['#4472C4', '#2F5496', '#1F3763']], ['파랑 II', ['#0F6FC6', '#0B5294', '#073763']], ['녹색', ['#548235', '#385723', '#203316']], ['주황', ['#C55A11', '#843C0C', '#5A2908']], ['빨강', ['#C00000', '#7F0000', '#560000']], ['보라', ['#7030A0', '#4E2272', '#311648']], ['회색조', ['#404040', '#262626', '#0D0D0D']], ['청록', ['#00808A', '#005C63', '#00393D']]];
register({
  themeMenu: (a) => openMenu(a, [
    { title: '테마 (글꼴 + 색)' },
    ...THEME_FONT_PAIRS.slice(0, 6).map(([mj, mn, l], i) => ({ label: l, action: () => { run('themeFonts', mj, mn); run('themeColors', THEME_COLOR_SETS[i % THEME_COLOR_SETS.length][0]); } })),
  ]),
  themeColorsMenu: (a) => openMenu(a, THEME_COLOR_SETS.map(([l, cols]) => ({ label: l, node: el('button', { class: 'menu-item', onclick: () => { closeMenus(); run('themeColors', l); } }, el('span', { class: 'mi-icon' }), ...cols.map((c) => el('i', { style: { display: 'inline-block', width: '14px', height: '14px', background: c, marginRight: '2px' } })), el('span', { style: { marginLeft: '8px' } }, l)) }))),
  themeColors: (name) => {
    const set = THEME_COLOR_SETS.find((x) => x[0] === name)?.[1];
    if (!set) return;
    change((d) => {
      const st = d.prop('styles');
      for (let i = 1; i <= 9; i++) { const k = `Heading${i}`; const c = i <= 2 || i === 4 || i === 5 ? set[1] : set[2]; st[k] = { ...st[k], rPr: { ...st[k].rPr, color: i === 1 || i === 2 ? set[1] : i >= 8 ? '#272727' : c } }; }
      st.IntenseQuote = { ...st.IntenseQuote, rPr: { ...st.IntenseQuote.rPr, color: set[0] }, pPr: { ...st.IntenseQuote.pPr, border: { top: { style: 'single', w: 0.5, color: set[0], space: 10 }, bottom: { style: 'single', w: 0.5, color: set[0], space: 10 } } } };
      st.Hyperlink = { ...st.Hyperlink };
      d.doc.themeColors = name;
      return S.sel;
    });
  },
  themeFontsMenu: (a) => openMenu(a, THEME_FONT_PAIRS.map(([mj, mn, l]) => ({ label: l, node: el('button', { class: 'menu-item', onclick: () => { closeMenus(); run('themeFonts', mj, mn); } }, el('span', { class: 'mi-icon' }), el('span', { class: 'mi-text' }, el('b', { style: { fontFamily: `'${mj}'` } }, `${l}`), el('small', { style: { fontFamily: `'${mn}'` } }, `${mj} / ${mn}`))) })), { scroll: true }),
  paraSpacingMenu: (a) => openMenu(a, [
    ...[['default', '스타일 세트 기본값'], ['none', '단락 간격 없음'], ['compact', '좁게'], ['tight', '조금 좁게'], ['open', '넓게'], ['relaxed', '느슨하게'], ['double', '두 줄']].map(([k, l]) => ({ label: l, action: () => run('paragraphSpacingPreset', k) })),
    { sep: true }, { label: '사용자 지정 단락 간격...', action: () => run('paragraphDialog') },
  ]),
  setDefaultFormat: () => { try { localStorage.setItem('wiword:defaults', JSON.stringify({ styles: S.doc.styles, defaults: S.doc.defaults, themeFonts: S.doc.themeFonts })); toast('새 문서에 이 서식을 씁니다'); } catch { toast('저장하지 못했습니다'); } },
  watermarkMenu: (a) => openMenu(a, [
    { title: '기밀' }, { label: '기밀', action: () => run('watermark', '기밀') }, { label: '대외비', action: () => run('watermark', '대외비') },
    { title: '고지 사항' }, { label: '초안', action: () => run('watermark', '초안') }, { label: '샘플', action: () => run('watermark', '샘플') }, { label: '복사 금지', action: () => run('watermark', '복사 금지') }, { label: '긴급', action: () => run('watermark', '긴급') },
    { sep: true },
    { label: '사용자 지정 워터마크...', action: () => run('watermarkDialog') },
    { label: '워터마크 제거', action: () => run('watermark', '') },
  ]),
  pageColorMenu: (a) => openMenu(a, [{ node: colorGrid((c) => run('pageColor', c), { auto: null, none: '색 없음', recentKey: 'page' }) }]),
});

// ───────────── 레이아웃 메뉴 ─────────────
const mcm = (v) => `${f1(pxToCm(v))}cm`;
register({
  marginsMenu: (a) => openMenu(a, [
    ...Object.entries(MARGINS).map(([k, m]) => ({ label: m.label, desc: `위: ${mcm(m.t)}  아래: ${mcm(m.b)}  왼쪽: ${mcm(m.l)}  오른쪽: ${mcm(m.r)}`, icon: 'margins', action: () => run('margins', k) })),
    { sep: true }, { label: '사용자 지정 여백...', action: () => run('pageSetupDialog', 'margins') },
  ]),
  orientationMenu: (a) => openMenu(a, [{ label: '세로', icon: 'onePage', action: () => run('orientation', 'portrait') }, { label: '가로', icon: 'orientation', action: () => run('orientation', 'landscape') }]),
  sizeMenu: (a) => openMenu(a, [...Object.entries(PAPERS).map(([k, p]) => ({ label: p.label, desc: p.size, icon: 'pageSize', action: () => run('paperSize', k) })), { sep: true }, { label: '다른 용지 크기...', action: () => run('pageSetupDialog', 'paper') }]),
  columnsMenu: (a) => openMenu(a, [
    { label: '1단', icon: 'onePage', action: () => run('columnsPreset', 'one') },
    { label: '2단', icon: 'multiPage', action: () => run('columnsPreset', 'two') },
    { label: '3단', icon: 'columns', action: () => run('columnsPreset', 'three') },
    { label: '왼쪽', action: () => run('columnsPreset', 'left') },
    { label: '오른쪽', action: () => run('columnsPreset', 'right') },
    { sep: true }, { label: '기타 단...', action: () => run('columnsDialog') },
  ]),
  breaksMenu: (a) => openMenu(a, [
    { title: '페이지 나누기' },
    { label: '페이지', desc: '한 페이지가 끝나고 다음 페이지가 시작되는 지점을 표시합니다.', icon: 'pageBreak', action: () => run('pageBreak') },
    { label: '단', desc: '단 나누기 다음의 텍스트는 다음 단에서 시작됩니다.', icon: 'columns', action: () => run('columnBreak') },
    { label: '텍스트 줄 바꿈', desc: '웹 페이지 개체 주위의 텍스트를 구분합니다.', action: () => run('lineBreak') },
    { title: '구역 나누기' },
    { label: '다음 페이지부터', desc: '구역 나누기를 삽입하고 다음 페이지에서 새 구역을 시작합니다.', icon: 'breaks', action: () => run('sectionBreak', 'nextPage') },
    { label: '이어서', desc: '구역 나누기를 삽입하고 같은 페이지에서 새 구역을 시작합니다.', icon: 'breaks', action: () => run('sectionBreak', 'continuous') },
    { label: '짝수 페이지부터', desc: '다음 짝수 페이지에서 새 구역을 시작합니다.', icon: 'breaks', action: () => run('sectionBreak', 'evenPage') },
    { label: '홀수 페이지부터', desc: '다음 홀수 페이지에서 새 구역을 시작합니다.', icon: 'breaks', action: () => run('sectionBreak', 'oddPage') },
  ]),
  lineNumbersMenu: (a) => openMenu(a, [
    { label: '없음', action: () => run('lineNumbers', 'none') },
    { label: '연속', action: () => run('lineNumbers', 'continuous') },
    { label: '페이지마다 다시 매기기', action: () => run('lineNumbers', 'page') },
    { label: '구역마다 다시 매기기', action: () => run('lineNumbers', 'section') },
  ]),
  hyphenMenu: (a) => openMenu(a, [
    { label: '없음', checked: !S.doc.settings.hyphenate, action: () => change((d) => { d.doc.settings = { ...d.doc.settings, hyphenate: false }; return S.sel; }) },
    { label: '자동', checked: !!S.doc.settings.hyphenate, action: () => change((d) => { d.doc.settings = { ...d.doc.settings, hyphenate: true }; return S.sel; }) },
  ]),
  positionMenu: (a) => openMenu(a, [
    { label: '텍스트 줄 안', action: () => run('imageWrap', 'inline') },
    { sep: true },
    ...[['left', '왼쪽 위'], ['center', '가운데 위'], ['right', '오른쪽 위']].map(([s, l]) => ({ label: `${l} (텍스트 배치: 정사각형)`, action: () => run('imageWrap', 'square', s) })),
  ]),
  wrapMenu: (a) => openMenu(a, [
    { label: '텍스트 줄 안 (인라인)', action: () => run('imageWrap', 'inline') },
    { label: '정사각형 (왼쪽)', action: () => run('imageWrap', 'square', 'left') },
    { label: '정사각형 (오른쪽)', action: () => run('imageWrap', 'square', 'right') },
    { label: '위/아래', action: () => run('imageWrap', 'topBottom') },
    { label: '텍스트 뒤', action: () => run('imageWrap', 'behind') },
    { label: '텍스트 앞', action: () => run('imageWrap', 'front') },
  ]),
  objAlignMenu: (a) => openMenu(a, [{ label: '왼쪽 맞춤', action: () => run('imageWrap', 'square', 'left') }, { label: '오른쪽 맞춤', action: () => run('imageWrap', 'square', 'right') }, { label: '가운데 맞춤 (인라인)', action: () => { run('imageWrap', 'inline'); run('alignCenter'); } }]),
  rotateMenu: (a) => openMenu(a, [{ label: '오른쪽으로 90도 회전', action: () => run('rotateImage', 90) }, { label: '왼쪽으로 90도 회전', action: () => run('rotateImage', -90) }, { label: '회전 없음', action: () => run('rotateImage', 0) }]),
});

// ───────────── 참조 메뉴 ─────────────
register({
  tocMenu: (a) => openMenu(a, [
    { title: '기본 제공' },
    { label: '자동 목차 1', desc: '제목 1~3', icon: 'toc', action: () => run('insertToc', 3) },
    { label: '자동 목차 2', desc: '제목 1~2', icon: 'toc', action: () => run('insertToc', 2) },
    { label: '자세한 목차', desc: '제목 1~4', icon: 'toc', action: () => run('insertToc', 4) },
    { sep: true },
    { label: '목차 업데이트', icon: 'updateToc', action: () => run('updateToc') },
    { label: '목차 제거', icon: 'delete', action: () => run('removeToc') },
  ]),
  removeToc: () => change((d) => { d.doc.body = d.doc.body.filter((b) => !b.toc); if (!d.doc.body.length) d.doc.body = [para('')]; return { p: [0], o: 0 }; }),
  addTextMenu: (a) => openMenu(a, [
    { label: '목차에 표시 안 함', action: () => run('paragraphProps', { outline: 9 }) },
    { label: '수준 1', action: () => run('applyStyle', 'Heading1') },
    { label: '수준 2', action: () => run('applyStyle', 'Heading2') },
    { label: '수준 3', action: () => run('applyStyle', 'Heading3') },
  ]),
  insertEndnote: () => run('insertFootnote', 'end'),
  nextNoteMenu: (a) => openMenu(a, [{ label: '다음 각주', action: () => run('gotoNote', 1) }, { label: '이전 각주', action: () => run('gotoNote', -1) }]),
  citationMenu: (a) => openMenu(a, [
    ...(S.doc.sources ?? []).map((s) => ({ label: `${s.author} (${s.year}) ${s.title}`, action: () => run('insertCitation', s.id) })),
    { sep: true },
    { label: '새 원본 추가...', action: () => run('sourceDialog') },
  ]),
  citeStyleMenu: (a) => openMenu(a, ['APA', 'MLA', 'Chicago', 'IEEE', 'KS (한국 논문)'].map((st) => ({ label: st, checked: (S.doc.settings.citeStyle ?? 'APA') === st, action: () => change((d) => { d.doc.settings = { ...d.doc.settings, citeStyle: st }; return S.sel; }) }))),
  bibliographyMenu: (a) => openMenu(a, [{ label: '참고 문헌', action: () => run('insertBibliography', '참고 문헌') }, { label: 'References', action: () => run('insertBibliography', 'References') }, { label: '인용 문헌', action: () => run('insertBibliography', '인용 문헌') }]),
});

// ───────────── 편지 · 검토 · 표 메뉴 ─────────────
register({
  mergeStartMenu: (a) => openMenu(a, [
    { label: '편지', action: () => { S.merge = { ...(S.merge ?? {}), kind: 'letter' }; toast('편지 병합: [받는 사람 선택]으로 CSV 를 불러오세요'); } },
    { label: '레이블...', action: () => run('labelsDialog') },
    { label: '디렉터리', action: () => { S.merge = { ...(S.merge ?? {}), kind: 'directory' }; } },
    { label: '일반 Word 문서', action: () => { S.merge = null; emit('change', { scope: 'view' }); } },
  ]),
  recipientsMenu: (a) => openMenu(a, [{ label: '새 목록 입력...', action: () => run('editRecipients', true) }, { label: '기존 목록 사용 (CSV · Excel CSV)...', action: () => run('loadRecipients') }]),
  mergeFieldMenu: (a) => {
    const cols = S.merge?.cols ?? [];
    if (!cols.length) { toast('먼저 [받는 사람 선택]으로 목록을 불러오세요'); return; }
    openMenu(a, cols.map((c) => ({ label: c, action: () => run('insertMergeField', c) })));
  },
  mergeFinishMenu: (a) => openMenu(a, [{ label: '개별 문서 편집...', action: () => run('finishMerge', 'doc') }, { label: '문서 인쇄...', action: () => run('finishMerge', 'print') }]),
  markupMenu: (a) => openMenu(a, [
    ...[['simple', '간단한 태그'], ['all', '모든 태그'], ['none', '태그 없음'], ['original', '원본']].map(([k, l]) => ({ label: l, checked: S.markup === k, action: () => run('showMarkup', k) })),
  ]),
  acceptMenu: (a) => openMenu(a, [{ label: '적용 후 다음으로 이동', action: () => { run('acceptSel'); run('nextChange'); } }, { label: '이 변경 내용 적용', action: () => run('acceptSel') }, { label: '모든 변경 내용 적용', action: () => run('acceptAll') }, { label: '모든 변경 내용 적용 및 추적 중지', action: () => { run('acceptAll'); if (S.doc.settings.track) run('trackChanges'); } }]),
  rejectMenu: (a) => openMenu(a, [{ label: '거부 후 다음으로 이동', action: () => { run('rejectSel'); run('nextChange'); } }, { label: '이 변경 내용 거부', action: () => run('rejectSel') }, { label: '모든 변경 내용 거부', action: () => run('rejectAll') }, { label: '모든 변경 내용 거부 및 추적 중지', action: () => { run('rejectAll'); if (S.doc.settings.track) run('trackChanges'); } }]),
  prevChange: () => run('nextChange', -1),
  deleteCommentMenu: (a) => openMenu(a, [{ label: '삭제', action: () => run('deleteCurrentComment') }, { label: '문서의 모든 메모 삭제', action: () => run('deleteAllComments') }]),
  translateMenu: (a) => openMenu(a, [{ label: '선택 영역 번역', action: () => run('translateDialog', 'sel') }, { label: '문서 번역', action: () => run('translateDialog', 'doc') }]),
  cellShadingMenu: (a) => openMenu(a, [{ node: colorGrid((c) => { if (c !== 'none') S.lastCellShade = c; run('cellShading', c); }, { auto: null, none: '색 없음', recentKey: 'cell' }) }]),
  cellBordersMenu: (a) => openMenu(a, [...cellBorderItems(), { sep: true }, { label: '펜 두께', submenu: [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3].map((w) => ({ label: `${w}pt`, action: () => { S.borderPen = { ...(S.borderPen ?? { style: 'single', color: '#000000' }), w }; } })) }, { label: '펜 스타일', submenu: [['single', '실선'], ['double', '이중선'], ['dotted', '점선'], ['dashed', '파선']].map(([k, l]) => ({ label: l, action: () => { S.borderPen = { ...(S.borderPen ?? { w: 0.5, color: '#000000' }), style: k }; } })) }, { label: '펜 색', submenu: [{ node: colorGrid((c) => { S.borderPen = { ...(S.borderPen ?? { w: 0.5, style: 'single' }), color: c === 'auto' ? '#000000' : c }; }) }] }]),
  tableSelectMenu: (a) => openMenu(a, [{ label: '셀 선택', action: () => run('selectCell') }, { label: '열 선택', action: () => run('selectCol') }, { label: '행 선택', action: () => run('selectRow') }, { label: '표 선택', action: () => run('selectTable') }]),
  tableDeleteMenu: (a) => openMenu(a, [{ label: '셀 삭제...', action: () => run('deleteRow') }, { label: '열 삭제', icon: 'delete', action: () => run('deleteCol') }, { label: '행 삭제', icon: 'delete', action: () => run('deleteRow') }, { label: '표 삭제', icon: 'delete', action: () => run('deleteTable') }]),
  autofitMenu: (a) => openMenu(a, [{ label: '내용에 자동으로 맞춤', action: () => run('autoFitContents') }, { label: '창에 자동으로 맞춤', action: () => run('autoFitWindow') }, { label: '고정 열 너비', action: () => toast('열 너비가 고정되어 있습니다') }]),
  picCorrectionsMenu: (a) => openMenu(a, [
    { title: '밝기/대비' },
    ...[[-40, -40], [-20, -20], [0, 0], [20, 20], [40, 40], [20, 0], [0, 20]].map(([b, c]) => ({ label: `밝기 ${b > 0 ? '+' : ''}${b}%  대비 ${c > 0 ? '+' : ''}${c}%`, action: () => run('imageProps', { bright: b || undefined, contrast: c || undefined }) })),
  ]),
  picColorMenu: (a) => openMenu(a, [{ label: '회색조', action: () => run('imageProps', { gray: true }) }, { label: '원래 색', action: () => run('imageProps', { gray: undefined, bright: undefined, contrast: undefined }) }]),
  picBorderMenu: (a) => openMenu(a, [{ node: colorGrid((c) => run('imageProps', { border: c === 'none' ? undefined : { style: 'single', w: 1, color: c === 'auto' ? '#000000' : c } }), { none: '윤곽선 없음' }) }]),
  picEffectsMenu: (a) => openMenu(a, [{ label: '그림자', action: () => run('imageProps', { shadow: true }) }, { label: '효과 없음', action: () => run('imageProps', { shadow: undefined, rot: undefined }) }]),
  docInfoMenu: (a) => openMenu(a, [{ label: '작성자', action: () => run('insertField', 'AUTHOR') }, { label: '문서 제목', action: () => run('insertField', 'TITLE') }, { label: '날짜', action: () => run('insertField', 'DATE') }]),
});

// ───────────── 갤러리 ─────────────
export const GALLERIES = {
  styles: () => {
    const wrap = el('div', { class: 'style-gal-wrap' });
    const gal = el('div', { class: 'style-gal' });
    const cur = (() => { try { return blockAt(S.doc, curSel().f.p)?.style; } catch { return null; } })();
    const list = Object.values(S.doc.styles).filter((s) => s.quick && !s.hidden).sort((x, y) => (x.order ?? 99) - (y.order ?? 99));
    for (const st of list) gal.append(styleCell(st, cur));
    const more = el('div', { class: 'style-gal-more' },
      el('button', { title: '위로', onclick: () => { gal.scrollTop -= 35; } }, '▲'),
      el('button', { title: '아래로', onclick: () => { gal.scrollTop += 35; } }, '▼'),
      el('button', { class: 'gal-more', title: '자세히', onclick: (e) => run('stylesGalleryMenu', e.currentTarget) }, '▾'));
    wrap.append(gal, more);
    return wrap;
  },
  styleSets: () => {
    const g = el('div', { class: 'style-gal', style: { height: '70px', maxWidth: '420px' } });
    for (const [k, l] of STYLE_SET_LIST) g.append(el('button', { class: 'style-cell', style: { width: '76px', height: '68px' }, title: l, onclick: () => run('styleSet', k) }, el('span', { class: 'sp1', style: { fontWeight: 700, fontSize: '11px' } }, '제목'), el('span', { class: 'sn' }, l)));
    return g;
  },
  tableStyles: () => {
    const g = el('div', { class: 'style-gal', style: { maxWidth: '380px' } });
    for (const [id, l] of tableStyleList().slice(0, 30)) g.append(el('button', { class: 'style-cell tsc', title: l, style: { width: '62px', height: '46px', padding: '2px' }, html: tableThumb(id), onclick: () => run('tableStyle', id) }));
    const more = el('div', { class: 'style-gal-more' }, el('button', { class: 'gal-more', title: '자세히', onclick: (e) => run('tableStylesMenu', e.currentTarget) }, '▾'));
    return el('div', { class: 'style-gal-wrap' }, g, more);
  },
  paraSpin: () => {
    const p = (() => { try { return blockAt(S.doc, curSel().f.p); } catch { return null; } })();
    const ep = p?.type === 'p' ? effPara(S.doc, p) : {};
    const box = (label, val, apply, unit = 'cm') => {
      const inp = el('input', { type: 'number', step: unit === 'cm' ? 0.1 : 1, value: unit === 'cm' ? f1(pxToCm(val ?? 0)) : Math.round(((val ?? 0) * 3) / 4), style: { width: '56px' } });
      inp.addEventListener('change', () => apply(unit === 'cm' ? cm(Number(inp.value)) : (Number(inp.value) * 4) / 3));
      inp.addEventListener('keydown', (e) => e.stopPropagation());
      return el('label', { class: 'spin-row' }, el('span', {}, label), inp, el('small', {}, unit === 'cm' ? 'cm' : 'pt'));
    };
    return el('div', { class: 'para-spin' },
      el('div', { class: 'ps-col' }, el('b', {}, '들여쓰기'), box('왼쪽:', ep.ind?.l, (v) => run('setIndent', { l: v })), box('오른쪽:', ep.ind?.r, (v) => run('setIndent', { r: v }))),
      el('div', { class: 'ps-col' }, el('b', {}, '간격'), box('앞:', ep.spc?.before, (v) => run('spaceBefore', v), 'pt'), box('뒤:', ep.spc?.after, (v) => run('spaceAfter', v), 'pt')));
  },
  cellSize: () => {
    const p = curSel().f.p;
    if (p.length < 4) return el('span');
    const t = blockAt(S.doc, p.slice(0, -3));
    const r = p[p.length - 3]; const c = p[p.length - 2];
    let gc = 0; for (let k = 0; k < c; k++) gc += t.rows[r].cells[k].span ?? 1;
    const h = el('input', { type: 'number', step: 0.1, value: f1(pxToCm(t.rows[r].h ?? 0)), style: { width: '56px' } });
    const w = el('input', { type: 'number', step: 0.1, value: f1(pxToCm(t.grid[gc] ?? 0)), style: { width: '56px' } });
    h.addEventListener('change', () => run('setRowHeight', cm(Number(h.value))));
    w.addEventListener('change', () => run('setColWidth', gc, cm(Number(w.value))));
    for (const i of [h, w]) i.addEventListener('keydown', (e) => e.stopPropagation());
    return el('div', { class: 'ps-col' }, el('label', { class: 'spin-row' }, el('span', { html: ICONS.rowHeight }), el('span', {}, '높이:'), h, el('small', {}, 'cm')), el('label', { class: 'spin-row' }, el('span', { html: ICONS.colWidth }), el('span', {}, '너비:'), w, el('small', {}, 'cm')));
  },
  picSize: () => {
    const o = S.objSel;
    const r = o ? (() => { const p = blockAt(S.doc, o.p); let pos = 0; for (const x of p.runs) { if (pos === o.o) return x; pos += x.obj ? 1 : x.text.length; } return null; })() : null;
    const h = el('input', { type: 'number', step: 0.1, value: r ? f1(pxToCm(r.h)) : '', style: { width: '56px' } });
    const w = el('input', { type: 'number', step: 0.1, value: r ? f1(pxToCm(r.w)) : '', style: { width: '56px' } });
    const ratio = r ? r.w / r.h : 1;
    h.addEventListener('change', () => run('imageProps', { h: cm(Number(h.value)), w: cm(Number(h.value)) * ratio }));
    w.addEventListener('change', () => run('imageProps', { w: cm(Number(w.value)), h: cm(Number(w.value)) / ratio }));
    for (const i of [h, w]) i.addEventListener('keydown', (e) => e.stopPropagation());
    return el('div', { class: 'ps-col' }, el('label', { class: 'spin-row' }, el('span', {}, '높이:'), h, el('small', {}, 'cm')), el('label', { class: 'spin-row' }, el('span', {}, '너비:'), w, el('small', {}, 'cm')));
  },
  pictureStyles: () => {
    const g = el('div', { class: 'style-gal', style: { maxWidth: '300px', height: '64px' } });
    const P = [['단순 테두리', { border: { style: 'single', w: 1, color: '#000000' }, shadow: undefined }], ['흰 테두리 그림자', { border: { style: 'single', w: 4, color: '#FFFFFF' }, shadow: true }], ['파랑 테두리', { border: { style: 'single', w: 2, color: '#2B579A' } }], ['그림자', { shadow: true, border: undefined }], ['없음', { shadow: undefined, border: undefined }]];
    for (const [l, props] of P) g.append(el('button', { class: 'style-cell', style: { width: '56px', height: '58px' }, title: l, onclick: () => run('imageProps', props) }, el('span', { class: 'sn' }, l)));
    return g;
  },
  hfPos: () => {
    const sec = sectionAtTop(S.doc, 0);
    const mk = (label, key) => { const inp = el('input', { type: 'number', step: 0.1, value: f1(pxToCm(sec.margin[key] ?? 48)), style: { width: '56px' } }); inp.addEventListener('change', () => run('pageSetup', { margin: { [key]: cm(Number(inp.value)) } })); inp.addEventListener('keydown', (e) => e.stopPropagation()); return el('label', { class: 'spin-row' }, el('span', {}, label), inp, el('small', {}, 'cm')); };
    return el('div', { class: 'ps-col' }, mk('위쪽 머리글:', 'header'), mk('아래쪽 바닥글:', 'footer'));
  },
  mergeRecord: () => {
    const n = S.merge?.rows?.length ?? 0;
    const inp = el('input', { type: 'number', min: 1, max: Math.max(1, n), value: (S.merge?.index ?? 0) + 1, style: { width: '44px' } });
    inp.addEventListener('change', () => run('mergeGoto', Number(inp.value) - 1));
    inp.addEventListener('keydown', (e) => e.stopPropagation());
    return inp;
  },
};
function styleCell(st, cur) {
  const rp = { ...styleChain(S.doc.styles, st.id).rPr };
  const sz = Math.min(16, Math.max(9, rp.sz ?? 10));
  const css = runCss(S.doc, { ...S.doc.defaults.rPr, ...rp, sz });
  return el('button', { class: `style-cell${st.id === cur ? ' on' : ''}`, title: st.name, onclick: () => run('applyStyle', st.id), oncontextmenu: (e) => { e.preventDefault(); run('styleContextMenu', st.id, e.clientX, e.clientY); } },
    el('span', { class: 'sp1', style: Object.fromEntries(css.split(';').filter(Boolean).map((x) => { const i = x.indexOf(':'); return [x.slice(0, i).replace(/-([a-z])/g, (m, c) => c.toUpperCase()), x.slice(i + 1)]; })) }, 'AaBb가나'),
    el('span', { class: 'sn' }, st.name));
}
function tableThumb(id) {
  const t = { type: 'tbl', id: 'th', style: id, tblPr: {}, grid: [10, 10, 10, 10, 10], rows: Array.from({ length: 4 }, () => ({ cells: Array.from({ length: 5 }, () => ({ blocks: [] })) })) };
  return tableHtml(S.doc, t, {}).replace('<table class="tbl"', '<table class="tbl" style="width:52px;height:38px;table-layout:fixed;border-collapse:collapse"').replace(/<td /g, '<td height="8" ');
}
register({
  stylesGalleryMenu: (a) => {
    const cur = blockAt(S.doc, curSel().f.p)?.style;
    const g = el('div', { class: 'styles-menu' });
    for (const st of Object.values(S.doc.styles).filter((s) => (s.quick || s.type === 'para') && !s.hidden && s.type !== 'table').sort((x, y) => (x.order ?? 99) - (y.order ?? 99)).slice(0, 40)) g.append(styleCell(st, cur));
    openMenu(a, [{ node: g }, { sep: true }, { label: '스타일 만들기', action: () => run('newStyleDialog') }, { label: '서식 지우기', action: () => run('clearFormat') }, { label: '스타일 창 (Alt+Ctrl+Shift+S)', action: () => run('stylesPane') }]);
  },
  tableStylesMenu: (a) => {
    const g = el('div', { class: 'tstyle-gal' });
    for (const [id, l] of tableStyleList()) g.append(el('button', { title: l, html: tableThumb(id), onclick: () => { closeMenus(); run('tableStyle', id); } }));
    openMenu(a, [{ node: g }, { sep: true }, { label: '표 서식 지우기', action: () => run('tableStyle', 'TableNormal') }]);
  },
  styleContextMenu: (id, x, y) => openMenu({ x, y }, [
    { label: '선택 영역과 일치하도록 업데이트', action: () => run('updateStyleFromSel', id) },
    { label: '수정...', action: () => run('modifyStyleDialog', id) },
    { label: '스타일 갤러리에서 제거', action: () => run('modifyStyle', id, { quick: false }) },
    { label: '삭제', disabled: ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Title'].includes(id), action: () => run('deleteStyle', id) },
  ]),
});
void numberList; void bulletList; void formatNumber; void newDoc; void clone; void walkParas; void sections; void uid; void ACCENTS; void openSubmenu; void hydrateIcons; void TEMPLATES; void orderPos; void posEq; void selParas;
