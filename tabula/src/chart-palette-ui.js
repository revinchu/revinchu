import { el } from './ui.js';
import { CHART_PALETTE_GROUPS, chartPaletteOptions, normalizeChartPalette } from './chart-palette-options.js';

/** 색 견본과 이름을 함께 표시하며, 검색은 적용 전 목록만 바꾼다. */
export function createChartPalettePicker({ chart, palettes, onPick, onCustom }) {
  const search = el('input', { type: 'search', placeholder: '색 구성 검색', 'aria-label': '색 구성 검색', 'data-menu-search-target': '.chart-palette-results', autocomplete: 'off' });
  const group = el('select', { 'aria-label': '색 구성 분류' }, el('option', { value: 'all' }, '모든 색'),
    CHART_PALETTE_GROUPS.map(g => el('option', { value: g.id }, g.label)));
  const list = el('div', { class: 'chart-palette-results' }), count = el('span', { class: 'chart-palette-count', role: 'status' });
  const draw = () => {
    const items = chartPaletteOptions(palettes, search.value, group.value);
    count.textContent = items.length + '개 색 구성';
    list.replaceChildren(...items.map(p => el('button', { type: 'button', class: 'pal-row chart-palette-option' + ((chart.palette ?? 'office') === p.id ? ' on' : ''),
      title: p.label, 'aria-label': p.label, 'aria-pressed': String((chart.palette ?? 'office') === p.id),
      onclick: () => onPick(p.id) }, el('span', { class: 'chart-palette-name' }, p.label),
      el('span', { class: 'chart-palette-swatches', 'aria-hidden': 'true' }, p.colors.map(c => el('i', { style: { background: c } }))))));
    if (!items.length) list.append(el('p', { class: 'chart-palette-empty' }, '일치하는 색 구성이 없습니다.'));
  };
  search.addEventListener('input', draw); group.addEventListener('change', draw);
  const root = el('div', { class: 'pal-list chart-palette-picker' }, el('div', { class: 'chart-palette-search' }, search, group), count, list,
    el('button', { type: 'button', class: 'btn chart-palette-custom', onclick: onCustom }, '사용자 지정 색 구성…'));
  root.addEventListener('keydown', e => {
    if (e.target.matches('input,select') && !['Escape', 'Tab'].includes(e.key)
      && !(e.target === search && ['ArrowDown', 'ArrowUp'].includes(e.key))) e.stopPropagation();
  });
  draw(); return root;
}

/** 편집 중에는 문서에 쓰지 않는다. 유효한 최종 색만 getColors()로 반환한다. */
export function createChartPaletteEditor({ colors, renderPreview, maxColors = 32 }) {
  let draft = [...colors], bulkDirty = false;
  const rows = el('div', { class: 'chart-palette-edit-rows' }), preview = el('div', { class: 'chart-palette-preview', 'aria-label': '색 구성 미리보기' });
  const error = el('p', { class: 'warn chart-palette-error', role: 'alert', hidden: true }), count = el('span', { role: 'status' });
  const bulk = el('textarea', { rows: 2, 'aria-label': '색 코드 목록', placeholder: '#4472C4, #ED7D31, #70AD47', spellcheck: 'false', 'data-access-key': 'none' });
  const showError = text => { error.textContent = text || ''; error.hidden = !text; };
  bulk.addEventListener('input', () => { bulkDirty = true; });
  const normalized = () => normalizeChartPalette(draft, { maxColors });
  const sync = () => {
    count.textContent = draft.length + ' / ' + maxColors + '색';
    add.disabled = draft.length >= maxColors; reverse.disabled = draft.length < 2;
    if (!bulkDirty) bulk.value = draft.join(', ');
    const value = normalized(); showError(value.error);
    if (!value.error) preview.replaceChildren(el('div', { html: renderPreview(value.colors) }));
  };
  const move = (i, step) => { [draft[i], draft[i + step]] = [draft[i + step], draft[i]]; draw(); };
  const draw = () => {
    rows.replaceChildren(...draft.map((v, i) => {
      const text = el('input', { type: 'text', value: v, 'aria-label': (i + 1) + '번째 색 코드', spellcheck: 'false', maxlength: 9 });
      const one = normalizeChartPalette([v]), picker = el('input', { type: 'color', value: one.error ? '#4472c4' : one.colors[0], 'aria-label': (i + 1) + '번째 색 선택' });
      text.addEventListener('input', () => { bulkDirty = false; draft[i] = text.value; const n = normalizeChartPalette([text.value]); if (!n.error) { picker.value = n.colors[0]; picker.style.backgroundColor = n.colors[0]; } sync(); });
      picker.addEventListener('input', () => { bulkDirty = false; draft[i] = picker.value; text.value = picker.value.toUpperCase(); picker.style.backgroundColor = picker.value; sync(); });
      picker.style.backgroundColor = picker.value;
      return el('div', { class: 'chart-palette-edit-row' }, el('span', {}, i + 1), picker, text,
        el('button', { type: 'button', class: 'btn small', 'aria-label': (i + 1) + '번째 색 위로', title: '위로', disabled: i === 0, onclick: () => move(i, -1) }, '↑'),
        el('button', { type: 'button', class: 'btn small', 'aria-label': (i + 1) + '번째 색 아래로', title: '아래로', disabled: i === draft.length - 1, onclick: () => move(i, 1) }, '↓'),
        el('button', { type: 'button', class: 'btn small', 'aria-label': (i + 1) + '번째 색 삭제', title: '삭제', disabled: draft.length === 1, onclick: () => { draft.splice(i, 1); draw(); } }, '×'));
    }));
    rows.querySelectorAll('input,button').forEach(node => node.dataset.accessKey = 'none');
    bulkDirty = false; bulk.value = draft.join(', '); sync();
  };
  const add = el('button', { type: 'button', class: 'btn small', onclick: () => { if (draft.length < maxColors) { draft.push('#4472C4'); draw(); rows.scrollTop = rows.scrollHeight; } } }, '색 추가');
  const reverse = el('button', { type: 'button', class: 'btn small', onclick: () => { draft.reverse(); draw(); } }, '순서 뒤집기');
  const body = el('div', { class: 'chart-palette-editor' }, el('p', {}, '1~' + maxColors + '개의 색을 지정하세요. 색 순서대로 계열과 항목에 적용합니다.'),
    el('div', { class: 'chart-palette-editor-layout' }, el('div', {}, rows, el('div', { class: 'chart-palette-editor-tools' }, add, reverse, count)), preview),
    el('label', { class: 'chart-palette-bulk-label' }, '여러 색 한 번에 입력 (HEX)', bulk),
    el('button', { type: 'button', class: 'btn small', onclick: () => { const n = normalizeChartPalette(bulk.value, { maxColors }); if (n.error) { showError(n.error); return; } draft = n.colors; draw(); } }, '목록 불러오기'), error);
  draw();
  return { body, getColors: () => { const value = bulkDirty ? normalizeChartPalette(bulk.value, { maxColors }) : normalized(); showError(value.error); return value.error ? null : value.colors; } };
}
