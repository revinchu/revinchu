import { el } from './ui.js';
import { filterSelection } from './filter-selection.js';

// 일반 필터와 피벗 필터가 같은 검색·선택·키보드 동작을 사용한다.
export function filterChecklist(items, initial = null, label = (value) => value || '(필드 값 없음)') {
  const state = filterSelection(items, initial, label);
  const search = el('input', { type: 'search', placeholder: '검색', 'aria-label': '필터 항목 검색' });
  const list = el('div', { class: 'filter-list' });
  const all = el('input', { type: 'checkbox', 'aria-label': '표시된 항목 모두 선택' });
  const allText = document.createTextNode('(모두 선택)');
  const checks = new Map();
  const empty = el('div', { class: 'muted', role: 'status', style: { padding: '8px', display: 'none' } }, '검색 결과가 없습니다.');
  const add = el('input', { type: 'checkbox' });
  const addRow = el('label', { class: 'filter-add', style: { display: 'none' } }, add, '필터에 현재 선택 내용 추가');
  // Keep every value in the selection model, but create DOM only for the viewport.
  const rowHeight = 26, overscan = 8;
  const header = el('label', { style: { height: `${rowHeight}px`, boxSizing: 'border-box', position: 'sticky', top: '0', zIndex: '1', background: 'var(--surface, #fff)' } }, all, allText);
  const rows = el('div', { class: 'filter-window', style: { position: 'relative' } });
  let drawnValues = null, drawnStart = -1, drawnEnd = -1;
  const draw = () => {
    const visible = state.visible();
    const start = Math.max(0, Math.floor((list.scrollTop - rowHeight) / rowHeight) - overscan);
    const end = Math.min(visible.length, start + Math.ceil((list.clientHeight || 180) / rowHeight) + overscan * 2);
    rows.style.height = `${visible.length * rowHeight}px`;
    if (visible !== drawnValues || start !== drawnStart || end !== drawnEnd) {
      drawnValues = visible; drawnStart = start; drawnEnd = end;
      checks.clear();
      const nodes = [];
      for (let i = start; i < end; i++) {
        const value = visible[i];
        const cb = el('input', { type: 'checkbox', 'aria-label': String(label(value)), 'data-filter-index': String(i) });
        cb.addEventListener('change', () => { state.toggle(value, cb.checked); sync(); });
        checks.set(value, cb);
        nodes.push(el('label', { title: String(label(value)), style: { position: 'absolute', top: `${i * rowHeight}px`, left: '0', right: '0', height: `${rowHeight}px`, boxSizing: 'border-box' } }, cb, label(value)));
      }
      rows.replaceChildren(...nodes);
    }
    for (const [value, cb] of checks) cb.checked = state.checked(value);
  };
  const sync = () => {
    const visible = state.visible();
    let count = 0;
    for (const value of visible) if (state.checked(value)) count++;
    all.checked = visible.length > 0 && count === visible.length;
    all.indeterminate = count > 0 && count < visible.length;
    all.disabled = visible.length === 0;
    allText.textContent = state.searching() ? '(검색 결과 모두 선택)' : '(모두 선택)';
    addRow.style.display = state.searching() ? '' : 'none';
    empty.style.display = visible.length ? 'none' : '';
    draw();
  };
  const focusIndex = (index) => {
    if (index < 0) { list.scrollTop = 0; draw(); if (!all.disabled) all.focus({ preventScroll: true }); return; }
    const visible = state.visible();
    if (!visible.length) return;
    index = Math.max(0, Math.min(visible.length - 1, index));
    const top = (index + 1) * rowHeight, bottom = top + rowHeight;
    if (top < list.scrollTop + rowHeight) list.scrollTop = top - rowHeight;
    else if (bottom > list.scrollTop + (list.clientHeight || 180)) list.scrollTop = bottom - (list.clientHeight || 180);
    draw(); checks.get(visible[index])?.focus({ preventScroll: true });
  };
  list.append(header, rows, empty);
  list.addEventListener('scroll', draw);
  all.addEventListener('change', () => { state.selectVisible(all.checked); sync(); });
  search.addEventListener('input', () => { state.search(search.value); list.scrollTop = 0; sync(); });
  search.addEventListener('keydown', (e) => { if (e.key === 'ArrowDown') { e.preventDefault(); focusIndex(-1); } });
  list.addEventListener('keydown', (e) => {
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
    const at = e.target === all ? -1 : Number(e.target.dataset.filterIndex);
    if (!Number.isInteger(at)) return;
    e.preventDefault(); e.stopPropagation();
    const next = e.key === 'Home' ? -1 : e.key === 'End' ? state.visible().length - 1 : Math.max(-1, at + (e.key === 'ArrowDown' ? 1 : -1));
    focusIndex(next);
  });
  sync();
  return { search, list, addRow, result: () => state.result(add.checked) };
}

export function searchableFieldPicker(fields, initial = []) {
  const selected = new Set(initial);
  const search = el('input', { type: 'search', placeholder: '필드 검색', 'aria-label': '슬라이서 필드 검색' });
  const list = el('div', { class: 'fc-list', style: { maxHeight: '260px', padding: '4px 8px', gap: '4px' } });
  const count = el('div', { class: 'muted', role: 'status' });
  const empty = el('div', { class: 'muted', style: { padding: '8px' } }, '검색 결과가 없습니다.');
  const rows = fields.map((name) => {
    const cb = el('input', { type: 'checkbox', checked: selected.has(name), 'aria-label': name });
    cb.addEventListener('change', () => { if (cb.checked) selected.add(name); else selected.delete(name); sync(); });
    return { name, cb, row: el('label', { class: 'fc-check' }, cb, name) };
  });
  const sync = () => {
    const q = search.value.trim().toLocaleLowerCase(); let visible = 0;
    for (const item of rows) {
      item.row.hidden = !!q && !item.name.toLocaleLowerCase().includes(q);
      item.row.style.display = item.row.hidden ? 'none' : '';
      item.cb.checked = selected.has(item.name);
      if (!item.row.hidden) visible++;
    }
    count.textContent = `${fields.length}개 필드 중 ${selected.size}개 선택 · 검색 결과 ${visible}개`;
    empty.hidden = visible > 0;
  };
  const choose = (on) => { for (const item of rows) if (!item.row.hidden) { if (on) selected.add(item.name); else selected.delete(item.name); } sync(); };
  search.addEventListener('input', sync);
  search.addEventListener('keydown', (e) => { if (e.key === 'ArrowDown') { e.preventDefault(); rows.find((item) => !item.row.hidden)?.cb.focus(); } });
  list.append(...rows.map((item) => item.row), empty); sync();
  const root = el('div', { class: 'slicer-field-picker', style: { display: 'grid', gap: '6px' } }, search,
    el('div', { style: { display: 'flex', gap: '6px' } },
      el('button', { type: 'button', class: 'btn', onclick: () => choose(true) }, '검색 결과 모두 선택'),
      el('button', { type: 'button', class: 'btn', onclick: () => choose(false) }, '검색 결과 선택 해제')), count, list);
  return { root, selected: () => fields.filter((name) => selected.has(name)) };
}
