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
  const sync = () => {
    const visible = new Set(state.visible());
    let count = 0;
    for (const [value, cb] of checks) {
      cb.parentElement.style.display = visible.has(value) ? '' : 'none';
      cb.checked = state.checked(value);
      if (visible.has(value) && cb.checked) count++;
    }
    all.checked = visible.size > 0 && count === visible.size;
    all.indeterminate = count > 0 && count < visible.size;
    all.disabled = visible.size === 0;
    allText.textContent = state.searching() ? '(검색 결과 모두 선택)' : '(모두 선택)';
    addRow.style.display = state.searching() ? '' : 'none';
    empty.style.display = visible.size ? 'none' : '';
  };
  list.append(el('label', {}, all, allText));
  for (const value of items) {
    const cb = el('input', { type: 'checkbox', 'aria-label': String(label(value)) });
    cb.addEventListener('change', () => { state.toggle(value, cb.checked); sync(); });
    checks.set(value, cb); list.append(el('label', {}, cb, label(value)));
  }
  list.append(empty);
  all.addEventListener('change', () => { state.selectVisible(all.checked); sync(); });
  search.addEventListener('input', () => { state.search(search.value); sync(); });
  const visibleChecks = () => [all, ...state.visible().map((value) => checks.get(value))].filter((cb) => !cb.disabled);
  search.addEventListener('keydown', (e) => { if (e.key === 'ArrowDown') { e.preventDefault(); visibleChecks()[0]?.focus(); } });
  list.addEventListener('keydown', (e) => {
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) return;
    const inputs = visibleChecks(), at = inputs.indexOf(e.target);
    if (at < 0) return;
    e.preventDefault(); e.stopPropagation();
    const next = e.key === 'Home' ? 0 : e.key === 'End' ? inputs.length - 1 : Math.max(0, Math.min(inputs.length - 1, at + (e.key === 'ArrowDown' ? 1 : -1)));
    inputs[next]?.focus();
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
