import { el } from './ui.js';

/** 앱 내부 시트 목록. iOS 네이티브 select/picker를 열지 않는다. */
export function createSheetPicker(items, { label = '시트 목록', multiple = false, selected = [], onChange, onAccept } = {}) {
  const selectedKeys = new Set(selected), buttons = [];
  let current = Math.max(0, items.findIndex(item => selectedKeys.has(item.value)));
  const body = el('div', { class: 'sheet-picker', role: 'listbox', 'aria-label': label,
    'aria-multiselectable': String(multiple), 'data-access-key': 'none' });
  const read = () => items.filter(item => selectedKeys.has(item.value)).map(item => item.value);
  const paint = () => {
    buttons.forEach((button, i) => {
      const checked = selectedKeys.has(items[i].value);
      button.setAttribute('aria-selected', String(checked));
      button.tabIndex = i === current ? 0 : -1;
      button.querySelector('.sheet-picker-mark').textContent = checked ? (multiple ? '✓' : '●') : '';
    });
  };
  const choose = index => {
    current = index;
    const key = items[index].value;
    if (multiple) { if (selectedKeys.has(key)) selectedKeys.delete(key); else selectedKeys.add(key); }
    else { selectedKeys.clear(); selectedKeys.add(key); }
    paint(); onChange?.(read());
  };
  items.forEach((item, i) => {
    const button = el('button', { type: 'button', class: 'sheet-picker-option', role: 'option',
      'data-access-key': 'none', 'data-sheet-value': String(item.value),
      onclick: () => { choose(i); button.focus({ preventScroll: true }); },
      onfocus: () => { current = i; paint(); },
    }, el('span', { class: 'sheet-picker-mark', 'aria-hidden': 'true' }), el('span', { class: 'sheet-picker-name' }, item.label));
    buttons.push(button); body.append(button);
  });
  if (!items.length) body.append(el('div', { class: 'sheet-picker-empty', role: 'status' }, '선택할 시트가 없습니다.'));
  body.addEventListener('keydown', event => {
    if (event.isComposing || event.keyCode === 229 || !buttons.length) return;
    let next = current;
    if (event.key === 'ArrowDown') next = Math.min(items.length - 1, current + 1);
    else if (event.key === 'ArrowUp') next = Math.max(0, current - 1);
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = items.length - 1;
    else if (event.key === ' ') { event.preventDefault(); event.stopPropagation(); choose(current); return; }
    else if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); onAccept?.(); return; }
    else return;
    event.preventDefault(); event.stopPropagation();
    current = next;
    if (!multiple) { selectedKeys.clear(); selectedKeys.add(items[current].value); }
    paint(); buttons[current].focus({ preventScroll: true }); buttons[current].scrollIntoView({ block: 'nearest' });
    if (!multiple) onChange?.(read());
  });
  paint();
  return { body, read, focus: () => buttons[current]?.focus({ preventScroll: true }) };
}
