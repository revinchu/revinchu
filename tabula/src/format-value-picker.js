import { el, openMenu, openDialog } from './ui.js';

/** 직접 입력은 데스크톱에 유지하고 터치 화면에서는 전체 목록 버튼을 사용한다. */
export function createFormatValuePicker({ label, className = '', inputClass = '', buttonClass = '', values = [], value = '', number = false, disabled = () => false, onChange, onOpen, onEscape, afterCommit } = {}) {
  let current = String(value ?? '');
  const locked = () => typeof disabled === 'function' ? !!disabled() : !!disabled;
  const input = el('input', { class: `format-picker-input ${inputClass}`, type: 'text', 'aria-label': label, title: label, autocomplete: 'off', spellcheck: false, inputmode: number ? 'decimal' : undefined });
  const caption = el('span', { class: 'format-picker-value', 'aria-hidden': 'true' }, current);
  const button = el('button', { type: 'button', class: `format-picker-toggle ${buttonClass}`, title: label + ' 목록', 'aria-label': label + ' 목록', 'aria-haspopup': 'menu', 'data-access-key': 'none' }, caption, el('span', { class: 'format-picker-arrow', 'aria-hidden': 'true' }, '▾'));
  const root = el('span', { class: 'format-value-picker ' + className }, input, button);
  const parse = (raw) => {
    const text = String(raw).trim();
    if (!text) return { error: label + '을(를) 입력하세요.' };
    if (!number) return { value: text };
    const n = Number(text);
    return Number.isFinite(n) && n > 0 && n <= 409 ? { value: n } : { error: '글꼴 크기는 0보다 크고 409 이하인 숫자여야 합니다.' };
  };
  const update = (next = current) => {
    current = String(next ?? '');
    if (document.activeElement !== input) input.value = current;
    caption.textContent = current;
    button.title = label + ' 목록: ' + current;
    input.disabled = button.disabled = locked();
  };
  const commit = (raw, field = input) => {
    if (locked()) return false;
    const parsed = parse(raw);
    field.setCustomValidity(parsed.error || '');
    if (parsed.error) { field.reportValidity(); return false; }
    if (String(parsed.value) !== current) {
      current = String(parsed.value); caption.textContent = current;
      onChange?.(parsed.value);
    }
    afterCommit?.();
    return true;
  };
  const direct = () => {
    if (locked()) return;
    const edit = el('input', { type: 'text', 'aria-label': label, value: current, inputmode: number ? 'decimal' : undefined, autocomplete: 'off', style: { fontSize: '16px', width: '100%', boxSizing: 'border-box' } });
    edit.addEventListener('input', () => edit.setCustomValidity(''));
    openDialog({ title: label + ' 직접 입력', body: el('label', { class: 'format-picker-direct' }, el('span', {}, label), edit), width: 380,
      buttons: [{ label: '확인', primary: true, action: () => commit(edit.value, edit) ? undefined : false }, { label: '취소' }] });
  };
  const open = () => {
    if (locked()) return;
    // Finish the input change and its focus callback before the menu captures its target.
    if (document.activeElement === input) input.blur();
    if (onOpen) { onOpen(root); return; }
    const list = typeof values === 'function' ? values() : values;
    openMenu(button, [
      ...list.map((v) => ({ label: String(v), checked: String(v) === current, accessKey: 'none', action: () => commit(v) })),
      { sep: true }, { label: '직접 입력…', accessKey: 'none', action: direct },
    ], { minWidth: number ? 120 : 220, scroll: true });
  };
  button.addEventListener('mousedown', (event) => event.preventDefault());
  button.addEventListener('click', open);
  input.addEventListener('change', () => commit(input.value));
  input.addEventListener('input', () => input.setCustomValidity(''));
  input.addEventListener('keydown', (event) => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); if (commit(input.value) && document.activeElement === input) input.select(); }
    else if (event.key === 'ArrowDown' && event.altKey) { event.preventDefault(); event.stopPropagation(); open(); }
    else if (event.key === 'Escape' && onEscape) { event.preventDefault(); input.value = current; onEscape(); }
  });
  input.value = current; update();
  return { root, input, button, update, open };
}
