import { el, openDialog } from './ui.js';
import { PASTE_SPECIAL_TYPES, PASTE_SPECIAL_OPERATIONS, pasteSourceFromText } from './paste-special.js';

/** 외부 텍스트도 같은 선택 창에서 사용. 권한 거부 시 사용자가 Ctrl+V로 직접 제공한다. */
export function showPasteSpecial({ source, readText, date1904 = false, onApply }) {
  let current = source, alive = true, loading = false, manual = false;
  const status = el('div', { class: 'muted', role: 'status' });
  const input = el('textarea', { rows: 3, 'aria-label': '외부 클립보드 데이터', placeholder: '클립보드 읽기가 허용되지 않으면 여기를 클릭하고 Ctrl+V로 붙여 넣으세요.', 'data-access-key': 'p', style: { width: '100%', boxSizing: 'border-box' } });
  const fallback = el('div', {}, input, el('div', { class: 'muted' }, '외부 텍스트에는 셀 서식·메모·유효성 검사·조건부 서식·열 너비 정보가 없습니다.'));
  const radio = (name, rows, selected) => el('div', { class: 'ps-grid' }, rows.map(([value, label, key]) => el('label', { class: 'fc-check' }, el('input', { type: 'radio', name, value, checked: value === selected, 'aria-label': label, 'data-access-key': key }), `${label}(${key.toUpperCase()})`)));
  const types = radio('psWhat', PASTE_SPECIAL_TYPES, 'all'), operations = radio('psOp', PASTE_SPECIAL_OPERATIONS, '');
  types.style.gridAutoFlow = 'column'; types.style.gridTemplateRows = 'repeat(6, auto)';
  operations.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
  const skip = el('input', { type: 'checkbox', 'aria-label': '내용 있는 셀만 붙여넣기', 'data-access-key': 'b' });
  const transpose = el('input', { type: 'checkbox', 'aria-label': '행/열 바꿈', 'data-access-key': 'e' });
  const body = el('div', { class: 'paste-special-dialog', style: { display: 'flex', flexDirection: 'column', gap: '10px' } }, status, fallback,
    el('fieldset', {}, el('legend', {}, '붙여넣기'), types), el('fieldset', {}, el('legend', {}, '연산'), operations),
    el('div', { style: { display: 'flex', gap: '16px', flexWrap: 'wrap' } }, el('label', { class: 'fc-check' }, skip, '내용 있는 셀만 붙여넣기(B)'), el('label', { class: 'fc-check' }, transpose, '행/열 바꿈(E)')));
  const options = () => ({ what: types.querySelector(':checked').value, op: operations.querySelector(':checked').value || null, skipBlanks: skip.checked, transpose: transpose.checked });
  const apply = link => { if (loading || !current) { status.textContent = '먼저 데이터를 복사하거나 아래 입력란에 Ctrl+V로 붙여 넣으세요.'; input.focus(); return false; } return onApply(current, { ...options(), ...(link ? { what: 'link' } : {}) }); };
  fallback.hidden = !!current && !current.external;
  const dialog = openDialog({ title: '선택하여 붙여넣기', width: 520, body, onClose: () => { alive = false; }, buttons: [
    { label: '연결하여 붙여넣기(L)', accessKey: 'l', action: () => apply(true) }, { label: '확인', accessKey: 'none', primary: true, action: () => apply(false) }, { label: '취소', accessKey: 'none' },
  ] });
  const closeButton = dialog.root.querySelector('.dialog-head button');
  closeButton.dataset.accessKey = 'none'; closeButton.removeAttribute('data-dialog-close-head');
  const link = dialog.root.querySelector('.dialog-foot button');
  const refresh = () => {
    fallback.hidden = !!current && !current.external;
    const unavailable = current?.external ? new Set(['formats', 'comments', 'validation', 'sourceTheme', 'noBorders', 'colWidths', 'mergeCond']) : new Set();
    for (const node of types.querySelectorAll('input')) node.disabled = unavailable.has(node.value);
    if (types.querySelector(':checked').disabled) types.querySelector('[value="values"]').checked = true;
    const what = options().what, canOperate = ['all', 'sourceTheme', 'noBorders', 'mergeCond', 'formulas', 'values', 'formulasNum', 'valuesNum'].includes(what);
    for (const node of operations.querySelectorAll('input')) node.disabled = !canOperate && !!node.value;
    if (!canOperate) operations.querySelector('[value=""]').checked = true;
    transpose.disabled = what === 'colWidths'; if (transpose.disabled) transpose.checked = false;
    link.disabled = loading || !current || current.external || what !== 'all' || !!options().op || skip.checked || transpose.checked;
    status.textContent = loading ? '시스템 클립보드를 확인하는 중…' : current ? `${current.external ? '외부 텍스트' : '복사한 셀'}: ${current.data.length}행 × ${current.data[0].length}열${current.external ? ' · 서식 정보가 없는 옵션은 사용할 수 없습니다.' : ''}` : '클립보드를 읽을 수 없습니다. 아래 입력란에 Ctrl+V로 붙여 넣으세요.';
  };
  const setText = text => { manual = true; loading = false; if (!text) { current = null; refresh(); return; } try { current = pasteSourceFromText(text, date1904); refresh(); } catch (e) { current = null; refresh(); status.textContent = e.message; } };
  input.addEventListener('input', () => setText(input.value));
  input.addEventListener('paste', event => { const text = event.clipboardData?.getData('text/plain'); if (text !== undefined) { event.preventDefault(); input.value = text; setText(text); } });
  body.addEventListener('change', refresh);
  body.addEventListener('keydown', e => {
    if (e.altKey || e.ctrlKey || e.metaKey || e.isComposing || e.target === input) return;
    const key = /^Key[A-Z]$/.test(e.code) ? e.code.slice(3).toLowerCase() : e.key.toLowerCase();
    const node = body.querySelector(`[data-access-key="${/^[a-z]$/.test(key) ? key : '_'}"]`);
    if (node && !node.disabled && node !== input) { e.preventDefault(); e.stopPropagation(); node.focus(); node.click(); }
  });
  refresh();
  if (readText) {
    loading = true; refresh();
    Promise.race([Promise.resolve().then(readText), new Promise(resolve => setTimeout(() => resolve(null), 3000))]).then(text => {
      if (!alive || manual) return;
      if (text && (!source || text.replace(/\r\n?/g, '\n').replace(/\n$/, '') !== String(source.text ?? '').replace(/\r\n?/g, '\n').replace(/\n$/, ''))) { input.value = text; current = pasteSourceFromText(text, date1904); }
    }).catch(() => {}).finally(() => { if (alive) { loading = false; refresh(); if (!current) input.focus(); } });
  }
  return dialog;
}
