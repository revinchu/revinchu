import { el } from './ui.js';
import { normalizePhonetic, phoneticHtml } from './phonetic.js';
import { setSafeHtml } from './safe-html.js';

export function phoneticEditor({ text = '', phonetic } = {}) {
  const original = normalizePhonetic(phonetic, text) ?? { runs: [], visible: true, type: 'noConversion', alignment: 'left', font: { size: 8 } };
  const rows = [], runList = el('div', { class: 'phonetic-runs' });
  const base = el('textarea', { 'aria-label': '원문', rows: 2, value: text });
  base.value = text;
  const visible = el('input', { type: 'checkbox', checked: original.visible });
  const font = el('input', { 'aria-label': '윗주 글꼴', value: original.font?.font ?? '' });
  const size = el('input', { type: 'number', min: 1, max: 409, 'aria-label': '윗주 크기(pt)', value: original.font?.size ?? 8 });
  const alignment = el('select', { 'aria-label': '윗주 맞춤' }, [['left', '왼쪽'], ['center', '가운데'], ['distributed', '균등 분할'], ['noControl', '원본 설정']].map(([value, label]) => el('option', { value, selected: original.alignment === value }, label)));
  const preview = el('div', { class: 'phonetic-preview', 'aria-label': '윗주 미리보기' });
  const read = () => {
    if (!base.value || base.value.length > 32767) throw new Error('원문은 1~32,767자로 입력하세요.');
    if (!Number.isFinite(Number(size.value)) || Number(size.value) <= 0 || Number(size.value) > 409) throw new Error('윗주 크기는 1~409pt로 입력하세요.');
    const meta = normalizePhonetic({ ...original, runs: rows.map(row => ({ sb: Number(row.start.value) - 1, eb: Number(row.end.value), text: row.hint.value })), visible: visible.checked, alignment: alignment.value, font: { ...original.font, font: font.value, size: Number(size.value) } }, base.value, true);
    return { text: base.value, phonetic: meta };
  };
  const update = () => { try { const data = read(); setSafeHtml(preview, phoneticHtml(data.text, { ...data.phonetic, visible: true }) ?? ''); } catch { preview.textContent = '구간과 원문을 확인하세요.'; } };
  const add = (run = { sb: 0, eb: base.value.length, text: '' }) => {
    const start = el('input', { type: 'number', min: 1, 'aria-label': '윗주 시작 글자', value: run.sb + 1 });
    const end = el('input', { type: 'number', min: 1, 'aria-label': '윗주 끝 글자', value: run.eb });
    const hint = el('input', { 'aria-label': '윗주', value: run.text });
    const row = { start, end, hint };
    const node = el('div', { class: 'phonetic-run' }, el('label', {}, '시작', start), el('label', {}, '끝', end), el('label', {}, '윗주', hint), el('button', { type: 'button', onclick: () => { rows.splice(rows.indexOf(row), 1); node.remove(); update(); } }, '삭제'));
    rows.push(row); runList.append(node); node.addEventListener('input', update); return hint;
  };
  for (const run of original.runs.length ? original.runs : [{ sb: 0, eb: text.length, text: '' }]) add(run);
  const body = el('div', { class: 'phonetic-editor' }, el('label', {}, '원문', base),
    el('p', { class: 'muted' }, '시작과 끝은 1부터 세는 글자 위치입니다. 원문 위에 표시할 윗주를 직접 입력하세요.'), runList,
    el('button', { type: 'button', onclick: () => { add().focus(); update(); } }, '윗주 구간 추가'),
    el('div', { class: 'phonetic-options' }, el('label', {}, '윗주 글꼴', font), el('label', {}, '크기(pt)', size), el('label', {}, '맞춤', alignment)),
    el('label', { class: 'phonetic-visible' }, visible, '윗주 필드 표시'), preview);
  body.addEventListener('input', update); body.addEventListener('change', update); update();
  return { body, read };
}
