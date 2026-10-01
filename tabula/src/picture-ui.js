import { el } from './ui.js';
import { resizePicture, setPictureCrop, pictureCropStyle, pictureTransform, resetPictureFormatting } from './picture.js';

// 모달 안의 초안만 변경합니다. 확인 전에는 통합 문서와 실행 취소 기록을 건드리지 않습니다.
export function pictureEditor(original) {
  const draft = { ...original, crop: original.crop ? { ...original.crop } : undefined, lockAspect: original.lockAspect !== false };
  const fields = {}, numbers = [], natural = { w: 0, h: 0 };
  const previewImage = el('img', { src: original.src, alt: original.alt ?? original.name ?? '', draggable: false });
  const frame = el('div', { class: 'picture-preview-frame' }, previewImage);
  const caption = el('div', { class: 'picture-preview-caption', 'aria-live': 'polite' });
  const preview = el('div', { class: 'picture-preview', 'aria-label': '그림 미리보기' }, el('div', { class: 'picture-preview-stage' }, frame), caption);
  const refresh = () => {
    const angle = (Number(draft.rot) || 0) * Math.PI / 180;
    const bw = Math.abs(draft.w * Math.cos(angle)) + Math.abs(draft.h * Math.sin(angle));
    const bh = Math.abs(draft.w * Math.sin(angle)) + Math.abs(draft.h * Math.cos(angle));
    const k = Math.min(1, 260 / Math.max(4, bw), 210 / Math.max(4, bh));
    Object.assign(frame.style, { width: `${draft.w * k}px`, height: `${draft.h * k}px`, transform: pictureTransform(draft), border: draft.border ? `${(draft.borderW ?? 2) * k}px solid ${draft.border}` : '', borderRadius: draft.radius ? `${draft.radius * k}px` : '', opacity: draft.opacity ?? 1, boxShadow: draft.shadow ? '3px 3px 8px #0006' : '' });
    Object.assign(previewImage.style, pictureCropStyle(draft.crop)); previewImage.alt = draft.alt ?? draft.name ?? '';
    caption.textContent = `${draft.w} × ${draft.h}px · 회전 ${draft.rot ?? 0}°`;
  };
  const sync = () => {
    for (const [key, input] of Object.entries(fields)) {
      if (key.startsWith('crop.')) input.value = Math.round((draft.crop?.[key.slice(5)] ?? 0) * 100000) / 1000;
      else if (input.type === 'checkbox') input.checked = !!draft[key];
      else input.value = draft[key] ?? (key === 'placement' ? 'twoCell' : key === 'rot' ? 0 : '');
    }
    refresh();
  };
  const row = (label, input) => el('label', { class: 'picture-field' }, el('span', {}, label), input);
  const num = (key, label, value, min, max, fn) => {
    const input = fields[key] = el('input', { type: 'number', 'aria-label': label, value, min, max, step: 'any' }); numbers.push(input);
    input.addEventListener('input', () => {
      if (!input.checkValidity() || input.value === '') return;
      const n = Number(input.value); if (fn) fn(n); else draft[key] = n;
      refresh();
    });
    return row(label, input);
  };
  const check = (key, label) => {
    const input = fields[key] = el('input', { type: 'checkbox', 'aria-label': label, checked: !!draft[key] });
    input.addEventListener('change', () => { draft[key] = input.checked; refresh(); }); return row(label, input);
  };
  const text = (key, label, multiline = false) => {
    const input = fields[key] = el(multiline ? 'textarea' : 'input', { 'aria-label': label, ...(multiline ? { rows: 3 } : { type: 'text' }), value: draft[key] ?? '' });
    if (multiline) input.value = draft[key] ?? '';
    input.addEventListener('input', () => { draft[key] = input.value; refresh(); }); return row(label, input);
  };
  const section = (name, ...children) => el('details', { class: 'picture-section', open: true }, el('summary', {}, name), ...children);
  const setSize = (axis, n) => { Object.assign(draft, resizePicture(draft, axis, n)); fields.w.value = draft.w; fields.h.value = draft.h; };
  const originalSize = el('button', { type: 'button', class: 'btn', disabled: true, onclick: () => {
    Object.assign(draft, resizePicture({ w: natural.w, h: natural.h }, 'w', natural.w, true)); sync();
  } }, '원본 크기');
  const source = new Image();
  source.onload = () => { natural.w = source.naturalWidth; natural.h = source.naturalHeight; originalSize.disabled = !(natural.w > 0 && natural.h > 0); originalSize.title = `${natural.w} × ${natural.h}px`; };
  source.onerror = () => { originalSize.title = '원본 이미지의 크기를 읽지 못했습니다.'; };
  if (!original.linked) source.src = original.src;
  const placement = fields.placement = el('select', { 'aria-label': '셀 크기가 바뀔 때' }, [['twoCell', '위치와 크기 변함'], ['oneCell', '위치만 변함'], ['absolute', '변하지 않음 (위치 고정)']].map(([value, name]) => el('option', { value }, name)));
  placement.value = draft.placement ?? 'twoCell'; placement.addEventListener('change', () => { draft.placement = placement.value; });
  const size = section('크기 및 속성', text('name', '이름'), check('lockAspect', '가로 세로 비율 고정'),
    num('w', '너비(px)', draft.w, 4, 20000, n => setSize('w', n)), num('h', '높이(px)', draft.h, 4, 20000, n => setSize('h', n)),
    num('x', '가로 위치(px)', draft.x, 0, 1000000), num('y', '세로 위치(px)', draft.y, 0, 10000000), row('셀 크기가 바뀔 때', placement), originalSize);
  const rotation = section('회전 및 대칭', num('rot', '회전(°)', draft.rot ?? 0, -360, 360), check('flip', '좌우 대칭'), check('flipV', '상하 대칭'));
  const crop = section('자르기', el('p', { class: 'muted' }, '원본에서 잘라낼 비율입니다. 음수는 그림 바깥 여백을 늘립니다. 원본은 삭제되지 않습니다.'),
    ...[['l', '왼쪽 자르기(%)'], ['r', '오른쪽 자르기(%)'], ['t', '위쪽 자르기(%)'], ['b', '아래쪽 자르기(%)']].map(([side, label]) => num('crop.' + side, label, (draft.crop?.[side] ?? 0) * 100, -1000, 99, n => {
      draft.crop = setPictureCrop(draft.crop, side, n); fields['crop.' + side].value = Math.round(draft.crop[side] * 100000) / 1000;
    })), el('button', { type: 'button', class: 'btn', onclick: () => { draft.crop = undefined; sync(); } }, '자르기 초기화'));
  const body = el('div', { class: 'picture-editor' }, preview, el('div', { class: 'picture-settings' }, size, rotation, original.linked ? null : crop,
    section('대체 텍스트', text('alt', '그림 설명', true)),
    el('button', { type: 'button', class: 'btn', onclick: () => { Object.assign(draft, resetPictureFormatting()); sync(); } }, '그림 서식 원래대로')));
  refresh();
  return { body, read() {
    const bad = numbers.find(input => input.isConnected && (input.value === '' || !input.checkValidity()));
    if (bad) { bad.focus(); throw new Error(`${bad.getAttribute('aria-label')} 값을 범위 안에서 입력하세요.`); }
    const { name, alt, w, h, x, y, rot, flip, flipV, crop, lockAspect, placement, border, borderW, radius, shadow, opacity } = draft;
    return { name: name?.trim() || original.name, alt: alt ?? '', w, h, x, y, rot: rot || undefined, flip: flip || undefined, flipV: flipV || undefined, crop, lockAspect, placement, border, borderW, radius, shadow, opacity };
  } };
}
