import { el } from './ui.js';
import { setSafeHtml } from './safe-html.js';
import { resizePicture, setPictureCrop, dragPictureCrop, nudgePictureCrop, pictureTransform, resetPictureFormatting, pictureEffects, PICTURE_CROP_PRESETS, pictureCropPreset, pictureCm, picturePx, PICTURE_STYLES, pictureStylePatch, pictureSourcePatch, resetPictureSource } from './picture.js';
import { PICTURE_ARTISTIC, PICTURE_RECOLOR, normalizePictureVisual } from './picture-filter.js';
import { pictureMarkup } from './picture-render.js';
import { loadPictureBitmap, pictureBitmapCanvas, compressPicture } from './picture-raster.js';
import { removePictureBackground, pictureDataBytes, PICTURE_PIXEL_LIMIT } from './picture-pixels.js';

let pictureEditorId = 0;
const STYLE_SAMPLE = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="90" height="58"><rect width="90" height="58" fill="#daedf4"/><circle cx="66" cy="15" r="8" fill="#ffc75f"/><path d="M0 58V39L24 20 48 43 64 29 90 51V58Z" fill="#348668"/></svg>');
// 확인 전 초안만 변경합니다. 비동기 작업도 모달의 수명/요청 순서를 넘어서 적용하지 않습니다.
export function pictureEditor(original, { section = 'size' } = {}) {
  const draft = structuredClone(original); draft.lockAspect = original.lockAspect !== false;
  const fields = {}, numbers = [], tabs = [], panels = [], natural = { w: 0, h: 0 }, uid = 'picture-v2-' + ++pictureEditorId;
  let closed = false, busy = false, ticket = 0, bitmap = null, bitmapSrc = null, bgPixels = null, bgResult = null, bgSeed = null, compression = null;
  const marks = [], status = el('div', { class: 'picture-v2-status', role: 'status', 'aria-live': 'polite' });
  let previewScale = 1, activeSection = 'size', cropDrag = null;
  const pixels = el('div', { class: 'picture-v2-pixels' }), cropHandles = el('div', { class: 'picture-v2-crop-handles', hidden: true });
  const frame = el('div', { class: 'picture-preview-frame picture-v2-frame' }, pixels, cropHandles);
  const stage = el('div', { class: 'picture-preview-stage picture-v2-stage' }, frame);
  const caption = el('div', { class: 'picture-preview-caption', 'aria-live': 'polite' });
  const preview = el('div', { class: 'picture-preview picture-v2-preview', 'aria-label': '그림 미리보기' }, stage, caption);
  const nav = el('div', { class: 'picture-v2-tabs', role: 'tablist', 'aria-label': '그림 서식 범주' });
  const settings = el('div', { class: 'picture-settings picture-v2-settings' });
  const valueAt = key => key.split('.').reduce((o, k) => o?.[k], draft);
  const put = (key, value) => { const keys = key.split('.'); let obj = draft; for (const k of keys.slice(0, -1)) obj = obj[k] = { ...(typeof obj[k] === 'object' ? obj[k] : {}) }; obj[keys.at(-1)] = value; };
  const refresh = () => {
    if (closed) return;
    const angle = (Number(draft.rot) || 0) * Math.PI / 180, ref = normalizePictureVisual(draft).reflection;
    const bw = Math.abs(draft.w * Math.cos(angle)) + Math.abs(draft.h * Math.sin(angle)), bh = Math.abs(draft.w * Math.sin(angle)) + Math.abs(draft.h * Math.cos(angle)) + (ref ? draft.h * ref.size + ref.gap : 0);
    const availableW = Math.max(80, stage.clientWidth ? stage.clientWidth - 24 : 260), availableH = Math.max(80, stage.clientHeight ? stage.clientHeight - 32 : 180);
    const k = Math.min(1, availableW / Math.max(4, bw), availableH / Math.max(4, bh));
    previewScale = k;
    Object.assign(frame.style, { width: draft.w + 'px', height: draft.h + 'px', transform: `translate(-50%,-50%) scale(${k}) ${pictureTransform(draft)}`, transformOrigin: 'center' });
    setSafeHtml(pixels, pictureMarkup(draft, uid));
    for (const button of cropHandles.children) { button.style.width = button.style.height = `${14 / k}px`; button.style.borderWidth = `${1 / k}px`; }
    caption.textContent = `${Math.round(draft.w * 100) / 100} × ${Math.round(draft.h * 100) / 100}px · ${pictureCm(draft.w)} × ${pictureCm(draft.h)}cm · 회전 ${draft.rot ?? 0}°`;
  };
  const row = (label, input) => el('label', { class: 'picture-field picture-v2-field' }, el('span', {}, label), input);
  const note = text => el('p', { class: 'muted picture-v2-note' }, text);
  const getters = {};
  const sync = () => {
    for (const [key, input] of Object.entries(fields)) {
      const value = getters[key] ? getters[key]() : valueAt(key);
      if (input.type === 'checkbox') input.checked = !!value; else input.value = value ?? '';
    }
    refresh();
  };
  const num = (key, label, min, max, get = () => valueAt(key) ?? 0, set = value => put(key, value), step = 'any') => {
    const input = fields[key] = el('input', { type: 'number', 'aria-label': label, min, max, step, value: get() }); numbers.push(input); getters[key] = get;
    input.addEventListener('input', () => { if (input.value === '' || !input.checkValidity()) return; set(Number(input.value)); refresh(); }); return row(label, input);
  };
  const check = (key, label, get = () => valueAt(key), set = value => put(key, value)) => {
    const input = fields[key] = el('input', { type: 'checkbox', 'aria-label': label, checked: !!get() }); getters[key] = get;
    input.addEventListener('change', () => { set(input.checked); refresh(); }); return row(label, input);
  };
  const select = (key, label, options, get = () => valueAt(key), set = value => put(key, value)) => {
    const input = fields[key] = el('select', { 'aria-label': label }, options.map(({ id, name }) => el('option', { value: id }, name))); getters[key] = get; input.value = get();
    input.addEventListener('change', () => { set(input.value); refresh(); }); return row(label, input);
  };
  const text = (key, label, multiline = false) => {
    const input = fields[key] = el(multiline ? 'textarea' : 'input', { 'aria-label': label, ...(multiline ? { rows: 3 } : { type: 'text' }) }); input.value = valueAt(key) ?? '';
    input.addEventListener('input', () => { put(key, input.value); }); return row(label, input);
  };
  const color = (key, label, get, set) => {
    const input = fields[key] = el('input', { type: 'color', 'aria-label': label, value: get() }); getters[key] = get;
    input.addEventListener('input', () => { set(input.value); refresh(); }); return row(label, input);
  };
  const button = (label, action, cls = '') => el('button', { type: 'button', class: 'btn ' + cls, onclick: action }, label);
  const activate = name => {
    const index = tabs.findIndex(t => t.dataset.section === name), active = index < 0 ? 0 : index;
    tabs.forEach((tab, i) => { tab.setAttribute('aria-selected', String(i === active)); tab.tabIndex = i === active ? 0 : -1; panels[i].hidden = i !== active; });
    activeSection = tabs[active]?.dataset.section ?? 'size'; cropHandles.hidden = activeSection !== 'crop' || !!original.linked || !!original.media;
  };
  const panel = (id, name, ...children) => {
    const tab = el('button', { type: 'button', role: 'tab', id: uid + '-tab-' + id, 'aria-controls': uid + '-panel-' + id, dataset: { section: id }, onclick: () => activate(id) }, name);
    tab.addEventListener('keydown', e => { const i = tabs.indexOf(tab), next = e.key === 'ArrowRight' ? (i + 1) % tabs.length : e.key === 'ArrowLeft' ? (i + tabs.length - 1) % tabs.length : e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : -1; if (next >= 0) { e.preventDefault(); e.stopPropagation(); activate(tabs[next].dataset.section); tabs[next].focus(); } });
    const node = el('section', { role: 'tabpanel', id: uid + '-panel-' + id, 'aria-labelledby': tab.id, class: 'picture-v2-panel' }, ...children); tabs.push(tab); panels.push(node); nav.append(tab); settings.append(node); return node;
  };
  const setSize = (axis, value) => { Object.assign(draft, resizePicture(draft, axis, value)); for (const key of ['w', 'h', 'widthCm', 'heightCm']) if (fields[key]) fields[key].value = getters[key](); };
  const originalSize = button('원본 크기', () => { if (natural.w) { Object.assign(draft, resizePicture({ w: natural.w, h: natural.h }, 'w', natural.w)); sync(); } }); originalSize.disabled = true;
  panel('size', '크기 및 속성', text('name', '이름'), check('lockAspect', '가로 세로 비율 고정'),
    num('w', '너비(px)', 4, 20000, () => draft.w, n => setSize('w', n)), num('h', '높이(px)', 4, 20000, () => draft.h, n => setSize('h', n)),
    num('widthCm', '너비(cm)', .106, 529.167, () => pictureCm(draft.w), n => setSize('w', picturePx(n))), num('heightCm', '높이(cm)', .106, 529.167, () => pictureCm(draft.h), n => setSize('h', picturePx(n))),
    num('x', '가로 위치(px)', 0, 1000000), num('y', '세로 위치(px)', 0, 10000000),
    select('placement', '셀 크기가 바뀔 때', [{ id: 'twoCell', name: '위치와 크기 변함' }, { id: 'oneCell', name: '위치만 변함' }, { id: 'absolute', name: '변하지 않음 (위치 고정)' }], () => draft.placement ?? 'twoCell'),
    num('rot', '회전(°)', -360, 360), check('flip', '좌우 대칭'), check('flipV', '상하 대칭'), originalSize);
  panel('corrections', '보정',
    num('brightness', '밝기(%)', -100, 100, () => (draft.correction?.brightness ?? 0) * 100, n => put('correction.brightness', n / 100)),
    num('contrast', '대비(%)', -100, 100, () => (draft.correction?.contrast ?? 0) * 100, n => put('correction.contrast', n / 100)),
    num('sharpness', '선명하게(%)', 0, 100, () => (draft.correction?.sharpness ?? 0) * 100, n => put('correction.sharpness', n / 100)),
    button('보정 초기화', () => { draft.correction = undefined; sync(); }), note('원본 픽셀은 유지하며 화면에서 보정합니다. 지원되지 않는 Excel 효과는 저장할 때 보정된 그림으로 함께 보존합니다.'));
  panel('color', '색',
    num('saturation', '채도(%)', 0, 200, () => (draft.color?.saturation ?? 1) * 100, n => put('color.saturation', n / 100)),
    num('temperature', '색 온도(차갑게 −100 / 따뜻하게 100)', -100, 100, () => (draft.color?.temperature ?? 0) * 100, n => put('color.temperature', n / 100)),
    select('recolor', '다시 칠하기', [...PICTURE_RECOLOR, ...(draft.color?.duotone ? [{ id: 'imported', name: '가져온 두 색' }] : [])], () => draft.color?.duotone ? 'imported' : draft.color?.recolor ?? 'none', value => { if (value !== 'imported') { put('color.recolor', value); draft.color.duotone = undefined; } }),
    button('색 초기화', () => { draft.color = undefined; sync(); }));
  panel('artistic', '꾸밈 효과', select('artisticType', '꾸밈 효과 종류', PICTURE_ARTISTIC, () => draft.artistic?.type ?? 'none', v => put('artistic.type', v)),
    num('artisticAmount', '꾸밈 효과 강도(%)', 0, 100, () => (draft.artistic?.amount ?? .5) * 100, n => put('artistic.amount', n / 100)), note('연필·포스터·물감·입자는 WIXEL의 픽셀 효과입니다. Excel의 동명 효과와 완전히 같지는 않습니다.'));
  const gallery = el('div', { class: 'picture-v2-gallery', 'aria-label': '그림 스타일' });
  for (const style of PICTURE_STYLES) {
    const thumb = el('span', { class: 'picture-v2-style-thumb', 'aria-hidden': 'true' });
    setSafeHtml(thumb, pictureMarkup({ src: STYLE_SAMPLE, ...pictureStylePatch(style.id), w: 90, h: 58 }, uid + '-style-' + style.id));
    gallery.append(button(style.name, () => { Object.assign(draft, pictureStylePatch(style.id, draft)); sync(); gallery.querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.style === style.id))); }, 'picture-v2-style'));
    const b = gallery.lastChild; b.prepend(thumb); b.dataset.style = style.id; b.setAttribute('aria-pressed', 'false');
  }
  panel('styles', '그림 스타일', note('24가지 테두리·그림자·네온·반사 조합. 크기·자르기·원본 내용은 바뀌지 않습니다.'), gallery);
  const shadow = () => pictureEffects({ ...draft, shadow: draft.shadow || true }).shadow;
  const updateShadow = (key, value) => { draft.shadow = { ...shadow(), [key]: value }; if (fields.shadow) fields.shadow.checked = true; };
  const glow = () => normalizePictureVisual({ ...draft, glow: draft.glow || {} }).glow;
  const reflect = () => normalizePictureVisual({ ...draft, reflection: draft.reflection || {} }).reflection;
  panel('effects', '테두리 및 효과',
    num('transparency', '그림 투명도(%)', 0, 100, () => (1 - pictureEffects(draft).opacity) * 100, n => { draft.opacity = 1 - n / 100; }),
    check('borderEnabled', '그림 테두리', () => !!draft.border, on => { draft.border = on ? fields.borderColor.value : undefined; }),
    color('borderColor', '그림 테두리 색', () => draft.border ?? '#000000', value => { draft.border = value; fields.borderEnabled.checked = true; }),
    num('borderW', '그림 테두리 두께(px)', 0, 100, () => draft.borderW ?? 2),
    select('borderDash', '그림 테두리 선 종류', [{ id: 'solid', name: '실선' }, { id: 'dash', name: '파선' }, { id: 'dot', name: '점선' }, { id: 'dashDot', name: '일점쇄선' }], () => draft.borderDash ?? 'solid'),
    num('radius', '둥근 모서리(px)', 0, 10000, () => draft.radius ?? 0, n => { draft.radius = pictureEffects({ ...draft, radius: n }).radius; fields.radius.value = draft.radius; }),
    check('shadow', '그림자 표시', () => !!draft.shadow, on => { draft.shadow = on ? { color: fields['shadow.color'].value, opacity: 1 - Number(fields['shadow.opacity'].value) / 100, blur: Number(fields['shadow.blur'].value), dx: Number(fields['shadow.dx'].value), dy: Number(fields['shadow.dy'].value) } : undefined; }),
    color('shadow.color', '그림자 색', () => shadow().color, value => updateShadow('color', value)),
    num('shadow.opacity', '그림자 투명도(%)', 0, 100, () => (1 - shadow().opacity) * 100, n => updateShadow('opacity', 1 - n / 100)),
    num('shadow.blur', '그림자 흐리게(px)', 0, 1000, () => shadow().blur, n => updateShadow('blur', n)),
    num('shadow.dx', '그림자 가로 거리(px)', -1000, 1000, () => shadow().dx, n => updateShadow('dx', n)),
    num('shadow.dy', '그림자 세로 거리(px)', -1000, 1000, () => shadow().dy, n => updateShadow('dy', n)),
    check('glowEnabled', '네온 표시', () => !!draft.glow, on => { draft.glow = on ? glow() : undefined; }),
    color('glow.color', '네온 색', () => glow().color, v => { draft.glow = { ...glow(), color: v }; fields.glowEnabled.checked = true; }),
    num('glow.size', '네온 크기(px)', 0, 100, () => glow().size, n => { draft.glow = { ...glow(), size: n }; fields.glowEnabled.checked = true; }),
    num('glow.opacity', '네온 투명도(%)', 0, 100, () => (1 - glow().opacity) * 100, n => { draft.glow = { ...glow(), opacity: 1 - n / 100 }; fields.glowEnabled.checked = true; }),
    num('softEdge', '부드러운 가장자리(px)', 0, 100, () => draft.softEdge ?? 0),
    check('reflectionEnabled', '반사 표시', () => !!draft.reflection, on => { draft.reflection = on ? reflect() : undefined; }),
    num('reflection.opacity', '반사 투명도(%)', 0, 100, () => (1 - reflect().opacity) * 100, n => { draft.reflection = { ...reflect(), opacity: 1 - n / 100 }; fields.reflectionEnabled.checked = true; }),
    num('reflection.size', '반사 크기(%)', 0, 100, () => reflect().size * 100, n => { draft.reflection = { ...reflect(), size: n / 100 }; fields.reflectionEnabled.checked = true; }),
    num('reflection.gap', '반사 거리(px)', 0, 100, () => reflect().gap, n => { draft.reflection = { ...reflect(), gap: n }; fields.reflectionEnabled.checked = true; }));
  const cropPanel = panel('crop', '자르기', note('잘라낼 비율을 지정합니다. 음수는 여백을 늘립니다. 원본은 보존됩니다.'),
    select('cropPreset', '자르기 가로 세로 비율', PICTURE_CROP_PRESETS, () => 'original', id => { Object.assign(draft, pictureCropPreset(draft, PICTURE_CROP_PRESETS.find(p => p.id === id)?.ratio, natural.w ? natural : draft)); sync(); fields.cropPreset.value = id; }),
    ...[['l', '왼쪽 자르기(%)'], ['r', '오른쪽 자르기(%)'], ['t', '위쪽 자르기(%)'], ['b', '아래쪽 자르기(%)']].map(([side, label]) => num('crop.' + side, label, -1000, 99, () => Math.round((draft.crop?.[side] ?? 0) * 100000) / 1000, n => { draft.crop = setPictureCrop(draft.crop, side, n); fields['crop.' + side].value = Math.round(draft.crop[side] * 100000) / 1000; })),
    button('자르기 초기화', () => { draft.crop = undefined; sync(); }));
  const syncCrop = () => { for (const side of ['l', 'r', 't', 'b']) fields['crop.' + side].value = Math.round((draft.crop?.[side] ?? 0) * 100000) / 1000; refresh(); };
  for (const [edge, label, x, y] of [['lt', '왼쪽 위', 0, 0], ['t', '위쪽', 50, 0], ['rt', '오른쪽 위', 100, 0], ['r', '오른쪽', 100, 50], ['rb', '오른쪽 아래', 100, 100], ['b', '아래쪽', 50, 100], ['lb', '왼쪽 아래', 0, 100], ['l', '왼쪽', 0, 50]]) {
    const handle = el('button', { type: 'button', class: 'picture-v2-crop-handle', 'aria-label': `${label} 자르기 조절점`, title: '드래그하거나 방향키로 자르기 조정', style: { left: x + '%', top: y + '%' } });
    handle.addEventListener('pointerdown', event => { event.preventDefault(); event.stopPropagation(); handle.focus({ preventScroll: true }); cropDrag = { crop: { ...draft.crop }, x: event.clientX, y: event.clientY, edge, pointer: event.pointerId, handle }; handle.setPointerCapture(event.pointerId); });
    handle.addEventListener('pointermove', event => { if (!cropDrag || cropDrag.handle !== handle) return; const angle = (draft.rot || 0) * Math.PI / 180, dx = (event.clientX - cropDrag.x) / previewScale, dy = (event.clientY - cropDrag.y) / previewScale;
      const u = (Math.cos(angle) * dx + Math.sin(angle) * dy) * (draft.flip ? -1 : 1) / draft.w, v = (-Math.sin(angle) * dx + Math.cos(angle) * dy) * (draft.flipV ? -1 : 1) / draft.h;
      draft.crop = dragPictureCrop(cropDrag.crop, edge, u, v); syncCrop(); });
    const finishCrop = () => { cropDrag = null; }; handle.addEventListener('pointerup', finishCrop); handle.addEventListener('lostpointercapture', finishCrop); handle.addEventListener('pointercancel', () => { if (cropDrag) { draft.crop = cropDrag.crop; syncCrop(); } finishCrop(); });
    handle.addEventListener('keydown', event => { if (event.key === 'Escape' && cropDrag) { event.preventDefault(); event.stopPropagation(); draft.crop = cropDrag.crop; cropDrag = null; syncCrop(); return; }
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return; event.preventDefault(); event.stopPropagation(); const step = event.shiftKey ? .05 : .01;
      draft.crop = nudgePictureCrop(draft.crop, edge, event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0, event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0); syncCrop(); });
    cropHandles.append(handle);
  }
  cropPanel.prepend(note('미리보기의 8개 조절점을 드래그할 수 있습니다. 조절점에 초점을 두고 방향키로 1%포인트, Shift+방향키로 5%포인트씩 조정합니다.'));
  panel('alt', '대체 텍스트', text('alt', '그림 설명', true), note('화면을 보지 않고도 그림을 이해할 수 있는 설명을 입력하세요.'));
  const readBitmap = async () => { if (bitmap && bitmapSrc === draft.src) return bitmap; const src = draft.src, loaded = await loadPictureBitmap(src); if (draft.src === src) { bitmap = loaded; bitmapSrc = src; } return loaded; };
  const run = async fn => { const current = ++ticket; busy = true; status.textContent = '그림을 처리하는 중입니다…'; try { await fn(current); } catch (e) { if (!closed && current === ticket) status.textContent = e.message; } finally { if (current === ticket) busy = false; } };
  const valid = current => !closed && current === ticket;
  const bgCanvas = el('canvas', { class: 'picture-v2-background-canvas', 'aria-label': '배경 제거 미리보기 · 색 선택 또는 유지·제거 표시', tabindex: '0' });
  const bgColor = el('input', { type: 'color', value: '#ffffff', 'aria-label': '제거할 배경 색' }), tolerance = el('input', { type: 'number', min: 0, max: 100, value: 18, 'aria-label': '배경 색 허용 오차(%)' });
  const brush = el('input', { type: 'number', min: 1, max: 100, value: 8, 'aria-label': '유지·제거 붓 크기(px)' });
  const mode = el('select', { 'aria-label': '배경 제거 표시 도구' }, [['pick', '색 선택'], ['keep', '유지할 영역 표시'], ['remove', '제거할 영역 표시']].map(([value, name]) => el('option', { value }, name)));
  const drawBackground = () => {
    if (!bgPixels) return; bgCanvas.width = bgPixels.width; bgCanvas.height = bgPixels.height;
    const ctx = bgCanvas.getContext('2d'); ctx.putImageData(new ImageData(bgResult?.data ?? bgPixels.data, bgPixels.width, bgPixels.height), 0, 0);
  };
  const previewBackground = async current => {
    const image = await readBitmap(); if (!valid(current)) return;
    if (image.naturalWidth * image.naturalHeight > PICTURE_PIXEL_LIMIT) throw new Error('배경 제거는 400만 픽셀 이하에서 지원합니다. 먼저 그림을 압축하세요.');
    if (!bgPixels) { const canvas = pictureBitmapCanvas(image); bgPixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height); }
    const hex = bgColor.value, target = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
    bgResult = removePictureBackground(bgPixels, { color: target, tolerance: Number(tolerance.value), seed: bgSeed, marks });
    if (!valid(current)) return; drawBackground(); status.textContent = '배경 제거 미리보기를 만들었습니다. 적용하면 초안에 반영됩니다.'; bgApply.disabled = false;
  };
  const bgApply = button('배경 제거 적용', () => { if (!bgResult || busy) return; const canvas = document.createElement('canvas'); canvas.width = bgResult.width; canvas.height = bgResult.height; canvas.getContext('2d').putImageData(new ImageData(bgResult.data, bgResult.width, bgResult.height), 0, 0); Object.assign(draft, pictureSourcePatch(draft, canvas.toDataURL('image/png'), { w: bgPixels.width, h: bgPixels.height })); bitmap = null; compression = null; refresh(); status.textContent = '배경 제거를 초안에 적용했습니다. 확인을 눌러 문서에 반영하세요.'; }); bgApply.disabled = true;
  for (const control of [bgColor, tolerance]) control.addEventListener('input', () => { ticket++; busy = false; bgResult = null; bgApply.disabled = true; if (control === bgColor) bgSeed = null; });
  let painting = false;
  const markAt = event => { if (!bgPixels || busy) return; const rect = bgCanvas.getBoundingClientRect(), x = Math.max(0, Math.min(bgPixels.width - 1, (event.clientX - rect.left) * bgPixels.width / rect.width)), y = Math.max(0, Math.min(bgPixels.height - 1, (event.clientY - rect.top) * bgPixels.height / rect.height));
    if (mode.value === 'pick') { const at = (Math.floor(y) * bgPixels.width + Math.floor(x)) * 4; bgColor.value = '#' + [...bgPixels.data.slice(at, at + 3)].map(n => n.toString(16).padStart(2, '0')).join(''); bgSeed = { x, y }; }
    else if (marks.length < 2000) marks.push({ x, y, radius: Math.max(1, Math.min(100, Number(brush.value) || 8)), mode: mode.value });
    else { status.textContent = '표시는 2,000개 이내로 지정하세요. 표시 초기화 후 다시 시도할 수 있습니다.'; return; }
    if (mode.value === 'pick') run(previewBackground);
    else {
      bgApply.disabled = true; status.textContent = '유지·제거 표시 중입니다. 손을 놓으면 미리보기를 계산합니다.';
      const ctx = bgCanvas.getContext('2d'), radius = Math.max(1, Math.min(100, Number(brush.value) || 8));
      ctx.strokeStyle = mode.value === 'keep' ? '#008f39' : '#d8203e'; ctx.lineWidth = Math.max(1, bgPixels.width / Math.max(1, bgCanvas.clientWidth)); ctx.beginPath(); ctx.arc(x, y, radius, 0, Math.PI * 2); ctx.stroke();
    }
  };
  bgCanvas.addEventListener('pointerdown', e => { e.preventDefault(); painting = true; bgCanvas.setPointerCapture(e.pointerId); markAt(e); });
  bgCanvas.addEventListener('pointermove', e => { if (painting && mode.value !== 'pick') markAt(e); });
  const finishMarking = () => { if (painting && mode.value !== 'pick') run(previewBackground); painting = false; };
  bgCanvas.addEventListener('pointerup', finishMarking); bgCanvas.addEventListener('pointercancel', finishMarking);
  const backgroundPanel = panel('background', '배경 제거', note('단색 또는 비슷한 색의 배경에 적합합니다. 색을 선택하고 유지·제거 표시로 다듬으세요. 최대 400만 픽셀.'), row('제거할 배경 색', bgColor), row('배경 색 허용 오차(%)', tolerance), row('표시 도구', mode), row('붓 크기(px)', brush),
    button('배경 제거 미리보기', () => run(previewBackground)), bgCanvas, button('표시 초기화', () => { marks.length = 0; bgSeed = null; run(previewBackground); }), bgApply);
  const resolution = el('select', { 'aria-label': '압축 해상도' }, [['original', '원본 해상도'], ['2560', '긴 변 2560px'], ['1920', '긴 변 1920px'], ['1280', '긴 변 1280px'], ['640', '긴 변 640px']].map(([value, name]) => el('option', { value }, name))); resolution.value = '1920';
  const quality = el('input', { type: 'number', 'aria-label': '압축 품질(%)', min: 10, max: 100, value: 85 });
  const mime = el('select', { 'aria-label': '압축 파일 형식' }, [['image/jpeg', 'JPEG · 투명 배경은 흰색'], ['image/png', 'PNG · 투명도 보존'], ['image/webp', 'WebP · 투명도 보존']].map(([value, name]) => el('option', { value }, name)));
  const compare = el('div', { class: 'picture-v2-compress-result', role: 'status' }), compressedImage = el('img', { class: 'picture-v2-compress-image', alt: '압축 결과 미리보기' });
  const keepOriginal = el('input', { type: 'checkbox', checked: true, 'aria-label': '원본 복원 데이터 보관' });
  const compressionSummary = () => { if (!compression) return; const before = pictureDataBytes(draft.src), after = pictureDataBytes(compression.src), savedOriginal = keepOriginal.checked ? pictureDataBytes(draft.originalSrc || draft.src) ?? 0 : 0;
    compare.textContent = `${before === null ? '원본 용량 확인 불가' : '현재 ' + Math.round(before / 1024) + 'KB'} → 압축 그림 ${Math.round(after / 1024)}KB · ${compression.w} × ${compression.h}px · 복원 데이터 포함 약 ${Math.round((after + savedOriginal) / 1024)}KB${!keepOriginal.checked ? ' · 원본 복원 데이터 삭제: 저장 후 원본 복원이 불가능합니다.' : ''}${before && after > before ? ' · 압축 그림 용량이 더 큽니다.' : ''}`; };
  keepOriginal.addEventListener('change', compressionSummary);
  const compressApply = button('압축 적용', () => { if (!compression || busy) return; Object.assign(draft, pictureSourcePatch(draft, compression.src, { w: compression.originalWidth, h: compression.originalHeight })); if (!keepOriginal.checked) for (const key of ['originalSrc', 'originalWidth', 'originalHeight', 'originalEmf', 'originalPng']) draft[key] = undefined; bitmap = null; bgPixels = bgResult = null; marks.length = 0; refresh(); status.textContent = keepOriginal.checked ? '압축을 초안에 적용했습니다. 그림 크기와 원본 보관본은 유지됩니다.' : '압축을 초안에 적용했습니다. 원본 복원 데이터는 확인 후 제거됩니다. 실행 취소는 가능합니다.'; }); compressApply.disabled = true;
  const previewCompression = () => run(async current => { if (!quality.checkValidity()) throw new Error('압축 품질은 10~100%로 입력하세요.'); const max = resolution.value === 'original' ? 32000 : Number(resolution.value); const result = await compressPicture(draft.src, { maxWidth: max, maxHeight: max, quality: Number(quality.value) / 100, type: mime.value }); if (!valid(current)) return; compression = result; compressedImage.src = result.src; compressionSummary(); compressApply.disabled = false; status.textContent = '압축 미리보기가 준비되었습니다.'; });
  for (const control of [resolution, quality, mime]) control.addEventListener('change', () => { ticket++; busy = false; compression = null; compressApply.disabled = true; });
  const compressionPanel = panel('compress', '압축', note('이 브라우저에서만 처리합니다. 원본 보관본을 유지하면 문서 전체 용량이 바로 줄지는 않을 수 있습니다. JPEG는 투명 배경을 흰색으로 바꿉니다. 합계는 그림 데이터 기준이며 XLSX 파일 크기와 다릅니다.'), row('해상도', resolution), row('품질(%)', quality), row('파일 형식', mime), row('원본 복원 데이터 보관', keepOriginal), note('원본 보관을 해제하고 저장하면 이 문서의 원본 그림 복원이 불가능합니다. 외부 원본 파일은 변경하지 않습니다.'), button('압축 미리보기', previewCompression), compare, compressedImage, compressApply,
    button('원본 그림 복원', () => { ticket++; busy = false; Object.assign(draft, resetPictureSource(draft)); bitmap = null; bgPixels = bgResult = compression = null; marks.length = 0; sync(); status.textContent = '보관한 원본 그림을 초안에 복원했습니다.'; }));
  if (/^data:image\/gif[;,]|\.gif(?:[?#]|$)/i.test(original.src ?? '')) for (const node of [backgroundPanel, compressionPanel]) node.prepend(note('움직이는 GIF에 적용하면 현재 한 프레임의 정지 그림이 됩니다. 움직임을 유지하려면 적용하지 마세요. 원본 보관 시 복원할 수 있습니다.'));
  if (original.linked || original.media) { for (const item of [cropPanel, backgroundPanel, compressionPanel]) { for (const control of item.querySelectorAll('input,select,button')) control.disabled = true; item.prepend(note('연결된 그림과 미디어에는 이 픽셀 편집을 적용하지 않습니다.')); } }
  const body = el('div', { class: 'picture-editor picture-v2' }, nav, el('div', { class: 'picture-v2-body' }, preview, settings), status,
    button('그림 서식 원래대로', () => { Object.assign(draft, resetPictureFormatting()); sync(); }));
  if (!original.linked && !original.media) {
    const source = new Image(); source.referrerPolicy = 'no-referrer'; source.onload = () => { if (closed) return; natural.w = source.naturalWidth; natural.h = source.naturalHeight; originalSize.disabled = !natural.w; originalSize.title = natural.w + ' × ' + natural.h + 'px'; }; source.onerror = () => { if (!closed) originalSize.title = '원본 이미지 크기를 읽지 못했습니다.'; }; source.src = original.originalSrc || original.src;
  }
  const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(() => refresh()) : null; observer?.observe(stage);
  activate({ transparency: 'effects', border: 'effects' }[section] ?? section); sync();
  return { body, read() {
    if (closed) throw new Error('그림 편집 창이 닫혔습니다.');
    if (busy) throw new Error('그림 처리가 끝난 뒤 확인을 누르세요.');
    const bad = numbers.find(input => input.value === '' || !input.checkValidity());
    if (bad) { const p = bad.closest('[role=tabpanel]'), index = panels.indexOf(p); if (index >= 0) activate(tabs[index].dataset.section); bad.focus(); throw new Error(`${bad.getAttribute('aria-label')} 값을 범위 안에서 입력하세요.`); }
    const result = structuredClone(draft); result.name = result.name?.trim() || original.name; result.alt = result.alt ?? ''; result.rot = result.rot || undefined; result.flip = result.flip || undefined; result.flipV = result.flipV || undefined; return result;
  }, dispose() { closed = true; ticket++; observer?.disconnect(); bitmap = bgPixels = bgResult = compression = null; marks.length = 0; } };
}
