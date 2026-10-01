import { el, openDialog } from './ui.js';
import { setSafeHtml } from './safe-html.js';
import { SMARTART_CATEGORIES, SMARTART_LAYOUTS, SMARTART_PALETTES, normalizeSmartArt, editSmartArt, smartArtPicture } from './smartart.js';
import { smartArtSvg } from './smartart-render.js';

/** Draft-only modal: the caller performs one guarded transaction on confirmation. */
export function openSmartArtEditor({ shape, onCommit, onConvert }) {
  let draft = structuredClone(shape), selected = draft.smartArt.nodes[0].id, category = SMARTART_LAYOUTS.find(x => x.id === draft.smartArt.layout)?.category ?? '목록', closed = false;
  const preview = el('div', { class: 'sa-preview', 'aria-label': 'SmartArt 미리 보기' }), gallery = el('div', { class: 'sa-gallery', role: 'group', 'aria-label': 'SmartArt 레이아웃' }), list = el('div', { class: 'sa-text-list', 'aria-label': '텍스트 창' }), notice = el('div', { class: 'muted sa-note', role: 'status' });
  const redrawPreview = () => { setSafeHtml(preview, smartArtSvg({ ...draft, w: 560, h: 300 })); for (const p of preview.querySelectorAll('[data-smartart-node]')) { p.style.cursor = 'pointer'; p.addEventListener('click', () => { selected = p.dataset.smartartNode; drawText(); list.querySelector(`[data-node-index="${draft.smartArt.nodes.findIndex(n => n.id === selected)}"] textarea`)?.focus(); }); } };
  const apply = action => { try { const before = draft.smartArt.nodes, at = before.findIndex(n => n.id === selected); draft.smartArt = editSmartArt(draft.smartArt, selected, action); if (action === 'add') selected = draft.smartArt.nodes.find(n => !before.some(o => o.id === n.id)).id; if (!draft.smartArt.nodes.some(n => n.id === selected)) selected = draft.smartArt.nodes[Math.min(at, draft.smartArt.nodes.length - 1)].id; drawText(); redrawPreview(); notice.textContent = ''; } catch (e) { notice.textContent = e.message; } };
  const drawText = () => {
    list.replaceChildren(...draft.smartArt.nodes.map((node, i) => {
      const input = el('textarea', { rows: 2, maxlength: 500, value: node.text, 'aria-label': `항목 ${i + 1} 텍스트` });
      input.addEventListener('focus', () => { selected = node.id; list.querySelectorAll('.sa-text-row').forEach(row => row.classList.toggle('on', row.dataset.nodeIndex === String(i))); });
      input.addEventListener('input', () => { draft.smartArt = editSmartArt(draft.smartArt, node.id, 'text', input.value); redrawPreview(); });
      input.addEventListener('keydown', event => { if (event.key === 'Tab' && !event.ctrlKey && !event.altKey) { event.preventDefault(); selected = node.id; apply(event.shiftKey ? 'promote' : 'demote'); list.querySelector(`[data-node-index="${draft.smartArt.nodes.findIndex(n => n.id === selected)}"] textarea`)?.focus(); } });
      return el('div', { class: `sa-text-row${node.id === selected ? ' on' : ''}`, 'data-node-index': i, style: { marginLeft: `${node.level * 14}px` } }, el('span', {}, `${i + 1}.`), input);
    }));
  };
  const drawGallery = () => {
    gallery.replaceChildren(...SMARTART_LAYOUTS.filter(x => x.category === category).map(layout => {
      const sample = { ...draft, id: `thumb-${layout.id}`, w: 140, h: 82, smartArt: { ...draft.smartArt, layout: layout.id, nodes: draft.smartArt.nodes.slice(0, 4).map((n, i) => ({ ...n, text: '', picture: undefined, level: layout.category === '계층' && i ? 1 : 0 })) } };
      return el('button', { type: 'button', class: `sa-layout${layout.id === draft.smartArt.layout ? ' on' : ''}`, 'aria-label': layout.name, 'aria-pressed': String(layout.id === draft.smartArt.layout), onclick: () => { draft.smartArt.layout = layout.id; drawGallery(); redrawPreview(); } }, el('span', { html: smartArtSvg(sample) }), el('span', {}, layout.name));
    }));
  };
  const categories = el('select', { 'aria-label': 'SmartArt 범주' }, SMARTART_CATEGORIES.map(c => el('option', { value: c, selected: c === category }, c)));
  categories.addEventListener('change', () => { category = categories.value; drawGallery(); });
  const palette = el('select', { 'aria-label': 'SmartArt 색 변경' }, el('option', { value: 'current' }, '현재 색'), SMARTART_PALETTES.map((p, i) => el('option', { value: i }, p[0])));
  palette.addEventListener('change', () => { if (palette.value !== 'current') { draft.smartArt.palette = SMARTART_PALETTES[Number(palette.value)].slice(1); redrawPreview(); } });
  const style = el('select', { 'aria-label': 'SmartArt 스타일' }, [['flat','단색'],['outline','윤곽선'],['shadow','그림자']].map(([v,t]) => el('option', { value: v, selected: v === draft.smartArt.style }, t)));
  style.addEventListener('change', () => { draft.smartArt.style = style.value; redrawPreview(); });
  const file = el('input', { type: 'file', accept: 'image/png,image/jpeg', 'aria-label': '선택 항목 그림 파일', style: { maxWidth: '210px' } });
  file.addEventListener('change', async () => {
    const f = file.files?.[0], target = selected; if (!f) return;
    if (f.size > 2 * 1024 * 1024 || !['image/png', 'image/jpeg'].includes(f.type)) { notice.textContent = '2MB 이하 PNG/JPEG 그림을 선택하세요.'; return; }
    try {
      const src = await new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = () => reject(new Error('그림 파일을 읽지 못했습니다.')); reader.readAsDataURL(f); });
      if (closed) return; if (!smartArtPicture(src)) throw new Error('PNG/JPEG 그림을 확인하세요.');
      const model = structuredClone(draft.smartArt), node = model.nodes.find(n => n.id === target); if (!node) return; node.picture = src; draft.smartArt = normalizeSmartArt(model); if (!draft.smartArt.layout.startsWith('picture')) { draft.smartArt.layout = 'pictureCards'; category = '그림'; categories.value = category; drawGallery(); } redrawPreview(); notice.textContent = '';
    } catch (e) { notice.textContent = e.message; }
  });
  const commands = [['add','항목 추가'],['delete','항목 삭제'],['promote','수준 올리기'],['demote','수준 내리기'],['up','위로 이동'],['down','아래로 이동']];
  const body = el('div', { class: 'sa-editor' },
    el('div', { class: 'sa-controls' }, el('label', {}, '범주 ', categories), el('label', {}, '색 변경 ', palette), el('label', {}, '스타일 ', style)), gallery,
    el('div', { class: 'sa-workarea' }, el('section', {}, el('div', { class: 'sa-toolbar' }, commands.map(([a,t]) => el('button', { class: 'btn', type: 'button', onclick: () => apply(a) }, t))), list), preview),
    el('label', { class: 'sa-picture' }, '선택 항목 그림 ', file), notice,
    el('div', { class: 'muted sa-note' }, '텍스트 창: Tab 수준 내리기 · Shift+Tab 수준 올리기. 최대 60개. Excel에는 편집 가능한 일반 도형 그룹으로 저장되며, Excel의 기본 SmartArt는 아닙니다.'));
  drawGallery(); drawText(); redrawPreview();
  return openDialog({ title: 'SmartArt 그래픽', body, width: 960, onClose: () => { closed = true; }, buttons: [
    ...(onConvert ? [{ label: '일반 도형으로 변환', action: () => onConvert({ ...draft, smartArt: normalizeSmartArt(draft.smartArt) }) }] : []),
    { label: '확인', primary: true, action: () => onCommit({ ...draft, smartArt: normalizeSmartArt(draft.smartArt) }) }, { label: '취소' },
  ] });
}
