// 색 고르기 메뉴 (PowerPoint 색 상자: 테마 색 · 표준 색 · 최근 색 · 다른 색)
import { S } from './state.js';
import { resolveColor, themePaletteRows, STANDARD_COLORS, SLOT_LABEL, parseColorRef } from './themes.js';
import { el, openMenu, closeMenus } from './ui.js';
import { startEyedrop } from './eyedrop.js';

const recent = [];
const label = (c) => { const p = parseColorRef(c); if (!p) return c; if (p.slot) return `${SLOT_LABEL[p.slot] ?? p.slot}${p.mods.length ? ` (${p.mods.map(([k, v]) => `${k} ${Math.round(v * 100)}%`).join(', ')})` : ''}`; return p.hex; };

/**
 * opts: { none: '채우기 없음' | '윤곽선 없음', auto: '자동', extra: [menu items] }
 * onPick(color|null)
 */
export function colorMenu(anchor, onPick, opts = {}) {
  const theme = S.pres.theme;
  const pick = (c) => { closeMenus(); if (c && !recent.includes(c)) { recent.unshift(c); recent.length = Math.min(recent.length, 10); } onPick(c); };
  const sw = (c) => el('button', { class: 'swatch', title: label(c), style: { background: resolveColor(theme, c) }, onmousedown: (e) => e.preventDefault(), onclick: () => pick(c) });
  const rows = themePaletteRows();
  const node = el('div', { class: 'palette' },
    opts.auto ? el('button', { class: 'menu-item', onclick: () => pick(undefined) }, el('span', { class: 'mi-swatch', style: { background: '#000', width: '14px', height: '14px', display: 'inline-block' } }), opts.auto) : null,
    el('div', { class: 'menu-title' }, '테마 색'),
    el('div', { class: 'palette-row gap' }, rows[0].map(sw)),
    rows.slice(1).map((r) => el('div', { class: 'palette-row' }, r.map(sw))),
    el('div', { class: 'menu-title' }, '표준 색'),
    el('div', { class: 'palette-row' }, STANDARD_COLORS.map(sw)),
    recent.length ? [el('div', { class: 'menu-title' }, '최근에 사용한 색'), el('div', { class: 'palette-row' }, recent.map(sw))] : null);
  const items = [{ node }];
  if (opts.none) items.push({ sep: true }, { label: opts.none, icon: 'borderNone', action: () => pick(null) });
  items.push({ label: '다른 색...', icon: 'theme', action: () => moreColors(pick) });
  items.push({ label: '스포이트', icon: 'eyedropper', action: () => startEyedrop(opts.target ?? 'fill', pick) });
  if ('EyeDropper' in window) items.push({ label: '스포이트 (화면 어디서나)', icon: 'eyedropper', action: async () => { try { const r = await new window.EyeDropper().open(); pick(r.sRGBHex.toUpperCase()); } catch { /* 취소 */ } } });
  for (const x of opts.extra ?? []) items.push(x);
  return openMenu(anchor, items, { minWidth: 230 });
}

function moreColors(pick) {
  const inp = el('input', { type: 'color', value: '#4472C4', style: { position: 'fixed', left: '-100px' } });
  document.body.append(inp);
  inp.addEventListener('change', () => { pick(inp.value.toUpperCase()); inp.remove(); });
  inp.addEventListener('blur', () => setTimeout(() => inp.remove(), 500));
  inp.click();
}

/** 작은 색 단추 (서식 창): 누르면 색 메뉴 */
export function colorButton(value, onPick, opts = {}) {
  const theme = S.pres.theme;
  const b = el('button', { class: 'color-btn-wide', type: 'button' },
    el('i', { style: { background: value ? resolveColor(theme, value) : 'repeating-linear-gradient(45deg,#fff 0 4px,#ddd 4px 8px)' } }), '▾');
  b.addEventListener('click', () => colorMenu(b, onPick, opts));
  return b;
}
