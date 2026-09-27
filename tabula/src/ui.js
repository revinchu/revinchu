// 공통 UI: 요소 생성 · 메뉴 · 대화상자 · 알림
import { ICONS } from './icons.js';

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

export function icon(name) {
  const span = el('span', { class: 'ico', html: ICONS[name] ?? '' });
  span.style.display = 'inline-flex';
  return span;
}

/** data-icon 속성이 있는 요소에 아이콘 채우기 */
export function hydrateIcons(root = document) {
  for (const node of root.querySelectorAll('[data-icon]')) {
    if (!node.dataset.hydrated) {
      node.insertAdjacentHTML('afterbegin', ICONS[node.dataset.icon] ?? '');
      node.dataset.hydrated = '1';
    }
  }
}

// ───────────── 알림 ─────────────
let toastTimer;
export function toast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 1800);
}

// ───────────── 메뉴 ─────────────
let openMenus = [];
let onMenuClose = null;

export function setMenuCloseHandler(fn) { onMenuClose = fn; }
export const isMenuOpen = () => openMenus.length > 0;

export function closeMenus() {
  if (!openMenus.length) return;
  for (const m of openMenus) m.remove();
  openMenus = [];
  onMenuClose?.();
}

/**
 * items: [{label, icon, key, action, disabled, checked}] | {sep:true} | {title} | {node}
 * anchor: HTMLElement(아래쪽에 표시) 또는 {x, y}
 */
export function openMenu(anchor, items, { minWidth } = {}) {
  closeMenus();
  const menu = el('div', { class: 'menu', role: 'menu' });
  if (minWidth) menu.style.minWidth = `${minWidth}px`;
  for (const it of items) {
    if (!it) continue;
    if (it.sep) { menu.append(el('div', { class: 'menu-sep' })); continue; }
    if (it.title) { menu.append(el('div', { class: 'menu-title' }, it.title)); continue; }
    if (it.node) { menu.append(it.node); continue; }
    const btn = el('button', {
      class: 'menu-item', role: 'menuitem', disabled: it.disabled,
      onmousedown: (e) => e.preventDefault(),
      onclick: () => { closeMenus(); it.action?.(); },
    },
    el('span', { class: 'mi-icon', html: it.checked ? ICONS.check : (it.icon ? ICONS[it.icon] ?? it.icon : '') }),
    el('span', {}, it.label),
    it.key ? el('span', { class: 'mi-key' }, it.key) : null);
    menu.append(btn);
  }
  document.getElementById('menuLayer').append(menu);
  placeMenu(menu, anchor);
  openMenus.push(menu);
  return menu;
}

function placeMenu(menu, anchor) {
  let x;
  let y;
  if (anchor instanceof Element) {
    const r = anchor.getBoundingClientRect();
    x = r.left;
    y = r.bottom + 2;
  } else {
    ({ x, y } = anchor);
  }
  const mw = menu.offsetWidth;
  const mh = menu.offsetHeight;
  if (x + mw > innerWidth - 4) x = Math.max(4, innerWidth - mw - 4);
  if (y + mh > innerHeight - 4) y = Math.max(4, innerHeight - mh - 4);
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
}

document.addEventListener('mousedown', (e) => {
  if (openMenus.length && !openMenus.some((m) => m.contains(e.target))) closeMenus();
}, true);

// ───────────── 대화상자 ─────────────
let dialogCloseHandler = null;
export function setDialogCloseHandler(fn) { dialogCloseHandler = fn; }
export const isDialogOpen = () => document.getElementById('dialogLayer').childElementCount > 0;

/**
 * buttons: [{label, primary, action}] — action이 false를 반환하면 닫지 않음
 * 반환: { close, root }
 */
export function openDialog({ title, body, buttons = [], onOpen, width }) {
  const layer = document.getElementById('dialogLayer');
  const close = () => {
    backdrop.remove();
    dialogCloseHandler?.();
  };
  const content = typeof body === 'string' ? el('div', { html: body }) : body;
  const dialog = el('div', { class: 'dialog', role: 'dialog', 'aria-label': title },
    el('div', { class: 'dialog-head' }, title, el('button', { title: '닫기', onclick: close }, '✕')),
    el('div', { class: 'dialog-body' }, content),
    buttons.length ? el('div', { class: 'dialog-foot' }, buttons.map((b) => el('button', {
      class: `btn${b.primary ? ' primary' : ''}`,
      onclick: () => { if (b.action?.() !== false) close(); },
    }, b.label))) : null);
  if (width) dialog.style.width = `${width}px`;
  const backdrop = el('div', { class: 'dialog-backdrop' }, dialog);
  backdrop.addEventListener('mousedown', (e) => { if (e.target === backdrop) close(); });
  dialog.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') close();
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON') {
      const primary = buttons.find((b) => b.primary);
      if (primary && primary.action?.() !== false) close();
    }
  });
  layer.append(backdrop);
  const first = dialog.querySelector('input, select, textarea');
  (first ?? dialog.querySelector('.btn.primary'))?.focus();
  first?.select?.();
  onOpen?.(dialog);
  return { close, root: dialog };
}

export function alertDialog(title, message) {
  return new Promise((resolve) => {
    openDialog({ title, body: el('div', {}, message), buttons: [{ label: '확인', primary: true, action: resolve }] });
  });
}

/** 입력 필드 여러 개를 받는 대화상자. fields: [{name, label, type, value, options}] */
export function formDialog(title, fields, onSubmit, { okLabel = '확인', note } = {}) {
  const inputs = {};
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    note ? el('div', { class: 'muted' }, note) : null,
    fields.map((f) => {
      let input;
      if (f.type === 'select') {
        input = el('select', {}, f.options.map((o) => el('option', { value: o.value, selected: o.value === f.value }, o.label)));
      } else if (f.type === 'checkbox') {
        input = el('input', { type: 'checkbox', checked: !!f.value });
      } else if (f.type === 'textarea') {
        input = el('textarea', {}, f.value ?? '');
      } else {
        input = el('input', { type: f.type || 'text', value: f.value ?? '' });
      }
      inputs[f.name] = input;
      return el('label', {}, el('span', {}, f.label), input);
    }));
  const read = () => Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.type === 'checkbox' ? i.checked : i.value]));
  return openDialog({
    title, body,
    buttons: [
      { label: okLabel, primary: true, action: () => onSubmit(read()) },
      { label: '취소' },
    ],
  });
}
