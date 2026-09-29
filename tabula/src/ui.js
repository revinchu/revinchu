// 공통 UI: 요소 생성 · 메뉴 · 대화상자 · 알림
import { ICONS } from './icons.js';

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k === 'style' && typeof v === 'object') {
      // CSS 사용자 속성(--이름)은 대입이 안 되므로 setProperty
      for (const [sk, sv] of Object.entries(v)) {
        if (sv === undefined || sv === null) continue;
        if (sk.startsWith('--')) node.style.setProperty(sk, sv); else node.style[sk] = sv;
      }
    }
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat(Infinity)) {
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
  menuAnchor = null;
  onMenuClose?.();
}

/**
 * items: [{label, icon, key, action, disabled, checked}] | {sep:true} | {title} | {node}
 * anchor: HTMLElement(아래쪽에 표시) 또는 {x, y}
 */
// 같은 단추를 다시 누르면 메뉴를 닫음 (열 때 단추를 기억, mousedown 에서 닫힌 직후의 click 은 다시 열지 않음)
let menuAnchor = null;
let suppress = null;
// 화면을 다시 그려 단추 요소가 바뀌어도 같은 단추로 알아보도록 (종류 · 데이터 · 제목)
const anchorKey = (el) => `${el.tagName}|${String(el.className).replace(/\b(on|active|open|pressed)\b/g, '').trim()}|${JSON.stringify({ ...el.dataset })}|${el.getAttribute('title') ?? ''}`;
export function openMenu(anchor, items, { minWidth, scroll } = {}) {
  if (anchor instanceof Element && suppress && (suppress.el === anchor || suppress.key === anchorKey(anchor)) && Date.now() - suppress.t < 600) {
    suppress = null;
    return document.createElement('div'); // 호출한 쪽이 style 등을 만져도 안전하게
  }
  closeMenus();
  menuAnchor = anchor instanceof Element ? anchor : null;
  return buildMenu(anchor, items, { minWidth, scroll });
}

/** 메뉴 하나 (submenu: 오른쪽에 하위 메뉴, swatch: 색 견본, header: 제목 줄) */
function buildMenu(anchor, items, { minWidth, scroll, level = 0 } = {}) {
  const menu = el('div', { class: 'menu', role: 'menu' });
  menu.dataset.level = String(level);
  if (minWidth) menu.style.minWidth = `${minWidth}px`;
  if (scroll || level) { menu.style.maxHeight = '60vh'; menu.style.overflowY = 'auto'; }
  const closeDeeper = () => {
    for (const m of openMenus.filter((x) => Number(x.dataset.level) > level)) m.remove();
    openMenus = openMenus.filter((x) => Number(x.dataset.level) <= level);
  };
  for (const it of items) {
    if (!it) continue;
    if (it.sep) { menu.append(el('div', { class: 'menu-sep' })); continue; }
    if (it.title || it.header) { menu.append(el('div', { class: 'menu-title' }, it.title ?? it.label)); continue; }
    if (it.node) { menu.append(it.node); continue; }
    const openSub = () => {
      closeDeeper();
      const r = btn.getBoundingClientRect();
      buildMenu({ x: r.right - 2, y: r.top - 4 }, it.submenu, { level: level + 1 });
    };
    const btn = el('button', {
      class: `menu-item${it.submenu ? ' has-sub' : ''}`, role: 'menuitem', disabled: it.disabled,
      onmousedown: (e) => e.preventDefault(),
      onmouseenter: () => { if (it.submenu) openSub(); else closeDeeper(); },
      onclick: () => { if (it.submenu) { openSub(); return; } closeMenus(); it.action?.(); },
    },
    el('span', { class: 'mi-icon', html: it.checked ? ICONS.check : (it.icon ? ICONS[it.icon] ?? it.icon : '') }),
    it.swatch !== undefined ? el('i', { class: 'mi-swatch', style: { background: it.swatch ?? 'transparent' } }) : null,
    el('span', {}, it.swatch !== undefined ? it.label.replace(/^(■|A) /, '') : it.label),
    it.key ? el('span', { class: 'mi-key' }, it.key) : null,
    it.submenu ? el('span', { class: 'mi-key' }, '▸') : null);
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
  if (openMenus.length && !openMenus.some((m) => m.contains(e.target))) {
    // 메뉴를 연 단추를 다시 누름 → 닫기만 (바로 이어지는 click 으로 다시 열리지 않게)
    const a = menuAnchor;
    if (a) {
      const key = anchorKey(a);
      const hit = e.target.closest?.('button, [data-c], [data-k], .fbtn, .dv-btn, .rbtn, .pp-drop, .pp-menu');
      if ((a.isConnected && a.contains(e.target)) || (hit && anchorKey(hit) === key)) suppress = { el: a, key, t: Date.now() };
    }
    closeMenus();
  }
}, true);

// ───────────── 대화상자 ─────────────
let dialogCloseHandler = null;
export function setDialogCloseHandler(fn) { dialogCloseHandler = fn; }
// 창을 띄워 둔 채 시트를 쓸 수 있는 대화상자(찾기 및 바꾸기 등)는 열린 것으로 치지 않음
export const isDialogOpen = () => [...document.getElementById('dialogLayer').children].some((x) => !x.classList.contains('modeless'));

/**
 * buttons: [{label, primary, action}] — action이 false를 반환하면 닫지 않음
 * 반환: { close, root }
 */
export function openDialog({ title, body, buttons = [], onOpen, width, modeless = false, onClose }) {
  const layer = document.getElementById('dialogLayer');
  const close = () => {
    if (!backdrop.isConnected) return;
    backdrop.remove();
    onClose?.();
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
  const backdrop = el('div', { class: `dialog-backdrop${modeless ? ' modeless' : ''}` }, dialog);
  backdrop.addEventListener('mousedown', (e) => { if (e.target === backdrop && !modeless) close(); });
  if (modeless) dragByHead(dialog);
  dialog.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') close();
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON') {
      e.preventDefault(); // 대화상자가 닫힌 뒤 Enter 가 셀 편집기에 들어가지 않도록
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

/** 대화상자 머리글을 끌어 옮기기 */
function dragByHead(dialog) {
  const head = dialog.querySelector('.dialog-head');
  head.style.cursor = 'move';
  head.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'BUTTON') return;
    const r = dialog.getBoundingClientRect();
    const dx = e.clientX - r.left;
    const dy = e.clientY - r.top;
    const move = (ev) => {
      dialog.style.position = 'fixed';
      dialog.style.margin = '0';
      dialog.style.left = `${Math.max(0, Math.min(innerWidth - 60, ev.clientX - dx))}px`;
      dialog.style.top = `${Math.max(0, Math.min(innerHeight - 30, ev.clientY - dy))}px`;
    };
    const up = () => { removeEventListener('mousemove', move); removeEventListener('mouseup', up); };
    addEventListener('mousemove', move);
    addEventListener('mouseup', up);
    e.preventDefault();
  });
}

export function alertDialog(title, message) {
  return new Promise((resolve) => {
    openDialog({ title, body: el('div', {}, message), buttons: [{ label: '확인', primary: true, action: resolve }] });
  });
}

/** 입력 필드 여러 개를 받는 대화상자. fields: [{name, label, type, value, options}] */
export function formDialog(title, fields, onSubmit, { okLabel = '확인', note, onChange } = {}) {
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
  // onChange(inputs): 값이 바뀔 때마다 (다른 칸 보이기 · 목록 바꾸기)
  if (onChange) { body.addEventListener('change', () => onChange(inputs)); onChange(inputs); }
  return openDialog({
    title, body,
    buttons: [
      { label: okLabel, primary: true, action: () => onSubmit(read()) },
      { label: '취소' },
    ],
  });
}
