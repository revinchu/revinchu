// 공통 UI: 요소 생성 · 메뉴 · 대화상자 · 알림
import { ICONS } from './icons.js';
import { sanitizeHtml, setSafeHtml } from './safe-html.js';
import { accessKeyFromLabel, accessKeyFromEvent, accessKeyHint, accessKeyAliases, dialogButtonAccessKey, allocateAccessKeys } from './access-keys.js';

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'html') setSafeHtml(node, v);
    else if (k === 'style' && typeof v === 'object') {
      // CSS 사용자 속성(--이름)은 대입이 안 되므로 setProperty
      for (const [sk, sv] of Object.entries(v)) {
        if (sv === undefined || sv === null) continue;
        if (sk.startsWith('--')) node.style.setProperty(sk, sv); else node.style[sk] = sv;
      }
    }
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else if (k === 'accessKey') node.dataset.accessKey = v;
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
      node.prepend(sanitizeHtml(ICONS[node.dataset.icon] ?? '', node.ownerDocument));
      node.dataset.hydrated = '1';
    }
  }
}

// ───────────── 창·메뉴 접근키 ─────────────
let accessKeyHandler = null;
export function setAccessKeyHandler(fn) { accessKeyHandler = fn; }
const accessMemory = new WeakMap(), menuAccessOwners = new WeakMap(), accessScopeOrder = new WeakMap(), accessScopeClose = new WeakMap(), menuToolbars = new WeakMap(), menuToolbarObservers = new WeakMap(), menuAnchors = new WeakMap(), menuMinimumWidths = new WeakMap(), consumedAccessKeys = new Set();
let accessOrder = 0;
let accessMode = false, accessScope = null, accessLayer = null;
/** 공통 메뉴/대화상자 밖의 팝업을 등록. owner는 여는 창/팝업, onClose는 Escape 종료 동작. */
export function registerAccessKeyScope(root, { owner = document.activeElement?.closest('.dialog,[data-access-scope]'), onClose } = {}) {
  root.dataset.accessScope = 'popup';
  menuAccessOwners.set(root, owner === root ? null : owner);
  accessScopeOrder.set(root, ++accessOrder);
  if (onClose) accessScopeClose.set(root, onClose);
  prepareAccessKeys(root);
  return () => {
    if (accessScope === root) endAccessKeys();
    menuAccessOwners.delete(root); accessScopeOrder.delete(root); accessScopeClose.delete(root);
    delete root.dataset.accessScope;
  };
}
const ACCESS_CONTROLS = 'button,input:not([type="hidden"]),select,textarea,a[href],[role="tab"],[role="button"],[role="menuitem"],[role="checkbox"],[role="radio"],[role="option"],[role="listbox"],[tabindex]';
function accessVisible(node, includeClipped = false) {
  if (!node.isConnected || node.matches(':disabled,[aria-disabled="true"]') || node.closest('[hidden],[inert],[aria-hidden="true"]')) return false;
  const rect = node.getBoundingClientRect();
  if (!rect.width || !rect.height || (!includeClipped && (rect.bottom <= 0 || rect.right <= 0 || rect.top >= innerHeight || rect.left >= innerWidth))) return false;
  for (let p = node; p && p !== document.body; p = p.parentElement) {
    const style = getComputedStyle(p);
    if (style.visibility === 'hidden' || style.visibility === 'collapse' || style.display === 'none') return false;
    if (!includeClipped && p !== node && /(auto|scroll|hidden|clip)/.test(style.overflow + style.overflowX + style.overflowY)) {
      const clip = p.getBoundingClientRect();
      if (rect.bottom <= clip.top || rect.top >= clip.bottom || rect.right <= clip.left || rect.left >= clip.right) return false;
    }
  }
  return true;
}
function activeAccessScope() {
  // 리본 키팁 자체의 명령 목록은 app.js의 순차 키 처리에 맡긴다.
  const dialogs = [...document.querySelectorAll('#dialogLayer .dialog')].filter((node) => accessVisible(node));
  const modal = dialogs.filter((d) => !d.closest('.dialog-backdrop')?.classList.contains('modeless')).at(-1);
  const owned = (scope) => {
    const seen = new Set(); let hasModal = !modal;
    for (let node = scope; node; node = menuAccessOwners.get(node)) {
      if (!node.isConnected || seen.has(node)) return false;
      seen.add(node); if (node === modal) hasModal = true;
    }
    return hasModal;
  };
  const menus = [...document.querySelectorAll('#menuLayer > .menu:not(.keytip-command-menu),[data-access-scope="popup"]')].filter((menu) => accessScopeOrder.has(menu) && accessVisible(menu) && owned(menu)).sort((a, b) => accessScopeOrder.get(a) - accessScopeOrder.get(b));
  if (menus.length) return menus.at(-1);
  if (modal) return modal;
  // 찾기 같은 modeless 창은 격자에 초점이 있을 때 리본 Alt 키를 빼앗지 않는다.
  return dialogs.findLast((d) => d.contains(document.activeElement)) ?? null;
}
function associatedAccessLabel(target) {
  const own = target.getAttribute('aria-label') ?? target.textContent?.trim() ?? '';
  if (accessKeyFromLabel(own)) return own;
  const labels = [...(target.labels ?? [])];
  const wrapper = target.closest('label'); if (wrapper) labels.push(wrapper);
  for (const id of (target.getAttribute('aria-labelledby') ?? '').split(/\s+/)) { const label = id && document.getElementById(id); if (label) labels.push(label); }
  const textOf = (label) => {
    // label 안의 select 옵션 텍스트가 '범위시트통합문서'처럼 레이블에 붙지 않게 한다.
    const clone = label.cloneNode(true);
    clone.querySelectorAll('input,select,textarea,button,svg').forEach((node) => node.remove());
    return clone.textContent?.trim() ?? '';
  };
  const texts = labels.map(textOf);
  for (const text of texts) if (accessKeyFromLabel(text)) return text;
  for (const text of [own, ...texts, target.getAttribute('title')]) if (accessKeyHint(text)) return text;
  // 레거시 범위 입력기는 label 대신 span + refInput wrapper 구조를 사용한다.
  for (let at = target, i = 0; at && i < 3; at = at.parentElement, i++) {
    const previous = at.previousElementSibling;
    if (previous && !previous.matches(ACCESS_CONTROLS) && !previous.querySelector(ACCESS_CONTROLS) && (accessKeyFromLabel(previous.textContent) || accessKeyHint(previous.textContent))) return previous.textContent;
    if (at.parentElement?.matches('.dialog,.dialog-body,.menu')) break;
  }
  return texts[0] ?? own;
}
function prepareAccessKeys(scope) {
  const closeHead = scope.querySelector(':scope > .dialog-head button[data-dialog-close-head]');
  if (closeHead) {
    const otherClose = [...scope.querySelectorAll('button')].some((node) => node !== closeHead && node.dataset.accessKey !== 'none' && accessVisible(node) && dialogButtonAccessKey(node.getAttribute('aria-label') ?? node.textContent) === 'd');
    closeHead.dataset.accessKey = otherClose ? 'none' : 'd';
  }
  for (const node of scope.querySelectorAll('[data-access-key="none"]')) {
    node.removeAttribute('aria-keyshortcuts'); delete node.dataset.resolvedAccessKey; delete node.dataset.accessKeySource;
  }
  const toolbar = menuToolbars.get(scope);
  const targets = [...scope.querySelectorAll(ACCESS_CONTROLS), ...(toolbar?.isConnected ? toolbar.querySelectorAll(ACCESS_CONTROLS) : [])].filter((node) => node !== scope && accessVisible(node) && node.dataset.accessKey !== 'none' && !node.closest('.access-key-layer'));
  const details = targets.map((node) => ({ explicit: node.dataset.accessKey, aliases: node.dataset.accessAliases, label: associatedAccessLabel(node), previous: accessMemory.get(node) }));
  // 작은 화면에서 메뉴 아래쪽 항목이 스크롤 밖이어도 그 명시 키를 미니 단추가 빼앗지 않는다.
  const reservations = scope.dataset.contextMenu ? [...scope.querySelectorAll('[data-access-key],[data-access-aliases]')].filter((node) => !targets.includes(node) && (/^[a-z0-9]$/i.test(node.dataset.accessKey ?? '') || accessKeyAliases(node.dataset.accessAliases).length)).map((node) => ({ explicit: node.dataset.accessKey, aliases: node.dataset.accessAliases })) : [];
  const keys = allocateAccessKeys([...details, ...reservations]);
  return targets.map((target, i) => {
    const entry = { target, ...keys[i], aliases: accessKeyAliases(details[i].aliases) };
    if (!entry.key) { target.removeAttribute('aria-keyshortcuts'); delete target.dataset.resolvedAccessKey; delete target.dataset.accessKeySource; return null; }
    accessMemory.set(target, entry.key);
    target.setAttribute('aria-keyshortcuts', [entry.key, ...entry.aliases.filter((key) => key !== entry.key)].map((key) => 'Alt+' + key.toUpperCase()).join(' '));
    target.dataset.resolvedAccessKey = entry.key;
    target.dataset.accessKeySource = entry.automatic ? 'wixel' : details[i].explicit || accessKeyFromLabel(details[i].label) ? 'label' : 'excel';
    return entry;
  }).filter(Boolean);
}
function endAccessKeys() {
  accessMode = false; accessScope = null; accessLayer?.remove(); accessLayer = null;
}
function drawAccessKeys(scope) {
  accessLayer?.remove();
  accessLayer = el('div', { class: 'access-key-layer', 'aria-hidden': 'true' });
  for (const item of prepareAccessKeys(scope)) {
    const r = item.target.getBoundingClientRect();
    const badge = el('span', { class: 'access-key-badge', 'data-key': item.key, title: item.automatic ? 'WIXEL 자동 접근키' : item.target.dataset.accessKeySource === 'excel' ? '한국어 Excel 레이블 접근키' : '지정 접근키', style: { left: `${Math.max(2, Math.min(innerWidth - 24, r.right - 18))}px`, top: `${Math.max(2, Math.min(innerHeight - 22, r.top - 6))}px` } }, item.key.toUpperCase());
    accessLayer.append(badge);
  }
  document.body.append(accessLayer);
}
function consumeAccessKey(event, phase, key = '', target = null) {
  event.preventDefault(); event.stopImmediatePropagation();
  consumedAccessKeys.add(event.code || event.key);
  accessKeyHandler?.(event, { phase, key, target });
}
function activateAccessTarget(target) {
  target.focus();
  if (target.matches('button,a[href],input[type="checkbox"],input[type="radio"],input[type="button"],input[type="submit"],input[type="reset"],[role="button"],[role="tab"],[role="menuitem"],[role="checkbox"],[role="radio"],[role="option"]')) target.click();
  else if (target.matches('input:not([type="checkbox"]):not([type="radio"]),textarea')) target.select?.();
}
window.addEventListener('keydown', (event) => {
  if (event.ctrlKey || event.metaKey || event.getModifierState?.('AltGraph')) { endAccessKeys(); return; }
  const scope = activeAccessScope();
  if (!scope) { endAccessKeys(); return; }
  if (accessScope && accessScope !== scope) endAccessKeys();
  if (event.key === 'Alt') {
    consumeAccessKey(event, accessMode ? 'cancel' : 'mode');
    if (event.repeat) return;
    if (accessMode) endAccessKeys(); else { accessMode = true; accessScope = scope; drawAccessKeys(scope); }
    return;
  }
  if (accessMode && event.key === 'Escape') { consumeAccessKey(event, 'cancel'); endAccessKeys(); return; }
  if (event.key === 'Escape' && accessScopeClose.has(scope)) { consumeAccessKey(event, 'cancel'); endAccessKeys(); accessScopeClose.get(scope)(); return; }
  if (['Tab', 'Enter', ' ', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { endAccessKeys(); return; }
  const editable = event.target instanceof Element && (event.target.matches('input,textarea,select') || event.target.isContentEditable);
  const menuTyping = scope.classList.contains('menu') && !editable && !event.altKey && !event.shiftKey && !event.isComposing && event.keyCode !== 229;
  if (!event.altKey && !accessMode && !menuTyping) return;
  const key = accessKeyFromEvent(event);
  if (!key) return;
  if (scope.dataset.contextMenu) {
    const explicit = [...scope.querySelectorAll(':scope > .menu-item')].find((node) => accessVisible(node, true) && (node.dataset.accessKey?.toLowerCase() === key || accessKeyAliases(node.dataset.accessAliases).includes(key)));
    explicit?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  const matches = prepareAccessKeys(scope).filter((entry) => entry.key === key || entry.aliases.includes(key));
  if (!matches.length || scope.getAttribute('aria-busy') === 'true') { if (event.altKey || accessMode) consumeAccessKey(event, 'mode', key); return; }
  if (event.repeat) { consumeAccessKey(event, 'mode', key); return; }
  if (matches.length > 1) {
    const at = matches.findIndex((entry) => entry.target === document.activeElement), target = matches[(at + 1) % matches.length].target;
    consumeAccessKey(event, 'cycle', key, target); target.focus(); accessMode = true; accessScope = scope; drawAccessKeys(scope); return;
  }
  const target = matches[0].target;
  consumeAccessKey(event, 'activate', key, target); endAccessKeys(); activateAccessTarget(target);
  const next = activeAccessScope();
  if (next && next !== scope && (next.classList.contains('menu') || next.dataset.accessScope === 'popup')) { accessMode = true; accessScope = next; drawAccessKeys(next); }
}, true);
window.addEventListener('keyup', (event) => {
  if (!consumedAccessKeys.delete(event.code || event.key)) return;
  event.preventDefault(); event.stopImmediatePropagation();
}, true);
window.addEventListener('blur', () => { endAccessKeys(); consumedAccessKeys.clear(); });
document.addEventListener('mousedown', endAccessKeys, true);
document.addEventListener('focusin', () => { if (accessMode && activeAccessScope() !== accessScope) endAccessKeys(); }, true);
window.addEventListener('scroll', () => { if (accessMode && accessScope?.isConnected) drawAccessKeys(accessScope); }, true);
window.addEventListener('resize', () => { if (accessMode && accessScope?.isConnected) drawAccessKeys(accessScope); });

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
  endAccessKeys();
  for (const m of openMenus) { menuToolbarObservers.get(m)?.disconnect(); m.remove(); }
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
export function openMenu(anchor, items, { minWidth, scroll, toolbar } = {}) {
  if (anchor instanceof Element && suppress && (suppress.el === anchor || suppress.key === anchorKey(anchor)) && Date.now() - suppress.t < 600) {
    suppress = null;
    return document.createElement('div'); // 호출한 쪽이 style 등을 만져도 안전하게
  }
  const anchorRect = anchor instanceof Element ? anchor.getBoundingClientRect() : null;
  closeMenus();
  menuAnchor = anchor instanceof Element ? anchor : null;
  const position = anchorRect && !anchor.isConnected ? { x: anchorRect.left, y: anchorRect.bottom + 2 } : anchor;
  return buildMenu(position, items, { minWidth, scroll, toolbar });
}

/** 열린 메뉴 안의 단추에서 오른쪽에 하위 메뉴 (앞 메뉴는 그대로 둠) — 피벗 필터의 [레이블 필터 ▸] 등 */
export function closeSubmenus() {
  for (const m of openMenus.filter((x) => Number(x.dataset.level) >= 1)) m.remove();
  openMenus = openMenus.filter((x) => Number(x.dataset.level) < 1);
}
export function openSubmenu(anchorEl, items) {
  for (const m of openMenus.filter((x) => Number(x.dataset.level) >= 1)) m.remove();
  openMenus = openMenus.filter((x) => Number(x.dataset.level) < 1);
  const r = anchorEl.getBoundingClientRect();
  return buildMenu({ x: r.right - 2, y: r.top - 4 }, items, { level: 1, parentItem: anchorEl, focus: document.activeElement === anchorEl });
}

/** 메뉴 하나 (submenu: 오른쪽에 하위 메뉴, swatch: 색 견본, header: 제목 줄) */
function buildMenu(anchor, items, { minWidth, scroll, toolbar, level = 0, parentItem = null, focus = true } = {}) {
  const menu = el('div', { class: 'menu', role: 'menu' });
  menuAccessOwners.set(menu, parentItem ? menuAccessOwners.get(parentItem.closest('.menu')) : (anchor instanceof Element ? anchor.closest('.dialog,[data-access-scope]') : null) ?? document.activeElement?.closest('.dialog,[data-access-scope]'));
  accessScopeOrder.set(menu, ++accessOrder);
  menu.dataset.level = String(level);
  if (toolbar || parentItem?.closest('.menu')?.dataset.contextMenu) menu.dataset.contextMenu = 'true';
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
    const openSub = (focus = false) => {
      closeDeeper();
      const r = btn.getBoundingClientRect();
      buildMenu({ x: r.right - 2, y: r.top - 4 }, it.submenu, { level: level + 1, parentItem: btn, focus });
    };
    const btn = el('button', {
      class: `menu-item${it.submenu ? ' has-sub' : ''}`, role: 'menuitem', disabled: it.disabled, 'aria-haspopup': it.submenu ? 'menu' : null, 'data-access-key': it.accessKey, 'data-access-aliases': it.accessAliases,
      onmousedown: (e) => e.preventDefault(),
      onmouseenter: () => { if (it.submenu) openSub(); else closeDeeper(); },
      onclick: () => { if (it.submenu) { openSub(true); return; } closeMenus(); it.action?.(); },
    },
    el('span', { class: 'mi-icon', html: it.checked ? ICONS.check : (it.icon ? ICONS[it.icon] ?? it.icon : '') }),
    it.swatch !== undefined ? el('i', { class: 'mi-swatch', style: { background: it.swatch ?? 'transparent' } }) : null,
    it.desc ? el('span', { class: 'mi-text' }, el('b', {}, it.label), el('small', {}, it.desc)) : el('span', {}, it.swatch !== undefined ? it.label.replace(/^(■|A) /, '') : it.label),
    it.key ? el('span', { class: 'mi-key' }, it.key) : null,
    it.submenu ? el('span', { class: 'mi-key' }, '▸') : null);
    menu.append(btn);
  }
  document.getElementById('menuLayer').append(menu);
  openMenus.push(menu);
  if (toolbar instanceof HTMLElement && !level) attachMenuToolbar(menu, toolbar, anchor);
  placeMenu(menu, anchor);
  prepareAccessKeys(menu);
  menu.addEventListener('keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === 'Tab' && toolbar?.isConnected) {
      const controls = toolbarControls(toolbar);
      if (controls.length) { e.preventDefault(); e.stopPropagation(); controls[e.shiftKey ? controls.length - 1 : 0].focus(); return; }
    }
    if (e.key === 'Escape' || (e.key === 'ArrowLeft' && level && !/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName))) {
      e.preventDefault(); e.stopPropagation();
      if (level) {
        for (const m of openMenus.filter((x) => Number(x.dataset.level) >= level)) m.remove();
        openMenus = openMenus.filter((x) => Number(x.dataset.level) < level);
        parentItem?.focus();
      } else closeMenus();
      return;
    }
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || e.target.isContentEditable) return;
    const entries = [...menu.querySelectorAll(':scope > .menu-item:not(:disabled)')];
    if (!entries.length) return;
    const at = entries.indexOf(document.activeElement);
    if (['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(e.key)) {
      e.preventDefault(); e.stopPropagation(); closeDeeper();
      const i = e.key === 'Home' ? 0 : e.key === 'End' ? entries.length - 1 : (at + (e.key === 'ArrowDown' ? 1 : -1) + entries.length) % entries.length;
      entries[i].focus();
    } else if (e.key === 'ArrowRight' && document.activeElement?.classList.contains('has-sub')) {
      e.preventDefault(); e.stopPropagation(); document.activeElement.click();
    } else if (e.key === 'Enter' || e.key === ' ') {
      if (at >= 0) { e.preventDefault(); e.stopPropagation(); entries[at].click(); }
    }
  });
  if (focus) (menu.querySelector('input:not(:disabled), select:not(:disabled), textarea:not(:disabled)') ?? menu.querySelector(':scope > .menu-item:not(:disabled)'))?.focus();
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
  const toolbar = menuToolbars.get(menu);
  if (toolbar?.isConnected) {
    const margin = 4, gap = 5;
    toolbar.style.maxWidth = `${Math.max(1, innerWidth - margin * 2)}px`;
    toolbar.style.maxHeight = `${Math.max(28, Math.floor((innerHeight - margin * 2 - gap) * .45))}px`;
    const height = toolbar.offsetHeight;
    // 두 영역이 화면보다 크면 메뉴만 스크롤. 작은 화면에서 위쪽 도구와 항목이 겹치지 않는다.
    menu.style.maxHeight = `${Math.max(24, innerHeight - margin * 2 - gap - height)}px`;
    menu.style.maxWidth = `${Math.max(1, innerWidth - margin * 2)}px`;
    menu.style.minWidth = `${Math.min(menuMinimumWidths.get(menu), Math.max(1, innerWidth - margin * 2))}px`;
    const menuHeight = menu.offsetHeight;
    y = Math.max(margin + height + gap, Math.min(y, innerHeight - margin - menuHeight));
    x = Math.max(margin, Math.min(x, innerWidth - margin - menu.offsetWidth));
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;
    toolbar.style.left = `${Math.max(margin, Math.min(x, innerWidth - margin - toolbar.offsetWidth))}px`;
    toolbar.style.top = `${Math.max(margin, y - height - gap)}px`;
  }
}

function toolbarControls(toolbar) { return [...toolbar.querySelectorAll('button,input,select,textarea,[tabindex]')].filter((node) => accessVisible(node) && node.tabIndex >= 0); }
function attachMenuToolbar(menu, toolbar, anchor) {
  toolbar.classList.add('context-mini-toolbar'); toolbar.dataset.level = '0';
  toolbar.setAttribute('role', 'toolbar');
  if (!toolbar.getAttribute('aria-label')) toolbar.setAttribute('aria-label', '미니 서식 도구 모음');
  menuToolbars.set(menu, toolbar); menuAnchors.set(menu, anchor);
  menuMinimumWidths.set(menu, parseFloat(getComputedStyle(menu).minWidth) || 180);
  document.getElementById('menuLayer').append(toolbar); openMenus.push(toolbar);
  // 외부 코드가 메뉴 DOM만 지우는 경우에도 보조 막대가 홀로 남지 않는다.
  const observer = new MutationObserver(() => { if (!menu.isConnected) { toolbar.remove(); observer.disconnect(); openMenus = openMenus.filter((node) => node.isConnected); } });
  menuToolbarObservers.set(menu, observer);
  observer.observe(menu.parentNode, { childList: true });
  toolbar.addEventListener('mousedown', (event) => { if (event.target.closest('button')) event.preventDefault(); });
  toolbar.addEventListener('keydown', (event) => {
    event.stopPropagation();
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape') { event.preventDefault(); closeMenus(); return; }
    const controls = toolbarControls(toolbar), at = controls.indexOf(document.activeElement);
    const editable = event.target.matches('input,select,textarea') || event.target.isContentEditable;
    if (event.key === 'Tab') {
      event.preventDefault(); const next = at + (event.shiftKey ? -1 : 1);
      if (next >= 0 && next < controls.length) controls[next].focus();
      else { const entries = [...menu.querySelectorAll(':scope > .menu-item:not(:disabled)')]; entries[event.shiftKey ? entries.length - 1 : 0]?.focus(); }
    } else if (!editable && ['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key) && controls.length) {
      event.preventDefault(); controls[event.key === 'Home' ? 0 : event.key === 'End' ? controls.length - 1 : (at + (event.key === 'ArrowRight' ? 1 : -1) + controls.length) % controls.length].focus();
    } else if (!editable && event.key === 'ArrowDown') { event.preventDefault(); menu.querySelector(':scope > .menu-item:not(:disabled)')?.focus(); }
  });
}
window.addEventListener('resize', () => { for (const menu of openMenus) if (menuToolbars.has(menu) && menu.isConnected) placeMenu(menu, menuAnchors.get(menu)); });

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
  const returnFocus = document.activeElement;
  let busy = false;
  const focusable = (root) => [...root.querySelectorAll('input, select, textarea, button, a[href], [tabindex]')]
    .filter((x) => !x.disabled && x.tabIndex >= 0 && x.getClientRects().length);
  const close = () => {
    if (!backdrop.isConnected || busy) return;
    endAccessKeys();
    backdrop.remove();
    onClose?.();
    const remaining = [...layer.querySelectorAll('.dialog')].at(-1);
    if (remaining) {
      if (!remaining.contains(document.activeElement)) (remaining.contains(returnFocus) ? returnFocus : focusable(remaining)[0])?.focus();
    } else dialogCloseHandler?.();
  };
  const content = typeof body === 'string' ? el('div', { html: body }) : body;
  const bodyButtons = [...(content?.matches?.('button') ? [content] : []), ...(content?.querySelectorAll('button') ?? [])];
  for (const button of bodyButtons) {
    const key = dialogButtonAccessKey(button.getAttribute('aria-label') ?? button.textContent);
    if (key && button.dataset.accessKey === undefined) button.dataset.accessKey = key;
  }
  const dialog = el('div', { class: 'dialog', role: 'dialog', 'aria-label': title, 'aria-modal': String(!modeless), tabindex: '-1' },
    el('div', { class: 'dialog-head' }, title, el('button', { title: '닫기', 'aria-label': '닫기', 'data-access-key': 'd', 'data-dialog-close-head': true, onclick: close }, '✕')),
    el('div', { class: 'dialog-body' }, content),
    buttons.length ? el('div', { class: 'dialog-foot' }, buttons.map((b) => el('button', {
      class: `btn${b.primary ? ' primary' : ''}`, 'data-access-key': b.accessKey ?? dialogButtonAccessKey(b.label), 'data-access-aliases': b.accessAliases,
      onclick: () => invoke(b),
    }, b.label))) : null);
  if (width) dialog.style.width = `${width}px`;
  const backdrop = el('div', { class: `dialog-backdrop${modeless ? ' modeless' : ''}` }, dialog);
  const error = el('div', { class: 'warn', role: 'alert', hidden: true });
  dialog.querySelector('.dialog-body').append(error);
  const invoke = (button) => {
    if (busy) return;
    error.hidden = true;
    const fail = (err) => { error.textContent = `작업을 완료하지 못했습니다: ${err?.message ?? String(err)}`; error.hidden = false; };
    try {
      const result = button.action?.();
      if (result && typeof result.then === 'function') {
        busy = true;
        dialog.setAttribute('aria-busy', 'true');
        const controls = [...dialog.querySelectorAll('.dialog-foot button, .dialog-head button')];
        const disabled = controls.map((x) => x.disabled);
        controls.forEach((x) => { x.disabled = true; });
        Promise.resolve(result).then((value) => {
          busy = false;
          if (value !== false) close();
        }, fail).finally(() => {
          busy = false;
          dialog.removeAttribute('aria-busy');
          controls.forEach((x, i) => { x.disabled = disabled[i]; });
          if (dialog.isConnected && !dialog.contains(document.activeElement)) (dialog.querySelector('.btn.primary:not(:disabled)') ?? focusable(dialog)[0] ?? dialog).focus();
        });
      } else if (result !== false) close();
    } catch (err) { fail(err); }
  };
  backdrop.addEventListener('mousedown', (e) => { if (e.target === backdrop && !modeless) close(); });
  if (modeless) dragByHead(dialog);
  dialog.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === 'Escape') { e.preventDefault(); close(); return; }
    if (e.key === 'Tab' && !modeless) {
      const list = focusable(dialog);
      const first = list[0]; const last = list.at(-1);
      if (!first) { e.preventDefault(); dialog.focus(); }
      else if (e.shiftKey && (document.activeElement === first || !list.includes(document.activeElement))) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && (document.activeElement === last || !list.includes(document.activeElement))) { e.preventDefault(); first.focus(); }
    }
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON') {
      e.preventDefault(); // 대화상자가 닫힌 뒤 Enter 가 셀 편집기에 들어가지 않도록
      const primary = buttons.find((b) => b.primary);
      if (primary) invoke(primary);
    }
  });
  layer.append(backdrop);
  const first = focusable(dialog).find((x) => /^(INPUT|SELECT|TEXTAREA)$/.test(x.tagName));
  (first ?? dialog.querySelector('.btn.primary') ?? focusable(dialog)[0] ?? dialog).focus();
  first?.select?.();
  onOpen?.(dialog);
  prepareAccessKeys(dialog);
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
    openDialog({ title, body: el('div', {}, message), onClose: resolve, buttons: [{ label: '확인', primary: true }] });
  });
}

/** 입력 필드 여러 개를 받는 대화상자. fields: [{name, label, type, value, options}] */
export function formDialog(title, fields, onSubmit, { okLabel = '확인', note, onChange, onClose } = {}) {
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
      } else if (f.type === 'combo') {
        // 직접 입력 + 목록에서 고르기
        const id = `dl-${Math.random().toString(36).slice(2, 8)}`;
        input = el('input', { type: 'text', value: f.value ?? '', list: id, 'data-access-key': f.accessKey });
        inputs[f.name] = input;
        return el('label', {}, el('span', {}, f.label), input, el('datalist', { id }, (f.options ?? []).map((o) => el('option', { value: o.value }, o.label))));
      } else {
        input = el('input', { type: f.type || 'text', value: f.value ?? '' });
      }
      if (f.accessKey) input.dataset.accessKey = f.accessKey;
      inputs[f.name] = input;
      return el('label', {}, el('span', {}, f.label), input);
    }));
  const read = () => Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.type === 'checkbox' ? i.checked : i.value]));
  // onChange(inputs): 값이 바뀔 때마다 (다른 칸 보이기 · 목록 바꾸기)
  if (onChange) { body.addEventListener('change', () => onChange(inputs)); onChange(inputs); }
  return openDialog({
    title, body, onClose,
    buttons: [
      { label: okLabel, primary: true, action: () => onSubmit(read()) },
      { label: '취소' },
    ],
  });
}
