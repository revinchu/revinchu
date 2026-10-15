// 키 팁 (PowerPoint 의 Alt 키 탐색): Alt 또는 F10 → 탭 · 빠른 실행 도구 모음에 글자 표시 → 글자를 누르면 그 탭의 모든 명령에 글자 표시
// Alt+H, Alt+N … 처럼 바로 눌러도 되고, Alt+1 … Alt+9 는 빠른 실행 도구 모음 명령. 메뉴가 열리면 메뉴 항목에도 글자, 화살표 · Enter 로 이동.
import { isDialogOpen, isMenuOpen, closeMenus, closeSubmenus } from './ui.js';

let layer = null;
let mode = null;          // null | 'top' | 'tab' | 'menu' | 'backstage'
let buf = '';
let altDown = false;
let altUsed = false;
let blocked = () => false;

const visible = (e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && r.right > 0 && r.left < innerWidth; };
const topMenu = () => { const ms = [...document.querySelectorAll('#menuLayer .menu')]; return ms[ms.length - 1] ?? null; };

/** 지금 단계에서 글자를 받을 요소들: [{ el, kt }] */
function candidates() {
  if (mode === 'top') {
    return [...document.querySelectorAll('#ribbonTabs [data-kt], #qat [data-kt], #qatTop [data-kt]')].filter(visible).map((el) => ({ el, kt: el.dataset.kt }));
  }
  if (mode === 'tab') return [...document.querySelectorAll('#ribbon [data-kt]')].map((el) => ({ el, kt: el.dataset.kt }));
  if (mode === 'backstage') return [...document.querySelectorAll('.backstage [data-kt]')].filter(visible).map((el) => ({ el, kt: el.dataset.kt }));
  if (mode === 'menu') {
    const m = topMenu();
    const keys = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ123456789';
    const items = [...(m?.querySelectorAll('.menu-item:not(:disabled), .swatch') ?? [])];
    // 26 + 9 개가 넘으면 두 글자
    if (items.length <= keys.length) return items.map((el, i) => ({ el, kt: keys[i] }));
    return items.map((el, i) => ({ el, kt: keys[Math.floor(i / 26) % 26] + keys[i % 26] }));
  }
  return [];
}

function draw() {
  if (!layer) layer = document.getElementById('keytipLayer');
  layer.innerHTML = '';
  if (!mode) return;
  for (const { el, kt } of candidates()) {
    if (!kt || (buf && !kt.startsWith(buf))) continue;
    if (mode === 'tab' && !visible(el)) continue;
    const r = el.getBoundingClientRect();
    const b = document.createElement('span');
    b.className = 'keytip';
    b.textContent = kt;
    const big = r.height > 40;
    let x = r.left + r.width / 2;
    let y = big ? r.bottom - 12 : r.bottom - 7;
    if (mode === 'menu') { x = r.left + 12; y = r.top + r.height / 2 - 7; }
    if (mode === 'top') y = r.bottom - 10;
    b.style.left = `${Math.max(10, Math.min(innerWidth - 10, x))}px`;
    b.style.top = `${Math.max(0, Math.min(innerHeight - 16, y))}px`;
    layer.append(b);
  }
}

export function keytipsActive() { return !!mode; }
export function exitKeytips() { mode = null; buf = ''; draw(); }
function enter(m) { mode = m; buf = ''; draw(); }

function afterAction() {
  // 메뉴가 열렸으면 메뉴 항목 글자, 아니면 끝 (다음 키가 곧바로 와도 맞는 단계에서 받도록 바로 정함)
  const decide = () => {
    if (isMenuOpen() && topMenu()) enter('menu');
    else if (document.querySelector('.backstage')) enter('backstage');
    else exitKeytips();
  };
  decide();
  requestAnimationFrame(() => { if (mode) decide(); });
}

function activate(c) {
  const el = c.el;
  buf = '';
  if (mode === 'top') {
    if (el.dataset.tab) {
      el.click();
      enter('tab');
      scrollToRibbonStart();
      requestAnimationFrame(draw);
      return;
    }
    el.click();
    afterAction();
    return;
  }
  if (mode === 'tab') {
    el.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    if (el.ktAction) el.ktAction(); else el.click();
    const ae = document.activeElement;
    if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'SELECT') && el.contains(ae)) { exitKeytips(); return; }
    afterAction();
    return;
  }
  if (mode === 'menu') {
    el.click();
    afterAction();
    return;
  }
  if (mode === 'backstage') { el.click(); afterAction(); }
}
function scrollToRibbonStart() { const r = document.getElementById('ribbon'); if (r) r.scrollLeft = 0; draw(); }

function typeChar(ch) {
  const next = buf + ch;
  const list = candidates().filter((c) => c.kt && c.kt.startsWith(next));
  if (!list.length) return; // 맞는 키 팁 없음: 무시 (Office 처럼 그대로 기다림)
  const exact = list.find((c) => c.kt === next);
  if (exact) { activate(exact); return; }
  buf = next;
  draw();
}

function back() {
  if (buf) { buf = ''; draw(); return; }
  if (mode === 'menu') {
    const m = topMenu();
    if (m && Number(m.dataset.level) >= 1) { closeSubmenus(); draw(); return; }
    closeMenus();
    enter(document.querySelector('.backstage') ? 'backstage' : 'tab');
    return;
  }
  if (mode === 'tab') { enter('top'); return; }
  exitKeytips();
}

const charOf = (e) => {
  if (/^Key[A-Z]$/.test(e.code)) return e.code.slice(3);
  if (/^Digit\d$/.test(e.code)) return e.code.slice(5);
  if (/^Numpad\d$/.test(e.code)) return e.code.slice(6);
  return null;
};

/** 메뉴 안 화살표 · Enter (키 팁 모드가 아니어도) */
function menuNav(e) {
  const m = topMenu();
  if (!m) return false;
  const items = [...m.querySelectorAll('.menu-item:not(:disabled)')];
  if (!items.length) return false;
  const i = items.indexOf(document.activeElement);
  const go = (j) => { const t = items[(j + items.length) % items.length]; t.focus(); t.scrollIntoView?.({ block: 'nearest' }); };
  switch (e.key) {
    case 'ArrowDown': go(i + 1); return true;
    case 'ArrowUp': go(i < 0 ? items.length - 1 : i - 1); return true;
    case 'Home': go(0); return true;
    case 'End': go(items.length - 1); return true;
    case 'ArrowRight': if (i >= 0 && items[i].classList.contains('has-sub')) { items[i].click(); requestAnimationFrame(() => { const n = topMenu(); n?.querySelector('.menu-item:not(:disabled)')?.focus(); if (mode) draw(); }); return true; } return false;
    case 'ArrowLeft': if (Number(m.dataset.level) >= 1) { closeSubmenus(); if (mode) draw(); return true; } return false;
    case 'Enter': case ' ': if (i >= 0) { items[i].click(); if (mode) afterAction(); return true; } return false;
    default: {
      // PowerPoint 오른쪽 클릭 메뉴의 괄호 글자 (잘라내기(T) …)
      if (mode || e.key.length !== 1) return false;
      const k = e.key.toUpperCase();
      const hit = items.find((x) => x.dataset.acc === k);
      if (!hit) return false;
      hit.click();
      if (hit.classList.contains('has-sub')) requestAnimationFrame(() => topMenu()?.querySelector('.menu-item:not(:disabled)')?.focus());
      return true;
    }
  }
}

function onKeyDown(e) {
  if (blocked() || isDialogOpen()) { if (mode) exitKeytips(); return; }
  if (e.key === 'Alt') { altDown = true; altUsed = false; return; }
  if (altDown) altUsed = true;
  // 메뉴 화살표 이동
  if (isMenuOpen() && !e.ctrlKey && !e.metaKey && !e.altKey && menuNav(e)) { e.preventDefault(); e.stopPropagation(); return; }
  if (e.key === 'F10' && !e.shiftKey && !e.ctrlKey && !e.altKey) { e.preventDefault(); e.stopPropagation(); if (mode) exitKeytips(); else enter(document.querySelector('.backstage') ? 'backstage' : 'top'); return; }
  const ch = charOf(e);
  if (mode) {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); back(); return; }
    if (ch && !e.ctrlKey && !e.metaKey) { e.preventDefault(); e.stopPropagation(); typeChar(ch); return; }
    if (e.key === 'Tab' || e.key.startsWith('Arrow') || e.key === 'Enter') { exitKeytips(); return; }
    if (!['Shift', 'Control', 'Meta', 'CapsLock'].includes(e.key)) exitKeytips();
    return;
  }
  // Alt+글자 / Alt+숫자 바로 누르기 (Ctrl·Shift 없이)
  if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && ch) {
    e.preventDefault();
    e.stopPropagation();
    enter(document.querySelector('.backstage') ? 'backstage' : 'top');
    typeChar(ch);
  }
}
function onKeyUp(e) {
  if (e.key !== 'Alt') return;
  const solo = altDown && !altUsed;
  altDown = false;
  if (!solo || blocked() || isDialogOpen()) return;
  e.preventDefault();
  if (mode) exitKeytips();
  else enter(document.querySelector('.backstage') ? 'backstage' : 'top');
}

export function initKeytips({ isBlocked } = {}) {
  if (isBlocked) blocked = isBlocked;
  addEventListener('keydown', onKeyDown, true);
  addEventListener('keyup', onKeyUp, true);
  addEventListener('pointerdown', () => { if (mode) exitKeytips(); }, true);
  addEventListener('resize', () => { if (mode) draw(); });
  addEventListener('blur', () => { altDown = false; if (mode) exitKeytips(); });
  document.getElementById('ribbon')?.addEventListener('scroll', () => { if (mode === 'tab') draw(); });
}
/** 리본이 다시 그려진 뒤 (탭 전환 등) 글자 위치 갱신 */
export function redrawKeytips() { if (mode) draw(); }
