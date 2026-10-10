// 모바일 화면 · 끌어서 스크롤
//   body.mobile: 좁은 화면(760px 이하)이면 자동, [보기 › 모바일 보기]로 켜고 끌 수 있음 (localStorage 'wipoint:mobile' = auto|on|off)
//   끌어서 스크롤: 리본 탭 · 리본 · 빠른 실행 · 슬라이드 목록 · 메뉴 · 창 등 넘치는 영역을 손가락/마우스로 끌어 옮김
//   (단추 위에서 시작해도 6px 넘게 움직이면 스크롤로 보고 그 클릭은 무시)

const NARROW = '(max-width: 760px)';
let mode = 'auto';
try { mode = localStorage.getItem('wipoint:mobile') ?? 'auto'; } catch { /* 저장소 없음 */ }

export const mobileMode = () => mode;
export const isMobile = () => document.body.classList.contains('mobile');

export function applyMobile() {
  const narrow = typeof matchMedia === 'function' && matchMedia(NARROW).matches;
  const on = mode === 'on' || (mode === 'auto' && narrow);
  const was = document.body.classList.contains('mobile');
  document.body.classList.toggle('mobile', on);
  return was !== on;
}
export function setMobileMode(m) {
  mode = m;
  try { localStorage.setItem('wipoint:mobile', m); } catch { /* 없음 */ }
  return applyMobile();
}

/** 넘치는 영역을 끌어서 스크롤 (axis: 'x' | 'y' | 'auto') */
export function dragScroll(box, axis = 'auto', { mobileOnly = false } = {}) {
  if (!box || box.dataset.dragScroll) return;
  box.dataset.dragScroll = axis;
  box.addEventListener('pointerdown', (e) => {
    if (mobileOnly && !isMobile() && e.pointerType !== 'touch') return;
    if (e.button !== 0 || e.target.closest('input, textarea, select, [contenteditable="true"], [contenteditable="plaintext-only"]')) return;
    // 터치는 브라우저가 스크롤함 (손가락) — 마우스 · 펜만 직접 처리, 단 터치도 클릭 억제는 같이
    const x0 = e.clientX; const y0 = e.clientY;
    const sl = box.scrollLeft; const st = box.scrollTop;
    const canX = box.scrollWidth > box.clientWidth + 1;
    const canY = box.scrollHeight > box.clientHeight + 1;
    if (!canX && !canY) return;
    let moved = false;
    const move = (ev) => {
      const dx = ev.clientX - x0; const dy = ev.clientY - y0;
      if (!moved && Math.hypot(dx, dy) < 6) return;
      if (!moved) {
        // 가로만 되는 상자에서 세로로 크게 움직이면 (예: 목록을 바깥으로 끌기) 스크롤하지 않음
        const ax = axis === 'auto' ? (canX && (!canY || Math.abs(dx) > Math.abs(dy)) ? 'x' : 'y') : axis;
        if ((ax === 'x' && (!canX || Math.abs(dy) > Math.abs(dx) * 1.5)) || (ax === 'y' && (!canY || Math.abs(dx) > Math.abs(dy) * 1.5))) { stop(); return; }
        moved = ax;
        box.classList.add('drag-scrolling');
      }
      if (e.pointerType === 'touch') return; // 손가락은 브라우저의 기본 스크롤
      if (moved === 'x') box.scrollLeft = sl - dx; else box.scrollTop = st - dy;
      ev.preventDefault();
    };
    const stop = () => {
      removeEventListener('pointermove', move, true);
      removeEventListener('pointerup', up, true);
      removeEventListener('pointercancel', up, true);
    };
    const up = () => {
      stop();
      box.classList.remove('drag-scrolling');
      if (moved) {
        // 끌기 끝의 click 은 단추를 누른 것이 아님
        const kill = (ev) => { ev.stopPropagation(); ev.preventDefault(); };
        box.addEventListener('click', kill, { capture: true, once: true });
        setTimeout(() => box.removeEventListener('click', kill, { capture: true }), 50);
      }
    };
    addEventListener('pointermove', move, true);
    addEventListener('pointerup', up, true);
    addEventListener('pointercancel', up, true);
  });
  // 마우스 휠: 가로로만 넘치는 상자(리본)는 휠로 가로 이동
  if (axis === 'x') {
    box.addEventListener('wheel', (e) => {
      if (Math.abs(e.deltaY) <= Math.abs(e.deltaX) || box.scrollWidth <= box.clientWidth + 1 || e.ctrlKey) return;
      box.scrollLeft += e.deltaY;
      e.preventDefault();
    }, { passive: false });
  }
}

/** 새로 생기는 메뉴 · 창에도 붙임 */
export function initMobile(roots) {
  applyMobile();
  if (typeof matchMedia === 'function') matchMedia(NARROW).addEventListener?.('change', () => { if (applyMobile()) dispatchEvent(new Event('wipoint-mobile')); });
  for (const [el, axis, opts] of roots) dragScroll(el, axis, opts);
  const watch = (layer, sel, axis) => {
    if (!layer) return;
    new MutationObserver(() => { for (const m of layer.querySelectorAll(sel)) dragScroll(m, axis); }).observe(layer, { childList: true, subtree: true });
  };
  watch(document.getElementById('menuLayer'), '.menu', 'y');
  // 파일(백스테이지) 화면은 body 에 바로 붙음
  new MutationObserver(() => { for (const m of document.querySelectorAll('body > .backstage .backstage-nav')) dragScroll(m, 'auto', { mobileOnly: true }); }).observe(document.body, { childList: true });
  watch(document.getElementById('dialogLayer'), '.dialog-body, .sym-grid, .ic-grid, .tpl-cards, .zoom-pick, .eq-gal', 'auto');
}
