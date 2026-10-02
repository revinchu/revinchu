// Device-only layout preferences; this module never changes workbook data.
export const MOBILE_MODE_KEY = 'wixel.mobile-work.v1';
export const MOBILE_DENSITY_KEY = 'wixel.mobile-density.v1';
export function mobileDensity(value) { return value === 'comfortable' ? 'comfortable' : 'compact'; }
export function mobileLayout({ width, height, layoutHeight = height, coarse = false, preference = 'auto', typing = false }) {
  const active = preference === 'on' || (preference !== 'off' && (width <= 720 || (coarse && width <= 1100)));
  const keyboard = active && typing && layoutHeight - height > 120;
  return { active, keyboard, short: active && height <= 600, width, height };
}
export function mobileSheetZoom(width, saved = 100) {
  const target = width < 360 ? 75 : width < 480 ? 85 : width < 720 ? 90 : 100;
  return Math.max(25, Math.min(400, Math.round(Math.min(Number(saved) || 100, target))));
}
export function installMobileWork({ button, onChange }) {
  let preference = 'auto', density = 'compact', state = null, frame = 0;
  try { const saved = localStorage.getItem(MOBILE_MODE_KEY); if (['auto', 'on', 'off'].includes(saved)) preference = saved; } catch {}
  try { density = mobileDensity(localStorage.getItem(MOBILE_DENSITY_KEY)); } catch {}
  const viewport = window.visualViewport;
  const media = window.matchMedia('(pointer: coarse)');
  const refresh = () => {
    frame = 0;
    const width = viewport?.width ?? window.innerWidth, height = viewport?.height ?? window.innerHeight;
    const focus = document.activeElement;
    const typing = !!focus?.matches('input:not([type=checkbox]):not([type=radio]):not([type=range]), textarea:not(.idle), [contenteditable=true]');
    const next = mobileLayout({ width: window.innerWidth, height, layoutHeight: window.innerHeight, coarse: media.matches, preference, typing });
    next.density = density;
    const root = document.documentElement;
    root.style.setProperty('--mobile-vw', `${width}px`); root.style.setProperty('--mobile-vh', `${height}px`);
    root.style.setProperty('--mobile-left', `${viewport?.offsetLeft ?? 0}px`); root.style.setProperty('--mobile-top', `${viewport?.offsetTop ?? 0}px`);
    document.body.classList.toggle('mobile-work-mode', next.active);
    document.body.classList.toggle('mobile-short', next.short);
    document.body.classList.toggle('mobile-compact', next.active && density === 'compact');
    document.body.classList.toggle('mobile-keyboard', next.keyboard);
    button?.setAttribute('aria-pressed', String(next.active));
    if (button) button.title = next.active ? '모바일 작업 모드 끄기' : '모바일 작업 모드 · 화면 맞춤';
    const previous = state; state = next;
    onChange?.(next, previous);
  };
  const schedule = () => { if (!frame) frame = requestAnimationFrame(refresh); };
  window.addEventListener('resize', schedule); window.addEventListener('orientationchange', schedule);
  viewport?.addEventListener('resize', schedule); viewport?.addEventListener('scroll', schedule);
  document.addEventListener('focusin', schedule); document.addEventListener('focusout', schedule);
  media.addEventListener?.('change', schedule);
  const setPreference = value => {
    preference = ['auto', 'on', 'off'].includes(value) ? value : 'auto';
    try { localStorage.setItem(MOBILE_MODE_KEY, preference); } catch {}
    refresh(); window.dispatchEvent(new Event('resize'));
  };
  const setDensity = value => {
    density = mobileDensity(value);
    try { localStorage.setItem(MOBILE_DENSITY_KEY, density); } catch {}
    refresh(); window.dispatchEvent(new Event('resize'));
  };
  refresh();
  return { get density() { return density; }, setDensity, get active() { return !!state?.active; }, get preference() { return preference; },
    toggle: () => setPreference(state?.active ? 'off' : 'on'), setPreference, refresh };
}
