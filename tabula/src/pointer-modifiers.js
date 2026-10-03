import { appleTouchDevice } from './keyboard-shortcuts.js';

// Apple touch browsers can omit modifier flags from wheel/compatibility mouse
// events. Keep this fallback local to the current key/primary-pointer gesture.
const tracked = new WeakMap();
const mouseModifiers = new WeakMap();
const explicitModifier = event => !!(event?.ctrlKey || event?.metaKey);
const eventWindow = event => event?.view || event?.target?.ownerDocument?.defaultView || globalThis.window;
const modifierKey = event => {
  if (event.code === 'ControlLeft' || event.code === 'ControlRight' || event.key === 'Control' || event.keyCode === 17) return 'control';
  if (event.code === 'MetaLeft' || event.code === 'MetaRight' || event.key === 'Meta' || event.key === 'OS' || event.keyCode === 91 || event.keyCode === 92) return 'meta';
  return '';
};

export function primaryPointerModifier(event) {
  if (explicitModifier(event)) return true;
  // A mobile mouse-down may be delivered to the grid only after pointer-up.
  // Its modifier belongs to that event, never to a later keyboard gesture.
  if (event && mouseModifiers.has(event)) return mouseModifiers.get(event);
  const win = eventWindow(event), state = win && tracked.get(win);
  return !!(state && appleTouchDevice(win.navigator) && (state.control || state.meta));
}

export function installPointerModifierTracking(win = globalThis.window) {
  if (!win?.addEventListener) return () => {};
  let state = tracked.get(win);
  if (!state) {
    state = { control: false, meta: false, pointer: null, refs: 0, remove: null };
    tracked.set(win, state);
    const clear = () => { state.control = state.meta = false; state.pointer = null; };
    const syncKeys = (event, pressed) => {
      const key = modifierKey(event);
      // Keyboard flags also repair a missed key-up. A released right Control
      // may still report ctrlKey=true while left Control remains held.
      if (typeof event.ctrlKey === 'boolean') state.control = event.ctrlKey;
      if (typeof event.metaKey === 'boolean') state.meta = event.metaKey;
      if (key) state[key] = pressed || !!(key === 'control' ? event.ctrlKey : event.metaKey);
      if ((!pressed && key) || (!state.control && !state.meta)) state.pointer = null;
    };
    const down = event => syncKeys(event, true);
    const up = event => syncKeys(event, false);
    const pointerDown = event => {
      state.pointer = null;
      if (!appleTouchDevice(win.navigator) || event.button !== 0 || event.isPrimary === false || !explicitModifier(event)) return;
      state.pointer = { target: event.target, x: event.clientX, y: event.clientY, id: event.pointerId, at: Date.now() };
    };
    const mouseDown = event => {
      const pointer = state.pointer;
      state.pointer = null;
      const same = pointer && event.button === 0 && event.target === pointer.target && Math.abs(event.clientX - pointer.x) <= 1 && Math.abs(event.clientY - pointer.y) <= 1 && Date.now() - pointer.at < 600;
      mouseModifiers.set(event, explicitModifier(event) || !!(appleTouchDevice(win.navigator) && (state.control || state.meta || same)));
    };
    const pointerEnd = event => { if (!state.pointer || event.pointerId === state.pointer.id) state.pointer = null; };
    const mouseEnd = () => { state.pointer = null; };
    const blur = event => { if (event.target === win) clear(); };
    const hidden = () => { if (win.document?.hidden) clear(); };
    const listeners = [['keydown', down], ['keyup', up], ['pointerdown', pointerDown], ['mousedown', mouseDown], ['pointerup', pointerEnd], ['pointercancel', pointerEnd], ['mouseup', mouseEnd], ['blur', blur], ['pageshow', clear], ['pagehide', clear]];
    for (const [type, fn] of listeners) win.addEventListener(type, fn, true);
    win.document?.addEventListener('visibilitychange', hidden, true);
    state.remove = () => { clear(); for (const [type, fn] of listeners) win.removeEventListener(type, fn, true); win.document?.removeEventListener('visibilitychange', hidden, true); tracked.delete(win); };
  }
  state.refs++;
  let disposed = false;
  return () => { if (disposed) return; disposed = true; if (--state.refs === 0) state.remove(); };
}
