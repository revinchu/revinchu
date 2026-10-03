// Mouse navigation is a view gesture: never write cells, styles or undo history.
export function installRibbonMouseDrag({ enabled }) {
  let gesture = null, swallowClick = false;
  const stop = e => { e.preventDefault(); e.stopImmediatePropagation(); };
  const end = () => { gesture = null; document.body.classList.remove('ribbon-mouse-drag'); };
  const replayed = new WeakSet();
  const cancel = () => { if (gesture) swallowClick = true; end(); };
  document.addEventListener('pointerdown', e => {
    if (e.pointerType !== 'mouse') { end(); swallowClick = false; return; }
    if (e.button !== 0) return;
    swallowClick = false;
    if (!enabled() || e.ctrlKey || e.metaKey || e.altKey) return;
    const strip = e.target.closest?.('.ribbon-tabs, .ribbon, .quick-access.below, .sheet-tabs');
    if (!strip || strip.scrollWidth <= strip.clientWidth + 1 || e.target.closest('input,textarea,select,[contenteditable=true]')) return;
    gesture = { strip, id:e.pointerId, x:e.clientX, scroll:strip.scrollLeft, active:false };
  }, true);
  // Sheet tabs switch on mousedown. Delay that action until this is known to be a click.
  document.addEventListener('mousedown', e => {
    if (!gesture || replayed.has(e) || !gesture.strip.contains(e.target)) return;
    gesture.down = e; stop(e);
  }, true);
  document.addEventListener('pointermove', e => {
    const g = gesture;
    if (!g || g.id !== e.pointerId) return;
    if (!enabled() || !g.strip.isConnected || !(e.buttons & 1)) { end(); return; }
    const dx = e.clientX - g.x;
    if (!g.active && Math.abs(dx) < 4) return;
    g.active = true; swallowClick = true;
    document.body.classList.add('ribbon-mouse-drag');
    g.strip.scrollLeft = g.scroll - dx;
    stop(e);
  }, true);
  document.addEventListener('pointerup', e => {
    if (!gesture || gesture.id !== e.pointerId) return;
    const g = gesture;
    if (g.active) stop(e);
    end();
    if (!g.active && g.down?.target.isConnected && enabled()) {
      const down = new MouseEvent('mousedown', { bubbles:true, cancelable:true, view:window,
        button:0, buttons:1, clientX:g.down.clientX, clientY:g.down.clientY,
        ctrlKey:g.down.ctrlKey, shiftKey:g.down.shiftKey, altKey:g.down.altKey, metaKey:g.down.metaKey });
      replayed.add(down); g.down.target.dispatchEvent(down);
    }
  }, true);
  document.addEventListener('click', e => { if (swallowClick) { stop(e); swallowClick = false; } }, true);
  document.addEventListener('pointercancel', cancel, true);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cancel(); else if (!gesture) swallowClick = false; }, true);
  window.addEventListener('blur', cancel);
}

export function installGridMousePan({ view, enabled, context, canStart, onDown, onStart, scroll, zoom }) {
  let gesture = null, blocked = false, swallow = false, recentTouch = -Infinity;
  const stop = e => { e.preventDefault(); e.stopImmediatePropagation(); };
  const clear = () => { gesture = null; document.body.classList.remove('grid-mouse-pan'); };
  const cancel = () => { if (gesture) { blocked = true; swallow = true; clear(); } };
  const valid = g => enabled() && g.context === context() && view.isConnected && (g.pan || g.target.isConnected);
  const forward = g => { if (valid(g)) onDown(g.down); };
  const replayed = new WeakSet();
  document.addEventListener('pointerdown', e => {
    if (e.pointerType === 'touch' || e.pointerType === 'pen') {
      recentTouch = performance.now(); clear(); blocked = false; swallow = false;
    } else if (e.pointerType === 'mouse') recentTouch = -Infinity;
  }, true);
  document.addEventListener('mousedown', e => {
    if (e.sourceCapabilities?.firesTouchEvents || performance.now() - recentTouch < 700 || ![0,2].includes(e.button)) return;
    if (gesture && !valid(gesture)) cancel();
    const bits = e.buttons & 3;
    // A new solitary press also recovers from a release lost on window blur.
    if (blocked && bits !== 3) blocked = false;
    if (blocked) { stop(e); return; }
    if (gesture) {
      if (bits === 3) {
        gesture.pan = true; gesture.x = e.clientX; gesture.y = e.clientY;
        swallow = true; onStart?.(); document.body.classList.add('grid-mouse-pan'); stop(e);
      }
      return;
    }
    swallow = false;
    if (!enabled() || !canStart() || !view.contains(e.target) || bits === 3 || e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;
    if (e.target.closest('input,textarea:not(.idle),select,[contenteditable=true]')) return;
    gesture = { down:e, target:e.target, context:context(), x:e.clientX, y:e.clientY, pan:false, menu:null };
    stop(e);
  }, true);
  document.addEventListener('mousemove', e => {
    if (blocked) { stop(e); return; }
    const g = gesture;
    if (!g) return;
    if (!valid(g)) { cancel(); stop(e); return; }
    if (g.pan) {
      if ((e.buttons & 3) !== 3) { cancel(); stop(e); return; }
      scroll((g.x-e.clientX)/zoom(), (g.y-e.clientY)/zoom());
      g.x = e.clientX; g.y = e.clientY; stop(e); return;
    }
    if (g.down.button === 0 && Math.hypot(e.clientX-g.x,e.clientY-g.y) >= 4) {
      clear(); forward(g); // Existing selection/drawing drag receives this move normally.
      return;
    }
    stop(e);
  }, true);
  document.addEventListener('mouseup', e => {
    if (blocked) { stop(e); if (!(e.buttons & 3)) blocked = false; return; }
    const g = gesture;
    if (!g) return;
    if (!valid(g) || g.pan) {
      swallow = true; blocked = !!(e.buttons & 3); clear(); stop(e); return;
    }
    clear(); forward(g);
    // Some browsers emit contextmenu on press, others after release.
    if (g.menu && g.down.button === 2) {
      const menu = new MouseEvent('contextmenu', { bubbles:true, cancelable:true, view:window, button:2, clientX:e.clientX, clientY:e.clientY });
      replayed.add(menu); g.target.dispatchEvent(menu);
    }
  }, true);
  document.addEventListener('contextmenu', e => {
    if (replayed.has(e)) return;
    if (gesture) { gesture.menu = e; stop(e); }
    else if (swallow || blocked) stop(e);
  }, true);
  document.addEventListener('click', e => { if (swallow || blocked) stop(e); }, true);
  document.addEventListener('auxclick', e => { if (swallow || blocked) stop(e); }, true);
  document.addEventListener('pointercancel', cancel, true);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cancel(); else if (!gesture) swallow = false; }, true);
  window.addEventListener('blur', cancel);
}
