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

export function installGridMousePan({ view, enabled, context, canStart, onDown, onStart, scroll, zoom, onArmedChange }) {
  let gesture = null, blocked = false, swallow = false, recentTouch = -Infinity, armed = false, nativeDrag = false, recentMenu = null;
  const stop = e => { e.preventDefault(); e.stopImmediatePropagation(); };
  const replayed = new WeakSet();
  // iPhone의 시스템 포인터가 CSS cursor를 표시하지 않아도 이동 상태를 확인할 수 있다.
  const cursor = document.createElement('div'); cursor.id = 'gridPanCursor'; cursor.hidden = true; cursor.setAttribute('aria-hidden', 'true');
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', '0 0 24 24'); svg.setAttribute('fill', 'none'); svg.setAttribute('stroke', 'currentColor'); svg.setAttribute('stroke-width', '1.7'); svg.setAttribute('stroke-linecap', 'round'); svg.setAttribute('stroke-linejoin', 'round');
  const path = document.createElementNS(svg.namespaceURI, 'path');
  path.setAttribute('d', 'M8 13V5a2 2 0 0 1 4 0v6-7a2 2 0 0 1 4 0v7-5a2 2 0 0 1 4 0v8c0 5-3 8-7 8-3 0-4-1-6-4l-4-5a2 2 0 0 1 3-3l2 3');
  svg.append(path); cursor.append(svg); document.body.append(cursor);
  const showCursor = e => {
    const v = window.visualViewport, left = v?.offsetLeft ?? 0, top = v?.offsetTop ?? 0;
    cursor.hidden = false;
    cursor.style.left = Math.max(left + 2, Math.min(e.clientX + 12, left + (v?.width ?? innerWidth) - 30)) + 'px';
    cursor.style.top = Math.max(top + 2, Math.min(e.clientY + 12, top + (v?.height ?? innerHeight) - 30)) + 'px';
  };
  const clear = () => { gesture = null; document.body.classList.remove('grid-mouse-pan'); cursor.hidden = true; };
  const cancel = () => { if (gesture) { blocked = true; swallow = true; clear(); } };
  const valid = g => enabled() && g.context === context() && view.isConnected && canStart() && (g.pan || g.target.isConnected);
  const forward = g => { if (valid(g)) onDown(g.down); };
  const eligible = e => enabled() && canStart() && view.contains(e.target) &&
    !e.target.closest?.('input,textarea:not(.idle),select,[contenteditable=true]') && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey;
  const startPan = (g, e) => {
    if (!g.pan) { g.pan = true; g.x = e.clientX; g.y = e.clientY; swallow = true; onStart?.(); }
    document.body.classList.add('grid-mouse-pan'); showCursor(e);
  };
  const begin = (e, source, manual = false) => {
    if (!eligible(e)) return false;
    gesture = { down:e, target:e.target, context:context(), x:e.clientX, y:e.clientY, pan:false, menu:null, source, manual,
      id:e.pointerId, pointerType:e.pointerType ?? 'mouse' };
    swallow = false; nativeDrag = false;
    if (manual || (e.buttons & 3) === 3) startPan(gesture, e);
    return true;
  };
  const setArmed = value => {
    const held = !!gesture || blocked;
    armed = !!value && enabled(); cancel(); blocked = held; swallow = held; nativeDrag = false;
    document.body.classList.toggle('grid-pan-armed', armed);
    if (armed) onStart?.();
    onArmedChange?.(armed);
  };
  const down = (e, source) => {
    if (![0, 2].includes(e.button)) return;
    if (gesture && !valid(gesture)) cancel();
    const bits = e.buttons & 3;
    if (blocked && bits !== 3) blocked = false;
    if (blocked) { stop(e); return; }
    if (gesture) {
      if (bits === 3) { startPan(gesture, e); stop(e); }
      else if (source === 'mouse') {
        // pointerdown과 뒤따르는 호환 mousedown은 같은 한 번의 눌림이다.
        if (!gesture.pan && e.button === gesture.down.button) { gesture.down = e; gesture.source = source; }
        stop(e);
      }
      return;
    }
    if (nativeDrag && bits === 3) return;
    swallow = false;
    if (begin(e, source, armed && e.button === 0)) {
      // 일반 pointerdown의 기본 동작은 유지하여 호환 mousedown을 받을 수 있게 한다.
      if (source === 'mouse' || gesture.pan) stop(e);
    }
  };
  document.addEventListener('pointerdown', e => {
    if (e.pointerType === 'touch' || e.pointerType === 'pen') {
      recentTouch = performance.now();
      if (gesture?.manual && gesture.id !== e.pointerId) { cancel(); stop(e); return; }
      clear(); blocked = false; swallow = false;
      if (armed && e.isPrimary !== false && begin(e, 'pointer', true)) stop(e);
    } else if (e.pointerType === 'mouse') { recentTouch = -Infinity; down(e, 'pointer'); }
  }, true);
  document.addEventListener('mousedown', e => {
    if (e.sourceCapabilities?.firesTouchEvents || performance.now() - recentTouch < 700) return;
    down(e, 'mouse');
  }, true);
  const move = (e, source) => {
    // 펜의 측면 버튼은 양버튼 마우스가 아니다. 명시적 손 모드만 터치/펜을 받는다.
    if (source === 'pointer' && e.pointerType !== 'mouse' && !(gesture?.manual && gesture.pointerType === e.pointerType)) return;
    if (blocked) { stop(e); return; }
    let g = gesture;
    // Pointer Events는 두 번째 버튼을 pointerdown 대신 pointermove(buttons=3)로 전달한다.
    if (!g && !nativeDrag && (e.buttons & 3) === 3 && source === 'pointer') {
      const menu = recentMenu && recentMenu.context === context() && performance.now() - recentMenu.time < 1200 ? recentMenu.event : null;
      if (begin(eligible(e) ? e : menu ?? e, source)) { g = gesture; startPan(g, e); }
    }
    if (!g) return;
    if (!valid(g)) { cancel(); stop(e); return; }
    if (source === 'pointer' && g.id !== undefined && e.pointerId !== g.id) return;
    if (!g.pan && (e.buttons & 3) === 3) startPan(g, e);
    if (g.pan) {
      const held = g.manual ? g.pointerType !== 'mouse' || !!(e.buttons & 1) : (e.buttons & 3) === 3;
      if (!held) { cancel(); stop(e); return; }
      // 호환 mousemove가 같은 좌표로 따라와도 이동량은 0이므로 두 번 스크롤되지 않는다.
      scroll((g.x - e.clientX) / zoom(), (g.y - e.clientY) / zoom());
      g.x = e.clientX; g.y = e.clientY; showCursor(e); stop(e); return;
    }
    if (source === 'pointer') return; // 일반 셀 선택은 기존 mouse 경로에 한 번만 넘긴다.
    if (g.down.button === 0 && Math.hypot(e.clientX - g.x, e.clientY - g.y) >= 4) {
      clear(); nativeDrag = true; forward(g); return;
    }
    stop(e);
  };
  document.addEventListener('pointermove', e => move(e, 'pointer'), true);
  document.addEventListener('mousemove', e => move(e, 'mouse'), true);
  const up = (e, source) => {
    if (!(e.buttons & 3)) nativeDrag = false;
    if (blocked) { stop(e); if (!(e.buttons & 3)) blocked = false; return; }
    const g = gesture;
    if (!g || (source === 'pointer' && g.id !== undefined && e.pointerId !== g.id)) return;
    if (!valid(g) || g.pan) { swallow = true; blocked = !!(e.buttons & 3); clear(); stop(e); return; }
    if (source === 'pointer' && g.source === 'mouse') return;
    clear(); forward(g);
    if (source === 'pointer') {
      // 호환 mouse 이벤트가 없는 장치에서도 단일 클릭의 드래그 상태를 남기지 않는다.
      const release = new MouseEvent('mouseup', { bubbles:true, cancelable:true, button:e.button, buttons:e.buttons, clientX:e.clientX, clientY:e.clientY });
      replayed.add(release); view.dispatchEvent(release);
    }
    if (g.menu && g.down.button === 2) {
      const menu = new MouseEvent('contextmenu', { bubbles:true, cancelable:true, view:window, button:2, clientX:e.clientX, clientY:e.clientY });
      replayed.add(menu); (g.target.isConnected ? g.target : view).dispatchEvent(menu);
    }
  };
  document.addEventListener('pointerup', e => up(e, 'pointer'), true);
  document.addEventListener('mouseup', e => { if (!replayed.has(e)) up(e, 'mouse'); }, true);
  // AssistiveTouch가 마우스를 터치로 변환하거나 Pointer Events가 없는 경우의 명시적 손 도구.
  const touchPoint = (e, t, buttons) => ({ target:e.target, clientX:t.clientX, clientY:t.clientY, button:0, buttons, pointerType:'touch' });
  document.addEventListener('touchstart', e => {
    if (!armed || !enabled() || !view.contains(e.target)) return;
    if (e.touches.length !== 1) { cancel(); stop(e); return; }
    if (!gesture && begin(touchPoint(e, e.touches[0], 1), 'touch', true)) recentTouch = performance.now();
    if (gesture?.manual) stop(e);
  }, { capture:true, passive:false });
  document.addEventListener('touchmove', e => {
    if (!gesture?.manual || gesture.pointerType !== 'touch') return;
    if (e.touches.length !== 1) { cancel(); stop(e); return; }
    const point = touchPoint(e, e.touches[0], 1); Object.assign(point, { preventDefault:()=>e.preventDefault(), stopImmediatePropagation:()=>e.stopImmediatePropagation() });
    move(point, 'touch');
  }, { capture:true, passive:false });
  document.addEventListener('touchend', e => {
    if (gesture?.manual && gesture.pointerType === 'touch') { swallow = true; clear(); stop(e); }
    if (!e.touches.length) blocked = false;
  }, { capture:true, passive:false });
  document.addEventListener('contextmenu', e => {
    if (replayed.has(e)) return;
    if (eligible(e)) recentMenu = { event:e, context:context(), time:performance.now() };
    if (gesture) { gesture.menu = e; stop(e); }
    else if (swallow || blocked || (armed && view.contains(e.target))) stop(e);
  }, true);
  document.addEventListener('click', e => { if (swallow || blocked) stop(e); }, true);
  document.addEventListener('auxclick', e => { if (swallow || blocked) stop(e); }, true);
  document.addEventListener('pointercancel', cancel, true);
  document.addEventListener('touchcancel', cancel, { capture:true, passive:false });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') { cancel(); if (armed) setArmed(false); } else if (!gesture) swallow = false; }, true);
  window.addEventListener('blur', cancel);
  return { get armed() { return armed; }, toggle:()=>setArmed(!armed), setArmed,
    refresh:()=>{ if (!enabled()) setArmed(false); else if (gesture && !valid(gesture)) cancel(); } };
}
