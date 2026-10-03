// 피벗 필드 목록의 포인터 이동만 담당합니다. 모델 변경은 유효한 드롭 콜백 한 번으로 제한합니다.
export function installPivotFieldDrag(pane, { getItem, canStart, resolveTarget, onDrop, onCancel }) {
  const doc = pane.ownerDocument, win = doc.defaultView;
  const pointer = !!win.PointerEvent, listeners = [], ignoredDown = new WeakSet();
  let session = null, frame = 0, observer = null, suppressed = null, disposed = false, releaseSuppression = null;
  const listen = (node, type, fn, options = true) => { node.addEventListener(type, fn, options); listeners.push(() => node.removeEventListener(type, fn, options)); };
  const stop = event => { if (event.cancelable) event.preventDefault(); event.stopImmediatePropagation(); };
  const valid = current => !disposed && pane.isConnected && current.item.node.isConnected && pane.contains(current.item.node) && canStart(current.item);
  const controls = node => node.closest?.('input,button,select,textarea,a,[contenteditable]:not([contenteditable="false"])');
  const point = event => ({ x: event.clientX, y: event.clientY });
  const matchesSuppression = event => suppressed && win.performance.now() <= suppressed.until &&
    !(event.type === 'click' && event.detail === 0 && !(event.pointerId > 0)) &&
    Math.hypot(event.clientX - suppressed.x, event.clientY - suppressed.y) <= 16 &&
    (!(event.pointerId > 0) || event.pointerId === suppressed.id);
  const remember = current => {
    if (!current.active) return;
    releaseSuppression?.();
    suppressed = { ...current.last, id: current.id, until: win.performance.now() + 650 };
    // onDrop이 패널과 helper를 즉시 교체해도 이전 드래그의 click은 새 체크 상자를 누르지 않습니다.
    const trailing = event => { if (matchesSuppression(event)) stop(event); };
    const nextPress = () => release();
    const timer = win.setTimeout(() => release(), 660);
    const release = () => {
      win.clearTimeout(timer);
      doc.removeEventListener('click', trailing, true); doc.removeEventListener('contextmenu', trailing, true);
      doc.removeEventListener(pointer ? 'pointerdown' : 'mousedown', nextPress, true);
      suppressed = null; if (releaseSuppression === release) releaseSuppression = null;
    };
    releaseSuppression = release;
    doc.addEventListener('click', trailing, true); doc.addEventListener('contextmenu', trailing, true);
    doc.addEventListener(pointer ? 'pointerdown' : 'mousedown', nextPress, true);
  };
  const clean = current => {
    if (frame) win.cancelAnimationFrame(frame); frame = 0;
    observer?.disconnect(); observer = null;
    current.target?.element?.classList.remove('pivot-field-drop-target');
    current.ghost?.remove(); current.marker?.remove();
    doc.body.classList.remove('pivot-field-dragging', 'pivot-grid-dragging');
    session = null;
    if (pointer && current.id !== undefined && pane.hasPointerCapture?.(current.id)) {
      try { pane.releasePointerCapture(current.id); } catch { /* 이미 취소된 포인터 */ }
    }
  };
  const cancel = (reason = 'cancel') => {
    const current = session; if (!current) return;
    remember(current); clean(current); onCancel?.(current.item, reason);
  };
  const viewport = () => {
    const v = win.visualViewport;
    return { left: v?.offsetLeft ?? 0, top: v?.offsetTop ?? 0, width: v?.width ?? win.innerWidth, height: v?.height ?? win.innerHeight };
  };
  const updateTarget = current => {
    const next = resolveTarget(current.item, current.last.x, current.last.y);
    current.target?.element?.classList.remove('pivot-field-drop-target');
    current.target = next?.element?.isConnected ? next : null;
    const v = viewport(), ghost = current.ghost;
    ghost.style.maxWidth = Math.max(40, v.width - 16) + 'px';
    ghost.style.left = Math.max(v.left + 4, Math.min(current.last.x + 12, v.left + v.width - ghost.offsetWidth - 4)) + 'px';
    ghost.style.top = Math.max(v.top + 4, Math.min(current.last.y + 14, v.top + v.height - ghost.offsetHeight - 4)) + 'px';
    ghost.dataset.valid = String(!!current.target);
    if (!current.target) { current.marker.hidden = true; ghost.setAttribute('aria-label', current.item.label); return; }
    const { element, index, label } = current.target;
    element.classList.add('pivot-field-drop-target');
    ghost.setAttribute('aria-label', label ? `${current.item.label}: ${label}` : current.item.label);
    const rect = element.getBoundingClientRect(), rows = [...element.querySelectorAll('.pp-item')];
    const at = Math.max(0, Math.min(rows.length, Number.isFinite(index) ? index : rows.length));
    const nextRow = rows[at]?.getBoundingClientRect(), previous = rows[at - 1]?.getBoundingClientRect();
    const markerY = Number.isFinite(current.target.markerY) ? current.target.markerY : nextRow?.top ?? previous?.bottom ?? rect.top + 3;
    const left = Math.max(v.left + 2, rect.left + 2), right = Math.min(v.left + v.width - 2, rect.right - 2);
    current.marker.hidden = right <= left || markerY < v.top || markerY > v.top + v.height;
    current.marker.style.left = left + 'px'; current.marker.style.top = markerY + 'px'; current.marker.style.width = Math.max(0, right - left) + 'px';
  };
  const edgeSpeed = (value, low, high) => {
    const edge = Math.min(24, (high - low) / 3);
    if (value < low || value > high) return 0;
    if (value < low + edge) return -Math.ceil(12 * (low + edge - value) / edge);
    if (value > high - edge) return Math.ceil(12 * (value - high + edge) / edge);
    return 0;
  };
  const autoScroll = current => {
    const { x, y } = current.last, outer = pane.getBoundingClientRect();
    if (x < outer.left || x > outer.right || y < outer.top || y > outer.bottom) return;
    const hit = doc.elementFromPoint(x, y), inner = hit?.closest?.('.pp-area,.pp-fields');
    const boxes = inner && pane.contains(inner) ? [inner, pane] : [pane];
    for (const box of boxes) {
      const r = box.getBoundingClientRect(), dx = box.scrollWidth > box.clientWidth + 1 ? edgeSpeed(x, r.left, r.right) : 0;
      const dy = box.scrollHeight > box.clientHeight + 1 ? edgeSpeed(y, r.top, r.bottom) : 0;
      const oldX = box.scrollLeft, oldY = box.scrollTop;
      if (dx) box.scrollLeft += dx; if (dy) box.scrollTop += dy;
      if (oldX !== box.scrollLeft || oldY !== box.scrollTop) break;
    }
  };
  const tick = () => {
    frame = 0; const current = session; if (!current?.active) return;
    if (!valid(current)) { cancel('context'); return; }
    autoScroll(current); updateTarget(current);
    frame = win.requestAnimationFrame(tick);
  };
  const activate = current => {
    current.active = true;
    const ghost = doc.createElement('div'), marker = doc.createElement('div');
    ghost.className = 'pivot-field-ghost'; ghost.textContent = current.item.label;
    marker.className = 'pivot-field-marker'; marker.hidden = true;
    for (const node of [ghost, marker]) { node.style.position = 'fixed'; node.style.pointerEvents = 'none'; node.setAttribute('aria-hidden', 'true'); }
    doc.body.append(ghost, marker); current.ghost = ghost; current.marker = marker;
    doc.body.classList.add('pivot-field-dragging', 'pivot-grid-dragging');
    if (pointer && current.id !== undefined) { try { pane.setPointerCapture(current.id); } catch { /* 합성 입력 또는 이미 해제된 포인터 */ } }
    updateTarget(current); frame = win.requestAnimationFrame(tick);
  };
  const down = event => {
    if (disposed || session || ignoredDown.has(event) || event.defaultPrevented || event.button !== 0 || event.isPrimary === false || controls(event.target)) return;
    const item = getItem(event.target);
    if (!item?.node || !pane.contains(item.node) || !canStart(item)) return;
    suppressed = null;
    session = { item, id: pointer ? event.pointerId : undefined, start: point(event), last: point(event), active: false, target: null };
    observer = new win.MutationObserver(() => { if (session && !valid(session)) cancel('node'); });
    observer.observe(doc.body, { childList: true, subtree: true });
    // 5px 전에는 클릭/체크 상자를 그대로 둡니다. 터치 스크롤은 행의 touch-action CSS로 제어합니다.
  };
  const move = event => {
    const current = session; if (!current || pointer && event.pointerId !== current.id) return;
    if (!valid(current)) { cancel('context'); return; }
    if (event.pointerType !== 'touch' && (event.buttons & 1) !== 1 || event.buttons & 2) { cancel('buttons'); return; }
    current.last = point(event);
    if (!current.active && Math.hypot(current.last.x - current.start.x, current.last.y - current.start.y) < 5) return;
    if (!current.active) activate(current); else updateTarget(current);
    stop(event);
  };
  const up = event => {
    const current = session; if (!current || pointer && event.pointerId !== current.id) return;
    if (!current.active) { clean(current); return; }
    stop(event); current.last = point(event);
    if (event.button !== 0) { cancel('buttons'); return; }
    if (!valid(current)) { cancel('context'); return; }
    updateTarget(current); const target = current.target;
    remember(current); clean(current);
    if (target) onDrop(current.item, target); else onCancel?.(current.item, 'outside');
  };
  listen(pane, pointer ? 'pointerdown' : 'mousedown', down);
  listen(doc, pointer ? 'pointermove' : 'mousemove', move, { capture: true, passive: false });
  listen(doc, pointer ? 'pointerup' : 'mouseup', up);
  if (pointer) {
    listen(doc, 'pointerdown', event => { if (session && event.pointerId !== session.id) { ignoredDown.add(event); cancel('pointer'); } });
    listen(doc, 'pointercancel', event => { if (session && event.pointerId === session.id) cancel('pointercancel'); });
    listen(pane, 'lostpointercapture', event => { if (event.target === pane && session?.active && event.pointerId === session.id && !pane.hasPointerCapture?.(event.pointerId)) cancel('capture'); });
  }
  listen(pane, 'dragstart', event => { if (session) stop(event); });
  listen(doc, 'keydown', event => { if (session && event.key === 'Escape') { if (session.active) stop(event); cancel('escape'); } });
  listen(win, 'blur', event => { if (event.target === win) cancel('blur'); });
  const suppress = event => {
    if (session?.active) { stop(event); return; }
    if (matchesSuppression(event)) stop(event);
  };
  listen(doc, 'click', suppress); listen(doc, 'contextmenu', suppress);
  return { cancel, dispose() { if (disposed) return; cancel('dispose'); disposed = true; listeners.forEach(remove => remove()); } };
}
