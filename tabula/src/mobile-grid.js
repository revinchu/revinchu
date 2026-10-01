// Touch-only grid navigation. The state machine has no DOM dependencies.
const distance = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
const midpoint = p => ({ x: (p[0].x + p[1].x) / 2, y: (p[0].y + p[1].y) / 2 });
const sameHit = (a, b) => a?.zone === b?.zone && a?.r === b?.r && a?.c === b?.c;
export class TouchGridGesture {
  constructor(host, { slop = 8, holdMs = 450, doubleMs = 350 } = {}) {
    this.host = host; this.slop = slop; this.holdMs = holdMs; this.doubleMs = doubleMs;
    this.phase = 'idle'; this.lastTap = null;
  }
  valid() { return this.context === this.host.context?.(); }
  start(points, time) {
    if (!points.length) return false;
    if (this.phase === 'idle') {
      const hit = this.host.hit(points[0].x, points[0].y);
      if (!['cell', 'rowHeader', 'colHeader', 'corner'].includes(hit?.zone) || this.host.begin?.() === false) return false;
      this.context = this.host.context?.(); this.origin = { ...points[0] }; this.last = { ...points[0] };
      this.originHit = hit; this.started = time; this.phase = 'pending';
    }
    if (!this.valid()) { this.phase = 'blocked'; this.lastTap = null; return true; }
    if (points.length >= 2) {
      this.phase = 'pinch'; this.lastTap = null;
      this.pinchDistance = Math.max(1, distance(points[0], points[1]));
      this.pinchZoom = this.host.zoom(); this.pinchCenter = midpoint(points);
    }
    return true;
  }
  move(points, time) {
    if (this.phase === 'idle' || !points.length) return false;
    if (!this.valid()) { this.phase = 'blocked'; this.lastTap = null; return true; }
    if (this.phase === 'blocked') return true;
    if (points.length >= 2 && this.phase !== 'pinch') this.start(points, time);
    if (this.phase === 'pinch') {
      if (points.length < 2) { this.phase = 'blocked'; return true; }
      const center = midpoint(points), zoom = Math.max(25, Math.min(400, Math.round(this.pinchZoom * distance(points[0], points[1]) / this.pinchDistance)));
      this.host.pinch?.(zoom, this.pinchCenter, center); this.pinchCenter = center;
      return true;
    }
    const point = points[0];
    if (this.phase === 'pending' && distance(point, this.origin) > this.slop) { this.phase = 'scroll'; this.lastTap = null; }
    if (this.phase === 'pending') this.hold(time);
    if (this.phase === 'scroll') this.host.scroll?.(this.last.x - point.x, this.last.y - point.y);
    else if (this.phase === 'range') this.host.range?.(this.originHit, this.host.hit(point.x, point.y, true));
    this.last = { ...point }; return true;
  }
  hold(time) {
    if (this.phase !== 'pending' || time - this.started < this.holdMs || !this.valid() || this.originHit.zone === 'corner') return false;
    this.phase = 'range'; this.lastTap = null; this.host.range?.(this.originHit, this.originHit); return true;
  }
  end(points, time, cancelled = false) {
    if (this.phase === 'idle') return false;
    if (points.length) { this.phase = 'blocked'; this.lastTap = null; return true; }
    const phase = this.phase, invalid = cancelled || !this.valid();
    if (!invalid && phase === 'pending') {
      const hit = this.originHit, tap = this.lastTap;
      const selectable = this.host.tap?.(hit) !== false;
      if (!selectable) this.lastTap = null;
      else if (hit.zone === 'cell' && tap?.context === this.context && sameHit(tap.hit, hit) && time - tap.time <= this.doubleMs && distance(tap.point, this.origin) <= 24) {
        this.lastTap = null; this.host.edit?.(hit);
      } else this.lastTap = { hit, context: this.context, time, point: this.origin };
    } else this.lastTap = null;
    this.phase = 'idle'; this.host.end?.(phase, invalid); return true;
  }
  cancel() { this.phase = 'idle'; this.lastTap = null; this.host.end?.('cancel', true); }
}

/** Input controls and object UI retain their own native touch/click behavior. */
export function gridTouchIgnored(target) {
  if (target?.id === 'cellEditor' && target.classList?.contains('idle')) return false;
  return !!target?.closest?.('input,textarea,select,button,a,[contenteditable="true"],.obj,.fbtn,.pbtn,.pxbtn,.dv-btn,.olb,.olv,.pv-classic-field,.pv-classic-zone');
}
export function bindGridTouch(element, host) {
  const gesture = new TouchGridGesture(host); let timer = null, consumedAt = -Infinity;
  const now = () => performance.now();
  const points = event => Array.from(event.touches, t => ({ id: t.identifier, x: t.clientX, y: t.clientY }));
  const clear = () => { if (timer !== null) clearTimeout(timer); timer = null; };
  const consume = event => { if (event.cancelable) event.preventDefault(); consumedAt = now(); };
  const start = event => {
    if (gesture.phase === 'idle' && gridTouchIgnored(event.target)) { consumedAt = -Infinity; return; }
    if (!gesture.start(points(event), now())) { consumedAt = -Infinity; return; }
    consume(event); clear();
    if (gesture.phase === 'pending') timer = setTimeout(() => { timer = null; gesture.hold(now()); }, gesture.holdMs);
  };
  const move = event => { if (gesture.move(points(event), now())) { consume(event); if (gesture.phase !== 'pending') clear(); } };
  const finish = event => { clear(); if (gesture.end(points(event), now(), event.type === 'touchcancel')) consume(event); };
  const mouse = event => { if (event.sourceCapabilities?.firesTouchEvents && now() - consumedAt < 800) { event.preventDefault(); event.stopImmediatePropagation(); } };
  const menu = event => { if (gesture.phase !== 'idle' || now() - consumedAt < 800) { event.preventDefault(); event.stopImmediatePropagation(); } };
  const options = { passive: false };
  element.addEventListener('touchstart', start, options); element.addEventListener('touchmove', move, options);
  element.addEventListener('touchend', finish, options); element.addEventListener('touchcancel', finish, options);
  element.addEventListener('mousedown', mouse, true); element.addEventListener('contextmenu', menu, true);
  const blur = () => { clear(); gesture.cancel(); };
  const win = element.ownerDocument?.defaultView; win?.addEventListener('blur', blur);
  return () => {
    blur(); element.removeEventListener('touchstart', start); element.removeEventListener('touchmove', move);
    element.removeEventListener('touchend', finish); element.removeEventListener('touchcancel', finish);
    element.removeEventListener('mousedown', mouse, true); element.removeEventListener('contextmenu', menu, true); win?.removeEventListener('blur', blur);
  };
}
