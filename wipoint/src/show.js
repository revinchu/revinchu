// 슬라이드 쇼: 전환 효과 · 클릭 단계별 애니메이션 · 발표자 보기 · 펜/레이저 · 예행 연습
import { S, change, emit } from './state.js';
import { slideHtml, SLIDE_CSS, escHtml } from './render.js';
import { animSteps, stepTimeline, slideTitle, motionPoints, morphPairs } from './model.js';
import { el, openMenu, toast } from './ui.js';

let show = null; // 현재 쇼 상태

/** 쇼 순서: 사용자 지정 쇼(opts.custom = 이름)면 그 슬라이드들, 아니면 숨기지 않은 슬라이드 */
const visibleSlides = (opts) => {
  const cs = opts.custom ? (S.pres.customShows ?? []).find((c) => c.name === opts.custom) : null;
  if (cs) return cs.slides.map((id) => S.pres.slides.findIndex((x) => x.id === id)).filter((i) => i >= 0);
  return S.pres.slides.map((s, i) => i).filter((i) => !S.pres.slides[i].hidden || opts.includeHidden);
};

/** opts: { from, presenter, windowed, rehearse } */
export function startShow(opts = {}) {
  if (show) endShow();
  const order = visibleSlides(opts);
  if (!order.length) { toast('보여 줄 슬라이드가 없습니다 (모두 숨김)'); return; }
  const setup = S.pres.show ?? {};
  let pos = opts.custom ? 0 : order.indexOf(opts.from ?? 0);
  if (pos < 0) pos = order.findIndex((i) => i > (opts.from ?? 0));
  if (pos < 0) pos = 0;
  const root = el('div', { class: `show${opts.windowed ? ' windowed' : ''}${opts.presenter ? ' presenter' : ''}`, tabIndex: -1 });
  const stage = el('div', { class: 'show-stage' });
  const ink = el('canvas', { class: 'show-ink' });
  const laser = el('div', { class: 'show-laser', hidden: true });
  const curtain = el('div', { class: 'show-curtain', hidden: true });
  const bar = el('div', { class: 'show-bar' },
    el('button', { title: '이전', onclick: (e) => { e.stopPropagation(); prev(); } }, '◀'),
    el('button', { title: '다음', onclick: (e) => { e.stopPropagation(); next(); } }, '▶'),
    el('button', { title: '펜 (Ctrl+P)', onclick: (e) => { e.stopPropagation(); setTool(show.tool === 'pen' ? null : 'pen'); } }, '✎'),
    el('button', { title: '레이저 포인터 (Ctrl+L)', onclick: (e) => { e.stopPropagation(); setTool(show.tool === 'laser' ? null : 'laser'); } }, '●'),
    el('button', { title: '메뉴', onclick: (e) => { e.stopPropagation(); showMenu(e); } }, '⋯'));
  const view = el('div', { class: 'show-view' }, stage, ink, laser, curtain);
  root.append(view, bar);
  show = { root, stage, view, ink, laser, curtain, order, pos, step: 0, steps: [], opts, setup, timers: [], started: Date.now(), slideStart: Date.now(), times: {}, tool: null, typed: '', audience: null, presenterEls: null };
  if (opts.presenter) buildPresenter();
  document.body.append(root);
  root.addEventListener('keydown', onKey);
  root.addEventListener('click', onClick);
  root.addEventListener('contextmenu', (e) => { e.preventDefault(); showMenu(e); });
  root.addEventListener('wheel', (e) => { if (Math.abs(e.deltaY) > 20) { if (e.deltaY > 0) next(); else prev(); } }, { passive: true });
  root.addEventListener('pointermove', onPointerMove);
  addEventListener('resize', fit);
  if (!opts.windowed && !opts.presenter && root.requestFullscreen) root.requestFullscreen().catch(() => {});
  document.addEventListener('fullscreenchange', onFsChange);
  root.focus();
  enterSlide(pos, { transition: false });
  fit();
}

function onFsChange() { if (show && !document.fullscreenElement && !show.opts.windowed && !show.opts.presenter && show.wasFull) endShow(); if (show && document.fullscreenElement) show.wasFull = true; }

export function endShow() {
  if (!show) return;
  const s = show;
  clearTimers();
  clearInterval(s.clock);
  if (s.opts.rehearse) { recordTime(); saveRehearsal(s.times); }
  s.root.remove();
  s.audience?.close?.();
  removeEventListener('resize', fit);
  document.removeEventListener('fullscreenchange', onFsChange);
  if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
  show = null;
  emit('showEnded');
}
export const showActive = () => !!show;

function clearTimers() { for (const t of show?.timers ?? []) clearTimeout(t); if (show) show.timers = []; }
const later = (ms, fn) => { const t = setTimeout(fn, ms); show.timers.push(t); return t; };

/** 슬라이드 크기를 창에 맞춤 */
function fit() {
  if (!show) return;
  const v = show.view.getBoundingClientRect();
  const { w, h } = S.pres.size;
  const sc = Math.min(v.width / w, v.height / h);
  show.scale = sc;
  for (const st of show.stage.querySelectorAll('.show-slide')) {
    st.style.transform = `translate(${(v.width - w * sc) / 2}px, ${(v.height - h * sc) / 2}px) scale(${sc})`;
  }
  show.ink.width = v.width;
  show.ink.height = v.height;
  if (show.presenterEls) updatePresenter();
}

/** 슬라이드 HTML (단계 stepsDone 까지 진행한 상태의 보이기/숨기기) */
function slideState(idx, stepsDone) {
  const slide = S.pres.slides[idx];
  const steps = animSteps(slide);
  const hide = new Set();
  const paraHide = new Set(); // 'objId#단락' (단락별 애니메이션)
  const moved = new Map();    // 이동 경로가 끝난 개체 → 마지막 위치
  const key = (a) => (a.para != null ? `${a.obj}#${a.para}` : a.obj);
  const first = new Map();
  for (const st of steps) for (const a of st) if (!first.has(key(a))) first.set(key(a), a);
  for (const [k, a] of first) if (a.cls === 'entr') (a.para != null ? paraHide : hide).add(k);
  for (let k = 0; k < Math.min(stepsDone, steps.length); k++) {
    for (const a of steps[k]) {
      const set = a.para != null ? paraHide : hide;
      if (a.cls === 'entr') set.delete(key(a));
      if (a.cls === 'exit') set.add(key(a));
      if (a.cls === 'path') { const pts = motionPoints(a.motion); const last = pts[pts.length - 1]; const prev = moved.get(a.obj) ?? [0, 0]; moved.set(a.obj, [prev[0] + last[0], prev[1] + last[1]]); }
    }
  }
  return { slide, steps, hide, paraHide, moved };
}

function makeSlideEl(idx, stepsDone) {
  const { slide, hide, paraHide, moved } = slideState(idx, stepsDone);
  const d = el('div', { class: 'show-slide', html: slideHtml(S.pres, slide, { index: idx, hide, links: true, live: true }) });
  const { w, h } = S.pres.size;
  for (const k of paraHide) { const [id, pi] = k.split('#'); const p = d.querySelectorAll(`.ob[data-id="${CSS.escape(id)}"] .p`)[Number(pi)]; if (p) p.style.visibility = 'hidden'; }
  for (const [id, [fx, fy]] of moved) { const ob = d.querySelector(`.ob[data-id="${CSS.escape(id)}"]`); if (ob) ob.style.translate = `${fx * w}px ${fy * h}px`; }
  Object.assign(d.style, { width: `${w}px`, height: `${h}px` });
  return d;
}

function enterSlide(pos, { transition = true, stepsDone = 0, back = false } = {}) {
  clearTimers();
  recordTime();
  show.pos = pos;
  const idx = show.order[pos];
  const slide = S.pres.slides[idx];
  show.steps = animSteps(slide);
  show.step = Math.min(stepsDone, show.steps.length);
  show.slideStart = Date.now();
  clearInk();
  const old = show.stage.querySelector('.show-slide');
  const neu = makeSlideEl(idx, show.step);
  show.stage.append(neu);
  fit();
  const tr = slide.transition;
  if (old && transition && tr && tr.type !== 'none' && tr.type !== 'cut' && !back) {
    runTransition(old, neu, tr).then(() => { old.remove(); afterEnter(); });
  } else {
    old?.remove();
    afterEnter();
  }
  syncAudience();
}

function afterEnter() {
  if (!show) return;
  // 자동 재생 비디오 · 오디오
  for (const m of show.stage.querySelectorAll('.show-slide .av[data-autoplay]')) m.play?.().catch(() => {});
  // 첫 단계가 '이전 효과와 함께/다음에' 로 시작하면 바로 재생
  if (show.step === 0 && show.steps[0] && show.steps[0][0].start !== 'click') playStep();
  scheduleAuto();
  if (show.presenterEls) updatePresenter();
}

function scheduleAuto() {
  const slide = S.pres.slides[show.order[show.pos]];
  const t = slide.transition;
  const useTimings = show.setup.useTimings !== false;
  if (useTimings && t?.advAfter != null && !show.opts.rehearse) later(Math.max(0.1, t.advAfter) * 1000, () => next());
}

// ───────────── 전환 효과 (Web Animations) ─────────────
function runTransition(old, neu, tr) {
  const dur = Math.max(0.05, tr.dur ?? 1) * 1000;
  const easing = 'ease-in-out';
  const opt = { duration: dur, easing, fill: 'both' };
  const { w, h } = S.pres.size;
  const base = neu.style.transform;
  const dir = tr.dir ?? 'b';
  const off = { b: [0, h], t: [0, -h], l: [-w, 0], r: [w, 0] }[dir] ?? [0, h];
  const T = (x, y) => `${base} translate(${x}px, ${y}px)`;
  const anims = [];
  switch (tr.type) {
    case 'fade': anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], opt)); break;
    case 'dissolve': anims.push(neu.animate([{ opacity: 0, filter: 'blur(6px)' }, { opacity: 1, filter: 'blur(0)' }], opt)); break;
    case 'push':
      anims.push(neu.animate([{ transform: T(off[0], off[1]) }, { transform: T(0, 0) }], opt));
      anims.push(old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], opt));
      break;
    case 'cover': anims.push(neu.animate([{ transform: T(off[0], off[1]) }, { transform: T(0, 0) }], opt)); break;
    case 'uncover':
      show.stage.append(old); // 이전 슬라이드를 위로 올려서 밀어냄
      anims.push(old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], opt));
      break;
    case 'wipe': {
      const from = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[dir] ?? 'inset(0 0 0 100%)';
      anims.push(neu.animate([{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], opt));
      break;
    }
    case 'split': anims.push(neu.animate([{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], opt)); break;
    case 'zoom': anims.push(neu.animate([{ transform: `${base} scale(.3)`, opacity: 0, transformOrigin: '50% 50%' }, { transform: base, opacity: 1 }], opt)); break;
    case 'circle': anims.push(neu.animate([{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }], opt)); break;
    case 'morph': {
      // 모핑: 짝이 되는 개체는 앞 슬라이드 자리 · 크기 · 회전에서 새 자리로, 나머지는 서서히
      const prevSlide = S.pres.slides[show.order[show.pos - 1] ?? -1] ?? S.pres.slides[show.order[show.pos]];
      const curSlide = S.pres.slides[show.order[show.pos]];
      const pairs = prevSlide === curSlide ? [] : morphPairs(prevSlide, curSlide);
      const matchedNew = new Set(pairs.map(([, b]) => b.id));
      const matchedOld = new Set(pairs.map(([a]) => a.id));
      const mopt = { duration: dur, easing: 'ease-in-out', fill: 'both' };
      for (const [a, b] of pairs) {
        const nEl = neu.querySelector(`.ob[data-id="${CSS.escape(b.id)}"]`);
        const oEl = old.querySelector(`.ob[data-id="${CSS.escape(a.id)}"]`);
        if (oEl) oEl.style.visibility = 'hidden';
        if (!nEl) continue;
        const dx = (a.x + a.w / 2) - (b.x + b.w / 2);
        const dy = (a.y + a.h / 2) - (b.y + b.h / 2);
        const sx = b.w ? a.w / b.w : 1;
        const sy = b.h ? a.h / b.h : 1;
        const dr = (a.rot ?? 0) - (b.rot ?? 0);
        anims.push(nEl.animate([{ translate: `${dx}px ${dy}px`, scale: `${sx} ${sy}`, rotate: `${dr}deg`, opacity: a.hidden ? 0 : 1 }, { translate: '0px 0px', scale: '1 1', rotate: '0deg', opacity: 1 }], mopt));
      }
      for (const elx of neu.querySelectorAll('.ob:not(.decor)')) if (!matchedNew.has(elx.dataset.id)) anims.push(elx.animate([{ opacity: 0 }, { opacity: 1 }], mopt));
      for (const elx of old.querySelectorAll('.ob:not(.decor)')) if (!matchedOld.has(elx.dataset.id)) anims.push(elx.animate([{ opacity: 1 }, { opacity: 0 }], mopt));
      // 배경 · 장식은 겹치며 바뀜
      anims.push(neu.animate([{ opacity: 0.999 }, { opacity: 1 }], mopt));
      break;
    }
    case 'flip':
      anims.push(old.animate([{ transform: `${base} perspective(2000px) rotateY(0)`, opacity: 1 }, { transform: `${base} perspective(2000px) rotateY(90deg)`, opacity: 0 }], { ...opt, duration: dur / 2 }));
      anims.push(neu.animate([{ transform: `${base} perspective(2000px) rotateY(-90deg)`, opacity: 0 }, { transform: `${base} perspective(2000px) rotateY(0)`, opacity: 1 }], { ...opt, duration: dur / 2, delay: dur / 2 }));
      break;
    default: anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], opt));
  }
  show.anims = anims;
  return Promise.all(anims.map((a) => a.finished.catch(() => {}))).then(() => { for (const a of anims) { try { a.cancel(); } catch { /* 끝남 */ } } });
}

// ───────────── 개체 애니메이션 ─────────────
function keyframes(a, o) {
  const { w, h } = S.pres.size;
  const dir = a.dir ?? 'b';
  const fly = { b: `0 ${h - o.y + 10}px`, t: `0 ${-(o.y + o.h + 10)}px`, l: `${-(o.x + o.w + 10)}px 0`, r: `${w - o.x + 10}px 0` }[dir];
  const wipeFrom = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[dir];
  switch (a.effect) {
    case 'appear': case 'disappear': return null;
    case 'fade': return [{ opacity: 0 }, { opacity: 1 }];
    case 'fly': return [{ translate: fly }, { translate: '0 0' }];
    case 'float': return [{ translate: '0 40px', opacity: 0 }, { translate: '0 0', opacity: 1 }];
    case 'split': return [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }];
    case 'wipe': return [{ clipPath: wipeFrom }, { clipPath: 'inset(0 0 0 0)' }];
    case 'zoom': return [{ scale: '0.3', opacity: 0 }, { scale: '1', opacity: 1 }];
    case 'grow': return [{ scale: '0', rotate: '90deg', opacity: 0 }, { scale: '1', rotate: '0deg', opacity: 1 }];
    case 'wheel': return [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }];
    case 'bounce': return [{ translate: `0 ${-(o.y + o.h)}px`, offset: 0 }, { translate: '0 0', offset: 0.55, easing: 'ease-in' }, { translate: '0 -30px', offset: 0.75, easing: 'ease-out' }, { translate: '0 0', offset: 0.9 }, { translate: '0 -8px', offset: 0.95 }, { translate: '0 0', offset: 1 }];
    case 'pulse': return [{ scale: '1' }, { scale: '1.08' }, { scale: '1' }];
    case 'spin': return [{ rotate: '0deg' }, { rotate: '360deg' }];
    case 'growShrink': return [{ scale: '1' }, { scale: '1.5' }];
    case 'teeter': return [{ rotate: '0deg' }, { rotate: '6deg' }, { rotate: '-6deg' }, { rotate: '4deg' }, { rotate: '0deg' }];
    case 'flash': return [{ opacity: 1 }, { opacity: 0 }, { opacity: 1 }];
    default: return [{ opacity: 0 }, { opacity: 1 }];
  }
}

/** 개체 하나 재생 (편집 화면 미리 보기에서도 씀) */
export function animateObject(elm, a, o) {
  if (a.cls === 'path') {
    // 이동 경로: 지금 위치(앞 경로 끝)에서 경로만큼 이동하고 그 자리에 머묾
    const { w, h } = S.pres.size;
    const [bx, by] = (elm.style.translate || '0px 0px').split(' ').map((v) => parseFloat(v) || 0);
    const pts = motionPoints(a.motion);
    const kf = pts.map(([x, y]) => ({ translate: `${bx + x * w}px ${by + y * h}px` }));
    const last = kf[kf.length - 1].translate;
    const an = elm.animate(kf, { duration: Math.max(1, (a.dur ?? 2) * 1000), easing: 'ease-in-out', fill: 'forwards' });
    return an.finished.then(() => { elm.style.translate = last; an.cancel(); }).catch(() => {});
  }
  const kf = keyframes(a, o);
  const dur = Math.max(1, (a.dur ?? 0.5) * 1000);
  if (a.cls === 'entr') {
    elm.style.visibility = 'visible';
    if (kf) return elm.animate(kf, { duration: dur, easing: 'ease-out', fill: 'none' }).finished.catch(() => {});
    return Promise.resolve();
  }
  if (a.cls === 'exit') {
    if (!kf) { elm.style.visibility = 'hidden'; return Promise.resolve(); }
    const anim = elm.animate([...kf].reverse(), { duration: dur, easing: 'ease-in', fill: 'forwards' });
    return anim.finished.then(() => { elm.style.visibility = 'hidden'; anim.cancel(); }).catch(() => {});
  }
  const fillMode = a.effect === 'growShrink' ? 'forwards' : 'none';
  return kf ? elm.animate(kf, { duration: dur, easing: 'ease-in-out', fill: fillMode }).finished.catch(() => {}) : Promise.resolve();
}

/** 한 클릭 묶음 실행 */
export function playStepOn(container, step, slide) {
  const tl = stepTimeline(step);
  const done = [];
  for (const { anim, at } of tl) {
    const o = slide.objects.find((x) => x.id === anim.obj);
    const ob = container.querySelector(`.ob[data-id="${CSS.escape(anim.obj)}"]`);
    const elm = anim.para != null ? ob?.querySelectorAll('.p')[anim.para] : ob;
    if (!o || !elm) continue;
    done.push(new Promise((res) => setTimeout(() => animateObject(elm, anim, o).then(res), at * 1000)));
  }
  return Promise.all(done);
}

function playStep() {
  const step = show.steps[show.step];
  if (!step) return;
  show.step++;
  const slide = S.pres.slides[show.order[show.pos]];
  const cur = show.stage.querySelector('.show-slide:last-child');
  playStepOn(cur, step, slide).then(() => {
    // 다음 묶음이 '이전 효과 다음에'로 이어지면 자동 재생 (animSteps 가 묶어 둠) — 여기서는 없음
    if (show?.presenterEls) updatePresenter();
  });
  syncAudience();
  if (show.presenterEls) updatePresenter();
}

// ───────────── 이동 ─────────────
export function next() {
  if (!show) return;
  if (show.ended) { endShow(); return; }
  if (show.blank) { setBlank(null); return; }
  if (show.step < show.steps.length) { playStep(); return; }
  if (show.pos + 1 < show.order.length) { enterSlide(show.pos + 1); return; }
  if (show.setup.loop) { enterSlide(0); return; }
  endScreen();
}
export function prev() {
  if (!show) return;
  if (show.ended) { show.ended = false; show.stage.querySelector('.show-end')?.remove(); enterSlide(show.pos, { transition: false, stepsDone: 999, back: true }); return; }
  if (show.step > 0) { show.step--; rerender(); return; }
  if (show.pos > 0) enterSlide(show.pos - 1, { transition: false, stepsDone: 999, back: true });
}
function goTo(pos) { if (pos >= 0 && pos < show.order.length) { show.ended = false; show.stage.querySelector('.show-end')?.remove(); enterSlide(pos, { transition: false }); } }
function rerender() {
  const old = show.stage.querySelector('.show-slide');
  const neu = makeSlideEl(show.order[show.pos], show.step);
  old?.replaceWith(neu);
  fit();
  syncAudience();
  if (show.presenterEls) updatePresenter();
}
function endScreen() {
  recordTime();
  show.ended = true;
  show.stage.querySelector('.show-slide')?.remove();
  show.stage.append(el('div', { class: 'show-end' }, '슬라이드 쇼의 끝입니다. 끝내려면 클릭하십시오.'));
  syncAudience();
}
function setBlank(color) {
  show.blank = color;
  show.curtain.hidden = !color;
  show.curtain.style.background = color ?? '';
}

function onKey(e) {
  if (!show) return;
  const k = e.key;
  e.stopPropagation();
  if (/^[0-9]$/.test(k)) { show.typed += k; return; }
  if (k === 'Enter' && show.typed) { const n = Number(show.typed) - 1; show.typed = ''; const pos = show.order.indexOf(n); if (pos >= 0) goTo(pos); return; }
  show.typed = '';
  if ((e.ctrlKey || e.metaKey) && (k === 'p' || k === 'P')) { e.preventDefault(); setTool(show.tool === 'pen' ? null : 'pen'); return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'l' || k === 'L')) { e.preventDefault(); setTool(show.tool === 'laser' ? null : 'laser'); return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'a' || k === 'A')) { e.preventDefault(); setTool(null); return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'e' || k === 'E')) { e.preventDefault(); clearInk(); setTool(null); return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'm' || k === 'M')) { e.preventDefault(); show.ink.hidden = !show.ink.hidden; return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'h' || k === 'H')) { e.preventDefault(); show.root.classList.add('no-cursor'); return; }
  if ((e.ctrlKey || e.metaKey) && (k === 'u' || k === 'U')) { e.preventDefault(); show.root.classList.remove('no-cursor'); return; }
  if (((e.ctrlKey || e.metaKey) && (k === 's' || k === 'S')) || k === 'g' || k === 'G') { e.preventDefault(); slideList(); return; }
  if (k === 'e' || k === 'E') { clearInk(); return; }
  if (k === 'h' || k === 'H') { e.preventDefault(); next(); return; }
  switch (k) {
    case 'Escape': case '-': e.preventDefault(); if (show.tool) setTool(null); else endShow(); break;
    case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': case 'N': e.preventDefault(); next(); break;
    case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P': e.preventDefault(); prev(); break;
    case 'Home': e.preventDefault(); goTo(0); break;
    case 'End': e.preventDefault(); goTo(show.order.length - 1); break;
    case 'b': case 'B': case '.': setBlank(show.blank === '#000' ? null : '#000'); break;
    case 'w': case 'W': case ',': setBlank(show.blank === '#fff' ? null : '#fff'); break;
    case 'F5': e.preventDefault(); break;
    default:
  }
}

function onClick(e) {
  if (!show || e.button !== 0) return;
  if (e.target.closest('.show-bar, .pv-side, .pv-tools')) return;
  if (show.tool === 'pen') return;
  // 비디오는 자체 재생 단추, 오디오 아이콘은 눌러서 재생/일시 정지
  if (e.target.closest('video')) return;
  const au = e.target.closest('.av-audio')?.querySelector('audio');
  if (au) { if (au.paused) au.play().catch(() => {}); else au.pause(); return; }
  const link = e.target.closest('a.lnk, [data-link]');
  if (link) {
    const href = link.getAttribute('href') ?? link.dataset.link;
    const m = /^#slide(\d+)$/.exec(href ?? '');
    e.preventDefault();
    if (m) { const pos = show.order.indexOf(Number(m[1]) - 1); if (pos >= 0) goTo(pos); } else if (href) window.open(href, '_blank', 'noopener');
    return;
  }
  next();
}

function showMenu(e) {
  openMenu({ x: e.clientX, y: e.clientY }, [
    { label: '다음', key: 'N', action: next },
    { label: '이전', key: 'P', action: prev },
    { label: '슬라이드로 이동', submenu: show.order.map((idx, pos) => ({ label: `${idx + 1}  ${slideTitle(S.pres.slides[idx]) || '(제목 없음)'}`, checked: pos === show.pos, action: () => goTo(pos) })) },
    { sep: true },
    { label: '레이저 포인터', key: 'Ctrl+L', checked: show.tool === 'laser', action: () => setTool(show.tool === 'laser' ? null : 'laser') },
    { label: '펜', key: 'Ctrl+P', checked: show.tool === 'pen', action: () => setTool(show.tool === 'pen' ? null : 'pen') },
    { label: '잉크 지우기', key: 'E', action: clearInk },
    { sep: true },
    { label: '화면 검정', key: 'B', action: () => setBlank('#000') },
    { label: '화면 흰색', key: 'W', action: () => setBlank('#fff') },
    { label: show.presenterEls ? '청중 창 열기' : '발표자 보기 표시', action: () => { if (show.presenterEls) openAudience(); else { const pos = show.pos; const f = show.order[pos]; endShow(); startShow({ from: f, presenter: true }); } } },
    { sep: true },
    { label: '쇼 마침', key: 'Esc', action: endShow },
  ]);
}

/** Ctrl+S / G: 모든 슬라이드 목록에서 고르기 */
function slideList() {
  const r = show.root.getBoundingClientRect();
  openMenu({ x: r.left + r.width / 2 - 150, y: r.top + 40 }, show.order.map((n, pos) => ({ label: `${n + 1}. ${slideTitle(S.pres.slides[n]) || '(제목 없음)'}`, checked: pos === show.pos, action: () => goTo(pos) })), { scroll: true, minWidth: 300 });
}

// ───────────── 펜 · 레이저 ─────────────
function setTool(t) {
  show.tool = t;
  show.root.classList.toggle('pen', t === 'pen');
  show.root.classList.toggle('laser-on', t === 'laser');
  show.laser.hidden = t !== 'laser';
}
function clearInk() { const c = show?.ink; if (c) c.getContext('2d').clearRect(0, 0, c.width, c.height); }
let drawing = null;
function onPointerMove(e) {
  if (!show) return;
  show.root.classList.remove('idle');
  clearTimeout(show.idleT);
  show.idleT = setTimeout(() => show?.root.classList.add('idle'), 2500);
  const r = show.view.getBoundingClientRect();
  if (show.tool === 'laser') { show.laser.style.left = `${e.clientX - r.left}px`; show.laser.style.top = `${e.clientY - r.top}px`; }
  if (show.tool === 'pen') {
    if (e.buttons === 1) {
      const ctx = show.ink.getContext('2d');
      const p = [e.clientX - r.left, e.clientY - r.top];
      if (drawing) {
        ctx.strokeStyle = '#e81123';
        ctx.lineWidth = 3;
        ctx.lineCap = 'round';
        ctx.beginPath();
        ctx.moveTo(...drawing);
        ctx.lineTo(...p);
        ctx.stroke();
      }
      drawing = p;
    } else drawing = null;
  }
}

// ───────────── 발표자 보기 ─────────────
function buildPresenter() {
  const side = el('div', { class: 'pv-side' });
  const nextBox = el('div', { class: 'pv-next' });
  const notes = el('div', { class: 'pv-notes' });
  const timer = el('div', { class: 'pv-timer' });
  const counter = el('div', { class: 'pv-counter' });
  let fs = 22;
  const tools = el('div', { class: 'pv-tools' },
    el('button', { onclick: () => prev() }, '◀ 이전'),
    counter,
    el('button', { onclick: () => next() }, '다음 ▶'),
    el('button', { title: '글자 크게', onclick: () => { fs = Math.min(48, fs + 2); notes.style.fontSize = `${fs}px`; } }, 'A+'),
    el('button', { title: '글자 작게', onclick: () => { fs = Math.max(12, fs - 2); notes.style.fontSize = `${fs}px`; } }, 'A−'),
    el('button', { onclick: () => { show.started = Date.now(); } }, '타이머 다시'),
    el('button', { onclick: openAudience }, '청중 창 열기'),
    el('button', { onclick: endShow }, '쇼 마침'));
  side.append(timer, el('div', { class: 'pv-label' }, '다음 슬라이드'), nextBox, el('div', { class: 'pv-label' }, '메모'), notes);
  show.root.append(side, tools);
  show.presenterEls = { side, nextBox, notes, timer, counter };
  show.clock = setInterval(updateTimer, 1000);
}
function updateTimer() {
  if (!show?.presenterEls) { clearInterval(show?.clock); return; }
  const s = Math.floor((Date.now() - show.started) / 1000);
  const hh = String(Math.floor(s / 3600)).padStart(2, '0');
  const mm = String(Math.floor((s % 3600) / 60)).padStart(2, '0');
  const ss = String(s % 60).padStart(2, '0');
  show.presenterEls.timer.textContent = `${hh}:${mm}:${ss}   ·   ${new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })}`;
}
function updatePresenter() {
  const p = show.presenterEls;
  const idx = show.order[show.pos];
  const slide = S.pres.slides[idx];
  p.counter.textContent = `슬라이드 ${show.pos + 1} / ${show.order.length}`;
  p.notes.innerHTML = slide.notes ? escHtml(slide.notes).replace(/\n/g, '<br>') : '<span class="muted">메모 없음</span>';
  const nextPos = show.step < show.steps.length ? show.pos : show.pos + 1;
  const nextIdx = show.order[nextPos];
  const w = p.nextBox.clientWidth || 300;
  const sc = w / S.pres.size.w;
  p.nextBox.innerHTML = nextIdx == null ? '<div class="pv-end">프레젠테이션 끝</div>' : `<div class="tw" style="width:${w}px;height:${S.pres.size.h * sc}px"><div class="tw-in" style="transform:scale(${sc})">${slideHtml(S.pres, S.pres.slides[nextIdx], { index: nextIdx, hide: nextPos === show.pos ? slideState(nextIdx, show.step + 1).hide : slideState(nextIdx, 0).hide })}</div></div>`;
  updateTimer();
}

/** 청중 창: 같은 출처의 새 창에 현재 슬라이드를 그대로 보여 줌 (두 번째 모니터로 옮겨 전체 화면) */
function openAudience() {
  const w = window.open('', 'wipoint-audience', 'popup,width=960,height=560');
  if (!w) { toast('팝업이 차단되었습니다. 브라우저에서 팝업을 허용해 주세요.'); return; }
  w.document.open();
  w.document.write(`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>WIPOINT - 청중 화면 (두 번 클릭: 전체 화면)</title><style>${SLIDE_CSS}html,body{margin:0;height:100%;background:#000;overflow:hidden}#v{position:absolute;inset:0}.show-slide{position:absolute;left:0;top:0;transform-origin:0 0}.show-end{color:#fff;font:20px sans-serif;padding:20px}</style></head><body><div id="v"></div></body></html>`);
  w.document.close();
  // 청중 창을 두 번 클릭하면 전체 화면 (보안 정책상 인라인 스크립트 없이 이 창에서 연결)
  w.document.addEventListener('dblclick', () => w.document.documentElement.requestFullscreen?.().catch(() => {}));
  show.audience = w;
  w.addEventListener('resize', syncAudience);
  syncAudience();
}
function syncAudience() {
  const w = show?.audience;
  if (!w || w.closed) return;
  const v = w.document.getElementById('v');
  if (!v) return;
  if (show.ended) { v.innerHTML = '<div class="show-end">슬라이드 쇼의 끝입니다.</div>'; return; }
  const idx = show.order[show.pos];
  const { hide } = slideState(idx, show.step);
  const { w: sw, h: sh } = S.pres.size;
  const sc = Math.min(w.innerWidth / sw, w.innerHeight / sh);
  v.innerHTML = `<div class="show-slide" style="width:${sw}px;height:${sh}px;transform:translate(${(w.innerWidth - sw * sc) / 2}px,${(w.innerHeight - sh * sc) / 2}px) scale(${sc})">${slideHtml(S.pres, S.pres.slides[idx], { index: idx, hide })}</div>`;
}

// ───────────── 예행 연습 ─────────────
function recordTime() {
  if (!show?.opts.rehearse || show.pos == null) return;
  const idx = show.order[show.pos];
  const sec = (Date.now() - show.slideStart) / 1000;
  show.times[idx] = Math.round(((show.times[idx] ?? 0) + sec) * 10) / 10;
  show.slideStart = Date.now();
}
function saveRehearsal(times) {
  const entries = Object.entries(times);
  if (!entries.length) return;
  const total = entries.reduce((s, [, t]) => s + t, 0);
  setTimeout(() => {
    if (!confirm(`슬라이드 쇼에 걸린 총 시간은 ${Math.floor(total / 60)}분 ${Math.round(total % 60)}초입니다. 슬라이드 시간을 저장하시겠습니까?`)) return;
    change(() => {
      for (const [i, t] of entries) {
        const s = S.pres.slides[Number(i)];
        s.transition = { ...(s.transition ?? { type: 'none' }), advAfter: Math.max(1, Math.round(t)) };
      }
    }, { scope: 'all' });
  }, 50);
}
