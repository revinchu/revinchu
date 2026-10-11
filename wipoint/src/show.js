// 슬라이드 쇼: 전환 효과 · 클릭 단계별 애니메이션 · 발표자 보기 · 펜/레이저 · 예행 연습
import { S, change, emit } from './state.js';
import { uid } from './model.js';
import { slideHtml, SLIDE_CSS, escHtml } from './render.js';
import { animSteps, triggerSteps, stepTimeline, slideTitle, motionPoints, morphPairs } from './model.js';
import { el, openMenu, toast, openDialog as import_openDialog } from './ui.js';
import { mixDown, resample, encodeWav, inkUntil, captionText } from './recording.js';
import { blobUrl } from './blobstore.js';
import { playSound } from './sfx.js';
import { blindsPoly, combPoly, checkerPoly, randomBarsFrames, randomCellsFrames, sweepFrames, diamondPoly, plusPoly, stripsPoly, frames } from './transfx.js';

let show = null; // 현재 쇼 상태

/** 쇼 순서: 사용자 지정 쇼(opts.custom = 이름)면 그 슬라이드들, 아니면 숨기지 않은 슬라이드 */
const visibleSlides = (opts) => {
  const setup = S.pres.show ?? {};
  // 쇼 설정의 [슬라이드 표시]: 사용자 지정 쇼 · 범위 (리본에서 고른 사용자 지정 쇼가 먼저)
  const name = opts.custom ?? (setup.slides === 'custom' ? setup.custom : null);
  const cs = name ? (S.pres.customShows ?? []).find((c) => c.name === name) : null;
  if (cs) return cs.slides.map((id) => S.pres.slides.findIndex((x) => x.id === id)).filter((i) => i >= 0);
  const [a, b] = setup.slides === 'range' ? [Math.max(1, setup.from ?? 1) - 1, Math.max(1, setup.to ?? S.pres.slides.length) - 1] : [0, Infinity];
  return S.pres.slides.map((s, i) => i).filter((i) => i >= a && i <= b && (!S.pres.slides[i].hidden || opts.includeHidden));
};

// ───────────── 여러 모니터 (Window Management API) ─────────────
/** 화면 목록 (모니터가 2대 이상일 때만 권한을 물음) */
export async function screenList(ask = false) {
  if (!('getScreenDetails' in window)) return null;
  if (!ask && !window.screen.isExtended) return null;
  try { return await window.getScreenDetails(); } catch { return null; }
}
/** 쇼 설정의 모니터 → 슬라이드를 보여 줄 화면 (자동 = 지금 창이 없는 다른 화면) */
function targetScreen(sd, monitor) {
  if (!sd || sd.screens.length < 2) return null;
  if (monitor === 'primary') return sd.screens.find((x) => x.isPrimary) ?? null;
  if (Number.isInteger(monitor) && sd.screens[monitor]) return sd.screens[monitor];
  return sd.screens.find((x) => x !== sd.currentScreen) ?? null;
}

/** opts: { from, presenter, windowed, rehearse } */
export function startShow(opts = {}) {
  if (show) endShow();
  // 녹화: 마이크(· 카메라) 권한을 먼저 받음 — 거절되면 타이밍 · 잉크만 기록
  if (opts.record && !opts.stream) {
    const want = { audio: opts.record.mic !== false, video: opts.record.camera ? { width: 640, height: 480 } : false };
    if (!want.audio && !want.video) { startShow({ ...opts, stream: 'none' }); return; }
    navigator.mediaDevices?.getUserMedia(want).then((stream) => startShow({ ...opts, stream }), () => { toast('마이크/카메라를 사용할 수 없어 타이밍과 잉크만 기록합니다'); startShow({ ...opts, stream: 'none' }); })
      ?? startShow({ ...opts, stream: 'none' });
    return;
  }
  const setup0 = S.pres.show ?? {};
  if (setup0.type === 'browse') opts = { ...opts, windowed: true };
  // 모니터가 여러 대면: 슬라이드는 고른 모니터에 전체 화면, 이 화면에는 발표자 보기 (PowerPoint 와 같음)
  if (!opts.windowed && !opts.screens && !opts.rehearse && window.screen?.isExtended && 'getScreenDetails' in window && setup0.monitor !== 'off') {
    screenList().then((sd) => startShow({ ...opts, screens: sd ?? 'none' })).catch(() => startShow({ ...opts, screens: 'none' }));
    return;
  }
  const sd = opts.screens && opts.screens !== 'none' ? opts.screens : null;
  const target = targetScreen(sd, setup0.monitor);
  if (target && target !== sd.currentScreen && setup0.usePresenter !== false && !opts.presenter && !opts.noPresenter) opts = { ...opts, presenter: true, audienceOn: target };
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
  if (opts.record) startRecording(opts.stream === 'none' ? null : opts.stream, opts.record);
  setupCaptions();
  document.body.append(root);
  root.addEventListener('keydown', onKey);
  root.addEventListener('click', onClick);
  root.addEventListener('contextmenu', (e) => { e.preventDefault(); showMenu(e); });
  root.addEventListener('wheel', (e) => { if (Math.abs(e.deltaY) > 20) { if (e.deltaY > 0) next(); else prev(); } }, { passive: true });
  root.addEventListener('pointermove', onPointerMove);
  addEventListener('resize', fit);
  if (!opts.windowed && !opts.presenter && root.requestFullscreen) root.requestFullscreen(target && setup0.usePresenter === false ? { screen: target } : undefined).catch(() => {});
  if (opts.presenter && opts.audienceOn) { root.requestFullscreen?.({ screen: sd.currentScreen }).catch(() => {}); openAudience(opts.audienceOn); }
  if (setup.penColor) show.penColor = setup.penColor;
  if (setup.laserColor) laser.style.setProperty('--laser', setup.laserColor);
  document.addEventListener('fullscreenchange', onFsChange);
  root.focus();
  enterSlide(pos, { transition: false });
  fit();
}

function onFsChange() { if (show && !document.fullscreenElement && !show.opts.windowed && !show.opts.presenter && show.wasFull) endShow(); if (show && document.fullscreenElement) show.wasFull = true; }

export function endShow() {
  if (!show) return;
  const s = show;
  const ink = [...(s.strokes ?? new Map())].filter(([, l]) => l.some((st) => st.pts.length > 1));
  if (ink.length && !s.viewOnlyInk && !s.rec) setTimeout(() => keepInkDialog(ink), 50);
  clearTimers();
  clearInterval(s.clock);
  if (s.opts.rehearse) { recordTime(); saveRehearsal(s.times); }
  if (s.rec) finishRecording(s.rec);
  stopCaptions(s);
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
  redrawInk();
  if (show.presenterEls) updatePresenter();
}

/** 슬라이드 HTML (단계 stepsDone 까지 진행한 상태의 보이기/숨기기) */
function slideState(idx, stepsDone) {
  const slide = S.pres.slides[idx];
  if (show?.setup?.noAnim) return { slide, hide: new Set(), paraHide: new Set(), moved: new Map() };
  const steps = animSteps(slide);
  const hide = new Set();
  const paraHide = new Set(); // 'objId#단락' (단락별 애니메이션)
  const moved = new Map();    // 이동 경로가 끝난 개체 → 마지막 위치
  const key = (a) => (a.para != null ? `${a.obj}#${a.para}` : a.obj);
  const first = new Map();
  for (const st of steps) for (const a of st) if (!first.has(key(a))) first.set(key(a), a);
  // 트리거 애니메이션: 처음 효과가 나타내기면 숨김, 이미 누른 트리거 (show.trigDone) 는 끝난 상태
  const trig = triggerSteps(slide);
  const done = [];
  for (const [t, tsteps] of trig) {
    for (const st of tsteps) for (const a of st) if (!first.has(key(a))) first.set(key(a), a);
    const n = show?.order?.[show.pos] === idx ? (show.trigDone?.get(t) ?? 0) : 0;
    done.push(...tsteps.slice(0, n));
  }
  for (const [k, a] of first) if (a.cls === 'entr') (a.para != null ? paraHide : hide).add(k);
  for (const st of [...steps.slice(0, Math.min(stepsDone, steps.length)), ...done]) {
    for (const a of st) {
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
  for (const t of triggerSteps(slide).keys()) d.querySelector(`.ob[data-id="${CSS.escape(t)}"]`)?.classList.add('trig');
  Object.assign(d.style, { width: `${w}px`, height: `${h}px` });
  return d;
}

function enterSlide(pos, { transition = true, stepsDone = 0, back = false } = {}) {
  clearTimers();
  recordTime();
  show.pos = pos;
  const idx = show.order[pos];
  const slide = S.pres.slides[idx];
  if (show.rec) recSegment(idx);
  show.steps = show.setup.noAnim ? [] : animSteps(slide);
  show.step = Math.min(stepsDone, show.steps.length);
  show.triggers = triggerSteps(slide);
  show.trigDone = new Map();
  show.slideStart = Date.now();
  clearInk(false);
  const old = show.stage.querySelector('.show-slide');
  const neu = makeSlideEl(idx, show.step);
  show.stage.append(neu);
  fit();
  const tr = slide.transition;
  if (transition && tr?.sound && !back) playSound(tr.sound);
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
  for (const m of show.stage.querySelectorAll('.show-slide .av[data-autoplay]')) {
    // 녹화 중이거나 [내레이션 재생] 을 끈 경우 내레이션은 재생하지 않음
    if (m.dataset.narration && (show.rec || show.setup.noNarration)) continue;
    m.play?.().catch(() => {});
  }
  replayInk();
  // 첫 단계가 '이전 효과와 함께/다음에' 로 시작하면 바로 재생
  if (show.step === 0 && show.steps[0] && show.steps[0][0].start !== 'click') playStep();
  scheduleAuto();
  if (show.presenterEls) updatePresenter();
}

function scheduleAuto() {
  const slide = S.pres.slides[show.order[show.pos]];
  const t = slide.transition;
  const useTimings = show.setup.useTimings !== false || kiosk();
  if (useTimings && t?.advAfter != null && !show.opts.rehearse) later(Math.max(0.1, t.advAfter) * 1000, () => next());
}

// ───────────── 전환 효과 (Web Animations) ─────────────
function runTransition(old, neu, tr) {
  const cur = S.pres.slides[show.order[show.pos]];
  const prev = S.pres.slides[show.order[show.pos - 1] ?? -1] ?? cur;
  const r = playTransition(old, neu, tr, { host: show.stage, raise: (e) => show.stage.append(e), prev, cur });
  show.anims = r.anims;
  return r.done;
}

/** 전환 효과 재생 (슬라이드 쇼 · 편집 화면 미리 보기 공용) — old/neu: transform-origin 0 0 인 슬라이드 요소,
 *  ctx = { host (조각을 넣을 부모), raise(el) (요소를 맨 위로), prev, cur (모핑용 슬라이드) } → { anims, done } */
export function playTransition(old, neu, tr, ctx) {
  const dur = Math.max(0.05, tr.dur ?? 1) * 1000;
  const easing = 'ease-in-out';
  const opt = { duration: dur, easing, fill: 'both' };
  const { w, h } = S.pres.size;
  const base = neu.style.transform;
  const dir = tr.dir ?? 'b';
  const off = { b: [0, h], t: [0, -h], l: [-w, 0], r: [w, 0] }[dir] ?? [0, h];
  const T = (x, y) => `${base} translate(${x}px, ${y}px)`;
  const anims = [];
  const ghosts = [];
  // 슬라이드 가운데(또는 지정한 점)를 기준으로 3D 변환 — 슬라이드 요소는 transform-origin 0 0 + 맞춤 배율
  const P = `perspective(${Math.round(w * 1.6)}px)`;
  const around = (ox, oy, t) => `${base} translate(${ox}px, ${oy}px) ${P} ${t} translate(${-ox}px, ${-oy}px)`;
  const C = (t) => around(w / 2, h / 2, t);
  const clipAnim = (elx, list, o = opt) => anims.push(elx.animate(list.map((clipPath) => ({ clipPath })), o));
  /** 앞 슬라이드 조각 (복제본) — 끝나면 지움. top: 새 슬라이드 위에 */
  const ghost = (clip, top = true) => {
    const g = old.cloneNode(true);
    g.classList.remove('show-slide');
    g.classList.add('show-slide-ghost');
    g.style.visibility = 'visible';
    if (clip) g.style.clipPath = clip;
    for (const m of g.querySelectorAll('video, audio')) { m.removeAttribute('autoplay'); m.pause?.(); }
    ctx.host.insertBefore(g, neu);
    if (top) ctx.raise(g);
    ghosts.push(g);
    return g;
  };
  const oldOnTop = () => { ctx.raise(old); };
  const pieces = (cols, rows, fn) => {
    old.style.visibility = 'hidden';
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
      const x1 = (c / cols) * 100; const y1 = (r / rows) * 100;
      const g = ghost(`polygon(${x1}% ${y1}%, ${x1 + 100 / cols}% ${y1}%, ${x1 + 100 / cols}% ${y1 + 100 / rows}%, ${x1}% ${y1 + 100 / rows}%)`);
      anims.push(fn(g, c, r));
    }
  };
  const rnd = (i, k = 1) => { const x = Math.sin(i * 12.9898 + k * 78.233) * 43758.5453; return x - Math.floor(x); };
  const half = { ...opt, duration: dur / 2 };
  const second = { ...opt, duration: dur / 2, delay: dur / 2 };
  const vert = dir === 'l' || dir === 'r';
  switch (tr.type) {
    case 'reveal':
      // 앞 슬라이드가 검게 사라진 뒤 새 슬라이드가 밝아짐
      oldOnTop();
      anims.push(old.animate([{ opacity: 1, filter: 'brightness(1)' }, { opacity: 0, filter: 'brightness(0)' }], half));
      anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], second));
      break;
    case 'flash':
      anims.push(neu.animate([{ opacity: 0, filter: 'brightness(4)' }, { opacity: 1, filter: 'brightness(3)', offset: 0.35 }, { opacity: 1, filter: 'brightness(1)' }], opt));
      break;
    case 'randomBars': clipAnim(neu, randomBarsFrames(16, 60, vert)); break;
    case 'blinds': clipAnim(neu, frames((t) => blindsPoly(t, 6, vert), 2)); break;
    case 'checker': clipAnim(neu, frames((t) => checkerPoly(t, 10, 6), 2)); break;
    case 'comb': clipAnim(neu, frames((t) => combPoly(t, 10, vert), 2)); break;
    case 'strips': clipAnim(neu, frames((t) => stripsPoly(t, { b: 'ru', t: 'rd', l: 'ld', r: 'rd' }[dir] ?? 'rd'), 2)); break;
    case 'diamond': clipAnim(neu, frames(diamondPoly, 2)); break;
    case 'plus': clipAnim(neu, frames(plusPoly, 2)); break;
    case 'clock': clipAnim(neu, sweepFrames(24, 'clock', 1), { ...opt, easing: 'linear' }); break;
    case 'dissolve': clipAnim(neu, randomCellsFrames(16, 32, 18), { ...opt, easing: 'linear' }); break;
    case 'glitter':
      clipAnim(neu, randomCellsFrames(16, 24, 14, 5), { ...opt, easing: 'linear' });
      anims.push(neu.animate([{ filter: 'brightness(1.8)' }, { filter: 'brightness(1)' }], opt));
      break;
    case 'honeycomb':
      oldOnTop();
      clipAnim(old, randomCellsFrames(16, 16, 9, 3).reverse(), { ...opt, easing: 'linear' });
      anims.push(old.animate([{ transform: base }, { transform: C('scale(1.15)') }], opt));
      break;
    case 'ripple':
      anims.push(neu.animate([{ clipPath: 'circle(0% at 50% 50%)', transform: C('scale(1.08)') }, { clipPath: 'circle(40% at 50% 50%)', transform: C('scale(.97)'), offset: 0.5 }, { clipPath: 'circle(75% at 50% 50%)', transform: base }], opt));
      break;
    case 'vortex':
      oldOnTop();
      anims.push(old.animate([{ transform: base, opacity: 1, filter: 'blur(0)' }, { transform: C('rotate(-200deg) scale(.2)'), opacity: 0, filter: 'blur(4px)' }], opt));
      anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], second));
      break;
    case 'shred':
      // 앞 슬라이드를 세로 띠로 잘라 아래로 떨어뜨림
      pieces(12, 1, (g, c) => g.animate([{ transform: base, opacity: 1 }, { transform: `${base} translate(0px, ${h * (0.9 + rnd(c) * 0.4)}px)`, opacity: 0 }], { ...opt, easing: 'ease-in', delay: rnd(c, 2) * dur * 0.3, duration: dur * 0.7 }));
      break;
    case 'fracture':
      pieces(5, 4, (g, c, r) => {
        const i = r * 5 + c; const dx = (c - 2) * 160 + (rnd(i) - 0.5) * 200; const dy = (r - 1.5) * 160 + (rnd(i, 3) - 0.5) * 200;
        return g.animate([{ transform: base, opacity: 1 }, { transform: around(((c + 0.5) * w) / 5, ((r + 0.5) * h) / 4, `translate3d(${dx}px, ${dy}px, ${200 + rnd(i, 5) * 400}px) rotateX(${rnd(i, 7) * 180}deg) rotateY(${rnd(i, 9) * 180}deg)`), opacity: 0 }], { ...opt, easing: 'ease-in', delay: rnd(i, 4) * dur * 0.2, duration: dur * 0.8 });
      });
      break;
    case 'crush':
      oldOnTop();
      anims.push(old.animate([{ transform: base, filter: 'blur(0) contrast(1)', opacity: 1 }, { transform: C('scale(.35) rotate(25deg)'), filter: 'blur(1px) contrast(1.6)', opacity: 1, offset: 0.6 }, { transform: C('translate(0px, 300px) scale(.05) rotate(80deg)'), opacity: 0 }], opt));
      break;
    case 'peelOff':
      oldOnTop();
      anims.push(old.animate([{ transform: base, opacity: 1 }, { transform: around(w, h, 'rotate3d(1, -1, 0, 100deg)'), opacity: 0.2 }], { ...opt, easing: 'ease-in' }));
      break;
    case 'pageCurl':
      oldOnTop();
      anims.push(old.animate([{ transform: around(0, h / 2, 'rotateY(0deg)'), filter: 'brightness(1)' }, { transform: around(0, h / 2, 'rotateY(-95deg)'), filter: 'brightness(.6)' }], { ...opt, easing: 'ease-in' }));
      break;
    case 'airplane':
      oldOnTop();
      anims.push(old.animate([{ transform: base, opacity: 1 }, { transform: C('scale(.25) rotateX(40deg)'), opacity: 1, offset: 0.45 }, { transform: C(`translate(${w}px, ${-h}px) scale(.08) rotate(-30deg)`), opacity: 0 }], opt));
      break;
    case 'origami':
      oldOnTop();
      anims.push(old.animate([{ transform: base }, { transform: C('rotateX(70deg) scale(.6)'), offset: 0.5 }, { transform: C(`rotateX(80deg) rotateZ(45deg) translate(0px, ${-h}px) scale(.1)`), opacity: 0 }], opt));
      break;
    case 'curtains': {
      old.style.visibility = 'hidden';
      const l = ghost('inset(0 50% 0 0)');
      const r = ghost('inset(0 0 0 50%)');
      anims.push(l.animate([{ transform: around(0, h / 2, 'scaleX(1) skewY(0deg)') }, { transform: around(0, h / 2, 'scaleX(.05) skewY(6deg)'), opacity: 0.6 }], opt));
      anims.push(r.animate([{ transform: around(w, h / 2, 'scaleX(1) skewY(0deg)') }, { transform: around(w, h / 2, 'scaleX(.05) skewY(-6deg)'), opacity: 0.6 }], opt));
      break;
    }
    case 'drape':
      oldOnTop();
      anims.push(old.animate([{ transform: around(w / 2, 0, 'scaleY(1) skewX(0deg)'), opacity: 1 }, { transform: around(w / 2, 0, 'scaleY(.05) skewX(-20deg)'), opacity: 0.4 }], { ...opt, easing: 'ease-in' }));
      break;
    case 'wind':
      oldOnTop();
      anims.push(old.animate([{ transform: base, filter: 'blur(0)', opacity: 1 }, { transform: `${base} translate(${dir === 'l' ? w : -w}px, -40px) skewX(-20deg)`, filter: 'blur(6px)', opacity: 0 }], { ...opt, easing: 'ease-in' }));
      anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], opt));
      break;
    case 'prestige':
      oldOnTop();
      anims.push(old.animate([{ transform: around(w / 2, h, 'rotateX(0deg)'), opacity: 1 }, { transform: around(w / 2, h, 'rotateX(70deg) scale(.8)'), opacity: 0 }], half));
      anims.push(neu.animate([{ transform: around(w / 2, h, 'rotateX(-50deg) scale(.8)'), opacity: 0 }, { transform: around(w / 2, h, 'rotateX(0deg) scale(1)'), opacity: 1 }], second));
      break;
    case 'fall':
      oldOnTop();
      anims.push(old.animate([{ transform: around(w / 2, h, 'rotateX(0deg)'), filter: 'brightness(1)' }, { transform: around(w / 2, h, 'rotateX(90deg)'), filter: 'brightness(.4)' }], { ...opt, easing: 'ease-in' }));
      break;
    case 'switch': {
      const sx = dir === 'l' ? -1 : 1;
      oldOnTop();
      anims.push(old.animate([{ transform: C('translateX(0px) rotateY(0deg) scale(1)'), opacity: 1 }, { transform: C(`translateX(${-sx * w * 0.6}px) rotateY(${sx * 50}deg) scale(.6)`), opacity: 0.5, offset: 0.5 }, { transform: C(`translateX(0px) rotateY(0deg) translateZ(-600px)`), opacity: 0 }], opt));
      anims.push(neu.animate([{ transform: C(`translateX(${sx * w * 0.6}px) rotateY(${-sx * 50}deg) scale(.6)`), opacity: 0.4 }, { transform: C('translateX(0px) rotateY(0deg) scale(1)'), opacity: 1 }], second));
      break;
    }
    case 'gallery': case 'conveyor': {
      const sx = dir === 'l' ? -1 : 1;
      const deg = tr.type === 'gallery' ? 35 : 15;
      anims.push(old.animate([{ transform: C('translateX(0px) rotateY(0deg) scale(1)') }, { transform: C(`translateX(${-sx * w * 1.05}px) rotateY(${sx * deg}deg) scale(.8)`) }], opt));
      anims.push(neu.animate([{ transform: C(`translateX(${sx * w * 1.05}px) rotateY(${-sx * deg}deg) scale(.8)`) }, { transform: C('translateX(0px) rotateY(0deg) scale(1)') }], opt));
      break;
    }
    case 'cube': case 'box': case 'rotate': {
      const sx = dir === 'l' ? -1 : 1;
      const depth = tr.type === 'box' ? 0.75 : 1;
      if (tr.type === 'rotate') {
        anims.push(old.animate([{ transform: C('rotateY(0deg)') }, { transform: C(`rotateY(${sx * 90}deg)`) }], half));
        anims.push(neu.animate([{ transform: C(`rotateY(${-sx * 90}deg)`) }, { transform: C('rotateY(0deg)') }], second));
        break;
      }
      anims.push(old.animate([{ transform: around(sx > 0 ? 0 : w, h / 2, 'translateX(0px) rotateY(0deg) scale(1)') }, { transform: around(sx > 0 ? 0 : w, h / 2, `translateX(${-sx * w}px) rotateY(${-sx * 90}deg) scale(${depth})`) }], opt));
      anims.push(neu.animate([{ transform: around(sx > 0 ? 0 : w, h / 2, `translateX(${sx * w}px) rotateY(${sx * 90}deg) scale(${depth})`) }, { transform: around(sx > 0 ? 0 : w, h / 2, 'translateX(0px) rotateY(0deg) scale(1)') }], opt));
      break;
    }
    case 'doors': case 'window': {
      old.style.visibility = 'hidden';
      const l = ghost('inset(0 50% 0 0)');
      const r = ghost('inset(0 0 0 50%)');
      const deg = tr.type === 'doors' ? 100 : 80;
      anims.push(l.animate([{ transform: around(0, h / 2, 'rotateY(0deg)') }, { transform: around(0, h / 2, `rotateY(${tr.type === 'doors' ? deg : -deg}deg)`), opacity: 0 }], { ...opt, easing: 'ease-in' }));
      anims.push(r.animate([{ transform: around(w, h / 2, 'rotateY(0deg)') }, { transform: around(w, h / 2, `rotateY(${tr.type === 'doors' ? -deg : deg}deg)`), opacity: 0 }], { ...opt, easing: 'ease-in' }));
      anims.push(neu.animate([{ transform: C('scale(.85)'), opacity: 0.4 }, { transform: C('scale(1)'), opacity: 1 }], opt));
      break;
    }
    case 'pan': {
      const [ox, oy] = { b: [0, h], t: [0, -h], l: [-w, 0], r: [w, 0] }[dir] ?? [0, h];
      anims.push(neu.animate([{ transform: C(`translate(${ox}px, ${oy}px) scale(.85)`) }, { transform: C(`translate(${ox * 0.5}px, ${oy * 0.5}px) scale(.85)`), offset: 0.5 }, { transform: C('translate(0px, 0px) scale(1)') }], opt));
      anims.push(old.animate([{ transform: C('translate(0px, 0px) scale(1)') }, { transform: C(`translate(${-ox * 0.5}px, ${-oy * 0.5}px) scale(.85)`), offset: 0.5 }, { transform: C(`translate(${-ox}px, ${-oy}px) scale(.85)`) }], opt));
      break;
    }
    case 'ferris': {
      const sx = dir === 'l' ? -1 : 1;
      anims.push(old.animate([{ transform: around(w / 2, h * 3, 'rotate(0deg)') }, { transform: around(w / 2, h * 3, `rotate(${-sx * 25}deg)`) }], opt));
      anims.push(neu.animate([{ transform: around(w / 2, h * 3, `rotate(${sx * 25}deg)`) }, { transform: around(w / 2, h * 3, 'rotate(0deg)') }], opt));
      break;
    }
    case 'orbit': {
      const sx = dir === 'l' ? -1 : 1;
      anims.push(old.animate([{ transform: C('rotateY(0deg) translateZ(0px)'), opacity: 1 }, { transform: C(`rotateY(${sx * 60}deg) translateZ(-${w}px)`), opacity: 0 }], opt));
      anims.push(neu.animate([{ transform: C(`rotateY(${-sx * 60}deg) translateZ(-${w}px)`), opacity: 0 }, { transform: C('rotateY(0deg) translateZ(0px)'), opacity: 1 }], opt));
      break;
    }
    case 'flythrough':
      oldOnTop();
      anims.push(old.animate([{ transform: C('scale(1)'), opacity: 1 }, { transform: C('scale(2.5)'), opacity: 0 }], { ...opt, easing: 'ease-in' }));
      anims.push(neu.animate([{ transform: C('scale(.3)'), opacity: 0 }, { transform: C('scale(1)'), opacity: 1 }], opt));
      break;
    case 'random': {
      const pool = ['fade', 'push', 'wipe', 'split', 'blinds', 'checker', 'dissolve', 'diamond', 'clock', 'cube', 'gallery', 'vortex', 'ripple', 'flythrough'];
      return playTransition(old, neu, { ...tr, type: pool[Math.floor(Math.random() * pool.length)] }, ctx);
    }
    case 'fade': anims.push(neu.animate([{ opacity: 0 }, { opacity: 1 }], opt)); break;
    case 'dissolve': anims.push(neu.animate([{ opacity: 0, filter: 'blur(6px)' }, { opacity: 1, filter: 'blur(0)' }], opt)); break;
    case 'push':
      anims.push(neu.animate([{ transform: T(off[0], off[1]) }, { transform: T(0, 0) }], opt));
      anims.push(old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], opt));
      break;
    case 'cover': anims.push(neu.animate([{ transform: T(off[0], off[1]) }, { transform: T(0, 0) }], opt)); break;
    case 'uncover':
      ctx.raise(old); // 이전 슬라이드를 위로 올려서 밀어냄
      anims.push(old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], opt));
      break;
    case 'wipe': {
      const from = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[dir] ?? 'inset(0 0 0 100%)';
      anims.push(neu.animate([{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], opt));
      break;
    }
    case 'split': anims.push(neu.animate(vert ? [{ clipPath: 'inset(50% 0 50% 0)' }, { clipPath: 'inset(0% 0 0% 0)' }] : [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], opt)); break;
    case 'zoom': anims.push(neu.animate([{ transform: `${base} scale(.3)`, opacity: 0, transformOrigin: '50% 50%' }, { transform: base, opacity: 1 }], opt)); break;
    case 'circle': anims.push(neu.animate([{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }], opt)); break;
    case 'morph': {
      // 모핑: 짝이 되는 개체는 앞 슬라이드 자리 · 크기 · 회전에서 새 자리로, 나머지는 서서히
      const prevSlide = ctx.prev ?? ctx.cur;
      const curSlide = ctx.cur;
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
  const done = Promise.all(anims.map((a) => a.finished.catch(() => {}))).then(() => {
    for (const a of anims) { try { a.cancel(); } catch { /* 끝남 */ } }
    for (const g of ghosts) g.remove();
    old.style.visibility = '';
  });
  return { anims, done };
}

// ───────────── 개체 애니메이션 ─────────────
const clipKf = (list) => list.map((clipPath) => ({ clipPath }));
/** 강조 효과 중 끝난 뒤에도 그대로 남는 것 (PowerPoint: 크게/작게, 투명, 어둡게 …) */
const EMPH_HOLD = new Set(['growShrink', 'desaturate', 'darken', 'lighten', 'transparency', 'objectColor', 'complementary', 'lineColor', 'fillColor', 'brushColor', 'bold', 'underline']);
export function keyframes(a, o, size = S.pres.size) {
  const { w, h } = size;
  const dir = a.dir ?? 'b';
  const fly = { b: `0 ${h - o.y + 10}px`, t: `0 ${-(o.y + o.h + 10)}px`, l: `${-(o.x + o.w + 10)}px 0`, r: `${w - o.x + 10}px 0` }[dir];
  const wipeFrom = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[dir];
  const vert = dir === 'l' || dir === 'r';
  switch (a.effect) {
    case 'appear': case 'disappear': return null;
    case 'fade': return [{ opacity: 0 }, { opacity: 1 }];
    case 'fly': return [{ translate: fly }, { translate: '0 0' }];
    case 'crawl': case 'credits': return [{ translate: fly, easing: 'linear' }, { translate: '0 0' }];
    case 'float': return [{ translate: '0 40px', opacity: 0 }, { translate: '0 0', opacity: 1 }];
    case 'split': return [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }];
    case 'wipe': return [{ clipPath: wipeFrom }, { clipPath: 'inset(0 0 0 0)' }];
    case 'zoom': return [{ scale: '0.3', opacity: 0 }, { scale: '1', opacity: 1 }];
    case 'basicZoom': return [{ scale: '0' }, { scale: '1' }];
    case 'grow': return [{ scale: '0', rotate: '90deg', opacity: 0 }, { scale: '1', rotate: '0deg', opacity: 1 }];
    case 'shape': case 'circleIn': return [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }];
    case 'box': return [{ clipPath: 'inset(50% 50% 50% 50%)' }, { clipPath: 'inset(0% 0% 0% 0%)' }];
    case 'wheel': return clipKf(sweepFrames(16, 'clock', 1));
    case 'wedge': return clipKf(sweepFrames(16, 'wedge'));
    case 'diamond': return clipKf(frames(diamondPoly, 2));
    case 'plus': return clipKf(frames(plusPoly, 2));
    case 'strips': return clipKf(frames((t) => stripsPoly(t), 2));
    case 'blinds': return clipKf(frames((t) => blindsPoly(t, 6, vert), 2));
    case 'checker': return clipKf(frames((t) => checkerPoly(t, 8, 6), 2));
    case 'randomBars': return clipKf(randomBarsFrames(12, 40, vert));
    case 'dissolveIn': return clipKf(randomCellsFrames(12, 12, 8));
    case 'peek': return [{ translate: `0 ${o.h}px`, clipPath: 'inset(0 0 100% 0)' }, { translate: '0 0', clipPath: 'inset(0 0 0% 0)' }];
    case 'swivel': return [{ rotate: 'y 630deg', opacity: 0 }, { rotate: 'y 0deg', opacity: 1 }];
    case 'flip': return [{ rotate: 'x 90deg', opacity: 0 }, { rotate: 'x 0deg', opacity: 1 }];
    case 'expand': return [{ scale: '0.7 1', opacity: 0 }, { scale: '1 1', opacity: 1 }];
    case 'compress': return [{ scale: '1.5 1', opacity: 0 }, { scale: '1 1', opacity: 1 }];
    case 'stretch': return [{ scale: '0 1' }, { scale: '1 1' }];
    case 'riseUp': return [{ translate: '0 120px', opacity: 0 }, { translate: '0 -12px', opacity: 1, offset: 0.8 }, { translate: '0 0' }];
    case 'sinkDown': return [{ translate: '0 -120px', opacity: 0 }, { translate: '0 10px', opacity: 1, offset: 0.8 }, { translate: '0 0' }];
    case 'centerRevolve': return [{ translate: '0 90px', scale: '0.3', opacity: 0 }, { translate: '60px 40px', scale: '0.7', opacity: 0.7 }, { translate: '0 0', scale: '1', opacity: 1 }];
    case 'pinwheel': return [{ scale: '0', rotate: '720deg', opacity: 0 }, { scale: '1', rotate: '0deg', opacity: 1 }];
    case 'boomerang': return [{ translate: `${Math.round(w * 0.3)}px -80px`, rotate: 'y 90deg', scale: '0.3', opacity: 0 }, { translate: '-60px 30px', rotate: 'y 30deg', scale: '0.8', opacity: 1 }, { translate: '0 0', rotate: 'y 0deg', scale: '1' }];
    case 'curveUp': return [{ translate: '-220px 220px', scale: '0.6', opacity: 0 }, { translate: '-60px -40px', scale: '1.05', opacity: 1 }, { translate: '0 0', scale: '1' }];
    case 'spiralIn': {
      const out = [];
      for (let k = 0; k <= 12; k++) {
        const t = k / 12; const r = (1 - t) * Math.max(w, h) * 0.6; const ang = t * Math.PI * 3;
        out.push({ translate: `${Math.round(r * Math.cos(ang))}px ${Math.round(r * Math.sin(ang))}px`, scale: String(0.1 + 0.9 * t), opacity: Math.min(1, t * 3) });
      }
      return out;
    }
    case 'dropIn': return [{ translate: `0 ${-(o.y + o.h)}px`, opacity: 0, easing: 'ease-in' }, { translate: '0 0', opacity: 1, offset: 0.65, easing: 'ease-out' }, { translate: '0 -16px', offset: 0.82, easing: 'ease-in' }, { translate: '0 0' }];
    case 'bounce': return [{ translate: `0 ${-(o.y + o.h)}px`, offset: 0 }, { translate: '0 0', offset: 0.55, easing: 'ease-in' }, { translate: '0 -30px', offset: 0.75, easing: 'ease-out' }, { translate: '0 0', offset: 0.9 }, { translate: '0 -8px', offset: 0.95 }, { translate: '0 0', offset: 1 }];
    // 강조
    case 'pulse': return [{ scale: '1' }, { scale: '1.08' }, { scale: '1' }];
    case 'colorPulse': return [{ filter: 'none' }, { filter: 'brightness(1.35) saturate(2.2) hue-rotate(25deg)' }, { filter: 'none' }];
    case 'spin': return [{ rotate: '0deg' }, { rotate: '360deg' }];
    case 'growShrink': return [{ scale: '1' }, { scale: '1.5' }];
    case 'teeter': return [{ rotate: '0deg' }, { rotate: '6deg' }, { rotate: '-6deg' }, { rotate: '4deg' }, { rotate: '0deg' }];
    case 'flash': return [{ opacity: 1 }, { opacity: 0 }, { opacity: 1 }];
    case 'boldFlash': return [{ filter: 'none', scale: '1' }, { filter: 'contrast(1.6) brightness(.75)', scale: '1.03' }, { filter: 'none', scale: '1' }];
    case 'shimmer': return [{ opacity: 1, scale: '1' }, { opacity: 0.45, scale: '0.97' }, { opacity: 1, scale: '1' }];
    case 'wave': return [{ translate: '0 0', rotate: '0deg' }, { translate: '0 -10px', rotate: '-3deg' }, { translate: '0 8px', rotate: '3deg' }, { translate: '0 -4px', rotate: '-1deg' }, { translate: '0 0', rotate: '0deg' }];
    case 'desaturate': return [{ filter: 'grayscale(0)' }, { filter: 'grayscale(1)' }];
    case 'darken': return [{ filter: 'brightness(1)' }, { filter: 'brightness(.55)' }];
    case 'lighten': return [{ filter: 'brightness(1)' }, { filter: 'brightness(1.45)' }];
    case 'transparency': return [{ opacity: 1 }, { opacity: 0.5 }];
    case 'objectColor': case 'fillColor': case 'brushColor': return [{ filter: 'hue-rotate(0deg)' }, { filter: 'hue-rotate(120deg)' }];
    case 'lineColor': return [{ filter: 'hue-rotate(0deg)' }, { filter: 'hue-rotate(60deg)' }];
    case 'complementary': return [{ filter: 'hue-rotate(0deg)' }, { filter: 'hue-rotate(180deg)' }];
    case 'bold': return [{ fontWeight: 400 }, { fontWeight: 700 }];
    case 'underline': return [{ textDecorationLine: 'none' }, { textDecorationLine: 'underline' }];
    default: return [{ opacity: 0 }, { opacity: 1 }];
  }
}

/** 끝내기 = 나타내기를 거꾸로 (offset 도 1 - offset, easing 은 앞 구간 것으로) */
export function reverseKf(kf) {
  const out = kf.map((k) => ({ ...k })).reverse();
  for (const k of out) if (k.offset != null) k.offset = 1 - k.offset;
  for (let i = 0; i < out.length; i++) { delete out[i].easing; const src = kf[kf.length - 2 - i]; if (src?.easing) out[i].easing = src.easing; }
  return out;
}

/** 글을 단어 · 문자 단위 span 으로 나눔 (PowerPoint [텍스트 애니메이션: 단어별 / 문자별]) */
export function splitTextUnits(elm, unit) {
  const units = [];
  const walk = (node) => {
    for (const c of [...node.childNodes]) {
      if (c.nodeType === 3) {
        const parts = unit === 'letter' ? [...c.textContent] : c.textContent.split(/(\s+)/);
        const frag = document.createDocumentFragment();
        for (const t of parts) {
          if (!t) continue;
          if (/^\s+$/.test(t)) { frag.append(t); continue; }
          const sp = document.createElement('span');
          sp.className = 'tu';
          sp.textContent = t;
          frag.append(sp);
          units.push(sp);
        }
        c.replaceWith(frag);
      } else if (c.nodeType === 1 && !c.classList.contains('tu')) walk(c);
    }
  };
  for (const p of elm.matches('.p') ? [elm] : elm.querySelectorAll('.p')) walk(p);
  return units;
}
function animateUnits(elm, units, a, o) {
  const kf0 = keyframes(a, o);
  const kf = kf0 ?? [{ opacity: 0 }, { opacity: 0, offset: 0.999 }, { opacity: 1 }];
  const dur = Math.max(1, (a.dur ?? 0.5) * 1000);
  const gap = dur * (a.unitDelay ?? 10) / 100;
  const easing = a.cls === 'exit' ? 'ease-in' : a.cls === 'emph' ? 'ease-in-out' : 'ease-out';
  const frames = a.cls === 'exit' ? reverseKf(kf) : kf;
  const fill = a.cls === 'entr' ? 'backwards' : a.cls === 'exit' || EMPH_HOLD.has(a.effect) ? 'both' : 'none';
  elm.style.visibility = 'visible';
  const anims = units.map((u, i) => u.animate(frames, { duration: dur, delay: i * gap, easing, fill }));
  const all = Promise.all(anims.map((x) => x.finished.catch(() => {})));
  if (a.cls === 'exit') return all.then(() => { elm.style.visibility = 'hidden'; for (const x of anims) x.cancel(); });
  return all;
}

/** 개체 하나 재생 (편집 화면 미리 보기에서도 씀) */
export function animateObject(elm, a, o) {
  if (a.textUnit && a.cls !== 'path' && a.para == null) {
    const units = splitTextUnits(elm, a.textUnit);
    if (units.length > 1) return animateUnits(elm, units, a, o);
  }
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
    const anim = elm.animate(reverseKf(kf), { duration: dur, easing: 'ease-in', fill: 'forwards' });
    return anim.finished.then(() => { elm.style.visibility = 'hidden'; anim.cancel(); }).catch(() => {});
  }
  const fillMode = EMPH_HOLD.has(a.effect) ? 'forwards' : 'none';
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
    done.push(new Promise((res) => setTimeout(() => { if (anim.sound) playSound(anim.sound); animateObject(elm, anim, o).then(res); }, at * 1000)));
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

// ───────────── 확대/축소 ─────────────
function zoomInto(zo) {
  const idx = S.pres.slides.findIndex((x) => x.id === zo.zoom.slide);
  const pos = show.order.indexOf(idx);
  if (pos < 0) return;
  // 구역 확대/축소: 다음 구역 시작 전까지 (쇼 순서 기준)
  let end = pos;
  if (zo.zoom.section) while (end + 1 < show.order.length && !S.pres.slides[show.order[end + 1]].section) end++;
  const from = show.pos;
  show.zoomBack = zo.zoom.ret !== false ? { pos: from, end, rect: zo } : null;
  enterSlide(pos, { transition: false });
  zoomAnim(show.stage.querySelector('.show-slide:last-child'), zo, false);
}
/** 축소판 사각형 ⇄ 전체 화면 (Web Animations) */
function zoomAnim(elm, r, out) {
  if (!elm) return;
  const { w, h } = S.pres.size;
  const base = elm.style.transform;
  const small = `${base} translate(${r.x}px, ${r.y}px) scale(${r.w / w}, ${r.h / h})`;
  const kf = [{ transform: small, opacity: 0.6 }, { transform: base, opacity: 1 }];
  elm.style.transformOrigin = '0 0';
  elm.animate(out ? [...kf].reverse() : kf, { duration: 700, easing: 'ease-in-out' });
}
function zoomReturn() {
  const zb = show.zoomBack;
  show.zoomBack = null;
  const cur = show.stage.querySelector('.show-slide:last-child');
  const done = () => { if (!show) return; enterSlide(zb.pos, { transition: false, stepsDone: 999, back: true }); };
  if (!cur) { done(); return; }
  const { w, h } = S.pres.size;
  const base = cur.style.transform;
  cur.style.transformOrigin = '0 0';
  cur.animate([{ transform: base, opacity: 1 }, { transform: `${base} translate(${zb.rect.x}px, ${zb.rect.y}px) scale(${zb.rect.w / w}, ${zb.rect.h / h})`, opacity: 0.4 }], { duration: 600, easing: 'ease-in-out', fill: 'forwards' }).finished.then(done, done);
}

// ───────────── 이동 ─────────────
export function next() {
  if (!show) return;
  if (show.ended) { endShow(); return; }
  if (show.blank) { setBlank(null); return; }
  if (show.step < show.steps.length) { playStep(); return; }
  if (show.zoomBack && show.pos >= show.zoomBack.end) { zoomReturn(); return; }
  if (show.pos + 1 < show.order.length) { enterSlide(show.pos + 1); return; }
  if (show.setup.loop || kiosk()) { enterSlide(0); return; }
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
  syncAudience();
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
  if (k === 'j' || k === 'J') { toggleCaptions(); return; }
  if (k === 'h' || k === 'H') { e.preventDefault(); next(); return; }
  switch (k) {
    case 'Escape': case '-': e.preventDefault(); if (show.tool) setTool(null); else endShow(); break;
    case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': case 'N': e.preventDefault(); if (!kiosk()) next(); break;
    case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P': e.preventDefault(); if (!kiosk()) prev(); break;
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
  // 확대/축소: 축소판에서 대상 슬라이드로 커지며 이동, 다 본 뒤 (ret) 다시 축소판 슬라이드로
  const zEl = e.target.closest('.show-slide:last-child .ob[data-id]');
  const zo = zEl && S.pres.slides[show.order[show.pos]].objects.find((o) => o.id === zEl.dataset.id && o.type === 'zoom');
  if (zo) { zoomInto(zo); return; }
  // 트리거: 누른 개체에 걸린 애니메이션 묶음을 차례로 (모두 끝나면 다시 처음부터)
  const tob = e.target.closest('.show-slide:last-child .ob[data-id]');
  const tsteps = tob && show.triggers?.get(tob.dataset.id);
  if (tsteps) {
    const t = tob.dataset.id;
    let n = show.trigDone.get(t) ?? 0;
    if (n >= tsteps.length) { n = 0; show.trigDone.set(t, 0); rerender(); }
    show.trigDone.set(t, n + 1);
    const slide = S.pres.slides[show.order[show.pos]];
    playStepOn(show.stage.querySelector('.show-slide:last-child'), tsteps[n], slide);
    syncAudience();
    return;
  }
  const link = e.target.closest('a.lnk, [data-link]');
  if (link) {
    const href = link.getAttribute('href') ?? link.dataset.link;
    const m = /^#slide(\d+)$/.exec(href ?? '');
    e.preventDefault();
    if (m) { const pos = show.order.indexOf(Number(m[1]) - 1); if (pos >= 0) goTo(pos); } else if (href) window.open(href, '_blank', 'noopener');
    return;
  }
  if (!kiosk()) next();
}
/** 대화형 자동 진행 (키오스크): 클릭 · 키로 넘기지 않음 (링크 · 트리거만), 설정된 시간으로 반복 */
const kiosk = () => show?.setup?.type === 'kiosk';

function showMenu(e) {
  openMenu({ x: e.clientX, y: e.clientY }, [
    { label: '다음', key: 'N', action: next },
    { label: '이전', key: 'P', action: prev },
    { label: '슬라이드로 이동', submenu: show.order.map((idx, pos) => ({ label: `${idx + 1}  ${slideTitle(S.pres.slides[idx]) || '(제목 없음)'}`, checked: pos === show.pos, action: () => goTo(pos) })) },
    { sep: true },
    { label: '레이저 포인터', key: 'Ctrl+L', checked: show.tool === 'laser', action: () => setTool(show.tool === 'laser' ? null : 'laser') },
    { label: '펜', key: 'Ctrl+P', checked: show.tool === 'pen', action: () => setTool(show.tool === 'pen' ? null : 'pen') },
    { label: '잉크 지우기', key: 'E', action: clearInk },
    { label: '자막 켜기/끄기', key: 'J', checked: !!show.cap?.on, action: toggleCaptions },
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

/** 쇼를 마칠 때: 잉크 주석을 슬라이드에 자유형 도형으로 남길지 묻기 (PowerPoint 와 같음) */
function keepInkDialog(ink) {
  if (S.viewOnly) return;
  import_openDialog({
    title: 'Microsoft PowerPoint',
    body: el('p', {}, '잉크 주석을 유지하시겠습니까?'),
    buttons: [{ label: '유지', primary: true, action: () => {
      change(() => {
        for (const [idx, list] of ink) {
          const slide = S.pres.slides[idx];
          if (!slide) continue;
          for (const st of list) {
            if (st.pts.length < 2) continue;
            const xs = st.pts.map((q) => q[0]);
            const ys = st.pts.map((q) => q[1]);
            const x = Math.min(...xs); const y = Math.min(...ys);
            const w = Math.max(1, Math.max(...xs) - x); const h = Math.max(1, Math.max(...ys) - y);
            const cmds = st.pts.map(([px, py], i) => [i ? 'L' : 'M', Math.round((px - x) * 10) / 10, Math.round((py - y) * 10) / 10]);
            slide.objects.push({ id: uid(), type: 'shape', shape: 'rect', name: '잉크', ink: true, x, y, w, h, rot: 0, fill: null, line: { color: st.color, width: Math.max(1, st.width * 1.5), dash: 'solid' }, path: [{ w, h, cmds }], openPath: true });
          }
        }
      }, { scope: 'all' });
    } }, { label: '취소' }],
  });
}

// ───────────── 펜 · 레이저 ─────────────
function setTool(t) {
  show.tool = t;
  show.root.classList.toggle('pen', t === 'pen');
  show.root.classList.toggle('laser-on', t === 'laser');
  show.laser.hidden = t !== 'laser';
}
function clearInk(forget = true) {
  const c = show?.ink;
  if (c) c.getContext('2d').clearRect(0, 0, c.width, c.height);
  if (forget && show) show.strokes?.delete(show.order[show.pos]);
}
/** 화면 좌표 ↔ 슬라이드 좌표 */
function toSlidePt(x, y) {
  const r = show.view.getBoundingClientRect();
  const { w, h } = S.pres.size;
  const sc = show.scale || 1;
  return [(x - r.left - (r.width - w * sc) / 2) / sc, (y - r.top - (r.height - h * sc) / 2) / sc];
}
/** 이 슬라이드에 남긴 잉크 다시 그리기 (슬라이드를 오갈 때) */
function redrawInk() {
  if (!show) return;
  const c = show.ink;
  const g = c.getContext('2d');
  g.clearRect(0, 0, c.width, c.height);
  const list = [...(show.replay?.shown ?? []), ...(show.strokes?.get(show.order[show.pos]) ?? [])];
  const r = show.view.getBoundingClientRect();
  const { w, h } = S.pres.size;
  const sc = show.scale || 1;
  const ox = (r.width - w * sc) / 2;
  const oy = (r.height - h * sc) / 2;
  for (const st of list) {
    g.strokeStyle = st.color;
    g.lineWidth = st.width * sc;
    g.lineCap = 'round';
    g.lineJoin = 'round';
    g.beginPath();
    st.pts.forEach(([x, y], i) => (i ? g.lineTo(ox + x * sc, oy + y * sc) : g.moveTo(ox + x * sc, oy + y * sc)));
    g.stroke();
  }
}
let drawing = null;
let stroke = null;
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
      // 슬라이드 좌표로도 기록 (쇼를 마칠 때 잉크를 도형으로 남길 수 있게)
      if (!stroke) {
        stroke = { color: show.penColor ?? '#E81123', width: 3 / (show.scale || 1), pts: [] };
        show.strokes ??= new Map();
        const key = show.order[show.pos];
        if (!show.strokes.has(key)) show.strokes.set(key, []);
        show.strokes.get(key).push(stroke);
      }
      stroke.pts.push(toSlidePt(e.clientX, e.clientY));
      if (show.rec?.cur) { stroke.ts ??= []; stroke.ts.push(Math.round(recElapsed() * 100) / 100); }
      if (drawing) {
        ctx.strokeStyle = show.penColor ?? '#e81123';
        ctx.lineWidth = 3;
        ctx.lineCap = 'round';
        ctx.beginPath();
        ctx.moveTo(...drawing);
        ctx.lineTo(...p);
        ctx.stroke();
      }
      drawing = p;
    } else { drawing = null; stroke = null; }
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
    el('button', { title: '타이머 일시 중지 / 다시 시작', onclick: (e) => { if (show.paused) { show.started += Date.now() - show.paused; show.paused = null; e.target.textContent = '⏸ 일시 중지'; } else { show.paused = Date.now(); e.target.textContent = '▶ 계속'; } } }, '⏸ 일시 중지'),
    el('button', { onclick: () => { show.started = Date.now(); if (show.paused) show.paused = Date.now(); } }, '타이머 다시'),
    el('button', { title: '펜 (Ctrl+P)', onclick: () => setTool(show.tool === 'pen' ? null : 'pen') }, '✎ 펜'),
    el('button', { title: '레이저 포인터 (Ctrl+L)', onclick: () => setTool(show.tool === 'laser' ? null : 'laser') }, '● 레이저'),
    el('button', { title: '모든 슬라이드 보기 (G)', onclick: () => slideList() }, '▦ 모든 슬라이드'),
    el('button', { title: '검은 화면 (B)', onclick: () => toggleCurtain('#000') }, '■ 검은 화면'),
    el('button', { title: '흰 화면 (W)', onclick: () => toggleCurtain('#fff') }, '□ 흰 화면'),
    el('button', { title: '발표자 보기와 슬라이드 쇼 화면 바꾸기', onclick: () => swapDisplays() }, '⇄ 화면 바꾸기'),
    el('button', { onclick: () => openAudience() }, '청중 창 열기'),
    el('button', { onclick: endShow }, '쇼 마침'));
  side.append(timer, el('div', { class: 'pv-label' }, '다음 슬라이드'), nextBox, el('div', { class: 'pv-label' }, '메모'), notes);
  show.root.append(side, tools);
  show.presenterEls = { side, nextBox, notes, timer, counter };
  show.clock = setInterval(updateTimer, 1000);
}
/** 발표자 보기 · 슬라이드 쇼 모니터 바꾸기 (모니터 2대 이상) */
async function swapDisplays() {
  const sd = await screenList(true);
  if (!sd || sd.screens.length < 2) { toast('모니터가 한 대입니다 — [청중 창 열기]로 창을 띄워 옮기세요'); return; }
  const cur = sd.currentScreen;
  const other = sd.screens.find((x) => x !== cur);
  show.audience?.close?.();
  show.root.requestFullscreen?.({ screen: other }).catch(() => {});
  openAudience(cur);
}
const toggleCurtain = (c) => setBlank(show.blank === c ? null : c);
function updateTimer() {
  if (!show?.presenterEls) { clearInterval(show?.clock); return; }
  const s = Math.floor(((show.paused ?? Date.now()) - show.started) / 1000);
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
function openAudience(scr) {
  // scr: 보여 줄 모니터 (Window Management API) — 그 화면 전체를 덮는 창 (+ 가능하면 전체 화면)
  const geo = scr && typeof scr === 'object' && 'availLeft' in scr ? `,left=${scr.left},top=${scr.top},width=${scr.width},height=${scr.height},fullscreen` : ',width=960,height=560';
  const w = window.open('', 'wipoint-audience', `popup${geo}`);
  if (!w) { toast('팝업이 차단되었습니다. 브라우저에서 팝업을 허용해 주세요.'); return; }
  w.document.open();
  w.document.write(`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>WIPOINT - 청중 화면 (두 번 클릭: 전체 화면)</title><style>${SLIDE_CSS}html,body{margin:0;height:100%;background:#000;overflow:hidden}#v{position:absolute;inset:0}.show-slide{position:absolute;left:0;top:0;transform-origin:0 0}.show-end{color:#fff;font:20px sans-serif;padding:20px}#cap{position:absolute;left:5%;right:5%;text-align:center;color:#fff;background:rgba(0,0,0,.78);font:500 28px/1.35 'Malgun Gothic',sans-serif;padding:.25em .6em;border-radius:4px;z-index:5}#cap.cap-bottom{bottom:4%}#cap.cap-top{top:4%}</style></head><body><div id="v"></div><div id="cap" hidden></div></body></html>`);
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
  if (show.blank) { v.innerHTML = `<div style="position:absolute;inset:0;background:${show.blank}"></div>`; return; }
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

// ───────────── 쇼 녹화 (PowerPoint [녹화]: 내레이션 · 카메라 · 타이밍 · 잉크) ─────────────
/** 오디오 형식: PowerPoint 가 재생하는 MP4(AAC) 를 먼저, 안 되면 WebM 으로 받아서 WAV 로 바꿈 */
function recMime(video) {
  const list = video ? ['video/mp4;codecs=avc1,mp4a.40.2', 'video/mp4', 'video/webm;codecs=vp9,opus', 'video/webm'] : ['audio/mp4;codecs=mp4a.40.2', 'audio/mp4', 'audio/webm;codecs=opus', 'audio/webm'];
  return list.find((m) => window.MediaRecorder?.isTypeSupported?.(m)) ?? '';
}
function startRecording(stream, o) {
  const video = !!stream?.getVideoTracks?.().length;
  const dot = el('span', { class: 'rec-dot' });
  const time = el('span', { class: 'rec-time' }, '00:00');
  const pauseBtn = el('button', { title: '녹화 일시 중지', onclick: (e) => { e.stopPropagation(); recPause(); } }, '⏸ 일시 중지');
  const bar = el('div', { class: 'rec-bar' }, dot, el('span', {}, '녹화 중'), time, pauseBtn, el('button', { title: '녹화 중지 (쇼 마침)', onclick: (e) => { e.stopPropagation(); endShow(); } }, '■ 중지'));
  let cam = null;
  if (video) { cam = el('video', { class: 'rec-cam', muted: true, autoplay: true, playsInline: true }); cam.srcObject = stream; cam.muted = true; }
  show.root.append(bar);
  if (cam) show.view.append(cam);
  show.rec = { stream, video, mime: stream ? recMime(video) : '', segs: new Map(), cur: null, paused: null, bar, pauseBtn, time, started: Date.now() };
  show.rec.tick = setInterval(() => {
    const r = show?.rec;
    if (!r) return;
    const t = Math.floor(recElapsed(true));
    r.time.textContent = `${String(Math.floor(t / 60)).padStart(2, '0')}:${String(t % 60).padStart(2, '0')}`;
  }, 500);
  return o;
}
/** 지금 슬라이드에서 녹화된 시간(초, 일시 중지 뺌) — total: 처음부터 */
function recElapsed(total = false) {
  const r = show?.rec;
  const c = r?.cur;
  if (!c) return 0;
  const now = r.paused ?? Date.now();
  if (total) return [...r.segs.values()].reduce((a, sg) => a + sg.dur, 0) + (now - c.start - c.pausedMs) / 1000;
  return (now - c.start - c.pausedMs) / 1000;
}
function recPause() {
  const r = show.rec;
  if (!r.paused) { r.paused = Date.now(); r.cur?.mr?.state === 'recording' && r.cur.mr.pause(); r.pauseBtn.textContent = '● 다시 시작'; r.bar.classList.add('paused'); } else { if (r.cur) r.cur.pausedMs += Date.now() - r.paused; r.paused = null; r.cur?.mr?.state === 'paused' && r.cur.mr.resume(); r.pauseBtn.textContent = '⏸ 일시 중지'; r.bar.classList.remove('paused'); }
}
/** 슬라이드마다 따로 녹음 (PowerPoint 와 같이 슬라이드별 내레이션 · 시간) */
function recSegment(idx) {
  const r = show.rec;
  recCloseSegment();
  const c = { idx, start: Date.now(), pausedMs: 0, chunks: [], mr: null };
  if (r.paused) { c.pausedMs = 0; r.paused = Date.now(); }
  if (r.stream && window.MediaRecorder) {
    try {
      c.mr = new MediaRecorder(r.stream, r.mime ? { mimeType: r.mime } : undefined);
      c.mr.ondataavailable = (e) => { if (e.data?.size) c.chunks.push(e.data); };
      c.done = new Promise((res) => { c.mr.onstop = () => res(new Blob(c.chunks, { type: c.mr.mimeType || r.mime || 'audio/webm' })); });
      c.mr.start(1000);
      if (r.paused) c.mr.pause();
    } catch { c.mr = null; }
  }
  r.cur = c;
}
function recCloseSegment() {
  const r = show?.rec;
  const c = r?.cur;
  if (!c) return;
  const dur = Math.max(0.1, recElapsed());
  if (c.mr && c.mr.state !== 'inactive') c.mr.stop();
  const ink = (show.strokes?.get(c.idx) ?? []).filter((st) => st.ts?.length);
  // 같은 슬라이드를 다시 녹화하면 마지막 것으로 바꿈 (PowerPoint 와 같음)
  r.segs.set(c.idx, { dur: Math.round(dur * 10) / 10, blob: c.done ?? null, ink: ink.map((st) => ({ color: st.color, width: st.width, pts: st.pts.map(([x, y]) => [Math.round(x * 10) / 10, Math.round(y * 10) / 10]), ts: st.ts })) });
  show.strokes?.delete(c.idx);
  r.cur = null;
}
/** 내레이션 → WAV (MP4 오디오면 그대로) */
async function narrationUrl(blob) {
  if (!blob?.size) return null;
  if (/^audio\/mp4|^video\//.test(blob.type)) return blobUrl(blob);
  try {
    const ac = new (window.AudioContext || window.webkitAudioContext)();
    const buf = await ac.decodeAudioData(await blob.arrayBuffer());
    ac.close?.();
    const chans = Array.from({ length: buf.numberOfChannels }, (_, i) => buf.getChannelData(i));
    const rate = 22050;
    const wav = encodeWav(resample(mixDown(chans), buf.sampleRate, rate), rate);
    return blobUrl(new Blob([wav], { type: 'audio/wav' }));
  } catch { return blobUrl(blob); }
}
async function finishRecording(r) {
  clearInterval(r.tick);
  const saved = show;
  show = show ?? { rec: r };
  recCloseSegment();
  show = saved;
  for (const t of r.stream?.getTracks?.() ?? []) t.stop();
  const segs = [...r.segs];
  if (!segs.length) return;
  toast('녹화를 저장하는 중…');
  const out = [];
  for (const [idx, sg] of segs) out.push([idx, sg, sg.blob ? await narrationUrl(await sg.blob) : null]);
  const { w, h } = S.pres.size;
  change(() => {
    for (const [idx, sg, url] of out) {
      const slide = S.pres.slides[idx];
      if (!slide) continue;
      slide.transition = { ...(slide.transition ?? { type: 'none' }), advAfter: Math.max(0.1, sg.dur) };
      if (sg.ink.length) slide.inkRec = sg.ink; else delete slide.inkRec;
      if (!url) continue;
      slide.objects = slide.objects.filter((o) => !o.narration);
      const id = uid('m');
      S.pres.media[id] = url;
      const cam = r.video;
      // PowerPoint 와 같이 오른쪽 아래: 카메라는 작은 동영상, 내레이션은 쇼에서 보이지 않는 소리 아이콘
      slide.objects.push(cam
        ? { id: uid(), type: 'media', kind: 'video', media: id, name: '녹화', narration: true, autoplay: true, x: w - 256 - 16, y: h - 192 - 16, w: 256, h: 192, rot: 0 }
        : { id: uid(), type: 'media', kind: 'audio', media: id, name: '내레이션', narration: true, autoplay: true, hideInShow: true, x: w - 48 - 12, y: h - 48 - 12, w: 48, h: 48, rot: 0 });
    }
  }, { scope: 'all' });
  const total = segs.reduce((a, [, sg]) => a + sg.dur, 0);
  toast(`녹화를 저장했습니다: 슬라이드 ${segs.length}장, ${Math.floor(total / 60)}분 ${Math.round(total % 60)}초`);
}
/** 녹화된 잉크를 시간에 맞춰 다시 그림 ([설정된 시간 사용] 일 때) */
function replayInk() {
  if (!show) return;
  const slide = S.pres.slides[show.order[show.pos]];
  show.replay = null;
  if (!slide?.inkRec?.length || show.rec || show.setup.useTimings === false) { redrawInk(); return; }
  const t0 = Date.now();
  const rp = { strokes: slide.inkRec, shown: [] };
  show.replay = rp;
  const end = Math.max(...slide.inkRec.map((st) => st.ts?.[st.ts.length - 1] ?? 0));
  const step = () => {
    if (!show || show.replay !== rp) return;
    const t = (Date.now() - t0) / 1000;
    rp.shown = inkUntil(rp.strokes, t);
    redrawInk();
    if (t <= end) requestAnimationFrame(step);
  };
  step();
}

// ───────────── 실시간 자막 (Web Speech API) ─────────────
/** 자막 설정 (이 브라우저에 저장): { on, lang, pos, size } */
export function captionSettings() {
  try { return { on: false, lang: 'ko-KR', target: '', pos: 'bottom', size: 28, ...JSON.parse(localStorage.getItem('wipoint:captions') ?? '{}') }; } catch { return { on: false, lang: 'ko-KR', target: '', pos: 'bottom', size: 28 }; }
}
export function setCaptionSettings(v) {
  try { localStorage.setItem('wipoint:captions', JSON.stringify({ ...captionSettings(), ...v })); } catch { /* 저장 안 됨 */ }
}
export const captionsSupported = () => !!(window.SpeechRecognition || window.webkitSpeechRecognition);
function setupCaptions() {
  const cs = captionSettings();
  show.cap = { on: false, text: '', final: '', rec: null, el: null, cs };
  if (cs.on) startCaptions();
}
function startCaptions() {
  const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SR) { toast('이 브라우저는 음성 인식을 지원하지 않습니다 (Chrome · Edge 에서 사용하세요)'); return; }
  const c = show.cap;
  c.on = true;
  const cs = captionSettings();
  c.cs = cs;
  c.el?.remove();
  c.el = el('div', { class: `show-caption cap-${cs.pos}` });
  c.el.style.setProperty('--cap-size', `${cs.size}px`);
  show.root.classList.toggle('cap-below', cs.pos === 'below');
  show.root.classList.toggle('cap-above', cs.pos === 'above');
  (cs.pos === 'below' || cs.pos === 'above' ? show.root : show.view).append(c.el);
  const rec = new SR();
  rec.lang = cs.lang;
  rec.continuous = true;
  rec.interimResults = true;
  // 번역 자막: 이 기기의 번역 모델 (Chrome Translator API) — 없으면 말한 언어 그대로
  const src = cs.lang.slice(0, 2);
  if (cs.target && cs.target !== src) {
    if (!('Translator' in window)) toast('이 브라우저는 자막 번역을 지원하지 않습니다 (Chrome 138 이상) — 말한 언어로 표시합니다');
    else window.Translator.create({ sourceLanguage: src, targetLanguage: cs.target }).then((t) => { c.tr = t; }, () => toast('번역 모델을 준비하지 못해 말한 언어로 표시합니다'));
  }
  let seq = 0;
  rec.onresult = (e) => {
    let interim = '';
    for (let i = e.resultIndex; i < e.results.length; i++) {
      if (e.results[i].isFinal) c.final += `${e.results[i][0].transcript} `; else interim += e.results[i][0].transcript;
    }
    c.final = c.final.slice(-400);
    const said = captionText(`${c.final}${interim}`);
    if (!c.tr) { c.text = said; showCaption(); return; }
    // 마지막 요청 결과만 보여 줌 (순서가 바뀌어 도착해도)
    const my = ++seq;
    c.tr.translate(said).then((t) => { if (my === seq && show?.cap === c) { c.text = captionText(t); showCaption(); } }, () => { c.text = said; showCaption(); });
  };
  // 잠깐 말이 끊기면 브라우저가 멈추므로 다시 시작
  rec.onend = () => { if (show?.cap === c && c.on) { try { rec.start(); } catch { /* 이미 시작 */ } } };
  rec.onerror = (e) => { if (e.error === 'not-allowed' || e.error === 'service-not-allowed') { c.on = false; toast('마이크 권한이 없어 자막을 켤 수 없습니다'); c.el?.remove(); } };
  c.rec = rec;
  try { rec.start(); } catch { /* 이미 시작 */ }
  fit();
}
function showCaption() {
  const c = show?.cap;
  if (!c?.el) return;
  c.el.textContent = c.text;
  c.el.hidden = !c.text;
  // 청중 창: 슬라이드는 그대로 두고 자막 줄만 바꿈
  const w = show.audience;
  const ce = w && !w.closed ? w.document.getElementById('cap') : null;
  if (ce) { ce.textContent = c.text; ce.hidden = !c.text; ce.className = `cap-${c.cs.pos === 'top' || c.cs.pos === 'above' ? 'top' : 'bottom'}`; ce.style.fontSize = `${c.cs.size}px`; }
}
function stopCaptions(s = show) {
  const c = s?.cap;
  if (!c) return;
  c.on = false;
  try { c.rec?.stop(); } catch { /* 이미 멈춤 */ }
  c.el?.remove();
  c.el = null;
  c.text = '';
  const ce = s.audience && !s.audience.closed ? s.audience.document.getElementById('cap') : null;
  if (ce) ce.hidden = true;
  s.root.classList.remove('cap-below', 'cap-above');
}
function toggleCaptions() {
  if (show.cap?.on) { stopCaptions(); setCaptionSettings({ on: false }); fit(); syncAudience(); } else { setCaptionSettings({ on: true }); startCaptions(); }
}
