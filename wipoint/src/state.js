// 앱 공용 상태 + 명령 등록소 + 변경 알림 (모듈끼리 서로 import 하지 않도록 이곳을 거침)
import { History, newPresentation } from './model.js';

export const S = {
  pres: newPresentation(),
  cur: 0,                 // 현재 슬라이드 번호
  sel: new Set(),         // 선택한 개체 id
  slideSel: new Set(),    // 미리 보기 창에서 고른 슬라이드 id (여러 개)
  editing: null,          // { id, cell?: [r, c] } 글 편집 중
  history: new History(),
  zoom: 1,
  fitZoom: true,
  view: 'normal',         // normal | sorter | reading
  // 휴대폰 세로 화면에서는 노트를 처음엔 숨김 (편집 화면을 넓게)
  showNotes: !(typeof matchMedia === 'function' && matchMedia('(max-width: 760px)').matches),
  showRuler: false,
  showGuides: false,
  showGrid: false,
  formatPane: null,       // null | 'shape' | 'bg' | 'anim' | 'selection'
  docName: '프레젠테이션1',
  dirty: false,
  fileHandle: null,
  painter: null,          // 서식 복사 { style, keep }
  drawShape: null,        // 그릴 도형 종류 (삽입 › 도형)
  clipboard: null,
  opts: {},
};

export const curSlide = () => S.pres.slides[S.cur] ?? S.pres.slides[0];
export const selObjects = () => { const s = curSlide(); return s ? s.objects.filter((o) => S.sel.has(o.id)) : []; };
export const selOne = () => { const l = selObjects(); return l.length === 1 ? l[0] : null; };
export const objById = (id) => curSlide()?.objects.find((o) => o.id === id) ?? null;

// ───────────── 알림 ─────────────
const listeners = new Map();
export function on(ev, fn) { if (!listeners.has(ev)) listeners.set(ev, []); listeners.get(ev).push(fn); }
export function emit(ev, data) { for (const fn of listeners.get(ev) ?? []) { try { fn(data); } catch (e) { console.error(e); } } }

/**
 * 문서 바꾸기 (실행 취소 기록 → 변경 → 다시 그리기)
 * scope: 'slide'(현재 슬라이드) | 'all' | 'sel'(선택 표시만) ; key: 같은 key 연속 변경은 한 번으로 묶음
 */
export function change(fn, { scope = 'slide', key = null, slides = null } = {}) {
  S.history.record(S.pres, key);
  const out = fn();
  S.dirty = true;
  emit('change', { scope, slides });
  return out;
}
export const refresh = (scope = 'slide') => emit('change', { scope });

// ───────────── 명령 ─────────────
export const COMMANDS = {};
export function register(map) { Object.assign(COMMANDS, map); }
export function run(name, ...args) {
  const fn = COMMANDS[name];
  if (!fn) { console.warn('알 수 없는 명령', name); return undefined; }
  try {
    const r = fn(...args);
    if (r && typeof r.catch === 'function') r.catch((e) => emit('error', e));
    return r;
  } catch (e) {
    emit('error', e);
    return undefined;
  }
}

/** 슬라이드 이동 */
export function goSlide(i, { keepSel = false } = {}) {
  const n = S.pres.slides.length;
  if (S.editing) emit('endEdit');
  S.cur = Math.max(0, Math.min(n - 1, i));
  if (!keepSel) S.sel.clear();
  S.editing = null;
  S.slideSel = new Set([S.pres.slides[S.cur]?.id]);
  emit('change', { scope: 'nav' });
}
