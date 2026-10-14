// 앱 공용 상태 + 명령 등록소 + 변경 알림 (모듈끼리 서로 import 하지 않도록 이곳을 거침)
import { History, newDoc, draft, firstPos } from './model.js';

export const S = {
  doc: newDoc(),
  sel: null,              // { a: 위치(고정 끝), f: 위치(움직이는 끝) } — 위치 = { p: [..], o }
  history: new History(),
  zoom: 1,
  zoomMode: null,         // null | 'pageWidth' | 'onePage' | 'multi'
  view: 'print',          // print | web | read | outline | draft
  marks: false,           // 편집 기호 표시 (Ctrl+*)
  markup: 'all',          // 변경 내용 표시: all | simple | none | original
  showRuler: true,
  showNav: false,
  pane: null,             // null | 'styles' | 'comments' | 'markdown' | 'revisions' | 'picture' | 'mcp'
  docName: '문서1',
  dirty: false,
  fileHandle: null,
  painter: null,          // 서식 복사 { run, para, keep }
  typing: null,           // 다음에 칠 글자 서식 (빈 선택에서 Ctrl+B 등)
  compose: null,          // IME 조합 중 { start, len }
  layout: null,           // 마지막 쪽 나누기 결과
  objSel: null,           // 선택한 그림 { p, o }
  hf: null,               // 머리글/바닥글 편집 중 { kind: 'header'|'footer', which: 'default'|'first'|'even', si }
  opts: {},
};

export const curSel = () => S.sel ?? { a: firstPos(S.doc.body), f: firstPos(S.doc.body) };

// ───────────── 알림 ─────────────
const listeners = new Map();
export function on(ev, fn) { if (!listeners.has(ev)) listeners.set(ev, []); listeners.get(ev).push(fn); }
export function emit(ev, data) { for (const fn of listeners.get(ev) ?? []) { try { fn(data); } catch (e) { console.error(e); } } }

/**
 * 문서 바꾸기: fn(d) — d = draft(S.doc) (d.doc 이 새 문서). fn 이 돌려준 선택(또는 위치)을 새 선택으로.
 * key: 같은 key 의 연속 변경은 실행 취소 한 번으로 묶음 (입력)
 */
export function change(fn, { key = null, sel = undefined, noHistory = false, scope = 'doc' } = {}) {
  const before = S.doc;
  const selBefore = S.sel;
  const d = draft(before);
  const out = fn(d);
  if (out === false) return false;
  if (!noHistory) S.history.record(before, selBefore, key);
  d.doc.props = { ...d.doc.props, modified: new Date().toISOString() };
  S.doc = d.doc;
  if (out && out.p) S.sel = { a: out, f: out };
  else if (out && out.a) S.sel = out;
  else if (sel !== undefined) S.sel = sel;
  S.dirty = true;
  emit('change', { scope, key });
  return out;
}
/** 문서 통째로 바꾸기 (열기 · 새 문서) */
export function setDoc(doc, name = null) {
  S.doc = doc;
  S.history = new History();
  S.sel = { a: firstPos(doc.body), f: firstPos(doc.body) };
  if (name) S.docName = name;
  S.dirty = false;
  S.objSel = null;
  S.hf = null;
  emit('change', { scope: 'all', newDoc: true });
}
export const refresh = (scope = 'view') => emit('change', { scope });

// ───────────── 명령 ─────────────
export const COMMANDS = {};
export const COMMAND_INFO = {}; // 이름 → { label, args } (MCP · 명령 검색)
export function register(map, info = {}) { Object.assign(COMMANDS, map); Object.assign(COMMAND_INFO, info); }
const REPEATABLE = new Set(['bold', 'italic', 'underline', 'strike', 'fontColor', 'highlight', 'fontFamily', 'fontSize', 'growFont', 'shrinkFont', 'alignLeft', 'alignCenter', 'alignRight', 'alignJustify', 'applyStyle', 'bullets', 'numbering', 'indentMore', 'indentLess', 'clearFormat', 'superscript', 'subscript', 'lineSpacing', 'insertTable', 'insertRowBelow', 'insertRowAbove', 'deleteRow', 'insertColRight', 'insertColLeft', 'deleteCol', 'pageBreak']);
export function run(name, ...args) {
  const fn = COMMANDS[name];
  if (!fn) { console.warn('알 수 없는 명령', name); emit('error', new Error(`알 수 없는 명령: ${name}`)); return undefined; }
  if (REPEATABLE.has(name) && args.every((a) => a == null || typeof a !== 'object')) S.lastRepeat = { name, args };
  try {
    const r = fn(...args);
    if (r && typeof r.catch === 'function') r.catch((e) => emit('error', e));
    return r;
  } catch (e) {
    emit('error', e);
    return undefined;
  }
}
