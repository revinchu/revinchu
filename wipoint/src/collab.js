// 공동 편집 (브라우저): 같은 '방' 사람끼리 슬라이드 변경분 · 참가자(이름 · 보고 있는 슬라이드 · 선택) 주고받기
//   전송: server.js 의 /api/collab 중계(SSE + POST)가 있으면 그것 (다른 기기와), 없으면 BroadcastChannel (같은 브라우저의 다른 탭)
import { S, curSlide, on, emit } from './state.js';
import { uid } from './model.js';
import { baseline, diffPres, applyPatch, fullState, applyFull, peerColor } from './collabcore.js';

export const collab = { room: null, me: uid('u'), mode: null, peers: new Map(), base: null, joining: false, queue: [] };
let transport = null;
let flushTimer = 0;
let presTimer = 0;
let beat = 0;

export const myName = () => S.userName ?? (() => { try { return localStorage.getItem('wipoint:user'); } catch { return null; } })() ?? '사용자';
export const collabLink = (room = collab.room) => `${location.origin}${location.pathname}?collab=${encodeURIComponent(room)}`;

function send(msg) { transport?.send(JSON.stringify({ ...msg, from: collab.me })); }

function sseTransport(room) {
  const es = new EventSource(`/api/collab/${room}/events?c=${collab.me}`);
  es.onmessage = (e) => receive(e.data);
  es.onerror = () => emit('collabStatus', '다시 연결하는 중…');
  return {
    send: (data) => fetch(`/api/collab/${room}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: data }).catch(() => emit('collabStatus', '보내지 못했습니다 (서버 연결 확인)')),
    close: () => es.close(),
  };
}
function bcTransport(room) {
  const bc = new BroadcastChannel(`wipoint-collab-${room}`);
  bc.onmessage = (e) => receive(e.data);
  return { send: (data) => bc.postMessage(data), close: () => bc.close() };
}

/** 방 열기 (join = 다른 사람이 연 방에 들어감: 그 문서를 받아 옴) */
export async function startCollab(room, { join = false } = {}) {
  stopCollab(false);
  collab.room = room;
  collab.joining = join;
  collab.base = baseline(S.pres);
  const server = await fetch('/api/collab/ping').then((r) => r.ok && r.headers.get('content-type')?.includes('json')).catch(() => false);
  collab.mode = server ? 'server' : 'local';
  transport = server ? sseTransport(room) : bcTransport(room);
  if (join) send({ type: 'hello' });
  else send({ type: 'state', ...fullState(S.pres) });
  presence();
  beat = setInterval(() => { presence(); prunePeers(); }, 5000);
  emit('collabPeers');
}
export function stopCollab(notify = true) {
  if (!collab.room) return;
  flush();
  send({ type: 'bye' });
  transport?.close();
  transport = null;
  clearInterval(beat);
  collab.room = null;
  collab.peers.clear();
  if (notify) emit('collabPeers');
}

/** 내 변경분 보내기 */
function flush() {
  clearTimeout(flushTimer);
  if (!collab.room || collab.joining || !collab.base) return;
  const p = diffPres(collab.base, S.pres);
  if (p) send({ type: 'patch', ...p });
}
function presence() {
  if (!collab.room) return;
  send({ type: 'presence', name: myName(), slide: curSlide()?.id ?? null, sel: [...S.sel], editing: S.editing?.id ?? null });
}
function prunePeers() {
  const now = Date.now();
  let gone = false;
  for (const [id, p] of collab.peers) if (now - p.at > 30000) { collab.peers.delete(id); gone = true; }
  if (gone) emit('collabPeers');
}

function receive(data) {
  let m;
  try { m = JSON.parse(data); } catch { return; }
  if (!m || m.from === collab.me || !collab.room) return;
  switch (m.type) {
    case 'hello':
      if (!collab.joining) { flush(); send({ type: 'state', to: m.from, ...fullState(S.pres) }); }
      presence();
      break;
    case 'state':
      if (m.to && m.to !== collab.me) break;
      if (!collab.joining) break;
      collab.joining = false;
      if (S.editing) emit('endEdit');
      collab.base = applyFull(S.pres, m);
      S.history.clear();
      S.cur = Math.min(S.cur, S.pres.slides.length - 1);
      S.sel.clear();
      emit('change', { scope: 'all' });
      emit('collabStatus', '문서를 받았습니다');
      drainQueue();
      break;
    case 'patch': {
      if (collab.joining) { collab.queue.push(m); break; }
      // 내가 글을 고치는 슬라이드가 바뀌었으면 편집이 끝날 때 적용
      const editingSlide = S.editing ? curSlide()?.id : null;
      if (editingSlide && m.slides?.[editingSlide]) { collab.queue.push(m); break; }
      applyRemote(m);
      break;
    }
    case 'presence':
      collab.peers.set(m.from, { name: m.name ?? '사용자', color: peerColor(m.from), slide: m.slide, sel: m.sel ?? [], editing: m.editing, at: Date.now() });
      emit('collabPeers');
      break;
    case 'bye':
      collab.peers.delete(m.from);
      emit('collabPeers');
      break;
    default:
  }
}
function applyRemote(m) {
  flush();
  const curId = curSlide()?.id;
  applyPatch(S.pres, m, collab.base);
  // 보고 있던 슬라이드가 지워졌으면 가까운 슬라이드로
  const i = S.pres.slides.findIndex((s) => s.id === curId);
  if (i >= 0) S.cur = i; else S.cur = Math.min(S.cur, S.pres.slides.length - 1);
  const objIds = new Set(curSlide()?.objects.map((o) => o.id) ?? []);
  for (const id of [...S.sel]) if (!objIds.has(id)) S.sel.delete(id);
  S.dirty = true;
  emit('change', { scope: 'all', remote: true });
}

// 내 변경 → 잠시 모아서 보내기, 선택 · 슬라이드 이동 → 참가자 표시
on('change', ({ scope = 'slide', remote } = {}) => {
  if (!collab.room || remote) return;
  if (scope === 'nav' || scope === 'view') { clearTimeout(presTimer); presTimer = setTimeout(presence, 200); return; }
  if (scope === 'none') return;
  if (!S.editing) setTimeout(drainQueue, 0);
  clearTimeout(flushTimer);
  flushTimer = setTimeout(flush, scope === 'text' ? 400 : 150);
});
on('selection', () => { if (collab.room) { clearTimeout(presTimer); presTimer = setTimeout(presence, 200); drainQueue(); } });
function drainQueue() { if (S.editing || !collab.queue.length || collab.joining) return; const q = collab.queue.splice(0); for (const m of q) applyRemote(m); }
on('endEdit', () => setTimeout(drainQueue, 0));
addEventListener('beforeunload', () => { if (collab.room) send({ type: 'bye' }); });

/** 이 슬라이드를 보고 있는 다른 사람들 */
export const peersOn = (slideId) => [...collab.peers.values()].filter((p) => p.slide === slideId);
