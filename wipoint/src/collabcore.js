// 공동 편집 핵심 (DOM 없음): 문서를 슬라이드 단위로 비교해 바뀐 것만 보내고(patch), 받은 patch 를 적용
//   충돌은 슬라이드 단위 '나중에 쓴 사람 우선'. 문서 수준 값(테마 · 크기 · 바닥글 · 마스터 …)은 meta 하나로.

/** 기준선: { slides: Map(id → json), order, meta, media: Set } */
export function baseline(pres) {
  return {
    slides: new Map(pres.slides.map((s) => [s.id, JSON.stringify(s)])),
    order: pres.slides.map((s) => s.id).join(','),
    meta: metaJson(pres),
    media: new Set(Object.keys(pres.media ?? {})),
  };
}
function metaJson(pres) {
  const { slides: _s, media: _m, fonts: _f, ...rest } = pres;
  return JSON.stringify(rest);
}

/** 기준선 이후 바뀐 것 → patch (없으면 null). 기준선은 새 상태로 바뀜 */
export function diffPres(base, pres) {
  const patch = {};
  const seen = new Set();
  for (const s of pres.slides) {
    seen.add(s.id);
    const j = JSON.stringify(s);
    if (base.slides.get(s.id) !== j) { (patch.slides ??= {})[s.id] = j; base.slides.set(s.id, j); }
  }
  for (const id of [...base.slides.keys()]) if (!seen.has(id)) { (patch.removed ??= []).push(id); base.slides.delete(id); }
  const order = pres.slides.map((s) => s.id).join(',');
  if (order !== base.order) { patch.order = pres.slides.map((s) => s.id); base.order = order; }
  const meta = metaJson(pres);
  if (meta !== base.meta) { patch.meta = meta; base.meta = meta; }
  for (const [id, url] of Object.entries(pres.media ?? {})) if (!base.media.has(id)) { (patch.media ??= {})[id] = url; base.media.add(id); }
  return Object.keys(patch).length ? patch : null;
}

/** patch 적용 (기준선도 함께 맞춰서 되돌려 보내지 않음). 바뀐 슬라이드 id 목록을 돌려줌 */
export function applyPatch(pres, patch, base = null) {
  const changed = [];
  if (patch.media) { pres.media ??= {}; for (const [id, url] of Object.entries(patch.media)) { pres.media[id] = url; base?.media.add(id); } }
  if (patch.meta) {
    const meta = JSON.parse(patch.meta);
    for (const k of Object.keys(pres)) if (!['slides', 'media', 'fonts'].includes(k) && !(k in meta)) delete pres[k];
    Object.assign(pres, meta);
    if (base) base.meta = patch.meta;
  }
  const byId = new Map(pres.slides.map((s) => [s.id, s]));
  for (const [id, j] of Object.entries(patch.slides ?? {})) {
    const s = JSON.parse(j);
    const cur = byId.get(id);
    if (cur) { for (const k of Object.keys(cur)) delete cur[k]; Object.assign(cur, s); } else { byId.set(id, s); pres.slides.push(s); }
    base?.slides.set(id, j);
    changed.push(id);
  }
  if (patch.removed) {
    const gone = new Set(patch.removed);
    pres.slides = pres.slides.filter((s) => !gone.has(s.id));
    for (const id of gone) base?.slides.delete(id);
  }
  if (patch.order) {
    const pos = new Map(patch.order.map((id, i) => [id, i]));
    // 순서에 없는 (내가 방금 만든) 슬라이드는 원래 앞 슬라이드 뒤에 남김
    pres.slides = pres.slides.map((s, i) => [s, pos.has(s.id) ? pos.get(s.id) : (pos.get(pres.slides[i - 1]?.id) ?? i) + 0.5]).sort((a, b) => a[1] - b[1]).map(([s]) => s);
  }
  if (base) base.order = pres.slides.map((s) => s.id).join(',');
  return changed;
}

/** 전체 상태 메시지 (처음 들어온 사람에게) */
export const fullState = (pres) => ({ slides: Object.fromEntries(pres.slides.map((s) => [s.id, JSON.stringify(s)])), order: pres.slides.map((s) => s.id), meta: metaJson(pres), media: { ...(pres.media ?? {}) } });

const COLORS = ['#c43e1c', '#2b7cd3', '#107c41', '#8764b8', '#d83b01', '#00838f', '#b4009e', '#5c2e91'];
export const peerColor = (id) => COLORS[[...String(id)].reduce((a, c) => a + c.charCodeAt(0), 0) % COLORS.length];
/** 전체 상태로 바꿈 (들어온 사람) */
export function applyFull(pres, st) {
  const meta = JSON.parse(st.meta);
  for (const k of Object.keys(pres)) if (k !== 'fonts') delete pres[k];
  Object.assign(pres, meta, { slides: st.order.map((id) => JSON.parse(st.slides[id])), media: { ...st.media } });
  return baseline(pres);
}
