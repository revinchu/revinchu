// 이 브라우저의 문서 보관함: 최근 문서 최대 30개 + 문서마다 버전 기록 최대 20개 (IndexedDB, gzip 압축)
// 창을 닫아도 남고, 다른 문서를 열어도 이전 문서가 사라지지 않음 (구글 스프레드시트의 '최근 문서 · 버전 기록'과 비슷)
import { idbGet, idbUpdate } from './storage.js';

export const LIB_MAX = 30;
export const VER_MAX = 20;
const INDEX = 'lib:index';
const docKey = (id) => `lib:doc:${id}`;
const verKey = (id, ts) => `lib:ver:${id}:${ts}`;
// 읽거나 저장한 본문의 revision. 다른 탭에서 바뀐 본문을 조용히 덮어쓰지 않습니다.
const observed = new Map();
const revisionOf = (packed) => packed ? packed.revision ?? 'legacy' : null;
const conflict = () => Object.assign(new Error('다른 탭에서 이 보관 문서가 바뀌었습니다. 현재 내용을 파일로 보관한 뒤 최신 문서를 다시 여세요.'), { code: 'LIB_CONFLICT' });
function checkRevision(id, packed) {
  if (observed.has(id) && observed.get(id) !== revisionOf(packed)) throw conflict();
}

export const newDocId = () => `d${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;

/** 문자열 → gzip Blob (압축을 못 하는 브라우저는 그대로) */
export async function packText(text) {
  if (globalThis.CompressionStream) {
    const stream = new Blob([text]).stream().pipeThrough(new CompressionStream('gzip'));
    return { gz: true, blob: await new Response(stream).blob() };
  }
  return { gz: false, blob: new Blob([text]) };
}
export async function unpackText(rec) {
  if (!rec) return null;
  if (rec.gz && globalThis.DecompressionStream) return new Response(rec.blob.stream().pipeThrough(new DecompressionStream('gzip'))).text();
  return rec.blob.text();
}

let queue = Promise.resolve();
/** 목록 갱신은 한 번에 하나씩 (자동 저장이 겹쳐도 목록이 깨지지 않게) */
function serial(fn) {
  const run = queue.then(fn, fn);
  queue = run.catch(() => {});
  return run;
}

export async function libList() {
  const list = (await idbGet(INDEX)) ?? [];
  return [...list].sort((a, b) => (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0) || b.updated - a.updated);
}

const dropKeys = (e) => [docKey(e.id), ...(e.versions ?? []).map((v) => verKey(e.id, v.ts))];

/**
 * 문서 저장 (json: 스냅샷 JSON 문자열). version: 버전 기록에도 남김 ({ label })
 * 반환: 갱신된 목록 항목
 */
export function libSave(id, name, json, { version = null, info = {}, max = LIB_MAX, keepIds = [] } = {}) {
  return serial(async () => {
    if (!observed.has(id)) {
      // 자동 복원/옛 저장 형식에는 관측 revision이 없을 수 있습니다. 현재 본문과
      // 같은 사본임을 확인한 경우에만 채택하며, 읽은 뒤의 변경도 아래 tx에서 재검사합니다.
      const previous = await idbGet(docKey(id));
      if (previous && await unpackText(previous) !== json) throw conflict();
      observed.set(id, revisionOf(previous));
    }
    const packed = await packText(json);
    // gzip은 트랜잭션 밖에서 준비하고, 문서·이력·정리·목록은 한 번에 커밋합니다.
    const result = await idbUpdate([INDEX, docKey(id)], (stored) => {
      checkRevision(id, stored.get(docKey(id)));
      const list = stored.get(INDEX) ?? [], set = [], remove = [];
      let e = list.find((x) => x.id === id);
      if (!e) { e = { id, created: Date.now(), versions: [] }; list.push(e); }
      const revision = newDocId();
      const record = { ...packed, revision };
      set.push([docKey(id), record]);
      Object.assign(e, { ...info, id, name, created: e.created, versions: e.versions ?? [], updated: Math.max(Date.now(), (e.updated ?? 0) + 1), size: packed.blob.size, revision });
      if (version) {
        // 같은 밀리초에 저장해도 과거 버전의 key가 겹치지 않게 합니다.
        let ts = Date.now();
        for (const v of e.versions) ts = Math.max(ts, v.ts + 1);
        set.push([verKey(id, ts), record]);
        e.versions = [...e.versions, { ts, size: packed.blob.size, label: version.label ?? '', named: !!version.named }];
        while (e.versions.length > VER_MAX) {
          const k = e.versions.findIndex((v) => !v.named);
          const [old] = e.versions.splice(k >= 0 ? k : 0, 1);
          remove.push(verKey(id, old.ts));
        }
      }
      const sorted = [...list].sort((a, b) => a.updated - b.updated);
      while (list.length > max) {
        const victim = sorted.find((x) => !x.pinned && x.id !== id && !keepIds.includes(x.id) && list.includes(x));
        if (!victim) break;
        list.splice(list.indexOf(victim), 1);
        for (const key of dropKeys(victim)) remove.push(key);
      }
      set.push([INDEX, list]);
      return { set, delete: remove, result: e };
    });
    observed.set(id, result.revision);
    return result;
  });
}

export function libLoad(id) {
  return serial(async () => {
    const rec = await idbGet(docKey(id));
    const text = await unpackText(rec);
    const data = text ? JSON.parse(text) : null;
    observed.set(id, revisionOf(rec));
    return data;
  });
}
export function libLoadVersion(id, ts) {
  return serial(async () => {
    // 과거본과 그 시점의 현재 revision을 함께 읽어 복원 직후 다른 탭 저장도 감지합니다.
    const { rec, revision } = await idbUpdate([docKey(id), verKey(id, ts)], (stored) => ({ result: { rec: stored.get(verKey(id, ts)), revision: revisionOf(stored.get(docKey(id))) } }));
    const text = await unpackText(rec);
    const data = text ? JSON.parse(text) : null;
    if (!observed.has(id)) observed.set(id, revision);
    return data;
  });
}
export function libUpdate(id, patch) {
  return serial(() => idbUpdate([INDEX], (stored) => {
    const list = stored.get(INDEX) ?? [];
    const e = list.find((x) => x.id === id);
    if (!e) return { result: null };
    // 본문·버전의 식별자는 목록 편집으로 바뀌면 안 됩니다.
    for (const key of ['name', 'pinned']) if (Object.hasOwn(patch, key)) e[key] = patch[key];
    return { set: [[INDEX, list]], result: e };
  }));
}
/** 버전 이름 붙이기 (이름 붙인 버전은 정리할 때 뒤로 미룸) */
export function libNameVersion(id, ts, label) {
  return serial(() => idbUpdate([INDEX], (stored) => {
    const list = stored.get(INDEX) ?? [];
    const v = list.find((x) => x.id === id)?.versions?.find((x) => x.ts === ts);
    if (!v) return {};
    v.label = label;
    v.named = !!label;
    return { set: [[INDEX, list]] };
  }));
}
export function libRemove(id) {
  return serial(async () => {
    await idbUpdate([INDEX, docKey(id)], (stored) => {
      checkRevision(id, stored.get(docKey(id)));
      const list = stored.get(INDEX) ?? [];
      const e = list.find((x) => x.id === id);
      if (!e) return {};
      list.splice(list.indexOf(e), 1);
      return { set: [[INDEX, list]], delete: dropKeys(e) };
    });
    observed.set(id, null);
  });
}
