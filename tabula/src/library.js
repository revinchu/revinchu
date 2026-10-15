// 이 브라우저의 문서 보관함: 최근 문서 최대 30개 + 문서마다 버전 기록 최대 20개 (IndexedDB, gzip 압축)
// 창을 닫아도 남고, 다른 문서를 열어도 이전 문서가 사라지지 않음 (구글 스프레드시트의 '최근 문서 · 버전 기록'과 비슷)
import { idbGet, idbSet, idbDel } from './storage.js';

export const LIB_MAX = 30;
export const VER_MAX = 20;
const INDEX = 'lib:index';
const docKey = (id) => `lib:doc:${id}`;
const verKey = (id, ts) => `lib:ver:${id}:${ts}`;

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

async function dropDoc(e) {
  await idbDel(docKey(e.id));
  for (const v of e.versions ?? []) await idbDel(verKey(e.id, v.ts));
}

/**
 * 문서 저장 (json: 스냅샷 JSON 문자열). version: 버전 기록에도 남김 ({ label })
 * 반환: 갱신된 목록 항목
 */
export function libSave(id, name, json, { version = null, info = {}, max = LIB_MAX } = {}) {
  return serial(async () => {
    const packed = await packText(json);
    await idbSet(docKey(id), packed);
    const list = (await idbGet(INDEX)) ?? [];
    let e = list.find((x) => x.id === id);
    if (!e) { e = { id, created: Date.now(), versions: [] }; list.push(e); }
    Object.assign(e, { name, updated: Date.now(), size: packed.blob.size, ...info });
    if (version) {
      const ts = Date.now();
      await idbSet(verKey(id, ts), packed);
      e.versions = [...(e.versions ?? []), { ts, size: packed.blob.size, label: version.label ?? '' }];
      while (e.versions.length > VER_MAX) {
        // 이름 붙인 버전은 되도록 남김
        const k = e.versions.findIndex((v) => !v.named);
        const [old] = e.versions.splice(k >= 0 ? k : 0, 1);
        await idbDel(verKey(id, old.ts));
      }
    }
    // 오래된 문서부터 정리 (고정한 문서 · 지금 문서 제외)
    const sorted = [...list].sort((a, b) => a.updated - b.updated);
    while (list.length > max) {
      const victim = sorted.find((x) => !x.pinned && x.id !== id && list.includes(x));
      if (!victim) break;
      list.splice(list.indexOf(victim), 1);
      await dropDoc(victim);
    }
    await idbSet(INDEX, list);
    return e;
  });
}

export async function libLoad(id) {
  const text = await unpackText(await idbGet(docKey(id)));
  return text ? JSON.parse(text) : null;
}
export async function libLoadVersion(id, ts) {
  const text = await unpackText(await idbGet(verKey(id, ts)));
  return text ? JSON.parse(text) : null;
}
export function libUpdate(id, patch) {
  return serial(async () => {
    const list = (await idbGet(INDEX)) ?? [];
    const e = list.find((x) => x.id === id);
    if (!e) return null;
    Object.assign(e, patch);
    await idbSet(INDEX, list);
    return e;
  });
}
/** 버전 이름 붙이기 (이름 붙인 버전은 정리할 때 뒤로 미룸) */
export function libNameVersion(id, ts, label) {
  return serial(async () => {
    const list = (await idbGet(INDEX)) ?? [];
    const v = list.find((x) => x.id === id)?.versions?.find((x) => x.ts === ts);
    if (!v) return;
    v.label = label;
    v.named = !!label;
    await idbSet(INDEX, list);
  });
}
export function libRemove(id) {
  return serial(async () => {
    const list = (await idbGet(INDEX)) ?? [];
    const e = list.find((x) => x.id === id);
    if (!e) return;
    list.splice(list.indexOf(e), 1);
    await dropDoc(e);
    await idbSet(INDEX, list);
  });
}
