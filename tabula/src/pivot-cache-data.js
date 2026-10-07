import { Cube, Column } from './cube.js';
import { parseXml, unx, decodeEntities } from './xml.js';
import { scanXmlChildren } from './xml-stream.js';
import { serialOf } from './format.js';
import { jsonReviver } from './block.js';

const TAIL_MIN = 256;
const tailValueValid = value => value === null || typeof value === 'string' || typeof value === 'boolean'
  || typeof value === 'number' && Number.isFinite(value)
  || value && typeof value === 'object' && !Array.isArray(value)
    && (typeof value.error === 'string' || typeof value.code === 'string' && value.code[0] === '#');
function columnValue(col, row) {
  if (col.str && col.str[row] >= 0) return col.dict[col.str[row]];
  const value = col.num?.[row]; return typeof value === 'number' && value === value ? value : null;
}

function growSnapshotColumn(col, need) {
  let cap = col.cap;
  while (cap < need) cap = Math.min(cap * 2, col.limit >= need ? col.limit : Infinity);
  if (cap === col.cap) return;
  if (col.num) { const values = new Float64Array(cap).fill(NaN); values.set(col.num); col.num = values; }
  if (col.str) { const values = new Int32Array(cap).fill(-1); values.set(col.str); col.str = values; }
  col.cap = cap;
}
function flushSnapshotTail(col, end, need) {
  if (col.num || col.str) growSnapshotColumn(col, need);
  if (col.kind === 1) col.num.fill(col.last, col.start, end);
  else if (col.kind === 2) col.str.fill(col.last, col.start, end);
  // Missing tails already have NaN/-1 defaults; never allocate a missing-only column.
  col.pending = false;
}

/** 저장된 캐시 값을 행 객체 없이 열별 숫자·사전 배열에 보관합니다. 빈 문자열은 빈 칸과 구별합니다. */
export class PivotSnapshotBuilder {
  constructor(header, expected = 0) {
    const limit = Math.max(16, Math.min(1048576, expected || 1024));
    this.header = header; this.n = 0; this.cap = Math.min(1024, limit);
    this.columns = header.map(() => ({ cap: this.cap, limit, num: null, str: null, dict: [], index: new Map(), kind: -1, last: undefined, start: 0, run: 0, pending: false }));
  }
  add(row) {
    for (let j = 0; j < this.columns.length; j++) {
      const value = row[j], col = this.columns[j]; let kind = 0, item = null;
      if (typeof value === 'number') {
        if (!Number.isFinite(value)) throw new Error('피벗 캐시에 올바르지 않은 숫자가 있습니다.');
        kind = 1; item = value;
      } else if (value != null) {
        kind = 2;
        const key = typeof value === 'object' ? 'error:' + (value.error ?? value.code) : typeof value + ':' + value;
        item = col.index.get(key);
        if (item === undefined) { item = col.dict.length; col.dict.push(value); col.index.set(key, item); }
      }
      const same = kind === col.kind && Object.is(item, col.last);
      if (same) {
        col.run++;
        if (col.pending) continue;
        // Scalar JSON cannot retain -0. Its full typed storage remains unchanged.
        if (col.run === TAIL_MIN && !Object.is(item, -0) && tailValueValid(kind === 2 ? col.dict[item] : item)) { col.pending = true; continue; }
      } else {
        if (col.pending) flushSnapshotTail(col, this.n, this.n + 1);
        col.kind = kind; col.last = item; col.start = this.n; col.run = 1;
      }
      if (kind === 0) continue;
      growSnapshotColumn(col, this.n + 1);
      if (kind === 1) { col.num ??= new Float64Array(col.cap).fill(NaN); col.num[this.n] = item; }
      else { col.str ??= new Int32Array(col.cap).fill(-1); col.str[this.n] = item; }
    }
    this.n++;
  }
  finish() {
    const columns = this.columns.map(col => {
      const tail = col.pending && (col.num || col.str) ? { start: col.start, value: col.kind === 2 ? col.dict[col.last] : col.last } : null;
      const end = tail ? tail.start : this.n;
      if (col.num || col.str) growSnapshotColumn(col, end);
      // Uniform pending rows remain logical values. Only the retained prefix is copied.
      const value = { num: col.num?.slice(0, end) ?? null, str: col.str?.slice(0, end) ?? null, dict: col.dict, ...(tail ? { tail } : {}) };
      col.num = null; col.str = null; col.index.clear(); col.pending = false; col.last = undefined; return value;
    });
    return { kind: 'pivot-cache', version: 1, header: this.header, n: this.n, columns };
  }
}

export function isPivotSnapshot(value) { return value?.kind === 'pivot-cache' && value.version === 1; }
export function pivotSnapshotValue(snapshot, row, column) {
  const col = snapshot.columns[column]; if (!col) return null;
  if (col.tail && row >= col.tail.start && row < snapshot.n) return col.tail.value;
  return columnValue(col, row);
}
const snapshotCubes = new WeakMap();
export function cubeFromPivotSnapshot(snapshot) {
  let cube = snapshotCubes.get(snapshot); if (cube) return cube;
  cube = new Cube(snapshot.n, snapshot.header, j => new Column(snapshot.n, i => pivotSnapshotValue(snapshot, i, j), { num: snapshot.columns[j].tail ? null : snapshot.columns[j].num }), i => snapshot.header.map((_, j) => pivotSnapshotValue(snapshot, i, j)));
  snapshotCubes.set(snapshot, cube); return cube;
}

/** 기존 행 배열, 살아 있는 형식화 배열, JSON·Blob 보관본을 모두 읽습니다. */
export function restorePivotSnapshot(data) {
  if (!isPivotSnapshot(data)) return data;
  if (!Number.isSafeInteger(data.n) || data.n < 0 || !Array.isArray(data.header) || !Array.isArray(data.columns) || data.columns.length !== data.header.length) throw new Error('피벗 캐시 보관본이 올바르지 않습니다.');
  const array = (raw, Type, empty, length, strict) => {
    if (raw == null) return null;
    const value = raw.__ta ? jsonReviver('', raw) : raw;
    if (value instanceof Type) { if (value.length !== length) throw new Error('피벗 캐시 열의 길이가 올바르지 않습니다.'); return value; }
    if (strict && (ArrayBuffer.isView(value) || Array.isArray(value) && value.length !== length)) throw new Error('피벗 캐시 열의 길이가 올바르지 않습니다.');
    const out = new Type(length).fill(empty);
    for (const key of Object.keys(value)) {
      const i = Number(key), v = value[key], index = Number.isSafeInteger(i) && i >= 0 && String(i) === key;
      if (strict && (!index || i >= length)) throw new Error('피벗 캐시 열의 길이가 올바르지 않습니다.');
      if (index && i < length && typeof v === 'number' && Number.isFinite(v)) out[i] = v;
    }
    return out;
  };
  let changed = false;
  const columns = data.columns.map(col => {
    const tail = col.tail;
    if (tail !== undefined && (!tail || typeof tail !== 'object' || Array.isArray(tail) || !Number.isSafeInteger(tail.start) || tail.start < 0 || tail.start >= data.n || !Object.hasOwn(tail, 'value') || !tailValueValid(tail.value))) throw new Error('피벗 캐시 꼬리 값이 올바르지 않습니다.');
    const length = tail ? tail.start : data.n, num = array(col.num, Float64Array, NaN, length, !!tail), str = array(col.str, Int32Array, -1, length, !!tail);
    if (num === col.num && str === col.str) return col;
    changed = true; return { ...col, num, str };
  });
  return changed ? { ...data, columns } : data;
}

export function pivotCacheDate(value, date1904 = false) {
  const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(value ?? '');
  if (!m) return value ?? '';
  return serialOf(+m[1], +m[2], +m[3], date1904) + (Number(m[4] ?? 0) * 3600000 + Number(m[5] ?? 0) * 60000 + Math.round(Number(m[6] ?? 0) * 1000)) / 86400000;
}

const recordSpace = code => code === 32 || code === 9 || code === 10 || code === 13;
const recordName = code => code >= 65 && code <= 90 || code >= 97 && code <= 122 || code >= 48 && code <= 57 || code === 95 || code === 45 || code === 46 || code === 58;

/** 표준 자기닫힘 cache cell만 직접 읽습니다. 전체 레코드가 이 형태가 아니면
 * null을 반환하여 기존 XML 파서로 처리합니다. fields는 database 열 순서입니다. */
export function scanPivotCacheRecord(xml, fields, date1904 = false) {
  let pos = 0;
  while (recordSpace(xml.charCodeAt(pos))) pos++;
  if (xml[pos++] !== '<') return null;
  const rootStart = pos;
  while (recordName(xml.charCodeAt(pos))) pos++;
  const root = xml.slice(rootStart, pos);
  if (!/^(?:[A-Za-z_][\w.-]*:)?r$/.test(root)) return null;
  // Namespace declarations on the record are harmless; unfamiliar attrs fall back.
  while (recordSpace(xml.charCodeAt(pos))) {
    while (recordSpace(xml.charCodeAt(pos))) pos++;
    if (xml[pos] === '>' || xml[pos] === '/') break;
    const start = pos;
    while (recordName(xml.charCodeAt(pos))) pos++;
    const name = xml.slice(start, pos);
    if (name !== 'xmlns' && !/^xmlns:[A-Za-z_][\w.-]*$/.test(name)) return null;
    while (recordSpace(xml.charCodeAt(pos))) pos++;
    if (xml[pos++] !== '=') return null;
    while (recordSpace(xml.charCodeAt(pos))) pos++;
    const quote = xml[pos++]; if (quote !== '"' && quote !== "'") return null;
    const end = xml.indexOf(quote, pos); if (end < 0) return null;
    const text = xml.slice(pos, end); if (text.includes('<')) return null;
    decodeEntities(text); pos = end + 1;
  }
  const row = new Array(fields.length).fill(null);
  if (xml[pos] === '/') {
    if (xml[++pos] !== '>') return null;
    pos++; while (recordSpace(xml.charCodeAt(pos))) pos++;
    return pos === xml.length ? row : null;
  }
  if (xml[pos++] !== '>') return null;
  let column = 0;
  for (;;) {
    while (recordSpace(xml.charCodeAt(pos))) pos++;
    if (xml[pos++] !== '<') return null;
    if (xml[pos] === '/') {
      pos++;
      if (!xml.startsWith(root, pos)) return null;
      pos += root.length;
      while (recordSpace(xml.charCodeAt(pos))) pos++;
      if (xml[pos++] !== '>') return null;
      while (recordSpace(xml.charCodeAt(pos))) pos++;
      return pos === xml.length ? row : null;
    }
    const start = pos; let colon = -1;
    while (recordName(xml.charCodeAt(pos))) {
      if (xml[pos] === ':') { if (colon >= 0) return null; colon = pos; }
      pos++;
    }
    const local = colon < 0 ? start : colon + 1, kind = xml.charCodeAt(local);
    if (pos !== local + 1 || colon === start || colon >= 0 && !(/[A-Za-z_]/.test(xml[start]))) return null;
    if (kind !== 120 && kind !== 110 && kind !== 115 && kind !== 100 && kind !== 98 && kind !== 101 && kind !== 109) return null;
    let value, seen = false;
    while (recordSpace(xml.charCodeAt(pos))) {
      while (recordSpace(xml.charCodeAt(pos))) pos++;
      if (xml[pos] === '/') break;
      if (xml[pos++] !== 'v' || seen) return null;
      while (recordSpace(xml.charCodeAt(pos))) pos++;
      if (xml[pos++] !== '=') return null;
      while (recordSpace(xml.charCodeAt(pos))) pos++;
      const quote = xml[pos++]; if (quote !== '"' && quote !== "'") return null;
      const end = xml.indexOf(quote, pos); if (end < 0) return null;
      const text = xml.slice(pos, end); if (text.includes('<')) return null;
      value = decodeEntities(text); pos = end + 1; seen = true;
    }
    if (xml[pos++] !== '/' || xml[pos++] !== '>') return null;
    if (column < fields.length) {
      row[column] = kind === 120 ? fields[column].shared[Number(value)] ?? null
        : kind === 110 ? Number(value) : kind === 100 ? pivotCacheDate(value, date1904)
          : kind === 98 ? value === '1' || value === 'true' : kind === 109 ? null
            : kind === 101 ? { error: value ?? '#N/A' } : unx(value ?? '');
    }
    column++;
  }
}

/** 큰 XML도 완전한 레코드 단위로 읽으며 저장된 모든 행을 보존합니다. */
export function* readPivotSnapshotXml(bytes, fields, date1904 = false, expected = 0) {
  const db = fields.filter(field => field.db), builder = new PivotSnapshotBuilder(db.map(field => field.name), expected);
  for (const xml of scanXmlChildren(bytes, 'r')) {
    let row = scanPivotCacheRecord(xml, db, date1904);
    if (!row) {
      const node = parseXml(xml); row = new Array(db.length).fill(null);
      for (let j = 0; j < Math.min(node.children.length, db.length); j++) {
        const cell = node.children[j], value = cell.attrs.v;
        row[j] = cell.name === 'x' ? db[j].shared[Number(value)] ?? null : cell.name === 'n' ? Number(value) : cell.name === 'd' ? pivotCacheDate(value, date1904) : cell.name === 'b' ? value === '1' || value === 'true' : cell.name === 'm' ? null : cell.name === 'e' ? { error: value ?? '#N/A' } : unx(value ?? '');
      }
    }
    builder.add(row);
    if (builder.n % 2048 === 0) yield builder.n;
  }
  if (expected && builder.n !== expected) throw new Error('피벗 캐시의 저장된 행 수가 올바르지 않습니다.');
  return builder.finish();
}
