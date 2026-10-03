import { Cube, Column } from './cube.js';
import { parseXml, unx } from './xml.js';
import { scanXmlChildren } from './xml-stream.js';
import { serialOf } from './format.js';
import { jsonReviver } from './block.js';

/** 저장된 캐시 값을 행 객체 없이 열별 숫자·사전 배열에 보관합니다. 빈 문자열은 빈 칸과 구별합니다. */
export class PivotSnapshotBuilder {
  constructor(header, expected = 0) {
    this.header = header; this.n = 0; this.cap = Math.max(16, Math.min(1048576, expected || 1024));
    this.columns = header.map(() => ({ num: null, str: null, dict: [], index: new Map() }));
  }
  add(row) {
    if (this.n >= this.cap) {
      const cap = this.cap * 2;
      for (const col of this.columns) {
        if (col.num) { const values = new Float64Array(cap).fill(NaN); values.set(col.num); col.num = values; }
        if (col.str) { const values = new Int32Array(cap).fill(-1); values.set(col.str); col.str = values; }
      }
      this.cap = cap;
    }
    for (let j = 0; j < this.columns.length; j++) {
      const value = row[j], col = this.columns[j];
      if (value == null) continue;
      if (typeof value === 'number') {
        if (!Number.isFinite(value)) throw new Error('피벗 캐시에 올바르지 않은 숫자가 있습니다.');
        col.num ??= new Float64Array(this.cap).fill(NaN); col.num[this.n] = value;
      } else {
        col.str ??= new Int32Array(this.cap).fill(-1);
        const key = typeof value === 'object' ? 'error:' + (value.error ?? value.code) : typeof value + ':' + value;
        let code = col.index.get(key);
        if (code === undefined) { code = col.dict.length; col.dict.push(value); col.index.set(key, code); }
        col.str[this.n] = code;
      }
    }
    this.n++;
  }
  finish() {
    const columns = this.columns.map(col => {
      const value = { num: col.num?.slice(0, this.n) ?? null, str: col.str?.slice(0, this.n) ?? null, dict: col.dict };
      col.num = null; col.str = null; col.index.clear(); return value;
    });
    return { kind: 'pivot-cache', version: 1, header: this.header, n: this.n, columns };
  }
}

export function isPivotSnapshot(value) { return value?.kind === 'pivot-cache' && value.version === 1; }
export function pivotSnapshotValue(snapshot, row, column) {
  const col = snapshot.columns[column]; if (!col) return null;
  if (col.str && col.str[row] >= 0) return col.dict[col.str[row]];
  const value = col.num?.[row]; return typeof value === 'number' && value === value ? value : null;
}
const snapshotCubes = new WeakMap();
export function cubeFromPivotSnapshot(snapshot) {
  let cube = snapshotCubes.get(snapshot); if (cube) return cube;
  cube = new Cube(snapshot.n, snapshot.header, j => new Column(snapshot.n, i => pivotSnapshotValue(snapshot, i, j), { num: snapshot.columns[j].num }), i => snapshot.header.map((_, j) => pivotSnapshotValue(snapshot, i, j)));
  snapshotCubes.set(snapshot, cube); return cube;
}

/** 기존 행 배열, 살아 있는 형식화 배열, JSON·Blob 보관본을 모두 읽습니다. */
export function restorePivotSnapshot(data) {
  if (!isPivotSnapshot(data)) return data;
  if (!Number.isSafeInteger(data.n) || data.n < 0 || !Array.isArray(data.header) || !Array.isArray(data.columns) || data.columns.length !== data.header.length) throw new Error('피벗 캐시 보관본이 올바르지 않습니다.');
  const array = (raw, Type, empty) => {
    if (raw == null) return null;
    const value = raw.__ta ? jsonReviver('', raw) : raw;
    if (value instanceof Type) { if (value.length !== data.n) throw new Error('피벗 캐시 열의 길이가 올바르지 않습니다.'); return value; }
    const out = new Type(data.n).fill(empty);
    for (const key of Object.keys(value)) { const i = Number(key), v = value[key]; if (Number.isSafeInteger(i) && i >= 0 && i < data.n && String(i) === key && typeof v === 'number' && Number.isFinite(v)) out[i] = v; }
    return out;
  };
  let changed = false;
  const columns = data.columns.map(col => { const num = array(col.num, Float64Array, NaN), str = array(col.str, Int32Array, -1); if (num === col.num && str === col.str) return col; changed = true; return { num, str, dict: col.dict }; });
  return changed ? { ...data, columns } : data;
}

export function pivotCacheDate(value, date1904 = false) {
  const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(value ?? '');
  if (!m) return value ?? '';
  return serialOf(+m[1], +m[2], +m[3], date1904) + (Number(m[4] ?? 0) * 3600000 + Number(m[5] ?? 0) * 60000 + Math.round(Number(m[6] ?? 0) * 1000)) / 86400000;
}

/** 큰 XML도 완전한 레코드 단위로 읽으며 저장된 모든 행을 보존합니다. */
export function* readPivotSnapshotXml(bytes, fields, date1904 = false, expected = 0) {
  const db = fields.filter(field => field.db), builder = new PivotSnapshotBuilder(db.map(field => field.name), expected);
  for (const xml of scanXmlChildren(bytes, 'r')) {
    const node = parseXml(xml), row = new Array(db.length).fill(null);
    for (let j = 0; j < Math.min(node.children.length, db.length); j++) {
      const cell = node.children[j], value = cell.attrs.v;
      row[j] = cell.name === 'x' ? db[j].shared[Number(value)] ?? null : cell.name === 'n' ? Number(value) : cell.name === 'd' ? pivotCacheDate(value, date1904) : cell.name === 'b' ? value === '1' || value === 'true' : cell.name === 'm' ? null : cell.name === 'e' ? { error: value ?? '#N/A' } : unx(value ?? '');
    }
    builder.add(row);
    if (builder.n % 2048 === 0) yield builder.n;
  }
  if (expected && builder.n !== expected) throw new Error('피벗 캐시의 저장된 행 수가 올바르지 않습니다.');
  return builder.finish();
}
