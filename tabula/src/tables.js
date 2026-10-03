// 엑셀 '표'(ListObject): 스타일, 열 이름, 구조적 참조, 자동 확장 (DOM 없음)
// 표 모델: { id, name, r1, c1, r2, c2, header, totals, style, banded, bandedCols, firstCol, lastCol,
//           filter: { criteria, hidden, sort } | null, totalsFns: { [열]: 'sum' | ... }, columns: [이름] }
// r1 은 머리글 행(header 가 true 일 때), r2 는 요약 행(totals 가 true 일 때)까지 포함.

import { PRESET_STYLES, presetSwatch, tablePresetCell, MODERN_STYLES, styleElementsPreset } from './stylepresets.js';

// ───────────── 스타일 ─────────────
export const ACCENTS = [
  { label: '검정', hex: '#000000', mid: '#808080' },
  { label: '파랑', hex: '#4472c4' },
  { label: '주황', hex: '#ed7d31' },
  { label: '회색', hex: '#a5a5a5' },
  { label: '금색', hex: '#ffc000' },
  { label: '하늘색', hex: '#5b9bd5' },
  { label: '녹색', hex: '#70ad47' },
];

function mix(hex, other, t) {
  const a = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const b = [1, 3, 5].map((i) => parseInt(other.slice(i, i + 2), 16));
  return `#${a.map((x, i) => Math.round(x + (b[i] - x) * t).toString(16).padStart(2, '0')).join('')}`;
}
export const tint = (hex, t) => mix(hex, '#ffffff', t);
export const shade = (hex, t) => mix(hex, '#000000', t);

/** 표 스타일 60개: 밝게 1~21, 보통 1~28, 어둡게 1~11 (엑셀 이름 · 정의를 그대로 씀 — stylepresets.js) */
const TABLE_KINDS = [['Light', '밝게', 21], ['Medium', '보통', 28], ['Dark', '어둡게', 11]];
export const TABLE_STYLES = [
  // WIXEL 모던 스타일 (엑셀 파일에는 사용자 지정 표 스타일로 저장)
  ...MODERN_STYLES.filter((s) => !s.pivot).map((s) => ({ ...s, get swatch() { return presetSwatch(this.name); } })),
  ...TABLE_KINDS.flatMap(([k, group, n]) => Array.from({ length: n }, (_, i) => ({
    name: `TableStyle${k}${i + 1}`, group, label: `표 스타일 ${group} ${i + 1}`,
    get swatch() { return presetSwatch(this.name); },
  }))),
];
export const TABLE_STYLE_GROUPS = [...new Set(TABLE_STYLES.map((s) => s.group))];
export const DEFAULT_TABLE_STYLE = 'TableStyleMedium2';

/** 엑셀 스타일 이름 → 지원하는 스타일 (알 수 없는 이름은 기본 스타일) */
export function normalizeStyleName(name) {
  if (name === 'None' || name === '') return 'None';
  if (PRESET_STYLES[name]) return name;
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/i.exec(name ?? '');
  if (!m) return DEFAULT_TABLE_STYLE;
  const kind = m[1][0].toUpperCase() + m[1].slice(1).toLowerCase();
  return PRESET_STYLES[`TableStyle${kind}${Number(m[2])}`] ? `TableStyle${kind}${Number(m[2])}` : DEFAULT_TABLE_STYLE;
}

export const styleByName = (name) => TABLE_STYLES.find((s) => s.name === name) ?? null;

// ───────────── 위치 ─────────────
export const headerRow = (t) => (t.header ? t.r1 : null);
export const totalsRow = (t) => (t.totals ? t.r2 : null);
export const dataTop = (t) => t.r1 + (t.header ? 1 : 0);
export const dataBottom = (t) => t.r2 - (t.totals ? 1 : 0);

export function tableAt(sheet, r, c) {
  for (const t of sheet.tables ?? []) if (r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2) return t;
  return null;
}

/** 셀에 입힐 표 서식 (셀에 직접 지정한 서식이 우선) — 엑셀 기본 제공 스타일 정의로 계산 */
export function tableCellStyle(t, r, c) {
  if (t.style === 'None' || t.style === '') return null;
  if (Array.isArray(t.styleElements)) return tablePresetCell(styleElementsPreset(t.styleElements), t, r, c);
  if (!styleByName(t.style)) return null;
  return tablePresetCell(t.style, t, r, c);
}

// ───────────── 이름 ─────────────
/** 표 이름 규칙: 글자나 _ 로 시작, 공백 없음, 셀 주소처럼 보이면 안 됨 */
export function validTableName(name) {
  if (!/^[A-Za-z_À-￿][\w.À-￿]{0,254}$/.test(name)) return false;
  if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^(r|c|rc)$/i.test(name) || /^r\d*c\d*$/i.test(name)) return false;
  return true;
}

export function nextTableName(wb, base = '표') {
  const used = new Set(wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.name.toLowerCase())));
  let n = 1;
  while (used.has(`${base}${n}`.toLowerCase())) n++;
  return `${base}${n}`;
}

export function findTable(wb, name) {
  const low = String(name).toLowerCase();
  for (let si = 0; si < wb.sheets.length; si++) {
    const t = (wb.sheets[si].tables ?? []).find((x) => x.name.toLowerCase() === low);
    if (t) return { si, t };
  }
  return null;
}

/** 머리글 텍스트 목록 → 비어 있거나 겹치지 않는 열 이름 */
export function uniqueNames(values) {
  const out = [];
  const seen = new Set();
  values.forEach((v, i) => {
    let base = String(v ?? '').trim() || `열${i + 1}`;
    let name = base;
    for (let n = 2; seen.has(name.toLowerCase()); n++) name = `${base}${n}`;
    seen.add(name.toLowerCase());
    out.push(name);
  });
  return out;
}

/** 표의 열 이름 (머리글 행이 있으면 그 텍스트) */
export function columnNames(wb, si, t, text = null) {
  const shown = text ?? ((r, c) => {
    const v = wb.getValue(si, r, c);
    return v === null ? '' : typeof v === 'object' ? v.code : String(v);
  });
  const raw = [];
  for (let c = t.c1; c <= t.c2; c++) raw.push(t.header ? shown(t.r1, c) : t.columns?.[c - t.c1] ?? '');
  return uniqueNames(raw);
}

/** 표 열 이름(소문자) → 열 순서. 시트가 바뀌기 전까지 표마다 한 번만 계산 (구조적 참조 수십만 개도 빠르게) */
const colIdxMemo = new WeakMap();
function columnIndex(wb, si, t) {
  const ver = `${wb.sheetVersion?.(si) ?? ''}:${t.c1}:${t.c2}:${t.r1}:${t.header}`;
  const hit = colIdxMemo.get(t);
  if (hit && hit.ver === ver) return hit.idx;
  const idx = new Map(columnNames(wb, si, t).map((n, i) => [n.toLowerCase(), i]));
  colIdxMemo.set(t, { ver, idx });
  return idx;
}

// ───────────── 구조적 참조 ─────────────
const AREA = {
  '#all': 'all', '#모두': 'all', '#data': 'data', '#데이터': 'data', '#headers': 'headers', '#머리글': 'headers',
  '#totals': 'totals', '#요약': 'totals', '#this row': 'thisrow', '#현재 행': 'thisrow', '@': 'thisrow',
};

/** 대괄호 안 글자 → { areas: [...], c1: 열 이름|null, c2: 열 이름|null } (결과는 공유하므로 바꾸지 말 것) */
const specMemo = new Map();
export function parseSpec(spec) {
  let p = specMemo.get(spec);
  if (!p) {
    p = parseSpec0(spec);
    if (specMemo.size > 5000) specMemo.clear();
    specMemo.set(spec, p);
  }
  return p;
}
function parseSpec0(spec) {
  const s = spec.trim();
  const unesc = (x) => x.replace(/'(.)/g, '$1').trim();
  if (s === '') return { areas: ['data'], c1: null, c2: null };
  if (s.startsWith('@')) {
    const rest = s.slice(1).trim();
    const col = rest.startsWith('[') ? rest.slice(1, rest.lastIndexOf(']')) : rest;
    return { areas: ['thisrow'], c1: col ? unesc(col) : null, c2: null };
  }
  if (!s.startsWith('[')) {
    const a = AREA[s.toLowerCase()];
    return a ? { areas: [a], c1: null, c2: null } : { areas: ['data'], c1: unesc(s), c2: null };
  }
  // [[#Headers],[열1]:[열2]] 형태
  const items = [];
  let depth = 0;
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'" && i + 1 < s.length) { cur += ch + s[i + 1]; i++; continue; }
    if (ch === '[') { if (depth++ === 0) { cur = ''; continue; } }
    if (ch === ']') { if (--depth === 0) { items.push({ v: cur }); continue; } }
    if (depth === 0) { if (ch === ':') items.push({ colon: true }); continue; }
    cur += ch;
  }
  const areas = [];
  const cols = [];
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.colon) continue;
    const a = AREA[it.v.trim().toLowerCase()];
    if (a) areas.push(a);
    else cols.push(unesc(it.v));
  }
  return { areas: areas.length ? areas : ['data'], c1: cols[0] ?? null, c2: cols[1] ?? null };
}

/**
 * 구조적 참조 → { si, r1, c1, r2, c2 } | null
 * here: 수식이 있는 셀 { si, r, c } (표 이름이 없거나 @ 일 때 필요)
 */
export function resolveStructRef(wb, tableName, spec, here) {
  let found;
  if (tableName) found = findTable(wb, tableName);
  else if (here) {
    const t = tableAt(wb.sheets[here.si], here.r, here.c);
    found = t ? { si: here.si, t } : null;
  }
  if (!found) return null;
  const { si, t } = found;
  const p = parseSpec(spec);
  let c1 = t.c1;
  let c2 = t.c2;
  if (p.c1) {
    const idx = columnIndex(wb, si, t);
    const a = idx.get(p.c1.toLowerCase()) ?? -1;
    const b = p.c2 ? idx.get(p.c2.toLowerCase()) ?? -1 : a;
    if (a < 0 || b < 0) return null;
    c1 = t.c1 + Math.min(a, b);
    c2 = t.c1 + Math.max(a, b);
  }
  let r1 = Infinity;
  let r2 = -Infinity;
  const add = (a, b) => { if (a <= b) { r1 = Math.min(r1, a); r2 = Math.max(r2, b); } };
  for (const area of p.areas) {
    if (area === 'all') add(t.r1, t.r2);
    else if (area === 'headers') { if (!t.header) return null; add(t.r1, t.r1); }
    else if (area === 'totals') { if (!t.totals) return null; add(t.r2, t.r2); }
    else if (area === 'thisrow') {
      if (!here || here.si !== si || here.r < dataTop(t) || here.r > dataBottom(t)) return null;
      add(here.r, here.r);
    } else add(dataTop(t), Math.max(dataTop(t), dataBottom(t)));
  }
  if (r1 === Infinity) return null;
  return { si, r1, c1, r2, c2 };
}

/** 파일(xlsx)에 쓰는 표준 형식: 표1[[#This Row],[열]] 등 */
export function canonicalRef(tableName, spec, hereTableName) {
  const p = parseSpec(spec);
  const esc = (n) => n.replace(/(['[\]#@])/g, "'$1");
  const name = tableName ?? hereTableName ?? '';
  const areaName = { all: '#All', data: '#Data', headers: '#Headers', totals: '#Totals', thisrow: '#This Row' };
  const cols = p.c1 ? (p.c2 ? `[${esc(p.c1)}]:[${esc(p.c2)}]` : `[${esc(p.c1)}]`) : '';
  const areas = p.areas.filter((a) => a !== 'data' || p.areas.length > 1 || !p.c1);
  if (!areas.length) return p.c2 ? `${name}[${cols}]` : `${name}[${esc(p.c1)}]`;
  if (!cols && areas.length === 1 && areas[0] === 'data') return `${name}[]`;
  if (!cols && areas.length === 1) return `${name}[${areaName[areas[0]]}]`; // 엑셀 표기: 표1[#All]
  const parts = [...areas.map((a) => `[${areaName[a]}]`), ...(cols ? [cols] : [])];
  return `${name}[${parts.join(',')}]`;
}

// ───────────── 자동 확장 ─────────────
/**
 * 방금 값을 넣은 범위(rg) 때문에 커져야 하는 표 → [{ id, r2?, c2? }]
 * 엑셀처럼 표 바로 아래 행이나 바로 오른쪽 열에 입력하면 표가 늘어남
 */
export function expansionFor(sheet, rg) {
  const out = [];
  for (const t of sheet.tables ?? []) {
    const others = (sheet.tables ?? []).filter((x) => x !== t);
    const hits = (b) => others.some((o) => o.r1 <= b.r2 && o.r2 >= b.r1 && o.c1 <= b.c2 && o.c2 >= b.c1);
    // 표 바로 아래에서 시작하거나, 표 안에서 시작해 표 아래까지 이어지는 붙여넣기 (원본 데이터를 한 번에 붙여넣을 때)
    if (!t.totals && rg.r1 <= t.r2 + 1 && rg.r1 >= t.r1 + (t.header ? 1 : 0) && rg.r2 > t.r2 && rg.c1 <= t.c2 && rg.c2 >= t.c1) {
      const nb = { ...t, r2: rg.r2 };
      if (!hits(nb)) out.push({ id: t.id, r2: rg.r2 });
    } else if (rg.c1 === t.c2 + 1 && rg.r1 <= t.r2 && rg.r2 >= t.r1) {
      const nb = { ...t, c2: rg.c2 };
      if (!hits(nb)) out.push({ id: t.id, c2: rg.c2 });
    }
  }
  return out;
}

/** 표를 필터 대상 모양으로 ({ r1(머리글), c1, r2(마지막 데이터 행), c2, criteria, hidden, sort }) */
export function tableFilterRange(t) {
  return { ...(t.filter ?? {}), r1: t.r1, c1: t.c1, r2: dataBottom(t), c2: t.c2, criteria: t.filter?.criteria ?? {}, hidden: t.filter?.hidden ?? {} };
}

export const TOTAL_FUNCS = [
  { id: 'none', label: '없음' },
  { id: 'average', label: '평균', code: 101 },
  { id: 'count', label: '개수', code: 103 },
  { id: 'countNums', label: '숫자 개수', code: 102 },
  { id: 'max', label: '최대값', code: 104 },
  { id: 'min', label: '최소값', code: 105 },
  { id: 'sum', label: '합계', code: 109 },
  { id: 'stdDev', label: '표본 표준 편차', code: 107 },
  { id: 'var', label: '표본 분산', code: 110 },
];
