// 시트 보호 (DOM 없음): 검토 → 시트 보호 / 셀 서식 → 보호(잠금 · 수식 숨기기)
// 시트 속성 protect = { on: true, hash: '엑셀 16비트 암호 해시' | null, allow: { selectLocked, selectUnlocked, formatCells, … } }
// 셀 서식: locked === false 면 잠기지 않은 셀 (엑셀 기본은 모두 잠김), hideFormula 면 보호 중 수식 입력줄에 수식을 숨김
export const PROTECT_OPTIONS = [
  { id: 'selectLocked', label: '잠긴 셀 선택', attr: 'selectLockedCells', def: true, select: true },
  { id: 'selectUnlocked', label: '잠기지 않은 셀 선택', attr: 'selectUnlockedCells', def: true, select: true },
  { id: 'formatCells', label: '셀 서식', attr: 'formatCells', def: false },
  { id: 'formatColumns', label: '열 서식', attr: 'formatColumns', def: false },
  { id: 'formatRows', label: '행 서식', attr: 'formatRows', def: false },
  { id: 'insertColumns', label: '열 삽입', attr: 'insertColumns', def: false },
  { id: 'insertRows', label: '행 삽입', attr: 'insertRows', def: false },
  { id: 'insertHyperlinks', label: '하이퍼링크 삽입', attr: 'insertHyperlinks', def: false },
  { id: 'deleteColumns', label: '열 삭제', attr: 'deleteColumns', def: false },
  { id: 'deleteRows', label: '행 삭제', attr: 'deleteRows', def: false },
  { id: 'sort', label: '정렬', attr: 'sort', def: false },
  { id: 'autoFilter', label: '자동 필터 사용', attr: 'autoFilter', def: false },
  { id: 'pivotTables', label: '피벗 테이블 및 피벗 차트 사용', attr: 'pivotTables', def: false },
  { id: 'objects', label: '개체 편집', attr: 'objects', def: false },
];

export const defaultAllow = () => Object.fromEntries(PROTECT_OPTIONS.map((o) => [o.id, o.def]));

/** 엑셀의 시트 보호 암호 해시 (파일의 password 속성, 16비트 16진수) */
export function excelHash(pw) {
  if (!pw) return null;
  let h = 0;
  for (let i = pw.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
    h ^= pw.charCodeAt(i);
  }
  h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
  h ^= pw.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

/** 보호 중인지 */
export const isProtected = (sheet) => !!sheet?.protect?.on;
/** 셀이 잠겨 있는지 (서식의 locked 가 false 가 아니면 잠김) */
export const isLockedStyle = (style) => style?.locked !== false;
/** 보호된 시트에서 action 이 허용되는지 */
export const allowed = (sheet, action) => !isProtected(sheet) || !!(sheet.protect.allow ?? defaultAllow())[action];

/** xlsx <sheetProtection> 속성 → protect */
export function protectFromAttrs(a) {
  if (!(a.sheet === '1' || a.sheet === 'true')) return null;
  const on = (v, dflt) => (v === undefined ? dflt : v === '1' || v === 'true');
  const allow = {};
  for (const o of PROTECT_OPTIONS) {
    // select* : 1 = 선택 막음, 나머지: 1(기본) = 막음, 0 = 허용
    allow[o.id] = o.select ? !on(a[o.attr], false) : !on(a[o.attr], true);
  }
  if (a.objects !== undefined) allow.objects = !on(a.objects, false);
  return { on: true, hash: a.password ?? null, allow, modern: a.algorithmName ? { algorithmName: a.algorithmName, hashValue: a.hashValue, saltValue: a.saltValue, spinCount: a.spinCount } : undefined };
}

/** protect → xlsx <sheetProtection …/> */
export function protectXml(p) {
  if (!p?.on) return '';
  const allow = { ...defaultAllow(), ...(p.allow ?? {}) };
  const attrs = [];
  if (p.hash) attrs.push(`password="${p.hash}"`);
  else if (p.modern?.hashValue) attrs.push(`algorithmName="${p.modern.algorithmName}" hashValue="${p.modern.hashValue}" saltValue="${p.modern.saltValue}" spinCount="${p.modern.spinCount}"`);
  attrs.push('sheet="1"');
  attrs.push(`objects="${allow.objects ? 0 : 1}"`, 'scenarios="1"');
  for (const o of PROTECT_OPTIONS) {
    if (o.id === 'objects') continue;
    if (o.select) { if (!allow[o.id]) attrs.push(`${o.attr}="1"`); } else if (allow[o.id]) attrs.push(`${o.attr}="0"`);
  }
  return `<sheetProtection ${attrs.join(' ')}/>`;
}
