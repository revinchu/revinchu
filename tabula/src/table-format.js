// 표 스타일과 직접 서식의 공통 합성. 계산/DOM 상태를 변경하지 않는다.
import { tableAt, tableCellStyle } from './tables.js';
import { mergeObjectStyle } from './object-styles.js';

const FILL_KEYS = ['fill', 'gradient', 'pattern', 'patternColor'];
const BORDER_KEYS = ['bt', 'bb', 'bl', 'br', 'dd', 'du'];
export const TABLE_VISUAL_KEYS = [
  ...FILL_KEYS, 'color', 'bold', 'italic', 'underline', 'strike',
  ...BORDER_KEYS.flatMap(k => [k, `${k}s`, `${k}c`]),
];
const VISUAL_KEYS = new Set(TABLE_VISUAL_KEYS);
const BOOL_KEYS = new Set(['bold', 'italic', 'underline', 'strike', ...BORDER_KEYS]);
const present = value => Object.fromEntries(Object.entries(value ?? {}).filter(([, v]) => v !== undefined));

/** 표 색을 지울 때 행/열 전체의 직접 서식까지 바꾸지 않도록 셀 안에서만 상속을 취소한다. */
export function clearedTableCellStyle(ownStyle, inheritedStyle) {
  const next = { ...ownStyle }, inherit = {};
  delete next.tableStyleInherit;
  // 가져온 이름 있는 셀 스타일의 시각 서식이 다시 상속되지 않도록 연결만 해제한다.
  delete next.cellStyleName;
  for (const key of TABLE_VISUAL_KEYS) {
    delete next[key];
    if (inheritedStyle?.[key] !== undefined && inheritedStyle[key] !== null) {
      next[key] = inherit[key] = BOOL_KEYS.has(key) ? false : '';
    }
  }
  if (Object.keys(inherit).length) next.tableStyleInherit = inherit;
  return next;
}

/** 나중에 직접 지정한 채우기/글자/테두리는 표 스타일보다 우선한다. */
export function explicitTableStylePatch(style, patch) {
  const next = { ...style, ...patch };
  if (!style?.tableStyleInherit || Object.hasOwn(patch, 'tableStyleInherit')) return next;
  const inherit = { ...style.tableStyleInherit };
  for (const key of Object.keys(patch)) {
    if (!VISUAL_KEYS.has(key)) continue;
    let group = [key];
    if (FILL_KEYS.includes(key)) group = FILL_KEYS;
    else {
      const border = BORDER_KEYS.find(b => key === b || key === `${b}s` || key === `${b}c`);
      if (border) group = [border, `${border}s`, `${border}c`];
    }
    for (const channel of group) delete inherit[channel];
  }
  if (Object.keys(inherit).length) next.tableStyleInherit = inherit;
  else delete next.tableStyleInherit;
  return next;
}

/** 기본 서식 < 표 스타일 < 직접 셀/행/열 서식. 조건부 서식은 호출자가 마지막에 합친다. */
export function tableCellDisplayStyle(wb, si, r, c, table = undefined) {
  const sheet = wb.sheets[si], effective = wb.styleAt(si, r, c);
  table ??= tableAt(sheet, r, c);
  const tableStyle = table && tableCellStyle(table, r, c);
  if (!tableStyle) {
    if (!effective.tableStyleInherit) return effective;
    const copy = { ...effective }; delete copy.tableStyleInherit; return copy;
  }
  const own = wb.getCell(si, r, c)?.style;
  const direct = mergeObjectStyle(present(sheet.allStyle), present(sheet.colStyles[c]), present(sheet.rowStyles[r]), present(own));
  // QUERY/CSV의 동적 표시 형식 힌트는 styleAt에서 계산하므로 그대로 보존한다.
  for (const [key, value] of Object.entries(effective)) {
    if (value !== undefined && !Object.hasOwn(direct, key) && (!VISUAL_KEYS.has(key) || !Object.hasOwn(wb.baseStyle ?? {}, key))) direct[key] = value;
  }
  if (effective.queryFormat != null) {
    for (const key of ['numFmt', 'code', 'queryFormat']) direct[key] = effective[key];
  }
  for (const [key, neutral] of Object.entries(direct.tableStyleInherit ?? {})) {
    if (VISUAL_KEYS.has(key) && direct[key] === neutral) delete direct[key];
  }
  delete direct.tableStyleInherit;
  return mergeObjectStyle(present(wb.baseStyle), tableStyle, direct);
}
