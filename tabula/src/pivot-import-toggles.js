import { child, kids } from './xml.js';
import { keyOf, itemText, itemIdentity, itemProperty } from './cube.js';
import { parseInput } from './format.js';

const uint = (value, fallback = null) => value === undefined ? fallback
  : /^(0|[1-9]\d*)$/.test(String(value)) && Number.isSafeInteger(Number(value)) ? Number(value) : null;
const inside = (a, r, c) => r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2;
const falseAttr = value => value === '0' || value === 'false';
const validArea = a => a && ['r1', 'c1', 'r2', 'c2'].every(k => Number.isSafeInteger(a[k]) && a[k] >= 0)
  && a.r1 <= a.r2 && a.c1 <= a.c2;

// 피벗 레이블은 숫자 모양 문자열도 숫자로 저장할 수 있습니다. 항목 타입은
// items 순번으로 확인하고, 셀은 표시 캡션으로 확인합니다. 수식은 평가하지 않습니다.
function matchesCaption(cell, key, caption, date1904) {
  if (!cell || typeof cell.raw !== 'string' || !cell.raw || cell.formula || cell.fx
    || cell.raw[0] === '=' && cell.inputType !== 'text' && cell.style?.numFmt !== 'text') return false;
  const textInput = cell.inputType === 'text' || cell.inputType !== 'value' && cell.style?.numFmt === 'text';
  const text = !textInput && cell.raw[0] === "'" ? cell.raw.slice(1) : cell.raw;
  if (text === String(caption)) return true;
  if (caption === itemText(key) && typeof key === 'boolean' && text.toUpperCase() === String(key).toUpperCase()) return true;
  if (caption !== itemText(key) || typeof key !== 'number' || !Number.isFinite(key) || textInput) return false;
  // 날짜 서식 숫자는 reader에서 ISO 입력으로 바뀌므로 같은 일련번호인지 확인합니다.
  const value = Object.hasOwn(cell, 'v') ? cell.v : parseInput(cell.raw, date1904).value;
  return typeof value === 'number' && Object.is(value, key);
}

/**
 * 저장된 축 항목과 원문 레이블에서 펼침 단추만 읽습니다. 셀/서식/역할/캐시를
 * 쓰거나 피벗을 계산하지 않으며, 자료가 불완전한 경로는 추측하지 않습니다.
 * root: pivotTableDefinition XML, cache.fields[].items: reader의 의미 값,
 * body: location.ref 범위, cellAt(r,c): 원문 셀 조회.
 */
export function importedPivotToggleButtons({ root, cache, def, body, cellAt } = {}) {
  if (!root || !cache?.fields || !def || def.showExpand === false || !validArea(body) || typeof cellAt !== 'function') return [];
  const location = child(root, 'location')?.attrs;
  const firstDataRow = uint(location?.firstDataRow), firstDataCol = uint(location?.firstDataCol);
  const firstHeaderRow = uint(location?.firstHeaderRow);
  if (firstDataRow === null || firstDataCol === null || firstHeaderRow === null
    || firstDataRow > body.r2 - body.r1 + 1 || firstDataCol > body.c2 - body.c1 + 1) return [];
  const layout = def.layout ?? 'compact';
  const fields = cache.fields, pfs = kids(child(root, 'pivotFields'), 'pivotField'), buttons = [];
  const fieldName = f => fields[f]?.name || '열' + (f + 1);
  const axisOf = tag => {
    const axis = [];
    for (const node of child(root, tag)?.children ?? []) {
      if (node.name !== 'field') continue;
      const f = node.attrs.x === '-2' ? -2 : uint(node.attrs.x);
      if (f === null || f >= fields.length || axis.includes(f)) return null;
      axis.push(f);
    }
    return axis;
  };
  // 필드별 items는 cache sharedItems와 순서가 다를 수 있습니다.
  const itemLists = new Map();
  const listOf = f => {
    if (!itemLists.has(f)) itemLists.set(f, kids(child(pfs[f], 'items'), 'item'));
    return itemLists.get(f);
  };
  const itemAt = (f, index) => {
    if (f === -2) return index < (def.values?.length ?? 0) ? { sigma: true, index } : null;
    const item = listOf(f)[index], x = uint(item?.attrs.x);
    if (!item || item.attrs.t && item.attrs.t !== 'data' || x === null || x >= (fields[f]?.items?.length ?? 0)) return null;
    const key = keyOf(fields[f].items[x]), name = fieldName(f), label = itemText(key);
    const caption = itemProperty(def.itemCaptions?.[name], label) ?? label;
    return { key, name, label, caption, collapsed: falseAttr(item.attrs.sd)
      || (def.collapsed?.[name] ?? []).some(value => itemIdentity(value) === itemIdentity(label)) };
  };
  for (const [fieldTag, itemTag, names, rowAxis] of [
    ['rowFields', 'rowItems', def.rows ?? [], true], ['colFields', 'colItems', def.cols ?? [], false],
  ]) {
    const axis = axisOf(fieldTag), real = axis?.filter(f => f !== -2);
    if (!real || real.length < 2 || real.length !== names.length || real.some((f, i) => fieldName(f) !== names[i])) continue;
    if (rowAxis && !firstDataCol || !rowAxis && firstHeaderRow + axis.length > firstDataRow) continue;
    const lastReal = axis.lastIndexOf(real[real.length - 1]), seen = new Set();
    let previous = [], ordinal = 0;
    for (const node of child(root, itemTag)?.children ?? []) {
      if (node.name !== 'i') continue;
      const position = ordinal++;
      const type = node.attrs.t;
      // i>0인 행은 두 번째 이후 값 필드의 행입니다. 축소된 상위 항목
      // 아래에서는 r이 Σ의 전체 축 위치보다 작아도 새 그룹이 아닙니다.
      const measure = uint(node.attrs.i, 0);
      if (measure === null || rowAxis && measure > 0) continue;
      if (type === 'grand') { previous = []; continue; }
      const base = uint(node.attrs.r, 0), xs = kids(node, 'x');
      if (base === null || base > axis.length || base > previous.length || !xs.length || base + xs.length > axis.length) { previous = []; continue; }
      // r 앞의 경로는 이전 행에서 상속하되 새 경로 뒤의 옛 suffix는 버립니다.
      const path = previous.slice(0, base);
      let valid = true;
      for (const x of xs) {
        const index = uint(x.attrs.v, 0), f = axis[path.length];
        if (index === null || !itemAt(f, index)) { valid = false; break; }
        path.push(index);
      }
      if (!valid) { previous = []; continue; }
      previous = path;
      if (type && type !== 'data') continue; // subtotal 등도 물리 순번에는 포함합니다.
      // 값 번호만 있는 행은 이미 표시한 그룹의 새 항목 레이블이 아닙니다.
      if (rowAxis && axis[base] === -2 && xs.length === 1) continue;
      const ownLevel = path.length - 1;
      for (let level = 0; level < path.length; level++) {
        const f = axis[level];
        if (f === -2 || level === lastReal || rowAxis && layout !== 'tabular' && level !== ownLevel) continue;
        const item = itemAt(f, path[level]);
        if (!item) continue;
        const prefix = [];
        for (let k = 0; k <= level; k++) {
          const at = itemAt(axis[k], path[k]);
          prefix.push([axis[k], at.sigma ? 'sigma:' + at.index : itemIdentity(at.key)]);
        }
        const identity = JSON.stringify(prefix);
        if (seen.has(identity)) continue;
        const r = rowAxis ? body.r1 + firstDataRow + position : body.r1 + firstHeaderRow + level;
        const c = rowAxis ? body.c1 + (layout === 'compact' ? 0 : level) : body.c1 + firstDataCol + position;
        if (!inside(body, r, c) || rowAxis && c >= body.c1 + firstDataCol || !rowAxis && r >= body.r1 + firstDataRow) continue;
        if (!matchesCaption(cellAt(r, c), item.key, item.caption, cache.date1904 ?? def.date1904 ?? false)) continue;
        seen.add(identity);
        buttons.push({ r, c, kind: 'toggle', field: item.name, item: item.label, collapsed: item.collapsed });
      }
    }
  }
  return buttons;
}
