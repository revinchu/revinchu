// 계산 결과의 역할을 사용하므로 같은 원본 필드를 값 영역에 여러 번 넣어도 정확한 항목을 고릅니다.
export function pivotContextTarget(def, result, r, c) {
  const row = result.grid[r], role = row?.[c]?.role ?? '', meta = result.meta ?? {};
  const body = r - (meta.pageRows ?? 0) - (meta.headerRows ?? 0), rowItem = meta.rowItems?.[body];
  let valueIndex = null, field = null, area = null;
  const value = /^(data|subData|groupData|grandData|grandColData|colSubData|valueHead|grandHead):(\d+)$/.exec(role);
  if (value) valueIndex = +value[2];
  else if (role === 'valueCaption' && def.values?.length === 1) valueIndex = 0;
  const dimension = /^(rowHead|rowItem|rowGroup|rowSub|colItem):(\d+)$/.exec(role);
  if (dimension) {
    area = dimension[1] === 'colItem' ? 'cols' : 'rows'; field = def[area]?.[+dimension[2]] ?? null;
    if (!field && def.valuesOnRows && Number.isInteger(rowItem?.vi)) valueIndex = rowItem.vi;
  } else if (role === 'grandLabel' && def.valuesOnRows && Number.isInteger(rowItem?.vi)) valueIndex = rowItem.vi;
  else if (role === 'pageLabel' || role === 'pageValue') { area = 'pages'; field = def.pages?.[r] ?? null; }
  else if (role === 'colHead') {
    area = 'cols'; const raw = String(row[c]?.raw ?? '').replace(/^'/, '');
    field = def.cols?.find(f => f === raw || def.fieldCaptions?.[f] === raw) ?? def.cols?.[0] ?? null;
  }
  if (!def.values?.[valueIndex]) valueIndex = null;
  if (valueIndex !== null) { field = def.values[valueIndex].field; area = 'values'; }
  const sortField = area && area !== 'values' && area !== 'pages' ? field : def.rows?.[rowItem?.node?.depth] ?? def.rows?.at(-1) ?? def.cols?.at(-1) ?? null;
  return { kind: valueIndex !== null ? 'value' : field ? 'label' : 'pivot', role, field, area, valueIndex, sortField, detail: /^(data|subData|grandData|grandColData|colSubData):/.test(role) && def.enableDrill !== false };
}

export function pivotValueDef(def, index, patch) {
  if (!def.values?.[index]) throw new Error('선택한 값 필드를 찾을 수 없습니다.');
  const values = def.values.map((v, i) => i === index ? { ...v, ...patch } : v), next = { ...def, values };
  for (const k of Object.keys(values[index])) if (values[index][k] === undefined) delete values[index][k];
  if (('numFmt' in patch && JSON.stringify(patch.numFmt ?? null) !== JSON.stringify(def.values[index].numFmt ?? null)) || ('showAs' in patch && (patch.showAs ?? 'normal') !== (def.values[index].showAs ?? 'normal'))) {
    const role = new RegExp(`^(data|subData|groupData|grandData|grandColData|colSubData):${index}$`);
    if (def.cellFmt) next.cellFmt = Object.fromEntries(Object.entries(def.cellFmt).map(([key, st]) => {
      if (!role.test(key)) return [key, st];
      const { numFmt, code, decimals, queryFormat, ...rest } = st; return [key, rest];
    }));
  }
  return next;
}

export function pivotRemoveContextField(def, target) {
  const next = { ...def };
  if (target.area === 'values') {
    const index = target.valueIndex;
    if (!Number.isInteger(index) || !def.values?.[index]) return def;
    next.values = def.values.filter((_, i) => i !== index);
    for (const key of ['sort', 'fieldFilters']) if (def[key]) {
      const out = {};
      for (const [f, rule] of Object.entries(def[key])) {
        if (Number.isInteger(rule?.by) && rule.by === index) { if (key === 'sort') { const { by, ...rest } = rule; out[f] = rest; } }
        else out[f] = Number.isInteger(rule?.by) && rule.by > index ? { ...rule, by: rule.by - 1 } : rule;
      }
      next[key] = out;
    }
    if (def.cellFmt) {
      const out = {};
      for (const [role, st] of Object.entries(def.cellFmt)) {
        const m = /^(data|subData|groupData|grandData|grandColData|colSubData|valueHead|grandHead):(\d+)$/.exec(role);
        if (m && +m[2] === index) continue;
        out[m && +m[2] > index ? `${m[1]}:${+m[2] - 1}` : role] = st;
      }
      next.cellFmt = out;
    }
    if (next.values.length < 2) { delete next.valuesPos; delete next.valuesOnRows; }
  } else if (['rows', 'cols', 'pages'].includes(target.area) && target.field) {
    next[target.area] = (def[target.area] ?? []).filter(f => f !== target.field);
    for (const key of ['filters', 'fieldFilters', 'sort', 'order', 'collapsed']) if (def[key]) { next[key] = { ...def[key] }; delete next[key][target.field]; }
    if (target.area === 'cols' && def.valuesPos != null) { const before = new Set(def.cols.slice(0, def.valuesPos)); next.valuesPos = next.cols.filter(f => before.has(f)).length; }
  }
  return next;
}
