// 검색 중 선택은 임시 상태다. 검색을 지우면 검색 전의 체크 선택을 복원한다.
export function filterSelection(items, initial = null, label = (value) => value) {
  const values = [...new Set(items)];
  const base = new Set(initial ?? values);
  const text = new Map(values.map((value) => [value, `${value}\n${label(value)}`.toLocaleLowerCase()]));
  let query = '', visible = values, selected = base;
  return {
    search(value) {
      const next = String(value).trim().toLocaleLowerCase();
      if (next === query) return;
      const previousVisible = new Set(visible);
      visible = next ? values.filter((item) => text.get(item).includes(next)) : values;
      selected = next ? new Set(visible.filter((item) => !query || !previousVisible.has(item) || selected.has(item))) : base;
      query = next;
    },
    visible: () => visible,
    checked: (value) => selected.has(value),
    toggle(value, checked) { if (checked) selected.add(value); else selected.delete(value); },
    selectVisible(checked) { for (const value of visible) { if (checked) selected.add(value); else selected.delete(value); } },
    result(add = false) { return [...(query && add ? new Set([...base, ...selected]) : selected)]; },
    searching: () => !!query,
  };
}
