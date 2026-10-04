// 슬라이서 표시 항목과 저장된 선택 집합을 분리한다. 숨김은 필터 값을 변경하지 않는다.
export function slicerDisplayModel(slicer, model, customLists = []) {
  if (!model.items?.length) return model;
  const allItems = model.items;
  let items = slicer.source?.kind === 'pivot' && slicer.showDeleted !== true ? allItems.filter(item => !item.deleted) : allItems;
  if (slicer.customList !== false) {
    const list = customLists.find(values => items.every(item => item.key === '' || values.includes(item.text)));
    if (list) items = [...items].sort((a, b) => a.key === '' ? 1 : b.key === '' ? -1 : list.indexOf(a.text) - list.indexOf(b.text));
  }
  if (slicer.sort === 'desc') items = [...items.filter(item => item.key !== '').reverse(), ...items.filter(item => item.key === '')];
  // Ctrl/다중 선택은 숨겨진 선택도 보존한다. 그 선택 때문에 '데이터 없는 항목 숨기기'를 풀지 않는다.
  if (slicer.hideNoData) items = items.filter(item => item.hasData);
  else if (slicer.noDataLast !== false) items = [...items.filter(item => item.hasData), ...items.filter(item => !item.hasData)];
  if (slicer.markNoData === false) items = items.map(item => ({ ...item, hasData: true }));
  return { ...model, items, allItems };
}
