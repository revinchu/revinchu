// 표시 단추와 적용된 필터 조건은 서로 독립된 상태다. 열 번호는 시트의 절대 인덱스다.
export function filterButtonVisible(filter, col) {
  return !!filter && !filter.hiddenButtons?.[col];
}

export function filterButtonsVisible(filter, c1, c2) {
  if (!filter) return false;
  for (let col = c1; col <= c2; col++) if (filterButtonVisible(filter, col)) return true;
  return false;
}

export function filterWithButtons(filter, c1, c2, visible) {
  const next = { criteria: {}, hidden: {}, ...filter };
  if (visible) delete next.hiddenButtons;
  else {
    next.hiddenButtons = {};
    for (let col = c1; col <= c2; col++) next.hiddenButtons[col] = true;
  }
  return next;
}
