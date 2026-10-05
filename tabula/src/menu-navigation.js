/** DOM 순서와 실제 견본 위치로 메뉴의 다음 실행 항목을 정한다. */
export function menuNavigationTarget(items, current, key, { whole = false } = {}) {
  if (!items.length) return -1;
  const last = items.length - 1;
  if (whole && (key === 'Home' || key === 'End')) return key === 'Home' ? 0 : last;
  if (current < 0 || current > last) return key === 'ArrowUp' || key === 'End' ? last : 0;
  const item = items[current];
  if (!item.group) {
    if (key === 'Home') return 0;
    if (key === 'End') return last;
    if (key === 'ArrowDown' || key === 'ArrowRight') return (current + 1) % items.length;
    if (key === 'ArrowUp' || key === 'ArrowLeft') return (current - 1 + items.length) % items.length;
    return current;
  }
  const peers = items.map((entry, index) => ({ ...entry, index })).filter(entry => entry.group === item.group);
  const center = entry => ({ x: entry.rect.left + entry.rect.width / 2, y: entry.rect.top + entry.rect.height / 2 });
  const origin = center(item);
  const sameRow = (a, b) => Math.abs(center(a).y - center(b).y) <= Math.max(2, Math.min(a.rect.height, b.rect.height) / 2);
  if (key === 'Home' || key === 'End') {
    const row = peers.filter(entry => sameRow(item, entry));
    return (key === 'Home' ? row[0] : row.at(-1))?.index ?? current;
  }
  if (key === 'ArrowLeft' || key === 'ArrowRight') return Math.max(0, Math.min(last, current + (key === 'ArrowRight' ? 1 : -1)));
  if (key !== 'ArrowDown' && key !== 'ArrowUp') return current;
  const direction = key === 'ArrowDown' ? 1 : -1;
  const rows = peers.filter(entry => !sameRow(item, entry) && direction * (center(entry).y - origin.y) > 0);
  if (!rows.length) {
    const boundary = direction > 0 ? peers.at(-1).index + 1 : peers[0].index - 1;
    // 첫 색 위에 '자동' 같은 명령이 없으면 첫 견본을 유지한다.
    return Math.max(0, Math.min(last, boundary));
  }
  const nearest = rows.reduce((best, entry) => Math.abs(center(entry).y - origin.y) < Math.abs(center(best).y - origin.y) ? entry : best);
  const row = rows.filter(entry => sameRow(nearest, entry));
  return row.reduce((best, entry) => Math.abs(center(entry).x - origin.x) < Math.abs(center(best).x - origin.x) ? entry : best).index;
}
