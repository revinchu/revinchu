/** 빈 보이는 셀로만 텍스트가 이어지도록, 정렬 방향마다 허용 폭을 구한다. */
export function cellTextOverflow(axis, column, align, textWidth, blocked) {
  const width = axis.size(column);
  const reach = Math.max(0, textWidth - width) / (align === 'center' ? 2 : 1);
  const room = dir => {
    let available = 0;
    let c = column + dir;
    while (c >= 0 && c < axis.max) {
      c = axis.nextVisible(c, dir);
      if (c < 0 || c >= axis.max || axis.size(c) <= 0 || blocked(c)) break;
      available += axis.size(c);
      if (available >= reach) break;
      c += dir;
    }
    return available;
  };
  return {
    left: align === 'right' || align === 'center' ? room(-1) : 0,
    right: align === 'left' || align === 'center' ? room(1) : 0,
  };
}
