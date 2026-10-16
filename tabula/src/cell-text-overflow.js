// 날짜 범위는 문자열/수식 결과여도 보고서의 한 값이다. 일반 제목·문장을
// 축소하지 않도록 전체 문자열이 유효한 두 날짜인 경우만 구분한다.
export function isCellDateRange(text) {
  if (typeof text !== 'string' || text.length > 64) return false;
  const m = /^\s*(\d{4})([-/.])(\d{1,2})\2(\d{1,2})\s*[~～〜–—−-]\s*(\d{4})([-/.])(\d{1,2})\6(\d{1,2})\s*$/.exec(text);
  if (!m) return false;
  const valid = (year, month, day) => {
    const y = Number(year), mo = Number(month), d = Number(day);
    const leap = y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0);
    const days = mo === 2 ? (leap ? 29 : 28) : [4, 6, 9, 11].includes(mo) ? 30 : 31;
    return y > 0 && mo >= 1 && mo <= 12 && d >= 1 && d <= days;
  };
  return valid(m[1], m[3], m[4]) && valid(m[5], m[7], m[8]);
}

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
