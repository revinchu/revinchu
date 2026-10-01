import { parseRangeName, MAX_ROWS, MAX_COLS } from './formula.js';

export function parseTableRange(text, sheetName) {
  let body = String(text ?? '').trim().replace(/^=/, '');
  const named = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(body);
  if (named) {
    const name = (named[1] ?? named[2]).replace(/''/g, "'");
    if (name.toLocaleLowerCase() !== String(sheetName).toLocaleLowerCase()) return null;
    body = named[3];
  }
  if (!/^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i.test(body)) return null;
  const range = parseRangeName(body.replace(/\$/g, ''));
  return range && range.r1 >= 0 && range.c1 >= 0 && range.r2 < MAX_ROWS && range.c2 < MAX_COLS ? range : null;
}

export function tableRangeProblem(wb, si, range, ignoreTableId = null) {
  const sheet = wb.sheets[si];
  const overlap = (other) => other && other.r1 <= range.r2 && other.r2 >= range.r1 && other.c1 <= range.c2 && other.c2 >= range.c1;
  if (range.r2 >= MAX_ROWS || range.c2 >= MAX_COLS) return '표의 머리글과 데이터가 시트 범위를 벗어납니다.';
  if (range.c2 - range.c1 + 1 > 500) return '표는 500열까지 만들 수 있습니다.';
  if ((sheet.tables ?? []).some((table) => table.id !== ignoreTableId && overlap(table))) return '표는 다른 표와 겹칠 수 없습니다.';
  if ((sheet.merges ?? []).some(overlap)) return '범위에 병합된 셀이 있습니다. 병합을 해제하고 다시 시도하세요.';
  if ([sheet.pivot, ...(sheet.pivotsExtra ?? [])].some((pivot) => overlap(pivot?.area))) return '범위에 피벗 테이블이 있습니다. 피벗 테이블 밖의 범위를 선택하세요.';
  if (wb.spillsOf(si).some((spill) => overlap({ r1: spill.r, c1: spill.c, r2: spill.r + spill.h - 1, c2: spill.c + spill.w - 1 }))) return '범위에 동적 배열의 분산 결과가 있습니다. 표 밖에 배열 수식을 두거나 값으로 변환하세요.';
  return '';
}
