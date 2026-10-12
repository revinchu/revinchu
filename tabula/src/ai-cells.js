function colName(c) { let text = ''; for (c++; c; c = Math.floor((c - 1) / 26)) text = String.fromCharCode(65 + (c - 1) % 26) + text; return text; }

export function aiAddress(address) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/i.exec(address ?? '');
  if (!match) throw new Error('셀 주소는 A1 형식으로 입력하세요.');
  let c = 0; for (const char of match[1].toUpperCase()) c = c * 26 + char.charCodeAt(0) - 64;
  const r = Number(match[2]);
  if (r > 1048576 || c > 16384) throw new Error('Excel 셀 범위를 초과했습니다.');
  return { r: r - 1, c: c - 1, cell: colName(c - 1) + r };
}
export function aiRange(value, maxCells = 1000) {
  if (typeof value !== 'string' || value.length > 40) throw new Error('A1:B10 형식의 범위를 입력하세요.');
  const parts = value.split(':'); if (parts.length > 2) throw new Error('범위 형식을 확인하세요.');
  const a = aiAddress(parts[0]), b = aiAddress(parts[1] ?? parts[0]);
  if (a.r > b.r || a.c > b.c || (b.r - a.r + 1) * (b.c - a.c + 1) > maxCells) throw new Error(`한 번에 ${maxCells}셀까지 읽을 수 있습니다. 범위를 줄여 주세요.`);
  return { r1: a.r, c1: a.c, r2: b.r, c2: b.c };
}
export function aiChanges(changes) {
  if (!Array.isArray(changes) || !changes.length || changes.length > 200) throw new Error('변경할 셀은 1~200개여야 합니다.');
  const seen = new Set(); let length = 0;
  return changes.map(change => {
    const position = aiAddress(change?.cell);
    if (seen.has(position.cell)) throw new Error('같은 셀을 두 번 변경할 수 없습니다.'); seen.add(position.cell);
    if (typeof change.input !== 'string' || change.input.length > 8192 || (length += change.input.length) > 64000) throw new Error('셀 입력이 너무 깁니다.');
    // External references and asynchronous network functions need separate user workflows.
    if (change.input.startsWith('=') && (/\[[^\]]+\][^!\r\n]*!/.test(change.input) || /\b(?:WEBSERVICE|IMPORT\w*|HYPERLINK|IMAGE|GOOGLE\w*|STOCKHISTORY|TRANSLATE|RTD|DDE)\s*\(/i.test(change.input) || /\|/.test(change.input))) throw new Error('AI 변경에서는 외부 참조·네트워크 수식을 사용할 수 없습니다.');
    return { ...position, input: change.input };
  });
}
export function aiReadCells(book, si, range) {
  const cells = [], sh = book.sheets[si];
  for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) {
    const cell = book.getCell(si, r, c), hidden = sh.protect?.on && book.styleAt(si, r, c)?.hideFormula;
    const value = cell?.cached ?? cell?.v ?? null;
    cells.push({ cell: colName(c) + (r + 1), input: hidden ? null : String(cell?.raw ?? '').slice(0, 8192),
      value: typeof value === 'string' ? value.slice(0, 8192) : typeof value === 'number' || typeof value === 'boolean' ? value : value?.code ?? null });
  }
  return cells;
}
export function aiWritable(book, si, changes) {
  const sheet = book.sheets[si];
  if (!sheet) throw new Error('시트를 찾을 수 없습니다.');
  if (book.props?.markedFinal) throw new Error('최종본으로 표시된 문서는 AI로 변경할 수 없습니다.');
  if (sheet.protect?.on) throw new Error('보호된 시트는 AI로 변경할 수 없습니다.');
  // Conservatively protect derived output without recomputing pivots or spills on a server.
  if (sheet.pivot || sheet.pivotsExtra?.length) throw new Error('피벗 테이블이 있는 시트는 원본 데이터 시트에서 수정하세요.');
  let arrays = false;
  sheet.cells.forEachFormulaRC(cell => { if (cell.maybeArray || cell.cachedArray) arrays = true; });
  if (arrays) throw new Error('배열 수식이 있는 시트는 AI 자동 편집을 지원하지 않습니다. 별도 입력 시트에서 수정하세요.');
  for (const { r, c } of changes) {
    if ((sheet.merges ?? []).some(a => r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2 && (r !== a.r1 || c !== a.c1))) throw new Error('병합된 셀의 첫 셀만 변경할 수 있습니다.');
    const cell = book.getCell(si, r, c);
    if (cell?.array || cell?.spill || cell?.spillRef || cell?.arrayRef) throw new Error('배열 수식의 일부는 변경할 수 없습니다.');
  }
}
