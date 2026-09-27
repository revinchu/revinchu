// CSV / TSV 읽기·쓰기
export function parseDelimited(text, delim = ',') {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let i = 0;
  const src = text.replace(/\r\n?/g, '\n');
  while (i < src.length) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i += 2; continue; }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') { quoted = true; i++; continue; }
    if (ch === delim) { row.push(field); field = ''; i++; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
    field += ch;
    i++;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function quote(s, delim) {
  return s.includes(delim) || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
}

export function toDelimited(rows, delim = ',', eol = '\r\n') {
  return rows.map((r) => r.map((s) => quote(s, delim)).join(delim)).join(eol);
}

/** 쉼표/탭 중 구분 기호 추정 */
export function guessDelimiter(text) {
  const first = text.split(/\r?\n/, 1)[0];
  return (first.match(/\t/g)?.length ?? 0) > (first.match(/,/g)?.length ?? 0) ? '\t' : ',';
}
