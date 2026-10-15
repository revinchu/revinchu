// CSV / TSV 읽기 (DOM 없음): 따옴표 · 줄바꿈 포함 칸 · 구분자 자동 감지
export function detectSep(text) {
  const line = text.slice(0, 4000).split(/\r?\n/)[0] ?? '';
  const counts = [[',', 0], ['\t', 0], [';', 0], ['|', 0]].map(([c]) => [c, line.split(c).length - 1]);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ',';
}

export function parseCsv(text, sep = null) {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  sep = sep ?? detectSep(text);
  const rows = [];
  let row = [];
  let cell = '';
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch;
    } else if (ch === '"' && cell === '') q = true;
    else if (ch === sep) { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

export function toCsv(rows, sep = ',') {
  return rows.map((r) => r.map((c) => { const s = String(c ?? ''); return /["\n\r]/.test(s) || s.includes(sep) ? `"${s.replace(/"/g, '""')}"` : s; }).join(sep)).join('\r\n');
}
