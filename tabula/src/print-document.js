// 인쇄 페이지 분할과 이미지 PDF 컨테이너. DOM/외부 라이브러리에 의존하지 않는다.
export function splitPrintIndexes(sizes, capacity, repeated = []) {
  const safe = sizes.map(n => Math.max(0, Number(n) || 0));
  const repeats = [...new Set(repeated)].filter(i => Number.isInteger(i) && i >= 0 && i < safe.length);
  const repeatSet = new Set(repeats), repeatSize = repeats.reduce((sum, i) => sum + safe[i], 0);
  const available = Math.max(1, Number(capacity) - repeatSize), pages = [];
  let current = [], used = 0;
  for (let i = 0; i < safe.length; i++) {
    if (repeatSet.has(i)) continue;
    if (current.length && used + safe[i] > available + 0.01) { pages.push([...repeats, ...current]); current = []; used = 0; }
    current.push(i); used += safe[i];
  }
  if (current.length || !pages.length) pages.push([...repeats, ...current]);
  return pages;
}

/** 선택한 보이는 행·열 안에서만 병합을 구성한다. 숨긴 행/열과 인쇄 영역 밖 셀은 확장하지 않는다. */
export function printMergeMap(merges, columns, rows) {
  const result = new Map();
  const intersect = (list, a, b) => {
    let lo = 0, hi = list.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (list[mid] < a) lo = mid + 1; else hi = mid; }
    const out = []; for (let i = lo; i < list.length && list[i] <= b; i++) out.push(list[i]); return out;
  };
  for (const merge of merges || []) {
    const cs = intersect(columns, merge.c1, merge.c2), rs = intersect(rows, merge.r1, merge.r2);
    if (!cs.length || !rs.length) continue;
    const anchor = { r: merge.r1, c: merge.c1, rowSpan: rs.length, colSpan: cs.length };
    for (const r of rs) for (const c of cs) result.set(`${r},${c}`, r === rs[0] && c === cs[0] ? anchor : null);
  }
  return result;
}

function pdfText(value) {
  let result = 'FEFF';
  for (const ch of String(value ?? '').slice(0, 500)) for (let i = 0; i < ch.length; i++) result += ch.charCodeAt(i).toString(16).padStart(4, '0');
  return `<${result}>`;
}

/** JPEG 페이지를 PDF에 그대로 넣는다. 이미지 PDF이므로 본문 텍스트 검색은 지원하지 않는다. */
export function imagePagesPdf(pages, { title = '인쇄 문서' } = {}) {
  if (!Array.isArray(pages) || !pages.length) throw new Error('PDF에 저장할 페이지가 없습니다.');
  const encode = new TextEncoder(), parts = [], offsets = [0]; let size = 0;
  const add = data => { const bytes = typeof data === 'string' ? encode.encode(data) : data; parts.push(bytes); size += bytes.length; };
  const object = (id, content, bytes) => {
    offsets[id] = size; add(`${id} 0 obj\n${content}`);
    if (bytes) { add('\nstream\n'); add(bytes); add('\nendstream'); }
    add('\nendobj\n');
  };
  add('%PDF-1.4\n%WIXEL\n');
  object(1, '<< /Type /Catalog /Pages 2 0 R >>');
  object(2, `<< /Type /Pages /Count ${pages.length} /Kids [${pages.map((_, i) => `${3 + i * 3} 0 R`).join(' ')}] >>`);
  pages.forEach((page, i) => {
    const { jpeg, width, height, paperWidth, paperHeight } = page;
    if (!(jpeg instanceof Uint8Array) || jpeg.length < 4 || jpeg[0] !== 255 || jpeg[1] !== 216 || ![width, height, paperWidth, paperHeight].every(v => Number.isFinite(v) && v > 0)) throw new Error('PDF 페이지 이미지가 올바르지 않습니다.');
    const id = 3 + i * 3, w = Number(paperWidth.toFixed(4)), h = Number(paperHeight.toFixed(4));
    const stream = encode.encode(`q\n${w} 0 0 ${h} 0 0 cm\n/Im0 Do\nQ`);
    object(id, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /XObject << /Im0 ${id + 2} 0 R >> >> /Contents ${id + 1} 0 R >>`);
    object(id + 1, `<< /Length ${stream.length} >>`, stream);
    object(id + 2, `<< /Type /XObject /Subtype /Image /Width ${Math.round(width)} /Height ${Math.round(height)} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpeg.length} >>`, jpeg);
  });
  const infoId = 3 + pages.length * 3;
  object(infoId, `<< /Title ${pdfText(title)} /Producer (WIXEL) >>`);
  const xref = size;
  add(`xref\n0 ${infoId + 1}\n0000000000 65535 f \n`);
  for (let id = 1; id <= infoId; id++) add(`${String(offsets[id]).padStart(10, '0')} 00000 n \n`);
  add(`trailer\n<< /Size ${infoId + 1} /Root 1 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
  const result = new Uint8Array(size); let at = 0;
  for (const bytes of parts) { result.set(bytes, at); at += bytes.length; }
  return result;
}
