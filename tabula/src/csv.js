// CSV / TSV 읽기·쓰기
import { ColBuilder, textValue } from './block.js';
export function delimitedRowWidth(rows) {
  let width = 1;
  for (const row of rows) if (row.length > width) width = row.length;
  return width;
}
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
  return s.includes(delim) || s.includes('"') || s.includes('\n') || s.includes('\r') ? `"${s.replace(/"/g, '""')}"` : s;
}

export function toDelimited(rows, delim = ',', eol = '\r\n') {
  return rows.map((r) => r.map((s) => quote(s, delim)).join(delim)).join(eol);
}

// CSV output is lazy all the way from rows to UTF-8. Large quoted cells are
// escaped in pieces too, so no row or workbook-sized replacement/join is built.
function* delimitedTextParts(rows, delim, eol, bom) {
  if (bom) yield '\ufeff';
  let firstRow = true;
  for (const row of rows) {
    if (!firstRow) yield eol;
    firstRow = false;
    let firstCell = true;
    for (const value of row) {
      if (!firstCell) yield delim;
      firstCell = false;
      const text = String(value ?? '');
      const quoted = text.includes(delim) || text.includes('"') || text.includes('\n') || text.includes('\r');
      if (quoted) yield '"';
      for (let at = 0; at < text.length; at += 32768) {
        const piece = text.slice(at, at + 32768);
        yield quoted ? piece.replace(/"/g, '""') : piece;
      }
      if (quoted) yield '"';
    }
  }
}
/** Iterable rows → bounded UTF-8 chunks, at most 64 Ki UTF-16 code units each. */
export function* toDelimitedChunks(rows, delim = ',', eol = '\r\n', { bom = false } = {}) {
  const encoder = new TextEncoder(), limit = 65536;
  let pending = '';
  for (const text of delimitedTextParts(rows, delim, eol, bom)) {
    for (let at = 0; at < text.length;) {
      const count = Math.min(limit - pending.length, text.length - at);
      pending += text.slice(at, at + count); at += count;
      if (pending.length === limit) {
        let cut = pending.length;
        const last = pending.charCodeAt(cut - 1);
        if (last >= 0xd800 && last <= 0xdbff) cut--;
        yield encoder.encode(pending.slice(0, cut));
        pending = pending.slice(cut);
      }
    }
  }
  if (pending) yield encoder.encode(pending);
}

/** 쉼표/탭 중 구분 기호 추정 */
export function guessDelimiter(text) {
  const first = text.split(/\r\n?|\n/, 1)[0];
  return (first.match(/\t/g)?.length ?? 0) > (first.match(/,/g)?.length ?? 0) ? '\t' : ',';
}

// ───────────── 큰 CSV: 조각씩 읽어 바로 열 블록으로 ─────────────
/** 글자 조각을 받아 레코드(필드 배열)를 onRow 로 넘김. 따옴표 안 줄바꿈 · "" 처리 */
export class CsvStream {
  constructor(delim, onRow) {
    this.delim = delim;
    this.onRow = onRow;
    this.row = [];
    this.field = '';
    this.quoted = false;
    this.quotePending = false;
    this.skipLF = false;
    this.first = true;
  }

  push(text, last = false) {
    // 작은 CSV와 같이 따옴표 안에서도 CR/CRLF를 LF로 정규화한다.
    // 앞 조각의 마지막 CR은 이미 처리했으므로 뒤따르는 LF만 한 번 건너뛴다.
    if (text.length) {
      if (this.first) { text = text.replace(/^\ufeff/, ''); this.first = false; }
      const endsCR = text.endsWith('\r');
      if (this.skipLF && text[0] === '\n') text = text.slice(1);
      this.skipLF = endsCR;
      text = text.replace(/\r\n?/g, '\n');
    }
    let i = 0;
    if (this.quotePending && text.length) {
      this.quotePending = false;
      if (text[0] === '"') { this.field += '"'; i = 1; }
      else this.quoted = false;
    }
    while (i < text.length) {
      if (this.quoted) {
        const q = text.indexOf('"', i);
        if (q < 0) { this.field += text.slice(i); break; }
        this.field += text.slice(i, q);
        if (q + 1 === text.length && !last) { this.quotePending = true; break; }
        if (text[q + 1] === '"') { this.field += '"'; i = q + 2; }
        else { this.quoted = false; i = q + 1; }
        continue;
      }
      // 완성된 비인용 줄은 기존처럼 native split으로 빠르게 처리한다. 빈 줄도 한 행이다.
      if (!this.row.length && this.field === '') {
        const nl = text.indexOf('\n', i), q = text.indexOf('"', i);
        if (nl >= 0 && (q < 0 || q >= nl)) { this.onRow(text.slice(i, nl).split(this.delim)); i = nl + 1; continue; }
        if (text[i] === '"') { this.quoted = true; i++; continue; }
      } else if (this.field === '' && text[i] === '"') { this.quoted = true; i++; continue; }
      const d = text.indexOf(this.delim, i), nl = text.indexOf('\n', i);
      const at = d < 0 ? nl : nl < 0 ? d : Math.min(d, nl);
      if (at < 0) { this.field += text.slice(i); break; }
      this.field += text.slice(i, at);
      this.row.push(this.field); this.field = '';
      if (text[at] === '\n') { const row = this.row; this.row = []; this.onRow(row); }
      i = at + 1;
    }
    if (last) {
      this.quotePending = false; this.quoted = false;
      // 빈 파일/명시적 빈 인용 필드만 있는 EOF/마지막 줄바꿈은 parseDelimited의 기존 정의를 유지한다.
      if (this.field !== '' || this.row.length) {
        this.row.push(this.field); const row = this.row; this.row = []; this.field = ''; this.onRow(row);
      }
    }
  }

  end() { this.push('', true); }
}

const ISO_DATE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
/** 2026-07-01 → 엑셀 날짜 일련번호 (아니면 null) */
function isoSerial(s) {
  const m = ISO_DATE.exec(s);
  if (!m) return null;
  const t = Date.UTC(+m[1], +m[2] - 1, +m[3]);
  return Number.isFinite(t) ? Math.round(t / 86400000) + 25569 : null;
}

/**
 * 큰 CSV → { header: [글자], block: 열 블록(r0=1, c0=0) } 을 만드는 도구.
 * push(글자 조각) 을 여러 번, 마지막에 finish(). 숫자 · TRUE/FALSE · yyyy-mm-dd 날짜는 값으로
 */
export class CsvBlockReader {
  constructor(delim, capHint = 1 << 16) {
    this.header = null;
    this.cols = [];
    this.fmts = [];
    this.n = 0;
    this.capHint = capHint;
    this.stream = new CsvStream(delim, (fields) => this.row(fields));
  }

  row(fields) {
    if (!this.header) { this.header = fields; return; }
    const i = this.n++;
    for (let j = 0; j < fields.length; j++) {
      const s = fields[j];
      if (s === '') continue;
      let b = this.cols[j];
      if (!b) { b = new ColBuilder(this.capHint); this.cols[j] = b; }
      let v = textValue(s);
      if (typeof v === 'string' && s.length <= 10 && s[4] === '-') {
        const d = isoSerial(s);
        if (d !== null) { v = d; this.fmts[j] ??= { numFmt: 'date' }; }
      }
      b.set(i, v);
    }
  }

  push(text) { this.stream.push(text); }

  finish() {
    this.stream.end();
    const width = Math.max(this.header?.length ?? 0, this.cols.length);
    const cols = Array.from({ length: width }, (_, j) => (this.cols[j] ? this.cols[j].finish(this.n, this.fmts[j] ?? null) : { num: null, str: null, dict: [], fmt: null }));
    return { header: this.header ?? [], block: { r0: 1, c0: 0, n: this.n, ver: 0, cols } };
  }
}
