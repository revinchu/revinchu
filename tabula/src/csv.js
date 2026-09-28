// CSV / TSV 읽기·쓰기
import { ColBuilder, textValue } from './block.js';
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

// ───────────── 큰 CSV: 조각씩 읽어 바로 열 블록으로 ─────────────
/** 글자 조각을 받아 레코드(필드 배열)를 onRow 로 넘김. 따옴표 안 줄바꿈 · "" 처리 */
export class CsvStream {
  constructor(delim, onRow) {
    this.delim = delim;
    this.onRow = onRow;
    this.rest = '';
    this.first = true;
  }

  push(text, last = false) {
    let buf = this.rest + text;
    if (this.first) { buf = buf.replace(/^\ufeff/, ''); this.first = false; }
    const d = this.delim;
    let pos = 0;
    const n = buf.length;
    while (pos < n) {
      const nl = buf.indexOf('\n', pos);
      if (nl < 0 && !last) break;
      const end = nl < 0 ? n : nl;
      const q = buf.indexOf('"', pos);
      if (q < 0 || q >= end) {
        // 빠른 길: 따옴표 없는 줄
        let line = buf.slice(pos, end);
        if (line.endsWith('\r')) line = line.slice(0, -1);
        if (line !== '') this.onRow(line.split(d));
        pos = end + 1;
        continue;
      }
      // 따옴표가 있는 레코드: 글자 단위 (줄바꿈이 따옴표 안에 있을 수 있음)
      const row = [];
      let field = '';
      let quoted = false;
      let i = pos;
      let done = false;
      while (i < n) {
        const ch = buf[i];
        if (quoted) {
          if (ch === '"') {
            if (buf[i + 1] === '"') { field += '"'; i += 2; continue; }
            if (i + 1 >= n && !last) break;
            quoted = false;
            i++;
            continue;
          }
          field += ch;
          i++;
          continue;
        }
        if (ch === '"' && field === '') { quoted = true; i++; continue; }
        if (ch === d) { row.push(field); field = ''; i++; continue; }
        if (ch === '\n' || ch === '\r') {
          row.push(field);
          i += ch === '\r' && buf[i + 1] === '\n' ? 2 : 1;
          done = true;
          break;
        }
        field += ch;
        i++;
      }
      if (!done) {
        if (!last) break; // 다음 조각을 기다림
        row.push(field);
        i = n;
      }
      this.onRow(row);
      pos = i;
    }
    this.rest = pos < n ? buf.slice(pos) : '';
  }

  end() {
    if (this.rest) this.push('', true);
  }
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
