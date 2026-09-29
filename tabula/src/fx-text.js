// 텍스트 함수 (DOM 없음)
import {
  ERR, Range, isError, scalar, toNum, toStr, toInt, toBool, optInt, optBool, flat, asRange, lift,
  parseNumberText, wildcardRegex,
} from './fxcore.js';
import { formatWithPattern } from './format.js';

const chars = (s) => [...s];

/** 엑셀 텍스트 연결용 값 → 문자열 (오류는 그대로 던짐) */
function textOf(v) {
  if (isError(v)) throw v;
  return toStr(v);
}

/** 전각/반각 바이트 수 (LENB 등): 한글·한자 등은 2바이트 */
const isWide = (ch) => ch.codePointAt(0) > 0xff;
const byteLen = (s) => chars(s).reduce((n, ch) => n + (isWide(ch) ? 2 : 1), 0);
function byteSlice(s, start, len) {
  // start: 0 기준 바이트 위치
  let pos = 0;
  let out = '';
  for (const ch of s) {
    const w = isWide(ch) ? 2 : 1;
    if (pos >= start && pos + w <= start + len) out += ch;
    pos += w;
  }
  return out;
}

function fixedText(n, decimals, noCommas) {
  const d = Math.max(-15, Math.min(127, decimals));
  let x = n;
  if (d < 0) { const f = 10 ** -d; x = Math.round(Math.abs(n) / f) * f * Math.sign(n); }
  const s = Math.abs(x).toFixed(Math.max(0, Math.min(100, d)));
  const [ip, fp] = s.split('.');
  const intPart = noCommas ? ip : ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (x < 0 && Number(s) !== 0 ? '-' : '') + intPart + (fp ? `.${fp}` : '');
}

/** TEXTSPLIT/TEXTBEFORE 구분 기호 목록 (배열 가능) */
function delimList(v) {
  if (v instanceof Range) return [...v.values()].map(textOf);
  return [textOf(v)];
}

function findDelims(text, delims, matchMode, from) {
  // 가장 앞에 나오는 구분 기호 위치 {i, len}
  const hay = matchMode ? text.toLowerCase() : text;
  let best = null;
  for (const d of delims) {
    if (d === '') continue;
    const i = hay.indexOf(matchMode ? d.toLowerCase() : d, from);
    if (i >= 0 && (!best || i < best.i || (i === best.i && d.length > best.len))) best = { i, len: d.length };
  }
  return best;
}

function splitAll(text, delims, matchMode) {
  const parts = [];
  let pos = 0;
  for (;;) {
    const f = findDelims(text, delims, matchMode, pos);
    if (!f) { parts.push(text.slice(pos)); break; }
    parts.push(text.slice(pos, f.i));
    pos = f.i + f.len;
  }
  return parts;
}

/** n번째 구분 기호 위치 (음수면 뒤에서) */
function nthDelim(text, delims, n, matchMode, matchEnd) {
  const hits = [];
  let pos = 0;
  for (;;) {
    const f = findDelims(text, delims, matchMode, pos);
    if (!f) break;
    hits.push(f);
    pos = f.i + Math.max(1, f.len);
  }
  if (matchEnd) {
    // 텍스트 끝도 구분 기호로 취급
    if (n > 0) hits.push({ i: text.length, len: 0 });
    else hits.unshift({ i: 0, len: 0 });
  }
  const k = n > 0 ? n - 1 : hits.length + n;
  return hits[k] ?? null;
}

const KO_DIGITS = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
const HANJA_DIGITS = ['', '壹', '貳', '參', '四', '伍', '六', '七', '八', '九'];
const KO_UNITS = ['', '십', '백', '천'];
const HANJA_UNITS = ['', '拾', '百', '阡'];
const KO_BIG = ['', '만', '억', '조', '경'];
const HANJA_BIG = ['', '萬', '億', '兆', '京'];

/** 한국어 엑셀 NUMBERSTRING: 1 = 일백이십삼, 2 = 壹百貳拾參, 3 = 일이삼 */
function numberString(n, type) {
  n = Math.round(Math.abs(n));
  if (type === 3) return String(n).split('').map((d) => (d === '0' ? '영' : KO_DIGITS[+d])).join('');
  if (n === 0) return type === 2 ? '零' : '영';
  const digits = type === 2 ? HANJA_DIGITS : KO_DIGITS;
  const units = type === 2 ? HANJA_UNITS : KO_UNITS;
  const big = type === 2 ? HANJA_BIG : KO_BIG;
  const s = String(n);
  let out = '';
  const groups = [];
  for (let i = s.length; i > 0; i -= 4) groups.unshift(s.slice(Math.max(0, i - 4), i));
  groups.forEach((g, gi) => {
    let part = '';
    const gs = g.padStart(4, '0');
    for (let i = 0; i < 4; i++) {
      const d = +gs[i];
      if (!d) continue;
      part += digits[d] + units[3 - i];
    }
    if (part) out += part + big[groups.length - 1 - gi];
  });
  return out;
}

const REGEX_FLAGS = (caseInsensitive) => (caseInsensitive ? 'giu' : 'gu');
function makeRegex(pattern, ci) {
  try {
    return new RegExp(pattern, REGEX_FLAGS(ci));
  } catch {
    throw ERR.VALUE;
  }
}

/** 배열 → 텍스트 (ARRAYTOTEXT/VALUETOTEXT) */
function valueText(v, strict) {
  if (isError(v)) return v.code;
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  return toStr(v);
}

const SCALAR = {
  LEN: ([s]) => chars(toStr(s)).length,
  LENB: ([s]) => byteLen(toStr(s)),
  LEFT: ([s, n]) => {
    const k = optInt(n, 1);
    if (k < 0) throw ERR.VALUE;
    return chars(toStr(s)).slice(0, k).join('');
  },
  LEFTB: ([s, n]) => {
    const k = optInt(n, 1);
    if (k < 0) throw ERR.VALUE;
    return byteSlice(toStr(s), 0, k);
  },
  RIGHT: ([s, n]) => {
    const k = optInt(n, 1);
    if (k < 0) throw ERR.VALUE;
    return k === 0 ? '' : chars(toStr(s)).slice(-k).join('');
  },
  RIGHTB: ([s, n]) => {
    const k = optInt(n, 1);
    if (k < 0) throw ERR.VALUE;
    const t = toStr(s);
    const len = byteLen(t);
    return byteSlice(t, Math.max(0, len - k), k);
  },
  MID: ([s, start, n]) => {
    const st = toInt(start);
    const k = toInt(n);
    if (st < 1 || k < 0) throw ERR.VALUE;
    return chars(toStr(s)).slice(st - 1, st - 1 + k).join('');
  },
  MIDB: ([s, start, n]) => {
    const st = toInt(start);
    const k = toInt(n);
    if (st < 1 || k < 0) throw ERR.VALUE;
    return byteSlice(toStr(s), st - 1, k);
  },
  UPPER: ([s]) => toStr(s).toUpperCase(),
  LOWER: ([s]) => toStr(s).toLowerCase(),
  PROPER: ([s]) => toStr(s).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, p, ch) => p + ch.toUpperCase()),
  TRIM: ([s]) => toStr(s).replace(/ +/g, ' ').replace(/^ | $/g, ''),
  CLEAN: ([s]) => toStr(s).replace(/[\x00-\x1f]/g, ''),
  REPT: ([s, n]) => {
    const k = toInt(n);
    const t = toStr(s);
    if (k < 0 || t.length * k > 32767) throw ERR.VALUE;
    return t.repeat(k);
  },
  SUBSTITUTE: ([s, oldT, newT, inst]) => {
    const text = toStr(s);
    const o = toStr(oldT);
    const nw = toStr(newT);
    if (o === '') return text;
    if (inst === undefined || inst === null) return text.split(o).join(nw);
    const k = toInt(inst);
    if (k < 1) throw ERR.VALUE;
    let idx = -1;
    for (let i = 0; i < k; i++) {
      idx = text.indexOf(o, idx + 1);
      if (idx < 0) return text;
    }
    return text.slice(0, idx) + nw + text.slice(idx + o.length);
  },
  REPLACE: ([s, start, n, nw]) => {
    const st = toInt(start);
    const k = toInt(n);
    if (st < 1 || k < 0) throw ERR.VALUE;
    const c = chars(toStr(s));
    return c.slice(0, st - 1).join('') + toStr(nw) + c.slice(st - 1 + k).join('');
  },
  REPLACEB: ([s, start, n, nw]) => {
    const st = toInt(start);
    const k = toInt(n);
    if (st < 1 || k < 0) throw ERR.VALUE;
    const t = toStr(s);
    const len = byteLen(t);
    return byteSlice(t, 0, st - 1) + toStr(nw) + byteSlice(t, st - 1 + k, len);
  },
  FIND: ([f, s, start]) => {
    const st = optInt(start, 1);
    const text = chars(toStr(s));
    const needle = toStr(f);
    if (st < 1 || st > text.length + 1) throw ERR.VALUE;
    const tail = text.slice(st - 1).join('');
    const idx = tail.indexOf(needle);
    if (idx < 0) throw ERR.VALUE;
    return st + chars(tail.slice(0, idx)).length;
  },
  FINDB: ([f, s, start]) => {
    const st = optInt(start, 1);
    const text = toStr(s);
    const needle = toStr(f);
    if (st < 1) throw ERR.VALUE;
    for (let i = 0, pos = 0; i <= text.length; ) {
      if (pos + 1 >= st && text.startsWith(needle, i)) return pos + 1;
      if (i >= text.length) break;
      const ch = String.fromCodePoint(text.codePointAt(i));
      pos += isWide(ch) ? 2 : 1;
      i += ch.length;
    }
    throw ERR.VALUE;
  },
  SEARCH: ([f, s, start]) => {
    const st = optInt(start, 1);
    const text = chars(toStr(s));
    if (st < 1 || st > text.length + 1) throw ERR.VALUE;
    const pat = toStr(f);
    if (pat === '') return st;
    // 와일드카드: 패턴을 앞쪽 부분 일치 정규식으로
    const re = new RegExp(wildcardRegex(pat).source.slice(1, -1), 'is');
    const tail = text.slice(st - 1).join('');
    const m = re.exec(tail);
    if (!m) throw ERR.VALUE;
    return st + chars(tail.slice(0, m.index)).length;
  },
  EXACT: ([a, b]) => toStr(a) === toStr(b),
  VALUE: ([s]) => {
    const v = scalar(s);
    if (typeof v === 'boolean') throw ERR.VALUE;
    return toNum(v);
  },
  NUMBERVALUE: ([s, dec, grp]) => {
    let t = toStr(s).replace(/\s/g, '');
    const d = dec === undefined || dec === null ? '.' : toStr(dec)[0];
    const g = grp === undefined || grp === null ? ',' : toStr(grp)[0];
    if (!d || d === g) throw ERR.VALUE;
    if (t === '') return 0;
    let pct = 0;
    while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
    const i = t.indexOf(d);
    const intPart = (i < 0 ? t : t.slice(0, i)).split(g).join('');
    const frac = i < 0 ? '' : t.slice(i + 1);
    if (frac.includes(g) || frac.includes(d)) throw ERR.VALUE;
    const n = Number(`${intPart}${frac ? `.${frac}` : ''}`);
    if (!Number.isFinite(n) || intPart + frac === '') throw ERR.VALUE;
    return n / 100 ** pct;
  },
  TEXT: ([v, fmt]) => {
    const x = scalar(v);
    const f = toStr(fmt);
    if (typeof x === 'string') {
      const n = parseNumberText(x);
      if (n === null) return f.includes('@') ? formatWithPattern(x, f) : x;
      return formatWithPattern(n, f);
    }
    if (typeof x === 'boolean') return x ? 'TRUE' : 'FALSE';
    return formatWithPattern(x ?? 0, f);
  },
  FIXED: ([n, d, noCommas]) => fixedText(toNum(n), optInt(d, 2), optBool(noCommas, false)),
  DOLLAR: ([n, d]) => {
    const x = toNum(n);
    const k = optInt(d, 2);
    const s = fixedText(Math.abs(x), k, false);
    return x < 0 ? `-₩${s}` : `₩${s}`;
  },
  WON: ([n, d]) => SCALAR.DOLLAR([n, d]),
  CHAR: ([n]) => {
    const k = toInt(n);
    if (k < 1 || k > 255) throw ERR.VALUE;
    return String.fromCharCode(k);
  },
  UNICHAR: ([n]) => {
    const k = toInt(n);
    if (k < 1 || k > 0x10ffff) throw ERR.VALUE;
    return String.fromCodePoint(k);
  },
  CODE: ([s]) => {
    const t = toStr(s);
    if (!t) throw ERR.VALUE;
    return t.charCodeAt(0);
  },
  UNICODE: ([s]) => {
    const t = toStr(s);
    if (!t) throw ERR.VALUE;
    return t.codePointAt(0);
  },
  ASC: ([s]) => toStr(s).replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0)).replace(/　/g, ' '),
  DBCS: ([s]) => toStr(s).replace(/[!-~]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) + 0xfee0)).replace(/ /g, '　'),
  JIS: ([s]) => SCALAR.DBCS([s]),
  PHONETIC: ([s]) => toStr(s),
  NUMBERSTRING: ([n, type]) => {
    const t = toInt(type);
    if (t < 1 || t > 3) throw ERR.VALUE;
    return numberString(toNum(n), t);
  },
  TEXTBEFORE: ([s, delim, inst, mode, matchEnd, notFound]) => {
    const text = toStr(s);
    const n = optInt(inst, 1);
    if (n === 0 || Math.abs(n) > text.length + 1) throw ERR.VALUE;
    const ds = delimList(delim);
    if (ds.every((d) => d === '')) return n > 0 ? '' : text;
    const hit = nthDelim(text, ds, n, optInt(mode, 0) === 1, optBool(matchEnd, false));
    if (!hit) {
      if (notFound !== undefined && notFound !== null) return scalar(notFound);
      throw ERR.NA;
    }
    return text.slice(0, hit.i);
  },
  TEXTAFTER: ([s, delim, inst, mode, matchEnd, notFound]) => {
    const text = toStr(s);
    const n = optInt(inst, 1);
    if (n === 0 || Math.abs(n) > text.length + 1) throw ERR.VALUE;
    const ds = delimList(delim);
    if (ds.every((d) => d === '')) return n > 0 ? text : '';
    const hit = nthDelim(text, ds, n, optInt(mode, 0) === 1, optBool(matchEnd, false));
    if (!hit) {
      if (notFound !== undefined && notFound !== null) return scalar(notFound);
      throw ERR.NA;
    }
    return text.slice(hit.i + hit.len);
  },
  REGEXTEST: ([s, pattern, ci]) => makeRegex(toStr(pattern), optInt(ci, 0) === 1).test(toStr(s)),
  // 구글 스프레드시트: REGEXMATCH(텍스트, 정규식) — 대소문자 구분
  REGEXMATCH: ([s, pattern]) => makeRegex(toStr(pattern), false).test(toStr(s)),
  REGEXREPLACE: ([s, pattern, repl, occ, ci]) => {
    const text = toStr(s);
    const re = makeRegex(toStr(pattern), optInt(ci, 0) === 1);
    const r = toStr(repl);
    const k = optInt(occ, 0);
    if (k === 0) return text.replace(re, r);
    const all = [...text.matchAll(re)];
    const m = all[k > 0 ? k - 1 : all.length + k];
    if (!m) return text;
    const one = new RegExp(re.source, re.flags.replace('g', ''));
    return text.slice(0, m.index) + m[0].replace(one, r) + text.slice(m.index + m[0].length);
  },
};

export const TEXT = {
  CONCATENATE: (a) => {
    // 배열 인수는 원소별로 (동적 배열 엑셀)
    if (a.some((x) => x instanceof Range && (x.height > 1 || x.width > 1))) {
      return lift((args) => args.map(textOf).join(''))(a);
    }
    return a.map(textOf).join('');
  },
  CONCAT: (a) => {
    const s = flat(a).map(textOf).join('');
    if (s.length > 32767) throw ERR.VALUE;
    return s;
  },
  TEXTJOIN: ([delim, ignoreEmpty, ...rest]) => {
    const ds = delimList(delim);
    const skip = toBool(ignoreEmpty);
    const parts = flat(rest).map(textOf).filter((s) => !(skip && s === ''));
    let out = '';
    parts.forEach((p, i) => { out += (i ? ds[(i - 1) % ds.length] : '') + p; });
    if (out.length > 32767) throw ERR.VALUE;
    return out;
  },
  TEXTSPLIT: ([s, colDelim, rowDelim, ignoreEmpty, mode, pad]) => {
    const text = toStr(s);
    const ci = optInt(mode, 0) === 1;
    const skip = optBool(ignoreEmpty, false);
    const cd = colDelim === undefined || colDelim === null ? [] : delimList(colDelim).filter((d) => d !== '');
    const rd = rowDelim === undefined || rowDelim === null ? [] : delimList(rowDelim).filter((d) => d !== '');
    if (!cd.length && !rd.length) throw ERR.VALUE;
    let lines = rd.length ? splitAll(text, rd, ci) : [text];
    if (skip) lines = lines.filter((l) => l !== '');
    let rows = lines.map((l) => {
      let cells = cd.length ? splitAll(l, cd, ci) : [l];
      if (skip) cells = cells.filter((x) => x !== '');
      return cells;
    });
    if (!rows.length) rows = [['']];
    const w = Math.max(1, ...rows.map((r) => r.length));
    const fill = pad === undefined || pad === null ? ERR.NA : scalar(pad);
    rows = rows.map((r) => (r.length ? r : ['']).concat(Array(w - Math.max(1, r.length)).fill(fill)));
    return new Range(rows);
  },
  ARRAYTOTEXT: ([v, fmt]) => {
    const strict = optInt(fmt, 0) === 1;
    const rg = asRange(v);
    if (!strict) return [...rg.values()].map((x) => valueText(x, false)).join(', ');
    return `{${rg.rows.map((r) => r.map((x) => valueText(x, true)).join(',')).join(';')}}`;
  },
  VALUETOTEXT: lift(([v, fmt]) => valueText(v === undefined ? null : v, optInt(fmt, 0) === 1), [0]),
  T: lift(([v]) => {
    if (isError(v)) throw v;
    return typeof v === 'string' ? v : '';
  }),
  REGEXEXTRACT: ([s, pattern, mode, ci]) => {
    const text = toStr(s);
    const re = makeRegex(toStr(pattern), optInt(ci, 0) === 1);
    const m = optInt(mode, 0);
    if (m === 1) {
      const all = [...text.matchAll(re)].map((x) => [x[0]]);
      if (!all.length) throw ERR.NA;
      return new Range(all);
    }
    const one = new RegExp(re.source, re.flags.replace('g', '')).exec(text);
    if (!one) throw ERR.NA;
    if (m === 2) return new Range([one.slice(1).map((x) => x ?? '')].map((r) => (r.length ? r : [''])));
    return one[0];
  },
};
for (const [k, fn] of Object.entries(SCALAR)) TEXT[k] = lift(fn);

