// 표시 형식: 일반 / 숫자 / 통화 / 회계 / 백분율 / 날짜 / 시간 / 텍스트

const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86400000;
const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

export const NUMBER_FORMATS = [
  { id: 'general', label: '일반' },
  { id: 'number', label: '숫자' },
  { id: 'currency', label: '통화' },
  { id: 'accounting', label: '회계' },
  { id: 'date', label: '간단한 날짜' },
  { id: 'longdate', label: '자세한 날짜' },
  { id: 'time', label: '시간' },
  { id: 'percent', label: '백분율' },
  { id: 'fraction', label: '분수' },
  { id: 'scientific', label: '지수' },
  { id: 'text', label: '텍스트' },
  { id: 'custom', label: '사용자 지정' },
  { id: 'more', label: '기타 표시 형식...' },
];

const DEFAULT_DECIMALS = { number: 2, currency: 0, accounting: 0, percent: 0, scientific: 2, comma: 0 };

/** 엑셀 "일반" 형식 숫자 표시 (최대 유효 자릿수 11) */
export function formatGeneral(n) {
  if (!Number.isFinite(n)) return '#NUM!';
  if (Object.is(n, -0)) n = 0;
  const abs = Math.abs(n);
  if (Number.isInteger(n) && abs < 1e11) return String(n);
  if (abs >= 1e11 || (abs !== 0 && abs < 1e-9)) {
    const [m, e] = n.toExponential(5).split('e');
    const mant = m.replace(/\.?0+$/, '');
    const exp = Number(e);
    return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  const intDigits = abs >= 1 ? Math.floor(Math.log10(abs)) + 1 : 1;
  const decimals = Math.max(0, Math.min(15, 10 - intDigits));
  let s = n.toFixed(decimals);
  if (s.includes('.')) s = s.replace(/\.?0+$/, '');
  return s === '-0' ? '0' : s;
}

function groupThousands(intStr) {
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function fixed(n, decimals, thousands) {
  const s = Math.abs(n).toFixed(decimals);
  const [i, f] = s.split('.');
  const body = (thousands ? groupThousands(i) : i) + (f ? '.' + f : '');
  const neg = n < 0 && Number(s) !== 0;
  return { neg, body };
}

function pad(n, w = 2) { return String(n).padStart(w, '0'); }

/** 날짜 → 일련번호 (엑셀 1900 체계: 1900-03-01 앞은 하루 당김, 1900-01-01 = 1) */
export function serialOf(y, m, d) {
  const s = (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
  return s >= 1 && s < 61 ? s - 1 : s;
}

/**
 * 일련번호 → 날짜 부분. 엑셀 1900 날짜 체계: 0 = 1900-01-00, 1 = 1900-01-01, 60 = 1900-02-29(없는 날),
 * 61 부터 실제 날짜. 요일도 엑셀처럼 1 = 일요일 기준 (step: 반올림 단위 ms)
 */
export function dateParts(serial, step = 1) {
  const ms = Math.round(serial * DAY_MS / step) * step;
  const day = Math.floor(ms / DAY_MS);
  const dow = (((day + 6) % 7) + 7) % 7;
  if (day >= 0 && day < 61) {
    const t = new Date(ms - day * DAY_MS + EPOCH);
    const time = { hh: t.getUTCHours(), mm: t.getUTCMinutes(), ss: t.getUTCSeconds(), dow };
    if (day === 0) return { y: 1900, m: 1, d: 0, ...time };
    if (day === 60) return { y: 1900, m: 2, d: 29, ...time };
    const d = new Date(EPOCH + ms + DAY_MS);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), ...time };
  }
  const d = new Date(EPOCH + ms);
  return {
    y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds(), dow,
  };
}

function toFraction(n) {
  const sign = n < 0 ? '-' : '';
  const x = Math.abs(n);
  const whole = Math.floor(x);
  const frac = x - whole;
  if (frac < 1e-9) return sign + whole;
  let best = { num: 0, den: 1, err: frac };
  for (let den = 1; den <= 9; den++) {
    const num = Math.round(frac * den);
    const err = Math.abs(frac - num / den);
    if (err < best.err - 1e-12) best = { num, den, err };
  }
  if (best.num === 0) return sign + whole;
  if (best.num === best.den) return sign + (whole + 1);
  return `${sign}${whole ? whole + ' ' : ''}${best.num}/${best.den}`;
}

/** 숫자에 셀 서식 적용 */
export function formatNumber(n, fmt = 'general', decimals) {
  const d = decimals ?? DEFAULT_DECIMALS[fmt];
  switch (fmt) {
    case 'number': {
      const { neg, body } = fixed(n, d, true);
      return (neg ? '-' : '') + body;
    }
    case 'comma': {
      const { neg, body } = fixed(n, d, true);
      return (neg ? '-' : '') + body;
    }
    case 'currency':
    case 'accounting': {
      const { neg, body } = fixed(n, d, true);
      return `${neg ? '-' : ''}₩${body}`;
    }
    case 'percent': {
      const { neg, body } = fixed(n * 100, d, false);
      return `${neg ? '-' : ''}${body}%`;
    }
    case 'scientific': {
      const [m, e] = n.toExponential(d).split('e');
      const exp = Number(e);
      return `${m}E${exp < 0 ? '-' : '+'}${pad(Math.abs(exp))}`;
    }
    case 'fraction':
      return toFraction(n);
    case 'date': {
      if (n < 0) return '#'.repeat(8);
      const p = dateParts(n);
      return `${p.y}-${pad(p.m)}-${pad(p.d)}`;
    }
    case 'longdate': {
      if (n < 0) return '#'.repeat(8);
      const p = dateParts(n);
      return `${p.y}년 ${p.m}월 ${p.d}일 ${WEEKDAYS[p.dow]}요일`;
    }
    case 'time': {
      if (n < 0) return '#'.repeat(8);
      const p = dateParts(n);
      const ampm = p.hh < 12 ? '오전' : '오후';
      const h12 = p.hh % 12 || 12;
      return `${ampm} ${h12}:${pad(p.mm)}:${pad(p.ss)}`;
    }
    case 'datetime': {
      if (n < 0) return '#'.repeat(8);
      const p = dateParts(n);
      return `${p.y}-${pad(p.m)}-${pad(p.d)} ${p.hh}:${pad(p.mm)}`;
    }
    default:
      if (decimals != null) {
        const { neg, body } = fixed(n, decimals, false);
        return (neg ? '-' : '') + body;
      }
      return formatGeneral(n);
  }
}

/** TEXT() 함수용 간단한 서식 코드 해석 */
export function formatWithPattern(n, pattern) {
  try {
    return formatCode(n, pattern).text;
  } catch {
    return legacyPattern(n, pattern);
  }
}

function legacyPattern(n, pattern) {
  const p = pattern.trim();
  if (/^[yYmMdDhHsS\-/.: 년월일시분초]+$/.test(p) && /[yYdD]|hh|ss/i.test(p)) {
    const t = dateParts(n);
    // 시:분(:초)를 먼저 처리해야 'mm'이 월로 해석되지 않음
    const TIME = '\u0000';
    const times = [];
    return p
      .replace(/hh?:mm(:ss)?/gi, (_, sec) => { times.push(`${pad(t.hh)}:${pad(t.mm)}${sec ? `:${pad(t.ss)}` : ''}`); return TIME; })
      .replace(/yyyy/gi, t.y).replace(/yy/gi, pad(t.y % 100))
      .replace(/mm/gi, pad(t.m)).replace(/(?<![a-z])m(?![a-z])/gi, t.m)
      .replace(/dd/gi, pad(t.d)).replace(/(?<![a-z])d(?![a-z])/gi, t.d)
      .replace(/hh/gi, pad(t.hh)).replace(/ss/gi, pad(t.ss))
      .replace(new RegExp(TIME, 'g'), () => times.shift());
  }
  const pct = p.endsWith('%');
  const core = pct ? p.slice(0, -1) : p;
  const m = /^([^#0,.]*)([#0,]*)(?:\.(0+|#+))?(.*)$/.exec(core);
  if (!m || !m[2]) return formatGeneral(pct ? n * 100 : n) + (pct ? '%' : '');
  const decimals = m[3] ? m[3].length : 0;
  const { neg, body } = fixed(pct ? n * 100 : n, decimals, m[2].includes(','));
  return `${neg ? '-' : ''}${m[1]}${body}${m[4]}${pct ? '%' : ''}`;
}

/** 현재 표시되는 소수 자릿수 (자릿수 늘림/줄임 용) */
export function displayedDecimals(n, fmt, decimals) {
  if (decimals != null) return decimals;
  if (fmt && fmt !== 'general' && DEFAULT_DECIMALS[fmt] != null) return DEFAULT_DECIMALS[fmt];
  const s = formatGeneral(n);
  const dot = s.indexOf('.');
  return dot < 0 || s.includes('E') ? 0 : s.length - dot - 1;
}

/**
 * 셀 값 → { text, align } (align은 서식에서 정렬을 지정하지 않았을 때 기본값)
 */
export function formatValue(v, style) {
  style ??= {};
  if (v === null || v === undefined || v === '') return { text: '', align: 'left' };
  if (typeof v === 'object' && v.type === 'image') return { text: v.alt || '', align: 'center', image: v };
  if (typeof v === 'object' && 'code' in v) return { text: v.code, align: 'center' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: 'center' };
  if (style.numFmt === 'custom' && style.code) {
    const r = formatCode(v, style.code);
    return { text: r.text, align: typeof v === 'number' ? 'right' : 'left', color: r.color };
  }
  if (typeof v === 'number') {
    if (style.numFmt === 'text') return { text: formatGeneral(v), align: 'left' };
    return { text: formatNumber(v, style.numFmt, style.decimals), align: 'right' };
  }
  return { text: String(v), align: 'left' };
}

/**
 * 사용자가 입력한 문자열 해석 → { value, numFmt? }
 * 수식('=')은 호출하는 쪽에서 처리
 */
export function parseInput(text) {
  if (text === '') return { value: null };
  if (text.startsWith("'")) return { value: text.slice(1) };
  const t = text.trim();
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return { value: Number(t) };
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return { value: Number(t.replace(/,/g, '')), numFmt: 'comma' };
  let m = /^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)%$/.exec(t);
  if (m) {
    const dec = (m[1].split('.')[1] || '').length;
    return { value: Number(m[1].replace(/,/g, '')) / 100, numFmt: 'percent', decimals: dec || undefined };
  }
  m = /^(-?)[₩\\$]\s?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)$/.exec(t);
  if (m) {
    const dec = (m[2].split('.')[1] || '').length;
    return { value: Number(m[1] + m[2].replace(/,/g, '')), numFmt: 'currency', decimals: dec || undefined };
  }
  m = /^(\d{4})\s*[-/.년]\s*(\d{1,2})\s*[-/.월]\s*(\d{1,2})\s*일?$/.exec(t);
  if (m) {
    const y = +m[1];
    const mo = +m[2];
    const d = +m[3];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      return { value: serialOf(y, mo, d), numFmt: 'date' };
    }
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31 && +m[4] < 24 && +m[5] < 60 && +(m[6] || 0) < 60) {
    const day = serialOf(+m[1], +m[2], +m[3]);
    return { value: day + (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400, numFmt: 'datetime' };
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && +m[1] < 24 && +m[2] < 60 && +(m[3] || 0) < 60) {
    return { value: (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, numFmt: 'time' };
  }
  const up = t.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return { value: up === 'TRUE' };
  return { value: text };
}

// ───────────────────────── 사용자 지정 서식 코드 (엑셀 형식) ─────────────────────────
const MONTHS_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const COLOR_NAMES = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', green: '#00ff00', magenta: '#ff00ff', red: '#ff0000', white: '#ffffff', yellow: '#ffff00',
  검정: '#000000', 파랑: '#0000ff', 녹청: '#00ffff', 녹색: '#00ff00', 자홍: '#ff00ff', 빨강: '#ff0000', 흰색: '#ffffff', 노랑: '#ffff00',
};
const COLOR_INDEX = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080'];

/** 세미콜론으로 구역 나누기 (따옴표·\·[] 안은 무시) */
function splitSections(code) {
  const out = [];
  let cur = '';
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const j = code.indexOf('"', i + 1);
      const end = j < 0 ? code.length : j;
      cur += code.slice(i, end + 1);
      i = end;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur += code.slice(i, i + 2);
      i++;
    } else if (ch === '[') {
      const j = code.indexOf(']', i);
      const end = j < 0 ? code.length : j;
      cur += code.slice(i, end + 1);
      i = end;
    } else if (ch === ';') {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

/** 구역 하나 → 토큰 */
function tokenizeSection(src) {
  const toks = [];
  const sec = { toks, color: null, cond: null, korean: false, date: false, elapsed: false, text: false, general: false, ampm: false };
  const lit = (v) => {
    const last = toks[toks.length - 1];
    if (last?.t === 'lit') last.v += v; else toks.push({ t: 'lit', v });
  };
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      lit(src.slice(i + 1, j < 0 ? src.length : j));
      i = j < 0 ? src.length : j + 1;
      continue;
    }
    if (ch === '\\') { lit(src[i + 1] ?? ''); i += 2; continue; }
    if (ch === '_') { lit(' '); i += 2; continue; }
    if (ch === '*') { i += 2; continue; } // 채우기 문자는 무시
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const low = inner.toLowerCase();
      let m;
      if (COLOR_NAMES[low] || COLOR_NAMES[inner]) sec.color = COLOR_NAMES[low] ?? COLOR_NAMES[inner];
      else if ((m = /^(?:color|색)\s*(\d+)$/i.exec(inner))) sec.color = COLOR_INDEX[Number(m[1]) - 1] ?? null;
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)$/i.exec(inner))) sec.cond = { op: m[1], v: Number(m[2]) };
      else if (/^(h+|m+|s+)$/i.test(inner)) { toks.push({ t: 'elapsed', v: low[0], n: inner.length }); sec.date = true; sec.elapsed = true; }
      else if (inner.startsWith('$')) {
        const dash = inner.indexOf('-');
        const sym = inner.slice(1, dash < 0 ? undefined : dash);
        const loc = dash < 0 ? '' : inner.slice(dash + 1).toLowerCase();
        if (/(^|0)412$/.test(loc) || loc.endsWith('412')) sec.korean = true;
        if (sym) lit(sym);
      }
      continue; // [DBNum1] 등은 무시
    }
    if (/^general/i.test(rest) || rest.startsWith('G/표준')) { toks.push({ t: 'general' }); sec.general = true; i += rest.startsWith('G/표준') ? 4 : 7; continue; }
    if (/^(am\/pm|a\/p)/i.test(rest)) {
      const m = /^(am\/pm|a\/p)/i.exec(rest)[0];
      toks.push({ t: 'ampm', v: m });
      sec.ampm = true;
      sec.date = true;
      i += m.length;
      continue;
    }
    if (rest.startsWith('오전/오후')) { toks.push({ t: 'ampm', v: '오전/오후' }); sec.ampm = true; sec.date = true; i += 5; continue; }
    const lc = ch.toLowerCase();
    if ('ymdhsae'.includes(lc) && !(lc === 'e' && /^e[+-]/i.test(rest))) {
      let j = i;
      while (j < src.length && src[j].toLowerCase() === lc) j++;
      toks.push({ t: 'dt', v: lc, n: j - i });
      sec.date = true;
      i = j;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') { toks.push({ t: 'digit', v: ch }); i++; continue; }
    if (ch === '.') { toks.push({ t: 'dot' }); i++; continue; }
    if (ch === ',') { toks.push({ t: 'comma' }); i++; continue; }
    if (ch === '%') { toks.push({ t: 'pct' }); i++; continue; }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) { toks.push({ t: 'exp', v: src[i + 1] }); i += 2; continue; }
    if (ch === '/') { toks.push({ t: 'slash' }); i++; continue; }
    if (ch === '@') { toks.push({ t: 'text' }); sec.text = true; i++; continue; }
    lit(ch);
    i++;
  }
  if (sec.date) {
    // 날짜 구역: 숫자 기호는 소수 초(.0)에만 쓰이므로 나머지는 글자로 취급
    for (let k = 0; k < toks.length; k++) {
      const t = toks[k];
      if (t.t === 'dot' && toks[k + 1]?.t === 'digit' && toks[k + 1].v === '0') {
        let n = 0;
        while (toks[k + 1 + n]?.t === 'digit' && toks[k + 1 + n].v === '0') n++;
        toks.splice(k, n + 1, { t: 'subsec', n });
      } else if (t.t === 'digit' || t.t === 'comma' || t.t === 'pct' || t.t === 'slash') {
        toks[k] = { t: 'lit', v: t.t === 'digit' ? t.v : t.t === 'comma' ? ',' : t.t === 'pct' ? '%' : '/' };
      } else if (t.t === 'dot') toks[k] = { t: 'lit', v: '.' };
    }
    // m: 시(h) 뒤나 초(s) 앞이면 '분'
    const dts = toks.filter((t) => t.t === 'dt' || t.t === 'elapsed');
    dts.forEach((t, k) => {
      if (t.t !== 'dt' || t.v !== 'm' || t.n > 2) return;
      const prev = dts[k - 1];
      const next = dts[k + 1];
      if ((prev && (prev.v === 'h')) || (next && next.v === 's')) t.minute = true;
    });
  }
  return sec;
}

const sectionCache = new Map();
function parseFormat(code) {
  let f = sectionCache.get(code);
  if (!f) {
    f = splitSections(code).map(tokenizeSection);
    if (sectionCache.size > 500) sectionCache.clear();
    sectionCache.set(code, f);
  }
  return f;
}

const testCond = (c, v) => {
  switch (c.op) {
    case '<': return v < c.v;
    case '<=': return v <= c.v;
    case '>': return v > c.v;
    case '>=': return v >= c.v;
    case '=': return v === c.v;
    default: return v !== c.v;
  }
};

function renderDate(sec, n) {
  if (n < 0 && !sec.elapsed) return '#'.repeat(8);
  const subsec = sec.toks.find((t) => t.t === 'subsec')?.n ?? 0;
  const secs = Math.round(n * 86400 * 10 ** subsec) / 10 ** subsec;
  const whole = Math.floor(secs + 1e-9);
  const frac = secs - whole;
  const p = dateParts(Math.floor(whole / 86400) + (whole % 86400) / 86400);
  const hour12 = sec.ampm;
  let out = '';
  for (const t of sec.toks) {
    switch (t.t) {
      case 'lit': out += t.v; break;
      case 'elapsed': {
        const total = t.v === 'h' ? Math.floor(whole / 3600) : t.v === 'm' ? Math.floor(whole / 60) : whole;
        out += String(total).padStart(t.n, '0');
        break;
      }
      case 'subsec': out += `.${frac.toFixed(t.n).slice(2)}`; break;
      case 'ampm': {
        const pm = p.hh >= 12;
        if (t.v === '오전/오후' || sec.korean) out += pm ? '오후' : '오전';
        else if (/^a\/p$/i.test(t.v)) out += (pm ? 'P' : 'A')[t.v[0] === 'a' ? 'toLowerCase' : 'toString']();
        else out += t.v[0] === 'a' ? (pm ? 'pm' : 'am') : pm ? 'PM' : 'AM';
        break;
      }
      case 'dt':
        switch (t.v) {
          case 'y': case 'e': out += t.n <= 2 && t.v === 'y' ? pad(p.y % 100) : String(p.y); break;
          case 'm':
            if (t.minute) out += t.n === 1 ? p.mm : pad(p.mm);
            else if (t.n === 1) out += p.m;
            else if (t.n === 2) out += pad(p.m);
            else if (t.n === 3) out += MONTHS_EN[p.m - 1].slice(0, 3);
            else if (t.n === 5) out += MONTHS_EN[p.m - 1][0];
            else out += MONTHS_EN[p.m - 1];
            break;
          case 'd':
            if (t.n === 1) out += p.d;
            else if (t.n === 2) out += pad(p.d);
            else if (t.n === 3) out += DAYS_EN[p.dow].slice(0, 3);
            else out += DAYS_EN[p.dow];
            break;
          case 'a': out += t.n >= 4 ? `${WEEKDAYS[p.dow]}요일` : WEEKDAYS[p.dow]; break;
          case 'h': {
            const h = hour12 ? p.hh % 12 || 12 : p.hh;
            out += t.n === 1 ? h : pad(h);
            break;
          }
          case 's': out += t.n === 1 ? p.ss : pad(p.ss); break;
          default:
        }
        break;
      default:
    }
  }
  return out;
}

function bestFraction(x, maxDen) {
  let best = { num: 0, den: 1, err: x };
  for (let den = 1; den <= maxDen; den++) {
    const num = Math.round(x * den);
    const err = Math.abs(x - num / den);
    if (err < best.err - 1e-12) best = { num, den, err };
    if (err < 1e-12) break;
  }
  return best;
}

/** 자리 표시자에 숫자 문자열 배치 (오른쪽부터). 그룹 쉼표는 digits 에 이미 반영 */
function placeInt(holders, digits) {
  const out = holders.map(() => '');
  let k = digits.length - 1;
  for (let h = holders.length - 1; h >= 0; h--) {
    if (k >= 0) out[h] = digits[k--];
    else out[h] = holders[h] === '0' ? '0' : holders[h] === '?' ? ' ' : '';
  }
  if (k >= 0 && holders.length) out[0] = digits.slice(0, k + 1) + out[0];
  return out;
}

function renderNumber(sec, n, autoMinus) {
  const toks = sec.toks;
  const neg = n < 0;
  let x = Math.abs(n);
  if (sec.general) {
    const g = formatGeneral(x);
    const body = toks.map((t) => (t.t === 'general' ? g : t.t === 'lit' ? t.v : t.t === 'text' ? '' : '')).join('');
    return (neg && autoMinus ? '-' : '') + body;
  }
  const digitIdx = toks.map((t, i) => (t.t === 'digit' ? i : -1)).filter((i) => i >= 0);
  if (!digitIdx.length) {
    // 숫자 자리 없음: 글자만 (예: "없음";"없음")
    return (neg && autoMinus ? '-' : '') + toks.map((t) => (t.t === 'lit' ? t.v : t.t === 'pct' ? '%' : '')).join('');
  }
  const pctCount = toks.filter((t) => t.t === 'pct').length;
  x *= 100 ** pctCount;
  const lastDigit = digitIdx[digitIdx.length - 1];
  const dotIdx = toks.findIndex((t) => t.t === 'dot');
  const expIdx = toks.findIndex((t) => t.t === 'exp');
  const slashIdx = toks.findIndex((t, i) => t.t === 'slash' && i > digitIdx[0]);
  // 쉼표: 숫자 사이 → 천 단위 구분, 정수부 끝 → 1000으로 나눔
  const intEnd = dotIdx >= 0 ? dotIdx : expIdx >= 0 ? expIdx : slashIdx >= 0 ? slashIdx : lastDigit + 1;
  let grouping = false;
  let scale = 0;
  toks.forEach((t, i) => {
    if (t.t !== 'comma') return;
    const before = digitIdx.some((d) => d < i);
    const afterInInt = digitIdx.some((d) => d > i && d < intEnd);
    const numEnd = expIdx >= 0 ? expIdx : toks.length;
    if (before && afterInInt && i < intEnd) grouping = true;
    else if (before && !digitIdx.some((d) => d > i && d < numEnd)) scale++;
  });
  x /= 1000 ** scale;
  const out = toks.map((t) => (t.t === 'lit' ? t.v : ''));
  const minus = neg && autoMinus;

  if (slashIdx >= 0) {
    // 분수
    const numIdx = [];
    for (let i = slashIdx - 1; i >= 0 && toks[i].t === 'digit'; i--) numIdx.unshift(i);
    const denIdx = [];
    let fixedDen = null;
    let j = slashIdx + 1;
    for (; j < toks.length && toks[j].t === 'digit'; j++) denIdx.push(j);
    if (!denIdx.length && toks[j]?.t === 'lit' && /^\d+/.test(toks[j].v)) fixedDen = Number(/^\d+/.exec(toks[j].v)[0]);
    if (fixedDen === null) {
      const allDigits = denIdx.map((d) => toks[d].v).join('');
      if (/^[1-9]\d*$/.test(allDigits)) fixedDen = Number(allDigits);
    }
    const intIdx = digitIdx.filter((d) => d < (numIdx[0] ?? slashIdx));
    let whole = intIdx.length ? Math.floor(x) : 0;
    let frac = intIdx.length ? x - whole : x;
    let fr;
    if (fixedDen) fr = { num: Math.round(frac * fixedDen), den: fixedDen };
    else fr = bestFraction(frac - Math.floor(frac), 10 ** Math.max(1, denIdx.length) - 1), fr.num += Math.floor(frac) * fr.den;
    if (fr.num === fr.den && intIdx.length) { whole++; fr.num = 0; }
    const intDigits = placeInt(intIdx.map((d) => toks[d].v), whole ? String(whole) : '');
    intIdx.forEach((d, k) => { out[d] = intDigits[k]; });
    if (fr.num === 0 && intIdx.length) {
      // 분수 부분 없음 → 공백
      for (const d of [...numIdx, ...denIdx]) out[d] = ' ';
      out[slashIdx] = ' ';
      if (fixedDen && toks[j]?.t === 'lit') out[j] = out[j].replace(/^\d+/, (s) => ' '.repeat(s.length));
      if (!whole) intIdx.forEach((d, k) => { out[d] = toks[d].v === '0' ? '0' : k === intIdx.length - 1 ? '0' : out[d]; });
    } else {
      const numDigits = placeInt(numIdx.map((d) => toks[d].v), String(fr.num));
      numIdx.forEach((d, k) => { out[d] = numDigits[k]; });
      out[slashIdx] = '/';
      if (!fixedDen || denIdx.length) {
        const ds = String(fr.den);
        denIdx.forEach((d, k) => { out[d] = k < ds.length ? ds[k] : toks[d].v === '0' ? '0' : toks[d].v === '?' ? ' ' : ''; });
        if (ds.length > denIdx.length && denIdx.length) out[denIdx[denIdx.length - 1]] += ds.slice(denIdx.length);
      }
    }
    return (minus ? '-' : '') + out.join('');
  }

  const intIdx = digitIdx.filter((d) => d < intEnd);
  const fracIdx = digitIdx.filter((d) => d > dotIdx && dotIdx >= 0 && (expIdx < 0 || d < expIdx));
  const expDigits = expIdx >= 0 ? digitIdx.filter((d) => d > expIdx) : [];
  let exponent = 0;
  if (expIdx >= 0 && x !== 0) {
    const intCount = Math.max(1, intIdx.length);
    const engineering = intIdx.length > 1 && toks[intIdx[0]].v === '#';
    exponent = Math.floor(Math.log10(x));
    if (engineering) exponent = Math.floor(exponent / intCount) * intCount;
    else exponent -= Math.max(1, intIdx.filter((d) => toks[d].v === '0').length) - 1;
    x /= 10 ** exponent;
    if (Number(x.toFixed(fracIdx.length)) >= 10 ** (engineering ? intCount : Math.max(1, intIdx.filter((d) => toks[d].v === '0').length))) {
      x /= 10;
      exponent += 1;
    }
  }
  const fixedStr = x.toFixed(fracIdx.length);
  let [ip, fp = ''] = fixedStr.split('.');
  if (ip === '0') ip = '';
  if (grouping && ip) ip = groupThousands(ip);
  const intDigits = placeInt(intIdx.map((d) => toks[d].v), ip);
  intIdx.forEach((d, k) => { out[d] = intDigits[k]; });
  if (dotIdx >= 0) out[dotIdx] = '.';
  if (fracIdx.length) {
    let lastNZ = fp.length - 1;
    while (lastNZ >= 0 && fp[lastNZ] === '0') lastNZ--;
    fracIdx.forEach((d, k) => {
      out[d] = k <= lastNZ ? fp[k] : toks[d].v === '0' ? '0' : toks[d].v === '?' ? ' ' : '';
    });
  }
  if (expIdx >= 0) {
    const sign = exponent < 0 ? '-' : toks[expIdx].v === '+' ? '+' : '';
    out[expIdx] = `E${sign}`;
    const es = placeInt(expDigits.map((d) => toks[d].v), String(Math.abs(exponent)));
    expDigits.forEach((d, k) => { out[d] = es[k]; });
  }
  toks.forEach((t, i) => { if (t.t === 'pct') out[i] = '%'; });
  const body = out.join('');
  // 반올림 결과가 0이면 음수 기호를 붙이지 않음
  const isZero = Number(fixedStr) === 0 && exponent === 0;
  return (minus && !isZero ? '-' : '') + body;
}

/**
 * 엑셀 서식 코드로 값 표시. 반환: { text, color } (color 없으면 null)
 * 숫자·날짜·시간·분수·지수·조건·색·텍스트 구역(@)을 지원
 */
export function formatCode(v, code) {
  const secs = parseFormat(code);
  if (typeof v === 'string') {
    const s = secs.length >= 4 ? secs[3] : secs.find((x) => x.text && !x.date && !x.toks.some((t) => t.t === 'digit'));
    if (!s) return { text: v, color: null };
    return { text: s.toks.map((t) => (t.t === 'text' ? v : t.t === 'lit' ? t.v : '')).join(''), color: s.color };
  }
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', color: null };
  const n = v;
  let sec;
  let autoMinus = true;
  const numeric = secs.slice(0, 3).filter((s, i) => i < 2 || !s.text || secs.length > 3 || s.toks.length);
  if (numeric.some((s) => s.cond)) {
    if (numeric[0].cond && testCond(numeric[0].cond, n)) sec = numeric[0];
    else if (numeric[1] && (!numeric[1].cond || testCond(numeric[1].cond, n))) sec = numeric[1];
    else if (numeric[2]) sec = numeric[2];
    else if (!numeric[0].cond) sec = numeric[0];
    if (!sec) return { text: '#'.repeat(8), color: null };
    // 조건 구역에서도 음수는 절댓값 대신 기호 표시 (첫 구역은 엑셀과 같이 '-' 자동)
    autoMinus = sec === numeric[0] || !sec.cond || sec.cond.v >= 0;
    if (sec !== numeric[0] && sec.cond && sec.cond.v <= 0 && /</.test(sec.cond.op)) autoMinus = false;
  } else if (numeric.length === 1 || (numeric.length >= 2 && n >= 0 && (numeric.length === 2 || n !== 0))) {
    sec = numeric[0];
  } else if (n < 0) {
    sec = numeric[1];
    autoMinus = false;
  } else {
    sec = numeric[2] ?? numeric[0];
  }
  if (sec.text && !sec.toks.some((t) => t.t === 'digit' || t.t === 'dt' || t.t === 'general')) {
    // "@" 만 있는 구역에 숫자 → 일반 형식
    return { text: sec.toks.map((t) => (t.t === 'text' ? formatGeneral(n) : t.t === 'lit' ? t.v : '')).join(''), color: sec.color };
  }
  const text = sec.date ? renderDate(sec, sec.elapsed ? n : n) : renderNumber(sec, n, autoMinus);
  return { text, color: sec.color };
}

/** 서식 코드가 날짜/시간 형식인지 */
export function isDateCode(code) {
  return parseFormat(code).slice(0, 1).some((s) => s.date);
}

// ───────────────────────── 기본 형식 ↔ 서식 코드 ─────────────────────────
/** 서식 코드 → 가장 가까운 기본 형식 (정확하지 않을 수 있음) */
export function fmtFromCode(code) {
  const first = code.split(';')[0];
  const plain = first.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, (m) => (/\$[₩$€¥£]/.test(m) ? '₩' : ''));
  const decimals = (plain.match(/\.(0+)/)?.[1].length) || undefined;
  if (plain.trim() === '@') return { numFmt: 'text' };
  if (/general/i.test(plain) && !/[0#]/.test(plain)) return {};
  if (/E[+-]/i.test(plain)) return { numFmt: 'scientific', decimals: decimals ?? 0 };
  const hasDate = /[yd]/i.test(plain) || /(^|[^a-z])m{1,4}([^a-z]|$)/i.test(plain) && !/h/i.test(plain);
  const hasTime = /[hs]/i.test(plain);
  if (hasDate || hasTime) {
    if (hasDate && hasTime) return { numFmt: 'datetime' };
    if (hasTime) return { numFmt: 'time' };
    return { numFmt: /dddd|aaaa/i.test(plain) || /년/.test(code) ? 'longdate' : 'date' };
  }
  if (plain.includes('%')) return { numFmt: 'percent', decimals };
  if (/[₩$€¥£]/.test(plain) || /[₩$€¥£]/.test(code.split(';')[0].replace(/"/g, ''))) return { numFmt: 'currency', decimals };
  if (/\?\/\?/.test(plain)) return { numFmt: 'fraction' };
  if (plain.includes('#,##0') || plain.includes('#,###')) return decimals ? { numFmt: 'number', decimals } : { numFmt: 'comma' };
  if (/0/.test(plain)) return { decimals: decimals ?? 0 };
  return {};
}

/** 기본 표시 형식 → 서식 코드 (사용자 지정이면 그 코드) */
export function fmtCode(style) {
  if (style.numFmt === 'custom') return style.code ?? null;
  const d = style.decimals;
  const dec = (n) => (n ? `.${'0'.repeat(n)}` : '');
  switch (style.numFmt) {
    case 'number': return `#,##0${dec(d ?? 2)}`;
    case 'comma': return `#,##0${dec(d ?? 0)}`;
    case 'currency': case 'accounting': return `"₩"#,##0${dec(d ?? 0)};\\-"₩"#,##0${dec(d ?? 0)}`;
    case 'percent': return `0${dec(d ?? 0)}%`;
    case 'scientific': return `0${dec(d ?? 2)}E+00`;
    case 'fraction': return '# ?/?';
    case 'date': return 'yyyy\\-mm\\-dd';
    case 'longdate': return 'yyyy"년" m"월" d"일" dddd';
    case 'time': return '[$-412]AM/PM h:mm:ss';
    case 'datetime': return 'yyyy\\-mm\\-dd h:mm';
    case 'text': return '@';
    default: return d !== undefined ? `0${dec(d)}` : null;
  }
}


const SAMPLES = [0, 1, -1, 7, 1234.5678, -1234.5678, 0.123456, 45366.5625, 1e-3, 98765432.1];

/**
 * 서식 코드 → 셀 스타일 조각. 기본 형식과 똑같이 보이면 기본 형식으로, 아니면 사용자 지정 코드로 보관
 */
export function styleForCode(code) {
  const c = String(code ?? '').trim();
  if (!c || /^(general|g\/표준)$/i.test(c)) return { numFmt: undefined, decimals: undefined, code: undefined };
  let mapped = {};
  try { mapped = fmtFromCode(c); } catch { /* 무시 */ }
  const same = SAMPLES.every((n) => {
    const r = formatCode(n, c);
    return !r.color && r.text === formatNumber(n, mapped.numFmt, mapped.decimals);
  }) && formatCode('가', c).text === '가';
  if (same) return { numFmt: mapped.numFmt, decimals: mapped.decimals, code: undefined };
  return { numFmt: 'custom', code: c, decimals: undefined };
}

/** 스타일의 현재 서식 코드 (대화상자 표시용) */
export function codeOfStyle(style) {
  return fmtCode(style ?? {}) ?? 'G/표준';
}

/** 자릿수 늘림/줄임: 각 숫자 구역의 마지막 숫자 자리 뒤에 0 을 더하거나 뺌 */
export function adjustCodeDecimals(code, delta) {
  return splitSections(code).map((sec) => {
    const t = tokenizeSection(sec);
    if (t.date || t.text && !t.toks.some((x) => x.t === 'digit')) return sec;
    // 따옴표·대괄호·\x 밖에서 마지막 숫자 자리 위치 찾기 (지수 앞)
    let last = -1;
    let dot = -1;
    let stop = sec.length;
    for (let i = 0; i < sec.length; i++) {
      const ch = sec[i];
      if (ch === '"') { const j = sec.indexOf('"', i + 1); i = j < 0 ? sec.length : j; continue; }
      if (ch === '[') { const j = sec.indexOf(']', i); i = j < 0 ? sec.length : j; continue; }
      if (ch === '\\' || ch === '_' || ch === '*') { i++; continue; }
      if ((ch === 'E' || ch === 'e') && /[+-]/.test(sec[i + 1] ?? '')) { stop = i; break; }
      if (ch === '0' || ch === '#' || ch === '?') last = i;
      if (ch === '.') dot = i;
    }
    if (last < 0 || last > stop) return sec;
    if (delta > 0) return dot >= 0 && dot < last ? `${sec.slice(0, last + 1)}0${sec.slice(last + 1)}` : `${sec.slice(0, last + 1)}.0${sec.slice(last + 1)}`;
    if (dot < 0 || dot > last) return sec;
    if (last === dot + 1) return sec.slice(0, dot) + sec.slice(last + 1);
    return sec.slice(0, last) + sec.slice(last + 1);
  }).join(';');
}
