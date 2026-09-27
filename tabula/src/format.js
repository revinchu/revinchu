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

function dateParts(serial) {
  const d = new Date(EPOCH + Math.round(serial * DAY_MS));
  return {
    y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds(), dow: d.getUTCDay(),
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
export function formatValue(v, style = {}) {
  if (v === null || v === undefined || v === '') return { text: '', align: 'left' };
  if (typeof v === 'object' && 'code' in v) return { text: v.code, align: 'center' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', align: 'center' };
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
      return { value: (Date.UTC(y, mo - 1, d) - EPOCH) / DAY_MS, numFmt: 'date' };
    }
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && +m[1] < 24 && +m[2] < 60 && +(m[3] || 0) < 60) {
    return { value: (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, numFmt: 'time' };
  }
  const up = t.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return { value: up === 'TRUE' };
  return { value: text };
}
