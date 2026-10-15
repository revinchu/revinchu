// 날짜/시간 함수 (DOM 없음). 날짜는 엑셀 일련번호 (1899-12-30 = 0)
import {
  ERR, Range, isError, scalar, toNum, toStr, toInt, optInt, optBool, lift, dateToSerial, serialToDate, todaySerial,
  toDate, DAY_MS,
} from './fxcore.js';
import { parseInput } from './format.js';

const ymd = (s) => serialToDate(toDate(s));
const daysInMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

function makeDate(y, m, d) {
  if (y < 0 || y > 9999) throw ERR.NUM;
  if (y < 1900) y += 1900;
  const s = dateToSerial(y, m, d);
  if (s < 0 || y > 9999) throw ERR.NUM;
  return s;
}

function addMonths(serial, months) {
  const t = serialToDate(serial);
  const m0 = t.m - 1 + months;
  const y = t.y + Math.floor(m0 / 12);
  const m = ((m0 % 12) + 12) % 12 + 1;
  return { y, m, d: t.d };
}

/** WEEKDAY 반환 유형 → 요일(0=일) 변환 */
function weekday(dow, type) {
  switch (type) {
    case 1: case 17: return dow + 1;
    case 2: case 11: return ((dow + 6) % 7) + 1;
    case 3: return (dow + 6) % 7;
    case 12: return ((dow + 5) % 7) + 1;
    case 13: return ((dow + 4) % 7) + 1;
    case 14: return ((dow + 3) % 7) + 1;
    case 15: return ((dow + 2) % 7) + 1;
    case 16: return ((dow + 1) % 7) + 1;
    default: throw ERR.NUM;
  }
}

/** WEEKNUM 유형 → 주 시작 요일 (0=일) */
const WEEK_START = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };

function isoWeek(serial) {
  const t = serialToDate(serial);
  const d = new Date(Date.UTC(t.y, t.m - 1, t.d));
  const day = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - day + 3);
  const firstThu = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d - firstThu) / DAY_MS - 3 + ((firstThu.getUTCDay() + 6) % 7)) / 7);
}

/** 주말 지정: 번호(1~7, 11~17) 또는 "0000011" 문자열 → 요일(0=일)별 주말 여부 */
function weekendMask(w) {
  if (w === undefined || w === null) return [true, false, false, false, false, false, true];
  const v = scalar(w);
  if (typeof v === 'string') {
    if (!/^[01]{7}$/.test(v) || v === '1111111') throw ERR.VALUE;
    // 문자열은 월요일부터
    const mask = Array(7).fill(false);
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = v[i] === '1';
    return mask;
  }
  const n = toInt(v);
  const mask = Array(7).fill(false);
  if (n >= 1 && n <= 7) {
    // 1: 토·일, 2: 일·월, ..., 7: 금·토
    const a = (n + 5) % 7;
    mask[a] = true;
    mask[(a + 1) % 7] = true;
  } else if (n >= 11 && n <= 17) {
    mask[(n - 11) % 7] = true; // 11: 일요일만 … 17: 토요일만
  } else throw ERR.NUM;
  return mask;
}

function holidaySet(h) {
  const set = new Set();
  if (h === undefined || h === null) return set;
  const vals = h instanceof Range ? [...h.values()] : [h];
  for (const v of vals) {
    if (isError(v)) throw v;
    if (v === null || v === '') continue;
    set.add(Math.floor(toNum(v)));
  }
  return set;
}

function networkDays(start, end, mask, hol) {
  let s = Math.floor(start);
  let e = Math.floor(end);
  const sign = s <= e ? 1 : -1;
  if (sign < 0) [s, e] = [e, s];
  let n = 0;
  for (let d = s; d <= e; d++) {
    const dow = serialToDate(d).dow;
    if (!mask[dow] && !hol.has(d)) n++;
  }
  return n * sign;
}

function workday(start, days, mask, hol) {
  if (mask.every(Boolean)) throw ERR.VALUE;
  let d = Math.floor(start);
  let left = Math.trunc(days);
  const step = left >= 0 ? 1 : -1;
  while (left !== 0) {
    d += step;
    const dow = serialToDate(d).dow;
    if (!mask[dow] && !hol.has(d)) left -= step;
  }
  return d;
}

function days360(a, b, european) {
  const x = serialToDate(a);
  const y = serialToDate(b);
  let d1 = x.d;
  let d2 = y.d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const lastFeb = (t) => t.m === 2 && t.d === daysInMonth(t.y, 2);
    if (lastFeb(x)) {
      if (lastFeb(y)) d2 = 30;
      d1 = 30;
    }
    if (d1 === 31) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
  }
  return (y.y - x.y) * 360 + (y.m - x.m) * 30 + (d2 - d1);
}

const isLeap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

export function yearFrac(a, b, basis) {
  if (a > b) [a, b] = [b, a];
  switch (basis) {
    case 0: return days360(a, b, false) / 360;
    case 1: {
      const x = serialToDate(a);
      const y = serialToDate(b);
      if (x.y === y.y || (y.y === x.y + 1 && (x.m > y.m || (x.m === y.m && x.d >= y.d)))) {
        let leap;
        if (x.y === y.y) leap = isLeap(x.y);
        else leap = (isLeap(x.y) && (x.m < 3)) || (isLeap(y.y) && (y.m > 2 || (y.m === 2 && y.d === 29)));
        return (b - a) / (leap ? 366 : 365);
      }
      let total = 0;
      for (let yy = x.y; yy <= y.y; yy++) total += isLeap(yy) ? 366 : 365;
      return (b - a) / (total / (y.y - x.y + 1));
    }
    case 2: return (b - a) / 360;
    case 3: return (b - a) / 365;
    case 4: return days360(a, b, true) / 360;
    default: throw ERR.NUM;
  }
}

function timeValue(text) {
  const p = parseInput(String(text).trim());
  if (typeof p.value !== 'number') throw ERR.VALUE;
  return p.value - Math.floor(p.value);
}

function dateValue(text) {
  const p = parseInput(String(text).trim());
  if (typeof p.value !== 'number' || !/date|time/.test(p.numFmt ?? '')) {
    // 숫자만 있는 텍스트는 날짜가 아님
    throw ERR.VALUE;
  }
  return Math.floor(p.value);
}

const SCALAR = {
  DATE: ([y, m, d]) => makeDate(toInt(y), toInt(m), toInt(d)),
  TIME: ([h, m, s]) => {
    const t = toInt(h) * 3600 + toInt(m) * 60 + toInt(s);
    if (t < 0) throw ERR.NUM;
    return (t % 86400) / 86400;
  },
  DATEVALUE: ([s]) => {
    const v = scalar(s);
    if (typeof v === 'number') throw ERR.VALUE;
    return dateValue(toStr(v));
  },
  TIMEVALUE: ([s]) => {
    const v = scalar(s);
    if (typeof v === 'number') throw ERR.VALUE;
    return timeValue(toStr(v));
  },
  YEAR: ([s]) => ymd(s).y,
  MONTH: ([s]) => ymd(s).m,
  DAY: ([s]) => ymd(s).d,
  HOUR: ([s]) => { const n = toNum(s); if (n < 0) throw ERR.NUM; return serialToDate(n).hh; },
  MINUTE: ([s]) => { const n = toNum(s); if (n < 0) throw ERR.NUM; return serialToDate(n).mm; },
  SECOND: ([s]) => { const n = toNum(s); if (n < 0) throw ERR.NUM; return serialToDate(n).ss; },
  WEEKDAY: ([s, type]) => weekday(ymd(s).dow, optInt(type, 1)),
  WEEKNUM: ([s, type]) => {
    const t = optInt(type, 1);
    if (t === 21) return isoWeek(toDate(s));
    const start = WEEK_START[t];
    if (start === undefined) throw ERR.NUM;
    const serial = toDate(s);
    const x = serialToDate(serial);
    const jan1 = dateToSerial(x.y, 1, 1);
    const jan1dow = serialToDate(jan1).dow;
    const offset = (jan1dow - start + 7) % 7;
    return Math.floor((serial - jan1 + offset) / 7) + 1;
  },
  ISOWEEKNUM: ([s]) => isoWeek(toDate(s)),
  EDATE: ([s, n]) => {
    const t = addMonths(toDate(s), toInt(n));
    return makeDate(t.y, t.m, Math.min(t.d, daysInMonth(t.y, t.m)));
  },
  EOMONTH: ([s, n]) => {
    const t = addMonths(toDate(s), toInt(n));
    return makeDate(t.y, t.m, daysInMonth(t.y, t.m));
  },
  DAYS: ([e, s]) => toDate(e) - toDate(s),
  DAYS360: ([s, e, method]) => days360(toDate(s), toDate(e), optBool(method, false)),
  YEARFRAC: ([s, e, basis]) => yearFrac(toDate(s), toDate(e), optInt(basis, 0)),
  DATEDIF: ([s, e, unit]) => {
    const a = toDate(s);
    const b = toDate(e);
    if (a > b) throw ERR.NUM;
    const x = serialToDate(a);
    const y = serialToDate(b);
    const u = toStr(unit).toUpperCase();
    let months = (y.y - x.y) * 12 + (y.m - x.m);
    if (y.d < x.d) months--;
    switch (u) {
      case 'D': return b - a;
      case 'M': return months;
      case 'Y': return Math.floor(months / 12);
      case 'YM': return ((months % 12) + 12) % 12;
      case 'MD': {
        if (y.d >= x.d) return y.d - x.d;
        const pm = y.m === 1 ? 12 : y.m - 1;
        const py = y.m === 1 ? y.y - 1 : y.y;
        return daysInMonth(py, pm) - x.d + y.d;
      }
      case 'YD': {
        let start = dateToSerial(y.y, x.m, Math.min(x.d, daysInMonth(y.y, x.m)));
        if (start > b) start = dateToSerial(y.y - 1, x.m, Math.min(x.d, daysInMonth(y.y - 1, x.m)));
        return b - start;
      }
      default: throw ERR.NUM;
    }
  },
  DATESTRING: ([s]) => {
    const t = ymd(s);
    return `${String(t.y % 100).padStart(2, '0')}년 ${String(t.m).padStart(2, '0')}월 ${String(t.d).padStart(2, '0')}일`;
  },
};

export const DATE = {
  TODAY: () => todaySerial(),
  NOW: () => {
    const n = new Date();
    return todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
  },
  NETWORKDAYS: ([s, e, hol]) => networkDays(toDate(s), toDate(e), weekendMask(null), holidaySet(hol)),
  'NETWORKDAYS.INTL': ([s, e, w, hol]) => networkDays(toDate(s), toDate(e), weekendMask(w), holidaySet(hol)),
  WORKDAY: ([s, n, hol]) => workday(toDate(s), toNum(n), weekendMask(null), holidaySet(hol)),
  'WORKDAY.INTL': ([s, n, w, hol]) => workday(toDate(s), toNum(n), weekendMask(w), holidaySet(hol)),
};
for (const [k, fn] of Object.entries(SCALAR)) DATE[k] = lift(fn);
// 공휴일 인수는 배열 그대로
for (const k of ['NETWORKDAYS', 'NETWORKDAYS.INTL', 'WORKDAY', 'WORKDAY.INTL']) {
  const fn = DATE[k];
  DATE[k] = lift(fn, k.endsWith('INTL') ? [0, 1, 2] : [0, 1]);
}

