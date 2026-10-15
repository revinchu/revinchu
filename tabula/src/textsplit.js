// 텍스트 나누기 (엑셀 '텍스트 마법사'): 구분 기호 / 너비가 일정함, 열 데이터 서식 (DOM 없음)

/**
 * 구분 기호로 나누기
 * o: { tab, semicolon, comma, space, other (문자열), consecutive (연속 구분 기호를 하나로), qualifier ('"' | "'" | '') }
 */
export function splitDelimited(line, o = {}) {
  const delims = new Set();
  if (o.tab) delims.add('\t');
  if (o.semicolon) delims.add(';');
  if (o.comma) delims.add(',');
  if (o.space) delims.add(' ');
  for (const ch of o.other ?? '') delims.add(ch);
  const q = o.qualifier ?? '"';
  const out = [];
  let cur = '';
  let quoted = false;
  let i = 0;
  let lastWasDelim = false;
  while (i < line.length) {
    const ch = line[i];
    if (q && ch === q && (cur === '' || quoted)) {
      if (quoted && line[i + 1] === q) { cur += q; i += 2; continue; }
      quoted = !quoted;
      i++;
      lastWasDelim = false;
      continue;
    }
    if (!quoted && delims.has(ch)) {
      if (!(o.consecutive && lastWasDelim)) out.push(cur);
      cur = '';
      lastWasDelim = true;
      i++;
      continue;
    }
    cur += ch;
    lastWasDelim = false;
    i++;
  }
  if (!(o.consecutive && lastWasDelim && out.length)) out.push(cur);
  return out;
}

/** 너비가 일정함: breaks 는 나눌 글자 위치(오름차순) */
export function splitFixed(line, breaks) {
  const pos = [...new Set(breaks)].filter((b) => b > 0).sort((a, b) => a - b);
  const out = [];
  let prev = 0;
  for (const b of pos) { out.push(line.slice(prev, b)); prev = b; }
  out.push(line.slice(prev));
  return out.map((s) => s.trim());
}

/** 너비가 일정한 데이터의 나눌 위치 추천: 모든 줄에서 공백인 열이 끝나는 곳 */
export function suggestBreaks(lines) {
  const sample = lines.filter((l) => l.trim()).slice(0, 200);
  if (!sample.length) return [];
  const width = Math.max(...sample.map((l) => l.length));
  const blank = [];
  for (let i = 0; i < width; i++) blank.push(sample.every((l) => i >= l.length || l[i] === ' '));
  const out = [];
  for (let i = 1; i < width; i++) if (blank[i - 1] && !blank[i]) out.push(i);
  return out;
}

export const DATE_ORDERS = ['YMD', 'MDY', 'DMY', 'YDM', 'MYD', 'DYM'];
/** 엑셀 텍스트 마법사의 날짜 순서 이름 */
export const DATE_ORDER_LABEL = { YMD: '년월일', MDY: '월일년', DMY: '일월년', YDM: '년일월', MYD: '월년일', DYM: '일년월' };

const EPOCH = Date.UTC(1899, 11, 30);

/** 순서(YMD 등)에 맞춰 날짜 텍스트 해석 → 'yyyy-mm-dd' (시간이 있으면 ' h:mm:ss') 또는 null */
export function parseDateOrder(text, order = 'YMD') {
  const t = String(text).trim();
  if (!t) return null;
  let time = '';
  let body = t;
  const tm = /\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|오전|오후)?$/i.exec(t);
  if (tm) {
    let h = Number(tm[1]);
    const pm = /pm|오후/i.test(tm[4] ?? '');
    const am = /am|오전/i.test(tm[4] ?? '');
    if (pm && h < 12) h += 12;
    if (am && h === 12) h = 0;
    time = ` ${h}:${tm[2]}${tm[3] ? `:${tm[3]}` : ''}`;
    body = t.slice(0, tm.index);
  }
  let parts = body.split(/[^0-9A-Za-z가-힣]+/).filter(Boolean).map((p) => p.replace(/[년월일]$/, ''));
  if (parts.length === 1 && /^\d+$/.test(parts[0])) {
    const d = parts[0];
    if (d.length === 8) {
      // 4자리 연도가 들어갈 위치에 맞춰 자르기
      const yi = order.indexOf('Y');
      parts = yi === 0 ? [d.slice(0, 4), d.slice(4, 6), d.slice(6)] : yi === 1 ? [d.slice(0, 2), d.slice(2, 6), d.slice(6)] : [d.slice(0, 2), d.slice(2, 4), d.slice(4)];
    } else if (d.length === 6) parts = [d.slice(0, 2), d.slice(2, 4), d.slice(4)];
    else return null;
  }
  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  if (parts.length !== 3) return null;
  const val = {};
  for (let i = 0; i < 3; i++) {
    const key = order[i];
    let p = parts[i];
    const mi = MONTHS.indexOf(p.slice(0, 3).toLowerCase());
    if (!/^\d+$/.test(p)) {
      if (mi < 0) return null;
      p = String(mi + 1);
    }
    val[key] = Number(p);
    if (key === 'Y' && p.length <= 2) val.Y = val.Y < 30 ? 2000 + val.Y : 1900 + val.Y;
  }
  const { Y, M, D } = val;
  if (!(M >= 1 && M <= 12 && D >= 1 && D <= 31 && Y >= 1900 && Y <= 9999)) return null;
  const dt = new Date(Date.UTC(Y, M - 1, D));
  if (dt.getUTCMonth() !== M - 1) return null; // 2월 30일 같은 날짜
  return `${Y}-${String(M).padStart(2, '0')}-${String(D).padStart(2, '0')}${time}`;
}

/** 날짜 텍스트 → 엑셀 일련 번호 (검사용) */
export function dateSerial(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(ymd);
  return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 86400000 : null;
}

/**
 * 열 데이터 서식 적용 → 셀에 넣을 입력 문자열 (null 이면 건너뜀)
 * fmt: 'general' | 'text' | 'date' | 'skip', order: 날짜 순서
 */
export function convertPart(text, fmt = 'general', order = 'YMD', num = null) {
  if (fmt === 'skip') return null;
  let s = String(text ?? '');
  // 텍스트 가져오기 고급 설정: 소수 · 1000 단위 구분 기호를 엑셀 기본(. ,)으로 바꿔 숫자로 인식
  if (num && fmt === 'general' && (num.decimal !== '.' || num.thousand !== ',')) {
    const t = s.trim();
    const esc = (c) => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`^-?[\\d${esc(num.thousand)}]*(${esc(num.decimal)}\\d+)?-?%?$`);
    if (num.decimal !== num.thousand && /\d/.test(t) && re.test(t)) s = t.split(num.thousand).join('').split(num.decimal).join('.');
  }
  if (num && num.trailingMinus === false && /^[\d,.]+-$/.test(s.trim())) return s;
  if (s === '') return '';
  if (fmt === 'text') return `'${s}`;
  if (fmt === 'date') {
    const d = parseDateOrder(s, order);
    return d ?? (s.startsWith('=') ? `'${s}` : s);
  }
  // 일반: 뒤에 붙은 빼기 기호(1234-)는 음수로, 수식처럼 보이는 글자는 텍스트로
  const t = s.trim();
  if (/^[\d,]*\.?\d+-$/.test(t)) return `-${t.slice(0, -1)}`;
  if (/^[=+@]/.test(s) && !/^\+\d/.test(s)) return `'${s}`;
  if (/^-/.test(s) && !/^-[\d.,]+%?$/.test(t)) return `'${s}`;
  return s;
}
