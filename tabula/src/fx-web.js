// 웹 · 구글 스프레드시트 전용 함수 (DOM 없음)
// IMPORTXML · IMPORTHTML · IMPORTDATA · IMPORTFEED · IMPORTRANGE · GOOGLEFINANCE · GOOGLETRANSLATE · QUERY · SPARKLINE …
// 네트워크 결과는 NET 캐시에 모아 두고(비동기), 받는 동안은 '로딩 중…' 을 돌려줌. 받으면 NET.onDone 이 요청한 시트를 다시 계산
import {
  ERR, Range, isError, toNum, toStr, optNum, optInt, optBool, asRange, lift, parseNumberText, makeCriteria,
  compareValues, CellImage, collectNums, serialToDate, dateToSerial, FormulaError,
} from './fxcore.js';

// ───────────── 네트워크 캐시 ─────────────
/**
 * fetcher(url) → Promise<string> (앱이 설정: 서버 프록시 /api/fetch 또는 직접 fetch)
 * onDone(sheetIndexes) → 결과가 도착하면 그 URL 을 요청한 시트를 다시 계산
 */
export const NET = { cache: new Map(), fetcher: null, onDone: null, waiting: 0 };
export const LOADING = '로딩 중…';

/** URL 의 텍스트: 있으면 문자열, 받는 중이면 LOADING, 실패하면 ERR.NA (NET.cache 의 message 에 이유) */
export function netText(url, ctx) {
  const hit = NET.cache.get(url);
  const si = ctx?.here?.si;
  if (hit) {
    if (hit.state === 'ok') return hit.data;
    if (hit.state === 'loading') { if (si != null) hit.sheets.add(si); return LOADING; }
    return ERR.NA;
  }
  if (!NET.fetcher) return ERR.NA;
  const rec = { state: 'loading', sheets: new Set(si != null ? [si] : []), t: Date.now() };
  NET.cache.set(url, rec);
  NET.waiting++;
  Promise.resolve().then(() => NET.fetcher(url)).then((text) => {
    rec.state = 'ok';
    rec.data = String(text ?? '');
  }, (e) => {
    rec.state = 'error';
    rec.message = String(e?.message ?? e ?? '가져오지 못했습니다');
  }).finally(() => {
    NET.waiting--;
    NET.onDone?.([...rec.sheets], url);
  });
  return LOADING;
}
/** 다시 가져오기 (F9 전체 계산) */
export function netClear() { for (const [k, v] of NET.cache) if (v.state !== 'loading') NET.cache.delete(k); }
const pending = (v) => v === LOADING || isError(v);

// ───────────── 도우미 ─────────────
const one = (v) => (v instanceof Range ? v.rows[0]?.[0] ?? null : v);
const str1 = (v) => { const x = one(v); if (isError(x)) throw x; return toStr(x ?? ''); };
const grid = (rows) => {
  if (!rows.length) return ERR.NA;
  const w = Math.max(1, ...rows.map((r) => r.length));
  return new Range(rows.map((r) => { const o = r.slice(); while (o.length < w) o.push(''); return o; }));
};
/** 가져온 글자 → 값 (숫자처럼 보이면 숫자: 스프레드시트와 같음) */
export function autoValue(s) {
  const t = String(s ?? '').trim();
  if (t === '') return '';
  if (/^[+-]?[\d,]*\.?\d+(e[+-]?\d+)?%?$/i.test(t) && !/^0\d/.test(t.replace(/^[+-]/, ''))) {
    const n = parseNumberText(t);
    if (n !== null) return n;
  }
  if (/^(true|false)$/i.test(t)) return t.toUpperCase() === 'TRUE';
  return t;
}
const isBlankV = (v) => v === null || v === undefined || v === '';

// ───────────── HTML · XML 파서 (관대한 트리) ─────────────
const VOID = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr']);
const RAW = new Set(['script', 'style', 'textarea', 'title']);
// 같은 태그가 열리면 앞의 것을 닫는 태그 (p · li · tr · td …)
const AUTO_CLOSE = { li: ['li'], p: ['p'], tr: ['tr', 'td', 'th'], td: ['td', 'th'], th: ['td', 'th'], option: ['option'], dt: ['dt', 'dd'], dd: ['dt', 'dd'], thead: ['tbody', 'thead', 'tr', 'td', 'th'], tbody: ['thead', 'tbody', 'tr', 'td', 'th'] };
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0', copy: '©', reg: '®', middot: '·', hellip: '…', ndash: '–', mdash: '—', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', times: '×', laquo: '«', raquo: '»', bull: '•', trade: '™', euro: '€', won: '₩', yen: '¥', pound: '£' };
export function decodeEntities(s) {
  return String(s).replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e) => {
    if (e[0] === '#') {
      const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(n) && n > 0 && n < 0x110000 ? String.fromCodePoint(n) : m;
    }
    return ENT[e.toLowerCase()] ?? m;
  });
}
/** 태그 트리: { tag, attrs, children: [노드 | 문자열], parent } — 태그 이름은 소문자 (xml 모드는 원래대로) */
export function parseMarkup(text, { xml = false } = {}) {
  const root = { tag: '#root', attrs: {}, children: [], parent: null };
  let cur = root;
  const src = String(text);
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<!(?:doctype|DOCTYPE)[^>]*>|<\?[\s\S]*?\?>|<(\/?)([A-Za-z][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let last = 0;
  let m;
  const addText = (t) => { if (t) cur.children.push(decodeEntities(t)); };
  const norm = (n) => (xml ? n : n.toLowerCase());
  const close = (name) => {
    for (let n = cur; n && n !== root; n = n.parent) {
      if (n.tag === name || (!xml && n.tag.toLowerCase() === name)) { cur = n.parent; return; }
    }
  };
  while ((m = re.exec(src))) {
    addText(src.slice(last, m.index));
    last = re.lastIndex;
    if (m[1] !== undefined) { cur.children.push(m[1]); continue; }
    if (!m[3]) continue; // 주석 · doctype · 처리 지시
    const name = norm(m[3]);
    if (m[2]) { close(name); continue; }
    if (!xml && AUTO_CLOSE[name]) {
      // 열린 li 안에서 li 가 또 열리면 앞의 li 를 닫음 (표 · 목록 경계를 넘지는 않음)
      for (let n = cur; n && n !== root; n = n.parent) {
        if (['table', 'ul', 'ol', 'dl', 'select'].includes(n.tag) && !['thead', 'tbody'].includes(name)) break;
        if (AUTO_CLOSE[name].includes(n.tag)) { cur = n.parent; break; }
      }
    }
    const attrs = {};
    const ar = /([^\s=/]+)(?:\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+)))?/g;
    let a;
    while ((a = ar.exec(m[4]))) attrs[xml ? a[1] : a[1].toLowerCase()] = decodeEntities(a[3] ?? a[4] ?? a[5] ?? '');
    const node = { tag: name, attrs, children: [], parent: cur };
    cur.children.push(node);
    if (m[5] || (!xml && VOID.has(name))) continue;
    if (!xml && RAW.has(name)) {
      const end = src.toLowerCase().indexOf(`</${name}`, last);
      const stop = end < 0 ? src.length : end;
      if (name === 'title' || name === 'textarea') node.children.push(decodeEntities(src.slice(last, stop)));
      else node.children.push(src.slice(last, stop));
      const gt = end < 0 ? src.length : src.indexOf('>', end);
      last = gt < 0 ? src.length : gt + 1;
      re.lastIndex = last;
      continue;
    }
    cur = node;
  }
  addText(src.slice(last));
  return root;
}
const isNode = (n) => n && typeof n === 'object' && n.tag;
/** 노드의 글자 (공백 정리, <br> 은 줄바꿈) */
export function textOf(n, raw = false) {
  if (typeof n === 'string') return raw ? n : n.replace(/\s+/g, ' ').trim();
  let out = '';
  const walk = (x) => {
    if (typeof x === 'string') { out += x; return; }
    if (!raw && (x.tag === 'script' || x.tag === 'style')) return;
    if (x.tag === 'br') { out += '\n'; return; }
    for (const ch of x.children) walk(ch);
    if (!raw && ['p', 'div', 'li', 'tr'].includes(x.tag)) out += ' ';
  };
  walk(n);
  return raw ? out : out.split('\n').map((l) => l.replace(/\s+/g, ' ').trim()).filter((l, i, a) => l || i < a.length - 1).join('\n').trim();
}
function* descendants(n) {
  for (const ch of n.children) {
    if (!isNode(ch)) continue;
    yield ch;
    yield* descendants(ch);
  }
}
const localName = (t) => t.replace(/^.*:/, '').toLowerCase();

// ───────────── XPath (자주 쓰는 부분) ─────────────
// 경로: /a/b, //a, //a[@x='y'], //a[2], //a[last()], //a/@href, //a/text(), //*[contains(@class,'x')], ., .., |(합집합)
function splitTop(s, sep) {
  const out = [];
  let depth = 0;
  let q = null;
  let start = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (q) { if (ch === q) q = null; continue; }
    if (ch === '"' || ch === "'") { q = ch; continue; }
    if (ch === '[' || ch === '(') depth++;
    else if (ch === ']' || ch === ')') depth--;
    else if (depth === 0 && s.startsWith(sep, i)) { out.push(s.slice(start, i)); start = i + sep.length; i += sep.length - 1; }
  }
  out.push(s.slice(start));
  return out;
}
function parseSteps(path) {
  const steps = [];
  let i = 0;
  const s = path.trim();
  let axis = 'child';
  if (s.startsWith('//')) { axis = 'desc'; i = 2; } else if (s.startsWith('/')) { i = 1; } else axis = 'rel';
  let buf = '';
  let depth = 0;
  let q = null;
  const push = () => {
    const t = buf.trim();
    if (!t) return;
    const pm = /^([^[]*)((?:\[[\s\S]*\])*)$/.exec(t);
    const preds = [];
    let rest = pm ? pm[2] : '';
    while (rest) {
      let d = 0;
      let k = 0;
      let qq = null;
      for (; k < rest.length; k++) {
        const ch = rest[k];
        if (qq) { if (ch === qq) qq = null; continue; }
        if (ch === '"' || ch === "'") qq = ch;
        else if (ch === '[') d++;
        else if (ch === ']') { d--; if (d === 0) break; }
      }
      preds.push(rest.slice(1, k));
      rest = rest.slice(k + 1);
    }
    steps.push({ axis: axis === 'rel' ? 'child' : axis, test: (pm ? pm[1] : t).trim(), preds });
    axis = 'child';
  };
  for (; i < s.length; i++) {
    const ch = s[i];
    if (q) { buf += ch; if (ch === q) q = null; continue; }
    if (ch === '"' || ch === "'") { q = ch; buf += ch; continue; }
    if (ch === '[') depth++;
    if (ch === ']') depth--;
    if (depth === 0 && ch === '/') {
      push();
      buf = '';
      if (s[i + 1] === '/') { axis = 'desc'; i++; } else axis = 'child';
      continue;
    }
    buf += ch;
  }
  push();
  return steps;
}
function matchTest(n, test) {
  if (!isNode(n)) return false;
  if (test === '*' || test === 'node()') return true;
  return n.tag === test || n.tag.toLowerCase() === test.toLowerCase() || localName(n.tag) === test.toLowerCase();
}
function literal(s) {
  const t = s.trim();
  if (/^(['"]).*\1$/s.test(t)) return t.slice(1, -1);
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return null;
}
/** 조건식 값 (문자열 · 수 · 불) */
function predValue(expr, n, pos, size) {
  const e = expr.trim();
  const lit = literal(e);
  if (lit !== null) return lit;
  if (e === 'last()') return size;
  if (e === 'position()') return pos;
  if (e === 'text()' || e === '.') return textOf(n);
  if (e.startsWith('@')) return n.attrs?.[e.slice(1)] ?? n.attrs?.[e.slice(1).toLowerCase()] ?? null;
  let f = /^(contains|starts-with|ends-with)\((.*)\)$/s.exec(e);
  if (f) {
    const [a, b] = splitTop(f[2], ',').map((x) => String(predValue(x, n, pos, size) ?? ''));
    return f[1] === 'contains' ? a.includes(b) : f[1] === 'starts-with' ? a.startsWith(b) : a.endsWith(b);
  }
  f = /^normalize-space\((.*)\)$/s.exec(e);
  if (f) return String(predValue(f[1] || '.', n, pos, size) ?? '').replace(/\s+/g, ' ').trim();
  f = /^not\((.*)\)$/s.exec(e);
  if (f) return !truthy(predValue(f[1], n, pos, size));
  f = /^count\((.*)\)$/s.exec(e);
  if (f) return evalPath(f[1], [n]).length;
  // 자식 요소 이름 (예: [title='x'])
  const kids = evalPath(e, [n]);
  if (kids.length) return typeof kids[0] === 'string' ? kids[0] : textOf(kids[0]);
  return null;
}
const truthy = (v) => (typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : v !== null && v !== '');
function predTest(expr, n, pos, size) {
  const ors = splitTop(expr, ' or ');
  if (ors.length > 1) return ors.some((x) => predTest(x, n, pos, size));
  const ands = splitTop(expr, ' and ');
  if (ands.length > 1) return ands.every((x) => predTest(x, n, pos, size));
  for (const op of ['!=', '<=', '>=', '=', '<', '>']) {
    const parts = splitTop(expr, op);
    if (parts.length === 2) {
      const a = predValue(parts[0], n, pos, size);
      const b = predValue(parts[1], n, pos, size);
      const na = Number(a);
      const nb = Number(b);
      const num = typeof a === 'number' || typeof b === 'number';
      const x = num ? na : String(a ?? '');
      const y = num ? nb : String(b ?? '');
      switch (op) {
        case '=': return x === y;
        case '!=': return x !== y;
        case '<': return x < y;
        case '>': return x > y;
        case '<=': return x <= y;
        default: return x >= y;
      }
    }
  }
  const v = predValue(expr, n, pos, size);
  if (typeof v === 'number') return pos === v;
  return truthy(v);
}
/** XPath → 노드 · 문자열 목록 */
export function evalPath(path, ctxNodes) {
  const alts = splitTop(path.trim(), '|');
  if (alts.length > 1) return alts.flatMap((p) => evalPath(p, ctxNodes));
  let set = ctxNodes;
  const absolute = path.trim().startsWith('/');
  if (absolute) {
    let r = ctxNodes[0];
    while (r?.parent) r = r.parent;
    set = r ? [r] : [];
  }
  for (const st of parseSteps(path)) {
    const out = [];
    for (const n of set) {
      if (!isNode(n)) continue;
      let cands;
      const t = st.test;
      if (t === '.') cands = [n];
      else if (t === '..') cands = n.parent ? [n.parent] : [];
      else if (t.startsWith('@')) {
        const nodes = st.axis === 'desc' ? [n, ...descendants(n)] : [n];
        const name = t.slice(1);
        cands = [];
        for (const x of nodes) {
          if (name === '*') cands.push(...Object.values(x.attrs));
          else if (x.attrs[name] !== undefined || x.attrs[name.toLowerCase()] !== undefined) cands.push(x.attrs[name] ?? x.attrs[name.toLowerCase()]);
        }
      } else if (t === 'text()') {
        const nodes = st.axis === 'desc' ? [n, ...descendants(n)] : [n];
        cands = nodes.flatMap((x) => x.children.filter((c) => typeof c === 'string' && c.trim()).map((c) => c.replace(/\s+/g, ' ').trim()));
      } else {
        const pool = st.axis === 'desc' ? [...descendants(n)] : n.children.filter(isNode);
        cands = pool.filter((x) => matchTest(x, t));
      }
      for (const p of st.preds) {
        const size = cands.length;
        cands = cands.filter((x, i) => (isNode(x) ? predTest(p, x, i + 1, size) : truthy(predValue(p, { attrs: {}, children: [x] }, i + 1, size))));
      }
      for (const c of cands) if (!out.includes(c)) out.push(c);
    }
    set = out;
  }
  return set;
}
const nodeText = (x) => (typeof x === 'string' ? x : textOf(x));

// ───────────── 표 · 목록 · 피드 · CSV ─────────────
export function htmlTables(root) { return [...descendants(root)].filter((n) => n.tag === 'table'); }
export function htmlLists(root) { return [...descendants(root)].filter((n) => n.tag === 'ul' || n.tag === 'ol'); }
export function tableRows(table) {
  const rows = [];
  const walk = (n) => {
    for (const ch of n.children) {
      if (!isNode(ch)) continue;
      if (ch.tag === 'table') continue; // 안쪽 표는 건너뜀
      if (ch.tag === 'tr') {
        const row = [];
        for (const td of ch.children) {
          if (!isNode(td) || (td.tag !== 'td' && td.tag !== 'th')) continue;
          row.push(autoValue(textOf(td)));
          const span = Math.min(50, Number(td.attrs.colspan) || 1);
          for (let k = 1; k < span; k++) row.push('');
        }
        rows.push(row);
      } else walk(ch);
    }
  };
  walk(table);
  return rows;
}
export function parseCsv(text, delim = null) {
  const src = String(text).replace(/^\ufeff/, '');
  const d = delim ?? (src.split('\n', 1)[0].includes('\t') && !src.split('\n', 1)[0].includes(',') ? '\t' : ',');
  const rows = [];
  let row = [];
  let f = '';
  let q = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (q) {
      if (ch === '"') { if (src[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch;
      continue;
    }
    if (ch === '"' && f === '') { q = true; continue; }
    if (ch === d) { row.push(f); f = ''; continue; }
    if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      row.push(f); f = ''; rows.push(row); row = [];
      continue;
    }
    f += ch;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}
/** RSS · Atom 피드 → { feed: {title, description, url}, items: [{title, url, date, summary, author}] } */
export function parseFeed(text) {
  const root = parseMarkup(text, { xml: true });
  const first = (n, names) => {
    for (const nm of names) {
      const hit = n.children.find((c) => isNode(c) && localName(c.tag) === nm);
      if (hit) return hit;
    }
    return null;
  };
  const txt = (n, names) => { const x = first(n, names); return x ? textOf(x) : ''; };
  const link = (n) => {
    const ls = n.children.filter((c) => isNode(c) && localName(c.tag) === 'link');
    const alt = ls.find((l) => !l.attrs.rel || l.attrs.rel === 'alternate') ?? ls[0];
    return alt ? alt.attrs.href ?? textOf(alt) : '';
  };
  const all = [...descendants(root)];
  const channel = all.find((n) => localName(n.tag) === 'channel') ?? all.find((n) => localName(n.tag) === 'feed') ?? root;
  const items = all.filter((n) => localName(n.tag) === 'item' || localName(n.tag) === 'entry').map((it) => ({
    title: txt(it, ['title']), url: link(it), date: txt(it, ['pubdate', 'published', 'updated', 'date']),
    summary: txt(it, ['description', 'summary', 'content']).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(), author: txt(it, ['author', 'creator']),
  }));
  return { feed: { title: txt(channel, ['title']), description: txt(channel, ['description', 'subtitle']), url: link(channel), author: txt(channel, ['author', 'managingeditor']) }, items };
}

// ───────────── 금융 (GOOGLEFINANCE · STOCKHISTORY) ─────────────
/** 티커 → { market: 'krx'|'us'|'fx', code } (KRX:005930, KOSDAQ:035720, 005930, NASDAQ:GOOG, AAPL, CURRENCY:USDKRW) */
export function parseTicker(t) {
  const s = String(t).trim().toUpperCase();
  let m = /^(?:KRX|KOSPI|KOSDAQ|KRX:KOSDAQ)[:.]?(\d{6})$/.exec(s) ?? /^(\d{6})(?:\.K[SQ])?$/.exec(s);
  if (m) return { market: 'krx', code: m[1] };
  m = /^CURRENCY:([A-Z]{3})([A-Z]{3})$/.exec(s);
  if (m) return { market: 'fx', code: `${m[1]}${m[2]}` };
  m = /^(?:NASDAQ|NYSE|NYSEARCA|NYSEAMERICAN|AMEX|BATS)?:?([A-Z][A-Z0-9.-]*)$/.exec(s);
  if (m) return { market: 'us', code: m[1] };
  return null;
}
const msToSerial = (ms) => ms / 86400000 + 25569;
const ymd = (serial) => { const d = serialToDate(serial); return `${d.y}${String(d.m).padStart(2, '0')}${String(d.d).padStart(2, '0')}`; };
export function quoteUrl(tk) {
  if (tk.market === 'krx') return `https://polling.finance.naver.com/api/realtime/domestic/stock/${tk.code}`;
  return `https://stooq.com/q/l/?s=${encodeURIComponent(tk.market === 'fx' ? tk.code.toLowerCase() : `${tk.code.toLowerCase()}.us`)}&f=sd2t2ohlcvn&h&e=csv`;
}
export function historyUrl(tk, from, to, interval = 'd') {
  if (tk.market === 'krx') return `https://api.finance.naver.com/siseJson.naver?symbol=${tk.code}&requestType=1&startTime=${ymd(from)}&endTime=${ymd(to)}&timeframe=${{ d: 'day', w: 'week', m: 'month' }[interval]}`;
  return `https://stooq.com/q/d/l/?s=${encodeURIComponent(tk.market === 'fx' ? tk.code.toLowerCase() : `${tk.code.toLowerCase()}.us`)}&d1=${ymd(from)}&d2=${ymd(to)}&i=${interval}`;
}
const num = (s) => { const n = Number(String(s ?? '').replace(/,/g, '')); return Number.isFinite(n) ? n : null; };
/** 시세 응답 → { price, open, high, low, volume, closeyest, change, changepct, name, currency, tradetime } */
export function parseQuote(tk, text) {
  if (tk.market === 'krx') {
    const j = JSON.parse(text);
    const d = j.datas?.[0] ?? j.result?.areas?.[0]?.datas?.[0];
    if (!d) return null;
    const price = num(d.closePrice ?? d.nv);
    const change = num(d.compareToPreviousClosePrice ?? d.cv);
    const sign = d.compareToPreviousPrice?.name === 'FALLING' || d.compareToPreviousPrice?.code === '5' ? -1 : 1;
    const ch = change === null ? null : Math.abs(change) * (String(d.compareToPreviousClosePrice ?? '').startsWith('-') ? -1 : sign);
    return {
      price, open: num(d.openPrice ?? d.ov), high: num(d.highPrice ?? d.hv), low: num(d.lowPrice ?? d.lv), volume: num(d.accumulatedTradingVolume ?? d.aq),
      change: ch, changepct: num(d.fluctuationsRatio ?? d.cr) === null ? null : num(d.fluctuationsRatio ?? d.cr) / 100, closeyest: price !== null && ch !== null ? price - ch : null,
      name: d.stockName ?? d.nm ?? tk.code, currency: 'KRW', marketcap: num(d.marketValue),
    };
  }
  const rows = parseCsv(text);
  const h = rows[0]?.map((x) => x.toLowerCase()) ?? [];
  const r = rows[1];
  if (!r || /N\/D/.test(r.join(''))) return null;
  const g = (k) => r[h.indexOf(k)];
  return { price: num(g('close')), open: num(g('open')), high: num(g('high')), low: num(g('low')), volume: num(g('volume')), name: g('name') ?? tk.code, currency: tk.market === 'fx' ? tk.code.slice(3) : 'USD', tradetime: g('date') };
}
/** 과거 시세 응답 → [[date serial, open, high, low, close, volume]] */
export function parseHistory(tk, text) {
  if (tk.market === 'krx') {
    // [['날짜','시가',…], ["20240102", 78200, …], …] (작은따옴표 · 줄바꿈이 섞인 유사 JSON)
    const rows = [];
    for (const m of String(text).matchAll(/\[\s*"(\d{8})"\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)/g)) {
      const s = m[1];
      rows.push([dateToSerial(+s.slice(0, 4), +s.slice(4, 6), +s.slice(6, 8)), +m[2], +m[3], +m[4], +m[5], +m[6]]);
    }
    return rows;
  }
  const rows = parseCsv(text);
  const h = rows[0]?.map((x) => x.toLowerCase()) ?? [];
  const at = (k) => h.indexOf(k);
  return rows.slice(1).filter((r) => /^\d{4}-\d\d-\d\d$/.test(r[at('date')] ?? '')).map((r) => {
    const [y, mo, d] = r[at('date')].split('-').map(Number);
    return [dateToSerial(y, mo, d), num(r[at('open')]), num(r[at('high')]), num(r[at('low')]), num(r[at('close')]), num(r[at('volume')]) ?? 0];
  });
}

// ───────────── QUERY (Google Visualization Query Language) ─────────────
const Q_KEYWORDS = ['select', 'where', 'group by', 'pivot', 'order by', 'limit', 'offset', 'label', 'format', 'options'];
function qTokens(s) {
  const out = [];
  const re = /\s*(?:(`[^`]*`)|('[^']*'|"[^"]*")|(\d+\.?\d*(?:e[+-]?\d+)?)|(<=|>=|!=|<>|[=<>(),*+\-/])|([A-Za-z_][\w.]*))/giy;
  let m;
  let pos = 0;
  while (pos < s.length) {
    re.lastIndex = pos;
    m = re.exec(s);
    if (!m || m[0].trim() === '' && re.lastIndex === pos) break;
    pos = re.lastIndex;
    if (m[1]) out.push({ t: 'id', v: m[1].slice(1, -1) });
    else if (m[2]) out.push({ t: 'str', v: m[2].slice(1, -1) });
    else if (m[3]) out.push({ t: 'num', v: Number(m[3]) });
    else if (m[4]) out.push({ t: 'op', v: m[4] === '<>' ? '!=' : m[4] });
    else if (m[5]) out.push({ t: 'id', v: m[5], w: m[5].toLowerCase() });
    if (/^\s*$/.test(s.slice(pos))) break;
  }
  return out;
}
const AGG = new Set(['sum', 'avg', 'count', 'max', 'min']);
const SCALAR = new Set(['year', 'month', 'day', 'quarter', 'hour', 'minute', 'second', 'dayofweek', 'upper', 'lower', 'todate', 'now', 'datediff', 'tostring']);
class QParser {
  constructor(toks) { this.k = toks; this.i = 0; }
  peek(o = 0) { return this.k[this.i + o]; }
  kw(word) {
    const parts = word.split(' ');
    for (let j = 0; j < parts.length; j++) if (this.peek(j)?.w !== parts[j]) return false;
    this.i += parts.length;
    return true;
  }
  atClause() { return Q_KEYWORDS.some((w) => { const p = w.split(' '); return p.every((x, j) => this.peek(j)?.w === x); }); }
  op(v) { if (this.peek()?.t === 'op' && this.peek().v === v) { this.i++; return true; } return false; }
  expect(v) { if (!this.op(v)) throw new Error(`QUERY: '${v}' 가 필요합니다`); }
  // 식: 산술 + 비교는 where 에서만
  primary() {
    const t = this.k[this.i++];
    if (!t) throw new Error('QUERY: 식이 끝났습니다');
    if (t.t === 'num') return { k: 'lit', v: t.v };
    if (t.t === 'str') return { k: 'lit', v: t.v };
    if (t.t === 'op' && t.v === '(') { const e = this.orExpr(); this.expect(')'); return e; }
    if (t.t === 'op' && t.v === '-') return { k: 'neg', a: this.primary() };
    if (t.t === 'id') {
      const w = t.w ?? t.v.toLowerCase();
      if (w === 'true' || w === 'false') return { k: 'lit', v: w === 'true' };
      if (w === 'null') return { k: 'lit', v: null };
      if ((w === 'date' || w === 'datetime' || w === 'timeofday') && this.peek()?.t === 'str') {
        const s = this.k[this.i++].v;
        const mm = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s.trim());
        if (w === 'timeofday') { const p = s.split(':').map(Number); return { k: 'lit', v: ((p[0] || 0) * 3600 + (p[1] || 0) * 60 + (p[2] || 0)) / 86400 }; }
        if (!mm) throw new Error(`QUERY: 날짜 '${s}' 를 읽을 수 없습니다`);
        return { k: 'lit', v: dateToSerial(+mm[1], +mm[2], +mm[3]) + ((+(mm[4] ?? 0)) * 3600 + (+(mm[5] ?? 0)) * 60 + (+(mm[6] ?? 0))) / 86400 };
      }
      if (this.peek()?.t === 'op' && this.peek().v === '(' && (AGG.has(w) || SCALAR.has(w))) {
        this.i++;
        const args = [];
        if (!this.op(')')) {
          do { args.push(this.op('*') ? { k: 'star' } : this.orExpr()); } while (this.op(','));
          this.expect(')');
        }
        return AGG.has(w) ? { k: 'agg', f: w, a: args[0] } : { k: 'fn', f: w, args };
      }
      return { k: 'col', v: t.v };
    }
    throw new Error(`QUERY: '${t.v}' 를 읽을 수 없습니다`);
  }
  mul() { let a = this.primary(); for (;;) { if (this.op('*')) a = { k: 'bin', o: '*', a, b: this.primary() }; else if (this.op('/')) a = { k: 'bin', o: '/', a, b: this.primary() }; else return a; } }
  add() { let a = this.mul(); for (;;) { if (this.op('+')) a = { k: 'bin', o: '+', a, b: this.mul() }; else if (this.op('-')) a = { k: 'bin', o: '-', a, b: this.mul() }; else return a; } }
  cmp() {
    const a = this.add();
    const t = this.peek();
    if (t?.t === 'op' && ['=', '!=', '<', '>', '<=', '>='].includes(t.v)) { this.i++; return { k: 'cmp', o: t.v, a, b: this.add() }; }
    if (t?.w === 'is') {
      this.i++;
      const not = this.kw('not');
      if (!this.kw('null')) throw new Error('QUERY: is null / is not null');
      return { k: 'isnull', not, a };
    }
    let not = false;
    if (t?.w === 'not' && ['contains', 'starts', 'ends', 'matches', 'like'].includes(this.peek(1)?.w)) { this.i++; not = true; }
    for (const w of ['contains', 'starts with', 'ends with', 'matches', 'like']) {
      if (this.kw(w)) { const e = { k: 'str', o: w, a, b: this.add() }; return not ? { k: 'not', a: e } : e; }
    }
    return a;
  }
  notExpr() { if (this.peek()?.w === 'not') { this.i++; return { k: 'not', a: this.notExpr() }; } return this.cmp(); }
  andExpr() { let a = this.notExpr(); while (this.kw('and')) a = { k: 'and', a, b: this.notExpr() }; return a; }
  orExpr() { let a = this.andExpr(); while (this.kw('or')) a = { k: 'or', a, b: this.andExpr() }; return a; }
  list() { const out = []; do out.push(this.orExpr()); while (this.op(',')); return out; }
}
export function parseQuery(text) {
  const p = new QParser(qTokens(String(text ?? '')));
  const q = { select: null, where: null, group: [], pivot: [], order: [], limit: null, offset: 0, labels: [] };
  while (p.peek()) {
    if (p.kw('select')) { q.select = p.op('*') ? null : p.list(); continue; }
    if (p.kw('where')) { q.where = p.orExpr(); continue; }
    if (p.kw('group by')) { q.group = p.list(); continue; }
    if (p.kw('pivot')) { q.pivot = p.list(); continue; }
    if (p.kw('order by')) {
      do {
        const e = p.orExpr();
        const desc = p.kw('desc');
        if (!desc) p.kw('asc');
        q.order.push({ e, desc });
      } while (p.op(','));
      continue;
    }
    if (p.kw('limit')) { q.limit = p.k[p.i++]?.v; continue; }
    if (p.kw('offset')) { q.offset = p.k[p.i++]?.v ?? 0; continue; }
    if (p.kw('label')) {
      do { const e = p.orExpr(); const t = p.k[p.i++]; q.labels.push({ e, text: t?.v ?? '' }); } while (p.op(','));
      continue;
    }
    if (p.kw('format') || p.kw('options')) { while (p.peek() && !p.atClause()) p.i++; continue; }
    throw new Error(`QUERY: '${p.peek().v}' 근처를 해석할 수 없습니다`);
  }
  return q;
}
const exprKey = (e) => JSON.stringify(e);
const hasAgg = (e) => !!e && typeof e === 'object' && (e.k === 'agg' || Object.values(e).some((x) => (Array.isArray(x) ? x.some(hasAgg) : x && typeof x === 'object' && hasAgg(x))));
function exprLabel(e, heads) {
  switch (e.k) {
    case 'col': return heads(e.v);
    case 'agg': return `${e.f} ${e.a?.k === 'star' ? '' : exprLabel(e.a, heads)}`.trim();
    case 'fn': return `${e.f}(${e.args.map((a) => exprLabel(a, heads)).join(', ')})`;
    case 'lit': return String(e.v);
    case 'bin': return `${exprLabel(e.a, heads)}${e.o}${exprLabel(e.b, heads)}`;
    default: return '';
  }
}
const isNum = (v) => typeof v === 'number';
function scalarFn(f, args) {
  const a = args[0];
  const d = isNum(a) ? serialToDate(a) : null;
  switch (f) {
    case 'year': return d ? d.y : null;
    case 'month': return d ? d.m - 1 : null; // 스프레드시트처럼 0 = 1월
    case 'day': return d ? d.d : null;
    case 'quarter': return d ? Math.floor((d.m - 1) / 3) + 1 : null;
    case 'dayofweek': return d ? d.dow + 1 : null;
    case 'hour': return isNum(a) ? Math.floor(((a % 1) * 24) + 1e-9) : null;
    case 'minute': return isNum(a) ? Math.floor(((a * 1440) % 60) + 1e-9) : null;
    case 'second': return isNum(a) ? Math.round((a * 86400) % 60) : null;
    case 'upper': return a == null ? null : String(a).toUpperCase();
    case 'lower': return a == null ? null : String(a).toLowerCase();
    case 'todate': return isNum(a) ? Math.floor(a) : null;
    case 'datediff': return isNum(a) && isNum(args[1]) ? Math.floor(a) - Math.floor(args[1]) : null;
    case 'tostring': return a == null ? null : String(a);
    case 'now': return msToSerial(Date.now());
    default: return null;
  }
}
function aggregate(f, vals) {
  const xs = vals.filter((v) => v !== null && v !== undefined && v !== '');
  if (f === 'count') return xs.length;
  const ns = xs.filter(isNum);
  if (f === 'sum') return ns.reduce((a, b) => a + b, 0);
  if (f === 'avg') return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : null;
  const pool = ns.length ? ns : xs;
  if (!pool.length) return null;
  return pool.reduce((m, v) => (compareValues(f === 'max' ? v : m, f === 'max' ? m : v) > 0 ? v : m));
}
/**
 * QUERY 실행: rows = 데이터 행(머리글 제외), cols = 열 식별자 (A,B… 또는 Col1…), heads = 머리글 글자
 */
export function runQuery(q, rows, cols, heads) {
  const idx = new Map(cols.map((c, i) => [c.toUpperCase(), i]));
  const colIndex = (name) => {
    const i = idx.get(String(name).toUpperCase());
    if (i === undefined) throw new Error(`QUERY: 열 '${name}' 이(가) 없습니다`);
    return i;
  };
  const headOf = (name) => heads[colIndex(name)] || String(name);
  // 행 문맥의 식 값
  const val = (e, row, group) => {
    switch (e.k) {
      case 'lit': return e.v;
      case 'col': return row[colIndex(e.v)] ?? null;
      case 'neg': { const v = val(e.a, row, group); return isNum(v) ? -v : null; }
      case 'bin': {
        const a = val(e.a, row, group);
        const b = val(e.b, row, group);
        if (!isNum(a) || !isNum(b)) return null;
        return e.o === '+' ? a + b : e.o === '-' ? a - b : e.o === '*' ? a * b : b === 0 ? null : a / b;
      }
      case 'fn': return scalarFn(e.f, e.args.map((x) => val(x, row, group)));
      case 'agg': {
        if (!group) throw new Error('QUERY: 집계 함수는 select · order by · label 에만 쓸 수 있습니다');
        return aggregate(e.f, group.map((r) => (e.a?.k === 'star' ? 1 : val(e.a, r, null))));
      }
      case 'cmp': {
        const a = val(e.a, row, group);
        const b = val(e.b, row, group);
        if (a === null || b === null || a === '' || b === '') return e.o === '!=' ? a !== b : e.o === '=' ? a === b : false;
        const c = compareValues(a, b);
        return { '=': c === 0, '!=': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[e.o];
      }
      case 'str': {
        const a = val(e.a, row, group);
        const b = String(val(e.b, row, group) ?? '');
        if (a === null || a === undefined) return false;
        const s = String(a);
        if (e.o === 'contains') return s.includes(b);
        if (e.o === 'starts with') return s.startsWith(b);
        if (e.o === 'ends with') return s.endsWith(b);
        if (e.o === 'matches') { try { return new RegExp(`^(?:${b})$`).test(s); } catch { return false; } }
        const re = new RegExp(`^${b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/%/g, '.*').replace(/_/g, '.')}$`, 's');
        return re.test(s);
      }
      case 'isnull': { const v = val(e.a, row, group); const n = v === null || v === undefined || v === ''; return e.not ? !n : n; }
      case 'not': return !val(e.a, row, group);
      case 'and': return !!val(e.a, row, group) && !!val(e.b, row, group);
      case 'or': return !!val(e.a, row, group) || !!val(e.b, row, group);
      default: return null;
    }
  };
  let data = q.where ? rows.filter((r) => val(q.where, r, null) === true) : rows.slice();
  const select = q.select ?? cols.map((c) => ({ k: 'col', v: c }));
  const aggregated = q.group.length || q.pivot.length || select.some(hasAgg);
  const labelOf = (e) => q.labels.find((l) => exprKey(l.e) === exprKey(e))?.text ?? exprLabel(e, headOf);
  let head = select.map(labelOf);
  let out;
  if (!aggregated) {
    const rowsWithSrc = data.map((r) => ({ src: r, out: select.map((e) => val(e, r, null)) }));
    if (q.order.length) {
      rowsWithSrc.sort((x, y) => {
        for (const o of q.order) {
          const c = compareNulls(val(o.e, x.src, null), val(o.e, y.src, null));
          if (c) return o.desc ? -c : c;
        }
        return 0;
      });
    }
    out = rowsWithSrc.map((x) => x.out);
  } else {
    // 묶기 (+ pivot: 집계 열을 피벗 값마다 펼침)
    const groups = new Map();
    for (const r of data) {
      const key = JSON.stringify(q.group.map((g) => val(g, r, null)));
      let g = groups.get(key);
      if (!g) groups.set(key, g = { vals: q.group.map((e) => val(e, r, null)), rows: [] });
      g.rows.push(r);
    }
    if (!q.group.length && !groups.size) groups.set('[]', { vals: [], rows: [] });
    let entries = [...groups.values()];
    if (q.order.length) {
      entries.sort((x, y) => {
        for (const o of q.order) {
          const c = compareNulls(val(o.e, x.rows[0] ?? [], x.rows), val(o.e, y.rows[0] ?? [], y.rows));
          if (c) return o.desc ? -c : c;
        }
        return 0;
      });
    } else if (q.group.length) {
      entries.sort((x, y) => { for (let i = 0; i < x.vals.length; i++) { const c = compareNulls(x.vals[i], y.vals[i]); if (c) return c; } return 0; });
    }
    if (q.pivot.length) {
      const pkeys = [...new Set(data.map((r) => JSON.stringify(q.pivot.map((p) => val(p, r, null)))))].sort((a, b) => compareNulls(JSON.parse(a)[0], JSON.parse(b)[0]));
      const plain = select.filter((e) => !hasAgg(e));
      const aggs = select.filter(hasAgg);
      head = [...plain.map(labelOf), ...pkeys.flatMap((pk) => aggs.map((e) => `${JSON.parse(pk).join(' ')}${aggs.length > 1 ? ` ${labelOf(e)}` : ''}`))];
      out = entries.map((g) => [
        ...plain.map((e) => val(e, g.rows[0] ?? [], g.rows)),
        ...pkeys.flatMap((pk) => {
          const sub = g.rows.filter((r) => JSON.stringify(q.pivot.map((p) => val(p, r, null))) === pk);
          return aggs.map((e) => (sub.length ? val(e, sub[0], sub) : null));
        }),
      ]);
    } else {
      out = entries.map((g) => select.map((e) => val(e, g.rows[0] ?? [], g.rows)));
    }
  }
  out = out.slice(q.offset ?? 0, q.limit != null ? (q.offset ?? 0) + q.limit : undefined);
  return { head, rows: out };
}
function compareNulls(a, b) {
  const na = a === null || a === undefined || a === '';
  const nb = b === null || b === undefined || b === '';
  if (na || nb) return na === nb ? 0 : na ? -1 : 1;
  return compareValues(a, b);
}

// ───────────── SPARKLINE ─────────────
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
export function sparklineSvg(values, o = {}) {
  const type = String(o.charttype ?? 'line').toLowerCase();
  const W = 100;
  const H = 30;
  const nums = values.map((v) => (isNum(v) ? v : null));
  const color = o.color ?? '#1a73e8';
  let body = '';
  if (type === 'bar') {
    // 가로 막대 1~2개 (값 / 최대값)
    const vs = nums.filter(isNum);
    const max = isNum(Number(o.max)) && o.max !== undefined ? Number(o.max) : vs.reduce((a, b) => a + Math.abs(b), 0) || 1;
    let x = 0;
    const cols = [o.color1 ?? color, o.color2 ?? '#a1c2fa'];
    vs.slice(0, 2).forEach((v, i) => {
      const w = Math.max(0, Math.min(W - x, (Math.abs(v) / max) * W));
      body += `<rect x="${x.toFixed(2)}" y="0" width="${w.toFixed(2)}" height="${H}" fill="${esc(cols[i])}"/>`;
      x += w;
    });
  } else {
    const vs = nums.map((v) => v ?? (o.empty === 'zero' ? 0 : null));
    const real = vs.filter(isNum);
    if (!real.length) return null;
    let lo = o.ymin !== undefined ? Number(o.ymin) : Math.min(...real);
    let hi = o.ymax !== undefined ? Number(o.ymax) : Math.max(...real);
    if (type === 'column' || type === 'winloss') { lo = Math.min(0, lo); hi = Math.max(0, hi); }
    if (hi === lo) { hi += 1; lo -= 1; }
    const n = vs.length;
    const y = (v) => H - ((v - lo) / (hi - lo)) * H;
    if (type === 'column' || type === 'winloss') {
      const bw = W / n;
      vs.forEach((v, i) => {
        if (!isNum(v)) return;
        const neg = v < 0;
        let fill = neg ? (o.negcolor ?? color) : color;
        if (i === 0 && o.firstcolor) fill = o.firstcolor;
        if (i === n - 1 && o.lastcolor) fill = o.lastcolor;
        if (o.highcolor && v === Math.max(...real)) fill = o.highcolor;
        if (o.lowcolor && v === Math.min(...real)) fill = o.lowcolor;
        const x0 = (i * bw + bw * 0.1).toFixed(2);
        if (type === 'winloss') {
          body += `<rect x="${x0}" y="${neg ? H / 2 : 0}" width="${(bw * 0.8).toFixed(2)}" height="${H / 2 - 1}" fill="${esc(fill)}"/>`;
        } else {
          const y0 = y(0);
          const y1 = y(v);
          body += `<rect x="${x0}" y="${Math.min(y0, y1).toFixed(2)}" width="${(bw * 0.8).toFixed(2)}" height="${Math.max(0.5, Math.abs(y1 - y0)).toFixed(2)}" fill="${esc(fill)}"/>`;
        }
      });
      if (o.axis) body += `<line x1="0" x2="${W}" y1="${y(0)}" y2="${y(0)}" stroke="${esc(o.axiscolor ?? '#999')}" stroke-width="0.6"/>`;
    } else {
      const step = n > 1 ? W / (n - 1) : W;
      let d = '';
      let pen = false;
      vs.forEach((v, i) => {
        if (!isNum(v)) { pen = false; return; }
        d += `${pen ? 'L' : 'M'}${(i * step).toFixed(2)},${y(v).toFixed(2)}`;
        pen = true;
      });
      body = `<path d="${d}" fill="none" stroke="${esc(color)}" stroke-width="${Number(o.linewidth) || 1.5}" stroke-linejoin="round" vector-effect="non-scaling-stroke"/>`;
      if (o.axis) body += `<line x1="0" x2="${W}" y1="${y(0)}" y2="${y(0)}" stroke="${esc(o.axiscolor ?? '#999')}" stroke-width="0.6"/>`;
    }
  }
  const flip = o.rtl ? ' transform="scale(-1,1) translate(-100,0)"' : '';
  return `data:image/svg+xml;utf8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 -2 ${W} ${H + 4}" preserveAspectRatio="none"><g${flip}>${body}</g></svg>`)}`;
}

// ───────────── 번역 ─────────────
export const translateUrl = (text, sl, tl) => `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${encodeURIComponent(sl || 'auto')}&tl=${encodeURIComponent(tl || 'ko')}&dt=t&q=${encodeURIComponent(text)}`;
export function parseTranslate(text) {
  const j = JSON.parse(text);
  return { text: (j[0] ?? []).map((x) => x?.[0] ?? '').join(''), lang: j[2] ?? '' };
}

// ───────────── 태국어 BAHTTEXT ─────────────
export function bahtText(n) {
  const digits = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const units = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
  const say = (x) => {
    if (x === 0) return '';
    if (x >= 1e6) return say(Math.floor(x / 1e6)) + 'ล้าน' + say(x % 1e6);
    const s = String(x);
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const d = +s[i];
      const pos = s.length - 1 - i;
      if (!d) continue;
      if (pos === 1 && d === 1) out += 'สิบ';
      else if (pos === 1 && d === 2) out += 'ยี่สิบ';
      else if (pos === 0 && d === 1 && s.length > 1) out += 'เอ็ด';
      else out += digits[d] + units[pos];
    }
    return out;
  };
  const neg = n < 0;
  const v = Math.round(Math.abs(n) * 100);
  const baht = Math.floor(v / 100);
  const satang = v % 100;
  let t = baht ? `${say(baht)}บาท` : '';
  t += satang ? `${say(satang)}สตางค์` : baht ? 'ถ้วน' : 'ศูนย์บาทถ้วน';
  return (neg ? 'ลบ' : '') + t;
}

// ───────────── 수 · 날짜 도우미 ─────────────
function normInv(p) {
  // Acklam 근사
  const a = [-39.6968302866538, 220.946098424521, -275.928510446969, 138.357751867269, -30.6647980661472, 2.50662827745924];
  const b = [-54.4760987982241, 161.585836858041, -155.698979859887, 66.8013118877197, -13.2806815528857];
  const c = [-0.00778489400243029, -0.322396458041136, -2.40075827716184, -2.54973253934373, 4.37466414146497, 2.93816398269878];
  const d = [0.00778469570904146, 0.32246712907004, 2.445134137143, 3.75440866190742];
  const pl = 0.02425;
  if (p < pl) { const q = Math.sqrt(-2 * Math.log(p)); return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  if (p > 1 - pl) { const q = Math.sqrt(-2 * Math.log(1 - p)); return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  const q = p - 0.5;
  const r = q * q;
  return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}
function yearFrac(s, e, basis) {
  const d1 = serialToDate(Math.min(s, e));
  const d2 = serialToDate(Math.max(s, e));
  const y1 = d1.y; const m1 = d1.m; let dd1 = d1.d;
  const y2 = d2.y; const m2 = d2.m; let dd2 = d2.d;
  const days = Math.abs(e - s);
  switch (basis) {
    case 1: {
      const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
      let len = 0;
      for (let y = y1; y <= y2; y++) len += leap(y) ? 366 : 365;
      return days / (len / (y2 - y1 + 1));
    }
    case 2: return days / 360;
    case 3: return days / 365;
    case 4: dd1 = Math.min(dd1, 30); dd2 = Math.min(dd2, 30); return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
    default:
      if (dd1 === 31) dd1 = 30;
      if (dd2 === 31 && dd1 >= 30) dd2 = 30;
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
  }
}
const binomPmf = (n, p, k) => {
  let lc = 0;
  for (let i = 1; i <= k; i++) lc += Math.log((n - k + i) / i);
  return Math.exp(lc + k * Math.log(p) + (n - k) * Math.log(1 - p));
};
const flatVals = (args) => {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) for (const v of a.values()) out.push(v);
    else out.push(a);
  }
  return out;
};
const wideLen = (ch) => (ch.codePointAt(0) > 0xff ? 2 : 1);

// ───────────── 함수 ─────────────
const netRange = (url, ctx, parse) => {
  const t = netText(url, ctx);
  if (pending(t)) return t;
  try {
    const v = parse(t);
    return v ?? ERR.NA;
  } catch (e) {
    if (e instanceof FormulaError) return e;
    return ERR.NA;
  }
};
const fixUrl = (u) => {
  const s = String(u ?? '').trim();
  if (!s) throw ERR.VALUE;
  return /^[a-z][\w+.-]*:/i.test(s) ? s : `https://${s}`;
};
function optionsFrom(v) {
  const o = {};
  if (v === undefined || v === null) return o;
  const r = asRange(v);
  if (r.width >= 2) for (const row of r.rows) { if (row[0] != null && row[0] !== '') o[String(row[0]).toLowerCase()] = row[1]; }
  else if (r.height >= 2) for (const c of r.rows[0].keys()) o[String(r.rows[0][c]).toLowerCase()] = r.rows[1]?.[c];
  for (const k of Object.keys(o)) { const x = o[k]; if (typeof x === 'string' && /^(true|false)$/i.test(x)) o[k] = /^true$/i.test(x); }
  return o;
}

export const WEB = {
  // ── 웹에서 가져오기 (구글 스프레드시트) ──
  IMPORTDATA: (args, ctx) => {
    const url = fixUrl(str1(args[0]));
    const delim = args[1] != null ? str1(args[1]) : null;
    return netRange(url, ctx, (t) => {
      const rows = parseCsv(t, delim || (/\.tsv(\?|$)/i.test(url) ? '\t' : null)).map((r) => r.map(autoValue));
      return grid(rows.filter((r, i) => r.some((x) => x !== '') || i < rows.length - 1));
    });
  },
  IMPORTHTML: (args, ctx) => {
    const url = fixUrl(str1(args[0]));
    const kind = str1(args[1] ?? 'table').toLowerCase();
    const index = Math.max(1, Math.trunc(toNum(one(args[2]) ?? 1)));
    if (kind !== 'table' && kind !== 'list') return ERR.VALUE;
    return netRange(url, ctx, (t) => {
      const root = parseMarkup(t);
      if (kind === 'table') {
        const tb = htmlTables(root)[index - 1];
        return tb ? grid(tableRows(tb)) : ERR.NA;
      }
      const ls = htmlLists(root)[index - 1];
      if (!ls) return ERR.NA;
      return grid(ls.children.filter((c) => isNode(c) && c.tag === 'li').map((li) => [autoValue(textOf(li))]));
    });
  },
  IMPORTXML: (args, ctx) => {
    const url = fixUrl(str1(args[0]));
    const xpath = str1(args[1]);
    return netRange(url, ctx, (t) => {
      const xml = /^\s*<\?xml/i.test(t) && !/<html/i.test(t.slice(0, 500));
      const root = parseMarkup(t, { xml });
      const hits = evalPath(xpath, [root]);
      if (!hits.length) return ERR.NA;
      return grid(hits.map((h) => [autoValue(nodeText(h))]));
    });
  },
  IMPORTFEED: (args, ctx) => {
    const url = fixUrl(str1(args[0]));
    const query = args[1] != null ? str1(args[1]).toLowerCase().trim() : 'items';
    const headers = optBool(one(args[2]), false);
    const count = args[3] != null ? Math.max(1, toNum(one(args[3]))) : Infinity;
    return netRange(url, ctx, (t) => {
      const f = parseFeed(t);
      const [what, field] = query.split(/\s+/);
      if (what === 'feed') {
        const map = { title: f.feed.title, description: f.feed.description, url: f.feed.url, author: f.feed.author };
        if (field) return grid([[map[field] ?? '']]);
        return grid([...(headers ? [['Title', 'Description', 'URL']] : []), [f.feed.title, f.feed.description, f.feed.url]]);
      }
      const items = f.items.slice(0, count);
      const fields = field ? [field === 'created' ? 'date' : field] : ['title', 'url', 'date', 'summary'];
      const head = fields.map((x) => ({ title: 'Title', url: 'URL', date: 'Date Created', summary: 'Summary', author: 'Author' }[x] ?? x));
      return grid([...(headers ? [head] : []), ...items.map((it) => fields.map((x) => it[x] ?? ''))]);
    });
  },
  IMPORTRANGE: (args, ctx) => {
    const src = str1(args[0]).trim();
    const spec = str1(args[1]).trim();
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(spec);
    const sheet = m ? (m[1] ?? m[2]).replace(/''/g, "'") : '';
    const area = m ? m[3] : spec;
    const gid = /\/spreadsheets\/d\/([\w-]{20,})/.exec(src)?.[1] ?? (/^[\w-]{30,}$/.test(src) ? src : null);
    const url = gid
      ? `https://docs.google.com/spreadsheets/d/${gid}/gviz/tq?tqx=out:csv${sheet ? `&sheet=${encodeURIComponent(sheet)}` : ''}&range=${encodeURIComponent(area)}`
      : `wixel-doc:${src}\u0001${sheet}\u0001${area}`;
    return netRange(url, ctx, (t) => grid(parseCsv(t).map((r) => r.map(autoValue))));
  },
  GOOGLEFINANCE: (args, ctx) => {
    const tk = parseTicker(str1(args[0]));
    if (!tk) return ERR.NA;
    const attr = args[1] != null ? str1(args[1]).toLowerCase() : 'price';
    if (args[2] != null) {
      // 과거 시세: GOOGLEFINANCE(티커, "close", 시작, [끝 | 일 수], ["DAILY"|"WEEKLY"])
      const start = Math.floor(toNum(one(args[2])));
      const endArg = args[3] != null ? toNum(one(args[3])) : start;
      const end = endArg < 10000 ? start + endArg : Math.floor(endArg);
      const iv = args[4] != null ? (/^(week|7)/i.test(str1(args[4])) ? 'w' : 'd') : 'd';
      return netRange(historyUrl(tk, start, end, iv), ctx, (t) => {
        const hist = parseHistory(tk, t);
        const col = { open: 1, high: 2, low: 3, close: 4, price: 4, volume: 5 }[attr];
        if (col === undefined) return ERR.VALUE;
        if (attr === 'all') return grid([['Date', 'Open', 'High', 'Low', 'Close', 'Volume'], ...hist]);
        return grid([['Date', attr[0].toUpperCase() + attr.slice(1)], ...hist.map((r) => [r[0], r[col]])]);
      });
    }
    return netRange(quoteUrl(tk), ctx, (t) => {
      const q = parseQuote(tk, t);
      if (!q) return ERR.NA;
      const map = { price: q.price, priceopen: q.open, open: q.open, high: q.high, low: q.low, volume: q.volume, name: q.name, currency: q.currency, closeyest: q.closeyest, change: q.change, changepct: q.changepct === null || q.changepct === undefined ? null : q.changepct * 100, marketcap: q.marketcap, tradetime: q.tradetime };
      const v = map[attr];
      return v === undefined || v === null ? ERR.NA : v;
    });
  },
  GOOGLETRANSLATE: (args, ctx) => {
    const text = str1(args[0]);
    if (!text.trim()) return '';
    const sl = args[1] != null ? str1(args[1]) : 'auto';
    const tl = args[2] != null ? str1(args[2]) : 'ko';
    return netRange(translateUrl(text, sl, tl), ctx, (t) => parseTranslate(t).text);
  },
  DETECTLANGUAGE: (args, ctx) => {
    const text = flatVals(args).filter((x) => typeof x === 'string').join(' ').slice(0, 1000);
    if (!text.trim()) return ERR.VALUE;
    return netRange(translateUrl(text, 'auto', 'en'), ctx, (t) => parseTranslate(t).lang);
  },
  SPARKLINE: (args) => {
    const data = asRange(args[0]);
    const o = optionsFrom(args[1]);
    const vals = data.height === 1 ? data.rows[0] : data.width === 1 ? data.rows.map((r) => r[0]) : [...data.values()];
    const src = sparklineSvg(vals, o);
    return src ? new CellImage({ src, alt: '스파크라인', sizing: 1 }) : '';
  },
  QUERY: (args) => {
    const data = asRange(args[0]);
    const text = str1(args[1] ?? 'select *');
    let headers = args[2] != null ? Math.trunc(toNum(one(args[2]))) : -1;
    const rows = data.rows.map((r) => r.map((v) => (v === undefined ? null : v)));
    if (headers < 0) {
      // 추측: 첫 행이 모두 글자이고 아래에 글자가 아닌 값이 있는 열이 있으면 머리글
      const first = rows[0] ?? [];
      const allText = first.every((v) => typeof v === 'string' && v !== '');
      headers = allText && rows.slice(1).some((r) => r.some((v, i) => typeof v === 'number' && typeof first[i] === 'string')) ? 1 : 0;
    }
    const cols = data.ref ? rows[0]?.map((_, i) => colName(data.ref.c1 + i)) ?? [] : (rows[0] ?? []).map((_, i) => `Col${i + 1}`);
    const heads = (rows[0] ?? []).map((_, i) => rows.slice(0, headers).map((r) => toStr(r[i] ?? '')).filter(Boolean).join(' '));
    let q;
    try { q = parseQuery(text); } catch { return ERR.VALUE; }
    let res;
    try { res = runQuery(q, rows.slice(headers).filter((r) => r.some((v) => v !== null && v !== '')), cols, heads); } catch { return ERR.VALUE; }
    const showHead = headers > 0 || q.labels.length || res.head.some((h, i) => (q.select?.[i] && hasAgg(q.select[i])));
    const out = [...(showHead ? [res.head] : []), ...res.rows.map((r) => r.map((v) => v ?? ''))];
    return out.length ? grid(out) : ERR.NA;
  },
  ARRAYFORMULA: (args) => (args.length ? args[0] : ERR.NA),
  SPLIT: (args) => {
    const src = args[0] instanceof Range ? args[0] : new Range([[args[0]]]);
    const delim = str1(args[1]);
    const each = optBool(one(args[2]), true);
    const dropEmpty = optBool(one(args[3]), true);
    if (delim === '') return ERR.VALUE;
    const splitOne = (v) => {
      if (isError(v)) return [v];
      const s = toStr(v ?? '');
      const parts = each ? s.split(new RegExp(`[${[...delim].map((c) => c.replace(/[\]\\^-]/g, '\\$&')).join('')}]`)) : s.split(delim);
      const kept = dropEmpty ? parts.filter((p) => p !== '') : parts;
      return (kept.length ? kept : ['']).map(autoValue);
    };
    const lines = [...src.values()].map(splitOne);
    return grid(src.height * src.width === 1 ? lines : lines);
  },
  JOIN: (args) => {
    if (args.length < 2) return ERR.NA;
    const d = str1(args[0]);
    const vals = flatVals(args.slice(1));
    const bad = vals.find(isError);
    if (bad) return bad;
    return vals.map((v) => toStr(v ?? '')).join(d);
  },
  FLATTEN: (args) => {
    const vals = flatVals(args).map((v) => (v === undefined || v === null ? '' : v));
    return vals.length ? new Range(vals.map((v) => [v])) : ERR.NA;
  },
  SORTN: (args) => {
    const data = asRange(args[0]);
    const n = args[1] != null ? Math.trunc(toNum(one(args[1]))) : 1;
    const mode = args[2] != null ? Math.trunc(toNum(one(args[2]))) : 0;
    const keys = [];
    for (let i = 3; i < args.length; i += 2) keys.push({ c: Math.trunc(toNum(one(args[i]))) - 1, asc: optBool(one(args[i + 1]), true) });
    if (!keys.length) keys.push({ c: 0, asc: true });
    const rows = data.rows.map((r, i) => ({ r, i }));
    rows.sort((a, b) => {
      for (const k of keys) {
        const c = compareNulls(a.r[k.c], b.r[k.c]);
        if (c) return k.asc ? c : -c;
      }
      return a.i - b.i;
    });
    const kf = (r) => JSON.stringify(keys.map((k) => r[k.c]));
    let out;
    if (mode === 1) {
      out = rows.slice(0, n);
      const last = out.at(-1);
      if (last) for (const x of rows.slice(n)) { if (kf(x.r) === kf(last.r)) out.push(x); else break; }
    } else if (mode === 2 || mode === 3) {
      const seen = new Set();
      out = [];
      for (const x of rows) {
        const k = kf(x.r);
        if (seen.has(k)) { if (mode === 3) out.push(x); continue; }
        if (seen.size >= n) { if (mode === 2) break; continue; }
        seen.add(k);
        out.push(x);
      }
    } else out = rows.slice(0, n);
    return out.length ? new Range(out.map((x) => x.r)) : ERR.NA;
  },
  ARRAY_CONSTRAIN: (args) => {
    const r = asRange(args[0]);
    const h = Math.trunc(toNum(one(args[1])));
    const w = Math.trunc(toNum(one(args[2])));
    if (h < 1 || w < 1) return ERR.VALUE;
    return new Range(r.rows.slice(0, h).map((row) => row.slice(0, w)));
  },
  COUNTUNIQUE: (args) => {
    const seen = new Set();
    for (const v of flatVals(args)) if (!isBlankV(v)) seen.add(typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`);
    return seen.size;
  },
  COUNTUNIQUEIFS: (args) => {
    const target = asRange(args[0]);
    const crits = [];
    for (let i = 1; i < args.length; i += 2) crits.push({ r: asRange(args[i]), f: makeCriteria(one(args[i + 1])) });
    const seen = new Set();
    target.rows.forEach((row, r) => row.forEach((v, c) => {
      if (isBlankV(v)) return;
      if (crits.every((k) => k.f(k.r.rows[r]?.[c] ?? null))) seen.add(typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`);
    }));
    return seen.size;
  },
  'AVERAGE.WEIGHTED': (args) => {
    let s = 0;
    let w = 0;
    for (let i = 0; i < args.length; i += 2) {
      const vs = flatVals([args[i]]);
      const ws = flatVals([args[i + 1] ?? 1]);
      if (ws.length !== 1 && ws.length !== vs.length) return ERR.VALUE;
      vs.forEach((v, k) => {
        const wt = ws.length === 1 ? ws[0] : ws[k];
        if (typeof v === 'number' && typeof wt === 'number') { s += v * wt; w += wt; }
      });
    }
    return w === 0 ? ERR.DIV0 : s / w;
  },
  MARGINOFERROR: (args) => {
    const xs = collectNums([args[0]]);
    const conf = toNum(one(args[1]));
    if (xs.length < 2 || !(conf > 0 && conf < 1)) return ERR.NUM;
    const mean = xs.reduce((a, b) => a + b, 0) / xs.length;
    const sd = Math.sqrt(xs.reduce((a, b) => a + (b - mean) ** 2, 0) / (xs.length - 1));
    return normInv(1 - (1 - conf) / 2) * sd / Math.sqrt(xs.length);
  },
  ISEMAIL: lift(([v]) => typeof v === 'string' && /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(v.trim())),
  ISURL: lift(([v]) => typeof v === 'string' && /^(https?:\/\/|ftp:\/\/|www\.)?[\w-]+(\.[\w-]+)+(:\d+)?(\/\S*)?$/i.test(v.trim()) && /\.[a-z]{2,}/i.test(v)),
  ISBETWEEN: lift(([v, lo, hi, li, hin]) => {
    const a = compareValues(v, lo);
    const b = compareValues(v, hi);
    return (optBool(li, true) ? a >= 0 : a > 0) && (optBool(hin, true) ? b <= 0 : b < 0);
  }),
  TO_DATE: lift(([v]) => (typeof v === 'number' ? v : v ?? '')),
  TO_PERCENT: lift(([v]) => (typeof v === 'number' ? v : v ?? '')),
  TO_DOLLARS: lift(([v]) => (typeof v === 'number' ? v : v ?? '')),
  TO_TEXT: lift(([v]) => (v === null || v === undefined ? '' : toStr(v))),
  TO_PURE_NUMBER: lift(([v]) => (typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : (parseNumberText(String(v ?? '')) ?? ERR.VALUE))),
  EPOCHTODATE: lift(([t, unit]) => {
    const u = optInt(unit, 1);
    const ms = toNum(t) * ({ 1: 1000, 2: 1, 3: 0.001 }[u] ?? NaN);
    if (!Number.isFinite(ms) || ms < 0) return ERR.NUM;
    return msToSerial(ms);
  }),
  ADD: lift(([a, b]) => toNum(a) + toNum(b)),
  MINUS: lift(([a, b]) => toNum(a) - toNum(b)),
  MULTIPLY: lift(([a, b]) => toNum(a) * toNum(b)),
  DIVIDE: lift(([a, b]) => { const d = toNum(b); return d === 0 ? ERR.DIV0 : toNum(a) / d; }),
  POW: lift(([a, b]) => toNum(a) ** toNum(b)),
  EQ: lift(([a, b]) => compareValues(a ?? '', b ?? '') === 0),
  NE: lift(([a, b]) => compareValues(a ?? '', b ?? '') !== 0),
  GT: lift(([a, b]) => compareValues(a ?? '', b ?? '') > 0),
  GTE: lift(([a, b]) => compareValues(a ?? '', b ?? '') >= 0),
  LT: lift(([a, b]) => compareValues(a ?? '', b ?? '') < 0),
  LTE: lift(([a, b]) => compareValues(a ?? '', b ?? '') <= 0),
  UMINUS: lift(([a]) => -toNum(a)),
  UPLUS: lift(([a]) => (a === null || a === undefined ? 0 : a)),
  UNARY_PERCENT: lift(([a]) => toNum(a) / 100),

  // ── 엑셀: 웹 ──
  WEBSERVICE: (args, ctx) => {
    const url = str1(args[0]);
    if (!/^https?:\/\//i.test(url) || url.length > 2048) return ERR.VALUE;
    const t = netText(url, ctx);
    return t.length > 32767 ? ERR.VALUE : t;
  },
  ENCODEURL: lift(([s]) => encodeURIComponent(toStr(s ?? '')).replace(/[!'()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`)),
  FILTERXML: (args) => {
    const xml = str1(args[0]);
    const xpath = str1(args[1]);
    let hits;
    try { hits = evalPath(xpath, [parseMarkup(xml, { xml: true })]); } catch { return ERR.VALUE; }
    if (!hits.length) return ERR.VALUE;
    const vals = hits.map((h) => autoValue(nodeText(h)));
    return vals.length === 1 ? vals[0] : new Range(vals.map((v) => [v]));
  },
  // ── 엑셀 365: 번역 · 주식 기록 · 범위 자르기 ──
  TRANSLATE: (args, ctx) => {
    const text = str1(args[0]);
    if (!text.trim()) return '';
    return netRange(translateUrl(text, args[1] != null ? str1(args[1]) : 'auto', args[2] != null ? str1(args[2]) : 'ko'), ctx, (t) => parseTranslate(t).text);
  },
  STOCKHISTORY: (args, ctx) => {
    const tk = parseTicker(str1(args[0]));
    if (!tk) return ERR.VALUE;
    const start = Math.floor(toNum(one(args[1])));
    const end = args[2] != null ? Math.floor(toNum(one(args[2]))) : start;
    const iv = ['d', 'w', 'm'][optInt(one(args[3]), 0)] ?? 'd';
    const headers = optInt(one(args[4]), 1);
    const props = args.slice(5).map((p) => Math.trunc(toNum(one(p))));
    const pick = props.length ? props : [0, 1];
    return netRange(historyUrl(tk, start, end, iv), ctx, (t) => {
      const hist = parseHistory(tk, t);
      if (!hist.length) return ERR.NA;
      const names = ['날짜', '종가', '시가', '고가', '저가', '거래량'];
      const colOf = [0, 4, 1, 2, 3, 5];
      return grid([...(headers ? [pick.map((p) => names[p] ?? '')] : []), ...hist.map((r) => pick.map((p) => r[colOf[p]] ?? ''))]);
    });
  },
  TRIMRANGE: (args) => {
    const r = asRange(args[0]);
    const tr = optInt(one(args[1]), 3);
    const tc = optInt(one(args[2]), 3);
    const empty = (v) => v === null || v === undefined || v === '';
    let r1 = 0; let r2 = r.height - 1; let c1 = 0; let c2 = r.width - 1;
    const rowEmpty = (i) => r.rows[i].every(empty);
    const colEmpty = (j) => r.rows.every((row) => empty(row[j]));
    if (tr & 1) while (r1 <= r2 && rowEmpty(r1)) r1++;
    if (tr & 2) while (r2 >= r1 && rowEmpty(r2)) r2--;
    if (tc & 1) while (c1 <= c2 && colEmpty(c1)) c1++;
    if (tc & 2) while (c2 >= c1 && colEmpty(c2)) c2--;
    if (r1 > r2 || c1 > c2) return ERR.REF;
    return new Range(r.rows.slice(r1, r2 + 1).map((row) => row.slice(c1, c2 + 1).map((v) => (v === undefined ? null : v))));
  },
  // ── 엑셀: 누락된 호환 함수 ──
  ACCRINT: (args) => {
    const [issue, first, settle, rate, par, freq, basis, method] = args.map(one);
    const f = toNum(freq);
    if (![1, 2, 4].includes(f) || toNum(rate) <= 0 || toNum(par ?? 1000) <= 0) return ERR.NUM;
    const b = optInt(basis, 0);
    if (b < 0 || b > 4) return ERR.NUM;
    const s = Math.floor(toNum(settle));
    const from = optBool(method, true) || s <= toNum(first) ? Math.floor(toNum(issue)) : Math.floor(toNum(first));
    if (from >= s) return ERR.NUM;
    return optNum(par, 1000) * toNum(rate) * yearFrac(from, s, b);
  },
  BAHTTEXT: lift(([n]) => bahtText(toNum(n))),
  SEARCHB: lift(([find, within, start]) => {
    const f = toStr(find ?? '');
    const w = toStr(within ?? '');
    const st = optInt(start, 1);
    // 바이트 위치 → 글자 위치
    let bytes = 0;
    let ci = 0;
    const chs = [...w];
    while (ci < chs.length && bytes + 1 < st) { bytes += wideLen(chs[ci]); ci++; }
    const re = new RegExp(f.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?'), 'i');
    const rest = chs.slice(ci).join('');
    const m = re.exec(rest);
    if (!m) return ERR.VALUE;
    const before = [...rest.slice(0, m.index)].reduce((a, ch) => a + wideLen(ch), 0);
    return bytes + before + 1;
  }),
  'ECMA.CEILING': lift(([n, sig]) => {
    const x = toNum(n);
    const s = Math.abs(toNum(sig));
    if (s === 0) return 0;
    return Math.ceil(x / s) * s;
  }),
  'BINOM.DIST.RANGE': lift(([trials, p, s1, s2]) => {
    const n = Math.trunc(toNum(trials));
    const pr = toNum(p);
    const a = Math.trunc(toNum(s1));
    const b = s2 === null || s2 === undefined ? a : Math.trunc(toNum(s2));
    if (n < 0 || pr < 0 || pr > 1 || a < 0 || a > n || b < a || b > n) return ERR.NUM;
    let sum = 0;
    for (let k = a; k <= b; k++) sum += binomPmf(n, pr, k);
    return sum;
  }),
};
for (const k of ['GOOGLETRANSLATE', 'TRANSLATE', 'DETECTLANGUAGE', 'IMPORTXML', 'IMPORTHTML', 'IMPORTDATA', 'IMPORTFEED', 'IMPORTRANGE', 'GOOGLEFINANCE', 'STOCKHISTORY', 'WEBSERVICE']) WEB[k].net = true;
function colName(c) { let s = ''; let n = c + 1; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
/** 결과가 여러 칸일 수 있는 웹 함수 (분산 후보) */
export const WEB_ARRAY = ['IMPORTDATA', 'IMPORTHTML', 'IMPORTXML', 'IMPORTFEED', 'IMPORTRANGE', 'QUERY', 'SPLIT', 'FLATTEN', 'SORTN', 'ARRAY_CONSTRAIN', 'ARRAYFORMULA', 'GOOGLEFINANCE', 'STOCKHISTORY', 'FILTERXML', 'TRIMRANGE'];
