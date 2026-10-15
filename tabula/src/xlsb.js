// XLSB (Excel 바이너리 통합 문서) 읽기 — DOM 없음
// 바이너리 부분(.bin: 통합 문서 · 스타일 · 공유 문자열 · 시트 · 표 · 피벗 · 슬라이서 · 메모)을 같은 경로의 xlsx XML 로 바꾸고
// 기존 xlsx 리더가 그대로 읽음. 드로잉 · 차트 · 테마는 xlsb 에서도 XML 이라 그대로 씀.
// 큰 시트의 셀은 XML 을 만들지 않고 readSheet 가 받는 행 모양({ attrs, cells })으로 바로 넘김 (수백만 행도 빠르게).
// 레코드 구조: [MS-XLSB] (레코드 번호 · 필드 순서는 LibreOffice oox 가져오기와 같음)

const enc = new TextEncoder();
const RK = new DataView(new ArrayBuffer(8));
const u16dec = new TextDecoder('utf-16le');

// ─────────────── 바이트 읽기 ───────────────
const dvCache = new WeakMap();
const dvOf = (u8) => { let d = dvCache.get(u8); if (!d) { d = new DataView(u8.buffer, u8.byteOffset, u8.byteLength); dvCache.set(u8, d); } return d; };

class Rd {
  constructor(u8, p = 0, e = u8.length) { this.u8 = u8; this.dv = dvOf(u8); this.p = p; this.e = e; }
  get left() { return this.e - this.p; }
  u8v() { return this.u8[this.p++]; }
  u16() { const v = this.dv.getUint16(this.p, true); this.p += 2; return v; }
  i16() { const v = this.dv.getInt16(this.p, true); this.p += 2; return v; }
  i32() { const v = this.dv.getInt32(this.p, true); this.p += 4; return v; }
  u32() { const v = this.dv.getUint32(this.p, true); this.p += 4; return v; }
  f64() { const v = this.dv.getFloat64(this.p, true); this.p += 8; return v; }
  skip(n) { this.p += n; return this; }
  chars(n) { const s = utf16(this.u8, this.p, n); this.p += n * 2; return s; }
  /** XLWideString (32비트 길이, -1 = 없음) */
  str() { if (this.left < 4) return null; const n = this.i32(); if (n < 0) return null; return this.chars(Math.min(n, this.left >> 1)); }
  str16() { const n = this.u16(); return this.chars(Math.min(n, this.left >> 1)); }
  rfx() { return { r1: this.i32(), r2: this.i32(), c1: this.i32(), c2: this.i32() }; }
}

function utf16(u8, p, n) {
  if (n <= 0) return '';
  if (n < 24) {
    let s = '';
    for (let i = 0; i < n; i++) s += String.fromCharCode(u8[p + 2 * i] | (u8[p + 2 * i + 1] << 8));
    return s;
  }
  return u16dec.decode(u8.subarray(p, p + n * 2));
}

/** 레코드 목록 [{ t, p, e }] (작은 부분용) */
function records(u8) {
  const out = [];
  let p = 0;
  const n = u8.length;
  while (p < n) {
    let t = u8[p++];
    if (t & 0x80) t = (t & 0x7f) | ((u8[p++] & 0x7f) << 7);
    let sz = 0;
    for (let i = 0, sh = 0; i < 4; i++, sh += 7) { const x = u8[p++]; sz |= (x & 0x7f) << sh; if (!(x & 0x80)) break; }
    out.push({ t, p, e: p + sz });
    p += sz;
  }
  return out;
}

const R = {
  ROW: 0, BLANK: 1, RK: 2, ERROR: 3, BOOL: 4, REAL: 5, ST: 6, ISST: 7, FSTR: 8, FNUM: 9, FBOOL: 10, FERR: 11,
  MBLANK: 12, MRK: 13, MERR: 14, MBOOL: 15, MREAL: 16, MST: 17, MISST: 18, SI: 19, NAME: 39,
  FONT: 43, NUMFMT: 44, FILL: 45, BORDER: 46, XF: 47, CELLSTYLE: 48, COL: 60, MRSTR: 61, RSTR: 62,
  SHEET: 156, WBPR: 153, BOOKVIEW: 158, WSPR: 147, DIM: 148, WSVIEW: 137, PANE: 151, SHEETDATA: 145, SHEETDATA_END: 146,
  AFILTER: 161, AFILTER_END: 162, FILTERCOL: 163, FILTERCOL_END: 164, FILTERS: 165, FILTER: 167, TOP10: 170, CUSTFILTERS: 172, CUSTFILTER: 174,
  MERGE: 176, EXTSELF: 357, EXTSAME: 358, EXTREF: 355, EXTADDIN: 667, EXTSHEETS: 362, PIVOTCACHE: 386,
  ARRAY: 426, SHRFMLA: 427, CF: 461, CF_END: 462, CFRULE: 463, CFRULE_END: 464, ICONSET: 465, DATABAR: 467, COLORSCALE: 469, CFVO: 471, CFCOLOR: 564,
  MARGINS: 476, WSFMT: 485, HLINK: 494, DXF: 507, TSTYLES: 508, TSTYLE: 510, TSTYLE_END: 511, TSELEM: 512, TSINFO: 513,
  PROTECT: 535, DRAWING: 550, LEGACY: 551, TABLEPART: 661,
  CSXFS: 626, CXFS: 617, CAUTHOR: 632, COMMENT: 635, COMMENT_END: 636, CTEXT: 637,
  TABLE: 343, TABLE_END: 344, LISTCOL: 347, LISTCOL_END: 348, LISTCCFMLA: 351,
  PCDEF: 179, PCDSOURCE: 185, PCDSHEETSRC: 187, PCDFIELD: 183, PCDFIELD_END: 184, PCDFSITEMS: 189, PCDFSITEMS_END: 190, PCITEM_ARRAY: 191,
  PCDFGROUP: 219, PCDFGROUP_END: 220, PCDFGROUPITEMS: 221, PCDFGROUPITEMS_END: 222, PCDPNAMES: 253, PCDPNAME: 255, PCDFGRANGE: 223,
  PTDEF: 280, PTFITEM: 282, PTFIELD: 285, PTFIELD_END: 286, PTLOCATION: 314, PTROWFIELDS: 309, PTCOLFIELDS: 311, PTPAGEFIELD: 289,
  PTDATAFIELD: 293, PTROWITEMS: 299, PTROWITEMS_END: 300, PTCOLITEMS: 301, PTCOLITEMS_END: 302, PTLINE: 297, PTLINE_X: 388, PTFILTER: 601, PTFILTER_END: 602, PTREFERENCE: 251, PTREFITEM: 382, AUTOSORTSCOPE: 459,
  SLC_DEF: 1077, SLC_PIVOTS: 1085, SLC_TABULAR: 1100, SLC_ITEMS: 1102, SLICER: 1083,
};

// ─────────────── XML 도우미 ───────────────
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[m]))
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, '');
const attrs = (o) => Object.entries(o).filter(([, v]) => v !== undefined && v !== null && v !== false).map(([k, v]) => ` ${k}="${esc(v === true ? 1 : v)}"`).join('');
const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const cellRef = (r, c) => `${colName(c)}${r + 1}`;
const rangeRef = ({ r1, c1, r2, c2 }) => (r1 === r2 && c1 === c2 ? cellRef(r1, c1) : `${cellRef(r1, c1)}:${cellRef(r2, c2)}`);
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const put = (files, path, xml) => { files[path] = enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`); };

const ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#SPILL!', 0x2C: '#CALC!' };
const errText = (e) => ERRORS[e] ?? '#N/A';

/** BIFF12 색 (8바이트) → xlsx 색 속성 */
function color(rd) {
  const flags = rd.u8v();
  const idx = rd.u8v();
  const tint = rd.i16();
  const rgb = [rd.u8v(), rd.u8v(), rd.u8v()];
  rd.skip(1);
  const type = flags >> 1;
  const t = tint ? { tint: Math.round((tint < 0 ? tint / 32768 : tint / 32767) * 1e6) / 1e6 } : {};
  if (type === 0) return { auto: 1 };
  if (type === 1) return { indexed: idx, ...t };
  if (type === 3) return { theme: idx, ...t };
  return { rgb: `FF${rgb.map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`, ...t };
}
const colorXml = (tag, c) => (c ? `<${tag}${attrs(c)}/>` : '');

// ─────────────── 수식 (rgce → A1 수식 글자) ───────────────
// BIFF12 함수 번호 → 이름 (':n' = 인수 개수 고정)
const FTAB = ('COUNT|IF|ISNA:1|ISERROR:1|SUM|AVERAGE|MIN|MAX|ROW|COLUMN|NA:0|NPV|STDEV|DOLLAR|FIXED|SIN:1|COS:1|TAN:1|ATAN:1|PI:0|SQRT:1|EXP:1|LN:1|LOG10:1|ABS:1|INT:1|SIGN:1|ROUND:2|LOOKUP|INDEX|REPT:2|MID:3|LEN:1|VALUE:1|TRUE:0|FALSE:0|AND|OR|NOT:1|MOD:2|DCOUNT:3|DSUM:3|DAVERAGE:3|DMIN:3|DMAX:3|DSTDEV:3|VAR|DVAR:3|TEXT:2|LINEST|TREND|LOGEST|GROWTH|||RETURN|PV|FV|NPER|PMT|RATE|MIRR:3|IRR|RAND:0|MATCH|DATE:3|TIME:3|DAY:1|MONTH:1|YEAR:1|WEEKDAY:1|HOUR:1|MINUTE:1|SECOND:1|NOW:0|AREAS:1|ROWS:1|COLUMNS:1|OFFSET|ABSREF:2||ADD.ARROW:0|SEARCH|TRANSPOSE:1|||TYPE:1||||||||ACTIVE.CELL:0|||ATAN2:2|ASIN:1|ACOS:1|CHOOSE|HLOOKUP:3|VLOOKUP:3|ACTIVATE|ACTIVATE.NEXT:0|ISREF:1||||LOG||CHAR:1|LOWER:1|UPPER:1|PROPER:1|LEFT|RIGHT|EXACT:2|TRIM:1|REPLACE:4|SUBSTITUTE|CODE:1|||FIND|CELL|ISERR:1|ISTEXT:1|ISNUMBER:1|ISBLANK:1|T:1|N:1|||||||||DATEVALUE:1|TIMEVALUE:1|SLN:3|SYD:4|DDB||||INDIRECT|||ADD.BAR:0|ADD.MENU:2|ADD.COMMAND:3|||||||||CLEAN:1|MDETERM:1|MINVERSE:1|MMULT:2||IPMT|PPMT|COUNTA||||||||||||||PRODUCT|FACT:1|||||DPRODUCT:3|ISNONTEXT:1|||STDEVP|VARP|DSTDEVP:3|DVARP:3|TRUNC:1|ISLOGICAL:1|DCOUNTA:3|||||USDOLLAR|FINDB|SEARCHB|REPLACEB:4|LEFTB|RIGHTB|MIDB:3|LENB:1|ROUNDUP:2|ROUNDDOWN:2|ASC:1|DBCS:1|RANK|||ADDRESS|DAYS360:2|TODAY:0|VDB|||||MEDIAN|SUMPRODUCT|SINH:1|COSH:1|TANH:1|ASINH:1|ACOSH:1|ATANH:1|DGET:3|||||||||INFO:1|||DB|||||FREQUENCY:2|||EXTERN.CALL||||||ERROR.TYPE:1||||||||AVEDEV|BETADIST|GAMMALN:1|BETAINV|BINOMDIST:4|CHIDIST:2|CHIINV:2|COMBIN:2|CONFIDENCE:3|CRITBINOM:3|EVEN:1|EXPONDIST:3|FDIST:3|FINV:3|FISHER:1|FISHERINV:1|FLOOR:2|GAMMADIST:4|GAMMAINV:3|CEILING:2|HYPGEOMDIST:4|LOGNORMDIST:3|LOGINV:3|NEGBINOMDIST:3|NORMDIST:4|NORMSDIST:1|NORMINV:3|NORMSINV:1|STANDARDIZE:3|ODD:1|PERMUT:2|POISSON:3|TDIST:3|WEIBULL:4|SUMXMY2:2|SUMX2MY2:2|SUMX2PY2:2|CHITEST:2|CORREL:2|COVAR:2|FORECAST:3|FTEST:2|INTERCEPT:2|PEARSON:2|RSQ:2|STEYX:2|SLOPE:2|TTEST:4|PROB|DEVSQ|GEOMEAN|HARMEAN|SUMSQ|KURT|SKEW|ZTEST|LARGE:2|SMALL:2|QUARTILE:2|PERCENTILE:2|PERCENTRANK|MODE|TRIMMEAN:2|TINV:2||||CONCATENATE|POWER:2|||||RADIANS:1|DEGREES:1|SUBTOTAL|SUMIF|COUNTIF:2|COUNTBLANK:1|||ISPMT:4|DATEDIF:3|DATESTRING:1|NUMBERSTRING:2|ROMAN||||GETPIVOTDATA|HYPERLINK|PHONETIC:1|AVERAGEA|MAXA|MINA|STDEVPA|VARPA|STDEVA|VARA|BAHTTEXT:1|THAIDAYOFWEEK:1|THAIDIGIT:1|THAIMONTHOFYEAR:1|THAINUMSOUND:1|THAINUMSTRING:1|THAISTRINGLENGTH:1|ISTHAIDIGIT:1|ROUNDBAHTDOWN:1|ROUNDBAHTUP:1|THAIYEAR:1|RTD:3|CUBEVALUE|CUBEMEMBER|CUBEMEMBERPROPERTY:3|CUBERANKEDMEMBER|HEX2BIN|HEX2DEC:1|HEX2OCT|DEC2BIN|DEC2HEX|DEC2OCT|OCT2BIN|OCT2HEX|OCT2DEC:1|BIN2DEC:1|BIN2OCT|BIN2HEX|IMSUB:2|IMDIV:2|IMPOWER:2|IMABS:1|IMSQRT:1|IMLN:1|IMLOG2:1|IMLOG10:1|IMSIN:1|IMCOS:1|IMEXP:1|IMARGUMENT:1|IMCONJUGATE:1|IMAGINARY:1|IMREAL:1|COMPLEX|IMSUM|IMPRODUCT|SERIESSUM:4|FACTDOUBLE:1|SQRTPI:1|QUOTIENT:2|DELTA|GESTEP|ISEVEN:1|ISODD:1|MROUND:2|ERF|ERFC:1|BESSELJ:2|BESSELK:2|BESSELY:2|BESSELI:2|XIRR|XNPV:3|PRICEMAT|YIELDMAT|INTRATE|RECEIVED|DISC|PRICEDISC|YIELDDISC|TBILLEQ:3|TBILLPRICE:3|TBILLYIELD:3|PRICE|YIELD|DOLLARDE:2|DOLLARFR:2|NOMINAL:2|EFFECT:2|CUMPRINC:6|CUMIPMT:6|EDATE:2|EOMONTH:2|YEARFRAC|COUPDAYBS|COUPDAYS|COUPDAYSNC|COUPNCD|COUPNUM|COUPPCD|DURATION|MDURATION|ODDLPRICE|ODDLYIELD|ODDFPRICE|ODDFYIELD|RANDBETWEEN:2|WEEKNUM|AMORDEGRC|AMORLINC|CONVERT:3|ACCRINT|ACCRINTM|WORKDAY|NETWORKDAYS|GCD|MULTINOMIAL|LCM|FVSCHEDULE:2|CUBEKPIMEMBER|CUBESET|CUBESETCOUNT:1|IFERROR:2|COUNTIFS|SUMIFS|AVERAGEIF|AVERAGEIFS')
  .split('|').map((s) => { const [n, a] = s.split(':'); return { name: n, argc: a === undefined ? -1 : Number(a) }; });

const OPS = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0A: '<=', 0x0B: '=', 0x0C: '>=', 0x0D: '>', 0x0E: '<>', 0x0F: ' ', 0x10: ',', 0x11: ':' };
const MAX_ROW = 1048575;
const MAX_COL = 16383;

/** 시트 이름 (필요하면 작은따옴표) */
function quoteSheet(name) {
  if (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name) && !/^(TRUE|FALSE)$/i.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

function refPart(row, col, rowRel, colRel) {
  return `${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
}

/** 셀 참조 필드 → { row, col, rowRel, colRel } (N 형식은 기준 셀에서의 상대 위치) */
function locOf(row, colField, base, relOffset) {
  let col = colField & 0x3fff;
  const colRel = !!(colField & 0x4000);
  const rowRel = !!(colField & 0x8000);
  let r = row;
  if (relOffset) {
    if (rowRel) { let o = row & 0xfffff; if (o > 0x7ffff) o -= 0x100000; r = (((base.r + o) % (MAX_ROW + 1)) + MAX_ROW + 1) % (MAX_ROW + 1); }
    if (colRel) { let o = col; if (o > 0x1fff) o -= 0x4000; col = (((base.c + o) % (MAX_COL + 1)) + MAX_COL + 1) % (MAX_COL + 1); }
  }
  return { row: r & 0xfffff, col, rowRel, colRel };
}

function areaText(a, b) {
  if (a.row === 0 && b.row === MAX_ROW && !a.rowRel === !b.rowRel) return `${a.colRel ? '' : '$'}${colName(a.col)}:${b.colRel ? '' : '$'}${colName(b.col)}`;
  if (a.col === 0 && b.col === MAX_COL) return `${a.rowRel ? '' : '$'}${a.row + 1}:${b.rowRel ? '' : '$'}${b.row + 1}`;
  const s1 = refPart(a.row, a.col, a.rowRel, a.colRel);
  const s2 = refPart(b.row, b.col, b.rowRel, b.colRel);
  return s1 === s2 ? s1 : `${s1}:${s2}`;
}

function numText(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  return String(n).toUpperCase().replace('E+', 'E+');
}

const strLit = (s) => `"${s.replace(/"/g, '""')}"`;

/** 구조적 참조 열 이름 이스케이프 ([ ] # ' 앞에 ') */
const colEsc = (s) => s.replace(/(['#[\]])/g, "'$1");

/**
 * rgce(+rgcb) → 수식 글자 (앞의 '=' 없음, 파일 형식 — _xlfn. 접두사 그대로)
 * env: { sheets, names, xti, tables, pnames }, base: 수식 셀 { r, c }, opts.relN: 참조를 모두 기준 셀 상대로 (공유 수식 · 조건부 서식)
 * 반환: { text, next } — text 가 null 이면 해석할 수 없는 수식 (값만 가져옴)
 */
export function decodeFormula(u8, p, env, base, opts = {}) {
  const dv0 = dvOf(u8);
  const cce = dv0.getInt32(p, true);
  const end = p + 4 + cce;
  if (cce < 0 || end + 4 > u8.length) return { text: null, next: Math.min(u8.length, end) };
  const cb = dv0.getInt32(end, true);
  const addEnd = end + 4 + cb;
  const next = addEnd;
  if (cb < 0 || addEnd > u8.length) return { text: null, next: end };
  // 같은 바이트의 수식(표의 계산 열 등)은 한 번만 해석 — 셀 위치에 따라 달라지는 토큰이 없을 때만
  const memo = env.memo;
  let key = null;
  if (memo && addEnd - p <= 512 && !opts.relN) {
    key = String.fromCharCode.apply(null, u8.subarray(p, addEnd));
    const hit = memo.get(key);
    if (hit !== undefined) return { text: hit, next };
  }
  let based = false;
  const rd = new Rd(u8, p + 4);
  const ex = new Rd(u8, end + 4);
  const st = [];
  const pop = (n) => st.splice(st.length - n, n);
  const sheetOf = (ixti) => {
    const x = env.xti[ixti];
    if (!x) return null;
    if (x.self) {
      if (x.first < 0 || x.last < 0) return x.first === -2 ? '' : null;
      const a = env.sheets[x.first];
      const b = env.sheets[x.last];
      if (a === undefined || b === undefined) return null;
      return `${a === b ? quoteSheet(a) : quoteSheet(`${a}:${b}`)}!`;
    }
    return null;
  };
  const nameOf = (idx) => env.names[idx - 1]?.name ?? null;
  try {
    while (rd.p < end) {
      const tok = rd.u8v();
      const cls = tok & 0x60;
      const id = tok & 0x1f;
      if (!cls) {
        if (OPS[id] !== undefined) { const [a, b] = pop(2); st.push(`${a}${OPS[id]}${b}`); continue; }
        switch (id) {
          case 0x01: { // PtgExp: 공유 · 배열 수식의 기준 셀 (열은 추가 데이터에)
            const row = rd.i32();
            const col = ex.i32();
            based = true;
            const t = env.shared?.(row, col, base);
            if (t == null) return { text: null, next, exp: { row, col } };
            st.push(t);
            break;
          }
          case 0x12: st.push(`+${st.pop()}`); break;
          case 0x13: st.push(`-${st.pop()}`); break;
          case 0x14: st.push(`${st.pop()}%`); break;
          case 0x15: st.push(`(${st.pop()})`); break;
          case 0x16: st.push(''); break;
          case 0x17: { const n = rd.u16(); st.push(strLit(rd.chars(n))); break; }
          case 0x18: { // NLR: 표 참조(PtgList) · 피벗 필드 이름
            const eptg = rd.u8v();
            if (eptg === 0x1D) { const i = rd.u32(); const f = env.pnames?.[i]; if (f === undefined) return { text: null, next }; st.push(/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(f) ? f : `'${f.replace(/'/g, "''")}'`); break; }
            if (eptg !== 0x19) return { text: null, next };
            rd.skip(2);
            const flags = rd.u16();
            const tid = rd.u16();
            rd.skip(2);
            const k1 = rd.u16();
            const k2 = rd.u16();
            const t = env.tables?.get(tid);
            if (!t) { st.push('#REF!'); break; }
            const parts = [];
            if (flags & 0x04) parts.push('[#All]');
            else {
              if (flags & 0x08) parts.push('[#Headers]');
              if (flags & 0x10) parts.push('[#Data]');
              if (flags & 0x20) parts.push('[#Totals]');
              if (flags & 0x40) parts.push('[#This Row]');
            }
            const cn = (k) => `[${colEsc(t.cols[k] ?? `열${k + 1}`)}]`;
            if (flags & 0x01) parts.push(cn(k1));
            else if (flags & 0x02) parts.push(`${cn(k1)}:${cn(k2)}`);
            if (!parts.length) st.push(`${t.name}[]`);
            else if (parts.length === 1) st.push(`${t.name}${parts[0]}`);
            else st.push(`${t.name}[${parts.join(',')}]`);
            break;
          }
          case 0x19: { // PtgAttr
            const ty = rd.u8v();
            if (ty === 0x04) { const n = rd.u16(); rd.skip(2 * n + 2); } else if (ty === 0x10) { rd.skip(2); st.push(`SUM(${st.pop()})`); } else rd.skip(2);
            break;
          }
          case 0x1C: st.push(errText(rd.u8v())); break;
          case 0x1D: st.push(rd.u8v() ? 'TRUE' : 'FALSE'); break;
          case 0x1E: st.push(String(rd.u16())); break;
          case 0x1F: st.push(numText(rd.f64())); break;
          default: return { text: null, next };
        }
        continue;
      }
      switch (id) {
        case 0x00: { // 배열 상수 (값은 추가 데이터에)
          rd.skip(14);
          const rows = ex.i32();
          const cols = ex.i32();
          const out = [];
          for (let i = 0; i < rows; i++) {
            const line = [];
            for (let j = 0; j < cols; j++) {
              const ty = ex.u8v();
              if (ty === 0) line.push(numText(ex.f64()));
              else if (ty === 1) line.push(strLit(ex.str16()));
              else if (ty === 2) line.push(ex.u8v() ? 'TRUE' : 'FALSE');
              else if (ty === 4) { line.push(errText(ex.u8v())); ex.skip(3); } else return { text: null, next };
            }
            out.push(line.join(','));
          }
          st.push(`{${out.join(';')}}`);
          break;
        }
        case 0x01: { // PtgFunc
          const fid = rd.u16();
          const f = FTAB[fid];
          if (!f?.name || f.argc < 0) return { text: null, next };
          st.push(`${f.name}(${pop(f.argc).join(',')})`);
          break;
        }
        case 0x02: { // PtgFuncVar
          const argc = rd.u8v() & 0x7f;
          const fid = rd.u16() & 0x7fff;
          const args = pop(argc);
          if (fid === 255) { const nm = args.shift(); st.push(`${nm}(${args.join(',')})`); break; }
          const f = FTAB[fid];
          if (!f?.name) return { text: null, next };
          st.push(`${f.name}(${args.join(',')})`);
          break;
        }
        case 0x03: { const n = nameOf(rd.i32()); if (n === null) return { text: null, next }; st.push(n); break; }
        case 0x04: case 0x0C: {
          if (id === 0x0C) based = true;
          const row = rd.i32(); const cf = rd.u16();
          const l = locOf(row, cf, base, id === 0x0C || opts.relN);
          st.push(refPart(l.row, l.col, l.rowRel, l.colRel));
          break;
        }
        case 0x05: case 0x0D: {
          const r1 = rd.i32(); const r2 = rd.i32(); const c1 = rd.u16(); const c2 = rd.u16();
          if (id === 0x0D) based = true;
          const n = id === 0x0D || opts.relN;
          st.push(areaText(locOf(r1, c1, base, n), locOf(r2, c2, base, n)));
          break;
        }
        case 0x06: { rd.skip(6); const n = ex.i32(); ex.skip(16 * Math.max(0, n)); break; } // PtgMemArea: 뒤따르는 식이 값
        case 0x07: case 0x08: rd.skip(6); break;
        case 0x09: case 0x0E: case 0x0F: rd.skip(2); break;
        case 0x0A: rd.skip(6); st.push('#REF!'); break;
        case 0x0B: rd.skip(12); st.push('#REF!'); break;
        case 0x19: { // PtgNameX (미래 함수 _xlfn.* · 다른 통합 문서 이름)
          const ixti = rd.i16(); const ni = rd.i32();
          const x = env.xti[ixti];
          const n = x?.self || !x ? nameOf(ni) : null;
          if (n === null) return { text: null, next };
          st.push(n);
          break;
        }
        case 0x1A: {
          const ixti = rd.i16(); const row = rd.i32(); const cf = rd.u16();
          const sp = sheetOf(ixti);
          const l = locOf(row, cf, base, opts.relN);
          st.push(sp === null ? '#REF!' : `${sp}${refPart(l.row, l.col, l.rowRel, l.colRel)}`);
          break;
        }
        case 0x1B: {
          const ixti = rd.i16(); const r1 = rd.i32(); const r2 = rd.i32(); const c1 = rd.u16(); const c2 = rd.u16();
          const sp = sheetOf(ixti);
          st.push(sp === null ? '#REF!' : `${sp}${areaText(locOf(r1, c1, base, opts.relN), locOf(r2, c2, base, opts.relN))}`);
          break;
        }
        case 0x1C: rd.skip(8); st.push('#REF!'); break;
        case 0x1D: rd.skip(14); st.push('#REF!'); break;
        default: return { text: null, next };
      }
    }
  } catch {
    return { text: null, next };
  }
  if (st.length !== 1) return { text: null, next };
  if (key !== null && !based) memo.set(key, st[0]);
  return { text: st[0], next };
}

// ─────────────── 통합 문서 ───────────────
function readWorkbook(u8) {
  const out = { sheets: [], names: [], xti: [], links: [], date1904: false, activeTab: 0, firstSheet: 0, pivotCaches: [] };
  for (const { t, p, e } of records(u8)) {
    const rd = new Rd(u8, p, e);
    if (t === R.SHEET) {
      const state = rd.u32(); const sheetId = rd.u32(); const rid = rd.str(); const name = rd.str() ?? '';
      out.sheets.push({ name, sheetId, rid, state: state === 1 ? 'hidden' : state === 2 ? 'veryHidden' : null });
    } else if (t === R.WBPR) {
      out.date1904 = !!(rd.u32() & 1);
    } else if (t === R.BOOKVIEW) {
      rd.skip(20); out.firstSheet = rd.u32(); out.activeTab = rd.u32();
    } else if (t === R.NAME) {
      const flags = rd.u32(); rd.skip(1); const itab = rd.i32(); let name = rd.str() ?? '';
      if (flags & 0x20) name = `_xlnm.${name}`;
      out.names.push({ name, itab, hidden: !!(flags & 1), func: !!(flags & 0x0e), fp: rd.p, fe: e });
    } else if (t === R.EXTSELF || t === R.EXTSAME) {
      out.links.push({ self: true });
    } else if (t === R.EXTREF || t === R.EXTADDIN) {
      out.links.push({ self: false });
    } else if (t === R.EXTSHEETS) {
      const n = rd.i32();
      for (let i = 0; i < n && rd.left >= 12; i++) {
        const li = rd.i32(); const first = rd.i32(); const last = rd.i32();
        out.xti.push({ self: !!out.links[li]?.self, first, last });
      }
    } else if (t === R.PIVOTCACHE) {
      out.pivotCaches.push({ id: rd.i32(), rid: rd.str() });
    }
  }
  return out;
}

function workbookXml(wbi, env, u8) {
  const names = [];
  wbi.names.forEach((n) => {
    if (n.func || /^_xlfn\./i.test(n.name)) return;
    const { text } = decodeFormula(u8, n.fp, env, { r: 0, c: 0 });
    if (text === null) return;
    names.push(`<definedName${attrs({ name: n.name, localSheetId: n.itab >= 0 ? n.itab : undefined, hidden: n.hidden || undefined })}>${esc(text)}</definedName>`);
  });
  return `<workbook xmlns="${NS}" xmlns:r="${NS_R}"><workbookPr${wbi.date1904 ? ' date1904="1"' : ''}/><bookViews><workbookView${attrs({ activeTab: wbi.activeTab, firstSheet: wbi.firstSheet || undefined })}/></bookViews><sheets>${wbi.sheets.map((s) => `<sheet${attrs({ name: s.name, sheetId: s.sheetId, state: s.state ?? undefined, 'r:id': s.rid })}/>`).join('')}</sheets>${names.length ? `<definedNames>${names.join('')}</definedNames>` : ''}${wbi.pivotCaches.length ? `<pivotCaches>${wbi.pivotCaches.map((c) => `<pivotCache cacheId="${c.id}" r:id="${esc(c.rid)}"/>`).join('')}</pivotCaches>` : ''}</workbook>`;
}

// ─────────────── 공유 문자열 ───────────────
function readStrings(u8) {
  const out = [];
  for (const { t, p, e } of records(u8)) {
    if (t !== R.SI) continue;
    const rd = new Rd(u8, p, e);
    rd.skip(1);
    out.push(rd.str() ?? '');
  }
  return out;
}

// ─────────────── 스타일 ───────────────
const H_ALIGN = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN = ['top', 'center', 'bottom', 'justify', 'distributed'];
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const BORDERS = ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
const UNDERLINE = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };
const TSE_TYPES = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell', 'firstSubtotalColumn', 'secondSubtotalColumn', 'thirdSubtotalColumn', 'firstSubtotalRow', 'secondSubtotalRow', 'thirdSubtotalRow', 'blankRow', 'firstColumnSubheading', 'secondColumnSubheading', 'thirdColumnSubheading', 'firstRowSubheading', 'secondRowSubheading', 'thirdRowSubheading', 'pageFieldLabels', 'pageFieldValues'];

function alignXml(flags) {
  const h = H_ALIGN[(flags >> 16) & 7];
  const v = V_ALIGN[(flags >> 19) & 7] ?? 'bottom';
  const a = {
    horizontal: h !== 'general' ? h : undefined, vertical: v !== 'bottom' ? v : undefined, textRotation: flags & 0xff || undefined,
    indent: (flags >> 8) & 0xff || undefined, wrapText: flags & 0x400000 ? 1 : undefined, shrinkToFit: flags & 0x1000000 ? 1 : undefined,
    readingOrder: (flags >> 26) & 3 || undefined,
  };
  return Object.values(a).some((x) => x !== undefined) ? `<alignment${attrs(a)}/>` : '';
}

function stylesXml(u8) {
  const numFmts = []; const fonts = []; const fills = []; const borders = []; const csxfs = []; const cxfs = []; const cellStyles = []; const dxfs = []; const tstyles = [];
  let xfTarget = null;
  let curTs = null;
  for (const { t, p, e } of records(u8)) {
    const rd = new Rd(u8, p, e);
    if (t === R.CSXFS) xfTarget = csxfs;
    else if (t === R.CXFS) xfTarget = cxfs;
    else if (t === R.NUMFMT) { const id = rd.u16(); numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(rd.str() ?? '')}"/>`); } else if (t === R.FONT) {
      const h = rd.u16(); const fl = rd.u16(); const w = rd.u16(); const escp = rd.u16(); const ul = rd.u8v(); const fam = rd.u8v(); const cs = rd.u8v(); rd.skip(1);
      const c = color(rd); const scheme = rd.u8v(); const name = rd.str() ?? '';
      fonts.push(`<font>${w >= 700 ? '<b/>' : ''}${fl & 2 ? '<i/>' : ''}${fl & 8 ? '<strike/>' : ''}${fl & 0x10 ? '<outline/>' : ''}${fl & 0x20 ? '<shadow/>' : ''}${UNDERLINE[ul] ? `<u${UNDERLINE[ul] === 'single' ? '' : ` val="${UNDERLINE[ul]}"`}/>` : ''}${escp === 1 ? '<vertAlign val="superscript"/>' : escp === 2 ? '<vertAlign val="subscript"/>' : ''}<sz val="${h / 20}"/>${colorXml('color', c.auto ? null : c)}<name val="${esc(name)}"/>${fam ? `<family val="${fam}"/>` : ''}${cs ? `<charset val="${cs}"/>` : ''}${scheme === 1 ? '<scheme val="major"/>' : scheme === 2 ? '<scheme val="minor"/>' : ''}</font>`);
    } else if (t === R.FILL) {
      const pat = rd.i32();
      if (pat === 40) {
        const fg = color(rd); color(rd); // 사용 안 함
        const type = rd.i32(); const deg = rd.f64(); const l = rd.f64(); const rr = rd.f64(); const tp = rd.f64(); const b = rd.f64();
        const n = rd.i32();
        const stops = [];
        for (let i = 0; i < n && rd.left >= 16; i++) { const c = color(rd); const pos = rd.f64(); stops.push(`<stop position="${pos}">${colorXml('color', c)}</stop>`); }
        fills.push(`<fill><gradientFill${attrs({ type: type === 1 ? 'path' : undefined, degree: deg || undefined, left: l || undefined, right: rr || undefined, top: tp || undefined, bottom: b || undefined })}>${stops.join('') || colorXml('color', fg)}</gradientFill></fill>`);
      } else {
        const fg = color(rd); const bg = color(rd);
        const pt = PATTERNS[pat] ?? 'none';
        fills.push(`<fill><patternFill patternType="${pt}">${pt !== 'none' ? colorXml('fgColor', fg) + colorXml('bgColor', bg) : ''}</patternFill></fill>`);
      }
    } else if (t === R.BORDER) {
      const fl = rd.u8v();
      const side = () => { const s = BORDERS[rd.u16()] ?? 'none'; const c = color(rd); return { s, c }; };
      const top = side(); const bottom = side(); const left = side(); const right = side(); const diag = side();
      const x = (tag, b) => (b.s === 'none' ? `<${tag}/>` : `<${tag} style="${b.s}">${colorXml('color', b.c)}</${tag}>`);
      borders.push(`<border${attrs({ diagonalDown: fl & 1 ? 1 : undefined, diagonalUp: fl & 2 ? 1 : undefined })}>${x('left', left)}${x('right', right)}${x('top', top)}${x('bottom', bottom)}${x('diagonal', diag)}</border>`);
    } else if (t === R.XF && xfTarget) {
      const parent = rd.u16(); const fmt = rd.u16(); const font = rd.u16(); const fill = rd.u16(); const border = rd.u16(); const fl = rd.u32();
      const prot = !(fl & 0x10000000) || fl & 0x20000000 ? `<protection${attrs({ locked: fl & 0x10000000 ? undefined : 0, hidden: fl & 0x20000000 ? 1 : undefined })}/>` : '';
      const al = alignXml(fl);
      xfTarget.push(`<xf${attrs({ numFmtId: fmt, fontId: font, fillId: fill, borderId: border, xfId: xfTarget === cxfs ? parent : undefined, applyNumberFormat: 1, applyFont: 1, applyFill: 1, applyBorder: 1, applyAlignment: al ? 1 : undefined, applyProtection: prot ? 1 : undefined })}>${al}${prot}</xf>`);
    } else if (t === R.CELLSTYLE) {
      const xf = rd.i32(); const fl = rd.u16(); const bi = rd.u8v(); const lvl = rd.u8v(); const name = rd.str() ?? '';
      cellStyles.push(`<cellStyle${attrs({ name, xfId: xf, builtinId: fl & 1 ? bi : undefined, iLevel: fl & 1 && (bi === 1 || bi === 2) ? lvl : undefined, hidden: fl & 2 ? 1 : undefined, customBuiltin: fl & 4 ? 1 : undefined })}/>`);
    } else if (t === R.DXF) {
      dxfs.push(dxfXml(rd));
    } else if (t === R.TSTYLES) {
      const n = rd.i32(); const defT = rd.str(); const defP = rd.str();
      tstyles.defs = { count: n, defaultTableStyle: defT, defaultPivotStyle: defP };
    } else if (t === R.TSTYLE) {
      const fl = rd.u16(); rd.skip(2); rd.u32(); const name = rd.str() ?? '';
      curTs = { name, pivot: !(fl & 2) ? undefined : 0, table: !(fl & 4) ? undefined : 0, els: [] };
    } else if (t === R.TSELEM && curTs) {
      const type = rd.i32(); const size = rd.i32(); const dxf = rd.i32();
      curTs.els.push(`<tableStyleElement${attrs({ type: TSE_TYPES[type] ?? 'wholeTable', size: size > 1 ? size : undefined, dxfId: dxf })}/>`);
    } else if (t === R.TSTYLE_END && curTs) {
      tstyles.push(`<tableStyle${attrs({ name: curTs.name, count: curTs.els.length })}>${curTs.els.join('')}</tableStyle>`);
      curTs = null;
    }
  }
  const sec = (tag, list) => `<${tag} count="${list.length}">${list.join('')}</${tag}>`;
  return `<styleSheet xmlns="${NS}">${numFmts.length ? sec('numFmts', numFmts) : ''}${sec('fonts', fonts)}${sec('fills', fills)}${sec('borders', borders)}${sec('cellStyleXfs', csxfs)}${sec('cellXfs', cxfs)}${sec('cellStyles', cellStyles)}${sec('dxfs', dxfs)}<tableStyles${attrs({ count: tstyles.length, defaultTableStyle: tstyles.defs?.defaultTableStyle ?? undefined, defaultPivotStyle: tstyles.defs?.defaultPivotStyle ?? undefined })}>${tstyles.join('')}</tableStyles></styleSheet>`;
}

/** BrtDXF → <dxf> */
function dxfXml(rd) {
  rd.skip(4);
  const n = rd.u16();
  const font = []; let fontColor = ''; let fill = { pat: null, fg: null, bg: null }; const border = {}; const align = {}; const prot = {};
  let fmtId = null; let fmtCode = null;
  for (let i = 0; i < n && rd.left >= 4; i++) {
    const at = rd.p;
    const type = rd.u16(); const size = rd.u16();
    const endAt = at + size;
    switch (type) {
      case 0: fill.pat = PATTERNS[rd.u8v()] ?? 'none'; break;
      case 1: fill.fg = color(rd); break;
      case 2: fill.bg = color(rd); break;
      case 5: fontColor = colorXml('color', color(rd)); break;
      case 6: case 7: case 8: case 9: case 10: case 11: case 12: {
        const c = color(rd); const s = BORDERS[rd.u16()] ?? 'none';
        border[['top', 'bottom', 'left', 'right', 'diagonal', 'vertical', 'horizontal'][type - 6]] = { s, c };
        break;
      }
      case 15: align.horizontal = H_ALIGN[rd.u8v()]; break;
      case 16: align.vertical = V_ALIGN[rd.u8v()]; break;
      case 17: align.textRotation = rd.u8v(); break;
      case 18: align.indent = rd.u16(); break;
      case 20: align.wrapText = rd.u8v() ? 1 : 0; break;
      case 22: align.shrinkToFit = rd.u8v() ? 1 : 0; break;
      case 24: font.push(`<name val="${esc(rd.str16())}"/>`); break;
      case 25: font.push(rd.u16() >= 700 ? '<b/>' : '<b val="0"/>'); break;
      case 26: { const u = UNDERLINE[rd.u16()]; font.push(u ? `<u${u === 'single' ? '' : ` val="${u}"`}/>` : '<u val="none"/>'); break; }
      case 27: { const v = rd.u16(); if (v === 1 || v === 2) font.push(`<vertAlign val="${v === 1 ? 'superscript' : 'subscript'}"/>`); break; }
      case 28: font.push(rd.u8v() ? '<i/>' : '<i val="0"/>'); break;
      case 29: font.push(rd.u8v() ? '<strike/>' : '<strike val="0"/>'); break;
      case 36: font.push(`<sz val="${rd.u32() / 20}"/>`); break;
      case 38: fmtCode = rd.str16(); break;
      case 41: fmtId = rd.u16(); break;
      case 43: prot.locked = rd.u8v() ? 1 : 0; break;
      case 44: prot.hidden = rd.u8v() ? 1 : 0; break;
      default: break;
    }
    rd.p = endAt;
  }
  const fontXml = font.length || fontColor ? `<font>${font.join('')}${fontColor}</font>` : '';
  const numXml = fmtCode !== null ? `<numFmt numFmtId="${fmtId ?? 164}" formatCode="${esc(fmtCode)}"/>` : '';
  const fillXml = fill.pat || fill.fg || fill.bg ? `<fill><patternFill${fill.pat ? ` patternType="${fill.pat}"` : ''}>${colorXml('fgColor', fill.fg)}${colorXml('bgColor', fill.bg)}</patternFill></fill>` : '';
  const bx = (k) => (border[k] ? (border[k].s === 'none' ? `<${k}/>` : `<${k} style="${border[k].s}">${colorXml('color', border[k].c)}</${k}>`) : '');
  const borderXml = Object.keys(border).length ? `<border>${['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'].map(bx).join('')}</border>` : '';
  const alignXml2 = Object.keys(align).length ? `<alignment${attrs(align)}/>` : '';
  const protXml = Object.keys(prot).length ? `<protection${attrs(prot)}/>` : '';
  return `<dxf>${fontXml}${numXml}${fillXml}${alignXml2}${borderXml}${protXml}</dxf>`;
}

// ─────────────── 표 ───────────────
const TOTALS_FN = ['none', 'sum', 'min', 'max', 'average', 'count', 'countNums', 'stdDev', 'var', 'custom'];

/** 표 (이름 · 열 이름) — 수식의 구조적 참조에 먼저 필요 */
function scanTable(u8) {
  let t = null;
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    if (r.t === R.TABLE) {
      const ref = rd.rfx(); rd.i32(); const id = rd.i32(); const hdr = rd.i32(); const tot = rd.i32(); rd.skip(32);
      const name = rd.str(); const display = rd.str();
      t = { id, ref, hdr, tot, name: display || name || `표${id}`, cols: [] };
    } else if (r.t === R.LISTCOL && t) {
      rd.skip(24); rd.str(); t.cols.push(rd.str() ?? `열${t.cols.length + 1}`);
    }
  }
  return t;
}

function tableXml(u8, env) {
  let head = '';
  let t = null;
  const cols = [];
  let col = null;
  let style = '';
  let af = '';
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    if (r.t === R.TABLE) {
      const ref = rd.rfx(); rd.i32(); const id = rd.i32(); const hdr = rd.i32(); const tot = rd.i32(); rd.skip(32);
      const name = rd.str(); const display = rd.str();
      t = { ref, id };
      head = attrs({ id, name: name || display, displayName: display || name, ref: rangeRef(ref), headerRowCount: hdr === 1 ? undefined : hdr, totalsRowCount: tot || undefined, totalsRowShown: tot ? undefined : 0 });
    } else if (r.t === R.AFILTER && t) {
      af = `<autoFilter ref="${rangeRef(rd.rfx())}"/>`;
    } else if (r.t === R.LISTCOL) {
      const fid = rd.i32(); const ilta = rd.i32(); rd.skip(16); rd.str(); const caption = rd.str() ?? '';
      col = { fid, name: caption, fn: TOTALS_FN[ilta] ?? 'none', f: null };
      cols.push(col);
    } else if (r.t === R.LISTCCFMLA && col && t) {
      rd.skip(1);
      const { text } = decodeFormula(u8, rd.p, env, { r: t.ref.r1 + 1, c: t.ref.c1 + cols.length - 1 });
      if (text !== null) col.f = text;
    } else if (r.t === R.TSINFO) {
      const fl = rd.u16(); const name = rd.str();
      style = `<tableStyleInfo${attrs({ name: name ?? undefined, showFirstColumn: fl & 1 ? 1 : 0, showLastColumn: fl & 2 ? 1 : 0, showRowStripes: fl & 4 ? 1 : 0, showColumnStripes: fl & 8 ? 1 : 0 })}/>`;
    }
  }
  if (!t) return null;
  const colsXml = cols.map((c) => `<tableColumn${attrs({ id: c.fid, name: c.name, totalsRowFunction: c.fn !== 'none' ? c.fn : undefined })}>${c.f ? `<calculatedColumnFormula>${esc(c.f)}</calculatedColumnFormula>` : ''}</tableColumn>`).join('');
  return `<table xmlns="${NS}"${head}>${af}<tableColumns count="${cols.length}">${colsXml}</tableColumns>${style}</table>`;
}

// ─────────────── 메모 ───────────────
function commentsXml(u8) {
  const authors = [];
  const list = [];
  let cur = null;
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    if (r.t === R.CAUTHOR) authors.push(rd.str() ?? '');
    else if (r.t === R.COMMENT) { const a = rd.i32(); cur = { a, ref: rangeRef(rd.rfx()), text: '' }; list.push(cur); } else if (r.t === R.CTEXT && cur) { rd.skip(1); cur.text = rd.str() ?? ''; }
  }
  return `<comments xmlns="${NS}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${list.map((c) => `<comment ref="${c.ref}" authorId="${Math.max(0, c.a)}"><text><r><t xml:space="preserve">${esc(c.text)}</t></r></text></comment>`).join('')}</commentList></comments>`;
}

// ─────────────── 피벗 캐시 ───────────────
function pad2(n) { return String(n).padStart(2, '0'); }

/** 캐시 항목 레코드 → XML 항목 */
function cacheItem(t, rd) {
  switch (t) {
    case 20: case 27: return '<m/>';
    case 21: case 28: return `<n v="${numText(rd.f64())}"/>`;
    case 22: case 29: return `<b v="${rd.u8v() ? 1 : 0}"/>`;
    case 23: case 30: return `<e v="${errText(rd.u8v())}"/>`;
    case 24: case 31: return `<s v="${esc(rd.str() ?? '')}"/>`;
    case 25: case 32: {
      const y = rd.u16(); const mo = rd.u16(); const d = rd.u8v(); const h = rd.u8v(); const mi = rd.u8v(); const s = rd.u8v();
      return `<d v="${y}-${pad2(mo)}-${pad2(d)}T${pad2(h)}:${pad2(mi)}:${pad2(s)}"/>`;
    }
    default: return '';
  }
}

/**
 * 피벗 캐시 레코드 (pivotCacheRecords.bin) → xlsx 의 <pivotCacheRecords> XML. 행마다 BrtPCRRecordDt(34) 뒤에 필드별 항목 기록,
 * 또는 BrtPCRRecord(33) 하나에 필드 값을 이어 붙인 형태 (공유 항목이 있는 필드 = u32 번호, 숫자 = f64, 글자 = XLWideString, 날짜 = 8바이트).
 * 형식을 알 수 없으면 null (그러면 원본에서 다시 계산)
 */
function pivotRecordsXml(u8, fields, recordCount) {
  const db = fields.filter((f) => f.database);
  const kinds = db.map((f) => {
    if (f.items.length) return 'x';
    const sf = f.sflags ?? 0;
    const num = sf & 0x40 || !(sf & 8);
    if (sf & 4 && !(sf & 8)) return 'd';
    if (sf & 8 && !(sf & 0x40) && !(sf & 4)) return 's';
    return num ? 'n' : null;
  });
  const out = [];
  let cur = null;
  let count = 0;
  let bad = false;
  const flush = () => { if (cur) { out.push(`<r>${cur.join('')}</r>`); count++; } cur = null; };
  for (const r of records(u8)) {
    if (r.t === 34) { flush(); cur = []; continue; }
    if (r.t === 33) {
      flush();
      const rd = new Rd(u8, r.p, r.e);
      const row = [];
      for (const k of kinds) {
        if (k === 'x') row.push(`<x v="${rd.u32()}"/>`);
        else if (k === 'n') row.push(`<n v="${numText(rd.f64())}"/>`);
        else if (k === 's') row.push(`<s v="${esc(rd.str() ?? '')}"/>`);
        else if (k === 'd') row.push(cacheItem(25, rd));
        else { bad = true; break; }
      }
      if (bad || rd.p !== r.e) return null;
      out.push(`<r>${row.join('')}</r>`);
      count++;
      continue;
    }
    if (!cur) continue;
    if (r.t === 26) cur.push(`<x v="${new Rd(u8, r.p, r.e).u32()}"/>`);
    else if (r.t >= 20 && r.t <= 32) cur.push(cacheItem(r.t, new Rd(u8, r.p, r.e)));
  }
  flush();
  if (!count || (recordCount && count !== recordCount)) return null;
  return `<pivotCacheRecords xmlns="${NS}" count="${count}">${out.join('')}</pivotCacheRecords>`;
}

function cacheArray(rd) {
  const type = rd.u16(); const n = rd.i32();
  const out = [];
  for (let i = 0; i < n && rd.left > 0; i++) {
    if (type === 1) out.push(`<n v="${numText(rd.f64())}"/>`);
    else if (type === 2) out.push(`<s v="${esc(rd.str() ?? '')}"/>`);
    else if (type === 0x10) out.push(`<e v="${errText(rd.u8v())}"/>`);
    else if (type === 0x20) out.push(cacheItem(25, rd));
    else break;
  }
  return out.join('');
}

const GROUP_BY = ['range', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'];
const serialIsoTime = (n) => new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000)).toISOString().slice(0, 19);
/** 캐시 필드 그룹 → <fieldGroup> (엑셀 xlsx 와 같은 모양) */
function groupXml(g) {
  if (!g || typeof g !== 'object') return '';
  const r = g.range;
  let rp = '';
  if (r) {
    const auto = { autoStart: r.fl & 1 ? undefined : 0, autoEnd: r.fl & 2 ? undefined : 0 };
    rp = r.by === 0
      ? `<rangePr${attrs({ ...auto, startNum: numText(r.start), endNum: numText(r.end), groupInterval: numText(r.step) })}/>`
      : `<rangePr${attrs({ ...auto, groupBy: GROUP_BY[r.by] ?? 'months', startDate: serialIsoTime(r.start), endDate: serialIsoTime(r.end), groupInterval: r.step !== 1 ? numText(r.step) : undefined })}/>`;
  }
  const items = g.items.join('');
  const n = (items.match(/<[a-z]/g) ?? []).length;
  return `<fieldGroup${attrs({ par: g.par >= 0 ? g.par : undefined, base: g.base >= 0 ? g.base : undefined })}>${rp}${n ? `<groupItems count="${n}">${items}</groupItems>` : ''}</fieldGroup>`;
}

function pivotCacheXml(u8, env, out = null) {
  let source = '';
  const fields = [];
  let f = null;
  let inShared = false;
  let inGroup = false;
  const pnames = [];
  let recordCount = 0;
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    switch (r.t) {
      case R.PCDEF: rd.skip(3); rd.u8v(); rd.i32(); rd.f64(); rd.u8v(); recordCount = rd.i32(); break;
      case R.PCDSHEETSRC: {
        const isName = rd.u8v(); const builtin = rd.u8v(); const fl = rd.u8v();
        const sheet = fl & 2 ? rd.str() : null;
        if (fl & 1) rd.str();
        if (!isName) source = `<worksheetSource${attrs({ ref: rangeRef(rd.rfx()), sheet: sheet ?? undefined })}/>`;
        else source = `<worksheetSource${attrs({ name: (builtin ? '_xlnm.' : '') + (rd.str() ?? ''), sheet: sheet ?? undefined })}/>`;
        break;
      }
      case R.PCDFIELD: {
        const fl = rd.u16(); const fmt = rd.i32(); rd.i16(); rd.i32(); rd.i32(); const mapCount = rd.i32();
        const name = rd.str() ?? '';
        if (fl & 8) rd.str();
        let fp = null;
        if (fl & 0x100) { fp = rd.p; const cce = rd.i32(); rd.skip(cce); const cb = rd.i32(); rd.skip(cb); }
        if (mapCount > 0 && rd.left >= 4) rd.skip(Math.max(0, rd.i32()));
        f = { name, fmt, fp, items: [], shared: '', group: false, database: !!(fl & 4) };
        fields.push(f);
        break;
      }
      case R.PCDFSITEMS: inShared = true; if (f) f.sflags = rd.u16(); break;
      case R.PCDFSITEMS_END: inShared = false; break;
      case R.PCDFGROUP: inGroup = true; if (f) f.group = { par: rd.i32(), base: rd.i32(), range: null, items: [] }; break;
      case R.PCDFGROUP_END: inGroup = false; break;
      // 그룹 범위 (날짜: 초 · 분 · 시 · 일 · 월 · 분기 · 연, 숫자: 구간)
      case R.PCDFGRANGE: if (f?.group) { const by = rd.u8v(); const fl = rd.u8v(); f.group.range = { by, fl, start: rd.f64(), end: rd.f64(), step: rd.f64() }; } break;
      case R.PCITEM_ARRAY:
        if (f && inShared && !inGroup) f.items.push(cacheArray(rd));
        else if (f?.group && inGroup) f.group.items.push(cacheArray(rd));
        break;
      case R.PCDPNAMES: if (f) f.pnames = []; break;
      // 계산 필드 수식의 이름 목록은 필드마다 따로 (PtgName 번호 = 그 필드 목록의 순서)
      case R.PCDPNAME: (f?.pnames ?? pnames).push(rd.u32()); break;
      default:
        if (f && inShared && !inGroup && r.t >= 20 && r.t <= 32 && r.t !== 26) f.items.push(cacheItem(r.t, rd));
        else if (f?.group && inGroup && r.t >= 20 && r.t <= 32 && r.t !== 26) f.group.items.push(cacheItem(r.t, rd));
    }
  }
  const fieldsXml = fields.map((x) => {
    let formula;
    if (x.fp !== null) {
      const penv = { ...env, memo: new Map(), pnames: (x.pnames ?? pnames).map((i) => fields[i]?.name) }; // 같은 바이트라도 이름 목록이 달라 기억해 두지 않음
      const { text } = decodeFormula(u8, x.fp, penv, { r: 0, c: 0 });
      if (text !== null) formula = text;
    }
    const items = x.items.join('');
    const sf = x.sflags ?? 0;
    const sa = { containsSemiMixedTypes: sf & 1 ? undefined : 0, containsNonDate: sf & 2 ? undefined : 0, containsDate: sf & 4 ? 1 : undefined, containsString: sf & 8 ? undefined : 0, containsBlank: sf & 0x10 ? 1 : undefined, containsNumber: sf & 0x40 ? 1 : undefined, containsInteger: sf & 0x80 ? 1 : undefined };
    return `<cacheField${attrs({ name: x.name, numFmtId: x.fmt >= 0 ? x.fmt : 0, formula, databaseField: x.database ? undefined : 0 })}><sharedItems${attrs(sa)}${items ? ` count="${(items.match(/<[a-z]/g) ?? []).length}">${items}</sharedItems>` : '/>'}${groupXml(x.group)}</cacheField>`;
  }).join('');
  if (out) { out.fields = fields; out.recordCount = recordCount; }
  // 캐시 레코드(엑셀이 저장한 원본)를 읽을 수 있으면 그것으로 (새로 고침 전의 엑셀 화면과 같게), 없으면 열 때 원본에서 계산
  return `<pivotCacheDefinition xmlns="${NS}" xmlns:r="${NS_R}"${out?.records ? '' : ' refreshOnLoad="1"'} recordCount="${recordCount}"><cacheSource type="worksheet">${source}</cacheSource><cacheFields count="${fields.length}">${fieldsXml}</cacheFields></pivotCacheDefinition>`;
}

// ─────────────── 피벗 테이블 ───────────────
const ITEM_TYPES = ['data', 'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank'];
const SUBTOTALS = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'];
const SHOW_AS = ['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index'];
const FIELD_SUBS = [[0x200, 'sumSubtotal'], [0x400, 'countASubtotal'], [0x800, 'avgSubtotal'], [0x1000, 'maxSubtotal'], [0x2000, 'minSubtotal'], [0x4000, 'productSubtotal'], [0x8000, 'countSubtotal'], [0x10000, 'stdDevSubtotal'], [0x20000, 'stdDevPSubtotal'], [0x40000, 'varSubtotal'], [0x80000, 'varPSubtotal']];
const PT_FILTERS = [null, 'count', 'percent', 'sum', 'captionEqual', 'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith', 'captionEndsWith', 'captionNotEndsWith', 'captionContains', 'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan', 'captionLessThanOrEqual', 'captionBetween', 'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan', 'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual', 'valueBetween', 'valueNotBetween'];

function pivotTableXml(u8) {
  let head = {};
  let loc = '';
  const fields = [];
  let fld = null;
  let rows = []; let cols = [];
  const pages = []; const datas = []; const filters = [];
  let style = '';
  let flt = null;
  let depth = 0; // PTFIELD 안의 참조 (자동 정렬 기준)
  // 표시된 행 · 열 항목 (rowItems · colItems): 값이 같은 항목의 순서를 엑셀과 똑같이 하는 데 씀
  const lines = { rowItems: [], colItems: [] };
  let lineList = null;
  let line = null;
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    switch (r.t) {
      case R.PTROWITEMS: lineList = lines.rowItems; break;
      case R.PTCOLITEMS: lineList = lines.colItems; break;
      case R.PTROWITEMS_END: case R.PTCOLITEMS_END: lineList = null; break;
      case R.PTLINE: if (lineList) { const rr = rd.u16(); const t = rd.u16(); rd.i32(); const i = rd.i32(); line = { r: rr, t, i, xs: [] }; lineList.push(line); } break;
      case R.PTLINE_X: if (line) while (rd.left >= 4) line.xs.push(rd.i32()); break;
      case R.PTDEF: {
        const f1 = rd.u32(); const f2 = rd.u32(); const f3 = rd.u32(); const dataAxis = rd.u8v(); rd.u8v(); rd.skip(2); rd.i32(); rd.u16(); rd.skip(2); rd.i32();
        const cacheId = rd.i32(); const name = rd.str() ?? '';
        const dataCaption = f2 & 0x80000 ? rd.str() : null;
        const grandCaption = f2 & 0x100000 ? rd.str() : null;
        const errorCaption = !(f3 & 0x40) ? rd.str() : null;
        const missingCaption = !(f3 & 0x80) ? rd.str() : null;
        if (f2 & 0x200000) rd.str();
        const ptStyle = f2 & 0x400000 ? rd.str() : null;
        if (f2 & 0x800000) rd.str();
        if (f2 & 0x40000000) rd.str();
        const colHeaderCaption = f3 & 0x800 ? rd.str() : null;
        const rowHeaderCaption = f3 & 0x400 ? rd.str() : null;
        head = {
          name, cacheId, dataCaption: dataCaption ?? 'Values', grandTotalCaption: grandCaption ?? undefined, dataOnRows: dataAxis === 1 ? 1 : undefined,
          showError: f2 & 0x200 ? 1 : undefined, errorCaption: errorCaption ?? undefined, showMissing: f2 & 0x400 ? undefined : 0, missingCaption: missingCaption ?? undefined,
          rowGrandTotals: f2 & 0x2000 ? undefined : 0, colGrandTotals: f2 & 0x4000 ? undefined : 0, mergeItem: f2 & 0x40000 ? 1 : undefined,
          preserveFormatting: f2 & 0x80 ? undefined : 0, useAutoFormatting: f2 & 0x100 ? 1 : undefined, enableDrill: f2 & 0x20 ? undefined : 0,
          showDrill: f1 & 0x100000 ? 0 : undefined, showHeaders: f1 & 0x80000000 ? 0 : undefined, indent: (f1 >>> 24) & 0x7f,
          rowHeaderCaption: rowHeaderCaption ?? undefined, colHeaderCaption: colHeaderCaption ?? undefined, _style: ptStyle,
        };
        break;
      }
      case R.PTLOCATION: {
        const rg = rd.rfx(); const fh = rd.i32() - rg.r1; const fd = rd.i32() - rg.r1; const fc = rd.i32() - rg.c1; const rp = rd.i32(); const cp = rd.i32();
        loc = `<location${attrs({ ref: rangeRef(rg), firstHeaderRow: fh, firstDataRow: fd, firstDataCol: fc, rowPageCount: rp || undefined, colPageCount: cp || undefined })}/>`;
        break;
      }
      case R.PTFIELD: {
        const f1 = rd.u32(); const fmt = rd.i32(); const f2 = rd.u32(); const asItems = rd.i32(); const asRank = rd.i32();
        const caption = f1 & 0x20000000 && rd.left >= 4 ? rd.str() : null; // 셀에서 바꾼 필드 이름 (pivotField@name)
        const axis = [null, 'axisRow', 'axisCol', null, 'axisPage'][f1 & 7];
        const a = {
          name: caption ?? undefined, axis: axis ?? undefined, dataField: f1 & 8 ? 1 : undefined, numFmtId: fmt || undefined, compact: f1 & 0x10000000 ? undefined : 0,
          outline: f2 & 0x40 ? undefined : 0, showAll: f2 & 0x20 ? undefined : 0, insertBlankRow: f2 & 0x80 ? 1 : undefined, subtotalTop: f2 & 0x100 ? undefined : 0,
          defaultSubtotal: f1 & 0x100 ? undefined : 0, sortType: f2 & 0x1000 ? (f2 & 0x2000 ? 'ascending' : 'descending') : undefined,
          autoShow: f2 & 0x4000 ? 1 : undefined, topAutoShow: f2 & 0x8000 ? undefined : 0, itemPageCount: undefined, multipleItemSelectionAllowed: f2 & 0x80000 ? 1 : undefined,
        };
        for (const [bit, k] of FIELD_SUBS) if (f1 & bit) a[k] = 1;
        fld = { a, items: [], sort: null, asItems, asRank };
        fields.push(fld);
        depth = 1;
        break;
      }
      case R.PTFIELD_END: depth = 0; break;
      case R.PTFITEM: if (fld) {
        const type = rd.u8v(); const fl = rd.u16(); const x = rd.i32();
        const t = ITEM_TYPES[type] ?? 'data';
        fld.items.push(`<item${attrs({ t: t !== 'data' ? t : undefined, h: fl & 1 ? 1 : undefined, sd: fl & 2 ? 0 : undefined, x: t === 'data' ? x : undefined })}/>`);
      } break;
      case R.PTREFERENCE: if (fld && depth) { fld.sort = { field: rd.i32(), items: [] }; } break;
      case R.PTREFITEM: if (fld?.sort) fld.sort.items.push(rd.i32()); break;
      case R.PTROWFIELDS: { const n = rd.i32(); rows = Array.from({ length: Math.min(n, rd.left >> 2) }, () => rd.i32()); break; }
      case R.PTCOLFIELDS: { const n = rd.i32(); cols = Array.from({ length: Math.min(n, rd.left >> 2) }, () => rd.i32()); break; }
      case R.PTPAGEFIELD: { const fl = rd.i32(); const item = rd.i32(); rd.skip(4); const hasName = rd.u8v() & 1; const name = hasName ? rd.str() : null; pages.push(`<pageField${attrs({ fld: fl, item: item === 0x001000FE || item < 0 ? undefined : item, hier: -1, name: name ?? undefined })}/>`); break; }
      case R.PTDATAFIELD: {
        const f = rd.i32(); const sub = rd.i32(); const as = rd.i32(); const bf = rd.i32(); const bi = rd.i32(); const fmt = rd.i32(); const hasName = rd.u8v();
        const name = hasName === 1 ? rd.str() : null;
        datas.push(`<dataField${attrs({ name: name ?? undefined, fld: f, subtotal: SUBTOTALS[sub] && sub ? SUBTOTALS[sub] : undefined, showDataAs: as ? SHOW_AS[as] : undefined, baseField: bf, baseItem: bi >>> 0 > 0x7fffffff ? 0 : bi, numFmtId: fmt || undefined })}/>`);
        break;
      }
      case R.PTFILTER: {
        const f = rd.i32(); rd.i32(); const type = rd.i32(); rd.skip(4); const id = rd.i32(); const mf = rd.i32(); rd.i32(); const fl = rd.u16();
        const name = fl & 1 ? rd.str() : null; if (fl & 2) rd.str();
        const v1 = fl & 4 ? rd.str() : null; const v2 = fl & 8 ? rd.str() : null;
        flt = { a: { fld: f, type: PT_FILTERS[type] ?? 'count', evalOrder: -1, id, iMeasureFld: mf >= 0 ? mf : undefined, name: name ?? undefined, stringValue1: v1 ?? undefined, stringValue2: v2 ?? undefined }, top: '' };
        filters.push(flt);
        break;
      }
      case R.TOP10: if (flt) { const fl = rd.u8v(); const v = rd.f64(); flt.top = `<autoFilter ref="A1"><filterColumn colId="0"><top10${attrs({ top: fl & 1 ? undefined : 0, percent: fl & 2 ? 1 : undefined, val: v })}/></filterColumn></autoFilter>`; } break;
      case R.TSINFO: {
        const fl = rd.u16(); const name = rd.str();
        style = `<pivotTableStyleInfo${attrs({ name: name ?? undefined, showRowHeaders: fl & 0x10 ? 1 : 0, showColHeaders: fl & 0x20 ? 1 : 0, showRowStripes: fl & 4 ? 1 : 0, showColStripes: fl & 8 ? 1 : 0, showLastColumn: fl & 2 ? 1 : 0 })}/>`;
        break;
      }
      default: break;
    }
  }
  if (!style && head._style) style = `<pivotTableStyleInfo name="${esc(head._style)}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`;
  delete head._style;
  const anyCompact = fields.some((x) => x.a.compact === undefined && x.a.axis);
  const anyOutline = fields.some((x) => x.a.outline === undefined && x.a.axis);
  const fieldsXml = fields.map((x) => {
    const sort = x.sort && x.a.sortType ? `<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="${x.sort.field >>> 0}" count="${x.sort.items.length}" selected="0">${x.sort.items.map((v) => `<x v="${v}"/>`).join('')}</reference></references></pivotArea></autoSortScope>` : '';
    return `<pivotField${attrs(x.a)}>${x.items.length ? `<items count="${x.items.length}">${x.items.join('')}</items>` : ''}${sort}</pivotField>`;
  }).join('');
  const fl = (tag, list) => (list.length ? `<${tag} count="${list.length}">${list.map((v) => `<field x="${v}"/>`).join('')}</${tag}>` : '');
  const ITEM_T = ['data', 'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank'];
  const lineXml = (tag, list) => (list.length ? `<${tag} count="${list.length}">${list.map((l) => `<i${attrs({ t: l.t ? ITEM_T[l.t] ?? 'data' : undefined, r: l.r || undefined, i: l.i || undefined })}>${l.xs.map((v) => `<x v="${v}"/>`).join('')}</i>`).join('')}</${tag}>` : '');
  return `<pivotTableDefinition xmlns="${NS}"${attrs({ ...head, compact: anyCompact ? undefined : 0, compactData: anyCompact ? undefined : 0, outline: anyOutline ? 1 : undefined, outlineData: anyOutline ? 1 : undefined })}>${loc}<pivotFields count="${fields.length}">${fieldsXml}</pivotFields>${fl('rowFields', rows)}${lineXml('rowItems', lines.rowItems)}${fl('colFields', cols)}${lineXml('colItems', lines.colItems)}${pages.length ? `<pageFields count="${pages.length}">${pages.join('')}</pageFields>` : ''}${datas.length ? `<dataFields count="${datas.length}">${datas.join('')}</dataFields>` : ''}${style}${filters.length ? `<filters count="${filters.length}">${filters.map((x) => `<filter${attrs(x.a)}>${x.top}</filter>`).join('')}</filters>` : ''}</pivotTableDefinition>`;
}

// ─────────────── 슬라이서 ───────────────
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

function slicerCacheXml(u8) {
  let name = ''; let source = '';
  const pts = []; const items = []; let tab = {};
  for (const r of records(u8)) {
    const rd = new Rd(u8, r.p, r.e);
    if (r.t === R.SLC_DEF) { name = rd.str() ?? ''; source = rd.str() ?? ''; } else if (r.t === R.SLC_PIVOTS) {
      const n = rd.i32();
      for (let i = 0; i < n && rd.left >= 8; i++) { const tabId = rd.i32(); pts.push({ tabId, name: rd.str() ?? '' }); }
    } else if (r.t === R.SLC_TABULAR) {
      const sort = rd.u32(); const cacheId = rd.u32();
      tab = { pivotCacheId: cacheId, sortOrder: sort === 1 ? 'descending' : undefined };
    } else if (r.t === R.SLC_ITEMS) {
      const n = rd.i32();
      for (let i = 0; i < n && rd.left >= 5; i++) { const x = rd.i32(); const fl = rd.u8v(); items.push(`<i x="${x}"${fl & 1 ? ' s="1"' : ''}${fl & 2 ? ' nd="1"' : ''}/>`); }
    }
  }
  return `<slicerCacheDefinition xmlns="${NS_X14}"${attrs({ name, sourceName: source })}>${pts.length ? `<pivotTables>${pts.map((p) => `<pivotTable tabId="${p.tabId}" name="${esc(p.name)}"/>`).join('')}</pivotTables>` : ''}<data><tabular${attrs(tab)}><items count="${items.length}">${items.join('')}</items></tabular></data></slicerCacheDefinition>`;
}

function slicersXml(u8) {
  const out = [];
  for (const r of records(u8)) {
    if (r.t !== R.SLICER) continue;
    const rd = new Rd(u8, r.p, r.e);
    const f0 = rd.u8v(); const f1 = rd.u8v(); rd.skip(3);
    const columns = rd.u32(); const start = rd.u32(); const rowHeight = rd.u32();
    const name = rd.str() ?? ''; const cache = rd.str() ?? ''; const caption = rd.str() ?? name; const style = f0 & 4 ? rd.str() : null;
    out.push(`<slicer${attrs({ name, cache, caption, startItem: start || undefined, columnCount: columns > 1 ? columns : undefined, showCaption: f0 & 1 ? undefined : 0, style: style ?? undefined, lockedPosition: f1 & 1 ? 1 : undefined, rowHeight })}/>`);
  }
  return `<slicers xmlns="${NS_X14}">${out.join('')}</slicers>`;
}

// ─────────────── 시트 ───────────────
const CFVO_TYPES = [null, 'num', 'min', 'max', 'percent', 'percentile', 'formula', 'min', 'max'];
const ICON_SETS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes'];
const CF_TIME = ['today', 'yesterday', 'last7Days', 'thisWeek', 'lastWeek', 'lastMonth', 'tomorrow', 'nextWeek', 'nextMonth', 'thisMonth'];
const CF_OPS = [null, 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
const CF_TEXT = ['containsText', 'notContainsText', 'beginsWith', 'endsWith'];
const CF_TEXT_OP = ['containsText', 'notContains', 'beginsWith', 'endsWith'];

/** 조건부 서식 규칙 머리 → xlsx 속성 */
function cfRuleAttrs(type, sub, op, flags) {
  const a = {};
  if (type === 1) { a.type = 'cellIs'; a.operator = CF_OPS[op] ?? 'equal'; } else if (type === 2) {
    switch (sub) {
      case 1: a.type = 'expression'; break;
      case 7: a.type = 'uniqueValues'; break;
      case 8: a.type = CF_TEXT[op] ?? 'containsText'; a.operator = CF_TEXT_OP[op] ?? 'containsText'; break;
      case 9: a.type = 'containsBlanks'; break;
      case 10: a.type = 'notContainsBlanks'; break;
      case 11: a.type = 'containsErrors'; break;
      case 12: a.type = 'notContainsErrors'; break;
      case 25: case 26: case 29: case 30: a.type = 'aboveAverage'; if (sub === 26 || sub === 30) a.aboveAverage = 0; if (sub >= 29) a.equalAverage = 1; if (op) a.stdDev = op; break;
      case 27: a.type = 'duplicateValues'; break;
      default: if (sub >= 15 && sub <= 24) { a.type = 'timePeriod'; a.timePeriod = CF_TIME[op] ?? 'today'; } else a.type = 'expression';
    }
  } else if (type === 3) a.type = 'colorScale';
  else if (type === 4) a.type = 'dataBar';
  else if (type === 5) { a.type = 'top10'; a.rank = op; if (flags & 8) a.bottom = 1; if (flags & 0x10) a.percent = 1; } else if (type === 6) a.type = 'iconSet';
  if (flags & 2) a.stopIfTrue = 1;
  return a;
}

/** 시트 머리(셀 데이터 밖의 레코드) → { xml, dataStart, dataEnd } */
function sheetHeadXml(u8, env) {
  const parts = { pr: '', dim: '', views: [], fmt: '', cols: [], merges: [], cf: [], links: [], af: '', margins: '', drawing: '', legacy: '', tables: [], protect: '' };
  let dataStart = -1; let dataEnd = -1;
  let p = 0;
  const n = u8.length;
  let view = null;
  let cf = null; let rule = null; let sub = null;
  let afCol = null;
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  while (p < n) {
    let t = u8[p++];
    if (t & 0x80) t = (t & 0x7f) | ((u8[p++] & 0x7f) << 7);
    let sz = 0;
    for (let i = 0, sh = 0; i < 4; i++, sh += 7) { const x = u8[p++]; sz |= (x & 0x7f) << sh; if (!(x & 0x80)) break; }
    const e = p + sz;
    if (t === R.SHEETDATA) {
      dataStart = e;
      // 셀 데이터는 머리만 훑어 건너뜀 (행 읽기는 rowsOf 가 따로)
      let q = e;
      for (;;) {
        if (q >= n) { dataEnd = n; p = n; break; }
        const q0 = q;
        let tt = u8[q++];
        if (tt & 0x80) tt = (tt & 0x7f) | ((u8[q++] & 0x7f) << 7);
        let s2 = 0;
        for (let i = 0, sh = 0; i < 4; i++, sh += 7) { const x = u8[q++]; s2 |= (x & 0x7f) << sh; if (!(x & 0x80)) break; }
        if (tt === R.SHEETDATA_END) { dataEnd = q0; p = q + s2; break; }
        q += s2;
      }
      continue;
    }
    const rd = new Rd(u8, p, e);
    switch (t) {
      case R.WSPR: {
        const f1 = rd.u16(); rd.u8v(); const tab = color(rd);
        const outlinePr = (f1 & 0x40) === 0 || (f1 & 0x80) === 0 ? `<outlinePr${attrs({ summaryBelow: f1 & 0x40 ? undefined : 0, summaryRight: f1 & 0x80 ? undefined : 0 })}/>` : '';
        const pageSetUp = f1 & 0x100 ? '<pageSetUpPr fitToPage="1"/>' : '';
        parts.pr = `<sheetPr>${tab.auto ? '' : colorXml('tabColor', tab)}${outlinePr}${pageSetUp}</sheetPr>`;
        break;
      }
      case R.DIM: { const rg = rd.rfx(); parts.dim = `<dimension ref="${rangeRef(rg)}"/>`; break; }
      case R.WSVIEW: {
        const fl = rd.u16(); rd.i32(); const tr = rd.i32(); const tc = rd.i32(); rd.skip(4); const zoom = rd.u16();
        view = { a: { showGridLines: fl & 4 ? undefined : 0, showRowColHeaders: fl & 8 ? undefined : 0, showZeros: fl & 0x10 ? undefined : 0, rightToLeft: fl & 0x20 ? 1 : undefined, tabSelected: fl & 0x40 ? 1 : undefined, zoomScale: zoom && zoom !== 100 ? zoom : undefined, topLeftCell: tr || tc ? cellRef(tr, tc) : undefined, workbookViewId: 0 }, pane: '' };
        parts.views.push(view);
        break;
      }
      case R.PANE: if (view) {
        const xs = rd.f64(); const ys = rd.f64(); const r2 = rd.i32(); const c2 = rd.i32(); const act = rd.i32(); const fl = rd.u8v();
        const frozen = fl & 1;
        view.pane = `<pane${attrs({ xSplit: xs || undefined, ySplit: ys || undefined, topLeftCell: cellRef(Math.max(0, r2), Math.max(0, c2)), activePane: ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'][act] ?? undefined, state: frozen ? 'frozen' : 'split' })}/>`;
      } break;
      case R.WSFMT: {
        const dw = rd.i32(); const base = rd.u16(); const h = rd.u16(); const fl = rd.u16();
        parts.fmt = `<sheetFormatPr${attrs({ baseColWidth: base !== 8 ? base : undefined, defaultColWidth: dw >= 0 ? dw / 256 : undefined, defaultRowHeight: h / 20, customHeight: fl & 1 ? 1 : undefined, zeroHeight: fl & 2 ? 1 : undefined })}/>`;
        break;
      }
      case R.COL: {
        const c1 = rd.i32(); const c2 = rd.i32(); const w = rd.i32(); const xf = rd.i32(); const fl = rd.u16();
        parts.cols.push(`<col${attrs({ min: c1 + 1, max: c2 + 1, width: w / 256, style: xf || undefined, hidden: fl & 1 ? 1 : undefined, customWidth: fl & 2 ? 1 : undefined, bestFit: fl & 4 ? 1 : undefined, outlineLevel: (fl >> 8) & 7 || undefined, collapsed: fl & 0x1000 ? 1 : undefined })}/>`);
        break;
      }
      case R.MERGE: parts.merges.push(`<mergeCell ref="${rangeRef(rd.rfx())}"/>`); break;
      case R.HLINK: {
        const rg = rd.rfx(); const rid = rd.str(); const loc = rd.str(); const tip = rd.str(); const disp = rd.str();
        parts.links.push(`<hyperlink${attrs({ ref: rangeRef(rg), 'r:id': rid || undefined, location: loc || undefined, tooltip: tip || undefined, display: disp || undefined })}/>`);
        break;
      }
      case R.AFILTER: parts.af = { ref: rangeRef(rd.rfx()), cols: [] }; break;
      case R.FILTERCOL: if (parts.af) { afCol = { id: rd.i32(), fl: rd.u16(), vals: [], blank: false }; parts.af.cols.push(afCol); } break;
      case R.FILTERS: if (afCol) afCol.blank = rd.i32() !== 0; break;
      case R.FILTER: if (afCol) afCol.vals.push(rd.str() ?? ''); break;
      case R.CF: {
        rd.skip(8);
        const cnt = rd.i32();
        const ranges = [];
        for (let i = 0; i < cnt && rd.left >= 16; i++) ranges.push(rd.rfx());
        cf = { ranges, rules: [] };
        parts.cf.push(cf);
        break;
      }
      case R.CFRULE: if (cf) {
        const type = rd.i32(); const st = rd.i32(); const dxf = rd.i32(); const pri = rd.i32(); const op = rd.i32(); rd.skip(8); const fl = rd.u16();
        const s1 = rd.i32(); const s2 = rd.i32(); const s3 = rd.i32(); const text = rd.str();
        const a = { ...cfRuleAttrs(type, st, op, fl), dxfId: dxf >= 0 ? dxf : undefined, priority: pri };
        if (text) a.text = text;
        const base = { r: cf.ranges[0]?.r1 ?? 0, c: cf.ranges[0]?.c1 ?? 0 };
        const fs = [];
        for (const s of [s1, s2, s3]) {
          if (s <= 0 || rd.left < 8) break;
          const { text: ft, next } = decodeFormula(u8, rd.p, env, base, { relN: false });
          rd.p = next;
          if (ft !== null) fs.push(ft);
        }
        rule = { a, fs, inner: '' };
        cf.rules.push(rule);
        break;
      }
      case R.COLORSCALE: sub = { tag: 'colorScale', a: {}, vos: [], colors: [] }; break;
      case R.DATABAR: sub = { tag: 'dataBar', a: { minLength: rd.u8v(), maxLength: rd.u8v(), showValue: rd.u8v() ? undefined : 0 }, vos: [], colors: [] }; break;
      case R.ICONSET: { const set = rd.i32(); const fl = rd.u16(); sub = { tag: 'iconSet', a: { iconSet: ICON_SETS[set] ?? '3TrafficLights1', showValue: fl & 2 ? 0 : undefined, reverse: fl & 4 ? 1 : undefined }, vos: [], colors: [] }; break; }
      case R.CFVO: if (sub) {
        const ty = rd.i32(); const v = rd.f64(); rd.skip(4); const hasF = rd.i32();
        let val = [2, 3, 7, 8].includes(ty) ? undefined : numText(v);
        if (ty === 6 && hasF > 0 && rd.left >= 8) { const { text } = decodeFormula(u8, rd.p, env, { r: cf?.ranges[0]?.r1 ?? 0, c: cf?.ranges[0]?.c1 ?? 0 }); if (text !== null) val = text; }
        sub.vos.push(`<cfvo${attrs({ type: CFVO_TYPES[ty] ?? 'num', val })}/>`);
      } break;
      case R.CFCOLOR: if (sub) sub.colors.push(colorXml('color', color(rd))); break;
      case R.CFRULE_END:
        if (rule && sub) rule.inner = `<${sub.tag}${attrs(sub.a)}>${sub.vos.join('')}${sub.colors.join('')}</${sub.tag}>`;
        sub = null; rule = null; break;
      case R.MARGINS: {
        const [l, rr, tp, b, h, f] = [rd.f64(), rd.f64(), rd.f64(), rd.f64(), rd.f64(), rd.f64()];
        parts.margins = `<pageMargins left="${l}" right="${rr}" top="${tp}" bottom="${b}" header="${h}" footer="${f}"/>`;
        break;
      }
      case R.DRAWING: parts.drawing = `<drawing r:id="${esc(rd.str() ?? '')}"/>`; break;
      case R.LEGACY: parts.legacy = `<legacyDrawing r:id="${esc(rd.str() ?? '')}"/>`; break;
      case R.TABLEPART: parts.tables.push(`<tablePart r:id="${esc(rd.str() ?? '')}"/>`); break;
      case R.PROTECT: parts.protect = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'; break;
      default: break;
    }
    p = e;
  }
  const af = parts.af ? `<autoFilter ref="${parts.af.ref}">${parts.af.cols.map((c) => `<filterColumn colId="${c.id}"${c.fl & 1 ? ' hiddenButton="1"' : ''}${c.fl & 2 ? ' showButton="0"' : ''}>${c.vals.length || c.blank ? `<filters${c.blank ? ' blank="1"' : ''}>${c.vals.map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters>` : ''}</filterColumn>`).join('')}</autoFilter>` : '';
  const cfXml = parts.cf.map((x) => `<conditionalFormatting sqref="${x.ranges.map(rangeRef).join(' ')}">${x.rules.map((r) => `<cfRule${attrs(r.a)}>${r.inner}${r.fs.map((f) => `<formula>${esc(f)}</formula>`).join('')}</cfRule>`).join('')}</conditionalFormatting>`).join('');
  const views = parts.views.length ? `<sheetViews>${parts.views.map((v) => `<sheetView${attrs(v.a)}>${v.pane}</sheetView>`).join('')}</sheetViews>` : '';
  const xml = `<worksheet xmlns="${NS}" xmlns:r="${NS_R}">${parts.pr}${parts.dim}${views}${parts.fmt}${parts.cols.length ? `<cols>${parts.cols.join('')}</cols>` : ''}<sheetData/>${parts.protect}${af}${parts.merges.length ? `<mergeCells count="${parts.merges.length}">${parts.merges.join('')}</mergeCells>` : ''}${cfXml}${parts.links.length ? `<hyperlinks>${parts.links.join('')}</hyperlinks>` : ''}${parts.margins}${parts.drawing}${parts.legacy}${parts.tables.length ? `<tableParts count="${parts.tables.length}">${parts.tables.join('')}</tableParts>` : ''}</worksheet>`;
  return { xml, dataStart, dataEnd };
}

/**
 * 셀 데이터 → readSheet 행 모양 { attrs, cells: [{ cc, attrs: { s, t }, v, f, fa }] } 을 차례로
 * 공유 수식 · 배열 수식은 기준 셀 바로 뒤의 레코드(BrtShrFmla · BrtArrFmla)에서 읽음
 */
function* rowsOf(u8, start, end, env, warn) {
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const shared = new Map(); // 'r,c' → { p (수식 위치), ref }
  const arrayAnchors = new Set(); // 여러 칸 배열 수식의 기준 칸 'r,c'
  const senv = { ...env, shared: null, memo: null };
  const sharedFn = (row, col, base) => {
    const s = shared.get(`${row},${col}`);
    if (!s) return null;
    const { text } = decodeFormula(u8, s.p, senv, base, { relN: false });
    return text;
  };
  const fenv = { ...env, shared: sharedFn };
  let p = start;
  let row = null;
  let col = -1;
  let pend = null; // 공유/배열 수식 기준 셀 (다음 레코드를 보고 결정)
  const resolvePending = (t, q, qe) => {
    const c = pend;
    pend = null;
    if (t === R.SHRFMLA) {
      const rd = new Rd(u8, q, qe);
      const rg = rd.rfx();
      shared.set(`${c.exp.row},${c.exp.col}`, { p: rd.p, ref: rg });
      if (c.done) return; // 기준 칸이 수식 전체를 가진 경우: 공유 수식만 등록
      const t2 = sharedFn(c.exp.row, c.exp.col, c.base);
      if (t2 !== null) { c.cell.f = t2; c.cell.fa = {}; } else warn();
    } else if (c.done) {
      // 공유 수식 기록 없이 끝난 보통 수식
    } else if (t === R.ARRAY) {
      const rd = new Rd(u8, q, qe);
      const rg = rd.rfx();
      rd.skip(1);
      const { text } = decodeFormula(u8, rd.p, fenv, c.base);
      arrayAnchors.add(`${c.exp.row},${c.exp.col}`);
      if (text !== null) { c.cell.f = text; c.cell.fa = { t: 'array', ref: rangeRef(rg) }; } else warn();
    } else warn();
  };
  while (p < end) {
    let t = u8[p++];
    if (t & 0x80) t = (t & 0x7f) | ((u8[p++] & 0x7f) << 7);
    let sz = 0;
    for (let i = 0, sh = 0; i < 4; i++, sh += 7) { const x = u8[p++]; sz |= (x & 0x7f) << sh; if (!(x & 0x80)) break; }
    const q = p;
    const e = p + sz;
    p = e;
    if (pend) resolvePending(t, q, e);
    if (t === R.ROW) {
      if (row) yield row;
      const rw = dv.getInt32(q, true);
      const xf = dv.getInt32(q + 4, true);
      const h = dv.getUint16(q + 8, true);
      const fl = dv.getUint16(q + 10, true);
      const a = { r: rw + 1, ht: h / 20 };
      if (fl & 0x2000) a.customHeight = '1';
      if (fl & 0x1000) a.hidden = '1';
      if ((fl >> 8) & 7) a.outlineLevel = String((fl >> 8) & 7);
      if (fl & 0x800) a.collapsed = '1';
      if (fl & 0x4000) { a.customFormat = '1'; a.s = xf; }
      row = { attrs: a, cells: [] };
      col = -1;
      continue;
    }
    if (t > 18 && t !== R.RSTR && t !== R.MRSTR) continue;
    if (!row) continue;
    let o = q;
    let cc;
    if (t >= 12 && t <= 18 || t === R.MRSTR) cc = col + 1;
    else { cc = dv.getInt32(o, true); o += 4; }
    col = cc;
    const s = dv.getUint32(o, true) & 0xffffff;
    o += 4;
    const cell = { cc, attrs: { s }, v: null, f: null, fa: null, is: null };
    switch (t) {
      case R.BLANK: case R.MBLANK: break;
      case R.RK: case R.MRK: {
        const rk = dv.getInt32(o, true);
        let v;
        if (rk & 2) v = rk >> 2;
        else { RK.setUint32(0, 0, true); RK.setUint32(4, rk & 0xfffffffc, true); v = RK.getFloat64(0, true); }
        cell.v = rk & 1 ? v / 100 : v;
        break;
      }
      case R.ERROR: case R.MERR: cell.attrs.t = 'e'; cell.v = errText(u8[o]); break;
      case R.BOOL: case R.MBOOL: cell.attrs.t = 'b'; cell.v = u8[o] ? '1' : '0'; break;
      case R.REAL: case R.MREAL: cell.v = dv.getFloat64(o, true); break;
      case R.ISST: case R.MISST: cell.attrs.t = 's'; cell.v = dv.getUint32(o, true); break;
      case R.ST: case R.MST: { const n = dv.getInt32(o, true); cell.attrs.t = 'str'; cell.v = utf16(u8, o + 4, Math.max(0, n)); break; }
      case R.RSTR: case R.MRSTR: { const n = dv.getInt32(o + 1, true); cell.attrs.t = 'str'; cell.v = utf16(u8, o + 5, Math.max(0, n)); break; }
      case R.FSTR: case R.FNUM: case R.FBOOL: case R.FERR: {
        let fp;
        if (t === R.FNUM) { cell.v = dv.getFloat64(o, true); fp = o + 8; } else if (t === R.FSTR) { const n = dv.getInt32(o, true); cell.attrs.t = 'str'; cell.v = utf16(u8, o + 4, Math.max(0, n)); fp = o + 4 + Math.max(0, n) * 2; } else if (t === R.FBOOL) { cell.attrs.t = 'b'; cell.v = u8[o] ? '1' : '0'; fp = o + 1; } else { cell.attrs.t = 'e'; cell.v = errText(u8[o]); fp = o + 1; }
        fp += 2;
        const base = { r: row.attrs.r - 1, c: cc };
        const res = decodeFormula(u8, fp, fenv, base);
        // 수식 전체를 가진 칸 뒤에 공유 수식 기록(BrtShrFmla)이 올 수 있음 → 다음 기록에서 등록
        if (res.text !== null) { cell.f = res.text; cell.fa = {}; pend = { cell, exp: { row: base.r, col: cc }, base, done: true }; } // 공유 수식의 기준 칸이 이 칸이 아니어도 (원래 왼쪽 위 칸을 지운 경우) 바로 뒤의 BrtShrFmla 가 그 기준으로 등록됨
        else if (res.exp && arrayAnchors.has(`${res.exp.row},${res.exp.col}`)) { /* 여러 칸 배열 수식의 나머지 칸: 값만 */ }
        else if (res.exp && !shared.has(`${res.exp.row},${res.exp.col}`)) pend = { cell, exp: res.exp, base };
        else warn();
        break;
      }
      default: break;
    }
    row.cells.push(cell);
  }
  if (pend) pend = null;
  if (row) yield row;
}

// ─────────────── 전체 변환 ───────────────
const relsPath = (path) => { const i = path.lastIndexOf('/'); return `${path.slice(0, i + 1)}_rels/${path.slice(i + 1)}.rels`; };
const RELS_RE = /<Relationship\b([^>]*)\/?>/g;

function relsOf(files, path) {
  const u8 = files[relsPath(path)];
  if (!u8) return [];
  const xml = new TextDecoder().decode(u8);
  const base = path.slice(0, path.lastIndexOf('/') + 1);
  const out = [];
  for (const m of xml.matchAll(RELS_RE)) {
    const at = {};
    for (const a of m[1].matchAll(/(\w+)="([^"]*)"/g)) at[a[1]] = a[2];
    if (at.TargetMode === 'External') continue;
    out.push({ id: at.Id, type: at.Type?.split('/').pop(), target: resolvePath(base, at.Target ?? '') });
  }
  return out;
}

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/');
  const out = [];
  for (const s of parts) { if (s === '..') out.pop(); else if (s !== '.') out.push(s); }
  return out.join('/');
}

/** 파일 목록이 xlsb 인지 */
export function isXlsb(files) {
  return Object.keys(files).some((k) => /^xl\/workbook\.bin$/i.test(k));
}

/**
 * xlsb 부분 파일을 제자리에서 xlsx XML 로 바꿈 (진행 상황을 yield)
 * files.__xlsb = { strings, rows: Map(시트 경로 → 행 생성기 함수), unsupported } (열거되지 않는 속성)
 */
export function* convertXlsb(files) {
  const wbPath = Object.keys(files).find((k) => /^xl\/workbook\.bin$/i.test(k));
  const wbU8 = files[wbPath];
  const wbi = readWorkbook(wbU8);
  const rels = relsOf(files, wbPath);
  const byType = (t) => rels.filter((r) => r.type === t && files[r.target]);
  const info = { strings: [], rows: new Map(), unsupported: 0 };
  Object.defineProperty(files, '__xlsb', { value: info, enumerable: false });
  // 표 (구조적 참조용: 수식보다 먼저)
  const tables = new Map();
  for (const k of Object.keys(files)) {
    if (!/^xl\/tables\/[^/]+\.bin$/i.test(k)) continue;
    const t = scanTable(files[k]);
    if (t) tables.set(t.id, t);
  }
  const env = { sheets: wbi.sheets.map((s) => s.name), names: wbi.names, xti: wbi.xti, tables, shared: null, memo: new Map() };
  put(files, wbPath, workbookXml(wbi, env, wbU8));
  yield { p: 0.08, msg: '스타일 읽는 중' };
  for (const r of byType('styles')) put(files, r.target, stylesXml(files[r.target]));
  for (const r of byType('sharedStrings')) { info.strings = readStrings(files[r.target]); delete files[r.target]; }
  for (const k of Object.keys(files)) {
    if (/^xl\/tables\/[^/]+\.bin$/i.test(k)) { const x = tableXml(files[k], env); if (x) put(files, k, x); } else if (/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.bin$/i.test(k)) {
      // 캐시 레코드도 xlsx 형식으로 (엑셀이 저장한 원본 = 새로 고치기 전 피벗 결과)
      const info0 = {};
      pivotCacheXml(files[k], env, info0);
      const rec = relsOf(files, k).find((x) => x.type === 'pivotCacheRecords');
      const recXml = rec && files[rec.target] && files[rec.target].length < (24 << 20) ? pivotRecordsXml(files[rec.target], info0.fields, info0.recordCount) : null;
      if (recXml) { put(files, rec.target, recXml); files[rec.target].__xml = true; }
      put(files, k, pivotCacheXml(files[k], env, { records: !!recXml }));
    }
    else if (/^xl\/pivotTables\/[^/]+\.bin$/i.test(k)) put(files, k, pivotTableXml(files[k]));
    else if (/^xl\/slicerCaches\/[^/]+\.bin$/i.test(k)) put(files, k, slicerCacheXml(files[k]));
    else if (/^xl\/slicers\/[^/]+\.bin$/i.test(k)) put(files, k, slicersXml(files[k]));
    else if (/^xl\/comments[^/]*\.bin$/i.test(k)) put(files, k, commentsXml(files[k]));
    else if (/^xl\/(calcChain|metadata)\.bin$/i.test(k) || (/pivotCacheRecords/i.test(k) && !files[k].__xml)) delete files[k];
  }
  yield { p: 0.1, msg: '시트 준비 중' };
  for (const r of byType('worksheet')) {
    const u8 = files[r.target];
    if (!(u8 instanceof Uint8Array) || !/\.bin$/i.test(r.target)) continue;
    const head = sheetHeadXml(u8, env);
    put(files, r.target, head.xml);
    if (head.dataStart >= 0) info.rows.set(r.target, () => rowsOf(u8, head.dataStart, head.dataEnd, env, () => { info.unsupported++; }));
    yield { p: 0.1, msg: '시트 준비 중' };
  }
}
