// 수식 ↔ 엑셀 파일 형식 변환 (DOM 없음)
// 파일에는 엑셀 2007 이후 함수에 _xlfn.(FILTER·SORT 는 _xlfn._xlws.), LAMBDA/LET 변수에 _xlpm., 함수 이름만 쓴 인수에 _xleta. 가 붙고,
// '@' 는 _xlfn.SINGLE(), 'A1#' 은 _xlfn.ANCHORARRAY(A1) 로 저장됨
import { parse, FUNCS, mayReturnArray } from './formula.js';
import { canonicalRef } from './tables.js';

/** 엑셀 2007 파일 형식에 원래 있던 함수 (접두사 없음) */
const BASE = new Set(`ABS ACCRINT ACCRINTM ACOS ACOSH ADDRESS AMORDEGRC AMORLINC AND AREAS ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE
AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT BESSELI BESSELJ BESSELK BESSELY BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT BINOMDIST CEILING CELL
CHAR CHIDIST CHIINV CHITEST CHOOSE CLEAN CODE COLUMN COLUMNS COMBIN COMPLEX CONCATENATE CONFIDENCE CONVERT CORREL COS COSH COUNT COUNTA
COUNTBLANK COUNTIF COUNTIFS COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR CRITBINOM CUBEKPIMEMBER CUBEMEMBER
CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS360 DB DCOUNT
DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR DPRODUCT DSTDEV DSTDEVP DSUM DURATION
DVAR DVARP ECMA.CEILING EDATE EFFECT EOMONTH ERF ERFC ERROR.TYPE EVEN EXACT EXP EXPONDIST FACT FACTDOUBLE FALSE FDIST FIND FINDB FINV
FISHER FISHERINV FIXED FLOOR FORECAST FREQUENCY FTEST FV FVSCHEDULE GAMMADIST GAMMAINV GAMMALN GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH
HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR HYPERLINK HYPGEOMDIST IF IFERROR IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP
IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSIN IMSQRT IMSUB IMSUM INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR
ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISO.CEILING ISODD ISPMT ISREF ISTEXT JIS KURT LARGE LCM LEFT LEFTB LEN LENB LINEST LN
LOG LOG10 LOGEST LOGINV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINUTE MINVERSE MIRR MMULT
MOD MODE MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER
NPV OCT2BIN OCT2DEC OCT2HEX ODD ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PEARSON PERCENTILE PERCENTRANK PERMUT PHONETIC PI PMT
POISSON POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER PV QUARTILE QUOTIENT RADIANS RAND RANDBETWEEN RANK RATE RECEIVED
REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS RSQ RTD SEARCH SEARCHB SECOND SERIESSUM SIGN SIN SINH SKEW SLN
SLOPE SMALL SQRT SQRTPI STANDARDIZE STDEV STDEVA STDEVP STDEVPA STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2
SUMX2PY2 SUMXMY2 SYD T TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TIME TIMEVALUE TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUE
TRUNC TTEST TYPE UPPER VALUE VAR VARA VARP VARPA VDB VLOOKUP WEEKDAY WEEKNUM WEIBULL WORKDAY WORKDAY.INTL XIRR XNPV YEAR YEARFRAC YIELD
YIELDDISC YIELDMAT ZTEST NUMBERSTRING DATESTRING WON`.split(/\s+/));
const XLWS = new Set(['FILTER', 'SORT']);

/** 파일에 쓸 함수 이름 */
export function fileFuncName(name) {
  if (BASE.has(name) || !FUNCS[name]) return name;
  return XLWS.has(name) ? `_xlfn._xlws.${name}` : `_xlfn.${name}`;
}

const fmtNum = (v) => {
  if (Number.isInteger(v) && Math.abs(v) < 1e21) return String(v);
  const s = String(v);
  return s.replace('e+', 'E+').replace('e-', 'E-');
};
const quoteStr = (s) => `"${String(s).replace(/"/g, '""')}"`;
const errCode = (e) => (e && typeof e === 'object' && e.code ? e.code : String(e));
const constText = (v) => (typeof v === 'number' ? fmtNum(v) : typeof v === 'string' ? quoteStr(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : errCode(v));

/**
 * 값 하나를 받는 인수 (범위를 주면 옛 형식 수식에서 암시적 교차: SUMIF 의 조건 B145:B147 → 수식 행의 B145).
 * 범위를 받는 인수는 제외 (SUMIF 의 range · sum_range 등)
 */
const SCALAR_ARGS = {
  SUMIF: [1], COUNTIF: [1], AVERAGEIF: [1],
  VLOOKUP: [0, 2, 3], HLOOKUP: [0, 2, 3], MATCH: [0, 2], XLOOKUP: [0, 4, 5], XMATCH: [0, 2, 3], LOOKUP: [0],
  INDEX: [1, 2, 3], OFFSET: [1, 2, 3, 4], LARGE: [1], SMALL: [1], PERCENTILE: [1], QUARTILE: [1], RANK: [0, 2],
};
function scalarArg(name, i) {
  const pos = SCALAR_ARGS[name];
  if (pos) return pos.includes(i);
  // 조건 쌍: SUMIFS · AVERAGEIFS · MAXIFS · MINIFS (범위, 조건 …) 의 조건 자리, COUNTIFS 는 홀수 자리
  if (name === 'COUNTIFS') return i % 2 === 1;
  if (name === 'SUMIFS' || name === 'AVERAGEIFS' || name === 'MAXIFS' || name === 'MINIFS') return i >= 2 && i % 2 === 0;
  return false;
}

/** LET · LAMBDA 가 만드는 지역 이름 */
function boundNames(node) {
  if (node.type !== 'func') return [];
  if (node.name === 'LET') return node.args.filter((_, i) => i % 2 === 0 && i < node.args.length - 1).filter((a) => a.type === 'name').map((a) => a.v.toLowerCase());
  if (node.name === 'LAMBDA') return node.args.slice(0, -1).filter((a) => a.type === 'name').map((a) => a.v.toLowerCase());
  return [];
}

/**
 * AST → 수식 텍스트 (맨 앞 '=' 없음)
 * mode 'file': 엑셀 파일 형식 (dynamic=true 면 '@' → SINGLE)
 * mode 'app' : 앱 형식 (SINGLE → '@', ANCHORARRAY → '#')
 */
function print(node, src, opt, scope = new Set(), scalar = false) {
  const P = (n, sc = false, sp = scope) => print(n, src, opt, sp, sc);
  const file = opt.mode === 'file';
  const at = (n, text) => {
    // 옛 형식(동적 배열 이전) 수식의 암시적 교차를 '@' 로 표시
    if (!file && opt.legacy && scalar && mayReturnArray(n)) {
      if (n.type === 'name' && !opt.nameMulti?.(n.v)) return text;
      return `@${text}`;
    }
    return text;
  };
  switch (node.type) {
    case 'num': return fmtNum(node.v);
    case 'str': return quoteStr(node.v);
    case 'bool': return node.v ? 'TRUE' : 'FALSE';
    case 'err': return node.v;
    case 'empty': return '';
    case 'array': return `{${node.rows.map((r) => r.map(constText).join(',')).join(';')}}`;
    case 'ref': return at(node, src.slice(node.s, node.e));
    case 'sref': {
      const text = file ? canonicalRef(node.table, node.spec, opt.hereTable) : src.slice(node.s, node.e);
      return at(node, text);
    }
    case 'name': {
      const key = node.v.toLowerCase();
      if (!node.sheet && scope.has(key)) return file ? `_xlpm.${node.v}` : node.v;
      const text = src.slice(node.s, node.e).replace(/^_xlpm\.|^_xleta\./i, '');
      return at(node, text);
    }
    case 'paren': return `(${P(node.a, scalar)})`;
    case 'union': return `(${node.items.map((x) => P(x)).join(',')})`;
    case 'range': return `${P(node.a)}:${P(node.b)}`;
    case 'spill': return file ? `_xlfn.ANCHORARRAY(${P(node.a)})` : `${P(node.a)}#`;
    case 'at': return file ? (opt.dynamic ? `_xlfn.SINGLE(${P(node.a)})` : P(node.a)) : `@${P(node.a)}`;
    case 'neg': return `-${P(node.a, scalar)}`;
    case 'pos': return `+${P(node.a, scalar)}`;
    case 'pct': return `${P(node.a, scalar)}%`;
    case 'bin': return `${P(node.a, scalar)}${node.op}${P(node.b, scalar)}`;
    case 'call': return `${P(node.fn)}(${node.args.map((a) => P(a)).join(',')})`;
    case 'func': {
      const name = node.name;
      if (!file && name === 'SINGLE') return `@${node.args.length === 1 && ['ref', 'name', 'sref', 'func', 'paren'].includes(node.args[0].type) ? P(node.args[0]) : `(${node.args.map((a) => P(a)).join(',')})`}`;
      if (!file && name === 'ANCHORARRAY' && node.args.length === 1) return `${P(node.args[0])}#`;
      const inner = new Set(scope);
      for (const n of boundNames(node)) inner.add(n);
      const fn = FUNCS[name];
      const lifted = (i) => (fn && (fn.liftPos === 'all' || (Array.isArray(fn.liftPos) && fn.liftPos.includes(i)))) || scalarArg(name, i);
      const pass = ['IF', 'IFS', 'SWITCH', 'IFERROR', 'IFNA', 'CHOOSE', 'NOT'].includes(name);
      const args = node.args.map((a, i) => {
        // 함수 이름만 쓴 인수 (GROUPBY(..., SUM)) → _xleta.
        const up = a.type === 'name' && !a.sheet ? a.v.toUpperCase() : null;
        if (up && FUNCS[up] && !inner.has(a.v.toLowerCase()) && !opt.isName?.(a.v)) return file ? `_xleta.${up}` : up;
        return P(a, scalar && (pass || lifted(i)), inner);
      }).join(',');
      const text = `${file ? fileFuncName(name) : name}(${args})`;
      return ['INDEX', 'OFFSET', 'INDIRECT'].includes(name) ? at(node, text) : text;
    }
    default:
      return src.slice(node.s, node.e);
  }
}

/**
 * 앱 수식(맨 앞 '=' 포함) → 파일 수식 텍스트('=' 없음)
 * opt: { hereTable, dynamic, isName }
 */
export function toFileFormula(raw, opt = {}) {
  const body = raw.startsWith('=') ? raw.slice(1) : raw;
  let ast;
  try { ast = parse(body); } catch { return body; }
  return print(ast, body, { ...opt, mode: 'file' });
}

/**
 * 파일 수식 텍스트('=' 없음) → 앱 수식 본문('=' 없음)
 * opt: { legacy: 동적 배열 표시가 없는 옛 수식이면 true (암시적 교차 '@' 추가), isName }
 */
export function fromFileFormula(text, opt = {}) {
  // 구글 스프레드시트가 내보낸 전용 함수: IFERROR(__xludf.DUMMYFUNCTION("IMPORTXML(…)"), 저장값) → 원래 수식
  const gs = /^\s*(?:IFERROR\(\s*)?__xludf\.DUMMYFUNCTION\(\s*"((?:[^"]|"")*)"\s*\)/i.exec(text);
  if (gs) return gs[1].replace(/""/g, '"');
  let ast;
  try { ast = parse(text); } catch { return text.replace(/_xlfn\.|_xlws\.|_xlpm\.|_xleta\./gi, ''); }
  return print(ast, text, { ...opt, mode: 'app' }, new Set(), !!opt.legacy);
}
