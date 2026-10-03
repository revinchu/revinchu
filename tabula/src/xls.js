// 엑셀 97-2003 통합 문서(.xls, BIFF8) 읽기 (DOM 없음)
// OLE 복합 문서(CFB)의 'Workbook' 스트림 → 레코드 → 시트 · 셀 · 서식 · 병합 · 열 너비 · 행 높이 · 틀 고정 · 이름 · 수식.
// 수식은 토큰(ptg)을 엑셀 파일 형식 글자로 되돌린 뒤 xlsx 와 같은 변환(fromFileFormula)을 거침. 계산 결과는 cached 로 보관.
import { CellMap } from './cellmap.js';
import { protectFromAttrs } from './protect.js';
import { relocateValidation, VALIDATION_IME_MODES } from './validation.js';
import { toBase64 } from './vba.js';
import { inflate } from './zip.js';
import { emfToSvg } from './emf.js';
import { styleForCode } from './format.js';
import { colToName, nameToCol, quoteSheetName } from './formula.js';
import { fromFileFormula } from './xlfn.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { textRaw, numberRaw, INDEXED, BUILTIN_FMT, BUILTIN_CODE, digitWidth, width2pxM, baseColPx, pt2px, parsePrintAreas } from './xlsx.js';

// 함수 번호 → 이름 · 고정 인수 개수(-1 = 가변) (MS-XLS 함수 표)
const FTAB_TEXT = [
  '0:COUNT:-1 1:IF:-1 2:ISNA:1 3:ISERROR:1 4:SUM:-1 5:AVERAGE:-1 6:MIN:-1 7:MAX:-1 8:ROW:-1 9:COLUMN:-1 10:NA:0 11:NPV:-1 12:STDEV:-1 13:DOLLAR:-1',
  '14:FIXED:-1 15:SIN:1 16:COS:1 17:TAN:1 18:ATAN:1 19:PI:0 20:SQRT:1 21:EXP:1 22:LN:1 23:LOG10:1 24:ABS:1 25:INT:1 26:SIGN:1 27:ROUND:2 28:LOOKUP:-1',
  '29:INDEX:-1 30:REPT:2 31:MID:3 32:LEN:1 33:VALUE:1 34:TRUE:0 35:FALSE:0 36:AND:-1 37:OR:-1 38:NOT:1 39:MOD:2 40:DCOUNT:3 41:DSUM:3 42:DAVERAGE:3',
  '43:DMIN:3 44:DMAX:3 45:DSTDEV:3 46:VAR:-1 47:DVAR:3 48:TEXT:2 49:LINEST:-1 50:TREND:-1 51:LOGEST:-1 52:GROWTH:-1 56:PV:-1 57:FV:-1 58:NPER:-1',
  '59:PMT:-1 60:RATE:-1 61:MIRR:3 62:IRR:-1 63:RAND:0 64:MATCH:-1 65:DATE:3 66:TIME:3 67:DAY:1 68:MONTH:1 69:YEAR:1 70:WEEKDAY:-1 71:HOUR:1 72:MINUTE:1',
  '73:SECOND:1 74:NOW:0 75:AREAS:1 76:ROWS:1 77:COLUMNS:1 78:OFFSET:-1 82:SEARCH:-1 83:TRANSPOSE:1 86:TYPE:1 92:SERIESSUM:4 97:ATAN2:2 98:ASIN:1',
  '99:ACOS:1 100:CHOOSE:-1 101:HLOOKUP:-1 102:VLOOKUP:-1 105:ISREF:1 109:LOG:-1 111:CHAR:1 112:LOWER:1 113:UPPER:1 114:PROPER:1 115:LEFT:-1 116:RIGHT:-1',
  '117:EXACT:2 118:TRIM:1 119:REPLACE:4 120:SUBSTITUTE:-1 121:CODE:1 124:FIND:-1 125:CELL:-1 126:ISERR:1 127:ISTEXT:1 128:ISNUMBER:1 129:ISBLANK:1',
  '130:T:1 131:N:1 140:DATEVALUE:1 141:TIMEVALUE:1 142:SLN:3 143:SYD:4 144:DDB:-1 148:INDIRECT:-1 162:CLEAN:1 163:MDETERM:1 164:MINVERSE:1 165:MMULT:2',
  '167:IPMT:-1 168:PPMT:-1 169:COUNTA:-1 183:PRODUCT:-1 184:FACT:1 189:DPRODUCT:3 190:ISNONTEXT:1 193:STDEVP:-1 194:VARP:-1 195:DSTDEVP:3 196:DVARP:3',
  '197:TRUNC:-1 198:ISLOGICAL:1 199:DCOUNTA:3 204:USDOLLAR:-1 205:FINDB:-1 206:SEARCHB:-1 207:REPLACEB:4 208:LEFTB:-1 209:RIGHTB:-1 210:MIDB:3 211:LENB:1',
  '212:ROUNDUP:2 213:ROUNDDOWN:2 214:ASC:1 215:DBCS:1 216:RANK:-1 219:ADDRESS:-1 220:DAYS360:-1 221:TODAY:0 222:VDB:-1 227:MEDIAN:-1 228:SUMPRODUCT:-1',
  '229:SINH:1 230:COSH:1 231:TANH:1 232:ASINH:1 233:ACOSH:1 234:ATANH:1 235:DGET:3 244:INFO:1 247:DB:-1 252:FREQUENCY:2 261:ERROR.TYPE:1 269:AVEDEV:-1',
  '270:BETADIST:-1 271:GAMMALN:1 272:BETAINV:-1 273:BINOMDIST:4 274:CHIDIST:2 275:CHIINV:2 276:COMBIN:2 277:CONFIDENCE:3 278:CRITBINOM:3 279:EVEN:1',
  '280:EXPONDIST:3 281:FDIST:3 282:FINV:3 283:FISHER:1 284:FISHERINV:1 285:FLOOR:2 286:GAMMADIST:4 287:GAMMAINV:3 288:CEILING:2 289:HYPGEOMDIST:4',
  '290:LOGNORMDIST:3 291:LOGINV:3 292:NEGBINOMDIST:3 293:NORMDIST:4 294:NORMSDIST:1 295:NORMINV:3 296:NORMSINV:1 297:STANDARDIZE:3 298:ODD:1 299:PERMUT:2',
  '300:POISSON:3 301:TDIST:3 302:WEIBULL:4 303:SUMXMY2:2 304:SUMX2MY2:2 305:SUMX2PY2:2 306:CHITEST:2 307:CORREL:2 308:COVAR:2 309:FORECAST:3 310:FTEST:2',
  '311:INTERCEPT:2 312:PEARSON:2 313:RSQ:2 314:STEYX:2 315:SLOPE:2 316:TTEST:4 317:PROB:-1 318:DEVSQ:-1 319:GEOMEAN:-1 320:HARMEAN:-1 321:SUMSQ:-1',
  '322:KURT:-1 323:SKEW:-1 324:ZTEST:-1 325:LARGE:2 326:SMALL:2 327:QUARTILE:2 328:PERCENTILE:2 329:PERCENTRANK:-1 330:MODE:-1 331:TRIMMEAN:2 332:TINV:2',
  '336:CONCATENATE:-1 337:POWER:2 342:RADIANS:1 343:DEGREES:1 344:SUBTOTAL:-1 345:SUMIF:-1 346:COUNTIF:2 347:COUNTBLANK:1 350:ISPMT:4 351:DATEDIF:3',
  '352:DATESTRING:1 353:NUMBERSTRING:2 354:ROMAN:-1 358:GETPIVOTDATA:2 359:HYPERLINK:-1 360:PHONETIC:1 361:AVERAGEA:-1 362:MAXA:-1 363:MINA:-1',
  '364:STDEVPA:-1 365:VARPA:-1 366:STDEVA:-1 367:VARA:-1 368:BAHTTEXT:1 369:THAIDAYOFWEEK:1 370:THAIDIGIT:1 371:THAIMONTHOFYEAR:1 372:THAINUMSOUND:1',
  '373:THAINUMSTRING:1 374:THAISTRINGLENGTH:1 375:ISTHAIDIGIT:1 376:ROUNDBAHTDOWN:1 377:ROUNDBAHTUP:1 378:THAIYEAR:1 379:RTD:-1',
].join(' ');

// ───────────────────────── 복합 문서(CFB) ─────────────────────────
/** OLE 복합 문서에서 스트림 하나 읽기 (없으면 null) */
function cfbStream(buf, names) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (SIG.some((b, i) => buf[i] !== b)) throw new Error('엑셀 97-2003 파일(.xls) 형식이 아닙니다.');
  const secSize = 1 << dv.getUint16(30, true);
  const miniSize = 1 << dv.getUint16(32, true);
  const dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true);
  const miniFatStart = dv.getUint32(60, true);
  let difatSec = dv.getUint32(68, true);
  const END = 0xfffffffe;
  const off = (s) => (s + 1) * secSize;
  // FAT 섹터 목록 (헤더의 109개 + DIFAT 사슬)
  const fatSecs = [];
  for (let i = 0; i < 109; i++) { const s = dv.getUint32(76 + i * 4, true); if (s < 0xfffffffa) fatSecs.push(s); }
  for (let guard = 0; difatSec < 0xfffffffa && guard < 100000; guard++) {
    const n = secSize / 4 - 1;
    for (let i = 0; i < n; i++) { const s = dv.getUint32(off(difatSec) + i * 4, true); if (s < 0xfffffffa) fatSecs.push(s); }
    difatSec = dv.getUint32(off(difatSec) + n * 4, true);
  }
  const perSec = secSize / 4;
  const fat = new Uint32Array(fatSecs.length * perSec);
  fatSecs.forEach((s, k) => { for (let i = 0; i < perSec && off(s) + i * 4 + 4 <= buf.length; i++) fat[k * perSec + i] = dv.getUint32(off(s) + i * 4, true); });
  const chain = (start, table) => {
    const out = [];
    for (let s = start, guard = 0; s < table.length && s !== END && guard < 1e7; s = table[s], guard++) out.push(s);
    return out;
  };
  const readChain = (start) => {
    const secs = chain(start, fat);
    const out = new Uint8Array(secs.length * secSize);
    secs.forEach((s, i) => out.set(buf.subarray(off(s), Math.min(buf.length, off(s) + secSize)), i * secSize));
    return out;
  };
  const dir = readChain(dirStart);
  const ddv = new DataView(dir.buffer);
  const entries = [];
  for (let p = 0; p + 128 <= dir.length; p += 128) {
    const len = ddv.getUint16(p + 64, true);
    let name = '';
    for (let i = 0; i + 2 < len; i += 2) name += String.fromCharCode(ddv.getUint16(p + i, true));
    entries.push({ name, type: dir[p + 66], start: ddv.getUint32(p + 116, true), size: ddv.getUint32(p + 120, true) });
  }
  const root = entries.find((e) => e.type === 5);
  const want = entries.find((e) => e.type === 2 && names.includes(e.name));
  if (!want) return null;
  if (want.size >= miniCutoff || !root) return readChain(want.start).subarray(0, want.size);
  // 작은 스트림은 미니 스트림에서
  const miniStream = readChain(root.start);
  const mf = readChain(miniFatStart);
  const mdv = new DataView(mf.buffer);
  const miniFat = new Uint32Array(mf.length / 4);
  for (let i = 0; i < miniFat.length; i++) miniFat[i] = mdv.getUint32(i * 4, true);
  const secs = chain(want.start, miniFat);
  const out = new Uint8Array(secs.length * miniSize);
  secs.forEach((s, i) => out.set(miniStream.subarray(s * miniSize, s * miniSize + miniSize), i * miniSize));
  return out.subarray(0, want.size);
}

// ───────────────────────── BIFF8 레코드 ─────────────────────────
/** 레코드 목록 [{ type, data, cont: [CONTINUE 데이터…], pos }] */
function records(stream) {
  const dv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const out = [];
  for (let p = 0; p + 4 <= stream.length;) {
    const type = dv.getUint16(p, true);
    const len = dv.getUint16(p + 2, true);
    const data = stream.subarray(p + 4, p + 4 + len);
    if (type === 0x3c && out.length) out[out.length - 1].cont.push(data);
    else out.push({ type, data, cont: [], pos: p });
    p += 4 + len;
  }
  return out;
}

/** 여러 조각(레코드 + CONTINUE)에 걸친 데이터를 차례로 읽기. 글자가 조각을 넘으면 새 조각 첫 바이트가 글자 폭 */
class Reader {
  constructor(chunks) { this.chunks = chunks; this.ci = 0; this.p = 0; }
  get cur() { return this.chunks[this.ci]; }
  ensure() { while (this.cur && this.p >= this.cur.length && this.ci < this.chunks.length - 1) { this.ci++; this.p = 0; } }
  u8() { this.ensure(); if (this.strict && (!this.cur || this.p >= this.cur.length)) throw new Error('XLS 설정 레코드가 잘렸습니다.'); return this.cur[this.p++]; }
  u16() { return this.u8() | (this.u8() << 8); }
  u32() { return (this.u16() | (this.u16() << 16)) >>> 0; }
  skip(n) { for (let i = 0; i < n; i++) this.u8(); }
  /** XLUnicodeString (cch 2바이트 또는 1바이트) */
  str(cchBytes = 2) {
    const cch = cchBytes === 1 ? this.u8() : this.u16();
    let flags = this.u8();
    let high = flags & 1;
    const runs = flags & 8 ? this.u16() : 0;
    const ext = flags & 4 ? this.u32() : 0;
    let s = '';
    for (let i = 0; i < cch; i++) {
      if (this.p >= this.cur.length && this.ci < this.chunks.length - 1) { this.ci++; this.p = 0; high = this.cur[this.p++] & 1; }
      s += String.fromCharCode(high ? this.u16() : this.u8());
    }
    this.skip(runs * 4 + ext);
    return s;
  }
}
const rd = (data) => new Reader([data]);
const f64 = (b, p) => new DataView(b.buffer, b.byteOffset + p, 8).getFloat64(0, true);
const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const i16 = (b, p) => { const v = u16(b, p); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, p) => (u16(b, p) | (u16(b, p + 2) << 16)) >>> 0;
function rkValue(rk) {
  let v;
  if (rk & 2) v = rk >> 2; // 정수 30비트 (부호 있음)
  else {
    const b = new DataView(new ArrayBuffer(8));
    b.setUint32(4, rk & 0xfffffffc, true);
    v = b.getFloat64(0, true);
  }
  return rk & 1 ? v / 100 : v;
}
const ERRS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A' };

// ───────────────────────── 수식 (ptg → 글자) ─────────────────────────
const FTAB = new Map(FTAB_TEXT.split(' ').map((t) => { const [i, name, n] = t.split(':'); return [Number(i), { name, n: Number(n) }]; }));
const BIN = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>', 0x0f: ' ', 0x10: ',', 0x11: ':' };
const quoteStr = (s) => `"${s.replace(/"/g, '""')}"`;

/**
 * rgce(수식 토큰) → 엑셀 파일 형식 수식 글자 (맨 앞 '=' 없음)
 * ctx: { sheetRef(ixti) → '시트!' 접두사, name(iname), nameX(ixti, iname), row, col (공유 수식 기준 칸) }
 */
function decodeFormula(rgce, rgcb, ctx) {
  const st = [];
  let p = 0;
  let q = 0; // rgcb(배열 상수 · 메모 영역) 읽는 위치
  const cell = (rw, cf, rel) => {
    // BIFF8: 행 16비트, 열 14비트 + 비트 14 = 열 상대, 비트 15 = 행 상대. 공유 수식(rel)의 상대 참조는 기준 칸에서의 거리
    const colRel = !!(cf & 0x4000);
    const rowRel = !!(cf & 0x8000);
    let r = rw;
    let c = cf & 0x3fff;
    if (rel) {
      if (rowRel) r = (ctx.row + (rw & 0x8000 ? rw - 0x10000 : rw) + 65536) % 65536;
      if (colRel) c = (ctx.col + ((c & 0xff) > 127 ? (c & 0xff) - 256 : c & 0xff) + 256) % 256;
    }
    return `${colRel ? '' : '$'}${colToName(c)}${rowRel ? '' : '$'}${r + 1}`;
  };
  // 영역: 행 전체(1~65536)면 열 참조 'A:B', 열 전체(A~IV)면 행 참조 '1:2' (엑셀 표시와 같음)
  const area = (r1, r2, c1f, c2f, rel) => {
    const whole = (f) => !rel || !(f & 0xc000);
    if (r1 === 0 && r2 === 0xffff && whole(c1f) && whole(c2f)) return `${c1f & 0x4000 ? '' : '$'}${colToName(c1f & 0x3fff)}:${c2f & 0x4000 ? '' : '$'}${colToName(c2f & 0x3fff)}`;
    if ((c1f & 0x3fff) === 0 && (c2f & 0x3fff) === 0xff && whole(c1f) && whole(c2f)) return `${c1f & 0x8000 ? '' : '$'}${r1 + 1}:${c2f & 0x8000 ? '' : '$'}${r2 + 1}`;
    return `${cell(r1, c1f, rel)}:${cell(r2, c2f, rel)}`;
  };
  const pop = () => st.pop() ?? '';
  while (p < rgce.length) {
    const ptg = rgce[p++];
    const base = ptg < 0x20 ? ptg : (ptg & 0x1f) | 0x20;
    switch (base) {
      case 0x01: case 0x02: p += 4; st.push(ctx.shared ?? '#REF!'); break; // ptgExp / ptgTbl (공유 · 배열 수식은 따로)
      case 0x12: st.push(`+${pop()}`); break;
      case 0x13: st.push(`-${pop()}`); break;
      case 0x14: st.push(`${pop()}%`); break;
      case 0x15: st.push(`(${pop()})`); break;
      case 0x16: st.push(''); break;
      case 0x17: { // 글자
        const cch = rgce[p];
        const high = rgce[p + 1] & 1;
        p += 2;
        let s = '';
        for (let i = 0; i < cch; i++) { s += String.fromCharCode(high ? u16(rgce, p) : rgce[p]); p += high ? 2 : 1; }
        st.push(quoteStr(s));
        break;
      }
      case 0x18: p += 5; break; // ptgExtend (드묾)
      case 0x19: { // ptgAttr
        const f = rgce[p];
        const w = u16(rgce, p + 1);
        p += 3;
        if (f & 0x04) p += (w + 1) * 2; // CHOOSE 점프 표
        if (f & 0x10) st.push(`SUM(${pop()})`);
        break;
      }
      case 0x1c: st.push(ERRS[rgce[p]] ?? '#N/A'); p += 1; break;
      case 0x1d: st.push(rgce[p] ? 'TRUE' : 'FALSE'); p += 1; break;
      case 0x1e: st.push(String(u16(rgce, p))); p += 2; break;
      case 0x1f: st.push(String(f64(rgce, p))); p += 8; break;
      case 0x20: { // 배열 상수: 값은 rgcb 에
        p += 7;
        const cols = rgcb[q] + 1;
        const rows = u16(rgcb, q + 1) + 1;
        q += 3;
        const out = [];
        for (let r = 0; r < rows; r++) {
          const row = [];
          for (let c = 0; c < cols; c++) {
            const t = rgcb[q++];
            if (t === 0x01) { row.push(String(f64(rgcb, q))); q += 8; }
            else if (t === 0x02) {
              const cch = u16(rgcb, q);
              const high = rgcb[q + 2] & 1;
              q += 3;
              let s = '';
              for (let i = 0; i < cch; i++) { s += String.fromCharCode(high ? u16(rgcb, q) : rgcb[q]); q += high ? 2 : 1; }
              row.push(quoteStr(s));
            } else if (t === 0x04) { row.push(rgcb[q] ? 'TRUE' : 'FALSE'); q += 8; }
            else if (t === 0x10) { row.push(ERRS[rgcb[q]] ?? '#N/A'); q += 8; }
            else { row.push(''); q += 8; }
          }
          out.push(row.join(','));
        }
        st.push(`{${out.join(';')}}`);
        break;
      }
      case 0x21: case 0x22: { // 함수
        let argc;
        let idx;
        if (base === 0x21) { idx = u16(rgce, p); p += 2; argc = FTAB.get(idx)?.n ?? 0; }
        else { argc = rgce[p] & 0x7f; idx = u16(rgce, p + 1) & 0x7fff; p += 3; }
        const args = st.splice(Math.max(0, st.length - argc), argc);
        if (idx === 255) { const fn = args.shift() ?? ''; st.push(`${fn}(${args.join(',')})`); } // 추가 기능 함수 (이름이 첫 인수)
        else st.push(`${FTAB.get(idx)?.name ?? `_xlfn.UNKNOWN${idx}`}(${args.join(',')})`);
        break;
      }
      case 0x23: st.push(ctx.name(u32(rgce, p))); p += 4; break; // ptgName: 이름 번호 4바이트
      case 0x24: st.push(cell(u16(rgce, p), u16(rgce, p + 2), false)); p += 4; break;
      case 0x25: st.push(area(u16(rgce, p), u16(rgce, p + 2), u16(rgce, p + 4), u16(rgce, p + 6), false)); p += 8; break;
      // 메모(계산 최적화) 토큰: 무시 (MemArea 는 영역 목록이 rgcb 에, MemFunc 는 길이만 2바이트)
      case 0x26: case 0x27: case 0x28: p += 6; if (base === 0x26) { const n = u16(rgcb, q); q += 2 + n * 8; } break;
      case 0x29: p += 2; break;
      case 0x2a: st.push('#REF!'); p += 4; break;
      case 0x2b: st.push('#REF!'); p += 8; break;
      case 0x2c: st.push(cell(u16(rgce, p), u16(rgce, p + 2), true)); p += 4; break;
      case 0x2d: st.push(area(u16(rgce, p), u16(rgce, p + 2), u16(rgce, p + 4), u16(rgce, p + 6), true)); p += 8; break;
      case 0x39: st.push(ctx.nameX(u16(rgce, p), u32(rgce, p + 2))); p += 6; break;
      case 0x3a: st.push(`${ctx.sheetRef(u16(rgce, p))}${cell(u16(rgce, p + 2), u16(rgce, p + 4), false)}`); p += 6; break;
      case 0x3b: st.push(`${ctx.sheetRef(u16(rgce, p))}${area(u16(rgce, p + 2), u16(rgce, p + 4), u16(rgce, p + 6), u16(rgce, p + 8), false)}`); p += 10; break;
      case 0x3c: st.push(`${ctx.sheetRef(u16(rgce, p))}#REF!`); p += 6; break;
      case 0x3d: st.push(`${ctx.sheetRef(u16(rgce, p))}#REF!`); p += 10; break;
      default:
        if (BIN[base]) { const b = pop(); const a = pop(); st.push(`${a}${BIN[base]}${b}`); break; }
        throw new Error(`지원하지 않는 수식 토큰 0x${ptg.toString(16)}`);
    }
  }
  return st.join('');
}

// ───────────────────────── 그림 (OfficeArt) ─────────────────────────
/** OfficeArt 레코드 목록 { type, inst, data, start } (컨테이너는 children 에 펼침) */
function artRecords(buf, start = 0, end = buf.length, out = []) {
  for (let p = start; p + 8 <= end;) {
    const vi = u16(buf, p);
    const type = u16(buf, p + 2);
    const len = u32(buf, p + 4);
    const rec = { type, inst: vi >> 4, ver: vi & 0xf, data: buf.subarray(p + 8, Math.min(end, p + 8 + len)), children: null };
    if ((vi & 0xf) === 0xf) rec.children = artRecords(buf, p + 8, Math.min(end, p + 8 + len), []);
    out.push(rec);
    p += 8 + len;
  }
  return out;
}
/** 그림 저장소(BStore)의 그림 → [{ mime, bytes } | null] (1부터 번호) */
function blipStore(dgg) {
  const out = [];
  const walk = (list) => {
    for (const r of list) {
      if (r.type === 0xf001 && r.children) {
        for (const fbse of r.children) {
          if (fbse.type !== 0xf007) { out.push(null); continue; }
          const d = fbse.data;
          const cbName = d[33];
          const blip = artRecords(d, 36 + cbName)[0];
          out.push(blip ? blipData(blip) : null);
        }
      } else if (r.children) walk(r.children);
    }
  };
  walk(artRecords(dgg));
  return out;
}
function blipData(b) {
  // 그림 앞의 고유 번호(16바이트, 두 개일 수도) + 태그 1바이트 뒤가 그림 자체
  const two = { 0xf01e: 0x6e1, 0xf01d: 0x46b, 0xf02a: 0x6e3, 0xf01f: 0x7a9 }[b.type] === b.inst || (b.type === 0xf01d && b.inst === 0x6e3);
  const skip = (two ? 32 : 16) + 1;
  const bytes = b.data.subarray(skip);
  if (b.type === 0xf01e) return { mime: 'image/png', bytes };
  if (b.type === 0xf01d || b.type === 0xf02a) return { mime: 'image/jpeg', bytes };
  if (b.type === 0xf01f) { // DIB → BMP 파일 머리말 붙이기
    const head = new Uint8Array(14);
    const size = bytes.length + 14;
    const hdrSize = u32(bytes, 0);
    const bits = u16(bytes, 14);
    const colors = u32(bytes, 32) || (bits <= 8 ? 1 << bits : 0);
    const offBits = 14 + hdrSize + colors * 4;
    head.set([0x42, 0x4d, size & 255, (size >> 8) & 255, (size >> 16) & 255, (size >>> 24) & 255, 0, 0, 0, 0, offBits & 255, (offBits >> 8) & 255, 0, 0]);
    const out = new Uint8Array(size);
    out.set(head);
    out.set(bytes, 14);
    return { mime: 'image/bmp', bytes: out };
  }
  if (b.type === 0xf01a) { // EMF: 고유 번호 + 메타파일 머리말(34바이트) 뒤에 zlib 압축 데이터 → SVG 로 바꿔 보여 줌
    const m = b.data.subarray(b.inst === 0x3d5 ? 32 : 16);
    if (m.length < 34) return null;
    const size = u32(m, 0), packed = m.subarray(34, 34 + u32(m, 28));
    let emf = packed;
    try { if (m[32] === 0) emf = inflate(packed.subarray(2), size); } catch { return null; }
    const svg = emfToSvg(emf);
    return svg ? { mime: 'image/svg+xml', bytes: new TextEncoder().encode(svg), emf } : null;
  }
  return null; // WMF · PICT: 브라우저가 그릴 수 없음
}
/** 시트의 그림 도형: [{ pib, anchor: {c1, dx1, r1, dy1, c2, dx2, r2, dy2}, name }] */
function sheetPictures(dg) {
  const out = [];
  const walk = (list) => {
    for (const r of list) {
      if (r.type === 0xf004 && r.children) {
        let pib = 0;
        let anchor = null;
        let crop = null;
        let name = '';
        for (const c of r.children) {
          if (c.type === 0xf00b) { // 속성 표
            const n = c.inst;
            let tail = n * 6;
            for (let k = 0; k < n; k++) {
              const id = u16(c.data, k * 6);
              const op = u32(c.data, k * 6 + 2);
              if ((id & 0x3fff) === 0x0104) pib = op;
              // 그림 자르기 (16.16 고정 소수: 위 · 아래 · 왼쪽 · 오른쪽 비율)
              if ((id & 0x3fff) >= 0x0100 && (id & 0x3fff) <= 0x0103 && op) (crop ??= {})['tblr'[(id & 0x3fff) - 0x0100]] = (op | 0) / 65536;
              if ((id & 0x3fff) === 0x0380 && id & 0x8000) { const bytes = c.data.subarray(tail, tail + op); for (let j = 0; j + 1 < bytes.length; j += 2) { const ch = u16(bytes, j); if (!ch) break; name += String.fromCharCode(ch); } }
              if (id & 0x8000) tail += op;
            }
          } else if (c.type === 0xf010 && c.data.length >= 18) {
            const d = c.data;
            anchor = { c1: u16(d, 2), dx1: u16(d, 4), r1: u16(d, 6), dy1: u16(d, 8), c2: u16(d, 10), dx2: u16(d, 12), r2: u16(d, 14), dy2: u16(d, 16) };
          }
        }
        if (pib && anchor) out.push({ pib, anchor, name, ...(crop ? { crop } : {}) });
      }
      if (r.children) walk(r.children);
    }
  };
  walk(artRecords(dg));
  return out;
}

// ───────────────────────── 서식 ─────────────────────────
const DG = ['', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
const FLS = [null, 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];

// XFExt의 실제 RGB는 기존 팔레트 색의 근사값보다 우선한다. 오래된 앱이 XF만
// 바꾼 파일의 낡은 확장을 적용하지 않도록 XFCRC(MS-OSHARED 2.4.3, 다항식 AF)를 확인.
function xfExtendedColors(xfs, extensions, check) {
  const out = new Map();
  if (!extensions.length || !check || check.length !== 20 || u16(check, 0) !== 0x087c || u16(check, 14) !== xfs.length) return out;
  let crc = 0;
  for (const xf of xfs) for (const byte of xf) {
    crc ^= byte << 24;
    for (let bit = 0; bit < 8; bit++) crc = (crc << 1) ^ (crc < 0 ? 0xaf : 0);
  }
  if ((crc >>> 0) !== u32(check, 16)) return out;
  for (const d of extensions) {
    if (d.length < 20 || u16(d, 0) !== 0x087d) continue;
    const ix = u16(d, 14), xf = xfs[ix];
    if (!xf || xf.length < 20 || !(u32(xf, 14) & 0x02000000)) continue;
    const colors = {}, count = u16(d, 18);
    let p = 20, valid = true;
    for (let k = 0; k < count; k++) {
      if (p + 4 > d.length) { valid = false; break; }
      const type = u16(d, p), size = u16(d, p + 2);
      if (size < 4 || p + size > d.length) { valid = false; break; }
      if ([4, 5, 7, 8, 9, 10, 11, 13].includes(type)) {
        if (size !== 20) { valid = false; break; }
        // FullColorExt: 명시적 불투명 RGB만. 자동/인덱스/테마/틴트/알파와
        // 미지원 속성은 추측하지 않고 원래 XF/FONT 팔레트 색을 유지한다.
        if (u16(d, p + 4) === 2 && i16(d, p + 6) === 0 && d[p + 11] === 255) {
          colors[type] = `#${Array.from(d.subarray(p + 8, p + 11), b => b.toString(16).padStart(2, '0')).join('')}`;
        }
      }
      p += size;
    }
    if (valid && p === d.length) out.set(ix, colors);
  }
  return out;
}

// ───────────────────────── 읽기 ─────────────────────────
/** .xls(엑셀 97-2003) → { data: { sheets, names, defaultFont }, warnings } — readXlsx 와 같은 모양 */
export function readXls(bytes) {
  const buf = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const stream = cfbStream(buf, ['Workbook', 'Book']);
  if (!stream) {
    // 암호를 건 .xlsx 는 같은 OLE 컨테이너 안에 EncryptionInfo · EncryptedPackage 로 들어 있음
    if (cfbStream(buf, ['EncryptionInfo', 'EncryptedPackage'])) throw new Error('암호로 보호된 파일입니다. 엑셀에서 암호를 해제(파일 → 정보 → 통합 문서 보호 → 암호 설정에서 암호 지우기)한 뒤 열어 주세요.');
    throw new Error('.xls 파일에서 통합 문서 데이터를 찾지 못했습니다.');
  }
  const recs = records(stream);
  // FILEPASS: 열기 암호가 걸린 .xls
  if (recs.some(r => r.type === 0x002f)) throw new Error('암호로 보호된 파일입니다. 엑셀에서 암호를 해제한 뒤 열어 주세요.');
  const warnings = [];
  if (recs[0]?.type !== 0x0809 || u16(recs[0].data, 0) !== 0x0600) warnings.push('엑셀 5.0/95 이전 형식은 일부만 읽을 수 있습니다.');
  const palette = INDEXED.slice();
  const fonts = [];
  const formats = {};
  const xfRaw = [];
  const xfExtensions = [];
  let xfCheck = null;
  const bound = [];
  let sst = [];
  const names = [];
  const supbooks = [];
  let xti = [];
  let date1904 = false;
  const calculation = {}, workbookProtection = {};
  const dggParts = [];
  // 전역 부분 (첫 EOF 까지)
  let i = 1;
  for (; i < recs.length && recs[i].type !== 0x000a; i++) {
    const { type, data, cont } = recs[i];
    switch (type) {
      case 0x000c: if (data.length >= 2) calculation.iterateCount = u16(data, 0); break;
      case 0x000d: if (data.length >= 2) calculation.mode = i16(data, 0) === 0 ? 'manual' : i16(data, 0) === -1 ? 'autoNoTable' : 'auto'; break;
      case 0x000e: if (data.length >= 2) calculation.fullPrecision = !!u16(data, 0); break;
      case 0x000f: if (data.length >= 2) calculation.refMode = u16(data, 0) ? 'A1' : 'R1C1'; break;
      case 0x0010: if (data.length >= 8) calculation.iterateDelta = f64(data, 0); break;
      case 0x0011: if (data.length >= 2) calculation.iterate = !!u16(data, 0); break;
      case 0x005f: if (data.length >= 2) calculation.calcOnSave = !!u16(data, 0); break;
      case 0x0012: if (data.length >= 2) workbookProtection.lockStructure = String(u16(data, 0) ? 1 : 0); break;
      case 0x0019: if (data.length >= 2) workbookProtection.lockWindows = String(u16(data, 0) ? 1 : 0); break;
      case 0x0013: if (data.length >= 2 && u16(data, 0)) workbookProtection.workbookPassword = u16(data, 0).toString(16).toUpperCase().padStart(4, '0'); break;
      case 0x0022: date1904 = u16(data, 0) === 1; break;
      case 0x0092: { const n = u16(data, 0); for (let k = 0; k < n; k++) palette[8 + k] = [data[2 + k * 4], data[3 + k * 4], data[4 + k * 4]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase(); break; }
      case 0x0031: {
        const r = rd(data);
        const height = r.u16(); const grbit = r.u16(); const icv = r.u16(); const bls = r.u16(); r.u16(); const uls = r.u8(); r.skip(3);
        fonts.push({ size: height / 20, italic: !!(grbit & 2), strike: !!(grbit & 8), icv, bold: bls >= 700, underline: uls > 0, name: r.str(1) });
        break;
      }
      case 0x041e: { const r = rd(data); const id = r.u16(); formats[id] = r.str(2); break; }
      case 0x00e0: xfRaw.push(data); break;
      case 0x087c: xfCheck = data; break;
      case 0x087d: xfExtensions.push(data); break;
      case 0x00eb: dggParts.push(data, ...cont); break;
      case 0x0085: { const r = new Reader([data]); const pos = r.u32(); const state = r.u8(); const dt = r.u8(); bound.push({ pos, state, dt, name: r.str(1) }); break; }
      case 0x00fc: { const r = new Reader([data, ...cont]); r.u32(); const n = r.u32(); sst = new Array(n); for (let k = 0; k < n; k++) sst[k] = r.str(2); break; }
      case 0x01ae: { const ctab = u16(data, 0); const cch = u16(data, 2); supbooks.push({ self: cch === 0x0401, addin: cch === 0x3a01, ctab, externs: [] }); break; }
      case 0x0023: { const sb = supbooks[supbooks.length - 1]; if (sb) { const r = rd(data); r.skip(6); sb.externs.push(r.str(1)); } break; }
      case 0x0017: { const n = u16(data, 0); xti = []; const all = new Uint8Array([...data, ...cont.flatMap((c) => [...c])]); for (let k = 0; k < n; k++) xti.push({ sb: u16(all, 2 + k * 6), first: i16(all, 4 + k * 6), last: i16(all, 6 + k * 6) }); break; }
      case 0x0018: {
        const flags = u16(data, 0);
        const cch = data[3];
        const cce = u16(data, 4);
        const itab = u16(data, 8);
        const r = new Reader([data]);
        r.p = 14;
        const high = r.u8() & 1;
        let name = '';
        for (let k = 0; k < cch; k++) name += String.fromCharCode(high ? r.u16() : r.u8());
        const builtin = !!(flags & 0x20);
        if (builtin) name = { '\u0006': 'Print_Area', '\u0007': 'Print_Titles', '\r': '_FilterDatabase', '\u0000': 'Consolidate_Area', '\u0001': 'Auto_Open' }[name[0]] ?? name;
        names.push({ name, builtin, hidden: !!(flags & 1), itab, rgce: data.subarray(r.p, r.p + cce), rgcb: data.subarray(r.p + cce) });
        break;
      }
      default: break;
    }
  }
  const concat = (parts) => { const n = parts.reduce((a, b) => a + b.length, 0); const out = new Uint8Array(n); let o = 0; for (const x of parts) { out.set(x, o); o += x.length; } return out; };
  let blips = [];
  try { if (dggParts.length) blips = blipStore(concat(dggParts)); } catch { blips = []; }
  const nFmtFont = (idx) => fonts[idx >= 4 ? idx - 1 : idx] ?? fonts[0];
  const color = (icv) => (icv === 0x40 || icv === 0x7fff || icv === undefined ? null : icv === 0x41 ? '#ffffff' : palette[icv] ? `#${palette[icv].toLowerCase()}` : null);
  const wbFont = { name: fonts[0]?.name ?? '맑은 고딕', size: fonts[0]?.size ?? 11 };
  const mdw = digitWidth(wbFont);
  // XF → 앱 셀 서식 (xlsx 가져오기와 같은 모양)
  const styleMemo = new Map();
  const extendedColors = xfExtendedColors(xfRaw, xfExtensions, xfCheck);
  const xfStyle = (ix) => {
    if (styleMemo.has(ix)) return styleMemo.get(ix);
    const d = xfRaw[ix];
    let st;
    if (!d) st = undefined;
    else {
      st = {};
      const f = nFmtFont(u16(d, 0));
      if (f) {
        if (f.bold) st.bold = true;
        if (f.italic) st.italic = true;
        if (f.underline) st.underline = true;
        if (f.strike) st.strike = true;
        if (f.size && f.size !== wbFont.size) st.size = f.size;
        if (f.name && f.name !== wbFont.name) st.font = f.name;
        const c = color(f.icv);
        if (c && c !== '#000000') st.color = c;
      }
      const ifmt = u16(d, 2);
      if (formats[ifmt] !== undefined) Object.assign(st, styleForCode(formats[ifmt]));
      else if (BUILTIN_CODE[ifmt]) Object.assign(st, styleForCode(BUILTIN_CODE[ifmt]));
      else if (BUILTIN_FMT[ifmt]) Object.assign(st, BUILTIN_FMT[ifmt]);
      const prot = u16(d, 4);
      if (!(prot & 1)) st.locked = false;
      if (prot & 2) st.hideFormula = true;
      const al = d[6];
      const h = al & 7;
      if (h === 1) st.align = 'left'; else if (h === 2 || h === 6 || h === 7) st.align = 'center'; else if (h === 3) st.align = 'right'; else if (h === 5) { st.align = 'left'; st.wrap = true; }
      if (al & 8) st.wrap = true;
      const v = (al >> 4) & 7;
      if (v === 0) st.valign = 'top'; else if (v === 1 || v === 3 || v === 4) st.valign = 'middle';
      const rot = d[7];
      if (rot) st.rotate = rot === 255 ? 255 : rot > 90 ? -(rot - 90) : rot;
      if (d[8] & 0x0f) st.indent = d[8] & 0x0f;
      if (d[8] & 0x10) st.shrink = true;
      const b1 = u32(d, 10);
      const b2 = u32(d, 14);
      const side = (key, dg, icv) => {
        if (!dg) return;
        st[key] = true;
        if (DG[dg] && DG[dg] !== 'thin') st[`${key}s`] = DG[dg];
        const c = color(icv);
        if (c && c !== '#000000') st[`${key}c`] = c;
      };
      side('bl', b1 & 0xf, (b1 >>> 16) & 0x7f);
      side('br', (b1 >>> 4) & 0xf, (b1 >>> 23) & 0x7f);
      side('bt', (b1 >>> 8) & 0xf, b2 & 0x7f);
      side('bb', (b1 >>> 12) & 0xf, (b2 >>> 7) & 0x7f);
      const diag = b1 >>> 30;
      if (diag) for (const k of [diag & 1 ? 'dd' : null, diag & 2 ? 'du' : null].filter(Boolean)) side(k, (b2 >>> 21) & 0xf, (b2 >>> 14) & 0x7f);
      const fls = (b2 >>> 26) & 0x3f;
      const colors = u16(d, 18);
      const fore = color(colors & 0x7f);
      const back = color((colors >>> 7) & 0x7f);
      if (fls === 1) { if (fore) st.fill = fore; }
      else if (FLS[fls]) { st.pattern = FLS[fls]; st.patternColor = fore ?? '#000000'; if (back) st.fill = back; }
      const ext = extendedColors.get(ix);
      if (ext) {
        if (ext[4]) { if (fls === 1) st.fill = ext[4]; else if (FLS[fls]) st.patternColor = ext[4]; }
        if (ext[5] && fls !== 1 && FLS[fls]) st.fill = ext[5];
        if (ext[13]) st.color = ext[13];
        for (const [type, key] of [[7, 'bt'], [8, 'bb'], [9, 'bl'], [10, 'br'], [11, 'dd'], [11, 'du']]) if (ext[type] && st[key]) st[`${key}c`] = ext[type];
      }
      for (const k of Object.keys(st)) if (st[k] === undefined) delete st[k];
      if (!Object.keys(st).length) st = undefined;
    }
    styleMemo.set(ix, st);
    return st;
  };
  // 시트 참조 (EXTERNSHEET 의 ixti → '시트!' 접두사)
  const allNames = bound.map((b) => b.name);
  const sheetRef = (ixti) => {
    const x = xti[ixti];
    const sb = x && supbooks[x.sb];
    if (!x || !sb || !sb.self) return '#REF!';
    if (x.first < 0) return '#REF!';
    const a = allNames[x.first] ?? '#REF!';
    const b = allNames[x.last] ?? a;
    return `${quoteSheetName(a === b ? a : `${a}:${b}`)}!`;
  };
  const nameText = (iname) => names[iname - 1]?.name ?? '#NAME?';
  const nameX = (ixti, iname) => {
    const x = xti[ixti];
    const sb = x && supbooks[x.sb];
    if (sb?.self) return nameText(iname);
    return sb?.externs[iname - 1] ?? '#NAME?';
  };
  const fctx = { sheetRef, name: nameText, nameX, row: 0, col: 0 };
  const decode = (rgce, rgcb, row, col) => {
    fctx.row = row;
    fctx.col = col;
    return fromFileFormula(decodeFormula(rgce, rgcb, fctx), { legacy: true });
  };

  // 시트
  const sheets = [];
  const byPos = new Map(recs.map((r, k) => [r.pos, k]));
  let unsupported = 0;
  for (const b of bound) {
    if (b.dt !== 0) { warnings.push(`'${b.name}' 차트 시트 · 매크로 시트는 가져오지 않았습니다.`); continue; }
    const sheet = {
      name: b.name.slice(0, 31), cells: new CellMap(), colWidths: {}, rowHeights: {}, merges: [], cond: [], colStyles: {}, rowStyles: {},
      hiddenRows: {}, hiddenCols: {}, rowManual: {}, freeze: { rows: 0, cols: 0 }, filter: null, charts: [], images: [], shapes: [], validations: [], slicers: [],
    };
    if (b.state === 1) sheet.state = 'hidden';
    else if (b.state === 2) sheet.state = 'veryHidden';
    let k = byPos.get(b.pos);
    if (k === undefined) { sheets.push(sheet); continue; }
    const shared = []; // 공유 수식 { r1, r2, c1, c2, rgce, rgcb }
    const pendingStr = [];
    let printFit = false, fitTo = null;
    const protectionAttrs = {};
    const put = (r, c, ixfe, value, formula) => {
      const style = xfStyle(ixfe);
      let raw;
      if (formula !== undefined) raw = `=${formula}`;
      else if (value === null || value === undefined) raw = '';
      else if (typeof value === 'number') raw = numberRaw(value, style, date1904);
      else if (typeof value === 'boolean') raw = value ? 'TRUE' : 'FALSE';
      else if (typeof value === 'object') raw = value.error;
      else raw = style?.numFmt === 'text' ? value : textRaw(value);
      if (!raw && !style) return;
      const d = { raw };
      if (style) d.style = style;
      if (formula !== undefined) {
        if (value !== undefined && value !== null) d.cached = value;
        if (style?.numFmt === 'text') d.fx = true;
      }
      sheet.cells.setRC(r, c, d);
    };
    let defColChars = 8;
    const dgParts = [];
    for (k += 1; k < recs.length && recs[k].type !== 0x000a; k++) {
      const { type, data } = recs[k];
      switch (type) {
        case 0x000c: if (data.length >= 2) calculation.iterateCount = u16(data, 0); break;
        case 0x000d: if (data.length >= 2) calculation.mode = i16(data, 0) === 0 ? 'manual' : i16(data, 0) === -1 ? 'autoNoTable' : 'auto'; break;
        case 0x000e: if (data.length >= 2) calculation.fullPrecision = !!u16(data, 0); break;
        case 0x000f: if (data.length >= 2) calculation.refMode = u16(data, 0) ? 'A1' : 'R1C1'; break;
        case 0x0010: if (data.length >= 8) calculation.iterateDelta = f64(data, 0); break;
        case 0x0011: if (data.length >= 2) calculation.iterate = !!u16(data, 0); break;
        case 0x005f: if (data.length >= 2) calculation.calcOnSave = !!u16(data, 0); break;
        case 0x01be: {
          // [MS-XLS] Dv / DVParsedFormula. 첫 영역 기준 상대 수식을 각 sqref로 옮긴다.
          const reader = new Reader([data, ...recs[k].cont]); reader.strict = true;
          const fl = reader.u32(), type = ['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'][fl & 15];
          if (!type) throw new Error('지원하지 않는 XLS 데이터 유효성 유형입니다. Excel에서 .xlsx로 저장한 파일을 여세요.');
          const rule = { type, op: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'][(fl >>> 20) & 15] ?? 'between', allowBlank: !!(fl & 256), showDropdown: !(fl & 512), showPrompt: !!(fl & 0x40000), showError: !!(fl & 0x80000), errorStyle: ['stop', 'warning', 'info'][(fl >>> 4) & 7] ?? 'stop' };
          const ime = VALIDATION_IME_MODES[(fl >>> 10) & 255]; if (ime) rule.imeMode = ime;
          for (const key of ['promptTitle', 'errorTitle', 'prompt', 'error']) { const value = reader.str(); if (value && value !== '\0') rule[key] = value; }
          const formulas = [];
          for (let i = 0; i < 2; i++) { const n = reader.u16(); reader.skip(2); const bytes = new Uint8Array(n); for (let j = 0; j < n; j++) bytes[j] = reader.u8(); formulas.push(bytes); }
          const count = reader.u16(), ranges = [];
          if (!count || count > 432) throw new Error('XLS 데이터 유효성 범위 개수가 올바르지 않습니다.');
          for (let i = 0; i < count; i++) { const r1 = reader.u16(), r2 = reader.u16(), c1 = reader.u16(), c2 = reader.u16(); if (r1 > r2 || c1 > c2 || c2 > 255) throw new Error('XLS 데이터 유효성 범위가 올바르지 않습니다.'); ranges.push({r1,r2,c1,c2}); }
          for (let i = 0; i < formulas.length; i++) if (formulas[i].length && type !== 'any') {
            let formula; try { formula = decode(formulas[i], new Uint8Array(), ranges[0].r1, ranges[0].c1); } catch { formula = null; }
            if (!formula) { formula = '#NAME?'; warnings.push('XLS 데이터 유효성 수식 일부를 해석하지 못해 #NAME? 조건으로 보존했습니다. 규칙을 확인하세요.'); }
            if (type === 'list' && i === 0) formula = fl & 128 ? formula.replace(/\0/g, ',') : '=' + formula.replace(/^=/, '');
            rule[`f${i+1}`] = formula;
          }
          for (const range of ranges) sheet.validations.push(relocateValidation({ ...ranges[0], ...rule }, range));
          break;
        }
        case 0x0012: if (data.length < 2) throw new Error('XLS 시트 보호 설정이 잘렸습니다.'); protectionAttrs.sheet = u16(data, 0) ? '1' : '0'; break;
        case 0x0013: if (data.length < 2) throw new Error('XLS 시트 보호 암호가 잘렸습니다.'); if (u16(data, 0)) protectionAttrs.password = u16(data, 0).toString(16).toUpperCase().padStart(4, '0'); break;
        case 0x0063: if (data.length < 2) throw new Error('XLS 개체 보호 설정이 잘렸습니다.'); protectionAttrs.objects = u16(data, 0) ? '1' : '0'; break;
        case 0x00dd: if (data.length < 2) throw new Error('XLS 시나리오 보호 설정이 잘렸습니다.'); protectionAttrs.scenarios = u16(data, 0) ? '1' : '0'; break;
        case 0x0867: if (data.length >= 19 && u16(data, 12) === 2 && u32(data, 15) === 0xffffffff) {
          if (data.length < 23) throw new Error('XLS 확장 시트 보호 설정이 잘렸습니다.');
          const flags = u32(data, 19), keys = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
          keys.forEach((key, i) => { protectionAttrs[key] = flags & (1 << i) ? '0' : '1'; });
        } break;
        case 0x0055: defColChars = u16(data, 0); break;
        case 0x0225: { const h = pt2px(u16(data, 2) / 20); if (h && h !== DEFAULT_ROW_HEIGHT) sheet.defRowH = h; break; }
        case 0x007d: {
          const c1 = u16(data, 0);
          const c2 = Math.min(u16(data, 2), 16383);
          const w = width2pxM(u16(data, 4) / 256, mdw);
          const ixfe = u16(data, 6);
          const flags = u16(data, 8);
          for (let c = c1; c <= c2 && c < c1 + 1024; c++) {
            sheet.colWidths[c] = w;
            if (flags & 1) sheet.hiddenCols[c] = true;
            if (ixfe !== 15 && xfStyle(ixfe)) sheet.colStyles[c] = xfStyle(ixfe);
          }
          break;
        }
        case 0x0208: {
          const r = u16(data, 0);
          const miy = u16(data, 6) & 0x7fff;
          const flags = u16(data, 12);
          if (flags & 0x20) sheet.hiddenRows[r] = true;
          // ROW.miyRw stores the saved height even when Excel fitted it automatically.
          // fUnsynced only distinguishes manually fixed heights; it does not gate miyRw.
          const h = pt2px(miy / 20);
          if (h !== (sheet.defRowH ?? DEFAULT_ROW_HEIGHT) || (flags & 0x40)) sheet.rowHeights[r] = h;
          if (flags & 0x40) sheet.rowManual[r] = true;
          if (flags & 0x80) { const st = xfStyle(u16(data, 14) & 0xfff); if (st) sheet.rowStyles[r] = st; }
          break;
        }
        case 0x00fd: put(u16(data, 0), u16(data, 2), u16(data, 4), sst[u32(data, 6)] ?? ''); break;
        case 0x0204: { const r = new Reader([data, ...recs[k].cont]); r.p = 6; put(u16(data, 0), u16(data, 2), u16(data, 4), r.str(2)); break; }
        case 0x00d6: { const r = new Reader([data, ...recs[k].cont]); r.p = 6; put(u16(data, 0), u16(data, 2), u16(data, 4), r.str(2)); break; } // RSTRING
        case 0x0203: put(u16(data, 0), u16(data, 2), u16(data, 4), f64(data, 6)); break;
        case 0x027e: put(u16(data, 0), u16(data, 2), u16(data, 4), rkValue(u32(data, 6))); break;
        case 0x00bd: {
          const r = u16(data, 0);
          const c0 = u16(data, 2);
          const n = (data.length - 6) / 6;
          for (let j = 0; j < n; j++) put(r, c0 + j, u16(data, 4 + j * 6), rkValue(u32(data, 6 + j * 6)));
          break;
        }
        case 0x0201: put(u16(data, 0), u16(data, 2), u16(data, 4), null); break;
        case 0x00be: {
          const r = u16(data, 0);
          const c0 = u16(data, 2);
          const n = (data.length - 6) / 2;
          for (let j = 0; j < n; j++) put(r, c0 + j, u16(data, 4 + j * 2), null);
          break;
        }
        case 0x0205: put(u16(data, 0), u16(data, 2), u16(data, 4), data[7] ? { error: ERRS[data[6]] ?? '#N/A' } : !!data[6]); break;
        case 0x04bc: {
          const cce = u16(data, 8);
          shared.push({ r1: u16(data, 0), r2: u16(data, 2), c1: data[4], c2: data[5], rgce: data.subarray(10, 10 + cce), rgcb: data.subarray(10 + cce) });
          break;
        }
        case 0x0221: {
          const cce = u16(data, 12);
          shared.push({ r1: u16(data, 0), r2: u16(data, 2), c1: data[4], c2: data[5], rgce: data.subarray(14, 14 + cce), rgcb: data.subarray(14 + cce), array: true });
          break;
        }
        case 0x0006: {
          const r = u16(data, 0);
          const c = u16(data, 2);
          const ixfe = u16(data, 4);
          let value;
          if (u16(data, 12) === 0xffff) {
            const t = data[6];
            if (t === 0) value = ''; // 결과 글자는 다음 STRING 레코드
            else if (t === 1) value = !!data[8];
            else if (t === 2) value = { error: ERRS[data[8]] ?? '#N/A' };
            else value = '';
            if (t === 0) pendingStr.push([r, c]);
          } else value = f64(data, 6);
          const cce = u16(data, 20);
          const rgce = data.subarray(22, 22 + cce);
          const rgcb = data.subarray(22 + cce);
          let text;
          try {
            if (rgce[0] === 0x01 && cce === 5) {
              // 공유 · 배열 수식: 기준 칸의 SHRFMLA/ARRAY (보통 바로 다음 레코드)를 이 칸 기준으로
              const br = u16(rgce, 1);
              const bc = u16(rgce, 3);
              let sh = shared.find((s) => s.r1 === br && s.c1 === bc);
              if (!sh) {
                for (let j = k + 1; j < Math.min(recs.length, k + 3); j++) {
                  const n = recs[j];
                  if (n.type === 0x04bc) { const ce = u16(n.data, 8); sh = { r1: u16(n.data, 0), r2: u16(n.data, 2), c1: n.data[4], c2: n.data[5], rgce: n.data.subarray(10, 10 + ce), rgcb: n.data.subarray(10 + ce) }; break; }
                  if (n.type === 0x0221) { const ce = u16(n.data, 12); sh = { r1: u16(n.data, 0), r2: u16(n.data, 2), c1: n.data[4], c2: n.data[5], rgce: n.data.subarray(14, 14 + ce), rgcb: n.data.subarray(14 + ce), array: true }; break; }
                }
                if (sh) shared.push(sh);
              }
              if (sh) text = sh.array ? decode(sh.rgce, sh.rgcb, sh.r1, sh.c1) : decode(sh.rgce, sh.rgcb, r, c);
              if (sh?.array && (r !== sh.r1 || c !== sh.c1)) { text = undefined; } // 배열 수식의 나머지 칸은 값만
            } else text = decode(rgce, rgcb, r, c);
          } catch {
            text = undefined;
            unsupported++;
          }
          put(r, c, ixfe, value, text);
          break;
        }
        case 0x0207: {
          const at = pendingStr.shift();
          if (at) {
            const r = new Reader([data, ...recs[k].cont]);
            const s = r.str(2);
            const cell = sheet.cells.getRC(at[0], at[1]);
            if (cell) { if (cell.raw.startsWith('=')) cell.cached = s; else cell.raw = textRaw(s); }
          }
          break;
        }
        case 0x00e5: {
          const n = u16(data, 0);
          for (let j = 0; j < n; j++) sheet.merges.push({ r1: u16(data, 2 + j * 8), r2: u16(data, 4 + j * 8), c1: u16(data, 6 + j * 8), c2: u16(data, 8 + j * 8) });
          break;
        }
        case 0x00ec: dgParts.push(data, ...recs[k].cont); break;
        case 0x0081: if (data.length >= 2) printFit = !!(u16(data, 0) & 0x100); break; // WsBool.fFitToPage
        case 0x0026: case 0x0027: case 0x0028: case 0x0029: {
          if (data.length < 8) break;
          const value = f64(data, 0), key = ['left', 'right', 'top', 'bottom'][type - 0x0026];
          if (Number.isFinite(value) && value >= 0) sheet.page = { ...(sheet.page ?? {}), margins: { ...(sheet.page?.margins ?? {}), [key]: value } };
          break;
        }
        case 0x0083: case 0x0084: case 0x002a: case 0x002b: {
          if (data.length < 2) break;
          const key = { 0x83: 'hCenter', 0x84: 'vCenter', 0x2a: 'headings', 0x2b: 'gridlines' }[type];
          sheet.page = { ...(sheet.page ?? {}), [key]: !!u16(data, 0) }; break;
        }
        case 0x0014: case 0x0015: {
          if (data.length < 2) break;
          sheet.page = { ...(sheet.page ?? {}), [type === 0x14 ? 'header' : 'footer']: rd(data).str(2) }; break;
        }
        case 0x00a1: { // Setup; fNoPls/fNoOrient mark undefined printer fields, not zero-valued options.
          if (data.length < 34) break;
          const g = u16(data, 10), page = { ...(sheet.page ?? {}), order: g & 1 ? 'overThenDown' : 'downThenOver' };
          fitTo = { fitW: Math.min(32767, u16(data, 6)), fitH: Math.min(32767, u16(data, 8)) };
          if (!(g & 4)) {
            const paper = u16(data, 0), scale = u16(data, 2);
            if (paper > 0 && paper < 256) page.paper = paper;
            if (scale >= 10 && scale <= 400) page.scale = scale;
            if (!(g & 0x40)) page.orientation = g & 2 ? 'portrait' : 'landscape';
          }
          for (const [key, offset] of [['header', 16], ['footer', 24]]) {
            const value = f64(data, offset);
            if (Number.isFinite(value) && value >= 0) page.margins = { ...(page.margins ?? {}), [key]: value };
          }
          sheet.page = page; break;
        }
        case 0x023e: {
          if (data.length < 10) break;
          const g = u16(data, 0); if (!(g & 2)) sheet.noGrid = true; if (!(g & 16)) sheet.noZeros = true; if (g & 8) sheet._frozen = true;
          sheet.view = { ...(sheet.view ?? {}), mode: g & 0x800 ? 'pageBreakPreview' : 'normal', headers: !!(g & 4), showFormulas: !!(g & 1), top: u16(data, 2), left: u16(data, 4) };
          if (!(g & 32) && u16(data, 6) < 64 && palette[u16(data, 6)]) sheet.view.gridColor = `#${palette[u16(data, 6)].toLowerCase()}`;
          break;
        }
        case 0x088b: if (data.length >= 16 && (u16(data, 14) & 1)) sheet.view = { ...(sheet.view ?? {}), mode: 'pageLayout' }; break;
        case 0x001b: case 0x001a: { // Horizontal/VerticalPageBreaks: each 6-byte structure is a manual break.
          if (data.length < 2) break;
          const values = [];
          for (let i = 0, n = Math.min(u16(data, 0), Math.floor((data.length - 2) / 6)); i < n; i++) { const id = u16(data, 2 + i * 6); if (id > 0 && id < (type === 0x001b ? 65536 : 256)) values.push(id); }
          if (values.length) sheet.page = { ...(sheet.page ?? {}), [type === 0x001b ? 'rowBreaks' : 'colBreaks']: [...new Set(values)].sort((a, b) => a - b) };
          break;
        }
        case 0x0041: if (sheet._frozen && data.length >= 9) {
          const rows = u16(data, 2), cols = u16(data, 0), top = rows ? sheet.view?.top ?? 0 : 0, left = cols ? sheet.view?.left ?? 0 : 0;
          sheet.freeze = { rows, cols, ...(top ? { top } : {}), ...(left ? { left } : {}) };
          sheet.view = { ...(sheet.view ?? {}), top: Math.max(top + rows, u16(data, 4)), left: Math.max(left + cols, u16(data, 6)), activePane: ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'][data[8]] ?? 'bottomRight' };
        } break;
        case 0x00a0: { // SCL is a ratio; the sheet model stores percent, as XLSX zoomScale does.
          const num = i16(data, 0); const den = i16(data, 2);
          if (num > 0 && den > 0 && num !== den) sheet.zoom = Math.max(10, Math.min(400, Math.round((num / den) * 100)));
          break;
        }
        case 0x0809: { // 시트 안의 차트 · 개체 하위 스트림은 건너뜀
          let depth = 1;
          while (depth && ++k < recs.length) { if (recs[k].type === 0x0809) depth++; else if (recs[k].type === 0x000a) depth--; }
          break;
        }
        default: break;
      }
    }
    if (protectionAttrs.sheet === '1') sheet.protect = protectFromAttrs({ objects: '0', ...protectionAttrs });
    if (printFit && fitTo) sheet.page = { ...(sheet.page ?? {}), ...fitTo };
    delete sheet._frozen;
    const defW = baseColPx(defColChars, mdw);
    if (defW !== DEFAULT_COL_WIDTH) sheet.defColW = defW;
    // 그림: 칸 기준 위치(가로 1/1024 열, 세로 1/256 행) → 시트 픽셀
    if (dgParts.length && blips.length) {
      const colW = (c) => (sheet.hiddenCols[c] ? 0 : sheet.colWidths[c] ?? sheet.defColW ?? DEFAULT_COL_WIDTH);
      const rowH = (r) => (sheet.hiddenRows[r] ? 0 : sheet.rowHeights[r] ?? sheet.defRowH ?? DEFAULT_ROW_HEIGHT);
      const xAt = (c, dx) => { let x = 0; for (let j = 0; j < c; j++) x += colW(j); return x + (Math.min(dx, 1024) / 1024) * colW(c); };
      const yAt = (r, dy) => { let y = 0; for (let j = 0; j < r; j++) y += rowH(j); return y + (Math.min(dy, 256) / 256) * rowH(r); };
      let z = 0;
      try {
        for (const pic of sheetPictures(concat(dgParts))) {
          const b = blips[pic.pib - 1];
          if (!b) { warnings.push('일부 그림(WMF · PICT 등)은 가져오지 않았습니다.'); continue; }
          const a = pic.anchor;
          const x = xAt(a.c1, a.dx1);
          const y = yAt(a.r1, a.dy1);
          sheet.images.push({
            id: `im${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, name: pic.name || '그림', z: ++z,
            x: Math.round(x), y: Math.round(y), w: Math.max(1, Math.round(xAt(a.c2, a.dx2) - x)), h: Math.max(1, Math.round(yAt(a.r2, a.dy2) - y)),
            src: `data:${b.mime};base64,${toBase64(b.bytes)}`,
            ...(pic.crop ? { crop: pic.crop } : {}),
            ...(b.emf ? { emf: `data:image/x-emf;base64,${toBase64(b.emf)}` } : {}),
          });
        }
      } catch { warnings.push('일부 그림을 읽지 못했습니다.'); }
    }
    sheets.push(sheet);
  }
  if (unsupported) warnings.push(`수식 ${unsupported}개는 해석하지 못해 계산 결과 값으로만 가져왔습니다.`);
  // 정의된 이름 (인쇄 영역 등 기본 이름은 제외)
  const outNames = [];
  for (const n of names) {
    // 인쇄 영역 · 반복 행/열: 그 시트의 페이지 설정으로.
    if (/^(_xlnm\.)?Print_(Area|Titles)$/i.test(n.name) && n.itab) {
      try {
        fctx.row = 0; fctx.col = 0;
        const sh = sheets.find((x) => x.name === allNames[n.itab - 1]);
        const areas = parsePrintAreas(decodeFormula(n.rgce, n.rgcb, fctx), sh?.name);
        if (sh && areas.length) {
          if (/Print_Area$/i.test(n.name)) sh.page = { ...(sh.page ?? {}), area: areas[0], ...(areas.length > 1 ? { areas } : {}) };
          else for (const rg of areas) {
            if (rg.c1 === 0 && rg.c2 === 16383) sh.page = { ...(sh.page ?? {}), titleRows: [rg.r1, rg.r2] };
            else if (rg.r1 === 0 && rg.r2 === 1048575) sh.page = { ...(sh.page ?? {}), titleCols: [rg.c1, rg.c2] };
          }
        }
      } catch { /* 무시 */ }
      continue;
    }
    if (n.builtin || /^_x(lfn|lnm)\./i.test(n.name)) continue; // _xlfn.IFERROR 등: 새 함수 자리 표시 이름
    try {
      fctx.row = 0; fctx.col = 0;
      const ref = fromFileFormula(decodeFormula(n.rgce, n.rgcb, fctx), {});
      outNames.push({ name: n.name, ref: `=${ref}`, sheet: n.itab ? allNames[n.itab - 1] ?? null : null, ...(n.hidden ? { hidden: true } : {}) });
    } catch { /* 해석하지 못한 이름은 건너뜀 */ }
  }
  const data = { sheets, defaultFont: wbFont, ...(date1904 ? { date1904: true } : {}) };
  if (Object.keys(calculation).length) data.calculation = { mode: 'auto', ...calculation };
  if (Object.keys(workbookProtection).length) data.props = { ...(workbookProtection.lockStructure === '1' ? { lockStructure: true } : {}), workbookProtection };
  if (calculation.iterate) warnings.push('반복 계산 설정은 보존하지만 현재 계산 엔진에서 반복 계산을 실행하지 않습니다.');
  if (calculation.fullPrecision === false) warnings.push('표시된 정밀도로 계산 설정은 보존하지만 현재 계산은 원래 숫자의 정밀도를 사용합니다.');
  if (outNames.length) data.names = outNames;
  if (!sheets.length) throw new Error('.xls 파일에 워크시트가 없습니다.');
  return { data, active: 0, warnings: [...new Set(warnings)] };
}
