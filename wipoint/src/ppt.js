// PowerPoint 97–2003 (.ppt, .pps, .pot) 읽기 — DOM 없음
// 복합 문서(cfb.js) 의 'PowerPoint Document' 스트림 → 레코드 → 영구 저장소 디렉터리 → 문서 · 마스터 · 슬라이드 · 메모
// 도형은 OfficeArt (FSP · FOPT · 앵커 · 글 상자), 그림은 'Pictures' 스트림의 BLIP. 결과는 readPptx 와 같은 pres 모양.
// 저장은 언제나 .pptx (PowerPoint 2007 이후 형식).
import { openCfb, isCfb as cfbSig } from './cfb.js';
import { inflate } from './zip.js';
import { uid, para, PRESET_ANIM, SUB_DIR, MOTION_PATHS } from './model.js';
import { THEMES, cloneTheme } from './themes.js';
import { placeInGroups } from './geom.js';

/** OLE 복합 문서인지 (.ppt 후보) */
export const isCfb = (b) => cfbSig(b);

// ───────────── 레코드 ─────────────
const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const i16 = (b, p) => { const v = u16(b, p); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, p) => (u16(b, p) | (u16(b, p + 2) << 16)) >>> 0;
const i32 = (b, p) => u32(b, p) | 0;

/** 레코드 목록: [{ type, inst, ver, off, data, kids }] (ver 0xF = 컨테이너) */
function records(buf, start = 0, end = buf.length, depth = 0) {
  const out = [];
  for (let p = start; p + 8 <= end;) {
    const vi = u16(buf, p);
    const type = u16(buf, p + 2);
    const len = u32(buf, p + 4);
    const e = Math.min(end, p + 8 + len);
    const rec = { type, inst: vi >> 4, ver: vi & 0xf, off: p, data: buf.subarray(p + 8, e), kids: null };
    if ((vi & 0xf) === 0xf && depth < 40) rec.kids = records(buf, p + 8, e, depth + 1);
    out.push(rec);
    if (len === 0 && type === 0) break;
    p += 8 + len;
  }
  return out;
}
const one = (list, type, inst) => list?.find((r) => r.type === type && (inst == null || r.inst === inst)) ?? null;
const all = (list, type) => (list ?? []).filter((r) => r.type === type);
function deep(list, type, out = []) { for (const r of list ?? []) { if (r.type === type) out.push(r); if (r.kids) deep(r.kids, type, out); } return out; }

const RT = {
  Document: 0x03e8, DocumentAtom: 0x03e9, Slide: 0x03ee, SlideAtom: 0x03ef, Notes: 0x03f0, NotesAtom: 0x03f1, Environment: 0x03f2, SlidePersistAtom: 0x03f3,
  MainMaster: 0x03f8, SlideShowSlideInfoAtom: 0x03f9, ExObjList: 0x0409, PPDrawingGroup: 0x040b, PPDrawing: 0x040c,
  ColorSchemeAtom: 0x07f0, FontCollection: 0x07d5, FontEntityAtom: 0x0fb7, TextHeaderAtom: 0x0f9f, TextCharsAtom: 0x0fa0, StyleTextPropAtom: 0x0fa1,
  TextMasterStyleAtom: 0x0fa3, TextBytesAtom: 0x0fa8, OutlineTextRefAtom: 0x0f9e, SlideListWithText: 0x0ff0, UserEditAtom: 0x0ff5, CurrentUserAtom: 0x0ff6,
  PersistDirectoryAtom: 0x1772, PlaceholderAtom: 0x0bc3, InteractiveInfo: 0x0ff2, InteractiveInfoAtom: 0x0ff3, TxInteractiveInfoAtom: 0x0fdf,
  ExHyperlink: 0x0fd7, ExHyperlinkAtom: 0x0fd3, CString: 0x0fba, HeadersFooters: 0x0fd9, HeadersFootersAtom: 0x0fda, TextRulerAtom: 0x0fa6,
  AnimationInfo: 0x1014, AnimationInfoAtom: 0x0ff1, ExOleObjStg: 0x1011, ExOleEmbed: 0x0fcc, ExOleObjAtom: 0x0fc3, ExObjRefAtom: 0x0bc1,
};
const ART = { DggContainer: 0xf000, BStoreContainer: 0xf001, DgContainer: 0xf002, SpgrContainer: 0xf003, SpContainer: 0xf004, FBSE: 0xf007, FSPGR: 0xf009, FSP: 0xf00a, FOPT: 0xf00b, ClientTextbox: 0xf00d, ChildAnchor: 0xf00f, ClientAnchor: 0xf010, ClientData: 0xf011, TertiaryFOPT: 0xf122 };

// ───────────── 그림 저장소 ─────────────
function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
/** BLIP 레코드 → { mime, bytes } */
function blipData(r) {
  const twoUid = r.inst & 1;
  const d = r.data;
  if (r.type >= 0xf01a && r.type <= 0xf01c) { // EMF · WMF · PICT: 메타파일 머리말 34바이트 + (보통 deflate) 데이터
    const h = twoUid ? 32 : 16;
    if (d.length < h + 34) return null;
    const size = u32(d, h); const cbSave = u32(d, h + 28); const comp = d[h + 32];
    let body = d.subarray(h + 34, h + 34 + cbSave);
    if (comp === 0) { try { body = inflate(body.subarray(2), size); } catch { try { body = inflate(body, size); } catch { return null; } } }
    if (r.type === 0xf01c) return null; // PICT (Mac) — 그리지 못함
    return { mime: r.type === 0xf01a ? 'image/x-emf' : 'image/x-wmf', bytes: body };
  }
  const skip = (twoUid ? 32 : 16) + 1;
  const bytes = d.subarray(skip);
  if (r.type === 0xf01e) return { mime: 'image/png', bytes };
  if (r.type === 0xf01d || r.type === 0xf02a) return { mime: 'image/jpeg', bytes };
  if (r.type === 0xf029) return { mime: 'image/tiff', bytes };
  if (r.type === 0xf01f) { // DIB → BMP
    const hdr = u32(bytes, 0); const bits = u16(bytes, 14); const colors = u32(bytes, 32) || (bits <= 8 ? 1 << bits : 0);
    const offBits = 14 + hdr + colors * 4; const size = bytes.length + 14;
    const out = new Uint8Array(size);
    out.set([0x42, 0x4d, size & 255, (size >> 8) & 255, (size >> 16) & 255, (size >>> 24) & 255, 0, 0, 0, 0, offBits & 255, (offBits >> 8) & 255, 0, 0]);
    out.set(bytes, 14);
    return { mime: 'image/bmp', bytes: out };
  }
  return null;
}

// ───────────── 도형 속성 (FOPT) ─────────────
function fopt(rec) {
  const props = new Map();
  if (!rec) return props;
  const d = rec.data; const n = rec.inst;
  let cpos = n * 6;
  for (let i = 0; i < n && i * 6 + 6 <= d.length; i++) {
    const id = u16(d, i * 6); const val = u32(d, i * 6 + 2);
    const pid = id & 0x3fff; const complex = !!(id & 0x8000);
    const p = { val };
    if (complex) {
      // 배열 속성 (IMsoArray): 크기에 6바이트 머리말이 빠진 파일이 있음 (nElems × 원소 크기 = 크기)
      let size = val;
      const ARRAYS = [0x0145, 0x0146, 0x0151, 0x0152, 0x0153, 0x0154, 0x0155, 0x0156, 0x0157, 0x0158, 0x0197, 0x01cf, 0x0340, 0x0341];
      if (ARRAYS.includes(pid) && cpos + 6 <= d.length) {
        const ne = u16(d, cpos); const cb = u16(d, cpos + 4);
        const es = cb === 0xfff0 ? 4 : cb;
        if (ne * es === val && val) size = val + 6;
      }
      p.data = d.subarray(cpos, cpos + size); cpos += size;
    }
    props.set(pid, p);
  }
  return props;
}
const fixed = (v) => (v | 0) / 65536;
/** TextRulerAtom → { levels: [{ left, indent }] (마스터 단위), tabs: [{ pos, algn }] } */
function parseRuler(d) {
  if (!d || d.length < 4) return null;
  const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
  const m = dv.getUint32(0, true);
  let p = 4;
  const rd = () => { if (p + 2 > d.length) return 0; const v = dv.getInt16(p, true); p += 2; return v; };
  const out = { levels: [], tabs: [] };
  if (m & 0x2) rd(); // cLevels
  if (m & 0x1) out.tabSize = rd();
  if (m & 0x4) {
    const n = rd();
    for (let i = 0; i < n && p + 4 <= d.length; i++) { const pos = rd(); const t = rd(); out.tabs.push({ pos, algn: ['l', 'ctr', 'r', 'dec'][t] ?? 'l' }); }
  }
  for (let i = 0; i < 5; i++) {
    const lv = {};
    if (m & (0x8 << i)) lv.left = rd();
    if (m & (0x100 << i)) lv.indent = rd();
    out.levels.push(lv.left != null || lv.indent != null ? lv : null);
  }
  return out;
}
/** 참/거짓 묶음 속성: bit 의 값 (use 비트가 켜져 있을 때만), 없으면 dflt */
function boolProp(props, pid, bit, dflt) {
  const p = props.get(pid);
  if (!p) return dflt;
  const use = 1 << (bit + 16);
  return p.val & use ? !!(p.val & (1 << bit)) : dflt;
}

/** 속성 겹치기 (b 가 a 를 덮음). 참/거짓 묶음 속성(pid & 0x3f = 0x3f)은 use 비트가 켜진 비트만 덮음 */
function mergeProps(a, b) {
  const out = new Map(a);
  for (const [pid, p] of b) {
    const q = out.get(pid);
    if ((pid & 0x3f) === 0x3f && q) {
      const use = (p.val >>> 16) & 0xffff;
      const val = ((q.val >>> 16) | use) << 16 | ((q.val & 0xffff) & ~use) | (p.val & use);
      out.set(pid, { val: val >>> 0 });
    } else out.set(pid, p);
  }
  return out;
}

// MSOSPT 모양 번호 → DrawingML 미리 설정 모양
const SPT = {
  1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon', 10: 'octagon', 11: 'plus', 12: 'star5', 13: 'rightArrow', 14: 'rightArrow', 15: 'homePlate', 16: 'cube', 17: 'wedgeRoundRectCallout', 18: 'star16', 19: 'arc', 20: 'line', 21: 'plaque', 22: 'can', 23: 'donut',
  32: 'straightConnector1', 33: 'bentConnector2', 34: 'bentConnector3', 35: 'bentConnector4', 36: 'bentConnector5', 37: 'curvedConnector2', 38: 'curvedConnector3', 39: 'curvedConnector4', 40: 'curvedConnector5',
  41: 'callout1', 42: 'callout2', 43: 'callout3', 44: 'accentCallout1', 45: 'accentCallout2', 46: 'accentCallout3', 47: 'borderCallout1', 48: 'borderCallout2', 49: 'borderCallout3', 50: 'accentBorderCallout1', 51: 'accentBorderCallout2', 52: 'accentBorderCallout3',
  53: 'ribbon', 54: 'ribbon2', 55: 'chevron', 56: 'pentagon', 57: 'noSmoking', 58: 'star8', 59: 'star16', 60: 'star32', 61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout', 64: 'wave', 65: 'foldedCorner', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow',
  71: 'irregularSeal1', 72: 'irregularSeal2', 73: 'lightningBolt', 74: 'heart', 75: 'rect', 76: 'quadArrow', 77: 'leftArrowCallout', 78: 'rightArrowCallout', 79: 'upArrowCallout', 80: 'downArrowCallout', 81: 'leftRightArrowCallout', 82: 'upDownArrowCallout', 83: 'quadArrowCallout',
  84: 'bevel', 85: 'leftBracket', 86: 'rightBracket', 87: 'leftBrace', 88: 'rightBrace', 89: 'leftUpArrow', 90: 'bentUpArrow', 91: 'bentArrow', 92: 'star24', 93: 'stripedRightArrow', 94: 'notchedRightArrow', 95: 'blockArc', 96: 'smileyFace', 97: 'verticalScroll', 98: 'horizontalScroll', 99: 'circularArrow', 100: 'circularArrow',
  101: 'uturnArrow', 102: 'curvedRightArrow', 103: 'curvedLeftArrow', 104: 'curvedUpArrow', 105: 'curvedDownArrow', 106: 'cloudCallout', 107: 'ellipseRibbon', 108: 'ellipseRibbon2',
  109: 'flowChartProcess', 110: 'flowChartDecision', 111: 'flowChartInputOutput', 112: 'flowChartPredefinedProcess', 113: 'flowChartInternalStorage', 114: 'flowChartDocument', 115: 'flowChartMultidocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation', 118: 'flowChartManualInput', 119: 'flowChartManualOperation', 120: 'flowChartConnector',
  121: 'flowChartPunchedCard', 122: 'flowChartPunchedTape', 123: 'flowChartSummingJunction', 124: 'flowChartOr', 125: 'flowChartCollate', 126: 'flowChartSort', 127: 'flowChartExtract', 128: 'flowChartMerge', 129: 'flowChartOfflineStorage', 130: 'flowChartOnlineStorage', 131: 'flowChartMagneticTape', 132: 'flowChartMagneticDisk', 133: 'flowChartMagneticDrum', 134: 'flowChartDisplay', 135: 'flowChartDelay',
  176: 'flowChartAlternateProcess', 177: 'flowChartOffpageConnector', 178: 'callout1', 179: 'accentCallout1', 180: 'borderCallout1', 181: 'accentBorderCallout1', 182: 'leftRightUpArrow', 183: 'sun', 184: 'moon', 185: 'bracketPair', 186: 'bracePair', 187: 'star4', 188: 'doubleWave',
  189: 'actionButtonBlank', 190: 'actionButtonHome', 191: 'actionButtonHelp', 192: 'actionButtonInformation', 193: 'actionButtonForwardNext', 194: 'actionButtonBackPrevious', 195: 'actionButtonEnd', 196: 'actionButtonBeginning', 197: 'actionButtonReturn', 198: 'actionButtonDocument', 199: 'actionButtonSound', 200: 'actionButtonMovie',
  202: 'rect',
};
const LINE_SHAPES = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40]);
const DASH = ['solid', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot'];
const ARROW = [null, 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];

/** PowerPoint 97 슬라이드 전환 종류 → WIPOINT 전환 */
const TRANS = { 0: 'cut', 1: 'random', 2: 'blinds', 3: 'checker', 4: 'cover', 5: 'dissolve', 6: 'fade', 7: 'uncover', 8: 'randomBars', 9: 'strips', 10: 'wipe', 11: 'zoom', 13: 'split', 17: 'diamond', 18: 'plus', 19: 'clock', 20: 'push', 21: 'comb', 22: 'zoom', 23: 'fade', 26: 'clock', 27: 'circle' };

/**
 * .ppt 바이트 → { pres, warnings }
 */
export function readPpt(bytes) {
  const cfb = openCfb(bytes);
  const warnings = new Set();
  if (cfb.names.some((n) => /EncryptedSummary|EncryptionInfo/i.test(n))) throw new Error('암호로 보호된 프레젠테이션은 열 수 없습니다. PowerPoint 에서 암호를 해제한 뒤 열어 주세요.');
  // PowerPoint 95 와 함께 저장한 파일: 97 형식은 PP97_DUALSTORAGE 안에
  const dual = cfb.names.includes('PP97_DUALSTORAGE/PowerPoint Document') ? 'PP97_DUALSTORAGE/' : '';
  const doc = cfb.stream(`${dual}PowerPoint Document`);
  if (!doc) throw new Error(cfb.names.includes('Text_Content') || cfb.names.includes('Header') ? 'PowerPoint 95 이전 형식(.ppt)은 열 수 없습니다. PowerPoint 에서 .pptx 로 저장한 뒤 열어 주세요.' : 'PowerPoint 97-2003 프레젠테이션(.ppt)이 아닙니다');
  const pics = cfb.stream(`${dual}Pictures`);

  // ── 영구 저장소 디렉터리: 가장 최근 편집(UserEditAtom)부터 이전 편집 순으로
  const persist = new Map();
  let docRef = null;
  const cu = cfb.stream(`${dual}Current User`);
  let editOff = cu && cu.length >= 20 ? u32(cu, 16) : null;
  const readEdit = (o) => {
    if (o == null || o + 8 > doc.length || u16(doc, o + 2) !== RT.UserEditAtom) return null;
    const d = doc.subarray(o + 8, o + 8 + u32(doc, o + 4));
    return { lastEdit: u32(d, 8), dirOff: u32(d, 12), docRef: u32(d, 16), encrypt: d.length >= 32 ? u32(d, 28) : 0 };
  };
  let ue = readEdit(editOff);
  if (!ue) {
    // Current User 가 없거나 틀리면: 스트림에서 마지막 UserEditAtom 찾기
    for (const r of records(doc)) if (r.type === RT.UserEditAtom) editOff = r.off;
    ue = readEdit(editOff);
  }
  if (!ue) throw new Error('PowerPoint 문서 구조를 읽지 못했습니다 (손상된 파일일 수 있습니다)');
  if (ue.encrypt) throw new Error('암호로 보호된 프레젠테이션은 열 수 없습니다. PowerPoint 에서 암호를 해제한 뒤 열어 주세요.');
  for (let guard = 0; ue && guard < 1000; guard++) {
    docRef ??= ue.docRef;
    const o = ue.dirOff;
    if (o + 8 <= doc.length && u16(doc, o + 2) === RT.PersistDirectoryAtom) {
      const d = doc.subarray(o + 8, o + 8 + u32(doc, o + 4));
      for (let p = 0; p + 4 <= d.length;) {
        const v = u32(d, p); const start = v & 0xfffff; const cnt = v >>> 20;
        p += 4;
        for (let k = 0; k < cnt && p + 4 <= d.length; k++, p += 4) if (!persist.has(start + k)) persist.set(start + k, u32(d, p));
      }
    }
    if (!ue.lastEdit) break;
    ue = readEdit(ue.lastEdit);
  }
  const recAt = (ref) => {
    const o = persist.get(ref);
    if (o == null || o + 8 > doc.length) return null;
    return records(doc, o, Math.min(doc.length, o + 8 + u32(doc, o + 4)))[0] ?? null;
  };
  const docRec = recAt(docRef);
  // 개요 글 (제목 · 본문 개체 틀) 의 PowerPoint 2000+ 단락 확장: 슬라이드 id + 글 종류 → 번호 매기기
  const outline9 = new Map();
  for (const blob of deep(docRec?.kids, 0x138b)) {
    for (const c of deep(records(blob.data), 0x0fae)) {
      let key = null;
      for (const k of c.kids ?? []) {
        if (k.type === 0x0faf && k.data.length >= 8) key = `${u32(k.data, 0)}:${u32(k.data, 4)}`;
        else if (k.type === 0x0fac && key) { outline9.set(key, parseStyle9(k.data)); key = null; }
      }
    }
  }
  if (!docRec?.kids) throw new Error('PowerPoint 문서 구조를 읽지 못했습니다 (손상된 파일일 수 있습니다)');

  // ── 문서 정보
  const da = one(docRec.kids, RT.DocumentAtom);
  const size = da ? { w: Math.round(i32(da.data, 0) / 6), h: Math.round(i32(da.data, 4) / 6) } : { w: 960, h: 720 };
  if (!(size.w > 50 && size.h > 50)) { size.w = 960; size.h = 720; }
  const env = one(docRec.kids, RT.Environment);
  const fonts = [];
  for (const fe of deep(env?.kids, RT.FontEntityAtom)) {
    let s = '';
    for (let i = 0; i < 32; i++) { const c = u16(fe.data, i * 2); if (!c) break; s += String.fromCharCode(c); }
    fonts[fe.inst] = s;
  }
  const fontName = (i) => fonts[i] ?? null;
  // 문서 기본 글자 서식 (Environment 의 TextMasterStyleAtom = 'other')
  const envStyles = parseMasterStyles(env?.kids);
  // 그림 저장소
  const dgg = one(deep(one(docRec.kids, RT.PPDrawingGroup)?.kids, ART.DggContainer), ART.DggContainer) ?? deep(one(docRec.kids, RT.PPDrawingGroup)?.kids, ART.DggContainer)[0];
  const blips = [];
  for (const fbse of deep(dgg?.kids, ART.FBSE)) {
    const d = fbse.data;
    const cbName = d[33];
    let rec = d.length > 36 + cbName + 8 ? records(d, 36 + cbName)[0] : null;
    if (!rec && pics && u32(d, 20) !== 0xffffffff) { const fo = u32(d, 28); if (fo + 8 <= pics.length) rec = records(pics, fo, Math.min(pics.length, fo + 8 + u32(pics, fo + 4)))[0]; }
    let m = null;
    try { m = rec ? blipData(rec) : null; } catch { m = null; }
    blips.push(m);
  }
  // 문서 기본 속성 (DGG FOPT) 은 새 도형용 기본값일 뿐 기존 도형이 물려받지 않음 → 명세 기본값 사용
  const defaultProps = new Map();
  const media = {};
  const mediaOf = new Map();
  // 소리 모음 (SoundCollectionContainer 0x07E4 → Sound 0x07E6: 이름 0 · 확장자 1 · id 2 · 자료 0x07E7) — 애니메이션 · 전환 소리
  const soundById = new Map();
  for (const sc of deep(docRec?.kids, 0x07e4)) {
    for (const sn of (sc.kids ?? []).filter((k) => k.type === 0x07e6)) {
      const str = (inst) => { const c = one(sn.kids, RT.CString, inst); let t = ''; for (let i = 0; c && i + 1 < c.data.length; i += 2) t += String.fromCharCode(u16(c.data, i)); return t.replace(/\0+$/, ''); };
      const id = Number(str(2));
      const blob = one(sn.kids, 0x07e7);
      if (Number.isFinite(id)) soundById.set(id, { name: str(0), ext: str(1), data: blob?.data ?? null });
    }
  }
  const soundMediaOf = new Map();
  /** 소리 id → { sound: 기본 소리 키 } 또는 { soundMedia, soundName } */
  const soundFor = (id) => {
    const sn = soundById.get(id);
    if (!sn) return null;
    const key = BUILTIN_SOUND[sn.name.toLowerCase().replace(/[^a-z]/g, '')];
    if (key) return { sound: key };
    if (!sn.data?.length) return null;
    if (!soundMediaOf.has(id)) {
      const ext = (sn.ext || '.wav').replace(/^\./, '').toLowerCase();
      const mime = { wav: 'audio/wav', mp3: 'audio/mpeg', mid: 'audio/midi', midi: 'audio/midi', aif: 'audio/aiff', aiff: 'audio/aiff', wma: 'audio/x-ms-wma' }[ext] ?? 'audio/wav';
      const mid = `snd${id}${uid('').slice(-4)}`;
      media[mid] = `data:${mime};base64,${b64(sn.data)}`;
      soundMediaOf.set(id, mid);
    }
    return { soundMedia: soundMediaOf.get(id), soundName: `${sn.name}${sn.ext || '.wav'}` };
  };
  const blipMedia = (pib) => {
    if (!pib || !blips[pib - 1]) return null;
    if (mediaOf.has(pib)) return mediaOf.get(pib);
    const b = blips[pib - 1];
    const id = `m${pib}${uid('').slice(-4)}`;
    media[id] = `data:${b.mime};base64,${b64(b.bytes)}`;
    mediaOf.set(pib, id);
    return id;
  };
  // 하이퍼링크 목록
  const links = new Map();
  for (const h of deep(one(docRec.kids, RT.ExObjList)?.kids, RT.ExHyperlink)) {
    const at = one(h.kids, RT.ExHyperlinkAtom);
    const str = (inst) => { const c = one(h.kids, RT.CString, inst); let s = ''; for (let i = 0; c && i + 1 < c.data.length; i += 2) s += String.fromCharCode(u16(c.data, i)); return s; };
    if (at) links.set(u32(at.data, 0), { target: str(1), loc: str(3) });
  }

  // ── 슬라이드 목록 (SlideListWithText): 영구 id · 개요 글
  const lists = all(docRec.kids, RT.SlideListWithText);
  const listOf = (inst) => {
    const out = [];
    const l = lists.find((x) => x.inst === inst);
    for (const r of l?.kids ?? []) {
      if (r.type === RT.SlidePersistAtom) out.push({ ref: u32(r.data, 0), slideId: u32(r.data, 12), texts: [] });
      else if (out.length && r.type === RT.TextHeaderAtom) out[out.length - 1].texts.push([r]);
      else if (out.length && out[out.length - 1].texts.length) out[out.length - 1].texts[out[out.length - 1].texts.length - 1].push(r);
    }
    return out;
  };
  const slideList = listOf(0);
  const masterList = listOf(1);
  const notesList = listOf(2);

  // ── 마스터
  const masters = new Map(); // slideId · persist ref → { styles, scheme, bg, decor }
  const readMaster = (entry) => {
    const rec = recAt(entry.ref);
    if (!rec?.kids) return null;
    const m = { rec, styles: rec.type === RT.MainMaster ? parseMasterStyles(rec.kids) : null, scheme: schemeOf(rec), atom: one(rec.kids, RT.SlideAtom) };
    // 개체 틀 글 종류별 눈금자 (슬라이드 개체 틀이 눈금자를 따로 갖지 않으면 이것을 씀)
    m.rulers = {};
    for (const tb of deep(one(rec.kids, RT.PPDrawing)?.kids, ART.ClientTextbox)) {
      const h = one(tb.kids, RT.TextHeaderAtom); const r = one(tb.kids, RT.TextRulerAtom);
      if (h && r) m.rulers[u32(h.data, 0)] ??= parseRuler(r.data);
    }
    // 마스터 개체 틀 속성 (슬라이드 개체 틀이 채우기 · 선을 따로 정하지 않으면 이것을 따름)
    m.phProps = {};
    for (const sp of deep(one(rec.kids, RT.PPDrawing)?.kids, ART.SpContainer)) {
      const ph = deep(one(sp.kids, ART.ClientData)?.kids, RT.PlaceholderAtom)[0];
      if (ph) m.phProps[ph.data[4]] ??= mergeProps(fopt(one(sp.kids, ART.FOPT)), fopt(one(sp.kids, ART.TertiaryFOPT)));
    }
    masters.set(entry.slideId, m);
    masters.set(entry.ref, m);
    return m;
  };
  for (const e of masterList) readMaster(e);
  // 제목 마스터는 MainMaster 를 가리킴 → 글 스타일을 거기서
  const mainMaster = [...masters.values()].find((m) => m.styles);
  for (const m of masters.values()) {
    if (!m.styles) {
      const ref = m.atom ? u32(m.atom.data, 12) : 0;
      const base = masters.get(ref) ?? mainMaster;
      m.styles = base?.styles ?? envStyles;
      m.base = base;
    }
  }

  // ── 색 구성 (8색: 배경, 글자 · 선, 그림자, 제목 글자, 채우기, 강조, 강조 · 하이퍼링크, 강조 · 열어 본 하이퍼링크)
  function schemeOf(rec) {
    const cs = all(rec?.kids, RT.ColorSchemeAtom).find((r) => r.inst === 1) ?? one(rec?.kids, RT.ColorSchemeAtom);
    if (!cs) return null;
    const out = [];
    for (let i = 0; i < 8; i++) out.push(`#${[0, 1, 2].map((k) => cs.data[i * 4 + k].toString(16).padStart(2, '0')).join('').toUpperCase()}`);
    return out;
  }
  const baseScheme = mainMaster?.scheme ?? ['#FFFFFF', '#000000', '#808080', '#000000', '#BBE0E3', '#333399', '#009999', '#99CC00'];
  const theme = cloneTheme(THEMES[0]);
  theme.name = 'ppt97'; theme.label = 'PowerPoint 97-2003';
  Object.assign(theme.colors, { lt1: baseScheme[0], dk1: baseScheme[1], lt2: baseScheme[2], dk2: baseScheme[3], accent1: baseScheme[4], accent2: baseScheme[5], accent3: baseScheme[6], accent4: baseScheme[7], hlink: baseScheme[6], folHlink: baseScheme[7] });
  const titleFont = mainMaster?.styles?.[0]?.[0]?.cf?.font;
  const bodyFont = mainMaster?.styles?.[1]?.[0]?.cf?.font;
  if (fontName(titleFont)) theme.fonts.major = fontName(titleFont);
  if (fontName(bodyFont)) theme.fonts.minor = fontName(bodyFont);

  /** 색 (OfficeArtCOLORREF) → '#RRGGBB' */
  const hexOf = (r, g, b) => `#${[r, g, b].map((x) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
  /** 색 (OfficeArtCOLORREF) → '#RRGGBB'. rel = { fill, line } (fSysIndex 0xF0… = 이 도형의 채우기 · 선 색을 바꾼 색) */
  const colorRef = (v, scheme, rel) => {
    const flags = v >>> 24;
    if (flags & 0x08) return (scheme ?? baseScheme)[v & 0xff] ?? '#000000';
    if (flags & 0x10) {
      const idx = v & 0xff;
      const base = { 0xf0: rel?.fill, 0xf1: rel?.line ?? rel?.fill, 0xf2: rel?.line, 0xf3: '#808080', 0xf4: rel?.fill, 0xf5: rel?.back, 0xf6: rel?.line, 0xf7: rel?.fill ?? rel?.line }[idx];
      if (!base) return idx === 0xf0 && !rel ? null : '#808080';
      const mod = (v >> 8) & 0x0f; const k = ((v >> 16) & 0xff) / 255;
      let [r, g, b] = [1, 3, 5].map((i) => parseInt(base.slice(i, i + 2), 16));
      if (mod === 1) [r, g, b] = [r * k, g * k, b * k];
      else if (mod === 2) [r, g, b] = [255 - (255 - r) * k, 255 - (255 - g) * k, 255 - (255 - b) * k];
      else if (mod === 3) [r, g, b] = [r + k * 255, g + k * 255, b + k * 255];
      else if (mod === 4) [r, g, b] = [r - k * 255, g - k * 255, b - k * 255];
      else if (mod === 5) [r, g, b] = [k * 255 - r, k * 255 - g, k * 255 - b];
      else if (mod === 6) { const y = 0.299 * r + 0.587 * g + 0.114 * b > 127 ? 255 : 0; r = g = b = y; }
      if ((v >> 8) & 0x20) [r, g, b] = [255 - r, 255 - g, 255 - b];
      return hexOf(r, g, b);
    }
    return hexOf(v & 255, (v >> 8) & 255, (v >> 16) & 255);
  };
  /** 글자 색 (ColorIndexStruct) */
  const textColor = (c, scheme) => (!c ? null : c.index === 0xfe ? c.hex : c.index < 8 ? (scheme ?? baseScheme)[c.index] : null);

  // ── 도형 읽기
  const readShapes = (rec, ctx) => {
    const out = [];
    const dg = deep(one(rec.kids, RT.PPDrawing)?.kids, ART.DgContainer)[0];
    if (!dg) return { objects: out, bg: undefined };
    let bg;
    const spgr = one(dg.kids, ART.SpgrContainer);
    const bgSp = one(dg.kids, ART.SpContainer);
    if (bgSp) { const f = fillOf(fopt(one(bgSp.kids, ART.FOPT)), ctx, true); if (f !== undefined) bg = f; }
    walkGroup(spgr, null, ctx, out, null);
    return { objects: out, bg };
  };
  function walkGroup(gr, xf, ctx, out, gid) {
    const kids = gr?.kids ?? [];
    let groupXf = xf;
    kids.forEach((k, i) => {
      if (k.type === ART.SpContainer) {
        const fsp = one(k.kids, ART.FSP);
        if (i === 0 && fsp && u32(fsp.data, 4) & 0x1) {
          // 그룹 자신: 그룹 좌표계 (FSPGR) + 앵커 (바깥 그룹 좌표) + 회전 · 대칭 → 자식 배치에 씀 (pptx 와 같은 placeInGroups)
          const g = one(k.kids, ART.FSPGR);
          const props = mergeProps(fopt(one(k.kids, ART.FOPT)), fopt(one(k.kids, ART.TertiaryFOPT)));
          const rot = props.get(0x0004) ? fixed(props.get(0x0004).val) : 0;
          const a = rawAnchor(k, xf, rot);
          const fl = u32(fsp.data, 4);
          if (g && a) groupXf = { x: a.x, y: a.y, w: a.w, h: a.h, chx: i32(g.data, 0), chy: i32(g.data, 4), chw: (i32(g.data, 8) - i32(g.data, 0)) || 1, chh: (i32(g.data, 12) - i32(g.data, 4)) || 1, rot: ((rot % 360) + 360) % 360, ...(fl & 0x40 ? { flipH: true } : {}), ...(fl & 0x80 ? { flipV: true } : {}), up: xf };
          return;
        }
        try {
          const o = readShape(k, groupXf, ctx);
          if (o) {
            if (gid) o.grp = gid;
            out.push(o);
            // 애니메이션 대상 (도형 id → 개체) · PowerPoint 97 빌드 정보 (AnimationInfoAtom)
            if (fsp) { ctx.spMap?.set(u32(fsp.data, 0), [o.id]); if (o.ph) ctx.phOf?.set(u32(fsp.data, 0), o.ph); }
            const ai = deep(one(k.kids, ART.ClientData)?.kids, RT.AnimationInfoAtom)[0];
            if (ai && ctx.legacyAnims) ctx.legacyAnims.push({ obj: o.id, d: ai.data });
          }
        } catch (e) { warnings.add(`일부 도형을 읽지 못했습니다: ${e.message}`); }
      } else if (k.type === ART.SpgrContainer) {
        const start = out.length;
        const gfsp = one(k.kids[0]?.kids, ART.FSP);
        walkGroup(k, groupXf, ctx, out, gid ?? (xf ? gid : uid('g')));
        // 그룹 전체에 건 애니메이션 → 그룹의 모든 개체
        if (gfsp) ctx.spMap?.set(u32(gfsp.data, 0), out.slice(start).map((o) => o.id));
        const ai = deep(one(k.kids[0]?.kids, ART.ClientData)?.kids, RT.AnimationInfoAtom)[0];
        if (ai && ctx.legacyAnims && out.length > start) ctx.legacyAnims.push({ obj: out[start].id, objs: out.slice(start).map((o) => o.id), d: ai.data });
      }
    });
  }
  /** 앵커 → 사각형: 그룹 안이면 그룹 좌표 단위 (ChildAnchor), 아니면 슬라이드 px (ClientAnchor).
   *  바이너리 형식은 45°~135° · 225°~315° 회전이면 앵커가 90° 돌려 저장되어 있어 가로 · 세로를 바꿈 */
  function rawAnchor(sp, xf, rot = 0) {
    const ca = one(sp.kids, ART.ClientAnchor);
    const ch = one(sp.kids, ART.ChildAnchor);
    let r = null;
    if (ch && xf) {
      const l = i32(ch.data, 0); const t = i32(ch.data, 4); const rr = i32(ch.data, 8); const b = i32(ch.data, 12);
      r = { x: l, y: t, w: rr - l, h: b - t };
    } else if (ca) {
      const d = ca.data;
      let t; let l; let rr; let b;
      if (d.length >= 16) { t = i32(d, 0); l = i32(d, 4); rr = i32(d, 8); b = i32(d, 12); } else { t = i16(d, 0); l = i16(d, 2); rr = i16(d, 4); b = i16(d, 6); }
      r = { x: l / 6, y: t / 6, w: (rr - l) / 6, h: (b - t) / 6 };
    }
    if (!r) return null;
    const nr = ((rot % 360) + 360) % 360;
    if ((nr >= 45 && nr < 135) || (nr >= 225 && nr < 315)) {
      const cx = r.x + r.w / 2; const cy = r.y + r.h / 2;
      r = { x: cx - r.h / 2, y: cy - r.w / 2, w: r.h, h: r.w };
    }
    return r;
  }
  function fillOf(props, ctx, isBg = false) {
    // 배경 도형은 문서 기본 속성(DGG)을 따르지 않음 (기본: 채움, 배경 색)
    const p = (pid) => props.get(pid) ?? (isBg ? undefined : defaultProps.get(pid));
    const filled = boolProp(props, 0x01bf, 4, isBg ? true : boolProp(defaultProps, 0x01bf, 4, true));
    if (!filled) return isBg ? undefined : null;
    const type = p(0x0180)?.val ?? 0;
    const c1 = p(0x0181) ? colorRef(p(0x0181).val, ctx.scheme) ?? '#FFFFFF' : isBg ? (ctx.scheme ?? baseScheme)[0] : '#FFFFFF';
    const c2 = p(0x0183) ? colorRef(p(0x0183).val, ctx.scheme, { fill: c1 }) ?? '#FFFFFF' : '#FFFFFF';
    const op = p(0x0182) ? fixed(p(0x0182).val) : 1;
    const withAlpha = (c) => (op < 1 && c ? `${c}:a${Math.round(op * 1000) / 10}` : c);
    if (type === 2 || type === 3) {
      const m = blipMedia(p(0x0186)?.val);
      if (m) return { type: 'image', media: m, ...(type === 2 ? { tile: true } : {}) };
      return { type: 'solid', color: withAlpha(c1) };
    }
    if (type === 1) return { type: 'pattern', prst: 'pct50', fg: c1, bg: c2 };
    if (type >= 4 && type <= 8) {
      // 그라데이션: 두 색 (focus 50 이면 가운데가 두 번째 색)
      const ang = p(0x018b) ? fixed(p(0x018b).val) : 0;
      const focus = p(0x018c) ? (p(0x018c).val | 0) : 0;
      // fillFocus = 두 번째 색(뒤 색)의 위치 % (음수면 두 색을 바꿈): 0 → 뒤 색에서 앞 색, 100 → 앞 색에서 뒤 색, 50 → 가운데가 뒤 색
      const [a, b] = focus < 0 ? [c2, c1] : [c1, c2];
      const fp = Math.min(1, Math.abs(focus) / 100);
      const stops = fp <= 0.001 ? [[0, b], [1, a]] : fp >= 0.999 ? [[0, a], [1, b]] : [[0, a], [fp, b], [1, a]];
      const out = { type: 'gradient', stops, angle: ((90 - ang) % 360 + 360) % 360 };
      if (type === 5 || type === 6) out.path = true;
      return out;
    }
    if (type === 9) return isBg ? undefined : null;
    return { type: 'solid', color: withAlpha(c1) };
  }
  function lineOf(props, ctx, isLine) {
    const p = (pid) => props.get(pid) ?? defaultProps.get(pid);
    const on = boolProp(props, 0x01ff, 3, boolProp(defaultProps, 0x01ff, 3, true));
    if (!on) return null;
    const color = p(0x01c0) ? colorRef(p(0x01c0).val, ctx.scheme, { fill: props.get(0x0181) ? colorRef(props.get(0x0181).val, ctx.scheme) : null }) : '#000000';
    const w = p(0x01cb) ? p(0x01cb).val / 9525 : 0.75 * (4 / 3);
    const out = { color: color ?? '#000000', width: Math.max(0.5, w) };
    const dash = DASH[p(0x01ce)?.val ?? 0];
    if (dash && dash !== 'solid') out.dash = dash;
    if (isLine) { const h = ARROW[p(0x01d0)?.val ?? 0]; const t = ARROW[p(0x01d1)?.val ?? 0]; if (h) out.head = h; if (t) out.tail = t; }
    return out;
  }
  /** 사용자 지정 도형 (pVertices · pSegmentInfo) → 경로 */
  function customGeom(props) {
    const v = props.get(0x0145)?.data; const s = props.get(0x0146)?.data;
    if (!v || v.length < 6) return null;
    const n = u16(v, 0); const cb = u16(v, 4) === 0xfff0 ? 4 : u16(v, 4);
    const pts = [];
    for (let i = 0; i < n && 6 + (i + 1) * cb <= v.length; i++) pts.push(cb === 4 ? [i16(v, 6 + i * 4), i16(v, 8 + i * 4)] : [i32(v, 6 + i * cb), i32(v, 10 + i * cb)]);
    const W = props.get(0x0142)?.val ?? 21600; const H = props.get(0x0143)?.val ?? 21600;
    const L = props.get(0x0140)?.val ?? 0; const T = props.get(0x0141)?.val ?? 0;
    const cmds = [];
    let k = 0;
    const P = () => { const q = pts[k++] ?? [0, 0]; return [q[0] - L, q[1] - T]; };
    if (s && s.length >= 6) {
      const ns = u16(s, 0); const cbs = u16(s, 4) === 0xfff0 ? 2 : u16(s, 4) || 2;
      for (let i = 0; i < ns; i++) {
        const seg = u16(s, 6 + i * cbs);
        const type = seg >> 13; const cnt = seg & 0x1fff;
        if (type === 0) for (let j = 0; j < Math.max(1, cnt); j++) cmds.push(['L', ...P()]);
        else if (type === 1) for (let j = 0; j < Math.max(1, cnt); j++) cmds.push(['C', ...P(), ...P(), ...P()]);
        else if (type === 2) cmds.push(['M', ...P()]);
        else if (type === 3) cmds.push(['Z']);
        else if (type === 4) break;
        else if (type === 5 || type === 6) { /* 확장 · 클라이언트 명령: 건너뜀 */ }
      }
    } else { cmds.push(['M', ...P()]); while (k < pts.length) cmds.push(['L', ...P()]); }
    if (!cmds.length || cmds[0][0] !== 'M') cmds.unshift(['M', ...(pts[0] ?? [0, 0])]);
    return [{ w: W - L || 21600, h: H - T || 21600, cmds }];
  }

  function readShape(sp, xf, ctx) {
    const fsp = one(sp.kids, ART.FSP);
    if (!fsp) return null;
    const flags = u32(fsp.data, 4);
    if (flags & 0x8 || flags & 0x400) return null; // 지움 · 배경
    const spt = fsp.inst;
    const props = mergeProps(fopt(one(sp.kids, ART.FOPT)), fopt(one(sp.kids, ART.TertiaryFOPT)));
    if (boolProp(props, 0x03bf, 1, false)) return null; // 숨김
    const rot = props.get(0x0004) ? fixed(props.get(0x0004).val) : 0;
    const raw = rawAnchor(sp, xf, rot);
    if (!raw) return null;
    const rect = placeInGroups(xf, { ...raw, rot: ((rot % 360) + 360) % 360, ...(flags & 0x40 ? { flipH: true } : {}), ...(flags & 0x80 ? { flipV: true } : {}) });
    const o = { id: uid(), x: rect.x, y: rect.y, w: Math.max(0, rect.w), h: Math.max(0, rect.h), rot: rect.rot ?? 0 };
    if (rect.flipH) o.flipH = true;
    if (rect.flipV) o.flipV = true;
    const cd = one(sp.kids, ART.ClientData);
    const phAtom = deep(cd?.kids, RT.PlaceholderAtom)[0];
    const phId = phAtom ? phAtom.data[4] : 0;
    // 하이퍼링크 (클릭 동작)
    const ia = deep(cd?.kids, RT.InteractiveInfoAtom)[0];
    if (ia) {
      const action = ia.data[8]; const jump = ia.data[9];
      if (action === 4) { const l = links.get(u32(ia.data, 4)); const t = linkTarget(l); if (t) o.link = t; } else if (action === 3) o.link = { 1: '#next', 2: '#prev', 3: '#first', 4: '#last' }[jump] ?? undefined;
      if (!o.link) delete o.link;
    }
    const pib = props.get(0x0104)?.val;
    const isPic = (spt === 75 || pib) && !(flags & 0x10 && !pib);
    if (isPic && pib) {
      o.type = 'image';
      o.media = blipMedia(pib);
      if (!o.media) { warnings.add('표시할 수 없는 그림(PICT 등)은 빈 칸으로 표시합니다'); o.media = null; }
      const cr = (pid) => (props.get(pid) ? fixed(props.get(pid).val) : 0);
      const crop = { t: cr(0x0100), b: cr(0x0101), l: cr(0x0102), r: cr(0x0103) };
      if (crop.t || crop.b || crop.l || crop.r) o.crop = crop;
      if (boolProp(props, 0x013f, 1, false)) o.gray = true;
      const ln = boolProp(props, 0x01ff, 3, false) ? lineOf(props, ctx, false) : null;
      if (ln) o.line = ln;
      if (flags & 0x10) o.ole = { progId: '', name: 'OLE' };
      if (phId) o.ph = 'pic';
      return o.media || o.ph ? o : null;
    }
    o.type = 'shape';
    const isLine = LINE_SHAPES.has(spt);
    o.shape = SPT[spt] ?? 'rect';
    if (spt === 0 || props.has(0x0145)) { const g = customGeom(props); if (g) { o.path = g; o.shape = 'rect'; } }
    o.fill = isLine ? null : fillOf(props, ctx);
    o.line = lineOf(props, ctx, isLine);
    // 글 상자 · 개체 틀 기본: 채우기 · 선 없음 (문서 기본 속성을 쓰지 않음)
    const ownBool = (pr, pid, bit) => (pr?.get(pid)?.val ?? 0) & (1 << (bit + 16));
    const mp = phId ? ctx.phProps?.[{ 13: 1, 15: 3, 16: 4, 17: 1 }[phId] ?? (phId >= 14 && phId <= 26 ? 2 : phId)] : null;
    if ((spt === 202 || phId) && !ownBool(props, 0x01bf, 4)) o.fill = mp && ownBool(mp, 0x01bf, 4) ? fillOf(mergeProps(mp, props), ctx) : null;
    if ((spt === 202 || phId) && !ownBool(props, 0x01ff, 3)) o.line = mp && ownBool(mp, 0x01ff, 3) ? lineOf(mergeProps(mp, props), ctx, false) : null;
    if (boolProp(props, 0x023f, 1, false)) o.fx = { shadow: { dist: 3, dir: 45, blur: 2, color: '#000000', alpha: 0.4 } };
    // 개체 틀 종류
    const PH = { 1: 'title', 2: 'body', 3: 'ctrTitle', 4: 'subTitle', 13: 'title', 14: 'body', 15: 'ctrTitle', 16: 'subTitle', 17: 'title', 18: 'body', 19: 'body', 20: 'body', 21: 'body', 22: 'body', 23: 'body', 24: 'body', 25: 'body', 26: 'pic' };
    if (phId && PH[phId]) o.ph = PH[phId];
    if ([7, 8, 9, 10].includes(phId)) return null; // 마스터 날짜 · 번호 · 바닥글 개체 틀
    // 글
    const tb = one(sp.kids, ART.ClientTextbox);
    const pp9 = style9In(cd?.kids);
    const txt = tb ? textBlock(tb.kids, pp9 ? { ...ctx, pp9 } : ctx) : null;
    const ins = [props.get(0x0081)?.val ?? 91440, props.get(0x0082)?.val ?? 45720, props.get(0x0083)?.val ?? 91440, props.get(0x0084)?.val ?? 45720].map((v) => v / 9525);
    const anchor = { 0: 't', 1: 'ctr', 2: 'b', 3: 't', 4: 'ctr', 5: 'b' }[props.get(0x0087)?.val ?? 0] ?? 't';
    if (txt) o.text = { paras: txt.paras, anchor, insets: ins, wrap: (props.get(0x0085)?.val ?? 0) !== 2, autofit: 'none' };
    if (spt === 202) o.txBox = true;
    if (!o.text && !o.fill && !o.line && !o.ph) return null;
    if (!o.text && !isLine) o.text = { paras: [para()], anchor: 'ctr', insets: ins, wrap: true, autofit: 'none' };
    if (isLine && o.text && !o.text.paras.some((p) => p.runs.length)) o.text = null;
    return o;
  }
  const linkTarget = (l) => {
    if (!l) return null;
    const m = /^(\d+),(\d+),/.exec(l.loc || l.target || '');
    if (m) return `#slide${Number(m[2])}`;
    if (l.target && /^(https?|mailto|ftp|file):/i.test(l.target)) return l.target;
    return l.target || null;
  };

  /** 글 상자 레코드 (TextHeader · Chars/Bytes · StyleTextProp) → 단락 */
  function textBlock(kids, ctx, ruler = null) {
    let type = 4; let text = null; let style = null;
    for (const r of kids ?? []) {
      if (r.type === RT.TextRulerAtom) ruler = parseRuler(r.data);
    }
    for (const r of kids ?? []) {
      if (r.type === RT.TextHeaderAtom) type = u32(r.data, 0);
      else if (r.type === RT.TextCharsAtom) { let s = ''; for (let i = 0; i + 1 < r.data.length; i += 2) s += String.fromCharCode(u16(r.data, i)); text = s; } else if (r.type === RT.TextBytesAtom) { let s = ''; for (const c of r.data) s += String.fromCharCode(c); text = s; } else if (r.type === RT.StyleTextPropAtom) style = r.data;
      else if (r.type === RT.OutlineTextRefAtom) {
        // 개요 글은 SlideListWithText 에
        const t = ctx.texts?.[u32(r.data, 0)];
        if (t) {
          // 개요 글의 번호 매기기 (문서의 OutlineTextProps9: 슬라이드 id + 글 종류)
          const tt = u32(one(t, RT.TextHeaderAtom)?.data ?? new Uint8Array(4), 0);
          const pp9 = ctx.pp9 ?? outline9.get(`${ctx.slideId}:${tt}`) ?? null;
          return textBlock(t, pp9 ? { ...ctx, pp9 } : ctx, ruler);
        }
      }
    }
    if (text == null) return null;
    const baseType = { 5: 1, 6: 0, 7: 1, 8: 1 }[type] ?? type;
    ruler ??= ctx.rulers?.[type] ?? ctx.rulers?.[baseType] ?? null;
    return buildText(text, style, type, { ...ctx, ruler });
  }

  function buildText(text, style, type, ctx) {
    const { pruns, cruns } = parseStyleText(style, text.length);
    const lv = ctx.styles ?? mainMaster?.styles ?? envStyles;
    const baseType = { 5: 1, 6: 0, 7: 1, 8: 1 }[type] ?? type;
    // 수준 스타일: 기타(4) → 기본 종류(제목 0 · 본문 1) → 이 종류(가운데 제목 6 등, 다른 점만 담김) 순으로 겹침
    const levelStyle = (lvl) => {
      const layers = [envStyles?.[4]?.[lvl] ?? envStyles?.[4]?.[0], lv?.[4]?.[lvl], lv?.[baseType]?.[0], lv?.[baseType]?.[lvl], type !== baseType ? lv?.[type]?.[lvl] : null].filter(Boolean);
      const out = { pf: {}, cf: {} };
      for (const L of layers) { Object.assign(out.pf, L.pf ?? {}); Object.assign(out.cf, L.cf ?? {}); }
      return out;
    };
    const paras = [];
    const lines = text.split('\r');
    let pos = 0; let pi = 0; let pLeft = pruns[0]?.count ?? Infinity; let ci = 0; let cLeft = cruns[0]?.count ?? Infinity;
    for (const line of lines) {
      while (pLeft <= 0 && pi < pruns.length - 1) { pi++; pLeft = pruns[pi].count; }
      const pr = pruns[pi] ?? { pf: {}, indent: 0 };
      const lvl = Math.min(4, pr.indent ?? 0);
      const ms = levelStyle(lvl);
      const pf = { ...(ms.pf ?? {}), ...(pr.pf ?? {}) };
      const p9 = ctx.pp9?.[pi];
      if (p9?.scheme != null && p9.auto !== false) { pf.bulletScheme = p9.scheme; pf.bulletStart = p9.start || 1; }
      else if (p9?.auto === false) delete pf.bulletScheme;
      const runs = [];
      let s = 0;
      while (s < line.length) {
        while (cLeft <= 0 && ci < cruns.length - 1) { ci++; cLeft = cruns[ci].count; }
        const cr = cruns[ci] ?? { cf: {} };
        const n = Math.max(1, Math.min(line.length - s, cLeft));
        const cf = { ...(ms.cf ?? {}), ...(cr.cf ?? {}) };
        runs.push(makeRun(line.slice(s, s + n).replace(/\u000b/g, '\v'), cf, ctx, type));
        s += n; cLeft -= n;
      }
      // 단락 끝 CR 한 글자
      cLeft -= 1; pLeft -= line.length + 1; pos += line.length + 1;
      const end = makeRun('', { ...(ms.cf ?? {}), ...(cruns[ci]?.cf ?? {}) }, ctx, type);
      delete end.t;
      // 들여쓰기: 단락 속성 (PowerPoint 2000+) 또는 눈금자 (TextRulerAtom: 글 시작 · 글머리 위치, 마스터 단위 576/inch)
      const rl = ctx.ruler?.levels?.[lvl];
      const left = pf.leftMargin ?? rl?.left ?? 0; const first = pf.indent ?? rl?.indent ?? 0;
      const P = { runs: runs.filter((r) => r.t !== ''), align: ['l', 'ctr', 'r', 'just', 'dist', 'dist'][pf.align ?? 0] ?? 'l', lvl, bullet: null, marL: left / 6, indent: (first - left) / 6, lineSpacing: 1, spcBef: 0, spcAft: 0, end };
      if (ctx.ruler?.tabs?.length) P.tabs = ctx.ruler.tabs.map((t) => ({ pos: t.pos / 6, algn: t.algn }));
      if (pf.lineSpacing != null) { if (pf.lineSpacing < 0) P.lineSpacingPt = -pf.lineSpacing / 8; else P.lineSpacing = pf.lineSpacing / 100; }
      if (pf.spaceBefore != null) P.spcBef = pf.spaceBefore < 0 ? -pf.spaceBefore / 8 : (pf.spaceBefore / 100) * (end.size ?? 18);
      if (pf.spaceAfter != null) P.spcAft = pf.spaceAfter < 0 ? -pf.spaceAfter / 8 : (pf.spaceAfter / 100) * (end.size ?? 18);
      if (pf.hasBullet && line.length) {
        const ch = pf.bulletChar ? String.fromCharCode(pf.bulletChar) : '•';
        P.bullet = { type: 'char', char: ch };
        if (pf.bulletHasFont && pf.bulletFont != null && fontName(pf.bulletFont)) P.bullet.font = fontName(pf.bulletFont);
        if (pf.bulletHasColor && pf.bulletColor) { const c = textColor(pf.bulletColor, ctx.scheme); if (c) P.bullet.color = c; }
        if (pf.bulletHasSize && pf.bulletSize > 0) P.bullet.size = pf.bulletSize / 100;
        if (pf.bulletScheme != null) P.bullet = { type: 'num', scheme: ['alphaLcPeriod', 'alphaUcPeriod', 'arabicParenR', 'arabicPeriod', 'romanLcParenBoth', 'romanLcParenR', 'romanLcPeriod', 'romanUcPeriod', 'alphaLcParenBoth', 'alphaLcParenR', 'alphaUcParenBoth', 'alphaUcParenR', 'arabicParenBoth', 'arabicPlain', 'circleNumDbPlain', 'circleNumWdBlackPlain', 'circleNumWdWhitePlain', 'arabicDbPeriod', 'arabicDbPlain', 'ea1ChsPeriod', 'ea1ChsPlain', 'ea1ChtPeriod', 'ea1ChtPlain', 'ea1JpnChsDbPeriod', 'ea1JpnKorPlain', 'ea1JpnKorPeriod', 'arabic1Minus', 'arabic2Minus', 'hebrew2Minus', 'thaiAlphaPeriod', 'thaiAlphaParenR', 'thaiAlphaParenBoth', 'thaiNumPeriod', 'thaiNumParenR', 'thaiNumParenBoth', 'hindiAlphaPeriod', 'hindiNumPeriod', 'kanjiAlphaPeriod', 'hindiNumParenR', 'hindiAlpha1Period'][pf.bulletScheme] ?? 'arabicPeriod', start: pf.bulletStart ?? 1 };
      }
      paras.push(P);
      pi += 0;
    }
    return { paras: paras.length ? paras : [para()] };
  }
  function makeRun(t, cf, ctx, type) {
    const r = { t, size: cf.size ?? 18 };
    if (cf.bold) r.b = true;
    if (cf.italic) r.i = true;
    if (cf.underline) r.u = true;
    if (cf.shadow) r.shadow = true;
    if (cf.position) r.base = cf.position > 0 ? 1 : -1;
    const fnt = fontName(cf.font) ?? (cf.eaFont != null ? fontName(cf.eaFont) : null);
    if (fnt) r.font = fnt;
    const ea = cf.eaFont != null ? fontName(cf.eaFont) : null;
    if (ea && /[가-힣]/.test(t)) r.font = ea;
    const sym = cf.symFont != null ? fontName(cf.symFont) : null;
    if (sym && /[-]/.test(t)) r.sym = sym;
    const c = textColor(cf.color, ctx.scheme) ?? (type === 0 || type === 6 ? (ctx.scheme ?? baseScheme)[3] : (ctx.scheme ?? baseScheme)[1]);
    if (c) r.color = c;
    return r;
  }

  // ── 슬라이드
  const slides = [];
  const slideIdToIndex = new Map(slideList.map((e, i) => [e.slideId, i]));
  const notesBySlide = new Map();
  for (const ne of notesList) {
    const nr = recAt(ne.ref);
    if (!nr?.kids) continue;
    const na = one(nr.kids, RT.NotesAtom);
    if (!na) continue;
    const sid = u32(na.data, 0);
    const texts = [];
    for (const sp of deep(one(nr.kids, RT.PPDrawing)?.kids, ART.SpContainer)) {
      const ph = deep(one(sp.kids, ART.ClientData)?.kids, RT.PlaceholderAtom)[0];
      const tb = one(sp.kids, ART.ClientTextbox);
      if (!tb || (ph && ph.data[4] !== 12 && ph.data[4] !== 6)) continue;
      const t = textBlock(tb.kids, { scheme: null, texts: ne.texts });
      if (t) texts.push(t.paras.map((p) => p.runs.map((r) => r.t).join('')).join('\n'));
    }
    notesBySlide.set(sid, texts.join('\n').trim());
  }
  slideList.forEach((entry, si) => {
    try {
      const rec = recAt(entry.ref);
      if (!rec?.kids || rec.type !== RT.Slide) return;
      const atom = one(rec.kids, RT.SlideAtom);
      const masterRef = atom ? u32(atom.data, 12) : 0;
      const flags = atom ? u16(atom.data, 20) : 0x7;
      const m = masters.get(masterRef) ?? mainMaster;
      const mainOf = m?.base ?? m;
      const scheme = flags & 0x2 ? (m?.scheme ?? mainOf?.scheme ?? baseScheme) : (schemeOf(rec) ?? m?.scheme ?? baseScheme);
      const rulers = { ...(mainOf?.rulers ?? {}), ...(m?.rulers ?? {}) };
      const ctx = { scheme, slideId: entry.slideId, styles: m?.styles ?? mainMaster?.styles, texts: entry.texts, rulers, phProps: { ...(mainOf?.phProps ?? {}), ...(m?.phProps ?? {}) } };
      ctx.spMap = new Map(); ctx.legacyAnims = [];
      const { objects, bg } = readShapes(rec, ctx);
      // 마스터 배경 · 마스터 그림 (자리 표시자 제외)
      let slideBg = bg;
      const bgObjects = [];
      let masterAnim = null;
      if (m) {
        const mctx = { scheme, styles: m.styles, texts: [], rulers, spMap: new Map(), legacyAnims: [], phOf: new Map() };
        const mm = readShapes(m.rec, mctx);
        masterAnim = { spMap: mctx.spMap, legacy: mctx.legacyAnims, phOf: mctx.phOf };
        if (slideBg === undefined || flags & 0x4) slideBg = mm.bg ?? slideBg;
        if (slideBg === undefined && m.base) slideBg = readShapes(m.base.rec, mctx).bg;
        if (flags & 0x1) {
          // 제목 마스터는 자기 개체 + (자기 개체가 마스터 개체를 따르면) 주 마스터 개체
          for (const o of mm.objects) if (!o.ph) bgObjects.push({ ...o, decor: true, id: `bg-${o.id}` });
        }
      }
      const tr = one(rec.kids, RT.SlideShowSlideInfoAtom);
      let transition = null;
      let hidden = false;
      if (tr) {
        const d = tr.data;
        const time = i32(d, 0); const effDir = d[8]; const effType = d[9]; const tflags = u16(d, 10); const speed = d[12];
        hidden = !!(tflags & 0x4);
        const t = TRANS[effType];
        if (t && t !== 'cut' || (t === 'cut' && effDir)) transition = { type: t ?? 'fade', dur: [1, 0.75, 0.5][speed] ?? 0.75, dir: ['l', 't', 'r', 'b'][effDir] ?? 'l' };
        if (tflags & 0x10 || tflags & 0x400) { transition ??= { type: 'none', dur: 0.75 }; transition.advAfter = Math.max(0, time / 1000); }
      }
      let anims = [];
      const aopts = { soundFor, scheme };
      try { anims = timeNodeAnims(rec, ctx.spMap, objects, aopts) ?? legacyAnims(ctx.legacyAnims, objects, aopts); } catch (e) { warnings.add(`일부 애니메이션을 읽지 못했습니다: ${e.message}`); }
      // 슬라이드 마스터에 건 애니메이션: 마스터 그래픽을 따르는 슬라이드에서 먼저 재생 (PowerPoint 와 같음).
      // 움직이는 마스터 개체는 이 슬라이드의 개체로 옮겨 (맨 뒤 층) 편집 · 쇼 · pptx 저장이 그대로 되게 함
      if (masterAnim && (bgObjects.length || masterAnim.phOf.size)) {
        try {
          // 마스터 개체 → 이 슬라이드의 배경 개체, 마스터 개체 틀 (제목 · 본문) → 이 슬라이드의 같은 개체 틀 (PowerPoint 애니메이션 구성표)
          const PH_KIND = { title: 'title', ctrTitle: 'title', body: 'body', subTitle: 'body', obj: 'body' };
          const spMapBg = new Map([...masterAnim.spMap].map(([k, ids]) => {
            const ph = masterAnim.phOf.get(k);
            if (ph) return [k, objects.filter((o) => o.ph && PH_KIND[o.ph] && PH_KIND[o.ph] === PH_KIND[ph] && o.text?.paras?.some((p) => p.runs.some((r) => r.t.trim()))).map((o) => o.id)];
            return [k, ids.map((id) => `bg-${id}`)];
          }));
          const animTargets = [...objects, ...bgObjects];
          const mAnims = timeNodeAnims(m.rec, spMapBg, animTargets, aopts) ?? legacyAnims(masterAnim.legacy.map((x) => ({ ...x, obj: `bg-${x.obj}`, objs: x.objs?.map((id) => `bg-${id}`) })), bgObjects, aopts);
          if (mAnims.length) {
            const moving = new Set(mAnims.map((a) => a.obj).filter((id) => String(id).startsWith('bg-')));
            const grps = new Set(bgObjects.filter((o) => moving.has(o.id) && o.grp).map((o) => o.grp));
            const lift = bgObjects.filter((o) => moving.has(o.id) || (o.grp && grps.has(o.grp)));
            for (const o of lift) { delete o.decor; o.fromMaster = true; }
            const liftIds = new Set(lift.map((o) => o.id));
            bgObjects.splice(0, bgObjects.length, ...bgObjects.filter((o) => !liftIds.has(o.id)));
            objects.unshift(...lift);
            anims = [...mAnims, ...anims];
          }
        } catch (e) { warnings.add(`마스터 애니메이션을 읽지 못했습니다: ${e.message}`); }
      }
      const slide = { id: uid('s'), layout: 'blank', bg: slideBg ?? { type: 'solid', color: scheme[0] }, hidden, transition, notes: notesBySlide.get(entry.slideId) ?? '', objects, anims, bgObjects };
      slides.push(slide);
      void si;
    } catch (e) { warnings.add(`일부 슬라이드를 읽지 못했습니다: ${e.message}`); }
  });
  // 슬라이드 번호 링크 (#slideN, N = 1부터)
  void slideIdToIndex;
  if (!slides.length) slides.push({ id: uid('s'), layout: 'blank', bg: { type: 'solid', color: baseScheme[0] }, hidden: false, transition: null, notes: '', objects: [], anims: [] });
  const pres = { version: 1, size, theme, slides, media, props: { title: '', author: '', created: new Date().toISOString() }, footer: { slideNum: false, date: false, text: '', hideOnTitle: true } };
  return { pres, warnings: [...warnings] };
}

// ───────────── 글 서식 (StyleTextPropAtom · TextMasterStyleAtom) ─────────────
/** TextPFException 읽기 → [pf, 다음 위치] */
function readPF(d, p) {
  const m = u32(d, p); p += 4;
  const pf = {};
  if (m & 0xf) { const f = u16(d, p); p += 2; pf.hasBullet = !!(f & 1); pf.bulletHasFont = !!(f & 2); pf.bulletHasColor = !!(f & 4); pf.bulletHasSize = !!(f & 8); }
  if (m & 0x80) { pf.bulletChar = u16(d, p); p += 2; }
  if (m & 0x10) { pf.bulletFont = u16(d, p); p += 2; }
  if (m & 0x40) { pf.bulletSize = i16(d, p); p += 2; }
  if (m & 0x20) { pf.bulletColor = { hex: `#${[d[p], d[p + 1], d[p + 2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`, index: d[p + 3] }; p += 4; }
  if (m & 0x800) { pf.align = u16(d, p); p += 2; }
  if (m & 0x1000) { pf.lineSpacing = i16(d, p); p += 2; }
  if (m & 0x2000) { pf.spaceBefore = i16(d, p); p += 2; }
  if (m & 0x4000) { pf.spaceAfter = i16(d, p); p += 2; }
  if (m & 0x100) { pf.leftMargin = i16(d, p); p += 2; }
  if (m & 0x400) { pf.indent = i16(d, p); p += 2; }
  if (m & 0x8000) { p += 2; }
  if (m & 0x100000) { const n = u16(d, p); p += 2 + n * 4; }
  if (m & 0x10000) { p += 2; }
  if (m & 0xe0000) { p += 2; }
  if (m & 0x200000) { p += 2; }
  if (m & 0x1000000) { pf.bulletScheme = u16(d, p); pf.bulletStart = u16(d, p + 2); p += 4; } // bulletScheme (PP9)
  return [pf, p];
}
/** TextCFException 읽기 → [cf, 다음 위치] */
function readCF(d, p) {
  const m = u32(d, p); p += 4;
  const cf = {};
  if (m & 0xffff) {
    const s = u16(d, p); p += 2;
    if (m & 0x1) cf.bold = !!(s & 0x1);
    if (m & 0x2) cf.italic = !!(s & 0x2);
    if (m & 0x4) cf.underline = !!(s & 0x4);
    if (m & 0x10) cf.shadow = !!(s & 0x10);
  }
  if (m & 0x10000) { cf.font = u16(d, p); p += 2; }
  if (m & 0x200000) { cf.eaFont = u16(d, p); p += 2; }
  if (m & 0x400000) { p += 2; }
  if (m & 0x800000) { cf.symFont = u16(d, p); p += 2; }
  if (m & 0x20000) { cf.size = u16(d, p); p += 2; }
  if (m & 0x40000) { cf.color = { hex: `#${[d[p], d[p + 1], d[p + 2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`, index: d[p + 3] }; p += 4; }
  if (m & 0x80000) { cf.position = i16(d, p); p += 2; }
  if (m & 0x100000) { p += 4; }
  if (m & 0x1000000) { p += 2; }
  if (m & 0x2000000) { p += 2; }
  if (m & 0x4000000) { p += 4; }
  return [cf, p];
}
/** StyleTextPropAtom → 단락 · 글자 구간 */
/** StyleTextProp9Atom (PowerPoint 2000+ 확장) → 단락 묶음마다 { auto, scheme, start } (번호 매기기) */
export function parseStyle9(d) {
  const out = [];
  let p = 0;
  try {
    while (p + 12 <= d.length) {
      const e = {};
      const pm = u32(d, p); p += 4;
      if (pm & 0x800000) p += 2; // bulletBlipRef
      if (pm & 0x1000000) { e.auto = !!u16(d, p); p += 2; } // fBulletHasAutoNumber
      if (pm & 0x2000000) { e.scheme = u16(d, p); e.start = i16(d, p + 2); p += 4; } // bulletAutoNumberScheme
      const cm = u32(d, p); p += 4;
      if (cm & 0x100000) p += 4; // pp10ext
      if (cm & 0x1000000) p += 2; // newEATypeface
      if (cm & 0x2000000) p += 2; // csTypeface
      if (cm & 0x4000000) p += 4; // pp11ext
      const sm = u32(d, p); p += 4;
      if (sm & 0x1) p += 2; if (sm & 0x2) p += 2; if (sm & 0x4) p += 2; if (sm & 0x40) p += 2; if (sm & 0x20) p += 4;
      if (sm & 0x200) { const n = u32(d, p); p += 4 + n * 4; }
      out.push(e);
    }
  } catch { /* 읽은 데까지 */ }
  return out;
}
/** 레코드 안의 ProgBinaryTag 자료(0x138B)들에서 StyleTextProp9Atom 찾기 */
function style9In(kids) {
  for (const blob of deep(kids, 0x138b)) {
    const rs = records(blob.data);
    const a = deep(rs, 0x0fac)[0];
    if (a) return parseStyle9(a.data);
  }
  return null;
}

function parseStyleText(d, len) {
  const pruns = []; const cruns = [];
  if (!d) return { pruns: [{ count: len + 1, indent: 0, pf: {} }], cruns: [{ count: len + 1, cf: {} }] };
  let p = 0;
  try {
    let left = len + 1;
    while (left > 0 && p + 6 <= d.length) {
      const count = u32(d, p); const indent = u16(d, p + 4); p += 6;
      const [pf, np] = readPF(d, p); p = np;
      pruns.push({ count, indent, pf });
      left -= count;
    }
    left = len + 1;
    while (left > 0 && p + 4 <= d.length) {
      const count = u32(d, p); p += 4;
      const [cf, np] = readCF(d, p); p = np;
      cruns.push({ count, cf });
      left -= count;
    }
  } catch { /* 남은 것은 기본값 */ }
  if (!pruns.length) pruns.push({ count: len + 1, indent: 0, pf: {} });
  if (!cruns.length) cruns.push({ count: len + 1, cf: {} });
  return { pruns, cruns };
}
/** TextMasterStyleAtom 들 → { 글 종류: [수준별 { pf, cf }] } */
function parseMasterStyles(kids) {
  const out = {};
  for (const r of kids ?? []) {
    if (r.type !== RT.TextMasterStyleAtom) continue;
    const d = r.data;
    const levels = [];
    try {
      const n = u16(d, 0);
      let p = 2;
      for (let i = 0; i < n && p < d.length; i++) {
        if (r.inst >= 5) p += 2; // 수준 번호
        const [pf, p1] = readPF(d, p);
        const [cf, p2] = readCF(d, p1);
        p = p2;
        levels.push({ pf, cf });
      }
    } catch { /* 일부만 */ }
    // 수준 상속: 아래 수준은 위 수준을 이어받음
    for (let i = 1; i < levels.length; i++) levels[i] = { pf: { ...levels[i - 1].pf, ...levels[i].pf }, cf: { ...levels[i - 1].cf, ...levels[i].cf } };
    out[r.inst] = levels;
  }
  return out;
}

// ───────────── 애니메이션 ─────────────
// PowerPoint 2002+ 시간 노드 (슬라이드 ProgTags ___PPT10 → ExtTimeNodeContainer 0xF144) — pptx 의 <p:timing> 과 같은 나무.
/** PowerPoint 기본 소리 이름 → WIPOINT 소리 키 */
const BUILTIN_SOUND = { applause: 'applause', arrow: 'arrow', bomb: 'bomb', breeze: 'breeze', camera: 'camera', cashregister: 'cashreg', cashreg: 'cashreg', chime: 'chime', click: 'click', coin: 'coin', drumroll: 'drumroll', explosion: 'explode', explode: 'explode', hammer: 'hammer', laser: 'laser', push: 'push', suction: 'suction', typewriter: 'typewriter', type: 'typewriter', voltage: 'voltage', whoosh: 'whoosh', wind: 'wind' };
const TN = { Sub: 0xf145, Group: 0xf144, Node: 0xf127, Props: 0xf13d, Value: 0xf142, Event: 0xf125, Trigger: 0xf128, Target: 0xf13c, Ref: 0x2afb, Motion: 0xf12e };
const TN_CLASS = { 1: 'entr', 2: 'exit', 3: 'emph', 4: 'path' };
/** 속성 값 (첫 바이트 = 형식: 0 바이트 · 1 int32 · 2 float · 3 UTF-16 글) */
function tnValue(d) {
  if (!d?.length) return null;
  const t = d[0];
  if (t === 0) return d[1];
  if (t === 1) return i32(d, 1);
  if (t === 2) return new DataView(d.buffer, d.byteOffset + 1, 4).getFloat32(0, true);
  if (t === 3) { let s = ''; for (let p = 1; p + 1 < d.length; p += 2) { const c = u16(d, p); if (!c) break; s += String.fromCharCode(c); } return s; }
  return null;
}
function tnProps(g) {
  const m = new Map();
  for (const v of one(g.kids, TN.Props)?.kids ?? []) if (v.type === TN.Value) m.set(v.inst, tnValue(v.data));
  return m;
}
const tnNode = (g) => { const a = one(g.kids, TN.Node)?.data; return a?.length >= 32 ? { group: i32(a, 8), dur: i32(a, 24), nodeType: i32(a, 28) } : null; };
/** 시작 조건 (instance 1) → { trigger, delay(ms), ref } */
function tnBegin(g) {
  const ev = (g.kids ?? []).find((k) => k.type === TN.Event && k.inst === 1);
  const tr = one(ev?.kids, TN.Trigger)?.data;
  const ref = deep(ev?.kids, TN.Ref)[0]?.data;
  return { trigger: tr ? i32(tr, 4) : 0, delay: tr ? i32(tr, 12) : 0, ref: ref?.length >= 12 ? { mode: i32(ref, 0), type: i32(ref, 4), id: u32(ref, 8) } : null };
}
const tnKids = (g) => (g.kids ?? []).filter((k) => k.type === TN.Group);

/** 슬라이드 레코드 → anims (시간 노드가 없으면 null) */
export function timeNodeAnims(rec, spMap, objects, opts = {}) {
  let root = null;
  for (const blob of deep(rec.kids, 0x138b)) { root = records(blob.data).find((r) => r.type === TN.Group); if (root) break; }
  if (!root) return null;
  const anims = [];
  const seqs = [];
  (function find(g, d) { const p = tnProps(g); const nt = p.get(20); if (nt === 4 || nt === 5) { seqs.push([g, nt]); return; } if (d < 6) for (const k of tnKids(g)) find(k, d + 1); })(root, 0);
  const objOf = (id) => spMap.get(id);
  for (const [seq, nt] of seqs) {
    let trigger = null;
    if (nt === 5) {
      const b = tnBegin(seq);
      trigger = b.ref ? objOf(b.ref.id)?.[0] : null;
      if (!trigger) continue;
    }
    const effects = [];
    (function walk(g, d) { for (const k of tnKids(g)) { const p = tnProps(k); if (p.has(11) && [1, 2, 3].includes(p.get(20))) effects.push([k, p]); else if (d < 6) walk(k, d + 1); } })(seq, 0);
    for (const [eff, p] of effects) {
      const cls = TN_CLASS[p.get(11)];
      if (!cls) continue;
      const main = (eff.kids ?? []).filter((k) => k.type !== TN.Sub);
      const subs = (eff.kids ?? []).filter((k) => k.type === TN.Sub);
      const refs = deep(main, TN.Ref).map((r) => r.data).filter((d) => d.length >= 12 && i32(d, 4) === 1);
      const r0 = refs[0];
      const objs = r0 ? objOf(u32(r0, 8)) : null;
      if (!objs?.length) continue;
      const pid = p.get(9) ?? 0;
      const effect = cls === 'path' ? 'pathRight' : PRESET_ANIM[cls]?.[pid] ?? (cls === 'emph' ? 'pulse' : 'fade');
      const durs = [];
      (function dur(g, d) { const n = tnNode(g); if (n && n.dur > 1) durs.push(n.dur + Math.max(0, tnBegin(g).delay)); if (d < 8) for (const k of tnKids(g)) dur(k, d + 1); })(eff, 0);
      const node = p.get(20);
      const delay = Math.max(0, tnBegin(eff).delay);
      const a = { id: uid('a'), obj: objs[0], cls, effect, start: node === 2 ? 'with' : node === 3 ? 'after' : 'click', dur: durs.length ? Math.max(...durs) / 1000 : effect === 'appear' || effect === 'disappear' ? 0 : 0.5, delay: delay / 1000, dir: SUB_DIR[p.get(10)] ?? 'b' };
      if (cls === 'path') {
        const mo = deep(eff.kids, TN.Motion)[0];
        const path = deep(mo?.kids, TN.Value).map((v) => tnValue(v.data)).find((v) => typeof v === 'string');
        a.motion = (path ?? 'M 0 0 L 0.25 0').replace(/\s*[Ee]\s*$/, '').trim();
        a.effect = Object.keys(MOTION_PATHS).find((k) => MOTION_PATHS[k] === a.motion) ?? 'pathRight';
        if (!durs.length) a.dur = 2;
      }
      if (trigger) a.trigger = trigger;
      // 하위 노드 (AnimSubGroup 0xF145 = pptx subTnLst): 소리 (대상 종류 2) · 애니메이션 후 (속성 13)
      for (const sg of subs) {
        const sp = tnProps(sg);
        const sref = deep(sg.kids, TN.Ref).map((r) => r.data).find((d) => d.length >= 12 && i32(d, 4) === 2);
        if (sref && opts.soundFor) Object.assign(a, opts.soundFor(u32(sref, 8)) ?? {});
        if (!sp.get(13)) continue;
        const names = deep(sg.kids, 0xf13e).flatMap((n) => deep(n.kids, TN.Value).map((v) => tnValue(v.data)));
        if (names.includes('style.visibility')) a.after = { type: sp.get(5) === 2 ? 'hideNext' : 'hide' };
        else if (names.some((n) => /color/i.test(String(n)))) {
          const cd = deep(sg.kids, 0xf135)[0]?.data;
          const rgb = cd?.length >= 52 && i32(cd, 36) === 0 ? [i32(cd, 40), i32(cd, 44), i32(cd, 48)] : null;
          a.after = { type: 'dim', color: rgb ? `#${rgb.map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase()}` : '#808080' };
        }
      }
      // 단락 대상 (VisualShapeAtom 종류 2 = 글 범위): 같은 개체 · 같은 효과가 이어지면 하나로 묶고 byPara
      if (i32(r0, 0) === 2) {
        if (anims.some((b) => b.obj === a.obj && b.cls === cls && b.effect === a.effect && b.byPara && (b.trigger ?? null) === (a.trigger ?? null))) continue;
        a.byPara = true;
      }
      anims.push(a);
      // 그룹 대상: 나머지 개체는 함께
      for (const id of objs.slice(1)) anims.push({ ...a, id: uid('a'), obj: id, start: 'with', delay: a.delay });
    }
  }
  return anims.filter((a) => objects.some((o) => o.id === a.obj));
}

// PowerPoint 97/2000 빌드 정보 (AnimationInfoAtom): 효과 · 방향 · 순서 · 자동 시작
const LEGACY_EFFECT = { 0: 'appear', 1: 'fade', 2: 'blinds', 3: 'checker', 4: 'peek', 5: 'dissolveIn', 6: 'fade', 7: 'peek', 8: 'randomBars', 9: 'strips', 10: 'wipe', 11: 'box', 13: 'split', 14: 'appear', 17: 'diamond', 18: 'plus', 19: 'wedge', 26: 'wheel', 27: 'circleIn' };
const LEGACY_DIR = ['l', 't', 'r', 'b'];
export function legacyAnims(list, objects, opts = {}) {
  const items = list.map((x) => ({ ...x, order: u16(x.d, 16), flags: u16(x.d, 4), delay: i32(x.d, 12), build: x.d[20], eff: x.d[21], dir: x.d[22], after: x.d[23], sound: u32(x.d, 8), dim: x.d.subarray(0, 4) }))
    .filter((x) => x.order > 0 && x.order < 0x8000 && objects.some((o) => o.id === x.obj))
    .sort((a, b) => a.order - b.order);
  const anims = [];
  for (const x of items) {
    let effect = LEGACY_EFFECT[x.eff] ?? 'fade';
    let dir = LEGACY_DIR[x.dir] ?? 'b';
    if (x.eff === 12) {
      // 날아오기 계열: 0–7 날아오기 · 8–11 살며시 · 12–17 확대/축소 · 18–22 늘이기 · 23 회전 · 24 나선
      const v = x.dir;
      effect = v < 8 ? 'fly' : v < 12 ? 'crawl' : v < 18 ? 'zoom' : v < 23 ? 'stretch' : v === 23 ? 'swivel' : 'spiralIn';
      dir = v < 8 ? (LEGACY_DIR[v] ?? (v === 4 || v === 5 ? 't' : 'b')) : v < 12 ? LEGACY_DIR[v - 8] : v > 18 && v < 23 ? LEGACY_DIR[v - 19] : 'b';
    } else if (x.eff === 10 || x.eff === 4 || x.eff === 7) dir = LEGACY_DIR[x.dir] ?? 'l';
    const a = { id: uid('a'), obj: x.obj, cls: 'entr', effect, start: x.flags & 0x4 ? 'after' : 'click', dur: effect === 'appear' ? 0 : 0.5, delay: x.flags & 0x4 ? Math.max(0, x.delay) / 1000 : 0, dir };
    if (x.build >= 2) a.byPara = true;
    // 소리 (fSound + soundIdRef) · 애니메이션 후 (1 흐리게 dimColor · 2 다음 클릭 때 숨기기 · 3 바로 숨기기)
    if (x.flags & 0x10 && opts.soundFor) Object.assign(a, opts.soundFor(x.sound) ?? {});
    if (x.after === 1) {
      const [r, g, b, idx] = x.dim;
      const c = idx === 0xfe ? `#${[r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase()}` : opts.scheme?.[idx] ?? '#808080';
      a.after = { type: 'dim', color: c };
    } else if (x.after === 2) a.after = { type: 'hideNext' };
    else if (x.after === 3) a.after = { type: 'hide' };
    anims.push(a);
    for (const id of x.objs?.slice(1) ?? []) anims.push({ ...a, id: uid('a'), obj: id, start: 'with', delay: 0 });
  }
  return anims;
}
