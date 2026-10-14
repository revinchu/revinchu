// 문서 모델 (DOM 없음 · Node 테스트 · MCP 서버에서 그대로 씀)
//
// doc = { version, props, styles, defaults, lists, section (마지막 구역), body: [블록], notes, comments, media, settings }
// 블록: 문단 { type:'p', id, style, pPr, runs, secBreak? } · 표 { type:'tbl', id, style, tblPr, grid:[px], rows:[{ h, hRule, header, cantSplit, cells:[{ span, vMerge, w, shd, borders, vAlign, mar, blocks }] }] }
// 런: 글 { text, ...글자 서식 } · 개체 { obj: 'tab'|'br'|'img'|'field'|'note'|'mark'|'sym', … } (개체 하나 = 위치 1칸)
// 위치: { p: [블록 번호, (행, 열, 블록 번호)…], o: 문단 안 위치 }
// 단위: 길이 px (96dpi), 글자 크기 pt. docx twip = px × 15, HWPUNIT = px × 75
//
// 변경은 '복사 후 바꾸기': 기록된(실행 취소에 남은) 블록 · 객체는 절대 고치지 않음. change() 안에서 mut.* 으로 복사본을 얻어 고침.

let idSeq = 0;
export const uid = (p = 'b') => `${p}${Date.now().toString(36).slice(-4)}${(idSeq++).toString(36)}${Math.random().toString(36).slice(2, 5)}`;
export const clone = (v) => (v == null ? v : typeof structuredClone === 'function' ? structuredClone(v) : JSON.parse(JSON.stringify(v)));
export const PX_PER_PT = 4 / 3;
export const TWIP = 15; // px → twip
export const cm = (v) => (v * 96) / 2.54;
export const mm = (v) => (v * 96) / 25.4;
export const inch = (v) => v * 96;
export const pxToCm = (px) => (px * 2.54) / 96;

// ───────────── 쪽 크기 ─────────────
export const PAPERS = {
  A4: { w: mm(210), h: mm(297), label: 'A4', size: '21cm × 29.7cm' },
  A3: { w: mm(297), h: mm(420), label: 'A3', size: '29.7cm × 42cm' },
  A5: { w: mm(148), h: mm(210), label: 'A5', size: '14.8cm × 21cm' },
  B4: { w: mm(257), h: mm(364), label: 'B4 (JIS)', size: '25.7cm × 36.4cm' },
  B5: { w: mm(182), h: mm(257), label: 'B5 (JIS)', size: '18.2cm × 25.7cm' },
  Letter: { w: inch(8.5), h: inch(11), label: 'Letter', size: '21.59cm × 27.94cm' },
  Legal: { w: inch(8.5), h: inch(14), label: 'Legal', size: '21.59cm × 35.56cm' },
  Executive: { w: inch(7.25), h: inch(10.5), label: 'Executive', size: '18.41cm × 26.67cm' },
};
/** Word 기본 여백 (보통: 위 · 아래 2.54cm, 좌우 2.54cm) — 한국어판 기본은 위 2.54 · 아래 2.54 · 왼쪽 3.17 · 오른쪽 3.17 */
export const MARGINS = {
  normal: { label: '보통', t: cm(2.54), b: cm(2.54), l: cm(3.17), r: cm(3.17) },
  narrow: { label: '좁게', t: cm(1.27), b: cm(1.27), l: cm(1.27), r: cm(1.27) },
  moderate: { label: '보통 (좁음)', t: cm(2.54), b: cm(2.54), l: cm(1.91), r: cm(1.91) },
  wide: { label: '넓게', t: cm(2.54), b: cm(2.54), l: cm(5.08), r: cm(5.08) },
  mirrored: { label: '대칭', t: cm(2.54), b: cm(2.54), l: cm(3.18), r: cm(2.54), mirror: true },
  office2003: { label: 'Office 2003 기본', t: cm(2.54), b: cm(2.54), l: cm(3.18), r: cm(3.18) },
  hwp: { label: '한글 기본 (위 2 · 아래 1.5 · 좌우 3)', t: mm(20), b: mm(15), l: mm(30), r: mm(30), header: mm(15), footer: mm(15) },
};
export function newSection(o = {}) {
  const p = PAPERS.A4;
  const m = MARGINS.normal;
  return {
    page: { w: p.w, h: p.h, orient: 'portrait' },
    margin: { t: m.t, b: m.b, l: m.l, r: m.r, header: cm(1.25), footer: cm(1.25), gutter: 0 },
    cols: { n: 1, space: cm(1.25), sep: false },
    headers: {}, // { default: [블록], first, even }
    footers: {},
    titlePg: false,
    pgNum: null, // { fmt: 'decimal'|'roman'|…, start }
    vAlign: 'top',
    ...o,
  };
}

// ───────────── 글자 · 문단 서식 ─────────────
// 글자: font (라틴), fontEa (한글), sz(pt), b, i, u (true|'double'|'thick'|'dotted'|'dash'|'wave'…), uColor, s, dstrike, color, hl (형광펜 이름 또는 #색), shd (글자 음영 #색),
//       va ('super'|'sub'), caps, smallCaps, spc (자간 pt), scale (장평 %), pos (위치 pt), hidden, link {url|anchor}, cstyle (글자 스타일 id), lang,
//       outline, shadow, emboss, imprint, glow, emph (강조점 'dot'|'comma'|'circle'), ins/del {author, date} (변경 내용)
export const RUN_KEYS = ['font', 'fontEa', 'sz', 'b', 'i', 'u', 'uColor', 's', 'dstrike', 'color', 'hl', 'shd', 'va', 'caps', 'smallCaps', 'spc', 'scale', 'pos', 'hidden', 'link', 'cstyle', 'lang', 'outline', 'shadow', 'emboss', 'imprint', 'emph', 'ins', 'del', 'kern', 'rtl'];
// 문단: align ('left'|'center'|'right'|'justify'|'distribute'), ind { l, r, first (+) | hanging (-) } (px), spc { before, after (px), line (배수 또는 px), rule ('auto'|'exact'|'atLeast') },
//       keepNext, keepLines, pageBreakBefore, widow (기본 true), num { id, lvl }, tabs [{ pos, align, leader }], border { top, bottom, left, right, between: { style, w(pt), color, space } },
//       shd (#색), outline (개요 수준 0–8), contextual (같은 스타일 사이 간격 없음), noHyphen, dropCap { lines, kind:'drop'|'margin', font }
export const PARA_KEYS = ['align', 'ind', 'spc', 'keepNext', 'keepLines', 'pageBreakBefore', 'widow', 'num', 'tabs', 'border', 'shd', 'outline', 'contextual', 'noHyphen', 'dropCap', 'bidi', 'snapGrid'];

export const HIGHLIGHTS = { yellow: '#FFFF00', brightGreen: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF' };
export const HIGHLIGHT_LABEL = { yellow: '노랑', brightGreen: '밝은 녹색', cyan: '옥색', magenta: '분홍', blue: '파랑', red: '빨강', darkBlue: '진한 파랑', darkCyan: '청록', darkGreen: '녹색', darkMagenta: '보라', darkRed: '진한 빨강', darkYellow: '황록', darkGray: '회색 50%', lightGray: '회색 25%', black: '검정' };

// ───────────── 기본 스타일 (Word 2016+ 기본 서식 파일, 한국어판) ─────────────
const H = (lvl, sz, color, extra = {}) => ({ name: `제목 ${lvl}`, type: 'para', basedOn: 'Normal', next: 'Normal', link: `Heading${lvl}Char`, quick: true, order: 9 + lvl,
  pPr: { keepNext: true, keepLines: true, spc: { before: lvl === 1 ? 32 : 2.67, after: 0 }, outline: lvl - 1, ...extra.pPr },
  rPr: { font: '+mj', fontEa: '+mj', sz, color, ...extra.rPr } });
export function builtinStyles() {
  const st = {
    Normal: { name: '표준', type: 'para', quick: true, order: 0, pPr: {}, rPr: {} },
    NoSpacing: { name: '간격 없음', type: 'para', basedOn: 'Normal', quick: true, order: 1, pPr: { spc: { after: 0, line: 1, rule: 'auto' } }, rPr: {} },
    Heading1: H(1, 16, '#2F5496'),
    Heading2: H(2, 13, '#2F5496'),
    Heading3: H(3, 12, '#1F3763'),
    Heading4: H(4, 11, '#2F5496', { rPr: { i: true } }),
    Heading5: H(5, 11, '#2F5496'),
    Heading6: H(6, 11, '#1F3763'),
    Heading7: H(7, 11, '#1F3763', { rPr: { i: true } }),
    Heading8: H(8, 10.5, '#272727'),
    Heading9: H(9, 10.5, '#272727', { rPr: { i: true } }),
    Title: { name: '제목', type: 'para', basedOn: 'Normal', next: 'Normal', quick: true, order: 2, pPr: { spc: { after: 0, line: 1, rule: 'auto' }, contextual: true }, rPr: { font: '+mj', fontEa: '+mj', sz: 28, spc: -0.5, kern: 14 } },
    Subtitle: { name: '부제', type: 'para', basedOn: 'Normal', next: 'Normal', quick: true, order: 3, pPr: { numSub: true }, rPr: { color: '#5A5A5A', spc: 0.75, sz: 11 } },
    Quote: { name: '인용', type: 'para', basedOn: 'Normal', next: 'Normal', quick: true, order: 20, pPr: { spc: { before: 13.33, after: 10.67 }, ind: { l: 57.6, r: 57.6 }, align: 'center' }, rPr: { i: true, color: '#404040' } },
    IntenseQuote: { name: '강한 인용', type: 'para', basedOn: 'Normal', next: 'Normal', quick: true, order: 21, pPr: { border: { top: { style: 'single', w: 0.5, color: '#4472C4', space: 10 }, bottom: { style: 'single', w: 0.5, color: '#4472C4', space: 10 } }, spc: { before: 48, after: 48 }, ind: { l: 115.2, r: 115.2 }, align: 'center' }, rPr: { i: true, color: '#4472C4' } },
    ListParagraph: { name: '목록 단락', type: 'para', basedOn: 'Normal', quick: true, order: 22, pPr: { ind: { l: 53.33 }, contextual: true }, rPr: {} },
    Caption: { name: '캡션', type: 'para', basedOn: 'Normal', next: 'Normal', order: 23, pPr: { spc: { after: 13.33, line: 1, rule: 'auto' } }, rPr: { b: true, sz: 10, color: '#44546A' } },
    Header: { name: '머리글', type: 'para', basedOn: 'Normal', order: 40, pPr: { spc: { after: 0, line: 1, rule: 'auto' }, tabs: [{ pos: 300.8, align: 'center' }, { pos: 601.6, align: 'right' }] }, rPr: {} },
    Footer: { name: '바닥글', type: 'para', basedOn: 'Normal', order: 41, pPr: { spc: { after: 0, line: 1, rule: 'auto' }, tabs: [{ pos: 300.8, align: 'center' }, { pos: 601.6, align: 'right' }] }, rPr: {} },
    FootnoteText: { name: '각주 텍스트', type: 'para', basedOn: 'Normal', order: 42, pPr: { spc: { after: 0, line: 1, rule: 'auto' } }, rPr: { sz: 10 } },
    EndnoteText: { name: '미주 텍스트', type: 'para', basedOn: 'Normal', order: 43, pPr: { spc: { after: 0, line: 1, rule: 'auto' } }, rPr: { sz: 10 } },
    TOCHeading: { name: '목차 제목', type: 'para', basedOn: 'Heading1', next: 'Normal', order: 44, pPr: { outline: 9 }, rPr: {} },
    Bibliography: { name: '참고 문헌', type: 'para', basedOn: 'Normal', next: 'Normal', order: 45, pPr: {}, rPr: {} },
    // 글자 스타일
    DefaultParagraphFont: { name: '기본 단락 글꼴', type: 'char', order: 50, rPr: {} },
    Strong: { name: '강한 강조', type: 'char', quick: true, order: 24, rPr: { b: true } },
    Emphasis: { name: '강조', type: 'char', quick: true, order: 25, rPr: { i: true } },
    SubtleEmphasis: { name: '약한 강조', type: 'char', quick: true, order: 26, rPr: { i: true, color: '#404040' } },
    IntenseEmphasis: { name: '강한 강조 2', type: 'char', quick: true, order: 27, rPr: { i: true, color: '#4472C4' } },
    SubtleReference: { name: '약한 참조', type: 'char', quick: true, order: 28, rPr: { smallCaps: true, color: '#5A5A5A' } },
    IntenseReference: { name: '강한 참조', type: 'char', quick: true, order: 29, rPr: { b: true, smallCaps: true, color: '#4472C4', spc: 0.25 } },
    BookTitle: { name: '책 제목', type: 'char', quick: true, order: 30, rPr: { b: true, i: true, spc: 0.25 } },
    Hyperlink: { name: '하이퍼링크', type: 'char', order: 51, rPr: { color: '#0563C1', u: true } },
    FollowedHyperlink: { name: '열어 본 하이퍼링크', type: 'char', order: 52, rPr: { color: '#954F72', u: true } },
    FootnoteReference: { name: '각주 참조', type: 'char', order: 53, rPr: { va: 'super' } },
    EndnoteReference: { name: '미주 참조', type: 'char', order: 54, rPr: { va: 'super' } },
    PlaceholderText: { name: '개체 틀 텍스트', type: 'char', order: 55, rPr: { color: '#808080' } },
    // 표 스타일
    TableNormal: { name: '표준 표', type: 'table', order: 60, tblPr: { cellMar: { l: 7.2, r: 7.2, t: 0, b: 0 } } },
    TableGrid: { name: '표 눈금', type: 'table', basedOn: 'TableNormal', order: 61, tblPr: { borders: allBorders({ style: 'single', w: 0.5, color: '#000000' }) }, pPr: { spc: { after: 0, line: 1, rule: 'auto' } } },
  };
  for (let i = 1; i <= 9; i++) st[`TOC${i}`] = { name: `목차 ${i}`, type: 'para', basedOn: 'Normal', next: 'Normal', order: 70 + i, pPr: { spc: { after: 6.67 }, ind: { l: (i - 1) * 14.67 } }, rPr: {} };
  for (let i = 1; i <= 9; i++) st[`Heading${i}Char`] = { name: `제목 ${i} 문자`, type: 'char', link: `Heading${i}`, hidden: true, order: 90 + i, rPr: { ...st[`Heading${i}`].rPr } };
  for (const [k, s] of Object.entries(st)) s.id = k;
  return st;
}
export function allBorders(b) { return { top: { ...b }, bottom: { ...b }, left: { ...b }, right: { ...b }, insideH: { ...b }, insideV: { ...b } }; }

// 문서 기본값 (Word 2016+ 한국어판: 맑은 고딕 10pt, 줄 간격 1.08 → 한국어판은 배수 1.08 · 단락 뒤 8pt)
export const DEFAULT_RPR = { font: '+mn', fontEa: '+mn', sz: 10, color: null, lang: 'ko-KR' };
export const DEFAULT_PPR = { align: 'justify', spc: { before: 0, after: 10.67, line: 1.08, rule: 'auto' }, widow: true };
export const THEME_FONTS = { major: { latin: '맑은 고딕', ea: '맑은 고딕' }, minor: { latin: '맑은 고딕', ea: '맑은 고딕' } };

// ───────────── 목록 (번호 매기기) ─────────────
export const NUM_FORMATS = {
  decimal: '1, 2, 3', upperRoman: 'I, II, III', lowerRoman: 'i, ii, iii', upperLetter: 'A, B, C', lowerLetter: 'a, b, c',
  ganada: '가, 나, 다', chosung: 'ㄱ, ㄴ, ㄷ', decimalCircle: '①, ②, ③', koreanDigital: '일, 이, 삼', koreanCounting: '하나, 둘, 셋', ideographDigital: '一, 二, 三',
  decimalZero: '01, 02, 03', bullet: '•', none: '(없음)',
};
const BULLET_CHARS = ['●', '○', '■', '◆', '➢', '✓', '–', '▪', '◇'];
export function bulletList(char = '●') {
  return { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: 'bullet', text: i === 0 ? char : BULLET_CHARS[(i % 3) + (i % 3 === 0 ? 0 : 0)] ?? '●', start: 1, ind: { l: 26.67 + i * 26.67, hanging: 26.67 }, font: null })) };
}
export function numberList(fmts = ['decimal', 'lowerLetter', 'lowerRoman'], texts = ['%1.', '%2)', '%3.']) {
  return { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: fmts[i % fmts.length], text: (texts[i % texts.length] ?? '%1.').replace(/%\d/, `%${i + 1}`), start: 1, ind: { l: 26.67 + i * 26.67, hanging: 26.67 } })) };
}
export function multiList(kind = 'outline') {
  if (kind === 'legal') return { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: 'decimal', text: Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.', start: 1, ind: { l: 24 + i * 24, hanging: 24 + Math.min(i, 4) * 8 } })) };
  if (kind === 'korean') {
    // 한국 공문서 항목 기호: 1. → 가. → 1) → 가) → (1) → (가) → ① → ㉮
    const lv = [['decimal', '%1.'], ['ganada', '%2.'], ['decimal', '%3)'], ['ganada', '%4)'], ['decimal', '(%5)'], ['ganada', '(%6)'], ['decimalCircle', '%7'], ['ganada', '%8'], ['bullet', '-']];
    return { levels: lv.map(([fmt, text], i) => ({ fmt, text, start: 1, ind: { l: 20 + i * 13.33, hanging: 20 } })) };
  }
  return { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: ['decimal', 'lowerLetter', 'lowerRoman'][i % 3], text: `%${i + 1}${i % 3 === 0 ? '.' : ')'}`, start: 1, ind: { l: 24 + i * 24, hanging: 24 } })) };
}

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const roman = (n) => { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const letters = (n, base) => { const c = String.fromCharCode(base + ((n - 1) % 26)); return c.repeat(Math.floor((n - 1) / 26) + 1); };
const GANADA = '가나다라마바사아자차카타파하';
const CHOSUNG = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ';
const KDIG = ['영', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
const KCOUNT = ['', '하나', '둘', '셋', '넷', '다섯', '여섯', '일곱', '여덟', '아홉', '열'];
const CJK = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
const sinoNum = (n, d) => { if (n < 10) return d[n]; if (n < 100) return `${Math.floor(n / 10) > 1 ? d[Math.floor(n / 10)] : ''}${d === KDIG ? '십' : '十'}${n % 10 ? d[n % 10] : ''}`; return String(n); };
/** 번호 → 글 */
export function formatNumber(n, fmt) {
  switch (fmt) {
    case 'upperRoman': return roman(n).toUpperCase();
    case 'lowerRoman': return roman(n);
    case 'upperLetter': return letters(n, 65);
    case 'lowerLetter': return letters(n, 97);
    case 'ganada': return GANADA[(n - 1) % 14] ?? String(n);
    case 'chosung': return CHOSUNG[(n - 1) % 14] ?? String(n);
    case 'decimalCircle': return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : n <= 35 ? String.fromCharCode(0x3251 + n - 21) : n <= 50 ? String.fromCharCode(0x32b1 + n - 36) : String(n);
    case 'decimalZero': return n < 10 ? `0${n}` : String(n);
    case 'koreanDigital': return sinoNum(n, KDIG);
    case 'koreanCounting': return KCOUNT[n] ?? String(n);
    case 'ideographDigital': return sinoNum(n, CJK);
    case 'none': return '';
    case 'bullet': return '';
    default: return String(n);
  }
}

// ───────────── 새 문서 ─────────────
export function newDoc(o = {}) {
  return {
    version: 1,
    props: { title: '', subject: '', author: '', keywords: '', comments: '', created: new Date().toISOString(), modified: new Date().toISOString() },
    styles: builtinStyles(),
    defaults: { rPr: { ...DEFAULT_RPR }, pPr: clone(DEFAULT_PPR) },
    themeFonts: clone(THEME_FONTS),
    lists: {},
    section: newSection(),
    body: [para('')],
    notes: { foot: {}, end: {} },
    comments: [],
    media: {},
    settings: { track: false, defaultTab: 48 },
    ...o,
  };
}
export function para(text = '', o = {}) {
  const runs = typeof text === 'string' ? (text ? [{ text }] : []) : text;
  return { type: 'p', id: uid('p'), style: o.style ?? 'Normal', pPr: o.pPr ?? {}, runs, ...(o.mark ? { mark: o.mark } : {}) };
}
export function table(rows, cols, o = {}) {
  const w = o.width ?? 600;
  const grid = Array.from({ length: cols }, () => w / cols);
  return {
    type: 'tbl', id: uid('t'), style: o.style ?? 'TableGrid', tblPr: { w: { type: 'auto' }, look: { firstRow: true, firstCol: true, bandRows: true }, ...o.tblPr },
    grid,
    rows: Array.from({ length: rows }, () => ({ cells: Array.from({ length: cols }, () => ({ blocks: [para('', { style: 'Normal' })] })) })),
  };
}

// ───────────── 런 · 길이 ─────────────
export const runLen = (r) => (r.obj ? (r.obj === 'mark' ? 0 : 1) : r.text.length);
export const paraLen = (p) => { let n = 0; for (const r of p.runs) n += runLen(r); return n; };
/** 문단의 순수 글 (개체: 탭 \t, 줄 바꿈 \n, 그 밖 U+FFFC) — 위치와 1:1 */
export function paraText(p) {
  let s = '';
  for (const r of p.runs) s += r.obj ? (r.obj === 'mark' ? '' : r.obj === 'tab' ? '\t' : r.obj === 'br' ? '\n' : r.obj === 'sym' ? r.char ?? '￼' : '￼') : r.text;
  return s;
}
/** 읽을 글 (개체 제외, 필드는 결과 글) */
export function paraPlain(p) {
  let s = '';
  for (const r of p.runs) {
    if (!r.obj) { if (!r.del) s += r.text; } else if (r.obj === 'tab') s += '\t'; else if (r.obj === 'br') s += '\n'; else if (r.obj === 'field') s += r.result ?? ''; else if (r.obj === 'sym') s += r.char ?? '';
  }
  return s;
}
export const runProps = (r) => { const o = {}; for (const k of RUN_KEYS) if (r[k] !== undefined) o[k] = r[k]; return o; };
const sameProps = (a, b) => {
  for (const k of RUN_KEYS) {
    const x = a[k]; const y = b[k];
    if (x === y) continue;
    if (x && y && typeof x === 'object' && JSON.stringify(x) === JSON.stringify(y)) continue;
    return false;
  }
  return true;
};
/** 서식이 같은 이웃 글 런 합치기 · 빈 글 런 지우기 (새 배열) */
export function normalizeRuns(runs) {
  const out = [];
  for (const r of runs) {
    if (!r.obj && !r.text) continue;
    const last = out[out.length - 1];
    if (last && !last.obj && !r.obj && sameProps(last, r)) out[out.length - 1] = { ...last, text: last.text + r.text };
    else out.push(r);
  }
  return out;
}
/** 런 배열을 위치 o 에서 나눔 → [앞, 뒤] (원래 런은 고치지 않음) */
export function splitRuns(runs, o) {
  const a = [];
  const b = [];
  let pos = 0;
  for (const r of runs) {
    const n = runLen(r);
    if (pos + n <= o) a.push(r);
    else if (pos >= o) b.push(r);
    else { a.push({ ...r, text: r.text.slice(0, o - pos) }); b.push({ ...r, text: r.text.slice(o - pos) }); }
    pos += n;
  }
  return [a, b];
}
/** 위치 o 의 글자 서식 (앞 글자 기준, 문단 처음이면 첫 런 · 빈 문단은 문단 표시 서식) */
export function propsAt(p, o) {
  let pos = 0;
  let prev = null;
  for (const r of p.runs) {
    const n = runLen(r);
    if (!r.obj || r.obj === 'tab' || r.obj === 'field' || r.obj === 'sym') {
      if (o > pos && o <= pos + n) return runProps(r);
      if (!prev && o === pos) prev = r;
    }
    pos += n;
  }
  const r0 = prev ?? p.runs.find((r) => !r.obj);
  const out = r0 ? runProps(r0) : { ...(p.mark ?? {}) };
  delete out.link; delete out.ins; delete out.del;
  return out;
}

// ───────────── 위치 · 블록 찾기 ─────────────
export const posEq = (a, b) => a && b && a.o === b.o && a.p.length === b.p.length && a.p.every((v, i) => v === b.p[i]);
/** 위치 비교 (문서 순서) */
export function cmpPos(a, b) {
  const n = Math.min(a.p.length, b.p.length);
  for (let i = 0; i < n; i++) if (a.p[i] !== b.p[i]) return typeof a.p[i] === 'string' || typeof b.p[i] === 'string' ? String(a.p[i]).localeCompare(String(b.p[i])) : a.p[i] - b.p[i];
  if (a.p.length !== b.p.length) return a.p.length - b.p.length; // 같은 표 안: 더 깊은 쪽이 뒤가 아님 — 호출자는 같은 깊이만 비교
  return a.o - b.o;
}
export const orderPos = (a, b) => (cmpPos(a, b) <= 0 ? [a, b] : [b, a]);
/** 경로 → 블록 목록 · 번호: [i] → (body, i), [i, r, c, j] → (cell.blocks, j) */
// 스토리 (본문 밖 블록 목록): 경로 첫 값이 글 'h:<구역>:<종류>' (머리글) · 'f:<구역>:<종류>' (바닥글) · 'n:<foot|end>:<id>' (각주 · 미주)
export const isStoryPath = (path) => typeof path[0] === 'string';
/** 구역 번호 → 구역 객체 (마지막 구역 = doc.section) */
export function sectionObj(doc, si) {
  let k = 0;
  for (const b of doc.body) if (b.type === 'p' && b.secBreak) { if (k === si) return b.secBreak.sec; k++; }
  return doc.section;
}
export function storyBlocks(doc, key) {
  const [kind, a, b] = key.split(':');
  if (kind === 'h' || kind === 'f') { const sec = sectionObj(doc, Number(a)); return sec?.[kind === 'h' ? 'headers' : 'footers']?.[b] ?? null; }
  if (kind === 'n') return doc.notes?.[a]?.[b] ?? null;
  return null;
}
/** 경로의 뿌리 → { blocks, base(앞 경로), rest(남은 경로) } */
export function rootOf(doc, path) {
  if (isStoryPath(path)) return { blocks: storyBlocks(doc, path[0]) ?? [], base: [path[0]], rest: path.slice(1) };
  return { blocks: doc.body, base: [], rest: path };
}
export function containerOf(doc, path, root = null) {
  if (!root) { const r = rootOf(doc, path); root = r.blocks; path = r.rest; }
  let list = root;
  for (let k = 0; k + 1 < path.length; k += 3) {
    const t = list[path[k]];
    if (!t || t.type !== 'tbl') return null;
    const cell = t.rows[path[k + 1]]?.cells[path[k + 2]];
    if (!cell) return null;
    list = cell.blocks;
  }
  return list;
}
export function blockAt(doc, path, root = null) {
  if (root) return containerOf(doc, path, root)?.[path[path.length - 1]] ?? null;
  const list = containerOf(doc, path);
  return list ? list[path[path.length - 1]] ?? null : null;
}
/** 모든 문단 (문서 순서, 표 안 포함) — 콜백 (p, path) */
export function walkParas(blocks, fn, base = []) {
  blocks.forEach((b, i) => {
    const path = [...base, i];
    if (b.type === 'p') fn(b, path);
    else if (b.type === 'tbl') b.rows.forEach((row, r) => row.cells.forEach((cell, c) => walkParas(cell.blocks, fn, [...path, r, c])));
  });
}
/** 문서 처음 · 끝 위치 */
export function firstPos(blocks, base = []) {
  const b = blocks[0];
  if (!b) return { p: [...base, 0], o: 0 };
  if (b.type === 'tbl') return firstPos(b.rows[0].cells[0].blocks, [...base, 0, 0, 0]);
  return { p: [...base, 0], o: 0 };
}
export function lastPos(blocks, base = []) {
  const i = blocks.length - 1;
  const b = blocks[i];
  if (b.type === 'tbl') { const r = b.rows.length - 1; const c = b.rows[r].cells.length - 1; return lastPos(b.rows[r].cells[c].blocks, [...base, i, r, c]); }
  return { p: [...base, i], o: paraLen(b) };
}
/** 문단 경로 목록 (문서 순서) — 커서 이동용 */
export function paraPaths(doc, story = null) {
  const out = [];
  if (story) walkParas(storyBlocks(doc, story) ?? [], (p, path) => out.push(path), [story]);
  else walkParas(doc.body, (p, path) => out.push(path));
  return out;
}
const samePath = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
/** 다음 · 앞 문단 경로 (표 칸을 지나감) */
export function nextParaPath(doc, path, dir = 1) {
  const list = paraPaths(doc, isStoryPath(path) ? path[0] : null);
  const i = list.findIndex((p) => samePath(p, path));
  return i < 0 ? null : list[i + dir] ?? null;
}

// ───────────── 스타일 해석 ─────────────
/** 스타일 체인 (basedOn) 을 따라 rPr · pPr 합침 (캐시: styles 객체가 바뀌면 새로 계산) */
const styleCache = new WeakMap();
function mergeDeep(a, b) {
  if (!b) return a;
  const out = { ...a };
  for (const [k, v] of Object.entries(b)) {
    if (v && typeof v === 'object' && !Array.isArray(v) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k]) && k !== 'num' && k !== 'link') out[k] = mergeDeep(out[k], v);
    else out[k] = v;
  }
  return out;
}
export function styleChain(styles, id) {
  let cache = styleCache.get(styles);
  if (!cache) { cache = new Map(); styleCache.set(styles, cache); }
  if (cache.has(id)) return cache.get(id);
  const seen = new Set();
  const chain = [];
  let s = styles[id];
  while (s && !seen.has(s.id ?? id)) { seen.add(s.id ?? id); chain.unshift(s); s = s.basedOn ? styles[s.basedOn] : null; }
  let pPr = {}; let rPr = {}; let tblPr = {};
  for (const x of chain) { pPr = mergeDeep(pPr, x.pPr); rPr = mergeDeep(rPr, x.rPr); tblPr = mergeDeep(tblPr, x.tblPr); }
  const out = { pPr, rPr, tblPr };
  cache.set(id, out);
  return out;
}
/** 문단의 실제 문단 서식 (문서 기본 → 표 스타일 문단 서식 → 문단 스타일 → 직접 서식) */
export function effPara(doc, p, tblStyle = null) {
  let pPr = doc.defaults.pPr;
  if (tblStyle) pPr = mergeDeep(pPr, styleChain(doc.styles, tblStyle).pPr);
  pPr = mergeDeep(pPr, styleChain(doc.styles, p.style ?? 'Normal').pPr);
  return mergeDeep(pPr, p.pPr);
}
/** 런의 실제 글자 서식 (기본 → 문단 스타일 → 글자 스타일 → 직접) */
export function effRun(doc, p, r, tblStyle = null) {
  let rPr = doc.defaults.rPr;
  if (tblStyle) rPr = { ...rPr, ...styleChain(doc.styles, tblStyle).rPr };
  rPr = { ...rPr, ...styleChain(doc.styles, p.style ?? 'Normal').rPr };
  if (r.cstyle) rPr = { ...rPr, ...styleChain(doc.styles, r.cstyle).rPr };
  if (r.link && !r.cstyle) rPr = { ...rPr, ...styleChain(doc.styles, 'Hyperlink').rPr };
  for (const k of RUN_KEYS) if (r[k] !== undefined) rPr[k] = r[k];
  return rPr;
}
/** 테마 글꼴 (+mj / +mn) → 실제 이름 */
export function fontName(doc, f, ea = false) {
  if (f === '+mj') return ea ? doc.themeFonts?.major?.ea ?? '맑은 고딕' : doc.themeFonts?.major?.latin ?? '맑은 고딕';
  if (f === '+mn') return ea ? doc.themeFonts?.minor?.ea ?? '맑은 고딕' : doc.themeFonts?.minor?.latin ?? '맑은 고딕';
  return f;
}
export const headingLevel = (doc, p) => { const o = effPara(doc, p).outline; return o != null && o < 9 ? o + 1 : 0; };

// ───────────── 번호 매기기 계산 (문서 순서) ─────────────
/** 문단 id → { text, fmt, lvl, numId, ind, bullet, font } */
export function computeNumbering(doc) {
  const out = new Map();
  const counters = new Map(); // numId → [9]
  const visit = (p) => {
    const ep = effPara(doc, p);
    const num = ep.num;
    if (!num || num.id == null || num.id === 0) return;
    const def = doc.lists[num.id];
    if (!def) return;
    const lvl = Math.max(0, Math.min(8, num.lvl ?? 0));
    const absKey = def.abs ?? num.id;
    let c = counters.get(absKey);
    if (!c) { c = Array(9).fill(0); counters.set(absKey, c); }
    const L = def.levels[lvl] ?? def.levels[0];
    const start = def.override?.[lvl]?.start ?? L.start ?? 1;
    if (num.restart) c[lvl] = start; else c[lvl] = c[lvl] ? c[lvl] + 1 : start;
    for (let k = lvl + 1; k < 9; k++) c[k] = 0;
    let text = L.text ?? '';
    if (L.fmt === 'bullet') text = L.text || '●';
    else text = text.replace(/%(\d)/g, (_, d) => { const k = Number(d) - 1; const lk = def.levels[k] ?? L; const v = c[k] || (lk.start ?? 1); return formatNumber(v, k === lvl ? L.fmt : (L.legal ? 'decimal' : lk.fmt)); });
    out.set(p.id, { text, fmt: L.fmt, lvl, numId: num.id, ind: L.ind, bullet: L.fmt === 'bullet', font: L.font, rPr: L.rPr, suff: L.suff ?? 'tab' });
  };
  walkParas(doc.body, (p) => visit(p));
  return out;
}

// ───────────── 글 통계 ─────────────
export function docText(doc, { sep = '\n' } = {}) {
  const out = [];
  walkParas(doc.body, (p) => out.push(paraPlain(p)));
  return out.join(sep);
}
/** Word 방식 단어 수: 공백으로 나눈 덩어리 (한글 어절 포함) */
export function wordCount(text) {
  const m = text.match(/[^\s　]+/g);
  return m ? m.length : 0;
}
export function docStats(doc) {
  const text = docText(doc);
  let paras = 0; let lines = 0;
  walkParas(doc.body, (p) => { if (paraLen(p)) paras++; });
  lines = paras;
  return { words: wordCount(text), chars: text.replace(/\n/g, '').length, charsNoSpace: text.replace(/\s/g, '').length, paras, lines };
}

// ───────────── 실행 취소 기록 ─────────────
// 문서 객체는 바꾸지 않으므로 이전 문서 객체를 통째로 보관 (얕은 복사, 바뀐 블록만 새 객체)
export class History {
  constructor(limit = 300) { this.undoStack = []; this.redoStack = []; this.limit = limit; this.lastKey = null; this.lastAt = 0; }
  record(doc, sel, key = null) {
    const now = Date.now();
    if (key && key === this.lastKey && now - this.lastAt < 1500) { this.lastAt = now; this.redoStack = []; return false; }
    this.undoStack.push({ doc, sel });
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
    this.lastKey = key; this.lastAt = now;
    return true;
  }
  undo(doc, sel) { const s = this.undoStack.pop(); if (!s) return null; this.redoStack.push({ doc, sel }); this.lastKey = null; return s; }
  redo(doc, sel) { const s = this.redoStack.pop(); if (!s) return null; this.undoStack.push({ doc, sel }); this.lastKey = null; return s; }
  get canUndo() { return this.undoStack.length > 0; }
  get canRedo() { return this.redoStack.length > 0; }
  breakChain() { this.lastKey = null; }
}

/**
 * 고칠 수 있는 문서 초안: draft(doc) → { doc: 새 문서(얕은 복사), block(path) → 그 블록의 고칠 수 있는 복사본 (상위 표도 복사), list(path) → 블록 목록 복사본, … }
 * 같은 초안 안에서는 한 번 복사한 객체를 다시 복사하지 않음
 */
export function draft(base) {
  const doc = { ...base, body: base.body.slice() };
  const owned = new Set([doc.body]);
  const own = (o) => { owned.add(o); return o; };
  /** 경로의 블록 목록 (복사본) */
  /** 스토리 블록 목록 (복사본) — 구역 · 머리글 묶음 · 각주 묶음도 복사 */
  const storyList = (key) => {
    if (owned.has(`story:${key}`)) return storyBlocks(doc, key);
    const [kind, a, b] = key.split(':');
    let lst;
    if (kind === 'h' || kind === 'f') {
      const prop = kind === 'h' ? 'headers' : 'footers';
      const si = Number(a);
      // 구역 객체 복사 (구역 나누기 문단 또는 doc.section)
      let k = 0; let found = -1;
      doc.body.forEach((blk, i) => { if (blk.type === 'p' && blk.secBreak) { if (k === si) found = i; k++; } });
      let sec;
      if (found >= 0) { const pb = block([found]); pb.secBreak = { ...pb.secBreak, sec: { ...pb.secBreak.sec } }; sec = pb.secBreak.sec; }
      else { doc.section = { ...doc.section }; sec = doc.section; }
      sec[prop] = { ...(sec[prop] ?? {}) };
      lst = (sec[prop][b] ?? [{ type: 'p', id: uid('p'), style: kind === 'h' ? 'Header' : 'Footer', pPr: {}, runs: [] }]).slice();
      sec[prop][b] = lst;
    } else if (kind === 'n') {
      if (!owned.has(doc.notes)) { doc.notes = { ...doc.notes }; owned.add(doc.notes); }
      doc.notes[a] = { ...(doc.notes[a] ?? {}) };
      lst = (doc.notes[a][b] ?? [{ type: 'p', id: uid('p'), style: a === 'end' ? 'EndnoteText' : 'FootnoteText', pPr: {}, runs: [] }]).slice();
      doc.notes[a][b] = lst;
    }
    owned.add(lst); owned.add(`story:${key}`);
    return lst;
  };
  const list = (path) => {
    let lst = doc.body;
    if (typeof path[0] === 'string') { lst = storyList(path[0]); path = path.slice(1); }
    for (let k = 0; k + 1 < path.length; k += 3) {
      const i = path[k];
      let t = lst[i];
      if (!owned.has(t)) { t = own({ ...t, rows: t.rows.slice() }); lst[i] = t; }
      const r = path[k + 1];
      let row = t.rows[r];
      if (!owned.has(row)) { row = own({ ...row, cells: row.cells.slice() }); t.rows[r] = row; }
      const c = path[k + 2];
      let cell = row.cells[c];
      if (!owned.has(cell)) { cell = own({ ...cell, blocks: cell.blocks.slice() }); row.cells[c] = cell; owned.add(cell.blocks); }
      lst = cell.blocks;
    }
    return lst;
  };
  const block = (path) => {
    const lst = list(path);
    const i = path[path.length - 1];
    let b = lst[i];
    if (!b) return null;
    if (!owned.has(b)) {
      b = b.type === 'p' ? own({ ...b, runs: b.runs.slice(), pPr: { ...b.pPr } }) : own({ ...b, rows: b.rows.slice(), tblPr: { ...b.tblPr }, grid: b.grid.slice() });
      lst[i] = b;
    }
    return b;
  };
  /** 표 행 · 칸 고치기 (표 경로) */
  const row = (tpath, r) => { const t = block(tpath); let rw = t.rows[r]; if (!owned.has(rw)) { rw = own({ ...rw, cells: rw.cells.slice() }); t.rows[r] = rw; } return rw; };
  const cell = (tpath, r, c) => { const rw = row(tpath, r); let cl = rw.cells[c]; if (!owned.has(cl)) { cl = own({ ...cl, blocks: cl.blocks.slice() }); rw.cells[c] = cl; owned.add(cl.blocks); } return cl; };
  const prop = (k) => { if (!owned.has(doc[k])) { doc[k] = Array.isArray(base[k]) ? base[k].slice() : { ...base[k] }; owned.add(doc[k]); } return doc[k]; };
  return { doc, list, block, row, cell, prop, owned };
}

// ───────────── 편집 연산 (초안 d 에 적용, 새 위치 반환) ─────────────
/** 글 넣기 (props: 넣을 글자 서식, 없으면 그 자리 서식) */
export function insertText(d, pos, text, props = null) {
  if (!text) return pos;
  d.list(pos.p); // 스토리가 아직 없으면 만듦
  const parts = text.replace(/\r\n?/g, '\n').split('\n');
  let cur = pos;
  parts.forEach((t, i) => {
    if (i > 0) cur = splitPara(d, cur);
    if (t) cur = insertRuns(d, cur, splitTabs(t, props ?? propsAt(blockAt(d.doc, cur.p), cur.o)));
  });
  return cur;
}
const splitTabs = (t, props) => {
  const out = [];
  t.split('\t').forEach((s, i) => { if (i) out.push({ obj: 'tab', ...props }); if (s) out.push({ ...props, text: s }); });
  return out;
};
/** 런 넣기 */
export function insertRuns(d, pos, runs) {
  const p = d.block(pos.p);
  const [a, b] = splitRuns(p.runs, pos.o);
  p.runs = normalizeRuns([...a, ...runs.map((r) => ({ ...r })), ...b]);
  let n = 0;
  for (const r of runs) n += runLen(r);
  return { p: pos.p, o: pos.o + n };
}
/** 문단 나누기 (Enter): 뒤 문단은 같은 서식, 제목 스타일이면 다음 스타일 */
export function splitPara(d, pos) {
  const lst = d.list(pos.p);
  const i = pos.p[pos.p.length - 1];
  const p = d.block(pos.p);
  const [a, b] = splitRuns(p.runs, pos.o);
  const st = d.doc.styles[p.style];
  const atEnd = pos.o >= paraLen(p) && pos.o > 0;
  const next = { ...p, id: uid('p'), runs: b, pPr: { ...p.pPr } };
  if (atEnd && st?.next && st.next !== p.style) { next.style = st.next; next.pPr = p.pPr.num ? { num: p.pPr.num } : {}; }
  const mark = propsAt(p, pos.o);
  delete next.secBreak;
  if (p.secBreak) { next.secBreak = p.secBreak; }
  p.runs = a;
  if (!a.length) p.mark = mark; // 빈 앞 문단은 서식 기억
  if (p.secBreak) delete p.secBreak;
  if (!b.length) next.mark = atEnd && st?.next && st.next !== p.style ? {} : mark;
  if (p.pPr.pageBreakBefore && pos.o === 0) { /* 그대로 */ }
  d.owned.add(next);
  lst.splice(i + 1, 0, next);
  return { p: [...pos.p.slice(0, -1), i + 1], o: 0 };
}
/** 두 위치 사이 지우기 (같은 블록 목록 · 문단들) → 시작 위치 */
export function deleteRange(d, a, b) {
  [a, b] = orderPos(a, b);
  if (posEq(a, b)) return a;
  const sameList = a.p.length === b.p.length && a.p.slice(0, -1).every((v, i) => v === b.p[i]);
  if (!sameList) {
    // 표 칸을 걸친 선택: 각 문단의 해당 부분만 비움 (표 구조는 유지)
    return clearAcross(d, a, b);
  }
  const lst = d.list(a.p);
  const ia = a.p[a.p.length - 1];
  const ib = b.p[b.p.length - 1];
  const pa = d.block(a.p);
  if (ia === ib) {
    const [x] = splitRuns(pa.runs, a.o);
    const [, y] = splitRuns(pa.runs, b.o);
    pa.runs = normalizeRuns([...x, ...y]);
    if (!pa.runs.length) pa.mark = propsAt(blockAt(d.doc, a.p), a.o);
    return a;
  }
  // 사이에 표가 있으면 표째 지움 (Word 와 같음)
  const pb = lst[ib].type === 'p' ? d.block(b.p) : null;
  const [x] = splitRuns(pa.runs, a.o);
  const tail = pb ? splitRuns(pb.runs, b.o)[1] : [];
  pa.runs = normalizeRuns([...x, ...tail]);
  if (pb?.secBreak) pa.secBreak = pb.secBreak;
  lst.splice(ia + 1, ib - ia);
  return a;
}
function clearAcross(d, a, b) {
  // a ~ b 사이의 모든 문단 (같은 최상위 표 · 다른 칸): 칸 안 글만 지움
  const all = [];
  const rt = rootOf(d.doc, a.p);
  walkParas(rt.blocks, (p, path) => all.push(path), rt.base);
  const ia = all.findIndex((p) => samePath(p, a.p));
  const ib = all.findIndex((p) => samePath(p, b.p));
  for (let k = ia; k <= ib; k++) {
    const path = all[k];
    const p = d.block(path);
    const s = k === ia ? a.o : 0;
    const e = k === ib ? b.o : paraLen(p);
    const [x] = splitRuns(p.runs, s);
    const [, y] = splitRuns(p.runs, e);
    p.runs = normalizeRuns([...x, ...y]);
  }
  return a;
}
/** 앞 문단과 합치기 (Backspace 를 문단 처음에서) → 합친 위치 */
export function joinWithPrev(d, pos) {
  const i = pos.p[pos.p.length - 1];
  if (i === 0) return null;
  const lst = d.list(pos.p);
  const prevPath = [...pos.p.slice(0, -1), i - 1];
  if (lst[i - 1].type !== 'p') return null;
  const prev = d.block(prevPath);
  const cur = lst[i];
  const at = paraLen(prev);
  prev.runs = normalizeRuns([...prev.runs, ...cur.runs]);
  if (cur.secBreak) prev.secBreak = cur.secBreak; else delete prev.secBreak;
  if (!paraLen(prev) && cur.mark) prev.mark = cur.mark;
  lst.splice(i, 1);
  return { p: prevPath, o: at };
}
/** 글자 서식 적용 (a~b, 같은 문단 또는 여러 문단) — props 값 undefined 는 지움, fn(rPr) 도 가능 */
export function formatRange(d, a, b, props) {
  [a, b] = orderPos(a, b);
  const paths = rangeParas(d.doc, a, b);
  for (const path of paths) {
    const p = d.block(path);
    const s = samePath(path, a.p) ? a.o : 0;
    const e = samePath(path, b.p) ? b.o : paraLen(p);
    if (s === e && paths.length > 1) continue;
    if (s === 0 && e === paraLen(p) && !p.runs.length) { p.mark = applyProps({ ...(p.mark ?? {}) }, props); continue; }
    const [x, rest] = splitRuns(p.runs, s);
    const [mid, y] = splitRuns(rest, e - s);
    p.runs = normalizeRuns([...x, ...mid.map((r) => applyProps({ ...r }, props)), ...y]);
    if (e === paraLen(p)) p.mark = applyProps({ ...(p.mark ?? {}) }, props);
  }
}
function applyProps(r, props) {
  const v = typeof props === 'function' ? props(r) : props;
  for (const [k, x] of Object.entries(v)) { if (x === undefined || x === null) delete r[k]; else r[k] = x; }
  return r;
}
/** a~b 사이 문단 경로 (같은 블록 목록이 아니어도 문서 순서대로) */
export function rangeParas(doc, a, b) {
  [a, b] = orderPos(a, b);
  const out = [];
  let on = false;
  const { blocks, base } = rootOf(doc, a.p);
  walkParas(blocks, (p, path) => {
    if (samePath(path, a.p)) on = true;
    if (on) out.push(path);
    if (samePath(path, b.p)) on = false;
  }, base);
  return out.length ? out : [a.p];
}
/** 문단 서식 적용 (경로 목록) */
export function formatParas(d, paths, fn) {
  for (const path of paths) {
    const p = d.block(path);
    if (!p || p.type !== 'p') continue;
    const r = fn(p);
    if (r) Object.assign(p, r);
  }
}
/** 범위의 글 · 런 잘라 오기 (복사용) → 블록 목록 (첫 · 끝 문단은 잘린 부분만) */
export function sliceRange(doc, a, b) {
  [a, b] = orderPos(a, b);
  const sameList = a.p.length === b.p.length && a.p.slice(0, -1).every((v, i) => v === b.p[i]);
  if (!sameList) {
    const out = [];
    for (const path of rangeParas(doc, a, b)) {
      const p = blockAt(doc, path);
      const s = samePath(path, a.p) ? a.o : 0;
      const e = samePath(path, b.p) ? b.o : paraLen(p);
      out.push({ ...clone(p), id: uid('p'), runs: clone(splitRuns(splitRuns(p.runs, s)[1], e - s)[0]) });
    }
    return out;
  }
  const lst = containerOf(doc, a.p);
  const ia = a.p[a.p.length - 1];
  const ib = b.p[b.p.length - 1];
  const out = [];
  for (let i = ia; i <= ib; i++) {
    const blk = lst[i];
    if (blk.type !== 'p') { out.push(clone(blk)); continue; }
    const s = i === ia ? a.o : 0;
    const e = i === ib ? b.o : paraLen(blk);
    const runs = splitRuns(splitRuns(blk.runs, s)[1], e - s)[0];
    const c = { ...clone(blk), id: uid('p'), runs: clone(runs) };
    delete c.secBreak;
    out.push(c);
  }
  return out;
}
/** 블록 목록 붙여넣기 (pos 에서) → 끝 위치 */
export function insertBlocks(d, pos, blocks) {
  if (!blocks.length) return pos;
  blocks = blocks.map((b) => reId(clone(b)));
  // 한 문단 조각: 그 자리에 런만 넣음
  if (blocks.length === 1 && blocks[0].type === 'p') return insertRuns(d, pos, blocks[0].runs);
  const lst = d.list(pos.p);
  const inTable = pos.p.length > 1;
  if (inTable && blocks.some((b) => b.type === 'tbl')) blocks = blocks.flatMap((b) => (b.type === 'tbl' ? tableToParas(b) : [b]));
  // 문단 나누고 사이에 넣기: 첫 문단은 앞쪽 문단에 이어 붙이고, 마지막 문단은 뒤쪽 문단 앞에 붙임
  const after = splitPara(d, pos);
  const i = pos.p[pos.p.length - 1];
  const head = d.block(pos.p);
  const first = blocks[0];
  let k = 0;
  if (first.type === 'p') { head.runs = normalizeRuns([...head.runs, ...first.runs]); if (!paraLen(head) || pos.o === 0) { head.style = first.style; head.pPr = { ...first.pPr }; } k = 1; }
  const last = blocks[blocks.length - 1];
  const tailPath = after.p;
  let endOff = 0;
  const middle = blocks.slice(k, last.type === 'p' && blocks.length - 1 >= k ? blocks.length - 1 : blocks.length);
  if (last.type === 'p' && blocks.length - 1 >= k) {
    const tail = d.block(tailPath);
    endOff = paraLen(last);
    tail.runs = normalizeRuns([...last.runs, ...tail.runs]);
    if (!tail.runs.length || endOff) { tail.style = last.style; tail.pPr = { ...last.pPr }; }
  }
  for (const m of middle) d.owned.add(m);
  lst.splice(i + 1, 0, ...middle);
  const endIdx = i + 1 + middle.length;
  return { p: [...pos.p.slice(0, -1), endIdx], o: endOff };
}
function reId(b) {
  if (b.type === 'p') b.id = uid('p');
  else if (b.type === 'tbl') { b.id = uid('t'); for (const r of b.rows) for (const c of r.cells) c.blocks = c.blocks.map(reId); }
  return b;
}
export function tableToParas(t, sep = '\t') {
  return t.rows.map((r) => {
    const runs = [];
    r.cells.forEach((c, i) => { if (i) runs.push({ obj: 'tab' }); runs.push({ text: c.blocks.filter((b) => b.type === 'p').map(paraPlain).join(' ') }); });
    void sep;
    return para(normalizeRuns(runs));
  });
}

// ───────────── 단어 · 문장 경계 ─────────────
const isWordChar = (ch) => /[\p{L}\p{N}_]/u.test(ch);
export function wordAt(p, o) {
  const t = paraText(p);
  let s = o; let e = o;
  if (s > 0 && !isWordChar(t[s] ?? '') && isWordChar(t[s - 1])) s--;
  if (!isWordChar(t[s] ?? '')) { e = Math.min(t.length, s + 1); while (e < t.length && t[e] === ' ') e++; return [s, e]; }
  while (s > 0 && isWordChar(t[s - 1])) s--;
  while (e < t.length && isWordChar(t[e])) e++;
  while (e < t.length && t[e] === ' ') e++; // Word: 단어 뒤 공백도 선택
  return [s, e];
}
/** Ctrl+← / → : 단어 처음으로 */
export function wordStep(p, o, dir) {
  const t = paraText(p);
  if (dir > 0) {
    let i = o;
    if (isWordChar(t[i] ?? '')) while (i < t.length && isWordChar(t[i])) i++;
    else if (i < t.length && t[i] !== ' ') i++;
    while (i < t.length && t[i] === ' ') i++;
    return i;
  }
  let i = o;
  while (i > 0 && t[i - 1] === ' ') i--;
  if (i > 0 && isWordChar(t[i - 1])) while (i > 0 && isWordChar(t[i - 1])) i--;
  else if (i > 0) i--;
  return i;
}
export function sentenceAt(p, o) {
  const t = paraText(p);
  let s = o; let e = o;
  while (s > 0 && !/[.!?。]\s/.test(t.slice(s - 2, s))) s--;
  while (e < t.length && !/[.!?。]/.test(t[e])) e++;
  if (e < t.length) e++;
  while (e < t.length && t[e] === ' ') e++;
  return [s, e];
}

// ───────────── 찾기 ─────────────
/** 찾기 → [{ p (경로), s, e }] (문단 단위, 문단을 넘는 일치는 없음) */
export function findAll(doc, query, { matchCase = false, wholeWord = false, regex = false, wildcards = false } = {}) {
  if (!query) return [];
  let re;
  try {
    if (regex) re = new RegExp(query, matchCase ? 'gu' : 'giu');
    else if (wildcards) re = new RegExp(query.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*?').replace(/\?/g, '.'), matchCase ? 'gu' : 'giu');
    else re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'gu' : 'giu');
  } catch { return []; }
  const out = [];
  walkParas(doc.body, (p, path) => {
    const t = paraText(p);
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(t))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      if (wholeWord && ((m.index > 0 && isWordChar(t[m.index - 1])) || isWordChar(t[m.index + m[0].length] ?? ''))) continue;
      out.push({ p: path, s: m.index, e: m.index + m[0].length, text: m[0], groups: m.slice(1) });
    }
  });
  return out;
}

// ───────────── 구역 ─────────────
/** 블록 → 구역 목록 [{ sec, start, end }] (최상위 블록 번호 범위, 구역 나누기는 그 문단에 저장 = docx 와 같음) */
export function sections(doc) {
  const out = [];
  let start = 0;
  doc.body.forEach((b, i) => {
    if (b.type === 'p' && b.secBreak) { out.push({ sec: b.secBreak.sec, kind: b.secBreak.kind ?? 'nextPage', start, end: i + 1 }); start = i + 1; }
  });
  out.push({ sec: doc.section, kind: doc.sectionKind ?? 'nextPage', start, end: doc.body.length });
  // 구역 시작 종류는 그 구역의 것 (앞 구역 나누기의 종류가 아님): docx 는 sectPr 의 type 이 그 구역의 시작 방식
  return out;
}
/** 최상위 블록 번호 i 의 구역 */
export function sectionOf(doc, i) {
  for (const s of sections(doc)) if (i >= s.start && i < s.end) return s;
  return sections(doc).pop();
}

// ───────────── 목차 (TOC) ─────────────
/** 제목 문단 목록 [{ p, path, level, text }] */
export function headings(doc, maxLevel = 9) {
  const out = [];
  doc.body.forEach((b, i) => {
    if (b.type !== 'p') return;
    const lv = headingLevel(doc, b);
    if (lv && lv <= maxLevel) out.push({ p: b, path: [i], level: lv, text: paraPlain(b).trim() });
  });
  return out;
}
/** 목차 블록 만들기 (쪽 번호는 pageOf(id) 로, 없으면 빈칸) */
export function buildToc(doc, { levels = 3, pageOf = null, title = '목차', tocId = uid('toc'), links = true, dots = true } = {}) {
  const out = [];
  if (title) out.push({ ...para(title, { style: 'TOCHeading' }), toc: tocId });
  for (const h of headings(doc, levels)) {
    if (h.p.toc) continue;
    const pg = pageOf ? pageOf(h.p.id) : null;
    const bm = h.p.id;
    out.push({
      ...para([{ text: h.text || ' ', ...(links ? { link: { anchor: bm } } : {}) }, { obj: 'tab' }, { obj: 'field', code: `PAGEREF ${bm} \\h`, result: pg != null ? String(pg) : '', ...(links ? { link: { anchor: bm } } : {}) }], { style: `TOC${h.level}`, pPr: { tabs: [{ pos: 9999, align: 'right', leader: dots ? 'dot' : 'none' }] } }),
      toc: tocId,
    });
  }
  if (out.length <= (title ? 1 : 0)) out.push({ ...para('목차 항목이 없습니다.', {}), toc: tocId });
  return out;
}

// ───────────── 스냅숏 (저장 · 자동 저장 · MCP) ─────────────
export function serialize(doc) { return JSON.stringify(doc); }
export function deserialize(json) {
  const d = typeof json === 'string' ? JSON.parse(json) : json;
  const base = newDoc();
  const doc = { ...base, ...d, styles: { ...builtinStyles(), ...(d.styles ?? {}) }, defaults: d.defaults ?? base.defaults, section: { ...newSection(), ...(d.section ?? {}) } };
  if (!doc.body?.length) doc.body = [para('')];
  return doc;
}
