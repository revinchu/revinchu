// 내부 모델은 96dpi 정수 CSS px. Excel 대화상자의 행 pt/열 문자 수로 표시한다.
// 아래 글꼴 MDW 추정은 xlsx.js의 digitWidth와 일치해야 한다(회귀 테스트로 검사).
// 브라우저/OS에서 실제 글꼴을 측정하지 않으므로 미등록·대체 글꼴에는 오차가 있을 수 있다.
const DIGIT_EM = { calibri: 0.507, 'calibri light': 0.49, '맑은 고딕': 0.55, 'malgun gothic': 0.55, arial: 0.556, '굴림': 0.5, gulim: 0.5, '굴림체': 0.5, '돋움': 0.5, dotum: 0.5, '돋움체': 0.5, '바탕': 0.5, batang: 0.5, '나눔고딕': 0.6, nanumgothic: 0.6, 'nanum gothic': 0.6, 'times new roman': 0.5, cambria: 0.556, 'segoe ui': 0.55, verdana: 0.636, tahoma: 0.546, 'meiryo ui': 0.55, 'ms gothic': 0.5, simsun: 0.5 };
export const MAX_ROW_POINTS = 409.5;
export const MAX_COLUMN_CHARS = 255;
export const fontDigitWidth = (font) => Math.max(4, Math.round(((font?.size || 11) * 96) / 72 * (DIGIT_EM[String(font?.name ?? '맑은 고딕').toLowerCase()] ?? 0.53)));

function checked(value, max, label) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > max) {
    const error = new RangeError(`${label}는 0 이상 ${max} 이하의 숫자로 입력하세요.`);
    error.code = 'INVALID_DIMENSION';
    throw error;
  }
  return value;
}

export const rowPointsToPixels = (points) => Math.round(checked(points, MAX_ROW_POINTS, '행 높이') * 4 / 3);
export const rowPixelsToPoints = (pixels) => Math.round(checked(pixels, Number.MAX_SAFE_INTEGER, '행 높이') * 75) / 100;

function mdwOf(font) {
  if (typeof font !== 'number') return fontDigitWidth(font);
  if (!Number.isFinite(font) || font <= 0) throw new RangeError('글꼴의 숫자 너비가 올바르지 않습니다.');
  return font;
}

/** UI 문자 수는 OOXML <col width>와 다르다. 1자 이상은 좌우 여백+격자선 5px를 더한다. */
export function columnCharsToPixels(chars, font) {
  checked(chars, MAX_COLUMN_CHARS, '열 너비');
  const mdw = mdwOf(font);
  return Math.round(chars < 1 ? chars * (mdw + 5) : chars * mdw + 5);
}

export function pixelsToColumnChars(pixels, font) {
  checked(pixels, Number.MAX_SAFE_INTEGER, '열 너비');
  const mdw = mdwOf(font);
  return Math.round((pixels < mdw + 5 ? pixels / (mdw + 5) : (pixels - 5) / mdw) * 100) / 100;
}
