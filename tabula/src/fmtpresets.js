// 셀 서식 대화상자의 '표시 형식' 범주와 서식 코드 만들기 (DOM 없음)

const dec = (n) => (n > 0 ? `.${'0'.repeat(n)}` : '');

export const CURRENCY_SYMBOLS = [
  { id: '₩', label: '₩ 한국어' },
  { id: '$', label: '$ 영어(미국)' },
  { id: '¥', label: '¥ 일본어' },
  { id: '€', label: '€ 유로' },
  { id: '', label: '없음' },
];

/** 음수 표시 형식 (숫자·통화 공통) */
export const NEGATIVE_STYLES = [
  { id: 'minus', label: '-1,234.10', make: (b) => b },
  { id: 'red', label: '1,234.10 (빨강)', red: true, make: (b) => `${b};[빨강]${b}` },
  { id: 'minusRed', label: '-1,234.10 (빨강)', red: true, make: (b) => `${b};[빨강]-${b}` },
  { id: 'paren', label: '(1,234.10)', make: (b) => `${b}_);(${b})` },
  { id: 'parenRed', label: '(1,234.10) (빨강)', red: true, make: (b) => `${b}_);[빨강](${b})` },
];

export const DATE_TYPES = [
  'yyyy-mm-dd', 'yyyy"년" m"월" d"일"', 'yyyy"년" m"월" d"일" aaaa', 'yyyy. m. d.', 'yyyy.mm.dd', 'yyyy/mm/dd',
  'yy-mm-dd', 'm"월" d"일"', 'mm"월" dd"일"', 'm/d', 'mm-dd', 'yyyy"년" m"월"', 'yyyy-mm', 'aaaa', 'aaa',
  'd-mmm-yy', 'mmm-yy', 'mmmm d, yyyy', 'yyyymmdd', 'yyyy-mm-dd h:mm', '[$-412]yyyy-mm-dd AM/PM h:mm',
];

export const TIME_TYPES = [
  'h:mm', 'h:mm:ss', '[$-412]AM/PM h:mm', '[$-412]AM/PM h:mm:ss', 'h"시" mm"분"', 'h"시" mm"분" ss"초"',
  '[$-412]AM/PM h"시" mm"분"', '[h]:mm:ss', 'mm:ss', 'mm:ss.0', 'h:mm AM/PM', 'yyyy-mm-dd h:mm:ss',
];

export const FRACTION_TYPES = [
  { code: '# ?/?', label: '한 자릿수 분모 (1/4)' },
  { code: '# ??/??', label: '두 자릿수 분모 (21/25)' },
  { code: '# ???/???', label: '세 자릿수 분모 (312/943)' },
  { code: '# ?/2', label: '2분의 (1/2)' },
  { code: '# ?/4', label: '4분의 (2/4)' },
  { code: '# ?/8', label: '8분의 (4/8)' },
  { code: '# ??/16', label: '16분의 (8/16)' },
  { code: '# ?/10', label: '10분의 (3/10)' },
  { code: '# ??/100', label: '100분의 (30/100)' },
];

export const SPECIAL_TYPES = [
  { code: '000-000', label: '우편 번호 (구)' },
  { code: '00000', label: '우편 번호' },
  { code: '[<=9999999]###-####;(0##) ###-####', label: '전화 번호 (국번 3자리)' },
  { code: '[<=99999999]####-####;(0##) ####-####', label: '전화 번호 (국번 4자리)' },
  { code: '000-0000-0000', label: '휴대폰 번호' },
  { code: '000000-0000000', label: '주민 등록 번호' },
  { code: '000-00-00000', label: '사업자 등록 번호' },
  { code: '#,##0"원"', label: '금액 (원)' },
  { code: '#,##0,"천원"', label: '금액 (천원 단위)' },
  { code: '#,##0,,"백만"', label: '금액 (백만 단위)' },
];

/** 사용자 지정 목록 (엑셀 기본 목록과 비슷하게) */
export const CUSTOM_LIST = [
  'G/표준', '0', '0.00', '#,##0', '#,##0.00', '#,##0;[빨강]-#,##0', '#,##0.00;[빨강]-#,##0.00',
  '"₩"#,##0', '"₩"#,##0;[빨강]"₩"-#,##0', '_-"₩"* #,##0_-;-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-',
  '0%', '0.00%', '0.00E+00', '##0.0E+0', '# ?/?', '# ??/??',
  'yyyy-mm-dd', 'yyyy"년" m"월" d"일"', 'mm"월" dd"일"', 'yy/mm/dd', 'd-mmm-yy', 'mmm-yy',
  'h:mm', 'h:mm:ss', '[$-412]AM/PM h:mm', '[h]:mm:ss', 'mm:ss.0', 'yyyy-mm-dd h:mm',
  '@', '@"님"', '#,##0"원"', '#,##0"개"', '0"%"', '[파랑]#,##0;[빨강]-#,##0;0', '[>=1000000]0.0,,"백만";[>=1000]0.0,"천";0',
];

export const CATEGORIES = [
  { id: 'general', label: '일반', note: '일반 서식 셀은 특정 서식을 지정하지 않습니다.' },
  { id: 'number', label: '숫자', note: '숫자는 일반적인 숫자를 나타내는 데 사용됩니다. 통화 및 회계 서식은 화폐 값을 표시합니다.' },
  { id: 'currency', label: '통화', note: '통화 서식은 일반 화폐 값에 사용됩니다. 소수점을 맞추려면 회계 서식을 사용하세요.' },
  { id: 'accounting', label: '회계', note: '회계 서식은 열에 있는 통화 기호와 소수점을 맞춥니다.' },
  { id: 'date', label: '날짜', note: '날짜 서식은 날짜 및 시간 일련 번호를 날짜 값으로 표시합니다.' },
  { id: 'time', label: '시간', note: '시간 서식은 날짜 및 시간 일련 번호를 시간 값으로 표시합니다.' },
  { id: 'percent', label: '백분율', note: '백분율 서식은 셀 값에 100을 곱한 결과를 % 기호와 함께 표시합니다.' },
  { id: 'fraction', label: '분수' },
  { id: 'scientific', label: '지수' },
  { id: 'text', label: '텍스트', note: '텍스트 서식 셀에서는 숫자도 텍스트로 취급되어 입력한 그대로 표시됩니다.' },
  { id: 'special', label: '기타', note: '기타 서식은 우편 번호, 전화 번호, 등록 번호 등에 유용합니다.' },
  { id: 'custom', label: '사용자 지정', note: '기존의 서식 코드 중 하나를 선택하거나 새로 입력하세요.' },
];

/**
 * 범주 + 옵션 → 서식 코드
 * opts: { decimals, thousands, negative, symbol, type }
 */
export function buildCode(cat, o = {}) {
  const d = Math.max(0, Math.min(30, Number(o.decimals ?? 2) || 0));
  switch (cat) {
    case 'general': return 'General';
    case 'number': {
      const base = `${o.thousands ? '#,##0' : '0'}${dec(d)}`;
      return (NEGATIVE_STYLES.find((n) => n.id === o.negative) ?? NEGATIVE_STYLES[0]).make(base);
    }
    case 'currency': {
      const sym = o.symbol ?? '₩';
      const base = `${sym ? `"${sym}"` : ''}#,##0${dec(d)}`;
      return (NEGATIVE_STYLES.find((n) => n.id === o.negative) ?? NEGATIVE_STYLES[0]).make(base);
    }
    case 'accounting': {
      const sym = o.symbol ?? '₩';
      const s = sym ? `"${sym}"` : '';
      const n = `#,##0${dec(d)}`;
      return `_-${s}* ${n}_-;-${s}* ${n}_-;_-${s}* "-"${d ? '?'.repeat(d) : ''}_-;_-@_-`;
    }
    case 'percent': return `0${dec(d)}%`;
    case 'scientific': return `0${dec(d)}E+00`;
    case 'text': return '@';
    case 'date': case 'time': case 'fraction': case 'special': case 'custom':
      return o.type || 'General';
    default: return 'General';
  }
}

/** 서식 코드 → 범주와 옵션 추정 (대화상자 초기값) */
export function describeCode(code) {
  const c = String(code ?? '').trim();
  if (!c || /^(general|g\/표준)$/i.test(c)) return { cat: 'general' };
  if (c === '@') return { cat: 'text' };
  if (DATE_TYPES.includes(c)) return { cat: 'date', type: c };
  if (TIME_TYPES.includes(c)) return { cat: 'time', type: c };
  if (FRACTION_TYPES.some((f) => f.code === c)) return { cat: 'fraction', type: c };
  if (SPECIAL_TYPES.some((f) => f.code === c)) return { cat: 'special', type: c };
  for (let d = 0; d <= 10; d++) {
    if (c === buildCode('percent', { decimals: d })) return { cat: 'percent', decimals: d };
    if (c === buildCode('scientific', { decimals: d })) return { cat: 'scientific', decimals: d };
    for (const neg of NEGATIVE_STYLES) {
      for (const thousands of [true, false]) if (c === buildCode('number', { decimals: d, thousands, negative: neg.id })) return { cat: 'number', decimals: d, thousands, negative: neg.id };
      for (const s of CURRENCY_SYMBOLS) if (c === buildCode('currency', { decimals: d, symbol: s.id, negative: neg.id })) return { cat: 'currency', decimals: d, symbol: s.id, negative: neg.id };
    }
    for (const s of CURRENCY_SYMBOLS) if (c === buildCode('accounting', { decimals: d, symbol: s.id })) return { cat: 'accounting', decimals: d, symbol: s.id };
  }
  return { cat: 'custom', type: c };
}
