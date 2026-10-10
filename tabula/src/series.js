// 채우기 핸들 / 연속 데이터 생성
import { shiftFormula } from './formula.js';

export const CUSTOM_LISTS = [
  ['일', '월', '화', '수', '목', '금', '토'],
  ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['1분기', '2분기', '3분기', '4분기'],
  ['갑', '을', '병', '정', '무', '기', '경', '신', '임', '계'],
];
const DATE_FMTS = new Set(['date', 'longdate', 'datetime']);

const clean = (n) => Number(n.toPrecision(15));
const isFormula = (d) => !!d && d.raw.startsWith('=');

/**
 * seq: 원본 셀 [{ data: {raw, style, comment}|null, value, pos }] (pos = 행 또는 열 번호)
 * forward: 아래/오른쪽 방향이면 true
 * axis: 'row'(세로 채우기) | 'col'(가로 채우기)
 * 반환: (k, targetPos) → 새 셀 데이터
 */
export function makeSeries(seq, forward, axis) {
  const n = seq.length;
  const styleOf = (i) => seq[i].data?.style;
  const cyc = (k) => (forward ? k % n : n - 1 - (k % n));
  const withStyle = (raw, k) => {
    const src = seq[cyc(k)].data;
    return raw === '' && !src?.style ? null : { raw, style: src?.style ? { ...src.style } : undefined };
  };

  const plain = seq.every((s) => s.data && s.data.raw !== '' && !isFormula(s.data));

  // 숫자 계열
  if (plain && seq.every((s) => typeof s.value === 'number')) {
    const vals = seq.map((s) => s.value);
    let step = 0;
    if (n > 1) step = (vals[n - 1] - vals[0]) / (n - 1);
    else if (DATE_FMTS.has(styleOf(0)?.numFmt)) step = 1;
    if (step !== 0 || n > 1) {
      return (k) => {
        const v = forward ? vals[n - 1] + step * (k + 1) : vals[0] - step * (k + 1);
        const src = seq[cyc(k)].data;
        const style = src.style ? { ...src.style } : undefined;
        if (DATE_FMTS.has(style?.numFmt) || style?.numFmt === 'time') {
          // 날짜/시간 서식은 표시 형식으로 보이도록 숫자 그대로 저장
          return { raw: String(clean(v)), style };
        }
        return { raw: String(clean(v)), style };
      };
    }
  }

  // 목록 (요일, 월 등)
  if (plain) {
    for (const list of CUSTOM_LISTS) {
      const idx = seq.map((s) => list.indexOf(s.data.raw));
      if (idx.every((i) => i >= 0)) {
        const step = n > 1 ? idx[n - 1] - idx[n - 2] : 1;
        const L = list.length;
        return (k) => {
          const base = forward ? idx[n - 1] + step * (k + 1) : idx[0] - step * (k + 1);
          return withStyle(list[((base % L) + L) % L], k);
        };
      }
    }
  }

  // "항목1" 처럼 끝에 숫자가 붙은 텍스트
  if (plain) {
    const parts = seq.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s.data.raw));
    if (parts.every(Boolean) && parts.every((p) => p[1] === parts[0][1] && p[3] === parts[0][3])
      && seq.every((s) => typeof s.value === 'string')) {
      const nums = parts.map((p) => Number(p[2]));
      const width = parts[0][2].startsWith('0') ? parts[0][2].length : 0;
      const step = n > 1 ? (nums[n - 1] - nums[0]) / (n - 1) : 1;
      return (k) => {
        const v = Math.max(0, Math.round(forward ? nums[n - 1] + step * (k + 1) : nums[0] - step * (k + 1)));
        return withStyle(parts[0][1] + String(v).padStart(width, '0') + parts[0][3], k);
      };
    }
  }

  // 그대로 복사 (수식은 상대 참조 이동)
  return (k, targetPos) => {
    const src = seq[cyc(k)];
    if (!src.data) return null;
    const delta = targetPos - src.pos;
    const raw = isFormula(src.data)
      ? shiftFormula(src.data.raw, axis === 'row' ? delta : 0, axis === 'col' ? delta : 0)
      : src.data.raw;
    return { ...src.data, raw, style: src.data.style ? { ...src.data.style } : undefined };
  };
}
