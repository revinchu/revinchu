// 문서 안의 하이퍼링크 참조만 해석한다. 이름의 수식 평가나 문서 변경은 하지 않는다.
import { parseRangeName, parseCellName, quoteSheetName, MAX_ROWS, MAX_COLS } from './formula.js';

const CONTROL = /[\u0000-\u001f\u007f-\u009f]/;
const NAME = /^[A-Za-z_\\\u00c0-\uffff][\w.\\\u00c0-\uffff]*$/;
const fail = (error) => ({ error });

function linkParts(value, nameRef = false) {
  if (typeof value !== 'string' || value.length > 8192 || CONTROL.test(value)) return fail('하이퍼링크 주소가 올바르지 않습니다.');
  let text = value.trim();
  if (text.startsWith('#')) text = text.slice(1);
  if (nameRef && text.startsWith('=')) text = text.slice(1).trim();
  if (!text) return fail('이동할 셀 주소나 이름이 없습니다.');
  if (text === '#REF!') return fail('링크가 가리키던 셀 또는 시트가 삭제되었습니다.');
  if (/[\[\]]/.test(text)) return fail('외부 통합 문서나 표의 구조적 참조로는 이동할 수 없습니다.');
  let sheetName = null, ref = text;
  if (text.includes('!')) {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text);
    if (!m) return fail('시트 참조가 올바르지 않습니다.');
    sheetName = (m[1] === undefined ? m[2] : m[1].replace(/''/g, "'"));
    ref = m[3].trim();
  }
  if (ref === '#REF!') return fail('링크가 가리키던 셀 또는 시트가 삭제되었습니다.');
  let range = parseRangeName(ref);
  if (!range) {
    const cols = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(ref);
    const rows = /^\$?(\d+):\$?(\d+)$/.exec(ref);
    if (cols) {
      const a = parseCellName(cols[1] + '1'), b = parseCellName(cols[2] + '1');
      if (a && b) range = { r1: 0, c1: Math.min(a.c, b.c), r2: MAX_ROWS - 1, c2: Math.max(a.c, b.c) };
    } else if (rows) {
      const a = Number(rows[1]) - 1, b = Number(rows[2]) - 1;
      if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) range = { r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAX_COLS - 1 };
    }
  }
  if (!range && !NAME.test(ref)) return fail(nameRef ? '계산식 또는 여러 영역으로 정의된 이름은 링크로 이동할 수 없습니다.' : '셀 주소나 이름이 올바르지 않습니다.');
  // A0/XFE1 같은 잘못된 셀 주소를 이름으로 잘못 취급하지 않는다.
  if (!range && /^\$?[A-Za-z]{1,3}\$?\d+$/.test(ref)) return fail('셀 주소가 시트 범위를 벗어났습니다.');
  return { sheetName, ref, range, text };
}

function decodedLink(value) {
  if (typeof value !== 'string' || !/%[0-9a-f]{2}/i.test(value)) return null;
  try { const decoded = decodeURIComponent(value); return decoded === value ? null : decoded; } catch { return null; }
}

/** 내부 #fragment 또는 내부 참조 본문 → 0-based 시트 번호와 사각형. 문서는 읽기만 한다. */
export function resolveWorkbookLink(wb, si, target) {
  if (!Number.isInteger(si) || !wb.sheets?.[si]) return fail('현재 시트를 찾을 수 없습니다.');
  const resolve = (value) => {
    let host = si, current = value, nameRef = false;
    const seen = new Set();
    for (let depth = 0; depth < 64; depth++) {
      const part = linkParts(current, nameRef);
      if (part.error) return part;
      let sheet = host;
      if (part.sheetName !== null) {
        sheet = wb.sheetIndexByName(part.sheetName);
        if (sheet < 0) return fail('링크가 가리키는 시트를 찾을 수 없습니다.');
      }
      if (part.range) {
        if (['hidden', 'veryHidden'].includes(wb.sheets[sheet].state)) return fail('숨겨진 시트입니다. 시트의 숨기기를 취소한 뒤 이동하세요.');
        return { sheet, range: part.range };
      }
      const entry = wb.findName(part.ref, host, part.sheetName);
      if (!entry) return fail('링크가 가리키는 이름을 찾을 수 없습니다.');
      if (seen.has(entry)) return fail('이름의 참조가 서로 반복되어 이동할 수 없습니다.');
      seen.add(entry);
      if (entry.sheet) {
        host = wb.sheetIndexByName(entry.sheet);
        if (host < 0) return fail('이름이 정의된 시트를 찾을 수 없습니다.');
      }
      current = entry.ref;
      nameRef = true;
    }
    return fail('이름의 참조 단계가 너무 많아 이동할 수 없습니다.');
  };
  const literal = resolve(target);
  if (!literal.error) return literal;
  // 리터럴 이름이 실제로 있으면 숨김/동적 참조 등의 오류도 그 대상의 결과이다.
  const original = linkParts(target);
  if (!original.error && ((original.sheetName !== null && wb.sheetIndexByName(original.sheetName) >= 0)
    || (!original.range && wb.findName(original.ref, si, original.sheetName)))) return literal;
  const decoded = decodedLink(target);
  return decoded === null ? literal : resolve(decoded);
}

/** 구조 변경의 정적 참조 재작성. 외부 주소·동적 식은 그대로 보존한다. */
export function rewriteWorkbookLink(target, transform, sheetExists, allowBare = false) {
  if (typeof target !== 'string' || (!target.startsWith('#') && !allowBare)) return target;
  let part = linkParts(target);
  const decoded = decodedLink(target);
  if (decoded !== null && (part.error || (part.sheetName !== null && !sheetExists(part.sheetName)))) {
    const candidate = linkParts(decoded);
    if (!candidate.error) part = candidate;
  }
  if (part.error) return target;
  const formula = '=' + (part.sheetName === null ? part.ref : quoteSheetName(part.sheetName) + '!' + part.ref);
  const result = transform(formula);
  return result === formula ? target : (target.startsWith('#') ? '#' : '') + result.slice(1);
}
