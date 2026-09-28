// 데이터 유효성 검사 (DOM 없음)
// 규칙: { r1,c1,r2,c2, type, op, f1, f2, allowBlank, showDropdown, errorStyle, errorTitle, error, promptTitle, prompt }
import { parse, evaluateFormula, shiftFormula, isError, parseRangeName } from './formula.js';
import { parseInput, formatValue } from './format.js';

export const VALIDATION_TYPES = [
  { id: 'any', label: '모든 값' },
  { id: 'whole', label: '정수' },
  { id: 'decimal', label: '소수점' },
  { id: 'list', label: '목록' },
  { id: 'date', label: '날짜' },
  { id: 'time', label: '시간' },
  { id: 'textLength', label: '텍스트 길이' },
  { id: 'custom', label: '사용자 지정' },
];

export const VALIDATION_OPS = [
  { id: 'between', label: '해당 범위' },
  { id: 'notBetween', label: '제외 범위' },
  { id: 'equal', label: '=' },
  { id: 'notEqual', label: '<>' },
  { id: 'greaterThan', label: '>' },
  { id: 'lessThan', label: '<' },
  { id: 'greaterThanOrEqual', label: '>=' },
  { id: 'lessThanOrEqual', label: '<=' },
];

export function validationAt(sheet, r, c) {
  const list = sheet.validations ?? [];
  for (let i = list.length - 1; i >= 0; i--) {
    const v = list[i];
    if (r >= v.r1 && r <= v.r2 && c >= v.c1 && c <= v.c2) return v;
  }
  return null;
}

/** 조건 값 계산: 숫자/날짜 문자열 또는 '=수식' (규칙 범위 왼쪽 위 기준 상대 참조) */
function operand(wb, si, rule, text, r, c) {
  if (text === undefined || text === null || text === '') return null;
  const t = String(text).trim();
  const formula = t.startsWith('=') ? t : /^[A-Za-z$]/.test(t) && !/^(TRUE|FALSE)$/i.test(t) ? `=${t}` : null;
  if (formula) {
    try {
      const shifted = shiftFormula(formula, r - rule.r1, c - rule.c1);
      return evaluateFormula(parse(shifted.slice(1)), wb.ctxFor(si));
    } catch {
      return null;
    }
  }
  return parseInput(t).value;
}

function compare(op, v, a, b) {
  switch (op) {
    case 'notBetween': return v < Math.min(a, b) || v > Math.max(a, b);
    case 'equal': return v === a;
    case 'notEqual': return v !== a;
    case 'greaterThan': return v > a;
    case 'lessThan': return v < a;
    case 'greaterThanOrEqual': return v >= a;
    case 'lessThanOrEqual': return v <= a;
    default: return v >= Math.min(a, b) && v <= Math.max(a, b);
  }
}

/** 목록 항목: "사과,배" 또는 범위 참조 ($A$1:$A$5, Sheet2!A1:A3) */
export function listItems(wb, si, rule) {
  const src = String(rule.f1 ?? '').trim().replace(/^=/, '');
  if (!src) return [];
  if (src.startsWith('"') && src.endsWith('"')) return splitList(src.slice(1, -1));
  const bang = src.lastIndexOf('!');
  const sheetName = bang > 0 ? src.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : null;
  const rg = parseRangeName((bang > 0 ? src.slice(bang + 1) : src).replace(/\$/g, ''));
  if (rg) {
    const s = sheetName ? wb.sheetIndexByName(sheetName) : si;
    if (s < 0) return [];
    const out = [];
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 5000); r++) {
      for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) {
        const t = formatValue(wb.getValue(s, r, c), wb.styleAt(s, r, c)).text;
        if (t !== '' && !out.includes(t)) out.push(t);
      }
    }
    return out;
  }
  return splitList(src);
}

const splitList = (s) => s.split(',').map((x) => x.trim()).filter((x) => x !== '');

/**
 * 입력한 텍스트가 규칙을 만족하는지. 반환: true | false
 * value: 입력 해석 결과 (숫자, 문자열, 논리값, null)
 */
export function checkValidation(wb, si, rule, r, c, text) {
  if (!rule || rule.type === 'any') return true;
  const blank = text === '' || text === null || text === undefined;
  if (blank) return rule.allowBlank !== false;
  if (String(text).startsWith('=')) return true; // 수식은 입력 시 검사하지 않음
  const value = parseInput(String(text)).value;
  switch (rule.type) {
    case 'list': {
      const items = listItems(wb, si, rule);
      const shown = typeof value === 'number' ? String(value) : String(text);
      return items.some((it) => it === String(text) || it === shown || it.toLowerCase() === String(text).toLowerCase());
    }
    case 'whole': case 'decimal': case 'date': case 'time': {
      if (typeof value !== 'number') return false;
      if (rule.type === 'whole' && !Number.isInteger(value)) return false;
      const a = operand(wb, si, rule, rule.f1, r, c);
      const b = operand(wb, si, rule, rule.f2, r, c);
      if (typeof a !== 'number') return true;
      if ((rule.op === 'between' || rule.op === 'notBetween' || !rule.op) && typeof b !== 'number') return true;
      return compare(rule.op ?? 'between', value, a, b);
    }
    case 'textLength': {
      const len = [...String(text)].length;
      const a = operand(wb, si, rule, rule.f1, r, c);
      const b = operand(wb, si, rule, rule.f2, r, c);
      if (typeof a !== 'number') return true;
      return compare(rule.op ?? 'between', len, a, typeof b === 'number' ? b : a);
    }
    case 'custom': {
      // 사용자 지정 수식은 입력값을 반영한 뒤 계산해야 하므로 임시로 셀 값을 바꿔 평가
      const key = `${r},${c}`;
      const sheet = wb.sheets[si];
      const prev = sheet.cells.get(key);
      sheet.cells.set(key, { raw: String(text), v: value, style: prev?.style });
      wb.invalidate(si);
      let ok = true;
      try {
        const res = operand(wb, si, rule, String(rule.f1 ?? '').startsWith('=') ? rule.f1 : `=${rule.f1}`, r, c);
        ok = res === true || (typeof res === 'number' && res !== 0) || res === null;
        if (isError(res)) ok = false;
      } finally {
        if (prev) sheet.cells.set(key, prev); else sheet.cells.delete(key);
        wb.invalidate(si);
      }
      return ok;
    }
    default:
      return true;
  }
}

/** 규칙 설명 (오류 메시지가 없을 때) */
export function describeRule(rule) {
  const t = VALIDATION_TYPES.find((x) => x.id === rule.type)?.label ?? rule.type;
  if (rule.type === 'list') return `목록에 있는 값만 입력할 수 있습니다.`;
  if (rule.type === 'custom') return '이 셀에 입력할 수 있는 값이 제한되어 있습니다.';
  const op = VALIDATION_OPS.find((x) => x.id === (rule.op ?? 'between'))?.label ?? '';
  const range = rule.op === 'between' || rule.op === 'notBetween' || !rule.op ? `${rule.f1 ?? ''} ~ ${rule.f2 ?? ''}` : `${op} ${rule.f1 ?? ''}`;
  return `${t} 값만 입력할 수 있습니다 (${rule.op === 'notBetween' ? '제외 범위 ' : ''}${range}).`;
}

/** 규칙 범위에서 rg 를 뺀 나머지 (최대 4조각) */
export function subtractRange(rule, rg) {
  if (rule.r2 < rg.r1 || rule.r1 > rg.r2 || rule.c2 < rg.c1 || rule.c1 > rg.c2) return [rule];
  const out = [];
  if (rule.r1 < rg.r1) out.push({ ...rule, r2: rg.r1 - 1 });
  if (rule.r2 > rg.r2) out.push({ ...rule, r1: rg.r2 + 1 });
  const r1 = Math.max(rule.r1, rg.r1);
  const r2 = Math.min(rule.r2, rg.r2);
  if (rule.c1 < rg.c1) out.push({ ...rule, r1, r2, c2: rg.c1 - 1 });
  if (rule.c2 > rg.c2) out.push({ ...rule, r1, r2, c1: rg.c2 + 1 });
  return out;
}

/** 범위 안에서 규칙에 맞지 않는 셀 (잘못된 데이터 표시) */
export function invalidCells(wb, si, limit = 2000) {
  const sheet = wb.sheets[si];
  const out = [];
  for (const [k, cell] of sheet.cells) {
    if (!cell.raw) continue;
    const i = k.indexOf(',');
    const r = Number(k.slice(0, i));
    const c = Number(k.slice(i + 1));
    const rule = validationAt(sheet, r, c);
    if (!rule || rule.type === 'any') continue;
    const v = wb.getValue(si, r, c);
    const text = cell.raw.startsWith('=') ? (typeof v === 'number' || typeof v === 'boolean' ? String(v) : formatValue(v, null).text) : cell.raw.replace(/^'/, '');
    if (!checkValidation(wb, si, rule, r, c, text)) out.push({ r, c });
    if (out.length >= limit) break;
  }
  return out;
}
