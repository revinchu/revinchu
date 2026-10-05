// 계산 가능 여부와 파일 저장값의 신뢰 상태. DOM·통합 문서 변경·네트워크 요청 없음.
import { FUNCS } from './formula.js';

/** 파일에 이미 저장된 이름 오류. 문자열 '#NAME?'는 오류값과 구분한다. */
export function hasSavedNameError(value) {
  return !!value && typeof value === 'object' && (value.error ?? value.code) === '#NAME?';
}

const syntaxMemo = new WeakMap();
function formulaNames(ast) {
  let result = syntaxMemo.get(ast);
  if (result) return result;
  const functions = new Set(), names = new Set();
  const walk = (node, local) => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) { for (const n of node) walk(n, local); return; }
    if (node.type === 'func') {
      if (node.name === 'LET') {
        const scope = new Set(local);
        for (let i = 0; i + 1 < node.args.length; i += 2) {
          walk(node.args[i + 1], scope);
          if (node.args[i].type === 'name') scope.add(node.args[i].v.toLowerCase());
        }
        walk(node.args[node.args.length - 1], scope);
        return;
      }
      if (node.name === 'LAMBDA') {
        const scope = new Set(local);
        for (const n of node.args.slice(0, -1)) if (n.type === 'name') scope.add(n.v.toLowerCase());
        walk(node.args[node.args.length - 1], scope);
        return;
      }
      if (!FUNCS[node.name] && !local.has(node.name.toLowerCase())) functions.add(node.name);
    } else if (node.type === 'name' && !local.has(node.v.toLowerCase()) && !FUNCS[node.v.toUpperCase()]) names.add(node.v);
    for (const [k, value] of Object.entries(node)) if (k !== 'ref' && value && typeof value === 'object') walk(value, local);
  };
  walk(ast, new Set());
  result = { functions: [...functions], names: [...names] };
  syntaxMemo.set(ast, result);
  return result;
}

/** 구문에서 확실히 지원하지 않는 함수/이름. IF 등의 실행되지 않는 분기일 수 있으므로 실제 값도 함께 판정한다. */
export function formulaSupportIssue(cell, isDefined = () => false) {
  if (!cell?.formula) return null;
  if (!cell.ast) return { reason: 'unreadable-formula', functions: [], names: [] };
  const all = formulaNames(cell.ast);
  const nested = [];
  const defined = (name) => {
    const value = isDefined(name);
    if (value && typeof value === 'object' && value.reason) nested.push(value);
    return !!value;
  };
  const functions = all.functions.filter(n => !defined(n));
  const names = all.names.filter(n => !defined(n));
  if (functions.length) return { reason: 'unsupported-function', functions, names };
  if (names.length) return { reason: 'unresolved-name', functions, names };
  return nested[0] ?? null;
}

/** 미지원 수식 상태. value는 호출자가 이미 계산/표시한 값이며 이 함수는 평가를 실행하지 않는다. */
export function calculationStatus(cell, value, support = null, usesSaved = false) {
  if (!cell?.formula) return null;
  const error = value && typeof value === 'object' && typeof value.code === 'string' ? value.code : null;
  const hasSaved = cell.cached !== undefined;
  const stale = !!cell.dirty && hasSaved;
  let status = 'calculated', reason = null, message = '현재 입력으로 계산한 결과입니다.';
  if (!cell.dirty && hasSavedNameError(cell.cached) && error === '#NAME?' && support) {
    status = 'source-error'; reason = 'source-name-error';
    message = '원본 Excel 파일에도 #NAME? 오류가 저장되어 있습니다. 원문의 함수 이름·정의된 이름 또는 수식으로 입력된 텍스트를 확인하세요.';
  } else if (hasSaved && !cell.dirty && (usesSaved || support)) {
    status = 'cached'; reason = support?.reason ?? 'saved-result';
    message = support ? '파일에 저장된 결과를 표시합니다. 이 수식의 재계산은 보장되지 않습니다.' : '파일에 저장된 계산 결과입니다. 현재 입력으로 다시 계산한 값과 다를 수 있습니다.';
  } else if (support && (error === '#NAME?' || support.reason === 'unreadable-formula')) {
    status = stale ? 'stale' : 'blocked'; reason = support.reason;
    message = stale ? '입력 또는 참조가 바뀌어 파일의 이전 결과를 사용할 수 없습니다.' : '이 수식은 현재 계산할 수 없습니다.';
  } else if (error === '#NAME?') {
    status = stale ? 'stale' : 'blocked'; reason = 'unresolved-result';
    message = '이름·함수를 찾을 수 없거나 참조한 수식을 계산할 수 없습니다.';
  }
  return { status, reason, message, formula: cell.raw, error, functions: support?.functions ?? [], names: support?.names ?? [], viaNames: support?.viaNames ?? [],
    ...(status !== 'calculated' && hasSaved ? { savedValue: cell.cached && typeof cell.cached === 'object' ? structuredClone(cell.cached) : cell.cached } : {}) };
}
