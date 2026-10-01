// 빠른 실행 도구 모음의 기본값과 저장된 이전 옵션 변환 (DOM 없음).
export const DEFAULT_QAT_ORDER = Object.freeze([
  'painter', 'mergeCenter', 'alignCenter', 'incDecimal', 'autosum', 'calcField',
  'toggleGrid', 'condColorScale', 'condDataBar', 'refreshAll', 'textToColumns', 'replace',
]);
export const DEFAULT_QAT_POSITION = 'below';

const LEGACY_QAT_ORDER = ['save', 'undo', 'redo'];
const validOrder = (order) => {
  if (!Array.isArray(order)) return false;
  for (const id of order) if (typeof id !== 'string' || id.length === 0 || id.trim() !== id) return false;
  return true;
};
const defaults = () => ({ qatOrder: [...DEFAULT_QAT_ORDER], qatPosition: DEFAULT_QAT_POSITION });

/**
 * 구 기본값만 새 기본으로 이전합니다. 명시적인 빈 배열과 사용자 순서는 보존합니다.
 * 명령 존재 여부는 앱에서 판단하며, 반환 배열은 원본 옵션/기본 배열과 공유하지 않습니다.
 */
export function normalizeQatOptions(o) {
  if (!o || typeof o !== 'object' || Array.isArray(o)) return defaults();
  const position = o.qatPosition === 'below' ? 'below' : 'above';
  if (o.qatOrder !== undefined && o.qatOrder !== null) {
    if (!validOrder(o.qatOrder)) return defaults();
    const oldDefault = o.qatOrder.length === LEGACY_QAT_ORDER.length && o.qatOrder.every((id, i) => id === LEGACY_QAT_ORDER[i]);
    // 구 기본과 같은 순서를 직접 지정했는지는 저장 데이터만으로 구분할 수 없습니다.
    if (oldDefault && position === 'above') return defaults();
    return { qatOrder: [...o.qatOrder], qatPosition: position };
  }
  if (o.qat === undefined || o.qat === null) return defaults();
  if (!validOrder(o.qat) || o.qat.length === 0) return defaults();
  const order = [...new Set([...LEGACY_QAT_ORDER, ...o.qat])];
  if (order.length === LEGACY_QAT_ORDER.length && position === 'above') return defaults();
  return { qatOrder: order, qatPosition: position };
}
