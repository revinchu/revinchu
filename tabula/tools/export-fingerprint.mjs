// Stress-test integrity checks. Do not call the content path before memory measurements.
function typed(value) {
  if (value === null) return 'null';
  if (typeof value === 'number') return 'number:' + (Object.is(value, -0) ? '-0' : String(value));
  if (typeof value === 'object' && typeof value.code === 'string') return 'error:' + value.code;
  return typeof value + ':' + JSON.stringify(value);
}
export function createCellFingerprint(content = false) {
  let formulas = 0, nonempty = 0, cached = 0;
  const raw = [0, 0], formula = [0, 0], value = [0, 0], cache = [0, 0];
  const hash = (text, into) => {
    let one = 2166136261;
    for (let i = 0; i < text.length; i++) one = Math.imul(one ^ text.charCodeAt(i), 16777619) >>> 0;
    into[0] = (into[0] + one) >>> 0; into[1] = (into[1] ^ one) >>> 0;
  };
  return {
    add(cell, r, c) {
      if (cell.formula) formulas++;
      if (cell.raw !== '' && cell.raw != null || cell.formula) {
        nonempty++;
        if (content) {
          const key = r + ',' + c + ':';
          hash(key + (cell.formula ? 'F:' : 'V:') + String(cell.raw), raw);
          // Formulas must retain their actual source text, not the boolean .formula flag.
          // Scalars compare evaluated literal types/values: Excel can omit an internal
          // input apostrophe without changing a string into a date, number, or formula.
          if (cell.formula) hash(key + String(cell.raw), formula);
          else hash(key + typed(cell.v), value);
        }
      }
      if (content && cell.formula && (cell.cached !== undefined || cell.cachedArray)) {
        cached++;
        const key = r + ',' + c + ':';
        hash(key + (cell.dirty ? 'stale:' : 'valid:') + typed(cell.cached), cache);
        if (cell.cachedArray) {
          hash(key + 'array:' + cell.cachedArray.h + ',' + cell.cachedArray.w, cache);
          for (let i = 0; i < (cell.cachedArray.values?.length ?? 0); i++) hash(key + 'a' + i + ':' + typed(cell.cachedArray.values[i]), cache);
        }
      }
    },
    result() {
      return {formulas, nonempty, ...(content ? {hash:raw[0], xor:raw[1], formulaHash:formula[0], formulaXor:formula[1], valueHash:value[0], valueXor:value[1], cached, cacheHash:cache[0], cacheXor:cache[1]} : {})};
    },
  };
}
