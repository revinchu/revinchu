// 도형 전체 텍스트 서식 변경. 부분 서식과 문단 구조는 대상 속성만 바꾼다.
export function shapeTextFormatPatch(shape, patch) {
  const out = { ...patch };
  if (Object.hasOwn(out, 'text')) {
    const current = shape.text ?? shape.paras?.map(p => (p.runs ?? []).map(r => r.t ?? '').join('')).join('\n') ?? '';
    if (String(out.text ?? '') !== String(current)) return { ...out, paras: undefined };
    delete out.text; // 내용이 같으면 기존 부분 서식을 평문으로 바꾸지 않는다.
  }
  if (!Array.isArray(shape.paras)) return out;
  const keys = { font: 'font', size: 'sz', color: 'color', bold: 'b', italic: 'i', underline: 'u', strike: 's' };
  const formats = Object.keys(keys).filter(k => Object.hasOwn(out, k));
  const align = Object.hasOwn(out, 'align');
  if (!formats.length && !align) return out;
  out.paras = shape.paras.map(p => {
    const para = { ...p, runs: (p.runs ?? []).map(r => {
      const run = { ...r };
      for (const key of formats) {
        const value = ['bold', 'italic', 'underline', 'strike'].includes(key) ? !!out[key] : out[key];
        if (value === undefined) delete run[keys[key]]; else run[keys[key]] = value;
      }
      return run;
    }) };
    if (align) { if (out.align === undefined) delete para.align; else para.align = out.align; }
    if (Object.hasOwn(out, 'size')) { if (out.size === undefined) delete para.sz; else if (!para.runs.length || p.sz !== undefined) para.sz = out.size; }
    return para;
  });
  return out;
}
